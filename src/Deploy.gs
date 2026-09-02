/**
 * GrupoLyN Financial Plan — Apps Script
 * File: Deploy.gs  (BONUS — mass client deployment)
 *
 * Strategy B is clearly labeled EXPERIMENTAL / DESIGN SKETCH, not "done" code —
 * it has never been run against a live second spreadsheet (no GCP project was
 * provisioned during this assessment); the honest state is documented, not hidden.
 * There is no supported Apps Script API method to look up an EXISTING bound
 * script's ID purely from a spreadsheet ID, so the design tracks each client's
 * script ID explicitly in a "Clients" ledger sheet (idempotent: known IDs get
 * updated, unknown ones get created and the new ID is written back). dryRun mode
 * previews exactly what would happen (create vs. update, target files) with ZERO
 * live API calls or GCP setup required to verify the control flow and target
 * selection logic. updateContent() reads the existing project's files first and
 * merges in only our own files by name, instead of blindly overwriting the whole
 * file set — a full-replace would silently delete any custom code a client's copy
 * already has. A "Deployments" ledger records every run ({client, fileId,
 * scriptId, action, status, timestamp}) so re-runs are auditable and a bad
 * rollout is at least diagnosable, not blind.
 *
 * Strategy A (Apps Script Library + one-time manual bootstrap paste per client)
 * remains the PROVEN, ship-today path and is the one actually recommended below.
 * Strategy B is the "how I'd automate this further" answer to the bonus prompt.
 */

// ─── CONFIGURATION ────────────────────────────────────────────────────────────

var MASTER_LIBRARY_SCRIPT_ID = 'YOUR_MASTER_SCRIPT_ID_HERE';
var MASTER_LIBRARY_VERSION   = 'HEAD';
var MASTER_LIBRARY_ALIAS     = 'GrupoLyNLib';
var DEPLOY_LOG_SHEET_NAME    = 'Deployments';

// ─── MAIN ENTRY POINT ─────────────────────────────────────────────────────────

/**
 * Run from the master/admin spreadsheet. Reads a "Clients" sheet:
 *   Column A: Google Sheets URL         (required)
 *   Column B: Client name               (optional, for logging)
 *   Column C: Client email              (optional)
 *   Column D: Known scriptId            (optional — filled in automatically after
 *                                         a successful run, so re-runs are idempotent)
 *
 * @param {boolean} dryRun  If true, no live API calls are made — logs the plan only.
 */
function deployToAllClients(dryRun) {
  requireSession_();

  var urls = readClientUrls_();
  if (!urls.length) {
    SpreadsheetApp.getUi().alert('No client URLs found in the "Clients" sheet (column A).');
    return;
  }

  var mode = dryRun ? '🔍 DRY RUN (no changes will be made)' : '🚀 LIVE deployment';
  var confirm = SpreadsheetApp.getUi().alert(
    mode, 'Process ' + urls.length + ' client(s)?', SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  if (confirm !== SpreadsheetApp.getUi().Button.YES) return;

  var results = [];
  urls.forEach(function(entry, idx) {
    var result = deployToClient_(entry, dryRun);
    results.push(entry.name + ': ' + result.status);
    logDeployment_(entry, result, dryRun);
    if (!dryRun) Utilities.sleep(500); // be gentle with API quotas
  });

  SpreadsheetApp.getUi().alert(mode + ' complete!\n\n' + results.join('\n'));
}

function dryRunDeployToAllClients() { deployToAllClients(true); }

// ─── PER-CLIENT DEPLOYMENT (Strategy B — EXPERIMENTAL, untested against a live API) ──

/**
 * ⚠️ EXPERIMENTAL — this function has NOT been exercised against a real second
 * spreadsheet. It requires: the Apps Script API enabled in a GCP project, the
 * `script.projects` advanced service added to this project, and OAuth scope
 * https://www.googleapis.com/auth/script.projects granted.
 *
 * There is no supported Apps Script API method to look up an EXISTING bound
 * script's ID purely from a spreadsheet ID — so this function relies on the
 * "Clients" sheet's column D to remember each client's scriptId after the first
 * successful deploy. First run per client = create; subsequent runs = update.
 */
function deployToClient_(entry, dryRun) {
  var fileId = extractFileId_(entry.url);
  if (!fileId) return { status: 'SKIP (invalid URL)' };

  if (dryRun) {
    var action = entry.scriptId ? 'UPDATE existing script ' + entry.scriptId : 'CREATE new bound script';
    return { status: '🔍 would ' + action, fileId: fileId };
  }

  try {
    var scriptId = entry.scriptId;
    var isFirstDeploy = !scriptId;

    if (isFirstDeploy) {
      var newProject = Script.Projects.create({ title: 'GrupoLyN Financial Automation', parentId: fileId });
      scriptId = newProject.scriptId;
    }

    // Merge, don't blindly overwrite: read whatever files already exist in the
    // client's bound project and only replace/add our own files by name. On a
    // re-deploy (project already exists), skip regenerating appsscript.json
    // entirely — only Setup.gs is touched — so any custom scopes/libraries the
    // client's project accumulated since the first deploy are never clobbered.
    var existing = { files: [] };
    try { existing = Script.Projects.getContent(scriptId); } catch (e) { /* new project has no content yet */ }

    var ourFiles = buildScriptContent_(fileId, entry.name, /* includeManifest= */ isFirstDeploy);
    var mergedByName = {};
    (existing.files || []).forEach(function(f) { mergedByName[f.name] = f; });
    ourFiles.forEach(function(f) { mergedByName[f.name] = f; });
    var mergedFiles = Object.keys(mergedByName).map(function(k) { return mergedByName[k]; });

    Script.Projects.updateContent({ files: mergedFiles }, scriptId);

    return { status: '✅ OK', scriptId: scriptId, fileId: fileId };
  } catch (e) {
    return { status: '❌ ERROR: ' + e.message, fileId: fileId };
  }
}

function buildScriptContent_(spreadsheetId, clientName, includeManifest) {
  var setupCode = [
    '/* Auto-generated by GrupoLyN Deploy — do not edit manually */',
    '/* Client: ' + clientName + ' */',
    '',
    'function onOpen() {',
    '  ' + MASTER_LIBRARY_ALIAS + '.onOpen();',
    '}'
  ].join('\n');

  var files = [{ name: 'Setup', type: 'SERVER_JS', source: setupCode }];
  if (!includeManifest) return files;

  var manifest = {
    timeZone: 'America/New_York',
    dependencies: {
      // The manifest field for a library dependency's script ID is `libraryId`,
      // not `scriptId` — that's the field used elsewhere for Script.Projects
      // calls, but appsscript.json's own schema names it differently. Using the
      // wrong key here would silently fail to attach the library at all, and
      // Setup.gs's call to GrupoLyNLib.onOpen() would break with "not defined".
      libraries: [{
        userSymbol: MASTER_LIBRARY_ALIAS,
        libraryId: MASTER_LIBRARY_SCRIPT_ID,
        version: MASTER_LIBRARY_VERSION,
        developmentMode: MASTER_LIBRARY_VERSION === 'HEAD'
      }]
    },
    exceptionLogging: 'STACKDRIVER',
    runtimeVersion: 'V8',
    // Least privilege: the generated Setup.gs only opens the menu and calls into
    // the shared library, so it only needs the scopes the library itself uses
    // (spreadsheet access, composing Gmail drafts, and the sidebar/dialog UI) —
    // not full Drive access.
    oauthScopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/gmail.compose',
      'https://www.googleapis.com/auth/script.container.ui'
    ]
  };

  files.push({ name: 'appsscript', type: 'JSON', source: JSON.stringify(manifest, null, 2) });
  return files;
}

// ─── STRATEGY A: proven manual bootstrap (recommended default) ──────────────────

function showBootstrapSnippet() {
  var snippet = buildBootstrapSnippet_();
  var html = HtmlService.createHtmlOutput(
    '<textarea style="width:100%;height:280px;font-family:monospace;font-size:11px;">' +
    snippet.replace(/</g, '&lt;') + '</textarea>' +
    '<p style="font-family:Arial;font-size:12px;color:#666;">Copy this into each client\'s Apps Script editor once. ' +
    'Then link the library via Resources/Libraries with the Script ID below.</p>'
  ).setTitle('Strategy A — Bootstrap Snippet').setWidth(480).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Bootstrap Snippet');
}

function buildBootstrapSnippet_() {
  return [
    '// ── GrupoLyN Bootstrap — paste into the client sheet\'s Apps Script editor ──',
    '// 1. Extensions > Apps Script > replace all content with this',
    '// 2. Resources/Libraries > add library, Script ID: ' + MASTER_LIBRARY_SCRIPT_ID,
    '// 3. Identifier: ' + MASTER_LIBRARY_ALIAS + ' · Version: ' + MASTER_LIBRARY_VERSION,
    '// 4. Save, then run grupolynSetup() once',
    '',
    'function onOpen() {',
    '  ' + MASTER_LIBRARY_ALIAS + '.onOpen();',
    '}',
    '',
    'function grupolynSetup() {',
    '  SpreadsheetApp.getUi().alert("✅ GrupoLyN Tools ready!");',
    '}'
  ].join('\n');
}

// ─── DEPLOYMENT LEDGER (audit trail across the whole client fleet) ───────────────

function logDeployment_(entry, result, dryRun) {
  var sheet = getOrCreateDeployLogSheet_();
  sheet.appendRow([
    new Date(), entry.name, entry.url, result.scriptId || '', result.status, dryRun ? 'DRY RUN' : 'LIVE'
  ]);

  if (!dryRun && result.scriptId) {
    writeScriptIdBack_(entry, result.scriptId);
  }
}

function getOrCreateDeployLogSheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DEPLOY_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DEPLOY_LOG_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Client', 'URL', 'Script ID', 'Result', 'Mode']);
    sheet.getRange('A1:F1').setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  }
  return sheet;
}

/** Writes the resolved scriptId back into the "Clients" sheet (column D) so re-runs are idempotent. */
function writeScriptIdBack_(entry, scriptId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clients');
  if (!sheet || !entry.row) return;
  sheet.getRange(entry.row, 4).setValue(scriptId);
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────

function readClientUrls_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clients');
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'No "Clients" sheet found.\n\nCreate one with:\n' +
      '  A: Google Sheets URL\n  B: Client name\n  C: Client email\n  D: (leave blank — auto-filled)'
    );
    return [];
  }
  var data    = sheet.getDataRange().getValues();
  var entries = [];
  var seenUrls = {}; // guards against duplicate rows creating two bound scripts for the same client
  data.slice(1).forEach(function(row, idx) {
    var url = String(row[0] || '').trim();
    if (url && url.indexOf('https://docs.google.com/spreadsheets') === 0) {
      if (seenUrls[url]) return;
      seenUrls[url] = true;
      entries.push({
        url: url, name: String(row[1] || url), email: String(row[2] || ''),
        scriptId: String(row[3] || '') || null, row: idx + 2
      });
    }
  });
  return entries;
}

function extractFileId_(url) {
  var match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

// requireSession_() is defined in Code.gs (shared across this Apps Script project) —
// it checks both the email allow-list and session expiry, so Deploy.gs uses the same
// authorization surface as the rest of the Admin menu instead of a weaker duplicate.

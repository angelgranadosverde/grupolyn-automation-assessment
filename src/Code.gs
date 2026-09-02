/**
 * GrupoLyN Financial Automation — Code.gs
 * Admin menu (document-scoped auth + optional email allow-list + audit log)
 * + Monthly Comparative Report engine (adapted to the real "Prospr Script" template)
 *
 * Design notes:
 *  - Admin code is DOCUMENT-scoped (shared per spreadsheet copy), not per-user.
 *    Sessions remain per-user (each browser must unlock its own 30-min window).
 *  - No hardcoded default password ships in code — first use forces a real setup.
 *  - Optional email allow-list (Document Properties) — if configured, only listed
 *    emails ever see the password prompt succeed.
 *  - Failed-attempt lockout (5 tries / 10 min).
 *  - Every unlock, failed attempt, code change and report run is logged to a
 *    hidden "Admin Audit Log" sheet.
 *  - Report engine: currency-string-safe number parsing (with a visible warning
 *    instead of a silent $0), duplicate-category-name collision guard, and a
 *    stale-"Total"-row cross-check.
 *  - Gmail draft: real recipient lookup (never sends to an empty address), and a
 *    branded HTML email instead of a plain-text dump.
 */

// ─── CONFIG ─────────────────────────────────────────────────────────────────

var ADMIN_CODE_PROP     = 'ADMIN_CODE';        // Document Properties
var ADMIN_EMAILS_PROP   = 'ADMIN_EMAILS';      // Document Properties (comma-separated, optional)
var CLIENT_EMAIL_PROP   = 'CLIENT_EMAIL';      // Document Properties (optional override)
var SESSION_PROP        = 'adminSession';      // User Properties
var FAILED_COUNT_PROP   = 'failedAttempts';    // User Properties
var LOCK_UNTIL_PROP     = 'lockUntil';         // User Properties
var SESSION_TTL_MS      = 30 * 60 * 1000;      // 30 min
var MAX_FAILED_ATTEMPTS = 5;
var LOCKOUT_MS          = 10 * 60 * 1000;      // 10 min
var AUDIT_SHEET_NAME    = 'Admin Audit Log';
var DEVIATION_THRESHOLD = 0.15;

// ─── MENU & AUTH ──────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔒 Admin')
    .addItem('Unlock Admin Panel…', 'promptAdminLogin')
    .addToUi();
}

function promptAdminLogin() {
  var ui = SpreadsheetApp.getUi();

  if (!isEmailAllowed_()) {
    ui.alert('❌ Not authorized', 'Your account is not on the admin allow-list for this spreadsheet.', ui.ButtonSet.OK);
    logAudit_('DENIED (not on allow-list)');
    return;
  }

  if (isSessionActive_()) { showAdminPanel_(); return; }

  var lockUntil = Number(PropertiesService.getUserProperties().getProperty(LOCK_UNTIL_PROP) || 0);
  if (Date.now() < lockUntil) {
    var minsLeft = Math.ceil((lockUntil - Date.now()) / 60000);
    ui.alert('🔒 Locked out', 'Too many failed attempts. Try again in ' + minsLeft + ' minute(s).', ui.ButtonSet.OK);
    return;
  }

  var docProps = PropertiesService.getDocumentProperties();
  var storedCode = docProps.getProperty(ADMIN_CODE_PROP);

  if (!storedCode) {
    setupInitialAdminCode_();
    return;
  }

  var result = ui.prompt('🔒 Admin Access', 'Enter admin code:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;

  if (result.getResponseText().trim() === storedCode) {
    resetFailedAttempts_();
    PropertiesService.getUserProperties().setProperty(SESSION_PROP, String(Date.now()));
    logAudit_('UNLOCK success');
    showAdminPanel_();
  } else {
    var attempts = registerFailedAttempt_();
    logAudit_('UNLOCK failed (attempt ' + attempts + ')');
    if (attempts >= MAX_FAILED_ATTEMPTS) {
      PropertiesService.getUserProperties().setProperty(LOCK_UNTIL_PROP, String(Date.now() + LOCKOUT_MS));
      ui.alert('❌ Access Denied', 'Too many failed attempts. Locked for 10 minutes.', ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Access Denied', 'Incorrect code. (' + attempts + '/' + MAX_FAILED_ATTEMPTS + ' attempts)', ui.ButtonSet.OK);
    }
  }
}

/** First-ever use on this spreadsheet copy: force the admin to set a real code (no shipped default). */
function setupInitialAdminCode_() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('👋 First-time setup', 'No admin code is set for this spreadsheet yet. You will now create one.', ui.ButtonSet.OK);

  var r1 = ui.prompt('Create Admin Code', 'Enter a new admin code (min 6 characters):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var code1 = r1.getResponseText().trim();
  if (code1.length < 6) { ui.alert('❌ Too short. Run "Unlock Admin Panel" again to retry.'); return; }

  var r2 = ui.prompt('Confirm Admin Code', 'Re-enter the same code to confirm:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  if (r2.getResponseText().trim() !== code1) { ui.alert('❌ Codes did not match. Run "Unlock Admin Panel" again to retry.'); return; }

  PropertiesService.getDocumentProperties().setProperty(ADMIN_CODE_PROP, code1);
  logAudit_('Admin code created (first-time setup)');
  ui.alert('✅ Admin code set. Click "Unlock Admin Panel" again to log in.');
}

function isEmailAllowed_() {
  var allowList = PropertiesService.getDocumentProperties().getProperty(ADMIN_EMAILS_PROP);
  if (!allowList) return true; // not configured yet — open by design until an admin sets it
  var email = activeUserEmailStrict_();
  // If Apps Script can't identify the active user (this genuinely happens for
  // external/off-domain collaborators), fail CLOSED rather than falling back to
  // the script owner's identity — a fallback here would let an unidentifiable
  // visitor in under the owner's name, which defeats the allow-list entirely.
  if (!email) return false;
  var allowed = allowList.split(',').map(function(e) { return e.trim().toLowerCase(); });
  return allowed.indexOf(email.toLowerCase()) !== -1;
}

/** Strictly the active user's email, or '' if Apps Script can't determine it. */
function activeUserEmailStrict_() {
  return Session.getActiveUser().getEmail() || '';
}

/**
 * The actual person operating the menu right now, for logging purposes — falls
 * back to the effective (script-owner) identity only so the audit log never has
 * a blank "User" column. NOT used for access-control decisions (see
 * activeUserEmailStrict_ / isEmailAllowed_): Session.getEffectiveUser() reports
 * the identity code executes AS (relevant for trigger delegation), which is the
 * wrong signal for "who is this" when denying access.
 */
function currentUserEmail_() {
  return activeUserEmailStrict_() || Session.getEffectiveUser().getEmail() || '(unknown)';
}

function isSessionActive_() {
  var ts = PropertiesService.getUserProperties().getProperty(SESSION_PROP);
  return ts && (Date.now() - parseInt(ts, 10)) < SESSION_TTL_MS;
}

/**
 * Guard for every sensitive action reachable from the Admin sidebar. The sidebar
 * itself persists in the UI well past the 30-minute session window (Sheets does not
 * auto-close it), so without a server-side check here, a stale/expired session's
 * buttons would keep working indefinitely. Every action below must call this first.
 */
function requireSession_() {
  if (!isEmailAllowed_()) throw new Error('Your account is no longer authorized for this spreadsheet.');
  if (!isSessionActive_()) throw new Error('Session expired. Close this panel and unlock Admin again.');
}

function registerFailedAttempt_() {
  var props = PropertiesService.getUserProperties();
  var n = Number(props.getProperty(FAILED_COUNT_PROP) || 0) + 1;
  props.setProperty(FAILED_COUNT_PROP, String(n));
  return n;
}

function resetFailedAttempts_() {
  var props = PropertiesService.getUserProperties();
  props.deleteProperty(FAILED_COUNT_PROP);
  props.deleteProperty(LOCK_UNTIL_PROP);
}

function showAdminPanel_() {
  var hasAllowList = !!PropertiesService.getDocumentProperties().getProperty(ADMIN_EMAILS_PROP);
  var allowListNote = hasAllowList
    ? '<p style="color:#137333;font-size:11px;">✅ Email allow-list configured</p>'
    : '<p style="color:#b06000;font-size:11px;">⚠️ No email allow-list set — anyone with the code can enter</p>';

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(
      '<style>' +
      'body{font-family:Google Sans,Arial,sans-serif;padding:16px;background:#f8f9fa;}' +
      'h3{color:#1a73e8;margin-top:0;}' +
      'p{font-size:12px;color:#666;margin-bottom:16px;}' +
      '.btn{display:block;width:100%;padding:10px;margin:8px 0;border:none;' +
      'border-radius:6px;font-size:13px;cursor:pointer;font-weight:600;}' +
      '.p{background:#1a73e8;color:#fff;}' +
      '.s{background:#e8f0fe;color:#1a73e8;}' +
      '.d{background:#fce8e6;color:#c5221f;}' +
      '</style>' +
      '<h3>✅ Admin Panel</h3>' +
      '<p>Session active · 30 minutes</p>' +
      allowListNote +
      '<button class="btn p" onclick="r(\'runMonthlyComparativeReport\')">📊 Monthly Report + Gmail Draft</button>' +
      '<button class="btn s" onclick="r(\'setNewAdminCode\')">🔑 Change Admin Code</button>' +
      '<button class="btn s" onclick="r(\'setAdminEmailAllowList\')">👥 Set Authorized Emails</button>' +
      '<button class="btn s" onclick="r(\'viewAuditLog\')">📋 View Audit Log</button>' +
      '<button class="btn d" onclick="r(\'lockPanel\')">🔒 Lock Panel</button>' +
      '<script>' +
      'function r(fn){' +
      '  google.script.run' +
      '    .withSuccessHandler(function(shouldClose){if(shouldClose)google.script.host.close();})' +
      '    .withFailureHandler(function(e){alert("Error: "+e.message);})[fn]();' +
      '}' +
      '</script>'
    ).setTitle('Admin Panel').setWidth(280).setHeight(320)
  );
}

function lockPanel() {
  PropertiesService.getUserProperties().deleteProperty(SESSION_PROP);
  logAudit_('Panel locked');
  SpreadsheetApp.getUi().alert('🔒 Admin panel locked.');
  return true;
}

// ─── PASSWORD & ALLOW-LIST MANAGEMENT ─────────────────────────────────────────

function setNewAdminCode() {
  requireSession_();
  var ui       = SpreadsheetApp.getUi();
  var docProps = PropertiesService.getDocumentProperties();

  var currentResp = ui.prompt('Verify Current Code', 'Enter your current admin code:', ui.ButtonSet.OK_CANCEL);
  if (currentResp.getSelectedButton() !== ui.Button.OK) return;

  var stored = docProps.getProperty(ADMIN_CODE_PROP);
  if (currentResp.getResponseText().trim() !== stored) {
    ui.alert('❌ Incorrect current code. Cannot change.');
    logAudit_('Code change DENIED (wrong current code)');
    return;
  }

  var newResp = ui.prompt('Set New Code', 'Enter the new admin code (min 6 characters):', ui.ButtonSet.OK_CANCEL);
  if (newResp.getSelectedButton() !== ui.Button.OK) return;

  var newCode = newResp.getResponseText().trim();
  if (newCode.length < 6) { ui.alert('❌ Code must be at least 6 characters.'); return; }

  docProps.setProperty(ADMIN_CODE_PROP, newCode);
  logAudit_('Admin code changed');
  ui.alert('✅ Admin code updated successfully!');
  return true;
}

function setAdminEmailAllowList() {
  requireSession_();
  var ui       = SpreadsheetApp.getUi();
  var docProps = PropertiesService.getDocumentProperties();
  var current  = docProps.getProperty(ADMIN_EMAILS_PROP) || '';

  var resp = ui.prompt(
    'Authorized Admin Emails',
    'Enter comma-separated emails allowed to use the Admin menu.\n' +
    'Leave blank to remove the restriction (anyone with the code can enter).\n\n' +
    'Current: ' + (current || '(none set — open to anyone with the code)'),
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var value = resp.getResponseText().trim();
  if (value) {
    docProps.setProperty(ADMIN_EMAILS_PROP, value);
    logAudit_('Allow-list updated: ' + value);
    ui.alert('✅ Allow-list set. Only listed emails can unlock the Admin menu.');
  } else {
    docProps.deleteProperty(ADMIN_EMAILS_PROP);
    logAudit_('Allow-list removed');
    ui.alert('✅ Allow-list removed. Anyone with the code can unlock the Admin menu.');
  }
  return true;
}

// ─── AUDIT LOG ────────────────────────────────────────────────────────────────

function logAudit_(action) {
  try {
    var sheet = getOrCreateAuditSheet_();
    sheet.appendRow([new Date(), currentUserEmail_(), action]);
  } catch (e) {
    // Audit logging must never block the primary action.
  }
}

function getOrCreateAuditSheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AUDIT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AUDIT_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'User', 'Action']);
    sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 320);
    sheet.hideSheet();
  }
  return sheet;
}

/**
 * Shows the last 25 audit entries in a read-only modal, never unhiding the
 * sheet itself — the audit log stays hidden from every collaborator except
 * through this modal.
 */
function viewAuditLog() {
  requireSession_();
  var sheet = getOrCreateAuditSheet_();
  var data  = sheet.getDataRange().getValues();
  var rows  = data.slice(1).slice(-25).reverse();

  var tableRows = rows.map(function(r) {
    var ts = (r[0] instanceof Date) ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'MM/dd HH:mm') : String(r[0]);
    return '<tr><td style="padding:6px 8px;border-bottom:1px solid #eee;white-space:nowrap;">' + ts + '</td>' +
      '<td style="padding:6px 8px;border-bottom:1px solid #eee;">' + escHtml_(r[1]) + '</td>' +
      '<td style="padding:6px 8px;border-bottom:1px solid #eee;">' + escHtml_(r[2]) + '</td></tr>';
  }).join('');

  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:12px;">' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr style="background:#1a73e8;color:#fff;">' +
    '<th style="padding:6px 8px;text-align:left;">Timestamp</th>' +
    '<th style="padding:6px 8px;text-align:left;">User</th>' +
    '<th style="padding:6px 8px;text-align:left;">Action</th></tr>' +
    (tableRows || '<tr><td colspan="3" style="padding:12px;color:#888;">No entries yet.</td></tr>') +
    '</table></div>'
  ).setWidth(560).setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, '📋 Admin Audit Log (last 25 actions)');
}

// ─── REPORT ENGINE ────────────────────────────────────────────────────────────

function runMonthlyComparativeReport() {
  requireSession_();
  var ui              = SpreadsheetApp.getUi();
  var ss               = SpreadsheetApp.getActiveSpreadsheet();
  var budgetSheetName  = 'Monthly Budget';
  var sheet            = ss.getSheetByName(budgetSheetName);

  if (!sheet) {
    ui.alert('Error: Sheet "' + budgetSheetName + '" not found!');
    return;
  }

  var clientEmail = getClientEmail_();
  if (!clientEmail) {
    ui.alert(
      '⚠️ No client email found',
      'No email was found on the "Cover" tab. You\'ll be asked for one now — it will be remembered for future reports.',
      ui.ButtonSet.OK
    );
    var emailResp = ui.prompt('Client Email', 'Enter the client\'s email for this report:', ui.ButtonSet.OK_CANCEL);
    if (emailResp.getSelectedButton() !== ui.Button.OK) return;
    clientEmail = emailResp.getResponseText().trim();
    if (!clientEmail || clientEmail.indexOf('@') < 1) {
      ui.alert('❌ Invalid email. Report generation cancelled.');
      return;
    }
    PropertiesService.getDocumentProperties().setProperty(CLIENT_EMAIL_PROP, clientEmail);
  }

  var year  = sheet.getRange('F2').getValue();
  var month = sheet.getRange('F3').getValue();
  var bom   = sheet.getRange('H2').getValue();
  var eom   = sheet.getRange('H3').getValue();

  var CATEGORY_COL  = 1;  // Column B
  var ITEM_DESC_COL = 2;  // Column C
  var BUDGET_COL    = 3;  // Column D
  var ACTUAL_COL    = 5;  // Column F
  var START_ROW     = 5;

  var data              = sheet.getDataRange().getValues();
  var allCategoriesData = {};
  var categoryOrder     = [];  // preserves first-seen order + guards name collisions
  var parseWarnings      = [];

  var currentCategoryName        = null;
  var currentCategoryKey         = null;
  var currentCategoryItems       = [];
  var currentCategoryItemsSum    = { budget: 0, actual: 0 };
  var currentCategoryTotalBudget = 0;
  var currentCategoryTotalActual = 0;

  function flushCurrent_() {
    if (!currentCategoryName) return;
    allCategoriesData[currentCategoryKey] = {
      displayName: currentCategoryName,
      items: currentCategoryItems,
      totalBudget: currentCategoryTotalBudget,
      totalActual: currentCategoryTotalActual,
      itemsSum: currentCategoryItemsSum
    };
    categoryOrder.push(currentCategoryKey);
  }

  function uniqueKey_(name, rowIndex) {
    var base = name;
    if (allCategoriesData.hasOwnProperty(base)) {
      base = name + ' (fila ' + (rowIndex + 1) + ')';
      parseWarnings.push('Nombre de categoría repetido "' + name + '" en la fila ' + (rowIndex + 1) + ' — renombrado para evitar mezclar datos.');
    }
    return base;
  }

  // A "Total" row is only ever "Total" alone or "Total <the category currently
  // being accumulated>" — matched against that specific name, not a generic
  // substring test. A plain indexOf('Total') check would misfire on a category
  // that happens to legitimately contain the word "Total" (e.g. "Total Rewards"
  // as a benefits category): it would neither start a new category (it "contains
  // Total") nor be excluded from closing the prior one (it "starts with Total"),
  // silently zeroing out the prior category's totals and dropping every line
  // item that actually belongs to the new one.
  function isTotalRowFor_(header, catName) {
    if (!catName) return false;
    return header === 'Total' || header === ('Total ' + catName);
  }

  for (var i = START_ROW - 1; i < data.length; i++) {
    var row = data[i];
    var categoryHeader  = (row[CATEGORY_COL] != null) ? String(row[CATEGORY_COL]).trim() : '';
    var itemDescription = (row[ITEM_DESC_COL] != null) ? String(row[ITEM_DESC_COL]).trim() : '';
    var budgetValue = toNumber_(row[BUDGET_COL], parseWarnings, 'B' + (i + 1) + '/D' + (i + 1));
    var actualValue = toNumber_(row[ACTUAL_COL], parseWarnings, 'F' + (i + 1));
    var isTotalRow = isTotalRowFor_(categoryHeader, currentCategoryName);

    if (categoryHeader && !isTotalRow) {
      flushCurrent_();
      currentCategoryName        = categoryHeader;
      currentCategoryKey         = uniqueKey_(categoryHeader, i);
      currentCategoryItems       = [];
      currentCategoryItemsSum    = { budget: 0, actual: 0 };
      currentCategoryTotalBudget = 0;
      currentCategoryTotalActual = 0;
    }

    if (currentCategoryName &&
        (itemDescription || budgetValue !== 0 || actualValue !== 0) &&
        !isTotalRow) {
      currentCategoryItems.push({ description: itemDescription, budget: budgetValue, actual: actualValue });
      currentCategoryItemsSum.budget += budgetValue;
      currentCategoryItemsSum.actual += actualValue;
    }

    if (isTotalRow && currentCategoryName) {
      currentCategoryTotalBudget = budgetValue;
      currentCategoryTotalActual = actualValue;
      flushCurrent_();
      currentCategoryName = null;
      currentCategoryKey  = null;
    }
  }
  flushCurrent_();

  var reportRowsForSheet = [[
    'Category', 'Item Description', 'Actual', 'Planned', 'Deviation ($)', 'Deviation (%)', 'Status'
  ]];
  var reportLinesForEmailHtml = [];
  var anyFlagged = false;

  categoryOrder.forEach(function(key) {
    var d      = allCategoriesData[key];
    var name   = d.displayName;
    var actual = d.totalActual || 0;
    var budget = d.totalBudget || 0;

    // Skip categories with no data at all (e.g. section wrappers like "Income"
    // that never hit their own "Total" row because they only contain sub-headers).
    if (budget === 0 && actual === 0 && d.items.length === 0) return;

    // Stale-Total-row cross-check: warn (don't silently trust) if the "Total" row's
    // own value diverges meaningfully from the sum of the line items under it.
    var itemsSumBudget = d.itemsSum.budget;
    var itemsSumActual = d.itemsSum.actual;
    var staleNote = '';
    if (Math.abs(itemsSumBudget - budget) > 1 || Math.abs(itemsSumActual - actual) > 1) {
      staleNote = ' ⚠️ Total row does not match sum of line items (possible stale formula).';
      parseWarnings.push('"' + name + '": la fila "Total" (' + fmtMoney_(budget) + ' / ' + fmtMoney_(actual) +
        ') no coincide con la suma de sus líneas (' + fmtMoney_(itemsSumBudget) + ' / ' + fmtMoney_(itemsSumActual) + ').');
    }

    var deviation    = actual - budget;
    var deviationPct = budget === 0 ? (actual === 0 ? 0 : (actual > 0 ? 1 : -1)) : deviation / budget;
    var deviationPctStr = (deviationPct * 100).toFixed(1) + '%';
    var status = Math.abs(deviationPct) > DEVIATION_THRESHOLD ? (deviationPct > 0 ? 'Over' : 'Under') : 'OK';
    var deviationSign = deviation > 0 ? '+' : '';

    var include = status !== 'OK' || (budget === 0 && actual !== 0) || (budget !== 0 && actual === 0);
    if (!include) return;
    anyFlagged = anyFlagged || (status !== 'OK');

    reportRowsForSheet.push([
      name, '', '$' + actual.toFixed(2), '$' + budget.toFixed(2),
      deviationSign + deviation.toFixed(2), deviationPctStr, status
    ]);

    var significantItems = [];
    d.items.forEach(function(item) {
      var itemDeviation = item.actual - item.budget;
      var itemDeviationPct = item.budget === 0
        ? (item.actual === 0 ? 0 : (item.actual > 0 ? 1 : -1))
        : itemDeviation / item.budget;
      if (Math.abs(itemDeviationPct) > DEVIATION_THRESHOLD ||
          (item.budget === 0 && item.actual !== 0) ||
          (item.budget !== 0 && item.actual === 0)) {
        var itemDiffSign = itemDeviation > 0 ? '+' : '';
        significantItems.push({
          description: item.description, actual: item.actual, budget: item.budget,
          diff: itemDiffSign + itemDeviation.toFixed(2), diffPct: (itemDeviationPct * 100).toFixed(1) + '%'
        });
        reportRowsForSheet.push([
          '', item.description, '$' + item.actual.toFixed(2), '$' + item.budget.toFixed(2),
          itemDiffSign + itemDeviation.toFixed(2), (itemDeviationPct * 100).toFixed(1) + '%', ''
        ]);
      }
    });
    reportRowsForSheet.push(['', '', '', '', '', '', '']);

    reportLinesForEmailHtml.push({
      name: name, status: status, actual: actual, budget: budget,
      deviationPctStr: deviationPctStr, staleNote: staleNote, items: significantItems
    });
  });

  generateTabularReportSheet(reportRowsForSheet, month, year, bom, eom);

  var subject = month + ' ' + year + ' Budget Comparison';
  generateReportAsEmailDraft(clientEmail, subject, month, year, bom, eom, reportLinesForEmailHtml);

  logAudit_('Report generated: ' + subject + (parseWarnings.length ? ' (' + parseWarnings.length + ' warning(s))' : ''));

  if (parseWarnings.length) {
    ui.alert(
      '⚠️ Report generated with warnings',
      'The report was created, but please review:\n\n' + parseWarnings.join('\n'),
      ui.ButtonSet.OK
    );
  }
  return true;
}

// ─── NUMBER PARSING (currency-safe) ───────────────────────────────────────────

function toNumber_(v, warnings, cellRef) {
  if (typeof v === 'number') return v;
  if (v === '' || v === null || v === undefined) return 0;
  var str = String(v).trim();
  if (!str || str === '-') return 0;
  // Accounting-style negatives — "($1,234.56)" — are common in Sheets/Excel
  // currency formatting; parseFloat can't read a string starting with "(", so
  // unwrap that before stripping the rest of the currency formatting.
  var isParenNegative = /^\(.*\)$/.test(str);
  if (isParenNegative) str = str.slice(1, -1);
  var cleaned = str.replace(/[$,\s]/g, '');
  var n = parseFloat(cleaned);
  if (isNaN(n)) {
    if (warnings) warnings.push('Valor no numérico en ' + cellRef + ': "' + str + '" — tratado como $0.');
    return 0;
  }
  return isParenNegative ? -Math.abs(n) : n;
}

function fmtMoney_(n) {
  return '$' + Math.abs(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// ─── SHEET REPORT WRITER ──────────────────────────────────────────────────────

function prettyDate(date) {
  if (!(date instanceof Date)) return date;
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

function generateTabularReportSheet(tableData, month, year, bom, eom) {
  var ss              = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheetName = month + ' Budget Comparison';
  var reportSheet     = ss.getSheetByName(reportSheetName);
  var preservedNote   = '';

  if (reportSheet) {
    if (looksLikeOurReport_(reportSheet)) {
      reportSheet.clear();
    } else {
      // A tab with this exact name already exists but doesn't look like a report
      // this tool generated (e.g. a pre-existing template tab with different
      // content). Rename it aside instead of silently destroying it with clear().
      var backupName = uniqueBackupName_(ss, reportSheetName);
      reportSheet.setName(backupName);
      preservedNote = '\n\n(The existing "' + reportSheetName + '" tab was not a report from ' +
        'this tool, so it was preserved as "' + backupName + '" instead of being overwritten.)';
      reportSheet = ss.insertSheet(reportSheetName);
    }
  } else {
    reportSheet = ss.insertSheet(reportSheetName);
  }

  reportSheet.getRange('C2').setValue('Year');
  reportSheet.getRange('D2').setValue(year);
  reportSheet.getRange('E2').setValue('BOM');
  reportSheet.getRange('F2').setValue(Utilities.formatDate(bom, Session.getScriptTimeZone(), 'M/d/yyyy'));
  reportSheet.getRange('C3').setValue('Month');
  reportSheet.getRange('D3').setValue(month);
  reportSheet.getRange('E3').setValue('EOM');
  reportSheet.getRange('F3').setValue(Utilities.formatDate(eom, Session.getScriptTimeZone(), 'M/d/yyyy'));

  var nRows        = tableData.length;
  var nCols        = 7;
  var dataStartRow = 5;
  reportSheet.getRange(dataStartRow, 1, nRows, nCols).setValues(tableData);

  var headerRange = reportSheet.getRange(dataStartRow, 1, 1, nCols);
  headerRange.setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');

  if (nRows > 1) {
    reportSheet.getRange(dataStartRow + 1, 1, nRows - 1, nCols).setFontFamily('Arial, Helvetica, sans-serif');
  }

  for (var i = 0; i < nRows - 1; i++) {
    var currentRowIndex = dataStartRow + 1 + i;
    var rowData         = tableData[i + 1];

    if (String(rowData[0]) === '' && String(rowData[1]) === '') {
      reportSheet.getRange(currentRowIndex, 1, 1, nCols).setBackground('#ffffff');
      continue;
    }
    if (rowData[0] === '' && rowData[1] !== '') {
      reportSheet.getRange(currentRowIndex, 1, 1, nCols).setBackground('#f7f7f7');
      reportSheet.getRange(currentRowIndex, 2).setFontSize(9);
    } else {
      var status     = rowData[6];
      var statusCell = reportSheet.getRange(currentRowIndex, 7);
      if (status === 'Over')  statusCell.setBackground('#ffd6d6');
      if (status === 'Under') statusCell.setBackground('#d6ffd6');
      reportSheet.getRange(currentRowIndex, 1).setFontWeight('bold');
      reportSheet.getRange(currentRowIndex, 1, 1, nCols).setBackground('#e8f0fe');
    }
  }

  for (var c = 1; c <= nCols; c++) { reportSheet.autoResizeColumn(c); }

  reportSheet.getRange(dataStartRow, 3, nRows, 1).setHorizontalAlignment('right');
  reportSheet.getRange(dataStartRow, 4, nRows, 1).setHorizontalAlignment('right');
  reportSheet.getRange(dataStartRow, 5, nRows, 1).setHorizontalAlignment('right');
  reportSheet.getRange(dataStartRow, 6, nRows, 1).setHorizontalAlignment('right');

  ss.setActiveSheet(reportSheet);
  SpreadsheetApp.getUi().alert('✅ Report generated in the "' + reportSheetName + '" sheet.' + preservedNote);
}

/**
 * Our report's data header row is always 'Category' / … / 'Status' at row 5.
 * Used to tell "a report this tool wrote before" apart from "some other tab
 * that happens to share this month's name" before deciding whether it's safe
 * to clear() it.
 */
function looksLikeOurReport_(sheet) {
  try {
    var header = sheet.getRange(5, 1, 1, 7).getValues()[0];
    return header[0] === 'Category' && header[6] === 'Status';
  } catch (e) {
    return false;
  }
}

function uniqueBackupName_(ss, baseName) {
  var name = baseName + ' (original)';
  var i = 2;
  while (ss.getSheetByName(name)) {
    name = baseName + ' (original ' + i + ')';
    i++;
  }
  return name;
}

// ─── CLIENT EMAIL LOOKUP ───────────────────────────────────────────────────────

function getClientEmail_() {
  var override = PropertiesService.getDocumentProperties().getProperty(CLIENT_EMAIL_PROP);
  if (override) return override;

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var cover = ss.getSheetByName('Cover');
  if (cover) {
    var data = cover.getDataRange().getValues();
    for (var r = 0; r < data.length; r++) {
      for (var c = 0; c < data[r].length; c++) {
        if (/email/i.test(String(data[r][c])) && c + 1 < data[r].length) {
          var email = String(data[r][c + 1]).trim();
          if (email && email.indexOf('@') > 0) return email;
        }
      }
    }
  }
  return null;
}

// ─── GMAIL DRAFT (branded HTML) ────────────────────────────────────────────────

function generateReportAsEmailDraft(clientEmail, subject, month, year, bom, eom, categories) {
  try {
    var htmlBody = buildEmailHtml_(month, year, bom, eom, categories);
    GmailApp.createDraft(clientEmail, subject, stripHtml_(htmlBody), { htmlBody: htmlBody });
    SpreadsheetApp.getUi().alert('✅ Draft saved in your Gmail, addressed to ' + clientEmail + '.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Could not create Gmail draft. Error: ' + e.message);
  }
}

function buildEmailHtml_(month, year, bom, eom, categories) {
  var rows = '';
  categories.forEach(function(cat) {
    var isOver = cat.status === 'Over';
    var isUnder = cat.status === 'Under';
    var rowBg = (isOver || isUnder) ? '#fce8e6' : '#ffffff';
    var arrow = isOver ? '▲' : (isUnder ? '▼' : '');

    rows += '<tr style="background:' + rowBg + ';">' +
      '<td style="padding:8px 12px;font-weight:bold;">' + escHtml_(cat.name) + '</td>' +
      '<td style="padding:8px 12px;text-align:right;">' + fmtMoney_(cat.budget) + '</td>' +
      '<td style="padding:8px 12px;text-align:right;">' + fmtMoney_(cat.actual) + '</td>' +
      '<td style="padding:8px 12px;text-align:right;color:' + (isOver ? '#c5221f' : '#137333') + ';">' +
      arrow + ' ' + cat.deviationPctStr + '</td></tr>';

    if (cat.items && cat.items.length) {
      rows += '<tr style="background:#fff3e0;"><td colspan="4" style="padding:6px 12px 8px 24px;color:#e65100;font-style:italic;">' +
        '⚠️ ' + escHtml_(cat.name) + ' is ' + (isOver ? 'over' : 'under') + ' budget by ' + cat.deviationPctStr + '.' + escHtml_(cat.staleNote || '');
      cat.items.forEach(function(item) {
        rows += '<br>&nbsp;&nbsp;&nbsp;↳ <strong>' + escHtml_(item.description) + '</strong>: ' +
          fmtMoney_(item.actual) + ' (Actual) vs ' + fmtMoney_(item.budget) + ' (Planned)';
      });
      rows += '</td></tr>';
    }
  });

  return '' +
    '<html><body style="font-family:Arial,sans-serif;color:#202124;max-width:680px;margin:0 auto;">' +
    '<div style="background:#1a73e8;padding:20px 24px;border-radius:8px 8px 0 0;">' +
    '<h2 style="color:#fff;margin:0;">GrupoLyN — Monthly Financial Report</h2>' +
    '<p style="color:#c5dcff;margin:4px 0 0;">' + escHtml_(month + ' ' + year) + ' · ' +
    escHtml_(prettyDate(bom)) + ' – ' + escHtml_(prettyDate(eom)) + '</p></div>' +
    '<p style="padding:16px 4px 0;">Dear Client,</p>' +
    '<p style="padding:0 4px;">Here is your monthly budget comparison. Categories highlighted in red exceeded the ' +
    (DEVIATION_THRESHOLD * 100) + '% deviation threshold and are broken down below.</p>' +
    '<table style="width:100%;border-collapse:collapse;margin:16px 0;">' +
    '<tr style="background:#e8f0fe;"><th style="padding:10px 12px;text-align:left;">Category</th>' +
    '<th style="padding:10px 12px;text-align:right;">Planned</th>' +
    '<th style="padding:10px 12px;text-align:right;">Actual</th>' +
    '<th style="padding:10px 12px;text-align:right;">Deviation</th></tr>' +
    rows + '</table>' +
    '<p style="padding:0 4px;color:#888;font-size:12px;">Generated automatically on ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy') +
    ' by the GrupoLyN financial automation tool.</p></body></html>';
}

function stripHtml_(html) {
  return html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function escHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

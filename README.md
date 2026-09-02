# GrupoLyN Financial Automation — Technical Assessment

**Role:** AI & Automation Specialist  
**Platform:** Google Sheets + Apps Script  
**Author:** Ángel Granados

---

## What this does

Adds a password-protected **🔒 Admin menu** to GrupoLyN's real "Prospr Script" financial planning template, automates the **Monthly Comparative Report** (budget vs. actual, ≥15% deviation flagging with the specific line items responsible), and includes a **mass deployment approach** to roll this out across client spreadsheets.

This was verified end-to-end against the real template: first-time admin code setup, unlocking the panel, generating the report for a real month, and confirming the resulting Gmail draft. Nothing here is unverified except where explicitly labeled "EXPERIMENTAL."

---

## Files

| File | Purpose |
|------|---------|
| `src/Code.gs` | Menu, document-scoped auth + email allow-list + audit log, and the full report engine |
| `src/Deploy.gs` | Mass deployment — Strategy A (proven) + Strategy B (labeled experimental design sketch) |
| `src/appsscript.json` | Manifest — OAuth scopes, V8 runtime |
| `tests/logic.test.js` | Standalone Node tests for the two pure-logic helpers (`toNumber_`, `isTotalRowFor_`) — `node tests/logic.test.js` |

---

## How to install (5 minutes)

1. Open the Google Sheet → **Extensions → Apps Script**
2. **Project Settings (gear icon) → check "Enable Chrome V8 runtime."** The template's pre-existing script was on the deprecated Rhino engine — without this the whole project silently fails to load the menu.
3. Delete `Code.gs`'s default content, paste `src/Code.gs`. Add `src/Deploy.gs` as a second file if you want the bonus deployment tool available.
4. Save (Ctrl+S), refresh the sheet — the **🔒 Admin** menu appears.
5. First unlock: since no code ships hardcoded, the tool walks you through creating one (min. 6 characters) the first time you click "Unlock Admin Panel."

---

## Security & reliability

| Concern | How it's handled |
|---|---|
| Admin password storage | Document-scoped (`DocumentProperties`), shared per spreadsheet copy — the correct scope for a password that applies to everyone using that client's copy. No default ships in code; first use forces a real code to be created. |
| Session handling | Unlock sessions are user-scoped and expire after 30 minutes; every sensitive action re-validates the session server-side before running, so a stale or expired session can't keep operating just because the sidebar is still open. |
| Failed-attempt lockout | 5 wrong attempts locks that browser out for 10 minutes. |
| Optional email allow-list | `Admin → Set Authorized Emails` restricts the menu to specific Google accounts, for clients who want more than "knows the code." Left optional so the tool isn't locked out of the box before anyone configures it. Access decisions use only `Session.getActiveUser()` and fail **closed** (deny) if Apps Script can't determine the active user's identity — they never fall back to the script owner's identity, which would otherwise let an unidentifiable visitor in under the owner's name. |
| Audit trail | Every unlock, failed attempt, code change, allow-list change, and report run writes a row (timestamp, user, action) to a hidden "Admin Audit Log" sheet, shown to admins only via a read-only modal — the sheet itself is never unhidden. |
| Currency-safe parsing | `toNumber_()` strips `$`/`,` formatting before parsing, and also unwraps accounting-style negatives like `"($1,234.56)"`. A cell that still can't be read surfaces a visible warning in the completion alert instead of silently reporting `$0`. |
| Duplicate category names | Category keys are collision-guarded — a repeat name (a real risk given the template's nested "Person 1 / Person 2" sections) is automatically disambiguated and flagged in the warnings list, instead of silently overwriting. |
| "Total"-row detection | A row only counts as a category's total when it's literally `"Total"` or `"Total <that category's own name>"` — matched against the specific category currently being read, not a generic substring test. That keeps a category that legitimately contains the word "Total" (e.g. "Total Rewards") from being mistaken for the total row of a different category. |
| Stale "Total" rows | A category's "Total" row is cross-checked against the sum of its own line items; a mismatch (common after a row insert/delete breaks a `SUM()` range) is flagged in the report rather than trusted blindly. |
| Overwrite protection | Before writing a month's report tab, the tool checks whether a tab of that name already exists and, if so, whether it matches this tool's own report format. If it doesn't (i.e. it's original template content), the existing tab is renamed to `"<name> (original)"` instead of being overwritten, and the completion alert says so. |
| Sidebar cancel behavior | Canceling a nested prompt (e.g. "Change Admin Code" → Cancelar) closes only that dialog — the Admin sidebar stays open. Every sensitive server function signals a real completion explicitly; the client only closes the panel on that signal, not on every round-trip. |
| Report delivery | The Gmail draft looks up the client's email from the "Cover" tab, or prompts once and remembers it. It refuses to run rather than create an unaddressed draft. The draft itself is branded HTML (GrupoLyN header, color-coded deviation table) — the same draft a consultant could forward to a client with zero edits — not a plain-text dump. |

---

## Key design decisions

- **Gmail draft, not auto-send** — financial reports need human review before reaching a client. The advisor reviews the draft and clicks Send.
- **Document-scoped password, user-scoped session** — the code itself is shared across whoever uses that spreadsheet copy (correct for a password); the 30-minute unlocked window is per-browser (so one person unlocking it doesn't unlock it for someone else mid-session).
- **Optional email allow-list** — not forced, so the tool isn't locked out of the box before anyone configures it.
- **Deviation threshold: 15%** — configurable via the `DEVIATION_THRESHOLD` constant in `Code.gs`.

---

## Deployment strategy (Bonus)

**Strategy A — proven, ship today:** Publish the master script as an Apps Script Library. Each client sheet gets a 10-line bootstrap script pasted once — run `showBootstrapSnippet()` from the Apps Script editor (Deploy.gs) to generate the exact snippet with your library's Script ID pre-filled. All future updates to the master library apply instantly to every linked client — this is the one actually recommended for GrupoLyN right now.

**Strategy B — labeled EXPERIMENTAL, design sketch:** Uses the Apps Script API to programmatically update the bound script in every client spreadsheet from a "Clients" ledger sheet. This was never run against a live second spreadsheet during this assessment (no GCP project was provisioned) — rather than ship untested code as if it were finished, it's clearly marked, includes a `dryRun` mode that verifies the targeting/control-flow logic with zero live API calls, merges into (rather than overwrites) any existing client code, and writes every attempt to a "Deployments" ledger so a rollout is auditable, not blind.

There is no supported Apps Script API to look up an existing bound script's ID from a spreadsheet ID alone, so the design tracks each client's resolved script ID explicitly in the "Clients" ledger instead (idempotent: known clients get updated, new ones get created and remembered).

The generated `appsscript.json` uses `libraryId` for the library dependency (the correct field name in the manifest schema) and requests only the OAuth scopes the generated `Setup.gs` actually needs (spreadsheets, Gmail drafts, sidebar UI) rather than a broad Drive scope. It's only written on a client's first deploy — a re-deploy only touches `Setup.gs`, so any scopes or libraries the client's project has since accumulated on its own are never clobbered. `readClientUrls_()` also de-duplicates the "Clients" sheet by URL, so a repeated row can't create two separate bound script projects for the same client.

---

## Base spreadsheet

[GrupoLyN DSM PROSPR Plan](https://docs.google.com/spreadsheets/d/1DmWspcWSL1YCqj2PlEOdrbKYMyvpieZKrKParZi1doM/edit?usp=sharing)

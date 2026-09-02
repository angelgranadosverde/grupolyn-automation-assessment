/**
 * Standalone Node.js tests for the pure-logic helpers in src/Code.gs.
 *
 * These two functions have no dependency on SpreadsheetApp/GmailApp, so they
 * run outside Apps Script exactly as they behave inside it. They are copied
 * here (not imported) because Code.gs is not a Node module — if either
 * implementation changes in Code.gs, mirror the change here too.
 *
 * Run: node tests/logic.test.js
 */

'use strict';
const assert = require('assert');

// ── mirrored from src/Code.gs: toNumber_() ──────────────────────────────────
function toNumber_(v, warnings, cellRef) {
  if (typeof v === 'number') return v;
  if (v === '' || v === null || v === undefined) return 0;
  var str = String(v).trim();
  if (!str || str === '-') return 0;
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

// ── mirrored from src/Code.gs: isTotalRowFor_() (nested inside runMonthlyComparativeReport) ──
function isTotalRowFor_(header, catName) {
  if (!catName) return false;
  return header === 'Total' || header === ('Total ' + catName);
}

// ── minimal test runner (no dependencies) ───────────────────────────────────
let pass = 0;
const failures = [];

function test(name, fn) {
  try {
    fn();
    pass++;
  } catch (err) {
    failures.push({ name, err });
  }
}

// ── toNumber_ ────────────────────────────────────────────────────────────────

test('toNumber_: passes a real number through unchanged', () => {
  assert.strictEqual(toNumber_(42), 42);
});

test('toNumber_: empty/null/undefined all become 0', () => {
  assert.strictEqual(toNumber_(''), 0);
  assert.strictEqual(toNumber_(null), 0);
  assert.strictEqual(toNumber_(undefined), 0);
});

test('toNumber_: a lone "-" (empty cell rendered as a dash) becomes 0', () => {
  assert.strictEqual(toNumber_('-'), 0);
});

test('toNumber_: strips "$" and thousands commas', () => {
  assert.strictEqual(toNumber_('$1,234.56'), 1234.56);
});

test('toNumber_: accounting-style parens are read as negative (the real bug this fixed)', () => {
  assert.strictEqual(toNumber_('($1,234.56)'), -1234.56);
});

test('toNumber_: a plain negative sign still works without parens', () => {
  assert.strictEqual(toNumber_('-50'), -50);
});

test('toNumber_: unparseable text becomes 0 and pushes a warning instead of failing silently', () => {
  const warnings = [];
  const result = toNumber_('n/a', warnings, 'B7');
  assert.strictEqual(result, 0);
  assert.strictEqual(warnings.length, 1);
  assert.ok(warnings[0].includes('B7'), 'warning should reference the cell that failed to parse');
});

// ── isTotalRowFor_ ───────────────────────────────────────────────────────────

test('isTotalRowFor_: a lone "Total" closes the category currently being read', () => {
  assert.strictEqual(isTotalRowFor_('Total', 'Food & Supplies'), true);
});

test('isTotalRowFor_: "Total <category name>" also closes it', () => {
  assert.strictEqual(isTotalRowFor_('Total Food & Supplies', 'Food & Supplies'), true);
});

test('isTotalRowFor_: a category literally named "Total Rewards" is NOT its own total row (the real bug this fixed)', () => {
  assert.strictEqual(isTotalRowFor_('Total Rewards', 'Total Rewards'), false);
});

test('isTotalRowFor_: with no category currently open, nothing counts as a total row', () => {
  assert.strictEqual(isTotalRowFor_('Total', null), false);
  assert.strictEqual(isTotalRowFor_('Total Rewards', null), false);
});

test('isTotalRowFor_: an unrelated header is never mistaken for a total row', () => {
  assert.strictEqual(isTotalRowFor_('Groceries', 'Food & Supplies'), false);
});

test('isTotalRowFor_: an empty header is never a total row', () => {
  assert.strictEqual(isTotalRowFor_('', 'Food & Supplies'), false);
});

// ── report ───────────────────────────────────────────────────────────────────

console.log(`${pass} passed, ${failures.length} failed`);
if (failures.length) {
  failures.forEach(({ name, err }) => {
    console.error(`\nFAIL: ${name}`);
    console.error(err.message);
  });
  process.exit(1);
}

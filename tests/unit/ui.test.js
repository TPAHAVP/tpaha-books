import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fmtMoney, fmtAmount, fmtNumberOrText, fmtDate, todayIso, monthRange, splitPhone } from '../../site/js/ui.js';

test('fmtMoney formats cents with thousands separators and a leading minus', () => {
  assert.equal(fmtMoney(0), '$0.00');
  assert.equal(fmtMoney(123456), '$1,234.56');
  assert.equal(fmtMoney(-1200), '-$12.00');
  assert.equal(fmtMoney(2205562), '$22,055.62');
  assert.equal(fmtAmount(789.9899999999999), '$789.99');
});

test('fmtNumberOrText passes Excel text through and trims float noise', () => {
  assert.equal(fmtNumberOrText('Check fee / qty'), 'Check fee / qty');
  assert.equal(fmtNumberOrText(''), '');
  assert.equal(fmtNumberOrText(null), '');
  assert.equal(fmtNumberOrText(24), '24');
  assert.equal(fmtNumberOrText(24.5), '24.5');
  assert.equal(fmtNumberOrText(0.1 + 0.2), '0.3');
});

test('fmtDate renders ISO dates for people', () => {
  assert.equal(fmtDate('2026-01-16'), '16 Jan');
  assert.equal(fmtDate('2026-01-16', 'long'), '16 Jan 2026');
  assert.equal(fmtDate(''), '');
});

test('todayIso is a valid yyyy-mm-dd string', () => {
  assert.match(todayIso(), /^\d{4}-\d{2}-\d{2}$/);
});

test('monthRange gives the first and last day of a month', () => {
  assert.deepEqual(monthRange(2026, 2), ['2026-02-01', '2026-02-28']);
  assert.deepEqual(monthRange(2028, 2), ['2028-02-01', '2028-02-29']);
  assert.deepEqual(monthRange(2026, 12), ['2026-12-01', '2026-12-31']);
});

test('splitPhone puts each number on its own line like the Print Sheet', () => {
  assert.deepEqual(splitPhone('204-555-0101/204-555-0102'), ['204-555-0101', '204-555-0102']);
  assert.deepEqual(splitPhone('204-555-0100'), ['204-555-0100']);
  assert.deepEqual(splitPhone(''), []);
});

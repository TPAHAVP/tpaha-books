// Monthly report row visibility (docs/workbook-mapping.md §8): a transaction saved from the website must not
// land in a row that Excel still hides. Test mode, synthetic data only.
import { test, expect } from '@playwright/test';

const ACK = () => { try { sessionStorage.setItem('tpaha:writer-ack', '1'); } catch { /* ignore */ } };
const open = async page => {
  await page.addInitScript(ACK);
  await page.goto('/ledger.html');
  await expect(page.locator('#mode-banner')).toContainText('Test mode');
  await expect(page.locator('#txn-table tbody tr[data-id]')).toHaveCount(12);
};
const fillEntry = async (page, over = {}) => {
  const v = { date: '2026-09-06', type: 'Deposit', category: 'Membership', amount: '25', description: 'Report row test', ...over };
  await page.fill('#entry-form [name=date]', v.date);
  await page.selectOption('#entry-form [name=type]', v.type);
  await page.selectOption('#entry-form [name=category]', v.category);
  await page.fill('#entry-form [name=amount]', v.amount);
  await page.fill('#entry-form [name=description]', v.description);
};
const visibleRows = (page, month) => page.evaluate(m => {
  const hidden = new Set(window.__tpahaMock.hiddenRows(m));
  return [...Array(30).keys()].map(i => i + 4).filter(r => !hidden.has(r));
}, month);
const rowsNamed = (page, description) => page.evaluate(d => window.__tpahaMock.table('LOG_Table').rows.filter(r => r[6] === d).length, description);
const writesSinceLoad = page => page.evaluate(() => window.__tpahaMock.log.filter(e => e.method !== 'GET' && !/Session/.test(e.url)).length);
/** Makes the first row-visibility PATCH for a month fail, as a lost connection would. */
const blockVisibility = (page, month) => page.evaluate(m => {
  window.__blocked = false;
  window.__tpahaMock.beforeRespond = async e => {
    if (!window.__blocked && e.method === 'PATCH' && e.url.includes(`worksheets/${m}/range`)) { window.__blocked = true; throw new TypeError('Failed to fetch'); }
  };
}, month);

test('a saved transaction leaves its month showing the populated rows and hiding the rest', async ({ page }) => {
  await open(page);
  expect(await visibleRows(page, 'September')).toHaveLength(30);        // nothing hidden to begin with
  await fillEntry(page);
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  expect(await visibleRows(page, 'September')).toEqual([4]);            // one transaction, one visible report row
  await expect(page.locator('#report-banner')).toBeHidden();            // nothing left to finish
  await expect(page.locator('#halt-banner')).toBeHidden();
  expect(await visibleRows(page, 'January')).toHaveLength(30);          // a month it did not touch is untouched
});

test('deleting a month\'s only transaction hides every report row for that month', async ({ page }) => {
  await open(page);
  await fillEntry(page, { description: 'Only one in September' });
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  expect(await visibleRows(page, 'September')).toEqual([4]);
  await page.click('#txn-table tbody tr[data-id="14"] .btn-delete');
  await page.click('dialog.confirm .btn-danger');
  await expect(page.locator('#txn-table tbody tr[data-id="14"]')).toHaveCount(0);
  expect(await visibleRows(page, 'September')).toEqual([]);
  await expect(page.locator('#report-banner')).toBeHidden();
});

test('a correction that moves a transaction to another month fixes both months', async ({ page }) => {
  await open(page);
  await fillEntry(page, { description: 'Moving month' });
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  expect(await visibleRows(page, 'September')).toEqual([4]);
  await page.click('#txn-table tbody tr[data-id="14"] .btn-edit');
  await page.fill('#edit-form [name=date]', '2026-10-02');
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-dialog')).toBeHidden();
  expect(await visibleRows(page, 'September')).toEqual([]);
  expect(await visibleRows(page, 'October')).toEqual([4]);
  await expect(page.locator('#report-banner')).toBeHidden();
});

test('when the formatting fails: the transaction is saved, the message says so, it survives a reload, nothing runs on its own, and the retry finishes it without a second transaction', async ({ page }) => {
  await open(page);
  await blockVisibility(page, 'September');
  await fillEntry(page, { description: 'Saved but unformatted' });
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  // The transaction landed exactly once, and the screen says what is left.
  expect(await rowsNamed(page, 'Saved but unformatted')).toBe(1);
  await expect(page.locator('#report-banner')).toBeVisible();
  await expect(page.locator('#report-banner')).toContainText('Transaction saved. Excel report formatting still needs updating.');
  await expect(page.locator('#report-banner')).toContainText('September');
  await expect(page.locator('#halt-banner')).toBeHidden();              // formatting is cosmetic; saving is not paused

  await page.reload();
  await expect(page.locator('#txn-table tbody tr[data-id="14"]')).toHaveCount(1);
  await expect(page.locator('#report-banner')).toBeVisible();           // the unfinished work survived
  await expect(page.locator('#report-banner')).toContainText('September');
  expect(await writesSinceLoad(page)).toBe(0);                          // opening the page runs nothing

  await page.click('#report-banner .btn-finish-report');
  await expect(page.locator('#report-banner')).toBeHidden();
  expect(await visibleRows(page, 'September')).toEqual([4]);
  expect(await rowsNamed(page, 'Saved but unformatted')).toBe(1);       // the retry never adds the transaction again
  expect(await page.evaluate(() => window.__tpahaMock.table('LOG_Table').rows.length)).toBe(14);
});

test('the retry never writes the transaction table, whatever else it has to repair', async ({ page }) => {
  await open(page);
  await blockVisibility(page, 'September');
  await fillEntry(page, { description: 'Formatting only' });
  await page.click('#btn-submit');
  await expect(page.locator('#report-banner')).toBeVisible();
  await page.evaluate(() => { window.__mark = window.__tpahaMock.log.length; });
  await page.click('#report-banner .btn-finish-report');
  await expect(page.locator('#report-banner')).toBeHidden();
  const touched = await page.evaluate(() => window.__tpahaMock.log.slice(window.__mark).map(e => `${e.method} ${e.url}`));
  // It may read the transaction table and repair the sorted helper table, but it must never WRITE a transaction.
  expect(touched.filter(u => !u.startsWith('GET ') && /LOG_Table/.test(u))).toEqual([]);
  expect(touched.filter(u => /protection/.test(u))).toEqual([]);        // protection is never touched
  expect(touched.filter(u => /ENTRY/.test(u))).toEqual([]);             // nor the cell holding the password
  expect(touched.some(u => u.startsWith('PATCH') && /September/.test(u))).toBe(true);
});

/** Holds the first row-visibility PATCH open so the page can be reloaded mid-operation. */
const holdVisibility = (page, month) => page.evaluate(m => {
  window.__held = false;
  window.__tpahaMock.beforeRespond = async e => {
    if (!window.__held && e.method === 'PATCH' && e.url.includes(`worksheets/${m}/range`)) { window.__held = true; await new Promise(() => {}); }
  };
}, month);

test('W2: reloading after the transaction lands but before the formatting returns still offers to finish it', async ({ page }) => {
  await open(page);
  await holdVisibility(page, 'September');
  await fillEntry(page, { description: 'Interrupted before formatting' });
  await page.click('#btn-submit');
  await expect.poll(() => page.evaluate(() => window.__held === true)).toBe(true);
  await expect.poll(() => rowsNamed(page, 'Interrupted before formatting')).toBe(1);   // the transaction has landed

  await page.reload();                                                  // the operation never returned
  await expect(page.locator('#txn-table tbody tr[data-id="14"]')).toHaveCount(1);
  await expect(page.locator('#report-banner')).toBeVisible();
  await expect(page.locator('#report-banner')).toContainText('interrupted');
  await expect(page.locator('#report-banner')).toContainText('September');
  expect(await writesSinceLoad(page)).toBe(0);                          // opening the page runs nothing

  await page.click('#report-banner .btn-finish-report');
  await expect(page.locator('#report-banner')).toBeHidden();
  expect(await visibleRows(page, 'September')).toEqual([4]);
  expect(await rowsNamed(page, 'Interrupted before formatting')).toBe(1);
});

test('W2: an ordinary save shows no banner, and a save that cannot be remembered says so', async ({ page }) => {
  await open(page);
  await fillEntry(page, { description: 'Plain save' });
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  await expect(page.locator('#report-banner')).toBeHidden();            // the reservation is cleared on completion

  await page.evaluate(() => {
    const real = Storage.prototype.setItem;
    Storage.prototype.setItem = function (k, v) { if (String(k).startsWith('tpaha:reportfmt:')) throw new DOMException('quota', 'QuotaExceededError'); return real.call(this, k, v); };
  });
  await fillEntry(page, { description: 'Storage refused' });
  await page.click('#btn-submit');
  await expect(page.locator('.toast-error').last()).toContainText('could not remember');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');   // the save still goes through
  expect(await rowsNamed(page, 'Storage refused')).toBe(1);
});

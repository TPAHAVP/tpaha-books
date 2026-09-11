// Treasurer Ledger screen on the in-memory sample workbook (TEST MODE). Nothing here touches Microsoft 365.
import { test, expect } from '@playwright/test';

const open = async page => {
  await page.addInitScript(() => { try { sessionStorage.setItem('tpaha:writer-ack', '1'); } catch { /* ignore */ } });   // pilot rule acknowledged once per session
  await page.goto('/ledger.html');
  await expect(page.locator('#mode-banner')).toContainText('Test mode');
  await expect(page.locator('#txn-table tbody tr[data-id]')).toHaveCount(12);
};
const fillEntry = async (page, over = {}) => {
  const v = { date: '2026-09-06', type: 'Deposit', category: 'Membership', amount: '25.50', description: 'Playwright add', ...over };
  await page.fill('#entry-form [name=date]', v.date);
  await page.selectOption('#entry-form [name=type]', v.type);
  await page.selectOption('#entry-form [name=category]', v.category);
  await page.fill('#entry-form [name=amount]', v.amount);
  await page.fill('#entry-form [name=description]', v.description);
};

test('shows the workbook identity, the records, and the test-mode banner', async ({ page }) => {
  await open(page);
  await expect(page.locator('#workbook-name')).toContainText('TPAHA_2026 (1).xlsx');
  await expect(page.locator('#txn-table tbody tr[data-id="13"]')).toContainText('Sample membership 2');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'idle');
});

test('Save writes the entry to the workbook, confirms after read-back, and the month figures come from the workbook', async ({ page }) => {
  await open(page);
  await fillEntry(page);
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'unsaved');
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  await expect(page.locator('#save-state')).toContainText('#14');
  await expect(page.locator('#txn-table tbody tr[data-id="14"]')).toContainText('Playwright add');
  await expect(page.locator('#entry-form [name=description]')).toHaveValue('');
  await page.click('[data-tab="months"]');
  await page.click('[data-month="9"]');
  await expect(page.locator('#month-view .subtotal [data-cat="Membership"]')).toHaveText('$25.50');
  await expect(page.locator('#month-source')).toContainText('Workbook figures');
  await expect(page.locator('#month-check')).toContainText('Matches');
  await expect(page.locator('#month-view')).not.toContainText('null');
});

test('a double click on Save creates one record', async ({ page }) => {
  await open(page);
  await fillEntry(page, { description: 'Double click' });
  await page.locator('#btn-submit').dblclick();
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  await expect(page.locator('#txn-table tbody tr[data-id="14"]')).toHaveCount(1);
  await expect(page.locator('#txn-table tbody tr[data-id="15"]')).toHaveCount(0);
});

test('a lost response after the write shows "Could not save"; Check workbook (reads only) finds the row, and Save finishes it instead of duplicating it', async ({ page }) => {
  await open(page);
  await page.evaluate(() => window.__tpahaMock.failNext({ match: 'rows/add', networkError: true, afterApply: true }));
  await fillEntry(page, { description: 'Lost response' });
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'failed');
  await expect(page.locator('#save-state')).toContainText('Could not save');
  await expect(page.locator('#btn-verify')).toBeVisible();
  const writes = () => page.evaluate(() => window.__tpahaMock.log.filter(e => e.method !== 'GET' && !/Session/.test(e.url)).length);
  const before = await writes();
  await page.click('#btn-verify');
  await expect(page.locator('#save-state')).toContainText(/in the workbook as #14/);
  expect(await writes()).toBe(before);
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  await expect(page.locator('#txn-table tbody tr[data-id="14"]')).toHaveCount(1);
  await expect(page.locator('#txn-table tbody tr[data-id="15"]')).toHaveCount(0);
});

test('a definite failure offers Retry and does not duplicate', async ({ page }) => {
  await open(page);
  await page.evaluate(() => window.__tpahaMock.failNext({ match: 'rows/add', status: 503, code: 'ServiceUnavailable', innerCode: 'serviceUnavailableUncategorized', retryAfter: 1 }));
  await fillEntry(page, { description: 'Retry me' });
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'failed');
  await expect(page.locator('#btn-retry')).toBeVisible();
  await page.click('#btn-retry');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  await expect(page.locator('#txn-table tbody tr[data-id="14"]')).toHaveCount(1);
});

test('deleting a record someone changed in Excel shows a conflict instead of deleting', async ({ page }) => {
  await open(page);
  await page.evaluate(() => { const t = window.__tpahaMock.table('LOG_Table'); const i = t.rows.findIndex(r => r[0] === 13); t.setCell(i, 6, 'Edited in Excel'); });
  await page.click('#txn-table tbody tr[data-id="13"] .btn-delete');
  await page.click('dialog.confirm .btn-danger');
  await expect(page.locator('dialog.conflict')).toBeVisible();
  await expect(page.locator('dialog.conflict')).toContainText('changed');
  await page.click('dialog.conflict .btn-reload');
  await expect(page.locator('#txn-table tbody tr[data-id="13"]')).toContainText('Edited in Excel');
  await expect(page.locator('#txn-table tbody tr[data-id]')).toHaveCount(12);
});

test('deleting with a current reference removes the record and the month figures follow', async ({ page }) => {
  await open(page);
  await page.click('#txn-table tbody tr[data-id="13"] .btn-delete');
  await page.click('dialog.confirm .btn-danger');
  await expect(page.locator('.toast').last()).toContainText('Deleted #13');
  await expect(page.locator('#txn-table tbody tr[data-id="13"]')).toHaveCount(0);
  await page.click('[data-tab="months"]');
  await page.click('[data-month="8"]');
  await expect(page.locator('#month-view .subtotal [data-cat="Membership"]')).toHaveText('$0.00');
});

test('changes made directly in Excel appear after Refresh', async ({ page }) => {
  await open(page);
  await page.evaluate(() => { const t = window.__tpahaMock.table('LOG_Table'); const i = t.rows.findIndex(r => r[0] === 2); t.setCell(i, 6, 'Changed in Excel'); });
  await expect(page.locator('#txn-table tbody tr[data-id="2"]')).not.toContainText('Changed in Excel');
  await page.click('#btn-refresh');
  await expect(page.locator('#txn-table tbody tr[data-id="2"]')).toContainText('Changed in Excel');
});

test('correcting a record saves only that row and verifies it', async ({ page }) => {
  await open(page);
  await page.click('#txn-table tbody tr[data-id="2"] .btn-edit');
  await expect(page.locator('#edit-dialog')).toBeVisible();
  await page.fill('#edit-form [name=description]', 'Corrected by test');
  await page.fill('#edit-form [name=amount]', '81');
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-dialog')).toBeHidden();
  await expect(page.locator('#txn-table tbody tr[data-id="2"]')).toContainText('Corrected by test');
  await expect(page.locator('#txn-table tbody tr[data-id="2"]')).toContainText('$81.00');
});

test('unsaved edits block sign-out until confirmed, and are reported to the unload guard', async ({ page }) => {
  await open(page);
  await page.fill('#entry-form [name=description]', 'half typed');
  expect(await page.evaluate(() => window.__tpahaHasUnsaved())).toBe(true);
  await page.evaluate(() => { window.__tpahaRequestSignOut(); });   // not awaited: it waits for the dialog
  await expect(page.locator('dialog.confirm')).toContainText('unsaved');
  await page.click('dialog.confirm .btn:not(.btn-danger):not(.btn-primary)');
  await expect(page.locator('#entry-form [name=description]')).toHaveValue('half typed');
  expect(page.url()).toContain('ledger.html');
});

test('prior-year balance is verified before writing and the workbook figures follow', async ({ page }) => {
  await open(page);
  await page.click('[data-tab="settings"]');
  await page.fill('#settings-form [name=priorYearBalance]', '6000');
  await page.click('#btn-save-settings');
  await expect(page.locator('#settings-state')).toHaveAttribute('data-state', 'saved');
  await page.click('[data-tab="months"]');
  await page.click('[data-month="1"]');
  await expect(page.locator('#recon [data-recon="opening"]')).toHaveText('$6,000.00');
});

test('annual view shows the workbook year totals', async ({ page }) => {
  await open(page);
  await page.click('[data-tab="annual"]');
  await expect(page.locator('#annual-view [data-cell="net-year"]')).toHaveText('$340.50');
  await expect(page.locator('#annual-view [data-cell="closing"]')).toHaveText('$5,340.50');
});

test('phone screenshot of each tab', async ({ page }, testInfo) => {
  await open(page);
  for (const tab of ['entry', 'months', 'annual', 'settings']) {
    await page.click(`[data-tab="${tab}"]`);
    if (tab === 'months') await page.click('[data-month="7"]');
    await page.screenshot({ path: testInfo.outputPath(`ledger-${tab}.png`), fullPage: true });
  }
});

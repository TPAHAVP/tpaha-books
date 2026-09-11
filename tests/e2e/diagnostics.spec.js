// Diagnostics page on the in-memory sample workbook (TEST MODE).
import { test, expect } from '@playwright/test';

test('diagnostics identifies the workbook, runs read-only checks and the prove-access self-test', async ({ page }) => {
  await page.goto('/diagnostics.html');
  await expect(page.locator('#diag-mode')).toContainText('Test mode');
  await page.click('#btn-find');
  await expect(page.locator('#candidates tbody tr')).toHaveCount(1);
  await expect(page.locator('#candidates tbody tr')).toContainText('TPAHA_2026 (1).xlsx');
  await page.click('#candidates .btn-choose');
  await expect(page.locator('#checks tbody tr[data-ok="true"]')).toHaveCount(6);
  await expect(page.locator('#checks tbody tr[data-ok="false"]')).toHaveCount(0);
  await expect(page.locator('#btn-selftest')).toBeDisabled();
  await page.fill('#confirm-name', 'TPAHA_2026 (1).xlsx');
  await expect(page.locator('#btn-selftest')).toBeEnabled();
  await page.click('#btn-selftest');
  await expect(page.locator('#selftest-result')).toContainText('restored');
  await expect(page.locator('#selftest-log li')).toHaveCount(8);
  await expect(page.locator('#selftest-log li.fail')).toHaveCount(0);
});

test('the self-test refuses to run when the typed name does not match', async ({ page }) => {
  await page.goto('/diagnostics.html');
  await page.click('#btn-find');
  await page.click('#candidates .btn-choose');
  await page.fill('#confirm-name', 'Some other file.xlsx');
  await expect(page.locator('#btn-selftest')).toBeDisabled();
});

test('the self-test stays disabled while a read-only check fails, and reports what it compared when it runs', async ({ page }) => {
  await page.goto('/diagnostics.html');
  await page.evaluate(() => window.__tpahaMock.table('LOG_Sorted_Table').append([99, 'junk', 46000, 'Deposit', 'Membership', 1, 'junk', '', '']));
  await page.click('#btn-find');
  await page.click('#candidates .btn-choose');
  await expect(page.locator('#checks tbody tr[data-ok="false"]')).toHaveCount(1);
  await page.fill('#confirm-name', 'TPAHA_2026 (1).xlsx');
  await expect(page.locator('#btn-selftest')).toBeDisabled();
  await expect(page.locator('#selftest-gate')).toContainText(/check/i);
});

test('the self-test result lists exactly what was compared', async ({ page }) => {
  await page.goto('/diagnostics.html');
  await page.click('#btn-find');
  await page.click('#candidates .btn-choose');
  await page.fill('#confirm-name', 'TPAHA_2026 (1).xlsx');
  await page.click('#btn-selftest');
  await expect(page.locator('#selftest-result')).toContainText('restored');
  await expect(page.locator('#selftest-compared')).toContainText('LOG values and number formats');
  await expect(page.locator('#selftest-compared')).toContainText('formulas');
  await expect(page.locator('#selftest-compared')).toContainText(/not compared/i);
});

// Review finding V1: the adapter's ordering rule and this page's read-only check must be the same rule.
const diagReady = page => page.evaluate(() => window.__tpahaDiagReady && window.__tpahaDiagReady());

test('V1: Diagnostics accepts the date-then-number order the adapter writes, after a same-day correction', async ({ page }) => {
  await page.goto('/diagnostics.html');
  await expect.poll(() => diagReady(page)).toBe(true);
  await page.evaluate(async () => {
    const { ExcelClient } = await import('/js/workbook/excel-client.js');
    const { LedgerWorkbook, makeMarker } = await import('/js/workbook/ledger-workbook.js');
    const mock = window.__tpahaMock;
    const wb = new LedgerWorkbook(new ExcelClient({ getToken: async () => 'synthetic', driveId: 'TEST', itemId: 'TEST', fetchImpl: (u, o) => mock.fetch(u, o) }));
    await wb.load();
    const base = { date: '2026-05-04', type: 'Deposit', category: 'Membership', amount: 25, chequeNum: '', notes: '' };
    const first = await wb.addTransaction({ ...base, description: 'Same day A' }, { marker: makeMarker() });
    await wb.addTransaction({ ...base, description: 'Same day B' }, { marker: makeMarker() });
    await wb.updateTransaction(first.txn, { description: 'Same day A corrected' });   // moves the lower number to the end of LOG
  });
  await page.click('#btn-find');
  await page.click('.btn-choose');
  await expect(page.locator('#checks tbody tr')).toHaveCount(6);
  await expect(page.locator('#checks tbody tr[data-ok="true"]')).toHaveCount(6);
  await page.fill('#confirm-name', 'TPAHA_2026 (1).xlsx');
  await expect(page.locator('#btn-selftest')).toBeEnabled();
});

test('V1: Diagnostics still rejects a helper table ordered by date alone, with the numbers reversed on one date', async ({ page }) => {
  await page.goto('/diagnostics.html');
  await expect.poll(() => diagReady(page)).toBe(true);
  await page.evaluate(() => {
    const m = window.__tpahaMock;
    const serial = Math.round((Date.UTC(2026, 11, 20) - Date.UTC(1899, 11, 30)) / 86400000);   // later than every sample row
    const higherFirst = [15, '2026-12-20T00:00:00.000000001Z', serial, 'Deposit', 'Membership', 25, 'Higher number first', '', ''];
    const lowerSecond = [14, '2026-12-20T00:00:00.000000002Z', serial, 'Deposit', 'Membership', 25, 'Lower number second', '', ''];
    for (const t of ['LOG_Table', 'LOG_Sorted_Table']) { m.table(t).append(higherFirst); m.table(t).append(lowerSecond); }
  });
  await page.click('#btn-find');
  await page.click('.btn-choose');
  await expect(page.locator('#checks tbody tr[data-ok="false"]')).toHaveCount(1);
  await expect(page.locator('#checks tbody tr[data-ok="false"]')).toContainText('transaction number');
  await page.fill('#confirm-name', 'TPAHA_2026 (1).xlsx');
  await expect(page.locator('#btn-selftest')).toBeDisabled();
});

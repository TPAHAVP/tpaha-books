// The treasurer page has no workbook selection: it opens the configured workbook, and when that cannot be
// opened it says so and stops. TEST MODE: the in-memory sample stands in for the configured workbook.
import { test, expect } from '@playwright/test';

const ACK = () => { try { sessionStorage.setItem('tpaha:writer-ack', '1'); } catch { /* ignore */ } };

test('the Settings tab offers no way to choose or change the workbook', async ({ page }) => {
  await page.addInitScript(ACK);
  await page.goto('/ledger.html');
  await expect(page.locator('#txn-table tbody tr[data-id]')).toHaveCount(12);
  await page.click('[data-tab="settings"]');
  await expect(page.locator('#workbook-info')).toContainText('Workbook');
  expect(await page.locator('#btn-change-workbook').count()).toBe(0);
  expect(await page.locator('#candidates, .btn-choose').count()).toBe(0);
  expect(await page.content()).not.toMatch(/Choose a different workbook|Which workbook should the ledger use/);
});

// The failure is injected on the session request, the first one the page cannot do without (the item-metadata
// read that precedes it is a courtesy the workbook tolerates losing).
for (const [label, failure, reason] of [
  ['refused with 403', { match: 'createSession', status: 403, code: 'AccessDenied', innerCode: ['accessDenied'], message: 'Access denied' }, /does not have access/i],
  ['not found (404)', { match: 'createSession', status: 404, code: 'itemNotFound', innerCode: ['itemNotFound'], message: 'The resource could not be found.' }, /not found|moved|renamed|deleted/i],
]) {
  test(`when the configured workbook is ${label} on the first read, the page shows an access error and searches for nothing`, async ({ page }) => {
    // Injected once: the init script runs again on the reload that "Try again" causes, when the failure must be gone.
    await page.addInitScript(f => {
      try {
        sessionStorage.setItem('tpaha:writer-ack', '1');
        if (!sessionStorage.getItem('tpaha:test-fail-injected')) { sessionStorage.setItem('tpaha:test-fail-injected', '1'); sessionStorage.setItem('tpaha:test-fail-first-request', JSON.stringify(f)); }
      } catch { /* ignore */ }
    }, failure);
    await page.goto('/ledger.html');
    const gate = page.locator('#access-error');
    await expect(gate).toBeVisible();
    await expect(gate).toContainText('Could not open');
    await expect(gate).toContainText(reason);
    await expect(gate).toContainText(String(failure.status));
    await expect(gate).toContainText('does not search for other copies');
    await expect(page.locator('#tabs')).toBeHidden();
    expect(await page.locator('#candidates, .btn-choose, #btn-change-workbook').count()).toBe(0);
    await expect(page.locator('#btn-retry')).toBeVisible();
    const requests = await page.evaluate(() => window.__tpahaMock.log.map(e => e.url));
    expect(requests.some(u => /search|sharedWithMe|\/me\/drive/.test(u))).toBe(false);
    // Try again reloads and, with the failure gone, the ledger opens.
    await page.click('#btn-retry');
    await expect(page.locator('#txn-table tbody tr[data-id]')).toHaveCount(12);
  });
}

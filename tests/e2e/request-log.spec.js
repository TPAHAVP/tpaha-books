// The request log on the ledger page must stay reachable when the clipboard is not: the same entries are shown
// on the page itself, selectable, and offered as a text file. Review finding of 2026-11-03: the clipboard
// fallback pointed at Diagnostics, whose log is a different page's.
import { test, expect } from '@playwright/test';

const open = async (page, { clipboard } = {}) => {
  await page.addInitScript(mode => {
    try { sessionStorage.setItem('tpaha:writer-ack', '1'); } catch { /* ignore */ }
    if (mode === 'reject') Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: () => Promise.reject(new Error('NotAllowedError')) } });
    if (mode === 'missing') Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined });
  }, clipboard);
  await page.goto('/ledger.html');
  await expect(page.locator('#mode-banner')).toContainText('Test mode');
  await expect(page.locator('#txn-table tbody tr[data-id]')).toHaveCount(12);
  await page.click('[data-tab="settings"]');
};

for (const mode of ['reject', 'missing']) {
  test(`when the clipboard ${mode === 'reject' ? 'rejects' : 'is absent'}, Copy shows the same log on the page, selectable, with a download`, async ({ page }) => {
    await open(page, { clipboard: mode });
    await expect(page.locator('#request-log-panel')).toBeHidden();
    await page.click('#btn-copy-log');
    await expect(page.locator('.toast').first()).toContainText('shown below');
    await expect(page.locator('#request-log-panel')).toBeVisible();
    const shown = await page.locator('#request-log-text').inputValue();
    const expected = await page.evaluate(() => window.__tpahaRequestLog());
    expect(shown).toBe(expected);
    expect(shown).toMatch(/GET \/workbook\/tables\/LOG_Table\/dataBodyRange.* → 200/);
    expect(shown).toMatch(/session opened/);
    expect(shown).not.toMatch(/session-\d+/);
    expect(shown).not.toContain('test-mode');
    await expect(page.locator('#request-log-text')).toHaveAttribute('readonly', '');
    const selected = await page.evaluate(() => { const t = document.getElementById('request-log-text'); return t.selectionEnd - t.selectionStart; });
    expect(selected).toBe(shown.length);
    await expect(page.locator('#btn-download-log')).toHaveAttribute('download', /^tpaha-request-log-.*\.txt$/);
    await expect(page.locator('#btn-download-log')).toHaveAttribute('href', /^blob:/);
  });
}

test('Show request log reveals the log without the clipboard and keeps it in step with later requests', async ({ page }) => {
  await open(page);
  await page.click('#btn-show-log');
  await expect(page.locator('#request-log-panel')).toBeVisible();
  const first = await page.locator('#request-log-text').inputValue();
  expect(first).toMatch(/session opened/);
  await page.click('[data-tab="entry"]');
  await page.click('#btn-refresh');
  await expect(page.locator('#txn-table tbody tr[data-id]')).toHaveCount(12);
  await page.click('[data-tab="settings"]');
  await page.click('#btn-show-log');     // hides
  await expect(page.locator('#request-log-panel')).toBeHidden();
  await page.click('#btn-show-log');     // shows again, refreshed
  const second = await page.locator('#request-log-text').inputValue();
  expect(second.length).toBeGreaterThan(first.length);
  expect(second).toBe(await page.evaluate(() => window.__tpahaRequestLog()));
});

test('the downloaded file holds the same text as the page', async ({ page }) => {
  await open(page);
  await page.click('#btn-show-log');
  const [download] = await Promise.all([page.waitForEvent('download'), page.click('#btn-download-log')]);
  expect(download.suggestedFilename()).toMatch(/^tpaha-request-log-.*\.txt$/);
  const path = await download.path();
  const fs = await import('node:fs');
  const text = fs.readFileSync(path, 'utf8');
  expect(text).toBe(await page.locator('#request-log-text').inputValue());
});

// Session errors on the ledger page (TEST MODE, in-memory sample workbook). After the live refusal of 2026-11-03.
import { test, expect } from '@playwright/test';

const open = async page => {
  await page.addInitScript(() => { try { sessionStorage.setItem('tpaha:writer-ack', '1'); } catch { /* ignore */ } });
  await page.goto('/ledger.html');
  await expect(page.locator('#mode-banner')).toContainText('Test mode');
  await expect(page.locator('#txn-table tbody tr[data-id]')).toHaveCount(12);
};
const writes = page => page.evaluate(() => window.__tpahaMock.log.filter(e => e.method !== 'GET' && !/Session/.test(e.url)).length);
const copies = (page, id) => page.evaluate(i => window.__tpahaMock.table('LOG_Table').rows.filter(r => r[0] === i).length, id);
const PAUSED = { status: 400, code: 'InvalidSession', innerCode: ['invalidSessionReCreatable', 'InvalidOrTimedOutSession'], message: 'The target session is invalid.' };

test('a session Microsoft paused while the Correct dialog was open is recovered on the read before the write; the correction saves once', async ({ page }) => {
  await open(page);
  await page.click('#txn-table tbody tr[data-id="2"] .btn-edit');
  await page.evaluate(() => window.__tpahaMock.pauseSession());
  await page.fill('#edit-form [name=amount]', '81');
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-dialog')).not.toBeVisible();
  await expect(page.locator('#txn-table tbody tr[data-id="2"]')).toContainText('$81.00');
  expect(await copies(page, 2)).toBe(1);
});

test('a paused session refusing the append: the dialog shows the refusal, Check workbook finds nothing written, Save writes once', async ({ page }) => {
  await open(page);
  await page.click('#txn-table tbody tr[data-id="2"] .btn-edit');
  await page.evaluate(p => window.__tpahaMock.failNext({ match: 'LOG_Table/rows/add', ...p }), PAUSED);
  await page.fill('#edit-form [name=amount]', '81');
  const before = await writes(page);
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-state')).toHaveAttribute('data-state', 'failed');
  await expect(page.locator('#edit-state')).toContainText('Could not save: The target session is invalid.');
  await expect(page.locator('#edit-state')).not.toContainText('lost');
  expect(await writes(page)).toBe(before + 1);                       // the one refused attempt
  await page.click('.btn-edit-verify');
  await expect(page.locator('#edit-state')).toContainText('This is not in the workbook yet');
  expect(await writes(page)).toBe(before + 1);                       // the check wrote nothing
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-dialog')).not.toBeVisible();
  await expect(page.locator('#txn-table tbody tr[data-id="2"]')).toContainText('$81.00');
  expect(await copies(page, 2)).toBe(1);
});

test('when the check itself cannot read, the dialog says the check failed and the attempt remains unconfirmed', async ({ page }) => {
  await open(page);
  await page.click('#txn-table tbody tr[data-id="2"] .btn-edit');
  await page.evaluate(() => window.__tpahaMock.failNext({ match: 'LOG_Table/rows/add', status: 400, code: 'InvalidSession', innerCode: ['invalidSessionNotFound'], message: 'The target session is invalid.' }));
  await page.fill('#edit-form [name=amount]', '81');
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-state')).toHaveAttribute('data-state', 'failed');
  const sent = await page.evaluate(() => window.__tpahaMock.log.length);
  await page.click('.btn-edit-verify');
  await expect(page.locator('#edit-state')).toContainText('Could not check the workbook');
  await expect(page.locator('#edit-state')).toContainText('remains unconfirmed');
  await expect(page.locator('#edit-state')).not.toContainText('answer from Microsoft was lost');
  expect(await page.evaluate(() => window.__tpahaMock.log.length)).toBe(sent);   // blocked: nothing sent
  await expect(page.locator('.btn-edit-verify')).toBeVisible();                   // the attempt is still pending
  expect(await copies(page, 2)).toBe(1);
});

test('the ledger page keeps a request log that can be copied and never contains a session id', async ({ page }) => {
  await open(page);
  await page.click('[data-tab="settings"]');
  await expect(page.locator('#btn-copy-log')).toBeVisible();
  const text = await page.evaluate(() => window.__tpahaRequestLog());
  expect(text).toMatch(/GET \/workbook\/tables\/LOG_Table\/dataBodyRange.* → 200/);
  expect(text).toMatch(/session opened/);
  expect(text).not.toMatch(/session-\d+/);
  expect(text).not.toContain('test-mode');
});

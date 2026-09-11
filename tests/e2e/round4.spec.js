// Regression tests for the round-4 review finding (U1): correction recovery must keep the reference of the
// version that was actually saved, and reloading a record must never discard typed text without asking.
// Test mode, synthetic data only. Nothing here touches Microsoft 365.
import { test, expect } from '@playwright/test';

const ACK = () => { try { sessionStorage.setItem('tpaha:writer-ack', '1'); } catch { /* ignore */ } };
const open = async page => {
  await page.addInitScript(ACK);
  await page.goto('/ledger.html');
  await expect(page.locator('#mode-banner')).toContainText('Test mode');
  await expect(page.locator('#txn-table tbody tr[data-id]')).toHaveCount(12);
};
/** Every row except the one being corrected, as stored: proves unrelated records are untouched. */
const otherRows = page => page.evaluate(() => JSON.stringify(window.__tpahaMock.table('LOG_Table').rows.filter(r => r[0] !== 2)));
const copiesOf2 = page => page.evaluate(() => window.__tpahaMock.table('LOG_Table').rows.filter(r => r[0] === 2));
const persistedCorrection = page => page.evaluate(() => {
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (!k.startsWith('tpaha:draft:')) continue;
    try { const rec = JSON.parse(localStorage.getItem(k)); if (rec && rec.correction) return rec.correction; } catch { /* ignore */ }
  }
  return null;
});
const writesSinceLoad = page => page.evaluate(() => window.__tpahaMock.log.filter(e => e.method !== 'GET' && !/Session/.test(e.url)).length);
const holdNextAppend = page => page.evaluate(() => {
  window.__held = false;
  window.__tpahaMock.beforeRespond = async e => {
    if (e.method === 'POST' && e.url.includes('LOG_Table/rows/add') && !window.__held) { window.__held = true; await new Promise(r => { window.__release = r; }); }
  };
});

test('U1: text typed while a correction saves is stored with the SAVED row as its reference; after a reload it saves without a false conflict', async ({ page }) => {
  await open(page);
  const before = await otherRows(page);
  await page.click('tr[data-id="2"] .btn-edit');
  await page.fill('#edit-form [name=description]', 'First correction saved');
  await holdNextAppend(page);
  await page.click('#btn-edit-save');
  await expect.poll(() => page.evaluate(() => Boolean(window.__release))).toBe(true);
  await page.fill('#edit-form [name=description]', 'Newer correction awaiting save');
  await page.evaluate(() => window.__release());
  await expect(page.locator('#edit-state')).toHaveAttribute('data-state', 'unsaved');
  await expect.poll(() => page.evaluate(() => window.__tpahaBusy())).toBe(false);
  // The stored reference must be the version that reached the workbook, not the one from before the save.
  const stored = await persistedCorrection(page);
  expect(stored.ref.description).toBe('First correction saved');
  expect(stored.draft.description).toBe('Newer correction awaiting save');
  expect((await copiesOf2(page))[0][6]).toBe('First correction saved');

  await page.reload();
  await expect(page.locator('#edit-dialog')).toBeVisible();
  await expect(page.locator('#edit-form [name=description]')).toHaveValue('Newer correction awaiting save');
  await expect(page.locator('#edit-state')).toHaveAttribute('data-state', 'unsaved');
  expect(await writesSinceLoad(page)).toBe(0);
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-dialog')).toBeHidden();
  await expect(page.locator('#edit-state')).not.toHaveAttribute('data-state', 'conflict');
  const copies = await copiesOf2(page);
  expect(copies).toHaveLength(1);
  expect(copies[0][6]).toBe('Newer correction awaiting save');
  expect(await otherRows(page)).toBe(before);
  await expect(page.locator('#halt-banner')).toBeHidden();
  expect(await persistedCorrection(page)).toBeNull();
});

test('U1: an unconfirmed correction that landed is finished after a reload, and newer text then saves against the finished row', async ({ page }) => {
  await open(page);
  const before = await otherRows(page);
  await page.click('tr[data-id="2"] .btn-edit');
  await page.fill('#edit-form [name=description]', 'Unconfirmed version');
  await page.evaluate(() => {
    window.__tpahaMock.beforeRespond = async e => {
      if (e.method === 'POST' && e.url.includes('LOG_Table/rows/add')) {
        window.__tpahaMock.beforeRespond = null;
        window.__tpahaMock.failNext({ match: 'LOG_Table/dataBodyRange', networkError: true });   // the read-back fails,
        window.__tpahaMock.failNext({ match: 'LOG_Table/dataBodyRange', networkError: true });   // and so does its one retry
      }
    };
  });
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-state')).toHaveAttribute('data-state', 'failed');
  await page.fill('#edit-form [name=description]', 'Newer text after the unconfirmed attempt');
  expect((await persistedCorrection(page)).marker).toBeTruthy();

  await page.reload();
  await expect(page.locator('#edit-dialog')).toBeVisible();
  await expect(page.locator('#edit-form [name=description]')).toHaveValue('Newer text after the unconfirmed attempt');
  await expect(page.locator('#edit-state')).toContainText(/removal of the old copy/);
  expect(await writesSinceLoad(page)).toBe(0);                       // the startup check only reads
  await page.click('#btn-edit-save');                                 // finishes the attempt that landed
  await expect(page.locator('#edit-dialog')).toBeVisible();
  await expect(page.locator('#edit-state')).toHaveAttribute('data-state', 'unsaved');
  let copies = await copiesOf2(page);
  expect(copies).toHaveLength(1);
  expect(copies[0][6]).toBe('Unconfirmed version');
  const stored = await persistedCorrection(page);
  expect(stored.ref.description).toBe('Unconfirmed version');        // the reference followed the finished row
  expect(stored.marker).toBeNull();
  await page.click('#btn-edit-save');                                 // now the newer text
  await expect(page.locator('#edit-dialog')).toBeHidden();
  copies = await copiesOf2(page);
  expect(copies).toHaveLength(1);
  expect(copies[0][6]).toBe('Newer text after the unconfirmed attempt');
  expect(await otherRows(page)).toBe(before);
  await expect(page.locator('#halt-banner')).toBeHidden();
});

test('U1: a real conflict is still reported, and Reload record asks before replacing typed text; keeping it saves onto the row just read', async ({ page }) => {
  await open(page);
  await page.click('tr[data-id="2"] .btn-edit');
  await page.fill('#edit-form [name=description]', 'My typed correction');
  await page.evaluate(() => { const t = window.__tpahaMock.table('LOG_Table'); t.setCell(t.rows.findIndex(r => r[0] === 2), 6, 'Changed in Excel'); });
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-state')).toHaveAttribute('data-state', 'conflict');   // the fingerprint check still works
  expect((await copiesOf2(page))[0][6]).toBe('Changed in Excel');
  await page.click('#edit-actions .btn-reload');
  const choice = page.locator('dialog.reload-choice');
  await expect(choice).toBeVisible();
  await expect(choice).toContainText('Changed in Excel');
  await expect(choice).toContainText('My typed correction');
  await choice.locator('.btn-cancel').click();                        // cancelling changes nothing
  await expect(page.locator('#edit-form [name=description]')).toHaveValue('My typed correction');
  await page.click('#edit-actions .btn-reload');
  await page.locator('dialog.reload-choice .btn-keep-typed').click();
  await expect(page.locator('#edit-form [name=description]')).toHaveValue('My typed correction');
  await expect(page.locator('#edit-state')).toHaveAttribute('data-state', 'unsaved');
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-dialog')).toBeHidden();
  const copies = await copiesOf2(page);
  expect(copies).toHaveLength(1);
  expect(copies[0][6]).toBe('My typed correction');
});

test('U1: Reload record can also replace the typed text with the workbook version, on purpose', async ({ page }) => {
  await open(page);
  await page.click('tr[data-id="2"] .btn-edit');
  await page.fill('#edit-form [name=description]', 'Text I am willing to lose');
  await page.evaluate(() => { const t = window.__tpahaMock.table('LOG_Table'); t.setCell(t.rows.findIndex(r => r[0] === 2), 6, 'Changed in Excel'); });
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-state')).toHaveAttribute('data-state', 'conflict');
  await page.click('#edit-actions .btn-reload');
  await page.locator('dialog.reload-choice .btn-use-workbook').click();
  await expect(page.locator('#edit-form [name=description]')).toHaveValue('Changed in Excel');
  await expect(page.locator('#edit-state')).toHaveAttribute('data-state', 'idle');
  expect(await persistedCorrection(page)).toBeNull();
  expect((await copiesOf2(page))[0][6]).toBe('Changed in Excel');
});

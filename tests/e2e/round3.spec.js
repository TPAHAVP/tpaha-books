// Regression tests for the round-3 review findings (T1–T4), asserting the SAFE outcomes. Test mode, synthetic data only.
import { test, expect } from '@playwright/test';

const ACK = () => { try { sessionStorage.setItem('tpaha:writer-ack', '1'); } catch { /* ignore */ } };
const open = async page => {
  await page.addInitScript(ACK);
  await page.goto('/ledger.html');
  await expect(page.locator('#mode-banner')).toContainText('Test mode');
  await expect(page.locator('#txn-table tbody tr[data-id]')).toHaveCount(12);
};
const fillEntry = async (page, over = {}) => {
  const v = { date: '2026-09-06', type: 'Deposit', category: 'Membership', amount: '25', description: 'Round 3 entry', ...over };
  await page.fill('#entry-form [name=date]', v.date);
  await page.selectOption('#entry-form [name=type]', v.type);
  await page.selectOption('#entry-form [name=category]', v.category);
  await page.fill('#entry-form [name=amount]', v.amount);
  await page.fill('#entry-form [name=description]', v.description);
};
const ids = page => page.evaluate(() => window.__tpahaMock.table('LOG_Table').rows.map(r => r[0]).filter(x => x !== ''));
const writes = page => page.evaluate(() => window.__tpahaMock.log.filter(e => e.method !== 'GET' && !/Session/.test(e.url)).length);
const persistedCorrection = page => page.evaluate(() => {
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (!k.startsWith('tpaha:draft:')) continue;
    try { const rec = JSON.parse(localStorage.getItem(k)); if (rec && rec.correction) return rec.correction; } catch { /* ignore */ }
  }
  return null;
});

test('T1: a second Delete confirmed while the first is still reading is refused; when the first finishes only its row is gone and #3, #4 survive', async ({ page }) => {
  await open(page);
  await page.evaluate(() => {
    window.__tpahaMock.beforeRespond = async e => {
      if (e.method === 'GET' && e.url.includes('LOG_Table/dataBodyRange') && !window.__held) { window.__held = true; await new Promise(r => { window.__release = r; }); }
    };
  });
  await page.click('tr[data-id="2"] .btn-delete');
  await page.click('dialog.confirm .btn-danger');
  await expect.poll(() => page.evaluate(() => Boolean(window.__release))).toBe(true);
  expect(await page.evaluate(() => window.__tpahaBusy())).toBe(true);
  await expect(page.locator('tr[data-id="3"] .btn-delete')).toBeDisabled();
  await page.locator('tr[data-id="3"] .btn-delete').click({ force: true });          // even a forced click is refused
  await expect(page.locator('dialog.confirm')).toHaveCount(0);
  await page.evaluate(() => window.__release());
  await expect(page.locator('#txn-table tbody tr[data-id="2"]')).toHaveCount(0);
  await expect(page.locator('#halt-banner')).toBeHidden();
  const remaining = await ids(page);
  expect(remaining).not.toContain(2);
  expect(remaining).toContain(3);
  expect(remaining).toContain(4);
  expect(remaining).toHaveLength(11);
  expect(await page.evaluate(() => window.__tpahaMock.log.filter(e => e.method === 'DELETE' && /LOG_Table\/rows/.test(e.url)).length)).toBe(1);
  await expect(page.locator('tr[data-id="3"] .btn-delete')).toBeEnabled();         // and it works afterwards, on fresh positions
  await page.click('tr[data-id="3"] .btn-delete');
  await page.click('dialog.confirm .btn-danger');
  await expect(page.locator('#txn-table tbody tr[data-id="3"]')).toHaveCount(0);
  expect(await ids(page)).toContain(4);
});

test('T1: Save is refused while another change is running, and typing stays possible', async ({ page }) => {
  await open(page);
  await page.evaluate(() => {
    window.__tpahaMock.beforeRespond = async e => {
      if (e.method === 'DELETE' && !window.__held) { window.__held = true; await new Promise(r => { window.__release = r; }); }
    };
  });
  await page.click('tr[data-id="13"] .btn-delete');
  await page.click('dialog.confirm .btn-danger');
  await expect.poll(() => page.evaluate(() => Boolean(window.__release))).toBe(true);
  await fillEntry(page, { description: 'typed during a delete' });
  await page.click('#btn-submit');
  await expect(page.locator('.toast').last()).toContainText(/still being saved/);
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'unsaved');
  await page.evaluate(() => window.__release());
  await expect(page.locator('#txn-table tbody tr[data-id="13"]')).toHaveCount(0);
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  expect(await page.evaluate(() => window.__tpahaMock.table('LOG_Table').rows.filter(r => r[6] === 'typed during a delete').length)).toBe(1);
});

test('T2: with the ledger open first, Diagnostics in a second tab keeps its read-only checks but refuses the write test, even when forced', async ({ page, context }) => {
  await open(page);
  expect(await page.evaluate(() => window.__tpahaSecondaryTab())).toBe(false);
  const diag = await context.newPage();
  await diag.goto('/diagnostics.html');
  await expect(diag.locator('#diag-mode')).toContainText(/another TPAHA Books tab/);
  await expect.poll(() => diag.evaluate(() => window.__tpahaDiagReady())).toBe(true);
  await diag.click('#btn-find');
  await diag.click('.btn-choose');
  await expect(diag.locator('#checks tbody tr[data-ok="true"]')).toHaveCount(6);      // reading still works
  await diag.fill('#confirm-name', 'TPAHA_2026 (1).xlsx');
  await expect(diag.locator('#btn-selftest')).toBeDisabled();
  await expect(diag.locator('#selftest-gate')).toContainText(/another tab/);
  await diag.evaluate(() => { const b = document.querySelector('#btn-selftest'); b.disabled = false; b.click(); });   // the handler checks too
  await expect(diag.locator('#selftest-result')).toContainText(/Refused/);
  expect(await writes(diag)).toBe(0);
  expect(await page.evaluate(() => window.__tpahaSecondaryTab())).toBe(false);
});

test('T2: with Diagnostics open first, the ledger in a second tab is the one that refuses to write', async ({ page, context }) => {
  await page.goto('/diagnostics.html');
  await expect(page.locator('#diag-mode')).toContainText('Test mode');
  await expect.poll(() => page.evaluate(() => window.__tpahaDiagReady && window.__tpahaDiagReady())).toBe(true);
  expect(await page.evaluate(() => window.__tpahaDiagSecondary())).toBe(false);
  const ledger = await context.newPage();
  await ledger.addInitScript(ACK);
  await ledger.goto('/ledger.html');
  await expect(ledger.locator('#pilot-banner')).toContainText(/another tab/);
  expect(await ledger.evaluate(() => window.__tpahaSecondaryTab())).toBe(true);
  await fillEntry(ledger, { description: 'from the second tab' });
  await ledger.click('#btn-submit');
  await expect(ledger.locator('.toast').last()).toContainText(/another tab/);
  expect(await ledger.evaluate(() => window.__tpahaMock.table('LOG_Table').rows.filter(r => r[6] === 'from the second tab').length)).toBe(0);
  await page.click('#btn-find');
  await page.click('.btn-choose');
  await page.fill('#confirm-name', 'TPAHA_2026 (1).xlsx');
  await expect(page.locator('#btn-selftest')).toBeEnabled();                          // the first tab still owns writing
});

test('T3: with the wall clock frozen, two distinct identical Saves create two rows with distinct nine-digit operation ids, and a retried lost attempt still makes one row', async ({ page }) => {
  await open(page);
  await page.clock.setFixedTime(new Date('2026-09-07T12:00:00.000Z'));
  for (let i = 0; i < 2; i++) {
    await fillEntry(page, { description: 'Two legitimate identical deposits' });
    await page.click('#btn-submit');
    await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
    await expect(page.locator('#entry-form [name=description]')).toHaveValue('');
  }
  const rows = await page.evaluate(() => window.__tpahaMock.table('LOG_Table').rows.filter(r => r[6] === 'Two legitimate identical deposits'));
  expect(rows).toHaveLength(2);
  expect(rows.map(r => r[0]).sort()).toEqual([14, 15]);
  for (const r of rows) expect(String(r[1])).toMatch(/^2026-09-07T12:00:00\.000\d{6}Z$/);   // the id the real Save button wrote
  expect(rows[0][1]).not.toBe(rows[1][1]);
  await page.evaluate(() => window.__tpahaMock.failNext({ match: 'LOG_Table/rows/add', networkError: true, afterApply: true }));
  await fillEntry(page, { description: 'Lost then retried at the same clock value' });
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'failed');
  await page.click('#btn-retry');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  expect(await page.evaluate(() => window.__tpahaMock.table('LOG_Table').rows.filter(r => r[6] === 'Lost then retried at the same clock value').length)).toBe(1);
  await expect(page.locator('#txn-table tbody tr[data-id="16"]')).toHaveCount(1);
  await expect(page.locator('#txn-table tbody tr[data-id="17"]')).toHaveCount(0);
});

test('T4: a correction typed while an entry is saving stays persisted after the entry succeeds, and reopens with its exact text after a reload', async ({ page }) => {
  await open(page);
  await fillEntry(page, { description: 'Entry saving while a correction is typed' });
  await page.evaluate(() => {
    window.__tpahaMock.beforeRespond = async e => {
      if (e.method === 'POST' && e.url.includes('LOG_Table/rows/add') && !window.__held) { window.__held = true; await new Promise(r => { window.__release = r; }); }
    };
  });
  await page.click('#btn-submit');
  await expect.poll(() => page.evaluate(() => Boolean(window.__release))).toBe(true);
  await page.click('tr[data-id="2"] .btn-edit');                                    // opening and typing are allowed while another change runs
  await page.fill('#edit-form [name=description]', 'Separate correction must survive');
  await expect.poll(async () => (await persistedCorrection(page))?.draft?.description).toBe('Separate correction must survive');
  await page.evaluate(() => window.__release());
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  await expect(page.locator('#edit-form [name=description]')).toHaveValue('Separate correction must survive');
  const after = await persistedCorrection(page);
  expect(after?.draft?.description).toBe('Separate correction must survive');       // the entry's success did not erase it
  expect(after?.ref?.id).toBe(2);
  await page.reload();
  await expect(page.locator('#edit-dialog')).toBeVisible();
  await expect(page.locator('#edit-title')).toContainText('#2');
  await expect(page.locator('#edit-form [name=description]')).toHaveValue('Separate correction must survive');
  await expect(page.locator('#edit-state')).toHaveAttribute('data-state', 'unsaved');
  expect(await page.evaluate(() => window.__tpahaHasUnsaved())).toBe(true);
});

// Note: this test does NOT perform an entry Save. Since T1 the page refuses a second write while one runs, and a
// pending correction keeps its modal dialog open, so the UI can no longer produce that overlap. The composition of a
// saved entry with a pending correction is proven in tests/unit/drafts.test.js ("T4: combineDrafts keeps each form
// part on its own"), and the typed-correction case through the UI is the test above.
test('T4: a pending (unconfirmed) correction survives a reload and a cancelled discard, and stays persisted', async ({ page }) => {
  await open(page);
  await page.click('tr[data-id="2"] .btn-edit');
  await page.fill('#edit-form [name=description]', 'Pending correction must survive');
  await page.evaluate(() => {
    window.__tpahaMock.beforeRespond = async e => {
      if (e.method === 'POST' && e.url.includes('LOG_Table/rows/add')) {
        window.__tpahaMock.beforeRespond = null;
        window.__tpahaMock.failNext({ match: 'LOG_Table/dataBodyRange', networkError: true });
        window.__tpahaMock.failNext({ match: 'LOG_Table/dataBodyRange', networkError: true });
      }
    };
  });
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-state')).toHaveAttribute('data-state', 'failed');
  expect((await persistedCorrection(page))?.marker).toBeTruthy();
  await page.keyboard.press('Escape');                                              // discard is asked, not done
  await expect(page.locator('dialog.confirm')).toContainText(/never confirmed/);
  await page.click('dialog.confirm .btn:not(.btn-danger):not(.btn-primary)');
  await expect(page.locator('#edit-dialog')).toBeVisible();
  // Reload: the pending correction reopens, is checked read-only, and stays persisted through a cancelled discard.
  await page.reload();
  await expect(page.locator('#edit-dialog')).toBeVisible();
  await expect(page.locator('#edit-state')).toContainText(/removal of the old copy/);
  expect((await persistedCorrection(page))?.marker).toBeTruthy();
  await page.keyboard.press('Escape');
  await page.click('dialog.confirm .btn:not(.btn-danger):not(.btn-primary)');
  expect((await persistedCorrection(page))?.marker).toBeTruthy();
});

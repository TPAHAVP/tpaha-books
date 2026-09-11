// Regression tests for the Checkpoint 1 review findings (R3, R6) and the re-review findings (S1–S4), asserting the SAFE outcomes.
// Test mode keeps the in-memory workbook in sessionStorage, so a page.reload() sees the same workbook state, as the real one would.
import { test, expect } from '@playwright/test';

const ACK = () => { try { sessionStorage.setItem('tpaha:writer-ack', '1'); } catch { /* ignore */ } };
const open = async (page, { ack = true } = {}) => {
  if (ack) await page.addInitScript(ACK);
  await page.goto('/ledger.html');
  await expect(page.locator('#mode-banner')).toContainText('Test mode');
  await expect(page.locator('#txn-table tbody tr[data-id]')).toHaveCount(12);
};
const fillEntry = async (page, over = {}) => {
  const v = { date: '2026-09-06', type: 'Deposit', category: 'Membership', amount: '25', description: 'Regression entry', ...over };
  await page.fill('#entry-form [name=date]', v.date);
  await page.selectOption('#entry-form [name=type]', v.type);
  await page.selectOption('#entry-form [name=category]', v.category);
  await page.fill('#entry-form [name=amount]', v.amount);
  await page.fill('#entry-form [name=description]', v.description);
};
const countRows = (page, description) => page.evaluate(d => window.__tpahaMock.table('LOG_Table').rows.filter(r => r[6] === d).length, description);
/** Writes seen by the mock since the page loaded (session creation excluded: it is needed for reads). */
const writesSinceLoad = page => page.evaluate(() => window.__tpahaMock.log.filter(e => e.method !== 'GET' && !/Session/.test(e.url)).map(e => `${e.method} ${e.url.replace(/^.*\/workbook\//, '')}`));
const loseAppend = page => page.evaluate(() => window.__tpahaMock.failNext({ match: 'LOG_Table/rows/add', networkError: true, afterApply: true }));

test('R3: pressing the ordinary Save again after a lost response checks the workbook first and does not duplicate', async ({ page }) => {
  await open(page);
  await loseAppend(page);
  await fillEntry(page, { description: 'Lost then saved again' });
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'failed');
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  expect(await countRows(page, 'Lost then saved again')).toBe(1);
  await expect(page.locator('#txn-table tbody tr[data-id="14"]')).toHaveCount(1);
  await expect(page.locator('#txn-table tbody tr[data-id="15"]')).toHaveCount(0);
});

test('R3: edits made after a lost response do not create a second identity; the earlier attempt is finished first and the newer text stays unsaved', async ({ page }) => {
  await open(page);
  await loseAppend(page);
  await fillEntry(page, { description: 'Landed first version' });
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'failed');
  await page.fill('#entry-form [name=description]', 'Edited after the failure');
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'unsaved');
  await expect(page.locator('#entry-form [name=description]')).toHaveValue('Edited after the failure');
  expect(await countRows(page, 'Landed first version')).toBe(1);
  expect(await countRows(page, 'Edited after the failure')).toBe(0);
  await expect(page.locator('#txn-table tbody tr[data-id="14"]')).toHaveCount(1);
});

test('S3: "Check workbook" only reads; a landed-but-unfinished entry is reported as not finished, and Save finishes it without a second row', async ({ page }) => {
  await open(page);
  await loseAppend(page);
  await fillEntry(page, { description: 'Check reads only' });
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'failed');
  const before = (await writesSinceLoad(page)).length;
  await page.click('#btn-verify');
  await expect(page.locator('#save-state')).toContainText(/in the workbook as #14/);
  await expect(page.locator('#save-state')).toContainText(/not finished|not confirmed/i);
  expect((await writesSinceLoad(page)).length).toBe(before);               // zero writes while checking
  await expect(page.locator('#btn-submit')).toHaveText(/Finish/);
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  expect(await countRows(page, 'Check reads only')).toBe(1);
  await expect(page.locator('#txn-table tbody tr[data-id="15"]')).toHaveCount(0);
});

test('S2: a newer draft survives a reload when the older attempt had landed: startup checking writes nothing, the newer text stays unsaved, and each version is saved once', async ({ page }) => {
  await open(page);
  await loseAppend(page);
  await fillEntry(page, { description: 'Earlier saved version' });
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'failed');
  await page.fill('#entry-form [name=description]', 'Later unsaved version');
  await page.reload();
  await expect(page.locator('#txn-table tbody tr[data-id="14"]')).toHaveCount(1);    // the landed row survived the reload, as in the real workbook
  await expect(page.locator('#entry-form [name=description]')).toHaveValue('Later unsaved version');
  await expect(page.locator('#save-state')).toContainText(/in the workbook as #14/);
  await expect(page.locator('#save-state')).not.toHaveAttribute('data-state', 'saved');
  expect(await page.evaluate(() => window.__tpahaHasUnsaved())).toBe(true);
  expect(await writesSinceLoad(page)).toEqual([]);                                  // S3: the automatic check after a reload only reads
  await page.click('#btn-submit');                                                   // finishes the earlier attempt
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'unsaved');
  await expect(page.locator('#entry-form [name=description]')).toHaveValue('Later unsaved version');
  expect(await countRows(page, 'Earlier saved version')).toBe(1);
  expect(await countRows(page, 'Later unsaved version')).toBe(0);
  await page.click('#btn-submit');                                                   // the newer draft is its own operation
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  expect(await countRows(page, 'Earlier saved version')).toBe(1);
  expect(await countRows(page, 'Later unsaved version')).toBe(1);
  await expect(page.locator('#txn-table tbody tr[data-id="15"]')).toHaveCount(1);
});

test('S4: an unconfirmed correction survives a reload: the dialog reopens, checking only reads, and Finish removes the old copy', async ({ page }) => {
  await open(page);
  await page.click('tr[data-id="2"] .btn-edit');
  await page.fill('#edit-form [name=description]', 'Corrected once');
  await page.evaluate(() => {
    window.__tpahaMock.beforeRespond = async e => {
      if (e.method === 'POST' && e.url.includes('LOG_Table/rows/add')) {
        window.__tpahaMock.beforeRespond = null;
        window.__tpahaMock.failNext({ match: 'LOG_Table/dataBodyRange', networkError: true });   // the read-back after the append fails,
        window.__tpahaMock.failNext({ match: 'LOG_Table/dataBodyRange', networkError: true });   // and so does its one retry
      }
    };
  });
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-state')).toHaveAttribute('data-state', 'failed');
  await page.reload();
  await expect(page.locator('#edit-dialog')).toBeVisible();
  await expect(page.locator('#edit-title')).toContainText('#2');
  await expect(page.locator('#edit-form [name=description]')).toHaveValue('Corrected once');
  await expect(page.locator('#edit-state')).toContainText(/removal of the old copy/);
  await expect(page.locator('#halt-banner')).toBeHidden();                          // a known in-progress correction is not reported as damage
  expect(await writesSinceLoad(page)).toEqual([]);
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-dialog')).toBeHidden();
  expect(await page.evaluate(() => window.__tpahaMock.table('LOG_Table').rows.filter(r => r[0] === 2).length)).toBe(1);
  expect(await page.evaluate(() => window.__tpahaMock.table('LOG_Table').rows.find(r => r[0] === 2)[6])).toBe('Corrected once');
});

test('R6: edits typed while a correction is saving stay on screen and unsaved; the second save applies them', async ({ page }) => {
  await open(page);
  await page.click('tr[data-id="2"] .btn-edit');
  await page.fill('#edit-form [name=description]', 'First correction');
  await page.evaluate(() => {
    window.__blocked = false;
    window.__tpahaMock.beforeRespond = async e => {
      if (e.method === 'POST' && e.url.includes('LOG_Table/rows/add') && !window.__blocked) { window.__blocked = true; await new Promise(r => { window.__release = r; }); }
    };
  });
  await page.click('#btn-edit-save');
  await expect.poll(() => page.evaluate(() => Boolean(window.__release))).toBe(true);
  await page.fill('#edit-form [name=description]', 'Second correction still unsaved');
  await page.evaluate(() => window.__release());
  await expect(page.locator('#edit-state')).toHaveAttribute('data-state', 'unsaved');
  await expect(page.locator('#edit-dialog')).toBeVisible();
  await expect(page.locator('#edit-form [name=description]')).toHaveValue('Second correction still unsaved');
  expect(await page.evaluate(() => window.__tpahaHasUnsaved())).toBe(true);
  expect(await page.evaluate(() => window.__tpahaMock.table('LOG_Table').rows.find(r => r[0] === 2)[6])).toBe('First correction');
  await page.click('#btn-edit-save');
  await expect(page.locator('#edit-dialog')).toBeHidden();
  expect(await page.evaluate(() => window.__tpahaMock.table('LOG_Table').rows.find(r => r[0] === 2)[6])).toBe('Second correction still unsaved');
  expect(await page.evaluate(() => window.__tpahaMock.table('LOG_Table').rows.filter(r => r[0] === 2).length)).toBe(1);
});

test('R6: a balance typed while the previous balance is saving stays unsaved and visible', async ({ page }) => {
  await open(page);
  await page.click('[data-tab="settings"]');
  await page.fill('#settings-form [name=priorYearBalance]', '6000');
  await page.evaluate(() => {
    window.__blocked = false;
    window.__tpahaMock.beforeRespond = async e => {
      if (e.method === 'PATCH' && e.url.includes('ENTRY') && !window.__blocked) { window.__blocked = true; await new Promise(r => { window.__release = r; }); }
    };
  });
  await page.click('#btn-save-settings');
  await expect.poll(() => page.evaluate(() => Boolean(window.__release))).toBe(true);
  await page.fill('#settings-form [name=priorYearBalance]', '7000');
  await page.evaluate(() => window.__release());
  await expect(page.locator('#settings-state')).toHaveAttribute('data-state', 'unsaved');
  await expect(page.locator('#settings-form [name=priorYearBalance]')).toHaveValue('7000');
  expect(await page.evaluate(() => window.__tpahaHasUnsaved())).toBe(true);
  expect(await page.evaluate(() => window.__tpahaMock.getCell('ENTRY', 'B23'))).toBe(6000);
});

test('S1/S4: a misdirected delete names the row and pauses saving; Refresh and a page reload keep the pause; only a verified explicit restore lifts it', async ({ page }) => {
  await open(page);
  await page.evaluate(() => {
    window.__once = false;
    window.__tpahaMock.beforeRespond = async e => { if (!window.__once && e.method === 'DELETE') { window.__once = true; window.__tpahaMock.table('LOG_Table').remove(0); } };   // a shift at the last instant
  });
  await page.click('#txn-table tbody tr[data-id="7"] .btn-delete');
  await page.click('dialog.confirm .btn-danger');
  await expect(page.locator('#halt-banner')).toBeVisible();
  await expect(page.locator('#halt-banner')).toContainText('paused');
  await expect(page.locator('#halt-banner')).toContainText('#8');
  await expect(page.locator('#halt-banner')).toContainText(/not recreated/i);
  await fillEntry(page, { description: 'blocked while paused' });
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'failed');
  await expect(page.locator('#save-state')).toContainText(/paused/i);
  expect(await countRows(page, 'blocked while paused')).toBe(0);
  await page.click('#btn-refresh');                                                  // S4: Refresh does not lift the pause
  await expect(page.locator('#halt-banner')).toBeVisible();
  await expect(page.locator('#halt-banner')).toContainText('paused');
  await page.reload();                                                               // S4: nor does a page reload
  await expect(page.locator('#halt-banner')).toBeVisible();
  await expect(page.locator('#halt-banner')).toContainText('paused');
  await expect(page.locator('#txn-table tbody tr[data-id="8"]')).toHaveCount(0);
  await fillEntry(page, { description: 'still blocked' });
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toContainText(/paused/i);
  expect(await countRows(page, 'still blocked')).toBe(0);
  await page.click('#halt-banner .btn-check-resolution');                            // read-only check: not resolved yet
  await expect(page.locator('#incident-checks li[data-ok="false"]')).toHaveCount(1);
  await expect(page.locator('#halt-banner')).toContainText('paused');
  await page.click('#halt-banner [data-repair^="restore:"]');                        // explicit, confirmed restore
  await page.click('dialog.confirm .btn-primary');
  await expect(page.locator('#halt-banner')).toContainText(/Resolved/);
  await expect(page.locator('#halt-banner')).toContainText(/verified/);
  await expect(page.locator('#txn-table tbody tr[data-id="8"]')).toHaveCount(1);
  await expect(page.locator('#txn-table tbody tr[data-id="7"]')).toHaveCount(1);
  await page.click('#halt-banner .btn-dismiss');
  await expect(page.locator('#halt-banner')).toBeHidden();
  await page.click('#btn-submit');
  await expect(page.locator('#save-state')).toHaveAttribute('data-state', 'saved');
  expect(await countRows(page, 'still blocked')).toBe(1);
});

test('S4: duplicate copies already in the workbook pause saving on load; removing one identical copy is explicit and verified', async ({ page }) => {
  await open(page);
  await page.evaluate(() => { const t = window.__tpahaMock.table('LOG_Table'); t.append(t.rows.find(r => r[0] === 5)); });   // pasted in Excel
  await page.click('#btn-refresh');
  await expect(page.locator('#halt-banner')).toContainText('paused');
  await expect(page.locator('#halt-banner')).toContainText('more than one copy');
  await expect(page.locator('#halt-banner')).toContainText('identical copies');
  await page.click('#halt-banner [data-repair^="remove:"]');
  await page.click('dialog.confirm .btn-danger');
  await expect(page.locator('#halt-banner')).toContainText(/Resolved/);
  expect(await page.evaluate(() => window.__tpahaMock.table('LOG_Table').rows.filter(r => r[0] === 5).length)).toBe(1);
});

test('pilot writer model: the first Correct or Delete of a session asks the member to confirm they are the only writer', async ({ page }) => {
  await open(page, { ack: false });
  await expect(page.locator('#pilot-banner')).toContainText(/one designated writer/i);
  await page.click('tr[data-id="2"] .btn-edit');
  await expect(page.locator('dialog.confirm')).toContainText('One writer at a time');
  await page.click('dialog.confirm .btn:not(.btn-primary):not(.btn-danger)');       // Cancel: nothing opens
  await expect(page.locator('#edit-dialog')).toBeHidden();
  await page.click('tr[data-id="2"] .btn-edit');
  await page.click('dialog.confirm .btn-primary');
  await expect(page.locator('#edit-dialog')).toBeVisible();
  await page.click('#btn-edit-cancel');
  await page.click('tr[data-id="3"] .btn-edit');                                     // asked once per session
  await expect(page.locator('#edit-dialog')).toBeVisible();
  await expect(page.locator('dialog.confirm')).toHaveCount(0);
});

test('pilot writer model: a second tab of the same browser is told it will not save, and its Save does nothing', async ({ page, context }) => {
  await open(page);
  const second = await context.newPage();
  await second.goto('/ledger.html');
  await expect(second.locator('#pilot-banner')).toContainText(/another tab/);
  expect(await second.evaluate(() => window.__tpahaSecondaryTab())).toBe(true);
  await fillEntry(second, { description: 'from the second tab' });
  await second.click('#btn-submit');
  await expect(second.locator('.toast').last()).toContainText(/another tab/);
  expect(await countRows(second, 'from the second tab')).toBe(0);
  expect(await page.evaluate(() => window.__tpahaSecondaryTab())).toBe(false);
});

test('annual view shows whether the workbook totals agree with the app calculation', async ({ page }) => {
  await open(page);
  await page.click('[data-tab="annual"]');
  await expect(page.locator('#annual-check')).toContainText('Matches');
});

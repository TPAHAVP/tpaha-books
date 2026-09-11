// Diagnostics: identify the workbook, run read-only checks, and (on a test copy) prove a write/read-back/restore cycle.
// This page can write (the connection test), so it takes part in the same one-writer tab guard as the ledger
// (review finding T2): a second TPAHA Books tab in this browser can read here but never runs the write test.
import { CONFIG } from './config.js';
import { initAuth } from './auth.js';
import { ExcelClient } from './workbook/excel-client.js';
import { MockWorkbook } from './workbook/mock-excel.js';
import { SAMPLE_WORKBOOK } from './workbook/sample-workbook.js';
import { LedgerWorkbook, LOG_COLUMNS, isBlankRow, sortedTableMatches } from './workbook/ledger-workbook.js';
import { findCandidates, rememberWorkbook, resolveConfigured } from './workbook/locate.js';
import { createTabGuard } from './save/tab-guard.js';
import { h, toast } from './ui.js';

const $ = sel => document.querySelector(sel);
const GRAPH = 'https://graph.microsoft.com/v1.0';
const SORTED_CHECK = 'LOG_Sorted matches LOG sorted by date, then transaction number';
const state = { auth: null, mode: 'test', mock: null, client: null, wb: null, chosen: null, who: null, requests: [], checksOk: null, tabGuard: null, tabReady: false, secondaryTab: false, tabUnsupported: false, running: false };

function logRequest(e) {
  state.requests.push(e);
  const li = h('li', {}, e.event ? `session ${e.id}` : `${e.method} ${e.path} → ${e.status || 'network error'} (${e.ms} ms)${e.requestId ? ' request-id ' + e.requestId : ''}${e.error ? ' ' + e.error : ''}`);
  $('#request-log').append(li);
}

async function main() {
  $('#app-version').textContent = `v${CONFIG.appVersion}`;
  state.auth = await initAuth(CONFIG);
  if (state.auth.mode === 'local') {
    state.mode = 'test';
    state.mock = new MockWorkbook(SAMPLE_WORKBOOK);
    window.__tpahaMock = state.mock;
    $('#diag-mode').textContent = 'Test mode (in-memory sample workbook)';
    $('#signin-info').textContent = 'Test mode: clientId is empty in config.js, so no Microsoft sign-in happens and an in-memory sample workbook is used.';
  } else {
    state.mode = 'connected';
    const a = state.auth.account;
    $('#diag-mode').textContent = a ? `Connected as ${a.username}` : 'Not signed in';
    $('#signin-info').textContent = a ? `Signed in as ${a.name} (${a.username}), tenant ${a.tenantId || CONFIG.tenantId}.` : 'Sign in with your TPAHA Microsoft account to continue.';
    $('#signin-actions').append(a
      ? h('button', { type: 'button', className: 'btn', onClick: () => state.auth.signOut() }, 'Sign out')
      : h('button', { type: 'button', className: 'btn btn-primary', onClick: () => state.auth.signIn() }, 'Sign in with Microsoft'));
    if (a) state.who = { tenantId: state.auth.tenantId || a.tenantId || 'tenant', homeAccountId: a.homeAccountId || a.username, app: 'ledger' };
  }
  window.__tpahaDiagSecondary = () => state.secondaryTab;
  window.__tpahaDiagReady = () => state.tabReady;
  $('#btn-find').addEventListener('click', find);
  $('#confirm-name').addEventListener('input', updateGate);
  updateGate();
  $('#btn-selftest').addEventListener('click', runSelfTest);
  $('#btn-copy-log').addEventListener('click', async () => {
    const text = [...$('#selftest-log').children, ...$('#request-log').children].map(li => li.textContent).join('\n');
    try { await navigator.clipboard.writeText(text); toast('Log copied'); } catch { toast('Could not copy automatically; select the text and copy it.', 'error'); }
  });
  // Same writer guard as the ledger: the first TPAHA Books tab in this browser owns writing. Reading works meanwhile;
  // the write test stays off until the guard has answered.
  state.tabGuard = createTabGuard();
  const tab = await state.tabGuard.start();
  state.secondaryTab = !tab.primary;
  state.tabUnsupported = Boolean(tab.unsupported);
  state.tabReady = true;
  if (state.secondaryTab) $('#diag-mode').textContent += ' · another TPAHA Books tab is open: this page reads only';
  else if (state.tabUnsupported) $('#diag-mode').textContent += ' · this browser cannot detect other tabs';
  updateGate();
}

async function graphGet(path) {
  const token = await state.auth.getToken();
  const res = await fetch(path.startsWith('http') ? path : GRAPH + path, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' } });
  if (!res.ok) throw new Error(`Graph ${res.status}`);
  return res.json();
}

async function find() {
  if (state.running) { toast('Wait for the running connection test to finish before changing the workbook.', 'error'); return; }
  const host = $('#candidate-host');
  host.replaceChildren(h('p', { className: 'muted' }, 'Searching…'));
  let cands;
  if (state.mode === 'test') {
    cands = [{ driveId: 'TEST', itemId: 'TEST', name: state.mock.name, extension: 'xlsx', path: '/Documents (sample)', size: 104054, lastModified: state.mock.lastModified, webUrl: '', source: 'mine', sharedBy: '' }];
    cands.warnings = [];
  } else {
    if (!state.auth.account) { host.replaceChildren(h('p', { className: 'warn' }, 'Sign in first.')); return; }
    const pinned = resolveConfigured(CONFIG, 'ledger');
    cands = await findCandidates(graphGet, { query: CONFIG.workbookSearch });
    if (pinned) cands.unshift({ ...pinned, extension: 'xlsx', path: '(pinned in config.js)', source: 'config', lastModified: '', size: null, warnings: [] });
  }
  if (!cands.length) { host.replaceChildren(h('p', { className: 'warn' }, 'No .xlsx workbook found.')); return; }
  host.replaceChildren(h('table', { id: 'candidates' },
    h('thead', {}, h('tr', {}, h('th', {}, 'Name'), h('th', {}, 'Where'), h('th', {}, 'Last changed'), h('th', { className: 'num' }, 'Size'), h('th'))),
    h('tbody', {}, cands.map(c => h('tr', {},
      h('td', {}, c.name, c.extension !== 'xlsx' ? h('span', { className: 'warn' }, ' (not .xlsx)') : null),
      h('td', {}, c.source === 'shared' ? `Shared by ${c.sharedBy || 'someone'}` : c.source === 'config' ? 'Pinned in config.js' : `My OneDrive ${c.path}`),
      h('td', {}, c.lastModified ? new Date(c.lastModified).toLocaleString() : ''),
      h('td', { className: 'num' }, c.size ? `${Math.round(c.size / 1024)} KB` : ''),
      h('td', {}, h('button', { type: 'button', className: 'btn btn-sm btn-primary btn-choose', onClick: () => choose(c) }, 'Choose')))))));
  for (const w of cands.warnings || []) host.append(h('p', { className: 'warn small' }, w));
}

async function choose(c) {
  if (state.running) { toast('Wait for the running connection test to finish before changing the workbook.', 'error'); return; }
  state.chosen = c;
  state.client = state.mode === 'test'
    ? new ExcelClient({ getToken: async () => 'test-mode', driveId: 'TEST', itemId: 'TEST', fetchImpl: (u, o) => state.mock.fetch(u, o), log: logRequest })
    : new ExcelClient({ getToken: state.auth.getToken, driveId: c.driveId, itemId: c.itemId, log: logRequest });
  state.wb = new LedgerWorkbook(state.client);
  if (state.mode === 'connected' && state.who) rememberWorkbook(localStorage, state.who, c);
  $('#workbook-info').replaceChildren(h('dl', { className: 'kv' },
    h('dt', {}, 'Chosen'), h('dd', {}, c.name), h('dt', {}, 'Drive id'), h('dd', {}, c.driveId), h('dt', {}, 'Item id'), h('dd', {}, c.itemId)));
  $('#confirm-name').value = '';
  $('#btn-selftest').disabled = true;
  $('#selftest-log').replaceChildren();
  $('#selftest-result').textContent = '';
  $('#selftest-compared').replaceChildren();
  state.checksOk = null;
  updateGate();
  await runChecks();
  updateGate();
}

/** The write test is offered only when this tab owns writing, a workbook is chosen, every read-only check passed, and its name was typed exactly. */
function updateGate() {
  const nameOk = Boolean(state.chosen) && $('#confirm-name').value.trim() === state.chosen.name;
  const ready = state.tabReady && !state.secondaryTab && !state.running && Boolean(state.chosen) && state.checksOk === true && nameOk;
  $('#btn-selftest').disabled = !ready;
  const tabNote = state.tabUnsupported ? ' This browser cannot detect other tabs; make sure no other TPAHA Books tab is open before running it.' : '';
  $('#selftest-gate').textContent = !state.tabReady ? 'Checking for other TPAHA Books tabs in this browser…'
    : state.secondaryTab ? 'TPAHA Books is open in another tab of this browser (the ledger or another Diagnostics page). Read-only checks still run here; the write test is refused until that tab is closed and this page is reloaded.'
    : state.running ? 'The connection test is running. Do not change the workbook or start it again until it finishes.'
    : !state.chosen ? 'Choose a workbook above first.'
    : state.checksOk === null ? 'Waiting for the read-only checks to finish.'
    : state.checksOk === false ? 'The connection test stays off until every read-only check above passes. Nothing is written while a check fails.'
    : !nameOk ? `All read-only checks passed. Type the workbook name exactly as shown to enable the test.${tabNote}`
    : `Ready. The test adds one row, reads it back, removes it, and compares the touched areas with their starting state.${tabNote}`;
}

async function runChecks() {
  const tbody = $('#checks tbody');
  tbody.replaceChildren();
  const results = [];
  const row = (name, ok, detail) => { results.push(ok); tbody.append(h('tr', { dataset: { ok: String(ok) } }, h('td', {}, name), h('td', { className: ok ? 'ok' : 'warn' }, ok ? 'OK' : 'Problem'), h('td', {}, detail))); };
  const finish = () => { state.checksOk = results.length > 0 && results.every(Boolean); };
  const client = state.client;
  let logBody = null, sortedBody = null;
  try { const meta = await client.getItemMeta(); row('Workbook opens', true, `${meta.name}, ${Math.round((meta.size || 0) / 1024)} KB, last changed ${meta.lastModifiedDateTime ? new Date(meta.lastModifiedDateTime).toLocaleString() : 'unknown'}`); }
  catch (e) { row('Workbook opens', false, e.message); finish(); return; }
  try {
    logBody = await client.getTableBody('LOG_Table');
    const sheet = logBody.address.split('!')[0];
    const header = await client.getRange(sheet, 'A1:I1');
    const got = header.values[0].map(String);
    const same = got.join('|') === LOG_COLUMNS.join('|');
    row('LOG_Table present with the expected columns', same, same ? `${logBody.values.filter(r => !isBlankRow(r)).length} transactions` : `columns are ${got.join(', ')}`);
  } catch (e) { row('LOG_Table present with the expected columns', false, e.message); }
  try { sortedBody = await client.getTableBody('LOG_Sorted_Table'); row('LOG_Sorted_Table present', true, `${sortedBody.rowCount} rows`); }
  catch (e) { row('LOG_Sorted_Table present', false, e.message); }
  try { const y = (await client.getRange('ConfigHidden', 'A1')).values[0][0]; const ok = Number.isInteger(Number(y)) && Number(y) > 2000; row('Workbook year (ConfigHidden!A1)', ok, String(y)); }
  catch (e) { row('Workbook year (ConfigHidden!A1)', false, e.message); }
  try { const b = (await client.getRange('ENTRY', 'B23')).values[0][0]; const ok = typeof b === 'number' || (b !== '' && Number.isFinite(Number(b))); row('Prior-year balance (ENTRY!B23) is a number', ok, ok ? 'present' : `found "${b}"`); }
  catch (e) { row('Prior-year balance (ENTRY!B23) is a number', false, e.message); }
  if (logBody && sortedBody) {
    // One shared rule with the adapter's rebuild, so the two can never disagree (review finding V1).
    const check = sortedTableMatches(logBody.values, sortedBody.values);
    row(SORTED_CHECK, check.ok, check.ok
      ? `consistent (${check.want.length} rows)`
      : `differs: ${check.have.length} rows present, ${check.want.length} expected, or a different order. This is the order the workbook's own Office Scripts require; the app rebuilds it after its next save and verifies the rebuild against a fresh read of LOG_Table.`);
  } else {
    row(SORTED_CHECK, false, 'could not compare because a table could not be read');
  }
  finish();
}

async function runSelfTest() {
  const result = $('#selftest-result');
  // The write handler checks ownership itself, not only the button gate (T2).
  if (!state.tabReady || state.secondaryTab) { result.textContent = state.tabReady ? 'Refused: TPAHA Books is open in another tab of this browser. Close it and reload this page before running the write test. Nothing was written.' : 'Refused: still checking for other TPAHA Books tabs. Nothing was written.'; result.className = 'status-line err'; updateGate(); return; }
  if (state.running || !state.wb) return;
  state.running = true;
  const btn = $('#btn-selftest');
  btn.disabled = true;
  $('#btn-find').disabled = true;
  for (const b of document.querySelectorAll('#candidates .btn-choose')) b.disabled = true;
  const log = $('#selftest-log');
  log.replaceChildren();
  result.textContent = 'Running…';
  try {
    const report = await state.wb.selfTest({ confirmTestCopy: true, onStep: s => log.append(h('li', { className: s.ok ? 'ok' : 'fail' }, `${s.name}: ${typeof s.detail === 'object' ? 'recorded' : s.detail}`)) });
    if (report.ok && report.restored) { result.textContent = 'All steps passed. The test row was added, read back and removed; the compared areas are back to their starting state (restored).'; result.className = 'status-line ok'; }
    else { result.textContent = `The connection test stopped: ${report.error || 'a step failed'}. ${report.restored ? '' : 'Check the workbook in Excel for a row described "TPAHA Books connection test" and remove it if present.'}`; result.className = 'status-line err'; }
    $('#selftest-compared').replaceChildren(
      h('p', {}, h('strong', {}, 'Compared with the starting state: '), (report.compared || []).join('; ') || 'nothing'),
      h('p', { className: 'muted' }, h('strong', {}, 'Not compared: '), (report.notCompared || []).join('; ') || 'nothing'));
  } finally {
    state.running = false;
    $('#btn-find').disabled = false;
    for (const b of document.querySelectorAll('#candidates .btn-choose')) b.disabled = false;
    updateGate();
  }
}

main().catch(e => { console.error(e); toast(e.message || String(e), 'error'); });

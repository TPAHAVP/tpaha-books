// Treasurer Ledger screen. The workbook is the record: every figure shown is read from it, and every
// Save is a targeted write to it followed by a read-back (see js/workbook/ledger-workbook.js).
//
// Saving rules on this screen (docs/workbook-mapping.md §3b):
// - every add carries one operation id; "Check workbook" and the automatic check after a reload only READ;
//   a Save (Retry) finishes or re-sends the same operation, never a second copy of it;
// - edits typed while a save is in flight, or after an unconfirmed attempt, stay on screen and are reported
//   as unsaved; they are never declared saved because an older attempt landed;
// - an unresolved change (a delete or correction that hit the wrong row, duplicate copies) is an INCIDENT:
//   it is kept in this browser, pauses all saving across Refresh and reload, and is lifted only after a
//   verified resolution (or an explicit acknowledgement when nothing can be verified). Repairs happen only
//   when a member presses a repair button and confirms;
// - pilot writer model: one designated writer at a time. The page asks before the first Correct/Delete of a
//   session, refuses to write from a second tab of the same browser, and says so.
import { CONFIG } from '../config.js';
import { initAuth } from '../auth.js';
import { ExcelClient } from '../workbook/excel-client.js';
import { MockWorkbook } from '../workbook/mock-excel.js';
import { SAMPLE_WORKBOOK } from '../workbook/sample-workbook.js';
import { LedgerWorkbook, ConflictError, VerificationError, UnresolvedOperationError, identityOf, makeMarker, monthOfIso } from '../workbook/ledger-workbook.js';
import { findCandidates, rememberWorkbook, recallWorkbook, forgetWorkbook, resolveConfigured, workbookKey } from '../workbook/locate.js';
import { SaveController } from '../save/save-controller.js';
import { draftKey, DraftStore, DraftStorageError, installUnloadGuard, clearDraftsForAccount, combineDrafts } from '../save/drafts.js';
import { incidentKey, IncidentStore } from '../save/incidents.js';
import { reportFormattingKey, ReportFormattingStore, ReportFormattingStorageError } from '../save/report-formatting.js';
import { createTabGuard } from '../save/tab-guard.js';
import { h, toast, confirmDialog, setupTabs, fmtMoney, fmtDate, todayIso, signInGate } from '../ui.js';
import { MONTH_NAMES, CATEGORIES, DEPOSIT_CATEGORIES, WITHDRAWAL_CATEGORIES, validateEntry, monthView, annualView, toCents } from './model.js';

const $ = sel => document.querySelector(sel);
const field = (form, name) => form.querySelector(`[name="${name}"]`);
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
const GRAPH = 'https://graph.microsoft.com/v1.0';
const MOCK_STATE_KEY = 'tpaha:test-mock-state';      // test mode only: the in-memory workbook survives a reload of this tab
const WRITER_ACK_KEY = 'tpaha:writer-ack';           // per browser session: the member confirmed the one-writer rule

const state = {
  auth: null, mode: 'test', client: null, wb: null, snapshot: null, workbook: null, who: null,
  month: new Date().getMonth() + 1, tab: 'entry', monthCache: new Map(), annualCache: null,
  entryCtl: null, editCtl: null, editingRef: null, drafts: null, draftWarned: false,
  settingsDirty: false, settingsSaving: false,
  incidents: null, repaired: new Set(), lastChecks: null,
  reportFmt: null, reportFmtRec: null, reportFmtWarned: false,
  tabGuard: null, secondaryTab: false,
};

// ---- boot --------------------------------------------------------------------------------
async function main() {
  $('#app-version').textContent = `v${CONFIG.appVersion}`;
  state.auth = await initAuth(CONFIG);
  if (state.auth.mode === 'local') {
    state.mode = 'test';
    const mock = new MockWorkbook(SAMPLE_WORKBOOK);
    try { const saved = sessionStorage.getItem(MOCK_STATE_KEY); if (saved) mock.importState(JSON.parse(saved)); } catch { /* fresh sample */ }
    mock.onChange = m => { try { sessionStorage.setItem(MOCK_STATE_KEY, JSON.stringify(m.exportState())); } catch { /* ignore */ } };
    window.__tpahaMock = mock;
    state.client = new ExcelClient({ getToken: async () => 'test-mode', driveId: 'TEST', itemId: 'TEST', fetchImpl: (u, o) => mock.fetch(u, o) });
    state.workbook = { driveId: 'TEST', itemId: 'TEST', name: mock.name, source: 'test' };
    state.who = { tenantId: 'test', homeAccountId: 'test-user', app: 'ledger' };
    showBanner('Test mode: in-memory copy of the workbook with made-up figures. Nothing is saved to Microsoft 365.');
  } else {
    state.mode = 'connected';
    if (!state.auth.account) { $('#main').replaceChildren(signInGate(state.auth, 'Treasurer Ledger')); $('#tabs').hidden = true; return; }
    state.who = { tenantId: state.auth.tenantId || state.auth.account.tenantId || 'tenant', homeAccountId: state.auth.account.homeAccountId || state.auth.account.username, app: 'ledger' };
    state.workbook = resolveConfigured(CONFIG, 'ledger') || recallWorkbook(localStorage, state.who);
    if (!state.workbook) { await showPicker(); return; }
    state.client = new ExcelClient({ getToken: state.auth.getToken, driveId: state.workbook.driveId, itemId: state.workbook.itemId });
  }
  state.wb = new LedgerWorkbook(state.client);
  const ids = { tenantId: state.who.tenantId, homeAccountId: state.who.homeAccountId, workbookId: workbookKey(state.workbook) };
  state.drafts = new DraftStore(localStorage, draftKey({ ...ids, schema: 3 }));
  state.incidents = new IncidentStore(localStorage, incidentKey(ids));
  const recorded = state.incidents.read();
  if (recorded) state.wb.restoreIncident(recorded);                    // a reload does not forget an unresolved change
  state.reportFmt = new ReportFormattingStore(localStorage, reportFormattingKey(ids));
  state.reportFmtRec = state.reportFmt.read();                         // shown on the next render; never run on its own
  state.tabGuard = createTabGuard();
  const tab = await state.tabGuard.start();
  state.secondaryTab = !tab.primary;
  showPilotBanner();
  buildEntryForm();
  buildSearch();
  buildEditDialog();
  buildSettings();
  setupTabs($('#tabs'), id => { state.tab = id; renderTab(); });
  installUnloadGuard(hasUnsaved);
  window.__tpahaHasUnsaved = hasUnsaved;
  window.__tpahaRequestSignOut = requestSignOut;
  window.__tpahaSecondaryTab = () => state.secondaryTab;
  window.__tpahaBusy = isBusy;
  $('#btn-refresh').addEventListener('click', () => reload('Refreshed from the workbook'));
  await reload();
  restoreDraft();
}

function showBanner(text) { const b = $('#mode-banner'); b.textContent = text; b.hidden = false; }
function showPilotBanner() {
  const b = $('#pilot-banner');
  const parts = [];
  if (CONFIG.pilot && CONFIG.pilot.writerModel === 'single') parts.push(CONFIG.pilot.note || 'Pilot: one designated writer at a time.');
  if (state.secondaryTab) parts.push('TPAHA Books is already open in another tab of this browser. This tab will not save anything: close it, or use the other tab.');
  if (!parts.length) { b.hidden = true; return; }
  b.replaceChildren(...parts.map(p => h('div', {}, p)));
  b.classList.toggle('warn-tab', state.secondaryTab);
  b.hidden = false;
}
function ctx() { return { user: state.auth.account ? state.auth.account.username : (state.mode === 'test' ? 'test-user' : '') }; }
/** Everything not yet confirmed in the workbook: typed entries, typed corrections (dialog open or not), unconfirmed attempts, a typed balance. */
function hasUnsaved() {
  return Boolean((state.entryCtl && state.entryCtl.hasUnsavedEdits) || (state.editCtl && state.editCtl.hasUnsavedEdits) || state.settingsDirty);
}
/** Writes are refused from a second tab of the same browser (pilot writer model). */
function tabAllowed() {
  if (!state.secondaryTab) return true;
  toast('This tab does not save: TPAHA Books is open in another tab of this browser. Use that tab.', 'error');
  return false;
}
/**
 * One complete operation at a time (review finding T1): while an add, correction, delete, finish, balance write
 * or repair is running, a second one is refused here, and the adapter itself queues anything that slips
 * through so that every operation reads fresh row positions before any positional request.
 */
let busyCount = 0;
function isBusy() { return busyCount > 0 || Boolean(state.wb && state.wb.busy); }
function writeAllowed() {
  if (!tabAllowed()) return false;
  if (isBusy()) { toast('Another change is still being saved or checked. Wait for it to finish, then try again.', 'error'); return false; }
  return true;
}
async function withBusy(fn) {
  busyCount += 1;
  renderBusy();
  try { return await fn(); } finally { busyCount -= 1; renderBusy(); }
}
function renderBusy() {
  const b = busyCount > 0;
  document.body.classList.toggle('busy', b);
  for (const el of document.querySelectorAll('#txn-table .btn-delete, #halt-banner .btn-repair, #btn-save-settings')) el.disabled = b;
}
/** Before the first Correct or Delete of a session: the member confirms nobody else is writing (pilot rule). */
async function confirmSoleWriter() {
  if (!CONFIG.pilot || CONFIG.pilot.writerModel !== 'single') return true;
  try { if (sessionStorage.getItem(WRITER_ACK_KEY) === '1') return true; } catch { /* ask */ }
  const ok = await confirmDialog('Pilot rule: one writer at a time. Is anyone else using this workbook right now: another member on this page, another device or tab, Excel, or the old script buttons? Corrections and deletes move rows, and two people changing rows at the same moment can affect the wrong row. Continue only if you are the only one writing.', { okLabel: 'I am the only one writing', title: 'One writer at a time' });
  if (ok) { try { sessionStorage.setItem(WRITER_ACK_KEY, '1'); } catch { /* ignore */ } }
  return ok;
}

/** Identities whose second copy is a correction this page has not finished yet (load() must not report them as damage). */
function expectedDuplicateIdentities() {
  const ids = new Set();
  if (state.editCtl && state.editCtl.pending && state.editingRef) ids.add(identityOf(state.editingRef));
  const d = state.drafts && state.drafts.read();
  if (d && d.correction && d.correction.marker && d.correction.ref) ids.add(identityOf(d.correction.ref));
  return [...ids];
}
async function reload(message) {
  try {
    $('#sync-status').firstElementChild.textContent = state.workbook.name || '';
    const snap = await state.wb.load({ expectedDuplicates: expectedDuplicateIdentities() });   // never lifts a pause; may add one (duplicate copies found)
    state.snapshot = snap;
    state.monthCache.clear();
    state.annualCache = null;
    $('#year').textContent = String(snap.year);
    document.title = `Treasurer Ledger ${snap.year} · ${CONFIG.siteName}`;
    if (state.mode === 'connected') {
      const when = snap.meta && snap.meta.lastModifiedDateTime ? new Date(snap.meta.lastModifiedDateTime).toLocaleString() : 'unknown';
      showBanner(`Connected to ${snap.meta && snap.meta.name ? snap.meta.name : state.workbook.name} · last changed ${when} · ${state.auth.account.username}`);
      if (snap.meta && snap.meta.name) state.workbook.name = snap.meta.name;
    }
    for (const n of snap.notices || []) toast(n, 'info');
    if (state.wb.halted) { persistIncident(); await checkResolution({ quiet: true }); }
    renderIncident();
    renderReportBanner();
    fillCategories(field($('#entry-form'), 'category'), field($('#entry-form'), 'type').value);
    renderList();
    renderTab();
    renderWorkbookInfo();
    if (message) toast(message);
  } catch (e) {
    console.error(e);
    toast(`Could not read the workbook: ${e.message}`, 'error');
    $('#list-summary').textContent = `Could not read the workbook: ${e.message}`;
  }
}

// ---- workbook picker (connected mode, first time) -----------------------------------------
async function graphGet(path) {
  const token = await state.auth.getToken();
  const res = await fetch(path.startsWith('http') ? path : GRAPH + path, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' } });
  if (!res.ok) throw new Error(`Graph ${res.status}`);
  return res.json();
}
async function showPicker() {
  $('#tabs').hidden = true;
  const main = $('#main');
  const list = h('div', { className: 'table-wrap' }, h('p', { className: 'muted' }, 'Searching your OneDrive…'));
  main.replaceChildren(h('section', { className: 'card gate' },
    h('h2', {}, 'Which workbook should the ledger use?'),
    h('p', {}, `Files named like "${CONFIG.workbookSearch}" in your OneDrive and files shared with you. Check the folder and last-changed time before choosing. Your choice is remembered on this device for your account.`),
    list));
  const cands = await findCandidates(graphGet, { query: CONFIG.workbookSearch });
  list.replaceChildren(candidatesTable(cands, async c => {
    rememberWorkbook(localStorage, state.who, c);
    location.reload();
  }));
  for (const w of cands.warnings) list.append(h('p', { className: 'warn small' }, w));
}
export function candidatesTable(cands, onChoose) {
  if (!cands.length) return h('p', { className: 'warn' }, 'No .xlsx workbook found. Make sure the file is in your OneDrive or has been shared with you.');
  return h('table', { id: 'candidates' },
    h('thead', {}, h('tr', {}, h('th', {}, 'Name'), h('th', {}, 'Where'), h('th', {}, 'Last changed'), h('th', { className: 'num' }, 'Size'), h('th'))),
    h('tbody', {}, cands.map(c => h('tr', {},
      h('td', {}, c.name),
      h('td', {}, c.source === 'shared' ? `Shared by ${c.sharedBy || 'someone'}` : `My OneDrive ${c.path}`),
      h('td', {}, c.lastModified ? new Date(c.lastModified).toLocaleString() : ''),
      h('td', { className: 'num' }, c.size ? `${Math.round(c.size / 1024)} KB` : ''),
      h('td', {}, h('button', { type: 'button', className: 'btn btn-sm btn-primary btn-choose', onClick: () => onChoose(c) }, 'Choose'))))));
}

// ---- shared form helpers -----------------------------------------------------------------
function fillCategories(select, type) {
  const names = type === 'Deposit' ? DEPOSIT_CATEGORIES : type === 'Withdrawal' ? WITHDRAWAL_CATEGORIES : [];
  const current = select.value;
  select.replaceChildren(h('option', { value: '' }, 'Choose a category'), ...names.map(n => h('option', { value: n }, n)));
  if (names.includes(current)) select.value = current;
}
function showErrors(form, errors) {
  for (const e of form.querySelectorAll('.error')) e.remove();
  for (const [name, msg] of Object.entries(errors || {})) {
    const input = field(form, name);
    if (input && input.closest('.field')) input.closest('.field').append(h('span', { className: 'error', role: 'alert' }, msg));
  }
}
function formValues(form) { return Object.fromEntries(new FormData(form).entries()); }
const describeTxn = t => t ? `#${t.id} · ${fmtDate(t.date, 'long')} · ${t.description} · ${fmtMoney(toCents(t.amount))} · ${t.type} · ${t.category}` : '';

function saveStateText(c) {
  switch (c.state) {
    case 'unsaved': return c.pending ? 'Unsaved changes (an earlier attempt is still unconfirmed; Save checks the workbook first)' : 'Unsaved changes';
    case 'saving': return c.phase === 'checking' ? 'Checking the workbook (reading only)…' : c.phase === 'resuming' ? 'Finishing the earlier attempt…' : 'Saving to the workbook…';
    case 'saved': {
      const r = c.result || {};
      const t = r.txn || r;
      const id = t && t.id ? `#${t.id}` : '';
      const extra = r.renumberedFrom ? ` (number changed from #${r.renumberedFrom} because another entry took it at the same time)` : '';
      return `Saved to workbook${id ? ' · ' + id : ''}${extra}`;
    }
    case 'failed': {
      const e = c.error || {};
      const retry = e.retryAfter ? ` Microsoft asked to wait ${e.retryAfter} s before trying again.` : '';
      if (c.incomplete) return `Not finished: ${e.message}`;
      return c.needsVerification
        ? `Could not save: the answer from Microsoft was lost, so this may or may not be in the workbook.${retry} Press "Check workbook" (reads only) or Save: the page checks the workbook before it writes anything.`
        : `Could not save: ${e.message || 'unknown error'}${retry}`;
    }
    case 'conflict': return c.conflict && c.conflict.name === 'MarkerCollisionError' ? `Could not save: ${c.conflict.message}` : `Someone else changed this record: ${c.conflict ? c.conflict.message : ''}`;
    default: return '';
  }
}

// ---- drafts (entry + correction, with their unconfirmed operations) ----------------------
/** The single persisted record for both forms; each part is kept or dropped on its own (T4). Never cleared from one form's success path. */
function persistDraft() {
  const e = state.entryCtl, c = state.editCtl;
  const rec = combineDrafts({
    entry: e ? { draft: e.draft, state: e.state, pending: e.pending } : null,
    correction: c && state.editingRef ? { ref: state.editingRef, draft: c.draft, pending: c.pending, hasUnsavedEdits: c.hasUnsavedEdits } : null,
  });
  if (!rec) { state.drafts.clear(); return; }
  try { state.drafts.write(rec); }
  catch (err) { if (err instanceof DraftStorageError && !state.draftWarned) { state.draftWarned = true; toast(err.message, 'error'); } }
}
function restoreDraft() {
  const d = state.drafts.read();
  if (!d) return;
  if (d.entry || d.pending) {
    const entry = d.entry || d.pending.payload;
    const form = $('#entry-form');
    for (const [k, v] of Object.entries(entry)) { const el = field(form, k); if (el && k !== 'category') el.value = v; }
    fillCategories(field(form, 'category'), entry.type);
    if (entry.category) field(form, 'category').value = entry.category;
    state.entryCtl.edit(entry);
    if (d.pending && d.pending.marker && d.pending.payload) {
      state.entryCtl.restorePending(d.pending);
      toast('An earlier save attempt was never confirmed. Checking the workbook (reading only)…', 'info');
      state.entryCtl.check().then(afterEntryCheck);
    } else {
      toast('Restored what you typed earlier. It is not in the workbook until you press Save.', 'info');
    }
  }
  if (d.correction && d.correction.ref) {
    const { ref, marker, payload, draft } = d.correction;
    openEdit(ref, { reset: false });
    const form = $('#edit-form');
    const shown = draft || payload;
    for (const [k, v] of Object.entries(shown)) { const el = field(form, k); if (el && k !== 'category') el.value = v; }
    fillCategories(field(form, 'category'), shown.type);
    if (shown.category) field(form, 'category').value = shown.category;
    if (draft) state.editCtl.edit(draft);
    if (marker && payload) {
      state.editCtl.restorePending({ marker, payload });
      toast(`An earlier correction to #${ref.id} was never confirmed. Checking the workbook (reading only)…`, 'info');
      state.editCtl.check().then(afterEditCheck);
    }
  }
}

// ---- entry form ----------------------------------------------------------------------------
function buildEntryForm() {
  const form = $('#entry-form');
  const stateEl = $('#save-state');
  const actions = $('#save-actions');
  const submit = $('#btn-submit');
  state.entryCtl = new SaveController({
    perform: (draft, { marker }) => withBusy(() => state.wb.addTransaction(draft, { marker, ...ctx() })),
    inspect: (payload, { marker }) => state.wb.inspectAppend(marker, payload),
    resume: (payload, { marker }) => withBusy(() => state.wb.completeAppend(marker, payload)),
    makeMarker,                                          // ISO 8601 with nine fraction digits, unique per operation (T3)
    onChange: c => {
      stateEl.dataset.state = c.state;
      stateEl.textContent = saveStateText(c);
      submit.disabled = c.state === 'saving' || c.state === 'idle' || c.state === 'saved';
      submit.textContent = c.state === 'saving' ? (c.phase === 'checking' ? 'Checking…' : 'Saving…') : (c.incomplete ? 'Finish saving' : 'Save to workbook');
      actions.replaceChildren();
      if (c.state === 'failed') {
        actions.append(h('button', { type: 'button', id: 'btn-retry', className: 'btn', onClick: () => { if (writeAllowed()) state.entryCtl.retry().then(afterEntrySave); } }, c.incomplete ? 'Finish' : 'Retry'));
        if (c.pending) actions.append(h('button', { type: 'button', id: 'btn-verify', className: 'btn btn-primary', onClick: () => state.entryCtl.check().then(afterEntryCheck) }, 'Check workbook'));
      }
      persistDraft();
    },
  });
  field(form, 'date').value = todayIso();
  field(form, 'type').addEventListener('change', () => fillCategories(field(form, 'category'), field(form, 'type').value));
  form.addEventListener('input', () => {
    const v = formValues(form);
    if (!v.description && !v.amount && !v.category && !v.chequeNum && !v.notes && !state.entryCtl.draft) return;   // nothing typed yet
    if (!same(v, state.entryCtl.draft)) state.entryCtl.edit(v);
  });
  form.addEventListener('submit', async e => {
    e.preventDefault();
    if (!writeAllowed()) return;
    const v = formValues(form);
    const check = validateEntry({ year: state.snapshot ? state.snapshot.year : new Date().getFullYear() }, v);
    showErrors(form, check.ok ? {} : check.errors);
    if (!check.ok) return;
    if (!same(v, state.entryCtl.draft)) state.entryCtl.edit(v);   // an unchanged form is not a new edit
    const reserved = reserveReportFormatting([monthOfIso(v.date)]);
    const r = await state.entryCtl.submit();
    if (r.ignored) { releaseReportFormatting(reserved); return; }
    afterEntrySave(r, reserved);
  });
}
function afterEntryCheck(c) {
  if (!c || c.ignored) return;
  if (c.found && c.outcome === 'landed') afterEntrySave({ ok: true, result: c.result, resolvedPending: true, currentDraftSaved: c.currentDraftSaved });
  else if (c.error instanceof UnresolvedOperationError) showIncident();
  else if (c.error) toast(`Could not check the workbook: ${c.error.message}`, 'error');
}
async function afterEntrySave(r, reserved = []) {
  if (!r || r.ignored) return;
  const c = state.entryCtl;
  const form = $('#entry-form');
  if (r.ok) {
    const res = r.result || {};
    const txn = res.txn || res;
    noteReportFormatting(res);
    if (c.state === 'saved') {
      for (const name of ['category', 'amount', 'description', 'chequeNum', 'notes']) field(form, name).value = '';
      fillCategories(field(form, 'category'), field(form, 'type').value);
      showErrors(form, {});
      persistDraft();                                    // drops the saved entry only; a correction's draft or pending operation stays (T4)
      toast(`Saved to workbook · #${txn.id}`);
    } else {
      toast(`Saved to workbook · #${txn.id} (the earlier attempt). What you typed after it is still unsaved and stays on screen.`, 'info');
    }
    if (r.resolvedPending && res.alreadyExisted) toast('The earlier attempt had reached the workbook. Nothing was written twice.', 'info');
    for (const w of res.warnings || []) toast(w, 'info');
    await reload();
    if (res.renumberedFrom) toast(`Another entry took #${res.renumberedFrom} at the same moment; yours is #${txn.id}.`, 'info');
  } else if (r.error instanceof UnresolvedOperationError) {
    showIncident();
  } else if (r.error instanceof VerificationError) {
    toast(r.error.message, 'error');
    await reload();
  }
  if (state.entryCtl.state === 'conflict') releaseReportFormatting(reserved);   // refused before anything was written
}

// ---- list ----------------------------------------------------------------------------------
function buildSearch() {
  const form = $('#search-form');
  form.addEventListener('input', renderList);
  form.addEventListener('submit', e => { e.preventDefault(); renderList(); });
}
function filteredTransactions() {
  if (!state.snapshot) return [];
  const form = $('#search-form');
  const from = field(form, 'from').value, to = field(form, 'to').value, text = field(form, 'text').value.trim().toLowerCase();
  return state.snapshot.transactions
    .filter(t => (!from || t.date >= from) && (!to || t.date <= to))
    .filter(t => !text || `${t.description} ${t.chequeNum} ${t.amount} ${t.category} ${t.id}`.toLowerCase().includes(text))
    .sort((a, b) => b.date.localeCompare(a.date) || b.id - a.id);
}
function renderList() {
  const rows = filteredTransactions();
  const table = $('#txn-table');
  table.querySelector('thead').replaceChildren(h('tr', {}, h('th', {}, 'Date'), h('th', {}, 'Description'), h('th', { className: 'num' }, 'Amount'), h('th', {}, 'Type · Category'), h('th', {}, 'Cheque'), h('th', {}, 'ID'), h('th', {}, '')));
  table.querySelector('tbody').replaceChildren(...rows.map(t => h('tr', { dataset: { id: String(t.id) } },
    h('td', {}, fmtDate(t.date, 'long')),
    h('td', {}, t.description, t.notes ? h('div', { className: 'small muted' }, t.notes) : null),
    h('td', { className: `num ${t.type === 'Deposit' ? 'dep' : 'wd'}` }, fmtMoney(toCents(t.amount))),
    h('td', {}, `${t.type} · ${t.category}`),
    h('td', {}, t.chequeNum || ''),
    h('td', {}, `#${t.id}`),
    h('td', { className: 'row', style: 'gap:.3rem;flex-wrap:nowrap' },
      h('button', { type: 'button', className: 'btn btn-sm btn-edit', onClick: async () => { if (tabAllowed() && await confirmSoleWriter()) openEdit(t); } }, 'Correct'),
      h('button', { type: 'button', className: 'btn btn-sm btn-danger btn-delete', disabled: busyCount > 0, onClick: () => deleteTxn(t) }, 'Delete')))));
  $('#list-summary').textContent = state.snapshot ? `${rows.length} of ${state.snapshot.transactions.length} transactions shown · read ${new Date(state.snapshot.loadedAt).toLocaleTimeString()}` : '';
}

function conflictDialog(err, { onReload }) {
  const cur = err.current;
  const dlg = h('dialog', { className: 'confirm conflict' },
    h('h3', {}, 'Someone else changed this record'),
    h('p', {}, err.reason === 'missing' ? (err.message || 'This record no longer exists in the workbook. It may have been deleted by another member or in Excel.') : 'The record in the workbook is different from the one you loaded, so nothing was changed.'),
    cur ? h('dl', { className: 'kv' }, h('dt', {}, 'Now in workbook'), h('dd', {}, describeTxn(cur))) : null,
    h('div', { className: 'row-end' },
      h('button', { type: 'button', className: 'btn', onClick: () => dlg.close() }, 'Cancel'),
      h('button', { type: 'button', className: 'btn btn-primary btn-reload', onClick: () => { dlg.close(); onReload(); } }, 'Reload')));
  dlg.addEventListener('close', () => dlg.remove());
  document.body.append(dlg);
  dlg.showModal();
}

async function deleteTxn(t) {
  if (!writeAllowed() || !(await confirmSoleWriter())) return;
  const ok = await confirmDialog(`Delete transaction ${describeTxn(t)} from the workbook?`, { okLabel: 'Delete from workbook', danger: true });
  if (!ok) return;
  if (!writeAllowed()) return;                           // something may have started while the dialog was open
  const reserved = reserveReportFormatting([monthOfIso(t.date)]);
  try {
    const r = await withBusy(() => state.wb.deleteTransaction(t));
    noteReportFormatting(r);
    toast(`Deleted #${t.id}`);
    for (const n of r.notices || []) toast(n, 'info');
    await reload();
  } catch (e) {
    if (e instanceof ConflictError) { releaseReportFormatting(reserved); conflictDialog(e, { onReload: () => reload() }); }
    else if (e instanceof UnresolvedOperationError) showIncident();
    else { toast(e.message, 'error'); await reload(); }
  }
}

// ---- incidents (unresolved changes): paused until verified -------------------------------------
function persistIncident() {
  if (!state.incidents) return;
  const rec = state.wb.incident;
  if (!rec) { state.incidents.clear(); return; }
  if (!state.incidents.write(rec)) toast('This browser could not keep the record of the unresolved change (storage full or blocked). It stays paused on this page; note the details before closing.', 'error');
}
function showIncident() {
  persistIncident();
  state.lastChecks = null;
  renderIncident();
  $('#halt-banner').scrollIntoView({ block: 'nearest' });
}
/** Read-only: asks the adapter whether the documented resolution conditions hold now. */
async function checkResolution({ quiet = false } = {}) {
  try {
    const v = await state.wb.verifyResolution();
    state.lastChecks = v;
    persistIncident();
    if (v.resolved && !quiet) toast('Verified: the unresolved change is resolved. Saving is enabled again.');
    else if (!quiet) toast(v.inconclusive ? 'Could not verify everything from the table alone; see the details.' : 'Not resolved yet; see which checks still fail.', 'info');
    return v;
  } catch (e) { toast(`Could not check the workbook: ${e.message}`, 'error'); return null; }
}
function dismissIncident() {
  if (state.wb.halted) return;
  state.wb.incident = null;
  state.incidents.clear();
  state.repaired.clear();
  state.lastChecks = null;
  renderIncident();
}
function renderIncident() {
  const inc = state.wb.incident;
  const b = $('#halt-banner');
  if (!inc) { b.hidden = true; b.classList.remove('resolved'); b.replaceChildren(); return; }
  const d = inc.details || {};
  const paused = Boolean(state.wb.halted);
  b.classList.toggle('resolved', !paused);
  const items = [];
  const repairButton = (key, label, action, danger = false) => state.repaired.has(key)
    ? h('span', { className: 'ok' }, ' Done.')
    : h('button', { type: 'button', className: `btn btn-sm ${danger ? 'btn-danger' : ''} btn-repair`, dataset: { repair: key }, onClick: action }, label);
  if (paused) {
    if (d.victim) {
      const key = `restore:${identityOf(d.victim)}`;
      items.push(h('li', {}, `Row removed by this change (values as this device last saw them; a newer change by someone else would not be included): ${describeTxn(d.victim)}. `,
        d.victim.values ? repairButton(key, `Restore #${d.victim.id} as last seen`, () => repairRestore(d.victim, key)) : null));
    } else if (inc.kind === 'wrong-row-deleted') {
      items.push(h('li', {}, 'The row that was removed could not be identified from the table. Compare with Excel\'s version history.'));
    }
    if (d.removedByOthers && d.removedByOthers.length) items.push(h('li', {}, `Removed by someone else meanwhile: ${d.removedByOthers.map(describeTxn).join('; ')}.`));
    if (d.missing && d.missing.length && !d.victim) items.push(h('li', {}, `Missing since the change: ${d.missing.map(describeTxn).join('; ')}.`));
    if (d.copies && d.copies.length) {
      const byFp = new Map();
      for (const c of d.copies) byFp.set(c.fingerprint, [...(byFp.get(c.fingerprint) || []), c]);
      items.push(h('li', {}, `Copies of #${d.id ?? d.copies[0].id} now in the workbook:`, h('ul', {}, [...byFp.values()].map(group => {
        const c = group[0];
        const tag = c.fingerprint === d.newFingerprint ? ' (your correction)' : c.fingerprint === d.oldFingerprint ? ' (old content)' : '';
        const key = `remove:${identityOf(c)}|${c.fingerprint}`;
        return group.length > 1
          ? h('li', {}, `${group.length} identical copies: ${describeTxn(c)}${tag} `, repairButton(key, 'Remove one of the identical copies', () => repairRemoveCopy(c, key, { allowIdentical: true }), true))
          : h('li', {}, `${describeTxn(c)}${tag} `, repairButton(key, 'Remove this copy', () => repairRemoveCopy(c, key), true));
      }))));
    } else if (d.duplicateCopy && d.identity && d.oldFingerprint) {
      const old = state.snapshot && state.snapshot.transactions.find(t => identityOf(t) === d.identity && t.fingerprint === d.oldFingerprint);
      const key = `remove:${d.identity}|${d.oldFingerprint}`;
      items.push(h('li', {}, `Your correction to #${d.id} is saved, but the OLD copy of #${d.id} may still be present${old ? ` (${describeTxn(old)})` : ''}. `,
        repairButton(key, `Remove the old copy of #${d.id}`, () => repairRemoveCopy({ id: d.id, timestamp: d.timestamp, fingerprint: d.oldFingerprint, ...(old || {}) }, key), true)));
    }
    if (d.possiblyAffected) items.push(h('li', {}, `Renumbered by mistake: ${describeTxn(d.possiblyAffected)} was #${d.previousNumber} and now carries #${d.number}. Restore its number in Excel, then check again.`));
  }
  const checks = state.lastChecks || (inc.checks ? { checks: inc.checks, resolved: inc.resolved, inconclusive: false } : null);
  const checksEl = checks && checks.checks.length ? h('ul', { className: 'checks', id: 'incident-checks' }, checks.checks.map(c => h('li', { dataset: { ok: String(c.ok) } },
    h('span', { className: c.ok === true ? 'ok' : c.ok === false ? 'no' : 'unknown' }, c.ok === true ? 'OK' : c.ok === false ? 'Not yet' : 'Cannot tell'), ` · ${c.name}${c.detail ? ` (${c.detail})` : ''}`))) : null;
  const actions = [];
  if (paused) {
    actions.push(h('button', { type: 'button', className: 'btn btn-primary btn-check-resolution', onClick: async () => { await checkResolution(); renderIncident(); } }, 'Check whether this is resolved'));
    if (checks && checks.inconclusive) actions.push(h('button', { type: 'button', className: 'btn btn-acknowledge', onClick: acknowledgeIncident }, 'I checked the workbook in Excel; mark as resolved'));
    actions.push(h('button', { type: 'button', className: 'btn btn-reload', onClick: () => reload('Refreshed from the workbook') }, 'Refresh (stays paused)'));
  } else {
    actions.push(h('button', { type: 'button', className: 'btn btn-primary btn-dismiss', onClick: dismissIncident }, 'Dismiss'));
  }
  b.replaceChildren(
    h('strong', {}, paused ? 'Changes are paused. ' : `Resolved (${inc.resolution === 'acknowledged' ? 'acknowledged by a member' : 'verified against the workbook'}). Saving is enabled again. `),
    h('span', {}, inc.message),
    items.length ? h('ul', { className: 'halt-details' }, items) : null,
    checksEl,
    paused ? h('p', { className: 'small' }, 'Nothing is repaired automatically. A repair happens only when you press one of the buttons above and confirm. Refresh and reloading the page keep this pause; it ends only when the checks above pass, or when nothing can be checked and you confirm you compared the workbook with Excel\'s version history.') : null,
    h('div', { className: 'row' }, actions));
  b.hidden = false;
}
async function acknowledgeIncident() {
  const ok = await confirmDialog('Mark this as resolved? Do this only after comparing the transaction table with Excel\'s version history and fixing anything missing. The app could not verify it from the table alone.', { okLabel: 'Mark as resolved', danger: true, title: 'Acknowledge' });
  if (!ok) return;
  try { await state.wb.acknowledgeIncident({ confirm: true, by: ctx().user }); persistIncident(); renderIncident(); toast('Marked as resolved. Saving is enabled again.'); }
  catch (e) { toast(e.message, 'error'); }
}
async function repairRestore(victim, key) {
  if (!writeAllowed()) return;
  const ok = await confirmDialog(`Add this row back exactly as this device last saw it? ${describeTxn(victim)}. If someone changed it after that, their change is not included; check Excel's version history first if unsure.`, { okLabel: 'Restore row' });
  if (!ok) return;
  if (!writeAllowed()) return;
  try {
    const t = await withBusy(() => state.wb.reappendRow(victim.values, { confirm: true }));
    state.repaired.add(key);
    toast(`Restored #${t.id}`);
    await checkResolution({ quiet: true });
    await reload();
  } catch (e) { if (e instanceof UnresolvedOperationError) showIncident(); else toast(e.message, 'error'); }
}
async function repairRemoveCopy(copy, key, { allowIdentical = false } = {}) {
  if (!writeAllowed()) return;
  const ok = await confirmDialog(`Remove ${allowIdentical ? 'one of the identical copies' : 'this copy'} from the workbook? ${describeTxn(copy) || `#${copy.id}`}`, { okLabel: 'Remove copy', danger: true });
  if (!ok) return;
  if (!writeAllowed()) return;
  try {
    const r = await withBusy(() => state.wb.removeCopy(identityOf(copy), copy.fingerprint, { confirm: true, allowIdentical }));
    if (r.removed) { state.repaired.add(key); toast('Copy removed'); for (const n of r.notices || []) toast(n, 'info'); await checkResolution({ quiet: true }); await reload(); }
    else toast(`Nothing was removed (${r.reason === 'not-found' ? 'that copy is no longer there' : r.reason === 'multiple' ? 'there is more than one identical copy' : r.reason === 'moving' ? 'the table kept changing' : 'the delete did not reach the workbook'}). Press Refresh and look again.`, 'error');
  } catch (e) { if (e instanceof UnresolvedOperationError) showIncident(); else toast(e.message, 'error'); }
}

// ---- unfinished monthly report formatting ----------------------------------------------------
const REPORT_INTERRUPTED = 'A change was interrupted before the monthly report rows were updated. Your transactions are unaffected; only which rows the monthly sheets show may be out of date.';
/**
 * Written BEFORE the first workbook change, not after it (review W2). If this page closes between the
 * transaction landing and the formatting finishing, the months are already recorded and the next visit offers
 * to finish them. Nothing is rendered here, so an ordinary save shows no banner on its way past. Returns the
 * months reserved, so a call that turns out to write nothing can release them again.
 */
function reserveReportFormatting(months) {
  const wanted = [...new Set((months || []).filter(m => m >= 1 && m <= 12))];
  if (!wanted.length || !state.reportFmt) return [];
  try { state.reportFmtRec = state.reportFmt.add(wanted, REPORT_INTERRUPTED, 'interrupted'); }
  catch (e) {
    if (e instanceof ReportFormattingStorageError && !state.reportFmtWarned) { state.reportFmtWarned = true; toast(e.message, 'error'); }
    state.reportFmtRec = { months: wanted, message: REPORT_INTERRUPTED, blockedBy: 'interrupted' };   // this page still knows
  }
  return wanted;
}
/** Drops a reservation for an operation that provably wrote nothing. */
function releaseReportFormatting(months) {
  if (!months || !months.length || !state.reportFmt) return;
  state.reportFmtRec = state.reportFmt.remove(months);
  renderReportBanner();
}
/**
 * Records what a finished operation managed to format and what it did not, then shows it. Months are cleared
 * only on verified completion. This never runs anything: a member presses the button.
 */
function noteReportFormatting(result) {
  if (!result || !state.reportFmt) return;
  const unfinished = result.reportFormatting;
  try {
    if (unfinished && unfinished.pending) state.reportFmtRec = state.reportFmt.add(unfinished.months, unfinished.message, unfinished.blockedBy);
    else if (result.formattedMonths) state.reportFmtRec = state.reportFmt.remove(result.formattedMonths);
  } catch (e) {
    if (e instanceof ReportFormattingStorageError && !state.reportFmtWarned) { state.reportFmtWarned = true; toast(e.message, 'error'); }
  }
  renderReportBanner();
}
function renderReportBanner() {
  const b = $('#report-banner');
  if (!b) return;
  const rec = state.reportFmtRec;
  if (!rec) { b.hidden = true; b.replaceChildren(); return; }
  const names = rec.months.map(m => MONTH_NAMES[m - 1]).join(', ');
  const sheets = rec.months.length === 1 ? 'the monthly report sheet' : 'the monthly report sheets';
  const what = rec.blockedBy === 'sorted-table'
    ? `The sorted helper table that ${sheets} for ${names} read could not be confirmed, so the rows were left alone rather than set from figures that may be out of date. Finishing will rebuild that helper table first, then set the rows.`
    : `What is left is only which rows are shown on ${sheets} for ${names}. Until it is done, a transaction can sit in a hidden row when that sheet is opened or printed in Excel, even though this website shows it.`;
  b.replaceChildren(
    h('strong', {}, rec.message || 'Excel report formatting still needs updating.'),
    h('p', { className: 'small' }, `${what} Your transactions are not touched by this: finishing writes to the helper table and the monthly sheets only, never to the transaction table, so it cannot add or remove a transaction however many times it runs.`),
    h('div', { className: 'row' }, h('button', { type: 'button', className: 'btn btn-primary btn-finish-report', onClick: finishReportFormatting }, 'Finish report formatting')));
  b.hidden = false;
}
async function finishReportFormatting() {
  if (!state.reportFmtRec || !writeAllowed()) return;
  const months = state.reportFmtRec.months;
  try {
    const r = await withBusy(() => state.wb.finishReportRows(months));
    if (r.ok) {
      state.reportFmtRec = state.reportFmt.remove(months);
      toast(r.repairedSortedTable ? 'Helper table rebuilt and Excel report formatting finished.' : 'Excel report formatting finished.');
    } else {
      try { state.reportFmtRec = state.reportFmt.add(months, state.reportFmtRec.message, r.blockedBy); } catch { /* already shown on this page */ }
      toast(`Still not finished: ${r.error}`, 'error');
    }
  } catch (e) {
    toast(`Still not finished: ${e.message}`, 'error');
  }
  renderReportBanner();
}

// ---- correct (edit) ------------------------------------------------------------------------
function buildEditDialog() {
  const dlg = $('#edit-dialog'), form = $('#edit-form'), stateEl = $('#edit-state'), actions = $('#edit-actions');
  state.editCtl = new SaveController({
    perform: draft => withBusy(() => state.wb.updateTransaction(state.editingRef, draft)),
    inspect: payload => state.wb.inspectCorrection(state.editingRef, payload),
    resume: payload => withBusy(() => state.wb.completeCorrection(state.editingRef, payload)),
    makeMarker: () => `correction:${state.editingRef ? identityOf(state.editingRef) : ''}:${Date.now()}`,
    onChange: c => {
      stateEl.dataset.state = c.state;
      stateEl.textContent = saveStateText(c);
      $('#btn-edit-save').disabled = c.state === 'saving';
      $('#btn-edit-save').textContent = c.incomplete ? 'Finish saving' : 'Save to workbook';
      actions.replaceChildren();
      if (c.state === 'conflict') actions.append(h('button', { type: 'button', className: 'btn btn-reload', onClick: reloadEditedRecord }, 'Reload record'));
      if (c.state === 'failed') {
        actions.append(h('button', { type: 'button', className: 'btn btn-edit-retry', onClick: () => { if (writeAllowed()) state.editCtl.retry().then(afterEditSave); } }, c.incomplete ? 'Finish' : 'Retry'));
        if (c.pending) actions.append(h('button', { type: 'button', className: 'btn btn-primary btn-edit-verify', onClick: () => state.editCtl.check().then(afterEditCheck) }, 'Check workbook'));
      }
      persistDraft();
    },
  });
  field(form, 'type').addEventListener('change', () => fillCategories(field(form, 'category'), field(form, 'type').value));
  form.addEventListener('input', () => { const v = formValues(form); if (!same(v, state.editCtl.draft)) state.editCtl.edit(v); });
  form.addEventListener('submit', async e => {
    e.preventDefault();
    if (!writeAllowed()) return;
    const v = formValues(form);
    const check = validateEntry({ year: state.snapshot.year }, v);
    showErrors(form, check.ok ? {} : check.errors);
    if (!check.ok) return;
    if (!same(v, state.editCtl.draft)) state.editCtl.edit(v);
    const reserved = reserveReportFormatting([monthOfIso(state.editingRef && state.editingRef.date), monthOfIso(v.date)]);
    const r = await state.editCtl.submit();
    if (r.ignored) { releaseReportFormatting(reserved); return; }
    afterEditSave(r, reserved);
  });
  $('#btn-edit-cancel').addEventListener('click', cancelEdit);
  dlg.addEventListener('cancel', e => { e.preventDefault(); cancelEdit(); });   // Escape goes through the same check
}
async function cancelEdit() {
  const dlg = $('#edit-dialog');
  if (state.editCtl.state === 'saving') { toast('Please wait for the save to finish.', 'info'); return; }
  if (state.editCtl.hasUnsavedEdits && !(await confirmDialog(state.editCtl.pending ? 'An attempt to save this correction was never confirmed. Discard it without checking the workbook?' : 'Discard the changes to this record? They are not in the workbook.', { okLabel: 'Discard', danger: true }))) return;
  state.editCtl.reset({ discardPending: true });
  state.editingRef = null;
  persistDraft();
  dlg.close();
}
/** The values a transaction puts into the Correct form, in the shape formValues() returns. */
const txnToForm = t => ({ date: t.date, type: t.type, category: t.category, amount: Number(t.amount).toFixed(2), description: String(t.description ?? ''), chequeNum: String(t.chequeNum ?? ''), notes: String(t.notes ?? '') });
/** Reloading a record must never throw away typed text without asking (review finding U1). */
function reloadChoiceDialog(fresh, typed) {
  return new Promise(resolve => {
    let done = false;
    const finish = v => { if (done) return; done = true; dlg.close(); resolve(v); };
    const dlg = h('dialog', { className: 'confirm reload-choice' },
      h('h3', {}, 'Keep what you typed, or use the workbook version?'),
      h('p', {}, 'This record was read again from the workbook and it differs from what is on your screen. Nothing is written until you press Save.'),
      h('dl', { className: 'kv' },
        h('dt', {}, 'Now in workbook'), h('dd', {}, describeTxn(fresh)),
        h('dt', {}, 'What you typed'), h('dd', {}, `${fmtDate(typed.date, 'long')} · ${typed.description} · ${fmtMoney(toCents(typed.amount))} · ${typed.type} · ${typed.category}`)),
      h('div', { className: 'row-end' },
        h('button', { type: 'button', className: 'btn btn-cancel', onClick: () => finish('cancel') }, 'Cancel'),
        h('button', { type: 'button', className: 'btn btn-use-workbook', onClick: () => finish('workbook') }, 'Use the workbook version'),
        h('button', { type: 'button', className: 'btn btn-primary btn-keep-typed', onClick: () => finish('keep') }, 'Keep what I typed')));
    dlg.addEventListener('close', () => { dlg.remove(); if (!done) { done = true; resolve('cancel'); } });
    document.body.append(dlg);
    dlg.showModal();
  });
}
async function reloadEditedRecord() {
  const ref = state.editingRef;
  if (!ref) return;
  await reload();
  const fresh = state.snapshot.transactions.find(x => x.id === ref.id && x.timestamp === ref.timestamp);
  if (!fresh) {
    state.editCtl.reset({ discardPending: true });
    state.editingRef = null;
    persistDraft();
    $('#edit-dialog').close();
    toast('That record no longer exists in the workbook.', 'error');
    return;
  }
  const typed = state.editCtl.draft;
  if (!typed || same(typed, txnToForm(fresh))) {
    openEdit(fresh);                                                  // nothing typed that reloading would lose
    toast('The record now shows the values in the workbook. Your change was not applied.', 'info');
    return;
  }
  const choice = await reloadChoiceDialog(fresh, typed);
  if (choice === 'cancel') return;                                    // the dialog stays exactly as it is
  if (choice === 'workbook') { openEdit(fresh); toast('The Correct form now shows the values in the workbook. What you had typed was replaced.', 'info'); return; }
  state.editingRef = fresh;                                           // keep the typed text, aimed at the version just read and shown
  state.editCtl.edit(typed);
  persistDraft();
  toast(`Your text is kept. Press Save to apply it to #${fresh.id} as it now stands in the workbook.`, 'info');
}
function openEdit(t, { reset = true } = {}) {
  state.editingRef = t;
  const form = $('#edit-form');
  $('#edit-title').textContent = `Correct transaction #${t.id}`;
  field(form, 'date').value = t.date;
  field(form, 'type').value = t.type;
  fillCategories(field(form, 'category'), t.type);
  field(form, 'category').value = t.category;
  field(form, 'amount').value = Number(t.amount).toFixed(2);
  field(form, 'description').value = t.description;
  field(form, 'chequeNum').value = t.chequeNum;
  field(form, 'notes').value = t.notes;
  showErrors(form, {});
  if (reset) state.editCtl.reset({ discardPending: true });
  if (!$('#edit-dialog').open) $('#edit-dialog').showModal();
}
function afterEditCheck(c) {
  if (!c || c.ignored) return;
  if (c.found && c.outcome === 'landed') afterEditSave({ ok: true, result: c.result, resolvedPending: true, currentDraftSaved: c.currentDraftSaved });
  else if (c.error instanceof UnresolvedOperationError) showIncident();
  else if (c.error) toast(`Could not check the workbook: ${c.error.message}`, 'error');
}
async function afterEditSave(r, reserved = []) {
  if (!r || r.ignored) return;
  const c = state.editCtl;
  if (r.ok) {
    const saved = r.result;
    noteReportFormatting(saved);
    if (saved && saved.id !== undefined) {
      state.editingRef = saved;      // later saves must target the corrected copy…
      persistDraft();                // …and the stored record must carry that reference with any newer draft (U1)
    }
    for (const w of saved.warnings || []) toast(w, 'info');
    for (const n of saved.notices || []) toast(n, 'info');
    if (c.state === 'saved') {
      $('#edit-dialog').close();
      state.editingRef = null;
      persistDraft();
      toast(`Saved to workbook · #${saved.id}`);
    } else {
      toast(`Saved to workbook · #${saved.id} (the earlier attempt). The changes you typed after it are still unsaved.`, 'info');
    }
    await reload();
  } else if (r.error instanceof UnresolvedOperationError) {
    showIncident();
  } else if (r.error instanceof VerificationError) {
    toast(r.error.message, 'error');
    await reload();
  }
  if (state.editCtl.state === 'conflict') releaseReportFormatting(reserved);   // nothing was written
}

// ---- months and annual (workbook figures) ------------------------------------------------
function docForCheck() {
  const s = state.snapshot;
  return { year: s.year, priorYearBalance: s.priorYearBalance, transactions: s.transactions.map(t => ({ ...t, number: t.id, createdAt: t.timestamp, deleted: false })) };
}
async function renderMonths() {
  const chips = $('#month-chips');
  chips.replaceChildren(...MONTH_NAMES.map((name, i) => h('button', { type: 'button', className: `chip${state.month === i + 1 ? ' active' : ''}`, dataset: { month: String(i + 1) }, onClick: () => { state.month = i + 1; renderMonths(); } }, name.slice(0, 3))));
  const view = $('#month-view');
  view.replaceChildren(h('p', { className: 'muted' }, 'Reading the workbook…'));
  let m;
  try {
    if (!state.monthCache.has(state.month)) state.monthCache.set(state.month, await state.wb.readMonth(state.month));
    m = state.monthCache.get(state.month);
  } catch (e) { view.replaceChildren(h('p', { className: 'warn' }, `Could not read the ${MONTH_NAMES[state.month - 1]} sheet: ${e.message}`)); return; }
  $('#month-source').textContent = `Workbook figures (${m.name} sheet)`;
  const app = monthView(docForCheck(), state.month);
  const agrees = toCents(m.subtotal.amount) === app.totals.amountCents && toCents(m.recon.closing) === app.recon.closingCents && CATEGORIES.every(c => toCents(m.subtotal.byCategory[c.name]) === app.totals.byCategory[c.name]);
  const check = $('#month-check');
  check.textContent = agrees ? 'Matches the app calculation' : 'Differs from the app calculation: the sorted table in the workbook may be out of date. Tell the treasurer.';
  check.className = agrees ? 'ok' : 'warn';
  const money = c => fmtMoney(toCents(c));
  const table = h('table', { id: 'month-table' },
    h('thead', {},
      h('tr', {}, h('th', { colspan: 4 }), h('th', { className: 'dep-head', colspan: 4 }, 'DEPOSITS'), h('th', { className: 'wd-head', colspan: 4 }, 'WITHDRAWALS'), h('th')),
      h('tr', {}, h('th', {}, 'Date'), h('th', {}, 'Description'), h('th', { className: 'num' }, 'Amount'), h('th', {}, 'Check#'), ...CATEGORIES.map(c => h('th', { className: 'num' }, c.name)), h('th', { className: 'num' }, 'Balance'))),
    h('tbody', {}, m.rows.length ? m.rows.map(r => h('tr', {},
      h('td', { className: 'date', dataset: { label: 'Date' } }, fmtDate(r.date)),
      h('td', { className: 'title', dataset: { label: 'Description' } }, r.description),
      h('td', { className: 'num', dataset: { label: 'Amount' } }, money(r.amount)),
      h('td', { className: r.chequeNum ? '' : 'empty', dataset: { label: 'Check#' } }, r.chequeNum),
      ...CATEGORIES.map(c => r.cells[c.name] === null ? h('td', { className: 'num empty', dataset: { label: c.name } }) : h('td', { className: `num ${DEPOSIT_CATEGORIES.includes(c.name) ? 'dep' : 'wd'}`, dataset: { label: c.name } }, money(r.cells[c.name]))),
      h('td', { className: `num ${r.net < 0 ? 'neg' : ''}`, dataset: { label: 'Balance' } }, money(r.net)),
    )) : h('tr', {}, h('td', { className: 'muted title', colspan: 13 }, 'No transactions this month.'))),
    h('tfoot', {}, h('tr', { className: 'subtotal' },
      h('td', { className: 'title' }, 'SUBTOTAL'), h('td', { className: 'empty' }),
      h('td', { className: 'num', dataset: { label: 'Amount', cat: 'amount' } }, money(m.subtotal.amount)), h('td', { className: 'empty' }),
      ...CATEGORIES.map(c => h('td', { className: 'num', dataset: { label: c.name, cat: c.name } }, money(m.subtotal.byCategory[c.name]))),
      h('td', { className: `num ${m.subtotal.net < 0 ? 'neg' : ''}`, dataset: { label: 'Balance', cat: 'net' } }, money(m.subtotal.net)))));
  const recon = h('dl', { className: 'kv recon', id: 'recon' },
    h('dt', {}, 'Opening Balance'), h('dd', { dataset: { recon: 'opening' } }, money(m.recon.opening)),
    h('dt', {}, 'Add: Total Deposits'), h('dd', { dataset: { recon: 'deposits' }, className: 'dep' }, money(m.recon.deposits)),
    h('dt', {}, 'Less: Total Expenses'), h('dd', { dataset: { recon: 'expenses' }, className: 'wd' }, money(m.recon.expenses)),
    h('dt', {}, 'Closing Balance'), h('dd', { dataset: { recon: 'closing' }, className: 'total' }, money(m.recon.closing)));
  view.replaceChildren(...[
    h('h2', {}, m.title || `${m.name} ${state.snapshot.year} Transactions`, h('span', { className: 'muted small' }, `${m.rows.length} transaction${m.rows.length === 1 ? '' : 's'}`)),
    m.truncated ? h('p', { className: 'warn small' }, 'The workbook sheet shows at most 30 transactions for a month; more exist in the transaction table.') : null,
    h('div', { className: 'table-wrap cards' }, table),
    h('h3', { style: 'margin-top:1rem' }, 'BANK RECONCILIATION'), recon,
  ].filter(Boolean));   // replaceChildren would render a null as the text "null"
}
/** Compares the year totals and balances on the workbook's annual sheet with the same figures computed from the transaction table. */
function annualCheck(a) {
  const app = annualView(docForCheck());
  const rowFor = re => a.rows.find(x => re.test(String(x[0])));
  const year = re => { const r = rowFor(re); return r && r[13] !== '' && r[13] !== undefined ? toCents(r[13]) : null; };
  const bal = re => { const r = rowFor(re); return r && r[1] !== '' && r[1] !== undefined ? toCents(r[1]) : null; };
  const pairs = [
    ['Total income', year(/^Total Income/), app.yearTotalIncome], ['Total expenses', year(/^Total Expenses/), app.yearTotalExpenses], ['Net', year(/^Net/), app.yearNet],
    ['Opening balance', bal(/Opening/), app.openingCents], ['Closing balance', bal(/Closing/), app.closingCents],
  ];
  if (pairs.some(([, w]) => w === null)) return { ok: null, text: 'Could not compare with the app calculation: the annual sheet layout was not recognised.' };
  const diffs = pairs.filter(([, w, c]) => w !== c);
  return diffs.length
    ? { ok: false, text: `Differs from the app calculation (${diffs.map(([n, w, c]) => `${n}: workbook ${fmtMoney(w)}, app ${fmtMoney(c)}`).join('; ')}). The monthly sheets or the sorted table may be out of date. Tell the treasurer.` }
    : { ok: true, text: 'Matches the app calculation' };
}
async function renderAnnual() {
  const view = $('#annual-view');
  view.replaceChildren(h('p', { className: 'muted' }, 'Reading the workbook…'));
  let a;
  try { a = state.annualCache || (state.annualCache = await state.wb.readAnnual()); }
  catch (e) { view.replaceChildren(h('p', { className: 'warn' }, `Could not read the annual sheet: ${e.message}`)); return; }
  const money = v => (typeof v === 'number' ? fmtMoney(toCents(v)) : String(v ?? ''));
  const keyFor = label => /^Total Income/.test(label) ? 'income-year' : /^Total Expenses/.test(label) ? 'expenses-year' : /^Net/.test(label) ? 'net-year' : null;
  const tableRows = a.rows.filter(r => r.length >= 14 && !/Balance/.test(String(r[0])));
  const balances = a.rows.filter(r => /Balance/.test(String(r[0])));
  const check = annualCheck(a);
  const checkEl = $('#annual-check');
  checkEl.textContent = check.text;
  checkEl.className = check.ok === true ? 'ok' : check.ok === false ? 'warn' : 'muted';
  view.replaceChildren(
    h('h2', {}, `${a.year} Financial Summary`, h('span', { className: 'muted small' }, 'Workbook figures')),
    h('div', { className: 'table-wrap' }, h('table', { id: 'annual-table' },
      h('thead', {}, h('tr', {}, ...a.header.map((x, i) => h('th', { className: i ? 'num' : '' }, String(x))))),
      h('tbody', {}, ...tableRows.map(r => { const key = keyFor(String(r[0])); return h('tr', { className: key ? 'subtotal' : '' }, h('th', { scope: 'row' }, String(r[0])), ...r.slice(1, 13).map(v => h('td', { className: 'num' }, money(v))), h('td', { className: 'num', dataset: key ? { cell: key } : {} }, money(r[13]))); })))),
    h('dl', { className: 'kv recon' }, ...balances.flatMap(r => [h('dt', {}, String(r[0])), h('dd', { dataset: { cell: /Opening/.test(String(r[0])) ? 'opening' : 'closing' } }, money(r[1]))])));
}
function renderTab() {
  if (!state.snapshot) return;
  if (state.tab === 'months') renderMonths();
  if (state.tab === 'annual') renderAnnual();
  if (state.tab === 'settings') renderSettings();
}

// ---- settings -----------------------------------------------------------------------------
function buildSettings() {
  const form = $('#settings-form'), stateEl = $('#settings-state'), btn = $('#btn-save-settings');
  const input = field(form, 'priorYearBalance');
  const setState = (s, text) => { stateEl.dataset.state = s; stateEl.textContent = text; };
  form.addEventListener('input', () => { state.settingsDirty = true; setState('unsaved', state.settingsSaving ? 'Saving the previous value… what you typed since is unsaved' : 'Unsaved changes'); });
  form.addEventListener('submit', async e => {
    e.preventDefault();
    if (!writeAllowed()) return;
    if (state.settingsSaving) return;                       // no duplicate submissions
    const raw = input.value;
    state.settingsSaving = true;
    btn.disabled = true;
    setState('saving', 'Saving to the workbook…');
    try {
      const v = await withBusy(() => state.wb.setPriorYearBalance(raw, state.snapshot.priorYearBalance));
      if (input.value !== raw) setState('unsaved', `Saved ${fmtMoney(toCents(v))} to the workbook. The value you typed since is not saved yet.`);
      else { state.settingsDirty = false; setState('saved', `Saved to workbook · ${fmtMoney(toCents(v))}`); }
      await reload();
    } catch (err) {
      if (err instanceof ConflictError) setState('conflict', `Someone else changed this record: ${err.message} Press Refresh, then try again.`);
      else if (err instanceof UnresolvedOperationError) { showIncident(); setState('failed', `Could not save: ${err.message}`); }
      else setState('failed', `Could not save: ${err.message}`);
    } finally {
      state.settingsSaving = false;
      btn.disabled = false;
    }
  });
  $('#btn-signout').addEventListener('click', requestSignOut);
  $('#btn-change-workbook').addEventListener('click', async () => {
    if (hasUnsaved() && !(await confirmDialog('You have unsaved changes. Choose a different workbook anyway?', { okLabel: 'Continue', danger: true }))) return;
    forgetWorkbook(localStorage, state.who);
    location.reload();
  });
}
function renderSettings() {
  const form = $('#settings-form');
  if (!state.settingsDirty) field(form, 'priorYearBalance').value = state.snapshot.priorYearBalance.toFixed(2);
  $('#account-info').textContent = state.auth.account ? `Signed in as ${state.auth.account.name} (${state.auth.account.username})` : 'Test mode: no Microsoft account is signed in.';
  renderWorkbookInfo();
}
function renderWorkbookInfo() {
  const info = $('#workbook-info');
  const m = (state.snapshot && state.snapshot.meta) || {};
  info.replaceChildren(h('dl', { className: 'kv' },
    h('dt', {}, 'Workbook'), h('dd', {}, m.name || state.workbook.name || ''),
    h('dt', {}, 'Last changed'), h('dd', {}, m.lastModifiedDateTime ? new Date(m.lastModifiedDateTime).toLocaleString() : ''),
    h('dt', {}, 'Where'), h('dd', {}, state.mode === 'test' ? 'In-memory sample (test mode)' : (state.workbook.source === 'config' ? 'Pinned in the site configuration' : 'Chosen on this device')),
    h('dt', {}, 'Read'), h('dd', {}, state.snapshot ? new Date(state.snapshot.loadedAt).toLocaleTimeString() : ''),
    h('dt', {}, 'Writer model'), h('dd', {}, CONFIG.pilot && CONFIG.pilot.writerModel === 'single' ? 'Pilot: one designated writer at a time; simultaneous editing is not supported.' : 'not set')));
  $('#btn-change-workbook').hidden = !(state.mode === 'connected' && state.workbook.source !== 'config');
}

async function requestSignOut() {
  if (hasUnsaved()) {
    const ok = await confirmDialog('You have unsaved changes that are not in the workbook. Sign out anyway and lose them?', { okLabel: 'Sign out anyway', danger: true });
    if (!ok) return 'cancelled';
  }
  if (state.wb && state.wb.halted) {
    const ok = await confirmDialog('An unresolved change is still paused for this workbook. Its record stays in this browser for the next sign-in. Sign out anyway?', { okLabel: 'Sign out', danger: true });
    if (!ok) return 'cancelled';
  }
  clearDraftsForAccount(localStorage, state.who);
  if (state.auth.account) await state.auth.signOut();
  else location.reload();
  return 'signed-out';
}

main().catch(e => { console.error(e); toast(e.message || String(e), 'error'); });

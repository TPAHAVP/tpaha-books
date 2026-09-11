// Winter Storage UI. All calculations come from ./model.js; persistence from ../store.js.
import { CONFIG } from '../config.js';
import { initAuth } from '../auth.js';
import { Store } from '../store.js';
import { h, toast, confirmDialog, setupTabs, fmtNumberOrText, splitPhone, downloadBlob, readFileText, printWithPage } from '../ui.js';
import { SPOT_IDS, MAIN_SPOTS, END_SPOTS, TYPES, newStorage, computeAll, storageList, validateSpot, updateSpot, clearSpot, swapSpots, addReserve, updateReserve, removeReserve, activeReserve, isBlank } from './model.js';
import { exportStorageXlsx } from '../xlsx-export.js';

const $ = sel => document.querySelector(sel);
// Fields are queried by attribute: a control named "length" would otherwise clash with the collection's length property.
const field = (form, name) => form.querySelector(`[name="${name}"]`);
const state = { auth: null, store: null, backend: null, tab: 'map', editing: null };
const money = n => (typeof n === 'number') ? '$' + n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',') : (n || '');

async function main() {
  $('#app-version').textContent = `v${CONFIG.appVersion}`;
  state.auth = await initAuth(CONFIG);
  const banner = $('#mode-banner');
  banner.textContent = 'Winter Storage is not connected to Microsoft 365 yet. This screen works on a copy kept in this browser only; nothing here is saved to the workbook.';
  banner.hidden = false;
  state.store = new Store({ key: 'storage', fileName: 'winter-storage.json', createEmpty: () => newStorage(new Date().getFullYear()) });
  state.store.onChange(onStoreChange);
  buildSwap();
  buildSpotDialog();
  buildReserveButtons();
  buildSettings();
  setupTabs($('#tabs'), id => { state.tab = id; renderTab(); });
  await state.store.load();
}

function ctx() { return { user: state.auth.account ? state.auth.account.username : '' }; }

function onStoreChange(doc, status) {
  $('#sync-status').textContent = 'Browser copy only';
  if (doc) $('#storage-year').textContent = String(doc.storageYear);
  renderTab();
}

function renderTab() {
  if (!state.store || !state.store.doc) return;
  renderHeader();
  if (state.tab === 'map') renderMap();
  if (state.tab === 'list') renderList();
  if (state.tab === 'reserve') renderReserve();
  if (state.tab === 'print') renderPrint();
  if (state.tab === 'settings') renderSettings();
}

const spotOf = id => state.store.doc.spots.find(s => s.spot === id);
const isWarn = check => check === 'Over length' || check === 'Too wide';

// ---- Map --------------------------------------------------------------------------------

function renderHeader() {
  const all = computeAll(state.store.doc);
  $('#hdr-assigned').textContent = all.header.assigned;
  $('#hdr-revenue').textContent = all.header.revenue;
  $('#hdr-revenue').classList.toggle('warn', all.header.revenueCents === null);
  $('#hdr-width').textContent = all.header.widthRemaining;
  $('#hdr-width').classList.toggle('warn', /exceeds/.test(all.header.widthRemaining));
}

function spotCard(s, c) {
  const open = isBlank(s.name);
  const dims = open ? '' : `L ${fmtNumberOrText(s.length) || '___'} x W ${fmtNumberOrText(s.width) || '___'} ft · Qty ${fmtNumberOrText(s.qty) || '___'}`;
  const years = typeof c.yearsStored === 'number' ? `${c.yearsStored} yrs` : c.yearsStored;
  const fee = typeof c.totalFee === 'number' ? `${s.type || ''}  ${money(c.totalFee)}`.trim() : c.totalFee;
  return h('article', {
    className: `spot-card${open ? ' open' : ''}`, dataset: { spot: s.spot }, tabindex: '0', role: 'button',
    'aria-label': `${s.spot}${open ? ', open space' : ', ' + s.name}`,
    onClick: () => openSpotDialog(s.spot),
    onKeydown: e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openSpotDialog(s.spot); } },
  },
    h('span', { className: 'spot-id' }, s.spot),
    h('div', { className: 'spot-name' }, open ? 'Open space' : s.name),
    open ? null : h('div', {}, s.equipment),
    open ? null : h('div', { className: 'spot-meta' }, [s.phone, s.plate ? `Plate ${s.plate}` : ''].filter(Boolean).join(' · ')),
    open ? null : h('div', { className: 'spot-meta' }, dims),
    open ? null : h('div', { className: 'spot-meta' }, [s.firstYear ? `Since ${s.firstYear}` : '', years ? `${years}` : ''].filter(Boolean).join(' · ')),
    open ? null : h('div', { className: `spot-fee${typeof c.totalFee === 'number' ? '' : ' warn-text'}` }, fee),
    s.notes ? h('div', { className: 'spot-meta' }, s.notes) : null,
    h('div', { className: `spot-status${isWarn(c.sizeCheck) || /Check|Add length|Measure|needed/.test(c.statusLine) && !open ? ' warn' : ''}` }, c.statusLine));
}

function renderMap() {
  const doc = state.store.doc;
  const all = computeAll(doc);
  const rows = [];
  for (let r = 1; r <= 5; r++) {
    rows.push(h('div', { className: 'map-grid' }, ...[1, 2, 3, 4, 5].map(c => { const id = `R${r}C${c}`; return spotCard(spotOf(id), all.spots[id]); })));
  }
  const footers = h('div', { className: 'col-footers' }, ...[1, 2, 3, 4, 5].map(c =>
    h('div', { className: `col-footer${all.footers[c] === 'Add lengths' ? ' warn' : ''}`, dataset: { col: String(c) } }, all.footers[c])));
  const ends = h('div', { className: 'end-grid' }, ...END_SPOTS.map(id => spotCard(spotOf(id), all.spots[id])));
  $('#map').replaceChildren(
    h('div', { className: 'card' }, h('h2', {}, `Storage year ${doc.storageYear} · Amenities end · 5 columns x 5 rows`), h('div', { style: 'display:grid;gap:.6rem' }, ...rows),
      h('h3', { style: 'margin-top:.75rem' }, 'Remaining after main rows, before sideways end units'), footers),
    h('div', { className: 'card' }, h('h2', {}, 'Sideways end spaces'), ends,
      h('p', { className: 'small muted', style: 'margin-top:.75rem' }, 'End spaces turn sideways and need measured usable dimensions (Settings). The map is schematic, not to scale.')));
}

// ---- Swap -------------------------------------------------------------------------------

function buildSwap() {
  for (const id of ['#swap-a', '#swap-b']) {
    $(id).replaceChildren(h('option', { value: '' }, 'Choose a spot'), ...SPOT_IDS.map(s => h('option', { value: s }, s)));
  }
  $('#btn-swap').addEventListener('click', async () => {
    const a = $('#swap-a').value, b = $('#swap-b').value;
    const status = $('#swap-status');
    const dry = swapSpots(JSON.parse(JSON.stringify(state.store.doc)), a, b, {});
    if (!dry.ok) { status.textContent = dry.message; status.className = 'status-line err'; return; }
    const nameA = spotOf(a).name || 'open space', nameB = spotOf(b).name || 'open space';
    const ok = await confirmDialog(`Swap ${a} (${nameA}) with ${b} (${nameB})? Plates, first years and fees move with the customers.`, { okLabel: 'Swap' });
    if (!ok) return;
    const r = await state.store.mutate(doc => swapSpots(doc, a, b, ctx()));
    status.textContent = r.message;
    status.className = `status-line ${r.ok ? 'ok' : 'err'}`;
    if (r.ok) { $('#swap-a').value = ''; $('#swap-b').value = ''; }
  });
}

// ---- Spot dialog ------------------------------------------------------------------------

function buildSpotDialog() {
  const dlg = $('#spot-dialog'), form = $('#spot-form');
  const preview = () => {
    const doc = JSON.parse(JSON.stringify(state.store.doc));
    const v = validateSpot(Object.fromEntries(new FormData(form).entries()));
    if (!v.ok) { field(form, 'yearsStored').value = ''; field(form, 'totalFee').value = ''; return; }
    updateSpot(doc, state.editing, v.value, {});
    const c = computeAll(doc).spots[state.editing];
    field(form, 'yearsStored').value = fmtNumberOrText(c.yearsStored);
    field(form, 'totalFee').value = typeof c.totalFee === 'number' ? money(c.totalFee) : c.totalFee;
  };
  form.addEventListener('input', preview);
  form.addEventListener('submit', async e => {
    e.preventDefault();
    const v = validateSpot(Object.fromEntries(new FormData(form).entries()));
    showErrors(form, v.ok ? {} : v.errors);
    if (!v.ok) return;
    await state.store.mutate(doc => updateSpot(doc, state.editing, v.value, ctx()));
    dlg.close();
    toast(`Saved ${state.editing}`);
  });
  $('#btn-spot-cancel').addEventListener('click', () => dlg.close());
  $('#btn-spot-clear').addEventListener('click', async () => {
    const s = spotOf(state.editing);
    const ok = await confirmDialog(`Clear ${state.editing}${s.name ? ` (${s.name})` : ''}? All customer fields for this space will be emptied.`, { okLabel: 'Clear space', danger: true });
    if (!ok) return;
    await state.store.mutate(doc => clearSpot(doc, state.editing, ctx()));
    dlg.close();
    toast(`${state.editing} is now an open space`);
  });
}

function showErrors(form, errors) {
  for (const e of form.querySelectorAll('.error')) e.remove();
  for (const [name, msg] of Object.entries(errors || {})) {
    const input = field(form, name);
    if (input && input.closest('.field')) input.closest('.field').append(h('span', { className: 'error', role: 'alert' }, msg));
  }
}

function openSpotDialog(id) {
  state.editing = id;
  const s = spotOf(id), form = $('#spot-form');
  $('#spot-dialog-title').textContent = `${id} · ${s.name || 'Open space'}`;
  for (const f of ['name', 'equipment', 'phone', 'plate', 'length', 'width', 'qty', 'firstYear', 'type', 'feeOverride', 'notes']) {
    field(form, f).value = isBlank(s[f]) ? '' : String(s[f]);
  }
  showErrors(form, {});
  form.dispatchEvent(new Event('input'));
  $('#spot-dialog').showModal();
}

// ---- List -------------------------------------------------------------------------------

const LIST_COLS = [
  ['spot', 'Spot'], ['name', 'Name'], ['equipment', 'Equipment'], ['phone', 'Phone'], ['plate', 'Licence plate'], ['firstYear', 'First year'],
  ['yearsStored', 'Years stored'], ['qty', 'Qty'], ['length', 'Length (ft)'], ['width', 'Width (ft)'], ['type', 'Equipment type'], ['feeOverride', 'Fee override (CAD)'],
  ['unitFee', 'Unit fee (CAD)'], ['totalFee', 'Total fee (CAD)'], ['customerTotal', 'Customer total (CAD)'], ['notes', 'Notes / arrival'], ['sourceRow', 'Source row'],
  ['start', 'Start (ft)'], ['end', 'End (ft)'], ['sizeCheck', 'Size check'],
];
function renderList() {
  const rows = storageList(state.store.doc);
  $('#storage-list').replaceChildren(
    h('thead', {}, h('tr', {}, ...LIST_COLS.map(([k, label]) => h('th', { className: /fee|total|qty|length|width|start|end|year/i.test(k) ? 'num' : '' }, label)))),
    h('tbody', {}, ...rows.map(r => h('tr', { dataset: { spot: r.spot }, onClick: () => openSpotDialog(r.spot), style: 'cursor:pointer' },
      ...LIST_COLS.map(([k]) => h('td', {
        className: `${/fee|total|qty|length|width|start|end|year/i.test(k) ? 'num' : ''}${k === 'sizeCheck' && isWarn(r[k]) ? ' warn' : ''}`,
        dataset: { col: k },
      }, fmtNumberOrText(r[k])))))));
}

// ---- Reserve ----------------------------------------------------------------------------

const RESERVE_COLS = [['name', 'Name', 'text'], ['yearsNote', 'Years / source note', 'text'], ['phone', 'Phone', 'tel'], ['equipment', 'Equipment', 'text'],
  ['qty', 'Qty', 'decimal'], ['length', 'Length (ft)', 'decimal'], ['notes', 'Notes', 'text'], ['plate', 'Licence plate', 'text']];

function setReserveStatus(msg, err) { const el = $('#reserve-status'); el.textContent = msg; el.className = `status-line ${err ? 'err' : 'ok'}`; }

function buildReserveButtons() {
  $('#btn-add-reserve').addEventListener('click', async () => {
    const row = await state.store.mutate(doc => addReserve(doc, ctx()));
    setReserveStatus('Added one blank reserve line. Enter the new customer’s details.', false);
    renderReserve();
    const input = document.querySelector(`#reserve-table tr[data-uid="${row.uid}"] [name=name]`);
    if (input) input.focus();
  });
}

function renderReserve() {
  const rows = activeReserve(state.store.doc);
  const cell = (r, [k, label, mode]) => {
    const input = h('input', { name: k, value: isBlank(r[k]) ? '' : String(r[k]), inputmode: mode === 'text' ? null : mode, 'aria-label': `${label} for reserve line` });
    input.addEventListener('change', async () => {
      const value = {}; value[k] = input.value;
      await state.store.mutate(doc => updateReserve(doc, r.uid, value, ctx()));
    });
    return h('td', { dataset: { label } }, input);
  };
  $('#reserve-table').replaceChildren(
    h('thead', {}, h('tr', {}, ...RESERVE_COLS.map(([, label]) => h('th', {}, label)), h('th', {}, ''))),
    h('tbody', {}, rows.length ? rows.map(r => h('tr', { dataset: { uid: r.uid } }, ...RESERVE_COLS.map(c => cell(r, c)),
      h('td', { className: 'actions' }, h('button', { type: 'button', className: 'btn btn-sm btn-danger btn-remove-reserve', onClick: () => removeReserveRow(r.uid) }, 'Remove Reserve Line'))))
      : h('tr', {}, h('td', { colspan: 9, className: 'muted' }, 'The reserve list is empty. Use Add Reserve Line.'))));
}

async function removeReserveRow(uid) {
  const row = state.store.doc.reserve.find(r => r.uid === uid && !r.deleted);
  if (!row) return;
  const ok = await confirmDialog(`Remove Reserve Line${row.name ? ` for ${row.name}` : ' (blank line)'}?`, { okLabel: 'Remove', danger: true });
  if (!ok) return;
  const r = await state.store.mutate(doc => removeReserve(doc, uid, ctx()));
  setReserveStatus(r.message, !r.ok);
  renderReserve();
}

// ---- Print sheet ------------------------------------------------------------------------

function printBlock(s, c, label) {
  const open = isBlank(s.name);
  const dims = `L ${open || isBlank(s.length) ? '___' : fmtNumberOrText(s.length)} x W ${open || isBlank(s.width) ? '___' : fmtNumberOrText(s.width)} ft   Qty ${open || isBlank(s.qty) ? '___' : fmtNumberOrText(s.qty)}`;
  const fee = open ? 'Fee: ' : `${s.type || ''}  ${typeof c.totalFee === 'number' ? money(c.totalFee) : c.totalFee}`;
  return h('div', { className: 'print-block', dataset: { spot: s.spot } },
    h('div', { className: 'pb-id' }, label || s.spot),
    h('div', { className: 'pb-name' }, open ? '' : s.name),
    h('div', {}, open ? '' : s.equipment),
    h('div', {}, 'Phone: ', ...splitPhone(s.phone).flatMap((p, i) => i ? [h('br'), p] : [p])),
    h('div', {}, `Plate: ${open ? '' : s.plate}`),
    h('div', {}, dims),
    h('div', {}, fee),
    h('div', {}, `Notes: ${open ? '' : s.notes}`),
    h('div', { className: 'lines' }, h('div'), h('div'), h('div')));
}

function renderPrint() {
  const doc = state.store.doc;
  const all = computeAll(doc);
  const rows = [];
  for (let r = 1; r <= 5; r++) rows.push(h('div', { className: 'print-row' }, ...[1, 2, 3, 4, 5].map(c => { const id = `R${r}C${c}`; return printBlock(spotOf(id), all.spots[id]); })));
  $('#print-sheet').replaceChildren(
    h('h1', {}, `TPAHA WINTER STORAGE - ${doc.storageYear}`),
    h('div', { className: 'print-sub' }, 'Current assignments with space for handwritten changes. Arena arrangement only; not to scale. Amenities end at the top.'),
    ...rows,
    h('div', { className: 'print-section' }, 'SIDEWAYS END SPACES'),
    h('div', { className: 'print-row ends' }, ...END_SPOTS.map(id => printBlock(spotOf(id), all.spots[id], `${id} - sideways end space`))),
    h('div', { className: 'print-footer' }, 'Checked by: __________________________    Date: __________________    Fees shown in CAD'));
  fitPrintSheet();
}

// Shrink the on-screen preview to the available width; printing always uses full size (see CSS).
function fitPrintSheet() {
  const wrap = document.querySelector('.print-wrap'), sheet = $('#print-sheet');
  if (!wrap || !sheet) return;
  const scale = Math.min(1, (wrap.clientWidth - 4) / 984);
  sheet.style.zoom = scale < 1 ? String(scale) : '';
}
window.addEventListener('resize', () => { if (state.tab === 'print') fitPrintSheet(); });

// ---- Settings ---------------------------------------------------------------------------

function buildSettings() {
  $('#settings-form').addEventListener('submit', async e => {
    e.preventDefault();
    const form = $('#settings-form');
    const num = (name, { min = 0, whole = false, allowBlank = false } = {}) => {
      const raw = field(form, name).value.trim();
      if (raw === '') return allowBlank ? null : NaN;
      const n = Number(raw);
      return (Number.isFinite(n) && n >= min && (!whole || Number.isInteger(n))) ? n : NaN;
    };
    const values = {
      storageYear: num('storageYear', { min: 1900, whole: true }),
      rateBoat: num('rateBoat'), ratePontoon: num('ratePontoon'),
      length: num('length', { min: 0.1 }), width: num('width', { min: 0.1 }), firstRowBlockDepth: num('firstRowBlockDepth', { min: 0.1 }),
      gapBetweenRows: num('gapBetweenRows'), mainColumnWidth: num('mainColumnWidth', { min: 0.1 }), leftMargin: num('leftMargin'), rightMargin: num('rightMargin'), gapBetweenColumns: num('gapBetweenColumns'),
      end1Length: num('end1Length', { min: 0.1, allowBlank: true }), end1Depth: num('end1Depth', { min: 0.1, allowBlank: true }),
      end2Length: num('end2Length', { min: 0.1, allowBlank: true }), end2Depth: num('end2Depth', { min: 0.1, allowBlank: true }),
    };
    const bad = Object.entries(values).filter(([, v]) => Number.isNaN(v)).map(([k]) => k);
    if (bad.length) { toast(`Check these settings: ${bad.join(', ')}`, 'error'); return; }
    await state.store.mutate(doc => {
      doc.storageYear = values.storageYear;
      doc.rates = { 'Boat': values.rateBoat, 'Pontoon / Camper': values.ratePontoon };
      for (const k of ['length', 'width', 'firstRowBlockDepth', 'gapBetweenRows', 'mainColumnWidth', 'leftMargin', 'rightMargin', 'gapBetweenColumns']) doc.arena[k] = values[k];
      doc.endSpaces = { 'END-1': { length: values.end1Length, depth: values.end1Depth }, 'END-2': { length: values.end2Length, depth: values.end2Depth } };
      doc.updatedAt = new Date().toISOString(); doc.updatedBy = ctx().user;
    });
    toast('Settings saved');
  });
  $('#btn-export-xlsx').addEventListener('click', () => { try { exportStorageXlsx(state.store.doc); } catch (e) { toast(e.message, 'error'); } });
  $('#btn-export-json').addEventListener('click', () => downloadBlob('winter-storage.json', new Blob([state.store.exportJson()], { type: 'application/json' })));
  $('#btn-import').addEventListener('click', () => $('#file-import').click());
  $('#file-import').addEventListener('change', async e => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const text = await readFileText(file);
      const parsed = JSON.parse(text);
      const ok = await confirmDialog(`Replace the storage data with "${file.name}" (storage year ${parsed.storageYear}, ${(parsed.spots || []).filter(s => s.name).length} assigned spaces)? The current data will be overwritten.`, { okLabel: 'Import', danger: true });
      if (!ok) return;
      await state.store.importJson(text);
      toast('Data file imported');
    } catch (err) { toast(err.message, 'error'); }
    e.target.value = '';
  });
  $('#btn-signout').addEventListener('click', async () => { state.store.clearCache(); await state.auth.signOut(); });
  $('#btn-print').addEventListener('click', () => printWithPage('size: 11in 17in portrait; margin: 0.375in'));
  $('#btn-print-list').addEventListener('click', () => printWithPage('size: 11in 17in landscape; margin: 0.4in'));
}

function renderSettings() {
  const doc = state.store.doc, form = $('#settings-form');
  if (form.contains(document.activeElement)) return;   // do not clobber typing
  const set = (name, v) => { field(form, name).value = (v === null || v === undefined) ? '' : v; };
  set('storageYear', doc.storageYear);
  set('rateBoat', doc.rates['Boat']);
  set('ratePontoon', doc.rates['Pontoon / Camper']);
  for (const k of ['length', 'width', 'firstRowBlockDepth', 'gapBetweenRows', 'mainColumnWidth', 'leftMargin', 'rightMargin', 'gapBetweenColumns']) set(k, doc.arena[k]);
  set('end1Length', doc.endSpaces['END-1'].length); set('end1Depth', doc.endSpaces['END-1'].depth);
  set('end2Length', doc.endSpaces['END-2'].length); set('end2Depth', doc.endSpaces['END-2'].depth);
  $('#data-location').textContent = 'Kept in this browser only (not connected to the workbook yet). Download the data file to keep a copy.';
  $('#account-info').textContent = state.auth.account ? `Signed in as ${state.auth.account.name} (${state.auth.account.username})` : 'Not signed in (local mode).';
  $('#btn-signout').hidden = !state.auth.account;
}

main().catch(e => { console.error(e); toast(e.message || String(e), 'error'); });

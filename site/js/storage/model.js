// Winter Storage model: the formulas of TPAHA_Winter_Storage.xlsx (Photo_Style_Layout, Storage List)
// and the four script buttons, as pure functions on the storage document.
// Computed fields return a number where Excel shows a number and Excel's exact text otherwise.

export const TYPES = ['Boat', 'Pontoon / Camper'];
export const MAIN_SPOTS = [];
for (let r = 1; r <= 5; r++) for (let c = 1; c <= 5; c++) MAIN_SPOTS.push(`R${r}C${c}`);
export const END_SPOTS = ['END-1', 'END-2'];
export const SPOT_IDS = [...MAIN_SPOTS, ...END_SPOTS];
export const CUSTOMER_FIELDS = ['name', 'equipment', 'phone', 'plate', 'length', 'width', 'qty', 'firstYear', 'type', 'feeOverride', 'notes', 'sourceRow'];
export const TOTAL_SPOTS = SPOT_IDS.length;

export const isBlank = v => v === null || v === undefined || v === '';
const isNum = v => typeof v === 'number' && Number.isFinite(v);
const isWhole = v => isNum(v) && Math.floor(v) === v;
const nowIso = () => new Date().toISOString();

export function newUid() {
  const bytes = new Uint8Array(9);
  globalThis.crypto.getRandomValues(bytes);
  return Array.from(bytes, b => b.toString(36).padStart(2, '0')).join('').slice(0, 12);
}

function blankSpot(spot) {
  return {
    spot, name: '', equipment: '', phone: '', plate: '', length: null, width: null, qty: null, firstYear: null,
    type: '', feeOverride: null, notes: '', sourceRow: null, updatedAt: '', updatedBy: '',
  };
}

export function newStorage(storageYear) {
  return {
    schema: 1,
    storageYear: Number(storageYear),
    rates: { 'Boat': 375, 'Pontoon / Camper': 400 },
    arena: { length: 160, width: 70, firstRowBlockDepth: 36, gapBetweenRows: 3, mainColumnWidth: 9, leftMargin: 6, rightMargin: 6, gapBetweenColumns: 3 },
    endSpaces: { 'END-1': { length: null, depth: null }, 'END-2': { length: null, depth: null } },
    spots: SPOT_IDS.map(blankSpot),
    reserve: [],
    updatedAt: nowIso(),
    updatedBy: '',
  };
}

const spotOf = (doc, id) => doc.spots.find(s => s.spot === id);
const money = n => '$' + n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');

// ---- per-spot formulas ------------------------------------------------------------------

function yearsStored(doc, s) {                    // "Years stored" cell (C20 pattern)
  if (isBlank(s.name)) return '';
  if (!isWhole(doc.storageYear)) return 'Check year';
  if (isBlank(s.firstYear)) return 'Add year';
  if (!isWhole(s.firstYear)) return 'Check year';
  return Math.max(0, doc.storageYear - s.firstYear + 1);
}

function unitFee(doc, s) {                        // hidden column AI
  if (isBlank(s.name)) return '';
  if (!isBlank(s.feeOverride)) return (isNum(s.feeOverride) && s.feeOverride >= 0) ? s.feeOverride : 'Check override';
  if (s.type === 'Boat') return doc.rates['Boat'];
  if (s.type === 'Pontoon / Camper') return doc.rates['Pontoon / Camper'];
  return 'Choose rate';
}

function totalFee(s, fee) {                       // "Total fee (CAD)" cell (C22 pattern)
  if (isBlank(s.name)) return '';
  return (isNum(fee) && isNum(s.qty) && s.qty >= 1 && isWhole(s.qty)) ? fee * s.qty : 'Check fee / qty';
}

function endOf(s, start) {                        // hidden column AH
  if (!isNum(start)) return '';
  if (isBlank(s.name)) return start;
  return (isNum(s.length) && s.length > 0) ? start + s.length : '';
}

function sizeCheckMain(doc, s, end, firstRow) {   // hidden column AJ, main rows
  const a = doc.arena;
  if (isBlank(s.name)) return 'Open';
  if (!isNum(s.length) || s.length <= 0) return 'Add length';
  if (end === '') return 'Check earlier length';
  if (end > a.length || (firstRow && s.length > a.firstRowBlockDepth - a.gapBetweenRows)) return 'Over length';
  if (isBlank(s.width)) return 'Width needed';
  if (!isNum(s.width) || s.width <= 0) return 'Check width';
  if (s.width > a.mainColumnWidth) return 'Too wide';
  return 'Fits';
}

function sizeCheckEnd(doc, s, id) {               // hidden cells AJ30 / AJ31
  const e = (doc.endSpaces && doc.endSpaces[id]) || {};
  if (isBlank(s.name)) return 'Open';
  if (!isNum(s.length) || s.length <= 0) return 'Add length';
  if (isBlank(e.length) || isBlank(e.depth)) return 'Measure end space';
  if (!isNum(e.length) || !isNum(e.depth) || e.length <= 0 || e.depth <= 0) return 'Check end size';
  if (s.length > e.length) return 'Over length';
  if (isBlank(s.width)) return 'Width needed';
  if (!isNum(s.width) || s.width <= 0) return 'Check width';
  if (s.width > e.depth) return 'Too wide';
  return 'Fits';
}

function statusLineMain(s, end, check) {          // status line under each block (B26 pattern)
  if (isBlank(s.name)) return 'Open space';
  if (check === 'Over length' || check === 'Too wide') return check;
  if (end === '') return 'Add length';
  return `${end} ft from arena top`;
}

/** Every derived value for the map, list, print sheet and header. */
export function computeAll(doc) {
  const out = { spots: {}, customerTotals: {}, footers: {}, header: {} };
  const a = doc.arena;
  for (let c = 1; c <= 5; c++) {
    const ends = [];
    let firstCheck = '';
    for (let r = 1; r <= 5; r++) {
      const id = `R${r}C${c}`;
      const s = spotOf(doc, id) || blankSpot(id);
      let start;
      if (r === 1) start = 0;
      else if (r === 2) start = a.firstRowBlockDepth;
      else start = isNum(ends[r - 2]) ? ends[r - 2] + a.gapBetweenRows : '';
      const end = endOf(s, start);
      ends[r - 1] = end;
      const check = sizeCheckMain(doc, s, end, r === 1);
      if (r === 1) firstCheck = check;
      const fee = unitFee(doc, s);
      out.spots[id] = {
        yearsStored: yearsStored(doc, s), unitFee: fee, totalFee: totalFee(s, fee),
        start: isNum(start) ? start : '', end, sizeCheck: check, statusLine: statusLineMain(s, end, check),
      };
    }
    out.footers[c] = (firstCheck === 'Add length' || ends.some(e => !isNum(e)))
      ? 'Add lengths'
      : `${a.length - ends[4]} ft remaining`;
  }
  for (const id of END_SPOTS) {
    const s = spotOf(doc, id) || blankSpot(id);
    const fee = unitFee(doc, s);
    const check = sizeCheckEnd(doc, s, id);
    out.spots[id] = {
      yearsStored: yearsStored(doc, s), unitFee: fee, totalFee: totalFee(s, fee), start: '', end: '',
      sizeCheck: check, statusLine: isBlank(s.name) ? 'Open space' : check,
    };
  }
  // Customer total: SUMIF over spots with the same name; any "Check fee / qty" wins.
  for (const s of doc.spots) {
    if (isBlank(s.name)) continue;
    const t = out.spots[s.spot].totalFee;
    const cur = out.customerTotals[s.name];
    if (t === 'Check fee / qty' || cur === 'Check fee / qty') out.customerTotals[s.name] = 'Check fee / qty';
    else out.customerTotals[s.name] = (isNum(cur) ? cur : 0) + (isNum(t) ? t : 0);
  }
  const assignedSpots = doc.spots.filter(s => !isBlank(s.name));
  const fees = assignedSpots.map(s => out.spots[s.spot].totalFee);
  const numericFees = fees.filter(isNum);
  const revenue = numericFees.reduce((x, y) => x + y, 0);
  const widthLeft = a.width - (5 * a.mainColumnWidth + 4 * a.gapBetweenColumns + a.leftMargin + a.rightMargin);
  out.header = {
    assignedCount: assignedSpots.length,
    assigned: `${assignedSpots.length} of ${TOTAL_SPOTS} spaces assigned`,
    revenueCents: numericFees.length === fees.length ? Math.round(revenue * 100) : null,
    revenue: numericFees.length !== fees.length ? 'Revenue: check fee / qty' : `Revenue: ${money(revenue)} CAD`,
    widthRemaining: widthLeft < 0 ? 'Planned width exceeds arena width' : `${widthLeft} ft width remaining`,
  };
  return out;
}

/** The Storage List sheet: one row per spot, 20 columns. */
export function storageList(doc) {
  const all = computeAll(doc);
  return SPOT_IDS.map(id => {
    const s = spotOf(doc, id) || blankSpot(id);
    const c = all.spots[id];
    return {
      spot: id, name: s.name, equipment: s.equipment, phone: s.phone, plate: s.plate, firstYear: s.firstYear,
      yearsStored: c.yearsStored, qty: s.qty, length: s.length, width: s.width, type: s.type, feeOverride: s.feeOverride,
      unitFee: c.unitFee, totalFee: c.totalFee, customerTotal: isBlank(s.name) ? '' : all.customerTotals[s.name],
      notes: s.notes, sourceRow: s.sourceRow, start: c.start, end: c.end, sizeCheck: c.sizeCheck,
    };
  });
}

// ---- entry validation and actions -----------------------------------------------------

function num(v) {
  if (isBlank(v)) return null;
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/[$,\s]/g, ''));
  return Number.isFinite(n) ? n : NaN;
}

/** Mirrors the sheet's data validation. Only the fields present in `input` are validated. */
export function validateSpot(input) {
  const errors = {};
  const value = {};
  for (const f of ['name', 'equipment', 'phone', 'plate', 'notes', 'type']) if (f in input) value[f] = String(input[f] ?? '').trim();
  if ('type' in value && value.type && !TYPES.includes(value.type)) errors.type = 'Choose Boat or Pontoon / Camper.';
  for (const f of ['length', 'width']) {
    if (!(f in input)) continue;
    const n = num(input[f]);
    if (Number.isNaN(n)) errors[f] = 'Enter a valid number.';
    else if (n !== null && !(n >= 0.1)) errors[f] = 'Enter a number of 0.1 or more, or leave blank.';
    value[f] = n;
  }
  if ('qty' in input) {
    const n = num(input.qty);
    if (Number.isNaN(n)) errors.qty = 'Enter a valid number.';
    else if (n !== null && !(isWhole(n) && n >= 1)) errors.qty = 'Enter a whole number of 1 or more.';
    value.qty = n;
  }
  if ('firstYear' in input) {
    const n = num(input.firstYear);
    if (Number.isNaN(n)) errors.firstYear = 'Enter a valid year.';
    else if (n !== null && !(isWhole(n) && n >= 1900 && n <= 9999)) errors.firstYear = 'Enter a year between 1900 and 9999.';
    value.firstYear = n;
  }
  if ('feeOverride' in input) {
    const n = num(input.feeOverride);
    if (Number.isNaN(n)) errors.feeOverride = 'Enter a valid amount.';
    else if (n !== null && !(n >= 0)) errors.feeOverride = 'Enter 0 or more, or leave blank to use the rate.';
    value.feeOverride = n;
  }
  if ('sourceRow' in input) { const n = num(input.sourceRow); value.sourceRow = Number.isNaN(n) ? null : n; }
  return Object.keys(errors).length ? { ok: false, errors } : { ok: true, value };
}

function stamp(doc, target, ctx) {
  const now = ctx.now || nowIso();
  const user = ctx.user || '';
  target.updatedAt = now; target.updatedBy = user;
  doc.updatedAt = now; doc.updatedBy = user;
  return now;
}

export function updateSpot(doc, spotId, value, ctx = {}) {
  const s = spotOf(doc, spotId);
  if (!s) throw new Error(`Unknown spot ${spotId}`);
  for (const f of CUSTOMER_FIELDS) if (f in value && value[f] !== undefined) s[f] = value[f];
  stamp(doc, s, ctx);
  return s;
}

/** "To empty a space, clear its yellow customer fields." */
export function clearSpot(doc, spotId, ctx = {}) {
  const blank = blankSpot(spotId);
  const value = {};
  for (const f of CUSTOMER_FIELDS) value[f] = blank[f];
  return updateSpot(doc, spotId, value, ctx);
}

/** Complete Storage Swap. */
export function swapSpots(doc, aId, bId, ctx = {}) {
  const A = String(aId || '').trim().toUpperCase();
  const B = String(bId || '').trim().toUpperCase();
  if (!A || !B) return { ok: false, message: 'Choose both spots first.' };
  if (A === B) return { ok: false, message: 'Choose two different spots. Nothing changed.' };
  const a = spotOf(doc, A);
  const b = spotOf(doc, B);
  if (!a || !b) return { ok: false, message: 'Choose valid labels from the spot dropdowns.' };
  const savedA = {};
  const savedB = {};
  for (const f of CUSTOMER_FIELDS) { savedA[f] = a[f]; savedB[f] = b[f]; }
  for (const f of CUSTOMER_FIELDS) { a[f] = savedB[f]; b[f] = savedA[f]; }
  stamp(doc, a, ctx);
  stamp(doc, b, ctx);
  return { ok: true, message: `Swapped ${A} and ${B}. Plates, first years and fees moved with the customers.` };
}

// ---- reserve list -----------------------------------------------------------------------

export function activeReserve(doc) { return doc.reserve.filter(r => !r.deleted); }

/** Add Reserve Line. */
export function addReserve(doc, ctx = {}) {
  const row = { uid: ctx.uid || newUid(), name: '', yearsNote: '', phone: '', equipment: '', qty: null, length: null, notes: '', plate: '', deleted: false, updatedAt: '', updatedBy: '' };
  doc.reserve.push(row);
  stamp(doc, row, ctx);
  return row;
}

export function updateReserve(doc, uid, value, ctx = {}) {
  const row = doc.reserve.find(r => r.uid === uid && !r.deleted);
  if (!row) return null;
  for (const f of ['name', 'yearsNote', 'phone', 'equipment', 'notes', 'plate']) if (f in value) row[f] = String(value[f] ?? '').trim();
  for (const f of ['qty', 'length']) if (f in value) { const n = num(value[f]); row[f] = Number.isNaN(n) ? null : n; }
  stamp(doc, row, ctx);
  return row;
}

/** Remove Reserve Line. */
export function removeReserve(doc, uid, ctx = {}) {
  const row = doc.reserve.find(r => r.uid === uid && !r.deleted);
  if (!row) return { ok: false, message: 'Select a reserve line to remove.' };
  row.deleted = true;
  stamp(doc, row, ctx);
  return { ok: true, message: row.name ? `Removed the reserve line for ${row.name}.` : 'Removed the selected blank reserve line.' };
}

// ---- merge ------------------------------------------------------------------------------

const laterOf = (x, y) => (String(x.updatedAt || '') >= String(y.updatedAt || '') ? x : y);

/** Conflict merge: per-spot later wins; reserve rows by uid with sticky deletes; settings from the newer document. */
export function mergeStorage(remote, local) {
  const spots = SPOT_IDS.map(id => {
    const r = spotOf(remote, id) || blankSpot(id);
    const l = spotOf(local, id) || blankSpot(id);
    return { ...laterOf(r, l) };
  });
  const byUid = new Map((remote.reserve || []).map(r => [r.uid, { ...r }]));
  for (const l of local.reserve || []) {
    const r = byUid.get(l.uid);
    if (!r) byUid.set(l.uid, { ...l });
    else byUid.set(l.uid, { ...laterOf(r, l), deleted: Boolean(r.deleted || l.deleted) });
  }
  const newer = laterOf(remote, local);
  return {
    ...remote,
    ...newer,
    spots,
    reserve: [...byUid.values()],
    storageYear: newer.storageYear,
    rates: newer.rates,
    arena: newer.arena,
    endSpaces: newer.endSpaces,
  };
}

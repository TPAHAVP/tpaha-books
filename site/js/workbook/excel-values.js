// Excel value helpers shared by the Graph adapter and the in-memory mock: A1 addresses, date serials, display text.

const LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

export function colToIndex(letters) {
  let n = 0;
  for (const ch of String(letters).toUpperCase()) n = n * 26 + (LETTERS.indexOf(ch) + 1);
  return n - 1;
}

export function indexToCol(i) {
  let s = '';
  let n = i + 1;
  while (n > 0) { const m = (n - 1) % 26; s = LETTERS[m] + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** 'Sheet!A2:I27' | 'A2:I27' | 'B23' -> { sheet, r1, c1, r2, c2 } with 1-based rows and 0-based columns. */
export function parseAddress(address) {
  const [maybeSheet, rest] = address.includes('!') ? address.split('!') : [null, address];
  const [p, q] = rest.replace(/\$/g, '').split(':');
  const m1 = /^([A-Za-z]+)(\d+)$/.exec(p);
  const m2 = q ? /^([A-Za-z]+)(\d+)$/.exec(q) : m1;
  if (!m1 || !m2) throw new Error(`Bad address: ${address}`);
  return { sheet: maybeSheet ? maybeSheet.replace(/^'|'$/g, '') : null, c1: colToIndex(m1[1]), r1: Number(m1[2]), c2: colToIndex(m2[1]), r2: Number(m2[2]) };
}

export function rangeAddress(sheet, r1, c1, r2, c2) {
  const a = `${indexToCol(c1)}${r1}`;
  const b = `${indexToCol(c2)}${r2}`;
  const body = a === b ? a : `${a}:${b}`;
  return sheet ? `${sheet}!${body}` : body;
}

const EPOCH = Date.UTC(1899, 11, 30);   // Excel 1900 date system
export function serialToIso(n) {
  return new Date(EPOCH + Math.round(Number(n)) * 86400000).toISOString().slice(0, 10);
}
export function isoToSerial(iso) {
  const [y, m, d] = String(iso).split('-').map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 86400000);
}
export const isSerialDate = v => typeof v === 'number' && v > 20000 && v < 80000;

/** Approximates Excel's display text for the number formats this workbook uses. */
export function formatText(v, nf = 'General') {
  if (v === '' || v === null || v === undefined) return '';
  if (typeof v === 'number') {
    if (/yyyy-mm-dd/i.test(nf) && !/hh/i.test(nf)) return serialToIso(v);
    if (/#,##0\.00/.test(nf)) {
      const s = Math.abs(v).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      return (v < 0 ? '-' : '') + (nf.includes('$') ? '$' : '') + s;
    }
    if (nf === '0') return String(Math.round(v));
    return Number.isInteger(v) ? String(v) : String(Number(v.toFixed(10)));
  }
  return String(v);
}

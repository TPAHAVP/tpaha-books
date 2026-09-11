// Excel and CSV exports. Row builders are pure (unit-tested); the download functions use the vendored SheetJS (globalThis.XLSX).
import { MONTH_NAMES, CATEGORIES, DEPOSIT_CATEGORIES, WITHDRAWAL_CATEGORIES, activeTransactions, monthView, annualView, fromCents } from './ledger/model.js';
import { storageList, activeReserve } from './storage/model.js';

const MONEY_FMT = '"$"#,##0.00';
const $ = c => fromCents(c);
const cellOrBlank = c => (c === null || c === undefined) ? '' : $(c);

export function buildLedgerWorkbookRows(doc) {
  const sheets = [];
  const log = [['Txn ID', 'Date', 'Type', 'Category', 'Amount', 'Description', 'Cheque #', 'Notes', 'Entered', 'Entered by']];
  for (const t of activeTransactions(doc)) log.push([t.number, t.date, t.type, t.category, t.amount, t.description, t.chequeNum || '', t.notes || '', t.createdAt || '', t.createdBy || '']);
  sheets.push({ name: 'LOG', rows: log, moneyCols: [4], widths: [8, 12, 12, 14, 12, 40, 10, 20, 24, 18] });

  for (let m = 1; m <= 12; m++) {
    const v = monthView(doc, m);
    const rows = [
      [`${MONTH_NAMES[m - 1]} ${doc.year} Transactions`],
      ['', '', '', '', '', 'DEPOSITS', '', '', '', 'WITHDRAWALS'],
      ['Date', 'Description', 'Amount', 'Check#', '', ...DEPOSIT_CATEGORIES, ...WITHDRAWAL_CATEGORIES, '', 'Balance'],
    ];
    for (const r of v.rows) {
      rows.push([r.txn.date, r.txn.description, r.txn.amount, r.txn.chequeNum || '', '',
        ...CATEGORIES.map(c => cellOrBlank(r.cells[c.name])), '', $(r.netCents)]);
    }
    rows.push(['SUBTOTAL', '', $(v.totals.amountCents), '', '', ...CATEGORIES.map(c => $(v.totals.byCategory[c.name])), '', $(v.totals.netCents)]);
    rows.push([]);
    rows.push(['BANK RECONCILIATION']);
    rows.push(['Opening Balance', $(v.recon.openingCents)]);
    rows.push(['Add: Total Deposits', $(v.recon.depositsCents)]);
    rows.push(['Less: Total Expenses', $(v.recon.expensesCents)]);
    rows.push(['Closing Balance', $(v.recon.closingCents)]);
    sheets.push({ name: MONTH_NAMES[m - 1], rows, moneyCols: [1, 2, 5, 6, 7, 8, 9, 10, 11, 12, 14], widths: [22, 30, 12, 10, 2, 13, 12, 14, 12, 12, 13, 12, 12, 2, 13] });
  }

  const a = annualView(doc);
  const header = ['Category', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Year Total'];
  const annual = [[`${doc.year} Financial Summary`], [], header];
  for (const c of DEPOSIT_CATEGORIES) annual.push([c, ...a.byCategoryByMonth[c].map($), $(a.yearTotalByCategory[c])]);
  annual.push(['Total Income', ...a.totalIncomeByMonth.map($), $(a.yearTotalIncome)]);
  annual.push([]);
  for (const c of WITHDRAWAL_CATEGORIES) annual.push([c, ...a.byCategoryByMonth[c].map($), $(a.yearTotalByCategory[c])]);
  annual.push(['Total Expenses', ...a.totalExpensesByMonth.map($), $(a.yearTotalExpenses)]);
  annual.push([]);
  annual.push(['Net', ...a.netByMonth.map($), $(a.yearNet)]);
  annual.push([]);
  annual.push([`Year Opening Balance (Jan 1, ${doc.year})`, $(a.openingCents)]);
  annual.push([`Year Closing Balance (Dec 31, ${doc.year})`, $(a.closingCents)]);
  sheets.push({ name: `Annual ${doc.year}`, rows: annual, moneyCols: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13], widths: [34, ...Array(13).fill(12)] });
  return { sheets };
}

export function buildStorageWorkbookRows(doc) {
  const list = [['Spot', 'Name', 'Equipment', 'Phone', 'Licence plate', 'First year', 'Years stored', 'Qty', 'Length (ft)', 'Width (ft)', 'Equipment type',
    'Fee override (CAD)', 'Unit fee (CAD)', 'Total fee (CAD)', 'Customer total (CAD)', 'Notes / arrival', 'Source row', 'Start (ft)', 'End (ft)', 'Size check']];
  const blank = v => (v === null || v === undefined) ? '' : v;
  for (const r of storageList(doc)) {
    list.push([r.spot, r.name, r.equipment, r.phone, r.plate, blank(r.firstYear), blank(r.yearsStored), blank(r.qty), blank(r.length), blank(r.width), r.type,
      blank(r.feeOverride), blank(r.unitFee), blank(r.totalFee), blank(r.customerTotal), r.notes, blank(r.sourceRow), blank(r.start), blank(r.end), r.sizeCheck]);
  }
  const reserve = [['Name', 'Years / source note', 'Phone', 'Equipment', 'Qty', 'Length (ft)', 'Notes', 'Licence plate']];
  for (const r of activeReserve(doc)) reserve.push([r.name, r.yearsNote, r.phone, r.equipment, blank(r.qty), blank(r.length), r.notes, r.plate]);
  return { sheets: [
    { name: 'Storage List', rows: list, moneyCols: [11, 12, 13, 14], widths: [8, 22, 24, 22, 14, 10, 10, 6, 10, 10, 16, 12, 12, 12, 14, 24, 10, 10, 10, 18] },
    { name: 'Reserve List', rows: reserve, moneyCols: [], widths: [22, 18, 22, 24, 6, 10, 24, 14] },
  ] };
}

export function toCsv(rows) {
  const esc = v => {
    if (v === null || v === undefined) return '';
    const s = String(v);
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  return rows.map(r => r.map(esc).join(',')).join('\r\n') + '\r\n';
}

// ---- browser side -----------------------------------------------------------------------

function toWorkbook(spec) {
  const XLSX = globalThis.XLSX;
  if (!XLSX) throw new Error('The Excel library did not load.');
  const wb = XLSX.utils.book_new();
  for (const sh of spec.sheets) {
    const ws = XLSX.utils.aoa_to_sheet(sh.rows);
    if (sh.widths) ws['!cols'] = sh.widths.map(w => ({ wch: w }));
    for (const addr of Object.keys(ws)) {
      if (addr[0] === '!') continue;
      const cell = ws[addr];
      const col = XLSX.utils.decode_cell(addr).c;
      if (cell.t === 'n' && sh.moneyCols && sh.moneyCols.includes(col)) cell.z = MONEY_FMT;
    }
    XLSX.utils.book_append_sheet(wb, ws, sh.name.slice(0, 31));
  }
  return wb;
}

const stamp = () => new Date().toISOString().slice(0, 10).replace(/-/g, '');

export function exportLedgerXlsx(doc) {
  globalThis.XLSX.writeFile(toWorkbook(buildLedgerWorkbookRows(doc)), `TPAHA_Ledger_${doc.year}_${stamp()}.xlsx`);
}
export function exportLedgerCsv(doc, download) {
  const rows = buildLedgerWorkbookRows(doc).sheets[0].rows;
  download(`TPAHA_Ledger_${doc.year}_transactions.csv`, new Blob(['﻿' + toCsv(rows)], { type: 'text/csv;charset=utf-8' }));
}
export function exportStorageXlsx(doc) {
  globalThis.XLSX.writeFile(toWorkbook(buildStorageWorkbookRows(doc)), `TPAHA_Winter_Storage_${doc.storageYear}.xlsx`);
}

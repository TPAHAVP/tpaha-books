// A fully synthetic workbook for TEST MODE and for automated tests. Nothing here comes from the
// real TPAHA ledger: the amounts, dates, cheque numbers and descriptions are made up.
// Shape matches what tools/build_seed.py produces from a real workbook (dates as Excel serials).
const D = (iso) => { const [y, m, d] = iso.split('-').map(Number); return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000); };
const row = (id, ts, date, type, category, amount, description, chequeNum = '', notes = '') => [id, ts, D(date), type, category, amount, description, chequeNum, notes];

const LOG_ROWS = [
  ['', '', '', '', '', '', '', '', ''],   // the real table's first body row is blank
  row(1, '2026-01-06T15:00:00.000Z', '2026-01-05', 'Deposit', 'Membership', 100, 'Sample membership'),
  row(2, '2026-01-21T15:00:00.000Z', '2026-01-20', 'Withdrawal', 'MB Hydro', 80.25, 'Sample hydro'),
  row(3, '2026-02-04T15:00:00.000Z', '2026-02-03', 'Deposit', 'Grub Box', 45.5, 'Sample grub box', 101),
  row(4, '2026-03-01T15:00:00.000Z', '2026-02-28', 'Withdrawal', 'Bank Chgs', 8, 'Sample bank fee'),
  row(5, '2026-03-16T15:00:00.000Z', '2026-03-15', 'Deposit', 'Arena Storage', 375, 'Sample storage fee'),
  row(6, '2026-03-17T15:00:00.000Z', '2026-03-16', 'Withdrawal', 'Taxes/Ins', 250, 'Sample insurance', 102),
  row(7, '2026-06-05T15:00:00.000Z', '2026-06-04', 'Withdrawal', 'Misc Exp', 120.5, 'Sample supplies', 103),
  row(8, '2026-07-11T15:00:00.000Z', '2026-07-10', 'Deposit', 'Misc Income', 500, 'Sample grant'),
  row(9, '2026-07-11T15:05:00.000Z', '2026-07-10', 'Deposit', 'Grub Box', 60, 'Sample grub box 2'),
  row(10, '2026-07-13T15:00:00.000Z', '2026-07-12', 'Withdrawal', 'Misc Exp', 300, 'Sample clinic'),
  row(11, '2026-08-15T15:00:00.000Z', '2026-08-14', 'Withdrawal', 'MB Hydro', 56.25, 'Sample hydro 2'),
  row(13, '2026-08-21T15:00:00.000Z', '2026-08-20', 'Deposit', 'Membership', 75, 'Sample membership 2'),   // id 12 deliberately unused: next id is 14
];
const SORTED_ROWS = LOG_ROWS.filter(r => r[2] !== '').slice().sort((a, b) => a[2] - b[2]);

export const SAMPLE_WORKBOOK = {
  source: 'synthetic sample (no real data)',
  year: 2026,
  priorYearBalance: 5000,
  logHeader: ['TransactionID', 'Timestamp', 'Date', 'Type', 'Category', 'Amount', 'Description', 'ChequeNum', 'Notes'],
  logRows: LOG_ROWS,
  logSortedRows: SORTED_ROWS,
  categories: [
    [1, 'Membership', 'Deposit', 'F'], [2, 'Grub Box', 'Deposit', 'G'], [3, 'Arena Storage', 'Deposit', 'H'], [4, 'Misc Income', 'Deposit', 'I'],
    [5, 'MB Hydro', 'Withdrawal', 'J'], [6, 'Taxes/Ins', 'Withdrawal', 'K'], [8, 'Misc Exp', 'Withdrawal', 'L'], [9, 'Bank Chgs', 'Withdrawal', 'M'],
    ['', '', '', ''],
  ],
  numberFormats: { Timestamp: 'yyyy-mm-dd hh:mm:ss', Date: 'yyyy-mm-dd', Amount: '$#,##0.00' },
  expectedNextId: 14,
};

// Handy expected figures for tests (computed by hand from the rows above):
//   January: total 180.25, Membership 100, MB Hydro 80.25, opening 5000, closing 5019.75
//   July: 3 rows, total 860, Misc Income 500, Grub Box 60, Misc Exp 300, net 260, opening 5061.75, closing 5321.75
//   Year: income 1155.50, expenses 815.00, net 340.50, closing 5340.50

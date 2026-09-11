// Treasurer Ledger model: every formula from TPAHA_2026.xlsx as pure functions on the ledger document.
// Money is handled in integer cents; documents store amounts as numbers with two decimals.

export const CATEGORIES = [
  { id: 1, name: 'Membership', type: 'Deposit' },
  { id: 2, name: 'Grub Box', type: 'Deposit' },
  { id: 3, name: 'Arena Storage', type: 'Deposit' },
  { id: 4, name: 'Misc Income', type: 'Deposit' },
  { id: 5, name: 'MB Hydro', type: 'Withdrawal' },
  { id: 6, name: 'Taxes/Ins', type: 'Withdrawal' },
  { id: 8, name: 'Misc Exp', type: 'Withdrawal' },
  { id: 9, name: 'Bank Chgs', type: 'Withdrawal' },
];
export const DEPOSIT_CATEGORIES = CATEGORIES.filter(c => c.type === 'Deposit').map(c => c.name);
export const WITHDRAWAL_CATEGORIES = CATEGORIES.filter(c => c.type === 'Withdrawal').map(c => c.name);
export const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

export const toCents = n => Math.round(Number(n) * 100);
export const fromCents = c => Math.round(c) / 100;

const nowIso = () => new Date().toISOString();

export function newUid() {
  const bytes = new Uint8Array(9);
  globalThis.crypto.getRandomValues(bytes);
  return Array.from(bytes, b => b.toString(36).padStart(2, '0')).join('').slice(0, 12);
}

export function newLedger(year, priorYearBalance = 0) {
  return {
    schema: 1,
    year: Number(year),
    priorYearBalance: fromCents(toCents(priorYearBalance)),
    categories: CATEGORIES.map(c => ({ ...c })),
    nextNumber: 1,
    transactions: [],
    updatedAt: nowIso(),
    updatedBy: '',
  };
}

const isoDateRe = /^\d{4}-\d{2}-\d{2}$/;
function validDate(s) {
  if (!isoDateRe.test(s || '')) return false;
  const d = new Date(s + 'T00:00:00Z');
  return !Number.isNaN(d.getTime()) && d.toISOString().slice(0, 10) === s;
}

/** Mirrors the ENTRY sheet's data validation rules. */
export function validateEntry(doc, input) {
  const errors = {};
  const date = String(input.date || '').trim();
  if (!validDate(date)) errors.date = 'Enter a valid date.';
  else if (Number(date.slice(0, 4)) !== Number(doc.year)) errors.date = `Date must fall within ${doc.year}.`;

  const type = input.type;
  if (type !== 'Deposit' && type !== 'Withdrawal') errors.type = "Type must be either 'Deposit' or 'Withdrawal'.";

  const category = input.category;
  const allowed = type === 'Deposit' ? DEPOSIT_CATEGORIES : type === 'Withdrawal' ? WITHDRAWAL_CATEGORIES : [];
  if (!allowed.includes(category)) errors.category = 'Please choose a category from the list.';

  const amountNum = Number(String(input.amount ?? '').replace(/[$,\s]/g, ''));
  if (!Number.isFinite(amountNum) || !(amountNum > 0)) errors.amount = 'Amount must be greater than 0.';

  const description = String(input.description || '').trim();
  if (!description) errors.description = 'Enter a description.';

  if (Object.keys(errors).length) return { ok: false, errors };
  return {
    ok: true,
    value: {
      date, type, category,
      amount: fromCents(toCents(amountNum)),
      description,
      chequeNum: String(input.chequeNum ?? '').trim(),
      notes: String(input.notes ?? '').trim(),
    },
  };
}

function maxNumber(doc) {
  return doc.transactions.reduce((m, t) => Math.max(m, Number(t.number) || 0), 0);
}

/** Submit Entry: appends a transaction (mutates doc) and returns it. */
export function addTransaction(doc, value, ctx = {}) {
  const now = ctx.now || nowIso();
  const user = ctx.user || '';
  const number = Math.max(Number(doc.nextNumber) || 1, maxNumber(doc) + 1);
  const txn = {
    uid: ctx.uid || newUid(),
    number,
    legacyId: null,
    date: value.date,
    type: value.type,
    category: value.category,
    amount: fromCents(toCents(value.amount)),
    description: value.description,
    chequeNum: value.chequeNum ?? '',
    notes: value.notes ?? '',
    createdAt: now,
    createdBy: user,
    updatedAt: now,
    updatedBy: user,
    deleted: false,
  };
  doc.transactions.push(txn);
  doc.nextNumber = number + 1;
  doc.updatedAt = now;
  doc.updatedBy = user;
  return txn;
}

/** Delete Transaction: marks the transaction with this Txn ID as deleted (tombstone). */
export function deleteTransaction(doc, number, ctx = {}) {
  const txn = doc.transactions.find(t => t.number === Number(number) && !t.deleted);
  if (!txn) return null;
  const now = ctx.now || nowIso();
  txn.deleted = true;
  txn.updatedAt = now;
  txn.updatedBy = ctx.user || '';
  doc.updatedAt = now;
  doc.updatedBy = ctx.user || '';
  return txn;
}

/** LOG_Sorted order: by date, then by entry time, then by number. */
export function compareTxn(a, b) {
  return a.date.localeCompare(b.date)
    || String(a.createdAt).localeCompare(String(b.createdAt))
    || (a.number - b.number);
}

export function activeTransactions(doc) {
  return doc.transactions.filter(t => !t.deleted).sort(compareTxn);
}

/** Find Transactions: inclusive date range. */
export function findTransactions(doc, fromISO, toISO) {
  return activeTransactions(doc).filter(t => (!fromISO || t.date >= fromISO) && (!toISO || t.date <= toISO));
}

const monthOf = txn => Number(txn.date.slice(5, 7));

function emptyTotals() {
  const byCategory = {};
  for (const c of CATEGORIES) byCategory[c.name] = 0;
  return { amountCents: 0, byCategory, depositsCents: 0, withdrawalsCents: 0, netCents: 0 };
}

function monthTotals(doc, month) {
  const totals = emptyTotals();
  for (const t of doc.transactions) {
    if (t.deleted || monthOf(t) !== month) continue;
    const c = toCents(t.amount);
    totals.amountCents += c;
    if (t.category in totals.byCategory) totals.byCategory[t.category] += c;
  }
  totals.depositsCents = DEPOSIT_CATEGORIES.reduce((s, n) => s + totals.byCategory[n], 0);
  totals.withdrawalsCents = WITHDRAWAL_CATEGORIES.reduce((s, n) => s + totals.byCategory[n], 0);
  totals.netCents = totals.depositsCents - totals.withdrawalsCents;
  return totals;
}

function openingFor(doc, month) {
  let balance = toCents(doc.priorYearBalance || 0);
  for (let m = 1; m < month; m++) balance += monthTotals(doc, m).netCents;
  return balance;
}

/** One month sheet: rows, SUBTOTAL row and BANK RECONCILIATION block. */
export function monthView(doc, month) {
  month = Number(month);
  const rows = activeTransactions(doc).filter(t => monthOf(t) === month).map(txn => {
    const cells = {};
    for (const c of CATEGORIES) cells[c.name] = null;
    const cents = toCents(txn.amount);
    if (txn.category in cells) cells[txn.category] = cents;
    const netCents = DEPOSIT_CATEGORIES.includes(txn.category) ? cents : -cents;
    return { txn, cells, netCents };
  });
  const totals = monthTotals(doc, month);
  const openingCents = openingFor(doc, month);
  const recon = {
    openingCents,
    depositsCents: totals.depositsCents,
    expensesCents: totals.withdrawalsCents,
    closingCents: openingCents + totals.depositsCents - totals.withdrawalsCents,
  };
  return { month, name: MONTH_NAMES[month - 1], rows, totals, recon };
}

/** Annual summary with correct year totals (the workbook's Year Total column has two formula errors). */
export function annualView(doc) {
  const months = Array.from({ length: 12 }, (_, i) => monthTotals(doc, i + 1));
  const byCategoryByMonth = {};
  const yearTotalByCategory = {};
  for (const c of CATEGORIES) {
    byCategoryByMonth[c.name] = months.map(m => m.byCategory[c.name]);
    yearTotalByCategory[c.name] = byCategoryByMonth[c.name].reduce((a, b) => a + b, 0);
  }
  const totalIncomeByMonth = months.map(m => m.depositsCents);
  const totalExpensesByMonth = months.map(m => m.withdrawalsCents);
  const netByMonth = months.map(m => m.netCents);
  const yearTotalIncome = totalIncomeByMonth.reduce((a, b) => a + b, 0);
  const yearTotalExpenses = totalExpensesByMonth.reduce((a, b) => a + b, 0);
  const openingCents = toCents(doc.priorYearBalance || 0);
  return {
    byCategoryByMonth, yearTotalByCategory, totalIncomeByMonth, totalExpensesByMonth, netByMonth,
    yearTotalIncome, yearTotalExpenses,
    yearNet: yearTotalIncome - yearTotalExpenses,
    openingCents,
    closingCents: openingCents + yearTotalIncome - yearTotalExpenses,
  };
}

/**
 * Conflict merge: union by uid; the later updatedAt wins; a delete on either side sticks.
 * Two different transactions that share a Txn ID keep the earlier one and renumber the later one.
 */
export function mergeLedgers(remote, local) {
  const notices = [];
  const byUid = new Map();
  for (const t of remote.transactions) byUid.set(t.uid, { ...t });
  for (const t of local.transactions) {
    const r = byUid.get(t.uid);
    if (!r) { byUid.set(t.uid, { ...t }); continue; }
    const winner = String(t.updatedAt) > String(r.updatedAt) ? t : r;
    byUid.set(t.uid, { ...winner, deleted: Boolean(r.deleted || t.deleted) });
  }
  const merged = [...byUid.values()].sort((a, b) =>
    String(a.createdAt).localeCompare(String(b.createdAt)) || (a.number - b.number));
  let max = merged.reduce((m, t) => Math.max(m, Number(t.number) || 0), 0);
  const seen = new Set();
  for (const t of merged) {
    if (seen.has(t.number)) {
      const old = t.number;
      t.number = ++max;
      notices.push(`Transaction "${t.description}" (${t.date}) was renumbered from #${old} to #${t.number} because two devices added #${old} at the same time.`);
    }
    seen.add(t.number);
  }
  const newer = String(local.updatedAt || '') >= String(remote.updatedAt || '') ? local : remote;
  const doc = {
    ...remote,
    ...newer,
    transactions: merged,
    nextNumber: max + 1,
    priorYearBalance: newer.priorYearBalance,
    categories: newer.categories || remote.categories,
  };
  return { doc, notices };
}

/** Start next year: a fresh ledger whose prior-year balance is this year's closing balance. */
export function nextYearLedger(doc, ctx = {}) {
  const closing = annualView(doc).closingCents;
  const next = newLedger(Number(doc.year) + 1, fromCents(closing));
  next.updatedAt = ctx.now || nowIso();
  next.updatedBy = ctx.user || '';
  return next;
}

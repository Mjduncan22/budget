// Port of Financials.kt: income statement, balance sheet, cash flow, realized gains, tax summary, donations, CSV.
import { T, category, needsReview } from "./analytics.js";
import { lc, eqi, sum, groupBy, monthLabel, monthYear, monthNum, endOfMonth, startOfMonth, todayEpochDay, currentMonth, ymd, monthOfDay, money, qty } from "./util.js";

const L = (label, amount = null, level = 0, bold = false, note = "") => ({ label, amount, level, bold, note });

export function monthsIn(month, scope) {
  if (scope === "MONTH") return [month];
  const y = monthYear(month);
  return Array.from({ length: monthNum(month) }, (_, i) => `${y}-${String(i + 1).padStart(2, "0")}`);
}

export function periodLabel(month, scope) {
  return scope === "MONTH" ? monthLabel(month) : `Jan – ${monthLabel(month, true)} ${monthYear(month)} (year to date)`;
}

const typeOf = (s, cat) => category(s, cat)?.type || T.REVIEW;

const CARD_WORDS = /credit|card|visa|mastercard|amex|american express|discover|capital one|savor|quicksilver|sapphire|freedom/;

export function accountKind(s, account) {
  const t = lc(s.accounts.find((a) => eqi(a.name, account))?.type);
  const n = lc(account);
  if (t.includes("credit") || ["capital one", "discover", "credit one"].includes(n)) return "CARD";
  if (t.includes("brokerage") || n.includes("investment") || n.includes("brokerage") || n.includes("401k") || n.includes("ira")) return "INVESTMENT";
  if (t.includes("wallet") || t.includes("exchange")) return "OTHER";
  if (!t) {
    // No type in the Accounts tab: guess from the name, then from SimpleFIN's sub-account name
    const subs = s.balances.filter((b) => eqi(b.account, account)).map((b) => lc(b.subAccount)).join(" ");
    if (CARD_WORDS.test(n) || CARD_WORDS.test(subs)) return "CARD";
  }
  return "CASH";
}

const inPeriod = (s, months) => { const set = new Set(months); return s.transactions.filter((t) => set.has(t.budgetMonth)); };

export function balanceAt(s, account, day) {
  const rows = s.balances.filter((b) => eqi(b.account, account));
  if (!rows.length) return null;
  const snapDay = Math.max(...rows.map((r) => r.date));
  const snap = sum(rows.filter((r) => r.date === snapDay), (r) => r.balance);
  const txs = s.transactions.filter((t) => eqi(t.account, account));
  return day >= snapDay ? snap + sum(txs.filter((t) => t.date > snapDay && t.date <= day), (t) => t.amount)
    : snap - sum(txs.filter((t) => t.date > day && t.date <= snapDay), (t) => t.amount);
}

function accountsOf(s) {
  const seen = new Set(), out = [];
  for (const a of [...s.accounts.map((x) => x.name), ...s.transactions.map((t) => t.account), ...s.balances.map((b) => b.account)]) {
    if (!a || seen.has(lc(a))) continue;
    seen.add(lc(a)); out.push(a);
  }
  return out;
}

export function incomeStatement(s, month, scope) {
  const txs = inPeriod(s, monthsIn(month, scope));
  const byCat = groupBy(txs, (t) => lc(t.category));
  const total = (c) => sum(byCat.get(lc(c.name)) || [], (t) => t.amount);
  const lines = [L("Revenue", null, 0, true)];
  let revenue = 0;
  for (const c of s.categories.filter((c) => c.type === T.INCOME)) { const v = total(c); if (v !== 0) { lines.push(L(c.name, v, 1)); revenue += v; } }
  lines.push(L("Total revenue", revenue, 0, true));
  lines.push(L("Expenses", null, 0, true));
  let expenses = 0;
  for (const [g, cats] of groupBy(s.categories.filter((c) => c.type === T.EXPENSE), (c) => c.group)) {
    const gt = sum(cats, (c) => 0 - total(c));
    if (gt === 0) continue;
    lines.push(L(g, gt, 1, true));
    for (const c of cats) { const v = 0 - total(c); if (v !== 0) lines.push(L(c.name, v, 2)); }
    expenses += gt;
  }
  lines.push(L("Total expenses", expenses, 0, true));
  lines.push(L("Net income", revenue - expenses, 0, true));
  const invested = sum(s.categories.filter((c) => c.type === T.INVESTMENT), (c) => 0 - total(c));
  const review = txs.filter((t) => needsReview(s, t));
  if (invested !== 0) lines.push(L("Invested (not an expense)", invested, 0, false, "crypto & stock purchases"));
  if (review.length) lines.push(L("Not yet categorized", sum(review, (t) => t.amount), 0, false, `${review.length} transactions — finish Bank Rec for an exact statement`));
  return { title: "Income statement", subtitle: periodLabel(month, scope), lines,
    footnote: "Card purchases count when you make them; card payments are transfers, so nothing is counted twice." };
}

export function balanceSheet(s, month, quotes) {
  const today = todayEpochDay();
  const end = Math.min(endOfMonth(month), today);
  const current = month === currentMonth();
  const lines = [], missing = [];
  const accounts = accountsOf(s);
  const section = (kind) => accounts.filter((a) => accountKind(s, a) === kind).map((a) => {
    const b = balanceAt(s, a, end);
    if (b == null) { if (s.transactions.some((t) => t.account === a && t.date > end - 120)) missing.push(a); return null; }
    return [a, b];
  }).filter(Boolean);
  const cash = section("CASH"), brokerage = section("INVESTMENT"), cards = section("CARD");
  // Holdings inside an account that already has a bank/brokerage balance (e.g. SPAXX in Fidelity Checking) are counted by that balance.
  const brokerageNames = new Set([...brokerage, ...cash].map((b) => lc(b[0])));
  const held = s.holdings.filter((h) => !brokerageNames.has(lc(h.heldIn)));
  const value = (h) => (current && quotes[h.symbol] ? quotes[h.symbol].price * h.quantity : h.costBasis);
  const cashH = held.filter((h) => eqi(h.assetType, "Cash"));
  const crypto = held.filter((h) => eqi(h.assetType, "Crypto"));
  const otherInv = held.filter((h) => !eqi(h.assetType, "Crypto") && !eqi(h.assetType, "Cash"));

  lines.push(L("Assets", null, 0, true));
  lines.push(L("Cash", sum(cash, (x) => x[1]) + sum(cashH, value), 1, true));
  cash.forEach((x) => lines.push(L(x[0], x[1], 2)));
  cashH.forEach((h) => lines.push(L(`${h.symbol} at ${h.heldIn}`, value(h), 2)));
  const invTotal = sum(brokerage, (x) => x[1]) + sum(otherInv, value);
  lines.push(L("Investments", invTotal, 1, true));
  brokerage.forEach((x) => lines.push(L(x[0], x[1], 2)));
  for (const [sym, hs] of groupBy(otherInv, (h) => h.symbol)) lines.push(L(sym, sum(hs, value), 2));
  const cryptoTotal = sum(crypto, value);
  lines.push(L("Crypto", cryptoTotal, 1, true, current ? "live prices" : "at cost"));
  const bySym = [...groupBy(crypto, (h) => h.symbol)].map(([k, hs]) => [k, sum(hs, value)]).sort((a, b) => b[1] - a[1]);
  bySym.slice(0, 8).forEach(([k, v]) => lines.push(L(k, v, 2)));
  if (bySym.length > 8) lines.push(L("Other coins", sum(bySym.slice(8), (x) => x[1]), 2));
  const assets = sum(cash, (x) => x[1]) + sum(cashH, value) + invTotal + cryptoTotal;
  lines.push(L("Total assets", assets, 0, true));
  lines.push(L("Liabilities", null, 0, true));
  cards.forEach((x) => lines.push(L(x[0], -x[1], 1)));
  const liabilities = sum(cards, (x) => -x[1]);
  lines.push(L("Total credit card balances", liabilities, 0, true));
  lines.push(L("Net worth", assets - liabilities, 0, true));
  let foot = `As of ${ymd(end)}. Bank and card balances come from SimpleFIN, worked back to month-end from your transactions.`;
  if (!s.balances.length) foot += " No balances recorded yet — run the sync once to fill this in.";
  if (missing.length) foot += ` No balance on file for: ${missing.join(", ")}.`;
  return { title: "Balance sheet", subtitle: `End of ${monthLabel(month)}`, lines, footnote: foot };
}

export function cashFlow(s, month, scope) {
  const months = monthsIn(month, scope);
  const txs = inPeriod(s, months);
  const sumType = (ty) => sum(txs.filter((t) => typeOf(s, t.category) === ty), (t) => t.amount);
  const income = sumType(T.INCOME), spending = 0 - sumType(T.EXPENSE), invested = 0 - sumType(T.INVESTMENT);
  const review = sum(txs.filter((t) => needsReview(s, t)), (t) => t.amount);
  const catSum = (name) => sum(txs.filter((t) => eqi(t.category, name)), (t) => t.amount);
  const ccCredit = catSum("Credit Card Credit"), ccDebit = catSum("Credit Card Debit"), netTransfers = catSum("Net Transfers");
  const chargedOnCards = 0 - sum(txs.filter((t) => accountKind(s, t.account) === "CARD" && typeOf(s, t.category) === T.EXPENSE), (t) => t.amount);
  const byKind = {};
  for (const t of txs) { const k = accountKind(s, t.account); byKind[k] = (byKind[k] || 0) + t.amount; }

  const lines = [];
  lines.push(L("1 · True income vs spending", null, 0, true));
  lines.push(L("Income", income, 1));
  lines.push(L("Spending", -spending, 1, false, `${money(chargedOnCards)} of it charged on credit cards`));
  lines.push(L("Invested in crypto / stocks", -invested, 1));
  const fromIncome = income - spending - invested;
  lines.push(L("Left over (saved)", fromIncome, 0, true));
  lines.push(L("2 · Credit cards", null, 0, true));
  lines.push(L("New charges", -chargedOnCards, 1));
  lines.push(L("Payments sent from your bank", ccCredit, 1));
  lines.push(L("Payments received by the cards", ccDebit, 1));
  const timing = ccCredit + ccDebit;
  lines.push(L("Payment timing difference", timing, 1, false, Math.abs(timing) < 1 ? "payments match" : "a payment left the bank in one period and posted to the card in another"));
  const cardChange = byKind.CARD || 0;
  lines.push(L("Change in what you owe", -cardChange, 0, true, cardChange < 0 ? "owe more" : "owe less"));
  lines.push(L("3 · Where the money went", null, 0, true));
  lines.push(L("Change in bank / cash accounts", byKind.CASH || 0, 1));
  lines.push(L("Change in brokerage cash", byKind.INVESTMENT || 0, 1));
  lines.push(L("Change in credit card balances", cardChange, 1));
  if (byKind.OTHER) lines.push(L("Other accounts", byKind.OTHER, 1));
  const total = sum(Object.values(byKind));
  lines.push(L("Total change", total, 0, true));
  lines.push(L("4 · Reconciliation", null, 0, true));
  lines.push(L("Left over from section 1", fromIncome, 1));
  lines.push(L("+ Card payment timing", timing, 1));
  lines.push(L("+ Transfers between accounts (net)", netTransfers, 1, false, Math.abs(netTransfers) < 1 ? "" : "transfers that don't net to zero — e.g. money moved to an account the app doesn't track"));
  if (review !== 0) lines.push(L("+ Not yet categorized", review, 1, false, "finish Bank Rec"));
  const explained = fromIncome + timing + netTransfers + review;
  lines.push(L("= Total change (section 3)", explained, 0, true));
  const diff = total - explained;
  lines.push(L("Unexplained difference", diff, 0, true, Math.abs(diff) < 0.01 ? "reconciled ✓" : "check categories"));
  if (s.balances.length) {
    const first = startOfMonth(months[0]) - 1;
    const last = Math.min(endOfMonth(months[months.length - 1]), todayEpochDay());
    const cashAccts = accountsOf(s).filter((a) => accountKind(s, a) === "CASH");
    const open = sum(cashAccts.map((a) => balanceAt(s, a, first)).filter((x) => x != null));
    const close = sum(cashAccts.map((a) => balanceAt(s, a, last)).filter((x) => x != null));
    lines.push(L("Bank balances", null, 0, true));
    lines.push(L(`Opening (${ymd(first)})`, open, 1));
    lines.push(L(`Closing (${ymd(last)})`, close, 1));
  }
  return { title: "Cash flow reconciliation", subtitle: periodLabel(month, scope), lines,
    footnote: "Every dollar in or out of a tracked account is either income, spending, investing, a transfer or uncategorized — so section 4 always ties to section 3." };
}

const gainOf = (sale) => (sale.costBasis == null ? null : sale.proceeds - sale.costBasis);

export function realizedGains(s, month, scope) {
  const months = new Set(monthsIn(month, scope));
  const sales = s.sales.filter((x) => months.has(monthOfDay(x.date))).sort((a, b) => a.date - b.date);
  const recorded = new Set(sales.map((x) => x.source));
  const pending = inPeriod(s, [...months]).filter((t) => lc(t.description).includes("you sold") && !recorded.has(t.id));
  const lines = [];
  const block = (title, list) => {
    if (!list.length) return;
    lines.push(L(title, null, 0, true));
    for (const x of list) lines.push(L(`${ymd(x.date)} · ${x.symbol} · ${qty(x.quantity)}`, gainOf(x), 1, false,
      `proceeds ${money(x.proceeds)}` + (x.costBasis != null ? ` · cost ${money(x.costBasis)}` : " · cost unknown")));
    lines.push(L(`${title} gain`, sum(list.map(gainOf).filter((g) => g != null)), 0, true));
  };
  const isCrypto = (x) => s.holdings.some((h) => h.symbol === x.symbol && eqi(h.assetType, "Crypto")) || eqi(x.account, "Coinbase") || eqi(x.account, "Kraken");
  block("Stocks & funds", sales.filter((x) => !isCrypto(x)));
  block("Crypto", sales.filter(isCrypto));
  if (pending.length) {
    lines.push(L("Sales waiting for cost basis", null, 0, true));
    pending.forEach((t) => lines.push(L(`${ymd(t.date)} · ${t.description.slice(0, 40)}`, null, 1, false, `proceeds ${money(t.amount)}`)));
  }
  lines.push(L("Total realized gain", sum(sales.map(gainOf).filter((g) => g != null)), 0, true));
  if (!sales.length && !pending.length) lines.push(L("No stock or crypto sales in this period", null, 1));
  return { title: "Investment gains (sold)", subtitle: periodLabel(month, scope), lines,
    footnote: "Only gains on things you sold. Cost basis is your average cost; your broker's 1099-B is the official number." };
}

const NOT_WAGES = new Set(["interest", "investing", "scholarship", "reimbursement", "misc income", "cc reward"]);

export function taxSummary(s, year) {
  const months = Array.from({ length: 12 }, (_, i) => `${year}-${String(i + 1).padStart(2, "0")}`);
  const txs = inPeriod(s, months);
  const cat = (...names) => sum(txs.filter((t) => names.some((n) => eqi(t.category, n))), (t) => t.amount);
  const lines = [L("Income", null, 0, true)];
  for (const c of s.categories.filter((c) => c.type === T.INCOME && !NOT_WAGES.has(lc(c.name)))) {
    const v = cat(c.name); if (v !== 0) lines.push(L(`Wages — ${c.name}`, v, 1, false, "W-2"));
  }
  lines.push(L("Interest", cat("Interest"), 1, false, "1099-INT"));
  lines.push(L("Dividends", cat("Investing"), 1, false, "1099-DIV"));
  const gains = sum(s.sales.filter((x) => monthYear(monthOfDay(x.date)) === year).map(gainOf).filter((g) => g != null));
  lines.push(L("Capital gains (realized)", gains, 1, false, "1099-B / Form 8949"));
  const sch = cat("Scholarship"); if (sch !== 0) lines.push(L("Scholarships", sch, 1, false, "taxable only above qualified tuition"));
  const misc = cat("Misc Income"); if (misc !== 0) lines.push(L("Other income", misc, 1));
  lines.push(L("Education", null, 0, true));
  lines.push(L("Tuition & fees paid", 0 - cat("Tuition"), 1, false, "1098-T — may qualify for an education credit"));
  lines.push(L("Deductions & credits", null, 0, true));
  lines.push(L("Charitable gifts (tithing, fast offerings)", 0 - cat("Tithing", "Fast Offerings"), 1, false, "deductible only if you itemize"));
  lines.push(L("Not counted", null, 0, true));
  lines.push(L("Reimbursements", cat("Reimbursement"), 1, false, "repayments, not income"));
  lines.push(L("Credit card rewards", cat("CC Reward"), 1, false, "usually treated as a rebate"));
  return { title: "Tax return summary", subtitle: String(year), lines,
    footnote: "A starting point for your return, built from your categories. Use your W-2, 1099 and 1098-T forms for the actual numbers. This isn't tax advice." };
}

export function defaultIncreaseCategories(s) {
  return s.categories.filter((c) => c.type === T.INCOME && !eqi(c.name, "Reimbursement")).map((c) => c.name);
}

export function donations(s, throughMonth, fastPerMonth, increaseCategories, paidCategories = ["Tithing", "Fast Offerings"]) {
  const y = monthYear(throughMonth);
  const months = Array.from({ length: monthNum(throughMonth) }, (_, i) => `${y}-${String(i + 1).padStart(2, "0")}`);
  const inc = new Set(increaseCategories.map(lc)), paidC = new Set(paidCategories.map(lc));
  const rows = months.map((m) => {
    const txs = s.transactions.filter((t) => t.budgetMonth === m);
    const increase = sum(txs.filter((t) => inc.has(lc(t.category))), (t) => t.amount);
    const paid = 0 - sum(txs.filter((t) => paidC.has(lc(t.category))), (t) => t.amount);
    return { month: m, increase, tithe: Math.max(0, increase) * 0.10, fast: fastPerMonth, paid };
  });
  const d = { year: y, months: rows, increase: sum(rows, (r) => r.increase), tithe: sum(rows, (r) => r.tithe), fast: sum(rows, (r) => r.fast), paid: sum(rows, (r) => r.paid) };
  d.owed = d.tithe + d.fast; d.remaining = d.owed - d.paid;
  return d;
}

// ---- CSV ---------------------------------------------------------------

const csvEsc = (v) => { const s = String(v ?? ""); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
const n2 = (v) => (v == null ? "" : v.toFixed(2));

export function reportCsv(r) {
  let out = `${csvEsc(r.title)},${csvEsc(r.subtitle)}\nLine,Amount,Note\n`;
  for (const l of r.lines) out += `${csvEsc("  ".repeat(l.level) + l.label)},${n2(l.amount)},${csvEsc(l.note)}\n`;
  if (r.footnote) out += `\n${csvEsc(r.footnote)}\n`;
  return out;
}

export function transactionsCsv(s, months = null) {
  const set = months ? new Set(months) : null;
  let out = "Date,Budget Month,Account,Description,Amount,Group,Category,Bank Category,Notes,Transaction ID\n";
  for (const t of s.transactions.filter((t) => !set || set.has(t.budgetMonth)).sort((a, b) => a.date - b.date)) {
    out += [ymd(t.date), t.budgetMonth, t.account, t.description, t.amount.toFixed(2), category(s, t.category)?.group || "Review",
      t.category, t.bankCategory, t.notes, t.id].map(csvEsc).join(",") + "\n";
  }
  return out;
}

export function donationsCsv(d) {
  let out = "Month,Increase,Tithing (10%),Fast offering,Paid\n";
  for (const m of d.months) out += `${m.month},${n2(m.increase)},${n2(m.tithe)},${n2(m.fast)},${n2(m.paid)}\n`;
  out += `Total,${n2(d.increase)},${n2(d.tithe)},${n2(d.fast)},${n2(d.paid)}\nOwed,${n2(d.owed)}\nRemaining,${n2(d.remaining)}\n`;
  return out;
}

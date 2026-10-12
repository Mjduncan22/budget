// Port of Analytics.kt — same rules as the Android app so both show identical numbers.
import { lc, eqi, sum, groupBy, addMonths, currentMonth, isMonth, todayEpochDay } from "./util.js";

export const T = { INCOME: "Income", EXPENSE: "Expense", INVESTMENT: "Investment", TRANSFER: "Transfer", REVIEW: "Review", UNCATEGORIZED: "Uncategorized" };

export function category(s, name) { return s.categories.find((c) => eqi(c.name, name)) || null; }

export function needsReview(s, tx) { return eqi(tx.category, T.UNCATEGORIZED) || !category(s, tx.category); }

export function summarize(s, month) {
  const inMonth = s.transactions.filter((t) => t.budgetMonth === month);
  const byCat = groupBy(inMonth, (t) => lc(t.category));
  const summary = (c) => {
    const txs = byCat.get(lc(c.name)) || [];
    const raw = sum(txs, (t) => t.amount);
    return { category: c, spent: c.type === T.INCOME ? raw : 0 - raw, count: txs.length };
  };
  const incomeCats = s.categories.filter((c) => c.type === T.INCOME).map(summary);
  const expenseCats = s.categories.filter((c) => c.type === T.EXPENSE).map(summary);
  const investCats = s.categories.filter((c) => c.type === T.INVESTMENT).map(summary);
  const groups = [...groupBy(expenseCats, (c) => c.category.group)].map(([name, categories]) => ({
    name, categories,
    spent: sum(categories, (c) => c.spent),
    budget: categories.some((c) => c.category.budget != null) ? sum(categories, (c) => c.category.budget) : null,
  }));
  const known = new Set(s.categories.map((c) => lc(c.name)));
  const review = s.transactions.filter((t) => eqi(t.category, T.UNCATEGORIZED) || !known.has(lc(t.category))).length;
  const income = sum(incomeCats, (c) => c.spent), spending = sum(expenseCats, (c) => c.spent);
  return { month, income, spending, invested: sum(investCats, (c) => c.spent), net: income - spending,
    incomeCategories: incomeCats, spendingGroups: groups, needsReview: review };
}

export function monthReview(s, month) {
  const prev = addMonths(month, -1);
  const now = summarize(s, month), before = summarize(s, prev);
  const prevBy = new Map(before.spendingGroups.flatMap((g) => g.categories).map((c) => [c.category.name, c.spent]));
  const changes = now.spendingGroups.flatMap((g) => g.categories)
    .map((c) => [c.category.name, c.spent, prevBy.get(c.category.name) || 0])
    .filter((x) => x[1] !== 0 || x[2] !== 0)
    .sort((a, b) => Math.abs(b[1] - b[2]) - Math.abs(a[1] - a[2]))
    .slice(0, 4);
  const largest = s.transactions.filter((t) => t.budgetMonth === month && t.amount < 0 && category(s, t.category)?.type === T.EXPENSE)
    .sort((a, b) => a.amount - b.amount).slice(0, 3);
  const reviewCount = s.transactions.filter((t) => t.budgetMonth === month && needsReview(s, t)).length;
  return { spending: now.spending, prevSpending: before.spending, changes, largest, reviewCount };
}

export function categoryTrend(s, cat, endMonth, count = 6) {
  const out = [];
  for (let back = count - 1; back >= 0; back--) {
    const m = addMonths(endMonth, -back);
    const raw = sum(s.transactions.filter((t) => t.budgetMonth === m && eqi(t.category, cat.name)), (t) => t.amount);
    out.push([m, cat.type === T.INCOME ? raw : 0 - raw]);
  }
  return out;
}

export function months(s) {
  const all = new Set([...s.transactions.map((t) => t.budgetMonth), currentMonth()]);
  return [...all].filter(isMonth).sort().reverse();
}

export function latestMonth(s) {
  if (!s.transactions.length) return currentMonth();
  let best = s.transactions[0];
  for (const t of s.transactions) if (t.date > best.date) best = t;
  return best.budgetMonth;
}

// ---- rules -------------------------------------------------------------

export function matchRule(rules, description, account) {
  let best = null;
  for (const r of rules) {
    if (!r.keyword || !lc(description).includes(lc(r.keyword))) continue;
    if (r.account && !eqi(r.account, account)) continue;
    if (!best || r.priority > best.priority || (r.priority === best.priority && r.keyword.length > best.keyword.length)) best = r;
  }
  return best;
}

export function suggestKeyword(description) {
  const cleaned = description.replace(/^(TST\*|SQ \*|SP |PAYPAL \*|GOOGLE \*|AMZN MKTP |DD \*)/i, "").trim();
  const words = cleaned.split(/\s+/).filter(Boolean);
  const generic = new Set(["external", "withdrawal", "deposit", "direct", "debit", "pos", "purchase", "ach"]);
  const take = [];
  for (const w of words) {
    if (/\d/.test(w) || w.startsWith("#")) break;
    take.push(w);
    const meaningful = take.filter((x) => !generic.has(x.toLowerCase())).length;
    if (meaningful >= 2 || (meaningful === 1 && w.length >= 8)) break;
  }
  return take.join(" ").replace(/[*\- ]+$/, "") || words[0] || "";
}

export function similar(s, text) {
  const t = (text || "").trim();
  if (!t) return [];
  return s.transactions.filter((x) => lc(x.description).includes(lc(t)));
}

const PREFIXES = /^(TST\* ?|SQ \*|SP |PAYPAL \*|GOOGLE \*|AMZN MKTP |DD \*|POS )+/i;
const GENERIC = new Set(["EXTERNAL", "WITHDRAWAL", "DEPOSIT", "DIRECT", "DEBIT", "POS", "PURCHASE", "ACH", "THE"]);

export function merchantKey(description) {
  const words = description.toUpperCase().replace(PREFIXES, "").split(/[^A-Z0-9&'.-]+/).filter(Boolean);
  const out = [];
  for (const w of words) {
    if (/\d/.test(w)) { if (out.length) break; else continue; }
    if (GENERIC.has(w)) continue;
    out.push(w.replace(/^[.-]+|[.-]+$/g, ""));
    if (out.length === 2) break;
  }
  return out.filter(Boolean).join(" ");
}

export function suggestions(s, tx, max = 3) {
  const key = merchantKey(tx.description);
  const counts = new Map();
  if (key) for (const t of s.transactions) {
    if (t.id === tx.id || needsReview(s, t) || merchantKey(t.description) !== key) continue;
    counts.set(t.category, (counts.get(t.category) || 0) + 1);
  }
  const fromHistory = [...counts].sort((a, b) => b[1] - a[1]).map((x) => x[0]);
  const rule = matchRule(s.rules, tx.description, tx.account);
  const prevMatch = /Review: was (.+?) - /.exec(tx.notes || "");
  const bc = lc(tx.bankCategory);
  const fromBank = bc.includes("dining") || bc.includes("restaurant") ? ["Dining"] : bc.includes("gas") ? ["Gas"] : bc.includes("grocer") ? ["Groceries"] : [];
  const all = [...(prevMatch ? [prevMatch[1]] : []), ...(rule ? [rule.category] : []), ...fromHistory, ...fromBank];
  return [...new Set(all.filter((c) => category(s, c)))].slice(0, max);
}

// ---- subscriptions -----------------------------------------------------

const CADENCES = [
  { label: "Weekly", days: 7, lo: 6, hi: 8 }, { label: "Monthly", days: 30, lo: 26, hi: 35 },
  { label: "Quarterly", days: 91, lo: 84, hi: 98 }, { label: "Yearly", days: 365, lo: 350, hi: 380 },
];
const EVERYDAY = new Set(["gas", "dining", "groceries", "dates", "parking"]);

export function subscriptions(s, today = todayEpochDay()) {
  const types = new Map(s.categories.map((c) => [lc(c.name), c.type]));
  const spending = s.transactions.filter((t) => {
    const ty = types.get(lc(t.category));
    return t.amount < 0 && ty !== T.TRANSFER && ty !== T.INVESTMENT && !EVERYDAY.has(lc(t.category));
  });
  const out = [];
  for (const [key, group] of groupBy(spending, (t) => merchantKey(t.description))) {
    if (!key || group.length < 2) continue;
    const amounts = group.map((t) => -t.amount).sort((a, b) => a - b);
    const median = amounts[Math.floor(amounts.length / 2)];
    const seen = new Set();
    const steady = group.filter((t) => Math.abs(-t.amount - median) <= Math.max(0.5, median * 0.10))
      .sort((a, b) => a.date - b.date)
      .filter((t) => (seen.has(t.date) ? false : (seen.add(t.date), true)));
    if (steady.length < 2) continue;
    const gaps = steady.slice(1).map((t, i) => t.date - steady[i].date);
    const need = Math.max(1, Math.floor(gaps.length * 0.7));
    const cad = CADENCES.find((c) => gaps.filter((g) => g >= c.lo && g <= c.hi).length >= need);
    if (!cad) continue;
    const last = steady[steady.length - 1];
    if (steady.length < 3 && !eqi(last.category, "Subscriptions")) continue;
    if (cad.label === "Weekly" && steady.length < 4) continue;
    const next = last.date + cad.days;
    const overdue = today - next;
    const status = overdue <= Math.max(5, Math.floor(cad.days / 6)) ? "Active" : overdue <= cad.days ? "Late" : "Stopped";
    out.push({
      name: key.split(" ").map((w) => w.charAt(0) + w.slice(1).toLowerCase()).join(" "),
      category: last.category, amount: -last.amount, cadence: cad.label, periodDays: cad.days,
      monthlyCost: -last.amount * 30.44 / cad.days, last: last.date, next,
      charges: [...steady].sort((a, b) => b.date - a.date), status, active: status !== "Stopped",
    });
  }
  const rank = { Active: 0, Late: 1, Stopped: 2 };
  return out.sort((a, b) => (rank[a.status] - rank[b.status]) || (b.monthlyCost - a.monthlyCost));
}

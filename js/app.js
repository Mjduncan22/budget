// Budget web app: same tabs and math as the Android app, for any phone browser.
import * as Auth from "./auth.js";
import { Workbook, newTxId } from "./data.js";
import * as A from "./analytics.js";
import * as F from "./financials.js";
import { quotes as fetchQuotes } from "./prices.js";
import { esc, money, pct, qty, monthLabel, dayLabel, ymd, parseYmd, todayEpochDay, monthOfDay, monthYear, lc, eqi, sum, groupBy } from "./util.js";

export const VERSION = "1.0.1";

// ---------------------------------------------------------------- state

const S = {
  signedIn: null, account: null, offline: false,
  snap: { transactions: [], categories: [], rules: [], holdings: [], accounts: [], balances: [], sales: [] },
  month: null, monthPicked: false, loading: false, saving: false,
  quotes: {}, quotesAt: 0, quotesLoading: false,
  path: null, wb: null, error: null,
};
const UI = {
  fin: { which: 0, ytd: false }, port: { by: 0 }, rec: { focused: null, skipped: new Set() },
  tx: { q: "", month: null, cat: null, review: false }, don: { fast: "20", included: null },
  open: new Set(), subsOpen: new Set(),
};

const store = {
  get(k, d = null) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* full or blocked */ } },
  del(k) { try { localStorage.removeItem(k); } catch { /* blocked */ } },
};
const userKey = (k) => `budget.${S.account?.oid || S.account?.user || "x"}.${k}`;

const $ = (sel, root = document) => root.querySelector(sel);
const main = () => $("#main");

// ---------------------------------------------------------------- toasts

let toastTimer;
function say(msg) {
  const t = $("#toast");
  t.textContent = msg; t.classList.add("show");
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove("show"), 3500);
}

// ---------------------------------------------------------------- data

function setSnap(snap, { keepMonth = false } = {}) {
  S.snap = snap;
  if (!keepMonth && !S.monthPicked) S.month = A.latestMonth(snap);
  if (!S.month) S.month = A.latestMonth(snap);
  if (!S.offline) store.set(userKey("snapshot"), snap);
}

async function refresh() {
  if (S.loading) return;
  if (S.offline) { render(); return; }
  S.loading = true; S.error = null; renderTop();
  try {
    if (!S.path) {
      S.path = await Workbook.find(store.get(userKey("path")));
      if (!S.path) { S.error = "nofile"; S.loading = false; render(); return; }
      store.set(userKey("path"), S.path);
    }
    S.wb = S.wb && S.wb.path === S.path ? S.wb : new Workbook(S.path);
    setSnap(await S.wb.load());
    S.loading = false;
    render();
    refreshPrices(); // always: fund prices come from the workbook, so new sync data shows right away
  } catch (e) {
    S.loading = false;
    if (e instanceof Auth.NotSignedIn) { S.signedIn = false; render(); return; }
    S.error = e.message || "Couldn't load your workbook";
    say(S.error); render();
  }
}

async function refreshPrices() {
  if (!S.snap.holdings.length || S.quotesLoading) return;
  S.quotesLoading = true;
  try {
    const q = await fetchQuotes(S.snap.holdings);
    S.quotes = { ...S.quotes, ...q };
    if (Object.keys(q).length) S.quotesAt = Date.now();
  } catch { say("Couldn't get live prices right now"); }
  S.quotesLoading = false;
  if (["portfolio", "financials", "overview"].includes(route().page)) render();
}

/** Runs a change against Excel (or just in memory in preview), then re-renders. */
async function write(done, fn) {
  S.saving = true; renderTop();
  try {
    const msg = await fn();
    say((msg || done) + (S.offline ? " (preview only — not saved to Excel)" : ""));
  } catch (e) {
    if (e instanceof Auth.NotSignedIn) { S.signedIn = false; }
    say(e.message || "Couldn't save to Excel");
  }
  S.saving = false;
  render();
}

async function recategorize(tx, cat, matchText, saveRule) {
  await write("Saved", async () => {
    const text = (matchText || "").trim();
    let targets = text ? A.similar(S.snap, text).filter((t) => !eqi(t.category, cat)) : [tx];
    if (text && !lc(tx.description).includes(lc(text))) targets.push(tx);
    targets = [...new Map(targets.map((t) => [t.id, t])).values()];
    let rows = {};
    if (!S.offline && targets.length) rows = await S.wb.setCategories(targets, cat);
    const ids = new Set(targets.map((t) => t.id));
    let snap = { ...S.snap, transactions: S.snap.transactions.map((t) => (ids.has(t.id) ? { ...t, category: cat, sheetRow: rows[t.id] || t.sheetRow } : t)) };
    let ruleNote = "";
    if (saveRule && text && !snap.rules.some((r) => eqi(r.keyword, text) && eqi(r.category, cat))) {
      const rule = { keyword: text, category: cat, account: "", priority: text.length + 100 };
      if (!S.offline) await S.wb.addRule(rule);
      snap = { ...snap, rules: [rule, ...snap.rules] };
      ruleNote = " · rule saved";
    }
    setSnap(snap, { keepMonth: true });
    return targets.length === 1 ? `Saved${ruleNote}` : `Moved ${targets.length} transactions to ${cat}${ruleNote}`;
  });
}

async function setBudget(cat, amount) {
  await write("Budget updated", async () => {
    if (!S.offline) await S.wb.setBudget(cat, amount);
    setSnap({ ...S.snap, categories: S.snap.categories.map((c) => (c.name === cat.name ? { ...c, budget: amount } : c)) }, { keepMonth: true });
  });
}

async function addTransaction(tx) {
  await write("Transaction added", async () => {
    if (!S.offline) await S.wb.addTransaction(tx);
    setSnap({ ...S.snap, transactions: [...S.snap.transactions, { ...tx, source: "App", bankCategory: "", sheetRow: 0 }] }, { keepMonth: true });
  });
}

function download(name, csv) {
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = name.endsWith(".csv") ? name : name + ".csv";
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
  say("Downloaded " + a.download);
}

// ---------------------------------------------------------------- routing

const TABS = [
  ["overview", "Overview", "M3 13h8V3H3v10zm0 8h8v-6H3v6zm10 0h8V11h-8v10zm0-18v6h8V3h-8z"],
  ["financials", "Financials", "M4 20V10h3v10H4zm6.5 0V4h3v16h-3zM17 20v-7h3v7h-3z"],
  ["bankrec", "Bank Rec", "M9 16.2 4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4z"],
  ["portfolio", "Portfolio", "M3.5 18.5 9.5 12.5l4 4L22 7.9l-1.4-1.4-7.1 8-4-4L2 17z"],
  ["more", "More", "M6 10c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm12 0c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm-6 0c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2z"],
];
const SUBPAGES = { tx: ["All transactions", "more"], subs: ["Subscriptions", "more"], donations: ["Church donations", "more"], cat: [null, "overview"], settings: ["Settings", "more"] };

function route() {
  const h = decodeURIComponent(location.hash.slice(1) || "overview");
  const [page, ...rest] = h.split("/");
  return { page: TABS.some((t) => t[0] === page) || SUBPAGES[page] ? page : "overview", arg: rest.join("/") };
}
function go(hash) { if (location.hash.slice(1) !== hash) location.hash = hash; else render(); }

// ---------------------------------------------------------------- shared bits

const icon = (d, cls = "") => `<svg class="ic ${cls}" viewBox="0 0 24 24" aria-hidden="true"><path d="${d}"/></svg>`;
const I = {
  back: "M20 11H7.8l5.6-5.6L12 4l-8 8 8 8 1.4-1.4L7.8 13H20z",
  left: "M15.4 7.4 14 6l-6 6 6 6 1.4-1.4L10.8 12z", right: "M8.6 16.6 10 18l6-6-6-6-1.4 1.4 4.6 4.6z",
  down: "M7 10l5 5 5-5z", chev: "M10 6 8.6 7.4 13.2 12l-4.6 4.6L10 18l6-6z",
  refresh: "M17.65 6.35A7.96 7.96 0 0 0 12 4a8 8 0 1 0 7.73 10h-2.08A6 6 0 1 1 12 6c1.66 0 3.14.69 4.22 1.78L13 11h7V4l-2.35 2.35z",
  warn: "M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z", check: "M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm-2 15-5-5 1.4-1.4L10 14.2l7.6-7.6L19 8l-9 9z",
  plus: "M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z", dl: "M5 20h14v-2H5v2zM19 9h-4V3H9v6H5l7 7 7-7z",
  search: "M15.5 14h-.8l-.3-.3A6.5 6.5 0 1 0 14 15.5l.3.3v.8l5 5 1.5-1.5-5-5zm-6 0a4.5 4.5 0 1 1 0-9 4.5 4.5 0 0 1 0 9z",
  close: "M19 6.4 17.6 5 12 10.6 6.4 5 5 6.4 10.6 12 5 17.6 6.4 19 12 13.4 17.6 19 19 17.6 13.4 12z",
};

const label = (t) => `<div class="label">${esc(t)}</div>`;
const card = (inner, cls = "") => `<section class="card ${cls}">${inner}</section>`;

/** kind: "budget" (amber past 85%, red past 100%), "share" (share of total) or "progress" (always teal). */
function bar(progress, kind = "budget") {
  if (progress == null) return "";
  if (kind === true) kind = "share";
  const p = Math.max(0, Math.min(1, progress));
  const cls = kind === "share" ? "share" : kind === "progress" ? "" : progress > 1 ? "over" : progress > 0.85 ? "warn" : "";
  return `<div class="bar"><div class="fill ${cls}" style="width:${(p * 100).toFixed(1)}%"></div></div>`;
}

function monthSwitcher() {
  const months = A.months(S.snap);
  const i = months.indexOf(S.month);
  const opts = months.map((m) => `<option value="${m}" ${m === S.month ? "selected" : ""}>${monthLabel(m)}</option>`).join("");
  return `<div class="monthsw">
    <button class="iconbtn" data-act="month" data-m="${months[i + 1] || ""}" ${i + 1 < months.length ? "" : "disabled"} aria-label="Previous month">${icon(I.left)}</button>
    <label class="monthpick"><span>${monthLabel(S.month)}</span>${icon(I.down, "sm")}
      <select data-change="month" aria-label="Choose month">${opts}</select></label>
    <button class="iconbtn" data-act="month" data-m="${months[i - 1] || ""}" ${i > 0 ? "" : "disabled"} aria-label="Next month">${icon(I.right)}</button>
  </div>`;
}

function txRow(tx, showCategory = true) {
  const review = A.needsReview(S.snap, tx);
  const sub = `${dayLabel(tx.date)} · ${esc(tx.account)}${showCategory ? " · " + esc(tx.category) : ""}`;
  return `<button class="row tx" data-act="edit" data-id="${esc(tx.id)}">
    <div class="grow"><div class="t1 ell">${esc(tx.description)}</div><div class="t2 ell ${review ? "warn" : ""}">${sub}</div></div>
    <div class="amt ${tx.amount > 0 ? "pos" : ""}">${money(tx.amount)}</div></button>`;
}

const amountRow = (l, v, act = "") => `<button class="row" ${act}><span class="grow">${esc(l)}</span><span class="amt">${v}</span></button>`;

function segmented(name, options, value) {
  return `<div class="seg">${options.map((o, i) => `<button data-act="seg" data-seg="${name}" data-v="${i}" class="${i === value ? "on" : ""}">${esc(o)}</button>`).join("")}</div>`;
}

// ---------------------------------------------------------------- pages

function pageOverview() {
  const s = S.snap, sum0 = A.summarize(s, S.month), r = A.monthReview(s, S.month);
  const budgetTotal = sum(sum0.spendingGroups.map((g) => g.budget).filter((b) => b != null));
  let h = monthSwitcher();
  h += `<div class="hero"><div class="cap">Spent</div><div class="big">${money(sum0.spending)}</div>`;
  if (budgetTotal > 0) {
    const left = budgetTotal - sum0.spending;
    h += bar(sum0.spending / budgetTotal) + `<div class="t2 ${left < 0 ? "neg" : ""}">${left >= 0 ? `${money(left, true)} left of ${money(budgetTotal, true)} budget` : `${money(-left, true)} over ${money(budgetTotal, true)} budget`}</div>`;
  }
  h += `<div class="tiles"><div class="tile"><div class="cap">Income</div><div class="v">${money(sum0.income, true)}</div></div>
    <div class="tile"><div class="cap">Net</div><div class="v ${sum0.net >= 0 ? "pos" : "neg"}">${money(sum0.net, true)}</div></div>
    <div class="tile"><div class="cap">Invested</div><div class="v">${money(sum0.invested, true)}</div></div></div></div>`;
  if (sum0.needsReview > 0) {
    h += `<button class="banner warnbg" data-go="bankrec">${icon(I.warn, "warn")}<span class="grow">${sum0.needsReview} transaction${sum0.needsReview === 1 ? "" : "s"} need a category</span><b class="link">Review</b></button>`;
  }
  // month in review
  const diff = r.spending - r.prevSpending;
  let mr = `<p class="lead ${diff > 0 && r.prevSpending > 0 ? "neg" : ""}">${r.prevSpending === 0 ? `Spent ${money(r.spending)} this month.` : diff > 0 ? `Spent ${money(diff)} more than last month (${money(r.prevSpending, true)}).` : diff < 0 ? `Spent ${money(-diff)} less than last month (${money(r.prevSpending, true)}).` : "Same spending as last month."}</p>`;
  if (r.changes.length) mr += `<div class="cap">Biggest changes</div>` + r.changes.map(([c, now, before]) => `<div class="kv"><span>${esc(c)}</span><span class="${now > before ? "neg" : "pos"} sm">${money(before, true)} → ${money(now, true)}</span></div>`).join("");
  if (r.largest.length) mr += `<div class="cap">Largest purchases</div>` + r.largest.map((t) => `<div class="kv"><span class="ell">${esc(t.description)}</span><span class="sm">${money(-t.amount)}</span></div>`).join("");
  if (r.reviewCount > 0) mr += `<button class="textbtn" data-go="bankrec">${r.reviewCount} uncategorized this month — open Bank Rec</button>`;
  h += label("Month in review") + card(`<div class="pad">${mr}</div>`);
  h += label("Spending by group");
  for (const g of sum0.spendingGroups.filter((g) => g.spent !== 0 || g.budget != null)) {
    const open = UI.open.has(g.name);
    let gi = `<button class="grouphead" data-act="toggle" data-k="${esc(g.name)}"><div class="kv"><b>${esc(g.name)}</b><span><b>${money(g.spent)}</b>${icon(I.down, open ? "flip sm" : "sm")}</span></div>`;
    gi += g.budget > 0 ? bar(g.spent / g.budget) + `<div class="t2">of ${money(g.budget, true)} budget</div>` : bar(sum0.spending > 0 ? g.spent / sum0.spending : 0, true);
    gi += `</button>`;
    if (open) gi += g.categories.map((c) => `<button class="row sub" data-go="cat/${esc(c.category.name)}"><div class="grow"><div class="kv"><span>${esc(c.category.name)}</span><span class="sm">${c.category.budget != null ? `${money(c.spent, true)} / ${money(c.category.budget, true)}` : money(c.spent)}</span></div>${c.category.budget > 0 ? bar(c.spent / c.category.budget) : ""}</div></button>`).join("");
    h += card(gi, "tight");
  }
  h += label("Income");
  const inc = sum0.incomeCategories.filter((c) => c.spent !== 0);
  h += card(inc.length ? inc.map((c) => amountRow(c.category.name, money(c.spent), `data-go="cat/${esc(c.category.name)}"`)).join('<hr>') : `<div class="pad t2">No income recorded this month</div>`);
  return h;
}

function pageCategory(name) {
  const s = S.snap, cat = A.category(s, name);
  if (!cat) return `<div class="pad">Category "${esc(name)}" isn't in your Categories table.</div>`;
  const trend = A.categoryTrend(s, cat, S.month);
  const spent = trend[trend.length - 1]?.[1] || 0;
  const txs = s.transactions.filter((t) => t.budgetMonth === S.month && eqi(t.category, cat.name)).sort((a, b) => b.date - a.date);
  const isIncome = cat.type === A.T.INCOME;
  let h = `<div class="hero"><div class="cap">${esc(cat.group)} · ${monthLabel(S.month)}</div><div class="big">${money(spent)}</div>`;
  if (cat.budget != null && !isIncome) {
    h += bar(spent / cat.budget) + `<div class="t2 ${spent > cat.budget ? "neg" : ""}">${spent <= cat.budget ? `${money(cat.budget - spent)} left of ${money(cat.budget, true)}` : `${money(spent - cat.budget)} over ${money(cat.budget, true)}`}</div>`;
  }
  if (!isIncome && cat.type !== A.T.TRANSFER) h += `<button class="outline" data-act="budget" data-c="${esc(cat.name)}">${cat.budget == null ? "Set monthly budget" : "Change budget"}</button>`;
  h += `</div>` + label("Last 6 months");
  const max = Math.max(1, cat.budget || 0, ...trend.map((x) => x[1]));
  const bars = trend.map(([m, v], i) => `<div class="tb"><span class="cap">${v > 0 ? money(v, true) : ""}</span><div class="tbwrap"><div class="tbfill ${i === trend.length - 1 ? "now" : ""}" style="height:${Math.max(2, (Math.max(0, v) / max) * 100)}%"></div></div><span class="cap">${monthLabel(m, true)}</span></div>`).join("");
  const avg = sum(trend, (x) => x[1]) / trend.length;
  h += card(`<div class="trend">${bars}</div><div class="pad t2">Average ${money(avg)} / month${cat.budget != null ? ` · budget ${money(cat.budget, true)}` : ""}</div>`);
  h += label(`${txs.length} transaction${txs.length === 1 ? "" : "s"} in ${monthLabel(S.month, true)}`);
  h += card(txs.map((t) => txRow(t, false)).join("<hr>") || `<div class="pad t2">None</div>`);
  return h;
}

const REPORTS = ["Income statement", "Balance sheet", "Cash flow", "Investments", "Tax return"];

function currentReport() {
  const { which, ytd } = UI.fin, scope = ytd ? "YTD" : "MONTH", m = S.month, s = S.snap;
  return [F.incomeStatement(s, m, scope), null, F.cashFlow(s, m, scope), F.realizedGains(s, m, scope), F.taxSummary(s, monthYear(m))][which] || F.balanceSheet(s, m, S.quotes);
}

function pageFinancials() {
  const { which, ytd } = UI.fin;
  const r = which === 1 ? F.balanceSheet(S.snap, S.month, S.quotes) : currentReport();
  let h = `<div class="chips scroll">${REPORTS.map((n, i) => `<button class="chip ${i === which ? "on" : ""}" data-act="report" data-i="${i}">${n}</button>`).join("")}</div>`;
  h += monthSwitcher();
  if ([0, 2, 3].includes(which)) h += `<div class="padx">${segmented("ytd", ["Month", "Year to date"], ytd ? 1 : 0)}</div>`;
  h += `<div class="rephead"><div class="grow"><h2>${esc(r.title)}</h2><div class="t2">${esc(r.subtitle)}</div></div><button class="outline sm" data-act="csv-report">${icon(I.dl, "sm")} CSV</button></div>`;
  h += card(r.lines.map((l, i) => {
    const heading = l.amount == null && l.bold;
    const sep = l.amount == null && l.level === 0 && i > 0 ? "<hr>" : "";
    return `${sep}<div class="rl ${heading ? "head" : ""} ${l.bold && !heading ? "b" : ""}" style="padding-left:${20 + 16 * l.level}px"><div class="kv"><span>${esc(l.label)}</span>${l.amount != null ? `<span class="${l.bold && l.amount < 0 ? "neg" : ""}">${money(l.amount)}</span>` : ""}</div>${l.note ? `<div class="t2">${esc(l.note)}</div>` : ""}</div>`;
  }).join(""));
  if (r.footnote) h += `<p class="foot">${esc(r.footnote)}</p>`;
  if (which === 0 || which === 2) h += `<div class="padx"><button class="outline" data-act="csv-tx">Download the transactions behind this report (CSV)</button></div>`;
  return h;
}

function pageBankRec() {
  const s = S.snap, m = S.month;
  const inMonth = s.transactions.filter((t) => t.budgetMonth === m);
  const todo = inMonth.filter((t) => A.needsReview(s, t)).sort((a, b) => a.date - b.date);
  const others = [...groupBy(s.transactions.filter((t) => A.needsReview(s, t) && t.budgetMonth !== m), (t) => t.budgetMonth)].sort((a, b) => (a[0] < b[0] ? 1 : -1));
  const cur = todo.find((t) => t.id === UI.rec.focused) || todo.find((t) => !UI.rec.skipped.has(t.id)) || todo[0];
  const done = inMonth.length - todo.length;
  let h = monthSwitcher();
  h += `<div class="hero"><div class="h1 ${todo.length ? "" : "pos"}">${todo.length ? `${todo.length} left to categorize` : "Reconciled"}</div>${bar(inMonth.length ? done / inMonth.length : 1, "progress")}
    <div class="t2">${done} of ${inMonth.length} transactions in ${monthLabel(m)} have a category · ${s.rules.length} saved rules</div></div>`;
  if (others.length) h += `<div class="chips scroll">${others.map(([mo, l]) => `<button class="chip" data-act="month" data-m="${mo}">${monthLabel(mo, true)} '${mo.slice(2, 4)} · ${l.length} left</button>`).join("")}</div>`;
  if (!cur) {
    h += `<div class="empty">${icon(I.check, "pos xl")}<p>${inMonth.length ? `Every transaction in ${monthLabel(m)} has a category.` : `No transactions in ${monthLabel(m)} yet.`}</p></div>`;
  } else {
    h += label("Categorize") + card(assignPanel(cur, todo.length > 1), "assign");
    const rest = todo.filter((t) => t.id !== cur.id);
    if (rest.length) h += label(`Still to do (${rest.length})`) + card(rest.map((t) => txRow(t, false).replace('data-act="edit"', 'data-act="focus"')).join("<hr>"));
  }
  return h;
}

/** Category picker for one transaction with "Just this one" / "All similar" (+ rule). */
function assignPanel(tx, canSkip = false) {
  const s = S.snap, sugg = A.suggestions(s, tx);
  const p = (assignPanel.state && assignPanel.state.id === tx.id) ? assignPanel.state
    : (assignPanel.state = { id: tx.id, selected: A.needsReview(s, tx) ? sugg[0] || "" : tx.category, all: false, match: A.suggestKeyword(tx.description), rule: true });
  const cats = s.categories.filter((c) => c.type !== A.T.REVIEW);
  const review = /Review: was (.+?) - (.+)/.exec(tx.notes || "");
  let h = `<div class="pad assignbody"><div class="t1 strong">${esc(tx.description)}</div>
    <div class="t2">${money(tx.amount)} · ${esc(tx.account)} · ${dayLabel(tx.date)}${tx.bankCategory ? " · bank says " + esc(tx.bankCategory) : ""}${!A.needsReview(s, tx) ? " · now " + esc(tx.category) : ""}</div>`;
  if (review) h += `<div class="t2 warn">Was ${esc(review[1])} — flagged because ${esc(review[2].split(" (")[0])}</div>`;
  if (sugg.length) h += `<div class="cap">Suggested</div><div class="chips">${sugg.map((c) => `<button class="chip ${p.selected === c ? "on" : ""}" data-act="pick" data-c="${esc(c)}">${esc(c)}</button>`).join("")}</div>`;
  for (const [g, list] of groupBy(cats, (c) => c.group)) {
    h += `<div class="cap">${esc(g)}</div><div class="chips">${list.map((c) => `<button class="chip ${eqi(p.selected, c.name) ? "on" : ""}" data-act="pick" data-c="${esc(c.name)}">${esc(c.name)}</button>`).join("")}</div>`;
  }
  h += `<hr><div class="cap">Apply to</div>${segmented("all", ["Just this one", "All similar"], p.all ? 1 : 0)}`;
  if (p.all) {
    h += `<label class="field"><span>Description contains</span><input data-input="match" value="${esc(p.match)}" autocomplete="off"></label><div data-region="matchinfo">${matchInfo(p)}</div>
      <label class="switchrow"><div class="grow"><div>Also for future transactions</div><div class="t2">Saves a rule in the Rules tab</div></div><input type="checkbox" data-change="rule" ${p.rule ? "checked" : ""}></label>`;
  }
  const matches = p.all ? A.similar(s, p.match) : [];
  const n = Math.max(1, matches.filter((t) => !eqi(t.category, p.selected)).length);
  const ok = A.category(s, p.selected) && (!p.all || p.match.trim());
  h += `<div class="actions">${canSkip ? `<button class="outline" data-act="skip" data-id="${esc(tx.id)}">Skip</button>` : ""}<button class="primary grow" data-act="save" data-id="${esc(tx.id)}" ${ok ? "" : "disabled"}>${p.all && matches.length > 1 ? `Save ${n}` : "Save"}</button></div></div>`;
  return h;
}

function matchInfo(p) {
  const s = S.snap, matches = A.similar(s, p.match);
  const already = matches.filter((t) => eqi(t.category, p.selected)).length;
  const unassigned = matches.filter((t) => A.needsReview(s, t)).length;
  const moving = matches.length - already;
  const txt = !p.match.trim() ? "Type part of the description to match"
    : !matches.length ? `No transactions contain "${esc(p.match)}"`
    : `${moving} of ${matches.length} matching transaction${matches.length === 1 ? "" : "s"} will move${unassigned ? ` (${unassigned} uncategorized)` : ""}${moving - unassigned > 0 ? ` · ${moving - unassigned} change from another category` : ""}`;
  const descs = [...new Set(matches.map((t) => t.description))];
  return `<div class="t2 ${moving - unassigned > 0 ? "warn" : ""}">${txt}</div>${descs.length ? `<div class="t2 ell2">${esc(descs.slice(0, 3).join(" · "))}${descs.length > 3 ? " …" : ""}</div>` : ""}`;
}

function pagePortfolio() {
  const pos = S.snap.holdings.map((h) => { const q = S.quotes[h.symbol]; return { h, q, value: q ? q.price * h.quantity : null }; });
  const val = (p) => (p.value ?? p.h.costBasis);
  const total = sum(pos, val), cost = sum(pos, (p) => p.h.costBasis);
  const change = sum(pos, (p) => (p.value != null && p.q.change != null ? p.value - p.value / (1 + p.q.change / 100) : 0));
  const missing = pos.filter((p) => !p.q).length;
  const gain = total - cost;
  const by = UI.port.by;
  const lines = [...groupBy(pos, (p) => (by === 0 ? p.h.symbol : p.h.heldIn || "Unassigned"))].map(([k, ps]) => ({
    title: k, sub: by === 0 ? ps[0].h.assetType : `${ps.length} holding${ps.length === 1 ? "" : "s"}`, ps,
    value: sum(ps, val), cost: sum(ps, (p) => p.h.costBasis), quantity: sum(ps, (p) => p.h.quantity), q: ps.find((p) => p.q)?.q,
  })).sort((a, b) => b.value - a.value);
  const upd = S.quotesAt ? `Prices updated ${new Date(S.quotesAt).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}` : S.quotesLoading ? "Getting live prices…" : "No live prices yet — tap refresh";
  let h = `<div class="hero"><div class="cap">Portfolio value</div><div class="big">${money(total)}</div>
    <div class="t2"><span class="${gain >= 0 ? "pos" : "neg"}">${money(gain)} (${pct(cost > 0 ? gain / cost * 100 : 0)}) all time</span>${change ? ` · <span class="${change >= 0 ? "pos" : "neg"}">${money(change)} today</span>` : ""}</div>
    <div class="t2">${upd}${missing ? ` · ${missing} without a price (shown at cost)` : ""} <button class="textbtn inline" data-act="prices">Refresh</button></div>
    ${segmented("port", ["By asset", "By wallet"], by)}</div>`;
  if (!pos.length) h += `<p class="pad t2">No holdings yet. Add rows to the Holdings tab in Budget Data.xlsx.</p>`;
  h += label(by === 0 ? "Holdings" : "Where it's held");
  h += card(lines.map((l) => {
    const g = l.value - l.cost;
    return `<div class="row col"><div class="kv"><div><div class="t1 strong">${esc(l.title)}</div><div class="t2">${by === 0 ? `${qty(l.quantity)}${l.q ? " @ " + money(l.q.price) : " · no price"}` : esc(l.sub)}</div></div>
      <div class="r"><div class="t1">${money(l.value)}</div><div class="t2 ${g >= 0 ? "pos" : "neg"}">${l.cost > 0 ? `${money(g, true)} (${pct(g / l.cost * 100)})` : ""}</div></div></div>${bar(total > 0 ? l.value / total : 0, true)}</div>`;
  }).join("<hr>"));
  h += `<p class="foot center">Quantities come from the Holdings tab, kept current by the sync. Fund prices are Fidelity's from the last sync.</p>`;
  return h;
}

function pageMore() {
  const subs = A.subscriptions(S.snap).filter((x) => x.active);
  const item = (go, title, sub, act = "") => `<button class="row" ${go ? `data-go="${go}"` : act}><div class="grow"><div class="t1">${esc(title)}</div><div class="t2">${esc(sub)}</div></div>${icon(I.chev, "dim")}</button>`;
  let h = label("Tools") + card([
    item("tx", "All transactions", `${S.snap.transactions.length} transactions · search & edit`),
    item("subs", "Subscriptions", `${subs.length} active · ${money(sum(subs, (x) => x.monthlyCost))}/month`),
    item("donations", "Church donations", "Tithing + fast offering vs. paid this year"),
  ].join("<hr>"));
  h += label("Data") + card([
    item("", "Download all transactions", "CSV file", 'data-act="csv-all"'),
    item("", "Reload from Excel", S.offline ? "Preview mode — sample data" : S.path || "Budget Data.xlsx on OneDrive", 'data-act="reload"'),
    item("settings", "Workbook & app settings", `Version ${VERSION}`),
    item("", S.offline ? "Exit preview" : "Sign out", S.offline ? "Back to the sign-in screen" : (S.account?.user || ""), 'data-act="signout"'),
  ].join("<hr>"));
  return h;
}

function pageSettings() {
  return label("Workbook") + card(`<div class="pad"><label class="field"><span>OneDrive path to your Budget Data.xlsx</span><input id="pathInput" value="${esc(S.path || "")}" placeholder="Documents/Finances/Budget/Budget Data.xlsx"></label>
    <p class="t2">Each person signs in with their own Microsoft account and sees only their own workbook. Lena's lives at <b>Lena Budget/Budget Data.xlsx</b> in her OneDrive.</p>
    <div class="actions"><button class="outline" data-act="findpath">Find it for me</button><button class="primary grow" data-act="savepath">Use this path</button></div></div>`)
    + label("Install on your phone") + card(`<div class="pad t2"><p><b>iPhone:</b> open this page in Safari, tap the Share button, then <b>Add to Home Screen</b>.</p><p><b>Android:</b> in Chrome tap ⋮ then <b>Add to Home screen</b> (or <b>Install app</b>).</p><p>Updates arrive automatically the next time you open it.</p></div>`)
    + label("About") + card(`<div class="pad t2">Budget web app ${VERSION}. Your data stays in your own OneDrive; this page only talks to Microsoft (sign-in and your workbook) and CoinGecko (crypto prices).</div>`);
}

function pageTransactions() {
  return `<div class="padx searchwrap">${icon(I.search, "dim")}<input type="search" data-input="txq" value="${esc(UI.tx.q)}" placeholder="Search description, account, category"></div>
    <div class="chips scroll padx">
      <button class="chip ${UI.tx.review ? "on" : ""}" data-act="txreview">Needs review</button>
      ${UI.tx.month ? `<button class="chip on" data-act="txmonth" data-m="">${monthLabel(UI.tx.month)} ✕</button>` : `<button class="chip" data-act="txmonth" data-m="${S.month}">This month only</button>`}
      ${UI.tx.cat ? `<button class="chip on" data-act="txcat">${esc(UI.tx.cat)} ✕</button>` : ""}
    </div><div data-region="txlist">${txList()}</div>
    <button class="fab" data-act="add" aria-label="Add transaction">${icon(I.plus)}</button>`;
}

function txList() {
  const s = S.snap, q = lc(UI.tx.q.trim());
  const list = s.transactions.filter((t) => (!UI.tx.month || t.budgetMonth === UI.tx.month) && (!UI.tx.cat || eqi(t.category, UI.tx.cat))
    && (!UI.tx.review || A.needsReview(s, t))
    && (!q || lc(t.description).includes(q) || lc(t.account).includes(q) || lc(t.category).includes(q) || lc(t.notes).includes(q)))
    .sort((a, b) => b.date - a.date || a.description.localeCompare(b.description));
  const shown = list.slice(0, 300);
  return `<div class="t2 padx">${list.length} transaction${list.length === 1 ? "" : "s"} · net ${money(sum(list, (t) => t.amount))}${list.length > shown.length ? " · showing newest 300 (search to narrow)" : ""}</div>`
    + (list.length ? card(shown.map((t) => txRow(t, !UI.tx.cat)).join("<hr>")) : `<p class="pad t2">${UI.tx.review ? "Nothing to review — every transaction has a category." : "No transactions match."}</p>`);
}

function pageSubs() {
  const subs = A.subscriptions(S.snap), cur = subs.filter((x) => x.active), stopped = subs.filter((x) => !x.active);
  const monthly = sum(cur, (x) => x.monthlyCost);
  const row = (x) => {
    const open = UI.subsOpen.has(x.name);
    const [st, cls] = x.status === "Active" ? [`Next ~${dayLabel(x.next)}`, "pos"] : x.status === "Late" ? [`Expected ${dayLabel(x.next)} — not charged yet`, "warn"] : [`Last charged ${dayLabel(x.last)}`, "dim"];
    return `<button class="row col" data-act="subtoggle" data-k="${esc(x.name)}"><div class="kv"><div><div class="t1 strong">${esc(x.name)}</div><div class="t2"><span class="dot ${cls}"></span>${st} · ${x.cadence.toLowerCase()} · ${esc(x.category)}</div></div>
      <div class="r"><div class="t1">${money(x.amount)}</div>${x.cadence !== "Monthly" ? `<div class="t2">${money(x.monthlyCost)}/mo</div>` : ""}</div></div>
      ${open ? `<div class="charges">${x.charges.map((c) => `<div class="kv t2"><span>${dayLabel(c.date)} · ${esc(c.account)}</span><span>${money(-c.amount)}</span></div>`).join("")}</div>` : ""}</button>`;
  };
  let h = `<div class="hero"><div class="cap">Recurring each month</div><div class="big">${money(monthly)}</div><div class="t2">${money(monthly * 12, true)} a year · ${cur.length} active</div></div>`;
  if (!subs.length) h += `<p class="pad t2">No recurring charges found yet. They show up once the same merchant charges a steady amount on a regular schedule.</p>`;
  if (cur.length) h += label("Active") + card(cur.map(row).join("<hr>"));
  if (stopped.length) h += label("Stopped — no recent charge") + card(stopped.map(row).join("<hr>"));
  h += `<p class="foot center">Found automatically from your transactions. Tag a charge as Subscriptions in Bank Rec to help it get picked up sooner.</p>`;
  return h;
}

function pageDonations() {
  const s = S.snap;
  const incomeCats = s.categories.filter((c) => c.type === A.T.INCOME).map((c) => c.name);
  if (!UI.don.included) UI.don.included = store.get(userKey("donInc"), null) || F.defaultIncreaseCategories(s);
  const fast = parseFloat(String(UI.don.fast).replace("$", "")) || 0;
  const d = F.donations(s, S.month, fast, UI.don.included);
  const ahead = d.remaining <= 0;
  const line = (l, v, b = false) => `<div class="kv pad1 ${b ? "strong" : ""}"><span>${esc(l)}</span><span>${money(v)}</span></div>`;
  let h = monthSwitcher();
  h += `<div class="hero"><div class="cap">Jan – ${monthLabel(S.month, true)} ${d.year}</div><div class="big ${ahead ? "pos" : ""}">${ahead ? `${money(-d.remaining)} ahead` : `${money(d.remaining)} to go`}</div>${bar(d.owed <= 0 ? 1 : d.paid / d.owed, "progress")}<div class="t2">Paid ${money(d.paid)} of ${money(d.owed)} owed so far this year</div></div>`;
  h += label("How it's worked out") + card(line("Increase (income counted below)", d.increase) + line("Tithing — 10% of each month's increase", d.tithe)
    + line(`Fast offering — ${money(fast)} × ${d.months.length} months`, d.fast) + "<hr>" + line("Total owed", d.owed, true)
    + line("Already paid (Tithing + Fast Offerings)", -d.paid) + "<hr>" + line(d.remaining > 0 ? "Still to pay" : "Paid ahead", Math.abs(d.remaining), true));
  h += label("Settings") + `<div class="padx"><label class="field"><span>Fast offering per month ($)</span><input inputmode="decimal" data-input="fast" value="${esc(UI.don.fast)}"></label>
    <div class="t2">Count as increase:</div><div class="chips">${incomeCats.map((c) => `<button class="chip ${UI.don.included.includes(c) ? "on" : ""}" data-act="doninc" data-c="${esc(c)}">${esc(c)}</button>`).join("")}</div>
    <p class="t2">Reimbursements are left out by default because they're money paid back to you, not new income.</p></div>`;
  h += label("By month") + card(`<div class="grid4 cap"><span>Month</span><span>Increase</span><span>Owed</span><span>Paid</span></div>` + d.months.map((r) => `<div class="grid4"><span>${monthLabel(r.month, true)}</span><span class="sm">${money(r.increase, true)}</span><span class="sm">${money(r.tithe + r.fast, true)}</span><span class="sm ${r.paid > 0 ? "pos" : "dim"}">${money(r.paid, true)}</span></div>`).join(""));
  h += `<div class="padx"><button class="outline" data-act="csv-don">Download CSV</button></div>`;
  return h;
}

// ---------------------------------------------------------------- sheets & dialogs

function openSheet(html, onBind) {
  const sh = $("#sheet");
  sh.innerHTML = `<div class="scrim" data-act="closesheet"></div><div class="panel" role="dialog" aria-modal="true"><div class="grab"></div>${html}</div>`;
  sh.classList.add("open");
  document.body.classList.add("noscroll");
  onBind?.(sh);
}
function closeSheet() { const sh = $("#sheet"); sh.classList.remove("open"); sh.innerHTML = ""; document.body.classList.remove("noscroll"); assignPanel.state = null; }

function editTx(id) {
  const tx = S.snap.transactions.find((t) => t.id === id);
  if (!tx) return;
  assignPanel.state = null;
  openSheet(`<div data-region="sheetassign" data-id="${esc(id)}">${assignPanel(tx)}</div>`);
}

function budgetDialog(cat) {
  openSheet(`<div class="pad"><h2>${esc(cat.name)} budget</h2><p class="t2">Monthly amount. Saved to the Categories tab in Excel.</p>
    <label class="field"><span>Amount ($)</span><input id="budgetIn" inputmode="decimal" value="${cat.budget != null ? cat.budget.toFixed(2) : ""}"></label>
    <div class="actions"><button class="outline" data-act="closesheet">Cancel</button><button class="primary grow" data-act="savebudget" data-c="${esc(cat.name)}">Save</button></div></div>`);
}

function addDialog() {
  const s = S.snap;
  const accounts = [...new Set(["Cash", ...s.transactions.map((t) => t.account)])];
  const cats = s.categories.filter((c) => c.type !== A.T.REVIEW);
  const def = cats.find((c) => c.name === "Dining")?.name || cats[0]?.name || "";
  openSheet(`<div class="pad"><h2>Add transaction</h2>
    ${segmented("addout", ["Money out", "Money in"], 0)}
    <label class="field"><span>Description</span><input id="addDesc"></label>
    <label class="field"><span>Amount ($)</span><input id="addAmt" inputmode="decimal"></label>
    <label class="field"><span>Date</span><input id="addDate" type="date" value="${ymd(todayEpochDay())}"></label>
    <label class="field"><span>Account</span><select id="addAcct">${accounts.map((a) => `<option>${esc(a)}</option>`).join("")}</select></label>
    <label class="field"><span>Category</span><select id="addCat">${cats.map((c) => `<option ${c.name === def ? "selected" : ""}>${esc(c.name)}</option>`).join("")}</select></label>
    <label class="field"><span>Notes (optional)</span><input id="addNotes"></label>
    <div class="actions"><button class="outline" data-act="closesheet">Cancel</button><button class="primary grow" data-act="saveadd">Add</button></div></div>`);
  UI.addOut = true;
}

// ---------------------------------------------------------------- render

function renderTop() {
  const r = route(), sub = SUBPAGES[r.page];
  const title = sub ? (sub[0] || r.arg) : TABS.find((t) => t[0] === r.page)?.[1];
  $("#top").innerHTML = `${sub ? `<button class="iconbtn" data-act="back" aria-label="Back">${icon(I.back)}</button>` : ""}
    <h1 class="grow">${esc(title)}</h1>
    ${S.offline ? `<span class="pill">Preview</span>` : ""}
    <button class="iconbtn ${S.loading || S.saving ? "spin" : ""}" data-act="reload" aria-label="Reload">${icon(I.refresh)}</button>`;
}

function render() {
  if (S.signedIn === null) { main().innerHTML = `<div class="center pad">Loading…</div>`; return; }
  document.body.classList.toggle("signedout", !S.signedIn);
  if (!S.signedIn) { renderSignIn(); return; }
  renderTop();
  const r = route();
  if (S.error === "nofile") { main().innerHTML = noFile(); renderNav(r); return; }
  const pages = { overview: pageOverview, financials: pageFinancials, bankrec: pageBankRec, portfolio: pagePortfolio, more: pageMore,
    tx: pageTransactions, subs: pageSubs, donations: pageDonations, settings: pageSettings, cat: () => pageCategory(r.arg) };
  const y = window.scrollY;
  main().innerHTML = (pages[r.page] || pageOverview)() + `<div class="spacer"></div>`;
  if (render.keepScroll) window.scrollTo(0, y);
  render.keepScroll = false;
  renderNav(r);
}

function renderNav(r) {
  const active = SUBPAGES[r.page] ? SUBPAGES[r.page][1] : r.page;
  $("#nav").innerHTML = TABS.map(([k, l, d]) => `<button data-go="${k}" class="${k === active ? "on" : ""}" aria-label="${l}">${icon(d)}<span>${l}</span></button>`).join("");
}

function noFile() {
  return `<div class="pad"><h2>Couldn't find Budget Data.xlsx</h2><p class="t2">Signed in as ${esc(S.account?.user || "")}. Put your Budget Data.xlsx in OneDrive (for Lena: a folder named <b>Lena Budget</b>), or type where it is.</p>
    <label class="field"><span>OneDrive path</span><input id="pathInput" placeholder="Lena Budget/Budget Data.xlsx"></label>
    <div class="actions"><button class="outline" data-act="signout">Use another account</button><button class="primary grow" data-act="savepath">Open</button></div></div>`;
}

function renderSignIn(err = "") {
  $("#top").innerHTML = "";
  $("#nav").innerHTML = "";
  const demo = ["localhost", "127.0.0.1"].includes(location.hostname);
  main().innerHTML = `<div class="signin"><div class="logo">${icon("M21 18v1a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v1h-9a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h9zm-9-2h10V8H12v8zm4-2.5a1.5 1.5 0 1 1 0-3 1.5 1.5 0 0 1 0 3z", "xl")}</div>
    <h1>Budget</h1><p class="t2">Sign in with the Microsoft account that holds your Budget Data.xlsx in OneDrive. The app reads and updates that workbook directly — nothing is stored anywhere else.</p>
    ${err ? `<p class="neg">${esc(err)}</p>` : ""}
    <button class="primary wide" data-act="signin">Sign in with Microsoft</button>
    ${demo ? `<button class="outline wide" data-act="preview">Preview with sample data</button>` : ""}
    <p class="t2 small">Works on iPhone and Android. Add it to your home screen to use it like an app.</p></div>`;
}

// ---------------------------------------------------------------- events

document.addEventListener("click", async (e) => {
  const el = e.target.closest("[data-act],[data-go]");
  if (!el || el.disabled) return;
  if (el.dataset.go != null) {
    e.preventDefault();
    const g = el.dataset.go;
    if (g.startsWith("cat/") === false && g === "bankrec") UI.rec = { focused: null, skipped: new Set() };
    go(g); window.scrollTo(0, 0); return;
  }
  const act = el.dataset.act, d = el.dataset;
  switch (act) {
    case "signin": Auth.signIn(); break;
    case "preview": startPreview(); break;
    case "signout": Auth.signOut(); S.offline = false; S.signedIn = false; S.path = null; S.wb = null; S.error = null; render(); break;
    case "reload": if (S.offline) say("Preview data reloaded"); refresh(); break;
    case "back": history.length > 1 ? history.back() : go(SUBPAGES[route().page]?.[1] || "overview"); break;
    case "month": if (d.m) { S.month = d.m; S.monthPicked = true; UI.rec = { focused: null, skipped: new Set() }; render.keepScroll = true; render(); } break;
    case "toggle": UI.open.has(d.k) ? UI.open.delete(d.k) : UI.open.add(d.k); render.keepScroll = true; render(); break;
    case "subtoggle": UI.subsOpen.has(d.k) ? UI.subsOpen.delete(d.k) : UI.subsOpen.add(d.k); render.keepScroll = true; render(); break;
    case "report": UI.fin.which = +d.i; render.keepScroll = true; render(); break;
    case "seg": segClick(d.seg, +d.v, el); break;
    case "prices": refreshPrices(); say("Getting live prices…"); break;
    case "csv-report": {
      const r = UI.fin.which === 1 ? F.balanceSheet(S.snap, S.month, S.quotes) : currentReport();
      const slug = r.title.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
      const y = monthYear(S.month), period = UI.fin.which === 4 ? `${y}` : UI.fin.which === 1 ? S.month : UI.fin.ytd ? `${y}-ytd-${S.month}` : S.month;
      download(`${slug}-${period}`, F.reportCsv(r)); break;
    }
    case "csv-tx": {
      const ms = F.monthsIn(S.month, UI.fin.ytd ? "YTD" : "MONTH");
      download(`transactions-${UI.fin.ytd ? `${monthYear(S.month)}-ytd-${S.month}` : S.month}`, F.transactionsCsv(S.snap, ms)); break;
    }
    case "csv-all": download("all-transactions", F.transactionsCsv(S.snap)); break;
    case "csv-don": download(`church-donations-${monthYear(S.month)}`, F.donationsCsv(F.donations(S.snap, S.month, parseFloat(UI.don.fast) || 0, UI.don.included))); break;
    case "doninc": { const inc = UI.don.included; UI.don.included = inc.includes(d.c) ? inc.filter((x) => x !== d.c) : [...inc, d.c]; store.set(userKey("donInc"), UI.don.included); render.keepScroll = true; render(); break; }
    case "edit": editTx(d.id); break;
    case "focus": UI.rec.focused = d.id; assignPanel.state = null; render(); window.scrollTo(0, 0); break;
    case "skip": UI.rec.skipped.add(d.id); UI.rec.focused = null; assignPanel.state = null; render(); break;
    case "pick": assignPanel.state.selected = d.c; rerenderAssign(); break;
    case "save": {
      const p = assignPanel.state, tx = S.snap.transactions.find((t) => t.id === d.id);
      if (!p || !tx) return;
      const inSheet = !!el.closest("#sheet");
      if (inSheet) closeSheet();
      UI.rec.focused = null; assignPanel.state = null;
      await recategorize(tx, p.selected, p.all ? p.match.trim() || null : null, p.all && p.rule);
      break;
    }
    case "closesheet": closeSheet(); break;
    case "budget": budgetDialog(A.category(S.snap, d.c)); break;
    case "savebudget": {
      const v = $("#budgetIn").value.replace(/[$,]/g, "").trim();
      const n = v === "" ? null : parseFloat(v);
      if (v !== "" && !(n >= 0)) { say("Enter an amount, or leave it blank to clear the budget"); return; }
      const cat = A.category(S.snap, d.c); closeSheet(); await setBudget(cat, n && n > 0 ? n : null); break;
    }
    case "add": addDialog(); break;
    case "saveadd": {
      const desc = $("#addDesc").value.trim(), amt = parseFloat($("#addAmt").value.replace(/[$,]/g, "")), day = parseYmd($("#addDate").value);
      if (!desc || !(amt > 0) || day == null) { say("Fill in a description, an amount above 0 and a date"); return; }
      const tx = { id: newTxId(), date: day, budgetMonth: monthOfDay(day), account: $("#addAcct").value, description: desc,
        amount: UI.addOut ? -amt : amt, category: $("#addCat").value, notes: $("#addNotes").value.trim() };
      closeSheet(); await addTransaction(tx); break;
    }
    case "txreview": UI.tx.review = !UI.tx.review; render.keepScroll = true; render(); break;
    case "txmonth": UI.tx.month = d.m || null; render.keepScroll = true; render(); break;
    case "txcat": UI.tx.cat = null; render.keepScroll = true; render(); break;
    case "findpath": { const p = await Workbook.find(null).catch(() => null); if (p) { $("#pathInput").value = p; say("Found " + p); } else say("Couldn't find Budget Data.xlsx in your OneDrive"); break; }
    case "savepath": {
      const p = $("#pathInput").value.trim().replace(/^\/+/, "");
      if (!p) return;
      S.path = p; S.wb = null; store.set(userKey("path"), p); S.error = null; go("overview"); refresh(); break;
    }
  }
});

function segClick(name, v, el) {
  if (name === "ytd") { UI.fin.ytd = v === 1; render.keepScroll = true; render(); }
  else if (name === "port") { UI.port.by = v; render.keepScroll = true; render(); }
  else if (name === "all") { assignPanel.state.all = v === 1; rerenderAssign(); }
  else if (name === "addout") { UI.addOut = v === 0; el.parentElement.querySelectorAll("button").forEach((b, i) => b.classList.toggle("on", i === v)); }
}

function rerenderAssign() {
  const p = assignPanel.state;
  const tx = S.snap.transactions.find((t) => t.id === p.id);
  const inSheet = $("#sheet [data-region=sheetassign]");
  if (inSheet) inSheet.innerHTML = assignPanel(tx);
  else { const c = $(".card.assign"); if (c) c.innerHTML = assignPanel(tx, true); }
}

document.addEventListener("input", (e) => {
  const k = e.target.dataset.input;
  if (!k) return;
  if (k === "txq") { UI.tx.q = e.target.value; const r = $("[data-region=txlist]"); if (r) r.innerHTML = txList(); }
  else if (k === "match") {
    const p = assignPanel.state; p.match = e.target.value;
    const box = e.target.closest(".assignbody");
    box.querySelector("[data-region=matchinfo]").innerHTML = matchInfo(p);
    const save = box.querySelector("[data-act=save]");
    const matches = A.similar(S.snap, p.match), n = Math.max(1, matches.filter((t) => !eqi(t.category, p.selected)).length);
    save.disabled = !(A.category(S.snap, p.selected) && p.match.trim());
    save.textContent = matches.length > 1 ? `Save ${n}` : "Save";
  } else if (k === "fast") { UI.don.fast = e.target.value; clearTimeout(UI.fastT); UI.fastT = setTimeout(() => { render.keepScroll = true; render(); const i = $("[data-input=fast]"); if (i) { i.focus(); i.setSelectionRange(i.value.length, i.value.length); } }, 500); }
});

document.addEventListener("change", (e) => {
  const k = e.target.dataset.change;
  if (k === "month") { S.month = e.target.value; S.monthPicked = true; UI.rec = { focused: null, skipped: new Set() }; render.keepScroll = true; render(); }
  if (k === "rule") assignPanel.state.rule = e.target.checked;
});

window.addEventListener("hashchange", () => { closeSheet(); render(); });

// ---------------------------------------------------------------- start

async function startPreview() {
  try {
    const r = await fetch("demo.json", { cache: "no-store" });
    if (!r.ok) throw new Error();
    S.offline = true; S.signedIn = true; S.account = { user: "Offline preview", oid: "preview" };
    setSnap(await r.json());
    render(); refreshPrices();
  } catch { say("No preview data here"); }
}

async function start() {
  render();
  let res = {};
  try { res = await Auth.handleRedirect(); } catch (e) { res = { error: e.message }; }
  S.account = res.account || Auth.account();
  S.signedIn = !!S.account;
  if (!S.signedIn) { renderSignIn(res.error ? `Sign-in failed: ${res.error}` : res.needsSignIn ? "Please sign in again." : ""); return; }
  const cached = store.get(userKey("snapshot"));
  S.path = store.get(userKey("path"));
  if (cached) setSnap(cached);
  render();
  refresh();
}

if ("serviceWorker" in navigator && !["localhost", "127.0.0.1"].includes(location.hostname)) {
  navigator.serviceWorker.register("sw.js").then((reg) => {
    reg.addEventListener("updatefound", () => {
      const nw = reg.installing;
      nw?.addEventListener("statechange", () => { if (nw.state === "installed" && navigator.serviceWorker.controller) say("App updated — it'll use the new version next time you open it"); });
    });
  }).catch(() => {});
}

start();

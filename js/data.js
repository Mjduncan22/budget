// Reads and writes the tables in Budget Data.xlsx through Microsoft Graph (same endpoints the
// Android app uses: table range read, worksheet range PATCH, table rows/add).
import { accessToken } from "./auth.js";
import { lc, todayEpochDay, ymd, monthOfDay } from "./util.js";

const BASE = "https://graph.microsoft.com/v1.0";
export const DEFAULT_PATHS = ["Documents/Finances/Budget/Budget Data.xlsx", "Lena Budget/Budget Data.xlsx", "Budget Data.xlsx"];
const EXCEL_EPOCH_OFFSET = 25569; // Excel serial of 1970-01-01

export class GraphError extends Error { constructor(code, msg) { super(msg); this.code = code; } }

const enc = (s) => encodeURIComponent(s);
const encPath = (p) => p.replace(/^\/+|\/+$/g, "").split("/").map(enc).join("/");

async function request(method, url, body) {
  for (let attempt = 0; ; attempt++) {
    const token = await accessToken();
    const r = await fetch(url, {
      method, headers: { Authorization: `Bearer ${token}`, ...(body ? { "Content-Type": "application/json" } : {}) },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (r.ok) { const t = await r.text(); return t ? JSON.parse(t) : null; }
    if ([429, 503, 504].includes(r.status) && attempt < 3) { await new Promise((res) => setTimeout(res, 1500 * (attempt + 1))); continue; }
    let msg = "";
    try { msg = (await r.json()).error?.message || ""; } catch { /* not json */ }
    if (r.status === 401) throw new GraphError(401, "Microsoft sign-in expired. Please sign in again.");
    if (r.status === 404) throw new GraphError(404, "Couldn't find the workbook or table in OneDrive. " + msg);
    throw new GraphError(r.status, `Excel/OneDrive error ${r.status}: ${msg}`);
  }
}

const colIndex = (letters) => [...letters.toUpperCase()].reduce((a, c) => a * 26 + (c.charCodeAt(0) - 64), 0) - 1;
export function colLetters(i) { let n = i + 1, s = ""; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }

function parseStart(address) {
  const cell = address.split("!").pop().split(":")[0].replace(/\$/g, "");
  const letters = cell.match(/^[A-Z]+/i)?.[0] || "A";
  return { col: colIndex(letters), row: +cell.slice(letters.length) || 1 };
}

const text = (v) => {
  if (v == null) return "";
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : String(v);
  return String(v);
};
const num = (v) => {
  if (v == null || v === "") return null;
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace(/[$,]/g, "").trim());
  return Number.isFinite(n) ? n : null;
};
const day = (v) => { const n = num(v); return n == null ? null : Math.floor(n) - EXCEL_EPOCH_OFFSET; };

class Table {
  constructor(obj) {
    const values = obj.values || [];
    const { col, row } = parseStart(obj.address || "A1");
    this.firstCol = col; this.firstRow = row;
    this.header = (values[0] || []).map(text);
    this.rows = values.slice(1);
  }
  c(name) { return this.header.findIndex((h) => lc(h) === lc(name)); }
  sheetRow(i) { return this.firstRow + 1 + i; }
  letter(name) { const i = this.c(name); return i < 0 ? null : colLetters(this.firstCol + i); }
}

export class Workbook {
  constructor(path) { this.path = path; this.itemId = null; this.cols = {}; }

  async open() {
    if (this.itemId) return;
    const o = await request("GET", `${BASE}/me/drive/root:/${encPath(this.path)}`);
    this.itemId = o.id;
  }

  /** Looks for the workbook in the usual places, then searches OneDrive. Returns the path or null. */
  static async find(preferred) {
    for (const p of [preferred, ...DEFAULT_PATHS].filter(Boolean)) {
      try { await request("GET", `${BASE}/me/drive/root:/${encPath(p)}?$select=id`); return p; } catch (e) { if (e.code !== 404) throw e; }
    }
    try {
      const r = await request("GET", `${BASE}/me/drive/root/search(q='Budget Data.xlsx')?$select=name,parentReference,id`);
      const hit = (r.value || []).find((x) => x.name === "Budget Data.xlsx" && x.parentReference?.path);
      if (hit) return (hit.parentReference.path.replace(/^\/drive\/root:\/?/, "") + "/" + hit.name).replace(/^\//, "");
    } catch { /* search not available */ }
    return null;
  }

  wb() { return `${BASE}/me/drive/items/${this.itemId}/workbook`; }
  async table(name) { return new Table(await request("GET", `${this.wb()}/tables/${enc(name)}/range?$select=address,values`)); }
  async cell(sheet, addr) { const o = await request("GET", `${this.wb()}/worksheets/${enc(sheet)}/range(address='${addr}')?$select=values`); return o?.values?.[0]?.[0]; }
  async write(sheet, addr, values) { await request("PATCH", `${this.wb()}/worksheets/${enc(sheet)}/range(address='${addr}')`, { values }); }
  async addRows(table, values) { await request("POST", `${this.wb()}/tables/${enc(table)}/rows/add`, { values }); }

  async load() {
    await this.open();
    const soft = (name) => this.table(name).catch(() => null);
    const [tx, cat, rules, hold, acct, bal, sales] = await Promise.all([
      this.table("tblTransactions"), this.table("tblCategories"), this.table("tblRules"), this.table("tblHoldings"),
      soft("tblAccounts"), soft("tblBalances"), soft("tblSales"),
    ]);
    return {
      transactions: this.parseTx(tx), categories: this.parseCats(cat), rules: parseRules(rules), holdings: parseHoldings(hold),
      accounts: acct ? parseAccounts(acct) : [], balances: bal ? parseBalances(bal) : [], sales: sales ? parseSales(sales) : [],
      loadedAt: Date.now(),
    };
  }

  parseTx(t) {
    const c = (n) => t.c(n);
    const id = c("Transaction ID"), date = c("Date"), month = c("Budget Month"), acct = c("Account"), desc = c("Description"),
      amt = c("Amount"), cat = c("Category"), bank = c("Bank Category"), src = c("Source"), notes = c("Notes");
    if (id < 0 || date < 0 || amt < 0 || cat < 0) throw new Error("Transactions table is missing required columns");
    this.cols.id = t.letter("Transaction ID"); this.cols.cat = t.letter("Category");
    const out = [];
    t.rows.forEach((r, i) => {
      const txId = text(r[id]); const d = day(r[date]);
      if (!txId || d == null) return;
      let bm = text(r[month]).trim();
      if (!/^\d{4}-\d{2}$/.test(bm)) { const n = num(r[month]); bm = n && n > 20000 ? monthOfDay(Math.floor(n) - EXCEL_EPOCH_OFFSET) : monthOfDay(d); }
      out.push({ id: txId, date: d, budgetMonth: bm, account: text(r[acct]), description: text(r[desc]), amount: num(r[amt]) || 0,
        category: text(r[cat]) || "Uncategorized", bankCategory: text(r[bank]), source: text(r[src]), notes: text(r[notes]), sheetRow: t.sheetRow(i) });
    });
    return out;
  }

  parseCats(t) {
    const g = t.c("Group"), n = t.c("Category"), ty = t.c("Type"), b = t.c("Monthly Budget");
    this.cols.budget = t.letter("Monthly Budget");
    const out = [];
    t.rows.forEach((r, i) => {
      const name = text(r[n]);
      if (!name) return;
      const bv = num(r[b]);
      out.push({ group: text(r[g]), name, type: text(r[ty]) || "Expense", budget: bv && bv > 0 ? bv : null, sheetRow: t.sheetRow(i) });
    });
    return out;
  }

  // ---- writes ---------------------------------------------------------

  async locateRow(tx) {
    const idCol = this.cols.id || "A";
    if (tx.sheetRow > 0 && text(await this.cell("Transactions", `${idCol}${tx.sheetRow}`)) === tx.id) return tx.sheetRow;
    const t = await this.table("tblTransactions");
    const c = t.c("Transaction ID");
    const i = t.rows.findIndex((r) => text(r[c]) === tx.id);
    if (i < 0) throw new Error("That transaction is no longer in the workbook. Reload and try again.");
    return t.sheetRow(i);
  }

  /** Sets the category of each transaction. Returns {id: sheetRow}. */
  async setCategories(txs, category) {
    await this.open();
    const rows = {};
    if (txs.length === 1) rows[txs[0].id] = await this.locateRow(txs[0]);
    else {
      const t = await this.table("tblTransactions");
      const c = t.c("Transaction ID");
      this.cols.cat = t.letter("Category") || this.cols.cat;
      t.rows.forEach((r, i) => { rows[text(r[c])] = t.sheetRow(i); });
    }
    const col = this.cols.cat || "H";
    for (const tx of txs) { const row = rows[tx.id]; if (row) await this.write("Transactions", `${col}${row}`, [[category]]); }
    return rows;
  }

  async addRule(rule) { await this.open(); await this.addRows("tblRules", [[rule.keyword, rule.category, "", rule.priority]]); }

  async setBudget(cat, amount) {
    await this.open();
    await this.write("Categories", `${this.cols.budget || "D"}${cat.sheetRow}`, [[amount == null ? "" : amount]]);
  }

  async addTransaction(tx) {
    await this.open();
    const serial = tx.date + EXCEL_EPOCH_OFFSET;
    const group = '=IFERROR(INDEX(Categories!$A:$A,MATCH([@Category],Categories!$B:$B,0)),"Review")';
    await this.addRows("tblTransactions", [[tx.id, serial, "'" + tx.budgetMonth, tx.account, tx.description, tx.amount, group, tx.category, "", "App", tx.notes]]);
  }
}

function parseRules(t) {
  const k = t.c("Keyword"), c = t.c("Category"), a = t.c("Account (optional)"), p = t.c("Priority");
  return t.rows.map((r) => {
    const kw = text(r[k]);
    return kw ? { keyword: kw, category: text(r[c]), account: text(r[a]), priority: Math.trunc(num(r[p]) ?? kw.length) } : null;
  }).filter(Boolean);
}

function parseHoldings(t) {
  const s = t.c("Symbol"), ty = t.c("Asset Type"), q = t.c("Quantity"), c = t.c("Cost Basis"), h = t.c("Held In"), lp = t.c("Last Price");
  return t.rows.map((r) => {
    const sym = text(r[s]).toUpperCase(), quantity = num(r[q]) || 0;
    if (!sym || quantity <= 0) return null;
    return { symbol: sym, assetType: text(r[ty]), quantity, costBasis: num(r[c]) || 0, heldIn: text(r[h]), lastPrice: lp >= 0 ? num(r[lp]) : null };
  }).filter(Boolean);
}

function parseAccounts(t) {
  const n = t.c("Account"), ty = t.c("Type");
  return t.rows.map((r) => (text(r[n]) ? { name: text(r[n]), type: text(r[ty]) } : null)).filter(Boolean);
}

function parseBalances(t) {
  const d = t.c("Date"), a = t.c("Account"), s = t.c("Sub-account"), b = t.c("Balance");
  return t.rows.map((r) => {
    const dd = day(r[d]), bal = num(r[b]);
    return dd == null || bal == null ? null : { date: dd, account: text(r[a]), subAccount: text(r[s]), balance: bal };
  }).filter(Boolean);
}

function parseSales(t) {
  const d = t.c("Date"), s = t.c("Symbol"), q = t.c("Quantity"), p = t.c("Proceeds"), c = t.c("Cost Basis"), a = t.c("Account"), src = t.c("Source ID");
  return t.rows.map((r) => {
    const dd = day(r[d]), sym = text(r[s]);
    if (dd == null || !sym) return null;
    return { date: dd, symbol: sym, quantity: num(r[q]) || 0, proceeds: num(r[p]) || 0, costBasis: num(r[c]), account: text(r[a]), source: text(r[src]) };
  }).filter(Boolean);
}

export function newTxId() { return "app-" + crypto.getRandomValues(new Uint32Array(1))[0].toString(16).padStart(8, "0"); }
export { todayEpochDay, ymd };

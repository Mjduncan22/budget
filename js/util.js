// Small shared helpers: dates as "epoch days" (days since 1970-01-01, UTC) and "YYYY-MM" month strings.

export const DAY_MS = 86400000;

export function todayEpochDay() {
  const d = new Date();
  return Math.floor(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / DAY_MS);
}

export function epochDayOf(y, m, d) { return Math.floor(Date.UTC(y, m - 1, d) / DAY_MS); }

export function ymd(epochDay) {
  const d = new Date(epochDay * DAY_MS);
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
}

export function parseYmd(s) {
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec((s || "").trim());
  if (!m) return null;
  const y = +m[1], mo = +m[2], d = +m[3];
  const day = epochDayOf(y, mo, d);
  return ymd(day) === `${m[1]}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}` ? day : null;
}

export function monthOfDay(epochDay) { return ymd(epochDay).slice(0, 7); }

export function currentMonth() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

export function isMonth(s) { return /^\d{4}-(0[1-9]|1[0-2])$/.test(s || ""); }

export function addMonths(month, n) {
  let [y, m] = month.split("-").map(Number);
  m += n;
  while (m < 1) { m += 12; y--; }
  while (m > 12) { m -= 12; y++; }
  return `${y}-${String(m).padStart(2, "0")}`;
}

export function monthYear(month) { return +month.slice(0, 4); }
export function monthNum(month) { return +month.slice(5, 7); }

/** Last day of the month as an epoch day. */
export function endOfMonth(month) {
  const [y, m] = month.split("-").map(Number);
  return Math.floor(Date.UTC(y, m, 0) / DAY_MS);
}
export function startOfMonth(month) {
  const [y, m] = month.split("-").map(Number);
  return epochDayOf(y, m, 1);
}

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

export function monthLabel(month, short = false) {
  if (!isMonth(month)) return month;
  const name = MONTHS[monthNum(month) - 1];
  return short ? name.slice(0, 3) : `${name} ${monthYear(month)}`;
}

export function dayLabel(epochDay) {
  const d = new Date(epochDay * DAY_MS);
  return `${MONTHS[d.getUTCMonth()].slice(0, 3)} ${d.getUTCDate()}`;
}

const cur = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
const curWhole = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0, minimumFractionDigits: 0 });

export function money(v, whole = false) {
  const f = whole ? curWhole : cur;
  const x = v || 0;
  return x < 0 ? "-" + f.format(-x) : f.format(x);
}

export function pct(v) { return (v >= 0 ? "+" : "") + v.toFixed(1) + "%"; }

export function qty(q) {
  return q.toLocaleString("en-US", { maximumFractionDigits: 8 });
}

export const lc = (s) => (s || "").toLowerCase();
export const eqi = (a, b) => lc(a) === lc(b);
export const sum = (arr, f = (x) => x) => arr.reduce((a, x) => a + (f(x) || 0), 0);

export function groupBy(arr, keyFn) {
  const m = new Map();
  for (const x of arr) {
    const k = keyFn(x);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(x);
  }
  return m;
}

export function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

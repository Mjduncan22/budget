// Microsoft sign-in for a single-page app: OAuth 2.0 authorization code + PKCE against the
// "consumers" (personal Microsoft account) endpoint. No library needed; tokens stay in this
// browser only. The app asks only for OneDrive file access (Files.ReadWrite) and your name.

export const CLIENT_ID = "285ac8de-696e-4a0f-be8b-ca6cec337e33";
const AUTHORITY = "https://login.microsoftonline.com/consumers/oauth2/v2.0";
const SCOPES = "openid profile offline_access User.Read Files.ReadWrite";
const KEY = "budget.auth.v1";
const PENDING = "budget.auth.pending";

export class NotSignedIn extends Error {}

/** The address Microsoft sends you back to: this folder, e.g. https://you.github.io/budget/ */
export function redirectUri() {
  const u = new URL(location.href);
  return u.origin + u.pathname.replace(/[^/]*$/, "");
}

const store = {
  get(k) { try { return JSON.parse(localStorage.getItem(k) || "null"); } catch { return null; } },
  set(k, v) { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, JSON.stringify(v)); } catch { /* storage blocked */ } },
};

function b64url(bytes) {
  let s = "";
  for (const b of new Uint8Array(bytes)) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function randomString(n = 48) { return b64url(crypto.getRandomValues(new Uint8Array(n))); }
async function challenge(verifier) { return b64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier))); }

function claims(idToken) {
  try {
    const p = idToken.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
    return JSON.parse(decodeURIComponent(escape(atob(p + "===".slice((p.length + 3) % 4)))));
  } catch { return {}; }
}

/** Sends the browser to Microsoft's sign-in page. silent = try without showing anything. */
export async function signIn({ silent = false, hint = "" } = {}) {
  const verifier = randomString(64), state = randomString(16);
  store.set(PENDING, { verifier, state, silent, back: location.hash });
  const q = new URLSearchParams({
    client_id: CLIENT_ID, response_type: "code", redirect_uri: redirectUri(), response_mode: "query",
    scope: SCOPES, state, code_challenge: await challenge(verifier), code_challenge_method: "S256",
  });
  if (silent) q.set("prompt", "none");
  else q.set("prompt", "select_account");
  if (hint) q.set("login_hint", hint);
  location.assign(`${AUTHORITY}/authorize?${q}`);
  return new Promise(() => {}); // page is navigating away
}

async function tokenRequest(params) {
  const r = await fetch(`${AUTHORITY}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: CLIENT_ID, scope: SCOPES, ...params }),
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) { const e = new Error(j.error_description?.split("\r\n")[0] || j.error || `Sign-in failed (${r.status})`); e.code = j.error; throw e; }
  return j;
}

function save(j, prev = {}) {
  const c = j.id_token ? claims(j.id_token) : {};
  const a = {
    access: j.access_token, refresh: j.refresh_token || prev.refresh,
    expires: Date.now() + (j.expires_in - 120) * 1000,
    name: c.name || prev.name || "", user: c.preferred_username || c.email || prev.user || "",
    oid: c.oid || prev.oid || "",
  };
  store.set(KEY, a);
  return a;
}

/**
 * Call once on page load. Finishes a sign-in that just came back from Microsoft.
 * Returns { account } when signed in, { error } when Microsoft said no, or {} otherwise.
 */
export async function handleRedirect() {
  const u = new URL(location.href);
  const code = u.searchParams.get("code"), err = u.searchParams.get("error");
  if (!code && !err) return {};
  const pending = store.get(PENDING);
  store.set(PENDING, null);
  history.replaceState(null, "", redirectUri() + (pending?.back || ""));
  if (err) {
    // A silent attempt that needs you to sign in again isn't an error worth showing.
    if (pending?.silent) return { needsSignIn: true };
    return { error: u.searchParams.get("error_description") || err };
  }
  if (!pending || pending.state !== u.searchParams.get("state")) return { error: "Sign-in expired. Please try again." };
  const j = await tokenRequest({ grant_type: "authorization_code", code, redirect_uri: redirectUri(), code_verifier: pending.verifier });
  return { account: save(j) };
}

export function account() { const a = store.get(KEY); return a && a.refresh ? a : null; }

let refreshing = null;

/** A valid access token, refreshing it when needed. Throws NotSignedIn when you must sign in again. */
export async function accessToken() {
  const a = store.get(KEY);
  if (!a) throw new NotSignedIn();
  if (a.access && Date.now() < a.expires) return a.access;
  if (!refreshing) {
    refreshing = tokenRequest({ grant_type: "refresh_token", refresh_token: a.refresh })
      .then((j) => save(j, a).access)
      .catch(async (e) => {
        // Browser sign-ins last about a day; Microsoft usually re-approves silently.
        if (["invalid_grant", "interaction_required", "login_required"].includes(e.code)) {
          await signIn({ silent: true, hint: a.user });
        }
        throw e;
      })
      .finally(() => { refreshing = null; });
  }
  return refreshing;
}

export function signOut() {
  store.set(KEY, null);
}

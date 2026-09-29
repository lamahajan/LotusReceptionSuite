// ═══════════════════════════════════════════════════════════
//  Razorpay Proxy — Cloudflare Worker (v10 — /create-qr now uses the Razorpay
//  QR Codes API (upi_qr) instead of Payment Links. The `short_url` field
//  (Firebase + Sheet + API response) now holds the UPI string
//  `upi://pay?...` (Razorpay's `image_content`). Needs the qr_image_content
//  feature enabled on the Razorpay / HDFC CollectNow account, and the
//  `qr_code.credited` webhook event turned on. Old plink_ records are still
//  handled by /close-qr, /delete-qr and the payment_link.paid webhook.)
//
//  v9 — /payment-status has a
//  4-step fallback chain instead of only checking the QR node:
//
//    1. QR Firebase ({QR_CODE}/{qr_id})            — live, fast, usual case
//    2. TXN Firebase ({TXN_CODE}/*, by qr_id)      — QR record already gone
//       (e.g. after the monthly cleanup) but the payment did happen
//    3. Google Sheet — TXN tab (historical, by qr_id) — Firebase itself got
//       wiped/lost but the sheet log still has the row
//    4. Google Sheet — QR tab (historical, by qr_id) — only used to tell
//       "closed with no payment" apart from "never existed": if the QR row
//       is there and status === "closed" → invalid
//    5. Nothing found anywhere → invalid
//
//  Steps 3 & 4 need a small doGet() lookup handler added to your Google
//  Apps Script — see sheets-lookup-addition.gs. Without that, the worker
//  just skips steps 3/4 (GOOGLE_SCRIPT_URL fetch will fail silently) and
//  falls straight through to "invalid".
// ═══════════════════════════════════════════════════════════
//  QR, Due List, Transactions all separate Firebase projects, each
//  addressed as URL + "code" path segment, no auth secret anymore):
//    {QR_FIREBASE_CODE}/{id}                            — payment link records (QR_FIREBASE_URL)
//    DUELIST/0,1,2...                                   — due-list, flat array (FIREBASE_URL)
//    {TXN_FIREBASE_CODE}/{payment_id or order_id}        — flat: captured payments AND
//                                                          pending orders live side by side
//                                                          (TXN_FIREBASE_URL) — keys never
//                                                          collide: pay_xxx vs order_xxx
//
//  No ?auth= secret is sent anymore — access relies entirely on your Firebase
//  Realtime Database security RULES. Since there's no secret gate at all now,
//  it is IMPORTANT that your Firebase rules are NOT public (.read/.write: true).
//  Lock them down (e.g. to a Firebase Auth check) or anyone with the URL can
//  read/write/delete everything directly, bypassing this worker entirely.
//
//  Endpoints:
//    POST /create-qr              → 🔒 requires X-App-Token — create UPI QR (QR Codes API) + save to Firebase
//    GET  /fetch-payment/:qr_id   → poll Razorpay directly (unauthenticated, low sensitivity)
//    GET  /payment-status/:qr_id  → read status, with fallback chain (see above)
//    POST /close-qr/:qr_id        → 🔒 requires X-App-Token — close QR + update Firebase
//    POST /webhook                → Razorpay webhook handler (protected by HMAC signature)
//    GET  /qr-by-adm/:adm         → 🔒 requires X-App-Token — find existing QR for an admission number
//    DELETE /delete-qr/:qr_id     → 🔒 requires X-App-Token — delete QR
//    POST /reconcile-qrs          → 🔒 requires X-App-Token — closes any pending/active QR
//                                    whose adm is NOT in the admList sent in the body.
//                                    Call right after a Due List "Upload xlsm & Full
//                                    Replace" so a student who paid manually (cash) and
//                                    got dropped from the sheet doesn't keep a stale QR
//                                    open until the monthly wipe.
//    POST /register-fcm-token     → 🔒 requires X-App-Token — stores a device's FCM
//                                    push-notification token (see reception.html's
//                                    "Enable Notifications" button). Every paid
//                                    transaction then pushes a notification to it.
//    POST /fcm-heartbeat          → 🔒 requires X-App-Token — called by reception.html
//                                    whenever the page is active. Marks the device as
//                                    "recently seen" (real-time push eligible) and, if
//                                    it had payments queued from while it was away,
//                                    sends one summary push and clears the queue.
//    POST /fee-lookup             → public-safe lookup requiring BOTH admission number AND
//                                    registered mobile number to match. Never exposes secrets.
//                                    If the admission number has no due-list record at all,
//                                    falls back to the master Student DB — if adm+mobile match
//                                    there, the student is reported as "paid" (balance 0) instead
//                                    of "not found", since no due-list entry means no pending fee.
//    POST /create-order           → public-safe (same adm+mobile dual-match design as
//                                    /fee-lookup). Creates a Razorpay ORDER (Orders API) for
//                                    Custom Checkout on the website pay page. Amount is NEVER
//                                    taken from the client — always the live Due List balance.
//                                    The order record ({TXN_CODE}/{order_id}) is saved straight
//                                    into the TXN Firebase project — this flow never creates
//                                    anything in the QR project at all.
//    POST /verify-payment         → public, protected by its own HMAC signature check (same
//                                    trust model as /webhook). LEGACY — was called by the
//                                    website pay page's Checkout `handler` under Standard
//                                    Checkout. Kept for backward compatibility, but the
//                                    website's Hosted Checkout flow no longer calls this route.
//    POST /checkout-callback      → public, protected by its own HMAC signature check (same
//                                    trust model as /verify-payment). This is the Hosted
//                                    Checkout `callback_url` — Razorpay's hosted payment page
//                                    POSTs here directly (a real browser navigation, not a
//                                    fetch() from the site) with razorpay_order_id /
//                                    razorpay_payment_id / razorpay_signature on success, or
//                                    error[...] fields on failure/decline. Does the same
//                                    settlement work /verify-payment used to do, then
//                                    303-redirects the browser back to the pay page with the
//                                    result in the query string. The /webhook payment.captured
//                                    handler (order_id based) is still a backup in case this
//                                    call never lands (e.g. the parent's connection drops mid
//                                    redirect).
//
//  CollectNow / HDFC compliance note (added when Standard Checkout was
//  replaced with Hosted Checkout): the website pay page now POSTs an HTML
//  form straight to https://api.razorpay.com/v1/checkout/embedded instead
//  of opening the checkout.js modal. See DigiFee_Web.html's payNow().
//
//  Two payment paths, two Firebase projects, one rule: the QR project only
//  ever holds PENDING qr data (created by /create-qr for the admin's
//  WhatsApp due-reminder flow). The instant a payment_link is paid, its
//  QR entry is deleted — the transaction lives ONLY in TXN from then
//  on ({TXN_CODE}/{payment_id}). The /create-order + /verify-payment
//  path (website search-and-pay) never touches the QR project at all —
//  its {order_id} record lives directly under TXN. Either path, once
//  a payment settles, the paid amount is subtracted from the matching
//  DUELIST row; if that brings the balance to 0 (or below), the whole
//  DUELIST record is deleted rather than left at balance: 0.
// ═══════════════════════════════════════════════════════════
//  Environment Variable (Cloudflare → Settings → Variables):
//    APP_TOKEN = <a long random string you generate yourself>
//
//  This token is a simple shared-secret gate — it stops random scanners / anyone who
//  doesn't know the token from calling admin-only endpoints. It is NOT a substitute for
//  real per-staff login (e.g. Cloudflare Access, Firebase Auth) if you need stronger
//  guarantees later, but it closes the "wide open" gap that existed before.
//
//  Generate one with, e.g.: openssl rand -hex 32
// ═══════════════════════════════════════════════════════════
//  Environment Variables (existing):
//    RZP_KEY_ID         = rzp_live_xxxxxxxx
//    RZP_KEY_SECRET     = your_razorpay_secret
//    RZP_WEBHOOK_SECRET = your_webhook_secret
//    ALLOWED_ORIGINS    = https://reception.lisdnn.org,https://fee.lisdnn.org
//    FIREBASE_URL        = https://your-due-project-default-rtdb.firebasedatabase.app
//                           (Due List only now — holds just the DUELIST node)
//    QR_FIREBASE_URL      = https://your-qr-project-default-rtdb.firebaseio.com
//                            (separate project, just for QR/payment-link records)
//    QR_FIREBASE_CODE     = QR    (path segment — data lives at {QR_FIREBASE_URL}/
//                                   {QR_FIREBASE_CODE}/{id}. No auth secret.)
//    TXN_FIREBASE_URL     = https://your-txn-project-default-rtdb.firebaseio.com
//                            (separate project, just for transactions)
//    TXN_FIREBASE_CODE    = TXN   (path segment under which transactions live —
//                                   same pattern as STUDENT_FIREBASE_CODE below.
//                                   Data lives at {TXN_FIREBASE_URL}/{TXN_FIREBASE_CODE}/
//                                   {payment_id}. No auth secret.)
//    GOOGLE_SCRIPT_URL    = the /exec URL of your deployed Google Apps Script Web
//                           App — see google-sheets-logger.gs + sheets-lookup-addition.gs.
//                           Used both to log every QR creation / captured payment as a
//                           row (fire-and-forget, existing behaviour) AND now as a
//                           read-side fallback for /payment-status (steps 3 & 4 above).
//    SITE_URL             = OPTIONAL, e.g. https://fee.lisdnn.org — the pay page's own
//                           base URL. Used by /checkout-callback to build the 303
//                           redirect back to the pay page once Hosted Checkout finishes.
//                           If unset, falls back to the ALLOWED_ORIGINS entry containing
//                           "fee.", then to the first ALLOWED_ORIGINS entry.
// ═══════════════════════════════════════════════════════════
//  Student DB (separate project — unrelated to the consolidation above):
//    STUDENT_FIREBASE_URL  = same Firebase URL you put into the dashboard's
//                             "Firebase — Students" (Reference Register) settings
//    STUDENT_FIREBASE_CODE = the "code" you put into that same settings box
//                             (e.g. STUDENT). This is the path segment the
//                             dashboard writes student records under: {code}/students
// ═══════════════════════════════════════════════════════════

const RZP_BASE = 'https://api.razorpay.com/v1';

// Endpoints that require the X-App-Token header to match env.APP_TOKEN.
// Everything NOT in this list stays open on purpose:
//  - /payment-status, /fetch-payment → parent's payment page needs these, no login there
//  - /webhook   → protected separately by Razorpay's HMAC signature
//  - /fee-lookup → protected separately by its own adm+mobile dual-match design
const PROTECTED_PREFIXES = ['/create-qr', '/close-qr/', '/delete-qr/', '/qr-by-adm/', '/reconcile-qrs', '/register-fcm-token', '/fcm-heartbeat'];

const now = () => Math.floor(Date.now() / 1000);

// ── Google Sheets logger (fire-and-forget) ──
// Mirrors QR-creation, captured-payment, close, and monthly-cleanup events
// to a Google Apps Script Web App, which appends a row to a Google Sheet.
// Never blocks or fails the caller — if GOOGLE_SCRIPT_URL isn't set, or the
// call fails, it's silently skipped (same pattern as the Firebase writes).
// Module-level (not inside fetch()) so both fetch() and scheduled() can use it.
const gsLog = (env, payload) => {
  if (!env.GOOGLE_SCRIPT_URL) return Promise.resolve();
  // Returns the promise (instead of firing-and-forgetting internally) so
  // call sites that need the write to actually finish before the Worker's
  // response goes out — e.g. /verify-payment — can `await gsLog(...)`.
  // Existing call sites that don't await it keep the exact same
  // fire-and-forget behaviour as before (they just ignore the returned
  // promise), so nothing else changes.
  return fetch(env.GOOGLE_SCRIPT_URL, {
    method : 'POST',
    headers: { 'Content-Type': 'application/json' },
    body   : JSON.stringify(payload),
  }).catch(() => {});
};

// ── Google Sheets lookup (read-side, used by /payment-status fallback) ──
// Calls GOOGLE_SCRIPT_URL as a GET with ?action=lookup&qr_id=..., expecting
// the doGet() handler from sheets-lookup-addition.gs. Returns null on any
// failure (missing config, network error, bad JSON, not found) so callers
// can just check "if (result) ...".
const gsLookup = async (env, qrId) => {
  if (!env.GOOGLE_SCRIPT_URL) return null;
  try {
    const res = await fetch(
      `${env.GOOGLE_SCRIPT_URL}?action=lookup&qr_id=${encodeURIComponent(qrId)}`
    );
    const data = await res.json();
    return (data && data.found) ? data : null;
  } catch (e) {
    return null;
  }
};

// ── Firebase Cloud Messaging (push notifications on every paid transaction) ──
//
//   Requires a Firebase service account (Project Settings → Service Accounts
//   → Generate new private key, in the Firebase console — same project the
//   reception app's FCM config points at). Three secrets, set with
//   `wrangler secret put`:
//     FCM_PROJECT_ID    = the service account JSON's "project_id"
//     FCM_CLIENT_EMAIL  = the service account JSON's "client_email"
//     FCM_PRIVATE_KEY   = the service account JSON's "private_key", EXACTLY
//                          as-is including the literal \n escapes — paste the
//                          whole "-----BEGIN PRIVATE KEY-----...END..." block.
//
//   Device tokens (one per browser/device that clicked "Enable
//   Notifications" in reception.html) are stored in the QR Firebase project,
//   as a TOP-LEVEL sibling node — `{QR_BASE}/fcm_tokens/{sanitized_token}` —
//   deliberately NOT nested under {QR_FIREBASE_CODE}. The monthly cleanup
//   below wipes the entire {QR_FIREBASE_CODE} subtree; nesting fcm_tokens
//   inside it would silently delete every registered device each month.
//   Registered via POST /register-fcm-token below.
//
//   If any of the three secrets are missing, sendFcmNotification() is a
//   silent no-op — nothing else in the payment flow depends on this.

// Base64url encode — JWTs use base64URL (no padding, - and _ instead of + and /).
const b64url = (bytes) => {
  let bin = '';
  const arr = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  for (let i = 0; i < arr.length; i++) bin += String.fromCharCode(arr[i]);
  return btoa(bin).replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_');
};

// Signs a Google OAuth2 service-account JWT (RS256) and exchanges it for a
// short-lived access token scoped to Firebase Cloud Messaging. No caching —
// this only runs in the background (ctx.waitUntil), once per paid
// transaction, so the extra ~100-200ms here never affects anything the
// parent or reception admin is waiting on.
async function getFcmAccessToken(env) {
  const iat = now();
  const header = { alg: 'RS256', typ: 'JWT' };
  const claim = {
    iss  : env.FCM_CLIENT_EMAIL,
    scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud  : 'https://oauth2.googleapis.com/token',
    iat,
    exp  : iat + 3600,
  };
  const enc = new TextEncoder();
  const unsigned = `${b64url(enc.encode(JSON.stringify(header)))}.${b64url(enc.encode(JSON.stringify(claim)))}`;

  const pem = String(env.FCM_PRIVATE_KEY || '').replace(/\\n/g, '\n');
  const der = pem
    .replace(/-----BEGIN PRIVATE KEY-----/, '')
    .replace(/-----END PRIVATE KEY-----/, '')
    .replace(/\s+/g, '');
  const derBytes = Uint8Array.from(atob(der), (c) => c.charCodeAt(0));

  const key = await crypto.subtle.importKey(
    'pkcs8', derBytes.buffer,
    { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
    false, ['sign']
  );
  const sig = await crypto.subtle.sign('RSASSA-PKCS1-v1_5', key, enc.encode(unsigned));
  const jwt = `${unsigned}.${b64url(sig)}`;

  const tokenRes = await fetch('https://oauth2.googleapis.com/token', {
    method : 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body   : `grant_type=${encodeURIComponent('urn:ietf:params:oauth:grant-type:jwt-bearer')}&assertion=${encodeURIComponent(jwt)}`,
  });
  const tokenData = await tokenRes.json();
  return tokenData.access_token || null;
}

// Fetches every registered device token, sends the notification to each
// (FCM's v1 API has no multicast — one HTTP call per token, fine for the
// handful of reception devices this is for), and prunes any token FCM
// reports as unregistered/invalid so the list doesn't grow stale forever.
// Sends one FCM push to a single token; prunes the token from Firebase if
// FCM reports it as dead (UNREGISTERED / invalid). Shared by both the
// real-time path and the heartbeat's summary-catch-up path below.
async function fcmSendToToken(env, accessToken, QR_BASE, key, token, { title, body, data }) {
  try {
    const res = await fetch(`https://fcm.googleapis.com/v1/projects/${env.FCM_PROJECT_ID}/messages:send`, {
      method : 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${accessToken}` },
      body   : JSON.stringify({
        message: {
          token,
          notification: { title, body },
          data: Object.fromEntries(Object.entries(data || {}).map(([k, v]) => [k, String(v)])),
          webpush: { fcm_options: { link: '/' } },
        },
      }),
    });
    if (res.status === 404 || res.status === 400) {
      const errBody = await res.json().catch(() => null);
      const code = errBody?.error?.details?.[0]?.errorCode || '';
      if (code === 'UNREGISTERED' || res.status === 404) {
        await fetch(`${QR_BASE}/fcm_tokens/${key}.json`, { method: 'DELETE' }).catch(() => {});
      }
    }
  } catch (e) {}
}

// How long since a device's last heartbeat before we stop pushing it
// real-time alerts and start queuing a summary for it instead. 10 minutes:
// long enough that a device mid-tab-switch or on a flaky connection doesn't
// get wrongly bucketed as "away", short enough that it kicks in fast for a
// genuinely closed/off device.
const FCM_AWAY_AFTER_SEC = 10 * 60;

// Called on every paid transaction. For each registered device: if its
// last heartbeat was recent, push the real-time "Payment Received" alert
// right now (unchanged behaviour). If it's been quiet longer than
// FCM_AWAY_AFTER_SEC, DON'T push — instead bump that device's
// pending_count/pending_amount in Firebase. fcmHeartbeat() (called by
// reception.html the moment that device comes back online) is what
// actually flushes the queued count into a single summary push — this
// function only ever decides real-time-vs-queue, never sends a summary
// itself.
async function notifyPaymentToDevices(env, { title, body, data, amount }) {
  if (!env.FCM_PROJECT_ID || !env.FCM_CLIENT_EMAIL || !env.FCM_PRIVATE_KEY) return;
  if (!env.QR_FIREBASE_URL || !env.QR_FIREBASE_CODE) return;

  const QR_BASE   = env.QR_FIREBASE_URL.replace(/\/+$/, '');
  const tokensUrl = `${QR_BASE}/fcm_tokens.json`;

  let tokenMap = {};
  try {
    const r = await fetch(tokensUrl);
    const parsed = await r.json();
    if (parsed && typeof parsed === 'object' && !parsed.error) tokenMap = parsed;
  } catch (e) {}

  const entries = Object.entries(tokenMap);
  if (!entries.length) return;

  const nowTs = now();
  let accessToken = null; // fetched lazily — only if at least one device is actually active

  for (const [key, rec] of entries) {
    if (!rec || !rec.token) continue;
    const lastSeen = Number(rec.last_seen || rec.registered_at || 0);
    const isAway   = (nowTs - lastSeen) > FCM_AWAY_AFTER_SEC;

    if (!isAway) {
      if (!accessToken) accessToken = await getFcmAccessToken(env);
      if (!accessToken) continue;
      await fcmSendToToken(env, accessToken, QR_BASE, key, rec.token, { title, body, data });
    } else {
      const newCount  = Number(rec.pending_count  || 0) + 1;
      const newAmount = Number(rec.pending_amount || 0) + Number(amount || 0);
      await fetch(`${QR_BASE}/fcm_tokens/${key}.json`, {
        method : 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body   : JSON.stringify({ pending_count: newCount, pending_amount: newAmount }),
      }).catch(() => {});
    }
  }
}

export default {
  async fetch(request, env, ctx) {

    const origin = request.headers.get('Origin') || '';


    // ── CORS ──
    const allowedOrigins = (env.ALLOWED_ORIGINS || '')
      .split(',').map(o => o.trim()).filter(Boolean);
    const allowedOrigin = allowedOrigins.includes(origin)
      ? origin : (allowedOrigins[0] || '*');
    const cors = {
      'Access-Control-Allow-Origin': allowedOrigin,
      'Access-Control-Allow-Methods': 'GET, POST, PATCH, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, x-razorpay-signature, X-App-Token',
      'Vary': 'Origin',
    };

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: cors });
    }

    const url  = new URL(request.url);
    const path = url.pathname;

    // ── Helpers ──
    const json = (data, status = 200) =>
      new Response(JSON.stringify(data), {
        status,
        headers: { ...cors, 'Content-Type': 'application/json' },
      });

    // ── 🔒 Shared-secret auth gate for admin-only endpoints ──
    const needsAuth = PROTECTED_PREFIXES.some(p => path.startsWith(p));
    if (needsAuth) {
      const suppliedToken = request.headers.get('X-App-Token') || '';
      if (!env.APP_TOKEN || suppliedToken !== env.APP_TOKEN) {
        return json({ error: 'Unauthorized — missing or invalid X-App-Token' }, 401);
      }
    }

    const rzpAuth = 'Basic ' + btoa(env.RZP_KEY_ID + ':' + env.RZP_KEY_SECRET);

    // ── Firebase — three SEPARATE projects, each addressed as URL + "code" path
    //    segment, no ?auth= secret. Access control lives entirely in your
    //    Firebase Realtime Database security RULES. ──

    // QR-map: {QR_FIREBASE_URL}/{QR_FIREBASE_CODE}/{id}  — flat, no qr_map/ wrapper
    const QR_BASE = (env.QR_FIREBASE_URL || '').replace(/\/+$/, '');
    const fbUrl = (node) => `${QR_BASE}/${env.QR_FIREBASE_CODE}/${node}.json`;

    const fbPut = (node, data) =>
      fetch(fbUrl(node), {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });

    const fbPatch = (node, data) =>
      fetch(fbUrl(node), {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });

    const fbGet = (node) => fetch(fbUrl(node));

    // Transactions: {TXN_FIREBASE_URL}/{TXN_FIREBASE_CODE}/{payment_id}  — flat, no
    // dl_transactions/ wrapper. Also now holds {order_id} directly — the
    // /create-order flow (website
    // search-and-pay) writes its order records straight into the TXN project
    // instead of the QR project, so that flow never touches QR at all.
    //
    // NOTE: the Cloudflare variable is actually named TXN_FIREBASE_SECRET on
    // this deployment (not TXN_FIREBASE_CODE, despite what the comments above
    // say) — resolve either name so this works regardless of which one is set.
    const TXN_BASE = (env.TXN_FIREBASE_URL || '').replace(/\/+$/, '');
    const TXN_CODE = env.TXN_FIREBASE_CODE || env.TXN_FIREBASE_SECRET || '';
    const txnFbUrl = (node) => `${TXN_BASE}/${TXN_CODE}/${node}.json`;

    const txnFbPut = (node, data) =>
      fetch(txnFbUrl(node), {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });

    const txnFbPatch = (node, data) =>
      fetch(txnFbUrl(node), {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });

    const txnFbGet = (node) => fetch(txnFbUrl(node));

    // Due List: {FIREBASE_URL}/DUELIST — flat array (0,1,2...), no wrapper node, no code needed
    const DUE_BASE = (env.DUE_FIREBASE_URL || '').replace(/\/+$/, '');
    const dueFbUrl = (query = '') => `${DUE_BASE}/DUELIST.json${query}`;

    // Student DB Firebase (separate project — master list of all students, used as a
    // fallback when a student has no due-list record at all — see /fee-lookup)
    const studentFbUrl = () =>
      `${env.STUDENT_FIREBASE_URL}/${env.STUDENT_FIREBASE_CODE}/students.json`;

    // ── Settle a Due List record after a captured payment ──
    //   Reads the CURRENT balance (not the amount that was charged — the
    //   two should match since accept_partial is false everywhere, but this
    //   stays correct even if the balance moved between order creation and
    //   payment), subtracts the amount just paid, and:
    //     • balance <= 0  → DELETE the record entirely from DUELIST
    //     • balance  > 0  → PATCH the reduced balance back
    //   Silently no-ops if dueKey/DUE_FIREBASE_URL is missing, or if the
    //   record is already gone (e.g. settled by a duplicate webhook fire).
    const settleDueBalance = async (dueKey, paidAmountPaise) => {
      if (!dueKey || !env.DUE_FIREBASE_URL) return;
      try {
        const r = await fetch(`${DUE_BASE}/DUELIST/${dueKey}.json`);
        if (!r.ok) {
          console.error('[settleDueBalance] GET failed', dueKey, r.status, await r.text().catch(() => ''));
          return;
        }
        const rec = await r.json();
        // rec can be null (already-gone record) OR a Firebase error object
        // like {"error":"Permission Denied"} if the RULES block this read —
        // either way, do NOT treat it as a real record with balance 0.
        if (!rec || typeof rec !== 'object' || rec.error) {
          if (rec && rec.error) console.error('[settleDueBalance] GET returned error', dueKey, rec.error);
          return;
        }
        const paidRupees = (Number(paidAmountPaise) || 0) / 100;
        const newBalance = Math.round(((Number(rec.balance) || 0) - paidRupees) * 100) / 100;
        if (newBalance <= 0) {
          const dr = await fetch(`${DUE_BASE}/DUELIST/${dueKey}.json`, { method: 'DELETE' });
          if (!dr.ok) console.error('[settleDueBalance] DELETE failed', dueKey, dr.status, await dr.text().catch(() => ''));
        } else {
          const pr = await fetch(`${DUE_BASE}/DUELIST/${dueKey}.json`, {
            method : 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body   : JSON.stringify({ balance: newBalance }),
          });
          if (!pr.ok) console.error('[settleDueBalance] PATCH failed', dueKey, pr.status, await pr.text().catch(() => ''));
        }
      } catch (e) {
        console.error('[settleDueBalance] exception', dueKey, e.message);
      }
    };

    // ── Find a Due List record + its Firebase key by admission number only
    //    (no mobile match required — used by /create-qr, which is an
    //    admin-only 🔒 endpoint, not a parent-facing one). Mirrors the
    //    adm-search logic in /create-order / /fee-lookup. ──
    const findDueKeyByAdm = async (admRaw) => {
      const norm   = (s) => String(s || '').trim().toLowerCase();
      const target = norm(admRaw);
      if (!target || !env.DUE_FIREBASE_URL) return { record: null, dueKey: null };

      try {
        const q = `&orderBy="adm"&equalTo="${encodeURIComponent(admRaw)}"`;
        const r = await fetch(dueFbUrl(q));
        const data = await r.json();
        if (data && typeof data === 'object' && !data.error) {
          const key = Object.keys(data)[0];
          if (key) return { record: data[key], dueKey: key };
        }
      } catch (e) {}

      try {
        const r = await fetch(dueFbUrl());
        const raw = await r.json();
        let list = [];
        if (Array.isArray(raw)) list = raw;
        else if (raw && typeof raw === 'object') list = Object.values(raw);
        const keys = Array.isArray(raw)
          ? raw.map((_, i) => String(i))
          : (raw && typeof raw === 'object' ? Object.keys(raw) : []);

        let idx = list.findIndex(x => norm(x.adm) === target);
        if (idx < 0) {
          idx = list.findIndex(x =>
            norm(x.adm).replace(/^0+/, '') === target.replace(/^0+/, '') &&
            norm(x.adm) !== ''
          );
        }
        if (idx >= 0) return { record: list[idx], dueKey: keys[idx] };
      } catch (e) {}

      return { record: null, dueKey: null };
    };

    // HMAC-SHA256 verify
    const verifySignature = async (body, signature, secret) => {
      const enc = new TextEncoder();
      const key = await crypto.subtle.importKey(
        'raw', enc.encode(secret),
        { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']
      );
      const sig = await crypto.subtle.sign('HMAC', key, enc.encode(body));
      const hex = Array.from(new Uint8Array(sig))
        .map(b => b.toString(16).padStart(2, '0')).join('');
      return hex === signature;
    };

    try {

      // ═══════════════════════════════════════════
      // POST /create-qr   🔒
      //   Creates a Razorpay UPI QR CODE (QR Codes API, type upi_qr),
      //   single-use + fixed amount, so it closes itself after one payment.
      //   The amount is locked on Razorpay's server the moment this runs —
      //   nothing about the amount ever depends on what the parent's
      //   browser sends later.
      //
      //   `short_url` (stored in Firebase, logged to the Sheet, and returned
      //   to the client) is now the UPI string from Razorpay's
      //   `image_content` field — upi://pay?pa=...&am=...&cu=INR...
      //   That field only exists in the response when the qr_image_content
      //   feature is enabled on the account. If it's missing we fall back to
      //   the rzp.io `image_url` so the field is never empty, and the
      //   response carries has_upi:false so the caller can tell.
      //
      //   Payment arrives via the `qr_code.credited` webhook (see /webhook).
      // ═══════════════════════════════════════════
      if (path === '/create-qr' && request.method === 'POST') {
        const body        = await request.json();
        const amountPaise = Math.round(Number(body.amount) * 100);

        if (!amountPaise || amountPaise < 100)
          return json({ error: 'amount must be at least ₹1' }, 400);

        const qrBody = {
          type          : 'upi_qr',
          name          : String(body.name || body.adm || 'Fee').slice(0, 40),
          usage         : 'single_use',
          fixed_amount  : true,
          payment_amount: amountPaise,
          description   : String(body.description || `Fee: ${body.name || body.adm || ''}`).slice(0, 200),
          notes         : {
            adm      : body.adm || '',
            name     : body.name || '',
            cls      : body.cls || '',
            reference: `${body.adm || 'stu'}-${now()}`,
          },
        };
        // Optional auto-close time (unix seconds). Razorpay needs it to be at
        // least 15 minutes in the future, so anything sooner is ignored.
        if (Number(body.close_by) > now() + 15 * 60) qrBody.close_by = Math.floor(Number(body.close_by));

        const rzpRes = await fetch(`${RZP_BASE}/payments/qr_codes`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'Authorization': rzpAuth },
          body: JSON.stringify(qrBody),
        });

        const rzpData = await rzpRes.json();
        if (!rzpRes.ok) return json({ error: rzpData }, rzpRes.status);

        // The UPI string that goes into the `short_url` field.
        const upiString = rzpData.image_content || '';
        const shortUrl  = upiString || rzpData.image_url || '';

        // Look up the matching Due List record so the qr_code.credited
        // webhook can settle (reduce / delete) it once this gets paid.
        let dueKey = null;
        if (body.adm) {
          try { ({ dueKey } = await findDueKeyByAdm(body.adm)); } catch (e) {}
        }

        await fbPut(`${rzpData.id}`, {
          adm        : body.adm  || '',
          name       : body.name || '',
          cls        : body.cls  || '',
          amount     : amountPaise,
          status     : 'pending',
          short_url  : shortUrl,
          image_url  : rzpData.image_url || '',
          due_key    : dueKey || '',
          created_at : now(),
        }).catch(() => {});

        gsLog(env, {
          type      : 'qr',
          qr_id     : rzpData.id,
          adm       : body.adm  || '',
          name      : body.name || '',
          cls       : body.cls  || '',
          amount    : amountPaise / 100,
          short_url : shortUrl,
          status    : rzpData.status,
          created_at: now(),
        });

        return json({
          qr_id         : rzpData.id,          // qr_xxx
          short_url     : shortUrl,            // upi://pay?... (falls back to rzp.io image_url if image_content isn't enabled)
          image_url     : rzpData.image_url || '',
          has_upi       : !!upiString,
          status        : rzpData.status,
          payment_amount: rzpData.payment_amount,
        });
      }

      // ═══════════════════════════════════════════
      // GET /fetch-payment/:qr_id  (unauthenticated — mirrors Razorpay poll)
      // ═══════════════════════════════════════════
      if (path.startsWith('/fetch-payment/') && request.method === 'GET') {
        const qrId = path.replace('/fetch-payment/', '').split('?')[0];
        if (!qrId) return json({ error: 'qr_id missing' }, 400);

        const rzpRes = await fetch(
          `${RZP_BASE}/payments/qr_codes/${qrId}/payments?count=1`,
          { headers: { 'Authorization': rzpAuth } }
        );
        const data = await rzpRes.json();
        if (!rzpRes.ok) return json({ error: data }, rzpRes.status);

        const items    = data.items || [];
        const captured = items.find(p => p.status === 'captured');

        return json({ captured: !!captured, payment: captured || null, count: items.length });
      }

      // ═══════════════════════════════════════════
      // GET /payment-status/:qr_id  (unauthenticated — parent's page polls this)
      //
      //   Fallback chain — each step only runs if the previous one found
      //   nothing:
      //     1. QR Firebase (qr_map/{qr_id})            — live, usual case
      //     2. TXN Firebase (dl_transactions, by qr_id) — QR record already
      //        gone (monthly cleanup / manual delete) but payment happened
      //     3. Google Sheet, TXN tab (historical, by qr_id)
      //     4. Google Sheet, QR tab — only to catch "closed, never paid";
      //        if found there with status "closed" → invalid
      //     5. Nothing found anywhere → invalid
      // ═══════════════════════════════════════════
      if (path.startsWith('/payment-status/') && request.method === 'GET') {
        const qrId = path.replace('/payment-status/', '').split('?')[0];
        if (!qrId) return json({ error: 'qr_id missing' }, 400);

        // 1. QR Firebase
        try {
          const fbRes  = await fbGet(`${qrId}`);
          const fbData = await fbRes.json();
          if (fbData) return json({ ...fbData, amount: (fbData.amount || 0) / 100, source: 'qr_node' });
        } catch (e) {}

        // 2. TXN Firebase (by qr_id)
        if (env.TXN_FIREBASE_URL && TXN_CODE) {
          try {
            const q = `?orderBy="qr_id"&equalTo="${encodeURIComponent(qrId)}"`;
            const r = await fetch(`${TXN_BASE}/${TXN_CODE}.json${q}`);
            const data = await r.json();
            if (data && typeof data === 'object' && Object.keys(data).length > 0) {
              const key = Object.keys(data)[0];
              const txn = data[key];
              return json({
                status     : 'captured',
                payment_id : txn.payment_id || key,
                adm        : txn.adm  || '',
                name       : txn.name || '',
                cls        : txn.cls  || '',
                amount     : (txn.amount || 0) / 100,
                method     : txn.method || '',
                vpa        : txn.vpa    || '',
                paid_at    : txn.paid_at || null,
                source     : 'txn_node',
              });
            }
          } catch (e) {}
        }

        // 3 & 4. Google Sheet — needs the doGet() lookup handler from
        //    sheets-lookup-addition.gs added to your Apps Script. If it's
        //    not there yet, gsLookup() just returns null and we fall
        //    straight through to step 5.
        const sheetHit = await gsLookup(env, qrId);

        // 3. Found a historical TXN row on the sheet
        if (sheetHit && sheetHit.sheet === 'TXN') {
          return json({
            status     : 'captured',
            payment_id : sheetHit.payment_id || '',
            adm        : sheetHit.adm  || '',
            name       : sheetHit.name || '',
            cls        : sheetHit.cls  || '',
            amount     : sheetHit.amount || 0,
            method     : sheetHit.method || '',
            vpa        : sheetHit.vpa    || '',
            paid_at    : sheetHit.paid_at || null,
            source     : 'sheet_txn',
          });
        }

        // 4. Found a QR row on the sheet, closed with no payment
        if (sheetHit && sheetHit.sheet === 'QR' && sheetHit.status === 'closed') {
          return json({ status: 'invalid', reason: 'closed_no_payment', source: 'sheet_qr' }, 404);
        }

        // 5. Nothing, anywhere
        return json({ status: 'invalid' }, 404);
      }

      // ═══════════════════════════════════════════
      // GET /txn-by-qr/:qr_id  (unauthenticated — same trust level as
      //   /payment-status: qr_id is unguessable. Kept as a standalone
      //   endpoint too, in case anything besides the payment page wants
      //   to check the TXN Firebase directly without the full fallback
      //   chain above.
      // ═══════════════════════════════════════════
      if (path.startsWith('/txn-by-qr/') && request.method === 'GET') {
        const qrId = path.replace('/txn-by-qr/', '').split('?')[0];
        if (!qrId) return json({ error: 'qr_id missing' }, 400);
        if (!env.TXN_FIREBASE_URL || !TXN_CODE)
          return json({ found: false });

        try {
          const q = `?orderBy="qr_id"&equalTo="${encodeURIComponent(qrId)}"`;
          const r = await fetch(`${TXN_BASE}/${TXN_CODE}.json${q}`);
          const data = await r.json();

          if (!data || typeof data !== 'object' || Object.keys(data).length === 0)
            return json({ found: false });

          const key = Object.keys(data)[0];
          return json({ found: true, ...data[key] });
        } catch (e) {
          return json({ found: false });
        }
      }

      // ═══════════════════════════════════════════
      // POST /close-qr/:qr_id   🔒   (closes the Razorpay QR Code; old plink_
      //   ids from before the QR Codes switch are still cancelled as links)
      // ═══════════════════════════════════════════
      if (path.startsWith('/close-qr/') && request.method === 'POST') {
        const qrId = path.replace('/close-qr/', '').split('?')[0];
        if (!qrId) return json({ error: 'qr_id missing' }, 400);

        const closeUrl = qrId.startsWith('plink_')
          ? `${RZP_BASE}/payment_links/${qrId}/cancel`
          : `${RZP_BASE}/payments/qr_codes/${qrId}/close`;
        const rzpRes = await fetch(
          closeUrl,
          { method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': rzpAuth } }
        );
        const data = await rzpRes.json();

        await fbPatch(`${qrId}`, { status: 'closed', closed_at: now() }).catch(() => {});

        gsLog(env, {
          type    : 'qr',
          qr_id   : qrId,
          status  : 'closed (manual)',
          closed_at: now(),
        });

        return json({
          closed: data.status === 'cancelled' || data.status === 'closed',
          status: data.status,
        }, rzpRes.ok ? 200 : rzpRes.status);
      }

      // ═══════════════════════════════════════════
      // POST /reconcile-qrs   🔒
      //   Fixes the gap where a student pays manually (cash etc.) and gets
      //   dropped from the Due List by the admin's "Upload xlsm & Full
      //   Replace" — their old pending QR used to stay open on Razorpay's
      //   side (and in Firebase) forever, until the monthly wipe.
      //
      //   Call this right after a Due List full-replace with
      //   { admList: [...every adm number in the NEW due list...] }.
      //   Any QR record in the QR Firebase project that is:
      //     - not already status "closed", AND
      //     - has an `adm` that is NOT in admList
      //   gets closed on Razorpay + its Firebase record deleted + logged to
      //   the Sheet, exactly like a manual /close-qr. QR records with no
      //   `adm` at all (shouldn't normally happen) are left untouched
      //   rather than guessed at.
      // ═══════════════════════════════════════════
      if (path === '/reconcile-qrs' && request.method === 'POST') {
        const body    = await request.json().catch(() => ({}));
        const admSet  = new Set((body.admList || []).map(a => String(a).trim()).filter(Boolean));

        let allQrs = {};
        try {
          const r = await fetch(`${QR_BASE}/${env.QR_FIREBASE_CODE}.json`);
          const parsed = await r.json();
          if (parsed && typeof parsed === 'object' && !parsed.error) allQrs = parsed;
        } catch (e) {}

        const closed = [];
        for (const [qrId, rec] of Object.entries(allQrs)) {
          if (!rec || typeof rec !== 'object') continue;
          if (rec.status === 'closed') continue;
          const adm = String(rec.adm || '').trim();
          if (!adm || admSet.has(adm)) continue; // no adm on record, or still due → leave it alone

          const closeUrl = qrId.startsWith('plink_')
            ? `${RZP_BASE}/payment_links/${qrId}/cancel`
            : `${RZP_BASE}/payments/qr_codes/${qrId}/close`;
          try {
            await fetch(closeUrl, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': rzpAuth } });
          } catch (e) {}

          await fetch(fbUrl(`${qrId}`), { method: 'DELETE' }).catch(() => {});
          gsLog(env, {
            type      : 'qr',
            qr_id     : qrId,
            adm       : adm,
            name      : rec.name || '',
            cls       : rec.cls  || '',
            amount    : (rec.amount || 0) / 100,
            status    : 'closed (due-list reconcile)',
            closed_at : now(),
          });
          closed.push({ qr_id: qrId, adm, name: rec.name || '' });
        }

        return json({ closed_count: closed.length, closed });
      }

      // ═══════════════════════════════════════════
      // POST /register-fcm-token   🔒
      //   Called once by reception.html when the admin clicks "Enable
      //   Notifications" and grants browser permission. Body: { token }
      //   (the FCM registration token from the Firebase Messaging SDK).
      //   Stored under fcm_tokens/{sanitized_token} in the QR Firebase
      //   project so sendFcmNotification() can find it on the next payment.
      //   Calling it again with the same token is harmless (same key).
      // ═══════════════════════════════════════════
      if (path === '/register-fcm-token' && request.method === 'POST') {
        const body  = await request.json().catch(() => ({}));
        const token = String(body.token || '').trim();
        if (!token) return json({ error: 'token missing' }, 400);
        if (!env.QR_FIREBASE_URL || !env.QR_FIREBASE_CODE) return json({ error: 'QR_FIREBASE_URL not configured on server' }, 500);

        // Firebase keys can't contain ".", "#", "$", "[", "]", "/" — FCM
        // tokens are base64url-ish already but colons/dots do appear.
        const key = token.replace(/[.#$\[\]/]/g, '_').slice(0, 400);
        const QR_BASE = env.QR_FIREBASE_URL.replace(/\/+$/, '');

        try {
          await fetch(`${QR_BASE}/fcm_tokens/${key}.json`, {
            method : 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body   : JSON.stringify({ token, registered_at: now(), last_seen: now() }),
          });
        } catch (e) {
          return json({ error: 'Failed to store token' }, 500);
        }

        return json({ registered: true });
      }

      // ═══════════════════════════════════════════
      // POST /fcm-heartbeat   🔒
      //   Called by reception.html on page load / visibility-change /
      //   window "online" event, for any device that already has a
      //   registered token. Body: { token }. Does two things:
      //     1. Bumps that device's last_seen to now — this is what keeps
      //        it in the "active, send real-time pushes" bucket in
      //        notifyPaymentToDevices() above.
      //     2. If it had queued pending_count/pending_amount from while it
      //        was away, sends ONE summary push right now ("🔔 X payments
      //        received while you were away — ₹Y total") and resets the
      //        counters to 0 — this is what turns an overnight laptop-off
      //        into a single catch-up notification instead of a burst of
      //        individual ones.
      // ═══════════════════════════════════════════
      if (path === '/fcm-heartbeat' && request.method === 'POST') {
        const body  = await request.json().catch(() => ({}));
        const token = String(body.token || '').trim();
        if (!token) return json({ error: 'token missing' }, 400);
        if (!env.QR_FIREBASE_URL || !env.QR_FIREBASE_CODE) return json({ error: 'QR_FIREBASE_URL not configured on server' }, 500);

        const key     = token.replace(/[.#$\[\]/]/g, '_').slice(0, 400);
        const QR_BASE = env.QR_FIREBASE_URL.replace(/\/+$/, '');
        const recUrl  = `${QR_BASE}/fcm_tokens/${key}.json`;

        let rec = null;
        try { rec = await (await fetch(recUrl)).json(); } catch (e) {}

        const pendingCount  = Number(rec?.pending_count  || 0);
        const pendingAmount = Number(rec?.pending_amount || 0);

        if (pendingCount > 0 && env.FCM_PROJECT_ID && env.FCM_CLIENT_EMAIL && env.FCM_PRIVATE_KEY) {
          const accessToken = await getFcmAccessToken(env);
          if (accessToken) {
            await fcmSendToToken(env, accessToken, QR_BASE, key, token, {
              title: '🔔 Welcome back',
              body : `${pendingCount} payment${pendingCount === 1 ? '' : 's'} received while you were away — ₹${pendingAmount.toLocaleString('en-IN')} total`,
              data : { type: 'summary', count: String(pendingCount), amount: String(pendingAmount) },
            });
          }
        }

        try {
          await fetch(recUrl, {
            method : 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body   : JSON.stringify({ last_seen: now(), pending_count: 0, pending_amount: 0 }),
          });
        } catch (e) {}

        return json({ ok: true, summary_sent: pendingCount > 0, pending_count: pendingCount, pending_amount: pendingAmount });
      }

      // ═══════════════════════════════════════════
      // POST /webhook  (protected by HMAC signature, not by APP_TOKEN)
      // ═══════════════════════════════════════════
      if (path === '/webhook' && request.method === 'POST') {
        const rawBody  = await request.text();
        const signature = request.headers.get('x-razorpay-signature') || '';

        const valid = await verifySignature(rawBody, signature, env.RZP_WEBHOOK_SECRET || '');
        if (!valid)
          return new Response(JSON.stringify({ error: 'Invalid signature' }), {
            status: 400, headers: { 'Content-Type': 'application/json' },
          });

        let payload;
        try { payload = JSON.parse(rawBody); }
        catch (e) {
          return new Response(JSON.stringify({ error: 'Invalid JSON' }), {
            status: 400, headers: { 'Content-Type': 'application/json' },
          });
        }

        const eventType = payload.event || '';

        // ── NEW: Payment Links flow (payment_link.paid) ──
        if (eventType === 'payment_link.paid') {
          const plink   = payload.payload?.payment_link?.entity || {};
          const payment = payload.payload?.payment?.entity || {};
          const plinkId = plink.id || '';

          if (plinkId) {
            // Auto-close the Razorpay link right after capture so it can never
            // be paid a second time.
            fetch(`${RZP_BASE}/payment_links/${plinkId}/cancel`, {
              method : 'POST',
              headers: { 'Content-Type': 'application/json', 'Authorization': rzpAuth },
            }).catch(() => {}); // Razorpay may reject cancelling an already-fully-paid
                                  // link — that's fine, it's already unusable for further
                                  // payment anyway (accept_partial was false).

            // Read the pending QR record first — this is still the ONLY
            // place with the student info (adm/name/cls) and the due_key.
            let studentInfo = { adm: '', name: '', cls: '', amount: plink.amount_paid || plink.amount || 0 };
            let dueKey = '';
            try {
              const qrRec = await (await fbGet(`${plinkId}`)).json();
              if (qrRec) {
                studentInfo.adm    = qrRec.adm    || '';
                studentInfo.name   = qrRec.name   || '';
                studentInfo.cls    = qrRec.cls    || '';
                studentInfo.amount = qrRec.amount || studentInfo.amount;
                dueKey             = qrRec.due_key || '';
              }
            } catch(e) {}

            const paymentId = (payment.id || plinkId).replace(/[^a-zA-Z0-9_-]/g, '_');
            if (paymentId && env.TXN_FIREBASE_URL && TXN_CODE) {
              // Awaited — this is the ONLY place the transaction record now
              // lives, so it must land before we clear the QR entry below.
              await txnFbPut(`${paymentId}`, {
                payment_id : payment.id || '',
                qr_id      : plinkId,
                adm        : studentInfo.adm,
                name       : studentInfo.name,
                cls        : studentInfo.cls,
                amount     : studentInfo.amount,
                method     : payment.method || '',
                vpa        : payment.vpa    || '',
                bank       : payment.bank   || '',
                rrn        : payment.acquirer_data?.rrn || '',
                ts         : now() * 1000,
                paid_at    : now(),
                source     : 'webhook',
              }).catch(() => {});
            }

            // Reduce (or delete, if fully paid) the matching Due List record.
            if (dueKey) await settleDueBalance(dueKey, studentInfo.amount);

            await gsLog(env, {
              type      : 'txn',
              payment_id: payment.id || plinkId,
              qr_id     : plinkId,
              adm       : studentInfo.adm,
              name      : studentInfo.name,
              cls       : studentInfo.cls,
              amount    : (studentInfo.amount || 0) / 100,
              method    : payment.method || '',
              vpa       : payment.vpa    || '',
              bank      : payment.bank   || '',
              paid_at   : now(),
            });

            ctx.waitUntil(notifyPaymentToDevices(env, {
              title : '💰 Payment Received',
              body  : `${studentInfo.name || studentInfo.adm} paid ₹${((studentInfo.amount || 0) / 100).toLocaleString('en-IN')}`,
              data  : { adm: studentInfo.adm, amount: String((studentInfo.amount || 0) / 100), type: 'payment' },
              amount: (studentInfo.amount || 0) / 100,
            }));

            // The QR node is only ever meant to hold PENDING qr data — now
            // that the full transaction lives in TXN, delete the QR
            // entry entirely instead of patching it with payment info.
            // /payment-status's fallback chain then finds it in TXN by
            // qr_id (source: "txn_node") the moment a poll lands.
            await fetch(fbUrl(`${plinkId}`), { method: 'DELETE' }).catch(() => {});
          }

          return new Response(JSON.stringify({ received: true }), {
            status: 200, headers: { 'Content-Type': 'application/json' },
          });
        }

        // ── UPI QR flow (QR Codes created by /create-qr) — qr_code.credited ──
        //   Does the same settlement the payment_link.paid branch above does:
        //   write the transaction to TXN, reduce/delete the Due List row, log
        //   to the Sheet, then delete the pending QR entry.
        //
        //   This must live BEFORE the payment.captured/Orders branch below:
        //   Razorpay auto-generates an order_* id for QR payments, so the
        //   Orders branch would otherwise swallow the event (no order_* record
        //   in TXN → it does nothing and returns).
        //
        //   Idempotent: the pending QR record in Firebase is the guard. A
        //   duplicate delivery finds it already deleted, then checks TXN for
        //   the pay_* row and skips.
        if (eventType === 'qr_code.credited') {
          const qrEntity  = payload.payload?.qr_code?.entity || {};
          const payment   = payload.payload?.payment?.entity || {};
          const qrCodeId  = qrEntity.id || payment.qr_code_id || '';

          if (qrCodeId) {
            const paymentId = (payment.id || qrCodeId).replace(/[^a-zA-Z0-9_-]/g, '_');

            let qrRec = null;
            try { qrRec = await (await fbGet(`${qrCodeId}`)).json(); } catch (e) {}

            let alreadyDone = false;
            if (!qrRec && env.TXN_FIREBASE_URL && TXN_CODE) {
              try { alreadyDone = !!(await (await txnFbGet(`${paymentId}`)).json()); } catch (e) {}
            }

            if (!alreadyDone) {
              const studentInfo = {
                adm   : (qrRec && qrRec.adm)  || '',
                name  : (qrRec && qrRec.name) || '',
                cls   : (qrRec && qrRec.cls)  || '',
                amount: (qrRec && qrRec.amount) || payment.amount || 0,
              };
              const dueKey = (qrRec && qrRec.due_key) || '';

              if (paymentId && env.TXN_FIREBASE_URL && TXN_CODE) {
                await txnFbPut(`${paymentId}`, {
                  payment_id : payment.id || '',
                  qr_id      : qrCodeId,
                  adm        : studentInfo.adm,
                  name       : studentInfo.name,
                  cls        : studentInfo.cls,
                  amount     : studentInfo.amount,
                  method     : payment.method || '',
                  vpa        : payment.vpa    || '',
                  bank       : payment.bank   || '',
                  rrn        : payment.acquirer_data?.rrn || '',
                  ts         : now() * 1000,
                  paid_at    : now(),
                  source     : 'webhook_qr',
                }).catch(() => {});
              }

              if (dueKey) await settleDueBalance(dueKey, studentInfo.amount);

              await gsLog(env, {
                type      : 'txn',
                payment_id: payment.id || qrCodeId,
                qr_id     : qrCodeId,
                adm       : studentInfo.adm,
                name      : studentInfo.name,
                cls       : studentInfo.cls,
                amount    : (studentInfo.amount || 0) / 100,
                method    : payment.method || '',
                vpa       : payment.vpa    || '',
                bank      : payment.bank   || '',
                paid_at   : now(),
              });

              ctx.waitUntil(notifyPaymentToDevices(env, {
                title : '💰 Payment Received',
                body  : `${studentInfo.name || studentInfo.adm} paid ₹${((studentInfo.amount || 0) / 100).toLocaleString('en-IN')}`,
                data  : { adm: studentInfo.adm, amount: String((studentInfo.amount || 0) / 100), type: 'payment' },
                amount: (studentInfo.amount || 0) / 100,
              }));

              // single_use QR closes itself on Razorpay's side after this
              // payment — only the pending Firebase entry needs clearing.
              await fetch(fbUrl(`${qrCodeId}`), { method: 'DELETE' }).catch(() => {});
            }
          }

          return new Response(JSON.stringify({ received: true }), {
            status: 200, headers: { 'Content-Type': 'application/json' },
          });
        }

        // ── NEW: Orders API flow (Custom Checkout on the website pay page)
        //         — backup for /verify-payment in case the browser closes
        //         before the Checkout `handler` callback can call it. ──
        if (eventType === 'payment.captured') {
          const orderPayEntity = payload.payload?.payment?.entity || {};
          const orderId        = orderPayEntity.order_id || '';

          if (orderId && orderId.startsWith('order_')) {
            const txnId = (orderPayEntity.id || orderId).replace(/[^a-zA-Z0-9_-]/g, '_');

            // ── Idempotency guard ──
            // The order_* record's own "status" field used to be the lock
            // against double-processing (webhook vs /verify-payment racing
            // each other) — but we now DELETE order_* right after a
            // successful run, so whichever path loses the race would see no
            // order_* record at all and treat this as a brand-new payment,
            // overwriting the correct pay_* row with blank student info and
            // double-subtracting the Due List balance. The pay_* record
            // itself is the only thing that survives, so check THAT first.
            let alreadyDone = false;
            if (txnId && env.TXN_FIREBASE_URL && TXN_CODE) {
              try { alreadyDone = !!(await (await txnFbGet(`${txnId}`)).json()); } catch (e) {}
            }

            if (!alreadyDone) {
              let orderRec = null;
              try { orderRec = await (await txnFbGet(`${orderId}`)).json(); } catch (e) {}

              // orderRec can legitimately be null here in a very tight race
              // (the other path read+wrote+deleted between our two awaits
              // above) — nothing to do in that case, the other path already
              // has it covered.
              if (orderRec) {
                // Awaited on purpose — see the note in /verify-payment above
                // about Workers killing in-flight fire-and-forget fetch()es.
                await txnFbPatch(`${orderId}`, {
                  status    : 'captured',
                  paid_at   : now(),
                  payment_id: orderPayEntity.id || '',
                }).catch(() => {});

                if (txnId && env.TXN_FIREBASE_URL && TXN_CODE) {
                  await txnFbPut(`${txnId}`, {
                    payment_id: orderPayEntity.id || '',
                    order_id  : orderId,
                    adm       : orderRec.adm  || '',
                    name      : orderRec.name || '',
                    cls       : orderRec.cls  || '',
                    amount    : orderRec.amount || orderPayEntity.amount || 0,
                    method    : orderPayEntity.method || '',
                    vpa       : orderPayEntity.vpa    || '',
                    bank      : orderPayEntity.bank   || '',
                    rrn       : orderPayEntity.acquirer_data?.rrn || '',
                    ts        : now() * 1000,
                    paid_at   : now(),
                    source    : 'webhook_order',
                  }).catch(() => {});
                }

                if (orderRec.due_key) {
                  await settleDueBalance(orderRec.due_key, orderRec.amount || orderPayEntity.amount || 0);
                }

                await gsLog(env, {
                  type      : 'txn',
                  payment_id: orderPayEntity.id || orderId,
                  order_id  : orderId,
                  adm       : orderRec.adm  || '',
                  name      : orderRec.name || '',
                  cls       : orderRec.cls  || '',
                  amount    : (orderRec.amount || 0) / 100,
                  method    : orderPayEntity.method || '',
                  vpa       : orderPayEntity.vpa    || '',
                  bank      : orderPayEntity.bank   || '',
                  paid_at   : now(),
                });

                ctx.waitUntil(notifyPaymentToDevices(env, {
                  title : '💰 Payment Received',
                  body  : `${orderRec.name || orderRec.adm} paid ₹${((orderRec.amount || 0) / 100).toLocaleString('en-IN')}`,
                  data  : { adm: orderRec.adm || '', amount: String((orderRec.amount || 0) / 100), type: 'payment' },
                  amount: (orderRec.amount || 0) / 100,
                }));

                // The order_* record's job is done — the pay_* record above
                // (txnId) is now the permanent one. Delete order_* so TXN only
                // ever holds one row per completed payment, not two.
                await fetch(txnFbUrl(orderId), { method: 'DELETE' }).catch(() => {});
              }
            }

            return new Response(JSON.stringify({ received: true }), {
              status: 200, headers: { 'Content-Type': 'application/json' },
            });
          }
        }

        // ── NEW: Orders API flow (Custom Checkout on the website pay page)
        //         — failed payments. Logged ONLY to the Google Sheet's
        //         "Failed" tab (not Firebase) — Firebase keeps the order_*
        //         record untouched so Razorpay can retry the same order_id
        //         without us having marked it "failed" there.
        // ══════════════════════════════════════════════════════════════
        if (eventType === 'payment.failed') {
          const orderPayEntity = payload.payload?.payment?.entity || {};
          const orderId        = orderPayEntity.order_id || '';
          const failPaymentId  = orderPayEntity.id || '';

          if (orderId && orderId.startsWith('order_')) {
            let orderRec = null;
            try { orderRec = await (await txnFbGet(`${orderId}`)).json(); } catch (e) {}

            if (orderRec) {
              // De-dupe guard — shared with /checkout-callback's failure
              // branch below via the same `failed_${key}` marker in the TXN
              // Firebase project. Without this, every Razorpay webhook
              // REDELIVERY of the exact same decline (which Razorpay does
              // when it doesn't get a fast 2xx) adds another identical row
              // to the Failed sheet — this was the cause of the same
              // order_id/payment_id appearing 5x in the Failed tab.
              //
              // Some declines never get a payment_id from Razorpay at all
              // (e.g. rejected before a payment object is created) — for
              // those, fall back to an order-level marker instead of
              // silently skipping the guard the way the old code did.
              const dedupeKey = failPaymentId ? `failed_${failPaymentId}` : `failed_order_${orderId}`;
              let alreadyLogged = false;
              try {
                const marker = await (await txnFbGet(dedupeKey)).json();
                alreadyLogged = !!marker;
              } catch (e) {}

              if (!alreadyLogged) {
                // Claim the slot FIRST, before the slow gsLog call — this
                // is the actual fix for the race above. A concurrent
                // duplicate delivery that lands while gsLog is still
                // running (often several seconds — Apps Script's exec
                // round-trip) used to still see the marker missing and log
                // a second row. Claiming here shrinks that window from
                // "however long gsLog takes" down to just this PUT's own
                // round trip (normally well under a second).
                await txnFbPut(dedupeKey, { logged_at: now(), order_id: orderId }).catch(() => {});
                await gsLog(env, {
                  type              : 'txn_failed',
                  payment_id        : failPaymentId,
                  order_id          : orderId,
                  adm               : orderRec.adm  || '',
                  name              : orderRec.name || '',
                  cls               : orderRec.cls  || '',
                  amount            : (orderRec.amount || orderPayEntity.amount || 0) / 100,
                  method            : orderPayEntity.method            || '',
                  error_code        : orderPayEntity.error_code        || '',
                  error_description : orderPayEntity.error_description || '',
                  failed_at         : now(),
                });
              } else {
                console.log('[webhook] payment.failed: duplicate delivery for already-logged decline, skipping re-log:', JSON.stringify({ orderId, paymentId: failPaymentId }));
              }
            }

            return new Response(JSON.stringify({ received: true }), {
              status: 200, headers: { 'Content-Type': 'application/json' },
            });
          }
        }

        // ── OLD: QR-code flow (kept for backward compatibility with any
        //         QR codes created before this update, or the separate
        //         Fee QR feature if it's ever wired up to this worker) ──
        const entity    = payload.payload?.payment?.entity || {};
        const qrId      = entity.qr_code_id || '';

        if (qrId && eventType === 'payment.captured') {

          fbPatch(`${qrId}`, {
            status     : 'captured',
            paid_at    : now(),
            payment_id : entity.id     || '',
            method     : entity.method || '',
            vpa        : entity.vpa    || '',
          }).catch(() => {});

          let studentInfo = { adm: '', name: '', cls: '', amount: entity.amount || 0 };
          try {
            const qrRec = await (await fbGet(`${qrId}`)).json();
            if (qrRec) {
              studentInfo.adm    = qrRec.adm    || '';
              studentInfo.name   = qrRec.name   || '';
              studentInfo.cls    = qrRec.cls    || '';
              studentInfo.amount = qrRec.amount || entity.amount || 0;
            }
          } catch(e) {}

          const paymentId = (entity.id || '').replace(/[^a-zA-Z0-9_-]/g, '_');
          if (paymentId && env.TXN_FIREBASE_URL && TXN_CODE) {
            txnFbPut(`${paymentId}`, {
              payment_id : entity.id         || '',
              qr_id      : qrId,
              adm        : studentInfo.adm,
              name       : studentInfo.name,
              cls        : studentInfo.cls,
              amount     : studentInfo.amount,
              method     : entity.method     || '',
              vpa        : entity.vpa        || '',
              bank       : entity.bank       || '',
              rrn        : entity.acquirer_data?.rrn || '',
              ts         : now() * 1000,
              paid_at    : now(),
              source     : 'webhook',
            }).catch(() => {});
          }

          gsLog(env, {
            type      : 'txn',
            payment_id: entity.id || qrId,
            qr_id     : qrId,
            adm       : studentInfo.adm,
            name      : studentInfo.name,
            cls       : studentInfo.cls,
            amount    : (studentInfo.amount || 0) / 100,
            method    : entity.method || '',
            vpa       : entity.vpa    || '',
            bank      : entity.bank   || '',
            paid_at   : now(),
          });

        } else if (qrId && eventType === 'payment.failed') {
          // Logged ONLY to the Google Sheet's "Failed" tab — Firebase's
          // qr_<id> record is left untouched (no status patch here), so
          // the QR link's live Firebase state doesn't change on failure.
          let failedStudentInfo = { adm: '', name: '', cls: '', amount: entity.amount || 0 };
          try {
            const qrRec = await (await fbGet(`${qrId}`)).json();
            if (qrRec) {
              failedStudentInfo.adm    = qrRec.adm    || '';
              failedStudentInfo.name   = qrRec.name   || '';
              failedStudentInfo.cls    = qrRec.cls    || '';
              failedStudentInfo.amount = qrRec.amount || entity.amount || 0;
            }
          } catch(e) {}

          gsLog(env, {
            type              : 'txn_failed',
            payment_id        : entity.id || '',
            qr_id             : qrId,
            adm               : failedStudentInfo.adm,
            name              : failedStudentInfo.name,
            cls               : failedStudentInfo.cls,
            amount            : (failedStudentInfo.amount || 0) / 100,
            method            : entity.method            || '',
            error_code        : entity.error_code        || '',
            error_description : entity.error_description || '',
            failed_at         : now(),
          });

        } else if (qrId) {
          fbPatch(`${qrId}`, {
            status     : 'unknown',
            event_type : eventType,
            logged_at  : now(),
          }).catch(() => {});
        }

        return new Response(JSON.stringify({ received: true }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }

      // ═══════════════════════════════════════════
      // GET /qr-by-adm/:adm   🔒
      // ═══════════════════════════════════════════
      if (path.startsWith('/qr-by-adm/') && request.method === 'GET') {
        const adm = decodeURIComponent(path.replace('/qr-by-adm/', '').split('?')[0]);
        if (!adm) return json({ error: 'adm missing' }, 400);

        const fbRes  = await fetch(
          `${QR_BASE}/${env.QR_FIREBASE_CODE}.json?orderBy="adm"&equalTo="${adm}"`
        );
        const data = await fbRes.json();

        if (!data || typeof data !== 'object' || Object.keys(data).length === 0)
          return json({ found: false });

        const qrId = Object.keys(data)[0];
        return json({ found: true, qr_id: qrId, ...data[qrId] });
      }

      // ═══════════════════════════════════════════
      // DELETE /delete-qr/:qr_id   🔒
      // ═══════════════════════════════════════════
      if (path.startsWith('/delete-qr/') && request.method === 'DELETE') {
        const qrId = path.replace('/delete-qr/', '').split('?')[0];
        if (!qrId) return json({ error: 'qr_id missing' }, 400);

        fetch(
          qrId.startsWith('plink_')
            ? `${RZP_BASE}/payment_links/${qrId}/cancel`
            : `${RZP_BASE}/payments/qr_codes/${qrId}/close`,
          {
            method : 'POST',
            headers: { 'Content-Type': 'application/json', 'Authorization': rzpAuth },
          }
        ).catch(() => {});

        await fetch(fbUrl(`${qrId}`), { method: 'DELETE' }).catch(() => {});

        return json({ deleted: true, qr_id: qrId });
      }

      // ═══════════════════════════════════════════
      // POST /create-order  (public-safe — same adm+mobile dual-match design
      //   as /fee-lookup below; requires BOTH to match a due-list record).
      //
      //   Creates a Razorpay ORDER (Orders API) for Custom Checkout, opened
      //   in-page on the website pay page via checkout.js — NOT a Payment
      //   Link. The WhatsApp flow keeps using /create-qr (UPI QR Codes API),
      //   untouched.
      //
      //   SECURITY: the amount is NEVER taken from the client. It is always
      //   the live balance read from the Due List at the moment the order
      //   is created, so nothing about what gets charged ever depends on
      //   what the parent's browser sends.
      //
      //   The order is saved directly to {<order_id>} in the TXN Firebase
      //   project (NOT the QR project — this flow never creates a QR
      //   entry at all) so /verify-payment and the webhook fallback can
      //   find the student + the exact Due List row to settle later.
      // ═══════════════════════════════════════════
      if (path === '/create-order' && request.method === 'POST') {
        let body;
        try { body = await request.json(); }
        catch (e) { return json({ error: 'Invalid JSON body' }, 400); }

        const admRaw    = String(body.adm    || '').trim();
        const mobileRaw = String(body.mobile || '').trim();

        if (!admRaw)    return json({ error: 'adm missing' }, 400);
        if (!mobileRaw) return json({ error: 'mobile missing' }, 400);
        if (!env.DUE_FIREBASE_URL) return json({ error: 'FIREBASE_URL not configured on server' }, 500);

        const norm       = (s) => String(s || '').trim().toLowerCase();
        const digitsOnly = (s) => String(s || '').replace(/\D/g, '');
        const target      = norm(admRaw);
        const targetMob   = digitsOnly(mobileRaw);
        const MOBILE_FIELD = 'mob';

        const mobileMatches = (record, field) => {
          const stored = digitsOnly(record[field]);
          if (!stored || !targetMob) return false;
          const a = stored.slice(-10), b = targetMob.slice(-10);
          return a.length === 10 && a === b;
        };

        // ── Find the due-list record AND its Firebase key (so /verify-payment
        //    can patch its balance to 0 once the payment is confirmed) ──
        let record = null, dueKey = null;
        try {
          const q = `&orderBy="adm"&equalTo="${encodeURIComponent(admRaw)}"`;
          const r = await fetch(dueFbUrl(q));
          const data = await r.json();
          if (data && typeof data === 'object' && !data.error) {
            const key = Object.keys(data)[0];
            if (key) { record = data[key]; dueKey = key; }
          }
        } catch (e) {}

        if (!record) {
          try {
            const r = await fetch(dueFbUrl());
            const raw = await r.json();
            let list = [];
            if (Array.isArray(raw)) list = raw;
            else if (raw && typeof raw === 'object') list = Object.values(raw);
            const keys = Array.isArray(raw)
              ? raw.map((_, i) => String(i))
              : (raw && typeof raw === 'object' ? Object.keys(raw) : []);

            let idx = list.findIndex(x => norm(x.adm) === target);
            if (idx < 0) {
              idx = list.findIndex(x =>
                norm(x.adm).replace(/^0+/, '') === target.replace(/^0+/, '') &&
                norm(x.adm) !== ''
              );
            }
            if (idx < 0) {
              idx = list.findIndex(x => norm(x.adm) === '' && mobileMatches(x, MOBILE_FIELD));
            }
            if (idx >= 0) { record = list[idx]; dueKey = keys[idx]; }
          } catch (e) {}
        }

        if (!record)
          return json({ error: 'No due-list record found for this admission number' }, 404);
        if (!mobileMatches(record, MOBILE_FIELD))
          return json({ error: 'Mobile number does not match our records' }, 403);

        const balance = Number(record.balance) || 0;
        if (balance <= 0)
          return json({ error: 'No pending fee — nothing to pay' }, 400);

        const amountPaise = Math.round(balance * 100);
        if (amountPaise < 100)
          return json({ error: 'Amount too small to process (minimum ₹1)' }, 400);

        const orderBody = {
          amount  : amountPaise,
          currency: 'INR',
          receipt : `${admRaw}-${now()}`.slice(0, 40),
          notes   : { adm: record.adm || admRaw, name: record.name || '', cls: record.cls || '', due_key: dueKey || '' },
        };

        const rzpRes = await fetch(`${RZP_BASE}/orders`, {
          method : 'POST',
          headers: { 'Content-Type': 'application/json', 'Authorization': rzpAuth },
          body   : JSON.stringify(orderBody),
        });
        const rzpData = await rzpRes.json();
        if (!rzpRes.ok) return json({ error: rzpData }, rzpRes.status);

        // NOT awaited — ctx.waitUntil() keeps this Worker invocation alive
        // to finish the write in the background, but the Response below no
        // longer waits on it. This was the #1 cause of the "Proceed to Pay"
        // → Razorpay redirect feeling slow: a synchronous Firebase PUT.
        ctx.waitUntil(txnFbPut(`${rzpData.id}`, {
          adm       : record.adm  || admRaw,
          name      : record.name || '',
          cls       : record.cls  || '',
          mobile    : targetMob,
          due_key   : dueKey || '',
          amount    : amountPaise,
          status    : 'created',
          created_at: now(),
        }).catch(() => {}));

        // NOT awaited either — same reasoning. Google Apps Script's exec
        // round-trip is usually the single slowest step in this whole
        // request (often 1–3s, sometimes more on a cold start), and it was
        // being awaited right before the redirect. ctx.waitUntil() still
        // guarantees it completes (this is what the old comment about
        // Workers killing un-awaited fetch() was warning about — waitUntil
        // is the actual fix for that, not the await that was here before),
        // it just no longer blocks the parent's browser.
        ctx.waitUntil(gsLog(env, {
          type      : 'order',
          order_id  : rzpData.id,
          adm       : record.adm  || admRaw,
          name      : record.name || '',
          cls       : record.cls  || '',
          amount    : amountPaise / 100,
          status    : rzpData.status,
          created_at: now(),
        }));

        return json({
          order_id: rzpData.id,
          amount  : rzpData.amount,
          currency: rzpData.currency,
          key_id  : env.RZP_KEY_ID,
          name    : record.name || '',
          cls     : record.cls  || '',
          adm     : record.adm  || admRaw,
          school  : env.SCHOOL_NAME || 'Lotus International School',
        });
      }

      // ═══════════════════════════════════════════
      // POST /verify-payment  (public — protected by its own HMAC signature
      //   check, same trust model as /webhook. Called by the website pay
      //   page's Checkout `handler`, right after Razorpay reports success in
      //   the browser. The /webhook payment.captured branch (order_id based,
      //   see below) is a backup in case this call never fires — e.g. the
      //   parent closes the tab before the handler runs.
      // ═══════════════════════════════════════════
      if (path === '/verify-payment' && request.method === 'POST') {
        let body;
        try { body = await request.json(); }
        catch (e) { return json({ error: 'Invalid JSON body' }, 400); }

        const orderId   = String(body.razorpay_order_id   || '').trim();
        const paymentId = String(body.razorpay_payment_id || '').trim();
        const signature = String(body.razorpay_signature  || '').trim();

        if (!orderId || !paymentId || !signature)
          return json({ error: 'Missing razorpay_order_id / razorpay_payment_id / razorpay_signature' }, 400);

        const valid = await verifySignature(`${orderId}|${paymentId}`, signature, env.RZP_KEY_SECRET || '');

        // AUDIT LOG — request in. Kept separate from the signature check so
        // even a failed/invalid signature attempt shows up in the trail.
        console.log('[verify-payment] request:', JSON.stringify({
          razorpay_order_id  : orderId,
          razorpay_payment_id: paymentId,
          signature_valid    : valid,
        }));

        if (!valid) return json({ verified: false, error: 'Signature mismatch' }, 400);

        const txnId = paymentId.replace(/[^a-zA-Z0-9_-]/g, '_');

        // ── Dual inquiry — calls Razorpay's Status API (GET /payments/:id)
        // and logs the request/response. This is the HDFC-mandated
        // independent server-to-server check: we do NOT take the browser's
        // word (or even just our own signature check) for it — we ask
        // Razorpay directly what the payment's current state is, and only
        // "captured" counts as successful (see FAQ: "Consider only the
        // captured state to mark the payment successful"). Called once,
        // up front, and reused both for the gating decision below AND for
        // the audit-trail log / payment-method lookup. ──
        const auditStatusCheck = async () => {
          try {
            const pr = await fetch(`${RZP_BASE}/payments/${paymentId}`, { headers: { 'Authorization': rzpAuth } });
            const pd = await pr.json().catch(() => null);
            console.log('[verify-payment] razorpay status response:', JSON.stringify({
              http_status: pr.status,
              payment_id : paymentId,
              data       : pd,
            }));
            return (pr.ok && pd) ? pd : null;
          } catch (e) {
            console.log('[verify-payment] razorpay status response: fetch failed', e.message);
            return null;
          }
        };
        const pd = await auditStatusCheck();
        const isCaptured = !!(pd && pd.status === 'captured');
        let payMethod = pd ? (pd.method || '') : '';
        let payVpa    = pd ? (pd.vpa    || '') : '';
        let payBank   = pd ? (pd.bank   || '') : '';
        let payRrn    = pd ? (pd.acquirer_data?.rrn || '') : '';

        // ── Idempotency guard — see the matching note in the webhook's
        //    payment.captured (order-based) branch. Check the pay_* record
        //    FIRST, before ever touching order_*: order_* gets deleted right
        //    after a successful run, so if the webhook already finished by
        //    the time this call lands, order_* is simply gone — checking its
        //    status first would find nothing and re-run everything from
        //    scratch, overwriting the correct pay_* row with blanks and
        //    double-subtracting the Due List balance. ──
        let existingTxn = null;
        if (txnId && env.TXN_FIREBASE_URL && TXN_CODE) {
          try { existingTxn = await (await txnFbGet(`${txnId}`)).json(); } catch (e) {}
        }
        if (existingTxn) {
          return json({ verified: true, already_processed: true, adm: existingTxn.adm, name: existingTxn.name });
        }

        // Signature valid → payment is genuine. Look up the order we saved
        // in /create-order to know which student / Due List row this is.
        let orderRec = null;
        try { orderRec = await (await txnFbGet(`${orderId}`)).json(); } catch (e) {}

        // Already processed (handler fired AND webhook also fired) — avoid
        // double-crediting / duplicate txn rows. (Also covers the very tight
        // race where order_* was patched to "captured" but the webhook
        // hadn't written pay_*/deleted order_* yet when our check above ran.)
        if (orderRec && orderRec.status === 'captured') {
          return json({ verified: true, already_processed: true, adm: orderRec.adm, name: orderRec.name });
        }

        // order_* already gone AND no pay_* record either — the webhook must
        // have deleted it between our two checks above (extremely tight
        // race). Nothing left for us to safely do; the webhook already has
        // (or is about to have) this covered.
        if (!orderRec) {
          return json({ verified: true, note: 'handled by webhook' });
        }

        // ── HDFC-mandated gate: do NOT settle, write the permanent pay_*
        // record, or mark the order captured unless the dual-inquiry Status
        // API independently confirmed "captured". Signature validity alone
        // proves the payment is genuine, but not that it has actually
        // captured — a manual-capture method or a transient Status API
        // hiccup could still leave it "authorized"/unknown at this instant.
        // In that case we leave order_* untouched and let the /webhook
        // payment.captured branch (which has its own strict, independent
        // "captured" check) settle it the moment Razorpay confirms it —
        // exactly the backup role that branch already documents itself as.
        if (!isCaptured) {
          console.log('[verify-payment] dual-inquiry did not confirm captured — deferring settlement to webhook:', JSON.stringify({
            orderId, paymentId, status: pd ? pd.status : 'status_check_failed',
          }));
          return json({
            verified: true,
            pending : true,
            adm     : orderRec.adm  || '',
            name    : orderRec.name || '',
            note    : 'Payment not yet confirmed captured; will finalize automatically once Razorpay confirms.',
          });
        }

        await txnFbPatch(`${orderId}`, {
          status    : 'captured',
          paid_at   : now(),
          payment_id: paymentId,
        }).catch(() => {});

        const studentInfo = {
          adm   : orderRec?.adm    || '',
          name  : orderRec?.name   || '',
          cls   : orderRec?.cls    || '',
          amount: orderRec?.amount || 0,
        };

        // NOTE: everything below is awaited on purpose (not fire-and-forget).
        // Cloudflare Workers can kill background fetch()es that are still
        // in flight once the Response has been sent back to the browser —
        // if we return before these finish, the balance/log writes can
        // silently never happen even though the client sees "verified".

        if (txnId && env.TXN_FIREBASE_URL && TXN_CODE) {
          await txnFbPut(`${txnId}`, {
            payment_id: paymentId,
            order_id  : orderId,
            adm       : studentInfo.adm,
            name      : studentInfo.name,
            cls       : studentInfo.cls,
            amount    : studentInfo.amount,
            method    : payMethod || 'checkout',
            vpa       : payVpa,
            bank      : payBank,
            rrn       : payRrn,
            ts        : now() * 1000,
            paid_at   : now(),
            source    : 'verify-payment',
          }).catch(() => {});
        }

        // Reduce (or delete, if fully paid) the Due List balance for this student.
        if (orderRec && orderRec.due_key) {
          await settleDueBalance(orderRec.due_key, studentInfo.amount);
        }

        await gsLog(env, {
          type      : 'txn',
          payment_id: paymentId,
          order_id  : orderId,
          adm       : studentInfo.adm,
          name      : studentInfo.name,
          cls       : studentInfo.cls,
          amount    : (studentInfo.amount || 0) / 100,
          method    : payMethod || 'checkout',
          vpa       : payVpa,
          bank      : payBank,
          paid_at   : now(),
        });

        ctx.waitUntil(notifyPaymentToDevices(env, {
          title : '💰 Payment Received',
          body  : `${studentInfo.name || studentInfo.adm} paid ₹${((studentInfo.amount || 0) / 100).toLocaleString('en-IN')}`,
          data  : { adm: studentInfo.adm, amount: String((studentInfo.amount || 0) / 100), type: 'payment' },
          amount: (studentInfo.amount || 0) / 100,
        }));

        // The order_* record's job is done — the pay_* record above (txnId)
        // is now the permanent one. Delete order_* so TXN only ever holds
        // one row per completed payment, not two.
        await fetch(txnFbUrl(orderId), { method: 'DELETE' }).catch(() => {});

        return json({ verified: true, adm: studentInfo.adm, name: studentInfo.name });
      }

      // ═══════════════════════════════════════════
      // POST /checkout-callback  (public — Razorpay HOSTED CHECKOUT posts
      //   here directly via a full-page browser redirect, NOT a fetch() from
      //   our own JS. This is the callback_url the website's payNow() puts
      //   into the hidden <form> it submits to
      //   https://api.razorpay.com/v1/checkout/embedded.
      //
      //   Success: Razorpay POSTs razorpay_order_id / razorpay_payment_id /
      //   razorpay_signature as form-data.
      //   Failure/decline: Razorpay POSTs error[code] / error[description] /
      //   error[metadata][order_id] / error[metadata][payment_id] instead —
      //   no razorpay_signature is present in that case.
      //
      //   Either way this does the SAME settlement work /verify-payment used
      //   to do (idempotency-guarded, so a /webhook payment.captured race is
      //   still safe), then 303-redirects the browser back to the pay page
      //   with the result in the query string, since the page's own JS state
      //   was lost during the redirect round-trip to Razorpay's hosted page.
      // ═══════════════════════════════════════════
      if (path === '/checkout-callback' && request.method === 'POST') {
        // Resolve siteBase defensively: a misconfigured env var (missing
        // "https://", a stray path segment, a typo, etc.) must NOT be able
        // to break the redirect for every single checkout. Try each
        // candidate in priority order and only accept ones that actually
        // parse as an absolute http(s) URL; log + skip anything that
        // doesn't, so a bad value shows up in the logs instead of silently
        // 500-ing on every payment.
        const HARDCODED_FALLBACK = 'https://fee.lisdnn.org/lsid';
        const siteBaseCandidates = [
          env.SITE_URL,
          allowedOrigins.find(o => o.includes('fee.')),
          allowedOrigins[0],
          HARDCODED_FALLBACK,
        ];

        const asValidOrigin = (candidate) => {
          if (!candidate) return null;
          const trimmed = String(candidate).trim();
          // Response.redirect() needs a scheme; a bare host like
          // "fee.lisdnn.org" (or "fee.lisdnn.org/lsid") is not one.
          const withScheme = /^https?:\/\//i.test(trimmed) ? trimmed : `https://${trimmed}`;
          try {
            const u = new URL(withScheme);
            if (u.protocol !== 'http:' && u.protocol !== 'https:') return null;
            // Keep host AND path (e.g. "/lsid") — only query/hash are
            // dropped, since redirectTo() appends its own query string.
            const path = u.pathname === '/' ? '' : u.pathname.replace(/\/+$/, '');
            return `${u.protocol}//${u.host}${path}`;
          } catch (e) {
            return null;
          }
        };

        let siteBase = null;
        for (const candidate of siteBaseCandidates) {
          const resolved = asValidOrigin(candidate);
          if (resolved) {
            if (resolved !== String(candidate).trim().replace(/\/+$/, '')) {
              console.log('[checkout-callback] siteBase candidate accepted with normalization:', JSON.stringify({ raw: candidate, resolved }));
            }
            siteBase = resolved;
            break;
          }
          if (candidate) {
            console.log('[checkout-callback] siteBase candidate rejected (not a valid URL):', JSON.stringify({ raw: candidate }));
          }
        }
        if (!siteBase) siteBase = HARDCODED_FALLBACK; // last resort, should be unreachable

        const redirectTo = (params) => {
          const q = new URLSearchParams(params);
          const sep = siteBase.includes('?') ? '&' : '?';
          return Response.redirect(`${siteBase}${sep}${q.toString()}`, 303);
        };

        let form;
        try { form = await request.formData(); }
        catch (e) {
          return redirectTo({ status: 'failed', reason: 'Malformed response from payment gateway.' });
        }

        // Log every field Razorpay actually posted. This is the only way to
        // know the real key names for a given failure type without
        // guessing — if the lookups below ever come up empty again, this
        // log line has the ground truth.
        const rawFormFields = {};
        for (const [k, v] of form.entries()) rawFormFields[k] = v;
        console.log('[checkout-callback] raw form fields:', JSON.stringify(rawFormFields));

        // Try known exact key names first; if none hit, fall back to
        // scanning ALL posted keys for one that looks like the field we
        // want. Razorpay's exact key naming for failed/declined attempts
        // isn't 100% consistent across failure types, so this is a safety
        // net rather than the primary path.
        const pick = (...keys) => {
          for (const k of keys) {
            const v = form.get(k);
            if (v) return String(v).trim();
          }
          return '';
        };
        const pickByPattern = (pattern) => {
          for (const [k, v] of form.entries()) {
            if (pattern.test(k) && v) return String(v).trim();
          }
          return '';
        };

        // order_id fallback chain: Razorpay's own posted fields first, then
        // (as of the index.html change that bakes order_id into callback_url
        // as a query param) our OWN query string — this is the reliable
        // path, since Razorpay does not consistently echo order_id in
        // error[metadata][...] for every decline type. We set this
        // ourselves before ever redirecting to Razorpay, so it's always
        // there regardless of what Razorpay chooses to send back.
        const urlOrderId = new URL(request.url).searchParams.get('order_id') || '';
        const orderId = pick('razorpay_order_id', 'error[metadata][order_id]', 'error[metadata][order_id][]')
          || pickByPattern(/order[_-]?id/i)
          || urlOrderId;
        const paymentId = pick('razorpay_payment_id', 'error[metadata][payment_id]', 'error[metadata][payment_id][]')
          || pickByPattern(/payment[_-]?id/i);
        const signature = String(form.get('razorpay_signature') || '').trim();

        // ── FAILURE branch — no signature means Razorpay posted an error
        //    payload, not a successful payment. Log it to the Sheet's
        //    "Failed" tab, same as the webhook's payment.failed branch does,
        //    then send the parent back with a readable reason. ──
        if (!signature || !orderId || !paymentId) {
          const errorCode = pick('error[code]') || pickByPattern(/error.*code/i);
          const errorDesc = pick('error[description]')
            || pickByPattern(/error.*desc/i)
            || 'The transaction could not be completed.';

          if (orderId) {
            let orderRec = null;
            try { orderRec = await (await txnFbGet(`${orderId}`)).json(); } catch (e) {}
            if (orderRec) {
              // De-dupe guard: the SAME failed callback (identical
              // payment_id) can land on this endpoint more than once — a
              // slow response, a browser "resubmit form" on refresh, etc.
              // Without this, every duplicate POST adds another identical
              // row to the Failed sheet for the exact same decline.
              //
              // Shared with the /webhook payment.failed branch above via the
              // same `failed_${key}` marker scheme, so whichever path (this
              // callback or the server-to-server webhook) fires FIRST for a
              // given decline "wins" and the other one sees the marker and
              // skips — previously each path only checked its own history,
              // so the same decline landing on both still produced 2 rows.
              //
              // When Razorpay doesn't hand back a payment_id at all (some
              // decline types never get one), fall back to an order-level
              // marker instead of skipping the guard entirely — the old
              // code's `if (paymentId)` left this case completely unguarded.
              const dedupeKey = paymentId ? `failed_${paymentId}` : `failed_order_${orderId}`;
              let alreadyLogged = false;
              try {
                const marker = await (await txnFbGet(dedupeKey)).json();
                alreadyLogged = !!marker;
              } catch (e) {}

              if (!alreadyLogged) {
                // Same fix as the webhook branch above — claim first,
                // slow gsLog second, so a duplicate landing on THIS
                // endpoint (browser resubmit) or on the webhook branch
                // (they share this exact dedupeKey) doesn't race in
                // during the multi-second gsLog call.
                await txnFbPut(dedupeKey, { logged_at: now(), order_id: orderId }).catch(() => {});
                await gsLog(env, {
                  type              : 'txn_failed',
                  payment_id        : paymentId,
                  order_id          : orderId,
                  adm               : orderRec.adm  || '',
                  name              : orderRec.name || '',
                  cls               : orderRec.cls  || '',
                  amount            : (orderRec.amount || 0) / 100,
                  error_code        : errorCode,
                  error_description : errorDesc,
                  failed_at         : now(),
                });
              } else {
                console.log('[checkout-callback] failure branch: duplicate callback for already-logged decline, skipping re-log:', JSON.stringify({ orderId, paymentId }));
              }
            } else {
              console.log('[checkout-callback] failure branch: orderId present but no matching Firebase record:', JSON.stringify({ orderId }));
            }
          } else {
            console.log('[checkout-callback] failure branch: could not determine orderId from posted form fields or URL query — see raw form fields logged above.');
          }

          return redirectTo({ status: 'failed', order_id: orderId, reason: errorDesc });
        }

        // ── SUCCESS branch — verify signature exactly like /verify-payment
        //    does: hmac_sha256(order_id + "|" + payment_id, key_secret). ──
        const valid = await verifySignature(`${orderId}|${paymentId}`, signature, env.RZP_KEY_SECRET || '');

        console.log('[checkout-callback] request:', JSON.stringify({
          razorpay_order_id  : orderId,
          razorpay_payment_id: paymentId,
          signature_valid    : valid,
        }));

        if (!valid) {
          return redirectTo({
            status  : 'failed',
            order_id: orderId,
            reason  : 'Payment could not be verified. Please contact the school office if any amount was debited.',
          });
        }

        const txnId = paymentId.replace(/[^a-zA-Z0-9_-]/g, '_');

        // ── Dual inquiry (HDFC-mandated) — see the matching comment in
        // /verify-payment. Called once, up front, reused for the gating
        // decision below, the audit log, AND the payment-method lookup. ──
        const auditStatusCheck = async () => {
          try {
            const pr = await fetch(`${RZP_BASE}/payments/${paymentId}`, { headers: { 'Authorization': rzpAuth } });
            const pd = await pr.json().catch(() => null);
            console.log('[checkout-callback] razorpay status response:', JSON.stringify({
              http_status: pr.status,
              payment_id : paymentId,
              data       : pd,
            }));
            return (pr.ok && pd) ? pd : null;
          } catch (e) {
            console.log('[checkout-callback] razorpay status response: fetch failed', e.message);
            return null;
          }
        };
        const pd = await auditStatusCheck();
        const isCaptured = !!(pd && pd.status === 'captured');
        let payMethod = pd ? (pd.method || '') : '';
        let payVpa    = pd ? (pd.vpa    || '') : '';
        let payBank   = pd ? (pd.bank   || '') : '';
        let payRrn    = pd ? (pd.acquirer_data?.rrn || '') : '';

        // ── Idempotency guard — identical shape to /verify-payment's, so a
        //    /webhook payment.captured race resolves the same safe way. ──
        let existingTxn = null;
        if (txnId && env.TXN_FIREBASE_URL && TXN_CODE) {
          try { existingTxn = await (await txnFbGet(`${txnId}`)).json(); } catch (e) {}
        }
        if (existingTxn) {
          return redirectTo({
            status    : 'success',
            order_id  : orderId,
            payment_id: paymentId,
            adm       : existingTxn.adm  || '',
            name      : existingTxn.name || '',
            cls       : existingTxn.cls  || '',
            amount    : String(Math.round((existingTxn.amount || 0) / 100)),
          });
        }

        let orderRec = null;
        try { orderRec = await (await txnFbGet(`${orderId}`)).json(); } catch (e) {}

        if (orderRec && orderRec.status === 'captured') {
          return redirectTo({
            status    : 'success',
            order_id  : orderId,
            payment_id: paymentId,
            adm       : orderRec.adm  || '',
            name      : orderRec.name || '',
            cls       : orderRec.cls  || '',
            amount    : String(Math.round((orderRec.amount || 0) / 100)),
          });
        }

        if (!orderRec) {
          // The webhook must have already fully processed and cleared this
          // order in a tight race — payment is genuine (signature checked
          // out), just nothing left here to read student details back from.
          return redirectTo({ status: 'success', order_id: orderId, payment_id: paymentId });
        }

        // ── HDFC-mandated gate: do NOT settle or write the permanent pay_*
        // record unless the dual-inquiry Status API independently confirmed
        // "captured". Signature validity alone proves the payment is
        // genuine, not that it has actually captured. If not yet confirmed,
        // leave order_* untouched — the /webhook payment.captured branch
        // (its own strict, independent check) will settle it the moment
        // Razorpay confirms, exactly the backup role it already documents. ──
        if (!isCaptured) {
          console.log('[checkout-callback] dual-inquiry did not confirm captured — deferring settlement to webhook:', JSON.stringify({
            orderId, paymentId, status: pd ? pd.status : 'status_check_failed',
          }));
          return redirectTo({
            status    : 'success',
            order_id  : orderId,
            payment_id: paymentId,
            adm       : orderRec.adm  || '',
            name      : orderRec.name || '',
            cls       : orderRec.cls  || '',
            amount    : String(Math.round((orderRec.amount || 0) / 100)),
          });
        }

        try {
          await txnFbPatch(`${orderId}`, {
            status    : 'captured',
            paid_at   : now(),
            payment_id: paymentId,
          });
        } catch (e) {
          // .catch() alone doesn't help here — txnFbPatch is a plain
          // (non-async) arrow fn that calls fetch() directly, so a
          // SYNCHRONOUS throw from fetch() (e.g. bad URL) happens before
          // any promise is returned for .catch() to attach to, and used to
          // escape straight to the outer try/catch → bare 500, no redirect.
          console.error('[checkout-callback] txnFbPatch threw:', e.message);
        }

        const studentInfo = {
          adm   : orderRec?.adm    || '',
          name  : orderRec?.name   || '',
          cls   : orderRec?.cls    || '',
          amount: orderRec?.amount || 0,
        };

        if (txnId && env.TXN_FIREBASE_URL && TXN_CODE) {
          try {
            await txnFbPut(`${txnId}`, {
              payment_id: paymentId,
              order_id  : orderId,
              adm       : studentInfo.adm,
              name      : studentInfo.name,
              cls       : studentInfo.cls,
              amount    : studentInfo.amount,
              method    : payMethod || 'checkout',
              vpa       : payVpa,
              bank      : payBank,
              rrn       : payRrn,
              ts        : now() * 1000,
              paid_at   : now(),
              source    : 'checkout-callback',
            });
          } catch (e) {
            console.error('[checkout-callback] txnFbPut threw:', e.message);
          }
        }

        if (orderRec && orderRec.due_key) {
          await settleDueBalance(orderRec.due_key, studentInfo.amount);
        }

        try {
          await gsLog(env, {
            type      : 'txn',
            payment_id: paymentId,
            order_id  : orderId,
            adm       : studentInfo.adm,
            name      : studentInfo.name,
            cls       : studentInfo.cls,
            amount    : (studentInfo.amount || 0) / 100,
            method    : payMethod || 'checkout',
            vpa       : payVpa,
            bank      : payBank,
            paid_at   : now(),
          });
        } catch (e) {
          console.error('[checkout-callback] gsLog threw:', e.message);
        }

        ctx.waitUntil(notifyPaymentToDevices(env, {
          title : '💰 Payment Received',
          body  : `${studentInfo.name || studentInfo.adm} paid ₹${((studentInfo.amount || 0) / 100).toLocaleString('en-IN')}`,
          data  : { adm: studentInfo.adm, amount: String((studentInfo.amount || 0) / 100), type: 'payment' },
          amount: (studentInfo.amount || 0) / 100,
        }));

        // Same cleanup /verify-payment does — order_* has done its job now
        // that the pay_* record (txnId) holds the permanent copy.
        try {
          await fetch(txnFbUrl(orderId), { method: 'DELETE' });
        } catch (e) {
          console.error('[checkout-callback] TXN order DELETE threw:', e.message);
        }

        return redirectTo({
          status    : 'success',
          order_id  : orderId,
          payment_id: paymentId,
          adm       : studentInfo.adm,
          name      : studentInfo.name,
          cls       : studentInfo.cls,
          amount    : String(Math.round((studentInfo.amount || 0) / 100)),
        });
      }

      // ═══════════════════════════════════════════
      // POST /fee-lookup  (public-safe, protected by its own adm+mobile design)
      //
      //   Search order:
      //     1. Due List (dl_due_data) — a real due-list record exists.
      //        → mobile must match → returns actual balance.
      //     2. NOT in Due List at all — falls back to the master Student DB.
      //        → mobile must match there too → returns balance: 0 ("paid up"),
      //          because a student with no due-list entry has nothing pending.
      //     3. Not found in either → { found: false }, same as before.
      // ═══════════════════════════════════════════
      if (path === '/fee-lookup' && request.method === 'POST') {
        let body;
        try { body = await request.json(); }
        catch (e) { return json({ error: 'Invalid JSON body' }, 400); }

        const admRaw    = String(body.adm    || '').trim();
        const mobileRaw = String(body.mobile || '').trim();

        if (!admRaw)    return json({ error: 'adm missing' }, 400);
        if (!mobileRaw) return json({ error: 'mobile missing' }, 400);

        if (!env.DUE_FIREBASE_URL) {
          return json({ error: 'FIREBASE_URL not configured on server' }, 500);
        }

        const norm      = (s) => String(s || '').trim().toLowerCase();
        const digitsOnly = (s) => String(s || '').replace(/\D/g, '');
        const target     = norm(admRaw);
        const targetMob  = digitsOnly(mobileRaw);

        const MOBILE_FIELD = 'mob';

        const mobileMatches = (record, field) => {
          const stored = digitsOnly(record[field]);
          if (!stored || !targetMob) return false;
          const a = stored.slice(-10);
          const b = targetMob.slice(-10);
          return a.length === 10 && a === b;
        };

        // ── Step 1: search the Due List ──
        let record = null;
        try {
          const q = `&orderBy="adm"&equalTo="${encodeURIComponent(admRaw)}"`;
          const r = await fetch(dueFbUrl(q));
          const data = await r.json();
          if (data && typeof data === 'object' && !data.error) {
            const key = Object.keys(data)[0];
            if (key) record = data[key];
          }
        } catch (e) {}

        if (!record) {
          try {
            const r = await fetch(dueFbUrl());
            const raw = await r.json();
            let list = [];
            if (Array.isArray(raw)) list = raw;
            else if (raw && typeof raw === 'object') list = Object.values(raw);

            record = list.find(x => norm(x.adm) === target);
            if (!record) {
              record = list.find(x =>
                norm(x.adm).replace(/^0+/, '') === target.replace(/^0+/, '') &&
                norm(x.adm) !== ''
              );
            }
            // Fallback: some Due List rows have no usable "adm" field at all
            // (e.g. row imported without it — this is what was happening for
            // adm 10003 / Vihaan Mahajan: balance, mob, name etc. were present
            // but "adm" was missing, so every match above silently failed).
            // Only applies to rows where adm is blank, so it can never override
            // a genuine adm mismatch on a row that does have one.
            if (!record) {
              record = list.find(x => norm(x.adm) === '' && mobileMatches(x, MOBILE_FIELD));
            }
          } catch (e) {}
        }

        if (record) {
          // Found in due list — mobile must match, real balance returned.
          if (!mobileMatches(record, MOBILE_FIELD)) {
            return json(body.debug
              ? { found: false, debug: 'mobile_mismatch', storedMobDigits: digitsOnly(record[MOBILE_FIELD]).slice(-10), targetMobDigits: targetMob.slice(-10) }
              : { found: false });
          }
          return json({
            found  : true,
            name   : record.name    || '',
            adm    : record.adm     || admRaw,
            cls    : record.cls     || '',
            balance: Number(record.balance) || 0,
          });
        }

        // ── Step 2: no due-list record at all — fall back to master Student DB ──
        if (env.STUDENT_FIREBASE_URL && env.STUDENT_FIREBASE_CODE) {
          try {
            const sr = await fetch(studentFbUrl());
            const sd = await sr.json();
            const list = Array.isArray(sd) ? sd : (sd ? Object.values(sd) : []);

            let stu = list.find(x => norm(x.admNo || x.admissionNumber) === target);
            if (!stu) {
              stu = list.find(x =>
                norm(x.admNo || x.admissionNumber).replace(/^0+/, '') === target.replace(/^0+/, '') &&
                norm(x.admNo || x.admissionNumber) !== ''
              );
            }

            if (stu && mobileMatches(stu, 'phone')) {
              return json({
                found  : true,
                name   : stu.name  || '',
                adm    : stu.admNo || stu.admissionNumber || admRaw,
                cls    : stu.cls   || stu.class || '',
                balance: 0,   // no due-list entry → nothing pending
              });
            }

            if (stu && body.debug) {
              return json({ found: false, debug: 'student_db_mobile_mismatch' });
            }
          } catch (e) {}
        }

        // ── Step 3: not found anywhere ──
        return json(body.debug ? { found: false, debug: 'no_record_matched_adm', admRaw } : { found: false });
      }

      return json({ error: 'Not found' }, 404);

    } catch (err) {
      return json({ error: err.message || 'Internal error' }, 500);
    }
  },
};

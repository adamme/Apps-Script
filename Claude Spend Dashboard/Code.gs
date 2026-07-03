

/**
 * Claude Enterprise Spend Tracker
 * =============================================================================
 * One Code.gs + one Dashboard.html. Pulls spend from the Anthropic
 * Analytics API daily and serves a dashboard either as a Sheet menu
 * dialog or as a standalone web app.
 *
 * Setup (one-time):
 *   1. Open a Google Sheet → Extensions → Apps Script
 *   2. Replace Code.gs with this file; add Dashboard.html
 *   3. Project Settings → Script Properties → add ANALYTICS_API_KEY (your Claude Analytics API key)
 *   4. (For web app) Project Settings → Script Properties → Add SPREADSHEET_ID to Script Properties
 *   5. Run refreshFullHistory once from the editor (authorize scopes)
 *   6. Add a time-based trigger: refreshDaily, every day at 6am
 *   7. (For web app) Deploy → New deployment → Web app
 *
 * Reference:
 * https://support.claude.com/en/articles/13703965-claude-enterprise-analytics-api-reference-guide
 */

// =============================================================================
// CONFIG
// =============================================================================

const CONFIG = {
  API_BASE: 'https://api.anthropic.com/v1/organizations/analytics',

  // The actual cost endpoints (confirmed from the help-center reference):
  COST_REPORT_PATH: '/cost_report',           // bucketed time series
  USER_COST_REPORT_PATH: '/user_cost_report', // per-user cost

  // Sheet tabs:
  SHEET_DAILY: 'daily_cost',
  SHEET_USERS: 'user_cost',
  SHEET_MONTHLY: 'monthly_summary',
  SHEET_PRODUCT_MONTH: 'product_month',
  SHEET_META: '_meta',

  // Gemini (Google Workspace) adoption — usage, NOT cost. Gemini in
  // Workspace is per-seat licensed, so there's no per-user dollar figure
  // to pull; we surface adoption instead. Sourced from the Admin SDK
  // Reports Activities API (applicationName=gemini_in_workspace_apps).
  SHEET_GEMINI: 'gemini_usage',
  GEMINI_APP_NAME: 'gemini_in_workspace_apps',
  GEMINI_WINDOW_DAYS: 30,            // rolling window recomputed each refresh (one row per user)

  // User-maintained sheet (read-only from the script's POV): one row
  // per person with org metadata used to roll spend up to department.
  // Expected columns: Email, Display Name, Title, Department, Function, Manager.
  SHEET_USERS_DIRECTORY: 'Users',

  // Refresh windows:
  BACKFILL_DAYS: 35,         // daily trigger re-pulls this many days
  INITIAL_HISTORY_DAYS: 365, // backfill goes back this far

  PAGE_LIMIT: 1000,
  TIMEZONE: 'America/New_York',

  // Per-user cost breakdown is fetched by calling /user_cost_report once
  // per product (the endpoint doesn't support group_by). This list is the
  // canonical set of product slugs accepted by the products[] filter —
  // taken verbatim from the API's own validation error. If a new product
  // launches it will need to be added here.
  PRODUCTS_FOR_USER_BREAKDOWN: [
    'chat', 'claude_code', 'code_review', 'cc_security_autopatch',
    'cowork', 'office_agent', 'claude_in_chrome', 'claude_design',
    'claude_in_slack', 'voice_mode', 'research', 'other'
  ],

  // The Analytics API only has data from this date forward. Requests for
  // chunks entirely before this date 400 with "data prior to X is not
  // available". The refresh code clamps to this floor automatically.
  API_DATA_FLOOR: '2026-01-01',
};

const DAILY_HEADERS = ['date', 'product', 'model', 'cost_usd', 'last_updated_utc', 'row_key'];
const USER_HEADERS = ['period_start', 'period_end', 'user_email', 'user_id', 'user_name', 'product', 'cost_usd', 'last_updated_utc', 'row_key'];
const MONTHLY_HEADERS = ['month', 'total_cost_usd', 'updated_at'];
const PRODUCT_MONTH_HEADERS = ['month', 'product', 'cost_usd', 'updated_at'];
const META_HEADERS = ['key', 'value', 'updated_at'];

// ONE ROW PER USER over a rolling GEMINI_WINDOW_DAYS window. Recomputed from
// scratch each refresh and fully replaced, so size is fixed at ~#users and the
// numbers track current adoption (users who fall out of the window drop off).
// No app dimension (the API's per-app attribution is unreliable).
const GEMINI_HEADERS = ['user_email', 'events', 'active_days', 'first_date', 'last_date', 'last_updated_utc'];

function getApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('ANALYTICS_API_KEY');
  if (!key) {
    throw new Error('No ANALYTICS_API_KEY in Script Properties. Add it under Project Settings → Script Properties.');
  }
  return key;
}

/**
 * Web-app-safe spreadsheet accessor.
 *
 * When the dashboard is opened via the menu (modal dialog), there's an
 * active spreadsheet and getActiveSpreadsheet() works. When the
 * dashboard is served as a web app via /exec, there is no active
 * spreadsheet and we need to open by ID. Set SPREADSHEET_ID in Script
 * Properties for web-app deployments.
 */
function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet and no SPREADSHEET_ID in Script Properties. Add the spreadsheet ID (from the sheet URL between /d/ and /edit) to Project Settings → Script Properties.');
  }
  return ss;
}

// =============================================================================
// MENU + DASHBOARD ENTRYPOINTS
// =============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Claude Spend')
    .addItem('Open dashboard', 'showDashboard')
    .addSeparator()
    .addItem('Refresh now (last 35 days)', 'refreshDaily')
    .addItem('Backfill full history', 'refreshFullHistory')
    .addSeparator()
    .addItem('Refresh Gemini usage', 'refreshGemini')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Diagnostics')
      .addItem('Reconcile check (validate totals)', 'reconcileCheck')
      .addItem('Debug dump one API row', 'debugDumpOneRow')
      .addItem('Clean up archived tabs', 'cleanupArchivedTabs'))
    .addToUi();
}

function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Claude Enterprise Spend')
    .setWidth(1400)
    .setHeight(900);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Claude Enterprise Spend');
}

/**
 * Web app entry point. After deploying as a web app, this function
 * serves the dashboard at the /exec URL. The dashboard reads from the
 * sheet identified by SPREADSHEET_ID in Script Properties.
 *
 * To deploy:
 *   1. Set SPREADSHEET_ID in Script Properties (paste the ID from the
 *      Sheet URL — the long string between /d/ and /edit).
 *   2. Apps Script editor → Deploy → New deployment → Type: Web app.
 *   3. Execute as: Me. Who has access: Anyone within Company.
 *   4. Copy the /exec URL and share it.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Claude Enterprise Spend')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================================================================
// HTTP CLIENT
// =============================================================================

function fetchCostOverTime_(startingAt, endingAt, opts) {
  opts = opts || {};
  const bucketWidth = opts.bucketWidth || '1d';
  // The cost_report endpoint caps `limit` at 31 when bucket_width=1d
  // (since the bucketed series can't exceed 31 days per request anyway).
  const limit = bucketWidth === '1d' ? 31 : CONFIG.PAGE_LIMIT;
  const params = {
    starting_at: toIsoTimestamp_(startingAt),
    ending_at: toIsoTimestamp_(endingAt),
    bucket_width: bucketWidth,
    limit: limit,
  };
  (opts.groupBy || ['product', 'model']).forEach(function (g) { addListParam_(params, 'group_by', g); });
  return paginate_(CONFIG.COST_REPORT_PATH, params);
}

/**
 * Per-user cost over a date range. The /user_cost_report endpoint does
 * NOT support group_by — it always returns one row per user with a
 * single total. To get a per-product breakdown, the caller must invoke
 * this once per product using the `products[]` filter, then aggregate.
 */
function fetchCostUsers_(startingAt, endingAt, opts) {
  opts = opts || {};
  const params = {
    starting_at: toIsoTimestamp_(startingAt),
    ending_at: toIsoTimestamp_(endingAt),
    limit: CONFIG.PAGE_LIMIT,
  };
  if (opts.products && opts.products.length) {
    opts.products.forEach(function (p) { addListParam_(params, 'products', p); });
  }
  if (opts.orderBy) params.order_by = opts.orderBy;
  return paginate_(CONFIG.USER_COST_REPORT_PATH, params);
}

function paginate_(path, params) {
  const all = [];
  let cursor = null;
  for (let pageCount = 0; pageCount < 100; pageCount++) {
    const req = Object.assign({}, params);
    if (cursor) req.page = cursor;
    const resp = requestJson_(path, req);
    (resp.data || []).forEach(function (r) { all.push(r); });
    if (!resp.has_more || !resp.next_page) break;
    cursor = resp.next_page;
  }
  return all;
}

function requestJson_(path, params) {
  const url = CONFIG.API_BASE + path + '?' + encodeParams_(params);
  const options = {
    method: 'get',
    headers: { 'x-api-key': getApiKey_(), 'anthropic-version': '2023-06-01' },
    muteHttpExceptions: true,
  };

  let lastErr = null;
  for (let attempt = 1; attempt <= 5; attempt++) {
    let resp;
    try {
      resp = UrlFetchApp.fetch(url, options);
    } catch (e) {
      lastErr = e;
      Utilities.sleep(Math.pow(2, attempt - 1) * 1000);
      continue;
    }
    const code = resp.getResponseCode();
    const body = resp.getContentText();
    if (attempt === 1) Logger.log('GET ' + url + ' -> ' + code);

    if (code >= 200 && code < 300) {
      try { return JSON.parse(body); }
      catch (e) { throw new Error('Non-JSON from ' + path + ': ' + body.slice(0, 200)); }
    }
    if (code === 429 || code >= 500) {
      lastErr = new Error('HTTP ' + code + ': ' + body.slice(0, 300));
      Utilities.sleep(Math.pow(2, attempt - 1) * 1000);
      continue;
    }
    // Gracefully handle the "data not available yet" 400 — happens at
    // the API's historical floor or right at the data_refreshed_at
    // watermark. Treat as empty rather than failing the whole refresh.
    if (code === 400 && body.indexOf('not available') !== -1) {
      Logger.log('Skipping unavailable window for ' + path + ': ' + body.slice(0, 200));
      return { data: [], has_more: false };
    }
    // 4xx: don't retry, surface full context
    throw new Error('HTTP ' + code + ' from ' + path + '\nURL: ' + url + '\nBody: ' + body.slice(0, 800));
  }
  throw lastErr;
}

function encodeParams_(params) {
  const parts = [];
  Object.keys(params).forEach(function (k) {
    const v = params[k];
    if (Array.isArray(v)) v.forEach(function (i) { parts.push(encodeURIComponent(k) + '=' + encodeURIComponent(i)); });
    else if (v !== null && v !== undefined) parts.push(encodeURIComponent(k) + '=' + encodeURIComponent(v));
  });
  return parts.join('&');
}

function addListParam_(params, key, value) {
  const k = key + '[]';
  if (!Array.isArray(params[k])) params[k] = [];
  params[k].push(value);
}

// =============================================================================
// RESPONSE FLATTENERS
//
// The API returns amounts as DECIMAL STRINGS IN CENTS, e.g. "41280.000000" =
// $412.80. We parse and divide by 100.
//
// Bucketed cost_report response shape:
//   { data: [ { starting_at: "...", results: [ { product, model, amount } ] } ] }
//
// Per-user user_cost_report response shape:
//   { data: [ { actor: { user_id, name, email }, results: [ { product, amount } ] } ] }
// =============================================================================

function flattenCostOverTime_(buckets) {
  const out = [];
  const nowIso = new Date().toISOString();
  buckets.forEach(function (b) {
    const date = String(b.starting_at || b.date || '').slice(0, 10);
    if (!date) return;
    (b.results || []).forEach(function (r) {
      const product = r.product || 'unknown';
      const model = r.model || 'unknown';
      const cost = centsStringToUsd_(r.amount);
      const key = date + '|' + product + '|' + model;
      out.push([date, product, model, cost, nowIso, key]);
    });
  });
  return out;
}

/**
 * Flatten /user_cost_report response into user_cost rows.
 * Every API response row is one user with a single total (the endpoint
 * doesn't break down per product). The caller is responsible for
 * filtering by product upstream and passing the product name here so
 * we tag rows correctly.
 */
function flattenCostUsers_(users, periodStart, periodEnd, product) {
  const out = [];
  const nowIso = new Date().toISOString();
  const monthKey = String(periodStart).slice(0, 7); // YYYY-MM
  users.forEach(function (u) {
    const actor = u.actor || {};
    const email = actor.email || u.email || '';
    const uid = actor.user_id || u.user_id || '';
    const name = actor.name || u.name || '';
    const total = centsStringToUsd_(u.amount);
    if (!email || total <= 0) return;  // skip empty/zero rows
    // Key on calendar month (not raw window dates) so re-pulling a month
    // overwrites the same rows instead of creating duplicates.
    const key = monthKey + '|' + email + '|' + product;
    out.push([periodStart, periodEnd, email, uid, name, product, total, nowIso, key]);
  });
  return out;
}

/**
 * "41280.000000" → 412.80
 * Anthropic returns amounts as decimal strings in cents.
 */
function centsStringToUsd_(amount) {
  if (amount === null || amount === undefined) return 0;
  const n = typeof amount === 'string' ? parseFloat(amount) : Number(amount);
  if (!isFinite(n)) return 0;
  return Math.round(n) / 100;
}

// =============================================================================
// REFRESH ORCHESTRATION
// =============================================================================

function refreshDaily() {
  runRefresh_(CONFIG.BACKFILL_DAYS, 'daily');
  // Also refresh Gemini adoption. Wrapped so a Gemini failure (e.g. the
  // running account lacks the Reports privilege) never blocks cost data.
  try { refreshGemini(); } catch (e) { Logger.log('Gemini refresh skipped: ' + e); }
}
function refreshFullHistory() { runRefresh_(CONFIG.INITIAL_HISTORY_DAYS, 'full'); }

function runRefresh_(daysBack, kind) {
  // Stale-status guard. If a previous refresh got killed (timeout, force-
  // close, etc.) before it could update meta, _meta keeps saying "running"
  // forever. Detect a stale running status (>15 min old) and treat as crashed.
  const lastStatus = getMeta_('last_refresh_status');
  if (lastStatus === 'running') {
    const lastStarted = getMeta_('last_refresh_kind_started_at');
    const ageMin = lastStarted ? (Date.now() - new Date(lastStarted).getTime()) / 60000 : 999;
    if (ageMin > 15) {
      Logger.log('Detected stale "running" status (' + ageMin.toFixed(1) + ' min old); recovering.');
      setMeta_('last_refresh_status', 'crashed');
    }
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.log('Another refresh in progress, skipping');
    return;
  }
  const startedAt = new Date();
  try {
    setMeta_('last_refresh_status', 'running');
    setMeta_('last_refresh_kind', kind);
    setMeta_('last_refresh_kind_started_at', startedAt.toISOString());

    const ending = isoDateUtc_(addDays_(new Date(), -1));
    let starting = isoDateUtc_(addDays_(new Date(), -daysBack));
    // Clamp to the API's earliest-available date so we don't 400 on
    // pre-history chunks.
    if (starting < CONFIG.API_DATA_FLOOR) {
      Logger.log('Clamping start date: ' + starting + ' → ' + CONFIG.API_DATA_FLOOR + ' (API data floor)');
      starting = CONFIG.API_DATA_FLOOR;
    }
    if (starting > ending) {
      Logger.log('Nothing to refresh: starting=' + starting + ' is after ending=' + ending);
      setMeta_('last_refresh_status', 'noop');
      return;
    }
    // Build the list of calendar months the refresh window touches. We
    // iterate by MONTH (not arbitrary 31-day chunks) so per-user rows get
    // stable keys (YYYY-MM|email|product) and re-pulling a month OVERWRITES
    // rather than duplicating. This is what prevents the double-counting
    // that happens when different refreshes use different window boundaries.
    const months = monthsInRange_(starting, ending);

    let dailyTotal = { updated: 0, inserted: 0 };
    let userTotal = { updated: 0, inserted: 0 };

    // One-time migration: archive then clear the legacy user_cost sheet if
    // it still contains rows keyed by arbitrary date windows. We detect
    // this by checking whether any row_key lacks the new YYYY-MM month form.
    migrateUserCostIfNeeded_();

    months.forEach(function (mo) {
      // mo = { key:'2026-05', start:'2026-05-01', end:'2026-05-31' }
      // Clamp the per-month window to [API floor, yesterday] so the first
      // and current months don't request unavailable dates.
      const moStart = mo.start < starting ? starting : mo.start;
      const moEnd = mo.end > ending ? ending : mo.end;

      // Bucketed daily — daily endpoint caps at 31-day requests, and a
      // single calendar month is always <= 31 days, so one call covers it.
      const overTimeRaw = fetchCostOverTime_(moStart, moEnd, { groupBy: ['product', 'model'] });
      const dailyRows = flattenCostOverTime_(overTimeRaw);
      const r1 = upsertByKey_(CONFIG.SHEET_DAILY, DAILY_HEADERS, dailyRows);
      dailyTotal.updated += r1.updated;
      dailyTotal.inserted += r1.inserted;

      // Before writing this month's per-user rows, delete any existing
      // rows for this month. Combined with stable month-keyed rows, this
      // guarantees a clean overwrite even if product slugs changed between
      // runs (e.g. a product that had spend last run but not this run).
      pruneUserCostByMonth_(mo.key);

      // Per-user — the endpoint doesn't support group_by, so we call it
      // once per product using the products[] filter, then aggregate
      // client-side. Rows are stamped with calendar-month boundaries.
      CONFIG.PRODUCTS_FOR_USER_BREAKDOWN.forEach(function (product) {
        const userRaw = fetchCostUsers_(moStart, moEnd, { products: [product] });
        const userRows = flattenCostUsers_(userRaw, mo.start, mo.end, product);
        const r2 = upsertByKey_(CONFIG.SHEET_USERS, USER_HEADERS, userRows);
        userTotal.updated += r2.updated;
        userTotal.inserted += r2.inserted;
      });
    });

    rebuildRollups_();

    const ms = new Date().getTime() - startedAt.getTime();
    setMeta_('last_refresh_at_utc', new Date().toISOString());
    setMeta_('last_refresh_status', 'ok');
    setMeta_('last_refresh_window', starting + ' → ' + ending);
    setMeta_('last_refresh_duration_ms', String(ms));
    setMeta_('last_daily_upserts', JSON.stringify(dailyTotal));
    setMeta_('last_user_upserts', JSON.stringify(userTotal));
    Logger.log('Refresh ok. daily=' + JSON.stringify(dailyTotal) + ' user=' + JSON.stringify(userTotal));
  } catch (e) {
    setMeta_('last_refresh_status', 'error');
    setMeta_('last_refresh_error', String(e && e.message || e));
    Logger.log('Refresh failed: ' + e);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function rebuildRollups_() {
  const daily = readDailyAll_();

  // monthly_summary
  const byMonth = {};
  daily.forEach(function (r) {
    const m = r.date.slice(0, 7);
    byMonth[m] = (byMonth[m] || 0) + r.cost_usd;
  });
  const monthRows = Object.keys(byMonth).sort().map(function (m) {
    return [m, round2_(byMonth[m]), new Date()];
  });
  replaceAll_(CONFIG.SHEET_MONTHLY, MONTHLY_HEADERS, monthRows);

  // product_month
  const byPm = {};
  daily.forEach(function (r) {
    const k = r.date.slice(0, 7) + '|' + r.product;
    byPm[k] = (byPm[k] || 0) + r.cost_usd;
  });
  const pmRows = Object.keys(byPm).sort().map(function (k) {
    const parts = k.split('|');
    return [parts[0], parts[1], round2_(byPm[k]), new Date()];
  });
  replaceAll_(CONFIG.SHEET_PRODUCT_MONTH, PRODUCT_MONTH_HEADERS, pmRows);
}

// =============================================================================
// GEMINI (Google Workspace) ADOPTION
//
// Gemini in Workspace is per-seat licensed — there is NO per-user cost in
// any Google API. So this tracks ADOPTION (active users, per-app usage,
// last-used), sourced from the Admin SDK Reports Activities API. Runs as
// the deploying account, which must hold the Workspace "Reports" admin
// privilege. Uses the AdminReports advanced service (no service account,
// no key — see appsscript.json).
// =============================================================================

function refreshGemini() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) { Logger.log('Gemini refresh: another run in progress'); return; }
  try {
    setMeta_('gemini_refresh_status', 'running');
    // Rolling window. We recompute a fixed trailing window from scratch every
    // run and fully replace the sheet — one row per user. Because each run
    // re-pulls the whole window, late-arriving events self-heal (no watermark,
    // no double-count) and users who fall out of the window drop off, so the
    // numbers reflect CURRENT adoption rather than an ever-growing total.
    const endDate = isoDateUtc_(addDays_(new Date(), -1)); // through yesterday (today is partial)
    const startDate = isoDateUtc_(addDays_(new Date(), -CONFIG.GEMINI_WINDOW_DAYS));

    const perUser = fetchGeminiUsage_(startDate, endDate); // inclusive range
    const nowIso = new Date().toISOString();
    const rows = Object.keys(perUser).map(function (email) {
      const u = perUser[email];
      // [user_email, events, active_days, first_date, last_date, last_updated_utc]
      return [email, u.events, Object.keys(u.days).length, u.firstDate, u.lastDate, nowIso];
    });
    replaceAll_(CONFIG.SHEET_GEMINI, GEMINI_HEADERS, rows);
    setMeta_('gemini_window', startDate + ' → ' + endDate);
    setMeta_('gemini_refresh_at_utc', new Date().toISOString());
    setMeta_('gemini_refresh_status', 'ok');
    Logger.log('Gemini refresh ok: ' + CONFIG.GEMINI_WINDOW_DAYS + 'd window ' + startDate + '→' + endDate + ', ' + rows.length + ' users');
  } catch (e) {
    setMeta_('gemini_refresh_status', 'error');
    setMeta_('gemini_refresh_error', String(e && e.message || e));
    Logger.log('Gemini refresh failed: ' + e);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Diagnostic: prints the exact identity the script is running as and the
 * raw Admin SDK error (not the wrapped one). Run this from the editor and
 * read the Execution log. It does NOT throw the friendly message, so you
 * see Google's actual response.
 */
function verifyGeminiAccess() {
  Logger.log('Active user (you):    ' + Session.getActiveUser().getEmail());
  Logger.log('Effective user (run): ' + Session.getEffectiveUser().getEmail());
  Logger.log('App name queried:     ' + CONFIG.GEMINI_APP_NAME);

  if (typeof AdminReports === 'undefined') {
    Logger.log('RESULT: AdminReports advanced service is NOT enabled in this project.');
    return;
  }
  try {
    var resp = AdminReports.Activities.list('all', CONFIG.GEMINI_APP_NAME, { maxResults: 1 });
    var n = (resp && resp.items) ? resp.items.length : 0;
    Logger.log('RESULT: OK — call succeeded, returned ' + n + ' item(s). Auth is fine.');
  } catch (e) {
    Logger.log('RESULT: RAW ERROR ->');
    Logger.log(e && e.stack ? e.stack : String(e));
    Logger.log('Message: ' + (e && e.message ? e.message : String(e)));
  }
}

/**
 * Page through Gemini activity events over an inclusive [startDate, endDate]
 * day range and aggregate PER USER. Returns:
 *   { email: { events: <int>, days: {YYYY-MM-DD:true}, firstDate, lastDate } }
 *
 * Each activity event = one Gemini action by one user. We deliberately keep
 * no app breakdown (the API's per-app attribution is unreliable).
 *
 * NOTE: Google's Admin-console "Gemini usage" report (Overall usage / Days
 * at limit / Studio columns) is NOT available via any API — only these raw
 * activity events are. So this is an adoption proxy, not a 1:1 match.
 */
function fetchGeminiUsage_(startDate, endDate) {
  if (typeof AdminReports === 'undefined') {
    throw new Error('AdminReports advanced service not enabled. Enable it in the Apps Script editor: Services (+) → Admin SDK API → Reports. Also confirm the manifest lists the "admin / reports_v1" advanced service.');
  }
  const startTime = startDate + 'T00:00:00.000Z';
  // endTime is exclusive, so push to the start of the day AFTER endDate.
  const endTime = isoDateUtc_(addDays_(new Date(endDate + 'T00:00:00.000Z'), 1)) + 'T00:00:00.000Z';

  const perUser = {};
  let pageToken = null;
  let pages = 0;
  do {
    const opts = {
      applicationName: CONFIG.GEMINI_APP_NAME,
      startTime: startTime,
      endTime: endTime,
      maxResults: 1000,
    };
    if (pageToken) opts.pageToken = pageToken;

    let resp;
    try {
      resp = AdminReports.Activities.list('all', CONFIG.GEMINI_APP_NAME, opts);
    } catch (e) {
      // Most common cause: the running account lacks the Reports admin
      // privilege, or the advanced service isn't enabled.
      throw new Error('Gemini activity fetch failed (' + e + '). The account running this needs the Workspace "Reports" admin privilege, and the Admin SDK Reports advanced service must be enabled.');
    }

    const items = (resp && resp.items) || [];
    items.forEach(function (ev) {
      const email = (ev.actor && ev.actor.email) ? ev.actor.email.toLowerCase() : '';
      if (!email) return;
      const date = String(ev.id && ev.id.time ? ev.id.time : '').slice(0, 10);
      if (!date) return;
      let u = perUser[email];
      if (!u) { u = perUser[email] = { events: 0, days: {}, firstDate: date, lastDate: date }; }
      u.events++;
      u.days[date] = true;
      if (date < u.firstDate) u.firstDate = date;
      if (date > u.lastDate) u.lastDate = date;
    });

    pageToken = resp && resp.nextPageToken;
    pages++;
  } while (pageToken && pages < 500);
  return perUser;
}

function readGeminiAll_() {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_GEMINI);
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const vals = sheet.getRange(2, 1, last - 1, GEMINI_HEADERS.length).getValues();
  return vals.map(function (r) {
    return {
      user_email: r[0],
      events: Number(r[1]) || 0,
      active_days: Number(r[2]) || 0,
      first_date: formatDateCell_(r[3]),
      last_date: formatDateCell_(r[4]),
    };
  });
}

// =============================================================================
// SHEET STORE
// =============================================================================

function getOrCreateSheet_(name, headers) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a1814').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a1814').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function upsertByKey_(sheetName, headers, rows) {
  if (!rows.length) return { updated: 0, inserted: 0 };
  const sheet = getOrCreateSheet_(sheetName, headers);
  const keyCol = headers.length;
  const lastRow = sheet.getLastRow();

  const existing = {};
  if (lastRow > 1) {
    const keys = sheet.getRange(2, keyCol, lastRow - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (keys[i][0]) existing[keys[i][0]] = i + 2;
    }
  }

  const updates = [];
  const inserts = [];
  rows.forEach(function (r) {
    const k = r[r.length - 1];
    if (existing[k]) updates.push({ row: existing[k], values: r });
    else inserts.push(r);
  });

  updates.forEach(function (u) {
    sheet.getRange(u.row, 1, 1, headers.length).setValues([u.values]);
  });

  if (inserts.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, inserts.length, headers.length).setValues(inserts);
  }
  return { updated: updates.length, inserted: inserts.length };
}

function replaceAll_(sheetName, headers, rows) {
  const sheet = getOrCreateSheet_(sheetName, headers);
  const last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function setMeta_(key, value) {
  const sheet = getOrCreateSheet_(CONFIG.SHEET_META, META_HEADERS);
  const last = sheet.getLastRow();
  const now = new Date();
  if (last > 1) {
    const keys = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (keys[i][0] === key) {
        sheet.getRange(i + 2, 1, 1, 3).setValues([[key, value, now]]);
        return;
      }
    }
  }
  sheet.appendRow([key, value, now]);
}

function getMeta_(key) {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_META);
  if (!sheet) return null;
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const rows = sheet.getRange(2, 1, last - 1, 2).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === key) return rows[i][1];
  }
  return null;
}

/**
 * Remove all rows from user_cost whose `product` column equals the given
 * value. Used to migrate legacy 'all'-tagged rows (which had no per-
 * product breakdown) so they don't double-count alongside the new
 * per-product rows.
 */
function pruneUserCostByProduct_(productValue) {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_USERS);
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  // product is column index 6 in USER_HEADERS (0-based 5, 1-based 6).
  const PRODUCT_COL = 6;
  const productCol = sheet.getRange(2, PRODUCT_COL, lastRow - 1, 1).getValues();
  // Walk bottom-up so deleteRow indices remain stable.
  let pruned = 0;
  for (let i = productCol.length - 1; i >= 0; i--) {
    if (String(productCol[i][0]).toLowerCase() === String(productValue).toLowerCase()) {
      sheet.deleteRow(i + 2);
      pruned++;
    }
  }
  if (pruned > 0) Logger.log('pruneUserCostByProduct_: removed ' + pruned + ' "' + productValue + '" rows');
  return pruned;
}

/**
 * Delete every tab whose name contains "__ARCHIVED_". These were created by
 * an older migration that renamed sheets instead of cleaning rows in place.
 * Safe to run anytime — archived tabs hold only superseded data that the
 * live sheets already rebuild from the API. Runnable from the Diagnostics
 * menu.
 */
function cleanupArchivedTabs() {
  const ss = getSpreadsheet_();
  const removed = [];
  ss.getSheets().forEach(function (sh) {
    if (sh.getName().indexOf('__ARCHIVED_') !== -1) {
      removed.push(sh.getName());
      ss.deleteSheet(sh);
    }
  });
  Logger.log('cleanupArchivedTabs: removed ' + removed.length + ' tab(s): ' + removed.join(', '));
  try { ss.toast('Removed ' + removed.length + ' archived tab(s).'); } catch (e) {}
  return removed;
}

/**
 * Return the list of calendar months a date range touches.
 *   monthsInRange_('2026-03-15', '2026-05-10')
 *   → [{key:'2026-03', start:'2026-03-01', end:'2026-03-31'},
 *      {key:'2026-04', start:'2026-04-01', end:'2026-04-30'},
 *      {key:'2026-05', start:'2026-05-01', end:'2026-05-31'}]
 */
function monthsInRange_(startIso, endIso) {
  const out = [];
  const start = parseIsoDate_(startIso);
  const end = parseIsoDate_(endIso);
  let y = start.getUTCFullYear();
  let m = start.getUTCMonth(); // 0-based
  while (true) {
    const first = new Date(Date.UTC(y, m, 1));
    const last = new Date(Date.UTC(y, m + 1, 0)); // day 0 of next month = last day
    const key = Utilities.formatDate(first, 'UTC', 'yyyy-MM');
    out.push({ key: key, start: isoDateUtc_(first), end: isoDateUtc_(last) });
    if (y === end.getUTCFullYear() && m === end.getUTCMonth()) break;
    m++;
    if (m > 11) { m = 0; y++; }
    if (out.length > 120) break; // safety
  }
  return out;
}

/**
 * Delete all user_cost rows whose row_key starts with the given month
 * (YYYY-MM|...). Called before re-writing a month so the overwrite is
 * clean even if the set of products with spend changed between runs.
 */
function pruneUserCostByMonth_(monthKey) {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_USERS);
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const KEY_COL = USER_HEADERS.length; // row_key is the last column
  const keys = sheet.getRange(2, KEY_COL, lastRow - 1, 1).getValues();
  const prefix = monthKey + '|';
  let pruned = 0;
  for (let i = keys.length - 1; i >= 0; i--) {
    if (String(keys[i][0]).indexOf(prefix) === 0) {
      sheet.deleteRow(i + 2);
      pruned++;
    }
  }
  return pruned;
}

/**
 * Migration / self-heal. Removes legacy user_cost rows whose row_key uses
 * the old "fulldate|fulldate|email|product" form (or the even older "|all"
 * form) so the month-keyed refresh can rebuild them cleanly. Cleans rows
 * IN PLACE — it never creates an archive tab (the old approach renamed the
 * whole sheet on every run, which spawned a new tab each time). Safe to run
 * every refresh: it's a no-op once the sheet is in the new format.
 */
function migrateUserCostIfNeeded_() {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_USERS);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const width = USER_HEADERS.length;
  const KEY_IDX = width - 1; // row_key is the last column
  const data = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  // New keys look like "2026-05|email|product" — the first segment is
  // exactly 7 chars (YYYY-MM). Legacy keys have a full date (10 chars) or
  // contain "|all".
  const survivors = [];
  let legacy = 0;
  data.forEach(function (row) {
    const k = String(row[KEY_IDX]);
    const firstSeg = k.split('|')[0];
    const isLegacy = firstSeg.length !== 7 || k.indexOf('|all') !== -1;
    if (isLegacy) legacy++; else survivors.push(row);
  });
  if (!legacy) return;

  Logger.log('migrateUserCostIfNeeded_: removing ' + legacy + ' legacy user_cost row(s) in place.');
  replaceAll_(CONFIG.SHEET_USERS, USER_HEADERS, survivors);
}

/**
 * Reconcile per-user totals against the bucketed daily totals for each
 * month. Writes one row per month to the log:
 *   2026-05: daily=$3,131.56  users=$3,090.21  ratio=0.987  ✓
 *
 * Ratios near 1.00 mean the two endpoints agree. A small residual (1-3%)
 * is expected — some spend isn't attributable to a named user (system
 * accounts, the "other" bucket). A ratio that's wildly off (0.1, 10) means
 * something's miscounted; first place to look is the product slug list.
 *
 * Run from the editor whenever totals look surprising.
 */
function reconcileCheck() {
  const daily = readDailyAll_();
  const users = readUsersAll_();

  // Sum daily by month
  const dailyByMonth = {};
  daily.forEach(function (r) {
    const m = (r.date || '').slice(0, 7);
    if (!m) return;
    dailyByMonth[m] = (dailyByMonth[m] || 0) + r.cost_usd;
  });

  // Sum users by month (using period_end as the month assignment)
  const usersByMonth = {};
  users.forEach(function (u) {
    const m = (u.period_end || '').slice(0, 7);
    if (!m) return;
    usersByMonth[m] = (usersByMonth[m] || 0) + u.cost_usd;
  });

  const allMonths = {};
  Object.keys(dailyByMonth).forEach(function (m) { allMonths[m] = true; });
  Object.keys(usersByMonth).forEach(function (m) { allMonths[m] = true; });

  Logger.log('=== reconcileCheck ===');
  Logger.log('month     daily      users      ratio   status');
  Logger.log('-------   --------   --------   -----   ------');
  Object.keys(allMonths).sort().forEach(function (m) {
    const d = dailyByMonth[m] || 0;
    const u = usersByMonth[m] || 0;
    const ratio = d > 0 ? u / d : 0;
    let status;
    if (d === 0 && u === 0) status = '—';
    else if (ratio >= 0.95 && ratio <= 1.05) status = '✓';
    else if (ratio >= 0.90 && ratio <= 1.10) status = 'OK';
    else if (ratio >= 0.50 && ratio <= 1.50) status = 'WARN';
    else status = 'FAIL';
    Logger.log(m + '   ' + ('$' + d.toFixed(2)).padEnd(8) + '   ' +
               ('$' + u.toFixed(2)).padEnd(8) + '   ' +
               ratio.toFixed(3) + '   ' + status);
  });
  Logger.log('=== end reconcileCheck ===');
}

/**
 * One-shot diagnostic: hit each cost endpoint with a tiny date window and
 * dump the first raw row from the response. Use to verify field names
 * (e.g. `amount` vs `cost_usd`) if numbers ever look wrong.
 */
function debugDumpOneRow() {
  const ending = isoDateUtc_(addDays_(new Date(), -3));
  const starting = isoDateUtc_(addDays_(new Date(), -5));
  Logger.log('Window: ' + starting + ' → ' + ending);

  Logger.log('\n--- cost_report (bucketed) ---');
  try {
    const r = fetchCostOverTime_(starting, ending, { groupBy: ['product'] });
    Logger.log('Buckets: ' + r.length);
    if (r.length) Logger.log('First bucket: ' + JSON.stringify(r[0]).slice(0, 800));
  } catch (e) { Logger.log('error: ' + e); }

  Logger.log('\n--- user_cost_report (filtered to cowork) ---');
  try {
    const r = fetchCostUsers_(starting, ending, { products: ['cowork'] });
    Logger.log('Users: ' + r.length);
    if (r.length) Logger.log('First user: ' + JSON.stringify(r[0]).slice(0, 800));
  } catch (e) { Logger.log('error: ' + e); }
}


function readDailyAll_() {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_DAILY);
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const vals = sheet.getRange(2, 1, last - 1, DAILY_HEADERS.length).getValues();
  return vals.map(function (r) {
    return {
      date: formatDateCell_(r[0]),
      product: r[1],
      model: r[2],
      cost_usd: Number(r[3]) || 0,
    };
  });
}

function readUsersAll_() {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_USERS);
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const vals = sheet.getRange(2, 1, last - 1, USER_HEADERS.length).getValues();
  return vals.map(function (r) {
    return {
      period_start: formatDateCell_(r[0]),
      period_end: formatDateCell_(r[1]),
      user_email: r[2],
      user_id: r[3],
      user_name: r[4],
      product: r[5],
      cost_usd: Number(r[6]) || 0,
    };
  });
}

/**
 * Read the user-maintained org directory tab. Returns a map of
 *   lowercased email → { email, displayName, title, department, fn, manager }
 *
 * Tolerant of column reordering — we look up positions by header name.
 * Returns an empty map if the sheet is missing or the headers don't match.
 */
function readUserDirectory_() {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_USERS_DIRECTORY);
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return {};

  // Build a header → column-index map (0-based, case- and space-insensitive).
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = {};
  headerRow.forEach(function (h, i) {
    const norm = String(h || '').toLowerCase().replace(/\s+/g, '').replace(/[^a-z]/g, '');
    if (norm) idx[norm] = i;
  });

  // Resolve the columns we care about, with a few common-spelling fallbacks.
  const col = {
    email: pickIdx_(idx, ['email', 'emailaddress']),
    displayName: pickIdx_(idx, ['displayname', 'name', 'fullname']),
    title: pickIdx_(idx, ['title', 'jobtitle']),
    department: pickIdx_(idx, ['department', 'dept']),
    fn: pickIdx_(idx, ['function', 'team']),
    manager: pickIdx_(idx, ['manager', 'reportsto']),
  };
  if (col.email < 0) {
    Logger.log('readUserDirectory_: no "Email" column found in ' + CONFIG.SHEET_USERS_DIRECTORY + ' tab');
    return {};
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const map = {};
  rows.forEach(function (r) {
    const email = String(r[col.email] || '').trim().toLowerCase();
    if (!email) return;
    map[email] = {
      email: email,
      displayName: col.displayName >= 0 ? String(r[col.displayName] || '').trim() : '',
      title:       col.title       >= 0 ? String(r[col.title]       || '').trim() : '',
      department:  col.department  >= 0 ? String(r[col.department]  || '').trim() : '',
      fn:          col.fn          >= 0 ? String(r[col.fn]          || '').trim() : '',
      manager:     col.manager     >= 0 ? String(r[col.manager]     || '').trim() : '',
    };
  });
  return map;
}

function pickIdx_(idx, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    if (idx[candidates[i]] !== undefined) return idx[candidates[i]];
  }
  return -1;
}

// =============================================================================
// DASHBOARD DATA (called by Dashboard.html via google.script.run)
// =============================================================================

function getDashboardData() {
  const daily = readDailyAll_();
  const users = readUsersAll_();
  const directory = readUserDirectory_();
  // Record join stats to _meta so it's visible from the sheet without
  // hunting through execution logs.
  setMeta_('directory_size', String(Object.keys(directory).length));
  setMeta_('directory_tab', CONFIG.SHEET_USERS_DIRECTORY);

  // Monthly trend
  const monthly = {};
  daily.forEach(function (r) {
    const m = r.date.slice(0, 7);
    monthly[m] = (monthly[m] || 0) + r.cost_usd;
  });
  const monthSeries = Object.keys(monthly).sort().map(function (m) {
    return { month: m, cost: round2_(monthly[m]) };
  });

  // Per-product, per-month
  const productByMonth = {};
  const productSet = {};
  daily.forEach(function (r) {
    const m = r.date.slice(0, 7);
    productByMonth[m] = productByMonth[m] || {};
    productByMonth[m][r.product] = (productByMonth[m][r.product] || 0) + r.cost_usd;
    productSet[r.product] = true;
  });
  const productList = Object.keys(productSet).sort();
  const productMonthly = Object.keys(productByMonth).sort().map(function (m) {
    const row = { month: m };
    productList.forEach(function (p) { row[p] = round2_(productByMonth[m][p] || 0); });
    return row;
  });

  // Product totals
  const productTotals = {};
  daily.forEach(function (r) {
    productTotals[r.product] = (productTotals[r.product] || 0) + r.cost_usd;
  });
  const productTotalsList = Object.keys(productTotals).sort(function (a, b) {
    return productTotals[b] - productTotals[a];
  }).map(function (p) {
    return { product: p, cost: round2_(productTotals[p]) };
  });

  // ---- Per-user, per-month aggregation ----
  // We send every user's spend broken out by month, so the dashboard can
  // (a) switch between This Month / YTD / All time client-side with no
  // server round-trip, and (b) draw a per-month sparkline when a row is
  // expanded. Each user: { email, name, dept, title, months: {YYYY-MM:
  // {cost, by_product}} }.
  const allMonthsSet = {};
  const userMonthly = {};
  users.forEach(function (u) {
    const m = (u.period_end || '').slice(0, 7); // YYYY-MM
    if (!m) return;
    allMonthsSet[m] = true;
    const email = u.user_email;
    if (!userMonthly[email]) {
      userMonthly[email] = {
        user_email: email,
        user_id: u.user_id,
        name: u.user_name || formatNameFromEmail_(email),
        months: {},
      };
    }
    const rec = userMonthly[email];
    if (!rec.months[m]) rec.months[m] = { cost: 0, by_product: {} };
    rec.months[m].cost += u.cost_usd;
    rec.months[m].by_product[u.product] = (rec.months[m].by_product[u.product] || 0) + u.cost_usd;
  });

  const allMonths = Object.keys(allMonthsSet).sort();
  const latestMonthKey = allMonths.length ? allMonths[allMonths.length - 1] : '';

  // Enrich with directory info and round.
  let directoryMatchCount = 0;
  const userList = Object.keys(userMonthly).map(function (email) {
    const u = userMonthly[email];
    Object.keys(u.months).forEach(function (m) {
      u.months[m].cost = round2_(u.months[m].cost);
      Object.keys(u.months[m].by_product).forEach(function (p) {
        u.months[m].by_product[p] = round2_(u.months[m].by_product[p]);
      });
    });
    const dirEntry = directory[String(email).trim().toLowerCase()];
    if (dirEntry) {
      directoryMatchCount++;
      u.department = dirEntry.department || '';
      u.title = dirEntry.title || '';
      if (dirEntry.displayName) u.name = dirEntry.displayName;
    } else {
      u.department = '';
      u.title = '';
    }
    return u;
  });
  setMeta_('directory_matches', directoryMatchCount + ' of ' + userList.length + ' users matched');

  // Date span for the latest month (display only).
  let latestStart = '';
  let latestEnd = '';
  users.forEach(function (u) {
    if ((u.period_end || '').slice(0, 7) !== latestMonthKey) return;
    if (!latestStart || (u.period_start || '') < latestStart) latestStart = u.period_start;
    if (!latestEnd || (u.period_end || '') > latestEnd) latestEnd = u.period_end;
  });

  const directoryEmailCount = Object.keys(directory).length;

  // ---- Gemini adoption (usage, not cost) ----
  // Rows are already one-per-user cumulative totals; just attach directory
  // metadata and the field names the dashboard expects.
  const gemUserList = readGeminiAll_().map(function (r) {
    const dirEntry = directory[r.user_email];
    return {
      user_email: r.user_email,
      events: r.events,
      active_days: r.active_days,
      lastDate: r.last_date,
      name: (dirEntry && dirEntry.displayName) ? dirEntry.displayName : formatNameFromEmail_(r.user_email),
      department: dirEntry ? (dirEntry.department || '') : '',
    };
  }).sort(function (a, b) { return b.events - a.events; });
  const gemTotalEvents = gemUserList.reduce(function (s, u) { return s + u.events; }, 0);

  const totalAllTime = daily.reduce(function (s, r) { return s + r.cost_usd; }, 0);
  const latestMonth = monthSeries.length ? monthSeries[monthSeries.length - 1] : null;
  const prevMonth = monthSeries.length > 1 ? monthSeries[monthSeries.length - 2] : null;

  return {
    generated_at: new Date().toISOString(),
    last_refresh_at: getMeta_('last_refresh_at_utc'),
    last_refresh_status: getMeta_('last_refresh_status'),
    last_refresh_window: getMeta_('last_refresh_window'),
    kpis: {
      total_all_time: round2_(totalAllTime),
      latest_month: latestMonth,
      prev_month: prevMonth,
      month_count: monthSeries.length,
      product_count: productList.length,
      user_count: userList.length,
    },
    month_series: monthSeries,
    products: productList,
    product_monthly: productMonthly,
    product_totals: productTotalsList,
    users: {
      latest_month: latestMonthKey,
      latest_start: latestStart,
      latest_end: latestEnd,
      months: allMonths,            // sorted list of YYYY-MM with data
      list: userList,               // each: {user_email, name, department, title, months:{YYYY-MM:{cost,by_product}}}
      directory_size: directoryEmailCount,
    },
    gemini: {
      refresh_status: getMeta_('gemini_refresh_status'),
      refresh_at: getMeta_('gemini_refresh_at_utc'),
      total_events: gemTotalEvents,
      active_users: gemUserList.length,
      users: gemUserList,
    },
  };
}

// =============================================================================
// HELPERS
// =============================================================================

function isoDateUtc_(d) { return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd'); }

// RFC3339 UTC timestamp for the Reports Activities API startTime param.
function isoTimestampUtc_(d) { return Utilities.formatDate(d, 'UTC', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'"); }

function addDays_(d, n) {
  const x = new Date(d.getTime());
  x.setUTCDate(x.getUTCDate() + n);
  return x;
}

function chunkDateRange_(start, end, maxDays) {
  const chunks = [];
  let cur = parseIsoDate_(start);
  const last = parseIsoDate_(end);
  while (cur.getTime() <= last.getTime()) {
    const chunkEnd = new Date(Math.min(addDays_(cur, maxDays - 1).getTime(), last.getTime()));
    chunks.push({ start: isoDateUtc_(cur), end: isoDateUtc_(chunkEnd) });
    cur = addDays_(chunkEnd, 1);
  }
  return chunks;
}

function parseIsoDate_(s) {
  const parts = s.split('-').map(Number);
  return new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
}

function toIsoTimestamp_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, 'UTC', "yyyy-MM-dd'T'00:00:00'Z'");
  if (typeof d === 'string') return d.length === 10 ? d + 'T00:00:00Z' : d;
  throw new Error('toIsoTimestamp_: bad value ' + d);
}

function round2_(n) { return Math.round(n * 100) / 100; }

function formatDateCell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'UTC', 'yyyy-MM-dd');
  return String(v);
}

function formatNameFromEmail_(email) {
  if (!email) return '';
  return email.split('@')[0].split('.').map(function (s) {
    return s.charAt(0).toUpperCase() + s.slice(1);
  }).join(' ');
}

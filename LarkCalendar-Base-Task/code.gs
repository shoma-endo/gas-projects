/* ===================== 0. 基本設定 =========================== */
// シート名
const ENV_SHEET_NAME = '環境変数'; 

// 同期デフォルト
const CAL_DEFAULT_PAST_DAYS = -14;
const CAL_DEFAULT_FUTURE_DAYS = 30;

// Bitable列名
const EN_FIELD = {
  EVENT_ID: 'event_id',
  CAL_ID: 'calendar_id',
  CAL_TITLE: 'calendar_title',
  TITLE: 'title',
  START: 'start_time', // DateTime(ms)
  END: 'end_time',     // DateTime(ms)
  LAST: 'last_seen_at',// DateTime(ms)
  DOC: 'doc_url',      // Text or Link
  TOTAL: 'total',      // Number(hours)
  EXPL: 'explanation'  // Text(description)
};

// 固定
const MY_TASKS_ID = 'my_tasks';     // Task v2の仮想カレンダーID
const CAL_BATCH_SIZE = 200;
const CAL_CREATE_MINUTES = true;    // meeting_minute自動作成
const CAL_SOFT_DELETE = false;
const SOFT_DELETE_FIELD = 'deleted';

// 正規表現
const KEY_RE = /^[A-Z][A-Z0-9_]{2,40}$/;

// doc_urlの型キャッシュ(true=Link型, false=Text型, null=未判定)
let DOC_FIELD_IS_LINK = null;
let __SCHEMA_READY = false;

/* ===================== 1. 便利ヘルパ ========================= */
// UI返却フォーマット
function uiOk(payload) {
  return Object.assign({ ok: true }, payload || {});
}
function uiErr(e) {
  // 例外内容を拾って UI に渡す
  return {
    ok: false,
    error: String(e && e.message ? e.message : e),
    detail: {
      http: e && e.http != null ? e.http : null,
      apiCode: e && e.apiCode != null ? e.apiCode : null,
      url: e && e.url ? e.url : null,
      body: e && e.body ? String(e.body).slice(0, 1000) : null,
      headers: e && e.headers ? e.headers : null
    }
  };
}

/* ===================== 1.1 環境変数読取 ====================== */
// フォーマット: 1行目=KEY群, 2行目=VALUE群(横長) もしくは A列=KEY, B列=VALUE(縦長)
function getEnv_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ENV_SHEET_NAME);
  if (!sh) throw new Error(`「${ENV_SHEET_NAME}」シートがありません。`);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const env = {};

  // 横長形式(1行目ヘッダ/2行目値)
  if (lastRow >= 2 && lastCol >= 1) {
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
    const values  = sh.getRange(2, 1, 1, lastCol).getValues()[0].map(v => String(v == null ? '' : v).trim());
    let any = false;
    for (let c = 0; c < headers.length; c++) {
      const k = headers[c], v = values[c];
      if (!k || v === '') continue;
      if (!KEY_RE.test(k)) continue;
      env[k] = v; any = true;
    }
    if (any) return env;
  }

  // 縦長形式(A=KEY, B=VALUE)
  const rows = sh.getRange(1, 1, lastRow, Math.min(2, lastCol)).getValues();
  for (let i = 1; i < rows.length; i++) {
    const k = String(rows[i][0] || '').trim();
    const v = String(rows[i][1] == null ? '' : rows[i][1]).trim();
    if (!k || v === '') continue;
    if (!KEY_RE.test(k)) continue;
    env[k] = v;
  }
  return env;
}

/* === 同期設定の保存/読取（ScriptProperties 版） === */
function saveSyncSettings_(selectedCalIds, pastDays, futureDays) {
  const sp = PropertiesService.getScriptProperties();
  sp.setProperty('SYNC_SELECTED_CAL_IDS', JSON.stringify(selectedCalIds || []));
  sp.setProperty('SYNC_PAST_DAYS',   String(Number(pastDays  ?? CAL_DEFAULT_PAST_DAYS)));
  sp.setProperty('SYNC_FUTURE_DAYS', String(Number(futureDays?? CAL_DEFAULT_FUTURE_DAYS)));
}
function loadSyncSettings_() {
  const sp = PropertiesService.getScriptProperties();
  const selected = JSON.parse(sp.getProperty('SYNC_SELECTED_CAL_IDS') || '[]');
  const past     = Number(sp.getProperty('SYNC_PAST_DAYS')   || CAL_DEFAULT_PAST_DAYS);
  const future   = Number(sp.getProperty('SYNC_FUTURE_DAYS') || CAL_DEFAULT_FUTURE_DAYS);
  return { selectedCalIds: selected, pastDays: past, futureDays: future };
}

/* ===================== 1.2 オリジン判定 ====================== */
function getOpenOrigin_() {
  const env = getEnv_();
  const ui = env['LARK_BASE_UI_URL'] || '';
  try {
    const u = new URL(ui);
    const host = u.hostname || '';
    if (host.endsWith('larksuite.com')) return 'https://open.larksuite.com';
    if (host.endsWith('feishu.cn')) return 'https://open.feishu.cn';
  } catch (_) {}
  return 'https://open.larksuite.com';
}
function bitableAppsRoot_() { return getOpenOrigin_() + '/open-apis/bitable/v1/apps'; }

/* ===================== 1.3 Base/Table 取得 =================== */
function parseBitableUiUrl_(url) {
  if (!url) return { baseId: '', tableId: '' };
  try {
    const u = new URL(url);
    const segs = u.pathname.split('/').filter(Boolean);
    const idx = segs.indexOf('base');
    const baseId = (idx >= 0 && idx + 1 < segs.length) ? segs[idx + 1] : '';
    const tableId = u.searchParams.get('table') || '';
    return { baseId, tableId };
  } catch (_) {
    const baseId =
      (url.match(/\/base\/([A-Za-z0-9_-]{20,40})/) || [])[1] ||
      (url.match(/\/app\/([A-Za-z0-9_-]{20,40})/) || [])[1] || '';
    const tableId = (url.match(/[?&]table=(tbl[A-Za-z0-9]+)/) || [])[1] || '';
    return { baseId, tableId };
  }
}
function getCfg_() {
  const env = getEnv_();
  let baseId = env['LARK_BASE_ID'] || '';
  let tableId = env['LARK_TABLE_ID'] || '';
  if (!baseId || !tableId) {
    const { baseId: b, tableId: t } = parseBitableUiUrl_(env['LARK_BASE_UI_URL'] || '');
    baseId = baseId || b;
    tableId = tableId || t;
  }
  if (!baseId || !tableId) throw new Error('LARK_BASE_ID/LARK_TABLE_ID(またはUI URLからの特定)が不足しています。');
  return { baseId, tableId };
}

/* ===================== 2. HTTP/認証 ========================== */
function fetchJson_(url, opt) {
  const res = UrlFetchApp.fetch(url, Object.assign({ muteHttpExceptions: true }, opt || {}));
  const http = res.getResponseCode();
  const text = res.getContentText() || '';
  let json = null; try { json = JSON.parse(text); } catch (_) {}
  if (http >= 200 && http < 300) return json || {};
  const apiCode = (json && typeof json.code === 'number') ? json.code : null;
  const msg = (json && (json.msg || json.message)) || ('HTTP ' + http);
  const err = new Error(msg);
  err.http = http; err.apiCode = apiCode; err.body = text; err.url = url;
  err.headers = (res.getAllHeaders && res.getAllHeaders()) || (res.getHeaders && res.getHeaders()) || {};
  throw err;
}
function fetchJsonAny_(url, opt) {
  const res = UrlFetchApp.fetch(url, Object.assign({ muteHttpExceptions: true }, opt || {}));
  const text = res.getContentText() || '';
  let json = null; try { json = JSON.parse(text); } catch (_) {}
  return { http: res.getResponseCode(), json: json || {}, text, headers: (res.getAllHeaders && res.getAllHeaders()) || (res.getHeaders && res.getHeaders()) || {} };
}

// tenant/userトークン
function getTenantAccessToken_() {
  const origin = getOpenOrigin_();
  const env = getEnv_();
  const appId = env['LARK_APP_ID'], secret = env['LARK_APP_SECRET'];
  if (!appId || !secret) throw new Error('LARK_APP_ID/LARK_APP_SECRETが未設定です(環境変数シート)。');

  const sp = PropertiesService.getScriptProperties();
  const keyTok = 'TENANT_TOKEN_' + origin, keyTtl = keyTok + '_TTL';
  const cached = sp.getProperty(keyTok), ttl = Number(sp.getProperty(keyTtl) || 0);
  if (cached && Date.now() < ttl) return cached;

  const js = fetchJson_(
    origin + '/open-apis/auth/v3/tenant_access_token/internal',
    { method: 'post', contentType: 'application/json', payload: JSON.stringify({ app_id: appId, app_secret: secret }) }
  );
  if (js.code !== 0) throw new Error('tenant_access_token取得失敗: ' + js.msg);
  const token = js.tenant_access_token, expire = js.expire || 1800;
  sp.setProperty(keyTok, token);
  sp.setProperty(keyTtl, String(Date.now() + (expire - 60) * 1000));
  return token;
}
function authHeaderTenant_() { return { Authorization: 'Bearer ' + getTenantAccessToken_() }; }

function userAuthHeader_() {
  const sp = PropertiesService.getScriptProperties();
  const now = Date.now();
  let at = sp.getProperty('USER_ACCESS_TOKEN');
  let exp = Number(sp.getProperty('USER_TOKEN_EXPIRE') || 0);
  if (!at) throw new Error('USER_ACCESS_TOKENがありません。まずOAuth認可してください。');
  if (now >= exp - 60000) {
    const d = refreshUserAccessToken();
    at = d.access_token;
    exp = Date.now() + d.expires_in * 1000;
  }
  return { Authorization: 'Bearer ' + at, 'Content-Type': 'application/json; charset=utf-8' };
}

/* ===================== 3. OAuth(ユーザー) ==================== */
function getAppAccessToken_() {
  const origin = getOpenOrigin_();
  const env = getEnv_();
  const js = fetchJson_(
    origin + '/open-apis/auth/v3/app_access_token/internal/',
    { method: 'post', contentType: 'application/json', payload: JSON.stringify({ app_id: env['LARK_APP_ID'], app_secret: env['LARK_APP_SECRET'] }) }
  );
  if (js.code !== 0) throw new Error('app_access_token失敗: ' + js.msg);
  return js.app_access_token;
}
function getOAuthURL() {
  const origin = getOpenOrigin_();
  const env = getEnv_();
  const appId = env['LARK_APP_ID'];
  const rawRedirect = env['OAUTH_REDIRECT_URI'];
  if (!appId || !rawRedirect) throw new Error('LARK_APP_ID / OAUTH_REDIRECT_URI が未設定です。');
  const state = Utilities.getUuid();
  CacheService.getScriptCache().put('OAUTH_STATE', state, 300);
  return `${origin}/open-apis/authen/v1/index?app_id=${appId}&redirect_uri=${encodeURIComponent(rawRedirect)}&state=${state}`;
}
function doGet(e) {
  try {
    const code = e?.parameter?.code;
    const state = e?.parameter?.state;
    const expect = CacheService.getScriptCache().get('OAUTH_STATE');
    if (!code) return HtmlService.createHtmlOutput('OK (no code)');
    if (expect && state !== expect) return HtmlService.createHtmlOutput('state mismatch');

    const data = exchangeAuthCodeForUserToken_(code);
    return HtmlService.createHtmlOutput('OK: user_access_token saved. ExpiresIn=' + data.expires_in);
  } catch (err) {
    // ブラウザで目視できるように
    return HtmlService.createHtmlOutput('ERROR: ' + String(err && err.message ? err.message : err));
  }
}
function exchangeAuthCodeForUserToken_(code) {
  const origin = getOpenOrigin_();
  const appTok = getAppAccessToken_();
  const js = fetchJson_(origin + '/open-apis/authen/v1/access_token', {
    method: 'post', headers: { Authorization: `Bearer ${appTok}` }, contentType: 'application/json',
    payload: JSON.stringify({ grant_type: 'authorization_code', code })
  });
  if (js.code !== 0) throw new Error('exchange failed: ' + js.msg);

  const now = Date.now(), d = js.data, sp = PropertiesService.getScriptProperties();
  sp.setProperty('USER_ACCESS_TOKEN', d.access_token);
  sp.setProperty('USER_TOKEN_EXPIRE', String(now + d.expires_in * 1000));
  sp.setProperty('USER_REFRESH_TOKEN', d.refresh_token);
  if (d.refresh_expires_in) sp.setProperty('USER_REFRESH_EXPIRE', String(now + d.refresh_expires_in * 1000));
  return d;
}
function refreshUserAccessToken() {
  const origin = getOpenOrigin_();
  const sp = PropertiesService.getScriptProperties();
  const rt = sp.getProperty('USER_REFRESH_TOKEN');
  if (!rt) throw new Error('USER_REFRESH_TOKENがありません。まず認可してください。');
  const appTok = getAppAccessToken_();
  const js = fetchJson_(origin + '/open-apis/authen/v1/refresh_access_token', {
    method: 'post', headers: { Authorization: `Bearer ${appTok}` }, contentType: 'application/json',
    payload: JSON.stringify({ grant_type: 'refresh_token', refresh_token: rt })
  });
  if (js.code !== 0) throw new Error('refresh failed: ' + js.msg);

  const now = Date.now(), d = js.data;
  const sp2 = PropertiesService.getScriptProperties();
  sp2.setProperty('USER_ACCESS_TOKEN', d.access_token);
  sp2.setProperty('USER_TOKEN_EXPIRE', String(now + d.expires_in * 1000));
  sp2.setProperty('USER_REFRESH_TOKEN', d.refresh_token);
  if (d.refresh_expires_in) sp2.setProperty('USER_REFRESH_EXPIRE', String(now + d.refresh_expires_in * 1000));
  return d;
}

/* ===================== 4. Bitableユーティリティ ================ */
function ensureSchema_() {
  if (__SCHEMA_READY) return;
  const { baseId, tableId } = getCfg_();
  const tokenH = authHeaderTenant_();

  const fieldsJ = fetchJson_(
    `${bitableAppsRoot_()}/${baseId}/tables/${tableId}/fields?page_size=200`,
    { method: 'get', headers: tokenH }
  );
  if (fieldsJ.code !== 0) throw new Error('フィールド一覧取得失敗: ' + fieldsJ.msg);

  const items = fieldsJ.data.items || [];
  const existing = new Set(items.map(f => f.field_name));
  const docField = items.find(f => f.field_name === EN_FIELD.DOC);
  DOC_FIELD_IS_LINK = docField ? (Number(docField.type) !== 1) : false;

  const WANT = [
    { name: EN_FIELD.EVENT_ID,  type: 1 },
    { name: EN_FIELD.CAL_ID,    type: 1 },
    { name: EN_FIELD.CAL_TITLE, type: 1 },
    { name: EN_FIELD.TITLE,     type: 1 },
    { name: EN_FIELD.START,     type: 5 },
    { name: EN_FIELD.END,       type: 5 },
    { name: EN_FIELD.LAST,      type: 5 },
    { name: EN_FIELD.DOC,       type: (docField ? Number(docField.type) : 1) },
    { name: EN_FIELD.TOTAL,     type: 2 },
    { name: EN_FIELD.EXPL,      type: 1 }
  ];
  if (CAL_SOFT_DELETE) WANT.push({ name: SOFT_DELETE_FIELD, type: 7 });

  let created = 0;
  for (const w of WANT) {
    if (existing.has(w.name)) continue;

    const res = UrlFetchApp.fetch(
      `${bitableAppsRoot_()}/${baseId}/tables/${tableId}/fields`,
      {
        method: 'post',
        contentType: 'application/json',
        headers: tokenH,
        payload: JSON.stringify({ field_name: w.name, type: w.type, property: {} }),
        muteHttpExceptions: true
      }
    );
    const js = JSON.parse(res.getContentText() || '{}');
    if (js.code === 0) { created++; continue; }

    // フォールバック(type:1)
    const res2 = UrlFetchApp.fetch(
      `${bitableAppsRoot_()}/${baseId}/tables/${tableId}/fields`,
      {
        method: 'post',
        contentType: 'application/json',
        headers: tokenH,
        payload: JSON.stringify({ field_name: w.name, type: 1, property: {} }),
        muteHttpExceptions: true
      }
    );
    const js2 = JSON.parse(res2.getContentText() || '{}');
    if (js2.code === 0) created++;
  }
  if (created) Logger.log('🧱 スキーマ作成: ' + created + ' 列追加');

  __SCHEMA_READY = true;
}

function putDoc_(fields, url, textOpt) {
  if (!url) return;
  const text = textOpt || 'Open';
  if (DOC_FIELD_IS_LINK) fields[EN_FIELD.DOC] = { text: text, link: url };
  else fields[EN_FIELD.DOC] = url;
}

function listAllRecords_Within_(fromSec, toSec, calendarIdOpt) {
  const { baseId, tableId } = getCfg_();
  const urlBase = `${bitableAppsRoot_()}/${baseId}/tables/${tableId}/records`;
  const headers = authHeaderTenant_();

  const fs = fromSec * 1000, ts = toSec * 1000;
  let items = [], page = '';
  do {
    const js = fetchJson_(`${urlBase}?page_size=500${page ? '&page_token=' + page : ''}&field_key=field_name`, { method: 'get', headers });
    if (js.code !== 0) throw new Error('Bitableレコード取得失敗: ' + js.msg);

    (js.data.items || []).forEach(r => {
      const f = r.fields || {};
      const s = Number(f[EN_FIELD.START] || 0);
      const e = Number(f[EN_FIELD.END] || 0);
      const inWindow = (s && s >= fs && s <= ts) || (e && e >= fs && e <= ts);
      if (!inWindow) return;
      if (calendarIdOpt && String(f[EN_FIELD.CAL_ID] || '') !== String(calendarIdOpt)) return;
      items.push({ record_id: r.record_id, event_id: f[EN_FIELD.EVENT_ID], last_seen: f[EN_FIELD.LAST] || 0 });
    });
    page = js.data.page_token || '';
  } while (page);

  const byEventId = {};
  items.forEach(x => { byEventId[String(x.event_id)] = x; });
  return byEventId;
}

function batchCreate_(records) {
  if (!Array.isArray(records) || records.length === 0) return 0;
  const { baseId, tableId } = getCfg_();
  const url = `${bitableAppsRoot_()}/${baseId}/tables/${tableId}/records/batch_create?field_key=field_name`;
  let total = 0;
  for (let i = 0; i < records.length; i += CAL_BATCH_SIZE) {
    const part = records.slice(i, i + CAL_BATCH_SIZE);
    if (!part.length) continue;
    const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', headers: authHeaderTenant_(), payload: JSON.stringify({ records: part }), muteHttpExceptions: true });
    const js = JSON.parse(res.getContentText() || '{}');
    if (res.getResponseCode() !== 200 || js.code !== 0) throw new Error('HTTP ' + res.getResponseCode() + ': ' + (res.getContentText() || ''));
    total += (js.data.records || []).length;
  }
  return total;
}
function batchUpdate_(records) {
  if (!Array.isArray(records) || records.length === 0) return 0;
  const { baseId, tableId } = getCfg_();
  const url = `${bitableAppsRoot_()}/${baseId}/tables/${tableId}/records/batch_update?field_key=field_name`;
  let total = 0;
  for (let i = 0; i < records.length; i += CAL_BATCH_SIZE) {
    const part = records.slice(i, i + CAL_BATCH_SIZE);
    if (!part.length) continue;
    const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', headers: authHeaderTenant_(), payload: JSON.stringify({ records: part }), muteHttpExceptions: true });
    const js = JSON.parse(res.getContentText() || '{}');
    if (res.getResponseCode() !== 200 || js.code !== 0) throw new Error('HTTP ' + res.getResponseCode() + ': ' + (res.getContentText() || ''));
    total += (js.data.records || []).length;
  }
  return total;
}
function batchDelete_(recordIds) {
  if (!Array.isArray(recordIds) || recordIds.length === 0) return 0;

  const { baseId, tableId } = getCfg_();
  const url = `${bitableAppsRoot_()}/${baseId}/tables/${tableId}/records/batch_delete`;

  let total = 0;

  for (let i = 0; i < recordIds.length; i += CAL_BATCH_SIZE) {
    const part = recordIds.slice(i, i + CAL_BATCH_SIZE);
    if (!part.length) continue;

    Logger.log(`[DEL] chunk=${Math.floor(i / CAL_BATCH_SIZE) + 1} size=${part.length} sample=${part.slice(0, 3).join(',')}`);

    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: authHeaderTenant_(),
      payload: JSON.stringify({ records: part }),
      muteHttpExceptions: true
    });

    const http = res.getResponseCode();
    const body = res.getContentText() || '';
    Logger.log(`[DEL] http=${http} body_head=${body.slice(0, 500)}`);

    const js = JSON.parse(body || '{}');
    if (http !== 200 || js.code !== 0) throw new Error('HTTP ' + http + ': ' + body);

    // ここが修正点:deleted_countが無い場合、data.records配列で数える
    const deletedByCount = Number(js.data && js.data.deleted_count) || 0;
    const deletedByRecords = Array.isArray(js.data && js.data.records) ? js.data.records.filter(r => r && r.deleted === true).length : 0;

    const added = deletedByCount || deletedByRecords || part.length; // 最後はフォールバック
    Logger.log(`[DEL] counted=${added} (deleted_count=${deletedByCount}, records_deleted=${deletedByRecords})`);

    total += added;
  }

  return total;
}

/* ===================== 5. カレンダー/イベント取得 ============= */
function toMsFromEventTime_(t, fallbackTz) {
  if (!t) return null;
  if (t.timestamp != null) {
    const raw = String(t.timestamp);
    if (/^\d{13}$/.test(raw)) return Number(raw);
    if (/^\d{10}$/.test(raw)) return Number(raw) * 1000;
    const n = Number(raw); return isFinite(n) ? n : null;
  }
  if (t.datetime) {
    const d = new Date(t.datetime);
    return isNaN(d) ? null : d.getTime();
  }
  if (t.date) {
    const tz = t.timezone || fallbackTz || Session.getScriptTimeZone() || 'Asia/Tokyo';
    const [y, m, d] = String(t.date).split('-').map(Number);
    const dt = new Date(y, m - 1, d, 0, 0, 0);
    return dt.getTime();
  }
  return null;
}

function getPrimaryCalendarId_() {
  const origin = getOpenOrigin_();
  const js = fetchJson_(origin + '/open-apis/calendar/v4/calendars/primary', { method: 'post', headers: userAuthHeader_() });
  if (js.code !== 0) throw new Error('primary取得失敗: ' + js.msg);
  const cal = js.data && js.data.calendar;
  const id = cal && (cal.calendar_id || cal.id);
  if (!id) throw new Error('primaryにcalendar_idがありません');
  return id;
}
function getCalendarInfo_(calId) {
  const origin = getOpenOrigin_();
  const js = fetchJson_(origin + '/open-apis/calendar/v4/calendars/' + calId, { method: 'get', headers: userAuthHeader_() });
  if (js.code !== 0) throw new Error('カレンダー情報取得失敗: ' + js.msg);
  const c = js.data.calendar || {};
  return { id: c.calendar_id || c.id, summary: c.summary || c.name || '', type: c.type || '', access_role: c.access_role || c.role || '', timezone: c.timezone || (Session.getScriptTimeZone() || 'Asia/Tokyo') };
}
function listMyCalendars() {
  const origin = getOpenOrigin_();
  const js = fetchJson_(origin + '/open-apis/calendar/v4/calendars?page_size=200', { method: 'get', headers: userAuthHeader_() });
  if (js.code !== 0) throw new Error('カレンダー一覧取得失敗: ' + js.msg);
  const items = (js.data.calendar_list || js.data.items || js.data.calendars || []).map(c => ({ calendar_id: c.calendar_id || c.id, summary: c.summary || c.name || '', access_role: c.access_role || c.role || '', type: c.type || '' }));
  // My tasksをUI選択用に追加(選択時のみ同期)
  items.push({ calendar_id: MY_TASKS_ID, summary: 'My tasks', access_role: 'owner', type: 'task' });
  return items;
}

function getEvents_(calId, fromSecIn, toSecIn) {
  const origin = getOpenOrigin_();
  let meta = { summary: '', timezone: Session.getScriptTimeZone() || 'Asia/Tokyo', type: '' };
  try { meta = getCalendarInfo_(calId); } catch (_) {}

  let fromSec = fromSecIn, toSec = toSecIn;
  if (fromSec >= toSec) { const t = fromSec; fromSec = toSec; toSec = t; }

  function buildUrl(noPageSize, pageToken, s, e) {
    const base = `${origin}/open-apis/calendar/v4/calendars/${calId}/events?start_time=${s}&end_time=${e}`;
    const ps = noPageSize ? '' : '&page_size=1000';
    const pt = pageToken ? `&page_token=${encodeURIComponent(pageToken)}` : '';
    return base + ps + pt;
  }
  function inWindowMs_(sMs, eMs) {
    const fs = fromSec * 1000, ts = toSec * 1000;
    return (sMs != null && sMs >= fs && sMs <= ts) || (eMs != null && eMs >= fs && eMs <= ts);
  }

  function fetchRange(s, e) {
    let noPageSize = false;
    const out = [];
    let pageToken = '';
    while (true) {
      const { http, json, text } =
        fetchJsonAny_(buildUrl(noPageSize, pageToken, s, e), { method: 'get', headers: userAuthHeader_() });
      const code = typeof json.code === 'number' ? json.code : null;

      if (http !== 200) {
        if (code === 191004) throw new Error('UNSUPPORTED:invalid calendar type');
        if (code === 191002) throw new Error('NOACCESS:no calendar access_role');
        if (code != null) throw new Error('ERROR:' + (json.msg || 'calendar error'));
        throw new Error('ERROR:HTTP ' + http);
      }
      if (json.code !== 0) throw new Error('ERROR:' + (json.msg || 'calendar error'));

      (json.data.items || []).forEach(ev => {
        if (ev.status === 'cancelled') return;
        const sMs = toMsFromEventTime_(ev.start_time, meta.timezone);
        const eMs = toMsFromEventTime_(ev.end_time, meta.timezone);
        if (!inWindowMs_(sMs, eMs)) return;
        out.push({
          event_id: ev.event_id,
          title: (ev.summary && ev.summary.trim()) ? ev.summary : '（無題のイベント）',
          description: ev.description || '',
          start_ms: sMs, end_ms: eMs
        });
      });

      pageToken = json.data.page_token || '';
      if (!pageToken) break;
    }
    return out;
  }

  try {
    return fetchRange(fromSec, toSec);
  } catch (e) {
    const msg = String(e.message || '');
    if (/190002|invalid parameters|ERROR:HTTP 400/i.test(msg)) {
      Logger.log('↻ 190002対策：週分割で再取得します');
      const WEEK = 7 * 86400;
      const items = [];
      for (let s = fromSec; s < toSec; s += WEEK) {
        const e2 = Math.min(s + WEEK, toSec);
        const chunk = fetchRange(s, e2);
        if (chunk && chunk.length) items.push.apply(items, chunk);
      }
      return items;
    }
    throw e;
  }
}

/* ===================== 6. Task v2(選択時のみに修正済み) ================= */
function extractTaskUrl_(t) { return t.web_url || t.url || t.task_url || t.jump_url || t.share_url || t.open_url || ''; }
function extractTimeMs_(timeField) {
  if (!timeField) return null;
  if (timeField.timestamp != null) {
    const raw = String(timeField.timestamp);
    if (/^\d{13}$/.test(raw)) return Number(raw);
    if (/^\d{10}$/.test(raw)) return Number(raw) * 1000;
    return Number(raw);
  }
  if (timeField.datetime) { const d = new Date(timeField.datetime); if (!isNaN(d)) return d.getTime(); }
  if (timeField.date) {
    const [y, m, d] = String(timeField.date).split('-').map(Number);
    return new Date(y, m - 1, d, 0, 0, 0).getTime();
  }
  return null;
}
function listMyTasks_(fromSec, toSec) {
  const origin = getOpenOrigin_();
  function makeUrl(pageSize, pageToken) {
    return `${origin}/open-apis/task/v2/tasks?type=my_tasks&page_size=${pageSize}` + (pageToken ? `&page_token=${encodeURIComponent(pageToken)}` : '');
  }
  let pageSize = 100, pageToken = '';
  const out = [], fs = fromSec * 1000, ts = toSec * 1000;

  while (true) {
    const url = makeUrl(pageSize, pageToken);
    const { http, json, text } = fetchJsonAny_(url, { method: 'get', headers: userAuthHeader_() });
    if (http !== 200 || json.code !== 0) {
      const msg = (json && json.msg) || text || '';
      if (/field validation failed/i.test(msg) && pageSize > 50) { pageSize = 50; pageToken = ''; continue; }
      throw new Error('TASKS_ERROR:' + msg);
    }
    const items = json.data.items || json.data.tasks || [];
    for (const t of items) {
      const tid = t.guid || t.task_guid || t.task_id || t.id;
      if (!tid) continue;
      const title = (t.summary && String(t.summary).trim()) || '（無題のタスク）';
      const description = t.description || '';

      let startMs = null;
      if (t.start) startMs = extractTimeMs_(t.start);
      else if (t.start_time) startMs = extractTimeMs_(t.start_time);
      else if (t.created_at) startMs = extractTimeMs_(t.created_at);

      let endMs = null;
      if (t.due) endMs = extractTimeMs_(t.due);
      else if (t.end) endMs = extractTimeMs_(t.end);
      else if (t.end_time) endMs = extractTimeMs_(t.end_time);
      else if (t.due_date) endMs = extractTimeMs_(t.due_date);

      if (startMs == null && endMs != null) startMs = endMs;
      if (endMs == null && startMs != null) endMs = startMs;

      if ((startMs == null && endMs == null) ||
          !((startMs && startMs >= fs && startMs <= ts) || (endMs && endMs >= fs && endMs <= ts))) {
        continue;
      }
      out.push({ event_id: `task:${tid}`, title, description, start_ms: startMs, end_ms: endMs, doc_url: extractTaskUrl_(t) || '' });
    }
    pageToken = json.data.page_token || '';
    if (!pageToken) break;
  }
  return out;
}

/* ===================== 7. 同期ロジック ======================== */
function createMinutesIfNeeded_(calId, eventId) {
  if (!CAL_CREATE_MINUTES) return '';
  try {
    const origin = getOpenOrigin_();
    const primaryId = getPrimaryCalendarId_();
    if (primaryId !== calId) return '';
    const js = fetchJson_(`${origin}/open-apis/calendar/v4/calendars/${calId}/events/${eventId}/meeting_minute`, {
      method: 'post', contentType: 'application/json', headers: userAuthHeader_(), payload: JSON.stringify({})
    });
    if (js.code !== 0) return '';
    return (js.data && js.data.doc_url) || '';
  } catch (_) { return ''; }
}
function calcHours_(sMs, eMs) {
  if (sMs == null || eMs == null) return null;
  const h = (eMs - sMs) / 3600000;
  if (!isFinite(h)) return null;
  return Math.round(h * 100) / 100;
}

// 指定カレンダー1つ同期
function runCalendarSyncOne_(calId, pastDays, futureDays) {
  const now = Date.now();
  let fromSec = Math.floor((now + (pastDays == null ? CAL_DEFAULT_PAST_DAYS : pastDays) * 86400 * 1000) / 1000);
  let toSec   = Math.floor((now + (futureDays == null ? CAL_DEFAULT_FUTURE_DAYS : futureDays) * 86400 * 1000) / 1000);
  if (fromSec >= toSec) { const t = fromSec; fromSec = toSec; toSec = t; }

  ensureSchema_();

  const calMeta = getCalendarInfo_(calId);
  const calTitle = calMeta.summary || '';

  Logger.log(`[SYNC] START cal=${calId} title=${calTitle} fromSec=${fromSec} toSec=${toSec}`);

  const events = getEvents_(calId, fromSec, toSec);
  const existingByEid = listAllRecords_Within_(fromSec, toSec, calId);

  Logger.log(`[SYNC] FETCH cal=${calId} events=${events.length} existing=${Object.keys(existingByEid).length}`);

  const nowMs = now;
  const creates = [], updates = [];
  let createdMinutes = 0;

  for (const ev of events) {
    let docUrl = '';
    if (CAL_CREATE_MINUTES) {
      const url = createMinutesIfNeeded_(calId, ev.event_id);
      if (url) { docUrl = url; createdMinutes++; }
    }

    const fields = {};
    fields[EN_FIELD.EVENT_ID]  = ev.event_id;
    fields[EN_FIELD.CAL_ID]    = calId;
    fields[EN_FIELD.CAL_TITLE] = calTitle;
    fields[EN_FIELD.TITLE]     = ev.title;
    if (ev.start_ms != null) fields[EN_FIELD.START] = ev.start_ms;
    if (ev.end_ms   != null) fields[EN_FIELD.END]   = ev.end_ms;
    fields[EN_FIELD.LAST] = nowMs;
    if (ev.description) fields[EN_FIELD.EXPL] = ev.description;

    const hours = calcHours_(ev.start_ms, ev.end_ms);
    if (hours != null) fields[EN_FIELD.TOTAL] = hours;

    if (docUrl) putDoc_(fields, docUrl, 'Minutes');

    const ex = existingByEid[String(ev.event_id)];
    if (ex) updates.push({ record_id: ex.record_id, fields });
    else creates.push({ fields });
  }

  let c = 0, u = 0, del = 0, softU = 0;
  if (creates.length) c = batchCreate_(creates);
  if (updates.length) u = batchUpdate_(updates);

  Logger.log(`[SYNC] UPSERT cal=${calId} created=${c} updated=${u} minutesCreated=${createdMinutes}`);

  // 差集合で削除判定（同期範囲内のみ）
  const seen = new Set(events.map(ev => String(ev.event_id)));
  const toDelete = Object.values(existingByEid).filter(r => !seen.has(String(r.event_id)));

  Logger.log(`[SYNC] DIFF cal=${calId} toDelete=${toDelete.length} sample=${toDelete.slice(0, 5).map(x => `${x.event_id}:${x.record_id}`).join(',')}`);

  if (toDelete.length) {
    if (CAL_SOFT_DELETE) {
      const soft = toDelete.map(r => ({
        record_id: r.record_id,
        fields: { [EN_FIELD.LAST]: nowMs, [SOFT_DELETE_FIELD]: true }
      }));

      Logger.log(`[SYNC] SOFT_DELETE cal=${calId} count=${soft.length} sample=${soft.slice(0, 5).map(x => x.record_id).join(',')}`);

      if (soft.length) softU = batchUpdate_(soft);

      Logger.log(`[SYNC] SOFT_DELETE_DONE cal=${calId} softUpdated=${softU}`);
    } else {
      const ids = toDelete.map(r => r.record_id);

      Logger.log(`[SYNC] DELETE cal=${calId} delete_ids=${ids.length} sample=${ids.slice(0, 5).join(',')}`);

      if (ids.length) del = batchDelete_(ids);

      Logger.log(`[SYNC] DELETE_DONE cal=${calId} deleted=${del}`);
    }
  }

  Logger.log(`[SYNC] END cal=${calId} created=${c} updated=${u} deleted=${del} softUpdated=${softU}`);

  return { created: c, updated: u, deleted: del, softUpdated: softU, events: events.length, minutesCreated: createdMinutes };
}

// タスク同期(選択時のみ呼ぶ)
function runTasksSync_(pastDays, futureDays) {
  const now = Date.now();
  let fromSec = Math.floor((now + (pastDays == null ? CAL_DEFAULT_PAST_DAYS : pastDays) * 86400 * 1000) / 1000);
  let toSec   = Math.floor((now + (futureDays == null ? CAL_DEFAULT_FUTURE_DAYS : futureDays) * 86400 * 1000) / 1000);
  if (fromSec >= toSec) { const t = fromSec; fromSec = toSec; toSec = t; }

  ensureSchema_();

  const tasks = listMyTasks_(fromSec, toSec);
  const existingByEid = listAllRecords_Within_(fromSec, toSec, MY_TASKS_ID);

  const nowMs = now;
  const creates = [], updates = [];

  for (const t of tasks) {
    const fields = {};
    fields[EN_FIELD.EVENT_ID]  = t.event_id;
    fields[EN_FIELD.CAL_ID]    = MY_TASKS_ID;
    fields[EN_FIELD.CAL_TITLE] = 'My tasks';
    fields[EN_FIELD.TITLE]     = t.title;
    if (t.start_ms != null) fields[EN_FIELD.START] = t.start_ms;
    if (t.end_ms   != null) fields[EN_FIELD.END]   = t.end_ms;
    fields[EN_FIELD.LAST] = nowMs;
    if (t.description) fields[EN_FIELD.EXPL] = t.description;
    const hours = calcHours_(t.start_ms, t.end_ms);
    if (hours != null) fields[EN_FIELD.TOTAL] = hours;
    if (t.doc_url) putDoc_(fields, t.doc_url, 'Task');

    const ex = existingByEid[String(t.event_id)];
    if (ex) updates.push({ record_id: ex.record_id, fields });
    else    creates.push({ fields });
  }

  let c = 0, u = 0, del = 0, softU = 0;
  if (creates.length) c = batchCreate_(creates);
  if (updates.length) u = batchUpdate_(updates);

  // 差集合で削除判定（再取得なし）
  const seen = new Set(tasks.map(ev => String(ev.event_id)));
  const toDelete = Object.values(existingByEid).filter(r => !seen.has(String(r.event_id)));
  if (toDelete.length) {
    if (CAL_SOFT_DELETE) {
      const soft = toDelete.map(r => ({
        record_id: r.record_id,
        fields: { [EN_FIELD.LAST]: nowMs, [SOFT_DELETE_FIELD]: true }
      }));
      if (soft.length) softU = batchUpdate_(soft);
    } else {
      const ids = toDelete.map(r => r.record_id);
      if (ids.length) del = batchDelete_(ids);
    }
  }

  Logger.log(`TASKS created=${c}, updated=${u}, deleted=${del}, softUpdated=${softU}, tasks=${tasks.length}`);
  return { created: c, updated: u, deleted: del, softUpdated: softU, tasks: tasks.length };
}

/* ===================== 8. スケジュール同期(設定反映) =========== */
function runScheduledSync() {
  try {
    const { selectedCalIds, pastDays, futureDays } = loadSyncSettings_();
    if (!selectedCalIds || !selectedCalIds.length) {
      Logger.log('保存された選択カレンダーがありません。同期スキップ');
      return;
    }
    const results = [];
    for (const calId of selectedCalIds) {
      if (calId === MY_TASKS_ID) {
        results.push({ calendar_id: calId, type: 'task', result: runTasksSync_(pastDays, futureDays) });
      } else {
        results.push({ calendar_id: calId, type: 'calendar', result: runCalendarSyncOne_(calId, pastDays, futureDays) });
      }
    }
    Logger.log('Scheduled sync OK: ' + JSON.stringify(results));
  } catch (e) {
    Logger.log('Scheduled sync NG: ' + (e && e.message ? e.message : e));
  }
}

/* ===================== 9. UI RPC ============================== */
// 準備チェック
function ensureSetupReady_() {
  const env = getEnv_();
  const lacks = [];
  if (!env['LARK_APP_ID']) lacks.push('LARK_APP_ID');
  if (!env['LARK_APP_SECRET']) lacks.push('LARK_APP_SECRET');
  try { getCfg_(); } catch (e) { lacks.push('LARK_BASE_ID/LARK_TABLE_ID/またはUI URL'); }
  const sp = PropertiesService.getScriptProperties();
  if (!sp.getProperty('USER_ACCESS_TOKEN')) lacks.push('USER_ACCESS_TOKEN（OAuth未実行）');
  if (lacks.length) throw new Error('セットアップ不足: ' + lacks.join(', '));
}

// ステータス
function uiGetStatus() {
  try {
    const cfg = getCfg_();
    const sp = PropertiesService.getScriptProperties();
    const now = Date.now();
    const exp = Number(sp.getProperty('USER_TOKEN_EXPIRE') || 0);
    const remainSec = exp ? Math.max(0, Math.floor((exp - now) / 1000)) : 0;
    const sync = loadSyncSettings_();
    return uiOk({
      baseId: cfg.baseId, tableId: cfg.tableId,
      hasUserToken: !!sp.getProperty('USER_ACCESS_TOKEN'),
      userTokenRemainSec: remainSec,
      origin: getOpenOrigin_(),
      saved: sync
    });
  } catch (e) { return uiErr(e); }
}

// 認可URL
function uiGetOAuthURL() {
  try { return uiOk({ url: getOAuthURL() }); }
  catch (e) { return uiErr(e); }
}

// カレンダー一覧（UIが呼ぶ旧関数名も用意）
function uiGetCalendars() {
  try {
    ensureSetupReady_();
    const list = listMyCalendars() || [];
    return uiOk({ calendars: list, defaults: { pastDays: Math.abs(CAL_DEFAULT_PAST_DAYS), futureDays: CAL_DEFAULT_FUTURE_DAYS } });
  } catch (e) { return uiErr(e); }
}
function getCalendarsForUI() { return uiGetCalendars(); }

// 設定の保存/読取
function uiSaveSyncSettings(calendarIds, pastDays, futureDays) {
  try {
    if (!Array.isArray(calendarIds)) throw new Error('calendarIdsは配列で指定してください。');
    const p = Number.isFinite(+pastDays) ? +pastDays : CAL_DEFAULT_PAST_DAYS;
    const f = Number.isFinite(+futureDays) ? +futureDays : CAL_DEFAULT_FUTURE_DAYS;
    saveSyncSettings_(calendarIds, p, f);
    return uiOk({ saved: { calendarIds, pastDays: p, futureDays: f } });
  } catch (e) { return uiErr(e); }
}
function uiLoadSyncSettings() {
  try { return uiOk(loadSyncSettings_()); } catch (e) { return uiErr(e); }
}

// 手動同期（UIが呼ぶ名前に合わせ、同時に設定も保存）
function syncSelectedCalendarsFromUI(calendarIds, opts) {
  try {
    ensureSetupReady_();
    const ids = Array.isArray(calendarIds) ? calendarIds : [];
    const p   = (opts && Number.isFinite(+opts.pastDays))   ? +opts.pastDays   : CAL_DEFAULT_PAST_DAYS;
    const f   = (opts && Number.isFinite(+opts.futureDays)) ? +opts.futureDays : CAL_DEFAULT_FUTURE_DAYS;
    if (!ids.length) throw new Error('カレンダーが選択されていません。');

    // 保存（→ トリガーが runScheduledSync で利用）
    saveSyncSettings_(ids, p, f);

    const out = [];
    for (const calId of ids) {
      try {
        if (calId === MY_TASKS_ID)
          out.push({ calendar_id: calId, ok: true, type: 'task',     result: runTasksSync_(p, f) });
        else
          out.push({ calendar_id: calId, ok: true, type: 'calendar', result: runCalendarSyncOne_(calId, p, f) });
      } catch (errOne) {
        const msg = String(errOne && errOne.message ? errOne.message : errOne);
        out.push({ calendar_id: calId, ok: false, error: msg });
      }
    }
    return uiOk({ results: out });
  } catch (e) { return uiErr(e); }
}

// 互換：新RPC名
function uiRunManualSync(calendarIds, pastDays, futureDays) {
  return syncSelectedCalendarsFromUI(calendarIds, { pastDays, futureDays });
}

/* ===================== 10. トリガー管理 ======================= */
// runScheduledSync を実行するハンドラー（UIはこの名前を使う）
function scheduledDaily()    { try { runScheduledSync(); } catch (e) { Logger.log('scheduledDaily NG: ' + (e.message || e)); } }
function frequentSync()      { try { runScheduledSync(); } catch (e) { Logger.log('frequentSync NG: ' + (e.message || e)); } }

// 保存用キー（時刻/分間隔の表示用）
const PROP_DAILY_H = 'DAILY_TRIGGER_HOUR';
const PROP_DAILY_M = 'DAILY_TRIGGER_MINUTE';
const PROP_FREQ_M  = 'FREQUENT_TRIGGER_MINUTES';

function uiGetDailyTriggerStatus() {
  try {
    const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'scheduledDaily');
    const sp = PropertiesService.getScriptProperties();
    return uiOk({
      status: {
        exists,
        hour: Number(sp.getProperty(PROP_DAILY_H) || ''),
        minute: Number(sp.getProperty(PROP_DAILY_M) || '')
      }
    });
  } catch (e) { return uiErr(e); }
}
function uiSetDailyTrigger(hour, minute) {
  try {
    ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'scheduledDaily') ScriptApp.deleteTrigger(t); });
    const h = Math.max(0, Math.min(23, Number(hour) || 3));
    const m = Math.max(0, Math.min(59, Number(minute) || 0));
    ScriptApp.newTrigger('scheduledDaily').timeBased().atHour(h).nearMinute(m).everyDays(1).create();
    const sp = PropertiesService.getScriptProperties();
    sp.setProperty(PROP_DAILY_H, String(h));
    sp.setProperty(PROP_DAILY_M, String(m));
    return uiOk({ message: `毎日 ${h}時${m}分に設定しました` });
  } catch (e) { return uiErr(e); }
}
function uiClearDailyTrigger() {
  try {
    let deleted = 0;
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === 'scheduledDaily') { ScriptApp.deleteTrigger(t); deleted++; }
    });
    const sp = PropertiesService.getScriptProperties();
    sp.deleteProperty(PROP_DAILY_H);
    sp.deleteProperty(PROP_DAILY_M);
    return uiOk({ message: `毎日トリガー ${deleted} 件削除` });
  } catch (e) { return uiErr(e); }
}

function uiGetFrequentTriggerStatus() {
  try {
    const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'frequentSync');
    const mins = Number(PropertiesService.getScriptProperties().getProperty(PROP_FREQ_M) || '');
    return uiOk({ status: { exists, minutes: Number.isFinite(mins) ? mins : null } });
  } catch (e) { return uiErr(e); }
}
function uiSetFrequentTrigger(minutes) {
  try {
    ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'frequentSync') ScriptApp.deleteTrigger(t); });
    const mins = Math.max(1, Math.min(60, Number(minutes) || 15));
    ScriptApp.newTrigger('frequentSync').timeBased().everyMinutes(mins).create();
    PropertiesService.getScriptProperties().setProperty(PROP_FREQ_M, String(mins));
    return uiOk({ message: `${mins}分間隔で設定しました` });
  } catch (e) { return uiErr(e); }
}
function uiClearFrequentTrigger() {
  try {
    let deleted = 0;
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === 'frequentSync') { ScriptApp.deleteTrigger(t); deleted++; }
    });
    PropertiesService.getScriptProperties().deleteProperty(PROP_FREQ_M);
    return uiOk({ message: `頻繁トリガー ${deleted} 件削除` });
  } catch (e) { return uiErr(e); }
}

/* ===================== 11. メニュー/サイドバー ================ */
function openCalendarPicker() {
  const html = HtmlService.createTemplateFromFile('cal_picker').evaluate()
    .setTitle('Lark セットアップ&同期')
    .setWidth(720);
  SpreadsheetApp.getUi().showSidebar(html);
}
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Lark設定')
      .addItem('セットアップ&同期', 'openCalendarPicker')
      .addToUi();
  } catch (_) {}
}

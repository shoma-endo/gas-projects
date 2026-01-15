/* ─── 0. 共通ヘルパ ─── */
function getSpreadId_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREAD_ID');
  if (!id) throw new Error('スクリプトプロパティ SPREAD_ID が未設定です');
  return id;
}

/** ヘッダー列番号マップ（5 分キャッシュ） */
function headerMap_() {
  const cache = CacheService.getScriptCache();
  const key   = 'propHeaderMap';
  const hit   = cache.get(key);
  if (hit) return JSON.parse(hit);

  const sheet   = SpreadsheetApp.openById(getSpreadId_()).getSheetByName('Property');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                       .getValues()[0]
                       .map(h => h.toString().trim());
  const map = headers.reduce((m, h, i) => { if (h) m[h] = i; return m; }, {});
  cache.put(key, JSON.stringify(map), 300);
  return map;
}

/** Property シートの任意行を配列で取得 */
function getPropertyRow_(rowNum = 2) {
  const sheet = SpreadsheetApp.openById(getSpreadId_()).getSheetByName('Property');
  return sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/* ─── 1. エントリポイント ─── */
function doGet(e) {
  if (e.parameter.app === 'migrate' && e.parameter.code) {
    return handleMigration(e);
  }
  return HtmlService
    .createTemplateFromFile('MigrationPage')
    .evaluate()
    .setTitle('LINEアカウント移行')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents || '{}');
  } catch {
    return createJsonResponse({ status: 'error', message: 'invalid json' });
  }

  if (Array.isArray(body.events)) {
    return safeExec(() => handleLineWebhook(body.events), 'line');
  }
  if (body.event === 'bitable_record_changed') {
    normalizeLarkKeys(body);
    return safeExec(() => handleAutomation(body), 'lark');
  }
  if (body.challenge) {
    return ContentService.createTextOutput(body.challenge)
                         .setMimeType(ContentService.MimeType.TEXT);
  }
  return createJsonResponse({ status: 'ignored', reason: 'unrecognized payload' });
}

/* ─── 2. LINE Bot Webhook ─── */
function handleLineWebhook(events) {
  const map = headerMap_();
  const row = getPropertyRow_(2);
  const env = {
    lineToken    : row[map['lineToken']],
    baseAppToken : row[map['baseAppToken']],
    baseTableId  : row[map['baseTableId']]
  };

  events.forEach(evt => {
    if (evt.type === 'follow') {
      const uid  = evt.source.userId;
      const prof = getLineUserProfile(env.lineToken, uid);
      upsertUserToBase(
        env,
        { userId: uid, displayName: prof.displayName }
      );
    }
  });
}

/* ─── 3. Lark → LINE 送信 ─── */
function handleAutomation(p) {
  const map = headerMap_();
  const row = getPropertyRow_(2);
  const env = {
    lineToken    : row[map['lineToken']],
    baseAppToken : row[map['baseAppToken']],
    baseTableId  : row[map['baseTableId']]
  };

  const msg = (p.record.fields?.MessageText || '').trim();
  if (!msg) {
    setStatus(env, p, 'Message空');
    return;
  }

  const uid = p.record.fields?.LineUserId;
  if (uid) {
    const ok = pushSafe(env.lineToken, uid, msg);
    setStatus(env, p, ok ? '送信済み' : '送信失敗');
    return;
  }

  let ids = [];
  if (p.userTableId) {
    try {
      ids = fetchIdsFromBase(env.baseAppToken, p.userTableId);
    } catch(e) {
      console.error(e);
    }
  }
  if (!ids.length) {
    setStatus(env, p, '対象なし');
    return;
  }

  const CHUNK = 500;
  let allOK = true;
  for (let i = 0; i < ids.length; i += CHUNK) {
    if (!multi(env.lineToken, ids.slice(i, i + CHUNK), msg)) allOK = false;
  }
  setStatus(env, p, allOK ? '送信済み' : '一部失敗');
}

/* ─── 4. LINE Login (UI & Callback) ─── */
function makeAuthUrl() {
  const map = headerMap_();
  const row = getPropertyRow_(2);
  const c   = {
    clientKey   : row[map['channel_ID']],
    loginSecret : row[map['AccessToken']].toString().trim().replace(/^"+|"+$/g,''),
    baseAppToken: row[map['baseAppToken']],
    baseTableId : row[map['baseTableId']]
  };

  const state = Utilities.getUuid();
  CacheService.getScriptCache().put(state, '1', 300);
  const redirect = encodeURIComponent(
    ScriptApp.getService().getUrl() + '?app=migrate'
  );

  return 'https://access.line.me/oauth2/v2.1/authorize?response_type=code' +
         `&client_id=${c.clientKey}&redirect_uri=${redirect}` +
         `&state=${state}&scope=openid%20profile`;
}

function handleMigration(e) {
  const { code, state } = e.parameter;
  if (!code || CacheService.getScriptCache().get(state) !== '1') {
    return HtmlService.createHtmlOutput('不正なリクエストです');
  }

  const map = headerMap_();
  const row = getPropertyRow_(2);
  const env = {
    clientKey   : row[map['channel_ID']],
    loginSecret : row[map['AccessToken']].toString().trim().replace(/^"+|"+$/g,''),
    baseAppToken: row[map['baseAppToken']],
    baseTableId : row[map['baseTableId']]
  };

  const tokenRes = UrlFetchApp.fetch(
    'https://api.line.me/oauth2/v2.1/token', {
      method      : 'post',
      contentType : 'application/x-www-form-urlencoded',
      payload     : [
        'grant_type=authorization_code',
        `code=${encodeURIComponent(code)}`,
        `redirect_uri=${encodeURIComponent(ScriptApp.getService().getUrl() + '?app=migrate')}`,
        `client_id=${env.clientKey}`,
        `client_secret=${env.loginSecret}`
      ].join('&'),
      muteHttpExceptions: true
    }
  );
  const accessToken = JSON.parse(tokenRes.getContentText()).access_token;
  if (!accessToken) {
    return HtmlService.createHtmlOutput(
      'アクセストークン取得失敗<br><pre>' + tokenRes.getContentText() + '</pre>'
    );
  }

  const profile = JSON.parse(
    UrlFetchApp.fetch('https://api.line.me/v2/profile', {
      headers: { Authorization: `Bearer ${accessToken}` }
    }).getContentText()
  );

  upsertUserToBase(
    { baseAppToken: env.baseAppToken, baseTableId: env.baseTableId },
    { userId: profile.userId, displayName: profile.displayName }
  );

  return HtmlService.createHtmlOutput(
    `<p>移行が完了しました！<br>ID: ${profile.userId}<br>名前: ${profile.displayName}</p>`
  );
}

/* ─── 5. Status 更新 ─── */
function setStatus(env, p, status) {
  const acc = baseTok();
  const url = `https://open.larksuite.com/open-apis/bitable/v1/apps/${env.baseAppToken}` +
              `/tables/${p.tableId}/records/${p.record.recordId}`;
  lark({
    url: url,
    method: 'put',
    headers: { Authorization: `Bearer ${acc}` },
    payload: { fields: { Status: status } }
  });
}

/* ─── 6. Lark API Token ─── */
function baseTok() {
  const prop = PropertiesService.getScriptProperties();
  const t    = prop.getProperty('BASE_ACCESS_TOKEN');
  const exp  = Number(prop.getProperty('BASE_TOKEN_EXPIRE') || 0);
  if (t && Date.now() < exp - 60000) return t;

  const res = UrlFetchApp.fetch(
    'https://open.larksuite.com/open-apis/auth/v3/app_access_token/internal/', {
      method      : 'post',
      contentType : 'application/json',
      payload     : JSON.stringify({
        app_id     : prop.getProperty('APP_ID'),
        app_secret : prop.getProperty('APP_SECRET')
      })
    }
  );
  const j = JSON.parse(res.getContentText());
  prop.setProperty('BASE_ACCESS_TOKEN', j.app_access_token);
  prop.setProperty('BASE_TOKEN_EXPIRE', String(Date.now() + j.expire * 1000));
  return j.app_access_token;
}
function lark(opt) {
  const r = UrlFetchApp.fetch(opt.url, {
    method            : opt.method || 'get',
    contentType       : 'application/json',
    headers           : opt.headers || {},
    payload           : opt.payload ? JSON.stringify(opt.payload) : undefined,
    muteHttpExceptions: true
  });
  const j = JSON.parse(r.getContentText());
  if (!String(r.getResponseCode()).startsWith('2') || j.code !== 0) {
    throw j.msg || r.getContentText();
  }
  return j;
}

/* ─── 7. LINE API Helpers ─── */
function getLineUserProfile(tok, uid) {
  const r = UrlFetchApp.fetch(
    `https://api.line.me/v2/bot/profile/${uid}`, {
      headers: { Authorization: `Bearer ${tok}` },
      muteHttpExceptions: true
    }
  );
  if (r.getResponseCode() !== 200) throw new Error('LINE profile error');
  return JSON.parse(r.getContentText());
}
function pushSafe(tok, to, txt) {
  const r = UrlFetchApp.fetch(
    'https://api.line.me/v2/bot/message/push', {
      method      : 'post',
      contentType : 'application/json',
      headers     : { Authorization: `Bearer ${tok}` },
      payload     : JSON.stringify({ to, messages: [{ type: 'text', text: txt }] }),
      muteHttpExceptions: true
    }
  );
  return String(r.getResponseCode()).startsWith('2');
}
function multi(tok, arr, txt) {
  const r = UrlFetchApp.fetch(
    'https://api.line.me/v2/bot/message/multicast', {
      method      : 'post',
      contentType : 'application/json',
      headers     : { Authorization: `Bearer ${tok}` },
      payload     : JSON.stringify({ to: arr, messages: [{ type: 'text', text: txt }] }),
      muteHttpExceptions: true
    }
  );
  return String(r.getResponseCode()).startsWith('2');
}

/* ─── 8. Bitable Helpers ─── */
function upsertUserToBase(env, p) {
  const acc    = baseTok();
  const filter = encodeURIComponent(`CurrentValue.[LINEユーザーID]="${p.userId}"`);
  const res    = lark({
    url: `https://open.larksuite.com/open-apis/bitable/v1/apps/${env.baseAppToken}` +
         `/tables/${env.baseTableId}/records?filter=${filter}`,
    headers: { Authorization: `Bearer ${acc}` }
  });
  const rec = res.data.records?.[0];
  const baseUrl = `https://open.larksuite.com/open-apis/bitable/v1/apps/${env.baseAppToken}` +
                  `/tables/${env.baseTableId}/records`;
  const body = { fields: { 'LINEユーザーID': p.userId, 'プロフィール表示名': p.displayName } };
  lark({
    url: rec ? `${baseUrl}/${rec.recordId}` : baseUrl,
    method: rec ? 'put' : 'post',
    headers: { Authorization: `Bearer ${acc}` },
    payload: body
  });
}
function fetchIdsFromBase(appTok, tblId) {
  const res = lark({
    url: `https://open.larksuite.com/open-apis/bitable/v1/apps/${appTok}` +
         `/tables/${tblId}/records`,
    headers: { Authorization: `Bearer ${baseTok()}` }
  });
  return [...new Set(
    (res.data.items || res.data.records || [])
      .map(r => r.fields && r.fields['LINEユーザーID'])
      .filter(Boolean)
  )];
}

/* ─── 9. 汎用 ─── */
function normalizeLarkKeys(o) {
  o.appToken        = o.app_token     || o.appToken;
  o.tableId         = o.table_id      || o.tableId;
  o.userTableId     = o.user_table_id || o.userTableId || o.UserTableId;
  o.record          = o.record || {};
  o.record.recordId = o.record.record_id || o.record.recordId;
}
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}
function safeExec(fn, src) {
  try {
    fn();
    return createJsonResponse({ status: 'ok', source: src });
  } catch (e) {
    console.error(e);
    return createJsonResponse({ status: 'error', message: String(e), source: src });
  }
}
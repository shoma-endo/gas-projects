/* ===================== 1. 定数 ============================== */
/** スクリプトプロパティのキー*/
const PROP = {
  LARK_BASE_ID: 'LARK_BASE_ID',
  LARK_TABLE_ID: 'LARK_TABLE_ID'
};

/**シート/列名など*/
const WEBHOOK_SHEET_NAME = 'Webhook受信';
const STATUS_HEADER_NAME = '送信済みフラグ';
const JSON_TYPE_HEADER_NAME = 'jsonデータの形式';                        // GET/POST を記録
const JSON_DATA_HEADER_NAME = 'jsonデータ（プロラインから受け取ったデータ）'; // json形式
const JSON_DATA_HEADER_ALIASES = [JSON_DATA_HEADER_NAME, 'jsonデータ'];
const OPTIONAL_RECORD_ID_COL = 'LarkRecordID';
const ERROR_DETAIL_COL = 'エラー詳細';

/**Lark/動作パラメータ*/
const LARK_DOMAIN = 'https://open.larksuite.com';
const MAX_ROWS_PER_DISPATCH = 100;
const MAX_RETRY = 4;
const BACKOFF_INITIAL_MS = 500;
const LOCK_TIMEOUT_MS = 10000;
const FIELD_CACHE_TTL_SEC = 300;

/** 受信日時ラベル */
const FRONTEND_DATA_KEYS = { DATE_HEADER: '受信日時' };

/** シート保存時のJSONサイズ上限（長大JSONの切り詰め用）*/
const JSON_SAVE_MAXLEN = 40000;

/** 重複ガード TTL（秒）: 6時間 */
const DEDUP_TTL_SEC = 21600;


/* ===================== 2. ヘルパー =========================== */
/** 'yyyy-MM-dd HH:mm:ss' で時刻整形 */
function fmtTs_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/** JSONレスポンス */
function json_(o, c) {
  if (c != null) o.statusCode = c;
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/** ログ */
function logErr_(tag, e) {
  Logger.log(tag + ':' + (e && e.stack || e));
}

/** シート保存用に文字列化しつつ長大JSONを丸める */
function stringifyForSheet_(o) {
  let s = JSON.stringify(o);
  if (s.length > JSON_SAVE_MAXLEN) s = s.slice(0, JSON_SAVE_MAXLEN - 20) + '...truncated';
  return s;
}

/** 安定化 stringify（キーをソート） */
function stableStringify_(v) {
  if (v === null || typeof v !== 'object') return JSON.stringify(v);
  if (Array.isArray(v)) return '[' + v.map(stableStringify_).join(',') + ']';
  const keys = Object.keys(v).sort();
  return '{' + keys.map(k => JSON.stringify(k) + ':' + stableStringify_(v[k])).join(',') + '}';
}

/** SHA-256 hex */
function sha256Hex_(s) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s);
  return bytes.map(b => ((b + 256) % 256).toString(16).padStart(2,'0')).join('');
}

/** 角括弧をドットに（例: user_data[passcode] => user_data.passcode） */
function sanitizeKey(k) {
  return String(k).replace(/\[([^\]]+)\]/g, '.$1');
}

/**
 * 送信先の列名を決める。
 * - 角括弧キー（xxx[leaf]）が来たら、Baseに leaf 列が既存なら leaf を優先（例: user_data[linename] -> linename）。
 * - それ以外（既存が無ければ）は sanitize した名前を使う（例: user_data[passcode] -> user_data.passcode）。
 *   ※ form_data[...] など他プレフィックスでも同様に leaf 優先。
 */
function pickFieldName(k, allowedSet) {
  const m = /^(.+?)\[(.+?)\]$/.exec(String(k));
  if (m) {
    const leaf = m[2];
    if (allowedSet.has(leaf)) return leaf;
  }
  return sanitizeKey(k);
}


/* ===================== 3. URL解析 ============================ */
function extractBaseAndTableIds(url) {
  if (!url) return { baseId: '', tableId: '' };
  return {
    baseId: (url.match(/\/base\/([A-Za-z0-9]{25,30})/) || [])[1] || '',
    tableId:
      (url.match(/[?&]table=(tbl[A-Za-z0-9]+)/) ||
       url.match(/(tbl[A-Za-z0-9]{14,20})/) ||
       [])[1] || ''
  };
}

/** URLから地域を判定、共有配布用ロジック*/
function getApiDomainFromUrl_(url) {
  if (!url) return LARK_DOMAIN;
  const u = String(url).toLowerCase();
  if (u.includes('feishu.cn') || u.includes('larkoffice.com')) return 'https://open.feishu.cn';
  return LARK_DOMAIN;
}


/* ===================== 4. Webhook受信 ======================== */
function doPost(e) { return handleReq_(e, 'POST'); }
function doGet(e)  { return handleReq_(e, 'GET'); }

function handleReq_(e, method) {
  try {
    const paramsRaw = (e && e.parameter) ? Object.assign({}, e.parameter) : {};
    const m = (method || '').toLowerCase();

    let data = {};
    if (m === 'get') {
      data = paramsRaw;
    } else if (m === 'post') {
      let bodyRaw = {};
      if (e && e.postData && typeof e.postData.contents === 'string') {
        const ct = (e.postData.type || '').toLowerCase();
        const contents = e.postData.contents;
        const looksJson = /^[\s]*[{[]/.test(contents);
        if ((/json/.test(ct) || looksJson)) {
          try { bodyRaw = JSON.parse(contents || '{}'); } catch (ignore) {}
        } else if (/application\/x-www-form-urlencoded|multipart\/form-data/.test(ct)) {
          bodyRaw = paramsRaw;
        } else {
          bodyRaw = paramsRaw;
        }
      } else {
        bodyRaw = paramsRaw;
      }
      data = flattenToOneLevel_(bodyRaw);
    } else {
      data = paramsRaw;
    }

    // 最小補完: ユーザーID/LINE登録名/シナリオ（既存があれば上書きしない）
    data = enrichCommonAliases_(data);

    handleWebhook_({ dataObj: data, httpMethod: m });
    dispatchPending(); // 公開関数（手動実行も可）

    return json_({ success: true });
  } catch (err) {
    logErr_('handleReq', err);
    return json_({ success: false, message: String(err) }, 500);
  }
}


/* ===================== 5. 受信データ→シート ================== */
function handleWebhook_(_param) {
  const dataObj = _param.dataObj;
  const method = _param.httpMethod || '';
  const pair = getOrCreateSheet_(WEBHOOK_SHEET_NAME);
  const sheet = pair.sheet;
  const headers = pair.headers;

  ensureEssentialColumns_(sheet, headers);
  const jsonHeaderName = getExistingHeaderName_(headers, JSON_DATA_HEADER_ALIASES) || JSON_DATA_HEADER_NAME;

  const ts = fmtTs_(new Date());

  // JSON列に保存するペイロードにも受信日時を同梱（Base送信時の欠落防止）
  const payloadForJson = Object.assign({}, dataObj);
  payloadForJson[FRONTEND_DATA_KEYS.DATE_HEADER] = ts;
  const jsonStr = stringifyForSheet_(payloadForJson);

  const row = headers.map(function(h) {
    if (h === FRONTEND_DATA_KEYS.DATE_HEADER) return ts;
    if (h === JSON_TYPE_HEADER_NAME)         return method.toUpperCase(); // GET/POST
    if (h === 'ユーザーID')                   return dataObj['ユーザーID'] || '';
    if (h === 'LINE登録名')                   return dataObj['LINE登録名'] || '';
    if (h === 'シナリオ')                     return dataObj['シナリオ']   || '';
    if (h === jsonHeaderName)                return jsonStr;
    if (h === STATUS_HEADER_NAME)            return '未送信';
    return '';
  });

  sheet.appendRow(row);
}


/* ===================== 6. ディスパッチ(同期) ================= */
function dispatchPending() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_TIMEOUT_MS)) return;

  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEBHOOK_SHEET_NAME);
    if (!sh) return;

    const meta = getSheetAndHeaders_(WEBHOOK_SHEET_NAME);
    if (!meta || meta.lastRow < 2) return;

    const headers = meta.headers;
    const lastRow = meta.lastRow;

    const statusCol   = headers.indexOf(STATUS_HEADER_NAME) + 1;
    const jsonHeaderName = getExistingHeaderName_(headers, JSON_DATA_HEADER_ALIASES) || JSON_DATA_HEADER_NAME;
    const jsonCol     = headers.indexOf(jsonHeaderName) + 1;
    const recordIdCol = headers.indexOf(OPTIONAL_RECORD_ID_COL) + 1;
    const errorCol    = headers.indexOf(ERROR_DETAIL_COL) + 1;
    const dateCol     = headers.indexOf(FRONTEND_DATA_KEYS.DATE_HEADER) + 1;

    const statusVals = sh.getRange(2, statusCol, lastRow - 1, 1).getValues();
    const jsonVals   = sh.getRange(2, jsonCol,   lastRow - 1, 1).getValues();
    const dateVals   = dateCol ? sh.getRange(2, dateCol, lastRow - 1, 1).getValues() : [];

    const targets = [];
    for (let i = 0; i < statusVals.length; i++) {
      const pending = statusVals[i][0] === '' || statusVals[i][0] === '未送信';
      if (pending && jsonVals[i][0]) {
        targets.push({
          row: i + 2,
          json: jsonVals[i][0],
          ts:  dateCol ? (dateVals[i][0] || '') : ''
        });
      }
      if (targets.length >= MAX_ROWS_PER_DISPATCH) break;
    }
    if (!targets.length) return;

    const cfg = getCfg_();
    const tok = getLarkTenantAccessToken_(cfg);

    const allowedSet = new Set(getLarkFieldNamesCached_(tok, cfg));

    // 重複ガード用キャッシュ
    const cache = CacheService.getScriptCache();

    for (let tIdx = 0; tIdx < targets.length; tIdx++) {
      const t = targets[tIdx];
      let status = '送信エラー';
      try {
        const raw = JSON.parse(t.json);

        // 受信日時が無ければシートの受信日時で補完（Baseへも送るため）
        if (!Object.prototype.hasOwnProperty.call(raw, FRONTEND_DATA_KEYS.DATE_HEADER) ||
            String(raw[FRONTEND_DATA_KEYS.DATE_HEADER] || '').trim() === '') {
          raw[FRONTEND_DATA_KEYS.DATE_HEADER] = t.ts || '';
        }

        // --- 重複ガード（id 優先 → 無い場合は内容ハッシュ） ---
        const eventId = String(raw.id || '').trim();
        const dedupKey = eventId
          ? ('dup:' + cfg.baseId + ':' + cfg.tableId + ':id:' + eventId)
          : (() => {
              const copy = Object.assign({}, raw);
              delete copy[FRONTEND_DATA_KEYS.DATE_HEADER];
              delete copy['http_method'];
              return 'dup:' + cfg.baseId + ':' + cfg.tableId + ':hash:' + sha256Hex_(stableStringify_(copy));
            })();

        if (cache.get(dedupKey)) {
          status = '重複スキップ';
          if (statusCol) sh.getRange(t.row, statusCol).setValue(status);
          continue;
        }

        // 念のためもう一度1階層化（POSTなら冪等、GETはそのままでもOK）
        const flat = flattenToOneLevel_(raw);

        // 値の最小整形のみ。フィルタ/マッピングは行わない（キー名はそのまま）
        const fieldsCandidate = toBaseFields_(flat);

        // --- 最小パッチ：角括弧キーのマッピング & 自動作成 ---
        const filteredPair = filterByAllowedFieldsWithAutoCreate_(fieldsCandidate, allowedSet, tok, cfg);
        const fields = filteredPair[0];

        // 追加: 可視化(未作成/除外フィールド)
        const ngFields = Object.keys(filteredPair[1] || {});
        if (ngFields.length && errorCol) {
          sh.getRange(t.row, errorCol).setValue('未作成/除外フィールド: ' + ngFields.join(', '));
        }

        if (!Object.keys(fields).length) throw new Error('no valid fields');

        const recordId = addWithRetry_(tok, cfg, fields);
        status = '送信済み';
        if (recordIdCol) sh.getRange(t.row, recordIdCol).setValue(recordId);

        // 重複キーを登録
        cache.put(dedupKey, '1', DEDUP_TTL_SEC);

      } catch (e) {
        if (errorCol) sh.getRange(t.row, errorCol).setValue(String(e));
      }
      if (statusCol) sh.getRange(t.row, statusCol).setValue(status);
    }
  } finally {
    lock.releaseLock();
  }
}


/* ===================== 7. Lark接続ユーティリティ ============= */
/** 設定シートから直接読み取りつつ、抽出した baseId/tableId を毎回スクリプトプロパティへ保存。
 *  抽出失敗時はプロパティ値をフォールバック利用。*/
function getCfg_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('設定');
  if (!sh) throw new Error('設定シートなし');

  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = function(h) {
    const idx = hdr.indexOf(h);
    if (idx === -1) throw new Error('設定シートに列が見つかりません: ' + h);
    return idx;
  };
  const row = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0];

  const url = String(row[col('LARK_URL')] || '').trim();
  const appId = String(row[col('LARK_APP_ID')] || '').trim();
  const appSecret = String(row[col('LARK_APP_SECRET')] || '').trim();

  // URL から抽出
  let ids = extractBaseAndTableIds(url);
  let baseId = ids.baseId || '';
  let tableId = ids.tableId || '';

  const sp = PropertiesService.getScriptProperties();

  // 抽出できたら毎回スクリプトプロパティへ書き込む
  if (baseId && tableId) {
    try {
      sp.setProperties({
        [PROP.LARK_BASE_ID]: baseId,
        [PROP.LARK_TABLE_ID]: tableId
      }, true);
    } catch (e) {
      Logger.log('setProperties failed: ' + e);
    }
  } else {
    // フォールバック：以前に保存した値を利用
    baseId = baseId || sp.getProperty(PROP.LARK_BASE_ID) || '';
    tableId = tableId || sp.getProperty(PROP.LARK_TABLE_ID) || '';
  }

  const domain = getApiDomainFromUrl_(url);

  if (!appId || !appSecret || !baseId || !tableId) {
    throw new Error('設定が不完全です(LARK_URL/LARK_APP_ID/LARK_APP_SECRET/baseId/tableIdのいずれかが空)');
  }

  return {
    appId: appId,
    appSecret: appSecret,
    baseId: baseId,
    tableId: tableId,
    domain: domain
  };
}

function getLarkTenantAccessToken_(_cfg) {
  const r = UrlFetchApp.fetch(_cfg.domain + '/open-apis/auth/v3/tenant_access_token/internal', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ app_id: _cfg.appId, app_secret: _cfg.appSecret })
  });
  const o = JSON.parse(r.getContentText());
  if (o.code !== 0) throw new Error(o.msg || o.code);
  return o.tenant_access_token;
}

function createLarkField_(tok, cfg, fieldName) {
  const payload = {
    field_name: fieldName,
    type: 1, // テキスト型
    property: {}
  };
  const r = UrlFetchApp.fetch(
    cfg.domain + '/open-apis/bitable/v1/apps/' + cfg.baseId + '/tables/' + cfg.tableId + '/fields',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + tok },
      payload: JSON.stringify(payload)
    }
  );
  const o = JSON.parse(r.getContentText());
  if (o.code !== 0) throw new Error('フィールド作成失敗: ' + fieldName + ' - ' + (o.msg || o.code));
  Logger.log('フィールド作成成功: ' + fieldName);
  return o.data.field;
}

function addWithRetry_(tok, cfg, fields) {
  let wait = BACKOFF_INITIAL_MS;
  for (let i = 0; i <= MAX_RETRY; i++) {
    try {
      return addRecord_(tok, cfg, fields);
    } catch (e) {
      if (i === MAX_RETRY) throw e;
      Utilities.sleep(wait);
      wait = wait * 2;
    }
  }
  throw new Error('unreachable');
}

function addRecord_(tok, cfg, fields) {
  const r = UrlFetchApp.fetch(
    cfg.domain + '/open-apis/bitable/v1/apps/' + cfg.baseId + '/tables/' + cfg.tableId + '/records',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + tok },
      payload: JSON.stringify({ fields: fields })
    }
  );
  const o = JSON.parse(r.getContentText());
  if (o.code !== 0) throw new Error(o.msg || o.code);
  return o.data.record.record_id;
}

function getLarkFieldNamesCached_(tok, cfg) {
  const cache = CacheService.getScriptCache();
  const key = 'fields:' + cfg.baseId + ':' + cfg.tableId;
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const res = UrlFetchApp.fetch(
    cfg.domain + '/open-apis/bitable/v1/apps/' + cfg.baseId + '/tables/' + cfg.tableId + '/fields?page_size=200',
    { headers: { Authorization: 'Bearer ' + tok } }
  );
  const o = JSON.parse(res.getContentText());
  if (o.code !== 0) throw new Error(o.msg || o.code);

  const items = o.data && o.data.items ? o.data.items : [];
  const names = items.map(function(f) { return f.field_name; });
  cache.put(key, JSON.stringify(names), FIELD_CACHE_TTL_SEC);
  return names;
}

function refreshLarkFieldNames_(tok, cfg) {
  const cache = CacheService.getScriptCache();
  const key = 'fields:' + cfg.baseId + ':' + cfg.tableId;
  cache.remove(key);
  return getLarkFieldNamesCached_(tok, cfg);
}

/** 許可されたフィールドだけに絞り込む + 未存在フィールドの自動作成（角括弧キー対応版） */
function filterByAllowedFieldsWithAutoCreate_(obj, allowedSet, tok, cfg) {
  const ok = {};
  const ng = {};
  const created = [];

  for (const k in obj) {
    const target = pickFieldName(k, allowedSet);
    if (allowedSet.has(target)) {
      ok[target] = obj[k];
      continue;
    }
    try {
      createLarkField_(tok, cfg, target);
      created.push(target);
      allowedSet.add(target);
      ok[target] = obj[k];
    } catch (e) {
      Logger.log('フィールド作成失敗: ' + target + ' - ' + String(e));
      ng[target] = obj[k];
    }
  }

  if (created.length > 0) {
    Logger.log('新規作成フィールド: ' + created.join(', '));
    refreshLarkFieldNames_(tok, cfg);
  }

  return [ok, ng];
}


/* ===================== 8. 汎用フラット化 =================== */
/** ネストJSONをドット区切りキーの1階層オブジェクトに変換（配列はJSON文字列化） */
function flattenToOneLevel_(obj) {
  const out = {};
  function rec(cur, path) {
    if (cur != null && typeof cur === 'object' && !Array.isArray(cur)) {
      let hasProp = false;
      for (const k in cur) {
        hasProp = true;
        rec(cur[k], path ? path + '.' + k : k);
      }
      if (!hasProp && path) out[path] = '{}';
    } else if (Array.isArray(cur)) {
      out[path] = JSON.stringify(cur);
    } else {
      out[path] = cur;
    }
  }
  if (obj && typeof obj === 'object') {
    for (const k in obj) rec(obj[k], k);
  }
  return out;
}

/** Base送信用に値を最小整形（オブジェクト/配列はJSON文字列化。null→空文字） */
function toBaseFields_(flat) {
  const out = {};
  for (const k in flat) {
    let v = flat[k];
    if (v === undefined) continue;
    if (v == null) { out[k] = ''; continue; }
    if (typeof v === 'object') v = JSON.stringify(v);
    out[k] = String(v);
  }
  return out;
}

/** 最小補完: ユーザーID/LINE登録名/シナリオ（既存があれば上書きしない） */
function enrichCommonAliases_(flat) {
  const out = Object.assign({}, flat);
  function pick(keys) {
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      if (Object.prototype.hasOwnProperty.call(flat, k)) {
        const v = flat[k];
        if (v != null && String(v).trim() !== '') return v;
      }
    }
    return null;
  }
  if (!Object.prototype.hasOwnProperty.call(out, 'ユーザーID')) {
    const v1 = pick(['uid','user_id','user_data.uid','user_data[uid]']);
    if (v1 != null) out['ユーザーID'] = v1;
  }
  if (!Object.prototype.hasOwnProperty.call(out, 'LINE登録名')) {
    const v2 = pick(['linename','user_data.linename','user_data[linename]','LINE名']);
    if (v2 != null) out['LINE登録名'] = v2;
  }
  if (!Object.prototype.hasOwnProperty.call(out, 'シナリオ')) {
    const v3 = pick(['event','form_name','custom_event']);
    if (v3 != null) out['シナリオ'] = v3;
  }
  return out;
}


/* ===================== 9. シート&汎用 ======================= */
function getExistingHeaderName_(headers, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const target = aliases[i];
    for (let j = 0; j < headers.length; j++) {
      const h = String(headers[j]).trim();
      if (h === JSON_TYPE_HEADER_NAME) continue; // 形式列はスキップ
      if (h === target) return h;                // 完全一致のみ
    }
  }
  return null;
}

function ensureHeader_(sh, headers, name) {
  if (headers.indexOf(name) === -1) {
    sh.insertColumnAfter(headers.length || 1);
    sh.getRange(1, headers.length + 1).setValue(name);
    headers.push(name);
  }
  return name;
}

function ensureHeaderWithAliases_(sh, headers, canonical, aliases) {
  const existing = getExistingHeaderName_(headers, aliases);
  if (existing) return existing;
  return ensureHeader_(sh, headers, canonical);
}

/** 対象シートの取得（なければ作成・ヘッダ初期化） */
function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  if (sh.getLastRow() < 1) {
    const hd = [
      '受信日時',
      'ユーザーID',
      'LINE登録名',
      'シナリオ',
      JSON_TYPE_HEADER_NAME,
      JSON_DATA_HEADER_NAME,
      STATUS_HEADER_NAME,
      ERROR_DETAIL_COL,
      OPTIONAL_RECORD_ID_COL
    ];
    sh.appendRow(hd);
    sh.getRange(1, 1, 1, hd.length).setBackground('#FFBB66');
    return { sheet: sh, headers: hd };
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(h) { return String(h).trim(); });
  return { sheet: sh, headers: headers };
}

/** シートとヘッダ情報の取得 */
function getSheetAndHeaders_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) return null;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(h) { return String(h).trim(); });
  return { headers: headers, lastRow: sh.getLastRow() };
}

/** Base 側の列名をログ出力（手動実行用公開関数） */
function listLarkFieldNames() {
  const cfg = getCfg_();
  const tok = getLarkTenantAccessToken_(cfg);
  const names = getLarkFieldNamesCached_(tok, cfg);
  Logger.log('FIELDS: ' + JSON.stringify(names));
}

/** 必須列が無ければ追加（エイリアス対応でjson列を保証） */
function ensureEssentialColumns_(sh, headers) {
  ensureHeader_(sh, headers, FRONTEND_DATA_KEYS.DATE_HEADER);
  ensureHeaderWithAliases_(sh, headers, JSON_DATA_HEADER_NAME, JSON_DATA_HEADER_ALIASES);
  ensureHeader_(sh, headers, JSON_TYPE_HEADER_NAME);
  ensureHeader_(sh, headers, 'ユーザーID');
  ensureHeader_(sh, headers, 'LINE登録名');
  ensureHeader_(sh, headers, 'シナリオ');
  ensureHeader_(sh, headers, STATUS_HEADER_NAME);
  ensureHeader_(sh, headers, ERROR_DETAIL_COL);
  ensureHeader_(sh, headers, OPTIONAL_RECORD_ID_COL);
}
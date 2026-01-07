const CONFIG = { 
  SPREADSHEET_ID: '', // 空なら自動検出。必要なら直書きOK
  SHEET_LOGS: 'logs',
  TIMEZONE: 'Asia/Tokyo',
  DATETIME_FORMAT: 'yyyy-MM-dd HH:mm:ss',
  TEST_MODE: false // trueにすると、パラメータなしでもテスト用デフォルト値で動作
};

const HEADERS = ['日時', '処理', '表示名'];
const ACTION_LABEL = { in: '入室', out: '退室' };
const PROP_KEY_SHEET_ID = 'SPREADSHEET_ID';
const PROP_KEY_TOKEN_PREFIX = 'TOKEN_';
const COLUMN_WIDTHS = { 1: 180, 2: 80, 3: 240 };

/** Entrypoints */
function doGet(e){ return handle_(e); }

/** Main */
function handle_(e){
  const p = parse_(e);
  
  // 完了ページ表示（PRGパターン）
  if (p.done && p.token){
    const result = getTokenResult_(p.token);
    if (result){
      return respondDone_(result);
    }
    // トークンが見つからない場合はエラー
    return respond_({ ok:false, code:'E20_TOKEN_NOT_FOUND', error:'無効なリクエストです。' });
  }
  
  // トークンがない場合は生成
  if (!p.token){
    p.token = generateToken_();
  }
  
  // トークンチェック：既に処理済みか確認
  const existingResult = getTokenResult_(p.token);
  if (existingResult){
    // 処理済みの場合はリダイレクト（PRGパターン）
    return redirectToDone_(p.token);
  }
  
  // バリデーション
  let actionKey = normalizeActionKey_(p.action);
  if (!actionKey){
    if (CONFIG.TEST_MODE){
      // テストモード：デフォルトで入室とする
      actionKey = 'in';
    } else {
      // デバッグ用：受け取ったパラメータをすべて表示
      let debugInfo = `受け取ったaction: "${p.action || '(空)'}", display_name: "${p.display_name || '(空)'}"`;
      if (p._debug) {
        debugInfo += `<br>パラメータキー: ${p._debug.paramKeys || '(なし)'}`;
        debugInfo += `<br>クエリ文字列: ${p._debug.queryString || '(なし)'}`;
        debugInfo += `<br>eオブジェクトのキー: ${p._debug.eKeys || '(なし)'}`;
        if (p._debug.queryString && p._debug.queryString !== '(なし)' && p._debug.queryString !== '(空文字列)') {
          debugInfo += `<br><small>クエリ文字列の内容: ${esc(p._debug.queryString)}</small>`;
        }
        if (p._debug.manualParams && p._debug.manualParams !== '(なし)') {
          debugInfo += `<br>手動パースしたパラメータ: ${esc(p._debug.manualParams)}`;
        }
        if (p._debug.finalParams && p._debug.finalParams !== '(空)') {
          debugInfo += `<br>最終パラメータ: ${esc(p._debug.finalParams)}`;
        }
      }
      return fail_('E30_INVALID_ACTION', `actionパラメータが無効です。${debugInfo}`);
    }
  }
  
  let displayName = p.display_name;
  if (!displayName){
    if (CONFIG.TEST_MODE){
      // テストモード：デフォルト値を設定
      displayName = 'テストユーザー';
    } else {
      const receivedName = p.display_name ? `受け取った値: "${p.display_name}"` : 'パラメータがありません';
      return fail_('E40_NO_NAME', `display_nameパラメータが必要です。${receivedName}`);
    }
  }
  
  const lock = LockService.getScriptLock();
  lock.waitLock(30 * 1000);
  try{
    const ss = openSpreadsheet_();
    const sh = getSheet_(ss, CONFIG.SHEET_LOGS);
    ensureHeadersAndLayout_(sh);

    const ts = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, CONFIG.DATETIME_FORMAT);
    const name = sanitizeName_(displayName);

    // 記録
    const row = [ts, actionLabel_(actionKey) || '', name];
    appendLogRow_(sh, row);

    // 処理結果を保存
    const result = { ok:true, action:actionKey, display_name:name, ts };
    saveTokenResult_(p.token, result);
    
    // PRGパターン：リダイレクト
    return redirectToDone_(p.token);
  } catch(err){
    const msg = String((err && (err.message || err)) || 'unknown');
    return fail_('E90_EXCEPTION', msg, p);
  } finally {
    lock.releaseLock();
  }
}

/** Parse: 値は加工せず保存 */
function parse_(e){
  // e.parameter と e.parameters の両方を確認
  const q = (e && e.parameter) ? e.parameter : (e && e.parameters) ? e.parameters : {};
  
  // デバッグ用：eオブジェクト全体を確認
  const paramKeys = Object.keys(q);
  const queryString = e && e.queryString !== undefined ? (e.queryString || '(空文字列)') : '(なし)';
  const eKeys = e ? Object.keys(e).join(', ') : '(eがnull)';
  
  // クエリ文字列から手動でパース（e.parameterが空の場合）
  let manualParams = {};
  if (e && e.queryString && e.queryString !== '' && Object.keys(q).length === 0) {
    try {
      const pairs = e.queryString.split('&');
      for (const pair of pairs) {
        const [key, value] = pair.split('=');
        if (key) {
          manualParams[decodeURIComponent(key)] = value ? decodeURIComponent(value) : '';
        }
      }
    } catch(err) {
      // パース失敗は無視
    }
  }
  
  // 手動パースしたパラメータを優先
  const finalParams = Object.keys(manualParams).length > 0 ? manualParams : q;
  
  return {
    action: String(finalParams.action || '').trim(),
    display_name: String(finalParams.display_name || '').trim(),
    token: String(finalParams.token || '').trim(),
    done: String(finalParams.done || '') === '1' || String(finalParams.done || '').toLowerCase() === 'true',
    _debug: {
      paramKeys: paramKeys.join(', ') || '(なし)',
      queryString: queryString,
      allParams: JSON.stringify(q) || '(空)',
      manualParams: JSON.stringify(manualParams) || '(なし)',
      finalParams: JSON.stringify(finalParams) || '(空)',
      eKeys: eKeys
    }
  };
}

/** エラー応答 */
function fail_(code, message, p){
  return respond_({ ok:false, code, error:message });
}

/** 応答(HTML) */
function respond_(payload){
  const ok = !!payload.ok;

  // HTML表示は日本語ラベルで
  const actDisp = actionLabel_(normalizeActionKey_(payload.action)) || '';

  // actionに応じてタイトルとメッセージを設定
  let title = '記録に失敗しました';
  let subHtml = '';
  let note = '';

  if (ok){
    if (actDisp === '入室'){
      title = '入室を受け付けました';
      subHtml = '入室を受け付けました';
    } else if (actDisp === '退室'){
      title = '退室が完了しました';
      subHtml = '退室が完了しました';
    } else {
      title = '記録しました';
      subHtml = `${esc(actDisp)}を記録しました`;
    }
    note = 'この画面を閉じてLINEに戻って大丈夫です。';
  } else {
    // エラーの詳細を表示
    const errorCode = esc(payload.code || '');
    const errorMsg = esc(payload.error || 'エラーが発生しました。');
    subHtml = `エラーコード: <strong>${errorCode}</strong><br>詳細: ${errorMsg}`;
    note = 'この画面を閉じて戻り、時間をおいて再度お試しください。';
  }

  const html = `<!doctype html>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  body{font-family:system-ui,-apple-system,BlinkMacSystemFont,Segoe UI,Roboto;margin:0;padding:24px;background:#f7f7f7;}
  .card{max-width:560px;margin:40px auto;background:#fff;border-radius:12px;box-shadow:0 6px 24px rgba(0,0,0,.08);padding:24px;text-align:center;}
  .icon{font-size:48px;margin-bottom:8px;${ok ? 'color:#2e7d32' : 'color:#c62828'};}
  h1{font-size:20px;margin:8px 0}
  p{color:#333;line-height:1.6;margin:8px 0 12px}
  small{color:#666}
</style>
<div class="card">
  <div class="icon">${ok ? '✓' : '✕'}</div>
  <h1>${esc(title)}</h1>
  <p>${subHtml}</p>
  <small>${esc(note)}</small>
</div>`;
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Layout */
function ensureHeadersAndLayout_(sheet){
  const range = sheet.getRange(1, 1, 1, HEADERS.length);
  const cur = (range.getValues()[0] || []).map(v => String(v || '').trim());
  const same = cur.length === HEADERS.length && cur.every((v,i)=>v===HEADERS[i]);
  if (!same){
    range.setValues([HEADERS]).setFontWeight('bold');
  }
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, COLUMN_WIDTHS[1]);
  sheet.setColumnWidth(2, COLUMN_WIDTHS[2]);
  sheet.setColumnWidth(3, COLUMN_WIDTHS[3]);
  const last = Math.max(sheet.getLastRow(), 2);
  sheet.getRange(2, 1, last - 1, 1).setNumberFormat(CONFIG.DATETIME_FORMAT);
  sheet.getRange(1, 1, sheet.getMaxRows(), HEADERS.length).setWrap(false);
}

/** Append without blank rows */
function appendLogRow_(sheet, rowValues){
  const r = Math.max(sheet.getLastRow() + 1, 2);
  sheet.getRange(r, 1, 1, rowValues.length).setValues([rowValues]);
  sheet.setRowHeight(r, 24);
  sheet.getRange(r, 1, 1, 3).setWrap(false);
  sheet.getRange(r, 1).setNumberFormat(CONFIG.DATETIME_FORMAT);
}

/** Spreadsheet helpers */
function openSpreadsheet_(){
  let id = CONFIG.SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty(PROP_KEY_SHEET_ID) || '';
  
  if (!id){
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active){
      id = active.getId();
      try { PropertiesService.getScriptProperties().setProperty(PROP_KEY_SHEET_ID, id); } catch(e){}
      return active;
    }
    throw new Error('E10_NO_SPREADSHEET: IDが未設定です。CONFIG.SPREADSHEET_IDにIDを設定してください。');
  }
  
  const ss = SpreadsheetApp.openById(id);
  // 初回のみ保存（既に同じIDなら保存不要）
  const savedId = PropertiesService.getScriptProperties().getProperty(PROP_KEY_SHEET_ID);
  if (savedId !== id){
    try { PropertiesService.getScriptProperties().setProperty(PROP_KEY_SHEET_ID, id); } catch(e){}
  }
  return ss;
}
function getSheet_(ss, name){
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** Utils */
function sanitizeName_(s){
  return String(s || '').replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim();
}
function esc(s){
  return String(s || '').replace(/[&<>"']/g, m => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[m]));
}

// 追加したactionを正規化→ラベル化の処理部分
function normalizeActionKey_(s){
  const v = String(s || '').trim().toLowerCase();
  if (v === 'in' || v === 'out') return v;
  if (v === '入' || v === '入室') return 'in';
  if (v === '退' || v === '退室') return 'out';
  return '';
}
function actionLabel_(key){
  return ACTION_LABEL[key] || '';
}

/** トークン管理 */
function generateToken_(){
  return Utilities.getUuid();
}

function saveTokenResult_(token, result){
  if (!token) return;
  const key = PROP_KEY_TOKEN_PREFIX + token;
  const data = { result: result };
  try {
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(data));
  } catch(e){
    // 保存失敗は無視（ログ記録は成功しているため）
  }
}

function getTokenResult_(token){
  if (!token) return null;
  const key = PROP_KEY_TOKEN_PREFIX + token;
  try {
    const dataStr = PropertiesService.getScriptProperties().getProperty(key);
    if (!dataStr) return null;
    
    const data = JSON.parse(dataStr);
    return data.result || null;
  } catch(e){
    return null;
  }
}

/** PRGパターン：リダイレクト */
function redirectToDone_(token){
  let scriptUrl;
  try {
    scriptUrl = ScriptApp.getService().getUrl();
  } catch(e){
    // デプロイされていない場合などは、現在のURLから取得を試みる
    scriptUrl = ScriptApp.getService() ? ScriptApp.getService().getUrl() : '';
  }
  
  if (!scriptUrl){
    // URL取得に失敗した場合は、トークンなしで完了ページを直接表示
    const result = getTokenResult_(token);
    if (result){
      return respondDone_(result);
    }
    return respond_({ ok:false, code:'E50_REDIRECT_FAILED', error:'リダイレクトに失敗しました。' });
  }
  
  // URLから既存のクエリパラメータを削除して、doneとtokenだけを付与
  const baseUrl = scriptUrl.split('?')[0];
  const redirectUrl = baseUrl + '?done=1&token=' + encodeURIComponent(token);
  
  const html = `<!doctype html>
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="refresh" content="0;url=${esc(redirectUrl)}">
<script>window.location.replace(${JSON.stringify(redirectUrl)});</script>
<body>リダイレクト中...</body>`;
  
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 完了ページ表示（respond_を呼び出してUIを統一） */
function respondDone_(result){
  return respond_(result);
}

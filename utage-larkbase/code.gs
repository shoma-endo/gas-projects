/* ───────── 1. 固定設定 ───────── */
const ENV_SHEET   = '環境変数シート';     // App ID / App Secret / BaseURL を置く
const DEBUG       = true;
const BATCH_LIMIT = 500;
const REG_NUM     = /(回数|金額|入金額)$/;

/** Google Sheets ↔ Bitable 対応表 */
const SHEETS = {
  common:      { name:'共通項目',        tableIdKey:'tableID_common'      },
  pay:         { name:'決済連携項目',    tableIdKey:'tableID_pay'         },
  event:       { name:'イベント連携項目', tableIdKey:'tableID_event'       },
  reservation: { name:'イベント予約情報', tableIdKey:'tableID_reservation' },
  scenario:    { name:'シナリオ固有項目', tableIdKey:'tableID_scenario'    },
  line:        { name:'LINE友だち項目',   tableIdKey:'tableID_line'        }
};

/* ───────── 2. グローバル ───────── */
const PROP       = PropertiesService.getScriptProperties();
const fieldCache = {};
function dlog(...a){ if (DEBUG) Logger.log(a.join(' ')); }

/* ───────── 3. ヘルパー ───────── */
function trimKey(k){ return (k||'').replace(/^[\u200B-\u200D\uFEFF\s]+|[\u200B-\u200D\uFEFF\s]+$/g,''); }
function getProp(key){
  const v = PROP.getProperty(key);
  if (v != null) return v;

  /* 誤綴り 'tabelID_*' → 修正コピー */
  const alt = key.replace(/^tableID_/,'tabelID_');
  const v2  = PROP.getProperty(alt);
  if (v2 != null){ PROP.setProperty(key,v2); PROP.deleteProperty(alt); return v2; }

  /* 不可視文字付きキーを救済 */
  const all = PROP.getProperties();
  for (const k in all){
    if (trimKey(k) === key){ PROP.setProperty(key,all[k]); PROP.deleteProperty(k); return all[k]; }
  }
  return null;
}
function setProp(k,v){ PROP.setProperty(trimKey(k),v); }

function toJsDate(v){
  if (v instanceof Date) return v;
  if (typeof v !== 'string') return null;
  const s = v.replace(/（.*?）|\(.*?\)/g,'').trim().replace(/\//g,'-');
  if (/^\d{4}-\d{2}-\d{2}$/.test(s))    return new Date(`${s}T00:00:00+09:00`);
  if (/^\d{4}-\d{2}-\d{2}[ T]/.test(s)) return new Date(s);
  return null;
}
const asNumber = v => { const n = Number(String(v).replace(/,/g,'')); return isFinite(n) ? n : null; };

/* ───────── 4. 認証 & 共通設定 ───────── */
function getCfg(){
  const env = SpreadsheetApp.getActive().getSheetByName(ENV_SHEET);
  if (!env) throw new Error(`ENV シート『${ENV_SHEET}』がありません`);
  const kv  = Object.fromEntries(env.getDataRange().getValues().filter(r=>r[0]&&r[1]));

  /* token キャッシュ */
  let tok = getProp('bitTok'), ttl = Number(getProp('bitTtl')||0);
  if (!tok || Date.now() > ttl){
    const js = JSON.parse(UrlFetchApp.fetch(
      'https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal',
      { method:'post', contentType:'application/json',
        payload:JSON.stringify({ app_id:kv['App ID'], app_secret:kv['App Secret'] }) }
    ).getContentText());
    if (js.code !== 0) throw new Error(js.msg);
    tok = js.tenant_access_token; ttl = Date.now() + (js.expire - 60)*1000;
    setProp('bitTok',tok); setProp('bitTtl',ttl);
  }

  /* BaseID / apiHost */
  let baseId = getProp('BaseID'), apiHost = getProp('apiHost');
  if (!baseId || !apiHost){
    const url = String(kv.BaseURL||'');
    const m   = url.match(/^(https?:\/\/[^/]+)\/(?:.*\/)?base\/([A-Za-z0-9]+)/);
    if (!m) throw new Error('BaseURL が不正です');
    baseId = m[2];

    if (/larksuite\.com/.test(url))
      apiHost = 'https://open.larksuite.com/open-apis/bitable/v1/apps';
    else if (/feishu\.cn/.test(url))
      apiHost = 'https://open.feishu.cn/open-apis/bitable/v1/apps';
    else  throw new Error('BaseURL ドメインが未対応です');

    setProp('BaseID',baseId); setProp('apiHost',apiHost);
  }
  return { token:tok, base:baseId, apiHost };
}

/* ───────── 5. Bitable API ───────── */
function callBase(cfg,method,path,body=null){
  const js = JSON.parse(UrlFetchApp.fetch(
    cfg.apiHost + path,
    { method, contentType:'application/json',
      headers:{ Authorization:`Bearer ${cfg.token}` },
      payload: body ? JSON.stringify(body) : undefined }
  ).getContentText());
  if (js.code !== 0) throw new Error(`API ${js.code}: ${js.msg}`);
  return js.data;
}
function getFieldMeta(cfg,tid){
  if (!fieldCache[tid])
    fieldCache[tid]=callBase(cfg,'get',`/${cfg.base}/tables/${tid}/fields`).items||[];
  return fieldCache[tid];
}

/* ───────── 6. 同期ロジック ───────── */
function extractRows(sh){
  const r=sh.getLastRow(), c=sh.getLastColumn();
  if (r<2) return{header:[],rows:[]};
  const v=sh.getRange(1,1,r,c).getValues();
  return{header:v[0].map(String),rows:v.slice(1)};
}
function convertRows(header,rows,exist,sentCol,ui){
  const rec=[], marks=[];
  rows.forEach((row,i)=>{
    if (sentCol>=0 && String(row[sentCol]).trim()==='済') return;
    const f={};
    header.forEach((h,idx)=>{
      if (!exist.has(h)) return;
      const v=row[idx]; if (v===''||v==null) return;
      const t=ui[h]||'';
      if (t.startsWith('Date')){ const d=toJsDate(v); if (d) f[h]=d.getTime(); }
      else if (REG_NUM.test(h)){ const n=asNumber(v); if (n!=null) f[h]=n; }
      else f[h]=String(v);
    });
    if (Object.keys(f).length){ rec.push({fields:f}); if (sentCol>=0) marks.push(i+2); }
  });
  return{records:rec, markRows:marks};
}
function uploadBatch(cfg,tid,rec,marks,sh,sentCol){
  while(rec.length){
    callBase(cfg,'post',`/${cfg.base}/tables/${tid}/records/batch_create?field_key=field_name`,
             {records:rec.splice(0,BATCH_LIMIT)});
    if (sentCol>=0 && marks.length)
      sh.getRangeList(marks.splice(0,BATCH_LIMIT).map(r=>`R${r}C${sentCol+1}`)).setValue('済');
  }
}
function syncSheet(def){
  const cfg=getCfg();
  const tid=getProp(def.tableIdKey);
  if (!tid) throw new Error(`${def.tableIdKey} が Script Properties にありません`);
  const sh=SpreadsheetApp.getActive().getSheetByName(def.name); if (!sh) return;
  const meta=getFieldMeta(cfg,tid);
  const exist=new Set(meta.map(f=>f.field_name));
  const ui=Object.fromEntries(meta.map(f=>[f.field_name,f.ui_type]));
  const {header,rows}=extractRows(sh); if (!header.length) return;
  const sent=header.indexOf('送信済');
  const {records,markRows}=convertRows(header,rows,exist,sent,ui);
  if (records.length) uploadBatch(cfg,tid,records,markRows,sh,sent);
}

/* ───────── 7. テーブル ID 自動補完 ───────── */
function autofillTableIDs(){
  const cfg=getCfg();
  const tbl=callBase(cfg,'get',`/${cfg.base}/tables`).items||[];
  const dict=Object.fromEntries(tbl.map(t=>[t.name,t.table_id]));
  let upd=0;
  Object.values(SHEETS).forEach(def=>{
    const id=dict[def.name]; if(!id){ dlog(`⚠ 未検出: ${def.name}`); return; }
    setProp(def.tableIdKey,id);
    PROP.deleteProperty('tabel'+def.tableIdKey.slice(5)); 
    upd++;
  });
  SpreadsheetApp.getUi().alert(`テーブル ID を ${upd} 件保存しました`);
}

/* ───────── 8. LINE シート即時同期 ───────── */
function onChange(e){
  const sh=e && e.source && e.source.getActiveSheet();
  if (!sh || sh.getName()!=='LINE友だち項目') return;

  /* ヘッダー取得 */
  const header=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const dateCol = header.indexOf('受信日時');
  const regCol  = header.indexOf('登録日時');

  /* ① 受信日時（日付のみ）を空欄行へ補完 */
  if (dateCol >= 0){
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    sh.getRange(2,dateCol+1, sh.getLastRow()-1,1).getValues()
      .forEach((v,i)=>{ if (!v[0]) sh.getRange(i+2, dateCol+1).setValue(today); });
  }

  /* ② 登録日時（日時）を空欄行へ補完 */
  if (regCol >= 0){
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sh.getRange(2,regCol+1, sh.getLastRow()-1,1).getValues()
      .forEach((v,i)=>{ if (!v[0]) sh.getRange(i+2, regCol+1).setValue(now); });
  }

  /* ③ 即時送信 */
  //syncSheet(SHEETS.line);
}

/* ───────── 9. トリガー（追加・一覧・個別削除） & メニュー ───────── */

/** 今すぐ全同期（LINE 含む） */
function runAllWithLine(){
  Object.values(SHEETS).forEach(def => syncSheet(def));
}

/** 今すぐLINEのみ同期（手動） */
function runLineNow(){
  syncSheet(SHEETS.line);
}

/** UIで 0..23 の「時」を取得（分は 00 固定） */
function promptHour_0to23_(title){
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(title, '0〜23 の整数を入力してください（分は 00 固定）', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return null;
  const h = Number(res.getResponseText());
  if (!Number.isInteger(h) || h < 0 || h > 23){
    ui.alert('0〜23 の整数を入力してください。');
    return null;
  }
  return h;
}

/* --- トリガー管理用メタデータ（Properties に保存） --- */
const TR_META_PREFIX = 'cron.runAllWithLine.'; // key例：cron.runAllWithLine.<uniqueId>

function saveTriggerMeta_(trigger, hour){
  const id = trigger.getUniqueId && trigger.getUniqueId();
  if (!id) return;
  const meta = { handler:'runAllWithLine', hour, createdAt:new Date().toISOString() };
  PROP.setProperty(TR_META_PREFIX + id, JSON.stringify(meta));
}
function removeTriggerMeta_(id){
  PROP.deleteProperty(TR_META_PREFIX + id);
}
function loadTriggerMetaMap_(){
  const all = PROP.getProperties();
  const map = {};
  const pfx = TR_META_PREFIX;
  for (const k in all){
    if (k.startsWith(pfx)){
      const id = k.slice(pfx.length);
      try{ map[id] = JSON.parse(all[k]); }
      catch(e){ map[id] = { handler:'runAllWithLine', parseError:true }; }
    }
  }
  return map;
}

/** 毎日トリガー作成—既存は消さずに増やす */
function createDailyTriggerAtHour(){
  const h = promptHour_0to23_('毎日トリガー追加（全シート）');
  if (h == null) return;
  const trig = ScriptApp.newTrigger('runAllWithLine').timeBased().atHour(h).nearMinute(0).everyDays(1).create();
  saveTriggerMeta_(trig, h);
  SpreadsheetApp.getUi().alert(`毎日 ${h}:00 のトリガーを追加しました。`);
}

/** （便利）カンマ区切りで複数時刻を一括追加：例「9,13,18」 */
function createDailyTriggersByCsv(){
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('複数時刻をカンマ区切りで入力', '例）9,13,18', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const raw = res.getResponseText() || '';
  const hours = [...new Set(raw.split(',').map(s=>s.trim()).filter(s=>s!=='').map(Number))]
    .filter(n=>Number.isInteger(n) && n>=0 && n<=23);
  if (!hours.length){ ui.alert('0〜23の整数をカンマ区切りで入力してください。'); return; }
  const made = [];
  hours.forEach(h=>{
    const trig = ScriptApp.newTrigger('runAllWithLine').timeBased().atHour(h).nearMinute(0).everyDays(1).create();
    saveTriggerMeta_(trig, h);
    made.push(h);
  });
  ui.alert(`毎日 ${made.map(h=>`${h}:00`).join(', ')} のトリガーを追加しました。`);
}

/** トリガー一覧文字列（runAllWithLine の CLOCK だけ） */
function buildTriggerListText_(){
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getEventType() === ScriptApp.EventType.CLOCK && t.getHandlerFunction() === 'runAllWithLine');
  const metaMap = loadTriggerMetaMap_();
  if (!triggers.length) return '（該当トリガーなし）';
  const lines = triggers.map((t, i)=>{
    const id = t.getUniqueId && t.getUniqueId();
    const meta = id ? metaMap[id] : null;
    const hh = (meta && Number.isInteger(meta.hour)) ? String(meta.hour).padStart(2,'0') : '??';
    const created = meta && meta.createdAt ? meta.createdAt : '';
    return `${i+1}. ${hh}:00  id=${id||'(unknown)'}  ${created}`;
  });
  return lines.join('\n');
}

/** 個別削除（一覧→番号入力で1件削除） */
function deleteOneDailyTrigger(){
  const ui = SpreadsheetApp.getUi();
  const list = buildTriggerListText_();
  const res = ui.prompt('削除するトリガー番号を入力', list + '\n\n番号を入力してください：', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const idx = Number(res.getResponseText());
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getEventType() === ScriptApp.EventType.CLOCK && t.getHandlerFunction() === 'runAllWithLine');
  if (!Number.isInteger(idx) || idx < 1 || idx > triggers.length){
    ui.alert('番号が不正です。'); return;
  }
  const target = triggers[idx-1];
  const id = target.getUniqueId && target.getUniqueId();
  ScriptApp.deleteTrigger(target);
  if (id) removeTriggerMeta_(id);
  ui.alert(`トリガー #${idx}（id=${id||'unknown'}）を削除しました。`);
}

/** 全削除（必要に応じて残しておく） */
function deleteAllDailyTriggers(){
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert('確認', 'runAllWithLine の毎日トリガーを全て削除します。よろしいですか？', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getEventType() === ScriptApp.EventType.CLOCK && t.getHandlerFunction() === 'runAllWithLine');
  const metaMap = loadTriggerMetaMap_();
  triggers.forEach(t=>{
    const id = t.getUniqueId && t.getUniqueId();
    ScriptApp.deleteTrigger(t);
    if (id && metaMap[id]) removeTriggerMeta_(id);
  });
  ui.alert('全ての runAllWithLine 用トリガーを削除しました。');
}

/** メニュー */
function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('UTAGE 設定')
    .addItem('テーブルID 自動取得','autofillTableIDs')
    .addSeparator()
    .addItem('今すぐ全同期','runAllWithLine')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('トリガー設定')
        .addItem('毎日トリガー追加（時刻指定・全シート）','createDailyTriggerAtHour')
        .addItem('毎日トリガー追加（複数時刻CSV）','createDailyTriggersByCsv')
        .addItem('トリガー一覧・個別削除','deleteOneDailyTrigger')
        .addItem('トリガー全削除','deleteAllDailyTriggers')
    )
    .addToUi();
}

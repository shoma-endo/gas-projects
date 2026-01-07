// ────────────────────────────────────────────────────────────────
// Lark Base Sync  (変更：2025-09-15)
// 変更点：投稿／リールの batch_create を「シートのレコードIDが無い場合のみ」に限定
// ────────────────────────────────────────────────────────────────

const ENV_SHEET = 'トークンとID';
const DEBUG     = true;

const SHEETS = {
  account : { name:'アカウント情報', tableIdKey:'tableID_account', keyField:'日付'   },
  follower: { name:'フォロワー情報', tableIdKey:'tableID_follower'                   },
  post    : { name:'投稿',           tableIdKey:'tableID_post', genre:'フィード', keyField:'投稿日' },
  reel    : { name:'リール',         tableIdKey:'tableID_post', genre:'リール',   keyField:'投稿日' },
};

const BASE       = 'https://open.larksuite.com/open-apis/bitable/v1/apps/';
const fieldCache = {};

// ⏱ small timer utility
function tic(lbl){const s=Date.now();return ()=>Logger.log(`⏱ ${lbl}: ${(Date.now()-s)/1000}s`);} 

// ───────── utility ─────────
function toJsDate(v){
  if(v instanceof Date) return v;
  if(typeof v!=='string') return null;
  const s=v.replace(/（.*?）|\(.*?\)/g,'').trim().replace(/\//g,'-');
  if(/^\d{4}-\d{2}-\d{2}$/.test(s))    return new Date(`${s}T00:00:00+09:00`);
  if(/^\d{4}-\d{2}-\d{2}[ T]/.test(s)) return new Date(s);
  return null;
}

function getCfg(){
  const kv  = SpreadsheetApp.getActive()
               .getSheetByName(ENV_SHEET)
               .getDataRange().getValues()
               .filter(r=>r[0]&&r[1]);
  const cfg = Object.fromEntries(kv);

  const prop=PropertiesService.getScriptProperties();
  let tok   = prop.getProperty('bitTok'),
      ttl   = Number(prop.getProperty('bitTtl')||0);

  if(!tok||Date.now()>ttl){
    const res=JSON.parse(UrlFetchApp.fetch(
      'https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal',{
        method:'post',
        contentType:'application/json',
        payload:JSON.stringify({app_id:cfg['App ID'],app_secret:cfg['App Secret']})
      }).getContentText());
    if(res.code!==0) throw new Error(`トークン取得失敗: ${res.msg}`);
    tok=res.tenant_access_token;
    prop.setProperty('bitTok',tok);
    prop.setProperty('bitTtl',String(Date.now()+(res.expire-60)*1000));
  }
  return Object.assign(cfg,{token:tok,base:cfg.BaseID});
}

function callBase(cfg,method,path,body=null){
  const opt={
    method:method.toUpperCase(),
    contentType:'application/json',
    headers:{Authorization:`Bearer ${cfg.token}`}
  };
  if(body) opt.payload=JSON.stringify(body);
  const raw=UrlFetchApp.fetch(BASE+path,opt).getContentText();
  if(DEBUG){
    const preview = raw.length > 500 ? raw.substring(0, 500) + '... (truncated)' : raw;
    Logger.log(`▶${opt.method} ${path}\n payload=${body?JSON.stringify(body):'null'}\n resp=${preview}`);
  }
  const js=JSON.parse(raw);
  if(js.code!==0) throw new Error(`API ${js.code}: ${js.msg}`);
  return js.data;
}

function getFieldMeta(cfg,tid){
  if(!fieldCache[tid]){
    const res=callBase(cfg,'get',`${cfg.base}/tables/${tid}/fields`);
    fieldCache[tid]=res.items||[];
  }
  return fieldCache[tid];
}

// ───────── URL×現状 で record_id 補完 ─────────
function syncRecordIds(def){
  const toc=tic(`${def.name} ▶ ID補完`);
  const cfg=getCfg(), tid=cfg[def.tableIdKey];
  const sh =SpreadsheetApp.getActive().getSheetByName(def.name);
  if(!sh){toc();return;}

  const headerIdx=(def.name==='投稿'||def.name==='リール')?2:1;
  const head=sh.getRange(headerIdx,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const ridIdx=head.indexOf('レコードID'), urlIdx=head.indexOf('投稿URL');
  if(ridIdx<0||urlIdx<0){toc();return;}

  const rows=sh.getRange(headerIdx+1,1,sh.getLastRow()-headerIdx,head.length).getValues();
  const targets=rows.map((r,i)=>({
      rowNum:i+headerIdx+1,
      url:String(r[urlIdx]).trim().replace(/\/+$/,''),
      rid:r[ridIdx]
    })).filter(o=>o.url&&!o.rid);
  if(!targets.length){toc();return;}

  let pageToken='', map={};
  do{
    const res=callBase(cfg,'get',
      `${cfg.base}/tables/${tid}/records?field_key=field_name${pageToken?`&page_token=${pageToken}`:''}`);
    const items=Array.isArray(res?.items)?res.items:[];
    items.forEach(rec=>{
      if(rec.fields['投稿後ステータス']!=='現状') return;
      const u=rec.fields['投稿URL'];
      if(u&&u.link){
        const clean=String(u.link).trim().replace(/\/+$/,'');
        map[clean]=rec.record_id;
      }
    });
    pageToken=res?.page_token||'';
  }while(pageToken);

  targets.forEach(o=>{
    if(map[o.url]) sh.getRange(o.rowNum,ridIdx+1).setValue(map[o.url]);
  });
  toc();
}

// ───────── 投稿／リール 送信（修正済み） ─────────
function _appendWithSentFlag(def){
  const toc=tic(`${def.name} ▶ batch処理`);

  const cfg=getCfg(), tid=cfg[def.tableIdKey];
  const sh =SpreadsheetApp.getActive().getSheetByName(def.name);
  if(!sh){toc();return;}

  const meta=getFieldMeta(cfg,tid);
  const exists=new Set(meta.map(f=>f.field_name));
  const uiType=Object.fromEntries(meta.map(f=>[f.field_name,f.ui_type]));
  const genreNames=meta.find(f=>f.field_name==='ジャンル')?.property?.options.map(o=>o.name)||[];
  const statusNames=meta.find(f=>f.field_name==='投稿後ステータス')?.property?.options.map(o=>o.name)||[];

  const headerIdx=(def.name==='投稿'||def.name==='リール')?2:1;
  if(sh.getLastRow()<headerIdx+1){toc();return;}

  const head=sh.getRange(headerIdx,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const statusRow=(def.name==='投稿'||def.name==='リール')
                  ?sh.getRange(headerIdx-1,1,1,head.length).getValues()[0]:[];

  const keyIdx=head.indexOf(def.keyField);
  const sentIdx=head.indexOf('送信済');
  const ridIdx=head.indexOf('レコードID');
  if(keyIdx<0) throw new Error(`ヘッダーに "${def.keyField}" が見つかりません`);
  if(sentIdx<0) throw new Error('"送信済" 列が必要です');
  if((def.name==='投稿'||def.name==='リール')&&ridIdx<0)
     throw new Error('"レコードID" 列が必要です');

  const nowStamp=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMdd');
  const rows=sh.getRange(headerIdx+1,1,sh.getLastRow()-headerIdx,head.length).getValues();

  const createBatch=[], updateBatch=[], newRows=[], updRows=[];

  rows.forEach((row,rIdx0)=>{
    const sentVal=String(row[sentIdx]).trim();
    const rowHasCurrent=head.some((c,ci)=>c==='いいね'&&String(statusRow[ci]||'').trim().replace(/数$/,'')==='現状');
    if(!rowHasCurrent && ['済','更新'].some(tag=>sentVal.startsWith(tag))) return;

    const dateJs=toJsDate(row[keyIdx]); if(!dateJs) return;
    const dateVal=uiType[def.keyField].startsWith('Date')
                 ?dateJs.getTime()
                 :Utilities.formatDate(dateJs,'Asia/Tokyo','yyyy-MM-dd');

    const url=head.includes('投稿URL')?row[head.indexOf('投稿URL')]:'';
    const cap=head.includes('キャプション')?row[head.indexOf('キャプション')]:'';

    if(def.name==='投稿'||def.name==='リール'){
      head.forEach((col,ci)=>{
        if(col!=='いいね') return;
        const st=String(statusRow[ci]||'').trim().replace(/数$/,'');
        if(!statusNames.includes(st)) return;

        // 送信するフィールド構築
        const fields={};
        fields[def.keyField]=dateVal;
        if(exists.has('ジャンル')&&genreNames.includes(def.genre)) fields['ジャンル']=def.genre;
        fields['投稿後ステータス']=st;
        if(exists.has('投稿URL')&&/^https?:\/\//.test(url)) fields['投稿URL']={link:String(url),text:''};
        if(exists.has('キャプション')&&cap) fields['キャプション']=String(cap);
        ['いいね','コメント','保存','エンゲージ','エンゲージ率','リーチ','IMP'].forEach((lab,k)=>{
          const v=row[ci+k];
          if(exists.has(lab)&&v!==''&&v!=null) {
            fields[lab] = uiType[lab] === 'Number' ? Number(v) : String(v);
          }
        });

        const rid=ridIdx>=0?String(row[ridIdx]||'').trim():'';
        const hasRid=!!rid; // ← 判定基準を RID 優先に変更

        // ◆ 新ロジック：
        //   - RID 無し        → 作成（batch_create）
        //   - RID 有り & 現状 → 更新（batch_update）
        //   - RID 有り & 非現状 → 何もしない（スキップ）
        if(!hasRid){
          createBatch.push({fields});
          newRows.push({rowNum:rIdx0+headerIdx+1,val:`済(${nowStamp})`, isCurrent:(st==='現状')});
          return;
        }

        if(st==='現状'){
          updateBatch.push({record_id:rid,fields});
          updRows.push({rowNum:rIdx0+headerIdx+1,val:`更新(${nowStamp})`});
          return;
        }

        // RID あり & 非現状 → スキップ
      });
    }else{
      // アカウント情報などは従来通り：常に作成
      // B~G列(インデックス1~6)が全て空ならスキップ
      if(def.name==='アカウント情報'){
        const allEmpty = row.slice(1, 7).every(v => v==='' || v==null);
        if(allEmpty) return;
      }

      const fields={};
      head.forEach((h,ci)=>{
        if(!h||!exists.has(h)) return;
        const v=row[ci]; if(v===''||v==null) return;
        if(uiType[h].startsWith('Date')){
          const d=toJsDate(v); if(d) fields[h]=d.getTime();
        }else{
          fields[h] = uiType[h] === 'Number' ? Number(v) : String(v);
        }
      });
      createBatch.push({fields});
      newRows.push({rowNum:rIdx0+headerIdx+1,val:`済(${nowStamp})`,isCurrent:false});
    }
  });

  // ─── 更新バッチ（500件ずつ） ───
  const BATCH_SIZE = 500;
  if(updateBatch.length){
    for(let i=0; i<updateBatch.length; i+=BATCH_SIZE){
      const chunk = updateBatch.slice(i, i+BATCH_SIZE);
      callBase(cfg,'post',`${cfg.base}/tables/${tid}/records/batch_update?field_key=field_name`,
               {records:chunk});
    }
    updRows.forEach(o=>sh.getRange(o.rowNum,sentIdx+1).setValue(o.val));
  }

  // ─── 作成バッチ（500件ずつ） ───
  if(createBatch.length){
    let allCreatedRecords = [];
    for(let i=0; i<createBatch.length; i+=BATCH_SIZE){
      const chunk = createBatch.slice(i, i+BATCH_SIZE);
      const res = callBase(cfg,'post',
        `${cfg.base}/tables/${tid}/records/batch_create?field_key=field_name`,
        {records:chunk});
      allCreatedRecords = allCreatedRecords.concat(res.records || []);
    }

    allCreatedRecords.forEach((rec,i)=>{
      const o=newRows[i];
      // ★ 現状レコードだけ RecordID を書く（従来挙動を維持）
      if(o.isCurrent && ridIdx>=0){
        const rid=rec.record_id||(rec.fields?.['レコードID']?.[0]?.text);
        if(rid) sh.getRange(o.rowNum,ridIdx+1).setValue(rid);
      }
      sh.getRange(o.rowNum,sentIdx+1).setValue(o.val);
    });
  }
  toc();
}

// ───────── フォロワー情報 ─────────
function _appendFollower(def){
  const toc=tic(`${def.name} ▶ アップロード`);
  const cfg=getCfg(), tid=cfg[def.tableIdKey];
  const sh=SpreadsheetApp.getActive().getSheetByName(def.name);
  if(!sh){toc();return;}

  const meta=getFieldMeta(cfg,tid);
  const exists=new Set(meta.map(f=>f.field_name));
  const uiType=Object.fromEntries(meta.map(f=>[f.field_name,f.ui_type]));

  const vals=sh.getDataRange().getValues();
  if(vals.length<3){toc();return;}

  const dateM=/\d{4}[\/\-]\d{2}[\/\-]\d{2}/.exec(String(vals[0][0]));
  if(!dateM) throw new Error('A1 に日付が見つかりません');
  let dateVal=dateM[0].replace(/\//g,'-');
  if(uiType['日付']?.startsWith('Date')){
    const d=toJsDate(dateVal); if(d) dateVal=d.getTime();
  }

  const headerIdx=vals.findIndex(r=>r.some(c=>String(c).trim()==='年齢')&&r.some(c=>String(c).trim()==='性別'));
  if(headerIdx<0) throw new Error('ヘッダー行が見つかりません');

  const header=vals[headerIdx].map(String);
  const ageIdx=header.indexOf('年齢'), ageCntIdx=header.indexOf('数',ageIdx),
        genIdx=header.indexOf('性別'), genCntIdx=header.indexOf('数',genIdx);

  const fields={'日付':dateVal};
  if(ageIdx>=0&&ageCntIdx>=0){
    for(let i=headerIdx+1;i<vals.length;i++){
      const grp=String(vals[i][ageIdx]).trim();
      const cnt=vals[i][ageCntIdx];
      const fName=`年齢 (${grp})`;
      if(grp&&cnt!=null&&exists.has(fName)) {
        fields[fName] = uiType[fName] === 'Number' ? Number(cnt) : String(cnt);
      }
    }
  }
  if(genIdx>=0&&genCntIdx>=0){
    for(let i=headerIdx+1;i<vals.length;i++){
      const g=String(vals[i][genIdx]).trim();
      const cnt=vals[i][genCntIdx];
      if(g&&cnt!=null&&exists.has(g)) {
        fields[g] = uiType[g] === 'Number' ? Number(cnt) : String(cnt);
      }
    }
  }

  callBase(cfg,'post',`${cfg.base}/tables/${tid}/records?field_key=field_name`,{fields});
  toc();
}

// ───────── 公開関数 ─────────
const syncAccount      = ()=>_appendWithSentFlag(SHEETS.account);
const syncFollowerInfo = ()=>_appendFollower(SHEETS.follower);
const syncPosts        = ()=>_appendWithSentFlag(SHEETS.post);
const syncReels        = ()=>_appendWithSentFlag(SHEETS.reel);

// ───────── メイン ─────────
function runAll(){
  const toc=tic('◎ runAll TOTAL');
  syncRecordIds(SHEETS.post);
  syncRecordIds(SHEETS.reel);

  syncAccount();
  syncFollowerInfo();
  syncPosts();
  syncReels();

  toc();
}

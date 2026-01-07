/******************************************************
 * Lark Base → GAS → Gemini OCR(JSON配列) → Baseに新規レコード追加
 * フィールド: 商品名 / 発注数量 / 発注者 / お届け先（会社名） / 注文番号 / 支店・営業所名 / 郵便番号 / 住所詳細 / 電話番号 / 都道府県
 * 
 * 【重要】Gemini APIモデル設定について
 * - 無料プラン（Free Tier）で使用可能なモデル:
 *   * gemini-2.5-flash (推奨: 最新版、高速で低コスト、OCRに最適)
 *   * gemini-2.0-flash (高速で低コスト)
 *   * gemini-1.5-flash (高速で低コスト)
 *   * gemini-1.5-pro (高精度が必要な場合)
 * - 無料プランで使用できないモデル:
 *   * gemini-2.5-pro (有料プランが必要)
 * 
 * GEMINI_MODELの設定場所: test.js で定義されています。
 * 無料プランを使用する場合は、test.js の GEMINI_MODEL を gemini-2.5-flash に変更してください。
 ******************************************************/

// 注意: 設定値は test.js で定義されています（GASのグローバルスコープで共有されます）

// ========== 商品名正規化機能 ==========

// 商品名マッピングテーブル
const PRODUCT_NAME_MAPPING = {
  'カルスNC-R_10kg': [
    'カルスNC-R 10',
    'カルスNC-R 10kg',
    'カルスNCR 10',
    'カルスNC-R10kg',
    'カルスNC-R_10kg'
  ],
  'アイデンカルス_10kg': [
    'アイデンカルス 10',
    'アイデンカルス 10kg',
    'アイテンカルス 10',
    'アイデンカルス10kg',
    'アイデンカルス_10kg'
  ],
  'アイデンマック_5kg': [
    'アイデンマック 5',
    'アイデンマック 5kg',
    'アイテンマック 5',
    'アイデンマック5kg',
    'アイデンマック_5kg'
  ],
  'カルスNC-R_1kg×20袋': [
    'カルスNC-R 1kg 20',
    'カルスNC-R 1kg×20袋',
    'カルスNC-R 1kg 20袋',
    'カルスNCR 1kg 20',
    'カルスNC-R_1kg×20袋',
    'カルスNC-R 1kg袋',   
  　'カルスNCR 1kg袋'        
  ],
  'リサールSE_5kg×4袋': [
    'リサールSE 5 4',
    'リサールSE 5kg×4袋',
    'リサールSE 5kg 4袋',
    'リサールSE 5 4袋',
    'リサールSE_5kg×4袋'
  ],
  '粒状カルスNC-R_10kg': [
    '粒状カルスNC-R 10',
    '粒状カルスNC-R 10kg',
    '粒状カルスNCR 10',
    '粒状カルスNC-R10kg',
    '粒状カルスNC-R_10kg',
    '粒状カルスNC-R/',        // ← 追加
    '粒状カルスNC-R\\'   
  ],
  'サルバーS_10kg': [
    'サルバーS',
    'サルバーS 10kg',
    'サルバーS 10',
    'サルバーS10kg',
    'サルバーS_10kg'
  ],
  '粒状サルバーS': [
    '粒状サルバーS',
    '粒状サルバー S',
    '粒状サルバーs',
    '粒状サルバ-S'
  ]
};

// 文字列を正規化
const normalizeString = (str) => {
  if (!str) return '';
  return str
    .replace(/[ー−]/g, '-')
    .replace(/[×xX]/g, '×')
    .replace(/[\s_]/g, '')
    .toLowerCase();
};

// 商品名を正規化
const normalizeProductName = (ocrName) => {
  if (!ocrName) return '';
  
  const normalizedOcr = normalizeString(ocrName);
  
  for (const [officialName, patterns] of Object.entries(PRODUCT_NAME_MAPPING)) {
    if (patterns.includes(ocrName)) {
      return officialName;
    }
    
    for (const pattern of patterns) {
      if (normalizeString(pattern) === normalizedOcr) {
        return officialName;
      }
    }
  }
  
  for (const [officialName, patterns] of Object.entries(PRODUCT_NAME_MAPPING)) {
    const corePattern = patterns[0];
    const normalizedCore = normalizeString(corePattern);
    
    if (normalizedOcr.includes(normalizedCore) || normalizedCore.includes(normalizedOcr)) {
      return officialName;
    }
  }
  
  Logger.log(`マッチング失敗: ${ocrName}`);
  return ocrName;
};

// 数量を数値に変換
const parseQuantity = (val) => {
  if (!val) return 0;
  // 数字と小数点以外を除去してから数値化
  const num = Number(String(val).replace(/[^0-9.]/g, ''));
  return isNaN(num) ? 0 : num;
};

// ========== ファイルタイプ判定 ==========

// サポートされているファイルタイプ
const SUPPORTED_MIME_TYPES = {
  'pdf': 'application/pdf',
  'jpg': 'image/jpeg',
  'jpeg': 'image/jpeg',
  'png': 'image/png',
  'bmp': 'image/bmp',
  'webp': 'image/webp'
};

// ファイル名からMIMEタイプを推定
const getMimeTypeFromFileName = (fileName) => {
  if (!fileName) return 'application/pdf';
  const ext = fileName.toLowerCase().split('.').pop();
  return SUPPORTED_MIME_TYPES[ext] || 'application/pdf';
};

// MIMEタイプが画像かどうかを判定
const isImageMimeType = (mimeType) => {
  return mimeType && mimeType.startsWith('image/');
};

// テスト用関数
function testProductNameMapping() {
  const testCases = [
    'カルスNC-R 10',
    'アイデンカルス 10',
    'アイデンマック 5',
    'カルスNC-R 1kg 20',
    'リサールSE 5 4',
    '粒状カルスNC-R 10',
    'サルバーS',
    '粒状サルバーS'
  ];
  
  testCases.forEach(testCase => {
    const result = normalizeProductName(testCase);
    Logger.log(`${testCase} → ${result}`);
  });
}
/** 共通: JSONレスポンス */
const jsonResponse = (obj) =>
  ContentService.createTextOutput(JSON.stringify(obj||{})).setMimeType(ContentService.MimeType.JSON);

/* ========== ログ出力 ========== */
function writeLog(message, data = null) {
  try {
    const sheet = SpreadsheetApp.openById(LOG_SPREADSHEET_ID).getSheetByName(LOG_SHEET_NAME);
    if (!sheet) {
      Logger.log(`[writeLog] シートが見つかりません: ${LOG_SHEET_NAME}`);
      return;
    }
    const row = [
      new Date(),
      message,
      data ? JSON.stringify(data, null, 2) : ''
    ];
    sheet.appendRow(row);
    Logger.log(`[writeLog] ${message}: ${data ? JSON.stringify(data) : ''}`);
  } catch(e) {
    Logger.log(`[writeLog Error] ${e}`);
  }
}

/* ========== Lark ========== */

const getTenantAccessToken = () => {
  const url = `${OPEN_API_HOST}/open-apis/auth/v3/tenant_access_token/internal`;
  const body = { app_id: LARK_APP_ID, app_secret: LARK_APP_SECRET };
  const res = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(body)
  });
  const json = JSON.parse(res.getContentText());
  if (!json.tenant_access_token) throw new Error('tenant_access_token not found');
  return json.tenant_access_token;
};

/** file_token → tmp_download_url */
const getTmpDownloadUrlByToken = (fileToken) => {
  const url = `${OPEN_API_HOST}/open-apis/drive/v1/medias/batch_get_tmp_download_url?file_tokens=${encodeURIComponent(fileToken)}`;
  const headers = { Authorization: 'Bearer ' + getTenantAccessToken() };
  const res = UrlFetchApp.fetch(url, { method: 'get', headers });
  const json = JSON.parse(res.getContentText());
  const arr = json?.data?.tmp_download_urls || [];
  const tmp = arr[0]?.tmp_download_url;
  if (!tmp) throw new Error('tmp_download_url not found for token=' + fileToken);
  return tmp;
};

/** Base: レコード新規作成 */
const createLarkRecord_ = (tenantTok, baseId, tableId, fieldsObj) => {
  const url = `${OPEN_API_HOST}/open-apis/bitable/v1/apps/${baseId}/tables/${tableId}/records/batch_create`;
  const payload = { records: [{ fields: fieldsObj }] };
  
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { Authorization: `Bearer ${tenantTok}`, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  const responseCode = res.getResponseCode();
  const responseText = res.getContentText();
  
  const json = JSON.parse(responseText);
  
  // レスポンスの code を確認（Lark APIはHTTP 200でも code !== 0 でエラーを返すことがある）
  if (json.code !== 0) {
    const errorMsg = `[createLarkRecord] ${json.code}: ${json.msg || responseText}`;
    writeLog('createLarkRecord エラー', { code: json.code, msg: json.msg, responseText, url, payload });
    throw new Error(errorMsg);
  }
  
  return json;
};

/* ========== Gemini ========== */

/** ファイルアップロード */
const geminiUploadFile_ = (bytes, mimeType, displayName) => {
  // GEMINI_API_KEYのチェック
  if (!GEMINI_API_KEY) {
    const errorMsg = 'GEMINI_API_KEYが設定されていません。スクリプトプロパティにGEMINI_API_KEYを設定してください。';
    writeLog('Gemini API Key エラー', { error: errorMsg });
    throw new Error(errorMsg);
  }
  
  const url = `${GEMINI_BASE}/upload/v1beta/files?key=${encodeURIComponent(GEMINI_API_KEY)}`;
  const headers = {
    'X-Goog-Upload-Protocol': 'raw',
    'X-Goog-Upload-Command': 'upload, finalize',
    'X-Goog-Upload-Header-Content-Length': String(bytes.length),
    'X-Goog-Upload-Header-Content-Type': mimeType,
    'X-Goog-Upload-File-Name': displayName || ('upload_' + Date.now())
  };
  
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers,
    contentType: mimeType,
    payload: bytes,
    muteHttpExceptions: true,
    fetchTimeout: 60 // タイムアウトを60秒に設定（大きなPDFファイルのアップロードに対応）
  });
  
  const responseCode = res.getResponseCode();
  const responseText = res.getContentText();
  
  if (responseCode >= 400) {
    writeLog('Gemini ファイルアップロード エラー', { 
      responseCode, 
      responseText,
      fileName: displayName,
      mimeType,
      size: bytes.length
    });
    throw new Error(`Gemini upload error (${responseCode}): ${responseText}`);
  }
  
  let json;
  try {
    json = JSON.parse(responseText);
  } catch(e) {
    writeLog('Gemini ファイルアップロード JSON解析エラー', { 
      responseText, 
      error: String(e),
      fileName: displayName
    });
    throw new Error(`Gemini upload response is not valid JSON: ${responseText.substring(0, 500)}`);
  }
  
  const file = json.file || json;
  if (!file || !file.uri) {
    writeLog('Gemini ファイルアップロード file.uriが存在しない', { json, fileName: displayName });
    throw new Error('Gemini upload did not return file.uri. Response: ' + JSON.stringify(json).substring(0, 500));
  }
  
  return file;
};

/** OCR(JSON配列で返す) - PDF/画像両対応 */
const geminiExtractFieldsFromFileUri_ = (fileUri, mimeType) => {
  // GEMINI_API_KEYのチェック
  if (!GEMINI_API_KEY) {
    const errorMsg = 'GEMINI_API_KEYが設定されていません。スクリプトプロパティにGEMINI_API_KEYを設定してください。';
    writeLog('Gemini API Key エラー', { error: errorMsg });
    throw new Error(errorMsg);
  }
  
  const url = `${GEMINI_BASE}/v1beta/models/${encodeURIComponent(GEMINI_MODEL)}:generateContent`;
  
  // ファイルタイプに応じた説明文
  const fileTypeDesc = isImageMimeType(mimeType) ? '画像' : 'PDF';
  const multiPageNote = isImageMimeType(mimeType) 
    ? '画像に複数の発注書が含まれている場合は、それぞれを別のレコードとして抽出してください。'
    : 'PDFに複数の発注書が含まれている場合は、ページごと（または発注書ごと）に分割して、それぞれを別のレコードとして抽出してください。';
  
  const prompt = `
次の${fileTypeDesc}をOCRして、以下のフィールドを必ず抽出してください。
${multiPageNote}
必ず有効な JSON 配列だけを返してください。余計な文章やコードブロックは書かないでください。

[
  {
    "item_name": "商品名",
    "quantity": "数量",
    "orderer": "発注者",
    "destination": "送付先(宛先)",
    "order_number": "注文番号",
    "branch_name": "支店・営業所名",
    "postal_code": "郵便番号",
    "address_detail": "住所詳細",
    "phone_number": "電話番号",
    "prefecture": "都道府県"
  }
]
`;
  const payload = {
    contents: [{
      role: 'user',
      parts: [
        { file_data: { file_uri: fileUri, mime_type: mimeType } },
        { text: prompt }
      ]
    }]
  };

  // PDF処理は時間がかかるため、タイムアウトエラーを適切に処理
  let res;
  let responseCode;
  let responseText;
  
  try {
    res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { 'x-goog-api-key': GEMINI_API_KEY, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      fetchTimeout: 60 // タイムアウトを60秒に設定（PDF処理に時間がかかる場合に対応）
    });
    
    responseCode = res.getResponseCode();
    responseText = res.getContentText();
  } catch(e) {
    const errorStr = String(e);
    // タイムアウトエラーの場合（GASのUrlFetchApp.fetchは60秒でタイムアウト）
    if (errorStr.includes('timeout') || errorStr.includes('Timeout') || 
        errorStr.includes('Execution') || errorStr.includes('time limit')) {
      const fileSizeNote = isImageMimeType(mimeType) 
        ? '画像ファイルが大きすぎるか、複雑すぎる可能性があります。'
        : 'PDFファイルが大きすぎるか、複雑すぎる可能性があります。';
      const errorMsg = `Gemini OCR API タイムアウト: ${fileTypeDesc}処理に時間がかかりすぎています（60秒制限）。` +
        `${fileSizeNote} ファイルタイプ: ${mimeType}, URI: ${fileUri}`;
      writeLog('Gemini OCR API タイムアウト', { 
        mimeType,
        fileTypeDesc,
        fileUri,
        error: errorStr
      });
      throw new Error(errorMsg);
    }
    // その他のエラー
    writeLog('Gemini OCR API リクエストエラー', { 
      mimeType,
      fileTypeDesc,
      fileUri,
      error: errorStr
    });
    throw e;
  }
  
  if (responseCode >= 300) {
    writeLog('Gemini OCR API エラー', { responseCode, responseText, mimeType, fileUri });
    
    // 429エラー（レート制限）の場合、詳細なエラーメッセージを生成
    if (responseCode === 429) {
      let errorMsg = 'Gemini APIのレート制限に達しました。';
      try {
        const errorJson = JSON.parse(responseText);
        if (errorJson.error) {
          const error = errorJson.error;
          errorMsg += `\nエラーコード: ${error.code || responseCode}`;
          errorMsg += `\nステータス: ${error.status || 'RESOURCE_EXHAUSTED'}`;
          errorMsg += `\nメッセージ: ${error.message || 'クォータ超過'}`;
          
          // リトライ情報を抽出
          if (error.details) {
            const retryInfo = error.details.find(d => d['@type'] === 'type.googleapis.com/google.rpc.RetryInfo');
            if (retryInfo && retryInfo.retryDelay) {
              const retrySeconds = parseInt(retryInfo.retryDelay.replace('s', '')) || 0;
              errorMsg += `\n\n推奨リトライ時間: ${retrySeconds}秒後`;
            }
            
            // クォータ違反の詳細を抽出
            const quotaFailure = error.details.find(d => d['@type'] === 'type.googleapis.com/google.rpc.QuotaFailure');
            if (quotaFailure && quotaFailure.violations) {
              errorMsg += '\n\n超過したクォータ:';
              quotaFailure.violations.forEach((v, i) => {
                errorMsg += `\n${i+1}. ${v.quotaMetric || '不明'}`;
                if (v.quotaDimensions && v.quotaDimensions.model) {
                  errorMsg += ` (モデル: ${v.quotaDimensions.model})`;
                }
              });
            }
          }
          
          // モデル名が無料プランで使用不可の場合の警告
          if (error.message && error.message.includes('limit: 0')) {
            errorMsg += '\n\n⚠️ 警告: 使用しているモデルが無料プランでは使用できません。';
            errorMsg += '\n無料プランで使用可能なモデル（推奨: gemini-2.5-flash, または gemini-2.0-flash, gemini-1.5-flash）に変更してください。';
          }
        }
      } catch(e) {
        // JSON解析に失敗した場合は元のメッセージを使用
      }
      throw new Error(errorMsg);
    }
    
    throw new Error(`generateContent failed (${responseCode}): ${responseText}`);
  }
  
  let json;
  try {
    json = JSON.parse(responseText);
  } catch(e) {
    writeLog('Gemini OCR API JSON解析エラー', { responseText, error: String(e) });
    throw new Error(`Gemini OCR API response is not valid JSON: ${responseText.substring(0, 500)}`);
  }
  
  // Gemini APIのエラーレスポンスをチェック
  if (json.error) {
    writeLog('Gemini OCR API エラーレスポンス', { error: json.error });
    
    // 429エラーの場合、詳細なメッセージを生成
    if (json.error.code === 429) {
      let errorMsg = 'Gemini APIのレート制限に達しました。';
      errorMsg += `\nエラーコード: ${json.error.code}`;
      errorMsg += `\nステータス: ${json.error.status || 'RESOURCE_EXHAUSTED'}`;
      errorMsg += `\nメッセージ: ${json.error.message || 'クォータ超過'}`;
      
      // リトライ情報を抽出
      if (json.error.details) {
        const retryInfo = json.error.details.find(d => d['@type'] === 'type.googleapis.com/google.rpc.RetryInfo');
        if (retryInfo && retryInfo.retryDelay) {
          const retrySeconds = parseInt(retryInfo.retryDelay.replace('s', '')) || 0;
          errorMsg += `\n\n推奨リトライ時間: ${retrySeconds}秒後`;
        }
        
        // クォータ違反の詳細を抽出
        const quotaFailure = json.error.details.find(d => d['@type'] === 'type.googleapis.com/google.rpc.QuotaFailure');
        if (quotaFailure && quotaFailure.violations) {
          errorMsg += '\n\n超過したクォータ:';
          quotaFailure.violations.forEach((v, i) => {
            errorMsg += `\n${i+1}. ${v.quotaMetric || '不明'}`;
            if (v.quotaDimensions && v.quotaDimensions.model) {
              errorMsg += ` (モデル: ${v.quotaDimensions.model})`;
            }
          });
        }
      }
      
      // モデル名が無料プランで使用不可の場合の警告
      if (json.error.message && json.error.message.includes('limit: 0')) {
        errorMsg += '\n\n⚠️ 警告: 使用しているモデル（gemini-2.5-pro）が無料プランでは使用できません。';
        errorMsg += '\n無料プランで使用可能なモデル（推奨: gemini-2.5-flash, または gemini-2.0-flash, gemini-1.5-flash）に変更してください。';
      }
      
      throw new Error(errorMsg);
    }
    
    throw new Error(`Gemini OCR API error: ${JSON.stringify(json.error)}`);
  }
  
  // candidatesが存在しない、または空の場合
  if (!json.candidates || json.candidates.length === 0) {
    writeLog('Gemini OCR API candidatesが空', { json });
    throw new Error('Gemini OCR API returned no candidates. Response: ' + JSON.stringify(json).substring(0, 500));
  }
  
  const parts = json.candidates[0]?.content?.parts || [];
  if (parts.length === 0) {
    writeLog('Gemini OCR API partsが空', { json });
    throw new Error('Gemini OCR API returned no parts. Response: ' + JSON.stringify(json).substring(0, 500));
  }
  
  const out = parts.map(p => p.text || '').join('').trim();
  
  if (!out || out.length === 0) {
    writeLog('Gemini OCR API テキストが空', { json, parts });
    throw new Error('Gemini OCR API returned empty text. Response: ' + JSON.stringify(json).substring(0, 500));
  }

  // コードブロック削除
  const cleaned = out.replace(/^```json/i, '').replace(/```$/i, '').trim();

  let arr;
  try {
    arr = JSON.parse(cleaned);
  } catch(e) {
    throw new Error('Gemini OCR result is not valid JSON: ' + out);
  }
  if (!Array.isArray(arr)) arr = [arr];
  return arr;
};

/* ========== Base URL管理 ========== */
// ↓↓↓ リンク先（代理店・支店マスタ）の情報をここに記入してください ↓↓↓
const ORDERER_TABLE_ID = "tbl8s8jEadxVXCLw"; // 代理店・支店マスタのテーブルID
const ORDERER_NAME_FIELD = "代理店・支店";    // 代理店・支店マスタ側のフィールド名
// ↑↑↑↑↑↑

/** Base: マスターテーブルから全レコードを取得（発注者のみ） */
const getAllMasterRecords_ = (tenantTok, baseId, tableId) => {
  const url = `${OPEN_API_HOST}/open-apis/bitable/v1/apps/${baseId}/tables/${tableId}/records/search`;
  const payload = {
    field_names: [ORDERER_NAME_FIELD],
    automatic_fields: false
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { Authorization: `Bearer ${tenantTok}`, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  const json = JSON.parse(res.getContentText());
  if (json.code !== 0) {
    Logger.log(`[getAllMasterRecords] Error: ${json.msg}`);
    return [];
  }
  
  const items = json.data?.items || [];
  return items.map(item => ({
    record_id: item.record_id,
    orderer: item.fields[ORDERER_NAME_FIELD] || ''
  }));
};

/** Gemini APIを使ってマスターレコードとマッチング */
const matchOrdererWithGemini_ = (ocrOrderer, ocrBranch, masterRecords) => {
  if (!ocrOrderer) return null;
  if (masterRecords.length === 0) return null;
  
  // GEMINI_API_KEYのチェック
  if (!GEMINI_API_KEY) {
    Logger.log('[matchOrdererWithGemini] GEMINI_API_KEYが設定されていません');
    return null;
  }
  
  // マスターレコードをJSON形式で準備（発注者のみ）
  const masterList = masterRecords.map((m, idx) => ({
    index: idx,
    record_id: m.record_id,
    orderer: m.orderer
  }));
  
  // OCR結果のキーを作成（発注者_支店・営業所名の組み合わせ）
  const ocrKey = ocrBranch ? `${ocrOrderer}_${ocrBranch}` : ocrOrderer;
  const hasBranch = ocrBranch && ocrBranch.trim() !== '';
  
  const prompt = `
以下のOCRで抽出された発注者情報と、マスターテーブルの発注者リストを比較して、最も一致するレコードのindexを選んでください。

【OCRで抽出された情報】
発注者: "${ocrOrderer}"
支店・営業所名: ${hasBranch ? `"${ocrBranch}"` : '(なし)'}
検索キー: "${ocrKey}"

【マスターテーブルの発注者リスト】
${JSON.stringify(masterList, null, 2)}

【重要】
- マスターテーブルには「発注者」のみが登録されています（支店・営業所名は含まれていません）
- OCRの支店・営業所名が「(なし)」の場合は、発注者のみでマッチングしてください
- OCRの支店・営業所名がある場合でも、マスターテーブルには支店・営業所名がないため、OCRの「発注者」部分とマスターテーブルの「発注者」を比較してください

【指示】
1. OCRで抽出された「発注者」とマスターテーブルの「発注者」を比較してください
2. 支店・営業所名は参考情報としてのみ使用し、マッチング判定には使用しません
3. 完全一致する場合はそのレコードのindexを返してください
4. 完全一致しない場合でも、表記の揺れ（スペースの有無、略称、別表記など）を考慮して、ほぼ同じと判断できる場合はそのレコードのindexを返してください
5. 一致するレコードがない場合は null を返してください

【出力形式】
必ず以下のJSON形式で返してください。余計な文章は書かないでください。
{
  "matched_index": 0 または null,
  "matched_record_id": "rec..." または null,
  "confidence": "high" または "medium" または "low" または "none"
}
`;

  const url = `${GEMINI_BASE}/v1beta/models/${encodeURIComponent(GEMINI_MODEL)}:generateContent`;
  const payload = {
    contents: [{
      role: 'user',
      parts: [{ text: prompt }]
    }]
  };

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { 'x-goog-api-key': GEMINI_API_KEY, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    if (res.getResponseCode() >= 300) {
      Logger.log(`[matchOrdererWithGemini] Gemini API Error: ${res.getContentText()}`);
      return null;
    }
    
    const json = JSON.parse(res.getContentText());
    const parts = json?.candidates?.[0]?.content?.parts || [];
    const out = parts.map(p => p.text || '').join('').trim();
    
    // コードブロック削除
    const cleaned = out.replace(/^```json/i, '').replace(/```$/i, '').trim();
    
    const result = JSON.parse(cleaned);
    
    if (result.matched_index !== null && result.matched_index !== undefined) {
      const matchedRecord = masterRecords[result.matched_index];
      if (matchedRecord) {
        writeLog('Geminiマッチング結果', { 
          ocrOrderer, 
          ocrBranch, 
          matchedIndex: result.matched_index,
          matchedRecordId: matchedRecord.record_id,
          confidence: result.confidence
        });
        return matchedRecord.record_id;
      }
    }
    
    writeLog('Geminiマッチング結果（一致なし）', { ocrOrderer, ocrBranch });
    return null;
  } catch(e) {
    Logger.log(`[matchOrdererWithGemini] Error: ${e}`);
    return null;
  }
};

/** Base: 発注者と支店・営業所名でマスターレコードを検索（Gemini API使用） */
const searchOrdererRecordId_ = (ocrOrderer, ocrBranch, masterRecords) => {
  if (!ocrOrderer) return null;
  
  if (!masterRecords || masterRecords.length === 0) {
    Logger.log(`[searchOrdererRecordId] マスターレコードが空です`);
    return null;
  }
  
  // Gemini APIでマッチング
  const matchedRecordId = matchOrdererWithGemini_(ocrOrderer, ocrBranch, masterRecords);
  
  return matchedRecordId;
};

/** Base: レコード検索してIDを返す（見つからなければnull）- 旧実装（後方互換性のため残す） */
const searchRecordId_ = (tenantTok, baseId, tableId, fieldName, value) => {
  if (!value) return null;
  
  const url = `${OPEN_API_HOST}/open-apis/bitable/v1/apps/${baseId}/tables/${tableId}/records/search`;
  const payload = {
    field_names: [], // IDだけあればいいのでrecord_idは返る
    filter: {
      conjunction: "and",
      conditions: [
        {
          field_name: fieldName,
          operator: "is",
          value: [value]
        }
      ]
    },
    automatic_fields: false
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { Authorization: `Bearer ${tenantTok}`, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  const json = JSON.parse(res.getContentText());
  if (json.code !== 0) {
    Logger.log(`[searchRecordId] Error: ${json.msg}`);
    return null;
  }
  
  const items = json.data?.items || [];
  if (items.length === 0) return null;
  return items[0].record_id; // 最初に見つかった1件を返す
};

function extractBaseAndTableIds(url) {
  // URL 全体からクエリ部分を分離
  const [pathPart, queryPart] = url.split('?');
  const query = queryPart || '';
  let tableId = '';
  // クエリから table パラメータを抽出
  query.split('&').forEach(pair => {
    const [k, v] = pair.split('=');
    if (k === 'table') {
      tableId = decodeURIComponent(v || '');
    }
  });
  const parts = pathPart.split('/').filter(s => s);
  const baseIndex = parts.findIndex(s => s.toLowerCase() === 'base');
  const baseId = (baseIndex !== -1 && baseIndex + 1 < parts.length)
    ? parts[baseIndex + 1]
    : '';
  return { baseId, tableId };
}


/* ========== メイン ========== */
function doPost(e) {
  const now = new Date();
  writeLog('doPost開始', { timestamp: now.toISOString() });
  
  const raw = e?.postData?.contents || '{}';
  let payload;
  try { 
    payload = JSON.parse(raw);
    writeLog('ペイロード受信', payload);
  } 
  catch(err) { 
    writeLog('JSON解析エラー', { raw, error: String(err) });
    return jsonResponse({ ok:false, error:'invalid JSON' }); 
  }

  const { baseId, tableId } = extractBaseAndTableIds(BASE_URL);
  
  if (!baseId || !tableId) {
    writeLog('BASE_URL解析失敗', { baseId, tableId, BASE_URL });
    return jsonResponse({ ok:false, error:'BASE_URL解析失敗' });
  }
  
  const base_id  = baseId;
  const table_id = tableId;
  // file_typeはファイルごとに判定するため、ここでは削除（後方互換性のためpayloadから取得可能）
  const defaultFileType = payload.file_type || null;
  
  // Baseオートメーションから送られてくるデータ形式に対応
  // file_token または attachmentToken の両方に対応
  let tokensArr = [];
  if (payload.file_token) {
    tokensArr = Array.isArray(payload.file_token) ? payload.file_token : [payload.file_token];
  } else if (payload.attachmentToken) {
    // Baseオートメーション拡張機能の添付ファイルコンポーネントから送られてくる場合
    tokensArr = Array.isArray(payload.attachmentToken) ? payload.attachmentToken : [payload.attachmentToken];
  }
  
  // ファイル名の取得（file_name または attachmentName）
  let fileNameArr = [];
  if (payload.file_name) {
    fileNameArr = Array.isArray(payload.file_name) ? payload.file_name : [payload.file_name];
  } else if (payload.attachmentName) {
    fileNameArr = Array.isArray(payload.attachmentName) ? payload.attachmentName : [payload.attachmentName];
  }
  
  // ファイル名が指定されていない場合、または配列の長さが一致しない場合は、デフォルト名を使用
  if (fileNameArr.length === 0 || fileNameArr.length !== tokensArr.length) {
    fileNameArr = tokensArr.map((_, i) => {
      if (i < fileNameArr.length && fileNameArr[i]) {
        return fileNameArr[i];
      }
      return `attachment_${i+1}.pdf`;
    });
  }

  if (!tokensArr.length) {
    writeLog('file_token/attachmentToken不足', { payload });
    return jsonResponse({ ok:false, error:'file_token or attachmentToken required' });
  }
  
  // サポート対象ファイルの拡張子一覧をログ出力
  const supportedExtensions = Object.keys(SUPPORTED_MIME_TYPES).join(', ');
  writeLog('ファイルトークン取得', { 
    tokenCount: tokensArr.length, 
    fileNameCount: fileNameArr.length,
    hasFileToken: !!payload.file_token,
    hasAttachmentToken: !!payload.attachmentToken,
    supportedFormats: supportedExtensions
  });

  try {
    writeLog('ファイルダウンロード開始', { fileCount: tokensArr.length });
    const tmpUrls = tokensArr.map(getTmpDownloadUrlByToken);
    const gemFileUris = [];
    const allRecords = [];

    tmpUrls.forEach((u,i) => {
      try {
        const fileName = fileNameArr[i] || `attachment_${i+1}.pdf`;
        // ファイル名からMIMEタイプを自動判定（payload指定があればそちらを優先）
        const mimeType = defaultFileType || getMimeTypeFromFileName(fileName);
        writeLog(`ファイル${i+1}処理開始`, { fileName, mimeType });
        
        const resp = UrlFetchApp.fetch(u, {
          fetchTimeout: 60 // タイムアウトを60秒に設定（大きなPDFファイルのダウンロードに対応）
        });
        if (resp.getResponseCode() !== 200) {
          throw new Error(`download failed: ${u} (${resp.getResponseCode()})`);
        }
        const bytes = resp.getContent();

        const file = geminiUploadFile_(bytes, mimeType, fileName);
        const records = geminiExtractFieldsFromFileUri_(file.uri, mimeType);
        writeLog(`ファイル${i+1} OCR完了`, { fileName, mimeType, recordsCount: records.length });
        
        gemFileUris.push(file.uri);
        allRecords.push(...records);
      } catch(err) {
        writeLog(`ファイル${i+1}処理エラー`, { fileName: fileNameArr[i], error: String(err) });
        throw err;
      }
    });

    writeLog('OCR処理完了', { totalRecords: allRecords.length });

    if (allRecords.length === 0) {
      writeLog('OCR結果が空', {});
      return jsonResponse({ ok:true, created: 0, message: 'OCR結果が空でした' });
    }

    // Baseに新規レコードを追加
    const tenantTok = getTenantAccessToken();
    const masterRecords = getAllMasterRecords_(tenantTok, base_id, ORDERER_TABLE_ID);
    
    let successCount = 0;
    let errorCount = 0;
    
    allRecords.forEach((r, index) => {
      try {
        // 発注者マスタからIDを検索（発注者と支店・営業所名の組み合わせでGemini APIマッチング）
        // キャッシュしたマスターレコードを使用
        const ordererId = searchOrdererRecordId_(r.orderer, r.branch_name, masterRecords);

        const fieldsObj = {
          "商品名": normalizeProductName(r.item_name) || "",
          "発注数量" : parseQuantity(r.quantity),
          "発注者" : ordererId ? [ordererId] : [],
          "お届け先（会社名）": r.destination || "",
          "注文番号": r.order_number || "",
          "支店・営業所名": r.branch_name || "",
          "郵便番号": r.postal_code || "",
          "住所詳細": r.address_detail || "",
          "電話番号": r.phone_number || "",
          "都道府県": r.prefecture || ""
        };
        
        createLarkRecord_(tenantTok, base_id, table_id, fieldsObj);
        successCount++;
      } catch(err) {
        writeLog(`レコード${index+1} 作成エラー`, { 
          index, 
          error: String(err), 
          record: r,
          stack: err.stack 
        });
        errorCount++;
        // エラーが発生しても続行（必要に応じて throw err で停止）
      }
    });

    writeLog('doPost完了', { 
      totalRecords: allRecords.length, 
      successCount, 
      errorCount 
    });
    
    return jsonResponse({ 
      ok: true, 
      created: successCount, 
      total: allRecords.length,
      errors: errorCount 
    });

  } catch(err) {
    writeLog('doPostエラー', { 
      error: String(err), 
      stack: err.stack 
    });
    return jsonResponse({ ok:false, error:String(err) });
  }
}


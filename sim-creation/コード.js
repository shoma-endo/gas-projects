/** @OnlyCurrentDoc
 * 推奨：スプレッドシートに「コンテナバインド」して使う
 * （シート→拡張機能→Apps Script）
 */

function doGet() {
  // 初期データを取得してHTMLに埋め込む（同期で取得）
  const outputType = "video"; // デフォルトは動画
  let initialData = {};
  try {
    initialData = getMediaMenus(outputType, false);
  } catch (e) {
    Logger.log('初期データの取得に失敗しました: ' + e);
  }

  const template = HtmlService.createTemplateFromFile('index');
  template.initialData = JSON.stringify(initialData);
  template.initialOutputType = outputType;
  
  return template
    .evaluate()
    .setTitle('SIM 入力フォーム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 行単位のデータを階層構造で返す
 * 媒体 → メニュー → KPI の階層で絞り込み可能な構造
 * CacheServiceで5分間キャッシュしてパフォーマンスと安定性を向上
 * @param {"video"|"static"} outputType
 * @param {boolean} forceRefresh キャッシュを無視して強制更新するか
 * @return {Object} { mediaKey: { name: "媒体名", menus: { menuKey: { name: "メニュー名", kpis: ["KPI1", "KPI2"] } } } }
 */
function getMediaMenus(outputType, forceRefresh = false) {
  const cacheKey = `mediaMenus_${outputType}`;
  const cache = CacheService.getScriptCache();
  
  // キャッシュから取得を試みる（強制更新でない場合）
  if (!forceRefresh) {
    const cached = cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
        Logger.log('キャッシュの解析に失敗しました: ' + e);
      }
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = outputType === "static" ? "静止画" : "動画";
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート『${sheetName}』が存在しません`);

  // 実際のデータ行数を取得（空行を考慮）
  const lastRow = sheet.getLastRow();
  const startRow = 3;
  const maxRows = Math.max(1000, lastRow - startRow + 1);
  
  // B列（媒体）、C列（メニュー）、E列（KPI）を取得
  // getValues()で値を取得することで、数式参照を値に変換して安定化
  const range = sheet.getRange(startRow, 2, Math.min(maxRows, lastRow - startRow + 1), 4);
  const values = range.getValues(); // B:E列（値として取得）
  
  const result = {};

  values.forEach(([rawMedia, rawMenu, , rawKpi]) => {
    // 値の正規化（数式の結果や空白を適切に処理）
    const mediaName = String(rawMedia || "").trim();
    const menuName = String(rawMenu || "").trim();
    const kpiName = String(rawKpi || "").trim();
    
    // 媒体・メニューが存在しない行はスキップ（空行除去）
    if (!mediaName || !menuName) return;

    const mediaKey = normalizeKey(mediaName);
    const menuKey = normalizeKey(menuName);

    // 媒体レベルの初期化
    if (!result[mediaKey]) {
      result[mediaKey] = { name: mediaName, menus: {} };
    }

    // メニューレベルの初期化
    if (!result[mediaKey].menus[menuKey]) {
      result[mediaKey].menus[menuKey] = { name: menuName, kpis: [] };
    }

    // KPIが存在する場合のみ重複チェックして追加
    if (kpiName && !result[mediaKey].menus[menuKey].kpis.includes(kpiName)) {
      result[mediaKey].menus[menuKey].kpis.push(kpiName);
    }
  });

  // 5分間キャッシュに保存（CacheServiceの最大サイズ制限を考慮してJSON文字列化）
  try {
    const cacheData = JSON.stringify(result);
    // CacheServiceの最大値は100KBなので、大きすぎる場合は分割保存しない（簡易実装）
    if (cacheData.length < 100000) {
      cache.put(cacheKey, cacheData, 300); // 300秒 = 5分
    }
  } catch (e) {
    Logger.log('キャッシュの保存に失敗しました: ' + e);
  }

  return result;
}


/** キー用に文字列を正規化する（日本語を保持） */
function normalizeKey(str) {
  return String(str).trim();
}

/** 送信ハンドラ（送信ボタンのクリックイベントで呼ばれる） */
function receivePlan(payload) {
  if (!payload) throw new Error('空のペイロードです');
  if (!payload.period || !payload.period.start || !payload.period.end) {
    throw new Error('期間が不正です');
  }
  const id = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss-SSS");

  // 保存ログ（Submissionsシート）
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Submissions');
  if (!sh) {
    sh = ss.insertSheet('Submissions');
    sh.getRange(1,1,1,3).setValues([['Timestamp','SubmissionId','Payload(JSON)']]);
  }
  sh.appendRow([new Date(), id, JSON.stringify(payload)]);

  // 本処理呼び出し（出力シートに書き出し）
  processPlan(payload);

  return { ok: true, id };
}

/** SIM出力処理 */
function processPlan(payload) {
  if (!payload || !payload.mediaPlans) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 動画の場合は複数シート、静止画の場合は1シートに出力
  const sheetNames = payload.outputType === "video"
    ? ["動画SIM出力", "動画SIM出力 v02", "動画SIM出力 v03"]
    : ["静止画SIM出力"];

  const formatDate = (dstr) => {
    if (!dstr) return "";
    const d = new Date(dstr);
    return `${d.getMonth() + 1}/${d.getDate()}`;
  };
  const periodStr = `${formatDate(payload.period.start)}-${formatDate(payload.period.end)}`;
  const marginPct = Number(payload.marginPct) || 0;

  // --- 行データの準備（各シート共通） ---
  const rows = payload.mediaPlans.map(plan => {
    const media = plan.media || "";
    const menu = plan.menu || "";
    const target = plan.target || "";
    const budget = Number(plan.budget) || 0;
    const kpi = plan.kpi || "";

    const rowData = getRowData(
      payload.outputType === "static" ? "静止画" : "動画",
      media, menu, kpi
    );
    const baseUnit = rowData ? rowData.unitCost : null;
    const ctrVal   = rowData ? rowData.ctr      : null;
    const duration = rowData ? rowData.duration : null;
    const completionRate = rowData ? rowData.completionRate : null;
    const viewDefinition = rowData ? rowData.viewDefinition : null;
    const viewUnitCost = rowData ? rowData.viewUnitCost : null;

    let unitCost = null;
    let cpmWithMargin = null;  // マージン率を考慮したCPM
    if (baseUnit !== null) {
      const rate = 1 - (marginPct / 100);
      // 静止画・動画ともにG列の値はCPM（1000回表示あたりの単価）なので、同じ計算式を使用
      unitCost = (baseUnit / 1000) / rate;
      // マージン率を考慮したCPM = 元のCPM / (1 - マージン率/100)
      cpmWithMargin = baseUnit / rate;
    }

    // 計算項目
    let impressions = null;          // 表示回数
    let clicks = null;               // クリック数
    let clickCost = null;            // クリック単価
    let completeViews = null;        // 完全視聴回数（動画）
    let completeViewCost = null;     // 完全視聴単価（動画）
    let estimatedViews = null;       // 想定視聴回数（動画 v03）
    let estimatedViewRate = null;    // 想定視聴率（動画 v03）
    let estimatedViewCost = null;    // 想定視聴単価（動画 v03）

    if (cpmWithMargin !== null && cpmWithMargin > 0) {
      impressions = (budget / cpmWithMargin) * 1000;  // 表示回数 = 予算 ÷ CPM × 1000

      // クリック数・クリック単価の計算（静止画・動画共通）
      if (ctrVal !== null && ctrVal > 0) {
        clicks = impressions * ctrVal;  // クリック数 = 表示回数 × CTR
        if (clicks > 0) {
          clickCost = budget / clicks;  // クリック単価 = 予算 / クリック数
        }
      }

      // 動画の場合：完全視聴回数・完全視聴単価の計算
      if (payload.outputType === "video") {
        if (completionRate !== null && completionRate > 0) {
          completeViews = impressions * completionRate;  // 完全視聴回数 = 表示回数 × 完全視聴率
          if (completeViews > 0) {
            completeViewCost = budget / completeViews;  // 完全視聴単価 = 予算 / 完全視聴回数
          }
        }

        // v03用：想定視聴回数・想定視聴率・想定視聴単価の計算
        if (viewUnitCost !== null && viewUnitCost > 0) {
          estimatedViews = budget / viewUnitCost;  // 想定視聴回数 = 予算 ÷ 視聴単価
          if (impressions !== null && impressions > 0) {
            estimatedViewRate = estimatedViews / impressions;  // 想定視聴率 = 想定視聴回数 ÷ 表示回数
          }
          if (estimatedViews > 0) {
            estimatedViewCost = budget / estimatedViews;  // 想定視聴単価 = 予算 ÷ 想定視聴回数
          }
        }
      }
    }

    return {
      media, menu, target, budget, unitCost, cpmWithMargin, ctrVal, kpi, duration, completionRate, viewDefinition,
      impressions, clicks, clickCost, completeViews, completeViewCost,
      estimatedViews, estimatedViewRate, estimatedViewCost
    };
  });

  // --- 各シートに出力 ---
  sheetNames.forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) {
      Logger.log(`警告: シート「${sheetName}」が見つかりません。スキップします。`);
      return;
    }

    // 3行目以降で最初の空行を探す
    const startRow = findFirstEmptyRow(sh);

    const rowNumbers = rows.map((_, idx) => startRow + idx);

    // --- 列ごとに書き込み ---
    writeColumn(sh, startRow, 1, rows.map(r => [r.media]));   // A 媒体
    writeColumn(sh, startRow, 2, rows.map(r => [r.menu]));    // B メニュー
    writeColumn(sh, startRow, 3, rows.map(_ => [periodStr])); // C 期間
    writeColumn(sh, startRow, 4, rows.map(r => [r.target]));  // D ターゲット

    if (payload.outputType === "static") {
      // 静止画出力
      const colBudget = columnToLetter(5);
      const colUnitCost = columnToLetter(7);  // G列の想定表示単価を参照
      const colImpressions = columnToLetter(6);
      const colCtr = columnToLetter(9);
      const colClicks = columnToLetter(8);

      writeColumn(sh, startRow, 5, rows.map(r => [r.budget]));      // E 予算
      writeColumn(
        sh,
        startRow,
        6,
        rowNumbers.map(rowNum => [
          `=IFERROR(IF(AND(ISNUMBER($${colBudget}${rowNum}),ISNUMBER($${colUnitCost}${rowNum}),$${colUnitCost}${rowNum}>0),$${colBudget}${rowNum}/$${colUnitCost}${rowNum},""),"")`
        ])
      ); // F 表示回数 = 予算 ÷ 想定表示単価（K列は使用しない）
      writeColumn(sh, startRow, 7, rows.map(r => [r.unitCost]));    // G 想定表示単価
      writeColumn(
        sh,
        startRow,
        8,
        rowNumbers.map(rowNum => [
          `=IFERROR(IF(AND(ISNUMBER($${colImpressions}${rowNum}),ISNUMBER($${colCtr}${rowNum})),$${colImpressions}${rowNum}*$${colCtr}${rowNum},""),"")`
        ])
      ); // H クリック数
      writeColumn(sh, startRow, 9, rows.map(r => [r.ctrVal]));      // I CTR
      writeColumn(
        sh,
        startRow,
        10,
        rowNumbers.map(rowNum => [
          `=IFERROR(IF(AND(ISNUMBER($${colClicks}${rowNum}),ISNUMBER($${colBudget}${rowNum}),$${colClicks}${rowNum}>0),$${colBudget}${rowNum}/$${colClicks}${rowNum},""),"")`
        ])
      ); // J クリック単価
      writeColumn(sh, startRow, 14, rows.map(r => [r.kpi]));        // N KPI
    } else {
      // 動画出力
      const colBudget = columnToLetter(6);
      const colCpm = sheetName === "動画SIM出力 v03" ? columnToLetter(8) : columnToLetter(13);  // v03シートはH列（想定表示単価）、それ以外はM列
      const colImpressions = columnToLetter(7);
      const colCompletionRate = columnToLetter(10);
      const colCompleteViews = columnToLetter(9);

      writeColumn(sh, startRow, 5, rows.map(r => [r.duration]));    // E 動画の尺
      writeColumn(sh, startRow, 6, rows.map(r => [r.budget]));      // F 予算
      // G列: 表示回数の計算（v03シートは想定表示単価を参照、それ以外はCPMを参照）
      if (sheetName === "動画SIM出力 v03") {
        // v03シート: 表示回数 = 予算 ÷ 想定表示単価（H列は1回あたりの単価）
        writeColumn(
          sh,
          startRow,
          7,
          rowNumbers.map(rowNum => [
            `=IFERROR(IF(AND(ISNUMBER($${colBudget}${rowNum}),ISNUMBER($${colCpm}${rowNum}),$${colCpm}${rowNum}>0),$${colBudget}${rowNum}/$${colCpm}${rowNum},""),"")`
          ])
        ); // G 表示回数 = 予算 ÷ 想定表示単価
      } else {
        // その他のシート: 表示回数 = 予算 ÷ CPM × 1000（CPMは1000回あたりの単価）
        writeColumn(
          sh,
          startRow,
          7,
          rowNumbers.map(rowNum => [
            `=IFERROR(IF(AND(ISNUMBER($${colBudget}${rowNum}),ISNUMBER($${colCpm}${rowNum}),$${colCpm}${rowNum}>0),$${colBudget}${rowNum}/$${colCpm}${rowNum}*1000,""),"")`
          ])
        ); // G 表示回数 = 予算 ÷ CPM × 1000
      }
      writeColumn(sh, startRow, 8, rows.map(r => [r.unitCost]));    // H 想定表示単価
      
      // v03シート以外の場合は通常のI列、J列、K列を出力
      if (sheetName !== "動画SIM出力 v03") {
        writeColumn(
          sh,
          startRow,
          9,
          rowNumbers.map(rowNum => [
            `=IFERROR(IF(AND(ISNUMBER($${colImpressions}${rowNum}),ISNUMBER($${colCompletionRate}${rowNum})),$${colImpressions}${rowNum}*$${colCompletionRate}${rowNum},""),"")`
          ])
        ); // I 完全視聴回数
        writeColumn(sh, startRow, 10, rows.map(r => [r.completionRate]));    // J 完全視聴率
        writeColumn(
          sh,
          startRow,
          11,
          rowNumbers.map(rowNum => [
            `=IFERROR(IF(AND(ISNUMBER($${colCompleteViews}${rowNum}),ISNUMBER($${colBudget}${rowNum}),$${colCompleteViews}${rowNum}>0),$${colBudget}${rowNum}/$${colCompleteViews}${rowNum},""),"")`
          ])
        ); // K 完全視聴単価
      }
      
      // v03シート以外の場合のみM列にCPMを出力（v03シートではP列に出力）
      if (sheetName !== "動画SIM出力 v03") {
        writeColumn(sh, startRow, 13, rows.map(r => [r.cpmWithMargin]));    // M CPM（マージン率を考慮、表示回数計算用、v02以外）
      }

      // v02シートのみ追加項目を出力
      if (sheetName === "動画SIM出力 v02") {
        const colCpmV02 = columnToLetter(16);  // P列にCPM値を出力（v02専用）
        const colCtr = columnToLetter(12);
        const colClicks = columnToLetter(13);
        const colBudgetV02 = columnToLetter(6);
        const colImpressionsV02 = columnToLetter(7);

        // v02シートの表示回数はP列のCPMを参照するように上書き
        writeColumn(
          sh,
          startRow,
          7,
          rowNumbers.map(rowNum => [
            `=IFERROR(IF(AND(ISNUMBER($${colBudgetV02}${rowNum}),ISNUMBER($${colCpmV02}${rowNum}),$${colCpmV02}${rowNum}>0),$${colBudgetV02}${rowNum}/$${colCpmV02}${rowNum}*1000,""),"")`
          ])
        ); // G 表示回数 = 予算 ÷ CPM × 1000（P列のCPMを参照）
        writeColumn(sh, startRow, 16, rows.map(r => [r.cpmWithMargin]));    // P CPM（マージン率を考慮、表示回数計算用）

        writeColumn(sh, startRow, 12, rows.map(r => [r.ctrVal]));        // L クリック率
        writeColumn(
          sh,
          startRow,
          13,
          rowNumbers.map(rowNum => [
            `=IFERROR(IF(AND(ISNUMBER($${colImpressions}${rowNum}),ISNUMBER($${colCtr}${rowNum})),$${colImpressions}${rowNum}*$${colCtr}${rowNum},""),"")`
          ])
        ); // M クリック数
        writeColumn(
          sh,
          startRow,
          14,
          rowNumbers.map(rowNum => [
            `=IFERROR(IF(AND(ISNUMBER($${colClicks}${rowNum}),ISNUMBER($${colBudget}${rowNum}),$${colClicks}${rowNum}>0),$${colBudget}${rowNum}/$${colClicks}${rowNum},""),"")`
          ])
        ); // N クリック単価
        writeColumn(sh, startRow, 15, rows.map(r => [r.viewDefinition]));// O 視聴定義
        writeColumn(sh, startRow, 17, rows.map(r => [r.kpi]));           // Q KPI
      } else if (sheetName === "動画SIM出力 v03") {
        // v03シートのみ追加項目を出力
        const colBudgetV03 = columnToLetter(6);
        const colEstimatedViews = columnToLetter(9);  // I列: 想定視聴回数
        const colImpressionsV03 = columnToLetter(7);  // G列: 表示回数
        const colCompleteViewsV03 = columnToLetter(12);  // L列: 完全視聴回数
        const colCompletionRateV03 = columnToLetter(13);  // M列: 完全視聴率

        // I列: 想定視聴回数 = 予算 ÷ 視聴単価（動画シートのK列の値を使用して計算済みの値）
        writeColumn(sh, startRow, 9, rows.map(r => [r.estimatedViews]));  // I 想定視聴回数

        // J列: 想定視聴率 = 想定視聴回数 ÷ 表示回数
        writeColumn(
          sh,
          startRow,
          10,
          rowNumbers.map(rowNum => [
            `=IFERROR(IF(AND(ISNUMBER($${colEstimatedViews}${rowNum}),ISNUMBER($${colImpressionsV03}${rowNum}),$${colImpressionsV03}${rowNum}>0),$${colEstimatedViews}${rowNum}/$${colImpressionsV03}${rowNum},""),"")`
          ])
        ); // J 想定視聴率

        // K列: 想定視聴単価 = 予算 ÷ 想定視聴回数
        writeColumn(
          sh,
          startRow,
          11,
          rowNumbers.map(rowNum => [
            `=IFERROR(IF(AND(ISNUMBER($${colEstimatedViews}${rowNum}),ISNUMBER($${colBudgetV03}${rowNum}),$${colEstimatedViews}${rowNum}>0),$${colBudgetV03}${rowNum}/$${colEstimatedViews}${rowNum},""),"")`
          ])
        ); // K 想定視聴単価

        // L列: 完全視聴回数（元I列から右にシフト）
        writeColumn(
          sh,
          startRow,
          12,
          rowNumbers.map(rowNum => [
            `=IFERROR(IF(AND(ISNUMBER($${colImpressionsV03}${rowNum}),ISNUMBER($${colCompletionRateV03}${rowNum})),$${colImpressionsV03}${rowNum}*$${colCompletionRateV03}${rowNum},""),"")`
          ])
        ); // L 完全視聴回数

        // M列: 完全視聴率（元J列から右にシフト）
        writeColumn(sh, startRow, 13, rows.map(r => [r.completionRate]));  // M 完全視聴率

        // N列: 完全視聴単価（元K列から右にシフト）
        writeColumn(
          sh,
          startRow,
          14,
          rowNumbers.map(rowNum => [
            `=IFERROR(IF(AND(ISNUMBER($${colCompleteViewsV03}${rowNum}),ISNUMBER($${colBudgetV03}${rowNum}),$${colCompleteViewsV03}${rowNum}>0),$${colBudgetV03}${rowNum}/$${colCompleteViewsV03}${rowNum},""),"")`
          ])
        ); // N 完全視聴単価

        // O列、P列、Q列: クリック関連
        const colCtrV03 = columnToLetter(15);  // O列: クリック率
        const colClicksV03 = columnToLetter(16);  // P列: クリック数
        writeColumn(sh, startRow, 15, rows.map(r => [r.ctrVal]));        // O クリック率
        writeColumn(
          sh,
          startRow,
          16,
          rowNumbers.map(rowNum => [
            `=IFERROR(IF(AND(ISNUMBER($${colImpressionsV03}${rowNum}),ISNUMBER($${colCtrV03}${rowNum})),$${colImpressionsV03}${rowNum}*$${colCtrV03}${rowNum},""),"")`
          ])
        ); // P クリック数
        writeColumn(
          sh,
          startRow,
          17,
          rowNumbers.map(rowNum => [
            `=IFERROR(IF(AND(ISNUMBER($${colClicksV03}${rowNum}),ISNUMBER($${colBudgetV03}${rowNum}),$${colClicksV03}${rowNum}>0),$${colBudgetV03}${rowNum}/$${colClicksV03}${rowNum},""),"")`
          ])
        ); // Q クリック単価
        writeColumn(sh, startRow, 18, rows.map(r => [r.viewDefinition]));  // R 視聴定義
        // S列: 入稿締切日（コードでは出力しない）
        writeColumn(sh, startRow, 20, rows.map(r => [r.kpi]));             // T KPI
      } else {
        writeColumn(sh, startRow, 12, rows.map(r => [r.viewDefinition]));// L 視聴定義
        writeColumn(sh, startRow, 14, rows.map(r => [r.kpi]));           // N KPI
      }
    }

    // --- フォーマット設定 ---
    if (payload.outputType === "static") {
      sh.getRange(startRow, 5, rows.length, 1).setNumberFormat("¥#,##0");    // E 予算
      sh.getRange(startRow, 6, rows.length, 1).setNumberFormat("#,##0");     // F 表示回数
      sh.getRange(startRow, 7, rows.length, 1).setNumberFormat("¥#,##0.00"); // G 想定表示単価
      sh.getRange(startRow, 8, rows.length, 1).setNumberFormat("#,##0");     // H クリック数
      sh.getRange(startRow, 9, rows.length, 1).setNumberFormat("0.00%");     // I CTR
      sh.getRange(startRow, 10, rows.length, 1).setNumberFormat("¥#,##0.00");// J クリック単価
    } else {
      sh.getRange(startRow, 6, rows.length, 1).setNumberFormat("¥#,##0");    // F 予算
      sh.getRange(startRow, 7, rows.length, 1).setNumberFormat("#,##0");     // G 表示回数
      sh.getRange(startRow, 8, rows.length, 1).setNumberFormat("¥#,##0.00"); // H 想定表示単価
      
      // v03シート以外の場合は通常のI列、J列、K列のフォーマットを設定
      if (sheetName !== "動画SIM出力 v03") {
        sh.getRange(startRow, 9, rows.length, 1).setNumberFormat("#,##0");     // I 完全視聴回数
        sh.getRange(startRow, 10, rows.length, 1).setNumberFormat("0.00%");    // J 完全視聴率
        sh.getRange(startRow, 11, rows.length, 1).setNumberFormat("¥#,##0.00");// K 完全視聴単価
      }

      // v02シートのみ追加項目のフォーマット設定
      if (sheetName === "動画SIM出力 v02") {
        sh.getRange(startRow, 12, rows.length, 1).setNumberFormat("0.00%");     // L クリック率
        sh.getRange(startRow, 13, rows.length, 1).setNumberFormat("#,##0");     // M クリック数
        sh.getRange(startRow, 14, rows.length, 1).setNumberFormat("¥#,##0.00"); // N クリック単価
      } else if (sheetName === "動画SIM出力 v03") {
        // v03シートのみ追加項目のフォーマット設定
        sh.getRange(startRow, 9, rows.length, 1).setNumberFormat("#,##0");      // I 想定視聴回数
        sh.getRange(startRow, 10, rows.length, 1).setNumberFormat("0.00%");     // J 想定視聴率
        sh.getRange(startRow, 11, rows.length, 1).setNumberFormat("¥#,##0.00"); // K 想定視聴単価
        sh.getRange(startRow, 12, rows.length, 1).setNumberFormat("#,##0");     // L 完全視聴回数
        sh.getRange(startRow, 13, rows.length, 1).setNumberFormat("0.00%");     // M 完全視聴率
        sh.getRange(startRow, 14, rows.length, 1).setNumberFormat("¥#,##0.00"); // N 完全視聴単価
        sh.getRange(startRow, 15, rows.length, 1).setNumberFormat("0.00%");     // O クリック率
        sh.getRange(startRow, 16, rows.length, 1).setNumberFormat("#,##0");     // P クリック数
        sh.getRange(startRow, 17, rows.length, 1).setNumberFormat("¥#,##0.00"); // Q クリック単価
      }
    }
  });
}

/** 指定列に values を一括書き込み */
function writeColumn(sheet, startRow, col, values) {
  if (!values || values.length === 0) return;
  sheet.getRange(startRow, col, values.length, 1).setValues(values);
}

/** 列番号をスプレッドシートの列記号に変換 */
function columnToLetter(columnNumber) {
  let temp = "";
  let letter = "";
  let colNum = Math.floor(columnNumber);
  if (!colNum || colNum < 1) return "";

  while (colNum > 0) {
    temp = (colNum - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colNum = Math.floor((colNum - temp - 1) / 26);
  }

  return letter;
}

/**
 * 3行目以降の最初の空行を探す
 * A列（媒体）が空なら「空行」とみなす
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @return {number} 書き込み開始行
 */
function findFirstEmptyRow(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return 3;

  const values = sh.getRange(3, 1, lastRow - 2, 1).getValues(); // A列チェック (3行目以降)
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) {
      return 3 + i; // 空セルが見つかった行
    }
  }
  return lastRow + 1; // 空きがなければ末尾に追記
}

/**
 * シートから媒体+メニュー+KPIに対応する行データを取得
 * @param {string} sheetName "静止画" または "動画"
 * @param {string} media 媒体名
 * @param {string} menu メニュー名
 * @param {string} kpi KPI名
 * @return {{unitCost:number|null, ctr:number|null, duration:string|null, completionRate:number|null, viewDefinition:string|null, viewUnitCost:number|null}}
 */
function getRowData(sheetName, media, menu, kpi) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`シート「${sheetName}」が見つかりません`);

  const lastRow = sh.getLastRow();
  if (lastRow < 3) return { unitCost: null, ctr: null, duration: null, completionRate: null, viewDefinition: null, viewUnitCost: null };

  // B列から始まり、S列（19列目）まで取得
  const values = sh.getRange(3, 2, lastRow - 2, 18).getValues();

  for (const row of values) {
    const mediaVal = String(row[0]).trim();  // B列: 媒体
    const menuVal  = String(row[1]).trim();  // C列: メニュー
    const kpiVal   = String(row[3]).trim();  // E列: KPI
    const durationVal = row[2];              // D列: 動画の尺
    const unitVal  = row[5];                 // G列: CPM/表示単価
    const ctrVal   = row[6];                 // H列: CTR
    const viewUnitCostVal = row[9];          // K列: 視聴単価
    const completionRateVal = row[10];       // L列: 完全視聴率
    const viewDefinitionVal = row[17];       // S列: 視聴定義

    if (mediaVal === media && menuVal === menu && kpiVal === kpi) {
      // CPM/表示単価
      let numericUnit = null;
      if (unitVal !== null && unitVal !== "") {
        const parsed = parseFloat(String(unitVal).replace(/[^\d.]/g, ""));
        numericUnit = isNaN(parsed) ? null : parsed;
      }
      // CTR（数値として返す。例: 0.0003 = 0.03%）
      let numericCtr = null;
      if (ctrVal !== null && ctrVal !== "") {
        if (typeof ctrVal === "number") {
          numericCtr = ctrVal;
        } else {
          const parsed = parseFloat(String(ctrVal).replace(/[^\d.]/g, ""));
          numericCtr = isNaN(parsed) ? null : parsed / 100;
        }
      }

      // 視聴単価（数値として返す）
      let numericViewUnitCost = null;
      if (viewUnitCostVal !== null && viewUnitCostVal !== "") {
        const parsed = parseFloat(String(viewUnitCostVal).replace(/[^\d.]/g, ""));
        numericViewUnitCost = isNaN(parsed) ? null : parsed;
      }

      // 完全視聴率（数値として返す）
      let numericCompletionRate = null;
      if (completionRateVal !== null && completionRateVal !== "") {
        if (typeof completionRateVal === "number") {
          numericCompletionRate = completionRateVal;
        } else {
          const parsed = parseFloat(String(completionRateVal).replace(/[^\d.]/g, ""));
          numericCompletionRate = isNaN(parsed) ? null : parsed / 100;
        }
      }

      return {
        unitCost: numericUnit,
        ctr: numericCtr,
        duration: durationVal ? String(durationVal).trim() : null,
        completionRate: numericCompletionRate,
        viewDefinition: viewDefinitionVal ? String(viewDefinitionVal).trim() : null,
        viewUnitCost: numericViewUnitCost
      };
    }
  }

  return { unitCost: null, ctr: null, duration: null, completionRate: null, viewDefinition: null, viewUnitCost: null };
}


/** 予算フォーマット（整数のみ、￥付き、カンマ区切り） */
function formatYen(num) {
  if (!num || isNaN(num)) return "";
  return "￥" + Math.round(num).toLocaleString('ja-JP');
}

function onOpen() {
SpreadsheetApp.getUi()
  .createMenu('SIMツール')
  .addItem('ダイアログで開く', 'openDialog')
  .addToUi();
}

function openDialog() {
  // 初期データを取得してHTMLに埋め込む（同期で取得）
  const outputType = "video"; // デフォルトは動画
  let initialData = {};
  try {
    initialData = getMediaMenus(outputType, false);
  } catch (e) {
    Logger.log('初期データの取得に失敗しました: ' + e);
  }

  const template = HtmlService.createTemplateFromFile('index');
  template.initialData = JSON.stringify(initialData);
  template.initialOutputType = outputType;
  
  const html = template
    .evaluate()
    .setWidth(1100)
    .setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, 'SIM 入力フォーム'); // モーダル表示
}

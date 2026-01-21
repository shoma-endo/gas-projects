function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('入稿シート自動作成フォーム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** HTMLテンプレートから部分をインクルードするためのヘルパー */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** 送信ハンドラ（送信ボタンのクリックイベントで呼ばれる） */
function receivePlan(payload) {
  if (!payload) throw new Error('空のペイロードです');
  if (!payload.requestText) throw new Error('与件情報が入力されていません');
  if (!payload.clientAgency) throw new Error('クライアント代理店が選択されていません');
  if (!payload.mediaList || payload.mediaList.length === 0) {
    throw new Error('媒体が選択されていません');
  }

  const id = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss-SSS");
  const startTime = Date.now();

  // 保存ログ（Submissionsシート）- 拡張版
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Submissions');
  if (!sh) {
    sh = ss.insertSheet('Submissions');
    sh.getRange(1, 1, 1, 10).setValues([[
      'Timestamp',
      'SubmissionId',
      'RequestText',
      'ClientAgency',
      'SelectedMedias',
      'ExtractedMedias',
      'Success',
      'Skipped',
      'Errors',
      'ElapsedMs'
    ]]);
  }

  // 処理実行
  let result;
  let extractedMedias = '';
  let successCount = 0;
  let skippedCount = 0;
  let errorCount = 0;
  let elapsed = 0;

  // mediaListを正しくフォーマット
  const selectedMediasStr = (payload.mediaList || [])
    .map(m => typeof m === 'string' ? m : (m.mediaId || ''))
    .filter(m => m.trim())
    .join(', ');

  try {
    result = processSubmission(payload, id); // IDを渡す
    elapsed = Date.now() - startTime;

    // 結果を集計
    if (result && result.results) {
      successCount = result.results.success.length;
      skippedCount = result.results.skipped.length;
      errorCount = result.results.errors.length;

      // 抽出された媒体リスト
      const allMedias = [
        ...result.results.success.map(m => m.mediaId),
        ...result.results.skipped.map(m => m.mediaId),
        ...result.results.errors.map(m => m.mediaId)
      ];
      extractedMedias = allMedias.join(', ');
    }

  } catch (e) {
    elapsed = Date.now() - startTime;
    errorCount = 1;
    Logger.log(`❌ receivePlan error [${id}]: ${e.message}`);

    // エラー情報をログに記録
    sh.appendRow([
      new Date(),
      id,
      payload.requestText,
      payload.clientAgency,
      selectedMediasStr,
      extractedMedias,
      0,
      0,
      1,
      elapsed
    ]);

    throw e; // 再スロー
  }

  // 正常終了時のログ記録
  sh.appendRow([
    new Date(),
    id,
    payload.requestText,
    payload.clientAgency,
    selectedMediasStr,
    extractedMedias,
    successCount,
    skippedCount,
    errorCount,
    elapsed
  ]);

  // 成功が1件もない場合は失敗扱い
  const hasErrors = errorCount > 0;
  const hasSkipped = skippedCount > 0;
  const noSuccess = successCount === 0;

  return {
    ok: !noSuccess, // 成功が1件もない場合はfalse
    id,
    results: result ? result.results : null,
    unmapped_notes: result ? result.unmapped_notes : null,
    hasPartialErrors: (hasErrors || hasSkipped) && successCount > 0 // 部分的失敗/スキップフラグ
  };
}

/**
 * フォーム送信データを処理（ワンショットAI抽出版）
 * @param {Object} payload - フォームからの送信データ
 * @param {string} submissionId - receivePlanで発行されたSubmission ID
 */
function processSubmission(payload, submissionId) {
  const startTime = Date.now();

  Logger.log(`🚀 processSubmission開始 [${submissionId}]`);

  // ====== 1. ワンショットでAI抽出 ======
  let aiResult;
  try {
    aiResult = analyzeTextAI(payload.requestText, submissionId);
  } catch (e) {
    Logger.log(`❌ AI抽出エラー [${submissionId}]: ${e.message}`);
    throw new Error(`AI抽出に失敗しました: ${e.message}`);
  }

  if (!aiResult || !Array.isArray(aiResult.medias) || aiResult.medias.length === 0) {
    Logger.log(`⚠️ AIが媒体を抽出できませんでした [${submissionId}]`);
    throw new Error('AIが媒体情報を抽出できませんでした。与件文を確認してください。');
  }

  Logger.log(`📊 抽出された媒体数: ${aiResult.medias.length}`);

  // ====== 2. ユーザー選択媒体との突合 ======
  const selectedMediaIds = (payload.mediaList || [])
    .map(entry => typeof entry === 'string' ? entry : (entry.mediaId || ''))
    .filter(name => name.trim());

  // ユーザーが選択した媒体に絞り込む（選択がない場合は全抽出媒体を使用）
  let targetMedias = aiResult.medias;
  if (selectedMediaIds.length > 0) {
    targetMedias = aiResult.medias.filter(media =>
      selectedMediaIds.includes(media.mediaId)
    );

    // 選択されたが抽出されなかった媒体を警告
    const extractedIds = aiResult.medias.map(m => m.mediaId);
    const missingMedias = selectedMediaIds.filter(id => !extractedIds.includes(id));
    if (missingMedias.length > 0) {
      Logger.log(`⚠️ 選択されたが抽出されなかった媒体: ${missingMedias.join(', ')}`);
    }
  }

  if (targetMedias.length === 0) {
    Logger.log(`⚠️ 処理対象の媒体がありません [${submissionId}]`);
    throw new Error('選択された媒体に該当する情報が抽出されませんでした。');
  }

  // ====== 3. 各媒体ごとにシート書き込み ======
  const results = {
    success: [],
    skipped: [],
    errors: []
  };

  targetMedias.forEach((media) => {
    const mediaId = media.mediaId;

    try {
      // MEDIA_CONFIGに存在確認
      const config = MEDIA_CONFIG[mediaId];
      if (!config) {
        results.skipped.push({
          mediaId,
          reason: 'MEDIA_CONFIGに未定義'
        });
        Logger.log(`⚠️ スキップ: ${mediaId} (MEDIA_CONFIGに存在しません)`);
        return;
      }

      Logger.log(`📋 ${mediaId}: MEDIA_CONFIG参照 → シート: ${config.sheet}, promap: ${config.promap}`);

      // シート判別
      const objSS = {
        adsheet: config.sheet,
        adpromap: config.promap,
        adsheetflag: 0
      };

      // 【自動マスタ同期】書き込み前に、最新のプルダウン検証ルールを適用
      // → なぜ: マスタシートの候補値とテンプレートシートのプルダウンを同期し、
      //         書き込み時のデータ検証エラーを防ぐため
      try {
        const ss = SpreadsheetApp.getActive();
        const masterSheetName = `【マスタ】${config.sheet}`;
        const masterSheet = ss.getSheetByName(masterSheetName);
        const templateSheet = ss.getSheetByName(config.sheet);

        if (masterSheet && templateSheet) {
          Logger.log(`🔄 書き込み前の自動マスタ同期: ${masterSheetName} → ${config.sheet}`);
          syncMasterToTemplate(masterSheet, templateSheet);
          Logger.log(`✅ 書き込み前の自動マスタ同期完了`);
        } else {
          if (!masterSheet) {
            Logger.log(`⚠️ マスタシート「${masterSheetName}」が見つかりません（同期スキップ）`);
          }
          if (!templateSheet) {
            Logger.log(`⚠️ テンプレートシート「${config.sheet}」が見つかりません（同期スキップ）`);
          }
        }
      } catch (syncError) {
        Logger.log(`⚠️ 書き込み前の自動マスタ同期エラー: ${syncError.message}`);
        // エラーがあっても書き込み処理は継続
      }

      // 【なぜ】フォームデータを取得
      // → payload.mediaListから、この媒体に対応するフォーム入力値を探す
      // → フォームで選択された値を「ディスカバリー運用記入」列に書き込むため
      let formFields = {};
      if (payload.mediaList && Array.isArray(payload.mediaList)) {
        // 【なぜこのログが必要か】
        // - ブラウザから送信されたペイロードがサーバーに正しく到達しているかを確認するため
        // - 問題: ブラウザ側のコンソールでは正常でも、サーバー側で受信できていない可能性がある
        // - 正常な場合: payload.mediaList は [{ mediaId: "...", "キャンペーン目標": "...", ... }] のような配列
        // - もし payload.mediaList が空配列 [] または undefined の場合、送信処理に問題がある
        // - もし payload.mediaList[0] に動的フィールドが含まれていない場合、ブラウザ側の buildPayload() に問題がある
        Logger.log(`  📦 payload.mediaList: ${JSON.stringify(payload.mediaList)}`);

        const formMedia = payload.mediaList.find(m => m.mediaId === mediaId);
        if (formMedia) {
          // 【なぜこのログが必要か】
          // - payload.mediaList から正しく mediaId で絞り込めているかを確認するため
          // - 問題: 複数の媒体がある場合、正しい媒体データが取得できているか確認
          // - 正常な場合: formMedia は { mediaId: "Googleリスティング", "キャンペーン目標": "売上", ... } のような形
          // - もし formMedia に mediaId しかない場合、ブラウザ側でのフィールド展開が失敗している
          Logger.log(`  🎯 見つかった媒体データ: ${JSON.stringify(formMedia)}`);

          // 【なぜ】mediaId以外のプロパティをformFieldsとして抽出
          // → フォームで入力された動的フィールドの値を取得するため
          const { mediaId: _, ...fields } = formMedia;
          formFields = fields;

          // 【なぜこのログが必要か】
          // - デストラクチャリング（分割代入）が正しく動作しているかを確認するため
          // - 問題: { mediaId: _, ...fields } という構文で、mediaId以外のプロパティが fields に抽出されているか確認
          // - 正常な場合: formFields は { "キャンペーン目標": "売上", "入札戦略": "目標コンバージョン単価" } のような形（mediaIdを除く）
          // - もし formFields が空オブジェクト {} の場合、formMedia に動的フィールドが含まれていない
          // - このログが正常でも書き込まれない場合、insertDataFromStructured() の処理に問題がある
          Logger.log(`  📝 抽出されたformFields: ${JSON.stringify(formFields)}`);
        } else {
          Logger.log(`  ⚠️ payload.mediaListに媒体「${mediaId}」が見つかりません`);
        }
      } else {
        Logger.log(`  ⚠️ payload.mediaListが存在しないか、配列ではありません`);
      }

      // 【なぜ】書き込み処理に3つの引数を渡す
      // 1. mediaId: 媒体ID
      // 2. media.fields: AIが抽出したフィールド（代理店記入列用）
      // 3. formFields: フォームで入力されたフィールド（ディスカバリー運用記入列用）
      insertDataFromStructured(mediaId, media.fields, formFields, objSS);

      results.success.push({
        mediaId,
        confidence: media.confidence,
        sheet: config.sheet
      });

      Logger.log(`✅ 成功: ${mediaId} → ${config.sheet}`);

    } catch (e) {
      results.errors.push({
        mediaId,
        error: e.message
      });
      Logger.log(`❌ エラー: ${mediaId} - ${e.message}`);
    }
  });

  // ====== 4. 処理結果サマリー ======
  const elapsed = Date.now() - startTime;
  Logger.log(`\n📈 処理完了サマリー [${submissionId}] ${elapsed}ms`);
  Logger.log(`  成功: ${results.success.length}件`);
  Logger.log(`  スキップ: ${results.skipped.length}件`);
  Logger.log(`  エラー: ${results.errors.length}件`);

  if (aiResult.unmapped_notes) {
    Logger.log(`  備考: ${aiResult.unmapped_notes}`);
  }

  // エラーがある場合はユーザーに通知
  if (results.errors.length > 0) {
    const errorMsg = results.errors.map(e => `${e.mediaId}: ${e.error}`).join('\n');
    Logger.log(`⚠️ 一部の媒体でエラーが発生:\n${errorMsg}`);
  }

  return {
    submissionId,
    results,
    unmapped_notes: aiResult.unmapped_notes
  };
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // 入稿シート自動作成メニュー
  ui.createMenu('入稿シート自動作成')
    .addItem('ダイアログで開く', 'openDialog')
    .addItem('READMEを表示', 'openReadmeDialog')
    .addToUi();

  // 入稿メールメニュー
  ui.createMenu('入稿メール')
    .addItem('チェック行 → 即時送信', 'sendCheckedRows')
    .addToUi();
}

function openDialog() {
  const html = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setWidth(1100)
    .setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, '入稿シート自動作成フォーム');
}

function openReadmeDialog() {
  const html = HtmlService.createTemplateFromFile('readme_view')
    .evaluate()
    .setWidth(960)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'README');
}

/**
 * 【見本】マスタシートからクライアント代理店リストを取得
 * @returns {string[]} - クライアント代理店のリスト
 */
function getClientAgencyList() {
  try {
    const ss = SpreadsheetApp.getActive();
    const masterSheet = ss.getSheetByName('【見本】マスタ');

    if (!masterSheet) {
      Logger.log('⚠️ 【見本】マスタシートが見つかりません');
      return [];
    }

    // ヘッダー行（1行目）を取得して「クライアント代理店」列を探す
    const headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    const agencyColIndex = headers.findIndex(h => h === 'クライアント代理店');

    if (agencyColIndex === -1) {
      Logger.log('⚠️ 【見本】マスタシートに「クライアント代理店」列が見つかりません');
      return [];
    }

    // 2行目以降のデータを取得（1ベースから0ベースに変換するため +1）
    const dataRange = masterSheet.getRange(2, agencyColIndex + 1, masterSheet.getLastRow() - 1, 1);
    const values = dataRange.getValues();

    // 空でない値のみをフィルタリング
    const agencyList = values
      .map(row => String(row[0]).trim())
      .filter(val => val.length > 0);

    Logger.log(`✅ クライアント代理店リストを取得: ${agencyList.length}件`);
    return agencyList;

  } catch (e) {
    Logger.log(`❌ getClientAgencyList error: ${e.message}`);
    return [];
  }
}

/**
 * 媒体別のマスタシートデータを取得（フォーム動的化用）
 *
 * 【なぜこの関数が必要か】
 * - フォーム上で媒体を選択した際に、その媒体専用のフィールド定義を動的に取得する必要がある
 * - 固定的なフィールド定義（MEDIA_FIELD_DEFS）ではなく、マスタシートから最新の定義を取得することで、
 *   マスタシートを更新するだけでフォームに反映される柔軟な設計を実現する
 *
 * 【処理の流れ】
 * 1. MEDIA_CONFIGから対応するテンプレートシート名を取得
 * 2. 「【マスタ】」+ テンプレートシート名 でマスタシートを特定
 * 3. マスタシートの1行目（ヘッダー）と2行目以降（候補値）を取得
 * 4. フロントエンドに返すデータ構造に整形
 *
 * @param {string} mediaId - 媒体ID（例: "Googleリスティング", "YDA"）
 * @returns {Object} - { success, mediaId, masterSheetName, columns: [{header, options, hasOptions}] }
 */
function getMasterSheetData(mediaId) {
  try {
    Logger.log(`📊 getMasterSheetData開始: ${mediaId}`);

    // 【なぜ】MEDIA_CONFIGから対応するシート名を取得
    // → 媒体IDからテンプレートシート名への変換を統一的に行うため
    const config = MEDIA_CONFIG[mediaId];
    if (!config) {
      Logger.log(`⚠️ MEDIA_CONFIGに媒体「${mediaId}」が見つかりません`);
      return {
        success: false,
        error: `媒体「${mediaId}」は未定義です`
      };
    }

    const templateSheetName = config.sheet;
    const masterSheetName = `【マスタ】${templateSheetName}`;

    Logger.log(`  テンプレートシート: ${templateSheetName}`);
    Logger.log(`  マスタシート: ${masterSheetName}`);

    const ss = SpreadsheetApp.getActive();

    // 【重要な変更】テンプレートシートからフィールド名を取得
    // → なぜ: 書き込み時はテンプレートシートの5-6行目の動的検出を使っているため
    // → マスタシートはプルダウン候補値の取得にのみ使用する
    const templateSheet = ss.getSheetByName(templateSheetName);
    if (!templateSheet) {
      Logger.log(`⚠️ テンプレートシート「${templateSheetName}」が見つかりません`);
      return {
        success: false,
        error: `テンプレートシート「${templateSheetName}」が存在しません`
      };
    }

    const masterSheet = ss.getSheetByName(masterSheetName);
    if (!masterSheet) {
      Logger.log(`⚠️ マスタシート「${masterSheetName}」が見つかりません`);
      return {
        success: false,
        error: `マスタシート「${masterSheetName}」が存在しません`
      };
    }

    // 【自動マスタ同期】フォームで媒体を選択したときに、最新のプルダウン候補を取得するため
    // → なぜ: ユーザーが手動で「マスタ同期」メニューを実行しなくても、常に最新の状態を保証
    // → マスタシートの候補値をテンプレートシートのプルダウンに反映
    try {
      Logger.log(`🔄 自動マスタ同期開始: ${masterSheetName} → ${templateSheetName}`);
      syncMasterToTemplate(masterSheet, templateSheet);
      Logger.log(`✅ 自動マスタ同期完了`);
    } catch (syncError) {
      Logger.log(`⚠️ 自動マスタ同期エラー: ${syncError.message}`);
      // エラーがあってもフォーム表示は継続
    }

    // 【重要な変更】マスタシート基準でフォームを生成
    // → なぜ: マスタシートの全項目をフォームに表示し、書き込み時にテンプレートシートに対応する列があるかチェックする
    // → これにより、マスタシートとテンプレートシートのフィールド名を完全一致させる必要がなくなる
    const masterLastCol = masterSheet.getLastColumn();
    const masterHeaders = masterSheet.getRange(1, 1, 1, masterLastCol).getValues()[0];
    const masterLastRow = masterSheet.getLastRow();

    if (masterLastCol === 0) {
      Logger.log(`⚠️ マスタシート「${masterSheetName}」にデータがありません`);
      return {
        success: false,
        error: 'マスタシートにデータがありません'
      };
    }

    const columns = [];
    const masterFieldNames = new Set(); // 重複チェック用

    // 【STEP 1】マスタシートの全列をフォームに表示（優先）
    // → マスタシートの1行目（ヘッダー）を走査し、各列の選択肢を取得
    masterHeaders.forEach((header, colIndex) => {
      // 【なぜ】空のヘッダーをスキップ
      // → 空列は意味を持たないため
      if (!header || header.toString().trim() === '') {
        return;
      }

      const fieldName = header.toString().trim();

      // 【なぜ】2行目以降の選択肢を取得
      // → マスタシートの2行目以降がプルダウン候補値のため
      let options = [];
      if (masterLastRow >= 2) {
        const colValues = masterSheet.getRange(2, colIndex + 1, masterLastRow - 1, 1).getValues();
        options = colValues
          .map(row => row[0])
          .filter(val => val !== null && val !== undefined && val.toString().trim() !== '')
          .map(val => val.toString().trim());
      }

      columns.push({
        header: fieldName, // マスタシートのヘッダー名
        options: options,
        hasOptions: options.length > 0,
        source: 'master' // 【追加】データソース識別
      });

      masterFieldNames.add(normalizeFieldName(fieldName)); // 正規化した名前で重複チェック
      Logger.log(`  📋 [マスタ] 「${fieldName}」: 選択肢${options.length}個`);
    });

    // 【STEP 2】ペアシート（テンプレートシート）の「ディスカバリー運用記入」列も取得
    // → マスタシートにない項目を追加（重複排除）
    try {
      const ranges = detectColumnRanges(templateSheet);
      if (ranges && ranges.discoveryRange && ranges.mainHeaderRow) {
        const subHeaderRow1 = ranges.mainHeaderRow + 1;
        const subHeaderRow2 = ranges.mainHeaderRow + 2;
        const templateLastCol = templateSheet.getLastColumn();

        if (templateLastCol > 0) {
          const templateSubHeaders1 = templateSheet.getRange(subHeaderRow1, 1, 1, templateLastCol).getValues()[0];
          const templateSubHeaders2 = templateSheet.getRange(subHeaderRow2, 1, 1, templateLastCol).getValues()[0];

          // 【なぜ】メインヘッダーを取得（各列が「ディスカバリー運用記入」かチェックするため）
          // → シート構造が複雑で、「代理店記入」列と「ディスカバリー運用記入」列が交互に配置されている場合があるため
          const mainHeaderRow = ranges.mainHeaderRow;
          const templateMainHeaders = templateSheet.getRange(mainHeaderRow, 1, 1, templateLastCol).getValues()[0];

          // 【なぜ】全列を走査して、メインヘッダーが「代理店記入」以外の列を処理
          // → 「ディスカバリー運用記入」「ディスカバリー営業記入」など、代理店記入以外の全ての列を含める
          // → 範囲ではなく、各列ごとに個別にチェックする

          for (let col = 0; col < templateLastCol; col++) {
            let mainHeader = templateMainHeaders[col];

            // 【なぜ】結合セル対応: メインヘッダーが空の場合、左側のセルから値を探す
            // → 行4のメインヘッダーも結合セルの可能性があるため
            if (!mainHeader || mainHeader.toString().trim() === '') {
              for (let leftCol = col - 1; leftCol >= Math.max(0, col - 10); leftCol--) {
                const leftMainHeader = templateMainHeaders[leftCol];
                if (leftMainHeader && leftMainHeader.toString().trim() !== '') {
                  mainHeader = leftMainHeader.toString().trim();
                  break;
                }
              }
            }

            // 【重要】メインヘッダーが空、または「代理店記入」を含む列はスキップ
            if (!mainHeader || mainHeader.toString().trim() === '' || mainHeader.toString().includes('代理店記入')) {
              continue; // 代理店記入列はスキップ
            }

            const header1 = templateSubHeaders1[col];
            const header2 = templateSubHeaders2[col];

            let part1 = header1 && header1.toString().trim() !== '' ? header1.toString().trim() : '';
            const part2 = header2 && header2.toString().trim() !== '' ? header2.toString().trim() : '';

            // 【なぜ】結合セル対応: 5行目が空欄の場合、左側のセルから値を探す
            // → スプレッドシートの結合セルは、getValues()で最初のセルにだけ値が入る
            // → 例: AC列「最適化と入札」がAC～AF列で結合されている場合、AD,AE,AF列は空欄になる
            if (!part1 && part2) {
              for (let leftCol = col - 1; leftCol >= Math.max(0, col - 10); leftCol--) {
                const leftHeader = templateSubHeaders1[leftCol];
                if (leftHeader && leftHeader.toString().trim() !== '') {
                  part1 = leftHeader.toString().trim();
                  break;
                }
              }
            }

            // 【なぜ】5行目と6行目を結合してフィールド名を作成
            // → フィールド名が2行にわたる場合があるため
            let fieldName = '';
            if (part1 && part2) {
              fieldName = part1 + '\n' + part2;
            } else if (part1) {
              fieldName = part1;
            } else if (part2) {
              fieldName = part2;
            }

            if (!fieldName || fieldName === '') continue;

            // 【なぜ】正規化して重複チェック
            // → マスタシートに同じ項目がある場合は、マスタを優先（プルダウンがあるため）
            const normalizedFieldName = normalizeFieldName(fieldName);
            if (masterFieldNames.has(normalizedFieldName)) {
              Logger.log(`  ⏩ [スキップ] 「${fieldName}」: マスタシートに存在するため`);
              continue; // マスタにある項目はスキップ
            }

            // 【なぜ】表示用のフィールド名を作成
            // → 改行を「 - 」に置き換えて見やすくする
            // → 例: 「オーディエンスセグメント\nオーディエンス」→「オーディエンスセグメント - オーディエンス」
            const displayName = fieldName.replace(/\n/g, ' - ');

            // 【なぜ】マスタシートに対応する列があるかチェック
            // → ある場合はプルダウン候補を取得、ない場合はテキスト入力
            let options = [];
            let hasOptions = false;

            // マスタシートの列を正規化して照合
            for (let masterColIndex = 0; masterColIndex < masterHeaders.length; masterColIndex++) {
              const masterHeader = masterHeaders[masterColIndex];
              if (masterHeader && normalizeFieldName(masterHeader) === normalizedFieldName) {
                // マスタシートに対応する列が見つかった
                if (masterLastRow >= 2) {
                  const colValues = masterSheet.getRange(2, masterColIndex + 1, masterLastRow - 1, 1).getValues();
                  options = colValues
                    .map(row => row[0])
                    .filter(val => val !== null && val !== undefined && val.toString().trim() !== '')
                    .map(val => val.toString().trim());
                  hasOptions = options.length > 0;
                }
                break;
              }
            }

            columns.push({
              header: displayName, // 表示用の名前（「 - 」区切り）
              originalHeader: fieldName, // 元のフィールド名（\n区切り、書き込み時に使用）
              options: options,
              hasOptions: hasOptions,
              source: 'template' // 【追加】データソース識別
            });

            masterFieldNames.add(normalizedFieldName); // 次回の重複チェック用
            Logger.log(`  📋 [ペアシート] 「${displayName}」: 選択肢${options.length}個`);
          }
        }
      } else {
        Logger.log(`  ⚠️ ペアシートの列範囲が検出できませんでした`);
      }
    } catch (templateError) {
      Logger.log(`  ⚠️ ペアシートのヘッダー取得エラー: ${templateError.message}`);
      // エラーがあってもマスタシートの項目は表示する
    }

    Logger.log(`✅ getMasterSheetData完了: ${columns.length}列取得（マスタ+ペアシート）`);

    return {
      success: true,
      mediaId: mediaId,
      masterSheetName: masterSheetName,
      columns: columns
    };

  } catch (e) {
    Logger.log(`❌ getMasterSheetData error: ${e.message}`);
    return {
      success: false,
      error: e.message
    };
  }
}

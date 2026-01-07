/**
 * マスタシートに値を追加する
 *
 * 【なぜこの関数が必要か】
 * - プルダウンの選択肢に存在しない値を書き込む場合、データ検証エラーになる
 * - 値をマスタシートに追加することで、自動的にプルダウン選択肢に反映される
 * - これによりデータ検証を保持したまま、新しい値を書き込める
 *
 * @param {string} masterSheetName - マスタシート名（例: "【マスタ】Googleリスティング広告"）
 * @param {string} fieldName - フィールド名（例: "キャンペーン目標の選択"）
 * @param {string} value - 追加する値
 * @returns {boolean} - 追加成功したらtrue
 */
function addValueToMasterSheet(masterSheetName, fieldName, value) {
  try {
    const ss = SpreadsheetApp.getActive();
    const masterSheet = ss.getSheetByName(masterSheetName);

    if (!masterSheet) {
      Logger.log(`  ⚠️ マスタシート「${masterSheetName}」が見つかりません`);
      return false;
    }

    // 【なぜ】1行目のヘッダーを取得
    // → フィールド名に対応する列を探すため
    const lastCol = masterSheet.getLastColumn();
    const headers = masterSheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // 【なぜ】フィールド名に一致する列を探す
    // → 正規化した値で照合することで、多少の表記揺れを許容する
    let targetCol = -1;
    const normalizedFieldName = normalizeFieldName(fieldName);
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] && normalizeFieldName(headers[i]) === normalizedFieldName) {
        targetCol = i + 1;
        break;
      }
    }

    if (targetCol === -1) {
      Logger.log(`  ⚠️ マスタシートに「${fieldName}」列が見つかりません`);
      return false;
    }

    // 【なぜ】その列の最終データ行を見つける
    // → シート全体の最終行ではなく、その列のデータが連続している最後の行を使う
    // → これにより、各列で連続してデータが追加される（バラバラにならない）
    const sheetLastRow = masterSheet.getLastRow();
    let columnLastRow = 1; // ヘッダー行から開始

    // 【なぜ】その列の2行目から順に走査して、最後のデータ行を見つける
    // → 空白行を挟まずに連続してデータが入っている最後の行を特定
    const columnValues = sheetLastRow >= 2
      ? masterSheet.getRange(2, targetCol, sheetLastRow - 1, 1).getValues()
      : [];

    for (let i = 0; i < columnValues.length; i++) {
      const cellValue = columnValues[i][0];
      if (cellValue !== null && cellValue !== undefined && cellValue.toString().trim() !== '') {
        columnLastRow = i + 2; // 2行目が i=0 なので、+2
      }
    }

    // 【なぜ】既存の値を取得（重複チェック用）
    // → 同じ値が既に存在する場合は追加しない
    const existingValues = columnLastRow >= 2
      ? masterSheet.getRange(2, targetCol, columnLastRow - 1, 1).getValues().map(row => row[0])
      : [];

    // 【なぜ】既に存在する場合はスキップ
    if (existingValues.some(v => v && v.toString().trim() === value.trim())) {
      Logger.log(`  ℹ️ 値「${value}」は既にマスタシートに存在します`);
      return true;
    }

    // 【なぜ】その列の最終データ行の次に値を追加
    // → 各列で連続してデータが入るようにする
    const newRow = columnLastRow + 1;
    masterSheet.getRange(newRow, targetCol).setValue(value);
    Logger.log(`  ✅ マスタシートに値を追加: 「${fieldName}」= "${value}" (行${newRow}列${targetCol})`);

    return true;

  } catch (e) {
    Logger.log(`  ❌ マスタシートへの値追加エラー: ${e.message}`);
    return false;
  }
}

/**
 * 構造化出力からスプレッドシートに書き込む（新版 - 2列対応）
 *
 * 【なぜ修正したか】
 * - シート構造が「代理店記入」列と「ディスカバリー運用記入」列の2つに分かれている
 * - AIの抽出結果 → 「代理店記入」列に書き込み
 * - フォームで選択した値 → 「ディスカバリー運用記入」列に書き込み
 *
 * @param {string} mediaId - 媒体ID
 * @param {Object} aiFields - AIが抽出したフィールド情報（代理店記入列用）
 * @param {Object} formFields - フォームで入力されたフィールド情報（ディスカバリー運用記入列用）
 * @param {Object} objSS - シート情報
 */
function insertDataFromStructured(mediaId, aiFields, formFields, objSS) {

  // ====== スプレッドシート書き込み ======
  // 書き込むシートを開く
  const targetSheetName = objSS.adsheet;
  // 【追加】マスタシート名を構築
  const masterSheetName = `【マスタ】${targetSheetName}`;

  Logger.log(`  🔍 ${mediaId}: シート「${targetSheetName}」に書き込み準備 (promap: ${objSS.adpromap})`);

  // 【修正】ハードコードIDを削除し、アクティブなスプレッドシートを使用
  const ss = SpreadsheetApp.getActive();
  const targetSheet = ss.getSheetByName(targetSheetName);

  if (!targetSheet) {
    throw new Error(`シート「${targetSheetName}」が見つかりません`);
  }

  const writeRow = targetSheet.getLastRow() + 1;
  const keyName = "mapping" + objSS.adpromap;
  const mappingArray = mappingObj[keyName];

  if (!mappingArray) {
    throw new Error(`mappingObj["${keyName}"]が見つかりません`);
  }

  // 【なぜ】動的マッピングを構築
  // → 同じフィールド名について、「代理店記入」列と「ディスカバリー運用記入」列を検出するため
  const dynamicMapping = buildDynamicMapping(targetSheet);

  // 【なぜこのログが必要か】
  // - buildDynamicMapping() が正しくフィールド名と列番号のマッピングを構築できているかを確認するため
  // - 問題: シート構造が想定と異なる場合、列範囲検出が失敗する可能性がある
  // - 正常な場合: dynamicMapping は { "キャンペーン目標": { agencyCol: 3, discoveryCol: 10 }, ... } のような形
  // - もし dynamicMapping が空オブジェクト {} の場合、4行目または6行目のヘッダー検出が失敗している
  // - もし discoveryCol が null の場合、そのフィールドが「ディスカバリー運用記入」列範囲に存在しない
  Logger.log(`  📋 動的マッピング: ${JSON.stringify(dynamicMapping)}`);

  // 【なぜこのログが必要か】
  // - processSubmission() から正しく formFields が渡されているかを確認するため
  // - 問題: processSubmission() でのデストラクチャリングが失敗していた場合、ここで空オブジェクトになる
  // - 正常な場合: formFields は { "キャンペーン目標": "売上", "入札戦略": "目標コンバージョン単価" } のような形
  // - もし formFields が空オブジェクト {} の場合、processSubmission() の抽出処理に問題がある
  // - もし formFields が undefined の場合、insertDataFromStructured() の引数渡しに問題がある
  Logger.log(`  📝 フォームフィールド: ${JSON.stringify(formFields)}`);

  // 【なぜ】2つの書き込み処理
  // 1. 代理店記入列: AIの抽出結果を書き込み（既存のマッピングを使用）
  // 2. ディスカバリー運用記入列: フォームの選択値を書き込み（動的マッピングを使用）

  let agencyWrittenCount = 0;
  let discoveryWrittenCount = 0;

  // ====== 1. 代理店記入列に書き込み（AIの抽出結果） ======
  mappingArray.forEach(({ key, col }) => {
    let value = aiFields[key];

    // 【なぜ】「ステータス」フィールドは常に「新規」を設定
    // → 新規行作成時のデフォルト値として「新規」を入れる仕様
    if (key === "ステータス") {
      value = "新規";
    } else {
      // nullまたはundefinedの場合は空文字
      if (value === null || value === undefined) {
        value = "";
      }
    }

    // 【新アプローチ】マスタシートに値を追加してから書き込み
    try {
      // 【なぜ】値をマスタシートに追加（フィールド名はkeyをそのまま使用）
      // → AIが抽出した値がプルダウンに無い場合、自動的に追加される
      if (value && value !== "" && key !== "ステータス") {
        addValueToMasterSheet(masterSheetName, key, value);
      }

      // 【なぜ】データ検証を保持して書き込み
      const range = targetSheet.getRange(writeRow, col);
      range.setValue(value);
      agencyWrittenCount++;

    } catch (e) {
      // 【なぜ】エラー時はデータ検証を「警告のみ」モードに変更して再試行
      // → setAllowInvalid(true)でプルダウンを残しつつ、無効な値も許容する
      Logger.log(`    ⚠️ 代理店記入列「${key}」の書き込みエラー (列${col}): ${e.message}`);
      try {
        const range = targetSheet.getRange(writeRow, col);
        const existingValidation = range.getDataValidation();
        if (existingValidation) {
          // データ検証を「警告のみ」モードに変更（プルダウンは残る）
          const newRule = existingValidation.copy().setAllowInvalid(true).build();
          range.setDataValidation(newRule);
        }
        range.setValue(value);          // 再書き込み
        agencyWrittenCount++;
        Logger.log(`    ✅ 再試行成功（データ検証を警告モードに変更）`);
      } catch (retryError) {
        Logger.log(`    ❌ 再試行も失敗: ${retryError.message}`);
      }
    }
  });

  // ====== 2. ディスカバリー運用記入列に書き込み（フォームの選択値） ======
  // 【なぜ】formFieldsが存在する場合のみ書き込み
  // → フォームで何も入力されていない場合はスキップ
  const skippedFields = []; // スキップされた項目を記録

  if (formFields && Object.keys(formFields).length > 0) {
    // 【なぜこのログが必要か】
    // - フォームフィールドの書き込み処理が開始されることを確認するため
    // - 問題: このログが表示されない場合、formFields が空または undefined
    // - このログが表示されれば、少なくとも formFields には何かしらのデータが含まれている
    Logger.log(`  🔍 フォームフィールド書き込み開始: ${Object.keys(formFields).length}フィールド`);

    Object.keys(formFields).forEach(fieldName => {
      const value = formFields[fieldName];

      // 【なぜ】正規化した値でマッピングを検索
      // → フィールド名の表記揺れ（全角/半角、スペース、改行など）を許容するため
      // → 例: マスタ「ターゲット リスト」とテンプレート「ターゲットリスト」を同一視
      let mapping = null;
      const normalizedFieldName = normalizeFieldName(fieldName);

      // 【なぜ】dynamicMappingの全キーを走査して正規化した値で照合
      // → オブジェクトのキーは文字列の完全一致が必要なため、正規化した値で検索する
      for (const mappingKey in dynamicMapping) {
        if (normalizeFieldName(mappingKey) === normalizedFieldName) {
          mapping = dynamicMapping[mappingKey];
          break;
        }
      }

      if (!mapping || !mapping.discoveryCol) {
        // 【新規追加】スキップされた項目を記録
        // → なぜ: マスタシートにあるがテンプレートシートにない項目を明確にするため
        // → これにより、どの項目をテンプレートシートに追加すべきかが分かる
        skippedFields.push({ fieldName, value, reason: 'テンプレートシートに対応する列がありません' });
        Logger.log(`    ⚠️ スキップ: フィールド「${fieldName}」（テンプレートシートに対応する列がありません）`);
        return;
      }

      const discoveryCol = mapping.discoveryCol;

      // 【新アプローチ】マスタシートに値を追加してから書き込み
      // → 問題: データ検証（プルダウン）の選択肢と値が完全一致しない場合、書き込みが拒否される
      // → 解決1: 値をマスタシートに追加 → プルダウン選択肢に自動反映
      // → 解決2: データ検証を保持したまま書き込み → プルダウンが残る
      try {
        const range = targetSheet.getRange(writeRow, discoveryCol);

        // 【なぜ】値をマスタシートに追加
        // → プルダウンの選択肢に含まれていない場合、自動的に追加される
        // → これにより次回から同じ値を選択できるようになる
        if (value && value !== "") {
          addValueToMasterSheet(masterSheetName, fieldName, value);
        }

        // 【なぜ】データ検証をそのまま保持して書き込み
        // → マスタシートに追加したため、データ検証は通過するはず
        // → プルダウンが残り、ユーザーは引き続き選択可能
        range.setValue(value || "");

        // 【重要】書き込み後に実際の値を読み取って検証
        // → setValue()が成功しても、実際にシートに値が入っていない場合があるため
        const actualValue = range.getValue();

        if (actualValue === value || (actualValue === "" && value === "")) {
          discoveryWrittenCount++;
        } else {
          Logger.log(`    ⚠️ フィールド「${fieldName}」書き込み後の検証失敗: 期待値="${value}", 実際値="${actualValue}"`);
        }

      } catch (e) {
        // 【なぜ】エラーが発生した場合はログに記録して、データ検証を「警告のみ」モードに変更して再試行
        // → setAllowInvalid(true)でプルダウンを残しつつ、無効な値も許容する
        Logger.log(`    ⚠️ フィールド「${fieldName}」の書き込みエラー: ${e.message}`);
        Logger.log(`    🔄 データ検証を警告モードに変更して再試行します`);

        try {
          const range = targetSheet.getRange(writeRow, discoveryCol);
          const existingValidation = range.getDataValidation();
          if (existingValidation) {
            // データ検証を「警告のみ」モードに変更（プルダウンは残る）
            const newRule = existingValidation.copy().setAllowInvalid(true).build();
            range.setDataValidation(newRule);
          }
          range.setValue(value || "");    // 再書き込み
          Logger.log(`    ✅ 再試行成功: フィールド「${fieldName}」を行${writeRow}列${discoveryCol}に書き込み（データ検証を警告モードに変更）`);
          discoveryWrittenCount++;
        } catch (retryError) {
          Logger.log(`    ❌ 再試行も失敗: ${retryError.message}`);
        }
      }
    });
  } else {
    // 【なぜこのログが必要か】
    // - formFields が空または undefined の場合に、その理由を明確にするため
    // - このログが表示される場合の原因:
    //   1. processSubmission() で formFields の抽出に失敗している
    //   2. ブラウザ側で動的フィールドの値が保存されていない
    //   3. payload.mediaList に該当する媒体が含まれていない
    Logger.log(`  ⚠️ フォームフィールドが空です`);
  }

  // 【新規追加】スキップされた項目のサマリーを出力
  if (skippedFields.length > 0) {
    Logger.log(`  ⚠️ ${mediaId}: ${skippedFields.length}個のフィールドがスキップされました（テンプレートシートに列がありません）:`);
    skippedFields.forEach(({ fieldName, value }) => {
      Logger.log(`    - "${fieldName}" = "${value}"`);
    });
    Logger.log(`  💡 これらの項目をテンプレートシートの「ディスカバリー運用記入」列範囲に追加すると、書き込まれるようになります`);
  }

  Logger.log(`  📝 ${mediaId}: 代理店記入=${agencyWrittenCount}フィールド, ディスカバリー運用記入=${discoveryWrittenCount}フィールドを行${writeRow}に書き込み`);
}


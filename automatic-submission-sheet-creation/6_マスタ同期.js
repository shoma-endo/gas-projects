/**
 * マスタシートとテンプレートシートのプルダウン同期機能
 *
 * 【背景】
 * スプレッドシート内に「【マスタ】◯◯」と「◯◯」というペアのシートが存在する。
 * マスタシート側でプルダウン候補を一元管理し、テンプレートシート側に自動反映することで、
 * データの整合性を保ち、手作業でのプルダウン設定を不要にする。
 *
 * 【仕様】
 * - マスタシート: 1行目がヘッダー、2行目以降が候補値
 * - テンプレートシート: 6行目がヘッダー、7行目以降にデータ検証（プルダウン）を設定
 * - ヘッダー名が一致する列について、マスタの値を参照する動的プルダウンを自動生成
 * - マスタに項目を追加・削除すれば、プルダウン候補も自動反映される
 */

/**
 * 単一のマスタシートとテンプレートシートを同期
 *
 * 【なぜこの関数が必要か】
 * - 2つのシート間でヘッダーを照合し、一致する列にのみデータ検証を設定する必要がある
 * - マスタとテンプレートで列の順序が異なる可能性があるため、ヘッダー名で照合する
 *
 * 【重要な仕様変更】
 * - テンプレートシートには「代理店記入」列と「ディスカバリー運用記入」列がある
 * - プルダウンは「ディスカバリー運用記入」列のみに設定する
 * - 「代理店記入」列はAIの抽出結果が入るため、プルダウン不要
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} masterSheet - マスタシート
 * @param {GoogleAppsScript.Spreadsheet.Sheet} templateSheet - テンプレートシート
 */
function syncMasterToTemplate(masterSheet, templateSheet) {
  const masterSheetName = masterSheet.getName();
  const templateSheetName = templateSheet.getName();

  // 【なぜ】マスタシートの1行目を取得
  // → 仕様: マスタシートの1行目がヘッダー行のため
  const masterLastCol = masterSheet.getLastColumn();
  if (masterLastCol === 0) {
    Logger.log(`  ⚠️ ${masterSheetName}: データがありません`);
    return;
  }

  const masterHeaders = masterSheet.getRange(1, 1, 1, masterLastCol).getValues()[0];

  // 【なぜ】「ディスカバリー運用記入」の列範囲を検出（メインヘッダー行も自動検出）
  // → この範囲の列のみにプルダウンを設定するため（代理店記入列は除外）
  // → シートによってメインヘッダー行が異なるため（Meta広告は行4、X広告は行3）
  const ranges = detectColumnRanges(templateSheet);
  if (!ranges || !ranges.discoveryRange) {
    Logger.log(`  ⚠️ ${templateSheetName}: ディスカバリー運用記入列が見つかりません`);
    return;
  }

  // 【なぜ】メインヘッダー行が見つからない場合はエラー
  if (!ranges.mainHeaderRow) {
    Logger.log(`  ⚠️ ${templateSheetName}: メインヘッダー行が検出できませんでした`);
    return;
  }

  Logger.log(`  📊 メインヘッダー行: ${ranges.mainHeaderRow}, ディスカバリー運用記入範囲: ${ranges.discoveryRange.start}～${ranges.discoveryRange.end}列`);

  // 【なぜ】サブヘッダー行を計算
  // → メインヘッダー行の次の行から2行分を読んで結合する
  // → Meta広告（mainHeaderRow=4）の場合: 5行目と6行目
  // → X広告（mainHeaderRow=3）の場合: 4行目と5行目
  const templateLastCol = templateSheet.getLastColumn();
  if (templateLastCol === 0) {
    Logger.log(`  ⚠️ ${templateSheetName}: データがありません`);
    return;
  }

  const subHeaderRow1 = ranges.mainHeaderRow + 1;
  const subHeaderRow2 = ranges.mainHeaderRow + 2;
  const templateSubHeaders1 = templateSheet.getRange(subHeaderRow1, 1, 1, templateLastCol).getValues()[0];
  const templateSubHeaders2 = templateSheet.getRange(subHeaderRow2, 1, 1, templateLastCol).getValues()[0];

  // 【なぜ】5行目と6行目を結合してサブヘッダーを作成
  // → フィールド名が2行にわたる場合があるため（例: "予算タイプ\n日予算・通算予算"）
  const templateSubHeaders = templateSubHeaders1.map((header1, colIndex) => {
    const header2 = templateSubHeaders2[colIndex];
    const part1 = header1 && header1.toString().trim() !== '' ? header1.toString().trim() : '';
    const part2 = header2 && header2.toString().trim() !== '' ? header2.toString().trim() : '';

    if (part1 && part2) {
      return part1 + '\n' + part2;
    } else if (part1) {
      return part1;
    } else if (part2) {
      return part2;
    }
    return '';
  });

  Logger.log(`  📋 ${masterSheetName}: ${masterHeaders.length}列, ${templateSheetName}: ${templateSubHeaders.length}列（サブヘッダー行${subHeaderRow1},${subHeaderRow2}）`);

  // 【なぜ】ヘッダー名が一致する列をマッピング（ディスカバリー運用記入範囲のみ）
  // → マスタとテンプレートで列の順序が異なる可能性があるため
  // → 「ディスカバリー運用記入」列のみにプルダウンを設定するため
  // → 正規化した値で照合することで、多少の表記揺れを許容する
  const columnMatches = [];

  masterHeaders.forEach((masterHeader, masterColIndex) => {
    // 【なぜ】空のヘッダーをスキップ
    // → 空列は意味を持たないため、処理対象外にする
    if (!masterHeader || masterHeader.toString().trim() === '') return;

    const normalizedMasterHeader = normalizeFieldName(masterHeader);

    // 【なぜ】テンプレート側で同じヘッダー名を持つ列を探す（ディスカバリー運用記入範囲内のみ）
    // → 「代理店記入」列は除外し、「ディスカバリー運用記入」列のみを対象にする
    // → 正規化した値で照合することで、多少の表記揺れを許容する
    for (let templateColIndex = ranges.discoveryRange.start - 1; templateColIndex < ranges.discoveryRange.end; templateColIndex++) {
      const templateHeader = templateSubHeaders[templateColIndex];
      if (templateHeader && normalizeFieldName(templateHeader) === normalizedMasterHeader) {
        columnMatches.push({
          header: masterHeader,
          masterCol: masterColIndex + 1, // 1ベース（GASの列番号は1から始まる）
          templateCol: templateColIndex + 1 // 1ベース
        });
        break; // 最初に見つかった列のみを使用
      }
    }
  });

  if (columnMatches.length === 0) {
    Logger.log(`  ⚠️ ${masterSheetName} → ${templateSheetName}: 一致するヘッダーがありません（ディスカバリー運用記入範囲）`);
    return;
  }

  Logger.log(`  🔗 一致する列（ディスカバリー運用記入範囲）: ${columnMatches.length}件`);

  // 【なぜ】各列についてデータ検証を設定
  // → ユーザーがテンプレートシートに入力する際、マスタの候補から選択できるようにするため
  let validationCount = 0;

  columnMatches.forEach(({ header, masterCol, templateCol }) => {
    try {
      // 【なぜ】マスタシートの2行目以降の値を確認
      // → 仕様: 2行目以降がプルダウン候補値のため
      const masterLastRow = masterSheet.getLastRow();

      if (masterLastRow < 2) {
        // 【なぜ】候補値がない場合はデータ検証を削除
        // → 仕様: マスタに候補がない列は自由入力を許可するため
        Logger.log(`    ⚠️ 「${header}」列: マスタに候補値がないため、データ検証を削除`);

        // テンプレートのデータ開始行以降の全行に対してデータ検証を削除
        const templateMaxRows = templateSheet.getMaxRows();
        if (templateMaxRows >= CONFIG.TEMPLATE_DATA_START_ROW) {
          const targetRange = templateSheet.getRange(CONFIG.TEMPLATE_DATA_START_ROW, templateCol, templateMaxRows - (CONFIG.TEMPLATE_DATA_START_ROW - 1), 1);
          targetRange.clearDataValidations();
        }
        return;
      }

      // 【なぜ】動的範囲参照を使用
      // → マスタシートに行を追加・削除した際、自動的にプルダウン候補が更新されるようにするため
      // → 静的範囲（例: A2:A10）だと、11行目以降を追加しても反映されない
      // 例: '【マスタ】Googleリスティング広告'!A2:A（A列の2行目以降すべて）
      const masterRangeA1 = `'${masterSheetName}'!${getColumnLetter(masterCol)}2:${getColumnLetter(masterCol)}`;

      // 【なぜ】データ検証ルールを作成
      // → テンプレートシートでユーザーがプルダウンから値を選択できるようにするため
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(masterSheet.getRange(`${getColumnLetter(masterCol)}2:${getColumnLetter(masterCol)}`), true)
        .setAllowInvalid(false) // 【なぜ】無効な値を拒否 → データの整合性を保つため
        .setHelpText(`マスタシート「${masterSheetName}」の「${header}」列から選択してください`) // 【なぜ】ヘルプテキスト → ユーザーに説明を表示
        .build();

      // 【なぜ】テンプレートシートのデータ開始行以降の全行に適用
      // → 仕様: テンプレートシートのヘッダーの次の行以降がデータ行のため
      const templateMaxRows = templateSheet.getMaxRows();
      if (templateMaxRows >= CONFIG.TEMPLATE_DATA_START_ROW) {
        const targetRange = templateSheet.getRange(CONFIG.TEMPLATE_DATA_START_ROW, templateCol, templateMaxRows - (CONFIG.TEMPLATE_DATA_START_ROW - 1), 1);
        targetRange.setDataValidation(rule);
        validationCount++;
        Logger.log(`    ✅ 「${header}」列: データ検証を設定 (${masterRangeA1} → ${CONFIG.TEMPLATE_DATA_START_ROW}行目以降)`);
      }

    } catch (e) {
      Logger.log(`    ❌ 「${header}」列: エラー - ${e.message}`);
    }
  });

  Logger.log(`  📝 ${templateSheetName}: ${validationCount}/${columnMatches.length}列にデータ検証を設定`);
}

/**
 * 列番号（1ベース）をA1形式の列文字に変換
 *
 * 【なぜこの関数が必要か】
 * - GASのデータ検証で範囲を指定する際、A1形式（例: A2:A）が必要
 * - 列番号（1, 2, 3...）を列文字（A, B, C...）に変換するため
 *
 * 【アルゴリズム】
 * - 26進数的な変換（ただしA=1から始まる）
 * - 例: 1→A, 26→Z, 27→AA, 52→AZ, 53→BA
 *
 * @param {number} column - 列番号（1ベース）
 * @return {string} - A1形式の列文字（例: 1→A, 27→AA）
 */
function getColumnLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter; // 65 = 'A'のASCIIコード
    column = (column - temp - 1) / 26;
  }
  return letter;
}


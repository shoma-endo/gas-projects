/**
 * シート列検出ユーティリティ
 *
 * 【背景】
 * スプレッドシートの構造：
 * - 4行目: メインヘッダー（「代理店記入」「ディスカバリー運用記入」）
 * - 6行目: サブヘッダー（実際のフィールド名、【マスタ】と一致）
 *
 * 【目的】
 * - 同じサブヘッダー名について、「代理店記入」列と「ディスカバリー運用記入」列を動的に検出
 * - これにより、マッピング定義を拡張せずに、動的に2つの列に書き込みできる
 */

/**
 * フィールド名を正規化して、多少の表記揺れを許容する
 *
 * 【なぜこの関数が必要か】
 * - マスタシートとテンプレートシートでフィールド名の表記が微妙に異なる場合がある
 * - 例: 全角/半角スペース、改行の有無、全角/半角英数字など
 *
 * 【正規化内容】
 * 1. 前後の空白を削除（trim）
 * 2. 全角スペース → 半角スペース
 * 3. 連続する空白 → 単一スペース
 * 4. エスケープされたダブルクォート → 通常のダブルクォート（\" → "）
 * 5. 改行文字を統一（\n に統一、複数改行は1つに）
 * 6. 全角英数字 → 半角英数字
 * 7. アルファベットを小文字に統一
 *
 * @param {string} fieldName - 正規化するフィールド名
 * @return {string} - 正規化されたフィールド名
 */
function normalizeFieldName(fieldName) {
  if (!fieldName || fieldName.toString().trim() === '') {
    return '';
  }

  let normalized = fieldName.toString();

  // 1. 前後の空白を削除
  normalized = normalized.trim();

  // 2. 全角スペース → 半角スペース
  normalized = normalized.replace(/　/g, ' ');

  // 3. 連続する空白 → 単一スペース
  normalized = normalized.replace(/\s+/g, ' ');

  // 4. エスケープされたダブルクォートを通常のダブルクォートに
  // 例: "入札戦略\n\"・バンパー広告..." → "入札戦略\n"・バンパー広告..."
  normalized = normalized.replace(/\\"/g, '"');

  // 5. 改行文字を統一（\r\n, \r → \n に統一、複数改行は1つに）
  normalized = normalized.replace(/\r\n/g, '\n');
  normalized = normalized.replace(/\r/g, '\n');
  normalized = normalized.replace(/\n+/g, '\n');

  // 6. 全角英数字 → 半角英数字
  // 全角英字（A-Z, a-z）と全角数字（0-9）を半角に変換
  normalized = normalized.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });

  // 7. アルファベットを小文字に統一
  normalized = normalized.toLowerCase();

  return normalized;
}

/**
 * シートから「代理店記入」と「ディスカバリー運用記入」の列範囲を検出
 *
 * 【なぜこの関数が必要か】
 * - 4行目のメインヘッダーを読み取って、どの列範囲が「代理店記入」で、
 *   どの列範囲が「ディスカバリー運用記入」かを特定する必要がある
 * - ハードコードせず、動的に検出することで、シート構造の変更に柔軟に対応
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @returns {Object} { agencyRange: {start, end}, discoveryRange: {start, end} }
 */
function detectColumnRanges(sheet) {
  try {
    // 【なぜ】行3～6を検索してメインヘッダーを自動検出
    // → シートによって構造が異なるため（Meta広告は行4、X広告は行3）
    // → 「代理店記入」「ディスカバリー運用記入」が含まれる行を探す
    const lastCol = sheet.getLastColumn();
    let mainHeaderRow = null;
    let mainHeaders = null;

    // 【なぜ】メインヘッダー検索範囲を順に検索
    // → どの行にメインヘッダーがあるか不明なため、見つかるまで探す
    for (let row = CONFIG.MAIN_HEADER_SEARCH_START; row <= CONFIG.MAIN_HEADER_SEARCH_END; row++) {
      const headers = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

      // 【なぜ】「代理店記入」または「ディスカバリー運用記入」が含まれているか確認
      // → どちらか一方でも見つかれば、その行がメインヘッダー行
      const hasAgency = headers.some(h => h && h.toString().includes('代理店記入'));
      const hasDiscovery = headers.some(h => h && h.toString().includes('ディスカバリー運用記入'));

      if (hasAgency || hasDiscovery) {
        mainHeaderRow = row;
        mainHeaders = headers;
        Logger.log(`📊 メインヘッダー行を検出: 行${mainHeaderRow}`);
        break;
      }
    }

    // 【なぜ】メインヘッダーが見つからない場合はエラー
    // → この場合、シート構造が想定外のため処理できない
    if (!mainHeaderRow || !mainHeaders) {
      Logger.log(`⚠️ メインヘッダー行（「代理店記入」「ディスカバリー運用記入」）が見つかりませんでした`);
      return {
        agencyRange: null,
        discoveryRange: null,
        mainHeaderRow: null
      };
    }

    let agencyStart = null;
    let agencyEnd = null;
    let discoveryStart = null;
    let discoveryEnd = null;

    // 【なぜ】結合セルを考慮して範囲を検出
    // → メインヘッダーは結合セルの可能性があるため
    for (let col = 0; col < mainHeaders.length; col++) {
      const value = mainHeaders[col];
      const colIndex = col + 1; // 1ベース

      if (value && value.toString().includes('代理店記入')) {
        if (agencyStart === null) agencyStart = colIndex;
        agencyEnd = colIndex;
      }

      if (value && value.toString().includes('ディスカバリー運用記入')) {
        if (discoveryStart === null) discoveryStart = colIndex;
        discoveryEnd = colIndex;
      }
    }

    // 【なぜ】結合セルの場合、最初の列だけに値が入っているため、終了列を推測
    // → 「代理店記入」の次が「ディスカバリー運用記入」の開始
    if (agencyStart !== null && discoveryStart !== null) {
      agencyEnd = discoveryStart - 1;
    } else if (agencyStart !== null && discoveryStart === null) {
      // ディスカバリー運用記入が見つからない場合、代理店記入のみ
      agencyEnd = lastCol;
    }

    Logger.log(`📊 列範囲検出: 代理店記入=${agencyStart}～${agencyEnd}, ディスカバリー運用記入=${discoveryStart}～${discoveryEnd}`);

    return {
      agencyRange: agencyStart !== null ? { start: agencyStart, end: agencyEnd } : null,
      discoveryRange: discoveryStart !== null ? { start: discoveryStart, end: discoveryEnd } : null,
      mainHeaderRow: mainHeaderRow // 【追加】メインヘッダー行番号を返す
    };

  } catch (e) {
    Logger.log(`❌ detectColumnRanges error: ${e.message}`);
    return {
      agencyRange: null,
      discoveryRange: null
    };
  }
}

/**
 * テンプレートシートの全フィールドについて、代理店記入列とディスカバリー運用記入列を取得
 *
 * 【なぜこの関数が必要か】
 * - 書き込み時に、各フィールドの2つの列番号が必要
 * - マッピング定義（0_マッピング.js）を拡張せず、動的に検出することで保守性向上
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @returns {Object} { fieldName: { agencyCol, discoveryCol }, ... }
 */
function buildDynamicMapping(sheet) {
  try {
    // 【なぜ】メインヘッダー行を検出してからサブヘッダー行を計算
    // → シートによってメインヘッダーの位置が異なるため（Meta広告は行4、X広告は行3）
    // → メインヘッダー行の次の行から2行分をサブヘッダーとして読む
    const ranges = detectColumnRanges(sheet);

    // 【なぜ】メインヘッダー行が見つからない場合はエラー
    // → mainHeaderRowがあれば、範囲オブジェクトがなくても全列を走査できる
    if (!ranges || !ranges.mainHeaderRow) {
      Logger.log(`⚠️ メインヘッダー行が検出できませんでした`);
      return {};
    }

    // 【なぜ】サブヘッダー行を計算
    // → メインヘッダー行の次の行から2行分を読む
    // → Meta広告（mainHeaderRow=4）の場合: 5行目と6行目
    // → X広告（mainHeaderRow=3）の場合: 4行目と5行目
    const subHeaderRow1 = ranges.mainHeaderRow + 1;
    const subHeaderRow2 = ranges.mainHeaderRow + 2;
    const lastCol = sheet.getLastColumn();
    const subHeaders1 = sheet.getRange(subHeaderRow1, 1, 1, lastCol).getValues()[0];
    const subHeaders2 = sheet.getRange(subHeaderRow2, 1, 1, lastCol).getValues()[0];

    Logger.log(`📊 サブヘッダー行: ${subHeaderRow1}行目と${subHeaderRow2}行目を使用`);

    // 【なぜ】メインヘッダーを取得（各列が「代理店記入」かチェックするため）
    // → getMasterSheetData()と同じロジックで、「代理店記入」列を個別に除外する
    const mainHeaderRow = ranges.mainHeaderRow;
    const mainHeaders = sheet.getRange(mainHeaderRow, 1, 1, lastCol).getValues()[0];

    const mapping = {};

    // 【なぜ】全列を走査して、メインヘッダーが「代理店記入」以外の列を処理
    // → getMasterSheetData()と同じロジックで列を選択することで、フィールド名の一致を保証
    // → 「ディスカバリー運用記入」「ディスカバリー営業記入」など、代理店記入以外の全ての列を含める
    for (let col = 0; col < lastCol; col++) {
      let mainHeader = mainHeaders[col];

      // 【なぜ】結合セル対応: メインヘッダーが空の場合、左側のセルから値を探す
      // → 行4のメインヘッダーも結合セルの可能性があるため
      // → 例: AC列「ディスカバリー運用記入」がAC～AF列で結合されている場合、AD,AE,AF列は空欄になる
      if (!mainHeader || mainHeader.toString().trim() === '') {
        for (let leftCol = col - 1; leftCol >= Math.max(0, col - 10); leftCol--) {
          const leftMainHeader = mainHeaders[leftCol];
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

      // 【なぜ】5行目と6行目を結合してフィールド名を作成
      // → 5行目だけ、6行目だけ、または両方ある場合に対応するため
      const header1 = subHeaders1[col];
      const header2 = subHeaders2[col];

      let part1 = header1 && header1.toString().trim() !== '' ? header1.toString().trim() : '';
      const part2 = header2 && header2.toString().trim() !== '' ? header2.toString().trim() : '';

      // 【なぜ】結合セル対応: 5行目が空欄の場合、左側のセルから値を探す
      // → スプレッドシートの結合セルは、getValues()で最初のセルにだけ値が入る
      // → 例: AC列「最適化と入札」がAC～AF列で結合されている場合、AD,AE,AF列は空欄になる
      // → この場合、AD列の part1 として AC列の値「最適化と入札」を使用する必要がある
      if (!part1 && part2) {
        // 【なぜ】左側の列を順に探して、最初に見つかった値を使用
        // → 結合セルの範囲は不明なため、最大10列左まで探す
        for (let leftCol = col - 1; leftCol >= Math.max(0, col - 10); leftCol--) {
          const leftHeader = subHeaders1[leftCol];
          if (leftHeader && leftHeader.toString().trim() !== '') {
            part1 = leftHeader.toString().trim();
            break; // 見つかったら終了
          }
        }
      }

      let fieldName = '';
      if (part1 && part2) {
        // 【なぜ】両方ある場合は改行で結合
        // → フォーム側も同じ形式（例: "予算タイプ\n日予算・通算予算"）で送信されるため
        fieldName = part1 + '\n' + part2;
      } else if (part1) {
        // 【なぜ】5行目だけの場合
        // → 6行目が空の場合（例: "購入タイプ"）
        fieldName = part1;
      } else if (part2) {
        // 【なぜ】6行目だけの場合
        // → 5行目が空の場合
        fieldName = part2;
      }

      if (!fieldName || fieldName === '') continue;

      const colIndex = col + 1; // 1ベース

      // 【なぜ】同じフィールド名の代理店記入列を探す
      // → AIの抽出結果を書き込む列を特定するため
      // → 代理店記入列も5行目と6行目を結合したフィールド名で探す
      // → 正規化した値で照合することで、多少の表記揺れを許容する
      let agencyColIndex = -1;
      const normalizedFieldName = normalizeFieldName(fieldName); // 【追加】照合用に正規化

      // 【なぜ】全列を走査して、メインヘッダーが「代理店記入」の列から対応する列を探す
      // → 範囲ではなく、各列ごとに個別にチェックする
      for (let agencyCol = 0; agencyCol < lastCol; agencyCol++) {
        const agencyMainHeader = mainHeaders[agencyCol];

        // 【重要】「代理店記入」列のみを対象
        if (!agencyMainHeader || !agencyMainHeader.toString().includes('代理店記入')) {
          continue; // 代理店記入列以外はスキップ
        }

        const agencyHeader1 = subHeaders1[agencyCol];
        const agencyHeader2 = subHeaders2[agencyCol];

        const agencyPart1 = agencyHeader1 && agencyHeader1.toString().trim() !== '' ? agencyHeader1.toString().trim() : '';
        const agencyPart2 = agencyHeader2 && agencyHeader2.toString().trim() !== '' ? agencyHeader2.toString().trim() : '';

        let agencyFieldName = '';
        if (agencyPart1 && agencyPart2) {
          agencyFieldName = agencyPart1 + '\n' + agencyPart2;
        } else if (agencyPart1) {
          agencyFieldName = agencyPart1;
        } else if (agencyPart2) {
          agencyFieldName = agencyPart2;
        }

        // 【変更】正規化した値で照合することで、多少の表記揺れを許容
        if (normalizeFieldName(agencyFieldName) === normalizedFieldName) {
          agencyColIndex = agencyCol;
          break;
        }
      }

      mapping[fieldName.toString().trim()] = {
        agencyCol: agencyColIndex !== -1 ? agencyColIndex + 1 : null,
        discoveryCol: colIndex
      };
    }

    Logger.log(`📋 動的マッピング構築完了: ${Object.keys(mapping).length}フィールド`);
    return mapping;

  } catch (e) {
    Logger.log(`❌ buildDynamicMapping error: ${e.message}`);
    return {};
  }
}

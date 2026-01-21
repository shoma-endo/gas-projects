/**
 * 8_入稿メール送信.js
 *
 * 掲載開始報告メール即時送信機能
 * - 対象：掲載開始報告をする = TRUE の行
 * - To：代理店担当者メール
 * - CC：nyukou@kakehashi.inc + 営業担当メールアドレス（あれば）
 * - Reply-To：nyukou@kakehashi.inc
 * - From：実行者のアカウント
 * - 送信後：掲載開始報告=済、入稿メールステータス=送信済
 */

/* =========================
 * 共通設定
 * ========================= */
const MAIL_COMMON = {
  TZ: Session.getScriptTimeZone(),
  FIXED_CC: 'nyukou@kakehashi.inc',
  REPLY_TO: 'nyukou@kakehashi.inc',
  FROM_NAME: '株式会社KAKEHASHI入稿チーム',
  // 送信者ドメインチェック（true: @kakehashi.incのみ許可、false: チェックなし）
  VALIDATE_SENDER_DOMAIN: false,
};

/* =========================
 * メール送信設定
 * ========================= */
const SEND_CFG = {
  CHECK_COL: '掲載開始報告をする',
  DONE_COL: '掲載開始報告',
  STATUS_COL: '入稿メールステータス',

  COLS: {
    TO: '代理店担当者メール',
    AGENCY: '代理店名',
    CASE: '案件タイトル',
    CLIENT: '広告主',
    MEDIA: '媒体名',
    MENU: 'メニュー名',
    START: '掲載開始日',
    SALES_EMAIL: '営業担当メールアドレス',
  }
};

/**
 * チェック行を即時送信（掲載開始報告メール）
 * ★修正：LockServiceで排他制御 + 送信直後に即時ステータス更新で重複送信防止
 */
function sendCheckedRows() {
  // ★排他制御：同時実行を防止
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    SpreadsheetApp.getUi().alert('別の送信処理が実行中です。しばらく待ってから再度お試しください。');
    return;
  }

  try {
    const sheet = SpreadsheetApp.getActiveSheet();

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    // 2: ヘッダー行(1行目) + データ行(1行以上)で最低2行必要
    // 1: 最低1列は必要
    if (lastRow < 2 || lastCol < 1) {
      SpreadsheetApp.getUi().alert('データがありません。');
      return;
    }

    // ヘッダー行を動的に検索（CHECK_COLを含む行を探す）
    const headerRowInfo = findHeaderRow_(sheet, SEND_CFG.CHECK_COL);
    if (!headerRowInfo) {
      SpreadsheetApp.getUi().alert(`エラー: 「${SEND_CFG.CHECK_COL}」列が見つかりません。`);
      return;
    }

    const headerRow = headerRowInfo.row;
    const dataStartRow = headerRow + 1;

    Logger.log(`ヘッダー行検出: ${headerRow}行目, データ開始: ${dataStartRow}行目`);

    // ヘッダー行からデータを取得
    const headerValues = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
    const header = headerValues.map(h => String(h || '').trim());

    // データ行を取得
    const dataRowCount = lastRow - headerRow;
    if (dataRowCount < 1) {
      SpreadsheetApp.getUi().alert('データ行がありません。');
      return;
    }
    const values = sheet.getRange(dataStartRow, 1, dataRowCount, lastCol).getValues();

    // デバッグ(ヘッダー列とチェック列の実値確認)
    debugHeaderColumnInThisSheet_(sheet, SEND_CFG.CHECK_COL, dataStartRow, 10);

    const map = buildMailHeaderMap_(header);

    // 必須列のチェック
    const checkCol = requireMailCol_(map, SEND_CFG.CHECK_COL);
    const doneCol = requireMailCol_(map, SEND_CFG.DONE_COL);
    const statusCol = requireMailCol_(map, SEND_CFG.STATUS_COL);

    const toCol = requireMailCol_(map, SEND_CFG.COLS.TO);
    const salesCol = findMailCol_(map, SEND_CFG.COLS.SALES_EMAIL);

    getExecutorEmail_(); //権限/取得可否の早期チェック

    let sentCount = 0;
    let errorCount = 0;

    for (let r = 0; r < values.length; r++) {
      const row = values[r];
      const actualRow = dataStartRow + r; // 実際のシート行番号

      // 完全空行はスキップ
      if (isBlankRow_(row)) continue;

      const checkValue = row[checkCol];
      const doneValue = String(row[doneCol] || '').trim();
      const statusValue = String(row[statusCol] || '').trim();

      Logger.log(`行${actualRow}: 掲載開始報告をする=${stringify_(checkValue)}(型:${typeof checkValue}), 掲載開始報告=${doneValue}, ステータス=${statusValue}`);

      // 既に送信済みの場合はスキップ
      if (doneValue === '済' || statusValue === '送信済') {
        Logger.log(`  →スキップ:既に送信済み`);
        continue;
      }

      // チェックがついていない場合はスキップ
      if (!toMailBool_(checkValue)) {
        Logger.log(`  →スキップ:チェックなし`);
        continue;
      }

      const to = String(row[toCol] || '').trim();
      if (!to) {
        Logger.log(`  →エラー:宛先メールアドレスなし`);
        errorCount++;
        continue;
      }

      // CC設定
      const ccList = [MAIL_COMMON.FIXED_CC];
      if (salesCol !== null) {
        const salesEmail = String(row[salesCol] || '').trim();
        if (salesEmail) ccList.push(salesEmail);
      }

      const subject = buildMailSubject_(row, map);
      const body = buildMailBody_(row, map);

      try {
        GmailApp.sendEmail(to, subject, body, {
          name: MAIL_COMMON.FROM_NAME,
          cc: ccList.join(','),
          replyTo: MAIL_COMMON.REPLY_TO,
        });

        // ★修正：送信直後に即時ステータス更新（レースコンディション防止）
        sheet.getRange(actualRow, doneCol + 1).setValue('済');
        sheet.getRange(actualRow, statusCol + 1).setValue('送信済');

        sentCount++;
        Logger.log(`  →送信成功:${to}`);

      } catch (e) {
        errorCount++;
        Logger.log(`  →送信エラー:${e && e.message ? e.message : e}`);
      }
    }

    SpreadsheetApp.getUi().alert(
      '送信完了',
      `送信成功:${sentCount}件\nエラー:${errorCount}件`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    Logger.log(`メール送信完了:成功=${sentCount},エラー=${errorCount}`);

  } finally {
    // ★排他制御：必ずロック解放
    lock.releaseLock();
  }
}

/**
 * メール件名を生成
 */
function buildMailSubject_(row, map) {
  const caseName = String(row[map[SEND_CFG.COLS.CASE]] || '').trim();
  const client = String(row[map[SEND_CFG.COLS.CLIENT]] || '').trim();
  return `【掲載開始】${caseName}｜${client}`;
}

/**
 * メール本文を生成
 */
function buildMailBody_(row, map) {
  const agency = String(row[map[SEND_CFG.COLS.AGENCY]] || '').trim();
  const caseName = String(row[map[SEND_CFG.COLS.CASE]] || '').trim();
  const client = String(row[map[SEND_CFG.COLS.CLIENT]] || '').trim();
  const media = String(row[map[SEND_CFG.COLS.MEDIA]] || '').trim();
  const menu = String(row[map[SEND_CFG.COLS.MENU]] || '').trim();
  const start = formatMailDateCell_(row[map[SEND_CFG.COLS.START]]);

  const dear = agency ? `${agency} ご担当者様` : 'ご担当者様';

  return [
    `${dear}`,
    ``,
    `平素より大変お世話になっております。`,
    `${MAIL_COMMON.FROM_NAME}です。`,
    ``,
    `下記案件につきまして、掲載開始のご連絡です。`,
    ``,
    `案件名：${caseName}`,
    `広告主：${client}`,
    `媒体名：${media}`,
    `メニュー名：${menu}`,
    `掲載開始日：${start}`,
    ``,
    `何卒よろしくお願いいたします。`,
    ``,
    `―――――――――――――――――`,
    `${MAIL_COMMON.FROM_NAME}`,
  ].join('\n');
}

/* =========================
 * ユーティリティ関数
 * ========================= */

/**
 * ヘッダー行を動的に検索
 * @param {Sheet} sheet - 対象シート
 * @param {string} headerText - 検索するヘッダー名
 * @returns {Object|null} - { row: ヘッダー行番号, col: 列番号 } または null
 */
function findHeaderRow_(sheet, headerText) {
  const ranges = sheet.createTextFinder(headerText).matchEntireCell(true).findAll();
  if (!ranges || ranges.length === 0) {
    return null;
  }
  // 最初に見つかったセルの行をヘッダー行とする
  const firstMatch = ranges[0];
  return {
    row: firstMatch.getRow(),
    col: firstMatch.getColumn()
  };
}

/**
 * ヘッダー行からマップを構築
 */
function buildMailHeaderMap_(headerRow) {
  const m = {};
  headerRow.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) m[key] = i;
  });
  return m;
}

/**
 * 必須列を取得（なければエラー）
 */
function requireMailCol_(map, name) {
  if (!(name in map)) throw new Error(`必須ヘッダー「${name}」が見つかりません。`);
  return map[name];
}

/**
 * 列を検索（なければnull）
 */
function findMailCol_(map, name) {
  return (name in map) ? map[name] : null;
}

/**
 * 値をbooleanに変換(checkbox/文字列/数値を許容)
 */
function toMailBool_(v) {
  if (v === true) return true;
  if (v === 1) return true;
  const s = String(v || '').trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'yes' || s === 'y';
}

/**
 * 完全空行判定
 */
function isBlankRow_(row) {
  for (let i = 0; i < row.length; i++) {
    const v = row[i];
    if (v === null || v === undefined) continue;
    if (typeof v === 'string' && v.trim() === '') continue;
    //checkboxのbooleanは空扱いに寄せる(書式行でfalseが並ぶ事故回避)
    if (typeof v === 'boolean') continue;
    return false;
  }
  return true;
}

/**
 * 日付セルをフォーマット
 */
function formatMailDateCell_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, MAIL_COMMON.TZ, 'yyyy/MM/dd');
  }
  return String(value).trim();
}

/**
 * 実行者のメールアドレスを取得
 * MAIL_COMMON.VALIDATE_SENDER_DOMAIN が true の場合のみドメインチェックを行う
 */
function getExecutorEmail_() {
  const e = Session.getActiveUser().getEmail();
  if (!e) throw new Error('実行者メールを取得できません。Google Workspaceアカウントで実行してください。');

  if (MAIL_COMMON.VALIDATE_SENDER_DOMAIN && !/@kakehashi\.inc$/i.test(e)) {
    throw new Error(`送信元は @kakehashi.inc のみ許可。現在：${e}`);
  }
  return e;
}

/* =========================
 * デバッグ(ヘッダー列の実値確認)
 * ========================= */

function debugHeaderColumnInThisSheet_(sheet, headerText, dataStartRow, sampleRows) {
  Logger.log(`debug:sheet=${sheet.getName()}`);

  const ranges = sheet.createTextFinder(headerText).matchEntireCell(true).findAll();
  if (!ranges || ranges.length === 0) {
    Logger.log(`debug:header not found:"${headerText}"`);
    return;
  }

  Logger.log(`debug:found header count=${ranges.length}`);

  const lastRow = sheet.getLastRow();

  ranges.forEach((r, i) => {
    const row = r.getRow();
    const col = r.getColumn();
    const a1 = r.getA1Notation();
    const colLetter = columnToLetter_(col);

    Logger.log(`debug:hit[${i}]:cell=${a1},row=${row},col=${col}(${colLetter})`);

    const start = Math.max(dataStartRow, row + 1);
    const end = Math.min(lastRow, start + sampleRows - 1);
    const n = Math.max(0, end - start + 1);

    Logger.log(`debug:sample ${colLetter}${start}:${colLetter}${end}(n=${n})`);
    if (n <= 0) return;

    const vals = sheet.getRange(start, col, n, 1).getValues();
    const disp = sheet.getRange(start, col, n, 1).getDisplayValues();

    for (let k = 0; k < n; k++) {
      const rr = start + k;
      const v = vals[k][0];
      const d = disp[k][0];
      Logger.log(`debug:${colLetter}${rr}:value=${stringify_(v)}(type:${typeof v}),display="${d}"`);
    }
  });
}

/**
 * 列番号をアルファベット表記に変換する（例: 1→"A", 27→"AA"）
 * デバッグログでセル位置を分かりやすく表示するために使用
 * @param {number} col1 - 1始まりの列番号
 * @returns {string} 列のアルファベット表記
 */
function columnToLetter_(col1) {
  let n = col1;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function stringify_(v) {
  try {
    return JSON.stringify(v);
  } catch (e) {
    return String(v);
  }
}
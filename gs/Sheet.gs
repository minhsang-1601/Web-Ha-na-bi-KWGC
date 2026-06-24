// ─── スプレッドシート書き込み ──────────────────────────────────────────────────

const _t = v => String(v || '').trim(); // TRIM ヘルパー

/** 手作業シートのヘッダーをリセット（古いシートの列を削除） */
function resetTesagyouSheetHeaders() {
  const ss = getDataSpreadsheet();
  const sheet = ss.getSheetByName(DEFAULT_SHEET_NAME2);
  if (!sheet) return;

  // 最大列数を取得
  const maxCol = sheet.getLastColumn();
  const headerCount = TESAGYOU_HEADERS.length;

  // 古い列を削除
  if (maxCol > headerCount) {
    sheet.deleteColumns(headerCount + 1, maxCol - headerCount);
  }

  // ヘッダーを再設定
  sheet.getRange(1, 1, 1, headerCount)
    .setValues([TESAGYOU_HEADERS])
    .setFontWeight('bold').setBackground('#fce8b2');

  // サブヘッダー（2行目）を再設定
  const subheaders = [
    '直接入力',         // A: 受付番号
    'XLOOKUP\n自動', 'XLOOKUP\n自動', 'XLOOKUP\n自動',  // B-D
    'XLOOKUP\n自動', 'XLOOKUP\n自動',                    // E-F
    'XLOOKUP\n自動', 'XLOOKUP\n自動',                    // G-H
    'checkbox\n手動', 'タイムスタンプ\n自動', 'checkbox\n手動',  // I-K
    'タイムスタンプ\n自動',                               // L
  ];

  sheet.getRange(2, 1, 1, headerCount)
    .setValues([subheaders])
    .setFontSize(8).setFontColor('#888888').setBackground('#fffbf0').setWrap(true);

  sheet.setRowHeight(2, 36);
  applyTesagyouColumnWidths(sheet);
}

function appendRow(data, sheetName) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss    = getDataSpreadsheet();
    let sheet   = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#d0e4f7');
      sheet.setFrozenRows(1);
      applyColumnWidths(sheet);
      _ensureFilter(sheet, 1, HEADERS.length);
    }

    const now      = new Date();
    const receptNo = data.receipt_no ||
                     getReceiptNoPrefix() + Utilities.formatDate(now, 'Asia/Tokyo', 'MMddHHmmssSSS');

    const newRow  = sheet.getLastRow() + 1;
    const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    // Build row data
    const rowData = [
      receptNo,                          // A: 受付番号
      dateStr,                           // B: 受付日時
      _t(data.company_name),            // C
      _t(data.company_furigana),        // D
      _t(data.rep_position),            // E: 代表者役職
      _t(data.rep_position_furigana),   // F: 代表者役職（フリガナ）
      _t(data.rep_name),                // G: 代表者名
      _t(data.rep_furigana),            // H: 代表者名（フリガナ）
      _t(data.staff_name),              // I
      _t(data.staff_furigana),          // J
      _t(data.zipcode),                 // K
      _t(data.address),                 // L
      _t(data.phone),                   // M
      _t(data.email),                   // N
      _t(data.category),                // O
      _t(data.website_url),             // P
      _t(data.alt_name),                // Q: 会社名・団体名と異なる名
    ];

    sheet.getRange(newRow, 1, 1, HEADERS.length).setValues([rowData]);

    // Set text format for zipcode (I) and phone (K) to preserve leading zeros
    sheet.getRange(newRow, 11).setNumberFormat('@').setValue(_t(data.zipcode));
    sheet.getRange(newRow, 13).setNumberFormat('@').setValue(_t(data.phone));

    SpreadsheetApp.flush();
    return receptNo;
  } finally {
    lock.releaseLock();
  }
}

function appendToTesagyouSheet(receptNo, sheetName2, data) {
  const ss    = getDataSpreadsheet();
  let sheet   = ss.getSheetByName(sheetName2);
  if (!sheet) sheet = ss.insertSheet(sheetName2);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(TESAGYOU_HEADERS);
    sheet.getRange(1, 1, 1, TESAGYOU_HEADERS.length)
      .setFontWeight('bold').setBackground('#fce8b2');
    // 2行目: サブヘッダー（操作種別メモ）
    sheet.appendRow([
      '直接入力',         // A: 受付番号
      'XLOOKUP\n自動', 'XLOOKUP\n自動', 'XLOOKUP\n自動',  // B-D
      'XLOOKUP\n自動', 'XLOOKUP\n自動', 'XLOOKUP\n自動', 'XLOOKUP\n自動', // E-H
      'checkbox\n手動', 'タイムスタンプ\n自動', 'checkbox\n手動',           // I-K
      'タイムスタンプ\n自動',                             // L
    ]);
    sheet.getRange(2, 1, 1, TESAGYOU_HEADERS.length)
      .setFontSize(8).setFontColor('#888888').setBackground('#fffbf0').setWrap(true);
    sheet.setRowHeight(2, 36);
    sheet.setFrozenRows(2);
    applyTesagyouColumnWidths(sheet);
    _ensureFilter(sheet, 1, TESAGYOU_HEADERS.length);
  }

  const newRow = sheet.getLastRow() + 1;
  const kubun    = _t(data ? (data.category || '') : '').toUpperCase();
  const autoSend = AUTO_SEND_KUBUN.includes(kubun);

  // A: 受付番号（直接入力）
  sheet.getRange(newRow, COL_RECEPT_NO).setValue(receptNo);

  // XLOOKUP: 受付番号（$A）をキーに協賛申込み一覧から各列を参照
  Object.entries(TESAGYOU_LOOKUP_COLS).forEach(([col, srcCol]) => {
    const formula =
      `=IFERROR(XLOOKUP($A${newRow},'${DEFAULT_SHEET_NAME}'!$A:$A,'${DEFAULT_SHEET_NAME}'!$${srcCol}:$${srcCol}),"見つかりません")`;
    sheet.getRange(newRow, Number(col)).setFormula(formula);
  });

  // I: 受付完了 checkbox — B〜E は申込時に自動完了
  sheet.getRange(newRow, COL_UKETSUKE).insertCheckboxes();
  if (autoSend) sheet.getRange(newRow, COL_UKETSUKE).setValue(true);

  // J: 請求書送信日時 — B〜E は申込時に自動送信済みのため現在日時
  if (autoSend) sheet.getRange(newRow, COL_INV_DATE).setValue(nowStr());

  // K: 入金完了 checkbox
  sheet.getRange(newRow, COL_NYUKIN).insertCheckboxes();

  // L: お礼状送信日時 は空（トリガーが自動設定）
}

/**
 * 指定行にフィルターを設定する（既存フィルターがあれば再作成しない）
 * @param {Sheet} sheet
 * @param {number} headerRow - フィルター基準行（1始まり）
 * @param {number} numCols
 */
/**
 * シート全体のデータ行を上揃え・左揃えに設定する
 * @param {Sheet} sheet
 * @param {number} firstDataRow - データ開始行（ヘッダー除く）
 * @param {number} numCols
 */
function _applyAlignment(sheet, firstDataRow, numCols) {
  try {
    const lastRow = sheet.getMaxRows();
    if (lastRow < firstDataRow) return;
    sheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, numCols)
      .setVerticalAlignment('top')
      .setHorizontalAlignment('left');
  } catch (e) {
    console.warn('アライメント設定エラー:', e.message);
  }
}

function _ensureFilter(sheet, headerRow, numCols) {
  try {
    const existing = sheet.getFilter();
    if (!existing) {
      sheet.getRange(headerRow, 1, 1, numCols).createFilter();
    }
  } catch (e) {
    console.warn('フィルター設定エラー:', e.message);
  }
}

function applyColumnWidths(sheet) {
  // A=受付番号 B=受付日時 C=会社名 D=会社名ふりがな E=代表者役職 F=役職ふりがな
  // G=代表者名 H=代表者名ふりがな I=担当者名 J=担当者ふりがな
  // K=郵便番号 L=住所 M=電話番号 N=メール O=区分 P=URL Q=異なる名
  [150, 150, 200, 180, 150, 150, 130, 130, 120, 120, 90, 220, 120, 220, 60, 200, 200]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  _applyAlignment(sheet, 2, HEADERS.length);
}

function applyTesagyouColumnWidths(sheet) {
  // A=受付番号 B=区分 C=電話 D=会社名 E=住所 F=代表者名
  // G=メール H=URL I=受付完了 J=請求書日時 K=入金 L=礼状日時
  [150, 60, 120, 200, 200, 130, 200, 160, 70, 150, 70, 150]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  _applyAlignment(sheet, 3, TESAGYOU_HEADERS.length);
}

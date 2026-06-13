// ─── スプレッドシート書き込み ──────────────────────────────────────────────────

const _t = v => String(v || '').trim(); // TRIM ヘルパー

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
    sheet.getRange(newRow, 1, 1, HEADERS.length).setValues([[
      receptNo,                   // A: 受付番号
      dateStr,                    // B: 受付日時
      _t(data.company_name),     // C
      _t(data.company_furigana), // D
      _t(data.rep_name),         // E
      _t(data.rep_furigana),     // F
      _t(data.staff_name),       // G
      _t(data.staff_furigana),   // H
      _t(data.zipcode),          // I
      _t(data.address),          // J
      _t(data.phone),            // K
      _t(data.email),            // L
      _t(data.category),         // M
      _t(data.website_url),      // N
    ]]);

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
      'タイムスタンプ\n自動', '区分＋7桁\n自動',                             // L-M
      'checkbox\n手動', 'タイムスタンプ\n自動', 'タイムスタンプ\n自動',       // N-P
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

  // B: 受付番号（直接入力）
  sheet.getRange(newRow, COL_RECEPT_NO).setValue(receptNo);

  // XLOOKUP: 受付番号（$A）をキーに協賛申込み一覧から各列を参照
  Object.entries(TESAGYOU_LOOKUP_COLS).forEach(([col, srcCol]) => {
    const formula =
      `=IFERROR(XLOOKUP($A${newRow},'${DEFAULT_SHEET_NAME}'!$A:$A,'${DEFAULT_SHEET_NAME}'!$${srcCol}:$${srcCol}),"見つかりません")`;
    sheet.getRange(newRow, Number(col)).setFormula(formula);
  });

  // J: 受付完了 checkbox — B〜E は申込時に自動完了
  sheet.getRange(newRow, COL_UKETSUKE).insertCheckboxes();
  if (autoSend) sheet.getRange(newRow, COL_UKETSUKE).setValue(true);

  // K: 請求書送信日時 — B〜E は申込時に自動送信済みのため現在日時
  if (autoSend) sheet.getRange(newRow, COL_INV_DATE).setValue(nowStr());

  // L: 入金完了 checkbox
  sheet.getRange(newRow, COL_NYUKIN).insertCheckboxes();

  // O: 案内実施 checkbox
  sheet.getRange(newRow, COL_ANNAIBUN).insertCheckboxes();

  // M, N, P, Q は空（各トリガーが自動設定）
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
  // A=受付番号 B=受付日時 C=個人名 D=ふりがな E=代表者 F=ふりがな
  // G=担当者 H=ふりがな I=郵便番号 J=住所 K=電話番号 L=メール M=区分 N=URL
  [150, 150, 200, 180, 150, 150, 120, 120, 90, 220, 120, 220, 60, 200]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  _applyAlignment(sheet, 2, HEADERS.length);
}

function applyTesagyouColumnWidths(sheet) {
  // A=受付番号 B=区分 C=電話 D=個人名 E=住所 F=代表者 G=メール H=URL
  // I=受付完了 J=請求書日時 K=入金 L=座席日時 M=座席番号 N=案内 O=案内日時 P=礼状日時
  [150, 60, 120, 200, 200, 150, 200, 160, 70, 150, 70, 150, 120, 70, 150, 150]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  _applyAlignment(sheet, 3, TESAGYOU_HEADERS.length);
}

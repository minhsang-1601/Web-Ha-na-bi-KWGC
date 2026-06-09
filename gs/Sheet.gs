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
    }

    const now      = new Date();
    const receptNo = data.receipt_no ||
                     getReceiptNoPrefix() + Utilities.formatDate(now, 'Asia/Tokyo', 'MMddHHmmssSSS');
    const newRow   = sheet.getLastRow() + 1;

    const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    sheet.getRange(newRow, 1, 1, HEADERS.length).setValues([[
      _t(data.kanri_id),         // A: 管理ID番号
      receptNo,                   // B: 受付番号
      dateStr,                    // C: 受付日時
      _t(data.company_name),     // D
      _t(data.company_furigana), // E
      _t(data.rep_name),         // F
      _t(data.rep_furigana),     // G
      _t(data.staff_name),       // H
      _t(data.staff_furigana),   // I
      _t(data.zipcode),          // J
      _t(data.address),          // K
      _t(data.phone),            // L
      _t(data.email),            // M
      _t(data.category),         // N
      _t(data.website_url),      // O
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
      'XLOOKUP\n自動',   // A: 管理ID番号
      '直接入力',         // B: 受付番号
      'XLOOKUP\n自動', 'XLOOKUP\n自動', 'XLOOKUP\n自動',  // C-E
      'XLOOKUP\n自動', 'XLOOKUP\n自動', 'XLOOKUP\n自動', 'XLOOKUP\n自動', // F-I
      'checkbox\n手動', 'タイムスタンプ\n自動', 'checkbox\n手動',           // J-L
      'タイムスタンプ\n自動', '区分＋7桁\n自動',                             // M-N
      'checkbox\n手動', 'タイムスタンプ\n自動', 'タイムスタンプ\n自動',       // O-Q
    ]);
    sheet.getRange(2, 1, 1, TESAGYOU_HEADERS.length)
      .setFontSize(8).setFontColor('#888888').setBackground('#fffbf0').setWrap(true);
    sheet.setRowHeight(2, 36);
    sheet.setFrozenRows(2);
    applyTesagyouColumnWidths(sheet);
  }

  const newRow   = sheet.getLastRow() + 1;
  const kubun    = _t(data ? (data.category || '') : '').toUpperCase();
  const autoSend = AUTO_SEND_KUBUN.includes(kubun);

  // B: 受付番号（直接入力）
  sheet.getRange(newRow, COL_RECEPT_NO).setValue(receptNo);

  // XLOOKUP: 受付番号（$B）をキーに協賛申込み一覧から各列を参照
  Object.entries(TESAGYOU_LOOKUP_COLS).forEach(([col, srcCol]) => {
    const formula =
      `=IFERROR(XLOOKUP($B${newRow},'${DEFAULT_SHEET_NAME}'!$B:$B,'${DEFAULT_SHEET_NAME}'!$${srcCol}:$${srcCol}),"見つかりません")`;
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

function applyColumnWidths(sheet) {
  // A=管理ID B=受付番号 C=受付日時 D=個人名 E=ふりがな F=代表者 G=ふりがな
  // H=担当者 I=ふりがな J=郵便番号 K=住所 L=電話番号 M=メール N=区分 O=URL
  [160, 150, 150, 200, 180, 150, 150, 120, 120, 90, 220, 120, 220, 60, 200]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

function applyTesagyouColumnWidths(sheet) {
  // A=管理ID B=受付番号 C=区分 D=電話 E=個人名 F=住所 G=代表者 H=メール I=URL
  // J=受付完了 K=請求書日時 L=入金 M=座席日時 N=座席番号 O=案内 P=案内日時 Q=礼状日時
  [160, 150, 60, 120, 200, 200, 150, 200, 160, 70, 150, 70, 150, 120, 70, 150, 150]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

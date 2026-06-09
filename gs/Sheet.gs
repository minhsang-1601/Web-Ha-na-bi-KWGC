// ─── スプレッドシート書き込み ──────────────────────────────────────────────────

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
      applyColumnWidths(sheet);
    }

    const now      = new Date();
    const receptNo = data.receipt_no ||
                     'KWGC' + Utilities.formatDate(now, 'Asia/Tokyo', 'MMddHHmmssSSS');
    const newRow   = sheet.getLastRow() + 1;

    const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    sheet.getRange(newRow, 1, 1, HEADERS.length).setValues([[
      receptNo, dateStr,
      data.company_name     || '',
      data.company_furigana || '',
      data.rep_name         || '',
      data.rep_furigana     || '',
      data.staff_name       || '',
      data.staff_furigana   || '',
      data.zipcode          || '',
      data.address          || '',
      data.phone            || '',
      data.email            || '',
      data.category         || '',
      data.website_url      || '',
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
      '', '', '', '', '', '', '', '',
      'checkbox\n手動', 'タイムスタンプ\n自動', 'checkbox\n手動',
      'タイムスタンプ\n自動', '区分＋7桁\n自動',
      'checkbox\n手動', 'タイムスタンプ\n自動', 'タイムスタンプ\n自動',
    ]);
    sheet.getRange(2, 1, 1, TESAGYOU_HEADERS.length)
      .setFontSize(8).setFontColor('#888888').setBackground('#fffbf0').setWrap(true);
    sheet.setRowHeight(2, 36);
    applyTesagyouColumnWidths(sheet);
  }

  const newRow   = sheet.getLastRow() + 1;
  const kubun    = (data ? (data.category || '') : '').trim().toUpperCase();
  const autoSend = AUTO_SEND_KUBUN.includes(kubun);

  // A: 受付番号
  sheet.getRange(newRow, 1).setValue(receptNo);

  // B〜H: XLOOKUP（セミコロン区切りはロケール依存のためカンマ使用）
  Object.entries(TESAGYOU_LOOKUP_COLS).forEach(([col, srcCol]) => {
    const formula =
      `=IFERROR(XLOOKUP($A${newRow},'${DEFAULT_SHEET_NAME}'!$A:$A,'${DEFAULT_SHEET_NAME}'!$${srcCol}:$${srcCol}),"見つかりません")`;
    sheet.getRange(newRow, Number(col)).setFormula(formula);
  });

  // I(9): 受付完了 — B〜E は申込時に自動完了
  sheet.getRange(newRow, COL_UKETSUKE).insertCheckboxes();
  if (autoSend) sheet.getRange(newRow, COL_UKETSUKE).setValue(true);

  // J(10): 請求書送信日時 — B〜E は申込時に自動送信済みのため現在日時
  if (autoSend) sheet.getRange(newRow, COL_INV_DATE).setValue(nowStr());

  // K(11): 入金完了 checkbox
  sheet.getRange(newRow, COL_NYUKIN).insertCheckboxes();

  // N(14): 案内実施 checkbox
  sheet.getRange(newRow, COL_ANNAIBUN).insertCheckboxes();

  // L, M, O, P は空（各トリガーが自動設定）
}

function applyColumnWidths(sheet) {
  [160,150,200,180,150,150,120,120,90,220,120,220,60,200]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

function applyTesagyouColumnWidths(sheet) {
  [160, 60, 120, 200, 200, 150, 200, 160, 70, 150, 70, 150, 110, 70, 150, 150]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

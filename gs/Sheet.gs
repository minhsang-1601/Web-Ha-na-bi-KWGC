// ─── スプレッドシート書き込み ──────────────────────────────────────────────────

function appendRow(data, sheetName) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#d0e4f7');
      sheet.getRange(2, 9,  sheet.getMaxRows() - 1).setNumberFormat('@');
      sheet.getRange(2, 11, sheet.getMaxRows() - 1).setNumberFormat('@');
      applyColumnWidths(sheet);
    }

    const now      = new Date();
    const receptNo = data.receipt_no ||
                     'KWGC' + Utilities.formatDate(now, 'Asia/Tokyo', 'MMddHHmmssSSS');
    const newRow   = sheet.getLastRow() + 1;

    sheet.getRange(newRow, 9).setNumberFormat('@');
    sheet.getRange(newRow, 11).setNumberFormat('@');

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
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(sheetName2);
  if (!sheet) sheet = ss.insertSheet(sheetName2);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(TESAGYOU_HEADERS);
    sheet.getRange(1, 1, 1, TESAGYOU_HEADERS.length).setFontWeight('bold').setBackground('#fce8b2');
    [160, 80, 120, 200, 200, 150, 200, 160, 80, 80, 80, 100].forEach((w, i) =>
      sheet.setColumnWidth(i + 1, w)
    );
    sheet.getRange(2, 3, sheet.getMaxRows() - 1).setNumberFormat('@');
  }

  const newRow   = sheet.getLastRow() + 1;
  const kubun    = data ? (data.category || '').trim().toUpperCase() : '';
  const autoSend = AUTO_SEND_KUBUN.includes(kubun);

  sheet.getRange(newRow, 1).setValue(receptNo);
  sheet.getRange(newRow, 3).setNumberFormat('@');

  Object.entries(TESAGYOU_LOOKUP_COLS).forEach(([col, srcCol]) => {
    const formula =
      `=IFERROR(XLOOKUP($A${newRow};'${DEFAULT_SHEET_NAME}'!$A:$A;'${DEFAULT_SHEET_NAME}'!$${srcCol}:$${srcCol});"見つかりません")`;
    sheet.getRange(newRow, Number(col)).setFormula(formula);
  });

  // I列: 受付完了
  sheet.getRange(newRow, 9).insertCheckboxes();
  if (autoSend) sheet.getRange(newRow, 9).setValue(true);
  // J列: 請求書送信
  sheet.getRange(newRow, 10).setValue(autoSend ? '済み' : '未');
  // K列: 入金完了
  sheet.getRange(newRow, 11).insertCheckboxes();
  // L列: お礼状送付
  sheet.getRange(newRow, 12).setValue('未');
}

function applyColumnWidths(sheet) {
  [160,150,200,180,150,150,120,120,90,220,120,220,60,200].forEach((w, i) =>
    sheet.setColumnWidth(i + 1, w)
  );
}

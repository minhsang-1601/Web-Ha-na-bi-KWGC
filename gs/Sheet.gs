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
    ]]);

    SpreadsheetApp.flush();
    return receptNo;
  } finally {
    lock.releaseLock();
  }
}

function appendToTetsugyoSheet(receptNo, sheetName2) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(sheetName2);
  if (!sheet) sheet = ss.insertSheet(sheetName2);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(TETSUGYO_HEADERS);
    sheet.getRange(1, 1, 1, TETSUGYO_HEADERS.length).setFontWeight('bold').setBackground('#fce8b2');
    [160, 80, 120, 200, 200, 150, 200, 160, 80, 80].forEach((w, i) =>
      sheet.setColumnWidth(i + 1, w)
    );
    sheet.getRange(2, 3, sheet.getMaxRows() - 1).setNumberFormat('@');
  }

  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1).setValue(receptNo);
  sheet.getRange(newRow, 3).setNumberFormat('@');

  Object.entries(TETSUGYO_LOOKUP_COLS).forEach(([col, srcCol]) => {
    const formula =
      `=IFERROR(XLOOKUP($A${newRow};'${DEFAULT_SHEET_NAME}'!$A:$A;'${DEFAULT_SHEET_NAME}'!$${srcCol}:$${srcCol});"見つかりません")`;
    sheet.getRange(newRow, Number(col)).setFormula(formula);
  });

  sheet.getRange(newRow, 9).insertCheckboxes();
  sheet.getRange(newRow, 10).setValue('未');
}

function applyColumnWidths(sheet) {
  [160,150,200,180,150,150,120,120,90,220,120,220,60].forEach((w, i) =>
    sheet.setColumnWidth(i + 1, w)
  );
}

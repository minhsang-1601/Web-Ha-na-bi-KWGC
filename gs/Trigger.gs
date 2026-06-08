// ─── チェックボックストリガー（手作業シート） ──────────────────────────────────
// ⚠️ GAS エディタ「トリガー」→ onEditInstallable → スプレッドシート → 編集時 で登録すること
//
// 列構成:
//   I (9)  : 受付完了   — checkbox (S/A: 手動, B〜E: 自動)
//   J (10) : 請求書送信  — 未/済み
//   K (11) : 入金完了   — checkbox
//   L (12) : お礼状送付  — 未/済み

function onEditInstallable(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== DEFAULT_SHEET_NAME2) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row <= 1) return;

  // I列（9）: 受付完了 → 請求書送信トリガー
  if (col === 9) {
    handleUketsuke(e, sheet, row);
    return;
  }

  // K列（11）: 入金完了 → お礼状送付トリガー
  if (col === 11) {
    handleNyukin(e, sheet, row);
    return;
  }
}

// ─── 受付完了チェック → 請求書送信（S/A のみ手動） ────────────────────────────

function handleUketsuke(e, sheet, row) {
  const invoiceStatusCell = sheet.getRange(row, 10); // J列: 請求書送信
  const invoiceStatus     = invoiceStatusCell.getValue();

  if (e.range.getValue() === false) {
    if (invoiceStatus === '済み') {
      const ui  = SpreadsheetApp.getUi();
      const res = ui.alert(
        '⚠️ チェックを外しますか？',
        '「請求書送信」が「済み」→「未」に戻ります。よろしいですか？',
        ui.ButtonSet.YES_NO
      );
      if (res === ui.Button.YES) {
        invoiceStatusCell.setValue('未');
      } else {
        e.range.setValue(true);
      }
    }
    return;
  }

  if (invoiceStatus === '済み') {
    SpreadsheetApp.getUi().alert('⚠️ 請求書はすでに「済み」です。');
    e.range.setValue(false);
    return;
  }

  const receptNo  = sheet.getRange(row, 1).getValue();
  const mainSheet = e.source.getSheetByName(DEFAULT_SHEET_NAME);
  if (!mainSheet) {
    SpreadsheetApp.getUi().alert('申込みシートが見つかりません。');
    e.range.setValue(false);
    return;
  }

  const data = findRowByReceptNo(mainSheet, receptNo);
  if (!data || !data.email) {
    SpreadsheetApp.getUi().alert('メールアドレスが見つかりません。');
    e.range.setValue(false);
    return;
  }

  // 確認ダイアログを表示してから送信
  const tpl = HtmlService.createTemplateFromFile('ConfirmInvoiceDialog');
  tpl.receptNo     = receptNo;
  tpl.company_name = data.company_name || '';
  tpl.rep_name     = data.rep_name     || '';
  tpl.email        = data.email        || '';
  tpl.row          = row;

  SpreadsheetApp.getUi().showModalDialog(
    tpl.evaluate().setWidth(420).setHeight(280),
    '請求書送信の確認'
  );
}

// ─── 入金完了チェック → お礼状送付トリガー ────────────────────────────────────

function handleNyukin(e, sheet, row) {
  const statusCell = sheet.getRange(row, 12); // L列: お礼状送付
  const status     = statusCell.getValue();

  if (e.range.getValue() === false) {
    if (status === '済み') {
      const ui  = SpreadsheetApp.getUi();
      const res = ui.alert(
        '⚠️ チェックを外しますか？',
        '「お礼状送付」が「済み」→「未」に戻ります。よろしいですか？',
        ui.ButtonSet.YES_NO
      );
      if (res === ui.Button.YES) {
        statusCell.setValue('未');
      } else {
        e.range.setValue(true);
      }
    }
    return;
  }

  if (status === '済み') {
    SpreadsheetApp.getUi().alert('⚠️ お礼状はすでに「済み」です。\n送付済みかどうかご確認ください。');
    e.range.setValue(false);
    return;
  }

  // 請求書が未送信の場合はブロック
  const invoiceStatus = sheet.getRange(row, 10).getValue(); // J列: 請求書送信
  if (invoiceStatus !== '済み') {
    SpreadsheetApp.getUi().alert('⚠️ 請求書がまだ送信されていません（「未」）。\nお礼状を送付する前に請求書を送信してください。');
    e.range.setValue(false);
    return;
  }

  const receptNo  = sheet.getRange(row, 1).getValue();
  const category  = sheet.getRange(row, 2).getValue();
  const mainSheet = e.source.getSheetByName(DEFAULT_SHEET_NAME);

  if (!mainSheet) {
    SpreadsheetApp.getUi().alert('申込みシートが見つかりません。');
    e.range.setValue(false);
    return;
  }

  const data = findRowByReceptNo(mainSheet, receptNo);
  if (!data || !data.email) {
    SpreadsheetApp.getUi().alert('メールアドレスが見つかりません。手動でご確認ください。');
    e.range.setValue(false);
    return;
  }

  const tpl = HtmlService.createTemplateFromFile('ConfirmOreijouDialog');
  tpl.receptNo     = receptNo;
  tpl.company_name = data.company_name || '';
  tpl.rep_name     = data.rep_name     || '';
  tpl.category     = category          || '';
  tpl.email        = data.email        || '';
  tpl.row          = row;

  SpreadsheetApp.getUi().showModalDialog(
    tpl.evaluate().setWidth(420).setHeight(310),
    'お礼状送付の確認'
  );
}

// ConfirmInvoiceDialog から呼ばれる（確認後に実際に送信）
function sendInvoiceConfirmed(row, receptNo) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const tetsuSheet = ss.getSheetByName(DEFAULT_SHEET_NAME2);
  const mainSheet  = ss.getSheetByName(DEFAULT_SHEET_NAME);

  const data = findRowByReceptNo(mainSheet, receptNo);
  if (!data) throw new Error('受付番号が見つかりません: ' + receptNo);

  const pdf = generateInvoicePdf(data, receptNo);
  sendInvoiceEmail(data, receptNo, pdf);
  tetsuSheet.getRange(row, 10).setValue('済み'); // J列: 請求書送信
}

// ConfirmInvoiceDialog からキャンセル時に呼ばれる
function cancelInvoiceSend(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DEFAULT_SHEET_NAME2);
  if (sheet) sheet.getRange(row, 9).setValue(false); // I列: 受付完了を戻す
}

function sendInvoiceEmail(data, receptNo, pdf) {
  const props = PropertiesService.getScriptProperties();
  let subject = props.getProperty('MAIL_SUBJECT') ||
    '【第5回川口花火大会】申込受理書兼請求書のご連絡（受付番号：{{receipt_no}}）';
  let body    = props.getProperty('MAIL_BODY') || '';

  const vars = {
    company_name: data.company_name || '',
    rep_name:     data.rep_name     || '',
    staff_name:   data.staff_name   || '',
    category:     data.category     || '',
    receipt_no:   receptNo          || '',
    date:         Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
  };
  Object.entries(vars).forEach(([key, val]) => {
    const re = new RegExp(`\\{\\{${key}\\}\\}`, 'g');
    subject  = subject.replace(re, val);
    body     = body.replace(re, val);
  });

  const mailOptions = { to: data.email, subject, body, replyTo: OFFICE_EMAIL };
  if (pdf) {
    mailOptions.attachments = [
      pdf.setName(`申込受理書兼請求書_${data.company_name || receptNo}.pdf`)
    ];
  }
  MailApp.sendEmail(mailOptions);
}

function sendOreijouConfirmed(row, receptNo) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const tetsuSheet = ss.getSheetByName(DEFAULT_SHEET_NAME2);
  const mainSheet  = ss.getSheetByName(DEFAULT_SHEET_NAME);

  const data = findRowByReceptNo(mainSheet, receptNo);
  if (!data) throw new Error('受付番号が見つかりません: ' + receptNo);

  sendOreijouEmail(data, receptNo);
  tetsuSheet.getRange(row, 12).setValue('済み'); // L列: お礼状送付
}

function cancelOreijou(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DEFAULT_SHEET_NAME2);
  if (sheet) sheet.getRange(row, 11).setValue(false); // K列: 入金完了
}

function findRowByReceptNo(sheet, receptNo) {
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === receptNo) {
      return {
        company_name:     rows[i][2],   // C列
        company_furigana: rows[i][3],   // D列
        rep_name:         rows[i][4],   // E列
        rep_furigana:     rows[i][5],   // F列
        staff_name:       rows[i][6],   // G列
        staff_furigana:   rows[i][7],   // H列
        zipcode:          rows[i][8],   // I列
        address:          rows[i][9],   // J列
        phone:            rows[i][10],  // K列
        email:            rows[i][11],  // L列
        category:         rows[i][12],  // M列
        website_url:      rows[i][13],  // N列
      };
    }
  }
  return null;
}

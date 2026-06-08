// ─── チェックボックストリガー（手作業シート I列） ──────────────────────────────
// ⚠️ GAS エディタ「トリガー」→ onEditInstallable → スプレッドシート → 編集時 で登録すること

function onEditInstallable(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== DEFAULT_SHEET_NAME2) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (col !== 9 || row <= 1) return;

  const statusCell = sheet.getRange(row, 10); // J列
  const status     = statusCell.getValue();

  // ── チェックを外した場合 ──
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

  // ── チェックを入れた場合 ──
  if (status === '済み') {
    SpreadsheetApp.getUi().alert(
      '⚠️ お礼状はすでに「済み」です。\n送付済みかどうかご確認ください。'
    );
    e.range.setValue(false);
    return;
  }

  // 未 or 空 → 確認ダイアログ表示
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

function sendOreijouConfirmed(row, receptNo) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const tetsuSheet = ss.getSheetByName(DEFAULT_SHEET_NAME2);
  const mainSheet  = ss.getSheetByName(DEFAULT_SHEET_NAME);

  const data = findRowByReceptNo(mainSheet, receptNo);
  if (!data) throw new Error('受付番号が見つかりません: ' + receptNo);

  sendOreijouEmail(data, receptNo);
  tetsuSheet.getRange(row, 10).setValue('済み');
}

function cancelOreijou(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DEFAULT_SHEET_NAME2);
  if (sheet) sheet.getRange(row, 9).setValue(false);
}

function findRowByReceptNo(sheet, receptNo) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === receptNo) {
      return {
        company_name: data[i][2],   // C列
        rep_name:     data[i][4],   // E列
        email:        data[i][11],  // L列
      };
    }
  }
  return null;
}

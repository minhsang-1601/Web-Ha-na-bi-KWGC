// ─── チェックボックストリガー（手作業シート） ──────────────────────────────────
// ⚠️ GAS エディタ「トリガー」→ onEditInstallable → スプレッドシート → 編集時 で登録
//
// 手作業 列構成（I〜L）:
//   I (9)  受付完了         checkbox  手動  → 請求書送信確認ダイアログ
//   J (10) 請求書送信日時   timestamp 自動  ← sendInvoiceConfirmed でセット
//   K (11) 入金完了         checkbox  手動  → お礼状送信確認ダイアログ ※J が必須
//   L (12) お礼状送信日時   timestamp 自動  ← sendNyukinConfirmed でセット

function onEditInstallable(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== DEFAULT_SHEET_NAME2) return;
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row <= 2) return; // ヘッダー行・サブヘッダー行をスキップ

  if (col === COL_UKETSUKE)  handleUketsuke(e, sheet, row);
  if (col === COL_NYUKIN)    handleNyukin(e, sheet, row);
}

// ─── I列: 受付完了 → 請求書送信 ────────────────────────────────────────────────

function handleUketsuke(e, sheet, row) {
  const invDateCell = sheet.getRange(row, COL_INV_DATE);
  const invDate     = invDateCell.getValue();

  if (e.range.getValue() === false) {
    // ── クリア順序チェック: 下流（K→L）を先に解除しないと外せない ──
    const oreijouDate = sheet.getRange(row, COL_OREIJOU_DATE).getValue();
    const nyukin      = sheet.getRange(row, COL_NYUKIN).getValue();

    // I を外す → 全下流（K・L）を自動クリア
    const hasDownstream = nyukin || oreijouDate || invDate;
    if (hasDownstream) {
      const detail = [
        oreijouDate ? `　お礼状送信日時（L）：${formatTs(oreijouDate)}` : null,
        nyukin      ? `　入金完了（K）：チェック済み`                    : null,
        invDate     ? `　請求書送信日時（J）：${formatTs(invDate)}`      : null,
      ].filter(Boolean).join('\n');
      const res = SpreadsheetApp.getUi().alert(
        '⚠️ チェックを外しますか？',
        `以下の値が自動でクリアされます。\n\n${detail}\n\nよろしいですか？`,
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      if (res === SpreadsheetApp.getUi().Button.YES) {
        sheet.getRange(row, COL_NYUKIN).setValue(false);
        sheet.getRange(row, COL_OREIJOU_DATE).clearContent();
        invDateCell.clearContent();
      } else {
        e.range.setValue(true);
      }
    }
    return;
  }

  if (invDate) {
    SpreadsheetApp.getUi().alert(`⚠️ 請求書はすでに送信済みです。\n送信日時：${formatTs(invDate)}`);
    e.range.setValue(false);
    return;
  }

  // メール残数チェック
  if (_blockSendIfLowQuota(e.range)) return;

  const receptNo  = sheet.getRange(row, COL_RECEPT_NO).getValue();
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

  const tpl = HtmlService.createTemplateFromFile('ConfirmInvoiceDialog');
  tpl.receptNo     = receptNo;
  tpl.company_name = data.company_name || '';
  tpl.rep_name     = data.rep_name     || '';
  tpl.email        = data.email        || '';
  tpl.row          = row;
  SpreadsheetApp.getUi().showModalDialog(
    tpl.evaluate().setWidth(420).setHeight(280), '請求書送信の確認'
  );
}

// ─── K列: 入金完了 → お礼状送信 ──────────────────────────────────────────────────

function handleNyukin(e, sheet, row) {
  const oreijouDateCell = sheet.getRange(row, COL_OREIJOU_DATE);
  const oreijouDate     = oreijouDateCell.getValue();

  if (e.range.getValue() === false) {
    // K を外す → L を自動クリア（値がある場合は確認）
    if (oreijouDate) {
      const detail = `　お礼状送信日時（L）：${formatTs(oreijouDate)}`;
      const res = SpreadsheetApp.getUi().alert(
        '⚠️ チェックを外しますか？',
        `以下の値が自動でクリアされます。\n\n${detail}\n\nよろしいですか？`,
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      if (res === SpreadsheetApp.getUi().Button.YES) {
        oreijouDateCell.clearContent();
      } else {
        e.range.setValue(true);
      }
    }
    return;
  }

  // 前提条件: 請求書送信済み（J に日時あり）
  const invDate = sheet.getRange(row, COL_INV_DATE).getValue();
  if (!invDate) {
    SpreadsheetApp.getUi().alert(
      '⚠️ 請求書がまだ送信されていません。\n先に「受付完了」をチェックして請求書を送信してください。'
    );
    e.range.setValue(false);
    return;
  }

  if (oreijouDate) {
    SpreadsheetApp.getUi().alert(`⚠️ お礼状はすでに送信済みです。\n送信日時：${formatTs(oreijouDate)}`);
    e.range.setValue(false);
    return;
  }

  // メール残数チェック
  if (_blockSendIfLowQuota(e.range)) return;

  const receptNo  = sheet.getRange(row, COL_RECEPT_NO).getValue();
  const mainSheet = e.source.getSheetByName(DEFAULT_SHEET_NAME);
  const data      = findRowByReceptNo(mainSheet, receptNo);
  if (!data || !data.email) {
    SpreadsheetApp.getUi().alert('メールアドレスが見つかりません。');
    e.range.setValue(false);
    return;
  }

  const tpl = HtmlService.createTemplateFromFile('ConfirmNyukinDialog');
  tpl.receptNo     = receptNo;
  tpl.company_name = data.company_name || '';
  tpl.email        = data.email        || '';
  tpl.row          = row;
  SpreadsheetApp.getUi().showModalDialog(
    tpl.evaluate().setWidth(420).setHeight(250), '入金確認・お礼状送信'
  );
}



/** データ用スプレッドシートにも同じメニューを表示（onOpen trigger として登録） */
function onOpenEventSheet() {
  SpreadsheetApp.getUi()
    .createMenu('📋 協賛管理')
    .addToUi();
}

// ─── ダイアログからのコールバック ─────────────────────────────────────────────

function sendInvoiceConfirmed(row, receptNo) {
  const dataSs     = getDataSpreadsheet();
  const tetsuSheet = dataSs.getSheetByName(DEFAULT_SHEET_NAME2);
  const mainSheet  = dataSs.getSheetByName(DEFAULT_SHEET_NAME);
  const data = findRowByReceptNo(mainSheet, receptNo);
  if (!data) throw new Error('受付番号が見つかりません: ' + receptNo);

  const pdf = generateInvoicePdf(data, receptNo);

  // 区分を取得して、S/A と B~E で異なるテンプレートで送信
  const kubun = String(tetsuSheet.getRange(row, 2).getValue()).trim().toUpperCase();
  if (['S', 'A'].includes(kubun)) {
    // S/A: 抽選確定・請求書送付メール + PDF
    sendSaInvoiceEmail(data, receptNo, pdf);
  } else {
    // B~E: 確認メール + PDF
    sendConfirmationEmail(data, receptNo, pdf);
  }

  tetsuSheet.getRange(row, COL_INV_DATE).setValue(nowStr());
}

function cancelInvoiceSend(row) {
  const sheet = getDataSpreadsheet().getSheetByName(DEFAULT_SHEET_NAME2);
  if (sheet) sheet.getRange(row, COL_UKETSUKE).setValue(false);
}

function sendNyukinConfirmed(row, receptNo) {
  const dataSs     = getDataSpreadsheet();
  const tetsuSheet = dataSs.getSheetByName(DEFAULT_SHEET_NAME2);
  const mainSheet  = dataSs.getSheetByName(DEFAULT_SHEET_NAME);
  const data = findRowByReceptNo(mainSheet, receptNo);
  if (!data) throw new Error('受付番号が見つかりません: ' + receptNo);

  console.log('DEBUG sendNyukinConfirmed: receptNo=', receptNo);
  console.log('DEBUG data.email=', data.email);

  // お礼状メール送信
  try {
    console.log('DEBUG about to send oreijou email');
    sendOreijouEmail(data, receptNo);
    console.log('DEBUG oreijou email sent successfully');
  } catch (e) {
    console.error('DEBUG oreijou email error:', e.message, e.stack);
    throw e;
  }

  tetsuSheet.getRange(row, COL_OREIJOU_DATE).setValue(nowStr());
}

function cancelNyukin(row) {
  const sheet = getDataSpreadsheet().getSheetByName(DEFAULT_SHEET_NAME2);
  if (sheet) sheet.getRange(row, COL_NYUKIN).setValue(false);
}

// ─── メール残数チェック（手作業シートの送信操作用） ──────────────────────────────

/**
 * メール残数が最低ラインを下回っていれば警告して送信をブロックする
 * @param {Object} revertCell - 残数不足時に false に戻すチェックボックスセル（任意）
 * @returns {boolean} true = ブロックした（送信中止）, false = 続行可
 */
function _blockSendIfLowQuota(revertCell) {
  let quota = 0;
  try { quota = MailApp.getRemainingDailyQuota(); } catch (_) { quota = 0; }

  let minQuota = 5;
  try { minQuota = getMinMailQuota(); } catch (_) {}

  if (quota < minQuota) {
    try { _notifyLowQuota(quota); } catch (_) {}
    SpreadsheetApp.getUi().alert(
      '⚠️ 送信できません',
      `本日のメール送信残数が不足しています。\n\n` +
      `　現在の残数：${quota} 通\n` +
      `　最低ライン：${minQuota} 通\n\n` +
      `残数は約1日で自動リセットされます。\n時間をおいて再度お試しください。`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    if (revertCell) {
      try { revertCell.setValue(false); } catch (_) {}
    }
    return true;
  }
  return false;
}

// ─── ユーティリティ ─────────────────────────────────────────────────────────────

function sendInvoiceEmail(data, receptNo, pdf) {
  const props = PropertiesService.getScriptProperties();
  let subject = props.getProperty('MAIL_SUBJECT') ||
    '【{{event_name}}】申込受理書兼請求書のご連絡（受付番号：{{receipt_no}}）';
  let body    = props.getProperty('MAIL_BODY') || _defaultConfirmBody();

  const vars = {
    company_name: data.company_name || '',
    rep_name:     data.rep_name     || '',
    staff_name:   data.staff_name   || '',
    category:     data.category     || '',
    receipt_no:   receptNo          || '',
    date:         nowStr(),
    event_name:   getEventName(),
    payment_due:  getPaymentDue(),
    office_email: getOfficeEmail(),
    office_hours: getOfficeHours(),
  };
  Object.entries(vars).forEach(([key, val]) => {
    const re = new RegExp(`\\{\\{${key}\\}\\}`, 'g');
    subject  = subject.replace(re, val);
    body     = body.replace(re, val);
  });

  const officeEmail = getOfficeEmail();
  const mailOptions = { to: data.email, subject, body };
  if (_validEmail(officeEmail)) { mailOptions.cc = officeEmail; mailOptions.replyTo = officeEmail; }
  if (pdf) {
    mailOptions.attachments = [
      pdf.setName(`申込受理書兼請求書_${data.company_name || receptNo}.pdf`)
    ];
  }
  MailApp.sendEmail(mailOptions);
}

function findRowByReceptNo(sheet, receptNo) {
  // 協賛申込み一覧: A=受付番号 B=受付日時 C=会社名 D=会社名ふりがな
  //                E=代表者役職 F=役職ふりがな G=代表者名 H=代表者名ふりがな
  //                I=担当者名 J=担当者ふりがな K=郵便番号 L=住所 M=電話番号 N=メール O=区分 P=会社HP URL
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(receptNo).trim()) { // A列: 受付番号
      return {
        company_name:          rows[i][2],   // C
        company_furigana:      rows[i][3],   // D
        rep_position:          rows[i][4],   // E
        rep_position_furigana: rows[i][5],   // F
        rep_name:              rows[i][6],   // G
        rep_furigana:          rows[i][7],   // H
        staff_name:            rows[i][8],   // I
        staff_furigana:        rows[i][9],   // J
        zipcode:               rows[i][10],  // K
        address:               rows[i][11],  // L
        phone:                 rows[i][12],  // M
        email:                 rows[i][13],  // N
        category:              rows[i][14],  // O
        website_url:           rows[i][15],  // P
      };
    }
  }
  return null;
}

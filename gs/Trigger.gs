// ─── チェックボックストリガー（手作業シート） ──────────────────────────────────
// ⚠️ GAS エディタ「トリガー」→ onEditInstallable → スプレッドシート → 編集時 で登録
//
// 手作業 列構成（I〜P）:
//   I (9)  受付完了         checkbox  手動  → 請求書送信確認ダイアログ
//   J (10) 請求書送信日時   timestamp 自動  ← sendInvoiceConfirmed でセット
//   K (11) 入金完了         checkbox  手動  → 座席割当確認ダイアログ ※J が必須
//   L (12) 座席割当送信日時 timestamp 自動  ← sendNyukinConfirmed でセット
//   M (13) 座席番号         text      自動  ← sendNyukinConfirmed でセット
//   N (14) 案内実施         checkbox  手動  → 案内送信確認ダイアログ ※L が必須
//   O (15) 案内送信日時     timestamp 自動  ← sendAnnaibunConfirmed でセット
//   P (16) お礼状送信日時   timestamp 自動  ← sendOreijouConfirmed でセット ※O が必須
//         お礼状はカスタムメニュー「お礼状送信」から起動

function onEditInstallable(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== DEFAULT_SHEET_NAME2) return;
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row <= 2) return; // ヘッダー行・サブヘッダー行をスキップ

  if (col === COL_UKETSUKE)  handleUketsuke(e, sheet, row);
  if (col === COL_NYUKIN)    handleNyukin(e, sheet, row);
  if (col === COL_ANNAIBUN)  handleAnnaibun(e, sheet, row);
}

// ─── I列: 受付完了 → 請求書送信 ────────────────────────────────────────────────

function handleUketsuke(e, sheet, row) {
  const invDateCell = sheet.getRange(row, COL_INV_DATE);
  const invDate     = invDateCell.getValue();

  if (e.range.getValue() === false) {
    // ── クリア順序チェック: 下流（L→O→P→Q）を先に解除しないと外せない ──
    const oreijouDate = sheet.getRange(row, COL_OREIJOU_DATE).getValue();
    const annaiDate   = sheet.getRange(row, COL_ANNAI_DATE).getValue();
    const annaibun    = sheet.getRange(row, COL_ANNAIBUN).getValue();
    const seatDate    = sheet.getRange(row, COL_SEAT_DATE).getValue();
    const nyukin      = sheet.getRange(row, COL_NYUKIN).getValue();

    // J を外す → 全下流（L・O・P・Q・M・N・K）を自動クリア
    const hasDownstream = nyukin || seatDate || annaibun || annaiDate || oreijouDate || invDate;
    if (hasDownstream) {
      const detail = [
        oreijouDate ? `　お礼状送信日時（Q）：${formatTs(oreijouDate)}` : null,
        annaiDate   ? `　案内送信日時（P）：${formatTs(annaiDate)}`      : null,
        annaibun    ? `　案内実施（O）：チェック済み`                    : null,
        seatDate    ? `　座席割当送信日時（M）：${formatTs(seatDate)}`   : null,
        seatDate    ? `　座席番号（N）`                                  : null,
        nyukin      ? `　入金完了（L）：チェック済み`                    : null,
        invDate     ? `　請求書送信日時（K）：${formatTs(invDate)}`      : null,
      ].filter(Boolean).join('\n');
      const res = SpreadsheetApp.getUi().alert(
        '⚠️ チェックを外しますか？',
        `以下の値が自動でクリアされます。\n\n${detail}\n\nよろしいですか？`,
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      if (res === SpreadsheetApp.getUi().Button.YES) {
        sheet.getRange(row, COL_NYUKIN).setValue(false);
        sheet.getRange(row, COL_SEAT_DATE).clearContent();
        sheet.getRange(row, COL_SEAT_NO).clearContent();
        sheet.getRange(row, COL_ANNAIBUN).setValue(false);
        sheet.getRange(row, COL_ANNAI_DATE).clearContent();
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

// ─── K列: 入金完了 → 座席割当 ──────────────────────────────────────────────────

function handleNyukin(e, sheet, row) {
  const seatDateCell = sheet.getRange(row, COL_SEAT_DATE);
  const seatDate     = seatDateCell.getValue();

  if (e.range.getValue() === false) {
    // ── クリア順序チェック: 下流（O→P→Q）を先に解除しないと外せない ──
    const oreijouDate = sheet.getRange(row, COL_OREIJOU_DATE).getValue();
    const annaiDate   = sheet.getRange(row, COL_ANNAI_DATE).getValue();
    const annaibun    = sheet.getRange(row, COL_ANNAIBUN).getValue();

    // L を外す → O・P・Q・M・N を自動クリア（値がある場合は確認）
    const hasDownstream = annaibun || annaiDate || oreijouDate || seatDate;
    if (hasDownstream) {
      const detail = [
        oreijouDate ? `　お礼状送信日時（Q）：${formatTs(oreijouDate)}` : null,
        annaiDate   ? `　案内送信日時（P）：${formatTs(annaiDate)}`      : null,
        annaibun    ? `　案内実施（O）：チェック済み`                    : null,
        seatDate    ? `　座席割当送信日時（M）：${formatTs(seatDate)}`   : null,
        seatDate    ? `　座席番号（N）`                                  : null,
      ].filter(Boolean).join('\n');
      const res = SpreadsheetApp.getUi().alert(
        '⚠️ チェックを外しますか？',
        `以下の値が自動でクリアされます。\n\n${detail}\n\nよろしいですか？`,
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      if (res === SpreadsheetApp.getUi().Button.YES) {
        sheet.getRange(row, COL_ANNAIBUN).setValue(false);
        sheet.getRange(row, COL_ANNAI_DATE).clearContent();
        sheet.getRange(row, COL_OREIJOU_DATE).clearContent();
        seatDateCell.clearContent();
        sheet.getRange(row, COL_SEAT_NO).clearContent();
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

  if (seatDate) {
    SpreadsheetApp.getUi().alert(`⚠️ 座席割当はすでに完了しています。\n割当日時：${formatTs(seatDate)}`);
    e.range.setValue(false);
    return;
  }

  // メール残数チェック
  if (_blockSendIfLowQuota(e.range)) return;

  const receptNo  = sheet.getRange(row, COL_RECEPT_NO).getValue();
  const kubun     = sheet.getRange(row, 3).getValue(); // C列: 区分
  const seatNo    = generateSeatNo(kubun, sheet);      // 区分ごとの連番
  const mainSheet = e.source.getSheetByName(DEFAULT_SHEET_NAME);
  const data      = findRowByReceptNo(mainSheet, receptNo);
  if (!data) {
    SpreadsheetApp.getUi().alert('申込みデータが見つかりません。');
    e.range.setValue(false);
    return;
  }

  const tpl = HtmlService.createTemplateFromFile('ConfirmNyukinDialog');
  tpl.receptNo     = receptNo;
  tpl.company_name = data.company_name || '';
  tpl.kubun        = kubun             || '';
  tpl.seatNo       = seatNo;
  tpl.email        = data.email        || '';
  tpl.row          = row;
  SpreadsheetApp.getUi().showModalDialog(
    tpl.evaluate().setWidth(420).setHeight(300), '入金確認・座席番号発行'
  );
}

// ─── N列: 案内実施 → 案内文送信 ───────────────────────────────────────────────
// 送信順序:  O(案内実施✔) → P(案内送信日時) → Q(お礼状送信日時)
// クリア順序: Q → P → O の順で削除が必要（逆順強制）

function handleAnnaibun(e, sheet, row) {
  const annaiDateCell  = sheet.getRange(row, COL_ANNAI_DATE);
  const oreijouDateCell = sheet.getRange(row, COL_OREIJOU_DATE);
  const annaiDate      = annaiDateCell.getValue();
  const oreijouDate    = oreijouDateCell.getValue();

  if (e.range.getValue() === false) {
    // O を外す → P・Q を自動クリア（値がある場合は確認）
    const hasData = annaiDate || oreijouDate;
    if (hasData) {
      const detail = [
        annaiDate   ? `　案内送信日時（P）：${formatTs(annaiDate)}`   : null,
        oreijouDate ? `　お礼状送信日時（Q）：${formatTs(oreijouDate)}` : null,
      ].filter(Boolean).join('\n');
      const res = SpreadsheetApp.getUi().alert(
        '⚠️ チェックを外しますか？',
        `以下の値が自動でクリアされます。\n\n${detail}\n\nよろしいですか？`,
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      if (res === SpreadsheetApp.getUi().Button.YES) {
        annaiDateCell.clearContent();
        oreijouDateCell.clearContent();
      } else {
        e.range.setValue(true);
      }
    }
    return;
  }

  // 前提条件: 座席割当済み（L に日時あり）
  const seatDate = sheet.getRange(row, COL_SEAT_DATE).getValue();
  if (!seatDate) {
    SpreadsheetApp.getUi().alert(
      '⚠️ 座席割当がまだ完了していません。\n先に「入金完了」をチェックして座席を割り当ててください。'
    );
    e.range.setValue(false);
    return;
  }

  if (annaiDate) {
    SpreadsheetApp.getUi().alert(`⚠️ 案内はすでに送信済みです。\n送信日時：${formatTs(annaiDate)}`);
    e.range.setValue(false);
    return;
  }

  // メール残数チェック
  if (_blockSendIfLowQuota(e.range)) return;

  const receptNo  = sheet.getRange(row, COL_RECEPT_NO).getValue();
  const seatNo    = sheet.getRange(row, COL_SEAT_NO).getValue();
  const mainSheet = e.source.getSheetByName(DEFAULT_SHEET_NAME);
  const data      = findRowByReceptNo(mainSheet, receptNo);
  if (!data || !data.email) {
    SpreadsheetApp.getUi().alert('メールアドレスが見つかりません。');
    e.range.setValue(false);
    return;
  }

  const tpl = HtmlService.createTemplateFromFile('ConfirmAnnaibunDialog');
  tpl.receptNo     = receptNo;
  tpl.company_name = data.company_name || '';
  tpl.rep_name     = data.rep_name     || '';
  tpl.seatNo       = seatNo            || '';
  tpl.email        = data.email        || '';
  tpl.row          = row;
  SpreadsheetApp.getUi().showModalDialog(
    tpl.evaluate().setWidth(420).setHeight(300), '案内送信の確認'
  );
}

// ─── お礼状送信（カスタムメニュー「お礼状送信」から起動） ────────────────────────

function sendOreijouFromMenu() {
  const dataSs  = getDataSpreadsheet();
  const sheet   = dataSs.getSheetByName(DEFAULT_SHEET_NAME2);
  if (!sheet) { SpreadsheetApp.getUi().alert('手作業シートが見つかりません。'); return; }

  const activeRange = SpreadsheetApp.getActiveRange();
  const row = activeRange ? activeRange.getRow() : 0;
  if (!row || row <= 2) {
    SpreadsheetApp.getUi().alert('手作業シートで対象の行を選択してから実行してください。');
    return;
  }

  const oreijouDate = sheet.getRange(row, COL_OREIJOU_DATE).getValue();
  if (oreijouDate) {
    SpreadsheetApp.getUi().alert(`⚠️ お礼状はすでに送信済みです。\n送信日時：${formatTs(oreijouDate)}`);
    return;
  }

  const annaiDate = sheet.getRange(row, COL_ANNAI_DATE).getValue();
  if (!annaiDate) {
    SpreadsheetApp.getUi().alert('⚠️ 案内がまだ送信されていません。\n先に「案内実施」をチェックして案内を送信してください。');
    return;
  }

  // メール残数チェック
  if (_blockSendIfLowQuota(null)) return;

  const receptNo  = sheet.getRange(row, COL_RECEPT_NO).getValue();
  const mainSheet = dataSs.getSheetByName(DEFAULT_SHEET_NAME);
  const data      = findRowByReceptNo(mainSheet, receptNo);
  if (!data || !data.email) {
    SpreadsheetApp.getUi().alert('メールアドレスが見つかりません。');
    return;
  }

  const tpl = HtmlService.createTemplateFromFile('ConfirmOreijouDialog');
  tpl.receptNo     = receptNo;
  tpl.company_name = data.company_name || '';
  tpl.rep_name     = data.rep_name     || '';
  tpl.category     = data.category     || '';
  tpl.email        = data.email        || '';
  tpl.row          = row;
  SpreadsheetApp.getUi().showModalDialog(
    tpl.evaluate().setWidth(420).setHeight(310), 'お礼状送付の確認'
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
    // S/A: 受付確認メール + PDF
    sendReceiptOnlyEmail(data, receptNo, pdf);
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

function sendNyukinConfirmed(row, receptNo, seatNo) {
  const tetsuSheet = getDataSpreadsheet().getSheetByName(DEFAULT_SHEET_NAME2);
  tetsuSheet.getRange(row, COL_SEAT_DATE).setValue(nowStr());
  tetsuSheet.getRange(row, COL_SEAT_NO).setValue(seatNo);
}

function cancelNyukin(row) {
  const sheet = getDataSpreadsheet().getSheetByName(DEFAULT_SHEET_NAME2);
  if (sheet) sheet.getRange(row, COL_NYUKIN).setValue(false);
}

function sendAnnaibunConfirmed(row, receptNo) {
  const dataSs     = getDataSpreadsheet();
  const tetsuSheet = dataSs.getSheetByName(DEFAULT_SHEET_NAME2);
  const mainSheet  = dataSs.getSheetByName(DEFAULT_SHEET_NAME);
  const data = findRowByReceptNo(mainSheet, receptNo);
  if (!data) throw new Error('受付番号が見つかりません: ' + receptNo);

  // 案内文送信 → O列にタイムスタンプ
  sendAnnaibunEmail(data, receptNo);
  tetsuSheet.getRange(row, COL_ANNAI_DATE).setValue(nowStr());

  // お礼状を自動送信 → P列にタイムスタンプ
  sendOreijouEmail(data, receptNo);
  tetsuSheet.getRange(row, COL_OREIJOU_DATE).setValue(nowStr());
}

function cancelAnnaibun(row) {
  const sheet = getDataSpreadsheet().getSheetByName(DEFAULT_SHEET_NAME2);
  if (sheet) sheet.getRange(row, COL_ANNAIBUN).setValue(false);
}

function sendOreijouConfirmed(row, receptNo) {
  const dataSs     = getDataSpreadsheet();
  const tetsuSheet = dataSs.getSheetByName(DEFAULT_SHEET_NAME2);
  const mainSheet  = dataSs.getSheetByName(DEFAULT_SHEET_NAME);
  const data = findRowByReceptNo(mainSheet, receptNo);
  if (!data) throw new Error('受付番号が見つかりません: ' + receptNo);

  sendOreijouEmail(data, receptNo);
  tetsuSheet.getRange(row, COL_OREIJOU_DATE).setValue(nowStr());
}

function cancelOreijou(row) {
  // お礼状はメニューから起動のため checkbox なし → 何もしない
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

  const mailOptions = { to: data.email, cc: getOfficeEmail(), subject, body, replyTo: getOfficeEmail() };
  if (pdf) {
    mailOptions.attachments = [
      pdf.setName(`申込受理書兼請求書_${data.company_name || receptNo}.pdf`)
    ];
  }
  MailApp.sendEmail(mailOptions);
}

function findRowByReceptNo(sheet, receptNo) {
  // 協賛申込み一覧: A=受付番号 B=受付日時 C=個人名 D=個人名ふりがな
  //                E=役職代表者 F=役職ふりがな G=担当者 H=担当者ふりがな
  //                I=郵便番号 J=住所 K=電話番号 L=メール M=区分 N=会社HP URL
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(receptNo).trim()) { // A列: 受付番号
      return {
        company_name:     rows[i][2],   // C
        company_furigana: rows[i][3],   // D
        rep_name:         rows[i][4],   // E
        rep_furigana:     rows[i][5],   // F
        staff_name:       rows[i][6],   // G
        staff_furigana:   rows[i][7],   // H
        zipcode:          rows[i][8],   // I
        address:          rows[i][9],   // J
        phone:            rows[i][10],  // K
        email:            rows[i][11],  // L
        category:         rows[i][12],  // M
        website_url:      rows[i][13],  // N
      };
    }
  }
  return null;
}

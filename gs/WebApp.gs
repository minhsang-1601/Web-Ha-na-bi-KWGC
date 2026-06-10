// ─── Web App エントリーポイント ────────────────────────────────────────────────

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(`【${getEventName()}】協賛申込み`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * クライアント(google.script.run)から呼ばれる設定取得
 */
function getConfig() {
  return {
    startDate:  String(getConfigVal('START_DATE',  '2025-01-01T00:00:00')),
    endDate:    String(getConfigVal('END_DATE',     '2026-10-01T23:59:59')),
    sheetName1: DEFAULT_SHEET_NAME,
    sheetName2: DEFAULT_SHEET_NAME2,
    eventName:        getEventName(),
    receiptNoPrefix:  getReceiptNoPrefix(),
  };
}

/**
 * クライアント(google.script.run)から呼ばれるフォーム送信
 */
function submitForm(data) {
  // ─── サーバー側バリデーション ────────────────────────────────────────────────
  const zipcode = String(data.zipcode || '').trim().replace(/\D/g, '');
  const phone   = String(data.phone   || '').trim().replace(/\D/g, '');
  if (!/^\d{7}$/.test(zipcode)) {
    throw new Error('郵便番号は7桁の数字で入力してください。（入力値: ' + (data.zipcode || '') + '）');
  }
  if (!/^\d{9,11}$/.test(phone)) {
    throw new Error('電話番号は9〜11桁の数字で入力してください。（入力値: ' + (data.phone || '') + '）');
  }
  // バリデーション済みの値で上書き
  data.zipcode = zipcode;
  data.phone   = phone;

  const kubun    = (data.category || '').trim().toUpperCase();
  const autoSend = AUTO_SEND_KUBUN.includes(kubun);

  // ─── 管理ID処理 ─────────────────────────────────────────────────────────────
  let kanriId = String(data.kanri_id || '').trim();

  let isNewKanri = false;
  if (kanriId) {
    // 既存IDの場合 → 変更チェック → update_kanri フラグがあれば更新
    const doUpdate = data.update_kanri === 'yes';
    saveKanriId(kanriId, data, doUpdate);
  } else {
    // 新規ユーザー → 新ID発行 → 保存
    kanriId    = generateKanriId();
    saveKanriId(kanriId, data, true);
    isNewKanri = true;
  }
  data.kanri_id = kanriId;

  // ─── スプレッドシート登録 ─────────────────────────────────────────────────
  const receptNo = appendRow(data, DEFAULT_SHEET_NAME);
  appendToTesagyouSheet(receptNo, DEFAULT_SHEET_NAME2, data);

  if (data.email) {
    // 新規ユーザーには管理ID発行通知メールを先に送信
    if (isNewKanri) sendKanriIdIssuedEmail(data, kanriId);

    if (autoSend) {
      const pdf = generateInvoicePdf(data, receptNo);
      sendConfirmationEmail(data, receptNo, pdf, kanriId);
    } else {
      sendReceiptOnlyEmail(data, receptNo, kanriId);
    }
    _notifyOffice(data, receptNo, autoSend, kanriId);
  }

  return { result: 'success', receipt_no: receptNo, kanri_id: kanriId };
}

/** カスタムメニュー登録 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 協賛管理')
    .addItem('お礼状送信',  'sendOreijouFromMenu')
    .addSeparator()
    .addItem('Info シート作成',         'setupInfoSheet')
    .addItem('申込みヘッダー作成',      'setupHeaders')
    .addItem('プロジェクト初期化',      'initProject')
    .addSeparator()
    .addItem('メールテンプレ保存',      'setupAllMailTemplates')
    .addItem('手作業ヘッダー修正',      'fixTesagyouHeaders')
    .addToUi();
}

// ─── 内部ユーティリティ ────────────────────────────────────────────────────────

function _notifyOffice(data, receptNo, autoSent, kanriId) {
  const subject = `【協賛申込】${data.company_name || ''} (${data.category || ''}) 受付番号:${receptNo}`;
  const body = [
    '新規協賛申込みがありました。',
    '',
    `会社名　：${data.company_name || ''}`,
    `区分　　：${data.category || ''} プラン`,
    `担当者　：${data.staff_name || ''}`,
    `メール　：${data.email || ''}`,
    `受付番号：${receptNo}`,
    `管理ID　：${kanriId || ''}`,
    `請求書　：${autoSent ? '自動送信済み' : '手動送信待ち（S/Aプラン）'}`,
  ].join('\n');
  try {
    MailApp.sendEmail({ to: getOfficeEmail(), subject, body });
  } catch (e) {
    console.warn('事務局通知メール送信失敗:', e.message);
  }
}

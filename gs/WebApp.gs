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
  };
}

/**
 * クライアント(google.script.run)から呼ばれるフォーム送信
 */
function submitForm(data) {
  const kubun    = (data.category || '').trim().toUpperCase();
  const autoSend = AUTO_SEND_KUBUN.includes(kubun);

  const receptNo = appendRow(data, DEFAULT_SHEET_NAME);
  appendToTesagyouSheet(receptNo, DEFAULT_SHEET_NAME2, data);

  if (data.email) {
    if (autoSend) {
      // B〜E: 受付確認 + 請求書PDF を自動送信
      const pdf = generateInvoicePdf(data, receptNo);
      sendConfirmationEmail(data, receptNo, pdf);
    } else {
      // S/A: 受付確認のみ（請求書は事務局が手動送信）
      sendReceiptOnlyEmail(data, receptNo);
    }

    // 事務局へも通知
    _notifyOffice(data, receptNo, autoSend);
  }

  return { result: 'success', receipt_no: receptNo };
}

/** カスタムメニュー登録 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 協賛管理')
    .addItem('お礼状送信',  'sendOreijouFromMenu')
    .addSeparator()
    .addItem('Info シート作成',     'setupInfoSheet')
    .addItem('申込みヘッダー作成',  'setupHeaders')
    .addItem('プロジェクト初期化',  'initProject')
    .addSeparator()
    .addItem('メールテンプレ保存',  'setupAllMailTemplates')
    .addToUi();
}

// ─── 内部ユーティリティ ────────────────────────────────────────────────────────

function _notifyOffice(data, receptNo, autoSent) {
  const subject = `【協賛申込】${data.company_name || ''} (${data.category || ''}) 受付番号:${receptNo}`;
  const body = [
    '新規協賛申込みがありました。',
    '',
    `会社名：${data.company_name || ''}`,
    `区分　：${data.category || ''} プラン`,
    `担当者：${data.staff_name || ''}`,
    `メール：${data.email || ''}`,
    `受付番号：${receptNo}`,
    `請求書：${autoSent ? '自動送信済み' : '手動送信待ち（S/Aプラン）'}`,
  ].join('\n');
  try {
    MailApp.sendEmail({ to: getOfficeEmail(), subject, body });
  } catch (e) {
    console.warn('事務局通知メール送信失敗:', e.message);
  }
}

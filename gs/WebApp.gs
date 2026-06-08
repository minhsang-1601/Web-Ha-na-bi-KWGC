// ─── Web App エントリーポイント ────────────────────────────────────────────────

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('【第5回川口花火大会】協賛申込み')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const data     = e.parameter;
    const receptNo = appendRow(data, data.sheetName || DEFAULT_SHEET_NAME);
    appendToTesagyouSheet(receptNo, data.sheetName2 || DEFAULT_SHEET_NAME2);
    if (data.email) {
      const pdf = generateInvoicePdf(data, receptNo);
      sendConfirmationEmail(data, receptNo, pdf);
    }
    return jsonResponse({ result: 'success' });
  } catch (err) {
    return jsonResponse({ result: 'error', message: err.message });
  }
}

/** HtmlService の <?!= include('FileName') ?> 用ヘルパー */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * クライアント(google.script.run)から呼ばれる設定取得。
 * @returns {{ startDate, endDate, sheetName1, sheetName2 }}
 */
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    startDate:  props.getProperty('START_DATE')  || '2025-01-01T00:00:00',
    endDate:    props.getProperty('END_DATE')     || '2026-10-01T23:59:59',
    sheetName1: props.getProperty('SHEET_NAME1') || DEFAULT_SHEET_NAME,
    sheetName2: props.getProperty('SHEET_NAME2') || DEFAULT_SHEET_NAME2,
  };
}

/**
 * クライアント(google.script.run)から呼ばれるフォーム送信。
 * @param {Object} data - フォームフィールドのプレーンオブジェクト
 */
function submitForm(data) {
  const config     = getConfig();
  const sheetName  = data.sheetName  || config.sheetName1;
  const sheetName2 = data.sheetName2 || config.sheetName2;
  const kubun      = (data.category || '').trim().toUpperCase();
  const autoSend   = AUTO_SEND_KUBUN.includes(kubun);

  const receptNo = appendRow(data, sheetName);
  appendToTesagyouSheet(receptNo, sheetName2, data);

  if (data.email) {
    // B〜E: 請求書PDF を添付して送信
    // S/A : 受付確認のみ（PDFなし）
    const pdf = autoSend ? generateInvoicePdf(data, receptNo) : null;
    sendConfirmationEmail(data, receptNo, pdf);
  }
  return { result: 'success', receipt_no: receptNo };
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

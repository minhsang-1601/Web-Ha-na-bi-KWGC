// ─── 申込確認メール ────────────────────────────────────────────────────────────

/** B〜E: 受付確認 + 請求書PDF 添付 */
function sendConfirmationEmail(data, receptNo, invoicePdf) {
  const props   = PropertiesService.getScriptProperties();
  let subject   = props.getProperty('MAIL_SUBJECT') ||
    `【${getEventName()}】申込受理書兼請求書のご連絡（受付番号：{{receipt_no}}）`;
  let body      = props.getProperty('MAIL_BODY') || _defaultConfirmBody();

  const vars = _buildVars(data, receptNo);
  subject = _replaceVars(subject, vars);
  body    = _replaceVars(body,    vars);

  const mailOptions = { to: data.email, subject, body, replyTo: getOfficeEmail() };
  if (invoicePdf) {
    mailOptions.attachments = [
      invoicePdf.setName(`申込受理書兼請求書_${data.company_name || receptNo}.pdf`)
    ];
  }
  MailApp.sendEmail(mailOptions);
}

/** S/A: 受付確認のみ（請求書は手動送信） */
function sendReceiptOnlyEmail(data, receptNo) {
  const props   = PropertiesService.getScriptProperties();
  let subject   = props.getProperty('RECEIPT_ONLY_SUBJECT') ||
    `【${getEventName()}】お申し込みを受け付けました（受付番号：{{receipt_no}}）`;
  let body      = props.getProperty('RECEIPT_ONLY_BODY') || _defaultReceiptOnlyBody();

  const vars = _buildVars(data, receptNo);
  subject = _replaceVars(subject, vars);
  body    = _replaceVars(body,    vars);

  MailApp.sendEmail({ to: data.email, subject, body, replyTo: getOfficeEmail() });
}

// ─── 請求書PDF生成 ─────────────────────────────────────────────────────────────

function generateInvoicePdf(data, receptNo) {
  const now       = new Date();
  const reiwa     = now.getFullYear() - 2018;
  const issueDate = `令和${reiwa}年${now.getMonth() + 1}月${now.getDate()}日`;

  const category   = (data.category || '').trim().toUpperCase();
  const totalPrice = getCategoryPrice(category);
  const subtotal   = Math.round(totalPrice / 1.1);
  const tax        = totalPrice - subtotal;
  const fmt        = n => n > 0 ? `¥${n.toLocaleString()}` : '―';

  let html = HtmlService.createHtmlOutputFromFile('invoice-template').getContent();
  html = html.replace('src="hanko.png"', `src="https://drive.google.com/uc?id=${getHankoFileId()}"`);

  const replacements = {
    '{{company_name}}': data.company_name || '',
    '{{issue_date}}':   issueDate,
    '{{receipt_no}}':   receptNo,
    '{{category}}':     category,
    '{{total}}':        fmt(totalPrice),
    '{{subtotal}}':     fmt(subtotal),
    '{{tax}}':          fmt(tax),
    '{{payment_due}}':  getPaymentDue(),
  };
  Object.entries(replacements).forEach(([k, v]) => { html = html.split(k).join(v); });

  const tmpFile = DriveApp.createFile(
    Utilities.newBlob(html, MimeType.HTML, `_tmp_${receptNo}.html`)
  );
  const pdf = tmpFile.getAs(MimeType.PDF);
  tmpFile.setTrashed(true);
  return pdf;
}

// ─── 案内文メール ──────────────────────────────────────────────────────────────

function sendAnnaibunEmail(data, receptNo) {
  const props   = PropertiesService.getScriptProperties();
  let subject   = props.getProperty('ANNAI_SUBJECT') ||
    `【${getEventName()}】ご案内（受付番号：{{receipt_no}}）`;
  let body      = props.getProperty('ANNAI_BODY') || _defaultAnnaiBody();

  const vars = _buildVars(data, receptNo);
  subject = _replaceVars(subject, vars);
  body    = _replaceVars(body,    vars);

  MailApp.sendEmail({ to: data.email, subject, body, replyTo: getOfficeEmail() });
}

// ─── お礼状メール ──────────────────────────────────────────────────────────────

function sendOreijouEmail(data, receptNo) {
  const props   = PropertiesService.getScriptProperties();
  let subject   = props.getProperty('OREIJOU_SUBJECT') ||
    `【${getEventName()}】ご協賛へのお礼（受付番号：{{receipt_no}}）`;
  let body      = props.getProperty('OREIJOU_BODY') || _defaultOreijouBody();

  const vars = _buildVars(data, receptNo);
  subject = _replaceVars(subject, vars);
  body    = _replaceVars(body,    vars);

  const pdf = generateOreijouPdf(data);
  const mailOptions = { to: data.email, subject, body, replyTo: getOfficeEmail() };
  if (pdf) mailOptions.attachments = [pdf.setName(`お礼状_${data.company_name || ''}.pdf`)];
  MailApp.sendEmail(mailOptions);
}

// ─── お礼状PDF生成 ────────────────────────────────────────────────────────────

function generateOreijouPdf(data) {
  try {
    const now       = new Date();
    const reiwa     = now.getFullYear() - 2018;
    const issueDate = `令和${reiwa}年${now.getMonth() + 1}月${now.getDate()}日`;

    let html = HtmlService.createHtmlOutputFromFile('oreijou-template').getContent();
    html = html.replace('src="hanko.png"', `src="https://drive.google.com/uc?id=${getHankoFileId()}"`);

    const replacements = {
      '{{company_name}}': data.company_name || '',
      '{{rep_name}}':     data.rep_name     || '',
      '{{issue_date}}':   issueDate,
    };
    Object.entries(replacements).forEach(([k, v]) => { html = html.split(k).join(v); });

    const tmpFile = DriveApp.createFile(
      Utilities.newBlob(html, MimeType.HTML, '_tmp_oreijou.html')
    );
    const pdf = tmpFile.getAs(MimeType.PDF);
    tmpFile.setTrashed(true);
    return pdf;
  } catch (err) {
    console.error('Oreijou PDF生成エラー:', err.message);
    return null;
  }
}

// ─── 内部ユーティリティ ────────────────────────────────────────────────────────

function _buildVars(data, receptNo) {
  return {
    company_name: data.company_name || '',
    rep_name:     data.rep_name     || '',
    staff_name:   data.staff_name   || '',
    category:     data.category     || '',
    receipt_no:   receptNo          || '',
    date:         nowStr(),
    event_name:   getEventName(),
    payment_due:  getPaymentDue(),
    office_email: getOfficeEmail(),
  };
}

function _replaceVars(str, vars) {
  Object.entries(vars).forEach(([key, val]) => {
    str = str.replace(new RegExp(`\\{\\{${key}\\}\\}`, 'g'), val);
  });
  return str;
}

function _defaultConfirmBody() {
  return [
    '{{company_name}}',
    '{{rep_name}} 様',
    '',
    '{{event_name}} 実行委員会 事務局でございます。',
    'このたびはご協賛のお申し込みをいただき、誠にありがとうございます。',
    '',
    '─────────────────────────────',
    '■ お申し込み内容',
    '　会社名・団体名　：{{company_name}}',
    '　ご担当者名　　　：{{staff_name}}',
    '　区分　　　　　　：{{category}} プラン',
    '　お申し込み日時　：{{date}}',
    '　受付番号　　　　：{{receipt_no}}',
    '─────────────────────────────',
    '',
    '本メールに「申込受理書兼請求書」をPDFにて添付しております。',
    'お振込期限（{{payment_due}}）までにお手続きくださいますようお願い申し上げます。',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
    '{{event_name}} 実行委員会 事務局',
    'E-mail：{{office_email}}',
    '受付時間：平日 10:00 〜 17:00',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
    '※ このメールは自動送信されています。',
  ].join('\n');
}

function _defaultReceiptOnlyBody() {
  return [
    '{{company_name}}',
    '{{rep_name}} 様',
    '',
    '{{event_name}} 実行委員会 事務局でございます。',
    'このたびはご協賛のお申し込みをいただき、誠にありがとうございます。',
    '',
    '─────────────────────────────',
    '■ お申し込み内容',
    '　会社名・団体名　：{{company_name}}',
    '　ご担当者名　　　：{{staff_name}}',
    '　区分　　　　　　：{{category}} プラン',
    '　お申し込み日時　：{{date}}',
    '　受付番号　　　　：{{receipt_no}}',
    '─────────────────────────────',
    '',
    '請求書は改めてご送付いたします。',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
    '{{event_name}} 実行委員会 事務局',
    'E-mail：{{office_email}}',
    '受付時間：平日 10:00 〜 17:00',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
    '※ このメールは自動送信されています。',
  ].join('\n');
}

function _defaultAnnaiBody() {
  return [
    '{{company_name}}',
    '{{rep_name}} 様',
    '',
    '{{event_name}} 実行委員会 事務局でございます。',
    'このたびはご協賛いただき、誠にありがとうございます。',
    '',
    '当日のご案内をお送りいたします。',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
    '{{event_name}} 実行委員会 事務局',
    'E-mail：{{office_email}}',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
  ].join('\n');
}

function _defaultOreijouBody() {
  return [
    '{{company_name}}',
    '{{rep_name}} 様',
    '',
    '{{event_name}} 実行委員会 事務局でございます。',
    'このたびはご協賛ならびにご入金いただき、誠にありがとうございます。',
    '',
    'なお、お礼状をPDFにて添付しておりますのでご確認ください。',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
    '{{event_name}} 実行委員会 事務局',
    'E-mail：{{office_email}}',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
    '※ このメールは自動送信されています。',
  ].join('\n');
}

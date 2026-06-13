// ─── メール送信残数の確認（GAS エディタから実行） ────────────────────────────────

/** 本日あと何通メール送信できるかを確認する */
function checkMailQuota() {
  const remaining = MailApp.getRemainingDailyQuota();
  const msg = `📧 本日のメール送信 残り回数: ${remaining} 通`;
  console.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (_) {}
  return remaining;
}

// ─── 申込確認メール ────────────────────────────────────────────────────────────

/** B〜E: 受付確認 + 請求書PDF 添付 */
function sendConfirmationEmail(data, receptNo, invoicePdf) {
  const props   = PropertiesService.getScriptProperties();
  let subject   = props.getProperty('MAIL_SUBJECT') ||
    `【${getEventName()}】協賛お申込みを受け付けました。`;
  let body      = props.getProperty('MAIL_BODY') || _defaultConfirmBody();

  const vars = _buildVars(data, receptNo);
  subject = _replaceVars(subject, vars);
  body    = _replaceVars(body,    vars);

  const mailOptions = { to: data.email, cc: getOfficeEmail(), subject, body, replyTo: getOfficeEmail() };
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
    `【${getEventName()}】協賛お申込みを受け付けました。`;
  let body      = props.getProperty('RECEIPT_ONLY_BODY') || _defaultReceiptOnlyBody();

  const vars = _buildVars(data, receptNo);
  subject = _replaceVars(subject, vars);
  body    = _replaceVars(body,    vars);

  MailApp.sendEmail({ to: data.email, cc: getOfficeEmail(), subject, body, replyTo: getOfficeEmail() });
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
    '{{company_name}}':  data.company_name || '',
    '{{issue_date}}':    issueDate,
    '{{receipt_no}}':    receptNo,
    '{{category}}':      category,
    '{{total}}':         fmt(totalPrice),
    '{{subtotal}}':      fmt(subtotal),
    '{{tax}}':           fmt(tax),
    '{{payment_due}}':   getPaymentDue(),
    '{{event_name}}':    getEventName(),
    '{{org_name}}':      getOrgName(),
    '{{org_rep}}':       getOrgRep(),
    '{{invoice_reg_no}}':getInvoiceRegNo(),
    '{{bank_name}}':     getBankName(),
    '{{bank_no}}':       getBankNo(),
    '{{bank_holder}}':   getBankHolder(),
    '{{bank_rep}}':      getBankRep(),
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

  MailApp.sendEmail({ to: data.email, cc: getOfficeEmail(), subject, body, replyTo: getOfficeEmail() });
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
  const mailOptions = { to: data.email, cc: getOfficeEmail(), subject, body, replyTo: getOfficeEmail() };
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
      '{{company_name}}':  data.company_name || '',
      '{{rep_name}}':      data.rep_name     || '',
      '{{issue_date}}':    issueDate,
      '{{event_name}}':    getEventName(),
      '{{org_name}}':      getOrgName(),
      '{{org_rep}}':       getOrgRep(),
      '{{org_location}}':  getOrgLocation(),
      '{{org_tel}}':       getOrgTel(),
      '{{org_fax}}':       getOrgFax(),
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
    office_hours: getOfficeHours(),
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
    '{{staff_name}} 様',
    '',
    '川口花火大会実行委員会でございます。',
    'このたびは、協賛にお申し込みいただき、誠にありがとうございます。',
    '本メールに「申込受理書兼請求書」をPDFにて添付しております。',
    'お振込み期限（{{payment_due}}）までに',
    'お手続きくださいますようお願い申しあげます。',
    '',
    '■お申込み内容',
    '　・会社名・団体名　：{{company_name}}',
    '　・ご担当者名　　　：{{staff_name}}',
    '　・区分　　　　　　：{{category}}',
    '　・お申込み日時　　：{{date}}',
    '　・受付番号　　　　：{{receipt_no}}',
    '',
    'ーーーーーーーーーーーーーーーーーーーーーーーーーー',
    'メールアドレス：{{office_email}}',
    'ーーーーーーーーーーーーーーーーーーーーーーーーーー',
  ].join('\n');
}

function _defaultReceiptOnlyBody() {
  return [
    '{{company_name}}',
    '{{staff_name}} 様',
    '',
    '川口花火大会実行委員会でございます。',
    'このたびは、協賛にお申し込みいただき、誠にありがとうございます。',
    '',
    'S・A協賛につきましては、募集枠を超えるお申し込みがあった場合、締切後に抽選を実施いたします。',
    '協賛の可否につきましては、締切後に改めてメールにてご連絡いたします。',
    '限られた募集枠となり恐縮ですが、何卒ご理解・ご協力のほどよろしくお願い申しあげます。',
    'ご不明な点がございましたら、お気軽にお問い合わせください。',
    'よろしくお願い申しあげます。',
    '',
    '■ お申込み内容',
    '　・会社名・団体名　：{{company_name}}',
    '　・ご担当者名　　　：{{staff_name}}',
    '　・区分　　　　　　：{{category}}',
    '　・お申込み日時　　：{{date}}',
    '　・受付番号　　　　：{{receipt_no}}',
    '',
    'ーーーーーーーーーーーーーーーーーーーーーーーーーー',
    'メールアドレス：{{office_email}}',
    'ーーーーーーーーーーーーーーーーーーーーーーーーーー',
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

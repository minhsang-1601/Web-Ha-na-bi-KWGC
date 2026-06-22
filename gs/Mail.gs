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
  let body      = props.getProperty('MAIL_BODY') || defaultConfirmBody();

  const vars = _buildVars(data, receptNo);
  subject = _replaceVars(subject, vars);
  body    = _replaceVars(body,    vars);

  const officeEmail = getOfficeEmail();
  const mailOptions = { to: data.email, subject, htmlBody: body };
  if (_validEmail(officeEmail)) { mailOptions.cc = officeEmail; mailOptions.replyTo = officeEmail; }
  if (invoicePdf) {
    mailOptions.attachments = [
      invoicePdf.setName(`申込受理書兼請求書_${data.company_name || receptNo}.pdf`)
    ];
  }
  MailApp.sendEmail(mailOptions);
}

/** S/A: 受付確認のみ（請求書は手動送信、またはPDF添付可能） */
function sendReceiptOnlyEmail(data, receptNo, invoicePdf) {
  const props   = PropertiesService.getScriptProperties();
  let subject   = props.getProperty('RECEIPT_ONLY_SUBJECT') ||
    `【${getEventName()}】協賛お申込みを受け付けました。`;
  let body      = props.getProperty('RECEIPT_ONLY_BODY') || defaultReceiptOnlyBody();

  const vars = _buildVars(data, receptNo);
  subject = _replaceVars(subject, vars);
  body    = _replaceVars(body,    vars);

  const officeEmail = getOfficeEmail();
  const mailOptions = { to: data.email, subject, htmlBody: body };
  if (_validEmail(officeEmail)) { mailOptions.cc = officeEmail; mailOptions.replyTo = officeEmail; }
  if (invoicePdf) {
    mailOptions.attachments = [
      invoicePdf.setName(`申込受理書兼請求書_${data.company_name || receptNo}.pdf`)
    ];
  }
  MailApp.sendEmail(mailOptions);
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
  // ── 印影画像をBase64埋め込みに変換（PDF生成時に外部URLは読み込めないため） ──
  try {
    const hankoBlob  = DriveApp.getFileById(getHankoFileId()).getBlob();
    const hankoB64   = Utilities.base64Encode(hankoBlob.getBytes());
    const hankoMime  = hankoBlob.getContentType() || 'image/png';
    html = html.replace('src="hanko.png"', `src="data:${hankoMime};base64,${hankoB64}"`);
  } catch (e) {
    console.warn('印影画像の読み込み失敗:', e.message);
    html = html.replace('src="hanko.png"', 'src=""');
  }

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

// ─── S/A 抽選メール（column I チェック時） ────────────────────────────────────

/** S/A: 抽選確定・請求書送付メール（column I チェック時に送信） */
// ─── S/A 当選通知 + 請求書メール（column I チェック時） ─────────────────────────

/** S/A: 抽選確定・請求書送付メール（column I チェック時に送信） */
function sendSaInvoiceEmail(data, receptNo, invoicePdf) {
  const props   = PropertiesService.getScriptProperties();
  let subject   = props.getProperty('SA_INVOICE_SUBJECT') ||
    `【{{event_name}}】協賛金のご請求書送付のご案内`;
  let body      = props.getProperty('SA_INVOICE_BODY') || defaultSaInvoiceBody();

  const vars = _buildSaInvoiceVars(data, receptNo);
  subject = _replaceVars(subject, vars);
  body    = _replaceVars(body,    vars);

  const officeEmail = getOfficeEmail();
  const mailOptions = { to: data.email, subject, htmlBody: body };
  if (_validEmail(officeEmail)) { mailOptions.cc = officeEmail; mailOptions.replyTo = officeEmail; }
  if (invoicePdf) {
    mailOptions.attachments = [
      invoicePdf.setName(`申込受理書兼請求書_${data.company_name || receptNo}.pdf`)
    ];
  }
  MailApp.sendEmail(mailOptions);
}

function _buildSaInvoiceVars(data, receptNo) {
  const category   = String(data.category || '').trim().toUpperCase();
  const totalPrice = getCategoryPrice(category);
  return {
    company_name: data.company_name || '',
    staff_name:   data.staff_name   || '',
    category:     category,
    receipt_no:   receptNo          || '',
    event_name:   getEventName(),
    payment_due:  getPaymentDue(),
    office_email: getOfficeEmail(),
    amount:       totalPrice > 0 ? totalPrice.toLocaleString() : '―',
  };
}

function defaultSaInvoiceBody() {
  return `<div style="font-family:'Meiryo',sans-serif;font-size:14px;line-height:1.9;">
{{company_name}}<br>
{{staff_name}} 様<br>
<br>
川口花火大会実行委員会でございます。<br>
<br>
このたびは、{{event_name}}の協賛にお申し込みいただき、誠にありがとうございます。<br>
厳正なる選考（抽選）の結果、このたび貴社（貴団体）の協賛が確定いたしましたので、ご連絡申しあげます。<br>
<br>
本メールに「申込受理書兼請求書」をPDFにて添付しております。<br>
お振込み期限<strong>【{{payment_due}}】</strong>までにお手続きくださいますようお願い申しあげます。<br>
<br>
■ご請求内容<br>
　・会社名・団体名　：{{company_name}}<br>
　・ご担当者名　　　：{{staff_name}}<br>
　・区分　　　　　　：{{category}}<br>
　・受付番号　　　　：{{receipt_no}}<br>
　・協賛金額　　　　：{{amount}}円（税込）<br>
<br>
ご不明な点がございましたら、お気軽にお問い合わせください。<br>
何卒よろしくお願い申しあげます。<br>
<br>
ーーーーーーーーーーーーーーーーーーーーーーーーーー<br>
メールアドレス：{{office_email}}<br>
ーーーーーーーーーーーーーーーーーーーーーーーーーー
</div>`;
}

// ─── お礼状メール ──────────────────────────────────────────────────────────────

function sendOreijouEmail(data, receptNo) {
  const props   = PropertiesService.getScriptProperties();
  let subject   = props.getProperty('OREIJOU_SUBJECT') ||
    `【${getEventName()}】ご協賛へのお礼`;
  let body      = props.getProperty('OREIJOU_BODY') || defaultOreijouBody();

  const vars = _buildVars(data, receptNo);
  subject = _replaceVars(subject, vars);
  body    = _replaceVars(body,    vars);

  const pdf = generateOreijouPdf(data);
  const officeEmail3 = getOfficeEmail();
  const mailOptions = { to: data.email, subject, htmlBody: body };
  if (_validEmail(officeEmail3)) { mailOptions.cc = officeEmail3; mailOptions.replyTo = officeEmail3; }
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
    try {
      const hankoBlob  = DriveApp.getFileById(getHankoFileId()).getBlob();
      const hankoB64   = Utilities.base64Encode(hankoBlob.getBytes());
      const hankoMime  = hankoBlob.getContentType() || 'image/png';
      html = html.replace('src="hanko.png"', `src="data:${hankoMime};base64,${hankoB64}"`);
    } catch (e) {
      console.warn('印影画像の読み込み失敗:', e.message);
      html = html.replace('src="hanko.png"', 'src=""');
    }

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

/** メールアドレスとして有効な文字列かチェック（プレースホルダーは無効） */
function _validEmail(email) {
  const s = String(email || '').trim();
  return s.includes('@') && !s.includes('Default');
}

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

function defaultConfirmBody() {
  return `<div style="font-family:'Meiryo',sans-serif;font-size:14px;line-height:1.9;">
{{company_name}}<br>
{{staff_name}} 様<br>
<br>
川口花火大会実行委員会でございます。<br>
このたびは、協賛にお申し込みいただき、誠にありがとうございます。<br>
本メールに「申込受理書兼請求書」をPDFにて添付しております。<br>
お振込み期限<strong>【{{payment_due}}】</strong>までにお手続きくださいますようお願い申しあげます。<br>
<br>
■お申込み内容<br>
　・会社名・団体名　：{{company_name}}<br>
　・ご担当者名　　　：{{staff_name}}<br>
　・区分　　　　　　：{{category}}<br>
　・お申込み日時　　：{{date}}<br>
　・受付番号　　　　：{{receipt_no}}<br>
<br>
ーーーーーーーーーーーーーーーーーーーーーーーーーー<br>
メールアドレス：{{office_email}}<br>
ーーーーーーーーーーーーーーーーーーーーーーーーーー
</div>`;
}

function defaultReceiptOnlyBody() {
  return `<div style="font-family:'Meiryo',sans-serif;font-size:14px;line-height:1.9;">
{{company_name}}<br>
{{staff_name}} 様<br>
<br>
川口花火大会実行委員会でございます。<br>
このたびは、協賛にお申し込みいただき、誠にありがとうございます。<br>
<br>
S・A協賛につきましては、募集枠を超えるお申し込みがあった場合、締切後に抽選を実施いたします。<br>
協賛の可否につきましては、締切後に改めてメールにてご連絡いたします。<br>
限られた募集枠となり恐縮ですが、何卒ご理解・ご協力のほどよろしくお願い申しあげます。<br>
ご不明な点がございましたら、お気軽にお問い合わせください。<br>
よろしくお願い申しあげます。<br>
<br>
■ お申込み内容<br>
　・会社名・団体名　：{{company_name}}<br>
　・ご担当者名　　　：{{staff_name}}<br>
　・区分　　　　　　：{{category}}<br>
　・お申込み日時　　：{{date}}<br>
　・受付番号　　　　：{{receipt_no}}<br>
<br>
ーーーーーーーーーーーーーーーーーーーーーーーーーー<br>
メールアドレス：{{office_email}}<br>
ーーーーーーーーーーーーーーーーーーーーーーーーーー
</div>`;
}

function defaultOreijouBody() {
  return `<div style="font-family:'Meiryo',sans-serif;font-size:14px;line-height:1.9;">
{{company_name}}<br>
{{rep_name}} 様<br>
<br>
{{event_name}} 実行委員会 事務局でございます。<br>
このたびはご協賛ならびにご入金いただき、誠にありがとうございます。<br>
<br>
なお、お礼状をPDFにて添付しておりますのでご確認ください。<br>
<br>
━━━━━━━━━━━━━━━━━━━━━━━━<br>
{{event_name}} 実行委員会 事務局<br>
E-mail：{{office_email}}<br>
━━━━━━━━━━━━━━━━━━━━━━━━<br>
※ このメールは自動送信されています。
</div>`;
}

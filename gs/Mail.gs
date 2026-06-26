// ─── メール送信残数の確認（GAS エディタから実行） ────────────────────────────────

/** 本日あと何通メール送信できるかを確認する */
function checkMailQuota() {
  const remaining = MailApp.getRemainingDailyQuota();
  const msg = `📧 本日のメール送信 残り回数: ${remaining} 通`;
  console.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (_) {}
  return remaining;
}

// ─── 申込み確認メール ────────────────────────────────────────────────────────────

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
    '{{title}}':         (category === 'S' || category === 'A') ? '請求書' : '申込受理書兼請求書', 
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
      invoicePdf.setName(`請求書_${data.company_name || receptNo}.pdf`)
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
    office_hours: getOfficeHours(),
    amount:       totalPrice > 0 ? Math.round(totalPrice / 10000) + '万' : '―',
    // 組織・振込先情報
    org_name:       getOrgName(),
    org_rep:        getOrgRep(),
    org_location:   getOrgLocation(),
    org_tel:        getOrgTel(),
    org_fax:        getOrgFax(),
    invoice_reg_no: getInvoiceRegNo(),
    bank_name:      getBankName(),
    bank_no:        getBankNo(),
    bank_holder:    getBankHolder(),
    bank_rep:       getBankRep(),
    // 区分別 申込み期間（テンプレートで使用）
    KUBUN_SA_END:     _fmtDateJa(getKubunSaEnd()),
    KUBUN_SA_START:   _fmtDateJa(getKubunSaStart()),
    KUBUN_BCDE_END:   _fmtDateJa(getKubunBcdeEnd()),
    KUBUN_BCDE_START: _fmtDateJa(getKubunBcdeStart()),
  };
}

/** ISO日時文字列を「yyyy年M月d日」に整形（変換できなければ元の文字列） */
function _fmtDateJa(iso) {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return String(iso || '');
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy年M月d日');
}

function defaultSaInvoiceBody() {
  return `<div style="font-family:'Meiryo',sans-serif;font-size:14px;line-height:1.9;">
{{company_name}}<br>
{{staff_name}} 様<br>
<br>
{{org_name}}でございます。<br>
<br>
このたびは、{{event_name}}の協賛にお申込みいただき、誠にありがとうございます。<br>
<strong>{{KUBUN_SA_END}}</strong>の申込み締め切りをもちまして、{{company_name}}様の{{category}}協賛枠が確定いたしましたので、ご連絡申しあげます。<br>
<br>
本メールに「請求書」をPDFにて添付しております。<br>
お振込み期限<strong>【{{payment_due}}】</strong>までにお手続きくださいますようお願い申しあげます。<br>
<br>
■ご請求内容<br>
　・会社名・団体名　：{{company_name}}<br>
　・担当者名　　　　：{{staff_name}}<br>
　・区分　　　　　　：{{category}}<br>
　・協賛金額　　　　：{{amount}}円（税込）<br>
　・受付番号　　　　：{{receipt_no}}<br>
<br>
ご不明な点がございましたら、お気軽にお問い合わせください。<br>
何卒よろしくお願い申しあげます。<br>
<br>
ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー<br>
※このメールは自動送信メールです。本メールへの返信はできませんのでご了承ください。<br>
なお、お心当たりのない場合やお申込み内容に誤りがある場合は、下記までご連絡ください。<br>
【お問い合わせ先】<br>
{{org_name}}{{org_location}}<br>
{{office_email}}<br>
ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー
</div>`;
}

// ─── お礼状メール ──────────────────────────────────────────────────────────────

function sendOreijouEmail(data, receptNo) {
  const props   = PropertiesService.getScriptProperties();
  let subject   = props.getProperty('OREIJOU_SUBJECT') ||
    `【${getEventName()}】ご協賛のお礼`;
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
  const cat        = String(data.category || '').trim().toUpperCase();
  const totalPrice = getCategoryPrice(cat);
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
    amount:       totalPrice > 0 ? Math.round(totalPrice / 10000) + '万' : '―',
    // 組織・振込先情報（テンプレートで使用される場合に置換）
    org_name:       getOrgName(),
    org_rep:        getOrgRep(),
    org_location:   getOrgLocation(),
    org_tel:        getOrgTel(),
    org_fax:        getOrgFax(),
    invoice_reg_no: getInvoiceRegNo(),
    bank_name:      getBankName(),
    bank_no:        getBankNo(),
    bank_holder:    getBankHolder(),
    bank_rep:       getBankRep(),
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
{{org_name}}でございます。<br>
このたびは、{{event_name}}の協賛にお申込みいただき、誠にありがとうございます。<br>
本メールに「申込受理書兼請求書」をPDFにて添付しております。<br>
お振込み期限<strong>【{{payment_due}}】</strong>までにお手続きくださいますようお願い申しあげます。<br>
<br>
■お申込み内容<br>
　・会社名・団体名　：{{company_name}}<br>
　・担当者名　　　　：{{staff_name}}<br>
　・区分　　　　　　：{{category}}<br>
　・協賛金額　　　　：{{amount}}円（税込）<br>
　・お申込み日時　　：{{date}}<br>
　・受付番号　　　　：{{receipt_no}}<br>
<br>
ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー<br>
※このメールは自動送信メールです。本メールへの返信はできませんのでご了承ください。<br>
なお、お心当たりのない場合やお申込み内容に誤りがある場合は、下記までご連絡ください。<br>
【お問い合わせ先】<br>
{{org_name}}{{org_location}}<br>
{{office_email}}<br>
ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー
</div>`;
}

function defaultReceiptOnlyBody() {
  return `<div style="font-family:'Meiryo',sans-serif;font-size:14px;line-height:1.9;">
{{company_name}}<br>
{{staff_name}} 様<br>
<br>
{{org_name}}でございます。<br>
このたびは、協賛にお申込みいただき、誠にありがとうございます。<br>
<br>
S・A協賛につきましては、募集枠を超えるお申込みがあった場合、締切後に抽選を実施いたします。<br>
協賛の可否につきましては、締切後に改めてメールにてご連絡いたします。<br>
限られた募集枠となり恐縮ですが、何卒ご理解・ご協力のほどよろしくお願い申しあげます。<br>
ご不明な点がございましたら、お気軽にお問い合わせください。<br>
よろしくお願い申しあげます。<br>
<br>
■ お申込み内容<br>
　・会社名・団体名　：{{company_name}}<br>
　・担当者名　　　　：{{staff_name}}<br>
　・区分　　　　　　：{{category}}<br>
　・協賛金額　　　　：{{amount}}円（税込）<br>
　・お申込み日時　　：{{date}}<br>
　・受付番号　　　　：{{receipt_no}}<br>
<br>
ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー<br>
※このメールは自動送信メールです。本メールへの返信はできませんのでご了承ください。<br>
なお、お心当たりのない場合やお申込み内容に誤りがある場合は、下記までご連絡ください。<br>
【お問い合わせ先】<br>
{{org_name}}{{org_location}}<br>
{{office_email}}<br>
ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー
</div>`;
}

function defaultOreijouBody() {
  return `<div style="font-family:'Meiryo',sans-serif;font-size:14px;line-height:1.9;">
{{company_name}}<br>
{{staff_name}} 様<br>
<br>
{{org_name}}でございます。<br>
ご協賛金のご入金を確認いたしましたので、ご連絡申しあげます。<br>
皆様のあたたかいご支援は、大会の開催に向けた大きな力となっております。<br>
<br>
心より厚く御礼申しあげます。<br>
なお、協賛企業決定通知書をPDFファイルにて添付いたしましたので、ご確認ください。<br>
<br>
今後とも、川口花火大会へのご支援・ご協力を賜りますよう、<br>
何卒よろしくお願い申しあげます。<br>
<br>
</div>`;
}

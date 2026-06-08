// ─── 申込確認メール（請求書PDF添付） ──────────────────────────────────────────

function sendConfirmationEmail(data, receptNo, invoicePdf) {
  const now  = new Date();
  const date = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  const vars = {
    company_name: data.company_name || '',
    rep_name:     data.rep_name     || '',
    staff_name:   data.staff_name   || '',
    category:     data.category     || '',
    date:         date,
    receipt_no:   receptNo          || '',
  };

  const props       = PropertiesService.getScriptProperties();
  let subject       = props.getProperty('MAIL_SUBJECT') || '【第5回川口花火大会】協賛お申し込み受付のご連絡';
  let body          = props.getProperty('MAIL_BODY')    || '{{company_name}}\n{{rep_name}} 様\n\nお申し込みを受け付けました。\n\n川口花火大会実行委員会 事務局';

  Object.entries(vars).forEach(([key, value]) => {
    const re = new RegExp(`\\{\\{${key}\\}\\}`, 'g');
    subject  = subject.replace(re, value);
    body     = body.replace(re, value);
  });

  const mailOptions = { to: data.email, subject, body, replyTo: OFFICE_EMAIL };
  if (invoicePdf) {
    mailOptions.attachments = [
      invoicePdf.setName(`申込受理書兼請求書_${data.company_name || receptNo}.pdf`)
    ];
  }

  MailApp.sendEmail(mailOptions);
}

function generateInvoicePdf(data, receptNo) {
  const now       = new Date();
  const reiwa     = now.getFullYear() - 2018;
  const issueDate = `令和${reiwa}年${now.getMonth() + 1}月${now.getDate()}日`;

  const category   = (data.category || '').trim().toUpperCase();
  const totalPrice = CATEGORY_PRICE[category] || 0;
  const subtotal   = Math.round(totalPrice / 1.1);
  const tax        = totalPrice - subtotal;
  const fmt        = (n) => n > 0 ? `¥${n.toLocaleString()}` : '';

  const marks = { A: '', B: '', C: '', D: '', E: '' };
  if (marks[category] !== undefined) marks[category] = '1';

  const amountLine = (cat) => cat === category ? fmt(totalPrice) : '';

  let html = HtmlService.createHtmlOutputFromFile('InvoiceTemplate').getContent();
  const hankoUrl = `https://drive.google.com/uc?id=${HANKO_FILE_ID}`;
  html = html.replace('src="hanko.png"', `src="${hankoUrl}"`);

  const replacements = {
    '{{company_name}}': data.company_name || '',
    '{{issue_date}}':   issueDate,
    '{{receipt_no}}':   receptNo,
    '{{total}}':        fmt(totalPrice),
    '{{subtotal}}':     fmt(subtotal),
    '{{tax}}':          fmt(tax),
    '{{mark_a}}': marks['A'], '{{mark_b}}': marks['B'],
    '{{mark_c}}': marks['C'], '{{mark_d}}': marks['D'], '{{mark_e}}': marks['E'],
    '{{amount_a}}': amountLine('A'), '{{amount_b}}': amountLine('B'),
    '{{amount_c}}': amountLine('C'), '{{amount_d}}': amountLine('D'), '{{amount_e}}': amountLine('E'),
  };

  Object.entries(replacements).forEach(([key, val]) => {
    html = html.split(key).join(val);
  });

  const tmpFile = DriveApp.createFile(Utilities.newBlob(html, MimeType.HTML, `_tmp_${receptNo}.html`));
  const pdf     = tmpFile.getAs(MimeType.PDF);
  tmpFile.setTrashed(true);
  return pdf;
}

// ─── お礼状メール ──────────────────────────────────────────────────────────────

function sendOreijouEmail(data, receptNo) {
  const props = PropertiesService.getScriptProperties();
  let subject = props.getProperty('OREIJOU_SUBJECT') || '【第5回川口花火大会】ご協賛へのお礼（受付番号：{{receipt_no}}）';
  let body    = props.getProperty('OREIJOU_BODY')    || 'お礼申し上げます。';

  const vars = {
    company_name: data.company_name || '',
    rep_name:     data.rep_name     || '',
    receipt_no:   receptNo          || '',
  };

  Object.entries(vars).forEach(([key, val]) => {
    const re = new RegExp(`\\{\\{${key}\\}\\}`, 'g');
    subject  = subject.replace(re, val);
    body     = body.replace(re, val);
  });

  const pdf = generateOreijouPdf(data);
  const mailOptions = { to: data.email, subject, body, replyTo: OFFICE_EMAIL };
  if (pdf) {
    mailOptions.attachments = [pdf.setName(`お礼状_${data.company_name || ''}.pdf`)];
  }

  MailApp.sendEmail(mailOptions);
}

function generateOreijouPdf(data) {
  try {
    const now       = new Date();
    const reiwa     = now.getFullYear() - 2018;
    const issueDate = `令和${reiwa}年${now.getMonth() + 1}月${now.getDate()}日`;

    let html = HtmlService.createHtmlOutputFromFile('oreijou-template').getContent();
    const hankoUrl = `https://drive.google.com/uc?id=${HANKO_FILE_ID}`;
    html = html.replace('src="hanko.png"', `src="${hankoUrl}"`);

    const replacements = {
      '{{company_name}}': data.company_name || '',
      '{{rep_name}}':     data.rep_name     || '',
      '{{issue_date}}':   issueDate,
    };
    Object.entries(replacements).forEach(([key, val]) => {
      html = html.split(key).join(val);
    });

    const tmpFile = DriveApp.createFile(Utilities.newBlob(html, MimeType.HTML, '_tmp_oreijou.html'));
    const pdf     = tmpFile.getAs(MimeType.PDF);
    tmpFile.setTrashed(true);
    return pdf;
  } catch (err) {
    console.error('Oreijou PDF生成エラー:', err.message);
    return null;
  }
}

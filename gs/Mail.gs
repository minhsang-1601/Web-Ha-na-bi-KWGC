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

  let html = HtmlService.createHtmlOutputFromFile('invoice-template').getContent();
  const hankoUrl = `https://drive.google.com/uc?id=${HANKO_FILE_ID}`;
  html = html.replace('src="hanko.png"', `src="${hankoUrl}"`);

  const replacements = {
    '{{company_name}}': data.company_name || '',
    '{{issue_date}}':   issueDate,
    '{{receipt_no}}':   receptNo,
    '{{category}}':     category,
    '{{total}}':        fmt(totalPrice),
    '{{subtotal}}':     fmt(subtotal),
    '{{tax}}':          fmt(tax),
  };

  Object.entries(replacements).forEach(([key, val]) => {
    html = html.split(key).join(val);
  });

  const tmpFile = DriveApp.createFile(Utilities.newBlob(html, MimeType.HTML, `_tmp_${receptNo}.html`));
  const pdf     = tmpFile.getAs(MimeType.PDF);
  tmpFile.setTrashed(true);
  return pdf;
}

// ─── 受付確認のみメール（S/A 申込時、請求書なし） ────────────────────────────────

function sendReceiptOnlyEmail(data, receptNo) {
  const props = PropertiesService.getScriptProperties();
  let subject = props.getProperty('RECEIPT_ONLY_SUBJECT') ||
    '【第5回川口花火大会】お申し込みを受け付けました（受付番号：{{receipt_no}}）';
  let body = props.getProperty('RECEIPT_ONLY_BODY') || [
    '{{company_name}}',
    '{{rep_name}} 様',
    '',
    '川口花火大会実行委員会 事務局でございます。',
    'このたびは第5回川口花火大会へのご協賛をお申し込みいただき、誠にありがとうございます。',
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
    'ご不明な点がございましたら、下記までご連絡ください。',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
    '第5回川口花火大会 実行委員会 事務局',
    'E-mail：' + OFFICE_EMAIL,
    '受付時間：平日 10:00 〜 17:00',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
  ].join('\n');

  const vars = {
    company_name: data.company_name || '',
    rep_name:     data.rep_name     || '',
    staff_name:   data.staff_name   || '',
    category:     data.category     || '',
    date:         Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
    receipt_no:   receptNo          || '',
  };
  Object.entries(vars).forEach(([key, val]) => {
    const re = new RegExp(`\\{\\{${key}\\}\\}`, 'g');
    subject  = subject.replace(re, val);
    body     = body.replace(re, val);
  });

  MailApp.sendEmail({ to: data.email, subject, body, replyTo: OFFICE_EMAIL });
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

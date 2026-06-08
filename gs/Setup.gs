// ─── 初回セットアップ用関数（エディタから手動実行） ────────────────────────────

/** 受付期間・シート名を Script Properties に保存する */
function setupWebAppConfig() {
  PropertiesService.getScriptProperties().setProperties({
    START_DATE:  '2025-01-01T00:00:00',
    END_DATE:    '2026-10-01T23:59:59',
    SHEET_NAME1: DEFAULT_SHEET_NAME,
    SHEET_NAME2: DEFAULT_SHEET_NAME2,
  });
  SpreadsheetApp.getUi().alert(
    'Web App 設定を保存しました。\n' +
    '変更: プロジェクトの設定 → スクリプト プロパティ'
  );
}

/** 申込確認メールのテンプレートを Script Properties に保存する */
function setupMailTemplate() {
  const subject = '【第5回川口花火大会】申込受理書兼請求書のご連絡（受付番号：{{receipt_no}}）';
  const body = [
    '{{company_name}}',
    '{{rep_name}} 様',
    '',
    '平素より格別のご高配を賜り、厚く御礼申し上げます。',
    '川口花火大会実行委員会 事務局でございます。',
    '',
    'このたびは、第5回川口花火大会へのご協賛をお申し込みいただき、',
    '誠にありがとうございます。',
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
    'ご確認のうえ、お振込期限までにお手続きくださいますようお願い申し上げます。',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
    '第5回川口花火大会 実行委員会 事務局',
    'E-mail：' + OFFICE_EMAIL,
    '受付時間：平日 10:00 〜 17:00',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
    '※ このメールは自動送信されています。',
  ].join('\n');

  PropertiesService.getScriptProperties().setProperties({
    MAIL_SUBJECT: subject,
    MAIL_BODY:    body,
    PAYMENT_DUE:  '9月11日（金）',
  });
  SpreadsheetApp.getUi().alert('申込確認メールテンプレートを保存しました。');
}

/** お礼状メールのテンプレートを Script Properties に保存する */
function setupOreijouTemplate() {
  const subject = '【第5回川口花火大会】ご協賛へのお礼（受付番号：{{receipt_no}}）';
  const body = [
    '{{company_name}}',
    '{{rep_name}} 様',
    '',
    '平素より格別のご高配を賜り、厚く御礼申し上げます。',
    '川口花火大会実行委員会 事務局でございます。',
    '',
    'このたびは、第5回川口花火大会へのご協賛ならびにご入金いただきまして、',
    '誠にありがとうございます。',
    '',
    'なお、お礼状をPDFにて添付しておりますのでご確認ください。',
    '',
    '当日は皆様に素晴らしい花火をお届けできるよう、実行委員会一同、精一杯努めてまいります。',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
    '第5回川口花火大会 実行委員会 事務局',
    'E-mail：' + OFFICE_EMAIL,
    '受付時間：平日 10:00 〜 17:00',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
    '※ このメールは自動送信されています。',
  ].join('\n');

  PropertiesService.getScriptProperties().setProperties({
    OREIJOU_SUBJECT: subject,
    OREIJOU_BODY:    body,
  });
  SpreadsheetApp.getUi().alert('お礼状テンプレートを保存しました。');
}

/** 申込みシートのヘッダー行を作成する（初回のみ） */
function setupHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DEFAULT_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert(`シート "${DEFAULT_SHEET_NAME}" が見つかりません。`); return; }
  if (sheet.getLastRow() > 0) { SpreadsheetApp.getUi().alert('ヘッダー行はすでに存在します。'); return; }
  sheet.appendRow(HEADERS);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#d0e4f7');
  sheet.getRange(2, 9,  sheet.getMaxRows() - 1).setNumberFormat('@');
  sheet.getRange(2, 11, sheet.getMaxRows() - 1).setNumberFormat('@');
  applyColumnWidths(sheet);
  SpreadsheetApp.getUi().alert('ヘッダー行を作成しました。');
}

/** 手作業シートの既存行のチェックボックス・J列を修正する */
function fixTetsugyoSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DEFAULT_SHEET_NAME2);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.getRange(2, 9, lastRow - 1).insertCheckboxes();
  for (let r = 2; r <= lastRow; r++) {
    const cell = sheet.getRange(r, 10);
    if (cell.getValue() === '' || cell.getValue() === null) cell.setValue('未');
  }
  SpreadsheetApp.getUi().alert('✅ チェックボックスとJ列を修正しました。');
}

/** 手作業シートを保護する（オーナーのみ編集可） */
function setupSheetProtection() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DEFAULT_SHEET_NAME2);
  if (!sheet) { SpreadsheetApp.getUi().alert('手作業シートが見つかりません。'); return; }

  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  const protection = sheet.protect();
  protection.setDescription('手作業シート保護 — オーナーのみ編集可');
  const me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors().filter(e => e.getEmail() !== me.getEmail()));
  protection.setWarningOnly(false);
  SpreadsheetApp.getUi().alert('✅ 手作業シートを保護しました。\nオーナー：' + me.getEmail());
}

/** 手作業シートの保護を解除する */
function removeSheetProtection() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DEFAULT_SHEET_NAME2);
  if (!sheet) return;
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  SpreadsheetApp.getUi().alert('手作業シートの保護を解除しました。');
}

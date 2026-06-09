// ─── 初回セットアップ用関数（カスタムメニューまたはエディタから手動実行） ──────

/** Info シートを作成・初期化する */
function setupInfoSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(INFO_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(INFO_SHEET_NAME, 0);

  if (sheet.getLastRow() > 1) {
    const ui  = SpreadsheetApp.getUi();
    const res = ui.alert('⚠️ Info シートは既に存在します。上書きしますか？', ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;
    sheet.clearContents();
  }

  const rows = [
    ['キー', '値', '説明'],
    ['EVENT_NAME',        '第5回川口花火大会',              'イベント名'],
    ['EVENT_DATE',        '2026-09-20',                     '開催日'],
    ['PAYMENT_DUE',       '9月12日（金）',                  '入金期限'],
    ['OFFICE_EMAIL',      OFFICE_EMAIL_DEFAULT,              '事務局メールアドレス'],
    ['HANKO_FILE_ID',     HANKO_FILE_ID_DEFAULT,             '印影PNG の Google Drive ファイルID'],
    ['ROOT_FOLDER_ID',    '',                                'プロジェクト作成先のフォルダID（Drive）'],
    ['START_DATE',        '2025-01-01T00:00:00',             'フォーム受付開始日時'],
    ['END_DATE',          '2026-10-01T23:59:59',             'フォーム受付終了日時'],
    ['PRICE_S',           2000000,                           'S協賛 金額（円）'],
    ['PRICE_A',           1000000,                           'A協賛 金額（円）'],
    ['PRICE_B',           500000,                            'B協賛 金額（円）'],
    ['PRICE_C',           300000,                            'C協賛 金額（円）'],
    ['PRICE_D',           200000,                            'D協賛 金額（円）'],
    ['PRICE_E',           100000,                            'E協賛 金額（円）'],
    ['DATA_SPREADSHEET_ID', '',                              'データシートID（initProject で自動設定）'],
    // ─── 組織情報 ───────────────────────────────────────────────────────────────
    ['ORG_NAME',         '川口花火大会実行委員会',              '発行者組織名（請求書・お礼状に使用）'],
    ['ORG_REP',          '委員長　廣瀬 進治',                   '代表者名（請求書・お礼状に使用）'],
    ['INVOICE_REG_NO',   'T9700150122003',                      'インボイス登録番号'],
    ['BANK_NAME',        '埼玉りそな銀行　川口支店',            '振込先銀行名・支店名'],
    ['BANK_NO',          '36216349',                            '口座番号（普通預金）'],
    ['BANK_HOLDER',      '川口商工会議所',                      '口座名義（略称）'],
    ['BANK_REP',         '川口花火大会実行委員会委員長　廣瀬 進治', '口座名義（正式）'],
    ['ORG_LOCATION',     '（川口商工会議所内）',                '事務所所在表示'],
    ['ORG_TEL',          '048-228-2220',                        '事務局電話番号'],
    ['ORG_FAX',          '048-228-2221',                        '事務局FAX番号'],
    ['OFFICE_HOURS',     '平日 10:00 〜 17:00',                 '受付時間（メール文末に表示）'],
    ['RECEIPT_NO_PREFIX','KWGC',                                '受付番号プレフィックス'],
  ];

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#d0e4f7');
  sheet.getRange(2, 1, rows.length - 1, 1).setBackground('#f8f8f8').setFontWeight('bold');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 300);
  sheet.setFrozenRows(1);

  _infoCache = null;
  SpreadsheetApp.getUi().alert('✅ Info シートを作成しました。\n各値を確認・編集してください。');
}

/** 申込みシートのヘッダー行を作成する（初回のみ） */
function setupHeaders() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DEFAULT_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert(`シート "${DEFAULT_SHEET_NAME}" が見つかりません。`); return; }
  if (sheet.getLastRow() > 0) { SpreadsheetApp.getUi().alert('ヘッダー行はすでに存在します。'); return; }
  sheet.appendRow(HEADERS);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#d0e4f7');
  applyColumnWidths(sheet);
  SpreadsheetApp.getUi().alert('ヘッダー行を作成しました。');
}

/** すべてのメールテンプレートを Script Properties に保存する */
function setupAllMailTemplates() {
  setupMailTemplate();
  setupReceiptOnlyTemplate();
  setupOreijouTemplate();
  setupAnnaiTemplate();
  SpreadsheetApp.getUi().alert('✅ 全メールテンプレートを保存しました。');
}

function setupMailTemplate() {
  PropertiesService.getScriptProperties().setProperties({
    MAIL_SUBJECT: '【{{event_name}}】申込受理書兼請求書のご連絡（受付番号：{{receipt_no}}）',
    MAIL_BODY:    _defaultConfirmBody(),
  });
}

function setupReceiptOnlyTemplate() {
  PropertiesService.getScriptProperties().setProperties({
    RECEIPT_ONLY_SUBJECT: '【{{event_name}}】お申し込みを受け付けました（受付番号：{{receipt_no}}）',
    RECEIPT_ONLY_BODY:    _defaultReceiptOnlyBody(),
  });
}

function setupOreijouTemplate() {
  PropertiesService.getScriptProperties().setProperties({
    OREIJOU_SUBJECT: '【{{event_name}}】ご協賛へのお礼（受付番号：{{receipt_no}}）',
    OREIJOU_BODY:    _defaultOreijouBody(),
  });
}

function setupAnnaiTemplate() {
  PropertiesService.getScriptProperties().setProperties({
    ANNAI_SUBJECT: '【{{event_name}}】ご案内（受付番号：{{receipt_no}}）',
    ANNAI_BODY:    _defaultAnnaiBody(),
  });
}

/**
 * 手作業シートのヘッダー行・サブヘッダー行を最新の列構成に修正する
 * ★ GAS エディタから直接 "repairTesagyouNow" を実行してください
 */
function repairTesagyouNow() {
  // Main の Info シートから DATA_SPREADSHEET_ID を直接取得
  const mainSs = SpreadsheetApp.getActiveSpreadsheet(); // GAS エディタ実行時 = Main
  const infoSheet = mainSs.getSheetByName(INFO_SHEET_NAME);
  let dataSsId = '';
  if (infoSheet) {
    const rows = infoSheet.getDataRange().getValues();
    for (const row of rows) {
      if (String(row[0]).trim() === 'DATA_SPREADSHEET_ID') { dataSsId = String(row[1]).trim(); break; }
    }
  }

  let sheet = null;
  if (dataSsId) {
    try { sheet = SpreadsheetApp.openById(dataSsId).getSheetByName(DEFAULT_SHEET_NAME2); } catch (e) {}
  }
  if (!sheet) sheet = mainSs.getSheetByName(DEFAULT_SHEET_NAME2);
  if (!sheet) { console.error('手作業シートが見つかりません。'); return; }

  _applyTesagyouHeaders(sheet);
  console.log('✅ 手作業ヘッダーを修正しました。対象: ' + sheet.getParent().getName());
}

function fixTesagyouHeaders() {
  // Data spreadsheet → Active の順で手作業シートを探す
  let sheet = null;
  let ssName = '';
  const candidates = [getDataSpreadsheet(), SpreadsheetApp.getActiveSpreadsheet()];
  for (const ss of candidates) {
    try {
      const s = ss.getSheetByName(DEFAULT_SHEET_NAME2);
      if (s) { sheet = s; ssName = ss.getName(); break; }
    } catch (e) { /* skip */ }
  }
  if (!sheet) { SpreadsheetApp.getUi().alert('手作業シートが見つかりません。'); return; }
  _applyTesagyouHeaders(sheet);
  SpreadsheetApp.getUi().alert(`✅ 手作業ヘッダーを修正しました。\n対象: ${ssName}`);
  if (!sheet) { SpreadsheetApp.getUi().alert('手作業シートが見つかりません。'); return; }

  _applyTesagyouHeaders(sheet);
  SpreadsheetApp.getUi().alert('✅ 手作業シートのヘッダーを修正しました。');
}

/** 手作業シートのヘッダー/サブヘッダーを書き込む共通処理 */
function _applyTesagyouHeaders(sheet) {
  // 行1: ヘッダー
  sheet.getRange(1, 1, 1, TESAGYOU_HEADERS.length).setValues([TESAGYOU_HEADERS]);
  sheet.getRange(1, 1, 1, TESAGYOU_HEADERS.length)
    .setFontWeight('bold').setBackground('#fce8b2');

  // 行2: サブヘッダー（17列）
  const sub = [
    'XLOOKUP\n自動',  '直接入力',                                          // A,B
    'XLOOKUP\n自動',  'XLOOKUP\n自動', 'XLOOKUP\n自動',                    // C,D,E
    'XLOOKUP\n自動',  'XLOOKUP\n自動', 'XLOOKUP\n自動', 'XLOOKUP\n自動',   // F,G,H,I
    'checkbox\n手動', 'タイムスタンプ\n自動', 'checkbox\n手動',              // J,K,L
    'タイムスタンプ\n自動', '区分＋7桁\n自動',                               // M,N
    'checkbox\n手動', 'タイムスタンプ\n自動', 'タイムスタンプ\n自動',         // O,P,Q
  ];
  sheet.getRange(2, 1, 1, sub.length).setValues([sub]);
  sheet.getRange(2, 1, 1, sub.length)
    .setFontSize(8).setFontColor('#888888').setBackground('#fffbf0').setWrap(true);
  sheet.setRowHeight(2, 36);
  sheet.setFrozenRows(2);
  applyTesagyouColumnWidths(sheet);
}

/** 手作業シートの既存行を修正（チェックボックス列の修正） */
function fixTesagyouSheet() {
  const dataSs = getDataSpreadsheet();
  const sheet  = dataSs.getSheetByName(DEFAULT_SHEET_NAME2);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  sheet.getRange(3, COL_UKETSUKE, lastRow - 2).insertCheckboxes();
  sheet.getRange(3, COL_NYUKIN,   lastRow - 2).insertCheckboxes();
  sheet.getRange(3, COL_ANNAIBUN, lastRow - 2).insertCheckboxes();
  SpreadsheetApp.getUi().alert('✅ チェックボックスを修正しました。');
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

function removeSheetProtection() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DEFAULT_SHEET_NAME2);
  if (!sheet) return;
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  SpreadsheetApp.getUi().alert('手作業シートの保護を解除しました。');
}

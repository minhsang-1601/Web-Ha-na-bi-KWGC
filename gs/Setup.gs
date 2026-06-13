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

  const O = '⭕️';
  const rows = [
    ['キー', '値', '必須', '説明'],
    ['EVENT_NAME',        '第5回川口花火大会',              O, 'イベント名'],
    ['EVENT_DATE',        '2026-09-20',                     O, '開催日'],
    ['PAYMENT_DUE',       '9月18日（金）',                  O, '入金期限'],
    ['OFFICE_EMAIL',      OFFICE_EMAIL_DEFAULT,              O, '事務局メールアドレス'],
    ['HANKO_FILE_ID',     HANKO_FILE_ID_DEFAULT,             O, '印影PNG の Google Drive ファイルID'],
    ['ROOT_FOLDER_ID',    '',                                O, 'プロジェクト作成先のフォルダID（Drive）'],
    ['START_DATE',        '2025-01-01T00:00:00',             O, 'フォーム受付開始日時'],
    ['END_DATE',          '2026-10-01T23:59:59',             O, 'フォーム受付終了日時'],
    ['PRICE_S',           2000000,                           O, 'S協賛 金額（円）'],
    ['PRICE_A',           1000000,                           O, 'A協賛 金額（円）'],
    ['PRICE_B',           500000,                            O, 'B協賛 金額（円）'],
    ['PRICE_C',           300000,                            O, 'C協賛 金額（円）'],
    ['PRICE_D',           200000,                            O, 'D協賛 金額（円）'],
    ['PRICE_E',           100000,                            O, 'E協賛 金額（円）'],
    ['DATA_SPREADSHEET_ID', '',                              O, 'データシートID（initProject で自動設定）'],
    // ─── 組織情報 ───────────────────────────────────────────────────────────────
    ['ORG_NAME',         '川口花火大会実行委員会',              O, '発行者組織名（請求書・お礼状に使用）'],
    ['ORG_REP',          '委員長　廣瀬 進治',                   O, '代表者名（請求書・お礼状に使用）'],
    ['INVOICE_REG_NO',   'T9700150122003',                      O, 'インボイス登録番号'],
    ['BANK_NAME',        '埼玉りそな銀行　川口支店',            O, '振込先銀行名・支店名'],
    ['BANK_NO',          '6216349',                            O, '口座番号（普通預金）'],
    ['BANK_HOLDER',      '川口商工会議所',                      O, '口座名義（略称）'],
    ['BANK_REP',         '川口花火大会実行委員会委員長　田中　？？', O, '口座名義（正式）'],
    ['ORG_LOCATION',     '（川口商工会議所内）',                O, '事務所所在表示'],
    ['ORG_TEL',          '048-228-2220',                        O, '事務局電話番号'],
    ['ORG_FAX',          '048-228-2221',                        O, '事務局FAX番号'],
    ['OFFICE_HOURS',     '平日 10:00 〜 17:00',                 O, '受付時間（メール文末に表示）'],
    ['RECEIPT_NO_PREFIX','KWGC',                                O, '受付番号プレフィックス'],
    ['MIN_MAIL_QUOTA',   MIN_MAIL_QUOTA_DEFAULT,                O, 'メール送信残数の最低ライン（これ未満はフォーム停止・最低5）'],
    // ─── 区分別申込み期間 ────────────────────────────────────────────────────────
    ['KUBUN_SA_START',   '2026-06-01T00:00:00',                O, 'S・A 区分 申込み受付開始日時'],
    ['KUBUN_SA_END',     '2026-07-07T23:59:59',                O, 'S・A 区分 申込み受付終了日時'],
    ['KUBUN_BCDE_START', '2026-06-01T00:00:00',                O, 'B〜E 区分 申込み受付開始日時'],
    ['KUBUN_BCDE_END',   '2026-09-04T23:59:59',                O, 'B〜E 区分 申込み受付終了日時'],
  ];

  sheet.getRange(1, 1, rows.length, 4).setValues(rows);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#d0e4f7');
  sheet.getRange(2, 1, rows.length - 1, 1).setBackground('#f8f8f8').setFontWeight('bold');
  sheet.getRange(2, 3, rows.length - 1, 1).setHorizontalAlignment('center').setFontColor('#c0392b');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 60);
  sheet.autoResizeColumn(4);
  sheet.setFrozenRows(1);
  // フィルターを4列すべてに適用
  const lastRow = sheet.getLastRow();
  const filterRange = sheet.getRange(1, 1, lastRow, 4);
  try {
    const existing = sheet.getFilter();
    if (existing) existing.remove();
  } catch (_) {}
  filterRange.createFilter();
  _applyAlignment(sheet, 2, 4);

  _infoCache = null;
  SpreadsheetApp.getUi().alert('✅ Info シートを作成しました。\n各値を確認・編集してください。');
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
    MAIL_SUBJECT: '【{{event_name}}】協賛お申込みを受け付けました。',
    MAIL_BODY:    _defaultConfirmBody(),
  });
}

function setupReceiptOnlyTemplate() {
  PropertiesService.getScriptProperties().setProperties({
    RECEIPT_ONLY_SUBJECT: '【{{event_name}}】協賛お申込みを受け付けました。',
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

/** 手作業シートのヘッダー/サブヘッダーを書き込む共通処理 */
function _applyTesagyouHeaders(sheet) {
  // 行1: ヘッダー
  sheet.getRange(1, 1, 1, TESAGYOU_HEADERS.length).setValues([TESAGYOU_HEADERS]);
  sheet.getRange(1, 1, 1, TESAGYOU_HEADERS.length)
    .setFontWeight('bold').setBackground('#fce8b2');

  // 行2: サブヘッダー（16列）
  const sub = [
    '直接入力',                                                                    // A
    'XLOOKUP\n自動', 'XLOOKUP\n自動', 'XLOOKUP\n自動',                            // B,C,D
    'XLOOKUP\n自動', 'XLOOKUP\n自動', 'XLOOKUP\n自動', 'XLOOKUP\n自動',           // E,F,G,H
    'checkbox\n手動', 'タイムスタンプ\n自動', 'checkbox\n手動',                    // I,J,K
    'タイムスタンプ\n自動', '区分＋7桁\n自動',                                     // L,M
    'checkbox\n手動', 'タイムスタンプ\n自動', 'タイムスタンプ\n自動',               // N,O,P
  ];
  sheet.getRange(2, 1, 1, sub.length).setValues([sub]);
  sheet.getRange(2, 1, 1, sub.length)
    .setFontSize(8).setFontColor('#888888').setBackground('#fffbf0').setWrap(true);
  sheet.setRowHeight(2, 36);
  sheet.setFrozenRows(2);
  applyTesagyouColumnWidths(sheet);
  _ensureFilter(sheet, 1, TESAGYOU_HEADERS.length);
}

/**
 * 既存シートのヘッダー行の列名を最新の名称に一括更新する
 * GAS エディタから "renameHeaders" を実行してください
 * ※ データ行はそのまま、ヘッダー行（行1）のみ書き換えます
 */
function renameHeaders() {
  const dataSs = getDataSpreadsheet();

  // ── 協賛申込み一覧 ─────────────────────────────────────────────────────────
  const sheet1 = dataSs.getSheetByName(DEFAULT_SHEET_NAME);
  if (sheet1 && sheet1.getLastRow() > 0) {
    sheet1.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    console.log('✅ 協賛申込み一覧: ヘッダーを更新しました。');
  } else {
    console.log('⚠️ 協賛申込み一覧: シートが見つからないかデータがありません。');
  }

  // ── 手作業 ─────────────────────────────────────────────────────────────────
  const sheet2 = dataSs.getSheetByName(DEFAULT_SHEET_NAME2);
  if (sheet2 && sheet2.getLastRow() > 0) {
    sheet2.getRange(1, 1, 1, TESAGYOU_HEADERS.length).setValues([TESAGYOU_HEADERS]);
    console.log('✅ 手作業: ヘッダーを更新しました。');
  } else {
    console.log('⚠️ 手作業: シートが見つからないかデータがありません。');
  }

  try {
    SpreadsheetApp.getUi().alert('✅ ヘッダーを更新しました。\n・協賛申込み一覧\n・手作業');
  } catch (_) {}
}

/**
 * 全シートにフィルターを設定する（既存シートへの一括適用）
 * カスタムメニュー「フィルター設定」から実行
 */
function applyFiltersToAllSheets() {
  const mainSs = SpreadsheetApp.getActiveSpreadsheet();
  const dataSs = getDataSpreadsheet();

  // ── Main スプレッドシート ──────────────────────────────────────────
  // Info
  const infoSheet = mainSs.getSheetByName(INFO_SHEET_NAME);
  if (infoSheet && infoSheet.getLastRow() > 0) {
    _ensureFilter(infoSheet, 1, 3);
    _applyAlignment(infoSheet, 2, 3);
  }

  // CreateLog
  const logSheet = mainSs.getSheetByName(CREATELOG_SHEET);
  if (logSheet && logSheet.getLastRow() > 0) {
    _ensureFilter(logSheet, 1, CREATELOG_HEADERS.length);
    _applyAlignment(logSheet, 2, CREATELOG_HEADERS.length);
  }

  // ── データ スプレッドシート ────────────────────────────────────────
  // 協賛申込み一覧
  const sheet1 = dataSs.getSheetByName(DEFAULT_SHEET_NAME);
  if (sheet1 && sheet1.getLastRow() > 0) {
    _ensureFilter(sheet1, 1, HEADERS.length);
    _applyAlignment(sheet1, 2, HEADERS.length);
  }

  // 手作業
  const sheet2 = dataSs.getSheetByName(DEFAULT_SHEET_NAME2);
  if (sheet2 && sheet2.getLastRow() > 0) {
    _ensureFilter(sheet2, 1, TESAGYOU_HEADERS.length);
    _applyAlignment(sheet2, 3, TESAGYOU_HEADERS.length);
  }

  SpreadsheetApp.getUi().alert('✅ 全シートにフィルターを設定しました。');
}


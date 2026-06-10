// ─── プロジェクト初期化（フォルダ作成・スプレッドシート作成・トリガー設定） ────
// 使い方:
//   1. Info シートの ROOT_FOLDER_ID に Google Drive フォルダIDをセット
//   2. カスタムメニュー「協賛管理 → プロジェクト初期化」を実行
//   3. yyyyMMdd フォルダ内にデータ用スプレッドシートが自動作成される
//   4. onEditInstallable / onOpenEventSheet トリガーが自動登録される

function initProject() {
  const ui = SpreadsheetApp.getUi();

  // ── 1. Info シート確認 ────────────────────────────────────────────────────────
  const mainSs    = SpreadsheetApp.getActiveSpreadsheet();
  const infoSheet = mainSs.getSheetByName(INFO_SHEET_NAME);
  if (!infoSheet) {
    ui.alert('⚠️ Info シートが見つかりません。\n先に「Info シート作成」を実行してください。');
    return;
  }

  // ── 2. ROOT_FOLDER_ID 確認 ───────────────────────────────────────────────────
  const rootFolderId = getRootFolderId();
  if (!rootFolderId) {
    ui.alert('⚠️ ROOT_FOLDER_ID が設定されていません。\nInfo シートの ROOT_FOLDER_ID に Google Drive フォルダIDを入力してください。');
    return;
  }
  let rootFolder;
  try {
    rootFolder = DriveApp.getFolderById(rootFolderId);
  } catch (e) {
    ui.alert('❌ ROOT_FOLDER_ID が無効です: ' + rootFolderId);
    return;
  }

  // ── 3. 既に初期化済みの場合は確認 ────────────────────────────────────────────
  const existingId = getConfigVal('DATA_SPREADSHEET_ID', '');
  if (existingId) {
    const res = ui.alert(
      '⚠️ すでにプロジェクトが初期化されています。',
      `DATA_SPREADSHEET_ID: ${existingId}\n新しく作成しますか？（既存のトリガーも再登録されます）`,
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) return;
  }

  // ── 4. フォルダ作成 ───────────────────────────────────────────────────────────
  const yyyy      = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy');
  const eventName = getEventName();
  const projectId = `${yyyy}_${eventName}`;
  const yearFolders = rootFolder.getFoldersByName(projectId);
  const folder      = yearFolders.hasNext() ? yearFolders.next() : rootFolder.createFolder(projectId);

  // ── 5. データ用スプレッドシート作成 ──────────────────────────────────────────
  const ssName     = `協賛データ_${projectId}`;
  const existFiles = folder.getFilesByName(ssName);
  if (existFiles.hasNext()) {
    ui.alert(`⚠️ 「${ssName}」はすでに存在します。\n既存ファイルのIDを DATA_SPREADSHEET_ID に設定してください。\n\nファイルURL: ${existFiles.next().getUrl()}`);
    return;
  }
  const newSs = SpreadsheetApp.create(ssName);
  DriveApp.getFileById(newSs.getId()).moveTo(folder);

  // ── 6. 協賛申込み一覧 シート ──────────────────────────────────────────────────
  const sheet1 = newSs.getSheets()[0];
  sheet1.setName(DEFAULT_SHEET_NAME);
  sheet1.appendRow(HEADERS);
  sheet1.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#d0e4f7');
  sheet1.setFrozenRows(1);
  applyColumnWidths(sheet1);
  _ensureFilter(sheet1, 1, HEADERS.length);

  // ── 7. 手作業 シート（サブヘッダー・フリーズ・フィルター完全版） ────────────────
  const sheet2 = newSs.insertSheet(DEFAULT_SHEET_NAME2);
  _applyTesagyouHeaders(sheet2);   // ヘッダー＋サブヘッダー＋フリーズ＋幅＋フィルター

  // ── 8. セルの配置: 上揃え・左揃え ──────────────────────────────────────────
  sheet1.getRange(2, 1, sheet1.getMaxRows() - 1, HEADERS.length)
    .setVerticalAlignment('top').setHorizontalAlignment('left');
  sheet2.getRange(3, 1, sheet2.getMaxRows() - 2, TESAGYOU_HEADERS.length)
    .setVerticalAlignment('top').setHorizontalAlignment('left');

  // ── 9. 管理IDリストシート 初期化 ─────────────────────────────────────────────
  initKanriSheet();

  // ── 10. メールテンプレート保存 ───────────────────────────────────────────────
  setupMailTemplate();
  setupReceiptOnlyTemplate();
  setupOreijouTemplate();
  setupAnnaiTemplate();

  // ── 11. Info シートに ID を保存 ──────────────────────────────────────────────
  setInfoValue('DATA_SPREADSHEET_ID', newSs.getId());
  _infoCache = null;

  // ── 12. トリガー登録 ─────────────────────────────────────────────────────────
  _removeTriggersForFunction('onEditInstallable');
  _removeTriggersForFunction('onOpenEventSheet');
  ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(newSs).onEdit().create();
  ScriptApp.newTrigger('onOpenEventSheet').forSpreadsheet(newSs).onOpen().create();

  // ── 13. CreateLog に記録 ──────────────────────────────────────────────────────
  _writeCreateLog('initProject', [
    `パス: ${projectId}/`,
    `フォルダURL: ${folder.getUrl()}`,
    `スプレッドシート名: ${ssName}`,
    `スプレッドシートURL: ${newSs.getUrl()}`,
    `スプレッドシートID: ${newSs.getId()}`,
  ].join('\n'));

  // ── 12. 完了メッセージ ────────────────────────────────────────────────────────
  ui.alert(
    `✅ プロジェクトを初期化しました。\n\n` +
    `📁 フォルダ: ${projectId}\n` +
    `📊 シート: ${ssName}\n` +
    `📧 メールテンプレート: 保存済み\n` +
    `🔔 トリガー: onEditInstallable / onOpenEventSheet を登録済み\n\n` +
    `データ用スプレッドシートを開いてください:\n${newSs.getUrl()}`
  );
}

// ─── CreateLog 書き込み ────────────────────────────────────────────────────────

const CREATELOG_HEADERS = ['日時', '操作', '詳細', '実行者'];

function _writeCreateLog(action, detail) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet(); // 常に Main に書く
  let sheet = ss.getSheetByName(CREATELOG_SHEET);

  if (!sheet) {
    // シートが存在しない → 新規作成
    sheet = ss.insertSheet(CREATELOG_SHEET);
  }

  // ヘッダーが未設定（シートが空 or 先頭セルが違う）→ ヘッダー行を挿入
  const firstCell = String(sheet.getRange(1, 1).getValue()).trim();
  if (firstCell !== '日時') {
    if (sheet.getLastRow() > 0) sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, CREATELOG_HEADERS.length).setValues([CREATELOG_HEADERS]);
    sheet.getRange(1, 1, 1, CREATELOG_HEADERS.length)
      .setFontWeight('bold').setBackground('#e8f4f8');
    [170, 150, 400, 220].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    sheet.setFrozenRows(1);
    _ensureFilter(sheet, 1, CREATELOG_HEADERS.length);
    _applyAlignment(sheet, 2, CREATELOG_HEADERS.length);
  }

  const now  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const user = Session.getActiveUser().getEmail() || '(unknown)';
  sheet.appendRow([now, action, detail, user]);
}

// ─── トリガー管理 ─────────────────────────────────────────────────────────────

function _removeTriggersForFunction(funcName) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === funcName)
    .forEach(t => ScriptApp.deleteTrigger(t));
}


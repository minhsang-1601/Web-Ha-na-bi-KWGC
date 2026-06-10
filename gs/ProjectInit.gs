// ─── プロジェクト初期化（フォルダ作成・スプレッドシート作成・トリガー設定） ────
// 使い方:
//   1. Info シートの ROOT_FOLDER_ID に Google Drive フォルダIDをセット
//   2. カスタムメニュー「協賛管理 → プロジェクト初期化」を実行
//   3. yyyyMMdd フォルダ内にデータ用スプレッドシートが自動作成される
//   4. onEditInstallable / onOpenEventSheet トリガーが自動登録される

function initProject() {
  const rootFolderId = getRootFolderId();
  if (!rootFolderId) {
    SpreadsheetApp.getUi().alert(
      '⚠️ ROOT_FOLDER_ID が設定されていません。\nInfo シートの ROOT_FOLDER_ID に Google Drive フォルダIDを入力してください。'
    );
    return;
  }

  let rootFolder;
  try {
    rootFolder = DriveApp.getFolderById(rootFolderId);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ ROOT_FOLDER_ID が無効です: ' + rootFolderId);
    return;
  }

  // 既に初期化済みの場合は確認
  const existingId = getConfigVal('DATA_SPREADSHEET_ID', '');
  if (existingId) {
    const res = SpreadsheetApp.getUi().alert(
      '⚠️ すでにプロジェクトが初期化されています。',
      `DATA_SPREADSHEET_ID: ${existingId}\n新しく作成しますか？（既存のトリガーも再登録されます）`,
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (res !== SpreadsheetApp.getUi().Button.YES) return;
  }

  // yyyy_EVENT_NAME フォルダ作成（既存なら再利用）
  const yyyy      = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy');
  const eventName = getEventName();
  const projectId = `${yyyy}_${eventName}`;

  const yearFolders = rootFolder.getFoldersByName(projectId);
  const folder      = yearFolders.hasNext() ? yearFolders.next() : rootFolder.createFolder(projectId);

  // データ用スプレッドシート作成（同名ファイルが既にあればスキップ）
  const ssName     = `協賛データ_${projectId}`;
  const existFiles = folder.getFilesByName(ssName);
  if (existFiles.hasNext()) {
    SpreadsheetApp.getUi().alert(
      `⚠️ 「${ssName}」はすでに存在します。\n既存ファイルのIDを DATA_SPREADSHEET_ID に設定してください。\n\nファイルURL: ${existFiles.next().getUrl()}`
    );
    return;
  }
  const newSs = SpreadsheetApp.create(ssName);
  DriveApp.getFileById(newSs.getId()).moveTo(folder);

  // シート作成
  const sheet1 = newSs.getSheets()[0];
  sheet1.setName(DEFAULT_SHEET_NAME);
  sheet1.appendRow(HEADERS);
  sheet1.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#d0e4f7');
  applyColumnWidths(sheet1);

  const sheet2 = newSs.insertSheet(DEFAULT_SHEET_NAME2);
  sheet2.appendRow(TESAGYOU_HEADERS);
  sheet2.getRange(1, 1, 1, TESAGYOU_HEADERS.length).setFontWeight('bold').setBackground('#fce8b2');
  // サブヘッダー行
  sheet2.appendRow([
    '', '', '', '', '', '', '', '',
    'checkbox\n手動', 'タイムスタンプ\n自動', 'checkbox\n手動',
    'タイムスタンプ\n自動', '区分＋7桁\n自動',
    'checkbox\n手動', 'タイムスタンプ\n自動', 'タイムスタンプ\n自動',
  ]);
  sheet2.getRange(2, 1, 1, TESAGYOU_HEADERS.length)
    .setFontSize(8).setFontColor('#888888').setBackground('#fffbf0').setWrap(true);
  sheet2.setRowHeight(2, 36);
  applyTesagyouColumnWidths(sheet2);

  // Info シートに ID を保存
  setInfoValue('DATA_SPREADSHEET_ID', newSs.getId());

  // トリガーを登録（既存の同種トリガーを削除してから）
  _removeTriggersForFunction('onEditInstallable');
  _removeTriggersForFunction('onOpenEventSheet');

  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(newSs)
    .onEdit()
    .create();

  ScriptApp.newTrigger('onOpenEventSheet')
    .forSpreadsheet(newSs)
    .onOpen()
    .create();

  // CreateLog に記録
  _writeCreateLog('initProject', [
    `パス: ${projectId}/`,
    `フォルダURL: ${folder.getUrl()}`,
    `スプレッドシート名: ${ssName}`,
    `スプレッドシートURL: ${newSs.getUrl()}`,
    `スプレッドシートID: ${newSs.getId()}`,
  ].join('\n'));

  SpreadsheetApp.getUi().alert(
    `✅ プロジェクトを初期化しました。\n\n` +
    `📁 フォルダ: ${projectId}\n` +
    `📊 シート: ${ssName}\n` +
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

/** トリガー一覧を確認（デバッグ用） */
function listTriggers() {
  const list = ScriptApp.getProjectTriggers().map(t =>
    `${t.getHandlerFunction()} → ${t.getTriggerSource()} / ${t.getEventType()}`
  );
  SpreadsheetApp.getUi().alert('登録済みトリガー:\n' + (list.join('\n') || 'なし'));
}

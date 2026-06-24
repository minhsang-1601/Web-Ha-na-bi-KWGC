// ─── プロジェクト初期化（フォルダ作成・スプレッドシート作成・トリガー設定） ────
// 使い方:
//   1. Info シートの ROOT_FOLDER_ID に Google Drive フォルダIDをセット
//   2. カスタムメニュー「協賛管理 → プロジェクト初期化」を実行
//   3. yyyyMMdd フォルダ内にデータ用スプレッドシートが自動作成される
//   4. onEditInstallable / onOpenEventSheet トリガーが自動登録される

function initProject() {
  const ui = SpreadsheetApp.getUi();

  // ── 1. Info シート確認（なければ自動作成） ───────────────────────────────────
  const mainSs = SpreadsheetApp.getActiveSpreadsheet();
  if (!mainSs.getSheetByName(INFO_SHEET_NAME)) {
    setupInfoSheet();
    ui.alert('ℹ️ Info シートが存在しなかったため自動作成しました。\n必要に応じて値を確認・編集してください。');
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

  // ── 8.5 操作ログ シート（監査ログ）を作成 ───────────────────────────────────
  _getAuditSheet(newSs);

  // ── 9. メールテンプレート保存 ───────────────────────────────────────────────
  setupMailTemplate();
  setupReceiptOnlyTemplate();
  setupOreijouTemplate();

  // ── 11. Info シートに ID を保存 ──────────────────────────────────────────────
  setInfoValue('DATA_SPREADSHEET_ID', newSs.getId());
  _infoCache = null;

  // ── 12. トリガー登録 ─────────────────────────────────────────────────────────
  _removeTriggersForFunction('onEditInstallable');
  _removeTriggersForFunction('onOpenEventSheet');
  _removeTriggersForFunction('onChangeInstallable');
  ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(newSs).onEdit().create();
  ScriptApp.newTrigger('onOpenEventSheet').forSpreadsheet(newSs).onOpen().create();
  ScriptApp.newTrigger('onChangeInstallable').forSpreadsheet(newSs).onChange().create();

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

/**
 * 既存のデータ用スプレッドシートにトリガーを（再）登録する
 * ※ initProject を使わずチェックボックスのダイアログを復活させたいときに実行
 *   メニュー「⚙️ 初期設定 → トリガー再登録」または GAS エディタから実行
 */
function registerTriggers() {
  const ui = SpreadsheetApp.getUi();
  const dataSsId = getConfigVal('DATA_SPREADSHEET_ID', '');
  if (!dataSsId) {
    ui.alert('⚠️ DATA_SPREADSHEET_ID が未設定です。\n先に「プロジェクト初期化」を実行してください。');
    return;
  }

  let dataSs = null;
  try { dataSs = SpreadsheetApp.openById(dataSsId); } catch (e) {
    ui.alert('❌ データ用スプレッドシートを開けません: ' + dataSsId);
    return;
  }

  _removeTriggersForFunction('onEditInstallable');
  _removeTriggersForFunction('onOpenEventSheet');
  _removeTriggersForFunction('onChangeInstallable');
  ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(dataSs).onEdit().create();
  ScriptApp.newTrigger('onOpenEventSheet').forSpreadsheet(dataSs).onOpen().create();
  ScriptApp.newTrigger('onChangeInstallable').forSpreadsheet(dataSs).onChange().create();

  ui.alert(`✅ トリガーを再登録しました。\n対象: ${dataSs.getName()}\n\n・チェックボックス操作で確認ダイアログが表示されます。\n・全ての編集・行列の挿入削除が「操作ログ」シートに記録されます。`);
}

/**
 * トリガーの状態を診断する（どのファイルに紐づいているか確認）
 * GAS エディタから実行 → 実行ログを確認
 */
function diagnoseTriggers() {
  const dataSsId = getConfigVal('DATA_SPREADSHEET_ID', '');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('■ Info の DATA_SPREADSHEET_ID:');
  console.log('   ' + (dataSsId || '（未設定）'));
  let dataName = '';
  try { dataName = SpreadsheetApp.openById(dataSsId).getName(); } catch (_) {}
  console.log('   名前: ' + (dataName || '（開けません）'));
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('■ 登録済みトリガー:');

  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    console.log('   （トリガーなし）');
  }
  triggers.forEach((t, i) => {
    const fn  = t.getHandlerFunction();
    let srcId = '';
    try { srcId = t.getTriggerSourceId(); } catch (_) {}
    let srcName = '';
    try { srcName = SpreadsheetApp.openById(srcId).getName(); } catch (_) {}
    const match = (srcId === dataSsId) ? '✅一致' : '❌不一致';
    console.log(`   [${i + 1}] ${fn}`);
    console.log(`        紐付け先ID: ${srcId}`);
    console.log(`        紐付け先名: ${srcName || '(開けません)'}  ${match}`);
  });
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━');
}

function _removeTriggersForFunction(funcName) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === funcName)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/**
 * 全シートの保護をすべて解除する（以前の「シート保護」を取り消す用）。
 * ★ オーナーアカウントで実行すること。
 * メニュー「⚙️ 初期設定 → シート保護を解除」または GAS エディタから実行。
 */
function unprotectAllSheets() {
  const mainSs = SpreadsheetApp.getActiveSpreadsheet();
  const dataSs = getDataSpreadsheet();
  let count = 0;

  [mainSs, dataSs].forEach((ss, idx) => {
    if (idx === 1 && ss.getId() === mainSs.getId()) return; // 同一ファイルは1回だけ
    ss.getSheets().forEach(sheet => {
      sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
        .forEach(p => { try { p.remove(); count++; } catch (_) {} });
      sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
        .forEach(p => { try { p.remove(); count++; } catch (_) {} });
    });
  });

  try {
    SpreadsheetApp.getUi().alert(`✅ シート保護を解除しました（${count} 件）。\nすべてのシートが編集可能になりました。`);
  } catch (_) {}
}

/**
 * 【重要】ログが二重に記録される / メールが二重送信される場合に実行。
 * インストール型トリガーはアカウントごとに作成されるため、複数アカウントが
 * トリガーを持つと、1回の編集で全員のトリガーが発火し重複する。
 * → デプロイ者「以外」の各アカウントでこの関数を実行し、自分のトリガーを全削除する。
 *   （トリガーはデプロイ者の1セットだけ残す。デプロイ者のトリガーが
 *    全員の編集に対して発火するため、他アカウントのトリガーは不要。）
 * GAS エディタでこの関数を選んで実行 → 実行ログを確認。
 */
function removeMyTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const me = Session.getActiveUser().getEmail() || '(unknown)';
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  console.log(`✅ ${me} のトリガー ${triggers.length} 件を削除しました。`);
  try {
    SpreadsheetApp.getUi().alert(
      `✅ このアカウント（${me}）のトリガーを ${triggers.length} 件削除しました。\n\n` +
      `※ デプロイ者アカウントでは実行しないでください。\n` +
      `　 デプロイ者のトリガーが全員の操作に対して動作します。`
    );
  } catch (_) {}
}


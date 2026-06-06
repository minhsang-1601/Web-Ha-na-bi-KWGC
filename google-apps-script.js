/**
 * Google Apps Script — 川口花火大会 協賛申込みフォーム受信スクリプト
 *
 * 設定手順:
 *   1. Google スプレッドシートを開く
 *   2. メニュー「拡張機能」→「Apps Script」を開く
 *   3. このファイルの内容を貼り付けて保存する
 *   4. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」を選択
 *   5. 実行ユーザー: 自分、アクセスできるユーザー: 全員 に設定してデプロイ
 *   6. デプロイIDを config.json の deploymentId に貼り付ける
 *
 * ヘッダー行の手動作成:
 *   Apps Script エディタで setupHeaders() を選択して実行すると、
 *   シートにヘッダー行が即座に作成されます（初回のみ）。
 */

// ─── 設定 ──────────────────────────────────────────────────────────────────────

const DEFAULT_SHEET_NAME = '協賛申込み';

// フォームデータ列（受付日時 + 12フィールド）＋管理用列
const HEADERS = [
  // フォーム入力項目
  '受付日時',
  '個人名・会社名・団体名',
  '個人名・会社名（フリガナ）',
  '役職・代表者名',
  '役職・代表者名（フリガナ）',
  '担当者名',
  '担当者名（フリガナ）',
  '郵便番号', 
  '住所',
  '電話番号',
  'メールアドレス',
  '区分',
  // 管理用項目（手動入力）
  '請求書送付',
  '入金日',
  '確認者',
  '確認日',
  '受付済み',
  'お礼状送付',
];

// ─── エントリーポイント ────────────────────────────────────────────────────────

/**
 * フォームからPOSTリクエストを受け取り、スプレッドシートに1行追記する。
 * フォームは JSON を URL エンコードされたキー "payload" に包んで送信する。
 * @param {GoogleAppsScript.Events.DoPost} e
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
/** GETリクエストでフォームデータを受け取りシートに書き込む */
function doGet(e) {
  // パラメータなし（動作確認用アクセス）
  if (!e || !e.parameter || !e.parameter.company_name) {
    return jsonResponse({ result: 'ok' });
  }

  try {
    const data      = e.parameter;
    const sheetName = data.sheetName || DEFAULT_SHEET_NAME;
    appendRow(data, sheetName);
    return jsonResponse({ result: 'success' });
  } catch (err) {
    return jsonResponse({ result: 'error', message: err.message });
  }
}

function doPost(e) {
  try {
    const data      = e.parameter;
    const sheetName = data.sheetName || DEFAULT_SHEET_NAME;
    appendRow(data, sheetName);
    return jsonResponse({ result: 'success' });
  } catch (err) {
    return jsonResponse({ result: 'error', message: err.message });
  }
}

// ─── ヘッダー手動セットアップ ─────────────────────────────────────────────────

/**
 * Apps Script エディタから手動実行してヘッダー行を作成する。
 * シートが空の場合のみ実行される。
 */
function setupHeaders() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(DEFAULT_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(`シート "${DEFAULT_SHEET_NAME}" が見つかりません。`);
    return;
  }

  if (sheet.getLastRow() > 0) {
    SpreadsheetApp.getUi().alert('ヘッダー行はすでに存在します。');
    return;
  }

  sheet.appendRow(HEADERS);

  // ヘッダー行を太字・背景色で強調する
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d0e4f7');

  SpreadsheetApp.getUi().alert('ヘッダー行を作成しました。');
}

// ─── 内部処理 ──────────────────────────────────────────────────────────────────

/**
 * スプレッドシートの末尾に申込データを1行追記する。
 * @param {Object} data      - フォームから受け取ったデータオブジェクト
 * @param {string} sheetName - 書き込み先シート名
 */
function appendRow(data, sheetName) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(sheetName);

  if (!sheet) throw new Error(`シート "${sheetName}" が見つかりません。`);

  // ヘッダー行がなければ自動作成する
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#d0e4f7');
  }

  // フォームデータの12列 ＋ 管理用6列（空欄）
  sheet.appendRow([
    new Date(),
    data.company_name    || '',
    data.company_furigana || '',
    data.rep_name        || '',
    data.rep_furigana    || '',
    data.staff_name      || '',
    data.staff_furigana  || '',
    data.zipcode         || '',
    data.address         || '',
    data.phone           || '',
    data.email           || '',
    data.category        || '',
    // 管理用6列は空欄（手動入力）
    '', '', '', '', '', '',
  ]);
}

/**
 * JSON レスポンスを生成する。
 * @param {Object} payload
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

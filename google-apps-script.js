/**
 * Google Apps Script — 川口花火大会 協賛申込みフォーム受信スクリプト
 *
 * 設定手順:
 *   1. Google スプレッドシートを開く
 *   2. メニュー「拡張機能」→「Apps Script」を開く
 *   3. このファイルの内容を貼り付けて保存する
 *   4. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」を選択
 *   5. 実行ユーザー: 自分 / アクセスできるユーザー: 全員（匿名含む）に設定してデプロイ
 *   6. デプロイIDを config.json の deploymentId に貼り付ける
 *
 * メールテンプレート:
 *   mail-template.json を編集するだけでメール文面を変更できます。
 *   subject・body 内の {{変数名}} が自動置換されます。
 *   利用可能な変数: {{company_name}} {{rep_name}} {{staff_name}} {{category}} {{date}}
 */

// ─── 設定 ──────────────────────────────────────────────────────────────────────

const DEFAULT_SHEET_NAME  = '協賛申込み';
const TETSUGYO_SHEET_NAME = '手作業';

// 手作業シートのヘッダー定義
const TETSUGYO_HEADERS = [
  '受付番号',         // A — 自動入力
  '区分',             // B — XLOOKUP
  '電話番号',         // C — XLOOKUP
  '個人名・会社名・団体名', // D — XLOOKUP
  '住所',             // E — XLOOKUP
  '役職・代表者名',   // F — XLOOKUP
  'メールアドレス',   // G — XLOOKUP
  '会社HP URL',       // H — 手動入力
];

// 手作業シート: 協賛申込みシートの参照列（XLOOKUP用）
// キー = 手作業の列番号(1始まり), 値 = 協賛申込みの列アルファベット
const TETSUGYO_LOOKUP_COLS = {
  2: 'M',  // 区分
  3: 'K',  // 電話番号
  4: 'C',  // 個人名・会社名・団体名
  5: 'J',  // 住所
  6: 'E',  // 役職・代表者名
  7: 'L',  // メールアドレス
};

// 事務局情報（返信先メールアドレスに使用）
const OFFICE_EMAIL = 'gioitre.kamifukuoka2023@gmail.com';

// ─── ヘッダー定義 ──────────────────────────────────────────────────────────────

const HEADERS = [
  '受付番号',
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
];

// ─── エントリーポイント ────────────────────────────────────────────────────────

/** GET でフォームデータを受け取りシートに書き込む（CORS回避のため GET を使用） */
function doGet(e) {
  if (!e || !e.parameter || !e.parameter.company_name) {
    return jsonResponse({ result: 'ok' });
  }
  return handleSubmission(e.parameter);
}

/** POST でフォームデータを受け取る（フォールバック） */
function doPost(e) {
  return handleSubmission(e.parameter);
}

/**
 * シートへの書き込みとメール送信を共通処理する。
 * @param {Object} data - e.parameter
 */
function handleSubmission(data) {
  try {
    const receptNo = appendRow(data, data.sheetName || DEFAULT_SHEET_NAME);
    appendToTetsugyoSheet(receptNo);
    if (data.email) sendConfirmationEmail(data, receptNo);
    return jsonResponse({ result: 'success' });
  } catch (err) {
    return jsonResponse({ result: 'error', message: err.message });
  }
}

// ─── メールテンプレート初期設定 ───────────────────────────────────────────────

/**
 * メールテンプレートを Script Properties に保存する。
 * Apps Script エディタでこの関数を選択して▶ Run で初回設定完了。
 * 以後は「プロジェクトの設定」→「スクリプト プロパティ」から直接編集可能。
 */
function setupMailTemplate() {
  const subject = '【第5回川口花火大会】協賛お申し込み受付のご連絡（受付番号：{{receipt_no}}）';

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
    '下記の内容にてお申し込みを受け付けいたしました。',
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
    '今後の流れにつきましては、改めて担当者よりご連絡を差し上げます。',
    'ご請求書の送付をもって正式なお手続きとなりますので、',
    '今しばらくお待ちくださいますようお願い申し上げます。',
    '',
    'ご不明な点がございましたら、お気軽に下記事務局までお問い合わせください。',
    '引き続き、どうぞよろしくお願い申し上げます。',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
    '第5回川口花火大会 実行委員会 事務局',
    'TEL　　：048-XXX-XXXX',
    'E-mail：' + OFFICE_EMAIL,
    '受付時間：平日 10:00 〜 17:00',
    '━━━━━━━━━━━━━━━━━━━━━━━━',
    '',
    '※ このメールは自動送信されています。',
    '　 本メールへの返信はお受けできませんのでご了承ください。',
  ].join('\n');

  PropertiesService.getScriptProperties().setProperties({
    MAIL_SUBJECT: subject,
    MAIL_BODY:    body,
  });

  SpreadsheetApp.getUi().alert(
    'メールテンプレートを保存しました。\n\n' +
    '編集するには：\n「プロジェクトの設定」→「スクリプト プロパティ」\n' +
    '→ MAIL_SUBJECT / MAIL_BODY を直接編集してください。'
  );
}

// ─── ヘッダー手動セットアップ ─────────────────────────────────────────────────

/** Apps Script エディタから手動実行してヘッダー行を作成する（初回のみ）*/
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

// ─── 列幅設定 ──────────────────────────────────────────────────────────────────

function applyColumnWidths(sheet) {
  const widths = [160,150,200,180,150,150,120,120,90,220,120,220,60];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

// ─── 手作業シート書き込み ──────────────────────────────────────────────────────

/**
 * 手作業シートに受付番号を追記し、各列にXLOOKUP数式を挿入する。
 * @param {string} receptNo - 受付番号
 */
function appendToTetsugyoSheet(receptNo) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(TETSUGYO_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(TETSUGYO_SHEET_NAME);

  // ヘッダー行の初期化（初回のみ）
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(TETSUGYO_HEADERS);
    const hRange = sheet.getRange(1, 1, 1, TETSUGYO_HEADERS.length);
    hRange.setFontWeight('bold').setBackground('#fce8b2');  // 黄色ヘッダー
    // 列幅設定
    [160, 80, 120, 200, 200, 150, 200, 160].forEach((w, i) =>
      sheet.setColumnWidth(i + 1, w)
    );
    // 電話番号列（C=3列目）をテキスト形式に設定（先頭0を保持）
    sheet.getRange(2, 3, sheet.getMaxRows() - 1).setNumberFormat('@');
  }

  const newRow = sheet.getLastRow() + 1;

  // A列: 受付番号を直接書き込み
  sheet.getRange(newRow, 1).setValue(receptNo);

  // 電話番号セル（C列）をテキスト形式に設定（先頭0保持）
  sheet.getRange(newRow, 3).setNumberFormat('@');

  // B〜G列: XLOOKUPで協賛申込みシートから自動参照
  Object.entries(TETSUGYO_LOOKUP_COLS).forEach(([col, srcCol]) => {
    const formula =
      `=IFERROR(XLOOKUP($A${newRow};'協賛申込み'!$A:$A;'協賛申込み'!$${srcCol}:$${srcCol});"見つかりません")`;
    sheet.getRange(newRow, Number(col)).setFormula(formula);
  });

  // H列: 会社HP URL は手動入力のため空欄
  // （何も書かない）
}

// ─── シート書き込み ────────────────────────────────────────────────────────────

function appendRow(data, sheetName) {
  // LockService で同時アクセスを防ぐ（最大10秒待機）
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // シートが存在しない場合は自動作成する
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#d0e4f7');

      // 郵便番号（9列目）・電話番号（11列目）列全体をテキスト形式に設定
      sheet.getRange(2, 9,  sheet.getMaxRows() - 1).setNumberFormat('@');
      sheet.getRange(2, 11, sheet.getMaxRows() - 1).setNumberFormat('@');

      // 列幅は初回のみ設定（毎回 resize すると遅いため）
      applyColumnWidths(sheet);
    }

    // 受付番号: クライアントから送られた値を優先し、なければサーバーで生成
    const now      = new Date();
    const receptNo = data.receipt_no ||
                     'KWGC' + Utilities.formatDate(now, 'Asia/Tokyo', 'MMddHHmmssSSS');

    // 追記行番号を確定し、フォーマットを先に設定してから一括書き込み
    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 9).setNumberFormat('@');   // 郵便番号
    sheet.getRange(newRow, 11).setNumberFormat('@');  // 電話番号

    // setValues で1回のAPI呼び出しにまとめる（appendRow より高速）
    const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    sheet.getRange(newRow, 1, 1, HEADERS.length).setValues([[
      receptNo,
      dateStr,
      data.company_name     || '',
      data.company_furigana || '',
      data.rep_name         || '',
      data.rep_furigana     || '',
      data.staff_name       || '',
      data.staff_furigana   || '',
      data.zipcode          || '',
      data.address          || '',
      data.phone            || '',
      data.email            || '',
      data.category         || '',
    ]]);

    // 書き込みを確定してからロックを解放
    SpreadsheetApp.flush();

    return receptNo;
  } finally {
    lock.releaseLock();
  }
}

// ─── 自動返信メール ─────────────────────────────────────────────────────────────

/**
 * mail-template.json を取得し {{変数名}} を実データで置換してメール送信する。
 * テンプレート取得失敗時はシンプルなフォールバック文面を使用する。
 * @param {Object} data - フォームデータ
 */
function sendConfirmationEmail(data, receptNo) {
  const now  = new Date();
  const date = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  // 置換変数マップ
  const vars = {
    company_name: data.company_name || '',
    rep_name:     data.rep_name     || '',
    staff_name:   data.staff_name   || '',
    category:     data.category     || '',
    date:         date,
    receipt_no:   receptNo          || '',
  };

  // ── Script Properties からテンプレートを取得 ──
  const props   = PropertiesService.getScriptProperties();
  const subject_tpl = props.getProperty('MAIL_SUBJECT');
  const body_tpl    = props.getProperty('MAIL_BODY');

  // プロパティ未設定の場合はフォールバック文面を使用
  let subject = subject_tpl || '【第5回川口花火大会】協賛お申し込み受付のご連絡';
  let body    = body_tpl    || '{{company_name}}\n{{rep_name}} 様\n\nお申し込みを受け付けました。\nご確認のほどよろしくお願い申し上げます。\n\n川口花火大会実行委員会 事務局';

  // ── {{変数名}} を実データに置換 ──
  Object.entries(vars).forEach(([key, value]) => {
    const placeholder = new RegExp(`\\{\\{${key}\\}\\}`, 'g');
    subject = subject.replace(placeholder, value);
    body    = body.replace(placeholder, value);
  });

  MailApp.sendEmail({
    to:      data.email,
    subject: subject,
    body:    body,
    replyTo: OFFICE_EMAIL,
  });
}

// ─── ユーティリティ ────────────────────────────────────────────────────────────

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

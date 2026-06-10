// ─── 管理ID番号 管理 ────────────────────────────────────────────────────────────
// 管理IDリスト シート構成（13列）:
//   A: 管理ID番号
//   B: 登録日時
//   C: 個人名・会社名・団体名
//   D: 個人名・会社名（フリガナ）
//   E: 役職・代表者名
//   F: 役職・代表者名（フリガナ）
//   G: 担当者名
//   H: 担当者名（フリガナ）
//   I: 郵便番号
//   J: 住所
//   K: 電話番号
//   L: メールアドレス
//   M: 会社HP URL

const KANRI_SS_ID   = '12SP_fszizLYubFvhI0JYcgm_LUDesG4UAkG_Wi_uHvM';
const KANRI_SHEET   = '管理IDリスト';
const KANRI_LOG_SHEET = 'Log';
const KANRI_PREFIX  = 'KRNO';

const KANRI_HEADERS = [
  '管理ID番号', '登録日時',
  '個人名・会社名・団体名', '個人名・会社名（フリガナ）',
  '役職・代表者名', '役職・代表者名（フリガナ）',
  '担当者名', '担当者名（フリガナ）',
  '郵便番号', '住所', '電話番号', 'メールアドレス', '会社HP URL',
];

/**
 * 管理IDリストシートを初期化する（初回セットアップ・復旧用）
 * GAS エディタまたは initProject から呼び出す
 */
function initKanriSheet() {
  try {
    const ss = SpreadsheetApp.openById(KANRI_SS_ID);
    let sheet = ss.getSheetByName(KANRI_SHEET);

    if (sheet) {
      // 既に存在する場合はヘッダーのみ確認・修正
      const firstCell  = String(sheet.getRange(1, 1).getValue()).trim();
      const secondCell = String(sheet.getRange(1, 2).getValue()).trim();
      if (firstCell !== '管理ID番号' || secondCell !== '登録日時') {
        sheet.insertRowBefore(1);
        sheet.getRange(1, 1, 1, KANRI_HEADERS.length)
          .setValues([KANRI_HEADERS])
          .setFontWeight('bold').setBackground('#d0e4f7');
        sheet.setFrozenRows(1);
        _ensureFilter(sheet, 1, KANRI_HEADERS.length);
        _applyAlignment(sheet, 2, KANRI_HEADERS.length);
        console.log('✅ 管理IDリスト: ヘッダーを修正しました。');
      } else {
        console.log('✅ 管理IDリスト: シートは正常です。');
      }
    } else {
      // シートが存在しない → 新規作成
      sheet = ss.insertSheet(KANRI_SHEET);
      sheet.appendRow(KANRI_HEADERS);
      sheet.getRange(1, 1, 1, KANRI_HEADERS.length)
        .setFontWeight('bold').setBackground('#d0e4f7');
      sheet.setFrozenRows(1);
      _ensureFilter(sheet, 1, KANRI_HEADERS.length);
      _applyAlignment(sheet, 2, KANRI_HEADERS.length);
      console.log('✅ 管理IDリスト: シートを新規作成しました。');
    }

    // Log シートも確認・作成
    let logSheet = ss.getSheetByName(KANRI_LOG_SHEET);
    if (!logSheet) {
      logSheet = ss.insertSheet(KANRI_LOG_SHEET);
      logSheet.appendRow(KANRI_LOG_HEADERS);
      logSheet.getRange(1, 1, 1, KANRI_LOG_HEADERS.length)
        .setFontWeight('bold').setBackground('#d9ead3');
      [160, 100, 160, 220, 160].forEach((w, i) => logSheet.setColumnWidth(i + 1, w));
      logSheet.setFrozenRows(1);
      _ensureFilter(logSheet, 1, KANRI_LOG_HEADERS.length);
      _applyAlignment(logSheet, 2, KANRI_LOG_HEADERS.length);
      console.log('✅ 管理IDリスト Log: シートを新規作成しました。');
    }

    // メニューから実行された場合はアラート表示
    try {
      SpreadsheetApp.getUi().alert('✅ 管理IDリストシートの準備が完了しました。\nファイルID: ' + KANRI_SS_ID);
    } catch (_) { /* GAS エディタから実行時は無視 */ }

    return sheet;
  } catch (e) {
    console.error('initKanriSheet エラー:', e.message);
    try {
      SpreadsheetApp.getUi().alert('❌ エラー: ' + e.message);
    } catch (_) {}
    return null;
  }
}

function getKanriSheet() {
  try {
    const ss    = SpreadsheetApp.openById(KANRI_SS_ID);
    let   sheet = ss.getSheetByName(KANRI_SHEET);

    if (!sheet) {
      // シート自体が存在しない → 新規作成（書式含む）
      sheet = ss.insertSheet(KANRI_SHEET);
      sheet.appendRow(KANRI_HEADERS);
      sheet.getRange(1, 1, 1, KANRI_HEADERS.length)
        .setFontWeight('bold').setBackground('#d0e4f7');
      sheet.setFrozenRows(1);
      _ensureFilter(sheet, 1, KANRI_HEADERS.length);
      _applyAlignment(sheet, 2, KANRI_HEADERS.length);
      return sheet;
    }

    // ヘッダーがなければ先頭に挿入（書式も一緒に設定）
    // 列構成確認: A=管理ID番号, B=登録日時
    const firstCell  = String(sheet.getRange(1, 1).getValue()).trim();
    const secondCell = String(sheet.getRange(1, 2).getValue()).trim();
    if (firstCell !== '管理ID番号' || secondCell !== '登録日時') {
      sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, KANRI_HEADERS.length)
        .setValues([KANRI_HEADERS])
        .setFontWeight('bold').setBackground('#d0e4f7');
      sheet.setFrozenRows(1);
      _ensureFilter(sheet, 1, KANRI_HEADERS.length);
    }

    // ✅ 通常時はここで返す（書式操作なし）
    return sheet;
  } catch (e) {
    console.error('getKanriSheet エラー:', e.message);
    return null;
  }
}

// ─── 行データ → オブジェクト変換 ────────────────────────────────────────────────

function _rowToKanri(row) {
  return {
    found:            true,
    kanri_id:         row[0],
    registered_at:    row[1],
    company_name:     row[2],
    company_furigana: row[3],
    rep_name:         row[4],
    rep_furigana:     row[5],
    staff_name:       row[6],
    staff_furigana:   row[7],
    zipcode:          row[8],
    address:          row[9],
    phone:            row[10],
    email:            row[11],
    website_url:      row[12],
  };
}

function _dataToRow(kanriId, data, registeredAt) {
  return [
    kanriId,
    registeredAt || nowStr(),
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
    data.website_url      || '',
  ];
}

// ─── IDで情報を呼び出す（クライアントから呼び出し） ────────────────────────────

function lookupKanriId(kanriId) {
  if (!kanriId) return { found: false };
  const sheet = getKanriSheet();
  if (!sheet) return { found: false, error: 'シートが見つかりません' };

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(kanriId).trim()) {
      return _rowToKanri(rows[i]);
    }
  }
  return { found: false };
}

// ─── IDを忘れた場合（会社名で検索→最新IDをメール送信） ─────────────────────────

function sendForgottenId(companyName) {
  if (!companyName) return { found: false };
  const sheet = getKanriSheet();
  if (!sheet) return { found: false, error: 'シートが見つかりません' };

  const rows = sheet.getDataRange().getValues();

  // 会社名が含まれる行を全て収集
  const matches = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).includes(companyName)) {
      matches.push(rows[i]);
    }
  }

  if (matches.length === 0) return { found: false, reason: 'not_found' };

  // 複数ヒット → 登録日時（B列 = index 1）が最新の行を採用
  matches.sort((a, b) => {
    const tA = a[1] ? new Date(a[1]).getTime() : 0;
    const tB = b[1] ? new Date(b[1]).getTime() : 0;
    return tB - tA; // 降順（新しい順）
  });

  const row     = matches[0];
  const email   = String(row[11]).trim(); // L列: メール
  const kanriId = String(row[0]).trim();
  if (!email) return { found: false, reason: 'no_email' };

  const eventName = getEventName();
  MailApp.sendEmail({
    to:      email,
    cc:      getOfficeEmail(),
    subject: `【${eventName}】管理ID番号のご連絡`,
    body: [
      `${row[2]}`,
      `${row[4]} 様`,  // E列: 役職・代表者名
      '',
      `【${eventName}】へのご協賛登録情報についてお問い合わせをいただきました。`,
      '',
      `■ 管理ID番号: ${kanriId}`,
      '',
      '次回以降のお申し込みの際は、上記IDをご入力いただくと情報が自動入力されます。',
      '',
      '━━━━━━━━━━━━━━━━━━━━━━━━',
      `${eventName} 実行委員会 事務局`,
      `E-mail：${getOfficeEmail()}`,
      `受付時間：${getOfficeHours()}`,
      '━━━━━━━━━━━━━━━━━━━━━━━━',
    ].join('\n'),
  });

  const maskedEmail = email.replace(/^(.{1,3})(.*)(@.*)$/, (_, a, b, c) =>
    a + b.replace(/./g, '*') + c
  );
  return {
    found:        true,
    maskedEmail,
    multipleHits: matches.length > 1,  // 複数ヒットを通知
  };
}

// ─── 新規ID発行（重複防止付き） ────────────────────────────────────────────────

function generateKanriId() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const year       = String(new Date().getFullYear()); // "2026"
    const yearPrefix = `${KANRI_PREFIX}${year}`;         // "KRNO2026"
    const sheet = getKanriSheet();
    if (!sheet) return `${yearPrefix}${String(1).padStart(4, '0')}`;
    const rows  = sheet.getDataRange().getValues();
    let maxNum  = 0;
    for (let i = 1; i < rows.length; i++) {
      const id = String(rows[i][0]);
      if (id.startsWith(yearPrefix)) {
        const num = parseInt(id.slice(yearPrefix.length));
        if (!isNaN(num) && num > maxNum) maxNum = num;
      }
    }
    return `${yearPrefix}${String(maxNum + 1).padStart(4, '0')}`;
  } finally {
    lock.releaseLock();
  }
}

// ─── IDリストに保存・更新 ────────────────────────────────────────────────────────

/**
 * @param {string} kanriId
 * @param {object} data - 全フォームフィールド
 * @param {boolean} forceUpdate - true: 既存行を強制上書き（登録日時は保持）
 * @returns {{ isNew: boolean, changed: string[] }}
 */
function saveKanriId(kanriId, data, forceUpdate) {
  const sheet = getKanriSheet();
  if (!sheet) throw new Error('管理IDリストシートが見つかりません');

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === kanriId) {
      if (forceUpdate) {
        // 登録日時（B列 = index 1）は元の値を保持
        const newRow = _dataToRow(kanriId, data, rows[i][1] || nowStr());
        sheet.getRange(i + 1, 1, 1, KANRI_HEADERS.length).setValues([newRow]);
        _writeKanriLog(sheet.getParent(), kanriId, '更新', data.company_name, data.rep_name);
        return { isNew: false, changed: [] };
      }
      // 変更点を検出（C列〜M列 = index 2〜12）
      const labels = [
        '会社名', '会社名フリガナ', '役職・代表者名', '役職フリガナ',
        '担当者名', '担当者フリガナ', '郵便番号', '住所', '電話番号',
        'メールアドレス', 'HP URL',
      ];
      const newRow = _dataToRow(kanriId, data, rows[i][1] || nowStr());
      const changed = [];
      for (let c = 2; c <= 12; c++) {
        if (String(rows[i][c]).trim() !== String(newRow[c]).trim()) {
          changed.push(`${labels[c - 2]}: 「${rows[i][c]}」→「${newRow[c]}」`);
        }
      }
      return { isNew: false, changed };
    }
  }

  // 新規登録前に会社名の重複チェック（C列 = index 2）
  const companyName = String(data.company_name || '').trim();
  if (companyName) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][2]).trim() === companyName) {
        throw new Error(`この会社名はすでに登録されています（管理ID：${rows[i][0]}）。既存の管理ID番号をご入力ください。`);
      }
    }
  }
  // 新規: 登録日時付きで保存
  const newRow = _dataToRow(kanriId, data, nowStr());
  sheet.appendRow(newRow);
  _writeKanriLog(sheet.getParent(), kanriId, '新規登録', data.company_name, data.rep_name);
  return { isNew: true, changed: [] };
}

// ─── 会社名の重複チェック（クライアントからのリアルタイムチェック） ──────────────

/**
 * 会社名がすでに登録されているか確認する
 * @returns {{ exists: boolean, kanri_id_hint?: string }}
 */
function checkCompanyExists(companyName) {
  if (!companyName) return { exists: false };
  const sheet = getKanriSheet();
  if (!sheet) return { exists: false };
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).trim() === String(companyName).trim()) {  // C列 = index 2
      const id = String(rows[i][0]);
      // IDの末尾4桁のみマスクして返す（例: KRNO2026****）
      const hint = id.length > 4 ? id.slice(0, -4) + '****' : '****';
      return { exists: true, kanri_id_hint: hint };
    }
  }
  return { exists: false };
}


// ─── 管理IDリスト Log シート書き込み ─────────────────────────────────────────────

const KANRI_LOG_HEADERS = ['管理ID番号', 'ステータス', '日時', '会社名', '代表者'];

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - 管理IDリストのスプレッドシート
 * @param {string} kanriId
 * @param {string} action  - '新規登録' | '更新'
 * @param {string} companyName
 * @param {string} repName
 */
function _writeKanriLog(ss, kanriId, action, companyName, repName) {
  try {
    let sheet = ss.getSheetByName(KANRI_LOG_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(KANRI_LOG_SHEET);
    }

    // ヘッダーが未設定なら挿入
    const firstCell = String(sheet.getRange(1, 1).getValue()).trim();
    if (firstCell !== '管理ID番号') {
      if (sheet.getLastRow() > 0) sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, KANRI_LOG_HEADERS.length).setValues([KANRI_LOG_HEADERS]);
      sheet.getRange(1, 1, 1, KANRI_LOG_HEADERS.length)
        .setFontWeight('bold').setBackground('#d9ead3');
      [160, 100, 160, 220, 160].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
      sheet.setFrozenRows(1);
      _ensureFilter(sheet, 1, KANRI_LOG_HEADERS.length);
      _applyAlignment(sheet, 2, KANRI_LOG_HEADERS.length);
    }

    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    sheet.appendRow([kanriId, action, now, companyName || '', repName || '']);
  } catch (e) {
    console.warn('KanriLog 書き込みエラー:', e.message);
  }
}

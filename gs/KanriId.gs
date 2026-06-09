// ─── 管理ID番号 管理 ────────────────────────────────────────────────────────────
// 管理IDリスト シート構成（13列）:
//   A: 管理ID番号
//   B: 個人名・会社名・団体名
//   C: 個人名・会社名（フリガナ）
//   D: 役職・代表者名
//   E: 役職・代表者名（フリガナ）
//   F: 担当者名
//   G: 担当者名（フリガナ）
//   H: 郵便番号
//   I: 住所
//   J: 電話番号
//   K: メールアドレス
//   L: 会社HP URL
//   M: 登録日時

const KANRI_SS_ID  = '12SP_fszizLYubFvhI0JYcgm_LUDesG4UAkG_Wi_uHvM';
const KANRI_SHEET  = '管理IDリスト';
const KANRI_PREFIX = 'KRNO';

const KANRI_HEADERS = [
  '管理ID番号',
  '個人名・会社名・団体名', '個人名・会社名（フリガナ）',
  '役職・代表者名', '役職・代表者名（フリガナ）',
  '担当者名', '担当者名（フリガナ）',
  '郵便番号', '住所', '電話番号', 'メールアドレス', '会社HP URL',
  '登録日時',
];

function getKanriSheet() {
  try {
    const ss    = SpreadsheetApp.openById(KANRI_SS_ID);
    let   sheet = ss.getSheetByName(KANRI_SHEET);

    if (!sheet) {
      // シート自体が存在しない → 新規作成
      sheet = ss.insertSheet(KANRI_SHEET);
    }

    // ヘッダー行がなければ先頭に挿入
    const firstCell = String(sheet.getRange(1, 1).getValue()).trim();
    if (firstCell !== '管理ID番号') {
      sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, KANRI_HEADERS.length).setValues([KANRI_HEADERS]);
    }
    // 書式を常に適用（冪等）
    sheet.getRange(1, 1, 1, KANRI_HEADERS.length)
      .setFontWeight('bold').setBackground('#d0e4f7');
    sheet.setFrozenRows(1);

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
    company_name:     row[1],
    company_furigana: row[2],
    rep_name:         row[3],
    rep_furigana:     row[4],
    staff_name:       row[5],
    staff_furigana:   row[6],
    zipcode:          row[7],
    address:          row[8],
    phone:            row[9],
    email:            row[10],
    website_url:      row[11],
    registered_at:    row[12],
  };
}

function _dataToRow(kanriId, data) {
  return [
    kanriId,
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
    if (String(rows[i][1]).includes(companyName)) {
      matches.push(rows[i]);
    }
  }

  if (matches.length === 0) return { found: false, reason: 'not_found' };

  // 複数ヒット → 登録日時（M列 = index 12）が最新の行を採用
  matches.sort((a, b) => {
    const tA = a[12] ? new Date(a[12]).getTime() : 0;
    const tB = b[12] ? new Date(b[12]).getTime() : 0;
    return tB - tA; // 降順（新しい順）
  });

  const row     = matches[0];
  const email   = String(row[10]).trim(); // K列: メール
  const kanriId = String(row[0]).trim();
  if (!email) return { found: false, reason: 'no_email' };

  const eventName = getEventName();
  MailApp.sendEmail({
    to:      email,
    cc:      getOfficeEmail(),
    subject: `【${eventName}】管理ID番号のご連絡`,
    body: [
      `${row[1]}`,
      `${row[3]} 様`,  // D列: 役職・代表者名
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

  const rows   = sheet.getDataRange().getValues();
  const newRow = _dataToRow(kanriId, data);

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === kanriId) {
      if (forceUpdate) {
        // 登録日時（M列）は元の値を保持
        newRow.push(rows[i][12] || nowStr());
        sheet.getRange(i + 1, 1, 1, KANRI_HEADERS.length).setValues([newRow]);
        return { isNew: false, changed: [] };
      }
      // 変更点を検出（区分以外のフィールド）
      const labels = [
        '会社名', '会社名フリガナ', '役職・代表者名', '役職フリガナ',
        '担当者名', '担当者フリガナ', '郵便番号', '住所', '電話番号',
        'メールアドレス', 'HP URL',
      ];
      const changed = [];
      for (let c = 1; c <= 11; c++) {
        if (String(rows[i][c]).trim() !== String(newRow[c]).trim()) {
          changed.push(`${labels[c - 1]}: 「${rows[i][c]}」→「${newRow[c]}」`);
        }
      }
      return { isNew: false, changed };
    }
  }

  // 新規登録前に会社名の重複チェック
  const companyName = String(data.company_name || '').trim();
  if (companyName) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim() === companyName) {
        throw new Error(`この会社名はすでに登録されています（管理ID：${rows[i][0]}）。既存の管理ID番号をご入力ください。`);
      }
    }
  }
  // 新規: 登録日時を追加して保存
  newRow.push(nowStr());
  sheet.appendRow(newRow);
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
    if (String(rows[i][1]).trim() === String(companyName).trim()) {
      const id = String(rows[i][0]);
      // IDの末尾4桁のみマスクして返す（例: KRNO2026****）
      const hint = id.length > 4 ? id.slice(0, -4) + '****' : '****';
      return { exists: true, kanri_id_hint: hint };
    }
  }
  return { exists: false };
}

// ─── 既存IDの変更を検出（submitForm からの呼び出し用） ─────────────────────────

function detectKanriChanges(kanriId, data) {
  if (!kanriId) return [];
  const existing = lookupKanriId(kanriId);
  if (!existing.found) return [];

  const fields = [
    ['会社名',           'company_name'],
    ['会社名フリガナ',   'company_furigana'],
    ['役職・代表者名',   'rep_name'],
    ['役職フリガナ',     'rep_furigana'],
    ['担当者名',         'staff_name'],
    ['担当者フリガナ',   'staff_furigana'],
    ['郵便番号',         'zipcode'],
    ['住所',             'address'],
    ['電話番号',         'phone'],
    ['メールアドレス',   'email'],
    ['HP URL',           'website_url'],
  ];
  return fields
    .filter(([, key]) => String(existing[key] || '').trim() !== String(data[key] || '').trim())
    .map(([label, key]) => `${label}: 「${existing[key] || ''}」→「${data[key] || ''}」`);
}

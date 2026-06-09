// ─── 定数・設定 ────────────────────────────────────────────────────────────────

const DEFAULT_SHEET_NAME  = '協賛申込み一覧';
const DEFAULT_SHEET_NAME2 = '手作業';
const INFO_SHEET_NAME     = 'Info';
const CREATELOG_SHEET     = 'CreateLog';

// ─── 手作業 列番号定数 ──────────────────────────────────────────────────────────
const COL_RECEPT_NO    =  1; // A: 受付番号
const COL_UKETSUKE     =  9; // I: 受付完了         checkbox  手動
const COL_INV_DATE     = 10; // J: 請求書送信日時    timestamp 自動
const COL_NYUKIN       = 11; // K: 入金完了          checkbox  手動
const COL_SEAT_DATE    = 12; // L: 座席割当送信日時  timestamp 自動
const COL_SEAT_NO      = 13; // M: 座席番号          text      自動（区分＋7桁）
const COL_ANNAIBUN     = 14; // N: 案内実施          checkbox  手動
const COL_ANNAI_DATE   = 15; // O: 案内送信日時      timestamp 自動
const COL_OREIJOU_DATE = 16; // P: お礼状送信日時    timestamp 自動

// ─── デフォルト価格（Info シートの PRICE_X で上書き可） ────────────────────────
const DEFAULT_PRICES = { S: 2000000, A: 1000000, B: 500000, C: 300000, D: 200000, E: 100000 };

// B〜E は申込時に自動で請求書送信
const AUTO_SEND_KUBUN = ['B', 'C', 'D', 'E'];

// フォールバック定数（Info シートで上書き可）
const HANKO_FILE_ID_DEFAULT = '1Q5kBbhKKuIRSaNiSXfBbvTnO0W1niFyZ';
const OFFICE_EMAIL_DEFAULT  = 'gioitre.kamifukuoka2023@gmail.com';

// ─── シートヘッダー ─────────────────────────────────────────────────────────────

const HEADERS = [
  '受付番号', '受付日時',
  '個人名・会社名・団体名', '個人名・会社名（フリガナ）',
  '役職・代表者名', '役職・代表者名（フリガナ）',
  '担当者名', '担当者名（フリガナ）',
  '郵便番号', '住所', '電話番号', 'メールアドレス', '区分', '会社HP URL',
];

const TESAGYOU_HEADERS = [
  '受付番号',               // A (1)
  '区分',                   // B (2)  XLOOKUP
  '電話番号',               // C (3)  XLOOKUP
  '個人名・会社名・団体名', // D (4)  XLOOKUP
  '住所',                   // E (5)  XLOOKUP
  '役職・代表者名',         // F (6)  XLOOKUP
  'メールアドレス',         // G (7)  XLOOKUP
  '会社HP URL',             // H (8)  XLOOKUP
  '受付完了',               // I (9)  checkbox 手動
  '請求書送信日時',         // J (10) timestamp 自動
  '入金完了',               // K (11) checkbox 手動
  '座席割当送信日時',       // L (12) timestamp 自動
  '座席番号',               // M (13) text 自動（区分＋7桁）
  '案内実施',               // N (14) checkbox 手動
  '案内送信日時',           // O (15) timestamp 自動
  'お礼状送信日時',         // P (16) timestamp 自動
];

// 手作業の列番号 → 協賛申込み一覧の列アルファベット（XLOOKUP）
const TESAGYOU_LOOKUP_COLS = {
  2: 'M',  // 区分
  3: 'K',  // 電話番号
  4: 'C',  // 個人名・会社名・団体名
  5: 'J',  // 住所
  6: 'E',  // 役職・代表者名
  7: 'L',  // メールアドレス
  8: 'N',  // 会社HP URL
};

// ─── Info シート読み込み ────────────────────────────────────────────────────────

let _infoCache = null;

function getInfoConfig() {
  if (_infoCache) return _infoCache;
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(INFO_SHEET_NAME);
    const cfg   = {};
    if (sheet && sheet.getLastRow() > 0) {
      sheet.getDataRange().getValues().forEach(row => {
        if (row[0]) cfg[String(row[0]).trim()] = row[1];
      });
    }
    _infoCache = cfg;
    return cfg;
  } catch (e) {
    _infoCache = {};
    return {};
  }
}

function getConfigVal(key, fallback) {
  const v = getInfoConfig()[key];
  return (v !== undefined && v !== '') ? v : fallback;
}

function getOfficeEmail()  { return getConfigVal('OFFICE_EMAIL',  OFFICE_EMAIL_DEFAULT);  }
function getHankoFileId()  { return getConfigVal('HANKO_FILE_ID', HANKO_FILE_ID_DEFAULT); }
function getEventName()    { return getConfigVal('EVENT_NAME',    '第5回川口花火大会');    }
function getPaymentDue()   { return getConfigVal('PAYMENT_DUE',   '9月12日（金）');        }
function getRootFolderId() { return getConfigVal('ROOT_FOLDER_ID', '');                   }

function getCategoryPrice(category) {
  const key = `PRICE_${String(category || '').trim().toUpperCase()}`;
  const v   = getConfigVal(key, null);
  return v !== null ? Number(v) : (DEFAULT_PRICES[String(category || '').toUpperCase()] || 0);
}

/** Info シートのキーを更新（なければ追記） */
/** データ用スプレッドシートを返す（DATA_SPREADSHEET_ID があればそちら、なければ Main） */
function getDataSpreadsheet() {
  const id = getConfigVal('DATA_SPREADSHEET_ID', '');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* fallthrough */ }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function setInfoValue(key, value) {
  _infoCache = null;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INFO_SHEET_NAME);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value, '（自動設定）']);
}

// ─── ユーティリティ ─────────────────────────────────────────────────────────────

function formatTs(value) {
  if (!value) return '';
  if (value instanceof Date) return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  return String(value);
}

function generateSeatNo(kubun, dataRow) {
  return `${String(kubun || 'X').trim().toUpperCase()}-${String(dataRow - 1).padStart(7, '0')}`;
}

function nowStr() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

// ─── 定数・設定 ────────────────────────────────────────────────────────────────

const DEFAULT_SHEET_NAME  = '協賛申込み一覧';
const DEFAULT_SHEET_NAME2 = '手作業';
const INFO_SHEET_NAME     = 'Info';
const CREATELOG_SHEET     = 'CreateLog';

// ─── 手作業 列番号定数 ──────────────────────────────────────────────────────────
const COL_KANRI_ID     =  1; // A: 管理ID番号        XLOOKUP
const COL_RECEPT_NO    =  2; // B: 受付番号
const COL_UKETSUKE     = 10; // J: 受付完了         checkbox  手動
const COL_INV_DATE     = 11; // K: 請求書送信日時    timestamp 自動
const COL_NYUKIN       = 12; // L: 入金完了          checkbox  手動
const COL_SEAT_DATE    = 13; // M: 座席割当送信日時  timestamp 自動
const COL_SEAT_NO      = 14; // N: 座席番号          text      自動（区分＋7桁）
const COL_ANNAIBUN     = 15; // O: 案内実施          checkbox  手動
const COL_ANNAI_DATE   = 16; // P: 案内送信日時      timestamp 自動
const COL_OREIJOU_DATE = 17; // Q: お礼状送信日時    timestamp 自動

// ─── デフォルト価格（Info シートの PRICE_X で上書き可） ────────────────────────
const DEFAULT_PRICES = { S: 2000000, A: 1000000, B: 500000, C: 300000, D: 200000, E: 100000 };

// B〜E は申込時に自動で請求書送信
const AUTO_SEND_KUBUN = ['B', 'C', 'D', 'E'];

// フォールバック定数（Info シートで上書き可）
const HANKO_FILE_ID_DEFAULT = '1Q5kBbhKKuIRSaNiSXfBbvTnO0W1niFyZ';
const OFFICE_EMAIL_DEFAULT  = 'gioitre.kamifukuoka2023@gmail.com';

// ─── シートヘッダー ─────────────────────────────────────────────────────────────

const HEADERS = [
  '管理ID番号',             // A (1)
  '受付番号', '受付日時',   // B,C (2,3)
  '個人名・会社名・団体名', '個人名・会社名（フリガナ）', // D,E (4,5)
  '役職・代表者名', '役職・代表者名（フリガナ）',         // F,G (6,7)
  '担当者名', '担当者名（フリガナ）',                     // H,I (8,9)
  '郵便番号', '住所', '電話番号', 'メールアドレス', '区分', '会社HP URL', // J-O (10-15)
];

const TESAGYOU_HEADERS = [
  '管理ID番号',             // A (1)  XLOOKUP
  '受付番号',               // B (2)
  '区分',                   // C (3)  XLOOKUP
  '電話番号',               // D (4)  XLOOKUP
  '個人名・会社名・団体名', // E (5)  XLOOKUP
  '住所',                   // F (6)  XLOOKUP
  '役職・代表者名',         // G (7)  XLOOKUP
  'メールアドレス',         // H (8)  XLOOKUP
  '会社HP URL',             // I (9)  XLOOKUP
  '受付完了',               // J (10) checkbox 手動
  '請求書送信日時',         // K (11) timestamp 自動
  '入金完了',               // L (12) checkbox 手動
  '座席割当送信日時',       // M (13) timestamp 自動
  '座席番号',               // N (14) text 自動（区分＋7桁）
  '案内実施',               // O (15) checkbox 手動
  '案内送信日時',           // P (16) timestamp 自動
  'お礼状送信日時',         // Q (17) timestamp 自動
];

// 手作業の列番号 → 協賛申込み一覧の列アルファベット（XLOOKUP）
// ※ 協賛申込み一覧: A=管理ID B=受付番号 C=受付日時 D=個人名 E=個人名ふりがな
//                   F=役職・代表者名 G=役職ふりがな H=担当者名 I=担当者ふりがな
//                   J=郵便番号 K=住所 L=電話番号 M=メール N=区分 O=会社HP URL
const TESAGYOU_LOOKUP_COLS = {
  1:  'A',  // 管理ID番号
  3:  'N',  // 区分
  4:  'L',  // 電話番号
  5:  'D',  // 個人名・会社名・団体名
  6:  'K',  // 住所
  7:  'F',  // 役職・代表者名
  8:  'M',  // メールアドレス
  9:  'O',  // 会社HP URL
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

function getOfficeEmail()      { return getConfigVal('OFFICE_EMAIL',       OFFICE_EMAIL_DEFAULT);              }
function getHankoFileId()      { return getConfigVal('HANKO_FILE_ID',      HANKO_FILE_ID_DEFAULT);             }
function getEventName()        { return getConfigVal('EVENT_NAME',         '第5回川口花火大会');                }
function getPaymentDue()       { return getConfigVal('PAYMENT_DUE',        '9月12日（金）');                   }
function getRootFolderId()     { return getConfigVal('ROOT_FOLDER_ID',     '');                                }
function getOrgName()          { return getConfigVal('ORG_NAME',           '川口花火大会実行委員会');           }
function getOrgRep()           { return getConfigVal('ORG_REP',            '委員長　廣瀬 進治');                }
function getInvoiceRegNo()     { return getConfigVal('INVOICE_REG_NO',     'T9700150122003');                  }
function getBankName()         { return getConfigVal('BANK_NAME',          '埼玉りそな銀行　川口支店');         }
function getBankNo()           { return getConfigVal('BANK_NO',            '36216349');                        }
function getBankHolder()       { return getConfigVal('BANK_HOLDER',        '川口商工会議所');                   }
function getBankRep()          { return getConfigVal('BANK_REP',           '川口花火大会実行委員会委員長　廣瀬 進治'); }
function getOrgLocation()      { return getConfigVal('ORG_LOCATION',       '（川口商工会議所内）');             }
function getOrgTel()           { return getConfigVal('ORG_TEL',            '048-228-2220');                    }
function getOrgFax()           { return getConfigVal('ORG_FAX',            '048-228-2221');                    }
function getOfficeHours()      { return getConfigVal('OFFICE_HOURS',       '平日 10:00 〜 17:00');              }
function getReceiptNoPrefix()  { return getConfigVal('RECEIPT_NO_PREFIX',  'KWGC');                            }

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

/**
 * 座席番号を発行する（区分ごとに 1 から連番）
 * @param {string} kubun - 区分（B/C/D/E など）
 * @param {Sheet}  tesagyouSheet - 手作業シート（既存の座席番号を数えるため）
 */
function generateSeatNo(kubun, tesagyouSheet) {
  const k = String(kubun || 'X').trim().toUpperCase();
  let count = 0;
  if (tesagyouSheet) {
    const lastRow = tesagyouSheet.getLastRow();
    if (lastRow > 2) {
      const vals = tesagyouSheet
        .getRange(3, COL_SEAT_NO, lastRow - 2, 1)
        .getValues();
      vals.forEach(([v]) => {
        if (String(v).startsWith(k + '-')) count++;
      });
    }
  }
  return `${k}-${String(count + 1).padStart(7, '0')}`;
}

function nowStr() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

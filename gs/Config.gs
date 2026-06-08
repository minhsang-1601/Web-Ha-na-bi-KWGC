// ─── 定数・設定 ────────────────────────────────────────────────────────────────

const DEFAULT_SHEET_NAME  = '協賛申込み一覧';
const DEFAULT_SHEET_NAME2 = '手作業';

const CATEGORY_PRICE = { S: 2000000, A: 1000000, B: 700000, C: 300000, D: 200000, E: 100000 };

// S/A は請求書を手動送信、B〜E は申込時に自動送信
const AUTO_SEND_KUBUN = ['B', 'C', 'D', 'E'];

const HANKO_FILE_ID = '1Q5kBbhKKuIRSaNiSXfBbvTnO0W1niFyZ';

const OFFICE_EMAIL = 'gioitre.kamifukuoka2023@gmail.com';

const HEADERS = [
  '受付番号', '受付日時',
  '個人名・会社名・団体名', '個人名・会社名（フリガナ）',
  '役職・代表者名', '役職・代表者名（フリガナ）',
  '担当者名', '担当者名（フリガナ）',
  '郵便番号', '住所', '電話番号', 'メールアドレス', '区分', '会社HP URL',
];

const TESAGYOU_HEADERS = [
  '受付番号',           // A
  '区分',               // B — XLOOKUP
  '電話番号',           // C — XLOOKUP
  '個人名・会社名・団体名', // D — XLOOKUP
  '住所',               // E — XLOOKUP
  '役職・代表者名',     // F — XLOOKUP
  'メールアドレス',     // G — XLOOKUP
  '会社HP URL',         // H — XLOOKUP
  '受付完了',           // I — チェックボックス（S/A: 手動, B〜E: 自動）
  '請求書送信',         // J — 未/済み
  '入金完了',           // K — チェックボックス
  'お礼状送付',         // L — 未/済み
];

// キー = 手作業の列番号(1始まり), 値 = 協賛申込みの列アルファベット
const TESAGYOU_LOOKUP_COLS = {
  2: 'M',  // 区分
  3: 'K',  // 電話番号
  4: 'C',  // 個人名・会社名・団体名
  5: 'J',  // 住所
  6: 'E',  // 役職・代表者名
  7: 'L',  // メールアドレス
  8: 'N',  // 会社HP URL
};

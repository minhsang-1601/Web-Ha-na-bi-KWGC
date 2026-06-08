// ─── 定数・設定 ────────────────────────────────────────────────────────────────

const DEFAULT_SHEET_NAME  = '協賛申込み一覧';
const DEFAULT_SHEET_NAME2 = '手作業';

const CATEGORY_PRICE = { A: 1000000, B: 500000, C: 300000, D: 200000, E: 100000 };

const HANKO_FILE_ID = '1Q5kBbhKKuIRSaNiSXfBbvTnO0W1niFyZ';

const OFFICE_EMAIL = 'gioitre.kamifukuoka2023@gmail.com';

const HEADERS = [
  '受付番号', '受付日時',
  '個人名・会社名・団体名', '個人名・会社名（フリガナ）',
  '役職・代表者名', '役職・代表者名（フリガナ）',
  '担当者名', '担当者名（フリガナ）',
  '郵便番号', '住所', '電話番号', 'メールアドレス', '区分',
];

const TETSUGYO_HEADERS = [
  '受付番号', '区分', '電話番号', '個人名・会社名・団体名',
  '住所', '役職・代表者名', 'メールアドレス', '会社HP URL', '入金完了', 'お礼状送付',
];

// キー = 手作業の列番号(1始まり), 値 = 協賛申込みの列アルファベット
const TETSUGYO_LOOKUP_COLS = {
  2: 'M',  // 区分
  3: 'K',  // 電話番号
  4: 'C',  // 個人名・会社名・団体名
  5: 'J',  // 住所
  6: 'E',  // 役職・代表者名
  7: 'L',  // メールアドレス
};

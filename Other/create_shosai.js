"use strict";
const pptxgen = require("pptxgenjs");

const C = {
  navy:"1a3a5c", blue:"2e6da4", lblue:"dce9f5", llblue:"f0f5fb",
  green:"27ae60", lgreen:"d5f0e0", orange:"e67e22", lorange:"fde9d4",
  red:"c0392b", lred:"fce4e4", white:"FFFFFF", gray:"7f8c8d",
  lgray:"ecf0f1", black:"222222", lyellow:"fef9e7", purple:"7e57c2", lpurple:"ede7f6"
};
const FONT = "Meiryo";
const makeShadow = () => ({ type:"outer", blur:5, offset:2, color:"000000", opacity:0.12 });

function addNavySlide(pres, title, subtitle) {
  const sl = pres.addSlide();
  sl.background = { color: C.navy };
  sl.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:0.18, h:5.625, fill:{color:C.blue} });
  sl.addShape(pres.shapes.RECTANGLE, { x:0, y:5.05, w:10, h:0.575, fill:{color:C.blue}, transparency:60 });
  if (title) sl.addText(title, { x:0.4, y:1.7, w:9.2, h:1.3, fontFace:FONT, fontSize:38, bold:true, color:C.white, valign:"middle" });
  if (subtitle) sl.addText(subtitle, { x:0.4, y:3.1, w:9.2, h:0.7, fontFace:FONT, fontSize:17, color:"a8c4e0" });
  return sl;
}
function addContentSlide(pres, title) {
  const sl = pres.addSlide();
  sl.background = { color: C.white };
  sl.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.95, fill:{color:C.navy} });
  sl.addShape(pres.shapes.RECTANGLE, { x:0, y:0.95, w:10, h:0.06, fill:{color:C.blue} });
  if (title) sl.addText(title, { x:0.35, y:0.08, w:9.3, h:0.79, fontFace:FONT, fontSize:20, bold:true, color:C.white, valign:"middle" });
  return sl;
}
function flowBox(sl, pres, x, y, w, h, label, sub, fill, stroke) {
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill:{color:fill||C.lblue}, line:{color:stroke||C.blue, width:1.5}, rectRadius:0.07, shadow:makeShadow() });
  sl.addText(label, { x:x+0.06, y:y+0.05, w:w-0.12, h:sub?h*0.5:h-0.1, fontFace:FONT, fontSize:10, bold:true, color:C.black, align:"center", valign:"bottom", wrap:true });
  if (sub) sl.addText(sub, { x:x+0.06, y:y+h*0.5, w:w-0.12, h:h*0.45, fontFace:FONT, fontSize:8, color:stroke||C.blue, align:"center", valign:"top", wrap:true });
}
function arrowRight(sl, x, y, len) {
  sl.addShape(pres.shapes.LINE, { x, y, w:len*0.8, h:0, line:{color:C.gray, width:1.5} });
  sl.addText("▶", { x:x+len*0.78, y:y-0.1, w:0.2, h:0.2, fontFace:FONT, fontSize:9, color:C.gray, align:"center" });
}
function arrowDown(sl, x, y, h) {
  sl.addShape(pres.shapes.LINE, { x, y, w:0, h:h*0.8, line:{color:C.gray, width:1.5} });
  sl.addText("▼", { x:x-0.1, y:y+h*0.78, w:0.2, h:0.2, fontFace:FONT, fontSize:9, color:C.gray, align:"center" });
}

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "詳細設計書 — WebHanabi";

// Slide 1: Cover
{
  const sl = pres.addSlide();
  sl.background = { color: C.navy };
  for (const [x,y,r,t] of [[8.5,0.7,2.2,20],[1.0,4.9,1.4,15],[8.8,4.6,1.0,25]]) {
    sl.addShape(pres.shapes.OVAL, { x:x-r/2, y:y-r/2, w:r, h:r, fill:{color:C.blue, transparency:t} });
  }
  sl.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:0.22, h:5.625, fill:{color:C.blue} });
  sl.addText("詳細設計書", { x:0.5, y:1.2, w:9, h:1.6, fontFace:FONT, fontSize:52, bold:true, color:C.white, valign:"middle" });
  sl.addText("WebHanabi  —  川口花火大会 協賛管理システム", { x:0.5, y:3.0, w:9, h:0.65, fontFace:FONT, fontSize:17, color:"a8c4e0" });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:3.8, w:5, h:0.03, fill:{color:C.blue} });
  sl.addText("発行日：2026年06月18日　　バージョン：1.0", { x:0.5, y:3.95, w:9, h:0.45, fontFace:FONT, fontSize:12, color:"8ab4cc" });
  // Code decoration
  const codeLines = ["function doGet(e) {", "  const sheet = getInfoSheet();", "  return HtmlService...", "}"];
  codeLines.forEach((l,i) => {
    sl.addText(l, { x:6.8, y:2.0+i*0.35, w:3.0, h:0.32, fontFace:"Courier New", fontSize:9, color:"334466", transparency:0 });
  });
}

// Slide 2: 目次
{
  const sl = addContentSlide(pres, "目　次");
  const items = [
    ["1","WebApp.gs — エントリーポイント"],
    ["2","Config.gs — 設定管理"],
    ["3","Sheet.gs — スプレッドシート操作"],
    ["4","Mail.gs — メール送信・PDF生成"],
    ["5","Trigger.gs — イベントトリガー"],
    ["6","ProjectInit.gs — プロジェクト初期化"],
    ["7","index.html — Webフォーム詳細"],
    ["8","エラーハンドリング方針"],
  ];
  items.forEach((item,i)=>{
    const col=i<4?0:1, row=i<4?i:i-4;
    const x=0.4+col*4.85, y=1.18+row*1.05;
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w:4.55, h:0.88, fill:{color:i%2===0?C.llblue:C.white}, line:{color:C.lblue,width:0.8}, shadow:makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w:0.52, h:0.88, fill:{color:C.blue} });
    sl.addText(item[0], { x, y, w:0.52, h:0.88, fontFace:FONT, fontSize:22, bold:true, color:C.white, align:"center", valign:"middle" });
    sl.addText(item[1], { x:x+0.6, y:y+0.18, w:3.85, h:0.52, fontFace:FONT, fontSize:12, color:C.navy, valign:"middle" });
  });
}

// Slide 3: WebApp.gs doGet
{
  const sl = addContentSlide(pres, "1.  WebApp.gs — doGet() 処理フロー");
  sl.addText("HTTPGETリクエストを受け取り、設定確認・クォータチェックを行ってフォームHTMLを返す。", {
    x:0.4, y:1.08, w:9.2, h:0.4, fontFace:FONT, fontSize:11, color:C.black
  });
  const steps = [
    {n:"1",label:"設定確認",sub:"getInfoSheet()\nInfoシート取得",fill:C.lblue,stroke:C.blue},
    {n:"2",label:"必須項目\n検証",sub:"EVENT_NAME\nOFFICE_EMAIL 等",fill:C.lgreen,stroke:C.green},
    {n:"3",label:"クォータ\n確認",sub:"_getMailQuotaSafe()\n不足→停止HTML",fill:C.lorange,stroke:C.orange},
    {n:"4",label:"背景画像\n読込",sub:"BG_IMAGE_ID\nBase64変換",fill:C.lblue,stroke:C.blue},
    {n:"5",label:"テンプレート\n生成",sub:"createTemplateFrom\nFile('index')",fill:C.lgreen,stroke:C.green},
    {n:"6",label:"HTML返却",sub:"evaluate()\nsetXFrameOptions",fill:C.lorange,stroke:C.orange},
  ];
  const bw=1.35, bh=1.55, gap=0.18, startX=0.3, y=1.6;
  steps.forEach((s,i)=>{
    const x=startX+i*(bw+gap);
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x, y, w:bw, h:bh, fill:{color:s.fill}, line:{color:s.stroke,width:1.5}, rectRadius:0.07, shadow:makeShadow()});
    sl.addShape(pres.shapes.OVAL, {x:x+bw/2-0.22, y:y+0.06, w:0.44, h:0.44, fill:{color:s.stroke}});
    sl.addText(s.n, {x:x+bw/2-0.22, y:y+0.06, w:0.44, h:0.44, fontFace:FONT, fontSize:14, bold:true, color:C.white, align:"center", valign:"middle"});
    sl.addText(s.label, {x:x+0.06, y:y+0.58, w:bw-0.12, h:0.5, fontFace:FONT, fontSize:11, bold:true, color:C.black, align:"center", wrap:true});
    sl.addText(s.sub,   {x:x+0.06, y:y+1.08, w:bw-0.12, h:0.42, fontFace:"Courier New", fontSize:8, color:s.stroke, align:"center", wrap:true});
    if(i<steps.length-1) arrowRight(sl, x+bw+0.02, y+bh/2, gap+0.02);
  });
  sl.addText("エラー時: 設定不備またはクォータ不足の場合はエラーHTMLを返却してフォームは表示されない", {
    x:0.4, y:3.3, w:9.2, h:0.4, fontFace:FONT, fontSize:10, color:C.orange,
  });
}

// Slide 4: submitFormJson
{
  const sl = addContentSlide(pres, "1.  WebApp.gs — submitFormJson() 処理フロー");
  sl.addText("google.script.run から呼ばれるサーバーサイド関数。JSON文字列でフォームデータを受け取る。", {
    x:0.4, y:1.08, w:9.2, h:0.38, fontFace:FONT, fontSize:11, color:C.black
  });
  const steps = [
    {n:"1",label:"JSONパース",sub:'JSON.parse(jsonStr)',fill:C.lblue,stroke:C.blue},
    {n:"2",label:"サーバー\nバリデーション",sub:"郵便番号7桁\n電話10-11桁",fill:C.lgreen,stroke:C.green},
    {n:"3",label:"データ\n登録",sub:"appendRow(data)\n受付番号取得",fill:C.lorange,stroke:C.orange},
    {n:"4",label:"手作業シート\n登録",sub:"appendToTesagyou\nSheet(receiptNo)",fill:C.lblue,stroke:C.blue},
    {n:"5",label:"メール\n送信判定",sub:"B〜E→自動送信\nS・A→送信なし",fill:C.lgreen,stroke:C.green},
    {n:"6",label:"クォータ\n監視通知",sub:"低残量時に\n管理者通知",fill:C.lorange,stroke:C.orange},
    {n:"7",label:"受付番号\n返却",sub:"文字列を\nクライアントへ",fill:"e8f4fd",stroke:C.blue},
  ];
  const bw=1.2, bh=1.5, gap=0.12, startX=0.27, y=1.6;
  steps.forEach((s,i)=>{
    const x=startX+i*(bw+gap);
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x, y, w:bw, h:bh, fill:{color:s.fill}, line:{color:s.stroke,width:1.5}, rectRadius:0.06, shadow:makeShadow()});
    sl.addShape(pres.shapes.OVAL, {x:x+bw/2-0.2, y:y+0.05, w:0.4, h:0.4, fill:{color:s.stroke}});
    sl.addText(s.n, {x:x+bw/2-0.2, y:y+0.05, w:0.4, h:0.4, fontFace:FONT, fontSize:12, bold:true, color:C.white, align:"center", valign:"middle"});
    sl.addText(s.label, {x:x+0.05, y:y+0.52, w:bw-0.1, h:0.48, fontFace:FONT, fontSize:10, bold:true, color:C.black, align:"center", wrap:true});
    sl.addText(s.sub,   {x:x+0.05, y:y+1.0,  w:bw-0.1, h:0.45, fontFace:"Courier New", fontSize:7.5, color:s.stroke, align:"center", wrap:true});
    if(i<steps.length-1) arrowRight(sl, x+bw+0.01, y+bh/2, gap+0.01);
  });
  sl.addText("※ B〜Eは申込み直後に自動送信。S・A区分は担当者が手作業シートで受付完了チェック後に手動送信。", {
    x:0.4, y:3.26, w:9.2, h:0.38, fontFace:FONT, fontSize:10, color:C.orange
  });
}

// Slide 5: Config.gs
{
  const sl = addContentSlide(pres, "2.  Config.gs — 設定取得の優先順位");
  sl.addText("getConfigVal(key) は3層のフォールバックで値を取得。スプレッドシートへのアクセスを最小化しパフォーマンスを向上。", {
    x:0.4, y:1.08, w:9.2, h:0.4, fontFace:FONT, fontSize:11, color:C.black
  });
  // 3-layer diagram
  const layers = [
    {n:"1",label:"Script Properties",sub:"キャッシュ（高速）",note:"ScriptApp.getScriptProperties()\n.getProperty(key)",fill:C.lgreen,stroke:C.green},
    {n:"2",label:"Info シート",sub:"スプレッドシート（直接読込）",note:"getInfoSheet().createTextFinder(key)\n.findNext()",fill:C.lorange,stroke:C.orange},
    {n:"3",label:"デフォルト値",sub:"引数 defaultVal",note:"呼び出し元が指定した\nフォールバック値",fill:C.lred,stroke:C.red},
  ];
  layers.forEach((l,i)=>{
    const y=1.65+i*1.22;
    const w=3.8-i*0.3, x=(10-w)/2;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x, y, w, h:1.0, fill:{color:l.fill}, line:{color:l.stroke,width:2}, rectRadius:0.07, shadow:makeShadow()});
    sl.addShape(pres.shapes.OVAL, {x:x+0.1, y:y+0.28, w:0.44, h:0.44, fill:{color:l.stroke}});
    sl.addText(l.n, {x:x+0.1, y:y+0.28, w:0.44, h:0.44, fontFace:FONT, fontSize:14, bold:true, color:C.white, align:"center", valign:"middle"});
    sl.addText(l.label, {x:x+0.65, y:y+0.08, w:w-0.75, h:0.38, fontFace:FONT, fontSize:14, bold:true, color:C.black});
    sl.addText(l.sub,   {x:x+0.65, y:y+0.46, w:w-0.75, h:0.28, fontFace:FONT, fontSize:10, color:l.stroke});
    // Right side code
    sl.addText(l.note, {x:5.8, y:y+0.1, w:3.8, h:0.8, fontFace:"Courier New", fontSize:8.5, color:C.gray, wrap:true});
    if(i<layers.length-1) {
      sl.addText("値が見つからない場合 ↓", {x:3.8, y:y+1.06, w:2.8, h:0.18, fontFace:FONT, fontSize:9, color:C.gray, align:"center"});
    }
  });
}

// Slide 6: Config keys
{
  const sl = addContentSlide(pres, "2.  Config.gs — 主要設定キー一覧");
  const keys = [
    ["EVENT_NAME","イベント名（フォームタイトル等に使用）"],
    ["OFFICE_EMAIL","事務局メールアドレス（CC・通知先）"],
    ["ORG_NAME / ORG_REP","主催団体名・代表者名"],
    ["START_DATE / END_DATE","イベント開催日時"],
    ["PAYMENT_DUE","入金期限日"],
    ["BANK_NAME / BANK_NO\nBANK_HOLDER / BANK_REP","振込先銀行情報（4項目）"],
    ["PRICE_S 〜 PRICE_E","各区分の協賛金額（税抜・円）"],
    ["KUBUN_SA_START/END","S・A区分の申込み受付期間"],
    ["KUBUN_BCDE_START/END","B〜E区分の申込み受付期間"],
    ["BG_IMAGE_ID","背景画像のGoogle Drive ファイルID"],
    ["MIN_MAIL_QUOTA","メール送信停止閾値（残通数）デフォルト:5"],
    ["INVOICE_REG_NO","インボイス登録番号（T+13桁）"],
  ];
  keys.forEach((k,i)=>{
    const col=i<6?0:1, row=i<6?i:i-6;
    const x=0.35+col*4.85, y=1.18+row*0.72;
    sl.addShape(pres.shapes.RECTANGLE, {x, y, w:4.55, h:0.62, fill:{color:i%2===0?C.llblue:C.white}, line:{color:C.lblue,width:0.5}});
    sl.addText(k[0], {x:x+0.08, y:y+0.04, w:2.1, h:0.54, fontFace:"Courier New", fontSize:8, bold:true, color:C.blue, valign:"middle", wrap:true});
    sl.addShape(pres.shapes.LINE, {x:x+2.25, y:y+0.1, w:0, h:0.42, line:{color:C.lblue, width:0.5}});
    sl.addText(k[1], {x:x+2.35, y:y+0.04, w:2.1, h:0.54, fontFace:FONT, fontSize:9, color:C.black, valign:"middle", wrap:true});
  });
}

// Slide 7: Sheet.gs appendRow
{
  const sl = addContentSlide(pres, "3.  Sheet.gs — appendRow() 詳細");
  sl.addText("LockService によるスレッドセーフな書き込みで二重登録を防止。郵便番号・電話番号の先頭ゼロを保持。", {
    x:0.4, y:1.08, w:9.2, h:0.38, fontFace:FONT, fontSize:11, color:C.black
  });
  const steps = [
    {n:"1",label:"ロック取得",detail:"LockService.getScriptLock()\n最大30秒待機。取得失敗→例外スロー",fill:C.lorange,stroke:C.orange},
    {n:"2",label:"受付番号生成",detail:'generateReceiptNumber()\nprefix + "yyyyMMddHHmmssSSS"',fill:C.lblue,stroke:C.blue},
    {n:"3",label:"タイムゾーン変換",detail:'Utilities.formatDate(...,\n"Asia/Tokyo", ...)',fill:C.lgreen,stroke:C.green},
    {n:"4",label:"書式設定",detail:"郵便番号・電話番号列に\n\"@\" 書式を設定",fill:C.lblue,stroke:C.blue},
    {n:"5",label:"行追加",detail:"sheet.appendRow(rowData)\nシート末尾に1行追加",fill:C.lgreen,stroke:C.green},
    {n:"6",label:"ロック解放",detail:"finally ブロックで\nlock.releaseLock() 必ず実行",fill:C.lorange,stroke:C.orange},
  ];
  steps.forEach((s,i)=>{
    const y=1.6+i*0.63;
    sl.addShape(pres.shapes.RECTANGLE, {x:0.35, y, w:0.42, h:0.52, fill:{color:s.stroke}});
    sl.addText(s.n, {x:0.35, y, w:0.42, h:0.52, fontFace:FONT, fontSize:14, bold:true, color:C.white, align:"center", valign:"middle"});
    sl.addShape(pres.shapes.RECTANGLE, {x:0.85, y, w:9.15, h:0.52, fill:{color:s.fill}, line:{color:s.stroke,width:0.5}});
    sl.addText(s.label,  {x:0.95, y:y+0.04, w:2.2, h:0.44, fontFace:FONT, fontSize:11, bold:true, color:C.black, valign:"middle"});
    sl.addShape(pres.shapes.LINE, {x:3.2, y:y+0.08, w:0, h:0.36, line:{color:s.stroke,width:0.8}});
    sl.addText(s.detail, {x:3.3, y:y+0.04, w:6.6, h:0.44, fontFace:"Courier New", fontSize:9, color:C.black, valign:"middle", wrap:true});
  });
}

// Slide 8: appendToTesagyouSheet
{
  const sl = addContentSlide(pres, "3.  Sheet.gs — appendToTesagyouSheet()");
  sl.addText("受付番号をキーとして手作業シートに行を追加。B〜H列はXLOOKUP数式でメインシートからデータを自動取得。", {
    x:0.4, y:1.08, w:9.2, h:0.38, fontFace:FONT, fontSize:11, color:C.black
  });
  // Column visual
  const cols = [
    {col:"A",label:"受付番号",type:"手入力",fill:C.lorange,stroke:C.orange},
    {col:"B",label:"区分",type:"XLOOKUP",fill:C.lblue,stroke:C.blue},
    {col:"C",label:"電話番号",type:"XLOOKUP",fill:C.lblue,stroke:C.blue},
    {col:"D",label:"会社名",type:"XLOOKUP",fill:C.lblue,stroke:C.blue},
    {col:"E",label:"住所",type:"XLOOKUP",fill:C.lblue,stroke:C.blue},
    {col:"F",label:"代表者",type:"XLOOKUP",fill:C.lblue,stroke:C.blue},
    {col:"G",label:"メール",type:"XLOOKUP",fill:C.lblue,stroke:C.blue},
    {col:"H",label:"URL",type:"XLOOKUP",fill:C.lblue,stroke:C.blue},
    {col:"I",label:"受付完了",type:"checkbox",fill:C.lgreen,stroke:C.green},
    {col:"J",label:"請求書\n送信日時",type:"auto",fill:C.lyellow,stroke:C.orange},
    {col:"K",label:"入金完了",type:"checkbox",fill:C.lgreen,stroke:C.green},
    {col:"L",label:"お礼状\n送信日時",type:"auto",fill:C.lyellow,stroke:C.orange},
  ];
  const cw = 0.73, ch = 1.0, y0 = 1.6;
  cols.forEach((c,i)=>{
    const x = 0.35+i*cw;
    sl.addShape(pres.shapes.RECTANGLE, {x, y:y0, w:cw-0.04, h:0.38, fill:{color:c.stroke}});
    sl.addText(c.col, {x, y:y0, w:cw-0.04, h:0.38, fontFace:FONT, fontSize:12, bold:true, color:C.white, align:"center", valign:"middle"});
    sl.addShape(pres.shapes.RECTANGLE, {x, y:y0+0.38, w:cw-0.04, h:ch, fill:{color:c.fill}, line:{color:c.stroke,width:0.5}});
    sl.addText(c.label, {x, y:y0+0.42, w:cw-0.04, h:0.55, fontFace:FONT, fontSize:8, bold:true, color:C.black, align:"center", wrap:true});
    sl.addText(c.type, {x, y:y0+0.98, w:cw-0.04, h:0.35, fontFace:FONT, fontSize:7, color:c.stroke, align:"center"});
  });
  // XLOOKUP formula example
  sl.addText("XLOOKUP 数式例（B列）：", {x:0.35, y:2.85, w:2.5, h:0.3, fontFace:FONT, fontSize:10, bold:true, color:C.blue});
  sl.addShape(pres.shapes.RECTANGLE, {x:0.35, y:3.18, w:9.3, h:0.5, fill:{color:C.llblue}, line:{color:C.lblue,width:0.5}});
  sl.addText('=XLOOKUP(A{row}, 協賛申込み一覧!A:A, 協賛申込み一覧!M:M)', {x:0.45, y:3.22, w:9.1, h:0.42, fontFace:"Courier New", fontSize:11, color:C.blue, valign:"middle"});
  sl.addText("※ B〜E区分の場合、J列（請求書送信日時）にも申込み日時を即時セットする", {x:0.35, y:3.75, w:9.3, h:0.35, fontFace:FONT, fontSize:10, color:C.orange});
}

// Slide 9: Mail.gs
{
  const sl = addContentSlide(pres, "4.  Mail.gs — generateAndSendInvoice() フロー");
  sl.addText("請求書HTMLテンプレートからPDFを生成し、申込者にメール送信する。一時ファイルは送信後に必ず削除。", {
    x:0.4, y:1.08, w:9.2, h:0.38, fontFace:FONT, fontSize:11, color:C.black
  });
  const steps2 = [
    {n:"1",label:"テンプレート\n読込",sub:"createTemplateFromFile\n('invoice-template')",fill:C.lblue,stroke:C.blue},
    {n:"2",label:"変数\n置換",sub:"{{company}}等を\n実際の値に置換",fill:C.lgreen,stroke:C.green},
    {n:"3",label:"HTML→\nファイル変換",sub:"Drive に一時\nHTMLファイル作成",fill:C.lorange,stroke:C.orange},
    {n:"4",label:"PDF\n変換",sub:"exportLinks\n['application/pdf']",fill:C.lorange,stroke:C.orange},
    {n:"5",label:"メール\n送信",sub:"MailApp.sendEmail\nreplyTo, CC設定",fill:C.lgreen,stroke:C.green},
    {n:"6",label:"一時ファイル\n削除",sub:"DriveApp.getFile\n(id).setTrashed(true)",fill:C.lblue,stroke:C.blue},
  ];
  const bw=1.35, bh=1.55, gap=0.18, startX=0.3, y=1.6;
  steps2.forEach((s,i)=>{
    const x=startX+i*(bw+gap);
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x, y, w:bw, h:bh, fill:{color:s.fill}, line:{color:s.stroke,width:1.5}, rectRadius:0.07, shadow:makeShadow()});
    sl.addShape(pres.shapes.OVAL, {x:x+bw/2-0.22, y:y+0.06, w:0.44, h:0.44, fill:{color:s.stroke}});
    sl.addText(s.n, {x:x+bw/2-0.22, y:y+0.06, w:0.44, h:0.44, fontFace:FONT, fontSize:14, bold:true, color:C.white, align:"center", valign:"middle"});
    sl.addText(s.label, {x:x+0.06, y:y+0.58, w:bw-0.12, h:0.5, fontFace:FONT, fontSize:10, bold:true, color:C.black, align:"center", wrap:true});
    sl.addText(s.sub, {x:x+0.06, y:y+1.08, w:bw-0.12, h:0.42, fontFace:"Courier New", fontSize:8, color:s.stroke, align:"center", wrap:true});
    if(i<steps2.length-1) arrowRight(sl, x+bw+0.02, y+bh/2, gap+0.02);
  });
  sl.addText("同様のフロー: generateAndSendOreijou() — テンプレートは oreijou-template.html を使用", {
    x:0.4, y:3.3, w:9.2, h:0.35, fontFace:FONT, fontSize:10, color:C.gray
  });
}

// Slide 10: Mail template vars
{
  const sl = addContentSlide(pres, "4.  Mail.gs — メールテンプレート変数");
  const vars = [
    ["{{company}}","会社名・団体名"],
    ["{{rep}}","代表者役職・氏名"],
    ["{{kubun}}","協賛区分（S/A/B/C/D/E）"],
    ["{{amount}}","協賛金額（税込）"],
    ["{{receiptNo}}","受付番号"],
    ["{{eventName}}","イベント名（Config: EVENT_NAME）"],
    ["{{paymentDue}}","入金期限（Config: PAYMENT_DUE）"],
    ["{{orgName}}","主催団体名（Config: ORG_NAME）"],
    ["{{bankInfo}}","振込先銀行情報（複数Config値から生成）"],
    ["{{invoiceRegNo}}","インボイス登録番号（Config: INVOICE_REG_NO）"],
  ];
  const hdrRow = ["変数名","説明"].map(c=>({ text:c, options:{bold:true,fill:{color:C.navy},color:C.white,fontFace:FONT,fontSize:10}}));
  const rows = vars.map(r=>[
    { text:r[0], options:{fontFace:"Courier New",fontSize:10,color:C.blue,bold:true}},
    { text:r[1], options:{fontFace:FONT,fontSize:10,color:C.black}},
  ]);
  sl.addTable([hdrRow,...rows], { x:1.5, y:1.18, w:7, h:4.2, colW:[3.2,3.8], border:{pt:0.5,color:"bbbbbb"}, fill:{color:C.llblue} });
  sl.addText("使用例: 「{{company}} 御中」→「株式会社サンプル 御中」", {
    x:0.4, y:5.0, w:9.2, h:0.38, fontFace:FONT, fontSize:10, color:C.orange
  });
}

// Slide 11: Trigger.gs
{
  const sl = addContentSlide(pres, "5.  Trigger.gs — onEditInstallable() ロジック");
  sl.addText("手作業シートのセル編集を検知して適切なダイアログを表示するインストール型トリガー。", {
    x:0.4, y:1.08, w:9.2, h:0.35, fontFace:FONT, fontSize:11, color:C.black
  });
  // Decision diagram using shapes
  // Start
  sl.addShape(pres.shapes.OVAL, {x:4.3, y:1.55, w:1.4, h:0.42, fill:{color:C.navy}});
  sl.addText("編集検知", {x:4.3, y:1.55, w:1.4, h:0.42, fontFace:FONT, fontSize:10, bold:true, color:C.white, align:"center", valign:"middle"});
  arrowDown(sl, 5.0, 1.97, 0.38);
  // Check sheet
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x:3.5, y:2.35, w:3.0, h:0.44, fill:{color:C.lorange}, line:{color:C.orange,width:1.5}, rectRadius:0.06});
  sl.addText('シートが「手作業」か？', {x:3.5, y:2.35, w:3.0, h:0.44, fontFace:FONT, fontSize:10, bold:true, color:C.black, align:"center", valign:"middle"});
  sl.addText("NO →", {x:6.6, y:2.48, w:1.0, h:0.2, fontFace:FONT, fontSize:9, color:C.red});
  sl.addText("即 return", {x:7.6, y:2.45, w:1.5, h:0.3, fontFace:FONT, fontSize:9, color:C.red});
  arrowDown(sl, 5.0, 2.79, 0.3);
  // I or K
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x:3.5, y:3.09, w:3.0, h:0.44, fill:{color:C.lblue}, line:{color:C.blue,width:1.5}, rectRadius:0.06});
  sl.addText("I列 or K列 の変更？", {x:3.5, y:3.09, w:3.0, h:0.44, fontFace:FONT, fontSize:10, bold:true, color:C.black, align:"center", valign:"middle"});
  // I branch left
  sl.addText("I列 ←", {x:2.0, y:3.22, w:1.4, h:0.2, fontFace:FONT, fontSize:9, color:C.blue, align:"right"});
  sl.addShape(pres.shapes.RECTANGLE, {x:0.4, y:3.6, w:2.8, h:0.84, fill:{color:C.lblue}, line:{color:C.blue,width:1}});
  sl.addText("✓ ON\n→ ConfirmInvoiceDialog表示\n（J列既存→送信済み警告）", {x:0.5, y:3.64, w:2.6, h:0.76, fontFace:FONT, fontSize:8.5, color:C.black, wrap:true});
  sl.addShape(pres.shapes.RECTANGLE, {x:0.4, y:4.5, w:2.8, h:0.6, fill:{color:C.lred}, line:{color:C.red,width:1}});
  sl.addText("✓ OFF\n→ J・K・L列クリア確認", {x:0.5, y:4.52, w:2.6, h:0.56, fontFace:FONT, fontSize:8.5, color:C.black, wrap:true});
  // K branch right
  sl.addText("→ K列", {x:6.6, y:3.22, w:1.0, h:0.2, fontFace:FONT, fontSize:9, color:C.green});
  sl.addShape(pres.shapes.RECTANGLE, {x:6.8, y:3.6, w:2.8, h:0.84, fill:{color:C.lgreen}, line:{color:C.green,width:1}});
  sl.addText("✓ ON\n→ J列空→警告&リセット\n（正常）ConfirmNyukinDialog", {x:6.9, y:3.64, w:2.6, h:0.76, fontFace:FONT, fontSize:8.5, color:C.black, wrap:true});
  sl.addShape(pres.shapes.RECTANGLE, {x:6.8, y:4.5, w:2.8, h:0.6, fill:{color:C.lred}, line:{color:C.red,width:1}});
  sl.addText("✓ OFF\n→ L列クリア確認", {x:6.9, y:4.52, w:2.6, h:0.56, fontFace:FONT, fontSize:8.5, color:C.black, wrap:true});
}

// Slide 12: Trigger callbacks
{
  const sl = addContentSlide(pres, "5.  Trigger.gs — ダイアログコールバック関数");
  const cbs = [
    {fn:"sendInvoiceConfirmed(row)",caller:"ConfirmInvoiceDialog",action:"請求書PDF生成・送信 → J列に現在日時を記録",fill:C.lblue,stroke:C.blue},
    {fn:"cancelInvoiceSend(row)",caller:"ConfirmInvoiceDialog",action:"I列のチェックを FALSE に戻す",fill:C.lred,stroke:C.red},
    {fn:"sendNyukinConfirmed(row)",caller:"ConfirmNyukinDialog",action:"お礼状PDF生成・送信 → L列に現在日時を記録",fill:C.lgreen,stroke:C.green},
    {fn:"cancelNyukin(row)",caller:"ConfirmNyukinDialog",action:"K列のチェックを FALSE に戻す",fill:C.lred,stroke:C.red},
  ];
  cbs.forEach((c,i)=>{
    const y=1.25+i*1.0;
    sl.addShape(pres.shapes.RECTANGLE, {x:0.35, y, w:9.3, h:0.82, fill:{color:c.fill}, line:{color:c.stroke,width:1.5}, shadow:makeShadow()});
    sl.addShape(pres.shapes.RECTANGLE, {x:0.35, y, w:3.5, h:0.82, fill:{color:c.stroke}});
    sl.addText(c.fn, {x:0.45, y:y+0.08, w:3.3, h:0.34, fontFace:"Courier New", fontSize:9.5, bold:true, color:C.white, wrap:true});
    sl.addText("呼出元: "+c.caller, {x:0.45, y:y+0.44, w:3.3, h:0.28, fontFace:FONT, fontSize:8.5, color:"cce0f5"});
    sl.addText(c.action, {x:3.95, y:y+0.18, w:5.6, h:0.46, fontFace:FONT, fontSize:11, color:C.black, valign:"middle", wrap:true});
  });
  sl.addText("_blockSendIfLowQuota() — メール送信前ガード関数。残量が MIN_MAIL_QUOTA 未満なら例外スローし処理中断。", {
    x:0.35, y:5.25, w:9.3, h:0.3, fontFace:FONT, fontSize:9.5, color:C.orange
  });
}

// Slide 13: ProjectInit
{
  const sl = addContentSlide(pres, "6.  ProjectInit.gs — 初期化フロー");
  const steps3 = [
    {n:"1",label:"ルートフォルダ取得",detail:"ROOT_FOLDER_ID から\nDriveフォルダを取得"},
    {n:"2",label:"プロジェクト\nフォルダ作成",detail:"YYYY_EventName 形式の\nサブフォルダを作成"},
    {n:"3",label:"スプレッドシート\n作成",detail:"プロジェクトフォルダ内に\n新規スプレッドシート作成"},
    {n:"4",label:"シート構成\n設定",detail:"3シートを作成・\n列幅・書式・フィルター設定"},
    {n:"5",label:"ヘッダー\n設定",detail:"カラム名・書式・\nフィルターを設定"},
    {n:"6",label:"トリガー\n登録",detail:"onEditInstallable と\nonOpenEventSheet を登録"},
    {n:"7",label:"スプシID\n保存",detail:"作成IDを Infoシート\nDATA_SPREADSHEET_IDに書込"},
    {n:"8",label:"ログ\n記録",detail:"CreateLog シートに\n操作内容と日時を記録"},
  ];
  const bw=1.08, bh=1.35, gap=0.11, startX=0.28, y=1.55;
  const fills=[C.lblue,C.lgreen,C.lorange,C.lblue,C.lgreen,C.lorange,C.lblue,C.lgreen];
  const strokes=[C.blue,C.green,C.orange,C.blue,C.green,C.orange,C.blue,C.green];
  steps3.forEach((s,i)=>{
    const x=startX+i*(bw+gap);
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x, y, w:bw, h:bh, fill:{color:fills[i]}, line:{color:strokes[i],width:1.5}, rectRadius:0.06, shadow:makeShadow()});
    sl.addShape(pres.shapes.OVAL, {x:x+bw/2-0.2, y:y+0.05, w:0.4, h:0.4, fill:{color:strokes[i]}});
    sl.addText(s.n, {x:x+bw/2-0.2, y:y+0.05, w:0.4, h:0.4, fontFace:FONT, fontSize:12, bold:true, color:C.white, align:"center", valign:"middle"});
    sl.addText(s.label, {x:x+0.04, y:y+0.52, w:bw-0.08, h:0.42, fontFace:FONT, fontSize:9, bold:true, color:C.black, align:"center", wrap:true});
    sl.addText(s.detail,{x:x+0.04, y:y+0.94, w:bw-0.08, h:0.38, fontFace:FONT, fontSize:7.5, color:strokes[i], align:"center", wrap:true});
    if(i<steps3.length-1) arrowRight(sl, x+bw+0.01, y+bh/2, gap+0.02);
  });
  sl.addText("※ ProjectInit.gs は初回セットアップ時のみ実行。スプレッドシート作成後、GAS スクリプトに紐付けて使用する。", {
    x:0.35, y:3.1, w:9.3, h:0.38, fontFace:FONT, fontSize:10, color:C.gray
  });
}

// Slide 14: index.html
{
  const sl = addContentSlide(pres, "7.  index.html — Webフォーム詳細");
  // Screen transitions left column
  sl.addText("画面遷移", {x:0.35, y:1.1, w:4.2, h:0.35, fontFace:FONT, fontSize:13, bold:true, color:C.blue});
  const screens = [
    {name:"入力フォーム",cond:"初期表示",fill:C.lblue,stroke:C.blue},
    {name:"確認画面",cond:"バリデーション通過後",fill:C.lgreen,stroke:C.green},
    {name:"完了画面",cond:"サーバー送信成功",fill:C.lgreen,stroke:C.green},
    {name:"エラー画面",cond:"設定不備・クォータ不足",fill:C.lred,stroke:C.red},
  ];
  screens.forEach((sc,i)=>{
    const y=1.55+i*0.95;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x:0.35, y, w:4.2, h:0.78, fill:{color:sc.fill}, line:{color:sc.stroke,width:1.5}, rectRadius:0.06, shadow:makeShadow()});
    sl.addText(sc.name, {x:0.45, y:y+0.05, w:4.0, h:0.3, fontFace:FONT, fontSize:12, bold:true, color:C.black});
    sl.addText(sc.cond, {x:0.45, y:y+0.38, w:4.0, h:0.3, fontFace:FONT, fontSize:10, color:sc.stroke});
    if(i<screens.length-2) arrowDown(sl, 2.45, y+0.78, 0.17);
  });
  // Validation right column
  sl.addText("クライアントサイドバリデーション", {x:4.75, y:1.1, w:4.9, h:0.35, fontFace:FONT, fontSize:13, bold:true, color:C.blue});
  const vals = [
    ["フリガナ（4箇所）","全角カタカナのみ（/^[ァ-ヶー　]+$/）"],
    ["郵便番号","ハイフンなし7桁の数字"],
    ["電話番号","ハイフンなし10〜11桁の数字"],
    ["メールアドレス","HTML5 type=email 形式"],
    ["利用規約同意","最下部スクロール後にチェック可能"],
    ["区分選択","現在日時が受付期間内の区分のみ表示"],
  ];
  vals.forEach((v,i)=>{
    const y=1.55+i*0.65;
    sl.addShape(pres.shapes.RECTANGLE, {x:4.75, y, w:4.9, h:0.55, fill:{color:i%2===0?C.lgreen:C.white}, line:{color:"b8dcc8",width:0.5}});
    sl.addText(v[0], {x:4.85, y:y+0.04, w:1.8, h:0.47, fontFace:FONT, fontSize:9, bold:true, color:C.green, valign:"middle"});
    sl.addText(v[1], {x:6.75, y:y+0.04, w:2.8, h:0.47, fontFace:"Courier New", fontSize:8, color:C.black, valign:"middle", wrap:true});
  });
}

// Slide 15: Error handling
{
  const sl = addContentSlide(pres, "8.  エラーハンドリング方針");
  const errors = [
    ["設定値が空","doGet()","エラーHTMLを表示し、フォームは表示しない"],
    ["クォータ不足","doGet() /\nsubmitFormJson()","フォーム停止HTML返却または例外スロー"],
    ["バリデーション失敗","submitFormJson()","エラーメッセージをクライアントに返却"],
    ["ロック取得失敗","appendRow()","例外スロー → クライアントにエラー表示"],
    ["メール送信失敗","Mail.gs 各関数","try/catch で捕捉 → ログ出力 + 管理者通知"],
    ["PDF生成失敗","generateInvoicePdf()","try/catch → 一時ファイル削除してリスロー"],
    ["トリガー重複送信","onEditInstallable()","J/L列の日時チェックで二重送信を防止"],
  ];
  const hdrRow = ["エラー種別","発生箇所","対応方針"].map(c=>({ text:c, options:{bold:true,fill:{color:C.navy},color:C.white,fontFace:FONT,fontSize:10}}));
  const rows = errors.map(r=>r.map(c=>({ text:c, options:{fontFace:FONT,fontSize:9.5,color:C.black}})));
  sl.addTable([hdrRow,...rows], { x:0.35, y:1.15, w:9.3, h:4.2, colW:[2.8,1.9,4.6], border:{pt:0.5,color:"bbbbbb"}, fill:{color:C.llblue} });
}

pres.writeFile({ fileName:"詳細設計書_WebHanabi.pptx" })
  .then(()=>console.log("✓ 詳細設計書_WebHanabi.pptx generated"))
  .catch(e=>console.error(e));

"use strict";
const pptxgen = require("pptxgenjs");

const C = {
  navy:"1a3a5c", blue:"2e6da4", lblue:"dce9f5", llblue:"f0f5fb",
  green:"27ae60", lgreen:"d5f0e0", orange:"e67e22", lorange:"fde9d4",
  red:"c0392b", lred:"fce4e4", white:"FFFFFF", gray:"7f8c8d",
  lgray:"ecf0f1", black:"222222", lyellow:"fef9e7",
  purple:"6c3483", lpurple:"e8d5f5", teal:"117a65", lteal:"d1f2eb",
};
const FONT = "Meiryo";
const makeShadow = () => ({ type:"outer", blur:5, offset:2, color:"000000", opacity:0.12 });

function addContent(pres, title, accent) {
  const sl = pres.addSlide();
  sl.background = { color: C.white };
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:10,h:0.95,fill:{color:accent||C.navy}});
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0.95,w:10,h:0.06,fill:{color:C.blue}});
  if(title) sl.addText(title, {x:0.35,y:0.09,w:9.3,h:0.77,fontFace:FONT,fontSize:19,bold:true,color:C.white,valign:"middle"});
  return sl;
}
function addSection(pres, title, sub) {
  const sl = pres.addSlide();
  sl.background = { color: C.navy };
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:0.18,h:5.625,fill:{color:C.blue}});
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:5.05,w:10,h:0.575,fill:{color:C.blue},transparency:60});
  sl.addText(title, {x:0.4,y:1.7,w:9.2,h:1.2,fontFace:FONT,fontSize:36,bold:true,color:C.white,valign:"middle"});
  if(sub) sl.addText(sub, {x:0.4,y:3.1,w:9.2,h:0.6,fontFace:FONT,fontSize:15,color:"a8c4e0"});
  return sl;
}

// status badge
function statusBadge(sl, x, y, status) {
  const cfg = {
    "OK":   {fill:C.green,  text:"✓ OK"},
    "NG":   {fill:C.red,    text:"✗ NG"},
    "未":   {fill:C.gray,   text:"― 未実施"},
    "確認": {fill:C.orange, text:"△ 確認"},
  }[status] || {fill:C.gray, text:status};
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x,y,w:1.1,h:0.28,fill:{color:cfg.fill},line:{color:cfg.fill,width:0},rectRadius:0.04});
  sl.addText(cfg.text, {x,y,w:1.1,h:0.28,fontFace:FONT,fontSize:8.5,bold:true,color:C.white,align:"center",valign:"middle"});
}

// test case table
function testTable(sl, cases, y0, colW) {
  // cases: [{id, name, steps, expected, status}]
  const hdrRow = ["No.","テスト項目","操作手順","期待結果","判定"].map(c=>({
    text:c, options:{bold:true,fill:{color:C.navy},color:C.white,fontFace:FONT,fontSize:9,align:"center"}
  }));
  const rows = cases.map(c=>[
    {text:c.id,   options:{fontFace:"Courier New",fontSize:9,bold:true,color:C.blue,align:"center"}},
    {text:c.name, options:{fontFace:FONT,fontSize:9,bold:true,color:C.black}},
    {text:c.steps,options:{fontFace:FONT,fontSize:8.5,color:C.black}},
    {text:c.exp,  options:{fontFace:FONT,fontSize:8.5,color:C.black}},
    {text:c.status,options:{fontFace:FONT,fontSize:9,bold:true,color:
      c.status==="OK"?C.green:c.status==="NG"?C.red:C.gray,align:"center"}},
  ]);
  sl.addTable([hdrRow,...rows], {
    x:0.35, y:y0, w:9.3,
    colW: colW||[0.55,2.0,3.0,3.0,0.75],
    border:{pt:0.5,color:"bbbbbb"}, fill:{color:C.llblue}
  });
}

// ────────────────────────────────────────────────────────────────────────────
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "テスト仕様書 — WebHanabi";

// ══ Slide 1: Cover ══════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: C.navy };
  // accent shapes
  for(const [x,y,r,t] of [[8.6,0.8,2.0,18],[1.2,5.0,1.3,12],[9.1,4.7,1.0,22]]) {
    sl.addShape(pres.shapes.OVAL,{x:x-r/2,y:y-r/2,w:r,h:r,fill:{color:C.blue,transparency:t}});
  }
  sl.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.22,h:5.625,fill:{color:C.blue}});
  sl.addShape(pres.shapes.RECTANGLE,{x:0,y:4.8,w:10,h:0.825,fill:{color:C.blue}});
  sl.addText("テスト仕様書",{x:0.5,y:1.0,w:9,h:1.7,fontFace:FONT,fontSize:50,bold:true,color:C.white,valign:"middle"});
  sl.addText("WebHanabi — 川口花火大会 協賛管理システム",{x:0.5,y:2.85,w:9,h:0.6,fontFace:FONT,fontSize:17,color:"a8c4e0"});
  sl.addShape(pres.shapes.LINE,{x:0.5,y:3.55,w:9,h:0,line:{color:C.blue,width:1}});
  sl.addText("発行日：2026年06月18日　　バージョン：1.0　　対象：QA担当者・開発者",{x:0.5,y:3.65,w:9,h:0.38,fontFace:FONT,fontSize:11,color:"8ab4cc"});
  // test type chips
  for(const [i,label,fill] of [
    [0,"機能テスト",C.blue],[1,"バリデーションテスト",C.green],
    [2,"メール送信テスト",C.orange],[3,"ワークフローテスト",C.purple]
  ]) {
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.5+i*2.35,y:4.88,w:2.15,h:0.34,fill:{color:fill,transparency:40},line:{color:fill,width:0},rectRadius:0.05});
    sl.addText(label,{x:0.5+i*2.35,y:4.88,w:2.15,h:0.34,fontFace:FONT,fontSize:9.5,bold:true,color:C.white,align:"center",valign:"middle"});
  }
}

// ══ Slide 2: テスト概要 ══════════════════════════════════════════════════════
{
  const sl = addContent(pres, "テスト概要");
  // purpose
  sl.addShape(pres.shapes.RECTANGLE,{x:0.35,y:1.1,w:9.3,h:0.35,fill:{color:C.blue}});
  sl.addText("目的",{x:0.45,y:1.12,w:9.1,h:0.31,fontFace:FONT,fontSize:12,bold:true,color:C.white});
  sl.addText("本書は WebHanabi システムの各機能が仕様通りに動作することを確認するためのテスト仕様を定義する。申込みフォームの入力受付から、メール送信・スプレッドシート管理・ワークフロー制御までの全機能をカバーする。",{
    x:0.45,y:1.5,w:9.1,h:0.55,fontFace:FONT,fontSize:10.5,color:C.black,wrap:true});

  // scope cards
  const scopes = [
    {label:"対象",items:["協賛申込みWebフォーム","スプレッドシートへのデータ登録","メール自動送信（請求書・お礼状）","onEditトリガーとダイアログ動作","Infoシート設定値の反映"],fill:C.lblue,stroke:C.blue},
    {label:"対象外",items:["Google インフラの可用性テスト","メール到達性（Gmail SLA）","スプレッドシートの行上限テスト","LibreOffice / 他ブラウザ互換性"],fill:C.lred,stroke:C.red},
    {label:"テスト種別",items:["機能テスト（黒箱）","バリデーションテスト","異常系テスト","回帰テスト（設定変更後）"],fill:C.lgreen,stroke:C.green},
  ];
  scopes.forEach((s,i)=>{
    const x=0.35+i*3.15, y=2.2;
    sl.addShape(pres.shapes.RECTANGLE,{x,y,w:3.0,h:3.1,fill:{color:s.fill},line:{color:s.stroke,width:1.5},shadow:makeShadow()});
    sl.addShape(pres.shapes.RECTANGLE,{x,y,w:3.0,h:0.42,fill:{color:s.stroke}});
    sl.addText(s.label,{x:x+0.1,y:y+0.06,w:2.8,h:0.3,fontFace:FONT,fontSize:12,bold:true,color:C.white});
    s.items.forEach((item,ii)=>{
      const iy=y+0.55+ii*0.5;
      sl.addShape(pres.shapes.OVAL,{x:x+0.1,y:iy+0.06,w:0.22,h:0.22,fill:{color:s.stroke}});
      sl.addText(item,{x:x+0.4,y:iy,w:2.5,h:0.38,fontFace:FONT,fontSize:9.5,color:C.black,valign:"middle",wrap:true});
    });
  });
}

// ══ Slide 3: テスト環境 ═══════════════════════════════════════════════════════
{
  const sl = addContent(pres, "テスト環境");
  const envs = [
    {cat:"実行環境",items:[["GAS実行環境","Google Apps Script（V8ランタイム）"],["スプレッドシート","Google Sheets（テスト専用ファイル）"],["メール","Gmail（テスト用アカウント）"],["Drive","Google Drive（テスト用フォルダ）"]]},
    {cat:"テストツール",items:[["ブラウザ","Google Chrome（最新版）"],["GASデバッガ","Apps Script エディタ（実行ログ）"],["スプレッドシート確認","目視検証"],["メール確認","Gmailで受信確認"]]},
    {cat:"前提条件",items:[["Infoシート","テスト用イベント情報を設定済み"],["受付フォーム","Webアプリとしてデプロイ済み"],["テストデータ","申込みデータを事前準備"],["クォータ","MIN_MAIL_QUOTAを十分に確保"]]},
  ];
  envs.forEach((e,i)=>{
    const x=0.35+i*3.15, y=1.18;
    sl.addShape(pres.shapes.RECTANGLE,{x,y,w:3.0,h:4.18,fill:{color:C.llblue},line:{color:C.lblue,width:1},shadow:makeShadow()});
    sl.addShape(pres.shapes.RECTANGLE,{x,y,w:3.0,h:0.42,fill:{color:C.navy}});
    sl.addText(e.cat,{x:x+0.1,y:y+0.06,w:2.8,h:0.3,fontFace:FONT,fontSize:12,bold:true,color:C.white});
    e.items.forEach(([k,v],ii)=>{
      const ry=y+0.55+ii*0.87;
      sl.addShape(pres.shapes.RECTANGLE,{x:x+0.1,y:ry,w:2.8,h:0.75,fill:{color:C.white},line:{color:C.lblue,width:0.5}});
      sl.addText(k,{x:x+0.18,y:ry+0.04,w:2.65,h:0.28,fontFace:FONT,fontSize:9,bold:true,color:C.blue});
      sl.addText(v,{x:x+0.18,y:ry+0.32,w:2.65,h:0.38,fontFace:FONT,fontSize:9,color:C.black,wrap:true});
    });
  });
}

// ══ Slide 4: テスト項目一覧 ══════════════════════════════════════════════════
{
  const sl = addContent(pres, "テスト項目一覧（全体マップ）");
  const groups = [
    {id:"T-01〜07",label:"フォーム表示・\n入力バリデーション",fill:C.lblue,stroke:C.blue,count:7},
    {id:"T-08〜12",label:"申込みデータ\n登録・受付番号",fill:C.lgreen,stroke:C.green,count:5},
    {id:"T-13〜18",label:"メール自動送信\n（B〜E区分）",fill:C.lorange,stroke:C.orange,count:6},
    {id:"T-19〜23",label:"受付完了チェック\n（S・A区分）",fill:C.lblue,stroke:C.blue,count:5},
    {id:"T-24〜28",label:"入金確認・\nお礼状送信",fill:C.lgreen,stroke:C.green,count:5},
    {id:"T-29〜33",label:"メールクォータ\n管理",fill:C.lorange,stroke:C.orange,count:5},
    {id:"T-34〜38",label:"Infoシート\n設定反映",fill:C.lpurple,stroke:C.purple,count:5},
    {id:"T-39〜43",label:"onEditトリガー\n安全確認",fill:C.lteal,stroke:C.teal,count:5},
  ];
  // Total
  sl.addShape(pres.shapes.RECTANGLE,{x:0.35,y:1.12,w:9.3,h:0.38,fill:{color:C.navy}});
  sl.addText("合計テストケース数：43件　　優先度：高 23件 / 中 14件 / 低 6件",{
    x:0.45,y:1.14,w:9.1,h:0.34,fontFace:FONT,fontSize:11,color:C.white,valign:"middle"});

  groups.forEach((g,i)=>{
    const col=i%4, row=Math.floor(i/4);
    const x=0.35+col*2.35, y=1.6+row*1.85;
    sl.addShape(pres.shapes.RECTANGLE,{x,y,w:2.2,h:1.65,fill:{color:g.fill},line:{color:g.stroke,width:1.5},shadow:makeShadow()});
    sl.addShape(pres.shapes.RECTANGLE,{x,y,w:2.2,h:0.42,fill:{color:g.stroke}});
    sl.addText(g.id,{x:x+0.08,y:y+0.06,w:2.04,h:0.3,fontFace:"Courier New",fontSize:10,bold:true,color:C.white});
    sl.addText(g.label,{x:x+0.08,y:y+0.52,w:2.04,h:0.65,fontFace:FONT,fontSize:10.5,bold:true,color:C.black,wrap:true});
    sl.addShape(pres.shapes.OVAL,{x:x+1.62,y:y+1.24,w:0.46,h:0.3,fill:{color:g.stroke}});
    sl.addText(String(g.count)+"件",{x:x+1.62,y:y+1.24,w:0.46,h:0.3,fontFace:FONT,fontSize:9,bold:true,color:C.white,align:"center",valign:"middle"});
  });
}

// ══ Slide 5: T-01〜07 フォームバリデーション ══════════════════════════════════
addSection(pres, "T-01〜07  フォーム表示・\n入力バリデーション", "Webフォームの表示条件と入力値の検証テスト");

{
  const sl = addContent(pres, "T-01〜04  フォーム表示テスト");
  testTable(sl, [
    {id:"T-01",name:"フォーム\n正常表示",steps:"①デプロイURLにブラウザでアクセスする\n②ページが読み込まれることを確認",exp:"協賛申込みフォームが表示される\nイベント名・区分・価格が正しく表示される",status:"未"},
    {id:"T-02",name:"INFO未設定\n時のエラー",steps:"①InfoシートのEVENT_NAMEを空にする\n②フォームURLにアクセスする",exp:"「設定が不足しています」等のエラー画面が表示され、申込みフォームは表示されない",status:"未"},
    {id:"T-03",name:"申込み期間外\n区分の非表示",steps:"①KUBUN_SA_END を過去日付に設定\n②フォームにアクセスしてS・A区分の表示を確認",exp:"S・A区分のラジオボタンが表示されない\nB〜E区分のみ表示される",status:"未"},
    {id:"T-04",name:"背景画像\n表示確認",steps:"①BG_IMAGE_IDに有効なDriveファイルIDを設定\n②フォームにアクセスして背景を確認",exp:"フォームの背景に指定した画像が表示される",status:"未"},
  ], 1.18);
}

{
  const sl = addContent(pres, "T-05〜07  バリデーションテスト（フォーム入力）");
  testTable(sl, [
    {id:"T-05",name:"フリガナ\n全角カタカナ",steps:"①会社名フリガナ欄にひらがな「かわぐち」を入力\n②「確認」ボタンを押す",exp:"バリデーションエラーが表示される\n「全角カタカナで入力してください」等のメッセージ",status:"未"},
    {id:"T-06",name:"郵便番号\n桁数チェック",steps:"①郵便番号欄に6桁「123456」を入力\n②「確認」ボタンを押す",exp:"バリデーションエラーが表示される\n7桁の数字を要求するメッセージ",status:"未"},
    {id:"T-07",name:"利用規約\nスクロール制御",steps:"①利用規約エリアをスクロールせずに同意チェックボックスを押す\n②スクロール後にチェックを試みる",exp:"スクロール前：チェックボックスが無効（クリック不可）\nスクロール後：チェックボックスが有効になる",status:"未"},
  ], 1.18, [0.55,2.0,3.3,2.7,0.75]);
}

// ══ Slide 8: T-08〜12 データ登録 ══════════════════════════════════════════════
addSection(pres, "T-08〜12  申込みデータ登録", "スプレッドシートへのデータ保存と受付番号生成のテスト");

{
  const sl = addContent(pres, "T-08〜12  申込みデータ登録テスト");
  testTable(sl, [
    {id:"T-08",name:"正常登録",steps:"①全必須項目を正しく入力してフォーム送信\n②スプレッドシートの「協賛申込み一覧」を確認",exp:"送信したデータが最終行に追記される\n受付日時・区分・連絡先情報がすべて正しく登録される",status:"未"},
    {id:"T-09",name:"受付番号\n形式確認",steps:"①フォームを正常送信する\n②A列の受付番号を確認する",exp:"受付番号が「KWGC＋17桁数字」形式で生成される\nフォーム完了画面にも同じ番号が表示される",status:"未"},
    {id:"T-10",name:"郵便番号\n先頭ゼロ保持",steps:"①郵便番号「0123456」を入力して送信\n②スプレッドシートI列の値を確認",exp:"「0123456」として保存される（先頭ゼロが消えない）",status:"未"},
    {id:"T-11",name:"手作業シート\nXLOOKUP自動入力",steps:"①フォーム送信後、手作業シートのA列に受付番号を入力\n②B〜H列のデータを確認",exp:"B〜H列に会社名・住所等がXLOOKUPで自動入力される",status:"未"},
    {id:"T-12",name:"同時送信\n（ロック確認）",steps:"①2つのブラウザタブで同時にフォームを送信する\n②スプレッドシートの行を確認",exp:"両方の申込みが重複なく別行に登録される\nデータが上書きされたり欠損したりしない",status:"未"},
  ], 1.18);
}

// ══ Slide 10: T-13〜18 メール送信 ══════════════════════════════════════════════
addSection(pres, "T-13〜18  メール自動送信テスト", "B〜E区分の申込み直後メール送信と請求書PDFのテスト");

{
  const sl = addContent(pres, "T-13〜16  B〜E区分 申込み確認メールテスト");
  testTable(sl, [
    {id:"T-13",name:"B〜E区分\nメール自動送信",steps:"①B区分でフォームを正常送信する\n②申込者メールアドレスの受信ボックスを確認",exp:"申込み直後にメールが届く\n件名に「受付確認」または「請求書」が含まれる",status:"未"},
    {id:"T-14",name:"請求書PDF\n添付確認",steps:"①T-13のメールを開いてPDFを確認する",exp:"PDFが添付されている\nPDFに会社名・区分・金額（税込）・インボイス番号・振込先が含まれる",status:"未"},
    {id:"T-15",name:"CC送信\n確認",steps:"①T-13のメール送信後、事務局メールを確認\n②CCで届いているか確認する",exp:"OFFICE_EMAILにCCでメールが届いている",status:"未"},
    {id:"T-16",name:"S・A区分\nメール不送信",steps:"①S区分でフォームを正常送信する\n②申込者メールアドレスを確認する",exp:"フォーム送信直後にメールが届かない\n（S・A区分は手動送信のため）",status:"未"},
  ], 1.18);
}

{
  const sl = addContent(pres, "T-17〜18  PDFコンテンツ・メールアドレステスト");
  testTable(sl, [
    {id:"T-17",name:"PDF記載内容\n全項目確認",steps:"①B区分で申込み送信してPDFを取得する\n②PDF内の各項目を確認する",exp:[
      "・会社名・担当者名が正しい",
      "・協賛区分と金額（税抜・税込）が正しい",
      "・消費税10%計算が正しい",
      "・インボイス登録番号（T+13桁）が表示されている",
      "・振込先銀行情報が正しい",
      "・入金期限が正しい",
    ].join("\n"),status:"未"},
    {id:"T-18",name:"不正メール\nアドレス",steps:"①メールアドレス欄に「test@」（不完全）を入力\n②送信ボタンを押す",exp:"フォームのバリデーションエラーが表示される\nサーバーへの送信が行われない",status:"未"},
  ], 1.18, [0.55,1.8,2.8,3.4,0.75]);
}

// ══ Slide 13: T-19〜23 受付完了チェック ══════════════════════════════════════
addSection(pres, "T-19〜23  受付完了チェックテスト", "S・A区分の I列チェックボックスとダイアログ動作のテスト");

{
  const sl = addContent(pres, "T-19〜23  I列（受付完了）チェックボックステスト");
  testTable(sl, [
    {id:"T-19",name:"受付完了\nダイアログ表示",steps:"①手作業シートでS区分の行のI列チェックボックスをONにする",exp:"確認ダイアログが表示される\nダイアログに会社名・区分・金額が表示されている",status:"未"},
    {id:"T-20",name:"送信確認後\nメール送信",steps:"①T-19のダイアログで「送信する」をクリックする\n②申込者メールを確認する",exp:"申込者に請求書PDF添付メールが届く\nJ列（請求書送信日時）に現在日時が自動記録される",status:"未"},
    {id:"T-21",name:"キャンセル時\nI列リセット",steps:"①I列をONにしてダイアログを表示する\n②「キャンセル」をクリックする",exp:"I列のチェックが FALSE に戻る\nメールは送信されない",status:"未"},
    {id:"T-22",name:"二重送信\n防止確認",steps:"①J列に日時が入力済みの行のI列を再度ONにする",exp:"「送信済み」の警告アラートが表示される\nダイアログは表示されず、二重送信されない",status:"未"},
    {id:"T-23",name:"I列OFF時\n連鎖クリア確認",steps:"①J列・K列・L列にデータがある行のI列をOFFにする",exp:"「J列以降をクリアしますか？」の確認が表示される\nOKを選択するとJ・K・L列がクリアされる",status:"未"},
  ], 1.18);
}

// ══ Slide 15: T-24〜28 入金確認 ═══════════════════════════════════════════════
addSection(pres, "T-24〜28  入金確認・お礼状送信テスト", "K列チェックボックスとお礼状PDF送信のテスト");

{
  const sl = addContent(pres, "T-24〜28  K列（入金完了）チェックボックステスト");
  testTable(sl, [
    {id:"T-24",name:"J列空のまま\nK列チェック",steps:"①J列（請求書送信日時）が空の行のK列をONにする",exp:"「請求書を先に送信してください」等の警告アラートが表示される\nK列のチェックが FALSE にリセットされる",status:"未"},
    {id:"T-25",name:"入金完了\nダイアログ表示",steps:"①J列に日時がある行のK列チェックボックスをONにする",exp:"お礼状送信確認ダイアログが表示される\n会社名・区分が正しく表示されている",status:"未"},
    {id:"T-26",name:"お礼状メール\n送信確認",steps:"①T-25のダイアログで「送信する」をクリックする\n②申込者メールを確認する",exp:"お礼状PDF添付メールが届く\nL列（お礼状送信日時）に現在日時が自動記録される",status:"未"},
    {id:"T-27",name:"お礼状PDF\n内容確認",steps:"①T-26で受信したお礼状PDFを開く",exp:"会社名・代表者名・区分・イベント名が正しい\n主催団体名・代表者名・日付が正しい\n組織印影が表示されている",status:"未"},
    {id:"T-28",name:"K列OFF時\nL列クリア確認",steps:"①L列にデータがある行のK列をOFFにする",exp:"「L列をクリアしますか？」の確認が表示される\nOKを選択するとL列がクリアされる",status:"未"},
  ], 1.18);
}

// ══ Slide 17: T-29〜33 クォータ ══════════════════════════════════════════════
addSection(pres, "T-29〜33  メールクォータ管理テスト", "Gmail送信上限の監視・通知・停止機能のテスト");

{
  const sl = addContent(pres, "T-29〜33  メールクォータ管理テスト");
  testTable(sl, [
    {id:"T-29",name:"通常クォータ\n時の動作確認",steps:"①MIN_MAIL_QUOTAを5に設定（デフォルト）\n②残クォータが十分ある状態でフォームにアクセス",exp:"フォームが正常に表示される\nクォータ警告は表示されない",status:"未"},
    {id:"T-30",name:"低クォータ\n警告メール",steps:"①MIN_MAIL_QUOTAを実際の残クォータより大きく設定する\n②フォームを送信する",exp:"OFFICE_EMAILに「クォータ残量警告」メールが届く\nフォームは引き続き受付を行う",status:"未"},
    {id:"T-31",name:"クォータ0時\nフォーム停止",steps:"①MailApp.getRemainingDailyQuotaが0を返す状態を模擬する\n※Script Propertiesで制御するか翌日テスト",exp:"フォームURLにアクセスすると「本日の受付を停止しました」等のメッセージが表示される\n申込みボタンが非活性または画面が出ない",status:"未"},
    {id:"T-32",name:"MIN_MAIL_QUOTA\n変更反映",steps:"①InfoシートのMIN_MAIL_QUOTAを10に変更する\n②「設定を同期」を実行する\n③動作を確認する",exp:"新しい閾値（10）が反映される\n残クォータが10以下になると警告メールが送信される",status:"未"},
    {id:"T-33",name:"クォータ翌日\nリセット確認",steps:"①前日クォータ消費後、翌日にフォームにアクセスする",exp:"フォームが正常に表示される\nクォータが自動リセットされている",status:"未"},
  ], 1.18);
}

// ══ Slide 19: T-34〜38 Infoシート設定 ════════════════════════════════════════
addSection(pres, "T-34〜38  Infoシート設定反映テスト", "設定値の変更がフォーム・PDF・メールに正しく反映されるか確認");

{
  const sl = addContent(pres, "T-34〜38  Infoシート設定反映テスト");
  testTable(sl, [
    {id:"T-34",name:"価格変更\n反映確認",steps:"①InfoシートのPRICE_Bを変更して「設定を同期」を実行\n②フォームにアクセスしてB区分の価格を確認\n③フォーム送信して請求書PDFの金額を確認",exp:"フォーム上の表示価格が変更後の値になっている\n請求書PDFの金額も変更後の値になっている",status:"未"},
    {id:"T-35",name:"銀行情報変更\n反映確認",steps:"①BANK_NAMEを変更して「設定を同期」を実行\n②フォームを送信して請求書PDFを確認する",exp:"請求書PDFの振込先に変更後の銀行情報が表示される",status:"未"},
    {id:"T-36",name:"イベント名変更\n反映確認",steps:"①EVENT_NAMEを変更して「設定を同期」を実行\n②フォームにアクセスしてタイトルを確認",exp:"フォームのページタイトルとヘッダーに変更後のイベント名が表示される",status:"未"},
    {id:"T-37",name:"INVOICE_REG_NO\n反映確認",steps:"①INVOICE_REG_NOに「T1234567890123」を設定\n②フォームを送信して請求書PDFを確認",exp:"請求書PDFにインボイス登録番号「T1234567890123」が表示される",status:"未"},
    {id:"T-38",name:"申込み期間\n設定反映",steps:"①KUBUN_SA_STARTを明日の日付に設定する\n②今日フォームにアクセスしてS区分の表示を確認",exp:"S・A区分が表示されない（申込み期間外のため）\nB〜E区分のみ表示される",status:"未"},
  ], 1.18);
}

// ══ Slide 21: T-39〜43 onEditトリガー安全確認 ════════════════════════════════
addSection(pres, "T-39〜43  onEditトリガー安全確認テスト", "誤操作防止・異常系の動作確認テスト");

{
  const sl = addContent(pres, "T-39〜43  onEditトリガー異常系・安全確認テスト");
  testTable(sl, [
    {id:"T-39",name:"手作業以外\nシートの編集",steps:"①「協賛申込み一覧」シートのセルを編集する\n②ダイアログが表示されないことを確認",exp:"ダイアログが表示されない\nスクリプトが静かにreturnする",status:"未"},
    {id:"T-40",name:"存在しない\n受付番号入力",steps:"①手作業シートのA列に存在しない受付番号を入力する\n②B〜H列の値を確認する",exp:"B〜H列に「#N/A」エラーが表示される\nクラッシュは発生しない",status:"未"},
    {id:"T-41",name:"I列ON→即OFF\n（キャンセル後の状態）",steps:"①I列をONにしてダイアログを表示\n②「キャンセル」をクリックしてI列の状態を確認",exp:"I列が FALSE（未チェック）に戻っている\nJ列は変化していない",status:"未"},
    {id:"T-42",name:"トリガー未登録\n時の動作",steps:"①Apps Scriptのトリガー一覧からonEditInstallableを削除する\n②I列チェックボックスをONにする",exp:"ダイアログが表示されない\n（ProjectInitializeの再実行でトリガーが再登録される）",status:"未"},
    {id:"T-43",name:"複数行同時\nチェック確認",steps:"①手作業シートで複数行のI列を素早く連続してONにする\n②各行のダイアログと送信状況を確認",exp:"各行ごとにダイアログが表示・処理される\n処理が競合してクラッシュしない",status:"未"},
  ], 1.18);
}

// ══ Slide 23: テスト結果サマリー（記録用）══════════════════════════════════════
{
  const sl = addContent(pres, "テスト結果サマリー（記録用テンプレート）");
  // Summary counters
  const counters = [
    {label:"テスト総件数",val:"43",fill:C.navy,text:"件"},
    {label:"合格（OK）",val:"__",fill:C.green,text:"件"},
    {label:"不合格（NG）",val:"__",fill:C.red,text:"件"},
    {label:"未実施",val:"43",fill:C.gray,text:"件"},
    {label:"合格率",val:"__%",fill:C.blue,text:""},
  ];
  counters.forEach((c,i)=>{
    const x=0.35+i*1.85, y=1.12;
    sl.addShape(pres.shapes.RECTANGLE,{x,y,w:1.7,h:1.45,fill:{color:c.fill},line:{color:c.fill,width:0},shadow:makeShadow()});
    sl.addText(c.label,{x:x+0.05,y:y+0.08,w:1.6,h:0.35,fontFace:FONT,fontSize:9,bold:true,color:C.white,align:"center",wrap:true});
    sl.addText(c.val,{x:x+0.05,y:y+0.5,w:1.6,h:0.7,fontFace:FONT,fontSize:28,bold:true,color:C.white,align:"center",valign:"middle"});
    if(c.text) sl.addText(c.text,{x:x+0.05,y:y+1.18,w:1.6,h:0.22,fontFace:FONT,fontSize:10,color:C.white,align:"center"});
  });

  // Result table header
  sl.addText("テストグループ別 結果記録", {x:0.35,y:2.72,w:5,h:0.35,fontFace:FONT,fontSize:12,bold:true,color:C.navy});
  sl.addText("テスト実施者：___________________　　実施日：___________________", {x:5.0,y:2.72,w:4.65,h:0.35,fontFace:FONT,fontSize:10,color:C.gray});
  const groups2 = [
    ["T-01〜07","フォーム表示・入力バリデーション","7","___","___","___"],
    ["T-08〜12","申込みデータ登録・受付番号","5","___","___","___"],
    ["T-13〜18","メール自動送信（B〜E区分）","6","___","___","___"],
    ["T-19〜23","受付完了チェック（S・A区分）","5","___","___","___"],
    ["T-24〜28","入金確認・お礼状送信","5","___","___","___"],
    ["T-29〜33","メールクォータ管理","5","___","___","___"],
    ["T-34〜38","Infoシート設定反映","5","___","___","___"],
    ["T-39〜43","onEditトリガー安全確認","5","___","___","___"],
  ];
  const hdr2 = ["グループ","テスト内容","総数","OK","NG","備考"].map(c=>({
    text:c,options:{bold:true,fill:{color:C.navy},color:C.white,fontFace:FONT,fontSize:9,align:"center"}
  }));
  const rows2 = groups2.map(r=>r.map(c=>({text:c,options:{fontFace:FONT,fontSize:9,color:C.black}})));
  sl.addTable([hdr2,...rows2],{x:0.35,y:3.12,w:9.3,h:2.35,colW:[1.5,4.0,0.7,0.7,0.7,1.7],border:{pt:0.5,color:"bbbbbb"},fill:{color:C.llblue}});
}

// ══ Slide 24: まとめ・注意事項 ════════════════════════════════════════════════
{
  const sl = addSection(pres, "テスト完了基準・注意事項", "WebHanabi テスト仕様書");
  const criteria = [
    {icon:"✦",title:"完了基準",body:"全43テストケースが「OK」判定であること。NGが0件であること。",fill:C.blue},
    {icon:"⚠",title:"メール送信テスト時の注意",body:"実際のメールが送信されます。テスト用Gmailアカウントを使用し、申込者メールアドレスにはテスト用アドレスを使用してください。",fill:"7d3c00"},
    {icon:"✦",title:"クォータ消費に注意",body:"T-13〜28のメール送信テストはGmailクォータを消費します。1日の上限（50通）に注意して計画的に実施してください。",fill:"1a6e40"},
    {icon:"✦",title:"テストデータのクリーンアップ",body:"テスト完了後、テスト用スプレッドシートのデータを削除し、本番環境に影響がないことを確認してください。",fill:"6a1520"},
  ];
  criteria.forEach((c,i)=>{
    const y=2.05+i*0.8;
    sl.addShape(pres.shapes.RECTANGLE,{x:0.4,y,w:9.2,h:0.68,fill:{color:c.fill},line:{color:c.fill,width:0}});
    sl.addText(c.icon,{x:0.5,y:y+0.06,w:0.38,h:0.56,fontFace:FONT,fontSize:16,color:"FFD700",valign:"middle",align:"center"});
    sl.addText(c.title,{x:0.98,y:y+0.04,w:2.2,h:0.28,fontFace:FONT,fontSize:12,bold:true,color:C.white});
    sl.addText(c.body,{x:0.98,y:y+0.32,w:8.55,h:0.3,fontFace:FONT,fontSize:9.5,color:"d4e8f5",wrap:true});
  });
}

pres.writeFile({fileName:"テスト仕様書_WebHanabi.pptx"})
  .then(()=>console.log("✓ テスト仕様書_WebHanabi.pptx generated"))
  .catch(e=>console.error(e));

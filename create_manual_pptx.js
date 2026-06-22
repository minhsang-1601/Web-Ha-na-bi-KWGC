"use strict";
const pptxgen = require("pptxgenjs");

const C = {
  navy:"1a3a5c", blue:"2e6da4", lblue:"dce9f5", llblue:"f0f5fb",
  green:"27ae60", lgreen:"d5f0e0", orange:"e67e22", lorange:"fde9d4",
  red:"c0392b", lred:"fce4e4", white:"FFFFFF", gray:"7f8c8d",
  lgray:"ecf0f1", black:"222222", lyellow:"fef9e7", yellow:"f1c40f"
};
const FONT = "Meiryo";
const makeShadow = () => ({ type:"outer", blur:5, offset:2, color:"000000", opacity:0.12 });

function addNavy(pres, title, subtitle) {
  const sl = pres.addSlide();
  sl.background = { color: C.navy };
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:0.18,h:5.625,fill:{color:C.blue}});
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:5.05,w:10,h:0.575,fill:{color:C.blue},transparency:60});
  if(title) sl.addText(title, {x:0.4,y:1.7,w:9.2,h:1.3,fontFace:FONT,fontSize:36,bold:true,color:C.white,valign:"middle"});
  if(subtitle) sl.addText(subtitle, {x:0.4,y:3.1,w:9.2,h:0.65,fontFace:FONT,fontSize:16,color:"a8c4e0"});
  return sl;
}
function addContent(pres, title) {
  const sl = pres.addSlide();
  sl.background = { color: C.white };
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:10,h:0.95,fill:{color:C.navy}});
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0.95,w:10,h:0.06,fill:{color:C.blue}});
  if(title) sl.addText(title, {x:0.35,y:0.09,w:9.3,h:0.77,fontFace:FONT,fontSize:19,bold:true,color:C.white,valign:"middle"});
  return sl;
}
function arrowDown(sl, x, y, h) {
  sl.addShape(pres.shapes.LINE, {x,y,w:0,h:h*0.75,line:{color:C.gray,width:1.5}});
  sl.addText("▼", {x:x-0.1,y:y+h*0.72,w:0.2,h:0.18,fontFace:FONT,fontSize:9,color:C.gray,align:"center"});
}
function arrowRight(sl, x, y, len) {
  sl.addShape(pres.shapes.LINE, {x,y,w:len*0.78,h:0,line:{color:C.gray,width:1.5}});
  sl.addText("▶", {x:x+len*0.76,y:y-0.1,w:0.18,h:0.18,fontFace:FONT,fontSize:9,color:C.gray,align:"center"});
}
function stepBadge(sl, pres, n, x, y, color) {
  sl.addShape(pres.shapes.OVAL, {x,y,w:0.38,h:0.38,fill:{color:color||C.blue}});
  sl.addText(String(n), {x,y,w:0.38,h:0.38,fontFace:FONT,fontSize:11,bold:true,color:C.white,align:"center",valign:"middle"});
}

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "操作マニュアル — WebHanabi";

// ── Slide 1: Cover ────────────────────────────────────────────────────────
{
  const sl = pres.addSlide();
  sl.background = { color: C.navy };
  // Firework decorations
  const fwSpots = [{cx:8.6,cy:0.9,r:55,num:12,step:30,color:C.blue},{cx:1.4,cy:4.9,r:38,num:8,step:45,color:"1a6e40"},{cx:9.0,cy:4.5,r:46,num:10,step:36,color:C.orange}];
  for(const fw of fwSpots) {
    for(let i=0;i<fw.num;i++) {
      const angle = (i*fw.step)*Math.PI/180;
      const x1=fw.cx+0.1*Math.cos(angle), y1=fw.cy+0.1*Math.sin(angle);
      const x2=fw.cx+(fw.r/72)*Math.cos(angle), y2=fw.cy+(fw.r/72)*Math.sin(angle);
      sl.addShape(pres.shapes.LINE, {x:x1,y:y1,w:x2-x1,h:y2-y1,line:{color:fw.color,width:1,transparency:55}});
    }
  }
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:0.22,h:5.625,fill:{color:C.blue}});
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:4.8,w:10,h:0.825,fill:{color:C.blue}});
  sl.addText("操 作 マ ニ ュ ア ル", {x:0.5,y:1.0,w:9,h:1.7,fontFace:FONT,fontSize:44,bold:true,color:C.white,valign:"middle"});
  sl.addText("WebHanabi — 川口花火大会 協賛管理システム", {x:0.5,y:2.8,w:9,h:0.65,fontFace:FONT,fontSize:18,color:"a8c4e0"});
  sl.addShape(pres.shapes.LINE, {x:0.5,y:3.55,w:9,h:0,line:{color:C.blue,width:1}});
  sl.addText("2026年06月18日発行　　　対象：申込者・事務局担当者", {x:0.5,y:3.65,w:9,h:0.4,fontFace:FONT,fontSize:12,color:"8ab4cc"});
  sl.addText("申込みから感謝状まで — ワンストップ協賛管理", {x:0.5,y:4.88,w:9,h:0.42,fontFace:FONT,fontSize:13,color:C.white,align:"center"});
}

// ── Slide 2: 対象読者 ─────────────────────────────────────────────────────
{
  const sl = addContent(pres, "本書の対象読者");
  const personas = [
    {role:"申込者（一般）",icon:"👤",desc:"協賛申込みフォームを使って申込む方",tasks:["フォームにアクセスして情報を入力","確認画面で内容を確認して送信","受付番号と確認メールを受け取る"],fill:C.lblue,stroke:C.blue},
    {role:"事務局担当者",icon:"📋",desc:"申込み受付・請求・入金確認を行う方",tasks:["手作業シートで受付番号を管理","チェックボックスで請求書を送信","入金確認後にお礼状を送信する"],fill:C.lgreen,stroke:C.green},
    {role:"システム管理者",icon:"⚙️",desc:"システムの設定・初期化を行う方",tasks:["Infoシートに設定値を入力する","Project Initializeを実行する","Webアプリとしてデプロイする"],fill:C.lorange,stroke:C.orange},
  ];
  personas.forEach((p,i)=>{
    const x=0.35+i*3.2, y=1.15;
    sl.addShape(pres.shapes.RECTANGLE, {x,y,w:3.0,h:4.2,fill:{color:p.fill},line:{color:p.stroke,width:1.5},shadow:makeShadow()});
    sl.addShape(pres.shapes.RECTANGLE, {x,y,w:3.0,h:0.72,fill:{color:p.stroke}});
    sl.addText(p.role, {x:x+0.08,y:y+0.12,w:2.84,h:0.48,fontFace:FONT,fontSize:14,bold:true,color:C.white,valign:"middle"});
    sl.addText(p.desc, {x:x+0.08,y:y+0.82,w:2.84,h:0.45,fontFace:FONT,fontSize:9.5,color:C.gray,wrap:true});
    p.tasks.forEach((t,ti)=>{
      const ty=y+1.4+ti*0.85;
      sl.addShape(pres.shapes.OVAL, {x:x+0.1,y:ty+0.08,w:0.3,h:0.3,fill:{color:p.stroke}});
      sl.addText(String(ti+1), {x:x+0.1,y:ty+0.08,w:0.3,h:0.3,fontFace:FONT,fontSize:9,bold:true,color:C.white,align:"center",valign:"middle"});
      sl.addText(t, {x:x+0.5,y:ty+0.04,w:2.4,h:0.42,fontFace:FONT,fontSize:9.5,color:C.black,valign:"middle",wrap:true});
    });
  });
}

// ── Slide 3: 目次 ─────────────────────────────────────────────────────────
{
  const sl = addContent(pres, "目　次");
  const sections = [
    {n:"1",title:"初期セットアップ",sub:"管理者向け",fill:C.lblue,stroke:C.blue},
    {n:"2",title:"申込みフォームの使い方",sub:"申込者向け",fill:C.lgreen,stroke:C.green},
    {n:"3",title:"申込みデータの管理",sub:"事務局向け",fill:C.lorange,stroke:C.orange},
    {n:"4",title:"請求書メールの送信",sub:"事務局向け",fill:C.lred,stroke:C.red},
    {n:"5",title:"入金確認とお礼状送信",sub:"事務局向け",fill:C.lgreen,stroke:C.green},
    {n:"6",title:"メールクォータ管理・FAQ",sub:"管理者・全員向け",fill:C.lyellow,stroke:C.orange},
  ];
  sections.forEach((s,i)=>{
    const col=i%2, row=Math.floor(i/2);
    const x=0.35+col*4.85, y=1.18+row*1.38;
    sl.addShape(pres.shapes.RECTANGLE, {x,y,w:4.55,h:1.2,fill:{color:s.fill},line:{color:s.stroke,width:1.5},shadow:makeShadow()});
    sl.addShape(pres.shapes.RECTANGLE, {x,y,w:0.62,h:1.2,fill:{color:s.stroke}});
    sl.addText(s.n, {x,y,w:0.62,h:1.2,fontFace:FONT,fontSize:26,bold:true,color:C.white,align:"center",valign:"middle"});
    sl.addText(s.title, {x:x+0.72,y:y+0.2,w:3.7,h:0.5,fontFace:FONT,fontSize:15,bold:true,color:C.black});
    sl.addText(s.sub,   {x:x+0.72,y:y+0.72,w:3.7,h:0.35,fontFace:FONT,fontSize:11,color:s.stroke});
  });
}

// ── Slide 4: セクション1 区切り ───────────────────────────────────────────
addNavy(pres, "1.  初期セットアップ", "管理者向け — システムの初回設定手順");

// ── Slide 5: セットアップフロー ───────────────────────────────────────────
{
  const sl = addContent(pres, "初期セットアップの流れ");
  const steps = [
    {n:"1",label:"Infoシートに\n設定値を入力",fill:C.lblue,stroke:C.blue},
    {n:"2",label:"メニューから\nProject Initialize",fill:C.lgreen,stroke:C.green},
    {n:"3",label:"フォルダ・シート\n自動生成",fill:C.lorange,stroke:C.orange},
    {n:"4",label:"トリガー\n自動登録",fill:C.lyellow,stroke:C.orange},
    {n:"5",label:"Webアプリ\nとしてデプロイ",fill:C.lred,stroke:C.red},
    {n:"6",label:"公開URLを\n申込者に共有",fill:C.lgreen,stroke:C.green},
  ];
  const bw=1.4, bh=1.5, gap=0.17, sx=0.3, y=1.55;
  steps.forEach((s,i)=>{
    const x=sx+i*(bw+gap);
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x,y,w:bw,h:bh,fill:{color:s.fill},line:{color:s.stroke,width:1.5},rectRadius:0.07,shadow:makeShadow()});
    sl.addShape(pres.shapes.OVAL, {x:x+bw/2-0.22,y:y+0.06,w:0.44,h:0.44,fill:{color:s.stroke}});
    sl.addText(s.n, {x:x+bw/2-0.22,y:y+0.06,w:0.44,h:0.44,fontFace:FONT,fontSize:14,bold:true,color:C.white,align:"center",valign:"middle"});
    sl.addText(s.label, {x:x+0.06,y:y+0.6,w:bw-0.12,h:0.8,fontFace:FONT,fontSize:11,bold:true,color:C.black,align:"center",wrap:true});
    if(i<steps.length-1) arrowRight(sl, x+bw+0.02, y+bh/2, gap+0.01);
  });
  sl.addText("※ 一度完了すれば次回以降は不要。設定変更時は再デプロイが必要。", {x:0.35,y:3.22,w:9.3,h:0.35,fontFace:FONT,fontSize:10,color:C.orange});
}

// ── Slide 6: Infoシート設定 ───────────────────────────────────────────────
{
  const sl = addContent(pres, "Infoシート — 主要設定項目");
  const settings = [
    ["EVENT_NAME","第〇回 川口花火大会","イベント名（フォームタイトル等に表示）"],
    ["OFFICE_EMAIL","office@example.com","事務局メールアドレス（CC・通知先）"],
    ["ORG_NAME / ORG_REP","川口市花火実行委員会 / 川口太郎","主催団体名・代表者名"],
    ["START_DATE / END_DATE","2025年8月3日 19:00","イベント開催日時"],
    ["PAYMENT_DUE","2025年7月31日","入金期限日"],
    ["BANK_NAME / BANK_NO","○○銀行 ○○支店 / 普通1234567","振込先銀行情報"],
    ["PRICE_B 〜 PRICE_E","100000 等（税抜・円）","各区分の協賛金額"],
    ["ROOT_FOLDER_ID","1AbCdEfGh...（URLから取得）","保存先Google DriveフォルダID"],
    ["INVOICE_REG_NO","T1234567890123","インボイス登録番号"],
    ["MIN_MAIL_QUOTA","5（デフォルト）","メール送信停止閾値（残通数）"],
  ];
  const hdrRow = ["設定キー","入力例","説明"].map(c=>({text:c,options:{bold:true,fill:{color:C.navy},color:C.white,fontFace:FONT,fontSize:9}}));
  const rows = settings.map(r=>r.map(c=>({text:c,options:{fontFace:FONT,fontSize:9,color:C.black}})));
  sl.addTable([hdrRow,...rows], {x:0.35,y:1.15,w:9.3,h:4.2,colW:[2.6,2.5,4.2],border:{pt:0.5,color:"bbbbbb"},fill:{color:C.llblue}});
}

// ── Slide 7: セクション2 区切り ───────────────────────────────────────────
addNavy(pres, "2.  申込みフォームの使い方", "申込者向け — 協賛申込みの手順");

// ── Slide 8: フォームUI mockup ────────────────────────────────────────────
{
  const sl = addContent(pres, "フォーム画面の構成");
  // Browser mockup
  sl.addShape(pres.shapes.RECTANGLE, {x:0.35,y:1.1,w:9.3,h:4.3,fill:{color:C.lgray},line:{color:C.gray,width:0.5},shadow:makeShadow()});
  // Browser chrome
  sl.addShape(pres.shapes.RECTANGLE, {x:0.35,y:1.1,w:9.3,h:0.38,fill:{color:"e0e0e0"},line:{color:C.gray,width:0.5}});
  for(const [ox,oc] of [[0.52,C.red],[0.78,C.orange],[1.04,C.green]]) {
    sl.addShape(pres.shapes.OVAL, {x:ox,y:1.18,w:0.18,h:0.18,fill:{color:oc}});
  }
  sl.addText("https://script.google.com/macros/s/.../exec", {x:1.4,y:1.14,w:7.0,h:0.3,fontFace:"Courier New",fontSize:8,color:C.gray});
  // Form content area
  sl.addShape(pres.shapes.RECTANGLE, {x:0.35,y:1.48,w:9.3,h:3.92,fill:{color:C.white},line:{color:C.lblue,width:0.5}});
  // Header bar
  sl.addShape(pres.shapes.RECTANGLE, {x:0.35,y:1.48,w:9.3,h:0.55,fill:{color:C.navy}});
  sl.addText("川口花火大会 協賛申込みフォーム", {x:0.5,y:1.52,w:9,h:0.47,fontFace:FONT,fontSize:13,bold:true,color:C.white,align:"center",valign:"middle"});
  // Fields
  const fieldY = [2.12,2.55,2.98];
  const fieldLabels = ["会社名・団体名  ＊","会社名（フリガナ）  ＊","代表者役職・代表者名  ＊"];
  fieldLabels.forEach((f,i)=>{
    sl.addText(f, {x:0.6,y:fieldY[i],w:4.2,h:0.2,fontFace:FONT,fontSize:8.5,color:C.navy,bold:true});
    sl.addShape(pres.shapes.RECTANGLE, {x:0.6,y:fieldY[i]+0.2,w:4.2,h:0.23,fill:{color:C.llblue},line:{color:C.lblue,width:0.5}});
  });
  // Kubun buttons
  sl.addText("協賛区分  ＊", {x:5.0,y:2.12,w:4.2,h:0.2,fontFace:FONT,fontSize:8.5,color:C.navy,bold:true});
  const tiers = ["S","A","B","C","D","E"];
  tiers.forEach((t,i)=>{
    const bx=5.0+i*0.72, fill=i<2?C.lblue:C.lgreen, stroke=i<2?C.blue:C.green;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x:bx,y:2.35,w:0.65,h:0.35,fill:{color:fill},line:{color:stroke,width:1},rectRadius:0.04});
    sl.addText(t, {x:bx,y:2.35,w:0.65,h:0.35,fontFace:FONT,fontSize:10,bold:true,color:stroke,align:"center",valign:"middle"});
  });
  // Terms
  sl.addShape(pres.shapes.RECTANGLE, {x:0.6,y:2.82,w:8.65,h:0.48,fill:{color:C.llblue},line:{color:C.lblue,width:0.5}});
  sl.addText("利用規約  （スクロールすると同意チェックが有効になります）", {x:0.75,y:2.88,w:8.0,h:0.36,fontFace:FONT,fontSize:9,color:C.gray,valign:"middle"});
  // Submit button
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x:3.8,y:3.45,w:2.4,h:0.55,fill:{color:C.blue},line:{color:C.blue,width:0},rectRadius:0.06});
  sl.addText("確　認", {x:3.8,y:3.45,w:2.4,h:0.55,fontFace:FONT,fontSize:13,bold:true,color:C.white,align:"center",valign:"middle"});
}

// ── Slide 9: 入力項目 ─────────────────────────────────────────────────────
{
  const sl = addContent(pres, "入力項目の説明");
  const fields = [
    ["会社名・団体名","✓","法人・団体の正式名称","略称不可"],
    ["会社名（フリガナ）","✓","全角カタカナで入力","ひらがな・漢字不可"],
    ["代表者役職・代表者名","✓","例：代表取締役 山田太郎",""],
    ["代表者（フリガナ）","✓","全角カタカナ",""],
    ["担当者名","✓","連絡担当者の氏名",""],
    ["担当者名（フリガナ）","✓","全角カタカナ",""],
    ["郵便番号","✓","ハイフンなし7桁","例：3330852"],
    ["住所","✓","都道府県から番地まで",""],
    ["電話番号","✓","ハイフンなし10〜11桁","例：0482571111"],
    ["メールアドレス","✓","連絡先メールアドレス","請求書の送付先"],
    ["協賛区分","✓","S/A/B/C/D/E から選択","受付期間外は表示なし"],
    ["会社HP URL","","任意入力","https〜"],
    ["掲載名（任意）","","印刷物等への掲載名","空欄→会社名を使用"],
    ["利用規約への同意","✓","規約を最後まで読んでチェック","スクロール後に有効"],
  ];
  const hdrRow = ["項目名","必須","入力内容","注意点"].map(c=>({text:c,options:{bold:true,fill:{color:C.navy},color:C.white,fontFace:FONT,fontSize:9}}));
  const rows = fields.map(r=>r.map((c,ci)=>({text:c,options:{fontFace:FONT,fontSize:9,color:ci===1?(c==="✓"?C.green:C.gray):C.black,bold:ci===1&&c==="✓"}})));
  sl.addTable([hdrRow,...rows], {x:0.35,y:1.15,w:9.3,h:4.2,colW:[2.5,0.9,3.5,2.4],border:{pt:0.5,color:"bbbbbb"},fill:{color:C.llblue}});
}

// ── Slide 10: 協賛区分 ───────────────────────────────────────────────────
{
  const sl = addContent(pres, "協賛区分の選択");
  // S/A label
  sl.addShape(pres.shapes.RECTANGLE, {x:0.35,y:1.12,w:4.3,h:0.35,fill:{color:C.blue}});
  sl.addText("S・A区分 — 抽選制", {x:0.45,y:1.14,w:4.1,h:0.31,fontFace:FONT,fontSize:12,bold:true,color:C.white});
  sl.addShape(pres.shapes.RECTANGLE, {x:4.85,y:1.12,w:4.8,h:0.35,fill:{color:C.green}});
  sl.addText("B〜E区分 — 先着制", {x:4.95,y:1.14,w:4.6,h:0.31,fontFace:FONT,fontSize:12,bold:true,color:C.white});

  const tiers = [
    {name:"S",fill:C.lblue,stroke:C.blue,method:"抽選制",invoice:"抽選後\n当選通知と同時"},
    {name:"A",fill:C.lblue,stroke:C.blue,method:"抽選制",invoice:"抽選後\n当選通知と同時"},
    {name:"B",fill:C.lgreen,stroke:C.green,method:"先着制",invoice:"申込み直後\n自動送信"},
    {name:"C",fill:C.lgreen,stroke:C.green,method:"先着制",invoice:"申込み直後\n自動送信"},
    {name:"D",fill:C.lgreen,stroke:C.green,method:"先着制",invoice:"申込み直後\n自動送信"},
    {name:"E",fill:C.lgreen,stroke:C.green,method:"先着制",invoice:"申込み直後\n自動送信"},
  ];
  tiers.forEach((t,i)=>{
    const x=i<2?0.35+i*2.25:4.85+(i-2)*1.25;
    const w=i<2?2.0:1.1;
    const y=1.52, h=3.8;
    sl.addShape(pres.shapes.RECTANGLE, {x,y,w,h,fill:{color:t.fill},line:{color:t.stroke,width:1.5},shadow:makeShadow()});
    sl.addShape(pres.shapes.RECTANGLE, {x,y,w,h:0.55,fill:{color:t.stroke}});
    sl.addText(t.name+" 区分", {x,y:y+0.05,w,h:0.45,fontFace:FONT,fontSize:i<2?16:13,bold:true,color:C.white,align:"center"});
    sl.addText("申込み方式", {x:x+0.06,y:y+0.65,w:w-0.12,h:0.28,fontFace:FONT,fontSize:8,bold:true,color:t.stroke});
    sl.addText(t.method, {x:x+0.06,y:y+0.92,w:w-0.12,h:0.45,fontFace:FONT,fontSize:i<2?15:12,bold:true,color:C.black,align:"center"});
    sl.addText("請求書送信", {x:x+0.06,y:y+1.5,w:w-0.12,h:0.28,fontFace:FONT,fontSize:8,bold:true,color:t.stroke});
    sl.addText(t.invoice, {x:x+0.06,y:y+1.78,w:w-0.12,h:0.7,fontFace:FONT,fontSize:i<2?11:9,color:C.black,align:"center",wrap:true});
    sl.addText("申込み期間", {x:x+0.06,y:y+2.55,w:w-0.12,h:0.28,fontFace:FONT,fontSize:8,bold:true,color:t.stroke});
    const period = i<2?"KUBUN_SA_\nSTART/END":"KUBUN_BCDE_\nSTART/END";
    sl.addText(period, {x:x+0.06,y:y+2.82,w:w-0.12,h:0.75,fontFace:"Courier New",fontSize:i<2?8:7,color:t.stroke,align:"center",wrap:true});
  });
}

// ── Slide 11: 申込み後の流れ ─────────────────────────────────────────────
{
  const sl = addContent(pres, "申込み後の流れ");
  // Main flow
  const mainSteps = [
    {label:"送信ボタン\nを押す",fill:C.lorange,stroke:C.orange},
    {label:"受付番号が\n画面に表示",fill:C.lgreen,stroke:C.green},
  ];
  mainSteps.forEach((s,i)=>{
    const x=0.4+i*4.5, y=1.5;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x,y,w:3.8,h:0.9,fill:{color:s.fill},line:{color:s.stroke,width:1.5},rectRadius:0.07,shadow:makeShadow()});
    sl.addText(s.label, {x,y,w:3.8,h:0.9,fontFace:FONT,fontSize:14,bold:true,color:C.black,align:"center",valign:"middle",wrap:true});
    if(i===0) arrowRight(sl, x+3.8+0.05, y+0.45, 0.65);
  });
  // Branch point
  sl.addText("区分により分岐", {x:3.8,y:2.62,w:2.4,h:0.28,fontFace:FONT,fontSize:10,color:C.gray,align:"center"});
  arrowDown(sl, 2.3, 2.4, 0.55);
  arrowDown(sl, 7.7, 2.4, 0.55);
  // B-E branch
  sl.addShape(pres.shapes.RECTANGLE, {x:0.35,y:3.05,w:4.3,h:0.38,fill:{color:C.green}});
  sl.addText("B〜E区分", {x:0.45,y:3.07,w:4.1,h:0.34,fontFace:FONT,fontSize:12,bold:true,color:C.white});
  const beBranch = [
    "申込み完了後すぐに\n請求書PDF添付メールが届く",
    "メールに記載の口座へ\n期限内にお振込みください",
    "入金確認後、事務局より\nお礼状メールが届きます",
  ];
  beBranch.forEach((b,i)=>{
    const y=3.55+i*0.6;
    sl.addShape(pres.shapes.OVAL, {x:0.38,y:y+0.08,w:0.28,h:0.28,fill:{color:C.green}});
    sl.addText(String(i+1), {x:0.38,y:y+0.08,w:0.28,h:0.28,fontFace:FONT,fontSize:9,bold:true,color:C.white,align:"center",valign:"middle"});
    sl.addText(b, {x:0.76,y:y+0.02,w:3.8,h:0.52,fontFace:FONT,fontSize:9.5,color:C.black,valign:"middle",wrap:true});
  });
  // S-A branch
  sl.addShape(pres.shapes.RECTANGLE, {x:5.35,y:3.05,w:4.3,h:0.38,fill:{color:C.blue}});
  sl.addText("S・A区分", {x:5.45,y:3.07,w:4.1,h:0.34,fontFace:FONT,fontSize:12,bold:true,color:C.white});
  const saBranch = [
    "申込み後は受付完了の\nメールが届きます（PDF無）",
    "抽選結果が決まり次第\n別途当落メールをお送りします",
    "当選の場合、請求書PDF付き\nメールが届きます",
  ];
  saBranch.forEach((b,i)=>{
    const y=3.55+i*0.6;
    sl.addShape(pres.shapes.OVAL, {x:5.38,y:y+0.08,w:0.28,h:0.28,fill:{color:C.blue}});
    sl.addText(String(i+1), {x:5.38,y:y+0.08,w:0.28,h:0.28,fontFace:FONT,fontSize:9,bold:true,color:C.white,align:"center",valign:"middle"});
    sl.addText(b, {x:5.76,y:y+0.02,w:3.8,h:0.52,fontFace:FONT,fontSize:9.5,color:C.black,valign:"middle",wrap:true});
  });
}

// ── Slide 12: セクション3 区切り ──────────────────────────────────────────
addNavy(pres, "3.  申込みデータの管理", "事務局向け — スプレッドシートの操作手順");

// ── Slide 13: 手作業シート ─────────────────────────────────────────────────
{
  const sl = addContent(pres, "手作業シートの見方");
  // Column mockup
  const cols = [
    {c:"A",l:"受付番号",t:"手入力",fill:C.lorange,stroke:C.orange},
    {c:"B",l:"区分",t:"自動",fill:C.lblue,stroke:C.blue},
    {c:"C",l:"電話番号",t:"自動",fill:C.lblue,stroke:C.blue},
    {c:"D",l:"会社名",t:"自動",fill:C.lblue,stroke:C.blue},
    {c:"E",l:"住所",t:"自動",fill:C.lblue,stroke:C.blue},
    {c:"F",l:"代表者",t:"自動",fill:C.lblue,stroke:C.blue},
    {c:"G",l:"メール",t:"自動",fill:C.lblue,stroke:C.blue},
    {c:"H",l:"URL",t:"自動",fill:C.lblue,stroke:C.blue},
    {c:"I",l:"受付完了",t:"☑",fill:C.lgreen,stroke:C.green},
    {c:"J",l:"請求書\n送信日時",t:"自動",fill:C.lyellow,stroke:C.orange},
    {c:"K",l:"入金完了",t:"☑",fill:C.lgreen,stroke:C.green},
    {c:"L",l:"お礼状\n送信日時",t:"自動",fill:C.lyellow,stroke:C.orange},
  ];
  const cw=0.73, y0=1.15, hh=0.34, ch=0.82;
  cols.forEach((c,i)=>{
    const x=0.4+i*cw;
    // Header
    sl.addShape(pres.shapes.RECTANGLE, {x,y:y0,w:cw-0.05,h:hh,fill:{color:c.stroke}});
    sl.addText(c.c, {x,y:y0,w:cw-0.05,h:hh,fontFace:FONT,fontSize:12,bold:true,color:C.white,align:"center",valign:"middle"});
    // Cell
    sl.addShape(pres.shapes.RECTANGLE, {x,y:y0+hh,w:cw-0.05,h:ch,fill:{color:c.fill},line:{color:c.stroke,width:0.5}});
    sl.addText(c.l, {x,y:y0+hh+0.06,w:cw-0.05,h:0.45,fontFace:FONT,fontSize:8,bold:true,color:C.black,align:"center",wrap:true});
    sl.addText(c.t, {x,y:y0+hh+0.52,w:cw-0.05,h:0.28,fontFace:FONT,fontSize:8,color:c.stroke,align:"center"});
  });
  // Labels
  sl.addShape(pres.shapes.LINE, {x:0.4,y:y0+hh+ch+0.06,w:0.73,h:0,line:{color:C.orange,width:1.5}});
  sl.addText("← 手入力", {x:0.5,y:y0+hh+ch+0.1,w:1.4,h:0.25,fontFace:FONT,fontSize:9,color:C.orange,bold:true});
  sl.addShape(pres.shapes.LINE, {x:1.18,y:y0+hh+ch+0.06,w:5.84,h:0,line:{color:C.blue,width:1.5}});
  sl.addText("XLOOKUP 自動取得 →", {x:2.0,y:y0+hh+ch+0.1,w:3.8,h:0.25,fontFace:FONT,fontSize:9,color:C.blue,bold:true});
  sl.addShape(pres.shapes.LINE, {x:7.07,y:y0+hh+ch+0.06,w:1.46,h:0,line:{color:C.green,width:1.5}});
  sl.addText("チェックボックス", {x:7.1,y:y0+hh+ch+0.1,w:1.8,h:0.25,fontFace:FONT,fontSize:9,color:C.green,bold:true});
  sl.addShape(pres.shapes.LINE, {x:8.57,y:y0+hh+ch+0.06,w:1.14,h:0,line:{color:C.orange,width:1.5}});
  sl.addText("自動タイムスタンプ", {x:8.6,y:y0+hh+ch+0.1,w:1.0,h:0.35,fontFace:FONT,fontSize:7.5,color:C.orange,bold:true,wrap:true});

  // Sample rows
  const sampleData = [
    ["KWGC001","B","090-1234","株式会社ABC","東京都...","✓","2025/6/18","✓","2025/6/20"],
    ["KWGC002","S","03-5678","山田商会","埼玉県...","✓","2025/6/19","",""],
    ["KWGC003","A","048-9012","花火製作所","川口市...","","","",""],
  ];
  const sampleCols = [0,1,2,3,4,8,9,10,11];
  sampleData.forEach((row,ri)=>{
    cols.forEach((c,ci)=>{
      const x=0.4+ci*cw;
      const y=y0+hh+ch+0.45+ri*0.35;
      const dIdx = [0,1,2,3,4,null,null,null,5,6,7,8][ci];
      const val = dIdx!==null ? row[dIdx] : "";
      const bg = val==="✓" ? C.lgreen : (ri%2===0?C.white:C.llblue);
      sl.addShape(pres.shapes.RECTANGLE, {x,y,w:cw-0.05,h:0.3,fill:{color:bg},line:{color:C.lblue,width:0.3}});
      sl.addText(val, {x,y:y+0.04,w:cw-0.05,h:0.22,fontFace:FONT,fontSize:7.5,color:val==="✓"?C.green:C.black,align:"center"});
    });
  });
}

// ── Slide 14: 受付番号の入力 ─────────────────────────────────────────────
{
  const sl = addContent(pres, "受付番号の入力方法");
  const steps = [
    {n:"1",title:"申込み一覧シートを確認",body:"「協賛申込み一覧」シートを開き、新しい申込み行のA列（受付番号）を確認します。",fill:C.lblue,stroke:C.blue},
    {n:"2",title:"受付番号をコピー",body:"受付番号のセルを選択してコピーします（Ctrl+C / Cmd+C）。\n⚠️ 手入力は誤りの原因になるため必ずコピーしてください。",fill:C.lgreen,stroke:C.green},
    {n:"3",title:"手作業シートに貼り付け",body:"「手作業」シートを開き、空いているA列の行に貼り付けます。B〜H列のデータが自動で入力されます。",fill:C.lorange,stroke:C.orange},
    {n:"4",title:"内容を確認して次へ",body:"B〜H列の情報が正しいことを確認します。問題なければ次のステップ（請求書送信）へ進みます。",fill:C.lblue,stroke:C.blue},
  ];
  steps.forEach((s,i)=>{
    const y=1.15+i*1.08;
    sl.addShape(pres.shapes.RECTANGLE, {x:0.35,y,w:9.3,h:0.9,fill:{color:s.fill},line:{color:s.stroke,width:1.5},shadow:makeShadow()});
    stepBadge(sl, pres, s.n, 0.45, y+0.26, s.stroke);
    sl.addText(s.title, {x:0.95,y:y+0.04,w:8.6,h:0.35,fontFace:FONT,fontSize:12,bold:true,color:C.black});
    sl.addText(s.body,  {x:0.95,y:y+0.38,w:8.6,h:0.48,fontFace:FONT,fontSize:9.5,color:C.black,wrap:true});
  });
  sl.addText("受付番号例：KWGC0618120530123　（prefix + yyyyMMddHHmmssSSS）", {x:0.35,y:5.4,w:9.3,h:0.3,fontFace:"Courier New",fontSize:10,color:C.blue});
}

// ── Slide 15: セクション4 区切り ──────────────────────────────────────────
addNavy(pres, "4.  請求書メールの送信", "事務局向け — 請求書送信の確認手順");

// ── Slide 16: メール送信タイミング ────────────────────────────────────────
{
  const sl = addContent(pres, "メール送信タイミング一覧");
  const cols2 = [
    {title:"B〜E区分\n申込み直後（自動）",icon:"⚡",body:"フォーム送信と同時に\n自動でメールが送信されます",attach:"請求書PDF",fill:C.lgreen,stroke:C.green,note:"担当者の操作不要"},
    {title:"S・A区分\n受付完了チェック後",icon:"☑",body:"担当者がI列をチェックし\n確認ダイアログで送信",attach:"請求書PDF",fill:C.lblue,stroke:C.blue,note:"担当者が確認して送信"},
    {title:"全区分\n入金完了チェック後",icon:"💰",body:"担当者がK列をチェックし\n確認ダイアログで送信",attach:"お礼状PDF",fill:C.lorange,stroke:C.orange,note:"J列に日時があること"},
  ];
  cols2.forEach((c,i)=>{
    const x=0.35+i*3.2, y=1.15, w=3.0, h=4.2;
    sl.addShape(pres.shapes.RECTANGLE, {x,y,w,h,fill:{color:c.fill},line:{color:c.stroke,width:1.5},shadow:makeShadow()});
    sl.addShape(pres.shapes.RECTANGLE, {x,y,w,h:0.65,fill:{color:c.stroke}});
    sl.addText(c.title, {x:x+0.08,y:y+0.05,w:w-0.16,h:0.55,fontFace:FONT,fontSize:11,bold:true,color:C.white,wrap:true});
    sl.addText(c.body, {x:x+0.08,y:y+0.8,w:w-0.16,h:0.9,fontFace:FONT,fontSize:11,color:C.black,wrap:true,valign:"middle"});
    sl.addShape(pres.shapes.RECTANGLE, {x:x+0.08,y:y+1.8,w:w-0.16,h:0.04,fill:{color:c.stroke}});
    sl.addText("添付ファイル", {x:x+0.08,y:y+1.95,w:w-0.16,h:0.3,fontFace:FONT,fontSize:9,bold:true,color:c.stroke});
    sl.addText(c.attach, {x:x+0.08,y:y+2.28,w:w-0.16,h:0.45,fontFace:FONT,fontSize:13,bold:true,color:C.black,align:"center"});
    sl.addShape(pres.shapes.RECTANGLE, {x:x+0.08,y:y+2.85,w:w-0.16,h:0.04,fill:{color:c.stroke}});
    sl.addShape(pres.shapes.RECTANGLE, {x:x+0.08,y:y+3.05,w:w-0.16,h:0.88,fill:{color:c.stroke},transparency:70});
    sl.addText(c.note, {x:x+0.1,y:y+3.1,w:w-0.2,h:0.8,fontFace:FONT,fontSize:11,bold:true,color:c.stroke,align:"center",valign:"middle",wrap:true});
  });
}

// ── Slide 17: 請求書送信手順 ─────────────────────────────────────────────
{
  const sl = addContent(pres, "請求書メール送信の手順（S・A区分）");
  const steps = [
    {n:"1",body:"「手作業」シートで対象の行（B列が S または A の行）を確認します。",fill:C.lblue,stroke:C.blue},
    {n:"2",body:"I列（受付完了）のチェックボックスをクリックします。",fill:C.lgreen,stroke:C.green},
    {n:"3",body:"確認ダイアログが表示されます。会社名・区分・金額を確認します。",fill:C.lorange,stroke:C.orange},
    {n:"4",body:"「送信する」ボタンをクリックします。（「キャンセル」で操作を中止できます）",fill:C.lgreen,stroke:C.green},
    {n:"5",body:"J列（請求書送信日時）に送信日時が自動的に記録されます。",fill:C.lblue,stroke:C.blue},
  ];
  steps.forEach((s,i)=>{
    const y=1.15+i*0.82;
    sl.addShape(pres.shapes.RECTANGLE, {x:0.35,y,w:5.8,h:0.7,fill:{color:s.fill},line:{color:s.stroke,width:1.5},shadow:makeShadow()});
    stepBadge(sl, pres, s.n, 0.45, y+0.16, s.stroke);
    sl.addText(s.body, {x:0.95,y:y+0.06,w:5.1,h:0.58,fontFace:FONT,fontSize:10,color:C.black,valign:"middle",wrap:true});
  });
  // Dialog mockup
  sl.addShape(pres.shapes.RECTANGLE, {x:6.4,y:1.05,w:3.25,h:2.8,fill:{color:C.white},line:{color:C.blue,width:1.5},shadow:makeShadow()});
  sl.addShape(pres.shapes.RECTANGLE, {x:6.4,y:1.05,w:3.25,h:0.52,fill:{color:C.navy}});
  sl.addText("請求書送信の確認", {x:6.5,y:1.08,w:3.05,h:0.46,fontFace:FONT,fontSize:11,bold:true,color:C.white,align:"center",valign:"middle"});
  sl.addText([
    {text:"以下の申込者に送信します。\n", options:{fontSize:9,breakLine:false}},
    {text:"会社名：株式会社サンプル\n", options:{fontSize:9,breakLine:false}},
    {text:"区分：A区分　金額：¥50,000", options:{fontSize:9}},
  ], {x:6.5,y:1.65,w:3.05,h:0.95,fontFace:FONT,color:C.black,wrap:true});
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x:6.55,y:2.72,w:1.25,h:0.4,fill:{color:C.blue},line:{color:C.blue,width:0},rectRadius:0.05});
  sl.addText("送信する", {x:6.55,y:2.72,w:1.25,h:0.4,fontFace:FONT,fontSize:11,bold:true,color:C.white,align:"center",valign:"middle"});
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x:8.0,y:2.72,w:1.5,h:0.4,fill:{color:C.lgray},line:{color:C.gray,width:0},rectRadius:0.05});
  sl.addText("キャンセル", {x:8.0,y:2.72,w:1.5,h:0.4,fontFace:FONT,fontSize:11,bold:true,color:C.gray,align:"center",valign:"middle"});
  sl.addText("← 確認ダイアログ", {x:6.5,y:3.95,w:3.1,h:0.28,fontFace:FONT,fontSize:9,color:C.blue,align:"center"});
  sl.addText("送信済みの場合: J列に日時が入力されていれば送信完了。再チェックすると「送信済み」の警告が表示され二重送信を防止します。", {x:0.35,y:5.12,w:9.3,h:0.4,fontFace:FONT,fontSize:9.5,color:C.orange,wrap:true});
}

// ── Slide 18: セクション5 区切り ──────────────────────────────────────────
addNavy(pres, "5.  入金確認とお礼状送信", "事務局向け — 入金確認後の手順");

// ── Slide 19: 入金確認・お礼状 ───────────────────────────────────────────
{
  const sl = addContent(pres, "入金確認とお礼状送信の手順");
  const steps = [
    {n:"1",body:"銀行の入金確認を行い、振込人名と受付番号を照合します。",fill:C.lblue,stroke:C.blue},
    {n:"2",body:"「手作業」シートで対象の行を確認。J列（請求書送信日時）に値があることを確認します。",fill:C.lorange,stroke:C.orange,warn:"J列が空の場合はK列チェック不可！"},
    {n:"3",body:"K列（入金完了）のチェックボックスをクリックします。",fill:C.lgreen,stroke:C.green},
    {n:"4",body:"確認ダイアログが表示されます。会社名・区分を確認します。",fill:C.lorange,stroke:C.orange},
    {n:"5",body:"「送信する」をクリック。L列（お礼状送信日時）に送信日時が自動記録されます。",fill:C.lgreen,stroke:C.green},
  ];
  steps.forEach((s,i)=>{
    const y=1.15+i*0.82;
    sl.addShape(pres.shapes.RECTANGLE, {x:0.35,y,w:5.8,h:0.7,fill:{color:s.fill},line:{color:s.stroke,width:1.5},shadow:makeShadow()});
    stepBadge(sl, pres, s.n, 0.45, y+0.16, s.stroke);
    sl.addText(s.body, {x:0.95,y:y+0.06,w:5.1,h:0.58,fontFace:FONT,fontSize:10,color:C.black,valign:"middle",wrap:true});
    if(s.warn) sl.addText("⚠ "+s.warn, {x:0.95,y:y+0.5,w:5.0,h:0.2,fontFace:FONT,fontSize:8.5,color:C.red});
  });
  // Dialog mockup
  sl.addShape(pres.shapes.RECTANGLE, {x:6.4,y:1.05,w:3.25,h:2.8,fill:{color:C.white},line:{color:C.green,width:1.5},shadow:makeShadow()});
  sl.addShape(pres.shapes.RECTANGLE, {x:6.4,y:1.05,w:3.25,h:0.52,fill:{color:C.green}});
  sl.addText("お礼状送信の確認", {x:6.5,y:1.08,w:3.05,h:0.46,fontFace:FONT,fontSize:11,bold:true,color:C.white,align:"center",valign:"middle"});
  sl.addText([
    {text:"入金を確認しました。\n", options:{fontSize:9,breakLine:false}},
    {text:"お礼状メールを送信します。\n", options:{fontSize:9,breakLine:false}},
    {text:"会社名：山田商会\n区分：B区分", options:{fontSize:9}},
  ], {x:6.5,y:1.65,w:3.05,h:0.95,fontFace:FONT,color:C.black,wrap:true});
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x:6.55,y:2.72,w:1.25,h:0.4,fill:{color:C.green},line:{color:C.green,width:0},rectRadius:0.05});
  sl.addText("送信する", {x:6.55,y:2.72,w:1.25,h:0.4,fontFace:FONT,fontSize:11,bold:true,color:C.white,align:"center",valign:"middle"});
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {x:8.0,y:2.72,w:1.5,h:0.4,fill:{color:C.lgray},line:{color:C.gray,width:0},rectRadius:0.05});
  sl.addText("キャンセル", {x:8.0,y:2.72,w:1.5,h:0.4,fontFace:FONT,fontSize:11,bold:true,color:C.gray,align:"center",valign:"middle"});
  sl.addText("← お礼状送信ダイアログ", {x:6.5,y:3.95,w:3.1,h:0.28,fontFace:FONT,fontSize:9,color:C.green,align:"center"});
}

// ── Slide 20: メールクォータ ─────────────────────────────────────────────
{
  const sl = addContent(pres, "メールクォータの管理");
  sl.addText("Gmailの1日あたり送信上限（無料：50通/日）。自動で監視し、不足時に通知します。", {x:0.35,y:1.08,w:9.3,h:0.35,fontFace:FONT,fontSize:11,color:C.black});
  const states = [
    {state:"通　常",cond:"残り通数 >\nMIN_MAIL_QUOTA",behavior:"通常通り動作\n制限なし",action:"対応不要",fill:C.lgreen,stroke:C.green},
    {state:"警　告",cond:"残り通数 ≤\nMIN_MAIL_QUOTA",behavior:"管理者（OFFICE_EMAIL）\nに警告メールを送信\nフォームは引き続き受付",action:"翌日まで待つか\n手動対応に切替",fill:C.lorange,stroke:C.orange},
    {state:"上限超過",cond:"残り通数 = 0",behavior:"Webフォームの\n受付を停止\nエラー画面を表示",action:"翌日（太平洋時間\n深夜）のリセットを待つ",fill:C.lred,stroke:C.red},
  ];
  states.forEach((s,i)=>{
    const x=0.35+i*3.2, y=1.55, w=3.0, h=3.75;
    sl.addShape(pres.shapes.RECTANGLE, {x,y,w,h,fill:{color:s.fill},line:{color:s.stroke,width:1.5},shadow:makeShadow()});
    sl.addShape(pres.shapes.RECTANGLE, {x,y,w,h:0.55,fill:{color:s.stroke}});
    sl.addText(s.state, {x,y:y+0.05,w,h:0.45,fontFace:FONT,fontSize:16,bold:true,color:C.white,align:"center",valign:"middle"});
    sl.addText("条件", {x:x+0.1,y:y+0.65,w:w-0.2,h:0.25,fontFace:FONT,fontSize:9,bold:true,color:s.stroke});
    sl.addText(s.cond, {x:x+0.1,y:y+0.9,w:w-0.2,h:0.55,fontFace:FONT,fontSize:10,color:C.black,wrap:true});
    sl.addText("システムの動作", {x:x+0.1,y:y+1.55,w:w-0.2,h:0.25,fontFace:FONT,fontSize:9,bold:true,color:s.stroke});
    sl.addText(s.behavior, {x:x+0.1,y:y+1.8,w:w-0.2,h:0.75,fontFace:FONT,fontSize:10,color:C.black,wrap:true});
    sl.addText("対応方法", {x:x+0.1,y:y+2.65,w:w-0.2,h:0.25,fontFace:FONT,fontSize:9,bold:true,color:s.stroke});
    sl.addText(s.action, {x:x+0.1,y:y+2.9,w:w-0.2,h:0.6,fontFace:FONT,fontSize:10,color:C.black,wrap:true});
  });
}

// ── Slide 21: FAQ ─────────────────────────────────────────────────────────
{
  const sl = addContent(pres, "よくある質問（FAQ）");
  const faqs = [
    {q:"申込み後にメールが届きません。",a:"①迷惑メールフォルダをご確認ください。②事務局のGmailクォータが上限に達している場合があります。翌日以降に届くか、事務局から手動でご連絡します。"},
    {q:"フォームで区分が選択できません。",a:"申込み受付期間外の場合、該当区分は表示されません。InfoシートのKUBUN_SA_START/ENDまたはKUBUN_BCDE_START/ENDの日付設定をご確認ください。"},
    {q:"手作業シートにデータが表示されません（XLOOKUP が機能しない）。",a:"受付番号が正確に入力されているか確認してください。スペースや全角・半角の違いが原因のことが多いです。「協賛申込み一覧」シートからコピー＆ペーストすることを推奨します。"},
    {q:"請求書のインボイス登録番号が空欄です。",a:"InfoシートのINVOICE_REG_NOに正しい登録番号（T＋13桁）を入力してください。入力後、「WebHanabi」→「設定を同期」を実行してください。"},
    {q:"K列（入金完了）にチェックできません。",a:"J列（請求書送信日時）に値が入っているかを確認してください。J列が空の場合はK列をチェックできません。"},
  ];
  faqs.forEach((f,i)=>{
    const y=1.18+i*0.84;
    sl.addShape(pres.shapes.RECTANGLE, {x:0.35,y,w:9.3,h:0.72,fill:{color:i%2===0?C.llblue:C.white},line:{color:C.lblue,width:0.5}});
    sl.addText("Q", {x:0.35,y,w:0.38,h:0.72,fontFace:FONT,fontSize:16,bold:true,color:C.blue,align:"center",valign:"middle"});
    sl.addShape(pres.shapes.LINE, {x:0.78,y:y+0.08,w:0,h:0.56,line:{color:C.lblue,width:0.8}});
    sl.addText(f.q, {x:0.88,y:y+0.03,w:4.0,h:0.3,fontFace:FONT,fontSize:10,bold:true,color:C.navy,wrap:true});
    sl.addText("A: "+f.a, {x:0.88,y:y+0.33,w:8.65,h:0.34,fontFace:FONT,fontSize:8.5,color:C.black,wrap:true});
  });
}

// ── Slide 22: まとめ ─────────────────────────────────────────────────────
{
  const sl = addNavy(pres, "ご不明な点はお気軽に\n事務局へご連絡ください", "WebHanabi — 川口花火大会 協賛管理システム");
  const pts = [
    {icon:"★",text:"申込みフォームURL：事務局より共有済みのURLをご利用ください"},
    {icon:"★",text:"お問い合わせ先：InfoシートのOFFICE_EMAIL に設定されたアドレス"},
    {icon:"★",text:"操作不明時：本マニュアルの各セクションをご参照ください"},
  ];
  pts.forEach((p,i)=>{
    sl.addShape(pres.shapes.RECTANGLE, {x:0.5,y:3.28+i*0.6,w:9,h:0.48,fill:{color:C.blue},transparency:70});
    sl.addText(p.icon+" "+p.text, {x:0.65,y:3.3+i*0.6,w:8.8,h:0.44,fontFace:FONT,fontSize:11,color:C.white,valign:"middle"});
  });
}

pres.writeFile({fileName:"操作マニュアル_WebHanabi.pptx"})
  .then(()=>console.log("✓ 操作マニュアル_WebHanabi.pptx generated"))
  .catch(e=>console.error(e));

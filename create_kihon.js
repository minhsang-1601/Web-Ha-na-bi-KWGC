const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "基本設計書 WebHanabi";
pres.author = "WebHanabi";

const NAVY = "1a3a5c";
const BLUE = "2e6da4";
const LBLUE = "dce9f5";
const WHITE = "FFFFFF";
const ACCENT = "4a90c4";
const GREEN = "27ae60";

const makeShadow = () => ({ type: "outer", blur: 6, offset: 3, angle: 135, color: "000000", opacity: 0.15 });

// ─────────────────────────────────────────
// SLIDE 1 — Cover
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: NAVY };

  // decorative circles
  s.addShape(pres.shapes.OVAL, { x: -0.8, y: -0.8, w: 3, h: 3, fill: { color: BLUE, transparency: 60 }, line: { color: BLUE, width: 0 } });
  s.addShape(pres.shapes.OVAL, { x: 8.5, y: 3.5, w: 2.5, h: 2.5, fill: { color: LBLUE, transparency: 70 }, line: { color: LBLUE, width: 0 } });
  s.addShape(pres.shapes.OVAL, { x: 7.5, y: -0.5, w: 1.8, h: 1.8, fill: { color: ACCENT, transparency: 60 }, line: { color: ACCENT, width: 0 } });

  // accent bar on left
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.5, w: 0.12, h: 2.6, fill: { color: LBLUE }, line: { color: LBLUE, width: 0 } });

  // Title
  s.addText("基本設計書", {
    x: 0.9, y: 1.4, w: 8.5, h: 1.3,
    fontSize: 52, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0
  });

  // Subtitle
  s.addText("WebHanabi — 川口花火大会 協賛管理システム", {
    x: 0.9, y: 2.8, w: 8.5, h: 0.7,
    fontSize: 20, color: LBLUE, fontFace: "Meiryo", margin: 0
  });

  // Date
  s.addText("2026-06-18", {
    x: 0.9, y: 4.7, w: 4, h: 0.4,
    fontSize: 13, color: "aaaaaa", fontFace: "Meiryo", margin: 0
  });
}

// ─────────────────────────────────────────
// SLIDE 2 — 目次
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addText("目次", { x: 0.4, y: 0.1, w: 9, h: 0.6, fontSize: 28, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0 });

  const items = [
    "システム概要",
    "アーキテクチャ概要",
    "コンポーネント構成",
    "ファイル構成",
    "機能一覧",
    "協賛区分と申込み方式",
    "データモデル",
    "Infoシート設定キー",
    "外部インターフェース",
    "メールインターフェース",
    "非機能要件",
    "まとめ"
  ];

  const cols = 2;
  const perCol = Math.ceil(items.length / cols);
  const colX = [0.5, 5.2];
  const startY = 1.1;
  const rowH = 0.37;

  items.forEach((item, i) => {
    const col = Math.floor(i / perCol);
    const row = i % perCol;
    const x = colX[col];
    const y = startY + row * rowH;
    const num = i + 1;

    s.addShape(pres.shapes.OVAL, { x, y: y - 0.01, w: 0.3, h: 0.3, fill: { color: BLUE }, line: { color: BLUE, width: 0 } });
    s.addText(String(num), { x, y: y - 0.01, w: 0.3, h: 0.3, fontSize: 11, bold: true, color: WHITE, fontFace: "Meiryo", align: "center", valign: "middle", margin: 0 });
    s.addText(item, { x: x + 0.38, y, w: 4.2, h: 0.3, fontSize: 14, color: NAVY, fontFace: "Meiryo", margin: 0 });
  });

  // decorative shape
  s.addShape(pres.shapes.RECTANGLE, { x: 9.3, y: 0.85, w: 0.7, h: 4.775, fill: { color: LBLUE }, line: { color: LBLUE, width: 0 } });
}

// ─────────────────────────────────────────
// SLIDE 3 — システム概要
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addText("システム概要", { x: 0.4, y: 0.1, w: 9, h: 0.6, fontSize: 28, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0 });

  // Left text
  s.addText([
    { text: "WebHanabi システムは、川口花火大会の", options: { breakLine: true } },
    { text: "協賛申込みを一元管理するシステムです。", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "主な機能:", options: { bold: true, breakLine: true } },
    { text: "Webフォームによる申込み受付", options: { bullet: true, breakLine: true } },
    { text: "スプレッドシートによるデータ管理", options: { bullet: true, breakLine: true } },
    { text: "PDF請求書・感謝状の自動生成", options: { bullet: true, breakLine: true } },
    { text: "Gmailによる自動メール送信", options: { bullet: true, breakLine: true } },
    { text: "担当者向けスタッフワークフロー", options: { bullet: true } },
  ], { x: 0.4, y: 1.0, w: 4.4, h: 4.0, fontSize: 14, color: NAVY, fontFace: "Meiryo" });

  // Right: flow diagram
  const flowItems = ["申込者", "Webフォーム", "スプレッドシート", "Gmail", "担当者"];
  const flowColors = [ACCENT, BLUE, "1e7e34", "d04a02", NAVY];
  const fx = 5.4;
  const fw = 3.8;
  const fh = 0.55;
  const gap = 0.2;
  const totalH = flowItems.length * fh + (flowItems.length - 1) * (gap + 0.2);
  const startY = (5.625 - 0.85 - totalH) / 2 + 0.85;

  flowItems.forEach((label, i) => {
    const fy = startY + i * (fh + gap + 0.2);
    s.addShape(pres.shapes.RECTANGLE, {
      x: fx, y: fy, w: fw, h: fh,
      fill: { color: flowColors[i] }, line: { color: flowColors[i], width: 0 },
      shadow: makeShadow()
    });
    s.addText(label, { x: fx, y: fy, w: fw, h: fh, fontSize: 15, bold: true, color: WHITE, fontFace: "Meiryo", align: "center", valign: "middle", margin: 0 });

    if (i < flowItems.length - 1) {
      const arrowY = fy + fh;
      s.addShape(pres.shapes.RECTANGLE, { x: fx + fw / 2 - 0.04, y: arrowY, w: 0.08, h: 0.15, fill: { color: "999999" }, line: { color: "999999", width: 0 } });
      s.addText("▼", { x: fx + fw / 2 - 0.2, y: arrowY + 0.08, w: 0.4, h: 0.2, fontSize: 10, color: "999999", fontFace: "Meiryo", align: "center", margin: 0 });
    }
  });
}

// ─────────────────────────────────────────
// SLIDE 4 — アーキテクチャ概要
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LBLUE };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addText("アーキテクチャ概要", { x: 0.4, y: 0.1, w: 9, h: 0.6, fontSize: 28, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0 });

  // Left text
  s.addText([
    { text: "サーバーレス GAS アーキテクチャ", options: { bold: true, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Google Apps Script (GAS) をバックエンドとして使用し、インフラ管理不要なサーバーレス構成を実現。", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "特徴:", options: { bold: true, breakLine: true } },
    { text: "ランニングコストがほぼゼロ", options: { bullet: true, breakLine: true } },
    { text: "Google エコシステムとの完全統合", options: { bullet: true, breakLine: true } },
    { text: "スケーリング不要（GAS が自動管理）", options: { bullet: true, breakLine: true } },
    { text: "OAuth 認証を Google が管理", options: { bullet: true } },
  ], { x: 0.4, y: 1.0, w: 4.5, h: 4.0, fontSize: 13, color: NAVY, fontFace: "Meiryo" });

  // Right: stack visualization
  const stack = [
    { label: "HtmlService", color: ACCENT },
    { label: "Google Apps Script", color: BLUE },
    { label: "Google Sheets", color: "0f9d58" },
    { label: "Google Drive", color: "fbbc05" },
    { label: "Gmail", color: "ea4335" },
  ];
  const sx = 5.3;
  const sw = 4.1;
  const sh = 0.62;
  const sgap = 0.08;
  const startY2 = 1.15;

  stack.forEach((item, i) => {
    const sy = startY2 + i * (sh + sgap);
    s.addShape(pres.shapes.RECTANGLE, {
      x: sx, y: sy, w: sw, h: sh,
      fill: { color: item.color }, line: { color: item.color, width: 0 },
      shadow: makeShadow()
    });
    s.addText(item.label, { x: sx, y: sy, w: sw, h: sh, fontSize: 15, bold: true, color: WHITE, fontFace: "Meiryo", align: "center", valign: "middle", margin: 0 });
  });
}

// ─────────────────────────────────────────
// SLIDE 5 — コンポーネント構成
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addText("コンポーネント構成", { x: 0.4, y: 0.1, w: 9, h: 0.6, fontSize: 28, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0 });

  const components = [
    { name: "Webフォーム", tech: "HtmlService", role: "申込者向けUIの提供" },
    { name: "バックエンドAPI", tech: "Google Apps Script", role: "フォーム処理・ビジネスロジック" },
    { name: "データストア", tech: "Google Sheets", role: "申込みデータの永続化・管理" },
    { name: "ファイルストレージ", tech: "Google Drive", role: "PDF請求書・感謝状の保存" },
    { name: "メール配信", tech: "Gmail API", role: "自動メール送信・通知" },
    { name: "設定管理", tech: "Infoシート", role: "システム設定値の一元管理" },
  ];

  const cardW = 2.9;
  const cardH = 1.4;
  const startX = 0.35;
  const startY = 1.05;
  const colGap = 0.25;
  const rowGap = 0.3;

  components.forEach((c, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const cx = startX + col * (cardW + colGap);
    const cy = startY + row * (cardH + rowGap);

    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: cardW, h: cardH, fill: { color: LBLUE }, line: { color: BLUE, width: 1 }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: cardW, h: 0.36, fill: { color: BLUE }, line: { color: BLUE, width: 0 } });
    s.addText(c.name, { x: cx, y: cy, w: cardW, h: 0.36, fontSize: 13, bold: true, color: WHITE, fontFace: "Meiryo", align: "center", valign: "middle", margin: 0 });
    s.addText([
      { text: c.tech, options: { bold: true, breakLine: true } },
      { text: c.role }
    ], { x: cx + 0.1, y: cy + 0.42, w: cardW - 0.2, h: cardH - 0.52, fontSize: 11.5, color: NAVY, fontFace: "Meiryo" });
  });
}

// ─────────────────────────────────────────
// SLIDE 6 — ファイル構成
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addText("ファイル構成", { x: 0.4, y: 0.1, w: 9, h: 0.6, fontSize: 28, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0 });

  const gsFiles = [
    { name: "WebApp.gs", desc: "Web App エントリ（doGet・getConfig）" },
    { name: "Sheet.gs", desc: "スプレッドシート書込・受付番号生成" },
    { name: "Mail.gs", desc: "メール送信・PDF生成・クォータ確認" },
    { name: "Trigger.gs", desc: "onEditトリガー（請求書・お礼状送信）" },
    { name: "Config.gs", desc: "定数・設定値（シート名・Info）" },
    { name: "ProjectInit.gs", desc: "プロジェクト初期化（フォルダ・SS・トリガー）" },
    { name: "Setup.gs", desc: "初回セットアップ（Infoシート作成等）" },
  ];

  const htmlFiles = [
    { name: "index.html", desc: "フォームUI（HTML・JS・確認/完了画面）" },
    { name: "Stylesheet.html", desc: "CSSスタイル定義" },
    { name: "invoice-template.html", desc: "請求書PDFテンプレート" },
    { name: "oreijou-template.html", desc: "お礼状PDFテンプレート" },
    { name: "ConfirmInvoiceDialog.html", desc: "請求書送信確認ダイアログ" },
    { name: "ConfirmNyukinDialog.html", desc: "入金確認ダイアログ" },
    { name: "ConfirmOreijouDialog.html", desc: "お礼状送信確認ダイアログ" },
  ];

  // Left column header
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 0.9, w: 4.2, h: 0.38, fill: { color: BLUE }, line: { color: BLUE, width: 0 } });
  s.addText(".gs ファイル (GAS)", { x: 0.3, y: 0.9, w: 4.2, h: 0.38, fontSize: 13, bold: true, color: WHITE, fontFace: "Meiryo", align: "center", valign: "middle", margin: 0 });

  // Right column header
  s.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 0.9, w: 4.2, h: 0.38, fill: { color: ACCENT }, line: { color: ACCENT, width: 0 } });
  s.addText("HTML / CSS ファイル", { x: 5.3, y: 0.9, w: 4.2, h: 0.38, fontSize: 13, bold: true, color: WHITE, fontFace: "Meiryo", align: "center", valign: "middle", margin: 0 });

  const rowH = 0.55;
  gsFiles.forEach((f, i) => {
    const fy = 1.38 + i * rowH;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: fy, w: 4.2, h: rowH - 0.05, fill: { color: i % 2 === 0 ? "f0f4f8" : WHITE }, line: { color: "dddddd", width: 0.5 } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: fy, w: 0.07, h: rowH - 0.05, fill: { color: BLUE }, line: { color: BLUE, width: 0 } });
    s.addText([
      { text: f.name, options: { bold: true, breakLine: true } },
      { text: f.desc }
    ], { x: 0.48, y: fy + 0.02, w: 4.0, h: rowH - 0.09, fontSize: 11, color: NAVY, fontFace: "Meiryo" });
  });

  htmlFiles.forEach((f, i) => {
    const fy = 1.38 + i * rowH;
    s.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: fy, w: 4.2, h: rowH - 0.05, fill: { color: i % 2 === 0 ? "f0f4f8" : WHITE }, line: { color: "dddddd", width: 0.5 } });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: fy, w: 0.07, h: rowH - 0.05, fill: { color: ACCENT }, line: { color: ACCENT, width: 0 } });
    s.addText([
      { text: f.name, options: { bold: true, breakLine: true } },
      { text: f.desc }
    ], { x: 5.48, y: fy + 0.02, w: 4.0, h: rowH - 0.09, fontSize: 11, color: NAVY, fontFace: "Meiryo" });
  });
}

// ─────────────────────────────────────────
// SLIDE 7 — 機能一覧
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addText("機能一覧", { x: 0.4, y: 0.1, w: 9, h: 0.6, fontSize: 28, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0 });

  const funcs = [
    ["F-01", "申込みフォーム表示", "協賛区分に応じた申込フォームをWebで提供", "高"],
    ["F-02", "申込みデータ保存", "フォーム入力をスプレッドシートに記録", "高"],
    ["F-03", "申込み確認メール送信", "B-E区分申込み後に確認メールを自動送信", "高"],
    ["F-04", "受付完了メール送信", "S-A区分の受付完了後にメールを送信", "高"],
    ["F-05", "入金確認メール送信", "入金確認後に領収・感謝メールを送信", "高"],
    ["F-06", "PDF請求書生成", "Google Drive上にPDF請求書を自動生成", "高"],
    ["F-07", "PDF感謝状生成", "入金完了後にPDF感謝状を自動生成", "中"],
    ["F-08", "協賛区分管理", "S/A/B/C/D/E区分と抽選・先着方式の管理", "高"],
    ["F-09", "抽選処理", "S-A区分の抽選申込み管理", "高"],
    ["F-10", "設定管理", "Infoシートによるシステム設定の一元管理", "中"],
    ["F-11", "担当者ワークフロー", "手作業シートによるスタッフ作業管理", "中"],
  ];

  const headers = [
    { text: "機能ID", options: { fill: { color: NAVY }, color: WHITE, bold: true } },
    { text: "機能名", options: { fill: { color: NAVY }, color: WHITE, bold: true } },
    { text: "概要", options: { fill: { color: NAVY }, color: WHITE, bold: true } },
    { text: "優先度", options: { fill: { color: NAVY }, color: WHITE, bold: true } },
  ];

  const rows = [headers, ...funcs.map((f, i) => [
    { text: f[0], options: { fill: { color: i % 2 === 0 ? "eef4fb" : WHITE }, color: BLUE, bold: true } },
    { text: f[1], options: { fill: { color: i % 2 === 0 ? "eef4fb" : WHITE }, color: NAVY, bold: true } },
    { text: f[2], options: { fill: { color: i % 2 === 0 ? "eef4fb" : WHITE }, color: "444444" } },
    { text: f[3], options: { fill: { color: i % 2 === 0 ? "eef4fb" : WHITE }, color: f[3] === "高" ? "c0392b" : "2980b9", bold: true } },
  ])];

  s.addTable(rows, {
    x: 0.3, y: 0.95, w: 9.4, colW: [0.85, 1.85, 5.6, 0.9],
    border: { pt: 0.5, color: "cccccc" },
    fontFace: "Meiryo", fontSize: 11
  });
}

// ─────────────────────────────────────────
// SLIDE 8 — 協賛区分と申込み方式
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addText("協賛区分と申込み方式", { x: 0.4, y: 0.1, w: 9, h: 0.6, fontSize: 28, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0 });

  const tiers = [
    { tier: "S", method: "抽選", key: "lotStartS / lotEndS", color: NAVY },
    { tier: "A", method: "抽選", key: "lotStartA / lotEndA", color: BLUE },
    { tier: "B", method: "先着", key: "startB / endB", color: ACCENT },
    { tier: "C", method: "先着", key: "startC / endC", color: ACCENT },
    { tier: "D", method: "先着", key: "startD / endD", color: "5ba3c9" },
    { tier: "E", method: "先着", key: "startE / endE", color: "5ba3c9" },
  ];

  const cardW = 2.9;
  const cardH = 1.6;
  const startX = 0.35;
  const startY = 1.0;

  tiers.forEach((t, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const cx = startX + col * (cardW + 0.2);
    const cy = startY + row * (cardH + 0.3);

    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: cardW, h: cardH, fill: { color: WHITE }, line: { color: t.color, width: 2 }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: cardW, h: 0.5, fill: { color: t.color }, line: { color: t.color, width: 0 } });
    s.addText(`区分 ${t.tier}`, { x: cx, y: cy, w: cardW, h: 0.5, fontSize: 18, bold: true, color: WHITE, fontFace: "Meiryo", align: "center", valign: "middle", margin: 0 });
    s.addText([
      { text: "方式: ", options: { bold: true } },
      { text: t.method, options: { breakLine: true } },
      { text: "設定キー: ", options: { bold: true } },
      { text: t.key }
    ], { x: cx + 0.15, y: cy + 0.58, w: cardW - 0.3, h: cardH - 0.65, fontSize: 12, color: NAVY, fontFace: "Meiryo" });
  });
}

// ─────────────────────────────────────────
// SLIDE 9 — データモデル: 協賛申込み一覧シート
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addText("データモデル: 協賛申込み一覧シート", { x: 0.4, y: 0.1, w: 9, h: 0.6, fontSize: 24, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0 });

  const cols9 = [
    ["A", "タイムスタンプ", "Datetime", "申込み日時"],
    ["B", "企業名", "String", "申込み企業名"],
    ["C", "担当者名", "String", "企業担当者名"],
    ["D", "メールアドレス", "String", "連絡先メール"],
    ["E", "電話番号", "String", "連絡先電話"],
    ["F", "協賛区分", "String", "S/A/B/C/D/E"],
    ["G", "申込み方式", "String", "抽選 or 先着"],
    ["H", "金額", "Number", "協賛金額（円）"],
    ["I", "請求書番号", "String", "自動採番"],
    ["J", "請求書PDF", "URL", "Drive リンク"],
    ["K", "感謝状PDF", "URL", "Drive リンク"],
    ["L", "入金確認", "Boolean", "入金済みフラグ"],
    ["M", "入金日", "Date", "入金確認日"],
    ["N", "メール送信済み", "Boolean", "送信フラグ"],
    ["O", "備考", "String", "担当者メモ"],
  ];

  const headers9 = [
    { text: "列", options: { fill: { color: NAVY }, color: WHITE, bold: true } },
    { text: "カラム名", options: { fill: { color: NAVY }, color: WHITE, bold: true } },
    { text: "型", options: { fill: { color: NAVY }, color: WHITE, bold: true } },
    { text: "説明", options: { fill: { color: NAVY }, color: WHITE, bold: true } },
  ];

  const rows9 = [headers9, ...cols9.map((c, i) => [
    { text: c[0], options: { fill: { color: i % 2 === 0 ? "eef4fb" : WHITE }, color: BLUE, bold: true } },
    { text: c[1], options: { fill: { color: i % 2 === 0 ? "eef4fb" : WHITE }, color: NAVY, bold: true } },
    { text: c[2], options: { fill: { color: i % 2 === 0 ? "eef4fb" : WHITE }, color: "555555", italic: true } },
    { text: c[3], options: { fill: { color: i % 2 === 0 ? "eef4fb" : WHITE }, color: "444444" } },
  ])];

  s.addTable(rows9, {
    x: 0.3, y: 0.92, w: 9.4, colW: [0.45, 2.1, 1.2, 5.55],
    border: { pt: 0.5, color: "cccccc" },
    fontFace: "Meiryo", fontSize: 11
  });
}

// ─────────────────────────────────────────
// SLIDE 10 — データモデル: 手作業シート
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addText("データモデル: 手作業シート", { x: 0.4, y: 0.1, w: 9, h: 0.6, fontSize: 26, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0 });

  const cols10 = [
    ["A", "企業名", "String", "XLOOKUP元"],
    ["B", "担当者名", "String", "XLOOKUP元"],
    ["C", "協賛区分", "String", "XLOOKUP元"],
    ["D", "金額", "Number", "XLOOKUP元"],
    ["E", "請求書番号", "String", "XLOOKUP元"],
    ["F", "連絡先メール", "String", "XLOOKUP元"],
    ["G", "入金予定日", "Date", "担当者記入"],
    ["H", "受付番号", "String", "担当者記入"],
    ["I", "受付完了", "Boolean", "チェックボックス"],
    ["J", "請求書送付日", "Date", "担当者記入"],
    ["K", "入金完了", "Boolean", "チェックボックス"],
    ["L", "備考", "String", "担当者メモ"],
  ];

  const headers10 = [
    { text: "列", options: { fill: { color: NAVY }, color: WHITE, bold: true } },
    { text: "カラム名", options: { fill: { color: NAVY }, color: WHITE, bold: true } },
    { text: "型", options: { fill: { color: NAVY }, color: WHITE, bold: true } },
    { text: "説明", options: { fill: { color: NAVY }, color: WHITE, bold: true } },
  ];

  const rows10 = [headers10, ...cols10.map((c, i) => {
    const isCheck = c[0] === "I" || c[0] === "K";
    return [
      { text: c[0], options: { fill: { color: isCheck ? "e8f8f0" : (i % 2 === 0 ? "eef4fb" : WHITE) }, color: isCheck ? GREEN : BLUE, bold: true } },
      { text: c[1], options: { fill: { color: isCheck ? "e8f8f0" : (i % 2 === 0 ? "eef4fb" : WHITE) }, color: isCheck ? GREEN : NAVY, bold: isCheck } },
      { text: c[2], options: { fill: { color: isCheck ? "e8f8f0" : (i % 2 === 0 ? "eef4fb" : WHITE) }, color: "555555", italic: true } },
      { text: c[3], options: { fill: { color: isCheck ? "e8f8f0" : (i % 2 === 0 ? "eef4fb" : WHITE) }, color: isCheck ? GREEN : "444444", bold: isCheck } },
    ];
  })];

  s.addTable(rows10, {
    x: 0.3, y: 0.92, w: 9.4, colW: [0.45, 2.1, 1.2, 5.55],
    border: { pt: 0.5, color: "cccccc" },
    fontFace: "Meiryo", fontSize: 11
  });

  // XLOOKUP note
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.25, w: 9.4, h: 0.28, fill: { color: "fff3cd" }, line: { color: "ffc107", width: 1 } });
  s.addText("注: A-F列はXLOOKUP関数により協賛申込み一覧シートから自動参照。I列・K列はチェックボックス（緑色）。", {
    x: 0.35, y: 5.25, w: 9.3, h: 0.28, fontSize: 10.5, color: "7d6608", fontFace: "Meiryo", valign: "middle", margin: 0
  });
}

// ─────────────────────────────────────────
// SLIDE 11 — Infoシート設定キー
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LBLUE };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addText("Infoシート設定キー", { x: 0.4, y: 0.1, w: 9, h: 0.6, fontSize: 28, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0 });

  const keys = [
    ["sheetName", "申込みデータを格納するシート名"],
    ["manualSheetName", "手作業シートの名前"],
    ["startB / endB", "B区分の申込み受付期間"],
    ["startC / endC", "C区分の申込み受付期間"],
    ["startD / endD", "D区分の申込み受付期間"],
    ["startE / endE", "E区分の申込み受付期間"],
    ["lotStartS / lotEndS", "S区分の抽選申込み期間"],
    ["lotStartA / lotEndA", "A区分の抽選申込み期間"],
    ["pdfFolderID", "PDF保存先Google DriveフォルダID"],
    ["mailFrom", "送信元メールアドレス"],
    ["adminEmail", "管理者通知先メール"],
    ["eventName", "イベント名（メール・PDF内で使用）"],
  ];

  const colW = 4.4;
  const rowH = 0.38;
  const startY = 0.95;
  const perCol = 6;

  keys.forEach((kv, i) => {
    const col = Math.floor(i / perCol);
    const row = i % perCol;
    const kx = 0.3 + col * (colW + 0.6);
    const ky = startY + row * rowH;

    s.addShape(pres.shapes.RECTANGLE, { x: kx, y: ky, w: colW, h: rowH - 0.04, fill: { color: i % 2 === 0 ? WHITE : "e8f0fa" }, line: { color: "bbccdd", width: 0.5 } });
    s.addText([
      { text: kv[0], options: { bold: true } },
      { text: " — " + kv[1] }
    ], { x: kx + 0.1, y: ky + 0.03, w: colW - 0.2, h: rowH - 0.1, fontSize: 11, color: NAVY, fontFace: "Meiryo", margin: 0 });
  });
}

// ─────────────────────────────────────────
// SLIDE 12 — 外部インターフェース: Webフォーム
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addText("外部インターフェース: Webフォーム", { x: 0.4, y: 0.1, w: 9, h: 0.6, fontSize: 24, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0 });

  const kvItems = [
    { key: "URL形式", value: "https://script.google.com/macros/s/{SCRIPT_ID}/exec" },
    { key: "アクセス制御", value: "Google アカウント不要（公開アクセス）" },
    { key: "レスポンス形式", value: "HTML（HtmlService.createHtmlOutput）" },
    { key: "プロトコル", value: "HTTPS（Google Infrastructure）" },
    { key: "リクエスト方式", value: "GET（初期表示）/ POST（申込み送信）" },
    { key: "タイムアウト", value: "GAS 標準 6分（最大実行時間）" },
    { key: "CORS", value: "同一オリジン（GAS スクリプトURL）" },
    { key: "認証方式", value: "なし（フォームは誰でもアクセス可能）" },
  ];

  const rowH = 0.56;
  const startY = 0.95;

  kvItems.forEach((item, i) => {
    const ry = startY + i * rowH;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: ry, w: 2.2, h: rowH - 0.05, fill: { color: BLUE }, line: { color: BLUE, width: 0 } });
    s.addText(item.key, { x: 0.3, y: ry, w: 2.2, h: rowH - 0.05, fontSize: 12, bold: true, color: WHITE, fontFace: "Meiryo", align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 2.5, y: ry, w: 7.2, h: rowH - 0.05, fill: { color: i % 2 === 0 ? "f0f4f8" : WHITE }, line: { color: "dddddd", width: 0.5 } });
    s.addText(item.value, { x: 2.6, y: ry + 0.02, w: 7.0, h: rowH - 0.1, fontSize: 12, color: NAVY, fontFace: "Meiryo", valign: "middle", margin: 0 });
  });
}

// ─────────────────────────────────────────
// SLIDE 13 — メールインターフェース
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addText("メールインターフェース", { x: 0.4, y: 0.1, w: 9, h: 0.6, fontSize: 28, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0 });

  const mailCards = [
    {
      title: "B-E区分 申込み直後",
      color: ACCENT,
      timing: "申込み送信直後（自動）",
      dest: "申込み者メールアドレス",
      attach: "なし",
      content: "申込み受付確認・内容確認用"
    },
    {
      title: "S-A区分 受付完了後",
      color: BLUE,
      timing: "担当者が受付完了を確認後",
      dest: "申込み者メールアドレス",
      attach: "PDF請求書",
      content: "受付完了通知・請求書添付"
    },
    {
      title: "全区分 入金完了後",
      color: NAVY,
      timing: "入金確認後（担当者操作）",
      dest: "申込み者メールアドレス",
      attach: "PDF感謝状",
      content: "入金御礼・感謝状添付"
    },
  ];

  mailCards.forEach((card, i) => {
    const cy = 1.0 + i * 1.55;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: cy, w: 9.4, h: 1.4, fill: { color: WHITE }, line: { color: card.color, width: 1.5 }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: cy, w: 2.5, h: 1.4, fill: { color: card.color }, line: { color: card.color, width: 0 } });
    s.addText(card.title, { x: 0.3, y: cy, w: 2.5, h: 1.4, fontSize: 13, bold: true, color: WHITE, fontFace: "Meiryo", align: "center", valign: "middle", margin: 0 });
    s.addText([
      { text: "タイミング: ", options: { bold: true } }, { text: card.timing, options: { breakLine: true } },
      { text: "送信先: ", options: { bold: true } }, { text: card.dest, options: { breakLine: true } },
      { text: "添付ファイル: ", options: { bold: true } }, { text: card.attach, options: { breakLine: true } },
      { text: "内容: ", options: { bold: true } }, { text: card.content }
    ], { x: 2.95, y: cy + 0.1, w: 6.6, h: 1.2, fontSize: 12, color: NAVY, fontFace: "Meiryo" });
  });
}

// ─────────────────────────────────────────
// SLIDE 14 — 非機能要件
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addText("非機能要件", { x: 0.4, y: 0.1, w: 9, h: 0.6, fontSize: 28, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0 });

  const nfr = [
    ["可用性", "Google インフラによる 99.9%+ の稼働率保証"],
    ["スケーラビリティ", "GAS の自動スケーリング（同時実行数の制限あり）"],
    ["セキュリティ", "HTTPS 必須、Google OAuth によるスクリプト保護"],
    ["メール送信上限", "Gmail 1日100通（GAS 無料枠）/ G Suite は500通"],
    ["並行処理", "GAS トリガー同時実行制限: 30並行まで"],
    ["タイムゾーン", "Asia/Tokyo (JST) — GAS プロジェクト設定で固定"],
    ["バックアップ", "Google Sheets 版履歴による自動バックアップ"],
  ];

  const headers14 = [
    { text: "項目", options: { fill: { color: NAVY }, color: WHITE, bold: true } },
    { text: "要件", options: { fill: { color: NAVY }, color: WHITE, bold: true } },
  ];

  const rows14 = [headers14, ...nfr.map((r, i) => [
    { text: r[0], options: { fill: { color: i % 2 === 0 ? "eef4fb" : WHITE }, color: BLUE, bold: true } },
    { text: r[1], options: { fill: { color: i % 2 === 0 ? "eef4fb" : WHITE }, color: NAVY } },
  ])];

  s.addTable(rows14, {
    x: 0.3, y: 1.0, w: 9.4, colW: [2.4, 7.0],
    border: { pt: 0.5, color: "cccccc" },
    fontFace: "Meiryo", fontSize: 13, rowH: 0.55
  });

  s.addShape(pres.shapes.OVAL, { x: 8.5, y: 4.5, w: 1.8, h: 1.8, fill: { color: LBLUE, transparency: 30 }, line: { color: LBLUE, width: 0 } });
}

// ─────────────────────────────────────────
// SLIDE 15 — まとめ
// ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: NAVY };

  // Decorative shapes
  s.addShape(pres.shapes.OVAL, { x: -0.5, y: 3.5, w: 3, h: 3, fill: { color: BLUE, transparency: 70 }, line: { color: BLUE, width: 0 } });
  s.addShape(pres.shapes.OVAL, { x: 8.0, y: -0.5, w: 2.5, h: 2.5, fill: { color: LBLUE, transparency: 70 }, line: { color: LBLUE, width: 0 } });

  s.addText("まとめ", { x: 0.5, y: 0.25, w: 9, h: 0.8, fontSize: 38, bold: true, color: WHITE, fontFace: "Meiryo", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.98, w: 9, h: 0.03, fill: { color: LBLUE }, line: { color: LBLUE, width: 0 } });

  const strengths = [
    { icon: "コスト", text: "Google インフラによるゼロ運用コスト・サーバーレス構成" },
    { icon: "統合", text: "Sheets / Drive / Gmail の完全統合で一元データ管理" },
    { icon: "自動化", text: "PDF生成・メール送信の完全自動化でスタッフ負担を軽減" },
    { icon: "拡張性", text: "Infoシート設定で区分・期間を柔軟に変更可能" },
  ];

  const cardW = 4.3;
  const cardH = 1.5;

  strengths.forEach((item, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const cx = 0.5 + col * (cardW + 0.55);
    const cy = 1.2 + row * (cardH + 0.25);

    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: cardW, h: cardH, fill: { color: BLUE, transparency: 20 }, line: { color: LBLUE, width: 1 } });
    s.addShape(pres.shapes.OVAL, { x: cx + 0.15, y: cy + 0.15, w: 0.7, h: 0.7, fill: { color: LBLUE }, line: { color: LBLUE, width: 0 } });
    s.addText(item.icon, { x: cx + 0.15, y: cy + 0.15, w: 0.7, h: 0.7, fontSize: 10, bold: true, color: NAVY, fontFace: "Meiryo", align: "center", valign: "middle", margin: 0 });
    s.addText(item.text, { x: cx + 1.0, y: cy + 0.2, w: cardW - 1.15, h: cardH - 0.4, fontSize: 13, color: WHITE, fontFace: "Meiryo" });
  });

  s.addText("WebHanabi — 川口花火大会 協賛管理システム | 2026-06-18", {
    x: 0.5, y: 5.2, w: 9, h: 0.35, fontSize: 11, color: "aaaaaa", fontFace: "Meiryo", align: "center", margin: 0
  });
}

pres.writeFile({ fileName: "/Users/minhsang1601/GitClone/WebHanabi/基本設計書_WebHanabi.pptx" })
  .then(() => console.log("Done: 基本設計書_WebHanabi.pptx"))
  .catch(e => { console.error(e); process.exit(1); });

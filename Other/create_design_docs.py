"""
Generate 基本設計書 and 詳細設計書 for WebHanabi project.
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from datetime import date

# Register Japanese fonts
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))

JP_SERIF = 'HeiseiMin-W3'
JP_SANS = 'HeiseiKakuGo-W5'

PAGE_W, PAGE_H = A4
MARGIN = 20 * mm


def make_styles():
    base = getSampleStyleSheet()
    styles = {}

    styles['cover_title'] = ParagraphStyle(
        'cover_title',
        fontName=JP_SANS,
        fontSize=28,
        leading=40,
        textColor=colors.HexColor('#1a3a5c'),
        spaceAfter=8,
    )
    styles['cover_subtitle'] = ParagraphStyle(
        'cover_subtitle',
        fontName=JP_SANS,
        fontSize=16,
        leading=24,
        textColor=colors.HexColor('#2e6da4'),
        spaceAfter=6,
    )
    styles['cover_meta'] = ParagraphStyle(
        'cover_meta',
        fontName=JP_SERIF,
        fontSize=11,
        leading=18,
        textColor=colors.HexColor('#444444'),
    )
    styles['h1'] = ParagraphStyle(
        'h1',
        fontName=JP_SANS,
        fontSize=16,
        leading=24,
        textColor=colors.HexColor('#1a3a5c'),
        spaceBefore=14,
        spaceAfter=6,
        borderPad=4,
    )
    styles['h2'] = ParagraphStyle(
        'h2',
        fontName=JP_SANS,
        fontSize=13,
        leading=20,
        textColor=colors.HexColor('#2e6da4'),
        spaceBefore=10,
        spaceAfter=4,
    )
    styles['h3'] = ParagraphStyle(
        'h3',
        fontName=JP_SANS,
        fontSize=11,
        leading=18,
        textColor=colors.HexColor('#1a5276'),
        spaceBefore=8,
        spaceAfter=3,
    )
    styles['body'] = ParagraphStyle(
        'body',
        fontName=JP_SERIF,
        fontSize=10,
        leading=17,
        textColor=colors.HexColor('#222222'),
        spaceAfter=4,
    )
    styles['bullet'] = ParagraphStyle(
        'bullet',
        fontName=JP_SERIF,
        fontSize=10,
        leading=17,
        leftIndent=14,
        textColor=colors.HexColor('#222222'),
        spaceAfter=2,
    )
    styles['code'] = ParagraphStyle(
        'code',
        fontName='Courier',
        fontSize=8,
        leading=12,
        backColor=colors.HexColor('#f5f5f5'),
        leftIndent=10,
        rightIndent=10,
        spaceAfter=4,
    )
    styles['table_header'] = ParagraphStyle(
        'table_header',
        fontName=JP_SANS,
        fontSize=9,
        leading=13,
        textColor=colors.white,
    )
    styles['table_cell'] = ParagraphStyle(
        'table_cell',
        fontName=JP_SERIF,
        fontSize=9,
        leading=13,
        textColor=colors.HexColor('#222222'),
    )
    return styles


HEADER_COLOR = colors.HexColor('#1a3a5c')
SUBHEADER_COLOR = colors.HexColor('#2e6da4')
LIGHT_BLUE = colors.HexColor('#dce9f5')
ALT_ROW = colors.HexColor('#f0f5fb')


def tbl_style(has_header=True):
    cmds = [
        ('FONTNAME', (0, 0), (-1, -1), JP_SERIF),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#aaaaaa')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, ALT_ROW]),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]
    if has_header:
        cmds += [
            ('BACKGROUND', (0, 0), (-1, 0), HEADER_COLOR),
            ('FONTNAME', (0, 0), (-1, 0), JP_SANS),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ]
    return TableStyle(cmds)


def add_page_number(canvas, doc):
    canvas.saveState()
    canvas.setFont(JP_SERIF, 8)
    canvas.setFillColor(colors.HexColor('#888888'))
    page_num = canvas.getPageNumber()
    canvas.drawRightString(PAGE_W - MARGIN, 10 * mm, f"- {page_num} -")
    canvas.drawString(MARGIN, 10 * mm, "WebHanabi システム")
    canvas.restoreState()


# ============================================================
# 基本設計書
# ============================================================

def build_kihon():
    s = make_styles()
    story = []

    # Cover
    story.append(Spacer(1, 40 * mm))
    story.append(Paragraph("WebHanabi システム", s['cover_subtitle']))
    story.append(Paragraph("基本設計書", s['cover_title']))
    story.append(HRFlowable(width="100%", thickness=2, color=HEADER_COLOR, spaceAfter=10))
    story.append(Spacer(1, 8 * mm))
    meta = [
        ["プロジェクト名", "川口花火大会 協賛管理システム (WebHanabi)"],
        ["ドキュメント種別", "基本設計書 (Basic Design Document)"],
        ["バージョン", "1.0"],
        ["作成日", str(date.today())],
        ["対象システム", "Google Apps Script + Google Sheets"],
    ]
    t = Table(meta, colWidths=[45 * mm, 110 * mm])
    t.setStyle(tbl_style(has_header=False))
    story.append(t)
    story.append(PageBreak())

    # --- 1. システム概要 ---
    story.append(Paragraph("1. システム概要", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=SUBHEADER_COLOR, spaceAfter=6))

    story.append(Paragraph("1.1 目的", s['h2']))
    story.append(Paragraph(
        "本システムは、川口花火大会における協賛申込み受付から、請求書発行・入金確認・お礼状送付に至る"
        "一連のワークフローを自動化することを目的とする。紙ベースの手作業を排除し、担当者の業務効率化"
        "とヒューマンエラーの防止を図る。",
        s['body']
    ))

    story.append(Paragraph("1.2 スコープ", s['h2']))
    scope = [
        ["対象", "説明"],
        ["協賛申込みフォーム", "一般公開用Webフォーム（Google Apps Script Webアプリ）"],
        ["協賛管理スプレッドシート", "申込みデータの一元管理（Google Sheets）"],
        ["自動メール送信", "申込み確認・請求書・お礼状のPDF添付メール（Gmail）"],
        ["PDF生成", "請求書・お礼状のPDF自動生成（Google Drive）"],
        ["管理者ワークフロー", "受付完了・入金確認チェックボックスによるステータス管理"],
    ]
    t = Table(scope, colWidths=[50 * mm, 105 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("1.3 前提条件・制約", s['h2']))
    for item in [
        "Google Workspace（または個人Googleアカウント）が必要",
        "Gmail の1日あたりの送信上限（50通/日）に依存する",
        "Google Apps Script の実行時間上限（6分/実行）が制約となる",
        "データはすべてGoogle Sheetsで管理し、外部DBは使用しない",
        "PDFは Google Drive 上の一時ファイルとして生成し、メール添付後に削除する",
    ]:
        story.append(Paragraph(f"・{item}", s['bullet']))

    # --- 2. システム構成 ---
    story.append(Paragraph("2. システム構成", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=SUBHEADER_COLOR, spaceAfter=6))

    story.append(Paragraph("2.1 アーキテクチャ概要", s['h2']))
    story.append(Paragraph(
        "本システムはサーバーレスアーキテクチャを採用し、Google のクラウドインフラ上で完結する。"
        "フロントエンドは Google Apps Script の HtmlService で提供し、バックエンドは GAS スクリプト"
        "が担う。データストアは Google Sheets、ファイルストレージは Google Drive を使用する。",
        s['body']
    ))

    story.append(Paragraph("2.2 コンポーネント構成", s['h2']))
    comps = [
        ["コンポーネント", "技術", "役割"],
        ["Webフォーム", "GAS HtmlService\n(index.html)", "協賛申込みフォームの表示・入力受付"],
        ["バックエンドAPI", "Google Apps Script\n(.gs ファイル群)", "フォーム処理・メール送信・PDF生成"],
        ["データストア", "Google Sheets", "申込みデータの永続化・ワークフロー管理"],
        ["ファイルストレージ", "Google Drive", "PDFファイルの一時保存・プロジェクトフォルダ管理"],
        ["メール配信", "Gmail (GAS MailApp)", "申込み確認・請求書・お礼状の送信"],
        ["設定管理", "Info シート\n+ Script Properties", "システム設定の一元管理とキャッシュ"],
    ]
    t = Table(comps, colWidths=[38 * mm, 42 * mm, 75 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("2.3 ファイル構成", s['h2']))
    files = [
        ["ファイル名", "種別", "説明"],
        ["WebApp.gs", "GAS", "エントリーポイント（doGet・フォーム送信処理）"],
        ["Config.gs", "GAS", "設定値取得・キャッシュ管理"],
        ["Sheet.gs", "GAS", "スプレッドシート操作（行追加・XLOOKUP設定）"],
        ["Mail.gs", "GAS", "メール送信・PDF生成"],
        ["Trigger.gs", "GAS", "onEdit トリガー・ダイアログコールバック"],
        ["ProjectInit.gs", "GAS", "プロジェクト初期化（フォルダ・シート作成）"],
        ["Setup.gs", "GAS", "Infoシートセットアップ・設定同期"],
        ["index.html", "HTML/JS", "協賛申込みフォームUI"],
        ["invoice-template.html", "HTML", "請求書PDFテンプレート"],
        ["oreijou-template.html", "HTML", "お礼状PDFテンプレート"],
        ["Stylesheet.html", "CSS", "フォームスタイルシート"],
        ["ConfirmInvoiceDialog.html", "HTML/JS", "請求書送信確認ダイアログ"],
        ["ConfirmNyukinDialog.html", "HTML/JS", "入金確認ダイアログ"],
    ]
    t = Table(files, colWidths=[55 * mm, 18 * mm, 82 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    # --- 3. 機能設計 ---
    story.append(Paragraph("3. 機能設計", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=SUBHEADER_COLOR, spaceAfter=6))

    story.append(Paragraph("3.1 機能一覧", s['h2']))
    funcs = [
        ["機能ID", "機能名", "概要", "優先度"],
        ["F-01", "協賛申込みフォーム表示", "一般公開Webフォームの表示・動的価格表示", "高"],
        ["F-02", "申込みデータ登録", "フォーム送信データをスプレッドシートに保存", "高"],
        ["F-03", "受付番号生成", "一意の受付番号を自動生成して返却", "高"],
        ["F-04", "B〜E区分 自動メール送信", "申込み直後に請求書PDF添付メールを自動送信", "高"],
        ["F-05", "S・A区分 手動メール送信", "抽選後に担当者が確認してから請求書メールを送信", "高"],
        ["F-06", "入金完了後 お礼状送信", "入金確認チェック後にお礼状PDF添付メールを送信", "高"],
        ["F-07", "請求書PDF生成", "会社情報・金額・インボイス情報を含むPDF生成", "高"],
        ["F-08", "お礼状PDF生成", "正式なお礼状PDFの自動生成", "高"],
        ["F-09", "メールクォータ監視", "Gmail送信上限を監視し、超過時に通知・停止", "中"],
        ["F-10", "プロジェクト初期化", "フォルダ・スプレッドシートの自動作成と設定", "中"],
        ["F-11", "設定管理", "Infoシートから設定値を読み込みキャッシュする", "中"],
    ]
    t = Table(funcs, colWidths=[16 * mm, 48 * mm, 82 * mm, 14 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("3.2 協賛区分と価格", s['h2']))
    story.append(Paragraph(
        "協賛区分はS・A・B・C・D・Eの6段階。価格はInfoシートから動的に取得する。"
        "S・A区分は抽選制で申込み期間が異なる。",
        s['body']
    ))
    kubun = [
        ["区分", "申込み方式", "申込み期間設定"],
        ["S", "抽選制（S/A共通期間）", "KUBUN_SA_START / KUBUN_SA_END"],
        ["A", "抽選制（S/A共通期間）", "KUBUN_SA_START / KUBUN_SA_END"],
        ["B", "先着制（B〜E共通期間）", "KUBUN_BCDE_START / KUBUN_BCDE_END"],
        ["C", "先着制（B〜E共通期間）", "KUBUN_BCDE_START / KUBUN_BCDE_END"],
        ["D", "先着制（B〜E共通期間）", "KUBUN_BCDE_START / KUBUN_BCDE_END"],
        ["E", "先着制（B〜E共通期間）", "KUBUN_BCDE_START / KUBUN_BCDE_END"],
    ]
    t = Table(kubun, colWidths=[15 * mm, 55 * mm, 85 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    # --- 4. データ設計 ---
    story.append(Paragraph("4. データ設計", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=SUBHEADER_COLOR, spaceAfter=6))

    story.append(Paragraph("4.1 スプレッドシート構成", s['h2']))
    sheets = [
        ["シート名", "役割", "主な更新者"],
        ["Info", "システム設定の一元管理（Key-Value形式）", "管理者（手動）"],
        ["協賛申込み一覧", "申込みデータのメインテーブル", "システム（自動）"],
        ["手作業", "ワークフロー管理（XLOOKUP + チェックボックス）", "担当者（手動）+ システム（自動）"],
        ["CreateLog", "プロジェクト初期化の操作ログ", "システム（自動）"],
    ]
    t = Table(sheets, colWidths=[48 * mm, 80 * mm, 27 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("4.2 協賛申込み一覧シート — カラム定義", s['h2']))
    cols_main = [
        ["列", "項目名", "データ型", "備考"],
        ["A", "受付番号", "文字列", "例: KWGC0618120530123"],
        ["B", "受付日時", "日時", "Asia/Tokyo タイムゾーン"],
        ["C", "会社名・団体名", "文字列", ""],
        ["D", "会社名（フリガナ）", "文字列", "全角カタカナ"],
        ["E", "代表者役職・氏名", "文字列", ""],
        ["F", "代表者（フリガナ）", "文字列", "全角カタカナ"],
        ["G", "担当者名", "文字列", ""],
        ["H", "担当者名（フリガナ）", "文字列", "全角カタカナ"],
        ["I", "郵便番号", "テキスト", "先頭ゼロ保持（@ 書式）"],
        ["J", "住所", "文字列", ""],
        ["K", "電話番号", "テキスト", "先頭ゼロ保持（@ 書式）"],
        ["L", "メールアドレス", "文字列", ""],
        ["M", "区分", "文字列", "S / A / B / C / D / E"],
        ["N", "会社HP URL", "文字列", "任意"],
        ["O", "掲載名（任意）", "文字列", "空の場合は会社名を使用"],
    ]
    t = Table(cols_main, colWidths=[10 * mm, 50 * mm, 25 * mm, 70 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("4.3 手作業シート — カラム定義", s['h2']))
    cols_te = [
        ["列", "項目名", "データ型", "備考"],
        ["A", "受付番号", "文字列", "担当者が手入力（トリガーキー）"],
        ["B", "区分", "文字列", "XLOOKUP 自動取得"],
        ["C", "電話番号", "テキスト", "XLOOKUP 自動取得"],
        ["D", "会社名・団体名", "文字列", "XLOOKUP 自動取得"],
        ["E", "住所", "文字列", "XLOOKUP 自動取得"],
        ["F", "代表者役職・氏名", "文字列", "XLOOKUP 自動取得"],
        ["G", "メールアドレス", "文字列", "XLOOKUP 自動取得"],
        ["H", "会社HP URL", "文字列", "XLOOKUP 自動取得"],
        ["I", "受付完了", "チェックボックス", "担当者がチェック → 請求書ダイアログ表示"],
        ["J", "請求書送信日時", "日時", "送信確定時に自動設定"],
        ["K", "入金完了", "チェックボックス", "担当者がチェック → お礼状ダイアログ表示"],
        ["L", "お礼状送信日時", "日時", "送信確定時に自動設定"],
    ]
    t = Table(cols_te, colWidths=[10 * mm, 44 * mm, 25 * mm, 76 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    # --- 5. 外部インターフェース ---
    story.append(Paragraph("5. 外部インターフェース設計", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=SUBHEADER_COLOR, spaceAfter=6))

    story.append(Paragraph("5.1 Webフォーム（申込者向け）", s['h2']))
    story.append(Paragraph(
        "Google Apps Script の doGet() で提供する公開Webアプリ。"
        "申込者は URL にアクセスしてフォームに入力・送信する。"
        "送信後、受付番号が画面に表示され、確認メールが届く。",
        s['body']
    ))
    iface_form = [
        ["項目", "内容"],
        ["URL形式", "https://script.google.com/macros/s/{SCRIPT_ID}/exec"],
        ["アクセス制御", "全員公開（認証不要）"],
        ["レスポンス形式", "HTML（HtmlService.createTemplateFromFile）"],
        ["通信プロトコル", "HTTPS（Google管理）"],
        ["フォームデータ形式", "JSON文字列（google.script.run 経由）"],
    ]
    t = Table(iface_form, colWidths=[45 * mm, 110 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("5.2 メールインターフェース", s['h2']))
    emails = [
        ["メール種別", "送信タイミング", "宛先", "添付"],
        ["申込み確認（B〜E）", "フォーム送信直後（自動）", "申込者 + CC: 事務局", "請求書PDF"],
        ["申込み確認（S・A）", "担当者が受付完了チェック後（手動トリガー）", "申込者 + CC: 事務局", "請求書PDF"],
        ["お礼状", "担当者が入金完了チェック後（手動トリガー）", "申込者 + CC: 事務局", "お礼状PDF"],
    ]
    t = Table(emails, colWidths=[44 * mm, 52 * mm, 42 * mm, 17 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    # --- 6. 非機能要件 ---
    story.append(Paragraph("6. 非機能要件", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=SUBHEADER_COLOR, spaceAfter=6))

    nfr = [
        ["項目", "要件"],
        ["可用性", "Google インフラに依存（SLA 99.9%）"],
        ["スケーラビリティ", "Google Sheets の行上限（500万セル）以内に収まる想定"],
        ["セキュリティ", "フォームデータのサーバーサイドバリデーション実施\nスクリプトエディタへのアクセスは管理者のみ"],
        ["メール送信上限", "50通/日（無料アカウント）\nMIN_MAIL_QUOTA 以下になると管理者通知・フォーム停止"],
        ["並行処理", "LockService による書き込みロックで二重登録を防止"],
        ["タイムゾーン", "全日時処理を Asia/Tokyo (JST) で統一"],
        ["バックアップ", "Google Sheets の変更履歴機能を利用"],
    ]
    t = Table(nfr, colWidths=[40 * mm, 115 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    return story


# ============================================================
# 詳細設計書
# ============================================================

def build_shosai():
    s = make_styles()
    story = []

    # Cover
    story.append(Spacer(1, 40 * mm))
    story.append(Paragraph("WebHanabi システム", s['cover_subtitle']))
    story.append(Paragraph("詳細設計書", s['cover_title']))
    story.append(HRFlowable(width="100%", thickness=2, color=HEADER_COLOR, spaceAfter=10))
    story.append(Spacer(1, 8 * mm))
    meta = [
        ["プロジェクト名", "川口花火大会 協賛管理システム (WebHanabi)"],
        ["ドキュメント種別", "詳細設計書 (Detailed Design Document)"],
        ["バージョン", "1.0"],
        ["作成日", str(date.today())],
        ["対象システム", "Google Apps Script + Google Sheets"],
    ]
    t = Table(meta, colWidths=[45 * mm, 110 * mm])
    t.setStyle(tbl_style(has_header=False))
    story.append(t)
    story.append(PageBreak())

    # --- 1. WebApp.gs ---
    story.append(Paragraph("1. WebApp.gs — エントリーポイント", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=SUBHEADER_COLOR, spaceAfter=6))

    story.append(Paragraph("1.1 doGet(e)", s['h2']))
    story.append(Paragraph("GASウェブアプリのエントリーポイント。HTTPGETリクエストを受け取りフォームHTMLを返す。", s['body']))
    doget = [
        ["処理ステップ", "内容"],
        ["1. 設定確認", "getInfoSheet() で Info シートを取得。存在しない場合はエラーHTMLを返す"],
        ["2. 必須項目検証", "EVENT_NAME, OFFICE_EMAIL 等の必須設定値が空でないかチェック"],
        ["3. クォータ確認", "_getMailQuotaSafe() でGmailクォータを確認。不足時はフォーム停止HTML返却"],
        ["4. 背景画像読み込み", "BG_IMAGE_ID が設定されていれば Drive から Base64 DataURL に変換"],
        ["5. テンプレート生成", "createTemplateFromFile('index') でテンプレートを生成し設定値を注入"],
        ["6. HTML返却", "evaluate().setXFrameOptionsMode(ALLOWALL) でHTMLを返却"],
    ]
    t = Table(doget, colWidths=[45 * mm, 110 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("1.2 submitFormJson(jsonStr)", s['h2']))
    story.append(Paragraph("フォーム送信を受け取るサーバーサイド関数。google.script.run から呼ばれる。", s['body']))
    submit = [
        ["処理ステップ", "内容"],
        ["1. JSONパース", "JSON.parse(jsonStr) でフォームデータをオブジェクトに変換"],
        ["2. サーバーサイドバリデーション", "郵便番号（7桁数字）・電話番号（10-11桁数字）の形式チェック"],
        ["3. データ登録", "appendRow(data) でメインシートに1行追加、受付番号を取得"],
        ["4. 手作業シート登録", "appendToTesagyouSheet(receiptNo) で手作業シートに行追加"],
        ["5. メール送信判定", "区分がB〜Eの場合: generateAndSendInvoice() を即時実行\n区分がS・Aの場合: メール送信なし（手動対応）"],
        ["6. クォータ監視通知", "残クォータが MIN_MAIL_QUOTA 未満なら管理者に警告メール送信"],
        ["7. 受付番号返却", "受付番号文字列をクライアントに返却"],
    ]
    t = Table(submit, colWidths=[52 * mm, 103 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("1.3 メールクォータ管理", s['h2']))
    quota_detail = [
        ["関数名", "処理内容"],
        ["_getMailQuotaSafe()", "MailApp.getRemainingDailyQuota() を try/catch で安全に呼び出す。エラー時は -1 を返す"],
        ["_notifyLowQuota(remaining)", "残クォータが閾値以下の場合、OFFICE_EMAIL に警告メールを送信する"],
    ]
    t = Table(quota_detail, colWidths=[60 * mm, 95 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    # --- 2. Config.gs ---
    story.append(Paragraph("2. Config.gs — 設定管理", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=SUBHEADER_COLOR, spaceAfter=6))

    story.append(Paragraph("2.1 設定取得の優先順位", s['h2']))
    story.append(Paragraph(
        "getConfigVal(key) は以下の順序で値を検索する。これによりスプレッドシートへのアクセスを最小化し"
        "パフォーマンスを向上させる。",
        s['body']
    ))
    for item in [
        "1. Script Properties（キャッシュ）から取得を試みる",
        "2. Script Properties に存在しない場合、Info シートから直接読み込む",
        "3. Info シートにも存在しない場合、デフォルト値（引数 defaultVal）を返す",
    ]:
        story.append(Paragraph(item, s['bullet']))

    story.append(Paragraph("2.2 主要な設定キー一覧", s['h2']))
    config_keys = [
        ["設定キー", "説明", "デフォルト値"],
        ["EVENT_NAME", "イベント名（フォームタイトル等に使用）", "（必須）"],
        ["OFFICE_EMAIL", "事務局メールアドレス（CC・通知先）", "（必須）"],
        ["ORG_NAME", "主催団体名", "（必須）"],
        ["ORG_REP", "代表者名", "（必須）"],
        ["START_DATE / END_DATE", "イベント開催日時", "（必須）"],
        ["PAYMENT_DUE", "入金期限日", "（必須）"],
        ["BANK_NAME / BANK_NO\nBANK_HOLDER / BANK_REP", "振込先銀行情報", "（必須）"],
        ["PRICE_S 〜 PRICE_E", "各区分の協賛金額", "（必須）"],
        ["KUBUN_SA_START/END", "S・A区分の申込み受付期間", "空（常時受付）"],
        ["KUBUN_BCDE_START/END", "B〜E区分の申込み受付期間", "空（常時受付）"],
        ["BG_IMAGE_ID", "背景画像の Google Drive ファイルID", "空（背景なし）"],
        ["MIN_MAIL_QUOTA", "メール送信停止閾値（残通数）", "5"],
        ["DATA_SPREADSHEET_ID", "データシートのスプレッドシートID", "（初期化後設定）"],
        ["INVOICE_REG_NO", "インボイス登録番号", "（任意）"],
    ]
    t = Table(config_keys, colWidths=[48 * mm, 72 * mm, 35 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    # --- 3. Sheet.gs ---
    story.append(Paragraph("3. Sheet.gs — スプレッドシート操作", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=SUBHEADER_COLOR, spaceAfter=6))

    story.append(Paragraph("3.1 appendRow(data) — メインシート行追加", s['h2']))
    append_row = [
        ["処理", "詳細"],
        ["ロック取得", "LockService.getScriptLock() で最大30秒待機。取得失敗時は例外をスロー"],
        ["受付番号生成", "generateReceiptNumber() : prefix + yyyyMMddHHmmssSSS 形式\n例: KWGC0618120530123"],
        ["タイムゾーン変換", "Utilities.formatDate(..., 'Asia/Tokyo', ...) で JST 日時文字列を生成"],
        ["書式設定", "郵便番号・電話番号列に '@' 書式を設定し先頭ゼロを保持"],
        ["行追加", "sheet.appendRow(rowData) でシート末尾に1行追加"],
        ["ロック解放", "finally ブロックで必ず lock.releaseLock() を実行"],
    ]
    t = Table(append_row, colWidths=[38 * mm, 117 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("3.2 appendToTesagyouSheet(receiptNo) — 手作業シート行追加", s['h2']))
    story.append(Paragraph(
        "受付番号をキーとして手作業シートに行を追加する。B〜H列はXLOOKUP数式で"
        "メインシートからデータを自動取得する。",
        s['body']
    ))
    xlookup_note = [
        ["列", "設定内容"],
        ["A列", "受付番号（引数 receiptNo を直接セット）"],
        ["B〜H列", "=XLOOKUP(A{row}, メインシート!A:A, メインシート!{col}:{col}) 数式を挿入"],
        ["I列", "チェックボックス（受付完了）を挿入"],
        ["J列", "B〜E区分は請求書送信日時を即時セット。S・A区分は空白"],
        ["K列", "チェックボックス（入金完了）を挿入"],
        ["L列", "お礼状送信日時（空白で初期化）"],
    ]
    t = Table(xlookup_note, colWidths=[20 * mm, 135 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    # --- 4. Mail.gs ---
    story.append(Paragraph("4. Mail.gs — メール送信・PDF生成", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=SUBHEADER_COLOR, spaceAfter=6))

    story.append(Paragraph("4.1 generateAndSendInvoice(rowData)", s['h2']))
    story.append(Paragraph("請求書PDFを生成し、申込者にメール送信する。", s['body']))
    invoice_steps = [
        ["ステップ", "処理"],
        ["1. テンプレート読込", "HtmlService.createTemplateFromFile('invoice-template') で HTML テンプレートを取得"],
        ["2. プレースホルダー置換", "{{company}}, {{amount}}, {{tax}} 等を実際の値に置換"],
        ["3. HTML→ファイル変換", "Drive に一時 HTML ファイルを作成"],
        ["4. PDF変換", "DriveApp の exportLinks['application/pdf'] で PDF に変換・保存"],
        ["5. メール送信", "MailApp.sendEmail() で申込者宛に送信（replyTo, CC: 事務局）"],
        ["6. 一時ファイル削除", "Drive から一時 HTML ファイルを削除してクリーンアップ"],
    ]
    t = Table(invoice_steps, colWidths=[38 * mm, 117 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("4.2 請求書PDFの含有情報", s['h2']))
    invoice_contents = [
        ["項目", "データソース"],
        ["会社名・担当者名", "申込みデータ（スプレッドシート）"],
        ["協賛区分・金額（税抜）", "申込みデータ + Config"],
        ["消費税（10%）・合計金額", "金額から自動計算"],
        ["インボイス登録番号", "Config: INVOICE_REG_NO"],
        ["振込先銀行情報", "Config: BANK_NAME / BANK_NO / BANK_HOLDER"],
        ["入金期限", "Config: PAYMENT_DUE"],
        ["主催団体名・代表者名", "Config: ORG_NAME / ORG_REP"],
        ["組織印影（印鑑画像）", "Script Properties に Base64 PNG として保存"],
    ]
    t = Table(invoice_contents, colWidths=[70 * mm, 85 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("4.3 generateAndSendOreijou(rowData)", s['h2']))
    story.append(Paragraph(
        "入金確認後にお礼状PDFを生成・送信する。処理フローは請求書と同様（テンプレートは"
        " oreijou-template.html）。",
        s['body']
    ))

    story.append(Paragraph("4.4 メールテンプレート変数", s['h2']))
    mail_vars = [
        ["変数", "説明"],
        ["{{company}}", "会社名・団体名"],
        ["{{rep}}", "代表者役職・氏名"],
        ["{{kubun}}", "協賛区分（S/A/B/C/D/E）"],
        ["{{amount}}", "協賛金額（税込）"],
        ["{{receiptNo}}", "受付番号"],
        ["{{eventName}}", "イベント名（Config: EVENT_NAME）"],
        ["{{paymentDue}}", "入金期限（Config: PAYMENT_DUE）"],
        ["{{orgName}}", "主催団体名（Config: ORG_NAME）"],
    ]
    t = Table(mail_vars, colWidths=[40 * mm, 115 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    # --- 5. Trigger.gs ---
    story.append(Paragraph("5. Trigger.gs — イベントトリガー", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=SUBHEADER_COLOR, spaceAfter=6))

    story.append(Paragraph("5.1 onEditInstallable(e) — 編集トリガー", s['h2']))
    story.append(Paragraph(
        "手作業シートの編集を検知してダイアログを表示するインストール型トリガー。"
        "ProjectInit.gs の初期化時に自動登録される。",
        s['body']
    ))
    edit_trigger = [
        ["条件", "処理"],
        ["シートが「手作業」以外", "即時 return（対象外シート）"],
        ["I列（受付完了）をチェック ON", "ConfirmInvoiceDialog を表示\n（既にJ列に日時がある場合は「送信済み」アラートを表示して I をリセット）"],
        ["I列（受付完了）をチェック OFF", "確認ダイアログ表示。J・K・L列をクリアするか確認"],
        ["K列（入金完了）をチェック ON", "J列（請求書送信日時）が空の場合は警告して K をリセット\n（正常時）ConfirmNyukinDialog を表示"],
        ["K列（入金完了）をチェック OFF", "確認ダイアログ表示。L列をクリアするか確認"],
    ]
    t = Table(edit_trigger, colWidths=[62 * mm, 93 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("5.2 ダイアログコールバック関数", s['h2']))
    callbacks = [
        ["関数名", "呼び出し元", "処理"],
        ["sendInvoiceConfirmed(row)", "ConfirmInvoiceDialog", "請求書PDF生成・送信 → J列に現在日時を記録"],
        ["cancelInvoiceSend(row)", "ConfirmInvoiceDialog", "I列のチェックを FALSE に戻す"],
        ["sendNyukinConfirmed(row)", "ConfirmNyukinDialog", "お礼状PDF生成・送信 → L列に現在日時を記録"],
        ["cancelNyukin(row)", "ConfirmNyukinDialog", "K列のチェックを FALSE に戻す"],
    ]
    t = Table(callbacks, colWidths=[52 * mm, 42 * mm, 61 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("5.3 _blockSendIfLowQuota()", s['h2']))
    story.append(Paragraph(
        "メール送信前にクォータを確認するガード関数。残量が MIN_MAIL_QUOTA 未満の場合は"
        "例外をスローし、呼び出し元の送信処理を中断させる。",
        s['body']
    ))

    # --- 6. ProjectInit.gs ---
    story.append(Paragraph("6. ProjectInit.gs — プロジェクト初期化", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=SUBHEADER_COLOR, spaceAfter=6))

    story.append(Paragraph("6.1 initProject() — 初期化フロー", s['h2']))
    init_steps = [
        ["順序", "処理", "詳細"],
        ["1", "ルートフォルダ取得", "Info シートの ROOT_FOLDER_ID から Drive フォルダを取得"],
        ["2", "プロジェクトフォルダ作成", "YYYY_EventName 形式のサブフォルダを作成"],
        ["3", "スプレッドシート作成", "プロジェクトフォルダ内に新規スプレッドシートを作成"],
        ["4", "シート構成設定", "協賛申込み一覧・手作業・CreateLog の3シートを作成・整形"],
        ["5", "ヘッダー設定", "各シートにカラム名・書式・フィルターを設定"],
        ["6", "トリガー登録", "onEditInstallable と onOpenEventSheet を ScriptApp に登録"],
        ["7", "スプレッドシートID保存", "作成したIDを Info シートの DATA_SPREADSHEET_ID に書き込む"],
        ["8", "ログ記録", "CreateLog シートに操作内容と日時を記録"],
    ]
    t = Table(init_steps, colWidths=[12 * mm, 50 * mm, 93 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    # --- 7. index.html ---
    story.append(Paragraph("7. index.html — Webフォーム詳細", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=SUBHEADER_COLOR, spaceAfter=6))

    story.append(Paragraph("7.1 フォームの画面遷移", s['h2']))
    screens = [
        ["画面状態", "表示条件", "ユーザー操作"],
        ["入力フォーム", "初期表示", "各フィールドに入力後「確認」ボタンをクリック"],
        ["確認画面", "バリデーション通過後", "内容を確認して「送信」または「戻る」"],
        ["完了画面", "サーバー送信成功", "受付番号が表示される"],
        ["エラー画面", "設定不備・クォータ不足", "フォームは表示されずエラーメッセージのみ"],
    ]
    t = Table(screens, colWidths=[38 * mm, 48 * mm, 69 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("7.2 クライアントサイドバリデーション", s['h2']))
    validations = [
        ["フィールド", "バリデーションルール"],
        ["フリガナ（全4箇所）", "全角カタカナのみ許可（正規表現: /^[ァ-ヶー　]+$/）"],
        ["郵便番号", "7桁の数字のみ（ハイフンなし）"],
        ["電話番号", "10〜11桁の数字のみ（ハイフンなし）"],
        ["メールアドレス", "標準メール形式（HTML5 type=email）"],
        ["利用規約同意", "規約テキストを最下部までスクロールしないとチェック不可"],
        ["区分選択", "現在日時が申込み受付期間内の区分のみ表示"],
    ]
    t = Table(validations, colWidths=[50 * mm, 105 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    story.append(Paragraph("7.3 受付番号生成ロジック", s['h2']))
    story.append(Paragraph(
        "受付番号は以下の形式で生成される。Prefix は Config の RECEIPT_PREFIX（例: KWGC）から取得。"
        "タイムスタンプ部分はサーバーサイドで JST 時刻を使用する。",
        s['body']
    ))
    story.append(Paragraph("形式: {RECEIPT_PREFIX}{yyyyMMddHHmmssSSS}", s['code']))
    story.append(Paragraph("例:  KWGC0618120530123", s['code']))

    # --- 8. エラーハンドリング ---
    story.append(Paragraph("8. エラーハンドリング方針", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=SUBHEADER_COLOR, spaceAfter=6))

    errors = [
        ["エラー種別", "発生箇所", "対応方針"],
        ["設定値が空", "doGet()", "エラーHTMLを表示し、フォームは表示しない"],
        ["クォータ不足", "doGet() / submitFormJson()", "フォーム停止HTML返却 or 例外スロー"],
        ["バリデーション失敗", "submitFormJson()", "エラーメッセージをクライアントに返却"],
        ["スプレッドシートロック取得失敗", "appendRow()", "例外スロー → クライアントにエラー表示"],
        ["メール送信失敗", "Mail.gs 各関数", "try/catch で捕捉 → ログ出力 + 管理者通知"],
        ["PDF生成失敗", "generateInvoicePdf()", "try/catch で捕捉 → 一時ファイルを削除してリスロー"],
        ["トリガー重複送信", "onEditInstallable()", "J / L 列の日時チェックで二重送信を防止"],
    ]
    t = Table(errors, colWidths=[42 * mm, 40 * mm, 73 * mm])
    t.setStyle(tbl_style())
    story.append(t)

    return story


def generate_pdf(story, out_path, title):
    doc = SimpleDocTemplate(
        out_path,
        pagesize=A4,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=MARGIN,
        bottomMargin=20 * mm,
        title=title,
        author="WebHanabi Project",
    )
    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    print(f"Generated: {out_path}")


if __name__ == '__main__':
    import os
    base = '/Users/minhsang1601/GitClone/WebHanabi'

    generate_pdf(
        build_kihon(),
        os.path.join(base, '基本設計書_WebHanabi.pdf'),
        '基本設計書 — WebHanabi 川口花火大会 協賛管理システム',
    )
    generate_pdf(
        build_shosai(),
        os.path.join(base, '詳細設計書_WebHanabi.pdf'),
        '詳細設計書 — WebHanabi 川口花火大会 協賛管理システム',
    )

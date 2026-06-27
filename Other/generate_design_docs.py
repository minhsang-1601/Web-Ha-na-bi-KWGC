#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WebHanabi 協賛管理システム - 基本設計書・詳細設計書 PDF生成スクリプト
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
import datetime

# Japanese font registration
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))

FONT_BOLD  = 'HeiseiKakuGo-W5'
FONT_NORM  = 'HeiseiMin-W3'
FONT_SANS  = 'HeiseiKakuGo-W5'

# Color palette
COL_PRIMARY   = colors.HexColor('#4a3b68')
COL_SECONDARY = colors.HexColor('#826ea0')
COL_ACCENT    = colors.HexColor('#007BFF')
COL_LIGHT     = colors.HexColor('#f0eaf8')
COL_HEADER_BG = colors.HexColor('#d0e4f7')
COL_GRAY      = colors.HexColor('#555555')
COL_LIGHT_GRAY= colors.HexColor('#f8f8f8')

def styles():
    base = getSampleStyleSheet()
    def s(name, **kw):
        return ParagraphStyle(name, **kw)

    return {
        'cover_title': s('cover_title',
            fontName=FONT_BOLD, fontSize=26, textColor=COL_PRIMARY,
            leading=36, spaceAfter=10, alignment=1),
        'cover_sub': s('cover_sub',
            fontName=FONT_BOLD, fontSize=16, textColor=COL_SECONDARY,
            leading=24, spaceAfter=6, alignment=1),
        'cover_info': s('cover_info',
            fontName=FONT_NORM, fontSize=11, textColor=COL_GRAY,
            leading=18, spaceAfter=4, alignment=1),
        'h1': s('h1',
            fontName=FONT_BOLD, fontSize=16, textColor=colors.white,
            leading=22, spaceBefore=6, spaceAfter=6,
            backColor=COL_PRIMARY, leftIndent=-10, rightIndent=-10,
            borderPad=6),
        'h2': s('h2',
            fontName=FONT_BOLD, fontSize=13, textColor=COL_PRIMARY,
            leading=18, spaceBefore=14, spaceAfter=4,
            borderColor=COL_SECONDARY, borderWidth=0,
            leftIndent=0),
        'h3': s('h3',
            fontName=FONT_BOLD, fontSize=11, textColor=COL_SECONDARY,
            leading=16, spaceBefore=10, spaceAfter=3),
        'body': s('body',
            fontName=FONT_NORM, fontSize=10, textColor=colors.black,
            leading=16, spaceBefore=2, spaceAfter=2),
        'body_bold': s('body_bold',
            fontName=FONT_BOLD, fontSize=10, textColor=colors.black,
            leading=16, spaceBefore=2, spaceAfter=2),
        'bullet': s('bullet',
            fontName=FONT_NORM, fontSize=10, textColor=colors.black,
            leading=15, leftIndent=14, spaceBefore=1, spaceAfter=1,
            bulletIndent=4, bulletFontName=FONT_NORM),
        'code': s('code',
            fontName=FONT_SANS, fontSize=9, textColor=colors.HexColor('#1a1a2e'),
            leading=14, leftIndent=10, spaceBefore=2, spaceAfter=2,
            backColor=colors.HexColor('#f4f0fb')),
        'caption': s('caption',
            fontName=FONT_BOLD, fontSize=9, textColor=COL_GRAY,
            leading=13, spaceBefore=2, spaceAfter=6, alignment=1),
        'toc': s('toc',
            fontName=FONT_NORM, fontSize=11, textColor=COL_PRIMARY,
            leading=20, leftIndent=0),
        'toc_sub': s('toc_sub',
            fontName=FONT_NORM, fontSize=10, textColor=COL_GRAY,
            leading=18, leftIndent=16),
    }

ST = styles()

def h1(text):
    return [
        Spacer(1, 4*mm),
        Table([[Paragraph(text, ST['h1'])]],
              colWidths=[170*mm],
              style=TableStyle([
                  ('BACKGROUND', (0,0), (-1,-1), COL_PRIMARY),
                  ('LEFTPADDING',  (0,0), (-1,-1), 8),
                  ('RIGHTPADDING', (0,0), (-1,-1), 8),
                  ('TOPPADDING',   (0,0), (-1,-1), 5),
                  ('BOTTOMPADDING',(0,0), (-1,-1), 5),
              ])),
        Spacer(1, 3*mm),
    ]

def h2(text):
    return [
        Spacer(1, 3*mm),
        HRFlowable(width='100%', thickness=2, color=COL_SECONDARY),
        Paragraph(text, ST['h2']),
        HRFlowable(width='100%', thickness=0.5, color=COL_LIGHT),
        Spacer(1, 1*mm),
    ]

def h3(text):
    return [Paragraph(text, ST['h3'])]

def body(text):
    return Paragraph(text, ST['body'])

def bullet(text):
    return Paragraph(f'・ {text}', ST['bullet'])

def sp(h=3):
    return Spacer(1, h*mm)

def simple_table(headers, rows, col_widths=None):
    data = [headers] + rows
    if col_widths is None:
        w = 170*mm / len(headers)
        col_widths = [w] * len(headers)

    style = TableStyle([
        ('BACKGROUND',    (0,0), (-1,0),  COL_PRIMARY),
        ('TEXTCOLOR',     (0,0), (-1,0),  colors.white),
        ('FONTNAME',      (0,0), (-1,0),  FONT_BOLD),
        ('FONTSIZE',      (0,0), (-1,0),  9),
        ('FONTNAME',      (0,1), (-1,-1), FONT_NORM),
        ('FONTSIZE',      (0,1), (-1,-1), 9),
        ('BACKGROUND',    (0,1), (-1,-1), colors.white),
        ('ROWBACKGROUNDS',(0,1), (-1,-1), [colors.white, COL_LIGHT]),
        ('GRID',          (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN',        (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING',   (0,0), (-1,-1), 6),
        ('RIGHTPADDING',  (0,0), (-1,-1), 6),
        ('TOPPADDING',    (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('WORDWRAP',      (0,0), (-1,-1), True),
    ])

    # Convert all cells to Paragraphs
    def to_p(v, is_header=False):
        st = ST['body_bold'] if is_header else ST['body']
        return Paragraph(str(v), st)

    p_data = []
    for i, row in enumerate(data):
        p_data.append([to_p(cell, i==0) for cell in row])

    return Table(p_data, colWidths=col_widths, style=style, repeatRows=1)


# ═══════════════════════════════════════════════════════════════════════════
#  基本設計書
# ═══════════════════════════════════════════════════════════════════════════
def build_kihon():
    story = []
    today = datetime.date.today().strftime('%Y年%m月%d日')

    # ── 表紙 ──
    story.append(Spacer(1, 30*mm))
    story.append(Paragraph('川口花火大会', ST['cover_sub']))
    story.append(Paragraph('協賛管理システム', ST['cover_title']))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph('基本設計書', ParagraphStyle('ct2',
        fontName=FONT_BOLD, fontSize=20, textColor=COL_SECONDARY,
        leading=28, spaceAfter=6, alignment=1)))
    story.append(Spacer(1, 8*mm))
    story.append(HRFlowable(width='80%', thickness=2, color=COL_SECONDARY, hAlign='CENTER'))
    story.append(Spacer(1, 8*mm))
    info_rows = [
        ['バージョン', 'v1.0'],
        ['作成日',     today],
        ['対象システム', 'WebHanabi (Google Apps Script Web App)'],
        ['使用技術', 'Google Apps Script / Google Sheets / HTML・JavaScript'],
    ]
    it = Table(info_rows, colWidths=[50*mm, 110*mm],
               style=TableStyle([
                   ('FONTNAME',  (0,0),(-1,-1), FONT_NORM),
                   ('FONTSIZE',  (0,0),(-1,-1), 10),
                   ('FONTNAME',  (0,0),(0,-1),  FONT_BOLD),
                   ('TEXTCOLOR', (0,0),(0,-1),  COL_PRIMARY),
                   ('ALIGN',     (0,0),(0,-1),  'RIGHT'),
                   ('GRID',      (0,0),(-1,-1), 0.5, colors.HexColor('#dddddd')),
                   ('BACKGROUND',(0,0),(-1,-1), COL_LIGHT),
                   ('LEFTPADDING', (0,0),(-1,-1), 8),
                   ('RIGHTPADDING',(0,0),(-1,-1), 8),
                   ('TOPPADDING',  (0,0),(-1,-1), 5),
                   ('BOTTOMPADDING',(0,0),(-1,-1), 5),
               ]))
    story.append(it)
    story.append(PageBreak())

    # ── 目次 ──
    story += h1('目次')
    toc_items = [
        ('1. システム概要',                '3'),
        ('  1.1 目的',                     '3'),
        ('  1.2 対象ユーザー',             '3'),
        ('  1.3 システム構成',             '3'),
        ('2. 機能一覧',                    '4'),
        ('  2.1 Web フォーム（申込み受付）','4'),
        ('  2.2 スプレッドシート管理',     '4'),
        ('  2.3 メール・PDF自動送信',      '5'),
        ('  2.4 手作業シート管理',         '5'),
        ('3. 画面設計',                    '6'),
        ('  3.1 申込みフォーム画面',       '6'),
        ('  3.2 確認ダイアログ',           '6'),
        ('4. データ設計',                  '7'),
        ('  4.1 スプレッドシート構成',     '7'),
        ('  4.2 協賛申込み一覧シート',     '7'),
        ('  4.3 手作業シート',             '7'),
        ('  4.4 Info シート',              '8'),
        ('5. メール設計',                  '8'),
        ('  5.1 メール種別',               '8'),
        ('6. 非機能要件',                  '9'),
        ('  6.1 セキュリティ',             '9'),
        ('  6.2 可用性',                   '9'),
    ]
    for text, page in toc_items:
        indent = 16 if text.startswith('  ') else 0
        st = ParagraphStyle('toc_item', fontName=FONT_NORM, fontSize=10,
                            textColor=COL_PRIMARY if indent==0 else COL_GRAY,
                            leading=18, leftIndent=indent,
                            fontWeight='bold' if indent==0 else 'normal')
        if indent == 0:
            st.fontName = FONT_BOLD
        story.append(Paragraph(text, st))
    story.append(PageBreak())

    # ── 1. システム概要 ──
    story += h1('1. システム概要')

    story += h2('1.1 目的')
    story.append(body(
        '本システム（WebHanabi）は、川口花火大会の協賛企業・団体からの申込み受付、'
        '請求書発行、入金管理、お礼状送付を一元的に管理するためのWebアプリケーションである。'
        'Google Apps Script（GAS）上で動作し、Google スプレッドシートをデータベースとして使用する。'
    ))
    story.append(sp())

    story += h2('1.2 対象ユーザー')
    tbl = simple_table(
        ['ユーザー区分', '役割', '使用機能'],
        [
            ['協賛申込者（企業・団体）', 'Webフォームから協賛申込みを行う', 'フォーム画面・確認画面'],
            ['事務局スタッフ', 'スプレッドシート上で申込みを管理し、\n請求書・お礼状を送付する', '手作業シート・メール送信'],
            ['システム管理者', 'Info シートで設定値を管理し、\nGASトリガーを制御する', 'Info シート・GASエディタ'],
        ],
        col_widths=[40*mm, 65*mm, 65*mm]
    )
    story.append(tbl)
    story.append(sp())

    story += h2('1.3 システム構成')
    story.append(body('本システムは以下のコンポーネントで構成される。'))
    story.append(sp(2))
    arch_tbl = simple_table(
        ['コンポーネント', '技術', '役割'],
        [
            ['Webフォーム（フロントエンド）', 'HTML / CSS / JavaScript\n(GAS HtmlService)', '協賛申込みフォームの表示・バリデーション・送信'],
            ['サーバーサイドロジック', 'Google Apps Script (GAS)', 'フォームデータ受信、DB書込み、メール・PDF生成'],
            ['データベース', 'Google スプレッドシート', '申込みデータ・設定値の永続化'],
            ['ファイルストレージ', 'Google Drive', '印影画像・背景画像・PDFの一時保存'],
            ['メール送信', 'GAS MailApp', '申込確認・請求書・お礼状の自動送信'],
            ['設定ストア', 'Script Properties\n(GAS)', 'メールテンプレート・Info値のキャッシュ'],
        ],
        col_widths=[45*mm, 50*mm, 75*mm]
    )
    story.append(arch_tbl)
    story.append(PageBreak())

    # ── 2. 機能一覧 ──
    story += h1('2. 機能一覧')

    story += h2('2.1 Webフォーム（申込み受付）')
    story.append(simple_table(
        ['機能ID', '機能名', '説明'],
        [
            ['F-01', '誓約事項スクロールゲート', '利用規約を最後までスクロールしないと同意チェックが押せない'],
            ['F-02', '区分選択', 'S/A/B/C/D/E の6区分を申込み期間ごとに表示制御'],
            ['F-03', 'フォーム入力', '会社名・代表者・担当者・住所・電話番号・メール等の入力'],
            ['F-04', '入力バリデーション', '郵便番号7桁・電話番号9〜11桁・フリガナ（全角カタカナ）チェック'],
            ['F-05', '確認画面', '送信前に入力内容を一覧表示。「修正する」「送信する」ボタン'],
            ['F-06', '受付期間チェック', 'サーバーから取得した開始/終了日時で受付可否を判定'],
            ['F-07', 'メール残数チェック', 'GASのメール残数が不足の場合はフォームをメンテナンス表示'],
            ['F-08', '背景画像表示', 'Drive上の画像をBase64でフォーム背景に埋め込み'],
        ],
        col_widths=[20*mm, 52*mm, 98*mm]
    ))
    story.append(sp())

    story += h2('2.2 スプレッドシート管理')
    story.append(simple_table(
        ['機能ID', '機能名', '説明'],
        [
            ['F-09', '申込みデータ登録', 'フォーム送信時に「協賛申込み一覧」シートに行を追加'],
            ['F-10', '手作業シート登録', '申込みと同時に「手作業」シートにXLOOKUP数式行を追加'],
            ['F-11', 'ヘッダー自動作成', 'シートが空の場合、ヘッダー行を自動生成してフィルター設定'],
            ['F-12', 'カスタムメニュー', 'スプレッドシートUIに「協賛管理」メニューを追加'],
            ['F-13', 'Info設定管理', 'Info シートの値を Script Properties にキャッシュ同期'],
        ],
        col_widths=[20*mm, 52*mm, 98*mm]
    ))
    story.append(sp())

    story += h2('2.3 メール・PDF自動送信')
    story.append(simple_table(
        ['機能ID', '機能名', 'トリガー', '対象区分'],
        [
            ['F-14', '申込確認メール＋請求書PDF', 'フォーム送信時（自動）', 'B/C/D/E'],
            ['F-15', '申込受付メール（抽選待ち）', 'フォーム送信時（自動）', 'S/A'],
            ['F-16', '請求書メール＋PDF送付', '手作業シートI列チェック時', 'S/A（手動）'],
            ['F-17', 'お礼状メール＋PDF送付', '手作業シートK列チェック時', '全区分'],
        ],
        col_widths=[20*mm, 55*mm, 50*mm, 45*mm]
    ))
    story.append(sp())

    story += h2('2.4 手作業シート管理')
    story.append(simple_table(
        ['機能ID', '機能名', '説明'],
        [
            ['F-18', '請求書送信確認ダイアログ', 'I列チェック時にモーダルダイアログで確認後に請求書メール送信'],
            ['F-19', 'お礼状送信確認ダイアログ', 'K列チェック時にモーダルダイアログで確認後にお礼状メール送信'],
            ['F-20', 'チェック解除保護', '下流カラム（K→L）が設定済みの場合、解除前に確認アラートを表示'],
            ['F-21', '二重送信防止', '送信済みタイムスタンプ（J/L列）が存在する場合はチェックを無効化'],
            ['F-22', 'メール残数ブロック', '残数が最低ライン未満の場合は手動送信もブロック'],
        ],
        col_widths=[20*mm, 52*mm, 98*mm]
    ))
    story.append(PageBreak())

    # ── 3. 画面設計 ──
    story += h1('3. 画面設計')

    story += h2('3.1 申込みフォーム画面')
    story.append(body('フォームは以下の順序でセクションを表示する。受付期間外はフォームを非表示とする。'))
    story.append(sp(2))
    flow_data = [
        ['No', '表示要素', '条件'],
        ['1', '申込み期間表示', '常時表示'],
        ['2', '受付時間外メッセージ', '受付期間外の場合のみ'],
        ['3', 'メンテナンスメッセージ', 'メール残数不足の場合のみ'],
        ['4', '誓約事項テキストボックス', '受付期間中・残数OK'],
        ['5', '誓約事項チェックボックス', '誓約事項を最後までスクロール後に有効化'],
        ['6', '区分選択ラジオボタン', 'チェックボックス有効後に表示。区分ごとに受付期間を制御'],
        ['7', '入力フォーム各フィールド', '区分選択後に表示'],
        ['8', '「入力内容を確認する」ボタン', '同意チェック済みの場合のみ有効'],
        ['9', '確認画面', 'ボタン押下後に入力内容を一覧表示'],
        ['10', '「送信する」ボタン', '確認画面上で押下 → google.script.run.submitFormJson() 呼び出し'],
        ['11', '送信完了メッセージ + 受付番号', '送信成功後に表示'],
    ]
    story.append(simple_table(flow_data[0], flow_data[1:], col_widths=[12*mm, 80*mm, 78*mm]))
    story.append(sp())

    story += h2('3.2 確認ダイアログ（管理者向け）')
    story.append(simple_table(
        ['ダイアログ名', '表示タイミング', '操作'],
        [
            ['請求書送信確認ダイアログ\n(ConfirmInvoiceDialog.html)', '手作業シートI列をチェック時',
             '「送信する」→ sendInvoiceConfirmed()\n「キャンセル」→ cancelInvoiceSend()'],
            ['入金確認・お礼状送信ダイアログ\n(ConfirmNyukinDialog.html)', '手作業シートK列をチェック時',
             '「送信する」→ sendNyukinConfirmed()\n「キャンセル」→ cancelNyukin()'],
            ['お礼状確認ダイアログ\n(ConfirmOreijouDialog.html)', '（予備）', '（今後拡張用）'],
        ],
        col_widths=[60*mm, 50*mm, 60*mm]
    ))
    story.append(PageBreak())

    # ── 4. データ設計 ──
    story += h1('4. データ設計')

    story += h2('4.1 スプレッドシート構成')
    story.append(body(
        '本システムはMain スプレッドシート（GASが紐づくファイル）と、'
        'オプションで別のData スプレッドシート（INFO の DATA_SPREADSHEET_ID で指定）を使用する。'
    ))
    story.append(sp(2))
    story.append(simple_table(
        ['シート名', 'スプレッドシート', '概要'],
        [
            ['Info',           'Main',        '設定値を管理するマスターシート'],
            ['CreateLog',      'Main',        '（将来拡張用）'],
            ['協賛申込み一覧', 'Data（またはMain）', 'フォームからの申込みデータ（原本）'],
            ['手作業',         'Data（またはMain）', '事務局が操作する管理用シート（XLOOKUP参照）'],
        ],
        col_widths=[45*mm, 55*mm, 70*mm]
    ))
    story.append(sp())

    story += h2('4.2 協賛申込み一覧シート')
    story.append(simple_table(
        ['列', 'フィールド名', 'データ型', '説明'],
        [
            ['A', '受付番号',                'テキスト', 'プレフィックス+日時（例: KWGC0620123456789）'],
            ['B', '受付日時',                '日時',     'サーバー側で自動生成（yyyy/MM/dd HH:mm:ss）'],
            ['C', '会社名・団体名',           'テキスト', '申込者入力'],
            ['D', '会社名・団体名（フリガナ）','テキスト', '全角カタカナ'],
            ['E', '代表者役職・代表者名',     'テキスト', '役職と氏名を1フィールドで入力'],
            ['F', '代表者役職・代表者名（フリガナ）', 'テキスト', '全角カタカナ'],
            ['G', '担当者名',                'テキスト', ''],
            ['H', '担当者名（フリガナ）',    'テキスト', '全角カタカナ'],
            ['I', '郵便番号',                'テキスト（@）', '7桁数字。テキスト書式で先頭ゼロを保持'],
            ['J', '住所',                    'テキスト', ''],
            ['K', '電話番号',                'テキスト（@）', '9〜11桁数字。テキスト書式'],
            ['L', 'メールアドレス',           'テキスト', ''],
            ['M', '区分',                    'テキスト', 'S/A/B/C/D/E'],
            ['N', '会社HP URL',              'テキスト', '任意'],
            ['O', '掲載名（任意）',           'テキスト', '会社名と異なる場合のみ'],
        ],
        col_widths=[10*mm, 50*mm, 30*mm, 80*mm]
    ))
    story.append(sp())

    story += h2('4.3 手作業シート')
    story.append(body('1行目: ヘッダー（太字・背景#fce8b2）、2行目: サブヘッダー（種別メモ、文字サイズ8pt）、3行目以降: データ。フリーズは2行。'))
    story.append(sp(2))
    story.append(simple_table(
        ['列', 'フィールド名', '入力種別', '説明'],
        [
            ['A', '受付番号',        '直接入力',           '申込み時にサーバーが書き込み'],
            ['B', '区分',            'XLOOKUP（自動）',    '協賛申込み一覧M列を参照'],
            ['C', '電話番号',        'XLOOKUP（自動）',    '協賛申込み一覧K列を参照'],
            ['D', '会社名・団体名',   'XLOOKUP（自動）',    '協賛申込み一覧C列を参照'],
            ['E', '住所',            'XLOOKUP（自動）',    '協賛申込み一覧J列を参照'],
            ['F', '代表者役職・代表者名', 'XLOOKUP（自動）', '協賛申込み一覧E列を参照'],
            ['G', 'メールアドレス',   'XLOOKUP（自動）',    '協賛申込み一覧L列を参照'],
            ['H', '会社HP URL',      'XLOOKUP（自動）',    '協賛申込み一覧N列を参照'],
            ['I', '受付完了',        'チェックボックス（手動）', 'チェック → 請求書送信ダイアログ表示'],
            ['J', '請求書送信日時',   'タイムスタンプ（自動）', '送信確認後にサーバーが書き込み'],
            ['K', '入金完了',        'チェックボックス（手動）', 'チェック → お礼状送信ダイアログ表示（J必須）'],
            ['L', 'お礼状送信日時',   'タイムスタンプ（自動）', '送信確認後にサーバーが書き込み'],
        ],
        col_widths=[10*mm, 42*mm, 40*mm, 78*mm]
    ))
    story.append(PageBreak())

    story += h2('4.4 Info シート（設定マスター）')
    story.append(body('システム設定値をKey-Value形式で管理。Script Properties にキャッシュされ、トリガーコンテキストからも参照可能。'))
    story.append(sp(2))
    story.append(simple_table(
        ['キー', '説明', '例'],
        [
            ['EVENT_NAME',       'イベント名',          '第5回川口花火大会'],
            ['OFFICE_EMAIL',     '事務局メールアドレス', 'info@example.com'],
            ['START_DATE',       '受付開始日時（ISO8601）', '2025-01-01T00:00:00'],
            ['END_DATE',         '受付終了日時（ISO8601）', '2026-10-01T23:59:59'],
            ['KUBUN_SA_START/END','S・A区分 申込み期間',  '2026-06-01T00:00:00'],
            ['KUBUN_BCDE_START/END','B〜E区分 申込み期間', '2026-06-01T00:00:00'],
            ['PAYMENT_DUE',      '支払い期限文字列',    '9月18日（金）'],
            ['PRICE_S〜PRICE_E', '各区分協賛金額（円）', '2000000'],
            ['HANKO_FILE_ID',    '印影画像 DriveファイルID', '1aBc...'],
            ['BG_IMAGE_ID',      '背景画像 DriveファイルID', '1xYz...'],
            ['ORG_NAME',         '発行者組織名',        '川口花火大会実行委員会'],
            ['ORG_REP',          '代表者名（請求書用）', '委員長 廣瀬 進治'],
            ['INVOICE_REG_NO',   'インボイス登録番号',  'T9700150122003'],
            ['BANK_NAME',        '振込先銀行名',        '埼玉りそな銀行 川口支店'],
            ['BANK_NO',          '口座番号',            '6216349'],
            ['BANK_HOLDER',      '口座名義（略称）',    '川口商工会議所'],
            ['DATA_SPREADSHEET_ID', 'データシートID（別SS）', '1abc...（空ならMainを使用）'],
            ['MIN_MAIL_QUOTA',   'メール残数最低ライン', '5'],
            ['RECEIPT_NO_PREFIX','受付番号プレフィックス', 'KWGC'],
        ],
        col_widths=[60*mm, 65*mm, 45*mm]
    ))
    story.append(PageBreak())

    # ── 5. メール設計 ──
    story += h1('5. メール設計')

    story += h2('5.1 メール種別')
    story.append(simple_table(
        ['種別', '送信タイミング', '対象区分', '添付', 'テンプレートキー'],
        [
            ['申込確認メール\n＋請求書PDF',   'フォーム送信直後（自動）', 'B/C/D/E',
             '請求書PDF\n（invoice-template.html）', 'MAIL_SUBJECT\nMAIL_BODY'],
            ['申込受付メール\n（抽選待ち）',  'フォーム送信直後（自動）', 'S/A',
             'なし', 'RECEIPT_ONLY_SUBJECT\nRECEIPT_ONLY_BODY'],
            ['抽選確定・請求書メール',         '手作業I列チェック確認後（手動）', 'S/A',
             '請求書PDF', 'SA_INVOICE_SUBJECT\nSA_INVOICE_BODY'],
            ['お礼状メール＋PDF',             '手作業K列チェック確認後（手動）', '全区分',
             'お礼状PDF\n（oreijou-template.html）', 'OREIJOU_SUBJECT\nOREIJOU_BODY'],
        ],
        col_widths=[35*mm, 40*mm, 28*mm, 35*mm, 42*mm]
    ))
    story.append(sp(3))
    story += h3('メールテンプレート変数（共通）')
    story.append(simple_table(
        ['変数', '内容'],
        [
            ['{{company_name}}', '会社名・団体名'],
            ['{{rep_name}}',     '代表者役職・代表者名'],
            ['{{staff_name}}',   '担当者名'],
            ['{{category}}',     '区分 (S/A/B/C/D/E)'],
            ['{{receipt_no}}',   '受付番号'],
            ['{{event_name}}',   'イベント名（Info.EVENT_NAME）'],
            ['{{payment_due}}',  '支払い期限（Info.PAYMENT_DUE）'],
            ['{{office_email}}', '事務局メールアドレス'],
            ['{{office_hours}}', '受付時間'],
            ['{{date}}',         '現在日時'],
            ['{{amount}}',       '協賛金額（S/A請求書用）'],
        ],
        col_widths=[55*mm, 115*mm]
    ))
    story.append(PageBreak())

    # ── 6. 非機能要件 ──
    story += h1('6. 非機能要件')

    story += h2('6.1 セキュリティ')
    for t in [
        'サーバー側バリデーション: 郵便番号（7桁）・電話番号（9〜11桁）はGAS側で再検証。クライアントサイドのみのバリデーションには依存しない。',
        'LockService: appendRow() は ScriptLock を使用し、同時申込みによるデータ競合を防止する。',
        '印影画像: DriveファイルをBase64変換してHTMLに埋め込み。外部URLによるアクセス制限を回避。',
        '反社チェック: 誓約事項に暴力団・反社会的勢力への対応条項を明記。',
    ]:
        story.append(bullet(t))
    story.append(sp())

    story += h2('6.2 可用性・運用')
    for t in [
        'メール残数監視: GAS の 1日あたりのメール送信上限（MIN_MAIL_QUOTA、既定5通）を下回ると、フォームをメンテナンス表示にし、オーナーへ警告メールを30分に1通送信する。',
        'Info シート異常検知: Info シートが存在しない場合、フォームアクセス時にオーナーへ警告メールを送信（30分に1通）。',
        'Script Properties キャッシュ: Info シートの値を PropertiesService にキャッシュし、トリガーコンテキスト（onEditInstallable）からも参照可能にする。',
        'データ分離: DATA_SPREADSHEET_ID を設定することで、GASが紐づくMainスプレッドシートとは別のスプレッドシートにデータを保存できる。',
        'トリガー: onEditInstallable は Installed Trigger（GASエディタで手動登録）。シンプルトリガーではUIアラートが表示できないため。',
    ]:
        story.append(bullet(t))

    story.append(sp(5))
    story.append(HRFlowable(width='100%', thickness=1, color=COL_LIGHT))
    story.append(Paragraph('以上', ParagraphStyle('end', fontName=FONT_NORM, fontSize=9,
                                                    textColor=COL_GRAY, alignment=2, spaceBefore=4)))

    return story


# ═══════════════════════════════════════════════════════════════════════════
#  詳細設計書
# ═══════════════════════════════════════════════════════════════════════════
def build_shousai():
    story = []
    today = datetime.date.today().strftime('%Y年%m月%d日')

    # ── 表紙 ──
    story.append(Spacer(1, 30*mm))
    story.append(Paragraph('川口花火大会', ST['cover_sub']))
    story.append(Paragraph('協賛管理システム', ST['cover_title']))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph('詳細設計書', ParagraphStyle('ct3',
        fontName=FONT_BOLD, fontSize=20, textColor=COL_SECONDARY,
        leading=28, spaceAfter=6, alignment=1)))
    story.append(Spacer(1, 8*mm))
    story.append(HRFlowable(width='80%', thickness=2, color=COL_SECONDARY, hAlign='CENTER'))
    story.append(Spacer(1, 8*mm))
    info_rows = [
        ['バージョン', 'v1.0'],
        ['作成日',     today],
        ['対象ファイル', 'gs/*.gs, *.html (Google Apps Script プロジェクト)'],
    ]
    it = Table(info_rows, colWidths=[50*mm, 110*mm],
               style=TableStyle([
                   ('FONTNAME',  (0,0),(-1,-1), FONT_NORM),
                   ('FONTSIZE',  (0,0),(-1,-1), 10),
                   ('FONTNAME',  (0,0),(0,-1),  FONT_BOLD),
                   ('TEXTCOLOR', (0,0),(0,-1),  COL_PRIMARY),
                   ('ALIGN',     (0,0),(0,-1),  'RIGHT'),
                   ('GRID',      (0,0),(-1,-1), 0.5, colors.HexColor('#dddddd')),
                   ('BACKGROUND',(0,0),(-1,-1), COL_LIGHT),
                   ('LEFTPADDING', (0,0),(-1,-1), 8),
                   ('RIGHTPADDING',(0,0),(-1,-1), 8),
                   ('TOPPADDING',  (0,0),(-1,-1), 5),
                   ('BOTTOMPADDING',(0,0),(-1,-1), 5),
               ]))
    story.append(it)
    story.append(PageBreak())

    # ── 目次 ──
    story += h1('目次')
    toc = [
        ('1. ファイル構成', '3'),
        ('2. Config.gs', '3'),
        ('  2.1 定数・シートヘッダー', '3'),
        ('  2.2 Info シート読み込み', '3'),
        ('  2.3 設定取得関数', '4'),
        ('3. WebApp.gs', '4'),
        ('  3.1 doGet()', '4'),
        ('  3.2 getConfig()', '5'),
        ('  3.3 submitFormJson() / submitForm()', '5'),
        ('  3.4 メール残数管理', '6'),
        ('4. Sheet.gs', '6'),
        ('  4.1 appendRow()', '6'),
        ('  4.2 appendToTesagyouSheet()', '7'),
        ('  4.3 その他ユーティリティ', '7'),
        ('5. Mail.gs', '8'),
        ('  5.1 メール送信関数', '8'),
        ('  5.2 PDF生成関数', '9'),
        ('  5.3 テンプレート変数置換', '9'),
        ('6. Trigger.gs', '10'),
        ('  6.1 onEditInstallable()', '10'),
        ('  6.2 handleUketsuke()', '10'),
        ('  6.3 handleNyukin()', '11'),
        ('  6.4 ダイアログコールバック', '11'),
        ('7. Setup.gs', '12'),
        ('  7.1 setupInfoSheet()', '12'),
        ('  7.2 syncInfoToScriptProperties()', '12'),
        ('  7.3 メールテンプレート保存', '12'),
        ('8. index.html（フロントエンド）', '13'),
        ('  8.1 初期化・設定取得', '13'),
        ('  8.2 フォームバリデーション', '13'),
        ('  8.3 送信処理', '14'),
        ('9. シーケンス図', '15'),
    ]
    for text, page in toc:
        indent = 16 if text.startswith('  ') else 0
        st = ParagraphStyle('toc_item', fontName=FONT_NORM, fontSize=10,
                            textColor=COL_PRIMARY if indent==0 else COL_GRAY,
                            leading=18, leftIndent=indent)
        if indent == 0:
            st.fontName = FONT_BOLD
        story.append(Paragraph(text, st))
    story.append(PageBreak())

    # ── 1. ファイル構成 ──
    story += h1('1. ファイル構成')
    story.append(simple_table(
        ['ファイル', '役割'],
        [
            ['gs/Config.gs',   '定数・設定値・Infoシート読み込み・共通ユーティリティ'],
            ['gs/WebApp.gs',   'Webアプリエントリーポイント(doGet)・フォーム受信(submitForm)・メール残数管理'],
            ['gs/Sheet.gs',    'スプレッドシートへの行書き込み・ヘッダー管理・列幅設定'],
            ['gs/Mail.gs',     'メール送信・請求書PDF生成・お礼状PDF生成・テンプレート管理'],
            ['gs/Trigger.gs',  'onEditInstallable（手作業シートのチェックボックストリガー）・ダイアログコールバック'],
            ['gs/Setup.gs',    'InfoシートセットアップとScriptProperties同期・メールテンプレート保存'],
            ['gs/KanriId.gs',  '（管理ID機能は削除済み・空ファイル）'],
            ['gs/ProjectInit.gs', 'プロジェクト初期化（initProject）'],
            ['index.html',     '協賛申込みWebフォーム（メインUI）'],
            ['Stylesheet.html','CSSスタイルシート（include方式で埋め込み）'],
            ['invoice-template.html', '請求書PDFのHTMLテンプレート'],
            ['oreijou-template.html', 'お礼状PDFのHTMLテンプレート'],
            ['ConfirmInvoiceDialog.html', '請求書送信確認ダイアログUI'],
            ['ConfirmNyukinDialog.html',  '入金確認・お礼状送信ダイアログUI'],
            ['ConfirmOreijouDialog.html', 'お礼状確認ダイアログUI（予備）'],
        ],
        col_widths=[65*mm, 105*mm]
    ))
    story.append(PageBreak())

    # ── 2. Config.gs ──
    story += h1('2. Config.gs')

    story += h2('2.1 定数・シートヘッダー')
    story.append(simple_table(
        ['定数名', '値', '説明'],
        [
            ['DEFAULT_SHEET_NAME',  '協賛申込み一覧', '申込みデータシート名'],
            ['DEFAULT_SHEET_NAME2', '手作業',         '管理用手作業シート名'],
            ['INFO_SHEET_NAME',     'Info',           '設定マスターシート名'],
            ['CREATELOG_SHEET',     'CreateLog',      'ログシート名（拡張用）'],
            ['AUTO_SEND_KUBUN',     "['B','C','D','E']", '申込時に自動で請求書メールを送る区分'],
            ['DEFAULT_PRICES',      'S:200万, A:100万, B:50万,\nC:30万, D:20万, E:10万',
             'Info シート未設定時のデフォルト協賛金額'],
            ['COL_RECEPT_NO',       '1 (A列)',         '手作業シートの受付番号列'],
            ['COL_UKETSUKE',        '9 (I列)',         '手作業シートの受付完了チェックボックス列'],
            ['COL_INV_DATE',        '10 (J列)',        '手作業シートの請求書送信日時列'],
            ['COL_NYUKIN',          '11 (K列)',        '手作業シートの入金完了チェックボックス列'],
            ['COL_OREIJOU_DATE',    '12 (L列)',        '手作業シートのお礼状送信日時列'],
        ],
        col_widths=[52*mm, 48*mm, 70*mm]
    ))
    story.append(sp())

    story += h2('2.2 Info シート読み込み')
    story.append(body('_infoCache 変数でInメモリキャッシュ。getInfoConfig() が最初の呼び出しで全行を読み込む。'))
    story.append(sp(2))
    story.append(simple_table(
        ['関数名', 'シグネチャ', '処理内容'],
        [
            ['getInfoConfig()', '() → Object',
             'Infoシートを全行読み込みKey-Valueオブジェクトを返す。キャッシュ済みなら即返す。'],
            ['getConfigVal(key, fallback)', '(string, any) → any',
             '1. _infoCache → 2. Script Properties（INFO_+key）→ 3. fallback の優先順で値を返す'],
            ['setInfoValue(key, value)', '(string, any) → void',
             'Infoシートのキーに対して値を更新。なければ新行を追加。キャッシュクリア。'],
        ],
        col_widths=[40*mm, 40*mm, 90*mm]
    ))
    story.append(sp())

    story += h2('2.3 設定取得関数（ラッパー）')
    story.append(body('各設定値に対して getConfigVal() を呼ぶ専用ラッパー関数が定義されている。'))
    story.append(sp(2))
    story.append(simple_table(
        ['関数名', '返す値'],
        [
            ['getOfficeEmail()',     'OFFICE_EMAIL'],
            ['getHankoFileId()',     'HANKO_FILE_ID（印影PNG DriveID）'],
            ['getBgImageId()',       'BG_IMAGE_ID（背景画像 DriveID）'],
            ['getEventName()',       'EVENT_NAME'],
            ['getPaymentDue()',      'PAYMENT_DUE（支払期限文字列）'],
            ['getRootFolderId()',    'ROOT_FOLDER_ID'],
            ['getOrgName()',         'ORG_NAME'],
            ['getOrgRep()',          'ORG_REP'],
            ['getInvoiceRegNo()',    'INVOICE_REG_NO（インボイス登録番号）'],
            ['getBankName()',        'BANK_NAME'],
            ['getBankNo()',          'BANK_NO'],
            ['getBankHolder()',      'BANK_HOLDER'],
            ['getBankRep()',         'BANK_REP'],
            ['getOrgLocation()',     'ORG_LOCATION'],
            ['getOrgTel()',          'ORG_TEL'],
            ['getOrgFax()',          'ORG_FAX'],
            ['getOfficeHours()',     'OFFICE_HOURS'],
            ['getReceiptNoPrefix()', 'RECEIPT_NO_PREFIX'],
            ['getKubunSaStart/End()', 'KUBUN_SA_START / KUBUN_SA_END'],
            ['getKubunBcdeStart/End()', 'KUBUN_BCDE_START / KUBUN_BCDE_END'],
            ['getMinMailQuota()',    'MIN_MAIL_QUOTA（最小5）'],
            ['getCategoryPrice(cat)', 'PRICE_{S/A/B/C/D/E}（なければDEFAULT_PRICES）'],
            ['getDataSpreadsheet()', 'DATA_SPREADSHEET_IDが設定されていれば別SS、なければMain'],
        ],
        col_widths=[65*mm, 105*mm]
    ))
    story.append(PageBreak())

    # ── 3. WebApp.gs ──
    story += h1('3. WebApp.gs')

    story += h2('3.1 doGet()')
    story.append(body('GAS WebアプリのHTTP GETエントリーポイント。以下の順序で処理する。'))
    story.append(sp(2))
    steps = [
        ('Step 1', 'MAIN_SS_ID をScript Propertiesに保存（トリガーコンテキスト用）'),
        ('Step 2', 'Info シートの存在確認。存在しない場合: オーナーへ警告メール（30分に1通）→ エラーHTML返却'),
        ('Step 3', '_checkRequiredInfoKeys() で必須キーの未入力チェック。不足あり: _sendInfoIncompleteAlert() → エラーHTML返却'),
        ('Step 4', 'BG_IMAGE_ID が設定されていれば DriveApp でBlob取得 → Base64変換 → tpl.bgImageDataUrl にセット'),
        ('Step 5', 'HtmlService.createTemplateFromFile("Index") でテンプレート生成、タイトルを設定して返却'),
    ]
    for step, desc in steps:
        story.append(Paragraph(
            f'<b>{step}:</b> {desc}', ST['body']))
    story.append(sp())

    story += h3('必須Infoキー (_checkRequiredInfoKeys)')
    req_keys = ['EVENT_NAME', 'OFFICE_EMAIL', 'START_DATE', 'END_DATE',
                'PAYMENT_DUE', 'ORG_NAME', 'ORG_REP', 'BANK_NAME',
                'BANK_NO', 'BANK_HOLDER', 'BG_IMAGE_ID']
    story.append(body('以下のキーが空の場合はエラーとなる: ' + '、'.join(req_keys)))
    story.append(sp())

    story += h2('3.2 getConfig()')
    story.append(body('クライアント（google.script.run）から呼ばれる設定取得関数。メール残数もここで確認する。'))
    story.append(sp(2))
    story.append(simple_table(
        ['返却フィールド', '値'],
        [
            ['startDate',         'Info.START_DATE'],
            ['endDate',           'Info.END_DATE'],
            ['sheetName1',        'DEFAULT_SHEET_NAME（協賛申込み一覧）'],
            ['sheetName2',        'DEFAULT_SHEET_NAME2（手作業）'],
            ['eventName',         'Info.EVENT_NAME'],
            ['receiptNoPrefix',   'Info.RECEIPT_NO_PREFIX'],
            ['mailQuotaOk',       'MailApp.getRemainingDailyQuota() >= getMinMailQuota()'],
            ['kubunSaStart/End',  'Info.KUBUN_SA_START / KUBUN_SA_END'],
            ['kubunBcdeStart/End','Info.KUBUN_BCDE_START / KUBUN_BCDE_END'],
        ],
        col_widths=[60*mm, 110*mm]
    ))
    story.append(sp())

    story += h2('3.3 submitFormJson() / submitForm()')
    story.append(body('フォーム送信を処理するメイン関数。submitFormJson は JSON文字列をパースして submitForm() に委譲する。'))
    story.append(sp(2))
    steps2 = [
        ('Step 1', 'メール残数チェック（quota < MIN_MAIL_QUOTA）。不足の場合はデータ記録前に例外をスロー。_notifyLowQuota() を呼ぶ。'),
        ('Step 2', 'Info シート存在確認。存在しない場合はオーナーへ警告メールを送信し例外スロー。'),
        ('Step 3', 'サーバー側バリデーション: zipcode → /^\\d{7}$/ 、phone → /^\\d{9,11}$/ 。不正なら例外スロー。'),
        ('Step 4', 'appendRow(data, DEFAULT_SHEET_NAME) → 受付番号(receptNo)を取得。'),
        ('Step 5', 'appendToTesagyouSheet(receptNo, DEFAULT_SHEET_NAME2, data) → 手作業シートに行追加。'),
        ('Step 6', '区分が AUTO_SEND_KUBUN（B/C/D/E）なら generateInvoicePdf() → sendConfirmationEmail()。S/A なら sendReceiptOnlyEmail()。'),
        ('Step 7', '{ result: "success", receipt_no: receptNo } を返す。'),
    ]
    for step, desc in steps2:
        story.append(Paragraph(f'<b>{step}:</b> {desc}', ST['body']))
    story.append(sp())

    story += h2('3.4 メール残数管理')
    story.append(simple_table(
        ['関数名', 'シグネチャ', '処理内容'],
        [
            ['_getMailQuotaSafe()', '() → number',
             'MailApp.getRemainingDailyQuota() を try-catch で保護。失敗時は 0 を返す（送信不可扱い）。'],
            ['_notifyLowQuota(quota)', '(number) → void',
             'LockService で直列化。30分以内の重複通知を ScriptProperties で制御。オーナーへ警告メールを送信。'],
            ['getBgImageDataUrl()', '() → string',
             'クライアントから呼ばれる（フォールバック用）。DriveAPI経由でBase64データURLを返す。'],
        ],
        col_widths=[45*mm, 35*mm, 90*mm]
    ))
    story.append(PageBreak())

    # ── 4. Sheet.gs ──
    story += h1('4. Sheet.gs')

    story += h2('4.1 appendRow()')
    story.append(simple_table(
        ['項目', '詳細'],
        [
            ['シグネチャ', 'appendRow(data: Object, sheetName: string) → string (receptNo)'],
            ['排他制御',   'LockService.getScriptLock().waitLock(10000) で同時申込みの競合を防止'],
            ['シート自動作成', 'getSheetByName() が null の場合 insertSheet() で作成'],
            ['ヘッダー自動生成', 'getLastRow() === 0 の場合 HEADERS を書き込み、太字・背景・フリーズ・フィルターを設定'],
            ['受付番号生成', 'data.receipt_no があれば使用。なければ getReceiptNoPrefix() + yyyyMMddHHmmssSSS'],
            ['郵便番号・電話番号書式', 'setNumberFormat("@") でテキスト書式に設定し先頭ゼロを保持'],
        ],
        col_widths=[40*mm, 130*mm]
    ))
    story.append(sp())

    story += h2('4.2 appendToTesagyouSheet()')
    story.append(simple_table(
        ['項目', '詳細'],
        [
            ['シグネチャ', 'appendToTesagyouSheet(receptNo, sheetName2, data) → void'],
            ['XLOOKUP数式', 'TESAGYOU_LOOKUP_COLS の各列に\n=IFERROR(XLOOKUP($A{row},\'協賛申込み一覧\'!$A:$A,\'協賛申込み一覧\'!${col}:${col}),"見つかりません")\nを設定'],
            ['チェックボックス', 'I列（COL_UKETSUKE）・K列（COL_NYUKIN）に insertCheckboxes()'],
            ['自動チェック（B/C/D/E）', 'AUTO_SEND_KUBUN の場合: I列を true、J列（COL_INV_DATE）に nowStr() をセット'],
            ['ヘッダー/サブヘッダー', '行1: TESAGYOU_HEADERS（背景#fce8b2）\n行2: サブヘッダー（8pt, #888888, #fffbf0）\nフリーズ2行'],
        ],
        col_widths=[45*mm, 125*mm]
    ))
    story.append(sp())

    story += h2('4.3 その他ユーティリティ')
    story.append(simple_table(
        ['関数名', '処理内容'],
        [
            ['resetTesagyouSheetHeaders()', '手作業シートの余分な列を削除しヘッダーを再設定（カスタムメニューから実行）'],
            ['applyColumnWidths(sheet)', '協賛申込み一覧の列幅を固定値で設定。データ行を上揃え・左揃えにする'],
            ['applyTesagyouColumnWidths(sheet)', '手作業シートの列幅を固定値で設定（A〜L, 12列）'],
            ['_applyAlignment(sheet, firstDataRow, numCols)', 'データ行全体を verticalAlignment=top, horizontalAlignment=left に設定'],
            ['_ensureFilter(sheet, headerRow, numCols)', 'フィルターが未設定の場合のみ createFilter() を呼ぶ（冪等）'],
        ],
        col_widths=[70*mm, 100*mm]
    ))
    story.append(PageBreak())

    # ── 5. Mail.gs ──
    story += h1('5. Mail.gs')

    story += h2('5.1 メール送信関数')
    story.append(simple_table(
        ['関数名', 'シグネチャ', '処理内容'],
        [
            ['sendConfirmationEmail()', '(data, receptNo, invoicePdf) → void',
             'B/C/D/E: 申込確認メール + 請求書PDF添付。Script Properties の MAIL_SUBJECT/MAIL_BODY を使用。officeEmail が有効なら cc/replyTo にセット。'],
            ['sendReceiptOnlyEmail()', '(data, receptNo, invoicePdf?) → void',
             'S/A: 受付確認のみ（抽選待ち）。RECEIPT_ONLY_SUBJECT/RECEIPT_ONLY_BODY を使用。'],
            ['sendSaInvoiceEmail()', '(data, receptNo, invoicePdf) → void',
             'S/A が抽選確定後: SA_INVOICE_SUBJECT/SA_INVOICE_BODY を使用。_buildSaInvoiceVars() で変数を構築。'],
            ['sendOreijouEmail()', '(data, receptNo) → void',
             '入金確認後のお礼状メール。OREIJOU_SUBJECT/OREIJOU_BODY を使用。generateOreijouPdf() でPDFを生成して添付。'],
        ],
        col_widths=[48*mm, 48*mm, 74*mm]
    ))
    story.append(sp())

    story += h2('5.2 PDF生成関数')
    story.append(simple_table(
        ['関数名', 'シグネチャ', '処理内容'],
        [
            ['generateInvoicePdf()', '(data, receptNo) → Blob',
             '1. invoice-template.html の内容を取得\n2. 印影PNG（DriveAPI）をBase64変換してsrc="hanko.png"と置換\n3. テンプレート変数をすべて文字列置換\n4. DriveApp.createFile(html blob) → getAs(PDF) → 一時ファイルをゴミ箱へ'],
            ['generateOreijouPdf()', '(data) → Blob|null',
             '1. oreijou-template.html の内容を取得\n2. 同様に印影PNG置換\n3. お礼状固有変数（org_location, org_tel, org_fax等）を置換\n4. PDF変換。エラー時はnullを返す（try-catch）'],
        ],
        col_widths=[40*mm, 40*mm, 90*mm]
    ))
    story.append(sp())

    story += h3('請求書テンプレート変数一覧（invoice-template.html）')
    story.append(simple_table(
        ['変数', '内容'],
        [
            ['{{company_name}}',   'data.company_name'],
            ['{{issue_date}}',     '令和{n}年{m}月{d}日（サーバー現在時刻）'],
            ['{{receipt_no}}',     '受付番号'],
            ['{{category}}',       '区分（S/A/B/C/D/E）'],
            ['{{total}}',          '税込金額（¥xxx,xxx）'],
            ['{{subtotal}}',       '税抜金額（total / 1.1、端数切り捨て）'],
            ['{{tax}}',            '消費税（total - subtotal）'],
            ['{{payment_due}}',    'Info.PAYMENT_DUE'],
            ['{{event_name}}',     'Info.EVENT_NAME'],
            ['{{org_name}}',       'Info.ORG_NAME'],
            ['{{org_rep}}',        'Info.ORG_REP'],
            ['{{invoice_reg_no}}', 'Info.INVOICE_REG_NO'],
            ['{{bank_name}}',      'Info.BANK_NAME'],
            ['{{bank_no}}',        'Info.BANK_NO'],
            ['{{bank_holder}}',    'Info.BANK_HOLDER'],
            ['{{bank_rep}}',       'Info.BANK_REP'],
        ],
        col_widths=[55*mm, 115*mm]
    ))
    story.append(sp())

    story += h2('5.3 テンプレート変数置換')
    story.append(simple_table(
        ['関数名', '処理内容'],
        [
            ['_replaceVars(str, vars)', '{{key}} を vars[key] で置換。RegExp で全出現を置換（gフラグ）'],
            ['_buildVars(data, receptNo)', '共通変数オブジェクトを構築（company_name, staff_name, category, receipt_no, event_name, payment_due, office_email, office_hours, date）'],
            ['_buildSaInvoiceVars(data, receptNo)', 'S/A請求書メール用変数（_buildVars + amount）を構築'],
            ['_validEmail(email)', '"@" を含む かつ "Default" を含まない → true（プレースホルダーを弾く）'],
        ],
        col_widths=[55*mm, 115*mm]
    ))
    story.append(PageBreak())

    # ── 6. Trigger.gs ──
    story += h1('6. Trigger.gs')

    story += h2('6.1 onEditInstallable()')
    story.append(body(
        '登録方法: GASエディタ → トリガー → onEditInstallable → スプレッドシート → 編集時（Installed Trigger）。\n'
        'Simple Trigger ではダイアログ表示ができないためInstalled Triggerが必要。'
    ))
    story.append(sp(2))
    story.append(simple_table(
        ['条件', '処理'],
        [
            ['sheet.getName() !== DEFAULT_SHEET_NAME2', 'return（手作業シート以外は無視）'],
            ['row <= 2', 'return（ヘッダー行・サブヘッダー行は無視）'],
            ['col === COL_UKETSUKE (9)', 'handleUketsuke(e, sheet, row) 呼び出し'],
            ['col === COL_NYUKIN (11)',   'handleNyukin(e, sheet, row) 呼び出し'],
        ],
        col_widths=[80*mm, 90*mm]
    ))
    story.append(sp())

    story += h2('6.2 handleUketsuke() — I列受付完了')
    story.append(simple_table(
        ['状態', '処理'],
        [
            ['false にしようとした（チェック解除）', '下流（K列チェック/L列日時/J列日時）が存在する場合は YES/NO 確認ダイアログ。YES: K・L・Jをクリア。NO: I列を true に戻す。'],
            ['J列（請求書送信日時）が既に存在', 'アラート「送信済み」を表示し I列を false に戻す'],
            ['メール残数不足', '_blockSendIfLowQuota() → true の場合はブロック（I列を false に戻す）'],
            ['正常チェック', '申込みシートから受付番号で findRowByReceptNo() → data取得。ConfirmInvoiceDialog.html をモーダル表示（420×280px）'],
        ],
        col_widths=[55*mm, 115*mm]
    ))
    story.append(sp())

    story += h2('6.3 handleNyukin() — K列入金完了')
    story.append(simple_table(
        ['状態', '処理'],
        [
            ['false にしようとした（チェック解除）', 'L列（お礼状送信日時）が存在する場合は YES/NO 確認ダイアログ。YES: Lをクリア。NO: K列を true に戻す。'],
            ['J列（請求書送信日時）が未設定', 'アラート「請求書がまだ送信されていません」を表示し K列を false に戻す'],
            ['L列（お礼状送信日時）が既に存在', 'アラート「お礼状は送信済み」を表示し K列を false に戻す'],
            ['メール残数不足', '_blockSendIfLowQuota() でブロック'],
            ['正常チェック', 'findRowByReceptNo() → data取得。ConfirmNyukinDialog.html をモーダル表示（420×250px）'],
        ],
        col_widths=[55*mm, 115*mm]
    ))
    story.append(sp())

    story += h2('6.4 ダイアログコールバック（google.script.run から呼ばれる）')
    story.append(simple_table(
        ['関数名', '処理内容'],
        [
            ['sendInvoiceConfirmed(row, receptNo)',
             'mainSheetからfindRowByReceptNo()でdata取得 → generateInvoicePdf() → 区分がS/AならsendSaInvoiceEmail()、それ以外はsendConfirmationEmail() → J列に nowStr() をセット'],
            ['cancelInvoiceSend(row)',
             '手作業シートのI列（COL_UKETSUKE）を false に戻す'],
            ['sendNyukinConfirmed(row, receptNo)',
             'findRowByReceptNo() → sendOreijouEmail() → L列に nowStr() をセット'],
            ['cancelNyukin(row)',
             '手作業シートのK列（COL_NYUKIN）を false に戻す'],
            ['findRowByReceptNo(sheet, receptNo)',
             '協賛申込み一覧の全行を走査し受付番号（A列）が一致する行のdataオブジェクトを返す。見つからなければ null。'],
            ['_blockSendIfLowQuota(revertCell)',
             '残数 < MIN_MAIL_QUOTA の場合: _notifyLowQuota() → アラート表示 → revertCell を false に戻す → true返却。問題なければ false返却。'],
        ],
        col_widths=[60*mm, 110*mm]
    ))
    story.append(PageBreak())

    # ── 7. Setup.gs ──
    story += h1('7. Setup.gs')

    story += h2('7.1 setupInfoSheet()')
    story.append(body('カスタムメニュー「⚙️ 初期設定 → Info シート作成」から実行。既存シートがある場合は上書き確認。'))
    story.append(body('設定キー・デフォルト値・説明を4列（キー、値、必須、説明）で書き込む。ヘッダー行・フィルター・列幅・フリーズを設定後、syncInfoToScriptProperties() を呼ぶ。'))
    story.append(sp())

    story += h2('7.2 syncInfoToScriptProperties()')
    story.append(body('Info シートの全行を読み込み、PropertiesService.setProperties() で一括保存。キーは "INFO_" プレフィックスを付加する。キャッシュ（_infoCache）をリセット。'))
    story.append(sp())

    story += h2('7.3 メールテンプレート保存')
    story.append(simple_table(
        ['関数名', '保存するScript Propertiesキー'],
        [
            ['setupMailTemplate()',          'MAIL_SUBJECT, MAIL_BODY'],
            ['setupReceiptOnlyTemplate()',   'RECEIPT_ONLY_SUBJECT, RECEIPT_ONLY_BODY'],
            ['setupSaInvoiceTemplate()',     'SA_INVOICE_SUBJECT, SA_INVOICE_BODY'],
            ['setupOreijouTemplate()',       'OREIJOU_SUBJECT, OREIJOU_BODY'],
            ['setupAllMailTemplates()',      '上記4つをまとめて実行'],
        ],
        col_widths=[70*mm, 100*mm]
    ))
    story.append(sp())

    story += h2('7.4 その他メンテナンス関数')
    story.append(simple_table(
        ['関数名', '処理内容'],
        [
            ['repairTesagyouNow()', 'Main または DataスプレッドシートのTesagyouシートにヘッダーを再適用（_applyTesagyouHeaders）'],
            ['renameHeaders()', '協賛申込み一覧・手作業シートの1行目を最新ヘッダーに上書き（データ行は保持）'],
            ['applyFiltersToAllSheets()', 'Info/CreateLog/協賛申込み一覧/手作業 の全シートにフィルターを一括設定'],
        ],
        col_widths=[55*mm, 115*mm]
    ))
    story.append(PageBreak())

    # ── 8. index.html ──
    story += h1('8. index.html（フロントエンド）')

    story += h2('8.1 初期化・設定取得')
    story.append(body('DOMContentLoaded 後に google.script.run.getConfig() を呼び出し、返却値に基づいてUIを制御する。'))
    story.append(sp(2))
    story.append(simple_table(
        ['制御', '条件', '表示要素'],
        [
            ['受付時間外', 'now < startDate または now > endDate', 'form-closed-message を表示'],
            ['メンテナンス', 'config.mailQuotaOk === false', 'maintenance-message を表示'],
            ['受付中', '上記以外', 'registration-form を表示。_applyKubunPeriod() で区分ごとに表示制御'],
        ],
        col_widths=[35*mm, 60*mm, 75*mm]
    ))
    story.append(sp(2))
    story += h3('_applyKubunPeriod(config, now)')
    story.append(body('S/A: kubunSaStart 〜 kubunSaEnd の期間内のみ表示。B〜E: kubunBcdeStart 〜 kubunBcdeEnd の期間内のみ表示。両方外の場合はクローズメッセージを表示。'))
    story.append(sp())

    story += h2('8.2 フォームバリデーション')
    story.append(simple_table(
        ['バリデーション', '実装方法'],
        [
            ['必須チェック', 'required 属性 + setCustomValidity() でブラウザ標準バリデーション'],
            ['郵便番号', 'pattern=\\d{7}、入力時に \\D を除去（input イベント）'],
            ['電話番号', 'pattern=\\d{9,11}、入力時に \\D を除去（input イベント）'],
            ['フリガナ（会社名）', 'pattern=^[ァ-ヶー\\s　）]+$、blur時に全角閉じカッコ変換'],
            ['フリガナ（人名）', 'pattern=^[ァ-ヶー\\s　]+$'],
            ['誓約事項スクロール', 'terms-box の scroll イベントで scrollTop + clientHeight >= scrollHeight - 2 をチェック → チェックボックスを有効化'],
            ['同意チェック', 'checkbox の change イベントで「入力内容を確認する」ボタンの disabled を制御'],
        ],
        col_widths=[40*mm, 130*mm]
    ))
    story.append(sp())

    story += h2('8.3 送信処理')
    story.append(simple_table(
        ['ステップ', '処理'],
        [
            ['① フォーム submit', 'form.addEventListener("submit") でデフォルト動作を preventDefault()\nFormData から savedFormData（Object）を構築\n確認テーブルを生成し confirm-screen を表示'],
            ['② 「修正する」クリック', 'confirm-screen を非表示、form を再表示'],
            ['③ 「送信する」クリック', 'submitBtn・backBtn を disabled に\n受付番号をクライアント側で生成（prefix + MMddHHmmssSSS）\nフィールドをfilterして data オブジェクトを構築\nJSON.stringify(data) → submitFormJson() をgoogle.script.runで呼び出し'],
            ['④ 成功ハンドラー', 'confirm-screen を非表示\n受付番号を表示した「送信完了」メッセージを挿入'],
            ['⑤ 失敗ハンドラー', 'ボタンを再有効化\nalert() でエラーメッセージを表示'],
        ],
        col_widths=[30*mm, 140*mm]
    ))
    story.append(PageBreak())

    # ── 9. シーケンス図 ──
    story += h1('9. 主要シーケンス図（テキスト表現）')

    story += h2('9.1 B/C/D/E 区分 申込みフロー')
    seq1 = [
        '申込者 → [ブラウザ] → google.script.run.getConfig()',
        '[GAS] getConfig(): Info読込 → メール残数チェック → 設定返却',
        '[ブラウザ] フォーム表示 → 入力 → 確認画面 → 「送信する」',
        '[ブラウザ] → google.script.run.submitFormJson(dataJson)',
        '[GAS] submitForm():',
        '  ① メール残数チェック（不足なら例外）',
        '  ② Info シート確認',
        '  ③ サーバー側バリデーション（zipcode, phone）',
        '  ④ appendRow() → 協賛申込み一覧に書き込み → receptNo取得',
        '  ⑤ appendToTesagyouSheet() → 手作業シートに行追加・チェック済み・請求書日時セット',
        '  ⑥ generateInvoicePdf() → 請求書PDF生成',
        '  ⑦ sendConfirmationEmail() → 申込者へメール + PDF添付',
        '[GAS] → { result: "success", receipt_no: receptNo }',
        '[ブラウザ] 送信完了メッセージ表示',
    ]
    for line in seq1:
        indent = 14 if line.startswith('  ') else 0
        story.append(Paragraph(line, ParagraphStyle('seq',
            fontName=FONT_NORM, fontSize=9, textColor=colors.black,
            leading=14, leftIndent=indent, backColor=COL_LIGHT_GRAY if indent else colors.white)))
    story.append(sp())

    story += h2('9.2 S/A 区分 抽選確定後 請求書送信フロー')
    seq2 = [
        '事務局スタッフ → [スプレッドシート] 手作業シート I列（受付完了）をチェック',
        '[GAS] onEditInstallable():',
        '  col === COL_UKETSUKE(9) → handleUketsuke()',
        '  メール残数チェック',
        '  findRowByReceptNo() → data取得',
        '  ConfirmInvoiceDialog.html をモーダル表示',
        'スタッフ → 「送信する」クリック（ダイアログ）',
        '[GAS] sendInvoiceConfirmed(row, receptNo):',
        '  findRowByReceptNo() → data',
        '  generateInvoicePdf()',
        '  区分S/A → sendSaInvoiceEmail()',
        '  J列（COL_INV_DATE）に nowStr() をセット',
        '[メール送信] 申込者へ 抽選確定・請求書PDF 添付メール',
    ]
    for line in seq2:
        indent = 14 if line.startswith('  ') else 0
        story.append(Paragraph(line, ParagraphStyle('seq2',
            fontName=FONT_NORM, fontSize=9, textColor=colors.black,
            leading=14, leftIndent=indent, backColor=COL_LIGHT_GRAY if indent else colors.white)))
    story.append(sp())

    story += h2('9.3 お礼状送信フロー（入金確認後）')
    seq3 = [
        '事務局スタッフ → [スプレッドシート] 手作業シート K列（入金完了）をチェック',
        '[GAS] onEditInstallable():',
        '  col === COL_NYUKIN(11) → handleNyukin()',
        '  J列（請求書送信日時）の存在確認（なければブロック）',
        '  メール残数チェック',
        '  findRowByReceptNo() → data取得',
        '  ConfirmNyukinDialog.html をモーダル表示',
        'スタッフ → 「送信する」クリック（ダイアログ）',
        '[GAS] sendNyukinConfirmed(row, receptNo):',
        '  findRowByReceptNo() → data',
        '  sendOreijouEmail():',
        '    generateOreijouPdf()',
        '    MailApp.sendEmail() with PDF添付',
        '  L列（COL_OREIJOU_DATE）に nowStr() をセット',
        '[メール送信] 申込者へ お礼状PDF 添付メール',
    ]
    for line in seq3:
        indent = 14 if line.startswith('  ') else 0
        story.append(Paragraph(line, ParagraphStyle('seq3',
            fontName=FONT_NORM, fontSize=9, textColor=colors.black,
            leading=14, leftIndent=indent, backColor=COL_LIGHT_GRAY if indent else colors.white)))

    story.append(sp(5))
    story.append(HRFlowable(width='100%', thickness=1, color=COL_LIGHT))
    story.append(Paragraph('以上', ParagraphStyle('end', fontName=FONT_NORM, fontSize=9,
                                                    textColor=COL_GRAY, alignment=2, spaceBefore=4)))

    return story


def add_page_number(canvas, doc):
    canvas.saveState()
    canvas.setFont(FONT_NORM, 8)
    canvas.setFillColor(COL_GRAY)
    page_num = canvas.getPageNumber()
    canvas.drawRightString(195*mm, 10*mm, f'{page_num}')
    canvas.drawString(20*mm, 10*mm, '川口花火大会 協賛管理システム')
    canvas.restoreState()


def build_doc(filename, story_fn, title):
    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        rightMargin=20*mm,
        leftMargin=20*mm,
        topMargin=20*mm,
        bottomMargin=20*mm,
        title=title,
        author='WebHanabi System',
    )
    story = story_fn()
    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    print(f'✅ Created: {filename}')


if __name__ == '__main__':
    build_doc(
        '/Users/minhsang1601/GitClone/WebHanabi/基本設計書_WebHanabi.pdf',
        build_kihon,
        '川口花火大会 協賛管理システム 基本設計書'
    )
    build_doc(
        '/Users/minhsang1601/GitClone/WebHanabi/詳細設計書_WebHanabi.pdf',
        build_shousai,
        '川口花火大会 協賛管理システム 詳細設計書'
    )

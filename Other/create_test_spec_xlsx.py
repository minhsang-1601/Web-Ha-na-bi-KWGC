#!/usr/bin/env python3
"""テスト仕様書_WebHanabi.xlsx を生成する。"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# ── 配色 ────────────────────────────────────────────────
NAVY   = "1A3A5C"
BLUE   = "2E6DA4"
LBLUE  = "DCE9F5"
LLBLUE = "F0F5FB"
GREEN  = "27AE60"
LGREEN = "D5F0E0"
ORANGE = "E67E22"
LORANGE= "FDE9D4"
PURPLE = "6C3483"
LPURPLE= "E8D5F5"
TEAL   = "117A65"
LTEAL  = "D1F2EB"
GRAY   = "7F8C8D"
LGRAY  = "ECF0F1"
WHITE  = "FFFFFF"
BLACK  = "222222"
FONT = "Meiryo"

def F(size=10, bold=False, color=BLACK, name=FONT):
    return Font(name=name, size=size, bold=bold, color=color)
def fill(c):
    return PatternFill("solid", fgColor=c)
thin = Side(style="thin", color="BBBBBB")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
def cell(ws, coord, value, font=None, bg=None, align=None, wrap=False, bd=True):
    c = ws[coord]
    c.value = value
    if font: c.font = font
    if bg: c.fill = fill(bg)
    c.alignment = Alignment(horizontal=(align or "left"), vertical="center", wrap_text=wrap)
    if bd: c.border = border
    return c

wb = Workbook()

# ════════════════════════════════════════════════════════
# Sheet 1: 表紙・概要
# ════════════════════════════════════════════════════════
ws = wb.active
ws.title = "表紙・概要"
ws.sheet_view.showGridLines = False
for col, w in {"A":3, "B":22, "C":30, "D":30, "E":30, "F":12}.items():
    ws.column_dimensions[col].width = w

ws.merge_cells("B2:F2")
cell(ws, "B2", "テスト仕様書", F(28, True, WHITE), NAVY, "left", bd=False)
ws.row_dimensions[2].height = 48
ws.merge_cells("B3:F3")
cell(ws, "B3", "WebHanabi — 川口花火大会 協賛管理システム", F(13, True, WHITE), BLUE, "left", bd=False)
ws.row_dimensions[3].height = 26
ws.merge_cells("B4:F4")
cell(ws, "B4", "発行日：2026年06月20日　　バージョン：1.0　　対象：QA担当者・開発者", F(10, False, BLACK), LLBLUE, "left", bd=False)

# 目的
ws.merge_cells("B6:F6")
cell(ws, "B6", "■ 目的", F(12, True, WHITE), BLUE)
ws.merge_cells("B7:F8")
cell(ws, "B7", "本書は WebHanabi システムの各機能が仕様通りに動作することを確認するためのテスト仕様を定義する。"
              "申込みフォームの入力受付から、メール送信・スプレッドシート管理・ワークフロー制御までの全機能をカバーする。",
     F(10), LLBLUE, "left", wrap=True)
ws.row_dimensions[7].height = 22

# 対象 / 対象外 / 種別
heads = [("対象", BLUE, LBLUE, [
            "協賛申込みWebフォーム", "スプレッドシートへのデータ登録",
            "メール自動送信（請求書・お礼状）", "onEditトリガーとダイアログ動作", "Infoシート設定値の反映"]),
         ("対象外", "C0392B", "FCE4E4", [
            "Google インフラの可用性テスト", "メール到達性（Gmail SLA）",
            "スプレッドシートの行上限テスト", "ブラウザ互換性テスト", ""]),
         ("テスト種別", GREEN, LGREEN, [
            "機能テスト（黒箱）", "バリデーションテスト",
            "異常系テスト", "回帰テスト（設定変更後）", ""])]
r0 = 10
for i, (title, hc, bc, items) in enumerate(heads):
    col = get_column_letter(2 + 2*i) if i else "B"
# place them in columns B / C / D area — instead use one column block each
# Simpler: stack as three column groups B, D, F is messy. Use rows.
# Render each block in its own column: B, C+D, E+F
cols = ["B", "D", "E"]
# Lay out as three side-by-side blocks across B / C-D? keep simple: B, C, D-E-F single cols
blocks = [("B", heads[0]), ("D", heads[1]), ("E", heads[2])]
# Use merged 1-col-wide blocks: B(=B), C-D(=C:D), E-F(=E:F)
spans = [("B","B"), ("C","D"), ("E","F")]
for (c1, c2), (title, hc, bc, items) in zip(spans, heads):
    ws.merge_cells(f"{c1}{r0}:{c2}{r0}")
    cell(ws, f"{c1}{r0}", title, F(11, True, WHITE), hc, "center")
    for j, it in enumerate(items):
        rr = r0 + 1 + j
        ws.merge_cells(f"{c1}{rr}:{c2}{rr}")
        cell(ws, f"{c1}{rr}", ("• " + it) if it else "", F(9.5), bc, "left", wrap=True)
        ws.row_dimensions[rr].height = 30

# ════════════════════════════════════════════════════════
# Sheet 2: テスト環境
# ════════════════════════════════════════════════════════
ws2 = wb.create_sheet("テスト環境")
ws2.sheet_view.showGridLines = False
for col, w in {"A":3, "B":18, "C":42, "D":3, "E":18, "F":42}.items():
    ws2.column_dimensions[col].width = w
ws2.merge_cells("B2:F2")
cell(ws2, "B2", "テスト環境・前提条件", F(16, True, WHITE), NAVY, "left", bd=False)
ws2.row_dimensions[2].height = 30

env_groups = [
    ("実行環境", [
        ("GAS実行環境", "Google Apps Script（V8ランタイム）"),
        ("スプレッドシート", "Google Sheets（テスト専用ファイル）"),
        ("メール", "Gmail（テスト用アカウント）"),
        ("Drive", "Google Drive（テスト用フォルダ）")]),
    ("テストツール", [
        ("ブラウザ", "Google Chrome（最新版）"),
        ("GASデバッガ", "Apps Script エディタ（実行ログ）"),
        ("スプレッドシート確認", "目視検証"),
        ("メール確認", "Gmailで受信確認")]),
    ("前提条件", [
        ("Infoシート", "テスト用イベント情報を設定済み"),
        ("受付フォーム", "Webアプリとしてデプロイ済み"),
        ("テストデータ", "申込みデータを事前準備"),
        ("クォータ", "MIN_MAIL_QUOTAを十分に確保")]),
]
r = 4
for title, items in env_groups:
    ws2.merge_cells(f"B{r}:F{r}")
    cell(ws2, f"B{r}", title, F(12, True, WHITE), BLUE, "left")
    r += 1
    for k, v in items:
        cell(ws2, f"B{r}", k, F(10, True, NAVY), LBLUE, "left", wrap=True)
        ws2.merge_cells(f"C{r}:F{r}")
        cell(ws2, f"C{r}", v, F(10), LLBLUE, "left", wrap=True)
        r += 1
    r += 1

# ════════════════════════════════════════════════════════
# Sheet 3: テスト項目一覧
# ════════════════════════════════════════════════════════
ws3 = wb.create_sheet("テスト項目一覧")
ws3.sheet_view.showGridLines = False
for col, w in {"A":3, "B":14, "C":40, "D":10, "E":12}.items():
    ws3.column_dimensions[col].width = w
ws3.merge_cells("B2:E2")
cell(ws3, "B2", "テスト項目一覧（全体マップ）", F(16, True, WHITE), NAVY, "left", bd=False)
ws3.row_dimensions[2].height = 30
ws3.merge_cells("B3:E3")
cell(ws3, "B3", "合計テストケース数：43件　　優先度：高 23件 / 中 14件 / 低 6件", F(10, True, WHITE), BLUE, "left")

hdr = ["グループID", "テスト内容", "件数", "シート"]
for j, h in enumerate(hdr):
    cell(ws3, f"{get_column_letter(2+j)}4", h, F(10, True, WHITE), NAVY, "center")
groups = [
    ("T-01〜07", "フォーム表示・入力バリデーション", 7, "G1"),
    ("T-08〜12", "申込みデータ登録・受付番号", 5, "G2"),
    ("T-13〜18", "メール自動送信（B〜E区分）", 6, "G3"),
    ("T-19〜23", "受付完了チェック（S・A区分）", 5, "G4"),
    ("T-24〜28", "入金確認・お礼状送信", 5, "G5"),
    ("T-29〜33", "メールクォータ管理", 5, "G6"),
    ("T-34〜38", "Infoシート設定反映", 5, "G7"),
    ("T-39〜43", "onEditトリガー安全確認", 5, "G8"),
]
gcolors = [LBLUE, LGREEN, LORANGE, LBLUE, LGREEN, LORANGE, LPURPLE, LTEAL]
r = 5
for (gid, name, cnt, sh), bc in zip(groups, gcolors):
    cell(ws3, f"B{r}", gid, F(10, True, NAVY), bc, "center")
    cell(ws3, f"C{r}", name, F(10), bc, "left", wrap=True)
    cell(ws3, f"D{r}", f"{cnt}件", F(10, True), bc, "center")
    cell(ws3, f"E{r}", sh, F(10), bc, "center")
    r += 1
cell(ws3, f"B{r}", "合計", F(10, True, WHITE), NAVY, "center")
cell(ws3, f"C{r}", "全テストケース", F(10, True, WHITE), NAVY, "left")
cell(ws3, f"D{r}", "=SUM(D5:D12)", F(10, True, WHITE), NAVY, "center")
cell(ws3, f"E{r}", "", F(10, True, WHITE), NAVY, "center")

# ════════════════════════════════════════════════════════
# テストケース（グループ別シート）
# ════════════════════════════════════════════════════════
CASE_HDR = ["No.", "テスト項目", "操作手順", "期待結果", "判定", "実施日", "実施者", "備考"]
CASE_W   = [8, 22, 50, 55, 9, 13, 12, 24]

def add_case_sheet(sheet_name, title, accent, cases):
    w = wb.create_sheet(sheet_name)
    w.sheet_view.showGridLines = False
    for j, width in enumerate(CASE_W):
        w.column_dimensions[get_column_letter(1+j)].width = width
    w.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(CASE_HDR))
    cell(w, "A1", title, F(15, True, WHITE), accent, "left", bd=False)
    w.row_dimensions[1].height = 28
    # header
    for j, h in enumerate(CASE_HDR):
        cell(w, f"{get_column_letter(1+j)}2", h, F(10, True, WHITE), NAVY, "center")
    w.row_dimensions[2].height = 22
    # rows
    r = 3
    for c in cases:
        cell(w, f"A{r}", c["id"], F(10, True, BLUE, "Consolas"), LLBLUE, "center")
        cell(w, f"B{r}", c["name"], F(10, True, BLACK), LLBLUE, "left", wrap=True)
        cell(w, f"C{r}", c["steps"], F(9.5), WHITE, "left", wrap=True)
        cell(w, f"D{r}", c["exp"], F(9.5), WHITE, "left", wrap=True)
        cell(w, f"E{r}", "", F(10, True), WHITE, "center")
        cell(w, f"F{r}", "", F(9.5), WHITE, "center")
        cell(w, f"G{r}", "", F(9.5), WHITE, "center")
        cell(w, f"H{r}", "", F(9.5), WHITE, "left", wrap=True)
        # 行高をステップ/期待結果の行数で概算
        lines = max(c["steps"].count("\n"), c["exp"].count("\n")) + 1
        w.row_dimensions[r].height = max(34, lines * 15)
        r += 1
    # 判定列にプルダウン（OK/NG/未/確認）
    dv = DataValidation(type="list", formula1='"OK,NG,未,確認"', allow_blank=True)
    w.add_data_validation(dv)
    dv.add(f"E3:E{r-1}")
    w.freeze_panes = "A3"
    return w

groups_cases = {
"フォーム表示・入力バリデーション": ("T-01〜07  フォーム表示・入力バリデーション", BLUE, [
    {"id":"T-01","name":"フォーム正常表示","steps":"①デプロイURLにブラウザでアクセスする\n②ページが読み込まれることを確認","exp":"協賛申込みフォームが表示される\nイベント名・区分・価格が正しく表示される"},
    {"id":"T-02","name":"INFO未設定時のエラー","steps":"①InfoシートのEVENT_NAMEを空にする\n②フォームURLにアクセスする","exp":"「設定が不足しています」等のエラー画面が表示され、申込みフォームは表示されない"},
    {"id":"T-03","name":"申込み期間外区分の非表示","steps":"①KUBUN_SA_END を過去日付に設定\n②フォームにアクセスしてS・A区分の表示を確認","exp":"S・A区分のラジオボタンが表示されない\nB〜E区分のみ表示される"},
    {"id":"T-04","name":"背景画像表示確認","steps":"①BG_IMAGE_IDに有効なDriveファイルIDを設定\n②フォームにアクセスして背景を確認","exp":"フォームの背景に指定した画像が表示される"},
    {"id":"T-05","name":"フリガナ全角カタカナ","steps":"①会社名フリガナ欄にひらがな「かわぐち」を入力\n②「確認」ボタンを押す","exp":"バリデーションエラーが表示される\n「全角カタカナで入力してください」等のメッセージ"},
    {"id":"T-06","name":"郵便番号桁数チェック","steps":"①郵便番号欄に6桁「123456」を入力\n②「確認」ボタンを押す","exp":"バリデーションエラーが表示される\n7桁の数字を要求するメッセージ"},
    {"id":"T-07","name":"利用規約スクロール制御","steps":"①利用規約エリアをスクロールせずに同意チェックを押す\n②スクロール後にチェックを試みる","exp":"スクロール前：チェックボックスが無効（クリック不可）\nスクロール後：チェックボックスが有効になる"},
]),
"申込みデータ登録": ("T-08〜12  申込みデータ登録・受付番号", GREEN, [
    {"id":"T-08","name":"正常登録","steps":"①全必須項目を正しく入力してフォーム送信\n②スプレッドシートの「協賛申込み一覧」を確認","exp":"送信したデータが最終行に追記される\n受付日時・区分・連絡先情報がすべて正しく登録される"},
    {"id":"T-09","name":"受付番号形式確認","steps":"①フォームを正常送信する\n②A列の受付番号を確認する","exp":"受付番号が「KWGC＋17桁数字」形式で生成される\nフォーム完了画面にも同じ番号が表示される"},
    {"id":"T-10","name":"郵便番号先頭ゼロ保持","steps":"①郵便番号「0123456」を入力して送信\n②スプレッドシートI列の値を確認","exp":"「0123456」として保存される（先頭ゼロが消えない）"},
    {"id":"T-11","name":"手作業シートXLOOKUP自動入力","steps":"①フォーム送信後、手作業シートのA列に受付番号を入力\n②B〜H列のデータを確認","exp":"B〜H列に会社名・住所等がXLOOKUPで自動入力される"},
    {"id":"T-12","name":"同時送信（ロック確認）","steps":"①2つのブラウザタブで同時にフォームを送信する\n②スプレッドシートの行を確認","exp":"両方の申込みが重複なく別行に登録される\nデータが上書きされたり欠損したりしない"},
]),
"メール自動送信": ("T-13〜18  メール自動送信（B〜E区分）", ORANGE, [
    {"id":"T-13","name":"B〜E区分メール自動送信","steps":"①B区分でフォームを正常送信する\n②申込者メールアドレスの受信ボックスを確認","exp":"申込み直後にメールが届く\n件名に「受付確認」または「請求書」が含まれる"},
    {"id":"T-14","name":"請求書PDF添付確認","steps":"①T-13のメールを開いてPDFを確認する","exp":"PDFが添付されている\nPDFに会社名・区分・金額（税込）・インボイス番号・振込先が含まれる"},
    {"id":"T-15","name":"CC送信確認","steps":"①T-13のメール送信後、事務局メールを確認\n②CCで届いているか確認する","exp":"OFFICE_EMAILにCCでメールが届いている"},
    {"id":"T-16","name":"S・A区分メール不送信","steps":"①S区分でフォームを正常送信する\n②申込者メールアドレスを確認する","exp":"フォーム送信直後にメールが届かない\n（S・A区分は手動送信のため）"},
    {"id":"T-17","name":"PDF記載内容全項目確認","steps":"①B区分で申込み送信してPDFを取得する\n②PDF内の各項目を確認する","exp":"・会社名・担当者名が正しい\n・協賛区分と金額（税抜・税込）が正しい\n・消費税10%計算が正しい\n・インボイス登録番号（T+13桁）が表示\n・振込先銀行情報が正しい\n・入金期限が正しい"},
    {"id":"T-18","name":"不正メールアドレス","steps":"①メールアドレス欄に「test@」（不完全）を入力\n②送信ボタンを押す","exp":"フォームのバリデーションエラーが表示される\nサーバーへの送信が行われない"},
]),
"受付完了チェック": ("T-19〜23  受付完了チェック（S・A区分）", BLUE, [
    {"id":"T-19","name":"受付完了ダイアログ表示","steps":"①手作業シートでS区分の行のI列チェックボックスをONにする","exp":"確認ダイアログが表示される\nダイアログに会社名・区分・金額が表示されている"},
    {"id":"T-20","name":"送信確認後メール送信","steps":"①T-19のダイアログで「送信する」をクリックする\n②申込者メールを確認する","exp":"申込者に請求書PDF添付メールが届く\nJ列（請求書送信日時）に現在日時が自動記録される"},
    {"id":"T-21","name":"キャンセル時I列リセット","steps":"①I列をONにしてダイアログを表示する\n②「キャンセル」をクリックする","exp":"I列のチェックが FALSE に戻る\nメールは送信されない"},
    {"id":"T-22","name":"二重送信防止確認","steps":"①J列に日時が入力済みの行のI列を再度ONにする","exp":"「送信済み」の警告アラートが表示される\nダイアログは表示されず、二重送信されない"},
    {"id":"T-23","name":"I列OFF時連鎖クリア確認","steps":"①J列・K列・L列にデータがある行のI列をOFFにする","exp":"「J列以降をクリアしますか？」の確認が表示される\nOKを選択するとJ・K・L列がクリアされる"},
]),
"入金確認・お礼状送信": ("T-24〜28  入金確認・お礼状送信", GREEN, [
    {"id":"T-24","name":"J列空のままK列チェック","steps":"①J列（請求書送信日時）が空の行のK列をONにする","exp":"「請求書を先に送信してください」等の警告が表示される\nK列のチェックが FALSE にリセットされる"},
    {"id":"T-25","name":"入金完了ダイアログ表示","steps":"①J列に日時がある行のK列チェックボックスをONにする","exp":"お礼状送信確認ダイアログが表示される\n会社名・区分が正しく表示されている"},
    {"id":"T-26","name":"お礼状メール送信確認","steps":"①T-25のダイアログで「送信する」をクリックする\n②申込者メールを確認する","exp":"お礼状PDF添付メールが届く\nL列（お礼状送信日時）に現在日時が自動記録される"},
    {"id":"T-27","name":"お礼状PDF内容確認","steps":"①T-26で受信したお礼状PDFを開く","exp":"会社名・代表者名・区分・イベント名が正しい\n主催団体名・代表者名・日付が正しい\n組織印影が表示されている"},
    {"id":"T-28","name":"K列OFF時L列クリア確認","steps":"①L列にデータがある行のK列をOFFにする","exp":"「L列をクリアしますか？」の確認が表示される\nOKを選択するとL列がクリアされる"},
]),
"メールクォータ管理": ("T-29〜33  メールクォータ管理", ORANGE, [
    {"id":"T-29","name":"通常クォータ時の動作確認","steps":"①MIN_MAIL_QUOTAを5に設定（デフォルト）\n②残クォータが十分ある状態でフォームにアクセス","exp":"フォームが正常に表示される\nクォータ警告は表示されない"},
    {"id":"T-30","name":"低クォータ警告メール","steps":"①MIN_MAIL_QUOTAを実際の残クォータより大きく設定\n②フォームを送信する","exp":"OFFICE_EMAILに「クォータ残量警告」メールが届く\nフォームは引き続き受付を行う"},
    {"id":"T-31","name":"クォータ0時フォーム停止","steps":"①getRemainingDailyQuotaが0を返す状態を模擬する\n※Script Propertiesで制御するか翌日テスト","exp":"フォームに「本日の受付を停止しました」等が表示される\n申込みボタンが非活性または画面が出ない"},
    {"id":"T-32","name":"MIN_MAIL_QUOTA変更反映","steps":"①InfoのMIN_MAIL_QUOTAを10に変更\n②「設定を同期」を実行\n③動作を確認する","exp":"新しい閾値（10）が反映される\n残クォータが10以下になると警告メールが送信される"},
    {"id":"T-33","name":"クォータ翌日リセット確認","steps":"①前日クォータ消費後、翌日にフォームにアクセスする","exp":"フォームが正常に表示される\nクォータが自動リセットされている"},
]),
"Infoシート設定反映": ("T-34〜38  Infoシート設定反映", PURPLE, [
    {"id":"T-34","name":"価格変更反映確認","steps":"①InfoのPRICE_Bを変更して「設定を同期」を実行\n②フォームでB区分の価格を確認\n③送信して請求書PDFの金額を確認","exp":"フォーム上の表示価格が変更後の値になっている\n請求書PDFの金額も変更後の値になっている"},
    {"id":"T-35","name":"銀行情報変更反映確認","steps":"①BANK_NAMEを変更して「設定を同期」を実行\n②フォームを送信して請求書PDFを確認する","exp":"請求書PDFの振込先に変更後の銀行情報が表示される"},
    {"id":"T-36","name":"イベント名変更反映確認","steps":"①EVENT_NAMEを変更して「設定を同期」を実行\n②フォームにアクセスしてタイトルを確認","exp":"フォームのタイトルとヘッダーに変更後のイベント名が表示される"},
    {"id":"T-37","name":"INVOICE_REG_NO反映確認","steps":"①INVOICE_REG_NOに「T1234567890123」を設定\n②フォームを送信して請求書PDFを確認","exp":"請求書PDFにインボイス登録番号「T1234567890123」が表示される"},
    {"id":"T-38","name":"申込み期間設定反映","steps":"①KUBUN_SA_STARTを明日の日付に設定する\n②今日フォームにアクセスしてS区分の表示を確認","exp":"S・A区分が表示されない（申込み期間外のため）\nB〜E区分のみ表示される"},
]),
"onEditトリガー安全確認": ("T-39〜43  onEditトリガー安全確認", TEAL, [
    {"id":"T-39","name":"手作業以外シートの編集","steps":"①「協賛申込み一覧」シートのセルを編集する\n②ダイアログが表示されないことを確認","exp":"ダイアログが表示されない\nスクリプトが静かにreturnする"},
    {"id":"T-40","name":"存在しない受付番号入力","steps":"①手作業シートのA列に存在しない受付番号を入力する\n②B〜H列の値を確認する","exp":"B〜H列に「#N/A」エラーが表示される\nクラッシュは発生しない"},
    {"id":"T-41","name":"I列ON→即OFF","steps":"①I列をONにしてダイアログを表示\n②「キャンセル」をクリックしてI列の状態を確認","exp":"I列が FALSE（未チェック）に戻っている\nJ列は変化していない"},
    {"id":"T-42","name":"トリガー未登録時の動作","steps":"①トリガー一覧からonEditInstallableを削除する\n②I列チェックボックスをONにする","exp":"ダイアログが表示されない\n（ProjectInitialize再実行でトリガーが再登録される）"},
    {"id":"T-43","name":"複数行同時チェック確認","steps":"①手作業シートで複数行のI列を素早く連続してONにする\n②各行のダイアログと送信状況を確認","exp":"各行ごとにダイアログが表示・処理される\n処理が競合してクラッシュしない"},
]),
}

for name, (title, accent, cases) in groups_cases.items():
    add_case_sheet(name, title, accent, cases)

# ════════════════════════════════════════════════════════
# 結果サマリー
# ════════════════════════════════════════════════════════
wsum = wb.create_sheet("結果サマリー")
wsum.sheet_view.showGridLines = False
for col, w in {"A":3, "B":16, "C":34, "D":8, "E":8, "F":8, "G":8, "H":22}.items():
    wsum.column_dimensions[col].width = w
wsum.merge_cells("B2:H2")
cell(wsum, "B2", "テスト結果サマリー（記録用）", F(16, True, WHITE), NAVY, "left", bd=False)
wsum.row_dimensions[2].height = 30
cell(wsum, "B4", "テスト実施者：", F(10, True, BLACK), bd=False)
cell(wsum, "C4", "", F(10), bd=False)
cell(wsum, "E4", "実施期間：", F(10, True, BLACK), bd=False, align="right")
cell(wsum, "F4", "", F(10), bd=False)

# グループ別記録テーブル
hdr = ["グループ", "テスト内容", "総数", "OK", "NG", "未/確認", "備考"]
for j, h in enumerate(hdr):
    cell(wsum, f"{get_column_letter(2+j)}6", h, F(10, True, WHITE), NAVY, "center")
rows = [
    ("T-01〜07", "フォーム表示・入力バリデーション", 7),
    ("T-08〜12", "申込みデータ登録・受付番号", 5),
    ("T-13〜18", "メール自動送信（B〜E区分）", 6),
    ("T-19〜23", "受付完了チェック（S・A区分）", 5),
    ("T-24〜28", "入金確認・お礼状送信", 5),
    ("T-29〜33", "メールクォータ管理", 5),
    ("T-34〜38", "Infoシート設定反映", 5),
    ("T-39〜43", "onEディトトリガー安全確認".replace("ディ","ェディ").replace("ェディ","ェ"), 5),
]
rows[-1] = ("T-39〜43", "onEditトリガー安全確認", 5)
r = 7
for gid, nm, cnt in rows:
    cell(wsum, f"B{r}", gid, F(10, True, NAVY), LLBLUE, "center")
    cell(wsum, f"C{r}", nm, F(10), LLBLUE, "left", wrap=True)
    cell(wsum, f"D{r}", cnt, F(10, True), LLBLUE, "center")
    cell(wsum, f"E{r}", "", F(10, True, GREEN), WHITE, "center")
    cell(wsum, f"F{r}", "", F(10, True, "C0392B"), WHITE, "center")
    cell(wsum, f"G{r}", "", F(10), WHITE, "center")
    cell(wsum, f"H{r}", "", F(9.5), WHITE, "left", wrap=True)
    r += 1
# 合計行
cell(wsum, f"B{r}", "合計", F(10, True, WHITE), NAVY, "center")
cell(wsum, f"C{r}", "全テストケース", F(10, True, WHITE), NAVY, "left")
cell(wsum, f"D{r}", f"=SUM(D7:D{r-1})", F(10, True, WHITE), NAVY, "center")
cell(wsum, f"E{r}", f"=SUM(E7:E{r-1})", F(10, True, WHITE), NAVY, "center")
cell(wsum, f"F{r}", f"=SUM(F7:F{r-1})", F(10, True, WHITE), NAVY, "center")
cell(wsum, f"G{r}", f"=SUM(G7:G{r-1})", F(10, True, WHITE), NAVY, "center")
cell(wsum, f"H{r}", "", F(10, True, WHITE), NAVY, "center")
total_row = r

# 集計ボックス
r += 2
cell(wsum, f"B{r}", "テスト総件数", F(10, True, WHITE), NAVY, "center")
cell(wsum, f"C{r}", f"=D{total_row}", F(12, True, NAVY), LGRAY, "center")
r += 1
cell(wsum, f"B{r}", "合格（OK）", F(10, True, WHITE), GREEN, "center")
cell(wsum, f"C{r}", f"=E{total_row}", F(12, True, GREEN), LGREEN, "center")
r += 1
cell(wsum, f"B{r}", "不合格（NG）", F(10, True, WHITE), "C0392B", "center")
cell(wsum, f"C{r}", f"=F{total_row}", F(12, True, "C0392B"), "FCE4E4", "center")
r += 1
cell(wsum, f"B{r}", "合格率", F(10, True, WHITE), BLUE, "center")
cell(wsum, f"C{r}", f'=IF(D{total_row}=0,"-",E{total_row}/D{total_row})', F(12, True, BLUE), LBLUE, "center")
wsum[f"C{r}"].number_format = "0.0%"

# 完了基準
r += 2
wsum.merge_cells(f"B{r}:H{r}")
cell(wsum, f"B{r}", "■ 完了基準・注意事項", F(12, True, WHITE), BLUE, "left")
notes = [
    "完了基準：全43テストケースが「OK」判定であること。NGが0件であること。",
    "注意：T-13〜28のメール送信テストは実際にメールを送信し、Gmailクォータ（1日50通）を消費する。テスト用アカウントを使用すること。",
    "クリーンアップ：テスト完了後、テスト用スプレッドシートのデータを削除し、本番環境に影響がないことを確認する。",
]
for n in notes:
    r += 1
    wsum.merge_cells(f"B{r}:H{r}")
    cell(wsum, f"B{r}", "• " + n, F(9.5), LLBLUE, "left", wrap=True)
    wsum.row_dimensions[r].height = 30

wb.save("テスト仕様書_WebHanabi.xlsx")
print("✓ テスト仕様書_WebHanabi.xlsx generated — sheets:", wb.sheetnames)

"""
WebHanabi システム — 操作マニュアル PDF 生成スクリプト
ユーザー向けの操作手順書（図解入り）
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether, Image as RLImage,
    Flowable
)
from reportlab.graphics.shapes import (
    Drawing, Rect, String, Line, Circle, Polygon, Group
)
from reportlab.graphics import renderPDF, renderSVG
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from datetime import date

# ── Fonts ──────────────────────────────────────────────────────────────────
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
JP = 'HeiseiMin-W3'
JP_B = 'HeiseiKakuGo-W5'

PAGE_W, PAGE_H = A4
MARGIN = 18 * mm
CONTENT_W = PAGE_W - 2 * MARGIN

# ── Colors ─────────────────────────────────────────────────────────────────
C_NAVY   = colors.HexColor('#1a3a5c')
C_BLUE   = colors.HexColor('#2e6da4')
C_LBLUE  = colors.HexColor('#dce9f5')
C_LLBLUE = colors.HexColor('#f0f5fb')
C_GREEN  = colors.HexColor('#27ae60')
C_LGREEN = colors.HexColor('#d5f0e0')
C_ORANGE = colors.HexColor('#e67e22')
C_LORANGE= colors.HexColor('#fde9d4')
C_RED    = colors.HexColor('#c0392b')
C_LRED   = colors.HexColor('#fce4e4')
C_GRAY   = colors.HexColor('#7f8c8d')
C_LGRAY  = colors.HexColor('#ecf0f1')
C_WHITE  = colors.white
C_BLACK  = colors.HexColor('#222222')
C_YELLOW = colors.HexColor('#f1c40f')
C_LYELLOW= colors.HexColor('#fef9e7')


# ══════════════════════════════════════════════════════════════════════════
#  Styles
# ══════════════════════════════════════════════════════════════════════════
def make_styles():
    s = {}
    def ps(name, **kw):
        return ParagraphStyle(name, **kw)

    s['cover_main']  = ps('cover_main',  fontName=JP_B, fontSize=32, leading=46,
                           textColor=C_WHITE, alignment=TA_CENTER)
    s['cover_sub']   = ps('cover_sub',   fontName=JP_B, fontSize=16, leading=26,
                           textColor=colors.HexColor('#cce0f5'), alignment=TA_CENTER)
    s['cover_meta']  = ps('cover_meta',  fontName=JP,   fontSize=11, leading=18,
                           textColor=C_LGRAY, alignment=TA_CENTER)

    s['h1']  = ps('h1',  fontName=JP_B, fontSize=17, leading=26,
                  textColor=C_NAVY, spaceBefore=16, spaceAfter=6)
    s['h2']  = ps('h2',  fontName=JP_B, fontSize=13, leading=20,
                  textColor=C_BLUE, spaceBefore=12, spaceAfter=4)
    s['h3']  = ps('h3',  fontName=JP_B, fontSize=11, leading=18,
                  textColor=C_NAVY, spaceBefore=8,  spaceAfter=3)
    s['body'] = ps('body', fontName=JP,  fontSize=10, leading=17,
                   textColor=C_BLACK, spaceAfter=4)
    s['bodyC']= ps('bodyC',fontName=JP,  fontSize=10, leading=17,
                   textColor=C_BLACK, alignment=TA_CENTER)
    s['step'] = ps('step', fontName=JP,  fontSize=10, leading=17,
                   leftIndent=18, textColor=C_BLACK, spaceAfter=3)
    s['note'] = ps('note', fontName=JP,  fontSize=9,  leading=15,
                   textColor=colors.HexColor('#555555'),
                   backColor=C_LYELLOW, leftIndent=8, rightIndent=8,
                   borderPad=6, spaceAfter=6)
    s['warn'] = ps('warn', fontName=JP,  fontSize=9,  leading=15,
                   textColor=colors.HexColor('#7d2c00'),
                   backColor=C_LORANGE, leftIndent=8, rightIndent=8,
                   borderPad=6, spaceAfter=6)
    s['tip']  = ps('tip',  fontName=JP,  fontSize=9,  leading=15,
                   textColor=colors.HexColor('#145a32'),
                   backColor=C_LGREEN, leftIndent=8, rightIndent=8,
                   borderPad=6, spaceAfter=6)
    s['th']   = ps('th',   fontName=JP_B, fontSize=9, leading=13,
                   textColor=C_WHITE, alignment=TA_CENTER)
    s['td']   = ps('td',   fontName=JP,   fontSize=9, leading=13,
                   textColor=C_BLACK)
    s['tdC']  = ps('tdC',  fontName=JP,   fontSize=9, leading=13,
                   textColor=C_BLACK, alignment=TA_CENTER)
    s['caption'] = ps('caption', fontName=JP_B, fontSize=9, leading=14,
                       textColor=C_GRAY, alignment=TA_CENTER, spaceAfter=8)
    return s


def tbl(data, col_w, has_header=True, row_colors=None):
    t = Table(data, colWidths=col_w)
    cmds = [
        ('FONTNAME',      (0,0), (-1,-1), JP),
        ('FONTSIZE',      (0,0), (-1,-1), 9),
        ('GRID',          (0,0), (-1,-1), 0.5, colors.HexColor('#bbbbbb')),
        ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',    (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING',   (0,0), (-1,-1), 6),
        ('RIGHTPADDING',  (0,0), (-1,-1), 6),
        ('ROWBACKGROUNDS',(0,1), (-1,-1), [C_WHITE, C_LLBLUE]),
    ]
    if has_header:
        cmds += [
            ('BACKGROUND',  (0,0), (-1,0), C_NAVY),
            ('FONTNAME',    (0,0), (-1,0), JP_B),
            ('TEXTCOLOR',   (0,0), (-1,0), C_WHITE),
            ('ALIGN',       (0,0), (-1,0), 'CENTER'),
        ]
    t.setStyle(TableStyle(cmds))
    return t


def add_header_footer(canvas, doc):
    canvas.saveState()
    pg = canvas.getPageNumber()
    # Footer line
    canvas.setStrokeColor(C_LBLUE)
    canvas.setLineWidth(0.8)
    canvas.line(MARGIN, 14*mm, PAGE_W-MARGIN, 14*mm)
    canvas.setFont(JP, 8)
    canvas.setFillColor(C_GRAY)
    canvas.drawString(MARGIN, 10*mm, "WebHanabi — 川口花火大会 協賛管理システム 操作マニュアル")
    canvas.drawRightString(PAGE_W-MARGIN, 10*mm, f"{pg}")
    canvas.restoreState()


# ══════════════════════════════════════════════════════════════════════════
#  Drawing helpers
# ══════════════════════════════════════════════════════════════════════════

class DrawingFlowable(Flowable):
    """Wrap a reportlab Drawing so it flows in the story."""
    def __init__(self, drawing, caption=None):
        super().__init__()
        self.drawing = drawing
        self.caption = caption
        self.width  = drawing.width
        self.height = drawing.height + (6 if caption else 0)

    def draw(self):
        self.drawing.drawOn(self.canv, 0, 6 if self.caption else 0)
        if self.caption:
            self.canv.setFont(JP_B, 8)
            self.canv.setFillColor(C_GRAY)
            self.canv.drawCentredString(self.width/2, 0, self.caption)


def txt(s, x, y, size=9, color=C_BLACK, font=JP, anchor='start'):
    from reportlab.graphics.shapes import String
    st = String(x, y, s, fontName=font, fontSize=size, fillColor=color)
    if anchor == 'middle':
        st.textAnchor = 'middle'
    elif anchor == 'end':
        st.textAnchor = 'end'
    return st


def box(x, y, w, h, fill=C_LBLUE, stroke=C_BLUE, r=4):
    return RoundRect(x, y, w, h, r, fillColor=fill, strokeColor=stroke, strokeWidth=1)


def arrow_down(x, y, length=16, color=C_GRAY):
    """Draw a downward arrow."""
    hw = 5
    g = Group()
    g.add(Line(x, y, x, y-length+hw, strokeColor=color, strokeWidth=1.5))
    g.add(Polygon([x-hw, y-length+hw, x+hw, y-length+hw, x, y-length],
                  fillColor=color, strokeColor=color, strokeWidth=0.5))
    return g


def arrow_right(x, y, length=20, color=C_GRAY):
    hw = 4
    g = Group()
    g.add(Line(x, y, x+length-hw, y, strokeColor=color, strokeWidth=1.5))
    g.add(Polygon([x+length-hw, y+hw, x+length-hw, y-hw, x+length, y],
                  fillColor=color, strokeColor=color, strokeWidth=0.5))
    return g


def step_badge(x, y, n, color=C_BLUE):
    g = Group()
    g.add(Circle(x, y, 9, fillColor=color, strokeColor=color))
    g.add(txt(str(n), x, y-4, size=9, color=C_WHITE, font=JP_B, anchor='middle'))
    return g


# ══════════════════════════════════════════════════════════════════════════
#  Diagram: System overview flow
# ══════════════════════════════════════════════════════════════════════════
def draw_system_overview():
    W, H = CONTENT_W, 130
    d = Drawing(W, H)

    bw, bh, gap = 80, 36, 14
    nodes = [
        ("申込者", "フォーム入力・送信", C_LBLUE, C_BLUE),
        ("Webフォーム\n(GAS)", "申込み受付\n受付番号生成", C_LGREEN, C_GREEN),
        ("スプレッド\nシート", "データ保存\nXLOOKUP管理", C_LORANGE, C_ORANGE),
        ("Gmail", "PDF添付\nメール送信", C_LRED, C_RED),
        ("担当者", "ワークフロー\n管理", C_LYELLOW, C_YELLOW),
    ]
    total_w = len(nodes)*bw + (len(nodes)-1)*gap
    start_x = (W - total_w) / 2
    cy = H/2

    for i, (title, sub, fill, stroke) in enumerate(nodes):
        x = start_x + i*(bw+gap)
        d.add(Rect(x, cy-bh/2, bw, bh, rx=6, ry=6,
                        fillColor=fill, strokeColor=stroke, strokeWidth=1.5))
        lines = title.split('\n')
        d.add(txt(lines[0], x+bw/2, cy+4, size=9, font=JP_B, color=C_BLACK, anchor='middle'))
        if len(lines)>1:
            d.add(txt(lines[1], x+bw/2, cy-6, size=7, font=JP_B, color=stroke, anchor='middle'))
        slines = sub.split('\n')
        d.add(txt(slines[0], x+bw/2, cy-bh/2-10, size=7, color=C_GRAY, font=JP, anchor='middle'))
        if len(slines)>1:
            d.add(txt(slines[1], x+bw/2, cy-bh/2-19, size=7, color=C_GRAY, font=JP, anchor='middle'))
        if i < len(nodes)-1:
            ax = x + bw + 2
            d.add(arrow_right(ax, cy, length=gap-4))

    return d


# ══════════════════════════════════════════════════════════════════════════
#  Diagram: Application workflow
# ══════════════════════════════════════════════════════════════════════════
def draw_application_flow():
    W, H = CONTENT_W, 220
    d = Drawing(W, H)

    bw, bh = 140, 34
    cx = W / 2

    steps_left = [
        (cx-10, H-30, "① フォームにアクセス", "公開URL を開く", C_LBLUE, C_BLUE),
        (cx-10, H-100, "② 情報を入力", "会社情報・区分を選択", C_LGREEN, C_GREEN),
        (cx-10, H-170, "③ 確認画面で内容確認", "「確認」ボタンを押す", C_LORANGE, C_ORANGE),
    ]
    steps_right = [
        (cx+10+bw+16, H-30,  "⑥ 受付完了", "受付番号が画面に表示", C_LGREEN, C_GREEN),
        (cx+10+bw+16, H-100, "⑤ メール受信", "請求書PDF が届く（B〜E）", C_LRED, C_RED),
        (cx+10+bw+16, H-170, "④ 送信", "「送信」ボタンを押す", C_LYELLOW, C_YELLOW),
    ]

    # Left column
    for x, y, title, sub, fill, stroke in steps_left:
        rx = cx - bw - 8
        d.add(Rect(rx, y-bh/2, bw, bh, rx=5, ry=5, fillColor=fill, strokeColor=stroke, strokeWidth=1.5))
        d.add(txt(title, rx+bw/2, y+4, size=9, font=JP_B, color=C_BLACK, anchor='middle'))
        d.add(txt(sub,   rx+bw/2, y-7, size=8, font=JP,   color=stroke,  anchor='middle'))

    # Right column
    for x, y, title, sub, fill, stroke in steps_right:
        d.add(Rect(x, y-bh/2, bw, bh, rx=5, ry=5, fillColor=fill, strokeColor=stroke, strokeWidth=1.5))
        d.add(txt(title, x+bw/2, y+4, size=9, font=JP_B, color=C_BLACK, anchor='middle'))
        d.add(txt(sub,   x+bw/2, y-7, size=8, font=JP,   color=stroke,  anchor='middle'))

    # Arrows left column down
    lx = cx - bw/2 - 8
    for y_from, y_to in [(H-30-bh/2, H-100+bh/2+2), (H-100-bh/2, H-170+bh/2+2)]:
        d.add(arrow_down(lx, y_from, length=y_from-y_to))

    # Bottom connector: left → right
    bx_l = cx - 8
    bx_r = cx + 10 + bw + 16
    by = H - 170
    d.add(Line(bx_l-bw/2, by, cx+bw/2+26, by,
               strokeColor=C_ORANGE, strokeWidth=1.5, strokeDashArray=[3,2]))
    d.add(txt("送信", cx+bw/2+12, by+3, size=8, color=C_ORANGE, font=JP_B))

    # Arrows right column up
    rx = cx + 10 + bw + 16 + bw/2
    for y_from, y_to in [(H-170+bh/2+2, H-100-bh/2), (H-100+bh/2+2, H-30-bh/2)]:
        d.add(arrow_down(rx, y_to, length=y_to-y_from))

    return d


# ══════════════════════════════════════════════════════════════════════════
#  Diagram: Form UI mockup
# ══════════════════════════════════════════════════════════════════════════
def draw_form_mockup():
    W, H = CONTENT_W, 260
    d = Drawing(W, H)

    # Browser chrome
    d.add(Rect(0, 0, W, H, rx=6, ry=6, fillColor=C_LGRAY, strokeColor=C_GRAY, strokeWidth=1))
    d.add(Rect(0, H-22, W, 22, rx=6, ry=6, fillColor=colors.HexColor('#e8e8e8'),
                    strokeColor=C_GRAY, strokeWidth=1))
    d.add(txt("川口花火大会 協賛申込みフォーム", 40, H-14, size=8, color=C_GRAY, font=JP))
    # close buttons
    for i, c in enumerate([C_RED, C_ORANGE, C_GREEN]):
        d.add(Circle(10+i*14, H-11, 5, fillColor=c, strokeColor=c))

    # Form area (white bg)
    form_y = 10
    form_h = H - 40
    d.add(Rect(10, form_y, W-20, form_h, fillColor=C_WHITE,
               strokeColor=C_LBLUE, strokeWidth=1))

    # Header bar
    d.add(Rect(10, form_y+form_h-40, W-20, 40,
               fillColor=C_NAVY, strokeColor=C_NAVY))
    d.add(txt("川口花火大会 協賛申込みフォーム", W/2, form_y+form_h-16,
              size=12, font=JP_B, color=C_WHITE, anchor='middle'))

    # Form fields
    fields = [
        ("会社名・団体名", 60),
        ("会社名（フリガナ）", 45),
        ("代表者役職・代表者名", 30),
        ("担当者名", 15),
    ]
    for label, offset in fields:
        fy = form_y + form_h - 60 - (60 - offset)
        d.add(txt(label+"  ＊", 20, fy, size=8, color=C_NAVY, font=JP_B))
        d.add(Rect(20, fy-14, W-50, 12, rx=2, ry=2,
                        fillColor=C_WHITE, strokeColor=C_LBLUE, strokeWidth=0.8))

    # Kubun select
    ky = form_y + form_h - 160
    d.add(txt("協賛区分  ＊", 20, ky, size=8, color=C_NAVY, font=JP_B))
    kubun_labels = ["S区分", "A区分", "B区分", "C区分", "D区分", "E区分"]
    bkw = (W-40) / 6
    for i, lb in enumerate(kubun_labels):
        bkx = 20 + i*bkw
        fc = C_LBLUE if i < 2 else C_LGREEN
        sc = C_BLUE  if i < 2 else C_GREEN
        d.add(Rect(bkx+1, ky-18, bkw-4, 16, rx=3, ry=3, fillColor=fc, strokeColor=sc, strokeWidth=1))
        d.add(txt(lb, bkx+bkw/2, ky-12, size=7, color=sc, font=JP_B, anchor='middle'))

    # Terms checkbox area
    ty = form_y + form_h - 195
    d.add(Rect(20, ty-30, W-40, 30, rx=2, ry=2,
                    fillColor=C_LLBLUE, strokeColor=C_LBLUE, strokeWidth=0.8))
    d.add(txt("利用規約  （下部までスクロールすると同意チェックが有効になります）",
              30, ty-12, size=7, color=C_GRAY, font=JP))

    # Submit button
    d.add(Rect(W/2-40, form_y+14, 80, 22, rx=4, ry=4,
                    fillColor=C_BLUE, strokeColor=C_BLUE))
    d.add(txt("確認する", W/2, form_y+29, size=10, font=JP_B, color=C_WHITE, anchor='middle'))

    return d


# ══════════════════════════════════════════════════════════════════════════
#  Diagram: Spreadsheet mockup (手作業シート)
# ══════════════════════════════════════════════════════════════════════════
def draw_sheet_mockup():
    W, H = CONTENT_W, 130
    d = Drawing(W, H)

    cols = ["A\n受付番号", "B\n区分", "C\n電話番号", "D\n会社名", "E\n住所",
            "I\n受付完了", "J\n請求書\n送信日時", "K\n入金完了", "L\nお礼状\n送信日時"]
    cw_list = [38, 15, 28, 36, 40, 16, 22, 16, 22]  # mm equivalent in pts
    total_w = sum(cw_list)
    scale = (W - 20) / total_w
    cw_scaled = [c * scale for c in cw_list]

    row_h = 18
    header_h = 28
    rows = [
        ["KWGC001", "B", "090-1234", "株式会社ABC", "東京都...", "☑", "2025/06/18", "☑", "2025/06/20"],
        ["KWGC002", "S", "03-5678", "山田商会",    "埼玉県...", "☑", "2025/06/19", "", ""],
        ["KWGC003", "A", "048-9012", "花火製作所",  "川口市...", "", "",            "", ""],
    ]
    x0 = 10
    y0 = H - header_h - 4

    # Draw header
    cx = x0
    for i, col in enumerate(cols):
        d.add(Rect(cx, y0, cw_scaled[i], header_h,
                   fillColor=C_NAVY, strokeColor=C_WHITE, strokeWidth=0.5))
        lines = col.split('\n')
        for li, line in enumerate(lines):
            d.add(txt(line, cx+cw_scaled[i]/2, y0+header_h-10-li*10,
                      size=7, font=JP_B, color=C_WHITE, anchor='middle'))
        cx += cw_scaled[i]

    # Draw rows
    for ri, row in enumerate(rows):
        cx = x0
        row_y = y0 - (ri+1)*row_h
        bg = C_WHITE if ri%2==0 else C_LLBLUE
        for ci, cell in enumerate(row):
            fill = bg
            if ci == 5 and cell == "☑":   fill = C_LGREEN
            if ci == 7 and cell == "☑":   fill = C_LGREEN
            if ci == 5 and cell == "":     fill = C_LRED
            d.add(Rect(cx, row_y, cw_scaled[ci], row_h,
                       fillColor=fill, strokeColor=C_LBLUE, strokeWidth=0.5))
            cell_color = C_GREEN if cell == "☑" else C_BLACK
            d.add(txt(cell, cx+cw_scaled[ci]/2, row_y+5,
                      size=7, font=JP, color=cell_color, anchor='middle'))
            cx += cw_scaled[ci]

    # XLOOKUP annotation
    d.add(Line(x0 + cw_scaled[0], y0 - row_h*3,
               x0 + cw_scaled[0], y0,
               strokeColor=C_ORANGE, strokeWidth=1, strokeDashArray=[2,2]))
    d.add(txt("←手入力", x0+2, y0-row_h*3-10, size=7, color=C_ORANGE, font=JP_B))
    d.add(txt("XLOOKUP自動→", x0+cw_scaled[0]+4, y0-row_h*3-10, size=7, color=C_BLUE, font=JP_B))

    return d


# ══════════════════════════════════════════════════════════════════════════
#  Diagram: Staff workflow
# ══════════════════════════════════════════════════════════════════════════
def draw_staff_workflow():
    W, H = CONTENT_W, 180
    d = Drawing(W, H)

    steps = [
        ("申込み\n受信", C_LBLUE,   C_BLUE,   "スプレッドシートに\n自動登録"),
        ("受付番号\n入力", C_LGREEN, C_GREEN,  "手作業シートA列\nに受付番号を入力"),
        ("受付完了\nチェック", C_LORANGE, C_ORANGE, "I列チェック→\nダイアログ確認"),
        ("請求書\nメール送信", C_LRED,  C_RED,   "担当者が確認して\n送信ボタン押下"),
        ("入金確認\nチェック", C_LYELLOW, colors.HexColor('#b7950b'), "K列チェック→\nダイアログ確認"),
        ("お礼状\n送信完了", C_LGREEN, C_GREEN, "お礼状PDF添付\nメール自動送信"),
    ]

    bw = (W - 40) / len(steps) - 6
    bh = 36
    spacing = (W - 40) / len(steps)

    for i, (title, fill, stroke, note) in enumerate(steps):
        x = 20 + i * spacing
        cy = H / 2 + 10
        d.add(Rect(x, cy - bh/2, bw, bh, rx=5, ry=5,
                        fillColor=fill, strokeColor=stroke, strokeWidth=1.5))
        for li, line in enumerate(title.split('\n')):
            d.add(txt(line, x+bw/2, cy+6-li*11, size=8, font=JP_B,
                      color=C_BLACK, anchor='middle'))
        for li, line in enumerate(note.split('\n')):
            d.add(txt(line, x+bw/2, cy-bh/2-10-li*9, size=7, font=JP,
                      color=C_GRAY, anchor='middle'))
        # Step badge
        d.add(Circle(x+bw/2, cy+bh/2+10, 8, fillColor=stroke, strokeColor=stroke))
        d.add(txt(str(i+1), x+bw/2, cy+bh/2+6, size=8, font=JP_B,
                  color=C_WHITE, anchor='middle'))
        if i < len(steps)-1:
            ax = x + bw + 2
            d.add(arrow_right(ax, cy, length=spacing-bw-4))

    return d


# ══════════════════════════════════════════════════════════════════════════
#  Diagram: Invoice dialog mockup
# ══════════════════════════════════════════════════════════════════════════
def draw_dialog_mockup(title_text, body_lines, btn_texts, btn_colors):
    W, H = 240, 130
    d = Drawing(W, H)
    # Shadow
    d.add(Rect(6, 4, W-12, H-10, rx=8, ry=8, fillColor=C_LGRAY, strokeColor=C_LGRAY))
    # Dialog bg
    d.add(Rect(2, 8, W-12, H-10, rx=8, ry=8, fillColor=C_WHITE,
                    strokeColor=C_BLUE, strokeWidth=1.5))
    # Title bar
    d.add(Rect(2, H-30, W-12, 22, rx=8, ry=8, fillColor=C_NAVY, strokeColor=C_NAVY))
    d.add(Rect(2, H-22, W-12, 14, fillColor=C_NAVY, strokeColor=C_NAVY))
    d.add(txt(title_text, (W-10)/2+2, H-21, size=10, font=JP_B, color=C_WHITE, anchor='middle'))
    # Body
    for i, line in enumerate(body_lines):
        d.add(txt(line, 20, H-52-i*14, size=9, font=JP, color=C_BLACK))
    # Buttons
    bw2 = 70
    total_bw = len(btn_texts)*bw2 + (len(btn_texts)-1)*10
    bx0 = ((W-10) - total_bw) / 2 + 2
    for i, (btn_txt, btn_color) in enumerate(zip(btn_texts, btn_colors)):
        bx = bx0 + i*(bw2+10)
        d.add(Rect(bx, 18, bw2, 20, rx=4, ry=4, fillColor=btn_color, strokeColor=btn_color))
        d.add(txt(btn_txt, bx+bw2/2, 31, size=9, font=JP_B, color=C_WHITE, anchor='middle'))
    return d


# ══════════════════════════════════════════════════════════════════════════
#  Diagram: Email flow
# ══════════════════════════════════════════════════════════════════════════
def draw_email_flow():
    W, H = CONTENT_W, 110
    d = Drawing(W, H)

    items = [
        ("B〜E区分\n申込み直後", C_LGREEN, C_GREEN,
         "申込み確認メール\n＋請求書PDF"),
        ("S・A区分\n受付完了チェック後", C_LORANGE, C_ORANGE,
         "当選通知メール\n＋請求書PDF"),
        ("全区分\n入金完了チェック後", C_LRED, C_RED,
         "お礼状メール\n＋お礼状PDF"),
    ]

    bw = (W - 40) / len(items) - 8
    bh = 36

    for i, (trigger, fill, stroke, content) in enumerate(items):
        x = 20 + i * ((W-40)/len(items))
        # Trigger box
        d.add(Rect(x, H-50, bw, 36, rx=5, ry=5, fillColor=fill, strokeColor=stroke, strokeWidth=1.5))
        for li, line in enumerate(trigger.split('\n')):
            d.add(txt(line, x+bw/2, H-28-li*12, size=8, font=JP_B, color=C_BLACK, anchor='middle'))
        # Arrow
        ax = x + bw/2
        d.add(arrow_down(ax, H-50, length=16))
        # Content box
        d.add(Rect(x, 8, bw, 28, rx=5, ry=5, fillColor=C_LLBLUE, strokeColor=C_BLUE, strokeWidth=1))
        for li, line in enumerate(content.split('\n')):
            d.add(txt(line, x+bw/2, 26-li*12, size=8, font=JP_B, color=C_BLUE, anchor='middle'))

    return d


# ══════════════════════════════════════════════════════════════════════════
#  Diagram: Setup flow
# ══════════════════════════════════════════════════════════════════════════
def draw_setup_flow():
    W, H = CONTENT_W, 100
    d = Drawing(W, H)

    steps = [
        ("Infoシートに\nROOT_FOLDER_ID\nを入力", C_LBLUE, C_BLUE),
        ("メニュー→\nProject\nInitialize", C_LGREEN, C_GREEN),
        ("フォルダ・\nスプレッド\nシート生成", C_LORANGE, C_ORANGE),
        ("トリガー\n自動登録", C_LYELLOW, colors.HexColor('#b7950b')),
        ("Webアプリ\nとして\nデプロイ", C_LRED, C_RED),
        ("公開URL\nを申込者\nに共有", C_LGREEN, C_GREEN),
    ]

    bw = (W - 30) / len(steps) - 4
    bh = 50

    for i, (label, fill, stroke) in enumerate(steps):
        x = 15 + i * ((W-30)/len(steps))
        d.add(Rect(x, H-bh-10, bw, bh, rx=5, ry=5, fillColor=fill, strokeColor=stroke, strokeWidth=1.5))
        for li, line in enumerate(label.split('\n')):
            d.add(txt(line, x+bw/2, H-22-li*12, size=7.5, font=JP_B, color=C_BLACK, anchor='middle'))
        d.add(Circle(x+bw/2, H-bh-10-10, 7, fillColor=stroke, strokeColor=stroke))
        d.add(txt(str(i+1), x+bw/2, H-bh-10-14, size=7, font=JP_B, color=C_WHITE, anchor='middle'))
        if i < len(steps)-1:
            ax = x + bw + 2
            d.add(arrow_right(ax, H-bh/2-10, length=((W-30)/len(steps))-bw-4))

    return d


def draw_cover_background(canvas, doc):
    import math
    canvas.saveState()
    canvas.setFillColor(C_NAVY)
    canvas.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    for i in range(20):
        alpha = 0.03 + i*0.01
        canvas.setFillColorRGB(0.18, 0.45, 0.7, alpha=alpha)
        canvas.rect(0, i*(PAGE_H/20), PAGE_W, PAGE_H/20+2, fill=1, stroke=0)
    for angle_step, radius, num, cx, cy, color_hex in [
        (30, 60, 12, PAGE_W*0.12, PAGE_H*0.78, '#4a90d9'),
        (45, 40,  8, PAGE_W*0.88, PAGE_H*0.22, '#27ae60'),
        (36, 50, 10, PAGE_W*0.82, PAGE_H*0.70, '#e67e22'),
    ]:
        r, g, b = [int(color_hex[i:i+2],16)/255 for i in (1,3,5)]
        canvas.setStrokeColorRGB(r, g, b, alpha=0.4)
        canvas.setLineWidth(1)
        for i in range(num):
            angle = math.radians(i * angle_step)
            x1 = cx + 8*math.cos(angle)
            y1 = cy + 8*math.sin(angle)
            x2 = cx + radius*math.cos(angle)
            y2 = cy + radius*math.sin(angle)
            canvas.line(x1, y1, x2, y2)
            canvas.setFillColorRGB(r, g, b, alpha=0.5)
            canvas.circle(x2, y2, 2, fill=1, stroke=0)
    canvas.setFillColor(C_BLUE)
    canvas.rect(0, 0, PAGE_W, 22*mm, fill=1, stroke=0)
    canvas.restoreState()


# ══════════════════════════════════════════════════════════════════════════
#  Build story
# ══════════════════════════════════════════════════════════════════════════
def build_story():
    s = make_styles()
    story = []

    # ── Cover ─────────────────────────────────────────────────────────────
    story.append(Spacer(1, 55*mm))
    story.append(Paragraph("川口花火大会", s['cover_sub']))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph("協賛管理システム", s['cover_main']))
    story.append(Paragraph("WebHanabi", ParagraphStyle('cv2', fontName=JP_B, fontSize=22,
                            textColor=colors.HexColor('#7fb3d8'), alignment=TA_CENTER)))
    story.append(Spacer(1, 8*mm))
    story.append(HRFlowable(width=120*mm, thickness=1.5, color=C_BLUE,
                             hAlign='CENTER', spaceAfter=8))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph("操　作　マニュアル", ParagraphStyle('cv3', fontName=JP_B,
                            fontSize=20, textColor=C_WHITE, alignment=TA_CENTER)))
    story.append(Spacer(1, 12*mm))
    story.append(Paragraph(f"発行日：{date.today().strftime('%Y年%m月%d日')}", s['cover_meta']))
    story.append(Paragraph("対象：申込者・事務局担当者", s['cover_meta']))
    story.append(PageBreak())

    # ── 目次 ──────────────────────────────────────────────────────────────
    story.append(Paragraph("目　次", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1.5, color=C_NAVY, spaceAfter=8))
    toc_data = [
        ["1.", "システム概要", "3"],
        ["2.", "初期セットアップ（管理者向け）", "4"],
        ["3.", "協賛申込みフォームの使い方（申込者向け）", "5"],
        ["4.", "申込みデータの管理（事務局向け）", "7"],
        ["5.", "請求書メールの送信（事務局向け）", "9"],
        ["6.", "入金確認とお礼状送信（事務局向け）", "11"],
        ["7.", "メールクォータの管理", "13"],
        ["8.", "よくある質問（FAQ）", "14"],
        ["9.", "トラブルシューティング", "15"],
    ]
    toc = tbl(toc_data, [12*mm, 130*mm, 13*mm], has_header=False)
    story.append(toc)
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════
    # 1. システム概要
    # ══════════════════════════════════════════════════════════════════════
    story.append(Paragraph("1. システム概要", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=C_NAVY, spaceAfter=6))
    story.append(Paragraph(
        "WebHanabi は、川口花火大会の協賛申込みから請求書発行・入金確認・お礼状送付までを"
        "自動化する管理システムです。Google Sheets と Google Apps Script を基盤とし、"
        "インターネットブラウザだけで申込み受付・業務管理が完結します。",
        s['body']
    ))
    story.append(Spacer(1, 4*mm))

    d = draw_system_overview()
    story.append(DrawingFlowable(d, caption="図1-1  システム全体構成"))
    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("1.1 利用者の種類", s['h2']))
    roles = tbl([
        ["利用者", "主な操作", "使用する画面"],
        ["申込者（一般）",  "協賛申込みフォームへの入力・送信",        "公開Webフォーム（ブラウザ）"],
        ["事務局担当者",   "受付確認・請求書送信・入金確認・お礼状送信", "Google Sheets（手作業シート）"],
        ["システム管理者", "初期設定・設定変更・Infoシート管理",        "Google Sheets（Infoシート）＋ GAS エディタ"],
    ], [32*mm, 70*mm, 53*mm])
    story.append(roles)
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════
    # 2. 初期セットアップ
    # ══════════════════════════════════════════════════════════════════════
    story.append(Paragraph("2. 初期セットアップ（管理者向け）", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=C_NAVY, spaceAfter=6))
    story.append(Paragraph(
        "初めてシステムを利用する前に、以下の手順でセットアップを行ってください。"
        "一度完了すれば、次回以降は不要です。",
        s['body']
    ))
    story.append(Spacer(1, 3*mm))
    d = draw_setup_flow()
    story.append(DrawingFlowable(d, caption="図2-1  初期セットアップの流れ"))
    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("2.1 Infoシートの設定", s['h2']))
    story.append(Paragraph("スクリプトに紐付いた Google Sheets を開き、「Info」シートに以下の情報を入力します。", s['body']))
    info_data = tbl([
        ["設定キー", "説明", "入力例"],
        ["EVENT_NAME",       "イベント名",             "第〇回 川口花火大会"],
        ["OFFICE_EMAIL",     "事務局メールアドレス",     "office@example.com"],
        ["ORG_NAME",         "主催団体名",             "川口市花火実行委員会"],
        ["ORG_REP",          "代表者名",               "川口 太郎"],
        ["START_DATE",       "開催日時（開始）",        "2025年8月3日 19:00"],
        ["PAYMENT_DUE",      "入金期限",               "2025年7月31日"],
        ["BANK_NAME",        "振込先金融機関名",         "○○銀行 ○○支店"],
        ["BANK_NO",          "口座番号",               "普通 1234567"],
        ["BANK_HOLDER",      "口座名義（カナ）",         "カワグチハナビジッコウイインカイ"],
        ["PRICE_B / PRICE_C\n/ PRICE_D / PRICE_E", "各区分の協賛金額（税抜・円）", "100000"],
        ["INVOICE_REG_NO",   "インボイス登録番号",        "T1234567890123"],
        ["ROOT_FOLDER_ID",   "保存先 Google Drive フォルダID", "1AbCdEfGh...（URLから取得）"],
    ], [45*mm, 55*mm, 55*mm])
    story.append(info_data)
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(
        "【メモ】Google Drive フォルダIDはフォルダURLの末尾の英数字部分です。\n"
        "例：https://drive.google.com/drive/folders/ 【ここがID】",
        s['note']
    ))

    story.append(Paragraph("2.2 プロジェクト初期化の実行", s['h2']))
    for i, step in enumerate([
        "Google Sheets のメニューバーから「WebHanabi」→「Project Initialize」をクリックします。",
        "確認ダイアログが表示されたら「OK」をクリックします。",
        "処理が完了すると、Google Drive に新しいフォルダとスプレッドシートが作成されます。",
        "「Info」シートの DATA_SPREADSHEET_ID に作成されたスプレッドシートのIDが自動入力されます。",
    ], 1):
        story.append(Paragraph(f"　{i}. {step}", s['step']))

    story.append(Paragraph("2.3 Webアプリのデプロイ", s['h2']))
    for i, step in enumerate([
        "Apps Script エディタを開き、右上の「デプロイ」→「新しいデプロイ」を選択します。",
        "種類を「ウェブアプリ」に設定します。",
        "「次のユーザーとして実行」を「自分」、「アクセスできるユーザー」を「全員」に設定します。",
        "「デプロイ」ボタンをクリックし、表示されたURLをコピーします。",
        "コピーしたURLを申込者に共有します。",
    ], 1):
        story.append(Paragraph(f"　{i}. {step}", s['step']))

    story.append(Paragraph(
        "【重要】設定を変更した場合は、必ず「デプロイを管理」から「新しいバージョン」として"
        "再デプロイしてください。既存のデプロイを更新しないと変更が反映されません。",
        s['warn']
    ))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════
    # 3. 申込みフォームの使い方
    # ══════════════════════════════════════════════════════════════════════
    story.append(Paragraph("3. 協賛申込みフォームの使い方（申込者向け）", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=C_NAVY, spaceAfter=6))
    story.append(Paragraph("申込みの流れを図で確認してから、詳細手順をお読みください。", s['body']))
    story.append(Spacer(1, 3*mm))

    d = draw_application_flow()
    story.append(DrawingFlowable(d, caption="図3-1  申込みの流れ（申込者）"))
    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("3.1 フォーム画面の構成", s['h2']))
    story.append(Spacer(1, 2*mm))
    d = draw_form_mockup()
    story.append(DrawingFlowable(d, caption="図3-2  申込みフォーム画面イメージ"))
    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("3.2 入力項目の説明", s['h2']))
    fields_data = tbl([
        ["項目名", "必須", "入力内容", "注意点"],
        ["会社名・団体名",         "✓", "法人・団体の正式名称",           "略称不可"],
        ["会社名（フリガナ）",      "✓", "全角カタカナで入力",             "ひらがな・漢字不可"],
        ["代表者役職・代表者名",    "✓", "例：代表取締役　山田太郎",       ""],
        ["代表者（フリガナ）",      "✓", "全角カタカナで入力",             ""],
        ["担当者名",               "✓", "連絡担当者の氏名",              ""],
        ["担当者名（フリガナ）",    "✓", "全角カタカナで入力",             ""],
        ["郵便番号",               "✓", "ハイフンなし7桁",               "例：3330852"],
        ["住所",                   "✓", "都道府県から番地まで",           ""],
        ["電話番号",               "✓", "ハイフンなし10〜11桁",          "例：0482571111"],
        ["メールアドレス",          "✓", "連絡先メールアドレス",           "請求書の送付先"],
        ["協賛区分",               "✓", "S・A・B・C・D・E から選択",     "申込み期間外は表示されません"],
        ["会社HP URL",             "　", "任意入力",                     "httpsから始まるURL"],
        ["掲載名（任意）",          "　", "印刷物等への掲載名",            "空欄の場合は会社名を使用"],
        ["利用規約への同意",        "✓", "規約を最後まで読んでチェック",   "スクロール後にチェック可能"],
    ], [42*mm, 10*mm, 52*mm, 51*mm])
    story.append(fields_data)

    story.append(Paragraph("3.3 協賛区分について", s['h2']))
    story.append(Paragraph(
        "協賛区分はS〜Eの6段階があります。S・A区分は抽選制のため、"
        "申込み後に当落のご連絡をお送りします。B〜E区分は申込み順となります。",
        s['body']
    ))
    kubun_data = tbl([
        ["区分", "申込み方式", "請求書送信タイミング", "金額（税抜）"],
        ["S", "抽選制", "当選通知と同時", "Infoシートの PRICE_S"],
        ["A", "抽選制", "当選通知と同時", "Infoシートの PRICE_A"],
        ["B", "先着制", "申込み直後（自動）", "Infoシートの PRICE_B"],
        ["C", "先着制", "申込み直後（自動）", "Infoシートの PRICE_C"],
        ["D", "先着制", "申込み直後（自動）", "Infoシートの PRICE_D"],
        ["E", "先着制", "申込み直後（自動）", "Infoシートの PRICE_E"],
    ], [15*mm, 25*mm, 50*mm, 65*mm])
    story.append(kubun_data)

    story.append(Paragraph("3.4 送信後の流れ", s['h2']))
    for step in [
        "「送信」ボタンを押すと、確認画面に受付番号（例：KWGC0618120530123）が表示されます。",
        "B〜E区分の場合、申込み直後に請求書PDF付きのメールが届きます。",
        "S・A区分の場合、抽選結果が決まり次第、別途メールでご連絡します。",
        "メールが届かない場合は、迷惑メールフォルダをご確認ください。",
    ]:
        story.append(Paragraph(f"・{step}", s['step']))

    story.append(Paragraph(
        "【ヒント】送信直後にメールが届かない場合は、数分お待ちください。"
        "Gmail の送信上限に達している場合、事務局より手動でご連絡する場合があります。",
        s['tip']
    ))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════
    # 4. データ管理
    # ══════════════════════════════════════════════════════════════════════
    story.append(Paragraph("4. 申込みデータの管理（事務局向け）", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=C_NAVY, spaceAfter=6))
    story.append(Paragraph(
        "申込みデータは Google Sheets の「協賛申込み一覧」シートに自動保存されます。"
        "事務局は「手作業」シートを使って受付・請求・入金確認を行います。",
        s['body']
    ))

    story.append(Paragraph("4.1 手作業シートの構成", s['h2']))
    story.append(Spacer(1, 2*mm))
    d = draw_sheet_mockup()
    story.append(DrawingFlowable(d, caption="図4-1  手作業シートのイメージ（主要列）"))
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph("4.2 業務フロー全体図", s['h2']))
    story.append(Spacer(1, 2*mm))
    d = draw_staff_workflow()
    story.append(DrawingFlowable(d, caption="図4-2  事務局の業務フロー"))
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph("4.3 受付番号の入力方法", s['h2']))
    story.append(Paragraph(
        "申込みが行われると、「協賛申込み一覧」シートに自動的にデータが追記されます。"
        "事務局は以下の手順で「手作業」シートに受付番号を入力してください。",
        s['body']
    ))
    for i, step in enumerate([
        "「協賛申込み一覧」シートを開き、新しい申込み行の受付番号（A列）を確認します。",
        "「手作業」シートを開き、空いているA列の行に受付番号をコピー＆ペーストします。",
        "B〜H列には申込みデータが自動的にXLOOKUPで入力されます。",
        "内容を確認し、問題なければ次のステップへ進みます。",
    ], 1):
        story.append(Paragraph(f"　{i}. {step}", s['step']))

    story.append(Paragraph(
        "【重要】受付番号は必ず「協賛申込み一覧」シートからコピーしてください。"
        "手入力の場合、1文字でも違うとXLOOKUPが機能しません。",
        s['warn']
    ))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════
    # 5. 請求書メール
    # ══════════════════════════════════════════════════════════════════════
    story.append(Paragraph("5. 請求書メールの送信（事務局向け）", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=C_NAVY, spaceAfter=6))
    story.append(Paragraph(
        "B〜E区分は申込み直後に自動送信されます。S・A区分のみ、以下の手順で手動送信が必要です。",
        s['body']
    ))

    story.append(Paragraph("5.1 メール送信のタイミング", s['h2']))
    story.append(Spacer(1, 2*mm))
    d = draw_email_flow()
    story.append(DrawingFlowable(d, caption="図5-1  メール送信タイミングの一覧"))
    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("5.2 S・A区分の請求書送信手順", s['h2']))
    for i, step in enumerate([
        "「手作業」シートで対象の行を見つけます（B列が「S」または「A」の行）。",
        "I列（受付完了）のチェックボックスをクリックします。",
        "確認ダイアログが表示されます（図5-2参照）。",
        "内容を確認し「送信する」ボタンをクリックします。",
        "J列（請求書送信日時）に送信日時が自動的に記録されます。",
    ], 1):
        story.append(Paragraph(f"　{i}. {step}", s['step']))

    story.append(Spacer(1, 4*mm))
    dialog1 = draw_dialog_mockup(
        "請求書送信の確認",
        ["以下の申込者に請求書メールを送信します。",
         "会社名：株式会社サンプル",
         "区分：A区分　金額：¥50,000"],
        ["送信する", "キャンセル"],
        [C_BLUE, C_GRAY]
    )
    story.append(DrawingFlowable(dialog1, caption="図5-2  請求書送信確認ダイアログ"))
    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("5.3 請求書PDFの内容", s['h2']))
    story.append(Paragraph("送信される請求書PDFには以下の情報が含まれます。", s['body']))
    pdf_contents = tbl([
        ["項目", "データソース"],
        ["会社名・担当者名",    "申込みフォームのデータ"],
        ["協賛区分・金額（税抜）", "申込みデータ＋Infoシート"],
        ["消費税（10%）・合計金額", "自動計算"],
        ["インボイス登録番号",  "Infoシート INVOICE_REG_NO"],
        ["振込先銀行情報",     "Infoシート BANK_NAME / BANK_NO / BANK_HOLDER"],
        ["入金期限",          "Infoシート PAYMENT_DUE"],
        ["主催団体名・代表者名", "Infoシート ORG_NAME / ORG_REP"],
        ["組織印影（印鑑）",   "Script Properties に登録済みの画像"],
    ], [60*mm, 95*mm])
    story.append(pdf_contents)

    story.append(Paragraph("5.4 送信済みの確認", s['h2']))
    story.append(Paragraph(
        "J列（請求書送信日時）に日時が入力されていれば送信完了です。"
        "この列に値がある場合、I列を再チェックしても「送信済み」の警告が表示され、"
        "二重送信を防止します。",
        s['body']
    ))
    story.append(Paragraph(
        "【ヒント】送信取り消しはできません。誤送信の場合は申込者に直接ご連絡ください。",
        s['warn']
    ))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════
    # 6. 入金確認とお礼状
    # ══════════════════════════════════════════════════════════════════════
    story.append(Paragraph("6. 入金確認とお礼状送信（事務局向け）", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=C_NAVY, spaceAfter=6))
    story.append(Paragraph(
        "入金が確認できたら、K列（入金完了）をチェックします。"
        "チェックするとお礼状PDF付きのメールが申込者に自動送信されます。",
        s['body']
    ))

    story.append(Paragraph("6.1 入金確認の手順", s['h2']))
    for i, step in enumerate([
        "銀行の入金確認を行い、振込人名と受付番号を照合します。",
        "「手作業」シートで対象の行（J列に請求書送信日時が入っている行）を確認します。",
        "K列（入金完了）のチェックボックスをクリックします。",
        "確認ダイアログが表示されます（図6-1参照）。",
        "内容を確認し「送信する」ボタンをクリックします。",
        "L列（お礼状送信日時）に送信日時が自動記録されます。",
    ], 1):
        story.append(Paragraph(f"　{i}. {step}", s['step']))

    story.append(Spacer(1, 4*mm))
    dialog2 = draw_dialog_mockup(
        "お礼状送信の確認",
        ["入金を確認しました。お礼状メールを送信します。",
         "会社名：山田商会",
         "区分：B区分"],
        ["送信する", "キャンセル"],
        [C_GREEN, C_GRAY]
    )
    story.append(DrawingFlowable(dialog2, caption="図6-1  お礼状送信確認ダイアログ"))
    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("6.2 J列が空の場合の注意", s['h2']))
    story.append(Paragraph(
        "K列（入金完了）にチェックするためには、J列（請求書送信日時）に日時が入っている必要があります。"
        "J列が空の状態でK列をチェックすると、警告が表示されてチェックが自動的に解除されます。",
        s['body']
    ))
    story.append(Paragraph(
        "【重要】必ず請求書を送信（J列に日時が入る）してから入金確認を行ってください。",
        s['warn']
    ))

    story.append(Paragraph("6.3 チェックを誤って入れた場合", s['h2']))
    story.append(Paragraph(
        "I列またはK列のチェックを誤ってONにしてしまい、キャンセルした場合（または外した場合）は、"
        "確認ダイアログで「キャンセル」を選択するか、チェックを外す際の確認で「はい」を選択してください。"
        "関連する日時列（J・L列）の値も合わせてクリアするかどうかを確認されます。",
        s['body']
    ))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════
    # 7. メールクォータ管理
    # ══════════════════════════════════════════════════════════════════════
    story.append(Paragraph("7. メールクォータの管理", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=C_NAVY, spaceAfter=6))
    story.append(Paragraph(
        "Google の無料アカウントでは1日に送信できるメール数が最大50通に制限されています。"
        "このシステムはクォータ（残り送信可能数）を自動で監視し、不足時に通知します。",
        s['body']
    ))

    quota_info = tbl([
        ["状態", "システムの動作", "対応方法"],
        ["残り通数 > MIN_MAIL_QUOTA\n（通常）", "通常通り動作",
         "対応不要"],
        ["残り通数 ≤ MIN_MAIL_QUOTA\n（警告レベル）",
         "管理者（OFFICE_EMAIL）に警告メールを送信\nフォームは引き続き受付",
         "翌日まで待つか、メール送信を手動対応に切り替える"],
        ["残り通数 = 0\n（上限超過）",
         "Webフォームの受付を停止\nエラー画面を表示",
         "翌日（太平洋時間の深夜）にリセットされるまで待つ"],
    ], [40*mm, 60*mm, 55*mm])
    story.append(quota_info)

    story.append(Spacer(1, 3*mm))
    story.append(Paragraph("7.1 MIN_MAIL_QUOTA の変更", s['h2']))
    story.append(Paragraph(
        "Infoシートの MIN_MAIL_QUOTA に数値を入力することで、警告の閾値を変更できます。"
        "デフォルトは 5 です。申込み件数が多い時期は 10〜15 程度に設定することを推奨します。",
        s['body']
    ))
    story.append(Paragraph(
        "【ヒント】Google Workspace（有料）を使用している場合、送信上限は1,500通/日まで引き上げられます。",
        s['tip']
    ))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════
    # 8. FAQ
    # ══════════════════════════════════════════════════════════════════════
    story.append(Paragraph("8. よくある質問（FAQ）", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=C_NAVY, spaceAfter=6))

    faqs = [
        ("申込み後にメールが届きません。",
         "①迷惑メールフォルダをご確認ください。②事務局のメールクォータが上限に達している場合があります。"
         "この場合、翌日以降に届くか、事務局から手動でご連絡します。"),
        ("フォームで区分が選択できません。",
         "申込み受付期間外の場合、該当区分は表示されません。Infoシートの KUBUN_SA_START/END または "
         "KUBUN_BCDE_START/END の日付設定をご確認ください。"),
        ("手作業シートにデータが表示されません（XLOOKUP が機能しない）。",
         "受付番号が正確に入力されているか確認してください。スペースや全角・半角の違いが原因のことが多いです。"
         "「協賛申込み一覧」シートからコピー＆ペーストすることを推奨します。"),
        ("誤った受付番号を入力してしまいました。",
         "手作業シートのA列の値を正しい受付番号に修正すれば、B〜H列のXLOOKUPが自動で更新されます。"
         "ただし、既にI列やK列をチェックして操作を行っていた場合は、J・L列の値も確認してください。"),
        ("請求書のインボイス登録番号が空欄です。",
         "Infoシートの INVOICE_REG_NO に正しい登録番号（T＋13桁）を入力してください。"
         "入力後、Script Properties との同期のために「WebHanabi」→「設定を同期」を実行してください。"),
        ("お礼状のK列にチェックできません。",
         "J列（請求書送信日時）に値が入っているかを確認してください。J列が空の場合はK列をチェックできません。"),
        ("背景画像を変更したいです。",
         "変更したい画像を Google Drive にアップロードし、そのファイルIDを Infoシートの BG_IMAGE_ID に"
         "入力してください。その後、Webアプリを新バージョンとして再デプロイしてください。"),
    ]
    for q, a in faqs:
        story.append(Paragraph(f"Q：{q}", ParagraphStyle('fq', fontName=JP_B, fontSize=10,
                     leading=16, textColor=C_BLUE, spaceBefore=8, spaceAfter=2)))
        story.append(Paragraph(f"A：{a}", ParagraphStyle('fa', fontName=JP, fontSize=10,
                     leading=16, textColor=C_BLACK, leftIndent=14, spaceAfter=4,
                     backColor=C_LLBLUE, borderPad=4)))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════
    # 9. トラブルシューティング
    # ══════════════════════════════════════════════════════════════════════
    story.append(Paragraph("9. トラブルシューティング", s['h1']))
    story.append(HRFlowable(width="100%", thickness=1, color=C_NAVY, spaceAfter=6))

    issues = tbl([
        ["症状", "原因", "対処方法"],
        ["フォームを開くとエラーメッセージが表示される",
         "Infoシートの必須項目が未設定、または Data Spreadsheet ID が無効",
         "Infoシートの設定を確認し、空欄の必須項目を埋める。\nProject Initialize を再実行する。"],
        ["「メール送信数の上限に達しました」と表示される",
         "Gmail の1日送信上限（50通）に達した",
         "翌日まで待つ（太平洋時間 深夜にリセット）。\nGoogle Workspace への移行を検討する。"],
        ["チェックボックスをクリックしてもダイアログが出ない",
         "onEditInstallable トリガーが未登録、または削除された",
         "Apps Script エディタで「トリガー」を確認し、onEditInstallable が登録されているか確認する。\n未登録の場合は Project Initialize を再実行する。"],
        ["PDFが添付されていないメールが届いた",
         "PDF生成中にエラーが発生した",
         "Apps Script のログ（実行ログ）でエラーを確認する。\n手動でPDFを生成して添付し、再送信する。"],
        ["XLOOKUPが #N/A エラーになる",
         "受付番号が一致しない（スペース・文字種の違い）",
         "協賛申込み一覧シートの受付番号を確認し、コピー＆ペーストで再入力する。"],
        ["フォームのURLを開いても画面が真っ白",
         "Webアプリが正しくデプロイされていない",
         "Apps Script エディタで「デプロイを管理」を確認し、\n「アクセスできるユーザー：全員」になっているか確認する。"],
    ], [45*mm, 52*mm, 58*mm])
    story.append(issues)

    story.append(Spacer(1, 5*mm))
    story.append(Paragraph("9.1 ログの確認方法", s['h2']))
    for step in [
        "Google Apps Script エディタ（script.google.com）を開く。",
        "左側メニューの「実行数」（時計アイコン）をクリックする。",
        "最近の実行一覧が表示される。エラーが発生した実行は赤くハイライトされる。",
        "該当の実行行をクリックすると、エラーの詳細ログが確認できる。",
    ]:
        story.append(Paragraph(f"・{step}", s['step']))

    story.append(Spacer(1, 5*mm))
    story.append(HRFlowable(width="100%", thickness=1, color=C_LBLUE, spaceAfter=6))
    story.append(Paragraph(
        "本マニュアルに記載のない問題が発生した場合は、システム管理者にご連絡ください。",
        ParagraphStyle('footer_note', fontName=JP, fontSize=9, leading=15,
                       textColor=C_GRAY, alignment=TA_CENTER)
    ))

    return story


def main():
    import os
    out = '/Users/minhsang1601/GitClone/WebHanabi/操作マニュアル_WebHanabi.pdf'
    doc = SimpleDocTemplate(
        out,
        pagesize=A4,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=MARGIN,
        bottomMargin=22*mm,
        title="WebHanabi 川口花火大会 協賛管理システム 操作マニュアル",
        author="WebHanabi Project",
    )
    story = build_story()
    def on_first(canvas, doc):
        draw_cover_background(canvas, doc)
        # No footer on cover

    doc.build(story, onFirstPage=on_first, onLaterPages=add_header_footer)
    print(f"Generated: {out}")


if __name__ == '__main__':
    main()

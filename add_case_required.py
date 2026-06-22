#!/usr/bin/env python3
"""必須項目未入力時の警告テストケースを1行、既存フォーマットで生成。"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

NAVY="1A3A5C"; BLUE="2E6DA4"; LLBLUE="F0F5FB"; WHITE="FFFFFF"; BLACK="222222"
FONT="Meiryo"
def F(size=10,bold=False,color=BLACK,name=FONT): return Font(name=name,size=size,bold=bold,color=color)
thin=Side(style="thin",color="BBBBBB"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
def cell(ws,coord,v,font=None,bg=None,align="left",wrap=False):
    c=ws[coord]; c.value=v
    if font: c.font=font
    if bg: c.fill=PatternFill("solid",fgColor=bg)
    c.alignment=Alignment(horizontal=align,vertical="center",wrap_text=wrap); c.border=border
    return c

HDR=["No.","テスト項目","操作手順","期待結果","判定","実施日","実施者","備考"]
W=[8,22,50,55,9,13,12,24]

wb=Workbook(); ws=wb.active; ws.title="追加テストケース"
ws.sheet_view.showGridLines=False
for j,w in enumerate(W): ws.column_dimensions[get_column_letter(1+j)].width=w
for j,h in enumerate(HDR): cell(ws,f"{get_column_letter(1+j)}1",h,F(10,True,WHITE),NAVY,"center")
ws.row_dimensions[1].height=22

cases=[{
 "id":"T-44",
 "name":"必須項目未入力時の警告",
 "steps":"①区分・会社名・フリガナ・代表者名・担当者名・郵便番号・住所・電話番号・メール等の必須項目を空欄のままにする\n②「送信」ボタンをクリックする",
 "exp":"・最初の未入力必須項目にフォーカスが移動する\n・「こちらは入力必須項目です。」等の警告メッセージ（吹き出し）が表示される\n・フォームは送信されず、確認画面に進まない\n・スプレッドシートにデータが登録されない",
},{
 "id":"T-45",
 "name":"誓約事項未同意時の送信",
 "steps":"①全必須項目を入力する\n②「誓約事項に同意する」チェックを外したまま「送信」をクリックする",
 "exp":"・同意チェックボックスに「このチェックボックスをオンにしてください。」等の警告が表示される\n・フォームは送信されない",
}]

r=2
dv=DataValidation(type="list",formula1='"OK,NG,未,確認"',allow_blank=True); ws.add_data_validation(dv)
for c in cases:
    cell(ws,f"A{r}",c["id"],F(10,True,BLUE,"Consolas"),LLBLUE,"center")
    cell(ws,f"B{r}",c["name"],F(10,True,BLACK),LLBLUE,"left",wrap=True)
    cell(ws,f"C{r}",c["steps"],F(9.5),WHITE,"left",wrap=True)
    cell(ws,f"D{r}",c["exp"],F(9.5),WHITE,"left",wrap=True)
    cell(ws,f"E{r}","",F(10,True),WHITE,"center")
    cell(ws,f"F{r}","",F(9.5),WHITE,"center")
    cell(ws,f"G{r}","",F(9.5),WHITE,"center")
    cell(ws,f"H{r}","",F(9.5),WHITE,"left",wrap=True)
    ws.row_dimensions[r].height=max(34,(c["exp"].count("\n")+1)*15)
    r+=1
dv.add(f"E2:E{r-1}")

wb.save("追加テストケース_必須項目未入力.xlsx")
print("✓ saved — rows:", len(cases))

import win32com.client
from openpyxl import Workbook
import re

# === 設定 ===
input_path = r"C:\path\to\your\file.xlsx"          # Excel元ファイル
output_path = r"C:\path\to\chart_headers.xlsx"     # 出力ファイル
target_sheet_name = "グラフシート"                   # グラフがあるシート名

# === Excel起動 ===
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(input_path)

# シート辞書を用意
ws_dict = {ws.Name: ws for ws in wb.Sheets}

# === 出力用ブック作成 ===
output_wb = Workbook()
output_ws = output_wb.active
output_ws.title = "HeaderInfo"
output_ws.append([
    "Chart Index", "Series Name",
    "X Sheet", "X Column", "X Header1", "X Header2", "X Header3",
    "Y Sheet", "Y Column", "Y Header1", "Y Header2", "Y Header3"
])

# === ヘルパー関数 ===

def parse_formula(formula):
    """
    =SERIES("名前", 'シート'!$E$4:$E$404, 'シート'!$F$4:$F$404, 1)
    → x_range, y_range を抽出
    """
    try:
        inner = formula.strip()[len("=SERIES("):-1]
        parts = [p.strip() for p in inner.split(",")]
        if len(parts) >= 3:
            return parts[1], parts[2]  # x_range, y_range
    except:
        pass
    return None, None

def parse_range_for_header(range_str):
    """
    'シート'!$E$4:$E$404 → (シート名, 列)
    """
    if not range_str:
        return None, None
    match = re.match(
        r"^(?:'(?P<sheet_quoted>[^']+)'|(?P<sheet_unquoted>[^!]+))!\$?(?P<col>[A-Z]+)\$?\d+:\$?(?P=col)\$?\d+",
        range_str.strip()
    )
    if match:
        sheet_name = match.group("sheet_quoted") or match.group("sheet_unquoted")
        col = match.group("col")
        return sheet_name, col
    return None, None

def get_headers(sheet_name, col_letter):
    if not sheet_name or not col_letter:
        return ["", "", ""]
    try:
        ws = ws_dict.get(sheet_name)
        return [ws.Range(f"{col_letter}{r}").Value for r in range(1, 4)]
    except:
        return ["", "", ""]

# === メイン処理 ===
try:
    sheet = wb.Sheets(target_sheet_name)
    chart_objects = sheet.ChartObjects()

    for i in range(1, chart_objects.Count + 1):
        chart = chart_objects.Item(i).Chart
        for s in chart.SeriesCollection():
            name = str(s.Name)
            formula = str(s.Formula)
            x_range, y_range = parse_formula(formula)

            x_sheet, x_col = parse_range_for_header(x_range)
            y_sheet, y_col = parse_range_for_header(y_range)

            x_headers = get_headers(x_sheet, x_col)
            y_headers = get_headers(y_sheet, y_col)

            output_ws.append([
                i, name,
                x_sheet or "", x_col or "", *x_headers,
                y_sheet or "", y_col or "", *y_headers
            ])

    output_wb.save(output_path)
    print(f"✅ 出力完了: {output_path}")

except Exception as e:
    print(f"❌ エラー: {e}")

finally:
    wb.Close(False)
    excel.Quit()
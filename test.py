import win32com.client
from openpyxl import Workbook
import re

# === 設定 ===
input_path = r"C:\path\to\your\file.xlsx"
output_path = r"C:\path\to\chart_headers.xlsx"
target_sheet_name = "グラフシート"

# === Excel起動 ===
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(input_path)
ws_dict = {ws.Name: ws for ws in wb.Sheets}

# === 出力用ブック作成 ===
output_wb = Workbook()
output_ws = output_wb.active
output_ws.title = "HeaderInfo"
output_ws.append([
    "Chart Index (見た目順)", "Series Name",
    "X Sheet", "X Column", "X Header1", "X Header2", "X Header3",
    "Y Sheet", "Y Column", "Y Header1", "Y Header2", "Y Header3",
    "X Axis Label", "Y Axis Label"
])

# === ヘルパー関数 ===

def parse_formula(formula):
    try:
        inner = formula.strip()[len("=SERIES("):-1]
        parts = [p.strip() for p in inner.split(",")]
        if len(parts) >= 3:
            return parts[1], parts[2]
    except:
        pass
    return None, None

def parse_range_for_header(range_str):
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

def get_axis_title(chart, axis_type):
    try:
        axis = chart.Axes(axis_type)
        if axis.HasTitle:
            return axis.AxisTitle.Text
    except:
        pass
    return ""

# === メイン処理 ===

try:
    sheet = wb.Sheets(target_sheet_name)
    chart_objects = sheet.ChartObjects()

    # グラフと位置情報を取得（COMオブジェクト除外で比較）
    chart_info_list = []
    for i in range(1, chart_objects.Count + 1):
        chart_obj = chart_objects.Item(i)
        chart = chart_obj.Chart
        chart_info_list.append({
            "top": chart_obj.Top,
            "left": chart_obj.Left,
            "chart_obj": chart_obj,
            "chart": chart
        })

    # 見た目順にソート（左上 → 右下）
    chart_info_list.sort(key=lambda c: (c["top"], c["left"]))

    # 出力処理（並び順どおり）
    for display_index, info in enumerate(chart_info_list, start=1):
        chart = info["chart"]
        x_axis_label = get_axis_title(chart, 1)  # xlCategory
        y_axis_label = get_axis_title(chart, 2)  # xlValue

        for s in chart.SeriesCollection():
            name = str(s.Name)
            formula = str(s.Formula)
            x_range, y_range = parse_formula(formula)

            x_sheet, x_col = parse_range_for_header(x_range)
            y_sheet, y_col = parse_range_for_header(y_range)

            x_headers = get_headers(x_sheet, x_col)
            y_headers = get_headers(y_sheet, y_col)

            output_ws.append([
                display_index, name,
                x_sheet or "", x_col or "", *x_headers,
                y_sheet or "", y_col or "", *y_headers,
                x_axis_label, y_axis_label
            ])

    output_wb.save(output_path)
    print(f"✅ 出力完了（軸ラベル含む）: {output_path}")

except Exception as e:
    print(f"❌ エラー: {e}")

finally:
    wb.Close(False)
    excel.Quit()
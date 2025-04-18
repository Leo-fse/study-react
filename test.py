import win32com.client
from openpyxl import Workbook
import re

# Excelファイルのパス
input_path = r"C:\path\to\your\file.xlsx"
output_path = r"C:\path\to\chart_headers.xlsx"
target_sheet_name = "グラフシート"

# Excel起動
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(input_path)
ws_dict = {ws.Name: ws for ws in wb.Sheets}

# 出力用ブック
output_wb = Workbook()
output_ws = output_wb.active
output_ws.title = "HeaderInfo"
output_ws.append(["Chart Index", "Series Name", "X Column", "X Header1", "X Header2", "X Header3",
                                             "Y Column", "Y Header1", "Y Header2", "Y Header3"])

# セル範囲から列文字とシート名を抽出する関数
def parse_range(formula_str):
    if not formula_str or not formula_str.startswith("="):
        return None, None

    match = re.search(r"=(?:'?(?P<sheet>[^']+)'?)?!\$?(?P<col>[A-Z]+)\$?\d+:\$?(?P=col)\$?\d+", formula_str)
    if match:
        return match.group("sheet"), match.group("col")
    return None, None

try:
    sheet = wb.Sheets(target_sheet_name)
    chart_objects = sheet.ChartObjects()

    for i in range(1, chart_objects.Count + 1):
        chart = chart_objects.Item(i).Chart
        for s in chart.SeriesCollection():
            series_name = f"{s.Name}"

            # X, Y の範囲取得
            x_formula = f"{s.Formula}".replace("=", "=")  # 保持
            y_formula = f"{s.Formula}".replace("=", "=")

            try:
                x_range = f"{s.XValues}"
                x_sheet_name, x_col = parse_range(x_range)
            except:
                x_sheet_name, x_col = None, None

            try:
                y_range = f"{s.Values}"
                y_sheet_name, y_col = parse_range(y_range)
            except:
                y_sheet_name, y_col = None, None

            # ヘッダー取得
            def get_headers(sheet_name, col):
                if not sheet_name or not col:
                    return ["", "", ""]
                try:
                    ws = ws_dict.get(sheet_name)
                    return [ws.Range(f"{col}{r}").Value for r in range(1, 4)]
                except:
                    return ["", "", ""]

            x_headers = get_headers(x_sheet_name, x_col)
            y_headers = get_headers(y_sheet_name, y_col)

            output_ws.append([
                i, series_name,
                x_col, *x_headers,
                y_col, *y_headers
            ])

    output_wb.save(output_path)
    print(f"✅ ヘッダー情報出力完了: {output_path}")

except Exception as e:
    print(f"❌ エラー: {e}")

finally:
    wb.Close(False)
    excel.Quit()
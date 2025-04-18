import win32com.client
import openpyxl
from openpyxl import Workbook

# 元のExcelファイルと出力ファイルのパス
input_path = r"C:\path\to\your\file.xlsx"
output_path = r"C:\path\to\chart_series_list.xlsx"
target_sheet_name = "グラフシート"  # 対象シート名

# Excelアプリケーションを起動
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

# 元のファイルを開く
wb = excel.Workbooks.Open(input_path)

# 出力用Excel作成
output_wb = Workbook()
output_ws = output_wb.active
output_ws.title = "SeriesList"
output_ws.append(["Chart Index", "Series Name", "Formula"])

try:
    sheet = wb.Sheets(target_sheet_name)
    chart_objects = sheet.ChartObjects()

    for i in range(1, chart_objects.Count + 1):
        chart = chart_objects.Item(i).Chart
        for s in chart.SeriesCollection():
            series_name = str(s.Name)
            formula = str(s.FormulaLocal)  # ← Formula → FormulaLocal + 明示的に文字列化
            output_ws.append([i, series_name, formula])

    output_wb.save(output_path)
    print(f"出力完了: {output_path}")

except Exception as e:
    print(f"エラー: {e}")

finally:
    wb.Close(False)
    excel.Quit()
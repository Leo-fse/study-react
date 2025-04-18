import win32com.client
from openpyxl import Workbook

input_path = r"C:\path\to\your\file.xlsx"
output_path = r"C:\path\to\chart_series_list.xlsx"
target_sheet_name = "グラフシート"

# Excel 起動
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(input_path)

# 出力用ワークブック作成
output_wb = Workbook()
output_ws = output_wb.active
output_ws.title = "SeriesList"
output_ws.append(["Chart Index", "Series Name", "XValues", "YValues", "Formula"])

try:
    sheet = wb.Sheets(target_sheet_name)
    chart_objects = sheet.ChartObjects()

    for i in range(1, chart_objects.Count + 1):
        chart = chart_objects.Item(i).Chart
        for s in chart.SeriesCollection():
            name = f"{s.Name}"
            formula = f"{s.Formula}".replace("=", "")  # 数式扱い防止

            # XValues・YValues は範囲参照 or 実データ
            try:
                xvalues = f"{s.XValues}"
            except Exception:
                xvalues = "N/A"

            try:
                yvalues = f"{s.Values}"
            except Exception:
                yvalues = "N/A"

            print(f"[DEBUG] name={name}, x={xvalues}, y={yvalues}")
            output_ws.append([i, name, xvalues, yvalues, formula])

    output_wb.save(output_path)
    print(f"✅ 出力完了: {output_path}")

except Exception as e:
    print(f"❌ エラー: {e}")

finally:
    wb.Close(False)
    excel.Quit()
import win32com.client

# Excelを起動
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

# 対象ファイルとシート
input_path = r"C:\path\to\your\file.xlsx"
target_sheet_name = "グラフシート"

# Excelブックを開く
wb = excel.Workbooks.Open(input_path)
sheet = wb.Sheets(target_sheet_name)
chart_objects = sheet.ChartObjects()

print(f"== シート: {target_sheet_name} ==")
for i in range(1, chart_objects.Count + 1):
    chart = chart_objects.Item(i).Chart
    print(f"  [Chart {i}]")
    for j, s in enumerate(chart.SeriesCollection(), start=1):
        try:
            formula = s.Formula  # ← COM型の可能性がある
            print(f"    Series {j} Name: {s.Name}")
            print(f"    Formula: {formula}")
        except Exception as e:
            print(f"    Series {j} Formula取得失敗: {e}")

wb.Close(False)
excel.Quit()
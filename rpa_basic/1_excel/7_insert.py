from openpyxl import load_workbook
wb = load_workbook("sample.xlsx")
ws = wb.active

ws.insert_rows(8, 5) # 8번 째 줄부터 새로운 row 5행이 생김
wb.save("sample_insert_rows.xlsx")

ws.insert_cols(2, 3) # B번 째 줄부터 새로운 cols 3열이 생김
wb.save("sample_insert_rows.xlsx")
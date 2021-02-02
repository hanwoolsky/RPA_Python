from openpyxl import load_workbook
wb = load_workbook("sample.xlsx")
ws = wb.active

ws.delete_rows(8, 3) # 8번 째 줄에 있는 7번 학생부터 3줄의 데이터 삭제
ws.save("sample_delete_row.xlsx")

ws.delete_cols(2, 2) # B번 째 줄에 있는 과목부터 2열의 데이터 삭제
ws.save("sample_delete_col.xlsx")
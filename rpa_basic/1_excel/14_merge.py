from openpyxl import workbook
wb = workbook()
ws = wb.active

# 셀 합치기
ws.merge_cells("B2:D2") #B2부터 D2까지 합치겠음
ws["B2"].value = "Merged Cell"

wb.save("sample_merge.xlsx")
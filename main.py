import openpyxl

wb = openpyxl.load_workbook('原始数据.xlsx')
sheet = wb['Sheet1']

rowIndex = 0
sourceData = []
for row in sheet.rows:
    if rowIndex < 1:
        rowIndex = rowIndex + 1
        continue
    rowData = []
    for cell in row:
        rowData.append(cell.value)
    print(rowData)

wb.close()
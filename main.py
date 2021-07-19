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

# 读客户对照表
wbClient = openpyxl.load_workbook('客户对照.xlsx')
sheetClient = wbClient['客户']
sheetStaff = wbClient['员工']
# 存储客户对照数据
clientStaffDict = {}
rowIndex = 0
for row in sheetClient.rows:
    if rowIndex < 1:
        rowIndex = rowIndex + 1
        continue
    clientName = row[0].value
    staffName = row[1].value
    clientStaffDict[clientName] = staffName
print(clientStaffDict)
# 存储员工数据
staffs = []
for row in sheetStaff.rows:
    staffs.append(row[0].value)
print(staffs)

wbClient.save('客户对照.xlsx')
wbClient.close()

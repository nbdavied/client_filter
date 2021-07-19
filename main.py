import openpyxl

wb = openpyxl.load_workbook('原始数据.xlsx')
sheet = wb['Sheet1']

wb.close()
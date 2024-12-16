import openpyxl
from prettytable import PrettyTable #For better visualisation

ss = openpyxl.load_workbook("D:\Selenium Practices\Week 4\Test Form.xlsx")

sheetData = ss.active

maxCol = sheetData.max_column
maxRow = sheetData.max_row

title = []
for i in range(1, maxCol + 1):
    cellObj = sheetData.cell(row=1, column=i)
    title.append(cellObj.value)
    
table = PrettyTable(title) #Table initialisation

rowCount = 2
for i in range(2, maxRow + 1):
    value = []
    for j in range(1, maxCol + 1):
        cellObj = sheetData.cell(row=i, column=j)
        value.append(cellObj.value)
    table.add_row(value)

print(table)
from openpyxl import Workbook
from openpyxl import load_workbook
#wb1 = Workbook()
wb1 = load_workbook('Workbook1.xlsx')
ws1 = wb1.active
ws1.title = "Worksheet1"
ws1.sheet_properties.tabColor = "1072BA"
for sheet in wb1:
	print(sheet.title)
wb1.save('Workbook1.xlsx')

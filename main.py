from openpyxl import Workbook
from openpyxl.styles import *
from openpyxl.worksheet.table import Table, TableStyleInfo

text_file = open("./employees.txt")

records = []

text_file.seek(0)

for record in text_file.readlines():
    records.append(record.rstrip("\n").split(";"))

print(records)

workbook = Workbook()

file_path = "./MyCompanyStaff.xlsx"

workbook.save(file_path)

#The default name is sheet
sheet = workbook["Sheet"]

sheet.title = "Employees"

for row in records:
    sheet.append(row)
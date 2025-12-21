import openpyxl
import os

path = r"c:\Users\zhangyh\Desktop\AutoPackage\④各店铺明细_模板.xlsx"
print(f"Checking path: {path}, Exists: {os.path.exists(path)}")

try:
    wb = openpyxl.load_workbook(path)
    sheet = wb.active
    print(f"Sheet Name: {sheet.title}")
    print("Rows 1-10:")
    for row in sheet.iter_rows(min_row=1, max_row=10):
        print([cell.value for cell in row])
except Exception as e:
    print(f"Error: {e}")

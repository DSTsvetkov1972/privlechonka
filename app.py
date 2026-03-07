import openpyxl
import os


file = os.path.join("files", "свод по каждому заказу.xlsx")

wb = openpyxl.load_workbook(file)

print("Получаем список листов книги")
sheet_names = wb.sheetnames


for n, sheet in enumerate(sheet_names):
    ws = wb[sheet]
    print(f" лист {sheet} {n} из { len(sheet_names) }")
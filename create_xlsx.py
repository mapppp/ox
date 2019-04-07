from openpyxl import Workbook

wb = Workbook()
sheet = wb.active
for i in range(1, 30001):
    sheet["A"+str(i)] = str(i)
wb.save("总表.xlsx")

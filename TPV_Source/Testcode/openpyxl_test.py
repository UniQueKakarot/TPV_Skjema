import openpyxl as op

wb = op.Workbook() #Opens up a new excel workbook

ws = wb.active #refrencing the active worksheet in the workbook

months = ['Januar', 'Februar', 'Mars', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Desember']

print(len(months))

for i in months:
    wb.create_sheet(i)

for sheet in wb:
    print(sheet.title)

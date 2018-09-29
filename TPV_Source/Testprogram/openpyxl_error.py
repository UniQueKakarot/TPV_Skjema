import openpyxl as op

wb = op.Workbook()
try:
    ws = wb['Test']
except KeyError:
    print('Hello')

print(ws)
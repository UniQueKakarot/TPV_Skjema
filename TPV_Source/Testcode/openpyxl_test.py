import openpyxl as op

wb = op.Workbook() #Opens up a new excel workbook


#Sheet names:
months = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
          'August', 'September', 'October', 'November', 'December']

colum_head = ['Spindel', 'Spindel Oljetank', 'Hydraulisk Enhet', 'Kjølevannsystem',
              'Maskindeksling & ATC', 'Roterende Akser', 'Kabinett Filter', 'Verktøymagasin', 'Spontransportør']
index = []

for i in months:
    wb.create_sheet(i)

sh = wb.get_sheet_names()
print(sh)
x = wb.get_sheet_by_name('Sheet')
wb.remove_sheet(x)

print(sh)

wb.save('test1.xlsx')

# for sheet in wb:
#     print(sheet.title)

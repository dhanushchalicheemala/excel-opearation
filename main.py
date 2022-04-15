import openpyxl as xl
toexcel = xl.load_workbook('practise.xlsx')
tosheet = toexcel['Dinesh']
fromexcel = xl.load_workbook('from excel.xlsx')
fromsheet = fromexcel['Dinesh']
for i in range(2, tosheet.max_row+1):
    tem = tosheet.cell(i, 2)
    for j in range(2, fromsheet.max_row + 1):
        if tem.value == fromsheet.cell(j, 2).value:
            tosheet.cell(i, 14).value = fromsheet.cell(j, 24).value


toexcel.save('new.xlsx')

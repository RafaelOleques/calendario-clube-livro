import xlwings as xw

wb = xw.Book('./calendario_clube.xlsx')
sh = wb.sheets[0]
sh.range('A1').value = 'hello'
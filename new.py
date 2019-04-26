import xlrd
import xlwt
from xlutils.copy import copy
import numpy as np
import os


data = xlrd.open_workbook('test1.xlsx')
table = data.sheets()[0]

nrows = table.nrows


list_values = []
for x in range(0,nrows):
    list_values.append(table.cell_value(x,3))
print((list_values))

rb = xlrd.open_workbook('test2.xls',formatting_info=True)
wb = copy(rb)
ws = wb.get_sheet(0)
for x in range(len(list_values)):
    ws.write(x,2,list_values[x])
wb.save('test2.xls')


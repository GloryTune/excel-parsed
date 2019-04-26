import pandas as pd
import xlrd
import xlwt
from xlutils.copy import copy
df1 = pd.read_excel("D:\\资料\\data.xls", encoding = 'utf-8')
df2 = pd.read_excel("D:\\资料\\机动车.xlsx", encoding = 'utf-8')

c=df1.merge(df2[['纳税人名称','机动']],how='left',on='纳税人名称')

c.to_excel('D:\\资料\\test2.xls', encoding = 'utf-8')
'''wb = load_workbook ('C:\\Users\\杨天佑\\Desktop\\资料\\test2.xlsx')
ws = wb['Sheet1']
sheet = wb.active
a = sheet.max_row
b=[]
i=0
while i < a:
    i = i+2
    #b.insert(i,ws.cell(i,14))
    b.append(ws.cell(i,14))
    print(b)
wbtest = load_workbook('C:\\Users\\杨天佑\\Desktop\\资料\\data.xlsx')
ws = wbtest['Sheet1']
sheet = wb.active
i = 0
while i < a:
    ws.cell (i+2,7).value =b[i]
    i = i+1
wbtest.save('C:\\Users\\杨天佑\\Desktop\\资料\\data.xlsx')'''
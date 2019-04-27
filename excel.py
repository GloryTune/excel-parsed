import pandas as pd
import xlrd
from xlutils.copy import copy
from tkinter.filedialog import askdirectory
from tkinter import *
import os
import tkinter as tk
root = Tk()
root.geometry('200x80+450+200')
root.title('天韵')
def xuanze():
    global a,m
    a = askdirectory()
    m = str(a)
    label_1 = tk.Label(root, text=m)
    label_1.place(x=80, y=0, width=120, height=25)
def jisuan():
    # 机动车Vlookup
    df1 = pd.read_excel(a+"/data.xls", encoding = 'utf-8')
    df2 = pd.read_excel(a+"/机动车.xls", encoding = 'utf-8')
    c=df1.merge(df2[['纳税人名称','机动']],how='left',on='纳税人名称')
    c.to_excel(a+'/test1.xls', encoding = 'utf-8')
    #机动车转移数据
    data = xlrd.open_workbook(a+'/test1.xls')
    table = data.sheets()[0]
    nrows = table.nrows
    list_values = []
    for x in range(1,nrows):
        list_values.append(table.cell_value(x,13))

    rb = xlrd.open_workbook(a+'/data.xls',formatting_info=True)
    wb = copy(rb)
    ws = wb.get_sheet(0)
    for x in range(len(list_values)):
        ws.write(x+1,6,list_values[x])
    wb.save(a+'/data.xls')
    # 海关
    af1 = pd.read_excel(a + "/data.xls", encoding='utf-8')
    af2 = pd.read_excel(a + "/海关.xls", encoding='utf-8')
    d = af1.merge(af2[['纳税人名称', '海关']], how='left', on='纳税人名称')
    d.to_excel(a + '/test2.xls', encoding='utf-8')
    # 海关转移数据
    data1 = xlrd.open_workbook(a + '/test2.xls')
    table1 = data1.sheets()[0]
    nrows1 = table1.nrows
    list_values1 = []
    for x in range(1, nrows1):
        list_values1.append(table1.cell_value(x, 13))

    rb1 = xlrd.open_workbook(a + '/data.xls', formatting_info=True)
    wb1 = copy(rb1)
    ws1 = wb1.get_sheet(0)
    for x in range(len(list_values1)):
        ws1.write(x + 1, 7, list_values1[x])
    wb1.save(a + '/data.xls')
    #认证
    bf1 = pd.read_excel(a + "/data.xls", encoding='utf-8')
    bf2 = pd.read_excel(a + "/认证.xls", encoding='utf-8')
    e = bf1.merge(bf2[['纳税人名称', '认证']], how='left', on='纳税人名称')
    e.to_excel(a + '/test3.xls', encoding='utf-8')
    # 认证转移数据
    data2 = xlrd.open_workbook(a + '/test3.xls')
    table2 = data2.sheets()[0]
    nrows2 = table2.nrows
    list_values2 = []
    for x in range(1, nrows2):
        list_values2.append(table2.cell_value(x, 13))

    rb2 = xlrd.open_workbook(a + '/data.xls', formatting_info=True)
    wb2 = copy(rb2)
    ws2 = wb2.get_sheet(0)
    for x in range(len(list_values2)):
        ws2.write(x + 1, 5, list_values2[x])
    wb2.save(a + '/data.xls')
    # 免抵退
    ef1 = pd.read_excel(a + "/data.xls", encoding='utf-8')
    ef2 = pd.read_excel(a + "/免抵退.xls", encoding='utf-8')
    f = ef1.merge(ef2[['纳税人名称', '免抵退']], how='left', on='纳税人名称')
    f.to_excel(a + '/test4.xls', encoding='utf-8')
    # 免抵退转移数据
    data3 = xlrd.open_workbook(a + '/test4.xls')
    table3 = data3.sheets()[0]
    nrows3 = table3.nrows
    list_values3 = []
    for x in range(1, nrows3):
        list_values3.append(table3.cell_value(x, 13))

    rb3 = xlrd.open_workbook(a + '/data.xls', formatting_info=True)
    wb3 = copy(rb3)
    ws3 = wb3.get_sheet(0)
    for x in range(len(list_values3)):
        ws3.write(x + 1, 3, list_values3[x])
    wb3.save(a + '/data.xls')

    #空位补零
    rf = pd.DataFrame(pd.read_excel(a + "/data.xls"))
    r= rf.fillna(0)
    r.to_excel(a + '/test5.xls', encoding='utf-8')
    #转移数据免抵退
    data4 = xlrd.open_workbook(a + '/test5.xls')
    table4 = data4.sheets()[0]
    nrows4 = table4.nrows
    list_values4 = []
    for x in range(1, nrows4):
        list_values4.append(table4.cell_value(x, 4))
    #print((list_values3))
    rb4 = xlrd.open_workbook(a + '/data.xls', formatting_info=True)
    wb4 = copy(rb4)
    ws4 = wb4.get_sheet(0)
    for x in range(len(list_values4)):
        ws4.write(x + 1, 3, list_values4[x])
    wb4.save(a + '/data.xls')
    #转移数据机动车
    data5 = xlrd.open_workbook(a + '/test5.xls')
    table5 = data5.sheets()[0]
    nrows5 = table5.nrows
    list_values5 = []
    for x in range(1, nrows5):
        list_values5.append(table5.cell_value(x, 7))
    # print((list_values3))
    rb5 = xlrd.open_workbook(a + '/data.xls', formatting_info=True)
    wb5 = copy(rb5)
    ws5 = wb5.get_sheet(0)
    for x in range(len(list_values5)):
        ws5.write(x + 1, 6, list_values5[x])
    wb5.save(a + '/data.xls')
    #转移数据海关
    data6 = xlrd.open_workbook(a + '/test5.xls')
    table6 = data6.sheets()[0]
    nrows6 = table6.nrows
    list_values6 = []
    for x in range(1, nrows6):
        list_values6.append(table6.cell_value(x, 8))
    # print((list_values3))
    rb6 = xlrd.open_workbook(a + '/data.xls', formatting_info=True)
    wb6 = copy(rb6)
    ws6 = wb6.get_sheet(0)
    for x in range(len(list_values6)):
        ws6.write(x + 1, 7, list_values6[x])
    wb6.save(a + '/data.xls')
    #转移数据认证
    data7 = xlrd.open_workbook(a + '/test5.xls')
    table7 = data7.sheets()[0]
    nrows7 = table7.nrows
    list_values7 = []
    for x in range(1, nrows7):
        list_values7.append(table7.cell_value(x, 6))

    rb7 = xlrd.open_workbook(a + '/data.xls', formatting_info=True)
    wb7 = copy(rb7)
    ws7 = wb7.get_sheet(0)
    for x in range(len(list_values7)):
        ws7.write(x + 1, 5, list_values7[x])
    wb7.save(a + '/data.xls')
    vf = pd.read_excel(a+"/data.xls", encoding = 'utf-8')
    vf.eval('sum = 报税销项税额 + 免抵退 -认证进项税额-认证-机动车-海关-上期留抵',inplace=True)
    vf.to_excel(a + '/test6.xls', encoding='utf-8')
    #总数中负数转为零并将数据转移到总表
    data9 = xlrd.open_workbook(a + '/test6.xls')
    table9 = data9.sheets()[0]
    nrows9 = table9.nrows
    list_values9 = []
    for x in range(1, nrows9):
        list_values9.append(table9.cell_value(x, 13))#list_values9现在存储着和那一列 是列表形式
    for w in list_values9:
        if w < 0:
            t = list_values9.index(w)
            list_values9[t]=0

    rb8 = xlrd.open_workbook(a + '/data.xls', formatting_info=True)
    wb8 = copy(rb8)
    ws8 = wb8.get_sheet(0)
    for x in range(len(list_values9)):
        ws8.write(x + 1, 9, list_values9[x])
    wb8.save(a + '/data.xls')
    #
    for name in ['test1.xls','test2.xls','test3.xls','test4.xls','test5.xls','test6.xls']:
        del_file = os.path.join(a, name)
        os.remove(del_file)
    m =str('转换完成')
    label_1 = tk.Label(root, text=m)
    label_1.place(x=80, y=25, width=100, height=25)
def caidan():
    mm =e_user.get()
    if mm == '1015':
        lable = tk.Label(root,text='我爱你')
        lable.place(x=80, y=25, width=120, height=25)
    else:
        lable2 = tk.Label(root, text='请输入正确的四位数')
        lable2.place(x=80, y=25, width=120, height=25)

btn=Button(root,text="选择文件夹",command=xuanze)
btn.place(x=0,y=0,width=80, height=25)
btn1=Button(root,text="一键转换",command=jisuan)
btn1.place(x=0,y=25,width=80, height=25)
btn2=Button(root,text="韵",command=caidan)
btn2.place(x=170,y=50,width=30,height=30)
e_user =Entry(root)
e_user.place(x=0,y=50,width=38, height=25)
root.mainloop()
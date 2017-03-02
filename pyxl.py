# -*- coding: utf-8 -*-
from __future__ import division    # 整数int相除后结果变为浮点型
import xlrd, xlwt, re, string
from xlutils.copy import copy

# 获取excel工作簿workbook表单sheet
data = xlrd.open_workbook('pyxl.xls')
table = data.sheets()[0]
nrows = table.nrows

# 拷贝工作簿进行修改
wb = copy(data)
ws = wb.get_sheet(0)

# 正则表达式匹配修改excel并保存
for rownum in range(0,nrows):
    mat =  table.cell(rownum,2).value
    m = re.match(r'^(\d{1,4})\((\d{2,4})', mat)
    # print m.group(1)
    m1 = string.atoi(m.group(1))
    m2 = string.atoi(m.group(2))
    # print round(m1*100/m2, 2)
    ws.write(rownum, 3, str(round(m1*100/m2, 2))+'%')

# 保存修改结果
wb.save('pyxl1.xls')

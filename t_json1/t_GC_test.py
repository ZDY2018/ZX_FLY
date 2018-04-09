# -*- coding:utf-8 -*-
import xlrd
import json

filename = u"2014全国最新行政区划及代号区号精确到县区级-全国省市区一览表 - 副本.xls"
book = xlrd.open_workbook(filename)  # ‘rb’可以加
for i in book.sheets():
    print (i.name)  # 省市区
sheet = book.sheet_by_index(0)  # 根据excle索引返回值
print (sheet, sheet.nrows, sheet.ncols)  # 此处有坑 ncols在excle中从B开始 所以列数比正常多一个
# for i in range(sheet.nrows):
#     print (sheet.row_values(i)[1])#索引第i行第1（b）列 //如果不标明列 则显示所有列 输出每一行地址
data = {}
for i in range(1, sheet.nrows):
    try:
        row = sheet.row_values(i)  # 索引第i行
        if (row[5] == 0):  # row[5](区划等级)为0时
            data[row[1]] = {}  # row[1](区划名称)
        if (row[5] == 1):
            data[row[2]][row[1]] = {}
        if (row[5] == 2):
            data[row[2]][row[3]][row[1]] = {}
    except:
        print(" ", i, row[1])

print json.dumps(data[u'黑龙江省'], ensure_ascii=False)#unicode转中文

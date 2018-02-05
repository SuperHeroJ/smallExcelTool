# -*- coding: utf-8 -*-
import  xdrlib ,sys

import xlrd
import xlwt


#打开excel文件
from orderedset._orderedset import OrderedSet


def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except:
        print("open failed")

def GoThroughCol(colNum,file):
    # 打开excel文件
    data = open_excel(file)
    # 打开第一张表
    table = data.sheets()[0]
    # 用于存放数据
    list = []
    # 将table中某列的数据读取并添加到list中
    list.extend(table.col_values(colNum))
    return list

def GoThroughRow(rowNum,file):
    # 打开excel文件
    data = open_excel(file)
    # 打开第一张表
    table = data.sheets()[0]
    # 用于存放数据
    list = []
    # 将table中某行的数据读取并添加到list中
    list.extend(table.row_values(rowNum))
    return list

def WriteToNewXl(list,RowNum,sheet1):
    j = 0
    for x in list:
        sheet1.write(RowNum, j, x) #在新sheet中的第RowNum行第j列写入读取到的x值
        j = j+1


#主函数
def main():

    usrIndex = 2
    passIndex = 3

    book = xlwt.Workbook()  # 创建一个Excel
    sheet1 = book.add_sheet('fuck')  # 在其中创建一个名为hello的sheet

    data = open_excel('test.xlsx')
    table = data.sheets()[0]
    nrows = table.nrows

    d = {}
    for i in range(nrows):
        list = []
        list.extend(table.row_values(i))
        key = list[usrIndex]
        if key not in d.keys():
            d[key] = []
            for j in range(len(list)):
                d[key].append(OrderedSet())#初始化d.value(),其中d.value()是一个以多个OrderdeSet()为内容的list
        for k in range(len(list)):
            if k == usrIndex:
                continue
            elif k == passIndex:
                values = list[passIndex].split()
                for v in values:
                    d[key][passIndex].add(v)
            else:
                d[key][k].add(list[k])

    for index, usr in enumerate(d.keys()):
        wlist = []
        for v in d[usr]:
            wlist.append(' '.join(v))
        wlist[usrIndex] = usr
        WriteToNewXl(wlist, index, sheet1)


    book.save('new.xls')  # 创建保存文件
if __name__ == "__main__":
    main()
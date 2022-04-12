# -*- coding: utf-8 -*-
# @Time: 2022/1/26 11:17
# @Author: shenyuming
import pandas
#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
    excel大量数据对比, 对比多列数据
"""


#导入模块 openpyxl
# !/usr/bin/env python
# -*- coding: utf-8 -*-
# 导入模块 openpyxl
import xlsxwriter
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
import copy

class conTrastExcel:
    def __init__(self):
        # 读取excel文件
        # 括号中的字符串为你要比较的两个excel的路径，注意用“/”
        wb_a = openpyxl.load_workbook('D:\Mine\scripeFile\GitHub\python-dataExcel\excel\excel1.xlsx')
        wb_b = openpyxl.load_workbook('D:\Mine\scripeFile\GitHub\python-dataExcel\excel\excel2.xlsx')
        wb_s = Workbook()
        self.ws = wb_s.active
        self.ws.title = '对比结果'

        self.wa = wb_a.worksheets[0]
        self.wb = wb_b.worksheets[0]

    print('-------第一张表---------')
    def excel1(self):
        lista = []
        for i in self.wa.iter_rows():
            list = []
            for cell in i:
                a = cell.value
                list.append(a)
            # print('一条数据',list)
            lista.append(list)
        # print('excel_a的全部值',lista)
        # print('lista的长度',len(lista))
        return lista

    print('-------第二张表---------')
    def excel2(self):
        listb = []
        for j in self.wb.iter_rows():
            list = []
            for cell in j:
                b = cell.value
                list.append(b)
            # print('一条数据：',list)
            listb.append(list)
        # print('excel_b的全部值',listb)
        # print('list_b的长度',len(listb))
        return listb

    def con(self,a,b):
        print("---excel1值： {}".format(a))
        print("---excel2值： {}".format(b))

        if len(a) > len(b):
            print("---len(a)值：{}".format(len(a)))
            c = len(a)
            d = len(b)
        else:
            print("---len(b)值：{}".format(len(b)))
            c = len(b)
            d = len(a)

        datas = []

        # c 为最大值；d 为最小值
        for i in range(c):
            print("---i值：{}".format(i))
            if i+1 <= d:
                tempList = []
                for j in range(len(a[i])):
                    if a[i][j] != b[i][j]:
                        tempList.append("{} 和 {} 不一致".format(a[i][j], b[i][j]))
                    else:
                        tempList.append(a[i][j])
                datas.append(tempList)
            else:
                if len(a) > len(b):
                    datas.append(a[i])
                else:
                    datas.append(b[i])
        print("---datas值：{}".format(datas))


        workbook = xlsxwriter.Workbook('excel/excel3.xlsx')  # 建立文件
        worksheet = workbook.add_worksheet()  # 建立sheet

        for index, item in enumerate(datas[0]):
            worksheet.write(0, index, '{}'.format(item))

        for index, item in enumerate(datas):
            print("---item： {}".format(item))
            worksheet.write(index, 0, '{}'.format(item[0]))
            worksheet.write(index, 1, '{}'.format(item[1]))
            worksheet.write(index, 2, '{}'.format(item[2]))



        workbook.close()



a = conTrastExcel().excel1()
b = conTrastExcel().excel2()
conTrastExcel().con(a,b)


























# for i in enumerate(a):
#     for j in enumerate(a):
#         if a != b:
#             resule = str(self.wa.values)+'和'+ str(self.wb.values)+'不一致'
#             self.ws.cell(i,j,resule)
#         else:
#             self.ws.cell(i,j,self.wa.cell(i,j).value)
# self.wb_s.save(f'E:/sym/共享文件夹/4.2取值_excel/11_布尔333333.xlsx')


# for i in range(wa.rows):
    #     for j in range(wa.ncols):
    #         if i == 0:
    #             ws.write(i, j, str(wb.cell(i, j).value))  # 第一行数据不比对
    #         else:
    #             if str(wa.cell(i, j).value) != str(wb.cell(i, j).value):  # 将两个excel表格中同行同列进行比较
    #                 # style = xlwt.easyxf('font:bold 1, color blue')  # 设置不匹配内容的字体及其颜色
    #                 result = str(wa.cell(i, j).value) + "和" + str(wb.cell(i, j).value) + "不一致"
    #                 ws.write(i, j, result)
    #             else:
    #                 ws.write(i, j, wa.cell(i, j).value)
    #     wbs.save('E:/sym/pi解析/对比结果2022218/Interpolated/1y/ScanClassInformation.xlsx')

    # # 定义一个方法来获取表格中某一列的内容，返回一个列表
    # # 将每一列输出为一个列表（temp表示列的名字）
    # def getIP(wb, temp):
    #     sheet = wb.active
    #     ip = []
    #     for cellobj in sheet[temp]:
    #         ip.append(cellobj.value)
    #     return ip
    #
    # # 想比较哪几列就输入那几列的名称
    # list1 = ['A', 'B', 'C', 'D', 'E']
    # list2 = []  # 用于存每列不同的值
    # differ1 = {}  # 第一个文件中每列不同的列表组成字典
    # differ2 = {}  # 第二个文件中每列不同的列表组成字典
    # for temp1 in list1:
    #     # 获得ip列表
    #     ip_a = getIP(wb_a, temp1)
    #     ip_b = getIP(wb_b, temp1)
    #     # 将两个列表转换成集合
    #     aa = set(ip_a)
    #     bb = set(ip_b)
    #     # 找出两个列表的不同行，并转换成列表
    #     difference = list(aa ^ bb)
    #     # 打印出列表中的元素
    #     # 到这一步，两个表格中不同的数据已经被找出来了
    #     # for i in difference:
    #     #     print (i)
    #
    #     # 将不同行高亮显示
    #     # print ("开始第一张表" + "----" *10)
    #     del list2[0:]
    #     a = wb_a.active[temp1]
    #     for cellobj in a:
    #         if cellobj.value in difference:
    #             # print (cellobj.value)
    #             cellobj.font = Font(color=colors.BLACK, italic=True, bold=True)
    #             cellobj.fill = PatternFill("solid", fgColor="DDDDDD")
    #             list2.append(cellobj.value)
    #     if list2 != []:
    #         differ1[temp1] = copy.deepcopy(list2)
    #     # print ("开始第二张表" + "----" *10)
    #     del list2[0:]
    #     b = wb_b.active[temp1]
    #     for cellobj in b:
    #         if cellobj.value in difference:
    #             # print (cellobj.value)
    #             cellobj.font = Font(color=colors.BLACK, italic=True, bold=True)
    #             cellobj.fill = PatternFill("solid", fgColor="DDDDDD")
    #             list2.append(cellobj.value)
    #     if list2 != []:
    #         differ2[temp1] = copy.deepcopy(list2)
    #
    # print(differ1.items())
    # print(differ2.items())
    # wb_a.save(f'E:/sym/共享文件夹/4.2取值_excel/B.xlsx')
    # wb_b.save(f'E:/sym/共享文件夹/4.2取值_excel/C.xlsx')

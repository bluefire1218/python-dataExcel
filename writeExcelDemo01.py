import xlwt
import xlsxwriter

""" xlsxwriter 方式写入 """
workbook = xlsxwriter.Workbook('excel/hello.xlsx') # 建立文件

worksheet = workbook.add_worksheet() # 建立sheet， 可以work.add_worksheet('employee')来指定sheet名，但中文名会报UnicodeDecodeErro的错误
worksheet.write(0, 0, 'Unformatted value') # 向A1写入

workbook.close()


""" xlwt 方式写入 """
# workbook = xlwt.Workbook(encoding = 'ascii')
# worksheet = workbook.add_sheet('My Worksheet')
# style = xlwt.XFStyle() # 初始化样式
# font = xlwt.Font() # 为样式创建字体
# font.name = 'Times New Roman'
# font.bold = True # 黑体
# font.underline = True # 下划线
# font.italic = True # 斜体字
# style.font = font # 设定样式
# worksheet.write(0, 0, 'Unformatted value') # 不带样式的写入
#
# worksheet.write(1, 0, 'Formatted value', style) # 带样式的写入
#
# workbook.save('excel/formatting.xls') # 保存文件



# # 创建一个workbook 设置编码
# workbook = xlwt.Workbook(encoding = 'utf-8')
# # 创建一个worksheet
# worksheet = workbook.add_sheet('My Worksheet')
#
# # 写入excel
# # 参数对应 行, 列, 值
# worksheet.write(1,0, label = 'this is test')
#
# # 保存
# workbook.save('excel/Excel_test.xls')

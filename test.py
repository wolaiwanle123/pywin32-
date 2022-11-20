# # def copy_table_from_excel_to_word():
# #     import time
# #     from win32com import client
# #
# #     excel = client.Dispatch('Excel.Application')
# #     word = client.Dispatch('Word.Application')
# #
# #     doc = word.Documents.Open(r"D:\03 python代码\01Pywin32批量写报告\test.docx")
# #     wb = excel.Workbooks.Open(r"D:\03 python代码\01Pywin32批量写报告\test.xlsx")
# #     sheet = wb.Worksheets(1)
# #
# #     tn = sheet.Cells(2, 1).Value
# #     print(tn)
# #     start_row = 2
# #     end_row = 2
# #
# #     while True:
# #         if sheet.Cells(start_row, 1).Value == '':
# #             print('finish')
# #             break
# #         if sheet.Cells(end_row + 1, 1).Value == tn:
# #             end_row += 1
# #         else:
# #             print(tn)
# #             word.Selection.InsertAfter('\n%s\n' % tn)
# #             word.Selection.InsertAfter('%s\n' % sheet.Cells(start_row, 2).Value)
# #             time.sleep(0.3)
# #             _ = word.Selection.MoveRight()
# #             time.sleep(0.3)
# #             _ = sheet.Range('C1:H1').Copy()
# #             word.Selection.PasteExcelTable(False, True, False)
# #             time.sleep(0.3)
# #             _ = sheet.Range('C%d:H%d' % (start_row, end_row)).Copy()
# #             word.Selection.PasteExcelTable(False, True, False)
# #             time.sleep(0.3)
# #
# #             start_row = end_row + 1
# #             end_row += 1
# #             tn = sheet.Cells(start_row, 1).Value
# #
# #     doc.Close()
# #     wb.Close()
# # copy_table_from_excel_to_word()
# # #
#
# #################################################################
# import time
# from win32com import client
# from win32com.client import constants
# excel = client.Dispatch('Excel.Application')
# word = client.Dispatch('Word.Application')
# doc = word.Documents.Open(r"D:\03 python代码\01Pywin32批量写报告\test.docx")
# wb = excel.Workbooks.Open(r"D:\03 python代码\01Pywin32批量写报告\test.xlsx")
#
# sheet=wb.Worksheets(1)#将变量sheet指向excel的第一张表
# sheet.Range('B2:H5').Copy()# 复制表中A1到B5的范围，A1为左上角的单元格坐标，B5为右下角的坐标
# # parag=doc.Paragraphs.Last#将变量parag指向word文档中最后一段的段尾
#
# 范围 = doc.Range()
# 范围.Find.Execute("表3-2 自动化检测路段路面状况 PQI 分项指标统计表")
# 范围.InsertParagraphAfter()
# doc.Range(范围.End, 范围.End).Paste()
# table = doc.Tables(3)
# table.Columns(1).Width = 120
# doc.Close()
# wb.Close()




from win32com.client import Dispatch
from win32com.client.gencache import EnsureDispatch
from win32com.client import constants
import openpyxl
from PIL import ImageGrab, Image    #PIL图像处理库
import time
import win32com.client as win32
# excel = EnsureDispatch('Excel.Application') #启动Excel程序
# excel.Visible = True  # 可视化                              #是否显示Excel界面
# excel.DisplayAlerts = False  # 是否显示警告
# wordorder = EnsureDispatch('Word.Application')  #启动word程序
# wordorder.Visible = True                        #程序可见
# lujing = r"D:\03 python代码\01Pywin32批量写报告\03 报告\07 2021年西宁市湟源县报告-0528.docx"  #word报告模板路径
# document = wordorder.Documents.Open(lujing)
# table4 = document.Tables(10)
# for i in range(2,table4.Rows.Count+1):
#
#     if float(92)>=90:
#         def rgbToInt(rgb):
#             colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
#             return colorInt
#         # table4.Cell(i,2).Range.Text.Interior.Color = rgbToInt((255, 255, 128))
#         # table4.Cell(i,2).Range.Font.Color = rgbToInt((120, 224, 57))
#         # table4.Cell(i, 2).Interior.Color = rgbToInt((120, 224, 57))
#         table4.Cell(i, 2).Range.Interior.Color = rgbToInt((120, 224, 57))
#         # table4.Cell(i, 2).Range.Text.Interior.Color = rgbToInt((120, 224, 57))
#         print(table4.Cell(i,2).Range.Text)
    # else :
    #     table4.Cell(i,2).Interior.PatternColor = (97,251,231)
    # elif float(table4.Cell(i,2).Range.Text)>=70:
    #     table4.Cell(i,2).Interior.PatternColor = RGB(224,238,115)
    # elif float(table4.Cell(i,2).Range.Text)>=60:
    #     table4.Cell(i,2).Interior.PatternColor = RGB(255,170,82)
    # else:
    #     table4.Cell(i,2).Interior.PatternColor = RGB(250,84,2)
import os
os.chdir(r'D:\03 python代码\01Pywin32批量写报告\02 Excel报表')
filename = os.listdir()
print(filename[0])
print(filename[1])
print(filename[2])
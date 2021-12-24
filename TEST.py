from openpyxl.styles import PatternFill, Alignment
from openpyxl.drawing.image import Image
from xlrd.xlsx import cooked_text  
import xlwt
import xlrd
from xlutils.copy import copy
import xlsxwriter
import datetime
import collections
import openpyxl

dataOT = openpyxl.load_workbook('OT.xlsx')
sheet_name_OT = dataOT.sheetnames[0]
sheet1_OT = dataOT[sheet_name_OT]

rows_OT = sheet1_OT.max_row #9
cols_OT = sheet1_OT.max_column #10

for col in range(4,cols_OT):
    time_OT = sheet1_OT.cell(col,10).coordinate
    day_OT = sheet1_OT.cell(col,11).coordinate

    x = sheet1_OT.cell(col,5).coordinate
    y = sheet1_OT.cell(col,6).coordinate
    sheet1_OT[time_OT].value = '=hour('+y+'-'+x+')'
    sheet1_OT[day_OT].value = '=day('+x+')'
    
dataOT.save('OT.xlsx')

# chamcong = xlrd.open_workbook('data.xlsx')
# data = chamcong.sheet_by_index(0)
# wb = copy(chamcong)
# w_sheet = wb.get_sheet(0)

# OT_approve = xlrd.open_workbook('OT_convert.xlsx')
# dataOT_approve = OT_approve.sheet_by_index(0)

# for i in range(dataOT_approve.nrows):
#     for j in range(data.nrows-3):
#         if dataOT_approve.cell_value(i, 0) == data.cell_value(j+3, 0) and dataOT_approve.cell_value(i, 2) == data.cell_value(j+3, 3):
#             w_sheet.write(j+3, 41,float( data.cell_value(j+3, 32))+float( data.cell_value(j+3, 33)))
#         else:
#             w_sheet.write(j+3, 41,float( data.cell_value(j+3, 32))+float( data.cell_value(j+3, 34)))

# # =========================== Mã hoá ca =====================
# for m in range(data.nrows-3):
#     if(float(data.cell_value(m+3, 41))>=1):
#         if(data.cell_value(m+3, 5) == "Cuoi tuan Ca Toi"):
#             w_sheet.write(m+3, 42, "CN")
#         elif(data.cell_value(m+3, 5) == "Cuoi tuan Ca Sang"):
#             w_sheet.write(m+3, 42, "CN")
#         else:
#             w_sheet.write(m+3, 42, "RCN")
#     if(float(data.cell_value(m+3, 25))>=5):
#         if(data.cell_value(m+3, 5) == "San xuat sang"):
#             w_sheet.write(m+3, 42, "A")
#         elif(data.cell_value(m+3, 5) == "San xuat toi"):
#             w_sheet.write(m+3, 42, "C")
#         elif(data.cell_value(m+3, 5) == "San xuat Ca C"):
#             w_sheet.write(m+3, 42, "B")
#         elif(data.cell_value(m+3, 5) == "Ca Hanh Chính"):
#             w_sheet.write(m+3, 42, "D")
#         else:
#             w_sheet.write(m+3, 42, "RR")
#     elif(float(data.cell_value(m+3, 25))<5 or float(data.cell_value(m+3, 25))>0):
#         if(data.cell_value(m+3, 5) == "San xuat sang"):
#             w_sheet.write(m+3, 42, "R")
#         elif(data.cell_value(m+3, 5) == "San xuat toi"):
#             w_sheet.write(m+3, 42, "R")
#         elif(data.cell_value(m+3, 5) == "San xuat Ca C"):
#             w_sheet.write(m+3, 42, "R")
#         elif(data.cell_value(m+3, 5) == "Ca Hanh Chính"):
#             w_sheet.write(m+3, 42, "RR")

# wb.save('data.xlsx')


# #Data
# all_rows_data = []
# for row in range(data.nrows):
#     curr_row = []
#     for col in range(data.ncols):
#         curr_row.append(data.cell_value(row, col))
#     all_rows_data.append(curr_row)

# chamcong1 = xlsxwriter.Workbook('data1.xlsx')
# data1 = chamcong1.add_worksheet()

# for row in range(len(all_rows_data)):
#     for col in range(len(all_rows_data[0])):
#         data1.write(row, col, all_rows_data[row][col])
# chamcong1.close()


# #Tạo template báo cáo

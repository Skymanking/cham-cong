import xlwt
import xlrd
from xlutils.copy import copy
import xlsxwriter
import datetime
import collections
import openpyxl
from openpyxl.styles import PatternFill, Alignment
from tqdm import tqdm, trange

# =========================== convert OT =====================
dataOT = xlrd.open_workbook('ott.xlsx')
ot = dataOT.sheet_by_index(0)
ot_convert = xlsxwriter.Workbook('OT_convert.xlsx')
add_sheet = ot_convert.add_worksheet()
for i in range(ot.nrows-3):
    id = ot.cell_value(i+3, 0)
    date = ot.cell_value(i+3, 4)
    date1 =ot.cell_value(i+3, 5)
    x =(datetime.datetime.strptime(date,"%Y-%m-%d %H:%M:%S"))
    y =(datetime.datetime.strptime(date1,"%Y-%m-%d %H:%M:%S"))
    timeOT = y - x
    hh, mm , ss = map(int, str(timeOT).split(':'))
    ot3 = hh + mm/60
    d = x.strftime("%Y-%m-%d")
    add_sheet.write(i,0, str(id))
    add_sheet.write(i,1, ot3)
    add_sheet.write(i,2, d)
ot_convert.close()

# =========================== Xoá data =====================

baocao_del = xlrd.open_workbook('baocao.xlsx')
data_baocao_del = baocao_del.sheet_by_index(0)

all_rows_baocao_del = []
for row in tqdm(range(data_baocao_del.nrows)):
    curr_row = []
    for col in range(data_baocao_del.ncols):
        curr_row.append(data_baocao_del.cell_value(row, col))
    all_rows_baocao_del.append(curr_row)

delete_baocao = xlsxwriter.Workbook('baocao.xlsx')
delete = delete_baocao.add_worksheet()

for row in range(len(all_rows_baocao_del)):
    for col in range(len(all_rows_baocao_del[0])):
        delete.write(row, col, "")
delete_baocao.close()


# =========================== Get data =====================

chamcong = xlrd.open_workbook('dataa.xlsx')
data = chamcong.sheet_by_index(0)
wb = copy(chamcong)
w_sheet = wb.get_sheet(0)

OT_approve = xlrd.open_workbook('OT_convert.xlsx')
dataOT_approve = OT_approve.sheet_by_index(0)

#Tạo template báo cáo

baocao = xlrd.open_workbook('baocao.xlsx')
mod_baocao = copy(baocao)
w_sheet_baocao = mod_baocao.get_sheet(0)

colect_id = []
colect_date = []
for id in range(data.nrows-3):
    colect_id.append(data.cell_value(id+3, 0))
    colect_date.append(data.cell_value(id+3, 3))
c = collections.Counter(colect_id)
d = collections.Counter(colect_date)
ID_baocao = c.keys()
date_baocao = d.keys()

colen2 = 0
for y in date_baocao:
    w_sheet_baocao.write(5, colen2+6, y)
    colen2 = colen2+2

colen = 0
for z in ID_baocao:
    for j in range(data.nrows -3):
        if z == data.cell_value(j+3, 0):
            w_sheet_baocao.write(colen+7, 0, data.cell_value(j+3, 0))
            w_sheet_baocao.write(colen+7, 1, data.cell_value(j+3, 1))
            w_sheet_baocao.write(colen+7, 2, data.cell_value(j+3, 2))
            w_sheet_baocao.write(colen+7, 3, data.cell_value(j+3, 3))
            w_sheet_baocao.write(colen+7, 4, data.cell_value(j+3, 4))
            w_sheet_baocao.write(colen+7, 5, data.cell_value(j+3, 5))
            colen = colen+1
            break

for o in range(1,1):
    w_sheet_baocao.write(5, o+67, o)

mod_baocao.save('baocao.xlsx')




# =========================== Duyệt OT =====================

for j in range(data.nrows-3):
    for i in range(dataOT_approve.nrows):
        if  dataOT_approve.cell_value(i, 0) == data.cell_value(j+3, 0) and dataOT_approve.cell_value(i, 2) == data.cell_value(j+3, 3):
            w_sheet.write(j+3, 35,float( data.cell_value(j+3, 29))+float( data.cell_value(j+3, 32)))
            break
        else:
            w_sheet.write(j+3, 35,float( data.cell_value(j+3, 29))+float( data.cell_value(j+3, 30))+float( data.cell_value(j+3, 31)))





baocao_1 = xlrd.open_workbook('baocao.xlsx')
data_baocao = baocao_1.sheet_by_index(0)

mod_day_baocao = copy(baocao_1)
w_sheet_baocao_day = mod_day_baocao.get_sheet(0)

# =========================== Mã hoá ca =====================
for m in range(data.nrows-3):
    if(float(data.cell_value(m+3, 30))>=1):
        if(data.cell_value(m+3, 5) == "Cuoi tuan Ca Toi"):
            w_sheet.write(m+3, 36, "CN")
        elif(data.cell_value(m+3, 5) == "Cuoi tuan Ca Sang"):
            w_sheet.write(m+3, 36, "CN")
        else:
            w_sheet.write(m+3, 36, "RCN")
    if(float(data.cell_value(m+3, 25))>=7):
        if(data.cell_value(m+3, 5) == "San xuat Sang"):
            w_sheet.write(m+3, 36, "A")
        elif(data.cell_value(m+3, 5) == "San xuat Toi"):
            w_sheet.write(m+3, 36, "C")
        elif(data.cell_value(m+3, 5) == "San xuat Ca C"):
            w_sheet.write(m+3, 36, "B")
        elif(data.cell_value(m+3, 5) == "Ca Hanh Chính"):
            w_sheet.write(m+3, 36, "D")
        else:
            w_sheet.write(m+3, 36, "RR1")
    elif(float(data.cell_value(m+3, 25))<7 or float(data.cell_value(m+3, 25))>0):
        if(data.cell_value(m+3, 5) == "San xuat Sang"):
            w_sheet.write(m+3, 36, "R")
        elif(data.cell_value(m+3, 5) == "San xuat Toi"):
            w_sheet.write(m+3, 36, "R")
        elif(data.cell_value(m+3, 5) == "San xuat Ca C"):
            w_sheet.write(m+3, 36, "R")
        elif(data.cell_value(m+3, 5) == "Ca Hanh Chính"):
            w_sheet.write(m+3, 36, "RR2")

wb.save('baocao.xlsx')

# =========================== Xử lý báo cáo =============================================
chamcong = xlrd.open_workbook('baocao.xlsx')
data = chamcong.sheet_by_index(0)
wb = copy(chamcong)
w_sheet = wb.get_sheet(0)


# # Chuyển ngày

for i in range(data_baocao.nrows-7):
    for j in range(data.nrows-3):
        if data_baocao.cell_value(i+7, 0) == data.cell_value(j+3, 0):
            for k in range(0,62,2):
                if data_baocao.cell_value(5, k+6) == data.cell_value(j+3, 3):
                        w_sheet_baocao_day.write(i+7,  k+6, data.cell_value(j+3, 36))
                        w_sheet_baocao_day.write(i+7,  k+7, data.cell_value(j+3, 35))

mod_day_baocao.save('baocao.xlsx')




# =========================== xử lý file mở k được =====================
#Data
all_rows_data = []
for row in range(data.nrows):
    curr_row = []
    for col in range(data.ncols):
        curr_row.append(data.cell_value(row, col))
    all_rows_data.append(curr_row)

chamcong1 = xlsxwriter.Workbook('data1.xlsx')
data1 = chamcong1.add_worksheet()

for row in range(len(all_rows_data)):
    for col in range(len(all_rows_data[0])):
        data1.write(row, col, all_rows_data[row][col])
chamcong1.close()

baocao_2 = xlrd.open_workbook('baocao.xlsx')
data_baocao = baocao_2.sheet_by_index(0)

# BAO CAO
all_rows_baocao = []
for row in range(data_baocao.nrows):
    curr_row = []
    for col in range(data_baocao.ncols):
        curr_row.append(data_baocao.cell_value(row, col))
    all_rows_baocao.append(curr_row)

baocao1 = xlsxwriter.Workbook('baocao1.xlsx')
data2 = baocao1.add_worksheet()

for row in range(len(all_rows_baocao)):
    for col in range(len(all_rows_baocao[0])):
        data2.write(row, col, all_rows_baocao[row][col])
baocao1.close()


baocao_1 = xlrd.open_workbook('baocao.xlsx')
data_baocao = baocao_1.sheet_by_index(0)

# Chuyen du lieu vao report


all_rows_baocao = []
for row in range(7 ,data_baocao.nrows):
    curr_row = []
    for col in range(data_baocao.ncols):
        curr_row.append(data_baocao.cell_value(row, col))
    all_rows_baocao.append(curr_row)


data_convert = openpyxl.load_workbook('Template_report.xlsx')
sheet_name_data_convert = data_convert.sheetnames[0]
sh_data_convert = data_convert[sheet_name_data_convert]

rows_data_convert= sh_data_convert.max_row #9
cols_data_convert = sh_data_convert.max_column #9

for row in range(1, len(all_rows_baocao)+1):
    for col in range(1, len(all_rows_baocao[0])+1):
        sh_data_convert.cell(row+10, col).value = all_rows_baocao[row-1][col-1]
data_convert.save("report1.xlsx")
print("done")
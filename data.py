import xlwt
import xlrd
from xlutils.copy import copy
import xlsxwriter
import datetime
import collections
import openpyxl
from openpyxl.styles import PatternFill, Alignment

# =========================== convert OT =====================
dataOT = xlrd.open_workbook('OT.xlsx')
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
    add_sheet.write(i,0, id)
    add_sheet.write(i,1, ot3)
    add_sheet.write(i,2, d)
ot_convert.close()

# =========================== Xoá data =====================

baocao_del = xlrd.open_workbook('baocao.xlsx')
data_baocao_del = baocao_del.sheet_by_index(0)

all_rows_baocao_del = []
for row in range(data_baocao_del.nrows):
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

chamcong = xlrd.open_workbook('data.xlsx')
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
    w_sheet_baocao.write(colen+7, 0, z)
    colen = colen+1

for o in range(1,23):
    w_sheet_baocao.write(5, o+67, o)

mod_baocao.save('baocao.xlsx')






# =========================== Duyệt OT =====================

for i in range(dataOT_approve.nrows):
    for j in range(data.nrows-3):
        if dataOT_approve.cell_value(i, 0) == data.cell_value(j+3, 0) and dataOT_approve.cell_value(i, 2) == data.cell_value(j+3, 3):
            w_sheet.write(j+3, 41,float( data.cell_value(j+3, 32))+float( data.cell_value(j+3, 33)))
        else:
            w_sheet.write(j+3, 41,float( data.cell_value(j+3, 32))+float( data.cell_value(j+3, 34)))





baocao_1 = xlrd.open_workbook('baocao.xlsx')
data_baocao = baocao_1.sheet_by_index(0)

mod_day_baocao = copy(baocao_1)
w_sheet_baocao_day = mod_day_baocao.get_sheet(0)


# =========================== Mã hoá ca =====================
for m in range(data.nrows-3):
    if(float(data.cell_value(m+3, 41))>=1):
        if(data.cell_value(m+3, 5) == "Cuoi tuan Ca Toi"):
            w_sheet.write(m+3, 42, "CN")
        elif(data.cell_value(m+3, 5) == "Cuoi tuan Ca Sang"):
            w_sheet.write(m+3, 42, "CN")
        else:
            w_sheet.write(m+3, 42, "RCN")
    if(float(data.cell_value(m+3, 25))>=5):
        if(data.cell_value(m+3, 5) == "San xuat sang"):
            w_sheet.write(m+3, 42, "A")
        elif(data.cell_value(m+3, 5) == "San xuat toi"):
            w_sheet.write(m+3, 42, "C")
        elif(data.cell_value(m+3, 5) == "San xuat Ca C"):
            w_sheet.write(m+3, 42, "B")
        elif(data.cell_value(m+3, 5) == "Ca Hanh Chính"):
            w_sheet.write(m+3, 42, "D")
        else:
            w_sheet.write(m+3, 42, "RR")
    elif(float(data.cell_value(m+3, 25))<5 or float(data.cell_value(m+3, 25))>0):
        if(data.cell_value(m+3, 5) == "San xuat sang"):
            w_sheet.write(m+3, 42, "R")
        elif(data.cell_value(m+3, 5) == "San xuat toi"):
            w_sheet.write(m+3, 42, "R")
        elif(data.cell_value(m+3, 5) == "San xuat Ca C"):
            w_sheet.write(m+3, 42, "R")
        elif(data.cell_value(m+3, 5) == "Ca Hanh Chính"):
            w_sheet.write(m+3, 42, "RR")

wb.save('data.xlsx')

# =========================== Xử lý báo cáo =============================================



# # Chuyển ngày

for i in range(data_baocao.nrows-7):
    for j in range(data.nrows-3):
        if data_baocao.cell_value(i+7, 0) == data.cell_value(j+3, 0):
            for k in range(0,62,2):
                if data_baocao.cell_value(5, k+6) == data.cell_value(j+3, 3):
                        w_sheet_baocao_day.write(i+7,  k+6, data.cell_value(j+3, 42))
                        w_sheet_baocao_day.write(i+7,  k+7, data.cell_value(j+3, 41))

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



data_report = openpyxl.load_workbook('baocao1.xlsx')
sheet_name = data_report.sheetnames[0]
sheet1 = data_report[sheet_name]

rows = sheet1.max_row #29
cols = sheet1.max_column #90
vi_tri = sheet1[7][cols-23].column_letter


sheet1['A5'].value = "BẢNG CHẤM CÔNG THÁNG"
sheet1.merge_cells('A5:B5') 
sheet1['C5'].value = "10"
sheet1['D5'].value = "Năm"
sheet1['E5'].value = "2021"
sheet1['A7'].value = "Mã NV"
for mer in range(7, cols -22,2):
    o1 = sheet1.cell(6, mer).coordinate
    o2 = sheet1.cell(6, mer +1).coordinate
    sheet1.merge_cells(o1+':'+o2)  
    sheet1[7][mer-1].value = "NC"
    sheet1[6][mer-1].alignment  = Alignment(horizontal='center')
    sheet1[7][mer].value = "TC"
    sheet1[6][mer].alignment  = Alignment(horizontal='center')
    can_chinh = sheet1.cell(1,mer).column_letter
    sheet1.column_dimensions[can_chinh].width = 5
    can_chinh1 = sheet1.cell(1,mer+1).column_letter
    sheet1.column_dimensions[can_chinh1].width = 5

for cot in range(0, rows-7):
    x = cot + 8
    sheet1[x][cols-22].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"a")+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"r0,a5")/2+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,a5")/2'
    sheet1[x][cols-21].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"d")+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"r0,d5")/2+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,d5")/2'
    sheet1[x][cols-20].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"b")+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"r0,b5")/2+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,b5")/2'
    sheet1[x][cols-19].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"c")+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"r0,c5")/2+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,c5")/2'
    sheet1[x][cols-18].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"KH")+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"TL")+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"L")'
    sheet1[x][cols-17].value = ''
    sheet1[x][cols-16].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"hh")'
    sheet1[x][cols-15].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"nt7")'
    sheet1[x][cols-14].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"P")+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,a5")/2+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,b5")/2+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,c5")/2+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,d5")/2'
    sheet1[x][cols-13].value = '=SUBTOTAL(9,BQ7:BY7)'
    sheet1[x][cols-12].value = '=SUM(J7,L7,N7,P7,R7,T7,X7,Z7,AB7,AD7,AF7,AH7,AL7,AN7,AP7,AR7,AT7,AV7,AZ7,BB7,BD7,BF7,BH7,BP7)'
    sheet1[x][cols-11].value = '=SUM(H7,V7,AJ7,AX7,BL7)'
    sheet1[x][cols-10].value = '=SUM(BJ7,BN7)'
    sheet1[x][cols-9].value = '=IF(BZ7>25,4,IF(BZ7>18,3,IF(BZ7>12,2,IF(BZ7>4,1,0))))-BX7-2+1'
    sheet1[x][cols-8].value = ''
    sheet1[x][cols-7].value = '=CD7*CI7'
    sheet1[x][cols-6].value = '=F7-BY7+1'
    sheet1[x][cols-5].value = '=CE7*16'
    sheet1[x][cols-4].value = ''
    sheet1[x][cols-3].value = '=SUM(CA12:CC12)'
    sheet1[x][cols-2].value = '=(SUM(IF(H13>4,H13-4,0),IF(J13>4,J13-4,0),IF(N13>4,N13-4,0),IF(P13>4,P13-4,0),IF(R13>4,R13-4,0),IF(T13>4,T13-4,0),IF(V13>4,V13-4,0),IF(X13>4,X13-4,0),IF(AB13>4,AB13-4,0),IF(AD13>4,AD13-4,0),IF(AF13>4,AH13-4,0),IF(AJ13>4,AJ13-4,0),IF(AL13>4,AL13-4,0),IF(AP13>4,AP13-4,0),IF(AR13>4,AR13-4,0),IF(AT13>4,AT13-4,0),IF(AV13>4,AV13-4,0),IF(AX13>4,AX13-4,0),IF(AZ13>4,AZ13-4,0),IF(BD13>4,BD13-4,0),IF(BF13>4,BF13-4,0),IF(BH13>4,BJ13-4,0),IF(BL13>4,BL13-4,0),IF(BN13>4,BN13-4,0))+CB13+CC13)'
    sheet1[x][cols-1].value = ''

for cot in range(7, rows+1):
    for mer in range(7, cols -22,2):
        # print(sheet1[cot][mer])
        sheet1[cot][mer].fill = PatternFill("solid", fgColor="F3F30B")
        sheet1[cot][mer].alignment  = Alignment(horizontal='center')

data_report.save('baocao1.xlsx')

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
import xlwings as xw


chamcong = xlrd.open_workbook('data.xlsx')
data = chamcong.sheet_by_index(0)
wb = copy(chamcong)
w_sheet = wb.get_sheet(0)

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

mod_baocao.save('baocao2.xlsx')

# BAO CAO

baocao_2 = xlrd.open_workbook('baocao2.xlsx')
data_baocao = baocao_2.sheet_by_index(0)
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
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

#=================== Time OT =========================



baocao_1 = xlrd.open_workbook('baocao.xlsx')
data_baocao = baocao_1.sheet_by_index(0)

# BAO CAO
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
data_convert.save("test1.xlsx")
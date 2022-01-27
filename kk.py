import xlwt
import xlrd
from xlutils.copy import copy
import xlsxwriter
import collections
import openpyxl
from openpyxl.styles import PatternFill, Alignment
from tqdm import tqdm, trange
from datetime import date, datetime


baocao_1 = xlrd.open_workbook('baocao.xlsx')
data_baocao = baocao_1.sheet_by_index(0)

nhanvien = xlrd.open_workbook("../cham-cong/input/nhanvien.xlsx")
sh_nhanvien = nhanvien.sheet_by_index(0)
for row_baocao in range(data_baocao.nrows):
    for row_nhanvien in range(sh_nhanvien.nrows):
        if data_baocao.cell_value(row_baocao, 0) == sh_nhanvien.cell_value(row_nhanvien, 0):
            print(data_baocao.cell_value(row_baocao, 0))
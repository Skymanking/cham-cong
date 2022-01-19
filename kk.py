import xlwt
import xlrd
from xlutils.copy import copy
import xlsxwriter
import datetime
import collections
import openpyxl
from openpyxl.styles import PatternFill, Alignment
from tqdm import tqdm, trange
from datetime import date, datetime

chamcong = xlrd.open_workbook('time.xlsx')
data = chamcong.sheet_by_index(0)
wb = copy(chamcong)
w_sheet = wb.get_sheet(0)

OT_approve = xlrd.open_workbook('OT_convert.xlsx')
dataOT_approve = OT_approve.sheet_by_index(0)

#Tạo template báo cáo

baocao = xlrd.open_workbook('baocao.xlsx')
mod_baocao = copy(baocao)
w_sheet_baocao = mod_baocao.get_sheet(0)

print(datetime.strptime(data.cell_value(4, 3), "%Y-%m-%d").weekday()==5)
print(data.cell_value(3, 5) == "San xuat Sang")
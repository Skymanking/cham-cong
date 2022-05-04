from asyncio.windows_events import NULL
from pandas import isnull
import xlwt
import xlrd
import khaibao
from xlutils.copy import copy
import xlsxwriter
import collections
import openpyxl
from openpyxl.styles import PatternFill, Alignment
from tqdm import tqdm, trange
from datetime import date, datetime
dataOT = xlrd.open_workbook('../cham-cong/input/Overtime_20220502032855.xlsx')
chamcong = xlrd.open_workbook('../cham-cong/convert/Du lieu hop nhat.xlsx', formatting_info=True)
data = chamcong.sheet_by_index(0)
ot = dataOT.sheet_by_index(0)
wb = copy(chamcong)
w_sheet = wb.get_sheet(0)
for m in tqdm(range(data.nrows-3)):
    for i in range(ot.nrows-3):
        if((data.cell_value(m+3, khaibao.MaNV) in ot.cell_value(i+3, khaibao.OTMaNV)) and (data.cell_value(m+3, khaibao.Ngay) in ot.cell_value(i+3, khaibao.OTStart))):
            date1 = ot.cell_value(i+3, khaibao.OTStart)
            date2 =ot.cell_value(i+3, khaibao.OTEnd)
            x =(datetime.strptime(date1,"%Y-%m-%d %H:%M:%S"))
            y =(datetime.strptime(date2,"%Y-%m-%d %H:%M:%S"))
            timeOT = y - x
            hh, mm , ss = map(int, str(timeOT).split(':'))
            ot3 = hh + mm/60
            w_sheet.write(m + 3,khaibao.Xinlamthem, ot3)
            print(ot.cell_value(i+3, khaibao.OTStart), data.cell_value(m+3, khaibao.MaNV), m, i)
wb.save('../cham-cong/convert/baocao.xlsx')
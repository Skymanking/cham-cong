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

data_convert = openpyxl.load_workbook('baocao1.xlsx')
sheet_name_data_convert = data_convert.sheetnames[0]
sh_data_convert = data_convert[sheet_name_data_convert]

rows_data_convert= sh_data_convert.max_row #9
cols_data_convert = sh_data_convert.max_column #9

print(rows_data_convert, cols_data_convert)

for row_convert in range(1, rows_data_convert + 1):
    for col_convert in range(1, cols_data_convert):
        print()



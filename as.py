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

k = input()
if k != "":
    h, m = k.split(":")
    x = float(h) + float(m)/60
    print(type(x))
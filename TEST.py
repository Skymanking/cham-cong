import tkinter.scrolledtext as sc
from tkinter import *
import tkinter.ttk as cm
from tkinter import filedialog

import xlwt
import xlrd
from xlutils.copy import copy
import xlsxwriter
import datetime
import collections
import openpyxl
from openpyxl.styles import PatternFill, Alignment

class Giaodien(Frame):

    def Clear(self):
        self.update()

    def Open_data(self):
        self.update()
        GD.filename_data = filedialog.askopenfilename()
        self.data_chamcong_link['text'] = "File đã chọn: " + GD.filename_data

    def Open_OT(self):
        self.update()
        GD.filename_OT = filedialog.askopenfilename()
        self.data_OT_link['text'] = "File đã chọn: " + GD.filename_OT

    def Open_nhanvien(self):
        self.update()
        GD.filename_nhanvien = filedialog.askopenfilename()
        self.data_nhanvien_link['text'] = "File đã chọn: " + GD.filename_nhanvien

    def Chon(self):
        self.update()
        data_convert = openpyxl.load_workbook('Template_report.xlsx')
        sheet_name_data_convert = data_convert.sheetnames[0]
        sh_data_convert = data_convert[sheet_name_data_convert]

        sh_data_convert.cell(6, 4).value = self.valuemonth.get()
        sh_data_convert.cell(6, 6).value = self.valueyear.get()
        data_convert.save("Template_report.xlsx")
        

        # =========================== convert OT =====================
        dataOT = xlrd.open_workbook(GD.filename_OT)
        ot = dataOT.sheet_by_index(0)
        ot_convert = xlsxwriter.Workbook('OT_convert.xlsx')
        add_sheet = ot_convert.add_worksheet()
        for i in range(ot.nrows-3):
            id = ot.cell_value(i+3, 0)
            date = ot.cell_value(i+3, 4)
            date1 =ot.cell_value(i+3, 5)
            x =(datetime.datetime.strptime(date,"%Y-%m-%d %H:%M:%S"))
            y =(datetime.datetime.strptime(date1,"%Y-%m-%d %H:%M:%S"))
            print(x)
            print(y)
            timeOT = y - x
            print(timeOT)
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

        chamcong = xlrd.open_workbook(GD.filename_data)
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

        wb.save(GD.filename_data)

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
        print("DONE")



    def __init__(self, master):
        super().__init__(master)
        self.Company = cm.Label(self, text = "HPT", font = ("Time New Roman", 30))

        self.Title = cm.Label(self, text = "BẢNG CHẤM CÔNG", font = ("Time New Roman", 24))
        self.Month = cm.Label(self, text = "THÁNG: ", font = ("Time New Roman", 24))

        self.month_title = cm.Label(self, text = "Tháng", font = ("Time New Roman", 12))
        self.valuemonth = cm.Combobox(self) 
        self.valuemonth['value'] = ("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12")

        self.year_title = cm.Label(self, text = "Năm", font = ("Time New Roman", 12))
        self.valueyear = cm.Combobox(self) 
        self.valueyear['value'] = ("2021","2022","2023","2024")
        
        self.data_chamcong = cm.Label(self, text = "Chọn file cham cong:", font = ("Time New Roman", 12))
        self.data_chamcong_link = cm.Label(self, text = "", font = ("Time New Roman", 12))
        self.button_chamcong=cm.Button(self, text = "Chọn file", command = self.Open_data)

        self.data_OT = cm.Label(self, text = "Chọn file OT:", font = ("Time New Roman", 12))
        self.data_OT_link = cm.Label(self, text = "", font = ("Time New Roman", 12))
        self.button_OT=cm.Button(self, text = "Chọn file", command = self.Open_OT)

        self.data_nhanvien = cm.Label(self, text = "Chọn file Nhân Viên:", font = ("Time New Roman", 12))
        self.data_nhanvien_link = cm.Label(self, text = "", font = ("Time New Roman", 12))
        self.button_nhanvien=cm.Button(self, text = "Chọn file", command = self.Open_nhanvien)



        self.Clear=cm.Button(self, text = "Clear data", command = self.Clear)
        self.Run = cm.Button(self, text = "RUN", command = self.Chon)
        master.bind("<Configure>", self.placeGD)

    def placeGD (self, even):
        self.update()
        selfW = self.winfo_width()
        selfH = self.winfo_height()

        self.Company.place(height = 100, width = 170, x = 30, y = 10)
        self.Title.place(height = 50, width = 350, x =270 , y = 50)
        self.Month.place(height = 40, width = 350, x =320 , y = 100)

        self.valuemonth.place(height = 30 , width = 80, x = 70, y = 160)
        self.month_title.place(height = 30, width = 50, x =10 , y = 160)

        self.valueyear.place(height = 30 , width = 80, x = 220, y = 160)
        self.year_title.place(height = 30, width = 50, x =170 , y = 160)

        self.data_chamcong.place(height = 40, width = 400, x =10 , y = 230)
        self.button_chamcong.place(height = 25, width = 70, x = 10, y = 265)
        self.data_chamcong_link.place(height = 40, width = 700, x =170 , y = 230)

        self.data_OT.place(height = 40, width = 400, x =10 , y = 300)
        self.button_OT.place(height = 25, width = 70, x = 10, y = 335)
        self.data_OT_link.place(height = 40, width = 700, x =170 , y = 300)

        self.data_nhanvien.place(height = 40, width = 400, x =10 , y = 370)
        self.button_nhanvien.place(height = 25, width = 70, x = 10, y = 405)
        self.data_nhanvien_link.place(height = 40, width = 700, x =170 , y = 370)


        self.Run.place(height = 100, width = 100, x = 680, y = 480)

        self.Clear.place(height = 40, width = 100, x =55 , y = 520)

 
GD = Tk()
GD.title("GROUP 4")
GD.geometry('800x600+0+0')
GD.configure(bg = 'red')
sky = Giaodien(GD)
sky.place(relwidth = 1, relheight = 1)
GD.mainloop()

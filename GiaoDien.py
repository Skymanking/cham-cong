from queue import Empty
import tkinter.scrolledtext as sc
from tkinter import *
from tkinter.ttk import *
import tkinter.ttk as cm
from tkinter import filedialog
from test import xuly
from datetime import datetime
dem = 0
day_now = datetime.today()
class Giaodien(Frame):

    def Clear(self):
        self.update()
        self.data_nhanvien_link['text'] = " " 
        self.data_OT_link['text'] = " " 
        self.data_chamcong_link['text'] = " " 

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
        def timerun():
            global dem
            dem += 1
            self.countdown['text'] = "Thoi gian: " + str(dem)
            GD.after(1000, timerun)
        timerun()
        xuly(GD.filename_data, GD.filename_OT, GD.filename_nhanvien, GD.text_nam.get(), GD.text_thang.get(), self.holiday_link.get())
        self.thongbao['text'] = "XONG "
    
    def __init__(self, master):
        super().__init__(master)
        GD.text_thang = StringVar()
        GD.text_nam = StringVar()
        GD.Holiday = StringVar()
        self.Company = cm.Label(self, text = "HPT", font = ("Time New Roman", 30))

        self.Title = cm.Label(self, text = "BẢNG CHẤM CÔNG", font = ("Time New Roman", 24))
        self.Month = cm.Label(self, text = "THÁNG: " + str(day_now.month - 1), font = ("Time New Roman", 24))

        self.month_title = cm.Label(self, text = "Tháng", font = ("Time New Roman", 12))
        self.valuemonth = cm.Combobox(self, textvariable= GD.text_thang) 
        self.valuemonth['value'] = ("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12")

        self.year_title = cm.Label(self, text = "Năm", font = ("Time New Roman", 12))
        self.valueyear = cm.Combobox(self, textvariable= GD.text_nam) 
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

        self.holiday = cm.Label(self, text = "Ngày lễ: (phân biệt bởi dấu ',')", font = ("Time New Roman", 12))
        self.holiday_link = Entry(GD, width= 500)

        self.thongbao = cm.Label(self, text = "", font = ("Time New Roman", 36))

        self.countdown = cm.Label(self, text = "", font = ("Time New Roman", 12))

        self.Clear=cm.Button(self, text = "Clear data", command = self.Clear)
        self.Run = cm.Button(self, text = "RUN", command = self.Chon)
        master.bind("<Configure>", self.placeGD)

    def placeGD (self, even):
        self.update()

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

        self.holiday.place(height = 40, width = 400, x =10 , y = 435)
        self.holiday_link.place(height = 30, width = 500, x =10 , y = 465)
 
        self.thongbao.place(height = 60, width = 700, x =250 , y = 450)

        self.countdown.place(height = 60, width = 700, x =250 , y = 500)

        self.Run.place(height = 100, width = 100, x = 680, y = 480)

        self.Clear.place(height = 40, width = 100, x =55 , y = 520)

GD = Tk()
GD.title("CHAM CONG")
GD.geometry('800x600+0+0')
GD.configure(bg = 'red')
sky = Giaodien(GD)
sky.place(relwidth = 1, relheight = 1)

GD.mainloop()
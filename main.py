import tkinter as tk
import tkinter.ttk as ttk
import global_Var as gvar
from tkinter import messagebox, StringVar, ttk
import datetime
import Catia_File
import main_program as mprog

# from pymssql import programmingError
data = []
Name = []
import serial, time, io
import pymssql
from string import punctuation

folderdir = "C:\\xampp\\htdocs\\CATIA\\CATIA_image_file\\"

for key in gvar.Path_Name:
    Name.append(key)


class Main_Frame:
    def __init__(self, master):
        # FRAME Settings(宣告此表單；最上方標題列文字；背景顏色；位置)
        self.master = master
        self.master.title('練習使用人機介面')
        self.master.configure(background='white')
        self.master.grid_columnconfigure(1, weight=0)
        self.combo()

        # 確認是否已有現成資料，若有則保留，無則空白
        try:
            self.entryvar1 = tk.StringVar()
            self.entryvar1.set(data[0])
            self.entryvar2 = tk.StringVar()
            self.entryvar2.set(data[1])
            self.entryvar3 = tk.StringVar()
            self.entryvar3.set(data[2])
        except:
            self.entryvar1 = tk.StringVar()
            self.entryvar2 = tk.StringVar()
            self.entryvar3 = tk.StringVar()

        self.entryvar4 = tk.StringVar()
        self.geo_values = ('desk', '3D_printing_machine', 'blackbroad', 'box', 'cabinet', 'part', 'shelf')

        # ELEMENT Settings(標題設定：所在視窗,文字內容,字型及字體大小,位置)
        self.text1 = tk.Label(self.master, text='20220128', font=("Arial", 16), anchor='center')
        self.text1.grid(row=0, column=0, columnspan=5, padx=5, sticky='we')
        self.text2a = tk.Label(self.master, text='1.型號(Scale): ', font=("Arial", 12), anchor='w')
        self.text2a.grid(row=1, column=0, columnspan=5, sticky='w')
        self.text2b = tk.Label(self.master, text='2.版型(version): ', font=("Arial", 12), anchor='w')
        self.text2b.grid(row=3, column=0, columnspan=5, sticky='w')
        self.text2c = tk.Label(self.master, text='3. 長(Length): ', font=("Arial", 12), anchor='w')
        self.text2c.grid(row=5, column=0, columnspan=5, sticky='w')
        self.text2d = tk.Label(self.master, text='4. 寬(Width): ', font=("Arial", 12), anchor='w')
        self.text2d.grid(row=7, column=0, columnspan=5, sticky='w')
        self.text2e = tk.Label(self.master, text='5. 高(Height): ', font=("Arial", 12), anchor='w')
        self.text2e.grid(row=9, column=0, columnspan=5, sticky='w')
        self.entry1 = tk.Entry(self.master, textvariable=self.entryvar1)
        self.entry1.grid(row=6, column=0, padx=20, pady=20)
        self.entry2 = tk.Entry(self.master, textvariable=self.entryvar2)
        self.entry2.grid(row=8, column=0, padx=20, pady=20)
        self.entry3 = tk.Entry(self.master, textvariable=self.entryvar3)
        self.entry3.grid(row=10, column=0, padx=20, pady=20)
        self.label1 = tk.Label(self.master, text='6. 架子層數(Shelf_Plane): ', font=("Arial", 12), anchor='w')
        # self.label1.grid(row=11, column=0, columnspan=5, sticky='w')
        self.label1.grid_forget()
        self.text2f = tk.Entry(self.master, textvariable=self.entryvar4)
        self.text2f.grid_forget()
        # lambda:[函式,函式](button多功能使用方式)
        self.button3 = tk.Button(self.master, text='匯入參數\nSet Value', command=lambda: [
            mprog.File_open(self.box2.get(), folderdir + "%s\\" % (self.box.get())), self.confirm_data(),
            self.Change_Parameter_In_Catia()])
        self.button3.grid(row=10, column=4, padx=5)

    def selected(self, e):
        self.box2.configure(values=gvar.Path_Name[self.box_value.get()])
        self.box2.current(0)
        self.geo_confirm()

    def combo(self):
        # First Combobox
        self.box_value = StringVar()
        self.box = ttk.Combobox(self.master, textvariable=self.box_value,
                                state='readonly')
        # Modification 1
        self.box.bind("<<ComboboxSelected>>", self.selected)
        self.box['values'] = Name
        self.box.current(0)
        self.box.grid(row=2, column=0, padx=20, pady=20)

        # Second Combobox
        self.box_value2 = StringVar()
        self.box2 = ttk.Combobox(self.master, textvariable=self.box_value2,
                                 state='readonly')
        self.box2['values'] = gvar.Path_Name[self.box.get()]

        self.box2.grid(row=4, column=0, padx=20, pady=20)

    # 新增架子欄位
    def geo_confirm(self):
        comboflag = self.box_value.get()
        if comboflag == self.geo_values[6]:
            self.label1.grid(row=11, column=0)
            self.text2f.grid(row=12, column=0)
        else:
            self.label1.grid_forget()
            self.text2f.grid_forget()



    # 設定參數，將輸入的參數放至主程式串列
    def confirm_data(self):
        try:
            # 確定是否有輸入數值進入方框中
            for item in range(0, len(data)):
                del data[-1]
            if self.entryvar1.get == '' or self.entryvar2.get() == '' or self.entryvar3.get() == '':
                raise ValueError('No Input')
                # 利用判斷式確定數值的大小

            try:
                if str(self.box.get()) == "shelf":
                    param = [self.box.get(), self.box2.get(), self.entryvar1.get(), self.entryvar2.get(),
                             self.entryvar3.get(), self.entryvar4.get()]
                else:
                    param = [self.box.get(), self.box2.get(), self.entryvar1.get(), self.entryvar2.get(),
                             self.entryvar3.get(), 0]
            except AttributeError:
                param = ['self.entryvars1.get()', 'self.entryvars2.get()', 'self.entryvars3.get()']
            for i in range(0, len(param)):
                data.append(param[i])
            else:
                data.append(True)
                temp = data.copy()
                messagebox.showinfo("選擇型號為：{}".format(temp[0]),
                                    "選擇版型為：{0[1]}\n選擇長為：{0[2]}\n選擇寬為：{0[3]}\n選擇高為：{0[4]}".format(temp))
                # opd.data = data.copy()

        # 顯示錯誤：沒有輸入
        except ValueError:
            tk.messagebox.showerror('WARNING', 'no Input!!', parent=self.master)

        try:
            # 輸入至資料庫
            conn = pymssql.connect(
                server='192.168.0.168',
                user='sa', password='mbil2235',
                database='catia')
            cursor = conn.cursor(as_dict=True)
            print(temp)
            cursor.execute(
                "INSERT INTO [dbo].[CATIA_Part] ([Path],[FileName],[Length],[Width],[Height],[Shelf_Plane]) VALUES(%s,%s,%s,%s,%s,%s)",
                tuple(temp))
            conn.commit()
            print(cursor.rowcount, "Record inserted successfully into CATIA_Part table")
            cursor.close()

        except pymssql.Error as error:
            print("Failed to insert record into CATIA_Part table {}".format(error))

        print(data)

    def Change_Parameter_In_Catia(self):
        gvar.Length = int(data[2])
        gvar.Width = int(data[3])
        gvar.Height = int(data[4])
        gvar.Shelf_Plane = int(data[5])

        mprog.Param_change("Length", gvar.Length)
        mprog.Param_change("Width", gvar.Width)
        mprog.Param_change("Height", gvar.Height)
        if str(self.box.get()) == "shelf":
            mprog.Param_change("Shelf_Plane", gvar.Shelf_Plane)
        else:
            pass

# mprog.Param_change("Height",data[0])

# 讓程式執行這個主程式，不執行其他瑣碎程式
def main():
    root = tk.Tk()
    app = Main_Frame(root)
    root.resizable(0, 0)
    root.mainloop()


if __name__ == '__main__':
    main()

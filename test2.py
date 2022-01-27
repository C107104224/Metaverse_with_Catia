import tkinter as tk
import tkinter.ttk as ttk
import global_Var as gvar
from tkinter import messagebox, StringVar, ttk
import datetime

# from pymssql import programmingError
data = []
Name = []
import serial, time, io
import pymssql
from string import punctuation

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
            # self.combovar0 = StringVar()
            # self.combovar1 = StringVar()
            self.entryvar1 = tk.StringVar()
            self.entryvar1.set(data[0])
            self.entryvar2 = tk.StringVar()
            self.entryvar2.set(data[1])
            self.entryvar3 = tk.StringVar()
            self.entryvar3.set(data[2])
        except:
            # self.combovar0 = StringVar()
            # self.combovar1 = StringVar()
            self.entryvar1 = tk.StringVar()
            self.entryvar2 = tk.StringVar()
            self.entryvar3 = tk.StringVar()

        # ELEMENT Settings(標題設定：所在視窗,文字內容,字型及字體大小,位置)
        self.text1 = tk.Label(self.master,
                              text='20220126', font=("Arial", 16),
                              anchor='center').grid(row=0, column=0, columnspan=5, padx=5, sticky='we')
        self.text2c = tk.Label(self.master, text='1.型號(Scale): ', font=("Arial", 12)
                               , anchor='w').grid(row=1, column=0, columnspan=5, sticky='w')

        # self.combobox0 = ttk.Combobox(self.master, textvariable=self.combovar0, state='readonly')
        # self.combobox0.bind("<<ComboboxSelected>>",self.selected)
        # self.combobox0['values'] = Name
        # self.combobox0.current(0)
        # self.combobox0.grid(row=2, column=0, padx=20, pady=20)
        self.text2d = tk.Label(self.master, text='2.版型(version): ', font=("Arial", 12)
                               , anchor='w').grid(row=3, column=0, columnspan=5, sticky='w')

        # self.combobox1 = ttk.Combobox(self.master, textvariable=self.combovar1, state='readonly')
        # self.combobox1['values'] = gvar.Path_Name[self.combobox0.get()]
        # self.combobox1.grid(row=4, column=0, padx=20, pady=20)
        self.text2 = tk.Label(self.master, text='3. 長(Length): ', font=("Arial", 12)
                              , anchor='w').grid(row=5, column=0, columnspan=5, sticky='w')
        self.text2a = tk.Label(self.master, text='4. 寬(Width): ', font=("Arial", 12),
                               anchor='w').grid(row=7, column=0, columnspan=5, sticky='w')
        self.text2b = tk.Label(self.master, text='5. 高(Height): ', font=("Arial", 12),
                               anchor='w').grid(row=9, column=0, columnspan=5, sticky='w')
        self.entry1 = tk.Entry(self.master, textvariable=self.entryvar1)
        self.entry1.grid(row=6, column=0, padx=20, pady=20)
        self.entry2 = tk.Entry(self.master, textvariable=self.entryvar2)
        self.entry2.grid(row=8, column=0, padx=20, pady=20)
        self.entry3 = tk.Entry(self.master, textvariable=self.entryvar3)
        self.entry3.grid(row=10, column=0, padx=20, pady=20)
        self.button3 = tk.Button(self.master, text='匯入參數\nSet Value',
                                 command=self.confirm_data).grid(row=6, column=4, padx=5)

    def selected(self, e):
        self.box2.configure(values=gvar.Path_Name[self.box_value.get()])
        self.box2.current(0)

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
                param = [self.box.get(), self.box2.get(), self.entryvar1.get(), self.entryvar2.get(),
                         self.entryvar3.get()]
            except AttributeError:
                param = ['self.entryvars1.get()', 'self.entryvars2.get()', 'self.entryvars3.get()']
            for i in range(0, len(param)):
                data.append(param[i])
            else:
                data.append(True)
                print(data)
                temp = data.copy()
                print(temp[1:3])
                messagebox.showinfo("選擇型號為：{}".format(temp[0]),
                                    "選擇版型為：{0[1]}\n選擇長為：{0[2]}\n選擇寬為：{0[3]}\n選擇高為：{0[4]}".format(temp))
                # opd.data = data.copy()

        # 顯示錯誤：沒有輸入
        except ValueError:
            tk.messagebox.showerror('WARNING', 'no Input!!', parent=self.master)

        try:
            conn = pymssql.connect(
                server='192.168.0.168',
                user='sa', password='mbil2235',
                database='catia')
            cursor = conn.cursor(as_dict=True)

            cursor.execute(
                "INSERT INTO [dbo].[CATIA_Part] ([Path],[FileName],[Length],[Width],[Height]) VALUES(%s,%s,%s,%s,%s)",
                tuple(temp))
            conn.commit()
            print(cursor.rowcount, "Record inserted successfully into CATIA_Part table")
            cursor.close()

        except pymssql.Error as error:
            print("Failed to insert record into CATIA_Part table {}".format(error))


# 讓程式執行這個主程式，不執行其他瑣碎程式
def main():
    root = tk.Tk()
    app = Main_Frame(root)
    root.resizable(0, 0)
    root.mainloop()


if __name__ == '__main__':
    main()

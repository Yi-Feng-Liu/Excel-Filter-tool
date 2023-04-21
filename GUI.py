from tkinter import *
import pandas as pd
from PIL import Image, ImageTk 
import cv2 as cv
from tkinter import filedialog, messagebox
from tkinter.font import Font
import threading
from tkinter import ttk
import numpy as np
from process import Judge_Metabolic_Syndrome
import webbrowser


class EntryWithPlaceholder(Entry):
    def __init__(self, master=None, placeholder="PLACEHOLDER", color='grey', width=30):
        super().__init__(master)
        self.placeholder = placeholder
        self.placeholder_color = color
        self['width'] = width
        self.default_fg_color = self['fg']

        self.bind("<FocusIn>", self.foc_in)
        self.bind("<FocusOut>", self.foc_out)

        self.put_placeholder()

    def put_placeholder(self):
        self.insert(0, self.placeholder)
        self['fg'] = self.placeholder_color

    def foc_in(self, *args):
        if self['fg'] == self.placeholder_color:
            self.delete('0', 'end')
            self['fg'] = self.default_fg_color

    def foc_out(self, *args):
        if not self.get():
            self.put_placeholder()
            
class Excel_GUI:
    def __init__(self, root):
        self.root = root
        # set background color
        self.bgcolor = "#d2efd3"
        self.empty_ls = []
        # set windows size
        self.windows_w = 500
        self.windows_h = 500
        self.font = Font(family="微軟正黑體", size=12)

        # set windows position
        self.geometry = root.geometry(f"{str(self.windows_w)}x{str(self.windows_h)}+800+300")
        
        # inintialize label
        self.greet = Label(root, text="", height="3", bg=self.bgcolor, font=self.font)
        self.sheets_names_combobox_label = Label(root, text="", font=self.font, bg=self.bgcolor)
        self.years_combobox_label = Label(root, text="", font=self.font, bg=self.bgcolor)
        self.save_sheet_name_label = Label(root, text="", font=self.font, bg=self.bgcolor)
        self.select_button = Button(root, text="", command=self.openFile)
        self.label = Label(root, text="", font=self.font, bg=self.bgcolor)
        # can't not sesize wiondow
        self.resizable = root.resizable(width=0, height=0)
        self.attributes = root.attributes('-alpha', 1)

        # set the app icon
        self.iconimg = self.getIcon('icon.jpg')
        root.iconphoto(True, self.iconimg)
        root.configure(bg=self.bgcolor)
        # set app title
        self.title = root.title("Excel Filter Tool")
        
        # set radio box to select choose file type
        self.select_input_file_type()
        
        # set decoraction position
        self.origin_img_path = self.resize('link_img.jpg', resize_w=100, resize_h=100, changeBG=True)
        self.decoration = self.getIcon(self.origin_img_path)
        width = self.decoration.width()
        height = self.decoration.height()
        self.pikachu = Button(root, image=self.decoration, command=self.linked_to_github)
        self.pikachu.place(x=self.windows_w - width, y=self.windows_h - height)

        # set the confirm button using image
        self.button_img = self.resize('confirm.jpg', resize_w=70, resize_h=60, changeBG=False)
        self.button_img = Image.open(self.button_img)
        self.comfirm_button_img = ImageTk.PhotoImage(self.button_img) 
        self.comfirm_button = Button(root, text="確認", command=self.processExcel, image=self.comfirm_button_img)
        self.comfirm_button.pack()
        self.comfirm_button.place(x=self.windows_w-250, y=self.windows_h-50, anchor='center')
        self.comfirm_button["state"] = "disabled"

        self.Copyright_label = Label(self.root, text='Copyright © 2023 YF Liu. All rights reserved.', font=("微軟正黑體", 7) , bg=self.bgcolor)
        self.Copyright_label.place(x=0, y=self.windows_h-20)
    
    def select_input_file_type(self):
        # global 
        methods = [
            ('電子檔健檢資料', 1, self.choose_e_file),
            ('手Key健檢資料', 2, self.choose_paper_file)
        ]
        self.v = IntVar()
        self.v.set(0)
    
        self.e_file = Radiobutton(self.root, text=methods[0][0], variable=self.v, value=methods[0][1], command=methods[0][2],bg=self.bgcolor, font=self.font)
        self.e_file.place(x=160, y=30, anchor='center')
        self.paper_file = Radiobutton(self.root, text=methods[1][0], variable=self.v, value=methods[1][1], command=methods[1][2], bg=self.bgcolor, font=self.font)
        self.paper_file.place(x=500-160, y=30, anchor='center')

    def initialize_button_and_label(self):
        self.greet.destroy()
        self.years_combobox_label.destroy()
        self.sheets_names_combobox_label.destroy()
        self.select_button.destroy()
        self.label.destroy()
        

    def choose_e_file(self):
        self.initialize_button_and_label()
        
        self.greet = Label(root, text="類型", height="3", bg=self.bgcolor, font=self.font)
        self.greet.place(x=250, y=85, anchor='center')
        

        # create combobox
        self.options = ['代謝症候群']
        self.box = ttk.Combobox(root, values=self.options)
        # use default
        self.box.current()
        self.box.place(x=250, y=110, anchor='center')

        # set choose file button
        self.select_button = Button(root, text="選擇檔案", command=self.openFile)
        self.select_button.place(x=225, y=140)

        # update choose file path
        self.label = Label(root, text='', font=self.font, bg=self.bgcolor)
        self.label.place(x=250, y=210, anchor='center')
        
        self.combobox_xlabel_place = 47
        self.combobox_place = 150
        self.sheets_names_combobox_label = Label(root, text='存檔位置:', font=self.font, bg=self.bgcolor)
        self.sheets_names_combobox_label.place(x=self.combobox_xlabel_place, y=self.windows_h-220)
        self.sheets_names_combobox = ttk.Combobox(self.root, values=None, width=27)
        self.sheets_names_combobox.place(x=self.combobox_place, y=self.windows_h-217)

        # set years combobox
        self.years_combobox_label = Label(root, text='儲存工作表:', font=self.font, bg=self.bgcolor)
        self.years_combobox_label.place(x=self.combobox_xlabel_place, y=self.windows_h-190)
        self.years_combobox = ttk.Combobox(self.root, values=None, width=27)
        self.years_combobox.place(x=self.combobox_place, y=self.windows_h-187)
        # set entry & save sheet
        self.save_sheet_name_label = Label(root, text='Sheet Name:', font=self.font, bg=self.bgcolor)
        self.save_sheet_name_label.place(x=self.combobox_xlabel_place, y=self.windows_h-160)
        self.defalut_placeholder_message = '輸入要儲存的Sheet名稱'
        self.save_sheet_name_entry = EntryWithPlaceholder(master=root, placeholder=self.defalut_placeholder_message, width=30)
        self.save_sheet_name_entry.place(x=self.combobox_place, y=self.windows_h-157)

        # set sheet comfirm button
        self.sheet_comfirm_button = Button(root, text="Submit", command=self.get_sheet_name)
        self.sheet_comfirm_button.pack()
        self.sheet_comfirm_button.place(x=390, y=self.windows_h-208, anchor='center')
        self.sheet_comfirm_button["state"] = "active"
    
    def choose_paper_file(self):
        self.initialize_button_and_label()
        
        self.greet = Label(root, text="選擇判斷類型", height="3", bg=self.bgcolor, font=self.font)
        self.greet.place(x=250, y=85, anchor='center')
        
        # create combobox
        self.options = ['代謝症候群']
        self.box = ttk.Combobox(root, values=self.options)

        # use default
        self.box.current()
        self.box.place(x=250, y=110, anchor='center')

        # set choose file button
        self.select_button = Button(root, text="選擇檔案", command=self.openFile)
        self.select_button.place(x=225, y=140)

        # update choose file path
        self.label = Label(root, text='', font=self.font, bg=self.bgcolor)
        self.label.place(x=250, y=210, anchor='center')
        
        self.combobox_xlabel_place = 47
        self.combobox_place = 150
        self.sheets_names_combobox_label = Label(root, text='Select Sheet:', font=self.font, bg=self.bgcolor)
        self.sheets_names_combobox_label.place(x=self.combobox_xlabel_place, y=self.windows_h-220)
        self.sheets_names_combobox = ttk.Combobox(self.root, values=None, width=27)
        self.sheets_names_combobox.place(x=self.combobox_place, y=self.windows_h-217)

        # set years combobox
        self.years_combobox_label = Label(root, text='Select Years:', font=self.font, bg=self.bgcolor)
        self.years_combobox_label.place(x=self.combobox_xlabel_place, y=self.windows_h-190)
        self.years_combobox = ttk.Combobox(self.root, values=None, width=27)
        self.years_combobox.place(x=self.combobox_place, y=self.windows_h-187)
        # set entry & save sheet
        self.save_sheet_name_label = Label(root, text='Sheet Name:', font=self.font, bg=self.bgcolor)
        self.save_sheet_name_label.place(x=self.combobox_xlabel_place, y=self.windows_h-160)
        self.defalut_placeholder_message = '輸入要儲存的Sheet名稱'
        self.save_sheet_name_entry = EntryWithPlaceholder(master=root, placeholder=self.defalut_placeholder_message, width=30)
        self.save_sheet_name_entry.place(x=self.combobox_place, y=self.windows_h-157)

        # set sheet comfirm button
        self.sheet_comfirm_button = Button(root, text="Submit", command=self.get_sheet_name)
        self.sheet_comfirm_button.pack()
        self.sheet_comfirm_button.place(x=390, y=self.windows_h-208, anchor='center')
        self.sheet_comfirm_button["state"] = "active"
        
    
    def resize(self, imgpath:str, resize_w:int, resize_h:int, changeBG=True):
        image_name = imgpath.split("\\")[-1].split(".")[0]
        
        origin_img = cv.imread(imgpath)
        # replace the iamge white background to #d2efd3
        if changeBG:
            origin_img[np.where((origin_img == [255, 255, 255]).all(axis=2))] = [211, 239, 211]
        resize_img = cv.resize(origin_img, (resize_w, resize_h))
        save_path = f'{image_name}_resize.jpg'
        cv.imwrite(save_path, resize_img)
        return save_path


    def finishInfo(self):
        messagebox.showinfo(title="通知", message="OK")


    def openFile(self):
        global filepath
        global tabs
        
        filepath = filedialog.askopenfilename(title="選擇檔案", filetypes=[("Excel files", ".xlsx .xls")])
        if filepath:
            self.label.config(text=f'選取的檔案路徑為\n {filepath}')
            self.comfirm_button["state"] = "active"
            # get list of sheet names 
            tabs = pd.ExcelFile(filepath).sheet_names 
            self.sheets_names_combobox['values'] = tabs
        else:
            self.sheets_names_combobox['values'] = self.empty_ls
            self.label.config(text=f'沒有選取的檔案!')
            self.comfirm_button["state"] = "disabled"
        

    def get_sheet_name(self):
        global sheetName
        sheetName = self.sheets_names_combobox.get()
        try:
            df = pd.read_excel(filepath, sheet_name=sheetName, engine = 'openpyxl')
            df_years_list = list(dict.fromkeys(df['年度代碼'].tolist()))
            self.years_combobox['values'] = df_years_list     
        except Exception as Error:
            self.years_combobox['values'] = self.empty_ls
            messagebox.showerror(title='Error', message=Error)


    def processExcel(self):
        select_year = self.years_combobox.get()
        save_sheet = self.save_sheet_name_entry.get()
        
        if save_sheet in tabs:
            messagebox.showerror(title='Error', message='The Same of Sheet Names That is Not Allow')
        elif save_sheet == self.defalut_placeholder_message:
            messagebox.showerror(title='Error', message='儲存的工作表名稱不可空白')
        else:
            types = self.box.get()
            if types == self.options[0]:
                Judge_Metabolic_Syndrome(io=filepath, select_years=select_year, save_sheet_name=save_sheet, save_file_path=None)
                self.finishInfo()
            else:
                messagebox.showerror(title="錯誤", message="請選擇檔案類型")

    def getIcon(self, img_path):
        img = Image.open(img_path)
        icon = ImageTk.PhotoImage(img) 
        return icon
    

    def linked_to_github(self):
        url = 'https://github.com/Yi-Feng-Liu/Excel-Filter-Tool'
        webbrowser.open(url)
        

    
        
    
 
import tkinter as tk
root = tk.Tk()
Excel_GUI(root)
root.mainloop() 

from tkinter import *
import pandas as pd
from PIL import Image, ImageTk 
import cv2 as cv
import os, sys
from tkinter import filedialog, messagebox
from tkinter.font import Font
from tkinter import ttk
import numpy as np
from util.process import Judge_Metabolic_Syndrome, Metabolic_Syndrome_From_Summary
import webbrowser



def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
 
    return os.path.join(base_path, relative_path)

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
        self.sheet_comfirm_button = Button(root, text="")
        self.years_combobox_label = Label(root, text="", font=self.font, bg=self.bgcolor)
        self.years_entry_label = Label(root, text="", font=self.font, bg=self.bgcolor)
        self.save_sheet_name_label = Label(root, text="", font=self.font, bg=self.bgcolor)
        self.select_button = Button(root, text="")
        self.label = Label(root, text="", font=self.font, bg=self.bgcolor)
        self.savelabel= Label(root, text="", font=self.font, bg=self.bgcolor)
        self.save_sheet_name_entry = EntryWithPlaceholder(master=root, placeholder="", width=30)

        # can't not sesize wiondow
        self.resizable = root.resizable(width=0, height=0)
        self.attributes = root.attributes('-alpha', 1)

        # set the app icon
        self.iconimg = self.getIcon('Images/icon.jpg')
        root.iconphoto(True, self.iconimg)
        root.configure(bg=self.bgcolor)
        # set app title
        self.title = root.title("Excel Filter Tool")
        
        # set radio box to select choose file type
        self.select_input_file_type()
        
        # set decoraction position
        self.origin_img_path = self.resize('Images/link_img.jpg', resize_w=100, resize_h=100, changeBG=True)
        self.decoration = self.getIcon(self.origin_img_path)
        width = self.decoration.width()
        height = self.decoration.height()
        self.pikachu = Button(root, image=self.decoration, command=self.linked_to_github)
        self.pikachu.place(x=self.windows_w - width, y=self.windows_h - height)

        self.Copyright_label = Label(self.root, text='Copyright © 2023 YF Liu. All rights reserved.', font=("微軟正黑體", 7) , bg=self.bgcolor)
        self.Copyright_label.place(x=0, y=self.windows_h-20)


    def create_confirm_botton(self, command):
        # set the confirm button using image
        self.button_img = self.resize('Images/confirm.jpg', resize_w=70, resize_h=60, changeBG=False)
        self.button_img = Image.open(resource_path(self.button_img))
        self.comfirm_button_img = ImageTk.PhotoImage(self.button_img) 
        self.comfirm_button = Button(self.root, command=command, image=self.comfirm_button_img)
        self.comfirm_button.place(x=self.windows_w-250, y=self.windows_h-50, anchor='center')
        self.comfirm_button["state"] = "disabled"

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
        self.save_sheet_name_label.destroy()
        self.select_button.destroy()
        self.label.destroy()
        self.savelabel.destroy()
        self.save_sheet_name_entry.destroy()
        self.years_entry_label.destroy()
        

    def choose_e_file(self):
        """處理電子檔
        """
        self.initialize_button_and_label()
        self.sheet_comfirm_button.destroy()
        
        self.greet = Label(self.root, text="選擇判斷類型", height="3", bg=self.bgcolor, font=self.font)
        self.greet.place(x=250, y=85, anchor='center')
        
        # create combobox
        self.options = ['代謝症候群']
        self.box = ttk.Combobox(self.root, values=self.options)
        # use default
        self.box.current()
        self.box.place(x=250, y=110, anchor='center')

        # set choose file button
        self.select_button = Button(self.root, text="選擇檔案", command=self.open_E_File)
        self.select_button.place(x=225, y=140)

        # update choose file path
        self.label = Label(self.root, text='', font=self.font, bg=self.bgcolor)
        self.label.place(x=250, y=210, anchor='center')
        
        self.combobox_xlabel_place = 30
        self.combobox_place = 150
        self.sheets_names_combobox_label = Label(self.root, text='Process Sheet:', font=self.font, bg=self.bgcolor)
        self.sheets_names_combobox_label.place(x=self.combobox_xlabel_place, y=self.windows_h-250)
        self.sheets_names_combobox = ttk.Combobox(self.root, values=None, width=27)
        self.sheets_names_combobox.place(x=self.combobox_place, y=self.windows_h-247)

        # set years combobox
        self.years_entry_label = Label(self.root, text='年度代碼:', font=self.font, bg=self.bgcolor)
        self.years_entry_label.place(x=self.combobox_xlabel_place, y=self.windows_h-220)
        self.defalut_years_placeholder_message = '輸入年度代碼名稱 (Ex: 111年度健檢)'
        self.years_text_entry = EntryWithPlaceholder(master=self.root, placeholder=self.defalut_years_placeholder_message, width=30)
        self.years_text_entry.place(x=self.combobox_place, y=self.windows_h-217)

        self.save_sheet_name_label = Label(self.root, text='工作表名稱:', font=self.font, bg=self.bgcolor)
        self.save_sheet_name_label.place(x=self.combobox_xlabel_place, y=self.windows_h-190)
        self.defalut_placeholder_message = '輸入要儲存的Sheet名稱'
        self.save_sheet_name_entry = EntryWithPlaceholder(master=self.root, placeholder=self.defalut_placeholder_message, width=30)
        self.save_sheet_name_entry.place(x=self.combobox_place, y=self.windows_h-187)

        self.savelabel= Label(self.root, text='', font=self.font, bg=self.bgcolor)
        self.savelabel.place(x=250, y=370, anchor='center')
        self.create_confirm_botton(command=self.process_Excel_from_E_file)
        # self.comfirm_button = Button(root, command=self.process_Excel_from_E_file, image=self.comfirm_button_img)


    def choose_paper_file(self):
        """處理手key資料
        """
        self.initialize_button_and_label()
        
        self.greet = Label(self.root, text="選擇判斷類型", height="3", bg=self.bgcolor, font=self.font)
        self.greet.place(x=250, y=85, anchor='center')
        
        # create combobox
        self.options = ['代謝症候群']
        self.box = ttk.Combobox(self.root, values=self.options)

        # use default
        self.box.current()
        self.box.place(x=250, y=110, anchor='center')

        # set choose file button
        self.select_button = Button(self.root, text="選擇檔案", command=self.openFile)
        self.select_button.place(x=225, y=140)

        # update choose file path
        self.label = Label(self.root, text='', font=self.font, bg=self.bgcolor)
        self.label.place(x=250, y=210, anchor='center')
        
        self.combobox_xlabel_place = 47
        self.combobox_place = 150
        self.sheets_names_combobox_label = Label(self.root, text='Select Sheet:', font=self.font, bg=self.bgcolor)
        self.sheets_names_combobox_label.place(x=self.combobox_xlabel_place, y=self.windows_h-250)
        self.sheets_names_combobox = ttk.Combobox(self.root, values=None, width=27)
        self.sheets_names_combobox.place(x=self.combobox_place, y=self.windows_h-247)

        # set years combobox
        self.years_combobox_label = Label(self.root, text='Select Years:', font=self.font, bg=self.bgcolor)
        self.years_combobox_label.place(x=self.combobox_xlabel_place, y=self.windows_h-220)
        self.years_combobox = ttk.Combobox(self.root, values=None, width=27)
        self.years_combobox.place(x=self.combobox_place, y=self.windows_h-217)
        # set entry & save sheet
        self.save_sheet_name_label = Label(self.root, text='Sheet Name:', font=self.font, bg=self.bgcolor)
        self.save_sheet_name_label.place(x=self.combobox_xlabel_place, y=self.windows_h-190)
        self.defalut_placeholder_message = '輸入要儲存的Sheet名稱'
        self.save_sheet_name_entry = EntryWithPlaceholder(master=self.root, placeholder=self.defalut_placeholder_message, width=30)
        self.save_sheet_name_entry.place(x=self.combobox_place, y=self.windows_h-187)

        # set sheet comfirm button
        self.sheet_comfirm_button = Button(self.root, text="Submit", command=self.get_sheet_name)
        self.sheet_comfirm_button.pack()
        self.sheet_comfirm_button.place(x=390, y=self.windows_h-238, anchor='center')
        self.sheet_comfirm_button["state"] = "disabled"
        self.create_confirm_botton(command=self.process_Excel_from_paper)
        # self.comfirm_button = Button(root, command=self.process_Excel_from_paper, image=self.comfirm_button_img)
    
    def resize(self, imgpath:str, resize_w:int, resize_h:int, changeBG=True):
        image_name = imgpath.split("\\")[-1].split(".")[0]
        
        origin_img = cv.imread(imgpath)
        # replace the iamge white background to #d2efd3
        if changeBG:
            origin_img[np.where((origin_img == [255, 255, 255]).all(axis=2))] = [211, 239, 211]
        resize_img = cv.resize(origin_img, (resize_w, resize_h))
        save_path = f'Images/{image_name}_resize.jpg'
        cv.imwrite(save_path, resize_img)
        return save_path


    def finishInfo(self):
        messagebox.showinfo(title="通知", message="OK")


    def openFile(self):
        global filepath

        filepath = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
        if filepath:
            self.label.config(text=f'選取的檔案路徑為\n {filepath}')
            self.comfirm_button["state"] = "active"
            # get list of sheet names 
            self.tabs = pd.ExcelFile(filepath).sheet_names 
            self.sheets_names_combobox['values'] = self.tabs
            self.sheet_comfirm_button["state"] = "active"
        else:
            self.sheets_names_combobox['values'] = self.empty_ls
            self.label.config(text=f'沒有選取的檔案!')
            self.comfirm_button["state"] = "disabled"
            self.sheet_comfirm_button["state"] = "disabled"
    
    def open_E_File(self):
        global filepath

        filepath = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
        if filepath:
            self.label.config(text=f'選取的檔案路徑為\n {filepath}')
            self.comfirm_button["state"] = "active"
            # get list of sheet names 
            self.tabs = pd.ExcelFile(filepath).sheet_names 
            self.sheets_names_combobox['values'] = self.tabs
        else:
            self.sheets_names_combobox['values'] = self.empty_ls
            self.label.config(text=f'沒有選取的檔案!')
            self.comfirm_button["state"] = "disabled"
        
    def get_sheet_name(self):
        global sheetName
        sheetName = self.sheets_names_combobox.get()
        self.comfirm_button["state"] = "active"
        try:
            if len(sheetName) == 0:
                messagebox.showerror(title='Error', message='請選擇要處理的工作表')
            elif sheetName not in self.tabs:
                messagebox.showerror(title='Error', message='工作表不存在')
            else:
                df = pd.read_excel(filepath, sheet_name=sheetName, engine = 'openpyxl')
                df_years_list = list(dict.fromkeys(df['年度代碼'].tolist()))
                self.years_combobox['values'] = df_years_list  
        except Exception as Error:
            self.years_combobox['values'] = self.empty_ls
            messagebox.showerror(title='Error', message=Error)
            self.comfirm_button["state"] = "disabled"

    def process_Excel_from_paper(self):
        """處理手key的新人體檢資料
        """
        select_year = self.years_combobox.get()
        save_sheet = self.save_sheet_name_entry.get()


        if save_sheet in self.tabs:
            messagebox.showerror(title='Error', message='工作表名稱已存在')
        elif select_year not in self.years_combobox['values'] and save_sheet == self.defalut_placeholder_message:
            messagebox.showerror(title='Error', message='年份以及儲存的工作表不可空白')
        elif save_sheet == self.defalut_placeholder_message:
            messagebox.showerror(title='Error', message='儲存的工作表名稱不可空白')
        elif select_year == 'NaN':
            messagebox.showerror(title='Error', message='不存在的年份')
        elif select_year not in self.years_combobox['values']:
            messagebox.showerror(title='Error', message='請選擇年份')
        else:
            try:
                types = self.box.get()
                if types == self.options[0]:
                    Judge_Metabolic_Syndrome(io=filepath, tab=sheetName, select_years=select_year, save_sheet_name=save_sheet).main_procesdure()

                    self.finishInfo()
                else:
                    messagebox.showerror(title="錯誤", message="請選擇檔案類型")
            except:
                messagebox.showerror(title="錯誤", message='請選擇年分')


    def process_Excel_from_E_file(self):
        """處理電子檔的健檢資料
        """
        sheetName = self.sheets_names_combobox.get()
        years_text = self.years_text_entry.get()
        save_sheet = self.save_sheet_name_entry.get()
        if save_sheet in self.tabs:
            messagebox.showerror(title='Error', message='工作表名稱已存在')
        elif len(sheetName)==0 :
            messagebox.showerror(title='Error', message='請選擇工作表')
        elif sheetName not in self.tabs:
            messagebox.showerror(title='Error', message='工作表不存在')
        elif save_sheet == self.defalut_placeholder_message and years_text == self.defalut_years_placeholder_message:
            messagebox.showerror(title='Error', message='年度代碼 & 儲存的工作表名稱不可空白')
        elif years_text == self.defalut_years_placeholder_message:
            messagebox.showerror(title='Error', message='年度代碼不可空白')
        elif save_sheet == self.defalut_placeholder_message:
            messagebox.showerror(title='Error', message='儲存的工作表名稱不可空白')
        else:
            try:
                types = self.box.get()
                if types == self.options[0]:
                    Metabolic_Syndrome_From_Summary(io=filepath, tab=sheetName, save_sheet_name=save_sheet, years_text=years_text).main_procesdure()
                    self.finishInfo()
                else:
                    messagebox.showerror(title="錯誤", message="請選擇檔案類型")
            except Exception as e:
                messagebox.showerror(title="錯誤", message=e)

    def getIcon(self, img_path):
        img = Image.open(resource_path(img_path))
        icon = ImageTk.PhotoImage(img) 
        return icon
    

    def linked_to_github(self):
        url = 'https://github.com/Yi-Feng-Liu/Excel-Filter-Tool'
        webbrowser.open(url)
        

    
        
    
 
# import tkinter as tk
# root = tk.Tk()
# Excel_GUI(root)
# root.mainloop() 

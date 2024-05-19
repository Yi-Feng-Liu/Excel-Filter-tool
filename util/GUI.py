from tkinter import *
import pandas as pd
from PIL import Image, ImageTk 
import cv2 as cv
import os, sys
from tkinter import filedialog, messagebox
from tkinter import font
from tkinter import ttk, END
import numpy as np
from util.process import *
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
        self.bgcolor = "#d2efd3"
        self.empty_ls = []

        # set windows size
        self.windows_w = 500
        self.windows_h = 500
        self.font = font.Font(family="微軟正黑體", size=12)

        # set windows position
        self.geometry = root.geometry(f"{str(self.windows_w)}x{str(self.windows_h)}+800+300")
        
        # inintialize label
        self.filepath = []
        self.greet = Label(root, text="", height="3", bg=self.bgcolor, font=self.font)
        self.sheets_names_combobox_label = Label(root, text="", font=self.font, bg=self.bgcolor)
        self.sheet_comfirm_button = Button(root, text="")
        self.years_combobox_label = Label(root, text="", font=self.font, bg=self.bgcolor)
        self.years_entry_label = Label(root, text="", font=self.font, bg=self.bgcolor)
        self.save_sheet_name_label = Label(root, text="", font=self.font, bg=self.bgcolor)
        self.select_button = Button(root, text="")
        self.label = Label(root, text="", font=self.font, bg=self.bgcolor)
        self.savelabel= Label(root, text="", font=self.font, bg=self.bgcolor)
        self.show_save_dir_entry = Entry(master=root)
        self.save_sheet_name_entry = EntryWithPlaceholder(master=root, placeholder="", width=30)
        self.box = ttk.Combobox(self.root, values=[])
        self.sheets_names_combobox = ttk.Combobox(self.root, values=None)
        self.saveFile_button = Button(self.root, text="", command=None)
        self.years_combobox = ttk.Combobox(self.root, values=None)
        self.years_text_entry = EntryWithPlaceholder(
            master = self.root, 
            placeholder = '', 
            width = 30
        )
        

        # can't not sesize wiondow
        self.resizable = root.resizable(width=0, height=0)
        self.attributes = root.attributes('-alpha', 1)

        # set the app icon
        self.iconimg = self.getIcon('Images\\icon.jpg')
        root.iconphoto(True, self.iconimg)
        root.configure(bg=self.bgcolor)

        # set app title
        self.title = root.title("Excel Filter Tool")
        
        # set radio box to select choose file type
        self.select_mode()
        
        # set decoraction position
        self.origin_img_path = self.resize(resource_path('Images\\link_img.jpg'), resize_w=100, resize_h=100, changeBG=True)
        self.decoration = self.getIcon(self.origin_img_path)
        width = self.decoration.width()
        height = self.decoration.height()
        self.pikachu = Button(root, image=self.decoration, command=self.linked_to_github)
        self.pikachu.place(x=self.windows_w - width, y=self.windows_h - height)

        self.Copyright_label = Label(self.root, text='Copyright © 2023 YF Liu. All rights reserved.', font=("微軟正黑體", 7) , bg=self.bgcolor)
        self.Copyright_label.place(x=0, y=self.windows_h-20)


    def create_confirm_botton(self, command):
        # set the confirm button using image
        self.button_img = self.resize('Images\\confirm.jpg', resize_w=70, resize_h=60, changeBG=False)
        self.button_img = Image.open(resource_path(self.button_img))
        self.comfirm_button_img = ImageTk.PhotoImage(self.button_img) 
        self.comfirm_button = Button(self.root, command=command, image=self.comfirm_button_img)
        self.comfirm_button.place(x=self.windows_w-250, y=self.windows_h-50, anchor='center')
        self.comfirm_button["state"] = "disabled"


    def select_mode(self):
        # global 
        methods = [
            ('健檢問卷', 1, self.show_e_file_ui),
            ('健檢資料', 2, self.show_manual_file_ui),
            ('轉換模式', 3, self.show_excel_transfer_ui)

        ]
        self.v = IntVar()
        self.v.set(0)
    
        self.e_file = Radiobutton(
            self.root, 
            text=methods[0][0], 
            variable=self.v, 
            value=methods[0][1], 
            command=methods[0][2],
            bg=self.bgcolor, 
            font=self.font
        )
        self.e_file.place(x=80, y=30, anchor='center')

        self.manual_file = Radiobutton(
            self.root, 
            text=methods[1][0], 
            variable=self.v, 
            value=methods[1][1], 
            command=methods[1][2], 
            bg=self.bgcolor, 
            font=self.font
        )
        self.manual_file.place(x=260, y=30, anchor='center')

        self.excelTransferToWord = Radiobutton(
            self.root, 
            text=methods[2][0], 
            variable=self.v, 
            value=methods[2][1], 
            command=methods[2][2], 
            bg=self.bgcolor, 
            font=self.font
        )
        self.excelTransferToWord.place(x=420, y=30, anchor='center')


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
        self.box.destroy()
        self.sheets_names_combobox.destroy()
        self.years_combobox.destroy()
        self.years_text_entry.destroy()
        self.saveFile_button.destroy()
        self.show_save_dir_entry.destroy()
        self.sheet_comfirm_button.destroy()
        
        

    def show_e_file_ui(self):
        """
        處理電子檔
        """
        self.initialize_button_and_label()
        self.sheet_comfirm_button.destroy()
        
        self.greet = Label(self.root, text="選擇判斷類型", height="3", bg=self.bgcolor, font=self.font)
        self.greet.place(x=250, y=85, anchor='center')
        
        # create combobox
        self.options = ['工作過勞量表']
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
        self.years_text_entry = EntryWithPlaceholder(
            master = self.root, 
            placeholder = self.defalut_years_placeholder_message, 
            width = 30
        )
        self.years_text_entry.place(x=self.combobox_place, y=self.windows_h-217)

        self.save_sheet_name_label = Label(self.root, text='工作表名稱:', font=self.font, bg=self.bgcolor)
        self.save_sheet_name_label.place(x=self.combobox_xlabel_place, y=self.windows_h-190)
        self.defalut_placeholder_message = '輸入要儲存的Sheet名稱'
        self.save_sheet_name_entry = EntryWithPlaceholder(
            master = self.root, 
            placeholder = self.defalut_placeholder_message, 
            width = 30
        )
        self.save_sheet_name_entry.place(x=self.combobox_place, y=self.windows_h-187)

        self.savelabel= Label(self.root, text='', font=self.font, bg=self.bgcolor)
        self.savelabel.place(x=250, y=370, anchor='center')
        self.create_confirm_botton(command=self.process_Excel_from_E_file)
        # self.comfirm_button = Button(root, command=self.process_Excel_from_E_file, image=self.comfirm_button_img)


    def show_manual_file_ui(self):
        """
        處理手key資料
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
        self.save_sheet_name_entry = EntryWithPlaceholder(
            master = self.root, 
            placeholder = self.defalut_placeholder_message, 
            width = 30
        )
        self.save_sheet_name_entry.place(x=self.combobox_place, y=self.windows_h-187)

        # set sheet comfirm button
        self.sheet_comfirm_button = Button(self.root, text="Submit", command=self.get_sheet_name)
        self.sheet_comfirm_button.pack()
        self.sheet_comfirm_button.place(x=390, y=self.windows_h-238, anchor='center')
        self.sheet_comfirm_button["state"] = "disabled"
        self.create_confirm_botton(command=self.process_Excel_from_manual)


    def show_excel_transfer_ui(self):
        self.initialize_button_and_label()
        self.select_button = Button(self.root, text="選擇檔案", command=self.open_muti_excelFile)
        self.select_button.place(x=225, y=80)
        self.label = Label(self.root, text='', font=self.font, bg=self.bgcolor)
        self.label.place(x=250, y=140, anchor='center')

        self.saveFile_button = Button(self.root, text="...", command=self.decide_zip_save_path)
        self.saveFile_button.place(x=390, y=249)
        self.show_save_dir_entry = Entry(master=self.root, width=40)
        self.show_save_dir_entry.place(x=100, y=250)
        try:
            self.create_confirm_botton(command=self.run_excel_to_word)
        except Exception as e:
            messagebox.showerror(title="錯誤", message=e)
        

    def resize(self, imgpath:str, resize_w:int, resize_h:int, changeBG=True):
        image_name = imgpath.split("\\")[-1].split(".")[0]
        
        origin_img = cv.imread(resource_path(imgpath))
        # replace the iamge white background to #d2efd3
        if changeBG:
            if origin_img.ndim > 0:
                mask = np.logical_and.reduce(origin_img == 255, axis=2)
                origin_img[mask] = [211, 239, 211]
        resize_img = cv.resize(origin_img, (resize_w, resize_h))
        save_path = resource_path(f'Images\\{image_name}_resize.jpg')
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
            self.tabs = pd.ExcelFile(resource_path(filepath)).sheet_names 
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
            self.tabs = pd.ExcelFile(resource_path(filepath)).sheet_names 
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


    def process_Excel_from_manual(self):
        """
        處理新人體檢資料
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
                    run_jms = Judge_Metabolic_Syndrome(
                        io = filepath, 
                        src_worksheet = sheetName, 
                        select_years = select_year, 
                        save_sheet_name = save_sheet
                    )
                    run_jms()
                    self.finishInfo()

                else:
                    messagebox.showerror(title="錯誤", message="請選擇檔案類型")

            except:
                messagebox.showerror(title="錯誤", message='請選擇年分')


    def process_Excel_from_E_file(self):
        """
        處理健檢問卷資料
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
                    run_mode1 = Metabolic_Syndrome_From_Summary(
                        io = filepath, 
                        src_worksheet = sheetName, 
                        save_sheet_name = save_sheet, 
                        years_text = years_text
                    )
                    run_mode1()
                    self.finishInfo()

                # elif types == self.options[1]:
                #     run_mode2 = Judge_Work_Pressure(
                #         io = filepath, 
                #         src_worksheet = sheetName, 
                #         save_sheet_name = save_sheet, 
                #         years_text = years_text
                #     )
                #     run_mode2()
                #     self.finishInfo()

                else:
                    messagebox.showerror(title="錯誤", message="請選擇檔案類型")

            except Exception as e:
                messagebox.showerror(title="錯誤", message=e)


    def getIcon(self, img_path):
        img = Image.open(resource_path(img_path))
        icon = ImageTk.PhotoImage(img) 
        return icon
    

    def linked_to_github(self):
        url = 'https://github.com/Yi-Feng-Liu/Excel-Filter-Tool/releases'
        webbrowser.open(url)


    def package_into_zip(self, save_zip_path):
        import zipfile
        from glob import glob
        folder_path = 'documents\\'
        files = glob(resource_path(folder_path) + '*.docx')
        
        with zipfile.ZipFile(save_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file in files:
                arc_name = os.path.basename(file)
                messagebox.showinfo(title="完成通知", message=arc_name)
                zipf.write(file, arcname=arc_name)
                os.remove(file)


    def open_muti_excelFile(self):
        self.filepath = filedialog.askopenfilenames(filetypes=[("Excel files", ".xlsx .xls")])
        if len(self.filepath) > 0:
            if len(self.filepath) <= 2:
                self.label.config(text=f'選擇\n{(self.filepath)}')
            else:
                self.label.config(text=f'已選擇{len(self.filepath)}個檔案')
            self.comfirm_button["state"] = "active"
        else:
            self.label.config(text=f'沒有選取的檔案!')
            self.comfirm_button["state"] = "disabled"


    def check_savefilepath_entry(self):
        if self.show_save_dir_entry.get() != '':
            self.comfirm_button["state"] = "active"
        else:
            self.comfirm_button["state"] = "disabled"


    def decide_zip_save_path(self):
        from datetime import datetime
        savefilepath = filedialog.asksaveasfilename(filetypes=[("壓縮檔", ".zip")])
        # clear text every time when who reselect save path
        formatted_datetime_str = datetime.now().replace(second=0, microsecond=0).strftime('%Y%m%d%H%M')
        self.show_save_dir_entry.delete(0, END)
        self.show_save_dir_entry.insert(0, resource_path(savefilepath + f"_{formatted_datetime_str}.zip"))
        self.check_savefilepath_entry()


    def run_excel_to_word(self):
        for file in self.filepath:
            try:
                etwt = excel_to_word_table(excel_source=file)
                etwt.run_convert()
            except Exception as e:
                messagebox.showerror(title="錯誤1", message=e)
        try:
            self.package_into_zip(save_zip_path=(self.show_save_dir_entry.get()))
            self.finishInfo()
        except Exception as e:
            messagebox.showerror(title="錯誤2", message=e)

    # def click_download_lastest_version(self):
    #     from github import Github

    #     # 設置你的 GitHub 身份驗證信息 #90天後要更新
    #     g = Github(self.personal_token)

    #     # 指定你的存儲庫名稱和發布版本名稱
    #     repo_name = self.repo_name
    #     release_name = self.release_name

    #     # 獲取存儲庫對象
    #     repo = g.get_repo(repo_name)

    #     # 獲取最新的發布版本對象
    #     release = repo.get_release(release_name)

    #     # 獲取執行檔的下載鏈接
    #     download_url = release.get_assets()[0].browser_download_url

    #     # 下載執行檔
    #     import urllib.request
    #     urllib.request.urlretrieve(download_url, self.download_rar_name)
       
from tkinter import *
import time
from PIL import Image, ImageTk 
import cv2 as cv
from tkinter import filedialog, messagebox
from tkinter.font import Font
import threading
from tkinter import ttk
import numpy as np
from process import Judge_Metabolic_Syndrome
import webbrowser


class Excel_GUI:
    def __init__(self, root):
        self.root = root
        # set background color
        self.bgcolor = "#d2efd3"

        # set windows size
        self.windows_w = 500
        self.windows_h = 400
        self.font = Font(family="微軟正黑體", size=12)

        # set windows position
        self.geometry = root.geometry(f"{str(self.windows_w)}x{str(self.windows_h)}+800+300")

        # can't not sesize wiondow
        self.resizable = root.resizable(width=0, height=0)
        self.attributes = root.attributes('-alpha', 1)
        # set greet label at the top
        self.greet = Label(root, text="選擇判斷類型", height="3", bg=self.bgcolor, font=self.font)
        self.greet.pack()

        # set the app icon
        self.iconimg = self.getIcon('icon.jpg')
        root.iconphoto(True, self.iconimg)
        root.configure(bg=self.bgcolor)
        # set app title
        self.title = root.title("Excel Filter Tool")

        # create combobox
        self.options = ['代謝症候群','選項2','選項3','選項4','選項5', '選項6', '選項7', '選項8', '選項9', '選項10', '選項11']
        self.box = ttk.Combobox(root, values=self.options)

        # use default
        self.box.current()
        self.box.pack(pady=10)

        # set choose file button
        self.select_button = Button(root, text="選擇檔案", command=self.openFile)
        self.select_button.pack(pady=10) 

        # update choose file path
        self.label = Label(root, text='', font=self.font, bg=self.bgcolor)
        self.label.pack(pady=10)

        # set decoraction position
        self.origin_img_path = self.resize('link_img.jpg', resize_w=100, resize_h=100, changeBG=True)
        self.decoration = self.getIcon(self.origin_img_path)
        width = self.decoration.width()
        height = self.decoration.height()
        self.pikachu = Button(root, image=self.decoration, command=self.linked_to_github)
        self.pikachu.place(x=self.windows_w - width, y=self.windows_h - height)

        # set entry for enter process sheet & save sheet
        self.process_sheet_name_label = Label(root, text='Sheet 1:', font=self.font, bg=self.bgcolor)
        self.process_sheet_name_label.place(x=80, y=230)
        self.process_sheet_name_entry = Entry(root, width=30)
        self.process_sheet_name_entry.place(x=150, y=233)
        self.save_sheet_name_label = Label(root, text='Sheet 2:', font=self.font, bg=self.bgcolor)
        self.save_sheet_name_label.place(x=80, y=250)
        self.save_sheet_name_entry = Entry(root, width=30)
        self.save_sheet_name_entry.place(x=150, y=253)

        # set the confirm button using image
        self.button_img = self.resize('confirm.jpg', resize_w=70, resize_h=60, changeBG=False)
        self.button_img = Image.open(self.button_img)
        self.comfirm_button_img = ImageTk.PhotoImage(self.button_img) 
        self.comfirm_button = Button(root, text="確認", command=self.processExcel, image=self.comfirm_button_img)
        self.comfirm_button.pack()
        self.comfirm_button.place(x=self.windows_w-250, y=self.windows_h-50, anchor='center')
        self.comfirm_button["state"] = "disabled"

        self.Copyright_label = Label(self.root, text='Copyright © 2023 YF Liu. All rights reserved.', font=("微軟正黑體", 7) , bg=self.bgcolor)
        self.Copyright_label.place(x=0, y=380)

        
        
    def resize(self, imgpath:str, resize_w:int, resize_h:int, changeBG=True):
        image_name = imgpath.split("\\")[-1].split(".")[0]
        
        origin_img = cv.imread(imgpath)
        # replace the iamge white background to #d2efd3
        if changeBG:
            origin_img[np.where((origin_img == [255, 255, 255]).all(axis=2))] = [211, 239, 211]
        resize_img = cv.resize(origin_img, (resize_w, resize_h))
        save_path = f'D:\\GUI\\Test\\{image_name}_resize.jpg'
        cv.imwrite(save_path, resize_img)
        return save_path


    def finishInfo(self):
        messagebox.showinfo(title="通知", message="OK")


    def openFile(self):
        global filepath
        filepath = filedialog.askopenfilename(title="選擇檔案", filetypes=[("Excel files", ".xlsx .xls")])

        if filepath:
            self.label.config(text=f'選取的檔案路徑為\n {filepath}')
            self.comfirm_button["state"] = "active"
        else:
            self.label.config(text=f'沒有選取的檔案!')
            self.comfirm_button["state"] = "disabled"

    def processExcel(self):
        process_sheet = self.process_sheet_name_entry.get()
        save_sheet = self.save_sheet_name_entry.get()
        if process_sheet == save_sheet:
            messagebox.showerror(title='Error', message='The Same of Sheet Names That is Not Allow')
        else:
            types = self.box.get()
            if types == self.options[0]:
                JMS = Judge_Metabolic_Syndrome()
                t = threading.Thread(target=JMS.process_Metabolic_Syndrome(filepath, src_worksheet=process_sheet, dst_worksheet=save_sheet))
                t.start()
                self.update_clock()
                time.sleep(2)
                self.clock_label.destroy() 
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
        

    def update_clock(self):
        current_time = time.strftime("%S")
        clock_label = tk.Label(root, font=("Helvetica", 24))
        clock_label.config(text=current_time)
        clock_label.place(x=250, y=295, anchor="center")
        self.root.after(1000, self.update_clock)
    
 
import tkinter as tk
root = tk.Tk()
Excel_GUI(root)
root.mainloop() 
from tkinter import *
from PIL import Image, ImageTk 
import cv2 as cv
from tkinter import filedialog, messagebox
from tkinter.font import Font
import openpyxl
import pandas as pd
from tkinter import ttk
import numpy as np


global filepath
filepath = ""


def resize(imgpath:str, resize_w:int, resize_h:int, changeBG=True):
    image_name = imgpath.split("\\")[-1].split(".")[0]
    
    origin_img = cv.imread(imgpath)
    # replace the iamge white background to #d2efd3
    if changeBG:
        origin_img[np.where((origin_img == [255, 255, 255]).all(axis=2))] = [211, 239, 211]
    resize_img = cv.resize(origin_img, (resize_w, resize_h))
    save_path = f'D:\\GUI\\Test\\{image_name}_resize.jpg'
    cv.imwrite(save_path, resize_img)
    return save_path


def finishInfo():
    messagebox.showinfo(title="通知", message="已篩選完成")

def openFile():
    filepath = filedialog.askopenfilename(title="選擇檔案", filetypes=[("Excel files", ".xlsx .xls")])
    
    if (filepath):
        label.config(text=f'選取的檔案路徑為\n\n {filepath}')
        # messagebox.showinfo(title="通知", message="選擇成功")
        comfirm_button["state"] = "active"
    else:
        label.config(text=f'沒有選取的檔案!')
        comfirm_button["state"] = "disabled"
    

def processExcel():
    types = box.get()
    if (types):
        print(types)
        finishInfo()
    else:
        messagebox.showerror(title="錯誤", message="請選擇檔案類型")

def getIcon(img_path):
    img = Image.open(img_path)
    icon = ImageTk.PhotoImage(img) 
    return icon
    

bgcolor = "#d2efd3"

# set windows size
windows_w = 500
windows_h = 400

root = Tk()
font = Font(family="微軟正黑體", size=12)
root.geometry(f"{str(windows_w)}x{str(windows_h)}+800+300")
root.resizable(width=0, height=0)
root.attributes('-alpha', 1)


root.title("Excel Filter Tool")

# set icon.
iconimg = getIcon('D:\\GUI\\Test\\icon.jpg')
root.iconphoto(True, iconimg)
root.configure(bg=bgcolor)

greet = Label(root, text="選擇檔案類型", height="3", bg=bgcolor, font=font)
greet.pack()

# read the img and set the img at right down
origin_img_path = resize('D:\\GUI\\Test\\456.jpg', resize_w=100, resize_h=100, changeBG=True)
decoration = getIcon(origin_img_path)
width = decoration.width()
height = decoration.height()

pikachu = Label(root, image=decoration)
pikachu.place(x=windows_w-width, y=windows_h-height)

# options list
options = ['代謝症候群','選項2','選項3','選項4','選項5', '選項6', '選項7', '選項8', '選項9', '選項10', '選項11']
box = ttk.Combobox(root, values=options)

# use default
box.current()
box.pack(pady=10)

select_button = Button(root, text="選擇檔案", command=openFile)
select_button.pack(pady=10) 

label = Label(root, text='', font=font, bg=bgcolor)
label.pack(pady=10)


button_img = resize('D:\\GUI\\Test\\confirm.jpg', resize_w=70, resize_h=60, changeBG=False)
button_img = Image.open(button_img)
comfirm_button_img = ImageTk.PhotoImage(button_img) 
comfirm_button = Button(root, text="確認", command=processExcel, image=comfirm_button_img)
comfirm_button.pack(pady=10)
comfirm_button["state"] = "disabled"

root.mainloop()

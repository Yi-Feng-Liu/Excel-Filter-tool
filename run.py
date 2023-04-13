from tkinter import *
from PIL import Image, ImageTk 
import cv2 as cv
from tkinter import filedialog, messagebox
from tkinter.font import Font
import openpyxl
import pandas as pd
from tkinter import ttk

global filepath
filepath = ""

def finishInfo():
    messagebox.showinfo(title="通知", message="已篩選完成")

def openFile():
    filepath = filedialog.askopenfilename(title="選擇檔案", filetypes=[("Excel files", ".xlsx .xls")])
    
    if (filepath):
        label.config(text=f'選取的檔案路徑為\n {filepath}')
        # messagebox.showinfo(title="通知", message="選擇成功")
        comfirm_button["state"] = "active"
    else:
        label.config(text=f'無選取檔案')
        comfirm_button["state"] = "disabled"
    

def processExcel():
    types = box.get()
    if (types):
        print(types)
        finishInfo()
    else:
        messagebox.showerror(title="錯誤", message="請選擇檔案類型")

bgcolor = "#ffffff"

root = Tk()
font = Font(family="微軟正黑體", size=12)
root.geometry("500x400+800+300")
root.resizable(width=0, height=0)
root.attributes('-alpha', 1)


root.title("Excel Filter Tool")

# set icon
img = Image.open('123.jpg')
icon = ImageTk.PhotoImage(img) 
root.iconphoto(True, icon)
root.configure(bg=bgcolor)
 
greet = Label(root, text="選擇檔案類型", height="3", bg=bgcolor, font=font)
greet.pack()

# set the img at right down
img2 = Image.open('456.jpg')
icon2 = ImageTk.PhotoImage(img2) 
pikachu = Label(root, image=icon2)
pikachu.pack(side='bottom', padx=1, pady=2)

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

comfirm_button = Button(root, text="確認", command=processExcel, bg='#f09595', )
comfirm_button.pack(pady=10)
comfirm_button["state"] = "disabled"

root.mainloop()

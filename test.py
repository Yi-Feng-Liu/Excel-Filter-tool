import tkinter as tk
from tkinter import font

root = tk.Tk()

my_font = font.Font(family='Helvetica', size=24, weight='bold')


my_label = tk.Label(root, text='Hello, world!', font=my_font)
my_label.pack()

root.mainloop()
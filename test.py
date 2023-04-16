import tkinter as tk
import time

def update_clock():
    current_time = time.strftime("%S")
    clock_label.config(text=current_time)
    root.after(1000, update_clock)

root = tk.Tk()
root.title("Clock")
root.geometry("200x50")

clock_label = tk.Label(root, font=("Helvetica", 24))
clock_label.pack(expand=True)

update_clock()

root.mainloop()

import tkinter as tk
from util.GUI import Excel_GUI


def main():
    root = tk.Tk()
    Excel_GUI(
        root
        # personal_token='personal_token', 
        # repo_name='Yi-Feng-Liu/Excel-Filter-Tool', 
        # release_name='v2.0.0', 
        # download_rar_name='Widget.rar'
    )
    root.mainloop()



main()
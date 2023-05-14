import tkinter as tk
from util.GUI import Excel_GUI
from threading import Thread

def main():
    root = tk.Tk()
    Thread(target=Excel_GUI, args=(root)).start()
    # Excel_GUI(
    #     root, 
    #     personal_token='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', 
    #     repo_name='Yi-Feng-Liu/Excel-Filter-Tool', 
    #     release_name='v2.0.0', 
    #     download_rar_name='Widget.rar'
    # )
    root.mainloop()



main()
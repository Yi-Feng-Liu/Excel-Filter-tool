# Excel-Filter-Tool version 1.0(This tool made for Lyra)

A tool used to filter Excel file and will create the new worksheet that you need to key in the entry. 
 

### *If name of sheet name dose not exist, it will be auto created and save it. The combobox you can select the type which you want to filter.*

# Method

> 1. Select the type you want to filter.

> 2. Click the choose file button to select file. If you see the submit button, pleat click it.

> 3. Please enter the names of sheet you want to process and save, respectively.

> 4. Press the confirm button. The tool will be start to process file. You only need to wait the 'OK' window appear.

> 5. Check the result is you want, if not please contact me.

# Environment

To run this program, your virtual environment will need the following modules.

1. pandas

2. numpy

3. openpyxl

4. tkinter

5. openCV

6. PIL

> Let all its dependencies into a single package:

    pyinstaller --onefile -w --icon=Images/app_icon.ico run.py --name "Excel Filter Tool" --add-data "util/*py;util" --add-data   "Images/*jpg;Images"






# Excel-Filter-Tool  (This tool made for Lyra)


## version 1.0:

This tool is used to filter and analyze Metabolic Syndrome data in an Excel file, and generate a new worksheet where you can enter your data.

## version 2.0:

Add a 'Process Work Pressure' option to the E file processing options.

Add a new function that can be used to process an Excel file on work pressure and quantify the text in each cell.

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

> Combine all its dependencies into a single package.:

    pyinstaller --onefile -w --icon=Images/app_icon.ico run.py --name "Excel Filter Tool" --add-data "util/*py;util" --add-data   "Images/*jpg;Images"






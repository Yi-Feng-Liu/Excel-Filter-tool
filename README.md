# Excel-Filter-Tool  PyQt6 version (This tool is made for Lyra)

In this version, Tkinter version is no longer being updated.

## version 1.0:

This tool is used to filter and analyze Metabolic Syndrome data in an Excel file, and generate a new worksheet where you can enter your data.

## version 2.0:

1. Added a 'Process Work Pressure' option to the Excel file processing combobox.

2. Added a new function that can process Excel files on work pressure and quantify the text in each cell.

### *If the sheet name does not exist, it will be automatically created and saved. You can select the type you want to filter from the combobox.*

# Method

> 1. Select the type you want to filter.

> 2. Click the 'Choose File' button to select a file. If you see the 'Submit' button, please click it.

> 3. Please enter the names of the sheets you want to process and save, respectively.

> 4. Press the confirm button. The tool will start processing the file. You only need to wait for the 'OK' window to appear.

> 5. Check if the result is what you want. If not, please contact me.

# Environment

To run this program, your virtual environment will need the following modules.

1. pandas

2. numpy

3. openpyxl

4. PyQt6

5. openCV

6. PIL

> Combine all its dependencies into a single package.:

    pyinstaller --onefile -w --icon=Images/app_icon.ico run.py --name "Excel Filter Tool" --add-data "util/*py;util" --add-data   "Images/*jpg;Images" --add-data "util/*json;util"






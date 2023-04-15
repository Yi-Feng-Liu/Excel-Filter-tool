import openpyxl
import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.styles import Alignment, numbers
import copy
import time
import threading


class Judge_Metabolic_Syndrome:
    def __init__(self):
        self.standard_waistline = 90
        self.standard_systolic_blood_pressure = 130
        self.standard_diastolic_blood_pressure = 85
        self.standard_glucose = 100
        self.standard_triglycerides = 150
        self.hdlc = 40
        self.red_font = Font(color='FF0000')
        self.metabolic_syndrome_column = [5, 9, 10, 11, 12, 14, 15]
        self.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        self.font = Font(color='FF0000')
        self.over_standard_cut = 0

    def change_date_time(self, worksheet, number_of_column):
        """Remove hours:minute:second format of the datetime 

        Args:
            worksheet : excel worksheet
            number_of_column : If column name is G, the number of column is 7, etc.

        Returns:
            _worksheet: 
        """
        # 變更時間格式
        for row in worksheet.iter_rows(min_row=2, min_col=number_of_column, max_col=number_of_column):
            for cell in row:
                cell.number_format = numbers.FORMAT_DATE_YYYYMMDD2
        return worksheet


    def place_center(self, worksheet):
        align = Alignment(horizontal='center', vertical='center')
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = align
        return worksheet


    def set_specific_column_format(self, worksheet:str, eng_column:str):
        """set the specific column format.

        Args:
            worksheet (str): Excel worksheet
            eng_column (str): like 'A' or 'G' column
        """
        worksheet.column_dimensions[eng_column].width = 20
        worksheet[f'{eng_column}1'].fill = self.fill
        worksheet[f'{eng_column}1'].font = self.font
        return worksheet


    def set_red_font(self, val):
        return 'color: red'


    def change_font_color_format(self, cell, fill_cell=True):
        """Change cell color and font color 

        Args:
            cell : the cell coordinate
        """
        if fill_cell:
            cell.fill = self.fill
        cell.font = self.font

    def label_over_standard(self, worksheet):
        """Use to label the cell, if cell's value exceed the standard

        Args:
            worksheet: the excel work sheet 

        Returns:
            worksheet
        """
        for row in worksheet.iter_rows(min_row=2):
            people_name = row[1]
            
            gender = row[self.metabolic_syndrome_column[0]] 
            waistline = row[self.metabolic_syndrome_column[1]]
            systolic = row[self.metabolic_syndrome_column[2]]
            diastolic = row[self.metabolic_syndrome_column[3]]
            glucose = row[self.metabolic_syndrome_column[4]]
            triglycerides = row[self.metabolic_syndrome_column[5]]
            hdlc = row[self.metabolic_syndrome_column[6]]
            
            if gender.value == '男':
                if waistline.value >= self.standard_waistline:
                    self.change_font_color_format(waistline)
                    self.over_standard_cut += 1
                if systolic.value >= self.standard_systolic_blood_pressure:
                    self.change_font_color_format(systolic)
                    self.over_standard_cut += 1
                if diastolic.value >= self.standard_diastolic_blood_pressure:
                    self.change_font_color_format(diastolic)
                    self.over_standard_cut += 1
                if glucose.value >= self.standard_glucose:
                    self.change_font_color_format(glucose)
                    self.over_standard_cut += 1
                if triglycerides.value >= self.standard_triglycerides:
                    self.change_font_color_format(triglycerides)
                    self.over_standard_cut += 1
                if hdlc.value < self.hdlc:
                    self.change_font_color_format(hdlc)
                    self.over_standard_cut += 1
                if self.over_standard_cut >= 3:
                    self.change_font_color_format(people_name, fill_cell=False)

            elif gender.value == '女':
                if waistline.value >= self.standard_waistline-10:
                    self.change_font_color_format(waistline)
                    self.over_standard_cut += 1
                if systolic.value >= self.standard_systolic_blood_pressure:
                    self.change_font_color_format(systolic)
                    self.over_standard_cut += 1
                if diastolic.value >= self.standard_diastolic_blood_pressure:
                    self.change_font_color_format(diastolic)
                    self.over_standard_cut += 1
                if glucose.value >= self.standard_glucose:
                    self.change_font_color_format(glucose)
                    self.over_standard_cut += 1
                if triglycerides.value >= self.standard_triglycerides:
                    self.change_font_color_format(triglycerides)
                    self.over_standard_cut += 1
                if hdlc.value < self.hdlc+10:
                    self.change_font_color_format(hdlc)
                    self.over_standard_cut += 1
                if self.over_standard_cut >= 3:
                    self.change_font_color_format(people_name, fill_cell=False)
        return worksheet
    

    def copy_format_from_sheet1(self, io, src_worksheet:str, dst_worksheet:str):
        """Copy the original sheet header format to specific sheet

        Including cell's fill, font, color, alignment, dimensions.

        Args:
            io (str): Fila path
            src_worksheet (str): Name of process sheet
            dst_worksheet (str): Name of saving sheet
        """
        workbook = openpyxl.load_workbook(io)
        
        ws1 = workbook[src_worksheet]
        ws2 = workbook[dst_worksheet]
        # copy format sheet1 header to sheet2 header
        
        for row in ws1.iter_rows(min_row=1, max_row=1):
            for cell in row:
                # copy cell
                ws2[cell.coordinate].font = copy.copy(cell.font)
                ws2[cell.coordinate].fill = copy.copy(cell.fill)
                ws2[cell.coordinate].border = copy.copy(cell.border)
                ws2[cell.coordinate].alignment = copy.copy(cell.alignment)
                ws2[cell.coordinate].number_format = copy.copy(cell.number_format)
                ws2.row_dimensions[cell.row].height = copy.copy(ws1.row_dimensions[cell.row].height)
                ws2.column_dimensions[cell.column_letter].width = copy.copy(ws1.column_dimensions[cell.column_letter].width)
        
        ws2 = self.set_specific_column_format(worksheet=ws2, eng_column='U')
        time.sleep(0.5)
        ws2 = self.change_date_time(worksheet=ws2, number_of_column=7)
        ws2 = self.place_center(worksheet=ws2)

        # label_over_standard worksheet
        ws1 = self.label_over_standard(worksheet=ws1)
        ws2 = self.label_over_standard(worksheet=ws2)
        workbook.save(io)
        print("saved")


    def process_Metabolic_Syndrome(self, io:str, src_worksheet:str, dst_worksheet:str):
        """篩選代謝症候群的Excel檔案

        Args:
            io (str): Fila path
            src_worksheet (str): Name of process sheet
            dst_worksheet (str): Name of saving sheet
        """

        df = pd.read_excel(io, sheet_name=src_worksheet, engine = 'openpyxl')
        df['超過標準數'] = 0
             
        for i in range(len(df.index)):
            gender = df.iloc[i,self.metabolic_syndrome_column[0]] 
            waistline = df.iloc[i,self.metabolic_syndrome_column[1]]
            systolic = df.iloc[i,self.metabolic_syndrome_column[2]]
            diastolic = df.iloc[i,self.metabolic_syndrome_column[3]]
            glucose = df.iloc[i,self.metabolic_syndrome_column[4]]
            triglycerides = df.iloc[i,self.metabolic_syndrome_column[5]]
            hdlc = df.iloc[i,self.metabolic_syndrome_column[6]]
            
            if gender == '男':
                if waistline >= self.standard_waistline:
                    self.over_standard_cnt += 1
                if systolic >= self.standard_systolic_blood_pressure:
                    self.over_standard_cnt += 1
                if diastolic >= self.standard_diastolic_blood_pressure:
                    self.over_standard_cnt += 1
                if glucose >= self.standard_glucose:
                    self.over_standard_cnt += 1
                if triglycerides >= self.standard_triglycerides:
                    self.over_standard_cnt += 1
                if hdlc < self.hdlc:
                    self.over_standard_cnt += 1
                df.iloc[i, len(df.columns)-1] = self.over_standard_cnt
                
            elif gender == '女':
                if waistline >= self.standard_waistline-10:
                    self.over_standard_cnt += 1
                if systolic >= self.standard_systolic_blood_pressure:
                    self.over_standard_cnt += 1
                if diastolic >= self.standard_diastolic_blood_pressure:
                    self.over_standard_cnt += 1
                if glucose >= self.standard_glucose:
                    self.over_standard_cnt += 1
                if triglycerides >= self.standard_triglycerides:
                    self.over_standard_cnt += 1
                if hdlc < self.hdlc+10:
                    self.over_standard_cnt += 1              
                df.iloc[i, len(df.columns)-1] = self.over_standard_cnt

        for i in range(len(df['超過標準數'])):
            if df['超過標準數'][i] < 3:
                df = df.drop(labels=i, axis=0)
        
        df = df.sort_values(by=['超過標準數'], ascending=False)
        # 建立一個新的 ExcelWriter 物件
        writer = pd.ExcelWriter(io, mode='a', engine='openpyxl', if_sheet_exists='replace')
        
        df.to_excel(writer, sheet_name=dst_worksheet, index=False)
        writer.close() 
        return io, src_worksheet, dst_worksheet

    def main_processdure(self, io, src_worksheet, dst_worksheet):
        filepath, original_sheet, save_sheet = self.process_Metabolic_Syndrome(io=io, src_worksheet=src_worksheet, dst_worksheet=dst_worksheet)
        self.copy_format_from_sheet1(io=filepath, src_worksheet=original_sheet, dst_worksheet=save_sheet)
    

def main():
    Judge_Metabolic_Syndrome().main_processdure('111.xlsx', src_worksheet='健檢資料', dst_worksheet='工作表1')
    print("Finish")

if __name__ == '__main__':

    main()
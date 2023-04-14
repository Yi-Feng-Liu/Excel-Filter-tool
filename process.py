import openpyxl
import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.styles import Alignment, numbers
import copy
import time


class Judge_Metabolic_Syndrome:
    def __init__(self) -> None:
        self.standard_waistline = 90
        self.standard_systolic_blood_pressure = 130
        self.standard_diastolic_blood_pressure = 85
        self.standard_glucose = 100
        self.standard_triglycerides = 150
        self.hdlc = 40
        self.red_font = Font(color='FF0000')
        self.metabolic_syndrome_column = [5, 9, 10, 11, 12, 14, 15]

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
        # 設定儲存格顏色
        fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        worksheet[f'{eng_column}1'].fill = fill
        # 設定儲存格字體顏色
        font = Font(color='FF0000')
        worksheet[f'{eng_column}1'].font = font
        return worksheet


    def set_red_font(self,val):
        return 'color: red'


    def label_over_standard(self, worksheet):
        fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        font = Font(color='FF0000')
        for row in worksheet.iter_rows(min_row=2):
            gender = row[self.metabolic_syndrome_column[0]] 
            waistline = row[self.metabolic_syndrome_column[1]]
            systolic = row[self.metabolic_syndrome_column[2]]
            diastolic = row[self.metabolic_syndrome_column[3]]
            glucose = row[self.metabolic_syndrome_column[4]]
            triglycerides = row[self.metabolic_syndrome_column[5]]
            hdlc = row[self.metabolic_syndrome_column[6]]
            if gender.value == '男':
                if waistline.value >= self.standard_waistline:
                    waistline.fill = fill
                    waistline.font = font
                if systolic.value >= self.standard_systolic_blood_pressure:
                    systolic.fill = fill
                    systolic.font = font
                if diastolic.value >= self.standard_diastolic_blood_pressure:
                    diastolic.fill = fill
                    diastolic.font = font
                if glucose.value >= self.standard_glucose:
                    glucose.fill = fill
                    glucose.font = font
                if triglycerides.value >= self.standard_triglycerides:
                    triglycerides.fill = fill
                    triglycerides.font = font
                if hdlc.value < self.hdlc:
                    hdlc.fill = fill
                    hdlc.font = font

            elif gender.value == '女':
                if waistline.value >= self.standard_waistline-10:
                    waistline.fill = fill
                    waistline.font = font
                if systolic.value >= self.standard_systolic_blood_pressure:
                    systolic.fill = fill
                    systolic.font = font
                if diastolic.value >= self.standard_diastolic_blood_pressure:
                    diastolic.fill = fill
                    diastolic.font = font
                if glucose.value >= self.standard_glucose:
                    glucose.fill = fill
                    glucose.font = font
                if triglycerides.value >= self.standard_triglycerides:
                    triglycerides.fill = fill
                    triglycerides.font = font
                if hdlc.value < self.hdlc+10:
                    hdlc.fill = fill
                    hdlc.font = font
        return worksheet
    

    def copy_format_from_sheet1(self, io):
        workbook = openpyxl.load_workbook(io)
        
        ws1 = workbook["健檢資料"]
        ws2 = workbook["工作表1"]
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


    def process_Metabolic_Syndrome(self, io:str):
        """篩選代謝症候群的Excel檔案

        Args:
            io : file_path
        """

        df = pd.read_excel(io, sheet_name='健檢資料', engine = 'openpyxl')
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
                over_standard_cnt = 0
                if waistline >= self.standard_waistline:
                    over_standard_cnt += 1
                if systolic >= self.standard_systolic_blood_pressure:
                    over_standard_cnt += 1
                if diastolic >= self.standard_diastolic_blood_pressure:
                    over_standard_cnt += 1
                if glucose >= self.standard_glucose:
                    over_standard_cnt += 1
                if triglycerides >= self.standard_triglycerides:
                    over_standard_cnt += 1
                if hdlc < self.hdlc:
                    over_standard_cnt += 1
                df.iloc[i, len(df.columns)-1] = over_standard_cnt
                
            elif gender == '女':
                over_standard_cnt = 0
                if waistline >= self.standard_waistline-10:
                    over_standard_cnt += 1
                if systolic >= self.standard_systolic_blood_pressure:
                    over_standard_cnt += 1
                if diastolic >= self.standard_diastolic_blood_pressure:
                    over_standard_cnt += 1
                if glucose >= self.standard_glucose:
                    over_standard_cnt += 1
                if triglycerides >= self.standard_triglycerides:
                    over_standard_cnt += 1
                if hdlc < self.hdlc+10:
                    over_standard_cnt += 1              
                df.iloc[i, len(df.columns)-1] = over_standard_cnt

        for i in range(len(df['超過標準數'])):
            if df['超過標準數'][i] < 3:
                df = df.drop(labels=i, axis=0)
        
        # df = df.style.applymap(self.set_red_font, subset=pd.IndexSlice[5, '高密度膽固醇']).set_properties(**{'text-align': 'center'})
        # 建立一個新的 ExcelWriter 物件
        writer = pd.ExcelWriter(io, mode='a', engine='openpyxl', if_sheet_exists='replace')
        
        df.to_excel(writer, sheet_name='工作表1', index=False)
        writer.close()        
        
        self.copy_format_from_sheet1(io=io)
    

# Judge_Metabolic_Syndrome().process_Metabolic_Syndrome('111.xlsx')

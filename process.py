import openpyxl
import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.styles import Alignment, numbers
import copy
import time


class Judge_Metabolic_Syndrome:
    def __init__(self, io, select_years, dst_worksheet):
        self.io = io
        self.dst_worksheet = dst_worksheet
        self.select_years = select_years
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
        self.main_procesdure()

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


    def change_font_color_format(self, cell):
        """Change cell color and font color 

        Args:
            cell : the cell coordinate
        """
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
            over_standard_cnt = 0
            if gender.value == '男':
                if waistline.value >= self.standard_waistline:
                    self.change_font_color_format(waistline)
                    over_standard_cnt += 1
                if systolic.value >= self.standard_systolic_blood_pressure:
                    self.change_font_color_format(systolic)
                    over_standard_cnt += 1
                if diastolic.value >= self.standard_diastolic_blood_pressure:
                    self.change_font_color_format(diastolic)
                    over_standard_cnt += 1
                if glucose.value >= self.standard_glucose:
                    self.change_font_color_format(glucose)
                    over_standard_cnt += 1
                if triglycerides.value >= self.standard_triglycerides:
                    self.change_font_color_format(triglycerides)
                    over_standard_cnt += 1
                if hdlc.value < self.hdlc:
                    self.change_font_color_format(hdlc)
                    over_standard_cnt += 1
                if over_standard_cnt >= 3:
                    self.change_font_color_format(people_name)

            elif gender.value == '女':
                if waistline.value >= self.standard_waistline-10:
                    self.change_font_color_format(waistline)
                    over_standard_cnt += 1
                if systolic.value >= self.standard_systolic_blood_pressure:
                    self.change_font_color_format(systolic)
                    over_standard_cnt += 1
                if diastolic.value >= self.standard_diastolic_blood_pressure:
                    self.change_font_color_format(diastolic)
                    over_standard_cnt += 1
                if glucose.value >= self.standard_glucose:
                    self.change_font_color_format(glucose)
                    over_standard_cnt += 1
                if triglycerides.value >= self.standard_triglycerides:
                    self.change_font_color_format(triglycerides)
                    over_standard_cnt += 1
                if hdlc.value < self.hdlc+10:
                    self.change_font_color_format(hdlc)
                    over_standard_cnt += 1
                if over_standard_cnt >= 3:
                    self.change_font_color_format(people_name)
        return worksheet
    

    def copy_title_format(self, ws1, ws2):
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
        return ws2
    
    def copy_format_from_sheet1(self):
        """Copy the original sheet header format to specific sheet

        Including cell's fill, font, color, alignment, dimensions.
        """
        workbook = openpyxl.load_workbook(self.io)
        
        ws1 = workbook['健檢資料']
        ws2 = workbook[self.dst_worksheet]
        # copy format sheet1 header to sheet2 header
        ws2 = self.copy_title_format(ws1=ws1, ws2=ws2)
        ws2 = self.set_specific_column_format(worksheet=ws2, eng_column='U')
        ws2 = self.change_date_time(worksheet=ws2, number_of_column=7)
        ws2 = self.place_center(worksheet=ws2)

        # label_over_standard worksheet
        # ws1 = self.label_over_standard(worksheet=ws1)
        # only process new sheet
        ws2 = self.label_over_standard(worksheet=ws2)
        workbook.save(self.io)
        print("saved")


    def read_file(self):
        df = pd.read_excel(self.io, sheet_name='健檢資料', engine='openpyxl')
        return df
    
    
    def process_Metabolic_Syndrome(self, df:pd.DataFrame):
        """Filter file and add new column named 超過標準數,

        Args:
            df (pd.DataFrame): Original dataframe
        Returns:
            dataframe
        """
        df['超過標準數'] = 0
        df = df[df['年度代碼'].str.startswith(self.select_years)]
             
        for i in range(len(df.index)):
            gender = df.iloc[i,self.metabolic_syndrome_column[0]] 
            waistline = df.iloc[i,self.metabolic_syndrome_column[1]]
            systolic = df.iloc[i,self.metabolic_syndrome_column[2]]
            diastolic = df.iloc[i,self.metabolic_syndrome_column[3]]
            glucose = df.iloc[i,self.metabolic_syndrome_column[4]]
            triglycerides = df.iloc[i,self.metabolic_syndrome_column[5]]
            hdlc = df.iloc[i,self.metabolic_syndrome_column[6]]
            over_standard_cnt = 0
            if gender == '男':
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
        df = df.sort_values(by=['超過標準數'], ascending=False)
        
        return df
        
    
    def save_file_and_copy_title(self, df):
        # 建立一個新的 ExcelWriter 物件
        writer = pd.ExcelWriter(self.io, mode='a', engine='openpyxl', if_sheet_exists='replace')
        df.to_excel(writer, sheet_name=self.dst_worksheet, index=False)
        writer.close() 
        self.copy_format_from_sheet1()


    def main_procesdure(self):
        df = self.read_file()
        df = self.process_Metabolic_Syndrome(df)
        self.save_file_and_copy_title(df)



class Metabolic_Syndrome_From_Summary(Judge_Metabolic_Syndrome):
    def __init__(self, summary_file_fath, select_years, saving_file_path, save_sheet_name):
        super().__init__()
        self.summary_file_fath = summary_file_fath
        self.saving_file_path = saving_file_path
        self.select_years = select_years
        self.save_sheet_name = save_sheet_name
        self.df = pd.read_excel(summary_file_fath, sheet_name='工作表1', engine='openpyxl')
        self.df2 = pd.read_excel(saving_file_path, sheet_name='健檢資料', engine='openpyxl')
        self.goal_column_name = self.df2.columns.tolist()
        self.main_procesdure()
        

    def append_column(self):
        if 'SGPT' not in self.df2.columns:
            self.df2['SGPT'] = 0
        if 'SGOT' not in self.df2.columns:
            self.df2['SGOT'] = 0
        self.goal_column_name = self.df2.columns.tolist()
        
        return self.goal_column_name
    
    def change_column_name(self):
        self.goal_column_name = self.append_column()
        self.df['年度代碼'] = '111年度體檢'
        self.df['部門代號'] = 'X001'
        self.df['健檢過程備註說明'] = ''
        self.df = self.df.drop(labels=0, axis=0)
        
        specific_column = ['年度代碼', '姓名', '工/學號', '部門代號', '部門/科系', '性別', '年齡', '身高', '體重', '腰圍', '收縮壓', '舒張壓', 'AC飯前血糖', 'T-CHO總膽固醇', 'TG三酸甘油脂', 'HDL高密度脂蛋白', 'LDL低密度脂蛋白', '請問您過去一個月內是否有吸菸？', '既往病史', '健檢過程備註說明', 'SGOT血清麩酸草酸轉氨脢', 'SGPT血清麩酸丙銅轉氨脢']

        speific_df = self.df[specific_column].copy()
        speific_df_columns = speific_df.columns.to_list()

        for i in range(len(speific_df_columns)):
            speific_df.rename(columns={speific_df_columns[i]: self.goal_column_name[i]}, inplace=True)  

        # save the sheet 
        # self.process_Metabolic_Syndrome(io='test.xlsx', select_years=111, dst_worksheet='test1')

        return speific_df
    
    def main_procesdure(self):
        # writer = pd.ExcelWriter(self.saving_file_path , mode='a', engine='openpyxl', if_sheet_exists='replace')
        # speific_df.to_excel(writer, sheet_name=self.save_sheet_name, index=False)
        # writer.close()
        speific_df = self.change_column_name()
        speific_df = self.process_Metabolic_Syndrome(speific_df)
        self.save_file_and_copy_title(speific_df)
        print('OK')
        
        
        
    
def main():
    # Judge_Metabolic_Syndrome('test.xlsx', '111', '工作表1')
    Metabolic_Syndrome_From_Summary(summary_file_fath='一般作業總表.xlsx', select_years='111', saving_file_path='test.xlsx', save_sheet_name='test_data')
    # print("Finish")

if __name__ == '__main__':
    main()

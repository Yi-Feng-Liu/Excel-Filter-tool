import openpyxl
import pandas as pd


class Heath_Judge:
    def __init__(self) -> None:
        self.standar_waistline = 90
        self.standar_systolic_blood_pressure = 130
        self.standar_diastolic_blood_pressure = 85
        self.standar_glucose = 100
        self.standar_triglycerides = 150
        self.hdlc = 40
    
    def process_Metabolic_Syndrome(self, io:str):
        """篩選代謝症候群的Excel檔案

        Args:
            io : file_path
        """
        df = pd.read_excel(io, sheet_name='健檢資料')
        print(df.iloc[0,5])
        df['超過標準數'] = 0
        
        # birth_day, height, weight, total_cholesterol. low_cholesterol
        over_standar_cnt_ls = []
        over_standar_cnt = 1
        for i in range(len(df.index)):
            gender = df.iloc[i,5] 
            waistline = df.iloc[i,9]
            systolic = df.iloc[i,10]
            diastolic = df.iloc[i,11]
            glucose = df.iloc[i,12]
            triglycerides = df.iloc[i,13]
            hdlc = df.iloc[i,14]
            if gender == '男':
                if waistline >= self.standar_waistline:
                    over_standar_cnt_ls.append(over_standar_cnt)
                if systolic >= self.standar_systolic_blood_pressure or diastolic >= self.standar_diastolic_blood_pressure:
                    over_standar_cnt_ls.append(over_standar_cnt)
                if glucose >= self.standar_glucose:
                    over_standar_cnt_ls.append(over_standar_cnt)
                if triglycerides < self.standar_triglycerides:
                    over_standar_cnt_ls.append(over_standar_cnt)
                if hdlc < self.hdlc:
                    over_standar_cnt_ls.append(over_standar_cnt) 
                cnt = (sum(over_standar_cnt_ls))   
                df.iloc[i, len(df.columns)-1] = cnt
                
            elif gender == '女':
                if waistline >= self.standar_waistline-10:
                    over_standar_cnt_ls.append(over_standar_cnt)
                if systolic >= self.standar_systolic_blood_pressure or diastolic >= self.standar_diastolic_blood_pressure:
                    over_standar_cnt_ls.append(over_standar_cnt)
                if glucose >= self.standar_glucose:
                    over_standar_cnt_ls.append(over_standar_cnt)
                if triglycerides < self.standar_triglycerides:
                    over_standar_cnt_ls.append(over_standar_cnt)
                if hdlc < self.hdlc+10:
                    over_standar_cnt_ls.append(over_standar_cnt)
                cnt = (sum(over_standar_cnt_ls))              
                df.iloc[i, len(df.columns)-1] = cnt
        print(df.columns)
    
Heath_Judge().process_Metabolic_Syndrome('test.xlsx')

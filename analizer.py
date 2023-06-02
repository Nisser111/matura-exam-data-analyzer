import pandas as pd

class Analiza:
    def __init__(self,excel_file,school):
        self.school = school
        self.df = pd.read_excel(excel_file, sheet_name='2021',header=1)
        self.school_index = self.df.copy()
        self.school_index.columns = self.school_index.iloc[0]
        self.school_index = self.school_index[2:]
        self.school_index = self.school_index.loc[(self.school_index['Nazwa szkoły'] == school)].index
    def get_subject_data(self,subject):
        subject_index = self.df.columns.get_loc(subject)
        if subject=='dla całego egzaminu dojrzałości':
            subject_df = self.df.iloc[self.school_index,subject_index:subject_index+3]
        else:
            subject_df = self.df.iloc[self.school_index,subject_index:subject_index+7]
        subject_df.columns = subject_df.iloc[0]  
        subject_df = subject_df[:1]   
        subject_df.to_excel(subject+'_data.xlsx',sheet_name=subject,header=['Liczba zdających','liczba laureatów/finalistów','zdawalność (%)','średni wynik (%)','odchylenie standardowe (%)','mediana (%)','modalna (%)'],index=False)
    def get_school_data(self,*subject): # *zmiana na zdawało*
        dataframes=[]
        subject = list(subject)
        for i in subject:
            subject_index = self.df.columns.get_loc(i)
            subject_df = self.df.iloc[self.school_index,subject_index]
            subject_df.columns = subject_df.iloc[0]
            subject_df = subject_df[:1]
            dataframes.append(subject_df)
            if i == subject[0]:
                dataframes.append(subject_df)
        result = pd.concat(dataframes, axis=1)
        subject.insert(0,'ogółem')
        result.to_excel(''.join(subject)+'.xlsx',sheet_name=''.join(subject),header=subject)
    def get_school_comparison(self,province):
        schools_list = self.df.copy()
        schools_list.columns = schools_list.iloc[0]
        schools_list = schools_list[2:]
        province_school_list = schools_list.copy()
        province_school_list = province_school_list.loc[(schools_list['województwo - nazwa']==province)]
        technical_schools_list = pd.DataFrame()
        for i,row in province_school_list.iterrows():
            if 'TECHNIKUM' in row['Nazwa szkoły']:
                technical_schools_list = pd.concat([technical_schools_list,row])
        technical_schools_list.to_excel('test.xlsx',sheet_name='test')




analize = Analiza(r'C:\Users\Max\Desktop\analiza\Analiza matur\wyniki_em_szkoly_2021.xlsx','TECHNIKUM NR 2 W ZDUŃSKIEJ WOLI')

analize.get_school_comparison('Łódzkie')


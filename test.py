
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import math
import numpy as np
from matplotlib import pyplot as plt

def Average(list):
    sum = 0
    length = len(list)
    for i in list:
        if math.isnan(i):
            length = length - 1
        else:
            sum = sum+float(i)
    return sum/length

class Analiza:
    def __init__(self,excel_file,school='TECHNIKUM NR 2 W ZDUŃSKIEJ WOLI'):
        self.school = school
        self.df = pd.read_excel(excel_file, sheet_name='2021',header=1)
        self.school_index = self.df.copy()
        self.school_index.columns = self.school_index.iloc[0]
        self.school_index = self.school_index[2:]
        self.school_index = self.school_index.loc[(self.school_index['Nazwa szkoły'] == school)].index
    def get_subject_data(self,subject,get_graph=True):
        subject_index = self.df.columns.get_loc(subject)
        if subject=='dla całego egzaminu dojrzałości':
            subject_df = self.df.iloc[self.school_index,subject_index:subject_index+3]
        else:
            subject_df = self.df.iloc[self.school_index,subject_index:subject_index+7]
        subject_df.columns = subject_df.iloc[0]  
        subject_df = subject_df[:1]  
        if get_graph:
            df_dict = subject_df.copy()
            df_dict = df_dict.to_dict()
            x = ['Liczba zdających','liczba laureatów/finalistów','zdawalność (%)','średni wynik (%)','odchylenie standardowe (%)','mediana (%)','modalna (%)']
            yd = list(df_dict.values())
            y = []
            for i in yd:
                y.append(list(i.values())[0])
            plt.bar(x, y, color='b')
            plt.title('Wykres danych z danego przedmiotu')
            plt.yticks(range(0,200,5))
            plt.xticks(rotation=45)
            plt.grid(True, linestyle='--', linewidth=0.5, color='gray')
            plt.savefig('subject_data_graph.png',bbox_inches ="tight")
        subject_df.to_excel(subject+'_data.xlsx',sheet_name=subject,header=['Liczba zdających','liczba laureatów/finalistów','zdawalność (%)','średni wynik (%)','odchylenie standardowe (%)','mediana (%)','modalna (%)'],index=False)
    def get_school_data(self,get_graph=True,*subject):
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
        result = pd.concat(dataframes, axis=1,ignore_index=True)
        subject.insert(0,'ogółem')
        if get_graph:
            df_dict = result.copy()
            df_dict = result.to_dict()
            x = subject
            yd = list(df_dict.values())
            y = []
            for i in yd:
                if math.isnan(list(i.values())[0]):
                    y.append(0)
                else:
                    y.append(list(i.values())[0])
            plt.bar(x, y, color='b')
            plt.title('Wykres danych z danego przedmiotu')
            plt.yticks(range(0,200,5))
            plt.grid(True, linestyle='--', linewidth=0.5, color='gray')
            plt.xticks(rotation=45)
            plt.savefig('school_data.png',bbox_inches ="tight")
        result.to_excel(''.join(subject)+'.xlsx',sheet_name=''.join(subject),header=subject)
    def get_school_comparison(self,province,subject,OKE_avg,get_graph=True):
        province_schools_list = self.df.copy()
        province_schools_list.columns = province_schools_list.iloc[0]
        province_schools_list = province_schools_list[1:]
        province_schools_list = province_schools_list.loc[(province_schools_list['województwo - nazwa']==province)]
        technical_schools_list = province_schools_list.copy()
        technical_schools_indexes = []
        for i,row in province_schools_list.iterrows():
            if 'TECHNIKUM' in row[6]:
                technical_schools_indexes.append(i)
        technical_schools_list = technical_schools_list[technical_schools_list.index.isin(technical_schools_indexes)]
        subject_index = self.df.columns.get_loc(subject)
        province_schools_avg = []
        for i,row in province_schools_list.iterrows():
            province_schools_avg.append(row[subject_index+2])
        province_schools_avg = Average(province_schools_avg)
        technical_schools_avg = []
        for i,row in technical_schools_list.iterrows():
            technical_schools_avg.append(row[subject_index+2])
        technical_schools_avg = Average(technical_schools_avg)
        result = {self.school:round(self.df.iloc[self.school_index,subject_index+2].values[0],1),
                  'Technika woj. '+province:round(technical_schools_avg,1),
                  'Ogólem woj. '+province:round(province_schools_avg,1),
                  'Ogółem OKE':OKE_avg}
        if get_graph:
            x = list(result.keys())
            y = list(result.values())
            plt.bar(x, y, color='b')
            plt.ylabel("Wynik w procentach")
            plt.title("Wykres pokazujący wyniki"+self.school.lower()+'na tle innych szkół')
            plt.xticks(rotation=45)
            plt.savefig('comparison.png',bbox_inches ="tight")
        result = pd.DataFrame(data=result,index=['Zdało'])
        result.to_excel('comparison.xlsx', sheet_name=self.school+'_comparison')


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("School Data Analysis")
        self.geometry("400x400")

        self.file_path = None
        self.subject_entry = None
        self.oke_avg_entry = None

        self.label = tk.Label(self, text="Select Excel File:")
        self.label.pack(pady=10)

        self.select_button = tk.Button(self, text="Select File", command=self.select_file)
        self.select_button.pack()

        self.subject_label = tk.Label(self, text="Enter Subject:")
        self.subject_label.pack()
        self.subject_entry = tk.Entry(self)
        self.subject_entry.pack()

        self.oke_avg_label = tk.Label(self, text="Enter OKE Average:")
        self.oke_avg_label.pack()
        self.oke_avg_entry = tk.Entry(self)
        self.oke_avg_entry.pack()

        self.analyze_subject_button = tk.Button(self, text="Analyze Subject Data", command=self.analyze_subject_data)
        self.analyze_subject_button.pack(pady=10)
        self.analyze_school_button = tk.Button(self, text="Analyze School Data", command=self.analyze_school_data)
        self.analyze_school_button.pack()
    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    def analyze_subject_data(self):
        subject = self.subject_entry.get()
        analize = Analiza(self.file_path)
        if self.file_path and self.subject_entry.get():
            try:
                analize.get_subject_data(subject)
            except Exception as e:
                print("Error:", e)
        else:
            print("Please select an Excel file and enter a subject and OKE average.")
    def analyze_school_data(self):
        subject = self.subject_entry.get()
        analize = Analiza(self.file_path)
        if self.file_path and self.subject_entry.get():
            try:
                analize.get_school_data(True,subject)
            except Exception as e:
                print("Error:", e)
        else:
            print("Please select an Excel file and enter a subject and OKE average.")

if __name__ == "__main__":
    app = App()
    app.mainloop()


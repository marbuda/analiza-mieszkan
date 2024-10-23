import os
import sys
import pandas as pd
from analizaMieszkana import read_data
import numpy as np

current_path = os.getcwd()
sys.path.append(current_path)

input_path = 'tabele_z_notatkami/'
df_baza = read_data()

work_table_1 = pd.read_excel(input_path+'work_table_1.xlsx', usecols=['l.p', 'Notatka']).rename(columns = {'Notatka': 'Notatka_wt1'})
work_table_2 = pd.read_excel(input_path+'work_table_2.xlsx', usecols=['l.p', 'Notatka']).rename(columns = {'Notatka': 'Notatka_wt2'})



df_baza = df_baza.merge(work_table_1, how = 'left', on = 'l.p')
df_baza = df_baza.merge(work_table_2, how = 'left', on = 'l.p')

#df_baza[['Notatka_wt1', 'Notatka_wt2']].fillna("", inplace = True)

def _update_notes_in_row(row):
    if pd.notna(row['Notatka_wt1']):
        return row['Notatka_wt1']
    elif pd.notna(row['Notatka_wt2']):
        return row['Notatka_wt2']
    else:
        return row['Notatka']


df_baza['Notatka_new'] = df_baza.apply(_update_notes_in_row, axis = 1)
save_tables([df_baza], ['Baza danych'], file_name='output/Baza danych output.xlsx')

df_baza.to_excel('test.xlsx')
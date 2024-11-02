#import importlib
import os
import sys
import pandas as pd
import analizaMieszkana, Formatowanie, utils
# importlib.reload(analizaMieszkana)
# importlib.reload(Formatowanie)
# importlib.reload(utils)
from analizaMieszkana import read_data, calculate_max_price, create_work_table_1, create_work_table_2, save_tables, \
    find_duplicates, update_notes
from Formatowanie import format_file


df_baza = read_data()
input_path = 'tabele_z_notatkami/'
work_table_1 = pd.read_excel(input_path + 'work_table_1.xlsx', sheet_name='sheet1', usecols=['l.p', 'Tabela dzwonienie']).rename(
    columns={'Tabela dzwonienie': 'Tabela_dzwonienie_wt1'})
work_table_2 = pd.read_excel(input_path + 'work_table_1.xlsx', sheet_name='sheet2', usecols=['l.p', 'Tabela dzwonienie']).rename(
    columns={'Tabela dzwonienie': 'Tabela_dzwonienie_wt2'})
work_table_3 = pd.read_excel(input_path + 'work_table_2.xlsx', sheet_name='sheet1', usecols=['l.p', 'Tabela dzwonienie']).rename(
    columns={'Tabela dzwonienie': 'Tabela_dzwonienie_wt3'})

def _update_dzwonienie_in_row(row):
    #sprawdza czy są nowe flagi tabela dzwonienie w tabelach roboczych, jesli tak, to orgyginalne są nadpisywane
    if pd.notna(row['Tabela_dzwonienie_wt1']):
        return row['Tabela_dzwonienie_wt1']
    elif pd.notna(row['Tabela_dzwonienie_wt2']):
        return row['Tabela_dzwonienie_wt2']
    elif pd.notna(row['Tabela_dzwonienie_wt3']):
        return row['Tabela_dzwonienie_wt3']
    else:
        return row['Tabela dzwonienie']


df_baza = df_baza.merge(work_table_1, how='left', on='l.p')
df_baza = df_baza.merge(work_table_2, how='left', on='l.p')
df_baza = df_baza.merge(work_table_3, how='left', on='l.p')

df_baza['Tabela dzwonienie'] = df_baza.apply(_update_dzwonienie_in_row, axis=1)
df_baza.drop(['Tabela_dzwonienie_wt1', 'Tabela_dzwonienie_wt2', 'Tabela_dzwonienie_wt3'], axis = 1, inplace= True)

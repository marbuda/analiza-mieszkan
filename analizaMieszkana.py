import os
import sys

import numpy as np
import openpyxl
import pandas as pd
import math

current_path = os.getcwd()
sys.path.append(current_path)


def read_data(path='input/Baza danych.xlsx'):
    df_baza = pd.read_excel(path)

    return df_baza


def _read_additional_data(path='input/Dane dodatkowe.xls'):
    df_dodatkowe = pd.read_excel(path, header=None)
    df_dodatkowe.columns = ['Category', 'Value']

    zysk = df_dodatkowe[df_dodatkowe['Category'] == 'Zysk']['Value'].iloc[0]
    przygotowanie = df_dodatkowe[df_dodatkowe['Category'] == 'Przygotowanie']['Value'].iloc[0]
    notariusz = df_dodatkowe[df_dodatkowe['Category'] == 'Notariusz']['Value'].iloc[0]
    posrednik = df_dodatkowe[df_dodatkowe['Category'] == 'Pośrednik']['Value'].iloc[0]
    zmniejszenie_rynkowe = df_dodatkowe[df_dodatkowe['Category'] == 'Zmniejszenie rynkowej']['Value'].iloc[0]
    wysokosc_negocjacji = df_dodatkowe[df_dodatkowe['Category'] == 'Wysokość negocjacji']['Value'].iloc[0]

    nominal = [zysk, przygotowanie]
    perc = [notariusz, posrednik]

    return nominal, perc, zmniejszenie_rynkowe, wysokosc_negocjacji


def _classify_value(val):
    if val <= 34:
        return '<=34'
    elif 34 < val <= 46:
        return '>34 & <=46'
    else:
        return '>46'


def calculate_max_price(df_baza, target, new_col_name):
    df_baza['m2_category'] = df_baza['m2'].apply(_classify_value)

    grouped_df = df_baza.groupby([target, 'm2_category'])['Cena/m2'].mean().reset_index()
    df_baza = pd.merge(df_baza, grouped_df, on=[target, 'm2_category'], suffixes=('', '_mean_' + target))

    nominal, perc, zmniejszenie_rynkowe, wysokosc_negocjacji = _read_additional_data()
    df_baza[new_col_name] = ((df_baza['Cena/m2_mean_' + target] - zmniejszenie_rynkowe) * df_baza['m2'] - sum(
        nominal)) / (1 + sum(perc))
    #round down result
    df_baza[new_col_name] = df_baza[new_col_name]/1000
    df_baza[new_col_name] = df_baza[new_col_name].apply(math.floor) * 1000


    df_baza['additional_profit_' + target] = df_baza[new_col_name] - df_baza['Cena']

    return df_baza


def find_duplicates(df_baza, columns, index_column = 'l.p'):
    # Create a 'Duplicate' column initialized with empty strings
    df_baza['Duplicate'] = ""

    # Iterate through DataFrame to find duplicates
    for index, row in df_baza.iterrows():
        # Get current row's values in specified columns
        current_values = tuple(row[col] for col in columns)

        # Find duplicates in the DataFrame based on specified columns
        duplicates = df_baza[df_baza[columns].apply(tuple, axis=1) == current_values]

        # If there are duplicates, exclude the current row and get 'l.p' values
        if len(duplicates) > 1:  # More than one means there are duplicates
            duplicate_values = duplicates[index_column].tolist()
            duplicate_values.remove(row[index_column])  # Remove current row's 'l.p' value
            df_baza.at[index, 'Duplicate'] = ', '.join(map(str, duplicate_values))

    return df_baza


def save_tables(dfs, sheet_names, file_name):
    # Write clickable hyperlinks to column A
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='w') as writer:
        for df, sheet_name in zip(dfs, sheet_names):
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Data saved to {file_name} successfully.")

    wb = openpyxl.load_workbook(file_name)

    # Convert the 'Link' column values to clickable hyperlinks for each sheet
    for sheet_name, df in zip(sheet_names, dfs):
        ws = wb[sheet_name]
        #link_col_text = df.columns.get_loc("Link") + 1  # Find the 'Link' column index, 1-based for openpyxl
        link_col_hyper = df.columns.get_loc("Link2") + 1

        for row_num, link in enumerate(df["Link"], start=2):  # Start from row 2 to skip headers
            cell = ws.cell(row=row_num, column=link_col_hyper)
            cell.value = f'=HYPERLINK("{link}", "{link}")'  # Use the link itself as the clickable text

    wb.save(file_name)

    return True


def create_work_table_1(df_baza):

    # 1) Aktualne z zyskiem powyzje -35000 LUB flaga
    nominal, perc, zmniejszenie_rynkowe, wysokosc_negocjacji = _read_additional_data()
    work_table_1_sheet_1 = df_baza[
        ((df_baza['Sprzedane?'] == 'NIE') & (df_baza['additional_profit_podzielnica'] >= wysokosc_negocjacji)) | (
                df_baza['Tabela dzwonienie'] == 'TAK')]

    work_table_1_sheet_1 = work_table_1_sheet_1[work_table_1_sheet_1['Tabela dzwonienie'] != 'NIE']

    work_table_1_sheet_1 = work_table_1_sheet_1[['l.p', 'Data dodania', 'Link', 'Link2', 'Adres',
                                                 'podzielnica', 'Pośrednik?', 'Telefon',
                                                 'Cena', 'Cena/m2', 'm2', 'Pokoje', 'Piętro',
                                                 'Rodzaj mieszkania', 'Tabela dzwonienie',
                                                 'Max Cena Kupna', 'Notatka']]

    # 2) Aktualne dodane powyzej 3 miesiecy temu i zysk powyzej -50000 oraz poniżej wysokosc negocjacji
    nominal, perc, zmniejszenie_rynkowe, wysokosc_negocjacji = _read_additional_data()
    three_months_ago = pd.Timestamp.today() - pd.DateOffset(months=3)
    work_table_1_sheet_2 = df_baza[(df_baza['Sprzedane?'] == 'NIE') & (df_baza['Data dodania'] < three_months_ago) & (
            df_baza['additional_profit_podzielnica'] >= -50000) & (df_baza['additional_profit_podzielnica'] < wysokosc_negocjacji)]

    work_table_1_sheet_2 = work_table_1_sheet_2[work_table_1_sheet_2['Tabela dzwonienie'] != 'NIE']

    work_table_1_sheet_2 = work_table_1_sheet_2[['l.p', 'Data dodania', 'Link', 'Link2', 'Adres',
                                                 'podzielnica', 'Pośrednik?', 'Telefon',
                                                 'Cena', 'Cena/m2', 'm2', 'Pokoje', 'Piętro',
                                                 'Rodzaj mieszkania', 'Tabela dzwonienie',
                                                 'Max Cena Kupna', 'Notatka']]


    work_tables_1 = [work_table_1_sheet_1, work_table_1_sheet_2]

    return work_tables_1


def create_work_table_2(df_baza):

    # 1) Wygalsny w zeszlym miesiacu i zysk powyzej -50000
    df_baza['Data wygaśnięcia'] = pd.to_datetime(df_baza['Data wygaśnięcia'], errors='coerce')
    one_month_ago = pd.Timestamp.today() - pd.DateOffset(months=1)
    work_table_2_sheet_1 = df_baza[
        ((df_baza['Sprzedane?'] == 'TAK') & (df_baza['additional_profit_podzielnica'] >= -50000))]
    work_table_2_sheet_1 = work_table_2_sheet_1[df_baza['Data wygaśnięcia'] > one_month_ago]

    work_table_2_sheet_1 = work_table_2_sheet_1[work_table_2_sheet_1['Tabela dzwonienie'] != 'NIE']

    work_table_2_sheet_1 = work_table_2_sheet_1[['l.p', 'Data dodania', 'Link', 'Link2', 'Adres',
                                                 'podzielnica', 'Pośrednik?', 'Telefon',
                                                 'Cena', 'Cena/m2', 'm2', 'Pokoje', 'Piętro',
                                                 'Rodzaj mieszkania', 'Tabela dzwonienie',
                                                 'Max Cena Kupna', 'Notatka']]
    work_tables_2 = [work_table_2_sheet_1]

    return work_tables_2

def _update_notes_in_row(row):
    #sprawdza czy są nowe notatki w tabelach roboczych, jesli tak, to orgyginalne są nadpisywane
    if pd.notna(row['Notatka_wt1']):
        return row['Notatka_wt1']
    elif pd.notna(row['Notatka_wt2']):
        return row['Notatka_wt2']
    elif pd.notna(row['Notatka_wt3']):
        return row['Notatka_wt3']
    else:
        return row['Notatka']



def update_notes(input_path = 'aktualizacja_notatek_i_mieszkan/', output_path = 'input/'):
    df_baza = read_data()

    work_table_1 = pd.read_excel(input_path + 'work_table_1.xlsx', usecols=['l.p', 'Notatka']).rename(
        columns={'Notatka': 'Notatka_wt1'})
    work_table_2 = pd.read_excel(input_path + 'work_table_1.xlsx', sheet_name= 'sheet2', usecols=['l.p', 'Notatka']).rename(
        columns={'Notatka': 'Notatka_wt2'})
    work_table_3 = pd.read_excel(input_path + 'work_table_2.xlsx', sheet_name= 'sheet1', usecols=['l.p', 'Notatka']).rename(
        columns={'Notatka': 'Notatka_wt3'})


    df_baza = df_baza.merge(work_table_1, how='left', on='l.p')
    df_baza = df_baza.merge(work_table_2, how='left', on='l.p')
    df_baza = df_baza.merge(work_table_3, how='left', on='l.p')

    df_baza['Notatka'] = df_baza.apply(_update_notes_in_row, axis=1)
    df_baza.drop(['Notatka_wt1', 'Notatka_wt2', 'Notatka_wt3'], axis = 1, inplace= True)
    save_tables([df_baza], ['Baza danych'], file_name=output_path+'Baza danych.xlsx')


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


def update_tabela_dzwonienie(input_path = 'aktualizacja_notatek_i_mieszkan/', output_path = 'input/'):
    df_baza = read_data()

    work_table_1 = pd.read_excel(input_path + 'work_table_1.xlsx', sheet_name='sheet1',
                                 usecols=['l.p', 'Tabela dzwonienie']).rename(
        columns={'Tabela dzwonienie': 'Tabela_dzwonienie_wt1'})
    work_table_2 = pd.read_excel(input_path + 'work_table_1.xlsx', sheet_name='sheet2',
                                 usecols=['l.p', 'Tabela dzwonienie']).rename(
        columns={'Tabela dzwonienie': 'Tabela_dzwonienie_wt2'})
    work_table_3 = pd.read_excel(input_path + 'work_table_2.xlsx', sheet_name='sheet1',
                                 usecols=['l.p', 'Tabela dzwonienie']).rename(
        columns={'Tabela dzwonienie': 'Tabela_dzwonienie_wt3'})


    df_baza = df_baza.merge(work_table_1, how='left', on='l.p')
    df_baza = df_baza.merge(work_table_2, how='left', on='l.p')
    df_baza = df_baza.merge(work_table_3, how='left', on='l.p')

    df_baza['Tabela dzwonienie'] = df_baza.apply(_update_dzwonienie_in_row, axis=1)
    df_baza.drop(['Tabela_dzwonienie_wt1', 'Tabela_dzwonienie_wt2', 'Tabela_dzwonienie_wt3'], axis=1, inplace=True)
    save_tables([df_baza], ['Baza danych'], file_name=output_path+'Baza danych.xlsx')
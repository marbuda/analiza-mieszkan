import os
import sys
import pandas as pd

current_path = os.getcwd()
sys.path.append(current_path)


def read_data(path='input/Baza danych.xls'):
    df_baza = pd.read_excel(path)

    return df_baza


def _read_additional_data(path='input/Dane dodatkowe.xls'):
    df_dodatkowe = pd.read_excel(path, header=None)
    df_dodatkowe.columns = ['Category', 'Value']

    zysk = df_dodatkowe[df_dodatkowe['Category'] == 'Zysk']['Value'].iloc[0]
    przygotowanie = df_dodatkowe[df_dodatkowe['Category'] == 'Przygotowanie']['Value'].iloc[0]
    notariusz = df_dodatkowe[df_dodatkowe['Category'] == 'Notariusz']['Value'].iloc[0]
    posrednik = df_dodatkowe[df_dodatkowe['Category'] == 'Pośrednik']['Value'].iloc[0]

    nominal = [zysk, przygotowanie]
    perc = [notariusz, posrednik]

    return nominal, perc


def _classify_value(val):
    if val <= 34:
        return '<=34'
    elif 34 < val <= 46:
        return '>34 & <=46'
    else:
        return '>46'


def calculate_max_price(df_baza, target):
    df_baza['m2_category'] = df_baza['m2'].apply(_classify_value)

    grouped_df = df_baza.groupby([target, 'm2_category'])['Cena/m2'].mean().reset_index()
    df_baza = pd.merge(df_baza, grouped_df, on=[target, 'm2_category'], suffixes=('', '_mean_' + target))

    nominal, perc = _read_additional_data()
    df_baza['maksymalna_cena_' + target] = ((df_baza['Cena/m2_mean_' + target] - 100) * df_baza['m2'] - sum(
        nominal)) / (1 + sum(perc))
    df_baza['additional_profit_' + target] = df_baza['maksymalna_cena_' + target] - df_baza['Cena']

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
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='w') as writer:
        for df, sheet_name in zip(dfs, sheet_names):
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Data saved to {file_name} successfully.")

    return True


def create_work_table_1(df_baza):
    # TODO change additional_profit to target column?

    # 1) Aktualne z zyskiem powyzje -35000 LUB flaga
    work_table_1_sheet_1 = df_baza[
        ((df_baza['Sprzedane?'] == 'NIE') & (df_baza['additional_profit_Dzielnica'] >= -35000)) | (
                df_baza['Tabela dzwonienie'] == 'TAK')]

    # 2) Aktualne dodane powyzej 3 miesiecy temu i zysk powyzej -50000
    three_months_ago = pd.Timestamp.today() - pd.DateOffset(months=3)
    work_table_1_sheet_2 = df_baza[(df_baza['Sprzedane?'] == 'NIE') & (df_baza['Data dodania'] < three_months_ago) & (
            df_baza['additional_profit_Dzielnica'] >= -50000)]

    work_tables_1 = [work_table_1_sheet_1, work_table_1_sheet_2]

    return work_tables_1


def create_work_table_2(df_baza):
    # TODO change additional_profit to target column?

    # 1) Wygalsny w zeszlym miesiacu i zysk powyzej -50000
    df_baza['Data wygaśnięcia'] = pd.to_datetime(df_baza['Data wygaśnięcia'], errors='coerce')
    one_month_ago = pd.Timestamp.today() - pd.DateOffset(months=1)
    work_table_2_sheet_1 = df_baza[
        ((df_baza['Sprzedane?'] == 'TAK') & (df_baza['additional_profit_Dzielnica'] >= -50000))]
    work_table_2_sheet_1 = work_table_2_sheet_1[df_baza['Data wygaśnięcia'] > one_month_ago]

    work_tables_2 = [work_table_2_sheet_1]

    return work_tables_2

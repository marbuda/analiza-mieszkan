#import importlib
import os
import sys
import analizaMieszkana, Formatowanie, utils
# importlib.reload(analizaMieszkana)
# importlib.reload(Formatowanie)
# importlib.reload(utils)
from analizaMieszkana import read_data, calculate_max_price, create_work_table_1, create_work_table_2, save_tables, \
    find_duplicates, update_notes
from Formatowanie import format_file

sys.path.append(os.getcwd())


df_baza = read_data()
df_baza = calculate_max_price(df_baza, 'Dzielnica')
df_baza = calculate_max_price(df_baza, 'podzielnica')
df_baza = find_duplicates(df_baza, ['Adres', 'm2', 'Piętro'])
save_tables([df_baza], ['Baza danych'], file_name='output/Baza danych output.xlsx')

work_tables_1 = create_work_table_1(df_baza)
save_tables(work_tables_1, ['sheet1', 'sheet2'], file_name='output/work_table_1.xlsx')

work_tables_2 = create_work_table_2(df_baza)
save_tables(work_tables_2, ['sheet1'], file_name='output/work_table_2.xlsx')

format_file('output/Baza danych output.xlsx')
# import datetime
# from typing import Union
#
# import random
# from contextlib import suppress
# from pathlib import Path
# from time import sleep
#
# from openpyxl.reader.excel import load_workbook
# from sqlalchemy import create_engine
# from sqlalchemy.orm import sessionmaker
#
# from config import engine_kwargs
# from models import Base, add_to_db, get_all_data, update_in_db, get_all_data_by_status
from tools.app import App
# import pyautogui as pag
#
# import pandas as pd
# import datetime
# import os
# from time import sleep
#
# import psycopg2

# from tools.xlsx_fix import fix_excel_file_error

# fix_excel_file_error(Path(r'C:\Users\Abdykarim.D\Documents\lolus.xlsx'))
#
# df = pd.read_excel(r'C:\Users\Abdykarim.D\Documents\lolus.xlsx', header=3)
#
# print(df[df['ПометкаУдаления'] == 'Нет']['Номер'])

# fix_excel_file_error(Path(r'C:\Users\Abdykarim.D\Documents\subconto.xlsx'))
# main_df = pd.read_excel(r'C:\Users\Abdykarim.D\Documents\subconto.xlsx', header=8)
#
# main_df = main_df.drop(['Unnamed: 0'], axis=1)
#
# main_df.columns = ['Subconto', 'Debit', 'Credit', 'Debit.1', 'Credit.1', 'Debit.2', 'Credit.2']
#
# contragent = 'ТОО Kobako Tsuvari'
# summ = 9173
# branch = None
# contract = None
#
# print(main_df)
#
# all_contracts = main_df[main_df['Subconto'] == contragent]
#
# print(all_contracts.index)
#
# if len(all_contracts) == 0:
#
#     pass
#
# if len(all_contracts) == 1:
#
#     ind = all_contracts.index[0]
#
#     print(main_df['Subconto'].iloc[ind - 2], main_df['Debit.2'].iloc[ind - 2])
#     branch = main_df['Subconto'].iloc[ind - 2]
#     contract = main_df['Subconto'].iloc[ind - 1]
#
# if len(all_contracts) > 1:
#
#     for i in all_contracts.index:
#         print(main_df['Subconto'].iloc[i - 2], main_df['Debit.2'].iloc[i - 2])
#         if float(summ) == float(main_df['Debit.2'].iloc[i - 2]):
#             branch = main_df['Subconto'].iloc[i - 2]
#             contract = main_df['Subconto'].iloc[i - 1]
#
# print('-----------------------')
# print(branch)
# print(contract)

# Session = sessionmaker()
#
# engine = create_engine(
#     'postgresql+psycopg2://{username}:{password}@{host}:{port}/{base}'.format(**engine_kwargs),
#     connect_args={'options': '-csearch_path=robot'}
# )
# Base.metadata.create_all(bind=engine)
# Session.configure(bind=engine)
# session = Session()
#
# second_excel_path = r'C:\Users\Abdykarim.D\Documents\chpokus.xlsx'
# main_df = pd.read_excel(second_excel_path)
#
# ids = []
# summs = []
# contragents = []
# branches = []
#
# # for ind in range(len(main_df)):
# #     ids.append(main_df['Номер'].iloc[ind])
# #     summs.append(main_df['Сумма документа'].iloc[ind])
# #     contragents.append(main_df['Контрагент'].iloc[ind])
# #     branches.append(main_df['Филиал'].iloc[ind])
# #
# #     add_to_db(session, 'new', main_df['Дата'].iloc[ind], str(main_df['Номер'].iloc[ind]), main_df['Сумма документа'].iloc[ind], main_df['Контрагент'].iloc[ind],
# #               main_df['Филиал'].iloc[ind], None, None, None, None, None)
#
# rows = get_all_data_by_status(session, 'processing')
#
# for ind, row in enumerate(rows):
#     update_in_db(session, row, 'processing', row.branch, None,
#                  False, None, False)
#
# session.close()


# --------------------------------------------------------------------------------------------------------------------------------
# contragent = 'ПРОКТЕР ЭНД ГЭМБЛ КАЗАХСТАН ДИСТРИБЬЮШН ТОО (13623)'
# cur_date = '16.03.2023'
# search_date = datetime.datetime.strptime(cur_date, '%d.%m.%Y') - datetime.timedelta(days=1)
# summ = 13459033.0
#
# print(cur_date, search_date)
#
# if contragent == 'ПРОКТЕР ЭНД ГЭМБЛ КАЗАХСТАН ДИСТРИБЬЮШН ТОО (13623)':
#
#     df = pd.read_excel(r'C:\Users\Abdykarim.D\Documents\проктер.xlsx')
#
#     # print(df[df['Payment Date'] == search_date.strftime('%Y-%m-%d')])
#
#     filtered_df = df[df['Payment Date'] == search_date.strftime('%Y-%m-%d')]
#
#     all_invoices = []
#
#     for i in range(len(filtered_df)):
#
#         print(filtered_df['Invoice Number'].iloc[i], filtered_df['Invoice Amount'].iloc[i])
#
#
# if contragent == 'КИМБЕРЛИ-КЛАРК КАЗАХСТАН ТОО (3199)':
#
#     df = pd.read_excel(r'C:\Users\Abdykarim.D\Documents\кимберли.xlsx')


# b = datetime.datetime(2024, 2, 19, 15, 21, 0)
# a = datetime.datetime(2024, 2, 16, 11, 21, 0)
#
# one = a.strftime("%d.%m.%Y %H:%M:%S").split('.')[0]
# two = b.strftime("%d.%m.%Y %H:%M:%S").split('.')[0]
# # two = datetime.datetime.now().strftime("%d.%m.%Y").split('.')[0]
#
# print((b - a).total_seconds() / 86400)
#
# print(int(two) - int(one))
#
# print((b - a))


# import pandas as pd
# c = 0
# for file in os.listdir(r'C:\Users\Abdykarim.D\Documents\BI'):
#
#     if 'xls' not in file:
#         continue
#     df = pd.read_excel(os.path.join(r'C:\Users\Abdykarim.D\Documents\BI', file))
#
#     df_notna = df[(df['ФИО'].str.len() > 1) & (df['ФИО'].notna())]
#     df_notna['Окончание смены_1'] = df_notna['Окончание смены'].apply(lambda x: datetime.datetime.strptime(x, '%d.%m.%Y %H:%M:%S').strftime('%d.%m.%Y'))
#
#     dates_list = [datetime.datetime.strptime(date_str, '%d.%m.%Y') for date_str in df_notna['Окончание смены_1'].unique()]
#
#     oldest_date = min(dates_list)
#
#     oldest_date_str = oldest_date.strftime('%d.%m.%Y')
#
#     for dates in df_notna['Окончание смены_1'].unique():
#         # print('OLD:',oldest_date_str)
#         if oldest_date_str == dates and os.path.isfile(fr'C:\Users\Abdykarim.D\Documents\BI\A\outsourcingshifts {str(dates.split()[0]).replace(".", "_")}.xlsx'):
#             df1 = pd.read_excel(fr'C:\Users\Abdykarim.D\Documents\BI\A\outsourcingshifts {str(dates.split()[0]).replace(".", "_")}.xlsx')
#             # print('FOUND HERE!!!', dates, file, f'SAVING AS: {str(dates.split()[0])}')
#             df_notna = pd.concat([df1, df_notna])
#             df_notna = df_notna.drop_duplicates()
#             # df__ = df_notna[df_notna['Окончание смены_1'] == dates].copy().drop(columns=['Окончание смены_1'])
#             df_notna[df_notna['Окончание смены_1'] == dates].to_excel(fr'C:\Users\Abdykarim.D\Documents\BI\A\outsourcingshifts {str(dates.split()[0]).replace(".", "_")}.xlsx', index=False)
#
#         else:
#             # print('---', dates, file, f'SAVING AS: {str(dates.split()[0])}')
#
#             # df__ = df_notna[df_notna['Окончание смены_1'] == dates].copy().drop(columns=['Окончание смены_1'])
#             df_notna[df_notna['Окончание смены_1'] == dates].to_excel(fr'C:\Users\Abdykarim.D\Documents\BI\A\outsourcingshifts {str(dates.split()[0]).replace(".", "_")}.xlsx', index=False)

app = App('')

a = app.find_element({"title": "Записать", "class_name": "", "control_type": "Button",
                      "visible_only": True, "enabled_only": True, "found_index": 0})
print(app.parent.element.children())
print()



# import pandas as pd
# conn = psycopg2.connect(dbname='adb', host='172.16.10.22', port='5432',
#                         user='rpa_robot', password='Qaz123123+')
#
# cur = conn.cursor()
#
# cur.execute(f"""select distinct(name_sale_object_for_print) from dwh_data.dim_branches_src dbs""")
# df = pd.DataFrame(cur.fetchall())
# conn.close()
#
# all_branches = []
#
# for branchos in df[df.columns[0]]:
#     all_branches.append(branchos.replace(' ', '').lower())
#
# all_branches.append('ТОО "Magnum Cash&Carry"'.replace(' ', '').lower())
#
# main_df = pd.read_excel(r'C:\Users\Abdykarim.D\Desktop\SUBCONTO.xlsx', header=8)
#
# main_df = main_df.drop(['Unnamed: 0'], axis=1)
#
# main_df.columns = ['Subconto', 'Debit', 'Credit', 'Debit.1', 'Credit.1', 'Debit.2', 'Credit.2']
#
# contragent = 'GAMMA D`ORO ТОО (20170)'
# summ = 666
# branch = None
# invoice = None
#
# row = main_df[main_df['Debit.2'] == summ]
# if isinstance(row['Subconto'].iloc[0], int):
#     for ind_ in range(row.index[0], -1, -1):
#         print(main_df['Subconto'].iloc[ind_])
#         if (invoice is None and isinstance(main_df['Subconto'].iloc[ind_], str)
#                 and main_df['Subconto'].iloc[ind_] != contragent)\
#                 and main_df['Subconto'].iloc[ind_].replace(' ', '').lower() not in all_branches:
#             invoice = main_df['Subconto'].iloc[ind_]
#
#         if isinstance(main_df['Subconto'].iloc[ind_], str) and main_df['Subconto'].iloc[ind_].replace(' ', '').lower() in all_branches:
#             branch = main_df['Subconto'].iloc[ind_]
#             break
#
# elif isinstance(row['Subconto'].iloc[0], str):
#     for ind_ in range(row.index[0] + 1, -1, -1):
#         print(main_df['Subconto'].iloc[ind_])
#         if (invoice is None and isinstance(main_df['Subconto'].iloc[ind_], str)
#                 and main_df['Subconto'].iloc[ind_] != contragent)\
#                 and main_df['Subconto'].iloc[ind_].replace(' ', '').lower() not in all_branches:
#             invoice = main_df['Subconto'].iloc[ind_]
#
#         if isinstance(main_df['Subconto'].iloc[ind_], str) and main_df['Subconto'].iloc[ind_].replace(' ', '').lower() in all_branches:
#             branch = main_df['Subconto'].iloc[ind_]
#             break
#
# print(row.index[0])
#
# print(f'\n------------------------------------------------------')
# print(f'BRANCH: {branch}')
# print(f'INVOICE: {invoice}')
#

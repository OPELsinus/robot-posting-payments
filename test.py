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
# from tools.app import App
# import pyautogui as pag
#
# import pandas as pd
import datetime
import os

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


import pandas as pd
c = 0
for file in os.listdir(r'C:\Users\Abdykarim.D\Documents\BI'):
    if 'xls' not in file:
        continue
    df = pd.read_excel(os.path.join(r'C:\Users\Abdykarim.D\Documents\BI', file))

    df_notna = df[(df['ФИО'].str.len() > 1) & (df['ФИО'].notna())]
    df_notna['Окончание смены'] = df_notna['Окончание смены'].apply(lambda x: datetime.datetime.strptime(x, '%d.%m.%Y %H:%M:%S').strftime('%d.%m.%Y'))
    print(len(df_notna))

    for dates in df_notna['Окончание смены'].unique():
        print(len(df_notna[df_notna['Окончание смены'] == dates]))
        c += len(df_notna[df_notna['Окончание смены'] == dates])
        if os.path.isfile(fr'C:\Users\Abdykarim.D\Documents\BI\A\outsourcingshifts {str(dates.split()[0]).replace(".", "_")}.xlsx'):
            df1 = pd.read_excel(fr'C:\Users\Abdykarim.D\Documents\BI\A\outsourcingshifts {str(dates.split()[0]).replace(".", "_")}.xlsx')
            df_notna = pd.concat([df1, df_notna])
            df_notna[df_notna['Окончание смены'] == dates].to_excel(fr'C:\Users\Abdykarim.D\Documents\BI\A\outsourcingshifts {str(dates.split()[0]).replace(".", "_")}.xlsx', index=False)

        else:
            df_notna[df_notna['Окончание смены'] == dates].to_excel(fr'C:\Users\Abdykarim.D\Documents\BI\A\outsourcingshifts {str(dates.split()[0]).replace(".", "_")}.xlsx', index=False)
print(c)

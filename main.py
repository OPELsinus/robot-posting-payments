import datetime
import os
import sys
import traceback
from contextlib import suppress
from pathlib import Path
from time import sleep

import pyautogui as pag
from pywinauto import keyboard
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from config import logger, process_list_path, production_calendar_path, form_document_path, engine_kwargs, main_executor, ip_address
from models import Base, add_to_db, get_all_data, get_all_data_by_status, update_in_db
from tools.app import App
from tools.odines import Odines
from tools.process import kill_process_list

import pandas as pd

from tools.xlsx_fix import fix_excel_file_error


def save_the_report(app: Odines, savepath: str):
    app.open('Файл', 'Сохранить как...')

    app.parent_switch({"title": "Сохранение", "class_name": "#32770", "control_type": "Window",
                       "visible_only": True, "enabled_only": True, "found_index": 0})

    app.find_element(
        {"title": "Имя файла:", "class_name": "Edit", "control_type": "Edit", "visible_only": True,
         "enabled_only": True, "found_index": 0}).type_keys(savepath)

    sleep(.1)
    for _ in range(5):
        try:
            while not app.wait_element(
                    {"title": "Лист Excel2007-... (*.xlsx)", "class_name": "", "control_type": "ListItem",
                     "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1):
                app.find_element(
                    {"title": "Тип файла:", "class_name": "AppControlHost", "control_type": "ComboBox",
                     "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=5).click(double=True)
                pag.click()

            # * save it as xlsx
            app.find_element(
                {"title": "Лист Excel2007-... (*.xlsx)", "class_name": "", "control_type": "ListItem",
                 "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=2).click()
            break

        except:
            pass
    sleep(.5)
    # * click save button
    for _ in range(10):

        try:
            app.find_element(
                {"title": "Сохранить", "class_name": "Button", "control_type": "Button", "visible_only": True,
                 "enabled_only": True, "found_index": 0}, timeout=2).click()
            sleep(0.3)
            if app.wait_element(
                    {"title": "Сохранить", "class_name": "Button", "control_type": "Button", "visible_only": True,
                     "enabled_only": True, "found_index": 0}, timeout=2):
                app.find_element(
                    {"title": "Сохранить", "class_name": "Button", "control_type": "Button", "visible_only": True,
                     "enabled_only": True, "found_index": 0}, timeout=2).click()

            doc_already_exists = app.wait_element(
                {"title": "Подтвердить сохранение в виде", "class_name": "#32770", "control_type": "Window",
                 "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=2)

            if doc_already_exists:
                app.find_element(
                    {"title": "Да", "class_name": "CCPushButton", "control_type": "Button", "visible_only": True,
                     "enabled_only": True, "found_index": 0}, timeout=2).click()
                sleep(0.3)
                if app.wait_element(
                        {"title": "Да", "class_name": "CCPushButton", "control_type": "Button", "visible_only": True,
                         "enabled_only": True, "found_index": 0}, timeout=2):
                    app.find_element(
                        {"title": "Да", "class_name": "CCPushButton", "control_type": "Button", "visible_only": True,
                         "enabled_only": True, "found_index": 0}, timeout=2).click()
            break

        except:
            pass

    app.parent_back(1)


def performer(processing_date, processing_date_short, half_year_back_date):
    Session = sessionmaker()

    engine = create_engine(
        'postgresql+psycopg2://{username}:{password}@{host}:{port}/{base}'.format(**engine_kwargs),
        connect_args={'options': '-csearch_path=robot'}
    )
    Base.metadata.create_all(bind=engine)
    Session.configure(bind=engine)
    session = Session()

    get_all_payments_excel = True
    get_all_needed_payments = True
    check_payment_to_contragent_ = True
    check_payment_tmz_realization_ = True
    check_payment_factura_ = True
    check_payments_subconto_ = True
    check_fill_final_step_ = True

    first_excel_path = fr'\\172.16.8.87\d\.rpa\.agent\robot-posting-payments\Temp\lolus_{ip_address.replace(".", "_")}.xlsx'
    second_excel_path = fr'\\172.16.8.87\d\.rpa\.agent\robot-posting-payments\Temp\chpokus_{ip_address.replace(".", "_")}.xlsx'
    subconto_path = fr'\\172.16.8.87\d\.rpa\.agent\robot-posting-payments\Temp\subconto_{ip_address.replace(".", "_")}.xlsx'

    procter_path = r'\\172.16.8.87\d\.rpa\.agent\robot-posting-payments\Temp\проктер.xlsx'

    temp_path = fr'\\172.16.8.87\d\.rpa\.agent\robot-posting-payments\Temp\temp_file_{ip_address.replace(".", "_")}.xlsx'  # os.path.join(working_path)

    calendar = pd.read_excel(os.path.join(production_calendar_path, f'Производственный календарь {processing_date[-4:]}.xlsx'))

    cur_day_index = calendar[calendar['Day'] == processing_date_short]['Type'].index[0]
    cur_day_type = calendar[calendar['Day'] == processing_date_short]['Type'].iloc[0]

    if cur_day_type == 'Holiday':
        return 0

    weekends = []
    weekends_type = []

    for i in range(cur_day_index - 1, -1, -1):
        print(i, calendar['Day'].iloc[i][:6] + '20' + calendar['Day'].iloc[i][-2:])
        weekends.append(calendar['Day'].iloc[i][:6] + '20' + calendar['Day'].iloc[i][-2:])
        weekends_type.append(calendar['Type'].iloc[i])
        if calendar['Type'].iloc[i] == 'Working':
            processing_date = calendar['Day'].iloc[i]
            processing_date = f"{processing_date.split('.')[0]}.{processing_date.split('.')[1]}.20{processing_date.split('.')[2]}"

            break

    print(processing_date)

    # half_year_back_date = '17.08.2023'
    print(ip_address)

    # if ip_address == main_executor:
    if True:
        if get_all_payments_excel:

            app = Odines()
            # app.run()
            app.auth()

            app.open('Банк и касса', 'Платежное поручение входящее')

            app.open('Файл', 'Открыть...')

            app1 = App('')
            app1.wait_element({"title": "Открытие", "class_name": "#32770", "control_type": "Window",
                               "visible_only": True, "enabled_only": True, "found_index": 0})

            app1.find_element({"title": "Имя файла:", "class_name": "Edit", "control_type": "Edit",
                               "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app1.find_element({"title": "Имя файла:", "class_name": "Edit", "control_type": "Edit",
                               "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(form_document_path, app.keys.ENTER)

            # if app.wait_element({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
            #                      "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1):
            #     app.find_element({"title": "Да", "class_name": "", "control_type": "Button",
            #                       "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.parent_switch({"title": "", "class_name": "", "control_type": "Pane",
                               "visible_only": True, "enabled_only": True, "found_index": 34})

            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}, timeout=3).click()
            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}, timeout=3).type_keys("{BACKSPACE}" * 15)
            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}, timeout=3).type_keys(processing_date)

            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 1, "parent": app.root}, timeout=3).click()
            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}, timeout=3).type_keys("{BACKSPACE}" * 15)
            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 1, "parent": app.root}, timeout=3).type_keys(processing_date)

            app.find_element({"title": "", "class_name": "", "control_type": "ComboBox",
                              "visible_only": True, "enabled_only": True, "found_index": 1, "parent": app.root}, timeout=3).click()
            app.find_element({"title": "", "class_name": "", "control_type": "ComboBox",
                              "visible_only": True, "enabled_only": True, "found_index": 1, "parent": app.root}, timeout=3).type_keys(app.keys.CONTROL, app.keys.DOWN)

            found = False
            previous_list = None

            while not found:

                current_list = app.find_elements(
                    selector={
                        "control_type": "ListItem",
                        "visible_only": True,
                        "enabled_only": True,
                        "parent": app.root
                    },
                    timeout=3,
                )

                current_list_texts = [
                    item.element.element_info.rich_text for item in current_list
                ]

                for item, text in zip(current_list, current_list_texts):
                    # print(item, text, sep=' | ')
                    if text.replace(' ', '') == 'ПлатежноеПоручениеВходящее':
                        item.click()
                        found = True
                        break

                if previous_list is not None:
                    # Проверка, изменился ли список после прокрутки
                    if current_list_texts == previous_list:
                        break

                for _ in range(10):
                    pag.hotkey('down')
                previous_list = current_list_texts

            app.find_element({"title": "Выполнить", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": None}, timeout=10).click(double=True)

            app.find_element({"title": "", "class_name": "", "control_type": "DataGrid",
                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}, timeout=3).click()

            save_the_report(app, first_excel_path)

            app.open("Окна", "Закрыть все")
            # app.quit()

        if get_all_needed_payments:

            # app = Odines()
            # app.auth()

            app.parent_switch(app.root)
            app.open('Банк и касса', 'Платежное поручение входящее')
            app.find_element({"title": "Установить интервал дат...", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.parent_switch({"title": "Настройка периода", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                               "visible_only": True, "enabled_only": True, "found_index": 0})

            app.find_element({"title": "Период", "class_name": "", "control_type": "TabItem",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "День", "class_name": "", "control_type": "RadioButton",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(processing_date)

            app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            # app.parent_switch(app.root)
            app.parent_back(1)

            app.find_element({"title": "Установить отбор и сортировку списка...", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.parent_switch({"title": "Отбор и сортировка", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                               "visible_only": True, "enabled_only": True, "found_index": 0})

            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 7}).click()
            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 7}).type_keys('KZT')

            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 9}).click()
            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 9}).type_keys('Оплата от покупателя')

            app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.parent_back(1)

            app.find_element({"title_re": ".* Номер", "class_name": "", "control_type": "Custom",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click(right=True)

            app.find_element({"title": "Вывести список...", "class_name": "", "control_type": "MenuItem",
                              "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=360).click()

            app.parent_switch({"title": "Вывести список", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                               "visible_only": True, "enabled_only": True, "found_index": 0})

            app.find_element({"title": "Выключить все", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "Включить все", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "ОК", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            save_the_report(app, second_excel_path)

            app.quit()

            fix_excel_file_error(Path(first_excel_path))

            df = pd.read_excel(first_excel_path, header=3)

            # print(df[df['ПометкаУдаления'] == 'Нет']['Номер'])

            fix_excel_file_error(Path(second_excel_path))

            main_df = pd.read_excel(second_excel_path)

            ids = []
            summs = []
            contragents = []
            branches = []

            for ind in range(len(main_df)):
                if len(df[(df['ПометкаУдаления'] == 'Нет') & (df['Номер'] == main_df['Номер'].iloc[ind])]) != 0:
                    ids.append(main_df['Номер'].iloc[ind])
                    summs.append(main_df['Сумма документа'].iloc[ind])
                    contragents.append(main_df['Контрагент'].iloc[ind])
                    branches.append(main_df['Филиал'].iloc[ind])

                    add_to_db(session, 'new', main_df['Дата'].iloc[ind], str(main_df['Номер'].iloc[ind]), main_df['Сумма документа'].iloc[ind], main_df['Контрагент'].iloc[ind],
                              main_df['Филиал'].iloc[ind], None, None, None, None, None)

    #
    #     rows = get_all_data_by_status(session, ['new'])
    #
    #     while len(rows) == 0:
    #
    #         rows = get_all_data_by_status(session, ['new'])
    #
    #         sleep(10)

    rows = get_all_data_by_status(session, ['new'])
    #
    # if len(rows) == 0:
    #     df = pd.read_excel(first_excel_path, header=3)
    #     main_df = pd.read_excel(second_excel_path)
    #
    #     ids = []
    #     summs = []
    #     contragents = []
    #     branches = []
    #
    #     for ind in range(len(main_df)):
    #         if len(df[(df['ПометкаУдаления'] == 'Нет') & (df['Номер'] == main_df['Номер'].iloc[ind])]) != 0:
    #             ids.append(main_df['Номер'].iloc[ind])
    #             summs.append(main_df['Сумма документа'].iloc[ind])
    #             contragents.append(main_df['Контрагент'].iloc[ind])
    #             branches.append(main_df['Филиал'].iloc[ind])
    #
    #             # add_to_db(session, 'new', main_df['Дата'].iloc[ind], str(main_df['Номер'].iloc[ind]), main_df['Сумма документа'].iloc[ind], main_df['Контрагент'].iloc[ind],
    #             #           main_df['Филиал'].iloc[ind], None, None, None, None, None)
    #             add_to_db(session, 'new', main_df['Дата'].iloc[ind], str(main_df['Номер'].iloc[ind]), main_df['Сумма документа'].iloc[ind], main_df['Контрагент'].iloc[ind],
    #                       None, None, None, None, None, None)

    rows = get_all_data_by_status(session, ['new'])
    ind = -1

    while len(rows) != 0:

        rows = get_all_data_by_status(session, ['new'])

        print(f'Total rows: {len(rows)}')

        sleep(1)

        # for ind, row in enumerate(rows):

        row = rows[0]
        ind += 1

        logger.info(row.contragent)

        if 'проктер' in str(row.contragent).lower():
            continue

        # if row.contragent != 'РХМ Казахстан ТОО':
        #     continue

        # if '(' not in row.contragent:
        #     continue
        # continue

        # if True:
        try:

            logger.warning(f'Started {row.payment_id} | {row.contragent}')

            check_payment_to_contragent = check_payment_to_contragent_
            check_payment_tmz_realization = check_payment_tmz_realization_
            check_payment_factura = check_payment_factura_
            check_payments_subconto = check_payments_subconto_
            check_fill_final_step = check_fill_final_step_

            app = Odines()
            app.auth()

            if row.contragent not in ['ПРОКТЕР ЭНД ГЭМБЛ КАЗАХСТАН ДИСТРИБЬЮШН ТОО (13623)', 'КИМБЕРЛИ-КЛАРК КАЗАХСТАН ТОО (3199)']:

                # ONE -------------------------------------------------------------------------------------------------------------------------------------------------

                logger.warning(f'STATUS FOR check_payment_to_contragent: {row.invoice_payment_to_contragent}')

                if check_payment_to_contragent:

                    logger.warning(f'Started check_payment_to_contragent {ind} | {len(rows)}')

                    try:

                        logger.warning(f'Checkpoint1 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.open('Продажа', 'Счет на оплату покупателю')

                        app.find_element({"title": "Установить интервал дат...", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.parent_switch({"title": "Настройка периода", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0})

                        app.find_element({"title": "Период", "class_name": "", "control_type": "TabItem",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "Произвольный интервал", "class_name": "", "control_type": "RadioButton",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(app.keys.BACKSPACE * 15, half_year_back_date)

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 1}).click()

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys(app.keys.BACKSPACE * 15, processing_date)

                        app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint2 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.parent_back(1)

                        app.find_element({"title": "Установить отбор и сортировку списка...", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint3 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.parent_switch({"title": "Отбор и сортировка", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0}, maximize=True)

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 13}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 13}).type_keys(row.contragent, protect_first=True)

                        app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint4 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        if app.wait_element({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                             "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1):
                            app.parent_switch({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                               "visible_only": True, "enabled_only": True, "found_index": 0})
                            app.find_element({"title": "Нет", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                            update_in_db(session, row, 'processing', None, None,
                                         False, None, None, None)

                            # continue

                        logger.warning(f'Checkpoint5 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.parent_back(1)

                        app.find_element({"title": "Действия", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "Вывести список...", "class_name": "", "control_type": "MenuItem",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=360).click()

                        logger.warning(f'Checkpoint6 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.parent_switch({"title": "Вывести список", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0})

                        logger.warning(f'Checkpoint7 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Включить все", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "ОК", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint8 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        save_the_report(app, temp_path)

                        logger.warning(f'Checkpoint9 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        fix_excel_file_error(temp_path)

                        df_ = pd.read_excel(temp_path)

                        found_ = False
                        branch = None

                        for i in range(len(df_) - 1, -1, -1):

                            if float(df_['Сумма'].iloc[i]) == float(row.payment_sum):
                                found_ = True
                                branch = df_['Организация'].iloc[i]

                                check_payment_tmz_realization = False
                                check_payment_factura = False
                                check_payments_subconto = False

                                break

                        if found_:
                            update_in_db(session, row, 'processing', branch, None,
                                         True, False, False, False)
                            # check_payment_tmz_realization_ = False
                            # check_payment_factura_ = False
                            # check_payments_subconto_ = False

                        else:
                            update_in_db(session, row, 'processing', None, None,
                                         False, None, None, None)

                        print()

                        logger.warning(f'Checkpoint10 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        # app.quit()

                        app.open("Окна", "Закрыть все")

                        # app.parent = app.root

                        logger.warning(f'Checkpoint11 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                    except Exception as err1:
                        # Add Error reason
                        traceback.print_exc()
                        app.open("Окна", "Закрыть все")
                        logger.warning(f'Error1 occured: {err1}')
                        update_in_db(session, row, 'processing', None, None,
                                     False, None, None, None, error_reason_=str(traceback.format_exc())[:500])

                # TWO -------------------------------------------------------------------------------------------------------------------------------------------------

                logger.warning(f'STATUS FOR tmz_realization: {row.tmz_realization}')

                if check_payment_tmz_realization and row.tmz_realization is None:

                    logger.warning(f'Started check_payment_tmz_realization {ind} | {len(rows)}')

                    try:

                        app.parent_switch(app.root)

                        # app = Odines()
                        # app.auth()

                        logger.warning(f'Checkpoint0 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.open('Продажа', 'Реализация ТМЗ и услуг')

                        logger.warning(f'Checkpoint0.0 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        # app.parent_switch({"title": "", "class_name": "", "control_type": "Pane",
                        #                    "visible_only": True, "enabled_only": True, "found_index": 29, "parent": app.root}, timeout=1000)

                        logger.warning(f'Checkpoint0.1 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Установить интервал дат...", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=360).click()

                        logger.warning(f'Checkpoint1 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.parent_switch({"title": "Настройка периода", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0})

                        logger.warning(f'Checkpoint2 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Период", "class_name": "", "control_type": "TabItem",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "Произвольный интервал", "class_name": "", "control_type": "RadioButton",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(app.keys.BACKSPACE * 15, half_year_back_date)

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 1}).click()

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys(app.keys.BACKSPACE * 15, processing_date)

                        app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint3 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.parent_back(1)

                        logger.warning(f'Checkpoint4 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Установить отбор и сортировку списка...", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint5 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.parent_switch({"title": "Отбор и сортировка", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0}, maximize=True)

                        logger.warning(f'Checkpoint6 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.filter({'Контрагент': ('Равно', row.contragent)})

                        # app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                        #                   "visible_only": True, "enabled_only": True, "found_index": 21}).click()
                        # app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                        #                   "visible_only": True, "enabled_only": True, "found_index": 21}).type_keys(row.contragent, protect_first=True)

                        # app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                        #                   "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint7 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        if app.wait_element({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                             "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1):
                            app.parent_switch({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                               "visible_only": True, "enabled_only": True, "found_index": 0})
                            app.find_element({"title": "Нет", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                            update_in_db(session, row, 'processing', None, None,
                                         None, False, None, None)

                            # continue

                        app.parent_back(1)

                        logger.warning(f'Checkpoint8 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Действия", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint8.0 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Вывести список...", "class_name": "", "control_type": "MenuItem",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=360).click()

                        logger.warning(f'Checkpoint9 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.parent_switch({"title": "Вывести список", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0})

                        logger.warning(f'Checkpoint10 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Включить все", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "ОК", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint11 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        save_the_report(app, temp_path)

                        logger.warning(f'Checkpoint12 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        fix_excel_file_error(temp_path)

                        df_ = pd.read_excel(temp_path)

                        found_ = False
                        branch = None

                        for i in range(len(df_) - 1, -1, -1):

                            if float(df_['Сумма'].iloc[i]) == float(row.payment_sum):
                                found_ = True
                                branch = df_['Организация'].iloc[i]

                                check_payment_factura = False
                                check_payments_subconto = False
                                break

                        if found_:
                            logger.info(f'BRANCHOOSS: {branch}')
                            update_in_db(session, row, 'processing', branch, None,
                                         None, True, False, False)
                        else:
                            update_in_db(session, row, 'processing', None, None,
                                         None, False, None, None)

                        print()

                        logger.warning(f'Checkpoint13 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        # app.quit()

                        app.open("Окна", "Закрыть все")

                        # app.parent = app.root

                        logger.warning(f'Checkpoint14 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                    except Exception as err2:
                        traceback.print_exc()
                        app.open("Окна", "Закрыть все")
                        logger.warning(f'Error2 occured: {err2}')
                        update_in_db(session, row, 'processing', None, None,
                                     None, False, None, None, error_reason_=str(traceback.format_exc())[:500])

                # THREE -----------------------------------------------------------------------------------------------------------------------------------------------

                logger.warning(f'STATUS FOR invoice_factura: {row.invoice_factura}')
                if check_payment_factura and row.invoice_factura is None:

                    # logger.info(f'----- Started check_payment_factura -----')
                    # logger.warning(f'----- Started check_payment_factura -----')
                    #
                    # for ind, row in enumerate(rows):

                    logger.warning(f'Started check_payment_factura {ind} | {len(rows)}')

                    try:

                        app.parent_switch(app.root)

                        # app = Odines()
                        # app.auth()

                        logger.warning(f'Checkpoint0 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.open('Продажа', 'Счета-фактуры выданные', 'Счет-фактура выданный')

                        logger.warning(f'Checkpoint0.0 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        # app.parent_switch({"title": "", "class_name": "", "control_type": "Pane",
                        #                    "visible_only": True, "enabled_only": True, "found_index": 29}, timeout=360)

                        logger.warning(f'Checkpoint1 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Установить интервал дат...", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=60).click()

                        logger.warning(f'Checkpoint1.1 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.parent_switch({"title": "Настройка периода", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0})

                        logger.warning(f'Checkpoint2 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Период", "class_name": "", "control_type": "TabItem",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "Произвольный интервал", "class_name": "", "control_type": "RadioButton",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(app.keys.BACKSPACE * 15, half_year_back_date)

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 1}).click()

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys(app.keys.BACKSPACE * 15, processing_date)

                        app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint3 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.parent_back(1)

                        logger.warning(f'Checkpoint4 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Установить отбор и сортировку списка...", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint5 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.parent_switch({"title": "Отбор и сортировка", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0}, maximize=True)

                        logger.warning(f'Checkpoint6 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 9}).click(double=True)

                        # for tab in range(15):
                        #     # with suppress(Exception):
                        #     #     app.find_element({"title": "Контрагент", "class_name": "", "control_type": "CheckBox",
                        #     #                               "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=0.1).click()
                        #     #     break
                        #     pag.hotkey('TAB')

                        logger.warning(f'Checkpoint7 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.filter({'Контрагент': ('Равно', row.contragent)})
                        # 1
                        # app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                        #                   "visible_only": True, "enabled_only": True, "found_index": 19}).click()
                        # app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                        #                   "visible_only": True, "enabled_only": True, "found_index": 19}).type_keys(row.contragent, protect_first=True)

                        # app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                        #                   "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint8 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        if app.wait_element({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                             "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1):
                            app.parent_switch({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                               "visible_only": True, "enabled_only": True, "found_index": 0})
                            app.find_element({"title": "Нет", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                            update_in_db(session, row, 'processing', None, None,
                                         None, None, False, None)

                            # continue

                        app.parent_back(1)

                        logger.warning(f'Checkpoint9 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Действия", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "Вывести список...", "class_name": "", "control_type": "MenuItem",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=360).click()

                        logger.warning(f'Checkpoint10 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.parent_switch({"title": "Вывести список", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0})

                        logger.warning(f'Checkpoint11 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Включить все", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "ОК", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint12 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        save_the_report(app, temp_path)

                        logger.warning(f'Checkpoint13 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        fix_excel_file_error(temp_path)

                        df_ = pd.read_excel(temp_path)

                        found_ = False
                        branch = None

                        for i in range(len(df_) - 1, -1, -1):

                            if float(df_['Сумма документа'].iloc[i]) == float(row.payment_sum):
                                found_ = True
                                branch = df_['Организация'].iloc[i]

                                check_payments_subconto = False
                                break

                        if found_:
                            update_in_db(session, row, 'processing', branch, None,
                                         None, None, True, False)
                        else:
                            update_in_db(session, row, 'processing', None, None,
                                         None, None, False, None)

                        print()

                        logger.warning(f'Checkpoint14 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        # app.quit()

                        app.open("Окна", "Закрыть все")

                        # app.parent = app.root

                        logger.warning(f'Checkpoint15 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                    except Exception as err3:
                        traceback.print_exc()
                        app.open("Окна", "Закрыть все")
                        logger.warning(f'Error3 occured: {err3}')
                        update_in_db(session, row, 'processing', None, None,
                                     None, None, False, None, error_reason_=str(traceback.format_exc())[:500])

                # SUBCONTO --------------------------------------------------------------------------------------------------------------------------------------------

                logger.warning(f'STATUS FOR subconto: {row.subconto}')

                if check_payments_subconto and row.subconto is None:

                    # logger.info(f'----- Started check_payments_subconto -----')
                    # logger.warning(f'----- Started check_payments_subconto -----')
                    #
                    # for ind, row in range(len(rows)):

                    logger.warning(f'Started check_payments_subconto {ind} | {len(rows)}')

                    try:

                        app.parent_switch(app.root)

                        # app = Odines()
                        # app.auth()

                        logger.warning(f'Checkpoint0 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.open('Отчеты', 'Анализ субконто')

                        logger.warning(f'Checkpoint0.0 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        # app.parent_switch({"title": "", "class_name": "", "control_type": "Pane",
                        #                    "visible_only": True, "enabled_only": True, "found_index": 29}, timeout=360)

                        logger.warning(f'Checkpoint1 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 1}, timeout=1000).click()

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys(app.keys.BACKSPACE * 15, half_year_back_date)

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 2}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 2}).type_keys(app.keys.BACKSPACE * 15, processing_date)

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(app.keys.F4)

                        logger.warning(f'Checkpoint2 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.parent_switch({"title": "Структурные единицы", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0})

                        logger.warning(f'Checkpoint3 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Установить флаги", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click(double=True)

                        app.find_element({"title": "ОК", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint4 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.parent_back(1)

                        logger.warning(f'Checkpoint5 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Субконто", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        for _ in range(10):

                            if app.wait_element({"title": "Удалить", "class_name": "", "control_type": "Button",
                                                 "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1):
                                app.find_element({"title": "Удалить", "class_name": "", "control_type": "Button",
                                                  "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                            if app.wait_element({"title": "Удалить", "class_name": "", "control_type": "Button",
                                                 "visible_only": True, "enabled_only": True, "found_index": 1}, timeout=1):
                                app.find_element({"title": "Удалить", "class_name": "", "control_type": "Button",
                                                  "visible_only": True, "enabled_only": True, "found_index": 1}).click()

                        logger.warning(f'Checkpoint6 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Добавить", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint7 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Вид субконто", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()
                        app.find_element({"title": "Вид субконто", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('Договоры', app.keys.ENTER)

                        logger.warning(f'Checkpoint8 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Добавить", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 1}).click()

                        app.find_element({"title": "Вид субконто", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()
                        app.find_element({"title": "Вид субконто", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('Контрагенты', app.keys.ENTER)

                        logger.warning(f'Checkpoint9 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": " Значение", "class_name": "", "control_type": "Custom",
                                          "visible_only": True, "enabled_only": True, "found_index": 1}).click(double=True)
                        app.find_element({"title": " Значение", "class_name": "", "control_type": "Custom",
                                          "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys(row.contragent, app.keys.ENTER, protect_first=True)

                        if app.wait_element({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                             "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1):
                            app.parent_switch({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                               "visible_only": True, "enabled_only": True, "found_index": 0})
                            app.find_element({"title": "Нет", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                            update_in_db(session, row, 'processing', None, None,
                                         None, None, None, False)

                            # continue

                        logger.warning(f'Checkpoint10 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.find_element({"title": "Сформировать", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        logger.warning(f'Checkpoint11 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        save_the_report(app, subconto_path)

                        logger.warning(f'Checkpoint12 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        # app.quit()

                        app.open("Окна", "Закрыть все")

                        # app.parent = app.root

                        logger.warning(f'Checkpoint13 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        # ----- Subconto analysis ------

                        fix_excel_file_error(Path(subconto_path))

                        main_df = pd.read_excel(subconto_path, header=8)

                        main_df = main_df.drop(['Unnamed: 0'], axis=1)

                        main_df.columns = ['Subconto', 'Debit', 'Credit', 'Debit.1', 'Credit.1', 'Debit.2', 'Credit.2']

                        # contragent = 'ТОО Kobako Tsuvari'
                        # summ = 9173

                        contragent = row.contragent
                        summ = row.payment_sum
                        branch = None
                        contract = None

                        # print(main_df)

                        all_contracts = main_df[main_df['Subconto'] == contragent]

                        # print(all_contracts.index)

                        if len(all_contracts) == 0:
                            pass

                        if len(all_contracts) == 1:
                            ind = all_contracts.index[0]

                            print(main_df['Subconto'].iloc[ind - 2], main_df['Debit.2'].iloc[ind - 2])
                            branch = main_df['Subconto'].iloc[ind - 2]
                            contract = main_df['Subconto'].iloc[ind - 1]

                        if len(all_contracts) > 1:

                            for i in all_contracts.index:
                                print(main_df['Subconto'].iloc[i - 2], main_df['Debit.2'].iloc[i - 2])
                                if float(summ) == float(main_df['Debit.2'].iloc[i - 2]):
                                    branch = main_df['Subconto'].iloc[i - 2]
                                    contract = main_df['Subconto'].iloc[i - 1]

                        if branch is not None:
                            update_in_db(session, row, 'processing', branch, None,
                                         None, None, None, True)
                        else:
                            update_in_db(session, row, 'processing', branch, None,
                                         None, None, None, False)

                        print('-----------------------')
                        # print(branch)
                        # print(contract)

                        logger.warning(f'Checkpoint14 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                    except Exception as err4:
                        traceback.print_exc()
                        app.open("Окна", "Закрыть все")
                        logger.warning(f'Error4 occured: {err4}')
                        update_in_db(session, row, 'processing', None, None,
                                     None, None, None, False, error_reason_=str(traceback.format_exc())[:500])

                # FINAL -----------------------------------------------------------------------------------------------------------------------------------------------

                if check_fill_final_step:

                    # for ind, row in enumerate(rows):

                    app.parent_switch(app.root)

                    # app = Odines()
                    # app.auth()

                    app.open('Банк и касса', 'Платежное поручение входящее')
                    app.find_element({"title": "Установить интервал дат...", "class_name": "", "control_type": "Button",
                                      "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                    app.parent_switch({"title": "Настройка периода", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                       "visible_only": True, "enabled_only": True, "found_index": 0})

                    app.find_element({"title": "Период", "class_name": "", "control_type": "TabItem",
                                      "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                    app.find_element({"title": "День", "class_name": "", "control_type": "RadioButton",
                                      "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(row.payment_date.strftime('%d.%m.%Y'))

                    app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                                      "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                    # app.parent_switch(app.root)
                    app.parent_back(1)

                    # app.find_element({"title": "Установить отбор и сортировку списка...", "class_name": "", "control_type": "Button",
                    #                   "visible_only": True, "enabled_only": True, "found_index": 0}).click()
                    #
                    # app.parent_switch({"title": "Отбор и сортировка", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                    #                    "visible_only": True, "enabled_only": True, "found_index": 0})
                    #
                    # app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                    #                   "visible_only": True, "enabled_only": True, "found_index": 3}).click()
                    # app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                    #                   "visible_only": True, "enabled_only": True, "found_index": 3}).type_keys(row.payment_id)

                    # app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                    #                   "visible_only": True, "enabled_only": True, "found_index": 7}).click()
                    # app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                    #                   "visible_only": True, "enabled_only": True, "found_index": 7}).type_keys('KZT')
                    #
                    # app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                    #                   "visible_only": True, "enabled_only": True, "found_index": 9}).click()
                    # app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                    #                   "visible_only": True, "enabled_only": True, "found_index": 9}).type_keys('Оплата от покупателя')

                    app.filter({'Номер': ('Равно', row.payment_id), 'Валюта документа': ('Равно', 'KZT'), 'Вид операции': ('Равно', 'Оплата от покупателя')})

                    # app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                    #                   "visible_only": True, "enabled_only": True, "found_index": 0}).click()
                    print('finished filter')
                    app.parent_back(1)
                    print('searching for the row')
                    app.find_element({"title_re": ".* Номер", "class_name": "", "control_type": "Custom",
                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                    app.find_element({"title_re": ".* Номер", "class_name": "", "control_type": "Custom",
                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(app.keys.ENTER)

                    app.parent_switch({"title": "", "class_name": "", "control_type": "Pane",
                                       "visible_only": True, "enabled_only": True, "found_index": 34})

                    comment = str(app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                                    "visible_only": True, "enabled_only": True, "found_index": 2, "parent": app.root}).element.iface_value.CurrentValue)

                    print(comment)
                    print('---')
                    print('!')
                    invoice = None
                    branch = None
                    dds_statement = None
                    pnl = None
                    bill_calculations = None
                    bill_advances = None
                    skip = False

                    for _ in range(1):  # To avoid value conflict when there are 2 statements are True

                        if 'товар' in comment.lower() or 'продукты' in comment.lower():
                            print('CHECKPOINT FINAL 1')
                            invoice = 'договор реализации'
                            branch = row.branch
                            dds_statement = 'Авансы полученные за товар'
                            pnl = 'Поступления от реализации товаров'
                            bill_calculations = 1210
                            bill_advances = 3510
                            break

                        if 'арен' in comment.lower():
                            print('CHECKPOINT FINAL 2')
                            invoice = row.invoice_id
                            branch = row.branch
                            dds_statement = 'Авансы полученные за услуги'
                            pnl = 'Поступления арендных платежей'
                            bill_calculations = 1260
                            bill_advances = 3510
                            break

                        if 'обесп взнос' in comment.lower():
                            print('CHECKPOINT FINAL 3')
                            invoice = row.invoice_id
                            branch = row.branch
                            dds_statement = 'Авансы полученные за услуги'
                            pnl = 'Поступления арендных платежей'
                            bill_calculations = 1260
                            bill_advances = 4150
                            break

                        if 'ком услуг' in comment.lower() or 'электроэнерги' in comment.lower():
                            print('CHECKPOINT FINAL 4')
                            invoice = row.invoice_id
                            branch = row.branch
                            dds_statement = 'Крат. деб. задолж-ть покупат-ей за услуги в тенге'
                            pnl = 'Поступление за оказанные услуги'
                            bill_calculations = 1210
                            bill_advances = 3510
                            break

                        if 'маркетинг' in comment.lower() or 'проф услуги' in comment.lower():
                            print('CHECKPOINT FINAL 5')
                            invoice = row.invoice_id
                            branch = 'ТОО “Magnum Cash&Carry”'
                            dds_statement = 'Крат. деб. задолж-ть покупат-ей за услуги в тенге'
                            pnl = 'Поступления от маркетинговой деятельности'
                            bill_calculations = 1210
                            bill_advances = 3510
                            break

                        if any(item in comment.lower() for item in ['собственных средств на свой счет', 'возврат', 'возврат ошиб.платежей',
                                                                    'для зачисления на картсчет сотрудникам', 'предоставление займа']):
                            print('CHECKPOINT FINAL 6')
                            skip = True
                            break

                    if skip:
                        print('CHECKPOINT FINAL 7')
                        # continue
                        pass

                    sleep(1)

                    if branch is None:
                        app.find_element({"title": "Закрыть", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 3}).click()
                        update_in_db(session, row, 'finished', None, None,
                                     None, None, None, None)

                    if dds_statement is None:
                        invoice = ''
                        branch = row.branch

                    sleep(1)

                    if invoice is not None:
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 12, "parent": app.root}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 12, "parent": app.root}).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100, invoice, app.keys.TAB)
                    sleep(1)
                    if dds_statement is not None:
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 15, "parent": app.root}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 15, "parent": app.root}).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100, dds_statement, app.keys.TAB)

                    sleep(1)
                    if pnl is not None:
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 18, "parent": app.root}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 18, "parent": app.root}).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100, pnl, app.keys.TAB)
                        if app.wait_element({"class_name": "", "control_type": "ListItem",
                                             "visible_only": True, "enabled_only": True, "parent": app.root}, timeout=1):
                            els = app.find_elements({"class_name": "", "control_type": "ListItem", "visible_only": True, "enabled_only": True, "parent": app.root}, timeout=10)

                            print('LIST DROPPED DOWN')
                            sleep(1)
                            # for el in els:
                            #     sleep(1)
                            #     if len(el.element.element_info.rich_text) - len(pnl) <= 5:
                            #         print(len(el.element.element_info.rich_text), el.element.element_info.rich_text)
                            # print('123-12-10-0-0')
                            for el in els:
                                sleep(1)
                                if len(el.element.element_info.rich_text) - len(pnl) <= 5:
                                    print('Clicking!!!!')
                                    sleep(5)
                                    el.click()
                                    break
                                # app.find_elements({"class_name": "", "control_type": "ListItem", "visible_only": True, "enabled_only": True, "parent": app.root}, timeout=10)[0].element.element_info.rich_text

                            # app.find_element({"title": "", "class_name": "", "control_type": "ListItem",
                            #                   "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                    sleep(1)
                    if bill_calculations is not None:
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 16, "parent": app.root}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 16, "parent": app.root}).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100, bill_calculations)

                    sleep(1)
                    if bill_advances is not None:
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 17, "parent": app.root}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 17, "parent": app.root}).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100, bill_advances)

                    sleep(1)
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": 19, "parent": app.root}).click()
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": 19, "parent": app.root}).type_keys(app.keys.RIGHT * 70, app.keys.BACKSPACE * 130, branch)

                    update_in_db(session, row, 'success', None, None,
                                 None, None, None, None)
                    app.find_element({"title": "ОК", "class_name": "", "control_type": "Button",
                                      "visible_only": True, "enabled_only": True, "found_index": 0}).click()
                    print()

            else:

                # contragent = 'ПРОКТЕР ЭНД ГЭМБЛ КАЗАХСТАН ДИСТРИБЬЮШН ТОО (13623)'
                cur_date = processing_date
                # cur_date = '16.03.2023'
                search_date = datetime.datetime.strptime(cur_date, '%d.%m.%Y') - datetime.timedelta(days=1)
                summ = 13459033.0

                print(cur_date, search_date)

                if row.contragent == 'ПРОКТЕР ЭНД ГЭМБЛ КАЗАХСТАН ДИСТРИБЬЮШН ТОО (13623)':

                    print('P&G!!!')

                    df = pd.read_excel(procter_path)

                    # print(df[df['Payment Date'] == search_date.strftime('%Y-%m-%d')])

                    filtered_df = df[df['Payment Date'] == search_date.strftime('%Y-%m-%d')]

                    all_invoices = []

                    for ind_ in range(len(filtered_df)):

                        print('NEW LINE', ind_)
                        # print(filtered_df['Invoice Number'].iloc[i], filtered_df['Invoice Amount'].iloc[i])

                        app.open('Продажа', 'Счета-фактуры выданные', 'Счет-фактура выданный')

                        app.find_element({"title": "Установить интервал дат...", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.parent_switch({"title": "Настройка периода", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0})

                        app.find_element({"title": "Интервал", "class_name": "", "control_type": "TabItem",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        # app.find_element({"title": "Произвольный интервал", "class_name": "", "control_type": "RadioButton",
                        #                   "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "Без ограничения", "class_name": "", "control_type": "RadioButton",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.parent_back(1)

                        app.find_element({"title": "Установить отбор и сортировку списка...", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.parent_switch({"title": "Отбор и сортировка", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0}, maximize=True)

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 9}).click(double=True)

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 5}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 5}).type_keys(filtered_df['Invoice Number'].iloc[ind_])

                        for tab in range(15):
                            pag.hotkey('TAB')

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 19}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 19}).type_keys(row.contragent, protect_first=True)

                        app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        if app.wait_element({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                             "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1):
                            app.parent_switch({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                               "visible_only": True, "enabled_only": True, "found_index": 0})
                            app.find_element({"title": "Нет", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                            update_in_db(session, row, 'processing', None, None,
                                         None, None, None, None)

                            # continue

                        app.parent_back(1)

                        invoice = app.find_element({"title_re": ".* Договор контрагента", "class_name": "", "control_type": "Custom",
                                                    "visible_only": True, "enabled_only": True, "found_index": 0}).element.element_info.element.CurrentName

                        logger.warning(f'Found invoice: {invoice}')

                        all_invoices.append({invoice: row.payment_sum})

                        app.open("Окна", "Закрыть все")

                        app.parent_switch(app.root)
                        # app.quit()

                    print(all_invoices)
                    logger.warning(all_invoices)

                    print('LAST STEP')
                    if check_fill_final_step:

                        # app = Odines()
                        # app.auth()

                        app.open('Банк и касса', 'Платежное поручение входящее')
                        app.find_element({"title": "Установить интервал дат...", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.parent_switch({"title": "Настройка периода", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0})

                        app.find_element({"title": "Период", "class_name": "", "control_type": "TabItem",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "День", "class_name": "", "control_type": "RadioButton",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(row.payment_date.strftime('%d.%m.%Y'))

                        app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.parent_back(1)

                        app.find_element({"title": "Установить отбор и сортировку списка...", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.parent_switch({"title": "Отбор и сортировка", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0})

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 3}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 3}).type_keys(row.payment_id)

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 7}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 7}).type_keys('KZT')

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 9}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": 9}).type_keys('Оплата от покупателя')

                        app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        app.parent_back(1)

                        app.find_element({"title_re": ".* Номер", "class_name": "", "control_type": "Custom",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()
                        app.find_element({"title_re": ".* Номер", "class_name": "", "control_type": "Custom",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(app.keys.ENTER)

                        app.parent_switch({"title": "", "class_name": "", "control_type": "Pane",
                                           "visible_only": True, "enabled_only": True, "found_index": 34})

                        if len(all_invoices) > 1:
                            app.find_element({"title": "Список", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                        else:
                            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                              "visible_only": True, "enabled_only": True, "found_index": 12}).click()
                            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                              "visible_only": True, "enabled_only": True, "found_index": 12}).type_keys()
                        print()

                else:

                    if 2.9 <= (datetime.datetime.now() - row.date_created).total_seconds() / 86400 <= 3.1:
                        pass

        except Exception as row_error:
            traceback.print_exc()
            logger.warning(f'Error on the row occured: {row_error}')
            update_in_db(session, row, 'failed', None, None,
                         None, None, None, None, error_reason_=str(traceback.format_exc())[:500])


def main():
    if True:

        logger.info('Process started')
        logger.warning('Process started')

        for day in range(13, 14):
            processing_date_ = datetime.date.today().strftime('%d.%m.%Y')
            processing_date_short_ = datetime.date.today().strftime('%d.%m.%y')


            if day < 10:
                processing_date_ = f'0{day}.02.2024'
                processing_date_short_ = f'0{day}.02.24'
            else:
                processing_date_ = f'{day}.02.2024'
                processing_date_short_ = f'{day}.02.24'

            half_year_back_date = (datetime.date(int(processing_date_.split('.')[2]), int(processing_date_.split('.')[1]), int(processing_date_.split('.')[0])) - datetime.timedelta(days=180)).strftime('%d.%m.%Y')

            performer(processing_date_, processing_date_short_, half_year_back_date)

    # except Exception as error:
    #         logger.info(f'Error occured: {error}')
    #         logger.warning(f'Error occured: {error}')
    #         traceback.print_exc()


if __name__ == '__main__':

    try:

        main()

    except (Exception,):
        traceback.print_exc()
        kill_process_list(process_list_path)
        sys.exit(1)

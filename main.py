import datetime
import os
import sys
from contextlib import suppress
from pathlib import Path
from time import sleep

import pyautogui as pag
from pywinauto import keyboard
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from config import logger, process_list_path, production_calendar_path, form_document_path, engine_kwargs
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


def main(processing_date, processing_date_short):
    Session = sessionmaker()

    engine = create_engine(
        'postgresql+psycopg2://{username}:{password}@{host}:{port}/{base}'.format(**engine_kwargs),
        connect_args={'options': '-csearch_path=robot'}
    )
    Base.metadata.create_all(bind=engine)
    Session.configure(bind=engine)
    session = Session()

    get_all_payments_excel = False
    get_all_needed_payments = False
    check_payment_to_contragent = True
    check_payment_tmz_realization = True
    check_payment_factura = True
    check_payments_subconto = True

    first_excel_path = r'\\172.16.8.87\d\.rpa\.agent\robot-posting-payments\Temp\lolus.xlsx'
    second_excel_path = r'\\172.16.8.87\d\.rpa\.agent\robot-posting-payments\Temp\chpokus.xlsx'
    subconto_path = r'\\172.16.8.87\d\.rpa\.agent\robot-posting-payments\Temp\subconto.xlsx'

    temp_path = r'\\172.16.8.87\d\.rpa\.agent\robot-posting-payments\Temp\temp_file.xlsx'  # os.path.join(working_path)

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

    half_year_back_date = '23.07.2023'

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
                    "parent": None
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

        app.quit()

    if get_all_needed_payments:

        app = Odines()
        app.auth()

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

        print(app.root, app.parent)
        # app.parent_switch(app.root)
        app.parent_back(1)

        print('Here3')
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
        print('Here1')
        app.find_element({"title_re": ".* Номер", "class_name": "", "control_type": "Custom",
                          "visible_only": True, "enabled_only": True, "found_index": 0}).click(right=True)

        print('Here2')
        app.find_element({"title": "Вывести список...", "class_name": "", "control_type": "MenuItem",
                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

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

        print(df[df['ПометкаУдаления'] == 'Нет']['Номер'])

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

    rows = get_all_data_by_status(session, 'new')

    # ONE -------------------------------------------------------------------------------------------------------------------------------------------------

    logger.info(f'----- Started check_payment_to_contragent -----')
    logger.warning(f'----- Started check_payment_to_contragent -----')

    for ind, row in enumerate(rows):

        logger.info(f'Started check_payment_to_contragent {ind} | {len(rows)}')
        logger.warning(f'Started check_payment_to_contragent {ind} | {len(rows)}')

        if check_payment_to_contragent:

            app = Odines()
            app.auth()

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

            app.parent_back(1)
            print('kpp1')
            app.find_element({"title": "Установить отбор и сортировку списка...", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.parent_switch({"title": "Отбор и сортировка", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                               "visible_only": True, "enabled_only": True, "found_index": 0}, maximize=True)
            print('kpp2')

            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 13}).click()
            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 13}).type_keys(row.contragent)

            app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.parent_back(1)

            app.find_element({"title": "Действия", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.parent_switch({"title": "Вывести список", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                               "visible_only": True, "enabled_only": True, "found_index": 0})

            app.find_element({"title": "Вывести список...", "class_name": "", "control_type": "MenuItem",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "Включить все", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "ОК", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            save_the_report(app, temp_path)

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
                             True, None, None, None)
            else:
                update_in_db(session, row, 'processing', None, None,
                             False, None, None, None)

            # update_in_db(session: Session, row: Table_, status_: str, branch_: str or None, invoice_id_: str or None,
            # invoice_payment_: bool or None, tmz_realization_: bool or None, invoice_factura_: bool or None, subconto_: bool or None):

            print()

            app.quit()

    # TWO -------------------------------------------------------------------------------------------------------------------------------------------------

    logger.info(f'----- Started check_payment_tmz_realization -----')
    logger.warning(f'----- Started check_payment_tmz_realization -----')

    for ind, row in enumerate(rows):

        logger.info(f'Started check_payment_tmz_realization {ind} | {len(rows)}')
        logger.warning(f'Started check_payment_tmz_realization {ind} | {len(rows)}')

        if check_payment_tmz_realization:

            app = Odines()
            app.auth()

            app.open('Продажа', 'Реализация ТМЗ и услуг')

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

            app.parent_back(1)
            print('kpp1')
            app.find_element({"title": "Установить отбор и сортировку списка...", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.parent_switch({"title": "Отбор и сортировка", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                               "visible_only": True, "enabled_only": True, "found_index": 0}, maximize=True)
            print('kpp2')

            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 21}).click()
            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 21}).type_keys(row.contragent)

            app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.parent_back(1)

            app.find_element({"title": "Действия", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.parent_switch({"title": "Вывести список", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                               "visible_only": True, "enabled_only": True, "found_index": 0})

            app.find_element({"title": "Вывести список...", "class_name": "", "control_type": "MenuItem",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "Включить все", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "ОК", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            save_the_report(app, temp_path)

            fix_excel_file_error(temp_path)

            df_ = pd.read_excel(temp_path)

            found_ = False
            branch = None

            for i in range(len(df_) - 1, -1, -1):

                if float(df_['Сумма документа'].iloc[i]) == float(row.payment_sum):
                    found_ = True
                    branch = df_['Организация'].iloc[i]

                    check_payment_factura = False
                    check_payments_subconto = False
                    break

            if found_:
                update_in_db(session, row, 'processing', branch, None,
                             None, True, None, None)
            else:
                update_in_db(session, row, 'processing', None, None,
                             None, False, None, None)

            # update_in_db(session: Session, row: Table_, status_: str, branch_: str or None, invoice_id_: str or None,
            # invoice_payment_: bool or None, tmz_realization_: bool or None, invoice_factura_: bool or None, subconto_: bool or None):

            print()

            app.quit()

    # THREE -----------------------------------------------------------------------------------------------------------------------------------------------

    logger.info(f'----- Started check_payment_factura -----')
    logger.warning(f'----- Started check_payment_factura -----')

    for ind, row in enumerate(rows):

        logger.info(f'Started check_payment_factura {ind} | {len(rows)}')
        logger.warning(f'Started check_payment_factura {ind} | {len(rows)}')

        if check_payment_factura:

            app = Odines()
            app.auth()

            app.open('Продажа', 'Счета-фактуры выданные', 'Счет-фактура выданный')

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

            app.parent_back(1)
            print('kpp1')
            app.find_element({"title": "Установить отбор и сортировку списка...", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.parent_switch({"title": "Отбор и сортировка", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                               "visible_only": True, "enabled_only": True, "found_index": 0}, maximize=True)
            print('kpp2')

            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 9}).click(double=True)

            for tab in range(15):
                # with suppress(Exception):
                #     app.find_element({"title": "Контрагент", "class_name": "", "control_type": "CheckBox",
                #                               "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=0.1).click()
                #     break
                pag.hotkey('TAB')

            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 19}).click()
            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 19}).type_keys(row.contragent)

            app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.parent_back(1)

            app.find_element({"title": "Действия", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.parent_switch({"title": "Вывести список", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                               "visible_only": True, "enabled_only": True, "found_index": 0})

            app.find_element({"title": "Вывести список...", "class_name": "", "control_type": "MenuItem",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "Включить все", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "ОК", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            save_the_report(app, temp_path)

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
                             None, None, True, None)
            else:
                update_in_db(session, row, 'processing', None, None,
                             None, None, False, None)

            # update_in_db(session: Session, row: Table_, status_: str, branch_: str or None, invoice_id_: str or None,
            # invoice_payment_: bool or None, tmz_realization_: bool or None, invoice_factura_: bool or None, subconto_: bool or None):

            print()

            app.quit()

    # SUBCONTO --------------------------------------------------------------------------------------------------------------------------------------------

    logger.info(f'----- Started check_payments_subconto -----')
    logger.warning(f'----- Started check_payments_subconto -----')

    for ind, row in range(len(rows)):

        logger.info(f'Started check_payments_subconto {ind} | {len(rows)}')
        logger.warning(f'Started check_payments_subconto {ind} | {len(rows)}')

        if check_payments_subconto:

            app = Odines()
            app.auth()

            app.open('Отчеты', 'Анализ субконто')
            #
            # app.parent_switch({"title": "", "class_name": "", "control_type": "Pane",
            #                    "visible_only": True, "enabled_only": True, "found_index": 29}, maximize=True)
            #
            # app.parent_switch({"title": "", "class_name": "", "control_type": "Pane",
            #                    "visible_only": True, "enabled_only": True, "found_index": 28})

            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 1}, timeout=180).click()
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

            app.parent_switch({"title": "Структурные единицы", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                               "visible_only": True, "enabled_only": True, "found_index": 0})

            app.find_element({"title": "Установить флаги", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click(double=True)

            app.find_element({"title": "ОК", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.parent_back(1)

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

            app.find_element({"title": "Добавить", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            app.find_element({"title": "Вид субконто", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()
            app.find_element({"title": "Вид субконто", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('Договоры', app.keys.ENTER)

            app.find_element({"title": "Добавить", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 1}).click()

            app.find_element({"title": "Вид субконто", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()
            app.find_element({"title": "Вид субконто", "class_name": "", "control_type": "Edit",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('Контрагенты', app.keys.ENTER)

            app.find_element({"title": " Значение", "class_name": "", "control_type": "Custom",
                              "visible_only": True, "enabled_only": True, "found_index": 1}).click(double=True)
            app.find_element({"title": " Значение", "class_name": "", "control_type": "Custom",
                              "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys(row.contragent, app.keys.ENTER)

            print()

            app.find_element({"title": "Сформировать", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            save_the_report(app, subconto_path)

            app.quit()

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

            print(main_df)

            all_contracts = main_df[main_df['Subconto'] == contragent]

            print(all_contracts.index)

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
            print(branch)
            print(contract)

    # FINAL -----------------------------------------------------------------------------------------------------------------------------------------------


if __name__ == '__main__':

    if True:

        logger.info('Process started')
        logger.warning('Process started')

        for day in range(31, 32):
            processing_date_ = datetime.date.today().strftime('%d.%m.%Y')
            processing_date_short_ = datetime.date.today().strftime('%d.%m.%y')

            if day < 10:
                processing_date_ = f'0{day}.01.2024'
                processing_date_short_ = f'0{day}.01.24'
            else:
                processing_date_ = f'{day}.01.2024'
                processing_date_short_ = f'{day}.01.24'

            main(processing_date_, processing_date_short_)

    # except Exception as error:
    #     logger.info(f'Error occured: {error}')
    #     logger.warning(f'Error occured: {error}')

    # except (Exception,):
    #     kill_process_list(process_list_path)
    #     sys.exit(1)

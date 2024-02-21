import datetime
import os
import sys
import traceback
from contextlib import suppress
from pathlib import Path
from time import sleep

import psycopg2
import pyautogui as pag
from pywinauto import keyboard
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from config import logger, process_list_path, production_calendar_path, form_document_path, engine_kwargs, main_executor, ip_address, download_path
from models import Base, add_to_db, get_all_data, get_all_data_by_status, update_in_db
from tools.app import App
from tools.odines import Odines
from tools.process import kill_process_list

import pandas as pd

from tools.xlsx_fix import fix_excel_file_error


def save_the_report(app: Odines, savepath: str):
    with suppress(Exception):
        Path.unlink(Path(savepath))

    app.open('Файл', 'Сохранить как...')

    app.parent_switch({"title": "Сохранение", "class_name": "#32770", "control_type": "Window",
                       "visible_only": True, "enabled_only": True, "found_index": 0})

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
    app.find_element(
        {"title": "Имя файла:", "class_name": "Edit", "control_type": "Edit", "visible_only": True,
         "enabled_only": True, "found_index": 0}).click()
    app.find_element(
        {"title": "Имя файла:", "class_name": "Edit", "control_type": "Edit", "visible_only": True,
         "enabled_only": True, "found_index": 0}).type_keys(savepath)

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

    first_excel_path = os.path.join(download_path, f'lolus_{ip_address.replace(".", "_")}.xlsx')
    second_excel_path = os.path.join(download_path, f'chpokus_{ip_address.replace(".", "_")}.xlsx')
    subconto_path = os.path.join(download_path, f'subconto_{ip_address.replace(".", "_")}.xlsx')

    procter_path = r'\\vault.magnum.local\common\Stuff\_06_Бухгалтерия\Выписки\Для Робота\Для Проктер\PaymentReportDetails (13).xlsx'
    kimberly_path = r'\\vault.magnum.local\common\Stuff\_06_Бухгалтерия\Выписки\Для Робота\Для Кимберли-Кларк'

    temp_path = os.path.join(download_path, f'temp_file_{ip_address.replace(".", "_")}.xlsx')  # os.path.join(working_path)

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
    if get_all_payments_excel or get_all_needed_payments:

        app = Odines()
        app.auth()

        if get_all_payments_excel:

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
                              None, None, None, None, None, None)  # main_df['Филиал'].iloc[ind]

    #
    #     rows = get_all_data_by_status(session, ['new'])
    #
    #     while len(rows) == 0:
    #
    #         rows = get_all_data_by_status(session, ['new'])
    #
    #         sleep(10)

    # rows = get_all_data_by_status(session, ['new'])
    #
    # if len(rows) == 0:
    # if True:
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

    rows = get_all_data_by_status(session, ['new', 'processing'])
    ind = -1
    rows_ = None
    break_counter = 0

    while len(rows) != 0:

        rows = get_all_data_by_status(session, ['new', 'processing'])

        if rows_ == rows:
            break_counter += 1

        if break_counter >= 4:
            break

        rows_ = rows
        print(f'Total rows: {len(rows)}')

        sleep(1)

        # for ind, row in enumerate(rows):

        row = rows[0]
        ind += 1

        logger.info(row.contragent)

        # print('DDD:', datetime.datetime.now() - row.payment_date)

        # if 'проктер' in str(row.contragent).lower():  # or 'кимберли' in str(row.contragent).lower():
        #     continue

        # if str(row.contragent).strip() != 'Ахметов ИП':
        #     continue

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

                update_in_db(session, row, 'processing', None, None,
                             None, None, None, None)

                if_tovar = False

                if True:
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

                    app.filter({'Номер': ('Равно', row.payment_id), 'Валюта документа': ('Равно', 'KZT'), 'Вид операции': ('Равно', 'Оплата от покупателя')})

                    app.parent_back(1)

                    app.find_element({"title_re": ".* Номер", "class_name": "", "control_type": "Custom",
                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                    app.find_element({"title_re": ".* Номер", "class_name": "", "control_type": "Custom",
                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(app.keys.ENTER)

                    app.parent_switch({"title": "", "class_name": "", "control_type": "Pane",
                                       "visible_only": True, "enabled_only": True, "found_index": 34})

                    cancel_check = {"title": "Отмена проведения", "class_name": "", "control_type": "Button",
                                    "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}

                    if app.wait_element(cancel_check, timeout=5):
                        app.find_element(cancel_check).click()

                        app.find_element({"title": "Записать", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()

                    app.open("Окна", "Закрыть все")

                # ONE -------------------------------------------------------------------------------------------------------------------------------------------------

                logger.warning(f'STATUS FOR check_payment_to_contragent: {row.invoice_payment_to_contragent}')

                if check_payment_to_contragent:

                    logger.warning(f'Started check_payment_to_contragent {ind} | {len(rows)}')

                    try:

                        app.parent_switch(app.root)

                        logger.warning(f'Checkpoint1 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.open('Продажа', 'Счет на оплату покупателю', maximize_inner=True)

                        parent_ = app.parent

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

                        # app.find_element({"title": "Установить отбор и сортировку списка...", "class_name": "", "control_type": "Button",
                        #                   "visible_only": True, "enabled_only": True, "found_index": 0}).click()
                        #
                        # logger.warning(f'Checkpoint3 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')
                        #
                        # app.parent_switch({"title": "Отбор и сортировка", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                        #                    "visible_only": True, "enabled_only": True, "found_index": 0}, maximize=True)

                        app.filter({'Контрагент': ('Равно', row.contragent), 'Сумма документа': ('Равно', row.payment_sum)})

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

                        # app.parent_back(1)

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
                        index = None

                        for i in range(len(df_) - 1, -1, -1):

                            if float(df_['Сумма'].iloc[i]) == float(row.payment_sum):
                                found_ = True
                                branch = df_['Организация'].iloc[i]
                                index = i

                                # check_payment_tmz_realization = False
                                # check_payment_factura = False
                                # check_payments_subconto = False

                                break

                        if found_:

                            app.open("Окна", "Закрыть")

                            app.parent_switch(parent_, maximize=True)

                            if app.wait_element({"title": "Развернуть", "class_name": "", "control_type": "Button",
                                                 "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1):
                                app.find_element({"title": "Развернуть", "class_name": "", "control_type": "Button",
                                                  "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                            app.find_element({"title_re": ".* Контрагент", "class_name": "", "control_type": "Custom",
                                              "visible_only": True, "enabled_only": True, "found_index": index}).click(double=True)

                            invoice = str(app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                                            "visible_only": True, "enabled_only": True, "found_index": 3}).element.iface_value.CurrentValue)

                            nomenclatura = str(app.find_element({"title_re": ".* Номенклатура", "class_name": "", "control_type": "Custom",
                                                             "visible_only": True, "enabled_only": True, "found_index": 0}).element.element_info.name)

                            if 'товар' in nomenclatura.lower():
                                if_tovar = True

                            app.find_element({"title": "Закрыть", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 3}).click()

                            app.parent_back(1)

                            logger.warning(f'Found invoice: {len(invoice)} | {invoice}*')

                            update_in_db(session, row, 'processing', branch, invoice,
                                         True, None if invoice is None or invoice == '' else False, None if invoice is None or invoice == '' else False, None if invoice is None or invoice == '' else False)
                            # check_payment_tmz_realization_ = False
                            # check_payment_factura_ = False
                            # check_payments_subconto_ = False

                        else:
                            update_in_db(session, row, 'processing', None, None,
                                         False, None, None, None)

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

                logger.warning(f'STATUS FOR tmz_realization: {row.tmz_realization} | {check_payment_tmz_realization}')

                if check_payment_tmz_realization and row.tmz_realization is None:

                    logger.warning(f'Started check_payment_tmz_realization {ind} | {len(rows)}')

                    try:

                        app.parent_switch(app.root)

                        # app = Odines()
                        # app.auth()

                        logger.warning(f'Checkpoint0 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.open('Продажа', 'Реализация ТМЗ и услуг', maximize_inner=True)

                        parent_ = app.parent

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

                        app.filter({'Контрагент': ('Равно', row.contragent), 'Сумма документа': ('Равно', row.payment_sum)})

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
                        index = None

                        for i in range(len(df_) - 1, -1, -1):

                            if float(df_['Сумма'].iloc[i]) == float(row.payment_sum):
                                found_ = True
                                branch = df_['Организация'].iloc[i]
                                index = i

                                # check_payment_factura = False
                                # check_payments_subconto = False
                                break

                        if found_:

                            logger.info(f'BRANCHOOSS: {branch}')

                            app.open("Окна", "Закрыть")

                            app.parent_switch(parent_, maximize=True)

                            if app.wait_element({"title": "Развернуть", "class_name": "", "control_type": "Button",
                                                 "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1):
                                app.find_element({"title": "Развернуть", "class_name": "", "control_type": "Button",
                                                  "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                            app.find_element({"title_re": ".* Контрагент", "class_name": "", "control_type": "Custom",
                                              "visible_only": True, "enabled_only": True, "found_index": index}).click(double=True)

                            invoice = str(app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                                            "visible_only": True, "enabled_only": True, "found_index": 3}).element.iface_value.CurrentValue)

                            nomenclatura = str(app.find_element({"title_re": ".* Номенклатура", "class_name": "", "control_type": "Custom",
                                                                 "visible_only": True, "enabled_only": True, "found_index": 0}).element.element_info.name)

                            if 'товар' in nomenclatura.lower():
                                if_tovar = True
                            print(f"FOUND INVOICE: {invoice}")
                            app.find_element({"title": "Закрыть", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 3}).click()

                            app.parent_back(1)

                            update_in_db(session, row, 'processing', branch, invoice,
                                         None, True, None if invoice is None or invoice == '' else False, None if invoice is None or invoice == '' else False)
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

                        app.open('Продажа', 'Счета-фактуры выданные', 'Счет-фактура выданный', maximize_inner=True)

                        parent_ = app.parent

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

                        app.filter({'Контрагент': ('Равно', row.contragent), 'Сумма документа': ('Равно', row.payment_sum)})
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
                        index = None

                        for i in range(len(df_) - 1, -1, -1):

                            if float(df_['Сумма документа'].iloc[i]) == float(row.payment_sum):
                                found_ = True
                                branch = df_['Организация'].iloc[i]
                                index = i

                                # check_payments_subconto = False
                                break

                        if found_:

                            app.open("Окна", "Закрыть")

                            app.parent_switch(parent_, maximize=True)

                            if app.wait_element({"title": "Развернуть", "class_name": "", "control_type": "Button",
                                                 "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1):
                                app.find_element({"title": "Развернуть", "class_name": "", "control_type": "Button",
                                                  "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                            app.find_element({"title_re": ".* Контрагент", "class_name": "", "control_type": "Custom",
                                              "visible_only": True, "enabled_only": True, "found_index": index}).click(double=True)

                            invoice = str(app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                                            "visible_only": True, "enabled_only": True, "found_index": 3}).element.iface_value.CurrentValue)

                            nomenclatura = str(app.find_element({"title_re": ".* Номенклатура", "class_name": "", "control_type": "Custom",
                                                                 "visible_only": True, "enabled_only": True, "found_index": 0}).element.element_info.name)

                            if 'товар' in nomenclatura.lower():
                                if_tovar = True

                            app.find_element({"title": "Закрыть", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 3}).click()

                            app.parent_back(1)

                            update_in_db(session, row, 'processing', branch, invoice,
                                         None, None, True, None if invoice is None or invoice == '' else False)
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

                        app.open('Отчеты', 'Анализ субконто', maximize_inner=True)

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

                        if app.wait_element({"class_name": "", "control_type": "ListItem",
                                             "visible_only": True, "enabled_only": True, "parent": app.root}, timeout=1):
                            els = app.find_elements({"class_name": "", "control_type": "ListItem", "visible_only": True, "enabled_only": True, "parent": app.root}, timeout=10)
                            for el in els:
                                el.click()
                                break

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

                            conn = psycopg2.connect(dbname='adb', host='172.16.10.22', port='5432',
                                                    user='rpa_robot', password='Qaz123123+')

                            cur = conn.cursor()

                            cur.execute(f"""select distinct(name_sale_object_for_print) from dwh_data.dim_branches_src dbs""")
                            df = pd.DataFrame(cur.fetchall())
                            conn.close()

                            all_branches = []

                            for branchos in df[df.columns[0]]:
                                all_branches.append(branchos.replace(' ', '').lower())

                            all_branches.append('ТОО "Magnum Cash&Carry"'.replace(' ', '').lower())
                            row_ = main_df[main_df['Debit.2'] == summ]
                            print(row_)
                            print()
                            if len(row_) != 0:

                                if isinstance(row_['Subconto'].iloc[0], int):
                                    for ind_ in range(row_.index[0], -1, -1):
                                        print(main_df['Subconto'].iloc[ind_])
                                        if (contract is None and isinstance(main_df['Subconto'].iloc[ind_], str)
                                            and main_df['Subconto'].iloc[ind_] != contragent) \
                                                and main_df['Subconto'].iloc[ind_].replace(' ', '').lower() not in all_branches:
                                            contract = main_df['Subconto'].iloc[ind_]

                                        if isinstance(main_df['Subconto'].iloc[ind_], str) and main_df['Subconto'].iloc[ind_].replace(' ', '').lower() in all_branches:
                                            branch = main_df['Subconto'].iloc[ind_]
                                            break

                                elif isinstance(row_['Subconto'].iloc[0], str):
                                    for ind_ in range(row_.index[0] + 1, -1, -1):
                                        print(main_df['Subconto'].iloc[ind_])
                                        if (contract is None and isinstance(main_df['Subconto'].iloc[ind_], str)
                                            and main_df['Subconto'].iloc[ind_] != contragent) \
                                                and main_df['Subconto'].iloc[ind_].replace(' ', '').lower() not in all_branches:
                                            contract = main_df['Subconto'].iloc[ind_]

                                        if isinstance(main_df['Subconto'].iloc[ind_], str) and main_df['Subconto'].iloc[ind_].replace(' ', '').lower() in all_branches:
                                            branch = main_df['Subconto'].iloc[ind_]
                                            break

                            # for i in all_contracts.index:
                            #     print(main_df['Subconto'].iloc[i - 2], main_df['Debit.2'].iloc[i - 2])
                            #     if float(summ) == float(main_df['Debit.2'].iloc[i - 2]):
                            #         branch = main_df['Subconto'].iloc[i - 2]
                            #         contract = main_df['Subconto'].iloc[i - 1]

                        if branch is not None:
                            update_in_db(session, row, 'processing', branch, contract,
                                         None, None, None, True)
                        else:
                            update_in_db(session, row, 'processing', branch, contract,
                                         None, None, None, False)

                        print('-----------------------')
                        # print(branch)
                        # print(contract)

                        logger.warning(f'Checkpoint14 | {datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')

                        app.open("Окна", "Закрыть все")

                    except Exception as err4:
                        traceback.print_exc()
                        app.open("Окна", "Закрыть все")
                        logger.warning(f'Error4 occured: {err4}')
                        update_in_db(session, row, 'processing', None, None,
                                     None, None, None, False, error_reason_=str(traceback.format_exc())[:500])

                # FINAL -----------------------------------------------------------------------------------------------------------------------------------------------

                if check_fill_final_step:  # and any([row.invoice_payment_to_contragent, row.tmz_realization, row.invoice_factura, row.subconto]):

                    # for ind, row in enumerate(rows):

                    app.open("Окна", "Закрыть все")

                    app.parent_switch(app.root)

                    els = app.find_elements({"title": "Закрыть", "class_name": "", "control_type": "Button",
                                                    "visible_only": True, "enabled_only": True, "parent": app.root}, timeout=15)
                    for ids_ in range(len(els)):
                        try:
                            app.find_element({"title": "Закрыть", "class_name": "", "control_type": "Button",
                                                      "visible_only": True, "enabled_only": True, "found_index": 2, "parent": app.root}, timeout=15).click()
                        except:
                            break

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

                    app.filter({'Номер': ('Равно', row.payment_id), 'Валюта документа': ('Равно', 'KZT'), 'Вид операции': ('Равно', 'Оплата от покупателя')})

                    print('finished filter')
                    app.parent_back(1)
                    print('searching for the row')
                    app.find_element({"title_re": ".* Номер", "class_name": "", "control_type": "Custom",
                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                    app.find_element({"title_re": ".* Номер", "class_name": "", "control_type": "Custom",
                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(app.keys.ENTER)

                    app.parent_switch({"title": "", "class_name": "", "control_type": "Pane",
                                       "visible_only": True, "enabled_only": True, "found_index": 34})

                    cancel_check = {"title": "Отмена проведения", "class_name": "", "control_type": "Button",
                                    "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}

                    if app.wait_element(cancel_check, timeout=5):
                        app.find_element(cancel_check).click()

                    comment = str(app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                                    "visible_only": True, "enabled_only": True, "found_index": 2, "parent": app.root}).element.iface_value.CurrentValue)

                    print(comment)
                    print('---')
                    print('!')

                    if row.invoice_id == '':
                        row.invoice_id = None

                    check_for_index = (app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                                         "visible_only": True, "enabled_only": True, "found_index": 19, "parent": app.root}).element.element_info.element.CurrentHelpText)

                    invoice_index = 12 if check_for_index == 'Филиал' else 15
                    branch_index = 19 if check_for_index == 'Филиал' else 22
                    dds_index = 15 if check_for_index == 'Филиал' else 18
                    pnl_index = 18 if check_for_index == 'Филиал' else 21
                    bill_calculations_index = 16 if check_for_index == 'Филиал' else 19
                    bill_advances_index = 17 if check_for_index == 'Филиал' else 20

                    invoice = 'Без договора'
                    branch = row.branch
                    dds_statement = None
                    pnl = None
                    bill_calculations = None
                    bill_advances = None
                    skip = False

                    for _ in range(1):  # To avoid value conflict when there are 2 statements are True

                        if 'маркетинг' in comment.lower() or ('проф' in comment.lower() and 'услуг' in comment.lower()):
                            print('CHECKPOINT FINAL 5')
                            invoice = row.invoice_id
                            branch = 'ТОО “Magnum Cash&Carry”'
                            dds_statement = 'Крат.деб. задолж-ть покупат-ей за услуги в тенге'
                            pnl = 'Поступления от маркетинговой деятельности'
                            bill_calculations = 1210
                            bill_advances = 3510
                            break

                        if ('ком' in comment.lower() and 'услуг' in comment.lower()) or 'электроэнерг' in comment.lower() \
                                or 'комун' in comment.lower() or 'коммун' in comment.lower():
                            print('CHECKPOINT FINAL 4')
                            invoice = row.invoice_id
                            branch = row.branch
                            dds_statement = 'Авансы полученные за услуги'
                            pnl = 'Поступления арендных платежей'
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

                        if 'обесп' in comment.lower() and 'взнос' in comment.lower():
                            print('CHECKPOINT FINAL 3')
                            invoice = row.invoice_id
                            branch = row.branch
                            dds_statement = 'Авансы полученные за услуги'
                            pnl = 'Поступления арендных платежей'
                            bill_calculations = 1260
                            bill_advances = 4150
                            break

                        if 'товар' in comment.lower() or 'продукт' in comment.lower():
                            print('CHECKPOINT FINAL 1')
                            invoice = 'Договор реализации' if row.invoice_id is None else row.invoice_id
                            branch = row.branch
                            dds_statement = 'Авансы полученные за товар'
                            pnl = 'Поступления от реализации товаров'
                            bill_calculations = 1210
                            bill_advances = 3510
                            break

                        if 'сыр' in comment.lower():
                            print('CHECKPOINT FINAL 1.1')
                            invoice = row.invoice_id
                            branch = row.branch
                            dds_statement = 'Авансы полученные за товар'
                            pnl = 'Поступления от реализации сырья и материалов'
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

                        update_in_db(session, row, 'skipped.1', branch, invoice,
                                     None, None, None, None)

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": branch_index, "parent": app.root}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": branch_index, "parent": app.root}).type_keys(app.keys.RIGHT * 70, app.keys.BACKSPACE * 130)

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": invoice_index, "parent": app.root}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": invoice_index, "parent": app.root}).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100)

                        app.find_element({"title": "Записать", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()

                        app.find_element({"title": "Закрыть", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 4}).click()

                        update_in_db(session, row, 'skipped', branch, invoice,
                                     None, None, None, None)
                        continue

                    sleep(1)

                    if branch is None or invoice == 'Без договора':
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": branch_index, "parent": app.root}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": branch_index, "parent": app.root}).type_keys(app.keys.RIGHT * 70, app.keys.BACKSPACE * 130)

                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": invoice_index, "parent": app.root}).click()
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": invoice_index, "parent": app.root}).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100)

                        app.find_element({"title": "Записать", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()

                        app.find_element({"title": "Закрыть", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 4}).click()

                        update_in_db(session, row, 'finished del', branch, invoice,
                                     None, None, None, None)

                        app.quit()

                        continue

                    # if dds_statement is None:
                    #     invoice = None
                    #     branch = row.branch

                    if if_tovar:
                        invoice = 'Договор реализации'

                    sleep(1)
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": invoice_index, "parent": app.root}).click()
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": invoice_index, "parent": app.root}).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100)
                    if invoice is not None:
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": invoice_index, "parent": app.root}).type_keys(invoice, protect_first=True)
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": invoice_index, "parent": app.root}).type_keys(app.keys.TAB)

                    sleep(1)

                    # DDS ----------------------------------------------------------------------------------------------------------------------------------------------------
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": dds_index, "parent": app.root}).click()
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": dds_index, "parent": app.root}).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100)
                    if dds_statement is not None:
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": dds_index, "parent": app.root}).type_keys(dds_statement, protect_first=True)
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": dds_index, "parent": app.root}).type_keys(app.keys.TAB)

                    sleep(1)

                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": pnl_index, "parent": app.root}).click()
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": pnl_index, "parent": app.root}).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100)
                    if pnl is not None:
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": pnl_index, "parent": app.root}).type_keys(pnl)
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": pnl_index, "parent": app.root}).type_keys(app.keys.ENTER, app.keys.ENTER)
                        if app.wait_element({"class_name": "", "control_type": "ListItem",
                                             "visible_only": True, "enabled_only": True, "parent": app.root}, timeout=1):
                            els = app.find_elements({"class_name": "", "control_type": "ListItem", "visible_only": True, "enabled_only": True, "parent": app.root}, timeout=10)

                            print('LIST DROPPED DOWN')
                            sleep(1)

                            for el in els:
                                sleep(1)
                                if len(el.element.element_info.rich_text) - len(pnl) <= 5:
                                    print('Clicking!!!!')
                                    sleep(1)
                                    el.click()
                                    break
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": pnl_index, "parent": app.root}).type_keys(app.keys.ENTER)

                    sleep(1)

                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": bill_calculations_index, "parent": app.root}).click()
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": bill_calculations_index, "parent": app.root}).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100)
                    if bill_calculations is not None:
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": bill_calculations_index, "parent": app.root}).type_keys(bill_calculations)
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": bill_calculations_index, "parent": app.root}).type_keys(app.keys.TAB)

                    sleep(1)

                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": bill_advances_index, "parent": app.root}).click()
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": bill_advances_index, "parent": app.root}).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100)
                    if bill_advances is not None:
                        app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                          "visible_only": True, "enabled_only": True, "found_index": bill_advances_index, "parent": app.root}).type_keys(bill_advances)
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": bill_advances_index, "parent": app.root}).type_keys(app.keys.TAB)

                    sleep(1)

                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": branch_index, "parent": app.root}).click()
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": branch_index, "parent": app.root}).type_keys(app.keys.RIGHT * 70, app.keys.BACKSPACE * 130)
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": branch_index, "parent": app.root}).type_keys(branch, protect_first=True)
                    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                      "visible_only": True, "enabled_only": True, "found_index": branch_index, "parent": app.root}).type_keys(app.keys.TAB)

                    if app.wait_element({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                         "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1):
                        app.parent_switch({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0})
                        app.find_element({"title": "Нет", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        update_in_db(session, row, 'failed last', None, None,
                                     None, None, None, False)

                    print()

                    app.find_element({"title": "Провести", "class_name": "", "control_type": "Button",
                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}, timeout=1).click()
                    # sleep(1.5)
                    # if app.wait_element({"title": "Отмена проведения", "class_name": "", "control_type": "Button",
                    #                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}, timeout=.5):
                    #     break

                    if app.wait_element({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                         "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1):
                        app.parent_switch({"title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                                           "visible_only": True, "enabled_only": True, "found_index": 0})
                        app.find_element({"title": "Нет", "class_name": "", "control_type": "Button",
                                          "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                        update_in_db(session, row, 'failed last', None, None,
                                     None, None, None, False)

                    with suppress(Exception):
                        app.find_element({"title_re": r"Конфликт блокировок.*", "parent": None}, timeout=7).parent(1).find_element(
                            {
                                "title_re": "OK",
                                "control_type": "Button",
                            },
                            timeout=1,
                        ).click()

                    app.find_element({"title": "ОК", "class_name": "", "control_type": "Button",
                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()

                    update_in_db(session, row, 'success', branch, invoice,
                                 None, None, None, None)

                    print()

                    app.quit()

                else:

                    update_in_db(session, row, 'finished w/o', row.branch, None,
                                 None, None, None, None)

            else:

                if row.contragent == 'ПРОКТЕР ЭНД ГЭМБЛ КАЗАХСТАН ДИСТРИБЬЮШН ТОО (13623)':

                    update_in_db(session, row, 'processing', None, None,
                                 None, None, None, None)

                    cur_date = row.payment_date.strftime('%d.%m.%y')  # processing_date
                    search_date = cur_date
                    # search_date = cur_date - datetime.timedelta(days=1)  # datetime.datetime.strptime(cur_date, '%d.%m.%Y')

                    cur_day_index_ = calendar[calendar['Day'] == cur_date]['Type'].index[0]
                    cur_day_type_ = calendar[calendar['Day'] == cur_date]['Type'].iloc[0]

                    if cur_day_type_ == 'Holiday':
                        return 0

                    for i in range(cur_day_index_ - 1, -1, -1):
                        if calendar['Type'].iloc[i] == 'Working':
                            search_date = calendar['Day'].iloc[i]
                            search_date = datetime.date(int(f'20{search_date.split('.')[2]}'), int(search_date.split('.')[1]), int(search_date.split('.')[0]))
                            break

                    print('DATES TO PROC:', cur_date, search_date)

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

                        app.filter({'Номер': ('Равно', filtered_df['Invoice Number'].iloc[ind_]), 'Контрагент': ('Равно', row.contragent)})

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

                        invoice = str(app.find_element({"title_re": ".* Договор контрагента", "class_name": "", "control_type": "Custom",
                                                        "visible_only": True, "enabled_only": True, "found_index": 0}).element.element_info.element.CurrentName).replace(' Договор контрагента', '')
                        summ_ = str(app.find_element({"title_re": ".* Сумма документа", "class_name": "", "control_type": "Custom",
                                                      "visible_only": True, "enabled_only": True, "found_index": 0}).element.element_info.element.CurrentName).replace(' Сумма документа', '')

                        logger.warning(f'Found invoice: {invoice}')

                        all_invoices.append({invoice: summ_})

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

                        cancel_check = {"title": "Отмена проведения", "class_name": "", "control_type": "Button",
                                        "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}

                        if app.wait_element(cancel_check, timeout=5):
                            app.find_element(cancel_check).click()

                            app.find_element({"title": "Записать", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()

                        if len(all_invoices) > 1:
                            app.find_element({"title": "Список", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()

                            dds_statement = 'Крат.деб. задолж-ть покупат-ей за услуги в тенге'
                            pnl = 'Поступления от маркетинговой деятельности'
                            bill_calculations = 1210
                            bill_advances = 3510

                            for ind__, invoice in enumerate(all_invoices):

                                for key_, val_ in invoice.items():

                                    if ind__ == 0:
                                        print('Kekus')
                                        invoice_contragent = app.find_element({"title_re": ".* Договор контрагента", "class_name": "", "control_type": "Custom",
                                                                               "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}, timeout=120)
                                        invoice_contragent.click(double=True)
                                        invoice_contragent.type_keys(key_, protect_first=True)
                                        invoice_contragent.type_keys(app.keys.TAB)

                                        app.find_element({"title_re": ".*Сумма платежа", "class_name": "", "control_type": "Edit",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}, timeout=130).type_keys(val_, app.keys.TAB)

                                    else:
                                        app.find_element({"title": "Добавить", "class_name": "", "control_type": "Button",
                                                          "visible_only": True, "enabled_only": True, "found_index": 1, "parent": app.root}).click()

                                        app.find_element({"title": "Договор контрагента", "class_name": "", "control_type": "Edit",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                                        app.find_element({"title": "Договор контрагента", "class_name": "", "control_type": "Edit",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(key_, protect_first=True)
                                        app.find_element({"title": "Договор контрагента", "class_name": "", "control_type": "Edit",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(app.keys.TAB)

                                        app.find_element({"title": "Сумма платежа", "class_name": "", "control_type": "Edit",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                                        app.find_element({"title": "Сумма платежа", "class_name": "", "control_type": "Edit",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(val_, app.keys.TAB)
                                        print()
                                        app.find_element({"title": " Ставка НДС", "class_name": "", "control_type": "Custom",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                                        app.find_element({"title": " Ставка НДС", "class_name": "", "control_type": "Custom",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys('12%', protect_first=True)
                                        app.find_element({"title": " Ставка НДС", "class_name": "", "control_type": "Custom",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(app.keys.TAB * 2)

                                    app.find_element({"title": " Статья ДДС", "class_name": "", "control_type": "Custom",
                                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                                    app.find_element({"title": " Статья ДДС", "class_name": "", "control_type": "Custom",
                                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(dds_statement)

                                    app.find_element({"title": " Доход расход", "class_name": "", "control_type": "Custom",
                                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                                    app.find_element({"title": " Доход расход", "class_name": "", "control_type": "Custom",
                                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(pnl)

                                    app.find_element({"title": " Филиал", "class_name": "", "control_type": "Custom",
                                                      "visible_only": True, "enabled_only": True, "found_index": 1, "parent": app.root}).click()
                                    app.find_element({"title": " Филиал", "class_name": "", "control_type": "Custom",
                                                      "visible_only": True, "enabled_only": True, "found_index": 1, "parent": app.root}).type_keys('ТОО "Magnum Cash', app.keys.TAB)

                            print()

                            app.find_element({"title": "Провести", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()

                            app.find_element({"title": "ОК", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()

                        else:

                            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                              "visible_only": True, "enabled_only": True, "found_index": 15, "parent": app.root}).click()
                            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                              "visible_only": True, "enabled_only": True, "found_index": 15, "parent": app.root}
                                             ).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100)
                            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                              "visible_only": True, "enabled_only": True, "found_index": 15, "parent": app.root}
                                             ).type_keys('Крат.деб. задолж-ть покупат-ей за услуги в тенге (поступление) 012', protect_first=True)
                            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                              "visible_only": True, "enabled_only": True, "found_index": 15, "parent": app.root}
                                             ).type_keys(app.keys.TAB)

                            sleep(1)

                            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                              "visible_only": True, "enabled_only": True, "found_index": 18, "parent": app.root}).click()
                            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                              "visible_only": True, "enabled_only": True, "found_index": 18, "parent": app.root}
                                             ).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100, 'Поступления от маркетинговой деятельности', app.keys.TAB)

                            sleep(1)

                            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                              "visible_only": True, "enabled_only": True, "found_index": 16, "parent": app.root}).click()
                            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                              "visible_only": True, "enabled_only": True, "found_index": 16, "parent": app.root}
                                             ).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100, '1210', app.keys.TAB)

                            sleep(1)

                            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                              "visible_only": True, "enabled_only": True, "found_index": 17, "parent": app.root}).click()
                            app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                              "visible_only": True, "enabled_only": True, "found_index": 17, "parent": app.root}
                                             ).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100, '3510', app.keys.TAB)

                            sleep(1)

                            for invoice_id in all_invoices[0].keys():

                                app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                                  "visible_only": True, "enabled_only": True, "found_index": 12, "parent": app.root}).click()
                                app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                                  "visible_only": True, "enabled_only": True, "found_index": 12, "parent": app.root}
                                                 ).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100)
                                app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                                  "visible_only": True, "enabled_only": True, "found_index": 12, "parent": app.root}
                                                 ).type_keys(invoice_id, protect_first=True)
                                app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                                  "visible_only": True, "enabled_only": True, "found_index": 12, "parent": app.root}
                                                 ).type_keys(app.keys.TAB)

                                if app.wait_element({"class_name": "", "control_type": "ListItem",
                                                     "visible_only": True, "enabled_only": True, "parent": app.root}, timeout=1):
                                    els = app.find_elements({"class_name": "", "control_type": "ListItem", "visible_only": True, "enabled_only": True, "parent": app.root}, timeout=10)

                                    print('LIST DROPPED DOWN')

                                    for el in els:
                                        el.click()

                                        break

                            with suppress(Exception):

                                app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                                  "visible_only": True, "enabled_only": True, "found_index": 19, "parent": app.root}).click()
                                app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                                  "visible_only": True, "enabled_only": True, "found_index": 19, "parent": app.root}
                                                 ).type_keys(app.keys.RIGHT * 50, app.keys.BACKSPACE * 100, 'ТОО "Magnum Cash', app.keys.TAB)
                            print()
                            app.find_element({"title": "ОК", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                            update_in_db(session, row, 'success', row.branch, None,
                                         None, None, None, None)
                            print()

                else:

                    cur_date = row.payment_date  # processing_date
                    search_date = cur_date - datetime.timedelta(days=1)  # datetime.datetime.strptime(cur_date, '%d.%m.%Y')

                    print(cur_date, search_date)

                    if 2.7 <= (datetime.datetime.now() - row.payment_date).total_seconds() / 86400:
                        app = Odines()
                        app.auth()
                        kimberly_file = None

                        for file in os.listdir(kimberly_path):
                            if row.payment_date.strftime('%d.%m.%Y') in file:
                                kimberly_file = os.path.join(kimberly_path, file)

                        print('FOUND EXCEL', kimberly_file)

                        if kimberly_file is None:
                            continue

                        df = pd.read_excel(kimberly_file)

                        all_invoices = []

                        for index in range(len(df)):

                            app.parent_switch(app.root)

                            app.open('Продажа', 'Счета-фактуры выданные', 'Счет-фактура выданный', maximize_inner=True)

                            parent_ = app.parent

                            app.filter({'Номер': ('Равно', df['Ссылка'].iloc[index]), 'Контрагент': ('Равно', row.contragent), 'Сумма документа': ('Равно', 0 - df['Сумма в валюте документа'].iloc[index])})

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

                            # app.parent_switch(parent_, maximize=True)

                            if app.wait_element({"title": "Развернуть", "class_name": "", "control_type": "Button",
                                                 "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1):
                                app.find_element({"title": "Развернуть", "class_name": "", "control_type": "Button",
                                                  "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                            app.find_element({"title_re": ".* Контрагент", "class_name": "", "control_type": "Custom",
                                              "visible_only": True, "enabled_only": True, "found_index": 0}).click(double=True)

                            invoice = str(app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                                            "visible_only": True, "enabled_only": True, "found_index": 3}).element.iface_value.CurrentValue)

                            app.find_element({"title": "Закрыть", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 3}).click()

                            all_invoices.append({invoice: df['Сумма в валюте документа'].iloc[index]})

                            app.parent_back(1)

                            app.open("Окна", "Закрыть все")

                        # all_invoices = [{'ДС№1 к ДС АБ-13(04951-2016)F19 от 01 апрель 2023г': -3511800}, {'М/С-23/604169 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -459510}, {'М/С-23/605796 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -8946363}, {'М/С-23/608751 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -98976}, {'М/С-23/609514 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -14687156}, {'М/С-23/609683 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -168412}, {'М/С-23/610662 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -30671}, {'М/С-23/611527 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -16374}, {'М/С-23/612732 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -17552270}, {'М/С-23/603303 к ДМУ-12(04951-2016)-КИМБЕРЛИ': -28121895}, {'М/С-23/612583 к ДМУ-12(04951-2016)-КИМБЕРЛИ': -40084026}, {'М/С-23/612592 к ДМУ-12(04951-2016)-КИМБЕРЛИ': -86970371}, {'М/С-23/612596 к ДМУ-12(04951-2016)-КИМБЕРЛИ': -106764157}, {'М/С-23/613838 к ДМУ-12(04951-2016)-КИМБЕРЛИ': -32458138}, {'М/С-23/615780 к ДМУ-12(04951-2016)-КИМБЕРЛИ': -34940864}, {'М/С-23/616090 к ДМУ-12(04951-2016)-КИМБЕРЛИ': -32073467}, {'М/С-23/619560 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -1000000}, {'М/С-23/566254 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -10000000}, {'М/С-23/568650 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -1260000}, {'М/С-23/568783 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -135000}, {'М/С-23/568786 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -1440000}, {'М/С-23/615926 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -36658}, {'М/С-23/619567 к ДС АБ-13(04951-2016)-КИМБЕРЛИ': -16000000}, {'М/С-9/565880 к ДМУ-12(04951-2016)-КИМБЕРЛИ': -2520000}, {'ДС №3 от 01.11.2023г к ДС АБ-13(04951-2016)-F19 от 01.07.2023г.': -349096}]

                        print(all_invoices)

                        if check_fill_final_step:
                            app.parent_switch(app.root)

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
                                              "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(row.payment_date.strftime('%d.%m.%Y'))

                            app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                            app.parent_back(1)

                            app.filter({'Номер': ('Равно', row.payment_id), 'Валюта документа': ('Равно', 'KZT'), 'Вид операции': ('Равно', 'Оплата от покупателя')})

                            app.parent_back(1)

                            app.find_element({"title_re": ".* Номер", "class_name": "", "control_type": "Custom",
                                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                            app.find_element({"title_re": ".* Номер", "class_name": "", "control_type": "Custom",
                                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(app.keys.ENTER)

                            app.parent_switch({"title": "", "class_name": "", "control_type": "Pane",
                                               "visible_only": True, "enabled_only": True, "found_index": 34})

                            app.find_element({"title": "Список", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()

                            comment = str(app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                                                            "visible_only": True, "enabled_only": True, "found_index": 2, "parent": app.root}).element.iface_value.CurrentValue)

                            # for _ in range(100):
                            #     try:
                            #         app.find_element({"title": "Удалить", "class_name": "", "control_type": "Button",
                            #                           "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}, timeout=5).click()
                            #     except:
                            #         break

                            dds_statement = 'Крат.деб. задолж-ть покупат-ей за услуги в тенге'
                            pnl = 'Поступления от маркетинговой деятельности'
                            bill_calculations = 1210
                            bill_advances = 3510

                            for ind__, invoice in enumerate(all_invoices):

                                for key_, val_ in invoice.items():

                                    if ind__ == 0:
                                        print('Kekus')
                                        invoice_contragent = app.find_element({"title_re": ".* Договор контрагента", "class_name": "", "control_type": "Custom",
                                                                               "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}, timeout=120)
                                        invoice_contragent.click(double=True)
                                        invoice_contragent.type_keys(key_, protect_first=True)
                                        invoice_contragent.type_keys(app.keys.TAB)

                                        app.find_element({"title_re": ".*Сумма платежа", "class_name": "", "control_type": "Edit",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}, timeout=130).type_keys(val_, app.keys.TAB)

                                    else:
                                        app.find_element({"title": "Добавить", "class_name": "", "control_type": "Button",
                                                          "visible_only": True, "enabled_only": True, "found_index": 2, "parent": app.root}).click()

                                        app.find_element({"title": "Договор контрагента", "class_name": "", "control_type": "Edit",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                                        app.find_element({"title": "Договор контрагента", "class_name": "", "control_type": "Edit",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(key_, protect_first=True)
                                        app.find_element({"title": "Договор контрагента", "class_name": "", "control_type": "Edit",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(app.keys.TAB)

                                        app.find_element({"title": "Сумма платежа", "class_name": "", "control_type": "Edit",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                                        app.find_element({"title": "Сумма платежа", "class_name": "", "control_type": "Edit",
                                                          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(val_, app.keys.TAB)

                                    app.find_element({"title": " Статья ДДС", "class_name": "", "control_type": "Custom",
                                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                                    app.find_element({"title": " Статья ДДС", "class_name": "", "control_type": "Custom",
                                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(dds_statement)

                                    app.find_element({"title": " Доход расход", "class_name": "", "control_type": "Custom",
                                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                                    app.find_element({"title": " Доход расход", "class_name": "", "control_type": "Custom",
                                                      "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).type_keys(pnl)

                                    # app.find_element({"title": " Счет расчетов", "class_name": "", "control_type": "Custom",
                                    #                   "visible_only": True, "enabled_only": True, "found_index": ind__ + 1, "parent": app.root}).click()
                                    # app.find_element({"title": " Счет расчетов", "class_name": "", "control_type": "Custom",
                                    #                   "visible_only": True, "enabled_only": True, "found_index": ind__ + 1, "parent": app.root}).type_keys(bill_calculations)
                                    #
                                    # app.find_element({"title": " Счет авансов", "class_name": "", "control_type": "Custom",
                                    #                   "visible_only": True, "enabled_only": True, "found_index": ind__ + 1, "parent": app.root}).click()
                                    # app.find_element({"title": " Счет авансов", "class_name": "", "control_type": "Custom",
                                    #                   "visible_only": True, "enabled_only": True, "found_index": ind__ + 1, "parent": app.root}).type_keys(bill_advances)

                                    app.find_element({"title": " Филиал", "class_name": "", "control_type": "Custom",
                                                      "visible_only": True, "enabled_only": True, "found_index": 1, "parent": app.root}).click()
                                    app.find_element({"title": " Филиал", "class_name": "", "control_type": "Custom",
                                                      "visible_only": True, "enabled_only": True, "found_index": 1, "parent": app.root}).type_keys('ТОО "Magnum Cash', app.keys.TAB)

                            print()

                            app.find_element({"title": "Провести", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()

                            app.find_element({"title": "ОК", "class_name": "", "control_type": "Button",
                                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": app.root}).click()
                        # update_in_db(session, row, 'kekus', None, None,
                        #              None, None, None, None)

        except Exception as row_error:
            traceback.print_exc()
            logger.warning(f'Error on the row occured: {row_error}')
            update_in_db(session, row, 'failed', None, None,
                         None, None, None, None, error_reason_=str(traceback.format_exc())[:500])


def main():
    try:

        logger.info('Process started')
        logger.warning('Process started')

        day1_ = 20
        day2_ = 21

        # if ip_address == '10.70.2.50':
        #     day1_ = 13
        #     day2_ = 14
        # if ip_address == '10.70.2.51':
        #     day1_ = 14
        #     day2_ = 15
        # if ip_address == '10.70.2.52':
        #     day1_ = 15
        #     day2_ = 16

        for day in range(day1_, day2_):

            processing_date_ = datetime.date.today().strftime('%d.%m.%Y')
            processing_date_short_ = datetime.date.today().strftime('%d.%m.%y')

            # if day < 10:
            #     processing_date_ = f'0{day}.02.2024'
            #     processing_date_short_ = f'0{day}.02.24'
            # else:
            #     processing_date_ = f'{day}.02.2024'
            #     processing_date_short_ = f'{day}.02.24'

            logger.warning(f'Process date: {processing_date_}')

            half_year_back_date = (datetime.date(int(processing_date_.split('.')[2]), int(processing_date_.split('.')[1]), int(processing_date_.split('.')[0])) - datetime.timedelta(days=180)).strftime('%d.%m.%Y')

            performer(processing_date_, processing_date_short_, half_year_back_date)

    except Exception as error:
        logger.info(f'Error occured: {error}')
        logger.warning(f'Error occured: {error}')
        traceback.print_exc()


if __name__ == '__main__':

    try:

        main()

    except (Exception,):
        traceback.print_exc()
        kill_process_list(process_list_path)
        sys.exit(1)

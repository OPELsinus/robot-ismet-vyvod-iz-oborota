# Standard
import calendar
import datetime
import locale
import logging
import random
import string
import traceback
from pathlib import Path
from time import sleep, time
import os
import glob

# func
from datetime import timedelta
import pandas as pd
from selenium.webdriver.support.expected_conditions import visibility_of_element_located
from sqlalchemy import create_engine, Result
from sqlalchemy.exc import IntegrityError, DataError
from sqlalchemy.orm import sessionmaker

from config import logger, smtp_host, smtp_author
from models import InventoryTask
from tools import smtp
from tools.get_hostname import get_hostname
from tools.json_rw import json_read
from tools.logs import init_main_logs, init_logger
from tools.sql_support import (get_unfinished_tasks,
                               get_unfinished_task, complete_task, save_to_excel, update_res_counters_task,
                               update_res_stringtype, update_res_saving_errors, change_task_status,
                               update_res_projection, get_message_info, get_current_task_statuses)
from tools.retry import find_title_or_new_window
from tools.web import Web
from tools.tg import tg_send

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from bs4 import BeautifulSoup

import math

# Exceptions
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException

locale.setlocale(locale.LC_ALL, 'ru_RU')

project_name = 'robot-report-gold'  # !!!
local_env_data = json_read(Path.home().joinpath(f'AppData\\Local\\.rpa').joinpath('env.json'))
global_path = Path(local_env_data['global_path'])
global_env_data = json_read(global_path.joinpath('env.json'))
project_path = global_path.joinpath(fr'.agent\{project_name}')
project_config_path = project_path.joinpath('config.json')
project_branches_path = project_path.joinpath('Branches_codes.xlsx')
screens_and_logs_path = project_path.joinpath(f'{get_hostname()}\\screens_and_logs')
project_config_data = json_read(project_config_path)
username = project_config_data['username']
password = project_config_data['password']
tg_token = project_config_data['tg_token']
chat_id = project_config_data['chat_id']

engine_kwargs = {
    'username': global_env_data['postgre_db_username'],
    'password': global_env_data['postgre_db_password'],
    'host': global_env_data['postgre_ip'],
    'port': global_env_data['postgre_port'],
    # 'base': project_config_data['base_name']
}

db_url = 'postgresql+psycopg2://{username}:{password}@{host}:{port}/orchestrator'.format(**engine_kwargs)
engine = create_engine(db_url)

url_gold = project_config_data['url_gold']

gold_login = project_config_data['username']
test_gold_login = project_config_data['test_login']

gold_password = project_config_data['password']
test_password = project_config_data['test_password']


path_spreadsheet = project_config_data['spreadsheet']

path_downloaded = Path.home().joinpath('Downloads')

months_dict = {
    "January": "01",
    "February": "02",
    "March": "03",
    "April": "04",
    "May": "05",
    "June": "06",
    "July": "07",
    "August": "08",
    "September": "09",
    "October": "10",
    "November": "11",
    "December": "12"
}


# Просто рандомный текст
# noinspection PyShadowingNames
def generate_random_text(min_length, max_length):
    characters = string.ascii_letters + string.digits

    # Randomly choose the length between min_length and max_length
    length = random.randint(min_length, max_length)

    # Generate a random string of the chosen length
    random_text = ''.join(random.choice(characters) for _ in range(length))

    return random_text


def send_message(message: str = None, html: str = None, receivers: list = None, attachments: list = None):
    subject = "test"
    host = smtp_host
    author = smtp_author
    html = html if html else None
    smtp.smtp_send(message + "\nTEST" if message else "", subject=subject, username=author,
                   url=host, to=receivers, html=html, attachments=attachments)
    return True


# noinspection PyShadowingNames
def delete_files(files):
    """Удаляет список файлов."""
    for file_path in files:
        try:
            os.remove(file_path)
        except OSError as e:
            print(f"Ошибка при удалении файла {file_path}: {e.strerror}")


def convert_share(share_str):
    numerator_str, denominator_str = share_str.split("/")
    numerator = float(numerator_str)
    float(denominator_str)
    return numerator / 100


# noinspection PyShadowingNames
def extract_info(path: Path = None, eng=engine):
    # Открытие таблицы, конверт в датафреймы
    excel_file = pd.ExcelFile(path)
    df = pd.read_excel(excel_file, 'ЛИ в ТЗ')
    df['Индекс'] = df.index + 2
    df2 = pd.read_excel(excel_file, 'Периодичность')

    df = df.dropna(subset=['Доля товарного запаса ТЗ/Склад'])
    df['Доля товарного запаса ТЗ/Склад'] = df['Доля товарного запаса ТЗ/Склад'].apply(convert_share)
    print(1)
    df['minus_day'] = pd.to_datetime(df['Дата'], format='%Y.%m.%d') - pd.to_timedelta(
        df['Кол-во дней до инвентаризации'], unit='d')
    # res df со всеми задачами на сегодня
    today = pd.Timestamp('now').normalize()  # !!!
    str(today)
    # TODO Проверка на ошибку в слове повтор или ошибку в пустой ячейке
    # ПОВТОР В СТАТУС ОБРАБОТКИ? ИЛИ ПРИМЕЧАНИЕ
    # FIXME ASKME ЕСЛИ В ПРИМЕЧАНИИ ЧТО-ТО СТОИТ, ДЕЛАТЬ ИЛИ НЕТ?
    # res_df = df[(df['minus_day'] == pd.to_datetime('2023-11-02')) | (
    #         df['Статус отработки'] == "Повтор")].reset_index(drop=True)
    condition_A = (
                          (df['minus_day'] == today) & df['Статус отработки'].isna()
                  ) | (
                          df['Статус отработки'] == "Повтор"
                  )

    condition_B = df['Примечание'].isna()
    res_df = df[condition_A & condition_B].reset_index(drop=True)

    if res_df.empty and len(res_df.index) == 0:
        return False

    # | (df['Статус отработки'] == "")
    # Добавить в задачу столбец со значениями отделов которые нужно внести в ракурс
    def add_list_value(r):
        key = r['Периодичность']
        its_list = df2[df2['Периодичность'].str.lower() == key.lower()]['Отдел'].tolist()
        return its_list

    res_df = res_df.copy()
    res_df['Отдел'] = res_df.apply(add_list_value, axis=1)

    # add a database
    Session = sessionmaker(bind=eng)
    with Session() as session:
        # Создание задачи в бд
        # Копии задач не создаются, только если Повтор
        res_df = res_df.where(pd.notna(res_df), None)
        try:
            for _, row in res_df.iterrows():

                existing_record = session.query(InventoryTask).filter_by(
                    index=int(row['Индекс']),
                    branch=row['Филиал'],
                    location_code=row['Код площадки'],
                    inventory_type=row['Тип инвентаризации'],
                    frequency=row['Периодичность'],
                    date=row['Дата'],
                    rto=row['РТО'],
                    regional_director=row['Региональный директор'],
                    security_service_head=row['Начальник службы безопасности']
                ).first()

                if not existing_record or (row['Статус отработки'] == 'Повтор' and existing_record):
                    inventory_task = InventoryTask(
                        index=int(row['Индекс']),
                        branch=row['Филиал'],
                        location_code=row['Код площадки'],
                        inventory_type=row['Тип инвентаризации'],
                        frequency=row['Периодичность'],
                        date=row['Дата'],
                        note=row['Примечание'],
                        auditor=row['ревизор'],
                        auditor_us_gold=row['Ревизор в УС Голд'],
                        format=row['Формат'],
                        processing_status=row['Статус отработки'],
                        mailing_status=row['Статус отправки рассылки'],
                        days_until_inventory=row['Кол-во дней до инвентаризации'],
                        stock_share_tz_warehouse=row['Доля товарного запаса ТЗ/Склад'],
                        rto=row['РТО'],
                        regional_director=row['Региональный директор'],
                        territorial_director=row['Территориальный директор'],
                        branch_director=row['Директор филиала'],
                        deputy_director_administrator=row['Заместитель директора филиала/Администратор торгового зала'],
                        security_service_head=row['Начальник службы безопасности'],
                        minus_day=row['minus_day'],
                        division=row['Отдел'],

                        done_projection=None,
                        done_pallets=None,
                        done_closed_docs=None,
                        done_counters_task=None,
                        done_mailing_superiors=None,
                        done_saving_errors=None
                    )
                    session.add(inventory_task)

        except IntegrityError as e:
            # Обработка ошибок целостности данных (например, дублирование ключей)
            session.rollback()
            logger.error("Ошибка IntegrityError при записи в базу данных: %s", str(e), f"\n{traceback.format_exc()}")
            return False
        except DataError as e:
            # Обработка ошибок данных (например, неверный формат данных)

            session.rollback()
            logger.error("Ошибка DataError при записи в базу данных: %s", str(e), f"\n{traceback.format_exc()}")
            return False
        except Exception as e:
            # Обработка других неожиданных ошибок
            session.rollback()
            logger.error("Неожиданная ошибка при записи в базу данных: %s", str(e), f"\n{traceback.format_exc()}")
            return False

        finally:
            if excel_file:
                excel_file.close()
            session.commit()
    return True


def log_into(depth: int = 0):
    def log():
        if app.wait_element('//input[@id="5Инвентаризацияlabel"]', timeout=2, event=visibility_of_element_located):
            return True
        else:
            app.get(url_gold)

            web_element = app.find_element('//input[@id="loginUserName"]')
            web_element.type_keys(gold_login, clear=True)
            sleep(1.5)

            web_element = app.find_element('//input[@id="loginPassword"]')
            web_element.type_keys(gold_password, clear=True)
            sleep(1.5)

            time1 = time()
            print(f"time 1 = {time1}")
            web_element = app.find_element('//tbody/tr[4]/td[2]/input[1]')
            web_element.type_keys(task.location_code)
            sleep(1.5)

            web_element = app.find_element('//input[@id="DOMButton_loginButton"]')
            web_element.click(double=False, delay=0.5, page_load=False, scroll=True)

            if app.wait_element('//input[@id="5Инвентаризацияlabel"]', timeout=10, event=visibility_of_element_located):
                return True
            else:
                return False

    if depth == 5:
        return False
    elif log():
        return True
    else:
        return log_into(depth + 1)


def bs_parse_find_by_class(html, class_str: str = None):
    """
    Searches for the text that is unreachable by selenium methds
    :param class_str: class name str
    :param html: html page str like
    :return: value: str
    """
    soup = BeautifulSoup(html, "html.parser")
    element = soup.find("td", class_=class_str)
    value = element.text
    return value


# noinspection PyBroadException,PyShadowingNames
def gold_projection(task: Result = None):
    r"""
    Creates a projection on US GOLD
    Создаёт ракурс на УС ГОЛД

    * **params**  description
    :return: bool
    """

    gold_report_pallets_logger = init_inside_logger('gold_projection', task.id)
    error_fp = screens_and_logs_path.joinpath(f"{task.id}")

    try:
        assert task is not None
        logger.info(f"Task id - {task.id}, specification - projection")

        # Проверка, запущен ли драйвер
        if app.driver.current_url != url_gold:
            app.get(url_gold)

            web_element = app.find_element("//input[@id='loginUserName']")
            # web_element.type_keys("Kuanyshov.A", clear=True)
            web_element.type_keys(gold_login, clear=True)
            sleep(1.5)

            web_element = app.find_element("//input[@id='loginPassword']")
            # web_element.type_keys(gold_password, clear=True)
            web_element.type_keys(gold_password, clear=True)
            sleep(1.5)

            time1 = time()
            print(f"time 1 = {time1}")
            web_element = app.find_element("//tbody/tr[4]/td[2]/input[1]")
            web_element.type_keys(task.location_code)
            sleep(1.5)

            web_element = app.find_element("//input[@id='DOMButton_loginButton']")
            web_element.click(double=False, delay=0.5, page_load=False, scroll=True)

        sleep(0.5)

        app.get(url_gold)
        # print(app.wait_element("//div[@id='popupBlock']", timeout=500, event=visibility_of_element_located))

        # TODO optimize the time

        # print(f'''{web_element.wait_element(selector=' // input[ @ id = "5Инвентаризацияlabel"]', timeout=15)}''')
        print(app.wait_element(' // input[ @ id = "5Инвентаризацияlabel"]', timeout=15,
                               event=visibility_of_element_located))

        # TODO Добавить ожидание прогрузки hidden=True
        sleep(5)

        print('этап 2')
        # Раскрывающееся окно
        web_element = app.find_element("//input[@id='5Инвентаризацияlabel']")
        web_element.click(double=False, delay=0.2, page_load=False, scroll=True)
        sleep(1)

        # Выбор варианта
        web_element = app.find_element("//input[@id='5-2Создание инвентаризацииlabel']")
        web_element.click(double=False, delay=0.2, page_load=False, scroll=True)
        sleep(1.5)

        # Ожидание
        find_title_or_new_window(app=app, selector="//input["
                                                   "@id='GSOInventoryInitializationAddPanel_inventoryDescription' and"
                                                   " @style='width: 300px;']", initial_window_count=len(
            app.driver.window_handles))
        # ==========================================Ввод данных=============================================

        # Тип инвентаризации
        web_element = app.find_element("//input[@id='GSOInventoryInitializationAddPanel_inventoryType']", timeout=3)
        web_element.type_keys("22")


        # Выбор и ввод описания
        web_element = app.find_element("//input[@id='GSOInventoryInitializationAddPanel_inventoryDescription' and "
                                       "@style='width: 300px;']", timeout=2)
        formatted_date = task.date.strftime('%d.%m.%Y')
        # Разделение даты на день, месяц и год

        month_name = calendar.month_name[task.date.month]
        formatted_text = f"{task.frequency} {month_name.lower()} {formatted_date}"
        web_element.type_keys(formatted_text)
        sleep(1)

        # Дата создания инвентаризации
        web_element = app.find_element("//input[@id='GSOInventoryInitializationAddPanel_inventoryCreationDate'"
                                       " and @title='DD/MM/YY']", timeout=3)
        web_element.click(delay=0.2, scroll=True)
        sleep(1)
        # Получение HTML-кода текущей страницы
        html = app.driver.page_source
        # Значение строки с текстом Месяца и года, недоступным из selenium
        value = bs_parse_find_by_class(html=html, class_str="calendarHead calendarMonth")
        # Защитный клик для открытия календаря
        web_element.click(delay=0.2, scroll=True)

        formatted_date = task.minus_day.strftime('%d.%m.%Y')
        day, month, year = formatted_date.split('.')
        day = day.lstrip("0")
        value_m, value_y = value.split(" ")
        value_m = months_dict[value_m]
        value = datetime.datetime.strptime(str(value_m) + " " + value_y, "%m %Y")
        formatted_date = datetime.datetime.strptime(formatted_date, '%d.%m.%Y')
        # Вычислите разницу в месяцах между formatted_date и value
        months_difference = (formatted_date.year - value.year) * 12 + (formatted_date.month - value.month)
        print(formatted_date)
        print(f"Month difference is {months_difference}")
        if months_difference > 0:
            for _ in range(months_difference):
                next_month_button = app.find_element(selector="//td[@class='calendarHead' and contains(text(), '>')]",
                                                     timeout=1)
                next_month_button.click(delay=0.3, scroll=True)
        elif months_difference < 0:
            for _ in range(months_difference):
                next_month_button = app.find_element(selector="//td[@class='calendarHead' and contains(text(), '<')]",
                                                     timeout=1)
                next_month_button.click(delay=0.3, scroll=True)
        else:
            pass

        web_element = app.find_element(selector=f"//td[contains(@class, 'calendarCell') and contains(text(), {day})]",
                                       timeout=2)

        web_element.click(delay=0.5, scroll=True)

        sleep(1)

        # noinspection PyBroadException
        try:
            web_element = app.find_element("//td[normalize-space()='x']", timeout=10)
            web_element.click(double=False, delay=0.2, page_load=False, scroll=True)
            sleep(1)
            logger.info('Найден х')
        except NoSuchElementException:
            logger.info('Не найден х')
        except ElementNotInteractableException:
            logger.info('Не найден х')
        except Exception:
            logger.error(f'Unexpected error\n{traceback.format_exc()}')

        # Торг. структ
        plus_button = app.find_element("//input[@id='boutonAddMSNodeLinesTable']")
        n = 0
        for item in task.division:
            plus_button.click(double=False, delay=0.2, page_load=False, scroll=True)
            if app.wait_element(f"//*[@id='MSNodeLinesTable_{n}_merchandiseStructureNodeCode']", timeout=5,
                                event=visibility_of_element_located) is False:
                plus_button.click(double=False, delay=0.2, page_load=False, scroll=True)
            web_element = app.find_element(f"//*[@id='MSNodeLinesTable_{n}_merchandiseStructureNodeCode']")
            web_element.click(double=False, delay=0.2, page_load=False, timeout=2, scroll=True)
            sleep(0.3)
            line_n = app.find_element(f"//input[@id='MSNodeLinesTable_{n}_merchandiseStructureNodeCode_Field']")
            line_n.type_keys(item)
            if n == len(task.division) - 1:
                # Если это последняя итерация, то добавляем задержку и нажимаем клавишу Enter
                sleep(0.5)  # Задержка, если необходимо
                line_n.type_keys(app.keys.ENTER)
            n += 1

        app.find_element("//input[@id='GSOInventoryInitializationAddPanel_inventoryDescription' and "
                         "@style='width: 300px;']", timeout=2).click(timeout=0.5, scroll=True)

        # Найдите элемент таблицы
        table_element = app.find_element('//div[@id="tableContentFrameMSNodeLinesTable"]', timeout=1)

        # Найдите все строки таблицы
        table_rows = table_element.find_elements(selector='.//div[contains(@class, "slick-row")]', timeout=3)

        # Инициализируйте пустой список для хранения данных
        table_data = []

        # Пройдите по каждой строке таблицы и извлеките данные
        for row in table_rows:
            code_element = row.find_element(selector='.//div[@id="MSNodeLinesTable_' + row.get_attr('linenumber')
                                                     + '_merchandiseStructureNodeCode"]', timeout=1)
            description_element = \
                row.find_element(selector='.//div[@id="MSNodeLinesTable_' + row.get_attr('linenumber')
                                          + '_merchandiseStructureNodeDescription"]', timeout=1)

            code = code_element.element.get_attribute('text')
            description = description_element.element.get_attribute('text')

            # Добавьте данные в виде кортежа в список
            table_data.append([code, description])

        app.find_element(selector="//button[@title='Подтвердить (ALT+CTRL+V)']").click(scroll=True, timeout=0.5)
        # press validate and check for errors
        wait = WebDriverWait(app.driver, 600)  # Максимальное время ожидания в секундах

        # Замените следующую строку на поиск вашего элемента
        try:
            if not wait.until(
                    EC.visibility_of_element_located((By.XPATH, "//pre[contains(text(), '№ инвентаризации')]"))):
                time0 = time()
                found = False
                while not found or time() - time0 <= 600:
                    app.find_element("//pre[contains(text(), '№ инвентаризации')]")

            result_message = app.find_element(selector="//pre[contains(text(), '№ инвентаризации')]", timeout=1)
            pre_text = result_message.element.get_attribute("text")
            parts = pre_text.splitlines()
            inventory_number = None
            row_count = None
            for part in parts:
                if part.startswith('№ инвентаризации :'):
                    inventory_number = (part.split(':')[-1])
                elif part.startswith('Кол-во строк :'):
                    row_count = (part.split(':')[-1])

            if inventory_number is None or row_count is None:
                raise ValueError("Не найден нужный текст")

            finish_window = app.find_element(selector="//button[@id='validatePopup']", timeout=20)
            finish_window.click(timeout=10, scroll=True)
            try:
                app.find_element(selector="//pre[contains(text(), 'Подтверждение завершено успешно')]", timeout=15)
                app.find_element(selector="//button[@id='validatePopup']", timeout=20).click(timeout=10, scroll=True)
                sleep(7)
            except:
                logger.error(f"Couldn't get the projection.")
                save_screen(task.id, 'gold_projection')
                ok = True
                while ok:
                    try:
                        app.find_element(selector="//button[@id='validatePopup']", timeout=1).click(timeout=0.5,
                                                                                                    scroll=True)
                    except NoSuchElementException:
                        ok = False
                    except Exception:
                        ok = False
                        log_error(gold_report_pallets_logger,
                                  f"Ошибка во закрытия окон ОК. \n{traceback.format_exc()}")
                return False, error_fp

        except Exception:
            logger.error(f"Couldn't get the projection. See below:\n{traceback.format_exc()}")
            save_screen(task.id, 'gold_projection')
            log_error(gold_report_pallets_logger, f"Ошибка во время валидации ракурса. \n{traceback.format_exc()}")
            ok = True
            while ok:
                try:
                    app.find_element(selector="//button[@id='validatePopup']", timeout=1).click(timeout=0.5,
                                                                                                scroll=True)
                except NoSuchElementException:
                    ok = False
                except Exception:
                    ok = False
                    log_error(gold_report_pallets_logger,
                              f"Ошибка во закрытия окон ОК. \n{traceback.format_exc()}")
            return False, error_fp

        sleep(15)

        return True, (inventory_number, table_data)
    except AssertionError as e:
        logger.error(f"Assertion error в моменте бд\n{e}")
        save_screen(task.id, 'gold_projection')
        log_error(gold_report_pallets_logger, f"Assertion error в моменте бд\n{traceback.format_exc()}")
        return False, error_fp
    except Exception:
        logger.error(f"Unknown problem. See below:\n{traceback.format_exc()}")
        save_screen(task.id, 'gold_projection')
        log_error(gold_report_pallets_logger, f"НЕОЖИДАННАЯ ОШИБКА FATAL ERROR \n{traceback.format_exc()}")
        return False, error_fp


def find_file_in_download_folder(target_string):
    files = glob.glob(os.path.join(path_downloaded, "*.xlsx"))  # Поиск всех файлов с расширением .xlsx
    files = sorted(files, key=os.path.getmtime, reverse=True)
    for file_path in files:
        if target_string in file_path:
            return file_path  # Возврат пути к файлу, если строка найдена
    for file_path in files:
        return file_path
    return None  # Если файл не найден


# report !!!
# noinspection PyShadowingNames
def gold_report_pallets(task: Result = None):
    r"""

    :return: (res: bool, path: Path = Path(fp))
    """

    gold_report_pallets_logger = init_inside_logger('gold_report_pallets', task.id)
    error_fp = screens_and_logs_path.joinpath(f"{task.id}")

    logger.info(f"Task id - {task.id}, specification - report pallets")

    # Проверка, запущен ли драйвер
    if app.driver.current_url != url_gold or not app.wait_element(selector="//input[@id='1Администрированиеlabel']",
                                                                  timeout=60, event=visibility_of_element_located):
        if log_into():
            pass
        else:
            sleep(25)
            if log_into():
                pass
            else:
                save_screen(task.id, 'gold_report_pallets')
                log_error(gold_report_pallets_logger, 'Не удалось залогиниться')
                return False, error_fp

    web_element = app.find_element("//input[@id='1Администрированиеlabel']")
    app.driver.execute_script("window.scrollTo(0, 0);")
    web_element.click(double=False, delay=0.2, page_load=False, scroll=True)

    sleep(1)

    web_element = app.find_element("//input[@id='1-1Запуск отчетовlabel']")
    web_element.click(double=False, delay=0.2, page_load=False, scroll=True)

    # Открытие отчёта на новой странице
    web_element = app.find_element(selector="//div[@class='treeNode treeRootNode']")
    text_selector = f'./div[@nodecode="GOLDREPORT-148"]'
    laying_selector = "//div[@class='childrenContainer']"
    laying_brother_selector = "//div[@class='treeRow']"
    child_selector = "./div[@class='treeNode']"
    if web_element.recursive_open_until_interactable_and_click(text_selector, laying_selector, laying_brother_selector,
                                                               child_selector):
        pass
    else:
        # Screen save
        # Не НАЙДЕН ДОК В СПИСКЕ ОТЧЕТОВ
        save_screen(task.id, 'gold_report_pallets')
        log_error(gold_report_pallets_logger, 'НЕ НАЙДЕН ДОК В СПИСКЕ ОТЧЕТОВ')
        return False, error_fp

    # Кликаем на
    try:
        sleep(2)
        app.driver.execute_script("window.scrollTo(0, 0);")
        app.find_element(selector="//button[@id='toolbarValidateGSOReportAndQueryExecution']",
                         timeout=15, event=visibility_of_element_located).click(delay=1, scroll=True)
    except Exception as e:
        # Screen save
        save_screen(task.id, 'gold_report_pallets')
        log_error(gold_report_pallets_logger, f'НЕ УДАЛОСЬ ВАЛИДИРОВАТЬ НАЖАВ НА ГАЛОЧКУ\n{e}')
        return False, error_fp

    # Сохранить идентификатор основного окна
    main_window = app.driver.current_window_handle

    sleep(5)
    all_windows = app.driver.window_handles
    for window in all_windows:
        if window != main_window:
            app.driver.switch_to.window(window)

    try:
        # Загрузка данных на новой странице
        # ============================================================================================================ #
        # Дата начала и дата конца
        try:
            web_element = app.find_element(selector='//*[@id="username"]', timeout=1)
            web_element.type_keys(gold_login, delay=0.1, clear=True)
            app.find_element(selector='//*[@id="password"]', timeout=1).type_keys(gold_password, delay=0.1, clear=True)
            app.find_element(selector='//*[@id="pageContent"]/div/div/div/form/fieldset/div[5]/div/button',
                             timeout=1).click(delay=0.2, scroll=True)
        except:
            pass

        # Done Дата С
        web_element = app.find_element("//input[@id='p-PRM_BEGIN_DATE']", timeout=2)
        web_element.type_keys((task.date - timedelta(days=30)).strftime("%d.%m.%Y"),
                              delay=0.01, clear=True)

        # DONE Дата по
        web_element = app.find_element(selector="//input[@id='p-PRM_END_DATE']", timeout=1)
        web_element.type_keys(task.date.strftime("%d.%m.%Y"), delay=0.01, clear=True)

        # Выбрать все РЦ
        web_element = app.find_element(selector="//button[@data-id='p-PRM_GOLD_SITES_DC_WMS']", timeout=1)
        web_element.click(double=False, delay=0.2, page_load=False, scroll=True)
        sleep(1)
        app.find_element(selector="//button[@type='button'][contains(text(),'Выбрать все')]",
                         timeout=3).click(double=False, delay=0.2, page_load=False, scroll=True)
        # Всё выбрано, закрыть окно
        web_element = app.find_element(selector="//button[@data-id='p-PRM_GOLD_SITES_DC_WMS']", timeout=1)
        web_element.click(double=False, delay=0.2, page_load=False, scroll=True)

        # Выбрать нужную торговую площадку
        web_element = app.find_element(selector="//button[@data-id='p-PRM_GOLD_SITES_USER']", timeout=1)
        web_element.click(double=False, delay=0.2, page_load=False, scroll=True)
        # input_el = app.find_element(selector="//input[@aria-label='Search']", timeout=1)
        sleep(1)
        app.find_element(selector=f"//option[@value='{task.location_code}']").click(delay=0.2, scroll=True,
                                                                                    page_load=False,
                                                                                    double=False)
        # input_el.type_keys(value=task.location_code, delay=0.2, scroll=False, clear=True)
        # input_el.type_keys(value=app.keys.ENTER, delay=0.1, scroll=False, clear=True)

        # ???
        # web_element = app.find_element(selector="//button[@data-id='p-PRM_GOLD_SITES_USER']", timeout=1)
        # web_element.click(double=False, delay=0.2, page_load=False, scroll=True)

        # Статус паллеты
        web_element = app.find_element(selector="//button[@data-id='p-PRM_GOLD_STATE_PALLET']", timeout=1)
        web_element.click(double=False, delay=0.2, page_load=False, scroll=True)
        sleep(3)
        input_el = app.find_element(selector="//div[@class='btn-group bootstrap-select show-tick form-control "
                                             "open']//input[@aria-label='Search']", timeout=1)
        # Отгруженно
        input_el.click(double=False, delay=0.2, page_load=False, scroll=True)
        input_el.type_keys("Отгруженно", delay=0.1, scroll=False, clear=True)
        input_el.type_keys(app.keys.ENTER, delay=0.1, scroll=False, clear=True)

        # Получено
        input_el.click(double=False, delay=0.2, page_load=False, scroll=True)
        input_el.type_keys("Получено", delay=0.1, scroll=False, clear=True)
        input_el.type_keys(app.keys.ENTER, delay=0.1, scroll=False, clear=True)

        # web_element.type_keys(value=app.keys.ENTER, delay=0.1, scroll=False, clear=True)
        web_element = app.find_element(selector="//button[@data-id='p-PRM_GOLD_STATE_PALLET']", timeout=1)
        web_element.click(double=False, delay=0.2, page_load=False, scroll=True)

        try:
            # web_element = app.find_element(selector="//ul[@class='jstree-container-ul jstree-children']", timeout=1)
            web_element = app.find_element(selector="//li[a[text()='ROOT']]", timeout=1)
            web_element.find_element(selector="//i[@class='jstree-icon "
                                              "jstree-ocl']", timeout=0.2).click(delay=0.2, scroll=True)
            for item in task.division:
                web_element.open_until_interactable(item)


        except Exception as e:
            # Screen save
            save_screen(task.id, 'gold_report_pallets')
            log_error(gold_report_pallets_logger, f'НЕ УДАЛОСЬ {e}')
            return False, error_fp

        fp = find_new_xlsx_file(path_downloaded, sel="//button[@type='submit' and contains(text(), 'Запустить')]",
                                timeout=900)

    except Exception:
        save_screen(task.id, 'gold_report_pallets')
        log_error(gold_report_pallets_logger, f'{traceback.format_exc()}')
        return False, error_fp

    finally:
        if app.driver.current_window_handle != main_window:
            app.driver.close()
            app.driver.switch_to.window(main_window)

    return True, fp


def find_new_xlsx_file(folder, sel: str, timeout=600):
    start_time = time()
    app.find_element(selector=sel,
                     timeout=1).click(double=False, delay=0, page_load=False, scroll=True)
    while time() - start_time < timeout:
        for file in folder.glob("*.xlsx"):
            if file.stat().st_mtime > start_time and file.stem.startswith(
                    f"GOLD"):
                return file
        sleep(5)
    return None


# noinspection PyBroadException,PyShadowingNames
def gold_closed_docs(task: Result = None):
    gold_counters_logger = init_inside_logger('gold_closed_docs', task.id)
    error_fp = screens_and_logs_path.joinpath(f"{task.id}")

    logger.info(f"Task id - {task.id}, specification - report closed_docs")

    # Проверка, запущен ли драйвер
    if app.driver.current_url != url_gold or not app.wait_element(selector="//input[@id='1Администрированиеlabel']",
                                                                  timeout=60, event=visibility_of_element_located):
        if log_into():
            pass
        else:
            sleep(25)
            if log_into():
                pass
            else:
                save_screen(task.id, 'gold_closed_docs')
                log_error(gold_counters_logger, 'Не удалось залогиниться')
                return False, error_fp

    web_element = app.find_element("//input[@id='1Администрированиеlabel']")
    app.driver.execute_script("window.scrollTo(0, 0);")
    web_element.click(double=False, delay=0.2, page_load=False, scroll=True)

    sleep(1)

    web_element = app.find_element("//input[@id='1-1Запуск отчетовlabel']")
    web_element.click(double=False, delay=0.2, page_load=False, scroll=True)

    # Открытие отчёта на новой странице
    web_element = app.find_element(selector="//div[@class='treeNode treeRootNode']")
    text_selector = f'./div[@nodecode="GOLDREPORT-26"]'
    laying_selector = "//div[@class='childrenContainer']"
    laying_brother_selector = "//div[@class='treeRow']"
    child_selector = "./div[@class='treeNode']"
    if web_element.recursive_open_until_interactable_and_click(text_selector, laying_selector, laying_brother_selector,
                                                               child_selector):
        pass
    else:
        # Screen save
        save_screen(task.id, 'gold_closed_docs')
        log_error(gold_counters_logger, 'Не найден Документ в списке отчётов \n '
                                        'recursive_open_until_interactable_and_click')
        return False, error_fp

    # Кликаем на галочку
    try:
        sleep(2)
        app.driver.execute_script("window.scrollTo(0, 0);")
        app.find_element(selector="//button[@id='toolbarValidateGSOReportAndQueryExecution']",
                         timeout=15, event=visibility_of_element_located).click(delay=0.4, scroll=True)
    except Exception as e:
        # Screen save
        save_screen(task.id, 'gold_closed_docs')
        log_error(gold_counters_logger, f'Не удалось валидировать, нажав на галочку \n{e}')
        return False, error_fp

    # Сохранить идентификатор основного окна
    main_window = app.driver.current_window_handle

    sleep(5)
    all_windows = app.driver.window_handles
    for window in all_windows:
        if window != main_window:
            app.driver.switch_to.window(window)

    try:
        # Загрузка данных на новой странице
        # ============================================================================================================ #
        # Дата начала и дата конца
        try:
            web_element = app.find_element(selector='//*[@id="username"]', timeout=1)
            web_element.type_keys(gold_login, delay=0.1, clear=True)
            app.find_element(selector='//*[@id="password"]', timeout=1).type_keys(gold_password, delay=0.1, clear=True)
            app.find_element(selector='//*[@id="pageContent"]/div/div/div/form/fieldset/div[5]/div/button',
                             timeout=1).click(delay=0.2, scroll=True)
        except:
            pass

        # Выбрать нужную торговую площадку
        web_element = app.find_element(selector="//button[@data-id='p-PRM_GOLD_SITES_USER']", timeout=1)
        web_element.click(double=False, delay=0.2, page_load=False, scroll=True)
        # input_el = app.find_element(selector="//input[@aria-label='Search']", timeout=1)
        sleep(1)
        app.find_element(selector=f"//option[@value='{task.location_code}']").click(delay=0.2, scroll=True,
                                                                                    page_load=False,
                                                                                    double=False)

        #
        web_element = app.find_element(selector="//button[@data-id='p-PRM_GOLD_SITES_USER']", timeout=1)
        if web_element.get_attr('aria-expanded') == 'true':
            web_element.click(double=False, delay=0.2, page_load=False, scroll=True)

    except Exception as e:
        save_screen(task.id, 'gold_closed_docs')
        log_error(gold_counters_logger, f'Загрузка данных на новой странице\n{e}')
        return False, error_fp

    # Группы товаров ROOT-ROOT
    try:
        web_element = app.find_element(selector="//li[a[text()='ROOT']]", timeout=1)
        web_element.find_element(selector="//i[@class='jstree-icon "
                                          "jstree-ocl']", timeout=0.2).click(delay=0.2, scroll=True)
        for item in task.division:
            web_element.open_until_interactable(item)

        fp = find_new_xlsx_file(path_downloaded, sel="//button[@type='submit' and contains(text(), 'Запустить')]")

    except Exception as e:
        # Screen save
        save_screen(task.id, 'gold_closed_docs')
        log_error(gold_counters_logger, f'Группы товаров ROOT-ROOT\n{e}')
        return False, error_fp

    finally:
        if app.driver.current_window_handle != main_window:
            app.driver.close()
            app.driver.switch_to.window(main_window)

    return True, fp


# noinspection PyBroadException, PyShadowingNames
def gold_counters_report(task: Result = None):
    gold_counters_logger = init_inside_logger('gold_counters_report', task.id)
    error_fp = screens_and_logs_path.joinpath(f"{task.id}")

    logger.info(f"Task id - {task.id}, specification - report counters report")

    # Проверка, запущен ли драйвер
    if app.driver.current_url != url_gold or not app.wait_element(selector="//input[@id='1Администрированиеlabel']",
                                                                  timeout=60, event=visibility_of_element_located):
        if log_into():
            pass
        else:
            sleep(25)
            if log_into():
                pass
            else:
                save_screen(task.id, 'gold_counters_report')
                log_error(gold_counters_logger, 'Не удалось залогиниться')
                return False, (error_fp, "")

    web_element = app.find_element("//input[@id='1Администрированиеlabel']")
    app.driver.execute_script("window.scrollTo(0, 0);")
    web_element.click(double=False, delay=0.2, page_load=False, scroll=True)

    sleep(1)

    web_element = app.find_element("//input[@id='1-1Запуск отчетовlabel']")
    web_element.click(double=False, delay=0.2, page_load=False, scroll=True)

    # Открытие отчёта на новой странице
    web_element = app.find_element(selector="//div[@class='treeNode treeRootNode']")
    text_selector = f'./div[@nodecode="GOLDREPORT-33"]'
    laying_selector = "//div[@class='childrenContainer']"
    laying_brother_selector = "//div[@class='treeRow']"
    child_selector = "./div[@class='treeNode']"
    if web_element.recursive_open_until_interactable_and_click(text_selector, laying_selector, laying_brother_selector,
                                                               child_selector):
        pass
    else:
        # Screen save
        save_screen(task.id, 'gold_counters_report')
        log_error(gold_counters_logger, 'Не найден Документ в списке отчётов')
        return False, (error_fp, "")

    # Кликаем на галочку
    try:
        sleep(2)
        app.driver.execute_script("window.scrollTo(0, 0);")
        app.find_element(selector="//button[@id='toolbarValidateGSOReportAndQueryExecution']",
                         timeout=15, event=visibility_of_element_located).click(delay=0.4, scroll=True)
    except:
        # Screen save
        save_screen(task.id, 'gold_counters_report')
        log_error(gold_counters_logger, 'Не удалось валидировать по галочке')
        return False, (error_fp, "")

    # Сохранить идентификатор основного окна
    main_window = app.driver.current_window_handle

    sleep(5)
    all_windows = app.driver.window_handles
    for window in all_windows:
        if window != main_window:
            app.driver.switch_to.window(window)

    try:
        # Загрузка данных на новой странице
        # ============================================================================================================ #
        # Дата начала и дата конца
        try:
            web_element = app.find_element(selector='//*[@id="username"]', timeout=1)
            web_element.type_keys(gold_login, delay=0.1, clear=True)
            app.find_element(selector='//*[@id="password"]', timeout=1).type_keys(gold_password, delay=0.1, clear=True)
            app.find_element(selector='//*[@id="pageContent"]/div/div/div/form/fieldset/div[5]/div/button',
                             timeout=1).click(delay=0.2, scroll=True)
        except:
            pass

        # Выбрать нужную торговую площадку
        web_element = app.find_element(selector="//button[@data-id='p-PRM_GOLD_SITES_USER']", timeout=1)
        web_element.click(double=False, delay=0.2, page_load=False, scroll=True)
        sleep(1)
        app.find_element(selector=f"//span[contains(text(), '{task.location_code}')]/../..").click(delay=0.2,
                                                                                                   scroll=True,
                                                                                                   page_load=False,
                                                                                                   double=False)
        sleep(0.5)
        web_element = app.find_element(selector="//button[@data-id='p-PRM_GOLD_SITES_USER']", timeout=1)
        web_element.click(double=False, delay=0.2, page_load=False, scroll=True)

    except Exception:
        save_screen(task.id, 'gold_counters_report')
        log_error(gold_counters_logger, f'{traceback.format_exc()}')
        return False, (error_fp, "")

    try:
        web_element = app.find_element(selector="//li[a[text()='ROOT']]", timeout=1)
        web_element.find_element(selector="//i[@class='jstree-icon "
                                          "jstree-ocl']", timeout=0.2).click(delay=0.2, scroll=True)
        for item in task.division:
            web_element.open_until_interactable(item)

        sleep(1)
        # Запуск

        fp = find_new_xlsx_file(path_downloaded,
                                sel="//button[@type='submit' "
                                    "and contains(text(), "
                                    "'Запустить')]")

        df_, error_msg = get_df(fp, 8)
        if df_ is not None:
            res_sk, res_tz = counters_df(df_)
            update_res_counters_task(session, InventoryTask, task.id, res_sk, res_tz)
            return True, (res_sk, res_tz)
        else:
            # return False, screen_fp to the screens with errors
            save_screen(task.id, 'gold_counters_report')
            log_error(inside_logger=gold_counters_logger, context=f'Не удалось выгрузить таблицу\n'
                                                                  f'if df_ is not None:')
            return False, (error_fp, "")

    except Exception as e:
        # Screen save
        save_screen(task.id, 'gold_counters_report')
        log_error(inside_logger=gold_counters_logger, context=f'Не удалось найти в списке элемент '
                                                              f'//li[a[text()="ROOT"]]'
                                                              f' по причине\n{e}')
        return False, (error_fp, "")

    finally:
        if app.driver.current_window_handle != main_window:
            app.driver.close()
            app.driver.switch_to.window(main_window)


def get_df(fp: Path, skip):
    if fp.is_file() and fp.suffix == '.xlsx':
        try:
            excel_data = pd.read_excel(fp, skiprows=skip)
            return excel_data, None
        except Exception as e:
            print(f"Ошибка при открытии файла: {e}")
            return None, e
    else:
        print("Файл Excel не найден.")
        return None, "gold_counters_report Файл Excel не найден."


def counters_df(df: pd.DataFrame):
    summ = df['Количество учетного остатка'].dropna().sum()
    sklad = summ * (1 - float(task.stock_share_tz_warehouse))
    tz = summ * float(task.stock_share_tz_warehouse)
    result_sklad = math.ceil(sklad / 6080 / float(task.days_until_inventory))
    result_tz = math.ceil(tz / 6080)
    return result_sklad, result_tz


# noinspection PyBroadException
def save_screen(id, func_name):
    path = screens_and_logs_path.joinpath(f"{id}\\{func_name}.png")
    try:
        app.driver.save_screenshot(filename=path)
        traceback.print_exc()
        return path
    except:
        logger.error("Не получилось сохранить скриншот")
    finally:
        return path


def log_error(inside_logger: logging.Logger, context: 'str'):
    inside_logger.error(context)


def init_inside_logger(func_name: 'str', id, level=logging.ERROR):
    path = screens_and_logs_path.joinpath(f"{id}\\{func_name}.txt")
    inside_logger = init_logger(
        f'{func_name}',
        level=level,
        tg_token=tg_token,
        chat_id=chat_id,
        log_path=path
    )
    return inside_logger


# noinspection PyUnusedLocal,PyShadowingNames,PyBroadException
def mailing_superiors(task: Result = None):
    """

    :type task: object
    """
    mailing_superiors_logger = init_inside_logger('mailing_superiors', task.id)
    error_fp = screens_and_logs_path.joinpath(f"{task.id}")

    def rename_file(original_path, new_name):
        """Переименовывает файл и возвращает новый путь к файлу."""
        dir_path, old_name = os.path.split(original_path)
        file_extension = os.path.splitext(old_name)[1]
        new_file_path = os.path.join(dir_path, new_name + file_extension)

        os.rename(original_path, new_file_path)

        return new_file_path

    try:
        info = get_message_info(session=session, model=InventoryTask, id=task.id)
        inv_number = info['res_projection1']
        table_data = info['res_table']
        res_sk, res_tz = info['res_counters_task']
        html_cont = smtp.generate_html(task, inv_number=inv_number, table_data=table_data, res_sk=res_sk, res_tz=res_tz)
        mails = smtp.find_receivers(task)

        new_path_pallets = rename_file(info['res_pallets'], f"Report_pallets_{task.branch}")
        new_path_closed = rename_file(info['res_closed'], f"Report_unclosed_docs_{task.branch}")

        if send_message(html=html_cont, receivers=mails, attachments=[new_path_pallets, new_path_closed]):
            delete_files([new_path_pallets, new_path_closed])
    except:
        log_error(mailing_superiors_logger,
                  f"ОШИБКА ВО ВРЕМЯ ОТПРАВКИ СООБЩЕНИЯ для {task.branch} \n traceback \n\n {traceback.format_exc()}")
        return False, error_fp
    return True, error_fp


# noinspection PyShadowingNames
def make_1_part(name, task):
    if name in globals():
        res, file_n_part = globals()[name](task)
        if res:
            # Задача успешно завершена
            logger.info(f"{name} is done")
            return True, file_n_part
        else:
            logger.error(f"Задача {str(name)} не выполнена")
            return False, file_n_part
    else:
        logger.info(f"Ошибка по причине, отсутствия нужной функции {name}")
        return False, f"Ошибка по причине, отсутствия нужной функции {name}. {name} IS NOT IN globals()"


# noinspection PyShadowingNames,PyBroadException
def execute_parts(task: Result = None):
    r"""
    func : db field mapping applied on the task row to
    :param task: sqlalchemy.engine.result.Result Результат query от sqlalchemy
    :return: bool
    """

    task_parts: dict = {
        "gold_projection": task.done_projection,
        "gold_report_pallets": task.done_pallets,
        "gold_closed_docs": task.done_closed_docs,
        "gold_counters_report": task.done_counters_task,
        "mailing_superiors": task.done_mailing_superiors,
    }

    task_done: dict = {
        "gold_projection": "done_projection",
        "gold_report_pallets": "done_pallets",
        "gold_closed_docs": "done_closed_docs",
        "gold_counters_report": "done_counters_task",
        "mailing_superiors": "done_mailing_superiors",
    }

    task_res: dict = {
        "gold_projection": "res_projection",
        "gold_report_pallets": "res_pallets",
        "gold_closed_docs": "res_closed_docs",
        "gold_counters_report": "res_counters_task",
        "mailing_superiors": "res_mailing_superiors",
    }

    mapping_russian = {
        "gold_projection": "Ошибка во время создания ракурса. Папка - done_projection",
        "gold_report_pallets": "Ошибка во время создания отчёта по паллетам. Папка - done_pallets",
        "gold_closed_docs": "Ошибка во время создания отчёта по незакрытым документам. Папка - done_closed_docs",
        "gold_counters_report": "Ошибка во время подсчёта требуемых работников. Папка - done_counters_task",
        "mailing_superiors": "Ошибка во время рассылки сообщений в филлиалы. Папка - done_mailing_superiors",
    }

    verif: dict = {}  # Ошибки хранятся в этой переменной
    try:
        for name in task_parts:
            logger.info(f"Задача под индексом {task.index},\n{name}")
            if not task_parts[name]:
                if name == "mailing_superiors":
                    statuses = get_current_task_statuses(session=session, model=InventoryTask, id=task.id)
                    if statuses['res_projection'] and statuses['res_pallets'] and statuses['res_closed_docs'] and \
                            statuses['res_counters_task']:
                        pass
                    else:
                        logger.error("Не выполнены все условия для рассылки")
                        return False, verif
                status, res_minitask = make_1_part(name, task)
                if status:
                    # Мини задача выполнена
                    complete_task(session=session, model=InventoryTask, id=task.id, done_task=task_done[name])
                    if res_minitask:
                        if name == "gold_counters_report":
                            update_res_counters_task(session=session, model=InventoryTask, id=task.id,
                                                     res_sk=res_minitask[0], res_tz=res_minitask[1])

                        elif name == "gold_projection":
                            res_number = res_minitask[0]
                            res_table = res_minitask[1]

                            column1_values = []
                            column2_values = []
                            for row in res_table:
                                column1_values.append(row[0])
                                column2_values.append(row[1])
                            update_res_projection(session, InventoryTask, task.id, res_number, column1_values,
                                                  column2_values)

                        else:
                            if isinstance(res_minitask, Path):
                                update_res_stringtype(session=session, model=InventoryTask, id=task.id,
                                                      value=res_minitask.as_posix(),
                                                      column_name=task_res[name])
                            else:
                                update_res_stringtype(session=session, model=InventoryTask, id=task.id,
                                                      value=res_minitask,
                                                      column_name=task_res[name])
                    logger.info(f"ЗАДАЧА {name} ВЫПОЛНЕНА УСПЕШНО")
                else:
                    # Встретилась / встретились ошибки
                    if res_minitask:
                        if isinstance(res_minitask, tuple):
                            verif[mapping_russian[name]] = res_minitask[0]
                        else:
                            verif[mapping_russian[name]] = res_minitask
                        if isinstance(verif[mapping_russian[name]], Path):
                            verif[mapping_russian[name]] = verif[mapping_russian[name]].as_posix()
                    if name == 'gold_projection':
                        return False, verif
                        pass
            else:
                # Do nothing. it's done
                pass

    except Exception:
        logger.error("Unexpected error! See below:\n")
        logger.error(traceback.format_exc())
        # noinspection PyUnboundLocalVariable
        verif[mapping_russian[name]] = "Exception"
        # if name == 'gold_projection':
        return False, verif

    if len(verif) > 0:
        return False, verif
    else:
        return True, verif


# True, verif:
# fp
# (res_sk, res_tz)
# None

# False, verif:
# "Exception"
# "(error_fp, "")"
# error_fp

# noinspection PyShadowingNames
def create_email_content(branches, success_tasks, fail_tasks):
    current_dir = os.path.dirname(__file__)
    project_dir = os.path.dirname(current_dir)
    template_path = os.path.join(project_dir, 'static', 'for_revisers.html')

    with open(template_path, 'r', encoding='utf-8') as file:
        html_template = file.read()

    # Формируем HTML для списка филиалов
    branches_html = ''.join(f'<li style="font-family: sans-serif;">{branch}</li>' for branch in branches)

    # Формируем HTML для задач
    tasks_html = '\n'.join(
        f'<tr style="border: 1px solid #000;"><td style="border: 1px solid #000;">{id_}</td><td style="border: 1px '
        f'solid #000;">{branch}</td>'
        f'<td style="color: red; border: 1px solid #000; font-weight: bold; font-size: larger;">ПРОВАЛ</td>'
        f'<td style="border: 1px solid #ddd; background-color: #ffeeee;">{"<br>".join(map(str, level))}</td></tr>'
        for id_, branch, level in fail_tasks
    )

    tasks_html += '\n'.join(
        f'<tr style="border: 1px solid #000;"><td style="border: 1px solid #000;">{id_}</td>'
        f'<td style="border: 1px solid #000;">{branch}</td>'
        f'<td style="color: green; border: 1px solid #000; font-weight: bold; font-size: larger;">УСПЕХ</td>'
        f'<td style="border: 1px solid #000;"></td></tr>'
        for id_, branch, _ in success_tasks
    )

    image_html = '<img src="magnum.jpg" alt="Magnum" style="display: block; margin-top: 20px;">'

    # Заменяем плейсхолдеры в шаблоне на соответствующий HTML
    html_content = html_template.format(branches=branches_html,
                                        tasks=tasks_html, image=image_html)
    return html_content


def send_errors(branches, success_tasks, fail_tasks):
    html_content = create_email_content(branches, success_tasks, fail_tasks)
    # TODO Изменить на реальные почты в env
    mails = [project_config_data['email1'], project_config_data['email2'], project_config_data['email3'],
             project_config_data['email4']]
    send_message("TEST", html_content, mails)


# Использование !!! GOLD
if __name__ == "__main__":
    failed_tasks = []
    success_tasks = []
    Session = sessionmaker(bind=engine)
    session = Session()
    app = Web()
    app.run()
    logger_main = init_inside_logger('main', 0)
    try:
        gold_password = gold_password
        gold_login = gold_login
        text = ("Deploy. Запуск на машине")
        init_main_logs(text)

        path_inventory_spreadsheet = Path(path_spreadsheet)

        get_tasks_flag = extract_info(path_inventory_spreadsheet)
        tasks = get_unfinished_tasks(session, InventoryTask)
        if not get_tasks_flag and len(tasks) == 0:
            raise ValueError("Пустой задачник на сегодня")

        # Идея такая. Все результаты записываются в этот лист

        for task in tasks:
            old_status = task.done_full_task
            change_task_status(session, InventoryTask, task.id, "Processing")
            res_per_task, result = execute_parts(task)
            """
                flag True if finished the task, False if encountered Exception
                result is dict which is empty if not any errors and not empty if any part is not done
                and has "Exception" if it sees Exception
            """
            if res_per_task:
                change_task_status(session, InventoryTask, task.id, "Success")
                # Выполнено всё успешно
                if len(result) == 0:
                    success_tasks.append((task.id, task.branch, ""))
                    save_to_excel(excel_path=path_inventory_spreadsheet, column="J", row=int(task.index) - 2,
                                  new_value="Success")

                    pass
                else:
                    # Нереальный кейс. Ожидается всегда пустой словарь при res_per_task == True
                    failed_tasks.append((task.id, task.branch, "Нереальный кейс. Задачи кажутся 'успешными'"))
                    save_to_excel(excel_path=path_inventory_spreadsheet, column="J", row=int(task.index) - 2,
                                  new_value="Fail")
                    raise ValueError(f"Невозможное значение словаря. Словарь накопил ошибки, которых быть не может. "
                                     f"Задача {task.id}")

            else:
                if old_status == "Repeatx2":
                    change_task_status(session, InventoryTask, task.id, "Repeatx3")
                elif old_status == "New":
                    change_task_status(session, InventoryTask, task.id, "Repeatx2")
                elif old_status == "Repeatx3":
                    change_task_status(session, InventoryTask, task.id, "Fail")
                failed_tasks.append((task.id, task.branch, result.keys()))
                save_to_excel(excel_path=path_inventory_spreadsheet, column="J", row=int(task.index) - 2,
                              new_value="Провал")
                if "Exception" in result.items():
                    # Send for me
                    tg_send("Exception", bot_token=tg_token, chat_id=chat_id)
                    # get_receivers_emails()
                    # send_fail_exception()
                    pass
                for bad in result:
                    if bad == "Exception":
                        tg_send(f"Exception in {task.id}", bot_token=tg_token, chat_id=chat_id)
                    else:
                        tg_send(f"\n\n=========\n\n{result[bad]} ВЫЗВАЛ ОШИБКУ", bot_token=tg_token,
                                chat_id=chat_id)
    except Exception as e:
        logger.error(f"==========================================================\n====================================="
                     f"=====================\n==========================================================\n"
                     f"НЕПРЕДВИДЕННАЯ ОШИБКА {e}\n ВО ВРЕМЯ \n{traceback.format_exc(limit=25)}")

    finally:
        # noinspection PyBroadException
        session.close()
        app.driver.quit()
        try:
            branches = [task[1] for task in success_tasks]
            branches.extend([task[1] for task in failed_tasks])
            if len(branches) > 0:
                send_errors(branches, success_tasks=success_tasks, fail_tasks=failed_tasks)
            for id in [id[0] for id in failed_tasks]:
                update_res_saving_errors(session=session, model=InventoryTask, id=id, value=True)
            for id in [id[0] for id in success_tasks]:
                update_res_saving_errors(session=session, model=InventoryTask, id=id, value=False)
        except Exception as e:
            logger.error(f"ОШИБКА ВО ВРЕМЯ ОТПРАВКИ ОШИБОК {e}")


if __name__ == "__main1__":
    try:
        text = "Test 'Проверка всего по отдельности'"
        init_main_logs(text)

        Session = sessionmaker(bind=engine)
        session = Session()

        task = get_unfinished_task(session=session, model=InventoryTask, task_ids=[1])

        app = Web()
        app.run()

        # gold_report_pallets(task)
        gold_report_pallets(task)
        # gold_closed_docs(task)
        # gold_counters_report(task)
        # mailing_superiors(task, branch_codes_path=project_branches_path)
        path_inventory_spreadsheet = Path.home().joinpath("Documents\\RPA\\Задача 1\\График_инвентаризаций_y2023.xlsx")
        # save_to_excel(excel_path=path_inventory_spreadsheet, column="J", row=int(task.index) - 2,
        #               new_value="Success")
    finally:
        session.close()
        app.driver.quit()

import datetime
import email
import json
import os
import time
import urllib.parse
from contextlib import suppress
from math import ceil
from time import sleep

import keyboard
import pandas as pd
import requests
import win32com.client as win32

from openpyxl import load_workbook, Workbook

from config import logger, engine_kwargs, robot_name, smtp_host, smtp_author, owa_username, owa_password, ecp_paths, ip_address

from sqlalchemy import create_engine, Column, Integer, String, DateTime, MetaData, Table, Date, Boolean, select, update
from sqlalchemy.orm import declarative_base, sessionmaker

from tools.net_use import net_use
from utils.fetching import fetching_unique_codes
from utils.parse_gtins import parse_all_gtins_to_out
from utils.wait_report import wait_report_to_download
from utils.website import ismet_auth, load_document_to_out, select_all_wares_to_dropout

Base = declarative_base()


class Table(Base):

    __tablename__ = f"{robot_name.replace('-', '_')}_dispatcher_whole"

    start_time = Column(DateTime, default=None)
    end_time = Column(DateTime, default=None)
    status = Column(String(128), default=None)
    error_message = Column(String(512), default=None)

    DATA_MATRIX_CODE = Column(String(256), primary_key=True)
    URL_INVOICE = Column(String(512))
    C_NAME_SOURCE_INVOICE = Column(String(512))
    C_NAME_SHOP = Column(String(512))
    DATE_INVOICE = Column(DateTime)

    @property
    def dict(self):
        m = self.__dict__
        return m


if __name__ == '__main__':

    if True:

        net_use(ecp_paths, owa_username, owa_password)

        check_ = False

        branches = list(os.listdir(ecp_paths))[::2]

        # if ip_address == '10.70.2.9':
        #     branches = list(os.listdir(ecp_paths))[1::2]
        # if ip_address == '10.70.2.11':
        #     branches = list(os.listdir(ecp_paths))[::2]
        # if ip_address == '172.20.1.24':
        #     branches = list(os.listdir(ecp_paths))[::-1]

        # if ip_address == '10.70.2.2':
        #     branches = list(os.listdir(ecp_paths))[100::]
        # if ip_address == '10.70.2.9':
        #     branches = list(os.listdir(ecp_paths))[1::2]
        # if ip_address == '10.70.2.11':
        #     branches = list(os.listdir(ecp_paths))[::2]
        # if ip_address == '172.20.1.24':
        #     branches = list(os.listdir(ecp_paths))[::-1]

        for folder in branches:

            # if folder in ['Торговый зал АСФ №1', 'Торговый зал АСФ №11', 'Торговый зал АСФ №12', 'Торговый зал АСФ №13', 'Торговый зал АСФ №16', 'Торговый зал АСФ №17', 'Торговый зал АСФ №18', 'Торговый зал АСФ №19', 'Торговый зал АСФ №2', 'Торговый зал АСФ №20', 'Торговый зал АСФ №21', 'Торговый зал АСФ №23', 'Торговый зал АСФ №24', 'Торговый зал АСФ №25', 'Торговый зал АСФ №26', 'Торговый зал АСФ №27', 'Торговый зал АСФ №28', 'Торговый зал АСФ №29', 'Торговый зал АСФ №3', 'Торговый зал АСФ №30', 'Торговый зал АСФ №31', 'Торговый зал АСФ №32', 'Торговый зал АСФ №33', 'Торговый зал АСФ №34', 'Торговый зал АСФ №37', 'Торговый зал АСФ №39', 'Торговый зал АСФ №4', 'Торговый зал АСФ №40', 'Торговый зал АСФ №41', 'Торговый зал АСФ №42', 'Торговый зал АСФ №45', 'Торговый зал АСФ №5', 'Торговый зал АСФ №50', 'Торговый зал АСФ №51', 'Торговый зал АСФ №53', 'Торговый зал АСФ №54', 'Торговый зал АСФ №56', 'Торговый зал АСФ №57', 'Торговый зал АСФ №58', 'Торговый зал АСФ №63', 'Торговый зал АСФ №7', 'Торговый зал АСФ №8', 'Торговый зал АСФ №81', 'Торговый зал АСФ №9', 'Торговый зал АФ №1', 'Торговый зал АФ №11', 'Торговый зал АФ №16', 'Торговый зал АФ №17', 'Торговый зал АФ №18', 'Торговый зал АФ №19', 'Торговый зал АФ №2', 'Торговый зал АФ №20', 'Торговый зал АФ №21', 'Торговый зал АФ №23', 'Торговый зал АФ №24', 'Торговый зал АФ №25', 'Торговый зал АФ №26', 'Торговый зал АФ №28', 'Торговый зал АФ №29', 'Торговый зал АФ №3', 'Торговый зал АФ №30', 'Торговый зал АФ №31', 'Торговый зал АФ №32', 'Торговый зал АФ №33', 'Торговый зал АФ №34', 'Торговый зал АФ №36', 'Торговый зал АФ №37', 'Торговый зал АФ №39', 'Торговый зал АФ №4', 'Торговый зал АФ №40', 'Торговый зал АФ №41', 'Торговый зал АФ №43', 'Торговый зал АФ №44', 'Торговый зал АФ №45', 'Торговый зал АФ №46', 'Торговый зал АФ №47', 'Торговый зал АФ №48', 'Торговый зал АФ №49', 'Торговый зал АФ №5', 'Торговый зал АФ №51', 'Торговый зал АФ №52', 'Торговый зал АФ №53', 'Торговый зал АФ №54', 'Торговый зал АФ №56', 'Торговый зал АФ №58', 'Торговый зал АФ №59', 'Торговый зал АФ №6', 'Торговый зал АФ №61', 'Торговый зал АФ №64', 'Торговый зал АФ №65', 'Торговый зал АФ №67', 'Торговый зал АФ №68', 'Торговый зал АФ №69', 'Торговый зал АФ №9', 'Торговый зал ЕКФ №1', 'Торговый зал КЗФ №1', 'Торговый зал КЗФ №2', 'Торговый зал КФ №4', 'Торговый зал КФ №6', 'Торговый зал КФ №7', 'Торговый зал ППФ №1', 'Торговый зал ППФ №10', 'Торговый зал ППФ №11', 'Торговый зал ППФ №13', 'Торговый зал ППФ №14', 'Торговый зал ППФ №16', 'Торговый зал ППФ №17', 'Торговый зал ППФ №2', 'Торговый зал ППФ №20', 'Торговый зал ППФ №21', 'Торговый зал ППФ №22', 'Торговый зал ППФ №3', 'Торговый зал ППФ №5', 'Торговый зал ППФ №6', 'Торговый зал ППФ №9', 'Торговый зал ТФ №1', 'Торговый зал ТФ №2', 'Торговый зал ТФ №3', 'Торговый зал УКФ №2', 'Торговый зал ФКС №2', 'Торговый зал ШФ №1', 'Торговый зал ШФ №10', 'Торговый зал ШФ №12', 'Торговый зал ШФ №14', 'Торговый зал ШФ №2', 'Торговый зал ШФ №4', 'Торговый зал ШФ №5', 'Торговый зал ШФ №6', 'Торговый зал ШФ №7']:
            #     continue
            if 'РЦ' in folder:
                continue

            if folder != 'Торговый зал АСФ №40':
                # check_ = True
                continue
            #
            # if not check_:
            #     continue

            logger.warning(f"STARTED {folder}")
            print(f"STARTED {folder}")

            Session = sessionmaker()

            engine = create_engine(
                'postgresql+psycopg2://{username}:{password}@{host}:{port}/{base}'.format(**engine_kwargs),
                connect_args={'options': '-csearch_path=robot'}
            )
            Base.metadata.create_all(bind=engine)
            Session.configure(bind=engine)
            session = Session()

            ecp_auth, ecp_sign = None, None
            folder_ = os.path.join(ecp_paths, folder)

            for file in os.listdir(folder_):

                if 'AUTH' in file:
                    ecp_auth = os.path.join(folder_, file)
                if 'GOST' in file:
                    ecp_sign = os.path.join(folder_, file)

            print(ecp_auth)
            web = ismet_auth(ecp_auth=ecp_auth, ecp_sign=ecp_sign)

            if web is None:
                continue

            web.get('https://goods.prod.markirovka.ismet.kz/cis/list')

            # sleep(1000)
            all_cookies = web.driver.get_cookies()
            cookies_dict = {}
            for cookie in all_cookies:
                cookies_dict[cookie['name']] = cookie['value']

            now = datetime.datetime.utcnow()
            formatted_date = time.strftime('%a, %d %b %Y %H:%M:%S GMT', now.timetuple())

            bearer_token = f"{cookies_dict.get('tokenPart1')}{cookies_dict.get('tokenPart2')}"
            print("bearer:", bearer_token)
            headers = {
                'Authorization': f'Bearer {bearer_token}',
                'Access-Control-Allow-Credentials': 'true',
                'Access-Control-Allow-Origin': 'https://goods.prod.markirovka.ismet.kz',
                'Access-Control-Expose-Headers': 'Authorization, Link, X-Total-Count',
                'Cache-Control': 'no-cache, no-store, max-age=0, must-revalidate',
                'Connection': 'keep-alive',
                'Content-Type': 'application/json;charset=UTF-8',
                'Date': formatted_date,
                'Expires': '0',
                'Pragma': 'no-cache',
                'Server': 'nginx',
                'Vary': 'Access-Control-Request-Headers, Access-Control-Request-Method, Origin',
                'X-Content-Type-Options': 'nosniff',
                'X-Frame-Options': 'DENY',
                'X-Xss-Protection': '1; mode=block'
            }

            total_pages = 3
            pageSize = 1
            page = 0

            for year in [2022, 2023]:
                for month in range(1, 13):

                    total_pages = 1
                    page = 0

                    if month == 12:
                        last_day = (datetime.date(year + 1, 1, 1) - datetime.timedelta(days=1)).day
                    else:
                        last_day = (datetime.date(year, month + 1, 1) - datetime.timedelta(days=1)).day

                    month_ = f"0{month}" if month < 10 else str(month)
                    last_day = f"0{last_day}" if last_day < 10 else str(last_day)
                    print(f"Started year: {year}, month: {month_}, last day: {last_day}")

                    # * ----- Get total els in the month -----
                    data = {
                        "pageSize": pageSize,
                        "currentPage": page,
                        "filters": [
                            {"operator": "=", "column": "owner", "filterTerm": "120641002344"},
                            {"operator": "=", "column": "status", "filterTerm": 2},
                            {"operator": ">=", "column": "emissionDate", "filterTerm": f"{year}-{month_}-01T00:00:00"},
                            {"operator": "<=", "column": "emissionDate", "filterTerm": f"{year}-{month_}-{last_day}T23:59:59"}
                        ],
                        "groupBy": [],
                        "sorts": [{"column": "status", "direction": "ASC"}]
                    }
                    for i in range(3):
                        try:
                            # print(1)
                            r = requests.post('https://goods.prod.markirovka.ismet.kz/api/km-grid/getPageData', data=json.dumps(data), cookies=cookies_dict, headers=headers,
                                              verify=False, timeout=180)
                            quantity = r.text
                            json_ = json.loads(quantity)
                            total_pages = ceil(int(json_["count"]) / pageSize)
                            break
                        except Exception as err1:
                            print("EROR")
                            logger.warning(f"ERRORKIN: {err1}")
                            if 'count' in str(err1):
                                break
                    # * --------------------------------------

                    pageSizeCalc = 1

                    for mnozhitel in range(70, 1, -1):
                        if total_pages % mnozhitel == 0:
                            pageSizeCalc = mnozhitel
                            break

                    logger.warning(f"MNOZHITEL: {pageSizeCalc} | TOTAL PAGES: {total_pages}")

                    print(page, total_pages % pageSizeCalc, page <= total_pages % pageSizeCalc)
                    total_iters = total_pages / pageSizeCalc
                    # while page <= total_pages % pageSizeCalc + 1:
                    for i in range(int(total_pages / pageSizeCalc)):
                        if page % 1 == 0:
                            logger.warning(f'PAGE: {page}')
                        data = {
                            "pageSize": pageSizeCalc,
                            "currentPage": page,
                            "filters": [
                                {"operator": "=", "column": "owner", "filterTerm": "120641002344"},
                                {"operator": "=", "column": "status", "filterTerm": 2},
                                {"operator": ">=", "column": "emissionDate", "filterTerm": f"{year}-{month_}-01T00:00:00"},
                                {"operator": "<=", "column": "emissionDate", "filterTerm": f"{year}-{month_}-{last_day}T23:59:59"}
                            ],
                            "groupBy": [],
                            "sorts": [{"column": "status", "direction": "ASC"}]
                        }
                        # print(f"requests.post('https://goods.prod.markirovka.ismet.kz/api/km-grid/getPageData', data=json.dumps(data), cookies={cookies_dict}, headers=headers, verify=False, timeout=180)")

                        # print(cookies_dict)
                        # print(headers)
                        for i in range(3):
                            try:
                                # print(1)
                                r = requests.post('https://goods.prod.markirovka.ismet.kz/api/km-grid/getPageData', data=json.dumps(data), cookies=cookies_dict, headers=headers,
                                                  verify=False, timeout=180)
                                quantity = r.text
                                # print(quantity)
                                # logger.warning('------------------------------')
                                # print(r)
                                json_ = json.loads(quantity)
                                # total_pages = ceil(int(json_["count"]) / pageSize)
                                # print('------------------------------')
                                # print(total_pages)
                                # print(json_)
                                # print(json_["kizes"])
                                # print('===')
                                # for indd, product in enumerate(json_["kizes"]):
                                #     print(indd, product['cis'])
                                # print('------------------------------')
                                # print(2)
                                for product in json_["kizes"]:

                                    select_query = (
                                        session.query(Table)
                                            .filter(Table.DATA_MATRIX_CODE == product['cis'])
                                            .all()
                                    )
                                    if len(select_query) != 0:
                                        continue
                                    # print(product['cis'], product['producerId']['name'], f"https://goods.prod.markirovka.ismet.kz/{urllib.parse.quote(product['cis'])}")
                                    try:
                                        session.add(Table(
                                            start_time=datetime.datetime.now(),
                                            status='new',
                                            URL_INVOICE=f"https://goods.prod.markirovka.ismet.kz/cis/list/{urllib.parse.quote(product['cis'])}",
                                            DATA_MATRIX_CODE=product['cis'],
                                            C_NAME_SOURCE_INVOICE=product['producerId']['name'],
                                            C_NAME_SHOP=folder,
                                            DATE_INVOICE=datetime.datetime.fromtimestamp(product['emissionDate'] / 1000.0).strftime('%d.%m.%Y')
                                        ))
                                    except Exception as err:
                                        print("EROR1")
                                        logger.warning(f"ERRORCHIK: {err}")
                                # sleep(10000)
                                # print(3)
                                break
                            except Exception as err1:
                                print("EROR")
                                logger.warning(f"ERRORKIN: {err1}")
                                if 'count' in str(err1):
                                    break
                        page += 1
                    # sleep(1000)
            logger.warning(f'FINISHED {folder}')
            # sleep(10000)

            session.commit()
            session.close()

            web.quit()

            # sleep(10000)

            # web.get('https://goods.prod.markirovka.ismet.kz/cis/list')
            # sleep(5)
            #
            # web.find_element("//div[text()='Статус кода']/../following-sibling::button").click()
            # sleep(0.05)
            # for i in range(10):
            #     keyboard.press_and_release("ctrl+-")
            #
            # sleep(0.1)
            #
            # web.find_element("//div[text()='Статус кода']/../following-sibling::button").click()
            # sleep(0.1)
            # web.find_element("//div[text()='Статус']/following-sibling::input").click()
            # sleep(0.05)
            # web.find_element("//span[text()='В обороте']").click()
            # sleep(0.05)
            # web.find_element("//div[text()='Статус']/../../../following-sibling::button[text()='Применить']").click()
            # sleep(0.05)

            #     page_num = 1
            #
            #     while True:
            #
            #         logger.warning(f"Current page: {page_num}")
            #
            #         sleep(1)
            #         links = web.find_elements("//a[@class='sc-hGoxap jcKonU']", timeout=10)
            #         gtin_codes = web.find_elements("//a[@class='sc-hGoxap jcKonU']/div", timeout=10)
            #         act_dates = web.find_elements("(//div[@class='sc-TFwJa gEcuGv'])[position() >= 0 and (position() - 2) mod 10 = 0]", timeout=10)
            #         # act_types = web.find_elements("(//div[@class='sc-jhaWeW dOwKeN'])[position() mod 2 = 1]", timeout=10)
            #         suppliers = web.find_elements("(//div[@class='sc-TFwJa gEcuGv'])[position() >= 0 and (position() - 8) mod 10 = 0]", timeout=10)
            #         # is_succeeds = web.find_elements("(//div[@class='sc-jhaWeW dOwKeN'])[position() mod 2 = 0]", timeout=10)
            #         sleep(.1)
            #         for ind in range(len(links)):
            #
            #             select_query = (
            #                 session.query(Table)
            #                     .filter(Table.URL_INVOICE == links[ind].get_attr('href'))
            #                     .all()
            #             )
            #
            #             # print('len:', len(select_query), links[ind].get_attr('href'))
            #
            #             if len(select_query) != 0:
            #                 continue
            #
            #             try:
            #                 session.add(Table(
            #                     start_time=datetime.datetime.now(),
            #                     status='new',
            #                     ID_INVOICE=links[ind].get_attr('text'),
            #                     URL_INVOICE=links[ind].get_attr('href'),
            #                     DATA_MATRIX_CODE=gtin_codes[ind].get_attr('text'),
            #                     C_NAME_SOURCE_INVOICE=suppliers[ind].get_attr('text'),
            #                     C_NAME_SHOP=folder,
            #                     DATE_INVOICE=act_dates[ind].get_attr('text')
            #                 ))
            #             except:
            #                 logger.warning("ERROR???")
            #                 logger.warning(links)
            #                 logger.warning(links[ind])
            #
            #             session.commit()
            #
            #         try:
            #             web.execute_script_click_xpath("//*[@id='root']/div/div[1]/div[2]/div/div[2]/div/div/div[3]/div/div[3]/div/div[3]/div")
            #             page_num += 1
            #         except:
            #             break
            #         # try:
            #         #     current_page = int(web.find_element("//div[@class='PageWrapper__ActiveWrapper-iSeyvZ jbDrJd']", timeout=5).get_attr('text'))
            #         # except:
            #         #     break
            #         # available_page = None
            #         # for pages in web.find_elements("//div[@class='PageWrapper__Wrapper-japepR faAVRk']"):
            #         #     with suppress(Exception):
            #         #         if int(pages.get_attr('text')) - 1 == current_page:
            #         #             available_page = int(pages.get_attr('text'))
            #         #             break
            #         #
            #         # if available_page is None:
            #         #     break
            #         #
            #         # web.find_element(f"//div[@class='PageWrapper__Wrapper-japepR faAVRk']/div[text()='{available_page}']/..").click()
            #
            #         # try:
            #         #     old_url = web.driver.current_url
            #         #     web.find_element("//div[@class='RightPaginationTextWrap-jaPepQ gJUHHY']", timeout=3).click()
            #         #     sleep(1)
            #         #     new_url = web.driver.current_url
            #         #
            #         #     if old_url == new_url:
            #         #         break
            #         #
            #         # except:
            #         #     break
            #
            #     web.quit()
            # session.close()


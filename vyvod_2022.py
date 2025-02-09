import datetime
import os
import shutil
import time
from contextlib import suppress
from pathlib import Path
from time import sleep
import pandas as pd
import win32com.client as win32

from openpyxl import load_workbook, Workbook

from config import logger, engine_kwargs, robot_name, smtp_host, smtp_author, owa_username, owa_password, ecp_paths, ip_address, working_path, download_path

from sqlalchemy import create_engine, Column, Integer, String, DateTime, MetaData, Table, Date, Boolean, select, update
from sqlalchemy.orm import declarative_base, sessionmaker

from tools.net_use import net_use
from utils.fetching import fetching_unique_codes
from utils.parse_gtins import parse_all_gtins_to_out
from utils.wait_report import wait_report_to_download
from utils.website import ismet_auth, load_document_to_out, select_all_wares_to_dropout

Base = declarative_base()


class Table(Base):

    __tablename__ = "robot_ismet_vyvod_iz_oborota_performer_prod"  # robot_name.replace('-', '_')

    start_time = Column(DateTime, default=None)
    end_time = Column(DateTime, default=None)
    status = Column(String(128), default=None)
    error_message = Column(String(512), default=None)

    DATA_MATRIX_CODE = Column(String(256), primary_key=True)
    GTIN_CODE = Column(String(256))
    ID_INVOICE = Column(String(256))
    NUMBER_INVOICE = Column(String(256))
    URL_INVOICE = Column(String(512))
    NEW_URL_INVOICE = Column(String(512))
    FILE_SAVED_PATH = Column(String(512))
    C_NAME_SOURCE_INVOICE = Column(String(512))
    C_NAME_SHOP = Column(String(512))
    DATE_INVOICE = Column(DateTime)
    NAME_WARES = Column(String(512))

    @property
    def dict(self):
        m = self.__dict__
        return m


class Table2022(Base):

    __tablename__ = f"{robot_name.replace('-', '_')}_dispatcher_prod"

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


def fetching_unique_codes_2022(branch: str, update_to_success=False):

    Session = sessionmaker()

    engine = create_engine(
        'postgresql+psycopg2://{username}:{password}@{host}:{port}/{base}'.format(**engine_kwargs),
        connect_args={'options': '-csearch_path=robot'}
    )
    Base.metadata.create_all(bind=engine)
    Session.configure(bind=engine)
    session2022 = Session()

    select_query1 = (
        session2022.query(Table2022)
            .filter(Table2022.C_NAME_SHOP == branch)
            .filter(Table2022.status == 'new')
            .all()
    )

    # * Fetching all number invoices from the db

    if not update_to_success:
        vals = []
        for ind, row in enumerate(select_query1):
            vals.append(f"{row.DATA_MATRIX_CODE} |-|-| {row.C_NAME_SOURCE_INVOICE} |-|-| {row.DATE_INVOICE.strftime('%d.%m.%Y')}")

        session2022.close()

        return vals

    if update_to_success:
        for record in select_query1:
            record.status = 'success'

        session2022.commit()
        session2022.close()

        return 0


def vyvodbek():

    Session = sessionmaker()

    engine = create_engine(
        'postgresql+psycopg2://{username}:{password}@{host}:{port}/{base}'.format(**engine_kwargs),
        connect_args={'options': '-csearch_path=robot'}
    )
    Base.metadata.create_all(bind=engine)
    Session.configure(bind=engine)
    session = Session()

    net_use(ecp_paths, owa_username, owa_password)

    check_ = False

    branches = list(os.listdir(ecp_paths))[::]

    # if ip_address == '10.70.2.9':
    #     branches = list(os.listdir(ecp_paths))[1::2]
    # if ip_address == '10.70.2.11':
    #     branches = list(os.listdir(ecp_paths))[::2]
    # if ip_address == '172.20.1.24':
    #     branches = list(os.listdir(ecp_paths))[::-1]

    # part_length = len(list(os.listdir(ecp_paths))) // 5

    # if ip_address == '10.70.2.6':
    #     branches = list(os.listdir(ecp_paths))[::2]
    # if ip_address == '10.70.2.52':
    #     branches = list(os.listdir(ecp_paths))[1::2]
    # if ip_address == '10.70.2.5':
    #     branches = list(os.listdir(ecp_paths))[::-2]
    # if ip_address == '10.70.2.51':
    #     branches = list(os.listdir(ecp_paths))[-2::-2]
    # if ip_address == '172.20.1.24':
    #     branches = list(os.listdir(ecp_paths))[::-1]

    for folder in branches:

        # if ip_address == '10.70.2.11':
        #     if folder == 'Торговый зал АСФ №55':
        #         check_ = True
        #         # continue
        #
        #     if not check_:
        #         continue

        # if folder not in ['Торговый зал АСФ №1', 'Торговый зал АСФ №18', 'Торговый зал АСФ №31', 'Торговый зал АСФ №50', 'Торговый зал АСФ №74', 'Торговый зал АФ №1', 'Торговый зал АФ №11', 'Торговый зал АФ №12', 'Торговый зал АФ №2', 'Торговый зал АФ №24', 'Торговый зал АФ №29', 'Торговый зал АФ №3', 'Торговый зал АФ №32', 'Торговый зал АФ №35', 'Торговый зал АФ №4', 'Торговый зал АФ №40', 'Торговый зал АФ №61', 'Торговый зал АФ №69', 'Торговый зал ТФ №1', 'Торговый зал ШФ №1']: # , 'Торговый зал ЕКФ №1'
        #     continue

        # if ip_address == '172.20.1.24':
        #     if folder == 'Торговый зал АСФ №41':
        #         check_ = True
        #         # continue
        #
        #     if not check_:
        #         continue

        # if folder == 'Торговый зал АСФ №40':
        #     # check_ = True
        #     continue

        logger.warning(f"Started {folder}")

        ecp_auth, ecp_sign = None, None
        folder_ = os.path.join(ecp_paths, folder)
        for file in os.listdir(folder_):

            if 'AUTH' in file:
                ecp_auth = os.path.join(folder_, file)
            if 'GOST' in file:
                ecp_sign = os.path.join(folder_, file)

        print(folder)
        vals: list = fetching_unique_codes_2022(branch=folder, update_to_success=False)
        print('finished fetching')

        # for val in vals:
        #     print(val)
        print(len(vals))
        # sleep(10000)
        # print(vals)
        session.close()

        # sleep(10000)
        # print(vals)
        start, end = 0, 1000
        while start < len(vals):

            logger.warning(f"Current slice: {start} - {end} | {len(vals)}")

            val = vals[start:end]
            start = end
            end += 1000

            with suppress(Exception):
                os.system("taskkill /im excel.exe")

            book = Workbook()
            sheet = book.active

            last_row = 1
            added_any_row = False

            time_overall_start = time.time()
            avg_time = 0
            cc = 0

            for val_ in val:
                # print(val)
                # print(val_)
                data_matrix_code = str(val_.split('|-|-|')[0]).strip()
                name_source = val_.split('|-|-|')[1]
                date_invoice = val_.split('|-|-|')[2]

                avg_time_start = time.time()

                select_query = (
                    session.query(Table)
                        .filter(Table.DATA_MATRIX_CODE == data_matrix_code)
                        .all()
                )

                # print('len:', len(select_query))

                if len(select_query) != 0:
                    # logger.info('----- ALREADY IN DB | NEXT -----')
                    continue

                sheet[f'A{last_row}'].value = str(data_matrix_code).strip()
                last_row += 1

                session.add(Table(
                    start_time=datetime.datetime.now(),
                    status='new',
                    DATA_MATRIX_CODE=data_matrix_code,
                    C_NAME_SOURCE_INVOICE=name_source,
                    C_NAME_SHOP=folder,
                    DATE_INVOICE=date_invoice,
                ))

                avg_time_end = time.time()

                avg_time += float(avg_time_end - avg_time_start)
                cc += 1

                added_any_row = True

                if not added_any_row:
                    # logger.info('----- ALREADY IN DB | NEXT -----')
                    continue

            time_overall_end = time.time()

            if cc == 0:
                cc = 1

            print("Total time to add new rows in db:", time_overall_end - time_overall_start)
            print("Average time to add new rows in db:", avg_time, avg_time / cc)

            if last_row == 1:
                continue

            session.commit()

            error_msg = None
            new_url = None

            report_path = ''

            sheet[f'A{last_row}'].value = ''

            file_path = os.path.join(download_path, f'{folder}.xlsx')

            book.save(file_path)

            book.close()

            saved = False

            web = ismet_auth(ecp_auth=ecp_auth, ecp_sign=ecp_sign)

            if web is None:
                break

            document_api_url = 'https://goods.prod.markirovka.ismet.kz/api/v3/facade/marked_products/search-by-xls-cis-list'

            for _ in range(1000):
                try:

                    with suppress(Exception):
                        os.system("taskkill /im excel.exe")

                    excel = win32.gencache.EnsureDispatch('Excel.Application')
                    excel.Visible = False
                    excel.DisplayAlerts = False

                    wb = excel.Workbooks.Open(file_path)
                    wb.Save()
                    wb.Close()

                    print('saved')

                    load_document_to_out(web=web, filepath=file_path, year=2023, month=12, day=31, url=document_api_url)

                    select_all_wares_to_dropout(web=web, ecp_sign=ecp_sign)

                    # new_url = web.driver.current_url
                    #
                    # print(new_url)
                    #
                    # report_path = wait_report_to_download(branch=folder, date_=f"_whole_2022_{start % 1500}")
                    saved = True

                    break

                except Exception as error:
                    error_msg = str(error)[:500]
                    logger.info(f"ERROR OCCURED: {error}")
                    web.driver.refresh()

            sleep(0)

            update_values_list = []

            if saved:
                logger.warning(f'{datetime.datetime.now()} | Marking as Succeed')
                print(f'{datetime.datetime.now()} | Marking as Succeed')
                for val_ in val:
                    data_matrix_code = str(val_.split('|-|-|')[0]).strip()
                    stmt = update(Table).where(Table.DATA_MATRIX_CODE == data_matrix_code).values(
                        status='success',
                        end_time=datetime.datetime.now(),
                        error_message='',
                        NEW_URL_INVOICE=new_url,
                        FILE_SAVED_PATH=report_path
                    )
                    session.execute(stmt)

            else:
                logger.warning(f'{datetime.datetime.now()} | Marking as Failed')
                print(f'{datetime.datetime.now()} | Marking as Failed')
                for val_ in val:
                    data_matrix_code = str(val_.split('|-|-|')[0]).strip()
                    stmt = update(Table).where(
                        Table.DATA_MATRIX_CODE == data_matrix_code
                    ).values(
                        status='failed',
                        end_time=datetime.datetime.now(),
                        error_message=error_msg,
                        NEW_URL_INVOICE=new_url,
                        FILE_SAVED_PATH=report_path
                    )
                    session.execute(stmt)

            logger.warning(f'{datetime.datetime.now()} | Finished marking')
            session.commit()

            # Path(file_path).unlink()

            web.quit()

        # fetching_unique_codes_2022(branch=folder, update_to_success=True)

        logger.info('----- NEXT -----')

    session.close()


if __name__ == '__main__':
    vyvodbek()
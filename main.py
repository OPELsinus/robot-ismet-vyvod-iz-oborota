import datetime
import os
from time import sleep
import pandas as pd
import win32com.client as win32

from openpyxl import load_workbook, Workbook

from config import logger, engine_kwargs, robot_name, smtp_host, smtp_author, owa_username, owa_password, ecp_paths

from sqlalchemy import create_engine, Column, Integer, String, DateTime, MetaData, Table, Date, Boolean, select, update
from sqlalchemy.orm import declarative_base, sessionmaker

from tools.net_use import net_use
from utils.fetching import fetching_unique_codes
from utils.parse_gtins import parse_all_gtins_to_out
from utils.wait_report import wait_report_to_download
from utils.website import ismet_auth, load_document_to_out, select_all_wares_to_dropout

Base = declarative_base()


class Table(Base):

    __tablename__ = robot_name.replace('-', '_')

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


if __name__ == '__main__':

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

    for folder in os.listdir(ecp_paths):

        if folder != 'Торговый зал АСФ №1':
            check_ = True
            continue

        # if not check_:
        #     continue

        ecp_auth, ecp_sign = None, None
        folder_ = os.path.join(ecp_paths, folder)
        for file in os.listdir(folder_):

            if 'AUTH' in file:
                ecp_auth = os.path.join(folder_, file)
            if 'GOST' in file:
                ecp_sign = os.path.join(folder_, file)

        urls: dict = fetching_unique_codes(branch=folder)

        print(folder, len(urls))
        if len(urls) == 0:
            continue
        # continue
        print(len(urls))
        for val, key in urls.items():

            print(val, key)

        web = ismet_auth(ecp_auth=ecp_auth, ecp_sign=ecp_sign)

        for url in urls:

            print(f"----- {url} -----")
            all_goods: dict = parse_all_gtins_to_out(web=web, url=url)

            book = Workbook()
            sheet = book.active

            last_row = 1

            added_any_row = False

            for key, val in all_goods.items():
                print(key, val)
                for ind1 in range(len(val[0])):

                    select_query = (
                        session.query(Table)
                            .filter(Table.DATA_MATRIX_CODE == val[0][ind1])
                            .all()
                    )

                    # print('len:', len(select_query))

                    if len(select_query) != 0:
                        continue

                    sheet[f'A{last_row}'].value = str(val[0][ind1]).strip()
                    last_row += 1

                    session.add(Table(
                        start_time=datetime.datetime.now(),
                        status='new',
                        DATA_MATRIX_CODE=val[0][ind1],
                        GTIN_CODE=val[1][ind1],
                        ID_INVOICE=urls.get(key)[0],
                        URL_INVOICE=key,
                        NEW_URL_INVOICE='',
                        FILE_SAVED_PATH='',
                        NUMBER_INVOICE=urls.get(key)[1],
                        C_NAME_SOURCE_INVOICE=urls.get(key)[2],
                        C_NAME_SHOP=urls.get(key)[3],
                        DATE_INVOICE=urls.get(key)[4],
                        NAME_WARES=val[2][ind1]
                    ))

                    added_any_row = True

            if not added_any_row:
                logger.info('----- ALREADY IN DB | NEXT -----')
                continue

            session.commit()

            error_msg = None
            new_url = None

            report_path = ''

            try:
                sheet[f'A{last_row}'].value = ''

                file_path = fr'C:\Users\Abdykarim.D\PycharmProjects\robot-ismet-vyvod-iz-oborota\{folder}.xlsx'

                book.save(file_path)

                book.close()

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False
                excel.DisplayAlerts = False

                wb = excel.Workbooks.Open(file_path)
                wb.Save()
                wb.Close()

                invoice_date: datetime = urls.get(url)[4]
                print(invoice_date, invoice_date.strftime('%d_%m_%Y'))
                print(invoice_date.year, invoice_date.month, invoice_date.day)
                load_document_to_out(web=web, filepath=file_path, year=invoice_date.year, month=invoice_date.month, day=invoice_date.day)

                select_all_wares_to_dropout(web=web, ecp_sign=ecp_sign)

                new_url = web.driver.current_url

                print(new_url)

                report_path = wait_report_to_download(branch=folder, date_=invoice_date.strftime('%d_%m_%Y'))

            except Exception as error:
                error_msg = str(error)[:500]
                logger.info(f"ERROR OCCURED: {error}")

            sleep(0)
            for key, val in all_goods.items():
                for ind1 in range(len(val[0])):

                    stmt = update(Table).where(
                        Table.DATA_MATRIX_CODE == val[0][ind1]
                    ).values(
                        status='success',
                        end_time=datetime.datetime.now(),
                        error_message=error_msg,
                        NEW_URL_INVOICE=new_url,
                        FILE_SAVED_PATH=report_path
                    )
                    session.execute(stmt)

            session.commit()

            logger.info('----- NEXT -----')

        web.quit()

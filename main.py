import datetime
import os
from time import sleep

from openpyxl import load_workbook

from config import logger, engine_kwargs, robot_name, smtp_host, smtp_author, owa_username, owa_password, ecp_paths

from sqlalchemy import create_engine, Column, Integer, String, DateTime, MetaData, Table, Date, Boolean, select
from sqlalchemy.orm import declarative_base, sessionmaker

from tools.net_use import net_use
from utils.fetching import fetching_unique_codes
from utils.parse_gtins import parse_all_gtins_to_out
from utils.website import ismet_auth, load_document_to_out

Base = declarative_base()


class Table(Base):

    __tablename__ = robot_name.replace('-', '_')

    date_created = Column(DateTime, default=None)
    invoice_date = Column(DateTime, default=None)
    id_invoice = Column(String(512), primary_key=True)
    reason_invoice = Column(String(512), default=None)
    store_name = Column(String(512), default=None)
    supplier_name = Column(String(512), default=None)

    status = Column(String(16), default=None)

    @property
    def dict(self):
        m = self.__dict__.copy()
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

    # for key, val in invoices.items():
    #     # print(key, val)
    #     # print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), val[3])
    #     session.add(Table(
    #         date_created=datetime.datetime.now(),
    #         invoice_date=val[3],
    #         id_invoice=key,
    #         reason_invoice=val[0],
    #         store_name=val[1],
    #         supplier_name=val[2],
    #         status='new'
    #     ))
    # session.commit()

    net_use(ecp_paths, owa_username, owa_password)

    for folder in os.listdir(ecp_paths):

        if folder != 'Торговый зал АФ №68':
            continue

        ecp_auth, ecp_sign = None, None
        folder_ = os.path.join(ecp_paths, folder)
        for file in os.listdir(folder_):

            if 'AUTH' in file:
                ecp_auth = os.path.join(folder_, file)
            if 'GOST' in file:
                ecp_sign = os.path.join(folder_, file)

        urls = fetching_unique_codes(branch=folder)

        web = ismet_auth(ecp_auth=ecp_auth, ecp_sign=ecp_sign)

        parse_all_gtins_to_out(web=web, urls=urls)

        load_document_to_out(web=web, year=2023, month=10, day=13)


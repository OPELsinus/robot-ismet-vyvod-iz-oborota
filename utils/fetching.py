import datetime

from sqlalchemy import Column, Integer, String, BigInteger, DateTime, create_engine, or_
from sqlalchemy.orm import sessionmaker, declarative_base

from config import global_env_data

import numpy as np

Base = declarative_base()


class IsmetTable(Base):

    __tablename__ = "parse_all"
    id = Column(Integer, primary_key=True)
    edit_time = Column(DateTime, default=None)
    status = Column(String(128), default=None)
    error_message = Column(String(128), default=None)

    ID_INVOICE = Column(String(128))
    NUMBER_INVOICE = Column(String(128))
    URL_INVOICE = Column(String(128))
    C_NAME_SOURCE_INVOICE = Column(String(128))
    C_NAME_SHOP = Column(String(128))
    DATE_INVOICE = Column(DateTime)
    BAR_CODE_WARES = Column(BigInteger)
    NAME_WARES = Column(String(128))
    QUANTITY = Column(Integer)
    APPROVE_FLAG = Column(Integer, default=None)

    @property
    def dict(self):
        m = self.__dict__
        return m


def fetching_unique_codes(branch: str):

    # * Creating connection to the ismet table

    Session_ismet = sessionmaker()

    engine_kwargs_ismet = {
        'username': global_env_data['postgre_db_username'],
        'password': global_env_data['postgre_db_password'],
        'host': global_env_data['postgre_ip'],
        'port': global_env_data['postgre_port'],
        'base': 'ismet'
    }

    engine = create_engine(
        'postgresql+psycopg2://{username}:{password}@{host}:{port}/{base}'.format(**engine_kwargs_ismet),
        connect_args={'options': '-csearch_path=public'}
    )

    Base.metadata.create_all(bind=engine)
    Session_ismet.configure(bind=engine)
    session_ismet = Session_ismet()

    select_query = (
        session_ismet.query(IsmetTable)
            .filter(IsmetTable.C_NAME_SHOP == branch)
            .filter(IsmetTable.DATE_INVOICE >= datetime.date(2023, 1, 1))
            .filter(IsmetTable.DATE_INVOICE <= datetime.date(2023, 5, 31))
            .filter(IsmetTable.APPROVE_FLAG == 1)
            .filter(IsmetTable.status == 'Success')
            .filter(or_(IsmetTable.NUMBER_INVOICE.is_(None), IsmetTable.NUMBER_INVOICE.notlike('!%')))
            .all()
    )

    # * Fetching all number invoices from the db

    id_invoice, num_invoice, c_name, name_shop, date_invoice = [], [], [], [], []
    vals = dict()
    for ind, row in enumerate(select_query):
        vals.update({row.URL_INVOICE: [row.ID_INVOICE, row.NUMBER_INVOICE, row.C_NAME_SOURCE_INVOICE, row.C_NAME_SHOP, row.DATE_INVOICE]})
        # print(ind, [row.ID_INVOICE, row.NUMBER_INVOICE, row.C_NAME_SOURCE_INVOICE, row.C_NAME_SHOP, row.DATE_INVOICE])
        # values.append([row.ID_INVOICE, row.NUMBER_INVOICE, row.C_NAME_SOURCE_INVOICE, row.C_NAME_SHOP, row.DATE_INVOICE])
        # urls.append(row.URL_INVOICE)
        # id_invoice.append(row.ID_INVOICE)
        # num_invoice.append(row.NUMBER_INVOICE)
        # c_name.append(row.C_NAME_SOURCE_INVOICE)
        # name_shop.append(row.C_NAME_SHOP)
        # date_invoice.append(row.DATE_INVOICE)

    session_ismet.close()

    return vals # np.unique(urls), np.unique(id_invoice), np.unique(num_invoice), np.unique(c_name), np.unique(name_shop), np.unique(date_invoice)

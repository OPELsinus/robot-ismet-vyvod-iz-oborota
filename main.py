import datetime
from time import sleep

from openpyxl import load_workbook

from config import logger, engine_kwargs, robot_name, smtp_host, smtp_author

from sqlalchemy import create_engine, Column, Integer, String, DateTime, MetaData, Table, Date, Boolean, select
from sqlalchemy.orm import declarative_base, sessionmaker

from utils.website import open_ismet_document

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

    open_ismet_document(r'\\vault.magnum.local\common\Stuff\_06_Бухгалтерия\! Актуальные ЭЦП\Торговый зал АФ №68\AUTH_RSA256_e6a89349ddb123800cd7e20da7ab61ee51b86c10.p12')



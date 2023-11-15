from sqlalchemy import Column, Integer, String, BigInteger, DateTime, create_engine
from sqlalchemy.orm import sessionmaker, declarative_base

from config import global_env_data

Base = declarative_base()


class IsmetTable(Base):

    __tablename__ = "parse_all"
    id = Column(Integer, primary_key=True)
    # add_time = Column(DateTime, default=None)
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
            .filter(IsmetTable.NUMBER_INVOICE.notlike('!%'))
            .all()
    )

    # * Fetching all number invoices from the db

    urls = []

    for ind, row in enumerate(select_query):
        print(ind, row.URL_INVOICE, row.C_NAME_SHOP)
        urls.append(row.URL_INVOICE)

    session_ismet.close()

    return urls

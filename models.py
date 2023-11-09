from sqlalchemy import create_engine, Column, Integer, String, Numeric, Date, Boolean, Table, column, MetaData
from sqlalchemy.orm import sessionmaker, declarative_base
from sqlalchemy.dialects.postgresql import ARRAY

from tools.json_rw import json_read
from tools.get_hostname import get_hostname
from pathlib import Path

project_name = 'Test_report_gold'  # !!!
local_env_data = json_read(Path.home().joinpath(f'AppData\\Local\\.rpa').joinpath('env.json'))
global_path = Path(local_env_data['global_path'])
global_env_data = json_read(global_path.joinpath('env.json'))
project_path = global_path.joinpath(fr'.agent\{project_name}')
project_config_path = project_path.joinpath(f'{get_hostname()}\\config.json')
# project_config_data = json_read(project_config_path)

engine_kwargs = {
    'username': global_env_data['postgre_db_username'],
    'password': global_env_data['postgre_db_password'],
    'host': global_env_data['postgre_ip'],
    'port': global_env_data['postgre_port'],
    # 'base': project_config_data['base_name']
}

db_url = 'postgresql+psycopg2://{username}:{password}@{host}:{port}/orchestrator'.format(**engine_kwargs)

engine = create_engine(db_url)
Session = sessionmaker(bind=engine)
session = Session()

Base = declarative_base()


class InventoryTask(Base):
    __tablename__ = 'inventory'
    __table_args__ = {"schema": "robot_ars"}

    id = Column(Integer, primary_key=True, autoincrement=True)

    index = Column(Integer, default=0)  # Индекс строки
    branch = Column(String, default=None)  # Филиал
    location_code = Column(Integer, default=None)  # Код площадки
    inventory_type = Column(String, default=None)
    frequency = Column(String, default=None)
    date = Column(Date, default=None)
    note = Column(String, default=None)
    auditor = Column(String, default=None)
    auditor_us_gold = Column(String, default=None)
    format = Column(String, default=None)
    processing_status = Column(String, default=None)
    mailing_status = Column(String, default=None)
    days_until_inventory = Column(Integer, default=None)
    stock_share_tz_warehouse = Column(Numeric(5, 3), default=None)
    rto = Column(String, default=None)
    regional_director = Column(String, default=None)
    territorial_director = Column(String, default=None)
    branch_director = Column(String, default=None)
    deputy_director_administrator = Column(String, default=None)
    security_service_head = Column(String, default=None)
    division = Column(ARRAY(String), default=None)
    minus_day = Column(Date, default=None)

    # Флаги завершения задач
    done_projection = Column(Boolean, default=False)
    done_pallets = Column(Boolean, default=False)
    done_closed_docs = Column(Boolean, default=False)
    done_counters_task = Column(Boolean, default=False)
    done_mailing_superiors = Column(Boolean, default=False)
    done_saving_errors = Column(Boolean, default=False)
    done_full_task = Column(String, default='New')

    # Результаты задач
    res_projection1 = Column((String), default=None)
    res_projection2_1 = Column(ARRAY(String), default=None)
    res_projection2_2 = Column(ARRAY(String), default=None)
    res_pallets = Column(String, default=None)
    res_closed_docs = Column(String, default=None)
    res_counters_task = Column(ARRAY(String), default=None)
    res_mailing_superiors = Column(String, default=None)
    res_saving_errors = Column(Boolean, default=None)

    # Статусы ошибок
    # Нужны ли они?

    def __repr__(self):
        return "<Book(title='{}', author='{}', pages={}, published={})>" \
            .format(self.title, self.author, self.pages, self.published)


# Create the table in the "robot_ars" schema
Base.metadata.create_all(engine)

# Close the session
session.close()

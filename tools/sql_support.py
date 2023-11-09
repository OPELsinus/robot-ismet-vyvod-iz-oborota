import os
import sys
from pathlib import Path

import deprecated
import openpyxl
import sqlalchemy
from openpyxl import load_workbook
from pandas import ExcelWriter
from sqlalchemy import select, or_, update, and_
from sqlalchemy.engine.base import Engine
from sqlalchemy.orm import declarative_base, sessionmaker
from sqlalchemy.orm.session import Session

from typing import Sequence

sys.path.append(os.path.abspath(os.path.join(os.path.dirname('models.py'), '..')))

Base = declarative_base()


@deprecated.deprecated(reason='Неактуально')
def get_unfinished_tasks_deprecated(eng: Engine, model: Base) -> Sequence[sqlalchemy.engine.ScalarResult]:
    Session = sessionmaker(bind=eng)
    with Session() as session:
        query = select(model).where(
            or_(
                ~model.done_projection,
                ~model.done_pallets,
                ~model.done_closed_docs,
                ~model.done_counters_task,
                ~model.done_mailing_superiors,
                model.done_projection.is_(None),
                model.done_pallets.is_(None),
                model.done_closed_docs.is_(None),
                model.done_counters_task.is_(None),
                model.done_mailing_superiors.is_(None)
            )
        )
        unfinished_tasks = session.execute(query)
        return unfinished_tasks.scalars().all()


def get_unfinished_tasks(session: Session, model: Base) -> Sequence[sqlalchemy.engine.ScalarResult]:
    query = select(model).where(
        # TODO OR
        or_(
            model.done_full_task.in_(["New", "Repeatx2", "Repeatx3", "Fail"]),
        )
    )
    unfinished_tasks = session.execute(query)
    return unfinished_tasks.scalars().all()


@deprecated.deprecated(reason='Неактуально')
# Single
def get_unfinished_task_deprecated(eng: Engine, model: Base) -> sqlalchemy.engine.Result:
    """

    :rtype: object
    """
    Session = sessionmaker(bind=eng)
    with Session() as session:
        query = select(model).where(
            or_(
                ~model.done_projection,
                ~model.done_pallets,
                ~model.done_closed_docs,
                ~model.done_counters_task,
                ~model.done_mailing_superiors,
                model.done_projection.is_(None),
                model.done_pallets.is_(None),
                model.done_closed_docs.is_(None),
                model.done_counters_task.is_(None),
                model.done_mailing_superiors.is_(None)
            )
        ).limit(1)
        unfinished_task = session.execute(query).scalar()
        return unfinished_task


def get_unfinished_task(session: Session, model: Base, task_ids: list = None) -> sqlalchemy.engine.Result:
    """

    :rtype: object
    """
    if task_ids:
        query = select(model).where(
            and_(
                model.done_full_task.in_(["New", "Repeat_x2", "Repeat_x3"]),
                model.id.in_(task_ids)
            )
        ).limit(1)
    else:
        query = select(model).where(
            or_(
                model.done_full_task.in_(["New", "Repeat_x2", "Repeat_x3"])
            )
        ).limit(1)
    unfinished_task = session.execute(query).scalar()
    session.commit()
    return unfinished_task


def complete_task(session: Session, model: Base, id: int, done_task: str):
    """
    Updates done_task column in a normal db
    :param session: Session from sqlalchemy orm
    :param id: int value for the id of the task in db
    :param done_task: string value of a column name for "done..." task recognition
    :type model: Base, base model of sqlalchemy orm
    """
    column = getattr(model, done_task, None)
    if column is None:
        raise ValueError(f"Колонка '{done_task}' не существует в модели.")

    try:
        # Создаем объект query
        query = update(model).where(id == model.id).values(**{done_task: True})

        # Выполняем запрос
        session.execute(query)
    finally:
        # Завершаем транзакцию
        session.commit()


def update_res_counters_task(session: Session, model: Base, id: int, res_sk, res_tz):
    """
    Updates string-type column in a normal db
    :param session: Session from sqlalchemy orm
    :param model: Base, base model of sqlalchemy orm
    :param id: int value for the id of the task in db
    :param res_sk: value for res_sk
    :param res_tz: value for res_tz
    """
    try:
        # Создаем объект query
        query = update(model).where(model.id == id).values(res_counters_task=(str(res_sk), str(res_tz)))

        # Выполняем запрос
        session.execute(query)
    finally:
        # Завершаем транзакцию
        session.commit()


def update_res_projection(session: Session, model: Base, id: int, number, table_1, table_2):
    """
    Updates string-type column in a normal db
    :param session: Session from sqlalchemy orm
    :param model: Base, base model of sqlalchemy orm
    :param id: int value for the id of the task in db
    :param res_sk: value for res_sk
    :param res_tz: value for res_tz
    """
    try:
        # Создаем объект query
        query = update(model).where(model.id == id).values(res_projection1=number, res_projection2_1=table_1,
                                                           res_projection2_2=table_2)

        # Выполняем запрос
        session.execute(query)
    finally:
        # Завершаем транзакцию
        session.commit()



def update_res_stringtype(session: Session, model: Base, id: int, value: str, column_name) -> None:
    """
    Updates string-type column in a normal db
    :param session: Session from sqlalchemy orm
    :param model: Base, base model of sqlalchemy orm
    :param id: int value for the id of the task in db
    :param value: Normally a string value to put a path to the folder with response from the function
    :param column_name: Column name in db
    :rtype: None
    """
    # Создаем объект query
    if column_name == "res_projection":
        query = update(model).where(model.id == id).values(res_projection=value)
    elif column_name == "res_pallets":
        query = update(model).where(model.id == id).values(res_pallets=value)
    elif column_name == "res_closed_docs":
        query = update(model).where(model.id == id).values(res_closed_docs=value)
    elif column_name == "res_mailing_superiors":
        query = update(model).where(model.id == id).values(res_mailing_superiors=value)
    else:
        raise ValueError("Нету такой колонны в модели")

    # Выполняем запрос
    try:
        session.execute(query)
    finally:
        # Завершаем транзакцию
        session.commit()


def update_res_saving_errors(session: Session, model: Base, id: int, value: bool) -> None:
    """
    Updates bool column res_saving_errors in a normal db
    :param session: Session from sqlalchemy orm
    :param model: Base, base model of sqlalchemy orm
    :param id: int value for the id of the task in db
    :param value: Normally a string value to put a path to the folder with response from the function
    :rtype: None
    """
    # Создаем объект query
    query = update(model).where(model.id == id).values(res_saving_errors=value)

    try:
        # Выполняем запрос
        session.execute(query)
    finally:
        # Завершаем транзакцию
        session.commit()


def make_done(engine: Engine, model: Base, id: int, done_task: str):
    pass


def save_to_excel(excel_path, sheet_name="ЛИ в ТЗ", column="J", row=1, new_value: str = None):
    # Открываем Excel-файл
    try:
        workbook = openpyxl.load_workbook(excel_path)
        # Выбираем лист по имени
        sheet = workbook[sheet_name]

        cell = column + str(row)

        # Изменяем значение ячейки
        sheet[cell] = new_value

        # Сохраняем изменения
        workbook.save(excel_path)
    except Exception as e:
        print(f"Произошла ошибка: {str(e)}")
        raise Exception


def change_task_status(session: Session, model: Base, id: int, value: str) -> None:
    """
        Updates bool column res_saving_errors in a normal db
        :param session: Session from sqlalchemy orm
        :param model: Base, base model of sqlalchemy orm
        :param id: int value for the id of the task in db
        :param value: Normally a string value to put a path to the folder with response from the function
        :rtype: None
    """
    # Создаем объект query
    query = update(model).where(model.id == id).values(done_full_task=value)

    try:
        # Выполняем запрос
        session.execute(query)
    finally:
        # Завершаем транзакцию
        session.commit()


def get_message_info(session: Session, model: Base, id: int):
    info = {}
    # record_to_update = session.query(model).filter(model.id == id).first()
    query = (
        select(model).where(model.id == id))
    try:
        # Выполняем запрос
        result = session.execute(query).fetchone()[0]
        res_projection1 = result.res_projection1
        res_projection2_1 = result.res_projection2_1
        res_projection2_2 = result.res_projection2_2
        res_pallets = result.res_pallets
        res_closed = result.res_closed_docs
        res_counters_task = result.res_counters_task
        table = []
        for i in range(len(res_projection2_1)):
            # Создайте пару значений и добавьте их в таблицу
            pair = [res_projection2_1[i], res_projection2_2[i]]
            table.append(pair)
        info['res_projection1'] = res_projection1
        info['res_table'] = table
        info['res_pallets'] = res_pallets
        info['res_closed'] = res_closed
        info['res_counters_task'] = res_counters_task
    finally:
        # Завершаем транзакцию
        pass
    return info


def get_current_task_statuses(session: Session, model: Base, id: int):
    res = {}
    query = (
        select(model).where(model.id == id))
    try:
        # Выполняем запрос
        result = session.execute(query).fetchone()[0]
        done_projection = result.done_projection
        done_pallets = result.done_pallets
        done_closed_docs = result.done_closed_docs
        done_counters_task = result.done_counters_task
        res['res_projection'] = done_projection
        res['res_pallets'] = done_pallets
        res['res_closed_docs'] = done_closed_docs
        res['res_counters_task'] = done_counters_task
    finally:
        # Завершаем транзакцию
        pass
    return res
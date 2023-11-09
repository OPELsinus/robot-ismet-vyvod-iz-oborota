from pathlib import Path
from typing import Union, List
import calendar
from datetime import datetime, timedelta
import requests as rq
from rpamini.main import project_branches_path


# ? tested
def smtp_send(*args, subject: str, url: str, to: Union[list, str], username: str, password: str = None,
              html: str = None, attachments: List[Union[Path, str]] = None) -> None:
    # Нужны обяз. Текст в *args,subject - тема, url-ip БЕЗ ПОРТА!, to майл адреса, username - логин от кого,
    # password НЕ НУЖЕН! по желанию, html укзать в стринге, атачментс
    import smtplib
    from email.mime.application import MIMEApplication
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    body = ' '.join([str(i) for i in args])
    with smtplib.SMTP(url, 25) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        if password:
            smtp.login(username, password)

        msg = MIMEMultipart('alternative')
        msg["From"] = username
        msg["To"] = ';'.join(to) if type(to) is list else to
        msg["Subject"] = subject
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        if html:
            msg.attach(MIMEText(html, 'html', 'utf-8'))

        if attachments and isinstance(attachments, list):
            for each in attachments:
                path = Path(each).resolve()
                with open(path.__str__(), 'rb') as f:
                    part = MIMEApplication(f.read(), Name=path.name)
                    part['Content-Disposition'] = 'attachment; filename="%s"' % path.name
                    msg.attach(part)

        smtp.send_message(msg=msg)


def generate_html(task, inv_number, table_data, res_sk, res_tz):
    formatted_date = task.date.strftime('%d.%m.%Y')
    month_name = calendar.month_name[task.date.month]
    formatted_text = f"{task.frequency} {month_name.lower()} {formatted_date}"

    # ===
    date_obj = datetime.strptime(formatted_date, '%d.%m.%Y')
    new_date = date_obj - timedelta(days=1)
    new_formatted_date = new_date.strftime('%d.%m.%Y')

    html = f"""
            <p style="color: blue;">Коллеги, напоминаю вам о проведении пересчётов согласно Графика инвентаризаций ТМЦ на 2023г.</p>
            <p style="color: blue;">Прошу принять в работу.</p>
            
            <ul>
              <li>Инвентаризацию {formatted_text}, необходимо провести в ночь с {new_formatted_date} на {formatted_date}</li>
              <li>Ракурс по инвентаризации {formatted_text} под номером {inv_number}, создан в программе ГОЛД</li>
              <li>Ракурс созданной инвентаризации, для ознакомления (см. ниже):</li>
            </ul>
            <p><u>{formatted_text}</u></p>
            <table>
              <tr>
                <th>Merch.struct. node</th>
                <th>Описание торг. структуры</th>
              </tr>
        """

    for row in table_data:
        html += f"<tr><td>{row[0]}</td><td>{row[1]}</td></tr>"

    html += f"""\
        </table>
        <p>Рекомендуемое количество счетчиков для пересчета секции Бакалея:</p>
            <table>
                <tr>
                    <th>Филиал</th>
                    <th>Дата проведения ЛИ</th>
                    <th>Кол-во счётчиков для пересчёта запасов</th>
                    <th>Кол-во счётчиков для пересчёта ТЗ</th>
                </tr>
                <tr>
                    <td>{task.branch}</td>
                    <td>{formatted_date}</td>
                    <td>По {res_sk} счетчика за {task.days_until_inventory} дней</td>
                    <td>{res_tz}</td>
                </tr>
            </table>
        <p style="color: white;">    Инструкция “Проведение инвентаризации товаров реализуемых в 
        торговых залах филиалов ТОО Magnum Cash&Carry” расположена в сетевом диске M:\Stuff\_56_ЛИ\2023 </p>
        <p style="color: black;"><u>Во вложении список не принятых паллет и отчёт 26 по незакрытым документам.
         Необходимо принять в работу и отработать до начала основного пересчёта</u></p>
        <p>При возникновении проблем технического характера необходимо по электронной почте отправлять заявки
          через <u style="color: blue;">SD@magnum.kz</u> по установленной форме и с обязательным прикреплением 
          скриншотов ошибки в учетной системе или в ТСД.
        </p>
        <p><u>Прошу довести информацию до всех ответственных сотрудников по процессу 
        проведения инвентаризаций ТМЦ в ТЗ.</u>
        </p>
    """

    return html


def find_receivers(task):
    path = project_branches_path
    def filter_dicts(dicts, key, value):
        return filter(lambda d: d.get(key) == value, dicts)
    # TODO FINISH
    import pandas as pd
    df = pd.read_excel(path)
    url = "http://172.16.11.11/ZUP/hs/Report/3"

    resp = rq.get(url=url, data={}, headers={}, timeout=180)
    parsed_data = resp.json()

    branch_code = df[df['Филиал'] == task.branch]['Branch code'].item()

    filtered_dicts = list(filter_dicts(parsed_data, 'branch_code', branch_code))

    emails_DF = []
    if task.branch_director != "ДФ":
        for dict in filtered_dicts:
            if dict['full_name'] == task.branch_director:
                emails_DF.append(dict['email'])
    else:
        for dict in filtered_dicts:
            if dict['position'] == "Директор Филиала":
                emails_DF.append(dict['email'])

    emails_HDF = []
    if task.deputy_director_administrator != "ЗДФ":
        for dict in filtered_dicts:
            if dict['full_name'] == task.deputy_director_administrator:
                emails_HDF.append(dict['email'])

    else:
        for dict in filtered_dicts:
            if dict['position'] == "Заместитель директора филиала":
                emails_HDF.append(dict['email'])

    emails_DF.extend(emails_HDF)
    emails_DF.extend(['zetty305@mail.ru', 'Mukhtarova@magnum.kz', 'Osmanova@magnum.kz', 'Kuanyshov.A@magnum.kz'])
    return emails_DF
    # return ["zetty305@mail.ru"]

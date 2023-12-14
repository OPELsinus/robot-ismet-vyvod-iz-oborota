import datetime
import json
import time
from contextlib import suppress
from time import sleep

import requests
from selenium.webdriver.support.expected_conditions import visibility_of_element_located

from config import logger
from tools.app import App
from tools.web import Web

file_selector = {"title": "Открыть файл", "class_name": "SunAwtDialog", "found_index": 0}
pass_selector = {"title": "Формирование ЭЦП в формате CMS", "class_name": "SunAwtDialog", "found_index": 0}

months = ['', 'январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август',
          'сентябрь', 'октябрь', 'ноябрь', 'декабрь']


def ismet_auth(ecp_auth: str, ecp_sign: str):

    web = Web()
    web.run()

    login_url = 'https://elk.prod.markirovka.ismet.kz/login-kep'
    web.get(login_url)
    web.find_element('//button[contains(., "Выбрать сертификат")]').click()

    # * auth nca
    nca = App('')
    file_element = nca.find_element(file_selector)
    file_element.type_keys(ecp_auth, set_focus=True, protect_first=True)
    sleep(1.5)
    file_element.type_keys(nca.keys.ENTER, set_focus=True)
    pass_element = nca.find_element(pass_selector)
    pass_element.type_keys('Aa123456', nca.keys.ENTER, set_focus=True)
    sleep(1.5)
    pass_element.type_keys('Aa123456', nca.keys.ENTER, set_focus=True)
    sleep(1)

    # * check success auth

    web.get('https://goods.prod.markirovka.ismet.kz/documents/list')

    selector = '//span[text()="Добавить документ"]'

    # * raise no count
    if not web.wait_element(selector, timeout=30):
        logger.warning('Ошибка авторизации')
        return None

    return web


def load_document_to_out(web: Web, filepath: str, year: int = None, month: int = None, day: int = None, url: str = None):

    print()

    web.get('https://goods.prod.markirovka.ismet.kz/documents/list')

    selector = '//span[text()="Добавить документ"]'
    web.find_element(selector).click()
    sleep(0.3)
    web.find_element('//li[contains(text(), "Уведомление о выводе из оборота")]').click()

    print()
    web.find_element("//label[contains(text(), 'Причина вывода из оборота')]/following-sibling::div").click()
    web.find_element("//li[contains(text(), 'Розничная продажа')]").click()

    web.find_element("//label[contains(text(), 'Наименование документа основания')]/following-sibling::div").click()
    web.find_element("//label[contains(text(), 'Наименование документа основания')]/following-sibling::div/input").type_keys('Чек продажи')

    web.find_element("//label[contains(text(), 'Номер документа основания')]/following-sibling::div").click()
    web.find_element("//label[contains(text(), 'Номер документа основания')]/following-sibling::div/input").type_keys('-')

    web.find_element("//label[contains(text(), 'Дата документа основания')]/following-sibling::div/input").click()
    web.wait_element("//span[contains(text(), 'Применить')]")

    web.find_element("//label[contains(text(), 'Год')]/following-sibling::div").click()
    web.find_element(f"//li[contains(text(), '{year}')]").click()

    web.find_element("//label[contains(text(), 'Месяц')]/following-sibling::div").click()
    web.find_element(f"//li[contains(text(), '{months[month]}')]").click()

    web.find_element(f"//button[contains(text(), '{day}')]").click()

    web.execute_script_click_xpath("//span[contains(text(), 'Применить')]")

    web.execute_script_click_xpath("//span[contains(text(), 'Следующий')]")

    # * Next page
    # print('sending request')
    # all_cookies = web.driver.get_cookies()
    # cookies_dict = {}
    # for cookie in all_cookies:
    #     cookies_dict[cookie['name']] = cookie['value']
    #
    # now = datetime.datetime.utcnow()
    # formatted_date = time.strftime('%a, %d %b %Y %H:%M:%S GMT', now.timetuple())
    #
    # bearer_token = f"{cookies_dict.get('tokenPart1')}{cookies_dict.get('tokenPart2')}"
    # print("bearer:", bearer_token)
    # headers = {
    #     'Authorization': f'Bearer {bearer_token}',
    #     'Access-Control-Allow-Credentials': 'true',
    #     'Access-Control-Allow-Origin': 'https://goods.prod.markirovka.ismet.kz',
    #     'Access-Control-Expose-Headers': 'Authorization, Link, X-Total-Count',
    #     'Cache-Control': 'no-cache, no-store, max-age=0, must-revalidate',
    #     'Connection': 'keep-alive',
    #     'Content-Type': 'application/json;charset=UTF-8',
    #     'Date': formatted_date,
    #     'Expires': '0',
    #     'Pragma': 'no-cache',
    #     'Server': 'nginx',
    #     'Vary': 'Access-Control-Request-Headers, Access-Control-Request-Method, Origin',
    #     'X-Content-Type-Options': 'nosniff',
    #     'X-Frame-Options': 'DENY',
    #     'X-Xss-Protection': '1; mode=block'
    # }
    #
    # def send_excel_file(file_path, url, bearer_token, cookies_dict):
    #     multipart_form_data = {
    #         'properties': ('blob', json.dumps({"userId": 600128503, "organisationInn": "191041025674"}), 'application/json'),
    #         'file': ('Торговый зал ШФ №7.xlsx', open(file_path, 'rb'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    #     }
    #
    #     with requests.Session() as session:
    #         session.cookies.update(cookies_dict)
    #         response = session.post(url, files=multipart_form_data, headers=headers, verify=False)
    #
    #     return response
    #
    # # Send the file
    # response = send_excel_file(filepath, url, bearer_token, cookies_dict)
    # print(response.status_code)
    # print(response.text)
    # print('sent request')
    # sleep(1000)

    web.find_element("//span[contains(text(), 'ДОБАВИТЬ ТОВАР')]").click()

    web.find_element("//div[contains(text(), 'Загрузить из файла')]").click()

    web.find_element("//div[contains(text(), 'Перетащите или выберите файл')]").click()

    # * Opening a file with codes to dropout

    app = App('')

    app.find_element({"title": "Имя файла:", "class_name": "Edit", "control_type": "Edit",
                      "visible_only": True, "enabled_only": True, "found_index": 0}).click(set_focus=True)

    app.find_element({"title": "Имя файла:", "class_name": "Edit", "control_type": "Edit",
                      "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(filepath)

    app.find_element({"title": "Открыть", "class_name": "Button", "control_type": "Button",
                      "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    # print()


def select_all_wares_to_dropout(web: Web, ecp_sign: str):

    timeout_ = 70

    if web.wait_element("//div[contains(text(), 'Невозможно выполнить запрос')]", timeout=30):
        raise Exception("Error when loading an Excel")

    while True:

        web.wait_element("//div[@class='rt-th sc-kkGfuU dzsQrm']/div/div/div", timeout=timeout_)

        if timeout_ == 70:
            if web.wait_element("//span[text() = 'Отмена']", timeout=3):
                web.find_element("//span[text() = 'Отмена']", timeout=3).click()

        web.find_element("//div[@class='rt-th sc-kkGfuU dzsQrm']/div/div/div").click()

        try:
            current_page = int(web.find_element("//li[@class='sc-giadOv ekuBbr']", timeout=5).get_attr('text'))
        except:
            break

        available_page = None
        for pages in web.find_elements("//li[@class='sc-giadOv hALQxM']"):
            with suppress(Exception):
                if int(pages.get_attr('text')) > current_page:
                    available_page = int(pages.get_attr('text'))
                    break

        if available_page is None:
            break

        web.find_element(f"//li[text()='{available_page}']", timeout=5).click()

        timeout_ = 5

    # print()

    web.find_element('//span[text()="Отправить"]/../..').click()
    print(ecp_sign)
    nca = App('')
    file_element = nca.find_element(file_selector)
    file_element.type_keys(ecp_sign, set_focus=True, protect_first=True)
    sleep(1.5)
    file_element.type_keys(nca.keys.ENTER, set_focus=True)
    pass_element = nca.find_element(pass_selector)
    pass_element.type_keys('Aa123456', nca.keys.ENTER, set_focus=True)
    sleep(1.5)
    pass_element.type_keys('Aa123456', nca.keys.ENTER, set_focus=True)
    sleep(1)

    # * Ждём пока портал обработает наш отчёт (Чтобы избежать статуса Проверяется)

    # sleep(20)
    #
    # web.find_element('(//a[@class="sc-dHIava Fduvs"])[1]').click()
    #
    # web.find_element('//button[text()="Печать"]').click()
    #
    # print()

from time import sleep


from selenium.webdriver.support.expected_conditions import visibility_of_element_located

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
    if not web.wait_element(selector):
        raise Exception('Ошибка авторизации')

    return web


def load_document_to_out(web: Web, year: int = None, month: int = None, day: int = None):

    selector = '//span[text()="Добавить документ"]'
    web.find_element(selector).click()

    web.find_element('//li[contains(text(), "Уведомление о выводе из оборота")]').click()

    print()
    web.find_element("//label[contains(text(), 'Причина вывода из оборота')]/following-sibling::div").click()
    web.find_element("//li[contains(text(), 'Розничная продажа')]").click()

    web.find_element("//label[contains(text(), 'Наименование документа основания')]/following-sibling::div").click()
    web.find_element("//label[contains(text(), 'Наименование документа основания')]/following-sibling::div/input").type_keys('Тест')

    web.find_element("//label[contains(text(), 'Номер документа основания')]/following-sibling::div").click()
    web.find_element("//label[contains(text(), 'Номер документа основания')]/following-sibling::div/input").type_keys('1')

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

    web.find_element("//span[contains(text(), 'ДОБАВИТЬ ТОВАР')]").click()

    web.find_element("//div[contains(text(), 'Выбрать из списка')]").click()

    print()





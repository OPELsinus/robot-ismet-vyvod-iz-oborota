from __future__ import annotations

from contextlib import suppress
from datetime import datetime
from pathlib import Path
from time import sleep
from typing import Union
import regex as re

from pywinauto.timings import wait_until
from selenium import webdriver
from selenium.webdriver import ChromeOptions, Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.expected_conditions import visibility_of_element_located, presence_of_element_located, \
    element_to_be_clickable
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.common.action_chains import ActionChains  # Corrected import
from time import time


class Web:
    keys = Keys

    # noinspection SpellCheckingInspection,PyBroadException
    class Element:
        keys = Keys

        def __init__(self, element, selector, by, driver, parent):
            self.element: WebElement = element
            self.selector = selector
            self.by = by
            self.driver: WebDriver = driver
            self.parent: Web = parent

        def find_all_web_elements_with_element(self):
            """
            Возвращает список всех WebElement на странице, включая элемент, переданный в качестве аргумента.

            :param self: WebElement, с которого начинается поиск
            :return: список WebElement
            """
            # Получить родительский элемент переданного элемента
            parent_element = self.find_element('..')

            # Используйте XPath для поиска всех элементов на странице, находящихся внутри родительского элемента
            all_elements = parent_element.find_elements('.//*')

            return all_elements

        def page_load(self, operand: str, timeout=60):
            def wait_for_url():
                return operand != self.driver.current_url

            wait_until(timeout, 0.5, wait_for_url)

        def click(self, double=False, delay: int or float = 0, scroll=False, page_load=False, el_selector: str = None,
                  timeout: int = 10):
            # Если указывать el, то page_load будет True
            # from __future__ import annotations для лечения проблемы циклических зависимостей
            if el_selector:
                page_load = True
                find = True
            else:
                find = False
            sleep(delay)
            if scroll:
                with suppress(Exception):
                    self.scroll()
            url = self.driver.current_url
            ActionChains(self.driver).double_click(self.element).perform() if double else self.element.click()
            if page_load:
                if el_selector:
                    print(self.wait_element(selector=el_selector, timeout=timeout, find=find))
                else:
                    self.page_load(url)

        def scroll(self, delay=0):
            sleep(delay)
            ActionChains(self.driver).move_to_element(self.element).perform()

        def clear(self, delay=0):
            sleep(delay)
            self.element.clear()

        def get_attr(self, attr='text', delay=0, scroll=False):
            sleep(delay)
            if scroll:
                with suppress(Exception):
                    self.scroll()
            return getattr(self.element, attr) if attr in ['tag_name', 'text'] else self.element.get_attribute(attr)

        def set_attr(self, value=None, attr='value', delay=0, scroll=False):
            sleep(delay)
            if scroll:
                with suppress(Exception):
                    self.scroll()
            self.driver.execute_script(f"arguments[0].{attr} = arguments[1]", self.element, value)

        def type_keys(self, *value, delay=0, scroll=False, clear=False):
            sleep(delay)
            if scroll:
                with suppress(Exception):
                    self.scroll()
            if clear:
                self.clear()
            self.element.send_keys(*value)

        def select(self, value=None, select_type='select_by_value', delay=0, scroll=False):
            sleep(delay)
            if scroll:
                with suppress(Exception):
                    self.scroll()
            select = Select(self.element)
            function = getattr(select, select_type)
            if value is None:
                if select_type == 'deselect_all':
                    return function()
                else:
                    return select
            else:
                return function(value)

        def find_elements(self, selector, timeout=60, by='xpath'):
            selector = f'.{selector}' if selector[0] != '.' else selector
            if timeout:
                self.wait_element(selector, timeout, by)
            elements = self.element.find_elements(by, selector)
            selector = f'{self.selector}{selector[1:]}'
            elements = [Web.Element(element=element, selector=selector, by=by, driver=self.driver, parent=self.parent)
                        for element in elements]
            return elements

        def find_element(self, selector, timeout=60, by='xpath'):
            selector = f'.{selector}' if selector[0] != '.' else selector
            if timeout:
                self.wait_element(selector, timeout, by)
            element = self.element.find_element(by, selector)
            selector = f'{self.selector}{selector[1:]}'
            element = Web.Element(element=element, selector=selector, by=by, driver=self.driver, parent=self.parent)
            return element

        def wait_element(self, selector, timeout=60, by='xpath', until=True, find: bool = True):
            selector = f'.{selector}' if selector[0] != '.' else selector

            if find:
                def find():
                    try:
                        self.element.find_element(by, selector)
                        return True
                    except (Exception,):
                        return False
            else:
                def find():
                    try:
                        self.parent.wait_element(selector, event=visibility_of_element_located, timeout=15)
                        return True
                    except (Exception,) as e:
                        print(e)
                        return False

            try:
                return wait_until(timeout, 0.5, find, until)
            except (Exception,):
                return False

        def recursive_open_until_interactable_old2(self, text):
            """
            Раскрывать список, нажимая на кнопку плюсик, пока элемент a[contains(text(), {text})] не будет кликабельным
            :param text: text in element which we want to be interactable
            :return: res: bool. True if clickable, False if no more daughters to unfold
            """

            # Плюсик
            try:
                ocl = self.find_element(selector="//i[@class='jstree-icon jstree-ocl']", timeout=1)
                if 'jstree-last jstree-open' not in self.element.get_attribute("class"):
                    if ocl:
                        ocl.click(delay=0.2, scroll=True)
            except:

                pass

            try:
                self.find_element(selector=f'//a[contains(text(), "{text}")]/i[@class="jstree-icon jstree-checkbox"]',
                                  timeout=3).click(delay=0.2, scroll=True)
                return True
            except Exception:
                children = self.find_children()
                if not children:
                    return False

                for child in children:
                    if child.recursive_open_until_interactable(text):
                        return True
                    else:
                        return False

        def find_children(self):
            try:
                el_ul = self.find_element(selector="//ul[@class='jstree-children']", timeout=0.2)
                try:
                    level = self.find_element(selector="//a", timeout=0.2).element.get_attribute("aria-level")
                    daughters = el_ul.find_elements(
                        selector=f".//li[contains(@class, 'jstree-node') and"
                                 f" .//a[@aria-level='{str(int(level) + 1)}']]", timeout=0.2)
                    # daughters = el_ul.find_elements(selector="//li[contains(@class, 'jstree-node')]", timeout=2)
                    return daughters
                except:
                    return None
            except:
                return None

        def recursive_open_until_interactable_old3(self, text):
            """
            Раскрывать список, нажимая на кнопку плюсик, пока элемент a[contains(text(), {text})] не будет кликабельным
            :param text: text in element which we want to be interactable
            :return: res: bool. True if clickable, False if no more daughters to unfold
            """

            # Плюсик
            start_time = time()
            try:
                ocls = self.find_elements(selector="//i[@class='jstree-icon jstree-ocl']", timeout=0.2)
                for ocl in ocls:
                    if 'jstree-open' not in self.element.get_attribute("class"):
                        if ocl:
                            ocl.click(delay=0.2, scroll=True)
                            sleep(0.2)
            except:
                pass

            try:
                element = self.find_element(selector=f'//a[contains(text(), "[{text}]")]', timeout=0.2)
                element.click(delay=0.2, scroll=True)
                return True
            except Exception:
                children = self.find_children()
                if children:
                    for child in children:
                        success = child.recursive_open_until_interactable(text)
                        if success:
                            return True
            return False

        def open_until_interactable(self, text):
            """

            """

            def parse_text(text):
                match = re.match(r"([A-Za-z]+)([0-9]+)", text)
                if not match:
                    raise ValueError("Текст не соответствует ожидаемому формату: буквы за которыми следуют цифры")

                # Извлекаем буквенную и числовую части
                category = match.group(1)
                numeric_part = match.group(2)

                return category, numeric_part

            mapping = {
                'F': '1',
                'NF': '2',
                'CP': '3',
                'T': '6',
                'S': '7',
                'A': '8',
            }
            # Разделить текст на буквенную и числовую части
            category, numeric_part = parse_text(text)

            # Находим основную категорию по буквенной части и проверяем, открыта ли она
            if not self.is_category_open(mapping[category]):
                self.open_category(mapping[category])

            # Перебор числовой части, открываем вложенные элементы
            for i in range(0, len(numeric_part) - 2, 2):
                numeric_section = numeric_part[:i + 2]
                self.open_subcategory(category, numeric_section)

            # Ищем и кликаем по конечному элементу
            self.click_final_element(text)

        def is_category_open(self, category_name):
            el = self.find_element(selector=f"//a[contains(text(), '[{category_name}]')]/..", timeout=1)
            return 'jstree-open' in el.element.get_attribute("class")

        def open_category(self, category_name):
            el = self.find_element(selector=f"//a[contains(text(), '[{category_name}]')]/..",
                                   timeout=1).find_element(selector="//i[@class='jstree-icon jstree-ocl']", timeout=0.2)
            el.click(delay=0.5, scroll=True)

        def open_subcategory(self, category, numeric_section):
            el = self.find_element(
                selector=f"//a[contains(text(), '[{category + numeric_section}]')]/..", timeout=1)
            if 'jstree-open' not in el.element.get_attribute("class"):
                el.find_element(selector="//i[@class='jstree-icon jstree-ocl']", timeout=0.2).click(delay=0.2,
                                                                                                    scroll=True)

        def click_final_element(self, text):
            el = self.find_element(selector=f"//a[contains(text(), '[{text}]')]", timeout=1)
            el.click(delay=0.2, scroll=True)

        def recursive_open_until_interactable_and_click(self, text_selector, laying_selector, laying_brother_selector,
                                                        child_selector):
            """

            :param child_selector: "//div[@class='treeNode']"
            :param laying_brother_selector: "//div[@class='treeRow']"
            :param laying_selector: "//div[@class='childrenContainer']"
            :param text: "GOLDREPORT-148 - Отчет по паллетам"
            :return:
            """
            try:
                attribute_isvisible = self.find_element(selector=laying_selector, timeout=0.5).element.is_displayed()
            except NoSuchElementException or TimeoutException:
                attribute_isvisible = True
            except Exception:
                attribute_isvisible = True

            if not attribute_isvisible:
                self.find_element(selector=laying_brother_selector, timeout=0.5).click(delay=0.2, scroll=True)

            try:
                # WebDriverWait(self.driver, 2).until(
                #     element_to_be_clickable((By.XPATH, f"//*[contains(text(), '{text}')]")))

                self.find_element(selector=text_selector, timeout=0.5).click(double=True, delay=0.2, scroll=True)
                return True
            except (ElementNotInteractableException, Exception):
                try:
                    laying_el = self.find_element(selector=laying_selector, timeout=0.5)
                    children = laying_el.find_children_by_name([child_selector])
                except NoSuchElementException or TimeoutException:
                    return False
                except Exception:
                    return False

                counter = 0
                if children:
                    for child in children:
                        counter += 1
                        print(counter)
                        if child.recursive_open_until_interactable_and_click(text_selector, laying_selector,
                                                                             laying_brother_selector,
                                                                             child_selector):
                            return True
                else:
                    return False

        def find_children_by_name(self, children_names):
            children = []

            for child_name in children_names:
                daughters = self.find_elements(selector=child_name, timeout=2)
                if daughters:
                    children.extend(daughters)

            if children:
                return children
            else:
                return None

    def __init__(self, path=None, download_path=None, run=False, timeout=60):
        self.path = path or Path.home().joinpath(r"AppData\Local\.rpa\Chromium\chromedriver.exe")
        self.download_path = download_path or Path.home().joinpath('Downloads')
        self.run_flag = run
        self.timeout = timeout

        self.options = ChromeOptions()
        self.options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
        self.options.add_experimental_option("useAutomationExtension", False)
        self.options.add_experimental_option("prefs", {
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False,
            "profile.default_content_settings.popups": 0,
            "download.default_directory": self.download_path.__str__(),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": False,
            "profile.content_settings.exceptions.automatic_downloads.*.setting": 1
        })
        self.options.add_argument("--start-maximized")
        self.options.add_argument("--no-sandbox")
        self.options.add_argument("--disable-dev-shm-usage")
        self.options.add_argument("--disable-print-preview")
        self.options.add_argument("--disable-extensions")
        self.options.add_argument("--disable-notifications")
        self.options.add_argument("--ignore-ssl-errors=yes")
        self.options.add_argument("--ignore-certificate-errors")

        # noinspection PyTypeChecker
        self.driver: WebDriver = None

    def expand_and_click(self, element_text):
        # Находим главную ноду, начиная с ROOT - ROOT
        main_node = WebDriverWait(self.driver, 10).until(
            presence_of_element_located(
                (By.XPATH, '//*[@class="treeNode treeRootNode"]'))
        )

        # Рекурсивная функция для поиска элемента в дереве
        def find_and_click(node):
            # Нажимаем на элемент, чтобы развернуть childrenContainer
            node.find_element(By.CLASS_NAME, 'treeRow').click()
            # Ждем, пока childrenContainer станет видимым
            WebDriverWait(self.driver, 10).until(
                visibility_of_element_located((By.CLASS_NAME, 'childrenContainer'))
            )
            # Ищем элемент с текстом element_text внутри childrenContainer
            try:
                element = node.find_element(By.XPATH,
                                            f'.//span[@class="rowContent"]/span[@class="rowLabel"][contains(text(),'
                                            f' "{element_text}")]')
                action = webdriver.common.action_chains.ActionChains(self.driver)
                action.move_to_element(element).click().perform()
                return True
            except NoSuchElementException:
                # Если элемент не найден, рекурсивно вызываем функцию для дочерних элементов
                children = node.find_element(By.CLASS_NAME, 'childrenContainer').find_elements(By.CLASS_NAME,
                                                                                               'treeNode')
                for child in children:
                    if find_and_click(child):
                        return True
            return False

        # Запускаем поиск, начиная с главной ноды
        find_and_click(main_node)

    def run(self):
        self.quit()
        self.driver = webdriver.Chrome(service=Service(self.path.__str__()), options=self.options)

    def quit(self):
        if self.driver:
            self.driver.quit()

    def close(self):
        self.driver.close()

    def switch(self, switch_type='window', switch_index=-1, frame_selector=None):
        if switch_type == 'window':
            self.driver.switch_to.window(self.driver.window_handles[switch_index])
        elif switch_type == 'frame':
            if frame_selector:
                self.driver.switch_to.frame(self.find_elements(frame_selector)[switch_index].element)
            else:
                raise Exception('selected type is "frame", but didnt received frame_selector')
        elif switch_type == 'alert':
            self.driver.switch_to.alert.accept()
        raise Exception(f'switch_type "{switch_type}" didnt found')

    def get(self, url):
        self.driver.get(url)

    def find_elements(self, selector, timeout=None, event=None, by='xpath'):
        if event is None:
            event = expected_conditions.presence_of_element_located
        timeout = timeout if timeout is not None else self.timeout
        if timeout:
            self.wait_element(selector, timeout, event, by)
        elements = self.driver.find_elements(by, selector)
        elements = [self.Element(element=element, selector=selector, by=by, driver=self.driver, parent=self) for element
                    in elements]
        return elements

    def find_element(self, selector, timeout=None, event=None, by='xpath'):
        if event is None:
            event = expected_conditions.presence_of_element_located
        timeout = timeout if timeout is not None else self.timeout
        if timeout:
            self.wait_element(selector, timeout, event, by)
        element = self.driver.find_element(by, selector)
        element = self.Element(element=element, selector=selector, by=by, driver=self.driver, parent=self)
        return element

    def wait_element(self, selector, timeout=None, event=None, by='xpath', until=True):
        if event is None:
            event = expected_conditions.presence_of_element_located
        try:
            timeout = timeout if timeout is not None else self.timeout
            wait = WebDriverWait(self.driver, timeout)
            event = event((by, selector))
            wait.until(event) if until else wait.until_not(event)
            return True
        except (Exception,):
            return False

    @staticmethod
    def wait_downloaded(target: Union[Path, str], timeout: Union[int, float] = 60) -> Union[Path, None]:
        start_time = datetime.now()
        while True:
            target = Path(target)
            folder = target.parent
            files = folder.glob(target.name)
            for file_path in files:
                if not any(temp in str(file_path) for temp in ['.crdownload', '~$']):
                    if file_path.is_file() and file_path.stat().st_size > 0:
                        return file_path
            if int((datetime.now() - start_time).seconds) > timeout:
                return None
            sleep(1)

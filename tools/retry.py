import time
from selenium.common.exceptions import NoSuchElementException


def try_except_decorator(retry_cout=2, retry_delay=1):
    import traceback
    from time import sleep

    def decorator(func):
        def wrapper(*args, **kwargs):
            for _ in range(retry_cout):
                try:
                    result = func(*args, **kwargs)
                    return result
                except (Exception,):
                    traceback.print_exc()
                    sleep(retry_delay)
            raise Exception('retry_cout <= 0')

        return wrapper

    return decorator


def find_title_or_new_window(app, selector="//div[@class='GSOInventoryInitialization']", initial_window_count=2,
                             timeout=600):
    end_time = time.time() + timeout
    while True:
        try:
            # Пытаемся найти элемент с id 'divTitleGSOInventoryInitialization'
            element = app.find_element(selector=selector, timeout=0.3)
            print("Элемент найден!")
            return True
        except NoSuchElementException:
            # Если элемент не найден, проверяем количество окон
            current_window_count = len(app.driver.window_handles)
            if current_window_count > initial_window_count:
                print(f"Обнаружено новое окно. Количество окон: {current_window_count}")
                return False
            # Если элемент не найден и новых окон нет, ждем 10 секунд
            print("Элемент не найден, новых окон нет. Ждем 10 секунд...")
            time.sleep(10)

        # Проверяем, не истекло ли общее время ожидания
        if time.time() > end_time:
            print("Время ожидания истекло. Элемент или новое окно так и не было обнаружено.")
            return False

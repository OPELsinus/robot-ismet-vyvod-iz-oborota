import os
import shutil
from time import sleep

from config import download_path, working_path, saving_path


def wait_report_to_download(branch: str, date_: str):

    filepath = None
    found = False
    while True:
        for file in os.listdir(download_path):
            if 'уведомление' in file.lower() and 'вывод' in file.lower() and 'оборот' in file.lower() and '$' not in file and '.crdownload' not in file:
                sleep(0.03)
                shutil.move(os.path.join(download_path, file), os.path.join(saving_path, f"{file.split('.')[0]} {branch} {date_}.{file.split('.')[1]}"))
                filepath = os.path.join(saving_path, f"{file.split('.')[0]} {branch} {date_}.{file.split('.')[1]}")
                found = True
                break
        if found:
            break

    return filepath
























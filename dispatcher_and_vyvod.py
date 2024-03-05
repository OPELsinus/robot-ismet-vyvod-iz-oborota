from contextlib import suppress

from config import logger
from skipped_files_dispatcher import dispatcher
from vyvod_2022 import vyvodbek

if __name__ == '__main__':

    for i in range(1):

        status = None

        try:

            status = dispatcher()

        except Exception as err:

            logger.warning(f'Errorbek1: {err}')

        logger.warning(f'STATUS: {status}')

        if status == 'good':

            try:

                print('Started vyvod')

                vyvodbek()

            except Exception as err1:

                logger.warning(f'Errorbek2: {err1}')

        print('Kek')



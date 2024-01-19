import sys

from config import logger, process_list_path
from tools.process import kill_process_list

if __name__ == '__main__':
    # ? не убирать данный try, он необходим для того чтобы Pyinstaller не выводил traceback в окошко
    try:
        logger.info('info log')
        logger.warning('warning log')
    except (Exception,):
        kill_process_list(process_list_path)
        sys.exit(1)

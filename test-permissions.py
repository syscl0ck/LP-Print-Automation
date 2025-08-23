import os
import stat
from logging import Logger, INFO
from logging.handlers import RotatingFileHandler

logger = Logger(__name__)
logger.setLevel(INFO)
handler = RotatingFileHandler("test.log", maxBytes=1024*1024, backupCount=5)
logger.addHandler(handler)
logger.info(f'[+] Logfile loaded at test.log')

temp_xlsx_path = r'C:\Users\isaac\Desktop\LP-Print-Automation\tmp\8f70143a-35a3-41b2-825b-c405e29606b6.xlsx'

logger.info(f"[+] Checking file permissions...")
logger.info(f"Exists? {os.path.exists(temp_xlsx_path)}")
logger.info(f"Readable? {os.access(temp_xlsx_path, os.R_OK)}")
logger.info(f"Writable? {os.access(temp_xlsx_path, os.W_OK)}")
logger.info(f"Executable? {os.access(temp_xlsx_path, os.X_OK)}")
logger.info(f"Stat: {os.stat(temp_xlsx_path)}")
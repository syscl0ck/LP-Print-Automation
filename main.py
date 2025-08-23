import os
import io
import tempfile
import json
import tempfile
from pathlib import Path
import shutil
import stat
from logging import Logger, INFO
from logging.handlers import RotatingFileHandler
import win32com.client as win32
from threading import Thread
from time import sleep
from watchdog.observers.polling import PollingObserver as Observer
from watchdog.events import FileSystemEventHandler
from uuid import uuid4
import pythoncom


##################################################################
# Created by Mountain Iron Technology LLC for Heliene USA
# Author: Isaac Burton
# License: MIT
# Date: July 8, 2025
#
# Automatic Printing for Palletizer
##################################################################


with open('settings.json') as config_file:
	conf = json.load(config_file)

logger = Logger(__name__)
logger.setLevel(INFO)
handler = RotatingFileHandler(os.path.join(os.getcwd(),conf['LOGFILE']), maxBytes=1024*1024, backupCount=5)
logger.addHandler(handler)
logger.info(f'[+] Logfile loaded at {conf["LOGFILE"]}')

COPY_PATH = os.path.join(os.getcwd(),conf['COPY_PATH'])

PRINT_METHOD = conf['PRINT_METHOD']

BASEDIR = conf['BASEDIR']

# Time between filesystem checks - in seconds
POLL_DURATION = conf['POLL_DELAY_SECONDS']

TMP_PATH = conf['TMP_PATH']

DEBUG = conf['DEBUG']

class NewFileHandler(FileSystemEventHandler):
      def on_created(self, event):
            if not event.is_directory:
                  print(f"New file created: {event.src_path}")
                  worker = Thread(target=process_document)
                  worker.start()

def get_newest_spreadsheet(basedir):
	def descend_newest_path(path):
		while True:
			# Check to see if we are in the directory with Excel Spreadsheets
			xlsx_files = [f for f in os.listdir(path) 
							if f.lower().endswith('.xlsx') 
							and os.path.isfile(os.path.join(path, f))]
			if xlsx_files:
				return path

			# Create a list of subdirectories
			subdirs = [os.path.join(path, d) for d in os.listdir(path)
						if os.path.isdir(os.path.join(path, d))]
			# If there are none found, and no XLSX files were found, exit 
			if not subdirs:
				return None

			# Get the newest created subdir
			newest_subdir = max(subdirs, key=os.path.getctime)

			# Set path and continue loop
			path = newest_subdir
			print(f'[+] Newest subdirectory: {newest_subdir}')

	def get_newest_file(path):
		if not path:
			return None
		
		# Create list of all .xlsx files in the current path
		files = [os.path.join(path, f) for f in os.listdir(path)
						if f.lower().endswith('.xlsx')
						and os.path.isfile(os.path.join(path, f))]
		if not files:
			print(f'[!] No .xlsx files found at {path}')
			return None

		# Sort file list by creation time and get the newest
		newest_file = max(files, key=os.path.getctime)
		return newest_file

	final_path = descend_newest_path(basedir)
	if not final_path:
		print('[!] No .xlsx files found at newest directory.')

	newest_file = get_newest_file(final_path)
	if not newest_file:
		print(f'[!] No .xlsx files found at ')

	return newest_file

"""Find, convert, and print the newest XLSX file"""
def process_document():
    excel = None
    wb = None
    basedir = BASEDIR

    try:
        # Find the newest .xlsx file in the basedir.
        newest_file_path = get_newest_spreadsheet(basedir)
        if not newest_file_path:
            logger.error(f'[!] File not found, exiting!')
            raise FileNotFoundError

        logger.info(f'[+] Newest file found at {newest_file_path}')
        # safe_temp_dir = os.path.expanduser(TMP_PATH)
        # os.makedirs(safe_temp_dir, exist_ok=True)

        # temp_xlsx_path = os.path.join(safe_temp_dir, str(uuid4()) + ".xlsx")
        # logger.info(f'[+] Creating temporary file at {temp_xlsx_path}')
        # temp_xlsx_path = os.path.join(temp_file_path)
        # shutil.copy(newest_file_path, temp_xlsx_path)
        # os.close(temp_xlsx_path)  # close the file descriptor immediately to avoid locking issues
        temp_xlsx_path = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False).name
        
        logger.info(f'[+] Copying newest content to temporary file...')
        shutil.copy(newest_file_path, temp_xlsx_path)
        os.chmod(temp_xlsx_path, stat.S_IRUSR | stat.S_IWUSR)

        logger.info(f"[+] Checking file permissions...")
        logger.info(f"Exists? {os.path.exists(temp_xlsx_path)}")
        logger.info(f"Readable? {os.access(temp_xlsx_path, os.R_OK)}")
        logger.info(f"Writable? {os.access(temp_xlsx_path, os.W_OK)}")
        logger.info(f"Executable? {os.access(temp_xlsx_path, os.X_OK)}")
        logger.info(f"Stat: {os.stat(temp_xlsx_path)}")

        abs_path = os.path.abspath(temp_xlsx_path)
        
        pythoncom.CoInitialize()
        excel = win32.gencache.EnsureDispatch('Excel.Application')

        excel.Visible = False
        excel.DisplayAlerts = False

        if DEBUG:
            logger.info(f'[+] Opening Workbook...')
        wb = excel.Workbooks.Open(abs_path)

        wb.RefreshAll()
        excel.CalculateFull()

        # Fix the formatting
        sheet = wb.Sheets("Pallet")
        print_area = sheet.PageSetup.PrintArea

        # Pages setup object
        ps = sheet.PageSetup

        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = False

        sheet.Range("K3:L3").Font.Size = 7

        # Remove images other than the logo in the top left
        for shape in sheet.Shapes:
            #print(f"{shape.Name}: Top={shape.Top}, Left={shape.Left}, Width={shape.Width}, Height={shape.Height}")
            if shape.Type == 13 and shape.Name == "图片 2":
                shape.Placement = 2  # xlMoveAndSize = 1, xlMove = 2, xlFreeFloating = 3
                shape.LockAspectRatio = True
                shape.ScaleWidth(0.5, True)
                shape.ScaleHeight(0.5, True)
            elif shape.Type == 13:
                shape.Placement = 2  # xlMoveAndSize = 1, xlMove = 2, xlFreeFloating = 3
                shape.ScaleWidth(0, True)
                shape.ScaleHeight(0, True)

        logger.info(f'[+] Printing {abs_path}')
        print(f'Printing: {newest_file_path}')
        #excel.CalculateFull()
        if DEBUG:
            logger.warning(f'[-] DEBUG is enabled, set to false to enable printing.')
            raise Exception

        if conf['PRINT_METHOD'] == 'excel':
            wb.PrintOut()
            wb.PrintOut()
        else:            
            wb.ExportAsFixedFormat(
                Type=0,  # 0 = PDF, 1 = XPS
                Filename=TMP_PATH,
                Quality=0,  # 0 = Standard, 1 = Minimum
                IncludeDocProperties=True,
                IgnorePrintAreas=False,  # respect Print_Area
                OpenAfterPublish=False
            )
            os.startfile(TMP_PATH, "print")
            os.startfile(TMP_PATH, "print")
    except Exception as e:
        logger.error(e)
    finally:
        logger.info(f'[+] Closing WorkBook')
        try:
            wb.Close(False)
            excel.Quit()
        except:
            pass
        try:
            if os.path.exists(temp_xlsx_path):
                os.remove(temp_xlsx_path)
        except Exception as e:
            logger.error(e)
            sleep(5)
        pythoncom.CoUninitialize()

def main():
    # TODO: Ensure that once a second is fast enough
    print('[+] Starting')
    logger.info('[+] Started')

    path_to_watch = BASEDIR
    event_handler = NewFileHandler()
    observer = Observer()
    observer.schedule(event_handler, path=path_to_watch, recursive=False)
    observer.start()

    try:
        while True:
            sleep(POLL_DURATION)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == '__main__':
	main()
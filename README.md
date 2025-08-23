# Palletizer Print Automation

## Intro

This project uses the Python win32com library (pip install pywin32) to format an Excel workbook before printing it to the system's default printer.

## Build and Installation

Install python dependencies:

```bash
./venv/Scripts/activate
python3 -m pip install ./requirements.txt
python3 -m pip install pyinstaller
pyinstaller.exe --hidden-import=pywin32 --hidden-import=win32com --onefile ./main.py
```

The EXE file will be in the ./dist folder. Copy it and replace main.exe, then create a shortcut and press:

```
Win+R
shell:startup
```

Once the startup folder opens copy the shortcut into here.

Double click the free3of9.tff file and click "Install". Make sure Excel is installed on the system as well. These are the only two dependencies.

## Configuration

To modify the execution of the program, modify the `settings.json` file. An example is below:

```json
{
  "BASEDIR": "L:/csv/Palletizer/outbox/",
  "TMP_PATH": "C:/Users/operator/Desktop/2025-004/tmp.pdf",
  "POLL_DELAY_SECONDS": 1,
  "PRINT_METHOD": "excel",
  "PRINT_METHOD_2": "system",
  "LOGFILE": "prod.log"
}
```

The BASEDIR is the directory that is searched for new XLSX files. \

The TMP_PATH defines where to drop a PDF file that is generated for the printing process. Note the username in C:\Users\...\Desktop\ as this needs to change if you switch the user accounts.

The POLL_DELAY_SECONDS defines the time between filesystem checks in seconds. This can be a decimal if you need subsecond polling.

The PRINT_METHOD and PRINT_METHOD_2 options allow you to switch between using Excel's built-in print method and the default Windows print method. Whichever option is labeled PRINT_METHOD will be used, the other gets ignored.

LOGFILE defines the logfile path, which is in the same directory as the executable by default.

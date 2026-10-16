@echo off
rem Folder Session Manager launcher.
rem Uses the Python that has PySide6, pywin32, comtypes and pywinauto.
set "PY=C:\Python314\python.exe"
set "DIR=C:\AI_WORKSPACE\FOLDERS_SESSION_MANAGER"
rem Picks the newest version file in the folder.
set "APP="
for /f "delims=" %%f in ('dir /b /o:-n "%DIR%\FOLDERS_SESSION_MANAGER_v*.py"') do if not defined APP set "APP=%DIR%\%%f"
if not defined APP (
    echo No FOLDERS_SESSION_MANAGER_v*.py found in %DIR%
    pause
    exit /b 1
)

rem Hides the two startup console messages (OleInitialize and DPI awareness).
set "QT_LOGGING_RULES=qt.qpa.window=false;qt.default=false"

if not exist "%PY%" (
    echo Python not found: %PY%
    echo Install the modules and run: python "%APP%"
    pause
    exit /b 1
)

"%PY%" -c "import PySide6, win32com.client, comtypes, pywinauto" 2>nul
if errorlevel 1 (
    echo Missing modules. Install them with:
    echo   "%PY%" -m pip install PySide6 pywin32 comtypes pywinauto
    echo The app still runs, but UIA window reading and tab insertion are off.
)

"%PY%" "%APP%"
if errorlevel 1 pause

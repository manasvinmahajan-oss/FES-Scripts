@echo off
REM ====================================================================
REM FES Bids Manager - Launcher
REM Uses the correct Anaconda Python installation
REM ====================================================================

title FES Bids Manager

REM Change to the script directory
cd /d "%~dp0"

REM Check if Python file exists
if not exist "FES_Bids_Runner_PRODUCTION.py" (
    echo ERROR: FES_Bids_Runner_PRODUCTION.py not found!
    echo Make sure all Python files are in the same folder as this .bat file
    pause
    exit /b 1
)

REM Use the correct Anaconda Python path
set "PYTHON_PATH=C:\Users\enrolment\AppData\Local\anaconda3\pythonw.exe"

REM Check if Python exists at this location
if not exist "%PYTHON_PATH%" (
    echo ERROR: Python not found at expected location:
    echo %PYTHON_PATH%
    echo.
    echo Anaconda may have been reinstalled or moved.
    echo Please contact support.
    pause
    exit /b 1
)

REM Launch the GUI silently (no console window after this)
start "" "%PYTHON_PATH%" FES_Bids_Runner_PRODUCTION.py

REM Exit immediately (console window closes)
exit

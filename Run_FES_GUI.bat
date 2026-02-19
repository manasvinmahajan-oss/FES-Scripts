@echo off
REM ====================================================================
REM FES Bids Manager - Direct Launcher
REM Double-click this file to launch the GUI without opening Anaconda
REM ====================================================================

title FES Bids Manager

echo.
echo ========================================
echo   Starting FES Bids Manager...
echo ========================================
echo.

REM Change to the script directory (where this .bat file is located)
cd /d "%~dp0"

REM Try to find and use Anaconda Python directly
set "ANACONDA_PATH=%USERPROFILE%\anaconda3"
set "ANACONDA_PATH2=C:\Users\enrolment\AppData\Local\anaconda3"
set "ANACONDA_PATH3=C:\ProgramData\Anaconda3"

REM Check which Anaconda installation exists and run directly
if exist "%ANACONDA_PATH%\python.exe" (
    echo Found Anaconda at: %ANACONDA_PATH%
    echo Launching GUI...
    echo.
    start "" "%ANACONDA_PATH%\pythonw.exe" FES_Bids_Runner_PRODUCTION.py
    exit
)

if exist "%ANACONDA_PATH2%\python.exe" (
    echo Found Anaconda at: %ANACONDA_PATH2%
    echo Launching GUI...
    echo.
    start "" "%ANACONDA_PATH2%\pythonw.exe" FES_Bids_Runner_PRODUCTION.py
    exit
)

if exist "%ANACONDA_PATH3%\python.exe" (
    echo Found Anaconda at: %ANACONDA_PATH3%
    echo Launching GUI...
    echo.
    start "" "%ANACONDA_PATH3%\pythonw.exe" FES_Bids_Runner_PRODUCTION.py
    exit
)

REM If Anaconda not found at default locations, try system Python
echo Anaconda not found at default locations.
echo Trying system Python...
echo.

where python >nul 2>&1
if %ERRORLEVEL% equ 0 (
    start "" pythonw.exe FES_Bids_Runner_PRODUCTION.py
    exit
)

REM If nothing works, show error
echo.
echo ========================================
echo   ERROR: Could not find Python
echo ========================================
echo.
echo Please ensure Anaconda or Python is installed.
echo.
echo Expected Anaconda locations:
echo   - %USERPROFILE%\anaconda3
echo   - C:\Users\enrolment\AppData\Local\anaconda3
echo   - C:\ProgramData\Anaconda3
echo.
pause
exit /b 1

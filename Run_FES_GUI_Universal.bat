@echo off
REM ====================================================================
REM FES Bids Manager - Universal Launcher
REM Works on any computer with Python/Anaconda installed
REM ====================================================================

title FES Bids Manager

echo.
echo ========================================
echo   Starting FES Bids Manager...
echo ========================================
echo.

REM Change to the script directory (where this .bat file is located)
cd /d "%~dp0"

REM ============================================================
REM Method 1: Try pythonw in system PATH (works if Python is in PATH)
REM ============================================================
where pythonw >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo Found Python in system PATH
    echo Launching GUI...
    echo.
    start "" pythonw.exe FES_Bids_Runner_PRODUCTION.py
    exit
)

REM ============================================================
REM Method 2: Search common Anaconda locations for ANY user
REM ============================================================
echo Searching for Anaconda installation...
echo.

REM Current user's home directory
if exist "%USERPROFILE%\anaconda3\pythonw.exe" (
    echo Found Anaconda at: %USERPROFILE%\anaconda3
    echo Launching GUI...
    echo.
    start "" "%USERPROFILE%\anaconda3\pythonw.exe" FES_Bids_Runner_PRODUCTION.py
    exit
)

if exist "%USERPROFILE%\Anaconda3\pythonw.exe" (
    echo Found Anaconda at: %USERPROFILE%\Anaconda3
    echo Launching GUI...
    echo.
    start "" "%USERPROFILE%\Anaconda3\pythonw.exe" FES_Bids_Runner_PRODUCTION.py
    exit
)

REM AppData Local
if exist "%LOCALAPPDATA%\anaconda3\pythonw.exe" (
    echo Found Anaconda at: %LOCALAPPDATA%\anaconda3
    echo Launching GUI...
    echo.
    start "" "%LOCALAPPDATA%\anaconda3\pythonw.exe" FES_Bids_Runner_PRODUCTION.py
    exit
)

if exist "%LOCALAPPDATA%\Continuum\anaconda3\pythonw.exe" (
    echo Found Anaconda at: %LOCALAPPDATA%\Continuum\anaconda3
    echo Launching GUI...
    echo.
    start "" "%LOCALAPPDATA%\Continuum\anaconda3\pythonw.exe" FES_Bids_Runner_PRODUCTION.py
    exit
)

REM System-wide installations
if exist "C:\ProgramData\Anaconda3\pythonw.exe" (
    echo Found Anaconda at: C:\ProgramData\Anaconda3
    echo Launching GUI...
    echo.
    start "" "C:\ProgramData\Anaconda3\pythonw.exe" FES_Bids_Runner_PRODUCTION.py
    exit
)

if exist "C:\Anaconda3\pythonw.exe" (
    echo Found Anaconda at: C:\Anaconda3
    echo Launching GUI...
    echo.
    start "" "C:\Anaconda3\pythonw.exe" FES_Bids_Runner_PRODUCTION.py
    exit
)

REM ============================================================
REM Method 3: Try to find Python using Windows Registry
REM ============================================================
echo Searching Windows registry for Python...
echo.

for /f "tokens=*" %%i in ('reg query "HKEY_CURRENT_USER\Software\Python\PythonCore" /s /f "InstallPath" 2^>nul ^| findstr "InstallPath"') do (
    for /f "tokens=2*" %%a in ('reg query "%%i" /ve 2^>nul') do (
        if exist "%%b\pythonw.exe" (
            echo Found Python at: %%b
            echo Launching GUI...
            echo.
            start "" "%%b\pythonw.exe" FES_Bids_Runner_PRODUCTION.py
            exit
        )
    )
)

REM ============================================================
REM Method 4: Try standard python command
REM ============================================================
where python >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo Found Python via 'python' command
    echo Launching GUI...
    echo.
    start "" python.exe FES_Bids_Runner_PRODUCTION.py
    exit
)

REM ============================================================
REM Method 5: Try py launcher (comes with Python 3.3+)
REM ============================================================
where py >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo Found Python via 'py' launcher
    echo Launching GUI...
    echo.
    start "" py.exe -3 FES_Bids_Runner_PRODUCTION.py
    exit
)

REM ============================================================
REM If nothing works, show helpful error
REM ============================================================
echo.
echo ========================================
echo   ERROR: Could not find Python
echo ========================================
echo.
echo Python/Anaconda was not found on this computer.
echo.
echo Please ensure one of the following is installed:
echo   1. Anaconda (recommended)
echo   2. Python 3.x
echo.
echo Searched locations:
echo   - System PATH
echo   - %USERPROFILE%\anaconda3
echo   - %LOCALAPPDATA%\anaconda3
echo   - C:\ProgramData\Anaconda3
echo   - C:\Anaconda3
echo   - Windows Registry
echo.
echo If Python/Anaconda IS installed, you may need to:
echo   1. Add Python to system PATH, or
echo   2. Reinstall Anaconda with "Add to PATH" option
echo.
pause
exit /b 1

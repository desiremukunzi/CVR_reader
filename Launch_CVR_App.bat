@echo off
title CVR Audio Compliance Checker

echo ============================================
echo   CVR Audio Compliance Checker - Starting
echo ============================================
echo.

:: Navigate to the app directory
cd /d A:\CVR_reader

:: Check if directory exists
if not exist "A:\CVR_reader" (
    echo ERROR: Folder A:\CVR_reader not found!
    pause
    exit /b 1
)

:: Check if virtual environment exists
if not exist "venv-py311\Scripts\activate.bat" (
    echo ERROR: Virtual environment not found at venv-py311\Scripts\activate.bat
    pause
    exit /b 1
)

:: Activate the virtual environment
echo [1/3] Activating virtual environment...
call venv-py311\Scripts\activate.bat

:: Check if app.py exists
if not exist "app.py" (
    echo ERROR: app.py not found in A:\CVR_reader
    pause
    exit /b 1
)

:: Start the Flask app in the background
echo [2/3] Starting Flask server...
start "CVR Flask Server" /min cmd /k "cd /d A:\CVR_reader && call venv-py311\Scripts\activate.bat && python app.py"

:: Poll until Flask is actually ready (max 60 seconds)
echo [3/3] Waiting for server to be ready...
set /a TRIES=0
:WAIT_LOOP
set /a TRIES+=1
if %TRIES% GTR 60 (
    echo ERROR: Server did not start after 60 seconds.
    pause
    exit /b 1
)
powershell -Command "try { Invoke-WebRequest -Uri 'http://127.0.0.1:5000' -UseBasicParsing -TimeoutSec 1 -ErrorAction Stop | Out-Null; exit 0 } catch { exit 1 }" >nul 2>&1
if errorlevel 1 (
    timeout /t 1 /nobreak >nul
    goto WAIT_LOOP
)

:: Flask responded - open the browser
echo Opening browser at http://127.0.0.1:5000
start "" "http://127.0.0.1:5000"

echo.
echo ============================================
echo   App is running!
echo   Browser opened at: http://127.0.0.1:5000
echo   Close the Flask Server window to stop.
echo ============================================

:: Close this launcher window after 3 seconds
timeout /t 3 /nobreak >nul
exit

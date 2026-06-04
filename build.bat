@echo off
title PPTX Extractor - Build

echo.
echo =====================================================
echo   PPTX Text Extractor  ^|  EXE Build Script
echo =====================================================
echo.

set SCRIPT_NAME=pptx_text_extractor.py
set SCRIPT_URL=https://raw.githubusercontent.com/kwkim-gif/text/claude/keen-davinci-QBEal/pptx_text_extractor.py
set SCRIPT_DEST=%~dp0%SCRIPT_NAME%

echo Downloading latest %SCRIPT_NAME% from GitHub...
powershell -NoProfile -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;(New-Object Net.WebClient).DownloadFile('%SCRIPT_URL%','%SCRIPT_DEST%')"
if not exist "%SCRIPT_DEST%" (
    echo [ERROR] Failed to download %SCRIPT_NAME%.
    pause
    exit /b 1
)
echo [OK] %SCRIPT_NAME% ready.
echo.

python --version > nul 2>&1
if errorlevel 1 goto NO_PYTHON
goto HAS_PYTHON

:NO_PYTHON
echo [WARN] Python not found. Downloading Python 3.11.9...
echo.

set PY_VER=3.11.9
set PY_FILE=python-%PY_VER%-amd64.exe
set PY_URL=https://www.python.org/ftp/python/%PY_VER%/%PY_FILE%
set PY_DEST=%TEMP%\%PY_FILE%

echo Downloading from: %PY_URL%
echo Please wait...
echo.

powershell -NoProfile -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;(New-Object Net.WebClient).DownloadFile('%PY_URL%','%PY_DEST%')"

if not exist "%PY_DEST%" (
    echo [ERROR] Download failed. Check your internet connection.
    pause
    exit /b 1
)

echo Download complete. Installing silently...
"%PY_DEST%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0

if errorlevel 1 (
    echo [ERROR] Python install failed.
    del /q "%PY_DEST%" 2>nul
    pause
    exit /b 1
)

del /q "%PY_DEST%" 2>nul
echo [OK] Python installed.
echo.

set PYPATH=%LOCALAPPDATA%\Programs\Python\Python311
set PATH=%PYPATH%;%PYPATH%\Scripts;%PATH%

:HAS_PYTHON
for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PY_VER_STR=%%i
echo [OK] %PY_VER_STR% ready
echo.

echo [1/3] Installing packages...
python -m pip install python-pptx tkinterdnd2 Pillow pymupdf pyinstaller --quiet --disable-pip-version-check
if errorlevel 1 (
    echo [ERROR] Package installation failed.
    pause
    exit /b 1
)
echo       Done
echo.

echo [2/3] Cleaning previous build...
if exist "dist"  rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "PPTX_Utility.spec" del /q "PPTX_Utility.spec"
echo       Done
echo.

echo [3/3] Building EXE (may take 1-2 minutes)...
echo.
python -m PyInstaller --onefile --windowed --name "PPTX_Utility" --collect-all tkinterdnd2 --collect-all fitz "%SCRIPT_DEST%"

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo.
echo =====================================================
echo   Build complete!
echo   Output: dist\PPTX_Utility.exe
echo =====================================================
echo.
explorer dist
pause

@echo off
title PPTX Extractor - Build

echo.
echo =====================================================
echo   PPTX Text Extractor  ^|  EXE Build Script
echo =====================================================
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
python -m pip install python-pptx pyinstaller --quiet --disable-pip-version-check
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
if exist "PPTX_TextExtractor.spec" del /q "PPTX_TextExtractor.spec"
echo       Done
echo.

echo [3/3] Building EXE (may take 1-2 minutes)...
echo.
python -m PyInstaller --onefile --windowed --name "PPTX_TextExtractor" pptx_text_extractor.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo.
echo =====================================================
echo   Build complete!
echo   Output: dist\PPTX_TextExtractor.exe
echo =====================================================
echo.
explorer dist
pause

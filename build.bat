@echo off
title PPTX Extractor - Build

echo.
echo =====================================================
echo   PPTX Text Extractor  ^|  EXE Build Script
echo =====================================================
echo.

:: Check Python installation
python --version > nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed.
    echo         Please install from https://www.python.org
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version') do set PY_VER=%%i
echo [OK] %PY_VER% detected
echo.

:: Install required packages
echo [1/3] Installing required packages...
pip install python-pptx pyinstaller --quiet --disable-pip-version-check
if errorlevel 1 (
    echo [ERROR] Package installation failed.
    pause
    exit /b 1
)
echo       Done
echo.

:: Clean previous build artifacts
echo [2/3] Cleaning previous build...
if exist "dist"  rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del /q "*.spec"
echo       Done
echo.

:: Build single EXE with PyInstaller
echo [3/3] Building EXE...  (may take 1-2 minutes)
echo.

pyinstaller --onefile --windowed --name "PPTX_TextExtractor" pptx_text_extractor.py

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

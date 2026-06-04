@echo off
title PPTX Extractor - Build

echo.
echo =====================================================
echo   PPTX Text Extractor  ^|  EXE Build Script
echo =====================================================
echo.

:: ── Step 0: Python 설치 여부 확인 후 없으면 자동 다운로드 ──────────
python --version > nul 2>&1
if errorlevel 1 (
    echo [WARN] Python not found. Auto-installing Python 3.11...
    echo.
    call :install_python
    if errorlevel 1 (
        echo [ERROR] Python installation failed. Please install manually.
        echo         https://www.python.org/downloads/
        pause
        exit /b 1
    )
)

:: 설치 후 PATH 반영을 위해 환경 변수 갱신
for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PY_VER=%%i
echo [OK] %PY_VER% ready
echo.

:: ── Step 1: 필요 패키지 설치 ─────────────────────────────────────
echo [1/3] Installing packages: python-pptx, pyinstaller...
python -m pip install python-pptx pyinstaller --quiet --disable-pip-version-check
if errorlevel 1 (
    echo [ERROR] Package installation failed.
    pause
    exit /b 1
)
echo       Done
echo.

:: ── Step 2: 이전 빌드 산출물 정리 ────────────────────────────────
echo [2/3] Cleaning previous build artifacts...
if exist "dist"   rmdir /s /q "dist"
if exist "build"  rmdir /s /q "build"
for %%f in (*.spec) do del /q "%%f"
echo       Done
echo.

:: ── Step 3: PyInstaller 로 단일 EXE 빌드 ─────────────────────────
echo [3/3] Building EXE (may take 1-2 minutes)...
echo.

python -m PyInstaller --onefile --windowed --name "PPTX_TextExtractor" pptx_text_extractor.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed. See above for details.
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
exit /b 0


:: ════════════════════════════════════════════════════
:: Python 자동 설치 서브루틴
:: PowerShell 로 공식 인스톨러를 다운로드 후 조용히 설치
:: ════════════════════════════════════════════════════
:install_python
setlocal

set PY_VERSION=3.11.9
set PY_INSTALLER=python-%PY_VERSION%-amd64.exe
set PY_URL=https://www.python.org/ftp/python/%PY_VERSION%/%PY_INSTALLER%
set PY_DEST=%TEMP%\%PY_INSTALLER%

:: 64bit / 32bit 구분
reg query "HKLM\HARDWARE\DESCRIPTION\System\CentralProcessor\0" /v Identifier | find "x86" > nul 2>&1
if not errorlevel 1 (
    set PY_INSTALLER=python-%PY_VERSION%.exe
    set PY_URL=https://www.python.org/ftp/python/%PY_VERSION%/python-%PY_VERSION%.exe
    set PY_DEST=%TEMP%\python-%PY_VERSION%.exe
)

echo Downloading Python %PY_VERSION% from python.org...
echo Source: %PY_URL%
echo.

:: PowerShell 로 다운로드 (Windows 7 이상 기본 내장)
powershell -NoProfile -Command ^
  "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;" ^
  "Write-Host 'Downloading...';" ^
  "(New-Object Net.WebClient).DownloadFile('%PY_URL%', '%PY_DEST%')" ^
  2>&1

if not exist "%PY_DEST%" (
    echo [ERROR] Download failed.
    endlocal
    exit /b 1
)

echo Download complete. Installing Python silently...
echo (This may take a minute)
echo.

:: 조용한 설치 옵션:
::   InstallAllUsers=0  -> 현재 사용자만 설치 (관리자 권한 불필요)
::   PrependPath=1      -> PATH 에 자동 등록
::   Include_test=0     -> 테스트 파일 제외 (용량 절약)
::   Include_launcher=1 -> py 런처 포함
"%PY_DEST%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_launcher=1

if errorlevel 1 (
    echo [ERROR] Python installer returned an error.
    del /q "%PY_DEST%" 2>nul
    endlocal
    exit /b 1
)

del /q "%PY_DEST%" 2>nul

:: 새로 등록된 PATH 를 현재 세션에 반영
for /f "tokens=2*" %%a in (
    'reg query "HKCU\Environment" /v PATH 2^>nul'
) do set "USER_PATH=%%b"

if defined USER_PATH (
    set "PATH=%USER_PATH%;%PATH%"
)

:: AppData 경로도 추가 (사용자 설치 기본 위치)
for /f "tokens=*" %%i in (
    'powershell -NoProfile -Command "[System.Environment]::GetFolderPath(\"LocalApplicationData\")"'
) do set "LOCALAPPDATA_PATH=%%i"

set "PATH=%LOCALAPPDATA_PATH%\Programs\Python\Python311;%LOCALAPPDATA_PATH%\Programs\Python\Python311\Scripts;%PATH%"

echo [OK] Python installed successfully.
echo.
endlocal & set "PATH=%PATH%"
exit /b 0

@echo off
chcp 65001 > nul
title PPTX 텍스트 추출기 - 빌드 스크립트

echo.
echo =====================================================
echo   PPTX 텍스트 추출기  ^|  EXE 빌드 스크립트
echo =====================================================
echo.

:: ── Python 설치 여부 확인 ────────────────────────────────
python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python 이 설치되어 있지 않습니다.
    echo        https://www.python.org 에서 설치 후 다시 실행하세요.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version') do set PY_VER=%%i
echo [OK] %PY_VER% 확인됨
echo.

:: ── 필요 패키지 설치 ─────────────────────────────────────
echo [1/3] 필요 패키지 설치 중...
pip install python-pptx pyinstaller --quiet --disable-pip-version-check
if errorlevel 1 (
    echo [오류] 패키지 설치에 실패했습니다.
    pause
    exit /b 1
)
echo       완료
echo.

:: ── 이전 빌드 산출물 정리 ─────────────────────────────────
echo [2/3] 이전 빌드 정리 중...
if exist "dist"  rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del /q "*.spec"
echo       완료
echo.

:: ── PyInstaller 로 단일 EXE 빌드 ─────────────────────────
echo [3/3] EXE 파일 빌드 중...  (1~2분 소요될 수 있습니다)
echo.

pyinstaller ^
    --onefile ^
    --windowed ^
    --name "PPTX텍스트추출기" ^
    --add-data "pptx_text_extractor.py;." ^
    pptx_text_extractor.py

if errorlevel 1 (
    echo.
    echo [오류] 빌드에 실패했습니다.
    pause
    exit /b 1
)

:: ── 완료 안내 ─────────────────────────────────────────────
echo.
echo =====================================================
echo   빌드 완료!
echo   실행 파일 위치: dist\PPTX텍스트추출기.exe
echo =====================================================
echo.
echo   이 EXE 파일만 복사해서 배포하면 됩니다.
echo   (Python 설치 없이 어디서든 실행 가능)
echo.

:: 빌드된 dist 폴더를 탐색기로 열기
explorer dist

pause

@echo off
chcp 65001 > nul
title AI 블로그 제목 생성기

echo ====================================================
echo         🤖 AI 블로그 제목 생성기 실행
echo ====================================================
echo.

echo 📦 Python 환경 확인 중...
python --version 2>nul
if %errorlevel% neq 0 (
    echo ❌ Python이 설치되지 않았습니다.
    echo    Python을 설치하고 다시 시도해주세요.
    echo    https://www.python.org/downloads/
    pause
    exit /b 1
)

echo ✅ Python 환경 확인됨
echo.

echo 📋 필수 패키지 설치 확인 중...
python -c "import PyQt6, openai, pandas, openpyxl, requests" 2>nul
if %errorlevel% neq 0 (
    echo ⚠️ 필수 패키지가 설치되지 않았습니다.
    echo    설치를 시작합니다...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo ❌ 패키지 설치에 실패했습니다.
        pause
        exit /b 1
    )
)

echo ✅ 필수 패키지 확인됨
echo.

echo 🚀 AI 블로그 제목 생성기를 시작합니다...
python blog_title_generator.py

if %errorlevel% neq 0 (
    echo.
    echo ❌ 프로그램 실행 중 오류가 발생했습니다.
    echo    오류 코드: %errorlevel%
    pause
)

echo.
echo 👋 프로그램이 종료되었습니다.
pause 
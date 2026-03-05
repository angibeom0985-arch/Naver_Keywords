@echo off
echo ====================================================
echo       🔓 파일 보호 해제 스크립트 (관리자 전용)
echo ====================================================
echo.
echo ⚠️  이 스크립트는 보호된 파일들을 해제합니다.
echo 계속하시겠습니까? (y/n)
set /p choice=선택: 
if /i not "%choice%"=="y" (
    echo 취소되었습니다.
    pause
    exit /b
)

echo.
echo 🔓 keyword_results 폴더 보호 해제 중...
if exist "keyword_results" (
    attrib -S -H -R "keyword_results" /S /D
    echo ✅ keyword_results 폴더 보호 해제 완료
) else (
    echo ⚠️ keyword_results 폴더가 존재하지 않습니다.
)

echo.
echo 🔓 dist 폴더 보호 해제 중...
if exist "dist" (
    attrib -S -H -R "dist" /S /D
    echo ✅ dist 폴더 보호 해제 완료
) else (
    echo ⚠️ dist 폴더가 존재하지 않습니다.
)

echo.
echo ✅ 모든 파일 보호가 해제되었습니다.
echo 이제 파일들을 삭제하거나 수정할 수 있습니다.
echo.
pause

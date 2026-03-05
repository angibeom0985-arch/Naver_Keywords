@echo off
chcp 65001 > nul
echo ====================================================
echo       🛡️ 파일 보호 스크립트 (간단 버전)
echo ====================================================
echo.

echo 📁 keyword_results 폴더 보호 중...
if exist "keyword_results" (
    echo 🔒 keyword_results 폴더를 읽기 전용으로 설정합니다...
    attrib +R "keyword_results" /S /D
    echo ✅ keyword_results 폴더 보호 완료
    
    echo 📊 보호된 파일 목록:
    dir "keyword_results" /B
) else (
    echo ⚠️ keyword_results 폴더가 존재하지 않습니다.
)

echo.
echo 📁 dist 폴더 보호 중...
if exist "dist" (
    echo 🔒 dist 폴더를 읽기 전용으로 설정합니다...
    attrib +R "dist" /S /D
    echo ✅ dist 폴더 보호 완료
    
    echo 📊 보호된 파일 목록:
    dir "dist" /B
) else (
    echo ⚠️ dist 폴더가 존재하지 않습니다.
)

echo.
echo 📄 보호 해제 스크립트 생성 중...
echo @echo off > unprotect_simple.bat
echo echo 🔓 파일 보호 해제 중... >> unprotect_simple.bat
echo if exist "keyword_results" attrib -R "keyword_results" /S /D >> unprotect_simple.bat
echo if exist "dist" attrib -R "dist" /S /D >> unprotect_simple.bat
echo echo ✅ 보호 해제 완료 >> unprotect_simple.bat
echo pause >> unprotect_simple.bat

echo ✅ 보호 해제 스크립트가 생성되었습니다: unprotect_simple.bat

echo.
echo ====================================================
echo 🎉 파일 보호 완료!
echo ====================================================
echo.
echo 📋 보호된 내용:
echo   • keyword_results 폴더와 모든 하위 파일
echo   • dist 폴더와 모든 하위 파일 (있는 경우)
echo   • 읽기 전용 속성으로 보호
echo.
echo 🔓 보호 해제 방법:
echo   1. unprotect_simple.bat 실행
echo   2. 또는 명령 프롬프트에서:
echo      attrib -R keyword_results /S /D
echo      attrib -R dist /S /D
echo.
echo ⚠️  주의: 이 보호는 기본적인 수준입니다.
echo    완전한 보호를 위해서는 관리자 권한이 필요합니다.
echo.
pause 
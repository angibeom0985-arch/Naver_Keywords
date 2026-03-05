@echo off 
echo 🔓 파일 보호 해제 중... 
if exist "keyword_results" attrib -R "keyword_results" /S /D 
if exist "dist" attrib -R "dist" /S /D 
echo ✅ 보호 해제 완료 
pause 

@echo off
set PYTHONIOENCODING=cp949:replace
cd /d "%~dp0"
echo  이미 받아 둔 폴더로 집계/보고서만 다시 만듭니다.
echo  사용:  집계만.bat 2026-08-31
echo.
py run_daily.py --no-collect --date %1
pause

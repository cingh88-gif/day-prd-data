@echo off
rem Windows 작업 스케줄러가 매일 부르는 진입점. 사람이 볼 창이 없으므로 pause 를 넣지 않는다.
rem 로그는 C:\ierp_exports\day_prd\logs\ 에 쌓인다.
set PYTHONIOENCODING=cp949:replace
cd /d "%~dp0"
py run_daily.py --auto
exit /b %errorlevel%

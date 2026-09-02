@echo off
rem Windows 작업 스케줄러가 매일 부르는 진입점.
rem 사람이 볼 창이 없으므로 pause 를 넣지 않는다. 로그는 C:\ierp_exports\day_prd\logs\ 에.
rem 실패하면 10분 뒤 이어받기로 다시 시도한다(최대 3회) - 사용자가 자리에 있거나
rem 다른 창이 화면을 덮고 있으면 한 번에 안 될 수 있다.
set PYTHONIOENCODING=cp949:replace
cd /d "%~dp0"
py run_daily.py --auto --retry 3 --retry-wait 10
exit /b %errorlevel%

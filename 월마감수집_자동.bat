@echo off
rem 작업 스케줄러가 매월 부르는 진입점 - 전월 전체를 받는다.
rem 실패하면 15분 뒤 이어받기로 재시도(최대 3회). 한 달치라 시간이 더 걸린다.
set PYTHONIOENCODING=cp949:replace
cd /d "%~dp0"
py run_daily.py --auto --last-month --retry 3 --retry-wait 15
exit /b %errorlevel%

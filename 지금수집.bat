@echo off
set PYTHONIOENCODING=cp949:replace
cd /d "%~dp0"
echo ============================================================
echo  일마감 공정진척 수집 - 전날치 (수집 -^> 집계 -^> 보고서)
echo ============================================================
echo.
echo  * iERP 에 로그인돼 iEMenu 가 떠 있어야 합니다.
echo  * 수집 중에는 마우스/키보드를 만지지 마세요.
echo  * 중단돼도 같은 날짜로 다시 실행하면 이어받습니다.
echo.
pause
py run_daily.py %*
echo.
pause

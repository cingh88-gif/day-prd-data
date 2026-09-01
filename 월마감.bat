@echo off
set PYTHONIOENCODING=cp949:replace
cd /d "%~dp0"
echo ============================================================
echo  월마감 공정진척 수집 - 한 달치 (수집 -^> 집계 -^> 보고서)
echo ============================================================
echo.
echo   사용:  월마감.bat 2026-08      (해당 월)
echo          월마감.bat              (전월)
echo.
echo  * iERP 에 로그인돼 iEMenu 가 떠 있어야 합니다.
echo  * 수집 중에는 마우스/키보드를 만지지 마세요.
echo  * 한 달치는 조회가 오래 걸립니다. 중단돼도 다시 실행하면 이어받습니다.
echo.
pause
if "%~1"=="" (
    py run_daily.py --last-month
) else (
    py run_daily.py --month %1
)
echo.
pause

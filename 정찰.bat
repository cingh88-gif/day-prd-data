@echo off
set PYTHONIOENCODING=cp949:replace
cd /d "%~dp0"
echo ============================================================
echo  공정진척현황(PM60250Rv3) 정찰 - 화면 컨트롤 실측
echo ============================================================
echo.
echo  * iERP 에 로그인돼 iEMenu 가 떠 있어야 합니다.
echo  * ★ 조회를 누르기 전 상태에서 돌리세요.
echo    (그리드에 행이 차면 컨트롤 탐색이 63초 뒤 실패합니다)
echo.
pause
py probe_progress.py
echo.
echo  결과: ierp_inspect\probe_progress_출력.txt  /  screen.json
pause

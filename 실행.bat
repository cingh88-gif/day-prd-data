@echo off
rem 공정진척현황(PM60250Rv3) 일마감 수집 - GUI 실행
set PYTHONIOENCODING=cp949:replace
cd /d "%~dp0"
start "" pyw gui.py

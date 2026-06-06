@echo off
cd /d "%~dp0"
"..\..\.venv\Scripts\python.exe" "..\..\_templates_machine_.py" "config_Auto.xlsx" all all
pause

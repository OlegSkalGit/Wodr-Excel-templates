@echo off
cd /d "%~dp0"
"..\..\.venv\Scripts\python.exe" "..\..\_templates_machine_.py" "w_config.xlsx" all all
pause

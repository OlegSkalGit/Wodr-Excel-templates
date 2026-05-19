@echo off
cd /d "%~dp0"
"..\.venv\Scripts\python.exe" "..\_templates_machine_.py" "Auto_Config.xlsx" all all
pause

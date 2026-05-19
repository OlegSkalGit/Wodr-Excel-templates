@echo off
cd /d "%~dp0"
chcp 1251 > nul

title TemplateMachine Control Center
echo ========================================================
echo  TEMPLATEMACHINE CONTROL CENTER
echo ========================================================
echo.

REM 1. Перевірка наявності Python у PATH
python --version >nul 2>nul
if %ERRORLEVEL% neq 0 (
    if not exist "setup\python-3.14.5-amd64.exe" (
        echo [Встановлення] Завантажуємо інсталятор Python...
        if not exist "setup" mkdir "setup"
        curl -L "https://www.python.org/ftp/python/3.14.5/python-3.14.5-amd64.exe" -o "setup\python-3.14.5-amd64.exe"
    )
    echo Python не знайдено у системі.
    echo Запуск тихого встановлення Python...
    .\setup\python-3.14.5-amd64.exe /quiet PrependPath=1 InstallAllUsers=1
    echo Очікування завершення встановлення...
    timeout /t 5 > nul
)

REM 2. Створення віртуального оточення
if not exist ".venv" (
    echo [!] Віртуальне оточення .venv відсутнє. Створюємо...
    echo.
    python -m venv .venv
)

REM 3. Перевірка наявності необхідних бібліотек
.venv\Scripts\python.exe -c "import streamlit, pandas, openpyxl, docxtpl, transformers, torch, accelerate" >nul 2>nul
if %ERRORLEVEL% == 0 goto skipbase
echo.
echo [!] Встановлення та оновлення бібліотек (streamlit, pandas, torch, transformers тощо)...
.venv\Scripts\pip install streamlit pandas openpyxl docxtpl transformers torch accelerate
if errorlevel 1 (
    echo [Помилка] Не вдалося встановити бібліотеки. Перевірте з'єднання з інтернетом.
    pause
    exit /b 1
)
:skipbase

REM 5. Запуск Streamlit
echo.
echo [Успішно] Все готово до запуску!
echo Запуск веб-інтерфейсу...
echo.
.venv\Scripts\streamlit run app.py

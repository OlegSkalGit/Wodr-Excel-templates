@echo off
cd /d "%~dp0"
chcp 1251 > nul

title Панель керування TemplateMachine
echo ========================================================
echo   ЗАПУСК ПАНЕЛІ КЕРУВАННЯ TEMPLATEMACHINE
echo ========================================================
echo.

REM 1. Перевірка наявності віртуального середовища
if not exist ".venv" (
    echo [Увага] Віртуальне середовище .venv не знайдено!
    echo Запуск встановлення залежностей через інсталятор...
    echo.
    call .\setup\INSTALL.bat
)

REM 2. Перевірка наявності необхідних бібліотек
echo Перевірка встановлених компонентів інтерфейсу...
.venv\Scripts\python.exe -c "import streamlit, pandas, openpyxl, docxtpl" >nul 2>nul
if %ERRORLEVEL% == 0 goto pylibok
echo.
echo [Інфо] Встановлюються додаткові бібліотеки для інтерфейсу (streamlit, pandas)...
.venv\Scripts\pip install streamlit pandas openpyxl docxtpl
if errorlevel 1 (
    echo [Помилка] Не вдалося встановити бібліотеки. Будь ласка, перевірте з'єднання з Інтернетом.
    pause
    exit /b 1
)

:pylibok

REM 3. Запуск Streamlit
echo.
echo [Успішно] Усі компоненти готові!
echo Запуск локального сервера...
echo Панель автоматично відкриється у вашому браузері.
echo.
echo Для завершення роботи просто закрийте це чорне вікно консолі.
echo ========================================================
echo.
.venv\Scripts\streamlit run app.py

pause

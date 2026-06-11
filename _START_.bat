@echo off
cd /d "%~dp0"
chcp 1251 > nul

setlocal enabledelayedexpansion

set "PY_VERSION=3.12.10"
set "PY_DIR_VER=312"

title TemplateMachine Control Center
echo ========================================================
echo  TEMPLATEMACHINE CONTROL CENTER
echo ========================================================
echo.

set "PYTHON_CMD=python"

REM 1. Перевірка наявності Python
python --version >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Python НЕ знайдено у системі.
    if not exist "setup\python-%PY_VERSION%-amd64.exe" (
        echo Інсталятор Python не знайдено. Завантажуємо версію %PY_VERSION%...
        if not exist "setup" mkdir "setup"
        curl -L "https://www.python.org/ftp/python/%PY_VERSION%/python-%PY_VERSION%-amd64.exe" -o "setup\python-%PY_VERSION%-amd64.exe"
        if !ERRORLEVEL! neq 0 (
            echo.
            echo ===============================================================================================
            echo Для ПЕРШОГО запуску програми необхідний інтернет для завантаження ДОДАТКОВИХ компонентів.
            echo ===============================================================================================
            echo.
            echo Не вдалося завантажити Python. Перевірте з'єднання з інтернетом.
            pause
            exit /b 1
        )
    )
    echo Інсталятор Python знайдено. Встановлюємо...
    start /wait "" ".\setup\python-%PY_VERSION%-amd64.exe" /quiet PrependPath=1 InstallAllUsers=0 Include_launcher=0 TargetDir="%LOCALAPPDATA%\Programs\Python\Python%PY_DIR_VER%"
    if !ERRORLEVEL! neq 0 (
        echo.
        echo ===============================================================================================
        echo Для ПЕРШОГО запуску програми необхідний інтернет для завантаження ДОДАТКОВИХ компонентів.
        echo ===============================================================================================
        echo.
        echo Не вдалося встановити Python. Перевірте з'єднання з інтернетом.
        pause
        exit /b 1
    )
    set "PYTHON_CMD="%LOCALAPPDATA%\Programs\Python\Python%PY_DIR_VER%\python.exe""
) else (
   echo Python знайдено у системі.
)

REM 2. Створення віртуального оточення
if not exist ".venv" (
    echo Віртуальне оточення Python .venv у проекті відсутнє. Створюємо...
    echo.
    !PYTHON_CMD! -m venv .venv
) else (
    echo Віртуальне оточення Python .venv у проекті знайдено.
)

REM 3. Перевірка наявності необхідних бібліотек
.venv\Scripts\python.exe -c "import streamlit, pandas, openpyxl, docxtpl" >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Необхідні бібліотеки Python НЕ знайдено у проекті. Завантажуємо і встановлюємо...
    .venv\Scripts\python.exe -m pip install -r requirements.txt
    if !ERRORLEVEL! neq 0 (
        echo.
        echo ===============================================================================================
        echo Для ПЕРШОГО запуску програми необхідний інтернет для завантаження ДОДАТКОВИХ компонентів.
        echo ===============================================================================================
        echo.
        echo [Помилка] Не вдалося встановити бібліотеки. Перевірте з'єднання з інтернетом.
        pause
        exit /b 1
    )
) else (
    echo Необхідні бібліотеки Python знайдено у проекті.
)

endlocal

REM 5. Запуск Streamlit
echo.
echo Все готово до запуску!
echo Запуск додатку...
echo.
".venv\Scripts\python.exe" -m streamlit run "app.py"
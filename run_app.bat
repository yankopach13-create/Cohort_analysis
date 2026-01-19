@echo off
chcp 65001 > nul
echo Запуск приложения когортного анализа...
echo.

REM Проверка наличия Python
python --version >nul 2>&1
if errorlevel 1 (
    echo Ошибка: Python не найден. Пожалуйста, установите Python.
    pause
    exit /b 1
)

REM Проверка наличия виртуального окружения
if exist "venv\Scripts\activate.bat" (
    echo Активация виртуального окружения...
    call venv\Scripts\activate.bat
) else (
    echo Создание виртуального окружения...
    python -m venv venv
    call venv\Scripts\activate.bat
    echo Установка зависимостей...
    pip install -r requirements.txt
)

REM Запуск Streamlit приложения
echo.
echo Запуск приложения...
echo.
streamlit run app.py

pause









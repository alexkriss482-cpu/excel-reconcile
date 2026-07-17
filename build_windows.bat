@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

cd /d "%~dp0"
echo ============================================
echo   Сборка SverkaOFD.exe для Windows
echo ============================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] Python не найден.
    echo Установите Python 3.10+ с https://www.python.org/downloads/
    echo При установке отметьте "Add python.exe to PATH".
    pause
    exit /b 1
)

echo [1/3] Установка зависимостей...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo [ОШИБКА] Не удалось установить зависимости.
    pause
    exit /b 1
)

echo.
echo [2/3] Сборка одного файла SverkaOFD.exe (1-3 минуты)...
python -m PyInstaller --noconfirm --clean SverkaOFD-win.spec
if errorlevel 1 (
    echo [ОШИБКА] Сборка не удалась.
    pause
    exit /b 1
)

echo.
echo [3/3] Копирование в папку dist\Для_Windows...
if not exist "dist\Для_Windows" mkdir "dist\Для_Windows"
copy /Y "dist\SverkaOFD.exe" "dist\Для_Windows\SverkaOFD.exe"
copy /Y "ИНСТРУКЦИЯ_ДЛЯ_WINDOWS.txt" "dist\Для_Windows\ИНСТРУКЦИЯ_ДЛЯ_WINDOWS.txt"

echo.
echo ============================================
echo   ГОТОВО!
echo   Файл: dist\Для_Windows\SverkaOFD.exe
echo ============================================
echo.
pause

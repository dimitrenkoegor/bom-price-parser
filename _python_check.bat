@echo off
rem Общая проверка для всех bat-файлов: Python 3.11+ и единственная зависимость openpyxl.
rem Возвращает код 1, если Python не найден или слишком старый.

where python >nul 2>nul
if errorlevel 1 (
  echo [!] Python не найден на компьютере.
  echo.
  echo Установите Python 3.11 или новее:
  echo   1^) Откройте https://www.python.org/downloads/
  echo   2^) Запустите установщик
  echo   3^) ВАЖНО: поставьте галочку "Add python.exe to PATH", затем Install Now
  echo   4^) После установки снова запустите этот файл
  echo.
  pause
  exit /b 1
)

python -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 11) else 1)" >nul 2>nul
if errorlevel 1 (
  echo [!] Нужен Python 3.11 или новее. Установлен:
  python --version
  pause
  exit /b 1
)

echo [1/2] Проверяю зависимости (openpyxl)...
python -m pip install --quiet openpyxl
echo      готово.
echo.
exit /b 0

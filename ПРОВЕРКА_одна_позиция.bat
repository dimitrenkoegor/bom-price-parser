@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Проверка одной позиции

where python >nul 2>nul
if errorlevel 1 (
  echo [!] Python не найден. Сначала запустите ЗАПУСТИТЬ.bat и установите Python.
  pause
  exit /b
)

python -m pip install --quiet undetected-chromedriver selenium beautifulsoup4 openpyxl webdriver-manager

echo Проверяю одну позицию: C5750X7R1H106KT, количество 320
echo Откроется окно Chrome - дайте ему пройти проверку Cloudflare, не закрывайте.
echo.
python price_parser.py --once C5750X7R1H106KT 320

echo.
echo Если выше видно строку с дистрибьютором и ценой - всё работает,
echo можно запускать ЗАПУСТИТЬ.bat для всего списка.
pause

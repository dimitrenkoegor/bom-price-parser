@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Распознать позиции (без поиска цен)

where python >nul 2>nul
if errorlevel 1 (
  echo [!] Python не найден. Сначала запустите ЗАПУСТИТЬ.bat и установите Python.
  pause
  exit /b
)
python -m pip install --quiet openpyxl

echo Распознаю позиции из Excel-запроса в "превью_позиции.xlsx" (цены не ищу)...
echo.
python price_parser.py --preview-only

echo.
echo Откройте "превью_позиции.xlsx", проверьте артикулы.
echo Жёлтым подсвечены строки, где артикул не распознан - впишите вручную.
pause

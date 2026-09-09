@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Парсер цен ЭКБ — файл из start\

call "%~dp0_python_check.bat" || exit /b 1

set "INPUT="
for %%f in ("start\*.xlsx") do (
  if not defined INPUT set "INPUT=%%f"
)
if not defined INPUT (
  echo [!] В папке start\ нет ни одного .xlsx — положите туда запрос.
  pause
  exit /b 1
)

echo Входной файл: %INPUT%
echo Запуск парсера (превью, затем поиск цен по API)...
echo.
python price_parser.py --input "%INPUT%"

echo.
if errorlevel 1 (
  echo [!] Завершилось с ошибкой — см. текст выше.
) else (
  echo Готово. Результат: final\BOM_Приложение_1.xlsx
)
pause

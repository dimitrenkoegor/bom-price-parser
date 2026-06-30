@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Парсер цен ЭКБ

echo ================================================
echo   Парсер цен на ЭКБ (oemsecrets)
echo ================================================
echo.
echo Скрипт сам найдёт свежий Excel-запрос в этой папке
echo (или в папке start), распознает позиции, покажет превью
echo и спросит подтверждение перед поиском цен.
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo [!] Python не найден на компьютере.
  echo.
  echo Установите Python:
  echo   1) Откройте https://www.python.org/downloads/
  echo   2) Запустите установщик
  echo   3) ВАЖНО: поставьте галочку "Add python.exe to PATH", затем Install Now
  echo   4) После установки снова дважды кликните этот файл
  echo.
  pause
  exit /b
)

echo [1/2] Проверяю зависимости (один раз, может занять пару минут)...
python -m pip install --quiet --upgrade pip
python -m pip install --quiet undetected-chromedriver selenium beautifulsoup4 openpyxl webdriver-manager
echo      готово.
echo.

echo [2/2] Запуск. Сначала будет ПРЕВЬЮ распознанных позиций.
echo      - проверьте файл "превью_позиции.xlsx" (откроется как обычный Excel),
echo        при необходимости поправьте артикулы и СОХРАНИТЕ его,
echo      - затем вернитесь сюда и нажмите Enter.
echo      Потом откроется окно Chrome для поиска цен - не закрывайте его.
echo.
python price_parser.py

echo.
echo ================================================
echo   Результат: final\BOM_Приложение_1.xlsx
echo ================================================
pause

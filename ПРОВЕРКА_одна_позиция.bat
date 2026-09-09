@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Проверка одной позиции

call "%~dp0_python_check.bat" || exit /b 1

echo Проверяю одну позицию: C5750X7R1H106KT (TDK), количество 320.
echo Цены идут через API из .env — окна браузера не будет.
echo.
python price_parser.py --once C5750X7R1H106KT 320 --manufacturer TDK --yes --output final\_проверка.xlsx

echo.
echo Если выше видна строка с дистрибьютором и ценой - всё работает,
echo можно запускать ЗАПУСТИТЬ.bat для всего списка.
echo Если везде RFQ и «нет ключа» - заполните .env (см. .env.example).
pause

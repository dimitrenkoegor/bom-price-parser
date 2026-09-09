@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Проверка окружения

call "%~dp0_python_check.bat" || exit /b 1
python --version
echo.
echo Готовность площадок по ключам из .env:
python -X utf8 -c "import sys; sys.path.insert(0, 'src'); from bomprice import registry; [print(f'  {s.id:<12} {registry.readiness(s)[0]:<12} {registry.readiness(s)[1]}') for s in registry.load_sources()]"
echo.
pause

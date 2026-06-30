@echo off
title Price Parser

cd /d "%~dp0"

py --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install from https://www.python.org
    pause
    exit /b 1
)

echo Installing dependencies...
py -m pip install undetected-chromedriver selenium requests beautifulsoup4 openpyxl webdriver-manager --quiet

echo.
if exist "start\Zapros.xlsx" (
    set INPUT=start\Zapros.xlsx
) else if exist "start\=D0=97=D0=B0=D0=BF=D1=80=D0=BE=D1=81.xlsx" (
    set INPUT=start\=D0=97=D0=B0=D0=BF=D1=80=D0=BE=D1=81.xlsx
) else (
    for %%f in ("start\*.xlsx") do set INPUT=%%f
)

if not defined INPUT (
    echo ERROR: No input file found in start\ folder
    pause
    exit /b 1
)

echo Input file: %INPUT%
echo Running parser...
echo.

py "%~dp0price_parser.py" --input "%INPUT%"

echo.
if errorlevel 1 (
    echo FAILED - see error above
) else (
    echo DONE - check price_results.xlsx
)
pause

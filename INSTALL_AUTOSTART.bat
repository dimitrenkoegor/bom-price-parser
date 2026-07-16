@echo off
rem Устанавливает автозапуск монитора почты при входе в Windows.
rem Запустите этот файл один раз двойным кликом.
copy /Y "%~dp0autostart_monitor.bat" "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\bom_mail_monitor.bat"
if %errorlevel%==0 (
    echo.
    echo Автозапуск установлен: монитор почты будет стартовать при входе в Windows.
    echo Удалить автозапуск: удалите файл bom_mail_monitor.bat из папки
    echo %APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\
) else (
    echo Не удалось скопировать файл автозапуска.
)
pause

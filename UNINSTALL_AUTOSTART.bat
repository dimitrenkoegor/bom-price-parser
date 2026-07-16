@echo off
rem Отключает автозапуск монитора почты при входе в Windows.
if exist "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\bom_mail_monitor.bat" (
    del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\bom_mail_monitor.bat"
    echo Автозапуск ОТКЛЮЧЁН: монитор больше не будет стартовать при входе в Windows.
    echo Если монитор сейчас запущен - он продолжает работать.
    echo Остановить работающий монитор: STOP_MAIL.bat
) else (
    echo Автозапуск и так не установлен.
)
pause

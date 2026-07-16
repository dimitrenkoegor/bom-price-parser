@echo off
rem Автостарт монитора почты BOM Price Parser при входе в Windows
rem (этот файл копируется в папку автозагрузки установщиком INSTALL_AUTOSTART.bat)
powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Users\Huawei\Desktop\Project C\BOM\start_mail.ps1"

# Общий запуск для всех .bat проекта: они однострочные и чисто ASCII, весь текст — здесь.
# Причина: cmd.exe с chcp 65001 читает UTF-8 bat со сбоем позиции и исполняет строки
# с середины ('жет' is not recognized ...). Файл хранить в UTF-8 с BOM и CRLF —
# без BOM Windows PowerShell 5.1 читает его как ANSI.
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('run', 'start', 'check', 'once', 'preview', 'autostart-install', 'autostart-remove')]
    [string]$Mode
)

$root = $PSScriptRoot
Set-Location -LiteralPath $root

function Say([string]$text = '') { Write-Host $text }

function Test-Python {
    $version = $null
    if (Get-Command python -ErrorAction SilentlyContinue) {
        $version = & python -c "import sys; print('%d.%d' % sys.version_info[:2])"
        if ($LASTEXITCODE -ne 0) { $version = $null }
    }
    if (-not $version) {
        Say '[!] Python не найден на компьютере.'
        Say
        Say 'Установите Python 3.10 или новее:'
        Say '  1) Откройте https://www.python.org/downloads/'
        Say '  2) Запустите установщик'
        Say '  3) ВАЖНО: поставьте галочку "Add python.exe to PATH", затем Install Now'
        Say '  4) После установки снова запустите этот файл'
        return $false
    }
    $parts = "$version".Trim().Split('.')
    if ([int]$parts[0] -lt 3 -or ([int]$parts[0] -eq 3 -and [int]$parts[1] -lt 10)) {
        Say "[!] Нужен Python 3.10 или новее. Установлен: $version"
        return $false
    }
    Say '[1/2] Проверяю зависимости (openpyxl)...'
    # Out-Host: вывод pip не должен попасть в возвращаемое значение функции
    & python -m pip install --quiet openpyxl | Out-Host
    if ($LASTEXITCODE -ne 0) {
        Say '     [!] pip не смог поставить openpyxl — проверьте интернет и запустите снова.'
    } else {
        Say '     готово.'
    }
    Say
    return $true
}

# Автозапуск — значение в HKCU\...\Run: строка там в Unicode, поэтому путь к проекту
# с кириллицей не ломается (bat в Startup его не прочитает, ярлык WScript.Shell портит).
$runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$runName = 'bom_mail_monitor'
$legacyAutostart = Join-Path ([Environment]::GetFolderPath('Startup')) 'bom_mail_monitor.bat'

switch ($Mode) {
    'run' {
        $Host.UI.RawUI.WindowTitle = 'Парсер цен ЭКБ'
        Say '================================================'
        Say '  Парсер цен на ЭКБ (API дистрибьюторов)'
        Say '================================================'
        Say
        Say 'Скрипт сам найдёт свежий Excel-запрос в папке start (или рядом),'
        Say 'распознает позиции, покажет превью и спросит подтверждение'
        Say 'перед поиском цен. Браузер не нужен: цены идут через API'
        Say 'DigiKey, Mouser, TME и Newark/Farnell по ключам из файла .env.'
        Say
        if (-not (Test-Python)) { exit 1 }
        Say '[2/2] Запуск. Сначала будет ПРЕВЬЮ распознанных позиций.'
        Say '     - проверьте файл "превью_позиции.xlsx" (откроется как обычный Excel),'
        Say '       при необходимости поправьте артикулы и производителей, СОХРАНИТЕ его,'
        Say '     - затем вернитесь сюда и нажмите Enter.'
        Say
        & python price_parser.py
        $code = $LASTEXITCODE
        Say
        Say '================================================'
        Say '  Результат: final\BOM_Приложение_1.xlsx'
        Say '================================================'
        exit $code
    }
    'start' {
        $Host.UI.RawUI.WindowTitle = 'Парсер цен ЭКБ — файл из start\'
        if (-not (Test-Python)) { exit 1 }
        $request = Get-ChildItem -LiteralPath (Join-Path $root 'start') -Filter '*.xlsx' -File -ErrorAction SilentlyContinue |
            Where-Object { -not $_.Name.StartsWith('~$') } | Sort-Object Name | Select-Object -First 1
        if (-not $request) {
            Say '[!] В папке start\ нет ни одного .xlsx — положите туда запрос.'
            exit 1
        }
        Say "Входной файл: start\$($request.Name)"
        Say 'Запуск парсера (превью, затем поиск цен по API)...'
        Say
        & python price_parser.py --input $request.FullName
        $code = $LASTEXITCODE
        Say
        if ($code -ne 0) {
            Say '[!] Завершилось с ошибкой — см. текст выше.'
        } else {
            Say 'Готово. Результат: final\BOM_Приложение_1.xlsx'
        }
        exit $code
    }
    'check' {
        $Host.UI.RawUI.WindowTitle = 'Проверка окружения'
        if (-not (Test-Python)) { exit 1 }
        & python --version
        Say
        Say 'Готовность площадок по ключам из .env:'
        # price_parser читает .env при импорте; сети здесь нет
        & python -X utf8 -c "import price_parser as p; rows = [('DigiKey', ['DIGIKEY_CLIENT_ID', 'DIGIKEY_CLIENT_SECRET']), ('Mouser', ['MOUSER_API_KEY']), ('Newark/Farnell', ['FARNELL_API_KEY']), ('TME', ['TME_TOKEN', 'TME_APP_SECRET']), ('oemsecrets', ['OEMSECRETS_API_KEY'])]; [print(f'  {name:<15} ' + ('готов' if all(p.ENV.get(k) for k in keys) else 'нет ключа    ' + ', '.join(k for k in keys if not p.ENV.get(k)))) for name, keys in rows]"
        $code = $LASTEXITCODE
        Say
        exit $code
    }
    'once' {
        $Host.UI.RawUI.WindowTitle = 'Проверка одной позиции'
        if (-not (Test-Python)) { exit 1 }
        Say 'Проверяю одну позицию: C5750X7R1H106KT (TDK), количество 320.'
        Say 'Цены идут через API из .env — окна браузера не будет.'
        Say
        & python price_parser.py --once C5750X7R1H106KT 320 --manufacturer TDK --yes --output 'final\_проверка.xlsx'
        $code = $LASTEXITCODE
        Say
        Say 'Если выше видна строка с дистрибьютором и ценой - всё работает,'
        Say 'можно запускать ЗАПУСТИТЬ.bat для всего списка.'
        Say 'Если везде RFQ и «нет ключа» - заполните .env (см. .env.example).'
        exit $code
    }
    'preview' {
        $Host.UI.RawUI.WindowTitle = 'Распознать позиции (без поиска цен)'
        if (-not (Test-Python)) { exit 1 }
        Say 'Распознаю позиции из Excel-запроса в "превью_позиции.xlsx" (цены не ищу)...'
        Say
        & python price_parser.py --preview-only
        $code = $LASTEXITCODE
        Say
        Say 'Откройте "превью_позиции.xlsx", проверьте артикулы и производителей.'
        Say 'Красным подсвечен нераспознанный артикул, жёлтым - пустой производитель'
        Say '(такая позиция уйдёт в RFQ). Впишите вручную и сохраните файл.'
        exit $code
    }
    'autostart-install' {
        # Старый bom_mail_monitor.bat из Startup убираем, иначе стартуют два монитора на один ящик.
        if (Test-Path -LiteralPath $legacyAutostart) { Remove-Item -LiteralPath $legacyAutostart -Force }
        $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $command = '"' + $powershell + '" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + (Join-Path $root 'start_mail.ps1') + '"'
        if (-not (Test-Path -LiteralPath $runKey)) { New-Item -Path $runKey -Force | Out-Null }
        Set-ItemProperty -LiteralPath $runKey -Name $runName -Value $command
        $saved = Get-ItemProperty -LiteralPath $runKey -Name $runName -ErrorAction SilentlyContinue
        if ($saved -and $saved.$runName -eq $command) {
            Say
            Say 'Автозапуск установлен: монитор почты будет стартовать при входе в Windows.'
            Say "Путь к проекту: $root"
            Say 'После переноса папки проекта запустите INSTALL_AUTOSTART.bat заново.'
            Say 'Удалить автозапуск: UNINSTALL_AUTOSTART.bat'
            exit 0
        }
        Say "Не удалось записать автозапуск в $runKey"
        exit 1
    }
    'autostart-remove' {
        $removed = $false
        if (Get-ItemProperty -LiteralPath $runKey -Name $runName -ErrorAction SilentlyContinue) {
            Remove-ItemProperty -LiteralPath $runKey -Name $runName
            $removed = $true
        }
        if (Test-Path -LiteralPath $legacyAutostart) {
            Remove-Item -LiteralPath $legacyAutostart -Force
            $removed = $true
        }
        if ($removed) {
            Say 'Автозапуск ОТКЛЮЧЁН: монитор больше не будет стартовать при входе в Windows.'
            Say 'Если монитор сейчас запущен - он продолжает работать.'
            Say 'Остановить работающий монитор: STOP_MAIL.bat'
        } else {
            Say 'Автозапуск и так не установлен.'
        }
        exit 0
    }
}

$ErrorActionPreference = 'Continue'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$pidFile = Join-Path $root 'monitor.pid'

# Мягкая остановка: демон увидит флаг на следующем цикле
Set-Content -LiteralPath (Join-Path $root 'STOP_MONITOR') -Value 'stop' -NoNewline
Write-Host 'Флаг STOP_MONITOR создан — демон завершится на следующем цикле (до POLL_SECONDS).'

if (Test-Path -LiteralPath $pidFile) {
    $monitorId = [int](Get-Content -LiteralPath $pidFile)
    Start-Sleep -Seconds 5
    $process = Get-Process -Id $monitorId -ErrorAction SilentlyContinue
    if ($process) {
        Stop-Process -Id $monitorId -Force
        Write-Host "Процесс остановлен принудительно (PID $monitorId)."
    } else {
        Write-Host 'Процесс уже завершился.'
    }
    Remove-Item -LiteralPath $pidFile -Force -ErrorAction SilentlyContinue
}

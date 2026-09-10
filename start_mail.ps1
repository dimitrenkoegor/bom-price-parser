$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$pidFile = Join-Path $root 'monitor.pid'
$stopFlag = Join-Path $root 'STOP_MONITOR'

if (Test-Path -LiteralPath $stopFlag) { Remove-Item -LiteralPath $stopFlag -Force }

if (Test-Path -LiteralPath $pidFile) {
    $existing = Get-Process -Id ([int](Get-Content -LiteralPath $pidFile)) -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "Монитор уже запущен (PID $($existing.Id))."
        exit 0
    }
    Remove-Item -LiteralPath $pidFile -Force
}

$python = (Get-Command python -ErrorAction Stop).Source
$process = Start-Process -FilePath $python -ArgumentList @('mail_monitor.py') -WorkingDirectory $root -WindowStyle Hidden -RedirectStandardOutput (Join-Path $root 'monitor.out.log') -RedirectStandardError (Join-Path $root 'monitor.err.log') -PassThru
Set-Content -LiteralPath $pidFile -Value $process.Id -NoNewline
Write-Host "Монитор почты запущен (PID $($process.Id)). Логи: monitor.log"

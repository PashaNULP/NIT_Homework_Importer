param(
    [ValidateSet(
        'start',
        'stop',
        'restart',
        'status',
        'install-autostart',
        'remove-autostart',
        'update'
    )]
    [string]$Action = 'start'
)

$ErrorActionPreference = 'Stop'
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = Join-Path $ProjectDir '.venv\Scripts\python.exe'
$AppScript = Join-Path $ProjectDir 'nit_web.py'
$RuntimeDir = Join-Path $ProjectDir '.nit2_data\server'
$LogDir = Join-Path $ProjectDir 'logs'
$PidFile = Join-Path $RuntimeDir 'nit2.pid'
$HealthUrl = 'http://127.0.0.1:8080/health'
$RepoZipUrl = 'https://github.com/PashaNULP/NIT_Homework_Importer/archive/refs/heads/main.zip'
$StartupFile = Join-Path ([Environment]::GetFolderPath('Startup')) 'NIT2 Server.cmd'

function Ensure-Directories {
    New-Item -ItemType Directory -Force -Path $RuntimeDir | Out-Null
    New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
}

function Require-Installation {
    if (-not (Test-Path -LiteralPath $Python)) {
        throw 'Не знайдено .venv. Спочатку запустіть "Встановити NIT Windows.bat".'
    }
    if (-not (Test-Path -LiteralPath $AppScript)) {
        throw 'Не знайдено nit_web.py у папці NIT2.'
    }
}

function Get-NitProcesses {
    if (-not (Test-Path -LiteralPath $Python)) {
        return @()
    }
    $expectedPython = [IO.Path]::GetFullPath($Python)
    $result = @()
    try {
        $items = Get-CimInstance Win32_Process -ErrorAction Stop
        foreach ($item in $items) {
            if (
                $item.ExecutablePath -and
                ([IO.Path]::GetFullPath($item.ExecutablePath) -ieq $expectedPython) -and
                $item.CommandLine -and
                $item.CommandLine -match 'nit_web\.py'
            ) {
                $process = Get-Process -Id $item.ProcessId -ErrorAction SilentlyContinue
                if ($null -ne $process) {
                    $result += $process
                }
            }
        }
    }
    catch {
        Write-Warning 'Windows не дозволив перевірити командний рядок процесів NIT2.'
    }
    return @($result | Sort-Object Id -Unique)
}

function Test-NitHealth {
    try {
        $response = Invoke-RestMethod -Uri $HealthUrl -TimeoutSec 3 -ErrorAction Stop
        return $response.status -eq 'ok' -and $response.app -eq 'NIT2'
    }
    catch {
        return $false
    }
}

function Remove-OldLogs {
    if (-not (Test-Path -LiteralPath $LogDir)) {
        return
    }
    Get-ChildItem -LiteralPath $LogDir -File -Filter 'nit2_*.log' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -Skip 30 |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

function Start-Nit {
    Require-Installation
    Ensure-Directories
    if (Test-NitHealth) {
        Write-Host 'NIT2 уже працює.' -ForegroundColor Yellow
        Write-Host 'Адреса: http://localhost:8080'
        return
    }
    $running = @(Get-NitProcesses)
    if ($running.Count -gt 0) {
        Write-Host "NIT2 уже працює. PID: $($running[0].Id)" -ForegroundColor Yellow
        Write-Host 'Адреса: http://localhost:8080'
        return
    }

    Remove-OldLogs
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $stdout = Join-Path $LogDir "nit2_${stamp}.out.log"
    $stderr = Join-Path $LogDir "nit2_${stamp}.err.log"

    $env:NIT_HOST = '0.0.0.0'
    $env:NIT_PORT = '8080'
    $env:NIT_SHOW_BROWSER = '0'
    $process = Start-Process `
        -FilePath $Python `
        -ArgumentList @('-u', 'nit_web.py') `
        -WorkingDirectory $ProjectDir `
        -WindowStyle Hidden `
        -RedirectStandardOutput $stdout `
        -RedirectStandardError $stderr `
        -PassThru
    Set-Content -LiteralPath $PidFile -Value $process.Id -Encoding ASCII

    $ready = $false
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        Start-Sleep -Milliseconds 500
        if ($process.HasExited) {
            break
        }
        if (Test-NitHealth) {
            $ready = $true
            break
        }
    }
    if (-not $ready) {
        if (-not $process.HasExited) {
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        }
        Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
        Write-Host "Лог помилки: $stderr" -ForegroundColor Yellow
        if (Test-Path -LiteralPath $stderr) {
            Get-Content -LiteralPath $stderr -Tail 30
        }
        throw 'NIT2 не відповів протягом 10 секунд.'
    }

    Write-Host "NIT2 запущено у фоні. PID: $($process.Id)" -ForegroundColor Green
    Write-Host 'На цьому комп''ютері: http://localhost:8080'
    $addresses = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*' } |
        Select-Object -ExpandProperty IPAddress -Unique
    foreach ($address in $addresses) {
        Write-Host "У локальній мережі: http://${address}:8080"
    }
    Write-Host "Логи: $LogDir"
}

function Stop-Nit {
    Ensure-Directories
    $running = @(Get-NitProcesses)
    if ($running.Count -eq 0) {
        if (Test-NitHealth) {
            throw 'Сайт NIT2 відповідає, але його процес не вдалося безпечно ідентифікувати. Перезапустіть Windows або завершіть nit_web.py через Диспетчер завдань.'
        }
        Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
        Write-Host 'NIT2 уже зупинено.' -ForegroundColor Yellow
        return
    }
    foreach ($process in $running) {
        Stop-Process -Id $process.Id -Force -ErrorAction Stop
        Write-Host "Зупинено NIT2, PID: $($process.Id)" -ForegroundColor Green
    }
    Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
}

function Show-Status {
    $running = @(Get-NitProcesses)
    if (Test-NitHealth) {
        $pidText = if ($running.Count -gt 0) { " PID: $($running[0].Id)" } else { '' }
        Write-Host "NIT2 працює.$pidText" -ForegroundColor Green
        Write-Host 'Адреса: http://localhost:8080'
        exit 0
    }
    if ($running.Count -gt 0) {
        Write-Host 'Процес NIT2 запущений, але сайт не відповідає.' -ForegroundColor Yellow
        exit 2
    }
    Write-Host 'NIT2 не запущено.' -ForegroundColor Red
    exit 1
}

function Install-Autostart {
    Require-Installation
    $scriptPath = Join-Path $ProjectDir 'NIT2-Windows.ps1'
    $content = "@echo off`r`npowershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`" -Action start`r`n"
    [IO.File]::WriteAllText($StartupFile, $content, [Text.UTF8Encoding]::new($false))
    Write-Host 'Автозапуск NIT2 увімкнено для поточного користувача Windows.' -ForegroundColor Green
    Write-Host "Файл автозапуску: $StartupFile"
    Start-Nit
}

function Remove-Autostart {
    if (Test-Path -LiteralPath $StartupFile) {
        Remove-Item -LiteralPath $StartupFile -Force
    }
    Write-Host 'Автозапуск NIT2 вимкнено.' -ForegroundColor Green
}

function Update-Nit {
    Require-Installation
    Ensure-Directories
    $temporary = Join-Path ([IO.Path]::GetTempPath()) ("nit2_update_" + [guid]::NewGuid().ToString('N'))
    $archive = Join-Path $temporary 'main.zip'
    $expanded = Join-Path $temporary 'expanded'
    New-Item -ItemType Directory -Force -Path $expanded | Out-Null
    try {
        Write-Host 'Завантажую останню версію NIT2 з GitHub...'
        Invoke-WebRequest -Uri $RepoZipUrl -OutFile $archive -UseBasicParsing
        Expand-Archive -LiteralPath $archive -DestinationPath $expanded -Force
        $source = Get-ChildItem -LiteralPath $expanded -Directory | Select-Object -First 1
        if ($null -eq $source -or -not (Test-Path (Join-Path $source.FullName 'nit_web.py'))) {
            throw 'Завантажений архів не містить nit_web.py.'
        }

        Write-Host 'Оновлюю залежності...'
        & $Python -m pip install -r (Join-Path $source.FullName 'requirements_web.txt')
        if ($LASTEXITCODE -ne 0) {
            throw 'Не вдалося оновити Python-залежності.'
        }

        Write-Host 'Перевіряю завантажену версію...'
        Push-Location $source.FullName
        try {
            & $Python -m unittest discover -s tests -q
            if ($LASTEXITCODE -ne 0) {
                throw 'Автоматичні тести нової версії не пройшли.'
            }
        }
        finally {
            Pop-Location
        }

        Stop-Nit
        $backupDir = Join-Path $RuntimeDir 'update-backups'
        New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
        $backup = Join-Path $backupDir ("before_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.zip')
        $backupFiles = Get-ChildItem -LiteralPath $ProjectDir -File | Where-Object {
            $_.Extension -in @('.py', '.bat', '.ps1', '.txt', '.md', '.json') -and
            $_.Name -ne 'nit_config.local.json'
        }
        if ($backupFiles.Count -gt 0) {
            Compress-Archive -Path $backupFiles.FullName -DestinationPath $backup -Force
        }

        try {
            Write-Host 'Встановлюю перевірену версію...'
            & robocopy $source.FullName $ProjectDir /E /R:2 /W:1 /XD .git .venv .nit2_data .nicegui backups logs /XF nit_config.local.json | Out-Null
            if ($LASTEXITCODE -ge 8) {
                throw "Robocopy завершився з кодом $LASTEXITCODE."
            }
        }
        catch {
            if (Test-Path -LiteralPath $backup) {
                Write-Host 'Відновлюю попередні файли після помилки...' -ForegroundColor Yellow
                Expand-Archive -LiteralPath $backup -DestinationPath $ProjectDir -Force
            }
            Start-Nit
            throw "Оновлення скасовано, попередні файли відновлено. Причина: $($_.Exception.Message)"
        }
        Write-Host 'NIT2 успішно оновлено.' -ForegroundColor Green
        Write-Host "Backup попередніх файлів: $backup"
        Start-Nit
    }
    finally {
        Remove-Item -LiteralPath $temporary -Recurse -Force -ErrorAction SilentlyContinue
    }
}

try {
    switch ($Action) {
        'start' { Start-Nit }
        'stop' { Stop-Nit }
        'restart' { Stop-Nit; Start-Nit }
        'status' { Show-Status }
        'install-autostart' { Install-Autostart }
        'remove-autostart' { Remove-Autostart }
        'update' { Update-Nit }
    }
}
catch {
    Write-Host "Помилка: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

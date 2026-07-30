# .SYNOPSIS
#    os4k2excel-server - Webserver fuer OpenScape 4000 Port & Lizenz Export
#
# .DESCRIPTION
#    Stellt eine Weboberflaeche bereit, ueber die das ETL-Script os4k2excel.ps1
#    per Button gestartet, der Fortschritt verfolgt und die fertige Excel-Datei
#    heruntergeladen oder per Email versendet werden kann.
#    Integrierter Zeitplaner fuer automatische Ausfuehrung.
#
#    Technologie: Reines PowerShell mit System.Net.HttpListener (Port 8080).
#
# .PARAMETER ApiPath
#    Pfad zur api2hipath.exe.
#
# .PARAMETER ApiHost
#    IP-Adresse des OpenScape 4000 (Pflicht).
#
# .PARAMETER ApiUser
#    API-Benutzername (Pflicht).
#
# .PARAMETER ApiPassword
#    API-Passwort (Pflicht).
#
# .PARAMETER OutputPath
#    Zielverzeichnis fuer Ausgabedateien.
#
# .PARAMETER Port
#    HTTP-Port (Standard: 8080).
#
# .PARAMETER WebPassword
#    Passwort fuer die Web-Anmeldung (Pflicht).
#
# .PARAMETER SmtpServer
#    SMTP-Server fuer Email-Versand (optional).
#
# .PARAMETER SmtpPort
#    SMTP-Port (Standard: 25).
#
# .PARAMETER SmtpFrom
#    Absender-Adresse fuer Emails (optional).
#
# .PARAMETER SmtpTo
#    Standard-Empfaenger fuer Emails (optional, wird im Dashboard vorausgefuellt).

param (
    [string]$ApiPath     = 'C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe',
    [string]$ApiHost     = '',
    [string]$ApiUser     = '',
    [string]$ApiPassword = '',
    [string]$OutputPath  = $PSScriptRoot,
    [int]$Port           = 8080,
    [string]$WebPassword = '',
    [string]$SmtpServer  = '',
    [int]$SmtpPort       = 25,
    [string]$SmtpFrom    = '',
    [string]$SmtpTo      = ''
)

# --- Parameterprufung ---
if ([string]::IsNullOrWhiteSpace($ApiHost) -or [string]::IsNullOrWhiteSpace($ApiUser) -or [string]::IsNullOrWhiteSpace($ApiPassword) -or [string]::IsNullOrWhiteSpace($WebPassword)) {
    Write-Host ""
    Write-Host "FEHLER: -ApiHost, -ApiUser, -ApiPassword und -WebPassword muessen angegeben werden." -ForegroundColor Red
    Write-Host ""
    Write-Host "Verwendung:" -ForegroundColor Yellow
    Write-Host '  .\os4k2excel-server.ps1 -ApiHost <IP> -ApiUser <USER> -ApiPassword <PW> -WebPassword <WEBPW>' -ForegroundColor White
    Write-Host ""
    Write-Host "Pflichtparameter:" -ForegroundColor Yellow
    Write-Host "  -ApiHost        IP-Adresse des OpenScape 4000" -ForegroundColor Gray
    Write-Host "  -ApiUser        API-Benutzername" -ForegroundColor Gray
    Write-Host "  -ApiPassword    API-Passwort" -ForegroundColor Gray
    Write-Host "  -WebPassword    Passwort fuer die Weboberflaeche" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Optionale Parameter:" -ForegroundColor Yellow
    Write-Host "  -Port           HTTP-Port (Standard: 8080)" -ForegroundColor Gray
    Write-Host "  -OutputPath     Zielverzeichnis (Standard: Skriptverzeichnis)" -ForegroundColor Gray
    Write-Host "  -SmtpServer     SMTP-Server fuer Email-Versand" -ForegroundColor Gray
    Write-Host "  -SmtpPort       SMTP-Port (Standard: 25)" -ForegroundColor Gray
    Write-Host "  -SmtpFrom       Absender-Adresse" -ForegroundColor Gray
    Write-Host "  -SmtpTo         Standard-Empfaenger (wird im Dashboard vorausgefuellt)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Beispiel:" -ForegroundColor Yellow
    Write-Host '  .\os4k2excel-server.ps1 -ApiHost 192.0.2.10 -ApiUser xieapi -ApiPassword geheim -WebPassword team2024 -SmtpServer mail.firma.de -SmtpFrom os4k@firma.de -SmtpTo admin@firma.de' -ForegroundColor White
    Write-Host ""
    exit 1
}

# --- Ausgabeverzeichnis pruefen ---
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# --- Session-Verwaltung ---
$script:Sessions = @{}
$script:SessionTimeout = 60  # Minuten

# --- Job-Status ---
$script:CurrentJob = $null
$script:JobStartTime = $null
$script:JobOutput = @()
$script:LastJobResult = $null
$script:RunHistory = [System.Collections.ArrayList]::new()

# --- Shutdown ---
$script:ShutdownRequested = $false
$script:StopFilePath = Join-Path $PSScriptRoot "os4k2excel-server.stop"

# --- Zeitplaner ---
$script:Schedule = @{
    Enabled    = $false
    Mode       = 'weekly'   # 'weekly' oder 'monthly'
    Time       = '02:00'
    Days       = @('Mo', 'Di', 'Mi', 'Do', 'Fr')
    DayOfMonth = 1           # 1-28 fuer monatliche Ausfuehrung
    LastRun    = $null
    EmailTo    = ''
}

# --- Pfad zum ETL-Script ---
$script:EtlScriptPath = Join-Path $PSScriptRoot "os4k2excel.ps1"

# --- Wochentag-Mapping (PowerShell DayOfWeek -> Kuerzel) ---
$script:DayMap = @{
    'Monday'    = 'Mo'
    'Tuesday'   = 'Di'
    'Wednesday' = 'Mi'
    'Thursday'  = 'Do'
    'Friday'    = 'Fr'
    'Saturday'  = 'Sa'
    'Sunday'    = 'So'
}

function Test-Session {
    param([string]$Token)
    if ([string]::IsNullOrWhiteSpace($Token)) { return $false }
    if (-not $script:Sessions.ContainsKey($Token)) { return $false }
    $session = $script:Sessions[$Token]
    if ((Get-Date) -gt $session.Expires) {
        $script:Sessions.Remove($Token)
        return $false
    }
    $script:Sessions[$Token].Expires = (Get-Date).AddMinutes($script:SessionTimeout)
    return $true
}

function Get-SessionToken {
    param($Request)
    $cookieHeader = $Request.Headers["Cookie"]
    if ($cookieHeader) {
        foreach ($part in $cookieHeader.Split(';')) {
            $kv = $part.Trim().Split('=', 2)
            if ($kv[0] -eq 'session' -and $kv.Length -eq 2) {
                return $kv[1]
            }
        }
    }
    return $null
}

function Clean-ExpiredSessions {
    $now = Get-Date
    $expired = @($script:Sessions.Keys | Where-Object { $script:Sessions[$_].Expires -lt $now })
    foreach ($key in $expired) {
        $script:Sessions.Remove($key)
    }
}

function Send-Response {
    param(
        $Response,
        [string]$Body,
        [string]$ContentType = 'text/html; charset=utf-8',
        [int]$StatusCode = 200
    )
    $Response.StatusCode = $StatusCode
    $Response.ContentType = $ContentType
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $Response.ContentLength64 = $buffer.Length
    $Response.OutputStream.Write($buffer, 0, $buffer.Length)
    $Response.OutputStream.Close()
}

function Send-Redirect {
    param($Response, [string]$Location)
    $Response.StatusCode = 302
    $Response.RedirectLocation = $Location
    $Response.OutputStream.Close()
}

function Send-Json {
    param($Response, $Data)
    $json = $Data | ConvertTo-Json -Depth 5
    Send-Response -Response $Response -Body $json -ContentType 'application/json; charset=utf-8'
}

function Find-LatestExcel {
    $files = Get-ChildItem -Path $OutputPath -Filter "OS4K-PORT-*.xlsx" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
    if ($files.Count -gt 0) { return $files[0] }
    return $null
}

function Start-ExportJob {
    if ($script:CurrentJob -and $script:CurrentJob.State -eq 'Running') { return $false }
    if ($script:CurrentJob) {
        Remove-Job -Job $script:CurrentJob -Force -ErrorAction SilentlyContinue
    }
    $script:JobOutput = @()
    $script:JobStartTime = Get-Date

    $jobApiPath = $ApiPath
    $jobApiHost = $ApiHost
    $jobApiUser = $ApiUser
    $jobApiPassword = $ApiPassword
    $jobOutputPath = $OutputPath
    $jobScriptPath = $script:EtlScriptPath

    $script:CurrentJob = Start-Job -ScriptBlock {
        param($ScriptPath, $ApiPath, $ApiHost, $ApiUser, $ApiPassword, $OutputPath)
        & $ScriptPath -ApiPath $ApiPath -ApiHost $ApiHost -ApiUser $ApiUser -ApiPassword $ApiPassword -OutputPath $OutputPath -MarkDuplicate
    } -ArgumentList $jobScriptPath, $jobApiPath, $jobApiHost, $jobApiUser, $jobApiPassword, $jobOutputPath

    return $true
}

# --- Zeitplaner-Pruefung ---
function Invoke-ScheduleCheck {
    if (-not $script:Schedule.Enabled) { return }
    if ($script:CurrentJob -and $script:CurrentJob.State -eq 'Running') { return }

    $now = Get-Date

    # Geplante Uhrzeit parsen
    $timeParts = $script:Schedule.Time.Split(':')
    if ($timeParts.Length -ne 2) { return }
    $schedHour = [int]$timeParts[0]
    $schedMin  = [int]$timeParts[1]

    # Pruefen ob wir in der richtigen Minute sind
    if ($now.Hour -ne $schedHour -or $now.Minute -ne $schedMin) { return }

    # Pruefen ob heute schon gelaufen
    if ($script:Schedule.LastRun -and $script:Schedule.LastRun.Date -eq $now.Date) { return }

    # Modus pruefen
    $triggerLabel = ''
    if ($script:Schedule.Mode -eq 'monthly') {
        if ($now.Day -ne $script:Schedule.DayOfMonth) { return }
        $triggerLabel = "$($now.Day). des Monats $($script:Schedule.Time)"
    } else {
        $todayShort = $script:DayMap[$now.DayOfWeek.ToString()]
        if ($todayShort -notin $script:Schedule.Days) { return }
        $triggerLabel = "$todayShort $($script:Schedule.Time)"
    }

    # Zeitplaner ausloesen
    $script:Schedule.LastRun = $now
    $script:JobStartedBySchedule = $true
    $started = Start-ExportJob
    if ($started) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] ZEITPLANER: Export automatisch gestartet ($triggerLabel)" -ForegroundColor Magenta
    }
}

# --- Pruefung auf Stop-Datei ---
function Test-StopFile {
    if (Test-Path $script:StopFilePath) {
        Remove-Item $script:StopFilePath -Force -ErrorAction SilentlyContinue
        return $true
    }
    return $false
}

# --- HTML-Seiten ---

# Optionales Logo (base64-PNG) - Platzhalter: 1x1 transparent
$script:LogoBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

function Get-LoginPage {
    param([string]$Error = '')
    $errorHtml = if ($Error) { "<div class='alert alert-error'>$Error</div>" } else { '' }
    return @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>os4k2excel - Login</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #F8F8F8; color: #3C3950; min-height: 100vh; display: flex; align-items: center; justify-content: center; }
        .login-container { background: #FFFFFF; border-radius: 8px; padding: 40px; width: 400px; max-width: 90vw; box-shadow: 0 4px 16px rgba(0,0,0,0.08); border-top: 3px solid #EC870E; }
        .login-logo { text-align: center; margin-bottom: 20px; }
        .login-logo img { height: 40px; }
        .login-container h1 { font-size: 24px; margin-bottom: 8px; color: #3C3950; }
        .login-container p.subtitle { color: #99A9B5; margin-bottom: 24px; font-size: 14px; }
        .form-group { margin-bottom: 16px; }
        .form-group label { display: block; margin-bottom: 6px; font-size: 14px; color: #666B6E; }
        .form-group input { width: 100%; padding: 10px 14px; border: 1px solid #ddd; border-radius: 5px; background: #FFFFFF; color: #3C3950; font-size: 16px; outline: none; }
        .form-group input:focus { border-color: #EC870E; }
        .btn { width: 100%; padding: 12px; border: none; border-radius: 5px; font-size: 16px; font-weight: 600; cursor: pointer; transition: background 0.2s; }
        .btn-primary { background: #EC870E; color: white; }
        .btn-primary:hover { background: #D07A0D; }
        .alert { padding: 10px 14px; border-radius: 5px; margin-bottom: 16px; font-size: 14px; }
        .alert-error { background: #FDF2F2; border: 1px solid #E53E3E; color: #C53030; }
        .footer { text-align: center; margin-top: 20px; font-size: 12px; color: #99A9B5; }
    </style>
</head>
<body>
    <div class="login-container">
        <div class="login-logo"><img src="data:image/png;base64,$($script:LogoBase64)" alt="Logo"></div>
        <h1>os4k2excel</h1>
        <p class="subtitle">OpenScape 4000 Port & Lizenz Export</p>
        $errorHtml
        <form method="POST" action="/login">
            <div class="form-group">
                <label for="password">Passwort</label>
                <input type="password" id="password" name="password" placeholder="Web-Passwort eingeben" autofocus required>
            </div>
            <button type="submit" class="btn btn-primary">Anmelden</button>
        </form>
        <div class="footer">Powered by PowerShell HttpListener</div>
    </div>
</body>
</html>
"@
}

function Get-DashboardPage {
    $smtpConfigured = -not [string]::IsNullOrWhiteSpace($SmtpServer)
    $smtpSection = if ($smtpConfigured) { 'true' } else { 'false' }
    $defaultEmailTo = if ($SmtpTo) { $SmtpTo } else { '' }

    # Zeitplaner-Status
    $scheduleEnabled = if ($script:Schedule.Enabled) { 'true' } else { 'false' }
    $scheduleMode = $script:Schedule.Mode
    $scheduleTime = $script:Schedule.Time
    $scheduleDaysJson = ($script:Schedule.Days | ForEach-Object { "`"$_`"" }) -join ','
    $scheduleDayOfMonth = $script:Schedule.DayOfMonth
    $scheduleEmailTo = if ($script:Schedule.EmailTo) { $script:Schedule.EmailTo } else { '' }
    $scheduleLastRun = if ($script:Schedule.LastRun) { $script:Schedule.LastRun.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }

    # Letzte Laeufe als JSON
    $historyJson = '[]'
    if ($script:RunHistory.Count -gt 0) {
        $historyJson = ($script:RunHistory | Select-Object -Last 10 | ConvertTo-Json -Depth 3)
        if ($script:RunHistory.Count -eq 1) { $historyJson = "[$historyJson]" }
    }

    return @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>os4k2excel - Dashboard</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #F8F8F8; color: #3C3950; min-height: 100vh; }
        .navbar { background: #FFFFFF; padding: 12px 24px; display: flex; align-items: center; justify-content: space-between; border-bottom: 2px solid #EC870E; box-shadow: 0 1px 4px rgba(0,0,0,0.06); }
        .navbar .nav-left { display: flex; align-items: center; gap: 12px; }
        .navbar .nav-left img { height: 32px; }
        .navbar h1 { font-size: 20px; color: #3C3950; }
        .navbar .nav-right { display: flex; align-items: center; gap: 16px; }
        .navbar a, .navbar button.nav-btn { color: #666B6E; text-decoration: none; font-size: 14px; background: none; border: none; cursor: pointer; }
        .navbar a:hover, .navbar button.nav-btn:hover { color: #EC870E; }
        .container { max-width: 900px; margin: 24px auto; padding: 0 16px; }
        .card { background: #FFFFFF; border-radius: 8px; padding: 24px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-top: 3px solid #EC870E; }
        .card h2 { font-size: 18px; margin-bottom: 16px; color: #3C3950; border-bottom: 1px solid #eee; padding-bottom: 8px; }
        .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
        .info-item { background: #F0F2F5; padding: 12px; border-radius: 5px; }
        .info-item .label { font-size: 12px; color: #99A9B5; text-transform: uppercase; }
        .info-item .value { font-size: 16px; font-weight: 600; margin-top: 4px; }
        .btn { padding: 10px 20px; border: none; border-radius: 5px; font-size: 14px; font-weight: 600; cursor: pointer; transition: all 0.2s; display: inline-flex; align-items: center; gap: 8px; }
        .btn:disabled { opacity: 0.5; cursor: not-allowed; }
        .btn-primary { background: #EC870E; color: white; }
        .btn-primary:hover:not(:disabled) { background: #D07A0D; }
        .btn-success { background: #38A169; color: white; }
        .btn-success:hover:not(:disabled) { background: #2F855A; }
        .btn-secondary { background: #E2E8F0; color: #3C3950; }
        .btn-secondary:hover:not(:disabled) { background: #CBD5E0; }
        .btn-danger { background: #FED7D7; color: #C53030; }
        .btn-danger:hover:not(:disabled) { background: #FEB2B2; }
        .btn-small { padding: 6px 12px; font-size: 12px; }
        .progress-container { margin-top: 16px; display: none; }
        .progress-bar-wrapper { background: #E2E8F0; border-radius: 8px; height: 24px; overflow: hidden; margin-bottom: 8px; }
        .progress-bar { height: 100%; background: linear-gradient(90deg, #EC870E, #D07A0D); transition: width 0.5s ease; width: 0%; border-radius: 8px; }
        .progress-text { font-size: 14px; color: #666B6E; }
        .elapsed-text { font-size: 13px; color: #99A9B5; margin-top: 4px; }
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
        .status-idle { background: #E2E8F0; color: #718096; }
        .status-running { background: #EBF8FF; color: #2B6CB0; }
        .status-completed { background: #F0FFF4; color: #276749; }
        .status-error { background: #FDF2F2; color: #C53030; }
        .download-section { margin-top: 16px; display: none; }
        .email-section { margin-top: 16px; }
        .input-field { padding: 8px 12px; border: 1px solid #ddd; border-radius: 5px; background: #FFFFFF; color: #3C3950; font-size: 14px; outline: none; }
        .input-field:focus { border-color: #EC870E; }
        .email-section input[type="email"] { width: 300px; max-width: 100%; }
        .email-row { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
        .history-table { width: 100%; border-collapse: collapse; margin-top: 8px; }
        .history-table th, .history-table td { padding: 8px 12px; text-align: left; border-bottom: 1px solid #eee; font-size: 13px; }
        .history-table th { color: #99A9B5; font-weight: 600; }
        .alert { padding: 10px 14px; border-radius: 5px; margin-top: 12px; font-size: 14px; display: none; }
        .alert-success { background: #F0FFF4; border: 1px solid #38A169; color: #276749; }
        .alert-error { background: #FDF2F2; border: 1px solid #E53E3E; color: #C53030; }
        .spinner { display: inline-block; width: 16px; height: 16px; border: 2px solid #EC870E; border-top-color: transparent; border-radius: 50%; animation: spin 0.8s linear infinite; }
        @keyframes spin { to { transform: rotate(360deg); } }
        /* Zeitplaner */
        .schedule-row { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin-bottom: 12px; }
        .schedule-row label { font-size: 14px; color: #666B6E; min-width: 70px; }
        .schedule-row input[type="time"] { padding: 6px 10px; border: 1px solid #ddd; border-radius: 5px; background: #FFFFFF; color: #3C3950; font-size: 14px; outline: none; }
        .schedule-row input[type="time"]:focus { border-color: #EC870E; }
        .day-checks { display: flex; gap: 6px; flex-wrap: wrap; }
        .day-check { position: relative; }
        .day-check input { display: none; }
        .day-check label { display: block; padding: 6px 10px; border-radius: 5px; background: #F0F2F5; color: #666B6E; font-size: 13px; cursor: pointer; border: 1px solid #ddd; transition: all 0.2s; user-select: none; }
        .day-check input:checked + label { background: #EC870E; color: white; border-color: #EC870E; }
        .toggle-container { display: flex; align-items: center; gap: 10px; margin-bottom: 16px; }
        .toggle { position: relative; width: 44px; height: 24px; }
        .toggle input { display: none; }
        .toggle-slider { position: absolute; inset: 0; background: #CBD5E0; border-radius: 12px; cursor: pointer; transition: background 0.2s; }
        .toggle-slider::before { content: ''; position: absolute; left: 3px; top: 3px; width: 18px; height: 18px; background: #FFFFFF; border-radius: 50%; transition: all 0.2s; box-shadow: 0 1px 3px rgba(0,0,0,0.15); }
        .toggle input:checked + .toggle-slider { background: #EC870E; }
        .toggle input:checked + .toggle-slider::before { transform: translateX(20px); background: white; }
        .schedule-info { font-size: 13px; color: #99A9B5; margin-top: 8px; }
        .mode-tabs { display: flex; gap: 4px; margin-bottom: 16px; }
        .mode-tab { padding: 6px 16px; border-radius: 5px; background: #F0F2F5; color: #666B6E; font-size: 13px; cursor: pointer; border: 1px solid #ddd; transition: all 0.2s; user-select: none; }
        .mode-tab.active { background: #EC870E; color: white; border-color: #EC870E; }
        .input-number { width: 60px; text-align: center; }
        @media (max-width: 600px) {
            .info-grid { grid-template-columns: 1fr; }
            .email-row { flex-direction: column; align-items: stretch; }
            .email-section input[type="email"] { width: 100%; }
            .schedule-row { flex-direction: column; align-items: flex-start; }
        }
    </style>
</head>
<body>
    <div class="navbar">
        <div class="nav-left">
            <img src="data:image/png;base64,$($script:LogoBase64)" alt="Logo">
            <h1>os4k2excel</h1>
        </div>
        <div class="nav-right">
            <span style="color:#99A9B5; font-size:13px;">OpenScape 4000 Export</span>
            <a href="/logout">Abmelden</a>
            <button class="nav-btn" onclick="shutdownServer()" title="Server beenden">Server beenden</button>
        </div>
    </div>

    <div class="container">
        <!-- Verbindungsinfo -->
        <div class="card">
            <h2>API-Verbindung</h2>
            <div class="info-grid">
                <div class="info-item">
                    <div class="label">Host</div>
                    <div class="value">$ApiHost</div>
                </div>
                <div class="info-item">
                    <div class="label">Benutzer</div>
                    <div class="value">$ApiUser</div>
                </div>
                <div class="info-item">
                    <div class="label">Ausgabepfad</div>
                    <div class="value">$OutputPath</div>
                </div>
                <div class="info-item">
                    <div class="label">Status</div>
                    <div class="value" id="job-status-badge"><span class="status-badge status-idle">Bereit</span></div>
                </div>
            </div>
        </div>

        <!-- Ausfuehrung -->
        <div class="card">
            <h2>Export starten</h2>
            <button id="btn-run" class="btn btn-primary" onclick="startRun()">Jetzt starten</button>

            <div class="progress-container" id="progress-container">
                <div class="progress-bar-wrapper">
                    <div class="progress-bar" id="progress-bar"></div>
                </div>
                <div class="progress-text" id="progress-text">Wird gestartet...</div>
                <div class="elapsed-text" id="elapsed-text"></div>
            </div>

            <div class="download-section" id="download-section">
                <button class="btn btn-success" onclick="window.location.href='/download'">Excel herunterladen</button>
            </div>

            <div id="run-alert" class="alert"></div>
        </div>

        <!-- Zeitplaner -->
        <div class="card">
            <h2>Zeitplaner</h2>
            <div class="toggle-container">
                <label class="toggle" for="sched-enabled">
                    <input type="checkbox" id="sched-enabled" onchange="scheduleChanged()" $( if ($script:Schedule.Enabled) { 'checked' } )>
                    <span class="toggle-slider"></span>
                </label>
                <span id="sched-status-text" style="font-size:14px;">$( if ($script:Schedule.Enabled) { 'Aktiviert' } else { 'Deaktiviert' } )</span>
            </div>

            <div class="mode-tabs">
                <div class="mode-tab $( if ($scheduleMode -eq 'weekly') { 'active' } )" id="mode-weekly" onclick="setMode('weekly')">Woechentlich</div>
                <div class="mode-tab $( if ($scheduleMode -eq 'monthly') { 'active' } )" id="mode-monthly" onclick="setMode('monthly')">Monatlich</div>
            </div>

            <div class="schedule-row">
                <label>Uhrzeit</label>
                <input type="time" id="sched-time" value="$scheduleTime" class="input-field" onchange="scheduleChanged()">
            </div>

            <div class="schedule-row" id="row-days" style="display: $( if ($scheduleMode -eq 'weekly') { 'flex' } else { 'none' } );">
                <label>Wochentage</label>
                <div class="day-checks">
                    <div class="day-check"><input type="checkbox" id="day-Mo" value="Mo" onchange="scheduleChanged()" $( if ('Mo' -in $script:Schedule.Days) { 'checked' } )><label for="day-Mo">Mo</label></div>
                    <div class="day-check"><input type="checkbox" id="day-Di" value="Di" onchange="scheduleChanged()" $( if ('Di' -in $script:Schedule.Days) { 'checked' } )><label for="day-Di">Di</label></div>
                    <div class="day-check"><input type="checkbox" id="day-Mi" value="Mi" onchange="scheduleChanged()" $( if ('Mi' -in $script:Schedule.Days) { 'checked' } )><label for="day-Mi">Mi</label></div>
                    <div class="day-check"><input type="checkbox" id="day-Do" value="Do" onchange="scheduleChanged()" $( if ('Do' -in $script:Schedule.Days) { 'checked' } )><label for="day-Do">Do</label></div>
                    <div class="day-check"><input type="checkbox" id="day-Fr" value="Fr" onchange="scheduleChanged()" $( if ('Fr' -in $script:Schedule.Days) { 'checked' } )><label for="day-Fr">Fr</label></div>
                    <div class="day-check"><input type="checkbox" id="day-Sa" value="Sa" onchange="scheduleChanged()" $( if ('Sa' -in $script:Schedule.Days) { 'checked' } )><label for="day-Sa">Sa</label></div>
                    <div class="day-check"><input type="checkbox" id="day-So" value="So" onchange="scheduleChanged()" $( if ('So' -in $script:Schedule.Days) { 'checked' } )><label for="day-So">So</label></div>
                </div>
            </div>

            <div class="schedule-row" id="row-dayofmonth" style="display: $( if ($scheduleMode -eq 'monthly') { 'flex' } else { 'none' } );">
                <label>Tag</label>
                <input type="number" id="sched-dayofmonth" value="$scheduleDayOfMonth" min="1" max="28" class="input-field input-number" onchange="scheduleChanged()">
                <span style="font-size:13px; color:#99A9B5;">des Monats (1-28)</span>
            </div>

            <div class="schedule-row" style="display: $( if ($smtpConfigured) { 'flex' } else { 'none' } );">
                <label>Email an</label>
                <input type="email" id="sched-email" value="$scheduleEmailTo" placeholder="Nach Export automatisch senden" class="input-field" style="width:300px; max-width:100%;" onchange="scheduleChanged()">
            </div>

            <button class="btn btn-secondary btn-small" id="btn-save-schedule" onclick="saveSchedule()" style="display:none;">Speichern</button>
            <div id="schedule-alert" class="alert"></div>
            <div class="schedule-info" id="sched-last-run">$( if ($scheduleLastRun) { "Letzter automatischer Lauf: $scheduleLastRun" } else { '' } )</div>
        </div>

        <!-- Email -->
        <div class="card" id="email-card" style="display: none;">
            <h2>Email-Versand</h2>
            <div class="email-section">
                <div class="email-row">
                    <input type="email" id="email-to" placeholder="empfaenger@firma.de" value="$defaultEmailTo" class="input-field">
                    <button id="btn-email" class="btn btn-secondary" onclick="sendEmail()" disabled>Senden</button>
                </div>
            </div>
            <div id="email-alert" class="alert"></div>
        </div>

        <!-- History -->
        <div class="card">
            <h2>Letzte Ausfuehrungen</h2>
            <table class="history-table" id="history-table">
                <thead>
                    <tr><th>Datum</th><th>Dauer</th><th>Quelle</th><th>Status</th></tr>
                </thead>
                <tbody id="history-body">
                </tbody>
            </table>
        </div>
    </div>

    <script>
        const smtpConfigured = $smtpSection;
        let polling = false;
        let pollInterval = null;
        let scheduleModified = false;

        // Email-Karte anzeigen/verstecken
        if (smtpConfigured) {
            document.getElementById('email-card').style.display = 'block';
        }

        // History laden
        const initialHistory = $historyJson;
        renderHistory(initialHistory);

        // Sofort Status pruefen
        checkStatus();

        function startRun() {
            const btn = document.getElementById('btn-run');
            btn.disabled = true;
            hideAlert('run-alert');

            fetch('/run', { method: 'POST' })
                .then(r => r.json())
                .then(data => {
                    if (data.error) {
                        showAlert('run-alert', data.error, 'error');
                        btn.disabled = false;
                    } else {
                        startPolling();
                    }
                })
                .catch(err => {
                    showAlert('run-alert', 'Verbindungsfehler: ' + err, 'error');
                    btn.disabled = false;
                });
        }

        function startPolling() {
            polling = true;
            document.getElementById('progress-container').style.display = 'block';
            document.getElementById('download-section').style.display = 'none';
            if (smtpConfigured) document.getElementById('btn-email').disabled = true;
            if (pollInterval) clearInterval(pollInterval);
            pollInterval = setInterval(checkStatus, 2000);
        }

        function checkStatus() {
            fetch('/status')
                .then(r => r.json())
                .then(data => { updateUI(data); })
                .catch(() => {});
        }

        function updateUI(data) {
            const btn = document.getElementById('btn-run');
            const badge = document.getElementById('job-status-badge');
            const bar = document.getElementById('progress-bar');
            const text = document.getElementById('progress-text');
            const elapsed = document.getElementById('elapsed-text');
            const progressContainer = document.getElementById('progress-container');
            const downloadSection = document.getElementById('download-section');

            if (data.state === 'running') {
                badge.innerHTML = '<span class="status-badge status-running"><span class="spinner"></span> Laeuft</span>';
                btn.disabled = true;
                progressContainer.style.display = 'block';
                bar.style.width = data.progress_pct + '%';
                bar.style.background = 'linear-gradient(90deg, #EC870E, #D07A0D)';
                text.textContent = data.progress || 'Verarbeite...';
                elapsed.textContent = data.elapsed ? 'Laufzeit: ' + data.elapsed : '';
                downloadSection.style.display = 'none';
                if (!polling) startPolling();
            } else if (data.state === 'completed') {
                badge.innerHTML = '<span class="status-badge status-completed">Abgeschlossen</span>';
                btn.disabled = false;
                bar.style.width = '100%';
                text.textContent = 'Export abgeschlossen';
                elapsed.textContent = data.elapsed ? 'Gesamtdauer: ' + data.elapsed : '';
                downloadSection.style.display = 'block';
                if (smtpConfigured) document.getElementById('btn-email').disabled = false;
                stopPolling();
            } else if (data.state === 'error') {
                badge.innerHTML = '<span class="status-badge status-error">Fehler</span>';
                btn.disabled = false;
                bar.style.width = '100%';
                bar.style.background = 'linear-gradient(90deg, #E53E3E, #C53030)';
                text.textContent = data.progress || 'Fehler aufgetreten';
                elapsed.textContent = data.elapsed ? 'Laufzeit: ' + data.elapsed : '';
                stopPolling();
            } else {
                badge.innerHTML = '<span class="status-badge status-idle">Bereit</span>';
                btn.disabled = false;
                if (data.has_file) {
                    downloadSection.style.display = 'block';
                    if (smtpConfigured) document.getElementById('btn-email').disabled = false;
                }
                stopPolling();
            }

            if (data.history) { renderHistory(data.history); }

            // Zeitplaner LastRun aktualisieren
            if (data.schedule_last_run) {
                document.getElementById('sched-last-run').textContent = 'Letzter automatischer Lauf: ' + data.schedule_last_run;
            }
        }

        function stopPolling() {
            polling = false;
            if (pollInterval) { clearInterval(pollInterval); pollInterval = null; }
        }

        function sendEmail() {
            const emailTo = document.getElementById('email-to').value.trim();
            if (!emailTo) {
                showAlert('email-alert', 'Bitte eine Empfaenger-Adresse eingeben.', 'error');
                return;
            }
            hideAlert('email-alert');
            document.getElementById('btn-email').disabled = true;

            fetch('/email', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ to: emailTo })
            })
            .then(r => r.json())
            .then(data => {
                if (data.error) { showAlert('email-alert', data.error, 'error'); }
                else { showAlert('email-alert', data.message, 'success'); }
                document.getElementById('btn-email').disabled = false;
            })
            .catch(err => {
                showAlert('email-alert', 'Fehler: ' + err, 'error');
                document.getElementById('btn-email').disabled = false;
            });
        }

        // --- Zeitplaner ---
        let schedMode = '$scheduleMode';

        function scheduleChanged() {
            scheduleModified = true;
            document.getElementById('btn-save-schedule').style.display = 'inline-flex';
            const enabled = document.getElementById('sched-enabled').checked;
            document.getElementById('sched-status-text').textContent = enabled ? 'Aktiviert' : 'Deaktiviert';
        }

        function setMode(mode) {
            schedMode = mode;
            document.getElementById('mode-weekly').className = 'mode-tab' + (mode === 'weekly' ? ' active' : '');
            document.getElementById('mode-monthly').className = 'mode-tab' + (mode === 'monthly' ? ' active' : '');
            document.getElementById('row-days').style.display = mode === 'weekly' ? 'flex' : 'none';
            document.getElementById('row-dayofmonth').style.display = mode === 'monthly' ? 'flex' : 'none';
            scheduleChanged();
        }

        function saveSchedule() {
            const enabled = document.getElementById('sched-enabled').checked;
            const time = document.getElementById('sched-time').value;
            const days = [];
            ['Mo','Di','Mi','Do','Fr','Sa','So'].forEach(d => {
                if (document.getElementById('day-' + d).checked) days.push(d);
            });
            const dayOfMonth = parseInt(document.getElementById('sched-dayofmonth').value) || 1;
            const emailTo = document.getElementById('sched-email') ? document.getElementById('sched-email').value.trim() : '';

            fetch('/schedule', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ enabled: enabled, mode: schedMode, time: time, days: days, dayOfMonth: dayOfMonth, emailTo: emailTo })
            })
            .then(r => r.json())
            .then(data => {
                if (data.error) {
                    showAlert('schedule-alert', data.error, 'error');
                } else {
                    showAlert('schedule-alert', data.message, 'success');
                    document.getElementById('btn-save-schedule').style.display = 'none';
                    scheduleModified = false;
                }
            })
            .catch(err => { showAlert('schedule-alert', 'Fehler: ' + err, 'error'); });
        }

        // --- Server beenden ---
        function shutdownServer() {
            if (!confirm('Server wirklich beenden?')) return;
            fetch('/shutdown', { method: 'POST' })
                .then(r => r.json())
                .then(data => {
                    document.body.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;min-height:100vh;font-family:sans-serif;background:#F8F8F8;color:#3C3950;"><div style="text-align:center;"><h1 style="color:#EC870E;">Server beendet</h1><p style="color:#99A9B5;margin-top:12px;">Der Webserver wurde heruntergefahren.</p></div></div>';
                })
                .catch(() => {
                    document.body.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;min-height:100vh;font-family:sans-serif;background:#F8F8F8;color:#3C3950;"><div style="text-align:center;"><h1 style="color:#EC870E;">Server beendet</h1><p style="color:#99A9B5;margin-top:12px;">Verbindung getrennt.</p></div></div>';
                });
        }

        function renderHistory(history) {
            const tbody = document.getElementById('history-body');
            if (!history || history.length === 0) {
                tbody.innerHTML = '<tr><td colspan="4" style="color:#99A9B5;">Keine Ausfuehrungen</td></tr>';
                return;
            }
            tbody.innerHTML = history.map(h => {
                const statusClass = h.State === 'completed' ? 'status-completed' : (h.State === 'error' ? 'status-error' : 'status-running');
                const statusLabel = h.State === 'completed' ? 'Abgeschlossen' : (h.State === 'error' ? 'Fehler' : 'Laeuft');
                const source = h.Source || 'Manuell';
                return '<tr><td>' + h.StartTime + '</td><td>' + (h.Duration || '-') + '</td><td>' + source + '</td><td><span class="status-badge ' + statusClass + '">' + statusLabel + '</span></td></tr>';
            }).reverse().join('');
        }

        function showAlert(id, msg, type) {
            const el = document.getElementById(id);
            el.textContent = msg;
            el.className = 'alert alert-' + type;
            el.style.display = 'block';
            setTimeout(() => { el.style.display = 'none'; }, 5000);
        }

        function hideAlert(id) {
            document.getElementById(id).style.display = 'none';
        }
    </script>
</body>
</html>
"@
}

# --- Hilfsfunktion: Fortschritt aus Job-Output parsen ---
function Get-JobProgress {
    $output = @()
    if ($script:CurrentJob) {
        try {
            $output = @(Receive-Job -Job $script:CurrentJob -Keep -ErrorAction SilentlyContinue)
        } catch {}
    }
    $script:JobOutput = $output

    $lastLine = ''
    $progressPct = 0

    foreach ($line in $output) {
        $lineStr = "$line"
        if ($lineStr -match '\[(\d+)/(\d+)\]') {
            $current = [int]$Matches[1]
            $total = [int]$Matches[2]
            $progressPct = [math]::Min(90, 40 + [math]::Floor(($current / [math]::Max(1, $total)) * 50))
            $lastLine = $lineStr
        }
        elseif ($lineStr -match 'Starte API-Abfrage') {
            $progressPct = [math]::Min(40, $progressPct + 5)
            $lastLine = $lineStr
        }
        elseif ($lineStr -match 'Gesamtverarbeitung abgeschlossen') {
            $progressPct = 100
            $lastLine = $lineStr
        }
        elseif ($lineStr -match 'Erstelle.*Worksheet|Dashboard|Formatierung|Datenexport') {
            $progressPct = [math]::Min(98, [math]::Max($progressPct, 90))
            $lastLine = $lineStr
        }
        elseif (-not [string]::IsNullOrWhiteSpace($lineStr)) {
            $lastLine = $lineStr
        }
    }

    return @{
        Progress    = $lastLine
        ProgressPct = $progressPct
    }
}

function Parse-FormBody {
    param([string]$Body)
    $result = @{}
    foreach ($pair in $Body.Split('&')) {
        $kv = $pair.Split('=', 2)
        if ($kv.Length -eq 2) {
            $result[[System.Uri]::UnescapeDataString($kv[0])] = [System.Uri]::UnescapeDataString($kv[1].Replace('+', ' '))
        }
    }
    return $result
}

# --- Hilfsfunktion: History-Eintrag erstellen ---
function Add-HistoryEntry {
    param([string]$State, [string]$Duration, [string]$Source = 'Manuell')
    $startStr = $script:JobStartTime.ToString('yyyy-MM-dd HH:mm:ss')
    $alreadyLogged = $script:RunHistory | Where-Object { $_.StartTime -eq $startStr }
    if (-not $alreadyLogged) {
        $null = $script:RunHistory.Add([PSCustomObject]@{
            StartTime = $startStr
            Duration  = $Duration
            State     = $State
            Source    = $Source
        })
    }
}

# --- Hilfsfunktion: Job-Quelle ermitteln ---
function Get-JobSource {
    if ($script:JobStartedBySchedule) { return 'Zeitplaner' }
    return 'Manuell'
}

$script:JobStartedBySchedule = $false

# --- HTTP-Server starten ---
$listener = New-Object System.Net.HttpListener
$prefix = "http://+:$Port/"
$listener.Prefixes.Add($prefix)

try {
    $listener.Start()
} catch {
    Write-Host "FEHLER: Kann Port $Port nicht oeffnen." -ForegroundColor Red
    Write-Host "Moeglicherweise wird der Port bereits verwendet oder es fehlen Berechtigungen." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Fuer nicht-Admin-Benutzer einmalig ausfuehren (als Admin):" -ForegroundColor Yellow
    Write-Host "  netsh http add urlacl url=http://+:$Port/ user=$env:USERDOMAIN\$env:USERNAME" -ForegroundColor White
    Write-Host ""
    exit 1
}

Write-Host "" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  os4k2excel Webserver gestartet" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  URL:    http://localhost:$Port/" -ForegroundColor Cyan
Write-Host "  Host:   $ApiHost" -ForegroundColor Gray
Write-Host "  User:   $ApiUser" -ForegroundColor Gray
Write-Host "  Output: $OutputPath" -ForegroundColor Gray
if ($SmtpServer) {
    Write-Host "  SMTP:   $SmtpServer`:$SmtpPort" -ForegroundColor Gray
    Write-Host "  From:   $SmtpFrom" -ForegroundColor Gray
    if ($SmtpTo) {
        Write-Host "  To:     $SmtpTo" -ForegroundColor Gray
    }
}
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  Beenden: Ctrl+C im Terminal" -ForegroundColor Yellow
Write-Host "           oder: Dashboard > Server beenden" -ForegroundColor Yellow
Write-Host "           oder: echo.> `"$($script:StopFilePath)`"" -ForegroundColor Yellow
Write-Host ""

# --- Hauptschleife (async mit Timeout fuer Zeitplaner + Stop-Datei) ---
try {
    while ($listener.IsListening -and -not $script:ShutdownRequested) {

        # Async auf Request warten (Timeout 2 Sekunden fuer Hintergrund-Checks)
        $asyncResult = $listener.BeginGetContext($null, $null)

        while (-not $asyncResult.AsyncWaitHandle.WaitOne(2000)) {
            # Hintergrund-Checks waehrend wir auf Requests warten
            Invoke-ScheduleCheck
            if (Test-StopFile) {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Stop-Datei erkannt. Server wird beendet..." -ForegroundColor Yellow
                $script:ShutdownRequested = $true
                break
            }
        }

        if ($script:ShutdownRequested) { break }

        $context = $listener.EndGetContext($asyncResult)
        $request = $context.Request
        $response = $context.Response

        $method = $request.HttpMethod
        $path = $request.Url.AbsolutePath.TrimEnd('/')
        if ($path -eq '') { $path = '/' }

        $timestamp = Get-Date -Format "HH:mm:ss"
        $token = Get-SessionToken -Request $request
        $authenticated = Test-Session -Token $token

        Clean-ExpiredSessions

        try {
            # --- Routing ---
            switch ("$method $path") {

                'GET /' {
                    if ($authenticated) {
                        Send-Redirect -Response $response -Location '/dashboard'
                    } else {
                        $html = Get-LoginPage
                        Send-Response -Response $response -Body $html
                    }
                    Write-Host "[$timestamp] $method $path -> $(if ($authenticated) { '302 /dashboard' } else { '200 Login' })" -ForegroundColor Gray
                }

                'POST /login' {
                    $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
                    $body = $reader.ReadToEnd()
                    $reader.Close()
                    $formData = Parse-FormBody -Body $body

                    if ($formData['password'] -eq $WebPassword) {
                        $newToken = [guid]::NewGuid().ToString()
                        $script:Sessions[$newToken] = @{
                            Created = Get-Date
                            Expires = (Get-Date).AddMinutes($script:SessionTimeout)
                        }
                        $response.AppendHeader('Set-Cookie', "session=$newToken; Path=/; HttpOnly")
                        Send-Redirect -Response $response -Location '/dashboard'
                        Write-Host "[$timestamp] POST /login -> Login erfolgreich" -ForegroundColor Green
                    } else {
                        $html = Get-LoginPage -Error 'Falsches Passwort.'
                        Send-Response -Response $response -Body $html -StatusCode 401
                        Write-Host "[$timestamp] POST /login -> Falsches Passwort" -ForegroundColor Red
                    }
                }

                'GET /dashboard' {
                    if (-not $authenticated) {
                        Send-Redirect -Response $response -Location '/'
                        Write-Host "[$timestamp] GET /dashboard -> 302 / (nicht angemeldet)" -ForegroundColor Yellow
                    } else {
                        $html = Get-DashboardPage
                        Send-Response -Response $response -Body $html
                        Write-Host "[$timestamp] GET /dashboard -> 200" -ForegroundColor Gray
                    }
                }

                'POST /run' {
                    if (-not $authenticated) {
                        Send-Json -Response $response -Data @{ error = 'Nicht angemeldet' }
                        Write-Host "[$timestamp] POST /run -> 401" -ForegroundColor Yellow
                    } else {
                        if ($script:CurrentJob -and $script:CurrentJob.State -eq 'Running') {
                            Send-Json -Response $response -Data @{ error = 'Es laeuft bereits ein Export. Bitte warten.' }
                            Write-Host "[$timestamp] POST /run -> Job laeuft bereits" -ForegroundColor Yellow
                        } else {
                            $script:JobStartedBySchedule = $false
                            $started = Start-ExportJob
                            if ($started) {
                                Send-Json -Response $response -Data @{ status = 'started' }
                                Write-Host "[$timestamp] POST /run -> Job gestartet (ID: $($script:CurrentJob.Id))" -ForegroundColor Green
                            } else {
                                Send-Json -Response $response -Data @{ error = 'Job konnte nicht gestartet werden.' }
                            }
                        }
                    }
                }

                'GET /status' {
                    if (-not $authenticated) {
                        Send-Json -Response $response -Data @{ error = 'Nicht angemeldet' }
                    } else {
                        $historyData = @($script:RunHistory | Select-Object -Last 10)
                        $schedLastRun = if ($script:Schedule.LastRun) { $script:Schedule.LastRun.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
                        $source = Get-JobSource

                        if ($script:CurrentJob -and $script:CurrentJob.State -eq 'Running') {
                            $prog = Get-JobProgress
                            $elapsedSpan = (Get-Date) - $script:JobStartTime
                            $elapsedStr = '{0:00}:{1:00}:{2:00}' -f [math]::Floor($elapsedSpan.TotalHours), $elapsedSpan.Minutes, $elapsedSpan.Seconds

                            Send-Json -Response $response -Data @{
                                state             = 'running'
                                progress          = $prog.Progress
                                progress_pct      = $prog.ProgressPct
                                elapsed           = $elapsedStr
                                history           = $historyData
                                schedule_last_run = $schedLastRun
                            }
                        }
                        elseif ($script:CurrentJob -and $script:CurrentJob.State -eq 'Completed') {
                            $elapsedSpan = $script:CurrentJob.PSEndTime - $script:JobStartTime
                            $elapsedStr = '{0:00}:{1:00}:{2:00}' -f [math]::Floor($elapsedSpan.TotalHours), $elapsedSpan.Minutes, $elapsedSpan.Seconds
                            $excelFile = Find-LatestExcel
                            Add-HistoryEntry -State 'completed' -Duration $elapsedStr -Source $source
                            $historyData = @($script:RunHistory | Select-Object -Last 10)

                            # Zeitplaner: automatische Email nach Abschluss
                            if ($script:JobStartedBySchedule -and $script:Schedule.EmailTo -and $SmtpServer -and $SmtpFrom) {
                                $schedEmailFile = Find-LatestExcel
                                if ($schedEmailFile) {
                                    try {
                                        $datum = Get-Date -Format 'yyyy-MM-dd'
                                        Send-MailMessage `
                                            -From $SmtpFrom `
                                            -To $script:Schedule.EmailTo `
                                            -Subject "OS4K Port & Lizenz Export - $datum (Zeitplaner)" `
                                            -Body "Anbei der automatische OpenScape 4000 Port & Lizenz Export vom $datum.`n`nDiese Email wurde automatisch vom os4k2excel Zeitplaner versendet." `
                                            -SmtpServer $SmtpServer `
                                            -Port $SmtpPort `
                                            -Attachments $schedEmailFile.FullName `
                                            -Encoding ([System.Text.Encoding]::UTF8) `
                                            -ErrorAction Stop
                                        Write-Host "[$timestamp] ZEITPLANER: Email an $($script:Schedule.EmailTo) gesendet" -ForegroundColor Magenta
                                    } catch {
                                        Write-Host "[$timestamp] ZEITPLANER: Email-Fehler: $($_.Exception.Message)" -ForegroundColor Red
                                    }
                                }
                                $script:JobStartedBySchedule = $false
                            }

                            Send-Json -Response $response -Data @{
                                state             = 'completed'
                                progress          = 'Export abgeschlossen'
                                progress_pct      = 100
                                elapsed           = $elapsedStr
                                has_file          = ($null -ne $excelFile)
                                history           = $historyData
                                schedule_last_run = $schedLastRun
                            }
                        }
                        elseif ($script:CurrentJob -and $script:CurrentJob.State -eq 'Failed') {
                            $errorMsg = ''
                            try { $errorMsg = ($script:CurrentJob.ChildJobs[0].Error | Out-String) } catch {}
                            $elapsedSpan = (Get-Date) - $script:JobStartTime
                            $elapsedStr = '{0:00}:{1:00}:{2:00}' -f [math]::Floor($elapsedSpan.TotalHours), $elapsedSpan.Minutes, $elapsedSpan.Seconds
                            Add-HistoryEntry -State 'error' -Duration $elapsedStr -Source $source
                            $historyData = @($script:RunHistory | Select-Object -Last 10)
                            $script:JobStartedBySchedule = $false

                            Send-Json -Response $response -Data @{
                                state             = 'error'
                                progress          = if ($errorMsg) { $errorMsg.Substring(0, [math]::Min(200, $errorMsg.Length)) } else { 'Unbekannter Fehler' }
                                progress_pct      = 100
                                elapsed           = $elapsedStr
                                history           = $historyData
                                schedule_last_run = $schedLastRun
                            }
                        }
                        else {
                            $excelFile = Find-LatestExcel
                            Send-Json -Response $response -Data @{
                                state             = 'idle'
                                has_file          = ($null -ne $excelFile)
                                history           = $historyData
                                schedule_last_run = $schedLastRun
                            }
                        }
                    }
                }

                'GET /download' {
                    if (-not $authenticated) {
                        Send-Redirect -Response $response -Location '/'
                    } else {
                        $excelFile = Find-LatestExcel
                        if ($excelFile) {
                            $response.StatusCode = 200
                            $response.ContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            $response.AppendHeader('Content-Disposition', "attachment; filename=`"$($excelFile.Name)`"")
                            $fileBytes = [System.IO.File]::ReadAllBytes($excelFile.FullName)
                            $response.ContentLength64 = $fileBytes.Length
                            $response.OutputStream.Write($fileBytes, 0, $fileBytes.Length)
                            $response.OutputStream.Close()
                            Write-Host "[$timestamp] GET /download -> $($excelFile.Name) ($([math]::Round($fileBytes.Length/1KB)) KB)" -ForegroundColor Green
                        } else {
                            Send-Response -Response $response -Body '{"error":"Keine Excel-Datei vorhanden"}' -ContentType 'application/json' -StatusCode 404
                            Write-Host "[$timestamp] GET /download -> 404" -ForegroundColor Yellow
                        }
                    }
                }

                'POST /email' {
                    if (-not $authenticated) {
                        Send-Json -Response $response -Data @{ error = 'Nicht angemeldet' }
                    } elseif ([string]::IsNullOrWhiteSpace($SmtpServer)) {
                        Send-Json -Response $response -Data @{ error = 'SMTP ist nicht konfiguriert.' }
                    } else {
                        $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
                        $body = $reader.ReadToEnd()
                        $reader.Close()

                        try {
                            $emailData = $body | ConvertFrom-Json
                        } catch {
                            Send-Json -Response $response -Data @{ error = 'Ungueltiges Format.' }
                            Write-Host "[$timestamp] POST /email -> Ungueltiges JSON" -ForegroundColor Red
                            continue
                        }

                        $emailTo = $emailData.to
                        if ([string]::IsNullOrWhiteSpace($emailTo)) {
                            Send-Json -Response $response -Data @{ error = 'Keine Empfaenger-Adresse angegeben.' }
                        } else {
                            $excelFile = Find-LatestExcel
                            if (-not $excelFile) {
                                Send-Json -Response $response -Data @{ error = 'Keine Excel-Datei zum Versenden vorhanden.' }
                            } else {
                                try {
                                    $datum = Get-Date -Format 'yyyy-MM-dd'
                                    Send-MailMessage `
                                        -From $SmtpFrom `
                                        -To $emailTo `
                                        -Subject "OS4K Port & Lizenz Export - $datum" `
                                        -Body "Anbei der aktuelle OpenScape 4000 Port & Lizenz Export vom $datum.`n`nDiese Email wurde automatisch von os4k2excel versendet." `
                                        -SmtpServer $SmtpServer `
                                        -Port $SmtpPort `
                                        -Attachments $excelFile.FullName `
                                        -Encoding ([System.Text.Encoding]::UTF8) `
                                        -ErrorAction Stop
                                    Send-Json -Response $response -Data @{ message = "Email an $emailTo gesendet." }
                                    Write-Host "[$timestamp] POST /email -> Email an $emailTo gesendet" -ForegroundColor Green
                                } catch {
                                    Send-Json -Response $response -Data @{ error = "Email-Versand fehlgeschlagen: $($_.Exception.Message)" }
                                    Write-Host "[$timestamp] POST /email -> Fehler: $($_.Exception.Message)" -ForegroundColor Red
                                }
                            }
                        }
                    }
                }

                'POST /schedule' {
                    if (-not $authenticated) {
                        Send-Json -Response $response -Data @{ error = 'Nicht angemeldet' }
                    } else {
                        $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
                        $body = $reader.ReadToEnd()
                        $reader.Close()

                        try {
                            $schedData = $body | ConvertFrom-Json
                        } catch {
                            Send-Json -Response $response -Data @{ error = 'Ungueltiges Format.' }
                            continue
                        }

                        # Validierung
                        $schedTime = "$($schedData.time)"
                        if ($schedTime -notmatch '^\d{2}:\d{2}$') {
                            Send-Json -Response $response -Data @{ error = 'Ungueltige Uhrzeit. Format: HH:MM' }
                        } else {
                            $validDays = @('Mo','Di','Mi','Do','Fr','Sa','So')
                            $selectedDays = @($schedData.days | Where-Object { $_ -in $validDays })
                            $schedMode = if ($schedData.mode -eq 'monthly') { 'monthly' } else { 'weekly' }
                            $schedDayOfMonth = [math]::Max(1, [math]::Min(28, [int]$schedData.dayOfMonth))

                            $script:Schedule.Enabled    = [bool]$schedData.enabled
                            $script:Schedule.Mode       = $schedMode
                            $script:Schedule.Time       = $schedTime
                            $script:Schedule.Days       = $selectedDays
                            $script:Schedule.DayOfMonth = $schedDayOfMonth
                            $script:Schedule.EmailTo    = "$($schedData.emailTo)"

                            $statusText = if ($script:Schedule.Enabled) {
                                if ($schedMode -eq 'monthly') {
                                    "Zeitplaner aktiviert: $($script:Schedule.Time) Uhr (monatlich am $($schedDayOfMonth).)"
                                } else {
                                    $daysList = ($script:Schedule.Days -join ', ')
                                    "Zeitplaner aktiviert: $($script:Schedule.Time) Uhr ($daysList)"
                                }
                            } else {
                                "Zeitplaner deaktiviert."
                            }

                            Send-Json -Response $response -Data @{ message = $statusText }
                            Write-Host "[$timestamp] POST /schedule -> $statusText" -ForegroundColor Cyan
                        }
                    }
                }

                'GET /schedule' {
                    if (-not $authenticated) {
                        Send-Json -Response $response -Data @{ error = 'Nicht angemeldet' }
                    } else {
                        Send-Json -Response $response -Data @{
                            enabled    = $script:Schedule.Enabled
                            mode       = $script:Schedule.Mode
                            time       = $script:Schedule.Time
                            days       = $script:Schedule.Days
                            dayOfMonth = $script:Schedule.DayOfMonth
                            emailTo    = $script:Schedule.EmailTo
                            last_run   = if ($script:Schedule.LastRun) { $script:Schedule.LastRun.ToString('yyyy-MM-dd HH:mm:ss') } else { $null }
                        }
                    }
                }

                'POST /shutdown' {
                    if (-not $authenticated) {
                        Send-Json -Response $response -Data @{ error = 'Nicht angemeldet' }
                    } else {
                        Send-Json -Response $response -Data @{ message = 'Server wird beendet...' }
                        Write-Host "[$timestamp] POST /shutdown -> Server wird beendet (via Dashboard)" -ForegroundColor Yellow
                        $script:ShutdownRequested = $true
                    }
                }

                'GET /logout' {
                    if ($token -and $script:Sessions.ContainsKey($token)) {
                        $script:Sessions.Remove($token)
                    }
                    $response.AppendHeader('Set-Cookie', 'session=; Path=/; HttpOnly; Max-Age=0')
                    Send-Redirect -Response $response -Location '/'
                    Write-Host "[$timestamp] GET /logout -> Abgemeldet" -ForegroundColor Gray
                }

                default {
                    Send-Response -Response $response -Body '<!DOCTYPE html><html><body><h1>404</h1><p>Seite nicht gefunden.</p><a href="/">Zur Startseite</a></body></html>' -StatusCode 404
                    Write-Host "[$timestamp] $method $path -> 404" -ForegroundColor Yellow
                }
            }
        }
        catch {
            Write-Host "[$timestamp] FEHLER bei $method ${path}: $($_.Exception.Message)" -ForegroundColor Red
            try {
                Send-Response -Response $response -Body '{"error":"Interner Serverfehler"}' -ContentType 'application/json' -StatusCode 500
            } catch {}
        }
    }
}
finally {
    Write-Host "`nServer wird beendet..." -ForegroundColor Yellow
    $listener.Stop()
    $listener.Close()
    if ($script:CurrentJob) {
        Stop-Job -Job $script:CurrentJob -ErrorAction SilentlyContinue
        Remove-Job -Job $script:CurrentJob -Force -ErrorAction SilentlyContinue
    }
    # Stop-Datei aufraeumen falls vorhanden
    if (Test-Path $script:StopFilePath) {
        Remove-Item $script:StopFilePath -Force -ErrorAction SilentlyContinue
    }
    Write-Host "Server beendet." -ForegroundColor Green
}

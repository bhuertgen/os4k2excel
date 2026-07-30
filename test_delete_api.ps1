<#
.SYNOPSIS
    test_delete_api.ps1 - OS4K-3 DELETE API Validierungsscript

.DESCRIPTION
    Testet die kritischen API-Verhalten fuer NSt-Loesung (OS4K-3)

.PARAMETER ApiHost
    Hostname/IP des OpenScape 4000 Managers (PFLICHT)

.PARAMETER ApiUser
    Benutzername fuer API (PFLICHT)

.PARAMETER ApiPassword
    Passwort fuer API (PFLICHT)

.PARAMETER TestExtension
    Test-Rufnummer (PFLICHT)

.PARAMETER TestUniqueKey
    unique_key aus PORT-Tabelle (PFLICHT)

.PARAMETER OutputPath
    Zielverzeichnis fuer Ausgabedateien
    Default: $PSScriptRoot

.PARAMETER DryRun
    Switch: Nur Dateien generieren, keine API-Aufrufe

.EXAMPLE
    .\test_delete_api.ps1 -ApiHost "192.0.2.100" -ApiUser "admin" -ApiPassword "secret" `
        -TestExtension "999" -TestUniqueKey "12345"
#>

param (
    [string]$ApiPath = 'C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe',
    [string]$ApiHost = '',
    [string]$ApiUser = '',
    [string]$ApiPassword = '',
    [string]$TestExtension = '',
    [string]$Domain = '',
    [string]$Switch = '',
    [string]$TestUniqueKey = '',
    [string]$OutputPath = $PSScriptRoot,
    [switch]$DryRun
)

# PARAMETER-VALIDIERUNG
if ([string]::IsNullOrWhiteSpace($ApiHost) -or `
    [string]::IsNullOrWhiteSpace($ApiUser) -or `
    [string]::IsNullOrWhiteSpace($ApiPassword) -or `
    [string]::IsNullOrWhiteSpace($TestExtension)) {
    Write-Host "FEHLER: Pflichtparameter fehlen!" -ForegroundColor Red
    Write-Host "Erforderlich: -ApiHost, -ApiUser, -ApiPassword, -TestExtension" -ForegroundColor Red
    exit 1
}

# Wenn TestUniqueKey nicht vorhanden, aber Domain und Switch vorhanden: unique_key ermitteln
if ([string]::IsNullOrWhiteSpace($TestUniqueKey)) {
    if ([string]::IsNullOrWhiteSpace($Domain) -or [string]::IsNullOrWhiteSpace($Switch)) {
        Write-Host "FEHLER: Entweder -TestUniqueKey ODER beide (-Domain UND -Switch) erforderlich!" -ForegroundColor Red
        Write-Host "Option 1: .\test_delete_api.ps1 -TestExtension '25037' -TestUniqueKey '1615167' ..." -ForegroundColor Yellow
        Write-Host "Option 2: .\test_delete_api.ps1 -TestExtension '25037' -Domain 'lev' -Switch '0001' ..." -ForegroundColor Yellow
        exit 1
    }
    Write-Host "INFO: unique_key nicht übergeben - wird aus Domain/Switch ermittelt" -ForegroundColor Cyan
}

if (-not (Test-Path $ApiPath)) {
    Write-Host "FEHLER: api2hipath.exe nicht gefunden: $ApiPath" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# INITIALISIERUNG
$ScriptStartTime = Get-Date
$AktuellesDatum = $ScriptStartTime.ToString("yyyy-MM-dd")
$LogFilePath = Join-Path $OutputPath "test_delete_${AktuellesDatum}.log"

"=========================================================" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
"OS4K-3 DELETE API VALIDIERUNGSSCRIPT" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
"Skript gestartet: $ScriptStartTime" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
"Test-Extension: $TestExtension | Test-UniqueKey: $TestUniqueKey" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
"DryRun: $DryRun" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
"=========================================================" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

# HILFSFUNKTION: ORDER_ID aus Timestamp generieren (Prüfsumme aus Datum+Uhrzeit+Sekunden+Millisekunden)
function New-OrderID {
    $now = Get-Date
    # Format: DDHHMMSSMM (10 digits) -> take first 8 digits
    # DD = Day (01-31), HH = Hour (00-23), MM = Minute (00-59), SS = Second (00-59), MM = Milliseconds/10 (00-99)
    $allDigits = "{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}" -f `
        $now.Day, $now.Hour, $now.Minute, $now.Second, [int]($now.Millisecond / 10)

    $checksum = $allDigits.Substring(0, 8)
    return "OS4K$checksum"
}

# HILFSFUNKTION: API-Aufruf loggen und anzeigen
function Invoke-ApiWithLogging {
    param(
        [string]$Label,
        [string]$ApiPath,
        [array]$ArgumentList,
        [string]$LogPath
    )

    # Vollständigen Befehl zusammenstellen (maskiert - ohne Passwort)
    $cmdPath = "`"$ApiPath`""
    $maskedArgs = $ArgumentList | ForEach-Object {
        if ($_ -eq $ApiPassword) { "***" } else { $_ }
    }
    $maskedCmdArgs = ($maskedArgs | ForEach-Object {
        if ($_ -match '\s') { "`"$_`"" } else { $_ }
    }) -join ' '
    $maskedCommand = "$cmdPath $maskedCmdArgs"

    # Zu Konsole und Log ausgeben
    Write-Host "   [API-AUFRUF]" -ForegroundColor Cyan
    Write-Host "   Kommando: $maskedCommand" -ForegroundColor DarkCyan
    "[API-AUFRUF] $Label" | Out-File -FilePath $LogPath -Append -Encoding UTF8
    "Kommando: $maskedCommand" | Out-File -FilePath $LogPath -Append -Encoding UTF8

    # Argument-Array für Debugging (maskiert)
    "Argumente (Array):" | Out-File -FilePath $LogPath -Append -Encoding UTF8
    for ($i = 0; $i -lt $maskedArgs.Count; $i++) {
        "  [$i] = '$($maskedArgs[$i])'" | Out-File -FilePath $LogPath -Append -Encoding UTF8
    }

    # API-Aufruf ausführen mit Ausgabeerfassung
    $tempOutput = [System.IO.Path]::GetTempFileName()
    $tempError = [System.IO.Path]::GetTempFileName()

    $process = Start-Process -FilePath $ApiPath `
        -ArgumentList $ArgumentList `
        -PassThru -NoNewWindow -Wait `
        -RedirectStandardOutput $tempOutput `
        -RedirectStandardError $tempError

    # Ausgabe erfassen und ausgeben
    $output = Get-Content -Path $tempOutput -Raw -ErrorAction SilentlyContinue
    $errorOutput = Get-Content -Path $tempError -Raw -ErrorAction SilentlyContinue

    # Ausgabe zur Konsole
    if ($output) {
        Write-Host $output -ForegroundColor Gray
    }
    if ($errorOutput) {
        Write-Host $errorOutput -ForegroundColor Yellow
    }

    # Ausgabe ins Log
    if ($output) {
        $output | Out-File -FilePath $LogPath -Append -Encoding UTF8
    }
    if ($errorOutput) {
        $errorOutput | Out-File -FilePath $LogPath -Append -Encoding UTF8
    }

    # Temp-Dateien löschen
    Remove-Item -Path $tempOutput -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $tempError -Force -ErrorAction SilentlyContinue

    return $process.ExitCode
}

# HILFSFUNKTION: Eingabedatei ohne BOM erstellen
function Write-ApiInputFile {
    param(
        [string]$FilePath,
        [string]$Content,
        [string]$LogPath = ""
    )

    # UTF8 ohne BOM
    $utf8NoBOM = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($FilePath, $Content, $utf8NoBOM)

    # Datei-Info ins Log schreiben
    Write-Host "   INPUT-Datei: $FilePath" -ForegroundColor Gray
    $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $hexDump = ($fileBytes | Select-Object -First 50 | ForEach-Object { $_.ToString("X2") }) -join " "
    Write-Host "   Hex-Dump (erste 50 Bytes): $hexDump" -ForegroundColor DarkGray

    if ($LogPath) {
        "   INPUT-Datei: $FilePath" | Out-File -FilePath $LogPath -Append -Encoding UTF8
        "   Hex-Dump (erste 50 Bytes): $hexDump" | Out-File -FilePath $LogPath -Append -Encoding UTF8
    }
}

Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host "OS4K-3: DELETE API VALIDIERUNGSSCRIPT" -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host ""

# GLOBALE ORDER_ID für GESAMTEN Lösch-Auftrag (wird in ALLEN Phasen verwendet)
$orderID = New-OrderID
Write-Host "ORDER_ID für diesen Lösch-Auftrag: $orderID" -ForegroundColor Cyan
"ORDER_ID für gesamten Lösch-Auftrag: $orderID" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

# GLOBALE Flag: Master SA erkannt (durch Status 064 in ACTION_CONTROL)
$script:masterSADetected = $false

$TestErgebnisse = @()

# PHASE 1: PORT-EXISTENZ PRUEFEN
Write-Host "[PHASE 1] PORT-Existenz pruefen..." -ForegroundColor Yellow
"[PHASE 1] PORT-Existenz pruefen..." | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

$portCsvPath = Join-Path $OutputPath "test_delete_p1_port_select.csv"

# Bestimme WHERE-Klausel basierend auf verfügbaren Parametern
if (-not [string]::IsNullOrWhiteSpace($TestUniqueKey)) {
    # Option 1: Mit unique_key (schneller)
    $whereClausePORTPhase1 = "unique_key='$TestUniqueKey'"
    Write-Host "   Suche mit unique_key: '$TestUniqueKey'" -ForegroundColor Gray
} else {
    # Option 2: Mit extension, domain, switch_name (benutzerfreundlicher)
    $whereClausePORTPhase1 = "extension='$TestExtension' AND domain='$Domain' AND switch_name='$Switch'"
    Write-Host "   Suche mit Extension/Domain/Switch: '$TestExtension' / '$Domain' / '$Switch'" -ForegroundColor Gray
}

$apiArgs = @("-l", $ApiUser, "-p", $ApiPassword, "-h", $ApiHost, `
             "-o", "PORT", "-s", "unique_key,extension,domain,switch_name", `
             "-c", "|", "-w", $whereClausePORTPhase1, $portCsvPath)

$exitCode = Invoke-ApiWithLogging -Label "PHASE 1: PORT SELECT" -ApiPath $ApiPath -ArgumentList $apiArgs -LogPath $LogFilePath

if ($exitCode -eq 0 -and (Test-Path $portCsvPath)) {
    $portData = Import-Csv -Path $portCsvPath -Delimiter "|"
    if ($portData.Count -gt 0 -or $null -ne $portData) {
        # Extrahiere PORT-Daten für spätere WHERE-Klauseln
        if ($portData -is [array]) {
            $script:portExtension = $portData[0].extension
            $script:portDomain = $portData[0].domain
            $script:portSwitchName = $portData[0].switch_name
            $script:portUniqueKey = $portData[0].unique_key
        } else {
            $script:portExtension = $portData.extension
            $script:portDomain = $portData.domain
            $script:portSwitchName = $portData.switch_name
            $script:portUniqueKey = $portData.unique_key
        }

        # Wenn unique_key nicht übergeben wurde, verwende den ermittelten Wert
        if ([string]::IsNullOrWhiteSpace($TestUniqueKey)) {
            $TestUniqueKey = $script:portUniqueKey
            Write-Host "   unique_key ermittelt: $TestUniqueKey" -ForegroundColor Cyan
        }

        Write-Host "OK: NSt $script:portExtension (unique_key: $TestUniqueKey) existiert" -ForegroundColor Green
        Write-Host "   Domain: $script:portDomain | Switch: $script:portSwitchName" -ForegroundColor Gray
        "OK: NSt existiert ($script:portExtension)" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
        "   Domain: $script:portDomain | Switch: $script:portSwitchName" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

        $TestErgebnisse += @{
            Phase = "1: PORT-Existenz"
            Status = "BESTANDEN"
            Details = "NSt existiert in PORT-Tabelle"
        }
    } else {
        Write-Host "FEHLER: NSt nicht gefunden in PORT-Tabelle" -ForegroundColor Red
        "FEHLER: NSt nicht in PORT-Tabelle" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
        exit 1
    }
} else {
    Write-Host "FEHLER: api2hipath.exe Fehler (Exit-Code: $exitCode)" -ForegroundColor Red
    "FEHLER: API-Fehler Exit-Code $exitCode" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
    exit 1
}

Write-Host ""

# PHASE 4: DELETE PORT
Write-Host "[PHASE 4] DELETE PORT (löscht automatisch alle Abhängigkeiten)..." -ForegroundColor Yellow
"[PHASE 4] DELETE PORT" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

$deleteInputPath = Join-Path $OutputPath "test_delete_p4_input_port_direct.txt"

$inputContent = "unique_key|`n$TestUniqueKey|`n"
Write-ApiInputFile -FilePath $deleteInputPath -Content $inputContent -LogPath $LogFilePath

"INPUT-Datei erstellt: $deleteInputPath" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

if (-not $DryRun) {
    Write-Host "   Rufe api2hipath.exe mit -d auf..." -ForegroundColor Gray

    # Verwende GLOBALE ORDER_ID (bereits generiert am Script-Anfang)
    $apiArgs = @("-d", "-l", $ApiUser, "-p", $ApiPassword, "-h", $ApiHost, `
                 "-o", "PORT", "-c", "|", "-r", "`"$orderID`"", $deleteInputPath)

    $exitCode = Invoke-ApiWithLogging -Label "PHASE 4: DELETE PORT" -ApiPath $ApiPath -ArgumentList $apiArgs -LogPath $LogFilePath

    "DELETE PORT: Exit-Code $exitCode | ORDER_ID=$orderID" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
    $script:phase4ExitCode = $exitCode

    if ($exitCode -eq 0) {
        Write-Host "OK: Exit-Code 0 - NSt erfolgreich geloescht" -ForegroundColor Green
        "OK: NSt erfolgreich geloescht (Regulaere NSt, kein Master SA)" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
        $TestErgebnisse += @{
            Phase = "4: DELETE PORT"
            Status = "ERFOLG"
            Details = "NSt automatisch geloescht"
        }
    } else {
        Write-Host "FEHLER: Exit-Code $exitCode - DELETE fehlgeschlagen" -ForegroundColor Yellow
        Write-Host "   Könnte Master SA sein - Phase 5 prüft..." -ForegroundColor Yellow
        "FEHLER: DELETE fehlgeschlagen - prüfe auf Master SA" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
        $TestErgebnisse += @{
            Phase = "4: DELETE PORT"
            Status = "FEHLER"
            Details = "DELETE fehlgeschlagen (könnte Master SA sein)"
        }
    }
} else {
    Write-Host "   [DryRun] api2hipath.exe wird uebersprungen" -ForegroundColor Cyan
    "   [DryRun] api2hipath.exe Aufruf uebersprungen" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
    $TestErgebnisse += @{
        Phase = "4: DELETE PORT direkt"
        Status = "DRY-RUN"
        Details = "Nicht ausgefuehrt"
    }
}

Write-Host ""

# PHASE 5: MASTER SA PRUEFEN (wenn Phase 4 fehlgeschlagen ODER Status 064 erkannt)
if ($script:phase4ExitCode -ne 0 -or $script:masterSADetected) {
    Write-Host "[PHASE 5] HUNTGRP Type pruefen (Master SA Detection)..." -ForegroundColor Yellow
    "[PHASE 5] Phase 4 fehlgeschlagen - prüfe auf Master SA" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

    Write-Host ""
    Write-Host "WARNUNG: DELETE fehlgeschlagen - könnte Master SA sein" -ForegroundColor Yellow
    Write-Host "   Extension: $TestExtension (unique_key: $TestUniqueKey)" -ForegroundColor Yellow
    Write-Host "   ORDER_ID: $orderID" -ForegroundColor Yellow
    Write-Host ""
    $fortsetzung = Read-Host "Fortfahren mit Master SA Erkennung? (j/n)"

    if ($fortsetzung -ne "j" -and $fortsetzung -ne "J") {
        Write-Host "Abgebrochen" -ForegroundColor Red
        "PHASE 5 abgebrochen vom Benutzer" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
        exit 0
    }

    Write-Host ""

    # PRÜFUNG: Ist die NSt eine Master SA?
    Write-Host "   Prüfe ob NSt eine Master SA ist (type=001)..." -ForegroundColor Gray
    "   Prüfe HUNTGRP ob type=001 (Master)" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

    $huntgrpTypeCsvPath = Join-Path $OutputPath "test_delete_p5_huntgrp_type.csv"

    if (-not $DryRun) {
        $apiArgs = @("-l", $ApiUser, "-p", $ApiPassword, "-h", $ApiHost, `
                     "-o", "HUNTGRP", "-s", "extension,domain,switch_name,type,huntgrpnum", `
                     "-c", "|", "-w", "extension='$TestExtension'", $huntgrpTypeCsvPath)

        $exitCode = Invoke-ApiWithLogging -Label "PHASE 5: HUNTGRP SELECT (type check)" -ApiPath $ApiPath -ArgumentList $apiArgs -LogPath $LogFilePath

        $isMasterSA = $false
        if ($exitCode -eq 0 -and (Test-Path $huntgrpTypeCsvPath)) {
            $huntgrpTypeData = Import-Csv -Path $huntgrpTypeCsvPath -Delimiter "|"
            if ($null -ne $huntgrpTypeData) {
                if ($huntgrpTypeData -is [array]) {
                    $isMasterSA = $huntgrpTypeData | Where-Object { $_.type -eq "001" }
                } else {
                    $isMasterSA = if ($huntgrpTypeData.type -eq "001") { $huntgrpTypeData } else { $null }
                }

                if ($isMasterSA) {
                    Write-Host "   ERKANNT: Master SA (type=001)" -ForegroundColor Yellow
                    "   ERKANNT: Master SA - wird nicht automatisch aus Hunt-Group geloescht" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

                    Write-Host "   Generiere .mac Datei für manuelle Master SA Löschung..." -ForegroundColor Yellow
                    "   Generiere .mac Datei statt DELETE zu versuchen" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

                    $macFileName = "Loesche-SA-$($script:portDomain)-$($script:portSwitchName)-$TestExtension.mac"
                    $macFilePath = Join-Path $OutputPath $macFileName

                    $macContent = @"
/* Loesche-SA fuer NSt $TestExtension (Master SA) */
/* Domain: $($script:portDomain) | Switch: $($script:portSwitchName) */
/* ORDER_ID: $orderID */
/* Generiert am: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') */
/* */
/* Diesen Befehl muss in ComWin manuell ausgefuehrt werden, da dieser Anschluss */
/* eine Master SA ist und nicht automatisch aus der Hunt-Group geloescht werden kann. */
/* */
/* Schritte in ComWin: */
/*   1. Gehe zu "Verwaltung" -> "Benutzer und Anlagen" -> "Anlagenmanagement" */
/*   2. Suche NSt $TestExtension */
/*   3. Wende den folgenden Befehl an */
/* */

LOESCHEN-SA:TYP=SA,RNR=$TestExtension,DIENST=SPR;
"@

                    $utf8NoBOM = New-Object System.Text.UTF8Encoding($false)
                    [System.IO.File]::WriteAllText($macFilePath, $macContent, $utf8NoBOM)

                    Write-Host "   .mac Datei generiert: $macFileName" -ForegroundColor Cyan
                    "   .mac Datei erstellt: $macFilePath" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

                    $TestErgebnisse += @{
                        Phase = "5: Master SA Detection"
                        Status = "MASTER_SA_ERKANNT"
                        Details = ".mac Datei generiert für manuelle Löschung"
                    }
                } else {
                    Write-Host "   OK: Reguläre NSt (type=000)" -ForegroundColor Green
                    "   OK: Regulaere NSt - Port wurde automatisch in Phase 4 geloescht" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
                    $TestErgebnisse += @{
                        Phase = "5: Master SA Detection"
                        Status = "REGULAERE_NST"
                        Details = "Keine manuelle Aktion erforderlich"
                    }
                }
            }
        }
    } else {
        Write-Host "   [DryRun] HUNTGRP Type Check uebersprungen" -ForegroundColor Cyan
        "   [DryRun] HUNTGRP SELECT uebersprungen" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
    }
}

Write-Host ""

# PHASE 6: ACTION_CONTROL POLLING
Write-Host "[PHASE 6] ACTION_CONTROL Polling..." -ForegroundColor Yellow
"[PHASE 6] ACTION_CONTROL Polling" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

if (-not $DryRun -and $script:phase4ExitCode -eq 0) {
    Write-Host "   Polling bis action_status='067'..." -ForegroundColor Gray

    $actionCtrlCsvPath = Join-Path $OutputPath "test_delete_p6_action_ctrl.csv"
    $pollingTimeout = 70
    $pollingInterval = 2
    $elapsed = 0
    $done = $false

    while ($elapsed -lt $pollingTimeout -and -not $done) {
        Start-Sleep -Seconds $pollingInterval
        $elapsed += $pollingInterval

        # Suche nach aktuelle ORDER_ID (nicht mit Wildcard - API unterstützt kein 'matches')
        # Alle Felder abrufen mit "*" für komplettes Audit-Trail
        $apiArgs = @("-l", $ApiUser, "-p", $ApiPassword, "-h", $ApiHost, `
                     "-o", "ACTION_CONTROL", `
                     "-s", "*", `
                     "-c", "|", `
                     "-w", "order_id='$orderID'", `
                     $actionCtrlCsvPath)

        $exitCode = Invoke-ApiWithLogging -Label "PHASE 6: ACTION_CONTROL POLL ($elapsed`s)" -ApiPath $ApiPath -ArgumentList $apiArgs -LogPath $LogFilePath

        if (Test-Path $actionCtrlCsvPath) {
            $actionData = Import-Csv -Path $actionCtrlCsvPath -Delimiter "|"

            if ($actionData -ne $null) {
                # Prüfe nur Hauptoperation (wo unique_key != "0") auf Status 067
                # Nebenoperationen haben unique_key="0", Hauptoperation hat unique_key=action_id
                if ($actionData -is [array]) {
                    # Hauptoperation ist die mit unique_key != "0"
                    $isDone = $actionData | Where-Object { ($_.unique_key -ne "0") -and ($_.action_status -eq "067") }
                } else {
                    $isDone = if (($actionData.unique_key -ne "0") -and ($actionData.action_status -eq "067")) { $actionData } else { $null }
                }

                if ($isDone) {
                    $done = $true
                    Write-Host "OK: Status 067 (Done) erreicht nach ${elapsed}s" -ForegroundColor Green
                    "OK: ACTION_CONTROL Status 067 nach ${elapsed}s" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

                    $errorCode = if ($isDone -is [array]) { $isDone[0].action_error } else { $isDone.action_error }
                    if ($errorCode -eq "065") {
                        Write-Host "OK: RMX-Synchronisierung erfolgreich (Error-Code 065)" -ForegroundColor Green
                        "OK: RMX-Sync erfolgreich" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
                        $TestErgebnisse += @{
                            Phase = "6: ACTION_CONTROL Polling"
                            Status = "BESTANDEN"
                            Details = "RMX-Sync erfolgreich"
                        }
                    } else {
                        Write-Host "WARNUNG: RMX-Fehler Error-Code $errorCode" -ForegroundColor Yellow
                        "WARNUNG: RMX-Fehler action_error=$errorCode" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
                        $TestErgebnisse += @{
                            Phase = "6: ACTION_CONTROL Polling"
                            Status = "WARNUNG"
                            Details = "RMX-Fehler (Error-Code $errorCode)"
                        }
                    }
                }
            }
        }
    }

    if (-not $done) {
        Write-Host "TIMEOUT: ACTION_CONTROL Status 067 nicht erreicht (${pollingTimeout}s)" -ForegroundColor Yellow
        "TIMEOUT bei ACTION_CONTROL Polling nach ${pollingTimeout}s" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
        $TestErgebnisse += @{
            Phase = "6: ACTION_CONTROL Polling"
            Status = "TIMEOUT"
            Details = "Status 067 nicht erreicht"
        }
    }
} else {
    Write-Host "   [SKIPPED] ACTION_CONTROL Polling nur nach erfolgreicher DELETE" -ForegroundColor Cyan
    "   [SKIPPED] Keine DELETE ausgefuehrt" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
}

# ZUSATZ: Prüfe auf Status 064 (Master SA Indikator) in ACTION_CONTROL
if (-not $DryRun -and $script:phase4ExitCode -eq 0) {
    Write-Host ""
    Write-Host "[ZUSATZ] Prüfe ACTION_CONTROL auf Status 064 (Master SA Indikator)..." -ForegroundColor Gray
    "   [ZUSATZ] Prüfe auf Status 064 in ACTION_CONTROL" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

    # Versuche die letzte ACTION_CONTROL CSV zu laden
    $actionCtrlCsvPath = Join-Path $OutputPath "test_delete_p6_action_ctrl.csv"

    if (Test-Path $actionCtrlCsvPath) {
        $actionData = Import-Csv -Path $actionCtrlCsvPath -Delimiter "|"

        if ($null -ne $actionData) {
            $has064 = $false

            # Prüfe nur Hauptoperation (wo unique_key != "0")
            if ($actionData -is [array]) {
                $has064 = @($actionData | Where-Object { ($_.unique_key -ne "0") -and ($_.action_status -eq "064") }).Count -gt 0
            } else {
                $has064 = ($actionData.unique_key -ne "0") -and ($actionData.action_status -eq "064")
            }

            if ($has064) {
                Write-Host "   [!] WARNUNG: Hauptoperation hat Status 064 - könnte Master SA sein!" -ForegroundColor Yellow
                "   [!] Hauptoperation Status 064 erkannt - Master SA Warnung" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
                $script:masterSADetected = $true
            }
        }
    }
}

Write-Host ""

# PHASE 7: ACTION_CONTROL LOG für ORDER_ID abfragen
Write-Host "[PHASE 7] ACTION_CONTROL Log-Abfrage für ORDER_ID..." -ForegroundColor Yellow
"[PHASE 7] ACTION_CONTROL Log für ORDER_ID: $orderID" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

if (-not $DryRun) {
    Write-Host "   Abfrage ACTION_CONTROL für ORDER_ID: $orderID" -ForegroundColor Gray

    $actionLogPath = Join-Path $OutputPath "test_delete_p7_action_log_${orderID}.csv"

    # Alle Felder abrufen mit "*" für komplettes Audit-Trail
    $apiArgs = @("-l", $ApiUser, "-p", $ApiPassword, "-h", $ApiHost, `
                 "-o", "ACTION_CONTROL", `
                 "-s", "*", `
                 "-c", "|", `
                 "-w", "order_id='$orderID'", `
                 $actionLogPath)

    $exitCode = Invoke-ApiWithLogging -Label "PHASE 7: ACTION_CONTROL LOG Query" -ApiPath $ApiPath -ArgumentList $apiArgs -LogPath $LogFilePath

    if ($exitCode -eq 0 -and (Test-Path $actionLogPath)) {
        $actionLogData = Import-Csv -Path $actionLogPath -Delimiter "|"

        if ($null -ne $actionLogData) {
            Write-Host "   ACTION_CONTROL Eintraege gefunden:" -ForegroundColor Green
            "   ACTION_CONTROL Eintraege gefunden:" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

            if ($actionLogData -is [array]) {
                foreach ($entry in $actionLogData) {
                    $statusStr = switch ($entry.action_status) {
                        "067" { "[+] FERTIG (067)" }
                        default { "Status: $($entry.action_status)" }
                    }

                    $errorStr = switch ($entry.action_error) {
                        "065" { "[OK] RMX-Sync (065)" }
                        "000" { "[OK] (000)" }
                        default { "[!] Fehler: $($entry.action_error)" }
                    }

                    Write-Host "     Extension: $($entry.extension) | $statusStr | $errorStr | Update: $($entry.act_time_update)" -ForegroundColor Gray
                    "     Extension: $($entry.extension) | Status: $($entry.action_status) | Error: $($entry.action_error) | Update: $($entry.act_time_update)" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
                }
            } else {
                $statusStr = switch ($actionLogData.action_status) {
                    "067" { "[+] FERTIG (067)" }
                    default { "Status: $($actionLogData.action_status)" }
                }

                $errorStr = switch ($actionLogData.action_error) {
                    "065" { "[OK] RMX-Sync (065)" }
                    "000" { "[OK] (000)" }
                    default { "[!] Fehler: $($actionLogData.action_error)" }
                }

                Write-Host "     Extension: $($actionLogData.extension) | $statusStr | $errorStr | Update: $($actionLogData.act_time_update)" -ForegroundColor Gray
                "     Extension: $($actionLogData.extension) | Status: $($actionLogData.action_status) | Error: $($actionLogData.action_error) | Update: $($actionLogData.act_time_update)" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
            }

            Write-Host "   Log-CSV: $actionLogPath" -ForegroundColor Cyan
            "   Log-CSV erstellt: $actionLogPath" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

            $TestErgebnisse += @{
                Phase = "7: ACTION_CONTROL Log"
                Status = "ABFRAGE_ERFOLGREICH"
                Details = "Log für ORDER_ID: $orderID"
            }
        } else {
            Write-Host "   KEINE Eintraege für ORDER_ID: $orderID gefunden" -ForegroundColor Yellow
            "   KEINE ACTION_CONTROL Eintraege für ORDER_ID: $orderID" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
            $TestErgebnisse += @{
                Phase = "7: ACTION_CONTROL Log"
                Status = "KEINE_EINTRAEGE"
                Details = "Keine Logs für ORDER_ID: $orderID"
            }
        }
    } else {
        Write-Host "   FEHLER bei ACTION_CONTROL Abfrage (Exit-Code: $exitCode)" -ForegroundColor Red
        "   ACTION_CONTROL Abfrage Fehler: Exit-Code $exitCode" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
        $TestErgebnisse += @{
            Phase = "7: ACTION_CONTROL Log"
            Status = "FEHLER"
            Details = "API-Fehler Exit-Code $exitCode"
        }
    }
} else {
    Write-Host "   [DryRun] ACTION_CONTROL Log-Abfrage uebersprungen" -ForegroundColor Cyan
    "   [DryRun] ACTION_CONTROL Log-Abfrage uebersprungen" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
    $TestErgebnisse += @{
        Phase = "7: ACTION_CONTROL Log"
        Status = "DRY-RUN"
        Details = "Nicht ausgefuehrt"
    }
}

Write-Host ""

# ZUSAMMENFASSUNG
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host "TESTERGEBNISSE ZUSAMMENFASSUNG" -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host ""

"=========================================================" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
"TESTERGEBNISSE ZUSAMMENFASSUNG" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
"=========================================================" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

foreach ($result in $TestErgebnisse) {
    Write-Host "[$($result.Phase)] $($result.Status)" -ForegroundColor White
    Write-Host "   $($result.Details)" -ForegroundColor Gray
    "[$($result.Phase)] $($result.Status): $($result.Details)" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
}

Write-Host ""
Write-Host "Ausgabedateien im Verzeichnis:" -ForegroundColor Cyan
Write-Host "  $OutputPath" -ForegroundColor White
Write-Host ""
Write-Host "Log-Datei:" -ForegroundColor Cyan
Write-Host "  $LogFilePath" -ForegroundColor White
Write-Host ""

$ScriptEndTime = Get-Date
$Duration = $ScriptEndTime - $ScriptStartTime

Write-Host "Laufzeit: $($Duration.TotalSeconds) Sekunden" -ForegroundColor Gray

"" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
"Laufzeit: $($Duration.TotalSeconds)s" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
"Skript beendet: $ScriptEndTime" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
"=========================================================" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8

Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host ""

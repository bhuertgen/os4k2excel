# .SYNOPSIS
#    os4k2excel - OpenScape 4000 Port & Lizenz Export
#    Version: M30.20260730.1649
#
#    Verarbeitet PORT-Daten aus der OpenScape 4000 API und erstellt eine Excel-Datei mit verknüpften Daten.
#
# .DESCRIPTION
#    Dieses Skript ruft Daten aus den Tabellen SWITCH, NUMBERING_PLAN, CFW, PORT, HUNTGRP, DEVCONST,
#    optional HUNTGRP_SERVICE, PERSPORT und PEN ab, verknüpft sie und exportiert die Ergebnisse in eine Excel-Datei.
#    Die Standorte (Domain/Switch-Paare) werden automatisch aus der SWITCH-Tabelle ermittelt.
#
#    Lizenz-Berechnung (Basis-Wert):
#      - pen leer → $null
#      - devtext = "RADIO" oder "EXTLINE" → 0
#      - devtext = "BASEST" → 4
#      - devtext = "SET600" → 2
#      - sonst → 1
#      - Zählung nur 1x pro Nebenstelle.
#
#    Lizenz-Aufteilung (Neu):
#      - Wenn connect_type == "IP2" → Wert in "Flex_Lizenz", TDM = 0/$null
#      - Sonst → Wert in "TDM_Lizenz", Flex = 0/$null
#
#    Anpassung History:
#      2025-06-30: Initiale Version
#      2025-11-28: Fixes (Lizenz, XML-Fehler, Farben).
#      2025-11-28: Neu: Sammelanschluss Auto-Fill Name (Master) + Graue Schrift.
#      2025-12-05: Neu: Trennung Flex vs. TDM Lizenzen + Dashboard Erweiterung.
#      2026-02-15: Standorte automatisch aus SWITCH-Tabelle (API) statt Hardcodierung.
#      2026-03-17: Neu: PEN/HVT-Zuordnung via -IncludePenData (OS4K-4).
#      2026-03-24: Neu: Interaktive Credential-Abfrage wenn -ApiUser/-ApiPassword fehlen (OS4K-5).
#
# .PARAMETER ApiPath
#    Pfad zur API-Executable.
#
# .PARAMETER ApiHost
#    IP-Adresse des API-Hosts.
#
# .PARAMETER ApiPassword
#    API-Passwort. Kann mit Wert (-ApiPassword geheim) oder ohne Wert
#    (-ApiPassword) angegeben werden; ohne Wert und bei komplett fehlendem
#    Parameter erfolgt eine maskierte Abfrage.
#
# .PARAMETER OutputPath
#    Zielverzeichnis.
#
# .PARAMETER ApiUser
#    API-Benutzername (z. B. 'engr').
#
# .PARAMETER MarkDuplicate
#    Schalter (Switch). Wenn gesetzt, werden doppelte PENs im Gesamt-Tab orange markiert.
#
# .PARAMETER ShowSecrets
#    Schalter (Switch). Wenn gesetzt, wird sip_secret im Klartext exportiert. Ohne diesen Parameter wird '***' angezeigt.

param (
    # -ApiPassword darf wahlweise mit oder ohne Wert angegeben werden.
    # PowerShell kann das für einen [string]-Parameter nicht leisten: dort führt
    # ein Aufruf ohne Wert zum Bindungsfehler, bevor das Skript startet
    # (GitHub-Issue #3). Deshalb ist der Schalter selbst ein [switch], der Wert
    # wird positional gebunden:
    #   -ApiPassword geheim   -> Wert 'geheim' wird verwendet
    #   -ApiPassword          -> kein Wert -> maskierte Abfrage
    #   (ganz weggelassen)    -> maskierte Abfrage
    #
    # $ApiPasswordWert steht bewusst an erster Stelle, damit es ohne
    # [Parameter(Position=0)] auf Position 0 liegt. Ein Parameter-Attribut würde
    # das Skript zur Advanced Function machen, wodurch -Debug mit dem
    # gleichnamigen Common-Parameter kollidiert und das Skript nicht mehr startet.
    [string]$ApiPasswordWert = '',
    [switch]$ApiPassword,

    [string]$ApiPath = 'C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe',
    [string]$ApiHost = '',
    [string]$ApiUser = '',
    [string]$OutputPath = $PSScriptRoot,
    [switch]$MarkDuplicate,
    [switch]$ShowSecrets,
    [switch]$IncludePenData,
    [switch]$Debug
)

# Klartext-Passwort: aus dem positionalen Wert übernommen, ggf. später
# durch die maskierte Abfrage gefüllt.
$ApiPasswordKlartext = $ApiPasswordWert

# Plausibilitätsprüfung: ein Wert ohne den zugehörigen Schalter deutet auf ein
# versehentlich angehängtes Argument hin.
if (-not $ApiPassword -and -not [string]::IsNullOrWhiteSpace($ApiPasswordWert)) {
    Write-Host ""
    Write-Host "FEHLER: Unerwartetes Argument '$ApiPasswordWert'." -ForegroundColor Red
    Write-Host "Das Passwort muss mit -ApiPassword <PASSWORT> angegeben werden." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

# --- Debug-Modus aktivieren ---
$DiagnosticMode = $Debug

# --- Modulprüfung: ImportExcel ---
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host ""
    Write-Host "FEHLER: Das PowerShell-Modul 'ImportExcel' ist nicht installiert." -ForegroundColor Red
    Write-Host ""
    Write-Host "Das Modul wird fuer die Excel-Erstellung benoetigt." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "=== Online-Installation (mit Internetzugang) ===" -ForegroundColor Cyan
    Write-Host "  Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor White
    Write-Host ""
    Write-Host "=== Offline-Installation (ohne Internetzugang) ===" -ForegroundColor Cyan
    Write-Host "  1. Auf einem PC MIT Internet ausfuehren:" -ForegroundColor White
    Write-Host "     Save-Module ImportExcel -Path C:\Temp\Modules" -ForegroundColor Gray
    Write-Host "  2. Ordner 'C:\Temp\Modules\ImportExcel' auf den Zielrechner kopieren nach:" -ForegroundColor White
    Write-Host "     $($env:USERPROFILE)\Documents\WindowsPowerShell\Modules\ImportExcel" -ForegroundColor Gray
    Write-Host "     (PS 5.1)" -ForegroundColor DarkGray
    Write-Host "     oder" -ForegroundColor White
    Write-Host "     $($env:USERPROFILE)\Documents\PowerShell\Modules\ImportExcel" -ForegroundColor Gray
    Write-Host "     (PS 7+)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "Nach der Installation das Script erneut starten." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

# --- Parameterprüfung ---

# [1] ApiHost ist immer Pflicht — keine interaktive Abfrage
if ([string]::IsNullOrWhiteSpace($ApiHost)) {
    Write-Host ""
    Write-Host "FEHLER: -ApiHost muss angegeben werden." -ForegroundColor Red
    Write-Host ""
    Write-Host "Verwendung:" -ForegroundColor Yellow
    Write-Host "  .\os4k2excel.ps1 -ApiHost <IP> -ApiUser <BENUTZER> -ApiPassword <PASSWORT> [Optionen]" -ForegroundColor White
    Write-Host ""
    Write-Host "Pflichtparameter:" -ForegroundColor Yellow
    Write-Host "  -ApiHost        IP-Adresse des OpenScape 4000" -ForegroundColor Gray
    Write-Host "  -ApiUser        API-Benutzername (interaktiv wenn nicht angegeben)" -ForegroundColor Gray
    Write-Host "  -ApiPassword    API-Passwort. Ohne Wert oder ganz weggelassen wird es" -ForegroundColor Gray
    Write-Host "                  maskiert abgefragt." -ForegroundColor Gray
    Write-Host ""
    Write-Host "Optionale Parameter:" -ForegroundColor Yellow
    Write-Host "  -ApiPath        Pfad zu api2hipath.exe (Standard: C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe)" -ForegroundColor Gray
    Write-Host "  -OutputPath     Zielverzeichnis (Standard: Skriptverzeichnis)" -ForegroundColor Gray
    Write-Host "  -MarkDuplicate  Doppelte PENs im Gesamt-Tab orange markieren" -ForegroundColor Gray
    Write-Host "  -ShowSecrets    SIP-Secrets im Klartext exportieren (Standard: maskiert als ***)" -ForegroundColor Gray
    Write-Host "  -IncludePenData PEN/HVT-Daten abfragen und exportieren (Hauptverteiler-Zuordnung)" -ForegroundColor Gray
    Write-Host "  -Debug          Detaillierte Diagnose-Ausgaben ins Log schreiben" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Beispiele:" -ForegroundColor Yellow
    Write-Host "  .\os4k2excel.ps1 -ApiHost 10.10.1.1 -ApiUser engr -ApiPassword MeinPasswort -MarkDuplicate" -ForegroundColor White
    Write-Host "  .\os4k2excel.ps1 -ApiHost 10.10.1.1 -ApiUser engr -ApiPassword -ShowSecrets" -ForegroundColor White
    Write-Host "     (Passwort wird maskiert abgefragt)" -ForegroundColor DarkGray
    Write-Host ""
    exit 1
}

# [2] Umgebungs-Detektion: Interaktive Shell?
$IsInteractiveSession = [Environment]::UserInteractive -and ($Host.Name -eq 'ConsoleHost' -or $Host.Name -eq 'Windows PowerShell ISE Host' -or $Host.Name -eq 'Visual Studio Code Host')

# [3] ApiUser — interaktiv abfragen wenn nicht angegeben
if ([string]::IsNullOrWhiteSpace($ApiUser)) {
    if (-not $IsInteractiveSession) {
        Write-Host ""
        Write-Host "FEHLER: -ApiUser muss in nicht-interaktiven Umgebungen als Parameter angegeben werden." -ForegroundColor Red
        Write-Host ""
        exit 1
    }
    Write-Host ""
    Write-Host "API-Benutzername fuer $ApiHost eingeben:" -ForegroundColor Yellow
    $ApiUser = Read-Host "ApiUser"
    if ([string]::IsNullOrWhiteSpace($ApiUser)) {
        Write-Host "FEHLER: Benutzername darf nicht leer sein." -ForegroundColor Red
        exit 1
    }
}

# [4] ApiPassword — interaktiv maskiert abfragen wenn kein Wert angegeben wurde.
# Greift sowohl bei '-ApiPassword' ohne Wert als auch bei komplett fehlendem
# Parameter (GitHub-Issue #3).
if ([string]::IsNullOrWhiteSpace($ApiPasswordKlartext)) {
    if (-not $IsInteractiveSession) {
        Write-Host ""
        Write-Host "FEHLER: -ApiPassword muss in nicht-interaktiven Umgebungen mit Wert angegeben werden." -ForegroundColor Red
        Write-Host ""
        exit 1
    }
    Write-Host ""
    Write-Host "API-Passwort fuer $ApiUser@$ApiHost eingeben (Eingabe wird maskiert):" -ForegroundColor Yellow
    $SecurePassword = Read-Host "ApiPassword" -AsSecureString
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    try {
        $ApiPasswordKlartext = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)
    } finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
    }
    if ([string]::IsNullOrWhiteSpace($ApiPasswordKlartext)) {
        Write-Host "FEHLER: Passwort darf nicht leer sein." -ForegroundColor Red
        exit 1
    }
}

# --- Version ---
$ScriptVersion = "M30.20260730.1649"
Write-Host "os4k2excel Version: $ScriptVersion" -ForegroundColor Cyan

# --- PERFORMANCE: Stapelgröße erhöhen ---
$Global:MaximumFunctionCount = 32768
$Global:MaximumVariableCount = 32768

# --- UTF-8 für korrekte Sonderzeichen ---
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

# --- CSV-Import mit Zeichensatz-Erkennung ---
# api2hipath.exe schreibt seine Exporte in Windows-1252. Ohne explizites Encoding
# liest Import-Csv sie als UTF-8, wodurch jeder Umlaut zu U+FFFD wird und der
# ursprüngliche Buchstabe unwiederbringlich verloren geht (GitHub-Issue #2).
#
# Statt ein Encoding fest vorzugeben, wird der Inhalt geprüft: Was gültiges UTF-8
# ist, wird als UTF-8 gelesen, alles andere als Windows-1252. Damit bleiben reine
# ASCII-Tabellen und künftige UTF-8-Exporte korrekt, ohne dass die Erkennung
# angepasst werden muss.
function Import-ApiCsv {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [string]$Delimiter = "|"
    )

    if (-not (Test-Path $Path)) { return @() }

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -eq 0) { return @() }

    $hatBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)

    if ($hatBom) {
        $encoding = New-Object System.Text.UTF8Encoding($false)
        $encodingName = "UTF-8 (BOM)"
    } else {
        # Strikte UTF-8-Prüfung: wirft bei Windows-1252-Umlauten eine Ausnahme
        try {
            $strikt = New-Object System.Text.UTF8Encoding($false, $true)
            [void]$strikt.GetString($bytes)
            $encoding = New-Object System.Text.UTF8Encoding($false)
            $encodingName = "UTF-8"
        } catch {
            $encoding = [System.Text.Encoding]::GetEncoding(1252)
            $encodingName = "Windows-1252"
        }
    }

    $text = $encoding.GetString($bytes).TrimStart([char]0xFEFF)

    if ($script:DiagnosticMode -and $script:LogFilePath) {
        "DIAGNOSE: $(Split-Path $Path -Leaf): Zeichensatz erkannt als $encodingName" |
            Out-File -FilePath $script:LogFilePath -Append
    }

    $zeilen = @($text -split "`r`n|`n|`r" | Where-Object { $_ -ne '' })
    if ($zeilen.Count -lt 2) { return @() }

    return @($zeilen | ConvertFrom-Csv -Delimiter $Delimiter)
}

# --- Datum für Dateinamen ---
$aktuellesDatum = Get-Date -Format "yyyy-MM-dd"
Write-Host "Aktuelles Datum: $aktuellesDatum" -ForegroundColor Green

# --- Ausgabeverzeichnis prüfen ---
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Write-Host "Verzeichnis erstellt: $OutputPath" -ForegroundColor Yellow
}

# --- Pfade für Excel und Log ---
$ExcelFilePath = Join-Path $OutputPath "OS4K-PORT-$aktuellesDatum.xlsx"
$LogFilePath = Join-Path $OutputPath "OS4K-PORT-$aktuellesDatum.log"
Write-Host "Ausgabeverzeichnis: $((Resolve-Path $OutputPath).Path)" -ForegroundColor Green
Write-Host "Excel: $ExcelFilePath" -ForegroundColor Green
Write-Host "Log:   $LogFilePath" -ForegroundColor Green

# --- Bestehende Excel-Datei entfernen (verhindert Dateisperren) ---
if (Test-Path $ExcelFilePath) {
    try {
        Remove-Item $ExcelFilePath -Force -ErrorAction Stop
        Write-Host "Bestehende Excel-Datei entfernt: $ExcelFilePath" -ForegroundColor Yellow
    } catch {
        Write-Host "FEHLER: Excel-Datei ist gesperrt oder kann nicht geloescht werden: $ExcelFilePath" -ForegroundColor Red
        Write-Host "Bitte die Datei in Excel schliessen und das Skript erneut starten." -ForegroundColor Red
        exit 1
    }
}

# --- Laufzeitmessung ---
$scriptStartTime = Get-Date
Write-Host "`n" -ForegroundColor White
Write-Host "════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "                    PHASE 1: INITIALISIERUNG" -ForegroundColor Cyan
Write-Host "════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "Skript gestartet: $scriptStartTime" -ForegroundColor Cyan
"Skript gestartet: $scriptStartTime" | Out-File -FilePath $LogFilePath -Append
$ImportExcelVersion = (Get-Module -ListAvailable -Name ImportExcel | Select-Object -First 1).Version.ToString()
"os4k2excel Version: $ScriptVersion" | Out-File -FilePath $LogFilePath -Append
"PowerShell Version: $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))" | Out-File -FilePath $LogFilePath -Append
"ImportExcel Version: $ImportExcelVersion" | Out-File -FilePath $LogFilePath -Append
"Betriebssystem: $([System.Environment]::OSVersion.VersionString)" | Out-File -FilePath $LogFilePath -Append
Write-Host "PowerShell: $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition)) | ImportExcel: $ImportExcelVersion" -ForegroundColor Gray

# --- API-Abfragen: Tabellen und Felder ---
$tables = @(
    @{ Name = "SWITCH"; Fields = "unique_key,switch_name,switch_typ,version,info,country,areacode,netcode,domain" },
    @{ Name = "NUMBERING_PLAN"; Fields = "domain,info,isdn_ac,isdn_cc,isdn_lc,isdn_skip_digits,isdn_ul,nodetrk,switch_name,unique_key,virtual_node_id,virtual_node_num" },
    @{ Name = "CFW"; Fields = "extension,dest,domain,switch_name,activate,name,dtype,itype,service,variant" },
    @{ Name = "HUNTGRP"; Fields = "domain,switch_name,huntgrpnum,type,extension,position,service,search_criteria,name,info,group_idx" },
    @{ Name = "HUNTGRP_SERVICE"; Fields = "domain,switch_name,group_idx,huntgrp_node_id,overflow_ext" },
    @{ Name = "PICKUPGRP"; Fields = "domain,switch_name,pickupgrpnum,displ,info" },
    @{ Name = "PICKUP_SUB"; Fields = "domain,switch_name,extension,pickupgrpnum" },
    @{ Name = "PERSPORT"; Fields = "domain,switch_name,extension,connect_type" },
    @{ Name = "DEVCONST"; Fields = "adaptor_1,adaptor_2,addon_text,addon_typ,chargeindic,cnt_of_devices,consistent,devconame,devtext1,devtext2,devtext3,devtext4,devtext5,devtext6,devtyp1,devtyp2,devtyp3,devtyp4,devtyp5,devtyp6,display,exact_dev_text,exact_dev_typ,funct_included,handsfree,headset,idcard_rd,info,opticom,recorder,sys_verify,unique_key" }
)

# --- PEN-Abfrage nur wenn -IncludePenData ---
if ($IncludePenData) {
    $tables += @{ Name = "PEN"; Fields = "unique_key,pen,switch_name,cardtyp,status,info,node_1,subnode_1,line_1,strip_1,pin_1,node_2,subnode_2,line_2,strip_2,pin_2,ctr_for_busy,domain,reserved,extension,pen_num,ltu,lin1,lin2,board_present,reserved_time,ltg,slot,circuit" }
    Write-Host "  -IncludePenData aktiv: PEN-Tabelle wird abgefragt" -ForegroundColor Cyan
    "-IncludePenData aktiv: PEN wird zur Abfrage hinzugefügt" | Out-File -FilePath $LogFilePath -Append
}

Write-Host "`n════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "                  PHASE 2: API-ABFRAGEN (Batch)" -ForegroundColor Cyan
Write-Host "════════════════════════════════════════════════════════════════`n" -ForegroundColor Cyan

# --- API-Abfragen ausführen ---
foreach ($t in $tables) {
    $tableName = $t.Name
    $fields = $t.Fields
    Write-Host "Starte API-Abfrage für $tableName..." -ForegroundColor Yellow
    "${tableName}: API-Abfrage gestartet, Felder: $fields" | Out-File -FilePath $LogFilePath -Append

    $apiStart = Get-Date
    $csvPath = Join-Path $OutputPath "OS4K-$tableName-$aktuellesDatum.csv"

    # API-Prozess mit Exit-Code-Überprüfung
    # -Wait ist zwingend: ohne diesen Schalter liefert $process.ExitCode unter
    # Windows PowerShell 5.1 immer $null (siehe GitHub-Issue #1).
    $process = Start-Process -FilePath $ApiPath -ArgumentList "-l","$ApiUser","-p","$ApiPasswordKlartext","-h","$ApiHost","-o","$tableName","-s","$fields","-c","|","$csvPath" -PassThru -NoNewWindow -Wait
    $exitCode = $process.ExitCode

    $apiEnd = Get-Date

    Write-Host "API-Abfrage für $tableName abgeschlossen. Dauer: $($apiEnd - $apiStart) (Exit-Code: $exitCode)" -ForegroundColor Green
    "${tableName}: API-Abfrage abgeschlossen. Dauer: $($apiEnd - $apiStart), Exit-Code: $exitCode" | Out-File -FilePath $LogFilePath -Append

    # --- Diagnose: Prüfe ob CSV existiert und wie groß sie ist (nur mit -Debug) ---
    if (Test-Path $csvPath) {
        if ($DiagnosticMode) {
            $fileSize = (Get-Item $csvPath).Length
            $fileLines = @(Get-Content $csvPath).Count
            "DIAGNOSE: ${tableName}: CSV existiert, Größe: $fileSize Bytes, Zeilen: $fileLines" | Out-File -FilePath $LogFilePath -Append
            Write-Host "  → CSV: $fileSize Bytes, ca. $fileLines Zeilen" -ForegroundColor Gray
        }
    } else {
        "FEHLER: ${tableName}: CSV-Datei nicht gefunden: $csvPath" | Out-File -FilePath $LogFilePath -Append
        Write-Host "  → FEHLER: CSV-Datei nicht gefunden!" -ForegroundColor Red
        if ($exitCode -ne 0) {
            "  Exit-Code war: $exitCode (wahrscheinlich Fehler bei API-Aufruf)" | Out-File -FilePath $LogFilePath -Append
        }
    }
}

Write-Host "`n════════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "                   PHASE 3: CSV-DATEN LADEN" -ForegroundColor Green
Write-Host "════════════════════════════════════════════════════════════════`n" -ForegroundColor Green

# --- CSV-Dateien importieren ---
Write-Host "Importiere CSV-Dateien..." -ForegroundColor Yellow
$importStart = Get-Date

$tableSwitch = @(if (Test-Path (Join-Path $OutputPath "OS4K-SWITCH-$aktuellesDatum.csv")) { Import-ApiCsv -Path (Join-Path $OutputPath "OS4K-SWITCH-$aktuellesDatum.csv") -Delimiter "|" })
$table2 = @(if (Test-Path (Join-Path $OutputPath "OS4K-NUMBERING_PLAN-$aktuellesDatum.csv")) { Import-ApiCsv -Path (Join-Path $OutputPath "OS4K-NUMBERING_PLAN-$aktuellesDatum.csv") -Delimiter "|" })
$table3 = @(if (Test-Path (Join-Path $OutputPath "OS4K-CFW-$aktuellesDatum.csv")) { Import-ApiCsv -Path (Join-Path $OutputPath "OS4K-CFW-$aktuellesDatum.csv") -Delimiter "|" })
$table4 = @(if (Test-Path (Join-Path $OutputPath "OS4K-HUNTGRP-$aktuellesDatum.csv")) { Import-ApiCsv -Path (Join-Path $OutputPath "OS4K-HUNTGRP-$aktuellesDatum.csv") -Delimiter "|" })
$table5 = @(if (Test-Path (Join-Path $OutputPath "OS4K-HUNTGRP_SERVICE-$aktuellesDatum.csv")) { Import-ApiCsv -Path (Join-Path $OutputPath "OS4K-HUNTGRP_SERVICE-$aktuellesDatum.csv") -Delimiter "|" })
$tablePickupGrp = @(if (Test-Path (Join-Path $OutputPath "OS4K-PICKUPGRP-$aktuellesDatum.csv")) { Import-ApiCsv -Path (Join-Path $OutputPath "OS4K-PICKUPGRP-$aktuellesDatum.csv") -Delimiter "|" })
$tablePickupSub = @(if (Test-Path (Join-Path $OutputPath "OS4K-PICKUP_SUB-$aktuellesDatum.csv")) { Import-ApiCsv -Path (Join-Path $OutputPath "OS4K-PICKUP_SUB-$aktuellesDatum.csv") -Delimiter "|" })
$table6 = @(if (Test-Path (Join-Path $OutputPath "OS4K-PERSPORT-$aktuellesDatum.csv")) { Import-ApiCsv -Path (Join-Path $OutputPath "OS4K-PERSPORT-$aktuellesDatum.csv") -Delimiter "|" })
$table7 = @(if (Test-Path (Join-Path $OutputPath "OS4K-DEVCONST-$aktuellesDatum.csv")) { Import-ApiCsv -Path (Join-Path $OutputPath "OS4K-DEVCONST-$aktuellesDatum.csv") -Delimiter "|" })

# --- PEN laden (nur wenn -IncludePenData) ---
$tablePenData = @()
if ($IncludePenData) {
    $penDataCsvPath = Join-Path $OutputPath "OS4K-PEN-$aktuellesDatum.csv"
    if (Test-Path $penDataCsvPath) {
        $tablePenData = @(Import-ApiCsv -Path $penDataCsvPath -Delimiter "|")
        Write-Host "  PEN: $($tablePenData.Count) Einträge geladen" -ForegroundColor Green
        "PEN: $($tablePenData.Count) Einträge geladen" | Out-File -FilePath $LogFilePath -Append
    } else {
        Write-Host "  WARNUNG: PEN-CSV nicht gefunden — HVT-Daten werden ohne PEN erstellt" -ForegroundColor Yellow
        "WARNUNG: PEN-CSV nicht gefunden: $penDataCsvPath" | Out-File -FilePath $LogFilePath -Append
    }
}

$importEnd = Get-Date
Write-Host "CSV-Import abgeschlossen. Dauer: $($importEnd - $importStart)" -ForegroundColor Green

# --- Diagnose: Zeilenzahl der geladenen Tabellen (nur mit -Debug) ---
if ($DiagnosticMode) {
    "=== DIAGNOSE: Zeilenzahl der Tabellen ===" | Out-File -FilePath $LogFilePath -Append
    "SWITCH: $($tableSwitch.Count) Zeilen" | Out-File -FilePath $LogFilePath -Append
    "NUMBERING_PLAN: $($table2.Count) Zeilen" | Out-File -FilePath $LogFilePath -Append
    "CFW: $($table3.Count) Zeilen" | Out-File -FilePath $LogFilePath -Append
    "HUNTGRP: $($table4.Count) Zeilen" | Out-File -FilePath $LogFilePath -Append
    "HUNTGRP_SERVICE: $($table5.Count) Zeilen" | Out-File -FilePath $LogFilePath -Append
    "PICKUPGRP: $($tablePickupGrp.Count) Zeilen" | Out-File -FilePath $LogFilePath -Append
    "PICKUP_SUB: $($tablePickupSub.Count) Zeilen" | Out-File -FilePath $LogFilePath -Append
    "PERSPORT: $($table6.Count) Zeilen" | Out-File -FilePath $LogFilePath -Append
    "DEVCONST: $($table7.Count) Zeilen" | Out-File -FilePath $LogFilePath -Append
    if ($IncludePenData) { "PEN: $($tablePenData.Count) Zeilen" | Out-File -FilePath $LogFilePath -Append }
    "=== ENDE DIAGNOSE ===" | Out-File -FilePath $LogFilePath -Append
}

# --- Zuordnungstabellen für lesbare Werte ---
$variantMapping = @{ "j07" = "STATION"; "j08" = "SYSTEM"; "j09" = "STATIONV" }
$serviceMapping = @{ "i90" = "VOICE"; "i91" = "AUDIO3K1"; "i92" = "FAXG23"; "i93" = "FAXG4"; "i94" = "GENSRVC"; "i95" = "SPR"; "i96" = "TEL3K1"; "i97" = "TEL7K"; "i98" = "TOENE"; "i99" = "UDI"; "j00" = "VIDEOTEL"; "j01" = "FAX"; "j02" = "DTE" }
$dtypeMapping = @{ "841" = "CFU"; "842" = "CFB"; "843" = "CFNR"; "844" = "CD"; "845" = "CFDEV"; "846" = "CFDND"; "847" = "CFBNR" }
$itypeMapping = @{ "j03" = "EXTERN"; "j04" = "INTERN"; "j05" = "GEN" }
$activateMapping = @{ "i70" = "Ausschalten"; "C56" = "Einschalten" }
$huntgrpTypeMapping = @{ "001" = "Master"; "000" = "regular hunt group" }
$connectTypeMapping = @{ "256" = "DIRECT"; "257" = "PNT"; "260" = "EXTERNAL"; "263" = "LOG"; "267" = "IP"; "268" = "IP2" }
$cardtypMapping = @{
    "300" = "SLMA"; "301" = "SLMB"; "302" = "SLMD"; "303" = "TMX21"; "304" = "SLMU"; "305" = "SLMS"; "306" = "STMD"
    "307" = "DIUS2"; "308" = "SLMQ"; "309" = "DIU24"; "310" = "DIU30"; "311" = "SLMR"; "312" = "ATI"; "313" = "OPS"
    "314" = "T1"; "315" = "RLI"; "316" = "TMD24"; "317" = "SLMAVAR"; "318" = "SLMQ_EXT"; "319" = "SLMO"
    "320" = "SLMA1"; "321" = "SLMU16"; "330" = "KLMU"; "331" = "SLCE8"; "332" = "SLC16"; "333" = "SLC16-100"
    "334" = "SLC16-200"; "335" = "SLC16-300"; "336" = "SLC24"; "337" = "SLCSM"; "340" = "SLMY"; "341" = "WAML2"
    "342" = "STMA"; "343" = "SLMAR"; "344" = "SLMO_ROUT_TEMP"; "345" = "SLMO_REM_TEMP"; "346" = "SLMO_ROUT_PERM"
    "347" = "SLMO_REM_PERM"; "348" = "STMA_PSW"; "349" = "SLMOP"; "350" = "STMD2"; "351" = "SLMS2"; "352" = "STHCO"
    "353" = "STHCD"; "355" = "DIU-N2"; "356" = "DIU-N4"; "357" = "SLMQ3"; "358" = "SLMA2"; "359" = "TMDNH"
    "360" = "STMI-HFA"; "361" = "STMI-IGW"; "362" = "APPS"; "363" = "BS22"; "364" = "CDG"; "365" = "CDG31-DE"
    "366" = "CDG31-FU"; "367" = "CDG31-RG"; "368" = "CGM&NEW"; "369" = "CSG"; "370" = "DIUC64"; "371" = "DPC5"
    "372" = "DSCX"; "373" = "LTUCX"; "374" = "MTSCG+CB"; "375" = "NCUI"; "376" = "PBCDG-DE"; "377" = "PBCDG-FU"
    "378" = "PBCDG-RG"; "379" = "PBPNE-DE"; "380" = "PBPNE-RG"; "381" = "PNEDIUS2"; "382" = "RG-USA"; "383" = "RGKOR"
    "384" = "SIUX2"; "385" = "SLMA3"; "386" = "SLMAB"; "387" = "SLMO16"; "388" = "SLOP2"; "389" = "SO-BG"
    "390" = "STHC"; "391" = "STHC2"; "392" = "TM2LP"; "393" = "TM3WI"; "394" = "TM3WO"; "395" = "TMACH"
    "396" = "TMAU"; "397" = "TMBD"; "398" = "TMBLN"; "399" = "TMCOW"; "400" = "TMEM"; "401" = "TMEW2"
    "402" = "TMLBL"; "403" = "TMLR"; "404" = "TMLRB"; "405" = "TMLSL"; "406" = "TMSFP"; "407" = "VCM-B15"
    "408" = "VCM-B7"; "409" = "PNE"; "410" = "STMI"; "411" = "STMI2-HFA"; "412" = "DIUN2U"
}
$penStatusMapping = @{ "002" = "Frei"; "003" = "Belegt" }
$reservedMapping = @{ "000" = "nein"; "001" = "ja" }

# --- Hauptfunktion: Verarbeitung der PORT-Daten ---
function Process-PortData {
    param (
        [Parameter(Mandatory=$true)][array]$DomainSwitchPairs,
        [Parameter(Mandatory=$false)][PSObject]$NumberingPlanTable,
        [Parameter(Mandatory=$false)][PSObject]$CfwTable,
        [Parameter(Mandatory=$false)][PSObject]$HuntgrpTable,
        [Parameter(Mandatory=$false)][PSObject]$HuntgrpServiceTable,
        [Parameter(Mandatory=$false)][PSObject]$PickupGrpTable,
        [Parameter(Mandatory=$false)][PSObject]$PickupSubTable,
        [Parameter(Mandatory=$false)][PSObject]$PersportTable,
        [Parameter(Mandatory=$false)][PSObject]$DevconstTable,
        [Parameter(Mandatory=$false)][array]$PenDataTable,
        [Parameter(Mandatory=$true)][string]$ExcelFilePath,
        [Parameter(Mandatory=$false)][string]$OutputPath,
        [Parameter(Mandatory=$false)][string]$LogFilePath,
        [Parameter(Mandatory=$false)][string]$ApiPath,
        [Parameter(Mandatory=$false)][string]$ApiHost,
        [Parameter(Mandatory=$false)][string]$ApiPassword,
        [Parameter(Mandatory=$false)][string]$ApiUser,
        [Parameter(Mandatory=$false)][string]$aktuellesDatum,
        [switch]$MarkDuplicate,
        [switch]$IncludePenData,
        [switch]$DiagnosticMode
    )

    $pairCount = 0
    $totalPairs = $DomainSwitchPairs.Count
    $allResults = @()
    $licenseSummary = @()

    # --- Hashtables für schnelle Verknüpfung ---
    $numberingHash = @{}
    foreach ($row in $NumberingPlanTable) { $key = "$($row.domain)|$($row.switch_name)|$($row.virtual_node_id)"; $numberingHash[$key] = $row }

    $cfwHash = @{}
    foreach ($row in $CfwTable) {
        $key = "$($row.domain)|$($row.switch_name)|$($row.extension)"
        if (-not $cfwHash.ContainsKey($key)) { $cfwHash[$key] = @() }
        $cfwHash[$key] += $row
    }

    $huntgrpHash = @{}
    foreach ($row in $HuntgrpTable) { $key = "$($row.domain)|$($row.switch_name)|$($row.huntgrpnum)"; $huntgrpHash[$key] = $row }

    # --- Hashtable: PICKUPGRP (AUN-Gruppen) nach Gruppennummer ---
    $pickupGrpHash = @{}
    if ($PickupGrpTable) {
        foreach ($row in $PickupGrpTable) {
            $key = "$($row.domain)|$($row.switch_name)|$($row.pickupgrpnum)"
            $pickupGrpHash[$key] = $row
        }
        if ($DiagnosticMode) { "DIAGNOSE: PickupGrpHash erstellt mit $($pickupGrpHash.Count) Einträgen" | Out-File -FilePath $LogFilePath -Append }
    } else {
        "FEHLER: PickupGrpTable ist leer oder null" | Out-File -FilePath $LogFilePath -Append
    }

    # --- Hashtable: Extension -> AUN-Gruppe (aus PICKUP_SUB) ---
    $extensionToAunHash = @{}
    if ($PickupSubTable) {
        foreach ($row in $PickupSubTable) {
            $key = "$($row.domain)|$($row.switch_name)|$($row.extension)"
            $extensionToAunHash[$key] = $row
        }
        if ($DiagnosticMode) { "DIAGNOSE: ExtensionToAunHash erstellt mit $($extensionToAunHash.Count) Einträgen" | Out-File -FilePath $LogFilePath -Append }
    } else {
        "FEHLER: PickupSubTable ist leer oder null" | Out-File -FilePath $LogFilePath -Append
    }

    $huntgrpServiceHash = @{}
    if ($HuntgrpServiceTable) {
        foreach ($row in $HuntgrpServiceTable) { $key = "$($row.domain)|$($row.switch_name)|$($row.group_idx)"; $huntgrpServiceHash[$key] = $row }
    }

    $persportHash = @{}
    if ($PersportTable) {
        foreach ($row in $PersportTable) { $key = "$($row.domain)|$($row.switch_name)|$($row.extension)"; $persportHash[$key] = $row }
    }

    # --- Hashtable: PEN für HVT-Zuordnung (nur wenn -IncludePenData) ---
    $penDataLookup = @{}
    $portPenLookup = @{}
    if ($IncludePenData -and $PenDataTable) {
        foreach ($row in $PenDataTable) {
            $key = "$($row.domain)|$($row.switch_name)|$($row.pen)"
            $penDataLookup[$key] = $row
        }
        if ($DiagnosticMode) { "DIAGNOSE: PenDataLookup erstellt mit $($penDataLookup.Count) Einträgen" | Out-File -FilePath $LogFilePath -Append }
    }

    # --- SCHRITT 1: Datenverarbeitung und EXPORT (Schreiben der Daten) ---
    foreach ($pair in $DomainSwitchPairs) {
        $pairCount++
        $domain = $pair.Domain
        $switchName = $pair.SwitchName
        $worksheetName = "$domain-$switchName"
        $portCsvPath = Join-Path $OutputPath "OS4K-PORT-$domain-$switchName-$aktuellesDatum.csv"

        Write-Host "[$pairCount/$totalPairs] Verarbeite $worksheetName..." -ForegroundColor Cyan

        # API Abruf mit Logging
        Write-Host "  Starte API-Abfrage für PORT ($switchName)..." -ForegroundColor Yellow
        "PORT ($switchName): API-Abfrage gestartet" | Out-File -FilePath $LogFilePath -Append

        $apiStart = Get-Date
        # -Wait ist zwingend, siehe GitHub-Issue #1 (ExitCode unter PS 5.1 sonst $null)
        $process = Start-Process -FilePath $ApiPath -ArgumentList "-l","$ApiUser","-p","$ApiPasswordKlartext","-h","$ApiHost","-o","PORT","-s","domain,switch_name,extension,e164_num,pen,displayname,devconame,devtext,status,unique_key,node_id,info,pubnum,cust_lan_ip_addr,sip_ip_address,sip_userid,sip_secret,sip_realm,lin1,lin2","-c","|","-w","switch_name='$switchName'","$portCsvPath" -PassThru -NoNewWindow -Wait
        $exitCode = $process.ExitCode
        $apiEnd = Get-Date

        Write-Host "  API-Abfrage für PORT ($switchName) abgeschlossen. Dauer: $($apiEnd - $apiStart) (Exit-Code: $exitCode)" -ForegroundColor Green
        "PORT ($switchName): API-Abfrage abgeschlossen. Dauer: $($apiEnd - $apiStart), Exit-Code: $exitCode" | Out-File -FilePath $LogFilePath -Append

        if (Test-Path $portCsvPath) {
            if ($DiagnosticMode) {
                $fileSize = (Get-Item $portCsvPath).Length
                $fileLines = @(Get-Content $portCsvPath).Count
                "DIAGNOSE: PORT ($switchName): CSV existiert, Größe: $fileSize Bytes, Zeilen: $fileLines" | Out-File -FilePath $LogFilePath -Append
                Write-Host "    → CSV: $fileSize Bytes, ca. $fileLines Zeilen" -ForegroundColor Gray
            }
        } else {
            "FEHLER: PORT ($switchName): CSV-Datei nicht gefunden: $portCsvPath" | Out-File -FilePath $LogFilePath -Append
            Write-Host "    → FEHLER: CSV-Datei nicht gefunden!" -ForegroundColor Red
            if ($exitCode -ne 0) {
                "  Exit-Code war: $exitCode (wahrscheinlich Fehler bei API-Aufruf)" | Out-File -FilePath $LogFilePath -Append
            }
        }

        $table1 = if (Test-Path $portCsvPath) { Import-ApiCsv -Path $portCsvPath -Delimiter "|" } else { @() }

        # --- lin1/lin2 für HVT-Sheet sammeln (nur wenn -IncludePenData) ---
        if ($IncludePenData) {
            foreach ($portRow in $table1) {
                if (-not [string]::IsNullOrWhiteSpace($portRow.pen)) {
                    $ppKey = "$($portRow.domain)|$($portRow.switch_name)|$($portRow.pen)"
                    $portPenLookup[$ppKey] = @{ lin1 = $portRow.lin1; lin2 = $portRow.lin2 }
                }
            }
        }

        $result = foreach ($row1 in $table1) {
            $npKey = "$($row1.domain)|$($row1.switch_name)|$($row1.node_id)"; $row2 = $numberingHash[$npKey]
            $cfwKey = "$($row1.domain)|$($row1.switch_name)|$($row1.extension)"; $rows3 = $cfwHash[$cfwKey]
            $huntgrpKey = "$($row1.domain)|$($row1.switch_name)|$($row1.extension)"; $masterSA = if ($huntgrpHash.ContainsKey($huntgrpKey)) { "JA" } else { "NEIN" }

            # --- AUN-Gruppennummer für diese Extension auslesen (max. eine) ---
            $aunGroupNum = $null
            $extensionToAunKey = "$($row1.domain)|$($row1.switch_name)|$($row1.extension)"
            if ($extensionToAunHash.ContainsKey($extensionToAunKey)) {
                $pickupSubRow = $extensionToAunHash[$extensionToAunKey]
                if ($pickupSubRow -and $pickupSubRow.pickupgrpnum) {
                    $aunGroupNum = $pickupSubRow.pickupgrpnum
                }
            }

            $persportKey = "$($row1.domain)|$($row1.switch_name)|$($row1.extension)"; $row6 = $persportHash[$persportKey]

            # Status
            $t_InBetrieb = "?"
            if (($row1.devtext -eq "ANAVAR") -and ($row1.status -like "*SIGNED_OFF*") -and ([string]::IsNullOrWhiteSpace($row1.pen))) { $t_InBetrieb = "AWTLN" }
            elseif ($row1.status -like "*GENNR*") { $t_InBetrieb = "nein" }
            elseif ($row1.status -like "*READY*" -or $row1.status -like "*DEFIL*" -or $row1.status -like "*TRS*") { $t_InBetrieb = "ja" }

            # Connect Type ermitteln
            $currentConnectType = if ($row6 -and $connectTypeMapping.ContainsKey($row6.connect_type)) { $connectTypeMapping[$row6.connect_type] } else { if ($row6) { $row6.connect_type } else { $null } }

            # 1. Basis-Lizenzwert berechnen
            $licenseBaseValue = if ([string]::IsNullOrWhiteSpace($row1.pen)) { 
                              $null 
                          } elseif ($row1.devtext -eq "RADIO" -or $row1.devtext -eq "EXTLINE") { 
                              0 
                          } elseif ($row1.devtext -eq "BASEST") { 
                              4 
                          } elseif ($row1.devtext -eq "SET600") { 
                              2 
                          } else { 
                              1 
                          }

            # 2. Aufteilung in Flex vs TDM
            # Standard: Alles ist TDM. Ausnahme: connect_type="IP2" -> Flex.
            $valFlex = $null
            $valTDM  = $null

            if ($licenseBaseValue -ne $null) {
                if ($currentConnectType -eq "IP2") {
                    $valFlex = $licenseBaseValue
                } else {
                    $valTDM = $licenseBaseValue
                }
            }

            # --- HVT-Daten nachschlagen (einmalig pro PORT-Zeile) ---
            $penRow = $null
            if ($IncludePenData -and -not [string]::IsNullOrWhiteSpace($row1.pen)) {
                $penKey = "$($row1.domain)|$($row1.switch_name)|$($row1.pen)"
                if ($penDataLookup.ContainsKey($penKey)) { $penRow = $penDataLookup[$penKey] }
                else { "WARNUNG: PEN $($row1.pen) (Extension: $($row1.extension)) nicht in PEN-Tabelle gefunden (Key: $penKey)" | Out-File -FilePath $LogFilePath -Append }
            }

            # Mit CFW
            if ($rows3) {
                $isFirstCfwRow = $true
                foreach ($row3 in $rows3) {
                    # Lizenz nur in der ersten Zeile der CFW-Gruppe zählen
                    $currFlex = if ($isFirstCfwRow) { $valFlex } else { $null }
                    $currTDM  = if ($isFirstCfwRow) { $valTDM } else { $null }
                    $isFirstCfwRow = $false

                    $obj = [ordered]@{
                        domain              = $row1.domain
                        switch_name         = $row1.switch_name
                        extension           = $row1.extension
                        e164_num            = $row1.e164_num
                        pen                 = $row1.pen
                        displayname         = $row1.displayname
                        devconame           = $row1.devconame
                        devtext             = $row1.devtext
                        status              = $row1.status
                        unique_key          = $row1.unique_key
                        info                = $row1.info
                        pubnum              = $row1.pubnum
                        node_id             = $row1.node_id
                        numbering_plan_info = if ($row2) { $row2.info } else { $null }
                        cfw_dest            = $row3.dest
                        cfw_dtype           = if ($dtypeMapping.ContainsKey($row3.dtype)) { $dtypeMapping[$row3.dtype] } else { $row3.dtype }
                        cfw_service         = if ($serviceMapping.ContainsKey($row3.service)) { $serviceMapping[$row3.service] } else { $row3.service }
                        cfw_name            = $row3.name
                        cfw_itype           = if ($itypeMapping.ContainsKey($row3.itype)) { $itypeMapping[$row3.itype] } else { $row3.itype }
                        cfw_variant         = if ($variantMapping.ContainsKey($row3.variant)) { $variantMapping[$row3.variant] } else { $row3.variant }
                        cfw_activate        = if ($activateMapping.ContainsKey($row3.activate)) { $activateMapping[$row3.activate] } else { $row3.activate }
                        connect_type        = $currentConnectType
                        t_InBetrieb         = $t_InBetrieb
                        cust_lan_ip_addr    = $row1.cust_lan_ip_addr
                        sip_ip_address      = $row1.sip_ip_address
                        sip_userid          = $row1.sip_userid
                        sip_secret          = if ($ShowSecrets) { $row1.sip_secret } elseif (-not [string]::IsNullOrWhiteSpace($row1.sip_secret)) { '***' } else { $null }
                        sip_realm           = $row1.sip_realm
                        Master_SA           = $masterSA
                        PickupGrpNum        = $aunGroupNum
                        Flex_Lizenz         = $currFlex
                        TDM_Lizenz          = $currTDM
                    }
                    if ($IncludePenData) {
                        $obj['HVT_Info']         = if ($penRow) { $penRow.info } else { $null }
                        $obj['HVT1_Knoten']      = if ($penRow) { $penRow.node_1 } else { $null }
                        $obj['HVT1_Unterknoten'] = if ($penRow) { $penRow.subnode_1 } else { $null }
                        $obj['HVT1_Bucht']       = if ($penRow) { $penRow.line_1 } else { $null }
                        $obj['HVT1_Leiste']      = if ($penRow) { $penRow.strip_1 } else { $null }
                        $obj['HVT1_Stift']       = if ($penRow) { $penRow.pin_1 } else { $null }
                        $obj['HVT2_Knoten']      = if ($penRow) { $penRow.node_2 } else { $null }
                        $obj['HVT2_Unterknoten'] = if ($penRow) { $penRow.subnode_2 } else { $null }
                        $obj['HVT2_Bucht']       = if ($penRow) { $penRow.line_2 } else { $null }
                        $obj['HVT2_Leiste']      = if ($penRow) { $penRow.strip_2 } else { $null }
                        $obj['HVT2_Stift']       = if ($penRow) { $penRow.pin_2 } else { $null }
                    }
                    [PSCustomObject]$obj
                }
            } else {
                # Ohne CFW
                $obj = [ordered]@{
                    domain              = $row1.domain
                    switch_name         = $row1.switch_name
                    extension           = $row1.extension
                    e164_num            = $row1.e164_num
                    pen                 = $row1.pen
                    displayname         = $row1.displayname
                    devconame           = $row1.devconame
                    devtext             = $row1.devtext
                    status              = $row1.status
                    unique_key          = $row1.unique_key
                    info                = $row1.info
                    pubnum              = $row1.pubnum
                    node_id             = $row1.node_id
                    numbering_plan_info = if ($row2) { $row2.info } else { $null }
                    cfw_dest            = $null
                    cfw_dtype           = $null
                    cfw_service         = $null
                    cfw_name            = $null
                    cfw_itype           = $null
                    cfw_variant         = $null
                    cfw_activate        = $null
                    connect_type        = $currentConnectType
                    t_InBetrieb         = $t_InBetrieb
                    cust_lan_ip_addr    = $row1.cust_lan_ip_addr
                    sip_ip_address      = $row1.sip_ip_address
                    sip_userid          = $row1.sip_userid
                    sip_secret          = if ($ShowSecrets) { $row1.sip_secret } elseif (-not [string]::IsNullOrWhiteSpace($row1.sip_secret)) { '***' } else { $null }
                    sip_realm           = $row1.sip_realm
                    Master_SA           = $masterSA
                    PickupGrpNum        = $aunGroupNum
                    Flex_Lizenz         = $valFlex
                    TDM_Lizenz          = $valTDM
                }
                if ($IncludePenData) {
                    $obj['HVT_Info']         = if ($penRow) { $penRow.info } else { $null }
                    $obj['HVT1_Knoten']      = if ($penRow) { $penRow.node_1 } else { $null }
                    $obj['HVT1_Unterknoten'] = if ($penRow) { $penRow.subnode_1 } else { $null }
                    $obj['HVT1_Bucht']       = if ($penRow) { $penRow.line_1 } else { $null }
                    $obj['HVT1_Leiste']      = if ($penRow) { $penRow.strip_1 } else { $null }
                    $obj['HVT1_Stift']       = if ($penRow) { $penRow.pin_1 } else { $null }
                    $obj['HVT2_Knoten']      = if ($penRow) { $penRow.node_2 } else { $null }
                    $obj['HVT2_Unterknoten'] = if ($penRow) { $penRow.subnode_2 } else { $null }
                    $obj['HVT2_Bucht']       = if ($penRow) { $penRow.line_2 } else { $null }
                    $obj['HVT2_Leiste']      = if ($penRow) { $penRow.strip_2 } else { $null }
                    $obj['HVT2_Stift']       = if ($penRow) { $penRow.pin_2 } else { $null }
                }
                [PSCustomObject]$obj
            }
        }

        $allResults += $result

        # Summen pro Standort
        $flexSum = ($result | Where-Object { $_.Flex_Lizenz -ne $null } | Measure-Object -Property Flex_Lizenz -Sum).Sum
        $flexSum = if ($null -eq $flexSum) { 0 } else { $flexSum }

        $tdmSum = ($result | Where-Object { $_.TDM_Lizenz -ne $null } | Measure-Object -Property TDM_Lizenz -Sum).Sum
        $tdmSum = if ($null -eq $tdmSum) { 0 } else { $tdmSum }

        $resultWithSum = $result + [PSCustomObject]@{ domain = "SUMME"; pubnum = $null; Flex_Lizenz = $flexSum; TDM_Lizenz = $tdmSum }
        
        # Export Standort
        $resultWithSum | Export-Excel -Path $ExcelFilePath -WorkSheetname $worksheetName -ClearSheet -NoNumberConversion * -BoldTopRow -AutoSize -AutoFilter -TitleBold -TableStyle Medium5 -FreezeTopRow

        # Dashboard Daten sammeln
        $licenseSummary += [PSCustomObject]@{ 
            Standort          = $worksheetName
            Flex_Lizenzen     = $flexSum
            TDM_Lizenzen      = $tdmSum
            Gesamt_Lizenzen   = ($flexSum + $tdmSum)
        }
    }

    # Gesamt-Summen berechnen
    $totalFlex = ($allResults | Where-Object { $_.Flex_Lizenz -ne $null } | Measure-Object -Property Flex_Lizenz -Sum).Sum
    $totalFlex = if ($null -eq $totalFlex) { 0 } else { $totalFlex }

    $totalTDM = ($allResults | Where-Object { $_.TDM_Lizenz -ne $null } | Measure-Object -Property TDM_Lizenz -Sum).Sum
    $totalTDM = if ($null -eq $totalTDM) { 0 } else { $totalTDM }

    # Export Gesamt
    $allResultsWithSum = $allResults + [PSCustomObject]@{ domain = "GESAMTSUMME"; pubnum = $null; Flex_Lizenz = $totalFlex; TDM_Lizenz = $totalTDM }
    $allResultsWithSum | Export-Excel -Path $ExcelFilePath -WorkSheetname "Gesamt" -ClearSheet -NoNumberConversion * -BoldTopRow -AutoSize -AutoFilter -TitleBold -TableStyle Medium5 -FreezeTopRow

    # Export Lizenz-Dashboard
    if ($licenseSummary.Count -gt 0) {
        $licenseSummary | Export-Excel -Path $ExcelFilePath -WorkSheetname "Lizenz Dashboard" -ClearSheet -BoldTopRow -AutoSize -AutoFilter -TableStyle Light9 -FreezeTopRow
    }

    Write-Host "Datenexport abgeschlossen. Starte Formatierung..." -ForegroundColor Yellow

    # --- SCHRITT 2: FORMATIERUNG (Datei öffnen und offen halten) ---
    $excel = Open-ExcelPackage -Path $ExcelFilePath

    # 2a. Standorte formatieren (Summenzeile gelb)
    foreach ($pair in $DomainSwitchPairs) {
        $wsName = "$($pair.Domain)-$($pair.SwitchName)"
        $ws = $excel.Workbook.Worksheets[$wsName]
        if ($ws) {
            $lastRow = $ws.Dimension.End.Row
            # Flex und TDM Spalten finden
            $flexCol = ($ws.Cells["1:1"] | Where-Object { $_.Value -eq "Flex_Lizenz" }).Start.Column
            $tdmCol  = ($ws.Cells["1:1"] | Where-Object { $_.Value -eq "TDM_Lizenz" }).Start.Column

            if ($flexCol -gt 0) {
                $ws.Row($lastRow).Style.Font.Bold = $true
                $ws.Row($lastRow).Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $ws.Row($lastRow).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightYellow)
                
                $ws.Cells[$lastRow, $flexCol].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
                if ($tdmCol -gt 0) {
                    $ws.Cells[$lastRow, $tdmCol].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
                }
            }
        }
    }

    # 2b. Gesamt formatieren
    $ws = $excel.Workbook.Worksheets["Gesamt"]
    if ($ws) {
        $lastRow = $ws.Dimension.End.Row
        
        $flexCol = ($ws.Cells["1:1"] | Where-Object { $_.Value -eq "Flex_Lizenz" }).Start.Column
        $tdmCol  = ($ws.Cells["1:1"] | Where-Object { $_.Value -eq "TDM_Lizenz" }).Start.Column
        
        if ($flexCol -gt 0) {
            $ws.Row($lastRow).Style.Font.Bold = $true
            $ws.Row($lastRow).Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $ws.Row($lastRow).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGreen)
            
            $ws.Cells[$lastRow, $flexCol].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LimeGreen)
            if ($tdmCol -gt 0) {
                $ws.Cells[$lastRow, $tdmCol].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LimeGreen)
            }
        }

        # --- Duplikats-Markierung (Nur wenn -MarkDuplicate gesetzt) ---
        if ($MarkDuplicate) {
            Write-Host "Berechne Duplikate für Farbmarkierung (Orange)..." -ForegroundColor Cyan
            $headerRow = $ws.Cells["1:1"]
            $switchColIndex = ($headerRow | Where-Object { $_.Value -eq "switch_name" }).Start.Column
            $penColIndex    = ($headerRow | Where-Object { $_.Value -eq "pen" }).Start.Column

            if ($switchColIndex -and $penColIndex -and $lastRow -gt 1) {
                $dupeLookup = @{}
                $validEntries = $allResults | Where-Object { -not [string]::IsNullOrWhiteSpace($_.pen) }
                $grouped = $validEntries | Group-Object switch_name, pen | Where-Object { $_.Count -gt 1 }
                
                foreach ($grp in $grouped) {
                    if ($grp.Group.Count -gt 0) {
                        $dupeLookup["$($grp.Group[0].switch_name)|$($grp.Group[0].pen)"] = $true
                    }
                }

                for ($r = 2; $r -lt $lastRow; $r++) {
                    $cellSwitch = $ws.Cells[$r, $switchColIndex].Text
                    $cellPen    = $ws.Cells[$r, $penColIndex].Text
                    if (-not [string]::IsNullOrWhiteSpace($cellPen)) {
                        if ($dupeLookup.ContainsKey("$cellSwitch|$cellPen")) {
                            $ws.Cells[$r, $penColIndex].Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed)
                            $ws.Cells[$r, $penColIndex].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $ws.Cells[$r, $penColIndex].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Orange)
                        }
                    }
                }
            }
        }
    }

    # 2c. Dashboard Header formatieren
    $ws = $excel.Workbook.Worksheets["Lizenz Dashboard"]
    if ($ws) {
        $ws.Cells["A1"].Value = "Standort"
        $ws.Cells["B1"].Value = "Genutzte Flex-Lizenzen"
        $ws.Cells["C1"].Value = "Genutzte TDM-Lizenzen"
        $ws.Cells["D1"].Value = "Gesamt Flex+TDM"

        # Formatierung Header (A1 bis D1)
        $ws.Cells["A1:D1"].Style.Font.Bold = $true
        $ws.Cells["A1:D1"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $ws.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::CornflowerBlue)
        $ws.Cells["A1:D1"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

        # Summenzeile unten
        $lastRow = [int]$ws.Dimension.End.Row + 1
        $ws.Cells[$lastRow, 1].Value = "GESAMT"
        $ws.Cells[$lastRow, 2].Value = $totalFlex
        $ws.Cells[$lastRow, 3].Value = $totalTDM
        $ws.Cells[$lastRow, 4].Value = ($totalFlex + $totalTDM)

        $ws.Cells[$lastRow, 1, $lastRow, 4].Style.Font.Bold = $true
        $ws.Cells[$lastRow, 2, $lastRow, 4].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $ws.Cells[$lastRow, 2, $lastRow, 4].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Gold)

        # Hinweis-Box: Lizenzkalkulation basiert auf ungefähren Näherungswerten
        $noticeStartRow = [int]($lastRow + 2)
        $noticeEndRow = [int]($noticeStartRow + 3)

        # Rahmen für die Hinweis-Box
        $ws.Cells[$noticeStartRow, 1, $noticeEndRow, 4].Style.Border.Top.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Medium
        $ws.Cells[$noticeStartRow, 1, $noticeEndRow, 4].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Medium
        $ws.Cells[$noticeStartRow, 1, $noticeEndRow, 4].Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Medium
        $ws.Cells[$noticeStartRow, 1, $noticeEndRow, 4].Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Medium

        $ws.Cells[$noticeStartRow, 1, $noticeEndRow, 4].Style.Border.Top.Color.SetColor([System.Drawing.Color]::DarkOrange)
        $ws.Cells[$noticeStartRow, 1, $noticeEndRow, 4].Style.Border.Bottom.Color.SetColor([System.Drawing.Color]::DarkOrange)
        $ws.Cells[$noticeStartRow, 1, $noticeEndRow, 4].Style.Border.Left.Color.SetColor([System.Drawing.Color]::DarkOrange)
        $ws.Cells[$noticeStartRow, 1, $noticeEndRow, 4].Style.Border.Right.Color.SetColor([System.Drawing.Color]::DarkOrange)

        # Hintergrundfarbe
        $ws.Cells[$noticeStartRow, 1, $noticeEndRow, 4].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $ws.Cells[$noticeStartRow, 1, $noticeEndRow, 4].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightYellow)

        # Zeilennummern vorberechnen
        $noticeRow1 = [int]($noticeStartRow + 1)
        $noticeRow2 = [int]($noticeStartRow + 2)
        $noticeRow3 = [int]($noticeStartRow + 3)

        # Titel der Hinweis-Box
        $ws.Cells[$noticeStartRow, 1].Value = "⚠️ HINWEIS"
        $ws.Cells[$noticeStartRow, 1, $noticeStartRow, 4].Merge = $true
        $ws.Cells[$noticeStartRow, 1].Style.Font.Bold = $true
        $ws.Cells[$noticeStartRow, 1].Style.Font.Size = 14
        $ws.Cells[$noticeStartRow, 1].Style.Font.Color.SetColor([System.Drawing.Color]::White)
        $ws.Cells[$noticeStartRow, 1].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $ws.Cells[$noticeStartRow, 1].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkOrange)
        $ws.Cells[$noticeStartRow, 1].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        $ws.Cells[$noticeStartRow, 1].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left

        # Hinweis-Text (in mehreren Zeilen)
        $ws.Cells[$noticeRow1, 1].Value = "Die Lizenzkalkulation kann nicht direkt aus der API abgerufen werden und basiert auf Näherungswerten."
        $ws.Cells[$noticeRow1, 1, $noticeRow1, 4].Merge = $true
        $ws.Cells[$noticeRow1, 1].Style.Font.Size = 11
        $ws.Cells[$noticeRow1, 1].Style.Font.Italic = $true
        $ws.Cells[$noticeRow1, 1].Style.WrapText = $true
        $ws.Cells[$noticeRow1, 1].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top

        $ws.Cells[$noticeRow2, 1].Value = "Diese können erheblich von der Anzeige des ABFRAGEN-CODEW; abweichen."
        $ws.Cells[$noticeRow2, 1, $noticeRow2, 4].Merge = $true
        $ws.Cells[$noticeRow2, 1].Style.Font.Size = 11
        $ws.Cells[$noticeRow2, 1].Style.Font.Italic = $true
        $ws.Cells[$noticeRow2, 1].Style.WrapText = $true
        $ws.Cells[$noticeRow2, 1].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top

        $ws.Cells[$noticeRow3, 1].Value = "Für genaue Lizenzinformationen wenden Sie sich bitte an Ihren Administrator."
        $ws.Cells[$noticeRow3, 1, $noticeRow3, 4].Merge = $true
        $ws.Cells[$noticeRow3, 1].Style.Font.Size = 10
        $ws.Cells[$noticeRow3, 1].Style.Font.Italic = $true
        $ws.Cells[$noticeRow3, 1].Style.WrapText = $true
        $ws.Cells[$noticeRow3, 1].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top

        # Zeilenhöhe für bessere Lesbarkeit anpassen
        $ws.Row($noticeStartRow).Height = 28
        $ws.Row($noticeRow1).Height = 40
        $ws.Row($noticeRow2).Height = 40
        $ws.Row($noticeRow3).Height = 35
    }

    # Formatierung speichern und schließen
    $formattingEnd = Get-Date
    Close-ExcelPackage $excel

    Write-Host "Lizenz-Dashboard aktualisiert. Gesamt: Flex=$totalFlex, TDM=$totalTDM (Formatierung: $($formattingEnd - $apiEnd))" -ForegroundColor Magenta
    "Lizenz-Dashboard aktualisiert. Gesamt: Flex=$totalFlex, TDM=$totalTDM (Formatierung: $($formattingEnd - $apiEnd))" | Out-File -FilePath $LogFilePath -Append

    # --- SCHRITT 3: Weitere Worksheets ---

    # Sammelanschluss
    Write-Host "Erstelle Sammelanschluss-Worksheet..." -ForegroundColor Yellow
    "Erstelle Sammelanschluss-Worksheet..." | Out-File -FilePath $LogFilePath -Append
    $sammelStart = Get-Date
    
    # Lookup für Displaynames erstellen
    $portDisplayLookup = @{}
    foreach ($item in $allResults) {
        if (-not [string]::IsNullOrWhiteSpace($item.displayname)) {
            $k = "$($item.domain)|$($item.switch_name)|$($item.extension)"
            if (-not $portDisplayLookup.ContainsKey($k)) { $portDisplayLookup[$k] = $item.displayname }
        }
    }

    $huntgrpGrayRows = @()
    $loopIndex = 0

    $huntgrpResults = foreach ($row in $HuntgrpTable) {
        
        $huntgrpServiceKey = "$($row.domain)|$($row.switch_name)|$($row.group_idx)"
        $huntgrpServiceRow = $huntgrpServiceHash[$huntgrpServiceKey]
        $numberingPlanRow = $null
        if ($huntgrpServiceRow -and $huntgrpServiceRow.huntgrp_node_id) {
            $npKey = "$($row.domain)|$($row.switch_name)|$($huntgrpServiceRow.huntgrp_node_id)"
            $numberingPlanRow = $numberingHash[$npKey]
        }
        $huntgrpE164 = $row.huntgrpnum
        if ($numberingPlanRow) {
            $huntgrpE164 = "$($numberingPlanRow.isdn_cc)$($numberingPlanRow.isdn_ac)$($numberingPlanRow.isdn_lc)$($row.huntgrpnum)"
        }

        # Name ergänzen Logik
        $currentName = $row.name
        $currentType = if ($huntgrpTypeMapping.ContainsKey($row.type)) { $huntgrpTypeMapping[$row.type] } else { $row.type }
        
        if ([string]::IsNullOrWhiteSpace($currentName) -and ($currentType -eq "Master")) {
            $checkKey = "$($row.domain)|$($row.switch_name)|$($row.extension)"
            if ($portDisplayLookup.ContainsKey($checkKey)) {
                $currentName = $portDisplayLookup[$checkKey]
                $huntgrpGrayRows += ($loopIndex + 2)
            }
        }
        $loopIndex++

        [PSCustomObject]@{
            huntgrp_domain         = $row.domain
            huntgrp_switch_name    = $row.switch_name
            huntgrp_huntgrpnum     = $row.huntgrpnum
            huntgrp_name           = $currentName
            huntgrp_type           = $currentType
            huntgrp_position       = $row.position
            huntgrp_extension      = $row.extension
            huntgrp_service        = $row.service
            huntgrp_search_criteria= $row.search_criteria
            huntgrp_info           = $row.info
            huntgrp_node_id        = if ($huntgrpServiceRow) { $huntgrpServiceRow.huntgrp_node_id } else { $null }
            huntgrp_overflow_ext   = if ($huntgrpServiceRow) { $huntgrpServiceRow.overflow_ext } else { $null }
            huntgrp_E164           = $huntgrpE164
        }
    }
    
    $huntgrpResults | Export-Excel -Path $ExcelFilePath -WorkSheetname "Sammelanschluss" -NoNumberConversion * -BoldTopRow -AutoSize -AutoFilter -TitleBold -TableStyle Medium5 -FreezeTopRow

    # Formatierung Graue Schrift Sammelanschluss
    if ($huntgrpGrayRows.Count -gt 0) {
        $excelSammel = Open-ExcelPackage -Path $ExcelFilePath
        $wsSammel = $excelSammel.Workbook.Worksheets["Sammelanschluss"]
        if ($wsSammel) {
            $nameColIndex = ($wsSammel.Cells["1:1"] | Where-Object { $_.Value -eq "huntgrp_name" }).Start.Column
            if ($nameColIndex) {
                foreach ($rIndex in $huntgrpGrayRows) {
                    $wsSammel.Cells[$rIndex, $nameColIndex].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)
                }
            }
        }
        Close-ExcelPackage $excelSammel
    }
    $sammelEnd = Get-Date
    Write-Host "  Sammelanschluss-Worksheet erstellt. Dauer: $($sammelEnd - $sammelStart)" -ForegroundColor Green
    "Sammelanschluss-Worksheet erstellt. Dauer: $($sammelEnd - $sammelStart)" | Out-File -FilePath $LogFilePath -Append

    # AUN-Gruppen (PICKUPGRP)
    Write-Host "Erstelle AUN-Gruppen-Worksheet..." -ForegroundColor Yellow
    "Erstelle AUN-Gruppen-Worksheet..." | Out-File -FilePath $LogFilePath -Append
    $pickupStart = Get-Date

    $pickupResults = foreach ($row in $PickupGrpTable) {
        [PSCustomObject]@{
            domain              = $row.domain
            switch_name         = $row.switch_name
            pickupgrpnum        = $row.pickupgrpnum
            displ               = $row.displ
            info                = $row.info
        }
    }

    if ($pickupResults) {
        $pickupResults | Export-Excel -Path $ExcelFilePath -WorkSheetname "AUN-Gruppen" -NoNumberConversion * -BoldTopRow -AutoSize -AutoFilter -TitleBold -TableStyle Medium5 -FreezeTopRow
    }
    $pickupEnd = Get-Date
    Write-Host "  AUN-Gruppen-Worksheet erstellt. Dauer: $($pickupEnd - $pickupStart)" -ForegroundColor Green
    "AUN-Gruppen-Worksheet erstellt. Dauer: $($pickupEnd - $pickupStart)" | Out-File -FilePath $LogFilePath -Append

    # NumberingPlan
    Write-Host "Erstelle NUMBERINGPLAN-Worksheet..." -ForegroundColor Yellow
    $numberingStart = Get-Date
    $numberingResults = foreach ($row in $NumberingPlanTable) {
        [PSCustomObject]@{
            domain              = $row.domain
            switch_name         = $row.switch_name
            info                = $row.info
            isdn_cc             = $row.isdn_cc
            isdn_ac             = $row.isdn_ac
            isdn_lc             = $row.isdn_lc
            isdn_skip_digits    = $row.isdn_skip_digits
            isdn_ul             = $row.isdn_ul
            nodetrk             = $row.nodetrk
            unique_key          = $row.unique_key
            virtual_node_id     = $row.virtual_node_id
            virtual_node_num    = $row.virtual_node_num
        }
    }
    $numberingResults | Export-Excel -Path $ExcelFilePath -WorkSheetname "NumberingPlan" -NoNumberConversion * -BoldTopRow -AutoSize -AutoFilter -TitleBold -TableStyle Medium5 -FreezeTopRow
    $numberingEnd = Get-Date
    Write-Host "  NUMBERINGPLAN-Worksheet erstellt. Dauer: $($numberingEnd - $numberingStart)" -ForegroundColor Green
    "NUMBERINGPLAN-Worksheet erstellt. Dauer: $($numberingEnd - $numberingStart)" | Out-File -FilePath $LogFilePath -Append

    # Devconst
    Write-Host "Erstelle DEVCONST-Worksheet..." -ForegroundColor Yellow
    $devconstStart = Get-Date
    $devconstResults = foreach ($row in $DevconstTable) {
        [PSCustomObject]@{
            adaptor_1         = $row.adaptor_1
            adaptor_2         = $row.adaptor_2
            addon_text        = $row.addon_text
            addon_typ         = $row.addon_typ
            chargeindic       = $row.chargeindic
            cnt_of_devices    = $row.cnt_of_devices
            consistent        = $row.consistent
            devconame         = $row.devconame
            devtext1          = $row.devtext1
            devtext2          = $row.devtext2
            devtext3          = $row.devtext3
            devtext4          = $row.devtext4
            devtext5          = $row.devtext5
            devtext6          = $row.devtext6
            devtyp1           = $row.devtyp1
            devtyp2           = $row.devtyp2
            devtyp3           = $row.devtyp3
            devtyp4           = $row.devtyp4
            devtyp5           = $row.devtyp5
            devtyp6           = $row.devtyp6
            display           = $row.display
            exact_dev_text    = $row.exact_dev_text
            exact_dev_typ     = $row.exact_dev_typ
            funct_included    = $row.funct_included
            handsfree         = $row.handsfree
            headset           = $row.headset
            idcard_rd         = $row.idcard_rd
            info              = $row.info
            opticom           = $row.opticom
            recorder          = $row.recorder
            sys_verify        = $row.sys_verify
            unique_key        = $row.unique_key
        }
    }
    $devconstResults | Export-Excel -Path $ExcelFilePath -WorkSheetname "Devconst" -NoNumberConversion * -BoldTopRow -AutoSize -AutoFilter -TitleBold -TableStyle Medium5 -FreezeTopRow
    $devconstEnd = Get-Date
    Write-Host "  DEVCONST-Worksheet erstellt. Dauer: $($devconstEnd - $devconstStart)" -ForegroundColor Green
    "DEVCONST-Worksheet erstellt. Dauer: $($devconstEnd - $devconstStart)" | Out-File -FilePath $LogFilePath -Append

    # HVT Sheet (nur wenn -IncludePenData)
    if ($IncludePenData -and $PenDataTable.Count -gt 0) {
        Write-Host "Erstelle HVT-Worksheet..." -ForegroundColor Yellow
        "Erstelle HVT-Worksheet..." | Out-File -FilePath $LogFilePath -Append
        $hvtStart = Get-Date

        $hvtResults = foreach ($row in $PenDataTable) {
            $resolvedCardtyp = if ($cardtypMapping.ContainsKey($row.cardtyp)) { $cardtypMapping[$row.cardtyp] } else { $row.cardtyp }
            $resolvedStatus = if ($penStatusMapping.ContainsKey($row.status)) { $penStatusMapping[$row.status] } else { $row.status }
            $resolvedReserved = if ($reservedMapping.ContainsKey($row.reserved)) { $reservedMapping[$row.reserved] } else { $row.reserved }

            [PSCustomObject]@{
                Domain           = $row.domain
                Switch           = $row.switch_name
                PEN              = $row.pen
                LTG              = $row.ltg
                LTU              = $row.ltu
                Slot             = $row.slot
                Circuit          = $row.circuit
                PEN_Num          = $row.pen_num
                Baugruppentyp    = $resolvedCardtyp
                Board_Present    = $row.board_present
                Status           = $resolvedStatus
                Extension        = $row.extension
                Reserviert       = $resolvedReserved
                Reserviert_Zeit  = $row.reserved_time
                Beschreibung     = $row.info
                LIN_1            = $row.lin1
                LIN_2            = $row.lin2
                S0_Bus           = $row.ctr_for_busy
                HVT1_Knoten      = $row.node_1
                HVT1_Unterknoten = $row.subnode_1
                HVT1_Bucht       = $row.line_1
                HVT1_Leiste      = $row.strip_1
                HVT1_Stift       = $row.pin_1
                HVT2_Knoten      = $row.node_2
                HVT2_Unterknoten = $row.subnode_2
                HVT2_Bucht       = $row.line_2
                HVT2_Leiste      = $row.strip_2
                HVT2_Stift       = $row.pin_2
            }
        }

        $hvtResults | Export-Excel -Path $ExcelFilePath -WorkSheetname "HVT" -NoNumberConversion * -BoldTopRow -AutoSize -AutoFilter -TitleBold -TableStyle Medium5 -FreezeTopRow

        # HVT-Sheet Header formatieren
        $excelHvt = Open-ExcelPackage -Path $ExcelFilePath
        $wsHvt = $excelHvt.Workbook.Worksheets["HVT"]
        if ($wsHvt) {
            # Header-Formatierung für HVT-Verbindungsgruppen
            $lastCol = $wsHvt.Dimension.End.Column
            for ($c = 1; $c -le $lastCol; $c++) {
                $headerVal = $wsHvt.Cells[1, $c].Value
                if ($headerVal -like "HVT1_*") {
                    $wsHvt.Cells[1, $c].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $wsHvt.Cells[1, $c].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightBlue)
                } elseif ($headerVal -like "HVT2_*") {
                    $wsHvt.Cells[1, $c].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $wsHvt.Cells[1, $c].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGreen)
                }
            }
        }
        Close-ExcelPackage $excelHvt

        $hvtEnd = Get-Date
        Write-Host "  HVT-Worksheet erstellt ($($hvtResults.Count) Einträge). Dauer: $($hvtEnd - $hvtStart)" -ForegroundColor Green
        "HVT-Worksheet erstellt. $($hvtResults.Count) Einträge. Dauer: $($hvtEnd - $hvtStart)" | Out-File -FilePath $LogFilePath -Append
    } elseif ($IncludePenData) {
        Write-Host "  HINWEIS: Keine PEN-Einträge — HVT-Sheet wird mit Header erstellt" -ForegroundColor Yellow
        "HINWEIS: Keine PEN-Einträge — HVT-Sheet mit leerem Header erstellt" | Out-File -FilePath $LogFilePath -Append
        # Leeres Sheet mit Header-Zeile erstellen
        $emptyHvt = @([PSCustomObject]@{
            Domain=""; Switch=""; PEN=""; LTG=""; LTU=""; Slot=""; Circuit=""; PEN_Num=""; Baugruppentyp=""; Board_Present=""
            Status=""; Extension=""; Reserviert=""; Reserviert_Zeit=""; Beschreibung=""; LIN_1=""; LIN_2=""; S0_Bus=""
            HVT1_Knoten=""; HVT1_Unterknoten=""; HVT1_Bucht=""; HVT1_Leiste=""; HVT1_Stift=""
            HVT2_Knoten=""; HVT2_Unterknoten=""; HVT2_Bucht=""; HVT2_Leiste=""; HVT2_Stift=""
        })
        $emptyHvt | Export-Excel -Path $ExcelFilePath -WorkSheetname "HVT" -NoNumberConversion * -BoldTopRow -AutoSize -AutoFilter -FreezeTopRow
        # Leere Datenzeile entfernen — nur Header behalten
        $excelHvtEmpty = Open-ExcelPackage -Path $ExcelFilePath
        $wsHvtEmpty = $excelHvtEmpty.Workbook.Worksheets["HVT"]
        if ($wsHvtEmpty -and $wsHvtEmpty.Dimension.Rows -gt 1) { $wsHvtEmpty.DeleteRow(2) }
        Close-ExcelPackage $excelHvtEmpty
    }

    Write-Host "Verarbeitung abgeschlossen." -ForegroundColor Green
    "Verarbeitung abgeschlossen." | Out-File -FilePath $LogFilePath -Append

    $worksheetList = "Lizenz Dashboard, Gesamt, Sammelanschluss, AUN-Gruppen, NumberingPlan, Devconst, Standorte"
    if ($IncludePenData) { $worksheetList += ", HVT" }
    Write-Host "Worksheets: $worksheetList" -ForegroundColor Green
}

Write-Host "`n════════════════════════════════════════════════════════════════" -ForegroundColor Magenta
Write-Host "         PHASE 4: ERMITTLUNG DOMAIN/SWITCH-PAARE" -ForegroundColor Magenta
Write-Host "════════════════════════════════════════════════════════════════`n" -ForegroundColor Magenta

# --- Standorte automatisch aus SWITCH-Tabelle ermitteln ---
$domainSwitchPairs = @($tableSwitch | Select-Object @{N='Domain';E={$_.domain}}, @{N='SwitchName';E={$_.switch_name}} | Sort-Object Domain, SwitchName)
Write-Host "Ermittelte Domain/Switch-Paare aus SWITCH-Tabelle: $($domainSwitchPairs.Count)" -ForegroundColor Green
foreach ($p in $domainSwitchPairs) { Write-Host "  $($p.Domain)-$($p.SwitchName)" -ForegroundColor Gray }
"Ermittelte Domain/Switch-Paare aus SWITCH: $($domainSwitchPairs.Count)" | Out-File -FilePath $LogFilePath -Append

# --- Skript ausführen ---
Write-Host "`n════════════════════════════════════════════════════════════════" -ForegroundColor Yellow
Write-Host "    PHASE 5: DATENVERARBEITUNG & EXPORT (pro Standort)" -ForegroundColor Yellow
Write-Host "════════════════════════════════════════════════════════════════`n" -ForegroundColor Yellow

# Aufruf der Funktion mit den einmalig eingelesenen Tabellen und dem zentralen Excel-Pfad
Write-Host "Starte Verarbeitung der Domain/Switch-Paare..." -ForegroundColor Cyan
"Starte Verarbeitung der Domain/Switch-Paare..." | Out-File -FilePath $LogFilePath -Append
$processStartTime = Get-Date
Process-PortData -DomainSwitchPairs $domainSwitchPairs -NumberingPlanTable $table2 -CfwTable $table3 -HuntgrpTable $table4 -HuntgrpServiceTable $table5 -PickupGrpTable $tablePickupGrp -PickupSubTable $tablePickupSub -PersportTable $table6 -DevconstTable $table7 -PenDataTable $tablePenData -ExcelFilePath $ExcelFilePath -OutputPath $OutputPath -LogFilePath $LogFilePath -ApiPath $ApiPath -ApiHost $ApiHost -ApiPassword $ApiPasswordKlartext -ApiUser $ApiUser -aktuellesDatum $aktuellesDatum -MarkDuplicate:$MarkDuplicate -IncludePenData:$IncludePenData -DiagnosticMode:$Debug

$processEndTime = Get-Date

Write-Host "`n════════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "                      ✅ ABGESCHLOSSEN" -ForegroundColor Green
Write-Host "════════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "Verarbeitungszeit (Phase 5):  $($processEndTime - $processStartTime)" -ForegroundColor Green
Write-Host "Gesamtlaufzeit:                $($processEndTime - $scriptStartTime)" -ForegroundColor Green
Write-Host "Excel-Datei:                   $ExcelFilePath" -ForegroundColor Green
Write-Host "Log-Datei:                     $LogFilePath" -ForegroundColor Green
Write-Host "════════════════════════════════════════════════════════════════`n" -ForegroundColor Green

"Gesamtverarbeitung abgeschlossen. Gesamtdauer: $($processEndTime - $scriptStartTime)" | Out-File -FilePath $LogFilePath -Append
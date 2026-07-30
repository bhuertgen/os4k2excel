# API-Parameter und Funktionen

Dokumentation aller verfügbaren Parameter und neuer Funktionen (Version M26+).

## os4k2excel.ps1 - Haupt-Export-Script

### Pflichtparameter

| Parameter | Typ | Beschreibung |
|---|---|---|
| `-ApiHost` | String | IP-Adresse oder Hostname des OpenScape 4000 Servers |
| `-ApiUser` | String | Benutzername für die XIE-API (z.B. `engr`, `xieapi`) |
| `-ApiPassword` | String | Passwort für den API-Benutzer |

### Optionale Parameter

| Parameter | Typ | Standard | Beschreibung |
|---|---|---|---|
| `-ApiPath` | String | `C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe` | Pfad zur `api2hipath.exe` Executable |
| `-OutputPath` | String | Skriptverzeichnis | Zielverzeichnis für Excel, Log und CSV-Dateien |
| `-MarkDuplicate` | Switch | - | Doppelte PENs im "Gesamt"-Tab mit orange Hintergrund markieren |
| `-ShowSecrets` | Switch | - | SIP-Secrets im Klartext in Excel exportieren (Standard: maskiert als `***`) |
| `-Debug` | Switch | - | **NEU (M26):** Detaillierte Diagnose-Ausgaben in das Log schreiben |

### Aufruf-Beispiele

```powershell
# Minimal mit Pflichtparametern
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "geheim"

# Mit Debug-Modus für Diagnose
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "geheim" -Debug

# Mit allen Optionen
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "geheim" `
  -OutputPath "C:\Exports" -MarkDuplicate -ShowSecrets -Debug
```

## Neue Funktionen (M26 - 2026-03-10)

### 1. PICKUPGRP und PICKUP_SUB Tabellen

Das Script extrahiert jetzt automatisch zwei zusätzliche Tabellen:

**PICKUPGRP** - Rufannahmegruppen
- Definition von Gruppen, die Anrufe annehmen können
- Felder: `domain`, `switch_name`, `pickupgrpnum`, `displ`, `info`

**PICKUP_SUB** - Mitglieder der Rufannahmegruppen
- Zuordnung von Erweiterungen zu Pickupgruppen
- Felder: `domain`, `switch_name`, `extension`, `pickupgrpnum`

Diese Daten sind intern verfügbar für Verknüpfungen und zukünftige Ausgabe-Sheets.

### 2. Verbesserte Logging-Struktur (M26)

Das Script gibt jetzt strukturierte Ausgabe mit Phasen-Headern aus:

```
════════════════════════════════════════════════════════════════
                    PHASE 1: INITIALISIERUNG
════════════════════════════════════════════════════════════════

════════════════════════════════════════════════════════════════
                  PHASE 2: API-ABFRAGEN (Batch)
════════════════════════════════════════════════════════════════
```

Dies erleichtert das Verfolgen des Fortschritts und Debugging.

### 3. Debug-Modus (-Debug Parameter)

Mit dem neuen `-Debug` Parameter werden detaillierte diagnostische Informationen ins Log geschrieben:

- **CSV-Datei-Größe**: Bytes und Zeilenzahl jeder API-Abfrage
- **Hashtable-Größen**: Anzahl der Einträge beim Laden der Tabellen
- **API Exit-Codes**: Fehler-Codes der `api2hipath.exe` Aufrufe
- **Datenbank-Zustand**: Anzahl der Einträge pro Tabelle

**Aktivierung:**
```powershell
.\os4k2excel.ps1 -ApiHost "..." -ApiUser "..." -ApiPassword "..." -Debug
```

**Nutzen:**
- Behebung von Datenextraktionsproblemen
- Überprüfung, ob alle erwarteten Daten abgerufen wurden
- Performance-Analyse und Debugging

## Extrahierte Tabellen (komplett)

Das Script extrahiert automatisch folgende Tabellen aus der OpenScape 4000:

| Tabelle | Beschreibung | Anzahl Einträge |
|---|---|---|
| **DOMAIN** | Verfügbare Domains | variabel |
| **SWITCH** | Switches mit Domain-Zuordnung | variabel |
| **NUMBERING_PLAN** | Rufnummernplan | variabel |
| **CFW** | Rufumleitungen (Call Forwarding) | variabel |
| **HUNTGRP** | Sammelanschluss-Gruppen | variabel |
| **HUNTGRP_SERVICE** | Sammelanschluss-Dienste | variabel |
| **PICKUPGRP** | Rufannahmegruppen | variabel |
| **PICKUP_SUB** | Mitglieder der Pickupgruppen | variabel |
| **PERSPORT** | Port-Verbindungstypen | variabel |
| **DEVCONST** | Gerätekonfigurationen | variabel |
| **PORT** | Port-Daten (pro Domain/Switch) | variabel |

## Lizenzberechnung

Die Lizenzberechnung erfolgt automatisch für jeden Port:

### Basis-Wert je Port
- PEN leer → `null` (keine Lizenz)
- Gerät = RADIO oder EXTLINE → `0`
- Gerät = BASEST → `4`
- Gerät = SET600 → `2`
- Sonstige → `1`

### Flex vs. TDM Aufteilung
- **IP2-Verbindung** (Code 268) → `Flex_Lizenz`
- **Alle anderen** → `TDM_Lizenz`

Nur 1x pro Nebenstelle gezählt.

## Ausgabe-Dateien

Alle Dateien werden mit Datumsstempel (`YYYY-MM-DD`) erstellt:

- `OS4K-PORT-{date}.xlsx` - Hauptreport mit mehreren Sheets
- `OS4K-PORT-{date}.log` - Ausführungsprotokoll
- `OS4K-{TABLE}-{date}.csv` - Zwischendateien (Rohexporte)

Mit `-Debug` enthalten die Logs detaillierte diagnostische Informationen.

## Webserver Parameter (os4k2excel-server.ps1)

Der Webserver verwendet die gleichen API-Parameter wie das Hauptscript:

```powershell
.\os4k2excel-server.ps1 `
  -ApiHost "192.0.2.10" `
  -ApiUser "engr" `
  -ApiPassword "geheim" `
  -WebPassword "team2024" `
  [-SmtpServer "mail.firma.de"] `
  [-SmtpFrom "os4k@firma.de"] `
  [-SmtpTo "admin@firma.de"]
```

Der Webserver startet das Haupt-Script mit allen Extraktionen und nutzt die neuen Funktionen automatisch.

---

**Version:** M26.20260310.1239 | **Release:** v2.26.0 | **Update:** 2026-03-10

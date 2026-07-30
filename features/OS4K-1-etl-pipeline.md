# OS4K-1: ETL-Pipeline (Extract, Transform, Excel-Export)

## Status: Deployed
**Created:** 2026-03-10
**Last Updated:** 2026-03-10

## Dependencies
- None (Kernfeature)

## User Stories

- Als IT-Administrator möchte ich per PowerShell-Befehl alle Port- und Lizenzdaten meiner OS4K-Anlage exportieren, damit ich einen vollständigen Überblick ohne manuelle Klickarbeit habe.
- Als externer Techniker möchte ich bei einem Vor-Ort-Einsatz schnell eine Excel-Datei mit allen Standorten, Ports und Lizenzen erzeugen, damit ich die Systemkonfiguration dokumentieren und auswerten kann.
- Als Administrator möchte ich, dass Standorte automatisch aus der SWITCH-Tabelle erkannt werden, damit ich keine manuelle Standortliste pflegen muss.
- Als Telefonverantwortlicher möchte ich die Aufteilung zwischen Flex- und TDM-Lizenzen sehen, damit ich die korrekte Lizenzverrechnung nachweisen kann.
- Als Datenschutzbeauftragter möchte ich, dass SIP-Secrets standardmäßig maskiert (`***`) exportiert werden, damit keine Zugangsdaten unbeabsichtigt in Reports landen.

## Acceptance Criteria

- [ ] Pflichtparameter `-ApiHost`, `-ApiUser`, `-ApiPassword` werden geprüft; bei Fehlen erscheint eine verständliche Fehlermeldung mit Verwendungsbeispielen
- [ ] Folgende Tabellen werden via `api2hipath.exe` abgerufen: DOMAIN, SWITCH, NUMBERING_PLAN, CFW, HUNTGRP, HUNTGRP_SERVICE, PICKUPGRP, PICKUP_SUB, PERSPORT, DEVCONST
- [ ] Domain/Switch-Paare werden automatisch aus der SWITCH-Tabelle ermittelt — keine Hardcodierung
- [ ] Pro Standort wird die PORT-Tabelle abgerufen und transformiert
- [ ] Lizenzwerte werden korrekt berechnet: BASEST=4, SET600=2, RADIO/EXTLINE=0, sonst=1, leer=null
- [ ] Lizenzen werden nach Verbindungstyp aufgeteilt: IP2 → Flex_Lizenz, alle anderen → TDM_Lizenz
- [ ] Excel-Datei enthält Arbeitsblätter: `{Domain}-{Switch}` pro Standort, Gesamt, Lizenz Dashboard, Sammelanschluss, NumberingPlan, Devconst
- [ ] `-MarkDuplicate` markiert doppelte PENs im Gesamt-Tab orange
- [ ] `-ShowSecrets` exportiert SIP-Secrets im Klartext; ohne den Parameter werden sie als `***` maskiert
- [ ] `-Debug` schreibt detaillierte Diagnose-Ausgaben ins Log
- [ ] Alle Ausgabedateien erhalten einen Datumsstempel (`YYYY-MM-DD`)
- [ ] Log-Datei wird parallel zur Excel-Datei geschrieben

## Edge Cases

- **API nicht erreichbar:** `api2hipath.exe` schlägt fehl → Fehlermeldung ins Log, Script bricht mit Exit-Code ab
- **Leere SWITCH-Tabelle:** Keine Standorte gefunden → Warnung, kein Excel wird erstellt
- **Standort ohne Ports:** PORT-Tabelle für einen Standort ist leer → leeres Arbeitsblatt oder Standort wird übersprungen
- **Doppelte PENs:** Gleiche Nebenstelle in mehreren Zeilen → `-MarkDuplicate` hebt sie hervor, keine Fehlermeldung
- **Ungültiges Ausgabeverzeichnis:** Pfad existiert nicht → wird automatisch erstellt
- **Excel-Datei bereits offen:** Schreibfehler → Fehlermeldung mit Hinweis zum Schließen der Datei
- **Sonderzeichen in Displaynames:** Umlaute und Sonderzeichen → korrekte UTF-8-Kodierung sichergestellt

## Technical Requirements

- **Plattform:** Windows 10/11/Server 2016+, PowerShell 5.1+, .NET 4.5+
- **Abhängigkeiten:** `ImportExcel`-Modul, `api2hipath.exe` (Unify OpenScape 4000 Assistant/Manager V11)
- **Dateiformat:** UTF-8 mit BOM
- **Performance:** Laufzeit unter 5 Minuten für typische Installationen (< 20 Standorte)
- **Logging:** Alle wichtigen Schritte werden zeitgestempelt ins Log geschrieben

---
<!-- Sections below are added by subsequent skills -->

## Tech Design (Solution Architect)
_Nicht zutreffend — Single-file PowerShell-Script ohne separates Architektur-Design_

## QA Test Results
**Tested:** 2026-03-17 | **Version:** M27.20260317.1500 | **Method:** Static code analysis + Security Audit

### Acceptance Criteria: 12/12 PASSED

### Edge Cases: 4/7 passed, 2 failed, 1 partial
- **FAIL EC-1:** Kein Script-Abbruch bei API-Fehler — Script läuft mit leeren Tabellen weiter
- **FAIL EC-2:** Keine Warnung bei leerer SWITCH-Tabelle — erzeugt unvollständiges Excel
- **PARTIAL EC-6:** Initiales File-Delete hat try/catch, aber Export-Excel/Open-ExcelPackage ohne Error-Handling

### Bugs

| Bug | Severity | Status | Beschreibung |
|-----|----------|--------|-------------|
| BUG-5 | Critical | **AUSNAHME (By Design)** | API-Passwort in Prozessliste sichtbar. `$ApiPassword` wird als Klartext-CLI-Argument an `api2hipath.exe` übergeben. **Begründung:** `api2hipath.exe` unterstützt ausschließlich `-p` als CLI-Parameter — keine alternative Übergabemethode verfügbar. Prozesse laufen nur wenige Sekunden. Risiko akzeptiert. |
| BUG-6 | High | **AUSNAHME (Gewünscht)** | SIP-Secrets in CSV-Dateien im Klartext. Die PORT-API-Abfrage enthält `sip_secret`, CSV-Zwischendateien speichern diese unverschlüsselt. **Begründung:** Die SIP-Anmeldedaten werden bewusst exportiert, da sie für die Konfiguration von Applikationen benötigt werden. Der `-ShowSecrets`-Parameter steuert lediglich die Excel-Ausgabe. |
| BUG-2 | Medium | Offen | Kein Script-Abbruch bei API-Fehler — Script loggt Warnung, läuft aber mit leeren Daten weiter. |
| BUG-3 | Medium | Offen | Keine Warnung bei leerer SWITCH-Tabelle — erzeugt Excel nur mit statischen Sheets. |
| BUG-4 | Medium | Offen | Kein try/catch um `Export-Excel` / `Open-ExcelPackage` / `Close-ExcelPackage`. |
| BUG-8 | Medium | Offen | Kein Timeout auf `$process.WaitForExit()` — Script hängt bei api2hipath.exe-Hänger. |
| BUG-7 | Low | Offen | Keine Sanitisierung von `$switchName` in WHERE-Clause (theoretisches Injection-Risiko). |

### Production Ready: JA (mit Ausnahmen)
- BUG-5 und BUG-6 sind dokumentierte Ausnahmen (by design / gewünscht)
- BUG-2/3/4/8 sollten im nächsten Sprint für Robustheit behoben werden

## Deployment
_Lokale Ausführung; kein separates Deployment notwendig_

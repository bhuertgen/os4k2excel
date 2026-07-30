# OS4K-2: Webserver (Browser-Oberfläche, Zeitplaner, E-Mail)

## Status: Deployed
**Created:** 2026-03-10
**Last Updated:** 2026-03-10

## Dependencies
- Requires: OS4K-1 (ETL-Pipeline) — Webserver startet `os4k2excel.ps1` als Background-Job

## User Stories

- Als Administrator möchte ich den OS4K-Export über einen Browser starten, damit ich kein PowerShell-Fenster öffnen muss.
- Als Administrator möchte ich den Fortschritt des Exports live im Browser verfolgen, damit ich nicht im Dunkeln tappen muss, ob das Script noch läuft.
- Als Administrator möchte ich die fertige Excel-Datei direkt im Browser herunterladen, damit kein manuelles Kopieren vom Server nötig ist.
- Als Administrator möchte ich den Export per E-Mail versenden lassen, damit Kollegen die Datei automatisch erhalten.
- Als Administrator möchte ich einen Zeitplaner einrichten (Wochentag + Uhrzeit), damit der Export automatisch und regelmäßig läuft — auch ohne manuelle Auslösung.
- Als Datenschutzbeauftragter möchte ich, dass der Webserver passwortgeschützt ist, damit nicht jeder im Netzwerk den Export starten kann.

## Acceptance Criteria

- [ ] Webserver wird per `os4k2excel-server.ps1` gestartet mit Pflichtparametern: `-ApiHost`, `-ApiUser`, `-ApiPassword`, `-WebPassword`
- [ ] Login-Seite mit Passwortschutz erscheint; nach erfolgreichem Login wird ein Session-Cookie gesetzt (60 Min. Timeout)
- [ ] Dashboard zeigt Verbindungsinfo (OS4K-Host), Start-Button und Statusanzeige
- [ ] Nur ein Export gleichzeitig möglich; zweiter Start wird verhindert, solange ein Job läuft
- [ ] Fortschrittsanzeige pollt alle 2 Sekunden und zeigt Live-Status aus dem Log
- [ ] Nach Abschluss: Download-Button für Excel-Datei erscheint
- [ ] Optionaler E-Mail-Versand: Excel als Anhang per SMTP (Parameter: `-SmtpServer`, `-SmtpPort`, `-SmtpFrom`, `-SmtpTo`)
- [ ] Zeitplaner: Wochentag und Uhrzeit konfigurierbar, optionaler automatischer E-Mail-Versand
- [ ] History: Übersicht der letzten Ausführungen (manuell/Zeitplaner) mit Zeitstempel und Status
- [ ] "Server beenden"-Button im Dashboard sowie Ctrl+C und Stop-Datei (`os4k2excel-server.stop`) als alternative Beendigungswege
- [ ] HTTP-Port konfigurierbar (Standard: 8080); Hinweis auf `netsh`-Freigabe für Nicht-Admin-Betrieb

## Edge Cases

- **Port bereits belegt:** Webserver kann nicht starten → Fehlermeldung mit Port-Nummer
- **ETL-Job schlägt fehl:** Fehlermeldung im Dashboard sichtbar; Download-Button erscheint nicht
- **SMTP-Fehler:** E-Mail kann nicht gesendet werden → Fehlermeldung im Dashboard; Excel bleibt zum Download verfügbar
- **Session abgelaufen:** Nutzer wird nach 60 Minuten ohne Aktivität zur Login-Seite weitergeleitet
- **Stop-Datei während laufendem Export:** Server beendet sich erst nach Abschluss des aktuellen Jobs oder bricht ab (je nach Implementierung)
- **Keine Netzwerkfreigabe (Nicht-Admin):** Webserver kann Port nicht binden → Hinweis auf `netsh http add urlacl`

## Technical Requirements

- **Plattform:** Windows 10/11/Server 2016+, PowerShell 5.1+
- **HTTP-Server:** PowerShell `System.Net.HttpListener` (kein IIS, kein externer Webserver)
- **Auth:** Single-Password per Session-Cookie (kein Multi-User)
- **Port:** Standard 8080, konfigurierbar per Parameter
- **Abhängigkeiten:** Verwendet `os4k2excel.ps1` (OS4K-1) als externen Prozess/Background-Job

---
<!-- Sections below are added by subsequent skills -->

## Tech Design (Solution Architect)
_Nicht zutreffend — Single-file PowerShell-Script ohne separates Architektur-Design_

## QA Test Results
_To be added by /qa_

## Deployment
_Lokale Ausführung; kein separates Deployment notwendig_

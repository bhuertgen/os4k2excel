# Webserver

Der os4k2excel Webserver (`os4k2excel-server.ps1`) stellt eine browserbasierte Oberflaeche bereit, ueber die der OpenScape 4000 Export gestartet, ueberwacht und heruntergeladen werden kann. Ein integrierter Zeitplaner ermoeglicht automatische Ausfuehrung.

## Architektur

```
Browser (Port 8080)
  |
  +-- GET  /           --> Login-Seite (Passwort-Eingabe)
  +-- POST /login      --> Passwort pruefen, Session-Cookie setzen
  +-- GET  /dashboard  --> Hauptseite (Start-Button, Status, Downloads, Zeitplaner)
  +-- POST /run        --> Script starten (Background-Job)
  +-- GET  /status     --> Job-Status als JSON (Polling alle 2s)
  +-- GET  /download   --> Excel-Datei herunterladen
  +-- POST /email      --> Excel per SMTP versenden
  +-- GET  /schedule   --> Zeitplaner-Konfiguration lesen (JSON)
  +-- POST /schedule   --> Zeitplaner-Konfiguration speichern
  +-- POST /shutdown   --> Server beenden
  +-- GET  /logout     --> Session beenden
```

**Technologie:** Reines PowerShell mit `System.Net.HttpListener`. Kein IIS, kein Node.js, keine externen Abhaengigkeiten. HTML/CSS/JS ist direkt im Script eingebettet.

## Installation

### Voraussetzungen

- Gleiche Voraussetzungen wie fuer `os4k2excel.ps1` (PowerShell 5.1+, ImportExcel-Modul, api2hipath.exe)
- Netzwerkzugang zum OpenScape 4000 Server
- Port 8080 (oder konfiguriert) muss frei sein

### Port-Freigabe

Fuer nicht-Administrator-Benutzer muss einmalig als Admin der Port freigegeben werden:

```powershell
# Port freigeben (einmalig als Administrator)
netsh http add urlacl url=http://+:8080/ user=DOMAIN\USERNAME

# Freigabe wieder entfernen
netsh http delete urlacl url=http://+:8080/
```

### Firewall (optional)

Falls der Server von anderen Rechnern erreichbar sein soll:

```powershell
# Windows-Firewall-Regel hinzufuegen (als Administrator)
New-NetFirewallRule -DisplayName "os4k2excel Webserver" -Direction Inbound -Protocol TCP -LocalPort 8080 -Action Allow
```

## Aufruf

```powershell
# Minimal
.\os4k2excel-server.ps1 -ApiHost "192.0.2.10" -ApiUser "xieapi" -ApiPassword "geheim" -WebPassword "team2024"

# Mit Email-Versand und Standard-Empfaenger
.\os4k2excel-server.ps1 -ApiHost "192.0.2.10" -ApiUser "xieapi" -ApiPassword "geheim" -WebPassword "team2024" -SmtpServer "mail.firma.de" -SmtpFrom "os4k@firma.de" -SmtpTo "admin@firma.de"

# Anderer Port
.\os4k2excel-server.ps1 -ApiHost "192.0.2.10" -ApiUser "xieapi" -ApiPassword "geheim" -WebPassword "team2024" -Port 9090
```

## Parameter

| Parameter | Pflicht | Standard | Beschreibung |
|---|---|---|---|
| `-ApiHost` | Ja | - | IP-Adresse des OpenScape 4000 |
| `-ApiUser` | Ja | - | API-Benutzername |
| `-ApiPassword` | Ja | - | API-Passwort |
| `-WebPassword` | Ja | - | Passwort fuer die Web-Anmeldung |
| `-ApiPath` | Nein | `C:\Program Files (x86)\...\api2hipath.exe` | Pfad zur API-Executable |
| `-OutputPath` | Nein | Skriptverzeichnis | Zielverzeichnis fuer Ausgabedateien |
| `-Port` | Nein | `8080` | HTTP-Port |
| `-SmtpServer` | Nein | - | SMTP-Server fuer Email-Versand |
| `-SmtpPort` | Nein | `25` | SMTP-Port |
| `-SmtpFrom` | Nein | - | Absender-Adresse fuer Emails |
| `-SmtpTo` | Nein | - | Standard-Empfaenger (wird im Dashboard vorausgefuellt) |

## Funktionen

### Login

- Einfache Passwort-Authentifizierung ueber Webformular
- Session-Cookie (GUID) wird bei erfolgreichem Login gesetzt
- Session-Timeout nach 60 Minuten Inaktivitaet
- Automatische Verlaengerung bei Aktivitaet

### Dashboard

- Zeigt API-Verbindungsinfo (Host, User - kein Passwort)
- "Jetzt starten" Button zum Ausfuehren des Exports
- Fortschrittsanzeige mit Ladebalken und aktueller Meldung
- Download-Button nach Abschluss
- Email-Versand Formular (wenn SMTP konfiguriert)
- Zeitplaner-Konfiguration
- History der letzten Ausfuehrungen (mit Quelle: Manuell/Zeitplaner)

### Export-Ausfuehrung

- Startet `os4k2excel.ps1` als PowerShell Background-Job (`Start-Job`)
- Nur ein Job gleichzeitig moeglich (Queue-Schutz)
- Fortschritts-Erkennung durch Parsen der Job-Ausgabe
- Phasen: API-Abfragen (0-40%) -> Port-Verarbeitung (40-90%) -> Formatierung (90-100%)
- **Neue Tabellen (seit M26):** PICKUPGRP und PICKUP_SUB werden automatisch extrahiert

### Download

- Stellt die neueste `OS4K-PORT-*.xlsx` Datei zum Download bereit
- Korrekter MIME-Type (`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`)

### Email-Versand

- Versendet die Excel-Datei als Attachment per SMTP
- Empfaenger-Adresse wird im Dashboard eingegeben
- Standard-Empfaenger kann per `-SmtpTo` Parameter vorgegeben werden (wird im Eingabefeld vorausgefuellt)
- Verwendet `Send-MailMessage` (PowerShell built-in)
- Nur verfuegbar wenn `-SmtpServer` und `-SmtpFrom` konfiguriert sind

### Zeitplaner

Der Zeitplaner ermoeglicht automatische Ausfuehrung des Exports zu festgelegten Zeiten.

**Konfiguration im Dashboard:**
- **Ein/Aus** - Toggle zum Aktivieren/Deaktivieren
- **Uhrzeit** - Ausfuehrungszeitpunkt (z.B. 02:00)
- **Wochentage** - An welchen Tagen ausgefuehrt werden soll (Mo-So)
- **Email an** - Optionale automatische Zustellung nach Abschluss (nur bei konfiguriertem SMTP)

**Funktionsweise:**
- Der Server prueft alle 2 Sekunden ob ein geplanter Lauf faellig ist
- Der Export wird nur einmal pro Tag zur konfigurierten Minute gestartet
- Laeuft bereits ein Export, wird kein zweiter gestartet
- Nach Abschluss wird optional eine Email mit der Excel-Datei versendet
- In der History wird die Quelle "Zeitplaner" statt "Manuell" angezeigt

**Hinweis:** Die Zeitplaner-Konfiguration wird im Arbeitsspeicher gehalten und geht beim Neustart des Servers verloren. Sie muss nach jedem Serverstart im Dashboard neu eingerichtet werden.

## Server beenden

Es gibt drei Moeglichkeiten den Server sauber zu beenden:

### 1. Ctrl+C im Terminal

Direktes Beenden im Terminal-Fenster in dem der Server laeuft.

### 2. Dashboard-Button

Im Dashboard oben rechts befindet sich der Button "Server beenden". Es erscheint eine Sicherheitsabfrage vor dem Herunterfahren.

### 3. Stop-Datei (von einem anderen Terminal)

Fuer Fernsteuerung oder Automatisierung kann aus einem beliebigen Terminal eine Stop-Datei erstellt werden:

```cmd
REM CMD
echo.> "C:\pfad\zum\skript\os4k2excel-server.stop"
```

```powershell
# PowerShell
New-Item "C:\pfad\zum\skript\os4k2excel-server.stop"
```

Der Server erkennt die Datei innerhalb von 2 Sekunden, loescht sie und faehrt herunter. Laufende Background-Jobs werden gestoppt.

## Sicherheit

- Passwort wird nur im Speicher gehalten, nie in Dateien geschrieben
- API-Passwort wird nicht im Frontend angezeigt
- Session-Cookie ist HttpOnly (kein JavaScript-Zugriff)
- Session-Timeout nach 60 Minuten Inaktivitaet
- Nur ein Script-Lauf gleichzeitig moeglich
- Server laeuft standardmaessig auf Port 8080 (kein Admin-Recht noetig)
- Shutdown-Endpoint erfordert gueltige Session

## Troubleshooting

| Problem | Loesung |
|---|---|
| Port bereits belegt | Anderen Port mit `-Port` Parameter waehlen |
| Zugriff verweigert auf Port | `netsh http add urlacl` ausfuehren (siehe Port-Freigabe) |
| Kein Zugriff von anderem Rechner | Firewall-Regel hinzufuegen (siehe Firewall) |
| Email-Versand schlaegt fehl | SMTP-Server und Port pruefen, ggf. Authentifizierung am SMTP |
| Job bleibt haengen | Server mit Ctrl+C beenden und neu starten |
| Zeitplaner loest nicht aus | Pruefen ob aktiviert, Uhrzeit und Wochentag korrekt, Server muss laufen |
| Zeitplaner-Einstellungen weg | Werden nicht persistent gespeichert, nach Neustart im Dashboard erneut konfigurieren |

---

**Version:** M26.20260310.1239 | **Release:** v2.26.0 | **Update:** 2026-03-10

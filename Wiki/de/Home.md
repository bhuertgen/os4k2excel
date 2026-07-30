# os4k2excel Wiki

Willkommen im Wiki für **os4k2excel** — Ein PowerShell ETL-Tool für OpenScape 4000 Export.

## Schnelleinstieg

os4k2excel extrahiert Port-, Lizenz- und Rufumleitungsdaten aus dem Unify/Mitel OpenScape 4000 Telefonsystem und exportiert sie als formatierter Excel-Report.

**Version:** M29.20260324.1807

### Installation

Siehe **[Installation](Installation.md)** für detaillierte Schritt-für-Schritt Anleitung:
1. Voraussetzungen pruefen
2. ImportExcel Modul installieren (Online oder Offline)
3. api2hipath.exe / Export Table Client installieren
4. API-Benutzer im OS4K Manager einrichten
5. Zugriffsrechte zuweisen (mit Screenshots)
6. Script ausfuehren

**Schnell gestartet:**
```powershell
# Interaktiv (Credentials werden abgefragt):
.\os4k2excel.ps1 -ApiHost "192.0.2.10"

# Oder vollstaendig per Parameter:
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "geheim"
```

## Hauptfunktionen

### 📊 Automatische Extraktion
- 10+ Tabellen aus OpenScape 4000 (DOMAIN, SWITCH, PORT, CFW, HUNTGRP, PICKUPGRP, etc.)
- Kein Datenbank-Zugriff notwendig, nur XIE-API über `api2hipath.exe`
- Automatische Domain/Switch-Erkennung (keine manuelle Konfiguration)

### 📈 Lizenzberechnung
- Automatische Berechnung von Flex- vs. TDM-Lizenzen
- Basis-Werte je Gerätetyp (BASEST=4, SET600=2, etc.)
- Dashborad mit Lizenzsummen pro Standort

### 🌐 Web-Interface
- **os4k2excel-server.ps1** stellt Webinterface bereit
- Start, Fortschritt, Download im Browser (Port 8080)
- Integrierter Zeitplaner für automatische Ausführung
- Email-Versand der Excel-Dateien

### 📝 Diagnose-Modus
- Neuer `-Debug` Parameter für detaillierte Log-Ausgaben
- CSV-Größen, Hashtable-Einträge, API Exit-Codes
- Perfekt zum Debugging und Performance-Analyse

## Neue Features (M29)

- ✅ **Interaktive Credential-Abfrage** — Passwort maskiert, User/Password werden abgefragt wenn nicht angegeben (OS4K-5)
- ✅ **PEN/HVT-Zuordnung** — Hauptverteiler-Dokumentation via `-IncludePenData` (OS4K-4)
- ✅ **ImportExcel-Modulpruefung** und Versionsinfo im Log
- ✅ **PtrToStringBSTR** fuer PS7-Kompatibilitaet bei SecureString-Konvertierung

## Dokumentation

- **[Installation](Installation.md)** — Voraussetzungen, ImportExcel (Online/Offline), api2hipath, API-Benutzer einrichten
- **[Aufruf und Parameter](Aufruf-und-Parameter.md)** — Skriptaufruf, Pflicht- und optionale Parameter, Ausgabedateien
- **[ETL-Pipeline](ETL-Pipeline.md)** — Extract, Transform, Load — Ablauf und Code-Mappings
- **[Lizenzberechnung](Lizenzberechnung.md)** — Flex vs. TDM Lizenzen, Basis-Werte, Dashboard
- **[API-Parameter](API-Parameter.md)** — Alle verfuegbaren API-Tabellen und Funktionen
- **[Webserver](Webserver.md)** — Konfiguration und Bedienung des Web-Interfaces
- **[Change History](Change-History.md)** — Versionshistorie und Aenderungen

## Häufig verwendete Befehle

```powershell
# Interaktiv (Credentials werden abgefragt)
.\os4k2excel.ps1 -ApiHost "192.0.2.10"

# Einfacher Export (non-interaktiv)
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "geheim"

# Mit HVT/Hauptverteiler-Daten
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "geheim" -IncludePenData

# Mit Duplikat-Markierung
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "geheim" -MarkDuplicate

# Webserver (Browser: http://localhost:8080)
.\os4k2excel-server.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "geheim" -WebPassword "team2024"
```

## Anforderungen

- **Windows:** 10, 11, Server 2016+
- **PowerShell:** 5.1+
- **Module:** ImportExcel (kostenlos)
- **Software:** api2hipath.exe (von OpenScape 4000 Manager/Assistant V11)
- **Zugriff:** XIE-API Credentials und Netzwerkzugriff zum OS4K Server

## Output

Alle Dateien werden mit Datum im Namen erstellt:

- **OS4K-PORT-2026-03-10.xlsx** — Hauptreport (mehrere Sheets pro Standort)
- **OS4K-PORT-2026-03-10.log** — Ausführungsprotokoll
- **OS4K-{TABLE}-2026-03-10.csv** — Rohexporte der einzelnen Tabellen

## Sicherheit

- ✅ Keine Passwörter in Dateien
- ✅ Interaktive Passwort-Eingabe maskiert (kein Echo auf Bildschirm)
- ✅ SIP-Secrets optional maskiert
- ✅ Session-basierte Web-Authentifizierung
- ✅ Nur lokal empfohlen oder in sicheren Netzwerken

---

**Letzte Änderung:** 2026-03-24 (Version M29)
**Dokumentation:** [README.md](../README.md)

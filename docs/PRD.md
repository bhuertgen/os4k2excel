# PRD: os4k2excel — OpenScape 4000 Port & Lizenz Export

**Version:** M26.20260310.1239
**Status:** Produktiv

---

## Vision

os4k2excel ist ein PowerShell-basiertes ETL-Tool, das Port-, Lizenz- und Rufumleitungsdaten aus einem Unify OpenScape 4000 Telefonsystem automatisch extrahiert, transformiert und als formatierte Excel-Berichte exportiert. Es löst das Problem der manuellen, fehleranfälligen Datenerfassung aus der OS4K-Oberfläche und liefert auf Knopfdruck einen vollständigen, strukturierten Überblick über alle Ports, Lizenzen und Konfigurationen.

---

## Zielnutzer

### Interne IT-/Telefonadministratoren
- Betreiben das OS4K-System im eigenen Unternehmen
- Benötigen regelmäßige Übersichten für Planung, Audits und Lizenzcontrolling
- Führen Änderungen direkt am System durch und nutzen den Report zur Verifikation

### Externe Techniker / Dienstleister
- Betreuen OS4K-Systeme bei mehreren Kunden
- Benötigen schnelle Bestandsaufnahme bei Vor-Ort-Einsätzen oder Remote-Support
- Arbeiten ohne dauerhaften Systemzugang — der Export ist ihr zentrales Arbeitsmittel

### Gemeinsame Bedürfnisse
- Kein manuelles Zusammenklicken von Daten aus der OS4K-Oberfläche
- Reproduzierbare, datumgestempelte Reports für Dokumentation und Vergleiche
- Klare Trennung von Flex- und TDM-Lizenzen für Lizenzcontrolling

---

## Core Features (Roadmap)

| ID | Feature | Priorität | Status |
|---|---|---|---|
| OS4K-1 | ETL-Pipeline (Extract, Transform, Excel-Export) | P0 (MVP) | Deployed |
| OS4K-2 | Webserver (Browser-Oberfläche, Zeitplaner, E-Mail) | P1 | Deployed |
| OS4K-3 | Delete NSt via CSV Import with Dependency Management | P2 | Planned |
| OS4K-5 | Interaktive Credential-Abfrage (Passwort maskiert) | P1 | Planned |

---

## Erfolgsmetriken

- Export läuft fehlerfrei durch alle Standorte (Domain/Switch-Paare)
- Excel-Datei enthält alle erwarteten Arbeitsblätter vollständig und korrekt formatiert
- Lizenzwerte (Flex/TDM) stimmen mit den Systemwerten überein
- Laufzeit unter 5 Minuten für typische Installationen (< 20 Standorte)
- Webserver-Export startet und liefert Download-Link ohne manuelle Eingriffe

---

## Constraints

- **Technisch:** Nur Windows (PowerShell 5.1 + .NET 4.5+); `api2hipath.exe` muss installiert sein und Netzwerkzugriff auf OS4K haben
- **Abhängigkeiten:** `ImportExcel`-Modul (kein Office erforderlich); `api2hipath.exe` ist proprietär (Unify/Mitel, nicht separat erhältlich)
- **Sicherheit:** API-Credentials werden als Parameter übergeben (keine Persistenz); SIP-Secrets standardmäßig maskiert
- **Skalierung:** Kein Parallelisierungskonzept — sequenzielle API-Abfragen pro Standort

---

## Non-Goals

- Kein Schreiben von Daten zurück ins OS4K-System (read-only)
- Keine Web-App oder Cloud-Deployment — lokale PowerShell-Ausführung
- Keine Unterstützung für andere Telefonsysteme (z.B. HiPath 3000, Cisco)
- Keine automatische Lizenzverwaltung oder -buchung
- Keine Multi-User-Authentifizierung im Webserver (Single-Password)

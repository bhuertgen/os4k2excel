# Change History

## Version M30.20260730 (Meilenstein 30)

| Datum | Änderung |
|---|---|
| 2026-07-30 | **Fix:** `-ApiPassword` ohne Wert brach mit einem PowerShell-Bindungsfehler ab, statt das Passwort abzufragen (Issue #3). Der Parameter nimmt jetzt beide Formen an. |
| 2026-07-30 | **Fix:** Umlaute in Teilnehmernamen wurden im Excel-Export zerstört (Issue #2). `api2hipath.exe` liefert Windows-1252, `Import-Csv` las die Dateien als UTF-8 — jeder Umlaut wurde zu `U+FFFD`. Betraf 255 Zellen je Lauf. |
| 2026-07-30 | **Neu:** Hilfsfunktion `Import-ApiCsv` mit Zeichensatz-Erkennung aus dem Dateiinhalt (gültiges UTF-8 bleibt UTF-8, sonst Windows-1252). Alle 11 CSV-Importe umgestellt. |
| 2026-07-30 | **Fix:** Exit-Code von `api2hipath.exe` wurde unter Windows PowerShell 5.1 nie ausgewertet (Issue #1). `-Wait` an beiden `Start-Process`-Aufrufen ergänzt. |
| 2026-07-30 | **Neu:** Wiki-Seiten [Sicherheit](Sicherheit.md) und [PowerShell 5.1 und 7.x](PowerShell-Kompatibilitaet.md) |
| 2026-07-30 | Sicherheitstest dokumentiert: Passwort in der Prozess-Kommandozeile ist ohne Administratorrechte lesbar und mit `api2hipath.exe` nicht vermeidbar |

## Version M29.20260324 (Meilenstein 29)

| Datum | Aenderung |
|---|---|
| 2026-03-24 | **NEU:** Interaktive Credential-Abfrage — Passwort maskiert, User abfragbar (OS4K-5) |
| 2026-03-24 | Nicht-interaktive Umgebungen (Scheduled Tasks) werden erkannt — sauberer Abbruch statt haengender Prompt |
| 2026-03-24 | ImportExcel-Modulpruefung beim Start mit Versionsinfo im Log |
| 2026-03-24 | Fix: PtrToStringBSTR statt PtrToStringAuto fuer PS7-Kompatibilitaet |
| 2026-03-24 | OS4K-4 und OS4K-5 als Deployed markiert (OS4K-3 bleibt In Review) |

## Version M27.20260317 (Meilenstein 27)

| Datum | Aenderung |
|---|---|
| 2026-03-17 | **NEU:** PEN/HVT-Zuordnung via `-IncludePenData` (OS4K-4) |
| 2026-03-17 | Neues HVT-Sheet mit allen PENs (belegt, frei, reserviert) und HVT-Verbindungsdaten |
| 2026-03-17 | HVT-Spalten in Standort- und Gesamt-Sheets (Info, HVT1, HVT2 je 5 Felder) |
| 2026-03-17 | Vollstaendiges Board-Typ-Mapping (80+ Typen, Doku Sektion 10.3.19-10.3.24) |
| 2026-03-17 | PEN-Tabelle statt PENDATA (korrekte ODF fuer Manager und Assistant) |
| 2026-03-17 | PEN-Tabelle liefert lin1/lin2 direkt (kein separater PORT-Lookup noetig) |
| 2026-03-17 | DOMAIN-Tabelle entfernt (wurde nicht funktional genutzt) |
| 2026-03-17 | Zusaetzliche PEN-Felder: LTG, LTU, Slot, Circuit, PEN_Num, Board_Present, Reserviert_Zeit |

## Version M26.20260310 (Meilenstein 26)

| Datum | Aenderung |
|---|---|
| 2026-03-10 | PICKUPGRP und PICKUP_SUB Tabellen-Extraktion |
| 2026-03-10 | Verbesserte Logging-Struktur mit Phasen-Header |
| 2026-03-10 | Neuer `-Debug` Parameter fuer detaillierte Diagnose-Ausgaben |
| 2026-03-10 | Verbesserte Fehlerbehandlung und API Exit-Code Ueberpruefung |
| 2026-03-10 | Security: API-Passwort in Logs maskiert |
| 2026-03-17 | Wiki: ImportExcel Online- und Offline-Installationsanleitung |
| 2026-03-17 | Wiki: ImportExcel Update- und Versionsmanagement-Anleitung |

## Version M22.20260215.1931 (Meilenstein 22)

| Datum | Aenderung |
|---|---|
| 2026-02-15 | Automatische Standort-Erkennung aus SWITCH-Tabelle (API) |
| 2026-02-15 | DOMAIN/SWITCH-Tabellen als API-Abfragen hinzugefuegt |
| 2026-02-15 | Parameter `-ApiUser` als Pflichtparameter (ersetzt `-ApiLanguage`) |
| 2026-02-15 | Parameter `-ShowSecrets` fuer optionale SIP-Secret Anzeige |
| 2026-02-15 | Parametervalidierung mit Hilfetext bei fehlenden Pflichtparametern |
| 2026-02-15 | Versionsanzeige und Ausgabepfad-Anzeige beim Skriptstart |
| 2026-02-15 | Fix: Single-Switch-Systeme (PowerShell Array-Handling) |
| 2026-02-15 | Fix: Leere Tabellen (HuntgrpTable, CfwTable) duerfen leer sein |
| 2026-02-15 | Fix: Excel-Dateisperre bei erneutem Aufruf am selben Tag |
| 2026-02-15 | Security: Hardcodierte Credentials entfernt |
| 2026-02-15 | Security: sip_secret im Export standardmaessig maskiert |
| 2026-02-15 | OutputPath Default: Skriptverzeichnis statt C:\script |
| 2026-02-15 | Kundenspezifische Referenzen entfernt, Dateiprefix auf OS4K |

## Fruehere Versionen

| Datum | Aenderung |
|---|---|
| 2025-12-05 | Trennung Flex vs. TDM Lizenzen + Dashboard Erweiterung |
| 2025-11-28 | Sammelanschluss Auto-Fill Name (Master) + Graue Schrift |
| 2025-11-28 | Fixes (Lizenz, XML-Fehler, Farben) |
| 2025-06-30 | Initiale Version |

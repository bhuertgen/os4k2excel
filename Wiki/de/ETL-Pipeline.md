# ETL-Pipeline

Das Skript folgt einer linearen ETL-Pipeline (Extract - Transform - Load) in einer einzigen Datei.

## 1. Extract (Datenextraktion)

Das Skript ruft via `api2hipath.exe` folgende Tabellen aus der OpenScape 4000 ab:

| Tabelle | Inhalt |
|---|---|
| **SWITCH** | Verfuegbare Switches mit Domain-Zuordnung |
| **NUMBERING_PLAN** | Rufnummernplan |
| **CFW** | Rufumleitungen (Call Forwarding) |
| **HUNTGRP** | Sammelanschluss-Gruppen |
| **HUNTGRP_SERVICE** | Sammelanschluss-Dienste |
| **PICKUPGRP** | Rufannahmegruppen — **NEU (M26)** |
| **PICKUP_SUB** | Mitglieder der Rufannahmegruppen — **NEU (M26)** |
| **PERSPORT** | Port-Verbindungstypen, Abteilung |
| **DEVCONST** | Geraetekonfigurationen |

Zusaetzlich wird pro Standort (Domain/Switch) die **PORT**-Tabelle abgerufen.

Optional (mit `-IncludePenData`):

| Tabelle | Inhalt |
|---|---|
| **PEN** | Physical Equipment Numbers mit HVT-Zuordnung (Hauptverteiler) — **NEU (M27)** |

## 2. Standort-Erkennung (automatisch)

Die Domain/Switch-Kombinationen werden **automatisch** aus der **SWITCH**-Tabelle ermittelt (Felder `domain` + `switch_name`). Es ist keine manuelle Konfiguration der Standorte notwendig.

## 3. Transform (Datenverarbeitung)

- Verknuepfung aller Tabellen ueber Hashtable-Lookups (Domain + Switch + Key)
- Mapping von Codes auf lesbare Werte (z.B. `841` -> `CFU`, `i90` -> `VOICE`)
- [Lizenzberechnung](Lizenzberechnung.md) mit Flex/TDM-Aufteilung

### Code-Mappings

| Kategorie | Beispiele |
|---|---|
| **Varianten** | j07=STATION, j08=SYSTEM, j09=STATIONV |
| **Services** | i90=VOICE, i91=AUDIO3K1, i92=FAXG23, j01=FAX |
| **Geraetetypen (dtype)** | 841=CFU, 842=CFB, 843=CFNR, 844=CD |
| **Interaktionstyp** | j03=EXTERN, j04=INTERN, j05=GEN |
| **Aktivierung** | i70=Ausschalten, C56=Einschalten |
| **Verbindungstyp** | 256=DIRECT, 257=PNT, 260=EXTERNAL, 263=LOG, 267=IP, 268=IP2 |
| **Baugruppen-Typ** | 300=SLMA, 306=STMD, 317=SLMAVAR, 349=SLMOP, 360=STMI-HFA, 410=STMI, 411=STMI2-HFA, ... (80+ Typen) — **NEU (M27)** |
| **PEN-Status** | 002=Frei, 003=Belegt — **NEU (M27)** |
| **Reserviert** | 000=nein, 001=ja — **NEU (M27)** |

## 4. Load (Excel-Export)

Export in eine Excel-Datei mit folgenden Arbeitsblaettern:

| Arbeitsblatt | Inhalt |
|---|---|
| **{Domain}-{Switch}** | Port-Daten pro Standort mit Lizenz-Summen |
| **Gesamt** | Alle Standorte zusammengefasst |
| **Lizenz Dashboard** | Zusammenfassung Flex- vs. TDM-Lizenzen pro Standort |
| **Sammelanschluss** | Hunting-Groups mit Auto-Fill Displaynames |
| **NumberingPlan** | Rufnummernplan |
| **Devconst** | Geraetekonfigurationen |
| **HVT** | Hauptverteiler-Zuordnung aller PENs (nur mit `-IncludePenData`) — **NEU (M27)** |

### Excel-Formatierung

| Element | Farbe |
|---|---|
| Standort-Summenzeile | Gelb |
| Gesamt-Summenzeile | Gruen |
| Dashboard-Header | Blau (weisse Schrift) |
| Dashboard-Gesamtsumme | Gold |
| Doppelte PENs (optional) | Orange |
| Auto-Fill Displaynames | Grau |
| HVT1-Header (HVT-Sheet) | Hellblau |
| HVT2-Header (HVT-Sheet) | Hellgruen |

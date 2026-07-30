# OS4K-4: PENDATA / HVT-Zuordnung (Hauptverteiler-Dokumentation)

## Status: Deployed
**Created:** 2026-03-17
**Last Updated:** 2026-03-24

## Dependencies
- Requires: OS4K-1 (ETL-Pipeline) — Erweiterung der `os4k2excel.ps1`, nutzt bestehende Extract/Transform/Load-Architektur, Join über `pen`-Feld

## Zusammenfassung

Optionale Erweiterung der ETL-Pipeline um die PENDATA-Tabelle der XIE API (aktivierbar via `-IncludePenData` Switch). Exportiert die physische Hauptverteiler-Zuordnung (HVT-Verbindung 1 und 2) zu jeder PEN (Physical Equipment Number). Besonders relevant für TDM-Anschlüsse.

**Kernnutzen:** Dokumentation der physischen Rangierungen, Enddosen und Raumzuordnungen (im `info`-Feld). Wenn TDM-Nebenstellen gelöscht werden (z.B. via OS4K-3), bleiben die physischen Rangierungen am Hauptverteiler oft bestehen. Über die gespeicherten HVT-Daten (Dosennummer, Raum im `info`-Feld) kann bei einer späteren Neueinrichtung die passende Lage wiedergefunden werden — ohne erneute Rangierarbeiten. Voraussetzung: Die Rangierungen wurden beim Löschen der NSt nicht aufgehoben.

## User Stories

- Als **TK-Techniker** möchte ich die HVT-Zuordnung (Knoten, Unterknoten, Bucht, Leiste, Stift) zu jedem Anschluss im Excel-Report sehen, damit ich bei Vor-Ort-Arbeiten am Hauptverteiler die richtige Rangierung finde.
- Als **TK-Administrator** möchte ich eine Gesamtübersicht aller PENs (belegt, frei, reserviert) mit HVT-Daten, damit ich die Kapazitätsplanung am Hauptverteiler durchführen kann.
- Als **TK-Techniker** möchte ich auf den Standort-Sheets die HVT-Verbindungsdaten direkt neben den Port-Daten sehen, damit ich nicht zwischen verschiedenen Sheets wechseln muss.
- Als **TK-Administrator** möchte ich freie und reservierte PENs mit HVT-Zuordnung sehen, damit ich weiß, welche Rangierungen noch verfügbar sind.
- Als **TK-Techniker** möchte ich bei PENs mit zwei HVT-Verbindungen (z.B. S0-Bus) beide Verbindungen sehen, damit die Dokumentation vollständig ist.
- Als **TK-Techniker** möchte ich nach dem Löschen einer TDM-NSt anhand der Dosennummer oder Raumnummer (aus dem `info`-Feld) die passende freie Lage wiederfinden, damit ich bei einer Neueinrichtung keine erneuten Rangierarbeiten durchführen muss.
- Als **TK-Administrator** möchte ich die PENDATA-Abfrage nur bei Bedarf aktivieren (`-IncludePenData`), damit der Standard-Report nicht durch zusätzliche API-Abfragen verlangsamt wird.

## Acceptance Criteria

- [ ] Neuer PowerShell Switch-Parameter `-IncludePenData` (Standard: aus) — nur wenn gesetzt, wird PENDATA abgefragt, das HVT-Sheet erstellt und HVT-Spalten in Standort-/Gesamt-Sheets hinzugefügt
- [ ] PENDATA-Tabelle wird via `api2hipath.exe` einmalig global abgefragt (SELECT) und als CSV exportiert — enthält alle Domains/Switches, Zuordnung per Hashtable-Lookup
- [ ] Alle PENDATA-Felder werden extrahiert: `domain`, `switch`, `pen`, `cardtyp`, `status`, `extension`, `reserved`, `info`, `node_1`, `subnode_1`, `line_1`, `strip_1`, `pin_1`, `node_2`, `subnode_2`, `line_2`, `strip_2`, `pin_2`, `ctr_for_busy`
- [ ] `cardtyp` wird von kodiertem Wert (z.B. 360) auf lesbaren Board-Namen aufgelöst (z.B. STMI-HFA). Mapping gemäß XIE API Doku Sektion 10.3.19-10.3.24
- [ ] `status` wird von kodiertem Wert aufgelöst: 002=Frei, 003=Belegt (gemäß General Coded Values 0-10, XIE API Doku Seite 426)
- [ ] `lin1` und `lin2` werden aus der bestehenden PORT-Abfrage mitextrahiert (keine zusätzliche API-Abfrage nötig) und nur im HVT-Sheet angezeigt (nicht in Standort-/Gesamt-Sheets)
- [ ] HVT-Spalten werden als zusätzliche Spalten in den bestehenden Standort-Sheets angezeigt (Join über `pen`-Feld aus PORT-Tabelle): Info, HVT1-Knoten, HVT1-Unterknoten, HVT1-Bucht, HVT1-Leiste, HVT1-Stift, HVT2-Knoten, HVT2-Unterknoten, HVT2-Bucht, HVT2-Leiste, HVT2-Stift
- [ ] Ein separates "HVT" Sheet wird erstellt mit allen PENs (belegt, frei, reserviert) und deren HVT-Verbindungsdaten
- [ ] Alle PENs werden exportiert: belegte (status=busy), freie (status=free) und reservierte (reserved)
- [ ] Das HVT-Sheet enthält: Domain, Switch, PEN, Baugruppen-Typ, Status, Extension, Reserviert, Info (Beschreibung), LIN 1, LIN 2, S0-Bus Teilnehmer (ctr_for_busy), HVT-Verbindung 1 (Knoten/Unterknoten/Bucht/Leiste/Stift), HVT-Verbindung 2 (Knoten/Unterknoten/Bucht/Leiste/Stift)
- [ ] Das "Gesamt"-Sheet enthält ebenfalls die HVT-Spalten (aus dem Join)
- [ ] Excel-Formatierung: HVT-Spalten erhalten passende Header und Spaltenbreiten
- [ ] CSV-Zwischendatei `OS4K-PENDATA-{date}.csv` wird erzeugt

## Edge Cases

- **PEN ohne HVT-Daten:** Auch IP-Anschlüsse (IP2/IP) können relevante Daten haben — z.B. CAT5-Dosennummer im `info`-Feld. Daher werden alle PENs gleich behandelt, unabhängig vom Anschlusstyp. HVT-Felder (Knoten/Bucht/Leiste/Stift) bleiben leer wenn nicht gepflegt, das `info`-Feld wird immer angezeigt.
- **Mehrfach-Zuordnung (HVT-Verbindung 1 + 2):** PENs mit zwei HVT-Verbindungen (z.B. bei S0-Bus mit `ctr_for_busy > 0`) zeigen beide Verbindungen vollständig an — separate Spaltengruppen für Verbindung 1 und 2.
- **Reservierte PENs:** PENs mit `reserved=001` (ja) aber ohne Extension werden im HVT-Sheet mit Status "Reserviert" gekennzeichnet. Diese Lagen sind für geplante Anschlüsse vorgemerkt. In den Standort-Sheets tauchen sie nicht auf (kein PORT-Eintrag).
- **Parameter nicht gesetzt:** Ohne `-IncludePenData` werden keine PENDATA-Abfragen durchgeführt, kein HVT-Sheet erstellt und keine HVT-Spalten in den Standort-Sheets angezeigt. Der Report verhält sich identisch zum bisherigen Verhalten.
- **PEN in PORT aber nicht in PENDATA:** Falls ein PORT eine PEN referenziert, die in PENDATA nicht existiert (Dateninkonsistenz), werden die HVT-Spalten im Standort-Sheet leer gelassen. Im Log wird eine Warnung ausgegeben.
- **PENDATA-Abfrage schlägt fehl:** Wenn die PENDATA-API-Abfrage fehlschlägt (z.B. Berechtigungsproblem), wird eine Warnung ins Log geschrieben. Der restliche Report wird ohne HVT-Daten erzeugt — kein Abbruch.
- **Leere PENDATA-Tabelle:** Falls keine PENDATA-Einträge vorhanden sind, wird das HVT-Sheet mit Header aber ohne Datenzeilen erstellt. Log-Hinweis.

## Technical Requirements

- **API-Tabelle:** PENDATA (XIE API Dokumentation, Seite 313-314)
- **API-Operation:** SELECT (Export only, read-only)
- **Join-Feld:** `pen` (Physical Equipment Number) — verknüpft PENDATA mit PORT-Tabelle
- **Sekundärer Join:** `domain` + `switch` + `extension` — für Zuordnung zu Standort-Sheets
- **PENDATA-Abfrage:** Einmalig global (nicht pro Domain/Switch), da die Tabelle bereits `domain` und `switch` als Felder enthält. Verknüpfung zu den Standort-Sheets erfolgt über Hashtable-Lookup nach `domain`+`switch`+`pen`
- **Keine Lizenz-Auswirkung:** HVT-Daten dienen rein der physischen Dokumentation, keine Änderung am Lizenz-Dashboard
- **Zieldatei:** `os4k2excel.ps1` — alle Änderungen in der bestehenden ETL-Pipeline
- **Neuer Parameter:** `-IncludePenData` [switch] — aktiviert PENDATA-Abfrage und HVT-Output
- **Performance:** Nur eine einzige zusätzliche API-Abfrage pro Script-Lauf (nicht pro Domain/Switch), nur wenn `-IncludePenData` gesetzt

## PENDATA API-Felder (Referenz)

| Feld | Typ | Beschreibung |
|---|---|---|
| `identifier` | int | unique_key |
| `domain` | char(8) | Domain |
| `switch` | char(4) | Switch Name |
| `pen` | char(11) | Physical Equipment Number (LTG/LTU/EBT/SATZ) |
| `cardtyp` | char(3) | Baugruppen-Typ, kodiert (z.B. 300=SLMA, 306=STMD, 360=STMI-HFA, 410=STMI, 411=STMI2-HFA). Vollständige Liste: XIE API Doku Sektion 10.3.19-10.3.24 (Seiten 433-438). Muss auf lesbaren Namen aufgelöst werden. |
| `status` | char(3) | Lage belegt? Kodiert gemäß General Coded Values (0-10): 002=Free (frei, keine NSt), 003=Used (belegt, NSt administriert) |
| `info` | char(50) | Beschreibungsfeld — typischerweise genutzt für Dosennummer, Raumnummer, Standorthinweis o.ä. Zentrales Feld für die Wiederauffindbarkeit von Lagen nach NSt-Löschung |
| `node_1` | char(6) | HVT-Verbindung 1: Knoten |
| `subnode_1` | char(2) | HVT-Verbindung 1: Unterknoten |
| `line_1` | char(3) | HVT-Verbindung 1: Bucht |
| `strip_1` | char(2) | HVT-Verbindung 1: Leiste |
| `pin_1` | char(3) | HVT-Verbindung 1: Stift |
| `node_2` | char(6) | HVT-Verbindung 2: Knoten |
| `subnode_2` | char(2) | HVT-Verbindung 2: Unterknoten |
| `line_2` | char(3) | HVT-Verbindung 2: Bucht |
| `strip_2` | char(2) | HVT-Verbindung 2: Leiste |
| `pin_2` | char(3) | HVT-Verbindung 2: Stift |
| `ctr_for_busy` | int | Anzahl zugeordneter S0-Bus Teilnehmer |
| `reserved` | char(3) | Lage reserviert? Kodiert: 000=nein, 001=ja. Entspricht der Checkbox "Reserviert" unter Lagenreservierung im Configuration Management. Eine reservierte Lage wird für einen geplanten Anschluss freigehalten, ohne dass bereits ein Gerät darauf administriert ist. |
| `extension` | char(6) | Zugeordnete Rufnummer |


---
<!-- Sections below are added by subsequent skills -->

## Tech Design (Solution Architect)

### Übersicht

OS4K-4 ist eine **reine Erweiterung der bestehenden ETL-Pipeline** in `os4k2excel.ps1`. Kein neues Script, keine neuen Abhängigkeiten — alles bleibt in der bewährten Struktur.

Das Feature ist vollständig hinter einem neuen `-IncludePenData` Switch versteckt. Ohne diesen Parameter verhält sich das Script **identisch** zum bisherigen Stand.

---

### Komponentenstruktur (Visuell)

```
os4k2excel.ps1 (erweitert)
│
├── [NEU] Parameter: -IncludePenData [switch]
│
├── PHASE 2: API-Abfragen (Batch)
│   ├── ... (bestehende Tabellen unverändert)
│   └── [NEU, nur wenn -IncludePenData] PENDATA-Abfrage → OS4K-PENDATA-{date}.csv
│
├── PHASE 3: CSV-Daten laden
│   ├── ... (bestehende Tabellen unverändert)
│   └── [NEU, nur wenn -IncludePenData] $tablePenData laden
│
├── [NEU, nur wenn -IncludePenData] PENDATA-Hashtable aufbauen
│   └── Schlüssel: "domain|switch|pen" → HVT-Datenzeile (O(1)-Lookup)
│
├── Process-PortData (Funktion, erweitert)
│   ├── Bestehender Port-Loop (unverändert)
│   └── [NEU, nur wenn -IncludePenData] 6 HVT-Spalten per Hashtable-Lookup anfügen
│       └── Beschreibung, HVT1 (Knoten/Unterknoten/Bucht/Leiste/Stift)
│
├── Excel-Export: Standort-Sheets (erweitert)
│   └── [NEU, nur wenn -IncludePenData] 6 zusätzliche HVT-Spalten rechts anhängen
│       └── Beschreibung, HVT-Knoten, HVT-Unterknoten, HVT-Bucht, HVT-Leiste, HVT-Stift
│
├── Excel-Export: "Gesamt"-Sheet (erweitert)
│   └── [NEU, nur wenn -IncludePenData] gleiche 6 HVT-Spalten
│
└── [NEU] Excel-Sheet: "HVT" (vollständige Dokumentation)
    └── Alle PENDATA-Zeilen (belegt, frei, reserviert) mit ALLEN Feldern:
        Baugruppentyp, Status, Reserviert, Beschreibung, LIN 1, LIN 2,
        HVT-Verbindung 1 (5 Felder), HVT-Verbindung 2 (5 Felder), ctr_for_busy
```

---

### Datenmodell (Plain Language)

**Neue Datenquelle: PENDATA-Tabelle**
- Enthält alle physischen Anschlusspunkte (PENs) der Anlage
- Jede Zeile = eine physische Lage (Port auf einer Baugruppe)
- Felder: Anlage, Baugruppen-Typ (kodiert), Status (frei/belegt), Extension, Info-Text, zwei HVT-Verbindungen (je 5 Felder: Knoten, Unterknoten, Bucht, Leiste, Stift), S0-Bus-Zähler, Reservierungs-Flag

**Join-Logik:**
- PORT-Tabelle enthält `pen`-Feld → Schlüssel für Lookup in PENDATA
- Zusammengesetzter Schlüssel: `domain + switch + pen` (eindeutig pro Anlage)
- Kein eigenes `pen`-Feld in den Standort-Sheets nötig — wird bereits exportiert

**Neue Code-Tabellen (Mappings):**
- `cardtyp` → lesbare Board-Namen (z.B. 300→SLMA, 306→STMD, 360→STMI-HFA, 410→STMI, 411→STMI2-HFA + alle weiteren aus Doku Sektion 10.3.19-10.3.24)
- `status` → 002=Frei, 003=Belegt
- `reserved` → 000=nein, 001=ja

**Wo gespeichert:** Nur in RAM während des Script-Laufs + CSV-Zwischendatei + Excel-Output (kein persistenter Zustand)

---

### Ablauf im Detail

```
Schritt 1: Parameter-Check
  └── Wenn -IncludePenData gesetzt → PENDATA in $tables-Array einfügen
      (nur in diesem Zweig, keine Änderung am Standard-Ablauf)

Schritt 2: API-Abfragen (Phase 2)
  └── PENDATA wird wie alle anderen Tabellen abgefragt:
      api2hipath.exe -o PENDATA -s <alle Felder> → OS4K-PENDATA-{date}.csv

Schritt 3: CSV laden (Phase 3)
  └── $tablePenData = Import-Csv "OS4K-PENDATA-{date}.csv"

Schritt 4: Hashtable aufbauen (einmalig, nach CSV-Import)
  └── $penDataLookup = @{}
      Schlüssel: "$domain|$switch|$pen"
      Wert: gesamte PENDATA-Zeile
      → Fehlerbehandlung: wenn Datei fehlt → Warning ins Log, kein Abbruch

Schritt 5: Process-PortData (Transform-Phase)
  └── Pro PORT-Zeile: $penDataLookup["$domain|$switch|$pen"] nachschlagen
      → 6 HVT-Felder der Ausgabe-Zeile anfügen (Beschreibung + HVT1)
      → Wenn kein Treffer: HVT-Spalten leer lassen + Warning ins Log

Schritt 6: Standort-Sheets + Gesamt-Sheet (Load-Phase)
  └── 6 neue Spaltenüberschriften:
      Beschreibung, HVT-Knoten, HVT-Unterknoten, HVT-Bucht, HVT-Leiste, HVT-Stift
      → Nur HVT-Verbindung 1 (Verbindung 2 nur im HVT-Sheet)
      → Formatierung: passende Spaltenbreiten, fetter Header

Schritt 7: Neues "HVT"-Sheet erstellen (vollständig)
  └── Alle $tablePenData-Zeilen exportieren (ungefiltert: frei + belegt + reserviert)
      → Spalten: Domain, Switch, PEN, Baugruppentyp (aufgelöst), Status (aufgelöst),
        Extension, Reserviert, Beschreibung, LIN 1, LIN 2,
        HVT1 (Knoten/Unterknoten/Bucht/Leiste/Stift),
        HVT2 (Knoten/Unterknoten/Bucht/Leiste/Stift), S0-Bus (ctr_for_busy)
      → Formatierung: farbiger Header, Spaltenbreiten
```

---

### Technische Entscheidungen (mit Begründung)

| Entscheidung | Warum |
|---|---|
| Einzelne globale PENDATA-Abfrage (nicht pro Domain/Switch) | PENDATA enthält bereits `domain`+`switch` als Felder — eine Abfrage reicht, spart API-Calls |
| Hashtable-Lookup (nicht Array-Suche) | Bestehende Architektur nutzt überall Hashtables für O(1)-Lookup — bewährtes Muster |
| `-IncludePenData` als Switch (Standard: aus) | PENDATA-Abfrage kann bei großen Anlagen mehrere Sekunden dauern — nur bei Bedarf |
| Kein eigenes Script | Alle Daten stammen aus derselben API, Excel-Export ist bereits vorhanden — Erweiterung sinnvoller als neues Script |
| lin1/lin2 nur im HVT-Sheet | Technische Leitungskennung — für TK-Techniker relevant, aber nicht für Standort-Übersichten |
| Fehlertoleranz (kein Abbruch bei PENDATA-Fehler) | Report soll auch ohne HVT-Daten nutzbar sein — Warnung genügt |

---

### Neue Abhängigkeiten

Keine neuen Pakete erforderlich. Alle genutzten Technologien sind bereits vorhanden:
- `api2hipath.exe` (schon vorhanden)
- `ImportExcel` Modul (schon vorhanden)
- PowerShell Hashtables, Import-Csv, Export-Excel (schon genutzt)

---

### Auswirkungen auf bestehende Sheets

| Sheet | Änderung |
|---|---|
| Standort-Sheets (z.B. "MUC1") | +6 HVT-Spalten rechts (Beschreibung + HVT1) — nur wenn `-IncludePenData` |
| Gesamt | +6 HVT-Spalten (gleich) — nur wenn `-IncludePenData` |
| Lizenz Dashboard | Keine Änderung |
| Sammelanschluss | Keine Änderung |
| NumberingPlan | Keine Änderung |
| Devconst | Keine Änderung |
| HVT | **Neu** — nur wenn `-IncludePenData` |

## QA Test Results
_To be added by /qa_

## Deployment
**Deployed:** 2026-03-24
**Version:** M29.20260324
**Tag:** v29.0
**Script:** `os4k2excel.ps1` mit Parameter `-IncludePenData`

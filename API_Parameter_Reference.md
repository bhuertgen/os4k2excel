# api2hipath.exe — Parameter Referenz

**Version:** 7.00.000
**Zweck:** Import/Export-Interface zu HiPath 4000 Manager
**Dateiformat:** CSV/Text mit Feldnamen in Kopfzeile

---

## Befehlszeilen-Syntax

```powershell
api2hipath.exe [SELECT|INSERT|UPDATE|DELETE|MIXED] -l usr -p passwd -h serverip -o odfname [options] filename
```

---

## Operationen

| Operation | Parameter | Zweck | Beispiel |
|-----------|-----------|-------|----------|
| **SELECT** (Lesen) | `-s fields [-w clause]` | Daten abfragen | `api2hipath -l user -p pass -h 10.0.0.1 -o PORT -s "extension,domain" -c "\|" output.csv` |
| **INSERT** | `-i` | Neue Datensätze einfügen | `api2hipath -i -l user -p pass -h 10.0.0.1 -o PORT -c "\|" input.txt` |
| **UPDATE** | `-u` | Vorhandene Datensätze ändern | `api2hipath -u -l user -p pass -h 10.0.0.1 -o PORT -c "\|" -r "myorder" input.txt` |
| **UPDATE/INSERT** | `-ui` | UPDATE oder INSERT falls Key nicht existiert | `api2hipath -ui -l user -p pass -h 10.0.0.1 -o PORT -c "\|" input.txt` |
| **DELETE** | `-d` | Datensätze löschen | `api2hipath -d -l user -p pass -h 10.0.0.1 -o PORT -c "\|" -r "myorder" input.txt` |

---

## Verbindungs-Parameter (immer erforderlich)

| Parameter | Pflicht | Beispiel | Beschreibung |
|-----------|---------|----------|-------------|
| `-l usr` | ✅ | `-l "admin"` | HiPath4000 Benutzername mit API-Berechtigung |
| `-p passwd` | ✅ | `-p "geheim123"` | HiPath4000 Passwort |
| `-h serverip` | ✅ | `-h "192.0.2.100"` | HiPath4000 Manager IP-Adresse oder Hostname |
| `-o odfname` | ✅ | `-o "PORT"` | ODF-Name (Tabellenname: PORT, PICKUP_SUB, HUNTGRP_SUB, etc.) |

---

## SELECT-Parameter (nur für Lesezugriff)

| Parameter | Pflicht | Beispiel | Beschreibung |
|-----------|---------|----------|-------------|
| `-s fields` | ✅ | `-s "extension,domain,switch_name"` | Feldnamen (kommagetrennt) oder `*` für alle Felder |
| `-w clause` | ❌ | `-w "extension='101' AND domain='DOM1'"` | WHERE-Klausel (SQL-Syntax) zur Filterung |
| `-t` | ❌ | `-t` | Ausgabe in GROSSBUCHSTABEN konvertieren |
| `-n` | ❌ | `-n` | Keine Kopfzeile (Feldnamen) in Ausgabe |
| `-f timefrom` | ❌ | `-f "2026-03-10 14:00:00"` | Delta-SELECT: Start-Zeit (Format: JJJJ-MM-DD HH:MM:SS) |
| `-e timeto` | ❌ | `-e "2026-03-10 15:00:00"` | Delta-SELECT: End-Zeit |

---

## INSERT/UPDATE/DELETE-Parameter

| Parameter | Pflicht | Beispiel | Beschreibung |
|-----------|---------|----------|-------------|
| `-i` / `-u` / `-d` / `-ui` | ✅ | `-d` | Operation: Insert, Update, Delete, oder Update-or-Insert |
| `-r orderid` | ❌ | `-r "OS4K000001"` | Order-ID für Tracking in ACTION_CONTROL (max. 12 Zeichen, wird kombiniert mit laufender Nummer) |

---

## Allgemeine Parameter

| Parameter | Pflicht | Beispiel | Beschreibung |
|-----------|---------|----------|-------------|
| `-c sep` | ❌ | `-c "\|"` | Feldtrennzeichen (Default: Komma `,`; üblich: Pipe `\|`) |
| `-m max` | ❌ | `-m "100"` | Max. Anzahl von Datensätzen (Limit für SELECT) |
| `-z` | ❌ | `-z` | Verschlüsselte Kommunikation (TLS) — nur V7.0+ |
| `-k` | ❌ | `-k` | Bei UPDATE/DELETE: Verwende `WHERE key_value IS NULL` wenn Feld leer |
| `-y` | ❌ | `-y` | Bei UPDATE: Ignoriere leere Felder |
| `filename` | ✅ | `output.csv` | Eingabe- oder Ausgabedatei (letztes Argument) |

---

## Input-Dateiformat

### Struktur
```
[Kopfzeile: Feldnamen]
[Datenzeilen: Werte]
```

### Beispiel (Port-Tabelle)
```
extension|domain|*switch_name|
101|DOM1|SWITCH1|
102|DOM1|SWITCH1|
```

### Regeln

- **Erste Zeile:** Feldnamen, getrennt durch `-c` Parameter (Default: Komma)
- **Key-Felder:** Mit `*` Präfix markieren (für UPDATE/DELETE)
- **Ignorierte Felder:** Mit `-` Präfix markieren (werden nicht verarbeitet)
- **DELETE-Mode:** Alle Felder sind automatisch Keys (außer `-` markierte)
- **Letztes Zeichen:** Trennzeichen am Ende jeder Zeile erforderlich (z.B. `|`)

---

## Praktische Beispiele

### 1. SELECT — Alle Extensions aus PORT

**PowerShell:**
```powershell
api2hipath.exe -l "admin" -p "passwd" -h "192.0.2.100" `
    -o "PORT" -s "extension,domain,switch_name" `
    -c "|" "output.csv"
```

**Ausgabe (output.csv):**
```
extension|domain|switch_name|
101|DOM1|SWITCH1|
102|DOM1|SWITCH1|
999|DOM2|SWITCH2|
```

---

### 2. SELECT mit WHERE — Nur Extensions in DOM1

**PowerShell:**
```powershell
api2hipath.exe -l "admin" -p "passwd" -h "192.0.2.100" `
    -o "PORT" -s "extension,domain" `
    -c "|" -w "domain='DOM1'" "dom1_ports.csv"
```

**Ausgabe:**
```
extension|domain|
101|DOM1|
102|DOM1|
```

---

### 3. INSERT — Neue Extensions hinzufügen

**Input-Datei (new_ports.txt):**
```
extension|domain|switch_name|
888|DOM1|SWITCH1|
889|DOM1|SWITCH1|
```

**PowerShell:**
```powershell
api2hipath.exe -i -l "admin" -p "passwd" -h "192.0.2.100" `
    -o "PORT" -c "|" -r "INS001" "new_ports.txt"
```

**Ergebnis:** 2 neue Extensions (888, 889) in PORT-Tabelle eingefügt

---

### 4. UPDATE — Extensions ändern (mit Key)

**Input-Datei (update_ports.txt):**
```
*extension|domain|
101|DOM2|
102|DOM2|
```

**PowerShell:**
```powershell
api2hipath.exe -u -l "admin" -p "passwd" -h "192.0.2.100" `
    -o "PORT" -c "|" -r "UPD001" "update_ports.txt"
```

**Ergebnis:** Extension 101 und 102 bekommen neuen Domain = DOM2

---

### 5. DELETE — Extensions löschen

**Input-Datei (delete_ports.txt):**
```
extension|
999|
888|
889|
```

**PowerShell:**
```powershell
api2hipath.exe -d -l "admin" -p "passwd" -h "192.0.2.100" `
    -o "PORT" -c "|" -r "DEL001" "delete_ports.txt"
```

**Ergebnis:** Extensions 999, 888, 889 gelöscht (Exit-Code 0)

**Mit Fehler (z.B. Extension in Gruppen):**
```
[Exit-Code: 13]
Error: "Operation not allowed"
```
(Bedeutet: Extension ist noch in PICKUP_SUB/HUNTGRP_SUB registriert)

---

### 6. UPDATE/INSERT kombiniert

**Input-Datei (mixed.txt):**
```
*extension|domain|status|
101|DOM1|active|
999|DOM1|new|
```

**PowerShell:**
```powershell
api2hipath.exe -ui -l "admin" -p "passwd" -h "192.0.2.100" `
    -o "PORT" -c "|" "mixed.txt"
```

**Ergebnis:**
- Extension 101: UPDATE (existiert bereits)
- Extension 999: INSERT (neue Extension)

---

### 7. DELETE mit Fehlerbehandlung

**Szenario:** Extension 101 noch in PICKUP_SUB registriert

**Input (delete_err.txt):**
```
extension|
101|
```

**PowerShell:**
```powershell
api2hipath.exe -d -l "admin" -p "passwd" -h "192.0.2.100" `
    -o "PORT" -c "|" "delete_err.txt"

# Exit-Code: 13 (Operation not allowed)
# → Extension 101 kann nicht gelöscht werden
```

**Lösung:**
1. Erst PICKUP_SUB löschen
2. Dann HUNTGRP_SUB löschen
3. Erst dann PORT löschen

---

## EXIT-CODES

| Code | Bedeutung | Aktion |
|------|-----------|--------|
| **0** | ✅ Erfolg | Operation abgeschlossen |
| **1** | ⚠️ Warnung | Teilweise erfolgreich |
| **13** | ❌ Operation not allowed | Keine Berechtigung oder ungültige Operation (z.B. DELETE bei existierenden Dependencies) |
| **Andere** | ❌ Fehler | Verbindung, Syntax, oder API-Fehler |

---

## Häufige Fehler & Lösungen

| Fehler | Ursache | Lösung |
|--------|--------|--------|
| `missing or wrong operation!` | Kein `-s` / `-i` / `-u` / `-d` | Parameter hinzufügen |
| `missing odfname !` | `-o TABELLE` fehlt | ODF-Namen angeben (PORT, PICKUP_SUB, etc.) |
| Exit-Code 13 bei DELETE | Dependencies existieren | Erst Gruppen-Einträge entfernen, dann PORT |
| Exit-Code 1 bei INSERT | Daten-Fehler (z.B. Duplikat Key) | Input-Datei prüfen, Keys eindeutig machen |
| Verbindungsfehler | `-h` / `-l` / `-p` falsch | IP, Benutzer, Passwort prüfen |

---

## OS4K-3 Spezifische Anwendung

### Delete NSt in korrekter Reihenfolge

```powershell
# 1. DELETE PICKUP_SUB
api2hipath.exe -d -l "admin" -p "passwd" -h "192.0.2.100" `
    -o "PICKUP_SUB" -c "|" -r "OS4Kdel1" "delete_pickup.txt"

# 2. DELETE HUNTGRP_SUB
api2hipath.exe -d -l "admin" -p "passwd" -h "192.0.2.100" `
    -o "HUNTGRP_SUB" -c "|" -r "OS4Kdel2" "delete_huntgrp.txt"

# 3. DELETE PORT
api2hipath.exe -d -l "admin" -p "passwd" -h "192.0.2.100" `
    -o "PORT" -c "|" -r "OS4Kdel3" "delete_port.txt"
```

**WICHTIG:** Reihenfolge ist ZWINGEND!
Wenn Step 1 oder 2 fehlschlägt → Step 3 NICHT durchführen!

---

## Quellen

- HiPath 4000 API Documentation
- OpenScape 4000 Manager V11 Export Table
- api2hipath.exe Help (Version 7.00.000)

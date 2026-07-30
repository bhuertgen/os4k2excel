# OS4K-3: Delete NSt via CSV Import with Dependency Management

## Status: In Review
**Created:** 2026-03-10
**Last Updated:** 2026-03-24

## Dependencies
- Requires: OS4K-1 (ETL-Pipeline, API connectivity)
- Requires: OS4K-2 (existing logging & backup infrastructure to reuse)

---

## Vision
Ein PowerShell-basiertes Verwaltungsscript, das es Administratoren ermöglicht, Rufnummern (NSts) sicher und nachverfolgbar über CSV-Import zu löschen. Das Script prüft automatisch, ob NSts in PICKUP- oder HUNTGRP-Gruppen vorhanden sind, entfernt sie dort ggf. zuerst und löscht erst dann die NSt — ohne die Gruppen selbst zu zerstören. Mit Backup, temporärem Script zur Bestätigung und vollständigem Logging aller Operationen.

---

## User Stories

- **Als IT-Administrator** möchte ich NSts per CSV-Datei (statt einzeln manuell) löschen können, **damit** ich mehrere Nummern in einem Arbeitsschritt verwalte.

- **Als Techniker vor Ort** möchte das Script automatisch prüfen, ob eine NSt noch in Umleitungsgruppen oder Sammelanschlüssen verwendet wird, **damit** ich keine Inkonsistenzen im System hinterlasse.

- **Als IT-Leiter** möchte ein Audit-Log aller Löschoperationen, **damit** ich nachvollziehen kann, wer wann welche NSts entfernt hat und warum.

- **Als Admin** möchte eine Datensicherung vor den Änderungen, **damit** ich im Fehlerfall zurückkehren kann.

- **Als Benutzer** möchte flexibel zwischen Einzeln- und Batch-Löschung wählen können und vor dem Löschen Best\u00e4tigung geben, **damit** ich Unfälle vermeide.

---

## Acceptance Criteria

### CSV Import & Validierung
- [ ] Script liest CSV-Datei mit Spalten: **unique_key** (Datenbank-ID), extension, domain, switch_name, delete_confirmation (ja/nein), reason (optional)
- [ ] Script validiert `unique_key` als numerisch und gegen die PORT-Tabelle via API
- [ ] Script warnt bei fehlenden/ungültigen `unique_key`-Werten; Admin kann skip/abort wählen
- [ ] CSV-Format ist dokumentiert und mit Beispiel-Template vorhanden
- [ ] Duplikate in `unique_key` werden erkannt und Warnung ausgegeben

### Abhängigkeits-Prüfung (Informativ)
- [ ] Script fragt automatisch API ab: Ist NSt in irgendeiner PICKUP-Gruppe vorhanden?
- [ ] Script fragt automatisch API ab: Ist NSt in irgendeiner HUNTGRP-Gruppe vorhanden?
- [ ] Script zeigt Liste aller Gruppen, in denen NSt vorhanden ist (mit Gruppennamen & Typ)
- [ ] **Diese Prüfung ist INFORMATIV** — die API wird Fehler zurückgeben wenn Dependencies nicht entfernt werden

### Abhängigkeiten Entfernen (NOTWENDIG vor PORT-DELETE!)
- [ ] **KRITISCH:** Script MUSS zuerst NSt aus PICKUP_SUB entfernen (separate DELETE-Befehle pro Gruppe)
- [ ] **KRITISCH:** Script MUSS zuerst NSt aus HUNTGRP_SUB entfernen (separate DELETE-Befehle pro Gruppe)
- [ ] **ERST DANN:** Script löscht NSt aus PORT (würde sonst mit Fehler -56 fehlschlagen)
- [ ] Wenn NSt in mehreren Gruppen: Alle PICKUP_SUB/HUNTGRP_SUB Einträge werden nacheinander in einer REQUEST gelöscht (max. 1 Fehler = ganze NSt skippt)
- [ ] Admin wird vor dem Entfernen gefragt: "NSt 101 in PICKUP 'Einkauf' gefunden. Entfernen? (j/n)"
- [ ] **Ohne Bestätigung:** NSt NICHT aus Gruppen entfernen, die NSt überspringen (Skip in Log)

### Backup vor Löschung
- [ ] Vor jeder Lösch-Aktion: Vollständiger Snapshot (PORT, PICKUP, HUNTGRP, CFW, PERSPORT) als Excel-Export
- [ ] Zusätzlich: Gezielter Export der NSts, die gelöscht werden (mit allen ihren Konfigurationen)
- [ ] Backup-Datei wird mit Zeitstempel benannt: `OS4K-NST-DELETE-BACKUP-{YYYY-MM-DD-HHmmss}.xlsx`

### Temporäres Script & Best\u00e4tigung
- [ ] Script erstellt temporäres PowerShell-File (`delete_nst_[timestamp].ps1`) mit allen Lösch-Befehlen
- [ ] Admin kann vor Ausführung das Script reviewen (Show-Content oder Datei öffnen)
- [ ] Zwei Modi:
  - **Manuell:** Admin drückt Taste zum Bestätigen (read-host "Fortfahren? (j/n)")
  - **Automatisch mit Timeout:** "Starte in 30 Sekunden... [30] [29]..." (Admin kann noch abbrechen)
- [ ] Admin wählt Modus beim Start des Hauptscripts

### Ausführung der Löschung
- [ ] Löschung nur über API (api2hipath.exe) — keine Direktzugriffe
- [ ] Für jede NSt: Erst aus Gruppen entfernen, dann NSt-Eintrag löschen
- [ ] Bei Fehler während der Ausführung: Admin wird gefragt pro NSt — Skip und weitermachen oder ganz abbrechen?

### Logging
- [ ] Log-Datei: `OS4K-NST-DELETE-{YYYY-MM-DD}.log`
- [ ] Einträge mit Zeitstempel für jede Operation:
  - NSt-Nummern-Validierung (existiert ja/nein)
  - Abhängigkeits-Prüfung (in welchen Gruppen)
  - Entfernen aus PICKUP/HUNTGRP (erfolg/fehler)
  - NSt-Löschung (erfolg/fehler)
  - Admin-Bestätigung (wer, wann, welche Entscheidung)
  - Fehler & Recovery-Schritte
- [ ] Log ist lesbar formatiert (kein reines syslog-Format)

### Modus-Flexibilität
- [ ] Batch-Modus: Mehrere NSts in einer CSV, alle mit einer Bestätigung verarbeiten
- [ ] Einzeln-Modus: Eine NSt pro Lauf
- [ ] Script erkennt Anzahl der NSts in CSV und schlägt Modus vor, lässt sich aber überschreiben

### Job Tracking & RMX Synchronisation
- [ ] Script generiert ORDER_ID im Format `OS4K[000001-999999]` (max 12 Zeichen)
- [ ] REQUEST-Datei enthält: `#ORDER_ID=...` vor DELETE-Befehlen
- [ ] Script pollt ACTION_CONTROL Tabelle bis `action_status='067'` (Done)
- [ ] Abfrage-Intervall: ~2-5 Sekunden, Timeout nach 5 Minuten
- [ ] `action_error` wird geprüft: `065`=OK, `061`=RMX-Fehler, `063`=Manager-Fehler
- [ ] Bei `action_status='064'` (Running): Admin wird informiert "Job läuft, bitte warten..."
- [ ] Bei `action_status='062'` (No Response): RMX-Timeout → Backup behalten, Manual-Check
- [ ] Bei Error (`061`, `063`, `066`): Log-Eintrag + Admin-Warnung (aber NSt ist aus Manager gelöscht!)
- [ ] Nach erfolgreicher RMX-Sync: NSt ist sowohl in Manager als auch in RMX gelöscht

---

### Multiple Identifier unter einer ORDER_ID

**Wichtig für Error-Handling:**

Jeder DELETE-Befehl bekommt seinen eigenen Identifier (1, 2, 3) UND seine EIGENE Response:

```
REQUEST-Datei:
#DeleteNSts
#ORDER_ID=OS4K000001
1;DELETE;PICKUP_SUB;12345
2;DELETE;HUNTGRP_SUB;12345
3;DELETE;PORT;12345
#@

RESPONSE-Datei (JEDER Befehl hat eigene Zeile!):
#DeleteNSts; 2026-03-10 14:32:15.00
#1;1              ← Identifier 1: erfolgreich (1 record deleted)
#2;1              ← Identifier 2: erfolgreich
#3;-56            ← Identifier 3: FEHLER! (Operation not allowed)
#@
```

**Auswertungs-Logik:**
- Wenn Identifier 1 oder 2 Fehler → NSt nicht aus Gruppen entfernt → Skip die ganze NSt, nicht Identifier 3 versuchen
- Wenn Identifier 3 Fehler → Sollte nicht vorkommen wenn 1+2 erfolgreich waren → Log-Fehler + Admin-Warnung
- Wenn alle erfolgreich → ACTION_CONTROL mit `order_id='OS4K000001'` polling für RMX-Sync Status

---

## Edge Cases

1. **NSt existiert nicht:** Script meldet "NSt 999 nicht gefunden" und fragt, ob weitere NSts verarbeitet werden sollen (Skip oder Abbruch).

2. **NSt ist in 5 Gruppen:** Script listet alle auf, fragt für jede Bestätigung. Wenn Admin eine ablehnt → nur diese Gruppe wird nicht bearbeitet, andere ja.

3. **API-Fehler während Abhängigkeitsprüfung:** Script meldet Fehler, fragt: "Abbrechen und Backup behalten? (j/n)" — im Fall von Ja wird alles zurückgerollt.

4. **Abhängigkeits-Entfernen schlägt fehl:** Z.B. "NSt kann nicht aus PICKUP entfernt werden, weil PICKUP der Gruppe ist." → Script meldet, fragt, ob NSt trotzdem gelöscht wird (riskant) oder abgebrochen wird.

5. **Duplikat-NSts in CSV:** Script warnt und entfernt Duplikate, zeigt Liste der eindeutigen NSts.

6. **CSV hat falsche Spalten:** Script zeigt erwartete vs. gefundene Spalten und bricht ab mit Hinweis auf Template.

7. **Leere Bemerkung/Grund:** Optional, nicht erforderlich. Script nutzt Default "Gelöscht via Batch-Import" im Log.

8. **Admin bricht während Timeout-Countdown ab:** Script stoppt sofort, Backup wird beibehalten, Log dokumentiert den Abbruch.

9. **Temporäres Lösch-Script beschädigt:** Script generiert neues, zeigt Unterschied und fragt zur Best\u00e4tigung.

10. **NSt wird gerade von anderem Admin bearbeitet:** Script kann das nicht prüfen (kein Lock im OS4K), aber Fehler der API wird gezeigt → Admin kann manuell entscheiden.

---

## Technical Requirements

### Performance
- CSV-Import: < 1 Sekunde für 100 NSts
- Abhängigkeits-Abfrage pro NSt: < 2 Sekunden (via API)
- Backup-Export: < 30 Sekunden (wie os4k2excel.ps1)
- Gesamtdauer für 10 NSts mit Abhängigkeiten: < 3 Minuten

### Security
- API-Credentials werden wie in os4k2excel.ps1 übergeben (Parameter, nicht geloggt)
- Backup-Dateien haben normale Filesystem-Permissions (Admin kann sie einsehen)
- Log-Datei dokumentiert Änderungen, aber nicht die Credentials
- Keine Speicherung von Secrets im temporären Script

### Compatibility
- PowerShell 5.1 (Windows)
- Benötigt: api2hipath.exe (wie OS4K-1)
- Benötigt: ImportExcel Modul (wie OS4K-1)
- Keine zusätzlichen Module erforderlich

### Logging
- Zentrale Log-Datei im gleichen Verzeichnis wie os4k2excel.ps1 Output
- Reuse von Logging-Funktionen aus OS4K-1 falls vorhanden
- Log-Retention: Mindestens 90 Tage (für Audit)

---

## Implementation Notes

- **Reuse von OS4K-1:** API-Aufrufe, Hashtable-Lookups für Domain/Switch/NSt-Mapping, Logging-Funktionen
- **Inspiration von OS4K-2:** Bestätigungs-Workflows (temporäres Script-Pattern)
- **CSV-Template:** Wird mit dem Script mitgeliefert oder als Beispiel im Repository vorhanden
- **Error Recovery:** Backup immer beibehalten; Admin kann bei Fehler manuell via os4k2excel.ps1 neuen Export machen

### KRITISCH: DELETE-Reihenfolge & Abhängigkeiten

**GUI-Verhalten:** Fehler/Hinweis wenn Extension noch in Gruppen vorhanden ist
→ **Folgerung für API:** DELETE;PORT wird wahrscheinlich mit Fehler -56 ("Operation not allowed") fehlschlagen wenn noch PICKUP_SUB/HUNTGRP_SUB Einträge existieren!

**Implementierungs-Konsequenz:**
```
REQUEST-Struktur MUSS sein:
#DeleteNSts
#ORDER_ID=OS4K000001
1;DELETE;PICKUP_SUB;[unique_key]    ← VORAB: Aus Pickup-Gruppen entfernen
2;DELETE;HUNTGRP_SUB;[unique_key]   ← VORAB: Aus Hunting-Gruppen entfernen
3;DELETE;PORT;[unique_key]          ← ERST DANN: NSt selbst löschen
#@
```

**Response-Auswertung (Pro Befehl separate Behandlung!):**
```
#1;1        ← PICKUP_SUB erfolgreich gelöscht
#2;1        ← HUNTGRP_SUB erfolgreich gelöscht
#3;1        ← PORT erfolgreich gelöscht
```

**Fehlerbehandlung:**
- IF #3 = -56 THEN NSt war noch in Gruppen registriert → Backup beibehalten, Admin manuell prüfen
- IF #1 oder #2 = -56 THEN Unerwarteter Fehler (Log + Admin-Warnung)

### Test-Plan für Implementation

**VOR Implementierung des Scripts testen:**
1. Eine Extension die noch in PICKUP_SUB ist versuchen mit DELETE;PORT zu löschen
   - Erwartet: Fehler -56 oder ähnlich
   - Dann: DELETE;PICKUP_SUB erst, dann DELETE;PORT
2. Eine Extension die noch in HUNTGRP_SUB ist, gleiches Szenario
3. Multiple Befehle unter einer ORDER_ID testen (ob jeder identifier sein Response bekommt)
4. ACTION_CONTROL Einträge abfragen pro identifier vs. pro ORDER_ID (wichtig für Fehlerbehandlung)

---

## Tech Design (Solution Architect)

### CSV-Template & Input-Format

**Spalten (Pipe-delimited):**
```
unique_key|extension|domain|switch_name|delete_confirmation|reason
```

**Beispiel-Zeilen:**
```
12345|101|DOMAIN1|SWITCH1|ja|Mitarbeiter ausgetreten
54321|102|DOMAIN1|SWITCH1|ja|Abteilung geschlossen
67890|103|DOMAIN2|SWITCH2|nein|Überprüfung erforderlich
```

**Format-Regeln:**
- `unique_key`: Numerisch, eindeutig, aus der Export-Excel
- `extension`: Lesbar, nur zur Bestätigung (nicht als Key verwendet)
- `delete_confirmation`: Genau "ja" oder "nein" (case-insensitive)
- `reason`: Optional, max. 200 Zeichen, im Log dokumentiert
- **Duplikate:** Werden erkannt und dedupliziert; Admin wird gewarnt

---

### Komponenten-Struktur

```
delete_nst.ps1 (Haupt-Script)
│
├── Phase 1: INIT
│   ├── Parameter einlesen (ApiHost, ApiUser, ApiPassword, CsvPath, OutputPath)
│   ├── Logging initialisieren
│   └── api2hipath.exe Verfügbarkeit prüfen
│
├── Phase 2: CSV-VERARBEITUNG
│   ├── CSV einlesen & Spalten validieren
│   ├── Duplikate entfernen (deduplizieren nach unique_key)
│   ├── Per unique_key: NSt-Existenz via API prüfen
│   │   └── SELECT unique_key, extension FROM PORT WHERE unique_key=X
│   └── Fehlerhafte Einträge dem Admin zeigen (skip/abort)
│
├── Phase 3: ABHÄNGIGKEITS-ANALYSE
│   ├── Per extension: Pickup-Abfrage
│   │   └── SELECT pickupgrpnum FROM PICKUP_SUB WHERE extension=X
│   ├── Per extension: Huntgrp-Abfrage
│   │   └── SELECT huntgrpnum, group_idx, service FROM HUNTGRP_SUB WHERE extension=X
│   └── Abhängigkeits-Report anzeigen (Gruppennamen + Typ)
│
├── Phase 4: BACKUP
│   ├── Vollständiger Snapshot via os4k2excel.ps1 Logik
│   │   (PORT, PICKUP_SUB, HUNTGRP_SUB, PICKUPGRP, HUNTGRP)
│   └── Gezielter NSt-Export → OS4K-NST-DELETE-BACKUP-{timestamp}.xlsx
│
├── Phase 5: BESTÄTIGUNG
│   ├── Modus A: Manuell (read-host "Fortfahren? (j/n)")
│   └── Modus B: Timeout-Countdown mit Abbruch möglich
│
├── Phase 6: AUSFÜHRUNG (KRITISCHE REIHENFOLGE!)
│   ├── Temporäres Script generieren (delete_nst_[timestamp].ps1)
│   │   └── Mit #ORDER_ID=OS4K[000001-999999] für Tracking
│   ├── Admin kann Script reviewen (MUSS Reihenfolge verstehen!)
│   ├── REQUEST-Struktur für jede NSt:
│   │   ```
│   │   #DeleteNSts_[unique_key]
│   │   #ORDER_ID=OS4K[sekunden-sequenz]
│   │   1;DELETE;PICKUP_SUB;[unique_key]      ← VORAB
│   │   2;DELETE;HUNTGRP_SUB;[unique_key]     ← VORAB
│   │   3;DELETE;PORT;[unique_key]            ← ERST DANN
│   │   #@
│   │   ```
│   ├── RESPONSE-Handling (JEDER Befehl hat eigenes Response!):
│   │   ├── IF #1 = -56 → NSt in PICKUP_SUB vergessen? → Abbruch + Backup behalten
│   │   ├── IF #2 = -56 → NSt in HUNTGRP_SUB vergessen? → Abbruch + Backup behalten
│   │   ├── IF #3 = -56 → Fehler! NSt konnte nicht gelöscht werden → Skip + Log
│   │   └── IF alle erfolgreich → ACTION_CONTROL polling für RMX-Sync
│   └── Bei Fehler in #1 oder #2: Admin entscheidet (skip NSt oder Abbruch alles)
│
└── Phase 7: ABSCHLUSS
    ├── Log abschließen
    └── Zusammenfassung: gelöscht / fehlgeschlagen / übersprungen
```

---

### API-Unterstützung & Datenmodell

**DELETE-Befehl auf API-Ebene unterstützt:**
```
DELETE;PORT;[unique_key]          → NSt löschen
DELETE;PICKUP_SUB;[unique_key]    → NSt aus Pickup entfernen
DELETE;HUNTGRP_SUB;[unique_key]   → NSt aus Huntgrp entfernen
```

**Abhängigkeits-Datenfluss:**
```
CSV (unique_key)
    ↓
→ SELECT PORT WHERE unique_key=X → extension
    ↓
→ SELECT PICKUP_SUB WHERE extension=Y → pickupgrpnum-Liste
→ SELECT HUNTGRP_SUB WHERE extension=Y → (huntgrpnum, group_idx, service)-Liste
    ↓
→ DELETE aus allen PICKUP_SUB-Einträgen
→ DELETE aus allen HUNTGRP_SUB-Einträgen
    ↓
→ DELETE FROM PORT WHERE unique_key=X
```

**Im RAM verwaltete Struktur (pro unique_key):**
```
[PSCustomObject]@{
    unique_key               = 12345
    extension                = "101"
    domain                   = "DOMAIN1"
    switch_name              = "SWITCH1"
    delete_confirmation      = "ja"
    reason                   = "Mitarbeiter ausgetreten"
    exists_in_port           = $true
    pickup_groups            = @(
        @{ num = "100"; name = "Einkauf"; members = 5 }
    )
    huntgrp_groups           = @(
        @{ num = "200"; name = "Support"; group_idx = 1; members = 8 }
    )
    status                   = "pending"  # oder "deleted", "failed", "skipped"
    error_message            = $null
}
```

---

### Ausgabe-Dateien

| Datei | Zweck | Format |
|-------|-------|--------|
| `OS4K-NST-DELETE-BACKUP-{YYYY-MM-DD-HHmmss}.xlsx` | Snapshot vor Löschung | Excel (MultiSheet wie os4k2excel.ps1) |
| `OS4K-NST-DELETE-{YYYY-MM-DD}.log` | Audit-Log | Text, UTF-8 with BOM |
| `delete_nst_{timestamp}.ps1` | Temporäres Lösch-Script | PowerShell (wird nach Bestätigung ausgeführt) |

---

### Technische Entscheidungen (PM-gerecht erklärt)

| Decision | Wahl | Grund |
|----------|------|-------|
| **Identifier** | `unique_key` (nicht Extension) | Eindeutig, direkt aus Export-Excel, DB-native Lösung |
| **Script-Struktur** | Eigenständig (`delete_nst.ps1`) | Klare Verantwortung, keine Vermischung mit ETL-Pipeline |
| **Bestätigung** | Zwei Modi (manuell + Timeout) | Flexibilität: sorgfältig vs. halbautomatisch |
| **Backup-Format** | Excel | Konsistenz mit bestehenden Tools, Admin kennt das Format |
| **Fehlerbehandlung** | Pro-NSt-Entscheidung | Admin bleibt in Kontrolle, kein stilles Fehlschlag |
| **API-Ebene** | DELETE-Befehle (keine Direktzugriffe) | Sicher, audit-trail via Log |

---

### Job Tracking & RMX Synchronisation

**REQUEST-Datei mit ORDER_ID (max 12 Zeichen):**
```
#DeleteNSts
#ORDER_ID=OS4K000001
1;DELETE;PICKUP_SUB;[unique_key_1]
2;DELETE;PICKUP_SUB;[unique_key_2]
3;DELETE;HUNTGRP_SUB;[unique_key_1]
4;DELETE;PORT;12345
5;DELETE;PORT;12345
#@
```

**ORDER_ID Format:**
- Prefix: `OS4K` (4 Zeichen max)
- Nummer: `000001` bis `999999` (6-9 Ziffern)
- Total: `OS4K000001` (10 Zeichen, max 12 möglich) ✓

**Automatische ACTION_CONTROL Einträge:**
Manager erstellt pro DELETE automatisch einen Eintrag mit:
- `order_id` = `OS4K000001`
- `action_status` = anfangs `064` (Running) → später `067` (Done)
- `action_error` = `065` (No Error) oder `061` (Error) oder `063` (DMS-Error)
- `extension` = die gelöschte NSt (Audit-Trail)
- `act_time_create` / `act_time_update` = Zeitstempel

**Tracking-Workflow (nach DELETE):**
```
LOOP (mit ~1-5 Sekunden Intervall):
  SELECT action_status, action_error, act_time_update
    FROM ACTION_CONTROL
    WHERE order_id='OS4K000001'
      AND extension IN ('101', '102', ...)

  IF action_status = '067' THEN
    → Job beendet, prüfe action_error
  ELSE IF action_status = '064' THEN
    → Noch laufend, nächste Abfrage in 2 Sekunden
  ELSE IF action_status = '062' THEN
    → RMX antwortet nicht (Timeout), Admin manuell prüfen
  END IF
END LOOP
```

**Error-Codes in action_error:**
| Code | Bedeutung | Aktion |
|------|-----------|--------|
| `065` | No Error (RMX-Sync OK) | ✓ NSt erfolgreich gelöscht |
| `060` | Done without error | ✓ NSt erfolgreich gelöscht |
| `061` | Done with error | ✗ Fehler bei RMX-Sync, Manual prüfen |
| `063` | DMS Error | ✗ Manager-Fehler, Support kontaktieren |
| `066` | Error | ✗ Generischer Fehler, Log prüfen |

**Logging mit ORDER_ID:**
```
[2026-03-10 14:32:15] DELETE START: unique_key=12345, extension=101
[2026-03-10 14:32:15] REQUEST-Datei: #ORDER_ID=OS4K000001
[2026-03-10 14:32:16] API-Response: Positive (1 record deleted)
[2026-03-10 14:32:17] ACTION_CONTROL polling started...
[2026-03-10 14:32:18] Status: 064 (Running) - waiting for RMX...
[2026-03-10 14:32:35] Status: 067 (Done), Error: 065 (No Error)
[2026-03-10 14:32:35] ✓ DELETE COMPLETE: extension 101 gelöscht (Manager + RMX)
```

---

### Abhängigkeiten & Ressourcen

**Keine neuen Module erforderlich:**
- `api2hipath.exe` — bereits vorhanden (OS4K-1)
- `ImportExcel` — bereits installiert (OS4K-1)
- PowerShell 5.1 — Standard auf Windows

**Reuse von OS4K-1:**
- API-Aufruf-Muster (Start-Process mit api2hipath.exe)
- Logging-Funktionen (falls vorhanden)
- Hashtable-Lookups für schnelle Verknüpfung
- Fehlerbehandlung & Retry-Logik

---

## QA Test Results

**Tested:** 2026-03-10 (Updated after commit 44c817e)
**Tester:** QA Engineer (AI)
**Scope:** Spec review, test_delete_api.ps1 code review (post-refactor), security audit
**Script version reviewed:** After 20 commits of iterative development (f8d5dbf simplified script, ddbe6ce fixed ACTION_CONTROL polling)

### BLOCKER: Implementation Does Not Exist

The main script `delete_nst.ps1` referenced in the tech design HAS NOT BEEN IMPLEMENTED. Only a validation/test script (`test_delete_api.ps1`) exists. The feature status in INDEX.md says "In Review" but there is no production implementation to functionally test against the acceptance criteria.

**QA was performed as a spec review + test script code review + security audit.**

---

### Acceptance Criteria Status

#### AC-1: CSV Import & Validierung
- [ ] NOT IMPLEMENTED: No `delete_nst.ps1` script exists
- [ ] NOT IMPLEMENTED: CSV reading with required columns (unique_key, extension, domain, switch_name, delete_confirmation, reason)
- [ ] NOT IMPLEMENTED: unique_key validation as numeric + against PORT table
- [ ] NOT IMPLEMENTED: Warning on missing/invalid unique_key values with skip/abort
- [ ] NOT TESTABLE: CSV format documented -- spec has documentation but no CSV template file ships with the repo
- [ ] NOT IMPLEMENTED: Duplicate unique_key detection

#### AC-2: Abhaengigkeits-Pruefung (Informativ)
- [ ] REMOVED: test_delete_api.ps1 no longer queries PICKUP_SUB (Phase 2 removed in commit f8d5dbf)
- [ ] REMOVED: test_delete_api.ps1 no longer queries HUNTGRP_SUB (Phase 3 removed in commit f8d5dbf)
- [ ] NOT IMPLEMENTED: Display of group names and types in a summary list
- [x] SPEC OK: Documented as informational (API will reject if deps not removed)

#### AC-3: Abhaengigkeiten Entfernen (NOTWENDIG vor PORT-DELETE!)
- [ ] BUG-SPEC-1 (HIGH): test_delete_api.ps1 was refactored (commit f8d5dbf) to remove Phases 2, 3, 5a, 5b, 5c. The script now performs a DIRECT PORT DELETE (Phase 4) without first removing PICKUP_SUB or HUNTGRP_SUB dependencies. This contradicts the spec's CRITICAL requirement that dependencies MUST be removed before PORT deletion. The test script's approach relies on the API to cascade-delete dependencies automatically, but the spec explicitly warns this will fail with error -56.
- [ ] NOT IMPLEMENTED: Per-group confirmation prompt ("NSt 101 in PICKUP 'Einkauf' gefunden. Entfernen? (j/n)")
- [ ] NOT IMPLEMENTED: Skip logic when admin declines removal from a group
- [ ] BUG-SPEC-2: Spec says "Alle PICKUP_SUB/HUNTGRP_SUB Eintraege werden nacheinander in einer REQUEST geloescht" but the tech design shows them as separate identifiers (1, 2, 3) in ONE request -- these are contradictory statements. Clarification needed: does the API cascade-delete dependencies when DELETE PORT is called, or must dependencies be removed first?

#### AC-4: Backup vor Loeschung
- [ ] NOT IMPLEMENTED: No backup generation code exists
- [ ] NOT IMPLEMENTED: No targeted NSt export before deletion
- [ ] NOT IMPLEMENTED: Backup file naming with timestamp

#### AC-5: Temporaeres Script & Bestaetigung
- [ ] PARTIAL: test_delete_api.ps1 has a Read-Host confirmation (Phase 5, only when Phase 4 fails) but not the full two-mode (manual/timeout) workflow
- [ ] NOT IMPLEMENTED: Temporary PowerShell script generation
- [ ] NOT IMPLEMENTED: Review capability before execution
- [ ] NOT IMPLEMENTED: Timeout countdown mode

#### AC-6: Ausfuehrung der Loeschung
- [x] SPEC OK: test_delete_api.ps1 uses api2hipath.exe exclusively (no direct access)
- [ ] BUG-11: Correct deletion order (PICKUP_SUB -> HUNTGRP_SUB -> PORT) is NOT demonstrated in the current test script. After refactoring, it does direct PORT DELETE only.
- [ ] NOT IMPLEMENTED: Per-NSt error handling with skip/abort choice

#### AC-7: Logging
- [x] PARTIAL: test_delete_api.ps1 has basic logging to a file with timestamps
- [ ] NOT IMPLEMENTED: Full audit log with all required entries (validation, dependency check, removal, deletion, admin confirmation, errors)
- [ ] NOT IMPLEMENTED: Log retention policy (90 days)
- [ ] BUG-SPEC-3: Log filename format inconsistency -- spec says `OS4K-NST-DELETE-{YYYY-MM-DD}.log` but test script uses `test_delete_{YYYY-MM-DD}.log`

#### AC-8: Modus-Flexibilitaet
- [ ] NOT IMPLEMENTED: Batch mode vs. single mode
- [ ] NOT IMPLEMENTED: Automatic mode suggestion based on CSV count

#### AC-9: Job Tracking & RMX Synchronisation
- [x] PASS: test_delete_api.ps1 demonstrates ACTION_CONTROL polling (Phase 6) with proper main-operation detection (unique_key != "0")
- [x] PASS: Polling with 2-second interval and 70-second timeout implemented
- [x] PARTIAL: ORDER_ID generation via New-OrderID function uses timestamp-based format (e.g., OS4K10143215). This is 12 chars and within the spec max, but uses a different scheme than spec's sequential OS4K[000001-999999].
- [ ] BUG-SPEC-4: ORDER_ID format mismatch -- spec says sequential `OS4K000001-999999`, implementation uses timestamp-based `OS4KDDHHMMSSMM`. The timestamp approach is arguably better for collision avoidance but diverges from the spec.
- [ ] PARTIAL: Error code handling covers 065 (OK) and 000 (OK) but does not handle 061, 063, 066 explicitly
- [x] PASS: Phase 7 queries ACTION_CONTROL log for complete audit trail per ORDER_ID

---

### Edge Cases Status

#### EC-1: NSt existiert nicht
- [x] PARTIAL: test_delete_api.ps1 exits on missing NSt (Phase 1) but no skip/continue option for batch processing

#### EC-2: NSt ist in 5 Gruppen
- [ ] NOT IMPLEMENTED: Per-group confirmation not built (dependency phases removed from test script)

#### EC-3: API-Fehler waehrend Abhaengigkeitspruefung
- [ ] NOT IMPLEMENTED: No abort-and-keep-backup flow

#### EC-4: Abhaengigkeits-Entfernen schlaegt fehl
- [ ] NOT IMPLEMENTED: No fallback prompt

#### EC-5: Duplikat-NSts in CSV
- [ ] NOT IMPLEMENTED: No deduplication logic

#### EC-6: CSV hat falsche Spalten
- [ ] NOT IMPLEMENTED: No column validation

#### EC-7: Leere Bemerkung/Grund
- [ ] NOT IMPLEMENTED: No default reason handling

#### EC-8: Admin bricht waehrend Timeout-Countdown ab
- [ ] NOT IMPLEMENTED: No countdown mode exists

#### EC-9: Temporaeres Loesch-Script beschaedigt
- [ ] NOT IMPLEMENTED: No temp script generation

#### EC-10: NSt wird gerade von anderem Admin bearbeitet
- [ ] NOT IMPLEMENTED: No concurrent-edit handling

#### EC-11 (NEW): Master SA Detection
- [x] PARTIAL: test_delete_api.ps1 detects Master SA (HUNTGRP type=001) after a failed PORT DELETE and generates a .mac file for manual ComWin deletion. This is not in the original spec but is a valid operational edge case. However, this flow only triggers when Phase 4 fails -- there is no upfront detection.

---

### Security Audit Results (Red Team)

#### SEC-1: Credential Handling in test_delete_api.ps1
- [ ] BUG-SEC-1 (HIGH): API credentials (ApiUser, ApiPassword) are logged in the log file via `Invoke-ApiWithLogging`. The function logs the full command including arguments at line 122 (`"Kommando: $fullCommand"`), and the argument array is dumped at lines 125-128. This means plaintext passwords appear in the log file. The `-p` argument containing the password is visible in both the console output (line 120) and the log file.
- [x] OK: Credentials are passed as parameters, not hardcoded in the script

#### SEC-2: Temporary File Security
- [ ] BUG-SEC-2 (MEDIUM): Input files for DELETE operations (`test_delete_p4_input_port_direct.txt`) are written to the OutputPath with no restricted permissions. Any user with read access to the directory can see what is being deleted.
- [ ] BUG-SEC-3 (LOW): Temporary input files and CSV output files (p1, p4, p5, p6, p7) are not cleaned up after execution. They remain on disk. The test script does clean up its own temp stdout/stderr files (lines 161-162) but not the API input/output files.

#### SEC-3: Input Validation
- [ ] BUG-SEC-4 (HIGH): TestExtension, TestUniqueKey, Domain, and Switch parameters are not sanitized before being interpolated into API command arguments (lines 215, 219, 356). If api2hipath.exe is susceptible to command injection via its `-w` (WHERE clause) parameter, an attacker could inject arbitrary queries. Example: `TestExtension` value of `' OR 1=1; --` could potentially be passed through. The Domain and Switch parameters (added since initial review) are also unsanitized at line 219.
- [ ] BUG-SEC-5 (MEDIUM): No validation that TestUniqueKey is actually numeric (spec requires numeric unique_key). Any string is accepted. The parameter validation at lines 49-56 only checks for empty strings, not format.

#### SEC-4: Exit Code Trust
- [ ] BUG-SEC-6 (MEDIUM): Phase 4 (line 297) checks only for exit code 0 (success) vs non-zero (failure). No documentation of what specific non-zero exit codes mean. An unexpected exit code could mask a security-relevant failure mode. The script assumes non-zero means "could be Master SA" which is an incorrect generalization.

#### SEC-5: Spec-Level Security Concerns
- [ ] BUG-SEC-7 (HIGH): The spec does not define any authentication/authorization for WHO can run the delete script. Any user with access to the script and API credentials can mass-delete NSts. There should be an explicit access control requirement.
- [ ] BUG-SEC-8 (MEDIUM): ORDER_ID collision risk. The timestamp-based New-OrderID uses DDHHMMSSMM format, which could collide if two admins run the script within the same 10ms window. The function truncates to 8 digits from a 10-digit timestamp (line 98), further increasing collision probability.
- [ ] BUG-SEC-9 (LOW): The spec states backup files use "normale Filesystem-Permissions" -- but for a deletion audit trail, backups should have read-only permissions set after creation to prevent tampering.

#### SEC-6: Data Exposure
- [x] OK: No secrets stored in generated .mac file (only extension, domain, switch info)
- [ ] BUG-SEC-10 (MEDIUM): The spec does not address what happens to the backup Excel file if deletion fails partway through. Partial state could be misleading for recovery.
- [ ] BUG-SEC-11 (LOW): The .mac file generated for Master SA contains the ORDER_ID which could be used to correlate with ACTION_CONTROL entries. This is acceptable for audit purposes but should be documented.

#### SEC-7: Hex Dump Exposure (NEW)
- [ ] BUG-SEC-12 (LOW): Write-ApiInputFile (line 182) outputs a hex dump of the first 50 bytes of every input file to both console and log. While useful for debugging, this exposes raw data content in the log permanently.

---

### Cross-Browser & Responsive Testing
- N/A: This is a PowerShell CLI tool, not a web application. No browser or responsive testing applicable.

---

### Bugs Found

#### BUG-1: No Implementation Exists (delete_nst.ps1)
- **Severity:** Critical
- **Steps to Reproduce:**
  1. Look for `delete_nst.ps1` in the repository
  2. Expected: The main deletion script exists as specified in the tech design
  3. Actual: Only `test_delete_api.ps1` (a validation script) exists. The feature cannot be used.
- **Priority:** Fix before deployment -- this IS the feature

#### BUG-2: API Credentials Logged in Plaintext
- **Severity:** High
- **Steps to Reproduce:**
  1. Run `test_delete_api.ps1` with valid credentials
  2. Open the generated log file `test_delete_{date}.log`
  3. Expected: Credentials are masked or omitted
  4. Actual: Full command line including `-p <PASSWORD>` is written to the log file (line 122 of `Invoke-ApiWithLogging`). The argument array dump at lines 125-128 also exposes the password as a separate indexed entry (e.g., `[3] = 'mypassword'`).
- **Priority:** Fix before deployment

#### BUG-3: No Input Sanitization on API Parameters
- **Severity:** High
- **Steps to Reproduce:**
  1. Run `test_delete_api.ps1 -TestExtension "' OR 1=1; --" -TestUniqueKey "99999"`
  2. Expected: Input is validated/sanitized before being passed to api2hipath.exe
  3. Actual: Raw input is interpolated directly into the `-w` WHERE clause argument (lines 215, 219, 356)
- **Priority:** Fix before deployment

#### BUG-4: TestUniqueKey Not Validated as Numeric
- **Severity:** Medium
- **Steps to Reproduce:**
  1. Run `test_delete_api.ps1 -TestUniqueKey "abc"`
  2. Expected: Script rejects non-numeric unique_key
  3. Actual: Script passes "abc" directly to the API without validation
- **Priority:** Fix before deployment

#### BUG-5: Spec Status Inconsistency (RESOLVED)
- **Severity:** Low
- **Steps to Reproduce:**
  1. Read INDEX.md: status now says "In Review"
  2. Read OS4K-3 spec header: status now says "In Review"
  3. Status is now consistent.
- **Priority:** Fixed during QA

#### BUG-6: No CSV Template File in Repository
- **Severity:** Medium
- **Steps to Reproduce:**
  1. Spec AC says "CSV-Format ist dokumentiert und mit Beispiel-Template vorhanden"
  2. Search repository for CSV template files related to deletion
  3. Expected: A template CSV file exists
  4. Actual: No template file found
- **Priority:** Fix before deployment

#### BUG-7: Temporary Files Not Cleaned Up
- **Severity:** Low
- **Steps to Reproduce:**
  1. Run test_delete_api.ps1
  2. After completion, check OutputPath
  3. Expected: Temporary input/output files are cleaned up or documented as artifacts
  4. Actual: Files like `test_delete_p4_input_port_direct.txt`, `test_delete_p1_port_select.csv`, `test_delete_p5_huntgrp_type.csv`, `test_delete_p6_action_ctrl.csv`, and `test_delete_p7_action_log_*.csv` remain on disk
- **Priority:** Fix in next sprint

#### BUG-8: ORDER_ID Format Mismatch Between Spec and Implementation (UPDATED)
- **Severity:** Medium
- **Steps to Reproduce:**
  1. Spec defines ORDER_ID format: `OS4K[000001-999999]` (sequential, max 12 chars)
  2. test_delete_api.ps1 uses New-OrderID function generating timestamp-based IDs (e.g., `OS4K10143215`)
  3. Expected: Implementation matches spec format
  4. Actual: Implementation uses timestamp-based format. While this is within 12-char limit and arguably better for avoiding collisions, it diverges from the spec. The spec should be updated to match, or vice versa.
- **Priority:** Clarify in spec before deployment

#### BUG-9: ORDER_ID Collision Risk
- **Severity:** Medium
- **Steps to Reproduce:**
  1. Two admins run the script within the same 10ms window on the same day
  2. New-OrderID generates DDHHMMSSMM then truncates to 8 digits
  3. Expected: Unique ORDER_IDs for concurrent executions
  4. Actual: Possible collision. The truncation at line 98 (`$allDigits.Substring(0, 8)`) discards the millisecond component, reducing uniqueness to 1-second granularity.
- **Priority:** Fix before deployment

#### BUG-10: No Access Control Requirements in Spec
- **Severity:** High
- **Steps to Reproduce:**
  1. Review spec for access control / authorization requirements
  2. Expected: Spec defines who is authorized to run mass deletions
  3. Actual: Any user with API credentials can delete any number of NSts without additional authorization checks
- **Priority:** Fix before deployment

#### BUG-11 (NEW): Test Script Skips Mandatory Dependency Removal
- **Severity:** High
- **Steps to Reproduce:**
  1. Review test_delete_api.ps1 -- Phases 2, 3, 5a, 5b, 5c were removed in commit f8d5dbf
  2. The script now performs DELETE PORT directly (Phase 4) without first removing PICKUP_SUB or HUNTGRP_SUB entries
  3. Expected: Script demonstrates the spec's CRITICAL deletion order (PICKUP_SUB -> HUNTGRP_SUB -> PORT)
  4. Actual: Script relies on automatic cascade deletion, which the spec explicitly warns will fail with error -56 for NSts that have group memberships
  5. Note: If testing showed that the API DOES cascade-delete dependencies, the spec's CRITICAL warning is incorrect and should be updated
- **Priority:** Clarify API behavior and update either the spec or the script

#### BUG-12 (NEW): Non-Zero Exit Code Incorrectly Assumed to Mean Master SA
- **Severity:** Medium
- **Steps to Reproduce:**
  1. Review Phase 4 error handling (lines 305-313)
  2. Any non-zero exit code triggers the message "Koennte Master SA sein - Phase 5 prueft..."
  3. Expected: Different exit codes are handled differently (network error, auth failure, permission denied, etc.)
  4. Actual: All non-zero codes are treated as potential Master SA, which could mask real errors like authentication failures or network timeouts
- **Priority:** Fix before deployment

---

### Regression Impact on Deployed Features

#### OS4K-1 (ETL-Pipeline) -- Deployed
- **Risk:** LOW. OS4K-3 is a separate script (`delete_nst.ps1` / `test_delete_api.ps1`), not modifying `os4k2excel.ps1`. However, after NSt deletions, the ETL pipeline should reflect the changes in the next export. No code changes to OS4K-1 are needed.
- **Concern:** The spec mentions "Reuse von OS4K-1" for API call patterns and logging functions, but does not specify whether functions will be imported or copy-pasted. If copied, future OS4K-1 changes will not propagate.
- **Verified:** `os4k2excel.ps1` was modified in recent commits (version bump only, lines 3 and 90). No functional changes to the ETL pipeline.

#### OS4K-2 (Webserver) -- Deployed
- **Risk:** LOW. OS4K-3 is CLI-only and does not interact with the web server.

---

### Summary
- **Acceptance Criteria:** 0/35 passed (0 fully implemented and testable)
- **Partial/Demonstrated in test script:** 5 criteria have working proof-of-concept in test_delete_api.ps1 (ACTION_CONTROL polling, API-only access, basic logging, NSt existence check, ORDER_ID generation)
- **Bugs Found:** 12 total (1 Critical, 4 High, 4 Medium, 3 Low)
- **Security Issues:** 7 findings (2 High, 3 Medium, 3 Low -- some overlap with bug list)
- **Production Ready:** NO
- **Recommendation:** The feature has NOT been implemented. Only a validation/test script exists. The spec is thorough and well-structured, but the following must happen before deployment:
  1. Implement `delete_nst.ps1` per the tech design
  2. CLARIFY: Does the OS4K API cascade-delete dependencies when DELETE PORT is called? If yes, update the spec. If no, the test script's approach (BUG-11) is wrong.
  3. Fix credential logging (BUG-2) in both test script and future implementation
  4. Add input sanitization (BUG-3, BUG-4)
  5. Add access control requirements to spec (BUG-10)
  6. Create CSV template file (BUG-6)
  7. Fix ORDER_ID collision risk (BUG-9) and align format with spec (BUG-8)
  8. Fix non-zero exit code handling (BUG-12)
  9. Re-run QA after implementation



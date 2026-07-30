# Aufruf und Parameter

## Aufruf

```powershell
# Interaktiv (User und Passwort werden abgefragt)
.\os4k2excel.ps1 -ApiHost "<IP>"

# Nur Passwort interaktiv (User per Parameter)
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<BENUTZER>"

# Vollstaendig non-interaktiv (alle Parameter angegeben)
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<BENUTZER>" -ApiPassword "<PASSWORT>"

# Mit Ausgabeverzeichnis
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<BENUTZER>" -ApiPassword "<PASSWORT>" -OutputPath "C:\export"

# Mit Duplikats-Markierung und SIP-Secrets im Klartext
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<BENUTZER>" -ApiPassword "<PASSWORT>" -MarkDuplicate -ShowSecrets

# Mit PEN/HVT-Daten (Hauptverteiler-Zuordnung)
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<BENUTZER>" -ApiPassword "<PASSWORT>" -IncludePenData

# Mit Debug-Modus fuer detaillierte Diagnose
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<BENUTZER>" -ApiPassword "<PASSWORT>" -Debug
```

## Parameter

### Pflichtparameter

| Parameter | Beschreibung |
|---|---|
| `-ApiHost` | IP-Adresse des OpenScape 4000 (muss immer angegeben werden) |
| `-ApiUser` | API-Benutzername — wird interaktiv abgefragt wenn nicht angegeben |
| `-ApiPassword` | API-Passwort — wird interaktiv mit maskierter Eingabe abgefragt wenn nicht angegeben |

> **Hinweis (ab M29):** Nur `-ApiHost` muss zwingend als Parameter angegeben werden. `-ApiUser` und `-ApiPassword` werden interaktiv abgefragt, wenn sie fehlen. Das Passwort wird dabei maskiert (keine Zeichen auf dem Bildschirm). In nicht-interaktiven Umgebungen (z.B. Scheduled Tasks) muessen alle drei Parameter angegeben werden.

### Optionale Parameter

| Parameter | Standard | Beschreibung |
|---|---|---|
| `-ApiPath` | `C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe` | Pfad zu api2hipath.exe |
| `-OutputPath` | Skriptverzeichnis | Zielverzeichnis fuer alle Ausgabedateien |
| `-MarkDuplicate` | aus | Doppelte PENs im Gesamt-Tab orange markieren |
| `-ShowSecrets` | aus | SIP-Secrets im Klartext exportieren (Standard: maskiert als `***`) |
| `-IncludePenData` | aus | PEN/HVT-Daten abfragen und exportieren (Hauptverteiler-Zuordnung). Erzeugt zusaetzliches HVT-Sheet und HVT-Spalten in Standort-/Gesamt-Sheets |
| `-Debug` | aus | Detaillierte Diagnose-Ausgaben ins Log schreiben (CSV-Groessen, Hashtable-Eintraege, API Exit-Codes) |

## Ausgabedateien

Alle Dateien werden im `OutputPath` mit Datumsstempel erzeugt:

| Datei | Beschreibung |
|---|---|
| `OS4K-PORT-YYYY-MM-DD.xlsx` | Hauptreport (Excel) |
| `OS4K-PORT-YYYY-MM-DD.log` | Ausfuehrungsprotokoll |
| `OS4K-{TABELLE}-YYYY-MM-DD.csv` | Zwischendateien der API-Abfragen |

## Interaktive Credential-Abfrage (ab M29)

Wenn `-ApiUser` oder `-ApiPassword` nicht angegeben werden, fragt das Script diese interaktiv ab:

- **ApiUser:** Sichtbare Eingabe (Klartext)
- **ApiPassword:** Maskierte Eingabe (keine Zeichen auf dem Bildschirm, `Read-Host -AsSecureString`)
- **Leere Eingabe:** Fehlermeldung und Abbruch
- **Scheduled Tasks:** Nicht-interaktive Umgebungen werden erkannt — sauberer Abbruch mit Fehlermeldung statt haengendem Prompt

## Hilfe bei fehlenden Parametern

Wird das Skript ohne `-ApiHost` aufgerufen, zeigt es automatisch eine Hilfe mit allen verfuegbaren Parametern und einem Beispielaufruf an.

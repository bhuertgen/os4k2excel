# OS4K-5: Interaktive Credential-Abfrage

## Status: Deployed
**Created:** 2026-03-24
**Last Updated:** 2026-03-24

## Dependencies
- Requires: OS4K-1 (ETL-Pipeline) — Parameterblock und Parameterprüfung

## User Stories
- As a Telefonadministrator, I want to das Script ohne -ApiPassword aufrufen können, so that mein Passwort nicht in der Kommandozeilen-History oder Bildschirmausgabe sichtbar ist.
- As a Telefonadministrator, I want to das Script ohne -ApiUser aufrufen können, so that ich den Benutzernamen interaktiv eingeben kann.
- As a externer Techniker, I want to Credentials interaktiv eingeben, so that ich keine sensiblen Daten in Batch-Dateien oder Shortcuts speichern muss.
- As a Sicherheitsverantwortlicher, I want to dass Passwörter bei der Eingabe maskiert werden, so that keine Credentials auf dem Bildschirm lesbar sind.

## Acceptance Criteria
- [x] Wenn `-ApiPassword` leer ist oder fehlt: Passwort interaktiv abfragen mit maskierter Eingabe (keine Zeichen auf dem Bildschirm)
- [x] Wenn `-ApiUser` leer ist oder fehlt: Benutzername interaktiv abfragen mit sichtbarer Eingabe
- [x] Wenn `-ApiHost` leer ist oder fehlt: weiterhin sofort Fehler + Abbruch (keine interaktive Abfrage)
- [x] Wenn bei der interaktiven Abfrage eine leere Eingabe erfolgt (nur Enter): Fehlermeldung anzeigen und Script beenden (exit 1)
- [x] Wenn `-ApiPassword` und/oder `-ApiUser` per Parameter übergeben werden: kein Prompt, bisheriges Verhalten bleibt erhalten
- [x] Maskierte Passwort-Eingabe nutzt PowerShell `Read-Host -AsSecureString` oder `$Host.UI.ReadLineAsSecureString()` und konvertiert zurück zu Klartext für die API-Nutzung
- [x] Das Script funktioniert weiterhin vollständig non-interaktiv wenn alle drei Pflichtparameter angegeben sind

## Edge Cases
- Benutzer gibt `-ApiPassword ""` (explizit leerer String) an → interaktive Abfrage auslösen
- Benutzer gibt `-ApiPassword " "` (nur Leerzeichen) an → interaktive Abfrage auslösen (IsNullOrWhiteSpace)
- Script wird in einer nicht-interaktiven Umgebung ausgeführt (z.B. Scheduled Task) ohne Parameter → Fehlermeldung und Abbruch, kein hängender Prompt
- SecureString-Konvertierung muss den vollen Unicode-Zeichensatz unterstützen (Sonderzeichen in Passwörtern)
- `-ApiUser` wird per Parameter übergeben, aber `-ApiPassword` nicht → nur Passwort abfragen

## Technical Requirements (optional)
- PowerShell 5.1 Kompatibilität (kein PowerShell 7-only Feature)
- `Read-Host -AsSecureString` für maskierte Passwort-Eingabe
- `Read-Host` (ohne -AsSecureString) für Benutzername-Eingabe
- Prüfung auf nicht-interaktive Umgebung via `[Environment]::UserInteractive` oder `$Host.Name`

---
<!-- Sections below are added by subsequent skills -->

## Tech Design (Solution Architect)

### Komponenten-Struktur (Ablauf-Logik)

```
os4k2excel.ps1 — Startsequenz
+-- [1] Parameterblock (bereits vorhanden)
|       -ApiHost, -ApiUser, -ApiPassword
|
+-- [2] NEU: Credential-Validierungsblock
|   +-- ApiHost-Prüfung
|   |   └── leer → Fehler + Abbruch (wie bisher)
|   |
|   +-- Umgebungs-Detektion
|   |   └── Nicht-interaktive Umgebung erkannt?
|   |       └── fehlende Credentials → Fehler + Abbruch (kein hängender Prompt)
|   |
|   +-- ApiUser-Prüfung
|   |   ├── angegeben → weiter (kein Prompt)
|   |   └── leer/fehlt → Interaktiver Prompt (sichtbare Eingabe)
|   |       └── leere Eingabe → Fehler + Abbruch
|   |
|   +-- ApiPassword-Prüfung
|       ├── angegeben → weiter (kein Prompt)
|       └── leer/fehlt → Maskierter Prompt (kein Echo auf Bildschirm)
|           └── leere Eingabe → Fehler + Abbruch
|
+-- [3] Bestehende ETL-Pipeline (unverändert)
        api2hipath → Transform → Excel-Export
```

### Datenmodell

Keine neuen Dateien oder Datenstrukturen. Credentials fließen als Strings wie bisher.

| Feld | Quelle | Typ | Anmerkung |
|---|---|---|---|
| `$ApiHost` | Nur Parameter | String | Pflicht, kein Fallback |
| `$ApiUser` | Parameter oder interaktiver Prompt | String | Sichtbare Eingabe |
| `$ApiPassword` | Parameter oder maskierter Prompt | String (Klartext) | Intern kurz als SecureString, sofort konvertiert |

### Tech-Entscheidungen

| Entscheidung | Warum |
|---|---|
| `Read-Host -AsSecureString` für Passwort | PS-Standard für maskierte Eingabe; PS 5.1 kompatibel |
| Sofortige Konvertierung SecureString → String | `api2hipath.exe` erwartet Klartext; wird nicht persistent gespeichert |
| `[Environment]::UserInteractive` für Umgebungsprüfung | Erkennt Scheduled Tasks und nicht-interaktive Shells zuverlässig |
| `[string]::IsNullOrWhiteSpace()` für Leer-Prüfung | Behandelt `""`, `" "` und `$null` gleich (Edge Cases) |
| Kein neues Modul / keine neue Datei | Single-File-Prinzip bleibt erhalten |

### Änderungs-Umfang

- **Geänderte Datei:** `os4k2excel.ps1` — Bereich nach Parameterblock, vor ETL-Start
- **Keine neuen Dateien, keine neuen Module**
- **Kein Einfluss auf bestehende Logik** wenn alle drei Parameter übergeben werden

## QA Test Results

**Tested:** 2026-03-24
**Script:** os4k2excel.ps1 (M28.20260324.1400)
**Tester:** QA Engineer (AI) -- Code Review + Static Analysis
**Method:** Code review of lines 54-139 (credential block) plus downstream usage (lines 226, 438)

### Acceptance Criteria Status

#### AC-1: ApiPassword leer/fehlt -> maskierte interaktive Abfrage
- [x] PASS: Line 119 checks `[string]::IsNullOrWhiteSpace($ApiPassword)` which covers missing, empty, and whitespace-only values
- [x] PASS: Line 128 uses `Read-Host "ApiPassword" -AsSecureString` for masked input
- [x] PASS: Lines 129-134 convert SecureString to plaintext via Marshal with proper BSTR cleanup in finally block

#### AC-2: ApiUser leer/fehlt -> sichtbare interaktive Abfrage
- [x] PASS: Line 102 checks `[string]::IsNullOrWhiteSpace($ApiUser)` which covers missing, empty, and whitespace-only values
- [x] PASS: Line 111 uses `Read-Host "ApiUser"` (no -AsSecureString) for visible input

#### AC-3: ApiHost leer/fehlt -> sofort Fehler + Abbruch (keine interaktive Abfrage)
- [x] PASS: Lines 72-96 check ApiHost first, display usage info, and `exit 1` without any prompt

#### AC-4: Leere Eingabe bei interaktiver Abfrage -> Fehlermeldung + exit 1
- [x] PASS: Line 112-115 checks `IsNullOrWhiteSpace($ApiUser)` after prompt and exits with error
- [x] PASS: Line 135-138 checks `IsNullOrWhiteSpace($ApiPassword)` after conversion and exits with error

#### AC-5: Parameter per Kommandozeile -> kein Prompt, bisheriges Verhalten
- [x] PASS: Both `if` blocks (line 102, 119) only trigger when the value IsNullOrWhiteSpace; if provided, they are skipped entirely

#### AC-6: Maskierte Eingabe nutzt Read-Host -AsSecureString + Konvertierung
- [x] PASS: Line 128 uses `Read-Host -AsSecureString`
- [x] PASS: Lines 129-134 use `SecureStringToBSTR` + `PtrToStringAuto` + `ZeroFreeBSTR` in finally block

#### AC-7: Vollstaendig non-interaktiv wenn alle drei Parameter angegeben
- [x] PASS: All three checks use `IsNullOrWhiteSpace` -- if all parameters have values, no prompts are triggered and the script proceeds directly to version output (line 142) and ETL pipeline

### Edge Cases Status

#### EC-1: `-ApiPassword ""` (explizit leerer String) -> interaktive Abfrage
- [x] PASS: `[string]::IsNullOrWhiteSpace("")` returns `$true`, triggers the prompt block (line 119)

#### EC-2: `-ApiPassword " "` (nur Leerzeichen) -> interaktive Abfrage
- [x] PASS: `[string]::IsNullOrWhiteSpace(" ")` returns `$true`, triggers the prompt block (line 119)

#### EC-3: Nicht-interaktive Umgebung ohne Parameter -> Fehler, kein haengender Prompt
- [x] PASS: Line 99 detects non-interactive sessions via `[Environment]::UserInteractive` combined with `$Host.Name` check
- [x] PASS: Lines 103-108 and 120-125 exit immediately with error message if not interactive and credentials missing

#### EC-4: SecureString Unicode-Unterstuetzung (Sonderzeichen in Passwoertern)
- [ ] BUG-1: `PtrToStringAuto` may not correctly decode BSTR on all platforms (see BUG-1 below)

#### EC-5: `-ApiUser` per Parameter, aber `-ApiPassword` nicht -> nur Passwort abfragen
- [x] PASS: The two checks are independent (`if` blocks at lines 102 and 119), so if ApiUser is provided but ApiPassword is not, only the password prompt appears. The password prompt message on line 127 correctly shows `$ApiUser@$ApiHost` confirming the user was already set.

### Security Audit Results

- [x] Password masking: Read-Host -AsSecureString masks console input correctly
- [x] BSTR memory cleanup: `ZeroFreeBSTR` is called in a `finally` block (line 133), ensuring the unmanaged BSTR is zeroed even if an exception occurs
- [x] No password in log file: The log file (lines 189ff) does not write ApiPassword anywhere
- [x] No credential leakage in error messages: Error messages reference parameter names, not values
- [ ] BUG-2: Password visible in process arguments (see BUG-2 below)
- [ ] BUG-3: Password visible in PowerShell command history when passed as parameter (see BUG-3 below)

### Regression Check

- [x] OS4K-1 (ETL-Pipeline): Parameter block (line 54-64) unchanged in structure; existing params retain default values; ETL pipeline code from line 141 onward is untouched by this feature
- [x] OS4K-4 (PENDATA): `-IncludePenData` switch still present (line 62), PEN table logic (lines 205-209) unaffected

### Bugs Found

#### BUG-1: PtrToStringAuto vs PtrToStringBSTR for SecureString conversion
- **Severity:** Low
- **Steps to Reproduce:**
  1. On Windows with PowerShell 5.1, use a password containing Unicode characters (e.g., umlauts, CJK characters)
  2. The `PtrToStringAuto` method (line 131) converts based on the platform character width, while `PtrToStringBSTR` would be the canonical counterpart to `SecureStringToBSTR`
  3. Expected: All Unicode passwords decode correctly
  4. Actual: On Windows (where BSTR is UTF-16LE and Auto is also UTF-16), this works identically to `PtrToStringBSTR`. The distinction only matters on non-Windows .NET Core where `Auto` might differ. Since the spec requires PS 5.1 (Windows-only), this is cosmetically incorrect but functionally safe.
- **Priority:** Nice to have -- consider changing to `PtrToStringBSTR` for correctness and forward-compatibility with PS 7 cross-platform

#### BUG-2: API password visible in process argument list
- **Severity:** Medium
- **Steps to Reproduce:**
  1. Run the script with interactive password prompt
  2. While `api2hipath.exe` is executing, open Task Manager or run `Get-Process api2hipath | Select-Object -Property CommandLine` from another PowerShell session
  3. Expected: Password should not be visible in the process command line
  4. Actual: Line 226 passes `-p "$ApiPassword"` as a command-line argument to `Start-Process`, making the password visible in the process list to any user on the system
- **Note:** This is a pre-existing issue (not introduced by OS4K-5) and is inherent to how `api2hipath.exe` works (it requires `-p` as a CLI argument). The feature spec's goal of hiding the password from "Kommandozeilen-History" and "Bildschirmausgabe" is met. However, the process argument list is a separate attack surface.
- **Priority:** Fix in next sprint -- investigate if api2hipath.exe supports reading credentials from stdin or environment variables

#### BUG-3: Password remains in $ApiPassword variable for script lifetime
- **Severity:** Low
- **Steps to Reproduce:**
  1. Run script with interactive password prompt
  2. If script hits an error and drops to a debugger, or if running in ISE, the `$ApiPassword` variable holds the plaintext password in memory for the entire script duration
  3. Expected: Password should be cleared from memory as soon as possible after use
  4. Actual: `$ApiPassword` remains as a .NET string in memory until garbage collection
- **Note:** This is standard PowerShell behavior and cannot be fully mitigated without significant architectural changes. The BSTR is properly cleaned up (line 133), but the resulting string persists.
- **Priority:** Nice to have -- document as a known limitation

### Summary
- **Acceptance Criteria:** 7/7 passed
- **Edge Cases:** 4/5 passed (1 cosmetic issue with PtrToStringAuto)
- **Bugs Found:** 3 total (0 critical, 0 high, 1 medium, 2 low)
- **Security:** Password input masking works correctly. Medium-severity finding on password visibility in process arguments (pre-existing, not introduced by this feature).
- **Production Ready:** YES
- **Recommendation:** Deploy. BUG-2 (process argument exposure) is pre-existing and not a regression from OS4K-5. All acceptance criteria pass. The two low-severity items are cosmetic/hardening improvements for a future sprint.

## Deployment
**Deployed:** 2026-03-24
**Version:** M29.20260324
**Tag:** v29.0
**Script:** `os4k2excel.ps1` — Credential-Abfrage integriert in Hauptscript

---

## Nachtrag 2026-07-30: Aufrufform `-ApiPassword` ohne Wert (Issue #3)

AC-1 war nur teilweise erfuellt. Die maskierte Abfrage griff, wenn
`-ApiPassword` **ganz weggelassen** oder als leerer String uebergeben wurde —
nicht aber bei der naheliegendsten Schreibweise `-ApiPassword` ohne Wert:

```
.\os4k2excel.ps1 -ApiHost <IP> -ApiUser <USER> -ApiPassword -ShowSecrets
-> Fehlendes Argument fuer den Parameter "ApiPassword".
```

Ursache: `[string]$ApiPassword` verlangt zwingend einen Wert. PowerShell erkennt
das folgende `-ShowSecrets` als Parameternamen und bricht bereits beim
Parameter-Binding ab — das Skript startet nicht, der Credential-Block wird nie
erreicht und kann auch keine verstaendliche Meldung ausgeben.

**Loesung:** `-ApiPassword` ist jetzt ein `[switch]`, der Wert wird ueber den
positional gebundenen Parameter `$ApiPasswordWert` aufgenommen. Damit
funktionieren beide Formen.

Wichtig fuer kuenftige Aenderungen: `$ApiPasswordWert` steht bewusst an **erster
Stelle** im param-Block und traegt **kein** `[Parameter(Position=0)]`-Attribut.
Ein Parameter-Attribut wuerde das Skript zur Advanced Function machen, wodurch
der eigene `[switch]$Debug` mit dem gleichnamigen Common-Parameter kollidiert
("Ein Parameter mit dem Namen Debug wurde mehrfach definiert") und das Skript
ueberhaupt nicht mehr startet.

### Verifizierte Aufrufformen

| Aufruf | Ergebnis |
|---|---|
| `-ApiPassword -ShowSecrets ...` | maskierte Abfrage |
| `... -ShowSecrets -ApiPassword` (letzte Position) | maskierte Abfrage |
| `-ApiPassword <Wert>` | Wert wird verwendet, keine Abfrage |
| `-ApiPassword ""` | maskierte Abfrage |
| ganz weggelassen | maskierte Abfrage |
| Aufrufform des Webservers `-ApiPassword $var` | unveraendert funktionsfaehig |
| verirrtes Argument ohne `-ApiPassword` | sauber abgewiesen |

Zusaetzlich geprueft: `-Debug` bindet weiterhin, Syntaxpruefung fehlerfrei in
PS 5.1 und PS 7.6.

---

## Security-Nachtest 2026-07-30

**Getestet:** 2026-07-30
**Script:** os4k2excel.ps1 (M29.20260324.1807), unveraendert
**Methode:** Dynamische Tests gegen laufende Prozesse, Byte-Analyse echter
Kundenexporte und -logs, Auswertung der `api2hipath.exe`-Hilfe (V7.00.000).
**Ergebnis:** Keine Code-Aenderung; Befunde dokumentiert in
`Wiki/de/Sicherheit.md` und `Wiki/en/Security.md`.

### Status der drei Alt-Bugs

#### BUG-1 (PtrToStringAuto) — BEHOBEN
Zeile 156 verwendet inzwischen `PtrToStringBSTR`. Damit ist EC-4
(Unicode-Sonderzeichen im Passwort) erfuellt. Punkt geschlossen.

#### BUG-2 (Passwort in Prozess-Argumenten) — BESTAETIGT, NICHT BEHEBBAR
Die offene Frage der urspruenglichen QA lautete: *"investigate if api2hipath.exe
supports reading credentials from stdin or environment variables"*. Diese
Untersuchung ist jetzt durchgefuehrt.

**Antwort: nein.** Die Hilfe von `api2hipath.exe` V7.00.000 kennt ausschliesslich
`-p passwd`. Es existiert kein stdin-Modus, keine Umgebungsvariable und keine
Credential-Datei. Ein Fix im Sinne von "Passwort nicht mehr auf der
Kommandozeile" ist mit diesem Werkzeug nicht moeglich.

Zusaetzlich praktisch nachgewiesen: Die Kommandozeile ist aus einer **separaten
Sitzung ohne Administratorrechte** auslesbar
(`Get-CimInstance Win32_Process | Select-Object CommandLine`). Die urspruengliche
Einstufung "Medium" ist damit eher zu niedrig — der Bewertung liegt zugrunde,
dass ein Lauf mehrere Dutzend solcher Prozesse erzeugt und die Kommandozeile
zusaetzlich in Sysmon (Event 1) und Prozessauditing (Event 4688) landet.

**Kompensierende Massnahme statt Fix:** Nur-Lese-API-Konto verwenden, dessen
Passwort nach dem Einsatz gewechselt wird. Damit ist eine etwaige Exposition
wertlos. Dokumentiert in der Sicherheitsseite des Wikis.

#### BUG-3 (Passwort bleibt im Speicher) — unveraendert
Weiterhin eine Eigenschaft von PowerShell, als bekannte Einschraenkung
dokumentiert.

### Neue Befunde ausserhalb des Feature-Umfangs

| Befund | Einstufung | Ort |
|---|---|---|
| `-z` (verschluesselte Verbindung) wird nicht genutzt, obwohl von V7 unterstuetzt | hoch | alle drei Scripts |
| Weboberflaeche laeuft ueber HTTP, `WebPassword` im Klartext, Session-Cookie ohne `Secure` | mittel | `os4k2excel-server.ps1:881` |
| Logdateien von `test_delete_api.ps1` **vor** dem 10.03.2026 enthalten das Passwort im Klartext (Maskierung kam erst mit Commit 060b125) | hoch, aber Altbestand | Kundendatenablage |

### Bestaetigt unauffaellig

- `OS4K-PORT-*.log`: keine Zugangsdaten (an drei echten Logs geprueft)
- Maskierung in `test_delete_api.ps1` ist wirksam — in PS 5.1 **und** PS 7
  verifiziert (Scope-Zugriff auf `$ApiPassword` funktioniert)
- PowerShell-Historie: kein Klartext-Passwort
- `.gitignore`: deckt Kundendaten und Credential-Dateien vollstaendig ab

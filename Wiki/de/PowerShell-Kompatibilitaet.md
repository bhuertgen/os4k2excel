# PowerShell 5.1 und 7.x

os4k2excel läuft unter Windows PowerShell 5.1 und unter PowerShell 7.x. Diese
Seite hält fest, was geprüft wurde und an welchen drei Stellen sich die Ausgabe
zwischen den Versionen unterscheidet.

Grundlage ist ein Vergleichstest vom 30.07.2026 mit PowerShell 5.1.26100.8875
und PowerShell 7.6.4 auf demselben System (ANSI-Codepage 1252, Kultur de-DE).

## Ergebnis in einem Satz

Die Verarbeitung und die erzeugte Excel-Datei sind in beiden Versionen gleich.
Seit **M30.20260730.1628** verbleiben zwei rein kosmetische Unterschiede in der
Konsolen- und Logausgabe.

## Behoben in M30: Exit-Code wurde unter 5.1 nicht angezeigt

`os4k2excel.ps1` startet `api2hipath.exe` über `Start-Process -PassThru` und
liest anschließend `$process.ExitCode` aus. Ohne den Schalter `-Wait` lieferte
diese Eigenschaft unter PowerShell 5.1 **immer `$null`**:

| Tatsächlicher Exit-Code | PowerShell 5.1 (vorher) | PowerShell 7.6 |
|---|---|---|
| 0 | `$null` | 0 |
| 7 | `$null` | 7 |
| 13 | `$null` | 13 |

Sichtbar war das in der Statuszeile je Tabelle:

```
vorher (PS 5.1):  API-Abfrage für SWITCH abgeschlossen. Dauer: ... (Exit-Code: )
jetzt  (beide) :  API-Abfrage für SWITCH abgeschlossen. Dauer: ... (Exit-Code: 0)
```

Die Daten waren nie betroffen — der Ablauf hängt nicht am Exit-Code, sondern
daran, ob die erwartete CSV-Datei erzeugt wurde (`Test-Path`). Fehlgeschlagene
API-Aufrufe ließen sich aber nicht aus dem Log heraus diagnostizieren.

**Behoben** durch Ergänzen von `-Wait` an beiden `Start-Process`-Aufrufen
(GitHub-Issue #1). `test_delete_api.ps1` verwendete `-Wait` bereits und war nie
betroffen — die dortige Ablaufsteuerung der Löschvorgänge, die auf Exit-Codes
beruht, arbeitete in beiden Versionen korrekt.

## Unterschied 1: Umleitung der Ausgabe in eine Datei

Die Variable `$OutputEncoding` steuert, wie Text bei einer Umleitung kodiert
wird:

| | PowerShell 5.1 | PowerShell 7.6 |
|---|---|---|
| `$OutputEncoding` | `us-ascii` | `utf-8` |

Wird die Ausgabe umgeleitet, ersetzt PowerShell 5.1 alle Zeichen außerhalb des
ASCII-Bereichs durch Fragezeichen. Betroffen sind die Rahmenlinien der
Phasenüberschriften, das Häkchen der Abschlussmeldung und Umlaute:

```powershell
.\os4k2excel.ps1 -ApiHost 10.10.1.1 -ApiUser engr > ausgabe.txt
```

Unter 5.1 steht in `ausgabe.txt` dann `???` statt `═══`. Auf dem Bildschirm
selbst wird korrekt dargestellt — das Script setzt dafür
`[Console]::OutputEncoding` auf UTF-8.

**Abhilfe:** unter 5.1 statt der Umleitung `Tee-Object` oder
`Start-Transcript` verwenden, oder die Ausgabe unter 7.x umleiten.

## Unterschied 2: Byte-Order-Mark in der Logdatei

`Out-File -Encoding utf8` verhält sich unterschiedlich:

| | PowerShell 5.1 | PowerShell 7.6 |
|---|---|---|
| BOM in `OS4K-PORT-<Datum>.log` | ja (`EF BB BF`) | nein |

Der Inhalt ist identisch, die Datei unterscheidet sich in den ersten drei Bytes.
Relevant nur, wenn die Logdatei maschinell weiterverarbeitet wird.

## Geprüft und identisch

| Prüfpunkt | Ergebnis |
|---|---|
| Syntaxprüfung aller vier Scripts | fehlerfrei in beiden Versionen |
| `Import-Csv` bei ASCII- und UTF-8-Eingabe | identisch |
| Kultur und Datumsformat (`yyyy-MM-dd`) | identisch (de-DE) |
| Zahlenformatierung in Excel-Zellen | identisch |
| Sortierung der Standorte (`Sort-Object Domain, SwitchName`) | identisch, da nach zwei Feldern sortiert wird |
| Duplikaterkennung (`Group-Object switch_name, pen`) | identisch |
| Excel-Formatierung über ImportExcel/EPPlus | identisch |

### Zeichensatz der API-Exporte — behoben in M30

Hier verhielten sich beide Versionen **gleich** — allerdings gleich falsch. Es
war damit nie ein Versionsunterschied, sondern ein Fehler unabhängig von der
PowerShell-Version (GitHub-Issue #2).

`api2hipath.exe` schreibt seine Exporte in **Windows-1252**. `Import-Csv` ohne
`-Encoding` las sie als UTF-8; jedes Umlaut-Byte ist dort ungültig und wurde
durch `U+FFFD` ersetzt. Aus `Wagenmeisterbüro` wurde `Wagenmeisterb<U+FFFD>ro`.
Das ist verlustbehaftet — der ursprüngliche Buchstabe lässt sich aus der
Arbeitsmappe nicht zurückgewinnen.

Nachgewiesen an den Exporten vom 29.07.2026: 336 Umlaut-Bytes, ausschließlich
in den PORT-Tabellen, die sich exakt auf die deutschen Umlaute verteilen
(`0xFC` ü, `0xF6` ö, `0xDF` ß, `0xDC` Ü, `0xE4` ä, `0xC4` Ä, `0xD6` Ö). In der
Arbeitsmappe schlugen sie als 255 beschädigte Zellen in der Spalte
`displayname` durch.

**Behoben** durch die Hilfsfunktion `Import-ApiCsv`, die den Zeichensatz aus dem
Inhalt bestimmt: Was gültiges UTF-8 ist, wird als UTF-8 gelesen, alles andere
als Windows-1252. Damit bleiben reine ASCII-Tabellen unverändert und ein
künftiger Wechsel von `api2hipath.exe` auf UTF-8 würde automatisch richtig
behandelt.

Gegengeprüft an allen Exporten vom 29.07.2026: 336 Umlaute korrekt gelesen,
kein einziges `U+FFFD` mehr, Zeilenzahlen unverändert — in PowerShell 5.1 und
7.x mit identischem Ergebnis.

## Stolperstelle bei der Installation

Die beiden PowerShell-Versionen verwenden **getrennte Modulverzeichnisse**:

| Version | Pfad für ImportExcel |
|---|---|
| Windows PowerShell 5.1 | `%USERPROFILE%\Documents\WindowsPowerShell\Modules` |
| PowerShell 7.x | `%USERPROFILE%\Documents\PowerShell\Modules` |

Ein `Install-Module ImportExcel` in der einen Version installiert das Modul
**nicht** für die andere. Wer zwischen den Versionen wechselt, muss das Modul
zweimal installieren. Das Script prüft dies beim Start und gibt bei fehlendem
Modul beide Pfade aus.

## Empfehlung

Für den produktiven Einsatz ist **PowerShell 7.x** zu bevorzugen: der Exit-Code
wird korrekt ausgegeben und die Umleitung der Ausgabe in eine Datei bleibt
lesbar. Unter 5.1 ist das Ergebnis jedoch fachlich gleichwertig.

---

[Zurück zur Übersicht](Home.md)

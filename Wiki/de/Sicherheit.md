# Sicherheit

Diese Seite beschreibt, wie os4k2excel mit Zugangsdaten umgeht, welche Risiken
bleiben und wie ein Einsatz beim Kunden abgesichert wird. Die Angaben beruhen
auf einem Sicherheitstest vom 30.07.2026 gegen `api2hipath.exe` Version 7.00.000.

## Übersicht

| Thema | Bewertung |
|---|---|
| Passwort-Eingabe (maskiert) | abgesichert |
| Passwort in der Prozess-Kommandozeile | Restrisiko, technisch nicht behebbar |
| Verschlüsselte API-Verbindung (`-z`) | derzeit nicht genutzt |
| Passwort in Logdateien | abgesichert (seit 10.03.2026) |
| Passwort in der PowerShell-Historie | abgesichert |
| Weboberfläche | überträgt unverschlüsselt (HTTP) |

## Zugangsdaten eingeben

Wird kein Passwort übergeben, fragt das Script es interaktiv ab. Die Eingabe
erfolgt maskiert über `Read-Host -AsSecureString`, es erscheint kein Zeichen auf
dem Bildschirm. Beide Schreibweisen führen zur Abfrage:

```powershell
# -ApiPassword ohne Wert
.\os4k2excel.ps1 -ApiHost 10.10.1.1 -ApiUser engr -ApiPassword -ShowSecrets

# oder ganz weggelassen
.\os4k2excel.ps1 -ApiHost 10.10.1.1 -ApiUser engr
```

Auch `-ApiUser` wird abgefragt, wenn es fehlt — dort allerdings sichtbar.

`-ApiHost` wird bewusst **nicht** abgefragt, sondern führt bei Fehlen sofort zum
Abbruch mit Hinweistext.

In nicht-interaktiven Umgebungen (Aufgabenplanung, Dienst, CI) bricht das Script
mit einer Fehlermeldung ab, statt auf eine Eingabe zu warten, die niemand
beantworten kann.

Die Eingabe wird unmittelbar nach dem Auslesen aus dem unmanaged Speicher
entfernt (`ZeroFreeBSTR` in einem `finally`-Block). Der daraus entstandene
.NET-String bleibt bis zur Garbage Collection im Prozessspeicher — das ist eine
Eigenschaft von PowerShell und lässt sich ohne Architekturbruch nicht vermeiden.

## Restrisiko: Passwort in der Prozess-Kommandozeile

`api2hipath.exe` nimmt das Passwort **ausschließlich** über den
Kommandozeilenschalter `-p` entgegen. Die Hilfe des Tools (Version 7.00.000)
kennt weder eine Übergabe über die Standardeingabe noch über eine
Umgebungsvariable oder eine Credential-Datei.

Damit steht das Passwort für die Dauer jedes API-Aufrufs im Klartext in der
Kommandozeile des Prozesses. Ein Test hat bestätigt, dass diese Kommandozeile
**ohne Administratorrechte** aus einer anderen Sitzung ausgelesen werden kann:

```powershell
Get-CimInstance Win32_Process -Filter "Name='api2hipath.exe'" |
    Select-Object ProcessId, CommandLine
```

Ebenso sichtbar ist sie im Task-Manager über die einblendbare Spalte
„Befehlszeile" sowie in Sysmon (Event ID 1), in der Prozessüberwachung
(Event ID 4688) und in vielen EDR-Produkten.

Ein Lauf erzeugt einen Aufruf je Tabelle plus einen Aufruf je Standort — bei
einem mittleren System also mehrere Dutzend Prozesse.

**Das ist mit diesem API-Werkzeug nicht abstellbar.** Die maskierte Eingabe
schützt vor Mitlesen am Bildschirm und vor Spuren in der Kommandohistorie, nicht
vor dieser Prozessliste.

## Empfehlung für Kundentermine

Wenn beim Kunden gearbeitet wird und Dritte zusehen, hat sich folgendes Vorgehen
bewährt:

1. **Wegwerf-Konto verwenden.** Der Kunde legt vor dem Termin ein API-Konto mit
   ausschließlich lesenden Rechten an und ändert dessen Passwort nach dem
   Einsatz wieder. Dann ist es unerheblich, ob das Passwort zwischenzeitlich in
   einer Prozessliste oder einem Ereignisprotokoll aufgetaucht ist — es ist
   danach wertlos. Das ist die einzige Maßnahme, die das Restrisiko oben
   tatsächlich auflöst.
2. **Der Kunde gibt das Passwort selbst ein.** Das Script ohne `-ApiPassword`
   starten und die Tastatur übergeben. Das Passwort wird maskiert eingegeben und
   ist zu keinem Zeitpunkt sichtbar — auch nicht für den Techniker.
3. **Keine Zugangsdaten in Verknüpfungen oder Batch-Dateien hinterlegen.** Wer
   das Script wiederkehrend ausführt, sollte die Anmeldedaten nicht als
   Klartextparameter in einer geplanten Aufgabe speichern.

## Logdateien

Die von `os4k2excel.ps1` erzeugte Logdatei `OS4K-PORT-<Datum>.log` enthält keine
Zugangsdaten. Das wurde an echten Logs geprüft.

`test_delete_api.ps1` protokolliert die vollständige API-Kommandozeile und
maskiert das Passwort dabei als `***`. Diese Maskierung wurde am **10.03.2026**
eingeführt.

> **Achtung bei Altbeständen:** Logdateien von Läufen **vor** dem 10.03.2026
> enthalten das damals verwendete API-Passwort im Klartext. Solche Dateien
> sollten gelöscht und das betroffene Passwort als kompromittiert behandelt und
> gewechselt werden.

## Nicht verschlüsselte Verbindung

`api2hipath.exe` unterstützt ab Version 7 den Schalter `-z` für eine
verschlüsselte Verbindung zum Server. **os4k2excel setzt diesen Schalter derzeit
nicht.**

Damit laufen sowohl die Anmeldung als auch sämtliche abgerufenen Daten —
Rufnummern, Teilnehmernamen, IP-Adressen und SIP-Secrets — unverschlüsselt über
das Kundennetz. Wer in einer Umgebung arbeitet, in der das nicht hinnehmbar ist,
sollte den Export über ein abgesichertes Netzsegment fahren oder den Einsatz von
`-z` mit dem Anlagenbetreuer abstimmen.

## Weboberfläche

`os4k2excel-server.ps1` bindet einen `HttpListener` auf `http://+:8080/`. Daraus
ergibt sich:

- Das mit `-WebPassword` gesetzte Passwort wird unverschlüsselt übertragen.
- Das Sitzungs-Cookie trägt `HttpOnly`, aber kein `Secure`-Attribut.
- Der Passwortvergleich erfolgt zeichenweise und ist damit theoretisch für
  Laufzeitmessungen anfällig.

Die Weboberfläche eignet sich deshalb für ein vertrauenswürdiges internes Netz,
nicht für eine Veröffentlichung über Netzgrenzen hinweg.

## Ausgabedateien enthalten Kundendaten

Die erzeugten CSV-, Log- und Excel-Dateien enthalten Rufnummern,
Teilnehmernamen, IP-Adressen und Anlagenkennungen. Ohne den Schalter
`-ShowSecrets` wird `sip_secret` als `***` exportiert; **mit** dem Schalter steht
es im Klartext in der Arbeitsmappe.

Die `.gitignore` des Projekts schließt diese Dateitypen aus. Beim Weitergeben
von Berichten oder Fehlerunterlagen ist der Inhalt trotzdem vorher zu prüfen.

## Zusammenfassung der geprüften Punkte

| Prüfung | Ergebnis |
|---|---|
| Maskierte Passwort-Eingabe | wirksam |
| Speicherfreigabe nach der Eingabe (`ZeroFreeBSTR`) | vorhanden |
| Passwort in `OS4K-PORT-*.log` | nicht enthalten |
| Passwortmaskierung in `test_delete_*.log` | wirksam seit 10.03.2026 |
| Passwort in der PowerShell-Historie | nicht enthalten |
| Kommandozeile ohne Administratorrechte auslesbar | ja — Restrisiko |
| Alternative Passwortübergabe in `api2hipath.exe` | existiert nicht |
| `.gitignore` gegen Kundendaten | vollständig |

---

[Zurück zur Übersicht](Home.md)

# os4k2excel

**Export port, license and call-forwarding data from a Unify OpenScape 4000 / HiPath 4000 PBX to Excel.**

PowerShell ETL tool. It queries the OpenScape 4000 database through the XIE
Import/Export API (`api2hipath.exe`), joins the tables, calculates the license
split and writes a formatted multi-sheet Excel workbook. A web interface for
running the export from a browser is included.

Runs on Windows PowerShell 5.1 and PowerShell 7.x. Microsoft Excel is **not**
required — the workbook is written through the ImportExcel module (EPPlus).

*Deutschsprachige Kurzfassung weiter unten, vollständige deutsche Dokumentation
im [Wiki](Wiki/de/Home.md).*

---

## Installation

```powershell
irm https://raw.githubusercontent.com/bhuertgen/os4k2excel/main/install.ps1 | iex
```

Installs into `%LOCALAPPDATA%\Programs\os4k2excel`, adds that folder to the user
PATH and installs the [ImportExcel](https://www.powershellgallery.com/packages/ImportExcel/)
module if it is missing.

With options:

```powershell
$u = 'https://raw.githubusercontent.com/bhuertgen/os4k2excel/main/install.ps1'
& ([scriptblock]::Create((irm $u))) -InstallDir 'C:\Tools' -NoPath
```

| Option | Meaning |
|---|---|
| `-InstallDir <path>` | different target directory |
| `-Ref <branch\|tag>` | install a specific version (default `main`) |
| `-NoPath` | do not modify PATH |
| `-NoModule` | do not install ImportExcel |
| `-UpdateModule` | also update ImportExcel if a newer version exists |
| `-Uninstall` | remove scripts, directory and PATH entry |

### Updating

Run the installer again over an existing installation — that is the intended
update path. It overwrites the scripts, reports the version change
(`updated M29… -> M30…`) and does **not** create a duplicate PATH entry.

```powershell
# update to the current development state
irm https://raw.githubusercontent.com/bhuertgen/os4k2excel/main/install.ps1 | iex

# pin to a release instead — recommended for production
$u = 'https://raw.githubusercontent.com/bhuertgen/os4k2excel/main/install.ps1'
& ([scriptblock]::Create((irm $u))) -Ref v30.0
```

An already installed ImportExcel module is left untouched unless you pass
`-UpdateModule`.

### Installing without `irm | iex`

Many corporate machines restrict running downloaded scripts. The tool works
without the installer — it has no dependencies apart from ImportExcel.

```powershell
# single file
$url = 'https://raw.githubusercontent.com/bhuertgen/os4k2excel/main/os4k2excel.ps1'
Invoke-WebRequest -Uri $url -OutFile .\os4k2excel.ps1 -UseBasicParsing
Unblock-File .\os4k2excel.ps1

# or the whole repository
git clone https://github.com/bhuertgen/os4k2excel.git
```

If execution is blocked:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
# or, without a permanent change
powershell -ExecutionPolicy Bypass -File .\os4k2excel.ps1 -ApiHost <host> -ApiUser <user>
```

Full step-by-step guide: [Wiki — Installation (EN)](Wiki/en/Installation.md) ·
[Installation (DE)](Wiki/de/Installation.md)

---

## Requirements

| Component | Version | Note |
|---|---|---|
| Windows | 10 / 11 / Server 2016+ | |
| PowerShell | 5.1 **or** 7.x | both tested, see below |
| .NET Framework | 4.5+ | required by Windows PowerShell 5.1 and ImportExcel |
| ImportExcel | 7.x | `Install-Module ImportExcel -Scope CurrentUser` |

### Tested PowerShell versions

Script and installer are verified on both editions — same data, same workbook:

| Edition | Tested with | Status |
|---|---|---|
| Windows PowerShell 5.1 (Desktop) | 5.1.26100.8875 | supported |
| PowerShell 7.x (Core) | 7.6.4 | supported, recommended |

PowerShell 7 is the better choice for production: the exit code of
`api2hipath.exe` is reported correctly and redirected console output stays
readable. On 5.1 the result is functionally equivalent — details in
[PowerShell compatibility](Wiki/en/PowerShell-Compatibility.md).

> The two editions use **separate module directories**. `Install-Module
> ImportExcel` in one does not install it for the other; if you switch editions,
> install the module twice.

### api2hipath.exe

Command line tool by Unify/Mitel that accesses the OpenScape 4000 database
through the XIE (Import/Export) API.

| | |
|---|---|
| **Vendor** | Unify / Mitel |
| **Part of** | OpenScape 4000 Assistant / Manager V11 |
| **Default path** | `C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe` |
| **Availability** | installed together with the OpenScape 4000 Assistant/Manager, not available separately |
| **Documentation** | *OpenScape 4000 Assistant/Manager V11, Import/Export (XIE) API, Service Documentation* (P31003-H34B0-S102-02-0020, 08/2024) — Mitel Document Center: [EN](https://www.mitel.com/document-center/business-phone-systems/openscape-4000-ecosystem/openscape-4000/110/en/openscape-4000-assistantmanager-v11-importexport-xie-api-service-documentation) \| [DE](https://www.mitel.com/document-center/business-phone-systems/openscape-4000-ecosystem/openscape-4000/110/de/openscape-4000-assistantmanager-v11-importexport-xie-api-service-documentation) |

The tool must be installed on the machine running the script and needs network
access to the OpenScape 4000 server.

Offline installation of ImportExcel (machine without internet):

```powershell
# on a machine WITH internet
Save-Module -Name ImportExcel -Path C:\Transfer
# copy the ImportExcel folder to the target machine, into
#   %USERPROFILE%\Documents\WindowsPowerShell\Modules   (PowerShell 5.1)
#   %USERPROFILE%\Documents\PowerShell\Modules          (PowerShell 7.x)
```

---

## Usage

```powershell
# standard run
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<USER>" -ApiPassword "<PASSWORD>" -OutputPath "C:\script"

# highlight duplicate PENs in orange
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<USER>" -ApiPassword "<PASSWORD>" -MarkDuplicate
```

| Parameter | Required | Description |
|---|---|---|
| `-ApiHost` | yes | IP address of the OpenScape 4000 |
| `-ApiUser` | yes | API user name |
| `-ApiPassword` | yes | API password |
| `-ApiPath` | no | path to `api2hipath.exe` |
| `-OutputPath` | no | output directory (default: script directory) |
| `-MarkDuplicate` | no | mark duplicate PENs orange in the combined sheet |
| `-ShowSecrets` | no | export SIP secrets in clear text (default: masked as `***`) |
| `-IncludePenData` | no | query and export PEN/MDF data (main distribution frame) |
| `-Debug` | no | write detailed diagnostics to the log |

`-ApiPassword` may be given without a value, or left out entirely — either way the
script prompts for it, masked. `-ApiUser` is prompted for as well when missing:

```powershell
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<USER>" -ApiPassword -ShowSecrets
.\os4k2excel.ps1 -ApiHost "<IP>"
```

> **Working at a customer site?** The password is passed to `api2hipath.exe` on
> the command line, because the tool accepts it no other way. It is therefore
> readable in the process list while a query runs — no administrator rights
> needed. Let the customer type the password themselves, and use a read-only
> account whose password is rotated afterwards. See
> [Security](Wiki/en/Security.md) · [Sicherheit](Wiki/de/Sicherheit.md).

---

## How it works — the ETL pipeline

### 1. Extract

The script queries these tables through `api2hipath.exe`:

| Table | Content |
|---|---|
| `SWITCH` | available switches with their domain |
| `NUMBERING_PLAN` | numbering plan |
| `CFW` | call forwarding |
| `HUNTGRP` | hunt groups |
| `HUNTGRP_SERVICE` | hunt group services |
| `PICKUPGRP` | pickup groups |
| `PICKUP_SUB` | pickup group members |
| `PERSPORT` | port connection types |
| `DEVCONST` | device configurations |
| `PORT` | queried per site (domain/switch) |
| `PEN` | Physical Equipment Numbers with MDF assignment — only with `-IncludePenData` |

### 2. Site discovery

Domain/switch combinations are discovered **automatically** from the `SWITCH`
table (`domain` + `switch_name`). No manual site configuration is needed.

### 3. Transform

- joins all tables through hashtable lookups (domain + switch + key)
- maps codes to readable values (for example `841` → `CFU`, `i90` → `VOICE`)
- calculates the license split between Flex and TDM

### 4. Load

| Worksheet | Content |
|---|---|
| `{Domain}-{Switch}` | port data per site with license totals |
| `Gesamt` | all sites combined |
| `Lizenz Dashboard` | Flex vs. TDM licenses per site |
| `Sammelanschluss` | hunt groups with auto-filled display names |
| `NumberingPlan` | numbering plan |
| `Devconst` | device configurations |
| `HVT` | MDF assignment of all PENs (only with `-IncludePenData`) |

---

## License calculation

Base value per port:

| Condition | License value |
|---|---|
| PEN empty | `null` (no license) |
| device type RADIO or EXTLINE | `0` |
| device type BASEST | `4` |
| device type SET600 | `2` |
| all others | `1` |

Flex vs. TDM split:

| Connection type | Assigned to |
|---|---|
| IP2 (code `268`) | **Flex license** |
| everything else | **TDM license** |

---

## Output files

Written to `OutputPath`, stamped with the current date:

- `OS4K-PORT-YYYY-MM-DD.xlsx` — main report
- `OS4K-PORT-YYYY-MM-DD.log` — execution log
- `OS4K-{TABLE}-YYYY-MM-DD.csv` — intermediate files of the API queries

> These files contain extension numbers, subscriber names and system identifiers
> of a live PBX. The `.gitignore` of this repository excludes them; keep them out
> of any repository or shared location.

---

## Web interface

`os4k2excel-server.ps1` provides a web interface to start, monitor and download
the export from a browser.

```powershell
# minimal
.\os4k2excel-server.ps1 -ApiHost "<IP>" -ApiUser "<USER>" -ApiPassword "<PW>" -WebPassword "<WEBPW>"

# with e-mail delivery and a default recipient
.\os4k2excel-server.ps1 -ApiHost "<IP>" -ApiUser "<USER>" -ApiPassword "<PW>" -WebPassword "<WEBPW>" `
    -SmtpServer "mail.example.com" -SmtpFrom "os4k@example.com" -SmtpTo "admin@example.com"
```

Then open `http://servername:8080` in a browser.

| Parameter | Required | Description |
|---|---|---|
| `-ApiHost` | yes | IP address of the OpenScape 4000 |
| `-ApiUser` | yes | API user name |
| `-ApiPassword` | yes | API password |
| `-WebPassword` | yes | password for the web login |
| `-Port` | no | HTTP port (default 8080) |
| `-OutputPath` | no | output directory (default: script directory) |
| `-SmtpServer` | no | SMTP server for e-mail delivery |
| `-SmtpPort` | no | SMTP port (default 25) |
| `-SmtpFrom` | no | sender address |
| `-SmtpTo` | no | default recipient, pre-filled in the dashboard |

Features: password-protected login with session cookie (60 min timeout),
dashboard with connection info and progress display, export as a background job
(one run at a time), live progress polling every 2 seconds, direct download,
optional e-mail delivery of the workbook, scheduler by weekday and time, and a
history of recent runs.

**Stopping the server** — three ways: `Ctrl+C` in the terminal, the *Server
beenden* button in the dashboard, or creating the stop file from another
terminal:

```cmd
echo.> "C:\path\os4k2excel-server.stop"
```

**Port reservation without admin rights** — run once as administrator:

```powershell
netsh http add urlacl url=http://+:8080/ user=DOMAIN\USERNAME
```

---

## Kurzfassung (Deutsch)

PowerShell-Werkzeug, das Port-, Lizenz- und Rufumleitungsdaten aus einer Unify
OpenScape 4000 über die XIE-API (`api2hipath.exe`) ausliest und als formatierten
Excel-Report mit mehreren Arbeitsblättern exportiert. Die Standorte werden
automatisch aus der SWITCH-Tabelle erkannt, die Lizenzen nach Flex und TDM
getrennt ausgewiesen. Ein Webserver für den Aufruf über den Browser liegt bei.

```powershell
# Installation
irm https://raw.githubusercontent.com/bhuertgen/os4k2excel/main/install.ps1 | iex

# Aufruf — Passwort wird maskiert abgefragt
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<BENUTZER>"
```

Voraussetzungen: Windows PowerShell 5.1 oder PowerShell 7.x — beide getestet
(5.1.26100.8875 und 7.6.4), das Modul
ImportExcel (wird vom Installer eingerichtet) und `api2hipath.exe` aus dem
OpenScape 4000 Assistant/Manager. Microsoft Excel wird nicht benötigt.

> **Beim Kunden vor Ort:** `api2hipath.exe` nimmt das Passwort ausschließlich
> über die Kommandozeile entgegen. Es ist deshalb während einer laufenden
> Abfrage in der Prozessliste lesbar — ohne Administratorrechte. Empfehlung: den
> Kunden das Passwort selbst eingeben lassen und ein Konto mit reinen
> Leserechten verwenden, dessen Passwort danach gewechselt wird. Details unter
> [Sicherheit](Wiki/de/Sicherheit.md).

**Vollständige deutsche Dokumentation:** [Wiki (DE)](Wiki/de/Home.md) —
[Installation](Wiki/de/Installation.md) ·
[Aufruf und Parameter](Wiki/de/Aufruf-und-Parameter.md) ·
[ETL-Pipeline](Wiki/de/ETL-Pipeline.md) ·
[Lizenzberechnung](Wiki/de/Lizenzberechnung.md) ·
[Webserver](Wiki/de/Webserver.md) ·
[API-Parameter](Wiki/de/API-Parameter.md) ·
[Sicherheit](Wiki/de/Sicherheit.md) ·
[PowerShell 5.1 / 7.x](Wiki/de/PowerShell-Kompatibilitaet.md)

---

## Documentation

| | English | Deutsch |
|---|---|---|
| Overview | [Home](Wiki/en/Home.md) | [Home](Wiki/de/Home.md) |
| Installation | [Installation](Wiki/en/Installation.md) | [Installation](Wiki/de/Installation.md) |
| Usage and parameters | [Usage](Wiki/en/Usage-and-Parameters.md) | [Aufruf und Parameter](Wiki/de/Aufruf-und-Parameter.md) |
| ETL pipeline | [ETL pipeline](Wiki/en/ETL-Pipeline.md) | [ETL-Pipeline](Wiki/de/ETL-Pipeline.md) |
| License calculation | [License calculation](Wiki/en/License-Calculation.md) | [Lizenzberechnung](Wiki/de/Lizenzberechnung.md) |
| Web interface | [Web interface](Wiki/en/Web-Interface.md) | [Webserver](Wiki/de/Webserver.md) |
| API parameters | [API parameters](Wiki/en/API-Parameters.md) | [API-Parameter](Wiki/de/API-Parameter.md) |
| Security | [Security](Wiki/en/Security.md) | [Sicherheit](Wiki/de/Sicherheit.md) |
| PowerShell 5.1 / 7.x | [Compatibility](Wiki/en/PowerShell-Compatibility.md) | [Kompatibilität](Wiki/de/PowerShell-Kompatibilitaet.md) |
| Change history | [Change history](Wiki/en/Change-History.md) | [Change History](Wiki/de/Change-History.md) |

---

## Related tools

- [os4k-wabe2excel](https://github.com/bhuertgen/os4k-wabe2excel) — dial plan of an
  OpenScape 4000 to Excel: free, assigned and undefined extension numbers from
  AMO WABE
- [Estos-Location-Rules-Synchronizer](https://github.com/bhuertgen/Estos-Location-Rules-Synchronizer)
  — keeps the cross-site number resolution in Estos ProCall in sync with an
  OpenScape 4000

---

## Search terms

OpenScape 4000, HiPath 4000, Unify, Atos, Mitel, XIE API, api2hipath, Export
Table Client, port list, license report, license count, Flex license, TDM
license, call forwarding, hunt group, pickup group, numbering plan, PEN, MDF,
main distribution frame, PBX inventory, PBX audit, telephony reporting —
Rufumleitung, Sammelanschluss, Rufannahmegruppe, Rufnummernplan, Lizenzzählung,
Hauptverteiler, Portliste, TK-Anlage, Telefonanlage, Anlagendokumentation.

---

## License

MIT — see [LICENSE](LICENSE).

Not affiliated with Unify, Mitel or Atos. OpenScape and HiPath are trademarks of
their respective owners.

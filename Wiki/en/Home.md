# os4k2excel Wiki

Welcome to the wiki for **os4k2excel** — a PowerShell ETL tool that exports data
from a Unify/Mitel OpenScape 4000 PBX.

*Diese Seite auf Deutsch: [Home (DE)](../de/Home.md)*

## Quick start

os4k2excel extracts port, license and call-forwarding data from an OpenScape 4000
telephone system and writes it to a formatted Excel report.

**Version:** M29.20260324.1807

### Installation

```powershell
irm https://raw.githubusercontent.com/bhuertgen/os4k2excel/main/install.ps1 | iex
```

See **[Installation](Installation.md)** for the detailed walkthrough:

1. check the requirements
2. install the ImportExcel module (online or offline)
3. install api2hipath.exe / the Export Table Client
4. create an API user in the OS4K Manager
5. assign access rights (with screenshots)
6. run the script

**Straight to it:**

```powershell
# interactive - credentials are prompted for
.\os4k2excel.ps1 -ApiHost "192.0.2.10"

# or fully by parameter
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "secret"
```

## Main features

### Automatic extraction

- 10+ tables from the OpenScape 4000 (SWITCH, PORT, CFW, HUNTGRP, PICKUPGRP, …)
- no direct database access needed, only the XIE API through `api2hipath.exe`
- automatic domain/switch discovery, no manual site configuration

### License calculation

- Flex and TDM licenses calculated separately
- base values per device type (BASEST=4, SET600=2, …)
- dashboard with license totals per site

### Web interface

- `os4k2excel-server.ps1` serves a web interface
- start, monitor and download in the browser (port 8080)
- built-in scheduler for unattended runs
- e-mail delivery of the workbook

### Diagnostic mode

- `-Debug` writes detailed information to the log
- CSV sizes, hashtable entries, API exit codes
- useful for troubleshooting and performance analysis

## New in M29

- **Interactive credential prompt** — user and password are requested when not
  supplied, the password is masked (OS4K-5)
- **PEN/MDF assignment** — main distribution frame documentation through
  `-IncludePenData` (OS4K-4)
- **ImportExcel module check** with version information in the log
- **PtrToStringBSTR** for PowerShell 7 compatibility when converting SecureStrings

## Documentation

- **[Installation](Installation.md)** — requirements, ImportExcel (online/offline),
  api2hipath, creating the API user
- **[Usage and parameters](Usage-and-Parameters.md)** — calling the script,
  required and optional parameters, output files
- **[ETL pipeline](ETL-Pipeline.md)** — extract, transform, load, and the code mappings
- **[License calculation](License-Calculation.md)** — Flex vs. TDM, base values, dashboard
- **[API parameters](API-Parameters.md)** — all available API tables and functions
- **[Web interface](Web-Interface.md)** — configuration and operation
- **[Change history](Change-History.md)** — version history

## Frequently used commands

```powershell
# interactive - credentials are prompted for
.\os4k2excel.ps1 -ApiHost "192.0.2.10"

# simple export, non-interactive
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "secret"

# including MDF (main distribution frame) data
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "secret" -IncludePenData

# highlight duplicate PENs
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "secret" -MarkDuplicate

# web interface, browser: http://localhost:8080
.\os4k2excel-server.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "secret" -WebPassword "team2024"
```

## Requirements

- **Windows:** 10, 11, Server 2016+
- **PowerShell:** 5.1+
- **Module:** ImportExcel (free)
- **Software:** api2hipath.exe, from the OpenScape 4000 Manager/Assistant V11
- **Access:** XIE API credentials and network access to the OS4K server

## Output

All files carry the run date in their name:

- **OS4K-PORT-2026-03-10.xlsx** — main report, one sheet per site plus summaries
- **OS4K-PORT-2026-03-10.log** — execution log
- **OS4K-{TABLE}-2026-03-10.csv** — raw exports of the individual tables

> These files contain extension numbers, subscriber names and system identifiers
> of a live PBX. Keep them out of repositories and shared locations.

## Security

- no passwords stored in files
- interactive password entry is masked, nothing is echoed to the screen
- SIP secrets masked by default, clear text only with `-ShowSecrets`
- session-based authentication for the web interface
- run locally or inside trusted networks only

---

**Last change:** 2026-03-24 (version M29)
**Project readme:** [README.md](../../README.md)

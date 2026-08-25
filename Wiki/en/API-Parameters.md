# API parameters and features

*Diese Seite auf Deutsch: [API-Parameter](../de/API-Parameter.md)*

Reference of all parameters and the features introduced from version M26 onwards.

## os4k2excel.ps1 — main export script

### Required parameters

| Parameter | Type | Description |
|---|---|---|
| `-ApiHost` | String | IP address or host name of the OpenScape 4000 server |
| `-ApiUser` | String | user name for the XIE API (for example `engr`, `xieapi`) |
| `-ApiPassword` | String | password of the API user |

From M29 onwards only `-ApiHost` has to be supplied; user and password are
requested interactively when missing. See
[Usage and parameters](Usage-and-Parameters.md).

### Optional parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-ApiPath` | String | `C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe` | path to the `api2hipath.exe` executable |
| `-OutputPath` | String | script directory | target directory for the Excel, log and CSV files |
| `-MarkDuplicate` | Switch | — | highlight duplicate PENs in the combined sheet with an orange background |
| `-ShowSecrets` | Switch | — | export SIP secrets in clear text (default: masked as `***`) |
| `-IncludePenData` | Switch | — | query and export PEN/MDF data |
| `-Debug` | Switch | — | write detailed diagnostics to the log (M26+) |

### Examples

```powershell
# minimal
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "secret"

# with diagnostics
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "secret" -Debug

# with all options
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "secret" `
  -OutputPath "C:\Exports" -MarkDuplicate -ShowSecrets -Debug
```

## Features added in M26

### PICKUPGRP and PICKUP_SUB tables

Two additional tables are extracted automatically:

**PICKUPGRP** — pickup groups
Definition of groups that may answer calls.
Fields: `domain`, `switch_name`, `pickupgrpnum`, `displ`, `info`

**PICKUP_SUB** — pickup group members
Assignment of extensions to pickup groups.
Fields: `domain`, `switch_name`, `extension`, `pickupgrpnum`

The data is available internally for joins and future output sheets.

### Structured logging

The script writes phase headers to the log, which makes following the progress
and troubleshooting considerably easier:

```
════════════════════════════════════════════════════════════════
                    PHASE 1: INITIALISIERUNG
════════════════════════════════════════════════════════════════

════════════════════════════════════════════════════════════════
                  PHASE 2: API-ABFRAGEN (Batch)
════════════════════════════════════════════════════════════════
```

### Debug mode

`-Debug` adds detailed diagnostic information to the log:

- **CSV file size** — bytes and row count of every API query
- **hashtable sizes** — number of entries when the tables are loaded
- **API exit codes** — return codes of the `api2hipath.exe` calls
- **database state** — number of entries per table

```powershell
.\os4k2excel.ps1 -ApiHost "..." -ApiUser "..." -ApiPassword "..." -Debug
```

Useful for tracking down extraction problems, verifying that all expected data
was retrieved, and for performance analysis.

## Extracted tables

| Table | Description |
|---|---|
| **SWITCH** | switches with their domain |
| **NUMBERING_PLAN** | numbering plan |
| **CFW** | call forwarding |
| **HUNTGRP** | hunt groups |
| **HUNTGRP_SERVICE** | hunt group services |
| **PICKUPGRP** | pickup groups |
| **PICKUP_SUB** | pickup group members |
| **PERSPORT** | port connection types, department |
| **DEVCONST** | device configurations |
| **PORT** | port data, queried per domain/switch |
| **PEN** | Physical Equipment Numbers with MDF assignment, only with `-IncludePenData` |

## License calculation

Base value per port:

- PEN empty → `null` (no license)
- device RADIO or EXTLINE → `0`
- device BASEST → `4`
- device SET600 → `2`
- all others → `1`

Flex vs. TDM split:

- **IP2 connection** (code 268) → Flex license
- **everything else** → TDM license

Counted once per extension. Details: [License calculation](License-Calculation.md).

## Output files

All files carry a date stamp (`YYYY-MM-DD`):

- `OS4K-PORT-{date}.xlsx` — main report with several sheets
- `OS4K-PORT-{date}.log` — execution log
- `OS4K-{TABLE}-{date}.csv` — intermediate raw exports

## Web server parameters

`os4k2excel-server.ps1` takes the same API parameters as the main script:

```powershell
.\os4k2excel-server.ps1 `
  -ApiHost "192.0.2.10" `
  -ApiUser "engr" `
  -ApiPassword "secret" `
  -WebPassword "team2024" `
  [-SmtpServer "mail.example.com"] `
  [-SmtpFrom "os4k@example.com"] `
  [-SmtpTo "admin@example.com"]
```

The server runs the main script with all extractions and picks up the newer
features automatically. Details: [Web interface](Web-Interface.md).

---

**Version:** M26.20260310.1239 | **Release:** v2.26.0 | **Updated:** 2026-03-10

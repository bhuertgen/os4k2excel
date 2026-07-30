# Usage and parameters

*Diese Seite auf Deutsch: [Aufruf und Parameter](../de/Aufruf-und-Parameter.md)*

## Calling the script

```powershell
# interactive - user and password are prompted for
.\os4k2excel.ps1 -ApiHost "<IP>"

# password only interactive, user supplied
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<USER>"

# fully non-interactive
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<USER>" -ApiPassword "<PASSWORD>"

# with an output directory
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<USER>" -ApiPassword "<PASSWORD>" -OutputPath "C:\export"

# highlight duplicates and export SIP secrets in clear text
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<USER>" -ApiPassword "<PASSWORD>" -MarkDuplicate -ShowSecrets

# including PEN/MDF data (main distribution frame assignment)
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<USER>" -ApiPassword "<PASSWORD>" -IncludePenData

# with detailed diagnostics
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<USER>" -ApiPassword "<PASSWORD>" -Debug
```

## Parameters

### Required

| Parameter | Description |
|---|---|
| `-ApiHost` | IP address of the OpenScape 4000 — always required |
| `-ApiUser` | API user name — prompted for when not supplied |
| `-ApiPassword` | API password — prompted for with masked input when not supplied |

> **Note (from M29):** only `-ApiHost` has to be passed as a parameter.
> `-ApiUser` and `-ApiPassword` are requested interactively when missing, with
> the password masked. In non-interactive environments such as scheduled tasks
> all three parameters must be supplied.

### Optional

| Parameter | Default | Description |
|---|---|---|
| `-ApiPath` | `C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe` | path to api2hipath.exe |
| `-OutputPath` | script directory | target directory for all output files |
| `-MarkDuplicate` | off | mark duplicate PENs orange in the combined sheet |
| `-ShowSecrets` | off | export SIP secrets in clear text (default: masked as `***`) |
| `-IncludePenData` | off | query and export PEN/MDF data. Adds an extra MDF sheet and MDF columns to the site and combined sheets |
| `-Debug` | off | write detailed diagnostics to the log: CSV sizes, hashtable entries, API exit codes |

## Output files

Written to `OutputPath` with a date stamp:

| File | Description |
|---|---|
| `OS4K-PORT-YYYY-MM-DD.xlsx` | main report (Excel) |
| `OS4K-PORT-YYYY-MM-DD.log` | execution log |
| `OS4K-{TABLE}-YYYY-MM-DD.csv` | intermediate files of the API queries |

> These files contain live PBX data — extension numbers, subscriber names,
> system identifiers. Keep them out of repositories and shared locations.

## Interactive credential prompt (from M29)

When `-ApiUser` or `-ApiPassword` are omitted, the script asks for them:

- **ApiUser:** visible input
- **ApiPassword:** masked input, nothing is echoed (`Read-Host -AsSecureString`)
- **empty input:** error message and abort
- **scheduled tasks:** non-interactive environments are detected — the script
  aborts with a clear message instead of hanging on a prompt

## Help when parameters are missing

Called without `-ApiHost`, the script prints an overview of all available
parameters together with an example invocation.

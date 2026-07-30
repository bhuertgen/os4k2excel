# Change history

*Diese Seite auf Deutsch: [Change History](../de/Change-History.md)*

## Version M30.20260730 (milestone 30)

| Date | Change |
|---|---|
| 2026-07-30 | **Fix:** `-ApiPassword` without a value aborted with a PowerShell binding error instead of prompting (issue #3). The parameter now accepts both forms. |
| 2026-07-30 | **Fix:** umlauts in subscriber names were destroyed in the Excel export (issue #2). `api2hipath.exe` writes Windows-1252, `Import-Csv` read the files as UTF-8 — every umlaut became `U+FFFD`. Affected 255 cells per run. |
| 2026-07-30 | **New:** helper `Import-ApiCsv` detecting the character set from file content (valid UTF-8 stays UTF-8, otherwise Windows-1252). All 11 CSV imports converted. |
| 2026-07-30 | **Fix:** the exit code of `api2hipath.exe` was never evaluated on Windows PowerShell 5.1 (issue #1). `-Wait` added to both `Start-Process` calls. |
| 2026-07-30 | **New:** wiki pages [Security](Security.md) and [PowerShell 5.1 and 7.x](PowerShell-Compatibility.md) |
| 2026-07-30 | security test documented: the password in the process command line is readable without administrator rights and cannot be avoided with `api2hipath.exe` |

## Version M29.20260324 (milestone 29)

| Date | Change |
|---|---|
| 2026-03-24 | **New:** interactive credential prompt — password masked, user requested when missing (OS4K-5) |
| 2026-03-24 | non-interactive environments such as scheduled tasks are detected — clean abort instead of a hanging prompt |
| 2026-03-24 | ImportExcel module check at startup, version written to the log |
| 2026-03-24 | Fix: PtrToStringBSTR instead of PtrToStringAuto for PowerShell 7 compatibility |
| 2026-03-24 | OS4K-4 and OS4K-5 marked as deployed, OS4K-3 stays in review |

## Version M27.20260317 (milestone 27)

| Date | Change |
|---|---|
| 2026-03-17 | **New:** PEN/MDF assignment through `-IncludePenData` (OS4K-4) |
| 2026-03-17 | new HVT sheet listing all PENs — in use, free, reserved — with MDF connection data |
| 2026-03-17 | MDF columns in the site and combined sheets (Info, HVT1, HVT2, five fields each) |
| 2026-03-17 | complete board type mapping, 80+ types (documentation sections 10.3.19–10.3.24) |
| 2026-03-17 | PEN table instead of PENDATA — the correct ODF for both Manager and Assistant |
| 2026-03-17 | the PEN table returns lin1/lin2 directly, no separate PORT lookup needed |
| 2026-03-17 | DOMAIN table removed, it was not used functionally |
| 2026-03-17 | additional PEN fields: LTG, LTU, slot, circuit, PEN_Num, Board_Present, reservation time |

## Version M26.20260310 (milestone 26)

| Date | Change |
|---|---|
| 2026-03-10 | PICKUPGRP and PICKUP_SUB table extraction |
| 2026-03-10 | improved logging structure with phase headers |
| 2026-03-10 | new `-Debug` parameter for detailed diagnostics |
| 2026-03-10 | better error handling and API exit code checking |
| 2026-03-10 | security: API password masked in the logs |
| 2026-03-17 | wiki: online and offline installation guide for ImportExcel |
| 2026-03-17 | wiki: update and version management guide for ImportExcel |

## Version M22.20260215.1931 (milestone 22)

| Date | Change |
|---|---|
| 2026-02-15 | automatic site discovery from the SWITCH table |
| 2026-02-15 | DOMAIN/SWITCH tables added as API queries |
| 2026-02-15 | `-ApiUser` became a required parameter, replacing `-ApiLanguage` |
| 2026-02-15 | `-ShowSecrets` added for optional SIP secret display |
| 2026-02-15 | parameter validation with help text when required parameters are missing |
| 2026-02-15 | version and output path shown at startup |
| 2026-02-15 | fix: single-switch systems (PowerShell array handling) |
| 2026-02-15 | fix: empty tables (HuntgrpTable, CfwTable) are allowed to be empty |
| 2026-02-15 | fix: Excel file lock when running twice on the same day |
| 2026-02-15 | security: hard-coded credentials removed |
| 2026-02-15 | security: sip_secret masked by default in the export |
| 2026-02-15 | OutputPath now defaults to the script directory |
| 2026-02-15 | customer-specific references removed, file prefix changed to OS4K |

## Earlier versions

| Date | Change |
|---|---|
| 2025-12-05 | Flex and TDM licenses separated, dashboard extended |
| 2025-11-28 | hunt group auto-fill display name (master) with grey text |
| 2025-11-28 | fixes: license calculation, XML errors, colours |
| 2025-06-30 | initial version |

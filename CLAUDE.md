# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Version

Current: **M31.20260825.1532** — Format: `M<Meilenstein>.<Datum:yyyymmdd>.<Zeit:hhmm>`

## Project Overview

Single-file PowerShell ETL utility (`os4k2excel.ps1`) that extracts port/license data from a Unify OpenScape 4000 telecom system via `api2hipath.exe`, transforms it, and exports formatted Excel reports.

## Running the Script

```powershell
# Minimal (ApiHost, ApiUser und ApiPassword sind Pflicht)
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<USER>" -ApiPassword "<PASSWORD>"

# Mit allen Optionen
.\os4k2excel.ps1 -ApiHost "<IP>" -ApiUser "<USER>" -ApiPassword "<PASSWORD>" -OutputPath "C:\script" -MarkDuplicate -ShowSecrets -IncludePenData
```

**Dependencies**:
- `ImportExcel` PowerShell module — install via `Install-Module ImportExcel`, update via `Update-Module ImportExcel`. Provides `Export-Excel`, `Open-ExcelPackage`, `Close-ExcelPackage` and the OfficeOpenXml/.NET classes for Excel formatting. No Microsoft Office installation required.
- `api2hipath.exe` — Unify/Mitel XIE API command-line tool, shipped with OpenScape 4000 Assistant/Manager V11. Must be installed locally with network access to the OS4K server.

No build step, test framework, or linting tools are configured.

## Architecture

The script follows a linear ETL pipeline in one file:

1. **Extract** (~lines 78-103): Calls `api2hipath.exe` to query 9 tables (SWITCH, NUMBERING_PLAN, CFW, HUNTGRP, HUNTGRP_SERVICE, PICKUPGRP, PICKUP_SUB, PERSPORT, DEVCONST), plus optionally PEN when `-IncludePenData` is set, producing pipe-delimited CSVs.
2. **Domain/Switch discovery** (~lines 611-616): Extracts domain/switch pairs automatically from the SWITCH table (API) — no hardcoded site list.
3. **Transform** (`Process-PortData` function): Per domain/switch pair, queries PORT table (incl. lin1/lin2), joins all tables via hashtable lookups, maps coded values to readable names, calculates license values (Flex vs TDM split based on IP2 connection type). If `-IncludePenData`: joins PEN table via pen field for HVT columns.
4. **Load**: Exports to multi-sheet Excel workbook with per-location sheets, "Gesamt" (overall), "Lizenz Dashboard", "Sammelanschluss" (hunting groups), "NumberingPlan", "Devconst", and optionally "HVT" sheets. Applies formatting via OfficeOpenXml.

### Key data mappings (lines 118-124)

- **License values**: BASEST=4, SET600=2, RADIO/EXTLINE=0, others=1; split into Flex_Lizenz (IP2) vs TDM_Lizenz (all others)
- **Variant/Service/Device codes**: Mapped from numeric codes to readable names (e.g., 841->CFU, i90->VOICE)
- **Connection types**: 256->DIRECT, 257->PNT, 260->EXTERNAL, 263->LOG, 267->IP, 268->IP2
- **Card types** (PENDATA): 300->SLMA, 306->STMD, 360->STMI-HFA, 410->STMI, 411->STMI2-HFA
- **PEN status**: 002->Frei, 003->Belegt; **Reserved**: 000->nein, 001->ja

### Output artifacts (timestamped YYYY-MM-DD)

- `OS4K-PORT-{date}.xlsx` — Main Excel report
- `OS4K-PORT-{date}.log` — Execution log
- `OS4K-{TABLE}-{date}.csv` — Intermediate CSV files

## Conventions

- Script is UTF-8 with BOM
- German variable names and comments throughout
- Domain/Switch pairs are auto-discovered from the SWITCH table via API (no manual configuration needed)
- Excel formatting uses System.Drawing.Color and OfficeOpenXml for styling (bold headers, yellow/green backgrounds, orange for duplicates, grey for auto-filled names)

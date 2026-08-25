# ETL pipeline

*Diese Seite auf Deutsch: [ETL-Pipeline](../de/ETL-Pipeline.md)*

The script follows a linear ETL pipeline — extract, transform, load — in a single
file.

## 1. Extract

These tables are queried through `api2hipath.exe`:

| Table | Content |
|---|---|
| **SWITCH** | available switches with their domain |
| **NUMBERING_PLAN** | numbering plan |
| **CFW** | call forwarding |
| **HUNTGRP** | hunt groups |
| **HUNTGRP_SERVICE** | hunt group services |
| **PICKUPGRP** | pickup groups — **new in M26** |
| **PICKUP_SUB** | pickup group members — **new in M26** |
| **PERSPORT** | port connection types, department |
| **DEVCONST** | device configurations |

The **PORT** table is queried once per site (domain/switch).

Optional, with `-IncludePenData`:

| Table | Content |
|---|---|
| **PEN** | Physical Equipment Numbers with MDF assignment — **new in M27** |

## 2. Site discovery

Domain/switch combinations are discovered **automatically** from the **SWITCH**
table (fields `domain` and `switch_name`). No manual site configuration is
required.

## 3. Transform

- joins all tables through hashtable lookups (domain + switch + key)
- maps codes to readable values, for example `841` → `CFU`, `i90` → `VOICE`
- [license calculation](License-Calculation.md) with the Flex/TDM split

### Code mappings

| Category | Examples |
|---|---|
| **Variants** | j07=STATION, j08=SYSTEM, j09=STATIONV |
| **Services** | i90=VOICE, i91=AUDIO3K1, i92=FAXG23, j01=FAX |
| **Forwarding types (dtype)** | 841=CFU, 842=CFB, 843=CFNR, 844=CD |
| **Interaction type** | j03=EXTERN, j04=INTERN, j05=GEN |
| **Activation** | i70=off, C56=on |
| **Connection type** | 256=DIRECT, 257=PNT, 260=EXTERNAL, 263=LOG, 267=IP, 268=IP2 |
| **Board type** | 300=SLMA, 306=STMD, 317=SLMAVAR, 349=SLMOP, 360=STMI-HFA, 410=STMI, 411=STMI2-HFA, … (80+ types) — **new in M27** |
| **PEN status** | 002=free, 003=in use — **new in M27** |
| **Reserved** | 000=no, 001=yes — **new in M27** |

## 4. Load

The workbook contains these worksheets:

| Worksheet | Content |
|---|---|
| **{Domain}-{Switch}** | port data per site with license totals |
| **Gesamt** | all sites combined |
| **Lizenz Dashboard** | Flex vs. TDM licenses per site |
| **Sammelanschluss** | hunt groups with auto-filled display names |
| **NumberingPlan** | numbering plan |
| **Devconst** | device configurations |
| **HVT** | MDF assignment of all PENs, only with `-IncludePenData` — **new in M27** |

> The worksheet names are kept in German because existing reports, templates and
> downstream evaluations refer to them.

### Excel formatting

| Element | Colour |
|---|---|
| site total row | yellow |
| grand total row | green |
| dashboard header | blue with white text |
| dashboard grand total | gold |
| duplicate PENs (optional) | orange |
| auto-filled display names | grey |
| MDF 1 header (HVT sheet) | light blue |
| MDF 2 header (HVT sheet) | light green |

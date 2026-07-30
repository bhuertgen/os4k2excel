# License calculation

*Diese Seite auf Deutsch: [Lizenzberechnung](../de/Lizenzberechnung.md)*

## Base value per port

| Condition | License value |
|---|---|
| PEN empty | `null` (no license) |
| device type RADIO or EXTLINE | `0` |
| device type BASEST | `4` |
| device type SET600 | `2` |
| all others | `1` |

A license is counted **once per extension**. When an extension has several CFW
entries, only the first row carries the license value.

## Flex vs. TDM split

The split is based on the connection type from the PERSPORT table:

| Connection type | Code | Assigned to |
|---|---|---|
| IP2 | `268` | **Flex license** |
| all others (DIRECT, PNT, EXTERNAL, LOG, IP) | `256`–`267` | **TDM license** |

## License dashboard

The worksheet *Lizenz Dashboard* shows one row per site:

| Column | Description |
|---|---|
| **Standort** | domain/switch combination |
| **Genutzte Flex-Lizenzen** | total of all IP2 licenses |
| **Genutzte TDM-Lizenzen** | total of all non-IP2 licenses |
| **Gesamt Flex+TDM** | grand total |

> Column headings are kept in German because existing reports and downstream
> evaluations refer to them.

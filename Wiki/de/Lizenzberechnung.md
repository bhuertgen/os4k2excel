# Lizenzberechnung

## Basis-Wert je Port

| Bedingung | Lizenzwert |
|---|---|
| PEN leer | `null` (keine Lizenz) |
| Geraetetyp RADIO oder EXTLINE | `0` |
| Geraetetyp BASEST | `4` |
| Geraetetyp SET600 | `2` |
| Sonstige | `1` |

Die Lizenz wird nur **1x pro Nebenstelle** gezaehlt (bei mehreren CFW-Eintraegen nur in der ersten Zeile).

## Flex vs. TDM Aufteilung

Die Aufteilung basiert auf dem Verbindungstyp aus der PERSPORT-Tabelle:

| Verbindungstyp | Code | Zuordnung |
|---|---|---|
| IP2 | `268` | **Flex_Lizenz** |
| Alle anderen (DIRECT, PNT, EXTERNAL, LOG, IP) | `256-267` | **TDM_Lizenz** |

## Lizenz Dashboard

Das Excel-Arbeitsblatt "Lizenz Dashboard" zeigt pro Standort:

| Spalte | Beschreibung |
|---|---|
| **Standort** | Domain-Switch Kombination |
| **Genutzte Flex-Lizenzen** | Summe aller IP2-Lizenzen |
| **Genutzte TDM-Lizenzen** | Summe aller Nicht-IP2-Lizenzen |
| **Gesamt Flex+TDM** | Gesamtsumme |

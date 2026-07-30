# Features Index — os4k2excel

**Next Available ID:** OS4K-6

## Features

| ID | Feature | Priorität | Status | Spec |
|---|---|---|---|---|
| OS4K-1 | ETL-Pipeline (Extract, Transform, Excel-Export) | P0 | Deployed | [OS4K-1-etl-pipeline.md](OS4K-1-etl-pipeline.md) |
| OS4K-2 | Webserver (Browser-Oberfläche, Zeitplaner, E-Mail) | P1 | Deployed | [OS4K-2-webserver.md](OS4K-2-webserver.md) |
| OS4K-3 | Delete NSt via CSV Import with Dependency Management | P2 | In Review | [OS4K-3-delete-nst-via-csv.md](OS4K-3-delete-nst-via-csv.md) |
| OS4K-4 | PENDATA / HVT-Zuordnung (Hauptverteiler-Dokumentation) | P1 | Deployed | [OS4K-4-pendata-hvt-zuordnung.md](OS4K-4-pendata-hvt-zuordnung.md) |
| OS4K-5 | Interaktive Credential-Abfrage (Passwort maskiert) | P1 | Deployed | [OS4K-5-interactive-credential-prompt.md](OS4K-5-interactive-credential-prompt.md) |

## Build-Reihenfolge

1. **OS4K-1** — ETL-Pipeline (Kern; unabhängig)
2. **OS4K-2** — Webserver (hängt von OS4K-1 ab)
3. **OS4K-4** — PENDATA / HVT-Zuordnung (hängt von OS4K-1 ab)

## Legende

| Status | Bedeutung |
|---|---|
| Planned | Spezifiziert, noch nicht begonnen |
| In Progress | In aktiver Entwicklung |
| In Review | Fertig, wird getestet/geprüft |
| Deployed | Produktiv ausgerollt |

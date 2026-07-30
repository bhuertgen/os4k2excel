# PowerShell 5.1 and 7.x

os4k2excel runs on Windows PowerShell 5.1 and on PowerShell 7.x. This page
records what was verified and the three places where output differs between the
two.

Based on a comparison test on 2026-07-30 using PowerShell 5.1.26100.8875 and
PowerShell 7.6.4 on the same machine (ANSI code page 1252, culture de-DE).

## Result in one sentence

Processing and the resulting Excel workbook are identical on both versions. As of
**M30.20260730.1628** two purely cosmetic differences remain in the console and
log output.

## Fixed in M30: exit code was not shown on 5.1

`os4k2excel.ps1` starts `api2hipath.exe` through `Start-Process -PassThru` and
then reads `$process.ExitCode`. Without `-Wait`, that property was **always
`$null`** on Windows PowerShell 5.1:

| Actual exit code | PowerShell 5.1 (before) | PowerShell 7.6 |
|---|---|---|
| 0 | `$null` | 0 |
| 7 | `$null` | 7 |
| 13 | `$null` | 13 |

It showed up in the per-table status line:

```
before (PS 5.1):  API-Abfrage für SWITCH abgeschlossen. Dauer: ... (Exit-Code: )
now    (both) :  API-Abfrage für SWITCH abgeschlossen. Dauer: ... (Exit-Code: 0)
```

The data was never affected — control flow depends on whether the expected CSV
file was produced (`Test-Path`), not on the exit code. Failed API calls could not
be diagnosed from the log, however.

**Fixed** by adding `-Wait` to both `Start-Process` calls (GitHub issue #1).
`test_delete_api.ps1` already used `-Wait` and was never affected — its deletion
sequencing, which does rely on exit codes, worked correctly on both versions.

## Difference 1: redirecting output to a file

The `$OutputEncoding` variable controls how text is encoded on redirection:

| | PowerShell 5.1 | PowerShell 7.6 |
|---|---|---|
| `$OutputEncoding` | `us-ascii` | `utf-8` |

On redirection, PowerShell 5.1 replaces every non-ASCII character with a question
mark. This affects the box-drawing rules of the phase headers, the check mark in
the completion message, and German umlauts:

```powershell
.\os4k2excel.ps1 -ApiHost 10.10.1.1 -ApiUser engr > output.txt
```

On 5.1, `output.txt` then contains `???` instead of `═══`. On screen the output
renders correctly — the script sets `[Console]::OutputEncoding` to UTF-8 for
that.

**Workaround:** on 5.1 use `Tee-Object` or `Start-Transcript` instead of
redirection, or redirect on 7.x.

## Difference 2: byte order mark in the log file

`Out-File -Encoding utf8` behaves differently:

| | PowerShell 5.1 | PowerShell 7.6 |
|---|---|---|
| BOM in `OS4K-PORT-<date>.log` | yes (`EF BB BF`) | no |

The content is identical; the files differ in their first three bytes. Relevant
only if the log is processed by another tool.

## Verified as identical

| Check | Result |
|---|---|
| Syntax check of all four scripts | clean on both versions |
| `Import-Csv` with ASCII and UTF-8 input | identical |
| Culture and date format (`yyyy-MM-dd`) | identical (de-DE) |
| Number formatting in Excel cells | identical |
| Site ordering (`Sort-Object Domain, SwitchName`) | identical, sorted on two fields |
| Duplicate detection (`Group-Object switch_name, pen`) | identical |
| Excel formatting via ImportExcel/EPPlus | identical |

### Character set of the API exports — fixed in M30

Here both versions behaved **the same** — but equally wrongly. It was therefore
never a version difference but a defect independent of the PowerShell version
(GitHub issue #2).

`api2hipath.exe` writes its exports in **Windows-1252**. `Import-Csv` without
`-Encoding` read them as UTF-8; every umlaut byte is invalid there and was
replaced by `U+FFFD`. `Wagenmeisterbüro` became `Wagenmeisterb<U+FFFD>ro`. That
is lossy — the original character cannot be recovered from the workbook.

Demonstrated on the exports of 2026-07-29: 336 umlaut bytes, all in the PORT
tables, distributed exactly across the German umlauts (`0xFC` ü, `0xF6` ö,
`0xDF` ß, `0xDC` Ü, `0xE4` ä, `0xC4` Ä, `0xD6` Ö). In the workbook they showed
up as 255 corrupted cells in the `displayname` column.

**Fixed** by the helper `Import-ApiCsv`, which determines the character set from
the content: valid UTF-8 is read as UTF-8, everything else as Windows-1252. Pure
ASCII tables are unaffected, and a future switch of `api2hipath.exe` to UTF-8
would be handled correctly on its own.

Verified against all exports of 2026-07-29: 336 umlauts read correctly, not a
single `U+FFFD` left, row counts unchanged — with identical results on
PowerShell 5.1 and 7.x.

## Installation pitfall

The two PowerShell versions use **separate module directories**:

| Version | Path for ImportExcel |
|---|---|
| Windows PowerShell 5.1 | `%USERPROFILE%\Documents\WindowsPowerShell\Modules` |
| PowerShell 7.x | `%USERPROFILE%\Documents\PowerShell\Modules` |

Running `Install-Module ImportExcel` on one version does **not** install it for
the other. If you switch between versions you have to install the module twice.
The script checks this at startup and prints both paths when the module is
missing.

## Recommendation

For production use, prefer **PowerShell 7.x**: the exit code is reported
correctly and redirected output stays readable. On 5.1 the result is nonetheless
functionally equivalent.

---

[Back to overview](Home.md)

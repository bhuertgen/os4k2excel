# Security

How os4k2excel handles credentials, which risks remain, and how to run it safely
at a customer site. Based on a security test carried out on 2026-07-30 against
`api2hipath.exe` version 7.00.000.

## Overview

| Topic | Assessment |
|---|---|
| Masked password entry | covered |
| Password in the process command line | residual risk, cannot be fixed |
| Encrypted API connection (`-z`) | currently not used |
| Password in log files | covered (since 2026-03-10) |
| Password in PowerShell history | covered |
| Web interface | transmits in the clear (HTTP) |

## Entering credentials

If no password is supplied, the script prompts for it. Input is masked through
`Read-Host -AsSecureString`, so nothing appears on screen. Both spellings trigger
the prompt:

```powershell
# -ApiPassword without a value
.\os4k2excel.ps1 -ApiHost 10.10.1.1 -ApiUser engr -ApiPassword -ShowSecrets

# or omitted entirely
.\os4k2excel.ps1 -ApiHost 10.10.1.1 -ApiUser engr
```

`-ApiUser` is prompted for as well when missing, though visibly.

`-ApiHost` is deliberately **not** prompted for; if it is missing the script
aborts with usage information.

In non-interactive environments (scheduled tasks, services, CI) the script exits
with an error instead of waiting on a prompt nobody can answer.

The entered value is wiped from unmanaged memory immediately after it is read
(`ZeroFreeBSTR` inside a `finally` block). The resulting .NET string stays in
process memory until garbage collection — a property of PowerShell that cannot
be avoided without a much larger redesign.

## Residual risk: password in the process command line

`api2hipath.exe` accepts the password **only** through the `-p` command-line
switch. Its help output (version 7.00.000) offers no way to pass credentials via
standard input, an environment variable, or a credential file.

The password therefore sits in the process command line in clear text for the
duration of every API call. Testing confirmed that this command line can be read
from another session **without administrator rights**:

```powershell
Get-CimInstance Win32_Process -Filter "Name='api2hipath.exe'" |
    Select-Object ProcessId, CommandLine
```

It is equally visible in Task Manager via the optional "Command line" column, in
Sysmon (event ID 1), in process auditing (event ID 4688), and in many EDR
products.

One run spawns a call per table plus a call per site — several dozen processes on
a mid-sized system.

**This cannot be avoided with this API tool.** The masked prompt protects against
shoulder surfing and against traces in the command history, not against this
process list.

## Recommendation for customer sites

When working on site with other people watching, the following approach works
well:

1. **Use a throwaway account.** Have the customer create a read-only API account
   before the session and change its password afterwards. It then no longer
   matters whether the password appeared in a process list or an event log — it
   is worthless afterwards. This is the only measure that genuinely removes the
   residual risk described above.
2. **Let the customer type the password.** Start the script without
   `-ApiPassword` and hand over the keyboard. The password is entered masked and
   is never visible — not even to the engineer.
3. **Never store credentials in shortcuts or batch files.** For recurring runs,
   do not put the credentials into a scheduled task as clear-text parameters.

## Log files

The log file `OS4K-PORT-<date>.log` written by `os4k2excel.ps1` contains no
credentials. This was verified against real logs.

`test_delete_api.ps1` logs the full API command line and masks the password as
`***`. That masking was introduced on **2026-03-10**.

> **Note on older files:** logs from runs **before** 2026-03-10 contain the API
> password of the time in clear text. Delete such files and treat the affected
> password as compromised and in need of rotation.

## Unencrypted connection

From version 7 onwards `api2hipath.exe` supports the `-z` switch for an
encrypted connection to the server. **os4k2excel does not currently set it.**

Both the logon and all retrieved data — extensions, subscriber names, IP
addresses and SIP secrets — therefore travel across the customer network
unencrypted. Where that is unacceptable, run the export over a protected network
segment or discuss enabling `-z` with whoever maintains the PBX.

## Web interface

`os4k2excel-server.ps1` binds an `HttpListener` to `http://+:8080/`, which means:

- The password set with `-WebPassword` is transmitted unencrypted.
- The session cookie carries `HttpOnly` but no `Secure` attribute.
- The password comparison is character-by-character and thus theoretically open
  to timing analysis.

The web interface is therefore suited to a trusted internal network, not to
exposure across network boundaries.

## Output files contain customer data

The generated CSV, log and Excel files contain extension numbers, subscriber
names, IP addresses and system identifiers. Without `-ShowSecrets` the
`sip_secret` field is exported as `***`; **with** the switch it appears in clear
text in the workbook.

The project `.gitignore` excludes these file types. Review the content before
sharing reports or diagnostic material regardless.

## Summary of the checks performed

| Check | Result |
|---|---|
| Masked password entry | effective |
| Memory wiped after entry (`ZeroFreeBSTR`) | present |
| Password in `OS4K-PORT-*.log` | not present |
| Password masking in `test_delete_*.log` | effective since 2026-03-10 |
| Password in PowerShell history | not present |
| Command line readable without admin rights | yes — residual risk |
| Alternative credential input in `api2hipath.exe` | does not exist |
| `.gitignore` against customer data | complete |

---

[Back to overview](Home.md)

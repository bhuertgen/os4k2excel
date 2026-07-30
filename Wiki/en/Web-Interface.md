# Web interface

*Diese Seite auf Deutsch: [Webserver](../de/Webserver.md)*

`os4k2excel-server.ps1` serves a browser-based interface for starting, monitoring
and downloading the OpenScape 4000 export. A built-in scheduler allows unattended
runs.

## Architecture

```
Browser (port 8080)
  |
  +-- GET  /           --> login page
  +-- POST /login      --> check password, set session cookie
  +-- GET  /dashboard  --> main page: start button, status, downloads, scheduler
  +-- POST /run        --> start the script as a background job
  +-- GET  /status     --> job status as JSON, polled every 2 s
  +-- GET  /download   --> download the Excel file
  +-- POST /email      --> send the workbook over SMTP
  +-- GET  /schedule   --> read the scheduler configuration (JSON)
  +-- POST /schedule   --> store the scheduler configuration
  +-- POST /shutdown   --> stop the server
  +-- GET  /logout     --> end the session
```

**Technology:** plain PowerShell with `System.Net.HttpListener`. No IIS, no
Node.js, no external dependencies — the HTML, CSS and JavaScript are embedded in
the script.

## Installation

### Requirements

- the same as for `os4k2excel.ps1`: PowerShell 5.1+, the ImportExcel module and
  `api2hipath.exe`
- network access to the OpenScape 4000 server
- port 8080 (or the configured port) has to be free

### Port reservation

For non-administrator users the port has to be reserved once, as administrator:

```powershell
# reserve the port
netsh http add urlacl url=http://+:8080/ user=DOMAIN\USERNAME

# remove the reservation again
netsh http delete urlacl url=http://+:8080/
```

### Firewall (optional)

If the server should be reachable from other machines:

```powershell
New-NetFirewallRule -DisplayName "os4k2excel web server" -Direction Inbound `
    -Protocol TCP -LocalPort 8080 -Action Allow
```

## Usage

```powershell
# minimal
.\os4k2excel-server.ps1 -ApiHost "<IP>" -ApiUser "<USER>" -ApiPassword "<PW>" -WebPassword "<WEBPW>"

# with e-mail delivery and a default recipient
.\os4k2excel-server.ps1 -ApiHost "<IP>" -ApiUser "<USER>" -ApiPassword "<PW>" -WebPassword "<WEBPW>" `
    -SmtpServer "mail.example.com" -SmtpFrom "os4k@example.com" -SmtpTo "admin@example.com"
```

Then open `http://servername:8080` in a browser.

### Parameters

| Parameter | Required | Description |
|---|---|---|
| `-ApiHost` | yes | IP address of the OpenScape 4000 |
| `-ApiUser` | yes | API user name |
| `-ApiPassword` | yes | API password |
| `-WebPassword` | yes | password for the web login |
| `-Port` | no | HTTP port, default 8080 |
| `-OutputPath` | no | output directory, default: script directory |
| `-SmtpServer` | no | SMTP server for e-mail delivery |
| `-SmtpPort` | no | SMTP port, default 25 |
| `-SmtpFrom` | no | sender address |
| `-SmtpTo` | no | default recipient, pre-filled in the dashboard |

## Features

- **Login** — password protected, session cookie with a 60 minute timeout
- **Dashboard** — connection information, start button, progress display
- **Start export** — runs `os4k2excel.ps1` as a background job, one run at a time
- **Progress** — live display, polled every 2 seconds
- **Download** — fetch the workbook directly once the run has finished
- **E-mail** — send the workbook as an attachment over SMTP (optional)
- **Scheduler** — unattended runs by weekday and time, optionally with e-mail
- **History** — overview of the recent runs, manual and scheduled

## Stopping the server

Three ways:

1. **Ctrl+C** in the terminal window
2. **Dashboard** → *Server beenden* button, top right
3. **stop file**, created from another terminal:

   ```cmd
   echo.> "C:\path\os4k2excel-server.stop"
   ```

## Security notes

- the web interface has a single shared password — it is meant for a trusted
  network, not for exposure to the internet
- HTTP only, no TLS. If the interface has to leave the local machine, put a
  reverse proxy with TLS in front of it
- the API password is passed as a parameter and kept in memory only; it is
  masked in the log
- SIP secrets are masked in the export unless `-ShowSecrets` is given

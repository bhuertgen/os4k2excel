# Installation

*Diese Seite auf Deutsch: [Installation](../de/Installation.md)*

Step-by-step guide for installing os4k2excel.

> **Shortcut:** the installer does steps 1 and 2 for you:
> ```powershell
> irm https://raw.githubusercontent.com/bhuertgen/os4k2excel/main/install.ps1 | iex
> ```
> Steps 3 to 5 — the Export Table Client and the API user on the PBX — always
> have to be done by hand.

## Step 1: Check the requirements

### System

- **Windows:** 10, 11, Server 2016+
- **PowerShell:** 5.1 or 7.x — both tested (5.1.26100.8875 and 7.6.4).
  PowerShell 7 is recommended for production, see
  [PowerShell 5.1 and 7.x](PowerShell-Compatibility.md)
- **.NET Framework:** 4.5+ (included with Windows)

### Software

- **api2hipath.exe** — XIE API tool by Unify/Mitel
  - part of the OpenScape 4000 Assistant/Manager V11
  - default path: `C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe`
  - must be installed locally
  - documentation: [XIE API Service Documentation](https://www.mitel.com/document-center/business-phone-systems/openscape-4000-ecosystem/openscape-4000/110/en/openscape-4000-assistantmanager-v11-importexport-xie-api-service-documentation)

### Network

- network access to the OpenScape 4000 server
- port 8080 free, if you want to use the web interface

## Step 2: Install the PowerShell module

The script needs the **[ImportExcel](https://www.powershellgallery.com/packages/ImportExcel/)**
module to write Excel files without Microsoft Office (based on EPPlus/OfficeOpenXml).

---

### Option A: Online installation

**1. Install the NuGet package provider** (if not present yet):

```powershell
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
```

**2. Trust the PSGallery** (optional, avoids prompts):

```powershell
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
```

**3. Install ImportExcel:**

```powershell
Install-Module -Name ImportExcel -Scope CurrentUser
```

On a first installation you will be asked to confirm installing the NuGet
provider and trusting the PSGallery — answer `Y` to both.

**4. Verify:**

```powershell
Get-Module -ListAvailable ImportExcel
```

**5. Update later on:**

```powershell
Update-Module -Name ImportExcel
```

---

### Option B: Offline installation

For machines without internet access, download the module elsewhere and copy it
over.

#### B.1 Download the module (on a machine WITH internet)

**Way 1 — PowerShell (recommended):**

```powershell
# downloads into a folder, does NOT install
Save-Module -Name ImportExcel -Path "C:\Temp\PSModules"
```

This creates the correct folder structure automatically:

```
C:\Temp\PSModules\
  └── ImportExcel\
      └── 7.8.10\          <- version number (example)
          ├── ImportExcel.psd1
          ├── ImportExcel.psm1
          └── ...
```

**Way 2 — manual download through a browser:**

1. open https://www.powershellgallery.com/packages/ImportExcel/
2. click the **Manual Download** tab
3. click **Download the raw nupkg file**
4. the file `importexcel.<version>.nupkg` is downloaded
5. rename the extension from `.nupkg` to **`.zip`**
6. extract the archive
7. the extracted folder holds the module files (`ImportExcel.psd1`, `ImportExcel.psm1`, …)
8. rename that folder to **`ImportExcel`**, without a version number
9. `[Content_Types].xml`, `_rels/` and `package/` can be deleted — they belong to
   the NuGet package format and are not needed

#### B.2 Transfer to the offline machine

Copy the `ImportExcel` folder including all subfolders by USB stick, network
share or any other medium.

#### B.3 Install on the offline machine

The module has to be placed in one of the PowerShell module paths.

**Option 1 — current user only** (no admin rights needed):

```powershell
$targetPath = "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\ImportExcel"
New-Item -Path $targetPath -ItemType Directory -Force
Copy-Item -Path "D:\ImportExcel\*" -Destination $targetPath -Recurse -Force
```

**Option 2 — all users** (start PowerShell as administrator):

```powershell
$targetPath = "C:\Program Files\WindowsPowerShell\Modules\ImportExcel"
New-Item -Path $targetPath -ItemType Directory -Force
Copy-Item -Path "D:\ImportExcel\*" -Destination $targetPath -Recurse -Force
```

#### B.4 Verify

```powershell
Get-Module -ListAvailable ImportExcel      # must appear in the list
Import-Module ImportExcel                  # test import
Get-Command -Module ImportExcel | Measure-Object   # around 40+ commands
```

#### Correct directory layout

PowerShell only finds the module when the files sit in the right place:

```
...\Modules\
  └── ImportExcel\                  <- folder name MUST be "ImportExcel"
      ├── ImportExcel.psd1          <- module manifest (required)
      ├── ImportExcel.psm1          <- module code
      └── ...
```

A version subfolder also works — that is what `Save-Module` creates:

```
...\Modules\
  └── ImportExcel\
      └── 7.8.10\
          ├── ImportExcel.psd1
          └── ...
```

PowerShell recognises both layouts.

---

### Looking up the module paths

```powershell
$env:PSModulePath -split ';'
```

| Path | Scope |
|---|---|
| `C:\Users\<name>\Documents\WindowsPowerShell\Modules` | current user |
| `C:\Program Files\WindowsPowerShell\Modules` | all users |
| `C:\Windows\System32\WindowsPowerShell\v1.0\Modules` | system — do not use |

### Updating ImportExcel

PowerShell modules are **not** updated automatically.

```powershell
# which version is installed
Get-Module -ListAvailable ImportExcel | Select-Object Name, Version, ModuleBase

# update (online)
Update-Module -Name ImportExcel
```

> `Update-Module` only works if the module was originally installed with
> `Install-Module`. After a manual offline installation, the update has to be
> done manually as well.

#### Cleaning up old versions

`Update-Module` installs the new version **next to** the old one and does not
remove the previous one — that is intentional, it allows a rollback.

```powershell
Get-Module -ListAvailable ImportExcel | Select-Object Name, Version
Uninstall-Module -Name ImportExcel -RequiredVersion 7.8.9   # adjust the version
```

If `Uninstall-Module` fails, for instance after an offline installation, remove
the folder by hand:

```powershell
Get-Module -ListAvailable ImportExcel | Select-Object Name, Version, ModuleBase
Remove-Item -Path "C:\Users\<name>\Documents\WindowsPowerShell\Modules\ImportExcel\7.8.9" -Recurse -Force
```

#### Offline update

Download the new version as described in B.1, transfer it, then replace the old
files:

```powershell
Get-Module -ListAvailable ImportExcel | Select-Object ModuleBase
Remove-Module ImportExcel -ErrorAction SilentlyContinue   # unload first

$modulePath = "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\ImportExcel"
Remove-Item -Path "$modulePath\*" -Recurse -Force
Copy-Item -Path "D:\ImportExcel\*" -Destination $modulePath -Recurse -Force
```

> **Important:** unload the module before overwriting (`Remove-Module`) — Windows
> locks files that are in use. Alternatively close PowerShell and open it again.

---

## Step 3: External software — api2hipath.exe

Command line tool by Unify/Mitel that accesses the OpenScape 4000 database
through the XIE (Import/Export) API.

| | |
|---|---|
| **Vendor** | Unify / Mitel |
| **Part of** | OpenScape 4000 Assistant / Manager V11 |
| **Default path** | `C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe` |
| **Availability** | installed with the OpenScape 4000 Assistant/Manager, not available separately |
| **Documentation** | *OpenScape 4000 Assistant/Manager V11, Import/Export (XIE) API, Service Documentation* (P31003-H34B0-S102-02-0020, 08/2024) — Mitel Document Center: [EN](https://www.mitel.com/document-center/business-phone-systems/openscape-4000-ecosystem/openscape-4000/110/en/openscape-4000-assistantmanager-v11-importexport-xie-api-service-documentation) \| [DE](https://www.mitel.com/document-center/business-phone-systems/openscape-4000-ecosystem/openscape-4000/110/de/openscape-4000-assistantmanager-v11-importexport-xie-api-service-documentation) |

The tool must be installed on the machine that runs the script and needs network
access to the OpenScape 4000 server.

### Installing the Export Table Client

`api2hipath.exe` ships with the **Export Table Client**, which is installed from
the OpenScape 4000 Assistant:

1. in the Assistant go to **Dienstprogramme** (utilities) → **Export Table Client installieren**
2. prerequisite: **Microsoft Visual C++ 2010 Redistributable Package x86** must be
   installed, also on 64-bit systems
3. download the **Export Table** installation file and run `ExTable.exe`

![Installing the Export Table Client](../5_Export_Table_Client_installieren.png)

Afterwards `api2hipath.exe` is located in
`C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\`.

## Step 4: Create the API user

A dedicated API user is needed to read the OpenScape 4000 database. Create it in
the Assistant under **Zugangsverwaltung** (access management) →
**Kennungsverwaltung** → **Benutzerkennungsverwaltung**.

### 4.1 Add the user

1. click **+ Benutzer hinzufügen**
2. **Neue Kennung** (new account): for example `xieapi`
3. **Beschreibung** (description): for example `Powershell Script`
4. **Sicherheitsprofil** (security profile): `Kunden-Benutzer (cust)` — the lowest privilege
5. click **Benutzer hinzufügen**

![Adding the API user](../1_Benutzerkennungsverwaltung.png)

### 4.2 Set the password and properties

1. select the new user and click **Benutzer bearbeiten**
2. set **Neues Passwort** and repeat it
3. under **Eigenschaften** tick **Passwort ist unbegrenzt gültig** (password never expires)
4. under **Automatisch sperren** set **Kennung automatisch sperren** to `nie` (never)
5. click **Speichern**

![Password and properties](../2_Benutzerkennungsverwaltung.png)

## Step 5: Assign access rights

The API user only needs read access to the tables the script queries through the
Import/Export API (XIE).

### 5.1 Create an access rights group

Under **Zugangsverwaltung** → **Zugriffsrechtegruppen-Konfiguration**:

1. create a new group, for example `XIEAPI`
2. under **Import/Export API (XIE)** grant these tables:
   `CFW`, `DEVCONST`, `DOMAIN`, `HUNTGRP`, `HUNTGRP_SERVICE`, `NUMBERING_PLAN`,
   `PERSPORT`, `PICKUPGRP` *(new in M26)*, `PICKUP_SUB` *(new in M26)*, `PORT`,
   `SWITCH`

![Access rights group XIEAPI with the XIE tables](../3_Zugriffsrechtegruppen_Konfiguration.png)

### 5.2 Assign the group to the user

Under **Zugangsverwaltung** → **Zugriffsrechtekonfiguration**:

1. select the user `xieapi`
2. assign the access rights group `XIEAPI`
3. the summary at the bottom shows the resulting rights

![Assigning the group to the user](../4_Zugriffsrechtekonfiguration.png)

> Only the minimum set of tables is granted deliberately. `PORT` is queried
> dynamically per domain/switch and therefore has to be part of the group as
> well. `PICKUPGRP` and `PICKUP_SUB` are required from version M26 onwards.

## Step 6: Check the configuration

```powershell
$path = 'C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe'
Test-Path $path   # should return True
```

If the tool sits elsewhere, pass its location with `-ApiPath`.

## Step 7: Run the script

```powershell
# minimal
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "secret"

# with diagnostics
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "secret" -Debug

# with all options
.\os4k2excel.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" -ApiPassword "secret" `
  -OutputPath "C:\Exports" -MarkDuplicate -ShowSecrets
```

See [Usage and parameters](Usage-and-Parameters.md) for the full parameter list.

## Step 8: Output files

- **OS4K-PORT-{date}.xlsx** — Excel report with several worksheets
- **OS4K-PORT-{date}.log** — execution log
- **OS4K-{table}-{date}.csv** — raw data of the queries

With `-Debug` the log contains detailed diagnostic information.

> These files contain live PBX data. Keep them out of repositories and shared
> locations.

## Optional: the web interface

```powershell
.\os4k2excel-server.ps1 -ApiHost "192.0.2.10" -ApiUser "engr" `
  -ApiPassword "secret" -WebPassword "team2024"
```

Then open `http://localhost:8080`. See [Web interface](Web-Interface.md) for the
full documentation.

## Troubleshooting

### "api2hipath.exe not found"

- check that the OpenScape 4000 Manager/Assistant V11 is installed
- verify the path with
  `Test-Path 'C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe'`
- if it differs, use `-ApiPath "C:\path\to\api2hipath.exe"`

### "ImportExcel module not found"

```powershell
Install-Module -Name ImportExcel -Scope CurrentUser -Force
```

### "Permission denied for XIE API"

- check the access rights group in the Assistant
- make sure every required table is granted for **reading**
- make sure the group is actually assigned to the API user

### No data is returned

- run the script with `-Debug` and read the log
- check the API exit codes in the log
- test the connection with `api2hipath.exe` manually

### "CSV file not found" in the debug log

- check the API exit codes in the log
- make sure all required tables are part of the access rights
- verify network connectivity to the OS4K server

---

**Last update:** 2026-03-10 | **Release:** v2.26.0 (M26.20260310.1239)

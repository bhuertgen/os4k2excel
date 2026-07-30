<#
.SYNOPSIS
    Installs os4k2excel - OpenScape 4000 port, license and call-forwarding export.

.DESCRIPTION
    Downloads the scripts from GitHub, places them in
    %LOCALAPPDATA%\Programs\os4k2excel and adds that folder to the user PATH, so
    that "os4k2excel.ps1" can be called from any directory.

    Quick install:
        irm https://raw.githubusercontent.com/bhuertgen/os4k2excel/main/install.ps1 | iex

    With options:
        & ([scriptblock]::Create((irm https://raw.githubusercontent.com/bhuertgen/os4k2excel/main/install.ps1))) -InstallDir 'C:\Tools' -NoPath

.PARAMETER InstallDir
    Target directory. Default: %LOCALAPPDATA%\Programs\os4k2excel

.PARAMETER Ref
    Branch, tag or commit to install from. Default: main

.PARAMETER NoPath
    Do not add the target directory to the user PATH.

.PARAMETER NoModule
    Do not install the ImportExcel module, even if it is missing.

.PARAMETER Uninstall
    Removes the scripts, the directory and the PATH entry again.

.NOTES
    Requires Windows PowerShell 5.1 or PowerShell 7.x.

    This file is deliberately stored as UTF-8 WITHOUT BOM: with "irm <url> | iex"
    the content is parsed as a string, and a BOM at the start of that string stops
    PowerShell from recognising the comment block that follows.
#>
[CmdletBinding()]
param(
    [string]$InstallDir,
    [string]$Ref = 'main',
    [switch]$NoPath,
    [switch]$NoModule,
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'

$Repo    = 'bhuertgen/os4k2excel'
$Scripts = @('os4k2excel.ps1', 'os4k2excel-server.ps1')
$Extras  = @('README.md', 'LICENSE')

if (-not $InstallDir) {
    $InstallDir = Join-Path $env:LOCALAPPDATA 'Programs\os4k2excel'
}

#---------------------------------------------------------------------- output
function Write-Step { param([string]$Text) Write-Host "  $Text" }
function Write-Ok   { param([string]$Text) Write-Host "  [ok] $Text" -ForegroundColor Green }
function Write-Note { param([string]$Text) Write-Host "  $Text" -ForegroundColor DarkGray }
function Write-Warn { param([string]$Text) Write-Host "  [!]  $Text" -ForegroundColor Yellow }

Write-Host ''
Write-Host 'os4k2excel' -ForegroundColor Cyan -NoNewline
Write-Host '  -  OpenScape 4000 port, license and call-forwarding data to Excel'
Write-Host ('-' * 74) -ForegroundColor DarkGray

#------------------------------------------------------------------- PATH help
function Get-UserPath {
    $p = [Environment]::GetEnvironmentVariable('Path', 'User')
    if ($p) { return $p } else { return '' }
}

function Test-InPath {
    param([string]$Directory)
    foreach ($t in ((Get-UserPath) -split ';' | Where-Object { $_ })) {
        if ($t.TrimEnd('\') -ieq $Directory.TrimEnd('\')) { return $true }
    }
    return $false
}

function Add-ToPath {
    param([string]$Directory)
    if (Test-InPath $Directory) { return $false }
    $old = Get-UserPath
    $new = if ($old) { $old.TrimEnd(';') + ';' + $Directory } else { $Directory }
    [Environment]::SetEnvironmentVariable('Path', $new, 'User')
    $env:Path = $env:Path.TrimEnd(';') + ';' + $Directory   # current session too
    return $true
}

function Remove-FromPath {
    param([string]$Directory)
    $parts = (Get-UserPath) -split ';' | Where-Object { $_ }
    $rest  = $parts | Where-Object { $_.TrimEnd('\') -ine $Directory.TrimEnd('\') }
    if ($parts.Count -eq $rest.Count) { return $false }
    [Environment]::SetEnvironmentVariable('Path', ($rest -join ';'), 'User')
    return $true
}

#-------------------------------------------------------------------- uninstall
if ($Uninstall) {
    Write-Host 'Uninstall' -ForegroundColor Cyan
    if (Test-Path -LiteralPath $InstallDir) {
        Remove-Item -LiteralPath $InstallDir -Recurse -Force
        Write-Ok "directory removed: $InstallDir"
    }
    else {
        Write-Note "directory did not exist: $InstallDir"
    }
    if (Remove-FromPath $InstallDir) { Write-Ok 'PATH entry removed' }
    else { Write-Note 'no PATH entry found' }
    Write-Host ''
    Write-Note 'The ImportExcel module was kept - other scripts may rely on it.'
    Write-Host ''
    return
}

#---------------------------------------------------------------- requirements
Write-Host 'Requirements' -ForegroundColor Cyan

$psv = $PSVersionTable.PSVersion
if ($psv.Major -lt 5 -or ($psv.Major -eq 5 -and $psv.Minor -lt 1)) {
    throw "PowerShell 5.1 or newer is required, found: $psv"
}
Write-Ok "PowerShell $psv"

# Windows PowerShell 5.1 needs TLS 1.2 enabled explicitly, otherwise the
# download from GitHub and the PowerShell Gallery fails.
if ($psv.Major -lt 6) {
    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }
    catch { Write-Warn 'Could not enable TLS 1.2 - the download may fail.' }
}

$ie = Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue |
      Sort-Object Version -Descending | Select-Object -First 1
if ($ie) {
    Write-Ok "ImportExcel module $($ie.Version)"
}
elseif ($NoModule) {
    Write-Warn 'ImportExcel module missing (-NoModule set, no installation).'
}
else {
    Write-Step 'ImportExcel module missing - installing from the PowerShell Gallery ...'

    # On a freshly installed Windows the NuGet provider is missing and the
    # PSGallery is registered as untrusted. Both make Install-Module ask a
    # question, which fails in a non-interactive session.
    try {
        $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        if (-not $nuget -or $nuget.Version -lt [version]'2.8.5.201') {
            Write-Step '  providing the NuGet package provider ...'
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 `
                -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
            Import-PackageProvider -Name NuGet -Force -ErrorAction SilentlyContinue | Out-Null
            Write-Ok 'NuGet package provider ready'
        }
    }
    catch { Write-Warn "Could not provide the NuGet package provider: $($_.Exception.Message)" }

    $policyBefore = $null
    try {
        $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($repo -and $repo.InstallationPolicy -ne 'Trusted') {
            $policyBefore = $repo.InstallationPolicy
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
        }
    }
    catch { Write-Warn "Could not change the PSGallery setting: $($_.Exception.Message)" }

    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        $ie = Get-Module -ListAvailable -Name ImportExcel |
              Sort-Object Version -Descending | Select-Object -First 1
        if ($ie) { Write-Ok "ImportExcel module $($ie.Version) installed" }
        else      { Write-Warn 'Installation reported success, but the module was not found.' }
    }
    catch {
        Write-Warn "Installation failed: $($_.Exception.Message)"
        Write-Note 'This is not fatal - the script also writes CSV without the module.'
        Write-Note ''
        Write-Note 'Later, with internet access:'
        Write-Note '  Install-Module ImportExcel -Scope CurrentUser'
        Write-Note 'Without internet access on this machine - on another PC:'
        Write-Note '  Save-Module ImportExcel -Path C:\Transfer'
        Write-Note 'then copy the ImportExcel folder to'
        Write-Note '  %USERPROFILE%\Documents\WindowsPowerShell\Modules   (PowerShell 5.1)'
        Write-Note '  %USERPROFILE%\Documents\PowerShell\Modules          (PowerShell 7.x)'
    }
    finally {
        if ($policyBefore) {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy $policyBefore } catch { }
        }
    }
}

#------------------------------------------------------------------- download
Write-Host ''
Write-Host 'Installation' -ForegroundColor Cyan

if (-not (Test-Path -LiteralPath $InstallDir)) {
    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
}
Write-Step "target directory: $InstallDir"

$base = "https://raw.githubusercontent.com/$Repo/$Ref"
foreach ($file in ($Scripts + $Extras)) {
    $url    = "$base/$file"
    $target = Join-Path $InstallDir $file
    try {
        Invoke-WebRequest -Uri $url -OutFile $target -UseBasicParsing
        $kb = [math]::Round((Get-Item -LiteralPath $target).Length / 1KB, 1)
        Write-Ok ("{0,-26} {1,7} KB" -f $file, $kb)
    }
    catch {
        if ($Scripts -contains $file) {
            throw "Download failed: $url`n$($_.Exception.Message)"
        }
        Write-Warn "$file not downloaded (optional)"
    }
}

# make sure the downloaded scripts actually parse
foreach ($file in $Scripts) {
    $path = Join-Path $InstallDir $file
    $errors = $null; $tokens = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        throw "$file is not a valid PowerShell script ($($errors.Count) syntax errors)."
    }
}
Write-Ok 'syntax check passed'

#------------------------------------------------------------------------ PATH
if ($NoPath) {
    Write-Note 'PATH unchanged (-NoPath set).'
}
elseif (Add-ToPath $InstallDir) {
    Write-Ok 'target directory added to the user PATH'
}
else {
    Write-Note 'target directory was already in PATH'
}

#-------------------------------------------------------------------- finished
Write-Host ''
Write-Host 'Done' -ForegroundColor Green
Write-Host ''
Write-Host '  Usage:' -ForegroundColor Cyan
if ($NoPath) {
    Write-Host "    & '$(Join-Path $InstallDir 'os4k2excel.ps1')' -ApiHost <host> -ApiUser <user>"
}
else {
    Write-Host '    os4k2excel.ps1 -ApiHost <host> -ApiUser <user> -ApiPassword <password>'
    Write-Host '    os4k2excel-server.ps1            web interface'
    Write-Host ''
    Write-Note 'In terminals that are already open, reload the PATH or restart the terminal.'
}
Write-Host ''
Write-Host '  Uninstall:' -ForegroundColor Cyan
Write-Host "    & ([scriptblock]::Create((irm $base/install.ps1))) -Uninstall"
Write-Host ''

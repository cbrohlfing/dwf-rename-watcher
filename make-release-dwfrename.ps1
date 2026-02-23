param(
    [string]$ToVersion
)

# ==========================================================
# DWF Rename - Make Release (Working -> Release)
#
# Source:
#   C:\Scripts\DWF Rename\working\Rename-IdwDwf-Watcher-Tray.ps1
#   C:\Scripts\DWF Rename\working\dwf_rename_settings.json (optional)
#
# Destination (created):
#   C:\Scripts\DWF Rename\vX.Y\
#     Rename-IdwDwf-Watcher-Tray.ps1
#     dwf_rename_settings.json (optional)
#     Run DWF Rename vX.Y.lnk
#
# Version rule: vX.Y only (legacy vX.Y.Z still detected as latest)
# ==========================================================

$root       = "C:\Scripts\DWF Rename"
$workingDir = Join-Path $root "working"
$workingPs1 = Join-Path $workingDir "Rename-IdwDwf-Watcher-Tray.ps1"
$workingCfg = Join-Path $workingDir "dwf_rename_settings.json"

function Bump-Minor([string]$vText) {
    if ($vText -notmatch '^v(\d+)\.(\d+)$') { throw "Expected vX.Y, got $vText" }
    $maj = [int]$matches[1]
    $min = [int]$matches[2]
    return ("v{0}.{1}" -f $maj, ($min + 1))
}

function Parse-Version([string]$name) {
    # Accept v1.2 OR v1.2.3 (legacy)
    if ($name -match '^v(\d+)\.(\d+)(?:\.(\d+))?$') {
        return [PSCustomObject]@{
            Major = [int]$matches[1]
            Minor = [int]$matches[2]
            Patch = if ($matches[3]) { [int]$matches[3] } else { 0 }
            Text  = $name
        }
    }
    return $null
}

function Get-LatestVersionDir([string]$rootPath) {
    $vers =
        Get-ChildItem -LiteralPath $rootPath -Directory -ErrorAction Stop |
        Where-Object { $_.Name -like "v*" } |
        ForEach-Object { Parse-Version $_.Name } |
        Where-Object { $_ -ne $null } |
        Sort-Object Major, Minor, Patch

    if (-not $vers -or $vers.Count -eq 0) { return $null }
    return $vers[-1]
}

function Next-VersionText($v) {
    # bump minor: v1.2(.3) -> v1.3
    return ("v{0}.{1}" -f $v.Major, ($v.Minor + 1))
}

function Assert-ToVersion([string]$v) {
    if ($v -notmatch '^v\d+\.\d+$') {
        throw "ToVersion must be in vX.Y format going forward (example: v1.7). You provided: $v"
    }
}

function Get-VersionFromRaw([string]$raw, [string]$pathForError) {
    if ($raw -match '\$script:Version\s*=\s*"([^"]+)"') {
        return $matches[1]
    }
    throw "Missing `$script:Version assignment in: $pathForError"
}

function Replace-AllExact([string]$raw, [string]$from, [string]$to) {
    return ($raw -replace [regex]::Escape($from), $to)
}

function New-Shortcut {
    param(
        [Parameter(Mandatory = $true)][string]$ShortcutPath,
        [Parameter(Mandatory = $true)][string]$TargetPath,
        [Parameter(Mandatory = $true)][string]$Arguments,
        [string]$WorkingDirectory,
        [string]$Description
    )

    $wsh = New-Object -ComObject WScript.Shell
    $sc = $wsh.CreateShortcut($ShortcutPath)
    $sc.TargetPath = $TargetPath
    $sc.Arguments = $Arguments
    if ($WorkingDirectory) { $sc.WorkingDirectory = $WorkingDirectory }
    if ($Description) { $sc.Description = $Description }
    $sc.Save()
}

if (!(Test-Path -LiteralPath $workingPs1)) {
    throw "Missing working script: $workingPs1"
}

$latest = Get-LatestVersionDir $root
if ($null -eq $latest) {
    $latest = [PSCustomObject]@{ Major = 1; Minor = 0; Patch = 0; Text = "v1.0" }
}

if ([string]::IsNullOrWhiteSpace($ToVersion)) {
    $ToVersion = Next-VersionText $latest
}
Assert-ToVersion $ToVersion

$toDir  = Join-Path $root $ToVersion
$dstPs1 = Join-Path $toDir "Rename-IdwDwf-Watcher-Tray.ps1"
$dstCfg = Join-Path $toDir "dwf_rename_settings.json"

if (Test-Path -LiteralPath $toDir) { throw "Target folder already exists: $toDir" }

New-Item -ItemType Directory -Path $toDir -Force | Out-Null
Copy-Item -LiteralPath $workingPs1 -Destination $dstPs1 -Force

if (Test-Path -LiteralPath $workingCfg) {
    Copy-Item -LiteralPath $workingCfg -Destination $dstCfg -Force
}

# ---- Update RELEASE script: replace EVERY vX.Y occurrence based on the file's own $script:Version
$raw = Get-Content -LiteralPath $dstPs1 -Raw
$fromVersion = Get-VersionFromRaw -raw $raw -pathForError $dstPs1

# Replace all exact occurrences of the old version string with the new version
$raw = Replace-AllExact -raw $raw -from $fromVersion -to $ToVersion

# Also fix any folder-layout comment paths that embed a version folder
$raw = $raw -replace '(?m)^(#\s*.*\\DWF Rename\\)v\d+\.\d+(?:\.\d+)?(\\)', "`$1$ToVersion`$2"

Set-Content -LiteralPath $dstPs1 -Value $raw -Encoding UTF8

# Parse test (hard fail if broken)
[void][ScriptBlock]::Create((Get-Content -LiteralPath $dstPs1 -Raw))

"OK: Released WORKING -> $ToVersion"
"OK: Parse passed for $dstPs1"

# Create launcher shortcut in the new folder
$psExe   = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
$lnkPath = Join-Path $toDir "Run DWF Rename $ToVersion.lnk"
$args    = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File `"$dstPs1`""

New-Shortcut -ShortcutPath $lnkPath `
    -TargetPath $psExe `
    -Arguments $args `
    -WorkingDirectory $toDir `
    -Description "Launch DWF Rename ($ToVersion)"

"OK: Shortcut created: $lnkPath"

# ---- Bump WORKING to next minor, and replace ALL old version strings there too
$nextWorking = Bump-Minor $ToVersion

$wraw = Get-Content -LiteralPath $workingPs1 -Raw
$wFromVersion = Get-VersionFromRaw -raw $wraw -pathForError $workingPs1

$wraw = Replace-AllExact -raw $wraw -from $wFromVersion -to $nextWorking

# Also fix any folder-layout comment paths that embed a version folder
$wraw = $wraw -replace '(?m)^(#\s*.*\\DWF Rename\\)v\d+\.\d+(?:\.\d+)?(\\)', "`$1$nextWorking`$2"

Set-Content -LiteralPath $workingPs1 -Value $wraw -Encoding UTF8

"OK: Working bumped to $nextWorking"

# Runs the latest version folder under C:\Scripts\DWF Rename (excluding "working")
$root = "C:\Scripts\DWF Rename"

function Parse-Version([string]$name) {
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

$vers = Get-ChildItem -LiteralPath $root -Directory -ErrorAction Stop |
    Where-Object { $_.Name -like "v*" } |
    ForEach-Object { Parse-Version $_.Name } |
    Where-Object { $_ -ne $null } |
    Sort-Object Major, Minor, Patch

if (-not $vers -or $vers.Count -eq 0) {
    throw "No version folders found under: $root"
}

$latest = $vers[-1]
$latestDir = Join-Path $root $latest.Text
$ps1 = Join-Path $latestDir "Rename-IdwDwf-Watcher-Tray.ps1"

if (!(Test-Path -LiteralPath $ps1)) {
    throw "Missing script in latest folder: $ps1"
}

$psExe = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
$args = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File `"$ps1`""

Start-Process -FilePath $psExe -ArgumentList $args -WorkingDirectory $latestDir | Out-Null

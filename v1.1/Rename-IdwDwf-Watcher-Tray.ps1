# ==========================================================
# DWF Auto-Renamer w/ Tray Icon (VPN/UNC safe) - v1.1
# - JSON config (no settings UI)
# - Double-click opens PRIMARY folder (configurable)
# - Multi-folder support via settings.json
# ==========================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Version-local paths (script folder) ---
$script:ScriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:SettingsPath = Join-Path $script:ScriptDir "dwf_rename_settings.json"
$script:LogPath      = Join-Path $script:ScriptDir "dwf_rename_watcher.log"

# Defaults if JSON missing/bad
$defaults = [PSCustomObject]@{
    baseServer = "\\Usvaldcfs1.na.valmont.com\Cadshare"
    paths = @("{BASE}\ORDERS\TP\{USERNAME}")
    startupWaitSeconds = 900
    healthCheckSeconds = 5
    sweepSeconds = 10
    maxWatcherAgeSeconds = 1800
    offlineRetrySeconds = 5
    logEnabled = $true
    logMaxLines = 1000
    logTrimEvery = 25
    openPrimaryOnDoubleClick = $true
}

function Load-Settings([string]$path, $fallback) {
    try {
        if (Test-Path -LiteralPath $path) {
            $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
            $cfg = $raw | ConvertFrom-Json -ErrorAction Stop
            return [PSCustomObject]@{
                baseServer = if ($cfg.baseServer) { [string]$cfg.baseServer } else { $fallback.baseServer }
                paths = if ($cfg.paths -and $cfg.paths.Count -gt 0) { @($cfg.paths | ForEach-Object { [string]$_ }) } else { @($fallback.paths) }
                startupWaitSeconds = if ($cfg.startupWaitSeconds) { [int]$cfg.startupWaitSeconds } else { $fallback.startupWaitSeconds }
                healthCheckSeconds = if ($cfg.healthCheckSeconds) { [int]$cfg.healthCheckSeconds } else { $fallback.healthCheckSeconds }
                sweepSeconds = if ($cfg.sweepSeconds) { [int]$cfg.sweepSeconds } else { $fallback.sweepSeconds }
                maxWatcherAgeSeconds = if ($cfg.maxWatcherAgeSeconds) { [int]$cfg.maxWatcherAgeSeconds } else { $fallback.maxWatcherAgeSeconds }
                offlineRetrySeconds = if ($cfg.offlineRetrySeconds) { [int]$cfg.offlineRetrySeconds } else { $fallback.offlineRetrySeconds }
                logEnabled = if ($null -ne $cfg.logEnabled) { [bool]$cfg.logEnabled } else { [bool]$fallback.logEnabled }
                logMaxLines = if ($null -ne $cfg.logMaxLines) { [int]$cfg.logMaxLines } else { [int]$fallback.logMaxLines }
                logTrimEvery = if ($null -ne $cfg.logTrimEvery) { [int]$cfg.logTrimEvery } else { [int]$fallback.logTrimEvery }
                openPrimaryOnDoubleClick = if ($null -ne $cfg.openPrimaryOnDoubleClick) { [bool]$cfg.openPrimaryOnDoubleClick } else { [bool]$fallback.openPrimaryOnDoubleClick }
            }
        }
    } catch {}
    return $fallback
}

$cfg = Load-Settings -path $script:SettingsPath -fallback $defaults

function Expand-PathTokens([string]$p, [string]$base, [string]$user) {
    return $p.Replace("{BASE}", $base).Replace("{USERNAME}", $user)
}

$BaseServer = $cfg.baseServer
$Username   = $env:USERNAME
$WatchPaths = @($cfg.paths | ForEach-Object { Expand-PathTokens $_ $BaseServer $Username })

$StartupWaitSeconds   = $cfg.startupWaitSeconds
$HealthCheckSeconds   = $cfg.healthCheckSeconds
$SweepSeconds         = $cfg.sweepSeconds
$MaxWatcherAgeSeconds = $cfg.maxWatcherAgeSeconds
$OfflineRetrySeconds  = $cfg.offlineRetrySeconds

$EnableLog   = $cfg.logEnabled
$MaxLogLines = $cfg.logMaxLines
$trimEvery   = $cfg.logTrimEvery

function Log($msg) {
    if (-not $EnableLog) { return }
    $line = "$(Get-Date -Format s) $msg"
    try {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
        if ($MaxLogLines -gt 0) {
            if (-not $script:logWriteCount) { $script:logWriteCount = 0 }
            $script:logWriteCount++
            if (($script:logWriteCount % $trimEvery) -eq 0) {
                $tail = Get-Content -LiteralPath $script:LogPath -Tail $MaxLogLines -ErrorAction SilentlyContinue
                if ($tail) { Set-Content -LiteralPath $script:LogPath -Value $tail -Encoding UTF8 }
            }
        }
    } catch {}
}

function New-ColorIcon([System.Drawing.Color]$color) {
    $bmp = New-Object System.Drawing.Bitmap 16, 16
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::Transparent)
    $brush = New-Object System.Drawing.SolidBrush $color
    $pen   = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(80,0,0,0)), 1
    $g.FillEllipse($brush, 1, 1, 14, 14)
    $g.DrawEllipse($pen,   1, 1, 14, 14)
    $g.Dispose(); $brush.Dispose(); $pen.Dispose()
    return [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
}

$IconStarting = New-ColorIcon ([System.Drawing.Color]::Gold)
$IconOnline   = New-ColorIcon ([System.Drawing.Color]::LimeGreen)
$IconOffline  = New-ColorIcon ([System.Drawing.Color]::Tomato)

$notify = New-Object System.Windows.Forms.NotifyIcon
$notify.Icon = $IconStarting
$notify.Text = "DWF Auto-Rename v1.1 (starting...)"
$notify.Visible = $true

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$miOpenPrimary = $menu.Items.Add("Open Folder (Primary)")
$miOpenPick    = $menu.Items.Add("Open Folder (Pick...)")
$miViewLog     = $menu.Items.Add("View Log")
$menu.Items.Add("-") | Out-Null
$miExit        = $menu.Items.Add("Exit")
$notify.ContextMenuStrip = $menu

$script:exitRequested = $false
$miExit.Add_Click({ $script:exitRequested = $true })

if ($cfg.openPrimaryOnDoubleClick) {
    $notify.add_DoubleClick({ try { Start-Process explorer.exe $WatchPaths[0] } catch {} })
}

$miOpenPrimary.Add_Click({ try { Start-Process explorer.exe $WatchPaths[0] } catch {} })

$miOpenPick.Add_Click({
    try {
        if ($WatchPaths.Count -le 1) { Start-Process explorer.exe $WatchPaths[0]; return }
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Open Folder (v1.1)"
        $form.StartPosition = "CenterScreen"
        $form.Width = 900
        $form.Height = 180
        $form.Topmost = $true
        $combo = New-Object System.Windows.Forms.ComboBox
        $combo.Left = 12; $combo.Top = 20; $combo.Width = 860
        $combo.DropDownStyle = "DropDownList"
        [void]$combo.Items.AddRange($WatchPaths)
        $combo.SelectedIndex = 0
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = "Open"; $btn.Left = 12; $btn.Top = 70; $btn.Width = 120
        $btn.Add_Click({ try { Start-Process explorer.exe $combo.SelectedItem } catch {}; $form.Close() })
        $form.Controls.Add($combo); $form.Controls.Add($btn)
        [void]$form.ShowDialog()
    } catch {}
})

$miViewLog.Add_Click({ try { Start-Process notepad.exe $script:LogPath } catch {} })

function Set-TrayStatus([string]$status) {
    switch ($status) {
        "starting" { $notify.Icon = $IconStarting; $notify.Text = "DWF Auto-Rename v1.1 (starting...)" }
        "online"   { $notify.Icon = $IconOnline;   $notify.Text = "DWF Auto-Rename v1.1 (online)" }
        "offline"  { $notify.Icon = $IconOffline;  $notify.Text = "DWF Auto-Rename v1.1 (offline)" }
        "mixed"    { $notify.Icon = $IconStarting; $notify.Text = "DWF Auto-Rename v1.1 (partial connectivity)" }
    }
}

function Wait-ForAnyPath([string[]]$paths, [int]$timeoutSeconds) {
    $deadline = (Get-Date).AddSeconds($timeoutSeconds)
    while ((Get-Date) -lt $deadline -and -not $script:exitRequested) {
        foreach ($p in $paths) { if (Test-Path -LiteralPath $p) { return $true } }
        Start-Sleep -Seconds $OfflineRetrySeconds
    }
    foreach ($p in $paths) { if (Test-Path -LiteralPath $p) { return $true } }
    return $false
}

function Is-ShareHealthy([string]$path) {
    try {
        if (-not (Test-Path -LiteralPath $path)) { return $false }
        $null = Get-ChildItem -LiteralPath $path -ErrorAction Stop -Force | Select-Object -First 1
        return $true
    } catch { return $false }
}

function Try-RenameIdwDwf([string]$path) {
    try {
        if (-not (Test-Path -LiteralPath $path)) { return }
        $name = [System.IO.Path]::GetFileName($path)
        if ($name -notlike "*.idw.dwf") { return }

        for ($t=0; $t -lt 60; $t++) {
            if (-not (Test-Path -LiteralPath $path)) { return }
            try {
                $fs = [System.IO.File]::Open($path, 'Open', 'ReadWrite', 'None')
                $fs.Close()
                break
            } catch {
                Start-Sleep -Milliseconds 250
            }
        }

        if (-not (Test-Path -LiteralPath $path)) { return }

        $newName = $name -replace "\.idw\.dwf$", ".dwf"
        $newPath = Join-Path (Split-Path $path -Parent) $newName
        if (Test-Path -LiteralPath $newPath) { return }

        Rename-Item -LiteralPath $path -NewName $newName -ErrorAction Stop
        Log "RENAMED: '$name' -> '$newName'"
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match "does not exist" -or $msg -match "cannot find" -or $msg -match "ItemNotFound") { return }
        Log "ERROR renaming '$path' : $msg"
    }
}

function Sweep-Renames([string]$path) {
    Get-ChildItem -LiteralPath $path -File -Filter "*.idw.dwf" -ErrorAction Stop | ForEach-Object {
        Try-RenameIdwDwf $_.FullName
    }
}

function Start-Watcher([string]$path) {
    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = $path
    $watcher.Filter = "*.dwf"
    $watcher.IncludeSubdirectories = $false
    $watcher.NotifyFilter = [System.IO.NotifyFilters]'FileName, LastWrite'
    $watcher.EnableRaisingEvents = $true

    $createdSub = Register-ObjectEvent $watcher Created -Action { Try-RenameIdwDwf $Event.SourceEventArgs.FullPath }
    $renamedSub = Register-ObjectEvent $watcher Renamed -Action { Try-RenameIdwDwf $Event.SourceEventArgs.FullPath }

    Log "WATCHER STARTED on $path"
    return [PSCustomObject]@{
        Path      = $path
        Watcher   = $watcher
        Subs      = @($createdSub, $renamedSub)
        StartedAt = Get-Date
    }
}

function Stop-Watcher($state) {
    if ($null -eq $state) { return }
    foreach ($s in $state.Subs) {
        try { Unregister-Event -SubscriptionId $s.Id -ErrorAction SilentlyContinue } catch {}
        try { Remove-Job -Id $s.Id -Force -ErrorAction SilentlyContinue } catch {}
    }
    try { $state.Watcher.EnableRaisingEvents = $false } catch {}
    try { $state.Watcher.Dispose() } catch {}
    Log "WATCHER STOPPED on $($state.Path)"
}

Log "SCRIPT START (v1.1): settings=$script:SettingsPath paths=$($WatchPaths -join '; ')"
Set-TrayStatus "starting"

[void](Wait-ForAnyPath -paths $WatchPaths -timeoutSeconds $StartupWaitSeconds)

$watcherStates = @{}
$lastHealth = Get-Date "2000-01-01"
$lastSweep  = Get-Date "2000-01-01"
$lastStatus = ""

while (-not $script:exitRequested) {
    [System.Windows.Forms.Application]::DoEvents()
    $now = Get-Date

    if (($now - $lastHealth).TotalSeconds -ge $HealthCheckSeconds) {
        $lastHealth = $now

        $onlineCount = 0
        foreach ($p in $WatchPaths) {
            $healthy = Is-ShareHealthy $p
            if ($healthy) { $onlineCount++ }

            if (-not $healthy) {
                if ($watcherStates.ContainsKey($p)) {
                    Stop-Watcher $watcherStates[$p]
                    $watcherStates.Remove($p) | Out-Null
                }
            } else {
                if (-not $watcherStates.ContainsKey($p)) {
                    $watcherStates[$p] = Start-Watcher $p
                    try { Sweep-Renames $p } catch {}
                } else {
                    if (($now - $watcherStates[$p].StartedAt).TotalSeconds -ge $MaxWatcherAgeSeconds) {
                        Log "REFRESHING WATCHER (age limit reached) on $p"
                        Stop-Watcher $watcherStates[$p]
                        $watcherStates[$p] = Start-Watcher $p
                    }
                }
            }
        }

        $status =
            if ($onlineCount -eq 0) { "offline" }
            elseif ($onlineCount -eq $WatchPaths.Count) { "online" }
            else { "mixed" }

        if ($status -ne $lastStatus) {
            Set-TrayStatus $status
            Log "STATUS: $status ($onlineCount/$($WatchPaths.Count) folders reachable)"
            $lastStatus = $status
        }

        if ($onlineCount -eq 0) {
            Start-Sleep -Seconds $OfflineRetrySeconds
            continue
        }
    }

    if (($now - $lastSweep).TotalSeconds -ge $SweepSeconds) {
        $lastSweep = $now
        foreach ($p in $WatchPaths) {
            try { if (Is-ShareHealthy $p) { Sweep-Renames $p } } catch {}
        }
    }

    Start-Sleep -Milliseconds 200
}

try {
    foreach ($p in @($watcherStates.Keys)) {
        Stop-Watcher $watcherStates[$p]
        $watcherStates.Remove($p) | Out-Null
    }
} catch {}

$notify.Visible = $false
$notify.Dispose()

Log "SCRIPT EXIT (v1.1)"

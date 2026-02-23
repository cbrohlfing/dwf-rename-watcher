# ==========================================================
# DWF Auto-Renamer w/ Tray Icon + Settings UI + Status Window
# Version: v1.2 (lock + housekeeping + status)
#
# Folder layout (versioned):
#   C:\Scripts\v1.2\Rename-IdwDwf-Watcher-Tray.ps1
#   C:\Scripts\v1.2\dwf_rename_settings.json
#   C:\Scripts\v1.2\dwf_rename_watcher.log
#
# Features:
# - Tray icon status: green=online(all), red=offline(all), yellow=mixed/starting
# - Double-click tray icon opens PRIMARY folder (configurable)
# - Right-click menu: Open Folder (Primary), Status..., Settings, View Log, Exit
# - Status window: shows each watched folder + ONLINE/OFFLINE + last checked (live)
# - Watches for *.idw.dwf and renames to *.dwf
# - Auto-recovers after VPN disconnect/reconnect
# - Periodic sweep fallback so missed events still get fixed
# - Multi-folder support via settings.json (tokens: {BASE}, {USERNAME})
# - Per-file lock (*.lock) to avoid multi-PC rename races
# - Auto-clears stale locks after 2 minutes + housekeeping sweep
# ==========================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -------------------------
# Version-local paths
# -------------------------
$script:ScriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:SettingsPath = Join-Path $script:ScriptDir "dwf_rename_settings.json"
$script:LogPath      = Join-Path $script:ScriptDir "dwf_rename_watcher.log"

# -------------------------
# Defaults (used if JSON missing/invalid)
# -------------------------
$script:DefaultSettings = [PSCustomObject]@{
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
                startupWaitSeconds = if ($null -ne $cfg.startupWaitSeconds) { [int]$cfg.startupWaitSeconds } else { [int]$fallback.startupWaitSeconds }
                healthCheckSeconds = if ($null -ne $cfg.healthCheckSeconds) { [int]$cfg.healthCheckSeconds } else { [int]$fallback.healthCheckSeconds }
                sweepSeconds = if ($null -ne $cfg.sweepSeconds) { [int]$cfg.sweepSeconds } else { [int]$fallback.sweepSeconds }
                maxWatcherAgeSeconds = if ($null -ne $cfg.maxWatcherAgeSeconds) { [int]$cfg.maxWatcherAgeSeconds } else { [int]$fallback.maxWatcherAgeSeconds }
                offlineRetrySeconds = if ($null -ne $cfg.offlineRetrySeconds) { [int]$cfg.offlineRetrySeconds } else { [int]$fallback.offlineRetrySeconds }
                logEnabled = if ($null -ne $cfg.logEnabled) { [bool]$cfg.logEnabled } else { [bool]$fallback.logEnabled }
                logMaxLines = if ($null -ne $cfg.logMaxLines) { [int]$cfg.logMaxLines } else { [int]$fallback.logMaxLines }
                logTrimEvery = if ($null -ne $cfg.logTrimEvery) { [int]$cfg.logTrimEvery } else { [int]$fallback.logTrimEvery }
                openPrimaryOnDoubleClick = if ($null -ne $cfg.openPrimaryOnDoubleClick) { [bool]$cfg.openPrimaryOnDoubleClick } else { [bool]$fallback.openPrimaryOnDoubleClick }
            }
        }
    } catch {
        # fall back below
    }
    return $fallback
}

function Save-Settings([string]$path, $settingsObject) {
    $json = $settingsObject | ConvertTo-Json -Depth 6
    Set-Content -LiteralPath $path -Value $json -Encoding UTF8
}

function Expand-PathTokens([string]$p, [string]$base, [string]$user) {
    return $p.Replace("{BASE}", $base).Replace("{USERNAME}", $user)
}

# -------------------------
# GLOBAL mirrors
# NOTE:
# WinForms event handlers (Timer/Click) often execute in PSEventHandler scope.
# $script: inside those handlers may NOT refer to this PS1.
# So we mirror shared state in $global: and have UI read from there.
# -------------------------
$global:DwfWatchPaths = @()
$global:DwfPathStatus = @{}
$global:DwfStatusForm = $null

# Load initial cfg
$script:cfg = Load-Settings -path $script:SettingsPath -fallback $script:DefaultSettings

# Apply cfg to runtime variables
function Apply-Config($newCfg) {
    $script:cfg = $newCfg

    $script:BaseServer = $script:cfg.baseServer
    $script:UserName   = $env:USERNAME

    $script:WatchPaths = @(
        $script:cfg.paths | ForEach-Object { Expand-PathTokens $_ $script:BaseServer $script:UserName }
    )

    # Keep global mirror in sync
    $global:DwfWatchPaths = @($script:WatchPaths)

    $script:StartupWaitSeconds   = [int]$script:cfg.startupWaitSeconds
    $script:HealthCheckSeconds   = [int]$script:cfg.healthCheckSeconds
    $script:SweepSeconds         = [int]$script:cfg.sweepSeconds
    $script:MaxWatcherAgeSeconds = [int]$script:cfg.maxWatcherAgeSeconds
    $script:OfflineRetrySeconds  = [int]$script:cfg.offlineRetrySeconds

    $script:EnableLog   = [bool]$script:cfg.logEnabled
    $script:MaxLogLines = [int]$script:cfg.logMaxLines
    $script:trimEvery   = [int]$script:cfg.logTrimEvery
    if ($script:trimEvery -lt 1) { $script:trimEvery = 25 }
}

Apply-Config $script:cfg

# -------------------------
# Logging (append at bottom, capped)
# -------------------------
function Log($msg) {
    if (-not $script:EnableLog) { return }
    $line = "$(Get-Date -Format s) $msg"
    try {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8

        if ($script:MaxLogLines -gt 0) {
            if (-not $script:logWriteCount) { $script:logWriteCount = 0 }
            $script:logWriteCount++
            if (($script:logWriteCount % $script:trimEvery) -eq 0) {
                $tail = Get-Content -LiteralPath $script:LogPath -Tail $script:MaxLogLines -ErrorAction SilentlyContinue
                if ($tail) { Set-Content -LiteralPath $script:LogPath -Value $tail -Encoding UTF8 }
            }
        }
    } catch {}
}

# -------------------------
# Tray icon helpers
# -------------------------
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
$notify.Text = "DWF Auto-Rename v1.2 (starting...)"
$notify.Visible = $true

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$miOpenPrimary = $menu.Items.Add("Open Folder (Primary)")
$miOpenPrimary.Font = New-Object System.Drawing.Font($miOpenPrimary.Font, [System.Drawing.FontStyle]::Bold)

$miStatus      = $menu.Items.Add("Status...")
$miSettings    = $menu.Items.Add("Settings...")
$miViewLog     = $menu.Items.Add("View Log")
$menu.Items.Add("-") | Out-Null
$miExit        = $menu.Items.Add("Exit")

$notify.ContextMenuStrip = $menu

function Set-TrayStatus([string]$status) {
    switch ($status) {
        "starting" { $notify.Icon = $IconStarting; $notify.Text = "DWF Auto-Rename v1.2 (starting...)" }
        "online"   { $notify.Icon = $IconOnline;   $notify.Text = "DWF Auto-Rename v1.2 (online)" }
        "offline"  { $notify.Icon = $IconOffline;  $notify.Text = "DWF Auto-Rename v1.2 (offline)" }
        "mixed"    { $notify.Icon = $IconStarting; $notify.Text = "DWF Auto-Rename v1.2 (partial connectivity)" }
    }
}

$script:exitRequested = $false
$miExit.Add_Click({ $script:exitRequested = $true })

# Double-click behavior
$notify.add_DoubleClick({
    try {
        if ($script:cfg.openPrimaryOnDoubleClick -and $script:WatchPaths.Count -ge 1) {
            Start-Process explorer.exe $script:WatchPaths[0]
        }
    } catch {}
})

$miOpenPrimary.Add_Click({
    try {
        if ($script:WatchPaths.Count -ge 1) { Start-Process explorer.exe $script:WatchPaths[0] }
    } catch {}
})

$miViewLog.Add_Click({
    try { Start-Process notepad.exe $script:LogPath } catch {}
})

# -------------------------
# Core functions
# -------------------------
function Wait-ForAnyPath([string[]]$paths, [int]$timeoutSeconds) {
    $deadline = (Get-Date).AddSeconds($timeoutSeconds)
    while ((Get-Date) -lt $deadline -and -not $script:exitRequested) {
        foreach ($p in $paths) { if (Test-Path -LiteralPath $p) { return $true } }
        Start-Sleep -Seconds $script:OfflineRetrySeconds
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

# -------------------------
# Lock + housekeeping
# -------------------------
$script:LockStaleMinutes = 2

function Get-LockPath([string]$targetPath) { return "$targetPath.lock" }

function Try-AcquireLock([string]$targetPath) {
    $lockPath = Get-LockPath $targetPath
    try {
        if (Test-Path -LiteralPath $lockPath) {
            try {
                $age = (Get-Date) - (Get-Item -LiteralPath $lockPath -ErrorAction Stop).LastWriteTime
                if ($age.TotalMinutes -ge $script:LockStaleMinutes) {
                    Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
                }
            } catch {}
        }

        $fs = [System.IO.File]::Open(
            $lockPath,
            [System.IO.FileMode]::CreateNew,
            [System.IO.FileAccess]::Write,
            [System.IO.FileShare]::None
        )

        $bytes = [System.Text.Encoding]::UTF8.GetBytes("$env:COMPUTERNAME\$env:USERNAME $(Get-Date -Format s)")
        $fs.Write($bytes, 0, $bytes.Length)
        $fs.Close()
        return $true
    } catch { return $false }
}

function Release-Lock([string]$targetPath) {
    $lockPath = Get-LockPath $targetPath
    try { Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue } catch {}
}

function Cleanup-OrphanLocks([string]$folderPath) {
    try {
        Get-ChildItem -LiteralPath $folderPath -File -Filter "*.dwf.lock" -ErrorAction Stop | ForEach-Object {
            $lockFile = $_.FullName
            $baseFile = $lockFile.Substring(0, $lockFile.Length - 5)  # remove ".lock"

            if (-not (Test-Path -LiteralPath $baseFile)) {
                try { Remove-Item -LiteralPath $lockFile -Force -ErrorAction SilentlyContinue } catch {}
            } else {
                try {
                    $age = (Get-Date) - $_.LastWriteTime
                    if ($age.TotalMinutes -ge $script:LockStaleMinutes) {
                        Remove-Item -LiteralPath $lockFile -Force -ErrorAction SilentlyContinue
                    }
                } catch {}
            }
        }
    } catch {}
}

function Try-RenameIdwDwf([string]$path) {
    $lockAcquired = $false
    try {
        if (-not (Test-Path -LiteralPath $path)) { return }

        $name = [System.IO.Path]::GetFileName($path)
        if ($name -notlike "*.idw.dwf") { return }

        $lockAcquired = Try-AcquireLock $path
        if (-not $lockAcquired) { return }

        for ($t=0; $t -lt 60; $t++) {
            if (-not (Test-Path -LiteralPath $path)) { return }
            try {
                $fs = [System.IO.File]::Open($path, 'Open', 'ReadWrite', 'None')
                $fs.Close()
                break
            } catch { Start-Sleep -Milliseconds 250 }
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
    } finally {
        if ($lockAcquired) { Release-Lock $path }
    }
}

function Sweep-Renames([string]$path) {
    Get-ChildItem -LiteralPath $path -File -Filter "*.idw.dwf" -ErrorAction Stop | ForEach-Object {
        Try-RenameIdwDwf $_.FullName
    }
    Cleanup-OrphanLocks $path
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

# -------------------------
# Status window (FIXED: uses GLOBAL mirrors)
# -------------------------
$script:StatusForm = $null

function Show-StatusWindow {
    if ($script:StatusForm -and -not $script:StatusForm.IsDisposed) {
        $script:StatusForm.Activate()
        return
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "DWF Auto-Rename Status (v1.2)"
    $form.StartPosition = "CenterScreen"
    $form.Width = 980
    $form.Height = 420
    $form.Topmost = $false

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Top'
    $panel.Height = 44

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh Now"
    $btnRefresh.Left = 12
    $btnRefresh.Top = 10
    $btnRefresh.Width = 110

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Left = 140
    $lblHint.Top = 14
    $lblHint.Width = 820
    $lblHint.Text = "Green tray = all reachable | Yellow = some reachable | Red = none reachable"

    $panel.Controls.Add($btnRefresh)
    $panel.Controls.Add($lblHint)

    $list = New-Object System.Windows.Forms.ListView
    $list.View = 'Details'
    $list.FullRowSelect = $true
    $list.GridLines = $true
    $list.Dock = 'Fill'
    $list.Font = New-Object System.Drawing.Font("Consolas", 10)

    $colFolder  = $list.Columns.Add("Folder", 720)
    $colStatus  = $list.Columns.Add("Status", 90)
    $colChecked = $list.Columns.Add("Last Checked", 140)

    $form.Controls.Add($list)
    $form.Controls.Add($panel)

    # Store all UI refs on the form so event handlers never lose scope
    $state = [PSCustomObject]@{
        List       = $list
        ColFolder  = $colFolder
        ColStatus  = $colStatus
        ColChecked = $colChecked
        Timer      = $null
        Refresh    = $null
    }

    $state.Refresh = {
        $checkNow = Get-Date

        try {
            $state.List.BeginUpdate()
            $state.List.Items.Clear()

            foreach ($p in @($global:DwfWatchPaths)) {
                $healthy = $false
                try { $healthy = Is-ShareHealthy $p } catch { $healthy = $false }

                # update global cache
                if (-not $global:DwfPathStatus.ContainsKey($p)) {
                    $global:DwfPathStatus[$p] = [PSCustomObject]@{
                        Healthy     = $healthy
                        LastChecked = $checkNow
                        LastChanged = $checkNow
                    }
                } else {
                    $prev = $global:DwfPathStatus[$p]
                    $changed = ([bool]$prev.Healthy) -ne ([bool]$healthy)
                    $global:DwfPathStatus[$p] = [PSCustomObject]@{
                        Healthy     = $healthy
                        LastChecked = $checkNow
                        LastChanged = if ($changed) { $checkNow } else { $prev.LastChanged }
                    }
                }

                $statusText = if ($healthy) { "ONLINE" } else { "OFFLINE" }
                $checked    = $checkNow.ToString("yyyy-MM-dd HH:mm:ss")

                $item = New-Object System.Windows.Forms.ListViewItem($p)
                [void]$item.SubItems.Add($statusText)
                [void]$item.SubItems.Add($checked)

                if (-not $healthy) { $item.BackColor = [System.Drawing.Color]::MistyRose }
                [void]$state.List.Items.Add($item)
            }

            # Auto-size ONLY Status + Last Checked (Folder stays user adjustable)
            try {
                $state.ColStatus.AutoResize([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
                $state.ColChecked.AutoResize([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
                $state.ColStatus.Width  = [Math]::Max($state.ColStatus.Width + 12, 80)
                $state.ColChecked.Width = [Math]::Max($state.ColChecked.Width + 12, 140)
            } catch {}

        } catch {
            Log "STATUS REFRESH ERROR: $($_.Exception.Message)"
        } finally {
            try { $state.List.EndUpdate() } catch {}
        }
    }.GetNewClosure()

    $form.Tag = $state

    $btnRefresh.Add_Click({
        try {
            $sb = $form.Tag.Refresh
            if ($sb -is [scriptblock]) { & $sb }
        } catch {
            Log "STATUS REFRESH CLICK ERROR: $($_.Exception.Message)"
        }
    })

    # NOTE: System.Windows.Forms.Timer runs on the UI thread (no InvokeRequired needed)
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
        try {
            if ($form.IsDisposed) { return }
            $sb = $form.Tag.Refresh
            if ($sb -is [scriptblock]) { & $sb }
        } catch {
            Log "STATUS TIMER ERROR: $($_.Exception.Message)"
        }
    })
    $timer.Start()
    $state.Timer = $timer

    $form.Add_FormClosed({
        try { $timer.Stop(); $timer.Dispose() } catch {}
        $script:StatusForm = $null
        $global:DwfStatusForm = $null
    })

    $script:StatusForm = $form
    $global:DwfStatusForm = $form

    # Initial refresh
    try { & $form.Tag.Refresh } catch {}

    $form.Show() | Out-Null
}



$miStatus.Add_Click({ try { Show-StatusWindow } catch {} })

function Refresh-StatusWindowNow {
    try {
        if ($global:DwfStatusForm -and -not $global:DwfStatusForm.IsDisposed) {
            $sb = $global:DwfStatusForm.Tag.Refresh
            if ($sb -is [scriptblock]) { & $sb }
        }
    } catch {
        Log "STATUS REFRESH NOW ERROR: $($_.Exception.Message)"
    }
}


# -------------------------
# Settings UI
# -------------------------
function Show-SettingsDialog {
    param($currentCfg)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "DWF Auto-Rename Settings (v1.2)"
    $form.StartPosition = "CenterScreen"
    $form.Width = 920
    $form.Height = 560
    $form.Topmost = $true

    $lblBase = New-Object System.Windows.Forms.Label
    $lblBase.Text = "Base Server:"
    $lblBase.Left = 12
    $lblBase.Top = 15
    $lblBase.Width = 120

    $txtBase = New-Object System.Windows.Forms.TextBox
    $txtBase.Left = 140
    $txtBase.Top = 12
    $txtBase.Width = 750
    $txtBase.Text = [string]$currentCfg.baseServer

    $lblPaths = New-Object System.Windows.Forms.Label
    $lblPaths.Text = "Folders to watch (one per line):"
    $lblPaths.Left = 12
    $lblPaths.Top = 50
    $lblPaths.Width = 320

    $txtPaths = New-Object System.Windows.Forms.TextBox
    $txtPaths.Left = 12
    $txtPaths.Top = 75
    $txtPaths.Width = 878
    $txtPaths.Height = 250
    $txtPaths.Multiline = $true
    $txtPaths.ScrollBars = "Vertical"
    $txtPaths.Font = New-Object System.Drawing.Font("Consolas", 10)
    $txtPaths.Text = (@($currentCfg.paths) -join "`r`n")

    $lblTokens = New-Object System.Windows.Forms.Label
    $lblTokens.Left = 12
    $lblTokens.Top = 330
    $lblTokens.Width = 878
    $lblTokens.Text = "Tip: Tokens supported: {BASE} and {USERNAME}. Example: {BASE}\ORDERS\TP\{USERNAME}"

    $chkDbl = New-Object System.Windows.Forms.CheckBox
    $chkDbl.Left = 12
    $chkDbl.Top = 360
    $chkDbl.Width = 520
    $chkDbl.Text = "Double-click tray icon opens PRIMARY folder"
    $chkDbl.Checked = [bool]$currentCfg.openPrimaryOnDoubleClick

    $lblHealth = New-Object System.Windows.Forms.Label
    $lblHealth.Left = 12
    $lblHealth.Top = 392
    $lblHealth.Width = 220
    $lblHealth.Text = "Health check seconds:"

    $numHealth = New-Object System.Windows.Forms.NumericUpDown
    $numHealth.Left = 240
    $numHealth.Top = 390
    $numHealth.Width = 90
    $numHealth.Minimum = 1
    $numHealth.Maximum = 60
    $numHealth.Value = [int]$currentCfg.healthCheckSeconds

    $lblSweep = New-Object System.Windows.Forms.Label
    $lblSweep.Left = 360
    $lblSweep.Top = 392
    $lblSweep.Width = 180
    $lblSweep.Text = "Sweep seconds:"

    $numSweep = New-Object System.Windows.Forms.NumericUpDown
    $numSweep.Left = 540
    $numSweep.Top = 390
    $numSweep.Width = 90
    $numSweep.Minimum = 1
    $numSweep.Maximum = 300
    $numSweep.Value = [int]$currentCfg.sweepSeconds

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save + Apply"
    $btnSave.Left = 660
    $btnSave.Top = 450
    $btnSave.Width = 110

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Left = 780
    $btnCancel.Top = 450
    $btnCancel.Width = 110

    $script:settingsDialogResult = $null
    $btnCancel.Add_Click({ $form.Close() })

    $btnSave.Add_Click({
        $lines = $txtPaths.Lines | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        if ($lines.Count -lt 1) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please specify at least one folder to watch.",
                "Settings",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }

        $newCfg = [PSCustomObject]@{
            baseServer = $txtBase.Text.Trim()
            paths = @($lines)
            startupWaitSeconds = [int]$currentCfg.startupWaitSeconds
            healthCheckSeconds = [int]$numHealth.Value
            sweepSeconds = [int]$numSweep.Value
            maxWatcherAgeSeconds = [int]$currentCfg.maxWatcherAgeSeconds
            offlineRetrySeconds = [int]$currentCfg.offlineRetrySeconds
            logEnabled = [bool]$currentCfg.logEnabled
            logMaxLines = [int]$currentCfg.logMaxLines
            logTrimEvery = [int]$currentCfg.logTrimEvery
            openPrimaryOnDoubleClick = [bool]$chkDbl.Checked
        }

        try {
            Save-Settings -path $script:SettingsPath -settingsObject $newCfg
            $script:settingsDialogResult = $newCfg
            $form.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to save settings:`r`n$($_.Exception.Message)",
                "Settings",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    })

    $form.Controls.AddRange(@(
        $lblBase, $txtBase,
        $lblPaths, $txtPaths, $lblTokens,
        $chkDbl,
        $lblHealth, $numHealth,
        $lblSweep, $numSweep,
        $btnSave, $btnCancel
    ))

    [void]$form.ShowDialog()
    return $script:settingsDialogResult
}

# -------------------------
# Watcher state + Apply settings live
# -------------------------
$script:watcherStates = @{}  # keyed by path

function Stop-AllWatchers {
    try {
        foreach ($p in @($script:watcherStates.Keys)) {
            Stop-Watcher $script:watcherStates[$p]
            $script:watcherStates.Remove($p) | Out-Null
        }
    } catch {}
}

function Apply-NewSettings($newCfg) {
    Stop-AllWatchers
    Apply-Config $newCfg

    # Prune global path-status to match new watch list
    try {
        $newMap = @{}
        foreach ($p in $global:DwfWatchPaths) {
            if ($global:DwfPathStatus.ContainsKey($p)) { $newMap[$p] = $global:DwfPathStatus[$p] }
        }
        $global:DwfPathStatus = $newMap
    } catch {}

    Log "APPLIED SETTINGS: paths=$($global:DwfWatchPaths -join '; ')"
    Set-TrayStatus "starting"

    # Force immediate status UI update if open
    Refresh-StatusWindowNow

    # Force immediate recompute on next loop
    $script:lastHealth = Get-Date "2000-01-01"
    $script:lastStatus = ""
}

$miSettings.Add_Click({
    try {
        $result = Show-SettingsDialog -currentCfg $script:cfg
        if ($null -ne $result) { Apply-NewSettings $result }
    } catch {
        Log "SETTINGS ERROR: $($_.Exception.Message)"
    }
})

# -------------------------
# Main loop
# -------------------------
Log "SCRIPT START (v1.2): settings=$script:SettingsPath paths=$($global:DwfWatchPaths -join '; ')"
Set-TrayStatus "starting"

[void](Wait-ForAnyPath -paths $global:DwfWatchPaths -timeoutSeconds $script:StartupWaitSeconds)

$script:lastHealth = Get-Date "2000-01-01"
$script:lastSweep  = Get-Date "2000-01-01"
$script:lastStatus = ""

while (-not $script:exitRequested) {
    [System.Windows.Forms.Application]::DoEvents()
    $now = Get-Date

    # Health check
    if (($now - $script:lastHealth).TotalSeconds -ge $script:HealthCheckSeconds) {
        $script:lastHealth = $now

        $onlineCount = 0
        foreach ($p in @($global:DwfWatchPaths)) {
            $healthy = Is-ShareHealthy $p

            if (-not $global:DwfPathStatus.ContainsKey($p)) {
                $global:DwfPathStatus[$p] = [PSCustomObject]@{
                    Healthy     = $healthy
                    LastChecked = $now
                    LastChanged = $now
                }
            } else {
                $prev = $global:DwfPathStatus[$p]
                $changed = ([bool]$prev.Healthy) -ne ([bool]$healthy)
                $global:DwfPathStatus[$p] = [PSCustomObject]@{
                    Healthy     = $healthy
                    LastChecked = $now
                    LastChanged = if ($changed) { $now } else { $prev.LastChanged }
                }
            }

            if ($healthy) { $onlineCount++ }

            if (-not $healthy) {
                if ($script:watcherStates.ContainsKey($p)) {
                    Stop-Watcher $script:watcherStates[$p]
                    $script:watcherStates.Remove($p) | Out-Null
                }
            } else {
                if (-not $script:watcherStates.ContainsKey($p)) {
                    $script:watcherStates[$p] = Start-Watcher $p
                    try { Sweep-Renames $p } catch {}
                } else {
                    if (($now - $script:watcherStates[$p].StartedAt).TotalSeconds -ge $script:MaxWatcherAgeSeconds) {
                        Log "REFRESHING WATCHER (age limit reached) on $p"
                        Stop-Watcher $script:watcherStates[$p]
                        $script:watcherStates[$p] = Start-Watcher $p
                    }
                }
            }
        }

        $status =
            if ($onlineCount -eq 0) { "offline" }
            elseif ($onlineCount -eq $global:DwfWatchPaths.Count) { "online" }
            else { "mixed" }

        if ($status -ne $script:lastStatus) {
            Set-TrayStatus $status
            Log "STATUS: $status ($onlineCount/$($global:DwfWatchPaths.Count) folders reachable)"
            $script:lastStatus = $status
        }

        if ($onlineCount -eq 0) {
            Start-Sleep -Seconds $script:OfflineRetrySeconds
            continue
        }
    }

    # Sweep fallback
    if (($now - $script:lastSweep).TotalSeconds -ge $script:SweepSeconds) {
        $script:lastSweep = $now
        foreach ($p in @($global:DwfWatchPaths)) {
            try { if (Is-ShareHealthy $p) { Sweep-Renames $p } } catch {}
        }
    }

    Start-Sleep -Milliseconds 200
}

# Cleanup
try {
    foreach ($p in @($script:watcherStates.Keys)) {
        Stop-Watcher $script:watcherStates[$p]
        $script:watcherStates.Remove($p) | Out-Null
    }
} catch {}

$notify.Visible = $false
$notify.Dispose()

Log "SCRIPT EXIT (v1.2)"

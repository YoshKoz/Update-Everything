<#
.SYNOPSIS
    System-wide update script for Windows
.DESCRIPTION
    Updates all package managers, system components, and development tools.
    Supports parallel execution of independent package managers for faster runs.
.VERSION
    2.6.2
.NOTES
    Run as Administrator for full functionality (some features require elevation).
    Requires PowerShell 7+ for -Parallel support.
.EXAMPLE
    .\updatescript.ps1
    .\updatescript.ps1 -FastMode
    .\updatescript.ps1 -Parallel
    .\updatescript.ps1 -AutoElevate
    .\updatescript.ps1 -Schedule -ScheduleTime "03:00"
    .\updatescript.ps1 -SkipCleanup -SkipWindowsUpdate
    .\updatescript.ps1 -SkipNode -SkipRust -SkipGo
    .\updatescript.ps1 -DeepClean
    .\updatescript.ps1 -UpdateOllamaModels
    .\updatescript.ps1 -WhatChanged
    .\updatescript.ps1 -DryRun
    .\updatescript.ps1 -NoParallel
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$SkipWindowsUpdate,
    [switch]$SkipReboot,
    [switch]$SkipDestructive,
    [switch]$FastMode,                  # Skip slower managers
    [switch]$NoElevate,                 # Compatibility switch (prevents auto-elevation)
    [switch]$AutoElevate,               # Opt-in: relaunch elevated (opens a new window due UAC)
    [switch]$NoPause,                   # Skip "Press Enter to close" prompt (for VS Code / CI)
    [switch]$Parallel,                  # Enable parallel updates where supported (e.g. PSResourceGet, PS7+)
    [switch]$SkipWSL,
    [switch]$SkipWSLDistros,            # Skip updating WSL distro packages (apt/pacman etc.)
    [switch]$SkipDefender,
    [switch]$SkipStoreApps,
    [switch]$SkipUVTools,
    [switch]$SkipVSCodeExtensions,
    [switch]$SkipPoetry,
    [switch]$SkipComposer,
    [switch]$SkipRuby,
    [switch]$SkipPowerShellModules,
    [switch]$SkipCleanup,               # Skip the system cleanup section
    [int]$WingetTimeoutSec = 300,       # Per-call timeout for winget (seconds, default 5 min)
    [switch]$Schedule,                  # Register a daily scheduled task to run this script
    [string]$ScheduleTime = "03:00",    # Time for the scheduled task (default: 3 AM)
    [string]$LogPath,
    [switch]$SkipNode,                  # Skip all Node.js toolchain updates (npm, pnpm, bun, deno, fnm, volta)
    [switch]$SkipRust,                  # Skip Rust toolchain updates (rustup, cargo)
    [switch]$SkipGo,                    # Skip Go toolchain updates
    [switch]$SkipFlutter,               # Skip Flutter SDK update
    [switch]$SkipGitLFS,                # Skip Git LFS client update
    [switch]$DeepClean,                 # Run DISM WinSxS cleanup, DO cache, and prefetch (adds ~7 min)
    [switch]$UpdateOllamaModels,        # Opt-in: pull latest for every installed Ollama model
    [switch]$WhatChanged,               # Show packages that changed since last run
    [switch]$NoParallel,                # Disable parallel execution of independent tools (PS7+)
    [switch]$DryRun                     # Show which steps would run without executing them
)

# ── Console encoding / external tool output ─────────────────────────────────────
try {
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [Console]::InputEncoding = $utf8NoBom
    [Console]::OutputEncoding = $utf8NoBom
    $OutputEncoding = $utf8NoBom
}
catch {
    # Best effort only; continue if host does not allow setting console encodings.
}

# ── Self-elevate ────────────────────────────────────────────────────────────────
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and $AutoElevate -and -not $NoElevate) {
    $params = @()
    if ($SkipWindowsUpdate) { $params += '-SkipWindowsUpdate' }
    if ($SkipReboot) { $params += '-SkipReboot' }
    if ($SkipDestructive) { $params += '-SkipDestructive' }
    if ($FastMode) { $params += '-FastMode' }
    if ($Parallel) { $params += '-Parallel' }
    if ($SkipWSL) { $params += '-SkipWSL' }
    if ($SkipWSLDistros) { $params += '-SkipWSLDistros' }
    if ($SkipDefender) { $params += '-SkipDefender' }
    if ($SkipStoreApps) { $params += '-SkipStoreApps' }
    if ($SkipUVTools) { $params += '-SkipUVTools' }
    if ($SkipVSCodeExtensions) { $params += '-SkipVSCodeExtensions' }
    if ($SkipPoetry) { $params += '-SkipPoetry' }
    if ($SkipComposer) { $params += '-SkipComposer' }
    if ($SkipRuby) { $params += '-SkipRuby' }
    if ($SkipPowerShellModules) { $params += '-SkipPowerShellModules' }
    if ($SkipCleanup) { $params += '-SkipCleanup' }
    if ($NoPause) { $params += '-NoPause' }
    if ($WingetTimeoutSec -ne 300) { $params += @('-WingetTimeoutSec', $WingetTimeoutSec) }
    if ($AutoElevate) { $params += '-AutoElevate' }
    if ($LogPath) { $params += @('-LogPath', $LogPath) }
    if ($SkipNode) { $params += '-SkipNode' }
    if ($SkipRust) { $params += '-SkipRust' }
    if ($SkipGo) { $params += '-SkipGo' }
    if ($SkipFlutter) { $params += '-SkipFlutter' }
    if ($SkipGitLFS) { $params += '-SkipGitLFS' }
    if ($DeepClean) { $params += '-DeepClean' }
    if ($UpdateOllamaModels) { $params += '-UpdateOllamaModels' }
    if ($WhatChanged) { $params += '-WhatChanged' }
    if ($NoParallel) { $params += '-NoParallel' }
    if ($DryRun) { $params += '-DryRun' }

    $pwshArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath) + $params
    try {
        $pwshCmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue

        if ($pwshCmd) {
            Start-Process -FilePath $pwshCmd.Source -Verb RunAs -ArgumentList $pwshArgs -Wait
        }
        else {
            Write-Host "WARNING: pwsh.exe not found. Falling back to Windows PowerShell." -ForegroundColor Yellow
            Start-Process powershell -Verb RunAs -ArgumentList $pwshArgs -Wait
        }
        exit
    }
    catch {
        Write-Host "WARNING: Could not elevate. Running without Administrator privileges.`n" -ForegroundColor Yellow
    }
}
elseif (-not $isAdmin -and -not $NoElevate) {
    Write-Host "INFO: Running in the current terminal without elevation (admin-only tasks may be skipped)." -ForegroundColor DarkYellow
    Write-Host "      To run elevated, start the terminal as Administrator or use -AutoElevate (opens a new window)." -ForegroundColor DarkYellow
}

# ── Scheduled task registration ─────────────────────────────────────────────────
if ($Schedule) {
    $taskName = "DailySystemUpdate"
    $action = New-ScheduledTaskAction -Execute 'pwsh.exe' `
        -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -SkipReboot -SkipWindowsUpdate"
    $trigger = New-ScheduledTaskTrigger -Daily -At $ScheduleTime
    $settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -StartWhenAvailable
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
        -Settings $settings -RunLevel Highest -Force | Out-Null
    Write-Host "[OK] Scheduled task '$taskName' registered to run daily at $ScheduleTime." -ForegroundColor Green
    exit
}

$ErrorActionPreference = "Continue"
$startTime = Get-Date
$commandCache = @{}
$script:sectionTimings = [ordered]@{}      # Name -> elapsed seconds

# ── Helper functions ─────────────────────────────────────────────────────────────

function Write-Section {
    param([string]$Title, [double]$ElapsedSeconds = 0)
    $elapsed = if ($ElapsedSeconds -gt 0) { "  ({0:F1}s)" -f $ElapsedSeconds } else { "" }
    Write-Host "`n$('=' * 54)" -ForegroundColor DarkGray
    Write-Host " $Title$elapsed" -ForegroundColor Cyan
    Write-Host "$('=' * 54)" -ForegroundColor DarkGray
}

function Test-Command([string]$Command) {
    if ($commandCache.ContainsKey($Command)) { return $commandCache[$Command] }
    $result = [bool](Get-Command $Command -ErrorAction SilentlyContinue)
    $commandCache[$Command] = $result
    return $result
}

function Get-VSCodeCliPath {
    # Always prefer the CLI shim (.cmd) to avoid launching Code.exe UI windows.
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code\bin\code.cmd'),
        (Join-Path $env:ProgramFiles 'Microsoft VS Code\bin\code.cmd'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft VS Code\bin\code.cmd'),
        (Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code Insiders\bin\code-insiders.cmd'),
        (Join-Path $env:ProgramFiles 'Microsoft VS Code Insiders\bin\code-insiders.cmd'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft VS Code Insiders\bin\code-insiders.cmd')
    ) | Where-Object { $_ -and (Test-Path $_) }

    if ($candidates.Count -gt 0) { return $candidates[0] }

    foreach ($name in @('code.cmd', 'code-insiders.cmd')) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($cmd -and $cmd.Source -and (Test-Path $cmd.Source)) { return $cmd.Source }
    }

    return $null
}

function Write-Status {
    param(
        [string]$Message,
        [ValidateSet('Success', 'Warning', 'Error', 'Info')]
        [string]$Type = 'Info'
    )
    $colors = @{ Success = 'Green'; Warning = 'Yellow'; Error = 'Red'; Info = 'Gray' }
    $symbols = @{ Success = '[OK]'; Warning = '[!]'; Error = '[X]'; Info = '[*]' }
    Write-Host "$($symbols[$Type]) $Message" -ForegroundColor $colors[$Type]
}

function Write-Detail {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Muted', 'Warning', 'Error')]
        [string]$Type = 'Info'
    )

    $colors = @{ Info = 'Gray'; Muted = 'DarkGray'; Warning = 'Yellow'; Error = 'Red' }
    $prefixes = @{ Info = '  >'; Muted = '  -'; Warning = '  !'; Error = '  x' }
    Write-Host "$($prefixes[$Type]) $Message" -ForegroundColor $colors[$Type]
}

function Write-IndentedOutput {
    param(
        [AllowNull()][string]$Text,
        [string]$Prefix = '  >',
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return }

    $normalized = $Text -replace '\x1b\[[0-9;?]*[ -/]*[@-~]', ''
    $normalized = $normalized -replace "`r", "`n"

    foreach ($rawLine in ($normalized -split "`n")) {
        $line = $rawLine.Trim()
        if (-not $line) { continue }
        Write-Host "$Prefix $line" -ForegroundColor $Color
    }
}

function Write-FilteredOutput {
    param(
        [AllowNull()][string]$Text,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return }

    # Remove common ANSI escape sequences and normalize carriage-return progress updates.
    $normalized = $Text -replace '\x1b\[[0-9;?]*[ -/]*[@-~]', ''
    $normalized = $normalized -replace "`r", "`n"

    foreach ($rawLine in ($normalized -split "`n")) {
        $line = $rawLine.TrimEnd()
        if (-not $line) { continue }

        $compact = $line.Trim()

        # Drop spinner-only frames emitted by winget and installers.
        if ($compact -match '^[\\/\|\-]+$') { continue }

        # Drop garbled/Unicode progress bar rows (block chars, regardless of whether a size suffix follows).
        if ($compact -match 'Γû|[█▓▒░▏▎▍▌▋▊▉■□▪▫]') { continue }

        # Drop noisy winget "cannot be determined" info lines.
        if ($compact -match 'package\(s\) have version numbers that cannot be determined') { continue }

        # Drop table separators and blank winget headers that add noise without new information.
        if ($compact -match '^[\-=]{6,}$') { continue }

        Write-Host $line -ForegroundColor $Color
    }
}

function Get-WingetFailedPackageIds {
    param([AllowNull()][string]$Text)

    $failedIds = [System.Collections.Generic.List[string]]::new()
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }

    $currentId = $null
    $normalized = $Text -replace "`r", "`n"

    foreach ($rawLine in ($normalized -split "`n")) {
        $line = $rawLine.Trim()
        if (-not $line) { continue }

        if ($line -match '^\(\d+/\d+\)\s+Found\s+.+\s+\[(?<id>[^\]]+)\]\s+Version\b') {
            $currentId = $Matches['id']
            continue
        }

        if (-not $currentId) { continue }

        # Terminal states for the current package
        if ($line -match '^(Successfully installed|No applicable upgrade found|Package already installed)') {
            $currentId = $null
            continue
        }

        # Common winget failure markers (including portable-package edge cases)
        if ($line -match '(?i)installer failed with exit code' -or
            $line -match '(?i)unable to remove portable package' -or
            $line -match '(?i)upgrade failed') {
            if (-not $failedIds.Contains($currentId)) {
                $failedIds.Add($currentId)
            }
            $currentId = $null
            continue
        }
    }

    return @($failedIds)
}

function Invoke-WingetWithTimeout {
    <#
    .SYNOPSIS
        Runs winget with a hard timeout, writing output to temp files.
        Avoids both the pipeline-buffering hang and .NET event-handler thread crashes.
    #>
    param(
        [string[]]$Arguments,
        [int]$TimeoutSec = 300
    )

    $stdoutFile = [System.IO.Path]::GetTempFileName()
    $stderrFile = [System.IO.Path]::GetTempFileName()

    try {
        # Start-Process with file redirection: no in-memory buffers, no threading issues.
        $proc = Start-Process -FilePath 'winget' `
            -ArgumentList $Arguments `
            -RedirectStandardOutput $stdoutFile `
            -RedirectStandardError  $stderrFile `
            -NoNewWindow -PassThru

        $exited = $proc.WaitForExit($TimeoutSec * 1000)

        if (-not $exited) {
            try { $proc.Kill() } catch { }
            throw "winget timed out after ${TimeoutSec}s — a hanging installer may require manual intervention"
        }

        $exitCode = $proc.ExitCode
        $stdout = Get-Content -Raw -Path $stdoutFile -ErrorAction SilentlyContinue
        $stderr = Get-Content -Raw -Path $stderrFile -ErrorAction SilentlyContinue
        $combined = (($stdout + $stderr) -replace '\x00', '').Trim()

        return [pscustomobject]@{ Output = $combined; ExitCode = $exitCode }
    }
    finally {
        Remove-Item $stdoutFile, $stderrFile -Force -ErrorAction SilentlyContinue
    }
}

$updateResults = @{
    Success = [System.Collections.Generic.List[string]]::new()
    Failed  = [System.Collections.Generic.List[string]]::new()
    Skipped = [System.Collections.Generic.List[pscustomobject]]::new()   # @{ Name; Reason }
}

# ── Winget upgrade hooks (pre/post per package ID) ──────────────────────────────
# Each key is a winget package ID.  Value is a hashtable with optional Pre / Post scriptblocks.
$script:WingetUpgradeHooks = @{
    'Spotify.Spotify' = @{
        Pre  = {
            $script:_spotifyWasRunning = [bool](Get-Process -Name Spotify -ErrorAction SilentlyContinue)
            Stop-Process -Name Spotify -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        }
        Post = {
            if ($script:_spotifyWasRunning) {
                Start-Process "$env:APPDATA\Spotify\Spotify.exe" -ErrorAction SilentlyContinue
            }
        }
    }
    'Microsoft.VisualStudioCode' = @{
        Pre  = {
            $script:_vscodeWasRunning = [bool](Get-Process -Name 'Code' -ErrorAction SilentlyContinue)
            if ($script:_vscodeWasRunning) {
                Write-Host "  Closing VS Code for upgrade..." -ForegroundColor Gray
                Stop-Process -Name 'Code' -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 3
            }
        }
        Post = {
            if ($script:_vscodeWasRunning) {
                $codePath = Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code\Code.exe'
                if (Test-Path $codePath) { Start-Process $codePath -ErrorAction SilentlyContinue }
            }
        }
    }
    'Microsoft.VisualStudioCode.Insiders' = @{
        Pre  = {
            $script:_vscodeInsidersWasRunning = [bool](Get-Process -Name 'Code - Insiders' -ErrorAction SilentlyContinue)
            if ($script:_vscodeInsidersWasRunning) {
                Write-Host "  Closing VS Code Insiders for upgrade..." -ForegroundColor Gray
                Stop-Process -Name 'Code - Insiders' -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 3
            }
        }
        Post = {
            if ($script:_vscodeInsidersWasRunning) {
                $codePath = Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code Insiders\Code - Insiders.exe'
                if (Test-Path $codePath) { Start-Process $codePath -ErrorAction SilentlyContinue }
            }
        }
    }
}

function Invoke-WingetUpgradeHooks {
    <#
    .SYNOPSIS  Run Pre or Post hooks for any package IDs detected in winget output.
    #>
    param(
        [string]$Phase,          # 'Pre' or 'Post'
        [AllowNull()][string]$WingetOutput
    )
    if (-not $WingetOutput) { return }
    foreach ($pkgId in $script:WingetUpgradeHooks.Keys) {
        if ($WingetOutput -match [regex]::Escape($pkgId)) {
            $hook = $script:WingetUpgradeHooks[$pkgId][$Phase]
            if ($hook) {
                try { & $hook } catch {
                    Write-Status "Hook ($Phase) for $pkgId failed: $_" -Type Warning
                }
            }
        }
    }
}

# ── Logging ──────────────────────────────────────────────────────────────────────
$transcriptStarted = $false
if ($LogPath) {
    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if ($logDir) { New-Item -ItemType Directory -Path $logDir -Force -ErrorAction SilentlyContinue | Out-Null }
        Start-Transcript -Path $LogPath -Append -Force | Out-Null
        $transcriptStarted = $true
    }
    catch {
        Write-Host "WARNING: Could not start transcript at $LogPath. $_" -ForegroundColor Yellow
    }
}

# ── Startup banner ───────────────────────────────────────────────────────────────
Write-Host "`n$('=' * 54)" -ForegroundColor DarkGray
Write-Host "  Update-Everything v2.6.2  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Cyan
Write-Host "  $(if ($isAdmin) { 'Running as Administrator' } else { 'Running as standard user (some tasks skipped)' })" -ForegroundColor $(if ($isAdmin) { 'Green' } else { 'Yellow' })
Write-Host "$('=' * 54)" -ForegroundColor DarkGray

# ── Core update wrapper ──────────────────────────────────────────────────────────
function Invoke-Update {
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$Title,
        [Parameter(Mandatory)][scriptblock]$Action,
        [string]$RequiresCommand,
        [string[]]$RequiresAnyCommand,
        [switch]$Disabled,
        [switch]$RequiresAdmin,
        [switch]$SlowOperation,
        [switch]$NoSection
    )
    if (-not $Title) { $Title = $Name }

    # Helper to record a skip with a reason
    $addSkip = { param([string]$r) $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = $r }) }

    if ($Disabled) { & $addSkip 'flag'; return }
    if ($RequiresCommand -and -not (Test-Command $RequiresCommand)) { & $addSkip 'not installed'; return }
    if ($RequiresAnyCommand -and -not ($RequiresAnyCommand | Where-Object { Test-Command $_ } | Select-Object -First 1)) { & $addSkip 'not installed'; return }
    if ($RequiresAdmin -and -not $isAdmin) { & $addSkip 'requires admin'; return }
    if ($SlowOperation -and $FastMode) { & $addSkip 'fast mode'; return }

    if ($DryRun) {
        Write-Host "  [DryRun] Would run: $Title" -ForegroundColor DarkCyan
        return
    }

    $sectionStart = Get-Date
    if (-not $NoSection) { Write-Section $Title }

    try {
        & $Action
        $elapsed = ((Get-Date) - $sectionStart).TotalSeconds
        Write-Status ("$Name updated ({0:F1}s)" -f $elapsed) -Type Success
        $updateResults.Success.Add($Name)
        $script:sectionTimings[$Name] = $elapsed
    }
    catch {
        Write-Status "$Name failed: $_" -Type Error
        $updateResults.Failed.Add($Name)
    }
}

# ── Parallel update runner ───────────────────────────────────────────────────────
function Invoke-UpdateParallel {
    <#
    .SYNOPSIS
        Runs multiple update sections in parallel using Start-Job.
        Each section must be self-contained (no calls to parent script functions).
        Falls back to sequential if -NoParallel is set or PS version < 7.
    .PARAMETER Sections
        Array of hashtables: @{ Name; Title; RequiresCommand; RequiresAnyCommand; Disabled; SlowOperation; Action }
    #>
    param([hashtable[]]$Sections)

    # Sequential fallback
    if ($NoParallel -or $PSVersionTable.PSVersion.Major -lt 7) {
        foreach ($sec in $Sections) {
            $invokeArgs = @{ Name = $sec.Name; Action = $sec.Action }
            if ($sec.Title) { $invokeArgs['Title'] = $sec.Title }
            if ($sec.RequiresCommand) { $invokeArgs['RequiresCommand'] = $sec.RequiresCommand }
            if ($sec.RequiresAnyCommand) { $invokeArgs['RequiresAnyCommand'] = $sec.RequiresAnyCommand }
            if ($sec.ContainsKey('Disabled')) { $invokeArgs['Disabled'] = $sec.Disabled }
            if ($sec.ContainsKey('SlowOperation')) { $invokeArgs['SlowOperation'] = $sec.SlowOperation }
            Invoke-Update @invokeArgs
        }
        return
    }

    # Parallel path: Start-ThreadJob per section, collect output in order
    # Serialize helper functions so thread jobs can call them
    $fnInit = [scriptblock]::Create(@"
`$commandCache = @{}
function Test-Command([string]`$c) {
    if (`$commandCache.ContainsKey(`$c)) { return `$commandCache[`$c] }
    `$r = [bool](Get-Command `$c -ErrorAction SilentlyContinue)
    `$commandCache[`$c] = `$r; return `$r
}
function Write-Section([string]`$Title, [double]`$ElapsedSeconds = 0) {
    `$e = if (`$ElapsedSeconds -gt 0) { '  ({0:F1}s)' -f `$ElapsedSeconds } else { '' }
    Write-Host "``n`$('=' * 54)" -ForegroundColor DarkGray
    Write-Host " `$Title`$e"      -ForegroundColor Cyan
    Write-Host "`$('=' * 54)"     -ForegroundColor DarkGray
}
function Write-Status([string]`$Message, [ValidateSet('Success','Warning','Error','Info')][string]`$Type='Info') {
    `$c = @{ Success='Green'; Warning='Yellow'; Error='Red'; Info='Gray' }
    `$s = @{ Success='[OK]';  Warning='[!]';    Error='[X]'; Info='[*]' }
    Write-Host "`$(`$s[`$Type]) `$Message" -ForegroundColor `$c[`$Type]
}
function Write-FilteredOutput([AllowNull()][string]`$Text, [ConsoleColor]`$Color = [ConsoleColor]::Gray) {
    if ([string]::IsNullOrWhiteSpace(`$Text)) { return }
    `$n = (`$Text -replace '\x1b\[[0-9;?]*[ -/]*[@-~]', '') -replace '`r', '`n'
    foreach (`$raw in (`$n -split '`n')) {
        `$l = `$raw.TrimEnd(); if (-not `$l) { continue }
        `$c2 = `$l.Trim()
        if (`$c2 -match '^[\\\/\|\-]+`$')                                                            { continue }
        if (`$c2 -match '[█▓▒░▏▎▍▌▋▊▉■□▪▫]')                                                       { continue }
        if (`$c2 -match 'package\(s\) have version numbers that cannot be determined')               { continue }
        Write-Host `$l -ForegroundColor `$Color
    }
}
function Invoke-WingetWithTimeout {
    param([string[]]`$Arguments, [int]`$TimeoutSec = 300)
    `$outFile = [System.IO.Path]::GetTempFileName()
    `$errFile = [System.IO.Path]::GetTempFileName()
    try {
        `$proc = Start-Process -FilePath 'winget' -ArgumentList `$Arguments ``
            -RedirectStandardOutput `$outFile -RedirectStandardError `$errFile -NoNewWindow -PassThru
        `$exited = `$proc.WaitForExit(`$TimeoutSec * 1000)
        if (-not `$exited) { try { `$proc.Kill() } catch { }; throw "winget timed out after `${TimeoutSec}s" }
        `$combined = ((Get-Content -Raw `$outFile -ErrorAction SilentlyContinue) + (Get-Content -Raw `$errFile -ErrorAction SilentlyContinue) -replace '\x00','').Trim()
        return [pscustomobject]@{ Output = `$combined; ExitCode = `$proc.ExitCode }
    } finally { Remove-Item `$outFile, `$errFile -Force -ErrorAction SilentlyContinue }
}
"@)

    $jobs = [ordered]@{}   # Name -> Job
    $jobTitles = @{}            # Name -> Title

    foreach ($sec in $Sections) {
        $name = $sec.Name
        $title = if ($sec.Title) { $sec.Title } else { $name }
        $jobTitles[$name] = $title

        # Apply guards (same as Invoke-Update)
        if ($sec.Disabled) {
            $script:updateResults.Skipped.Add([pscustomobject]@{ Name = $name; Reason = 'flag' })
            continue
        }
        if ($sec.RequiresCommand -and -not (Test-Command $sec.RequiresCommand)) {
            $script:updateResults.Skipped.Add([pscustomobject]@{ Name = $name; Reason = 'not installed' })
            continue
        }
        if ($sec.RequiresAnyCommand) {
            $anyFound = $sec.RequiresAnyCommand | Where-Object { Test-Command $_ } | Select-Object -First 1
            if (-not $anyFound) {
                $script:updateResults.Skipped.Add([pscustomobject]@{ Name = $name; Reason = 'not installed' })
                continue
            }
        }
        if ($sec.SlowOperation -and $FastMode) {
            $script:updateResults.Skipped.Add([pscustomobject]@{ Name = $name; Reason = 'fast mode' })
            continue
        }
        if ($DryRun) {
            Write-Host "  [DryRun] Would run: $title" -ForegroundColor DarkCyan
            continue
        }

        $actionStr = $sec.Action.ToString()   # pass as string; rebuilt with Create() in the thread job
        $secName = $name
        $secTitle = $title
        # Variables the action blocks may reference
        $varTable = @{
            WingetTimeoutSec = $WingetTimeoutSec
            SkipDestructive  = [bool]$SkipDestructive
            pythonCmd        = $pythonCmd
        }
        $jobs[$name] = Start-ThreadJob -Name "update-$name" `
            -InitializationScript $fnInit `
            -ScriptBlock {
            param($actionStr, $secName, $secTitle, $varTable)
            # Inject needed script-scope variables so actions can use them directly
            foreach ($kv in $varTable.GetEnumerator()) { Set-Variable -Name $kv.Key -Value $kv.Value }
            # Reconstruct the scriptblock natively in this runspace so cmdlets resolve correctly
            $action = [ScriptBlock]::Create($actionStr)
            $t = Get-Date
            Write-Section $secTitle
            try {
                & $action
                $elapsed = ((Get-Date) - $t).TotalSeconds
                Write-Status ("{0} updated ({1:F1}s)" -f $secName, $elapsed) -Type Success
                # sentinel on stdout so main thread knows elapsed
                "__TIMING__${secName}__${elapsed}"
            }
            catch {
                Write-Host "[X] $secName failed: $($_.Exception.Message)" -ForegroundColor Red
                "__FAILED__${secName}"
            }
        } -ArgumentList $actionStr, $secName, $secTitle, $varTable
    }

    if ($jobs.Count -eq 0) { return }

    $null = Wait-Job -Job @($jobs.Values) -Timeout 300

    foreach ($name in $jobs.Keys) {
        $job = $jobs[$name]

        # Replay Information-stream items (Write-Host calls from thread job)
        $rawItems = Receive-Job $job -ErrorAction SilentlyContinue 2>&1 6>&1
        foreach ($item in $rawItems) {
            if ($item -is [System.Management.Automation.InformationRecord]) {
                $hm = $item.MessageData
                if ($hm -is [System.Management.Automation.HostInformationMessage]) {
                    # ForegroundColor can be null or -1 in thread jobs — fall back to Gray safely
                    $fg = try {
                        $c = $hm.ForegroundColor
                        if ($null -ne $c -and [int]$c -ge 0 -and [int]$c -le 15) { [ConsoleColor]$c }
                        else { [ConsoleColor]::Gray }
                    }
                    catch { [ConsoleColor]::Gray }
                    try { Write-Host $hm.Message -ForegroundColor $fg -NoNewline:$hm.NoNewLine }
                    catch { Write-Host $hm.Message }
                }
                else {
                    Write-Host $hm.ToString()
                }
            }
            elseif ($item -is [string]) {
                if ($item -match '^__TIMING__(.+)__([0-9.,]+)$') {
                    $sn = $Matches[1]; $el = [double]($Matches[2] -replace ',', '.')
                    $script:updateResults.Success.Add($sn)
                    $script:sectionTimings[$sn] = $el
                }
                elseif ($item -match '^__FAILED__(.+)$') {
                    $script:updateResults.Failed.Add($Matches[1])
                }
                # control strings are not printed
            }
            elseif ($null -ne $item) {
                try { Write-Host ($item | Out-String).TrimEnd() } catch { }
            }
        }

        if ($job.State -eq 'Running') {
            Stop-Job $job -ErrorAction SilentlyContinue
            Write-Status "$name timed out" -Type Warning
            $script:updateResults.Failed.Add($name)
        }
        Remove-Job $job -Force -ErrorAction SilentlyContinue
    }
}

# ════════════════════════════════════════════════════════════
#  PACKAGE MANAGERS
# ════════════════════════════════════════════════════════════

# ── Scoop ────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Scoop' -RequiresCommand 'scoop' -Action {
    $scoopSelfOut = (scoop update 2>&1 | Out-String).Trim()
    if ($scoopSelfOut -match '(?i)error|failed') { Write-Status "Scoop self-update warning: $scoopSelfOut" -Type Warning }
    $out = (scoop update '*' 2>&1 | Out-String).Trim()
    Write-FilteredOutput $out
    scoop cleanup '*' 2>&1 | Out-Null
    scoop cache rm '*' 2>&1 | Out-Null
}

# Packages known to require a reboot to complete — retrying with --force can leave them broken.
$script:WingetRetryBlocklist = @(
    'Microsoft.VCRedist.2015+.x64',
    'Microsoft.VCRedist.2015+.x86',
    'Microsoft.DotNet.Runtime.8',
    'Microsoft.DotNet.Runtime.9',
    'Microsoft.DotNet.DesktopRuntime.8',
    'Microsoft.DotNet.DesktopRuntime.9',
    'Microsoft.WindowsTerminal'
)

# ── Winget ───────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Winget' -RequiresCommand 'winget' -Action {
    $bulkHadFailures = $false
    $retryIds = [System.Collections.Generic.List[string]]::new()

    # Refresh package sources to ensure the latest catalog is available
    Write-Host "  Refreshing winget sources..." -ForegroundColor Gray
    try { Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @('source', 'update') | Out-Null } catch { }

    # Run Pre hooks for packages with known conflicts (e.g. Spotify exit-code 23)
    try {
        $pendingCheck = Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @(
            'upgrade', '--include-unknown', '--source', 'winget', '--accept-source-agreements', '--disable-interactivity'
        )
        Invoke-WingetUpgradeHooks -Phase 'Pre' -WingetOutput $pendingCheck.Output
    }
    catch { }   # best-effort; don't block upgrades if list fails

    # Standard winget source — known-version packages only (safe, no unnecessary reinstalls)
    Write-Host "  Upgrading all (winget source, timeout: ${WingetTimeoutSec}s)..." -ForegroundColor Gray
    try {
        $result = Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
            'upgrade', '--all', '--source', 'winget',
            '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity'
        )
        Write-FilteredOutput $result.Output
        Invoke-WingetUpgradeHooks -Phase 'Post' -WingetOutput $result.Output
        foreach ($id in (Get-WingetFailedPackageIds $result.Output)) {
            if (-not $retryIds.Contains($id)) { $retryIds.Add($id) }
        }
        if ($result.ExitCode -ne 0 -or $result.Output -match 'Installer failed with exit code') { $bulkHadFailures = $true }
    }
    catch {
        Write-Status "Winget (winget source): $_" -Type Warning
        $bulkHadFailures = $true
    }

    # Unknown-version packages — only upgrade when a NEW Available version appears.
    # Winget can't detect the installed version for some apps (e.g. Node.js, ImageMagick).
    # We track which Available version we last upgraded to and skip if unchanged.
    $unknownStateFile = Join-Path (Join-Path $env:LOCALAPPDATA 'Update-Everything') 'unknown_versions.json'
    $unknownState = @{}
    if (Test-Path $unknownStateFile) {
        try { $unknownState = Get-Content $unknownStateFile -Raw | ConvertFrom-Json -AsHashtable } catch { $unknownState = @{} }
    }
    try {
        $unknownCheck = Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @(
            'upgrade', '--include-unknown', '--source', 'winget',
            '--accept-source-agreements', '--disable-interactivity'
        )
        # Parse unknown-version entries: lines where the Version column is "Unknown"
        # Winget IDs always contain dots (publisher.package), so we anchor on that pattern.
        $unknownPkgs = @()
        foreach ($line in ($unknownCheck.Output -split '\r?\n')) {
            if ($line -match '^\s*(.+?)\s+(\S+\.\S+)\s+Unknown\s+(\S+)\s*$') {
                $unknownPkgs += [pscustomobject]@{ Id = $Matches[2]; Available = $Matches[3] }
            }
        }
        if ($unknownPkgs.Count -gt 0) {
            $toUpgrade = @()
            foreach ($pkg in $unknownPkgs) {
                $prev = $unknownState[$pkg.Id]
                if ($prev -ne $pkg.Available) {
                    $toUpgrade += $pkg
                }
            }
            if ($toUpgrade.Count -gt 0) {
                Write-Host "  Upgrading $($toUpgrade.Count) unknown-version package(s)..." -ForegroundColor Gray
                foreach ($pkg in $toUpgrade) {
                    Write-Host "    $($pkg.Id) → $($pkg.Available) (installed version unknown)" -ForegroundColor Gray
                    try {
                        $upResult = Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
                            'upgrade', '--id', $pkg.Id,
                            '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity'
                        )
                        Write-FilteredOutput $upResult.Output
                        # Record the version we just installed
                        $unknownState[$pkg.Id] = $pkg.Available
                    }
                    catch {
                        Write-Status "  Unknown-version upgrade of $($pkg.Id) failed: $_" -Type Warning
                    }
                }
            }
            else {
                Write-Host "  $($unknownPkgs.Count) unknown-version package(s) already at latest available version" -ForegroundColor Gray
            }
            # Save state
            try {
                $dir = Split-Path $unknownStateFile
                New-Item -ItemType Directory -Path $dir -Force -ErrorAction SilentlyContinue | Out-Null
                $unknownState | ConvertTo-Json -Compress | Set-Content -Path $unknownStateFile -Force
            }
            catch { }
        }
    }
    catch { }

    # MS Store source (skipped when -SkipStoreApps)
    if (-not $SkipStoreApps) {
        Write-Host "  Checking MS Store source for upgrades..." -ForegroundColor Gray
        try {
            $storeCheck = Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @(
                'upgrade', '--source', 'msstore',
                '--accept-source-agreements', '--disable-interactivity'
            )
            if ($storeCheck.Output -match 'upgrades available') {
                Write-Host "  Upgrading MS Store apps (timeout: ${WingetTimeoutSec}s)..." -ForegroundColor Gray
                $storeResult = Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
                    'upgrade', '--all', '--source', 'msstore',
                    '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity'
                )
                Write-FilteredOutput $storeResult.Output
                foreach ($id in (Get-WingetFailedPackageIds $storeResult.Output)) {
                    if (-not $retryIds.Contains($id)) { $retryIds.Add($id) }
                }
                if ($storeResult.ExitCode -ne 0 -or $storeResult.Output -match 'Installer failed with exit code') { $bulkHadFailures = $true }
            }
        }
        catch {
            Write-Status "Winget (msstore source): $_" -Type Warning
        }
    }

    # Retry only packages that actually failed during bulk execution.
    if ($retryIds.Count -gt 0) {
        Write-Host "  Retrying $($retryIds.Count) failed package(s) individually..." -ForegroundColor Gray
        foreach ($pkgId in $retryIds) {
            if ($script:WingetRetryBlocklist -contains $pkgId) {
                Write-Status "  Skipping retry for $pkgId (blocklisted — may require reboot)" -Type Warning
                continue
            }
            Write-Host "    Updating $pkgId..." -ForegroundColor Gray
            try {
                Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
                    'upgrade', '--id', $pkgId,
                    '--accept-source-agreements', '--accept-package-agreements',
                    '--disable-interactivity', '--force'
                ) | Out-Null
            }
            catch {
                Write-Status "  Retry of $pkgId timed out or failed: $_" -Type Warning
            }
        }
    }
    elseif ($bulkHadFailures) {
        Write-Status 'Winget reported a bulk failure, but no specific failed package IDs were detected for retry' -Type Warning
    }
}

# Save a winget package snapshot for -WhatChanged diffing
$script:WingetSnapshotDir = Join-Path $env:LOCALAPPDATA 'Update-Everything'
$script:WingetSnapshotCurrent = Join-Path $script:WingetSnapshotDir 'update_log.txt'
$script:WingetSnapshotPrev = Join-Path $script:WingetSnapshotDir 'update_log_prev.txt'
if ((-not $DryRun) -and (Test-Command 'winget')) {
    try {
        New-Item -ItemType Directory -Path $script:WingetSnapshotDir -Force -ErrorAction SilentlyContinue | Out-Null
        # Rotate: current -> prev
        if (Test-Path $script:WingetSnapshotCurrent) {
            Copy-Item -Path $script:WingetSnapshotCurrent -Destination $script:WingetSnapshotPrev -Force
        }
        # Capture current state
        $listResult = Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @('list', '--accept-source-agreements', '--disable-interactivity')
        if ($listResult.Output) {
            Set-Content -Path $script:WingetSnapshotCurrent -Value $listResult.Output -Force
        }
    }
    catch {
        Write-Verbose "Could not save winget snapshot: $_"
    }
}

# ── Sequential: Chocolatey (admin-only) ─────────────────────────────────────────
Invoke-Update -Name 'Chocolatey' -RequiresCommand 'choco' -RequiresAdmin -Action {
    $out = (choco upgrade all -y 2>&1 | Out-String).Trim()
    Write-FilteredOutput $out
}

# ════════════════════════════════════════════════════════════
#  WINDOWS COMPONENTS
# ════════════════════════════════════════════════════════════

# ── Windows Update ───────────────────────────────────────────────────────────────
if (-not $SkipWindowsUpdate) {
    if (-not $isAdmin) {
        $updateResults.Skipped.Add([pscustomobject]@{ Name = 'WindowsUpdate'; Reason = 'requires admin' })
    }
    elseif (Get-Module -ListAvailable -Name PSWindowsUpdate) {
        Invoke-Update -Name 'WindowsUpdate' -Title 'Windows Update' -RequiresAdmin -Action {
            Import-Module PSWindowsUpdate

            # 0x800704c7 (ERROR_CANCELLED) is usually caused by the WU agent aborting
            # due to a pending reboot or -AutoReboot conflicting with the install.
            # Fix: use -IgnoreReboot during install, then handle reboot separately.
            $maxRetries = 2
            $attempt = 0
            $succeeded = $false

            while ($attempt -lt $maxRetries -and -not $succeeded) {
                $attempt++
                try {
                    # IgnoreReboot prevents the WU agent from self-cancelling mid-install.
                    # RecurseCycle handles chained dependencies that require multiple passes.
                    $wuParams = @{
                        Install      = $true
                        AcceptAll    = $true
                        NotCategory  = 'Drivers'
                        IgnoreReboot = $true
                        RecurseCycle = 3
                        Verbose      = $false
                        Confirm      = $false
                    }
                    $results = Get-WindowsUpdate @wuParams

                    if ($results) {
                        Write-Host "  Installed $($results.Count) update(s)." -ForegroundColor Gray
                    }
                    else {
                        Write-Host "  No updates available." -ForegroundColor Gray
                    }
                    $succeeded = $true
                }
                catch {
                    $hresult = $_.Exception.HResult
                    $msg = $_.Exception.Message

                    # 0x800704c7 = -2147023673 (ERROR_CANCELLED)
                    if ($hresult -eq -2147023673 -or $msg -match '0x800704c7') {
                        Write-Status "Windows Update cancelled by system (0x800704c7), attempt $attempt/$maxRetries" -Type Warning

                        if ($attempt -lt $maxRetries) {
                            Write-Host "  Restarting Windows Update service before retry..." -ForegroundColor Gray
                            Restart-Service wuauserv -Force -ErrorAction SilentlyContinue
                            Start-Sleep -Seconds 5
                        }
                    }
                    else {
                        # Unknown error — don't retry
                        throw
                    }
                }
            }

            if (-not $succeeded) {
                throw "Windows Update failed after $maxRetries attempts (0x800704c7). A pending reboot may be required."
            }

            # Handle reboot separately after successful install
            if (-not $SkipReboot) {
                $rebootRequired = (Get-WURebootStatus -Silent -ErrorAction SilentlyContinue)
                if ($rebootRequired) {
                    Write-Status 'A reboot is required to finish installing updates.' -Type Warning
                }
            }
        }
    }
    else {
        Write-Section 'Windows Update'
        Write-Status 'PSWindowsUpdate module not found. Install with: Install-Module PSWindowsUpdate -Force' -Type Warning
        $updateResults.Skipped.Add([pscustomobject]@{ Name = 'WindowsUpdate'; Reason = 'not installed' })
    }
}
else {
    $updateResults.Skipped.Add([pscustomobject]@{ Name = 'WindowsUpdate'; Reason = 'flag' })
}

# ── Microsoft Store Apps ─────────────────────────────────────────────────────────
Invoke-Update -Name 'StoreApps' -Title 'Microsoft Store Apps' -Disabled:$SkipStoreApps -RequiresAdmin -Action {
    try {
        Write-Detail "Checking Microsoft Store app upgrades via winget"
        $result = Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
            'upgrade', '--source', 'msstore', '--all', '--silent',
            '--accept-package-agreements', '--accept-source-agreements'
        )

        if ($result.Output -match 'No installed package found matching input criteria\.') {
            Write-Detail 'No Microsoft Store app upgrades available' -Type Muted
        }
        else {
            Write-FilteredOutput $result.Output
        }

        if ($result.ExitCode -ne 0) {
            Write-Status "winget msstore exited with code $($result.ExitCode)" -Type Warning
        }
    }
    catch {
        throw "StoreApps: $_"
    }
}

# ── WSL ──────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'WSL' -Title 'Windows Subsystem for Linux' -Disabled:$SkipWSL -RequiresCommand 'wsl' -RequiresAdmin -Action {
    # Update the WSL kernel/platform itself (sequential prerequisite)
    $out = (wsl --update 2>&1 | Out-String).Trim() -replace '\x00', ''
    if ($out) { Write-IndentedOutput $out -Prefix '  >' -Color ([ConsoleColor]::Gray) }
    if ($LASTEXITCODE -ne 0) {
        throw "wsl --update failed with exit code $LASTEXITCODE"
    }

    # Optionally update packages inside each distro — in parallel via Start-Job
    if (-not $SkipWSLDistros) {
        $distros = wsl --list --quiet 2>&1 | Where-Object { $_ -and $_ -notmatch '^\s*$' }
        $distroNames = @($distros | ForEach-Object { ($_.Trim() -replace '\x00', '') } | Where-Object { $_ })

        if ($distroNames.Count -eq 0) {
            Write-Detail 'No WSL distros found' -Type Muted
        }
        else {
            Write-Detail "Updating packages in $($distroNames.Count) distro(s) in parallel"

            $jobs = @()
            foreach ($distroName in $distroNames) {
                $jobs += Start-Job -Name "wsl-$distroName" -ArgumentList $distroName -ScriptBlock {
                    param($dn)
                    function Invoke-WslPackageUpdate {
                        param(
                            [string]$DistroName,
                            [string]$Manager,
                            [string]$Command
                        )

                        $result = (wsl -d $DistroName -- sh -lc $Command 2>&1 | Out-String).Trim()
                        $exitCode = $LASTEXITCODE

                        return [pscustomobject]@{
                            Distro = $DistroName
                            Manager = $Manager
                            ExitCode = $exitCode
                            Output = $result
                        }
                    }

                    $attempts = @(
                        @{ Manager = 'apt'; Command = 'if command -v apt-get >/dev/null 2>&1; then sudo apt-get update -qq && sudo apt-get upgrade -y -qq 2>&1 | tail -5; else exit 127; fi' },
                        @{ Manager = 'pacman'; Command = 'if command -v pacman >/dev/null 2>&1; then sudo pacman -Syu --noconfirm 2>&1 | tail -5; else exit 127; fi' },
                        @{ Manager = 'zypper'; Command = 'if command -v zypper >/dev/null 2>&1; then sudo zypper refresh && sudo zypper update -y 2>&1 | tail -5; else exit 127; fi' }
                    )

                    foreach ($attempt in $attempts) {
                        $result = Invoke-WslPackageUpdate -DistroName $dn -Manager $attempt.Manager -Command $attempt.Command
                        if ($result.ExitCode -eq 0) {
                            return [pscustomobject]@{
                                Distro = $dn
                                Success = $true
                                Warning = $false
                                Manager = $attempt.Manager
                                Output = $result.Output
                            }
                        }

                        if ($result.ExitCode -ne 127) {
                            return [pscustomobject]@{
                                Distro = $dn
                                Success = $false
                                Warning = $false
                                Manager = $attempt.Manager
                                Output = $result.Output
                            }
                        }
                    }

                    return [pscustomobject]@{
                        Distro = $dn
                        Success = $true
                        Warning = $true
                        Manager = 'none'
                        Output = 'No supported package manager detected (apt, pacman, zypper)'
                    }
                }
            }

            # Wait for all distro jobs to complete (5 min timeout)
            $null = Wait-Job -Job $jobs -Timeout 300
            $wslFailures = [System.Collections.Generic.List[string]]::new()

            foreach ($job in $jobs) {
                if ($job.State -eq 'Running') {
                    Stop-Job -Job $job -ErrorAction SilentlyContinue
                    Write-Status "  $($job.Name) timed out" -Type Warning
                    $null = $wslFailures.Add($job.Name -replace '^wsl-', '')
                    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                    continue
                }

                $jobResult = Receive-Job -Job $job -ErrorAction SilentlyContinue
                foreach ($result in @($jobResult)) {
                    if ($null -eq $result) { continue }

                    $label = "[$($result.Distro)] $($result.Manager)"
                    if ($result.Success -and -not $result.Warning) {
                        Write-Detail "$label update completed" -Type Info
                        Write-IndentedOutput $result.Output -Prefix '    -' -Color ([ConsoleColor]::DarkGray)
                    }
                    elseif ($result.Warning) {
                        Write-Detail "[$($result.Distro)] No supported package manager detected" -Type Warning
                    }
                    else {
                        Write-Detail "$label update failed" -Type Error
                        Write-IndentedOutput $result.Output -Prefix '    x' -Color ([ConsoleColor]::Red)
                    }

                    if (-not $result.Success) {
                        $null = $wslFailures.Add($result.Distro)
                    }
                }

                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            }

            if ($wslFailures.Count -gt 0) {
                throw "WSL distro update failed for: $($wslFailures -join ', ')"
            }
        }
    }
}

# ── Defender Signatures ──────────────────────────────────────────────────────────
Invoke-Update -Name 'DefenderSignatures' -Title 'Microsoft Defender Signatures' -Disabled:$SkipDefender -RequiresCommand 'Update-MpSignature' -RequiresAdmin -Action {
    $mpStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if ($mpStatus -and -not ($mpStatus.AMServiceEnabled -and $mpStatus.AntivirusEnabled)) {
        Write-Status 'Microsoft Defender is not the active AV; skipping signature update' -Type Info
        return
    }
    try {
        Update-MpSignature -ErrorAction Stop 2>&1 | Out-Null
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -match 'completed with errors') {
            $mpStatusAfter = Get-MpComputerStatus -ErrorAction SilentlyContinue
            if ($mpStatusAfter -and $mpStatusAfter.AntivirusSignatureLastUpdated) {
                if (((Get-Date) - $mpStatusAfter.AntivirusSignatureLastUpdated).TotalDays -lt 2) {
                    Write-Status 'Defender partial source errors, but signatures appear current' -Type Warning
                    return
                }
            }
        }
        throw
    }
}

# ════════════════════════════════════════════════════════════
#  DEVELOPMENT TOOLS (parallel batch)
# ════════════════════════════════════════════════════════════
Write-Host "`n$(if (-not $NoParallel -and $PSVersionTable.PSVersion.Major -ge 7) { '[Parallel]' } else { '[Sequential]' }) Running dev tool updates..." -ForegroundColor DarkCyan

Invoke-UpdateParallel @(
    @{
        Name = 'npm'; Title = 'npm (Node.js)'; RequiresCommand = 'npm'; Disabled = $SkipNode
        Action = {
            $currentNpm = (npm --version 2>&1).Trim()
            $latestNpm = (npm view npm version 2>&1).Trim()
            if ($currentNpm -ne $latestNpm) { npm install -g npm@latest 2>&1 | Out-Null }
            $outdatedJson = (npm outdated -g --json 2>$null | Out-String).Trim()
            if ($outdatedJson) {
                $outdated = $outdatedJson | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($outdated -and $outdated.PSObject.Properties.Count -gt 0) {
                    $pkgs = $outdated.PSObject.Properties.Name
                    Write-Host "  Updating $($pkgs.Count) package(s): $($pkgs -join ', ')" -ForegroundColor Gray
                    npm install -g $pkgs 2>&1 | Out-Null
                }
                else { Write-Host "  All global packages are up to date" -ForegroundColor Gray }
            }
            else { Write-Host "  All global packages are up to date" -ForegroundColor Gray }
            npm cache clean --force 2>&1 | Out-Null
        }
    }
    @{
        Name = 'pnpm'; RequiresCommand = 'pnpm'; Disabled = $SkipNode; SlowOperation = $true
        Action = {
            $out = (pnpm update -g 2>&1 | Out-String).Trim()
            if ($out) { Write-Host $out -ForegroundColor Gray }
        }
    }
    @{
        Name = 'Bun'; RequiresCommand = 'bun'; Disabled = $SkipNode; SlowOperation = $true
        Action = {
            $out = (bun upgrade 2>&1 | Out-String).Trim()
            if ($out) { Write-Host $out -ForegroundColor Gray }
        }
    }
    @{
        Name = 'Deno'; RequiresCommand = 'deno'; Disabled = $SkipNode; SlowOperation = $true
        Action = {
            $denoPath = (Get-Command deno).Source
            if ($denoPath -like '*scoop*') { Write-Host "  Managed by Scoop (already updated)"  -ForegroundColor Gray; return }
            if ($denoPath -like '*WinGet*' -or $denoPath -like '*winget*') { Write-Host "  Managed by winget (already updated)" -ForegroundColor Gray; return }
            $env:NO_COLOR = '1'
            try { $out = (deno upgrade 2>&1 | Out-String).Trim(); if ($out) { Write-Host $out -ForegroundColor Gray } }
            finally { Remove-Item Env:\NO_COLOR -ErrorAction SilentlyContinue }
        }
    }
    @{
        Name = 'Rust'; RequiresCommand = 'rustup'; Disabled = $SkipRust
        Action = {
            $out = (rustup update 2>&1 | Out-String).Trim()
            if ($out) { Write-Host $out -ForegroundColor Gray }
        }
    }
    @{
        Name = 'Go'; RequiresCommand = 'go'; Disabled = $SkipGo
        Action = {
            $goPath = (Get-Command go).Source
            if ($goPath -like '*scoop*') { Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray }
            elseif ($goPath -like '*winget*' -or $goPath -like '*WinGet*') { Write-Host "  Managed by winget (already updated)" -ForegroundColor Gray }
            else {
                try {
                    $r = Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @('upgrade', '--id', 'GoLang.Go', '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity')
                    if ($r.ExitCode -eq 0 -and $r.Output -match 'Successfully installed') { Write-Host "  Updated via winget (GoLang.Go)" -ForegroundColor Gray }
                    else { Write-Host "  No newer version available or not managed by winget" -ForegroundColor Gray }
                }
                catch { Write-Status "Go winget upgrade: $_" -Type Warning }
            }
            if (-not $SkipDestructive) {
                Write-Host "  Cleaning module cache..." -ForegroundColor Gray
                go clean -modcache 2>&1 | Out-Null
            }
        }
    }
    @{
        Name = 'gh-extensions'; Title = 'GitHub CLI Extensions'; RequiresCommand = 'gh'
        Action = {
            $ghExt = (gh extension list 2>&1 | Out-String).Trim()
            if ($ghExt) { gh extension upgrade --all 2>&1 | Out-Null }
            else { Write-Status 'No gh extensions installed' -Type Info }
        }
    }
    @{
        Name = 'pipx'; Title = 'pipx (Python CLI Tools)'; RequiresCommand = 'pipx'
        Action = { pipx upgrade-all 2>&1 | Out-Null }
    }
    @{
        Name = 'Poetry'; Title = 'Poetry (Python Packaging)'; RequiresCommand = 'poetry'; Disabled = $SkipPoetry
        Action = {
            $out = (poetry self update 2>&1 | Out-String).Trim()
            if ($out) { Write-Host $out -ForegroundColor Gray }
        }
    }
    @{
        Name = 'Composer'; Title = 'Composer (PHP)'; RequiresCommand = 'composer'; Disabled = $SkipComposer
        Action = {
            $selfOut = (composer self-update --no-interaction 2>&1 | Out-String).Trim()
            if ($selfOut) { Write-Host $selfOut   -ForegroundColor Gray }
            $globalOut = (composer global update --no-interaction 2>&1 | Out-String).Trim()
            if ($globalOut) { Write-Host $globalOut -ForegroundColor Gray }
        }
    }
    @{
        Name = 'RubyGems'; RequiresCommand = 'gem'; Disabled = $SkipRuby; SlowOperation = $true
        Action = {
            $sysOut = (gem update --system 2>&1 | Out-String).Trim()
            if ($sysOut) { Write-Host $sysOut -ForegroundColor Gray }
            $gemOut = (gem update 2>&1 | Out-String).Trim()
            if ($gemOut) { Write-Host $gemOut -ForegroundColor Gray }
        }
    }
    @{
        Name = 'yt-dlp'; RequiresCommand = 'yt-dlp'
        Action = {
            $ytdlpPath = (Get-Command yt-dlp).Source
            if ($ytdlpPath -like '*scoop*') { Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray }
            elseif ($ytdlpPath -like '*pip*' -or $ytdlpPath -like '*Python*' -or $ytdlpPath -like '*Scripts*') {
                $pCmd = if (Test-Command 'python') { 'python' } else { $null }
                if ($pCmd) { & $pCmd -m pip install --upgrade yt-dlp 2>&1 | Out-Null; Write-Host "  Updated via pip" -ForegroundColor Gray }
                else { Write-Status "yt-dlp installed via pip but Python not found" -Type Warning }
            }
            else {
                $out = (yt-dlp -U 2>&1 | Out-String).Trim()
                if ($out) { Write-Host $out -ForegroundColor Gray }
            }
        }
    }
    @{
        Name = 'tldr'; Title = 'tldr Cache'; RequiresAnyCommand = @('tldr', 'tealdeer')
        Action = {
            $tldrCmd = if (Test-Command 'tealdeer') { 'tealdeer' } else { 'tldr' }
            $out = (& $tldrCmd --update 2>&1 | Out-String).Trim()
            if ($out) { Write-Host $out -ForegroundColor Gray }
        }
    }
    @{
        Name = 'oh-my-posh'; Title = 'Oh My Posh'; RequiresCommand = 'oh-my-posh'
        Action = {
            $ompPath = (Get-Command oh-my-posh).Source
            if ($ompPath -like '*scoop*') { Write-Host "  Managed by Scoop (already updated)"               -ForegroundColor Gray }
            elseif ($ompPath -like '*winget*') { Write-Host "  Managed by winget (already updated)"              -ForegroundColor Gray }
            elseif ($ompPath -like '*WindowsApps*') { Write-Host "  Managed by winget/Store (already updated)"    -ForegroundColor Gray }
            else {
                $out = (oh-my-posh upgrade 2>&1 | Out-String).Trim()
                if ($out -match '(?i)not supported|error|failed') { Write-Status "oh-my-posh upgrade failed: $out" -Type Warning }
                elseif ($out) { Write-Host $out -ForegroundColor Gray }
            }
        }
    }
    @{
        Name = 'Volta'; Title = 'Volta (Node Version Manager)'; RequiresCommand = 'volta'; Disabled = $SkipNode; SlowOperation = $true
        Action = {
            $voltaPath = (Get-Command volta).Source
            if ($voltaPath -like '*scoop*') { Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray }
            else { Write-Status 'Volta installed standalone - update via its installer' -Type Info }
        }
    }
    @{
        Name = 'fnm'; Title = 'fnm (Fast Node Manager)'; RequiresCommand = 'fnm'; Disabled = $SkipNode; SlowOperation = $true
        Action = {
            $fnmPath = (Get-Command fnm).Source
            if ($fnmPath -like '*scoop*') { Write-Host "  Managed by Scoop (already updated)"  -ForegroundColor Gray }
            elseif ($fnmPath -like '*winget*' -or $fnmPath -like '*WinGet*') { Write-Host "  Managed by winget (already updated)" -ForegroundColor Gray }
            else { Write-Status 'fnm installed standalone - update via its installer or scoop' -Type Info }
        }
    }
    @{
        Name = 'mise'; Title = 'mise (Tool Version Manager)'; RequiresCommand = 'mise'; SlowOperation = $true
        Action = {
            $misePath = (Get-Command mise).Source
            if ($misePath -like '*scoop*') { Write-Host "  Managed by Scoop (already updated)"  -ForegroundColor Gray }
            elseif ($misePath -like '*winget*' -or $misePath -like '*WinGet*') { Write-Host "  Managed by winget (already updated)" -ForegroundColor Gray }
            else {
                $out = (mise self-update --yes 2>&1 | Out-String).Trim()
                if ($out) { Write-Host $out -ForegroundColor Gray }
            }
            $pluginsOut = (mise plugins ls 2>&1 | Out-String).Trim()
            if ($pluginsOut) {
                Write-Host "  Upgrading mise plugins..." -ForegroundColor Gray
                mise plugins upgrade 2>&1 | Out-Null
            }
            Write-Host "  Upgrading mise-managed runtimes..." -ForegroundColor Gray
            $upgradeOut = (mise upgrade 2>&1 | Out-String).Trim()
            if ($upgradeOut) { Write-Host $upgradeOut -ForegroundColor Gray }
        }
    }
    @{
        Name = 'juliaup'; Title = 'Julia (juliaup)'; RequiresCommand = 'juliaup'; SlowOperation = $true
        Action = {
            $juliaupPath = (Get-Command juliaup).Source
            if ($juliaupPath -like '*scoop*') { Write-Host "  Managed by Scoop (already updated)"  -ForegroundColor Gray }
            elseif ($juliaupPath -like '*winget*' -or $juliaupPath -like '*WinGet*') { Write-Host "  Managed by winget (already updated)" -ForegroundColor Gray }
            else {
                $out = (juliaup update 2>&1 | Out-String).Trim()
                if ($out) { Write-Host $out -ForegroundColor Gray }
            }
        }
    }
    @{
        Name = 'Ollama Models'; Title = 'Ollama Models'; RequiresCommand = 'ollama'
        Disabled = (-not $UpdateOllamaModels); SlowOperation = $true
        Action = {
            $listOut = (ollama list 2>&1 | Out-String).Trim()
            if (-not $listOut -or $listOut -match '(?i)error') { Write-Status 'Could not list Ollama models' -Type Warning; return }
            $modelNames = ($listOut -split "`n" | Select-Object -Skip 1 | ForEach-Object { ($_ -split '\s+')[0] } | Where-Object { $_ -and $_ -notmatch '^\s*$' })
            if (-not $modelNames -or $modelNames.Count -eq 0) { Write-Host "  No Ollama models installed" -ForegroundColor Gray; return }
            Write-Host "  Pulling updates for $($modelNames.Count) model(s)..." -ForegroundColor Gray
            $updatedCount = 0; $currentCount = 0
            foreach ($model in $modelNames) {
                Write-Host "    Pulling $model..." -ForegroundColor Gray
                $pullOut = (ollama pull $model 2>&1 | Out-String).Trim()
                if ($pullOut -match '(?i)up to date') { $currentCount++ } else { $updatedCount++ }
            }
            Write-Host "  $updatedCount updated, $currentCount already current" -ForegroundColor Gray
        }
    }
    @{
        Name = 'git-lfs'; Title = 'Git LFS'; RequiresCommand = 'git-lfs'; Disabled = $SkipGitLFS
        Action = {
            $lfsPath = (Get-Command git-lfs).Source
            if ($lfsPath -like '*scoop*') { Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray }
            else {
                try {
                    Invoke-WingetWithTimeout -TimeoutSec 120 -Arguments @('upgrade', '--id', 'GitHub.GitLFS', '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity') | Out-Null
                    Write-Host "  Updated via winget (GitHub.GitLFS)" -ForegroundColor Gray
                }
                catch { Write-Status "git-lfs winget upgrade: $_" -Type Warning }
            }
        }
    }
    @{
        Name = 'git-credential-manager'; Title = 'Git Credential Manager'; RequiresCommand = 'git-credential-manager'
        Action = {
            $gcmPath = (Get-Command git-credential-manager).Source
            if ($gcmPath -like '*scoop*') { Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray }
            else {
                try {
                    Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @('upgrade', '--id', 'Git.Git', '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity') | Out-Null
                    Write-Host "  Updated via winget (Git.Git)" -ForegroundColor Gray
                }
                catch { Write-Status "GCM winget upgrade: $_" -Type Warning }
            }
        }
    }
)

# ── Python / pip ─────────────────────────────────────────────────────────────────
$pythonCmd = $null
if (Test-Command 'python') {
    $pythonCmd = 'python'
}
else {
    foreach ($ver in @(314, 313, 312, 311, 310)) {
        $py = "C:\Program Files\Python$ver\python.exe"
        if (Test-Path $py) { $pythonCmd = $py; break }
    }
}
if ($pythonCmd) {
    Invoke-Update -Name 'pip' -Title 'Python / pip' -Action {
        Write-Host "  Using: $pythonCmd" -ForegroundColor Gray
        & $pythonCmd -m pip install --upgrade pip 2>&1 | Out-Null

        # Upgrade outdated global packages
        try {
            $outdatedJson = (& $pythonCmd -m pip list --outdated --format=json 2>$null | Out-String).Trim()
            if ($outdatedJson) {
                $outdated = $outdatedJson | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($outdated -and @($outdated).Count -gt 0) {
                    $pkgNames = @($outdated | ForEach-Object { $_.name })
                    Write-Host "  Upgrading $($pkgNames.Count) outdated package(s): $($pkgNames -join ', ')" -ForegroundColor Gray
                    foreach ($pkg in $pkgNames) {
                        & $pythonCmd -m pip install --upgrade $pkg 2>&1 | Out-Null
                    }
                }
                else {
                    Write-Host "  All global pip packages are up to date" -ForegroundColor Gray
                }
            }
        }
        catch {
            Write-Status "pip global upgrade check failed: $_" -Type Warning
        }
    }
}
else {
    Write-Section 'Python / pip'
    Write-Status 'Python not found' -Type Warning
    $updateResults.Skipped.Add([pscustomobject]@{ Name = 'pip'; Reason = 'not installed' })
}

# ── uv ───────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'uv' -Title 'UV Package Manager' -RequiresCommand 'uv' -Action {
    $uvPath = (Get-Command uv).Source
    if ($uvPath -like "*scoop*") {
        Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray
    }
    elseif ($uvPath -like "*pip*" -or $uvPath -like "*Python*") {
        Write-Host "  Managed by pip (update with: pip install --upgrade uv)" -ForegroundColor Gray
    }
    else {
        $out = (uv self update 2>&1 | Out-String).Trim()
        if ($out -match 'error') {
            Write-Host "  uv self-update not supported; trying winget..." -ForegroundColor Gray
            Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
                'upgrade', '--id', 'astral-sh.uv',
                '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity'
            ) | Out-Null
        }
        elseif ($out) {
            Write-Host $out -ForegroundColor Gray
        }
    }
}

# ── uv tools ─────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'uv-tools' -Title 'uv Tool Installs' -Disabled:$SkipUVTools -RequiresCommand 'uv' -SlowOperation -Action {
    $out = (uv tool upgrade --all 2>&1 | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -and $out -match '(?i)unknown command|unrecognized') {
        Write-Status 'uv tool upgrade --all not supported by this version' -Type Info
        return
    }
    if ($out) { Write-Host $out -ForegroundColor Gray }
}

# ── Claude CLI ────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Claude CLI' -Title 'Claude CLI (Anthropic)' -RequiresCommand 'npm' -Action {
    $currentVer = $null
    try {
        $rawVer = (claude --version 2>&1 | Out-String).Trim()
        # Strip labels like " (Claude Code)" to get the bare semver
        if ($rawVer -match '[\d]+\.[\d]+\.[\d]+') { $currentVer = $Matches[0] }
    }
    catch { }
    $latestVer = (npm show @anthropic-ai/claude-code version 2>&1 | Out-String).Trim()

    if (-not $latestVer -or $latestVer -match '(?i)error|ERR!') {
        Write-Status 'Could not determine latest Claude CLI version from npm' -Type Warning
        return
    }

    if (-not $currentVer -or $currentVer -ne $latestVer) {
        Write-Host "  Updating Claude CLI: $currentVer -> $latestVer" -ForegroundColor Gray

        # Check if the npm global prefix is in a protected location (requires admin)
        $npmPrefix = (npm prefix -g 2>&1 | Out-String).Trim()
        $needsAdmin = (-not $isAdmin) -and ($npmPrefix -like "${env:ProgramFiles}*" -or $npmPrefix -like "${env:ProgramFiles(x86)}*" -or $npmPrefix -like "C:\Program Files*")

        if ($needsAdmin) {
            # Try installing to the user-level npm prefix instead
            $userNpmDir = Join-Path $env:APPDATA 'npm'
            if (-not (Test-Path $userNpmDir)) { New-Item -ItemType Directory -Path $userNpmDir -Force | Out-Null }
            Write-Host "  npm global prefix is admin-protected; installing to user prefix ($userNpmDir)..." -ForegroundColor Gray
            $out = (npm install -g @anthropic-ai/claude-code --prefix "$userNpmDir" 2>&1 | Out-String).Trim()
            if ($out -match '(?i)EPERM|EACCES|error') {
                Write-Status "Claude CLI update requires admin (npm prefix: $npmPrefix). Run elevated or change npm prefix." -Type Warning
            }
            elseif ($out) { Write-Host $out -ForegroundColor Gray }
        }
        else {
            $out = (npm install -g @anthropic-ai/claude-code 2>&1 | Out-String).Trim()
            if ($out) { Write-Host $out -ForegroundColor Gray }
        }
    }
    else {
        Write-Host "  Claude CLI is up to date ($currentVer)" -ForegroundColor Gray
    }
}

# ── uv Python versions ────────────────────────────────────────────────────────────
Invoke-Update -Name 'uv-python' -Title 'uv Python Versions' -Disabled:$SkipUVTools -RequiresCommand 'uv' -SlowOperation -Action {
    $pyList = (uv python list --only-installed 2>&1 | Out-String).Trim()
    if (-not $pyList -or $pyList -match '(?i)no python|error') {
        Write-Status 'No uv-managed Python versions found' -Type Info
        return
    }
    Write-Host "  Upgrading uv-managed Python installs..." -ForegroundColor Gray

    # Extract unique major.minor versions from uv python list output and reinstall
    # to pull the latest patch releases.  Lines look like:
    #   cpython-3.13.2-windows-x86_64-none    C:\Users\...\python.exe
    $majMinors = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($line in ($pyList -split '\r?\n')) {
        if ($line -match 'cpython-(\d+\.\d+)') {
            $null = $majMinors.Add($Matches[1])
        }
    }

    if ($majMinors.Count -eq 0) {
        Write-Host "  No uv-managed cpython installs detected" -ForegroundColor Gray
        return
    }

    $versions = @($majMinors | Sort-Object)
    Write-Host "  Reinstalling latest patches for: $($versions -join ', ')" -ForegroundColor Gray
    $out = (uv python install @versions 2>&1 | Out-String).Trim()
    if ($out) { Write-Host $out -ForegroundColor Gray }
}

# ── Cargo global binaries ────────────────────────────────────────────────────────
if ((Test-Command 'cargo') -and -not $FastMode -and -not $SkipRust) {
    $hasCargoUpdate = $false
    try {
        $null = cargo install-update --version 2>$null
        $hasCargoUpdate = ($LASTEXITCODE -eq 0)
    }
    catch { }

    if ($hasCargoUpdate) {
        Invoke-Update -Name 'cargo-binaries' -Title 'Cargo Global Binaries' -Action {
            $out = (cargo install-update -a 2>&1 | Out-String).Trim()
            if ($out) { Write-Host $out -ForegroundColor Gray }
        }
    }
    else {
        Write-Status 'cargo-update not installed (run: cargo install cargo-update)' -Type Info
        $updateResults.Skipped.Add([pscustomobject]@{ Name = 'cargo-binaries'; Reason = 'not installed' })
    }
}
elseif ($SkipRust) {
    $updateResults.Skipped.Add([pscustomobject]@{ Name = 'cargo-binaries'; Reason = 'flag' })
}
elseif (Test-Command 'cargo') {
    $updateResults.Skipped.Add([pscustomobject]@{ Name = 'cargo-binaries'; Reason = 'fast mode' })
}
else {
    $updateResults.Skipped.Add([pscustomobject]@{ Name = 'cargo-binaries'; Reason = 'not installed' })
}

# ── .NET tools ───────────────────────────────────────────────────────────────────
Invoke-Update -Name 'dotnet' -Title '.NET Tools' -RequiresCommand 'dotnet' -Action {
    # Clean up broken .store entries that cause DirectoryNotFoundException
    $storePath = Join-Path $env:USERPROFILE '.dotnet\tools\.store'
    if (Test-Path $storePath) {
        Get-ChildItem -Path $storePath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $toolDir = $_
            # A valid store entry has a versioned subfolder with a 'tools' directory
            $hasValidLayout = Get-ChildItem -Path $toolDir.FullName -Directory -ErrorAction SilentlyContinue | Where-Object {
                Test-Path (Join-Path $_.FullName "$($toolDir.Name)\$($_.Name)\tools") -ErrorAction SilentlyContinue
            }
            if (-not $hasValidLayout) {
                Write-Host "  Removing broken tool store entry: $($toolDir.Name)" -ForegroundColor Gray
                Remove-Item -Path $toolDir.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    $toolLines = dotnet tool list -g 2>&1 | Select-Object -Skip 2 | Where-Object { $_ -match '\S' }
    if (-not $toolLines) { Write-Host "  No global .NET tools installed" -ForegroundColor Gray; return }

    $updatedCount = 0
    $errorCount = 0
    foreach ($line in $toolLines) {
        $parts = $line -split '\s+', 3
        if ($parts.Count -lt 2) { continue }
        $toolId = $parts[0].Trim()
        $currentVer = $parts[1].Trim()

        try {
            $meta = Invoke-RestMethod "https://api.nuget.org/v3-flatcontainer/$($toolId.ToLower())/index.json" -TimeoutSec 10 -ErrorAction Stop
            $latestVer = $meta.versions | Where-Object { $_ -notmatch '-' } | Select-Object -Last 1
            if (-not $latestVer) { $latestVer = $meta.versions | Select-Object -Last 1 }
        }
        catch { $latestVer = $null }

        if (-not $latestVer -or $latestVer -eq $currentVer) { continue }

        try {
            $out = (dotnet tool update -g $toolId 2>&1 | Out-String).Trim()
            if ($out -match 'Unhandled exception|DirectoryNotFoundException') {
                Write-Status "  $toolId update failed (broken store entry); uninstalling..." -Type Warning
                dotnet tool uninstall -g $toolId 2>&1 | Out-Null
                $out = (dotnet tool install -g $toolId 2>&1 | Out-String).Trim()
            }
            if ($out) { Write-Host $out -ForegroundColor Gray }
            $updatedCount++
        }
        catch {
            Write-Status "  $toolId failed: $_" -Type Warning
            $errorCount++
        }
    }

    if ($updatedCount -eq 0 -and $errorCount -eq 0) { Write-Host "  All .NET tools are up to date" -ForegroundColor Gray }
    if ($errorCount -gt 0) { Write-Status "$errorCount tool(s) had errors" -Type Warning }
}

# ── .NET workloads ───────────────────────────────────────────────────────────────
Invoke-Update -Name 'dotnet-workloads' -Title '.NET Workloads' -RequiresCommand 'dotnet' -Action {
    $workloads = (dotnet workload list 2>&1 | Out-String)
    if ($workloads -notmatch 'No workloads are installed') {
        $out = (dotnet workload update 2>&1 | Out-String).Trim()
        if ($out) {
            $lines = $out -split "`n"
            $filtered = $lines | Where-Object { $_ -notmatch 'Updated advertising manifest' }
            $verbose = $lines | Where-Object { $_ -match 'Updated advertising manifest' }
            if ($verbose) { Write-Verbose ($verbose -join "`n") }
            $filteredText = ($filtered -join "`n").Trim()
            if ($filteredText) { Write-Host $filteredText -ForegroundColor Gray }
        }
    }
    else {
        Write-Status 'No .NET workloads installed' -Type Info
    }
}

# ── VS Code extensions ───────────────────────────────────────────────────────────
Invoke-Update -Name 'vscode-extensions' -Title 'VS Code Extensions' -Disabled:$SkipVSCodeExtensions -RequiresAnyCommand @('code', 'code-insiders', 'code.cmd', 'code-insiders.cmd') -SlowOperation -Action {
    $codeCli = Get-VSCodeCliPath
    if (-not $codeCli) {
        Write-Status 'VS Code CLI shim not found (code.cmd/code-insiders.cmd); skipping extension update to avoid launching UI' -Type Warning
        return
    }
    $out = (& $codeCli --update-extensions 2>&1 | Out-String).Trim()
    if ($out) { Write-Host $out -ForegroundColor Gray }
}

# ── Flutter ───────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Flutter' -Disabled:$SkipFlutter -RequiresCommand 'flutter' -SlowOperation -Action {
    $flutterPath = (Get-Command flutter).Source
    if ($flutterPath -like '*scoop*') {
        Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray
    }
    else {
        $out = (flutter upgrade 2>&1 | Out-String).Trim()
        if ($out) { Write-Host $out -ForegroundColor Gray }
    }
}

# ── PowerShell modules / PSResources ─────────────────────────────────────────────
Invoke-Update -Name 'pwsh-resources' -Title 'PowerShell Modules / Resources' -Disabled:$SkipPowerShellModules -Action {
    $usedProvider = $false

    # PSResourceGet (modern, PS 7.4+) — parallelize per-resource updates
    if ((Test-Command 'Get-InstalledPSResource') -and (Test-Command 'Update-PSResource')) {
        $usedProvider = $true
        $resources = Get-InstalledPSResource -ErrorAction SilentlyContinue
        if ($resources) {
            Write-Host "  Updating $(@($resources).Count) PSResource(s)..." -ForegroundColor Gray
            $_psrCmd = Get-Command Update-PSResource -ErrorAction SilentlyContinue
            $supportsAcceptLicense = if ($_psrCmd) { $_psrCmd.Parameters.ContainsKey('AcceptLicense') } else { $false }

            # Use parallel if PS7 and more than a handful of resources
            if ($PSVersionTable.PSVersion.Major -ge 7 -and @($resources).Count -gt 3) {
                $resources | ForEach-Object -Parallel {
                    $psrArgs = @{ Name = $_.Name; ErrorAction = 'SilentlyContinue' }
                    if ($using:supportsAcceptLicense) { $psrArgs['AcceptLicense'] = $true }
                    try { Update-PSResource @psrArgs 2>&1 | Out-Null } catch { }
                } -ThrottleLimit 4
            }
            else {
                foreach ($resource in $resources) {
                    $psrArgs = @{ Name = $resource.Name; ErrorAction = 'SilentlyContinue' }
                    if ($supportsAcceptLicense) { $psrArgs['AcceptLicense'] = $true }
                    try { Update-PSResource @psrArgs 2>&1 | Out-Null } catch { }
                }
            }
        }
        else {
            Write-Status 'No installed PSResources found' -Type Info
        }
    }

    # PowerShellGet (legacy fallback)
    if ((Test-Command 'Get-InstalledModule') -and (Test-Command 'Update-Module')) {
        $usedProvider = $true
        $modules = Get-InstalledModule -ErrorAction SilentlyContinue
        if ($modules) {
            Write-Host "  Updating $(@($modules).Count) PowerShellGet module(s)..." -ForegroundColor Gray
            $_modCmd = Get-Command Update-Module -ErrorAction SilentlyContinue
            $supportsAcceptLicense = if ($_modCmd) { $_modCmd.Parameters.ContainsKey('AcceptLicense') } else { $false }
            $modules | ForEach-Object {
                $updateArgs = @{ Name = $_.Name; ErrorAction = 'SilentlyContinue' }
                if ($supportsAcceptLicense) { $updateArgs['AcceptLicense'] = $true }
                Update-Module @updateArgs 2>&1 | Out-Null
            }
        }
        else {
            Write-Status 'No installed PowerShellGet modules found' -Type Info
        }
    }

    if (-not $usedProvider) {
        Write-Status 'No PowerShell module update provider found (PSResourceGet or PowerShellGet)' -Type Info
    }
}

# ════════════════════════════════════════════════════════════
#  CLEANUP
# ════════════════════════════════════════════════════════════
if ($SkipCleanup) {
    Write-Section 'System Cleanup'
    Write-Status 'Cleanup skipped (-SkipCleanup flag set)' -Type Info
    $updateResults.Skipped.Add([pscustomobject]@{ Name = 'cleanup'; Reason = 'flag' })
}
elseif ($DryRun) {
    Write-Host "  [DryRun] Would run: System Cleanup" -ForegroundColor DarkCyan
}
else {

    Write-Section 'System Cleanup'

    # Temp files (older than 7 days)
    try {
        $tempPath = $env:TEMP
        if ($tempPath -and (Test-Path $tempPath) -and $tempPath -ne (Split-Path $tempPath -Qualifier)) {
            if ($PSCmdlet.ShouldProcess($tempPath, 'Remove temp files older than 7 days')) {
                $cutoff = (Get-Date).AddDays(-7)
                Get-ChildItem -Path $tempPath -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt $cutoff } |
                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                Write-Status "Temp files cleared (older than 7 days)" -Type Success
            }
        }
    }
    catch {
        Write-Status "Temp cleanup partially failed (normal)" -Type Warning
    }

    # Windows system temp (older than 7 days, admin only)
    if ($isAdmin) {
        try {
            $winTempPath = 'C:\Windows\Temp'
            if (Test-Path $winTempPath) {
                $cutoff = (Get-Date).AddDays(-7)
                Get-ChildItem -Path $winTempPath -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt $cutoff } |
                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                Write-Status "C:\Windows\Temp cleared (older than 7 days)" -Type Success
            }
        }
        catch {
            Write-Status "C:\Windows\Temp cleanup partially failed (normal)" -Type Warning
        }
    }

    # DNS cache
    try {
        Clear-DnsClientCache -ErrorAction SilentlyContinue
        Write-Status "DNS cache flushed" -Type Success
    }
    catch {
        Write-Status "DNS cache flush failed" -Type Warning
    }

    # Recycle Bin
    try {
        if ($PSCmdlet.ShouldProcess('Recycle Bin', 'Empty')) {
            Clear-RecycleBin -Force -ErrorAction SilentlyContinue
            Write-Status "Recycle Bin emptied" -Type Success
        }
    }
    catch {
        Write-Status "Recycle Bin cleanup failed" -Type Warning
    }

    # Crash dumps
    $crashDumpPath = Join-Path $env:LOCALAPPDATA 'CrashDumps'
    if (Test-Path $crashDumpPath) {
        try {
            $dumps = Get-ChildItem -Path $crashDumpPath -ErrorAction SilentlyContinue
            if ($dumps) {
                $dumps | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                Write-Status "Crash dumps cleared ($($dumps.Count) files)" -Type Success
            }
        }
        catch {
            Write-Status "Crash dump cleanup failed" -Type Warning
        }
    }

    # Windows Error Reporting queue (NEW)
    $werPath = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\WER\ReportQueue'
    if (Test-Path $werPath) {
        try {
            Get-ChildItem -Path $werPath -ErrorAction SilentlyContinue |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            Write-Status "WER report queue cleared" -Type Success
        }
        catch {
            Write-Status "WER cleanup failed" -Type Warning
        }
    }

    # Admin-only deep cleanup (opt-in via -DeepClean)
    if ($isAdmin -and $DeepClean) {
        try {
            Write-Host "  Cleaning WinSxS component store..." -ForegroundColor Gray
            DISM /Online /Cleanup-Image /StartComponentCleanup 2>&1 | Out-Null
            Write-Status "DISM component store cleaned" -Type Success
        }
        catch {
            Write-Status "DISM cleanup failed" -Type Warning
        }

        try {
            Clear-DeliveryOptimizationCache -Force -ErrorAction SilentlyContinue
            Write-Status "Delivery Optimization cache cleared" -Type Success
        }
        catch {
            Write-Status "Delivery Optimization cleanup failed" -Type Warning
        }

        # Prefetch cleanup (optional, rarely needed but useful on old HDDs)
        $prefetchPath = 'C:\Windows\Prefetch'
        if ((Test-Path $prefetchPath) -and -not $SkipDestructive) {
            try {
                Get-ChildItem -Path $prefetchPath -Filter '*.pf' -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
                Write-Status "Prefetch files cleared" -Type Success
            }
            catch {
                Write-Status "Prefetch cleanup failed (normal if Superfetch is disabled)" -Type Warning
            }
        }
    }

    $updateResults.Success.Add('cleanup')

} # end -not $SkipCleanup

# ════════════════════════════════════════════════════════════
#  SUMMARY
# ════════════════════════════════════════════════════════════
$duration = (Get-Date) - $startTime
Write-Host "`n$('=' * 54)" -ForegroundColor Green
Write-Host (" UPDATE COMPLETE -- {0}" -f $duration.ToString('hh\:mm\:ss')) -ForegroundColor Green
Write-Host "$('=' * 54)" -ForegroundColor Green

if ($updateResults.Success.Count -gt 0) {
    Write-Host "`n[OK] Succeeded ($($updateResults.Success.Count))" -ForegroundColor Green
    foreach ($name in $updateResults.Success) {
        Write-Detail $name -Type Info
    }
}
if ($updateResults.Failed.Count -gt 0) {
    Write-Host "[X] Failed    ($($updateResults.Failed.Count))" -ForegroundColor Red
    foreach ($name in $updateResults.Failed) {
        Write-Detail $name -Type Error
    }
}
if ($updateResults.Skipped.Count -gt 0) {
    # Known opt-in items (Disabled when their flag is NOT passed, not via -SkipX)
    $optInNames = @('Ollama Models')

    $skippedDisplay = $updateResults.Skipped | ForEach-Object {
        $item = $_
        $label = switch ($item.Reason) {
            'flag' {
                if ($item.Name -in $optInNames) {
                    'opt-in: -Update{0}' -f ($item.Name -replace '\s', '')
                }
                else {
                    'flag: -Skip{0}' -f ($item.Name -replace '\s', '')
                }
            }
            'not installed' { 'not installed' }
            'requires admin' { 'requires admin' }
            'fast mode' { 'fast mode' }
            default { $item.Reason }
        }
        '{0} ({1})' -f $item.Name, $label
    }
    Write-Host "[!] Skipped   ($($updateResults.Skipped.Count))" -ForegroundColor Yellow
    foreach ($item in $skippedDisplay) {
        Write-Detail $item -Type Warning
    }
}

# Per-section timing table
if ($script:sectionTimings.Count -gt 0) {
    Write-Host "`n  Section timings:" -ForegroundColor DarkGray
    $script:sectionTimings.GetEnumerator() |
    Sort-Object Value -Descending |
    ForEach-Object {
        Write-Host ("    {0,-30} {1,6:F1}s" -f $_.Key, $_.Value) -ForegroundColor DarkGray
    }
}

# Completion notification
try {
    if (Get-Module -ListAvailable -Name BurntToast -ErrorAction SilentlyContinue) {
        Import-Module BurntToast -ErrorAction SilentlyContinue
        $successCount = $updateResults.Success.Count
        $failCount = $updateResults.Failed.Count
        $msg = if ($failCount -gt 0) { "$successCount updated, $failCount failed" } else { "$successCount components updated" }
        New-BurntToastNotification -Text 'Update-Everything', $msg -ErrorAction SilentlyContinue
    }
    else {
        [System.Console]::Beep(880, 200)
        Start-Sleep -Milliseconds 50
        [System.Console]::Beep(1100, 300)
    }
}
catch { }

Write-Host ""

# ── WhatChanged diff ─────────────────────────────────────────────────────────────
if ($WhatChanged -and (Test-Path $script:WingetSnapshotCurrent) -and (Test-Path $script:WingetSnapshotPrev)) {
    Write-Section 'What changed since last run'
    try {
        # Parse winget list output into a hashtable of PackageId -> Version
        function ConvertFrom-WingetList {
            param([string]$RawText)
            $map = @{}
            foreach ($line in ($RawText -split "`n")) {
                $trimmed = $line.Trim()
                if (-not $trimmed -or $trimmed -match '^[-=\s]+$' -or $trimmed -match '^Name\s') { continue }
                # winget list columns are variable-width; extract the last two non-empty tokens as Id and Version
                $tokens = $trimmed -split '\s{2,}' | Where-Object { $_ }
                if ($tokens.Count -ge 3) {
                    $id = $tokens[-2].Trim()
                    $ver = $tokens[-1].Trim()
                    if ($id -match '^\S+\.\S+$') { $map[$id] = $ver }
                }
            }
            return $map
        }

        $prevContent = Get-Content -Raw -Path $script:WingetSnapshotPrev
        $currContent = Get-Content -Raw -Path $script:WingetSnapshotCurrent
        $prevMap = ConvertFrom-WingetList $prevContent
        $currMap = ConvertFrom-WingetList $currContent

        $changes = @()
        foreach ($id in $currMap.Keys) {
            if (-not $prevMap.ContainsKey($id)) {
                $changes += "  + $id $($currMap[$id]) (new)"
            }
            elseif ($prevMap[$id] -ne $currMap[$id]) {
                $changes += "  ~ $id $($prevMap[$id]) -> $($currMap[$id])"
            }
        }
        foreach ($id in $prevMap.Keys) {
            if (-not $currMap.ContainsKey($id)) {
                $changes += "  - $id $($prevMap[$id]) (removed)"
            }
        }

        if ($changes.Count -eq 0) {
            Write-Host "  No package changes detected." -ForegroundColor Gray
        }
        else {
            Write-Host "  $($changes.Count) change(s) detected:" -ForegroundColor Cyan
            $changes | ForEach-Object { Write-Host $_ -ForegroundColor Gray }
        }
    }
    catch {
        Write-Status "WhatChanged diff failed: $_" -Type Warning
    }
}
elseif ($WhatChanged) {
    Write-Host "`n[!] No previous snapshot found for comparison. Run the script once without -WhatChanged first." -ForegroundColor Yellow
}

if ($transcriptStarted) {
    try { Stop-Transcript | Out-Null } catch { }
}

if (-not $NoPause -and $AutoElevate) { Read-Host "`nPress Enter to close" }

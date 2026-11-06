<#
.SYNOPSIS
    Re-points Task Scheduler tasks through run-hidden.js so they run with no console window.

.DESCRIPTION
    Dry run by default: prints what would change and touches nothing.

    With -Apply, for each chosen task:
      1. the task is exported to <Root>\backups\<task>.xml, for rollback;
      2. a wrapper <Root>\hidden\<task>.cmd is written that carries the task's original
         command line verbatim;
      3. the task's action is replaced with
             wscript.exe //B "<Root>\run-hidden.js" "<Root>\hidden\<task>.cmd"
         so the job runs inside a hidden console. Its output goes to
         <Root>\hidden\<task>.log and its exit code still reaches Task Scheduler.

    Nothing else about the task changes. Triggers, the user it runs as, "run only when
    user is logged on", the working directory and every setting are kept, so jobs that
    need the desktop (DDE, Outlook COM, a headed browser) keep working.

    Candidates are tasks set to "Run only when user is logged on" (LogonType Interactive)
    whose actions run a console program: powershell, pwsh, cmd, cscript, python, or a
    .cmd/.bat file. pythonw.exe and wscript.exe actions are already windowless and are
    left alone, as is everything under \Microsoft\. An action whose command line would
    be read differently by cmd.exe (an unquoted & | < > ^, or a % that is not part of a
    %NAME% variable) is reported but skipped; convert those by hand.

.PARAMETER TaskName
    One or more task names to convert (exact match, case-insensitive).

.PARAMETER All
    Act on every candidate. Required with -Apply when -TaskName is not given.

.PARAMETER Apply
    Make the changes. Without it the script only reports.

.PARAMETER Root
    Folder holding run-hidden.js. Wrappers, logs and backups go in subfolders.
    Default C:\Automation.

.EXAMPLE
    .\Convert-TasksToHidden.ps1
    Preview every candidate.

.EXAMPLE
    .\Convert-TasksToHidden.ps1 -TaskName 'PricePublish','Job Health Check' -Apply
    Convert two tasks.

.EXAMPLE
    .\Convert-TasksToHidden.ps1 -All -Apply
    Convert everything that is still a candidate. Safe to re-run: converted tasks no
    longer match and are not touched again.

.NOTES
    Run from an elevated PowerShell (Run as administrator) so tasks that use
    "Run with highest privileges" can be modified.

    Roll a task back with:
    Register-ScheduledTask -TaskName 'PricePublish' `
        -Xml (Get-Content 'C:\Automation\backups\PricePublish.xml' -Raw) -Force
#>
[CmdletBinding()]
param(
    [string[]] $TaskName,
    [switch]   $All,
    [switch]   $Apply,
    [string]   $Root = 'C:\Automation'
)

$ErrorActionPreference = 'Stop'

$launcher  = Join-Path $Root 'run-hidden.js'
$hiddenDir = Join-Path $Root 'hidden'
$backupDir = Join-Path $Root 'backups'

if (-not (Test-Path -LiteralPath $launcher)) {
    throw "run-hidden.js was not found at $launcher. Save it there first, or pass -Root."
}
if ($Apply -and -not $All -and -not $TaskName) {
    throw '-Apply needs either -TaskName <names> or -All.'
}

# Console programs whose window we want rid of. pythonw.exe and wscript.exe do not match.
$consoleExe = '(^|\\)(powershell|pwsh|cmd|cscript|python\d*)(\.exe)?$|\.(cmd|bat)$'

function Get-SafeName([string] $Name) {
    ($Name -replace '[\\/:*?"<>|]', '_').Trim()
}

function Test-ExecAction($Action) {
    $Action.CimClass.CimClassName -eq 'MSFT_TaskExecAction'
}

function Test-ConvertibleAction($Action) {
    (Test-ExecAction $Action) -and ($Action.Execute -match $consoleExe)
}

# Task Scheduler hands the command line to the program untouched; inside the wrapper cmd.exe
# reads it first. Flag anything cmd.exe would interpret so those tasks are done by hand.
function Test-CmdHazard([string] $Text) {
    $unquoted = $Text -replace '"[^"]*"', ''
    if ($unquoted -match '[&|<>^]') { return $true }
    # %NAME% is expanded by Task Scheduler and by cmd.exe alike; any other % is a batch-parameter hazard.
    (($Text -replace '%[A-Za-z_][A-Za-z0-9_()]*%', '') -match '%')
}

function Get-WrapperPath([string] $Safe, [int] $Index, [int] $Count) {
    $suffix = if ($Count -gt 1) { "-$Index" } else { '' }
    Join-Path $hiddenDir "$Safe$suffix.cmd"
}

# ---- collect candidates ----------------------------------------------------------------

$tasks = @(Get-ScheduledTask | Where-Object {
    $_.TaskPath -notlike '\Microsoft\*' -and $_.Principal.LogonType -eq 'Interactive'
})

if ($TaskName) {
    $tasks = @($tasks | Where-Object { $TaskName -contains $_.TaskName })
    foreach ($name in $TaskName) {
        if (-not ($tasks | Where-Object TaskName -eq $name)) {
            Write-Warning "No interactive task named '$name' (already converted, not interactive, or misspelt)."
        }
    }
}

$plan = @(foreach ($t in $tasks) {
    $convert = @($t.Actions | Where-Object { Test-ConvertibleAction $_ })
    if (-not $convert) { continue }

    $skip = $null
    foreach ($a in $convert) {
        if (Test-CmdHazard "$($a.Execute) $($a.Arguments) $($a.WorkingDirectory)") {
            $skip = 'command line contains a character cmd.exe would reinterpret - convert by hand'
        }
    }
    [pscustomobject]@{ Task = $t; Convert = $convert; Skip = $skip }
})

if (-not $plan) {
    Write-Host 'Nothing to do: no interactive task runs a console program.'
    return
}

# ---- preview ---------------------------------------------------------------------------

foreach ($p in $plan) {
    $t    = $p.Task
    $safe = Get-SafeName $t.TaskName
    Write-Host ''
    Write-Host ('{0}{1}' -f $t.TaskPath, $t.TaskName) -ForegroundColor Cyan
    $i = 0
    foreach ($a in $p.Convert) {
        $i++
        Write-Host (('   now:  {0} {1}' -f $a.Execute, $a.Arguments).TrimEnd())
        if (-not $p.Skip) {
            $wrapper = Get-WrapperPath $safe $i $p.Convert.Count
            Write-Host ('   new:  wscript.exe //B "{0}" "{1}"' -f $launcher, $wrapper)
            Write-Host ('   log:  {0}' -f ($wrapper -replace '\.cmd$', '.log'))
        }
    }
    if ($p.Skip) {
        Write-Host ('   SKIP: {0}' -f $p.Skip) -ForegroundColor Yellow
    }
}

$ready = @($plan | Where-Object { -not $_.Skip })
Write-Host ''
Write-Host ('{0} task(s) can be converted, {1} need manual attention.' -f $ready.Count, ($plan.Count - $ready.Count))

if (-not $Apply) {
    Write-Host 'Preview only. Re-run with -Apply -All, or -Apply -TaskName <names>, to make the change.'
    return
}

# ---- apply -----------------------------------------------------------------------------

New-Item -ItemType Directory -Force -Path $hiddenDir, $backupDir | Out-Null

foreach ($p in $ready) {
    $t    = $p.Task
    $safe = Get-SafeName $t.TaskName
    try {
        $backup = Join-Path $backupDir "$safe.xml"
        Export-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath |
            Set-Content -LiteralPath $backup -Encoding Unicode

        $n = 0
        $newActions = @(foreach ($a in $t.Actions) {
            if (-not (Test-ConvertibleAction $a)) { $a; continue }

            $n++
            $wrapper = Get-WrapperPath $safe $n $p.Convert.Count
            $exe     = $a.Execute.Trim().Trim('"')

            $lines = @(
                '@echo off',
                ('rem Written by Convert-TasksToHidden.ps1 on {0} for task "{1}{2}".' -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), $t.TaskPath, $t.TaskName),
                ('rem Original task definition: {0}' -f $backup)
            )
            if ($a.WorkingDirectory) { $lines += ('cd /d "{0}"' -f $a.WorkingDirectory) }
            $lines += ('"{0}" {1}' -f $exe, $a.Arguments).TrimEnd()
            $lines += 'exit /b %ERRORLEVEL%'
            Set-Content -LiteralPath $wrapper -Value $lines -Encoding Oem

            $params = @{
                Execute  = 'wscript.exe'
                Argument = ('//B "{0}" "{1}"' -f $launcher, $wrapper)
            }
            if ($a.WorkingDirectory) { $params.WorkingDirectory = $a.WorkingDirectory }
            New-ScheduledTaskAction @params
        })

        Set-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -Action $newActions | Out-Null
        Write-Host ('Converted: {0}' -f $t.TaskName) -ForegroundColor Green
    }
    catch {
        Write-Warning ('FAILED:    {0} - {1}' -f $t.TaskName, $_.Exception.Message)
    }
}

Write-Host ''
Write-Host "Done. Test a task with:  Start-ScheduledTask -TaskName '<name>'   then read its log in $hiddenDir"
Write-Host "Roll one back with:     Register-ScheduledTask -TaskName '<name>' -Xml (Get-Content '$backupDir\<name>.xml' -Raw) -Force"

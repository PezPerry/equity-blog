<#
.SYNOPSIS
    One-shot setup on the RemotePC: fetch the hidden launcher and the converter, smoke-test the
    launcher, preview the tasks, and (with -Apply) convert them.

.DESCRIPTION
    1. Creates the install folder and downloads run-hidden.js and Convert-TasksToHidden.ps1
       from GitHub (PezPerry/equity-blog, tools/windows, at the ref given by -Ref).
    2. Smoke-tests the launcher with a throwaway .cmd: it must run hidden, log its output and
       pass its exit code back. No scheduled task is touched by this step.
    3. Runs the converter in preview mode, or with -Apply converts every candidate task
       (each task is backed up to <Root>\backups first).
    4. Lists the interactive tasks that still run a console program afterwards, which should
       be only the ones the converter marked SKIP.

.PARAMETER Apply
    Convert the tasks. Without it nothing on the machine changes except the files in -Root.

.PARAMETER Ref
    Git ref to download from. Default main; use the branch name while the PR is unmerged.

.PARAMETER Root
    Install folder. Default C:\Automation.

.EXAMPLE
    # From a file, in an administrator PowerShell:
    powershell -ExecutionPolicy Bypass -File .\install.ps1 -Apply

.EXAMPLE
    # Straight from GitHub, no file needed (one line):
    & ([scriptblock]::Create((irm 'https://raw.githubusercontent.com/PezPerry/equity-blog/main/tools/windows/install.ps1'))) -Apply
#>
[CmdletBinding()]
param(
    [switch] $Apply,
    [string] $Ref  = 'main',
    [string] $Root = 'C:\Automation'
)

$ErrorActionPreference = 'Stop'
$base = "https://raw.githubusercontent.com/PezPerry/equity-blog/$Ref/tools/windows"

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole('Administrators')
Write-Host ("Host: {0}   User: {1}   PowerShell: {2}   Elevated: {3}" -f $env:COMPUTERNAME, $env:USERNAME, $PSVersionTable.PSVersion, $isAdmin)
if (-not $isAdmin) {
    Write-Warning 'Not elevated. Tasks that use "Run with highest privileges" may fail to convert; re-run from an administrator PowerShell if any do.'
}

# ---- 1. files ---------------------------------------------------------------------------------
New-Item -ItemType Directory -Force -Path $Root | Out-Null
foreach ($f in 'run-hidden.js', 'Convert-TasksToHidden.ps1') {
    Invoke-WebRequest -UseBasicParsing -Uri "$base/$f" -OutFile (Join-Path $Root $f)
    Write-Host "Downloaded $f -> $Root"
}

# ---- 2. launcher smoke test -------------------------------------------------------------------
$launcher = Join-Path $Root 'run-hidden.js'
$smoke    = Join-Path $Root 'smoke-test.cmd'
$smokeLog = Join-Path $Root 'smoke-test.log'
Set-Content -LiteralPath $smoke -Value @('@echo off', 'echo launcher ok', 'exit /b 7') -Encoding Oem
Remove-Item -LiteralPath $smokeLog -ErrorAction SilentlyContinue

$p = Start-Process -FilePath 'wscript.exe' -ArgumentList ('//B "{0}" "{1}"' -f $launcher, $smoke) -Wait -PassThru
$logText = if (Test-Path -LiteralPath $smokeLog) { Get-Content -LiteralPath $smokeLog -Raw } else { '' }
Remove-Item -LiteralPath $smoke, $smokeLog -ErrorAction SilentlyContinue

if ($p.ExitCode -eq 7 -and $logText -match 'launcher ok') {
    Write-Host 'Launcher smoke test passed: ran hidden, output logged, exit code passed through.' -ForegroundColor Green
} else {
    throw ("Launcher smoke test FAILED (exit code {0}, log '{1}'). No task was changed." -f $p.ExitCode, $logText.Trim())
}

# ---- 3. convert --------------------------------------------------------------------------------
$conv = Join-Path $Root 'Convert-TasksToHidden.ps1'
if ($Apply) { & $conv -Root $Root -All -Apply } else { & $conv -Root $Root }

# ---- 4. what is left ---------------------------------------------------------------------------
Write-Host ''
Write-Host 'Interactive tasks still running a console program (expected: only the ones marked SKIP above):' -ForegroundColor Cyan
$left = @(Get-ScheduledTask | Where-Object { $_.TaskPath -notlike '\Microsoft\*' -and $_.Principal.LogonType -eq 'Interactive' } |
    ForEach-Object { $t = $_; $_.Actions | ForEach-Object {
        [pscustomobject]@{ Task = $t.TaskPath + $t.TaskName; Runs = $_.Execute; Args = $_.Arguments } } } |
    Where-Object { $_.Runs -match '(^|\\)(powershell|pwsh|cmd|cscript|python\d*)(\.exe)?$|\.(cmd|bat)$' })
if ($left) { $left | Format-Table -AutoSize -Wrap } else { Write-Host '(none)' }

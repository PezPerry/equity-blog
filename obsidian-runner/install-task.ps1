# Register a Windows scheduled task that runs the Obsidian website-notes runner
# every 10 minutes while you are logged on (so it uses your git and Claude logins).
#
#   powershell -ExecutionPolicy Bypass -File install-task.ps1
#   powershell -ExecutionPolicy Bypass -File install-task.ps1 -Minutes 5
#   Unregister-ScheduledTask -TaskName "Obsidian Website Runner"   # to remove it

param([int]$Minutes = 10)

$here   = Split-Path -Parent $MyInvocation.MyCommand.Path
$script = Join-Path $here "obsidian_runner.py"
$python = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
if (-not $python) { $python = (Get-Command python.exe -ErrorAction Stop).Source }

if (-not (Test-Path (Join-Path $here "config.json"))) {
    Write-Error "config.json not found. Copy config.example.json to config.json and fill it in first."
    exit 1
}

$action   = New-ScheduledTaskAction -Execute $python -Argument "`"$script`" --once" -WorkingDirectory $here
$trigger  = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes $Minutes)
# Windows defaults a new task to "don't start on battery" and "stop if the PC
# switches to battery", which silently skips runs on a laptop; turn both off.
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
            -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit (New-TimeSpan -Hours 2)

Register-ScheduledTask -TaskName "Obsidian Website Runner" -Action $action -Trigger $trigger `
    -Settings $settings -Description "Pushes new Obsidian website notes to equity-blog via Claude Code" -Force

Write-Host "Scheduled every $Minutes minutes. Log: $(Join-Path $here 'runner.log')"

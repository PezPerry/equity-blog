# Hidden launcher for Task Scheduler jobs (Windows)

Scheduled jobs on the RemotePC run PowerShell, cmd and Python with the task set to
"Run only when user is logged on". Windows opens a console window for each one inside
the interactive session and gives it keyboard focus, which interrupts whatever is being
typed. The two files here make those jobs run with no window while changing nothing
else about them.

| File | Purpose |
|---|---|
| `run-hidden.js` | Launcher. Starts a `.ps1`, `.py`, `.cmd`, `.bat` or `.exe` with its console hidden, logs its output beside it, passes its exit code back to Task Scheduler. |
| `Convert-TasksToHidden.ps1` | Re-points existing tasks through the launcher in bulk. Dry run by default, backs every task up first. |
| `install.ps1` | One-shot runner: downloads the two files above, smoke-tests the launcher, runs the converter, lists what is left. |

## Quickest route: one line on the RemotePC

In an administrator PowerShell on the RemotePC, this downloads both files to `C:\Automation`,
smoke-tests the launcher, previews every candidate task and converts them all (each one is
backed up first). Leave off `-Apply` to preview only.

```powershell
& ([scriptblock]::Create((irm 'https://raw.githubusercontent.com/PezPerry/equity-blog/main/tools/windows/install.ps1'))) -Apply
```

While the PR is unmerged, point it at the branch instead:

```powershell
& ([scriptblock]::Create((irm 'https://raw.githubusercontent.com/PezPerry/equity-blog/claude/powershell-background-windows-epgzi9/tools/windows/install.ps1'))) -Ref claude/powershell-background-windows-epgzi9 -Apply
```

## Manual install (once, on the RemotePC)

```powershell
New-Item -ItemType Directory -Force C:\Automation | Out-Null
$base = 'https://raw.githubusercontent.com/PezPerry/equity-blog/main/tools/windows'
Invoke-WebRequest -UseBasicParsing "$base/run-hidden.js" -OutFile C:\Automation\run-hidden.js
Invoke-WebRequest -UseBasicParsing "$base/Convert-TasksToHidden.ps1" -OutFile C:\Automation\Convert-TasksToHidden.ps1
```

## Convert the tasks

Open PowerShell as administrator, then:

```powershell
cd C:\Automation
powershell -ExecutionPolicy Bypass -File .\Convert-TasksToHidden.ps1                       # preview only
powershell -ExecutionPolicy Bypass -File .\Convert-TasksToHidden.ps1 -TaskName 'PricePublish' -Apply
Start-ScheduledTask -TaskName 'PricePublish'                                              # test it
Get-Content C:\Automation\hidden\PricePublish.log -Tail 20                                # read its output
powershell -ExecutionPolicy Bypass -File .\Convert-TasksToHidden.ps1 -All -Apply          # the rest
```

What the converter does to each task:

1. Exports the task to `C:\Automation\backups\<task>.xml`.
2. Writes `C:\Automation\hidden\<task>.cmd` holding the task's original command line.
3. Sets the task's action to `wscript.exe //B "C:\Automation\run-hidden.js" "C:\Automation\hidden\<task>.cmd"`.

Triggers, the account, "run only when logged on", the working directory and all settings
stay as they were, so jobs that need the desktop (DDE, Outlook COM, a headed browser)
keep working. Output goes to `C:\Automation\hidden\<task>.log`, rolled to `.log.old` at
5 MB. The task's Last Run Result is the job's real exit code.

Tasks whose actions already run `pythonw.exe` or `wscript.exe` are windowless and are
skipped. A task whose command line contains something cmd.exe would reinterpret (an
unquoted `& | < > ^`, or a `%` that is not part of a `%NAME%` variable) is listed as
SKIP; edit that task by hand in Task Scheduler instead.

Roll a task back:

```powershell
Register-ScheduledTask -TaskName 'PricePublish' -Xml (Get-Content 'C:\Automation\backups\PricePublish.xml' -Raw) -Force
```

## Using the launcher for a single job

For a new task, or one done by hand, set the action to:

```
Program:    wscript.exe
Arguments:  //B "C:\Automation\run-hidden.js" "C:\path\to\job.ps1"
Start in:   C:\path\to
```

A `.ps1` runs through `powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File`,
a `.py` through `python.exe`, anything else directly. Extra arguments after the script path
are passed on. Output goes to `job.log` next to the script.

## Jobs that start from pythonw.exe

`pythonw.exe` has no console, so every console program it launches (`powershell.exe`,
`python.exe`, `git.exe`) gets a brand new, visible window. The Task Scheduler change does
not help there; the fix is in the Python script itself. Pasting this near the top makes
every `subprocess` call in that script windowless:

```python
import subprocess, functools

if hasattr(subprocess, "CREATE_NO_WINDOW"):          # Windows only
    _popen_init = subprocess.Popen.__init__

    @functools.wraps(_popen_init)
    def _quiet_init(self, *args, **kwargs):
        kwargs["creationflags"] = kwargs.get("creationflags", 0) | subprocess.CREATE_NO_WINDOW
        _popen_init(self, *args, **kwargs)

    subprocess.Popen.__init__ = _quiet_init
```

`os.system()` is not covered; replace it with `subprocess.run(...)` where it is used.

## If a window still flashes

On Windows 11 with Windows Terminal set as the default terminal, a console can be shown
before the program gets a chance to hide it. Settings > Privacy & security > For developers
> Terminal > choose **Windows Console Host**.

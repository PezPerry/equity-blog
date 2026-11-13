// run-hidden.js - start a script with no console window, log its output, and pass its exit code back.
//
// Task Scheduler action:
//   Program:    wscript.exe
//   Arguments:  //B "C:\Automation\run-hidden.js" "C:\path\to\job.ps1"    (also .py, .cmd, .bat, .exe)
//   Start in:   the folder the job expects to run from
//
// wscript.exe has no window of its own, and the job is started with its console hidden, so
// nothing appears on screen and nothing steals keyboard focus. Anything the job launches in
// turn shares that hidden console. Output is appended to <job>.log beside the job; when the
// log passes 5 MB it is rolled to <job>.log.old first.

var shell = new ActiveXObject("WScript.Shell");
var fso   = new ActiveXObject("Scripting.FileSystemObject");
var args  = WScript.Arguments;
if (args.length === 0) { WScript.Quit(1); }

var target = args.Item(0);
var extra  = "";
for (var i = 1; i < args.length; i++) { extra += ' "' + args.Item(i) + '"'; }

var cmd;
if (/\.ps1$/i.test(target)) {
    cmd = 'powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + target + '"' + extra;
} else if (/\.py$/i.test(target)) {
    cmd = 'python.exe "' + target + '"' + extra;      // put a venv's full python.exe path here if a job needs one
} else {
    cmd = '"' + target + '"' + extra;                 // .cmd, .bat, .exe
}

var log = target.replace(/\.[^.\\]+$/, "") + ".log";
try {
    if (fso.FileExists(log) && fso.GetFile(log).Size > 5 * 1024 * 1024) {
        if (fso.FileExists(log + ".old")) { fso.DeleteFile(log + ".old", true); }
        fso.MoveFile(log, log + ".old");
    }
} catch (e) { /* log in use by another run; keep appending */ }

// Window style 0 = hidden. Waiting keeps the task "Running" in Task Scheduler for as long as the
// job really runs, so instance limits and "stop if it runs longer than" still apply, and the
// job's exit code becomes the task's Last Run Result.
var rc = shell.Run('cmd.exe /c "' + cmd + ' >> "' + log + '" 2>&1"', 0, true);
WScript.Quit(rc);

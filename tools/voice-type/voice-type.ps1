# voice-type.ps1 — launches voice-type.py silently in the background.
# Must be compatible with PowerShell 5.1 (used by wscript.exe launcher).
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PyScript  = Join-Path $ScriptDir "voice-type.py"

# Use 'python' (not 'pythonw' — on Windows Store Python, pythonw.exe is a
# stub that opens the Microsoft Store instead of running a script).
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if ($pythonCmd) {
    $python = $pythonCmd.Source
} else {
    $python = $null
}

if (-not $python) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "Python not found on PATH. Install Python and re-run install.ps1.",
        "voice-type", "OK", "Error"
    )
    exit 1
}

# Kill any existing voice-type instances before launching a fresh one.
# This prevents stale instances (which install keyboard hooks) from
# interfering with Ctrl+C, Ctrl+V, etc.
# Get-CimInstance is reliable on Win10/11; Get-WmiObject silently returns
# nothing on newer builds because CommandLine is null via that provider.
try {
    Get-CimInstance Win32_Process | Where-Object {
        $_.CommandLine -like "*voice-type.py*"
    } | ForEach-Object {
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }
} catch {
    Get-WmiObject Win32_Process | Where-Object {
        $_.CommandLine -like "*voice-type.py*"
    } | ForEach-Object {
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }
}

Start-Process -FilePath $python `
              -ArgumentList "`"$PyScript`"" `
              -WorkingDirectory $ScriptDir `
              -WindowStyle Hidden

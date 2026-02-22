# taskmon.ps1 -- Taskbar system monitor launcher
#
# The C# source lives in taskmon.cs.  Build it once with build.bat, then this
# script just loads the pre-built DLL and starts the UI.  No compilation here.
#
# First-time setup:
#   1. Open a terminal in this folder
#   2. Run:  build.bat
#   3. After that, taskmon.vbs starts instantly every time
#
# After any code change to taskmon.cs:
#   1. Kill the running taskmon (right-click overlay > Quit)
#   2. Run build.bat again
#   3. Relaunch via taskmon.vbs

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -- Single-instance guard -----------------------------------------------------
$_mtx = New-Object System.Threading.Mutex($false, 'Global\TaskMon_SingleInstance')
if (-not $_mtx.WaitOne(0)) {
    [System.Windows.Forms.MessageBox]::Show(
        'taskmon is already running.  Check the taskbar.', 'taskmon',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    exit
}

# -- Load pre-built DLL --------------------------------------------------------
$dll = Join-Path $env:LOCALAPPDATA 'taskmon\taskmon.dll'

if (-not (Test-Path $dll)) {
    $buildBat = Join-Path $PSScriptRoot 'build.bat'
    [System.Windows.Forms.MessageBox]::Show(
        "taskmon has not been built yet.`n`nPlease run build.bat first:`n`n  $buildBat`n`nThis only needs to be done once (and again after any code changes).",
        'taskmon -- not built',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    exit
}

[System.Reflection.Assembly]::LoadFrom($dll) | Out-Null
# Pass the script directory so the C# side can locate taskmon.vbs for the startup registry entry.
[TaskMon.App]::Run($PSScriptRoot)

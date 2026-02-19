# install.ps1 — wires up all tools so they are available on the PATH.
#
# Run once after cloning, and re-run whenever a new tool is added.
# Updating an existing tool only requires a `git pull` — no reinstall needed.
#
# What it does:
#   - Writes thin stub .bat files into $ToolsDir (which should be on your PATH)
#   - Stubs simply forward to the real scripts inside this repo
#   - Recreates the "Scale Monitor 4" taskbar shortcut
#   - Runs each tool's deps.ps1 (if present) to install/check dependencies
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File install.ps1
#   powershell -ExecutionPolicy Bypass -File install.ps1 -SkipDeps

param(
    [switch]$SkipDeps
)

$RepoDir  = Split-Path -Parent $MyInvocation.MyCommand.Path   # auto-resolved; move the repo freely
$ToolsDir = "C:\dev\tools"                                    # directory on your PATH

if (-not (Test-Path $ToolsDir)) {
    New-Item -ItemType Directory -Path $ToolsDir | Out-Null
    Write-Host "Created $ToolsDir" -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# PATH check — warn and offer to fix if $ToolsDir is not on PATH
# ---------------------------------------------------------------------------
$machinePath = [System.Environment]::GetEnvironmentVariable("Path", "Machine"); if (-not $machinePath) { $machinePath = "" }
$userPath    = [System.Environment]::GetEnvironmentVariable("Path", "User");    if (-not $userPath)    { $userPath    = "" }
$onPath      = ($machinePath -split ";") + ($userPath -split ";") |
               Where-Object { $_.TrimEnd("\") -ieq $ToolsDir.TrimEnd("\") }

if (-not $onPath) {
    Write-Host ""
    Write-Host "WARNING: '$ToolsDir' is not on your PATH." -ForegroundColor Yellow
    $ans = Read-Host "  Add it to your User PATH now? [Y/n]"
    if ($ans -eq "" -or $ans -imatch "^y") {
        $newUserPath = ($userPath.TrimEnd(";") + ";$ToolsDir").TrimStart(";")
        [System.Environment]::SetEnvironmentVariable("Path", $newUserPath, "User")
        $env:PATH += ";$ToolsDir"
        Write-Host "  Added '$ToolsDir' to User PATH. Open a new terminal to use the tools." -ForegroundColor Green
    } else {
        Write-Host "  Skipped. Add '$ToolsDir' to PATH manually to use the tools." -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "Installing mikes-windows-tools -> $ToolsDir" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------------------
# Helper: write a stub .bat that calls a target, preserving all arguments
# ---------------------------------------------------------------------------
function Write-BatStub($toolName, $content) {
    $dest = Join-Path $ToolsDir "$toolName.bat"
    Set-Content -Path $dest -Value $content -Encoding ASCII
    Write-Host "  [bat]  $dest" -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# transcribe  — needs EXEDIR so the exe files in c:\dev\tools are found
# ---------------------------------------------------------------------------
Write-BatStub "transcribe" @"
@echo off
set "EXEDIR=%~dp0"
call "$RepoDir\transcribe\transcribe.bat" %*
"@

# ---------------------------------------------------------------------------
# removebg
# ---------------------------------------------------------------------------
Write-BatStub "removebg" @"
@echo off
call "$RepoDir\removebg\removebg.bat" %*
"@

# ---------------------------------------------------------------------------
# all-hands
# ---------------------------------------------------------------------------
Write-BatStub "all-hands" @"
@echo off
call "$RepoDir\all-hands\all-hands.bat" %*
"@

# ---------------------------------------------------------------------------
# backup-phone
# ---------------------------------------------------------------------------
Write-BatStub "backup-phone" @"
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "$RepoDir\backup-phone\backup-phone.ps1" %*
"@

# ---------------------------------------------------------------------------
# scale-monitor4 — taskbar shortcut (no bat stub needed; launched via shortcut)
# ---------------------------------------------------------------------------
$vbsPath      = "$RepoDir\scale-monitor4\scale-monitor4.vbs"
$shortcutPath = Join-Path $ToolsDir "Scale Monitor 4.lnk"
$wsh          = New-Object -ComObject WScript.Shell
$sc           = $wsh.CreateShortcut($shortcutPath)
$sc.TargetPath       = "wscript.exe"
$sc.Arguments        = "`"$vbsPath`""
$sc.WorkingDirectory = "$RepoDir\scale-monitor4"
$sc.Description      = "Toggle Monitor 4 scale between 200% (normal) and 300% (filming)"
$sc.IconLocation     = "%SystemRoot%\System32\imageres.dll,109"
$sc.Save()
Write-Host "  [lnk]  $shortcutPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# taskmon — taskbar system monitor shortcut (launched via VBS for no console window)
# ---------------------------------------------------------------------------
$tmVbsPath      = "$RepoDir\taskmon\taskmon.vbs"
$tmShortcutPath = Join-Path $ToolsDir "Task Monitor.lnk"
$tmSc           = $wsh.CreateShortcut($tmShortcutPath)
$tmSc.TargetPath       = "wscript.exe"
$tmSc.Arguments        = "`"$tmVbsPath`""
$tmSc.WorkingDirectory = "$RepoDir\taskmon"
$tmSc.Description      = "Taskbar system monitor: NET / CPU / GPU / MEM sparklines"
$tmSc.IconLocation     = "%SystemRoot%\System32\imageres.dll,174"
$tmSc.Save()
Write-Host "  [lnk]  $tmShortcutPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# voice-type — taskbar shortcut (launched via VBS for no console window)
# ---------------------------------------------------------------------------
$vtVbsPath      = "$RepoDir\voice-type\voice-type.vbs"
$vtShortcutPath = Join-Path $ToolsDir "Voice Type.lnk"
$vtSc           = $wsh.CreateShortcut($vtShortcutPath)
$vtSc.TargetPath       = "wscript.exe"
$vtSc.Arguments        = "`"$vtVbsPath`""
$vtSc.WorkingDirectory = "$RepoDir\voice-type"
$vtSc.Description      = "Push-to-talk voice typing: hold Right Ctrl to record, release to transcribe and paste"
$vtSc.IconLocation     = "%SystemRoot%\System32\imageres.dll,109"
$vtSc.Save()
Write-Host "  [lnk]  $vtShortcutPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Dependencies — run each tool's deps.ps1 if present
# ---------------------------------------------------------------------------
if ($SkipDeps) {
    Write-Host "Skipping dependency checks (-SkipDeps was set)." -ForegroundColor DarkGray
} else {
    Write-Host ""
    Write-Host "Checking / installing tool dependencies..." -ForegroundColor Cyan

    $depsScripts = Get-ChildItem -Path $RepoDir -Recurse -Filter "deps.ps1" |
        Where-Object { $_.FullName -ne (Join-Path $RepoDir "deps.ps1") }

    if ($depsScripts.Count -eq 0) {
        Write-Host "  (no deps.ps1 files found)" -ForegroundColor DarkGray
    } else {
        foreach ($script in $depsScripts | Sort-Object FullName) {
            & $script.FullName
        }
    }
}

# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "Done. To update tools in future: git pull (no reinstall needed)." -ForegroundColor Yellow
Write-Host "To add a new tool: create its subfolder, then re-run install.ps1." -ForegroundColor Yellow
Write-Host "To skip dependency checks: install.ps1 -SkipDeps" -ForegroundColor Yellow
Write-Host ""
Write-Host "Reminder: right-click 'Scale Monitor 4.lnk' in $ToolsDir and pin to taskbar." -ForegroundColor Cyan
Write-Host "Reminder: right-click 'Task Monitor.lnk' in $ToolsDir and pin to taskbar." -ForegroundColor Cyan
Write-Host "Reminder: right-click 'Voice Type.lnk' in $ToolsDir and pin to taskbar (or run on login)." -ForegroundColor Cyan

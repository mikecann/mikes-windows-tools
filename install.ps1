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
# ghopen — open current repo/PR in browser
# ---------------------------------------------------------------------------
Write-BatStub "ghopen" @"
@echo off
call "$RepoDir\ghopen\ghopen.bat" %*
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
# Context menu - "Mike's Tools" submenu on video files in File Explorer
# Registered per-extension under SystemFileAssociations (more reliable than
# the perceived-type path which Explorer doesn't always pick up).
# ---------------------------------------------------------------------------
Write-Host "  [reg]  Registering 'Mike''s Tools' context menu for video files..." -ForegroundColor Green

# Convert a PNG to a 16x16 .ico file and return the output path.
# The registry Icon value requires .ico (or .dll,index) - raw PNG is not supported.
function ConvertTo-Ico($pngPath, $icoPath) {
    Add-Type -AssemblyName System.Drawing
    $bmp = [System.Drawing.Bitmap]::new($pngPath)
    # Resize to 16x16 - the standard shell menu icon size
    $small = [System.Drawing.Bitmap]::new(16, 16, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($small)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($bmp, 0, 0, 16, 16)
    $g.Dispose()
    $hicon = $small.GetHicon()
    $icon  = [System.Drawing.Icon]::FromHandle($hicon)
    $stream = [System.IO.FileStream]::new($icoPath, [System.IO.FileMode]::Create)
    $icon.Save($stream)
    $stream.Close()
    $icon.Dispose()
    $small.Dispose()
    $bmp.Dispose()
}

$iconsOut = "$env:LOCALAPPDATA\mikes-windows-tools\icons"
New-Item -ItemType Directory -Force $iconsOut | Out-Null

$wrenchIco = "$iconsOut\mikes-tools.ico"
$filmIco   = "$iconsOut\transcribe.ico"
ConvertTo-Ico "$RepoDir\transcribe\icons\wrench.png" $wrenchIco
ConvertTo-Ico "$RepoDir\transcribe\icons\film.png"   $filmIco
Write-Host "  [ico]  Icons written to $iconsOut" -ForegroundColor Green

$transcribeCmd = 'cmd.exe /k ""C:\dev\tools\transcribe.bat" "%1""'

$videoExts = @('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.webm', '.m4v', '.mpg', '.mpeg', '.ts', '.mts', '.m2ts', '.flv', '.f4v')

foreach ($ext in $videoExts) {
    $menuRoot  = "HKCU:\Software\Classes\SystemFileAssociations\$ext\shell\MikesTools"
    $transcKey = "$menuRoot\shell\Transcribe"
    $cmdKey    = "$transcKey\command"

    New-Item -Path $menuRoot  -Force | Out-Null
    New-Item -Path $transcKey -Force | Out-Null
    New-Item -Path $cmdKey    -Force | Out-Null

    Set-ItemProperty -Path $menuRoot  -Name "MUIVerb"     -Value "Mike's Tools"
    Set-ItemProperty -Path $menuRoot  -Name "SubCommands" -Value ""
    Set-ItemProperty -Path $menuRoot  -Name "Icon"        -Value $wrenchIco
    Set-ItemProperty -Path $transcKey -Name "MUIVerb"     -Value "Transcribe Video"
    Set-ItemProperty -Path $transcKey -Name "Icon"        -Value $filmIco
    Set-ItemProperty -Path $cmdKey    -Name "(Default)"   -Value $transcribeCmd
}

# Notify the shell that file associations changed so Explorer picks it up
# without needing a manual restart.
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class ShellNotify {
    [DllImport("shell32.dll")]
    public static extern void SHChangeNotify(int wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
}
'@
[ShellNotify]::SHChangeNotify(0x08000000, 0x0000, [IntPtr]::Zero, [IntPtr]::Zero)

Write-Host "  [reg]  Done. Right-click a video in Explorer to see Mike's Tools." -ForegroundColor Green

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

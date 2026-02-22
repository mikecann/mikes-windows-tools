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
Write-Host "Installing mikerosoft.app -> $ToolsDir" -ForegroundColor Cyan
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
call "$RepoDir\tools\transcribe\transcribe.bat" %*
"@

# ---------------------------------------------------------------------------
# removebg
# ---------------------------------------------------------------------------
Write-BatStub "removebg" @"
@echo off
call "$RepoDir\tools\removebg\removebg.bat" %*
"@

# ---------------------------------------------------------------------------
# all-hands
# ---------------------------------------------------------------------------
Write-BatStub "all-hands" @"
@echo off
call "$RepoDir\tools\all-hands\all-hands.bat" %*
"@

# ---------------------------------------------------------------------------
# ghopen — open current repo/PR in browser
# ---------------------------------------------------------------------------
Write-BatStub "ghopen" @"
@echo off
call "$RepoDir\tools\ghopen\ghopen.bat" %*
"@

# ---------------------------------------------------------------------------
# ctxmenu — context menu manager GUI
# ---------------------------------------------------------------------------
Write-BatStub "ctxmenu" @"
@echo off
wscript.exe "$RepoDir\tools\ctxmenu\ctxmenu.vbs"
"@

# ---------------------------------------------------------------------------
# backup-phone
# ---------------------------------------------------------------------------
Write-BatStub "backup-phone" @"
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "$RepoDir\tools\backup-phone\backup-phone.ps1" %*
"@

# ---------------------------------------------------------------------------
# copypath — copies current or specified path to clipboard
# ---------------------------------------------------------------------------
Write-BatStub "copypath" @"
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "$RepoDir\tools\copypath\copypath.ps1" %*
"@

# ---------------------------------------------------------------------------
# vid2md — YouTube URL to markdown clipboard
# ---------------------------------------------------------------------------
Write-BatStub "vid2md" @"
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -Sta -File "$RepoDir\tools\vid2md\vid2md.ps1" %*
"@

# ---------------------------------------------------------------------------
# scale-monitor4 — taskbar shortcut (no bat stub needed; launched via shortcut)
# ---------------------------------------------------------------------------
$vbsPath      = "$RepoDir\tools\scale-monitor4\scale-monitor4.vbs"
$shortcutPath = Join-Path $ToolsDir "Scale Monitor 4.lnk"
$wsh          = New-Object -ComObject WScript.Shell
$sc           = $wsh.CreateShortcut($shortcutPath)
$sc.TargetPath       = "wscript.exe"
$sc.Arguments        = "`"$vbsPath`""
$sc.WorkingDirectory = "$RepoDir\tools\scale-monitor4"
$sc.Description      = "Toggle Monitor 4 scale between 200% (normal) and 300% (filming)"
$sc.IconLocation     = "%SystemRoot%\System32\imageres.dll,109"
$sc.Save()
Write-Host "  [lnk]  $shortcutPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# taskmon — bat stub (terminal launch) + taskbar shortcut
# ---------------------------------------------------------------------------
Write-BatStub "taskmon" @"
@echo off
wscript.exe "$RepoDir\tools\taskmon\taskmon.vbs"
"@

# taskmon shortcut
$tmVbsPath      = "$RepoDir\tools\taskmon\taskmon.vbs"
$tmShortcutPath = Join-Path $ToolsDir "Task Monitor.lnk"
$tmSc           = $wsh.CreateShortcut($tmShortcutPath)
$tmSc.TargetPath       = "wscript.exe"
$tmSc.Arguments        = "`"$tmVbsPath`""
$tmSc.WorkingDirectory = "$RepoDir\tools\taskmon"
$tmSc.Description      = "Taskbar system monitor: NET / CPU / GPU / MEM sparklines"
$tmSc.IconLocation     = "%SystemRoot%\System32\imageres.dll,174"
$tmSc.Save()
Write-Host "  [lnk]  $tmShortcutPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# voice-type — taskbar shortcut (launched via VBS for no console window)
# ---------------------------------------------------------------------------
$vtVbsPath      = "$RepoDir\tools\voice-type\voice-type.vbs"
$vtShortcutPath = Join-Path $ToolsDir "Voice Type.lnk"
$vtSc           = $wsh.CreateShortcut($vtShortcutPath)
$vtSc.TargetPath       = "wscript.exe"
$vtSc.Arguments        = "`"$vtVbsPath`""
$vtSc.WorkingDirectory = "$RepoDir\tools\voice-type"
$vtSc.Description      = "Push-to-talk voice typing: hold Right Ctrl to record, release to transcribe and paste"
$vtSc.IconLocation     = "%SystemRoot%\System32\imageres.dll,109"
$vtSc.Save()
Write-Host "  [lnk]  $vtShortcutPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Context menu - "Mike's Tools" submenu in File Explorer
# Covers: video files (transcribe), image files (removebg), folders (ghopen)
# ---------------------------------------------------------------------------
Write-Host "  [reg]  Registering 'Mike''s Tools' context menus..." -ForegroundColor Green

# Convert a PNG to an .ico using PNG-in-ICO format (Vista+).
# Embeds raw PNG bytes into the ICO container to preserve full alpha
# transparency. GetHicon()/Icon.FromHandle() fills alpha with black.
function ConvertTo-Ico($pngPath, $icoPath) {
    $pngBytes = [System.IO.File]::ReadAllBytes($pngPath)
    $stream   = [System.IO.FileStream]::new($icoPath, [System.IO.FileMode]::Create)
    $w        = [System.IO.BinaryWriter]::new($stream)
    $w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]1)   # ICONDIR
    $w.Write([byte]16);  $w.Write([byte]16);  $w.Write([byte]0)     # ICONDIRENTRY width/height/colorcount
    $w.Write([byte]0);   $w.Write([uint16]1); $w.Write([uint16]32)  # reserved/planes/bitcount
    $w.Write([uint32]$pngBytes.Length); $w.Write([uint32]22)        # data size / offset
    $w.Write($pngBytes)
    $w.Close(); $stream.Close()
}

# Helper: ensure a "Mike's Tools" submenu root exists at $rootKey with $icon.
function Set-MikesToolsRoot($rootKey, $icon) {
    New-Item -Path $rootKey -Force | Out-Null
    Set-ItemProperty -Path $rootKey -Name "MUIVerb"     -Value "Mike's Tools"
    Set-ItemProperty -Path $rootKey -Name "SubCommands" -Value ""
    Set-ItemProperty -Path $rootKey -Name "Icon"        -Value $icon
}

# Helper: add a verb entry + command under an existing Mike's Tools root.
function Add-MikesVerb($rootKey, $verbName, $label, $icon, $command) {
    $verbKey = "$rootKey\shell\$verbName"
    $cmdKey  = "$verbKey\command"
    New-Item -Path $verbKey -Force | Out-Null
    New-Item -Path $cmdKey  -Force | Out-Null
    Set-ItemProperty -Path $verbKey -Name "MUIVerb" -Value $label
    Set-ItemProperty -Path $verbKey -Name "Icon"    -Value $icon
    Set-ItemProperty -Path $cmdKey  -Name "(Default)" -Value $command
}

$iconsOut = "$env:LOCALAPPDATA\mikerosoft.app\icons"
New-Item -ItemType Directory -Force $iconsOut | Out-Null

$wrenchIco  = "$iconsOut\mikes-tools.ico"
$filmIco    = "$iconsOut\transcribe.ico"
$pictureIco = "$iconsOut\removebg.ico"
$worldIco   = "$iconsOut\ghopen.ico"
$linkPageIco = "$iconsOut\vid2md.ico"
ConvertTo-Ico "$RepoDir\tools\transcribe\icons\wrench.png"       $wrenchIco
ConvertTo-Ico "$RepoDir\tools\transcribe\icons\film.png"         $filmIco
ConvertTo-Ico "$RepoDir\tools\removebg\icons\picture.png"        $pictureIco
ConvertTo-Ico "$RepoDir\tools\ghopen\icons\world_go.png"         $worldIco
ConvertTo-Ico "$RepoDir\tools\vid2md\icons\page_white_link.png"  $linkPageIco
Write-Host "  [ico]  Icons written to $iconsOut" -ForegroundColor Green

# --- transcribe + vid2md: video file extensions ---
$videoExts = @('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.webm', '.m4v', '.mpg', '.mpeg', '.ts', '.mts', '.m2ts', '.flv', '.f4v')
foreach ($ext in $videoExts) {
    $root = "HKCU:\Software\Classes\SystemFileAssociations\$ext\shell\MikesTools"
    Set-MikesToolsRoot $root $wrenchIco
    Add-MikesVerb $root "Transcribe" "Transcribe Video"   $filmIco    'cmd.exe /k ""C:\dev\tools\transcribe.bat" "%1""'
    Add-MikesVerb $root "Vid2md"    "Video to Markdown"  $linkPageIco "powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File `"$RepoDir\tools\vid2md\vid2md.ps1`" `"%1`""
}

# --- removebg: image file extensions ---
$imageExts = @('.jpg', '.jpeg', '.png', '.webp', '.bmp', '.tiff', '.tif')
foreach ($ext in $imageExts) {
    $root = "HKCU:\Software\Classes\SystemFileAssociations\$ext\shell\MikesTools"
    Set-MikesToolsRoot $root $wrenchIco
    Add-MikesVerb $root "RemoveBg" "Remove Background" $pictureIco 'cmd.exe /k ""C:\dev\tools\removebg.bat" "%1""'
}

# --- vid2md: Internet Shortcut files (.url) - YouTube links ---
$urlRoot = "HKCU:\Software\Classes\SystemFileAssociations\.url\shell\MikesTools"
Set-MikesToolsRoot $urlRoot $wrenchIco
Add-MikesVerb $urlRoot "Vid2md" "Video to Markdown" $linkPageIco "powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File `"$RepoDir\tools\vid2md\vid2md.ps1`" `"%1`""

# --- ghopen + vid2md: folders (right-click on folder icon) and folder background ---
$vid2mdCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File `"$RepoDir\tools\vid2md\vid2md.ps1`" `"%1`""

# Directory - right-clicking a folder item; %1 = folder path
$dirRoot = "HKCU:\Software\Classes\Directory\shell\MikesTools"
Set-MikesToolsRoot $dirRoot $wrenchIco
Add-MikesVerb $dirRoot "GhOpen" "Open on GitHub"   $worldIco  'cmd.exe /k "cd /d "%1" && "C:\dev\tools\ghopen.bat""'
Add-MikesVerb $dirRoot "Vid2md" "Video to Markdown" $linkPageIco $vid2mdCmd

# Directory\Background - right-clicking inside an open folder; %V = current folder
$bgRoot = "HKCU:\Software\Classes\Directory\Background\shell\MikesTools"
Set-MikesToolsRoot $bgRoot $wrenchIco
Add-MikesVerb $bgRoot "GhOpen" "Open on GitHub"   $worldIco  'cmd.exe /k "cd /d "%V" && "C:\dev\tools\ghopen.bat""'
Add-MikesVerb $bgRoot "Vid2md" "Video to Markdown" $linkPageIco "powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File `"$RepoDir\tools\vid2md\vid2md.ps1`""

# --- vid2md: all files and folders via AllFilesystemObjects ---
# AllFilesystemObjects is a Windows shell class that matches every file and folder.
# It has no literal '*' in the registry path so PowerShell's -Path handles it safely.
# Windows merges these verbs with any more-specific SystemFileAssociations entries
# (e.g. video files also get Transcribe from their own registration).
$afoRoot = "HKCU:\Software\Classes\AllFilesystemObjects\shell\MikesTools"
Set-MikesToolsRoot $afoRoot $wrenchIco
Add-MikesVerb $afoRoot "Vid2md" "Video to Markdown" $linkPageIco $vid2mdCmd

# Notify the shell so Explorer picks up all changes without a manual restart.
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class ShellNotify {
    [DllImport("shell32.dll")]
    public static extern void SHChangeNotify(int wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
}
'@
[ShellNotify]::SHChangeNotify(0x08000000, 0x0000, [IntPtr]::Zero, [IntPtr]::Zero)

Write-Host "  [reg]  Done. Mike's Tools in context menus for videos, images, and folders." -ForegroundColor Green

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

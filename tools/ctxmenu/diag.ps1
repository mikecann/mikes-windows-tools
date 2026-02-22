# ctxmenu-diag.ps1
# Dumps every registry location that contributes to the context menu for a given extension.
# Run:  powershell -ExecutionPolicy Bypass -File diag.ps1 -Ext .mp4
# Output is written to diag-<ext>.txt next to this script.

param([string]$Ext = '.mp4')

$ext = $Ext.ToLower()
if (-not $ext.StartsWith('.')) { $ext = ".$ext" }

$out = [System.Text.StringBuilder]::new()
function W([string]$s) { [void]$out.AppendLine($s) }
function WH([string]$s) { W ''; W "=== $s ==="; W '' }

W  "Context menu diagnostic for: $ext"
W  "Run at: $(Get-Date)"
W  "---"

# Helper: read a registry key by full path (HKCR, HKLM, HKCU)
function OpenKey([string]$fullPath) {
    $parts = $fullPath -split '\\', 2
    if ($parts.Count -lt 2) { return $null }
    $root = switch ($parts[0]) {
        'HKEY_CLASSES_ROOT'  { [Microsoft.Win32.Registry]::ClassesRoot }
        'HKEY_LOCAL_MACHINE' { [Microsoft.Win32.Registry]::LocalMachine }
        'HKEY_CURRENT_USER'  { [Microsoft.Win32.Registry]::CurrentUser }
        default { return $null }
    }
    return $root.OpenSubKey($parts[1], $false)
}
function VerbLabel([Microsoft.Win32.RegistryKey]$k) {
    if (-not $k) { return '?' }
    $v = $k.GetValue('MUIVerb')
    if (-not $v) { $v = $k.GetValue('') }
    if (-not $v) { $v = $k.Name.Split('\')[-1] }
    return "$v"
}
function DumpShell([string]$fullPath) {
    $k = OpenKey $fullPath
    if (-not $k) { W "  (key not found: $fullPath)"; return }
    $names = $k.GetSubKeyNames()
    if ($names.Count -eq 0) { W "  (empty)"; $k.Close(); return }
    foreach ($n in $names) {
        $sk = $k.OpenSubKey($n)
        $label = VerbLabel $sk
        $icon  = if ($sk) { $sk.GetValue('Icon') } else { '' }
        $cmd   = ''
        if ($sk) {
            $ck = $sk.OpenSubKey('command')
            if ($ck) { $cmd = "$($ck.GetValue(''))"; $ck.Close() }
        }
        if ($sk) { $sk.Close() }
        W  "  [$n]  label='$label'  icon='$icon'"
        if ($cmd) { W "    cmd: $cmd" }
    }
    $k.Close()
}
function DumpShellEx([string]$fullPath) {
    $k = OpenKey $fullPath
    if (-not $k) { W "  (key not found: $fullPath)"; return }
    $names = $k.GetSubKeyNames()
    if ($names.Count -eq 0) { W "  (empty)"; $k.Close(); return }
    foreach ($n in $names) {
        $sk = $k.OpenSubKey($n)
        $clsid = if ($sk) { "$($sk.GetValue(''))" } else { '' }
        if ($sk) { $sk.Close() }
        # Try to resolve CLSID to a friendly name
        $clsClean = $clsid.TrimStart('-')
        $friendly = ''
        if ($clsClean -match '^\{') {
            $ck = OpenKey "HKEY_CLASSES_ROOT\CLSID\$clsClean"
            if ($ck) { $friendly = "$($ck.GetValue(''))"; $ck.Close() }
        }
        W  "  [$n]  clsid='$clsid'  friendly='$friendly'"
    }
    $k.Close()
}

# ── 1. What ProgID does this extension map to? ───────────────────────────────
WH "1. Extension ProgID mapping"
foreach ($hive in @('HKEY_CLASSES_ROOT','HKEY_LOCAL_MACHINE','HKEY_CURRENT_USER')) {
    $sub = if ($hive -eq 'HKEY_LOCAL_MACHINE') { 'SOFTWARE\Classes' } `
           elseif ($hive -eq 'HKEY_CURRENT_USER') { 'Software\Classes' } `
           else { '' }
    $path = if ($sub) { "$hive\$sub\$ext" } else { "$hive\$ext" }
    $k = OpenKey $path
    if ($k) {
        $def = $k.GetValue('')
        W "  ${hive}: default='$def'"
        $owi = $k.OpenSubKey('OpenWithProgids')
        if ($owi) {
            W "  OpenWithProgids:"
            foreach ($n in $owi.GetValueNames()) { W "    $n" }
            $owi.Close()
        }
        $k.Close()
    } else {
        W "  ${hive}: (not found)"
    }
}

# ── 2. HKCR\*\shell  (All Files - verbs) ─────────────────────────────────────
WH "2. HKCR\*\shell  (All Files verbs)"
DumpShell 'HKEY_CLASSES_ROOT\*\shell'

# ── 3. HKCR\*\shellex\ContextMenuHandlers  (All Files - COM) ─────────────────
WH "3. HKCR\*\shellex\ContextMenuHandlers  (All Files COM)"
DumpShellEx 'HKEY_CLASSES_ROOT\*\shellex\ContextMenuHandlers'

# ── 4. SystemFileAssociations\<ext>\shell ────────────────────────────────────
WH "4. HKCR\SystemFileAssociations\$ext\shell"
DumpShell "HKEY_CLASSES_ROOT\SystemFileAssociations\$ext\shell"

WH "4b. HKCU SystemFileAssociations\$ext\shell"
DumpShell "HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\$ext\shell"

WH "4c. HKLM SystemFileAssociations\$ext\shell"
DumpShell "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\SystemFileAssociations\$ext\shell"

# ── 5. Perceived type (video / image) ────────────────────────────────────────
$perceivedType = ''
$ptKey = OpenKey "HKEY_CLASSES_ROOT\SystemFileAssociations\$ext"
if ($ptKey) {
    $perceivedType = "$($ptKey.GetValue('PerceivedType'))"; $ptKey.Close()
}
if (-not $perceivedType) {
    $perceivedType = switch -wildcard ($ext) {
        '.mp4' { 'video' } '.mkv' { 'video' } '.avi' { 'video' } '.mov' { 'video' }
        '.jpg' { 'image' } '.png' { 'image' } '.jpeg' { 'image' }
        default { '' }
    }
}
W "PerceivedType: $perceivedType"

if ($perceivedType) {
    WH "5. HKCR\SystemFileAssociations\$perceivedType\shell"
    DumpShell "HKEY_CLASSES_ROOT\SystemFileAssociations\$perceivedType\shell"

    WH "5b. HKLM SystemFileAssociations\$perceivedType\shell"
    DumpShell "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\SystemFileAssociations\$perceivedType\shell"

    WH "5c. HKCU SystemFileAssociations\$perceivedType\shell"
    DumpShell "HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\$perceivedType\shell"
}

# ── 6. ProgID shell verbs ─────────────────────────────────────────────────────
# Get all ProgIDs for this extension from HKCR
$progIds = [System.Collections.Generic.List[string]]::new()
foreach ($hive in @('HKEY_CLASSES_ROOT','HKEY_LOCAL_MACHINE','HKEY_CURRENT_USER')) {
    $sub = if ($hive -eq 'HKEY_LOCAL_MACHINE') { 'SOFTWARE\Classes' } `
           elseif ($hive -eq 'HKEY_CURRENT_USER') { 'Software\Classes' } `
           else { '' }
    $path = if ($sub) { "$hive\$sub\$ext" } else { "$hive\$ext" }
    $k = OpenKey $path
    if ($k) {
        $def = $k.GetValue('')
        if ($def -and -not $progIds.Contains($def)) { $progIds.Add($def) }
        $owi = $k.OpenSubKey('OpenWithProgids')
        if ($owi) {
            foreach ($n in $owi.GetValueNames()) {
                if ($n -and -not $progIds.Contains($n)) { $progIds.Add($n) }
            }
            $owi.Close()
        }
        $k.Close()
    }
}
WH "6. ProgID shell verbs  (ProgIDs: $($progIds -join ', '))"
foreach ($progId2 in $progIds) {
    W "  -- ProgID: $progId2 --"
    DumpShell "HKEY_CLASSES_ROOT\$progId2\shell"
}

# Also dump Clipchamp AppX keys specifically
WH "6b. Clipchamp AppX shell verbs"
foreach ($name in @('AppX77a4xf1yjnhecq6pjngc07j31r3q912','AppXyhj8780c19x3fmfxz91fzp1yzmr5wae0')) {
    W "  -- $name --"
    DumpShell "HKEY_CLASSES_ROOT\$name\shell"
    $sk = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($name)
    if ($sk) {
        $sup = $sk.OpenSubKey('SupportedTypes')
        if ($sup) { W "  SupportedTypes: $($sup.GetValueNames() -join ', ')"; $sup.Close() }
        $sk.Close()
    }
}

# ── 7. HKCR top-level keys that start with "AppX" (UWP app registrations) ──
WH "7. AppX/UWP shell verbs for $ext"
# UWP apps register under HKCR\AppX<hash>\ keys
# Look for AppX keys that have shell\open or similar for this extension
$hkcr = [Microsoft.Win32.Registry]::ClassesRoot
$appxMatches = [System.Collections.Generic.List[string]]::new()
try {
    foreach ($name in $hkcr.GetSubKeyNames()) {
        if ($name.StartsWith('AppX', [System.StringComparison]::OrdinalIgnoreCase)) {
            $ak = $hkcr.OpenSubKey($name)
            if (-not $ak) { continue }
            # Check if this AppX key has shell subkeys
            $shellk = $ak.OpenSubKey('shell')
            if ($shellk) {
                $hasExt = $false
                # Check SupportedTypes or capabilities
                $capk = $ak.OpenSubKey('SupportedTypes')
                if ($capk) {
                    if ($capk.GetValueNames() -icontains $ext) { $hasExt = $true }
                    $capk.Close()
                }
                if ($hasExt -or ($appxMatches.Count -lt 20)) {
                    # Dump if relevant or up to 20 for discovery
                    $appName = "$($ak.GetValue(''))"
                    W "  AppX key: $name  name='$appName'"
                    if ($hasExt) {
                        W "  -> supports $ext!"
                        DumpShell "HKEY_CLASSES_ROOT\$name\shell"
                    }
                }
                $shellk.Close()
            }
            $ak.Close()
        }
    }
} catch { W "  Error scanning AppX: $_" }

# ── 8. HKCU\Software\Classes - user-level registrations ─────────────────────
WH "8. HKCU\Software\Classes\$ext\shell  (user-level)"
DumpShell "HKEY_CURRENT_USER\Software\Classes\$ext\shell"

WH "8b. HKCU\Software\Classes\$ext\shellex\ContextMenuHandlers"
DumpShellEx "HKEY_CURRENT_USER\Software\Classes\$ext\shellex\ContextMenuHandlers"

# ── Done ──────────────────────────────────────────────────────────────────────
$result = $out.ToString()
$outFile = Join-Path $PSScriptRoot "diag$ext.txt"
$result | Set-Content $outFile -Encoding UTF8
Write-Host "Written to: $outFile"
Write-Host "--- summary ($($result.Split("`n").Count) lines) ---"

# ctxmenu.ps1 - Windows Explorer context menu manager
#
# Shows shell verbs and COM extension handlers from the registry.
# Toggle entries on/off using HKCU shadow keys - no admin rights needed,
# because Windows merges HKCU on top of HKLM when building HKCR.
#
# Disable mechanisms:
#   Static verbs   - add LegacyDisable (REG_SZ, empty) to the verb key
#   COM handlers   - prefix the CLSID value with '-' in the handler key

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

Add-Type -ReferencedAssemblies 'System.Drawing' @'
using System;
using System.Drawing;
using System.Runtime.InteropServices;

public class CmEntry {
    public string VerbName;     // registry key name
    public string Label;        // friendly display name
    public string AppliesTo;    // "All Files", "Folders", "Video Files", etc.
    public string Source;       // "HKCU" or "HKLM"
    public string Kind;         // "Verb", "Submenu", or "ShellEx"
    public string ReadPath;     // full HKEY_xxx\... path for reading
    public string ShadowPath;   // HKCU subkey path to write disable value
    public bool   Enabled;
    public bool   IsSubmenu;
    public string ClsId;        // ShellEx only - clean {GUID} without leading -
}

public class IconUtil {
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    private static extern int ExtractIconEx(
        string lpszFile, int nIconIndex,
        IntPtr[] phiconLarge, IntPtr[] phiconSmall, int nIcons);

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);

    public static Icon GetSmall(string file, int index) {
        try {
            var large = new IntPtr[1];
            var small = new IntPtr[1];
            if (ExtractIconEx(file, index, large, small, 1) == 0) return null;
            if (large[0] != IntPtr.Zero) DestroyIcon(large[0]);
            if (small[0] == IntPtr.Zero) return null;
            var icon = (Icon)Icon.FromHandle(small[0]).Clone();
            DestroyIcon(small[0]);
            return icon;
        } catch { return null; }
    }
}
'@

# ── Registry helpers ──────────────────────────────────────────────────────────
function rOpen([string]$fullPath) {
    $parts = $fullPath -split '\\', 2
    if ($parts.Count -lt 2) { return $null }
    $root = switch ($parts[0]) {
        'HKEY_LOCAL_MACHINE' { [Microsoft.Win32.Registry]::LocalMachine }
        'HKEY_CURRENT_USER'  { [Microsoft.Win32.Registry]::CurrentUser  }
        'HKEY_CLASSES_ROOT'  { [Microsoft.Win32.Registry]::ClassesRoot  }
        default              { return $null }
    }
    if (-not $root) { return $null }
    return $root.OpenSubKey($parts[1], $false)
}

function rLabel([Microsoft.Win32.RegistryKey]$k) {
    foreach ($name in @('MUIVerb', '')) {
        $v = $k.GetValue($name)
        if ($v -and [string]$v -ne '') { return [string]$v }
    }
    return $k.Name.Split('\')[-1]
}

function hkuShadow([string]$hive, [string]$subPath) {
    if ($hive -eq 'HKCU') { return $subPath }
    return $subPath -replace '^SOFTWARE\\Classes\\', 'Software\Classes\'
}

function isVerbDisabled([string]$readPath, [string]$shadow) {
    foreach ($p in @($readPath, "HKEY_CURRENT_USER\$shadow")) {
        $k = rOpen $p
        if ($k) {
            $dis = $k.GetValueNames() -icontains 'LegacyDisable'
            $k.Close()
            if ($dis) { return $true }
        }
    }
    return $false
}

function isShellExDisabled([string]$readPath, [string]$shadow) {
    foreach ($p in @("HKEY_CURRENT_USER\$shadow", $readPath)) {
        $k = rOpen $p
        if ($k) {
            $v = [string]$k.GetValue('')
            $k.Close()
            if ($v) { return $v.StartsWith('-') }
        }
    }
    return $false
}

# ── Scanners ──────────────────────────────────────────────────────────────────
function scanVerbs([string]$hive, [string]$subPath, [string]$appliesTo) {
    $results = [System.Collections.Generic.List[CmEntry]]::new()
    $hiveWord = if ($hive -eq 'HKLM') { 'HKEY_LOCAL_MACHINE' } else { 'HKEY_CURRENT_USER' }
    $shell = rOpen "$hiveWord\$subPath"
    if (-not $shell) { return $results }

    $shBase = hkuShadow $hive $subPath

    foreach ($verb in $shell.GetSubKeyNames()) {
        try {
            $vk = $shell.OpenSubKey($verb)
            if (-not $vk) { continue }
            $label     = rLabel $vk
            $isSubmenu = $null -ne $vk.GetValue('SubCommands')
            $vk.Close()

            $e = [CmEntry]::new()
            $e.VerbName   = $verb
            $e.Label      = $label
            $e.AppliesTo  = $appliesTo
            $e.Source     = $hive
            $e.Kind       = if ($isSubmenu) { 'Submenu' } else { 'Verb' }
            $e.ReadPath   = "$hiveWord\$subPath\$verb"
            $e.ShadowPath = "$shBase\$verb"
            $e.Enabled    = -not (isVerbDisabled $e.ReadPath $e.ShadowPath)
            $e.IsSubmenu  = $isSubmenu
            $results.Add($e)
        } catch { }
    }
    $shell.Close()
    return $results
}

function scanShellEx([string]$hive, [string]$subPath, [string]$appliesTo) {
    $results = [System.Collections.Generic.List[CmEntry]]::new()
    $hiveWord = if ($hive -eq 'HKLM') { 'HKEY_LOCAL_MACHINE' } else { 'HKEY_CURRENT_USER' }
    $handlers = rOpen "$hiveWord\$subPath"
    if (-not $handlers) { return $results }

    $shBase = hkuShadow $hive $subPath

    foreach ($name in $handlers.GetSubKeyNames()) {
        try {
            $hk = $handlers.OpenSubKey($name)
            if (-not $hk) { continue }
            $clsidRaw = [string]$hk.GetValue('')
            $hk.Close()
            if (-not $clsidRaw) { continue }

            $clsidClean = $clsidRaw.TrimStart('-')
            $label = $name
            $ck = rOpen "HKEY_CLASSES_ROOT\CLSID\$clsidClean"
            if ($ck) {
                $fn = [string]$ck.GetValue('')
                if ($fn) { $label = "$name  [$fn]" }
                $ck.Close()
            }

            $e = [CmEntry]::new()
            $e.VerbName   = $name
            $e.Label      = $label
            $e.AppliesTo  = $appliesTo
            $e.Source     = $hive
            $e.Kind       = 'ShellEx'
            $e.ReadPath   = "$hiveWord\$subPath\$name"
            $e.ShadowPath = "$shBase\$name"
            $e.ClsId      = $clsidClean
            $e.Enabled    = -not (isShellExDisabled $e.ReadPath $e.ShadowPath)
            $results.Add($e)
        } catch { }
    }
    $handlers.Close()
    return $results
}

function getExtEntries([string[]]$exts, [string]$typeName) {
    $seen    = [System.Collections.Generic.Dictionary[string,CmEntry]]::new()
    $shadows = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[string]]]::new()

    foreach ($ext in $exts) {
        foreach ($hive in @('HKCU', 'HKLM')) {
            $hiveWord = if ($hive -eq 'HKLM') { 'HKEY_LOCAL_MACHINE' } else { 'HKEY_CURRENT_USER' }
            $sub  = if ($hive -eq 'HKLM') { "SOFTWARE\Classes\SystemFileAssociations\$ext\shell" } `
                    else                   { "Software\Classes\SystemFileAssociations\$ext\shell" }
            $shell = rOpen "$hiveWord\$sub"
            if (-not $shell) { continue }

            foreach ($verb in $shell.GetSubKeyNames()) {
                try {
                    $vk = $shell.OpenSubKey($verb)
                    if (-not $vk) { continue }
                    $label     = rLabel $vk
                    $isSubmenu = $null -ne $vk.GetValue('SubCommands')
                    $vk.Close()

                    $shPath = "Software\Classes\SystemFileAssociations\$ext\shell\$verb"

                    if (-not $seen.ContainsKey($verb)) {
                        $e = [CmEntry]::new()
                        $e.VerbName   = $verb
                        $e.Label      = $label
                        $e.AppliesTo  = $typeName
                        $e.Source     = $hive
                        $e.Kind       = if ($isSubmenu) { 'Submenu' } else { 'Verb' }
                        $e.ReadPath   = "$hiveWord\$sub\$verb"
                        $e.ShadowPath = $shPath
                        $e.Enabled    = -not (isVerbDisabled $e.ReadPath $e.ShadowPath)
                        $e.IsSubmenu  = $isSubmenu
                        $seen[$verb]    = $e
                        $shadows[$verb] = [System.Collections.Generic.List[string]]::new()
                    }
                    $shadows[$verb].Add($shPath)
                } catch { }
            }
            $shell.Close()
        }
    }

    foreach ($verb in $seen.Keys) {
        $seen[$verb].ShadowPath = ($shadows[$verb] | Sort-Object -Unique) -join ';'
    }
    return $seen.Values
}

function getAllEntries {
    $all   = [System.Collections.Generic.List[CmEntry]]::new()
    $addAll = { param($col) foreach ($e in $col) { if ($e) { $all.Add($e) } } }

    @(
        @('HKLM','SOFTWARE\Classes\*\shell',                    'All Files'),
        @('HKCU','Software\Classes\*\shell',                    'All Files'),
        @('HKLM','SOFTWARE\Classes\Directory\shell',            'Folders'),
        @('HKCU','Software\Classes\Directory\shell',            'Folders'),
        @('HKLM','SOFTWARE\Classes\Directory\Background\shell', 'Folder Background'),
        @('HKCU','Software\Classes\Directory\Background\shell', 'Folder Background'),
        @('HKLM','SOFTWARE\Classes\Drive\shell',                'Drives'),
        @('HKCU','Software\Classes\Drive\shell',                'Drives')
    ) | ForEach-Object { & $addAll (scanVerbs $_[0] $_[1] $_[2]) }

    @(
        @('HKLM','SOFTWARE\Classes\*\shellex\ContextMenuHandlers',                    'All Files'),
        @('HKCU','Software\Classes\*\shellex\ContextMenuHandlers',                    'All Files'),
        @('HKLM','SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers',            'Folders'),
        @('HKCU','Software\Classes\Directory\shellex\ContextMenuHandlers',            'Folders'),
        @('HKLM','SOFTWARE\Classes\Directory\Background\shellex\ContextMenuHandlers', 'Folder Background'),
        @('HKCU','Software\Classes\Directory\Background\shellex\ContextMenuHandlers', 'Folder Background')
    ) | ForEach-Object { & $addAll (scanShellEx $_[0] $_[1] $_[2]) }

    $videoExts = @('.mp4','.mkv','.avi','.mov','.wmv','.webm','.m4v','.mpg','.mpeg','.ts','.mts','.m2ts','.flv','.f4v')
    $imageExts = @('.jpg','.jpeg','.png','.webp','.bmp','.tiff','.tif')
    & $addAll (getExtEntries $videoExts 'Video Files')
    & $addAll (getExtEntries $imageExts 'Image Files')

    return $all
}

# ── Apply enable/disable ──────────────────────────────────────────────────────
function applyEntry([CmEntry]$entry, [bool]$enable) {
    $hkcu = [Microsoft.Win32.Registry]::CurrentUser

    if ($entry.Kind -ne 'ShellEx') {
        foreach ($shadow in ($entry.ShadowPath -split ';')) {
            try {
                $k = $hkcu.OpenSubKey($shadow, $true)
                if (-not $k -and -not $enable) { $k = $hkcu.CreateSubKey($shadow) }
                if ($k) {
                    if ($enable) { try { $k.DeleteValue('LegacyDisable') } catch { } }
                    else         { $k.SetValue('LegacyDisable', '', [Microsoft.Win32.RegistryValueKind]::String) }
                    $k.Close()
                }
            } catch { }
        }
    } else {
        try {
            $k = $hkcu.OpenSubKey($entry.ShadowPath, $true)
            if (-not $k) { $k = $hkcu.CreateSubKey($entry.ShadowPath) }
            if ($k) {
                $val = if ($enable) { $entry.ClsId } else { "-$($entry.ClsId)" }
                $k.SetValue('', $val, [Microsoft.Win32.RegistryValueKind]::String)
                $k.Close()
            }
        } catch { }
    }
}

function notifyShell {
    Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public class CtxShell {
    [DllImport("shell32.dll")]
    public static extern void SHChangeNotify(int e, uint f, IntPtr a, IntPtr b);
}
'@ -ErrorAction SilentlyContinue
    try { [CtxShell]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero) } catch { }
}

# ── Icon helpers ──────────────────────────────────────────────────────────────
$script:imgCache = [System.Collections.Generic.Dictionary[string,int]]::new()

function initImageList {
    $il = New-Object System.Windows.Forms.ImageList
    $il.ImageSize  = New-Object System.Drawing.Size(16, 16)
    $il.ColorDepth = 'Depth32Bit'

    # Load fallback famfamfam icons by index (0=verb/wrench, 1=submenu/go, 2=shellex/cog)
    $root = $PSScriptRoot
    foreach ($rel in @(
        '..\transcribe\icons\wrench.png',   # 0 - Verb
        '..\taskmon\icons\bullet_go.png',   # 1 - Submenu
        '..\taskmon\icons\cog.png'          # 2 - ShellEx
    )) {
        $p = Join-Path $root $rel
        try {
            if (Test-Path $p) {
                $bmp = [System.Drawing.Bitmap]::new($p)
                $il.Images.Add($bmp)
                $bmp.Dispose()
            } else {
                $il.Images.Add([System.Drawing.Bitmap]::new(16,16))
            }
        } catch { $il.Images.Add([System.Drawing.Bitmap]::new(16,16)) }
    }
    return $il
}

function fallbackIndex([CmEntry]$e) {
    if ($e.Kind -eq 'ShellEx')  { return 2 }
    if ($e.Kind -eq 'Submenu')  { return 1 }
    return 0
}

function getIconIndex([CmEntry]$entry, [System.Windows.Forms.ImageList]$il) {
    # 1. Try Icon value on the verb key
    $iconSpec = $null
    $vk = rOpen $entry.ReadPath
    if ($vk) {
        $iconSpec = $vk.GetValue('Icon')
        $vk.Close()
    }

    # 2. For ShellEx with no Icon value, try InprocServer32 path
    if (-not $iconSpec -and $entry.Kind -eq 'ShellEx' -and $entry.ClsId) {
        $ipk = rOpen "HKEY_CLASSES_ROOT\CLSID\$($entry.ClsId)\InprocServer32"
        if ($ipk) {
            $dll = [string]$ipk.GetValue('')
            $ipk.Close()
            if ($dll) { $iconSpec = "$dll,0" }
        }
    }

    if (-not $iconSpec) { return fallbackIndex $entry }

    $cacheKey = [string]$iconSpec
    if ($script:imgCache.ContainsKey($cacheKey)) { return $script:imgCache[$cacheKey] }

    try {
        $raw = [System.Environment]::ExpandEnvironmentVariables([string]$iconSpec).Trim()
        # Skip MUI localized string refs (@file,-resId) - not icon paths
        if ($raw.StartsWith('@')) {
            $script:imgCache[$cacheKey] = fallbackIndex $entry
            return fallbackIndex $entry
        }

        $idx   = 0
        $comma = $raw.LastIndexOf(',')
        $spec  = if ($comma -gt 2) {
            [int]::TryParse($raw.Substring($comma + 1), [ref]$idx) | Out-Null
            $raw.Substring(0, $comma)
        } else { $raw }
        $spec = $spec.Trim().Trim('"').Trim("'").Trim()  # strip any surrounding quotes

        $pathOk = $false
        try { $pathOk = [System.IO.File]::Exists($spec) } catch { }
        if (-not $pathOk) {
            $script:imgCache[$cacheKey] = fallbackIndex $entry
            return fallbackIndex $entry
        }

        $icon = [IconUtil]::GetSmall($spec, $idx)
        if (-not $icon) { $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($spec) }

        if ($icon) {
            $bmp   = $icon.ToBitmap()
            $small = New-Object System.Drawing.Bitmap(16, 16)
            $g     = [System.Drawing.Graphics]::FromImage($small)
            $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $g.DrawImage($bmp, 0, 0, 16, 16)
            $g.Dispose(); $bmp.Dispose(); $icon.Dispose()

            $imgIdx = $il.Images.Count
            $il.Images.Add($small)
            $script:imgCache[$cacheKey] = $imgIdx
            return $imgIdx
        }
    } catch { }

    $script:imgCache[$cacheKey] = fallbackIndex $entry
    return fallbackIndex $entry
}

# ── Form icon (PNG-in-ICO via MemoryStream) ───────────────────────────────────
function pngToIcon([string]$path) {
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $ms = New-Object System.IO.MemoryStream
    $w  = New-Object System.IO.BinaryWriter($ms)
    $w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]1)
    $w.Write([byte]16);  $w.Write([byte]16);  $w.Write([byte]0)
    $w.Write([byte]0);   $w.Write([uint16]1); $w.Write([uint16]32)
    $w.Write([uint32]$bytes.Length); $w.Write([uint32]22)
    $w.Write($bytes)
    $ms.Position = 0
    return [System.Drawing.Icon]::new($ms)
}

# ── UI ────────────────────────────────────────────────────────────────────────
$script:entries  = [CmEntry[]]@()
$script:imageList = initImageList

$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Context Menu Manager'
$form.Size          = New-Object System.Drawing.Size(980, 580)
$form.MinimumSize   = New-Object System.Drawing.Size(700, 400)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

$appIconPath = Join-Path $PSScriptRoot '..\taskmon\icons\application_view_list.png'
if (Test-Path $appIconPath) {
    try { $form.Icon = pngToIcon $appIconPath } catch { }
}

# ── Toolbar ──
$toolbar = New-Object System.Windows.Forms.FlowLayoutPanel
$toolbar.Dock          = 'Top'
$toolbar.Height        = 34
$toolbar.Padding       = New-Object System.Windows.Forms.Padding(4, 4, 4, 0)
$toolbar.FlowDirection = 'LeftToRight'
$toolbar.WrapContents  = $false

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text    = 'Show:'
$lblFilter.AutoSize = $true
$lblFilter.Margin   = New-Object System.Windows.Forms.Padding(0, 4, 4, 0)

$cbFilter = New-Object System.Windows.Forms.ComboBox
$cbFilter.DropDownStyle = 'DropDownList'
$cbFilter.Width         = 160
$cbFilter.Margin        = New-Object System.Windows.Forms.Padding(0, 1, 8, 0)
$cbFilter.Items.AddRange(@('All', 'All Files', 'Folders', 'Folder Background', 'Video Files', 'Image Files', 'Drives'))
$cbFilter.SelectedIndex = 0

$chkDisabled = New-Object System.Windows.Forms.CheckBox
$chkDisabled.Text     = 'Show disabled'
$chkDisabled.Checked  = $true
$chkDisabled.AutoSize = $true
$chkDisabled.Margin   = New-Object System.Windows.Forms.Padding(0, 3, 12, 0)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text   = 'Refresh'
$btnRefresh.Width  = 70
$btnRefresh.Height = 24
$btnRefresh.Margin = New-Object System.Windows.Forms.Padding(0, 1, 0, 0)

$toolbar.Controls.AddRange(@($lblFilter, $cbFilter, $chkDisabled, $btnRefresh))

# ── ListView ──
$lv = New-Object System.Windows.Forms.ListView
$lv.Dock           = 'Fill'
$lv.View           = 'Details'
$lv.CheckBoxes     = $true
$lv.FullRowSelect  = $true
$lv.GridLines      = $true
$lv.SmallImageList = $script:imageList
[void]$lv.Columns.Add('Name',        220)
[void]$lv.Columns.Add('Applies To',  130)
[void]$lv.Columns.Add('Kind',         65)
[void]$lv.Columns.Add('Source',       55)
[void]$lv.Columns.Add('Status',       70)
[void]$lv.Columns.Add('Registry Key', 320)

# ── Bottom bar ──
$bottom = New-Object System.Windows.Forms.Panel
$bottom.Dock   = 'Bottom'
$bottom.Height = 38

$btnEnable = New-Object System.Windows.Forms.Button
$btnEnable.Text   = 'Enable Selected'
$btnEnable.Width  = 115
$btnEnable.Height = 26
$btnEnable.Left   = 6
$btnEnable.Top    = 6

$btnDisable = New-Object System.Windows.Forms.Button
$btnDisable.Text   = 'Disable Selected'
$btnDisable.Width  = 120
$btnDisable.Height = 26
$btnDisable.Left   = 128
$btnDisable.Top    = 6

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.AutoSize  = $true
$lblStatus.Left      = 262
$lblStatus.Top       = 12
$lblStatus.Text      = 'Loading...'
$lblStatus.ForeColor = [System.Drawing.Color]::Gray

$bottom.Controls.AddRange(@($btnEnable, $btnDisable, $lblStatus))

# Add controls: Fill control first so Top/Bottom docked controls overlay correctly
$form.Controls.Add($lv)
$form.Controls.Add($bottom)
$form.Controls.Add($toolbar)

# ── Populate list ──
function populateList {
    $lv.BeginUpdate()
    $lv.Items.Clear()

    $filter       = $cbFilter.SelectedItem
    $showDisabled = $chkDisabled.Checked
    $shown = 0; $disabled = 0

    foreach ($e in $script:entries) {
        if ($filter -ne 'All' -and $e.AppliesTo -ne $filter) { continue }
        if (-not $showDisabled -and -not $e.Enabled) { continue }

        $imgIdx = getIconIndex $e $script:imageList

        $item = New-Object System.Windows.Forms.ListViewItem($e.Label, $imgIdx)
        $item.Checked = $e.Enabled
        $item.Tag     = $e

        [void]$item.SubItems.Add($e.AppliesTo)
        [void]$item.SubItems.Add($e.Kind)
        [void]$item.SubItems.Add($e.Source)
        [void]$item.SubItems.Add($(if ($e.Enabled) { 'Enabled' } else { 'Disabled' }))
        [void]$item.SubItems.Add($e.ReadPath)

        if (-not $e.Enabled) { $item.ForeColor = [System.Drawing.Color]::Gray }

        [void]$lv.Items.Add($item)
        $shown++
        if (-not $e.Enabled) { $disabled++ }
    }

    $lv.EndUpdate()
    $lblStatus.Text = "$shown shown  |  $disabled disabled  |  $($script:entries.Count) total"
}

function reloadEntries {
    $lblStatus.Text  = 'Scanning registry...'
    $form.Cursor     = [System.Windows.Forms.Cursors]::WaitCursor
    $lv.Items.Clear()
    $form.Refresh()
    try   { $script:entries = getAllEntries }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    populateList
}

# ── Events ────────────────────────────────────────────────────────────────────
$lv.add_ItemCheck({
    param($s, $e)
    $item  = $lv.Items[$e.Index]
    $entry = [CmEntry]$item.Tag
    $on    = ($e.NewValue -eq 'Checked')
    applyEntry $entry $on
    $entry.Enabled         = $on
    $item.ForeColor        = if ($on) { [System.Drawing.SystemColors]::WindowText } else { [System.Drawing.Color]::Gray }
    $item.SubItems[4].Text = if ($on) { 'Enabled' } else { 'Disabled' }
    $script:pendingNotify  = $true
})

$lv.add_ItemChecked({
    if ($script:pendingNotify) {
        $script:pendingNotify = $false
        notifyShell
        $dis = ($lv.Items | Where-Object { -not $_.Checked }).Count
        $lblStatus.Text = "$($lv.Items.Count) shown  |  $dis disabled  |  $($script:entries.Count) total"
    }
})

$btnEnable.add_Click({
    $changed = $false
    foreach ($item in @($lv.SelectedItems)) {
        $entry = [CmEntry]$item.Tag
        if (-not $entry.Enabled) {
            applyEntry $entry $true
            $entry.Enabled = $true; $item.Checked = $true
            $item.ForeColor = [System.Drawing.SystemColors]::WindowText
            $item.SubItems[4].Text = 'Enabled'
            $changed = $true
        }
    }
    if ($changed) { notifyShell; populateList }
})

$btnDisable.add_Click({
    $changed = $false
    foreach ($item in @($lv.SelectedItems)) {
        $entry = [CmEntry]$item.Tag
        if ($entry.Enabled) {
            applyEntry $entry $false
            $entry.Enabled = $false; $item.Checked = $false
            $item.ForeColor = [System.Drawing.Color]::Gray
            $item.SubItems[4].Text = 'Disabled'
            $changed = $true
        }
    }
    if ($changed) { notifyShell; populateList }
})

$btnRefresh.add_Click({
    $script:imgCache.Clear()
    $script:imageList.Images.Clear()
    $script:imageList = initImageList
    $lv.SmallImageList = $script:imageList
    reloadEntries
})
$cbFilter.add_SelectedIndexChanged({ populateList })
$chkDisabled.add_CheckedChanged({ populateList })

$form.add_Shown({ reloadEntries })
[void]$form.ShowDialog()

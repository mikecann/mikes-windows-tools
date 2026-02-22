# test-compare.ps1
# Compares what ctxmenu.ps1 shows for a context against the ACTUAL
# context menu items Windows would show for a real file.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File test-compare.ps1 -File C:\path\to\video.mp4
#   powershell -ExecutionPolicy Bypass -File test-compare.ps1 -ContextType Folders
#
# For file contexts it invokes the shell IContextMenu COM interface to get
# the real item list, then diffs it against what ctxmenu reports.

param(
    [string]$TestFile  = '',           # path to an actual file/folder to test against
    [string]$ContextType = 'Video Files'  # fallback if no file given
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ── Pull in the same scan logic from ctxmenu.ps1 ─────────────────────────────
$scriptDir = Split-Path $MyInvocation.MyCommand.Path
$mainScript = Join-Path $scriptDir 'ctxmenu.ps1'

# We dot-source ctxmenu.ps1 but skip the UI parts by stubbing ShowDialog.
# Easier: just re-use the individual scan functions by loading the file up to the UI section.
$src = Get-Content $mainScript -Raw
# Stop before the UI setup (after all function definitions)
$cutAt = $src.IndexOf('# ── Form icon ──')
if ($cutAt -lt 0) { $cutAt = $src.IndexOf('$script:entries') }
$functionsOnly = $src.Substring(0, $cutAt)
Invoke-Expression $functionsOnly

# ── Get real context menu items via IContextMenu COM ─────────────────────────
Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public class ShellCtxMenu {
    [DllImport("shell32.dll", CharSet=CharSet.Auto)]
    static extern IntPtr ILCreateFromPath(string pszPath);
    [DllImport("shell32.dll")]
    static extern void ILFree(IntPtr pidl);
    [DllImport("shell32.dll")]
    static extern int SHGetDesktopFolder(out IShellFolder ppshf);

    [ComImport, Guid("000214E6-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IShellFolder {
        void ParseDisplayName(IntPtr hwnd, IntPtr pbc, [MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName, IntPtr pchEaten, out IntPtr ppidl, IntPtr pdwAttributes);
        void EnumObjects(IntPtr hwnd, uint grfFlags, out IntPtr ppenumIDList);
        void BindToObject(IntPtr pidl, IntPtr pbc, ref Guid riid, out IntPtr ppv);
        void BindToStorage(IntPtr pidl, IntPtr pbc, ref Guid riid, out IntPtr ppv);
        void CompareIDs(IntPtr lParam, IntPtr pidl1, IntPtr pidl2);
        void CreateViewObject(IntPtr hwndOwner, ref Guid riid, out IntPtr ppv);
        void GetAttributesOf(uint cidl, IntPtr apidl, ref uint rgfInOut);
        void GetUIObjectOf(IntPtr hwndOwner, uint cidl, [MarshalAs(UnmanagedType.LPArray)] IntPtr[] apidl, ref Guid riid, IntPtr rgfReserved, out IntPtr ppv);
        void GetDisplayNameOf(IntPtr pidl, uint uFlags, out IntPtr pName);
        void SetNameOf(IntPtr hwnd, IntPtr pidl, [MarshalAs(UnmanagedType.LPWStr)] string pszName, uint uFlags, out IntPtr ppidlOut);
    }

    [ComImport, Guid("000214e4-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IContextMenu {
        [PreserveSig] int QueryContextMenu(IntPtr hmenu, uint indexMenu, uint idCmdFirst, uint idCmdLast, uint uFlags);
        [PreserveSig] int InvokeCommand(IntPtr pici);
        [PreserveSig] int GetCommandString(UIntPtr idCmd, uint uType, IntPtr pReserved, StringBuilder pszName, uint cchMax);
    }

    [DllImport("user32.dll")] static extern IntPtr CreatePopupMenu();
    [DllImport("user32.dll")] static extern bool DestroyMenu(IntPtr hMenu);
    [DllImport("user32.dll")] static extern int GetMenuItemCount(IntPtr hMenu);
    [DllImport("user32.dll", CharSet=CharSet.Auto)] static extern bool GetMenuItemInfo(IntPtr hMenu, uint uItem, bool fByPosition, ref MENUITEMINFO lpmii);

    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto)]
    struct MENUITEMINFO {
        public uint   cbSize;
        public uint   fMask;
        public uint   fType;
        public uint   fState;
        public uint   wID;
        public IntPtr hSubMenu;
        public IntPtr hbmpChecked;
        public IntPtr hbmpUnchecked;
        public IntPtr dwItemData;
        [MarshalAs(UnmanagedType.LPTStr)] public string dwTypeData;
        public uint   cch;
        public IntPtr hbmpItem;
    }
    const uint MIIM_STRING = 0x40;
    const uint MIIM_FTYPE  = 0x100;
    const uint MFT_SEPARATOR = 0x800;

    public static List<string> GetItems(string path) {
        var result = new List<string>();
        try {
            IntPtr pidlFull = ILCreateFromPath(path);
            if (pidlFull == IntPtr.Zero) return result;

            IShellFolder desktop;
            SHGetDesktopFolder(out desktop);

            IntPtr pidlRel;
            desktop.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, path, IntPtr.Zero, out pidlRel, IntPtr.Zero);

            // Get parent folder
            string parentPath = System.IO.Path.GetDirectoryName(path);
            IntPtr pidlParent = ILCreateFromPath(parentPath);
            IShellFolder parentFolder;
            var sfGuid = typeof(IShellFolder).GUID;
            IntPtr ppv;
            Marshal.QueryInterface(Marshal.GetIUnknownForObject(desktop), ref sfGuid, out ppv);

            // Get child pidl relative to parent
            IntPtr pidlChild;
            IntPtr dummy = IntPtr.Zero;
            desktop.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, path, IntPtr.Zero, out pidlChild, IntPtr.Zero);

            // Get IContextMenu
            var ctxGuid = new Guid("000214e4-0000-0000-c000-000000000046");
            IntPtr pCtxMenu;
            var pidls = new IntPtr[] { pidlChild };
            desktop.GetUIObjectOf(IntPtr.Zero, 1, pidls, ref ctxGuid, IntPtr.Zero, out pCtxMenu);
            if (pCtxMenu == IntPtr.Zero) { ILFree(pidlFull); return result; }

            var ctxMenu = (IContextMenu)Marshal.GetObjectForIUnknown(pCtxMenu);
            IntPtr hMenu = CreatePopupMenu();
            ctxMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, 0x100); // CMF_EXPLORE

            int count = GetMenuItemCount(hMenu);
            for (int i = 0; i < count; i++) {
                var mii = new MENUITEMINFO();
                mii.cbSize = (uint)Marshal.SizeOf(mii);
                mii.fMask = MIIM_STRING | MIIM_FTYPE;
                mii.dwTypeData = new string(' ', 256);
                mii.cch = 256;
                if (GetMenuItemInfo(hMenu, (uint)i, true, ref mii)) {
                    if ((mii.fType & MFT_SEPARATOR) != 0) {
                        result.Add("---");
                    } else if (!string.IsNullOrWhiteSpace(mii.dwTypeData)) {
                        result.Add(mii.dwTypeData.Trim());
                    }
                }
            }
            DestroyMenu(hMenu);
            Marshal.ReleaseComObject(ctxMenu);
            ILFree(pidlFull);
        } catch (Exception ex) {
            result.Add("ERROR: " + ex.Message);
        }
        return result;
    }
}
'@ -ErrorAction SilentlyContinue

# ── Run the scan (re-use getAllEntries from ctxmenu.ps1) ─────────────────────
Write-Host "Scanning registry..." -ForegroundColor Cyan
$allEntries = getAllEntries

# ── Determine context type from file if provided ──────────────────────────────
if ($TestFile -and (Test-Path $TestFile)) {
    $item = Get-Item $TestFile
    if ($item.PSIsContainer) {
        $ContextType = 'Folders'
    } else {
        $ext = $item.Extension.ToLower()
        $ContextType = switch -wildcard ($ext) {
            '.mp4' {'Video Files'} '.mkv' {'Video Files'} '.avi' {'Video Files'}
            '.mov' {'Video Files'} '.wmv' {'Video Files'} '.webm' {'Video Files'}
            '.jpg' {'Image Files'} '.jpeg' {'Image Files'} '.png' {'Image Files'}
            '.gif' {'Image Files'} '.bmp' {'Image Files'}
            default { 'All Files' }
        }
    }
    Write-Host "File: $TestFile  => context: $ContextType" -ForegroundColor Cyan
}

# ── Get registered items for this context (same logic as buildDisplayItems) ──
$categories = if ($ContextType -in @('Video Files','Image Files','All Files')) {
    @($ContextType, 'All Files')
} else { @($ContextType) }

$contextItems = @($allEntries | Where-Object { $_.AppliesTo -in $categories })
$deduped = [System.Collections.Generic.Dictionary[string,CmEntry]]::new()
foreach ($e in ($contextItems | Where-Object { $_.Source -eq 'HKCU' })) { $deduped[$e.VerbName] = $e }
foreach ($e in ($contextItems | Where-Object { $_.Source -ne 'HKCU' })) {
    if (-not $deduped.ContainsKey($e.VerbName)) { $deduped[$e.VerbName] = $e }
}
$registeredItems = @($deduped.Values | Sort-Object Label)

# ── Get REAL context menu if a file path was given ────────────────────────────
$realItems = @()
if ($TestFile -and (Test-Path $TestFile) -and ([System.Management.Automation.PSTypeName]'ShellCtxMenu').Type) {
    Write-Host "Querying real context menu from shell..." -ForegroundColor Cyan
    $realItems = [ShellCtxMenu]::GetItems($TestFile)
    if ($realItems[0] -like 'ERROR:*') {
        Write-Host "Shell query failed: $($realItems[0])" -ForegroundColor Yellow
        $realItems = @()
    }
}

# ── Output report ─────────────────────────────────────────────────────────────
$outLines = [System.Text.StringBuilder]::new()
function L([string]$s) { [void]$outLines.AppendLine($s); Write-Host $s }

L "========================================================"
L "  Context Menu Manager - Comparison Test"
L "  Context: $ContextType"
if ($TestFile) { L "  File:    $TestFile" }
L "  Date:    $(Get-Date)"
L "========================================================"
L ""

L "REGISTERED ITEMS ($($registeredItems.Count) total - what ctxmenu.ps1 can manage):"
L "--------------------------------------------------------"
foreach ($e in $registeredItems) {
    $state  = if ($e.Enabled) { '[ON ]' } else { '[OFF]' }
    $source = "[$($e.Source)]"
    $kind   = "[$($e.Kind)]  "
    L ("  $state $source  $($e.Label.PadRight(45)) $($e.ReadPath)")
}

if ($realItems.Count -gt 0) {
    L ""
    L "ACTUAL CONTEXT MENU ($($realItems.Count) items - what Windows shows):"
    L "--------------------------------------------------------"
    $realLabels = @{}
    foreach ($item in $realItems) {
        $clean = ($item -replace '&','').Trim()
        if ($clean -and $clean -ne '---') {
            L "  $clean"
            $realLabels[$clean.ToLower()] = $true
        } else {
            L "  ----"
        }
    }

    L ""
    L "DIFF ANALYSIS:"
    L "--------------------------------------------------------"
    L "Items in ACTUAL menu but NOT in registered list (unmanageable):"
    foreach ($item in $realItems) {
        $clean = ($item -replace '&','').Trim()
        if (-not $clean -or $clean -eq '---') { continue }
        $matched = $registeredItems | Where-Object {
            $_.Label -replace '&','' -replace '\s+\[.*\]','' -like "*$($clean.Split(' ')[0])*"
        }
        if (-not $matched) {
            L "  MISSING: $clean"
        }
    }

    L ""
    L "Items REGISTERED but likely conditional (not always visible):"
    $regLabels = @{}
    foreach ($e in $registeredItems) { $regLabels[($e.Label -replace '&','' -replace '\s+\[.*\]','').ToLower()] = $e }
    foreach ($e in $registeredItems) {
        $lbl = ($e.Label -replace '&','' -replace '\s+\[.*\]','').ToLower()
        $found = $realItems | Where-Object { ($_ -replace '&','').ToLower() -like "*$($lbl.Split(' ')[0])*" }
        if (-not $found -and $e.Enabled) {
            L "  CONDITIONAL? $($e.Label)   [$($e.AppliesTo)]"
        }
    }
} else {
    L ""
    L "NOTE: No test file provided, so no real-menu comparison available."
    L "To get a comparison, run:"
    L "  powershell -ExecutionPolicy Bypass -File test-compare.ps1 -TestFile C:\path\to\video.mp4"
}

L ""
L "KNOWN LIMITATIONS:"
L "  - Clipchamp: UWP/AppX app - Windows injects this dynamically, not in registry"
L "  - Play (default verb): shown by Windows as first item, registered in ProgID as 'open'"
L "  - Cloud sync items (Dropbox, Google Drive): only shown inside synced folders"
L "  - Windows built-ins (Cast to Device, Share, Defender, Properties etc): not in registry"

# Write to file
$outFile = Join-Path $PSScriptRoot "test-compare-$($ContextType -replace ' ','_').txt"
$outLines.ToString() | Set-Content $outFile -Encoding UTF8
Write-Host ""
Write-Host "Report saved to: $outFile" -ForegroundColor Green

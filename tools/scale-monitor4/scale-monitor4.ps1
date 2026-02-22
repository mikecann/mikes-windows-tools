# Monitor 4 (HG584T05) DPI scale toggle
# DpiValue 4 = 200% (normal), DpiValue 7 = 300% (filming)

$regKey = "HKCU:\Control Panel\Desktop\PerMonitorSettings\RTK8405_0C_07E9_97^C9A428C8B2686559443005CCA2CE3E2E"

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class WinApi {
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessageTimeout(
        IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam,
        uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);

    [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", SetLastError = true)]
    public static extern IntPtr SendMessageTimeoutW(
        IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam,
        uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int ChangeDisplaySettingsEx(
        string lpszDeviceName, IntPtr lpDevMode, IntPtr hwnd,
        uint dwflags, IntPtr lParam);

    public static readonly IntPtr HWND_BROADCAST = new IntPtr(0xffff);
    public const uint WM_SETTINGCHANGE = 0x001A;
    public const uint SMTO_ABORTIFHUNG = 0x0002;
    public const uint CDS_RESET        = 0x40000000;
}
"@

function Apply-Scale($dpiValue) {
    Set-ItemProperty -Path $regKey -Name DpiValue -Value $dpiValue -Type DWord

    $r = [UIntPtr]::Zero
    [WinApi]::SendMessageTimeout(
        [WinApi]::HWND_BROADCAST, [WinApi]::WM_SETTINGCHANGE,
        [UIntPtr]::Zero, "Environment",
        [WinApi]::SMTO_ABORTIFHUNG, 3000, [ref]$r) | Out-Null
    [WinApi]::SendMessageTimeoutW(
        [WinApi]::HWND_BROADCAST, [WinApi]::WM_SETTINGCHANGE,
        [UIntPtr]([uint32]0x97), [IntPtr]::Zero,
        [WinApi]::SMTO_ABORTIFHUNG, 3000, [ref]$r) | Out-Null
    [WinApi]::ChangeDisplaySettingsEx("\\.\DISPLAY4", [IntPtr]::Zero,
        [IntPtr]::Zero, [WinApi]::CDS_RESET, [IntPtr]::Zero) | Out-Null
}

if (-not (Test-Path $regKey)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Registry key not found. Is Monitor 4 connected?",
        "Scale Monitor 4", "OK", "Error") | Out-Null
    exit 1
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$current = (Get-ItemProperty -Path $regKey -Name DpiValue).DpiValue

# --- Build the popup ---
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Monitor 4 Scale"
$form.Size            = New-Object System.Drawing.Size(260, 140)
$form.StartPosition   = "Manual"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.TopMost         = $true
$form.BackColor       = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.ForeColor       = [System.Drawing.Color]::White
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

# Position near bottom-right of screen (above taskbar)
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$form.Location = New-Object System.Drawing.Point(
    ($screen.Right - $form.Width - 12),
    ($screen.Bottom - $form.Height - 8)
)

$label = New-Object System.Windows.Forms.Label
$label.Text      = if ($current -eq 7) { "Currently: 300%  (filming)" } else { "Currently: 200%  (normal)" }
$label.Location  = New-Object System.Drawing.Point(16, 14)
$label.Size      = New-Object System.Drawing.Size(220, 20)
$label.ForeColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
$form.Controls.Add($label)

function Make-Button($text, $x, $active) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text      = $text
    $btn.Location  = New-Object System.Drawing.Point($x, 48)
    $btn.Size      = New-Object System.Drawing.Size(104, 44)
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 1
    if ($active) {
        $btn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $btn.ForeColor = [System.Drawing.Color]::White
        $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $btn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    } else {
        $btn.BackColor = [System.Drawing.Color]::FromArgb(55, 55, 55)
        $btn.ForeColor = [System.Drawing.Color]::White
        $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    }
    return $btn
}

$btn200 = Make-Button "200%`nNormal" 16  ($current -eq 4)
$btn300 = Make-Button "300%`nFilming" 130 ($current -eq 7)

$btn200.Add_Click({
    Apply-Scale 4
    $form.Close()
})

$btn300.Add_Click({
    Apply-Scale 7
    $form.Close()
})

$form.Controls.Add($btn200)
$form.Controls.Add($btn300)

# Close on Escape or clicking outside
$form.Add_Deactivate({ $form.Close() })
$form.Add_KeyDown({ if ($_.KeyCode -eq "Escape") { $form.Close() } })
$form.KeyPreview = $true

[System.Windows.Forms.Application]::Run($form)

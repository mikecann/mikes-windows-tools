Add-Type -TypeDefinition '
using System;
using System.Text;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    [DllImport("user32.dll")]
    public static extern IntPtr GetParent(IntPtr hWnd);
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
}'

$tray = [Win32]::FindWindowEx([IntPtr]::Zero, [IntPtr]::Zero, "Shell_TrayWnd", $null)
Write-Host "Tray: $tray"

$child = [Win32]::FindWindowEx($tray, [IntPtr]::Zero, $null, $null)
$i = 0
while ($child -ne [IntPtr]::Zero) {
    $rect = New-Object Win32+RECT
    [void][Win32]::GetWindowRect($child, [ref]$rect)
    $vis = [Win32]::IsWindowVisible($child)
    
    $cls = New-Object System.Text.StringBuilder 256
    [void][Win32]::GetClassName($child, $cls, 256)
    
    Write-Host "[$i] Child: $child | Vis: $vis | Rect: $($rect.Left), $($rect.Top), $($rect.Right), $($rect.Bottom) | Class: $($cls.ToString())"
    $child = [Win32]::FindWindowEx($tray, $child, $null, $null)
    $i++
}

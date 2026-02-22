using System;
using System.Runtime.InteropServices;
using System.Text;

public class Diag {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    public static void Main() {
        IntPtr tray = FindWindow("Shell_TrayWnd", null);
        Console.WriteLine("Tray: " + tray);
        
        int i = 0;
        EnumChildWindows(tray, (child, param) => {
            RECT rect;
            GetWindowRect(child, out rect);
            bool vis = IsWindowVisible(child);
            StringBuilder cls = new StringBuilder(256);
            GetClassName(child, cls, 256);
            
            Console.WriteLine(string.Format("[{0}] Child: {1} | Vis: {2} | Rect: {3}, {4}, {5}, {6} | Class: {7}", i, child, vis, rect.Left, rect.Top, rect.Right, rect.Bottom, cls));
            i++;
            return true;
        }, IntPtr.Zero);
    }
}

using System;
using System.Runtime.InteropServices;
using System.Text;

public class Diag2 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern IntPtr GetParent(IntPtr hWnd);

    public static void Main() {
        EnumWindows((child, param) => {
            StringBuilder cls = new StringBuilder(256);
            GetClassName(child, cls, 256);
            if (cls.ToString().Contains("WindowsForms")) {
                Console.WriteLine("Form: " + child + " Class: " + cls + " Parent: " + GetParent(child));
            }
            return true;
        }, IntPtr.Zero);
    }
}

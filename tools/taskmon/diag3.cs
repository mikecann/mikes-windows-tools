using System;
using System.Runtime.InteropServices;
using System.Text;

public class Diag3 {
    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr GetParent(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    public static void Main(string[] args) {
        IntPtr handle = new IntPtr(int.Parse(args[0]));
        Console.WriteLine("IsWindow: " + IsWindow(handle));
        Console.WriteLine("Parent: " + GetParent(handle));
        Console.WriteLine("Style: " + GetWindowLong(handle, -16).ToString("X"));
    }
}

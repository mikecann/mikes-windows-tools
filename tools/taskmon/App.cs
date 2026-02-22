using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;

namespace TaskMon {

// =============================================================================
// DarkRenderer -- makes the right-click context menu match the dark overlay theme
// =============================================================================
class DarkRenderer : ToolStripProfessionalRenderer {
    class DarkCT : ProfessionalColorTable {
        public override Color MenuItemSelected            { get { return Color.FromArgb(0x45, 0x45, 0x55); } }
        public override Color MenuItemBorder              { get { return Color.FromArgb(0x55, 0x55, 0x66); } }
        public override Color MenuBorder                  { get { return Color.FromArgb(0x44, 0x44, 0x44); } }
        public override Color ToolStripDropDownBackground { get { return Color.FromArgb(0x2D, 0x2D, 0x2D); } }
        public override Color ImageMarginGradientBegin    { get { return Color.FromArgb(0x2D, 0x2D, 0x2D); } }
        public override Color ImageMarginGradientMiddle   { get { return Color.FromArgb(0x2D, 0x2D, 0x2D); } }
        public override Color ImageMarginGradientEnd      { get { return Color.FromArgb(0x2D, 0x2D, 0x2D); } }
        public override Color SeparatorDark               { get { return Color.FromArgb(0x44, 0x44, 0x44); } }
        public override Color SeparatorLight              { get { return Color.FromArgb(0x44, 0x44, 0x44); } }
    }
    public DarkRenderer() : base(new DarkCT()) {}
}

// =============================================================================
// App -- entry point called by taskmon.ps1
// =============================================================================
public static class App {
    const string REG_KEY  = @"Software\Microsoft\Windows\CurrentVersion\Run";
    const string REG_NAME = "taskmon";

    // scriptDir is passed from taskmon.ps1 ($PSScriptRoot) so we can find taskmon.vbs.
    public static void Run(string scriptDir = null) {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        var s = Settings.Load();
        // Apply startup registration on every launch so it stays in sync with the setting.
        if (scriptDir != null)
            ApplyStartup(s.RunOnStartup, scriptDir);
        using (var f = new OverlayForm(s, scriptDir)) {
            f.Show();
            Application.Run(f);
        }
    }

    // Adds or removes the HKCU Run entry.
    // cmd = wscript.exe "<path to taskmon.vbs>"
    internal static void ApplyStartup(bool enable, string scriptDir) {
        try {
            using (var key = Registry.CurrentUser.OpenSubKey(REG_KEY, writable: true)) {
                if (key == null) return;
                if (enable) {
                    var vbs = Path.Combine(scriptDir, "taskmon.vbs");
                    key.SetValue(REG_NAME,
                        string.Format("wscript.exe \"{0}\"", vbs));
                } else {
                    key.DeleteValue(REG_NAME, throwOnMissingValue: false);
                }
            }
        } catch { }
    }
}

} // namespace TaskMon

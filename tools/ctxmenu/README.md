# ctxmenu

A GUI for managing Windows Explorer context menu entries. Shows shell verbs
and COM extension handlers from the registry, and lets you toggle them on or
off without needing admin rights.

## Usage

```
ctxmenu
```

Or run `ctxmenu.vbs` directly. A window opens immediately - no console.

## What it shows

| Category | What's scanned |
|---|---|
| All Files | `*\shell` and `*\shellex\ContextMenuHandlers` |
| Folders | `Directory\shell` and `Directory\shellex\ContextMenuHandlers` |
| Folder Background | `Directory\Background\shell` and `...\shellex\...` |
| Drives | `Drive\shell` |
| Video Files | `SystemFileAssociations\.<ext>\shell` (grouped across all video exts) |
| Image Files | `SystemFileAssociations\.<ext>\shell` (grouped across all image exts) |

Both HKCU (user) and HKLM (system/app-installed) entries are shown.

## Toggling entries

Click the checkbox next to an entry, or select rows and use **Enable Selected**
/ **Disable Selected**.

Changes take effect immediately - Explorer is notified via `SHChangeNotify`
so you don't need to restart it.

**How disabling works (no admin needed):**

- **Verb entries** (`Verb` / `Submenu` kind): adds an empty `LegacyDisable`
  value to a HKCU shadow key. Windows merges HKCU on top of HKLM when
  building HKCR, so this suppresses system-installed entries too.
- **COM handlers** (`COM` kind): creates a HKCU shadow key with the CLSID
  prefixed by `-` (e.g. `{ABC...}` becomes `-{ABC...}`), which causes COM
  activation to fail silently.

To re-enable: check the box again. The shadow key is cleaned up.

## What it won't show

- **ProgID-specific entries** (e.g. VLC's "Play with VLC" registered under
  `VLC.mp4\shell`) - these are tied to specific file-type ProgIDs and not
  enumerated here. Tools like [ShellMenuView](https://www.nirsoft.net/utils/shell_menu_view.html)
  by NirSoft cover those.
- Inline `Open with` suggestions (those come from app capabilities, not shell keys).

## Notes

- No external dependencies - uses built-in .NET WinForms.
- All writes go to HKCU - never modifies HKLM directly.
- Safe to run multiple times; toggling is fully reversible.

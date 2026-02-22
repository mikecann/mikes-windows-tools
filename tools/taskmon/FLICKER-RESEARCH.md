# taskmon — app-switch flicker research

Investigation into the 1–2 frame flicker visible when switching between
maximized applications while the taskmon overlay is running.

---

## The problem

When Alt+Tabbing or clicking between maximized windows, the taskmon overlay
briefly disappears for 1–2 frames and then reappears. It looks like a quick
flash of bare taskbar where the sparklines normally sit.

---

## Root cause

taskmon's overlay is a **separate top-level `WS_POPUP` window** using
`UpdateLayeredWindow` for per-pixel alpha rendering, positioned on top of the
taskbar with `HWND_TOPMOST`.

When you switch apps, the DWM compositor recomposes all visible surfaces. The
taskbar itself redraws (updating the active-app highlight), and for 1–2 frames
the compositor shows the taskbar's surface before compositing the overlay back
on top. The overlay's z-order may also briefly get demoted during the switch.

The fundamental issue: **two independent compositor surfaces can never be
perfectly synchronised frame-by-frame.** The taskbar repaints on one
schedule, the overlay repaints on another, and the DWM composes them with a
potential 1-frame lag between the two.

---

## Approaches tried

### 1. WM_WINDOWPOSCHANGING interception (partial fix)

Hook `WM_WINDOWPOSCHANGING` (0x0046) and rewrite the `WINDOWPOS` struct's
`hwndInsertAfter` field to `HWND_TOPMOST` before the position change takes
effect. This prevents z-order demotion before it happens, unlike the existing
`WM_WINDOWPOSCHANGED` handler which fixes it after.

**Result:** reduced the flicker slightly but did not eliminate it. The
remaining flicker comes from compositor recomposition, not z-order demotion.

### 2. DWM transition attributes (partial fix)

Set `DWMWA_TRANSITIONS_FORCEDISABLED` and `DWMWA_EXCLUDED_FROM_PEEK` on the
overlay window via `DwmSetWindowAttribute`. This tells the DWM to skip all
fade/slide animations and exclude the overlay from Aero Peek.

```csharp
int TRUE = 1;
Native.DwmSetWindowAttribute(Handle,
    Native.DWMWA_TRANSITIONS_FORCEDISABLED, ref TRUE, sizeof(int));
Native.DwmSetWindowAttribute(Handle,
    Native.DWMWA_EXCLUDED_FROM_PEEK, ref TRUE, sizeof(int));
```

**Result:** removed the longer animation-based flicker but the 1–2 frame
compositor flash remains.

### 3. SetParent into Shell_TrayWnd — WS_CHILD embedding (broken)

The theoretically correct fix: call `SetParent(overlayHwnd, shellTrayWnd)` to
make the overlay a child window of the taskbar. As a child, it shares the
same compositor surface — the taskbar can never "show through" underneath
because the overlay *is* part of the taskbar's visual tree.

Implementation:

```csharp
IntPtr tray = Native.FindWindow("Shell_TrayWnd", null);
int style = Native.GetWindowLong(Handle, GWL_STYLE);
style = (style & ~WS_POPUP) | WS_CHILD;
Native.SetWindowLong(Handle, GWL_STYLE, style);
Native.SetParent(Handle, tray);
```

**Result:** the overlay became invisible. `UpdateLayeredWindow` does not work
correctly on `WS_CHILD` windows. The `WS_EX_LAYERED` extended style is
technically supported on child windows since Windows 8, but
`UpdateLayeredWindow` with per-pixel alpha compositing (`AC_SRC_ALPHA`) fails
silently when the window is a child — the bits are never
composited onto the parent's surface.

This is confirmed by Microsoft docs: `UpdateLayeredWindow` is designed for
top-level windows. For child windows, you'd need to use `SetLayeredWindowAttributes`
(which only supports a global alpha or a transparency colour key, not per-pixel
alpha), or switch to regular `WM_PAINT` rendering.

### 4. What would actually work (not implemented)

**Full rendering rewrite:** replace `UpdateLayeredWindow` with standard
`WM_PAINT` rendering inside a `WS_CHILD` of `Shell_TrayWnd`. Use
`TransparencyKey` (a magic background colour) for the see-through portions.

This is what TrafficMonitor does — it's a pure Win32 `WS_CHILD` embedded in
the taskbar with `WM_PAINT` rendering.

**Trade-offs:**
- Would completely eliminate flicker (same compositor surface)
- Loses per-pixel alpha — transparency is binary (transparent or not), no
  smooth alpha blending for the sparkline fills
- The 18% alpha sparkline fill under text would need to become either fully
  opaque or fully transparent
- The `TransparencyKey` approach is fragile — if any pixel in the rendered
  content accidentally matches the key colour, it becomes a hole
- Significant code change: ~200 lines of `UpdateLayeredWindow` + DIB setup
  replaced with `OnPaint` + `TransparencyKey`
- Mouse hit-testing changes (transparent regions don't receive clicks)
- The global opacity slider (Settings → Behaviour) would stop working since
  `SourceConstantAlpha` is an `UpdateLayeredWindow` feature

---

## Current status

Living with the 1–2 frame flicker. The DWM attribute fixes
(`DWMWA_TRANSITIONS_FORCEDISABLED` + `DWMWA_EXCLUDED_FROM_PEEK`) and the
`WM_WINDOWPOSCHANGING` interception are **not** currently committed — they
helped slightly but added complexity for minimal benefit.

If the flicker becomes more bothersome in future, the path forward is
approach 4 (full rendering rewrite to `WM_PAINT` + `WS_CHILD`).

---

## References

- [UpdateLayeredWindow docs](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-updatelayeredwindow)
  — "A layered window need not be a top-level window" but AC_SRC_ALPHA compositing
  is unreliable on child windows in practice
- [DWMWA_TRANSITIONS_FORCEDISABLED](https://learn.microsoft.com/en-us/windows/win32/api/dwmapi/ne-dwmapi-dwmwindowattribute)
- [TrafficMonitor source](https://github.com/zhongyang219/TrafficMonitor) — uses
  WS_CHILD + WM_PAINT embedding into Shell_TrayWnd

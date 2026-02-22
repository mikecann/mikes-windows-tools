"""
voice-type.py — Push-to-talk voice typing tool.

Hold RIGHT_CTRL while speaking. Partial transcription appears in the overlay
as you talk. Release to paste the final text into the active window.

A microphone icon lives in the system tray; right-click for settings and exit.

Requirements: faster-whisper, sounddevice, numpy, Pillow, pystray
"""

import os
import sys
import time
import math
import json
import winreg
import threading
import queue
import ctypes
import ctypes.wintypes
import tkinter as tk

os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

_LOG_PATH    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "voice-type.log")
_SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
_LOG_MAX_MB  = 1       # rotate when log exceeds this size
_LOG_KEEP    = 200     # lines to keep after rotation

def _rotate_log():
    """On startup: if log > _LOG_MAX_MB, keep only the last _LOG_KEEP lines."""
    try:
        if not os.path.exists(_LOG_PATH):
            return
        if os.path.getsize(_LOG_PATH) < _LOG_MAX_MB * 1024 * 1024:
            return
        with open(_LOG_PATH, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
        kept = lines[-_LOG_KEEP:]
        with open(_LOG_PATH, "w", encoding="utf-8") as f:
            f.write(f"[log rotated — kept last {_LOG_KEEP} of {len(lines)} lines]\n")
            f.writelines(kept)
    except Exception:
        pass  # never crash on log housekeeping

_rotate_log()
_log_lock   = threading.Lock()
_log_file   = open(_LOG_PATH, "a", encoding="utf-8", buffering=1)


def log(msg: str):
    ts = time.strftime("%H:%M:%S")
    line = f"{ts}  {msg}"
    with _log_lock:
        _log_file.write(line + "\n")
        _log_file.flush()
    print(line, flush=True)


log(f"=== voice-type started === log: {_LOG_PATH}")

import numpy as np
import sounddevice as sd
from faster_whisper import WhisperModel

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

HOTKEY_VK        = 0xA3   # VK_RCONTROL — Right Ctrl only (0xA2 = Left Ctrl)
POLL_INTERVAL    = 0.01   # key-state poll rate (100 Hz)
STREAM_INTERVAL  = 0.5    # seconds between streaming preview passes
STREAM_MIN_AUDIO = 0.8    # don't start streaming until this many seconds recorded

# Final transcription model (accurate):
#   CPU → "small.en"        ~0.5–1.5s depending on clip length
#   GPU → "large-v3-turbo"  ~0.2s on CUDA
GPU_MODEL    = "large-v3-turbo"
CPU_MODEL    = "small.en"

# Streaming preview model (speed over accuracy — visual feedback only):
# tiny.en runs in ~0.1s on CPU so it never meaningfully blocks the final pass.
STREAM_MODEL = "tiny.en"

SAMPLE_RATE  = 16000
CHANNELS     = 1
DTYPE        = "float32"
DEVICE       = None       # None = system default mic
COMPUTE_TYPE = "float16"  # float16 on GPU; overridden to int8 on CPU

# Models available in the tray settings menu.
# Final model: accuracy matters most; stream model: speed matters most.
FINAL_MODEL_OPTIONS  = ["tiny.en", "base.en", "small.en", "medium.en",
                        "large-v2", "large-v3", "large-v3-turbo"]
STREAM_MODEL_OPTIONS = ["tiny.en", "base.en", "small.en"]

# ---------------------------------------------------------------------------
# Settings (persisted to settings.json beside the script)
# ---------------------------------------------------------------------------

_SETTINGS_PATH = os.path.join(_SCRIPT_DIR, "settings.json")
_settings: dict = {}


def _load_settings():
    """Load settings.json, filling missing keys with hardware-appropriate defaults."""
    global _settings
    if os.path.exists(_SETTINGS_PATH):
        try:
            with open(_SETTINGS_PATH, "r", encoding="utf-8") as f:
                _settings = json.load(f)
        except Exception as e:
            log(f"Settings load failed: {e}; using defaults.")
            _settings = {}
    # Defaults are resolved after CUDA detection so the right model is chosen.
    cuda = _cuda_available()
    _settings.setdefault("final_model",  GPU_MODEL if cuda else CPU_MODEL)
    _settings.setdefault("stream_model", GPU_MODEL if cuda else STREAM_MODEL)
    _save_settings()


def _save_settings():
    try:
        with open(_SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(_settings, f, indent=2)
    except Exception as e:
        log(f"Settings save failed: {e}")


# ---------------------------------------------------------------------------
# Win32 helpers
# ---------------------------------------------------------------------------

_user32 = ctypes.windll.user32

_KEYEVENTF_KEYUP    = 0x0002
_KEYEVENTF_UNICODE  = 0x0004
_VK_CONTROL         = 0x11
_VK_V               = 0x56

# SendInput structures for clipboard-free text injection
_INPUT_KEYBOARD = 1

class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk",         ctypes.c_ushort),
        ("wScan",       ctypes.c_ushort),
        ("dwFlags",     ctypes.c_uint32),
        ("time",        ctypes.c_uint32),
        ("dwExtraInfo", ctypes.c_uint64),
    ]

class _INPUT_UNION(ctypes.Union):
    _fields_ = [
        ("ki",   _KEYBDINPUT),
        ("_pad", ctypes.c_byte * 28),   # ensure union >= sizeof(MOUSEINPUT)
    ]

class _INPUT(ctypes.Structure):
    _fields_ = [
        ("type",  ctypes.c_uint32),
        ("union", _INPUT_UNION),
    ]

_REG_RUN  = r"Software\Microsoft\Windows\CurrentVersion\Run"
_REG_NAME = "VoiceType"
_VBS_PATH = os.path.join(_SCRIPT_DIR, "voice-type.vbs")


def _key_is_down(vk: int) -> bool:
    return bool(_user32.GetAsyncKeyState(vk) & 0x8000)


class _MONITORINFOEX(ctypes.Structure):
    _fields_ = [
        ("cbSize",    ctypes.c_uint32),
        ("rcMonitor", ctypes.wintypes.RECT),
        ("rcWork",    ctypes.wintypes.RECT),
        ("dwFlags",   ctypes.c_uint32),
    ]


def _foreground_monitor_work_area() -> tuple[int, int, int, int]:
    """Return (left, top, right, bottom) of the work area of the monitor
    that contains the current foreground window."""
    MONITOR_DEFAULTTONEAREST = 2
    hwnd = _user32.GetForegroundWindow()
    hmon = _user32.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
    info = _MONITORINFOEX()
    info.cbSize = ctypes.sizeof(_MONITORINFOEX)
    ctypes.windll.user32.GetMonitorInfoW(hmon, ctypes.byref(info))
    r = info.rcWork
    return r.left, r.top, r.right, r.bottom


def _send_text_input(text: str):
    """Inject text via SendInput (KEYEVENTF_UNICODE) — clipboard is never touched."""
    inputs = []
    for ch in text:
        code = ord(ch)
        if code > 0xFFFF:
            # Encode as surrogate pair for characters outside the BMP
            code -= 0x10000
            chars = [0xD800 | (code >> 10), 0xDC00 | (code & 0x3FF)]
        else:
            chars = [code]
        for scan in chars:
            for flags in (_KEYEVENTF_UNICODE, _KEYEVENTF_UNICODE | _KEYEVENTF_KEYUP):
                inp = _INPUT()
                inp.type                 = _INPUT_KEYBOARD
                inp.union.ki.wVk         = 0
                inp.union.ki.wScan       = scan
                inp.union.ki.dwFlags     = flags
                inp.union.ki.time        = 0
                inp.union.ki.dwExtraInfo = 0
                inputs.append(inp)
    if not inputs:
        return
    arr  = (_INPUT * len(inputs))(*inputs)
    sent = ctypes.windll.user32.SendInput(len(inputs), arr, ctypes.sizeof(_INPUT))
    log(f"SendInput: {sent}/{len(inputs) // 2} char events delivered")


def _startup_enabled() -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_RUN) as k:
            winreg.QueryValueEx(k, _REG_NAME)
            return True
    except OSError:
        return False


def _set_startup(enable: bool):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_RUN,
                            access=winreg.KEY_SET_VALUE) as k:
            if enable:
                winreg.SetValueEx(k, _REG_NAME, 0, winreg.REG_SZ,
                                  f'wscript.exe "{_VBS_PATH}"')
                log("Run on Startup enabled.")
            else:
                winreg.DeleteValue(k, _REG_NAME)
                log("Run on Startup disabled.")
    except Exception as e:
        log(f"Startup toggle failed: {e}")


# ---------------------------------------------------------------------------
# CUDA detection
# ---------------------------------------------------------------------------

def _cuda_available() -> bool:
    try:
        import ctranslate2
        if ctranslate2.get_cuda_device_count() == 0:
            return False
        ctypes.cdll.LoadLibrary("cublas64_12.dll")
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Models
#
# Two separate instances so streaming never contends with final transcription:
#   _stream_model  tiny.en   CPU int8  ~0.1s/pass  — live preview only
#   _model         small.en  CPU int8  ~0.5–1.5s   — accurate final result
#                  (large-v3-turbo on CUDA for both)
# ---------------------------------------------------------------------------

_model: WhisperModel | None = None
_model_lock = threading.Lock()


def get_model() -> WhisperModel:
    global _model
    if _model is None:
        with _model_lock:
            if _model is None:
                cuda   = _cuda_available()
                device = "cuda" if cuda else "cpu"
                ct     = COMPUTE_TYPE if cuda else "int8"
                name   = _settings.get("final_model", CPU_MODEL)
                log(f"Loading final model {name!r} on {device} ({ct})...")
                _model = WhisperModel(name, device=device, compute_type=ct)
                log("Final model ready.")
    return _model


_stream_model: WhisperModel | None = None
_stream_model_lock = threading.Lock()


def get_stream_model() -> WhisperModel | None:
    """Returns the streaming preview model, or None if not yet loaded."""
    return _stream_model


def _load_stream_model():
    """Load the stream model in the background. Waits for the final model first
    to avoid competing for CPU during initial warm-up."""
    global _stream_model
    get_model()   # ensure final model finishes first
    with _stream_model_lock:
        if _stream_model is None:
            cuda   = _cuda_available()
            name   = _settings.get("stream_model", STREAM_MODEL)
            device = "cuda" if cuda else "cpu"
            ct     = COMPUTE_TYPE if cuda else "int8"
            log(f"Loading stream model {name!r} on {device} ({ct})...")
            _stream_model = WhisperModel(name, device=device, compute_type=ct)
            log("Stream model ready.")


def _set_final_model(name: str):
    """Switch the final transcription model; reloads it in the background."""
    global _model
    if _settings.get("final_model") == name:
        return
    log(f"Final model switching to {name!r}...")
    _settings["final_model"] = name
    _save_settings()
    with _model_lock:
        _model = None
    threading.Thread(target=get_model, daemon=True).start()


def _set_stream_model(name: str):
    """Switch the streaming preview model; reloads it in the background."""
    global _stream_model
    if _settings.get("stream_model") == name:
        return
    log(f"Stream model switching to {name!r}...")
    _settings["stream_model"] = name
    _save_settings()
    with _stream_model_lock:
        _stream_model = None
    threading.Thread(target=_load_stream_model, daemon=True).start()


# ---------------------------------------------------------------------------
# Tray icon drawing (Pillow)
# ---------------------------------------------------------------------------

_TRAY_COLORS = {
    "idle":       (72,  72,  82),
    "recording":  (192, 57,  43),
    "processing": (211, 84,   0),
    "disabled":   (38,  38,  42),
}

_TRAY_LABELS = {
    "idle":       "Voice Type — Ready",
    "recording":  "Voice Type — Recording…",
    "processing": "Voice Type — Transcribing…",
    "disabled":   "Voice Type — Disabled",
}


def _make_tray_icon(state: str):
    from PIL import Image, ImageDraw

    fg   = _TRAY_COLORS.get(state, _TRAY_COLORS["idle"])
    size = 64
    img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d    = ImageDraw.Draw(img)
    cx   = size // 2

    # Coloured background circle
    d.ellipse([1, 1, size - 2, size - 2], fill=(*fg, 255))

    # Microphone body (white capsule)
    wh = (255, 255, 255, 230)
    bw, bh, radius = 16, 22, 8
    bx0, by0 = cx - bw // 2, 9
    bx1, by1 = cx + bw // 2, by0 + bh
    try:
        d.rounded_rectangle([bx0, by0, bx1, by1], radius=radius, fill=wh)
    except AttributeError:
        # Pillow < 8.2 fallback
        d.rectangle([bx0 + radius, by0, bx1 - radius, by1], fill=wh)
        d.rectangle([bx0, by0 + radius, bx1, by1 - radius], fill=wh)
        for ex, ey in [(bx0, by0), (bx1 - 2*radius, by0),
                       (bx0, by1 - 2*radius), (bx1 - 2*radius, by1 - 2*radius)]:
            d.ellipse([ex, ey, ex + 2*radius, ey + 2*radius], fill=wh)

    # Stand arc
    d.arc([cx - 15, by1 - 3, cx + 15, by1 + 13], start=0, end=180, fill=wh, width=3)
    # Stem
    d.line([cx, by1 + 10, cx, by1 + 15], fill=wh, width=3)
    # Base
    d.line([cx - 9, by1 + 15, cx + 9, by1 + 15], fill=wh, width=3)

    return img


# ---------------------------------------------------------------------------
# System tray icon (pystray — runs in its own background thread)
# ---------------------------------------------------------------------------

class TrayIcon:
    def __init__(self, overlay: "Overlay"):
        self._overlay = overlay
        self.enabled  = True         # read/written by hotkey thread & tray thread
        self._icon    = None

    def start(self):
        import pystray

        def _make_final_action(name):
            return lambda: _set_final_model(name)

        def _make_final_check(name):
            return lambda item: _settings.get("final_model") == name

        def _make_stream_action(name):
            return lambda: _set_stream_model(name)

        def _make_stream_check(name):
            return lambda item: _settings.get("stream_model") == name

        def _final_model_items():
            return [
                pystray.MenuItem(
                    m,
                    _make_final_action(m),
                    checked=_make_final_check(m),
                    radio=True,
                )
                for m in FINAL_MODEL_OPTIONS
            ]

        def _stream_model_items():
            return [
                pystray.MenuItem(
                    m,
                    _make_stream_action(m),
                    checked=_make_stream_check(m),
                    radio=True,
                )
                for m in STREAM_MODEL_OPTIONS
            ]

        menu = pystray.Menu(
            pystray.MenuItem("Voice Type", None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "Enabled",
                self._toggle_enabled,
                checked=lambda item: self.enabled,
            ),
            pystray.MenuItem("Open Log",        self._open_log),
            pystray.MenuItem(
                "Run on Startup",
                self._toggle_startup,
                checked=lambda item: _startup_enabled(),
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Final Model",  pystray.Menu(lambda: _final_model_items())),
            pystray.MenuItem("Preview Model", pystray.Menu(lambda: _stream_model_items())),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit", self._on_exit),
        )
        self._icon = pystray.Icon(
            "voice-type",
            _make_tray_icon("idle"),
            _TRAY_LABELS["idle"],
            menu,
        )
        self._icon.run_detached()
        log("Tray icon started.")

    def set_state(self, state: str):
        """Thread-safe: update icon colour and tooltip to reflect current state."""
        if self._icon is None:
            return
        effective = "disabled" if not self.enabled else state
        self._icon.icon  = _make_tray_icon(effective)
        self._icon.title = _TRAY_LABELS.get(effective, "Voice Type")

    # ---- Menu callbacks (called on pystray's thread) ----

    def _toggle_enabled(self, icon, item):
        self.enabled = not self.enabled
        log(f"Voice Type {'enabled' if self.enabled else 'disabled'} via tray.")
        self.set_state("idle")

    def _open_log(self, icon, item):
        os.startfile(_LOG_PATH)

    def _toggle_startup(self, icon, item):
        _set_startup(not _startup_enabled())

    def _on_exit(self, icon, item):
        log("Exit requested via tray.")
        icon.stop()
        self._overlay.quit()   # ask tkinter main loop to exit cleanly


# ---------------------------------------------------------------------------
# Overlay window — must be created and run on the MAIN thread (Windows/Tk rule)
# ---------------------------------------------------------------------------

# Colours
_OVL_BG      = "#1C1C1E"   # dark charcoal background
_COL_REC     = "#FF453A"   # iOS-style red
_COL_PROC    = "#FF9F0A"   # iOS-style amber
_COL_TEXT    = "#EBEBF5"   # near-white
_COL_PREVIEW = "#8E8E93"   # grey for partial text

# Waveform bar geometry
_N_BARS    = 7
_BAR_W     = 4
_BAR_GAP   = 3
_CANVAS_W  = _N_BARS * _BAR_W + (_N_BARS - 1) * _BAR_GAP  # 46 px
_CANVAS_H  = 28
_BAR_MAX_H = 20
_BAR_MIN_H = 3

_MAX_PREVIEW_CHARS = 58


def _wrap_preview(text: str) -> str:
    """Word-wrap text and show the last 2 lines so recent speech is always visible."""
    if not text:
        return ""
    words = text.split()
    # Build all wrapped lines (no early exit)
    lines, current = [], ""
    for w in words:
        if current and len(current) + 1 + len(w) > _MAX_PREVIEW_CHARS:
            lines.append(current)
            current = w
        else:
            current = (current + " " + w).lstrip()
    if current:
        lines.append(current)
    if not lines:
        return ""
    # Show only the last 2 lines so the display tracks what you're currently saying.
    # A leading "…" indicates earlier text is scrolled off.
    visible = lines[-2:]
    prefix = "…" if len(lines) > 2 else ""
    return prefix + "\n".join(visible)


class Overlay:
    GWL_EXSTYLE      = -20
    WS_EX_NOACTIVATE = 0x08000000
    WS_EX_TOOLWINDOW = 0x00000080

    def __init__(self, get_level):
        """
        get_level: callable() -> float  — returns current mic RMS (0.0–1.0).
        Used to drive the waveform animation while recording.
        """
        self._get_level = get_level
        self._state     = "hidden"   # "hidden" | "rec" | "processing"
        self._bar_h     = [float(_BAR_MIN_H)] * _N_BARS
        self._monitor   = None       # cached work-area tuple for reposition

        self._root = tk.Tk()
        self._root.withdraw()
        self._root.overrideredirect(True)
        self._root.attributes("-topmost", True)
        self._root.configure(bg=_OVL_BG)
        self._root.resizable(False, False)

        # ── Left accent bar (4 px wide, coloured by state) ──────────────
        self._accent = tk.Frame(self._root, width=4, bg=_COL_REC)
        self._accent.pack(side="left", fill="y")

        # ── Main content ────────────────────────────────────────────────
        body = tk.Frame(self._root, bg=_OVL_BG, padx=10, pady=8)
        body.pack(side="left", fill="both", expand=True)

        # Top row: dot · label · waveform canvas
        top = tk.Frame(body, bg=_OVL_BG)
        top.pack(fill="x")

        self._dot = tk.Label(top, text="●", fg=_COL_REC, bg=_OVL_BG,
                             font=("Segoe UI", 8))
        self._dot.pack(side="left")

        self._label = tk.Label(top, text=" REC", fg=_COL_TEXT, bg=_OVL_BG,
                               font=("Segoe UI", 10, "bold"))
        self._label.pack(side="left")

        self._canvas = tk.Canvas(top, width=_CANVAS_W + 4, height=_CANVAS_H,
                                 bg=_OVL_BG, highlightthickness=0)
        self._canvas.pack(side="left", padx=(12, 0))

        # Draw bar rectangles (initially at minimum height, bottom-anchored)
        self._bar_ids = []
        for i in range(_N_BARS):
            x0 = 2 + i * (_BAR_W + _BAR_GAP)
            x1 = x0 + _BAR_W
            y1 = _CANVAS_H - 2
            y0 = y1 - _BAR_MIN_H
            rid = self._canvas.create_rectangle(x0, y0, x1, y1,
                                                fill=_COL_REC, outline="")
            self._bar_ids.append(rid)

        # Preview text (only shown when streaming text is available)
        self._preview = tk.Label(body, text="", fg=_COL_PREVIEW, bg=_OVL_BG,
                                 font=("Segoe UI", 10), anchor="w",
                                 justify="left", wraplength=360,
                                 pady=2)

        # Win32 window style — no focus steal, hidden from Alt+Tab
        hwnd  = self._root.winfo_id()
        style = ctypes.windll.user32.GetWindowLongW(hwnd, self.GWL_EXSTYLE)
        ctypes.windll.user32.SetWindowLongW(
            hwnd, self.GWL_EXSTYLE,
            style | self.WS_EX_NOACTIVATE | self.WS_EX_TOOLWINDOW,
        )

        self._visible   = False
        self._cmd_queue: queue.Queue = queue.Queue()
        self._root.after(50,  self._poll)
        self._root.after(33,  self._animate)   # 30 fps animation loop

    # ── Thread-safe public commands ──────────────────────────────────────

    def show_rec(self, preview: str = ""):
        self._cmd_queue.put(("rec", preview))

    def show_processing(self, preview: str = ""):
        self._cmd_queue.put(("processing", preview))

    def hide(self):
        self._cmd_queue.put(("hide", ""))

    def quit(self):
        self._cmd_queue.put(("quit", ""))

    def mainloop(self):
        self._root.mainloop()

    # ── Internal (main thread only) ──────────────────────────────────────

    def _poll(self):
        try:
            while True:
                cmd, preview = self._cmd_queue.get_nowait()
                if cmd == "quit":
                    self._root.destroy()
                    sys.exit(0)
                elif cmd == "hide":
                    self._root.withdraw()
                    self._visible = False
                    self._state   = "hidden"
                else:
                    self._state = cmd
                    col   = _COL_REC  if cmd == "rec" else _COL_PROC
                    label = " REC"    if cmd == "rec" else " ..."
                    self._accent.configure(bg=col)
                    self._dot.configure(fg=col)
                    self._label.configure(text=label)
                    for rid in self._bar_ids:
                        self._canvas.itemconfigure(rid, fill=col)
                    if preview:
                        self._preview.configure(text=preview)
                        self._preview.pack(fill="x")
                    else:
                        self._preview.pack_forget()
                    if not self._visible:
                        self._position()
                        self._root.deiconify()
                        self._visible = True
                    else:
                        self._reposition()
        except queue.Empty:
            pass
        self._root.after(50, self._poll)

    def _animate(self):
        if self._visible and self._state != "hidden":
            t = time.perf_counter()
            if self._state == "rec":
                raw   = self._get_level()
                level = min(raw * 14.0, 1.0)   # typical mic RMS is 0.01–0.07
                for i in range(_N_BARS):
                    phase = i * 0.75
                    freq  = 4.5 + i * 0.4
                    wave  = (math.sin(t * freq + phase) + 1) / 2
                    # Quiet idle: gentle low ripple; loud: bars jump high
                    target = _BAR_MIN_H + (_BAR_MAX_H - _BAR_MIN_H) * (
                        level * 0.75 + wave * (0.25 + level * 0.15)
                    )
                    self._bar_h[i] = self._bar_h[i] * 0.5 + target * 0.5
            else:
                # Processing: smooth travelling sine sweep
                for i in range(_N_BARS):
                    wave   = (math.sin(t * 3.5 + i * 0.75) + 1) / 2
                    target = _BAR_MIN_H + (_BAR_MAX_H - _BAR_MIN_H) * wave * 0.55
                    self._bar_h[i] = self._bar_h[i] * 0.6 + target * 0.4

            y_base = _CANVAS_H - 2
            for i, (rid, h) in enumerate(zip(self._bar_ids, self._bar_h)):
                x0 = 2 + i * (_BAR_W + _BAR_GAP)
                x1 = x0 + _BAR_W
                self._canvas.coords(rid, x0, y_base - int(h), x1, y_base)

        self._root.after(33, self._animate)

    def _position(self):
        """Position at bottom-centre of the monitor holding the focused window."""
        try:
            self._monitor = _foreground_monitor_work_area()
        except Exception:
            self._monitor = (0, 0,
                             self._root.winfo_screenwidth(),
                             self._root.winfo_screenheight())
        self._do_geometry()

    def _reposition(self):
        """Re-centre after size changes (preview text appearing/disappearing)."""
        if self._monitor is None:
            self._position()
            return
        self._do_geometry()

    def _do_geometry(self):
        left, _, right, bottom = self._monitor
        self._root.update_idletasks()
        w = self._root.winfo_reqwidth()
        h = self._root.winfo_reqheight()
        x = left + (right - left) // 2 - w // 2
        y = bottom - h - 20
        self._root.geometry(f"+{x}+{y}")


# ---------------------------------------------------------------------------
# Audio recorder
# ---------------------------------------------------------------------------

class Recorder:
    """Keeps the microphone stream open permanently so there is no hardware
    activation delay when the key is pressed.  Audio is only captured into
    _frames while _recording is True."""

    def __init__(self):
        self._frames:    list[np.ndarray] = []
        self._lock       = threading.Lock()
        self._recording  = False
        info = sd.query_devices(DEVICE, "input")
        log(f"Mic: {info['name']!r}")
        # blocksize=256 → 16 ms per callback — low enough that the first
        # captured block is ≤16 ms after the key goes down.
        self._stream = sd.InputStream(
            samplerate=SAMPLE_RATE, channels=CHANNELS, dtype=DTYPE,
            device=DEVICE, callback=self._callback, blocksize=256,
        )
        self._stream.start()

    def start(self):
        with self._lock:
            self._frames    = []
            self._recording = True

    def peek(self) -> np.ndarray:
        """Non-destructive snapshot of all audio recorded so far."""
        with self._lock:
            if not self._frames:
                return np.array([], dtype=np.float32)
            return np.concatenate(self._frames, axis=0).flatten()

    def get_rms(self) -> float:
        """RMS of the last ~100 ms of audio — drives the waveform animation."""
        with self._lock:
            if not self._frames:
                return 0.0
            recent = np.concatenate(self._frames[-2:], axis=0).flatten()
            if len(recent) == 0:
                return 0.0
            return float(np.sqrt(np.mean(recent ** 2)))

    def stop(self) -> np.ndarray:
        with self._lock:
            self._recording = False
            if not self._frames:
                return np.array([], dtype=np.float32)
            audio = np.concatenate(self._frames, axis=0).flatten()
            dur   = len(audio) / SAMPLE_RATE
            rms   = float(np.sqrt(np.mean(audio ** 2)))
            peak  = float(np.max(np.abs(audio)))
            log(f"Stopped: {dur:.2f}s  rms={rms:.4f}  peak={peak:.4f}")
            return audio

    def _callback(self, indata, frames, time_info, status):
        if status:
            log(f"Audio status: {status}")
        if self._recording:
            self._frames.append(indata.copy())


# ---------------------------------------------------------------------------
# Transcription
# ---------------------------------------------------------------------------

def transcribe(audio: np.ndarray, verbose: bool = True) -> str:
    duration = len(audio) / SAMPLE_RATE
    if duration < 0.3:
        return ""
    model    = get_model()
    segments, info = model.transcribe(
        audio,
        language="en",
        vad_filter=False,
        beam_size=1,
        condition_on_previous_text=False,
    )
    parts = [seg.text.strip() for seg in segments]
    result = " ".join(parts).strip()
    if verbose:
        log(f"Transcribed {duration:.1f}s → {result!r}  "
            f"(lang={info.language} p={info.language_probability:.2f})")
    return result


# ---------------------------------------------------------------------------
# Streaming transcriber — runs while key is held
# ---------------------------------------------------------------------------

class StreamingTranscriber:
    def __init__(self, recorder: Recorder, overlay: Overlay):
        self._recorder = recorder
        self._overlay  = overlay
        self._active   = False
        self._last_text = ""

    def start(self):
        self._active    = True
        self._last_text = ""
        threading.Thread(target=self._loop, daemon=True).start()

    def stop(self):
        """Signal the streaming loop to stop. Final transcription is always done by the caller."""
        self._active = False

    @property
    def last_preview(self) -> str:
        return _wrap_preview(self._last_text)

    def _loop(self):
        time.sleep(STREAM_INTERVAL)
        while self._active:
            model = get_stream_model()
            if model is None:
                # Stream model still loading — skip this tick silently
                time.sleep(STREAM_INTERVAL)
                continue

            audio = self._recorder.peek()
            if len(audio) >= SAMPLE_RATE * STREAM_MIN_AUDIO:
                if not self._active:
                    break
                t0 = time.perf_counter()
                # Use the dedicated stream model — never contends with _model_lock
                segs, _ = model.transcribe(
                    audio, language="en", vad_filter=False,
                    beam_size=1, condition_on_previous_text=False,
                )
                text = " ".join(s.text.strip() for s in segs).strip()
                if not self._active:
                    break
                elapsed = time.perf_counter() - t0
                log(f"Stream pass: {len(audio)/SAMPLE_RATE:.1f}s → {elapsed:.2f}s → {text[:60]!r}")
                self._last_text = text
                self._overlay.show_rec(_wrap_preview(text))
            time.sleep(STREAM_INTERVAL)


# ---------------------------------------------------------------------------
# Text injection
# ---------------------------------------------------------------------------

def paste_text(text: str):
    if not text.strip():
        return
    hwnd = _user32.GetForegroundWindow()
    buf  = ctypes.create_unicode_buffer(256)
    _user32.GetWindowTextW(hwnd, buf, 256)
    log(f"Injecting into {buf.value!r}: {text!r}")
    time.sleep(0.05)
    _send_text_input(text)
    time.sleep(0.05)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run():
    _load_settings()
    log(f"Settings: final_model={_settings['final_model']!r}  stream_model={_settings['stream_model']!r}")

    recorder = Recorder()
    overlay  = Overlay(get_level=recorder.get_rms)
    streamer = StreamingTranscriber(recorder, overlay)
    tray     = TrayIcon(overlay)
    tray.start()

    def hotkey_worker():
        # Load main model first, then stream model (sequenced to avoid CPU contention)
        threading.Thread(target=get_model, daemon=True).start()
        threading.Thread(target=_load_stream_model, daemon=True).start()
        log("Ready. Hold Right Ctrl to record.")

        was_down = False
        while True:
            is_down = _key_is_down(HOTKEY_VK)

            if is_down and not was_down:
                if not tray.enabled:
                    pass  # silently ignore while disabled
                else:
                    log("--- Key DOWN ---")
                    tray.set_state("recording")
                    overlay.show_rec()
                    recorder.start()
                    streamer.start()

            elif not is_down and was_down:
                if tray.enabled:
                    log("--- Key UP ---")
                    audio = recorder.stop()
                    streamer.stop()   # signal stream loop; final pass always runs below

                    def _finish(audio=audio, preview=streamer.last_preview):
                        # Show "processing" with the last streaming preview so the
                        # user sees what was recognised so far while we finalise.
                        overlay.show_processing(preview)
                        tray.set_state("processing")
                        t0 = time.perf_counter()
                        try:
                            text = transcribe(audio)
                        except Exception as e:
                            log(f"Transcription error: {e}")
                            text = ""
                        elapsed = time.perf_counter() - t0
                        overlay.hide()
                        tray.set_state("idle")
                        if text:
                            log(f"Done ({elapsed:.2f}s): {text!r}")
                            paste_text(text)
                        else:
                            log(f"Nothing to paste ({elapsed:.2f}s).")

                    threading.Thread(target=_finish, daemon=True).start()

            was_down = is_down
            time.sleep(POLL_INTERVAL)

    threading.Thread(target=hotkey_worker, daemon=True).start()

    # Tkinter mainloop MUST run on the main thread on Windows
    overlay.mainloop()


if __name__ == "__main__":
    try:
        run()
    except KeyboardInterrupt:
        sys.exit(0)

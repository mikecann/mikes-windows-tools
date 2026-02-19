"""
voice-type.py — Push-to-talk voice typing tool.

Hold RIGHT_CTRL while speaking. Partial transcription appears in the overlay
as you talk. Release to paste the final text into the active window.

A microphone icon lives in the system tray; right-click for settings and exit.

Requirements: faster-whisper, sounddevice, pyperclip, numpy, Pillow, pystray
"""

import os
import sys
import time
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
import pyperclip
from faster_whisper import WhisperModel

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

HOTKEY_VK       = 0xA3    # VK_RCONTROL — Right Ctrl only (0xA2 = Left Ctrl)
POLL_INTERVAL   = 0.01    # key-state poll rate (100 Hz)
STREAM_INTERVAL = 0.6     # seconds between streaming transcriptions while recording
STREAM_MIN_AUDIO = 0.8    # don't start streaming until this many seconds recorded

# Model selection:
#   CPU → "small.en"        fast (~0.4s per pass), English only
#   GPU → "large-v3-turbo"  best quality (~0.2s on CUDA)
GPU_MODEL    = "large-v3-turbo"
CPU_MODEL    = "small.en"

SAMPLE_RATE  = 16000
CHANNELS     = 1
DTYPE        = "float32"
DEVICE       = None       # None = system default mic
COMPUTE_TYPE = "float16"  # float16 on GPU; overridden to int8 on CPU

# ---------------------------------------------------------------------------
# Win32 helpers
# ---------------------------------------------------------------------------

_user32 = ctypes.windll.user32

_KEYEVENTF_KEYUP = 0x0002
_VK_CONTROL      = 0x11
_VK_V            = 0x56

_REG_RUN  = r"Software\Microsoft\Windows\CurrentVersion\Run"
_REG_NAME = "VoiceType"
_VBS_PATH = os.path.join(_SCRIPT_DIR, "voice-type.vbs")


def _key_is_down(vk: int) -> bool:
    return bool(_user32.GetAsyncKeyState(vk) & 0x8000)


def _send_ctrl_v():
    _user32.keybd_event(_VK_CONTROL, 0, 0, 0)
    _user32.keybd_event(_VK_V,       0, 0, 0)
    time.sleep(0.05)
    _user32.keybd_event(_VK_V,       0, _KEYEVENTF_KEYUP, 0)
    _user32.keybd_event(_VK_CONTROL, 0, _KEYEVENTF_KEYUP, 0)


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
# Model
# ---------------------------------------------------------------------------

_model: WhisperModel | None = None
_model_lock = threading.Lock()


def get_model() -> WhisperModel:
    global _model
    if _model is None:
        with _model_lock:
            if _model is None:
                cuda = _cuda_available()
                device = "cuda" if cuda else "cpu"
                ct     = COMPUTE_TYPE if cuda else "int8"
                name   = GPU_MODEL if cuda else CPU_MODEL
                log(f"Loading model {name!r} on {device} ({ct})...")
                _model = WhisperModel(name, device=device, compute_type=ct)
                log("Model ready.")
    return _model


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

_REC_BG  = "#c0392b"   # red
_PROC_BG = "#d35400"   # orange
_MAX_PREVIEW_CHARS = 55


def _wrap_preview(text: str) -> str:
    """Trim and word-wrap to at most 2 lines of _MAX_PREVIEW_CHARS each."""
    if not text:
        return ""
    words = text.split()
    lines, current = [], ""
    for w in words:
        if current and len(current) + 1 + len(w) > _MAX_PREVIEW_CHARS:
            lines.append(current)
            current = w
            if len(lines) == 2:
                break
        else:
            current = (current + " " + w).lstrip()
    if current and len(lines) < 2:
        lines.append(current)
    result = "\n".join(lines)
    if len(" ".join(words)) > len(result.replace("\n", " ")):
        result = result.rstrip() + "…"
    return result


class Overlay:
    GWL_EXSTYLE      = -20
    WS_EX_NOACTIVATE = 0x08000000
    WS_EX_TOOLWINDOW = 0x00000080

    def __init__(self):
        self._root = tk.Tk()
        self._root.withdraw()
        self._root.overrideredirect(True)
        self._root.attributes("-topmost", True)
        self._root.configure(bg=_REC_BG)

        self._status = tk.Label(
            self._root, text="", fg="white", bg=_REC_BG,
            font=("Segoe UI", 11, "bold"), padx=8, pady=3, anchor="w",
        )
        self._status.pack(fill="x")

        self._preview = tk.Label(
            self._root, text="", fg="#ffe0e0", bg=_REC_BG,
            font=("Segoe UI", 10), padx=8, pady=2, anchor="w",
            justify="left", wraplength=420,
        )
        self._preview.pack(fill="x")

        hwnd  = self._root.winfo_id()
        style = ctypes.windll.user32.GetWindowLongW(hwnd, self.GWL_EXSTYLE)
        ctypes.windll.user32.SetWindowLongW(
            hwnd, self.GWL_EXSTYLE,
            style | self.WS_EX_NOACTIVATE | self.WS_EX_TOOLWINDOW,
        )

        self._visible   = False
        self._cmd_queue: queue.Queue = queue.Queue()
        self._root.after(50, self._poll)

    # Thread-safe commands -----------------------------------------------

    def show_rec(self, preview: str = ""):
        self._cmd_queue.put(("rec", preview))

    def show_processing(self, preview: str = ""):
        self._cmd_queue.put(("processing", preview))

    def hide(self):
        self._cmd_queue.put(("hide", ""))

    def quit(self):
        """Thread-safe: ask the main loop to exit."""
        self._cmd_queue.put(("quit", ""))

    def mainloop(self):
        self._root.mainloop()

    # Internal (main thread only) ----------------------------------------

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
                else:
                    bg     = _REC_BG if cmd == "rec" else _PROC_BG
                    status = "  \u25cf REC" if cmd == "rec" else "  \u25cc  ..."
                    self._root.configure(bg=bg)
                    self._status.configure(text=status, bg=bg)
                    self._preview.configure(text=preview, bg=bg)
                    if preview:
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

    def _position(self):
        sw = self._root.winfo_screenwidth()
        sh = self._root.winfo_screenheight()
        self._root.update_idletasks()
        w = self._root.winfo_reqwidth()
        h = self._root.winfo_reqheight()
        self._root.geometry(f"+{sw - w - 12}+{sh - h - 48}")

    def _reposition(self):
        self._root.update_idletasks()
        sw = self._root.winfo_screenwidth()
        sh = self._root.winfo_screenheight()
        w  = self._root.winfo_reqwidth()
        h  = self._root.winfo_reqheight()
        self._root.geometry(f"+{sw - w - 12}+{sh - h - 48}")


# ---------------------------------------------------------------------------
# Audio recorder
# ---------------------------------------------------------------------------

class Recorder:
    def __init__(self):
        self._frames: list[np.ndarray] = []
        self._lock   = threading.Lock()
        self._stream: sd.InputStream | None = None

    def start(self):
        with self._lock:
            self._frames = []
            info = sd.query_devices(DEVICE, "input")
            log(f"Recording on: {info['name']!r}")
            self._stream = sd.InputStream(
                samplerate=SAMPLE_RATE, channels=CHANNELS, dtype=DTYPE,
                device=DEVICE, callback=self._callback, blocksize=1024,
            )
            self._stream.start()

    def peek(self) -> np.ndarray:
        """Non-destructive snapshot of all audio recorded so far."""
        with self._lock:
            if not self._frames:
                return np.array([], dtype=np.float32)
            return np.concatenate(self._frames, axis=0).flatten()

    def stop(self) -> np.ndarray:
        with self._lock:
            if self._stream:
                self._stream.stop()
                self._stream.close()
                self._stream = None
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
            audio = self._recorder.peek()
            if len(audio) >= SAMPLE_RATE * STREAM_MIN_AUDIO:
                # Guard: skip if key was released while we were sleeping
                if not self._active:
                    break
                t0   = time.perf_counter()
                text = transcribe(audio, verbose=False)
                # Guard: don't touch the overlay if key was released during transcription
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
    log(f"Pasting into {buf.value!r}: {text!r}")
    pyperclip.copy(text)
    time.sleep(0.1)
    _send_ctrl_v()
    time.sleep(0.15)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run():
    overlay  = Overlay()
    recorder = Recorder()
    streamer = StreamingTranscriber(recorder, overlay)
    tray     = TrayIcon(overlay)
    tray.start()

    def hotkey_worker():
        threading.Thread(target=get_model, daemon=True).start()
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

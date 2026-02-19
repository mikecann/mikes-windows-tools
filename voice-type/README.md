# voice-type

Push-to-talk voice transcription that types directly into any focused window.
Hold **Right Ctrl** while speaking, release to paste. Runs entirely locally —
no cloud, no subscription.

---

## Quick start

```powershell
# First time: install dependencies
powershell -ExecutionPolicy Bypass -File .\voice-type\deps.ps1

# Add to taskbar (run install.ps1 from repo root if not already done)
powershell -ExecutionPolicy Bypass -File .\install.ps1

# Or launch manually for testing
wscript.exe "C:\dev\me\mikes-windows-tools\voice-type\voice-type.vbs"
```

Right-click `C:\dev\tools\Voice Type.lnk` → **Pin to taskbar** for one-click launch.

---

## Usage

| Action | What happens |
|---|---|
| Hold **Right Ctrl** | Recording starts — red `● REC` overlay appears bottom-right |
| Keep holding | Partial transcription builds up in the overlay as you speak |
| Release **Right Ctrl** | Transcription finalises and text is pasted into the active window |
| Right-click **tray icon** | Settings menu (see below) |

The tool types into whatever window had focus when you released the key —
text editors, browsers, chat apps, terminals, etc.

---

## System tray icon

A microphone icon sits in the system tray. Its colour reflects the current state:

| Colour | State |
|---|---|
| Dark grey | Idle — ready to record |
| Red | Recording |
| Orange | Transcribing |
| Very dark | Disabled |

**Right-click menu:**

| Item | Description |
|---|---|
| **Enabled** ✓ | Toggle the tool on/off without killing the process |
| **Open Log** | Opens `voice-type.log` in Notepad |
| **Run on Startup** | Add/remove from `HKCU\...\Run` (auto-start with Windows) |
| **Exit** | Quit cleanly |

---

## How it works

- **Hotkey polling** — `GetAsyncKeyState(VK_RCONTROL)` at 100 Hz. No global
  keyboard hook is installed, so `Ctrl+C`, `Ctrl+V`, etc. are never affected.
- **Audio capture** — `sounddevice` streams 16 kHz mono float32 from the
  default microphone into a NumPy buffer.
- **Streaming preview** — while the key is held, a background thread
  transcribes accumulated audio every 0.6 s and updates the overlay so you
  can see words appear as you speak.
- **Final transcription** — on key release a full transcription of all recorded
  audio runs, ensuring accuracy. Result is copied to clipboard and pasted via
  `Ctrl+V` (`keybd_event`).
- **Model** — [faster-whisper](https://github.com/SYSTRAN/faster-whisper)
  (`small.en` on CPU, `large-v3-turbo` on CUDA). English-only, greedy decode
  for maximum speed.
- **Overlay** — `tkinter` frameless window with `WS_EX_NOACTIVATE` so it
  never steals focus from the window you're typing into.

---

## Performance

| Hardware | Model | Typical transcription time |
|---|---|---|
| CPU (any) | `small.en` | ~0.4 s for 3 s of speech |
| NVIDIA GPU (CUDA) | `large-v3-turbo` | ~0.2 s for 3 s of speech |

CUDA is auto-detected at startup. If `cublas64_12.dll` is not loadable it
falls back to CPU automatically.

---

## Configuration

Edit the constants near the top of `voice-type.py`:

| Constant | Default | Description |
|---|---|---|
| `HOTKEY_VK` | `0xA3` (Right Ctrl) | Virtual key code for push-to-talk |
| `CPU_MODEL` | `"small.en"` | Whisper model used when no GPU |
| `GPU_MODEL` | `"large-v3-turbo"` | Whisper model used when CUDA available |
| `STREAM_INTERVAL` | `0.6` | Seconds between preview transcriptions |
| `DEVICE` | `None` | Mic device (`None` = system default) |

Common hotkey alternatives: `0xA5` = Right Alt, `0x14` = Caps Lock,
`0x91` = Scroll Lock.

---

## Dependencies

Installed automatically by `deps.ps1`:

```
faster-whisper   speech-to-text engine (CTranslate2 backend)
sounddevice      microphone capture
pyperclip        clipboard write
numpy            audio buffer maths
Pillow           tray icon drawing
pystray          system tray integration
```

The Whisper model (~250 MB for `small.en`) downloads automatically on first
use and is cached in `%USERPROFILE%\.cache\huggingface\`.

---

## Files

| File | Purpose |
|---|---|
| `voice-type.py` | Main script — all logic |
| `voice-type.vbs` | Silent launcher (no console window) |
| `voice-type.ps1` | PowerShell launcher called by the VBS |
| `deps.ps1` | Installs Python dependencies |
| `voice-type.log` | Runtime log (gitignored, auto-rotates at 1 MB) |

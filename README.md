# voice-to-text-system-tray

Minimal Windows background tool that transcribes speech to text on demand and types the result into the active window, controlled via a single keyboard shortcut.

## Screenshot

![System Tray](docs/screenshot.png)

*No screenshot available yet. Place a screenshot at `docs/screenshot.png`.*

## Features

- Runs silently in the Windows system tray with no persistent window
- Hold right Ctrl to record, release to transcribe and inject text into the focused window
- Single-instance enforcement via PID file — kills the previous instance automatically on restart
- Auto-installer: downloads and installs Python 3.12 silently if no compatible interpreter is found
- Structured file logging to `voice_to_text.log` with timestamps and log levels
- Optional Administrator elevation check with a user-facing warning dialog

## Stack

| Component | Library | Purpose |
|---|---|---|
| System tray | pystray | Tray icon and context menu |
| Speech recognition | SpeechRecognition | Google Speech API transcription |
| Audio capture | PyAudio | Microphone input stream |
| Global hotkey | keyboard / pynput | Right Ctrl press/release detection |
| Text injection | pyautogui / win32clipboard | Type transcription into active window |
| Error dialogs | tkinter (stdlib) | Pre-tray warning messages |
| Logging | logging (stdlib) | File + stdout structured logs |

## Setup / Installation

**Requirements:** Windows, Python 3.12 recommended (auto-installed if missing)

```bash
git clone https://github.com/gabrielcnb/voice-to-text-system-tray.git
cd voice-to-text-system-tray
pip install SpeechRecognition pyaudio pystray keyboard pynput pyautogui pywin32
```

**Run as Administrator** for reliable global hotkey capture:

Right-click `voice_to_text.py` and select "Run as administrator", or from an elevated prompt:

```bash
python voice_to_text.py
```

An internet connection is required. The tool uses Google Speech API with no offline fallback.

## Usage

1. Launch the script. A tray icon appears in the Windows notification area.
2. Focus any text input field (browser, editor, messaging app, etc.).
3. Hold **Right Ctrl** and speak.
4. Release **Right Ctrl**. The transcribed text is typed into the active window.

**Stop the application:**
Right-click the tray icon and select "Quit".

**Log file:**
```
voice_to_text.log   (same directory as the script, created on first run)
```

**PID file** (single-instance guard):
```
.voicetotext.pid    (same directory as the script)
```

## File Structure

```
voice-to-text-system-tray/
├── voice_to_text.py     # Main script: tray, audio, transcription, injection, auto-installer
├── voice_to_text.log    # Runtime log (auto-created)
├── .voicetotext.pid     # Single-instance PID file (auto-created)
└── README.md
```

## Known Limitations

- Windows only.
- Requires an internet connection for Google Speech API.
- Administrator privileges recommended for global hotkey capture across all applications.

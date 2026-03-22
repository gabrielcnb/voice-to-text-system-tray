# Voice to Text - System Tray

[![Python](https://img.shields.io/badge/Python-3.8--3.12-blue?logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows)](https://www.microsoft.com/windows)

Minimal Windows background tool that transcribes speech to text on demand and types the result into the active window, controlled via a configurable keyboard shortcut.

## Features

- Runs silently in the Windows system tray with no visible window
- Hold a hotkey to record, release to transcribe and inject text into the focused window
- Configurable hotkey (Right Ctrl, Left Ctrl, Shift, Alt, F8-F10, Scroll Lock, Pause)
- First-run configuration GUI for hotkey and auto-paste preferences
- Auto-punctuation: adds periods, question marks, or exclamation marks based on content (Portuguese-aware)
- Stutter correction: removes repeated consecutive words from transcription
- Recording overlay with real-time audio level visualization
- Audio feedback: distinct sounds for start, stop, success, and error events
- Clipboard-based paste or direct typing mode (configurable)
- Transcription history in the tray menu (last 10 results)
- Single-instance enforcement via PID file -- kills previous instance on restart
- Auto-installer: downloads and installs Python 3.12 silently if no compatible interpreter is found
- Auto-installs missing pip dependencies on first run with progress GUI
- Structured file logging to `voice_to_text.log`
- Administrator elevation check with user-facing warning dialog
- Persistent configuration via `config.json`

## Stack

| Component | Library | Purpose |
|---|---|---|
| System tray | pystray | Tray icon and context menu |
| Speech recognition | SpeechRecognition | Google Speech API transcription |
| Audio capture | PyAudio | Microphone input stream |
| Global hotkey | keyboard | Hotkey press/release detection |
| Text injection | pyautogui + pyperclip | Paste or type transcription into active window |
| Tray icon | Pillow | Dynamic icon rendering (green/red/yellow states) |
| Audio feedback | winsound (stdlib) | WAV playback for recording events |
| Overlay | tkinter (stdlib) | Recording indicator with audio level bars |
| Config | json (stdlib) | Persistent user preferences |
| Logging | logging (stdlib) | File + stdout structured logs |

## Setup

**Requirements:** Windows, Python 3.8-3.12 (auto-installed if missing), internet connection

```bash
git clone https://github.com/gabrielcnb/voice-to-text-system-tray.git
cd voice-to-text-system-tray
pip install -r requirements.txt
```

Or install manually:

```bash
pip install SpeechRecognition pyaudio pystray keyboard pyautogui pyperclip pillow
```

**Run as Administrator** for reliable global hotkey capture:

```bash
python voice_to_text.py
```

Or double-click `INICIAR VoiceToText.vbs` for a no-console launch.

On first run, a configuration window lets you pick the hotkey and paste behavior. Settings are saved to `config.json`.

## Usage

1. Launch the script. A green tray icon appears in the notification area.
2. Focus any text input field (browser, editor, chat, etc.).
3. Hold the configured hotkey (default: Right Ctrl) and speak.
4. Release the key. The transcribed text is pasted into the active window.

**Tray menu options:** view transcription history, reconfigure hotkey, quit.

## File Structure

```
voice-to-text-system-tray/
├── voice_to_text.py          # Main script: tray, audio, transcription, overlay, auto-installer
├── config.json               # User preferences (hotkey, paste mode)
├── INICIAR VoiceToText.vbs   # Windows launcher (no console window)
├── requirements.txt          # pip dependencies
├── LICENSE                   # MIT
└── README.md
```

## License

[MIT](LICENSE)

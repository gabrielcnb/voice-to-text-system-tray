# Voice to Text — System Tray

Windows background service that transcribes speech on demand. Hold the configured hotkey (default: Right Ctrl) to record, release to transcribe — text is automatically pasted at the cursor.

## Features

- Runs silently in the Windows system tray
- Configurable hotkey (Right Ctrl by default)
- Automatic paste after transcription
- Persistent log of all transcriptions
- Auto-start via `.vbs` launcher

## Stack

- Python 3.10+
- SpeechRecognition + PyAudio
- pynput (hotkey detection)
- pyperclip (clipboard)

## Setup

```bash
pip install SpeechRecognition pyaudio pynput pyperclip
python voice_to_text.py
```

To auto-start with Windows, run `INICIAR VoiceToText.vbs`.

## Configuration

Edit `config.json` (created on first run):

```json
{
  "tecla": "right ctrl",
  "tecla_display": "CTRL Direito",
  "colar_automatico": true
}
```

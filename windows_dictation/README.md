# Dictation Tool for Windows

Press **Ctrl+Shift+D** to start recording, press again to stop. Your speech is transcribed and pasted at the cursor.

## Quick Start

### 1. Install Python 3.10+
Download from [python.org](https://www.python.org/downloads/). Check "Add to PATH" during install.

### 2. Install dependencies
```
pip install -r requirements.txt
```

### 3. Run
```
python dictation_app.py
```

On first launch, a dialog will ask for your **OpenAI API key**. Get one from [platform.openai.com/api-keys](https://platform.openai.com/api-keys). The key is saved locally in `dictation_config.json` and remembered for future launches.

You can change the key later via the system tray icon (right-click > Change API Key).

## Usage

| Action | What happens |
|--------|-------------|
| **Ctrl+Shift+D** | Start recording (dot turns red) |
| **Ctrl+Shift+D** again | Stop recording, transcribe, paste at cursor |

The small floating bar shows the current state:
- **Gray dot** = idle, ready
- **Red dot** = recording
- **Yellow dot** = transcribing
- **Green dot** = done, text pasted

Drag the bar anywhere on screen. Right-click the tray icon to quit or change API key.

## Build Standalone .exe

Double-click `build_exe.bat` or run:
```
pyinstaller --onefile --windowed --name DictationTool dictation_app.py
```

The `.exe` will be in `dist/`. Share it with anyone -- they just double-click and enter their API key on first run.

## Troubleshooting

**"No speech detected"** - Check your microphone is set as default input in Windows Sound Settings.

**"Mic error"** - Another app may be using the microphone, or permissions are missing. Check Settings > Privacy > Microphone.

**Nothing pastes** - The tool uses Ctrl+V to paste. Make sure the target app accepts paste.

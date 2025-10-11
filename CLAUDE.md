# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is a voice-controlled dictation and AI assistant tool for macOS (with Windows support). The core application is **voice_shortcuts.py**, a PyQt5 GUI that provides three modes of operation:
1. **Writing Mode** (` + 1): AI-assisted text composition with clipboard context
2. **Speaking Mode** (` + 2): Verbal Q&A with spoken responses
3. **Dictation Mode** (` + 3): Pure speech-to-text transcription

## Architecture

### Core Components

**voice_shortcuts.py** (~1000 lines) - Single-file application with:
- **UI Layer** (lines 481-936): PyQt5 frameless, draggable window with FocusCircle widget
- **Audio Pipeline** (lines 809-841): sounddevice recording → numpy frames → WAV export
- **Transcription** (lines 357-472): ElevenLabs primary, Gemini fallback
- **LLM Processing** (lines 183-355): Gemini→Claude→OpenAI fallback chain with 5-second timeouts
- **Hotkey System** (lines 760-778): pynput global keyboard listener for ` or > + 1/2/3 combos

### Key Design Patterns

1. **Fallback Chain Architecture**: Every AI operation tries multiple providers sequentially
   - Transcription: ElevenLabs → Gemini
   - LLM responses: Gemini (flash) → Claude (3.7-sonnet) → OpenAI (gpt-4o)
   - Timeout: 5 seconds per API call (configurable at line 331)

2. **Mode-Specific Prompts**: System prompts defined at lines 93-139
   - `PROMPT_TEXT`: Direct writing assistant (no meta-commentary)
   - `PROMPT_SPOKEN`: Verbal Q&A with detail levels (short/medium/detailed)
   - Both support "ignore" prefix to bypass clipboard context

3. **Clipboard Integration**: Multimodal clipboard grabbing (lines 142-159)
   - Handles both text and base64-encoded images
   - Automatically included in LLM context unless "ignore" keyword used

## Development Commands

### Running the Application
```bash
# Standard run
python voice_shortcuts.py

# Or via shell wrapper
./dictate_tool.sh
```

### Building Executables
```bash
# macOS
./build_executable.sh

# Windows
build_executable.bat
```

### Environment Setup
```bash
# Create conda environment
conda create -n dictate_env python=3.12
conda activate dictate_env

# Install dependencies
pip install -r requirements.txt
```

## Configuration

### Required API Keys (.env file)
- **ANTHROPIC_API_KEY**: Claude LLM (fallback)
- **ELEVENLABS_API_KEY**: Primary transcription service

### Optional API Keys
- **GEMINI_API_KEY**: Fallback transcription and primary LLM
- **OPENAI_API_KEY**: Tertiary LLM fallback

### Voice Settings (macOS only)
- **VOICE**: System voice name (default: "Samantha")
- **VOICE_RATE**: Speech rate in WPM (default: 220)

## Critical Implementation Details

### Audio Processing
- Sample rate: 16kHz mono (line 41)
- Format: float32 → int16 conversion at line 840
- Frames buffered via callback at lines 829-831

### Hotkey Triggers
- Two trigger keys supported: `` ` `` (backtick) and `<` (less-than)
- Combos defined at lines 78-91
- Global keyboard listener persists across focus changes

### UI States
- **Mode Colors**: Blue (#00CCFF) = Write, Yellow (#FFCA28) = Speak, Red (#FF5252) = Dictate
- **FocusCircle**: Hollow when idle, filled when recording (lines 481-557)
- **Minimized UI**: Click circle to toggle between 120px and 60px width (lines 718-741)
- **Draggable**: Window can be repositioned via mouse drag (lines 743-758)

### Transcription Workflow
1. Audio recorded to temp WAV file (lines 836-841)
2. `transcribe_audio_with_elevenlabs()` attempts first (lines 357-403)
3. If ElevenLabs fails, `transcribe_audio_with_gemini()` called (lines 405-472)
4. Gemini uploads audio, polls for processing, generates transcription
5. Temp file deleted regardless of outcome (lines 858-861)

### LLM Response Generation
- `get_ai_response()` orchestrates fallback (lines 329-355)
- Each provider wrapped in `call_with_timeout()` for cancellation (lines 293-326)
- Response inserted via clipboard paste after deleting trigger keys (lines 866-918)

## Modifying the Application

### Adding New Hotkeys
1. Define new `keyboard.KeyCode` at line 78
2. Add combo to `COMBO_*` lists (lines 89-91)
3. Extend `on_press` handler (lines 763-771)
4. Implement mode in `_toggle()` method (lines 780-807)

### Changing AI Models
Model constants defined at lines 45-50:
- `CLAUDE_MODEL = "claude-3-7-sonnet-20250219"`
- `GEMINI_MODEL = "gemini-1.5-flash"` (LLM responses)
- `GEMINI_TRANSCRIPTION_MODEL = "gemini-2.0-flash-exp"`
- `OPENAI_MODEL = "gpt-4o"`
- `ELEVENLABS_TRANSCRIPTION_MODEL = "scribe_v1"`

### Adjusting Prompts
Edit system prompts at lines 93-139:
- `PROMPT_TEXT`: Writing assistant behavior
- `PROMPT_SPOKEN`: Q&A response style

**Important**: Prompts handle "ignore" keyword detection (lines 163-180)

### Timeout Configuration
- `API_TIMEOUT = 5` seconds per LLM call (line 331)
- `TRANSCRIPTION_TIMEOUT = 15` seconds (line 54, currently unused)

## Testing Notes

- No automated tests exist
- Manual testing requires:
  1. Valid API keys in `.env`
  2. Working microphone
  3. macOS accessibility permissions for global hotkeys
  4. Test all three modes separately
  5. Verify fallback behavior by temporarily invalidating API keys

## Platform-Specific Behavior

### macOS
- Uses `say` command for TTS (line 476)
- Requires Accessibility permissions for global hotkeys
- `.app` bundle created by PyInstaller

### Windows
- TTS not implemented (speak function will fail)
- Hotkey system should work via pynput
- `.exe` created by PyInstaller

## Build System

- **voice_shortcuts.spec**: PyInstaller configuration (not included in analysis)
- **build_executable.sh**: macOS build script
- **build_executable.bat**: Windows build script
- Distribution folder: `dist/VoiceShortcuts_Distribution/`

## Common Modification Scenarios

### Change Hotkey Trigger
Replace backtick/less-than at line 78:
```python
TRIGGER_KEYS = [keyboard.KeyCode.from_char('YOUR_KEY_HERE')]
```

### Add Fourth Mode
1. Add combo definition after line 91
2. Create prompt constant after line 139
3. Add color constant after line 581
4. Extend `_toggle()` switch at lines 791-802
5. Update `_toggle_settings()` menu at lines 652-710

### Disable Image Support
Comment out image handling in `grab_clipboard()` at lines 145-154

### Change Fallback Order
Reorder function calls in `get_ai_response()` at lines 334-352

## Related Files

- **activity_manager.py**: Legacy activity tracking (not integrated)
- **simple_dashboard.py**: Flask dashboard for activity logs
- **focus_timer.py**: Separate focus tracking app
- **manual_log_ui.py**: Manual activity logging interface

These are separate tools; only **voice_shortcuts.py** is the core dictation app.

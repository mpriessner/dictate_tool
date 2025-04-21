#!/usr/bin/env python3
# Load environment variables first
from dotenv import load_dotenv
load_dotenv()

"""
dictate_clipboard.py
A voice-plus-clipboard assistant that pastes Claude 3.7's direct answer.

🔑  ENV VARS   OPENAI_API_KEY   ANTHROPIC_API_KEY
🏷️  Hotkey     back-tick (`) + digit 1   (toggles record)
"""

import os
import sys
import time
import base64
import tempfile
import wave
import subprocess
import numpy as np
import sounddevice as sd
import pyperclip
from PyQt5.QtCore import Qt, QTimer, QBuffer, QByteArray
from PyQt5.QtGui import QImage
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QVBoxLayout, QWidget
from pynput import keyboard
import openai
from anthropic import Anthropic

# ──────────────────────────────────────────────────────
# Helper: grab clipboard (text or image) ──────────────
def grab_clipboard():
    cb = QApplication.clipboard()
    if cb.mimeData().hasImage():
        qimg: QImage = cb.image()
        ba = QByteArray()
        buf = QBuffer(ba)
        buf.open(QBuffer.WriteOnly)
        qimg.save(buf, "PNG")
        encoded = base64.b64encode(ba.data()).decode()
        return {
            "type": "image",
            "data": encoded,
            "mimetype": "image/png"
        }
    else:
        return {"type": "text", "data": cb.text()}

# ──────────────────────────────────────────────────────
# Helper: build Anthropic messages  ────────────────────
SYSTEM_PROMPT = """
You are a direct and concise writing assistant for Martin. Your task is to:
1. Write responses in the same language as the input/request
2. Write directly from Martin's perspective without any commentary
3. Focus on drafting emails, messages, and other written communication
4. Be extremely concise - use as few words as possible while maintaining clarity
5. Respond directly without preamble, acknowledgments, or meta-commentary
6. Match the tone and style appropriate for the context
7. Preserve any formatting from the clipboard content when relevant
8. Always end messages with an appropriate signature:
   - For German informal: "Liebe Grüße,\nMartin"
   - For English formal: "Best wishes,\nMartin"
"""

def build_messages(clip, transcript):
    if clip["type"] == "text" and clip["data"].strip():
        content = [{
            "type": "text",
            "text": f"Context from clipboard:\n{clip['data']}\n\nVoice request: {transcript.strip()}\n\nPlease answer the request directly—no preamble."
        }]
    elif clip["type"] == "image":
        content = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": clip["mimetype"],
                    "data": clip["data"]
                }
            },
            {
                "type": "text",
                "text": f"Voice request: {transcript.strip()}\n\nPlease answer the request directly—no preamble."
            }
        ]
    return [{"role": "user", "content": content}]

def ask_claude(msgs):
    client = Anthropic()  # picks up ANTHROPIC_API_KEY from env
    resp = client.messages.create(
        model="claude-3-7-sonnet-20250219",
        max_tokens=1024,
        temperature=0.3,
        system=SYSTEM_PROMPT,
        messages=msgs  # Pass the msgs directly
    )
    return resp.content[0].text

# ──────────────────────────────────────────────────────
class DictateClip(QMainWindow):
    SAMPLE_RATE = 16000
    HOTKEY = {keyboard.KeyCode.from_char('`'), keyboard.KeyCode.from_char('1')}  # ` + 1 combo

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dictate + Clipboard → Claude")
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        self.is_recording = False
        self.frames = []
        self._setup_ui()
        self._setup_hotkey()

    # UI -------------
    def _setup_ui(self):
        cw = QWidget()
        self.setCentralWidget(cw)
        lay = QVBoxLayout(cw)
        self.label = QLabel("Ready (press `+1)")
        btn = QPushButton("Hold to record")
        btn.setStyleSheet(
            "QPushButton{background:#4CAF50;color:#fff;padding:8px;border-radius:5px}"
            "QPushButton:pressed{background:#FF5733}"
        )
        lay.addWidget(self.label)
        lay.addWidget(btn)
        btn.pressed.connect(self.start_recording)
        btn.released.connect(self.stop_recording)

    # Global hotkey ---
    def _setup_hotkey(self):
        current = set()

        def on_press(key):
            if key in self.HOTKEY:
                current.add(key)
            if current == self.HOTKEY:
                self.toggle_recording()

        def on_release(key):
            if key in current:
                current.remove(key)

        self.listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        self.listener.start()

    # Audio callback --
    def _audio_cb(self, indata, frames, time_, status):
        if self.is_recording:
            self.frames.append(indata.copy())

    # Start / stop ----
    def start_recording(self):
        if self.is_recording:
            return
        self.frames.clear()
        self.is_recording = True
        self.label.setText("Recording…")
        self.stream = sd.InputStream(
            samplerate=self.SAMPLE_RATE,
            channels=1,
            dtype='float32',
            callback=self._audio_cb,
            blocksize=1024
        )
        self.stream.start()

    def stop_recording(self):
        if not self.is_recording:
            return
        self.is_recording = False
        self.stream.stop()
        self.stream.close()
        self.label.setText("Processing audio…")
        QTimer.singleShot(0, self._process)

    def toggle_recording(self):
        if self.is_recording:
            self.stop_recording()
        else:
            self.start_recording()

    # Core pipeline ----
    def _process(self):
        try:
            # 1. clipboard
            clip = grab_clipboard()
            self.label.setText("Captured clipboard content...")

            # 2. audio → wav tmp
            audio = np.concatenate(self.frames, axis=0)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as fp:
                with wave.open(fp, 'wb') as wf:
                    wf.setnchannels(1)
                    wf.setsampwidth(2)
                    wf.setframerate(self.SAMPLE_RATE)
                    wf.writeframes((audio * 32767).astype(np.int16).tobytes())
                wav_path = fp.name

            # 3. Whisper
            try:
                self.label.setText("Transcribing audio...")
                with open(wav_path, "rb") as af:
                    transcript = openai.Audio.transcribe("whisper-1", af)["text"]
                print(f"Transcription: {transcript}")
            finally:
                os.unlink(wav_path)

            # 4. Claude
            self.label.setText("Asking Claude...")
            msgs = build_messages(clip, transcript)
            try:
                answer = ask_claude(msgs)
                print(f"Claude's answer: {answer}")
            except Exception as e:
                self.label.setText(f"Claude error: {e}")
                return

            # 5. remove recording character and paste
            self.label.setText("Pasting response...")
            pyperclip.copy(answer)
            time.sleep(0.1)  # Brief pause before actions
            kb = keyboard.Controller()
            
            # Delete the backtick character that triggered recording
            kb.press(keyboard.Key.backspace)
            kb.release(keyboard.Key.backspace)
            kb.press(keyboard.Key.backspace)
            kb.release(keyboard.Key.backspace)
            kb.press(keyboard.Key.backspace)
            kb.release(keyboard.Key.backspace)
            kb.press(keyboard.Key.backspace)
            kb.release(keyboard.Key.backspace)
            time.sleep(0.1)  # Brief pause after delete
            
            # Paste the response
            with kb.pressed(keyboard.Key.cmd):  # Cmd+V (mac)
                kb.press('v')
                kb.release('v')

            self.label.setText("Done! (press `+1 to ask again)")

        except Exception as e:
            error_msg = f"Error in process: {str(e)}"
            print(error_msg)
            self.label.setText(error_msg)

    # Cleanup ----------
    def closeEvent(self, ev):
        if hasattr(self, 'listener') and self.listener.running:
            self.listener.stop()
        ev.accept()

# ──────────────────────────────────────────────────────
def main():
    openai_key = os.getenv("OPENAI_API_KEY")
    anthropic_key = os.getenv("ANTHROPIC_API_KEY")
    print("Environment check:")
    print(f"OpenAI API Key present: {bool(openai_key)}")
    print(f"Anthropic API Key present: {bool(anthropic_key)}")
    
    if not (openai_key and anthropic_key):
        print("Set OPENAI_API_KEY and ANTHROPIC_API_KEY env vars first")
        return

    app = QApplication(sys.argv)
    win = DictateClip()
    win.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

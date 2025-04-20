#!/usr/bin/env python3
# dictate_clip_voice.py
# 🔑 ENV: OPENAI_API_KEY, ANTHROPIC_API_KEY
# 🏷️ Hot‑key: back‑tick (`) + digit 2  – hold to record

from dotenv import load_dotenv; load_dotenv()

import os, sys, time, base64, tempfile, wave, subprocess
import numpy as np, sounddevice as sd, pyperclip
from PyQt5.QtCore import Qt, QTimer, QBuffer, QByteArray
from PyQt5.QtGui  import QImage
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QPushButton, 
                         QVBoxLayout, QHBoxLayout, QWidget, QSlider)
from pynput import keyboard
import openai                                     # Whisper API
from anthropic import Anthropic                   # Claude 3.x API

# ──────────────────────────────────────────────────────
# CONFIG
MODEL_NAME   = "claude-3-7-sonnet-20250219"        # vision-capable model
TEMP         = 0.3
SPEECH_RATE  = 220                                # words‑per‑minute for `say -r`
VOICE        = "Samantha"                         # macOS voice

SYSTEM_PROMPT = """
You are Martin's AI assistant. Your task is to provide quick, clear answers to verbal questions. Follow these rules:
1. First line: Briefly state what you understood the question/request to be
2. Second line onwards: Provide your direct answer
3. Keep responses extremely concise - use minimal words
4. Match the language of the question (English/German/etc.)
5. For questions about clipboard content, reference it naturally in your answer
6. No meta-commentary or explanations about your process
"""

# ──────────────────────────────────────────────────────
# Clipboard helper (text or screenshot)
def grab_clipboard():
    try:
        cb = QApplication.clipboard()
        if cb.mimeData().hasImage():
            qi: QImage = cb.image()
            if qi.isNull():
                return {"type": "text", "data": ""}
            ba, buf = QByteArray(), QBuffer(ba)
            buf.open(QBuffer.WriteOnly)
            qi.save(buf, "PNG")                                   # QImage→PNG bytes
            b64 = base64.b64encode(ba.data()).decode()
            return {"type": "image", "data": b64, "mimetype": "image/png"}
        text = cb.text()
        return {"type": "text", "data": text if text else ""}
    except Exception as e:
        print(f"Clipboard error: {e}")
        return {"type": "text", "data": ""}

# Anthropic messages builder
def build_messages(clip, question):
    blocks = []
    if clip["type"] == "text" and clip["data"].strip():
        blocks.append({
            "type": "text",
            "text": f"Context from clipboard:\\n{clip['data']}"
        })
    elif clip["type"] == "image":
        blocks.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": clip["mimetype"],
                "data": clip["data"]
            }
        })
    blocks.append({
        "type": "text",
        "text": f"Voice question: {question.strip()}"
    })
    return [{"role": "user", "content": blocks}]

def ask_claude(msgs):
    client = Anthropic()                                          # env ANTHROPIC_API_KEY
    r = client.messages.create(
        model=MODEL_NAME,
        max_tokens=1024,
        temperature=TEMP,
        system=SYSTEM_PROMPT,
        messages=msgs
    )
    return r.content[0].text

# Speak via macOS 'say'
class DictateClipVoice(QMainWindow):
    SAMPLE_RATE = 16000
    HOTKEY = {keyboard.KeyCode.from_char('`'), keyboard.KeyCode.from_char('2')}  # `+2

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dictate + Clipboard → Claude → Voice")
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        self.is_recording, self.frames = False, []
        self.speech_rate = SPEECH_RATE  # Default speech rate
        self._ui()
        self._hotkey()

    def speak(self, text):
        try:
            subprocess.run(["say", "-v", VOICE, "-r", str(self.speech_rate), text])
        except FileNotFoundError:
            print("macOS 'say' command not found. Install or use PyObjC fallback.")
        except Exception as e:
            print(f"Speech error: {e}")

    def _update_rate(self):
        self.speech_rate = self.rate_slider.value()
        self.rate_value.setText(f"{self.speech_rate} WPM")

# ──────────────────────────────────────────────────────

    def _ui(self):
        cw = QWidget()
        self.setCentralWidget(cw)
        lay = QVBoxLayout(cw)
        
        # Status label
        self.label = QLabel("Ready (press `+2)")
        lay.addWidget(self.label)
        
        # Record button
        btn = QPushButton("Hold to record")
        btn.setStyleSheet(
            "QPushButton{background:#4CAF50;color:#fff;padding:8px;border-radius:5px}"
            "QPushButton:pressed{background:#FF5733}"
        )
        lay.addWidget(btn)
        btn.pressed.connect(self.start_recording)
        btn.released.connect(self.stop_recording)
        
        # Speech rate slider
        rate_layout = QHBoxLayout()
        rate_label = QLabel("Speech Rate:")
        self.rate_slider = QSlider(Qt.Horizontal)
        self.rate_slider.setMinimum(100)
        self.rate_slider.setMaximum(400)
        self.rate_slider.setValue(self.speech_rate)
        self.rate_slider.setTickPosition(QSlider.TicksBelow)
        self.rate_slider.setTickInterval(50)
        self.rate_value = QLabel(f"{self.speech_rate} WPM")
        
        rate_layout.addWidget(rate_label)
        rate_layout.addWidget(self.rate_slider)
        rate_layout.addWidget(self.rate_value)
        lay.addLayout(rate_layout)
        
        self.rate_slider.valueChanged.connect(self._update_rate)

    def _hotkey(self):
        current = set()
        def press(k):
            if k in self.HOTKEY:
                current.add(k)
            if current == self.HOTKEY:
                self.toggle()
        def release(k):
            if k in current:
                current.remove(k)
        self.listener = keyboard.Listener(on_press=press, on_release=release)
        self.listener.start()

    # Audio stream callback
    def _cb(self, indata, frames, time_, status):
        if status:
            print(f"Stream error: {status}")
        if self.is_recording:
            self.frames.append(indata.copy())

    # Record control
    def start_recording(self):
        if self.is_recording:
            return
        self.frames.clear()
        self.is_recording = True
        self.label.setText("Recording…")
        try:
            self.stream = sd.InputStream(
                samplerate=self.SAMPLE_RATE,
                channels=1,
                dtype='float32',
                callback=self._cb,
                blocksize=1024
            )
            self.stream.start()
        except Exception as e:
            self.label.setText(f"Recording error: {e}")
            self.is_recording = False

    def stop_recording(self):
        if not self.is_recording:
            return
        self.is_recording = False
        try:
            self.stream.stop()
            self.stream.close()
        except Exception as e:
            print(f"Stream close error: {e}")
        self.label.setText("Processing…")
        QTimer.singleShot(0, self._process)

    def toggle(self):
        self.stop_recording() if self.is_recording else self.start_recording()

    # Pipeline
    def _process(self):
        try:
            # 1. Get clipboard content
            clip = grab_clipboard()
            self.label.setText("Transcribing…")

            # 2. Save and transcribe audio
            audio = np.concatenate(self.frames, axis=0)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as fp:
                with wave.open(fp, 'wb') as wf:
                    wf.setnchannels(1)
                    wf.setsampwidth(2)
                    wf.setframerate(self.SAMPLE_RATE)
                    wf.writeframes((audio * 32767).astype(np.int16).tobytes())
                wav_path = fp.name

            try:
                with open(wav_path, "rb") as af:
                    question = openai.Audio.transcribe("whisper-1", af)["text"]
                print(f"Transcription: {question}")
            finally:
                os.unlink(wav_path)

            # 3. Query Claude
            self.label.setText("Querying Claude…")
            answer = ask_claude(build_messages(clip, question))
            print(f"Claude's answer: {answer}")

            # 4. Copy to clipboard and speak
            pyperclip.copy(answer)
            
            # 5. Remove recording character and speak
            self.label.setText("Speaking answer…")
            kb = keyboard.Controller()
            
            # Delete the backtick and 2 that triggered recording
            for _ in range(2):  # Delete both ` and 2
                kb.press(keyboard.Key.backspace)
                kb.release(keyboard.Key.backspace)
            time.sleep(0.1)
            
            # Paste the response
            with kb.pressed(keyboard.Key.cmd):
                kb.press('v')
                kb.release('v')
            
            # Speak the response
            self.speak(answer)
            self.label.setText("Done. (`+2 to ask again)")

        except Exception as e:
            self.label.setText(f"Error: {e}")
            print("Error:", e)

    def closeEvent(self, ev):
        if getattr(self, 'listener', None):
            self.listener.stop()
        ev.accept()

# ──────────────────────────────────────────────────────
def main():
    if not (os.getenv("OPENAI_API_KEY") and os.getenv("ANTHROPIC_API_KEY")):
        print("Set OPENAI_API_KEY and ANTHROPIC_API_KEY first")
        return
    app = QApplication(sys.argv)
    win = DictateClipVoice()
    win.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

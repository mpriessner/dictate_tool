#!/usr/bin/env python3
# voice_shortcuts.py  —  unified tool with mode‑specific Claude prompts

from dotenv import load_dotenv; load_dotenv()

import os, sys, time, base64, tempfile, wave, subprocess
import numpy as np, sounddevice as sd, pyperclip
from PyQt5.QtCore import Qt, QTimer, QBuffer, QByteArray
from PyQt5.QtGui  import QImage
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QWidget, QVBoxLayout, QHBoxLayout, QSlider
from pynput import keyboard
import openai
from anthropic import Anthropic

# ───────────────────────── CONFIG & PROMPTS ─────────────────────────
SAMPLE_RATE   = 16000
VOICE         = os.getenv("VOICE", "Samantha")
VOICE_RATE    = int(os.getenv("VOICE_RATE", "220"))
MODEL_NAME    = "claude-3-7-sonnet-20250219"
TEMP          = 0.3

# keyboard combos - support both backtick (`) and greater-than (>) as triggers
TRIGGER_KEYS = [keyboard.KeyCode.from_char('`'), keyboard.KeyCode.from_char('<')]

# Define key combinations with both possible trigger keys
def make_combo(*nums):
    combos = []
    for trigger in TRIGGER_KEYS:
        for num in nums:
            combos.append({trigger, keyboard.KeyCode.from_char(str(num))})
    return combos

# Create combinations for each mode
COMBO_DICT = make_combo('3')    # ` or > + 3 for dictation
COMBO_TEXT = make_combo('1')    # ` or > + 1 for writing
COMBO_SPOKEN = make_combo('2')  # ` or > + 2 for speaking

# writing assistant prompt  (for ` + 1)
PROMPT_TEXT = """
You are a direct and concise writing assistant for Martin. Your task is to:

1. Determine the type of request:
   - If it sounds like a message/email response: Format as a proper reply
   - If it's a general request: Just fulfill it directly

2. For ALL requests:
   - Write in the same language as the input/request
   - Write directly from Martin's perspective without commentary
   - Be extremely concise while maintaining clarity
   - Respond without preamble or meta-commentary
   - Match the tone and style appropriate for the context
   - Preserve any relevant formatting from the clipboard content

3. For message/email responses ONLY:
   - End with an appropriate signature:
     • German informal → "Liebe Grüße,\nMartin"
     • English formal  → "Best wishes,\nMartin"
   - Format as a proper reply with appropriate greeting if needed

4. IMPORTANT: If the voice request starts with the word "ignore", completely disregard any clipboard content and only respond to what comes after "ignore" in the request
"""

# quick verbal answer prompt (for ` + 2)
PROMPT_SPOKEN = """
You are Martin's AI assistant. Provide quick, clear answers to verbal questions.
1. First line: briefly restate the understood question
2. Then give the direct answer
3. Be extremely concise
4. Match the language of the question
5. Reference clipboard content naturally when needed, UNLESS the question starts with the word "ignore"
6. IMPORTANT: If the question starts with "ignore", completely disregard any clipboard content and only respond to what comes after "ignore" in the question
7. No meta-commentary
"""

# ───────────────────────── HELPERS ─────────────────────────
def grab_clipboard():
    try:
        cb = QApplication.clipboard()
        if cb.mimeData().hasImage():
            img = cb.image()
            if img.isNull():
                return {"type": "text", "data": ""}
            ba = QByteArray()
            buf = QBuffer(ba)
            buf.open(QBuffer.WriteOnly)
            img.save(buf, "PNG")
            b64 = base64.b64encode(ba.data()).decode()
            return {"type": "image", "data": b64, "mimetype": "image/png"}
        text = cb.text()
        return {"type": "text", "data": text if text else ""}
    except Exception as e:
        print(f"Clipboard error: {e}")
        return {"type": "text", "data": ""}

def build_messages(clip, question):
    blocks = []
    # Check if question starts with 'ignore'
    question_text = question.strip()
    ignore_clipboard = question_text.lower().startswith("ignore")
    
    # If starting with 'ignore', remove it from the question and skip clipboard content
    if ignore_clipboard:
        # Remove 'ignore' and any following whitespace
        question_text = question_text[6:].lstrip()
    else:
        # Include clipboard content only if not ignoring
        if clip["type"] == "text" and clip["data"].strip():
            blocks.append({"type": "text", "text": f"Context from clipboard:\n{clip['data']}"})
        elif clip["type"] == "image":
            blocks.append({"type": "image",
                           "source": {"type": "base64", "media_type": clip["mimetype"], "data": clip["data"]}})
    
    # Add the question (with 'ignore' removed if it was present)
    blocks.append({"type": "text", "text": f"Voice request: {question_text}"})
    return [{"role": "user", "content": blocks}]

def ask_claude(sys_prompt, msgs):
    client = Anthropic()
    r = client.messages.create(model=MODEL_NAME, max_tokens=1024,
                               temperature=TEMP, system=sys_prompt, messages=msgs)
    return r.content[0].text

def speak(text, voice=VOICE, rate=VOICE_RATE):
    try:
        subprocess.run(["say", "-v", voice, "-r", str(rate), text])
    except Exception as e:
        print("say error:", e)

# ──────────────────────── MAIN WINDOW ────────────────────────
class VoiceTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Voice Shortcuts")
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        self.frames, self.is_recording, self.mode = [], False, None
        self.speech_rate = VOICE_RATE
        self._setup_ui()
        self._hotkey()
        
    def _setup_ui(self):
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Status label
        self.label = QLabel("Ready – `/> + 3 dict, `/> + 1 write, `/> + 2 speak")
        layout.addWidget(self.label)
        
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
        layout.addLayout(rate_layout)
        
        self.rate_slider.valueChanged.connect(self._update_rate)

    # Hot‑key listener
    def _hotkey(self):
        current = set()
        def on_press(k):
            if k in TRIGGER_KEYS or (hasattr(k, 'char') and k.char in '123'):
                current.add(k)
                
            # Check if current keys match any of our combinations
            for combo in COMBO_DICT:
                if current.issuperset(combo):
                    self._toggle("dict")
                    return
                    
            for combo in COMBO_TEXT:
                if current.issuperset(combo):
                    self._toggle("text")
                    return
                    
            for combo in COMBO_SPOKEN:
                if current.issuperset(combo):
                    self._toggle("spoken")
                    return
                    
        def on_release(k): current.discard(k)
        self.listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        self.listener.start()

    # Record control
    def _toggle(self, mode):
        if self.is_recording:
            self._stop()
        else:
            self.mode = mode
            self._start()

    def _start(self):
        self.frames.clear(); self.is_recording = True
        self.label.setText(f"Recording ({self.mode})…")
        self.stream = sd.InputStream(samplerate=SAMPLE_RATE, channels=1,
                                     dtype='float32', callback=self._cb, blocksize=1024)
        self.stream.start()

    def _stop(self):
        self.is_recording = False
        self.stream.stop(); self.stream.close()
        self.label.setText("Processing…")
        QTimer.singleShot(0, self._process)

    def _cb(self, indata, frames, *_):
        if self.is_recording:
            self.frames.append(indata.copy())

    # Pipeline
    def _process(self):
        try:
            # --- Whisper transcription
            audio = np.concatenate(self.frames, axis=0)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as fp:
                with wave.open(fp, 'wb') as wf:
                    wf.setnchannels(1); wf.setsampwidth(2); wf.setframerate(SAMPLE_RATE)
                    wf.writeframes((audio*32767).astype(np.int16).tobytes())
                wav_path = fp.name
            with open(wav_path, "rb") as af:
                question = openai.Audio.transcribe("whisper-1", af)["text"]
            os.unlink(wav_path)

            if self.mode == "dict":          # simple dictation branch
                self._paste(question, n_back=4)  # delete back-tick + 3
                self.label.setText("Dictation done.")
                return

            # --- Claude branch
            clip = grab_clipboard()
            sys_prompt = PROMPT_TEXT if self.mode == "text" else PROMPT_SPOKEN
            answer = ask_claude(sys_prompt, build_messages(clip, question))
            
            if self.mode == "text":
                # For writing mode, paste the answer
                self._paste(answer, n_back=4)    # delete back‑tick + 1/2 + additional chars
                self.label.setText("Done.")
            else:  # spoken mode
                # For spoken mode, just copy to clipboard and speak
                pyperclip.copy(answer)
                # Delete the trigger keys without pasting
                kb = keyboard.Controller()
                for _ in range(4):  # delete back‑tick + 2 + additional chars
                    kb.press(keyboard.Key.backspace); kb.release(keyboard.Key.backspace)
                # Speak the answer
                speak(answer, rate=self.speech_rate)
                self.label.setText("Answer copied to clipboard (not pasted).")
        except Exception as e:
            self.label.setText(f"Error: {e}")
            print("Error:", e)
            
    def _update_rate(self):
        self.speech_rate = self.rate_slider.value()
        self.rate_value.setText(f"{self.speech_rate} WPM")

    # helper: copy text & paste, removing arming keys
    def _paste(self, text, n_back):
        pyperclip.copy(text); time.sleep(0.05)
        kb = keyboard.Controller()
        for _ in range(n_back):
            kb.press(keyboard.Key.backspace); kb.release(keyboard.Key.backspace)
        with kb.pressed(keyboard.Key.cmd):
            kb.press('v'); kb.release('v')

    def closeEvent(self, ev):
        if getattr(self, 'listener', None): self.listener.stop()
        ev.accept()

# ───────────────────────────── main ─────────────────────────────
def main():
    if not (os.getenv("OPENAI_API_KEY") and os.getenv("ANTHROPIC_API_KEY")):
        print("Set OPENAI_API_KEY and ANTHROPIC_API_KEY"); return
    app = QApplication(sys.argv)
    win = VoiceTool(); win.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

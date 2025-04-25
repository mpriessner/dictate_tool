#!/usr/bin/env python3
# voice_shortcuts.py  —  unified tool with mode‑specific Claude prompts

# Make sure to load environment variables first, with verbose output
from dotenv import load_dotenv
env_loaded = load_dotenv()
print(f"Environment variables loaded from .env: {env_loaded}")

import os, sys, time, base64, tempfile, wave, subprocess
import numpy as np, sounddevice as sd, pyperclip
from PyQt5.QtCore import Qt, QTimer, QBuffer, QByteArray, QDateTime
from PyQt5.QtGui import QImage, QFont, QIcon, QPainter, QColor, QPen, QBrush
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QWidget, QVBoxLayout, 
                           QHBoxLayout, QSlider, QPushButton, QFrame, QSizePolicy, QAction, QMenu)
from pynput import keyboard
import openai
from anthropic import Anthropic
import json
import concurrent.futures
import threading

# Try to import Gemini API, but make it optional
HAS_GEMINI = False
try:
    import google.generativeai as genai
    HAS_GEMINI = True
except ImportError:
    print("Note: google-generativeai module not found. Gemini fallback will not be available.")

# ───────────────────────── CONFIG & PROMPTS ─────────────────────────
SAMPLE_RATE   = 16000
VOICE         = os.getenv("VOICE", "Samantha")
VOICE_RATE    = int(os.getenv("VOICE_RATE", "220"))

# AI Model Configuration
CLAUDE_MODEL  = "claude-3-7-sonnet-20250219"
GEMINI_MODEL  = "gemini-1.5-flash" # Use gemini-1.5-flash instead of 2.5-flash
OPENAI_MODEL  = "gpt-4o"
TEMP          = 0.3

# Initialize Gemini API if available
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or os.getenv("GOOGLE_GENERATIVE_AI_API_KEY")

# Debug environment variables
print(f"GEMINI_API_KEY present: {bool(os.getenv('GEMINI_API_KEY'))}")
print(f"GOOGLE_API_KEY present: {bool(os.getenv('GOOGLE_API_KEY'))}")
print(f"GOOGLE_GENERATIVE_AI_API_KEY present: {bool(os.getenv('GOOGLE_GENERATIVE_AI_API_KEY'))}")
print(f"Final GEMINI_API_KEY value present: {bool(GEMINI_API_KEY)}")

if HAS_GEMINI and GEMINI_API_KEY:
    print(f"Gemini API key found and configured")
    genai.configure(api_key=GEMINI_API_KEY)

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

# Anthropic Claude API
def ask_claude(sys_prompt, msgs):
    try:
        client = Anthropic()
        r = client.messages.create(
            model=CLAUDE_MODEL, 
            max_tokens=1024,
            temperature=TEMP, 
            system=sys_prompt, 
            messages=msgs
        )
        return r.content[0].text, "claude"
    except Exception as e:
        print(f"Claude API error: {e}")
        return None, "claude"

# Google Gemini API
def ask_gemini(sys_prompt, msgs):
    # Skip if Gemini is not available
    if not HAS_GEMINI:
        print("Gemini API not available: Module not installed")
        return None, "gemini"
        
    # Skip if API key is not available
    if not GEMINI_API_KEY:
        print("Gemini API not available: No API key found")
        return None, "gemini"
        
    # Ensure API key is configured
    genai.configure(api_key=GEMINI_API_KEY)
        
    try:
        # Configure Gemini model
        generation_config = {
            "temperature": TEMP,
            "max_output_tokens": 1024,
        }
        
        # Convert Anthropic message format to Gemini format
        prompt = f"{sys_prompt}\n\n"
        
        # Extract content from messages
        for msg in msgs:
            if msg["role"] == "user":
                for content in msg["content"]:
                    if content["type"] == "text":
                        prompt += f"{content['text']}\n"
                    elif content["type"] == "image":
                        # Skip images for now - we'll handle text only in fallback
                        prompt += "[Image content not processed in fallback]\n"
        
        # Create Gemini model
        model = genai.GenerativeModel(
            model_name=GEMINI_MODEL,
            generation_config=generation_config
        )
        
        # Generate response
        response = model.generate_content(prompt)
        return response.text, "gemini"
    except Exception as e:
        print(f"Gemini API error: {e}")
        return None, "gemini"

# OpenAI API
def ask_openai(sys_prompt, msgs):
    try:
        # Convert Anthropic message format to OpenAI format
        openai_msgs = []
        
        # Add system message
        openai_msgs.append({"role": "system", "content": sys_prompt})
        
        # Convert user messages
        for msg in msgs:
            if msg["role"] == "user":
                content_text = ""
                for content in msg["content"]:
                    if content["type"] == "text":
                        content_text += f"{content['text']}\n"
                    elif content["type"] == "image":
                        # Skip images for now - we'll handle text only in fallback
                        content_text += "[Image content not processed in fallback]\n"
                openai_msgs.append({"role": "user", "content": content_text})
        
        # Call OpenAI API (using the appropriate client based on version)
        try:
            # Try the new client first (OpenAI v1.0+)
            client = openai.OpenAI()
            response = client.chat.completions.create(
                model=OPENAI_MODEL,
                messages=openai_msgs,
                temperature=TEMP,
                max_tokens=1024
            )
            return response.choices[0].message.content, "openai"
        except (AttributeError, TypeError):
            # Fall back to legacy client
            response = openai.ChatCompletion.create(
                model=OPENAI_MODEL,
                messages=openai_msgs,
                temperature=TEMP,
                max_tokens=1024
            )
            return response.choices[0].message.content, "openai"
    except Exception as e:
        print(f"OpenAI API error: {e}")
        return None, "openai"

# Helper function to call API with timeout and proper cancellation
def call_with_timeout(func, *args, timeout=5):
    """Call a function with a timeout and cancel the thread if it times out"""
    # Use an event to signal cancellation
    cancel_event = threading.Event()
    result = [None, f"{func.__name__} (unknown error)"]
    
    def wrapped_func(*args):
        try:
            # Check if we should cancel before starting
            if cancel_event.is_set():
                return
                
            # Call the actual function
            nonlocal result
            result = func(*args)
        except Exception as e:
            print(f"Error in {func.__name__}: {e}")
            result = None, f"{func.__name__} (error: {str(e)[:100]})"
    
    # Start the function in a thread
    thread = threading.Thread(target=wrapped_func, args=args)
    thread.daemon = True  # Allow the thread to be killed when the program exits
    thread.start()
    
    # Wait for the thread to complete or timeout
    thread.join(timeout=timeout)
    
    # If the thread is still alive after the timeout, signal cancellation and return
    if thread.is_alive():
        print(f"Timeout after {timeout} seconds for {func.__name__}")
        cancel_event.set()  # Signal the thread to cancel
        return None, f"{func.__name__} (timeout)"
    
    return result

# Fallback mechanism
def get_ai_response(sys_prompt, msgs):
    # Set timeout in seconds
    API_TIMEOUT = 5
    
    # Try Gemini first
    print("\n--- Trying Gemini API ---")
    response, provider = call_with_timeout(ask_gemini, sys_prompt, msgs, timeout=API_TIMEOUT)
    if response:
        print(f"Using {provider.upper()} API response")
        return response
    
    # If Gemini fails, try Claude
    print(f"\n--- Gemini API failed ({provider}), trying Claude... ---")
    response, provider = call_with_timeout(ask_claude, sys_prompt, msgs, timeout=API_TIMEOUT)
    if response:
        print(f"Using {provider.upper()} API response")
        return response
    
    # If Claude fails, try OpenAI
    print(f"\n--- Claude API failed ({provider}), trying OpenAI... ---")
    response, provider = call_with_timeout(ask_openai, sys_prompt, msgs, timeout=API_TIMEOUT)
    if response:
        print(f"Using {provider.upper()} API response")
        return response
    
    # If all APIs fail, return error message
    return "Sorry, all AI providers are currently unavailable. Please try again later."

def speak(text, voice=VOICE, rate=VOICE_RATE):
    try:
        subprocess.run(["say", "-v", voice, "-r", str(rate), text])
    except Exception as e:
        print("say error:", e)

# ──────────────────────── MAIN WINDOW ────────────────────────
class FocusCircle(QWidget):
    """
    A custom widget that displays a colored circle that can be toggled
    between active (filled) and inactive (hollow) states.
    """
    def __init__(self, color="#00CCFF", parent=None):
        super().__init__(parent)
        self.color = QColor(color)
        self.active = False
        self.setCursor(Qt.PointingHandCursor)  # Change cursor to indicate clickable
        self.setToolTip("Click to minimize/maximize\nDouble-click to switch modes")
    
    def set_active(self, active):
        """Set the active state of the circle"""
        self.active = active
        self.update()
    
    def paintEvent(self, event):
        """Draw the circle"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Draw circle
        painter.setPen(Qt.NoPen)
        
        if self.active:
            # Filled circle when active
            painter.setBrush(self.color)
        else:
            # Hollow circle when inactive
            painter.setPen(self.color)
            painter.setBrush(Qt.NoBrush)
        
        painter.drawEllipse(1, 1, self.width() - 2, self.height() - 2)
    
    def mousePressEvent(self, event):
        """Handle mouse click on the focus circle to toggle UI state"""
        if event.button() == Qt.LeftButton:
            # Get the parent MainWindow instance
            parent = self.parent()
            while parent and not isinstance(parent, QMainWindow):
                parent = parent.parent()
                
            # Toggle UI state if parent is found
            if parent and hasattr(parent, 'ui_minimized'):
                if parent.ui_minimized:
                    parent.restore_ui()
                else:
                    parent.minimize_ui()
                    
            event.accept()
    
    def mouseDoubleClickEvent(self, event):
        """Handle double-click on the focus circle to toggle modes"""
        if event.button() == Qt.LeftButton:
            # Get the parent MainWindow instance
            parent = self.parent()
            while parent and not isinstance(parent, QMainWindow):
                parent = parent.parent()
                
            # Toggle mode if parent is found
            if parent:
                # Cycle through modes: text -> spoken -> dict -> text
                modes = ["text", "spoken", "dict"]
                if parent.mode:
                    try:
                        idx = modes.index(parent.mode)
                        next_mode = modes[(idx + 1) % len(modes)]
                    except ValueError:
                        next_mode = "text"
                else:
                    next_mode = "text"
                
                if hasattr(parent, "_toggle"):
                    parent._toggle(next_mode)
                
            event.accept()

class VoiceTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Voice Shortcuts")
        
        # Remove standard window frame and make it stay on top
        self.setWindowFlags(
            Qt.FramelessWindowHint |  # No frame
            Qt.WindowStaysOnTopHint |  # Stay on top
            Qt.Tool  # Doesn't show in taskbar
        )
        
        # Enable rounded corners by making window transparent
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        self.frames, self.is_recording, self.mode = [], False, None
        self.speech_rate = VOICE_RATE
        self.settings_visible = False
        self.ui_minimized = False  # Track if UI is minimized
        
        # Colors
        self.PRIMARY_COLOR = "#00CCFF"  # Blue
        self.SECONDARY_COLOR = "#FFCA28"  # Amber/yellow
        self.DICTATION_COLOR = "#FF5252"  # Red
        
        self._setup_ui()
        self._hotkey()
        
        # Set initial window position to top-left
        self.move(70, 50)  # Small offset from the very corner
        
        # Make window draggable
        self.old_pos = None
    
    def _setup_ui(self):
        # Set window style
        self.setStyleSheet("""
            QMainWindow {
                background-color: #252535;
                border-radius: 10px;
                border: 1px solid #333;
            }
            QPushButton { 
                background-color: transparent;
                color: white;
                border: none;
                padding: 5px;
            }
            QPushButton:hover { 
                background-color: rgba(255, 255, 255, 0.1);
            }
            QLabel { 
                color: white;
            }
            QSlider::groove:horizontal { 
                background: #555; height: 4px; 
            }
            QSlider::handle:horizontal { 
                background: #3498db; width: 16px; margin: -6px 0; border-radius: 8px; 
            }
        """)
        
        # Create central widget
        central_widget = QWidget()
        central_widget.setObjectName("centralWidget")
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 5, 10, 5)
        main_layout.setSpacing(0)
        main_layout.setAlignment(Qt.AlignVCenter)
        
        # Create focus circle
        self.focus_widget = FocusCircle(self.PRIMARY_COLOR)
        self.focus_widget.setFixedSize(19, 19)
        
        # Menu button
        self.menu_button = QPushButton("≡")
        self.menu_button.setFont(QFont("Arial", 14))
        self.menu_button.setFixedSize(22, 22)
        self.menu_button.clicked.connect(self._toggle_settings)
        
        # Status label
        self.status_label = QLabel("Ready")
        self.status_label.setFont(QFont("Arial", 12))
        self.status_label.setStyleSheet(f"color: {self.PRIMARY_COLOR};")
        
        # Add widgets to layout
        main_layout.addWidget(self.menu_button)
        main_layout.addWidget(self.focus_widget)
        main_layout.addWidget(self.status_label)
        
        # Set fixed size for the window
        self.setFixedSize(120, 32)
    
    def _toggle_settings(self):
        """Show the context menu"""
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: #252535;
                color: white;
                border: 1px solid #333;
                padding: 5px;
            }
            QMenu::item {
                padding: 5px 20px;
            }
            QMenu::item:selected {
                background-color: #3A3A4A;
            }
        """)
        
        # Add menu actions
        write_action = QAction("Write Mode (` + 1)", self)
        write_action.triggered.connect(lambda: self._toggle("text"))
        menu.addAction(write_action)
        
        speak_action = QAction("Speak Mode (` + 2)", self)
        speak_action.triggered.connect(lambda: self._toggle("spoken"))
        menu.addAction(speak_action)
        
        dict_action = QAction("Dictation Mode (` + 3)", self)
        dict_action.triggered.connect(lambda: self._toggle("dict"))
        menu.addAction(dict_action)
        
        # Speech rate submenu
        rate_menu = QMenu("Speech Rate", self)
        rate_menu.setStyleSheet(menu.styleSheet())
        
        for rate in [150, 200, 220, 250, 300]:
            rate_action = QAction(f"{rate} WPM", self)
            rate_action.triggered.connect(lambda checked, r=rate: self._set_rate(r))
            rate_menu.addAction(rate_action)
        
        menu.addMenu(rate_menu)
        
        # Toggle UI size
        if self.ui_minimized:
            expand_action = QAction("Expand UI", self)
            expand_action.triggered.connect(self.restore_ui)
            menu.addAction(expand_action)
        else:
            minimize_action = QAction("Minimize UI", self)
            minimize_action.triggered.connect(self.minimize_ui)
            menu.addAction(minimize_action)
        
        # Exit action
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        menu.addAction(exit_action)
        
        # Show menu at button position
        menu.exec_(self.menu_button.mapToGlobal(self.menu_button.rect().bottomLeft()))
    
    def _set_rate(self, rate):
        """Set speech rate directly"""
        self.speech_rate = rate
        global VOICE_RATE
        VOICE_RATE = rate
    
    def minimize_ui(self):
        """Shrink the window to just show the control elements"""
        if not self.ui_minimized:
            self.ui_minimized = True
            
            # Store original window size for restoration
            self.original_width = self.width()
            
            # Hide the status label
            self.status_label.hide()
            
            # Resize the window to be more compact
            self.setFixedWidth(60)
    
    def restore_ui(self):
        """Restore the window to its original size"""
        if self.ui_minimized:
            self.ui_minimized = False
            
            # Show all hidden elements
            self.status_label.show()
            
            # Restore original width
            self.setFixedWidth(self.original_width)
    
    def mousePressEvent(self, event):
        """Handle mouse press for window dragging"""
        if event.button() == Qt.LeftButton:
            self.old_pos = event.globalPos()
    
    def mouseMoveEvent(self, event):
        """Handle mouse movement for window dragging"""
        if self.old_pos:
            delta = event.globalPos() - self.old_pos
            self.move(self.pos() + delta)
            self.old_pos = event.globalPos()
    
    def mouseReleaseEvent(self, event):
        """Handle mouse release for window dragging"""
        if event.button() == Qt.LeftButton:
            self.old_pos = None
    
    # Hot‑key listener
    def _hotkey(self):
        current = set()
        def on_press(k):
            if k in TRIGGER_KEYS or isinstance(k, keyboard.KeyCode) and k.char in "123":
                current.add(k)
                if any(all(k in current for k in combo) for combo in COMBO_DICT):
                    self._toggle("dict")
                elif any(all(k in current for k in combo) for combo in COMBO_TEXT):
                    self._toggle("text")
                elif any(all(k in current for k in combo) for combo in COMBO_SPOKEN):
                    self._toggle("spoken")
        
        def on_release(k):
            try: current.remove(k)
            except: pass
        
        self.listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        self.listener.start()
    
    def _toggle(self, mode):
        """Toggle recording mode"""
        # If already in this mode, stop recording
        if self.mode == mode and self.is_recording:
            self._stop()
            return
        
        # Set mode and update status
        self.mode = mode
        
        # Update focus circle color based on mode
        if mode == "text":
            self.focus_widget.color = QColor(self.PRIMARY_COLOR)
            self.status_label.setStyleSheet(f"color: {self.PRIMARY_COLOR};")
            self.status_label.setText("Write")
        elif mode == "spoken":
            self.focus_widget.color = QColor(self.SECONDARY_COLOR)
            self.status_label.setStyleSheet(f"color: {self.SECONDARY_COLOR};")
            self.status_label.setText("Speak")
        elif mode == "dict":
            self.focus_widget.color = QColor(self.DICTATION_COLOR)
            self.status_label.setStyleSheet(f"color: {self.DICTATION_COLOR};")
            self.status_label.setText("Dictate")
        
        self.focus_widget.update()
        
        # Start recording
        self._start()
    
    def _start(self):
        """Start recording"""
        if not self.is_recording:
            self.frames = []
            self.is_recording = True
            self.focus_widget.set_active(True)
            self.stream = sd.InputStream(samplerate=SAMPLE_RATE, channels=1,
                                     dtype='float32', callback=self._cb, blocksize=1024)
            self.stream.start()
    
    def _stop(self):
        """Stop recording and process audio"""
        if self.is_recording:
            self.is_recording = False
            self.focus_widget.set_active(False)
            self.stream.stop()
            self.stream.close()
            self.status_label.setText("Processing...")
            QTimer.singleShot(0, self._process)
    
    def _cb(self, indata, frames, *_):
        if self.is_recording:
            self.frames.append(indata.copy())
    
    def _process(self):
        try:
            # Concatenate frames and convert to WAV
            audio = np.concatenate(self.frames)
            with tempfile.NamedTemporaryFile(suffix='.wav', delete=False) as f:
                with wave.open(f.name, 'wb') as wf:
                    wf.setnchannels(1)
                    wf.setsampwidth(2)
                    wf.setframerate(SAMPLE_RATE)
                    wf.writeframes((audio * 32767).astype(np.int16).tobytes())
                wav_path = f.name
                
            try:
                # Try the new client first (OpenAI v1.0+) with timeout
                client = openai.OpenAI(timeout=10)  # 10 second timeout for transcription
                with open(wav_path, "rb") as af:
                    response = client.audio.transcriptions.create(
                        model="whisper-1",
                        file=af
                    )
                    question = response.text
            except (AttributeError, TypeError):
                # Fall back to legacy client
                with open(wav_path, "rb") as af:
                    question = openai.Audio.transcribe("whisper-1", af)["text"]
            finally:
                os.unlink(wav_path)
            
            # Get clipboard content
            clip = grab_clipboard()
            
            # Delete the trigger keys from the input field
            with keyboard.Controller() as kb:
                for _ in range(4):  # Delete backtick + number
                    kb.press(keyboard.Key.backspace); kb.release(keyboard.Key.backspace)

            if self.mode == "dict":          # simple dictation branch
                self._paste(question, n_back=4)  # delete back-tick + 3
                self.status_label.setText("Dictation done.")
                return

            # --- AI response with fallback mechanism
            sys_prompt = PROMPT_TEXT if self.mode == "text" else PROMPT_SPOKEN
            self.status_label.setText("Querying AI services...")
            answer = get_ai_response(sys_prompt, build_messages(clip, question))
            
            if self.mode == "text":
                # For text mode, paste the answer
                self._paste(answer, n_back=4)    # delete back‑tick + 1/2 + additional chars
                self.status_label.setText("Done.")
            else:  # spoken mode
                # For spoken mode, just copy to clipboard and speak
                pyperclip.copy(answer)
                with keyboard.Controller() as kb:
                    kb.press(keyboard.Key.backspace); kb.release(keyboard.Key.backspace)
                # Speak the answer
                speak(answer, rate=self.speech_rate)
                self.status_label.setText("Answer copied to clipboard.")
        except Exception as e:
            self.status_label.setText(f"Error: {e}")
            print("Error:", e)
    
    def _update_rate(self):
        """This method is kept for compatibility but not used in the new UI"""
        pass
    
    def _update_time(self):
        """This method is kept for compatibility but not used in the new UI"""
        pass
    
    # helper: copy text & paste, removing arming keys
    def _paste(self, text, n_back):
        pyperclip.copy(text); time.sleep(0.05)
        with keyboard.Controller() as kb:
            for _ in range(n_back):
                kb.press(keyboard.Key.backspace); kb.release(keyboard.Key.backspace)
            kb.press(keyboard.Key.ctrl); kb.press('v'); kb.release('v'); kb.release(keyboard.Key.ctrl)
    
    def closeEvent(self, ev):
        try:
            self.listener.stop()
            if self.is_recording:
                sd.stop()
        except:
            pass
        super().closeEvent(ev)

# ───────────────────────────── main ─────────────────────────────
def main():
    # Check for required API keys (at least one LLM API key is required)
    has_anthropic = bool(os.getenv("ANTHROPIC_API_KEY"))
    has_gemini = bool(os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or os.getenv("GOOGLE_GENERATIVE_AI_API_KEY"))
    has_openai = bool(os.getenv("OPENAI_API_KEY"))
    
    if not has_openai:
        print("Warning: OPENAI_API_KEY not set. OpenAI fallback will not be available.")
    
    if not has_anthropic:
        print("Warning: ANTHROPIC_API_KEY not set. Claude will not be available.")
    
    if not has_gemini:
        print("Warning: GOOGLE_API_KEY not set. Gemini fallback will not be available.")
    
    # We need at least one LLM API key and the OpenAI key for Whisper
    if not has_openai:
        print("Error: OPENAI_API_KEY is required for Whisper transcription.")
        return
        
    if not (has_anthropic or has_gemini or has_openai):
        print("Error: At least one LLM API key (ANTHROPIC_API_KEY, GOOGLE_API_KEY, or OPENAI_API_KEY) is required.")
        return
    # Inform user about Gemini availability
    if not HAS_GEMINI:
        print("Note: To enable Gemini fallback, install the package with:")
        print("  pip install google-generativeai")
        
    app = QApplication(sys.argv)
    win = VoiceTool(); win.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

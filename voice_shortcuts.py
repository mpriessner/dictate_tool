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
                           QHBoxLayout, QSlider, QPushButton, QFrame, QSizePolicy, QAction, QMenu, QDesktopWidget)
from pynput import keyboard
import openai
from anthropic import Anthropic
import json
import concurrent.futures
import threading
import traceback

# Try to import Gemini API, but make it optional
HAS_GEMINI = False
try:
    import google.generativeai as genai
    from google.generativeai.types import HarmCategory, HarmBlockThreshold
    HAS_GEMINI = True
except ImportError:
    print("Note: google-generativeai module not found. Gemini fallback will not be available.")

# Try to import ElevenLabs API
HAS_ELEVENLABS = False
try:
    from elevenlabs.client import ElevenLabs
    HAS_ELEVENLABS = True
except ImportError:
    print("Note: elevenlabs module not found. ElevenLabs transcription will not be available.")

# ───────────────────────── CONFIG & PROMPTS ─────────────────────────
SAMPLE_RATE   = 16000
VOICE         = os.getenv("VOICE", "Samantha")
VOICE_RATE    = int(os.getenv("VOICE_RATE", "220"))

# AI Model Configuration
CLAUDE_MODEL  = "claude-3-7-sonnet-20250219"
GEMINI_MODEL  = "gemini-1.5-flash" # Use gemini-1.5-flash for LLM responses
GEMINI_TRANSCRIPTION_MODEL = "gemini-2.0-flash-exp"  # Use Gemini 2.0 Flash for audio transcription
OPENAI_MODEL  = "gpt-4o"
ELEVENLABS_TRANSCRIPTION_MODEL = "scribe_v1"  # ElevenLabs Speech-to-Text model
TEMP          = 0.3

# Transcription timeout (in seconds)
TRANSCRIPTION_TIMEOUT = 15

# Initialize Gemini API if available
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or os.getenv("GOOGLE_GENERATIVE_AI_API_KEY")

# Initialize ElevenLabs API if available
ELEVENLABS_API_KEY = os.getenv("ELEVENLABS_API_KEY")

# Debug environment variables
print(f"GEMINI_API_KEY present: {bool(os.getenv('GEMINI_API_KEY'))}")
print(f"ELEVENLABS_API_KEY present: {bool(ELEVENLABS_API_KEY)}")

if HAS_GEMINI and GEMINI_API_KEY:
    print(f"🤖 Gemini API key found and configured")
    genai.configure(api_key=GEMINI_API_KEY)

if HAS_ELEVENLABS and ELEVENLABS_API_KEY:
    print(f"🎤 ElevenLabs API key found and configured")
    ELEVENLABS_AVAILABLE = True
else:
    ELEVENLABS_AVAILABLE = False
    print("⚠️ ElevenLabs transcription disabled - API key not found or module not installed")

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
COMBO_CLEANUP = make_combo('4') # ` or > + 4 for text cleanup

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
   - Do NOT include a subject line unless the request explicitly asks for a subject.
   - End with an appropriate signature based on the language used *in the voice request*:
     • German informal → "Liebe Grüße,\nMartin"
     • English formal  → "Best wishes,\nMartin"
   - Format as a proper reply with appropriate greeting if needed

4. IMPORTANT: If the voice request starts with the word "ignore", completely disregard any clipboard content and only respond to what comes after "ignore" in the request
"""

# quick verbal answer prompt (for ` + 2)
PROMPT_SPOKEN = """
You are Martin's AI assistant. Provide clear answers to verbal questions.

1. By default, give direct, extremely concise answers (1-2 sentences maximum).

2. IMPORTANT: Check if the question ends with one of these detail level keywords:
   - "short": Give a concise answer (default behavior, 1-2 sentences)
   - "medium": Provide a moderately detailed answer (3-5 sentences with key points)
   - "detailed": Give a comprehensive answer with more context, examples, and explanation

3. Match the language of the question.

4. Reference clipboard content naturally when needed, UNLESS the question starts with the word "ignore".

5. IMPORTANT: If the question starts with "ignore", completely disregard any clipboard content and only respond to what comes after "ignore" in the question.

6. No meta-commentary or preambles like "here's a short answer" or "in summary".

7. If no detail level keyword is specified, default to the shortest, most concise response.
"""

# text cleanup and summarization prompt (for ` + 4)
PROMPT_CLEANUP = """
You are a text processing assistant that cleans and compacts text for efficient LLM consumption.

Your task is to take the provided text (which may contain terminal output, debug logs, code, data, or mixed content) and transform it into a clean, concise format suitable for feeding to another LLM.

RULES:
1. **Remove noise**: Strip out debug logs, stack traces, repetitive terminal output, timestamps, file paths that aren't essential
2. **Preserve code**: Keep code blocks intact and properly formatted
3. **Preserve data**: Keep structured data (JSON, tables, lists) but remove redundant entries
4. **Condense prose**: Summarize verbose explanations into key points
5. **Maintain context**: Ensure the cleaned text retains all essential information needed to understand the content
6. **No meta-commentary**: Do NOT add introductions like "Here's the cleaned text:" - just output the cleaned content directly
7. **Format for clarity**: Use markdown formatting (headers, lists, code blocks) to organize the output

If a voice instruction is provided, use it to guide what to focus on or what aspects to emphasize in the cleanup.

Output ONLY the cleaned and compacted text, ready to be pasted.
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

def transcribe_audio_with_elevenlabs(wav_path):
    """
    Transcribe audio using ElevenLabs Speech-to-Text API.
    
    Args:
        wav_path (str): Path to the WAV audio file
        
    Returns:
        str: Transcribed text, or None if transcription failed
    """
    if not HAS_ELEVENLABS:
        print("❌ ElevenLabs API not available: Module not installed")
        return None
        
    if not ELEVENLABS_API_KEY:
        print("❌ ElevenLabs API not available: No API key found")
        return None
        
    try:
        print("🎤 Transcribing audio with ElevenLabs...")
        
        # Initialize ElevenLabs client
        client = ElevenLabs(api_key=ELEVENLABS_API_KEY)
        
        # Open and read the audio file
        with open(wav_path, "rb") as audio_file:
            # Call the speech-to-text API
            # Note: Don't pass language_code at all for auto-detection
            transcription = client.speech_to_text.convert(
                file=audio_file,
                model_id=ELEVENLABS_TRANSCRIPTION_MODEL,
                tag_audio_events=False,  # Don't tag events like laughter
                diarize=False,  # Don't annotate speakers for single-speaker dictation
            )
        
        # Extract the text from the response
        if hasattr(transcription, 'text') and transcription.text:
            print(f"✅ ElevenLabs transcription successful")
            return transcription.text.strip()
        else:
            print(f"❌ ElevenLabs returned empty transcription")
            return None
            
    except Exception as e:
        print(f"❌ ElevenLabs transcription error: {str(e)}")
        traceback.print_exc()
        return None

def transcribe_audio_with_gemini(wav_path):
    """
    Transcribe audio using Gemini API (fallback method).
    
    Args:
        wav_path (str): Path to the WAV audio file
        
    Returns:
        str: Transcribed text, or None if transcription failed
    """
    if not HAS_GEMINI:
        print("❌ Gemini API not available: Module not installed")
        return None
        
    if not GEMINI_API_KEY:
        print("❌ Gemini API not available: No API key found")
        return None
        
    try:
        print("🎤 Transcribing audio with Gemini (fallback)...")
        
        # Upload audio file to Gemini
        audio_file = genai.upload_file(wav_path)
        
        # Wait for file to be processed
        while audio_file.state.name == "PROCESSING":
            print("⏳ Processing audio file...")
            time.sleep(0.5)
            audio_file = genai.get_file(audio_file.name)
        
        if audio_file.state.name == "FAILED":
            print("❌ Gemini audio processing failed")
            return None
            
        # Initialize Gemini model for transcription
        model = genai.GenerativeModel(GEMINI_TRANSCRIPTION_MODEL)
        
        # Generate transcription
        response = model.generate_content(
            [
                "Please transcribe this audio file accurately. Provide only the transcription text without any additional commentary, formatting, or preamble. Just the raw transcribed text.",
                audio_file
            ],
            safety_settings={
                HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
            }
        )
        
        # Clean up uploaded file
        try:
            genai.delete_file(audio_file.name)
        except:
            pass
            
        if response.text:
            print(f"✅ Gemini transcription successful")
            return response.text.strip()
        else:
            print(f"❌ Gemini returned empty transcription")
            return None
            
    except Exception as e:
        print(f"❌ Gemini transcription error: {str(e)}")
        traceback.print_exc()
        return None

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
                # Cycle through modes: text -> spoken -> dict -> cleanup -> text
                modes = ["text", "spoken", "dict", "cleanup"]
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
            Qt.WindowStaysOnTopHint   # Stay on top
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
        self.CLEANUP_COLOR = "#34C759"  # Green
        
        self._setup_ui()
        self._hotkey()
        
        # Set initial window position to top-left
        self.move(55, 60) # Small offset from the very corner
        
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
        self.status_text = QLabel("Ready")
        self.status_text.setFont(QFont("Arial", 12))
        self.status_text.setStyleSheet(f"color: {self.PRIMARY_COLOR};")
        
        # Add widgets to layout
        main_layout.addWidget(self.menu_button)
        main_layout.addWidget(self.focus_widget)
        main_layout.addWidget(self.status_text)
        
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

        cleanup_action = QAction("Cleanup Mode (` + 4)", self)
        cleanup_action.triggered.connect(lambda: self._toggle("cleanup"))
        menu.addAction(cleanup_action)
        
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
            self.status_text.hide()
            
            # Resize the window to be more compact
            self.setFixedWidth(60)
    
    def restore_ui(self):
        """Restore the window to its original size"""
        if self.ui_minimized:
            self.ui_minimized = False
            
            # Show all hidden elements
            self.status_text.show()
            
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
            if k is not None and (k in TRIGGER_KEYS or (hasattr(k, 'char') and k.char is not None and k.char in '1234')):
                current.add(k)
                if any(all(k in current for k in combo) for combo in COMBO_DICT):
                    self._toggle("dict")
                elif any(all(k in current for k in combo) for combo in COMBO_TEXT):
                    self._toggle("text")
                elif any(all(k in current for k in combo) for combo in COMBO_SPOKEN):
                    self._toggle("spoken")
                elif any(all(k in current for k in combo) for combo in COMBO_CLEANUP):
                    self._toggle("cleanup")
        
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
            self.status_text.setStyleSheet(f"color: {self.PRIMARY_COLOR};")
            self.status_text.setText("Write")
        elif mode == "spoken":
            self.focus_widget.color = QColor(self.SECONDARY_COLOR)
            self.status_text.setStyleSheet(f"color: {self.SECONDARY_COLOR};")
            self.status_text.setText("Speak")
        elif mode == "dict":
            self.focus_widget.color = QColor(self.DICTATION_COLOR)
            self.status_text.setStyleSheet(f"color: {self.DICTATION_COLOR};")
            self.status_text.setText("Dictate")
        elif mode == "cleanup":
            self.focus_widget.color = QColor(self.CLEANUP_COLOR)
            self.status_text.setStyleSheet(f"color: {self.CLEANUP_COLOR};")
            self.status_text.setText("Cleanup")

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
            self.status_text.setText("Processing...")
            QTimer.singleShot(0, self._process)
    
    def _cb(self, indata, frames, *_):
        if self.is_recording:
            self.frames.append(indata.copy())
    
    def _process(self):
        try:
            # Concatenate frames and convert to WAV
            audio = np.concatenate(self.frames, axis=0)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as f: 
                with wave.open(f, 'wb') as wf: # Changed f.name to f
                    wf.setnchannels(1); wf.setsampwidth(2); wf.setframerate(SAMPLE_RATE)
                    wf.writeframes((audio*32767).astype(np.int16).tobytes())
                wav_path = f.name
            
            # Transcribe audio using ElevenLabs API with Gemini fallback
            try:
                # Try ElevenLabs first
                question = transcribe_audio_with_elevenlabs(wav_path)
                
                # If ElevenLabs fails, try Gemini as fallback
                if not question:
                    print("🔄 ElevenLabs failed, trying Gemini fallback...")
                    question = transcribe_audio_with_gemini(wav_path)
                
                if not question:
                    self.status_text.setText("Transcription failed")
                    return
            finally:
                # Clean up the temporary WAV file
                try:
                    os.unlink(wav_path)
                except:
                    pass
            
            # Get clipboard content
            clip = grab_clipboard()
            
            # Delete the trigger keys from the input field
            kb = keyboard.Controller()
            for _ in range(4):  # Delete backtick + number
                kb.press(keyboard.Key.backspace); kb.release(keyboard.Key.backspace)

            if self.mode == "dict":          # simple dictation branch
                self._paste(question)  # paste after deleting trigger keys
                self.status_text.setText("Done.")
                self.activateWindow() # Try to bring window to front
                return

            # --- AI response with fallback mechanism
            if self.mode == "cleanup":
                sys_prompt = PROMPT_CLEANUP
            elif self.mode == "text":
                sys_prompt = PROMPT_TEXT
            else:  # spoken mode
                sys_prompt = PROMPT_SPOKEN

            self.status_text.setText("Querying AI services...")
            answer = get_ai_response(sys_prompt, build_messages(clip, question))

            if self.mode == "text" or self.mode == "cleanup":
                # For text and cleanup modes, paste the answer
                self._paste(answer)    # paste after deleting trigger keys
                self.status_text.setText("Done.")
            else:  # spoken mode
                # For spoken mode, just copy to clipboard and speak
                pyperclip.copy(answer)

                self._paste("", n_back=4)  # Only deletes 4 characters, does not paste anything
                speak(answer, rate=self.speech_rate)
                self.status_text.setText("Done.")
        except Exception as e:
            self.status_text.setText(f"Error: {e}")
            print("--- DETAILED ERROR TRACEBACK ---")
            traceback.print_exc()
            print("--------------------------------")
        finally:
            # Ensure window is attempted to be brought to front even after errors or success
            QTimer.singleShot(100, self.activateWindow)
    
    def _update_rate(self):
        """This method is kept for compatibility but not used in the new UI"""
        pass
    
    def _update_time(self):
        """This method is kept for compatibility but not used in the new UI"""
        pass
    
    # helper: copy text & paste, removing arming keys
    def _paste(self, text, n_back=0):
        pyperclip.copy(text); time.sleep(0.05)

        kb = keyboard.Controller()
        for _ in range(n_back):
            kb.press(keyboard.Key.backspace); kb.release(keyboard.Key.backspace)
        with kb.pressed(keyboard.Key.cmd):
            kb.press('v'); kb.release('v')

    def closeEvent(self, ev):
        try:
            self.listener.stop()
            if self.is_recording:
                sd.stop()
        except:
            pass
        super().closeEvent(ev)

    def _position_window(self):
        # Position window near top-right, shifted left
        screen_geo = QDesktopWidget().availableGeometry()
        window_width = self.width()
        target_x = screen_geo.width() - window_width - 5 # 5px left shift
        target_y = 50 # Position near the top
        self.move(target_x, target_y)

# ───────────────────────────── main ────────────────────────
def main():
    # Check for required API keys
    has_anthropic = bool(os.getenv("ANTHROPIC_API_KEY"))
    has_gemini = bool(os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or os.getenv("GOOGLE_GENERATIVE_AI_API_KEY"))
    has_openai = bool(os.getenv("OPENAI_API_KEY"))
    has_elevenlabs = bool(ELEVENLABS_API_KEY)
    
    # ElevenLabs is now PRIMARY for transcription, with Gemini as fallback
    if not has_elevenlabs:
        print("Warning: ELEVENLABS_API_KEY not found. ElevenLabs transcription will not be available.")
        if not has_gemini:
            print("Error: Neither ELEVENLABS_API_KEY nor GEMINI_API_KEY found.")
            print("At least one transcription API key is required.")
            print("Please add ELEVENLABS_API_KEY or GEMINI_API_KEY to your .env file.")
            return
        else:
            print("Using Gemini as primary transcription method.")
    
    if not HAS_ELEVENLABS and not HAS_GEMINI:
        print("Error: Neither elevenlabs nor google-generativeai module installed.")
        print("Please install at least one: pip install elevenlabs OR pip install google-generativeai")
        return
    
    # Check for LLM API keys (at least one is required)
    if not has_anthropic:
        print("Warning: ANTHROPIC_API_KEY not set. Claude will not be available.")
    
    if not has_openai:
        print("Warning: OPENAI_API_KEY not set. OpenAI fallback will not be available.")
    
    # We need at least one LLM API key for text processing
    if not (has_anthropic or has_gemini or has_openai):
        print("Error: At least one LLM API key (ANTHROPIC_API_KEY, GOOGLE_API_KEY, or OPENAI_API_KEY) is required.")
        return
    
    print("\n" + "="*60)
    print("🎤 Voice Shortcuts with ElevenLabs Transcription")
    print("="*60)
    if has_elevenlabs:
        print(f"Transcription Primary: ElevenLabs {ELEVENLABS_TRANSCRIPTION_MODEL}")
        if has_gemini:
            print(f"Transcription Fallback: Gemini {GEMINI_TRANSCRIPTION_MODEL}")
    else:
        print(f"Transcription: Gemini {GEMINI_TRANSCRIPTION_MODEL}")
    print(f"LLM Primary: {'Gemini' if has_gemini else 'Claude' if has_anthropic else 'OpenAI'}")
    print("="*60 + "\n")
        
    app = QApplication(sys.argv)
    win = VoiceTool()
    win.show()
    win.raise_()
    win.activateWindow()
    try:
        sys.exit(app.exec_())
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()

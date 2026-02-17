#!/usr/bin/env python3
"""
Windows Dictation Tool
Press Ctrl+Shift+D to start/stop recording. Transcribed text is pasted at cursor.
"""

from dotenv import load_dotenv
load_dotenv()

import os
import sys
import time
import tempfile
import wave
import threading
import numpy as np
import sounddevice as sd
import pyperclip
from openai import OpenAI
from pynput import keyboard
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QObject
from PyQt5.QtGui import QFont, QColor, QPainter, QPen, QBrush, QIcon
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLabel, QWidget,
    QHBoxLayout, QFrame, QDesktopWidget, QSystemTrayIcon, QMenu, QAction
)

# ── Config ──────────────────────────────────────────────────────────────────

SAMPLE_RATE = 16000
CHANNELS = 1
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")

if not OPENAI_API_KEY:
    print("WARNING: OPENAI_API_KEY not set. Add it to .env file.")


# ── Transcription ───────────────────────────────────────────────────────────

def transcribe_audio(wav_path: str) -> str:
    """Transcribe a WAV file using OpenAI Whisper API."""
    client = OpenAI(api_key=OPENAI_API_KEY)
    with open(wav_path, "rb") as f:
        response = client.audio.transcriptions.create(
            model="whisper-1",
            file=f,
            response_format="text",
        )
    return response.strip() if isinstance(response, str) else response.text.strip()


# ── Signals (thread-safe GUI updates) ──────────────────────────────────────

class Signals(QObject):
    status_changed = pyqtSignal(str)
    recording_changed = pyqtSignal(bool)
    flash_success = pyqtSignal()


# ── Status Dot Widget ──────────────────────────────────────────────────────

class StatusDot(QWidget):
    """Small colored dot: gray=idle, red=recording, green=done."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(18, 18)
        self._color = QColor("#888888")
        self._filled = False

    def set_state(self, color: str, filled: bool = False):
        self._color = QColor(color)
        self._filled = filled
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        rect = self.rect().adjusted(2, 2, -2, -2)
        if self._filled:
            p.setBrush(QBrush(self._color))
            p.setPen(Qt.NoPen)
        else:
            p.setBrush(Qt.NoBrush)
            p.setPen(QPen(self._color, 2))
        p.drawEllipse(rect)
        p.end()


# ── Main Window ────────────────────────────────────────────────────────────

class DictationWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # State
        self.is_recording = False
        self.frames = []
        self.signals = Signals()
        self.signals.status_changed.connect(self._set_status)
        self.signals.recording_changed.connect(self._set_recording_state)
        self.signals.flash_success.connect(self._flash_green)

        # Window setup
        self.setWindowTitle("Dictation")
        self.setWindowFlags(
            Qt.FramelessWindowHint
            | Qt.WindowStaysOnTopHint
            | Qt.Tool
        )
        self.setAttribute(Qt.WA_TranslucentBackground)

        # Central widget
        container = QFrame()
        container.setStyleSheet("""
            QFrame {
                background-color: rgba(30, 30, 30, 230);
                border-radius: 12px;
                border: 1px solid rgba(255, 255, 255, 40);
            }
        """)

        layout = QHBoxLayout(container)
        layout.setContentsMargins(12, 6, 14, 6)
        layout.setSpacing(8)

        # Dot
        self.dot = StatusDot()
        self.dot.set_state("#888888", False)
        layout.addWidget(self.dot)

        # Label
        self.label = QLabel("Ctrl+Shift+D")
        self.label.setFont(QFont("Segoe UI", 10))
        self.label.setStyleSheet("color: #cccccc; background: transparent; border: none;")
        layout.addWidget(self.label)

        self.setCentralWidget(container)
        self.setFixedHeight(36)
        self.adjustSize()

        # Position top-right
        self._position_window()

        # Drag support
        self._drag_pos = None

        # Keyboard listener
        self._setup_hotkey()

        # Auto-reset timer
        self._reset_timer = QTimer()
        self._reset_timer.setSingleShot(True)
        self._reset_timer.timeout.connect(self._reset_idle)

    # ── Window positioning & dragging ──────────────────────────────────

    def _position_window(self):
        screen = QDesktopWidget().availableGeometry()
        x = screen.width() - self.width() - 16
        y = 16
        self.move(x, y)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_pos = event.globalPos() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, event):
        if self._drag_pos and event.buttons() & Qt.LeftButton:
            self.move(event.globalPos() - self._drag_pos)

    def mouseReleaseEvent(self, event):
        self._drag_pos = None

    # ── GUI state updates (called from signals, always on main thread) ─

    def _set_status(self, text: str):
        self.label.setText(text)
        self.adjustSize()

    def _set_recording_state(self, recording: bool):
        if recording:
            self.dot.set_state("#FF4444", True)
            self.label.setText("Recording...")
        else:
            self.dot.set_state("#FFAA00", True)
            self.label.setText("Transcribing...")
        self.adjustSize()

    def _flash_green(self):
        self.dot.set_state("#44DD44", True)
        self.label.setText("Pasted")
        self.adjustSize()
        self._reset_timer.start(2000)

    def _reset_idle(self):
        self.dot.set_state("#888888", False)
        self.label.setText("Ctrl+Shift+D")
        self.adjustSize()

    # ── Hotkey ─────────────────────────────────────────────────────────

    def _setup_hotkey(self):
        combo = {keyboard.Key.ctrl_l, keyboard.Key.shift, keyboard.KeyCode.from_char('d')}
        self._hotkey_keys = set()
        self._combo = combo

        def on_press(key):
            # Normalize: treat ctrl_r as ctrl_l, shift_r as shift
            if key == keyboard.Key.ctrl_r:
                key = keyboard.Key.ctrl_l
            if key == keyboard.Key.shift_r:
                key = keyboard.Key.shift

            self._hotkey_keys.add(key)

            # Also check lowercase 'd' for the KeyCode
            try:
                if hasattr(key, 'char') and key.char and key.char.lower() == 'd':
                    self._hotkey_keys.add(keyboard.KeyCode.from_char('d'))
            except Exception:
                pass

            if self._combo.issubset(self._hotkey_keys):
                self._hotkey_keys.clear()
                self._on_toggle()

        def on_release(key):
            if key == keyboard.Key.ctrl_r:
                key = keyboard.Key.ctrl_l
            if key == keyboard.Key.shift_r:
                key = keyboard.Key.shift
            self._hotkey_keys.discard(key)
            try:
                if hasattr(key, 'char') and key.char and key.char.lower() == 'd':
                    self._hotkey_keys.discard(keyboard.KeyCode.from_char('d'))
            except Exception:
                pass

        self.listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        self.listener.daemon = True
        self.listener.start()

    def _on_toggle(self):
        if self.is_recording:
            self._stop_recording()
        else:
            self._start_recording()

    # ── Recording ──────────────────────────────────────────────────────

    def _start_recording(self):
        if self.is_recording:
            return
        self.is_recording = True
        self.frames = []
        self.signals.recording_changed.emit(True)

        def callback(indata, frame_count, time_info, status):
            if self.is_recording:
                self.frames.append(indata.copy())

        try:
            self.stream = sd.InputStream(
                samplerate=SAMPLE_RATE,
                channels=CHANNELS,
                dtype='float32',
                callback=callback,
                blocksize=1024,
            )
            self.stream.start()
        except Exception as e:
            print(f"Recording error: {e}")
            self.is_recording = False
            self.signals.status_changed.emit(f"Mic error")
            self._reset_timer.start(3000)

    def _stop_recording(self):
        if not self.is_recording:
            return
        self.is_recording = False
        self.signals.recording_changed.emit(False)

        try:
            self.stream.stop()
            self.stream.close()
        except Exception:
            pass

        if not self.frames:
            self.signals.status_changed.emit("No audio")
            self._reset_timer.start(2000)
            return

        audio = np.concatenate(self.frames, axis=0)
        duration = len(audio) / SAMPLE_RATE

        if duration < 0.5:
            self.signals.status_changed.emit("Too short")
            self._reset_timer.start(2000)
            return

        # Transcribe in background thread
        threading.Thread(target=self._transcribe_and_paste, args=(audio,), daemon=True).start()

    def _transcribe_and_paste(self, audio: np.ndarray):
        wav_path = None
        try:
            # Check audio level first
            peak = np.max(np.abs(audio))
            if peak < 0.005:
                print(f"WARNING: Audio peak is very low ({peak:.6f}). Check microphone.")
                self.signals.status_changed.emit("No speech detected")
                self._reset_timer.start(3000)
                return

            # Save to temp WAV (close file before wave.open for Windows compat)
            tmp = tempfile.NamedTemporaryFile(suffix=".wav", delete=False)
            wav_path = tmp.name
            tmp.close()

            pcm = (audio * 32767).astype(np.int16)
            with wave.open(wav_path, 'wb') as wf:
                wf.setnchannels(CHANNELS)
                wf.setsampwidth(2)  # 16-bit
                wf.setframerate(SAMPLE_RATE)
                wf.writeframes(pcm.tobytes())

            # Transcribe
            text = transcribe_audio(wav_path)

            if not text:
                self.signals.status_changed.emit("No speech detected")
                self._reset_timer.start(2000)
                return

            # Paste at cursor
            self._paste_text(text)
            self.signals.flash_success.emit()

        except Exception as e:
            print(f"Transcription error: {e}")
            self.signals.status_changed.emit("Error")
            self._reset_timer.start(3000)
        finally:
            if wav_path:
                try:
                    os.unlink(wav_path)
                except Exception:
                    pass

    def _paste_text(self, text: str):
        """Copy text to clipboard and paste via Ctrl+V."""
        pyperclip.copy(text)
        time.sleep(0.05)
        kb = keyboard.Controller()
        # Small delay to ensure hotkey keys are fully released
        time.sleep(0.15)
        with kb.pressed(keyboard.Key.ctrl):
            kb.press('v')
            kb.release('v')

    # ── Cleanup ────────────────────────────────────────────────────────

    def closeEvent(self, event):
        try:
            self.listener.stop()
        except Exception:
            pass
        if self.is_recording:
            try:
                self.stream.stop()
                self.stream.close()
            except Exception:
                pass
        super().closeEvent(event)


# ── Entry point ────────────────────────────────────────────────────────────

def main():
    if not OPENAI_API_KEY:
        print("=" * 50)
        print("ERROR: OPENAI_API_KEY not found!")
        print("")
        print("Create a .env file with:")
        print("  OPENAI_API_KEY=sk-your-key-here")
        print("")
        print("Or set the environment variable directly.")
        print("=" * 50)
        sys.exit(1)

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    window = DictationWindow()
    window.show()

    # System tray for easy exit
    tray = QSystemTrayIcon()
    tray_menu = QMenu()
    show_action = QAction("Show")
    show_action.triggered.connect(window.show)
    tray_menu.addAction(show_action)

    quit_action = QAction("Quit")
    quit_action.triggered.connect(app.quit)
    tray_menu.addAction(quit_action)

    tray.setContextMenu(tray_menu)
    tray.setToolTip("Dictation Tool - Ctrl+Shift+D")

    # Create a simple icon programmatically
    from PyQt5.QtGui import QPixmap
    pixmap = QPixmap(32, 32)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.setBrush(QBrush(QColor("#4488FF")))
    painter.setPen(Qt.NoPen)
    painter.drawEllipse(4, 4, 24, 24)
    painter.end()
    tray.setIcon(QIcon(pixmap))
    tray.show()

    print("Dictation Tool running. Press Ctrl+Shift+D to record.")
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

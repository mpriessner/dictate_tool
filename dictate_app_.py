import sys
import sounddevice as sd
import wave
import tempfile
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLabel
from PyQt5.QtCore import Qt, QTimer
import openai
import pyperclip
import keyboard
import numpy as np
import win32gui
import win32api
import win32con
import time
import ctypes
from ctypes import wintypes, windll, create_unicode_buffer

class DictateWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.is_recording = False
        self.is_recording_from_hotkey = False
        self.frames = []
        self.sample_rate = 16000
        self.transcribed_text = ""
        self.last_active_window = None
        
        # Create countdown timer
        self.countdown_timer = QTimer()
        self.countdown_timer.setSingleShot(True)
        self.countdown_timer.timeout.connect(self.paste_after_delay)
        
        # Register global hotkey (Ctrl+Alt+Down Arrow)
        keyboard.on_press_key("|", self.handle_hotkey, suppress=True)
        
        print("Initialization complete - press Ctrl+Alt+Down Arrow to start/stop recording")

    def init_ui(self):
        self.setWindowTitle("Dictation Tool")
        self.setGeometry(100, 100, 200, 100)
        
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Create status label
        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)
        
        # Create record button
        self.record_button = QPushButton("Press and Hold to Record")
        self.record_button.setStyleSheet(
            "QPushButton { background-color: #4CAF50; color: white; border-radius: 5px; padding: 10px; }"
            "QPushButton:pressed { background-color: #FF5733; }"
        )
        layout.addWidget(self.record_button)
        
        # Setup button events
        self.record_button.pressed.connect(self.start_recording)
        self.record_button.released.connect(self.stop_recording)
        
        # Set window flags to keep it always on top
        self.setWindowFlags(Qt.WindowStaysOnTopHint)

        # Test audio devices
        try:
            devices = sd.query_devices()
            input_device = sd.default.device[0]
            self.status_label.setText(f"Ready - Using input device: {devices[input_device]['name']}")
        except Exception as e:
            self.status_label.setText(f"Warning: Audio device error - {str(e)}")

    def get_focused_window(self):
        """Get the currently focused window handle and class name"""
        hwnd = win32gui.GetForegroundWindow()
        class_name = create_unicode_buffer(256)
        windll.user32.GetClassNameW(hwnd, class_name, 256)
        return hwnd, class_name.value

    def restore_window_focus(self, hwnd):
        """Restore focus to the previously active window"""
        if hwnd:
            # Bring the window to the foreground
            try:
                # Get current foreground window
                current_hwnd = win32gui.GetForegroundWindow()
                if current_hwnd != hwnd:
                    # Get current thread ID
                    current_thread = win32gui.GetWindowThreadProcessId(current_hwnd)[0]
                    # Get target thread ID
                    target_thread = win32gui.GetWindowThreadProcessId(hwnd)[0]
                    # Attach threads
                    windll.user32.AttachThreadInput(current_thread, target_thread, True)
                    # Set focus
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.SetFocus(hwnd)
                    # Detach threads
                    windll.user32.AttachThreadInput(current_thread, target_thread, False)
                    time.sleep(0.2)  # Give Windows time to switch focus
            except Exception as e:
                print(f"Error restoring focus: {e}")

    def paste_after_delay(self):
        """Paste the text after the timer expires"""
        try:
            # Get the currently focused window right when the timer expires
            self.last_active_window = win32gui.GetForegroundWindow()
            time.sleep(0.1)  # Small delay to ensure we have the right window
            
            # Stop the timer and paste
            self.countdown_timer.stop()
            self.paste_text()
        except Exception as e:
            print(f"Error in paste_after_delay: {e}")
            self.status_label.setText("Error pasting text")

    def paste_text(self):
        """Paste the text"""
        try:
            # For keyboard shortcut, we don't need window focus handling
            if not self.is_recording_from_hotkey and self.last_active_window:
                self.restore_window_focus(self.last_active_window)
                time.sleep(0.1)  # Small delay only for button press
            
            keyboard.press_and_release('ctrl+v')
            self.status_label.setText("Done!")
        except Exception as paste_error:
            print(f"Paste failed: {paste_error}")
            try:
                keyboard.write(self.transcribed_text)
                self.status_label.setText("Done!")
            except Exception as alt_error:
                print(f"Alternative paste failed: {alt_error}")
                self.status_label.setText("Warning: Automatic paste failed. Text is in clipboard.")

    def process_recording(self, was_from_hotkey=False):
        """Process the recorded audio and handle transcription"""
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.wav') as temp_wav:
                wav_file_path = temp_wav.name
                
                audio_data = np.concatenate(self.frames, axis=0)
                with wave.open(wav_file_path, 'wb') as wf:
                    wf.setnchannels(1)
                    wf.setsampwidth(2)  # 16-bit audio
                    wf.setframerate(self.sample_rate)
                    wf.writeframes((audio_data * 32767).astype(np.int16).tobytes())

                print(f"Audio saved to temporary file: {os.path.getsize(wav_file_path)} bytes")
                self.status_label.setText("Audio saved, sending to Whisper...")

            try:
                print("Starting Whisper transcription")
                with open(wav_file_path, "rb") as audio_file:
                    self.status_label.setText("Transcribing with Whisper...")
                    transcript = openai.Audio.transcribe(
                        "whisper-1",
                        audio_file
                    )
                
                self.transcribed_text = transcript["text"].strip()
                print(f"Transcription successful: '{self.transcribed_text}'")
                
                # Copy to clipboard
                pyperclip.copy(self.transcribed_text)
                
                if was_from_hotkey:
                    # For hotkey, paste immediately without any delays
                    print("Hotkey used - pasting immediately")
                    self.status_label.setText("Pasting...")
                    self.paste_text()
                else:
                    # For button press, use the 1-second delay
                    print("Button press - starting countdown")
                    self.status_label.setText("Click where you want to paste! 1 second...")
                    self.last_active_window = None  # Reset window for button press
                    self.countdown_timer.start(1000)
                
            except openai.error.AuthenticationError:
                error_msg = "OpenAI API key is invalid"
                print(f"Error: {error_msg}")
                self.status_label.setText(f"Error: {error_msg}")
            except openai.error.APIError as e:
                error_msg = f"OpenAI API error: {str(e)}"
                print(f"Error: {error_msg}")
                self.status_label.setText(f"Error: {error_msg}")
            
        except Exception as e:
            error_msg = f"Error: {str(e)}"
            print(f"Error in process_recording: {error_msg}")
            self.status_label.setText(error_msg)
        finally:
            if 'wav_file_path' in locals():
                os.unlink(wav_file_path)
                print("Temporary audio file deleted")

    def audio_callback(self, indata, frames, time, status):
        if status:
            print('Audio callback status:', status)
        if self.is_recording:
            self.frames.append(indata.copy())
            duration = len(self.frames) * frames / self.sample_rate
            self.status_label.setText(f"Recording... {duration:.1f}s")
            if len(self.frames) % 10 == 0:  # Print every 10 frames to avoid too much output
                print(f"Recording duration: {duration:.1f}s, Frames: {len(self.frames)}")

    def start_recording(self):
        print("\n--- Recording Debug ---")
        print(f"Starting recording (from hotkey: {self.is_recording_from_hotkey})")
        
        self.is_recording = True
        self.frames = []
        self.status_label.setText("Starting recording...")
        
        try:
            self.stream = sd.InputStream(
                channels=1,
                samplerate=self.sample_rate,
                callback=self.audio_callback,
                blocksize=1024
            )
            self.stream.start()
            print("Audio stream started successfully")
        except Exception as e:
            print(f"Error starting recording: {str(e)}")
            self.status_label.setText(f"Error starting recording: {str(e)}")
            self.is_recording = False
        print("--- End Recording Debug ---\n")

    def stop_recording(self):
        print("\n+++ Stop Recording Debug +++")
        print(f"Stop recording called (recording: {self.is_recording}, from hotkey: {self.is_recording_from_hotkey})")
        
        if not self.is_recording:
            print("Not recording, returning early")
            return
            
        self.is_recording = False
        self.status_label.setText("Processing...")
        
        if hasattr(self, 'stream'):
            print("Stopping audio stream")
            self.stream.stop()
            self.stream.close()
        
        if not self.frames:
            print("No frames recorded!")
            self.status_label.setText("Error: No audio recorded")
            return
            
        print(f"Number of frames recorded: {len(self.frames)}")
        was_from_hotkey = self.is_recording_from_hotkey  # Store the state
        
        # Process the recording with the hotkey state
        self.process_recording(was_from_hotkey)
        print("+++ End Stop Recording Debug +++\n")

    def start_countdown(self):
        """Start the countdown timer on the main thread"""
        print("Starting countdown timer on main thread")
        self.countdown_timer.start(1000)

    def handle_hotkey(self, e):
        """Handle the hotkey press event"""
        if keyboard.is_pressed('ctrl') and keyboard.is_pressed('alt'):
            # Use QTimer to ensure we're on the main Qt thread
            QTimer.singleShot(0, self.toggle_recording)

    def toggle_recording(self):
        """Toggle recording state when hotkey is pressed"""
        print("\n=== Keyboard Shortcut Debug ===")
        print(f"Current recording state: {self.is_recording}")
        print(f"Recording from hotkey: {self.is_recording_from_hotkey}")
        
        if not self.is_recording:
            print("Starting recording via keyboard shortcut")
            self.is_recording_from_hotkey = True
            self.start_recording()
        elif self.is_recording_from_hotkey:
            print("Stopping recording via keyboard shortcut")
            self.is_recording_from_hotkey = False
            self.stop_recording()
        print("=== End Keyboard Debug ===\n")

def main():
    # Set your OpenAI API key here
    openai.api_key = "api_key"
    
    app = QApplication(sys.argv)
    window = DictateWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

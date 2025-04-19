import sys
import sounddevice as sd
import wave
import tempfile
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLabel
from PyQt5.QtCore import Qt, QTimer
import openai
import pyperclip
import numpy as np
import time
from pynput import keyboard
import subprocess

class DictateWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.is_recording = False
        self.is_recording_from_hotkey = False
        self.frames = []
        self.sample_rate = 16000
        self.transcribed_text = ""
        self.transcription_start_time = 0
        
        # Create countdown timer
        self.countdown_timer = QTimer()
        self.countdown_timer.setSingleShot(True)
        self.countdown_timer.timeout.connect(self.paste_after_delay)
        
        # Add a status update timer to display Fn key state
        self.status_timer = QTimer()
        self.status_timer.setInterval(1000)  # Check every second
        self.status_timer.timeout.connect(self.update_fn_key_status)
        self.status_timer.start()
        
        # Setup hotkey listener
        self.hotkey_pressed = False
        # Use the '<' key (code 50) as the single toggle key
        self.toggle_key_code = 50  # '<' key code
        
        self.listener = keyboard.Listener(
            on_press=self.on_press,
            on_release=self.on_release
        )
        self.listener.start()
        
        print("Initialization complete - press the '<' key once to start recording, press again to stop")

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

    def on_press(self, key):
        """Handle key press events for hotkey detection"""
        try:
            # Check for the toggle key (code 50, '<' key)
            if hasattr(key, 'vk') and key.vk == self.toggle_key_code:
                # Toggle recording state
                self.toggle_recording()
        except Exception as e:
            print(f"Error in key detection: {e}")

    def on_release(self, key):
        """Handle key release events"""
        try:
            # We don't need to do anything special on key release
            # Recording is toggled by key press, not release
            pass
        except Exception as e:
            print(f"Error in key release detection: {e}")
            pass

    def paste_after_delay(self):
        """Paste the text after the timer expires"""
        try:
            # Small delay to ensure we have the right window
            time.sleep(0.1)
            
            # Stop the timer and paste
            self.countdown_timer.stop()
            
            # Hide the app window temporarily to ensure focus goes to the underlying application
            self.hide()
            time.sleep(0.2)  # Give time for focus to shift to the target application
            
            self.paste_text()
            
            # Show the window again after pasting
            self.show()
        except Exception as e:
            print(f"Error in paste_after_delay: {e}")
            self.status_label.setText("Error pasting text")
            self.show()  # Ensure window is shown even if there's an error

    def paste_text(self):
        """Paste the text"""
        try:
            # Use pynput to simulate cmd+v keystroke
            kb_controller = keyboard.Controller()
            
            # First, ensure we're properly focused on the target application
            time.sleep(0.2)  # Give time for focus to settle
            
            # Paste using cmd+v
            with kb_controller.pressed(keyboard.Key.cmd):
                kb_controller.press('v')
                kb_controller.release('v')
            
            # Add a small delay after pasting
            time.sleep(0.1)
                
            self.status_label.setText("Done!")
        except Exception as paste_error:
            print(f"Paste failed: {paste_error}")
            try:
                # Alternative: try using AppleScript to paste
                self.paste_with_applescript()
            except Exception as applescript_error:
                print(f"AppleScript paste failed: {applescript_error}")
                try:
                    # Last resort: try typing the text directly
                    kb_controller = keyboard.Controller()
                    kb_controller.type(self.transcribed_text)
                    self.status_label.setText("Done!")
                except Exception as alt_error:
                    print(f"Alternative paste failed: {alt_error}")
                    self.status_label.setText("Warning: Automatic paste failed. Text is in clipboard.")

    def paste_with_applescript(self):
        """Use AppleScript to paste text from clipboard"""
        try:
            # AppleScript to paste from clipboard at current cursor position
            script = '''
            tell application "System Events"
                keystroke "v" using command down
            end tell
            '''
            subprocess.run(["osascript", "-e", script], check=True)
            self.status_label.setText("Done with AppleScript!")
            return True
        except Exception as e:
            print(f"AppleScript paste error: {e}")
            return False
            
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

                file_size = os.path.getsize(wav_file_path)
                print(f"Audio saved to temporary file: {file_size} bytes")
                
                # Update status with file size to give user feedback on expected wait time
                duration_estimate = len(self.frames) * 1024 / self.sample_rate
                self.status_label.setText(f"Audio saved ({duration_estimate:.1f}s), sending to Whisper...")

            try:
                print("Starting Whisper transcription")
                
                # Start a timer to update the status periodically during transcription
                self.transcription_start_time = time.time()
                self.transcription_timer = QTimer()
                self.transcription_timer.setInterval(500)  # Update every 500ms
                self.transcription_timer.timeout.connect(self.update_transcription_status)
                self.transcription_timer.start()
                
                with open(wav_file_path, "rb") as audio_file:
                    self.status_label.setText("Transcribing with Whisper...")
                    transcript = openai.Audio.transcribe(
                      #  "gpt-4o-transcribe",
                        "whisper-1",
                        audio_file
                    )
                
                # Stop the timer once transcription is complete
                self.transcription_timer.stop()
                
                self.transcribed_text = transcript["text"].strip()
                print(f"Transcription successful: '{self.transcribed_text}'")
                
                # Copy to clipboard
                pyperclip.copy(self.transcribed_text)
                
                if was_from_hotkey:
                    # For hotkey, paste immediately without any delays
                    print("Hotkey used - pasting immediately")
                    self.status_label.setText("Pasting...")
                    # Hide window temporarily to ensure focus goes to the underlying app
                    self.hide()
                    time.sleep(0.2)  # Give time for focus to shift
                    self.paste_text()
                    self.show()
                else:
                    # For button press, use the 1-second delay
                    print("Button press - starting countdown")
                    self.status_label.setText("Click where you want to paste! 1 second...")
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
                try:
                    os.unlink(wav_file_path)
                    print("Temporary audio file deleted")
                except:
                    print("Failed to delete temporary audio file")

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
        print(f"Stop recording called (recording: {self.is_recording}, from hotkey: {self.is_recording_from_hotkey}")
        
        if not self.is_recording:
            print("Not recording, returning early")
            return
            
        self.is_recording = False
        duration = len(self.frames) * 1024 / self.sample_rate if self.frames else 0
        self.status_label.setText(f"Processing {duration:.1f}s of audio...")
        
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

    def toggle_recording(self):
        """Toggle recording state when hotkey is pressed"""
        print("\n=== Keyboard Shortcut Debug ===")
        print(f"Current recording state: {self.is_recording}")
        print(f"Recording from hotkey: {self.is_recording_from_hotkey}")
        
        if not self.is_recording:
            # First press - start recording
            print("Starting recording via keyboard shortcut")
            self.is_recording_from_hotkey = True
            self.status_label.setText("Recording started - Press Option+D again to stop")
            QTimer.singleShot(0, self.start_recording)
        elif self.is_recording:
            # Second press - stop recording and process
            print("Stopping recording via keyboard shortcut")
            self.is_recording_from_hotkey = False
            self.status_label.setText("Recording stopped - Processing...")
            QTimer.singleShot(0, self.stop_recording)
        print("=== End Keyboard Debug ===\n")
        
    def update_fn_key_status(self):
        """Update the status based on recording state"""
        if not self.is_recording:
            self.status_label.setText("Ready - Press the '<' key to start recording")
        elif self.is_recording and not self.status_label.text().startswith("Recording"):
            # Only update if not already showing a recording message
            self.status_label.setText("Recording in progress - Press the '<' key to stop")
    
    def update_transcription_status(self):
        """Update the status label with transcription progress"""
        elapsed = time.time() - self.transcription_start_time
        self.status_label.setText(f"Transcribing with Whisper... ({elapsed:.1f}s)")
        
        # If transcription is taking too long (over 30 seconds), provide more feedback
        if elapsed > 30:
            self.status_label.setText(f"Still transcribing... ({elapsed:.1f}s) Please wait.")
        if elapsed > 60:
            self.status_label.setText(f"Transcription taking longer than usual ({elapsed:.1f}s)")
    
    def detect_hotkey_state(self):
        """Try to detect hotkey state using alternative methods"""
        # This method is kept for compatibility but not actively used
        # since we're now using Option+D instead of Fn key
        return False
    
    def closeEvent(self, event):
        """Handle application close event"""
        if hasattr(self, 'listener') and self.listener.is_alive():
            self.listener.stop()
        if hasattr(self, 'status_timer'):
            self.status_timer.stop()
        event.accept()

def main():
    # Get OpenAI API key from environment variable
    openai.api_key = os.environ.get("OPENAI_API_KEY", "")
    
    # Check if API key is available
    if not openai.api_key:
        print("ERROR: OpenAI API key not found. Please set the OPENAI_API_KEY environment variable.")
        print("Example: export OPENAI_API_KEY='your-api-key-here'")
        return
    
    # Check for accessibility permissions
    try:
        # This will prompt for accessibility permissions if needed
        subprocess.run(["osascript", "-e", 'tell application "System Events" to return name of application processes'], 
                       capture_output=True, check=True)
    except subprocess.CalledProcessError:
        print("Warning: Accessibility permissions may be needed for keyboard functionality")
        print("Please go to System Preferences > Security & Privacy > Privacy > Accessibility")
        print("and add Terminal or your Python application to the list")
    
    app = QApplication(sys.argv)
    window = DictateWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
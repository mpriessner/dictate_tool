#!/usr/bin/env python3
# focus_timer.py - A work tracking tool with focus time tracking

import sys
import os
import time
from datetime import datetime, timedelta
import json
import csv
from pathlib import Path
import subprocess
from PyQt5.QtCore import Qt, QTimer, QSettings
from PyQt5.QtGui import QFont, QColor, QPainter, QPainterPath, QIcon
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QLabel, QPushButton, QMenu, QAction, QDialog, QFormLayout,
    QTimeEdit, QSpinBox, QComboBox, QCheckBox
)
from pynput import mouse, keyboard

class FocusTimer(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Load settings
        self.settings = QSettings("FocusTimer", "Settings")
        self.load_settings()
        
        # Timer variables
        self.focus_start_time = None
        self.work_hours = 0
        self.is_paused = False
        self.last_activity = time.time()
        self.elapsed_time_before_pause = 0  # Track elapsed time for resume functionality
        
        # Activity tracking variables
        self.activity_window = 5  # seconds to check for activity
        self.activity_events = []
        
        # UI state variables
        self.ui_minimized = False  # Track if UI is in minimized state
        
        # Set up activity listeners
        self.setup_activity_tracking()
        
        # Set up logging directory
        self.setup_logging()
        
        # Load today's work hours from logs
        self.load_todays_work_hours()
        
        # Enable rounded corners by making window transparent
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet("background:transparent;")
        
        # Setup UI
        self.setup_ui()
        
        # Start update timer (updates UI every second)
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self.update_display)
        self.update_timer.start(1000)
        
        # Initial display update
        self.update_display()
        
        # Auto-start focus timer when application launches
        self.start_focus()
        
        # Enable mouse tracking for hover effects
        self.setMouseTracking(True)
    
    def load_settings(self):
        # Load settings with defaults
        self.target_hours = self.settings.value("target_hours", 8, type=int)
        self.auto_pause = self.settings.value("auto_pause", False, type=bool)
        self.auto_pause_minutes = self.settings.value("auto_pause_minutes", 5, type=int)
        self.theme_color = self.settings.value("theme_color", "#00CCFF", type=str)
    
    def save_settings(self):
        # Save current settings
        self.settings.setValue("target_hours", self.target_hours)
        self.settings.setValue("auto_pause", self.auto_pause)
        self.settings.setValue("auto_pause_minutes", self.auto_pause_minutes)
        self.settings.setValue("theme_color", self.theme_color)
        self.settings.setValue("today_work_hours", self.work_hours)
        self.settings.setValue("last_date", datetime.now().strftime("%Y-%m-%d"))
    
    def setup_logging(self):
        """Set up the logging directory structure"""
        # Create logs directory in the same folder as the script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.logs_dir = os.path.join(script_dir, "focus_logs")
        os.makedirs(self.logs_dir, exist_ok=True)
        
        # Create today's directory
        today = datetime.now().strftime("%Y-%m-%d")
        self.today_dir = os.path.join(self.logs_dir, today)
        os.makedirs(self.today_dir, exist_ok=True)
        
        # Set up today's log file
        self.log_file = os.path.join(self.today_dir, "activity_log.csv")
        
        # Create the log file with headers if it doesn't exist
        if not os.path.exists(self.log_file):
            with open(self.log_file, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["timestamp", "active", "elapsed_time", "total_work_hours", "active_app", "note"])
    
    def get_active_application(self):
        """Get the name of the currently active application on macOS"""
        try:
            # Use AppleScript to get the name of the frontmost application
            cmd = ["osascript", "-e", "tell application \"System Events\" to get name of first application process whose frontmost is true"]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=1)
            app_name = result.stdout.strip()
            return app_name if app_name else "Unknown"
        except (subprocess.SubprocessError, subprocess.TimeoutExpired):
            return "Unknown"
    
    def log_activity(self, active, final=False, reset=False, pause=False, resume=False):
        """Log the current activity state"""
        # Check if we need to create a new log for a new day
        current_date = datetime.now().strftime("%Y-%m-%d")
        log_date = os.path.basename(os.path.dirname(self.log_file))
        
        # If the date has changed, set up new logging directory and file
        if current_date != log_date:
            print(f"Date changed from {log_date} to {current_date}, creating new log file")
            self.setup_logging()
            # Reset work hours for the new day
            if not reset:  # Only reset if not already being reset
                self.work_hours = 0
                self.settings.setValue("today_work_hours", 0)
                self.settings.setValue("last_date", current_date)
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Calculate elapsed time if focus is active
        elapsed_time = 0
        if self.focus_start_time and not self.is_paused:
            elapsed_time = time.time() - self.focus_start_time
        
        # Get the active application
        active_app = self.get_active_application()
        
        # Add note for special log entries
        note = ""
        if final:
            note = "final"
        elif reset:
            note = "reset"
        elif pause:
            note = "pause"
        elif resume:
            note = "resume"
            # For resume entries, show the elapsed time we're resuming from
            elapsed_time = self.elapsed_time_before_pause
        
        # Write to log file
        with open(self.log_file, 'a', newline='') as f:
            writer = csv.writer(f)
            if note:
                writer.writerow([timestamp, int(active), round(elapsed_time), round(self.work_hours, 2), active_app, note])
            else:
                writer.writerow([timestamp, int(active), round(elapsed_time), round(self.work_hours, 2), active_app])
    
    def load_todays_work_hours(self):
        """Load today's work hours from logs or settings"""
        today = datetime.now().strftime("%Y-%m-%d")
        last_date = self.settings.value("last_date", "", type=str)
        
        # If it's a new day, reset work hours
        if last_date != today:
            print(f"New day detected: {last_date} -> {today}")
            self.work_hours = 0
            self.settings.setValue("today_work_hours", 0)
            self.settings.setValue("last_date", today)
            
            # Ensure we have the correct log directory for today
            current_log_dir = os.path.dirname(self.log_file)
            expected_log_dir = os.path.join(self.logs_dir, today)
            
            # If the log directory doesn't match today's date, set up new logging
            if os.path.basename(current_log_dir) != today:
                print(f"Updating log directory for new day: {today}")
                self.setup_logging()
        else:
            # Calculate work hours from today's log file if it exists
            if os.path.exists(self.log_file):
                self.calculate_work_hours_from_log()
            else:
                # Fall back to settings if log file doesn't exist
                self.work_hours = self.settings.value("today_work_hours", 0, type=float)
    
    def calculate_work_hours_from_log(self):
        """Calculate total work hours from today's log file"""
        try:
            # Initialize work hours
            total_hours = 0
            last_active_time = None
            last_active_state = False
            
            # Read the log file
            with open(self.log_file, 'r', newline='') as f:
                reader = csv.reader(f)
                next(reader)  # Skip header row
                
                # Get all entries
                entries = list(reader)
                
                # Check for reset entries
                reset_index = -1
                for i, row in enumerate(entries):
                    # Check if this is a reset entry
                    if (len(row) >= 6 and row[5] == "reset") or (len(row) >= 5 and row[4] == "reset"):
                        reset_index = i
                
                # If we found a reset, only process entries after the reset
                if reset_index >= 0:
                    entries = entries[reset_index+1:]
                
                # Process active periods
                active_periods = []
                for row in entries:
                    if len(row) >= 4:
                        # Get timestamp and active state
                        timestamp_str = row[0]
                        is_active = int(row[1]) == 1
                        
                        try:
                            # Parse timestamp
                            timestamp = datetime.strptime(timestamp_str, "%Y-%m-%d %H:%M:%S")
                            
                            # If we have a transition from active to inactive, record the period
                            if last_active_state and not is_active and last_active_time is not None:
                                active_duration = (timestamp - last_active_time).total_seconds()
                                active_periods.append(active_duration)
                            
                            # If we have a transition to active, record the start time
                            if is_active and not last_active_state:
                                last_active_time = timestamp
                            
                            # Update last active state
                            last_active_state = is_active
                            
                        except ValueError:
                            # Skip entries with invalid timestamps
                            pass
                
                # If we're still active at the end, add the final period
                if last_active_state and last_active_time is not None:
                    now = datetime.now()
                    active_duration = (now - last_active_time).total_seconds()
                    active_periods.append(active_duration)
                
                # Sum up all active periods
                total_seconds = sum(active_periods)
                total_hours = total_seconds / 3600
            
            # Update work hours
            self.work_hours = total_hours
            
            # Save to settings
            self.settings.setValue("today_work_hours", self.work_hours)
            
            print(f"Calculated work hours from log: {self.work_hours:.2f} hours")
            
        except Exception as e:
            print(f"Error calculating work hours from log: {e}")
            # Fall back to settings
            self.work_hours = self.settings.value("today_work_hours", 0, type=float)
    
    def setup_ui(self):
        # Main window setup
        self.setWindowTitle("Focus Timer")
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        self.setFixedSize(360, 30)  # Reduced width and height
        
        # Store the original width for restoration
        self.original_width = 360
        
        # We'll use stylesheet for rounded corners instead of mask
        # as it's more compatible
        
        # Set stylesheet with theme color and rounded corners
        self.setStyleSheet("""
            QMainWindow { 
                background-color: #1E1E2A; 
                color: white;
                font-family: Arial;
                border-radius: 15px;
                border: 1px solid #333;
            }
            QWidget#centralWidget { 
                background-color: #1E1E2A; 
                color: white;
                font-family: Arial;
                border-radius: 15px;
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
        """)
        
        # Create central widget with object name for styling
        central_widget = QWidget()
        central_widget.setObjectName("centralWidget")
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 0, 10, 0)
        main_layout.setSpacing(0)  # No spacing to bring elements as close as possible
        
        # Focus widget and layout
        self.focus_widget = FocusCircle("#00CCFF")  # Blue color
        self.focus_widget.setFixedSize(22, 22)  # Slightly larger circle
        
        focus_layout = QVBoxLayout()
        focus_layout.setContentsMargins(0, 0, 0, 0)
        focus_layout.setSpacing(0)
        focus_layout.addWidget(self.focus_widget)
        
        # Focus time display
        self.focus_time_label = QLabel("00:00")
        self.focus_time_label.setFont(QFont("Arial", 16, QFont.Bold))  # Reduced by 2
        self.focus_time_label.setStyleSheet("color: #00CCFF;")  # Blue color
        self.focus_time_label.setAlignment(Qt.AlignCenter)
        
        # Focus label with reduced line spacing
        self.focus_label = QLabel("<html><div style='line-height:80%'>FOCUS TIME<br>ELAPSED</div></html>")
        self.focus_label.setFont(QFont("Arial", 7))  # Smaller font
        self.focus_label.setStyleSheet("color: #888; padding-top: 5px;")  # Added padding to move text down
        self.focus_label.setAlignment(Qt.AlignLeft)
        
        # Work time display
        self.work_time_label = QLabel("0 hr 0 min")
        self.work_time_label.setFont(QFont("Arial", 16, QFont.Bold))  # Reduced by 2
        self.work_time_label.setStyleSheet("color: #9370DB;")  # Purple color
        self.work_time_label.setAlignment(Qt.AlignCenter)
        
        # Work label with reduced line spacing
        self.work_label = QLabel("<html><div style='line-height:80%'>WORK<br>HOURS</div></html>")
        self.work_label.setFont(QFont("Arial", 7))  # Smaller font
        self.work_label.setStyleSheet("color: #888; padding-top: 5px;")  # Added padding to move text down
        self.work_label.setAlignment(Qt.AlignLeft)
        
        # Percent display (replacing the separator dashes)
        self.percent_display = QLabel("0%")
        self.percent_display.setFont(QFont("Arial", 16, QFont.Bold))  # Same size as other numbers
        self.percent_display.setStyleSheet("color: #888;")
        self.percent_display.setAlignment(Qt.AlignCenter)
        
        # Percent of day with reduced line spacing (now just the label)
        self.percent_label = QLabel("<html><div style='line-height:80%'>PERCENT<br>OF DAY</div></html>")
        self.percent_label.setFont(QFont("Arial", 7))  # Smaller font
        self.percent_label.setStyleSheet("color: #888; padding-top: 5px;")  # Added padding to move text down
        self.percent_label.setAlignment(Qt.AlignLeft)
        
        # Menu button
        self.menu_button = QPushButton("≡")
        self.menu_button.setFont(QFont("Arial", 14))
        self.menu_button.setFixedSize(22, 22)
        self.menu_button.clicked.connect(self.show_menu)
        
        # Add widgets to layout - menu button first, then focus widget
        main_layout.addWidget(self.menu_button)
        main_layout.addWidget(self.focus_widget)
        main_layout.addWidget(self.focus_time_label)
        main_layout.addWidget(self.focus_label)
        main_layout.addWidget(self.work_time_label)
        main_layout.addWidget(self.work_label)
        
        # Add spacing before percent display to move it further right
        spacer = QWidget()
        spacer.setFixedWidth(10)  # Adjust this value to control how far right it moves
        main_layout.addWidget(spacer)
        
        main_layout.addWidget(self.percent_display)
        main_layout.addWidget(self.percent_label)
        
        # Adjust widget widths to move labels closer to numbers
        self.focus_time_label.setFixedWidth(60)
        self.focus_label.setFixedWidth(40)  # Reduced width to bring work hours closer
        self.work_time_label.setFixedWidth(100)
        self.work_label.setFixedWidth(30)
        self.percent_display.setFixedWidth(40)
        self.percent_label.setFixedWidth(35)
        
        # Make window draggable
        self.old_pos = None
        
        # Set initial work time display
        self.format_work_time()
    
    def start_focus(self):
        """Start a new focus session (resets the timer)"""
        # Reset the focus timer
        self.focus_start_time = time.time()
        self.is_paused = False
        self.focus_widget.set_active(True)
        self.last_activity = time.time()
        
        # Reset elapsed time before pause when starting a new session
        self.elapsed_time_before_pause = 0
        
    def resume_focus(self):
        """Resume a paused focus session"""
        # Only resume if we're paused
        if self.is_paused:
            # Adjust the focus start time to account for the elapsed time before pause
            # This makes the timer continue from where it left off
            self.focus_start_time = time.time() - self.elapsed_time_before_pause
            self.is_paused = False
            self.focus_widget.set_active(True)
            self.last_activity = time.time()
            
            # Log the resume
            self.log_activity(True, resume=True)
    
    def pause_focus(self):
        """Pause the current focus session"""
        if self.focus_start_time and not self.is_paused:
            # Calculate elapsed time and add to work hours
            elapsed = time.time() - self.focus_start_time
            self.add_work_time(elapsed / 3600)  # Convert seconds to hours
            
            # Store the current elapsed time for resume functionality
            self.elapsed_time_before_pause = elapsed
            
            # Update state - keep focus_start_time as None to indicate we're fully paused
            self.is_paused = True
            self.focus_widget.set_active(False)
            self.focus_start_time = None
            
            # Log the pause
            self.log_activity(False, pause=True)
    
    def toggle_focus(self):
        """Toggle between focus and pause states"""
        if self.focus_start_time and not self.is_paused:
            # If active, pause
            self.pause_focus()
        elif self.is_paused:
            # If paused, resume
            self.resume_focus()
        else:
            # If stopped, start new
            self.start_focus()
    
    def start_timer(self):
        if not self.focus_start_time:
            self.focus_start_time = time.time()
            self.focus_widget.set_active(True)
    
    def pause_timer(self):
        if self.focus_start_time:
            self.is_paused = True
            self.focus_widget.set_active(False)
            # Add elapsed time to total
            elapsed = time.time() - self.focus_start_time
            self.add_work_time(elapsed / 3600)
            self.focus_start_time = None
    
    def reset_timer(self):
        """Reset the day's work hours and log the reset"""
        self.pause_timer()
        
        # Log final entry with previous work hours
        self.log_activity(False, final=True)
        
        # Reset work hours
        self.work_hours = 0
        
        # Update settings
        self.settings.setValue("today_work_hours", 0)
        
        # Log a new entry with reset work hours
        self.log_activity(False, reset=True)
        
        # Update display
        self.format_work_time()
        self.update_display()
    
    def format_work_time(self):
        """Format work time in hours and minutes"""
        total_hours = self.work_hours
        hours = int(total_hours)
        minutes = int((total_hours - hours) * 60)
        self.work_time_label.setText(f"{hours} hr {minutes} min")
    
    def add_work_time(self, hours):
        self.work_hours += hours
        self.format_work_time()
        self.update_display()
        # Save updated work hours to settings
        self.settings.setValue("today_work_hours", self.work_hours)
    
    def setup_activity_tracking(self):
        """Set up global mouse and keyboard listeners"""
        # Start mouse listener
        self.mouse_listener = mouse.Listener(
            on_move=self.on_mouse_move,
            on_click=self.on_mouse_click,
            on_scroll=self.on_mouse_scroll
        )
        self.mouse_listener.start()
        
        # Start keyboard listener
        self.keyboard_listener = keyboard.Listener(
            on_press=self.on_key_press,
            on_release=self.on_key_release
        )
        self.keyboard_listener.start()
    
    def on_mouse_move(self, x, y):
        self.record_activity()
    
    def on_mouse_click(self, x, y, button, pressed):
        self.record_activity()
    
    def on_mouse_scroll(self, x, y, dx, dy):
        self.record_activity()
    
    def on_key_press(self, key):
        self.record_activity()
    
    def on_key_release(self, key):
        self.record_activity()
    
    def record_activity(self):
        """Record an activity event with timestamp"""
        current_time = time.time()
        self.last_activity = current_time
        self.activity_events.append(current_time)
        
        # Clean up old events outside the activity window
        cutoff_time = current_time - self.activity_window
        self.activity_events = [t for t in self.activity_events if t >= cutoff_time]
    
    def is_active_in_window(self):
        """Check if there was activity in the activity window"""
        if not self.activity_events:
            return False
            
        current_time = time.time()
        cutoff_time = current_time - self.activity_window
        
        # Check if there are any events in the window
        return any(t >= cutoff_time for t in self.activity_events)
    
    def update_display(self):
        # Check for day change
        current_date = datetime.now().strftime("%Y-%m-%d")
        log_date = os.path.basename(os.path.dirname(self.log_file))
        
        # If the date has changed, update logging and reset work hours
        if current_date != log_date:
            print(f"Day changed during runtime: {log_date} -> {current_date}")
            # Log final entry for previous day
            self.log_activity(False, final=True)
            # Set up new logging for the new day
            self.setup_logging()
            # Reset work hours for the new day
            self.work_hours = 0
            self.settings.setValue("today_work_hours", 0)
            self.settings.setValue("last_date", current_date)
            # Update display with reset hours
            self.format_work_time()
        else:
            # Periodically recalculate work hours from log file (every 5 minutes)
            current_minute = int(time.time() / 60)
            if not hasattr(self, 'last_recalculated_minute') or current_minute - self.last_recalculated_minute >= 5:
                self.last_recalculated_minute = current_minute
                # Recalculate work hours from log
                self.calculate_work_hours_from_log()
                # Update work time display
                self.format_work_time()
        
        # Update focus time
        if self.focus_start_time and not self.is_paused:
            elapsed = time.time() - self.focus_start_time
            hours, remainder = divmod(elapsed, 3600)
            minutes, seconds = divmod(remainder, 60)
            self.focus_time_label.setText(f"{int(minutes):02d}:{int(seconds):02d}")
            
            # Log activity every minute
            current_minute = int(time.time() / 60)
            if not hasattr(self, 'last_logged_minute') or current_minute > self.last_logged_minute:
                self.last_logged_minute = current_minute
                # Use the 5-second window to determine if active
                is_active = self.is_active_in_window()
                self.log_activity(is_active)  # Log based on activity window
            
            # Auto-pause after inactivity
            if self.auto_pause and time.time() - self.last_activity > self.auto_pause_minutes * 60:
                self.pause_focus()
        elif self.is_paused or not self.focus_start_time:
            # Log inactivity every minute when paused
            current_minute = int(time.time() / 60)
            if not hasattr(self, 'last_logged_minute') or current_minute > self.last_logged_minute:
                self.last_logged_minute = current_minute
                # Still check activity window even when paused
                is_active = self.is_active_in_window()
                self.log_activity(is_active)
        
        # Calculate percent of day (allow going over 100%)
        percent = (self.work_hours / self.target_hours) * 100
        display_percent = min(100, percent)  # Cap at 100% for display purposes
        
        # Update the percent display
        if percent > 100:
            self.percent_display.setText(f"{percent:.0f}%")
        else:
            self.percent_display.setText(f"{display_percent:.0f}%")
        
        # Create a smooth color gradient from red to green based on progress
        # Start with red (255,0,0) and transition to green (0,255,0)
        if percent <= 50:
            # Red to Yellow: Increase green component
            red = 255
            green = int((percent / 50) * 255)
            blue = 0
        else:
            # Yellow to Green: Decrease red component
            red = int(255 - ((percent - 50) / 50) * 255)
            green = 255
            blue = 0
            
        # If over 100%, add a blue component for a special color
        if percent > 100:
            blue = min(255, int((percent - 100) * 2.55))
            
        # Ensure values are within valid range
        red = max(0, min(255, red))
        green = max(0, min(255, green))
        blue = max(0, min(255, blue))
        
        # Set the color
        color_hex = f"#{red:02x}{green:02x}{blue:02x}"
        self.percent_display.setStyleSheet(f"color: {color_hex}; font-weight: bold;")
    
    def show_menu(self):
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
        
        # Focus actions
        if self.focus_start_time and not self.is_paused:
            # Currently active - show pause option
            pause_action = QAction("Pause Focus", self)
            pause_action.triggered.connect(self.pause_focus)
            menu.addAction(pause_action)
        elif self.is_paused:
            # Currently paused - show both resume and restart options
            resume_action = QAction("Resume Focus", self)
            resume_action.triggered.connect(self.resume_focus)
            menu.addAction(resume_action)
            
            start_action = QAction("Start New Focus", self)
            start_action.triggered.connect(self.start_focus)
            menu.addAction(start_action)
        else:
            # Not started - show start option
            start_action = QAction("Start Focus", self)
            start_action.triggered.connect(self.start_focus)
            menu.addAction(start_action)
        
        # Settings
        settings_action = QAction("Settings", self)
        settings_action.triggered.connect(self.show_settings)
        menu.addAction(settings_action)
        
        # Reset
        reset_action = QAction("Reset Day", self)
        reset_action.triggered.connect(self.reset_timer)
        menu.addAction(reset_action)
        
        # Exit
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        menu.addAction(exit_action)
        
        # Show menu at button position
        menu.exec_(self.menu_button.mapToGlobal(self.menu_button.rect().bottomLeft()))
    
    def show_settings(self):
        dialog = SettingsDialog(self)
        if dialog.exec_():
            # Update settings from dialog
            self.target_hours = dialog.target_hours_spin.value()
            self.auto_pause = dialog.auto_pause_check.isChecked()
            self.auto_pause_minutes = dialog.auto_pause_minutes_spin.value()
            
            # Get theme color
            color_name = dialog.theme_color_combo.currentText()
            self.theme_color = dialog.theme_colors[color_name]
            
            # Update display
            self.setFixedHeight(80)  # Height with settings panel
        else:
            self.setFixedHeight(32)   # Height without settings panel
    
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.old_pos = event.globalPos()
            self.record_activity()
    
    def mouseMoveEvent(self, event):
        if self.old_pos:
            delta = event.globalPos() - self.old_pos
            self.move(self.pos() + delta)
            self.old_pos = event.globalPos()
            self.record_activity()
    
    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.old_pos = None
    
    # Removed roundedRect method as we're using stylesheet approach instead
    
    def enterEvent(self, event):
        """Handle mouse entering the window area"""
        # Minimize UI when mouse enters the window
        self.minimize_ui()
        event.accept()
    
    def leaveEvent(self, event):
        """Handle mouse leaving the window area"""
        # Restore UI when mouse leaves the window
        self.restore_ui()
        event.accept()
    
    def minimize_ui(self):
        """Shrink the window to just show the control elements"""
        if not self.ui_minimized:
            self.ui_minimized = True
            
            # Store original window size for restoration
            self.original_width = self.width()
            
            # Hide all the labels and numbers
            self.focus_time_label.hide()
            self.focus_label.hide()
            self.work_time_label.hide()
            self.work_label.hide()
            self.percent_display.hide()
            self.percent_label.hide()
            
            # Keep the focus widget and menu button visible
            self.focus_widget.show()
            self.menu_button.show()
            
            # Adjust spacing for minimized mode
            self.centralWidget().layout().setContentsMargins(10, 0, 10, 0)
            
            # Shrink the window to just show the control elements
            self.setFixedSize(60, 32)  # Width just enough for focus circle and menu button
    
    def restore_ui(self):
        """Restore the window to its original size and show all UI elements"""
        if self.ui_minimized:
            self.ui_minimized = False
            
            # Restore original window size
            self.setFixedSize(self.original_width, 32)  # Use stored original width
            
            # Restore original layout margins
            self.centralWidget().layout().setContentsMargins(10, 0, 10, 0)
            
            # Show all the labels and numbers
            self.focus_time_label.show()
            self.focus_label.show()
            self.work_time_label.show()
            self.work_label.show()
            self.percent_display.show()
            self.percent_label.show()
    
    def closeEvent(self, event):
        # Save settings on close
        self.save_settings()
        
        # Stop focus if active
        if self.focus_start_time and not self.is_paused:
            self.pause_timer()
        
        # Log final activity state
        is_active = self.is_active_in_window()
        self.log_activity(is_active, final=True)
        
        # Print confirmation
        print(f"Closing application. Final work hours: {self.work_hours:.2f}")
        
        # Stop listeners
        self.mouse_listener.stop()
        self.keyboard_listener.stop()
            
        event.accept()


class FocusCircle(QWidget):
    def __init__(self, color="#00CCFF"):
        super().__init__()
        self.color = QColor(color)
        self.active = False
    
    def set_active(self, active):
        self.active = active
        self.update()
    
    def paintEvent(self, event):
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


class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setWindowTitle("Settings")
        self.setFixedSize(280, 250)  # Slightly narrower to ensure it fits on screen
        self.setStyleSheet("""
            QDialog {
                background-color: #252535;
                color: white;
            }
            QLabel {
                color: white;
            }
            QSpinBox, QTimeEdit, QComboBox {
                background-color: #333345;
                color: white;
                border: 1px solid #444;
                padding: 5px;
            }
            QPushButton {
                background-color: #3A3A4A;
                color: white;
                border: none;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #4A4A5A;
            }
            /* Individual button styles will override these */
            QCheckBox {
                color: white;
            }
            QCheckBox::indicator {
                width: 15px;
                height: 15px;
            }
        """)
        
        layout = QFormLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Adjust label alignment to move fields to the left
        layout.setLabelAlignment(Qt.AlignLeft)
        layout.setFormAlignment(Qt.AlignLeft)
        
        # Set field growth policy to ensure fields don't expand too much
        layout.setFieldGrowthPolicy(QFormLayout.FieldsStayAtSizeHint)
        
        # Target hours setting
        self.target_hours_spin = QSpinBox()
        self.target_hours_spin.setRange(1, 24)
        self.target_hours_spin.setValue(parent.target_hours)
        self.target_hours_spin.setFixedWidth(60)  # Limit width
        layout.addRow("Target Hours:", self.target_hours_spin)
        
        # Auto-pause setting
        self.auto_pause_check = QCheckBox()
        self.auto_pause_check.setChecked(parent.auto_pause)
        layout.addRow("Auto-pause when inactive:", self.auto_pause_check)
        
        # Auto-pause minutes
        self.auto_pause_minutes_spin = QSpinBox()
        self.auto_pause_minutes_spin.setRange(1, 60)
        self.auto_pause_minutes_spin.setValue(parent.auto_pause_minutes)
        self.auto_pause_minutes_spin.setFixedWidth(60)  # Limit width
        layout.addRow("Auto-pause after (minutes):", self.auto_pause_minutes_spin)
        
        # Theme color
        self.theme_colors = {
            "Blue": "#00CCFF",
            "Purple": "#9370DB",
            "Green": "#4CAF50",
            "Orange": "#FF9800"
        }
        
        self.theme_color_combo = QComboBox()
        self.theme_color_combo.setFixedWidth(100)  # Limit width
        for name in self.theme_colors:
            self.theme_color_combo.addItem(name)
        
        # Set current color
        current_color = parent.theme_color
        for i, (name, color) in enumerate(self.theme_colors.items()):
            if color == current_color:
                self.theme_color_combo.setCurrentIndex(i)
                break
        
        layout.addRow("Theme Color:", self.theme_color_combo)
        
        # --- Button Layout Setup ---
        # Create a container widget for the buttons
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container) # Set the layout on the container
        button_layout.setContentsMargins(0, 10, 0, 0) # Add some top margin for spacing
        button_layout.setSpacing(10)
        
        # Add buttons
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setFixedSize(80, 20)  # Smaller size
        self.cancel_button.setStyleSheet("""
            QPushButton {
                background-color: #444455;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 5px 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #555566;
            }
        """)
        self.cancel_button.clicked.connect(self.reject)
        
        self.save_button = QPushButton("Save")
        self.save_button.setFixedSize(80, 20)  # Smaller size
        self.save_button.setStyleSheet("""
            QPushButton {
                background-color: #00AADD;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 5px 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #22CCFF;
            }
        """)
        self.save_button.clicked.connect(self.accept)
        
        # Add buttons with spacing between them
        button_layout.addWidget(self.cancel_button)
        button_layout.addWidget(self.save_button)
        
        # Add stretch only on the right side to push buttons to the left
        button_layout.addStretch(1)
        
        # Add the container widget to the form layout, spanning both columns
        layout.addRow(button_container)


def main():
    app = QApplication(sys.argv)
    window = FocusTimer()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

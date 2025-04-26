#!/usr/bin/env python3
# main_window.py - Main application window for Focus Timer

import time
from datetime import datetime
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QLabel, QPushButton, QMenu, QAction
)

from focus_timer_app.ui.focus_circle import FocusCircle
from focus_timer_app.ui.settings_dialog import SettingsDialog
from focus_timer_app.ui.styles import FocusTimerStyles

class FocusTimerWindow(QMainWindow):
    """Main window for the Focus Timer application"""
    
    def __init__(self, timer_core, activity_tracker, app_tracker, logger, analyzer):
        """Initialize the main window
        
        Args:
            timer_core: FocusTimerCore object
            activity_tracker: ActivityTracker object
            app_tracker: ApplicationTracker object
            logger: ActivityLogger object
            analyzer: ActivityAnalyzer object
        """
        super().__init__()
        
        # Store references to core components
        self.timer_core = timer_core
        self.activity_tracker = activity_tracker
        self.app_tracker = app_tracker
        self.logger = logger
        self.analyzer = analyzer
        
        # UI state variables
        self.ui_minimized = False  # Track if UI is in minimized state
        
        # Set up UI
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
        
        # Set initial window position to top-left
        self.move(150, 50)  # Small offset from the very corner
    
    def setup_ui(self):
        """Set up the main window UI"""
        # Main window setup
        self.setWindowTitle("Focus Timer")
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        self.setFixedSize(380, 30)  # Reduced width and height
        
        # Store the original width for restoration
        self.original_width = 380
        
        # Enable rounded corners by making window transparent
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        # Set stylesheet with theme color and rounded corners
        self.setStyleSheet(FocusTimerStyles.get_main_window_style())
        
        # Create central widget
        central_widget = QWidget(self)
        central_widget.setObjectName("centralWidget")
        self.setCentralWidget(central_widget)
        
        # Create layout
        layout = QHBoxLayout(central_widget)
        layout.setContentsMargins(10, 5, 10, 5)
        layout.setSpacing(0)  # No spacing to bring elements as close as possible
        layout.setAlignment(Qt.AlignVCenter)  # Vertically center all widgets in the layout
        
        # Add menu button first (matching original layout)
        self.menu_button = QPushButton("≡")
        self.menu_button.setFont(QFont("Arial", 14))
        self.menu_button.setFixedSize(22, 22)
        self.menu_button.setCursor(Qt.PointingHandCursor)
        self.menu_button.clicked.connect(self.show_menu)
        layout.addWidget(self.menu_button)
        
        # Add focus circle widget
        self.focus_widget = FocusCircle(
            color=FocusTimerStyles.WORK_COLOR if self.timer_core.focus_mode == self.timer_core.MODE_WORK 
                  else FocusTimerStyles.LEISURE_COLOR,
            parent=self
        )
        self.focus_widget.setFixedSize(19, 19)  # Smaller circle to match original
        layout.addWidget(self.focus_widget)
        
        # Add focus time label (MM:SS)
        self.focus_time_label = QLabel("00:00")
        self.focus_time_label.setFont(QFont("Arial", 16, QFont.Bold))  # Match original font size
        self.focus_time_label.setStyleSheet(f"color: {FocusTimerStyles.WORK_COLOR};")
        self.focus_time_label.setFixedWidth(50)
        self.focus_time_label.setAlignment(Qt.AlignCenter | Qt.AlignVCenter)  # Center both horizontally and vertically
        layout.addWidget(self.focus_time_label)
        
        # Add focus label with reduced line spacing
        self.focus_label = QLabel("<html><div style='line-height:80%'>FOCUS TIME<br>ELAPSED</div></html>")
        self.focus_label.setFont(QFont("Arial", 7))  # Smaller font
        self.focus_label.setStyleSheet("color: #888; padding-top: 2px;")  # Gray color with padding
        self.focus_label.setFixedWidth(50)
        self.focus_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)  # Left-aligned with vertical center
        layout.addWidget(self.focus_label)
        
        # Add work time label (H hr M min)
        self.work_time_label = QLabel("0 hr 0 min")
        self.work_time_label.setFont(QFont("Arial", 16, QFont.Bold))  # Match original font size
        self.work_time_label.setStyleSheet(f"color: {FocusTimerStyles.WORK_HOURS_COLOR};")
        self.work_time_label.setFixedWidth(100)  # Match original width
        self.work_time_label.setAlignment(Qt.AlignCenter | Qt.AlignVCenter)  # Center both horizontally and vertically
        layout.addWidget(self.work_time_label)
        
        # Add work label with reduced line spacing
        self.work_label = QLabel("<html><div style='line-height:80%'>WORK<br>HOURS</div></html>")
        self.work_label.setFont(QFont("Arial", 7))  # Smaller font
        self.work_label.setStyleSheet("color: #888; padding-top: 2px; padding-left: 30px;")  # Added left padding to move label right
        self.work_label.setFixedWidth(40)  # Match original width
        self.work_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)  # Left-aligned with vertical center
        layout.addWidget(self.work_label)
        
        # Add spacing before percent display to move it further right (matching original layout)
        spacer = QWidget()
        spacer.setFixedWidth(60)  # Increased width to move percent display further right
        layout.addWidget(spacer)
        
        # Add percent display
        self.percent_display = QLabel("0%")
        self.percent_display.setFont(QFont("Arial", 16, QFont.Bold))  # Match original font size
        self.percent_display.setStyleSheet("color: #888;")  # Start with gray, will be updated dynamically
        self.percent_display.setFixedWidth(30)
        self.percent_display.setAlignment(Qt.AlignCenter | Qt.AlignVCenter)  # Center both horizontally and vertically
        layout.addWidget(self.percent_display)
        
        # Add percent label with reduced line spacing
        self.percent_label = QLabel("<html><div style='line-height:80%'>PERCENT<br>OF DAY</div></html>")
        self.percent_label.setFont(QFont("Arial", 7))  # Smaller font
        self.percent_label.setStyleSheet("color: #888; padding-top: 2px;")  # Gray color with padding
        self.percent_label.setFixedWidth(60)  # Match original width
        self.percent_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)  # Left-aligned with vertical center
        layout.addWidget(self.percent_label)
        
        # Menu button is now added at the beginning of the layout
        
        # Update the focus circle active state
        self.focus_widget.set_active(self.timer_core.focus_start_time is not None and not self.timer_core.is_paused)
    
    def update_display(self):
        """Update the UI display"""
        # Update focus time if active
        if self.timer_core.focus_start_time and not self.timer_core.is_paused:
            # Update the focus time label
            self.focus_time_label.setText(self.timer_core.format_elapsed_time())
            
            # Log activity once per minute
            current_minute = int(time.time() / 60)
            if not hasattr(self, 'last_logged_minute') or current_minute > self.last_logged_minute:
                self.last_logged_minute = current_minute
                # Use the 5-second window to determine if active
                is_active = self.activity_tracker.is_active_in_window()
                self.log_activity(is_active)  # Log based on activity window
            
            # Auto-pause after inactivity
            if self.timer_core.settings.auto_pause and self.activity_tracker.time_since_last_activity() > self.timer_core.settings.auto_pause_minutes * 60:
                self.pause_focus()
        elif self.timer_core.is_paused or not self.timer_core.focus_start_time:
            # Log inactivity once per minute when paused
            current_minute = int(time.time() / 60)
            if not hasattr(self, 'last_logged_minute') or current_minute > self.last_logged_minute:
                self.last_logged_minute = current_minute
                # Still check activity window even when paused
                is_active = self.activity_tracker.is_active_in_window()
                self.log_activity(is_active)
        
        # Update work/leisure time display
        self.update_time_display()
        
        # Update the focus circle active state
        self.focus_widget.set_active(self.timer_core.focus_start_time is not None and not self.timer_core.is_paused)
    
    def update_time_display(self):
        """Update the work/leisure time display"""
        # Update the appropriate time display based on focus mode
        if self.timer_core.focus_mode == self.timer_core.MODE_WORK:
            self.work_time_label.setText(self.timer_core.format_work_time())
        else:
            self.work_time_label.setText(self.timer_core.format_leisure_time())
        
        # Calculate percent of day (allow going over 100%)
        percent = self.timer_core.get_progress_percent()
        display_percent = min(100, percent)  # Cap at 100% for display purposes
        
        # Update the percent display
        if percent > 100:
            self.percent_display.setText(f"{percent:.0f}%")
        else:
            self.percent_display.setText(f"{display_percent:.0f}%")
        
        # Set the color based on progress
        color_hex = FocusTimerStyles.get_progress_color(percent)
        self.percent_display.setStyleSheet(f"color: {color_hex}; font-weight: bold;")
    
    def update_percent_color(self, base_color):
        """Update the percent indicator color based on the current mode
        
        Args:
            base_color (str): Base color in hex format
        """
        # Get the current style
        current_style = self.percent_display.styleSheet()
        
        # If the percent indicator has a dynamic color based on progress,
        # we don't want to override it completely, just tint it
        if "font-weight: bold" in current_style:
            # Keep the font weight but change the color
            self.percent_display.setStyleSheet(f"color: {base_color}; font-weight: bold;")
        else:
            # Just change the color
            self.percent_display.setStyleSheet(f"color: {base_color};")
    
    def start_focus(self):
        """Start the focus timer"""
        # Start the timer
        self.timer_core.start_focus()
        
        # Update the focus circle
        self.focus_widget.set_active(True)
        
        # Log the start
        is_active = self.activity_tracker.is_active_in_window()
        self.log_activity(is_active, resume=True)
    
    def pause_focus(self):
        """Pause the focus timer"""
        # Pause the timer
        self.timer_core.pause_focus()
        
        # Update the focus circle
        self.focus_widget.set_active(False)
        
        # Log the pause
        is_active = self.activity_tracker.is_active_in_window()
        self.log_activity(is_active, pause=True)
    
    def reset_focus(self):
        """Reset the focus timer"""
        # Reset the timer
        self.timer_core.reset_focus()
        
        # Update the focus circle
        self.focus_widget.set_active(False)
        
        # Log the reset
        is_active = self.activity_tracker.is_active_in_window()
        self.log_activity(is_active, reset=True)
    
    def toggle_focus_mode(self):
        """Toggle between work and leisure focus modes"""
        # Toggle the focus mode
        new_mode = self.timer_core.toggle_focus_mode()
        
        # Update UI based on new mode
        if new_mode == self.timer_core.MODE_LEISURE:
            # Change to leisure color and update label
            self.focus_widget.set_color(FocusTimerStyles.LEISURE_COLOR)
            self.work_label.setText("<html><div style='line-height:80%'>LEISURE<br>HOURS</div></html>")
            # Change focus time elapsed counter to leisure color (yellow)
            self.focus_time_label.setStyleSheet(f"color: {FocusTimerStyles.LEISURE_COLOR};")
            # Change leisure hours counter to green
            self.work_time_label.setStyleSheet(f"color: {FocusTimerStyles.LEISURE_HOURS_COLOR}")
            # Update percent indicator color
            self.update_percent_color(FocusTimerStyles.LEISURE_COLOR)
        else:
            # Change back to work color and update label
            self.focus_widget.set_color(FocusTimerStyles.WORK_COLOR)
            self.work_label.setText("<html><div style='line-height:80%'>WORK<br>HOURS</div></html>")
            # Change focus time elapsed counter back to work color
            self.focus_time_label.setStyleSheet(f"color: {FocusTimerStyles.WORK_COLOR};")
            # Change work hours counter to purple
            self.work_time_label.setStyleSheet(f"color: {FocusTimerStyles.WORK_HOURS_COLOR};")
            # Update percent indicator color
            self.update_percent_color(FocusTimerStyles.WORK_COLOR)
        
        # Update the UI
        self.focus_widget.update()
        self.update_time_display()
    
    def minimize_ui(self):
        """Minimize the window to just show the focus circle"""
        if not self.ui_minimized:
            self.ui_minimized = True
            
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
            self.setFixedSize(self.original_width, 30)
            
            # Restore original layout margins
            self.centralWidget().layout().setContentsMargins(10, 5, 10, 5)
            
            # Show all the labels and numbers
            self.focus_time_label.show()
            self.focus_label.show()
            self.work_time_label.show()
            self.work_label.show()
            self.percent_display.show()
            self.percent_label.show()
    
    def show_menu(self):
        """Show the context menu"""
        menu = QMenu(self)
        menu.setStyleSheet(FocusTimerStyles.get_menu_style())
        
        # Focus actions
        if self.timer_core.focus_start_time and not self.timer_core.is_paused:
            # Currently active - show pause option
            pause_action = QAction("Pause Focus", self)
            pause_action.triggered.connect(self.pause_focus)
            menu.addAction(pause_action)
        elif self.timer_core.is_paused:
            # Currently paused - show resume option
            resume_action = QAction("Resume Focus", self)
            resume_action.triggered.connect(self.start_focus)
            menu.addAction(resume_action)
        else:
            # Not started - show start option
            start_action = QAction("Start Focus", self)
            start_action.triggered.connect(self.start_focus)
            menu.addAction(start_action)
        
        # Always show reset option
        reset_action = QAction("Reset Timer", self)
        reset_action.triggered.connect(self.reset_focus)
        menu.addAction(reset_action)
        
        # Add separator
        menu.addSeparator()
        
        # Toggle focus mode
        if self.timer_core.focus_mode == self.timer_core.MODE_WORK:
            toggle_action = QAction("Switch to Leisure Focus", self)
        else:
            toggle_action = QAction("Switch to Work Focus", self)
        toggle_action.triggered.connect(self.toggle_focus_mode)
        menu.addAction(toggle_action)
        
        # Add separator
        menu.addSeparator()
        
        # UI actions
        if self.ui_minimized:
            ui_action = QAction("Expand UI", self)
            ui_action.triggered.connect(self.restore_ui)
        else:
            ui_action = QAction("Minimize UI", self)
            ui_action.triggered.connect(self.minimize_ui)
        menu.addAction(ui_action)
        
        # Settings action
        settings_action = QAction("Settings", self)
        settings_action.triggered.connect(self.show_settings)
        menu.addAction(settings_action)
        
        # Quit action
        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self.close)
        menu.addAction(quit_action)
        
        # Show the menu
        menu.exec_(self.mapToGlobal(self.menu_button.pos()))
    
    def show_settings(self):
        """Show the settings dialog"""
        dialog = SettingsDialog(self.timer_core.settings, self)
        if dialog.exec_():
            # Get the updated settings
            updated_settings = dialog.get_settings()
            
            # Update the settings
            self.timer_core.settings.target_hours = updated_settings["target_hours"]
            self.timer_core.settings.auto_pause = updated_settings["auto_pause"]
            self.timer_core.settings.auto_pause_minutes = updated_settings["auto_pause_minutes"]
            self.timer_core.settings.theme_color = updated_settings["theme_color"]
            
            # Save the settings
            self.timer_core.settings.save()
            
            # Update the UI
            self.update_display()
    
    def log_activity(self, active, final=False, reset=False, pause=False, resume=False, restart=False):
        """Log the current activity state
        
        Args:
            active (bool): Whether the user is active
            final (bool, optional): Whether this is the final log entry
            reset (bool, optional): Whether the timer was reset
            pause (bool, optional): Whether the timer was paused
            resume (bool, optional): Whether the timer was resumed
            restart (bool, optional): Whether the timer was restarted
        """
        # Get active application and URL
        active_app = self.app_tracker.get_active_application()
        url = "n/a"
        
        # Store the application name in the activity tracker
        self.activity_tracker.record_activity(active_app)
        
        # If active app is a browser, get the URL
        if active_app in self.app_tracker.supported_browsers:
            url = self.app_tracker.get_browser_url(active_app)
        
        # Add a note for special events
        note = ""
        if final:
            note = "final"
        elif reset:
            note = "reset"
        elif pause:
            note = "pause"
        elif resume:
            note = "resume"
        elif restart:
            note = "restart"
        
        # Determine the active value: 0 = inactive, 1 = work focus, 2 = leisure focus
        active_val = 2 if (self.timer_core.focus_mode == self.timer_core.MODE_LEISURE and active) else (1 if active else 0)
        
        # Calculate work_elapsed and leisure_elapsed based on active state
        elapsed_time = self.timer_core.get_elapsed_time()
        work_elapsed = round(elapsed_time) if active_val == 1 else 0
        leisure_elapsed = round(elapsed_time) if active_val == 2 else 0
        
        # Log the activity
        self.logger.log_activity(
            active_val=active_val,
            elapsed_time=elapsed_time,
            work_elapsed=work_elapsed,
            leisure_elapsed=leisure_elapsed,
            total_work_hours=self.timer_core.work_hours,
            total_leisure_hours=self.timer_core.leisure_hours,
            active_app=active_app,
            url=url,
            note=note
        )
    
    def closeEvent(self, event):
        """Handle window close event
        
        Args:
            event: Close event
        """
        # Save settings on close
        self.timer_core.settings.save()
        
        # Stop focus if active
        if self.timer_core.focus_start_time and not self.timer_core.is_paused:
            self.pause_focus()
        
        # Log final activity state
        is_active = self.activity_tracker.is_active_in_window()
        self.log_activity(is_active, final=True)
        
        # Print confirmation
        print(f"Closing application. Final work hours: {self.timer_core.work_hours:.2f}, leisure hours: {self.timer_core.leisure_hours:.2f}")
        
        # Stop activity tracking
        self.activity_tracker.stop()
            
        event.accept()

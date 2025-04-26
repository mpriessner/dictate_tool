#!/usr/bin/env python3
# activity.py - Activity tracking for Focus Timer

import time
from pynput import mouse, keyboard

class ActivityTracker:
    """Tracks user activity via mouse and keyboard events"""
    
    def __init__(self, activity_window=5):
        """Initialize activity tracker
        
        Args:
            activity_window (int): Time window in seconds to check for activity
        """
        self.activity_window = activity_window  # seconds to check for activity
        self.activity_events = []
        self.last_activity = time.time()
        self.last_active_app_name = "Unknown"
        
        # Set up listeners
        self.mouse_listener = None
        self.keyboard_listener = None
    
    def start(self):
        """Start tracking activity"""
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
    
    def stop(self):
        """Stop tracking activity"""
        if self.mouse_listener:
            self.mouse_listener.stop()
        
        if self.keyboard_listener:
            self.keyboard_listener.stop()
    
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
    
    def record_activity(self, app_name=None):
        """Record an activity event with timestamp
        
        Args:
            app_name (str, optional): Name of the active application
        """
        current_time = time.time()
        self.last_activity = current_time
        self.activity_events.append(current_time)
        
        # Store the application active at this moment if provided
        if app_name:
            self.last_active_app_name = app_name
        
        # Clean up old events outside the activity window
        cutoff_time = current_time - self.activity_window
        self.activity_events = [t for t in self.activity_events if t >= cutoff_time]
    
    def is_active_in_window(self):
        """Check if there was activity in the activity window
        
        Returns:
            bool: True if there was activity in the window, False otherwise
        """
        if not self.activity_events:
            return False
            
        current_time = time.time()
        cutoff_time = current_time - self.activity_window
        
        # Check if there are any events in the window
        return any(t >= cutoff_time for t in self.activity_events)
    
    def time_since_last_activity(self):
        """Get time in seconds since last activity
        
        Returns:
            float: Time in seconds since last activity
        """
        return time.time() - self.last_activity

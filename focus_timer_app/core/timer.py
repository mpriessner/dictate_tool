#!/usr/bin/env python3
# timer.py - Core timer functionality for Focus Timer

import time
from datetime import datetime

class FocusTimerCore:
    """Core timer functionality for tracking focus time"""
    
    # Focus mode constants
    MODE_WORK = "work"
    MODE_LEISURE = "leisure"
    
    def __init__(self, settings):
        """Initialize the focus timer core
        
        Args:
            settings: Settings object containing timer settings
        """
        self.settings = settings
        
        # Timer variables
        self.focus_start_time = None
        self.is_paused = False
        self.elapsed_time_before_pause = 0  # Track elapsed time for resume functionality
        self.focus_mode = self.MODE_WORK  # Default to work mode
        
        # Time tracking
        self.work_hours = settings.work_hours
        self.leisure_hours = settings.leisure_hours
    
    def start_focus(self):
        """Start the focus timer"""
        if self.is_paused:
            # If resuming from pause, keep the elapsed time
            self.is_paused = False
        else:
            # If starting fresh, reset elapsed time
            self.elapsed_time_before_pause = 0
        
        self.focus_start_time = time.time()
        return True
    
    def pause_focus(self):
        """Pause the focus timer"""
        if self.focus_start_time and not self.is_paused:
            # Calculate elapsed time and add to total
            elapsed = time.time() - self.focus_start_time
            self.elapsed_time_before_pause += elapsed
            
            # Update work/leisure hours
            self.add_elapsed_time(elapsed / 3600)  # Convert seconds to hours
            
            # Set paused state
            self.is_paused = True
            return True
        return False
    
    def reset_focus(self):
        """Reset the focus timer"""
        self.focus_start_time = None
        self.is_paused = False
        self.elapsed_time_before_pause = 0
        return True
    
    def toggle_focus_mode(self):
        """Toggle between work and leisure focus modes"""
        # If timer is running, pause it to record current elapsed time
        was_running = False
        if self.focus_start_time and not self.is_paused:
            was_running = True
            self.pause_focus()
        
        # Toggle the focus mode
        if self.focus_mode == self.MODE_WORK:
            self.focus_mode = self.MODE_LEISURE
        else:
            self.focus_mode = self.MODE_WORK
        
        # If timer was running, restart it
        if was_running:
            self.start_focus()
        
        return self.focus_mode
    
    def get_elapsed_time(self):
        """Get the current elapsed time in seconds
        
        Returns:
            float: Elapsed time in seconds
        """
        if self.focus_start_time and not self.is_paused:
            return time.time() - self.focus_start_time + self.elapsed_time_before_pause
        return self.elapsed_time_before_pause
    
    def add_elapsed_time(self, hours):
        """Add elapsed time to the appropriate bucket (work or leisure)
        
        Args:
            hours (float): Hours to add
        """
        if self.focus_mode == self.MODE_WORK:
            self.work_hours += hours
        else:  # Leisure mode
            self.leisure_hours += hours
        
        # Update settings
        self.settings.work_hours = self.work_hours
        self.settings.leisure_hours = self.leisure_hours
        self.settings.save()
    
    def get_progress_percent(self):
        """Calculate progress as a percentage of target hours
        
        Returns:
            float: Progress percentage
        """
        return (self.work_hours / self.settings.target_hours) * 100 if self.settings.target_hours > 0 else 0
    
    def format_elapsed_time(self):
        """Format the current elapsed time as minutes:seconds
        
        Returns:
            str: Formatted time string (MM:SS)
        """
        elapsed = self.get_elapsed_time()
        minutes, seconds = divmod(int(elapsed), 60)
        return f"{minutes:02d}:{seconds:02d}"
    
    def format_work_time(self):
        """Format the total work time as hours and minutes
        
        Returns:
            str: Formatted time string (H hr M min)
        """
        hours = int(self.work_hours)
        minutes = int((self.work_hours - hours) * 60)
        return f"{hours} hr {minutes} min"
    
    def format_leisure_time(self):
        """Format the total leisure time as hours and minutes
        
        Returns:
            str: Formatted time string (H hr M min)
        """
        hours = int(self.leisure_hours)
        minutes = int((self.leisure_hours - hours) * 60)
        return f"{hours} hr {minutes} min"

#!/usr/bin/env python3
# settings.py - Settings management for Focus Timer

from PyQt5.QtCore import QSettings
from datetime import datetime

class Settings:
    """Manages application settings for Focus Timer"""
    
    def __init__(self):
        self.settings = QSettings("FocusTimer", "Settings")
        self.target_hours = 8
        self.auto_pause = False
        self.auto_pause_minutes = 5
        self.theme_color = "#00CCFF"
        self.work_hours = 0
        self.leisure_hours = 0
        self.load()
    
    def load(self):
        """Load settings with defaults"""
        self.target_hours = self.settings.value("target_hours", 8, type=int)
        self.auto_pause = self.settings.value("auto_pause", False, type=bool)
        self.auto_pause_minutes = self.settings.value("auto_pause_minutes", 5, type=int)
        self.theme_color = self.settings.value("theme_color", "#00CCFF", type=str)
        # Load work and leisure hours if available
        self.work_hours = self.settings.value("today_work_hours", 0, type=float)
        self.leisure_hours = self.settings.value("today_leisure_hours", 0, type=float)
    
    def save(self):
        """Save current settings"""
        self.settings.setValue("target_hours", self.target_hours)
        self.settings.setValue("auto_pause", self.auto_pause)
        self.settings.setValue("auto_pause_minutes", self.auto_pause_minutes)
        self.settings.setValue("theme_color", self.theme_color)
        self.settings.setValue("today_work_hours", self.work_hours)
        self.settings.setValue("today_leisure_hours", self.leisure_hours)
        self.settings.setValue("last_date", datetime.now().strftime("%Y-%m-%d"))
    
    def check_new_day(self):
        """Check if it's a new day and reset work hours if needed
        
        Returns:
            bool: True if it's a new day, False otherwise
        """
        today = datetime.now().strftime("%Y-%m-%d")
        last_date = self.settings.value("last_date", "", type=str)
        
        if last_date != today:
            print(f"New day detected: {last_date} -> {today}")
            self.work_hours = 0
            self.leisure_hours = 0
            self.settings.setValue("today_work_hours", 0)
            self.settings.setValue("today_leisure_hours", 0)
            self.settings.setValue("last_date", today)
            return True
        
        return False

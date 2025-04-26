#!/usr/bin/env python3
# logger.py - Logging functionality for Focus Timer

import os
import csv
from datetime import datetime
from pathlib import Path

class ActivityLogger:
    """Logs activity data to CSV files"""
    
    def __init__(self, base_dir=None):
        """Initialize the activity logger
        
        Args:
            base_dir (str, optional): Base directory for logs. If None, uses script directory.
        """
        # Set up logging directory
        if base_dir is None:
            script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            self.logs_dir = os.path.join(script_dir, "focus_logs")
        else:
            self.logs_dir = os.path.join(base_dir, "focus_logs")
        
        os.makedirs(self.logs_dir, exist_ok=True)
        
        # Set up today's directory and log file
        self.setup_logging()
    
    def setup_logging(self):
        """Set up the logging directory structure for today"""
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
                writer.writerow(["timestamp", "active", "elapsed_time", "work_elapsed", "leisure_elapsed", 
                               "total_work_hours", "total_leisure_hours", "active_app", "url", "note"])
    
    def log_activity(self, active_val, elapsed_time, work_elapsed, leisure_elapsed, 
                    total_work_hours, total_leisure_hours, active_app, url, note=""):
        """Log activity to the CSV file
        
        Args:
            active_val (int): 0=inactive, 1=work focus, 2=leisure focus
            elapsed_time (float): Current session elapsed time in seconds
            work_elapsed (float): Work elapsed time in seconds for this entry
            leisure_elapsed (float): Leisure elapsed time in seconds for this entry
            total_work_hours (float): Total accumulated work hours
            total_leisure_hours (float): Total accumulated leisure hours
            active_app (str): Name of the active application
            url (str): URL if browser is active
            note (str, optional): Additional note (e.g., "pause", "resume", "final")
        """
        # Check if we need to create a new log for a new day
        current_date = datetime.now().strftime("%Y-%m-%d")
        log_date = os.path.basename(os.path.dirname(self.log_file))
        
        # If the date has changed, set up new logging directory and file
        if current_date != log_date:
            print(f"Date changed from {log_date} to {current_date}, creating new log file")
            self.setup_logging()
        
        # Get current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Write to log file
        with open(self.log_file, 'a', newline='') as f:
            writer = csv.writer(f)
            row = [
                timestamp, 
                active_val, 
                round(elapsed_time), 
                round(work_elapsed), 
                round(leisure_elapsed),
                round(total_work_hours, 2), 
                round(total_leisure_hours, 2), 
                active_app, 
                url
            ]
            if note:
                row.append(note)
            writer.writerow(row)
    
    def get_log_file_path(self):
        """Get the path to the current log file
        
        Returns:
            str: Path to the current log file
        """
        return self.log_file

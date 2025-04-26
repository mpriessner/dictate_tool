#!/usr/bin/env python3
# helpers.py - Helper functions for Focus Timer

import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

def format_time_duration(seconds):
    """Format a time duration in seconds to a human-readable string
    
    Args:
        seconds (float): Time in seconds
        
    Returns:
        str: Formatted time string
    """
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    
    if hours > 0:
        return f"{int(hours)}h {int(minutes)}m"
    else:
        return f"{int(minutes)}m {int(seconds)}s"

def get_date_range(days=7):
    """Get a range of dates for the past N days
    
    Args:
        days (int): Number of days to include
        
    Returns:
        list: List of date strings in YYYY-MM-DD format
    """
    today = datetime.now().date()
    dates = []
    
    for i in range(days):
        date = today - timedelta(days=i)
        dates.append(date.strftime("%Y-%m-%d"))
    
    return dates

def ensure_path_exists(path):
    """Ensure a directory path exists
    
    Args:
        path (str): Path to ensure exists
        
    Returns:
        bool: True if the path exists or was created, False otherwise
    """
    try:
        os.makedirs(path, exist_ok=True)
        return True
    except Exception as e:
        print(f"Error creating directory {path}: {e}")
        return False

def get_app_directory():
    """Get the application directory
    
    Returns:
        str: Path to the application directory
    """
    # If running as a frozen executable
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    # If running as a script
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def get_logs_directory():
    """Get the logs directory
    
    Returns:
        str: Path to the logs directory
    """
    app_dir = get_app_directory()
    logs_dir = os.path.join(app_dir, "focus_logs")
    ensure_path_exists(logs_dir)
    return logs_dir

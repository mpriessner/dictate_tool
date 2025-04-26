#!/usr/bin/env python3
# analyzer.py - Time analysis functions for Focus Timer

import os
import csv
from datetime import datetime

class ActivityAnalyzer:
    """Analyzes activity logs to calculate focus times"""
    
    def __init__(self, logger):
        """Initialize the activity analyzer
        
        Args:
            logger: ActivityLogger object
        """
        self.logger = logger
    
    def calculate_hours_from_log(self):
        """Calculate total work and leisure hours from today's log file
        
        Returns:
            tuple: (work_hours, leisure_hours)
        """
        log_file = self.logger.get_log_file_path()
        
        if os.path.exists(log_file):
            try:
                # Read the log file
                with open(log_file, 'r', newline='') as f:
                    reader = csv.reader(f)
                    # Skip header
                    next(reader)
                    
                    # Load all entries
                    entries = list(reader)
                
                # Check if there's a reset or restart entry
                reset_index = -1
                for i, row in enumerate(entries):
                    # Check if this is a reset or restart entry
                    if (len(row) >= 10 and (row[9] == "reset" or row[9] == "restart")) or \
                       (len(row) >= 9 and (row[8] == "reset" or row[8] == "restart")):
                        reset_index = i
                
                # If we found a reset/restart, only process entries after the reset
                if reset_index >= 0:
                    entries = entries[reset_index+1:]
                
                # Initialize tracking variables
                work_periods = []
                leisure_periods = []
                last_active_time_work = None
                last_active_time_leisure = None
                last_active_state_work = False
                last_active_state_leisure = False
                
                for row in entries:
                    if len(row) >= 5:
                        # Get timestamp and active state
                        timestamp_str = row[0]
                        active_val = int(row[1])
                        is_active_work = active_val == 1
                        is_active_leisure = active_val == 2
                        
                        try:
                            # Parse timestamp
                            timestamp = datetime.strptime(timestamp_str, "%Y-%m-%d %H:%M:%S")
                            
                            # Process work periods
                            # If we have a transition from active work to inactive/leisure, record the period
                            if last_active_state_work and not is_active_work and last_active_time_work is not None:
                                active_duration = (timestamp - last_active_time_work).total_seconds()
                                # Only count if the difference is reasonable (less than 5 minutes)
                                if active_duration <= 300:
                                    work_periods.append(active_duration)
                            
                            # If we have a transition to active work, record the start time
                            if is_active_work and not last_active_state_work:
                                last_active_time_work = timestamp
                            
                            # Process leisure periods
                            # If we have a transition from active leisure to inactive/work, record the period
                            if last_active_state_leisure and not is_active_leisure and last_active_time_leisure is not None:
                                active_duration = (timestamp - last_active_time_leisure).total_seconds()
                                # Only count if the difference is reasonable (less than 5 minutes)
                                if active_duration <= 300:
                                    leisure_periods.append(active_duration)
                            
                            # If we have a transition to active leisure, record the start time
                            if is_active_leisure and not last_active_state_leisure:
                                last_active_time_leisure = timestamp
                            
                            # Update last active states
                            last_active_state_work = is_active_work
                            last_active_state_leisure = is_active_leisure
                            
                        except ValueError:
                            # Skip entries with invalid timestamps
                            pass
                
                # If we're still active at the end, add the final periods
                now = datetime.now()
                
                if last_active_state_work and last_active_time_work is not None:
                    active_duration = (now - last_active_time_work).total_seconds()
                    # Only count if the difference is reasonable (less than 5 minutes)
                    if active_duration <= 300:
                        work_periods.append(active_duration)
                
                if last_active_state_leisure and last_active_time_leisure is not None:
                    active_duration = (now - last_active_time_leisure).total_seconds()
                    # Only count if the difference is reasonable (less than 5 minutes)
                    if active_duration <= 300:
                        leisure_periods.append(active_duration)
                
                # Sum up all active periods
                total_work_seconds = sum(work_periods)
                total_leisure_seconds = sum(leisure_periods)
                total_work_hours = total_work_seconds / 3600
                total_leisure_hours = total_leisure_seconds / 3600
                
                print(f"Calculated from log: {total_work_hours:.2f} work hours, {total_leisure_hours:.2f} leisure hours")
                
                return total_work_hours, total_leisure_hours
            
            except Exception as e:
                print(f"Error calculating hours from log: {e}")
        
        return 0, 0
    
    def get_daily_summary(self, date=None):
        """Get a summary of activity for a specific day
        
        Args:
            date (str, optional): Date in YYYY-MM-DD format. If None, uses today.
            
        Returns:
            dict: Summary of activity
        """
        if date is None:
            date = datetime.now().strftime("%Y-%m-%d")
        
        # Get the log file for the specified date
        log_dir = os.path.dirname(os.path.dirname(self.logger.get_log_file_path()))
        log_file = os.path.join(log_dir, date, "activity_log.csv")
        
        summary = {
            "date": date,
            "work_hours": 0,
            "leisure_hours": 0,
            "active_apps": {},
            "active_urls": {}
        }
        
        if os.path.exists(log_file):
            try:
                # Read the log file
                with open(log_file, 'r', newline='') as f:
                    reader = csv.reader(f)
                    # Skip header
                    next(reader)
                    
                    # Process entries
                    for row in entries:
                        if len(row) >= 8:
                            # Get active app and URL
                            active_app = row[7]
                            url = row[8] if len(row) >= 9 else "n/a"
                            
                            # Count occurrences of apps and URLs
                            if active_app in summary["active_apps"]:
                                summary["active_apps"][active_app] += 1
                            else:
                                summary["active_apps"][active_app] = 1
                            
                            if url != "n/a":
                                if url in summary["active_urls"]:
                                    summary["active_urls"][url] += 1
                                else:
                                    summary["active_urls"][url] = 1
                
                # Calculate work and leisure hours
                work_hours, leisure_hours = self.calculate_hours_from_log()
                summary["work_hours"] = work_hours
                summary["leisure_hours"] = leisure_hours
                
            except Exception as e:
                print(f"Error generating daily summary: {e}")
        
        return summary

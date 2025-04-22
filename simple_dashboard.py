#!/usr/bin/env python3
"""
simple_dashboard.py - A simplified version of the focus dashboard
that only handles loading CSV files and sending data to the React app.
"""

import os
import csv
import sys
from datetime import datetime, timedelta
from pathlib import Path

from PyQt5.QtCore import QObject, pyqtSlot, QUrl, QVariant
from PyQt5.QtWidgets import QApplication, QDialog, QVBoxLayout
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtWebChannel import QWebChannel

class SimpleDashboardHandler(QObject):
    """Handles data loading and communication with the React app"""
    
    def __init__(self, log_dir):
        super().__init__()
        self.log_dir = Path(log_dir)
        print(f"Initialized with log directory: {self.log_dir}")
    
    def get_log_path(self, date_str):
        """Get the path to the activity log file for a specific date"""
        log_path = self.log_dir / date_str / "activity_log.csv"
        print(f"Log path: {log_path}")
        return log_path
    
    def get_day_data(self, date_str):
        """Get focus and break data for a specific day"""
        log_path = self.get_log_path(date_str)
        print(f"[DEBUG] Reading day data from: {log_path}")
        
        data = {
            'focus': 0,
            'break': 0,
            'raw_app_times': {},
            'app_usage': {},
            'intervals': []
        }
        
        if log_path.exists():
            print(f"Log file exists: {log_path}")
            try:
                with log_path.open('r') as f:
                    print(f"Opened file: {log_path}")
                    reader = csv.DictReader(f)
                    print(f"CSV headers: {reader.fieldnames}")
                    
                    for row in reader:
                        # Print a sample row to debug
                        if data['focus'] == 0 and data['break'] == 0:
                            print(f"Sample row: {row}")
                        
                        # Parse the activity log CSV format
                        try:
                            elapsed_time = float(row['elapsed_time'])
                            # Active is '1' for active, '0' for inactive
                            is_active = row['active'] == '1'
                            app_name = row['active_app']
                            timestamp = row['timestamp']
                            
                            # Track time per application (only for active time)
                            if is_active:
                                data['raw_app_times'][app_name] = data['raw_app_times'].get(app_name, 0) + elapsed_time
                        except Exception as e:
                            print(f"Error processing row: {e}")
                            print(f"Problematic row: {row}")
            except Exception as e:
                print(f"Error reading file: {e}")
        else:
            print(f"Log file does not exist: {log_path}")
        
        # Focus and break time will be calculated from the intervals
        
        # Calculate app usage percentages
        total_app_time = sum(data['raw_app_times'].values())
        if total_app_time > 0:
            print(f"[DEBUG] Raw app times: {data['raw_app_times']}")
            
            # Calculate percentages based on total active app time
            data['app_usage'] = {}
            for app, time in data['raw_app_times'].items():
                percentage = round((time / total_app_time) * 100, 1)
                data['app_usage'][app] = percentage
                print(f"[DEBUG] App {app}: {time} seconds = {percentage}% of total active time")
        else:
            print("[DEBUG] No active app time recorded, can't calculate app usage percentages")
        
        # Generate intervals for the timeline
        try:
            # Convert the date string to a datetime object
            base_date = datetime.strptime(date_str, '%Y-%m-%d')
            
            # Process the raw data to create intervals
            if log_path.exists():
                # First pass: collect all timestamps, active states, and app names
                all_logs = []
                with log_path.open('r') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        timestamp = datetime.strptime(row['timestamp'], '%Y-%m-%d %H:%M:%S')
                        is_active = row['active'] == '1'
                        app_name = row['active_app']
                        all_logs.append((timestamp, is_active, app_name))
                
                # Sort logs by timestamp
                all_logs.sort(key=lambda x: x[0])
                
                # Second pass: create intervals with the 5 consecutive inactive logs rule
                intervals = []
                inactive_count = 0
                current_focus = None
                current_start = None
                last_timestamp = None
                
                # Reset focus and break time counters
                data['focus'] = 0
                data['break'] = 0
                data['raw_app_times'] = {}
                
                for i, (timestamp, is_active, app_name) in enumerate(all_logs):
                    # Track consecutive inactive logs
                    if not is_active:
                        inactive_count += 1
                    else:
                        inactive_count = 0
                    
                    # Determine if this should be considered a focus or break period
                    # Only count as break if we have 5+ consecutive inactive logs
                    is_break = inactive_count >= 5
                    is_focus = is_active or (not is_active and inactive_count < 5)
                    
                    # For interval creation, we consider it a focus period unless it's a confirmed break
                    current_state = not is_break  # True for focus, False for break
                    
                    # If this is the first log or there's a change in state
                    if current_focus is None or current_focus != current_state:
                        # If we have a previous interval, add it
                        if current_start is not None and last_timestamp is not None:
                            # Calculate the duration of this interval in seconds
                            interval_duration = (last_timestamp - current_start).total_seconds()
                            
                            # Add to the appropriate counter
                            if current_focus:  # If it was a focus interval
                                data['focus'] += interval_duration
                            else:  # If it was a break interval
                                data['break'] += interval_duration
                                
                            intervals.append({
                                'start': current_start.isoformat(),
                                'end': last_timestamp.isoformat(),
                                'focus': current_focus
                            })
                        
                        # Start a new interval
                        current_focus = current_state
                        current_start = timestamp
                    
                    # Track app usage (only for active time)
                    if is_active:
                        # Calculate time since last log if not the first log
                        if i > 0 and all_logs[i-1][0] < timestamp:
                            time_diff = (timestamp - all_logs[i-1][0]).total_seconds()
                            # Only count if reasonable (less than 10 minutes)
                            if time_diff < 600:
                                data['raw_app_times'][app_name] = data['raw_app_times'].get(app_name, 0) + time_diff
                    
                    last_timestamp = timestamp
                
                # Add the final interval if there is one
                if current_start is not None and last_timestamp is not None:
                    # Calculate the duration of the final interval in seconds
                    interval_duration = (last_timestamp - current_start).total_seconds()
                    
                    # Add to the appropriate counter
                    if current_focus:  # If it was a focus interval
                        data['focus'] += interval_duration
                    else:  # If it was a break interval
                        data['break'] += interval_duration
                        
                    intervals.append({
                        'start': current_start.isoformat(),
                        'end': last_timestamp.isoformat(),
                        'focus': current_focus
                    })
                
                # Convert seconds to minutes
                data['focus'] = round(data['focus'] / 60)
                data['break'] = round(data['break'] / 60)
                
                # Cap total time at 24 hours (1440 minutes) if it exceeds that
                total_minutes = data['focus'] + data['break']
                if total_minutes > 1440:
                    scaling_factor = 1440 / total_minutes
                    data['focus'] = round(data['focus'] * scaling_factor)
                    data['break'] = round(data['break'] * scaling_factor)
                    print(f"[DEBUG] Scaled down times to fit within 24 hours. New focus: {data['focus']}m, break: {data['break']}m")
                
                data['intervals'] = intervals
                print(f"[DEBUG] Generated {len(intervals)} intervals for timeline")
        except Exception as e:
            print(f"[DEBUG] Error generating intervals: {e}")
            data['intervals'] = []
        
        print(f"[DEBUG] Day data for {date_str}: {data}")
        return data
    
    def get_month_data(self, date_str):
        """Get aggregated data for the month containing the specified date"""
        date = datetime.strptime(date_str, '%Y-%m-%d')
        
        # Get the first day of the month
        start_of_month = date.replace(day=1)
        
        # Get the number of days in the month
        if start_of_month.month == 12:
            next_month = start_of_month.replace(year=start_of_month.year + 1, month=1)
        else:
            next_month = start_of_month.replace(month=start_of_month.month + 1)
        days_in_month = (next_month - start_of_month).days
        
        month_data = {
            'focus': 0,
            'break': 0,
            'raw_app_times': {},
            'app_usage': {},
            'weekly_hours': []
        }
        
        # Find the Monday before or on the first day of the month
        first_monday = start_of_month - timedelta(days=start_of_month.weekday())
        
        # Process each week that overlaps with the month
        week_number = 1
        current_date = first_monday
        
        while current_date < next_month:
            week_start = current_date
            week_end = current_date + timedelta(days=6)
            
            # Skip weeks that are entirely before the month starts
            if week_end < start_of_month:
                current_date += timedelta(days=7)
                continue
            
            # Calculate focus and break time for this week
            week_focus = 0
            week_break = 0
            week_app_times = {}
            
            # Process each day in the week
            for i in range(7):
                day_date = week_start + timedelta(days=i)
                
                # Only include days in the target month
                if day_date.month == start_of_month.month and day_date.year == start_of_month.year:
                    day_str = day_date.strftime('%Y-%m-%d')
                    day_data = self.get_day_data(day_str)
                    
                    # Aggregate data
                    week_focus += day_data['focus']
                    week_break += day_data['break']
                    month_data['focus'] += day_data['focus']
                    month_data['break'] += day_data['break']
                    
                    # Aggregate app times
                    for app, time in day_data['raw_app_times'].items():
                        week_app_times[app] = week_app_times.get(app, 0) + time
                        month_data['raw_app_times'][app] = month_data['raw_app_times'].get(app, 0) + time
            
            # Format date range
            start_date = f"{week_start.day}/{week_start.month}"
            end_date = f"{week_end.day}/{week_end.month}"
            date_range = f"{start_date} - {end_date}"
            
            # Add week data
            total_minutes = week_focus + week_break
            if total_minutes > 0:
                month_data['weekly_hours'].append({
                    'week': f'Week {week_number}',
                    'dateRange': date_range,
                    'totalMinutes': total_minutes,
                    'focusMinutes': week_focus,
                    'hours': round(week_focus / 60, 1)  # For backward compatibility
                })
                week_number += 1
            
            # Move to next week
            current_date += timedelta(days=7)
        
        # Calculate app usage percentages for the month
        total_minutes = month_data['focus'] + month_data['break']
        if total_minutes > 0:
            month_data['app_usage'] = {
                app: round((time / 60 / total_minutes) * 100, 1)
                for app, time in month_data['raw_app_times'].items()
            }
        
        print(f"[DEBUG] Month data for {date_str[:7]}: {month_data}")
        return month_data
    
    def get_week_data(self, date_str):
        """Get aggregated data for the week containing the specified date"""
        date = datetime.strptime(date_str, '%Y-%m-%d')
        start_of_week = date - timedelta(days=date.weekday())  # Monday
        
        week_data = {
            'focus': 0,
            'break': 0,
            'raw_app_times': {},
            'app_usage': {},
            'daily_hours': []
        }
        
        # Process each day of the week
        day_names = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        for i in range(7):
            current_date = start_of_week + timedelta(days=i)
            day_str = current_date.strftime('%Y-%m-%d')
            day_data = self.get_day_data(day_str)
            
            # Aggregate the data
            week_data['focus'] += day_data['focus']
            week_data['break'] += day_data['break']
            
            # Track daily hours
            week_data['daily_hours'].append({
                'day': day_names[i],
                'hours': round(day_data['focus'] / 60, 1)  # Convert minutes to hours
            })
            
            # Aggregate app times
            for app, time in day_data['raw_app_times'].items():
                week_data['raw_app_times'][app] = week_data['raw_app_times'].get(app, 0) + time
        
        # Calculate app usage percentages for the week
        total_minutes = week_data['focus'] + week_data['break']
        if total_minutes > 0:
            week_data['app_usage'] = {
                app: round((time / 60 / total_minutes) * 100, 1)
                for app, time in week_data['raw_app_times'].items()
            }
        
        print(f"[DEBUG] Week data ending {date_str}: {week_data}")
        return week_data
    
    @pyqtSlot(str, str, result='QVariant')
    def get_data(self, view, date_str):
        """Get data for the specified view and date"""
        print(f"[DEBUG] Received request for {view} view, date: {date_str}")
        
        if view == 'day':
            data = self.get_day_data(date_str)
        elif view == 'week':
            data = self.get_week_data(date_str)
        else:  # month
            data = self.get_month_data(date_str)
        
        print(f"[DEBUG] Sending data to JS: {data}")
        return QVariant(data)

class SimpleDashboardView(QDialog):
    """Main window for the dashboard"""
    
    def __init__(self, parent=None, log_dir=None):
        super().__init__(parent)
        self.setWindowTitle("Simple Focus Dashboard")
        self.setMinimumSize(900, 700)
        
        # Create web view
        self.web_view = QWebEngineView(self)
        layout = QVBoxLayout(self)
        layout.addWidget(self.web_view)
        
        # Set up web channel for Python-JavaScript communication
        self.channel = QWebChannel()
        self.page = self.web_view.page()
        self.page.setWebChannel(self.channel)
        
        # Register handler object
        self.handler = SimpleDashboardHandler(log_dir)
        self.channel.registerObject("pyHandler", self.handler)
        
        # Load the HTML file
        html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dashboard.html')
        print(f"Loading HTML from: {html_path}")
        self.web_view.load(QUrl.fromLocalFile(html_path))

def main():
    """Main entry point"""
    app = QApplication(sys.argv)
    
    # Get the logs directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    logs_dir = os.path.join(script_dir, "focus_logs")
    
    print(f"Starting Simple Dashboard with logs directory: {logs_dir}")
    
    # Create and show the dashboard
    dashboard = SimpleDashboardView(log_dir=logs_dir)
    dashboard.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
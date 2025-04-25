#!/usr/bin/env python3
"""
simple_dashboard.py - A simplified version of the focus dashboard
that only handles loading CSV files and sending data to the React app.

Activity tracking modes:
- active = 0: Break/inactive time
- active = 1: Work focus time
- active = 2: Leisure focus time
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
            'focusWork': 0,
            'focusLeisure': 0,
            'focus': 0,  # Total focus (work + leisure) for backward compatibility
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
                    
                    rows = []
                    for row in reader:
                        rows.append(row)
                    
                    for row in rows:
                        # Parse the activity log CSV format
                        try:
                            elapsed_time = float(row['elapsed_time'])
                            # Active values: 0 = inactive, 1 = work focus, 2 = leisure focus
                            active_val = int(row['active'])
                            is_active_work = active_val == 1
                            is_active_leisure = active_val == 2
                            is_active = is_active_work or is_active_leisure
                            app_name = row['active_app']
                            timestamp = row['timestamp']
                            
                            # Check for work_elapsed and leisure_elapsed in the new format
                            work_elapsed = float(row.get('work_elapsed', 0))
                            leisure_elapsed = float(row.get('leisure_elapsed', 0))
                            
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
                        active_val = int(row['active'])
                        is_active = active_val > 0  # Either work (1) or leisure (2)
                        app_name = row['active_app']
                        all_logs.append((timestamp, active_val, app_name))
                
                # Sort logs by timestamp
                all_logs.sort(key=lambda x: x[0])
                
                # Second pass: create intervals with the 5 consecutive inactive logs rule
                intervals = []
                inactive_count = 0
                current_focus = None
                current_start = None
                last_timestamp = None
                
                # Reset focus and break time counters
                data['focusWork'] = 0
                data['focusLeisure'] = 0
                data['focus'] = 0
                data['break'] = 0
                data['raw_app_times'] = {}
                
                for i, (timestamp, is_active, app_name) in enumerate(all_logs):
                    # Track consecutive inactive logs
                    if not is_active:
                        inactive_count += 1
                    else:
                        inactive_count = 0
                        
                    # Determine if this is work or leisure focus
                    active_val = all_logs[i][1]  # Already an integer from the previous change
                    is_active_work = active_val == 1
                    is_active_leisure = active_val == 2
                    is_active = active_val > 0  # Either work (1) or leisure (2)
                    
                    # Check for large time gaps (> 2 minutes) between consecutive logs
                    # These should be considered breaks even if not explicitly marked as such
                    time_gap = False
                    if i > 0:
                        prev_timestamp = all_logs[i-1][0]
                        gap = (timestamp - prev_timestamp).total_seconds()
                        if gap > 120:  # More than 2 minutes
                            time_gap = True
                    
                    # Determine if this should be considered a focus or break period
                    # Count as break if we have 5+ consecutive inactive logs or a large time gap
                    is_break = inactive_count >= 5 or time_gap
                    is_focus = is_active or (not is_active and not is_break)  # is_active is true for both work and leisure
                    
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
                                # Determine if it was work or leisure focus
                                # Use the last known active state in this interval
                                last_active_val = 0
                                for j in range(i-1, -1, -1):
                                    if all_logs[j][0] >= current_start and all_logs[j][0] <= last_timestamp:
                                        val = all_logs[j][1]  # Already an integer
                                        if val > 0:  # If it was active (1 or 2)
                                            last_active_val = val
                                            break
                                
                                # Check for restart markers in this interval
                                has_restart = False
                                with log_path.open('r') as f:
                                    reader = csv.DictReader(f)
                                    for row in reader:
                                        row_time = datetime.strptime(row['timestamp'], '%Y-%m-%d %H:%M:%S')
                                        if row_time >= current_start and row_time <= last_timestamp:
                                            note = row.get('note', '')
                                            if note == 'restart':
                                                has_restart = True
                                                break
                                
                                # If there was a restart in this interval, adjust the duration
                                if not has_restart:
                                    if last_active_val == 1:  # Work focus
                                        data['focusWork'] += interval_duration
                                    elif last_active_val == 2:  # Leisure focus
                                        data['focusLeisure'] += interval_duration
                                    
                                    # Also add to total focus for backward compatibility
                                    data['focus'] += interval_duration
                            else:  # If it was a break interval
                                data['break'] += interval_duration
                                
                            # Determine the mode for this interval
                            mode = 'break'
                            if current_focus:
                                # Check if this was a work or leisure interval
                                # Use the last known active state in this interval
                                last_active_val = 0
                                for j in range(i-1, -1, -1):
                                    if all_logs[j][0] >= current_start and all_logs[j][0] <= last_timestamp:
                                        val = all_logs[j][1]  # Already an integer
                                        if val > 0:  # If it was active (1 or 2)
                                            last_active_val = val
                                            break
                                mode = 'work' if last_active_val == 1 else 'leisure'
                            
                            intervals.append({
                                'start': current_start.isoformat(),
                                'end': last_timestamp.isoformat(),
                                'focus': current_focus,
                                'mode': mode
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
                        # Determine if it was work or leisure focus
                        # Use the last known active state in this interval
                        last_active_val = 0
                        for j in range(len(all_logs)-1, -1, -1):
                            if all_logs[j][0] >= current_start and all_logs[j][0] <= last_timestamp:
                                val = all_logs[j][1]  # Already an integer
                                if val > 0:  # If it was active (1 or 2)
                                    last_active_val = val
                                    break
                        
                        if last_active_val == 1:  # Work focus
                            data['focusWork'] += interval_duration
                        elif last_active_val == 2:  # Leisure focus
                            data['focusLeisure'] += interval_duration
                        
                        # Also add to total focus for backward compatibility
                        data['focus'] += interval_duration
                    else:  # If it was a break interval
                        data['break'] += interval_duration
                        
                    # Determine the mode for this interval
                    mode = 'break'
                    if current_focus:
                        # Check if this was a work or leisure interval
                        # Use the last known active state in this interval
                        last_active_val = 0
                        for j in range(len(all_logs)-1, -1, -1):
                            if all_logs[j][0] >= current_start and all_logs[j][0] <= last_timestamp:
                                val = all_logs[j][1]  # Already an integer
                                if val > 0:  # If it was active (1 or 2)
                                    last_active_val = val
                                    break
                        
                        mode = 'work' if last_active_val == 1 else 'leisure'
                    
                    intervals.append({
                        'start': current_start.isoformat(),
                        'end': last_timestamp.isoformat(),
                        'focus': current_focus,
                        'mode': mode
                    })
                
                # Convert seconds to minutes
                data['focusWork'] = round(data['focusWork'] / 60)
                data['focusLeisure'] = round(data['focusLeisure'] / 60)
                data['focus'] = round(data['focus'] / 60)
                data['break'] = round(data['break'] / 60)
                
                # Cap total time at 24 hours (1440 minutes) if it exceeds that
                total_minutes = data['focus'] + data['break']
                if total_minutes > 1440:
                    scaling_factor = 1440 / total_minutes
                    data['focusWork'] = round(data['focusWork'] * scaling_factor)
                    data['focusLeisure'] = round(data['focusLeisure'] * scaling_factor)
                    data['focus'] = round(data['focus'] * scaling_factor)
                    data['break'] = round(data['break'] * scaling_factor)
                    print(f"[DEBUG] Scaled down times to fit within 24 hours. New focus work: {data['focusWork']}m, focus leisure: {data['focusLeisure']}m, break: {data['break']}m")
                
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
            'focusWork': 0,
            'focusLeisure': 0,
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
                    week_focus_work = day_data.get('focusWork', 0)
                    week_focus_leisure = day_data.get('focusLeisure', 0)
                    week_break += day_data['break']
                    
                    month_data['focus'] += day_data['focus']
                    month_data['focusWork'] += day_data.get('focusWork', 0)
                    month_data['focusLeisure'] += day_data.get('focusLeisure', 0)
                    month_data['break'] += day_data['break']
                    
                    # Aggregate app times
                    for app, time in day_data['raw_app_times'].items():
                        week_app_times[app] = week_app_times.get(app, 0) + time
                        month_data['raw_app_times'][app] = month_data['raw_app_times'].get(app, 0) + time
            
            # Format date range
            start_date = f"{week_start.day}/{week_start.month}"
            end_date = f"{week_end.day}/{week_end.month}"
            date_range = f"{start_date} - {end_date}"
            
            # Add week data with work focus, leisure focus, and break time
            total_minutes = week_focus + week_break
            if total_minutes > 0:
                month_data['weekly_hours'].append({
                    'week': f'Week {week_number}',
                    'dateRange': date_range,
                    'totalMinutes': total_minutes,
                    'focusMinutes': week_focus,
                    'focusWorkMinutes': week_focus_work,
                    'focusLeisureMinutes': week_focus_leisure,
                    'breakMinutes': week_break,
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
        print(f"[DEBUG] Month totals - Focus Work: {month_data['focusWork']}, Focus Leisure: {month_data['focusLeisure']}, Break: {month_data['break']}")
        return month_data
    
    def get_week_data(self, date_str):
        """Get aggregated data for the week containing the specified date"""
        date = datetime.strptime(date_str, '%Y-%m-%d')
        start_of_week = date - timedelta(days=date.weekday())  # Monday
        
        week_data = {
            'focus': 0,
            'focusWork': 0,
            'focusLeisure': 0,
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
            week_data['focusWork'] += day_data.get('focusWork', 0)
            week_data['focusLeisure'] += day_data.get('focusLeisure', 0)
            week_data['break'] += day_data['break']
            
            # Track daily hours with work focus, leisure focus, and break time
            day_focus = day_data['focus']
            day_focus_work = day_data.get('focusWork', 0)
            day_focus_leisure = day_data.get('focusLeisure', 0)
            day_break = day_data['break']
            
            # Always include the day, even if zero hours
            week_data['daily_hours'].append({
                'day': day_names[i],
                'date': current_date.strftime('%d/%m'),
                'focusHours': round(day_focus / 60, 1),  # Convert minutes to hours
                'focusWorkHours': round(day_focus_work / 60, 1),  # Convert minutes to hours
                'focusLeisureHours': round(day_focus_leisure / 60, 1),  # Convert minutes to hours
                'breakHours': round(day_break / 60, 1),  # Convert minutes to hours
                'focusMinutes': day_focus,
                'focusWorkMinutes': day_focus_work,
                'focusLeisureMinutes': day_focus_leisure,
                'breakMinutes': day_break,
                'totalMinutes': day_focus + day_break
            })
            
            # Aggregate app times
            for app, time in day_data['raw_app_times'].items():
                week_data['raw_app_times'][app] = week_data['raw_app_times'].get(app, 0) + time
        
        # Calculate app usage percentages for the week - only based on active app time
        total_app_time = sum(week_data['raw_app_times'].values())
        if total_app_time > 0:
            week_data['app_usage'] = {
                app: round((time / total_app_time) * 100, 1)
                for app, time in week_data['raw_app_times'].items()
            }
        
        print(f"[DEBUG] Week data ending {date_str}: {week_data}")
        print(f"[DEBUG] Focus Work: {week_data['focusWork']}, Focus Leisure: {week_data['focusLeisure']}, Break: {week_data['break']}")
        print(f"[DEBUG] daily_hours: {week_data['daily_hours']}")
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


class SimpleDashboardAdapter(QDialog):
    """
    Adapter class that wraps the SimpleDashboardView to make it compatible
    with the main focus timer application interface.
    """
    
    def __init__(self, parent=None, log_dir=None):
        """
        Initialize the dashboard adapter.
        
        Args:
            parent: Parent widget
            log_dir: Directory containing the focus logs
        """
        super().__init__(parent)
        
        # Create and show the dashboard
        self.dashboard = SimpleDashboardView(parent=self, log_dir=log_dir)
        
        # Copy the layout from the dashboard
        self.setLayout(self.dashboard.layout())
        
        # Set window properties
        self.setWindowTitle(self.dashboard.windowTitle())
        self.setMinimumSize(self.dashboard.minimumSize())


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
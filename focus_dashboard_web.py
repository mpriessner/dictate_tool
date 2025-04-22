#!/usr/bin/env python3
import os
from datetime import datetime
from pathlib import Path
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtWebChannel import QWebChannel
from PyQt5.QtCore import QObject, pyqtSlot, QUrl, QVariant
from PyQt5.QtWidgets import QDialog, QVBoxLayout
import csv
from datetime import timedelta

class DashboardHandler(QObject):
    def __init__(self, log_dir):
        super().__init__()
        self.log_dir = Path(log_dir)
        
    @pyqtSlot(str)
    def updateView(self, view):
        """Called when the view mode changes in the web UI"""
        print(f"View changed to: {view}")
        # TODO: Load and send data for the new view
        
    @pyqtSlot(str)
    def updateDate(self, date):
        """Called when the date changes in the web UI"""
        print(f"Date changed to: {date}")
        # TODO: Load and send data for the new date

    def get_log_path(self, date_str):
        """Get the path to the activity log file for a specific date"""
        return self.log_dir / date_str / "activity_log.csv"

    def get_day_data(self, date_str):
        """Get focus and break data for a specific day"""
        log_path = self.get_log_path(date_str)
        print(f"[DEBUG] Reading day data from: {log_path}")
        data = {
            'focus': 0,
            'break': 0,
            'raw_app_times': {},
            'app_usage': {}
        }

        if log_path.exists():
            with log_path.open('r') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    # Parse the activity log CSV format
                    elapsed_time = float(row['elapsed_time'])
                    is_active = row['active'].lower() == 'true'
                    app_name = row['active_app']

                    if is_active:
                        data['focus'] += elapsed_time
                        # Track time per application
                        data['raw_app_times'][app_name] = data['raw_app_times'].get(app_name, 0) + elapsed_time
                    else:
                        data['break'] += elapsed_time

        # Convert seconds to minutes
        data['focus'] = round(data['focus'] / 60)
        data['break'] = round(data['break'] / 60)

        # Calculate app usage percentages
        total_minutes = data['focus'] + data['break']
        if total_minutes > 0:
            data['app_usage'] = {
                app: round((time / 60 / total_minutes) * 100, 1)
                for app, time in data['raw_app_times'].items()
            }

        print(f"[DEBUG] Day data for {date_str}: {data}")
        return data

    def get_week_data(self, date_str):
        """Get aggregated data for the week containing the specified date"""
        date = datetime.strptime(date_str, '%Y-%m-%d')
        start_of_week = date - timedelta(days=date.weekday())  # Monday
        
        week_data = {
            'focus': 0,
            'break': 0,
            'raw_app_times': {},
            'app_usage': {},
            'daily_hours': []  # Track focus time for each day
        }

        # Process each day of the week
        for i in range(7):
            current_date = start_of_week + timedelta(days=i)
            day_str = current_date.strftime('%Y-%m-%d')
            day_data = self.get_day_data(day_str)
            
            # Aggregate the data
            week_data['focus'] += day_data['focus']
            week_data['break'] += day_data['break']
            
            # Track daily hours
            week_data['daily_hours'].append({
                'day': ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'][i],
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

    def get_month_data(self, date_str):
        """Get aggregated data for the month containing the specified date"""
        date = datetime.strptime(date_str, '%Y-%m-%d')
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
            'weekly_hours': []  # Track focus time for each week
        }

        # Process each day of the month
        current_week_hours = 0
        week_number = 1
        
        for i in range(days_in_month):
            current_date = start_of_month + timedelta(days=i)
            day_str = current_date.strftime('%Y-%m-%d')
            day_data = self.get_day_data(day_str)
            
            # Aggregate the data
            month_data['focus'] += day_data['focus']
            month_data['break'] += day_data['break']
            
            # Add to current week's hours
            current_week_hours += day_data['focus'] / 60  # Convert minutes to hours
            
            # If it's the end of a week or month, save the weekly total
            if current_date.weekday() == 6 or i == days_in_month - 1:
                month_data['weekly_hours'].append({
                    'week': f'Week {week_number}',
                    'hours': round(current_week_hours, 1)
                })
                current_week_hours = 0
                week_number += 1
            
            # Aggregate app times
            for app, time in day_data['raw_app_times'].items():
                month_data['raw_app_times'][app] = month_data['raw_app_times'].get(app, 0) + time

        # Calculate app usage percentages for the month
        total_minutes = month_data['focus'] + month_data['break']
        if total_minutes > 0:
            month_data['app_usage'] = {
                app: round((time / 60 / total_minutes) * 100, 1)
                for app, time in month_data['raw_app_times'].items()
            }

        print(f"[DEBUG] Month data for {date_str[:7]}: {month_data}")
        return month_data

    @pyqtSlot(str, str, result='QVariant')
    def get_data(self, view, date_str):
        """Get data for the specified view and date"""
        if view == 'day':
            data = self.get_day_data(date_str)
        elif view == 'week':
            data = self.get_week_data(date_str)
        elif view == 'month':
            data = self.get_month_data(date_str)
        else:
            data = {}

        print(f"[DEBUG] Sending data to JS: {data}")
        return QVariant(data)

class DashboardWebView(QDialog):
    def __init__(self, parent=None, log_dir=None):
        super().__init__(parent)
        self.setWindowTitle("Focus Dashboard")
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
        self.handler = DashboardHandler(log_dir)
        self.channel.registerObject("pyHandler", self.handler)
        
        # Load the HTML file
        html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dashboard.html')
        self.web_view.load(QUrl.fromLocalFile(html_path))

if __name__ == "__main__":
    import sys
    from PyQt5.QtWidgets import QApplication
    
    app = QApplication(sys.argv)
    logbase = Path(sys.argv[1] if len(sys.argv) > 1 else "./focus_logs")
    
    win = DashboardWebView(log_dir=logbase)
    win.show()
    sys.exit(app.exec_())

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
            'app_usage': {}
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
                            
                            if is_active:
                                data['focus'] += elapsed_time
                                # Track time per application
                                data['raw_app_times'][app_name] = data['raw_app_times'].get(app_name, 0) + elapsed_time
                            else:
                                data['break'] += elapsed_time
                        except Exception as e:
                            print(f"Error processing row: {e}")
                            print(f"Problematic row: {row}")
            except Exception as e:
                print(f"Error reading file: {e}")
        else:
            print(f"Log file does not exist: {log_path}")
        
        # Convert seconds to minutes
        data['focus'] = round(data['focus'] / 60)
        data['break'] = round(data['break'] / 60)
        
        # Calculate app usage percentages
        total_minutes = data['focus'] + data['break']
        if total_minutes > 0:
            print(f"[DEBUG] Raw app times: {data['raw_app_times']}")
            
            # Calculate percentages
            data['app_usage'] = {}
            for app, time in data['raw_app_times'].items():
                percentage = round((time / 60 / total_minutes) * 100, 1)
                data['app_usage'][app] = percentage
                print(f"[DEBUG] App {app}: {time} seconds = {percentage}% of total time")
        else:
            print("[DEBUG] No focus or break time recorded, can't calculate app usage percentages")
        
        print(f"[DEBUG] Day data for {date_str}: {data}")
        return data
    
    @pyqtSlot(str, str, result='QVariant')
    def get_data(self, view, date_str):
        """Get data for the specified view and date"""
        print(f"[DEBUG] Received request for {view} view, date: {date_str}")
        
        # For simplicity, just return day data for all views
        data = self.get_day_data(date_str)
        
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
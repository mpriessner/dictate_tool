#!/usr/bin/env python3
# activity_manager.py - A tool to manage focus timer activity logs

import os
import sys
import csv
import datetime
import pandas as pd
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QComboBox, QDateEdit, QTimeEdit, 
    QSpinBox, QDialog, QFormLayout, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QFileDialog, QTabWidget
)
from PyQt5.QtCore import Qt, QDate, QTime, QDateTime

class ActivityManager(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Constants
        self.MODE_WORK = 1  # Work focus mode
        self.MODE_LEISURE = 2  # Leisure focus mode
        self.MODE_BREAK = 0  # Break (inactive)
        
        # Set up the base log directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.logs_dir = os.path.join(script_dir, "focus_logs")
        
        # Ensure logs directory exists
        if not os.path.exists(self.logs_dir):
            os.makedirs(self.logs_dir, exist_ok=True)
            print(f"Created logs directory: {self.logs_dir}")
        else:
            print(f"Using existing logs directory: {self.logs_dir}")
        
        # Setup UI
        self.setup_ui()
        
        # Load available dates
        self.load_available_dates()
    
    def setup_ui(self):
        self.setWindowTitle("Focus Timer Activity Manager")
        self.setMinimumSize(800, 600)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Add status label at the top
        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("color: gray; font-style: italic;")
        main_layout.addWidget(self.status_label)
        
        # Create tab widget
        tab_widget = QTabWidget()
        main_layout.addWidget(tab_widget)
        
        # Add tab
        add_widget = QWidget()
        add_layout = QVBoxLayout(add_widget)
        tab_widget.addTab(add_widget, "Add Activity")
        
        # View/Delete tab
        view_widget = QWidget()
        view_layout = QVBoxLayout(view_widget)
        tab_widget.addTab(view_widget, "View/Delete Activities")
        
        # === ADD ACTIVITY TAB ===
        form_layout = QFormLayout()
        add_layout.addLayout(form_layout)
        
        # Date selector
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.dateChanged.connect(self.update_log_file_path)
        form_layout.addRow("Date:", self.date_edit)
        
        # Start time selector
        self.start_time_edit = QTimeEdit()
        self.start_time_edit.setDisplayFormat("HH:mm")
        self.start_time_edit.setTime(QTime.currentTime())
        form_layout.addRow("Start Time:", self.start_time_edit)
        
        # Duration in minutes
        self.duration_spin = QSpinBox()
        self.duration_spin.setRange(1, 480)  # 1 minute to 8 hours
        self.duration_spin.setValue(30)
        self.duration_spin.setSuffix(" minutes")
        form_layout.addRow("Duration:", self.duration_spin)
        
        # Activity type
        self.activity_type_combo = QComboBox()
        self.activity_type_combo.addItem("Work Focus", self.MODE_WORK)
        self.activity_type_combo.addItem("Leisure Focus", self.MODE_LEISURE)
        self.activity_type_combo.addItem("Break (Inactive)", self.MODE_BREAK)
        form_layout.addRow("Activity Type:", self.activity_type_combo)
        
        # Application name
        self.app_name_combo = QComboBox()
        self.app_name_combo.setEditable(True)
        self.app_name_combo.addItems(["Meeting", "Microsoft Teams", "Zoom", "Google Meet", "Microsoft Word", "Microsoft Excel", "Visual Studio Code", "PyCharm", "Safari", "Google Chrome", "Mail", "Calendar"])
        form_layout.addRow("Application:", self.app_name_combo)
        
        # URL/Note
        self.url_combo = QComboBox()
        self.url_combo.setEditable(True)
        self.url_combo.addItems(["n/a", "meeting", "work", "leisure", "break", "call", "presentation"])
        form_layout.addRow("URL/Note:", self.url_combo)
        
        # Log file path display
        self.log_path_label = QLabel()
        form_layout.addRow("Log File:", self.log_path_label)
        self.update_log_file_path()
        
        # Add activity button
        add_button_layout = QHBoxLayout()
        add_layout.addLayout(add_button_layout)
        
        self.add_activity_button = QPushButton("Add Activity")
        self.add_activity_button.clicked.connect(self.add_activity)
        add_button_layout.addWidget(self.add_activity_button)
        
        # === VIEW/DELETE TAB ===
        date_selector_layout = QHBoxLayout()
        view_layout.addLayout(date_selector_layout)
        
        date_selector_layout.addWidget(QLabel("Select Date:"))
        self.view_date_combo = QComboBox()
        self.view_date_combo.currentIndexChanged.connect(self.load_activities_for_date)
        date_selector_layout.addWidget(self.view_date_combo)
        
        date_selector_layout.addStretch()
        
        refresh_button = QPushButton("Refresh")
        refresh_button.clicked.connect(self.load_available_dates)
        date_selector_layout.addWidget(refresh_button)
        
        # Table for viewing activities
        self.activities_table = QTableWidget()
        self.activities_table.setColumnCount(9)
        self.activities_table.setHorizontalHeaderLabels([
            "Timestamp", "Type", "Duration (min)", "Work Time", "Leisure Time", 
            "Total Work (h)", "Total Leisure (h)", "App", "URL/Note"
        ])
        self.activities_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.activities_table.horizontalHeader().setStretchLastSection(True)
        view_layout.addWidget(self.activities_table)
        
        # Delete button
        delete_button = QPushButton("Delete Selected Activity")
        delete_button.clicked.connect(self.delete_activity)
        view_layout.addWidget(delete_button)
    
    def update_log_file_path(self):
        selected_date = self.date_edit.date().toString("yyyy-MM-dd")
        log_file_path = os.path.join(self.logs_dir, selected_date, "activity_log.csv")
        self.log_path_label.setText(log_file_path)
    
    def load_available_dates(self):
        # Clear existing items
        self.view_date_combo.clear()
        
        # Check if logs directory exists
        if not os.path.exists(self.logs_dir):
            print(f"Logs directory does not exist: {self.logs_dir}")
            return
        
        # Get all subdirectories in the logs directory (each is a date)
        date_dirs = []
        try:
            for item in os.listdir(self.logs_dir):
                item_path = os.path.join(self.logs_dir, item)
                if os.path.isdir(item_path):
                    # Check if the directory contains an activity_log.csv file
                    log_file = os.path.join(item_path, "activity_log.csv")
                    if os.path.exists(log_file):
                        try:
                            # Validate that the directory name is a date
                            datetime.datetime.strptime(item, "%Y-%m-%d")
                            date_dirs.append(item)
                            print(f"Found valid date directory: {item}")
                        except ValueError:
                            # Not a valid date format, skip
                            print(f"Skipping invalid date format: {item}")
                            continue
        except Exception as e:
            print(f"Error loading date directories: {e}")
        
        # Sort dates in reverse order (newest first)
        date_dirs.sort(reverse=True)
        print(f"Found {len(date_dirs)} date directories")
        
        # Add to combo box
        self.view_date_combo.addItems(date_dirs)
        
        # Load activities for the first date if available
        if self.view_date_combo.count() > 0:
            self.load_activities_for_date()
    
    def load_activities_for_date(self):
        # Clear existing table
        self.activities_table.setRowCount(0)
        
        # Get selected date
        if self.view_date_combo.count() == 0:
            return
        
        selected_date = self.view_date_combo.currentText()
        log_file_path = os.path.join(self.logs_dir, selected_date, "activity_log.csv")
        
        if not os.path.exists(log_file_path):
            return
        
        try:
            # Read the CSV file using pandas for better handling
            df = pd.read_csv(log_file_path)
            
            # Sort by timestamp to ensure chronological order
            if 'timestamp' in df.columns:
                df['timestamp'] = pd.to_datetime(df['timestamp'])
                df = df.sort_values('timestamp')
                df['timestamp'] = df['timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S')
            
            # Populate table
            self.activities_table.setRowCount(len(df))
            
            for i, (_, row) in enumerate(df.iterrows()):
                # Timestamp
                self.activities_table.setItem(i, 0, QTableWidgetItem(str(row['timestamp'])))
                
                # Activity type
                activity_type = int(row['active'])
                if activity_type == self.MODE_WORK:
                    type_text = "Work Focus"
                elif activity_type == self.MODE_LEISURE:
                    type_text = "Leisure Focus"
                else:
                    type_text = "Break (Inactive)"
                self.activities_table.setItem(i, 1, QTableWidgetItem(type_text))
                
                # Duration in minutes
                duration_sec = int(row['elapsed_time'])
                duration_min = round(duration_sec / 60)
                self.activities_table.setItem(i, 2, QTableWidgetItem(str(duration_min)))
                
                # Work elapsed
                work_elapsed_sec = int(row['work_elapsed'])
                work_elapsed_min = round(work_elapsed_sec / 60)
                self.activities_table.setItem(i, 3, QTableWidgetItem(str(work_elapsed_min)))
                
                # Leisure elapsed
                leisure_elapsed_sec = int(row['leisure_elapsed'])
                leisure_elapsed_min = round(leisure_elapsed_sec / 60)
                self.activities_table.setItem(i, 4, QTableWidgetItem(str(leisure_elapsed_min)))
                
                # Total work hours
                self.activities_table.setItem(i, 5, QTableWidgetItem(str(row['total_work_hours'])))
                
                # Total leisure hours
                self.activities_table.setItem(i, 6, QTableWidgetItem(str(row['total_leisure_hours'])))
                
                # App name
                self.activities_table.setItem(i, 7, QTableWidgetItem(str(row['active_app'])))
                
                # URL/Note
                if 'url' in row and pd.notna(row['url']):
                    self.activities_table.setItem(i, 8, QTableWidgetItem(str(row['url'])))
                else:
                    self.activities_table.setItem(i, 8, QTableWidgetItem(""))
                
                # Highlight manual entries
                if 'note' in row and row['note'] == 'manual':
                    for col in range(9):
                        item = self.activities_table.item(i, col)
                        if item:
                            item.setBackground(Qt.lightGray)
        
        except Exception as e:
            print(f"Error loading activities: {e}")
            import traceback
            traceback.print_exc()
    
    def add_activity(self):
        # Get selected date and create log directory if it doesn't exist
        selected_date = self.date_edit.date().toString("yyyy-MM-dd")
        date_dir = os.path.join(self.logs_dir, selected_date)
        os.makedirs(date_dir, exist_ok=True)
        
        log_file_path = os.path.join(date_dir, "activity_log.csv")
        
        # Check if log file exists, create with header if not
        if not os.path.exists(log_file_path):
            with open(log_file_path, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["timestamp", "active", "elapsed_time", "work_elapsed", "leisure_elapsed", 
                               "total_work_hours", "total_leisure_hours", "active_app", "url", "note"])
        
        # Get form values
        start_time = self.start_time_edit.time()
        duration_minutes = self.duration_spin.value()
        activity_type = self.activity_type_combo.currentData()
        app_name = self.app_name_combo.currentText()
        url_note = self.url_combo.currentText()
        
        # Create start and end datetime objects
        start_datetime = QDateTime(self.date_edit.date(), start_time).toPyDateTime()
        end_datetime = start_datetime + datetime.timedelta(minutes=duration_minutes)
        
        # Load the existing CSV file using pandas
        try:
            # Read the CSV file
            if os.path.exists(log_file_path) and os.path.getsize(log_file_path) > 0:
                df = pd.read_csv(log_file_path)
                print(f"Loaded CSV with {len(df)} rows")
                
                # Convert timestamp column to datetime
                df['timestamp'] = pd.to_datetime(df['timestamp'])
                
                # Get the latest work and leisure hours
                if not df.empty:
                    total_work_hours = float(df.iloc[-1]['total_work_hours'])
                    total_leisure_hours = float(df.iloc[-1]['total_leisure_hours'])
                    print(f"Loaded existing totals: Work={total_work_hours}h, Leisure={total_leisure_hours}h")
                else:
                    total_work_hours = 0
                    total_leisure_hours = 0
                    print("No existing entries found, starting with zero totals")
                
                # Check if there are any entries in the time range we're adding
                # Convert to pandas datetime for comparison
                start_pd = pd.to_datetime(start_datetime)
                end_pd = pd.to_datetime(end_datetime)
                
                # Find entries that fall within our time range
                overlapping_entries = df[(df['timestamp'] >= start_pd) & 
                                         (df['timestamp'] < end_pd))]
                
                if not overlapping_entries.empty:
                    print(f"Found {len(overlapping_entries)} overlapping entries in the time range")
                    # Ask user for confirmation to overwrite
                    reply = QMessageBox.question(
                        self,
                        "Overlapping Entries",
                        f"Found {len(overlapping_entries)} existing entries in the selected time range. Overwrite them?",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.No
                    )
                    
                    if reply == QMessageBox.No:
                        return
                    
                    # Remove the overlapping entries
                    df = df[~((df['timestamp'] >= start_pd) & (df['timestamp'] < end_pd))]
                    print(f"Removed overlapping entries, {len(df)} entries remaining")
                    
                    # Recalculate the total work and leisure hours after removing entries
                    # Find the last entry before our start time to get the correct starting totals
                    previous_entries = df[df['timestamp'] < start_pd]
                    if not previous_entries.empty:
                        last_previous = previous_entries.iloc[-1]
                        total_work_hours = float(last_previous['total_work_hours'])
                        total_leisure_hours = float(last_previous['total_leisure_hours'])
                        print(f"Using totals from previous entry: Work={total_work_hours}h, Leisure={total_leisure_hours}h")
            else:
                # Create a new DataFrame with the correct columns
                columns = ["timestamp", "active", "elapsed_time", "work_elapsed", "leisure_elapsed", 
                          "total_work_hours", "total_leisure_hours", "active_app", "url", "note"]
                df = pd.DataFrame(columns=columns)
                total_work_hours = 0
                total_leisure_hours = 0
                print("Created new DataFrame for log file")
            
            # Generate a row for each minute in the duration
            new_rows = []
            current_time = start_datetime
            
            # Get the base elapsed time from the start of the day (in seconds)
            # This is important to match the format of the existing log file
            day_start = datetime.datetime.combine(start_datetime.date(), datetime.time.min)
            base_elapsed = int((start_datetime - day_start).total_seconds())
            
            # For each minute in the duration
            for minute in range(duration_minutes):
                # Calculate the elapsed time for this minute (in seconds since start of day)
                elapsed_time_seconds = base_elapsed + (minute * 60)
                
                # Calculate work and leisure elapsed based on activity type (60 seconds per minute)
                work_elapsed = 60 if activity_type == self.MODE_WORK else 0
                leisure_elapsed = 60 if activity_type == self.MODE_LEISURE else 0
                
                # Update total hours
                minute_work_hours = work_elapsed / 3600  # Convert seconds to hours
                minute_leisure_hours = leisure_elapsed / 3600  # Convert seconds to hours
                
                # Always add the minute's contribution to the existing totals
                total_work_hours += minute_work_hours
                total_leisure_hours += minute_leisure_hours
                
                # Print debug information for the first and last minute
                if minute == 0 or minute == duration_minutes - 1:
                    print(f"Minute {minute}: Work={total_work_hours}h, Leisure={total_leisure_hours}h")
                
                # Round to 2 decimal places
                total_work_hours_rounded = round(total_work_hours, 2)
                total_leisure_hours_rounded = round(total_leisure_hours, 2)
                
                # Format the timestamp
                timestamp = current_time.strftime("%Y-%m-%d %H:%M:%S")
                
                # Create a row for this minute
                new_row = {
                    "timestamp": timestamp,
                    "active": activity_type,  # 0=break, 1=work, 2=leisure
                    "elapsed_time": elapsed_time_seconds,
                    "work_elapsed": work_elapsed,
                    "leisure_elapsed": leisure_elapsed,
                    "total_work_hours": total_work_hours_rounded,
                    "total_leisure_hours": total_leisure_hours_rounded,
                    "active_app": app_name,
                    "url": url_note,
                    "note": "manual"
                }
                
                new_rows.append(new_row)
                
                # Move to the next minute
                current_time += datetime.timedelta(minutes=1)
            
            # Add the new rows to the DataFrame
            new_df = pd.DataFrame(new_rows)
            if 'timestamp' in new_df.columns:
                new_df['timestamp'] = pd.to_datetime(new_df['timestamp'])
            
            # Concatenate with the existing DataFrame
            df = pd.concat([df, new_df], ignore_index=True)
            
            # Sort by timestamp
            df = df.sort_values('timestamp')
            
            # After sorting, we need to recalculate all cumulative totals to ensure consistency
            # Initialize counters
            running_work_hours = 0
            running_leisure_hours = 0
            
            # Process each row in chronological order
            for idx, row in df.iterrows():
                # Add the contribution from this minute
                if row['active'] == self.MODE_WORK:
                    running_work_hours += 60/3600  # 60 seconds = 1/60 hour
                elif row['active'] == self.MODE_LEISURE:
                    running_leisure_hours += 60/3600
                
                # Update the totals in the DataFrame
                df.at[idx, 'total_work_hours'] = round(running_work_hours, 2)
                df.at[idx, 'total_leisure_hours'] = round(running_leisure_hours, 2)
            
            print(f"Recalculated totals: Final Work={running_work_hours}h, Leisure={running_leisure_hours}h")
            
            # Convert timestamp back to string format before saving
            df['timestamp'] = df['timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S')
            
            # Ensure all numeric columns have the right data type
            df['active'] = df['active'].astype(int)
            df['elapsed_time'] = df['elapsed_time'].astype(int)
            df['work_elapsed'] = df['work_elapsed'].astype(int)
            df['leisure_elapsed'] = df['leisure_elapsed'].astype(int)
            
            # Make sure we have the correct column order to match the original CSV
            column_order = ["timestamp", "active", "elapsed_time", "work_elapsed", "leisure_elapsed", 
                           "total_work_hours", "total_leisure_hours", "active_app", "url", "note"]
            df = df[column_order]
            
            # Save the updated DataFrame to the CSV file - explicitly overwrite the file
            print(f"Writing {len(df)} rows to {log_file_path}")
            df.to_csv(log_file_path, index=False, mode='w')
            
            print(f"Successfully wrote {len(new_rows)} entries to {log_file_path}")
            
            # Show confirmation
            QMessageBox.information(
                self, 
                "Activity Added", 
                f"Added {duration_minutes} minutes of {self.activity_type_combo.currentText()} to {selected_date}."
            )
            
            # Update status label
            self.status_label.setText(f"Added {duration_minutes} minutes of {self.activity_type_combo.currentText()} to {selected_date}")
            self.status_label.setStyleSheet("color: green; font-weight: bold;")
            
            # Refresh the view if we're looking at the same date
            if self.view_date_combo.currentText() == selected_date:
                self.load_activities_for_date()
            
            # Refresh available dates in case we added a new date
            self.load_available_dates()
            
        except Exception as e:
            print(f"Error processing log file: {e}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to add activity: {str(e)}"
            )
    
    def delete_activity(self):
        # Check if a row is selected
        selected_rows = self.activities_table.selectedIndexes()
        if not selected_rows:
            QMessageBox.warning(self, "No Selection", "Please select an activity to delete.")
            return
        
        # Get the row of the first selected cell
        row = selected_rows[0].row()
        
        # Get the timestamp of the selected activity for identification
        timestamp = self.activities_table.item(row, 0).text()
        
        # Confirm deletion
        reply = QMessageBox.question(
            self, 
            "Confirm Deletion", 
            f"Are you sure you want to delete the activity at {timestamp}?",
            QMessageBox.Yes | QMessageBox.No, 
            QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
        
        # Get the log file path
        selected_date = self.view_date_combo.currentText()
        log_file_path = os.path.join(self.logs_dir, selected_date, "activity_log.csv")
        
        if not os.path.exists(log_file_path):
            QMessageBox.warning(self, "Error", "Log file not found.")
            return
        
        try:
            # Read the CSV file using pandas
            df = pd.read_csv(log_file_path)
            
            # Convert timestamp to datetime for proper comparison
            if 'timestamp' in df.columns:
                df['timestamp'] = pd.to_datetime(df['timestamp'])
            
            # Find the row with the matching timestamp
            timestamp_dt = pd.to_datetime(timestamp)
            df = df[df['timestamp'] != timestamp_dt]
            
            # Convert timestamp back to string format before saving
            df['timestamp'] = df['timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S')
            
            # Save the updated DataFrame to the CSV file
            df.to_csv(log_file_path, index=False)
            
            # Refresh the view
            self.load_activities_for_date()
            
            # Show confirmation
            QMessageBox.information(self, "Activity Deleted", f"Deleted activity at {timestamp}.")
            
        except Exception as e:
            print(f"Error deleting activity: {e}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Error", f"Failed to delete activity: {str(e)}")

def main():
    app = QApplication(sys.argv)
    window = ActivityManager()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
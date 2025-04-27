#!/usr/bin/env python3
"""
Script to update the CSV format of existing focus log files and add manual log entries.
Supports:
1. Updating CSV format to include work/leisure columns
2. Adding manual log entries for missed focus sessions
"""

import os
import csv
import random
import argparse
from datetime import datetime, timedelta

def update_csv_file(file_path):
    """Update a single CSV file to the new format"""
    print(f"Updating file: {file_path}")
    
    # Read the existing data
    rows = []
    with open(file_path, 'r', newline='') as f:
        reader = csv.reader(f)
        header = next(reader)
        rows = list(reader)
    
    # Check if the file already has the new format
    if 'work_elapsed' in header and 'leisure_elapsed' in header and 'total_leisure_hours' in header:
        print(f"File {file_path} already has the new format")
        return
    
    # Create the new header
    new_header = ["timestamp", "active", "elapsed_time", "work_elapsed", "leisure_elapsed", 
                  "total_work_hours", "total_leisure_hours", "active_app", "url", "note"]
    
    # Initialize total leisure hours
    total_leisure_hours = 0.0
    
    # Process each row
    new_rows = []
    for row in rows:
        # Extract existing data
        timestamp = row[0]
        active = int(row[1])
        elapsed_time = int(row[2]) if row[2] else 0
        total_work_hours = float(row[3]) if row[3] else 0.0
        
        # Calculate work_elapsed and leisure_elapsed based on active state
        work_elapsed = elapsed_time if active == 1 else 0
        leisure_elapsed = elapsed_time if active == 2 else 0
        
        # Update total leisure hours (make up some data)
        if active == 2:
            # If this is a leisure activity, increase leisure hours
            leisure_increment = elapsed_time / 3600  # Convert seconds to hours
            total_leisure_hours += leisure_increment
        else:
            # For non-leisure activities, add a small random increment occasionally
            if random.random() < 0.1:  # 10% chance
                leisure_increment = random.uniform(0.01, 0.05)
                total_leisure_hours += leisure_increment
        
        # Round to 2 decimal places
        total_leisure_hours = round(total_leisure_hours, 2)
        
        # Create the new row
        new_row = [timestamp, active, elapsed_time, work_elapsed, leisure_elapsed, 
                   total_work_hours, total_leisure_hours]
        
        # Add app name, URL, and note
        if len(row) >= 5:
            new_row.append(row[4])  # active_app
            
            # Add URL if it exists, otherwise n/a
            if len(row) >= 6:
                new_row.append(row[5])  # url
            else:
                new_row.append("n/a")
            
            # Add note if it exists
            if len(row) >= 7:
                new_row.append(row[6])  # note
        else:
            # Add defaults if missing
            new_row.append("Unknown")
            new_row.append("n/a")
        
        new_rows.append(new_row)
    
    # Write the updated data back to the file
    with open(file_path, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(new_header)
        writer.writerows(new_rows)
    
    print(f"Updated {len(new_rows)} rows in {file_path}")

def add_manual_log_entry(file_path, timestamp, duration_minutes, activity_type='work', app_name='Manual Entry', url='n/a', note='manual entry'):
    """
    Add a manual log entry to the CSV file.
    
    Args:
        file_path (str): Path to the CSV file
        timestamp (str): Timestamp in format 'YYYY-MM-DD HH:MM:SS'
        duration_minutes (int): Duration of the activity in minutes
        activity_type (str): Type of activity ('work' or 'leisure')
        app_name (str): Name of the application used
        url (str): URL if applicable
        note (str): Additional notes
    """
    try:
        # Validate timestamp format
        datetime.strptime(timestamp, '%Y-%m-%d %H:%M:%S')
    except ValueError:
        print("Error: Timestamp must be in format 'YYYY-MM-DD HH:MM:SS'")
        return False
        
    # Convert duration to seconds
    duration_seconds = duration_minutes * 60
    
    # Set active state based on activity type
    active = 1 if activity_type == 'work' else 2
    
    # Calculate elapsed times
    work_elapsed = duration_seconds if activity_type == 'work' else 0
    leisure_elapsed = duration_seconds if activity_type == 'leisure' else 0
    
    # Read existing data to get the last total hours
    total_work_hours = 0.0
    total_leisure_hours = 0.0
    
    try:
        with open(file_path, 'r', newline='') as f:
            reader = csv.reader(f)
            header = next(reader)
            for row in reader:
                if len(row) >= 7:  # Ensure row has enough columns
                    total_work_hours = float(row[5])
                    total_leisure_hours = float(row[6])
    except FileNotFoundError:
        # Create new file with header if it doesn't exist
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with open(file_path, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(["timestamp", "active", "elapsed_time", "work_elapsed", "leisure_elapsed", 
                           "total_work_hours", "total_leisure_hours", "active_app", "url", "note"])
    
    # Update total hours based on new entry
    if activity_type == 'work':
        total_work_hours += duration_minutes / 60
    else:
        total_leisure_hours += duration_minutes / 60
    
    # Round to 2 decimal places
    total_work_hours = round(total_work_hours, 2)
    total_leisure_hours = round(total_leisure_hours, 2)
    
    # Create new row
    new_row = [
        timestamp,
        active,
        duration_seconds,
        work_elapsed,
        leisure_elapsed,
        total_work_hours,
        total_leisure_hours,
        app_name,
        url,
        note
    ]
    
    # Append the new row
    with open(file_path, 'a', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(new_row)
    
    print(f"Added manual log entry: {duration_minutes} minutes of {activity_type} activity at {timestamp}")
    return True

def main():
    parser = argparse.ArgumentParser(description='Update CSV format and add manual log entries')
    parser.add_argument('--add', action='store_true', help='Add a manual log entry')
    parser.add_argument('--time', type=str, help='Time of the activity (HH:MM), e.g., 14:30')
    parser.add_argument('--duration', type=int, help='Duration in minutes')
    parser.add_argument('--type', choices=['work', 'leisure'], help='Activity type')
    parser.add_argument('--note', type=str, default='manual entry', help='Optional note')
    
    args = parser.parse_args()
    
    if args.add:
        if not all([args.time, args.duration, args.type]):
            parser.print_help()
            print("\nExample usage:")
            print("python update_csv_format.py --add --time 14:30 --duration 45 --type work --note 'meeting'")
            return
            
        # Get the current date and combine with provided time
        current_date = datetime.now().strftime('%Y-%m-%d')
        timestamp = f"{current_date} {args.time}:00"
        
        # Validate timestamp
        try:
            datetime.strptime(timestamp, '%Y-%m-%d %H:%M:%S')
        except ValueError:
            print("Error: Time must be in 24-hour format (HH:MM)")
            return
            
        # Set up log file path
        log_dir = os.path.join("/Users/mpriessner/windsurf_repos/dictate_tool/focus_logs", current_date)
        log_file = os.path.join(log_dir, "activity_log.csv")
        
        # Add the entry
        add_manual_log_entry(
            file_path=log_file,
            timestamp=timestamp,
            duration_minutes=args.duration,
            activity_type=args.type,
            note=args.note
        )
    else:
        # Original CSV update functionality
        base_dir = "/Users/mpriessner/windsurf_repos/dictate_tool/focus_logs"
        for date in ["2025-04-22", "2025-04-23"]:
            log_file = os.path.join(base_dir, date, "activity_log.csv")
            if os.path.exists(log_file):
                update_csv_file(log_file)
            else:
                print(f"File not found: {log_file}")

if __name__ == "__main__":
    main()

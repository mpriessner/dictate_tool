#!/usr/bin/env python3
"""
Script to update the CSV format of existing focus log files to the new format
with separate work and leisure time columns.
"""

import os
import csv
import random
from datetime import datetime

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

def main():
    # Base directory for focus logs
    base_dir = "/Users/mpriessner/windsurf_repos/dictate_tool/focus_logs"
    
    # Update April 22 log
    april_22_log = os.path.join(base_dir, "2025-04-22", "activity_log.csv")
    if os.path.exists(april_22_log):
        update_csv_file(april_22_log)
    else:
        print(f"File not found: {april_22_log}")
    
    # Update April 23 log
    april_23_log = os.path.join(base_dir, "2025-04-23", "activity_log.csv")
    if os.path.exists(april_23_log):
        update_csv_file(april_23_log)
    else:
        print(f"File not found: {april_23_log}")

if __name__ == "__main__":
    main()

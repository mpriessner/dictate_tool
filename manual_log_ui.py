#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta
import os
import pandas as pd
import numpy as np

class ManualLogUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Add Manual Focus Log")
        
        # Define activity types and their properties
        self.activity_types = {
            'work': {'active': 1, 'work_elapsed': True, 'leisure_elapsed': False},
            'leisure': {'active': 2, 'work_elapsed': False, 'leisure_elapsed': True},
            'break': {'active': 0, 'work_elapsed': False, 'leisure_elapsed': False},
            'inactive': {'active': 0, 'work_elapsed': False, 'leisure_elapsed': False}
        }
        
        # Date and time frame
        time_frame = ttk.LabelFrame(root, text="Time Range", padding="5 5 5 5")
        time_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        
        # Date picker
        ttk.Label(time_frame, text="Date (YYYY-MM-DD):").grid(row=0, column=0, padx=5, pady=5)
        self.date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        self.date_entry = ttk.Entry(time_frame, textvariable=self.date_var)
        self.date_entry.grid(row=0, column=1, columnspan=3, padx=5, pady=5)
        
        # Start time
        ttk.Label(time_frame, text="Start (HH:MM):").grid(row=1, column=0, padx=5, pady=5)
        self.start_time_var = tk.StringVar(value=datetime.now().strftime('%H:%M'))
        self.start_time_entry = ttk.Entry(time_frame, textvariable=self.start_time_var, width=10)
        self.start_time_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # End time
        ttk.Label(time_frame, text="End (HH:MM):").grid(row=1, column=2, padx=5, pady=5)
        end_time = (datetime.now() + timedelta(minutes=30)).strftime('%H:%M')
        self.end_time_var = tk.StringVar(value=end_time)
        self.end_time_entry = ttk.Entry(time_frame, textvariable=self.end_time_var, width=10)
        self.end_time_entry.grid(row=1, column=3, padx=5, pady=5)
        
        # Activity type frame
        type_frame = ttk.LabelFrame(root, text="Activity Type", padding="5 5 5 5")
        type_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        
        self.activity_type = tk.StringVar(value="work")
        row = 0
        col = 0
        for activity in ['work', 'leisure', 'break', 'inactive']:
            ttk.Radiobutton(type_frame, text=activity.capitalize(), 
                           variable=self.activity_type,
                           value=activity).grid(row=row, column=col, padx=5, pady=5)
            col += 1
        
        # Note frame
        note_frame = ttk.LabelFrame(root, text="Note", padding="5 5 5 5")
        note_frame.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")
        
        self.note_var = tk.StringVar()
        self.note_entry = ttk.Entry(note_frame, textvariable=self.note_var)
        self.note_entry.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        
        # Add button
        self.add_button = ttk.Button(root, text="Add Log Entry", command=self.add_log)
        self.add_button.grid(row=4, column=0, pady=10)
        
        # Status label
        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(root, textvariable=self.status_var)
        self.status_label.grid(row=5, column=0, pady=5)
        
    def add_log(self):
        try:
            # Get and validate inputs
            date_str = self.date_var.get()
            start_time_str = self.start_time_var.get()
            end_time_str = self.end_time_var.get()
            activity_type = self.activity_type.get()
            note = self.note_var.get()
            
            # Create timestamps
            start_time = datetime.strptime(f"{date_str} {start_time_str}:00", '%Y-%m-%d %H:%M:%S')
            end_time = datetime.strptime(f"{date_str} {end_time_str}:00", '%Y-%m-%d %H:%M:%S')
            
            if end_time <= start_time:
                raise ValueError("End time must be after start time")
            
            # Check for existing entries
            log_file = os.path.join("/Users/mpriessner/windsurf_repos/dictate_tool/focus_logs", 
                                   date_str, "activity_log.csv")
            
            if os.path.exists(log_file):
                df = pd.read_csv(log_file)
                if not df.empty:
                    df['timestamp'] = pd.to_datetime(df['timestamp'])
                    overlapping = df[(df['timestamp'] >= start_time) & 
                                   (df['timestamp'] <= end_time)]
                    
                    if not overlapping.empty:
                        if not messagebox.askyesno("Confirm Overwrite", 
                            f"There are {len(overlapping)} existing entries in this time range. "
                            "Do you want to overwrite them?"):
                            return
            
            # Process the log file
            if self.process_log_file(start_time, end_time, activity_type, note):
                self.status_var.set("Log entries added successfully!")
            else:
                self.status_var.set("Error: Failed to add log entries")
            
        except ValueError as e:
            self.status_var.set(f"Error: {str(e)}")
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
    
    def process_log_file(self, start_time, end_time, activity_type, note):
        """Process the log file using pandas for better data handling"""
        date_str = start_time.strftime('%Y-%m-%d')
        log_dir = os.path.join("/Users/mpriessner/windsurf_repos/dictate_tool/focus_logs", date_str)
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, "activity_log.csv")
        
        # Define columns for the DataFrame
        columns = ["timestamp", "active", "elapsed_time", "work_elapsed", 
                  "leisure_elapsed", "total_work_hours", "total_leisure_hours", 
                  "active_app", "url", "note"]
        
        # Read existing file or create new DataFrame
        if os.path.exists(log_file):
            df = pd.read_csv(log_file)
        else:
            df = pd.DataFrame(columns=columns)
        
        # Convert timestamp to datetime
        if not df.empty:
            df['timestamp'] = pd.to_datetime(df['timestamp'])
        
        # Remove any existing entries in the time range
        df = df[~((df['timestamp'] >= start_time) & (df['timestamp'] <= end_time))]
        
        # Generate new entries
        new_entries = []
        current_time = start_time
        activity_props = self.activity_types[activity_type]
        
        while current_time <= end_time:
            elapsed_seconds = (current_time - start_time).total_seconds()
            
            # Calculate work/leisure elapsed time
            work_elapsed = elapsed_seconds if activity_props['work_elapsed'] else 0
            leisure_elapsed = elapsed_seconds if activity_props['leisure_elapsed'] else 0
            
            # Calculate total hours (will be updated later)
            entry = {
                'timestamp': current_time,
                'active': activity_props['active'],
                'elapsed_time': elapsed_seconds,
                'work_elapsed': work_elapsed,
                'leisure_elapsed': leisure_elapsed,
                'total_work_hours': 0.0,  # Placeholder
                'total_leisure_hours': 0.0,  # Placeholder
                'active_app': 'Manual Entry',
                'url': 'n/a',
                'note': note if current_time == start_time else ''
            }
            new_entries.append(entry)
            current_time += timedelta(minutes=1)
        
        # Add final inactive entry
        final_entry = {
            'timestamp': end_time,
            'active': 0,
            'elapsed_time': (end_time - start_time).total_seconds(),
            'work_elapsed': 0,
            'leisure_elapsed': 0,
            'total_work_hours': 0.0,  # Placeholder
            'total_leisure_hours': 0.0,  # Placeholder
            'active_app': 'Manual Entry',
            'url': 'n/a',
            'note': ''
        }
        new_entries.append(final_entry)
        
        # Add new entries to DataFrame
        new_df = pd.DataFrame(new_entries)
        df = pd.concat([df, new_df], ignore_index=True)
        
        # Sort by timestamp
        df = df.sort_values('timestamp')
        
        # Recalculate cumulative hours
        total_work = 0.0
        total_leisure = 0.0
        
        for idx, row in df.iterrows():
            if row['active'] == 1:  # Work
                total_work += 1/60  # Add one minute of work
            elif row['active'] == 2:  # Leisure
                total_leisure += 1/60  # Add one minute of leisure
            
            df.at[idx, 'total_work_hours'] = round(total_work, 2)
            df.at[idx, 'total_leisure_hours'] = round(total_leisure, 2)
        
        # Save the updated DataFrame
        df.to_csv(log_file, index=False)
        return True

def main():
    root = tk.Tk()
    
    # Set window to be topmost
    root.attributes('-topmost', True)
    root.lift()
    
    # Create the UI
    app = ManualLogUI(root)
    
    # Center the window
    # First update to get actual window size
    root.update_idletasks()
    
    # Get screen dimensions
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    # Get window dimensions
    window_width = root.winfo_width()
    window_height = root.winfo_height()
    
    # Calculate center position
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    
    # Set window position
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    # Start the main loop
    root.mainloop()

if __name__ == "__main__":
    main()

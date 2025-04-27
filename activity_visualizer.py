import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import numpy as np

def load_activity_data(file_path):
    # Read CSV file
    df = pd.read_csv(file_path, parse_dates=['timestamp'])
    return df

def process_intervals(df):
    intervals = []
    current_mode = None
    start_time = None
    last_timestamp = None
    INACTIVE_THRESHOLD = pd.Timedelta(minutes=30)  # 30 minutes threshold
    
    # Variables for tracking consecutive zeros
    zero_count = 0
    zero_start_time = None
    
    # Get the day's start time (midnight)
    day_start = df['timestamp'].iloc[0].replace(hour=0, minute=0, second=0, microsecond=0)
    
    # Find first activity (including breaks)
    first_activity = None
    first_activity_mode = None
    for idx, row in df.iterrows():
        if row['elapsed_time'] == 0:  # First entry in a monitoring session
            first_activity = row['timestamp']
            # Set the correct mode for the first activity
            if row['active'] == 0:
                first_activity_mode = 'break'
            elif row['active'] == 1:
                first_activity_mode = 'work'
            elif row['active'] == 2:
                first_activity_mode = 'leisure'
            break
    
    # If we found a first activity and it's after day start, add inactive period
    if first_activity and first_activity > day_start:
        # Add the inactive period before monitoring starts
        intervals.append({
            'start': day_start,
            'end': first_activity,
            'mode': 'inactive'
        })
        # Add the first activity with its correct mode
        if first_activity_mode:
            current_mode = first_activity_mode
            start_time = first_activity
    
    # Process the rest of the intervals
    for idx, row in df.iterrows():
        timestamp = row['timestamp']
        active = row['active']
        
        # Check for gaps between timestamps (program shutdown periods)
        if last_timestamp is not None:
            time_gap = timestamp - last_timestamp
            if time_gap > INACTIVE_THRESHOLD:
                # Add an inactive period for the gap
                intervals.append({
                    'start': last_timestamp,
                    'end': timestamp,
                    'mode': 'inactive'
                })
                # Reset the current mode to force starting a new interval
                current_mode = None
        
        # Track consecutive zeros
        if active == 0:
            if zero_count == 0:  # Start of a new zero sequence
                zero_start_time = timestamp
            zero_count += 1
        else:
            # If we had 5 or more zeros, mark that period as a break
            if zero_count >= 5 and zero_start_time:
                intervals.append({
                    'start': zero_start_time,
                    'end': timestamp,
                    'mode': 'break'
                })
                current_mode = None  # Reset current mode to start fresh
                start_time = timestamp  # Start new interval from here
            zero_count = 0
            zero_start_time = None

        # Determine the mode
        if active == 1:
            mode = 'work'
        elif active == 2:
            mode = 'leisure'
        elif active == 0:
            # If we're in a sequence of 5+ zeros, force this to be a break
            if zero_count >= 5:
                mode = 'break'
            else:
                mode = 'break'  # Still mark individual zeros as breaks
        else:
            mode = 'inactive'
            
        # If this is the first entry or mode changed
        if current_mode is None:
            start_time = timestamp
            current_mode = mode
        elif mode != current_mode:
            # Add the completed interval
            intervals.append({
                'start': start_time,
                'end': timestamp,
                'mode': current_mode
            })
            # Start new interval
            start_time = timestamp
            current_mode = mode
        
        last_timestamp = timestamp
    
    # Add the last interval
    if start_time is not None and current_mode is not None:
        intervals.append({
            'start': start_time,
            'end': timestamp,
            'mode': current_mode
        })
    
    return intervals

def plot_activity_timeline(intervals):
    # Set up the plot
    plt.figure(figsize=(15, 4))
    
    # Define colors (matching dashboard colors)
    colors = {
        'work': '#2196F3',    # Blue
        'leisure': '#FFCA28',  # Yellow/Amber
        'break': '#00BCD4',    # Cyan
        'inactive': '#9E9E9E'  # Gray
    }
    
    # Plot each interval as a horizontal bar
    for interval in intervals:
        start = interval['start']
        end = interval['end']
        mode = interval['mode']
        
        # Convert to hours for x-axis
        start_hours = start.hour + start.minute/60
        end_hours = end.hour + end.minute/60
        
        plt.hlines(y=0, xmin=start_hours, xmax=end_hours, 
                  colors=colors[mode], linewidth=20, label=mode)
    
    # Customize the plot
    plt.xlim(0, 24)
    plt.ylim(-0.5, 0.5)
    plt.xlabel('Hour of Day')
    plt.title('Activity Timeline')
    
    # Create legend without duplicates
    handles, labels = plt.gca().get_legend_handles_labels()
    by_label = dict(zip(labels, handles))
    plt.legend(by_label.values(), by_label.keys(), 
              loc='upper center', bbox_to_anchor=(0.5, -0.15),
              ncol=4)
    
    # Add grid lines for hours
    plt.grid(True, axis='x', alpha=0.3)
    plt.xticks(range(0, 25, 2))
    
    # Remove y-axis ticks and labels
    plt.yticks([])
    
    plt.tight_layout()
    return plt

def main():
    # Load today's activity log
    file_path = 'focus_logs/2025-04-27/activity_log.csv'
    df = load_activity_data(file_path)
    
    # Process the data into intervals
    intervals = process_intervals(df)
    
    # Create and show the plot
    plt = plot_activity_timeline(intervals)
    plt.show()
    
    # Print summary statistics
    modes = [interval['mode'] for interval in intervals]
    durations = [(interval['end'] - interval['start']).total_seconds() / 3600 
                 for interval in intervals]
    
    print("\nActivity Summary:")
    print("-----------------")
    for mode in ['work', 'leisure', 'break', 'inactive']:
        total_hours = sum(dur for m, dur in zip(modes, durations) if m == mode)
        print(f"{mode.capitalize()}: {total_hours:.2f} hours")

if __name__ == "__main__":
    main()

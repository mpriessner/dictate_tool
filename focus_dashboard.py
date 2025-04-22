import sys
import os
import csv
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (QApplication, QDialog, QVBoxLayout, QHBoxLayout, 
                            QLabel, QFrame, QWidget, QPushButton, QDateEdit, QScrollArea)
from PyQt5.QtCore import Qt, QDate, QRectF, QPoint, QSize
from PyQt5.QtGui import QFont, QColor, QPainter, QPen, QBrush, QPainterPath

# --- Custom Donut Chart Widget ---
class DonutChartWidget(QWidget):
    """A custom widget to draw an interactive donut chart."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(200, 200)
        self.data = {}  # Data format: {'Focus': seconds, 'Break': seconds}
        self.colors = {'Focus': QColor("#89B4FA"), 'Break': QColor("#74C7EC")} # Example colors
        self.hovered_segment = None # Track which segment is hovered
        self.segment_paths = {} # To store paths for hit testing

        # Enable mouse tracking for hover effects
        self.setMouseTracking(True)

    def setData(self, data):
        self.data = data
        self.segment_paths = {} # Reset paths when data changes
        self.hovered_segment = None
        self.setToolTip("") # Clear tooltip
        self.update() # Trigger repaint

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        rect = self.rect().adjusted(10, 10, -10, -10) # Add padding
        center = rect.center()
        outer_radius = min(rect.width(), rect.height()) / 2
        inner_radius = outer_radius * 0.6 # Adjust for donut thickness

        if not self.data or outer_radius <= 0:
            # Draw placeholder if no data or size is too small
            painter.setPen(QColor("#CDD6F4"))
            font = QFont("Arial", 10)
            painter.setFont(font)
            painter.drawText(self.rect(), Qt.AlignCenter, "No Data")
            return
            
        total_seconds = sum(self.data.values())
        if total_seconds <= 0:
            # Draw placeholder if total is zero
            painter.setPen(QColor("#CDD6F4"))
            font = QFont("Arial", 10)
            painter.setFont(font)
            painter.drawText(self.rect(), Qt.AlignCenter, "No Time Recorded")
            return

        start_angle = 90 * 16 # Start at the top (angles are in 1/16th of a degree)
        self.segment_paths.clear() # Clear paths before redrawing

        # Draw segments
        for i, (label, value) in enumerate(self.data.items()):
            angle = (value / total_seconds) * 360 * 16
            if angle <= 0: continue

            # Define bounding box for the arc
            bounding_rect = QRectF(center.x() - outer_radius, center.y() - outer_radius, 
                                   outer_radius * 2, outer_radius * 2)
            
            path = QPainterPath()
            path.moveTo(center)
            path.arcTo(bounding_rect, start_angle / 16.0, angle / 16.0)
            path.lineTo(center) # Ensure the path is closed for filling

            # Create the donut hole path
            hole_rect = QRectF(center.x() - inner_radius, center.y() - inner_radius,
                               inner_radius * 2, inner_radius * 2)
            hole_path = QPainterPath()
            hole_path.addEllipse(hole_rect)
            
            # Subtract the hole from the segment path
            segment_path = path.subtracted(hole_path)

            color = self.colors.get(label, QColor("gray"))
            
            # Highlight if hovered
            if self.hovered_segment == label:
                 # Make slightly brighter or change border
                 pen = QPen(Qt.white, 1)
                 painter.setPen(pen)
                 painter.setBrush(QBrush(color.lighter(115)))
            else:
                 painter.setPen(Qt.NoPen)
                 painter.setBrush(QBrush(color))

            painter.drawPath(segment_path)
            
            # Store path for hit testing (hover)
            self.segment_paths[label] = segment_path

            start_angle += angle

    def mouseMoveEvent(self, event):
        """Handle mouse hover events to detect which segment is hovered."""
        pos = event.pos()
        found_segment = None
        # Check segments in reverse order of drawing (last drawn is topmost)
        for label in reversed(list(self.segment_paths.keys())):
             path = self.segment_paths[label]
             if path.contains(pos):
                 found_segment = label
                 break
        
        if found_segment != self.hovered_segment:
            self.hovered_segment = found_segment
            # Update tooltip or trigger repaint for visual feedback
            if self.hovered_segment:
                value = self.data.get(self.hovered_segment, 0)
                hours, rem = divmod(value, 3600)
                minutes, _ = divmod(rem, 60)
                self.setToolTip(f"{self.hovered_segment}: {int(hours)} hr {int(minutes)} min")
            else:
                self.setToolTip("")
            self.update() # Repaint to show hover effect

    def leaveEvent(self, event):
        """Reset hover effect when mouse leaves the widget."""
        if self.hovered_segment is not None:
            self.hovered_segment = None
            self.setToolTip("")
            self.update()


# --- Custom Bar Widget for Categories ---
class CategoryBarWidget(QWidget):
    def __init__(self, label, value, max_value, color, parent=None):
        super().__init__(parent)
        self.label = label
        self.value = value
        self.max_value = max_value
        self.color = color
        self.setToolTip(f"{label}: {self.format_time(value)}")
        self.setFixedHeight(30) # Fixed height for each bar
        
    def format_time(self, seconds):
        hours, rem = divmod(seconds, 3600)
        minutes, _ = divmod(rem, 60)
        if hours >= 1:
            return f"{int(hours)}h {int(minutes)}m"
        else:
            return f"{int(minutes)}m"

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        rect = self.rect()
        bar_height = rect.height() - 10 # Add some padding
        bar_y = (rect.height() - bar_height) / 2

        # Calculate bar width based on percentage of max_value
        bar_width = 0
        if self.max_value > 0:
            bar_width = (self.value / self.max_value) * rect.width() * 0.7 # Use 70% of width for the bar
        
        # Draw the background track (optional)
        # painter.setPen(Qt.NoPen)
        # painter.setBrush(QColor("#45475A")) # Dark grey background
        # painter.drawRect(0, bar_y, rect.width() * 0.7, bar_height)

        # Draw the actual bar
        painter.setPen(Qt.NoPen)
        painter.setBrush(QBrush(self.color))
        painter.drawRoundedRect(QRectF(0, bar_y, bar_width, bar_height), 5, 5) # Rounded corners

        # Draw the label and value text
        painter.setPen(QColor("#CDD6F4")) # Light text color
        font = QFont("Arial", 9)
        painter.setFont(font)
        text_rect = QRectF(rect.width() * 0.72, 0, rect.width() * 0.28, rect.height())
        # Combine label and formatted time, handle long labels
        display_text = f"{self.label[:15]}{'...' if len(self.label) > 15 else ''} ({self.format_time(self.value)})"
        painter.drawText(text_rect, Qt.AlignLeft | Qt.AlignVCenter, display_text)


# --- Dashboard Window --- 
class DashboardWindow(QDialog):
    """A dialog window to display focus and productivity statistics."""
    def __init__(self, parent=None, log_dir=None):
        super().__init__(parent)
        self.log_dir = log_dir
        self.current_date = QDate.currentDate()
        
        self.setWindowTitle("Focus Timer Dashboard")
        self.setMinimumSize(800, 650)
        self.setStyleSheet("""
            QDialog {
                background-color: #1E1E2E; /* Dark background */
                color: #CDD6F4; /* Light text */
            }
            QLabel {
                color: #CDD6F4;
            }
            QFrame#statsFrame, QFrame#chartFrame, QFrame#categoriesFrame {
                background-color: #28283D; /* Slightly lighter background for frames */
                border-radius: 8px;
            }
            QPushButton {
                background-color: #3A3A5A;
                color: white;
                border: none;
                padding: 5px 10px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #4A4A6A;
            }
            QDateEdit {
                 background-color: #3A3A5A;
                 color: white;
                 border: 1px solid #555;
                 padding: 3px;
            }
            QDateEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 15px;
                border-left-width: 1px;
                border-left-color: darkgray;
                border-left-style: solid;
                border-top-right-radius: 3px;
                border-bottom-right-radius: 3px;
            }
        """)
        
        self.setup_ui()
        self.load_data_for_day(self.current_date)

    def setup_ui(self):
        """Sets up the user interface layout and widgets."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # --- Header with Date Selector and View Options --- 
        header_layout = QHBoxLayout()
        title = QLabel("Productivity Dashboard")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        header_layout.addWidget(title)
        header_layout.addStretch()
        
        # View selection buttons (Day/Week/Month)
        view_layout = QHBoxLayout()
        view_layout.setSpacing(2)
        view_layout.setContentsMargins(0, 0, 0, 0)
        
        self.view_mode = "day"  # Default view mode
        
        self.day_btn = QPushButton("Day")
        self.week_btn = QPushButton("Week")
        self.month_btn = QPushButton("Month")
        
        # Style the buttons
        button_style = """
            QPushButton {
                background-color: #313244;
                color: #CDD6F4;
                border: none;
                padding: 5px 10px;
                border-radius: 4px;
                min-width: 60px;
            }
            QPushButton:hover {
                background-color: #45475A;
            }
            QPushButton:checked, QPushButton:pressed {
                background-color: #89B4FA;
                color: #1E1E2E;
            }
        """
        
        self.day_btn.setStyleSheet(button_style)
        self.week_btn.setStyleSheet(button_style)
        self.month_btn.setStyleSheet(button_style)
        
        # Make the buttons checkable (like radio buttons)
        self.day_btn.setCheckable(True)
        self.week_btn.setCheckable(True)
        self.month_btn.setCheckable(True)
        self.day_btn.setChecked(True)  # Default to day view
        
        # Connect signals
        self.day_btn.clicked.connect(lambda: self.change_view_mode("day"))
        self.week_btn.clicked.connect(lambda: self.change_view_mode("week"))
        self.month_btn.clicked.connect(lambda: self.change_view_mode("month"))
        
        view_layout.addWidget(self.day_btn)
        view_layout.addWidget(self.week_btn)
        view_layout.addWidget(self.month_btn)
        
        header_layout.addLayout(view_layout)
        header_layout.addSpacing(15)
        
        # Date selector
        self.date_edit = QDateEdit(self.current_date)
        self.date_edit.setCalendarPopup(True)
        self.date_edit.dateChanged.connect(self.date_changed)
        self.date_label = QLabel("Date:")
        header_layout.addWidget(self.date_label)
        header_layout.addWidget(self.date_edit)
        main_layout.addLayout(header_layout)

        # --- Top Row: Stats and Pie Chart --- 
        top_row_layout = QHBoxLayout()
        top_row_layout.setSpacing(15)

        # Left Side: Summary Stats
        stats_frame = QFrame()
        stats_frame.setObjectName("statsFrame")
        stats_layout = QVBoxLayout(stats_frame)
        stats_layout.setContentsMargins(15, 15, 15, 15)
        stats_layout.setSpacing(10)

        stats_title = QLabel("Daily Summary")
        stats_title.setFont(QFont("Arial", 14, QFont.Bold))
        stats_layout.addWidget(stats_title)

        # Placeholder for stats widgets
        self.work_hours_label = QLabel("Work Hours: --")
        self.work_hours_label.setFont(QFont("Arial", 12))
        stats_layout.addWidget(self.work_hours_label)
        
        self.focus_percent_label = QLabel("Focus: --%")
        self.focus_percent_label.setFont(QFont("Arial", 12))
        stats_layout.addWidget(self.focus_percent_label)

        stats_layout.addStretch()
        top_row_layout.addWidget(stats_frame, 1) # Takes 1/3 of the space

        # Right Side: Breakdown Pie Chart
        chart_section = QFrame()
        chart_section.setFrameShape(QFrame.StyledPanel)
        chart_layout = QVBoxLayout(chart_section)
        chart_title = QLabel("Time Breakdown")
        chart_title.setFont(QFont("Arial", 14, QFont.Bold))
        chart_title.setAlignment(Qt.AlignCenter)
        chart_layout.addWidget(chart_title)

        self.pie_chart_widget = DonutChartWidget(self) # Use the custom widget
        chart_layout.addWidget(self.pie_chart_widget)

        top_row_layout.addWidget(chart_section, 2) # Takes 2/3 of the space
        main_layout.addLayout(top_row_layout)

        # --- Bottom Section: Top Categories --- 
        categories_frame = QFrame()
        categories_frame.setObjectName("categoriesFrame")
        categories_layout = QVBoxLayout(categories_frame)
        categories_layout.setContentsMargins(15, 15, 15, 15)
        categories_layout.setSpacing(10)

        categories_title = QLabel("Top Categories")
        categories_title.setFont(QFont("Arial", 14, QFont.Bold))
        categories_layout.addWidget(categories_title)

        # Placeholder layout for category bars
        self.categories_container_layout = QVBoxLayout()
        self.categories_container_layout.setSpacing(5) # Spacing between bars
        self.categories_container_layout.setAlignment(Qt.AlignTop) # Align bars to the top
        categories_scroll_content = QWidget()
        categories_scroll_content.setLayout(self.categories_container_layout)
        categories_scroll_area = QScrollArea()
        categories_scroll_area.setWidgetResizable(True)
        categories_scroll_area.setWidget(categories_scroll_content)
        categories_layout.addWidget(categories_scroll_area)
        categories_layout.addStretch()

        main_layout.addWidget(categories_frame)
        
    def change_view_mode(self, mode):
        """Handles switching between day, week, and month views."""
        if mode == self.view_mode:
            return  # No change needed
            
        # Update the mode
        self.view_mode = mode
        
        # Update the button states
        self.day_btn.setChecked(mode == "day")
        self.week_btn.setChecked(mode == "week")
        self.month_btn.setChecked(mode == "month")
        
        # Update the date edit label and behavior
        if mode == "day":
            self.date_edit.setDisplayFormat("yyyy-MM-dd")
            self.date_label.setText("Date:")
        elif mode == "week":
            # Format to show the week number
            self.date_edit.setDisplayFormat("yyyy-'W'ww")
            self.date_label.setText("Week:")
        elif mode == "month":
            # Format to show month and year
            self.date_edit.setDisplayFormat("MMM yyyy")
            self.date_label.setText("Month:")
        
        # Load data for the new view
        self.date_changed(self.date_edit.date())
        

    def clear_data_display(self):
        """Clears all data displays to prepare for loading new data."""
        # Clear category bars
        while self.categories_container_layout.count():
            item = self.categories_container_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
                
        # Clear chart data
        self.pie_chart_widget.setData({})
        
        # Reset labels
        self.work_hours_label.setText("Work Hours: --")
        self.focus_percent_label.setText("Focus: --%")
        
    def get_week_dates(self, date):
        """Returns a list of dates for the week containing the given date."""
        # Get the first day of the week (Monday)
        start_of_week = date.addDays(-(date.dayOfWeek() - 1))
        
        # Create a list of 7 days starting from Monday
        week_dates = [start_of_week.addDays(i) for i in range(7)]
        return week_dates
        
    def date_changed(self, date):
        """Handles the date change event from the QDateEdit."""
        self.current_date = date
        
        # Clear previous content
        self.clear_data_display()
        
        # Load data based on the current view mode
        if self.view_mode == "day":
            self.load_data_for_day(date)
        elif self.view_mode == "week":
            self.load_data_for_week(date)
        elif self.view_mode == "month":
            self.load_data_for_month(date)
        
    def load_data_for_day(self, date):
        """Loads and processes data for the selected date."""
        date_str = date.toString("yyyy-MM-dd")
        log_path = os.path.join(self.log_dir, date_str, 'activity_log.csv')
        print(f"Attempting to load log file: {log_path}")

        total_hours = 0
        active_seconds = 0
        inactive_seconds = 0
        total_seconds = 0
        app_times = {} # Will be used later for categories
        
        if not os.path.exists(log_path):
            print(f"Log file not found for {date_str}")
            self.work_hours_label.setText("Work Hours: 0 hr 0 min")
            self.focus_percent_label.setText("Focus: 0%")
            self.pie_chart_widget.setData({}) # Clear chart data
            # Add placeholder text if no data
            self.categories_container_layout.addWidget(QLabel("No data for selected date."))
            return
            
        try:
            with open(log_path, 'r', newline='') as f:
                reader = csv.reader(f)
                try:
                    header = next(reader) # Skip header
                except StopIteration:
                     print(f"Log file is empty for {date_str}")
                     # Handle empty file case same as non-existent
                     self.work_hours_label.setText("Work Hours: 0 hr 0 min")
                     self.focus_percent_label.setText("Focus: 0%")
                     self.pie_chart_widget.setData({}) # Clear chart data
                     # Add placeholder text if no data
                     self.categories_container_layout.addWidget(QLabel("No data for selected date."))
                     return

                entries = list(reader)
                if not entries:
                    print(f"Log file has no data entries for {date_str}")
                    self.work_hours_label.setText("Work Hours: 0 hr 0 min")
                    self.focus_percent_label.setText("Focus: 0%")
                    self.pie_chart_widget.setData({}) # Clear chart data
                    # Add placeholder text if no data
                    self.categories_container_layout.addWidget(QLabel("No data for selected date."))
                    return

                last_entry = entries[-1]
                if len(last_entry) > 3:
                    try:
                        total_hours = float(last_entry[3])
                    except (ValueError, TypeError):
                        total_hours = 0
                
                previous_timestamp = None
                for i, row in enumerate(entries):
                    if len(row) < 5:
                        continue # Skip malformed rows
                    
                    try:
                        timestamp = datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S')
                        active = int(row[1]) if row[1].isdigit() else 0
                        # Elapsed time (column 2) represents duration *since last log entry*
                        # Let's calculate duration between entries instead for accuracy
                        duration = 0
                        if previous_timestamp:
                            duration = (timestamp - previous_timestamp).total_seconds()
                        
                        app_name = row[4] if len(row) > 4 else "Unknown"

                        if duration > 0: # Only count if there's a valid duration
                            if active:
                                active_seconds += duration
                                # Add to app times (handle later)
                                if app_name not in app_times:
                                    app_times[app_name] = 0
                                app_times[app_name] += duration
                            else:
                                inactive_seconds += duration
                                # Track 'Break' time explicitly if app is 'Unknown' or similar during inactive periods
                                break_app_name = 'Break' if app_name in ['Unknown', ''] else f"Inactive ({app_name})"
                                if break_app_name not in app_times:
                                    app_times[break_app_name] = 0
                                app_times[break_app_name] += duration
                                
                            total_seconds += duration
                            
                        previous_timestamp = timestamp
                        
                    except (ValueError, TypeError, IndexError) as row_e:
                        print(f"Skipping malformed row {i+1}: {row} - Error: {row_e}")
                        continue
                        
        except FileNotFoundError:
             print(f"Log file not found: {log_path}")
             self.work_hours_label.setText("Work Hours: 0 hr 0 min")
             self.focus_percent_label.setText("Focus: 0%")
             self.pie_chart_widget.setData({}) # Clear chart data
             # Add placeholder text if no data
             self.categories_container_layout.addWidget(QLabel("No data for selected date."))
             return
        except Exception as e:
            print(f"Error loading data from {log_path}: {e}")
            # Optionally display an error message in the UI
            self.work_hours_label.setText("Work Hours: Error")
            self.focus_percent_label.setText("Focus: Error")
            self.pie_chart_widget.setData({}) # Clear chart data
            # Add placeholder text if no data
            self.categories_container_layout.addWidget(QLabel(f"Error loading data: {e}"))
            return
            
        # --- Update UI --- 
        # Update Work Hours Label (using total_hours from last log entry for consistency with timer)
        hours, rem = divmod(total_hours * 3600, 3600)
        minutes, _ = divmod(rem, 60)
        self.work_hours_label.setText(f"Work Hours: {int(hours)} hr {int(minutes)} min")
        
        # Update Focus Percentage Label
        if total_seconds > 0:
            focus_percentage = (active_seconds / total_seconds) * 100
            self.focus_percent_label.setText(f"Focus: {focus_percentage:.0f}%")
        else:
            self.focus_percent_label.setText("Focus: 0%")
            
        # Update Pie Chart
        chart_data = {'Focus': active_seconds, 'Break': inactive_seconds}
        self.pie_chart_widget.setData(chart_data)
        
        # Update categories
        self.update_category_display(app_times)
        
    def load_data_for_week(self, date):
        """Loads and processes aggregated data for the week containing the given date."""
        # Get all dates in the week
        week_dates = self.get_week_dates(date)
        
        # Set up variables for aggregation
        total_hours = 0
        active_seconds = 0
        inactive_seconds = 0
        total_seconds = 0
        app_times = {}  # Aggregate app usage across the week
        days_with_data = 0
        
        # Update title to show date range
        week_start = week_dates[0].toString("MMM d")
        week_end = week_dates[-1].toString("MMM d, yyyy")
        self.setWindowTitle(f"Focus Timer Dashboard: Week of {week_start} - {week_end}")
        
        for day_date in week_dates:
            day_str = day_date.toString("yyyy-MM-dd")
            log_path = os.path.join(self.log_dir, day_str, 'activity_log.csv')
            
            if not os.path.exists(log_path):
                print(f"No log file for {day_str}")
                continue
                
            try:
                with open(log_path, 'r', newline='') as f:
                    reader = csv.reader(f)
                    
                    try:
                        next(reader)  # Skip header
                    except StopIteration:
                        print(f"Empty log file for {day_str}")
                        continue
                    
                    # Process the day's data
                    entries = list(reader)
                    if not entries:
                        continue
                        
                    days_with_data += 1
                    
                    # Get the day's work hours from the last entry
                    last_entry = entries[-1]
                    if len(last_entry) > 3:
                        try:
                            day_hours = float(last_entry[3])
                            total_hours += day_hours
                        except (ValueError, TypeError):
                            pass
                    
                    # Calculate active/inactive time
                    day_active = 0
                    day_inactive = 0
                    previous_timestamp = None
                    
                    for row in entries:
                        if len(row) < 5:
                            continue
                            
                        try:
                            timestamp = datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S')
                            active = int(row[1]) if row[1].isdigit() else 0
                            app_name = row[4] if len(row) > 4 else "Unknown"
                            
                            # Calculate duration
                            duration = 0
                            if previous_timestamp:
                                duration = (timestamp - previous_timestamp).total_seconds()
                                
                            if duration > 0:
                                if active:
                                    active_seconds += duration
                                    day_active += duration
                                    if app_name not in app_times:
                                        app_times[app_name] = 0
                                    app_times[app_name] += duration
                                else:
                                    inactive_seconds += duration
                                    day_inactive += duration
                                    break_app_name = 'Break' if app_name in ['Unknown', ''] else f"Inactive ({app_name})"
                                    if break_app_name not in app_times:
                                        app_times[break_app_name] = 0
                                    app_times[break_app_name] += duration
                                
                                total_seconds += duration
                                
                            previous_timestamp = timestamp
                                
                        except (ValueError, TypeError, IndexError) as e:
                            continue
                            
            except Exception as e:
                print(f"Error loading data for {day_str}: {e}")
                continue
                
        # Update UI with aggregated data
        if days_with_data == 0:
            self.work_hours_label.setText("Work Hours: No data")
            self.focus_percent_label.setText("Focus: No data")
            self.pie_chart_widget.setData({})
            self.categories_container_layout.addWidget(QLabel("No data available for selected week"))
            return
            
        # Update Work Hours
        hours, rem = divmod(total_hours * 3600, 3600)
        minutes, _ = divmod(rem, 60)
        self.work_hours_label.setText(f"Work Hours: {int(hours)} hr {int(minutes)} min")
        
        # Update Focus Percentage
        if total_seconds > 0:
            focus_percentage = (active_seconds / total_seconds) * 100
            self.focus_percent_label.setText(f"Focus: {focus_percentage:.0f}%")
        else:
            self.focus_percent_label.setText("Focus: 0%")
            
        # Update pie chart
        chart_data = {'Focus': active_seconds, 'Break': inactive_seconds}
        self.pie_chart_widget.setData(chart_data)
        
        # Update categories
        self.update_category_display(app_times)
        
    def get_month_dates(self, date):
        """Returns a list of dates for the month containing the given date."""
        # Create a date for the first day of the month
        first_day = QDate(date.year(), date.month(), 1)
        
        # Determine how many days in the month
        days_in_month = 31  # Maximum possible
        if date.month() in [4, 6, 9, 11]:
            days_in_month = 30
        elif date.month() == 2:  # February
            if date.year() % 4 == 0 and (date.year() % 100 != 0 or date.year() % 400 == 0):
                days_in_month = 29  # Leap year
            else:
                days_in_month = 28
        
        # Create a list of all days in the month
        month_dates = [QDate(date.year(), date.month(), i+1) for i in range(days_in_month) 
                     if QDate(date.year(), date.month(), i+1).isValid()]
        return month_dates
    
    def load_data_for_month(self, date):
        """Loads and processes aggregated data for the month containing the given date."""
        # Get all dates in the month
        month_dates = self.get_month_dates(date)
        
        # Set up variables for aggregation
        total_hours = 0
        active_seconds = 0
        inactive_seconds = 0
        total_seconds = 0
        app_times = {}  # Aggregate app usage across the month
        days_with_data = 0
        
        # Update title to show month
        month_name = date.toString("MMMM yyyy")
        self.setWindowTitle(f"Focus Timer Dashboard: {month_name}")
        
        for day_date in month_dates:
            day_str = day_date.toString("yyyy-MM-dd")
            log_path = os.path.join(self.log_dir, day_str, 'activity_log.csv')
            
            if not os.path.exists(log_path):
                continue
                
            try:
                with open(log_path, 'r', newline='') as f:
                    reader = csv.reader(f)
                    
                    try:
                        next(reader)  # Skip header
                    except StopIteration:
                        continue
                    
                    # Process the day's data
                    entries = list(reader)
                    if not entries:
                        continue
                        
                    days_with_data += 1
                    
                    # Get the day's work hours from the last entry
                    last_entry = entries[-1]
                    if len(last_entry) > 3:
                        try:
                            day_hours = float(last_entry[3])
                            total_hours += day_hours
                        except (ValueError, TypeError):
                            pass
                    
                    # Calculate active/inactive time
                    previous_timestamp = None
                    
                    for row in entries:
                        if len(row) < 5:
                            continue
                            
                        try:
                            timestamp = datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S')
                            active = int(row[1]) if row[1].isdigit() else 0
                            app_name = row[4] if len(row) > 4 else "Unknown"
                            
                            # Calculate duration
                            duration = 0
                            if previous_timestamp:
                                duration = (timestamp - previous_timestamp).total_seconds()
                                
                            if duration > 0:
                                if active:
                                    active_seconds += duration
                                    if app_name not in app_times:
                                        app_times[app_name] = 0
                                    app_times[app_name] += duration
                                else:
                                    inactive_seconds += duration
                                    break_app_name = 'Break' if app_name in ['Unknown', ''] else f"Inactive ({app_name})"
                                    if break_app_name not in app_times:
                                        app_times[break_app_name] = 0
                                    app_times[break_app_name] += duration
                                
                                total_seconds += duration
                                
                            previous_timestamp = timestamp
                                
                        except (ValueError, TypeError, IndexError) as e:
                            continue
                            
            except Exception as e:
                print(f"Error loading data for {day_str}: {e}")
                continue
                
        # Update UI with aggregated data
        if days_with_data == 0:
            self.work_hours_label.setText("Work Hours: No data")
            self.focus_percent_label.setText("Focus: No data")
            self.pie_chart_widget.setData({})
            self.categories_container_layout.addWidget(QLabel("No data available for selected month"))
            return
            
        # Update Work Hours
        hours, rem = divmod(total_hours * 3600, 3600)
        minutes, _ = divmod(rem, 60)
        self.work_hours_label.setText(f"Work Hours: {int(hours)} hr {int(minutes)} min")
        
        # Add average daily hours
        avg_hours = total_hours / days_with_data
        avg_hours_int = int(avg_hours)
        avg_minutes = int((avg_hours - avg_hours_int) * 60)
        self.categories_container_layout.addWidget(QLabel(f"Data from {days_with_data} days • Avg: {avg_hours_int}h {avg_minutes}m per day"))
        
        # Update Focus Percentage
        if total_seconds > 0:
            focus_percentage = (active_seconds / total_seconds) * 100
            self.focus_percent_label.setText(f"Focus: {focus_percentage:.0f}%")
        else:
            self.focus_percent_label.setText("Focus: 0%")
            
        # Update pie chart
        chart_data = {'Focus': active_seconds, 'Break': inactive_seconds}
        self.pie_chart_widget.setData(chart_data)
        
        # Update categories
        self.update_category_display(app_times)
        
    def update_category_display(self, app_times):
        """Updates the category display with the provided app times data."""
        # Sort apps by time spent, descending
        sorted_apps = sorted(app_times.items(), key=lambda item: item[1], reverse=True)
        
        if not sorted_apps:
            self.categories_container_layout.addWidget(QLabel("No application data found."))
            print("No app data to display")
            return
            
        max_time = sorted_apps[0][1] if sorted_apps else 1 # Avoid division by zero
        if max_time <= 0: max_time = 1 # Ensure max_time is positive
        
        print(f"Displaying Top Categories (Max time: {max_time:.0f}s):")
        # Define a color palette (can be expanded)
        colors = [QColor("#F2CDCD"), QColor("#DDB6F2"), QColor("#F5C2E7"), 
                  QColor("#E8A2AF"), QColor("#F28FAD"), QColor("#ABE9B3"),
                  QColor("#FAE3B0"), QColor("#F8BD96"), QColor("#EE99A0"),
                  QColor("#89DCEB")]
        
        # Display top N categories (e.g., top 10 or all)
        for i, (app_name, time_spent) in enumerate(sorted_apps[:10]): # Limit to top 10
            if time_spent <= 0: continue # Skip apps with no time
            color = colors[i % len(colors)] # Cycle through colors
            bar_widget = CategoryBarWidget(app_name, time_spent, max_time, color)
            self.categories_container_layout.addWidget(bar_widget)
            print(f"  - {app_name}: {time_spent:.0f}s")
 
# Example usage (for testing)
if __name__ == '__main__':
    app = QApplication(sys.argv)
    # Provide a dummy log directory for testing
    script_dir = os.path.dirname(os.path.abspath(__file__))
    test_log_dir = os.path.join(script_dir, 'focus_logs')
    # Ensure the directory exists for testing
    os.makedirs(test_log_dir, exist_ok=True) 
    # Create a dummy log file for today for testing
    today_str = datetime.now().strftime("%Y-%m-%d")
    today_log_path = os.path.join(test_log_dir, today_str, 'activity_log.csv')
    os.makedirs(os.path.dirname(today_log_path), exist_ok=True)
    if not os.path.exists(today_log_path):
        with open(today_log_path, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Timestamp', 'Active', 'Elapsed', 'WorkHours', 'Application', 'URL', 'Note'])
            writer.writerow(['2024-01-01 10:00:00', '1', '600', '0.17', 'VS Code', 'n/a', '']) # 10 mins
            writer.writerow(['2024-01-01 10:10:00', '0', '300', '0.17', 'Break', 'n/a', ''])   # 5 mins
            writer.writerow(['2024-01-01 10:15:00', '1', '900', '0.42', 'Chrome', 'example.com', '']) # 15 mins

    dashboard = DashboardWindow(log_dir=test_log_dir)
    dashboard.show()
    sys.exit(app.exec_())

#!/usr/bin/env python3
# main.py - Application entry point for Focus Timer

import sys
import os
from PyQt5.QtWidgets import QApplication

# Import core components
from focus_timer_app.core.settings import Settings
from focus_timer_app.core.timer import FocusTimerCore
from focus_timer_app.core.activity import ActivityTracker
from focus_timer_app.core.application import ApplicationTracker

# Import data components
from focus_timer_app.data.logger import ActivityLogger
from focus_timer_app.data.analyzer import ActivityAnalyzer

# Import UI components
from focus_timer_app.ui.main_window import FocusTimerWindow

def main():
    """Main application entry point"""
    # Create application
    app = QApplication(sys.argv)
    app.setApplicationName("Focus Timer")
    
    # Initialize settings
    settings = Settings()
    
    # Initialize core components
    timer_core = FocusTimerCore(settings)
    activity_tracker = ActivityTracker()
    app_tracker = ApplicationTracker()
    
    # Initialize data components
    logger = ActivityLogger()
    analyzer = ActivityAnalyzer(logger)
    
    # Check if it's a new day and update settings if needed
    if settings.check_new_day():
        # Set up new logging directory
        logger.setup_logging()
    else:
        # Load work hours from logs if available
        work_hours, leisure_hours = analyzer.calculate_hours_from_log()
        if work_hours > 0 or leisure_hours > 0:
            timer_core.work_hours = work_hours
            timer_core.leisure_hours = leisure_hours
            settings.work_hours = work_hours
            settings.leisure_hours = leisure_hours
            settings.save()
    
    # Create main window
    window = FocusTimerWindow(timer_core, activity_tracker, app_tracker, logger, analyzer)
    window.show()
    
    # Start activity tracking
    activity_tracker.start()
    
    # Run application
    return app.exec_()

if __name__ == "__main__":
    sys.exit(main())

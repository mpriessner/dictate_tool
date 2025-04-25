#!/usr/bin/env python3
"""
simple_dashboard_adapter.py - Adapter to connect the simple dashboard to the main focus timer application
"""

import os
from pathlib import Path
from PyQt5.QtWidgets import QDialog

# Import the SimpleDashboardView from simple_dashboard.py
from simple_dashboard import SimpleDashboardView

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

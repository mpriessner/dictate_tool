#!/usr/bin/env python3
# settings_dialog.py - Settings dialog for Focus Timer

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog, QFormLayout, QHBoxLayout, QVBoxLayout, QWidget,
    QLabel, QPushButton, QSpinBox, QComboBox, QCheckBox
)

class SettingsDialog(QDialog):
    """Dialog for configuring Focus Timer settings"""
    
    def __init__(self, settings, parent=None):
        """Initialize the settings dialog
        
        Args:
            settings: Settings object
            parent: Parent widget
        """
        super().__init__(parent)
        self.parent = parent
        self.settings = settings
        self.setWindowTitle("Settings")
        self.setFixedSize(280, 250)  # Slightly narrower to ensure it fits on screen
        self.setStyleSheet("""
            QDialog {
                background-color: #252535;
                color: white;
            }
            QLabel {
                color: white;
            }
            QSpinBox, QTimeEdit, QComboBox {
                background-color: #333345;
                color: white;
                border: 1px solid #444;
                padding: 5px;
            }
            QPushButton {
                background-color: #3A3A4A;
                color: white;
                border: none;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #4A4A5A;
            }
            /* Individual button styles will override these */
            QCheckBox {
                color: white;
            }
            QCheckBox::indicator {
                width: 15px;
                height: 15px;
            }
        """)
        
        self.setup_ui()
    
    def setup_ui(self):
        """Set up the dialog UI"""
        layout = QFormLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Adjust label alignment to move fields to the left
        layout.setLabelAlignment(Qt.AlignLeft)
        layout.setFormAlignment(Qt.AlignLeft)
        
        # Set field growth policy to ensure fields don't expand too much
        layout.setFieldGrowthPolicy(QFormLayout.FieldsStayAtSizeHint)
        
        # Target hours setting
        self.target_hours_spin = QSpinBox()
        self.target_hours_spin.setRange(1, 24)
        self.target_hours_spin.setValue(self.settings.target_hours)
        self.target_hours_spin.setFixedWidth(60)  # Limit width
        layout.addRow("Target Hours:", self.target_hours_spin)
        
        # Auto-pause setting
        self.auto_pause_check = QCheckBox()
        self.auto_pause_check.setChecked(self.settings.auto_pause)
        layout.addRow("Auto-pause when inactive:", self.auto_pause_check)
        
        # Auto-pause minutes
        self.auto_pause_minutes_spin = QSpinBox()
        self.auto_pause_minutes_spin.setRange(1, 60)
        self.auto_pause_minutes_spin.setValue(self.settings.auto_pause_minutes)
        self.auto_pause_minutes_spin.setFixedWidth(60)  # Limit width
        layout.addRow("Auto-pause after (minutes):", self.auto_pause_minutes_spin)
        
        # Theme color
        self.theme_colors = {
            "Blue": "#00CCFF",
            "Purple": "#9370DB",
            "Green": "#4CAF50",
            "Orange": "#FF9800"
        }
        
        self.theme_color_combo = QComboBox()
        self.theme_color_combo.setFixedWidth(100)  # Limit width
        for name in self.theme_colors:
            self.theme_color_combo.addItem(name)
        
        # Set current color
        current_color = self.settings.theme_color
        for i, (name, color) in enumerate(self.theme_colors.items()):
            if color == current_color:
                self.theme_color_combo.setCurrentIndex(i)
                break
        
        layout.addRow("Theme Color:", self.theme_color_combo)
        
        # --- Button Layout Setup ---
        # Create a container widget for the buttons
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container) # Set the layout on the container
        button_layout.setContentsMargins(0, 10, 0, 0) # Add some top margin for spacing
        button_layout.setSpacing(10)
        
        # Add buttons
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setFixedSize(80, 20)  # Smaller size
        self.cancel_button.setStyleSheet("""
            QPushButton {
                background-color: #444455;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 5px 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #555566;
            }
        """)
        self.cancel_button.clicked.connect(self.reject)
        
        self.save_button = QPushButton("Save")
        self.save_button.setFixedSize(80, 20)  # Smaller size
        self.save_button.setStyleSheet("""
            QPushButton {
                background-color: #00AADD;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 5px 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #22CCFF;
            }
        """)
        self.save_button.clicked.connect(self.accept)
        
        # Add buttons with spacing between them
        button_layout.addWidget(self.cancel_button)
        button_layout.addWidget(self.save_button)
        
        # Add stretch only on the right side to push buttons to the left
        button_layout.addStretch(1)
        
        # Add the container widget to the form layout, spanning both columns
        layout.addRow(button_container)
    
    def get_settings(self):
        """Get the updated settings from the dialog
        
        Returns:
            dict: Updated settings
        """
        # Get the selected color
        color_name = self.theme_color_combo.currentText()
        color_value = self.theme_colors.get(color_name, "#00CCFF")
        
        return {
            "target_hours": self.target_hours_spin.value(),
            "auto_pause": self.auto_pause_check.isChecked(),
            "auto_pause_minutes": self.auto_pause_minutes_spin.value(),
            "theme_color": color_value
        }

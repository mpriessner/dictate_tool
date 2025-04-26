#!/usr/bin/env python3
# styles.py - UI styles and themes for Focus Timer

class FocusTimerStyles:
    """Styles and themes for the Focus Timer UI"""
    
    # Focus mode colors
    WORK_COLOR = "#00CCFF"  # Blue
    WORK_HOURS_COLOR = "#9370DB"  # Purple
    LEISURE_COLOR = "#FFCA28"  # Amber/yellow
    LEISURE_HOURS_COLOR = "#4CAF50"  # Green
    
    @staticmethod
    def get_main_window_style():
        """Get the style for the main window
        
        Returns:
            str: CSS style string
        """
        return """
            QMainWindow { 
                background-color: #1E1E2A; 
                color: white;
                font-family: Arial;
                border-radius: 15px;
                border: 1px solid #333;
            }
            QWidget#centralWidget { 
                background-color: #1E1E2A; 
                color: white;
                font-family: Arial;
                border-radius: 15px;
            }
            QPushButton {
                background-color: transparent;
                border: none;
                color: white;
            }
            QPushButton:hover {
                background-color: #2A2A3A;
            }
            QLabel {
                color: white;
            }
        """
    
    @staticmethod
    def get_menu_style():
        """Get the style for the context menu
        
        Returns:
            str: CSS style string
        """
        return """
            QMenu {
                background-color: #252535;
                color: white;
                border: 1px solid #333;
                padding: 5px;
            }
            QMenu::item {
                padding: 5px 20px;
            }
            QMenu::item:selected {
                background-color: #3A3A4A;
            }
        """
    
    @staticmethod
    def get_progress_color(percent):
        """Get a color for the progress indicator based on percentage
        
        Args:
            percent (float): Progress percentage
            
        Returns:
            str: Hex color string
        """
        # Create a smooth color gradient from red to green based on progress
        # Start with red (255,0,0) and transition to green (0,255,0)
        if percent <= 50:
            # Red to Yellow: Increase green component
            red = 255
            green = int((percent / 50) * 255)
            blue = 0
        else:
            # Yellow to Green: Decrease red component
            red = int(255 - ((percent - 50) / 50) * 255)
            green = 255
            blue = 0
            
        # If over 100%, add a blue component for a special color
        if percent > 100:
            blue = min(255, int((percent - 100) * 2.55))
            
        # Ensure values are within valid range
        red = max(0, min(255, red))
        green = max(0, min(255, green))
        blue = max(0, min(255, blue))
        
        # Return the color as a hex string
        return f"#{red:02x}{green:02x}{blue:02x}"

#!/usr/bin/env python3
# focus_circle.py - Focus circle widget for Focus Timer

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QPainter
from PyQt5.QtWidgets import QWidget, QMainWindow

class FocusCircle(QWidget):
    """Circular widget that indicates focus state and allows mode switching"""
    
    def __init__(self, color="#00CCFF", parent=None):
        """Initialize the focus circle
        
        Args:
            color (str): Color of the circle in hex format
            parent: Parent widget
        """
        super().__init__(parent)
        self.color = QColor(color)
        self.active = False
        self.setCursor(Qt.PointingHandCursor)  # Change cursor to indicate clickable
        self.setToolTip("Click to switch Work ⇄ Leisure mode\nDouble-click to minimize/maximize")
    
    def set_active(self, active):
        """Set the active state of the circle
        
        Args:
            active (bool): Whether the circle is active
        """
        self.active = active
        self.update()
    
    def set_color(self, color):
        """Set the color of the circle
        
        Args:
            color (str): Color in hex format
        """
        self.color = QColor(color)
        self.update()
    
    def paintEvent(self, event):
        """Paint the circle
        
        Args:
            event: Paint event
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Draw circle
        painter.setPen(Qt.NoPen)
        
        if self.active:
            # Filled circle when active
            painter.setBrush(self.color)
        else:
            # Hollow circle when inactive
            painter.setPen(self.color)
            painter.setBrush(Qt.NoBrush)
        
        painter.drawEllipse(1, 1, self.width() - 2, self.height() - 2)
    
    def mousePressEvent(self, event):
        """Handle mouse click on the focus circle to toggle focus mode
        
        Args:
            event: Mouse event
        """
        if event.button() == Qt.LeftButton:
            # Get the parent FocusTimer instance
            parent = self.parent()
            while parent and not isinstance(parent, QMainWindow):
                parent = parent.parent()
                
            # Toggle focus mode if parent is found
            if parent and hasattr(parent, "toggle_focus_mode"):
                parent.toggle_focus_mode()
                    
            event.accept()
    
    def mouseDoubleClickEvent(self, event):
        """Handle double-click on the focus circle to toggle UI state
        
        Args:
            event: Mouse event
        """
        if event.button() == Qt.LeftButton:
            # Get the parent FocusTimer instance
            parent = self.parent()
            while parent and not isinstance(parent, QMainWindow):
                parent = parent.parent()
                
            # Toggle UI state if parent is found
            if parent:
                if hasattr(parent, "ui_minimized") and parent.ui_minimized:
                    if hasattr(parent, "restore_ui"):
                        parent.restore_ui()
                else:
                    if hasattr(parent, "minimize_ui"):
                        parent.minimize_ui()
                    
            event.accept()

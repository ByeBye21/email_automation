"""
Email Automation Application
Version 1.0.0

A comprehensive email automation tool with CSV recipients, template personalization,
SMTP sending, and professional GUI interface.

Author: ByeBye21
Email: younes0079@gmail.com
GitHub: https://github.com/ByeBye21
Copyright © 2025 ByeBye21. All rights reserved.

Date: July 10, 2025
"""

import sys
import os
import json
import traceback
from pathlib import Path
from typing import Dict, List, Any, Optional
from datetime import datetime
import re

# Set High DPI attributes BEFORE importing QtWidgets
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QPropertyAnimation, QEasingCurve, QRect
from PyQt5.QtWidgets import QApplication

# Set High DPI scaling before any Qt widgets are created
if hasattr(Qt, 'AA_EnableHighDpiScaling'):
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

# Now import PyQt5 widgets
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QGridLayout, QLabel, QLineEdit, QPushButton, QTextEdit, 
    QComboBox, QSpinBox, QCheckBox, QFileDialog, QMessageBox,
    QProgressBar, QGroupBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QSplitter, QMenuBar, QAction, QStatusBar,
    QDialog, QDialogButtonBox, QFormLayout, QScrollArea, QFrame,
    QStackedWidget, QListWidget, QListWidgetItem, QCompleter, 
    QSizePolicy, QMenu, QWidgetAction
)

from PyQt5.QtGui import QMovie, QTextCursor
from PyQt5.QtGui import QFont, QIcon, QPixmap, QTextCursor as QTextCursor2, QPalette, QPainter, QColor, QDesktopServices
from PyQt5.QtCore import QStringListModel, QUrl

# Import our email application backend
from email_app import (
    EmailApplication, ConfigurationManager, Logger,
    EmailAppError, ConfigurationError, EmailSendError,
    CSVError, TemplateError
)


# Enhanced Modern CSS Styling with Modern Dropdowns
MODERN_STYLESHEET = """
/* Main Application Styling */
QMainWindow {
    background-color: #f8fafc;
    color: #1e293b;
}

/* Sidebar Styling */
QFrame#sidebar {
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                                stop: 0 #1e293b, stop: 1 #334155);
    border: none;
    min-width: 280px;
    max-width: 280px;
}

QLabel#appHeader {
    color: white;
    font-size: 22px;
    font-weight: bold;
    padding: 20px;
    border-bottom: 1px solid #475569;
}

/* Sidebar Navigation */
QListWidget#navList {
    background: transparent;
    border: none;
    outline: none;
    padding: 10px 0;
}

QListWidget#navList::item {
    background: transparent;
    color: #cbd5e1;
    padding: 15px 20px;
    margin: 2px 10px;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 500;
}

QListWidget#navList::item:hover {
    background-color: rgba(255, 255, 255, 0.1);
    color: white;
}

QListWidget#navList::item:selected {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                               stop: 0 #3b82f6, stop: 1 #1d4ed8);
    color: white;
    border-left: 4px solid #60a5fa;
}

/* Top Bar */
QFrame#topBar {
    background: white;
    border-bottom: 1px solid #e2e8f0;
    min-height: 70px;
    max-height: 70px;
}

QLabel#pageTitle {
    font-size: 24px;
    font-weight: bold;
    color: #1e293b;
    padding: 20px;
}

/* Card Styling */
QFrame.card {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    margin: 10px;
    padding: 0;
}

QLabel.cardTitle {
    font-size: 18px;
    font-weight: bold;
    color: #1e293b;
    padding: 20px 20px 10px 20px;
    border-bottom: 1px solid #f1f5f9;
}

QLabel.cardSubtitle {
    font-size: 14px;
    color: #64748b;
    padding: 0 20px 15px 20px;
}

/* Form Elements */
QLineEdit, QTextEdit, QSpinBox {
    border: 2px solid #e2e8f0;
    border-radius: 8px;
    padding: 12px;
    font-size: 14px;
    background: white;
    selection-background-color: #3b82f6;
}

QLineEdit:focus, QTextEdit:focus, QSpinBox:focus {
    border-color: #3b82f6;
    outline: none;
}

QLineEdit:hover, QTextEdit:hover, QSpinBox:hover {
    border-color: #94a3b8;
}

/* Modern ComboBox Styling */
QComboBox {
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                               stop: 0 #ffffff, stop: 1 #f8fafc);
    border: 2px solid #e2e8f0;
    border-radius: 8px;
    padding: 12px 15px;
    font-size: 14px;
    font-weight: 500;
    color: #374151;
    selection-background-color: #3b82f6;
    min-height: 20px;
}

QComboBox:hover {
    border-color: #3b82f6;
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                               stop: 0 #fefefe, stop: 1 #f0f9ff);
}

QComboBox:focus {
    border-color: #2563eb;
    background: white;
}

QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 30px;
    border-left: 2px solid #e2e8f0;
    border-top-right-radius: 6px;
    border-bottom-right-radius: 6px;
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                               stop: 0 #f8fafc, stop: 1 #e2e8f0);
}

QComboBox::down-arrow {
    image: none;
    border: 2px solid #6b7280;
    border-top: none;
    border-left: none;
    width: 8px;
    height: 8px;
    transform: rotate(45deg);
    margin-right: 8px;
}

QComboBox QAbstractItemView {
    background-color: white;
    border: 2px solid #3b82f6;
    border-radius: 8px;
    selection-background-color: #eff6ff;
    selection-color: #1e40af;
    outline: none;
    padding: 5px;
}

QComboBox QAbstractItemView::item {
    height: 35px;
    padding: 8px 12px;
    border-radius: 6px;
    margin: 2px;
}

QComboBox QAbstractItemView::item:hover {
    background-color: #f0f9ff;
    color: #1e40af;
}

QComboBox QAbstractItemView::item:selected {
    background-color: #3b82f6;
    color: white;
}

/* Loading Button Styling */
QPushButton.loading {
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                               stop: 0 #6b7280, stop: 1 #4b5563) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 12px 24px !important;
    font-size: 14px !important;
    font-weight: bold !important;
}

/* Primary Buttons */
QPushButton.primary {
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                               stop: 0 #3b82f6, stop: 1 #1d4ed8) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 12px 24px !important;
    font-size: 14px !important;
    font-weight: bold !important;
    min-height: 20px;
}

QPushButton.primary:hover {
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                               stop: 0 #2563eb, stop: 1 #1e40af) !important;
}

QPushButton.primary:pressed {
    background: #1e40af !important;
}

QPushButton.primary:disabled {
    background: #9ca3af !important;
    color: #ffffff !important;
}

/* Secondary Buttons */
QPushButton.secondary {
    background: transparent !important;
    color: #475569 !important;
    border: 2px solid #e2e8f0 !important;
    border-radius: 8px !important;
    padding: 10px 20px !important;
    font-size: 12px !important;
    font-weight: 500 !important;
}

QPushButton.secondary:hover {
    border-color: #94a3b8 !important;
    background: #f8fafc !important;
}

QPushButton.secondary:disabled {
    background: #f8fafc !important;
    color: #9ca3af !important;
    border-color: #e5e7eb !important;
}

/* Success Buttons */
QPushButton.success {
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                               stop: 0 #10b981, stop: 1 #047857) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 12px 24px !important;
    font-size: 14px !important;
    font-weight: bold !important;
}

QPushButton.success:hover {
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                               stop: 0 #059669, stop: 1 #065f46) !important;
}

QPushButton.success:disabled {
    background: #9ca3af !important;
    color: #ffffff !important;
}

/* Warning Buttons */
QPushButton.warning {
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                               stop: 0 #f59e0b, stop: 1 #d97706) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 12px 24px !important;
    font-size: 14px !important;
    font-weight: bold !important;
}

QPushButton.warning:hover {
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                               stop: 0 #d97706, stop: 1 #b45309) !important;
}

/* Danger Buttons */
QPushButton.danger {
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                               stop: 0 #ef4444, stop: 1 #dc2626) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 12px 24px !important;
    font-size: 14px !important;
    font-weight: bold !important;
}

QPushButton.danger:hover {
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                               stop: 0 #dc2626, stop: 1 #b91c1c) !important;
}

/* Large Campaign Buttons */
QPushButton.campaign-btn {
    border: none !important;
    border-radius: 12px !important;
    padding: 20px 30px !important;
    font-size: 16px !important;
    font-weight: bold !important;
    min-height: 60px !important;
    min-width: 200px !important;
}

QPushButton.campaign-primary {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,
                               stop: 0 #3b82f6, stop: 1 #1d4ed8) !important;
    color: white !important;
    box-shadow: 0 4px 14px 0 rgba(59, 130, 246, 0.39);
}

QPushButton.campaign-primary:hover {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,
                               stop: 0 #2563eb, stop: 1 #1e40af) !important;
    transform: translateY(-2px);
}

QPushButton.campaign-success {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,
                               stop: 0 #10b981, stop: 1 #047857) !important;
    color: white !important;
    box-shadow: 0 4px 14px 0 rgba(16, 185, 129, 0.39);
}

QPushButton.campaign-success:hover {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,
                               stop: 0 #059669, stop: 1 #065f46) !important;
}

QPushButton.campaign-danger {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,
                               stop: 0 #ef4444, stop: 1 #dc2626) !important;
    color: white !important;
    box-shadow: 0 4px 14px 0 rgba(239, 68, 68, 0.39);
}

QPushButton.campaign-danger:hover {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,
                               stop: 0 #dc2626, stop: 1 #b91c1c) !important;
}

/* Connection Status */
QLabel.connection-status {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                               stop: 0 #ef4444, stop: 1 #dc2626);
    color: white;
    border-radius: 6px;
    padding: 6px 12px;
    font-size: 12px;
    font-weight: bold;
    margin: 2px;
    border: 1px solid #dc2626;
}

QLabel.connection-status.connected {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                               stop: 0 #10b981, stop: 1 #047857);
    border-color: #047857;
}

QLabel.connection-status.disconnected {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                               stop: 0 #64748b, stop: 1 #475569);
    border-color: #475569;
}

/* Checkboxes */
QCheckBox {
    font-size: 14px;
    color: #374151;
    spacing: 8px;
}

QCheckBox::indicator {
    width: 18px;
    height: 18px;
    border: 2px solid #d1d5db;
    border-radius: 4px;
    background: white;
}

QCheckBox::indicator:checked {
    background: #3b82f6;
    border-color: #3b82f6;
}

/* Tables */
QTableWidget {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    gridline-color: #f1f5f9;
    font-size: 13px;
}

QTableWidget::item {
    padding: 12px;
    border-bottom: 1px solid #f1f5f9;
}

QTableWidget::item:selected {
    background: #eff6ff;
    color: #1e40af;
}

QHeaderView::section {
    background: #f8fafc;
    color: #374151;
    font-weight: bold;
    font-size: 12px;
    padding: 12px;
    border: none;
    border-bottom: 2px solid #e2e8f0;
}

/* Modern Progress Bar */
QProgressBar {
    border: none;
    border-radius: 10px;
    background: #e2e8f0;
    height: 20px;
    text-align: center;
    font-weight: bold;
    color: #1e293b;
}

QProgressBar::chunk {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                               stop: 0 #3b82f6, stop: 1 #1d4ed8);
    border-radius: 10px;
}

/* Dropdown Menu for Attributes */
QMenu {
    background-color: white;
    border: 2px solid #3b82f6;
    border-radius: 10px;
    padding: 8px;
    box-shadow: 0 10px 25px rgba(0, 0, 0, 0.15);
}

QMenu::item {
    background-color: transparent;
    padding: 10px 15px;
    border-radius: 6px;
    color: #374151;
    font-size: 13px;
    font-weight: 500;
    margin: 2px;
}

QMenu::item:selected {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                               stop: 0 #3b82f6, stop: 1 #2563eb);
    color: white;
}

QMenu::item:hover {
    background: #f0f9ff;
    color: #1e40af;
}

/* Campaign Stats Cards */
QFrame.stats-card {
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                               stop: 0 #ffffff, stop: 1 #f8fafc);
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 20px;
    margin: 10px;
}

QFrame.stats-card:hover {
    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                               stop: 0 #fefefe, stop: 1 #f0f9ff);
    border-color: #3b82f6;
}

/* LOG DISPLAY - WHITE FONT */
QTextEdit.log-display {
    background-color: #1e293b !important;
    color: #ffffff !important;
    border: 1px solid #334155 !important;
    border-radius: 8px !important;
    padding: 15px !important;
    font-family: 'Consolas', 'Monaco', monospace !important;
    font-size: 11px !important;
    line-height: 1.4 !important;
}
"""


class LoadingButton(QPushButton):
    """Button with loading animation capability."""
    
    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.original_text = text
        self.original_class = ""
        self.is_loading = False
        self.loading_timer = QTimer()
        self.loading_timer.timeout.connect(self.update_loading_text)
        self.loading_dots = 0
    
    def start_loading(self, loading_text="Loading"):
        """Start loading animation."""
        if not self.is_loading:
            self.is_loading = True
            self.original_class = self.property("class") or ""
            self.setEnabled(False)
            self.setProperty("class", f"{self.original_class} loading")
            self.style().unpolish(self)
            self.style().polish(self)
            self.loading_text = loading_text
            self.loading_dots = 0
            self.loading_timer.start(500)
            self.update_loading_text()
    
    def stop_loading(self):
        """Stop loading animation."""
        if self.is_loading:
            self.is_loading = False
            self.loading_timer.stop()
            self.setText(self.original_text)
            self.setEnabled(True)
            self.setProperty("class", self.original_class)
            self.style().unpolish(self)
            self.style().polish(self)
    
    def update_loading_text(self):
        """Update loading text with animated dots."""
        if self.is_loading:
            dots = "." * (self.loading_dots % 4)
            self.setText(f"{self.loading_text}{dots}")
            self.loading_dots += 1


class ModernAttributeDropdown(QWidget):
    """Modern dropdown widget for inserting CSV attributes."""
    
    def __init__(self, text_widget, csv_columns=None, parent=None):
        super().__init__(parent)
        self.text_widget = text_widget
        self.csv_columns = csv_columns or []
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the modern attribute dropdown widget."""
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        self.setLayout(layout)
        
        # Modern label
        label = QLabel("Insert attributes:")
        label.setStyleSheet("""
            font-size: 14px; 
            color: #6b7280; 
            font-weight: 500;
            padding: 5px 0;
        """)
        layout.addWidget(label)
        
        # Modern dropdown button with gradient
        self.dropdown_btn = QPushButton("📋 Select Attribute ▼")
        self.dropdown_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                                           stop: 0 #f8fafc, stop: 1 #e2e8f0);
                border: 2px solid #d1d5db;
                border-radius: 8px;
                padding: 8px 16px;
                font-size: 12px;
                font-weight: 600;
                color: #374151;
                min-width: 140px;
            }
            QPushButton:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                                           stop: 0 #eff6ff, stop: 1 #dbeafe);
                border-color: #3b82f6;
                color: #1e40af;
            }
            QPushButton:pressed {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                                           stop: 0 #dbeafe, stop: 1 #bfdbfe);
            }
            QPushButton:disabled {
                background: #f1f5f9;
                color: #9ca3af;
                border-color: #e5e7eb;
            }
        """)
        self.dropdown_btn.clicked.connect(self.show_attributes_menu)
        layout.addWidget(self.dropdown_btn)
        
        layout.addStretch()
        self.update_attributes(self.csv_columns)
    
    def update_attributes(self, csv_columns):
        """Update available attributes based on CSV columns."""
        self.csv_columns = csv_columns
        
        if csv_columns:
            self.dropdown_btn.setText("📋 Select Attribute ▼")
            self.dropdown_btn.setEnabled(True)
        else:
            self.dropdown_btn.setText("📋 No Attributes Available")
            self.dropdown_btn.setEnabled(False)
    
    def show_attributes_menu(self):
        """Show modern dropdown menu with available attributes."""
        if not self.csv_columns:
            return
        
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                                           stop: 0 #ffffff, stop: 1 #f8fafc);
                border: 2px solid #3b82f6;
                border-radius: 12px;
                padding: 10px;
                box-shadow: 0 10px 25px rgba(0, 0, 0, 0.15);
            }
            QMenu::item {
                background-color: transparent;
                padding: 12px 18px;
                border-radius: 8px;
                color: #374151;
                font-size: 13px;
                font-weight: 500;
                margin: 2px;
                min-width: 120px;
            }
            QMenu::item:selected {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                                           stop: 0 #3b82f6, stop: 1 #2563eb);
                color: white;
            }
            QMenu::item:hover {
                background: #f0f9ff;
                color: #1e40af;
            }
        """)
        
        # Create a non-interactive label
        label = QLabel("🎯 Available Attributes")
        label.setStyleSheet("""
            QLabel {
                padding: 6px 12px;
                font-size: 16px;
                font-weight: bold;
                color: gray;
            }
        """)
        label.setEnabled(False)

        # Wrap it in a QWidgetAction
        header_action = QWidgetAction(menu)
        header_action.setDefaultWidget(label)

        # Add it to the menu
        menu.addAction(header_action)
        menu.addSeparator()
        
        # Add attributes
        for column in self.csv_columns:
            action = menu.addAction(f"📌 {{{{ {column} }}}}")
            action.triggered.connect(lambda checked, col=column: self.insert_attribute(col))
        
        # Show menu below the button
        button_rect = self.dropdown_btn.geometry()
        menu_pos = self.dropdown_btn.mapToGlobal(button_rect.bottomLeft())
        menu.exec_(menu_pos)
    
    def insert_attribute(self, column):
        """Insert selected attribute into text widget."""
        attribute_text = f"{{{{{column}}}}}"
        
        if isinstance(self.text_widget, QLineEdit):
            current_text = self.text_widget.text()
            cursor_pos = self.text_widget.cursorPosition()
            new_text = current_text[:cursor_pos] + attribute_text + current_text[cursor_pos:]
            self.text_widget.setText(new_text)
            self.text_widget.setCursorPosition(cursor_pos + len(attribute_text))
        elif isinstance(self.text_widget, QTextEdit):
            cursor = self.text_widget.textCursor()
            cursor.insertText(attribute_text)
        
        self.text_widget.setFocus()


class EmailColumnSelector(QWidget):
    """Modern widget to select email column from CSV."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.csv_columns = []
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the modern email column selector widget."""
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        self.setLayout(layout)
        
        label = QLabel("Email Column:")
        label.setStyleSheet("""
            font-size: 14px; 
            font-weight: bold; 
            color: #374151;
            padding: 5px 0;
        """)
        layout.addWidget(label)
        
        self.email_column_combo = QComboBox()
        self.email_column_combo.setMinimumHeight(40)
        self.email_column_combo.setMinimumWidth(150)
        layout.addWidget(self.email_column_combo)
        
        layout.addStretch()
    
    def update_columns(self, columns):
        """Update available columns in the dropdown."""
        self.csv_columns = columns
        self.email_column_combo.clear()
        
        if columns:
            # Auto-detect email column (more flexible detection)
            email_candidates = [col for col in columns if 'email' in col.lower() or 'mail' in col.lower()]
            
            # Add all columns
            self.email_column_combo.addItems(columns)
            
            # Select email column if found, otherwise select first column
            if email_candidates:
                self.email_column_combo.setCurrentText(email_candidates[0])
            else:
                # If no email-like column found, just select the first column
                self.email_column_combo.setCurrentIndex(0)

    def get_selected_column(self):
        """Get the selected email column."""
        return self.email_column_combo.currentText()


class EmailWorker(QThread):
    """Worker thread for email operations."""
    
    progress_updated = pyqtSignal(int, int)
    status_updated = pyqtSignal(str)
    email_sent = pyqtSignal(str, bool, str)
    finished = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)
    
    def __init__(self, email_app: EmailApplication, operation: str, **kwargs):
        super().__init__()
        self.email_app = email_app
        self.operation = operation
        self.kwargs = kwargs
        self.should_stop = False
    
    def run(self):
        """Run the email operation in a separate thread."""
        try:
            if self.operation == 'test':
                self._run_test_email()
            elif self.operation == 'bulk':
                self._run_bulk_email()
        except Exception as e:
            self.error_occurred.emit(f"Operation failed: {str(e)}")
    
    def _run_test_email(self):
        """Run test email operation."""
        try:
            self.status_updated.emit("Sending test email...")
            success = self.email_app.send_test_email(self.kwargs.get('test_recipient'))
            
            if success:
                self.finished.emit({'success': True, 'message': 'Test email sent'})
            else:
                self.finished.emit({'success': False, 'message': 'Test email failed'})
        except Exception as e:
            self.error_occurred.emit(f"Test email error: {str(e)}")
    
    def _run_bulk_email(self):
        """Run bulk email operation."""
        try:
            self.status_updated.emit("Reading recipients...")
            recipients = self.kwargs.get('recipients_data', [])
            
            total = len(recipients)
            self.progress_updated.emit(0, total)
            
            smtp_config = self.email_app.config['smtp']
            template_content = self.kwargs.get('template_content', '')
            subject = self.kwargs.get('subject', 'Email from Application')
            attachments = self.kwargs.get('attachments', [])
            
            results = {'total': total, 'sent': 0, 'failed': 0, 'errors': []}
            
            for i, recipient_data in enumerate(recipients):
                if self.should_stop:
                    break
                
                try:
                    self.progress_updated.emit(i + 1, total)
                    self.status_updated.emit(f"Sending to {recipient_data.get('email', 'Unknown')}...")
                    
                    # Render template with recipient data
                    personalized_body = self.render_template(template_content, recipient_data)
                    personalized_subject = self.render_template(subject, recipient_data)
                    
                    success = self.email_app.email_sender.send_email(
                        smtp_config=smtp_config,
                        sender=smtp_config['username'],
                        recipient=recipient_data.get('email'),
                        subject=personalized_subject,
                        body=personalized_body,
                        attachments=attachments
                    )
                    
                    if success:
                        results['sent'] += 1
                        self.email_sent.emit(recipient_data.get('email', 'Unknown'), True, "")
                    else:
                        results['failed'] += 1
                        self.email_sent.emit(recipient_data.get('email', 'Unknown'), False, "Send failed")
                        
                except Exception as e:
                    results['failed'] += 1
                    error_msg = str(e)
                    results['errors'].append(error_msg)
                    self.email_sent.emit(recipient_data.get('email', 'Unknown'), False, error_msg)
            
            results['success'] = results['failed'] == 0
            self.finished.emit(results)
            
        except Exception as e:
            self.error_occurred.emit(f"Bulk email error: {str(e)}")
    
    def render_template(self, template_content, data):
        """Simple template rendering with attribute replacement."""
        rendered = template_content
        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}"
            rendered = rendered.replace(placeholder, str(value))
        return rendered
    
    def stop(self):
        """Stop the email operation."""
        self.should_stop = True


class ModernCard(QFrame):
    """Modern card widget with rounded corners and shadow effect."""
    
    def __init__(self, title: str, subtitle: str = "", parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setProperty("class", "card")
        self.setup_ui(title, subtitle)
    
    def setup_ui(self, title: str, subtitle: str):
        """Setup card layout with title and content area."""
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 20)
        self.setLayout(layout)
        
        if title:
            # Card title
            title_label = QLabel(title)
            title_label.setProperty("class", "cardTitle")
            layout.addWidget(title_label)
        
        if subtitle:
            # Card subtitle
            subtitle_label = QLabel(subtitle)
            subtitle_label.setProperty("class", "cardSubtitle")
            layout.addWidget(subtitle_label)
        
        # Content area
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout()
        self.content_layout.setContentsMargins(20, 0, 20, 20)
        self.content_widget.setLayout(self.content_layout)
        layout.addWidget(self.content_widget)
    
    def add_content(self, widget):
        """Add content widget to the card."""
        self.content_layout.addWidget(widget)


class StatsCard(QFrame):
    """Statistics card widget for campaign display."""
    
    def __init__(self, title: str, value: str, icon: str = "📊", color: str = "#3b82f6", parent=None):
        super().__init__(parent)
        self.setProperty("class", "stats-card")
        self.setup_ui(title, value, icon, color)
    
    def setup_ui(self, title: str, value: str, icon: str, color: str):
        """Setup stats card layout."""
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(8)
        self.setLayout(layout)
        
        # Icon and value row
        top_row = QHBoxLayout()
        
        icon_label = QLabel(icon)
        icon_label.setStyleSheet(f"""
            font-size: 24px;
            color: {color};
            padding: 5px;
        """)
        top_row.addWidget(icon_label)
        
        top_row.addStretch()
        
        value_label = QLabel(value)
        value_label.setStyleSheet(f"""
            font-size: 28px;
            font-weight: bold;
            color: {color};
        """)
        value_label.setAlignment(Qt.AlignRight)
        top_row.addWidget(value_label)
        
        layout.addLayout(top_row)
        
        # Title
        title_label = QLabel(title)
        title_label.setStyleSheet("""
            font-size: 14px;
            font-weight: 500;
            color: #6b7280;
        """)
        layout.addWidget(title_label)
    
    def update_value(self, value: str):
        """Update the stats value."""
        # Find value label and update it
        for i in range(self.layout().count()):
            item = self.layout().itemAt(i)
            if isinstance(item, QHBoxLayout):
                for j in range(item.count()):
                    widget = item.itemAt(j).widget()
                    if isinstance(widget, QLabel) and "font-size: 28px" in widget.styleSheet():
                        widget.setText(value)
                        break


class SetupPage(QWidget):
    """Setup page for configuring email account and CSV file."""
    
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the account and file configuration interface."""
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)
        self.setLayout(main_layout)
        
        # Account Configuration Card
        account_card = ModernCard("🔐 Email Account", "Configure your email provider and credentials")
        account_content = QWidget()
        account_layout = QVBoxLayout()
        account_content.setLayout(account_layout)
        
        # Provider and credentials layout
        creds_layout = QVBoxLayout()
        creds_layout.setSpacing(15)
        
        # Provider row
        provider_row = QHBoxLayout()
        provider_row.setSpacing(10)
        provider_label = QLabel("Provider:")
        provider_label.setMinimumWidth(80)
        provider_row.addWidget(provider_label)
        
        self.provider_combo = QComboBox()
        self.provider_combo.addItems(["Gmail", "Outlook", "Yahoo", "iCloud", "Custom"])
        self.provider_combo.currentTextChanged.connect(self.on_provider_changed)
        self.provider_combo.setMinimumHeight(45)
        provider_row.addWidget(self.provider_combo)
        
        provider_row.addStretch()
        # Connection status
        self.connection_status = QLabel("⚫ Not Connected")
        self.connection_status.setProperty("class", "connection-status disconnected")
        provider_row.addWidget(self.connection_status)
        
        creds_layout.addLayout(provider_row)
        
        # Email row
        email_row = QHBoxLayout()
        email_row.setSpacing(10)
        email_label = QLabel("Email:")
        email_label.setMinimumWidth(80)
        email_row.addWidget(email_label)
        
        self.email_edit = QLineEdit()
        self.email_edit.setPlaceholderText("your.email@gmail.com")
        self.email_edit.setMinimumHeight(45)
        email_row.addWidget(self.email_edit)
        
        creds_layout.addLayout(email_row)
        
        # Password row
        password_row = QHBoxLayout()
        password_row.setSpacing(10)
        password_label = QLabel("Password:")
        password_label.setMinimumWidth(80)
        password_row.addWidget(password_label)
        
        # Password container that looks like other input fields
        password_container = QFrame()
        password_container.setMinimumHeight(45)
        password_container.setStyleSheet("""
            QFrame {
                border: 2px solid #e2e8f0;
                border-radius: 8px;
                background: white;
                padding: 0px;
            }
            QFrame:focus-within {
                border-color: #3b82f6;
            }
        """)
        password_container_layout = QHBoxLayout()
        password_container_layout.setContentsMargins(12, 0, 0, 0)
        password_container_layout.setSpacing(0)
        
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.password_edit.setPlaceholderText("Your email password or app password")
        self.password_edit.setStyleSheet("QLineEdit { border: none; background: transparent; padding: 8px 0; }")
        password_container_layout.addWidget(self.password_edit)
        
        # Password visibility toggle button
        self.password_eye_btn = QPushButton("👁️")
        self.password_eye_btn.setFixedSize(50, 35)
        self.password_eye_btn.setStyleSheet("""
            QPushButton {
                border: none;
                background: transparent;
                color: #6b7280;
                font-size: 16px;
                margin-right: 8px;
                border-radius: 6px;
            }
            QPushButton:hover {
                color: #374151;
                background: #e5e7eb;
            }
            QPushButton:pressed {
                background: #d1d5db;
            }
        """)
        self.password_eye_btn.clicked.connect(self.toggle_password_visibility)
        password_container_layout.addWidget(self.password_eye_btn)
        
        password_container.setLayout(password_container_layout)
        password_row.addWidget(password_container)
        
        # Test connection button
        self.test_btn = LoadingButton("🔗 Test Connection")
        self.test_btn.setProperty("class", "success")
        self.test_btn.setMinimumSize(140, 45)
        self.test_btn.clicked.connect(self.test_connection)
        password_row.addWidget(self.test_btn)
        
        creds_layout.addLayout(password_row)
        
        # Advanced settings (compact)
        advanced_layout = QHBoxLayout()
        advanced_layout.setSpacing(15)
        
        smtp_label = QLabel("SMTP:")
        smtp_label.setMinimumWidth(80)
        advanced_layout.addWidget(smtp_label)
        
        self.smtp_server_edit = QLineEdit()
        self.smtp_server_edit.setPlaceholderText("smtp.gmail.com")
        self.smtp_server_edit.setMinimumHeight(45)
        advanced_layout.addWidget(self.smtp_server_edit)
        
        port_label = QLabel("Port:")
        advanced_layout.addWidget(port_label)
        
        self.smtp_port_spin = QSpinBox()
        self.smtp_port_spin.setRange(1, 65535)
        self.smtp_port_spin.setValue(587)
        self.smtp_port_spin.setFixedWidth(80)
        self.smtp_port_spin.setMinimumHeight(40)
        advanced_layout.addWidget(self.smtp_port_spin)
        
        self.tls_check = QCheckBox("TLS")
        self.tls_check.setChecked(True)
        advanced_layout.addWidget(self.tls_check)
        
        self.ssl_check = QCheckBox("SSL")
        advanced_layout.addWidget(self.ssl_check)
        
        account_layout.addLayout(creds_layout)
        account_layout.addLayout(advanced_layout)
        
        account_card.add_content(account_content)
        main_layout.addWidget(account_card)
        
        # Files Configuration Card
        files_card = ModernCard("📁 Recipients File", "Upload your CSV file with recipient data")
        files_content = QWidget()
        files_layout = QVBoxLayout()
        files_content.setLayout(files_layout)
        
        # CSV file selection
        csv_layout = QHBoxLayout()
        csv_layout.addWidget(QLabel("CSV File:"))
        
        self.csv_file_edit = QLineEdit()
        self.csv_file_edit.setPlaceholderText("Select your recipients CSV file...")
        csv_layout.addWidget(self.csv_file_edit)
        
        self.csv_browse_btn = LoadingButton("📂 Browse")
        self.csv_browse_btn.setProperty("class", "secondary")
        self.csv_browse_btn.clicked.connect(self.browse_csv)
        csv_layout.addWidget(self.csv_browse_btn)
        
        files_layout.addLayout(csv_layout)
        
        files_card.add_content(files_content)
        main_layout.addWidget(files_card)
        
        main_layout.addStretch()
        
        # Set default provider
        self.on_provider_changed("Gmail")
    
    def on_provider_changed(self, provider):
        """Handle email provider selection change."""
        # SMTP configurations for different providers
        providers = {
            "Gmail": {"server": "smtp.gmail.com", "port": 587, "tls": True, "ssl": False},
            "Outlook": {"server": "smtp-mail.outlook.com", "port": 587, "tls": True, "ssl": False},
            "Yahoo": {"server": "smtp.mail.yahoo.com", "port": 587, "tls": True, "ssl": False},
            "iCloud": {"server": "smtp.mail.me.com", "port": 587, "tls": True, "ssl": False},
            "Custom": {"server": "", "port": 587, "tls": True, "ssl": False}
        }
        
        config = providers.get(provider, providers["Custom"])
        self.smtp_server_edit.setText(config["server"])
        self.smtp_port_spin.setValue(config["port"])
        self.tls_check.setChecked(config["tls"])
        self.ssl_check.setChecked(config["ssl"])
        
        # Update placeholder texts based on provider
        placeholders = {
            "Gmail": ("your.email@gmail.com", "Use App Password for Gmail"),
            "Outlook": ("your.email@outlook.com", "Your Outlook password"),
            "Yahoo": ("your.email@yahoo.com", "Your Yahoo password"), 
            "iCloud": ("your.email@icloud.com", "Your iCloud password"),
            "Custom": ("your.email@domain.com", "Your email password")
        }
        
        email_placeholder, pass_placeholder = placeholders.get(provider, placeholders["Custom"])
        self.email_edit.setPlaceholderText(email_placeholder)
        self.password_edit.setPlaceholderText(pass_placeholder)
    
    def toggle_password_visibility(self):
        """Toggle password field visibility between hidden and visible."""
        try:
            if hasattr(self, 'password_edit') and self.password_edit:
                if self.password_edit.echoMode() == QLineEdit.Password:
                    self.password_edit.setEchoMode(QLineEdit.Normal)
                    self.password_eye_btn.setText("🙈")
                else:
                    self.password_edit.setEchoMode(QLineEdit.Password)
                    self.password_eye_btn.setText("👁️")
        except:
            pass
    
    def test_connection(self):
        """Test SMTP connection with loading animation."""
        self.test_btn.start_loading("Testing")
        if self.main_window:
            self.main_window.test_smtp_connection()
    
    def browse_csv(self):
        """Browse and select CSV file with loading animation."""
        self.csv_browse_btn.start_loading("Loading")
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Recipients CSV File", "", "CSV Files (*.csv);;All Files (*)"
        )
        
        self.csv_browse_btn.stop_loading()
        
        if file_path:
            self.csv_file_edit.setText(file_path)
            if self.main_window:
                self.main_window.on_files_changed()
    
    def update_connection_status(self, connected: bool, message: str = ""):
        """Update the connection status indicator."""
        if connected:
            self.connection_status.setText("🟢 Connected")
            self.connection_status.setProperty("class", "connection-status connected")
        else:
            self.connection_status.setText("🔴 Disconnected")
            self.connection_status.setProperty("class", "connection-status disconnected")
        
        # Apply style changes
        self.connection_status.style().unpolish(self.connection_status)
        self.connection_status.style().polish(self.connection_status)
        self.test_btn.stop_loading()
        
        # Show result message if provided
        if message and self.main_window:
            if connected:
                self.main_window.show_info("Connection Test", f"✅ {message}")
            else:
                self.main_window.show_error("Connection Test", f"❌ {message}")
    
    def get_config(self):
        """Get current configuration settings."""
        # Safe password access to prevent widget deletion errors
        password_text = ""
        try:
            if hasattr(self, 'password_edit') and self.password_edit:
                password_text = self.password_edit.text()
        except:
            password_text = ""
        
        return {
            'smtp': {
                'server': self.smtp_server_edit.text(),
                'port': self.smtp_port_spin.value(),
                'username': self.email_edit.text(),
                'password': password_text,
                'use_tls': self.tls_check.isChecked(),
                'use_ssl': self.ssl_check.isChecked(),
                'timeout': 30
            },
            'files': {
                'csv_recipients': self.csv_file_edit.text(),
            },
            'email': {
                'test_email': self.email_edit.text()
            }
        }
    
    def set_config(self, config):
        """Set configuration settings in the UI."""
        smtp_config = config.get('smtp', {})
        self.smtp_server_edit.setText(smtp_config.get('server', ''))
        self.smtp_port_spin.setValue(smtp_config.get('port', 587))
        self.email_edit.setText(smtp_config.get('username', ''))
        
        # Safe password setting to prevent widget deletion errors
        try:
            if hasattr(self, 'password_edit') and self.password_edit:
                self.password_edit.setText(smtp_config.get('password', ''))
        except:
            pass
        
        self.tls_check.setChecked(smtp_config.get('use_tls', True))
        self.ssl_check.setChecked(smtp_config.get('use_ssl', False))
        
        files_config = config.get('files', {})
        self.csv_file_edit.setText(files_config.get('csv_recipients', ''))


class RecipientsPage(QWidget):
    """Recipients preview page with email column selection and centered content."""
    
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the recipients preview interface with centered content."""
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)
        self.setLayout(main_layout)
        
        # Recipients Card with centered content
        recipients_card = ModernCard("", "")
        recipients_content = QWidget()
        recipients_layout = QVBoxLayout()
        recipients_layout.setAlignment(Qt.AlignCenter)
        recipients_content.setLayout(recipients_layout)
        
        # Centered title
        title_label = QLabel("👥 Recipients Preview")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #1e293b; text-align: center;")
        title_label.setAlignment(Qt.AlignCenter)
        recipients_layout.addWidget(title_label)
        
        # Centered subtitle
        subtitle_label = QLabel("Review your email recipients and select email column")
        subtitle_label.setStyleSheet("font-size: 14px; color: #64748b; text-align: center; margin-bottom: 15px;")
        subtitle_label.setAlignment(Qt.AlignCenter)
        recipients_layout.addWidget(subtitle_label)
        
        # Top controls row - centered
        controls_layout = QHBoxLayout()
        controls_layout.setAlignment(Qt.AlignCenter)
        
        self.refresh_btn = LoadingButton("🔄 Refresh Recipients")
        self.refresh_btn.setProperty("class", "secondary")
        self.refresh_btn.clicked.connect(self.refresh_recipients)
        controls_layout.addWidget(self.refresh_btn)
        
        self.recipients_info_label = QLabel("No recipients loaded")
        self.recipients_info_label.setStyleSheet("color: #64748b; font-style: italic; margin-left: 10px;")
        controls_layout.addWidget(self.recipients_info_label)
        
        # Email column selector
        self.email_column_selector = EmailColumnSelector()
        controls_layout.addWidget(self.email_column_selector)
        
        recipients_layout.addLayout(controls_layout)
        
        # Available attributes display - centered
        self.attributes_label = QLabel("Available attributes: None")
        self.attributes_label.setStyleSheet("color: #3b82f6; font-weight: bold; font-size: 13px; margin: 10px 0; text-align: center;")
        self.attributes_label.setAlignment(Qt.AlignCenter)
        recipients_layout.addWidget(self.attributes_label)
        
        # Recipients table with auto-sizing and height expansion
        self.recipients_table = QTableWidget()
        self.recipients_table.setAlternatingRowColors(True)
        self.recipients_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.recipients_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.recipients_table.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        self.recipients_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.recipients_table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        # Auto-resize columns to content
        self.recipients_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        # Don't stretch last column to prevent overflow
        self.recipients_table.horizontalHeader().setStretchLastSection(False)
        
        recipients_layout.addWidget(self.recipients_table)
        
        recipients_card.add_content(recipients_content)
        main_layout.addWidget(recipients_card)
    
    def refresh_recipients(self):
        """Refresh recipients data with loading animation."""
        self.refresh_btn.start_loading("Loading")
        if self.main_window:
            self.main_window.refresh_recipients()
    
    def update_recipients(self, recipients_data, csv_file_name=""):
        """Update recipients display with auto-sizing columns and centered content."""
        self.refresh_btn.stop_loading()
        
        if not recipients_data:
            self.recipients_info_label.setText("No recipients loaded")
            self.attributes_label.setText("Available attributes: None")
            self.recipients_table.setRowCount(0)
            self.recipients_table.setColumnCount(0)
            self.email_column_selector.update_columns([])
            return
        
        # Update info label
        self.recipients_info_label.setText(
            f"📊 {len(recipients_data)} recipients loaded from {csv_file_name}"
        )
        
        if recipients_data:
            columns = list(recipients_data[0].keys())
            
            # Update email column selector
            self.email_column_selector.update_columns(columns)
            
            # Update attributes display
            attr_text = ", ".join([f"{{{{ {col} }}}}" for col in columns])
            self.attributes_label.setText(f"Available attributes: {attr_text}")
            
            # Update table
            self.recipients_table.setColumnCount(len(columns))
            self.recipients_table.setHorizontalHeaderLabels(columns)
            self.recipients_table.setRowCount(len(recipients_data))
            
            for row, recipient in enumerate(recipients_data):
                for col, header in enumerate(columns):
                    item = QTableWidgetItem(str(recipient.get(header, '')))
                    item.setTextAlignment(Qt.AlignCenter)
                    self.recipients_table.setItem(row, col, item)
            
            # Auto-resize columns to fit content
            self.recipients_table.resizeColumnsToContents()
            
            # Ensure table doesn't exceed card width
            header = self.recipients_table.horizontalHeader()
            total_width = sum(self.recipients_table.columnWidth(i) for i in range(len(columns)))
            available_width = self.recipients_table.parent().width() - 60
            
            if total_width > available_width:
                # Scale down columns proportionally
                scale_factor = available_width / total_width
                for col in range(len(columns)):
                    current_width = self.recipients_table.columnWidth(col)
                    new_width = int(current_width * scale_factor)
                    self.recipients_table.setColumnWidth(col, max(new_width, 80))
    
    def get_selected_email_column(self):
        """Get the selected email column name."""
        return self.email_column_selector.get_selected_column()


class ComposePage(QWidget):
    """Email composition page with centered content and modern dropdowns."""
    
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.csv_columns = []
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the email composition interface with centered content."""
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)
        self.setLayout(main_layout)
        
        # Email Composition Card
        compose_card = ModernCard("✏️ Compose Email", "Write your email subject and content with personalized attributes")
        compose_content = QWidget()
        compose_layout = QVBoxLayout()
        compose_content.setLayout(compose_layout)
        
        # Subject section
        subject_layout = QVBoxLayout()
        subject_layout.addWidget(QLabel("Subject:"))
        
        self.subject_edit = QLineEdit()
        self.subject_edit.setPlaceholderText("Enter your email subject... (use attributes like {{name}})")
        subject_layout.addWidget(self.subject_edit)
        
                # Subject attributes dropdown
        self.subject_attributes = ModernAttributeDropdown(self.subject_edit, self.csv_columns)
        subject_layout.addWidget(self.subject_attributes)
        
        compose_layout.addLayout(subject_layout)
        
        # Email content section
        content_layout = QVBoxLayout()
        content_layout.addWidget(QLabel("Email Content:"))
        
        self.email_content_edit = QTextEdit()
        self.email_content_edit.setPlaceholderText("""Write your email content here...

You can use attributes from your CSV file like:
- {{name}} for recipient name
- {{company}} for company name
- {{email}} for email address

Example:
Dear {{name}},

Thank you for your interest in our services. We're excited to work with {{company}}!

Best regards,
The Team""")
        self.email_content_edit.setMinimumHeight(400)
        content_layout.addWidget(self.email_content_edit)
        
        # Content attributes dropdown
        self.content_attributes = ModernAttributeDropdown(self.email_content_edit, self.csv_columns)
        content_layout.addWidget(self.content_attributes)
        
        compose_layout.addLayout(content_layout)
        
        compose_card.add_content(compose_content)
        main_layout.addWidget(compose_card)
        
        # Attachments Card with left-aligned content
        attachments_card = ModernCard("", "")
        attachments_content = QWidget()
        attachments_layout = QVBoxLayout()
        attachments_content.setLayout(attachments_layout)

        # Left-aligned title
        attach_title = QLabel("📎 Attachments")
        attach_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #1e293b;")
        attach_title.setAlignment(Qt.AlignLeft)  # Changed to left
        attachments_layout.addWidget(attach_title)

        # Left-aligned subtitle
        attach_subtitle = QLabel("Add files to your email campaign")
        attach_subtitle.setStyleSheet("font-size: 14px; color: #64748b; margin-bottom: 15px;")
        attach_subtitle.setAlignment(Qt.AlignLeft)  # Changed to left
        attachments_layout.addWidget(attach_subtitle)

        # Attachment controls - left-aligned
        attach_controls = QHBoxLayout()
        
        self.add_btn = LoadingButton("➕ Add Files")
        self.add_btn.setProperty("class", "success")
        self.add_btn.clicked.connect(self.add_attachment)
        attach_controls.addWidget(self.add_btn)
        
        remove_btn = QPushButton("➖ Remove")
        remove_btn.setProperty("class", "warning")
        remove_btn.clicked.connect(self.remove_attachment)
        attach_controls.addWidget(remove_btn)
        
        clear_btn = QPushButton("🗑️ Clear All")
        clear_btn.setProperty("class", "secondary")
        clear_btn.clicked.connect(self.clear_attachments)
        attach_controls.addWidget(clear_btn)
        
        attachments_layout.addLayout(attach_controls)
        
        # Attachments table with full width and proportional columns
        self.attachments_table = QTableWidget()
        self.attachments_table.setColumnCount(2)
        self.attachments_table.setHorizontalHeaderLabels(["📁 File", "📏 Size"])
        self.attachments_table.setAlternatingRowColors(True)
        self.attachments_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.attachments_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.attachments_table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        # Make table read-only
        self.attachments_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.attachments_table.setSelectionBehavior(QTableWidget.SelectRows)

        # Set column widths proportionally - name column gets 70%, size column gets 30%
        header = self.attachments_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)  # Name column stretches
        header.setSectionResizeMode(1, QHeaderView.Fixed)    # Size column fixed
        self.attachments_table.setColumnWidth(1, 100)  # Fixed width for size column

        # Set minimum height but allow expansion
        self.attachments_table.setMinimumHeight(400)

        # Add table with full width (no container layout)
        attachments_layout.addWidget(self.attachments_table)
        
        attachments_card.add_content(attachments_content)
        main_layout.addWidget(attachments_card)
    
    def update_csv_columns(self, columns):
        """Update available CSV columns for attribute insertion."""
        self.csv_columns = columns
        self.subject_attributes.update_attributes(columns)
        self.content_attributes.update_attributes(columns)
    
    def add_attachment(self):
        """Add file attachments with loading animation."""
        self.add_btn.start_loading("Adding")
        
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Files to Attach", "", "All Files (*)"
        )
        
        for file_path in file_paths:
            if file_path and os.path.exists(file_path):
                row = self.attachments_table.rowCount()
                self.attachments_table.insertRow(row)
                
                file_name = os.path.basename(file_path)
                file_size = os.path.getsize(file_path)
                size_str = self.format_file_size(file_size)
                
                # Create items with center alignment
                name_item = QTableWidgetItem(file_name)
                name_item.setTextAlignment(Qt.AlignCenter)
                self.attachments_table.setItem(row, 0, name_item)
                
                size_item = QTableWidgetItem(size_str)
                size_item.setTextAlignment(Qt.AlignCenter)
                self.attachments_table.setItem(row, 1, size_item)
                
                # Store full path in item data for later retrieval
                name_item.setData(Qt.UserRole, file_path)
        
        self.add_btn.stop_loading()
    
    def remove_attachment(self):
        """Remove selected attachment only if user has manually selected a row."""
        selected_items = self.attachments_table.selectedItems()
        current_row = self.attachments_table.currentRow()
        
        # Check if user has actually selected something (not just default selection)
        if selected_items and current_row >= 0:
            self.attachments_table.removeRow(current_row)
        else:
            # Show message if no row is manually selected
            if self.main_window:
                self.main_window.show_warning("No Selection", "Please select an attachment to remove.")
    
    def clear_attachments(self):
        """Clear all attachments."""
        self.attachments_table.setRowCount(0)
    
    def format_file_size(self, size_bytes):
        """Format file size in human-readable format."""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024**2:
            return f"{size_bytes/1024:.1f} KB"
        elif size_bytes < 1024**3:
            return f"{size_bytes/1024**2:.1f} MB"
        else:
            return f"{size_bytes/1024**3:.1f} GB"
    
    def get_attachments_list(self):
        """Get list of attachment file paths."""
        attachments = []
        for row in range(self.attachments_table.rowCount()):
            item = self.attachments_table.item(row, 0)
            if item:
                file_path = item.data(Qt.UserRole)
                if file_path:
                    attachments.append(file_path)
        return attachments


class SendPage(QWidget):
    """Modern campaign launch page with improved layout and statistics."""
    
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the modern campaign launch interface."""
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(25)
        self.setLayout(main_layout)
        
        # Campaign Overview Card with Stats
        overview_card = ModernCard("🚀 Campaign Dashboard", "Monitor and launch your email campaigns")
        overview_content = QWidget()
        overview_layout = QVBoxLayout()
        overview_content.setLayout(overview_layout)
        
        # Statistics Row
        stats_layout = QHBoxLayout()
        stats_layout.setSpacing(15)
        
        # Create stats cards
        self.recipients_stats = StatsCard("Total Recipients", "0", "👥", "#3b82f6")
        self.attachments_stats = StatsCard("Attachments", "0", "📎", "#10b981")
        self.status_stats = StatsCard("Status", "Ready", "📊", "#f59e0b")
        
        stats_layout.addWidget(self.recipients_stats)
        stats_layout.addWidget(self.attachments_stats)
        stats_layout.addWidget(self.status_stats)
        
        overview_layout.addLayout(stats_layout)
        
        overview_card.add_content(overview_content)
        main_layout.addWidget(overview_card)
        
        # Campaign Actions Card
        actions_card = ModernCard("🎯 Campaign Actions", "Test your setup or launch the full campaign")
        actions_content = QWidget()
        actions_layout = QVBoxLayout()
        actions_layout.setAlignment(Qt.AlignCenter)
        actions_content.setLayout(actions_layout)
        
        # Action buttons with modern styling
        buttons_layout = QHBoxLayout()
        buttons_layout.setAlignment(Qt.AlignCenter)
        buttons_layout.setSpacing(25)
        
        # Test Email Button
        self.test_btn = LoadingButton("🧪 Send Test Email")
        self.test_btn.setProperty("class", "campaign-btn campaign-primary")
        self.test_btn.clicked.connect(self.send_test)
        buttons_layout.addWidget(self.test_btn)
        
        # Launch Campaign Button
        self.send_all_btn = LoadingButton("🚀 Launch Campaign")
        self.send_all_btn.setProperty("class", "campaign-btn campaign-success")
        self.send_all_btn.clicked.connect(self.send_bulk)
        buttons_layout.addWidget(self.send_all_btn)
        
        # Stop Button
        self.stop_btn = QPushButton("⏹️ Stop Campaign")
        self.stop_btn.setProperty("class", "campaign-btn campaign-danger")
        self.stop_btn.clicked.connect(self.stop_sending)
        self.stop_btn.setEnabled(False)
        buttons_layout.addWidget(self.stop_btn)
        
        actions_layout.addLayout(buttons_layout)
        
        actions_card.add_content(actions_content)
        main_layout.addWidget(actions_card)
        
        # Progress Monitoring Card
        progress_card = ModernCard("📈 Campaign Progress", "Real-time monitoring of your email campaign")
        progress_content = QWidget()
        progress_layout = QVBoxLayout()
        progress_content.setLayout(progress_layout)
        
        # Progress bar with modern styling
        progress_container = QVBoxLayout()
        progress_container.setSpacing(15)
        
        # Progress info row
        progress_info_layout = QHBoxLayout()
        
        self.progress_label = QLabel("Campaign Progress")
        self.progress_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #374151;")
        progress_info_layout.addWidget(self.progress_label)
        
        progress_info_layout.addStretch()
        
        self.progress_percentage = QLabel("0%")
        self.progress_percentage.setStyleSheet("font-size: 16px; font-weight: bold; color: #3b82f6;")
        progress_info_layout.addWidget(self.progress_percentage)
        
        progress_container.addLayout(progress_info_layout)
        
        # Enhanced progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 12px;
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                                           stop: 0 #f1f5f9, stop: 1 #e2e8f0);
                height: 24px;
                text-align: center;
                font-weight: bold;
                color: #1e293b;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                                           stop: 0 #3b82f6, stop: 0.5 #60a5fa, stop: 1 #93c5fd);
                border-radius: 12px;
                margin: 2px;
            }
        """)
        progress_container.addWidget(self.progress_bar)
        
        # Status message with modern styling
        self.status_label = QLabel("Ready to launch email campaign")
        self.status_label.setStyleSheet("""
            font-size: 14px; 
            color: #6b7280; 
            font-style: italic;
            text-align: center;
            padding: 10px;
            background: #f8fafc;
            border-radius: 8px;
            border: 1px solid #e2e8f0;
        """)
        self.status_label.setAlignment(Qt.AlignCenter)
        progress_container.addWidget(self.status_label)
        
        progress_layout.addLayout(progress_container)
        
        progress_card.add_content(progress_content)
        main_layout.addWidget(progress_card)
        
        # Campaign Results Card (Initially hidden)
        self.results_card = ModernCard("🎉 Campaign Results", "Summary of your email campaign performance")
        results_content = QWidget()
        results_layout = QHBoxLayout()
        results_content.setLayout(results_layout)
        
        # Results will be populated dynamically
        self.results_stats_layout = QHBoxLayout()
        results_layout.addLayout(self.results_stats_layout)
        
        self.results_card.add_content(results_content)
        self.results_card.setVisible(False)
        main_layout.addWidget(self.results_card)
    
    def send_test(self):
        """Send test email with loading animation."""
        self.test_btn.start_loading("Sending")
        self.status_stats.update_value("Testing")
        if self.main_window:
            self.main_window.send_test_email()
    
    def send_bulk(self):
        """Send bulk emails with loading animation."""
        self.send_all_btn.start_loading("Launching")
        self.status_stats.update_value("Launching")
        if self.main_window:
            self.main_window.send_bulk_emails()
    
    def stop_sending(self):
        """Stop email campaign."""
        if self.main_window:
            self.main_window.stop_sending()
        self.status_stats.update_value("Stopping")
    
    def update_ui_for_sending(self, is_sending):
        """Update UI state during email sending operations."""
        if is_sending:
            self.send_all_btn.stop_loading()
            self.send_all_btn.setEnabled(False)
            self.stop_btn.setEnabled(True)
            self.status_stats.update_value("Sending")
        else:
            self.send_all_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            self.test_btn.stop_loading()
            self.status_stats.update_value("Ready")
    
    def update_stats(self, recipients_count=0, attachments_count=0):
        """Update campaign statistics."""
        self.recipients_stats.update_value(str(recipients_count))
        self.attachments_stats.update_value(str(attachments_count))
    
    def update_progress_info(self, current, total):
        """Update progress information."""
        if total > 0:
            percentage = int((current / total) * 100)
            self.progress_percentage.setText(f"{percentage}%")
            self.progress_label.setText(f"Sending emails... {current}/{total}")
        else:
            self.progress_percentage.setText("0%")
            self.progress_label.setText("Campaign Progress")
    
    def show_campaign_results(self, results):
        """Show campaign results in the results card."""
        # Clear previous results
        for i in reversed(range(self.results_stats_layout.count())):
            self.results_stats_layout.itemAt(i).widget().setParent(None)
        
        # Add new result stats
        total_stats = StatsCard("Total Sent", str(results.get('total', 0)), "📧", "#3b82f6")
        success_stats = StatsCard("Successful", str(results.get('sent', 0)), "✅", "#10b981")
        failed_stats = StatsCard("Failed", str(results.get('failed', 0)), "❌", "#ef4444")
        
        success_rate = 0
        if results.get('total', 0) > 0:
            success_rate = int((results.get('sent', 0) / results.get('total', 0)) * 100)
        rate_stats = StatsCard("Success Rate", f"{success_rate}%", "📊", "#f59e0b")
        
        self.results_stats_layout.addWidget(total_stats)
        self.results_stats_layout.addWidget(success_stats)
        self.results_stats_layout.addWidget(failed_stats)
        self.results_stats_layout.addWidget(rate_stats)
        
        # Show results card
        self.results_card.setVisible(True)


class LogPage(QWidget):
    """Activity log page with white font and export functionality."""
    
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the activity log interface."""
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)
        self.setLayout(main_layout)
        
        # Log Controls Card
        controls_card = ModernCard("📋 Activity Log", "Monitor email sending activity and system events")
        controls_content = QWidget()
        controls_layout = QHBoxLayout()
        controls_content.setLayout(controls_layout)
        
        clear_btn = QPushButton("🗑️ Clear Log")
        clear_btn.setProperty("class", "warning")
        clear_btn.clicked.connect(self.clear_log)
        controls_layout.addWidget(clear_btn)
        
        save_btn = QPushButton("💾 Export Log")
        save_btn.setProperty("class", "secondary")
        save_btn.clicked.connect(self.save_log)
        controls_layout.addWidget(save_btn)
        
        controls_layout.addStretch()
        
        controls_card.add_content(controls_content)
        main_layout.addWidget(controls_card)
        
        # Log Display Card
        log_card = ModernCard("", "")
        log_content = QWidget()
        log_layout = QVBoxLayout()
        log_content.setLayout(log_layout)
        
        # Log display with white font for better readability
        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        self.log_display.setProperty("class", "log-display")
        self.log_display.setFont(QFont("Consolas", 11))
        
        log_layout.addWidget(self.log_display)
        
        log_card.add_content(log_content)
        main_layout.addWidget(log_card)
    
    def clear_log(self):
        """Clear the activity log."""
        if self.main_window:
            self.main_window.clear_log()
    
    def save_log(self):
        """Export activity log to file."""
        if self.main_window:
            self.main_window.save_log()


class EmailAutomationApp(QMainWindow):
    """
    Main Email Automation Application
    
    A comprehensive email automation tool with modern GUI, CSV recipient management,
    template personalization, SMTP sending, and activity logging.
    """
    
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon('icon.png'))
        self.setWindowTitle("Email Automation")
        
        # Window configuration
        self.setMinimumSize(800, 600)
        self.resize(1400, 900)
        
        # Application state
        self.email_app = None
        self.config = {}
        self.worker = None
        self.recipients_data = []
        
        # Setup UI
        self.setup_ui()
        self.setup_menu()
        
        # Auto-refresh timer for dynamic content updates
        self.refresh_timer = QTimer()
        self.refresh_timer.timeout.connect(self.auto_refresh_content)
        self.refresh_timer.start(3000)
        
        # Load default configuration if available
        self.load_config_file('config.json')
    
    def setup_ui(self):
        """Setup the main user interface with sidebar navigation."""
        # Apply modern styling
        self.setStyleSheet(MODERN_STYLESHEET)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout with sidebar and content area
        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        central_widget.setLayout(main_layout)
        
        # Setup sidebar navigation
        self.setup_sidebar()
        main_layout.addWidget(self.sidebar)
        
        # Content area with scrollable pages
        content_widget = QWidget()
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)
        content_widget.setLayout(content_layout)
        
        # Top bar with page title and status
        self.setup_top_bar()
        content_layout.addWidget(self.top_bar)
        
        # Scrollable page container
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        
        self.page_stack = QStackedWidget()
        scroll_area.setWidget(self.page_stack)
        content_layout.addWidget(scroll_area)
        
        # Setup all application pages
        self.setup_pages()
        
        main_layout.addWidget(content_widget)
        
        # Set initial page
        self.nav_list.setCurrentRow(0)
        self.on_navigation_changed(0)
    
    def setup_sidebar(self):
        """Setup the sidebar navigation with modern styling."""
        self.sidebar = QFrame()
        self.sidebar.setObjectName("sidebar")
        
        sidebar_layout = QVBoxLayout()
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(0)
        self.sidebar.setLayout(sidebar_layout)
        
        # Application header
        header_label = QLabel("📧 Email Automation")
        header_label.setObjectName("appHeader")
        sidebar_layout.addWidget(header_label)
        
        # Navigation list
        self.nav_list = QListWidget()
        self.nav_list.setObjectName("navList")
        
        # Navigation items with icons and tooltips
        nav_items = [
            ("⚙️  Setup", "Configure account and upload CSV"),
            ("👥  Recipients", "Preview recipients and select email column"),
            ("✏️  Compose", "Write your email content"),
            ("🚀  Send", "Launch your email campaign"),
            ("📋  Logs", "View activity logs")
        ]
        
        for title, tooltip in nav_items:
            item = QListWidgetItem(title)
            item.setToolTip(tooltip)
            self.nav_list.addItem(item)
        
        self.nav_list.currentRowChanged.connect(self.on_navigation_changed)
        sidebar_layout.addWidget(self.nav_list)
        
        sidebar_layout.addStretch()
        
        # Connection status indicator
        self.connection_indicator = QLabel("🔴 Disconnected")
        self.connection_indicator.setAlignment(Qt.AlignCenter)
        self.connection_indicator.setStyleSheet("""
            QLabel {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                                           stop: 0 #64748b, stop: 1 #475569);
                color: white;
                border-radius: 12px;
                padding: 8px;
                margin: 10px;
                font-size: 12px;
                font-weight: bold;
                border: 1px solid #475569;
            }
        """)
        sidebar_layout.addWidget(self.connection_indicator)
    
    def setup_top_bar(self):
        """Setup the top bar with page title and status indicators."""
        self.top_bar = QFrame()
        self.top_bar.setObjectName("topBar")
        
        top_layout = QHBoxLayout()
        top_layout.setContentsMargins(20, 0, 20, 0)
        self.top_bar.setLayout(top_layout)
        
        # Page title
        self.page_title = QLabel("Setup Account")
        self.page_title.setObjectName("pageTitle")
        top_layout.addWidget(self.page_title)
        
        top_layout.addStretch()
        
        # Recipients count display
        self.recipients_count = QLabel("👥 Recipients: 0")
        self.recipients_count.setStyleSheet("color: #64748b; font-weight: bold;")
        top_layout.addWidget(self.recipients_count)
    
    def setup_pages(self):
        """Setup all application pages."""
        # Setup page for account and file configuration
        self.setup_page = SetupPage(self)
        self.page_stack.addWidget(self.setup_page)
        
        # Recipients page for preview and email column selection
        self.recipients_page = RecipientsPage(self)
        self.page_stack.addWidget(self.recipients_page)
        
        # Compose page for email content creation
        self.compose_page = ComposePage(self)
        self.page_stack.addWidget(self.compose_page)
        
        # Send page for campaign management
        self.send_page = SendPage(self)
        self.page_stack.addWidget(self.send_page)
        
        # Log page for activity monitoring
        self.log_page = LogPage(self)
        self.page_stack.addWidget(self.log_page)
    
    def setup_menu(self):
        """Setup the application menu bar with additional options."""
        menubar = self.menuBar()
        menubar.setStyleSheet("""
            QMenuBar {
                background-color: white;
                color: #1e293b;
                border-bottom: 1px solid #e2e8f0;
                padding: 4px;
            }
            QMenuBar::item {
                background: transparent;
                padding: 8px 12px;
                border-radius: 6px;
                margin: 2px;
            }
            QMenuBar::item:selected {
                background-color: #f1f5f9;
            }
        """)
        
        # File menu
        file_menu = menubar.addMenu('📁 File')
        
        load_action = QAction('📂 Load Configuration...', self)
        load_action.setShortcut('Ctrl+O')
        load_action.triggered.connect(self.load_config)
        file_menu.addAction(load_action)
        
        save_action = QAction('💾 Save Configuration...', self)
        save_action.setShortcut('Ctrl+S')
        save_action.triggered.connect(self.save_config)
        file_menu.addAction(save_action)
        
        file_menu.addSeparator()
        
        export_action = QAction('📤 Export Log...', self)
        export_action.triggered.connect(self.save_log)
        file_menu.addAction(export_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction('🚪 Exit', self)
        exit_action.setShortcut('Ctrl+Q')
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Tools menu
        tools_menu = menubar.addMenu('🛠️ Tools')
        
        test_config_action = QAction('🔧 Test SMTP Configuration', self)
        test_config_action.triggered.connect(self.test_smtp_connection)
        tools_menu.addAction(test_config_action)
        
        refresh_action = QAction('🔄 Refresh Recipients', self)
        refresh_action.setShortcut('F5')
        refresh_action.triggered.connect(self.refresh_recipients)
        tools_menu.addAction(refresh_action)
        
        tools_menu.addSeparator()
        
        clear_log_action = QAction('🗑️ Clear Activity Log', self)
        clear_log_action.triggered.connect(self.clear_log)
        tools_menu.addAction(clear_log_action)
        
        # Help menu
        help_menu = menubar.addMenu('❓ Help')
        
        github_action = QAction('📱 Visit GitHub', self)
        github_action.triggered.connect(lambda: QDesktopServices.openUrl(QUrl("https://github.com/ByeBye21")))
        help_menu.addAction(github_action)
        
        contact_action = QAction('📧 Contact Developer', self)
        contact_action.triggered.connect(lambda: QDesktopServices.openUrl(QUrl("mailto:younes0079@gmail.com")))
        help_menu.addAction(contact_action)
        
        help_menu.addSeparator()
        
        about_action = QAction('ℹ️ About Email Automation', self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
    
    def on_navigation_changed(self, index):
        """Handle navigation between different pages."""
        self.page_stack.setCurrentIndex(index)
        
        # Update page title based on current page
        titles = [
            "⚙️ Setup Account",
            "👥 Preview Recipients", 
            "✏️ Compose Email",
            "🚀 Launch Campaign",
            "📋 Activity Logs"
        ]
        
        if 0 <= index < len(titles):
            self.page_title.setText(titles[index])
        
        # Update send page stats when navigating to it
        if index == 3:  # Send page
            self.update_send_page_stats()
    
    def update_send_page_stats(self):
        """Update statistics on the send page."""
        recipients_count = len(self.recipients_data)
        attachments_count = self.compose_page.attachments_table.rowCount()
        self.send_page.update_stats(recipients_count, attachments_count)
    
    def update_connection_status(self, connected: bool):
        """Update the SMTP connection status indicator."""
        if connected:
            self.connection_indicator.setText("🟢 Connected")
            self.connection_indicator.setStyleSheet("""
                QLabel {
                    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                                               stop: 0 #10b981, stop: 1 #047857);
                    color: white;
                    border-radius: 12px;
                    padding: 8px;
                    margin: 10px;
                    font-size: 12px;
                    font-weight: bold;
                    border: 1px solid #047857;
                }
            """)
        else:
            self.connection_indicator.setText("🔴 Disconnected")
            self.connection_indicator.setStyleSheet("""
                QLabel {
                    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                                               stop: 0 #64748b, stop: 1 #475569);
                    color: white;
                    border-radius: 12px;
                    padding: 8px;
                    margin: 10px;
                    font-size: 12px;
                    font-weight: bold;
                    border: 1px solid #475569;
                }
            """)
        
        # Update setup page connection status
        self.setup_page.update_connection_status(connected)
    
    def update_recipients_count(self, count: int):
        """Update the recipients count display in the top bar."""
        self.recipients_count.setText(f"👥 Recipients: {count}")
        # Also update send page if currently viewing it
        if self.page_stack.currentIndex() == 3:
            self.update_send_page_stats()
    
    def on_files_changed(self):
        """Handle CSV file changes and refresh data."""
        self.refresh_recipients()
    
    def auto_refresh_content(self):
        """Automatically refresh content when files change."""
        if hasattr(self.setup_page, 'csv_file_edit'):
            current_csv = self.setup_page.csv_file_edit.text()
            
            if current_csv and os.path.exists(current_csv):
                # Check if CSV file has changed
                if (not hasattr(self, '_last_csv') or self._last_csv != current_csv):
                    self.refresh_recipients()
                    self._last_csv = current_csv
    
    def refresh_recipients(self):
        """Refresh recipients data from CSV file."""
        try:
            csv_file = self.setup_page.csv_file_edit.text()
            if not csv_file or not os.path.exists(csv_file):
                # Clear recipients if no file selected
                self.recipients_page.update_recipients([], "")
                self.compose_page.update_csv_columns([])
                self.update_recipients_count(0)
                return
            
            # Initialize email application if needed
            if not self.email_app:
                self.email_app = EmailApplication(log_level="INFO")
            
            # Read recipients from CSV
            self.recipients_data = self.email_app.csv_reader.read_recipients(csv_file)
            
            # Update recipients page
            self.recipients_page.update_recipients(self.recipients_data, os.path.basename(csv_file))
            
            # Update compose page with available columns for attributes
            if self.recipients_data:
                columns = list(self.recipients_data[0].keys())
                self.compose_page.update_csv_columns(columns)
            else:
                self.compose_page.update_csv_columns([])
            
            # Update recipient count
            self.update_recipients_count(len(self.recipients_data))
            self.log_message(f"✅ Recipients loaded: {len(self.recipients_data)} records", "SUCCESS")
            
        except Exception as e:
            # Handle errors gracefully
            self.recipients_page.update_recipients([], "")
            self.compose_page.update_csv_columns([])
            self.log_message(f"❌ Error loading recipients: {str(e)}", "ERROR")
            self.update_recipients_count(0)
    
    def test_smtp_connection(self):
        """Test SMTP connection configuration."""
        try:
            config = self.setup_page.get_config()
            smtp_config = config['smtp']
            
            # Validate SMTP settings
            if not smtp_config['server'] or not smtp_config['username'] or not smtp_config['password']:
                self.setup_page.update_connection_status(False, "Please complete all SMTP settings")
                return
            
            # Initialize email application
            if not self.email_app:
                self.email_app = EmailApplication(log_level="INFO")
            
            # Setup configuration
            self.email_app.config = config
            self.email_app.config_manager.config = config
            self.email_app.config_manager.is_loaded = True
            
            # Start test operation
            self.worker = EmailWorker(self.email_app, 'test', test_recipient=smtp_config['username'])
            self.worker.status_updated.connect(self.update_status)
            self.worker.finished.connect(self.test_finished)
            self.worker.error_occurred.connect(self.worker_error)
            self.worker.start()
            
        except Exception as e:
            self.setup_page.update_connection_status(False, f"Test failed: {str(e)}")
    
    def send_test_email(self):
        """Send a test email to verify configuration."""
        try:
            if not self.validate_configuration():
                self.send_page.test_btn.stop_loading()
                return
            
            config = self.setup_page.get_config()
            test_recipient = config['email']['test_email']
            
            # Start test email operation
            self.worker = EmailWorker(self.email_app, 'test', test_recipient=test_recipient)
            self.worker.status_updated.connect(self.update_status)
            self.worker.finished.connect(self.test_email_finished)
            self.worker.error_occurred.connect(self.worker_error)
            self.worker.start()
            
        except Exception as e:
            self.show_error("Send Test Error", f"Failed to send test email: {str(e)}")
            self.send_page.test_btn.stop_loading()
    
    def send_bulk_emails(self):
        """Send bulk emails to all recipients."""
        try:
            if not self.validate_configuration():
                self.send_page.send_all_btn.stop_loading()
                return
            
            if not self.recipients_data:
                self.show_warning("No Recipients", "Please load recipients from CSV file first.")
                self.send_page.send_all_btn.stop_loading()
                return
            
            # Get email content and settings
            subject = self.compose_page.subject_edit.text()
            content = self.compose_page.email_content_edit.toPlainText()
            attachments = self.compose_page.get_attachments_list()
            
            # Validate required fields
            if not subject:
                self.show_warning("Missing Subject", "Please enter an email subject.")
                self.send_page.send_all_btn.stop_loading()
                return
            
            if not content:
                self.show_warning("Missing Content", "Please write your email content.")
                self.send_page.send_all_btn.stop_loading()
                return
            
            # Get selected email column
            email_column = self.recipients_page.get_selected_email_column()
            if not email_column:
                self.show_warning("No Email Column", "Please select the email column from recipients.")
                self.send_page.send_all_btn.stop_loading()
                return
            
            # Update recipients data with selected email column
            updated_recipients = []
            for recipient in self.recipients_data:
                updated_recipient = recipient.copy()
                updated_recipient['email'] = recipient.get(email_column, '')
                updated_recipients.append(updated_recipient)
            
            # Confirmation dialog
            reply = QMessageBox.question(
                self, "Confirm Campaign Launch",
                f"🚀 Ready to send emails to {len(updated_recipients)} recipients?\n\n"
                f"📧 Subject: {subject}\n"
                f"📎 Attachments: {len(attachments)}\n"
                f"📬 Email Column: {email_column}\n\n"
                f"This action cannot be undone. Continue?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply != QMessageBox.Yes:
                self.send_page.send_all_btn.stop_loading()
                return
            
            # Start bulk email operation
            self.worker = EmailWorker(
                self.email_app, 'bulk',
                recipients_data=updated_recipients,
                template_content=content,
                subject=subject,
                attachments=attachments
            )
            
            # Connect worker signals
            self.worker.progress_updated.connect(self.update_progress)
            self.worker.status_updated.connect(self.update_status)
            self.worker.email_sent.connect(self.email_sent)
            self.worker.finished.connect(self.bulk_finished)
            self.worker.error_occurred.connect(self.worker_error)
            self.worker.start()
            
            # Update UI for sending state
            self.send_page.update_ui_for_sending(True)
            self.log_message("🚀 Starting email campaign...", "INFO")
            
        except Exception as e:
            self.show_error("Send Error", f"Failed to start email campaign: {str(e)}")
            self.send_page.send_all_btn.stop_loading()
    
    def stop_sending(self):
        """Stop the current email sending operation."""
        if self.worker:
            self.worker.stop()
            self.update_status("⏹️ Stopping email campaign...")
            self.log_message("⏹️ Email campaign stopped by user", "WARNING")
    
    def validate_configuration(self):
        """Validate configuration before sending emails."""
        config = self.setup_page.get_config()
        
        # Check SMTP configuration
        smtp = config['smtp']
        if not smtp['server'] or not smtp['username'] or not smtp['password']:
            self.show_warning("Incomplete SMTP Configuration", "Please complete SMTP settings in the Setup tab.")
            return False
        
        # Check recipients
        if not self.recipients_data:
            self.show_warning("No Recipients", "Please load recipients from CSV file.")
            return False
        
        # Initialize email application if needed
        if not self.email_app:
            self.email_app = EmailApplication(log_level="INFO")
        
        # Setup configuration
        self.email_app.config = config
        self.email_app.config_manager.config = config
        self.email_app.config_manager.is_loaded = True
        
        return True
    
    def update_progress(self, current, total):
        """Update progress bar during bulk email sending."""
        self.send_page.progress_bar.setMaximum(total)
        self.send_page.progress_bar.setValue(current)
        self.send_page.update_progress_info(current, total)
    
    def update_status(self, message):
        """Update status message during operations."""
        self.send_page.status_label.setText(message)
        self.log_message(message, "INFO")
    
    def test_finished(self, results):
        """Handle SMTP test completion."""
        if results['success']:
            self.update_connection_status(True)
            self.setup_page.update_connection_status(True, "SMTP connection successful")
            self.log_message("✅ SMTP connection test successful", "SUCCESS")
        else:
            self.update_connection_status(False)
            error_msg = results.get('message', 'Unknown error')
            self.setup_page.update_connection_status(False, f"SMTP connection failed: {error_msg}")
            self.log_message(f"❌ SMTP connection test failed: {error_msg}", "ERROR")
    
    def test_email_finished(self, results):
        """Handle test email completion."""
        if results['success']:
            self.show_info("🧪 Test Email Successful", "✅ Test email sent successfully!\n\nYour configuration is working correctly.")
        else:
            self.show_error("🧪 Test Email Failed", f"❌ Test email failed: {results.get('message', 'Unknown error')}\n\nPlease check your settings.")
        
        self.send_page.test_btn.stop_loading()
        self.send_page.status_label.setText("Ready to launch campaign")
    
    def bulk_finished(self, results):
        """Handle bulk email campaign completion."""
        # Reset UI state
        self.send_page.update_ui_for_sending(False)
        self.send_page.progress_bar.setValue(0)
        self.send_page.update_progress_info(0, 1)
        
        # Show campaign results
        self.send_page.show_campaign_results(results)
        
        # Calculate success rate
        success_rate = int((results['sent'] / results['total']) * 100) if results['total'] > 0 else 0
        
        # Create results message
        message = f"📊 Campaign Results\n\n"
        message += f"📧 Total Recipients: {results['total']}\n"
        message += f"✅ Successfully Sent: {results['sent']}\n"
        message += f"❌ Failed: {results['failed']}\n"
        message += f"📈 Success Rate: {success_rate}%"
        
        # Add error information if any
        if results['errors']:
            message += f"\n\n⚠️ Errors occurred:"
            for error in results['errors'][:3]:  # Show first 3 errors
                message += f"\n• {error}"
            if len(results['errors']) > 3:
                message += f"\n... and {len(results['errors']) - 3} more errors (see logs)"
        
        # Show appropriate result dialog
        if results['success']:
            self.show_info("🎉 Campaign Complete", message)
        else:
            self.show_warning("⚠️ Campaign Complete with Issues", message)
        
        self.send_page.status_label.setText("Campaign completed successfully")
        self.log_message(f"📊 Email campaign completed: {results['sent']}/{results['total']} sent successfully", "INFO")
    
    def email_sent(self, recipient, success, error):
        """Handle individual email send results."""
        if success:
            self.log_message(f"✅ Email sent to {recipient}", "SUCCESS")
        else:
            self.log_message(f"❌ Failed to send to {recipient}: {error}", "ERROR")
    
    def worker_error(self, error):
        """Handle worker thread errors."""
        self.show_error("⚠️ Operation Error", f"An error occurred during the operation:\n\n{error}")
        self.send_page.status_label.setText("Ready to launch campaign")
        self.send_page.update_ui_for_sending(False)
        self.send_page.test_btn.stop_loading()
        self.send_page.send_all_btn.stop_loading()
        self.setup_page.update_connection_status(False, "Operation failed")
        self.update_connection_status(False)
    
    def log_message(self, message, level="INFO"):
        """Add message to activity log with formatting and white font."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Define level styling
        level_info = {
            "INFO": ("ℹ️", "#3b82f6"),
            "SUCCESS": ("✅", "#10b981"),
            "ERROR": ("❌", "#ef4444"),
            "WARNING": ("⚠️", "#f59e0b")
        }
        
        emoji, color = level_info.get(level, ("📝", "#64748b"))
        
        # Create HTML-formatted message with white font
        formatted_message = f'''
        <div style="margin: 5px 0; padding: 8px; border-left: 3px solid {color}; background: rgba(59, 130, 246, 0.05);">
            <span style="color: {color}; font-weight: bold;">[{timestamp}] {emoji} {level}:</span>
            <span style="color: #ffffff; margin-left: 10px;">{message}</span>
        </div>
        '''
        
        self.log_page.log_display.append(formatted_message)
        
        # Auto-scroll to bottom
        scrollbar = self.log_page.log_display.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def show_error(self, title, message):
        """Show error message dialog with modern styling."""
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setStyleSheet("""
            QMessageBox {
                background-color: white;
                border-radius: 8px;
                min-width: 350px;
            }
            QMessageBox QPushButton {
                background-color: #ef4444;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
                min-width: 70px;
            }
            QMessageBox QPushButton:hover {
                background-color: #dc2626;
            }
        """)
        msg_box.exec_()
        self.log_message(f"❌ {message}", "ERROR")
    
    def show_info(self, title, message):
        """Show information message dialog with modern styling."""
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setStyleSheet("""
            QMessageBox {
                background-color: white;
                border-radius: 8px;
                min-width: 350px;
            }
            QMessageBox QPushButton {
                background-color: #3b82f6;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
                min-width: 70px;
            }
            QMessageBox QPushButton:hover {
                background-color: #2563eb;
            }
        """)
        msg_box.exec_()
        self.log_message(f"ℹ️ {message}", "INFO")
    
    def show_warning(self, title, message):
        """Show warning message dialog with modern styling."""
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Warning)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setStyleSheet("""
            QMessageBox {
                background-color: white;
                border-radius: 8px;
                min-width: 350px;
            }
            QMessageBox QPushButton {
                background-color: #f59e0b;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
                min-width: 70px;
            }
            QMessageBox QPushButton:hover {
                background-color: #d97706;
            }
        """)
        msg_box.exec_()
        self.log_message(f"⚠️ {message}", "WARNING")
    
    def load_config(self):
        """Load configuration from file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "🔧 Load Configuration", "", 
            "Configuration Files (*.json);;All Files (*)"
        )
        if file_path:
            self.load_config_file(file_path)
    
    def load_config_file(self, file_path):
        """Load configuration from specified file."""
        try:
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # Apply configuration to UI
                self.setup_page.set_config(config)
                self.config = config
                
                # Refresh data
                self.on_files_changed()
                
                self.show_info("✅ Configuration Loaded", 
                              f"Configuration successfully loaded from:\n{os.path.basename(file_path)}")
                self.log_message(f"Configuration loaded from {file_path}", "SUCCESS")
                
        except Exception as e:
            self.show_error("❌ Load Error", f"Failed to load configuration:\n\n{str(e)}")
    
    def save_config(self):
        """Save current configuration to file."""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "💾 Save Configuration", "email_config.json", 
            "Configuration Files (*.json);;All Files (*)"
        )
        if file_path:
            try:
                config = self.setup_page.get_config()
                
                # Add metadata
                config['metadata'] = {
                    'version': '1.0.0',
                    'created_by': 'ByeBye21',
                    'created_at': datetime.now().isoformat(),
                    'application': 'Email Automation'
                }
                
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=2, ensure_ascii=False)
                
                self.show_info("💾 Configuration Saved", 
                              f"Configuration successfully saved to:\n{os.path.basename(file_path)}")
                self.log_message(f"Configuration saved to {file_path}", "SUCCESS")
                
            except Exception as e:
                self.show_error("❌ Save Error", f"Failed to save configuration:\n\n{str(e)}")
    
    def clear_log(self):
        """Clear the activity log with confirmation."""
        reply = QMessageBox.question(
            self, "Clear Activity Log",
            "Are you sure you want to clear the activity log?\n\nThis action cannot be undone.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.log_page.log_display.clear()
            self.log_message("🗑️ Activity log cleared by user", "INFO")
    
    def save_log(self):
        """Export activity log to HTML file."""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "💾 Export Activity Log", 
            f"email_automation_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html", 
            "HTML Files (*.html);;Text Files (*.txt);;All Files (*)"
        )
        if file_path:
            try:
                # Create comprehensive HTML log
                html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Email Automation Activity Log</title>
    <style>
        body {{ 
            font-family: 'Segoe UI', Arial, sans-serif; 
            margin: 20px; 
            background: #1e293b; 
            color: white; 
            line-height: 1.6;
        }}
        .header {{ 
            color: #ffffff; 
            border-bottom: 2px solid #334155; 
            padding-bottom: 15px; 
            margin-bottom: 20px;
        }}
        .header h1 {{
            margin: 0;
            color: #3b82f6;
        }}
        .header .meta {{
            font-size: 14px;
            color: #cbd5e1;
            margin-top: 10px;
        }}
        .log-content {{ 
            margin-top: 20px; 
            background: #1e293b;
            padding: 15px;
            border-radius: 8px;
        }}
        .footer {{
            margin-top: 30px;
            padding-top: 15px;
            border-top: 1px solid #334155;
            font-size: 12px;
            color: #94a3b8;
            text-align: center;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📧 Email Automation Activity Log</h1>
        <div class="meta">
            <strong>Generated:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}<br>
            <strong>Application:</strong> Email Automation v1.0.0<br>
            <strong>Developer:</strong> ByeBye21
        </div>
    </div>
    <div class="log-content">
        {self.log_page.log_display.toHtml()}
    </div>
    <div class="footer">
        Email Automation v1.0.0 | Created by ByeBye21 | © 2025 All rights reserved<br>
        GitHub: <a href="https://github.com/ByeBye21" style="color: #3b82f6;">https://github.com/ByeBye21</a> | 
        Email: <a href="mailto:younes0079@gmail.com" style="color: #3b82f6;">younes0079@gmail.com</a>
    </div>
</body>
</html>
                """
                
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(html_content)
                
                self.show_info("💾 Log Exported", f"Activity log exported to:\n{os.path.basename(file_path)}")
                self.log_message(f"Activity log exported to {file_path}", "SUCCESS")
                
            except Exception as e:
                self.show_error("❌ Export Error", f"Failed to export log:\n\n{str(e)}")
    
    def show_about(self):
        """Show comprehensive about dialog."""
        about_text = """
        <div style="text-align: center; padding: 25px; line-height: 1.6;">
            <h2 style="color: #1e293b; margin-bottom: 5px;">📧 Email Automation</h2>
            <p style="color: #64748b; font-size: 16px; margin-bottom: 25px;"><strong>Version 1.0.0</strong></p>
            
            <div style="text-align: left; max-width: 500px; margin: 0 auto;">
                <h3 style="color: #3b82f6; margin-bottom: 15px;">✨ Features</h3>
                <ul style="color: #4b5563; line-height: 1.8; margin-bottom: 20px;">
                    <li><strong>🔧 Smart SMTP Setup</strong> - Automatic provider detection with secure password handling</li>
                    <li><strong>📊 CSV Management</strong> - Intelligent recipient loading with email column selection</li>
                    <li><strong>✏️ Rich Email Composer</strong> - Built-in editor with dynamic attribute insertion</li>
                    <li><strong>📎 File Attachments</strong> - Drag-and-drop attachment support with size validation</li>
                    <li><strong>🚀 Campaign Management</strong> - Real-time progress tracking and bulk sending</li>
                    <li><strong>📋 Activity Logging</strong> - Comprehensive logging with export functionality</li>
                    <li><strong>🎨 Modern Interface</strong> - Responsive design with professional styling</li>
                    <li><strong>⚡ Performance</strong> - Multi-threaded operations with smooth UI</li>
                </ul>
                
                <h3 style="color: #10b981; margin: 20px 0 15px 0;">🛠️ Technology Stack</h3>
                <p style="color: #4b5563; margin-bottom: 20px;">
                    Built with <strong>Python</strong>, <strong>PyQt5</strong>, and <strong>modern design principles</strong>
                </p>
                
                <h3 style="color: #f59e0b; margin: 20px 0 15px 0;">👨‍💻 Developer</h3>
                <p style="color: #4b5563; margin-bottom: 20px;">
                    <strong>ByeBye21</strong><br>
                    Created with ❤️ for efficient email automation
                </p>
                
                <h3 style="color: #ef4444; margin: 20px 0 15px 0;">📞 Contact & Support</h3>
                <p style="color: #4b5563; margin-bottom: 20px;">
                    <strong>Email:</strong> <a href="mailto:younes0079@gmail.com" style="color: #3b82f6; text-decoration: none;">younes0079@gmail.com</a><br>
                    <strong>GitHub:</strong> <a href="https://github.com/ByeBye21" style="color: #3b82f6; text-decoration: none;">github.com/ByeBye21</a>
                </p>
                
                <div style="text-align: center; margin-top: 25px; padding-top: 20px; border-top: 1px solid #e2e8f0;">
                    <p style="color: #94a3b8; font-size: 12px; margin: 0;">
                        © 2025 ByeBye21. All rights reserved.<br>
                        Email Automation v1.0.0
                    </p>
                </div>
            </div>
        </div>
        """
        
        about_dialog = QMessageBox(self)
        about_dialog.setWindowTitle("About Email Automation")
        about_dialog.setText(about_text)
        about_dialog.setStyleSheet("""
            QMessageBox {
                background-color: white;
                border-radius: 12px;
                min-width: 600px;
                min-height: 500px;
            }
            QMessageBox QPushButton {
                background-color: #3b82f6;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 12px 24px;
                font-weight: bold;
                min-width: 100px;
            }
            QMessageBox QPushButton:hover {
                background-color: #2563eb;
            }
        """)
        about_dialog.exec_()
    
    def closeEvent(self, event):
        """Handle application close event with proper cleanup."""
        # Check if email operation is in progress
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self, "🚪 Confirm Exit",
                "⚠️ Email operation in progress!\n\nAre you sure you want to exit?\nThis will stop the current operation.",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.worker.stop()
                self.worker.wait(3000)  # Wait up to 3 seconds
                event.accept()
            else:
                event.ignore()
        else:
            # Save window geometry for next session
            try:
                settings = {
                    'geometry': {
                        'x': self.x(),
                        'y': self.y(),
                        'width': self.width(),
                        'height': self.height()
                    },
                    'last_saved': datetime.now().isoformat()
                }
                
                with open('app_settings.json', 'w', encoding='utf-8') as f:
                    json.dump(settings, f, indent=2)
                    
            except Exception:
                pass  # Ignore settings save errors
            
            event.accept()


def main():
    """
    Main application entry point.
    
    Initializes the Qt application, creates the main window,
    and starts the event loop.
    """
    # Create Qt application
    app = QApplication(sys.argv)
    
    # Set application metadata
    app.setApplicationName("Email Automation")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("ByeBye21")
    app.setOrganizationDomain("github.com/ByeBye21")
    
    # Set application icon if available
    try:
        icon_path = "app_icon.png"
        if os.path.exists(icon_path):
            app.setWindowIcon(QIcon(icon_path))
    except Exception:
        pass  # Ignore icon loading errors
    
    # Create main application window
    window = EmailAutomationApp()
    
    # Restore window geometry if available
    try:
        with open('app_settings.json', 'r', encoding='utf-8') as f:
            settings = json.load(f)
        
        geometry = settings.get('geometry', {})
        if geometry:
            window.resize(
                geometry.get('width', 1400),
                geometry.get('height', 900)
            )
            window.move(
                geometry.get('x', 100),
                geometry.get('y', 100)
            )
    except Exception:
        pass  # Use default geometry if settings not found
    
    # Show window
    window.show()
    
    # Log startup
    window.log_message("🚀 Email Automation v1.0.0 started successfully", "SUCCESS")
    window.log_message("Created by ByeBye21 | GitHub: https://github.com/ByeBye21", "INFO")
    
    # Start application event loop
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

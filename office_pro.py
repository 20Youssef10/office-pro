#!/usr/bin/env python3
"""
Office Pro - Professional Office Suite
A comprehensive office application suite with support for DOCX, XLSX, PPTX, PDF and more.
Author: Office Pro Team
Version: 1.0.0
"""

import sys
import os
from pathlib import Path

# Add modules to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from PyQt6.QtWidgets import (
        QApplication,
        QMainWindow,
        QWidget,
        QVBoxLayout,
        QHBoxLayout,
        QPushButton,
        QLabel,
        QStackedWidget,
        QFileDialog,
        QMessageBox,
        QMenuBar,
        QToolBar,
        QStatusBar,
        QFrame,
        QSizePolicy,
    )
    from PyQt6.QtCore import Qt, QSize, QTimer
    from PyQt6.QtGui import QIcon, QFont, QAction, QPalette, QColor

    PYQT_AVAILABLE = True
except ImportError:
    PYQT_AVAILABLE = False
    print("PyQt6 not available. Installing required packages...")
    import subprocess

    subprocess.check_call([sys.executable, "-m", "pip", "install", "PyQt6", "-q"])

    from PyQt6.QtWidgets import (
        QApplication,
        QMainWindow,
        QWidget,
        QVBoxLayout,
        QHBoxLayout,
        QPushButton,
        QLabel,
        QStackedWidget,
        QFileDialog,
        QMessageBox,
        QMenuBar,
        QToolBar,
        QStatusBar,
        QFrame,
        QSizePolicy,
    )
    from PyQt6.QtCore import Qt, QSize, QTimer
    from PyQt6.QtGui import QIcon, QFont, QAction, QPalette, QColor

# Import modules
from modules.word_processor import WordProcessor
from modules.spreadsheet import SpreadsheetEditor
from modules.presentation import PresentationEditor
from modules.pdf_editor import PDFEditor
from modules.file_manager import FileManager


class ModernButton(QPushButton):
    """Custom styled button for the main menu"""

    def __init__(self, text, icon_text, parent=None):
        super().__init__(text, parent)
        self.setMinimumSize(200, 180)
        self.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        self.setStyleSheet("""
            QPushButton {
                background-color: #2b579a;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 20px;
                text-align: center;
            }
            QPushButton:hover {
                background-color: #1e3f6f;
            }
            QPushButton:pressed {
                background-color: #16325c;
            }
        """)


class MainMenu(QWidget):
    """Main menu widget with application selection"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(30)
        layout.setContentsMargins(50, 50, 50, 50)

        # Header
        header = QLabel("Office Pro")
        header.setFont(QFont("Segoe UI", 48, QFont.Weight.Bold))
        header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header.setStyleSheet("color: #2b579a; margin-bottom: 20px;")
        layout.addWidget(header)

        # Subtitle
        subtitle = QLabel("Professional Office Suite")
        subtitle.setFont(QFont("Segoe UI", 16))
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("color: #666; margin-bottom: 30px;")
        layout.addWidget(subtitle)

        # Buttons grid
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(30)

        # Word Processor Button
        self.btn_word = ModernButton("Word\nProcessor", "W")
        self.btn_word.clicked.connect(lambda: self.parent.open_module("word"))
        buttons_layout.addWidget(self.btn_word)

        # Spreadsheet Button
        self.btn_spreadsheet = ModernButton("Spreadsheet", "S")
        self.btn_spreadsheet.setStyleSheet("""
            QPushButton {
                background-color: #217346;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 20px;
                text-align: center;
            }
            QPushButton:hover {
                background-color: #1a5c38;
            }
            QPushButton:pressed {
                background-color: #154a2d;
            }
        """)
        self.btn_spreadsheet.clicked.connect(
            lambda: self.parent.open_module("spreadsheet")
        )
        buttons_layout.addWidget(self.btn_spreadsheet)

        # Presentation Button
        self.btn_presentation = ModernButton("Presentation", "P")
        self.btn_presentation.setStyleSheet("""
            QPushButton {
                background-color: #d24726;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 20px;
                text-align: center;
            }
            QPushButton:hover {
                background-color: #b33d20;
            }
            QPushButton:pressed {
                background-color: #9c351c;
            }
        """)
        self.btn_presentation.clicked.connect(
            lambda: self.parent.open_module("presentation")
        )
        buttons_layout.addWidget(self.btn_presentation)

        # PDF Editor Button
        self.btn_pdf = ModernButton("PDF\nEditor", "P")
        self.btn_pdf.setStyleSheet("""
            QPushButton {
                background-color: #f40f02;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 20px;
                text-align: center;
            }
            QPushButton:hover {
                background-color: #cc0c02;
            }
            QPushButton:pressed {
                background-color: #b30a01;
            }
        """)
        self.btn_pdf.clicked.connect(lambda: self.parent.open_module("pdf"))
        buttons_layout.addWidget(self.btn_pdf)

        layout.addLayout(buttons_layout)

        # Recent files section
        recent_label = QLabel("Recent Files")
        recent_label.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        recent_label.setStyleSheet("color: #333; margin-top: 30px;")
        layout.addWidget(recent_label)

        self.recent_files_widget = QFrame()
        self.recent_files_widget.setStyleSheet("""
            QFrame {
                background-color: #f5f5f5;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        recent_layout = QVBoxLayout(self.recent_files_widget)

        recent_placeholder = QLabel(
            "No recent files. Create or open a document to get started."
        )
        recent_placeholder.setFont(QFont("Segoe UI", 12))
        recent_placeholder.setStyleSheet("color: #999;")
        recent_layout.addWidget(recent_placeholder)

        layout.addWidget(self.recent_files_widget)

        # Footer
        footer = QLabel("© 2026 Office Pro. All rights reserved.")
        footer.setFont(QFont("Segoe UI", 10))
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer.setStyleSheet("color: #999; margin-top: 20px;")
        layout.addWidget(footer)

        layout.addStretch()


class OfficePro(QMainWindow):
    """Main application window"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Office Pro")
        self.setGeometry(100, 100, 1400, 900)

        # Initialize file manager
        self.file_manager = FileManager()

        # Initialize modules
        self.modules = {}
        self.current_module = None

        # Setup UI
        self.setup_ui()

    def setup_ui(self):
        """Setup the main user interface"""
        # Central widget
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # Main layout
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)

        # Stacked widget for different views
        self.stack = QStackedWidget()
        self.main_layout.addWidget(self.stack)

        # Main menu
        self.main_menu = MainMenu(self)
        self.stack.addWidget(self.main_menu)

        # Initialize modules (but don't show yet)
        self.word_processor = WordProcessor(self)
        self.spreadsheet_editor = SpreadsheetEditor(self)
        self.presentation_editor = PresentationEditor(self)
        self.pdf_editor = PDFEditor(self)

        # Add modules to stack
        self.stack.addWidget(self.word_processor)
        self.stack.addWidget(self.spreadsheet_editor)
        self.stack.addWidget(self.presentation_editor)
        self.stack.addWidget(self.pdf_editor)

        # Menu bar
        self.create_menu_bar()

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")

    def create_menu_bar(self):
        """Create the application menu bar"""
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu("&File")

        new_menu = file_menu.addMenu("&New")
        new_word = QAction("Word Document", self)
        new_word.triggered.connect(lambda: self.open_module("word"))
        new_menu.addAction(new_word)

        new_spreadsheet = QAction("Spreadsheet", self)
        new_spreadsheet.triggered.connect(lambda: self.open_module("spreadsheet"))
        new_menu.addAction(new_spreadsheet)

        new_presentation = QAction("Presentation", self)
        new_presentation.triggered.connect(lambda: self.open_module("presentation"))
        new_menu.addAction(new_presentation)

        file_menu.addSeparator()

        open_action = QAction("&Open...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)

        save_action = QAction("&Save", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)

        file_menu.addSeparator()

        exit_action = QAction("E&xit", self)
        exit_action.setShortcut("Alt+F4")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # View menu
        view_menu = menubar.addMenu("&View")

        home_action = QAction("&Home", self)
        home_action.triggered.connect(self.show_main_menu)
        view_menu.addAction(home_action)

        # Help menu
        help_menu = menubar.addMenu("&Help")

        about_action = QAction("&About", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def open_module(self, module_name):
        """Open a specific office module"""
        if module_name == "word":
            self.current_module = self.word_processor
            self.stack.setCurrentWidget(self.word_processor)
            self.setWindowTitle("Office Pro - Word Processor")
        elif module_name == "spreadsheet":
            self.current_module = self.spreadsheet_editor
            self.stack.setCurrentWidget(self.spreadsheet_editor)
            self.setWindowTitle("Office Pro - Spreadsheet")
        elif module_name == "presentation":
            self.current_module = self.presentation_editor
            self.stack.setCurrentWidget(self.presentation_editor)
            self.setWindowTitle("Office Pro - Presentation")
        elif module_name == "pdf":
            self.current_module = self.pdf_editor
            self.stack.setCurrentWidget(self.pdf_editor)
            self.setWindowTitle("Office Pro - PDF Editor")

    def show_main_menu(self):
        """Return to main menu"""
        self.current_module = None
        self.stack.setCurrentWidget(self.main_menu)
        self.setWindowTitle("Office Pro")

    def open_file(self):
        """Open a file dialog"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open File",
            "",
            "All Supported Files (*.docx *.doc *.xlsx *.xls *.pptx *.ppt *.pdf *.txt *.rtf *.odt *.ods *.odp);;"
            "Word Documents (*.docx *.doc *.odt);;"
            "Spreadsheets (*.xlsx *.xls *.ods *.csv);;"
            "Presentations (*.pptx *.ppt *.odp);;"
            "PDF Files (*.pdf);;"
            "Text Files (*.txt *.rtf);;"
            "All Files (*)",
        )

        if file_path:
            self.load_file(file_path)

    def load_file(self, file_path):
        """Load a file into the appropriate module"""
        ext = Path(file_path).suffix.lower()

        if ext in [".docx", ".doc", ".odt", ".rtf"]:
            self.open_module("word")
            self.word_processor.open_document(file_path)
        elif ext in [".xlsx", ".xls", ".ods", ".csv"]:
            self.open_module("spreadsheet")
            self.spreadsheet_editor.open_spreadsheet(file_path)
        elif ext in [".pptx", ".ppt", ".odp"]:
            self.open_module("presentation")
            self.presentation_editor.open_presentation(file_path)
        elif ext == ".pdf":
            self.open_module("pdf")
            self.pdf_editor.open_pdf(file_path)
        elif ext == ".txt":
            self.open_module("word")
            self.word_processor.open_document(file_path)
        else:
            QMessageBox.warning(
                self, "Unsupported File", f"The file type '{ext}' is not supported yet."
            )

    def save_file(self):
        """Save the current file"""
        if self.current_module:
            self.current_module.save()
        else:
            QMessageBox.information(self, "No File", "No file is currently open.")

    def show_about(self):
        """Show about dialog"""
        QMessageBox.about(
            self,
            "About Office Pro",
            """<h2>Office Pro 1.0</h2>
                         <p>A professional office suite built with Python.</p>
                         <p>Features:</p>
                         <ul>
                         <li>Word Processing (DOCX, DOC, ODT, RTF)</li>
                         <li>Spreadsheets (XLSX, XLS, ODS, CSV)</li>
                         <li>Presentations (PPTX, PPT, ODP)</li>
                         <li>PDF Editor (View, Annotate, Edit)</li>
                         <li>Text Files (TXT)</li>
                         </ul>
                         <p>Built with PyQt6, Python-docx, OpenPyXL, python-pptx, and PyMuPDF.</p>
                         <p>© 2026 Office Pro Team</p>""",
        )


def main():
    """Main entry point"""
    app = QApplication(sys.argv)

    # Set application style
    app.setStyle("Fusion")

    # Set palette for modern look
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(0, 0, 0))
    palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(240, 240, 240))
    palette.setColor(QPalette.ColorRole.Text, QColor(0, 0, 0))
    palette.setColor(QPalette.ColorRole.Button, QColor(240, 240, 240))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(0, 0, 0))
    app.setPalette(palette)

    # Create and show main window
    window = OfficePro()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()

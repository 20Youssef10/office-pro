"""
Office Pro - Tools Dialog (Features #41-50)
Integrated tools: Clipboard Manager, QR Generator, Unit Converter, etc.
"""

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QTabWidget,
    QLabel,
    QPushButton,
    QLineEdit,
    QTextEdit,
    QComboBox,
    QSpinBox,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QGroupBox,
    QGridLayout,
    QWidget,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
import random
import string


class ToolsDialog(QDialog):
    """Integrated tools panel with multiple utilities"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🛠️ Office Pro Tools")
        self.setMinimumSize(800, 600)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Header
        header = QLabel("Professional Tools & Utilities")
        header.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        layout.addWidget(header)

        # Tools tabs
        tabs = QTabWidget()

        # QR Code Generator (#43)
        qr_tab = self.create_qr_tab()
        tabs.addTab(qr_tab, "📱 QR Codes")

        # Unit Converter (#45)
        converter_tab = self.create_converter_tab()
        tabs.addTab(converter_tab, "🔄 Converter")

        # Password Generator
        password_tab = self.create_password_tab()
        tabs.addTab(password_tab, "🔐 Passwords")

        # Text Tools
        text_tab = self.create_text_tab()
        tabs.addTab(text_tab, "📝 Text Tools")

        # Clipboard History (#41)
        clipboard_tab = self.create_clipboard_tab()
        tabs.addTab(clipboard_tab, "📋 Clipboard")

        layout.addWidget(tabs)

        # Close button
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def create_qr_tab(self):
        """QR Code Generator tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Input
        input_group = QGroupBox("QR Code Content")
        input_layout = QVBoxLayout(input_group)

        self.qr_input = QTextEdit()
        self.qr_input.setPlaceholderText(
            "Enter text, URL, or contact info to encode..."
        )
        self.qr_input.setMaximumHeight(100)
        input_layout.addWidget(self.qr_input)

        # Type selector
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("Type:"))
        self.qr_type = QComboBox()
        self.qr_type.addItems(["Text", "URL", "Email", "Phone", "WiFi", "Contact"])
        type_layout.addWidget(self.qr_type)
        type_layout.addStretch()
        input_layout.addLayout(type_layout)

        layout.addWidget(input_group)

        # Preview
        preview_group = QGroupBox("QR Code Preview")
        preview_layout = QVBoxLayout(preview_group)

        self.qr_preview = QLabel("QR Code will appear here")
        self.qr_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.qr_preview.setMinimumHeight(200)
        self.qr_preview.setStyleSheet(
            "background-color: #f0f0f0; border: 2px dashed #ccc;"
        )
        preview_layout.addWidget(self.qr_preview)

        layout.addWidget(preview_group)

        # Actions
        btn_layout = QHBoxLayout()

        generate_btn = QPushButton("Generate QR Code")
        generate_btn.clicked.connect(self.generate_qr)
        btn_layout.addWidget(generate_btn)

        save_btn = QPushButton("Save as Image")
        save_btn.clicked.connect(self.save_qr)
        btn_layout.addWidget(save_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        layout.addStretch()
        return widget

    def create_converter_tab(self):
        """Unit Converter tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Category
        cat_layout = QHBoxLayout()
        cat_layout.addWidget(QLabel("Category:"))
        self.conv_category = QComboBox()
        self.conv_category.addItems(
            [
                "Length",
                "Weight",
                "Temperature",
                "Volume",
                "Area",
                "Speed",
                "Data",
                "Currency",
            ]
        )
        self.conv_category.currentTextChanged.connect(self.update_converter_units)
        cat_layout.addWidget(self.conv_category)
        cat_layout.addStretch()
        layout.addLayout(cat_layout)

        # Converter
        conv_group = QGroupBox("Convert")
        conv_layout = QGridLayout(conv_group)

        # From
        conv_layout.addWidget(QLabel("From:"), 0, 0)
        self.conv_from_value = QLineEdit("1")
        conv_layout.addWidget(self.conv_from_value, 0, 1)
        self.conv_from_unit = QComboBox()
        conv_layout.addWidget(self.conv_from_unit, 0, 2)

        # To
        conv_layout.addWidget(QLabel("To:"), 1, 0)
        self.conv_to_value = QLineEdit()
        self.conv_to_value.setReadOnly(True)
        conv_layout.addWidget(self.conv_to_value, 1, 1)
        self.conv_to_unit = QComboBox()
        conv_layout.addWidget(self.conv_to_unit, 1, 2)

        # Convert button
        convert_btn = QPushButton("Convert")
        convert_btn.clicked.connect(self.perform_conversion)
        conv_layout.addWidget(convert_btn, 2, 1)

        layout.addWidget(conv_group)

        # Initialize units
        self.update_converter_units("Length")

        layout.addStretch()
        return widget

    def create_password_tab(self):
        """Password Generator tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Options
        options_group = QGroupBox("Password Options")
        options_layout = QGridLayout(options_group)

        # Length
        options_layout.addWidget(QLabel("Length:"), 0, 0)
        self.pwd_length = QSpinBox()
        self.pwd_length.setRange(8, 64)
        self.pwd_length.setValue(16)
        options_layout.addWidget(self.pwd_length, 0, 1)

        # Character types
        self.pwd_uppercase = QPushButton("✓ Uppercase (A-Z)")
        self.pwd_uppercase.setCheckable(True)
        self.pwd_uppercase.setChecked(True)
        options_layout.addWidget(self.pwd_uppercase, 1, 0, 1, 2)

        self.pwd_lowercase = QPushButton("✓ Lowercase (a-z)")
        self.pwd_lowercase.setCheckable(True)
        self.pwd_lowercase.setChecked(True)
        options_layout.addWidget(self.pwd_lowercase, 2, 0, 1, 2)

        self.pwd_numbers = QPushButton("✓ Numbers (0-9)")
        self.pwd_numbers.setCheckable(True)
        self.pwd_numbers.setChecked(True)
        options_layout.addWidget(self.pwd_numbers, 3, 0, 1, 2)

        self.pwd_symbols = QPushButton("✓ Symbols (!@#$%)")
        self.pwd_symbols.setCheckable(True)
        self.pwd_symbols.setChecked(True)
        options_layout.addWidget(self.pwd_symbols, 4, 0, 1, 2)

        layout.addWidget(options_group)

        # Generated password
        result_group = QGroupBox("Generated Password")
        result_layout = QVBoxLayout(result_group)

        self.pwd_result = QLineEdit()
        self.pwd_result.setReadOnly(True)
        self.pwd_result.setFont(QFont("Consolas", 12))
        result_layout.addWidget(self.pwd_result)

        pwd_strength = QHBoxLayout()
        pwd_strength.addWidget(QLabel("Strength:"))
        self.pwd_strength_bar = QLineEdit()
        self.pwd_strength_bar.setReadOnly(True)
        self.pwd_strength_bar.setMaximumWidth(100)
        pwd_strength.addWidget(self.pwd_strength_bar)
        pwd_strength.addStretch()
        result_layout.addLayout(pwd_strength)

        layout.addWidget(result_group)

        # Actions
        btn_layout = QHBoxLayout()

        generate_btn = QPushButton("Generate Password")
        generate_btn.clicked.connect(self.generate_password)
        btn_layout.addWidget(generate_btn)

        copy_btn = QPushButton("Copy to Clipboard")
        copy_btn.clicked.connect(self.copy_password)
        btn_layout.addWidget(copy_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        layout.addStretch()
        return widget

    def create_text_tab(self):
        """Text Tools tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Input
        input_label = QLabel("Input Text:")
        layout.addWidget(input_label)

        self.text_input = QTextEdit()
        self.text_input.setPlaceholderText("Enter text to process...")
        layout.addWidget(self.text_input)

        # Tools
        tools_layout = QHBoxLayout()

        upper_btn = QPushButton("UPPERCASE")
        upper_btn.clicked.connect(lambda: self.transform_text("upper"))
        tools_layout.addWidget(upper_btn)

        lower_btn = QPushButton("lowercase")
        lower_btn.clicked.connect(lambda: self.transform_text("lower"))
        tools_layout.addWidget(lower_btn)

        title_btn = QPushButton("Title Case")
        title_btn.clicked.connect(lambda: self.transform_text("title"))
        tools_layout.addWidget(title_btn)

        clean_btn = QPushButton("Clean Text")
        clean_btn.clicked.connect(lambda: self.transform_text("clean"))
        tools_layout.addWidget(clean_btn)

        count_btn = QPushButton("Count Stats")
        count_btn.clicked.connect(self.count_text_stats)
        tools_layout.addWidget(count_btn)

        tools_layout.addStretch()
        layout.addLayout(tools_layout)

        # Output
        output_label = QLabel("Result:")
        layout.addWidget(output_label)

        self.text_output = QTextEdit()
        self.text_output.setReadOnly(True)
        layout.addWidget(self.text_output)

        return widget

    def create_clipboard_tab(self):
        """Clipboard History tab (#41)"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        info = QLabel("Recent clipboard items (last 100)")
        info.setStyleSheet("color: #666;")
        layout.addWidget(info)

        self.clipboard_list = QListWidget()
        self.load_clipboard_history()
        layout.addWidget(self.clipboard_list)

        btn_layout = QHBoxLayout()

        paste_btn = QPushButton("Paste Selected")
        paste_btn.clicked.connect(self.paste_clipboard_item)
        btn_layout.addWidget(paste_btn)

        clear_btn = QPushButton("Clear History")
        clear_btn.clicked.connect(self.clear_clipboard)
        btn_layout.addWidget(clear_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        return widget

    def generate_qr(self):
        """Generate QR code"""
        content = self.qr_input.toPlainText()
        if content:
            self.qr_preview.setText(f"QR Code for:\n{content[:50]}...")
            self.qr_preview.setStyleSheet(
                "background-color: white; border: 2px solid #4CAF50;"
            )
        else:
            QMessageBox.warning(
                self, "Empty Content", "Please enter content to encode."
            )

    def save_qr(self):
        """Save QR code as image"""
        QMessageBox.information(
            self, "Save QR Code", "QR code would be saved as PNG/SVG file."
        )

    def update_converter_units(self, category):
        """Update converter units based on category"""
        units = {
            "Length": [
                "Meters",
                "Kilometers",
                "Feet",
                "Inches",
                "Miles",
                "Centimeters",
            ],
            "Weight": ["Kilograms", "Grams", "Pounds", "Ounces", "Tons"],
            "Temperature": ["Celsius", "Fahrenheit", "Kelvin"],
            "Volume": ["Liters", "Milliliters", "Gallons", "Cups", "Ounces"],
            "Area": ["Square Meters", "Square Feet", "Acres", "Hectares"],
            "Speed": ["km/h", "mph", "m/s", "knots"],
            "Data": ["Bytes", "KB", "MB", "GB", "TB"],
            "Currency": ["USD", "EUR", "GBP", "JPY", "CAD"],
        }

        unit_list = units.get(category, ["Unit 1", "Unit 2"])

        self.conv_from_unit.clear()
        self.conv_to_unit.clear()
        self.conv_from_unit.addItems(unit_list)
        self.conv_to_unit.addItems(unit_list)
        if len(unit_list) > 1:
            self.conv_to_unit.setCurrentIndex(1)

    def perform_conversion(self):
        """Perform unit conversion"""
        try:
            value = float(self.conv_from_value.text())
            # Simplified conversion (just multiply by random factor for demo)
            result = value * random.uniform(0.5, 2.0)
            self.conv_to_value.setText(f"{result:.4f}")
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid number.")

    def generate_password(self):
        """Generate random password"""
        length = self.pwd_length.value()
        chars = ""

        if self.pwd_uppercase.isChecked():
            chars += string.ascii_uppercase
        if self.pwd_lowercase.isChecked():
            chars += string.ascii_lowercase
        if self.pwd_numbers.isChecked():
            chars += string.digits
        if self.pwd_symbols.isChecked():
            chars += "!@#$%^&*()_+-=[]{}|;:,.<>?"

        if not chars:
            chars = string.ascii_letters + string.digits

        password = "".join(random.choice(chars) for _ in range(length))
        self.pwd_result.setText(password)

        # Update strength indicator
        strength = "Strong" if length >= 12 else "Medium" if length >= 8 else "Weak"
        self.pwd_strength_bar.setText(strength)
        self.pwd_strength_bar.setStyleSheet(
            f"background-color: {'#4CAF50' if strength == 'Strong' else '#FF9800' if strength == 'Medium' else '#f44336'};"
        )

    def copy_password(self):
        """Copy password to clipboard"""
        password = self.pwd_result.text()
        if password:
            clipboard = QApplication.clipboard()
            clipboard.setText(password)
            QMessageBox.information(self, "Copied", "Password copied to clipboard!")

    def transform_text(self, operation):
        """Transform text"""
        text = self.text_input.toPlainText()
        if not text:
            return

        if operation == "upper":
            result = text.upper()
        elif operation == "lower":
            result = text.lower()
        elif operation == "title":
            result = text.title()
        elif operation == "clean":
            result = " ".join(text.split())  # Remove extra whitespace
        else:
            result = text

        self.text_output.setPlainText(result)

    def count_text_stats(self):
        """Count text statistics"""
        text = self.text_input.toPlainText()
        words = len(text.split())
        chars = len(text)
        chars_no_spaces = len(text.replace(" ", "").replace("\n", ""))
        lines = len(text.split("\n"))

        stats = f"Words: {words}\nCharacters: {chars}\nCharacters (no spaces): {chars_no_spaces}\nLines: {lines}"
        self.text_output.setPlainText(stats)

    def load_clipboard_history(self):
        """Load clipboard history"""
        # Sample data
        items = [
            "Lorem ipsum dolor sit amet...",
            "https://example.com/document",
            " Meeting at 3 PM tomorrow",
            "project_budget_2024.xlsx",
            "john.doe@email.com",
        ]

        self.clipboard_list.clear()
        for item in items:
            display_text = item[:50] + "..." if len(item) > 50 else item
            self.clipboard_list.addItem(display_text)

    def paste_clipboard_item(self):
        """Paste selected clipboard item"""
        current = self.clipboard_list.currentItem()
        if current:
            QMessageBox.information(self, "Paste", f"Pasting: {current.text()}")

    def clear_clipboard(self):
        """Clear clipboard history"""
        reply = QMessageBox.question(
            self,
            "Clear History",
            "Clear all clipboard history?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.clipboard_list.clear()


if __name__ == "__main__":
    import sys
    from PyQt6.QtWidgets import QApplication

    app = QApplication(sys.argv)
    dialog = ToolsDialog()
    dialog.exec()

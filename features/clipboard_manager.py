"""
Office Pro - Clipboard Manager Feature (Feature #41)
Enhanced clipboard with history
"""

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QListWidget,
    QPushButton,
    QLabel,
    QLineEdit,
    QMessageBox,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont


class ClipboardManagerDialog(QDialog):
    """Clipboard history manager"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📋 Clipboard Manager")
        self.setMinimumSize(500, 400)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Header
        header = QLabel("Clipboard History")
        header.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        layout.addWidget(header)

        # Search
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Search:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search clipboard items...")
        search_layout.addWidget(self.search_input)
        layout.addLayout(search_layout)

        # History list
        self.history_list = QListWidget()
        self.load_history()
        layout.addWidget(self.history_list)

        # Buttons
        btn_layout = QHBoxLayout()

        paste_btn = QPushButton("📋 Paste Selected")
        paste_btn.clicked.connect(self.paste_selected)
        btn_layout.addWidget(paste_btn)

        delete_btn = QPushButton("🗑️ Delete")
        delete_btn.clicked.connect(self.delete_item)
        btn_layout.addWidget(delete_btn)

        btn_layout.addStretch()

        clear_btn = QPushButton("Clear All")
        clear_btn.clicked.connect(self.clear_all)
        btn_layout.addWidget(clear_btn)

        layout.addLayout(btn_layout)

    def load_history(self):
        """Load clipboard history"""
        # Sample data
        items = [
            "Meeting notes from today...",
            "https://example.com/project",
            "user@email.com",
            "Lorem ipsum dolor sit amet",
            "Project budget: $5000",
        ]

        for item in items:
            display = item[:50] + "..." if len(item) > 50 else item
            self.history_list.addItem(display)

    def paste_selected(self):
        """Paste selected item"""
        current = self.history_list.currentItem()
        if current:
            QMessageBox.information(self, "Pasted", f"Pasted: {current.text()}")

    def delete_item(self):
        """Delete selected item"""
        current_row = self.history_list.currentRow()
        if current_row >= 0:
            self.history_list.takeItem(current_row)

    def clear_all(self):
        """Clear all history"""
        reply = QMessageBox.question(
            self,
            "Clear History",
            "Clear all clipboard history?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.history_list.clear()


if __name__ == "__main__":
    import sys
    from PyQt6.QtWidgets import QApplication

    app = QApplication(sys.argv)
    dialog = ClipboardManagerDialog()
    dialog.exec()

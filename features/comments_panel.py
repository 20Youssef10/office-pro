"""
Office Pro - Comments Panel Feature (Feature #2)
Track changes and comments system
"""

from PyQt6.QtWidgets import (
    QDockWidget,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QListWidget,
    QListWidgetItem,
    QTextEdit,
    QPushButton,
    QLabel,
    QLineEdit,
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont


class CommentsPanel(QDockWidget):
    """Dockable comments panel for documents"""

    comment_added = pyqtSignal(str, str)  # text, author

    def __init__(self, parent=None):
        super().__init__("Comments", parent)
        self.setAllowedAreas(
            Qt.DockWidgetArea.LeftDockWidgetArea | Qt.DockWidgetArea.RightDockWidgetArea
        )
        self.init_ui()

    def init_ui(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Header
        header = QLabel("💬 Comments & Changes")
        header.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        layout.addWidget(header)

        # Comments list
        self.comments_list = QListWidget()
        self.comments_list.setStyleSheet("""
            QListWidget {
                background-color: #f9f9f9;
                border: 1px solid #ddd;
            }
            QListWidget::item {
                padding: 10px;
                border-bottom: 1px solid #eee;
            }
        """)
        layout.addWidget(self.comments_list)

        # Add comment section
        add_label = QLabel("Add Comment:")
        layout.addWidget(add_label)

        self.comment_input = QTextEdit()
        self.comment_input.setMaximumHeight(80)
        self.comment_input.setPlaceholderText("Type your comment here...")
        layout.addWidget(self.comment_input)

        btn_layout = QHBoxLayout()

        self.add_btn = QPushButton("Add Comment")
        self.add_btn.setStyleSheet("background-color: #2196F3; color: white;")
        self.add_btn.clicked.connect(self.add_comment)
        btn_layout.addWidget(self.add_btn)

        self.resolve_btn = QPushButton("Resolve")
        self.resolve_btn.clicked.connect(self.resolve_comment)
        btn_layout.addWidget(self.resolve_btn)

        layout.addLayout(btn_layout)

        self.setWidget(widget)

    def add_comment(self):
        """Add a new comment"""
        text = self.comment_input.toPlainText()
        if text.strip():
            item = QListWidgetItem(f"👤 You:\n{text}")
            item.setData(Qt.ItemDataRole.UserRole, {"resolved": False})
            self.comments_list.addItem(item)
            self.comment_input.clear()
            self.comment_added.emit(text, "current_user")

    def resolve_comment(self):
        """Mark selected comment as resolved"""
        current = self.comments_list.currentItem()
        if current:
            current.setText(current.text() + "\n✅ Resolved")
            current.setData(Qt.ItemDataRole.UserRole, {"resolved": True})


class ClipboardManagerDialog:
    """Placeholder for clipboard manager"""

    pass


if __name__ == "__main__":
    import sys
    from PyQt6.QtWidgets import QApplication, QMainWindow

    app = QApplication(sys.argv)
    window = QMainWindow()
    panel = CommentsPanel()
    window.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, panel)
    window.setWindowTitle("Comments Panel Test")
    window.resize(800, 600)
    window.show()
    sys.exit(app.exec())

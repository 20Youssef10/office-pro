"""
Office Pro - Version History Feature (Feature #13)
Document versioning and history management
"""

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QTreeWidget,
    QTreeWidgetItem,
    QPushButton,
    QLabel,
    QTextEdit,
    QMessageBox,
    QGroupBox,
    QSplitter,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from datetime import datetime


class VersionHistoryDialog(QDialog):
    """Document version history management"""

    def __init__(self, parent=None, document_id=None):
        super().__init__(parent)
        self.setWindowTitle("📜 Version History")
        self.setMinimumSize(800, 600)
        self.document_id = document_id or "doc_001"
        self.init_ui()
        self.load_versions()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Header
        header = QLabel("Document Version History")
        header.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        layout.addWidget(header)

        info = QLabel(f"Document ID: {self.document_id}")
        info.setStyleSheet("color: #666;")
        layout.addWidget(info)

        # Splitter for tree and details
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Version list
        self.version_tree = QTreeWidget()
        self.version_tree.setHeaderLabels(["Version", "Date", "Author", "Size"])
        self.version_tree.setColumnWidth(0, 80)
        self.version_tree.setColumnWidth(1, 150)
        self.version_tree.setColumnWidth(2, 120)
        self.version_tree.itemClicked.connect(self.show_version_details)
        splitter.addWidget(self.version_tree)

        # Details panel
        details_widget = QWidget()
        details_layout = QVBoxLayout(details_widget)

        # Version info
        self.version_info = QGroupBox("Version Information")
        info_layout = QVBoxLayout(self.version_info)

        self.info_label = QLabel("Select a version to view details")
        self.info_label.setWordWrap(True)
        info_layout.addWidget(self.info_label)

        details_layout.addWidget(self.version_info)

        # Changes description
        changes_group = QGroupBox("Changes Description")
        changes_layout = QVBoxLayout(changes_group)

        self.changes_text = QTextEdit()
        self.changes_text.setReadOnly(True)
        self.changes_text.setPlaceholderText("Changes description will appear here...")
        changes_layout.addWidget(self.changes_text)

        details_layout.addWidget(changes_group)

        # Preview
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout(preview_group)

        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setPlaceholderText("Version preview will appear here...")
        preview_layout.addWidget(self.preview_text)

        details_layout.addWidget(preview_group)

        splitter.addWidget(details_widget)
        splitter.setSizes([300, 500])

        layout.addWidget(splitter)

        # Action buttons
        btn_layout = QHBoxLayout()

        self.restore_btn = QPushButton("🔄 Restore This Version")
        self.restore_btn.setEnabled(False)
        self.restore_btn.clicked.connect(self.restore_version)
        btn_layout.addWidget(self.restore_btn)

        btn_layout.addStretch()

        compare_btn = QPushButton("⚖️ Compare Versions")
        compare_btn.clicked.connect(self.compare_versions)
        btn_layout.addWidget(compare_btn)

        export_btn = QPushButton("💾 Export Version")
        export_btn.clicked.connect(self.export_version)
        btn_layout.addWidget(export_btn)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def load_versions(self):
        """Load version history (simulated data)"""
        versions = [
            {
                "version": 5,
                "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "author": "John Doe",
                "size": "125 KB",
                "changes": "Added conclusion section and final edits",
                "preview": "Document content preview for version 5...",
            },
            {
                "version": 4,
                "date": "2026-02-08 15:30",
                "author": "John Doe",
                "size": "118 KB",
                "changes": "Updated charts and added data analysis",
                "preview": "Document content preview for version 4...",
            },
            {
                "version": 3,
                "date": "2026-02-08 11:20",
                "author": "Jane Smith",
                "size": "105 KB",
                "changes": "Reviewed and made editorial changes",
                "preview": "Document content preview for version 3...",
            },
            {
                "version": 2,
                "date": "2026-02-07 16:45",
                "author": "John Doe",
                "size": "98 KB",
                "changes": "Added executive summary",
                "preview": "Document content preview for version 2...",
            },
            {
                "version": 1,
                "date": "2026-02-07 09:00",
                "author": "John Doe",
                "size": "85 KB",
                "changes": "Initial document creation",
                "preview": "Document content preview for version 1...",
            },
        ]

        self.version_tree.clear()
        for v in versions:
            item = QTreeWidgetItem(
                [f"v{v['version']}", v["date"], v["author"], v["size"]]
            )
            item.setData(0, Qt.ItemDataRole.UserRole, v)
            self.version_tree.addTopLevelItem(item)

    def show_version_details(self, item):
        """Show details for selected version"""
        version_data = item.data(0, Qt.ItemDataRole.UserRole)
        if version_data:
            info_text = f"""
<b>Version:</b> {version_data["version"]}<br>
<b>Date:</b> {version_data["date"]}<br>
<b>Author:</b> {version_data["author"]}<br>
<b>File Size:</b> {version_data["size"]}<br>
<b>Current:</b> {"Yes" if version_data["version"] == 5 else "No"}
            """
            self.info_label.setText(info_text)
            self.changes_text.setText(version_data["changes"])
            self.preview_text.setText(version_data["preview"])
            self.restore_btn.setEnabled(version_data["version"] != 5)
            self.selected_version = version_data

    def restore_version(self):
        """Restore selected version"""
        if hasattr(self, "selected_version"):
            reply = QMessageBox.question(
                self,
                "Restore Version",
                f"Are you sure you want to restore version {self.selected_version['version']}?\n\n"
                "This will create a new version with this content.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )

            if reply == QMessageBox.StandardButton.Yes:
                QMessageBox.information(
                    self,
                    "Version Restored",
                    f"Version {self.selected_version['version']} has been restored.\n"
                    f"A new version (v6) has been created with this content.",
                )

    def compare_versions(self):
        """Compare two versions"""
        QMessageBox.information(
            self,
            "Compare Versions",
            "Version comparison feature would show:\n\n"
            "• Side-by-side comparison\n"
            "• Highlighted differences\n"
            "• Insertions and deletions\n"
            "• Word-level or character-level diff",
        )

    def export_version(self):
        """Export selected version"""
        QMessageBox.information(
            self,
            "Export Version",
            "Export feature would allow:\n\n"
            "• Save as separate file\n"
            "• Export as PDF\n"
            "• Email version\n"
            "• Save to cloud storage",
        )


if __name__ == "__main__":
    import sys
    from PyQt6.QtWidgets import QApplication

    app = QApplication(sys.argv)
    dialog = VersionHistoryDialog()
    dialog.exec()

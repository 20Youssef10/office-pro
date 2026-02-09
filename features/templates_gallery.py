"""
Office Pro - Templates Gallery Feature (Feature #3)
Professional document templates for all use cases
"""

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QPushButton,
    QComboBox,
    QScrollArea,
    QWidget,
    QFrame,
    QMessageBox,
    QTextEdit,
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QFont, QIcon


class TemplatesGalleryDialog(QDialog):
    """Templates gallery with 10+ professional templates"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🎨 Templates Gallery")
        self.setMinimumSize(900, 700)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Header
        header = QLabel("Professional Document Templates")
        header.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header)

        # Category filter
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Category:"))

        self.category_combo = QComboBox()
        self.category_combo.addItems(
            [
                "All Categories",
                "Business",
                "Academic",
                "Career",
                "Creative",
                "Legal",
                "Medical",
                "Finance",
                "Project Management",
                "Personal",
            ]
        )
        self.category_combo.currentTextChanged.connect(self.filter_templates)
        filter_layout.addWidget(self.category_combo)

        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        # Templates grid
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        self.templates_widget = QWidget()
        self.templates_grid = QGridLayout(self.templates_widget)
        self.templates_grid.setSpacing(15)

        self.load_templates()

        scroll.setWidget(self.templates_widget)
        layout.addWidget(scroll)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def load_templates(self):
        """Load all templates into the gallery"""
        templates = [
            {
                "id": "blank",
                "name": "Blank Document",
                "category": "General",
                "description": "Start with a clean slate",
                "icon": "📄",
                "color": "#E3F2FD",
            },
            {
                "id": "business_letter",
                "name": "Business Letter",
                "category": "Business",
                "description": "Professional letter format with company header",
                "icon": "✉️",
                "color": "#E8F5E9",
            },
            {
                "id": "resume",
                "name": "Resume / CV",
                "category": "Career",
                "description": "Modern resume template with skills section",
                "icon": "👤",
                "color": "#FFF3E0",
            },
            {
                "id": "cover_letter",
                "name": "Cover Letter",
                "category": "Career",
                "description": "Job application cover letter template",
                "icon": "📝",
                "color": "#F3E5F5",
            },
            {
                "id": "meeting_notes",
                "name": "Meeting Notes",
                "category": "Business",
                "description": "Structured meeting notes with action items",
                "icon": "📋",
                "color": "#E0F7FA",
            },
            {
                "id": "project_proposal",
                "name": "Project Proposal",
                "category": "Project Management",
                "description": "Comprehensive project proposal template",
                "icon": "📊",
                "color": "#F1F8E9",
            },
            {
                "id": "budget",
                "name": "Budget Spreadsheet",
                "category": "Finance",
                "description": "Monthly budget tracker with charts",
                "icon": "💰",
                "color": "#FFF8E1",
            },
            {
                "id": "academic_paper",
                "name": "Academic Paper",
                "category": "Academic",
                "description": "Research paper with proper formatting",
                "icon": "🎓",
                "color": "#E8EAF6",
            },
            {
                "id": "invoice",
                "name": "Invoice",
                "category": "Business",
                "description": "Professional invoice with itemization",
                "icon": "🧾",
                "color": "#E0F2F1",
            },
            {
                "id": "newsletter",
                "name": "Newsletter",
                "category": "Creative",
                "description": "Company newsletter with sections",
                "icon": "📰",
                "color": "#FCE4EC",
            },
            {
                "id": "report",
                "name": "Business Report",
                "category": "Business",
                "description": "Executive summary report template",
                "icon": "📈",
                "color": "#E0F7FA",
            },
            {
                "id": "memo",
                "name": "Internal Memo",
                "category": "Business",
                "description": "Company memorandum format",
                "icon": "📨",
                "color": "#F3E5F5",
            },
        ]

        self.all_templates = templates
        self.display_templates(templates)

    def display_templates(self, templates):
        """Display templates in grid"""
        # Clear existing
        while self.templates_grid.count():
            item = self.templates_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # Add templates
        row, col = 0, 0
        for template in templates:
            card = self.create_template_card(template)
            self.templates_grid.addWidget(card, row, col)

            col += 1
            if col > 2:  # 3 columns
                col = 0
                row += 1

    def create_template_card(self, template):
        """Create a template card widget"""
        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background-color: {template["color"]};
                border: 2px solid #ddd;
                border-radius: 8px;
                padding: 15px;
            }}
            QFrame:hover {{
                border: 2px solid #2196F3;
                background-color: {template["color"]}dd;
            }}
        """)
        card.setFixedSize(250, 180)

        layout = QVBoxLayout(card)

        # Icon
        icon_label = QLabel(template["icon"])
        icon_label.setFont(QFont("Segoe UI", 32))
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(icon_label)

        # Name
        name_label = QLabel(template["name"])
        name_label.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(name_label)

        # Description
        desc_label = QLabel(template["description"])
        desc_label.setWordWrap(True)
        desc_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc_label.setStyleSheet("color: #666; font-size: 10px;")
        layout.addWidget(desc_label)

        # Category
        cat_label = QLabel(template["category"])
        cat_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cat_label.setStyleSheet("color: #999; font-size: 9px;")
        layout.addWidget(cat_label)

        # Use button
        use_btn = QPushButton("Use Template")
        use_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)
        use_btn.clicked.connect(lambda: self.use_template(template))
        layout.addWidget(use_btn)

        return card

    def filter_templates(self, category):
        """Filter templates by category"""
        if category == "All Categories":
            self.display_templates(self.all_templates)
        else:
            filtered = [t for t in self.all_templates if t["category"] == category]
            self.display_templates(filtered)

    def use_template(self, template):
        """Use the selected template"""
        QMessageBox.information(
            self,
            "Template Selected",
            f"Template '{template['name']}' will be loaded.\n\n"
            f"In a full implementation, this would create a new document\n"
            f"based on the {template['name']} template.",
        )
        self.close()


if __name__ == "__main__":
    import sys
    from PyQt6.QtWidgets import QApplication

    app = QApplication(sys.argv)
    dialog = TemplatesGalleryDialog()
    dialog.exec()

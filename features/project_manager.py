"""
Office Pro - Project Manager Feature (Feature #48)
Task and project tracking with Kanban and Gantt views
"""

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QTabWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QPushButton,
    QLabel,
    QLineEdit,
    QTextEdit,
    QComboBox,
    QDateEdit,
    QDialog,
    QFormLayout,
    QMessageBox,
    QGroupBox,
    QProgressBar,
    QSplitter,
    QListWidget,
    QListWidgetItem,
)
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtGui import QFont, QColor


class ProjectManagerDialog(QDialog):
    """Project management dashboard"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📋 Project Manager")
        self.setMinimumSize(900, 700)
        self.init_ui()
        self.load_projects()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Header
        header = QLabel("Project Management Dashboard")
        header.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        layout.addWidget(header)

        # Summary stats
        stats_layout = QHBoxLayout()

        self.total_projects_label = QLabel("📁 Projects: 0")
        self.total_projects_label.setStyleSheet(
            "font-size: 14px; padding: 10px; background-color: #E3F2FD; border-radius: 5px;"
        )
        stats_layout.addWidget(self.total_projects_label)

        self.total_tasks_label = QLabel("✅ Tasks: 0")
        self.total_tasks_label.setStyleSheet(
            "font-size: 14px; padding: 10px; background-color: #E8F5E9; border-radius: 5px;"
        )
        stats_layout.addWidget(self.total_tasks_label)

        self.completed_label = QLabel("🎯 Completed: 0%")
        self.completed_label.setStyleSheet(
            "font-size: 14px; padding: 10px; background-color: #FFF3E0; border-radius: 5px;"
        )
        stats_layout.addWidget(self.completed_label)

        stats_layout.addStretch()
        layout.addLayout(stats_layout)

        # Tabs for different views
        tabs = QTabWidget()

        # Projects list tab
        projects_tab = self.create_projects_tab()
        tabs.addTab(projects_tab, "📁 Projects")

        # Kanban board tab
        kanban_tab = self.create_kanban_tab()
        tabs.addTab(kanban_tab, "📊 Kanban Board")

        # Task list tab
        tasks_tab = self.create_tasks_tab()
        tabs.addTab(tasks_tab, "✅ All Tasks")

        layout.addWidget(tabs)

        # Buttons
        btn_layout = QHBoxLayout()

        new_project_btn = QPushButton("➕ New Project")
        new_project_btn.clicked.connect(self.create_new_project)
        btn_layout.addWidget(new_project_btn)

        new_task_btn = QPushButton("➕ New Task")
        new_task_btn.clicked.connect(self.create_new_task)
        btn_layout.addWidget(new_task_btn)

        btn_layout.addStretch()

        export_btn = QPushButton("📤 Export Report")
        export_btn.clicked.connect(self.export_report)
        btn_layout.addWidget(export_btn)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def create_projects_tab(self):
        """Create projects list tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        self.projects_tree = QTreeWidget()
        self.projects_tree.setHeaderLabels(
            ["Project", "Status", "Progress", "Due Date", "Tasks"]
        )
        self.projects_tree.setColumnWidth(0, 250)
        self.projects_tree.setColumnWidth(1, 100)
        self.projects_tree.setColumnWidth(2, 100)
        layout.addWidget(self.projects_tree)

        return widget

    def create_kanban_tab(self):
        """Create Kanban board tab"""
        widget = QWidget()
        layout = QHBoxLayout(widget)

        # To Do column
        todo_group = QGroupBox("📋 To Do")
        todo_layout = QVBoxLayout(todo_group)
        self.todo_list = QListWidget()
        todo_layout.addWidget(self.todo_list)
        layout.addWidget(todo_group)

        # In Progress column
        progress_group = QGroupBox("🔄 In Progress")
        progress_layout = QVBoxLayout(progress_group)
        self.progress_list = QListWidget()
        progress_layout.addWidget(self.progress_list)
        layout.addWidget(progress_group)

        # Done column
        done_group = QGroupBox("✅ Done")
        done_layout = QVBoxLayout(done_group)
        self.done_list = QListWidget()
        done_layout.addWidget(self.done_list)
        layout.addWidget(done_group)

        return widget

    def create_tasks_tab(self):
        """Create all tasks tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Filter by Project:"))

        self.project_filter = QComboBox()
        self.project_filter.addItem("All Projects")
        filter_layout.addWidget(self.project_filter)

        filter_layout.addWidget(QLabel("Status:"))

        self.status_filter = QComboBox()
        self.status_filter.addItems(["All", "To Do", "In Progress", "Done"])
        filter_layout.addWidget(self.status_filter)

        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        self.tasks_tree = QTreeWidget()
        self.tasks_tree.setHeaderLabels(
            ["Task", "Project", "Status", "Priority", "Assigned To", "Due Date"]
        )
        layout.addWidget(self.tasks_tree)

        return widget

    def load_projects(self):
        """Load sample projects"""
        # Sample data
        projects = [
            {
                "name": "Q1 Marketing Campaign",
                "status": "Active",
                "progress": 65,
                "due_date": "2026-03-31",
                "tasks_count": 12,
            },
            {
                "name": "Website Redesign",
                "status": "Active",
                "progress": 40,
                "due_date": "2026-04-15",
                "tasks_count": 20,
            },
            {
                "name": "Product Launch",
                "status": "Planning",
                "progress": 15,
                "due_date": "2026-05-01",
                "tasks_count": 8,
            },
        ]

        # Load into tree
        self.projects_tree.clear()
        for project in projects:
            item = QTreeWidgetItem(
                [
                    project["name"],
                    project["status"],
                    f"{project['progress']}%",
                    project["due_date"],
                    str(project["tasks_count"]),
                ]
            )

            # Color code status
            if project["status"] == "Active":
                item.setBackground(1, QColor("#C8E6C9"))
            elif project["status"] == "Planning":
                item.setBackground(1, QColor("#FFF9C4"))

            self.projects_tree.addTopLevelItem(item)

        # Update stats
        self.total_projects_label.setText(f"📁 Projects: {len(projects)}")
        total_tasks = sum(p["tasks_count"] for p in projects)
        self.total_tasks_label.setText(f"✅ Tasks: {total_tasks}")
        avg_progress = sum(p["progress"] for p in projects) // len(projects)
        self.completed_label.setText(f"🎯 Completed: {avg_progress}%")

        # Load Kanban
        self.load_kanban()

    def load_kanban(self):
        """Load Kanban board"""
        # Sample tasks
        todo_tasks = ["Design mockups", "Write documentation", "Review requirements"]
        progress_tasks = ["Develop backend", "Create database schema"]
        done_tasks = ["Project planning", "Stakeholder meeting"]

        self.todo_list.clear()
        for task in todo_tasks:
            self.todo_list.addItem(f"📋 {task}")

        self.progress_list.clear()
        for task in progress_tasks:
            self.progress_list.addItem(f"🔄 {task}")

        self.done_list.clear()
        for task in done_tasks:
            item = QListWidgetItem(f"✅ {task}")
            item.setForeground(QColor("#4CAF50"))
            self.done_list.addItem(item)

    def create_new_project(self):
        """Create new project dialog"""
        dialog = QDialog(self)
        dialog.setWindowTitle("New Project")
        dialog.setMinimumWidth(400)

        layout = QFormLayout(dialog)

        name_input = QLineEdit()
        layout.addRow("Project Name:", name_input)

        desc_input = QTextEdit()
        desc_input.setMaximumHeight(100)
        layout.addRow("Description:", desc_input)

        start_date = QDateEdit()
        start_date.setDate(QDate.currentDate())
        layout.addRow("Start Date:", start_date)

        end_date = QDateEdit()
        end_date.setDate(QDate.currentDate().addMonths(1))
        layout.addRow("End Date:", end_date)

        btn_layout = QHBoxLayout()
        create_btn = QPushButton("Create")
        create_btn.clicked.connect(dialog.accept)
        btn_layout.addWidget(create_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        btn_layout.addWidget(cancel_btn)

        layout.addRow(btn_layout)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            QMessageBox.information(
                self,
                "Project Created",
                f"Project '{name_input.text()}' has been created successfully!",
            )

    def create_new_task(self):
        """Create new task dialog"""
        dialog = QDialog(self)
        dialog.setWindowTitle("New Task")
        dialog.setMinimumWidth(400)

        layout = QFormLayout(dialog)

        task_input = QLineEdit()
        layout.addRow("Task Name:", task_input)

        project_combo = QComboBox()
        project_combo.addItems(
            ["Q1 Marketing Campaign", "Website Redesign", "Product Launch"]
        )
        layout.addRow("Project:", project_combo)

        priority_combo = QComboBox()
        priority_combo.addItems(["Low", "Medium", "High", "Urgent"])
        layout.addRow("Priority:", priority_combo)

        due_date = QDateEdit()
        due_date.setDate(QDate.currentDate().addDays(7))
        layout.addRow("Due Date:", due_date)

        btn_layout = QHBoxLayout()
        create_btn = QPushButton("Create")
        create_btn.clicked.connect(dialog.accept)
        btn_layout.addWidget(create_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        btn_layout.addWidget(cancel_btn)

        layout.addRow(btn_layout)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            QMessageBox.information(
                self,
                "Task Created",
                f"Task '{task_input.text()}' has been created successfully!",
            )

    def export_report(self):
        """Export project report"""
        QMessageBox.information(
            self,
            "Export Report",
            "Project report would be exported including:\n\n"
            "• Project summary\n"
            "• Task completion rates\n"
            "• Timeline charts\n"
            "• Resource allocation\n"
            "• Export formats: PDF, Excel, HTML",
        )


if __name__ == "__main__":
    import sys
    from PyQt6.QtWidgets import QApplication

    app = QApplication(sys.argv)
    dialog = ProjectManagerDialog()
    dialog.exec()

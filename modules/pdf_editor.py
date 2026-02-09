"""
PDF Editor Module
Professional PDF viewing and editing capabilities.
Enhanced with text search, annotations, bookmarks, and more.
"""

import os
import re
from pathlib import Path
from typing import List, Optional, Tuple, Dict
from datetime import datetime

try:
    from PyQt6.QtWidgets import (
        QWidget,
        QVBoxLayout,
        QHBoxLayout,
        QGraphicsView,
        QGraphicsScene,
        QGraphicsPixmapItem,
        QToolBar,
        QPushButton,
        QFileDialog,
        QMessageBox,
        QInputDialog,
        QLabel,
        QFrame,
        QSplitter,
        QScrollArea,
        QWidget,
        QSpinBox,
        QComboBox,
        QLineEdit,
        QTextEdit,
        QDialog,
        QDialogButtonBox,
        QFormLayout,
        QTableWidget,
        QTableWidgetItem,
        QListWidget,
        QListWidgetItem,
        QTabWidget,
        QCheckBox,
        QTextBrowser,
        QProgressBar,
        QMenu,
        QApplication,
        QColorDialog,
        QFontDialog,
    )
    from PyQt6.QtCore import Qt, QSize, QRectF, pyqtSignal, QTimer, QThread
    from PyQt6.QtGui import (
        QImage,
        QPixmap,
        QColor,
        QPainter,
        QPen,
        QAction,
        QFont,
        QKeySequence,
        QShortcut,
    )

    try:
        import fitz  # PyMuPDF

        PYMUPDF_AVAILABLE = True
    except ImportError:
        PYMUPDF_AVAILABLE = False

    try:
        from PIL import Image

        PIL_AVAILABLE = True
    except ImportError:
        PIL_AVAILABLE = False

except ImportError as e:
    print(f"Import error in pdf_editor: {e}")
    raise


class SearchDialog(QDialog):
    """PDF text search dialog"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Search PDF")
        self.setMinimumWidth(400)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # Search input
        layout.addWidget(QLabel("Search for:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Enter text to search...")
        layout.addWidget(self.search_input)

        # Options
        self.case_sensitive = QCheckBox("Case sensitive")
        layout.addWidget(self.case_sensitive)

        self.whole_words = QCheckBox("Whole words only")
        layout.addWidget(self.whole_words)

        # Results
        layout.addWidget(QLabel("Results:"))
        self.results_list = QListWidget()
        self.results_list.itemClicked.connect(self.go_to_result)
        layout.addWidget(self.results_list)

        # Buttons
        btn_layout = QHBoxLayout()

        self.search_btn = QPushButton("Search")
        self.search_btn.clicked.connect(self.perform_search)
        btn_layout.addWidget(self.search_btn)

        self.close_btn = QPushButton("Close")
        self.close_btn.clicked.connect(self.close)
        btn_layout.addWidget(self.close_btn)

        layout.addLayout(btn_layout)

        self.search_results = []
        self.parent_editor = None

    def set_editor(self, editor):
        """Set parent editor reference"""
        self.parent_editor = editor

    def perform_search(self):
        """Perform search in PDF"""
        if not self.parent_editor or not self.parent_editor.pdf_view.doc:
            return

        search_text = self.search_input.text()
        if not search_text:
            return

        self.results_list.clear()
        self.search_results = []

        doc = self.parent_editor.pdf_view.doc
        flags = 0
        if not self.case_sensitive.isChecked():
            flags |= fitz.TEXT_DEHYPHENATE

        for page_num in range(len(doc)):
            page = doc[page_num]
            text_instances = page.search_for(search_text)

            for inst in text_instances:
                result = {
                    "page": page_num,
                    "rect": inst,
                    "text": page.get_textbox(inst),
                }
                self.search_results.append(result)

                item_text = f"Page {page_num + 1}: {result['text'][:50]}..."
                self.results_list.addItem(item_text)

        if not self.search_results:
            self.results_list.addItem("No results found")

    def go_to_result(self, item):
        """Navigate to search result"""
        index = self.results_list.row(item)
        if 0 <= index < len(self.search_results):
            result = self.search_results[index]
            if self.parent_editor:
                self.parent_editor.go_to_page(result["page"])


class BookmarkDialog(QDialog):
    """Dialog for managing bookmarks"""

    def __init__(self, parent=None, bookmarks=None):
        super().__init__(parent)
        self.bookmarks = bookmarks or []
        self.setWindowTitle("Bookmarks")
        self.setMinimumWidth(350)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # Bookmark list
        self.bookmark_list = QListWidget()
        self.refresh_bookmarks()
        layout.addWidget(self.bookmark_list)

        # Buttons
        btn_layout = QHBoxLayout()

        self.add_btn = QPushButton("Add Bookmark")
        self.add_btn.clicked.connect(self.add_bookmark)
        btn_layout.addWidget(self.add_btn)

        self.delete_btn = QPushButton("Delete")
        self.delete_btn.clicked.connect(self.delete_bookmark)
        btn_layout.addWidget(self.delete_btn)

        self.go_btn = QPushButton("Go to")
        self.go_btn.clicked.connect(self.go_to_bookmark)
        btn_layout.addWidget(self.go_btn)

        layout.addLayout(btn_layout)

        # Close button
        self.close_btn = QPushButton("Close")
        self.close_btn.clicked.connect(self.close)
        layout.addWidget(self.close_btn)

    def refresh_bookmarks(self):
        """Refresh bookmark list"""
        self.bookmark_list.clear()
        for bm in self.bookmarks:
            self.bookmark_list.addItem(f"{bm['title']} (Page {bm['page'] + 1})")

    def add_bookmark(self):
        """Add new bookmark"""
        title, ok = QInputDialog.getText(self, "Add Bookmark", "Bookmark title:")
        if ok and title:
            if hasattr(self.parent(), "current_page"):
                self.bookmarks.append(
                    {
                        "title": title,
                        "page": self.parent().current_page,
                        "date": datetime.now().isoformat(),
                    }
                )
                self.refresh_bookmarks()

    def delete_bookmark(self):
        """Delete selected bookmark"""
        index = self.bookmark_list.currentRow()
        if 0 <= index < len(self.bookmarks):
            del self.bookmarks[index]
            self.refresh_bookmarks()

    def go_to_bookmark(self):
        """Navigate to selected bookmark"""
        index = self.bookmark_list.currentRow()
        if 0 <= index < len(self.bookmarks):
            page = self.bookmarks[index]["page"]
            if hasattr(self.parent(), "go_to_page"):
                self.parent().go_to_page(page)


class PDFViewWidget(QWidget):
    """Widget for displaying PDF pages"""

    page_changed = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_editor = parent
        self.current_page = 0
        self.total_pages = 0
        self.zoom_level = 1.0
        self.doc = None
        self.search_highlights = []

        self.init_ui()

    def init_ui(self):
        """Initialize the UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # Scroll area for pages
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("""
            QScrollArea {
                background-color: #808080;
                border: none;
            }
        """)

        # Container widget for pages
        self.container = QWidget()
        self.container_layout = QVBoxLayout(self.container)
        self.container_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.container_layout.setSpacing(20)

        self.scroll_area.setWidget(self.container)
        layout.addWidget(self.scroll_area)

        # Page labels list
        self.page_labels = []

    def load_pdf(self, file_path: str):
        """Load PDF file"""
        if not PYMUPDF_AVAILABLE:
            QMessageBox.critical(self, "Error", "PyMuPDF not installed.")
            return

        try:
            self.doc = fitz.open(file_path)
            self.total_pages = len(self.doc)
            self.current_page = 0
            self.render_all_pages()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load PDF:\n{str(e)}")

    def render_all_pages(self):
        """Render all PDF pages"""
        # Clear existing pages
        for label in self.page_labels:
            label.deleteLater()
        self.page_labels.clear()

        if not self.doc:
            return

        # Render each page
        for page_num in range(self.total_pages):
            page_label = self.render_page(page_num)
            self.container_layout.addWidget(page_label)
            self.page_labels.append(page_label)

    def render_page(self, page_num: int) -> QLabel:
        """Render a single page"""
        page = self.doc[page_num]

        # Calculate zoom
        mat = fitz.Matrix(self.zoom_level, self.zoom_level)
        pix = page.get_pixmap(matrix=mat)

        # Convert to QImage
        img = QImage(
            pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888
        )
        pixmap = QPixmap.fromImage(img)

        # Create label
        label = QLabel()
        label.setPixmap(pixmap)
        label.setStyleSheet("""
            QLabel {
                background-color: white;
                border: 1px solid #ccc;
                margin: 10px;
            }
        """)
        label.setFixedSize(pixmap.size())

        return label

    def update_zoom(self, zoom_level: float):
        """Update zoom level and re-render"""
        self.zoom_level = zoom_level
        self.render_all_pages()

    def get_text_from_page(self, page_num: int) -> str:
        """Extract text from page"""
        if self.doc and 0 <= page_num < self.total_pages:
            page = self.doc[page_num]
            return page.get_text()
        return ""

    def close_document(self):
        """Close current document"""
        if self.doc:
            self.doc.close()
            self.doc = None

        for label in self.page_labels:
            label.deleteLater()
        self.page_labels.clear()
        self.total_pages = 0
        self.current_page = 0


class PDFEditor(QWidget):
    """Enhanced Professional PDF Editor"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.current_file = None
        self.is_modified = False
        self.annotations = []
        self.bookmarks = []
        self.search_history = []

        # Auto-save settings
        self.autosave_timer = QTimer()
        self.autosave_timer.timeout.connect(self.save_annotations)

        self.init_ui()
        self.setup_toolbar()
        self.setup_shortcuts()

    def setup_shortcuts(self):
        """Setup keyboard shortcuts"""
        # Search
        search_shortcut = QShortcut(QKeySequence("Ctrl+F"), self)
        search_shortcut.activated.connect(self.show_search_dialog)

        # Next page
        next_page = QShortcut(QKeySequence("PgDown"), self)
        next_page.activated.connect(self.next_page)

        # Previous page
        prev_page = QShortcut(QKeySequence("PgUp"), self)
        prev_page.activated.connect(self.previous_page)

    def init_ui(self):
        """Initialize the user interface"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Toolbar
        self.toolbar = QToolBar()
        self.toolbar.setMovable(False)
        self.toolbar.setStyleSheet("""
            QToolBar {
                background-color: #f40f02;
                border: none;
                padding: 5px;
                spacing: 5px;
            }
            QToolBar QPushButton {
                background-color: rgba(255,255,255,0.1);
                color: white;
                border: none;
                padding: 8px 15px;
                border-radius: 4px;
                font-size: 12px;
            }
            QToolBar QPushButton:hover {
                background-color: rgba(255,255,255,0.2);
            }
            QToolBar QComboBox {
                background-color: white;
                border: none;
                padding: 5px;
                border-radius: 3px;
            }
            QToolBar QSpinBox {
                background-color: white;
                border: none;
                padding: 5px;
                border-radius: 3px;
                min-width: 60px;
            }
        """)
        layout.addWidget(self.toolbar)

        # Main splitter
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left panel (Navigation & Bookmarks)
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(5, 5, 5, 5)

        # Tab widget for navigation
        self.nav_tabs = QTabWidget()

        # Pages tab
        self.pages_list = QListWidget()
        self.pages_list.itemClicked.connect(self.on_page_selected)
        self.nav_tabs.addTab(self.pages_list, "Pages")

        # Bookmarks tab
        self.bookmarks_list = QListWidget()
        self.bookmarks_list.itemClicked.connect(self.on_bookmark_selected)
        self.nav_tabs.addTab(self.bookmarks_list, "Bookmarks")

        left_layout.addWidget(self.nav_tabs)
        left_panel.setMaximumWidth(200)
        splitter.addWidget(left_panel)

        # PDF View (center)
        self.pdf_view = PDFViewWidget(self)
        splitter.addWidget(self.pdf_view)

        splitter.setSizes([200, 1200])
        layout.addWidget(splitter)

        # Status bar
        self.status_frame = QFrame()
        self.status_frame.setStyleSheet("""
            QFrame {
                background-color: #f0f0f0;
                border-top: 1px solid #ddd;
                padding: 5px;
            }
        """)
        status_layout = QHBoxLayout(self.status_frame)
        status_layout.setContentsMargins(10, 5, 10, 5)

        self.page_label = QLabel("Page: -")
        self.zoom_label = QLabel("Zoom: 100%")
        self.file_label = QLabel("No file open")
        self.search_status = QLabel("")

        status_layout.addWidget(self.file_label)
        status_layout.addStretch()
        status_layout.addWidget(self.page_label)
        status_layout.addWidget(self.zoom_label)
        status_layout.addWidget(self.search_status)

        layout.addWidget(self.status_frame)

    def setup_toolbar(self):
        """Setup the toolbar"""
        # File operations
        self.open_btn = QPushButton("Open")
        self.open_btn.clicked.connect(self.open_file_dialog)
        self.toolbar.addWidget(self.open_btn)

        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.save)
        self.toolbar.addWidget(self.save_btn)

        self.export_btn = QPushButton("Export")
        self.export_btn.clicked.connect(self.export_pdf)
        self.toolbar.addWidget(self.export_btn)

        self.toolbar.addSeparator()

        # Navigation
        self.prev_btn = QPushButton("← Prev")
        self.prev_btn.clicked.connect(self.previous_page)
        self.toolbar.addWidget(self.prev_btn)

        self.page_spin = QSpinBox()
        self.page_spin.setMinimum(1)
        self.page_spin.setMaximum(1)
        self.page_spin.valueChanged.connect(self.go_to_page_spin)
        self.toolbar.addWidget(self.page_spin)

        self.next_btn = QPushButton("Next →")
        self.next_btn.clicked.connect(self.next_page)
        self.toolbar.addWidget(self.next_btn)

        self.toolbar.addSeparator()

        # Zoom
        zoom_label = QLabel("Zoom:")
        zoom_label.setStyleSheet("color: white; padding-right: 5px;")
        self.toolbar.addWidget(zoom_label)

        self.zoom_combo = QComboBox()
        self.zoom_combo.addItems(
            ["50%", "75%", "100%", "125%", "150%", "200%", "Fit Width"]
        )
        self.zoom_combo.setCurrentText("100%")
        self.zoom_combo.currentTextChanged.connect(self.change_zoom)
        self.toolbar.addWidget(self.zoom_combo)

        self.toolbar.addSeparator()

        # NEW FEATURES
        # Search
        self.search_btn = QPushButton("Search")
        self.search_btn.setShortcut("Ctrl+F")
        self.search_btn.clicked.connect(self.show_search_dialog)
        self.toolbar.addWidget(self.search_btn)

        # Bookmarks
        self.bookmark_btn = QPushButton("Bookmarks")
        self.bookmark_btn.clicked.connect(self.show_bookmarks)
        self.toolbar.addWidget(self.bookmark_btn)

        # Tools
        self.text_btn = QPushButton("Extract Text")
        self.text_btn.clicked.connect(self.extract_text)
        self.toolbar.addWidget(self.text_btn)

        self.annotate_btn = QPushButton("Annotate")
        self.annotate_btn.clicked.connect(self.add_annotation)
        self.toolbar.addWidget(self.annotate_btn)

        self.toolbar.addSeparator()

        # Home button
        self.home_btn = QPushButton("← Home")
        self.home_btn.clicked.connect(self.go_home)
        self.toolbar.addWidget(self.home_btn)

    def previous_page(self):
        """Go to previous page"""
        if self.current_page > 0:
            self.go_to_page(self.current_page - 1)

    def next_page(self):
        """Go to next page"""
        if self.pdf_view.doc and self.current_page < self.pdf_view.total_pages - 1:
            self.go_to_page(self.current_page + 1)

    def go_to_page_spin(self):
        """Go to page from spin box"""
        page = self.page_spin.value() - 1
        self.go_to_page(page)

    def go_to_page(self, page_num: int):
        """Go to specific page"""
        if self.pdf_view.doc and 0 <= page_num < self.pdf_view.total_pages:
            self.current_page = page_num
            # Scroll to page
            if page_num < len(self.pdf_view.page_labels):
                self.pdf_view.scroll_area.ensureWidgetVisible(
                    self.pdf_view.page_labels[page_num]
                )
            self.update_status()

    def on_page_selected(self, item):
        """Handle page selection from list"""
        index = self.pages_list.row(item)
        self.go_to_page(index)

    def on_bookmark_selected(self, item):
        """Handle bookmark selection"""
        index = self.bookmarks_list.row(item)
        if 0 <= index < len(self.bookmarks):
            self.go_to_page(self.bookmarks[index]["page"])

    def update_pages_list(self):
        """Update pages list"""
        self.pages_list.clear()
        for i in range(self.pdf_view.total_pages):
            self.pages_list.addItem(f"Page {i + 1}")

    def update_bookmarks_list(self):
        """Update bookmarks list"""
        self.bookmarks_list.clear()
        for bm in self.bookmarks:
            self.bookmarks_list.addItem(bm["title"])

    def open_file_dialog(self):
        """Open file dialog"""
        if self.check_save():
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Open PDF", "", "PDF Files (*.pdf);;All Files (*)"
            )
            if file_path:
                self.open_pdf(file_path)

    def open_pdf(self, file_path):
        """Open PDF file"""
        if not PYMUPDF_AVAILABLE:
            QMessageBox.critical(
                self, "Error", "PyMuPDF not installed. Run: pip install PyMuPDF"
            )
            return

        try:
            # Close existing
            self.pdf_view.close_document()

            # Load new
            self.pdf_view.load_pdf(file_path)
            self.current_file = file_path
            self.is_modified = False

            # Update navigation
            self.page_spin.setMaximum(self.pdf_view.total_pages)
            self.update_pages_list()
            self.bookmarks = []  # Clear bookmarks for new file
            self.update_bookmarks_list()

            # Update UI
            self.update_title()
            self.update_status()

            if self.parent:
                self.parent.status_bar.showMessage(f"Opened: {file_path}", 3000)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open PDF:\n{str(e)}")

    def save(self):
        """Save PDF"""
        if not self.current_file:
            self.save_as_dialog()
        else:
            self.save_as(self.current_file)

    def save_annotations(self):
        """Auto-save annotations"""
        if self.is_modified and self.current_file:
            # Save annotations to a sidecar file
            annotations_file = self.current_file + ".annotations.json"
            try:
                import json

                with open(annotations_file, "w") as f:
                    json.dump(
                        {"annotations": self.annotations, "bookmarks": self.bookmarks},
                        f,
                        indent=2,
                    )
            except:
                pass

    def save_as_dialog(self):
        """Save as dialog"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save PDF", "", "PDF Files (*.pdf);;All Files (*)"
        )
        if file_path:
            self.save_as(file_path)

    def save_as(self, file_path):
        """Save PDF to file"""
        if not PYMUPDF_AVAILABLE or not self.pdf_view.doc:
            return

        try:
            # Save document
            self.pdf_view.doc.save(file_path)

            self.current_file = file_path
            self.is_modified = False
            self.update_title()

            if self.parent:
                self.parent.status_bar.showMessage(f"Saved: {file_path}", 3000)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save PDF:\n{str(e)}")

    def export_pdf(self):
        """Export PDF to other formats"""
        if not self.pdf_view.doc:
            QMessageBox.warning(self, "No File", "No PDF is currently open.")
            return

        formats = ["Images (PNG)", "Text", "HTML", "Word (DOCX)", "Markdown"]
        format_choice, ok = QInputDialog.getItem(
            self, "Export", "Export format:", formats, 0, False
        )

        if ok:
            if format_choice == "Images (PNG)":
                self.export_as_images()
            elif format_choice == "Text":
                self.export_as_text()
            elif format_choice == "HTML":
                self.export_as_html()
            elif format_choice == "Word (DOCX)":
                self.export_as_docx()
            elif format_choice == "Markdown":
                self.export_as_markdown()

    def export_as_images(self):
        """Export PDF pages as images"""
        folder = QFileDialog.getExistingDirectory(self, "Select Export Folder")
        if folder and self.pdf_view.doc:
            try:
                for page_num in range(self.pdf_view.total_pages):
                    page = self.pdf_view.doc[page_num]
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    output_path = os.path.join(folder, f"page_{page_num + 1:03d}.png")
                    pix.save(output_path)

                QMessageBox.information(
                    self,
                    "Export Complete",
                    f"Exported {self.pdf_view.total_pages} pages to {folder}",
                )

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Export failed:\n{str(e)}")

    def export_as_text(self):
        """Export PDF as text"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export as Text", "", "Text Files (*.txt);;All Files (*)"
        )
        if file_path and self.pdf_view.doc:
            try:
                text = []
                for page_num in range(self.pdf_view.total_pages):
                    page = self.pdf_view.doc[page_num]
                    text.append(f"--- Page {page_num + 1} ---")
                    text.append(page.get_text())
                    text.append("")

                with open(file_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(text))

                QMessageBox.information(
                    self, "Export Complete", f"Exported to {file_path}"
                )

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Export failed:\n{str(e)}")

    def export_as_html(self):
        """Export PDF as HTML"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export as HTML", "", "HTML Files (*.html);;All Files (*)"
        )
        if file_path and self.pdf_view.doc:
            try:
                html_content = ["<html><body>"]

                for page_num in range(self.pdf_view.total_pages):
                    page = self.pdf_view.doc[page_num]
                    html_content.append(f"<h2>Page {page_num + 1}</h2>")
                    html_content.append(f"<pre>{page.get_text()}</pre>")
                    html_content.append("<hr>")

                html_content.append("</body></html>")

                with open(file_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(html_content))

                QMessageBox.information(
                    self, "Export Complete", f"Exported to {file_path}"
                )

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Export failed:\n{str(e)}")

    def export_as_docx(self):
        """Export PDF as Word document"""
        try:
            from docx import Document

            file_path, _ = QFileDialog.getSaveFileName(
                self, "Export as Word", "", "Word Documents (*.docx);;All Files (*)"
            )
            if file_path and self.pdf_view.doc:
                doc = Document()
                for page_num in range(self.pdf_view.total_pages):
                    page = self.pdf_view.doc[page_num]
                    text = page.get_text()
                    if text.strip():
                        doc.add_paragraph(text)
                doc.save(file_path)
                QMessageBox.information(
                    self, "Export Complete", f"Exported to {file_path}"
                )
        except ImportError:
            QMessageBox.warning(self, "Error", "python-docx not installed")

    def export_as_markdown(self):
        """Export PDF as Markdown"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export as Markdown", "", "Markdown Files (*.md);;All Files (*)"
        )
        if file_path and self.pdf_view.doc:
            try:
                md_content = [f"# {Path(self.current_file).name}\n"]
                for page_num in range(self.pdf_view.total_pages):
                    page = self.pdf_view.doc[page_num]
                    md_content.append(f"## Page {page_num + 1}\n")
                    md_content.append(page.get_text())
                    md_content.append("\n---\n")
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(md_content))
                QMessageBox.information(
                    self, "Export Complete", f"Exported to {file_path}"
                )
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Export failed:\n{str(e)}")

    def extract_text(self):
        """Extract and display text from current page"""
        if not self.pdf_view.doc:
            QMessageBox.warning(self, "No File", "No PDF is currently open.")
            return

        # Extract text from all pages
        all_text = []
        for page_num in range(self.pdf_view.total_pages):
            page = self.pdf_view.doc[page_num]
            all_text.append(f"=== PAGE {page_num + 1} ===")
            all_text.append(page.get_text())
            all_text.append("")

        # Show in dialog
        dialog = QDialog(self)
        dialog.setWindowTitle("Extracted Text")
        dialog.setMinimumSize(600, 500)

        layout = QVBoxLayout(dialog)

        text_edit = QTextEdit()
        text_edit.setPlainText("\n".join(all_text))
        text_edit.setReadOnly(True)
        layout.addWidget(text_edit)

        # Copy button
        btn_layout = QHBoxLayout()
        copy_btn = QPushButton("Copy All")
        copy_btn.clicked.connect(
            lambda: QApplication.clipboard().setText("\n".join(all_text))
        )
        btn_layout.addWidget(copy_btn)

        save_btn = QPushButton("Save as Text")
        save_btn.clicked.connect(lambda: self.export_as_text())
        btn_layout.addWidget(save_btn)

        layout.addLayout(btn_layout)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        dialog.exec()

    def add_annotation(self):
        """Add annotation to PDF"""
        if not self.pdf_view.doc:
            QMessageBox.warning(self, "No File", "No PDF is currently open.")
            return

        types = ["Highlight", "Underline", "Comment", "Text", "Rectangle", "Circle"]
        annotation_type, ok = QInputDialog.getItem(
            self, "Add Annotation", "Annotation type:", types, 0, False
        )

        if ok:
            text, ok2 = QInputDialog.getText(self, "Annotation", "Text (optional):")

            annotation = {
                "type": annotation_type,
                "page": self.current_page,
                "text": text if ok2 else "",
                "date": datetime.now().isoformat(),
            }

            self.annotations.append(annotation)
            self.is_modified = True
            self.update_title()

            QMessageBox.information(
                self,
                "Annotation Added",
                f"Added {annotation_type} annotation to page {self.current_page + 1}",
            )

    def show_search_dialog(self):
        """Show search dialog"""
        if not self.pdf_view.doc:
            QMessageBox.warning(self, "No File", "No PDF is currently open.")
            return
        dialog = SearchDialog(self)
        dialog.set_editor(self)
        dialog.exec()

    def show_bookmarks(self):
        """Show bookmarks dialog"""
        dialog = BookmarkDialog(self, self.bookmarks)
        dialog.exec()

    def change_zoom(self, zoom_text):
        """Change zoom level"""
        if zoom_text == "Fit Width":
            self.pdf_view.update_zoom(1.0)
        else:
            try:
                zoom = int(zoom_text.replace("%", "")) / 100.0
                self.pdf_view.update_zoom(zoom)
                self.update_status()
            except ValueError:
                pass

    def check_save(self):
        """Check if save is needed"""
        if self.is_modified:
            reply = QMessageBox.question(
                self,
                "Unsaved Changes",
                "The PDF has unsaved annotations. Save now?",
                QMessageBox.StandardButton.Save
                | QMessageBox.StandardButton.Discard
                | QMessageBox.StandardButton.Cancel,
            )

            if reply == QMessageBox.StandardButton.Save:
                self.save()
                return True
            elif reply == QMessageBox.StandardButton.Cancel:
                return False

        return True

    def update_title(self):
        """Update window title"""
        if self.parent:
            filename = "Untitled"
            if self.current_file:
                filename = Path(self.current_file).name
            modified = "*" if self.is_modified else ""
            self.parent.setWindowTitle(
                f"Office Pro - PDF Editor - {filename}{modified}"
            )

    def update_status(self):
        """Update status bar"""
        if self.pdf_view.doc:
            self.page_label.setText(
                f"Page {self.current_page + 1} of {self.pdf_view.total_pages}"
            )
            self.file_label.setText(
                Path(self.current_file).name if self.current_file else "No file"
            )
            zoom = int(self.pdf_view.zoom_level * 100)
            self.zoom_label.setText(f"Zoom: {zoom}%")
        else:
            self.page_label.setText("Page: -")
            self.file_label.setText("No file open")
            self.zoom_label.setText("Zoom: 100%")

    def go_home(self):
        """Return to main menu"""
        if self.check_save():
            self.pdf_view.close_document()
            self.current_file = None
            self.annotations.clear()
            self.bookmarks.clear()
            self.is_modified = False
            if self.parent:
                self.parent.show_main_menu()

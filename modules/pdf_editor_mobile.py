"""
PDF Editor Mobile Module
Kivy-based mobile-optimized PDF viewer - Minimal version for Android
"""

import os
from pathlib import Path

from kivy.uix.screenmanager import Screen
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.textinput import TextInput
from kivy.utils import platform


class PDFEditorMobile(Screen):
    """Mobile PDF editor screen - Basic text viewer"""

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.current_file = None
        self.file_content = ""
        self.build_ui()

    def build_ui(self):
        """Build mobile UI"""
        layout = BoxLayout(orientation="vertical")

        # Toolbar
        toolbar = GridLayout(cols=4, size_hint_y=0.08, spacing=2, padding=2)
        toolbar.add_widget(Button(text="Open", on_press=self.open_file))
        toolbar.add_widget(Button(text="Info", on_press=self.show_info))
        toolbar.add_widget(Button(text="Back", on_press=self.go_back))
        layout.add_widget(toolbar)

        # PDF display area (text only for basic version)
        scroll = ScrollView()
        self.content_label = Label(
            text="Open a PDF or text file to view",
            size_hint_y=None,
            text_size=(400, None),
            halign="left",
            valign="top",
        )
        self.content_label.bind(texture_size=self.content_label.setter("size"))
        scroll.add_widget(self.content_label)
        layout.add_widget(scroll)

        # Filename
        self.filename_label = Label(
            text="No file open", size_hint_y=0.05, font_size="12sp"
        )
        layout.add_widget(self.filename_label)

        self.add_widget(layout)

    def open_file(self, instance):
        """Open file dialog"""
        content = BoxLayout(orientation="vertical")
        filechooser = FileChooserListView(
            path=self.get_documents_path(), filters=["*.pdf", "*.txt"]
        )
        content.add_widget(filechooser)

        btn_layout = BoxLayout(size_hint_y=0.1)
        btn_layout.add_widget(
            Button(
                text="Open", on_press=lambda x: self.load_file(filechooser.selection)
            )
        )
        btn_layout.add_widget(
            Button(text="Cancel", on_press=lambda x: self.popup.dismiss())
        )
        content.add_widget(btn_layout)

        self.popup = Popup(title="Open File", content=content, size_hint=(0.9, 0.9))
        self.popup.open()

    def get_documents_path(self):
        """Get documents path"""
        if platform == "android":
            try:
                from android.storage import primary_external_storage_path

                return primary_external_storage_path()
            except:
                return os.path.expanduser("~")
        else:
            return os.path.expanduser("~")

    def load_file(self, selection):
        """Load selected file"""
        if not selection:
            return

        file_path = selection[0]

        try:
            if file_path.endswith(".txt"):
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    self.file_content = f.read()
                self.content_label.text = self.file_content[:5000]  # Limit display
                self.current_file = file_path
                self.filename_label.text = Path(file_path).name
            elif file_path.endswith(".pdf"):
                # Basic PDF - just show file info
                self.file_content = f"PDF File: {Path(file_path).name}\n\nNote: Full PDF rendering requires PyMuPDF library.\n\nFile path: {file_path}"
                self.content_label.text = self.file_content
                self.current_file = file_path
                self.filename_label.text = Path(file_path).name

            self.popup.dismiss()

        except Exception as e:
            self.show_error(f"Error loading file: {str(e)}")

    def show_info(self, instance):
        """Show file info"""
        if self.current_file:
            info = f"File: {Path(self.current_file).name}\n"
            info += f"Path: {self.current_file}\n"
            try:
                size = os.path.getsize(self.current_file)
                info += f"Size: {size} bytes"
            except:
                pass
        else:
            info = "No file open"

        popup = Popup(title="File Info", content=Label(text=info), size_hint=(0.8, 0.5))
        popup.open()

    def show_error(self, error):
        """Show error popup"""
        popup = Popup(title="Error", content=Label(text=error), size_hint=(0.8, 0.4))
        popup.open()

    def go_back(self, instance):
        """Go back to main menu"""
        self.manager.current = "menu"

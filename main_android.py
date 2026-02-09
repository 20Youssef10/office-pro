#!/usr/bin/env python3
"""
Office Pro - Android Entry Point
This is the main entry point for Android APK conversion using Kivy
"""

import os
import sys

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Kivy imports for Android compatibility
try:
    from kivy.app import App
    from kivy.uix.boxlayout import BoxLayout
    from kivy.uix.button import Button
    from kivy.uix.label import Label
    from kivy.uix.gridlayout import GridLayout
    from kivy.uix.screenmanager import ScreenManager, Screen
    from kivy.core.window import Window
    from kivy.config import Config
    from kivy.properties import ObjectProperty
    from kivy.uix.popup import Popup
    from kivy.uix.filechooser import FileChooserListView
    from kivy.uix.textinput import TextInput
    from kivy.uix.scrollview import ScrollView
    from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelHeader
    from kivy.graphics import Color, Rectangle
    from kivy.utils import platform

    KIVY_AVAILABLE = True
except ImportError:
    KIVY_AVAILABLE = False
    print("Kivy not available, falling back to console mode")

# Mobile-specific imports
if platform == "android":
    from android.permissions import request_permissions, Permission
    from android.storage import primary_external_storage_path
    from jnius import autoclass

    # Request permissions on Android
    request_permissions(
        [
            Permission.READ_EXTERNAL_STORAGE,
            Permission.WRITE_EXTERNAL_STORAGE,
            Permission.INTERNET,
        ]
    )

# Import Office Pro modules
from modules.word_processor_mobile import WordProcessorMobile
from modules.spreadsheet_mobile import SpreadsheetMobile
from modules.pdf_editor_mobile import PDFEditorMobile


class MainMenuScreen(Screen):
    """Main menu screen for mobile app"""

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.build_ui()

    def build_ui(self):
        """Build the main menu UI"""
        layout = BoxLayout(orientation="vertical", padding=20, spacing=10)

        # Title
        title = Label(
            text="[b]Office Pro[/b]",
            markup=True,
            font_size="32sp",
            size_hint_y=0.15,
            color=(0.17, 0.34, 0.6, 1),  # Office blue
        )
        layout.add_widget(title)

        subtitle = Label(
            text="Professional Office Suite",
            font_size="16sp",
            size_hint_y=0.08,
            color=(0.4, 0.4, 0.4, 1),
        )
        layout.add_widget(subtitle)

        # Buttons grid
        buttons_grid = GridLayout(cols=2, spacing=15, padding=20)

        # Word Processor
        word_btn = Button(
            text="Word\nProcessor",
            font_size="18sp",
            background_color=(0.17, 0.34, 0.6, 1),
            background_normal="",
            on_press=self.open_word,
        )
        buttons_grid.add_widget(word_btn)

        # Spreadsheet
        spreadsheet_btn = Button(
            text="Spreadsheet",
            font_size="18sp",
            background_color=(0.13, 0.45, 0.27, 1),
            background_normal="",
            on_press=self.open_spreadsheet,
        )
        buttons_grid.add_widget(spreadsheet_btn)

        # Presentation
        presentation_btn = Button(
            text="Presentation",
            font_size="18sp",
            background_color=(0.82, 0.28, 0.15, 1),
            background_normal="",
            on_press=self.open_presentation,
        )
        buttons_grid.add_widget(presentation_btn)

        # PDF Editor
        pdf_btn = Button(
            text="PDF\nEditor",
            font_size="18sp",
            background_color=(0.96, 0.06, 0.01, 1),
            background_normal="",
            on_press=self.open_pdf,
        )
        buttons_grid.add_widget(pdf_btn)

        layout.add_widget(buttons_grid)

        # Info label
        info = Label(
            text="Tap an app to launch",
            font_size="14sp",
            size_hint_y=0.1,
            color=(0.5, 0.5, 0.5, 1),
        )
        layout.add_widget(info)

        self.add_widget(layout)

    def open_word(self, instance):
        """Open Word Processor"""
        self.manager.current = "word"

    def open_spreadsheet(self, instance):
        """Open Spreadsheet"""
        self.manager.current = "spreadsheet"

    def open_presentation(self, instance):
        """Open Presentation"""
        self.show_popup("Presentation", "Coming soon in v2.0!")

    def open_pdf(self, instance):
        """Open PDF Editor"""
        self.manager.current = "pdf"

    def show_popup(self, title, message):
        """Show popup message"""
        popup = Popup(title=title, content=Label(text=message), size_hint=(0.8, 0.4))
        popup.open()


class OfficeProApp(App):
    """Main Office Pro Android Application"""

    def build(self):
        """Build the application"""
        # Set window background color
        Window.clearcolor = (0.95, 0.95, 0.95, 1)

        # Create screen manager
        sm = ScreenManager()

        # Add screens
        sm.add_widget(MainMenuScreen(name="menu"))
        sm.add_widget(WordProcessorMobile(name="word"))
        sm.add_widget(SpreadsheetMobile(name="spreadsheet"))
        sm.add_widget(PDFEditorMobile(name="pdf"))

        return sm

    def on_pause(self):
        """Handle app pause (Android)"""
        return True

    def on_resume(self):
        """Handle app resume (Android)"""
        pass


def main():
    """Main entry point"""
    if KIVY_AVAILABLE:
        # Run Kivy app
        OfficeProApp().run()
    else:
        # Fallback to console mode
        print("=" * 50)
        print("Office Pro - Console Mode")
        print("=" * 50)
        print("\nKivy not available. To run the full GUI:")
        print("1. Install Kivy: pip install kivy")
        print("2. Run: python main_android.py")
        print("\nFor Android APK, use Buildozer:")
        print("1. Install buildozer: pip install buildozer")
        print("2. Run: buildozer android debug")


if __name__ == "__main__":
    main()

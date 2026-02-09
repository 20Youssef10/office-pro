"""
PDF Editor Mobile Module
Kivy-based mobile-optimized PDF viewer
"""

import os
from pathlib import Path

try:
    from kivy.uix.screenmanager import Screen
    from kivy.uix.boxlayout import BoxLayout
    from kivy.uix.textinput import TextInput
    from kivy.uix.button import Button
    from kivy.uix.gridlayout import GridLayout
    from kivy.uix.scrollview import ScrollView
    from kivy.uix.popup import Popup
    from kivy.uix.label import Label
    from kivy.uix.filechooser import FileChooserListView
    from kivy.uix.image import Image
    from kivy.utils import platform
    from kivy.core.image import Image as CoreImage
    from io import BytesIO

    KIVY_AVAILABLE = True
except ImportError:
    KIVY_AVAILABLE = False

try:
    import fitz  # PyMuPDF

    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False


class PDFEditorMobile(Screen):
    """Mobile PDF editor screen"""

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.current_file = None
        self.current_page = 0
        self.total_pages = 0
        self.doc = None
        self.zoom_level = 1.0
        self.build_ui()

    def build_ui(self):
        """Build mobile UI"""
        layout = BoxLayout(orientation="vertical")

        # Toolbar
        toolbar = GridLayout(cols=6, size_hint_y=0.08, spacing=2, padding=2)
        toolbar.add_widget(Button(text="Open", on_press=self.open_file))
        toolbar.add_widget(Button(text="Prev", on_press=self.prev_page))
        toolbar.add_widget(Button(text="Next", on_press=self.next_page))
        toolbar.add_widget(Button(text="Text", on_press=self.extract_text))
        toolbar.add_widget(Button(text="Zoom", on_press=self.zoom_dialog))
        toolbar.add_widget(Button(text="Back", on_press=self.go_back))
        layout.add_widget(toolbar)

        # Page info
        self.page_info = Label(text="Page: -", size_hint_y=0.05, font_size="14sp")
        layout.add_widget(self.page_info)

        # PDF display
        scroll = ScrollView()
        self.pdf_image = Image(source="", allow_stretch=True, keep_ratio=True)
        scroll.add_widget(self.pdf_image)
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
            path=self.get_documents_path(), filters=["*.pdf"]
        )
        content.add_widget(filechooser)

        btn_layout = BoxLayout(size_hint_y=0.1)
        btn_layout.add_widget(
            Button(text="Open", on_press=lambda x: self.load_pdf(filechooser.selection))
        )
        btn_layout.add_widget(
            Button(text="Cancel", on_press=lambda x: self.popup.dismiss())
        )
        content.add_widget(btn_layout)

        self.popup = Popup(title="Open PDF", content=content, size_hint=(0.9, 0.9))
        self.popup.open()

    def get_documents_path(self):
        """Get documents path"""
        if platform == "android":
            from android.storage import primary_external_storage_path

            return primary_external_storage_path()
        else:
            return os.path.expanduser("~")

    def load_pdf(self, selection):
        """Load selected PDF"""
        if not selection:
            return

        file_path = selection[0]

        if not PYMUPDF_AVAILABLE:
            self.show_error("PyMuPDF not available")
            self.popup.dismiss()
            return

        try:
            # Close previous document
            if self.doc:
                self.doc.close()

            # Open new document
            self.doc = fitz.open(file_path)
            self.total_pages = len(self.doc)
            self.current_page = 0
            self.current_file = file_path

            self.popup.dismiss()
            self.render_page()
            self.update_ui()

        except Exception as e:
            self.show_error(f"Error loading PDF: {str(e)}")

    def render_page(self):
        """Render current page"""
        if not self.doc or not PYMUPDF_AVAILABLE:
            return

        try:
            page = self.doc[self.current_page]

            # Render at current zoom
            mat = fitz.Matrix(self.zoom_level, self.zoom_level)
            pix = page.get_pixmap(matrix=mat)

            # Convert to image
            img_data = pix.tobytes("png")
            img = CoreImage(BytesIO(img_data), ext="png")

            # Update image
            self.pdf_image.texture = img.texture
            self.pdf_image.size = img.size

        except Exception as e:
            self.show_error(f"Error rendering page: {str(e)}")

    def prev_page(self, instance):
        """Go to previous page"""
        if self.doc and self.current_page > 0:
            self.current_page -= 1
            self.render_page()
            self.update_ui()

    def next_page(self, instance):
        """Go to next page"""
        if self.doc and self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.render_page()
            self.update_ui()

    def update_ui(self):
        """Update UI elements"""
        if self.doc:
            self.page_info.text = f"Page {self.current_page + 1} of {self.total_pages}"
            self.filename_label.text = Path(self.current_file).name
        else:
            self.page_info.text = "Page: -"
            self.filename_label.text = "No file open"

    def zoom_dialog(self, instance):
        """Show zoom dialog"""
        content = BoxLayout(orientation="vertical")

        zoom_levels = ["50%", "75%", "100%", "125%", "150%", "200%"]
        for level in zoom_levels:
            btn = Button(text=level, size_hint_y=None, height=50)
            btn.bind(on_press=lambda x, l=level: self.set_zoom(l))
            content.add_widget(btn)

        btn_layout = BoxLayout(size_hint_y=0.1)
        btn_layout.add_widget(
            Button(text="Close", on_press=lambda x: self.popup.dismiss())
        )
        content.add_widget(btn_layout)

        self.popup = Popup(title="Zoom", content=content, size_hint=(0.7, 0.7))
        self.popup.open()

    def set_zoom(self, level):
        """Set zoom level"""
        try:
            self.zoom_level = int(level.replace("%", "")) / 100.0
            self.render_page()
            self.popup.dismiss()
        except:
            pass

    def extract_text(self, instance):
        """Extract text from current page"""
        if not self.doc:
            self.show_error("No PDF open")
            return

        try:
            page = self.doc[self.current_page]
            text = page.get_text()

            # Show text in popup
            content = BoxLayout(orientation="vertical")
            scroll = ScrollView()
            text_label = Label(text=text, size_hint_y=None, text_size=(400, None))
            text_label.bind(texture_size=text_label.setter("size"))
            scroll.add_widget(text_label)
            content.add_widget(scroll)

            btn_layout = BoxLayout(size_hint_y=0.1)
            btn_layout.add_widget(
                Button(text="Close", on_press=lambda x: self.popup.dismiss())
            )
            content.add_widget(btn_layout)

            self.popup = Popup(
                title=f"Text - Page {self.current_page + 1}",
                content=content,
                size_hint=(0.9, 0.7),
            )
            self.popup.open()

        except Exception as e:
            self.show_error(f"Error extracting text: {str(e)}")

    def show_error(self, error):
        """Show error popup"""
        popup = Popup(title="Error", content=Label(text=error), size_hint=(0.8, 0.4))
        popup.open()

    def go_back(self, instance):
        """Go back to main menu"""
        if self.doc:
            self.doc.close()
            self.doc = None
        self.manager.current = "menu"

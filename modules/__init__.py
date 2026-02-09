"""
Office Pro - Modules Package
"""

from .word_processor import WordProcessor
from .spreadsheet import SpreadsheetEditor
from .presentation import PresentationEditor
from .pdf_editor import PDFEditor
from .file_manager import FileManager

__all__ = [
    "WordProcessor",
    "SpreadsheetEditor",
    "PresentationEditor",
    "PDFEditor",
    "FileManager",
]

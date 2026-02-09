# Office Pro - Professional Office Suite

A comprehensive office application suite built with Python, featuring professional document editing capabilities.

## Features

### Word Processor
- Rich text editing with formatting options
- Support for DOCX, DOC, ODT, RTF, and TXT files
- Font styling (bold, italic, underline)
- Text alignment and color options
- Bullet and numbered lists
- Print preview and page setup

### Spreadsheet Editor
- Full-featured spreadsheet editing
- Support for XLSX, XLS, ODS, and CSV files
- Cell formatting and styling
- Formula support (SUM, AVERAGE, COUNT, MAX, MIN)
- Row and column management
- Auto-calculation

### Presentation Editor
- Create and edit presentations
- Support for PPTX, PPT, and ODP files
- Slide management
- Text boxes and shapes
- Slide transitions (coming soon)
- Presentation mode (coming soon)

### PDF Editor
- View PDF documents
- Extract text from PDFs
- Export to images, text, or HTML
- Annotation support
- Zoom and navigation controls

## Installation

### Prerequisites
- Python 3.8 or higher
- pip package manager

### Install Dependencies

```bash
pip install -r requirements.txt
```

### Required Packages
- PyQt6 - GUI framework
- python-docx - Word document handling
- openpyxl - Excel spreadsheet handling
- python-pptx - PowerPoint presentation handling
- PyMuPDF (fitz) - PDF handling
- pandas - Data manipulation
- Pillow - Image processing

## Usage

### Starting the Application

```bash
cd "Office Pro"
python office_pro.py
```

### Basic Operations

1. **Create New Document**: Click on Word Processor, Spreadsheet, Presentation, or PDF Editor from the main menu
2. **Open Existing File**: Use File > Open or Ctrl+O
3. **Save File**: Use File > Save or Ctrl+S
4. **Switch Between Apps**: Use View > Home to return to main menu

### Word Processor
- Type and format text using the toolbar
- Change fonts, sizes, and colors
- Use alignment buttons for text positioning
- Save as DOCX, TXT, or HTML

### Spreadsheet
- Enter data in cells
- Use formulas starting with = (e.g., =SUM(A1:B5))
- Format cells with bold, colors, etc.
- Navigate with arrow keys or mouse

### Presentation
- Add slides with the + Slide button
- Add text boxes and shapes
- Navigate between slides
- Save as PPTX format

### PDF Editor
- Open PDF files for viewing
- Extract text content
- Export to various formats
- Add annotations

## Keyboard Shortcuts

- **Ctrl+N**: New document
- **Ctrl+O**: Open file
- **Ctrl+S**: Save file
- **Ctrl+Z**: Undo (in supported modules)
- **Ctrl+Y**: Redo (in supported modules)
- **Ctrl+B**: Bold
- **Ctrl+I**: Italic
- **Ctrl+U**: Underline
- **Alt+F4**: Exit application

## Project Structure

```
Office Pro/
├── office_pro.py          # Main application launcher
├── requirements.txt       # Python dependencies
├── README.md             # This file
├── modules/
│   ├── __init__.py
│   ├── word_processor.py # Word document editor
│   ├── spreadsheet.py    # Spreadsheet editor
│   ├── presentation.py   # Presentation editor
│   ├── pdf_editor.py     # PDF viewer/editor
│   └── file_manager.py   # File operations manager
├── assets/               # Application assets
└── templates/            # Document templates
```

## Supported File Formats

### Word Processor
- .docx (Microsoft Word)
- .doc (Microsoft Word 97-2003)
- .odt (OpenDocument Text)
- .rtf (Rich Text Format)
- .txt (Plain Text)

### Spreadsheet
- .xlsx (Microsoft Excel)
- .xls (Microsoft Excel 97-2003)
- .ods (OpenDocument Spreadsheet)
- .csv (Comma Separated Values)

### Presentation
- .pptx (Microsoft PowerPoint)
- .ppt (Microsoft PowerPoint 97-2003)
- .odp (OpenDocument Presentation)

### PDF Editor
- .pdf (Portable Document Format)

## Development

### Adding New Features

1. Create new module in `modules/` directory
2. Inherit from QWidget for UI components
3. Implement standard interface methods:
   - `new_document()`: Create new file
   - `open_file()`: Open existing file
   - `save()`: Save current file
   - `check_save()`: Check for unsaved changes

### Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## License

This project is open source and available under the MIT License.

## Troubleshooting

### Common Issues

1. **Module not found errors**: Ensure all requirements are installed
2. **PDF not opening**: Install PyMuPDF with `pip install PyMuPDF`
3. **DOCX files not working**: Install python-docx with `pip install python-docx`

### Support

For issues and feature requests, please use the GitHub issue tracker.

## Credits

Built with:
- Python 3
- PyQt6
- python-docx
- openpyxl
- python-pptx
- PyMuPDF
- pandas

© 2026 Office Pro Team

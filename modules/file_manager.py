"""
File Manager Module
Handles file operations, recent files, and file type detection.
"""

import os
import json
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional


class FileManager:
    """Manages file operations and recent files"""

    def __init__(self, config_dir: Optional[str] = None):
        """Initialize file manager"""
        self.config_dir = config_dir or os.path.expanduser("~/.office_pro")
        self.recent_files_path = os.path.join(self.config_dir, "recent_files.json")
        self.max_recent_files = 20

        # Ensure config directory exists
        os.makedirs(self.config_dir, exist_ok=True)

        # Load recent files
        self.recent_files = self.load_recent_files()

    def load_recent_files(self) -> List[Dict]:
        """Load recent files list"""
        if os.path.exists(self.recent_files_path):
            try:
                with open(self.recent_files_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                return []
        return []

    def save_recent_files(self):
        """Save recent files list"""
        try:
            with open(self.recent_files_path, "w", encoding="utf-8") as f:
                json.dump(self.recent_files, f, indent=2)
        except Exception as e:
            print(f"Error saving recent files: {e}")

    def add_recent_file(self, file_path: str, file_type: str):
        """Add file to recent files list"""
        # Remove if already exists
        self.recent_files = [f for f in self.recent_files if f["path"] != file_path]

        # Add to beginning
        self.recent_files.insert(
            0,
            {
                "path": file_path,
                "type": file_type,
                "opened": datetime.now().isoformat(),
                "name": Path(file_path).name,
            },
        )

        # Limit size
        self.recent_files = self.recent_files[: self.max_recent_files]

        # Save
        self.save_recent_files()

    def get_recent_files(self, file_type: Optional[str] = None) -> List[Dict]:
        """Get recent files, optionally filtered by type"""
        if file_type:
            return [f for f in self.recent_files if f["type"] == file_type]
        return self.recent_files

    def clear_recent_files(self):
        """Clear recent files list"""
        self.recent_files = []
        self.save_recent_files()

    def get_file_type(self, file_path: str) -> str:
        """Determine file type from extension"""
        ext = Path(file_path).suffix.lower()

        file_types = {
            ".docx": "word",
            ".doc": "word",
            ".odt": "word",
            ".rtf": "word",
            ".txt": "word",
            ".xlsx": "spreadsheet",
            ".xls": "spreadsheet",
            ".ods": "spreadsheet",
            ".csv": "spreadsheet",
            ".pptx": "presentation",
            ".ppt": "presentation",
            ".odp": "presentation",
            ".pdf": "pdf",
        }

        return file_types.get(ext, "unknown")

    def file_exists(self, file_path: str) -> bool:
        """Check if file exists"""
        return os.path.exists(file_path)

    def get_file_size(self, file_path: str) -> int:
        """Get file size in bytes"""
        try:
            return os.path.getsize(file_path)
        except:
            return 0

    def format_file_size(self, size_bytes: int) -> str:
        """Format file size for display"""
        for unit in ["B", "KB", "MB", "GB"]:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} TB"

    def get_file_info(self, file_path: str) -> Dict:
        """Get comprehensive file information"""
        try:
            stat = os.stat(file_path)
            return {
                "path": file_path,
                "name": Path(file_path).name,
                "type": self.get_file_type(file_path),
                "size": self.format_file_size(stat.st_size),
                "size_bytes": stat.st_size,
                "created": datetime.fromtimestamp(stat.st_ctime).isoformat(),
                "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
                "accessed": datetime.fromtimestamp(stat.st_atime).isoformat(),
                "exists": True,
            }
        except:
            return {
                "path": file_path,
                "name": Path(file_path).name,
                "type": "unknown",
                "exists": False,
            }

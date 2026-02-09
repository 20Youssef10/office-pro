"""
Office Pro - Database Models for All 50 Features
Supports versioning, collaboration, AI features, and more
"""

import sqlite3
import json
import os
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict, Any
from dataclasses import dataclass, asdict
import uuid


class DatabaseManager:
    """Central database manager for Office Pro"""

    def __init__(self, db_path: str = None):
        self.db_path = db_path or os.path.expanduser("~/.office_pro/office_pro.db")
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self.create_tables()

    def create_tables(self):
        """Create all necessary tables for 50 features"""
        cursor = self.conn.cursor()

        # Documents table (for version history - Feature 13)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS documents (
                id TEXT PRIMARY KEY,
                file_path TEXT NOT NULL,
                file_name TEXT NOT NULL,
                file_type TEXT NOT NULL,
                current_version INTEGER DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                modified_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                owner_id TEXT,
                is_shared BOOLEAN DEFAULT 0
            )
        """)

        # Document versions table (Feature 13: Version History)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS document_versions (
                id TEXT PRIMARY KEY,
                document_id TEXT NOT NULL,
                version_number INTEGER NOT NULL,
                content BLOB,
                changes_description TEXT,
                created_by TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                file_size INTEGER,
                FOREIGN KEY (document_id) REFERENCES documents(id)
            )
        """)

        # Comments table (Feature 2: Track Changes & Comments)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS comments (
                id TEXT PRIMARY KEY,
                document_id TEXT NOT NULL,
                user_id TEXT NOT NULL,
                user_name TEXT,
                content TEXT NOT NULL,
                position_x INTEGER,
                position_y INTEGER,
                selection_text TEXT,
                parent_id TEXT,
                is_resolved BOOLEAN DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                modified_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (document_id) REFERENCES documents(id)
            )
        """)

        # Changes/Track changes table (Feature 2)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS track_changes (
                id TEXT PRIMARY KEY,
                document_id TEXT NOT NULL,
                user_id TEXT NOT NULL,
                user_name TEXT,
                change_type TEXT NOT NULL,
                old_content TEXT,
                new_content TEXT,
                position_start INTEGER,
                position_end INTEGER,
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                is_accepted BOOLEAN DEFAULT 0,
                is_rejected BOOLEAN DEFAULT 0,
                FOREIGN KEY (document_id) REFERENCES documents(id)
            )
        """)

        # Templates table (Feature 3: Templates Gallery)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS templates (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                category TEXT NOT NULL,
                description TEXT,
                content BLOB,
                thumbnail_path TEXT,
                is_builtin BOOLEAN DEFAULT 0,
                is_custom BOOLEAN DEFAULT 0,
                created_by TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                usage_count INTEGER DEFAULT 0
            )
        """)

        # Citations table (Feature 6: Citation Manager)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS citations (
                id TEXT PRIMARY KEY,
                document_id TEXT NOT NULL,
                citation_type TEXT NOT NULL,
                style TEXT NOT NULL,
                author TEXT,
                title TEXT,
                journal TEXT,
                year INTEGER,
                volume TEXT,
                issue TEXT,
                pages TEXT,
                doi TEXT,
                url TEXT,
                citation_key TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (document_id) REFERENCES documents(id)
            )
        """)

        # Spreadsheets metadata (Feature 16: Pivot Tables)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS spreadsheets (
                id TEXT PRIMARY KEY,
                file_path TEXT NOT NULL,
                file_name TEXT NOT NULL,
                row_count INTEGER,
                col_count INTEGER,
                last_calculation TIMESTAMP,
                metadata TEXT
            )
        """)

        # Pivot tables (Feature 16)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS pivot_tables (
                id TEXT PRIMARY KEY,
                spreadsheet_id TEXT NOT NULL,
                name TEXT NOT NULL,
                source_range TEXT NOT NULL,
                row_fields TEXT,
                column_fields TEXT,
                data_fields TEXT,
                filter_fields TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (spreadsheet_id) REFERENCES spreadsheets(id)
            )
        """)

        # Data validation rules (Feature 18)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS data_validation_rules (
                id TEXT PRIMARY KEY,
                spreadsheet_id TEXT NOT NULL,
                cell_range TEXT NOT NULL,
                rule_type TEXT NOT NULL,
                criteria TEXT,
                error_message TEXT,
                allow_blank BOOLEAN DEFAULT 1,
                FOREIGN KEY (spreadsheet_id) REFERENCES spreadsheets(id)
            )
        """)

        # Macros (Feature 21)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS macros (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                description TEXT,
                code TEXT NOT NULL,
                shortcut TEXT,
                created_by TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                modified_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                usage_count INTEGER DEFAULT 0
            )
        """)

        # PDF annotations (Feature 39)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS pdf_annotations (
                id TEXT PRIMARY KEY,
                pdf_path TEXT NOT NULL,
                user_id TEXT,
                annotation_type TEXT NOT NULL,
                page_number INTEGER NOT NULL,
                position_x REAL,
                position_y REAL,
                width REAL,
                height REAL,
                content TEXT,
                color TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                modified_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Projects (Feature 48: Project Management)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS projects (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                description TEXT,
                status TEXT DEFAULT 'active',
                start_date DATE,
                end_date DATE,
                owner_id TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Tasks (Feature 48)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS tasks (
                id TEXT PRIMARY KEY,
                project_id TEXT NOT NULL,
                title TEXT NOT NULL,
                description TEXT,
                status TEXT DEFAULT 'todo',
                priority TEXT DEFAULT 'medium',
                assigned_to TEXT,
                due_date DATE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                completed_at TIMESTAMP,
                FOREIGN KEY (project_id) REFERENCES projects(id)
            )
        """)

        # Clipboard history (Feature 41)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS clipboard_history (
                id TEXT PRIMARY KEY,
                content TEXT NOT NULL,
                content_type TEXT DEFAULT 'text',
                source_application TEXT,
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # User settings and preferences
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS user_settings (
                user_id TEXT PRIMARY KEY,
                settings_json TEXT NOT NULL,
                last_modified TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # AI usage tracking (Feature 50)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS ai_usage (
                id TEXT PRIMARY KEY,
                user_id TEXT,
                feature TEXT NOT NULL,
                tokens_used INTEGER,
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Collaboration sessions (Feature 11)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS collaboration_sessions (
                id TEXT PRIMARY KEY,
                document_id TEXT NOT NULL,
                session_token TEXT UNIQUE NOT NULL,
                created_by TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                expires_at TIMESTAMP,
                is_active BOOLEAN DEFAULT 1
            )
        """)

        # Session participants (Feature 11)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS session_participants (
                session_id TEXT,
                user_id TEXT,
                user_name TEXT,
                joined_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                last_activity TIMESTAMP,
                PRIMARY KEY (session_id, user_id),
                FOREIGN KEY (session_id) REFERENCES collaboration_sessions(id)
            )
        """)

        # Insert default templates
        self.insert_default_templates(cursor)

        self.conn.commit()

    def insert_default_templates(self, cursor):
        """Insert default document templates"""
        templates = [
            ("blank", "Blank Document", "General", "Start with a blank page", None, 1),
            (
                "business_letter",
                "Business Letter",
                "Business",
                "Professional letter format",
                None,
                1,
            ),
            ("resume", "Resume/CV", "Career", "Professional resume template", None, 1),
            (
                "report",
                "Business Report",
                "Business",
                "Structured business report",
                None,
                1,
            ),
            ("memo", "Internal Memo", "Business", "Company memorandum format", None, 1),
            (
                "academic_paper",
                "Academic Paper",
                "Academic",
                "Research paper format",
                None,
                1,
            ),
            (
                "cover_letter",
                "Cover Letter",
                "Career",
                "Job application cover letter",
                None,
                1,
            ),
            (
                "meeting_notes",
                "Meeting Notes",
                "Business",
                "Structured meeting notes",
                None,
                1,
            ),
            (
                "budget",
                "Budget Spreadsheet",
                "Finance",
                "Personal or business budget",
                None,
                1,
            ),
            (
                "schedule",
                "Project Schedule",
                "Project",
                "Gantt-style project schedule",
                None,
                1,
            ),
        ]

        for template_id, name, category, description, content, is_builtin in templates:
            cursor.execute(
                """
                INSERT OR IGNORE INTO templates 
                (id, name, category, description, content, is_builtin, is_custom)
                VALUES (?, ?, ?, ?, ?, ?, 0)
            """,
                (template_id, name, category, description, content, is_builtin),
            )

    # Document Version Methods (Feature 13)
    def save_version(
        self,
        document_id: str,
        content: bytes,
        changes_description: str = "",
        created_by: str = "",
    ) -> str:
        """Save a new version of a document"""
        version_id = str(uuid.uuid4())
        cursor = self.conn.cursor()

        # Get next version number
        cursor.execute(
            """
            SELECT MAX(version_number) FROM document_versions 
            WHERE document_id = ?
        """,
            (document_id,),
        )
        result = cursor.fetchone()
        version_number = (result[0] or 0) + 1

        cursor.execute(
            """
            INSERT INTO document_versions 
            (id, document_id, version_number, content, changes_description, created_by, file_size)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
            (
                version_id,
                document_id,
                version_number,
                content,
                changes_description,
                created_by,
                len(content),
            ),
        )

        # Update document's current version
        cursor.execute(
            """
            UPDATE documents SET current_version = ?, modified_at = CURRENT_TIMESTAMP 
            WHERE id = ?
        """,
            (version_number, document_id),
        )

        self.conn.commit()
        return version_id

    def get_versions(self, document_id: str) -> List[Dict]:
        """Get all versions of a document"""
        cursor = self.conn.cursor()
        cursor.execute(
            """
            SELECT * FROM document_versions 
            WHERE document_id = ? 
            ORDER BY version_number DESC
        """,
            (document_id,),
        )

        return [dict(row) for row in cursor.fetchall()]

    def restore_version(self, version_id: str) -> Optional[bytes]:
        """Restore a specific version"""
        cursor = self.conn.cursor()
        cursor.execute(
            """
            SELECT content, document_id, version_number FROM document_versions 
            WHERE id = ?
        """,
            (version_id,),
        )

        row = cursor.fetchone()
        if row:
            # Update document's current version
            cursor.execute(
                """
                UPDATE documents SET current_version = ?, modified_at = CURRENT_TIMESTAMP 
                WHERE id = ?
            """,
                (row["version_number"], row["document_id"]),
            )
            self.conn.commit()
            return row["content"]
        return None

    # Comments Methods (Feature 2)
    def add_comment(
        self,
        document_id: str,
        user_id: str,
        content: str,
        position_x: int = 0,
        position_y: int = 0,
        selection_text: str = "",
        parent_id: str = None,
    ) -> str:
        """Add a comment to a document"""
        comment_id = str(uuid.uuid4())
        cursor = self.conn.cursor()

        cursor.execute(
            """
            INSERT INTO comments 
            (id, document_id, user_id, content, position_x, position_y, selection_text, parent_id)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
            (
                comment_id,
                document_id,
                user_id,
                content,
                position_x,
                position_y,
                selection_text,
                parent_id,
            ),
        )

        self.conn.commit()
        return comment_id

    def get_comments(
        self, document_id: str, include_resolved: bool = False
    ) -> List[Dict]:
        """Get all comments for a document"""
        cursor = self.conn.cursor()

        if include_resolved:
            cursor.execute(
                """
                SELECT * FROM comments WHERE document_id = ? ORDER BY created_at DESC
            """,
                (document_id,),
            )
        else:
            cursor.execute(
                """
                SELECT * FROM comments 
                WHERE document_id = ? AND is_resolved = 0 
                ORDER BY created_at DESC
            """,
                (document_id,),
            )

        return [dict(row) for row in cursor.fetchall()]

    def resolve_comment(self, comment_id: str):
        """Mark a comment as resolved"""
        cursor = self.conn.cursor()
        cursor.execute(
            """
            UPDATE comments SET is_resolved = 1 WHERE id = ?
        """,
            (comment_id,),
        )
        self.conn.commit()

    # Track Changes Methods (Feature 2)
    def record_change(
        self,
        document_id: str,
        user_id: str,
        change_type: str,
        old_content: str,
        new_content: str,
        position_start: int,
        position_end: int,
    ) -> str:
        """Record a tracked change"""
        change_id = str(uuid.uuid4())
        cursor = self.conn.cursor()

        cursor.execute(
            """
            INSERT INTO track_changes 
            (id, document_id, user_id, change_type, old_content, new_content, 
             position_start, position_end)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
            (
                change_id,
                document_id,
                user_id,
                change_type,
                old_content,
                new_content,
                position_start,
                position_end,
            ),
        )

        self.conn.commit()
        return change_id

    def get_changes(self, document_id: str) -> List[Dict]:
        """Get all tracked changes for a document"""
        cursor = self.conn.cursor()
        cursor.execute(
            """
            SELECT * FROM track_changes 
            WHERE document_id = ? AND is_accepted = 0 AND is_rejected = 0
            ORDER BY timestamp ASC
        """,
            (document_id,),
        )

        return [dict(row) for row in cursor.fetchall()]

    def accept_change(self, change_id: str):
        """Accept a tracked change"""
        cursor = self.conn.cursor()
        cursor.execute(
            """
            UPDATE track_changes SET is_accepted = 1 WHERE id = ?
        """,
            (change_id,),
        )
        self.conn.commit()

    def reject_change(self, change_id: str):
        """Reject a tracked change"""
        cursor = self.conn.cursor()
        cursor.execute(
            """
            UPDATE track_changes SET is_rejected = 1 WHERE id = ?
        """,
            (change_id,),
        )
        self.conn.commit()

    # Template Methods (Feature 3)
    def get_templates(self, category: str = None) -> List[Dict]:
        """Get all templates, optionally filtered by category"""
        cursor = self.conn.cursor()

        if category:
            cursor.execute(
                """
                SELECT * FROM templates WHERE category = ? ORDER BY usage_count DESC
            """,
                (category,),
            )
        else:
            cursor.execute(
                "SELECT * FROM templates ORDER BY category, usage_count DESC"
            )

        return [dict(row) for row in cursor.fetchall()]

    def increment_template_usage(self, template_id: str):
        """Increment usage count for a template"""
        cursor = self.conn.cursor()
        cursor.execute(
            """
            UPDATE templates SET usage_count = usage_count + 1 WHERE id = ?
        """,
            (template_id,),
        )
        self.conn.commit()

    # Clipboard History (Feature 41)
    def add_clipboard_item(
        self, content: str, content_type: str = "text", source: str = None
    ) -> str:
        """Add item to clipboard history"""
        item_id = str(uuid.uuid4())
        cursor = self.conn.cursor()

        cursor.execute(
            """
            INSERT INTO clipboard_history (id, content, content_type, source_application)
            VALUES (?, ?, ?, ?)
        """,
            (item_id, content, content_type, source),
        )

        # Keep only last 100 items
        cursor.execute("""
            DELETE FROM clipboard_history WHERE id NOT IN (
                SELECT id FROM clipboard_history ORDER BY timestamp DESC LIMIT 100
            )
        """)

        self.conn.commit()
        return item_id

    def get_clipboard_history(self, limit: int = 20) -> List[Dict]:
        """Get clipboard history"""
        cursor = self.conn.cursor()
        cursor.execute(
            """
            SELECT * FROM clipboard_history 
            ORDER BY timestamp DESC LIMIT ?
        """,
            (limit,),
        )

        return [dict(row) for row in cursor.fetchall()]

    # Project Management (Feature 48)
    def create_project(
        self, name: str, description: str = "", owner_id: str = None
    ) -> str:
        """Create a new project"""
        project_id = str(uuid.uuid4())
        cursor = self.conn.cursor()

        cursor.execute(
            """
            INSERT INTO projects (id, name, description, owner_id)
            VALUES (?, ?, ?, ?)
        """,
            (project_id, name, description, owner_id),
        )

        self.conn.commit()
        return project_id

    def create_task(
        self,
        project_id: str,
        title: str,
        description: str = "",
        assigned_to: str = None,
        due_date: str = None,
    ) -> str:
        """Create a new task"""
        task_id = str(uuid.uuid4())
        cursor = self.conn.cursor()

        cursor.execute(
            """
            INSERT INTO tasks (id, project_id, title, description, assigned_to, due_date)
            VALUES (?, ?, ?, ?, ?, ?)
        """,
            (task_id, project_id, title, description, assigned_to, due_date),
        )

        self.conn.commit()
        return task_id

    def get_project_tasks(self, project_id: str) -> List[Dict]:
        """Get all tasks for a project"""
        cursor = self.conn.cursor()
        cursor.execute(
            """
            SELECT * FROM tasks WHERE project_id = ? ORDER BY created_at DESC
        """,
            (project_id,),
        )

        return [dict(row) for row in cursor.fetchall()]

    def update_task_status(self, task_id: str, status: str):
        """Update task status"""
        cursor = self.conn.cursor()

        if status == "completed":
            cursor.execute(
                """
                UPDATE tasks SET status = ?, completed_at = CURRENT_TIMESTAMP WHERE id = ?
            """,
                (status, task_id),
            )
        else:
            cursor.execute(
                """
                UPDATE tasks SET status = ?, completed_at = NULL WHERE id = ?
            """,
                (status, task_id),
            )

        self.conn.commit()

    # Cleanup methods
    def close(self):
        """Close database connection"""
        self.conn.close()

    def vacuum(self):
        """Optimize database"""
        self.conn.execute("VACUUM")


# Global database instance
db_manager = DatabaseManager()


if __name__ == "__main__":
    # Test database
    db = DatabaseManager()
    print("✓ Database initialized successfully")
    print(f"✓ Database path: {db.db_path}")

    # Test templates
    templates = db.get_templates()
    print(f"✓ Loaded {len(templates)} templates")

    db.close()
    print("✓ Database test complete")

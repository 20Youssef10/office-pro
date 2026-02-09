"""
Office Pro - Core Configuration and Feature Manager
Manages all 50 features with enable/disable capability
"""

import json
import os
from pathlib import Path
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, asdict
from datetime import datetime


@dataclass
class FeatureConfig:
    """Configuration for individual features"""

    name: str
    enabled: bool = True
    version: str = "1.0.0"
    settings: Dict[str, Any] = None

    def __post_init__(self):
        if self.settings is None:
            self.settings = {}


class FeatureManager:
    """Manages all 50 Office Pro features"""

    def __init__(self, config_path: str = None):
        self.config_path = config_path or os.path.expanduser(
            "~/.office_pro/features.json"
        )
        self.features: Dict[str, FeatureConfig] = {}
        self.load_config()
        self.initialize_features()

    def initialize_features(self):
        """Initialize all 50 features with default settings"""
        feature_definitions = {
            # Phase 1: Foundation (Features 1-10)
            "track_changes": FeatureConfig(
                "Track Changes & Comments",
                True,
                "1.0.0",
                {
                    "enabled_by_default": True,
                    "storage_limit_mb": 100,
                    "auto_save_interval": 300,
                },
            ),
            "version_history": FeatureConfig(
                "Version History",
                True,
                "1.0.0",
                {"max_versions": 50, "auto_save": True, "compression": True},
            ),
            "auto_save": FeatureConfig(
                "Auto-save", True, "1.0.0", {"interval_seconds": 300, "backup_count": 5}
            ),
            "data_validation": FeatureConfig(
                "Data Validation",
                True,
                "1.0.0",
                {"show_error_alerts": True, "allow_blank": True},
            ),
            "find_replace": FeatureConfig(
                "Find & Replace",
                True,
                "1.0.0",
                {"regex_support": True, "case_sensitive": False},
            ),
            "templates_gallery": FeatureConfig(
                "Templates Gallery",
                True,
                "1.0.0",
                {
                    "categories": [
                        "Business",
                        "Academic",
                        "Creative",
                        "Legal",
                        "Medical",
                    ],
                    "custom_templates_path": "",
                },
            ),
            "spell_check": FeatureConfig(
                "Spell Check",
                True,
                "1.0.0",
                {"languages": ["en_US", "en_GB"], "auto_check": True},
            ),
            "comments": FeatureConfig(
                "Comments System",
                True,
                "1.0.0",
                {"threaded": True, "notifications": True},
            ),
            "pivot_tables_basic": FeatureConfig(
                "Pivot Tables (Basic)",
                True,
                "1.0.0",
                {"max_rows": 10000, "max_cols": 100},
            ),
            "pdf_view": FeatureConfig(
                "PDF Viewer", True, "1.0.0", {"render_dpi": 150, "cache_size_mb": 50}
            ),
            # Phase 2: Productivity (Features 11-20)
            "grammar_checker": FeatureConfig(
                "Grammar & Style Checker",
                True,
                "1.0.0",
                {"enabled": True, "check_style": True, "check_readability": True},
            ),
            "mail_merge": FeatureConfig(
                "Mail Merge",
                True,
                "1.0.0",
                {"max_recipients": 10000, "email_support": True},
            ),
            "citation_manager": FeatureConfig(
                "Citation Manager",
                True,
                "1.0.0",
                {
                    "styles": ["APA", "MLA", "Chicago", "Harvard", "IEEE"],
                    "auto_format": True,
                },
            ),
            "equation_editor": FeatureConfig(
                "Equation Editor",
                True,
                "1.0.0",
                {"latex_support": True, "visual_editor": True},
            ),
            "data_cleaning": FeatureConfig(
                "Data Cleaning Tools",
                True,
                "1.0.0",
                {"auto_detect_errors": True, "suggestions": True},
            ),
            "pdf_forms": FeatureConfig(
                "PDF Form Creation",
                True,
                "1.0.0",
                {
                    "field_types": ["text", "checkbox", "radio", "dropdown"],
                    "validation": True,
                },
            ),
            "ocr": FeatureConfig(
                "PDF OCR",
                True,
                "1.0.0",
                {"engine": "tesseract", "languages": ["eng"], "auto_rotate": True},
            ),
            "screen_capture": FeatureConfig(
                "Screen Capture",
                True,
                "1.0.0",
                {
                    "modes": ["full", "window", "region", "scrolling"],
                    "annotations": True,
                },
            ),
            "qr_generator": FeatureConfig(
                "QR Code Generator",
                True,
                "1.0.0",
                {"types": ["text", "url", "contact", "wifi"], "customization": True},
            ),
            "unit_converter": FeatureConfig(
                "Unit Converter",
                True,
                "1.0.0",
                {"categories": ["length", "weight", "currency", "temperature", "data"]},
            ),
            # Phase 3: Collaboration (Features 21-30)
            "realtime_collab": FeatureConfig(
                "Real-time Collaboration",
                False,
                "1.0.0",
                {"max_users": 50, "cursor_tracking": True, "chat": True},
            ),
            "document_compare": FeatureConfig(
                "Document Compare",
                True,
                "1.0.0",
                {"side_by_side": True, "highlight_changes": True},
            ),
            "backup_sync": FeatureConfig(
                "Backup & Sync",
                True,
                "1.0.0",
                {
                    "providers": ["local", "google_drive", "dropbox"],
                    "auto_backup": True,
                },
            ),
            "project_management": FeatureConfig(
                "Project Management",
                True,
                "1.0.0",
                {"kanban": True, "gantt": True, "task_tracking": True},
            ),
            "collaborative_sheets": FeatureConfig(
                "Collaborative Spreadsheets",
                False,
                "1.0.0",
                {"realtime": True, "conflict_resolution": True},
            ),
            "advanced_charts": FeatureConfig(
                "3D Charts & Visualizations",
                True,
                "1.0.0",
                {
                    "types": ["3d_bar", "surface", "bubble", "radar", "gantt"],
                    "animations": True,
                },
            ),
            "database_integration": FeatureConfig(
                "Database Integration",
                True,
                "1.0.0",
                {"supported": ["sqlite", "mysql", "postgresql"], "live_refresh": False},
            ),
            "form_controls": FeatureConfig(
                "Form Controls & Dashboards",
                True,
                "1.0.0",
                {
                    "controls": ["button", "checkbox", "slider", "dropdown"],
                    "interactive": True,
                },
            ),
            "statistical_tools": FeatureConfig(
                "Statistical Analysis Tools",
                True,
                "1.0.0",
                {
                    "tests": ["regression", "anova", "t-test", "chi-square"],
                    "distributions": True,
                },
            ),
            "scenario_manager": FeatureConfig(
                "What-If Analysis",
                True,
                "1.0.0",
                {"max_scenarios": 10, "comparison": True},
            ),
            # Phase 4: AI & Intelligence (Features 31-40)
            "ai_writing": FeatureConfig(
                "AI Writing Assistant",
                False,
                "1.0.0",
                {
                    "enabled": False,
                    "api_key": "",
                    "tone_adjustment": True,
                    "summarization": True,
                },
            ),
            "voice_typing": FeatureConfig(
                "Voice Typing & Dictation",
                False,
                "1.0.0",
                {"enabled": False, "language": "en-US", "commands": True},
            ),
            "smart_autocorrect": FeatureConfig(
                "Smart Auto-Correct",
                True,
                "1.0.0",
                {"learn_patterns": True, "custom_shortcuts": True},
            ),
            "ai_document_intelligence": FeatureConfig(
                "AI Document Intelligence",
                False,
                "1.0.0",
                {
                    "enabled": False,
                    "summarization": True,
                    "entity_extraction": True,
                    "sentiment_analysis": True,
                    "classification": True,
                },
            ),
            "advanced_formulas": FeatureConfig(
                "Advanced Formula Engine",
                True,
                "1.0.0",
                {
                    "excel_compatible": True,
                    "custom_functions": True,
                    "total_functions": 500,
                },
            ),
            "goal_seek": FeatureConfig(
                "Goal Seek & Solver",
                True,
                "1.0.0",
                {"single_variable": True, "multi_variable": True, "constraints": True},
            ),
            "pivot_tables_advanced": FeatureConfig(
                "Pivot Tables (Advanced)",
                True,
                "1.0.0",
                {"calculated_fields": True, "grouping": True, "filtering": True},
            ),
            "web_import": FeatureConfig(
                "Import from Web & APIs",
                True,
                "1.0.0",
                {"rest_api": True, "json_parsing": True, "auto_refresh": True},
            ),
            "digital_signatures": FeatureConfig(
                "Digital Signatures",
                True,
                "1.0.0",
                {
                    "create_certificates": True,
                    "verify_signatures": True,
                    "timestamp": True,
                },
            ),
            "batch_processing": FeatureConfig(
                "Batch Processing & Automation",
                True,
                "1.0.0",
                {
                    "workflows": True,
                    "scheduler": True,
                    "operations": ["convert", "ocr", "sign", "watermark"],
                },
            ),
            # Phase 5: Advanced & Ecosystem (Features 41-50)
            "clipboard_manager": FeatureConfig(
                "Clipboard Manager",
                True,
                "1.0.0",
                {"history_size": 100, "searchable": True, "sync": False},
            ),
            "file_organizer": FeatureConfig(
                "File Organizer & Batch Renamer",
                True,
                "1.0.0",
                {"patterns": ["date", "sequence", "regex"], "preview": True},
            ),
            "macro_recording": FeatureConfig(
                "Macro Recording & Automation",
                True,
                "1.0.0",
                {"record_actions": True, "python_based": True, "scheduler": True},
            ),
            "plugin_system": FeatureConfig(
                "Plugin & Extension System",
                False,
                "1.0.0",
                {"enabled": False, "api_version": "1.0", "marketplace": False},
            ),
            "table_of_figures": FeatureConfig(
                "Table of Figures/Tables",
                True,
                "1.0.0",
                {"auto_update": True, "cross_references": True},
            ),
            "pdf_redaction": FeatureConfig(
                "PDF Redaction Tool",
                True,
                "1.0.0",
                {"permanent_removal": True, "verification": True},
            ),
            "pdf_merge_split": FeatureConfig(
                "PDF Merge, Split & Reorganize",
                True,
                "1.0.0",
                {"drag_drop": True, "batch_operations": True},
            ),
            "pdf_comparison": FeatureConfig(
                "PDF Comparison & Diff View",
                True,
                "1.0.0",
                {"visual_comparison": True, "text_comparison": True},
            ),
            "password_manager": FeatureConfig(
                "Password Manager Integration",
                False,
                "1.0.0",
                {"enabled": False, "generator": True, "autofill": True},
            ),
            "annotation_layers": FeatureConfig(
                "Annotation Layer Management",
                True,
                "1.0.0",
                {
                    "multiple_layers": True,
                    "show_hide": True,
                    "threaded_discussions": True,
                },
            ),
        }

        # Merge with existing config
        for key, config in feature_definitions.items():
            if key not in self.features:
                self.features[key] = config

    def load_config(self):
        """Load feature configuration from disk"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    for key, value in data.items():
                        self.features[key] = FeatureConfig(**value)
        except Exception as e:
            print(f"Error loading feature config: {e}")

    def save_config(self):
        """Save feature configuration to disk"""
        try:
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
            data = {k: asdict(v) for k, v in self.features.items()}
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            print(f"Error saving feature config: {e}")

    def is_enabled(self, feature_name: str) -> bool:
        """Check if a feature is enabled"""
        if feature_name in self.features:
            return self.features[feature_name].enabled
        return False

    def enable(self, feature_name: str):
        """Enable a feature"""
        if feature_name in self.features:
            self.features[feature_name].enabled = True
            self.save_config()

    def disable(self, feature_name: str):
        """Disable a feature"""
        if feature_name in self.features:
            self.features[feature_name].enabled = False
            self.save_config()

    def get_feature(self, feature_name: str) -> Optional[FeatureConfig]:
        """Get feature configuration"""
        return self.features.get(feature_name)

    def get_enabled_features(self) -> List[str]:
        """Get list of enabled features"""
        return [name for name, config in self.features.items() if config.enabled]

    def get_features_by_phase(self, phase: int) -> List[str]:
        """Get features by implementation phase"""
        phase_ranges = {
            1: list(self.features.keys())[:10],  # Foundation
            2: list(self.features.keys())[10:20],  # Productivity
            3: list(self.features.keys())[20:30],  # Collaboration
            4: list(self.features.keys())[30:40],  # AI/Intelligence
            5: list(self.features.keys())[40:50],  # Advanced
        }
        return phase_ranges.get(phase, [])

    def update_setting(self, feature_name: str, setting: str, value: Any):
        """Update a feature setting"""
        if feature_name in self.features:
            self.features[feature_name].settings[setting] = value
            self.save_config()


# Global feature manager instance
feature_manager = FeatureManager()


if __name__ == "__main__":
    # Test the feature manager
    fm = FeatureManager()
    print("Total features:", len(fm.features))
    print("Enabled features:", len(fm.get_enabled_features()))
    print("\nEnabled features:")
    for feature in fm.get_enabled_features():
        config = fm.get_feature(feature)
        print(f"  ✓ {config.name}")

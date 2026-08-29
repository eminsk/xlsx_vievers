"""
Application configuration, constants, and styling defaults for Excel Viewer Pro.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any
from openpyxl.utils import get_column_letter


# =============================================================================
# Application Configuration
# =============================================================================

class Config:
    """Application configuration constants."""

    # Window
    WINDOW_SIZE = "1440x880"
    MIN_SIZE = (1040, 660)
    THEME = "litera"  # Modern light theme
    APP_TITLE = "Excel Viewer Pro"
    ICON_PATH = "app_icon.ico"
    ICON_PNG_PATH = "app_icon.png"

    # Excel Theme Colors
    EXCEL_GREEN = "#107c41"
    EXCEL_DARK_GREEN = "#0b5a2f"
    EXCEL_LIGHT_GREEN = "#e7f4ec"
    EXCEL_ACCENT = "#217346"
    EXCEL_HEADER_TEXT = "#ffffff"

    # Grid & UI Colors
    HEADER_BG = "#f3f3f3"
    HEADER_FG = "#262626"
    HEADER_SELECTED_BG = "#c8e5d0"
    CELL_BG = "#ffffff"
    CELL_FG = "#1a1a1a"
    ROW_ALT_BG = "#f9fbfd"
    GRID_COLOR = "#d9d9d9"
    SELECTION_BG = "#d7eafd"
    SELECTION_BORDER = "#107c41"
    ACTIVE_CELL_BORDER = "#107c41"
    RANGE_BORDER = "#107c41"
    RANGE_FILL = "#e8f0fe"
    BORDER_WIDTH = 2
    EDIT_BG = "#ffffff"

    # Fill Handle
    FILL_HANDLE_SIZE = 6
    FILL_HANDLE_COLOR = "#107c41"

    # Dimensions
    ROW_NUM_WIDTH = 55
    DEFAULT_COL_WIDTH = 110
    MIN_COL_WIDTH = 30
    MAX_COL_WIDTH = 500
    DEFAULT_ROW_HEIGHT = 24
    MIN_ROW_HEIGHT = 18

    # Performance
    CHUNK_SIZE = 500
    UNDO_LIMIT = 100

    # Fonts
    FONT_FAMILY = "Segoe UI"
    FONT_SIZE = 10
    FONT = ("Segoe UI", 10)
    FONT_BOLD = ("Segoe UI", 10, "bold")
    FONT_ITALIC = ("Segoe UI", 10, "italic")
    FONT_SMALL = ("Segoe UI", 8)
    FONT_HEADER = ("Segoe UI", 10, "bold")
    FONT_MONO = ("Consolas", 10)

    # Number Format Categories
    NUMBER_FORMATS = {
        "General": "",
        "Number (0)": "0",
        "Number (0.00)": "0.00",
        "Currency ($)": "$#,##0.00",
        "Currency (₽)": "#,##0.00 ₽",
        "Currency (€)": "€#,##0.00",
        "Accounting": "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
        "Short Date": "YYYY-MM-DD",
        "Date (DD.MM.YYYY)": "DD.MM.YYYY",
        "Long Date": "DD MMMM YYYY",
        "Time": "HH:MM:SS",
        "Percentage (0%)": "0%",
        "Percentage (0.00%)": "0.00%",
        "Fraction": "# ?/?",
        "Scientific": "0.00E+00",
        "Text": "@"
    }

    # Available UI Themes
    AVAILABLE_THEMES = [
        "litera", "flatly", "cosmo", "minty", "pulse", "sandstone",
        "journal", "united", "yeti", "darkly", "cyborg", "superhero"
    ]

    # Standard Chart Palettes
    CHART_PALETTES = {
        "Excel Classic": ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"],
        "Modern Teal": ["#008080", "#20B2AA", "#48D1CC", "#00CED1", "#5F9EA0", "#4682B4"],
        "Vibrant": ["#E6194B", "#3CB44B", "#FFE119", "#4363D8", "#F58231", "#911EB4"],
        "Pastel": ["#AEC7E8", "#FFBB78", "#98DF8A", "#FF9896", "#C5B0D5", "#C49C94"],
        "Emerald Forest": ["#0F5132", "#198754", "#20C997", "#0DCAF0", "#3DD5F3", "#6EDFF6"],
        "Warm Sunset": ["#C0392B", "#E67E22", "#F39C12", "#F1C40F", "#D35400", "#E74C3C"],
        "Indigo Night": ["#3B4252", "#4C566A", "#5E81AC", "#81A1C1", "#88C0D0", "#8FBCBB"],
        "Monochrome": ["#252525", "#525252", "#737373", "#969696", "#BDBDBD", "#D9D9D9"]
    }

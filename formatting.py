"""
Number formatting and conditional formatting engine for Excel Viewer Pro.
"""

from __future__ import annotations

import re
from datetime import datetime, date, time
from typing import Any
from models import CellStyle


# =============================================================================
# Number Formatter
# =============================================================================

class NumberFormatter:
    """Formats numeric and date values according to Excel number formats."""

    @staticmethod
    def format_value(value: Any, fmt: str | None) -> str:
        """Format a value with an Excel format string."""
        if value is None or value == "":
            return ""

        if not fmt or fmt.upper() in ("GENERAL", ""):
            if isinstance(value, float):
                if value == int(value):
                    return str(int(value))
                return f"{round(value, 8):g}"
            return str(value)

        # Text format
        if fmt == "@":
            return str(value)

        # Date/Time formatting for datetime objects
        if isinstance(value, (datetime, date, time)):
            return NumberFormatter._format_datetime(value, fmt)

        # Check if fmt is a date format and value is a serial number
        is_date_fmt = any(token in fmt.upper() for token in ("YYYY", "YY", "MMMM", "MMM", "DD", "HH:MM", "SS"))
        if is_date_fmt:
            try:
                serial_num = float(str(value).replace(",", ".").replace(" ", ""))
                if 0 < serial_num < 2958465:  # Valid Excel date range (up to year 9999)
                    # Excel 1900 epoch adjustment
                    from datetime import timedelta
                    excel_epoch = datetime(1899, 12, 30)
                    dt_val = excel_epoch + timedelta(days=serial_num)
                    return NumberFormatter._format_datetime(dt_val, fmt)
            except (ValueError, TypeError, OverflowError):
                pass

        # Try converting to numeric
        try:
            num = float(str(value).replace(",", ".").replace(" ", "").replace("₽", "").replace("$", "").replace("€", "").replace("%", ""))
        except (ValueError, TypeError):
            # Not a number - return original string
            return str(value)

        # Percentage formats
        if "%" in fmt:
            pct_val = num * 100 if "%" not in str(value) else num
            m = re.search(r"\.(0+)", fmt)
            if m:
                decimals = len(m.group(1))
                return f"{pct_val:,.{decimals}f}%"
            return f"{round(pct_val):,.0f}%"

        # Accounting formats with parentheses: _($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)
        if "(*" in fmt or "_(" in fmt:
            m = re.search(r"\.(0+)", fmt)
            decimals = len(m.group(1)) if m else 2
            if num < 0:
                return f"(${abs(num):,.{decimals}f})"
            elif num == 0 and "-" in fmt:
                return "$ -"
            return f"${num:,.{decimals}f}"

        # Currency formats
        if "₽" in fmt or "RUB" in fmt.upper():
            sign = "-" if num < 0 else ""
            m = re.search(r"\.(0+)", fmt)
            if m:
                decimals = len(m.group(1))
                return f"{sign}{abs(num):,.{decimals}f} ₽".replace(",", " ")
            return f"{sign}{abs(round(num)):,} ₽".replace(",", " ")

        if "$" in fmt or "USD" in fmt.upper():
            sign = "-" if num < 0 else ""
            m = re.search(r"\.(0+)", fmt)
            if m:
                decimals = len(m.group(1))
                return f"{sign}${abs(num):,.{decimals}f}"
            return f"{sign}${abs(round(num)):,}"

        if "€" in fmt or "EUR" in fmt.upper():
            sign = "-" if num < 0 else ""
            m = re.search(r"\.(0+)", fmt)
            if m:
                decimals = len(m.group(1))
                return f"{sign}€{abs(num):,.{decimals}f}"
            return f"{sign}€{abs(round(num)):,}"

        # Scientific notation
        if "E+" in fmt.upper() or "E-" in fmt.upper():
            return f"{num:.2E}"

        # Standard decimal formats
        if "#,##0" in fmt or "0,000" in fmt or "#,##" in fmt:
            m = re.search(r"\.(0+)", fmt)
            if m:
                decimals = len(m.group(1))
                return f"{num:,.{decimals}f}"
            return f"{round(num):,}"

        if "0.0" in fmt:
            m = re.search(r"0\.(0+)", fmt)
            decimals = len(m.group(1)) if m else 2
            return f"{num:.{decimals}f}"

        if fmt == "0":
            return f"{round(num)}"

        return str(value)

    @staticmethod
    def _format_datetime(dt_val: datetime | date | time, fmt: str) -> str:
        # Standardize format tokens
        py_fmt = fmt.replace("YYYY", "%Y").replace("yyyy", "%Y")
        py_fmt = py_fmt.replace("YY", "%y").replace("yy", "%y")
        py_fmt = py_fmt.replace("MMMM", "%B").replace("mmmm", "%B")
        py_fmt = py_fmt.replace("MMM", "%b").replace("mmm", "%b")
        py_fmt = py_fmt.replace("MM", "%m").replace("mm", "%m")
        py_fmt = py_fmt.replace("DD", "%d").replace("dd", "%d")
        py_fmt = py_fmt.replace("HH", "%H").replace("hh", "%H")
        py_fmt = py_fmt.replace("SS", "%S").replace("ss", "%S")
        try:
            return dt_val.strftime(py_fmt)
        except Exception:
            return str(dt_val)


# =============================================================================
# Conditional Formatting Engine
# =============================================================================

class ConditionalFormattingEngine:
    """Evaluates conditional formatting rules and generates style overrides."""

    @staticmethod
    def evaluate_rule(val: Any, rule: dict[str, Any], all_values: list[float] | None = None) -> tuple[str | None, str | None]:
        """
        Evaluate a single rule against a cell value.
        Returns (bg_color, fg_color) if matched, else (None, None).
        """
        rule_type = rule.get("type")
        bg_color = rule.get("bg_color")
        fg_color = rule.get("fg_color")

        try:
            num = float(str(val).replace(",", ".").replace(" ", ""))
        except (ValueError, TypeError):
            num = None

        if rule_type == "greater_than" and num is not None:
            target = float(rule.get("value", 0))
            if num > target:
                return bg_color, fg_color

        elif rule_type == "less_than" and num is not None:
            target = float(rule.get("value", 0))
            if num < target:
                return bg_color, fg_color

        elif rule_type == "between" and num is not None:
            v1 = float(rule.get("value1", 0))
            v2 = float(rule.get("value2", 0))
            if min(v1, v2) <= num <= max(v1, v2):
                return bg_color, fg_color

        elif rule_type == "equal_to":
            target = str(rule.get("value", ""))
            if str(val).lower() == target.lower():
                return bg_color, fg_color

        elif rule_type == "text_contains":
            target = str(rule.get("value", "")).lower()
            if target in str(val if val is not None else "").lower():
                return bg_color, fg_color

        elif rule_type == "color_scale" and num is not None and all_values:
            valid_nums = [v for v in all_values if v is not None]
            if len(valid_nums) >= 2:
                min_v = min(valid_nums)
                max_v = max(valid_nums)
                if max_v > min_v:
                    ratio = (num - min_v) / (max_v - min_v)
                    # 3-color scale (Green -> Yellow -> Red or Red -> Yellow -> Green)
                    scale_type = rule.get("scale", "green_yellow_red")
                    if scale_type == "green_yellow_red":
                        # high = green, mid = yellow, low = red
                        hex_bg = ConditionalFormattingEngine._interpolate_color(ratio, "#f8696b", "#ffeb84", "#63be7b")
                        return hex_bg, "#000000"
                    elif scale_type == "red_yellow_green":
                        hex_bg = ConditionalFormattingEngine._interpolate_color(ratio, "#63be7b", "#ffeb84", "#f8696b")
                        return hex_bg, "#000000"

        return None, None

    @staticmethod
    def _interpolate_color(ratio: float, c_low: str, c_mid: str, c_high: str) -> str:
        """Interpolate hex colors at ratio (0.0 to 1.0)."""
        def hex_to_rgb(h: str) -> tuple[int, int, int]:
            h = h.lstrip("#")
            return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

        def rgb_to_hex(rgb: tuple[int, int, int]) -> str:
            return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"

        ratio = max(0.0, min(1.0, ratio))
        if ratio <= 0.5:
            r = ratio * 2.0
            r1, g1, b1 = hex_to_rgb(c_low)
            r2, g2, b2 = hex_to_rgb(c_mid)
            rgb = (int(r1 + (r2 - r1) * r), int(g1 + (g2 - g1) * r), int(b1 + (b2 - b1) * r))
        else:
            r = (ratio - 0.5) * 2.0
            r1, g1, b1 = hex_to_rgb(c_mid)
            r2, g2, b2 = hex_to_rgb(c_high)
            rgb = (int(r1 + (r2 - r1) * r), int(g1 + (g2 - g1) * r), int(b1 + (b2 - b1) * r))

        return rgb_to_hex(rgb)

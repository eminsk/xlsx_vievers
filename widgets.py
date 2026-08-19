"""
Professional UI widgets and Ribbon components for Excel Viewer Pro.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import Menu, colorchooser, simpledialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from typing import Callable, Any
from functools import partial

from config import Config
from models import CellPosition, CellRange
from formulas import FUNCTION_METADATA


# =============================================================================
# Ribbon Bar (Excel-Style Tabbed Toolbar)
# =============================================================================

class ExcelRibbon(ttk.Frame):
    """Excel-like Tabbed Ribbon Toolbar."""

    def __init__(self, parent: tk.Widget, callbacks: dict[str, Callable]) -> None:
        super().__init__(parent)
        self._cb = callbacks
        self._build_ui()

    def _build_ui(self) -> None:
        # File & Ribbon tabs notebook
        self._notebook = ttk.Notebook(self)
        self._notebook.pack(fill=tk.BOTH, expand=True)

        self._create_home_tab()
        self._create_insert_tab()
        self._create_data_tab()
        self._create_formulas_tab()
        self._create_view_tab()

    # -------------------------------------------------------------------------
    # Home Tab
    # -------------------------------------------------------------------------
    def _create_home_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=(4, 2))
        self._notebook.add(tab, text="  Home  ")

        # 1. Clipboard Group
        clip_grp = ttk.LabelFrame(tab, text="Clipboard", padding=2)
        clip_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        paste_btn = ttk.Menubutton(clip_grp, text="📋 Paste", bootstyle="primary")
        paste_menu = Menu(paste_btn, tearoff=0)
        paste_menu.add_command(label="Paste All (Ctrl+V)", command=self._cb.get("paste"))
        paste_menu.add_command(label="Paste Values Only", command=self._cb.get("paste_values"))
        paste_menu.add_command(label="Paste Formulas Only", command=self._cb.get("paste_formulas"))
        paste_menu.add_command(label="Paste Transposed", command=self._cb.get("paste_transpose"))
        paste_btn["menu"] = paste_menu
        paste_btn.pack(side=tk.LEFT, padx=1)

        clip_sub = ttk.Frame(clip_grp)
        clip_sub.pack(side=tk.LEFT, padx=1)
        ttk.Button(clip_sub, text="✂ Cut", width=6, bootstyle="outline", command=self._cb.get("cut")).pack(anchor=tk.W, pady=1)
        ttk.Button(clip_sub, text="📄 Copy", width=6, bootstyle="outline", command=self._cb.get("copy")).pack(anchor=tk.W, pady=1)
        ttk.Button(clip_sub, text="🖌 Format", width=6, bootstyle="outline", command=self._cb.get("format_painter")).pack(anchor=tk.W, pady=1)

        # 2. Font Group
        font_grp = ttk.LabelFrame(tab, text="Font", padding=2)
        font_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        font_top = ttk.Frame(font_grp)
        font_top.pack(fill=tk.X, pady=1)

        self._font_family = ttk.Combobox(
            font_top,
            values=["Calibri", "Segoe UI", "Arial", "Times New Roman", "Consolas", "Verdana", "Tahoma"],
            width=12, state="readonly"
        )
        self._font_family.set("Segoe UI")
        self._font_family.pack(side=tk.LEFT, padx=1)
        self._font_family.bind("<<ComboboxSelected>>", lambda e: self._cb.get("set_font_family")(self._font_family.get()))

        self._font_size = ttk.Combobox(
            font_top,
            values=["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36"],
            width=3, state="readonly"
        )
        self._font_size.set("10")
        self._font_size.pack(side=tk.LEFT, padx=1)
        self._font_size.bind("<<ComboboxSelected>>", lambda e: self._cb.get("set_font_size")(int(self._font_size.get())))

        ttk.Button(font_top, text="A+", width=3, bootstyle="secondary-outline", command=self._cb.get("increase_font_size")).pack(side=tk.LEFT, padx=1)
        ttk.Button(font_top, text="A-", width=3, bootstyle="secondary-outline", command=self._cb.get("decrease_font_size")).pack(side=tk.LEFT, padx=1)

        font_bot = ttk.Frame(font_grp)
        font_bot.pack(fill=tk.X, pady=1)

        ttk.Button(font_bot, text="𝐁", width=2, bootstyle="secondary-outline", command=self._cb.get("toggle_bold")).pack(side=tk.LEFT, padx=1)
        ttk.Button(font_bot, text="𝐼", width=2, bootstyle="secondary-outline", command=self._cb.get("toggle_italic")).pack(side=tk.LEFT, padx=1)
        ttk.Button(font_bot, text="𝐔", width=2, bootstyle="secondary-outline", command=self._cb.get("toggle_underline")).pack(side=tk.LEFT, padx=1)
        ttk.Button(font_bot, text="S̶", width=2, bootstyle="secondary-outline", command=self._cb.get("toggle_strikethrough")).pack(side=tk.LEFT, padx=1)

        # Border menu
        border_btn = ttk.Menubutton(font_bot, text="田", width=3, bootstyle="secondary-outline")
        border_menu = Menu(border_btn, tearoff=0)
        border_menu.add_command(label="All Borders", command=lambda: self._cb.get("set_borders")("all"))
        border_menu.add_command(label="Outside Borders", command=lambda: self._cb.get("set_borders")("outside"))
        border_menu.add_command(label="Thick Box Border", command=lambda: self._cb.get("set_borders")("thick"))
        border_menu.add_command(label="Bottom Border", command=lambda: self._cb.get("set_borders")("bottom"))
        border_menu.add_command(label="Top and Bottom Border", command=lambda: self._cb.get("set_borders")("top_bottom"))
        border_menu.add_separator()
        border_menu.add_command(label="No Border", command=lambda: self._cb.get("set_borders")("none"))
        border_btn["menu"] = border_menu
        border_btn.pack(side=tk.LEFT, padx=1)

        # Colors
        self._bg_btn = tk.Button(font_bot, text="🎨", width=2, bg="#FFFF00", command=self._cb.get("set_bg_color"))
        self._bg_btn.pack(side=tk.LEFT, padx=1)
        self._fg_btn = tk.Button(font_bot, text="A", width=2, fg="#FF0000", font=("Segoe UI", 9, "bold"), command=self._cb.get("set_fg_color"))
        self._fg_btn.pack(side=tk.LEFT, padx=1)

        # 3. Alignment Group
        align_grp = ttk.LabelFrame(tab, text="Alignment", padding=2)
        align_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        align_top = ttk.Frame(align_grp)
        align_top.pack(fill=tk.X, pady=1)
        ttk.Button(align_top, text="⬆ Top", width=6, bootstyle="outline", command=lambda: self._cb.get("set_valign")("top")).pack(side=tk.LEFT, padx=1)
        ttk.Button(align_top, text="⬍ Mid", width=6, bootstyle="outline", command=lambda: self._cb.get("set_valign")("center")).pack(side=tk.LEFT, padx=1)
        ttk.Button(align_top, text="⬇ Bot", width=6, bootstyle="outline", command=lambda: self._cb.get("set_valign")("bottom")).pack(side=tk.LEFT, padx=1)

        align_bot = ttk.Frame(align_grp)
        align_bot.pack(fill=tk.X, pady=1)
        ttk.Button(align_bot, text="⬅ Left", width=6, bootstyle="outline", command=lambda: self._cb.get("set_halign")("left")).pack(side=tk.LEFT, padx=1)
        ttk.Button(align_bot, text="⬌ Center", width=6, bootstyle="outline", command=lambda: self._cb.get("set_halign")("center")).pack(side=tk.LEFT, padx=1)
        ttk.Button(align_bot, text="➡ Right", width=6, bootstyle="outline", command=lambda: self._cb.get("set_halign")("right")).pack(side=tk.LEFT, padx=1)

        align_side = ttk.Frame(align_grp)
        align_side.pack(side=tk.LEFT, fill=tk.Y, padx=2)
        ttk.Button(align_side, text="Wrap Text", width=9, bootstyle="outline", command=self._cb.get("toggle_wrap_text")).pack(pady=1)
        ttk.Button(align_side, text="Merge & Ctr", width=9, bootstyle="outline", command=self._cb.get("merge_cells")).pack(pady=1)

        # 4. Number Group
        num_grp = ttk.LabelFrame(tab, text="Number", padding=2)
        num_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._num_fmt = ttk.Combobox(
            num_grp,
            values=list(Config.NUMBER_FORMATS.keys()),
            width=14, state="readonly"
        )
        self._num_fmt.set("General")
        self._num_fmt.pack(fill=tk.X, pady=1)
        self._num_fmt.bind("<<ComboboxSelected>>", lambda e: self._cb.get("set_number_format")(Config.NUMBER_FORMATS.get(self._num_fmt.get(), "")))

        num_btns = ttk.Frame(num_grp)
        num_btns.pack(fill=tk.X, pady=1)
        ttk.Button(num_btns, text="₽", width=2, bootstyle="outline", command=lambda: self._cb.get("set_number_format")("#,##0.00 ₽")).pack(side=tk.LEFT, padx=1)
        ttk.Button(num_btns, text="$", width=2, bootstyle="outline", command=lambda: self._cb.get("set_number_format")("$#,##0.00")).pack(side=tk.LEFT, padx=1)
        ttk.Button(num_btns, text="%", width=2, bootstyle="outline", command=lambda: self._cb.get("set_number_format")("0.00%")).pack(side=tk.LEFT, padx=1)
        ttk.Button(num_btns, text=",", width=2, bootstyle="outline", command=lambda: self._cb.get("set_number_format")("#,##0.00")).pack(side=tk.LEFT, padx=1)
        ttk.Button(num_btns, text=".0→", width=3, bootstyle="outline", command=self._cb.get("increase_decimals")).pack(side=tk.LEFT, padx=1)
        ttk.Button(num_btns, text="←.0", width=3, bootstyle="outline", command=self._cb.get("decrease_decimals")).pack(side=tk.LEFT, padx=1)

        # 5. Styles (Conditional Formatting)
        style_grp = ttk.LabelFrame(tab, text="Styles", padding=2)
        style_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        cf_btn = ttk.Menubutton(style_grp, text="Conditional\nFormatting", bootstyle="info-outline")
        cf_menu = Menu(cf_btn, tearoff=0)
        cf_menu.add_command(label="Highlight: Greater Than...", command=lambda: self._cb.get("add_cf_rule")("greater_than"))
        cf_menu.add_command(label="Highlight: Less Than...", command=lambda: self._cb.get("add_cf_rule")("less_than"))
        cf_menu.add_command(label="Highlight: Between...", command=lambda: self._cb.get("add_cf_rule")("between"))
        cf_menu.add_command(label="Highlight: Equal To...", command=lambda: self._cb.get("add_cf_rule")("equal_to"))
        cf_menu.add_command(label="Highlight: Text that Contains...", command=lambda: self._cb.get("add_cf_rule")("text_contains"))
        cf_menu.add_separator()
        cf_menu.add_command(label="Color Scale: Green - Yellow - Red", command=lambda: self._cb.get("add_cf_rule")("color_scale_gyr"))
        cf_menu.add_command(label="Color Scale: Red - Yellow - Green", command=lambda: self._cb.get("add_cf_rule")("color_scale_ryg"))
        cf_menu.add_separator()
        cf_menu.add_command(label="Clear Rules from Selected Cells", command=self._cb.get("clear_cf_rules"))
        cf_btn["menu"] = cf_menu
        cf_btn.pack(fill=tk.BOTH, expand=True)

        # 6. Cells Group (Insert, Delete, Format)
        cells_grp = ttk.LabelFrame(tab, text="Cells", padding=2)
        cells_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        ins_btn = ttk.Menubutton(cells_grp, text="➕ Insert", width=8, bootstyle="outline")
        ins_menu = Menu(ins_btn, tearoff=0)
        ins_menu.add_command(label="Insert Row Above", command=self._cb.get("insert_row_above"))
        ins_menu.add_command(label="Insert Row Below", command=self._cb.get("insert_row_below"))
        ins_menu.add_command(label="Insert Column Left", command=self._cb.get("insert_col_left"))
        ins_menu.add_command(label="Insert Column Right", command=self._cb.get("insert_col_right"))
        ins_btn["menu"] = ins_menu
        ins_btn.pack(pady=1)

        del_btn = ttk.Menubutton(cells_grp, text="➖ Delete", width=8, bootstyle="outline")
        del_menu = Menu(del_btn, tearoff=0)
        del_menu.add_command(label="Delete Row", command=self._cb.get("delete_row"))
        del_menu.add_command(label="Delete Column", command=self._cb.get("delete_column"))
        del_btn["menu"] = del_menu
        del_btn.pack(pady=1)

        # 7. Editing Group (AutoSum, Sort, Filter, Find, Clear)
        edit_grp = ttk.LabelFrame(tab, text="Editing", padding=2)
        edit_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        autosum_btn = ttk.Menubutton(edit_grp, text="Σ AutoSum", width=10, bootstyle="primary-outline")
        autosum_menu = Menu(autosum_btn, tearoff=0)
        autosum_menu.add_command(label="Sum (=SUM)", command=lambda: self._cb.get("apply_autosum")("SUM"))
        autosum_menu.add_command(label="Average (=AVERAGE)", command=lambda: self._cb.get("apply_autosum")("AVERAGE"))
        autosum_menu.add_command(label="Count Numbers (=COUNT)", command=lambda: self._cb.get("apply_autosum")("COUNT"))
        autosum_menu.add_command(label="Max (=MAX)", command=lambda: self._cb.get("apply_autosum")("MAX"))
        autosum_menu.add_command(label="Min (=MIN)", command=lambda: self._cb.get("apply_autosum")("MIN"))
        autosum_btn["menu"] = autosum_menu
        autosum_btn.pack(pady=1)

        sort_btn = ttk.Menubutton(edit_grp, text="Sort & Filter", width=10, bootstyle="outline")
        sort_menu = Menu(sort_btn, tearoff=0)
        sort_menu.add_command(label="Sort A to Z", command=lambda: self._cb.get("quick_sort")(False))
        sort_menu.add_command(label="Sort Z to A", command=lambda: self._cb.get("quick_sort")(True))
        sort_menu.add_command(label="Custom Sort...", command=self._cb.get("custom_sort_dialog"))
        sort_menu.add_separator()
        sort_menu.add_command(label="Toggle AutoFilter", command=self._cb.get("toggle_filter"))
        sort_menu.add_command(label="Clear Filters", command=self._cb.get("clear_filters"))
        sort_btn["menu"] = sort_menu
        sort_btn.pack(pady=1)

        find_btn = ttk.Menubutton(edit_grp, text="Find & Select", width=10, bootstyle="outline")
        find_menu = Menu(find_btn, tearoff=0)
        find_menu.add_command(label="Find... (Ctrl+F)", command=self._cb.get("find_dialog"))
        find_menu.add_command(label="Replace... (Ctrl+H)", command=self._cb.get("replace_dialog"))
        find_menu.add_command(label="Go to Cell... (Ctrl+G)", command=self._cb.get("goto_dialog"))
        find_btn["menu"] = find_menu
        find_btn.pack(pady=1)

    # -------------------------------------------------------------------------
    # Insert Tab
    # -------------------------------------------------------------------------
    def _create_insert_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=(4, 2))
        self._notebook.add(tab, text="  Insert  ")

        # Charts
        chart_grp = ttk.LabelFrame(tab, text="Charts", padding=4)
        chart_grp.pack(side=tk.LEFT, fill=tk.Y, padx=4)

        ttk.Button(chart_grp, text="📊 Chart Wizard", bootstyle="primary", command=self._cb.get("show_chart_wizard")).pack(side=tk.LEFT, padx=3)
        ttk.Button(chart_grp, text="📈 Line Chart", bootstyle="outline", command=lambda: self._cb.get("quick_chart")("Line")).pack(side=tk.LEFT, padx=2)
        ttk.Button(chart_grp, text="📊 Column Chart", bootstyle="outline", command=lambda: self._cb.get("quick_chart")("Clustered Column")).pack(side=tk.LEFT, padx=2)
        ttk.Button(chart_grp, text="🥧 Pie Chart", bootstyle="outline", command=lambda: self._cb.get("quick_chart")("Pie")).pack(side=tk.LEFT, padx=2)

        # Comments & Annotations
        comm_grp = ttk.LabelFrame(tab, text="Comments & Links", padding=4)
        comm_grp.pack(side=tk.LEFT, fill=tk.Y, padx=4)
        ttk.Button(comm_grp, text="💬 Add / Edit Comment", bootstyle="outline", command=self._cb.get("edit_comment")).pack(side=tk.LEFT, padx=3)

        # Functions
        fn_grp = ttk.LabelFrame(tab, text="Function", padding=4)
        fn_grp.pack(side=tk.LEFT, fill=tk.Y, padx=4)
        ttk.Button(fn_grp, text="fx Insert Function", bootstyle="success", command=self._cb.get("show_fx_wizard")).pack(side=tk.LEFT, padx=3)

    # -------------------------------------------------------------------------
    # Data Tab
    # -------------------------------------------------------------------------
    def _create_data_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=(4, 2))
        self._notebook.add(tab, text="  Data  ")

        # Sort & Filter
        sf_grp = ttk.LabelFrame(tab, text="Sort & Filter", padding=4)
        sf_grp.pack(side=tk.LEFT, fill=tk.Y, padx=4)
        ttk.Button(sf_grp, text="⬆ Sort A-Z", bootstyle="outline", command=lambda: self._cb.get("quick_sort")(False)).pack(side=tk.LEFT, padx=2)
        ttk.Button(sf_grp, text="⬇ Sort Z-A", bootstyle="outline", command=lambda: self._cb.get("quick_sort")(True)).pack(side=tk.LEFT, padx=2)
        ttk.Button(sf_grp, text="⚡ Custom Sort...", bootstyle="outline", command=self._cb.get("custom_sort_dialog")).pack(side=tk.LEFT, padx=2)
        ttk.Button(sf_grp, text="🔍 AutoFilter", bootstyle="primary-outline", command=self._cb.get("toggle_filter")).pack(side=tk.LEFT, padx=2)
        ttk.Button(sf_grp, text="❌ Clear Filters", bootstyle="secondary-outline", command=self._cb.get("clear_filters")).pack(side=tk.LEFT, padx=2)

        # Data Tools
        dt_grp = ttk.LabelFrame(tab, text="Data Tools", padding=4)
        dt_grp.pack(side=tk.LEFT, fill=tk.Y, padx=4)
        ttk.Button(dt_grp, text="✂ Text to Columns...", bootstyle="outline", command=self._cb.get("text_to_columns_dialog")).pack(side=tk.LEFT, padx=2)
        ttk.Button(dt_grp, text="🗑 Remove Duplicates...", bootstyle="outline", command=self._cb.get("remove_duplicates_dialog")).pack(side=tk.LEFT, padx=2)
        ttk.Button(dt_grp, text="🎯 Goal Seek...", bootstyle="outline", command=self._cb.get("goal_seek_dialog")).pack(side=tk.LEFT, padx=2)

    # -------------------------------------------------------------------------
    # Formulas Tab
    # -------------------------------------------------------------------------
    def _create_formulas_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=(4, 2))
        self._notebook.add(tab, text="  Formulas  ")

        # Function Library
        lib_grp = ttk.LabelFrame(tab, text="Function Library", padding=4)
        lib_grp.pack(side=tk.LEFT, fill=tk.Y, padx=4)
        ttk.Button(lib_grp, text="fx Insert Function", bootstyle="success", command=self._cb.get("show_fx_wizard")).pack(side=tk.LEFT, padx=3)
        ttk.Button(lib_grp, text="Σ AutoSum", bootstyle="outline", command=lambda: self._cb.get("apply_autosum")("SUM")).pack(side=tk.LEFT, padx=2)

        for cat in ["Financial", "Logical", "Text", "Date & Time", "Lookup & Reference", "Math & Trig", "Statistical"]:
            btn = ttk.Menubutton(lib_grp, text=cat, bootstyle="outline")
            m = Menu(btn, tearoff=0)
            for name, meta in sorted(FUNCTION_METADATA.items()):
                if meta.get("cat") == cat:
                    m.add_command(label=f"{name} — {meta.get('desc')[:35]}...", command=partial(self._cb.get("insert_formula_text"), f"={name}("))
            btn["menu"] = m
            btn.pack(side=tk.LEFT, padx=1)

        # Calculation
        calc_grp = ttk.LabelFrame(tab, text="Calculation", padding=4)
        calc_grp.pack(side=tk.LEFT, fill=tk.Y, padx=4)
        ttk.Button(calc_grp, text="🔄 Calculate Sheet (F9)", bootstyle="primary", command=self._cb.get("recalculate_sheet")).pack(side=tk.LEFT, padx=2)

    # -------------------------------------------------------------------------
    # View Tab
    # -------------------------------------------------------------------------
    def _create_view_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=(4, 2))
        self._notebook.add(tab, text="  View  ")

        # Show / Hide
        show_grp = ttk.LabelFrame(tab, text="Show / Hide", padding=4)
        show_grp.pack(side=tk.LEFT, fill=tk.Y, padx=4)
        ttk.Button(show_grp, text="Toggle Formula Bar", bootstyle="outline", command=self._cb.get("toggle_formula_bar")).pack(side=tk.LEFT, padx=2)
        ttk.Button(show_grp, text="Toggle Status Bar", bootstyle="outline", command=self._cb.get("toggle_status_bar")).pack(side=tk.LEFT, padx=2)

        # Window & Freeze
        win_grp = ttk.LabelFrame(tab, text="Window", padding=4)
        win_grp.pack(side=tk.LEFT, fill=tk.Y, padx=4)
        freeze_btn = ttk.Menubutton(win_grp, text="❄ Freeze Panes", bootstyle="outline")
        frz_menu = Menu(freeze_btn, tearoff=0)
        frz_menu.add_command(label="Freeze Top Row", command=self._cb.get("freeze_top_row"))
        frz_menu.add_command(label="Freeze First Column", command=self._cb.get("freeze_first_col"))
        frz_menu.add_command(label="Unfreeze All", command=self._cb.get("unfreeze_all"))
        freeze_btn["menu"] = frz_menu
        freeze_btn.pack(side=tk.LEFT, padx=2)

        ttk.Button(win_grp, text="Auto-fit Columns", bootstyle="outline", command=self._cb.get("autofit_all_cols")).pack(side=tk.LEFT, padx=2)

        # Theme Selector
        theme_grp = ttk.LabelFrame(tab, text="Theme", padding=4)
        theme_grp.pack(side=tk.LEFT, fill=tk.Y, padx=4)
        theme_cb = ttk.Combobox(theme_grp, values=Config.AVAILABLE_THEMES, width=12, state="readonly")
        theme_cb.set(Config.THEME)
        theme_cb.pack(side=tk.LEFT, padx=4)
        theme_cb.bind("<<ComboboxSelected>>", lambda e: self._cb.get("change_theme")(theme_cb.get()))

    def update_color_buttons(self, bg_hex: str | None, fg_hex: str | None) -> None:
        if hasattr(self, "_bg_btn"):
            self._bg_btn.config(bg=bg_hex or "#FFFF00")
        if hasattr(self, "_fg_btn"):
            self._fg_btn.config(fg=fg_hex or "#FF0000")


# =============================================================================
# Formula Bar with Name Box & Autocomplete
# =============================================================================

class FormulaBar(ttk.Frame):
    """Excel Formula Bar with Name Box, fx Wizard button, and Autocomplete."""

    def __init__(
        self,
        parent: tk.Widget,
        on_commit: Callable[[str], None],
        on_fx_clicked: Callable[[], None],
        on_goto_cell: Callable[[str], None]
    ) -> None:
        super().__init__(parent, padding=(4, 2))
        self._on_commit = on_commit
        self._on_fx = on_fx_clicked
        self._on_goto = on_goto_cell
        self._build_ui()

    def _build_ui(self) -> None:
        # Name Box (Cell Reference / Jump Box)
        self._name_box = ttk.Entry(self, font=Config.FONT_BOLD, width=10, justify=tk.CENTER)
        self._name_box.pack(side=tk.LEFT, padx=(0, 3))
        self._name_box.insert(0, "A1")
        self._name_box.bind("<Return>", self._on_name_box_enter)

        # Separator
        ttk.Separator(self, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=2)

        # fx Wizard Button
        ttk.Button(self, text="fx", width=3, bootstyle="secondary-outline", command=self._on_fx).pack(side=tk.LEFT, padx=(0, 4))

        # Main Formula / Value Entry
        self._entry = ttk.Entry(self, font=Config.FONT)
        self._entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self._entry.bind("<Return>", lambda e: self._on_commit(self._entry.get()))
        self._entry.bind("<Escape>", lambda e: self._entry.delete(0, tk.END))

    def _on_name_box_enter(self, event: tk.Event) -> None:
        target = self._name_box.get().strip().upper()
        if target:
            self._on_goto(target)

    def update_cell(self, cell_ref_text: str, value: str) -> None:
        self._name_box.delete(0, tk.END)
        self._name_box.insert(0, cell_ref_text)
        self._entry.delete(0, tk.END)
        self._entry.insert(0, value)

    def set_value(self, value: str, cursor: int | None = None) -> None:
        self._entry.delete(0, tk.END)
        self._entry.insert(0, value)
        if cursor is not None:
            self._entry.icursor(cursor)

    def get_value(self) -> str:
        return self._entry.get()

    def focus_entry(self) -> None:
        self._entry.focus_set()
        self._entry.select_range(0, tk.END)

    @property
    def entry(self) -> ttk.Entry:
        return self._entry


# =============================================================================
# Sheet Tabs with Add Tab (+) & Context Menu
# =============================================================================

class SheetTabs(ttk.Frame):
    """Excel-style scrollable sheet tab bar with '+' button and context menu."""

    def __init__(
        self,
        parent: tk.Widget,
        on_select_sheet: Callable[[str], None],
        on_add_sheet: Callable[[], None],
        on_rename_sheet: Callable[[str], None],
        on_delete_sheet: Callable[[str], None],
        on_duplicate_sheet: Callable[[str], None]
    ) -> None:
        super().__init__(parent, padding=(4, 2))
        self._on_select = on_select_sheet
        self._on_add = on_add_sheet
        self._on_rename = on_rename_sheet
        self._on_delete = on_delete_sheet
        self._on_duplicate = on_duplicate_sheet

        self._tabs: dict[str, ttk.Button] = {}
        self._tab_colors: dict[str, str] = {}
        self._active: str | None = None
        self._names: list[str] = []

        self._build_ui()

    def _build_ui(self) -> None:
        # Navigation arrows
        nav_frame = ttk.Frame(self)
        nav_frame.pack(side=tk.LEFT)

        w = 2
        ttk.Button(nav_frame, text="⏮", width=w, bootstyle="secondary-outline", command=self._scroll_start).pack(side=tk.LEFT, padx=1)
        ttk.Button(nav_frame, text="◀", width=w, bootstyle="secondary-outline", command=self._scroll_left).pack(side=tk.LEFT, padx=1)
        ttk.Button(nav_frame, text="▶", width=w, bootstyle="secondary-outline", command=self._scroll_right).pack(side=tk.LEFT, padx=1)
        ttk.Button(nav_frame, text="⏭", width=w, bootstyle="secondary-outline", command=self._scroll_end).pack(side=tk.LEFT, padx=1)

        # '+' Add sheet button
        ttk.Button(self, text=" ➕ ", bootstyle="success-outline", command=self._on_add).pack(side=tk.LEFT, padx=(3, 6))

        # Canvas for scrollable sheet tabs
        self._canvas = tk.Canvas(self, height=28, highlightthickness=0, bd=0)
        self._canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self._inner = ttk.Frame(self._canvas)
        self._canvas.create_window((0, 0), window=self._inner, anchor=tk.NW)

        self._inner.bind("<Configure>", lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<MouseWheel>", lambda e: self._canvas.xview_scroll(-1 if e.delta > 0 else 1, "units"))

    def _scroll_start(self) -> None:
        self._canvas.xview_moveto(0)

    def _scroll_end(self) -> None:
        self._canvas.xview_moveto(1)

    def _scroll_left(self) -> None:
        self._canvas.xview_scroll(-3, "units")

    def _scroll_right(self) -> None:
        self._canvas.xview_scroll(3, "units")

    def set_sheets(self, names: list[str]) -> None:
        for btn in self._tabs.values():
            btn.destroy()
        self._tabs.clear()
        self._names = list(names)

        for name in names:
            btn = ttk.Button(
                self._inner, text=f" {name} ",
                bootstyle="secondary-outline",
                command=partial(self._select, name)
            )
            btn.pack(side=tk.LEFT, padx=2)
            btn.bind("<Button-3>", partial(self._show_tab_context_menu, name))
            btn.bind("<Double-1>", partial(self._on_double_click_tab, name))
            self._tabs[name] = btn

        if names:
            self._select(names[0])

    def _select(self, name: str) -> None:
        if self._active and self._active in self._tabs:
            self._tabs[self._active].configure(bootstyle="secondary-outline")

        self._active = name
        if name in self._tabs:
            self._tabs[name].configure(bootstyle="success")
            self._scroll_to_tab(name)

        self._on_select(name)

    def _scroll_to_tab(self, name: str) -> None:
        if name not in self._tabs:
            return
        btn = self._tabs[name]
        self._canvas.update_idletasks()
        btn_x = btn.winfo_x()
        btn_w = btn.winfo_width()
        canvas_w = self._canvas.winfo_width()
        bbox = self._canvas.bbox("all")
        if not bbox:
            return
        total_w = bbox[2] - bbox[0]
        if total_w <= canvas_w:
            return
        left = self._canvas.xview()[0] * total_w
        right = left + canvas_w
        if btn_x < left:
            self._canvas.xview_moveto(btn_x / total_w)
        elif btn_x + btn_w > right:
            self._canvas.xview_moveto((btn_x + btn_w - canvas_w) / total_w)

    def _show_tab_context_menu(self, sheet_name: str, event: tk.Event) -> None:
        menu = Menu(self, tearoff=0)
        menu.add_command(label=f"Rename '{sheet_name}'...", command=lambda: self._on_rename(sheet_name))
        menu.add_command(label=f"Duplicate '{sheet_name}'", command=lambda: self._on_duplicate(sheet_name))
        menu.add_separator()
        menu.add_command(label=f"Delete '{sheet_name}'", command=lambda: self._on_delete(sheet_name))
        menu.tk_popup(event.x_root, event.y_root)

    def _on_double_click_tab(self, sheet_name: str, event: tk.Event) -> None:
        self._on_rename(sheet_name)

    @property
    def active(self) -> str | None:
        return self._active


# =============================================================================
# Range-Aware Status Bar with Dynamic Math Stats
# =============================================================================

class StatusBar(ttk.Frame):
    """Excel Status Bar with Live Dynamic Range Calculations."""

    def __init__(self, parent: tk.Widget, on_zoom_change: Callable[[int], None] | None = None) -> None:
        super().__init__(parent, padding=(6, 2))
        self._on_zoom = on_zoom_change
        self._current_zoom = 100
        self._build_ui()

    def _build_ui(self) -> None:
        # Ready / Mode indicator
        self._mode_label = ttk.Label(self, text="Ready", font=Config.FONT_BOLD, width=8)
        self._mode_label.pack(side=tk.LEFT)

        ttk.Separator(self, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6)

        # Cell coordinate info
        self._cell_label = ttk.Label(self, text="A1", width=12)
        self._cell_label.pack(side=tk.LEFT)

        # Dynamic Range Stats (Average, Count, Min, Max, Sum)
        self._stats_label = ttk.Label(self, text="", font=Config.FONT, bootstyle="inverse-light")
        self._stats_label.pack(side=tk.LEFT, padx=10)

        # Right side info
        self._modified_label = ttk.Label(self, text="", bootstyle="warning", width=10)
        self._modified_label.pack(side=tk.RIGHT, padx=4)

        self._grid_size_label = ttk.Label(self, text="", width=18)
        self._grid_size_label.pack(side=tk.RIGHT, padx=4)

        # Zoom Controls
        zoom_frame = ttk.Frame(self)
        zoom_frame.pack(side=tk.RIGHT, padx=6)
        ttk.Button(zoom_frame, text="−", width=2, bootstyle="secondary-outline", command=self._zoom_out).pack(side=tk.LEFT, padx=1)
        self._zoom_label = ttk.Label(zoom_frame, text="100%", width=5, anchor=tk.CENTER)
        self._zoom_label.pack(side=tk.LEFT, padx=2)
        ttk.Button(zoom_frame, text="+", width=2, bootstyle="secondary-outline", command=self._zoom_in).pack(side=tk.LEFT, padx=1)

    def set_mode(self, mode: str) -> None:
        self._mode_label.config(text=mode)

    def set_cell(self, text: str) -> None:
        self._cell_label.config(text=text)

    def set_stats(self, rows: int, cols: int) -> None:
        self._grid_size_label.config(text=f"{rows:,} R × {cols} C")

    def set_modified(self, modified: bool) -> None:
        self._modified_label.config(text="● Modified" if modified else "")

    def update_range_stats(self, values: list[Any]) -> None:
        """Dynamically compute and display Excel statistics for selected cells."""
        if not values or len(values) <= 1:
            self._stats_label.config(text="")
            return

        nums = []
        non_empty = 0
        for v in values:
            if v is not None and v != "":
                non_empty += 1
                try:
                    num = float(str(v).replace(",", ".").replace(" ", "").replace("₽", "").replace("$", "").replace("%", ""))
                    nums.append(num)
                except (ValueError, TypeError):
                    pass

        parts = []
        if nums:
            avg_val = sum(nums) / len(nums)
            parts.append(f"AVERAGE: {avg_val:,.2f}" if avg_val != int(avg_val) else f"AVERAGE: {int(avg_val)}")
            parts.append(f"COUNT: {non_empty}")
            parts.append(f"NUMS: {len(nums)}")
            min_val = min(nums)
            max_val = max(nums)
            sum_val = sum(nums)
            parts.append(f"MIN: {min_val:g}")
            parts.append(f"MAX: {max_val:g}")
            parts.append(f"SUM: {sum_val:,.2f}" if sum_val != int(sum_val) else f"SUM: {int(sum_val)}")
        elif non_empty > 0:
            parts.append(f"COUNT: {non_empty}")

        self._stats_label.config(text="   |   ".join(parts))

    def _zoom_in(self) -> None:
        if self._current_zoom < 200:
            self._current_zoom = min(200, self._current_zoom + 15)
            self._zoom_label.config(text=f"{self._current_zoom}%")
            if self._on_zoom:
                self._on_zoom(self._current_zoom)

    def _zoom_out(self) -> None:
        if self._current_zoom > 50:
            self._current_zoom = max(50, self._current_zoom - 15)
            self._zoom_label.config(text=f"{self._current_zoom}%")
            if self._on_zoom:
                self._on_zoom(self._current_zoom)


# =============================================================================
# Inline Cell Editor
# =============================================================================

class CellEditor:
    """Inline cell editor overlay inside Treeview."""

    __slots__ = ("_tree", "_entry", "_on_commit", "_on_cancel", "_commit_on_focus_out")

    def __init__(
        self,
        tree: ttk.Treeview,
        on_commit: Callable[[str], None],
        on_cancel: Callable[[], None]
    ) -> None:
        self._tree = tree
        self._entry: tk.Entry | None = None
        self._on_commit = on_commit
        self._on_cancel = on_cancel
        self._commit_on_focus_out = True

    def start(self, iid: str, col_id: str, value: str, initial_char: str | None = None) -> None:
        self.cancel()
        self._commit_on_focus_out = True

        bbox = self._tree.bbox(iid, col_id)
        if not bbox:
            self._tree.see(iid)
            self._tree.update_idletasks()
            bbox = self._tree.bbox(iid, col_id)
            if not bbox:
                return

        x, y, w, h = bbox

        self._entry = tk.Entry(
            self._tree,
            font=Config.FONT,
            bg=Config.EDIT_BG,
            fg=Config.CELL_FG,
            insertbackground=Config.CELL_FG,
            selectbackground=Config.SELECTION_BG,
            relief="solid",
            bd=2
        )
        self._entry.place(x=x, y=y, width=w, height=h)

        if initial_char:
            self._entry.insert(0, initial_char)
        else:
            self._entry.insert(0, value)
            self._entry.select_range(0, tk.END)

        self._entry.focus_set()
        self._entry.bind("<Return>", self._commit)
        self._entry.bind("<Tab>", self._commit_tab)
        self._entry.bind("<Escape>", lambda e: self.cancel())
        self._entry.bind("<FocusOut>", self._on_focus_out)

    def _on_focus_out(self, event: tk.Event = None) -> str | None:
        if not self._commit_on_focus_out:
            return None
        return self._commit(event)

    def _commit(self, event: tk.Event = None) -> str:
        if self._entry:
            val = self._entry.get()
            self.cancel()
            self._on_commit(val)
        return "break"

    def _commit_tab(self, event: tk.Event = None) -> str:
        if self._entry:
            val = self._entry.get()
            self.cancel()
            self._on_commit(val)
        return "break"

    def cancel(self) -> None:
        if self._entry:
            self._entry.destroy()
            self._entry = None
            self._on_cancel()

    def set_focus_commit_enabled(self, enabled: bool) -> None:
        self._commit_on_focus_out = enabled

    def get_value(self) -> str:
        return self._entry.get() if self._entry else ""

    def set_value(self, value: str, cursor: int | None = None) -> None:
        if not self._entry:
            return
        self._entry.delete(0, tk.END)
        self._entry.insert(0, value)
        if cursor is not None:
            self._entry.icursor(cursor)

    @property
    def is_active(self) -> bool:
        return self._entry is not None

    @property
    def entry(self) -> tk.Entry | None:
        return self._entry

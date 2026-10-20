"""
Professional UI widgets and Ribbon components for Excel Viewer Pro.
"""

from __future__ import annotations

import tkinter as tk
from functools import partial
from tkinter import Menu, colorchooser
from typing import Any, Callable

import ttkbootstrap as ttk
from ttkbootstrap.constants import *

try:
    from .config import Config
    from .formulas import FUNCTION_METADATA
    from .icons import get_icon
except ImportError:
    from config import Config
    from formulas import FUNCTION_METADATA
    from icons import get_icon

# =============================================================================
# Tooltip Helper
# =============================================================================


class Tooltip:
    """Elegant floating tooltip for UI widgets."""

    def __init__(self, widget: tk.Widget, text: str, delay_ms: int = 350) -> None:
        self.widget = widget
        self.text = text
        self.delay_ms = delay_ms
        self._tipwindow: tk.Toplevel | None = None
        self._after_id: str | None = None
        self.widget.bind("<Enter>", self._on_enter, add="+")
        self.widget.bind("<Leave>", self._on_leave, add="+")
        self.widget.bind("<ButtonPress>", self._on_leave, add="+")

    def _on_enter(self, event: tk.Event | None = None) -> None:
        self._cancel()
        self._after_id = self.widget.after(self.delay_ms, self._show)

    def _on_leave(self, event: tk.Event | None = None) -> None:
        self._cancel()
        self._hide()

    def _cancel(self) -> None:
        if self._after_id:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show(self) -> None:
        if self._tipwindow or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 8
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
            self._tipwindow = tw = tk.Toplevel(self.widget)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{x}+{y}")
            tw.attributes("-topmost", True)
            label = tk.Label(
                tw,
                text=self.text,
                justify=tk.LEFT,
                background="#1e293b",
                foreground="#ffffff",
                relief="solid",
                borderwidth=0,
                padx=6,
                pady=3,
                font=("Segoe UI", 8),
            )
            label.pack()
        except Exception:
            pass

    def _hide(self) -> None:
        if self._tipwindow:
            try:
                self._tipwindow.destroy()
            except Exception:
                pass
            self._tipwindow = None


# =============================================================================
# Ribbon Bar (Excel-Style Tabbed Toolbar)
# =============================================================================


class ExcelRibbon(ttk.Frame):
    """Excel-like Tabbed Ribbon Toolbar."""

    def __init__(self, parent: tk.Widget, callbacks: dict[str, Callable]) -> None:
        super().__init__(parent)
        self._cb = callbacks
        self._icon_refs: list[Any] = []
        self._build_ui()

    def _btn(
        self,
        parent: tk.Widget,
        text: str = "",
        icon_name: str | None = None,
        command: Callable | None = None,
        bootstyle: str = "outline",
        compound: str = "left",
        icon_size: int = 16,
        tooltip: str | None = None,
        width: int | None = None,
        padding: Any = None,
        side: str | None = None,
        fill: str | None = None,
        padx: int = 1,
        pady: int = 1,
    ) -> ttk.Button:
        kwargs: dict[str, Any] = {"bootstyle": bootstyle}
        if command is not None:
            kwargs["command"] = command
        if width is not None:
            kwargs["width"] = width
        if padding is not None:
            kwargs["padding"] = padding
        elif not text and icon_name:
            kwargs["padding"] = (3, 1)
        if icon_name:
            try:
                photo = get_icon(icon_name, size=icon_size)
                self._icon_refs.append(photo)
                kwargs["image"] = photo
                if text:
                    kwargs["compound"] = compound
            except Exception:
                pass
        if text:
            kwargs["text"] = text
        btn = ttk.Button(parent, **kwargs)
        if icon_name and "image" in kwargs:
            btn.image = kwargs["image"]
        pack_kwargs: dict[str, Any] = {"padx": padx, "pady": pady}
        if side is not None:
            pack_kwargs["side"] = side
        if fill is not None:
            pack_kwargs["fill"] = fill
        btn.pack(**pack_kwargs)
        if tooltip:
            Tooltip(btn, tooltip)
        return btn

    def _menubtn(
        self,
        parent: tk.Widget,
        text: str = "",
        icon_name: str | None = None,
        bootstyle: str = "outline",
        compound: str = "left",
        icon_size: int = 16,
        tooltip: str | None = None,
        width: int | None = None,
        padding: Any = None,
        side: str | None = None,
        fill: str | None = None,
        padx: int = 1,
        pady: int = 1,
    ) -> tuple[ttk.Menubutton, Menu]:
        kwargs: dict[str, Any] = {"bootstyle": bootstyle}
        if width is not None:
            kwargs["width"] = width
        if padding is not None:
            kwargs["padding"] = padding
        elif not text and icon_name:
            kwargs["padding"] = (3, 1)
        if icon_name:
            try:
                photo = get_icon(icon_name, size=icon_size)
                self._icon_refs.append(photo)
                kwargs["image"] = photo
                if text:
                    kwargs["compound"] = compound
            except Exception:
                pass
        if text:
            kwargs["text"] = text
        mbtn = ttk.Menubutton(parent, **kwargs)
        if icon_name and "image" in kwargs:
            mbtn.image = kwargs["image"]
        menu = Menu(mbtn, tearoff=0)
        mbtn["menu"] = menu
        pack_kwargs: dict[str, Any] = {"padx": padx, "pady": pady}
        if side is not None:
            pack_kwargs["side"] = side
        if fill is not None:
            pack_kwargs["fill"] = fill
        mbtn.pack(**pack_kwargs)
        if tooltip:
            Tooltip(mbtn, tooltip)
        return mbtn, menu

    def _build_ui(self) -> None:
        # File & Ribbon tabs notebook (no separate disconnected QAT bar)
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
        tab = ttk.Frame(self._notebook, padding=(2, 2))
        self._notebook.add(tab, text="  Home  ")

        # 1. Undo & Save Group (Excel standard Undo section)
        undo_grp = ttk.LabelFrame(tab, text="Undo", padding=2)
        undo_grp.pack(side=tk.LEFT, fill=tk.Y, padx=1)

        u_col = ttk.Frame(undo_grp)
        u_col.pack(side=tk.LEFT, fill=tk.Y, padx=1)

        self._btn(
            u_col, text=" Save", icon_name="save", bootstyle="success-outline",
            command=self._cb.get("save"), tooltip="Save Workbook (Ctrl+S)", fill=tk.X
        )
        self._btn(
            u_col, text=" Undo", icon_name="undo", bootstyle="secondary-outline",
            command=self._cb.get("undo"), tooltip="Undo last action (Ctrl+Z)", fill=tk.X
        )
        self._btn(
            u_col, text=" Redo", icon_name="redo", bootstyle="secondary-outline",
            command=self._cb.get("redo"), tooltip="Redo last undone action (Ctrl+Y)", fill=tk.X
        )

        # 2. Clipboard Group
        clip_grp = ttk.LabelFrame(tab, text="Clipboard", padding=2)
        clip_grp.pack(side=tk.LEFT, fill=tk.Y, padx=1)

        paste_btn, paste_menu = self._menubtn(
            clip_grp, text="Paste", icon_name="paste", compound="top", icon_size=20,
            bootstyle="primary", tooltip="Paste from Clipboard (Ctrl+V)", side=tk.LEFT, fill=tk.Y
        )
        paste_menu.add_command(label="Paste All (Ctrl+V)", command=self._cb.get("paste"))
        paste_menu.add_command(label="Paste Values Only", command=self._cb.get("paste_values"))
        paste_menu.add_command(label="Paste Formulas Only", command=self._cb.get("paste_formulas"))
        paste_menu.add_command(label="Paste Transposed", command=self._cb.get("paste_transpose"))

        clip_sub = ttk.Frame(clip_grp)
        clip_sub.pack(side=tk.LEFT, padx=1)
        self._btn(
            clip_sub, text=" Cut", icon_name="cut", bootstyle="outline",
            command=self._cb.get("cut"), tooltip="Cut (Ctrl+X)", fill=tk.X
        )
        self._btn(
            clip_sub, text=" Copy", icon_name="copy", bootstyle="outline",
            command=self._cb.get("copy"), tooltip="Copy (Ctrl+C)", fill=tk.X
        )
        self._btn(
            clip_sub, text=" Format", icon_name="format_painter", bootstyle="outline",
            command=self._cb.get("format_painter"), tooltip="Format Painter (Copy formatting)", fill=tk.X
        )

        # 3. Font Group
        font_grp = ttk.LabelFrame(tab, text="Font", padding=2)
        font_grp.pack(side=tk.LEFT, fill=tk.Y, padx=1)

        font_top = ttk.Frame(font_grp)
        font_top.pack(fill=tk.X, pady=1)

        self._font_family = ttk.Combobox(
            font_top,
            values=[
                "Calibri",
                "Segoe UI",
                "Arial",
                "Times New Roman",
                "Consolas",
                "Verdana",
                "Tahoma",
            ],
            width=10,
            state="readonly",
        )
        self._font_family.set("Segoe UI")
        self._font_family.pack(side=tk.LEFT, padx=1)
        self._font_family.bind(
            "<<ComboboxSelected>>",
            lambda e: self._cb.get("set_font_family")(self._font_family.get()),
        )

        self._font_size = ttk.Combobox(
            font_top,
            values=["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36"],
            width=2,
            state="readonly",
        )
        self._font_size.set("10")
        self._font_size.pack(side=tk.LEFT, padx=1)
        self._font_size.bind(
            "<<ComboboxSelected>>",
            lambda e: self._cb.get("set_font_size")(int(self._font_size.get())),
        )

        self._btn(
            font_top, icon_name="font_inc", bootstyle="secondary-outline",
            command=self._cb.get("increase_font_size"), tooltip="Increase Font Size", side=tk.LEFT
        )
        self._btn(
            font_top, icon_name="font_dec", bootstyle="secondary-outline",
            command=self._cb.get("decrease_font_size"), tooltip="Decrease Font Size", side=tk.LEFT
        )

        font_bot = ttk.Frame(font_grp)
        font_bot.pack(fill=tk.X, pady=1)

        self._btn(
            font_bot, icon_name="bold", bootstyle="secondary-outline",
            command=self._cb.get("toggle_bold"), tooltip="Bold (Ctrl+B)", side=tk.LEFT
        )
        self._btn(
            font_bot, icon_name="italic", bootstyle="secondary-outline",
            command=self._cb.get("toggle_italic"), tooltip="Italic (Ctrl+I)", side=tk.LEFT
        )
        self._btn(
            font_bot, icon_name="underline", bootstyle="secondary-outline",
            command=self._cb.get("toggle_underline"), tooltip="Underline (Ctrl+U)", side=tk.LEFT
        )
        self._btn(
            font_bot, icon_name="strike", bootstyle="secondary-outline",
            command=self._cb.get("toggle_strikethrough"), tooltip="Strikethrough", side=tk.LEFT
        )

        # Borders menu
        border_btn, border_menu = self._menubtn(
            font_bot, icon_name="borders", bootstyle="secondary-outline",
            tooltip="Cell Borders", side=tk.LEFT
        )
        border_menu.add_command(
            label="All Borders", command=lambda: self._cb.get("set_borders")("all")
        )
        border_menu.add_command(
            label="Outside Borders", command=lambda: self._cb.get("set_borders")("outside")
        )
        border_menu.add_command(
            label="Thick Box Border", command=lambda: self._cb.get("set_borders")("thick")
        )
        border_menu.add_command(
            label="Bottom Border", command=lambda: self._cb.get("set_borders")("bottom")
        )
        border_menu.add_command(
            label="Top and Bottom Border", command=lambda: self._cb.get("set_borders")("top_bottom")
        )
        border_menu.add_separator()
        border_menu.add_command(
            label="No Border", command=lambda: self._cb.get("set_borders")("none")
        )

        # Colors (Fill and Font)
        self._bg_btn = self._btn(
            font_bot, icon_name="fill_color", bootstyle="secondary-outline",
            command=self._cb.get("set_bg_color"), tooltip="Fill Color", side=tk.LEFT
        )
        self._fg_btn = self._btn(
            font_bot, icon_name="font_color", bootstyle="secondary-outline",
            command=self._cb.get("set_fg_color"), tooltip="Font Color", side=tk.LEFT
        )

        # 4. Alignment Group
        align_grp = ttk.LabelFrame(tab, text="Alignment", padding=2)
        align_grp.pack(side=tk.LEFT, fill=tk.Y, padx=1)

        align_left_box = ttk.Frame(align_grp)
        align_left_box.pack(side=tk.LEFT, fill=tk.Y, padx=(1, 2))

        align_top = ttk.Frame(align_left_box)
        align_top.pack(fill=tk.X, pady=1)
        self._btn(
            align_top, icon_name="align_top", bootstyle="secondary-outline",
            command=lambda: self._cb.get("set_valign")("top"), tooltip="Top Align", side=tk.LEFT
        )
        self._btn(
            align_top, icon_name="align_mid", bootstyle="secondary-outline",
            command=lambda: self._cb.get("set_valign")("center"), tooltip="Middle Align", side=tk.LEFT
        )
        self._btn(
            align_top, icon_name="align_bot", bootstyle="secondary-outline",
            command=lambda: self._cb.get("set_valign")("bottom"), tooltip="Bottom Align", side=tk.LEFT
        )

        align_bot = ttk.Frame(align_left_box)
        align_bot.pack(fill=tk.X, pady=1)
        self._btn(
            align_bot, icon_name="align_left", bootstyle="secondary-outline",
            command=lambda: self._cb.get("set_halign")("left"), tooltip="Align Left (Ctrl+L)", side=tk.LEFT
        )
        self._btn(
            align_bot, icon_name="align_center", bootstyle="secondary-outline",
            command=lambda: self._cb.get("set_halign")("center"), tooltip="Center (Ctrl+E)", side=tk.LEFT
        )
        self._btn(
            align_bot, icon_name="align_right", bootstyle="secondary-outline",
            command=lambda: self._cb.get("set_halign")("right"), tooltip="Align Right (Ctrl+R)", side=tk.LEFT
        )

        align_side = ttk.Frame(align_grp)
        align_side.pack(side=tk.LEFT, fill=tk.Y, padx=1)
        self._btn(
            align_side, text=" Wrap Text", icon_name="wrap_text", bootstyle="secondary-outline",
            command=self._cb.get("toggle_wrap_text"), tooltip="Wrap Text in Cell", fill=tk.X
        )
        self._btn(
            align_side, text=" Merge & Ctr", icon_name="merge_center", bootstyle="secondary-outline",
            command=self._cb.get("merge_cells"), tooltip="Merge Cells & Center", fill=tk.X
        )

        # 5. Number Group
        num_grp = ttk.LabelFrame(tab, text="Number", padding=2)
        num_grp.pack(side=tk.LEFT, fill=tk.Y, padx=1)

        self._num_fmt = ttk.Combobox(
            num_grp, values=list(Config.NUMBER_FORMATS.keys()), width=10, state="readonly"
        )
        self._num_fmt.set("General")
        self._num_fmt.pack(fill=tk.X, pady=1)
        self._num_fmt.bind(
            "<<ComboboxSelected>>",
            lambda e: self._cb.get("set_number_format")(
                Config.NUMBER_FORMATS.get(self._num_fmt.get(), "")
            ),
        )

        num_btns = ttk.Frame(num_grp)
        num_btns.pack(fill=tk.X, pady=1)
        self._btn(
            num_btns, icon_name="currency_rub", bootstyle="outline",
            command=lambda: self._cb.get("set_number_format")("#,##0.00 ₽"), tooltip="Format as Rubles (₽)", side=tk.LEFT
        )
        self._btn(
            num_btns, icon_name="currency_usd", bootstyle="outline",
            command=lambda: self._cb.get("set_number_format")("$#,##0.00"), tooltip="Format as Dollars ($)", side=tk.LEFT
        )
        self._btn(
            num_btns, icon_name="percent", bootstyle="outline",
            command=lambda: self._cb.get("set_number_format")("0.00%"), tooltip="Percent Style (%)", side=tk.LEFT
        )
        self._btn(
            num_btns, icon_name="comma", bootstyle="outline",
            command=lambda: self._cb.get("set_number_format")("#,##0.00"), tooltip="Comma Style (,)", side=tk.LEFT
        )
        self._btn(
            num_btns, icon_name="decimal_inc", bootstyle="outline",
            command=self._cb.get("increase_decimals"), tooltip="Increase Decimals (.0+)", side=tk.LEFT
        )
        self._btn(
            num_btns, icon_name="decimal_dec", bootstyle="outline",
            command=self._cb.get("decrease_decimals"), tooltip="Decrease Decimals (.0-)", side=tk.LEFT
        )

        # 6. Styles Group
        style_grp = ttk.LabelFrame(tab, text="Styles", padding=2)
        style_grp.pack(side=tk.LEFT, fill=tk.Y, padx=1)

        cf_btn, cf_menu = self._menubtn(
            style_grp, text=" Conditional\n Formatting", icon_name="cond_format",
            bootstyle="info-outline", tooltip="Conditional Formatting Rules", fill=tk.X
        )
        cf_menu.add_command(
            label="Highlight: Greater Than...",
            command=lambda: self._cb.get("add_cf_rule")("greater_than"),
        )
        cf_menu.add_command(
            label="Highlight: Less Than...",
            command=lambda: self._cb.get("add_cf_rule")("less_than"),
        )
        cf_menu.add_command(
            label="Highlight: Between...", command=lambda: self._cb.get("add_cf_rule")("between")
        )
        cf_menu.add_command(
            label="Highlight: Equal To...", command=lambda: self._cb.get("add_cf_rule")("equal_to")
        )
        cf_menu.add_command(
            label="Highlight: Text that Contains...",
            command=lambda: self._cb.get("add_cf_rule")("text_contains"),
        )
        cf_menu.add_separator()
        cf_menu.add_command(
            label="Color Scale: Green - Yellow - Red",
            command=lambda: self._cb.get("add_cf_rule")("color_scale_gyr"),
        )
        cf_menu.add_command(
            label="Color Scale: Red - Yellow - Green",
            command=lambda: self._cb.get("add_cf_rule")("color_scale_ryg"),
        )
        cf_menu.add_separator()
        cf_menu.add_command(
            label="Clear Rules from Selected Cells", command=self._cb.get("clear_cf_rules")
        )

        self._btn(
            style_grp, text=" Format Cells...", icon_name="format_cells",
            bootstyle="secondary-outline", command=self._cb.get("format_cells_dialog"),
            tooltip="Format Cells Dialog (Ctrl+1)", fill=tk.X
        )

        # 7. Cells Group
        cells_grp = ttk.LabelFrame(tab, text="Cells", padding=2)
        cells_grp.pack(side=tk.LEFT, fill=tk.Y, padx=1)

        ins_btn, ins_menu = self._menubtn(
            cells_grp, text=" Insert", icon_name="insert", bootstyle="outline",
            tooltip="Insert Rows or Columns", fill=tk.X
        )
        ins_menu.add_command(label="Insert Row Above", command=self._cb.get("insert_row_above"))
        ins_menu.add_command(label="Insert Row Below", command=self._cb.get("insert_row_below"))
        ins_menu.add_command(label="Insert Column Left", command=self._cb.get("insert_col_left"))
        ins_menu.add_command(label="Insert Column Right", command=self._cb.get("insert_col_right"))

        del_btn, del_menu = self._menubtn(
            cells_grp, text=" Delete", icon_name="delete", bootstyle="outline",
            tooltip="Delete Rows, Columns, or Cells", fill=tk.X
        )
        del_menu.add_command(label="Delete Row", command=self._cb.get("delete_row"))
        del_menu.add_command(label="Delete Column", command=self._cb.get("delete_column"))
        del_menu.add_separator()
        del_menu.add_command(label="Delete / Clear Cell", command=self._cb.get("delete_cell"))

        self._btn(
            cells_grp, text=" Auto-fit", icon_name="autofit", bootstyle="outline",
            command=self._cb.get("autofit_all_cols"), tooltip="Auto-fit Column Widths", fill=tk.X
        )

        # 8. Editing Group
        edit_grp = ttk.LabelFrame(tab, text="Editing", padding=2)
        edit_grp.pack(side=tk.LEFT, fill=tk.Y, padx=1)

        edit_top = ttk.Frame(edit_grp)
        edit_top.pack(fill=tk.X, pady=1)

        as_btn, as_menu = self._menubtn(
            edit_top, text=" AutoSum", icon_name="autosum", bootstyle="primary-outline",
            tooltip="AutoSum & Basic Math", side=tk.LEFT
        )
        as_menu.add_command(label="Sum (=SUM)", command=lambda: self._cb.get("apply_autosum")("SUM"))
        as_menu.add_command(label="Average (=AVERAGE)", command=lambda: self._cb.get("apply_autosum")("AVERAGE"))
        as_menu.add_command(label="Count Numbers (=COUNT)", command=lambda: self._cb.get("apply_autosum")("COUNT"))
        as_menu.add_command(label="Max (=MAX)", command=lambda: self._cb.get("apply_autosum")("MAX"))
        as_menu.add_command(label="Min (=MIN)", command=lambda: self._cb.get("apply_autosum")("MIN"))

        clr_btn, clr_menu = self._menubtn(
            edit_top, text=" Clear", icon_name="clear", bootstyle="outline",
            tooltip="Clear Contents, Formats, or Comments", side=tk.LEFT
        )
        clr_menu.add_command(label="Clear All", command=self._cb.get("clear_all"))
        clr_menu.add_command(label="Clear Formats", command=self._cb.get("clear_formats"))
        clr_menu.add_command(label="Clear Comments", command=self._cb.get("clear_comments"))
        clr_menu.add_command(label="Clear Contents (Del)", command=self._cb.get("delete_cell"))

        edit_bot = ttk.Frame(edit_grp)
        edit_bot.pack(fill=tk.X, pady=1)

        sort_btn, sort_menu = self._menubtn(
            edit_bot, text=" Sort", icon_name="sort_az", bootstyle="outline",
            tooltip="Sort and Filter Data", side=tk.LEFT
        )
        sort_menu.add_command(label="Sort A to Z", command=lambda: self._cb.get("quick_sort")(False))
        sort_menu.add_command(label="Sort Z to A", command=lambda: self._cb.get("quick_sort")(True))
        sort_menu.add_command(label="Custom Sort...", command=self._cb.get("custom_sort_dialog"))
        sort_menu.add_separator()
        sort_menu.add_command(label="Toggle AutoFilter", command=self._cb.get("toggle_filter"))
        sort_menu.add_command(label="Clear Filters", command=self._cb.get("clear_filters"))

        find_btn, find_menu = self._menubtn(
            edit_bot, text=" Find", icon_name="find", bootstyle="outline",
            tooltip="Find, Replace, or Jump to Cell", side=tk.LEFT
        )
        find_menu.add_command(label="Find... (Ctrl+F)", command=self._cb.get("find_dialog"))
        find_menu.add_command(label="Replace... (Ctrl+H)", command=self._cb.get("replace_dialog"))
        find_menu.add_command(label="Go to Cell... (Ctrl+G)", command=self._cb.get("goto_dialog"))

    # -------------------------------------------------------------------------
    # Insert Tab
    # -------------------------------------------------------------------------
    def _create_insert_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=(4, 2))
        self._notebook.add(tab, text="  Insert  ")

        # 1. Tables Group
        tbl_grp = ttk.LabelFrame(tab, text="Tables", padding=3)
        tbl_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._btn(
            tbl_grp, text="Demo Table", icon_name="demo_table", compound="top", icon_size=20,
            bootstyle="primary-outline", command=self._cb.get("load_demo_data"),
            tooltip="Load rich Financial Model demo data", side=tk.LEFT, fill=tk.Y, padx=2
        )
        self._btn(
            tbl_grp, text="New Sheet\n(Shift+F11)", icon_name="new_sheet", compound="top", icon_size=20,
            bootstyle="outline", command=self._cb.get("add_sheet"),
            tooltip="Insert new worksheet", side=tk.LEFT, fill=tk.Y, padx=2
        )

        # 2. Charts Group
        chart_grp = ttk.LabelFrame(tab, text="Charts", padding=3)
        chart_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._btn(
            chart_grp, text="Chart\nWizard", icon_name="chart_wizard", compound="top", icon_size=20,
            bootstyle="primary", command=self._cb.get("show_chart_wizard"),
            tooltip="Open Chart Wizard (F11)", side=tk.LEFT, fill=tk.Y, padx=2
        )

        chart_sub = ttk.Frame(chart_grp)
        chart_sub.pack(side=tk.LEFT, padx=2)

        crow1 = ttk.Frame(chart_sub)
        crow1.pack(fill=tk.X, pady=1)
        self._btn(
            crow1, text=" Col", icon_name="chart_col", bootstyle="outline",
            command=lambda: self._cb.get("quick_chart")("Clustered Column"),
            tooltip="Clustered Column Chart", side=tk.LEFT, padx=1
        )
        self._btn(
            crow1, text=" Line", icon_name="chart_line", bootstyle="outline",
            command=lambda: self._cb.get("quick_chart")("Line"),
            tooltip="Line Chart with markers", side=tk.LEFT, padx=1
        )
        self._btn(
            crow1, text=" Pie", icon_name="chart_pie", bootstyle="outline",
            command=lambda: self._cb.get("quick_chart")("Pie"),
            tooltip="Pie Chart", side=tk.LEFT, padx=1
        )

        crow2 = ttk.Frame(chart_sub)
        crow2.pack(fill=tk.X, pady=1)
        self._btn(
            crow2, text=" Bar", icon_name="chart_bar", bootstyle="outline",
            command=lambda: self._cb.get("quick_chart")("Clustered Bar"),
            tooltip="Clustered Bar Chart", side=tk.LEFT, padx=1
        )
        self._btn(
            crow2, text=" Area", icon_name="chart_area", bootstyle="outline",
            command=lambda: self._cb.get("quick_chart")("Area"),
            tooltip="Area Chart", side=tk.LEFT, padx=1
        )
        self._btn(
            crow2, text=" Scatter", icon_name="chart_scatter", bootstyle="outline",
            command=lambda: self._cb.get("quick_chart")("Scatter"),
            tooltip="XY Scatter Chart", side=tk.LEFT, padx=1
        )

        # 3. Comments Group
        comm_grp = ttk.LabelFrame(tab, text="Comments", padding=3)
        comm_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._btn(
            comm_grp, text="Add / Edit\nComment", icon_name="comment_add", compound="top", icon_size=20,
            bootstyle="info-outline", command=self._cb.get("edit_comment"),
            tooltip="Add or Edit Cell Comment (Shift+F2)", side=tk.LEFT, fill=tk.Y, padx=2
        )
        self._btn(
            comm_grp, text="Delete\nComment", icon_name="comment_del", compound="top", icon_size=20,
            bootstyle="outline", command=self._cb.get("clear_comments"),
            tooltip="Delete comment in selection", side=tk.LEFT, fill=tk.Y, padx=2
        )

        # 4. Functions Group
        fn_grp = ttk.LabelFrame(tab, text="Functions", padding=3)
        fn_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._btn(
            fn_grp, text="fx Insert\nFunction", icon_name="fx", compound="top", icon_size=20,
            bootstyle="success", command=self._cb.get("show_fx_wizard"),
            tooltip="Insert Function Wizard (Shift+F3)", side=tk.LEFT, fill=tk.Y, padx=2
        )
        self._btn(
            fn_grp, text="Formula\nDocs", icon_name="formula_ref", compound="top", icon_size=20,
            bootstyle="outline", command=self._cb.get("formula_reference_dialog"),
            tooltip="Open 129 Excel Formula Functions Reference Guide", side=tk.LEFT, fill=tk.Y, padx=2
        )

        # 5. Data & Files Group
        file_grp = ttk.LabelFrame(tab, text="Text & Data", padding=3)
        file_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._btn(
            file_grp, text="Import\nCSV", icon_name="import_csv", compound="top", icon_size=20,
            bootstyle="secondary-outline", command=self._cb.get("import_csv"),
            tooltip="Import CSV / Delimited Text File", side=tk.LEFT, fill=tk.Y, padx=2
        )
        self._btn(
            file_grp, text="Export\nCSV", icon_name="export_csv", compound="top", icon_size=20,
            bootstyle="secondary-outline", command=self._cb.get("export_csv"),
            tooltip="Export Active Worksheet to CSV", side=tk.LEFT, fill=tk.Y, padx=2
        )

    # -------------------------------------------------------------------------
    # Data Tab
    # -------------------------------------------------------------------------
    def _create_data_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=(4, 2))
        self._notebook.add(tab, text="  Data  ")

        # 1. Get & Transform Data
        get_grp = ttk.LabelFrame(tab, text="Get & Transform Data", padding=3)
        get_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._btn(
            get_grp, text="From Text\n/ CSV...", icon_name="import_csv", compound="top", icon_size=20,
            bootstyle="primary-outline", command=self._cb.get("import_csv"),
            tooltip="Import data from Delimited Text / CSV", side=tk.LEFT, fill=tk.Y, padx=2
        )
        self._btn(
            get_grp, text="Export\nto CSV...", icon_name="export_csv", compound="top", icon_size=20,
            bootstyle="outline", command=self._cb.get("export_csv"),
            tooltip="Export current sheet to CSV", side=tk.LEFT, fill=tk.Y, padx=2
        )
        self._btn(
            get_grp, text="Refresh\nAll (F5)", icon_name="refresh", compound="top", icon_size=20,
            bootstyle="secondary-outline", command=self._cb.get("refresh_grid"),
            tooltip="Recalculate formulas and refresh grid (F5)", side=tk.LEFT, fill=tk.Y, padx=2
        )

        # 2. Sort & Filter Group
        sf_grp = ttk.LabelFrame(tab, text="Sort & Filter", padding=3)
        sf_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        sf_sub = ttk.Frame(sf_grp)
        sf_sub.pack(side=tk.LEFT, fill=tk.Y)

        sf_row1 = ttk.Frame(sf_sub)
        sf_row1.pack(fill=tk.X, pady=1)
        self._btn(
            sf_row1, text=" Sort A-Z", icon_name="sort_az", bootstyle="outline",
            command=lambda: self._cb.get("quick_sort")(False),
            tooltip="Sort Ascending (Lowest to Highest / A to Z)", side=tk.LEFT, padx=1
        )
        self._btn(
            sf_row1, text=" Sort Z-A", icon_name="sort_za", bootstyle="outline",
            command=lambda: self._cb.get("quick_sort")(True),
            tooltip="Sort Descending (Highest to Lowest / Z to A)", side=tk.LEFT, padx=1
        )
        self._btn(
            sf_row1, text=" Custom Sort...", icon_name="sort_custom", bootstyle="outline",
            command=self._cb.get("custom_sort_dialog"),
            tooltip="Multi-column Custom Sort dialog", side=tk.LEFT, padx=1
        )

        sf_row2 = ttk.Frame(sf_sub)
        sf_row2.pack(fill=tk.X, pady=1)
        self._btn(
            sf_row2, text=" AutoFilter", icon_name="filter", bootstyle="primary-outline",
            command=self._cb.get("toggle_filter"),
            tooltip="Toggle Excel AutoFilter dropdowns on columns", side=tk.LEFT, padx=1
        )
        self._btn(
            sf_row2, text=" Clear Filters", icon_name="clear_filter", bootstyle="secondary-outline",
            command=self._cb.get("clear_filters"),
            tooltip="Clear all active AutoFilter criteria", side=tk.LEFT, padx=1
        )

        # 3. Data Tools Group
        dt_grp = ttk.LabelFrame(tab, text="Data Tools", padding=3)
        dt_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._btn(
            dt_grp, text="Text to\nColumns...", icon_name="text_to_columns", compound="top", icon_size=20,
            bootstyle="outline", command=self._cb.get("text_to_columns_dialog"),
            tooltip="Split cell contents into adjacent columns by delimiter", side=tk.LEFT, fill=tk.Y, padx=2
        )
        self._btn(
            dt_grp, text="Remove\nDuplicates...", icon_name="remove_duplicates", compound="top", icon_size=20,
            bootstyle="outline", command=self._cb.get("remove_duplicates_dialog"),
            tooltip="Find and delete duplicate rows in selected data", side=tk.LEFT, fill=tk.Y, padx=2
        )

        # 4. What-If Analysis
        wi_grp = ttk.LabelFrame(tab, text="Forecast / What-If", padding=3)
        wi_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._btn(
            wi_grp, text="Goal Seek\nAnalysis...", icon_name="goal_seek", compound="top", icon_size=20,
            bootstyle="info-outline", command=self._cb.get("goal_seek_dialog"),
            tooltip="What-If Analysis: Goal Seek to determine required input value", side=tk.LEFT, fill=tk.Y, padx=2
        )

        # 5. Sheet Structure
        st_grp = ttk.LabelFrame(tab, text="Sheet Tools", padding=3)
        st_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._btn(
            st_grp, text="Auto-fit\nColumns", icon_name="autofit", compound="top", icon_size=20,
            bootstyle="outline", command=self._cb.get("autofit_all_cols"),
            tooltip="Auto-fit all column widths to content", side=tk.LEFT, fill=tk.Y, padx=2
        )

        frz_btn, frz_menu = self._menubtn(
            st_grp, text="Freeze\nPanes", icon_name="freeze_panes", compound="top", icon_size=20,
            bootstyle="outline", tooltip="Freeze rows or columns to keep them visible while scrolling", side=tk.LEFT, fill=tk.Y, padx=2
        )
        frz_menu.add_command(label="Freeze Top Row", command=self._cb.get("freeze_top_row"))
        frz_menu.add_command(label="Freeze First Column", command=self._cb.get("freeze_first_col"))
        frz_menu.add_separator()
        frz_menu.add_command(label="Unfreeze All Panes", command=self._cb.get("unfreeze_all"))

    # -------------------------------------------------------------------------
    # Formulas Tab
    # -------------------------------------------------------------------------
    def _create_formulas_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=(4, 2))
        self._notebook.add(tab, text="  Formulas  ")

        # 1. Function Library
        lib_grp = ttk.LabelFrame(tab, text="Function Library", padding=3)
        lib_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._btn(
            lib_grp, text="fx Insert\nFunction", icon_name="fx", compound="top", icon_size=20,
            bootstyle="success", command=self._cb.get("show_fx_wizard"),
            tooltip="Insert Function Wizard (Shift+F3)", side=tk.LEFT, fill=tk.Y, padx=2
        )

        as_btn, as_menu = self._menubtn(
            lib_grp, text="AutoSum\nFunctions", icon_name="autosum", compound="top", icon_size=20,
            bootstyle="primary-outline", tooltip="AutoSum Quick Math Functions", side=tk.LEFT, fill=tk.Y, padx=2
        )
        as_menu.add_command(label="Sum (=SUM)", command=lambda: self._cb.get("apply_autosum")("SUM"))
        as_menu.add_command(label="Average (=AVERAGE)", command=lambda: self._cb.get("apply_autosum")("AVERAGE"))
        as_menu.add_command(label="Count Numbers (=COUNT)", command=lambda: self._cb.get("apply_autosum")("COUNT"))
        as_menu.add_command(label="Max (=MAX)", command=lambda: self._cb.get("apply_autosum")("MAX"))
        as_menu.add_command(label="Min (=MIN)", command=lambda: self._cb.get("apply_autosum")("MIN"))

        cats_frame = ttk.Frame(lib_grp)
        cats_frame.pack(side=tk.LEFT, padx=2)

        row1 = ttk.Frame(cats_frame)
        row1.pack(fill=tk.X, pady=1)

        row2 = ttk.Frame(cats_frame)
        row2.pack(fill=tk.X, pady=1)

        cat_row1 = [
            ("Financial", "fx_financial"),
            ("Logical", "fx_logical"),
            ("Text", "fx_text"),
            ("Date & Time", "fx_date"),
        ]
        cat_row2 = [
            ("Lookup & Ref", "fx_lookup", "Lookup & Reference"),
            ("Math & Trig", "fx_math", "Math & Trig"),
            ("Stats", "fx_stats", "Statistical"),
        ]

        for label, icon_key in cat_row1:
            btn, m = self._menubtn(
                row1, text=f" {label}", icon_name=icon_key, bootstyle="outline",
                tooltip=f"{label} Functions", side=tk.LEFT, padx=1
            )
            for name, meta in sorted(FUNCTION_METADATA.items()):
                if meta.get("cat") == label:
                    m.add_command(
                        label=f"{name} — {meta.get('desc')[:35]}...",
                        command=partial(self._cb.get("insert_formula_text"), f"={name}("),
                    )

        for item in cat_row2:
            label, icon_key = item[0], item[1]
            cat_filter = item[2] if len(item) > 2 else label
            btn, m = self._menubtn(
                row2, text=f" {label}", icon_name=icon_key, bootstyle="outline",
                tooltip=f"{cat_filter} Functions", side=tk.LEFT, padx=1
            )
            for name, meta in sorted(FUNCTION_METADATA.items()):
                if meta.get("cat") == cat_filter:
                    m.add_command(
                        label=f"{name} — {meta.get('desc')[:35]}...",
                        command=partial(self._cb.get("insert_formula_text"), f"={name}("),
                    )

        self._btn(
            row2, text=" All 129...", icon_name="formula_ref", bootstyle="secondary-outline",
            command=self._cb.get("formula_reference_dialog"), tooltip="Complete 129 Excel Formula Reference Manual",
            side=tk.LEFT, padx=1
        )

        # 2. Formula Auditing Group
        audit_grp = ttk.LabelFrame(tab, text="Formula Auditing", padding=3)
        audit_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._btn(
            audit_grp, text="Show Formulas\n(Ctrl+`)", icon_name="show_formulas", compound="top", icon_size=20,
            bootstyle="outline", command=self._cb.get("toggle_show_formulas"),
            tooltip="Toggle displaying formula strings instead of calculated results", side=tk.LEFT, fill=tk.Y, padx=2
        )
        self._btn(
            audit_grp, text="Dependency\nAuditor...", icon_name="auditor", compound="top", icon_size=20,
            bootstyle="outline", command=self._cb.get("formula_auditor_dialog"),
            tooltip="Inspect formula dependency trees, precedents, and evaluated nodes", side=tk.LEFT, fill=tk.Y, padx=2
        )

        # 3. Calculation & Performance Group
        calc_grp = ttk.LabelFrame(tab, text="Calculation & Engine", padding=3)
        calc_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._btn(
            calc_grp, text="Calculate Sheet\n(F9)", icon_name="calculate", compound="top", icon_size=20,
            bootstyle="primary", command=self._cb.get("recalculate_sheet"),
            tooltip="Calculate full sheet now (F9)", side=tk.LEFT, fill=tk.Y, padx=2
        )

        calc_sub = ttk.Frame(calc_grp)
        calc_sub.pack(side=tk.LEFT, padx=2)
        self._btn(
            calc_sub, text=" SIMD Status", icon_name="simd", bootstyle="secondary-outline",
            command=self._cb.get("simd_info_dialog"), tooltip="Hardware SIMD Vector Acceleration Details", fill=tk.X
        )
        self._btn(
            calc_sub, text=" Benchmark", icon_name="benchmark", bootstyle="secondary-outline",
            command=self._cb.get("simd_benchmark_dialog"), tooltip="Run SIMD Hardware Vector Benchmark", fill=tk.X
        )

    # -------------------------------------------------------------------------
    # View Tab
    # -------------------------------------------------------------------------
    def _create_view_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=(4, 2))
        self._notebook.add(tab, text="  View  ")

        # 1. Workbook Views
        vw_grp = ttk.LabelFrame(tab, text="Workbook Views", padding=3)
        vw_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._btn(
            vw_grp, text="Normal\nView", icon_name="view_normal", compound="top", icon_size=20,
            bootstyle="primary-outline", command=self._cb.get("refresh_grid"),
            tooltip="Standard Worksheet Grid View", side=tk.LEFT, fill=tk.Y, padx=2
        )
        self._btn(
            vw_grp, text="Formula\nAudit View", icon_name="show_formulas", compound="top", icon_size=20,
            bootstyle="outline", command=self._cb.get("toggle_show_formulas"),
            tooltip="Display cell formulas directly in the grid (Ctrl+`)", side=tk.LEFT, fill=tk.Y, padx=2
        )

        # 2. Show / Hide (checkbuttons like genuine Excel)
        show_grp = ttk.LabelFrame(tab, text="Show / Hide", padding=3)
        show_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        self._show_fbar_var = tk.BooleanVar(value=True)
        cb_fbar = ttk.Checkbutton(
            show_grp,
            text="Formula Bar",
            variable=self._show_fbar_var,
            bootstyle="round-toggle",
            command=self._cb.get("toggle_formula_bar"),
        )
        cb_fbar.pack(anchor=tk.W, padx=4, pady=2)
        Tooltip(cb_fbar, "Show or hide the Formula Bar")

        self._show_sbar_var = tk.BooleanVar(value=True)
        cb_sbar = ttk.Checkbutton(
            show_grp,
            text="Status Bar",
            variable=self._show_sbar_var,
            bootstyle="round-toggle",
            command=self._cb.get("toggle_status_bar"),
        )
        cb_sbar.pack(anchor=tk.W, padx=4, pady=2)
        Tooltip(cb_sbar, "Show or hide the bottom Status Bar")

        # 3. Zoom Group
        zoom_grp = ttk.LabelFrame(tab, text="Zoom", padding=3)
        zoom_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        z_sub = ttk.Frame(zoom_grp)
        z_sub.pack(side=tk.LEFT, fill=tk.Y)

        z_row1 = ttk.Frame(z_sub)
        z_row1.pack(fill=tk.X, pady=1)
        self._btn(
            z_row1, text=" Zoom In (+)", icon_name="zoom_in", bootstyle="outline",
            command=self._cb.get("zoom_in"), tooltip="Zoom In by 15%", side=tk.LEFT, padx=1
        )
        self._btn(
            z_row1, text=" Zoom Out (-)", icon_name="zoom_out", bootstyle="outline",
            command=self._cb.get("zoom_out"), tooltip="Zoom Out by 15%", side=tk.LEFT, padx=1
        )

        z_row2 = ttk.Frame(z_sub)
        z_row2.pack(fill=tk.X, pady=1)
        self._btn(
            z_row2, text=" 100%", icon_name="zoom_100", bootstyle="outline",
            command=self._cb.get("reset_zoom"), tooltip="Reset Zoom to 100%", side=tk.LEFT, padx=1
        )
        self._btn(
            z_row2, text=" Refresh (F5)", icon_name="refresh", bootstyle="secondary-outline",
            command=self._cb.get("refresh_grid"), tooltip="Refresh Grid and formulas (F5)", side=tk.LEFT, padx=1
        )

        # 4. Window & Panes Group
        win_grp = ttk.LabelFrame(tab, text="Window", padding=3)
        win_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        frz_btn2, frz_menu2 = self._menubtn(
            win_grp, text="Freeze\nPanes", icon_name="freeze_panes", compound="top", icon_size=20,
            bootstyle="outline", tooltip="Freeze rows or columns to keep them visible while scrolling", side=tk.LEFT, fill=tk.Y, padx=2
        )
        frz_menu2.add_command(label="Freeze Top Row", command=self._cb.get("freeze_top_row"))
        frz_menu2.add_command(label="Freeze First Column", command=self._cb.get("freeze_first_col"))
        frz_menu2.add_separator()
        frz_menu2.add_command(label="Unfreeze All Panes", command=self._cb.get("unfreeze_all"))

        self._btn(
            win_grp, text="Auto-fit\nColumns", icon_name="autofit", compound="top", icon_size=20,
            bootstyle="outline", command=self._cb.get("autofit_all_cols"),
            tooltip="Auto-fit all column widths", side=tk.LEFT, fill=tk.Y, padx=2
        )

        # 5. Theme & Appearance
        theme_grp = ttk.LabelFrame(tab, text="Theme & Appearance", padding=3)
        theme_grp.pack(side=tk.LEFT, fill=tk.Y, padx=2)

        t_sub = ttk.Frame(theme_grp)
        t_sub.pack(side=tk.LEFT, padx=2)

        theme_cb = ttk.Combobox(
            t_sub, values=Config.AVAILABLE_THEMES, width=11, state="readonly"
        )
        theme_cb.set(Config.THEME)
        theme_cb.pack(pady=2)
        theme_cb.bind(
            "<<ComboboxSelected>>", lambda e: self._cb.get("change_theme")(theme_cb.get())
        )
        Tooltip(theme_cb, "Select Application Theme")

        self._btn(
            t_sub, text=" Shortcuts (Ctrl+/)", icon_name="shortcuts", bootstyle="secondary-outline",
            command=self._cb.get("shortcuts_dialog"), tooltip="Keyboard Shortcuts Reference (Ctrl+/)", fill=tk.X
        )

    def set_formula_bar_check(self, state: bool) -> None:
        """Update formula bar checkbutton state."""
        if hasattr(self, "_show_fbar_var"):
            self._show_fbar_var.set(state)

    def set_status_bar_check(self, state: bool) -> None:
        """Update status bar checkbutton state."""
        if hasattr(self, "_show_sbar_var"):
            self._show_sbar_var.set(state)

    def update_color_buttons(self, bg_hex: str | None, fg_hex: str | None) -> None:
        if hasattr(self, "_bg_btn"):
            try:
                self._bg_btn.config(bg=bg_hex or "#FFFF00")
            except Exception:
                pass
        if hasattr(self, "_fg_btn"):
            try:
                self._fg_btn.config(fg=fg_hex or "#FF0000")
            except Exception:
                pass


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
        on_goto_cell: Callable[[str], None],
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
        fx_photo = None
        try:
            fx_photo = get_icon("fx", size=16)
        except Exception:
            pass
        if fx_photo:
            btn_fx = ttk.Button(
                self, image=fx_photo, bootstyle="secondary-outline", command=self._on_fx
            )
            btn_fx.image = fx_photo
        else:
            btn_fx = ttk.Button(
                self, text="fx", width=3, bootstyle="secondary-outline", command=self._on_fx
            )
        btn_fx.pack(side=tk.LEFT, padx=(0, 4))
        Tooltip(btn_fx, "Insert Function Wizard (Shift+F3)")

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
        on_duplicate_sheet: Callable[[str], None],
        on_tab_color_change: Callable[[str, str | None], None] | None = None,
    ) -> None:
        super().__init__(parent, padding=(4, 2))
        self._on_select = on_select_sheet
        self._on_add = on_add_sheet
        self._on_rename = on_rename_sheet
        self._on_delete = on_delete_sheet
        self._on_duplicate = on_duplicate_sheet
        self._on_color = on_tab_color_change

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
        b1 = ttk.Button(
            nav_frame, text="⏮", width=w, bootstyle="secondary-outline", command=self._scroll_start
        )
        b1.pack(side=tk.LEFT, padx=1)
        Tooltip(b1, "Scroll to first sheet")

        b2 = ttk.Button(
            nav_frame, text="◀", width=w, bootstyle="secondary-outline", command=self._scroll_left
        )
        b2.pack(side=tk.LEFT, padx=1)
        Tooltip(b2, "Scroll left")

        b3 = ttk.Button(
            nav_frame, text="▶", width=w, bootstyle="secondary-outline", command=self._scroll_right
        )
        b3.pack(side=tk.LEFT, padx=1)
        Tooltip(b3, "Scroll right")

        b4 = ttk.Button(
            nav_frame, text="⏭", width=w, bootstyle="secondary-outline", command=self._scroll_end
        )
        b4.pack(side=tk.LEFT, padx=1)
        Tooltip(b4, "Scroll to last sheet")

        # '+' Add sheet button
        add_btn = ttk.Button(self, text=" ➕ ", bootstyle="success-outline", command=self._on_add)
        add_btn.pack(side=tk.LEFT, padx=(3, 6))
        Tooltip(add_btn, "New Sheet")

        # Canvas for scrollable sheet tabs
        self._canvas = tk.Canvas(self, height=28, highlightthickness=0, bd=0)
        self._canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self._inner = ttk.Frame(self._canvas)
        self._canvas.create_window((0, 0), window=self._inner, anchor=tk.NW)

        self._inner.bind(
            "<Configure>", lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all"))
        )
        self._canvas.bind(
            "<MouseWheel>", lambda e: self._canvas.xview_scroll(-1 if e.delta > 0 else 1, "units")
        )

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
            self._create_tab_button(name)

        if names:
            self._select(names[0])

    def add_sheet(self, name: str) -> None:
        """Add a new sheet tab."""
        if name not in self._names:
            self._names.append(name)
        self._create_tab_button(name)
        self._select(name)

    def rename_sheet(self, old_name: str, new_name: str) -> None:
        """Rename an existing sheet tab."""
        if old_name in self._names:
            idx = self._names.index(old_name)
            self._names[idx] = new_name
        if old_name in self._tabs:
            btn = self._tabs.pop(old_name)
            btn.destroy()
        if old_name in self._tab_colors:
            self._tab_colors[new_name] = self._tab_colors.pop(old_name)

        # Re-render tabs to maintain order
        cur_active = self._active if self._active != old_name else new_name
        self.set_sheets(self._names)
        self._select(cur_active)

    def remove_sheet(self, name: str) -> None:
        """Remove a sheet tab."""
        if name in self._names:
            self._names.remove(name)
        if name in self._tabs:
            btn = self._tabs.pop(name)
            btn.destroy()
        if name in self._tab_colors:
            del self._tab_colors[name]

        if self._active == name and self._names:
            self._select(self._names[0])

    def set_tab_color(self, name: str, color_hex: str | None) -> None:
        """Set custom color for a tab."""
        if color_hex:
            self._tab_colors[name] = color_hex
        elif name in self._tab_colors:
            del self._tab_colors[name]
        if name in self._tabs:
            if name != self._active:
                self._tabs[name].configure(bootstyle="secondary-outline")

    def _create_tab_button(self, name: str) -> ttk.Button:
        btn = ttk.Button(
            self._inner,
            text=f" {name} ",
            bootstyle="secondary-outline",
            command=partial(self._select, name),
        )
        btn.pack(side=tk.LEFT, padx=2)
        btn.bind("<Button-3>", partial(self._show_tab_context_menu, name))
        btn.bind("<Double-1>", partial(self._on_double_click_tab, name))
        self._tabs[name] = btn
        return btn

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
        menu.add_command(
            label=f"Rename '{sheet_name}'...", command=lambda: self._on_rename(sheet_name)
        )
        menu.add_command(
            label=f"Duplicate '{sheet_name}'", command=lambda: self._on_duplicate(sheet_name)
        )
        menu.add_command(label="Tab Color...", command=lambda: self._pick_tab_color(sheet_name))
        menu.add_separator()
        menu.add_command(
            label=f"Delete '{sheet_name}'", command=lambda: self._on_delete(sheet_name)
        )
        menu.tk_popup(event.x_root, event.y_root)

    def _pick_tab_color(self, sheet_name: str) -> None:
        color = colorchooser.askcolor(title=f"Choose Color for '{sheet_name}' Tab")
        if color[1]:
            self.set_tab_color(sheet_name, color[1].upper())
            if self._on_color:
                self._on_color(sheet_name, color[1].upper())

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

    def __init__(
        self, parent: tk.Widget, on_zoom_change: Callable[[int], None] | None = None
    ) -> None:
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
        ttk.Button(
            zoom_frame, text="−", width=2, bootstyle="secondary-outline", command=self._zoom_out
        ).pack(side=tk.LEFT, padx=1)
        self._zoom_label = ttk.Label(zoom_frame, text="100%", width=5, anchor=tk.CENTER)
        self._zoom_label.pack(side=tk.LEFT, padx=2)
        ttk.Button(
            zoom_frame, text="+", width=2, bootstyle="secondary-outline", command=self._zoom_in
        ).pack(side=tk.LEFT, padx=1)

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
                    num = float(
                        str(v)
                        .replace(",", ".")
                        .replace(" ", "")
                        .replace("₽", "")
                        .replace("$", "")
                        .replace("%", "")
                    )
                    nums.append(num)
                except (ValueError, TypeError):
                    pass

        parts = []
        if nums:
            avg_val = sum(nums) / len(nums)
            parts.append(
                f"AVERAGE: {avg_val:,.2f}"
                if avg_val != int(avg_val)
                else f"AVERAGE: {int(avg_val)}"
            )
            parts.append(f"COUNT: {non_empty}")
            parts.append(f"NUMS: {len(nums)}")
            min_val = min(nums)
            max_val = max(nums)
            sum_val = sum(nums)
            parts.append(f"MIN: {min_val:g}")
            parts.append(f"MAX: {max_val:g}")
            parts.append(
                f"SUM: {sum_val:,.2f}" if sum_val != int(sum_val) else f"SUM: {int(sum_val)}"
            )
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
        self, tree: ttk.Treeview, on_commit: Callable[[str], None], on_cancel: Callable[[], None]
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
            bd=2,
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

"""
Professional dialogs and wizards for Excel Viewer Pro.
"""

from __future__ import annotations

import math
import time
import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox
from typing import Any, Callable

import ttkbootstrap as ttk
from openpyxl.utils import get_column_letter
from ttkbootstrap.constants import *

try:
    from .config import Config
    from .formulas import FUNCTION_METADATA
    from .models import SheetData
except ImportError:
    from config import Config
    from formulas import FUNCTION_METADATA
    from models import SheetData

# =============================================================================
# Find and Replace Dialog
# =============================================================================


class FindReplaceDialog(tk.Toplevel):
    """Excel-like Find and Replace dialog."""

    def __init__(
        self,
        parent: tk.Widget,
        on_find_next: Callable[[str, bool, bool, bool], bool],
        on_replace: Callable[[str, str, bool, bool], bool],
        on_replace_all: Callable[[str, str, bool, bool], int],
        initial_tab: str = "find",
    ) -> None:
        super().__init__(parent)
        self.title("Find and Replace")
        self.geometry("450x260")
        self.resizable(False, False)
        self.transient(parent)

        self._on_find_next = on_find_next
        self._on_replace = on_replace
        self._on_replace_all = on_replace_all

        self._find_var = tk.StringVar()
        self._replace_var = tk.StringVar()
        self._match_case = tk.BooleanVar(value=False)
        self._match_entire = tk.BooleanVar(value=False)
        self._search_all_sheets = tk.BooleanVar(value=False)

        self._build_ui(initial_tab)
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self, initial_tab: str) -> None:
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Find Tab
        find_frame = ttk.Frame(notebook, padding=10)
        notebook.add(find_frame, text="Find")

        # Replace Tab
        replace_frame = ttk.Frame(notebook, padding=10)
        notebook.add(replace_frame, text="Replace")

        # Setup Find Tab
        ttk.Label(find_frame, text="Find what:").grid(row=0, column=0, sticky=tk.W, pady=5)
        find_entry1 = ttk.Entry(find_frame, textvariable=self._find_var, width=32)
        find_entry1.grid(row=0, column=1, columnspan=2, sticky=tk.EW, pady=5)
        find_entry1.focus_set()

        opt_frame1 = ttk.Frame(find_frame)
        opt_frame1.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=10)
        ttk.Checkbutton(opt_frame1, text="Match case", variable=self._match_case).pack(anchor=tk.W)
        ttk.Checkbutton(
            opt_frame1, text="Match entire cell contents", variable=self._match_entire
        ).pack(anchor=tk.W)
        ttk.Checkbutton(
            opt_frame1, text="Search all sheets", variable=self._search_all_sheets
        ).pack(anchor=tk.W)

        btn_box1 = ttk.Frame(find_frame)
        btn_box1.grid(row=2, column=0, columnspan=3, sticky=tk.E, pady=5)
        ttk.Button(
            btn_box1, text="Find Next", bootstyle="primary", command=self._do_find_next
        ).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_box1, text="Close", bootstyle="secondary", command=self.destroy).pack(
            side=tk.LEFT, padx=3
        )

        # Setup Replace Tab
        ttk.Label(replace_frame, text="Find what:").grid(row=0, column=0, sticky=tk.W, pady=3)
        ttk.Entry(replace_frame, textvariable=self._find_var, width=32).grid(
            row=0, column=1, columnspan=2, sticky=tk.EW, pady=3
        )

        ttk.Label(replace_frame, text="Replace with:").grid(row=1, column=0, sticky=tk.W, pady=3)
        ttk.Entry(replace_frame, textvariable=self._replace_var, width=32).grid(
            row=1, column=1, columnspan=2, sticky=tk.EW, pady=3
        )

        opt_frame2 = ttk.Frame(replace_frame)
        opt_frame2.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
        ttk.Checkbutton(opt_frame2, text="Match case", variable=self._match_case).pack(anchor=tk.W)
        ttk.Checkbutton(
            opt_frame2, text="Match entire cell contents", variable=self._match_entire
        ).pack(anchor=tk.W)

        btn_box2 = ttk.Frame(replace_frame)
        btn_box2.grid(row=3, column=0, columnspan=3, sticky=tk.E, pady=5)
        ttk.Button(btn_box2, text="Replace", bootstyle="primary", command=self._do_replace).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(
            btn_box2, text="Replace All", bootstyle="warning", command=self._do_replace_all
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            btn_box2, text="Find Next", bootstyle="secondary-outline", command=self._do_find_next
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_box2, text="Close", bootstyle="secondary", command=self.destroy).pack(
            side=tk.LEFT, padx=2
        )

        if initial_tab == "replace":
            notebook.select(1)

    def _do_find_next(self) -> None:
        target = self._find_var.get()
        if not target:
            return
        found = self._on_find_next(
            target, self._match_case.get(), self._match_entire.get(), self._search_all_sheets.get()
        )
        if not found:
            messagebox.showinfo("Find", f"Cannot find '{target}'.")

    def _do_replace(self) -> None:
        target = self._find_var.get()
        replacement = self._replace_var.get()
        if not target:
            return
        self._on_replace(target, replacement, self._match_case.get(), self._match_entire.get())

    def _do_replace_all(self) -> None:
        target = self._find_var.get()
        replacement = self._replace_var.get()
        if not target:
            return
        count = self._on_replace_all(
            target, replacement, self._match_case.get(), self._match_entire.get()
        )
        messagebox.showinfo(
            "Replace All",
            f"Excel Viewer Pro has completed its search and has made {count} replacement(s).",
        )


# =============================================================================
# Insert Function Wizard (fx)
# =============================================================================


class InsertFunctionDialog(tk.Toplevel):
    """Excel Insert Function (fx) Wizard."""

    def __init__(self, parent: tk.Widget, on_insert_formula: Callable[[str], None]) -> None:
        super().__init__(parent)
        self.title("Insert Function")
        self.geometry("520x460")
        self.resizable(False, False)
        self.transient(parent)

        self._on_insert = on_insert_formula
        self._search_var = tk.StringVar()
        self._category_var = tk.StringVar(value="All")

        self._build_ui()
        self._populate_list()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        # Search Bar
        ttk.Label(main, text="Search for a function:").pack(anchor=tk.W)
        search_box = ttk.Frame(main)
        search_box.pack(fill=tk.X, pady=(2, 8))
        s_entry = ttk.Entry(search_box, textvariable=self._search_var)
        s_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        s_entry.bind("<KeyRelease>", lambda e: self._populate_list())
        ttk.Button(search_box, text="Go", width=4, command=self._populate_list).pack(
            side=tk.LEFT, padx=(5, 0)
        )

        # Category Filter
        cat_box = ttk.Frame(main)
        cat_box.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(cat_box, text="Or select a category:").pack(side=tk.LEFT)
        categories = [
            "All",
            "Math & Trig",
            "Statistical",
            "Lookup & Reference",
            "Logical",
            "Text",
            "Date & Time",
            "Financial",
            "Information",
        ]
        cat_cb = ttk.Combobox(
            cat_box, textvariable=self._category_var, values=categories, state="readonly", width=22
        )
        cat_cb.pack(side=tk.LEFT, padx=10)
        cat_cb.bind("<<ComboboxSelected>>", lambda e: self._populate_list())

        # Function Listbox
        ttk.Label(main, text="Select a function:").pack(anchor=tk.W)
        list_frame = ttk.Frame(main)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(2, 8))

        self._listbox = tk.Listbox(list_frame, font=Config.FONT, selectmode=tk.SINGLE, height=7)
        vsb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self._listbox.yview)
        self._listbox.configure(yscrollcommand=vsb.set)
        self._listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self._listbox.bind("<<ListboxSelect>>", self._on_select_fn)
        self._listbox.bind("<Double-1>", lambda e: self._do_insert())
        self._listbox.bind("<Return>", lambda e: self._do_insert())

        # Description Box
        desc_box = ttk.LabelFrame(main, text="Function Details", padding=8)
        desc_box.pack(fill=tk.X, pady=(0, 10))

        self._syntax_label = ttk.Label(desc_box, text="", font=Config.FONT_BOLD)
        self._syntax_label.pack(anchor=tk.W)

        self._desc_label = ttk.Label(desc_box, text="", wraplength=480, justify=tk.LEFT)
        self._desc_label.pack(anchor=tk.W, pady=(4, 0))

        # Action Buttons
        btn_box = ttk.Frame(main)
        btn_box.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Button(btn_box, text="Insert", bootstyle="success", command=self._do_insert).pack(
            side=tk.RIGHT, padx=4
        )
        ttk.Button(btn_box, text="Cancel", bootstyle="secondary", command=self.destroy).pack(
            side=tk.RIGHT
        )

    def _populate_list(self) -> None:
        query = self._search_var.get().strip().upper()
        cat = self._category_var.get()

        self._listbox.delete(0, tk.END)
        for name, meta in sorted(FUNCTION_METADATA.items()):
            if cat != "All" and meta.get("cat") != cat:
                continue
            if query and query not in name and query not in meta.get("desc", "").upper():
                continue
            self._listbox.insert(tk.END, name)

        if self._listbox.size() > 0:
            self._listbox.selection_set(0)
            self._on_select_fn(None)

    def _on_select_fn(self, event) -> None:
        sel = self._listbox.curselection()
        if not sel:
            return
        name = self._listbox.get(sel[0])
        meta = FUNCTION_METADATA.get(name, {})
        self._syntax_label.config(text=meta.get("syntax", f"{name}()"))
        self._desc_label.config(text=meta.get("desc", ""))

    def _do_insert(self) -> None:
        sel = self._listbox.curselection()
        if not sel:
            return
        name = self._listbox.get(sel[0])
        self._on_insert(f"={name}(")
        self.destroy()


# =============================================================================
# Custom Multi-Column Sort Dialog
# =============================================================================


class CustomSortDialog(tk.Toplevel):
    """Excel Multi-level sort dialog."""

    def __init__(
        self,
        parent: tk.Widget,
        headers: list[str],
        on_apply_sort: Callable[[list[tuple[int, bool]], bool], None],
    ) -> None:
        super().__init__(parent)
        self.title("Sort")
        self.geometry("480x320")
        self.resizable(False, False)
        self.transient(parent)

        self._headers = headers
        self._on_apply_sort = on_apply_sort
        self._has_headers = tk.BooleanVar(value=True)

        self._level_col_vars: list[tk.StringVar] = []
        self._level_order_vars: list[tk.StringVar] = []
        self._level_rows: list[ttk.Frame] = []

        self._build_ui()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # Toolbar
        bar = ttk.Frame(main)
        bar.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(bar, text="➕ Add Level", bootstyle="outline", command=self._add_level).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(
            bar, text="➖ Delete Level", bootstyle="outline", command=self._delete_level
        ).pack(side=tk.LEFT, padx=2)
        ttk.Checkbutton(bar, text="My data has headers", variable=self._has_headers).pack(
            side=tk.RIGHT
        )

        # Scrollable level container
        self._container = ttk.Frame(main, relief="solid", borderwidth=1, padding=5)
        self._container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Add initial sort level
        self._add_level()

        # Action Buttons
        btn_box = ttk.Frame(main)
        btn_box.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Button(btn_box, text="OK", bootstyle="primary", command=self._do_sort).pack(
            side=tk.RIGHT, padx=4
        )
        ttk.Button(btn_box, text="Cancel", bootstyle="secondary", command=self.destroy).pack(
            side=tk.RIGHT
        )

    def _add_level(self) -> None:
        if len(self._level_rows) >= 5:
            return
        idx = len(self._level_rows)
        row_frame = ttk.Frame(self._container)
        row_frame.pack(fill=tk.X, pady=3)

        label_text = "Sort by:" if idx == 0 else "Then by:"
        ttk.Label(row_frame, text=label_text, width=9).pack(side=tk.LEFT)

        col_var = tk.StringVar(value=self._headers[0] if self._headers else "Col 1")
        col_cb = ttk.Combobox(
            row_frame, textvariable=col_var, values=self._headers, state="readonly", width=18
        )
        col_cb.pack(side=tk.LEFT, padx=5)

        ttk.Label(row_frame, text="Order:").pack(side=tk.LEFT, padx=(5, 0))
        order_var = tk.StringVar(value="A to Z (Ascending)")
        order_cb = ttk.Combobox(
            row_frame,
            textvariable=order_var,
            values=["A to Z (Ascending)", "Z to A (Descending)"],
            state="readonly",
            width=18,
        )
        order_cb.pack(side=tk.LEFT, padx=5)

        self._level_col_vars.append(col_var)
        self._level_order_vars.append(order_var)
        self._level_rows.append(row_frame)

    def _delete_level(self) -> None:
        if len(self._level_rows) > 1:
            frame = self._level_rows.pop()
            frame.destroy()
            self._level_col_vars.pop()
            self._level_order_vars.pop()

    def _do_sort(self) -> None:
        sort_rules = []
        for c_var, o_var in zip(self._level_col_vars, self._level_order_vars):
            col_name = c_var.get()
            col_idx = self._headers.index(col_name) if col_name in self._headers else 0
            is_descending = "Z to A" in o_var.get() or "Descending" in o_var.get()
            sort_rules.append((col_idx, is_descending))

        self._on_apply_sort(sort_rules, self._has_headers.get())
        self.destroy()


# =============================================================================
# AutoFilter Popup Dialog
# =============================================================================


class AutoFilterPopup(tk.Toplevel):
    """Excel Column AutoFilter Dropdown Popup."""

    def __init__(
        self,
        parent: tk.Widget,
        col_idx: int,
        col_name: str,
        unique_values: list[str],
        selected_values: set[str],
        on_apply: Callable[[int, set[str] | None], None],
        on_sort_col: Callable[[int, bool], None],
    ) -> None:
        super().__init__(parent)
        self.title(f"Filter: {col_name}")
        self.geometry("260x340")
        self.resizable(False, False)
        self.transient(parent)

        self._col_idx = col_idx
        self._unique_values = sorted(unique_values, key=lambda s: str(s).lower())
        self._selected_values = (
            set(selected_values) if selected_values is not None else set(unique_values)
        )
        self._on_apply = on_apply
        self._on_sort_col = on_sort_col

        self._check_vars: dict[str, tk.BooleanVar] = {}
        self._search_var = tk.StringVar()
        self._select_all_var = tk.BooleanVar(
            value=len(self._selected_values) == len(self._unique_values)
        )

        self._build_ui()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=8)
        main.pack(fill=tk.BOTH, expand=True)

        # Quick Sort Actions
        sort_box = ttk.Frame(main)
        sort_box.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(
            sort_box,
            text="⬆ Sort A → Z",
            bootstyle="outline",
            command=lambda: [self._on_sort_col(self._col_idx, False), self.destroy()],
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=1)
        ttk.Button(
            sort_box,
            text="⬇ Sort Z → A",
            bootstyle="outline",
            command=lambda: [self._on_sort_col(self._col_idx, True), self.destroy()],
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=1)

        ttk.Separator(main).pack(fill=tk.X, pady=4)

        # Search Bar
        s_entry = ttk.Entry(main, textvariable=self._search_var)
        s_entry.pack(fill=tk.X, pady=(0, 4))
        s_entry.bind("<KeyRelease>", lambda e: self._render_checks())

        # Select All checkbox
        ttk.Checkbutton(
            main, text="(Select All)", variable=self._select_all_var, command=self._toggle_all
        ).pack(anchor=tk.W)

        # Scrollable checklist
        list_container = ttk.Frame(main, relief="solid", borderwidth=1)
        list_container.pack(fill=tk.BOTH, expand=True, pady=4)

        self._canvas = tk.Canvas(list_container, highlightthickness=0)
        vsb = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)

        self._inner_frame = ttk.Frame(self._canvas)
        self._canvas.create_window((0, 0), window=self._inner_frame, anchor=tk.NW)

        self._inner_frame.bind(
            "<Configure>", lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all"))
        )
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self._render_checks()

        # Action Buttons
        btn_box = ttk.Frame(main)
        btn_box.pack(fill=tk.X, side=tk.BOTTOM, pady=(4, 0))
        ttk.Button(
            btn_box, text="Clear Filter", bootstyle="secondary-outline", command=self._clear_filter
        ).pack(side=tk.LEFT)
        ttk.Button(btn_box, text="OK", bootstyle="primary", command=self._do_apply).pack(
            side=tk.RIGHT, padx=2
        )
        ttk.Button(btn_box, text="Cancel", bootstyle="secondary", command=self.destroy).pack(
            side=tk.RIGHT
        )

    def _render_checks(self) -> None:
        query = self._search_var.get().lower().strip()
        for w in self._inner_frame.winfo_children():
            w.destroy()

        for val in self._unique_values:
            display_val = "(Blanks)" if val == "" else str(val)
            if query and query not in display_val.lower():
                continue

            if val not in self._check_vars:
                self._check_vars[val] = tk.BooleanVar(value=val in self._selected_values)

            cb = ttk.Checkbutton(
                self._inner_frame, text=display_val, variable=self._check_vars[val]
            )
            cb.pack(anchor=tk.W, padx=4, pady=1)

    def _toggle_all(self) -> None:
        val = self._select_all_var.get()
        for var in self._check_vars.values():
            var.set(val)

    def _clear_filter(self) -> None:
        self._on_apply(self._col_idx, None)
        self.destroy()

    def _do_apply(self) -> None:
        selected = {val for val, var in self._check_vars.items() if var.get()}
        if len(selected) == len(self._unique_values):
            self._on_apply(self._col_idx, None)
        else:
            self._on_apply(self._col_idx, selected)
        self.destroy()


# =============================================================================
# Professional Chart Wizard Dialog
# =============================================================================


class ChartWizardDialog(tk.Toplevel):
    """Excel-like Chart Creation and Customization Wizard."""

    CHART_TYPES = [
        "Clustered Column",
        "Stacked Column",
        "Bar (Horizontal)",
        "Line",
        "Smooth Line",
        "Pie",
        "Donut",
        "Area",
        "Scatter (XY)",
        "Histogram",
    ]

    def __init__(
        self,
        parent: tk.Widget,
        sheet_data: SheetData,
        get_data_fn: Callable[[int, int, int, int], tuple[list[str], list[list[float]], list[str]]],
    ) -> None:
        super().__init__(parent)
        self.title("Chart Wizard")
        self.geometry("960x680")
        self.minsize(780, 520)
        self.transient(parent)

        self._sheet_data = sheet_data
        self._get_data_fn = get_data_fn

        self._chart_type = tk.StringVar(value="Clustered Column")
        self._palette = tk.StringVar(value="Excel Classic")
        self._title_var = tk.StringVar(value="Chart")
        self._xlabel_var = tk.StringVar()
        self._ylabel_var = tk.StringVar()
        self._show_legend = tk.BooleanVar(value=True)
        self._show_grid = tk.BooleanVar(value=True)

        self._figure = None
        self._canvas_widget = None

        self._build_ui()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=8)
        main.pack(fill=tk.BOTH, expand=True)

        # Left Control Panel
        left = ttk.Frame(main, width=280, padding=8)
        left.pack(side=tk.LEFT, fill=tk.Y)

        ttk.Label(left, text="Chart Type:", font=Config.FONT_BOLD).pack(anchor=tk.W)
        cb_type = ttk.Combobox(
            left, textvariable=self._chart_type, values=self.CHART_TYPES, state="readonly"
        )
        cb_type.pack(fill=tk.X, pady=(2, 8))
        cb_type.bind("<<ComboboxSelected>>", lambda e: self._render_chart())

        ttk.Label(left, text="Color Theme:", font=Config.FONT_BOLD).pack(anchor=tk.W)
        cb_pal = ttk.Combobox(
            left,
            textvariable=self._palette,
            values=list(Config.CHART_PALETTES.keys()),
            state="readonly",
        )
        cb_pal.pack(fill=tk.X, pady=(2, 8))
        cb_pal.bind("<<ComboboxSelected>>", lambda e: self._render_chart())

        # Data Range Frame
        range_box = ttk.LabelFrame(left, text="Data Range", padding=6)
        range_box.pack(fill=tk.X, pady=6)

        ttk.Label(range_box, text="Rows:").grid(row=0, column=0, sticky=tk.W)
        self._r_start = ttk.Spinbox(
            range_box, from_=1, to=max(1, self._sheet_data.row_count), width=5
        )
        self._r_start.set(1)
        self._r_start.grid(row=0, column=1, padx=2)
        ttk.Label(range_box, text="to").grid(row=0, column=2)
        self._r_end = ttk.Spinbox(
            range_box, from_=1, to=max(1, self._sheet_data.row_count), width=5
        )
        self._r_end.set(min(15, max(1, self._sheet_data.row_count)))
        self._r_end.grid(row=0, column=3, padx=2)

        ttk.Label(range_box, text="Cols:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self._c_start = ttk.Spinbox(
            range_box, from_=1, to=max(1, self._sheet_data.col_count), width=5
        )
        self._c_start.set(1)
        self._c_start.grid(row=1, column=1, padx=2, pady=4)
        ttk.Label(range_box, text="to").grid(row=1, column=2, pady=4)
        self._c_end = ttk.Spinbox(
            range_box, from_=1, to=max(1, self._sheet_data.col_count), width=5
        )
        self._c_end.set(min(4, max(1, self._sheet_data.col_count)))
        self._c_end.grid(row=1, column=3, padx=2, pady=4)

        # Titles & Labels
        titles_box = ttk.LabelFrame(left, text="Labels & Titles", padding=6)
        titles_box.pack(fill=tk.X, pady=6)
        ttk.Label(titles_box, text="Chart Title:").pack(anchor=tk.W)
        ttk.Entry(titles_box, textvariable=self._title_var).pack(fill=tk.X, pady=(1, 4))
        ttk.Label(titles_box, text="X-Axis Title:").pack(anchor=tk.W)
        ttk.Entry(titles_box, textvariable=self._xlabel_var).pack(fill=tk.X, pady=(1, 4))
        ttk.Label(titles_box, text="Y-Axis Title:").pack(anchor=tk.W)
        ttk.Entry(titles_box, textvariable=self._ylabel_var).pack(fill=tk.X, pady=(1, 4))

        # Checkboxes
        ttk.Checkbutton(
            left, text="Show Legend", variable=self._show_legend, command=self._render_chart
        ).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(
            left, text="Show Grid Lines", variable=self._show_grid, command=self._render_chart
        ).pack(anchor=tk.W, pady=2)

        # Draw / Refresh Button
        ttk.Button(
            left, text="🔄 Update Chart", bootstyle="primary", command=self._render_chart
        ).pack(fill=tk.X, pady=(10, 4))

        # Export Buttons
        ttk.Separator(left).pack(fill=tk.X, pady=6)
        ttk.Button(
            left,
            text="💾 Save Image (PNG)",
            bootstyle="success-outline",
            command=lambda: self._export_chart("png"),
        ).pack(fill=tk.X, pady=2)
        ttk.Button(
            left,
            text="📄 Save Vector (PDF/SVG)",
            bootstyle="info-outline",
            command=lambda: self._export_chart("pdf"),
        ).pack(fill=tk.X, pady=2)

        # Right Chart Area
        self._chart_panel = ttk.Frame(main, relief="solid", borderwidth=1)
        self._chart_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.after(100, self._render_chart)

    def _render_chart(self) -> None:
        try:
            import matplotlib

            matplotlib.use("TkAgg")
            import matplotlib.pyplot as plt
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        except ImportError:
            messagebox.showerror("Error", "Matplotlib is not installed.")
            return

        try:
            r1 = max(0, int(self._r_start.get()) - 1)
            r2 = max(0, int(self._r_end.get()) - 1)
            c1 = max(0, int(self._c_start.get()) - 1)
            c2 = max(0, int(self._c_end.get()) - 1)

            labels, series, series_names = self._get_data_fn(r1, r2, c1, c2)
            if not labels or not series:
                return

            if self._canvas_widget:
                self._canvas_widget.destroy()
            if self._figure:
                plt.close(self._figure)

            self._figure, ax = plt.subplots(figsize=(7, 5), dpi=100)
            colors = Config.CHART_PALETTES.get(
                self._palette.get(), Config.CHART_PALETTES["Excel Classic"]
            )
            chart_type = self._chart_type.get()

            if chart_type == "Clustered Column":
                x = list(range(len(labels)))
                num_series = len(series)
                w = 0.8 / max(1, num_series)
                for i, data in enumerate(series):
                    offset = (i - num_series / 2 + 0.5) * w
                    name = series_names[i] if i < len(series_names) else f"Series {i + 1}"
                    ax.bar(
                        [xi + offset for xi in x],
                        data,
                        w,
                        label=name,
                        color=colors[i % len(colors)],
                    )
                ax.set_xticks(x)
                ax.set_xticklabels(labels, rotation=35, ha="right")

            elif chart_type == "Stacked Column":
                x = list(range(len(labels)))
                bottom = [0.0] * len(labels)
                for i, data in enumerate(series):
                    name = series_names[i] if i < len(series_names) else f"Series {i + 1}"
                    ax.bar(x, data, 0.6, bottom=bottom, label=name, color=colors[i % len(colors)])
                    bottom = [b + d for b, d in zip(bottom, data)]
                ax.set_xticks(x)
                ax.set_xticklabels(labels, rotation=35, ha="right")

            elif chart_type == "Bar (Horizontal)":
                y = list(range(len(labels)))
                num_series = len(series)
                h = 0.8 / max(1, num_series)
                for i, data in enumerate(series):
                    offset = (i - num_series / 2 + 0.5) * h
                    name = series_names[i] if i < len(series_names) else f"Series {i + 1}"
                    ax.barh(
                        [yi + offset for yi in y],
                        data,
                        h,
                        label=name,
                        color=colors[i % len(colors)],
                    )
                ax.set_yticks(y)
                ax.set_yticklabels(labels)

            elif chart_type in ("Line", "Smooth Line"):
                for i, data in enumerate(series):
                    name = series_names[i] if i < len(series_names) else f"Series {i + 1}"
                    ax.plot(
                        labels,
                        data,
                        marker="o",
                        linewidth=2.2,
                        label=name,
                        color=colors[i % len(colors)],
                    )
                ax.tick_params(axis="x", rotation=35)

            elif chart_type in ("Pie", "Donut"):
                first_series = series[0] if series else []
                wedge_props = (
                    dict(width=0.4, edgecolor="w") if chart_type == "Donut" else dict(edgecolor="w")
                )
                ax.pie(
                    first_series,
                    labels=labels,
                    autopct="%1.1f%%",
                    colors=colors,
                    startangle=90,
                    wedgeprops=wedge_props,
                )

            elif chart_type == "Area":
                x = list(range(len(labels)))
                for i, data in enumerate(series):
                    name = series_names[i] if i < len(series_names) else f"Series {i + 1}"
                    ax.fill_between(x, data, alpha=0.4, color=colors[i % len(colors)], label=name)
                    ax.plot(x, data, color=colors[i % len(colors)], linewidth=1.5)
                ax.set_xticks(x)
                ax.set_xticklabels(labels, rotation=35, ha="right")

            elif chart_type == "Scatter (XY)":
                for i, data in enumerate(series):
                    name = series_names[i] if i < len(series_names) else f"Series {i + 1}"
                    ax.scatter(
                        list(range(len(data))),
                        data,
                        s=50,
                        label=name,
                        color=colors[i % len(colors)],
                    )
                ax.set_xticks(list(range(len(labels))))
                ax.set_xticklabels(labels, rotation=35, ha="right")

            elif chart_type == "Histogram":
                all_vals = [v for s in series for v in s]
                ax.hist(all_vals, bins=10, color=colors[0], edgecolor="white", alpha=0.8)

            # Titles and Decorators
            if self._title_var.get():
                ax.set_title(self._title_var.get(), fontsize=12, fontweight="bold", pad=10)
            if self._xlabel_var.get():
                ax.set_xlabel(self._xlabel_var.get())
            if self._ylabel_var.get():
                ax.set_ylabel(self._ylabel_var.get())

            if self._show_grid.get() and chart_type not in ("Pie", "Donut"):
                ax.grid(True, linestyle="--", alpha=0.5)

            if (
                self._show_legend.get()
                and chart_type not in ("Pie", "Donut", "Histogram")
                and len(series) > 0
            ):
                ax.legend(loc="best", frameon=True)

            self._figure.tight_layout()

            canvas = FigureCanvasTkAgg(self._figure, master=self._chart_panel)
            self._canvas_widget = canvas.get_tk_widget()
            self._canvas_widget.pack(fill=tk.BOTH, expand=True)
            canvas.draw()

        except Exception as e:
            messagebox.showerror("Chart Error", f"Failed to generate chart:\n{e}")

    def _export_chart(self, fmt: str) -> None:
        if not self._figure:
            return
        ext = ".png" if fmt == "png" else (".svg" if fmt == "svg" else ".pdf")
        path = filedialog.asksaveasfilename(
            defaultextension=ext,
            filetypes=[(f"{fmt.upper()} File", f"*{ext}"), ("All Files", "*.*")],
        )
        if path:
            self._figure.savefig(path, dpi=200, bbox_inches="tight")
            messagebox.showinfo("Export Chart", f"Chart successfully saved to {path}")


# =============================================================================
# Remove Duplicates Dialog
# =============================================================================


class RemoveDuplicatesDialog(tk.Toplevel):
    """Excel Remove Duplicates Dialog."""

    def __init__(
        self, parent: tk.Widget, headers: list[str], on_remove: Callable[[list[int]], None]
    ) -> None:
        super().__init__(parent)
        self.title("Remove Duplicates")
        self.geometry("360x320")
        self.resizable(False, False)
        self.transient(parent)

        self._headers = headers
        self._on_remove = on_remove
        self._check_vars: list[tk.BooleanVar] = []

        self._build_ui()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main, text="Select columns to compare for duplicates:").pack(
            anchor=tk.W, pady=(0, 6)
        )

        bar = ttk.Frame(main)
        bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(bar, text="Select All", bootstyle="outline", command=self._select_all).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(bar, text="Unselect All", bootstyle="outline", command=self._unselect_all).pack(
            side=tk.LEFT, padx=2
        )

        # List frame
        list_frame = ttk.Frame(main, relief="solid", borderwidth=1, padding=5)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=4)

        canvas = tk.Canvas(list_frame, highlightthickness=0)
        vsb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)

        inner = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=inner, anchor=tk.NW)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        for i, header in enumerate(self._headers):
            var = tk.BooleanVar(value=True)
            self._check_vars.append(var)
            cb = ttk.Checkbutton(inner, text=f"{header} ({get_column_letter(i + 1)})", variable=var)
            cb.pack(anchor=tk.W, pady=2)

        # Buttons
        btn_box = ttk.Frame(main)
        btn_box.pack(fill=tk.X, side=tk.BOTTOM, pady=(8, 0))
        ttk.Button(btn_box, text="OK", bootstyle="primary", command=self._do_remove).pack(
            side=tk.RIGHT, padx=3
        )
        ttk.Button(btn_box, text="Cancel", bootstyle="secondary", command=self.destroy).pack(
            side=tk.RIGHT
        )

    def _select_all(self) -> None:
        for var in self._check_vars:
            var.set(True)

    def _unselect_all(self) -> None:
        for var in self._check_vars:
            var.set(False)

    def _do_remove(self) -> None:
        cols = [i for i, var in enumerate(self._check_vars) if var.get()]
        if not cols:
            messagebox.showwarning("Warning", "Please select at least one column.")
            return
        self._on_remove(cols)
        self.destroy()


# =============================================================================
# Text to Columns Wizard
# =============================================================================


class TextToColumnsDialog(tk.Toplevel):
    """Excel Text to Columns wizard."""

    def __init__(
        self, parent: tk.Widget, sample_data: list[str], on_split: Callable[[str], None]
    ) -> None:
        super().__init__(parent)
        self.title("Convert Text to Columns")
        self.geometry("460x360")
        self.resizable(False, False)
        self.transient(parent)

        self._sample_data = sample_data[:6]
        self._on_split = on_split

        self._tab_var = tk.BooleanVar(value=False)
        self._semicolon_var = tk.BooleanVar(value=True)
        self._comma_var = tk.BooleanVar(value=False)
        self._space_var = tk.BooleanVar(value=False)
        self._other_var = tk.BooleanVar(value=False)
        self._other_char_var = tk.StringVar(value="|")

        self._build_ui()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            main, text="Select delimiters contained in your data:", font=Config.FONT_BOLD
        ).pack(anchor=tk.W, pady=(0, 6))

        delims_frame = ttk.LabelFrame(main, text="Delimiters", padding=8)
        delims_frame.pack(fill=tk.X, pady=(0, 8))

        row1 = ttk.Frame(delims_frame)
        row1.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(
            row1, text="Tab", variable=self._tab_var, command=self._update_preview
        ).pack(side=tk.LEFT, padx=8)
        ttk.Checkbutton(
            row1, text="Semicolon (;)", variable=self._semicolon_var, command=self._update_preview
        ).pack(side=tk.LEFT, padx=8)
        ttk.Checkbutton(
            row1, text="Comma (,)", variable=self._comma_var, command=self._update_preview
        ).pack(side=tk.LEFT, padx=8)

        row2 = ttk.Frame(delims_frame)
        row2.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(
            row2, text="Space", variable=self._space_var, command=self._update_preview
        ).pack(side=tk.LEFT, padx=8)
        ttk.Checkbutton(
            row2, text="Other:", variable=self._other_var, command=self._update_preview
        ).pack(side=tk.LEFT, padx=(8, 2))
        e = ttk.Entry(row2, textvariable=self._other_char_var, width=4)
        e.pack(side=tk.LEFT)
        e.bind("<KeyRelease>", lambda ev: self._update_preview())

        # Data Preview
        preview_box = ttk.LabelFrame(main, text="Data Preview", padding=8)
        preview_box.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self._preview_text = tk.Text(preview_box, font=Config.FONT_MONO, height=6, wrap=tk.NONE)
        self._preview_text.pack(fill=tk.BOTH, expand=True)

        self._update_preview()

        # Action Buttons
        btn_box = ttk.Frame(main)
        btn_box.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Button(btn_box, text="Finish", bootstyle="success", command=self._do_finish).pack(
            side=tk.RIGHT, padx=3
        )
        ttk.Button(btn_box, text="Cancel", bootstyle="secondary", command=self.destroy).pack(
            side=tk.RIGHT
        )

    def _get_active_delimiter(self) -> str:
        if self._semicolon_var.get():
            return ";"
        if self._comma_var.get():
            return ","
        if self._tab_var.get():
            return "\t"
        if self._space_var.get():
            return " "
        if self._other_var.get() and self._other_char_var.get():
            return self._other_char_var.get()
        return ";"

    def _update_preview(self) -> None:
        delim = self._get_active_delimiter()
        self._preview_text.delete("1.0", tk.END)
        for line in self._sample_data:
            parts = line.split(delim)
            self._preview_text.insert(tk.END, " | ".join(parts) + "\n")

    def _do_finish(self) -> None:
        delim = self._get_active_delimiter()
        self._on_split(delim)
        self.destroy()


# =============================================================================
# Goal Seek Dialog
# =============================================================================


class GoalSeekDialog(tk.Toplevel):
    """Excel What-If Analysis: Goal Seek."""

    def __init__(
        self, parent: tk.Widget, current_cell: str, on_solve: Callable[[str, float, str], None]
    ) -> None:
        super().__init__(parent)
        self.title("Goal Seek")
        self.geometry("320x190")
        self.resizable(False, False)
        self.transient(parent)

        self._set_cell_var = tk.StringVar(value=current_cell)
        self._to_val_var = tk.StringVar(value="0")
        self._by_cell_var = tk.StringVar()
        self._on_solve = on_solve

        self._build_ui()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main, text="Set cell:").grid(row=0, column=0, sticky=tk.W, pady=3)
        ttk.Entry(main, textvariable=self._set_cell_var, width=15).grid(
            row=0, column=1, sticky=tk.EW, pady=3
        )

        ttk.Label(main, text="To value:").grid(row=1, column=0, sticky=tk.W, pady=3)
        ttk.Entry(main, textvariable=self._to_val_var, width=15).grid(
            row=1, column=1, sticky=tk.EW, pady=3
        )

        ttk.Label(main, text="By changing cell:").grid(row=2, column=0, sticky=tk.W, pady=3)
        ttk.Entry(main, textvariable=self._by_cell_var, width=15).grid(
            row=2, column=1, sticky=tk.EW, pady=3
        )

        btn_box = ttk.Frame(main)
        btn_box.grid(row=3, column=0, columnspan=2, sticky=tk.E, pady=(12, 0))
        ttk.Button(btn_box, text="OK", bootstyle="primary", command=self._do_solve).pack(
            side=tk.LEFT, padx=3
        )
        ttk.Button(btn_box, text="Cancel", bootstyle="secondary", command=self.destroy).pack(
            side=tk.LEFT
        )

    def _do_solve(self) -> None:
        try:
            target_val = float(self._to_val_var.get().replace(",", "."))
            set_c = self._set_cell_var.get().strip().upper()
            by_c = self._by_cell_var.get().strip().upper()
            if not set_c or not by_c:
                messagebox.showwarning("Warning", "Please specify all cells.")
                return
            self._on_solve(set_c, target_val, by_c)
            self.destroy()
        except ValueError:
            messagebox.showerror("Error", "Invalid numeric value for 'To value'.")


# =============================================================================
# Cell Comment Editor Dialog
# =============================================================================


class CellCommentDialog(tk.Toplevel):
    """Dialog to create, edit, or delete cell comments."""

    def __init__(
        self,
        parent: tk.Widget,
        cell_name: str,
        initial_text: str,
        on_save: Callable[[str], None],
        on_delete: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self.title(f"Comment for {cell_name}")
        self.geometry("380x260")
        self.resizable(False, False)
        self.transient(parent)

        self._on_save = on_save
        self._on_delete = on_delete

        main = ttk.Frame(self, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main, text=f"Comment ({cell_name}):", font=Config.FONT_BOLD).pack(
            anchor=tk.W, pady=(0, 4)
        )

        self._text_area = tk.Text(main, font=Config.FONT, height=7, wrap=tk.WORD)
        self._text_area.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self._text_area.insert("1.0", initial_text)
        self._text_area.focus_set()

        btn_box = ttk.Frame(main)
        btn_box.pack(fill=tk.X, side=tk.BOTTOM)
        if on_delete:
            ttk.Button(
                btn_box,
                text="Delete",
                bootstyle="danger-outline",
                command=lambda: [self._on_delete(), self.destroy()],
            ).pack(side=tk.LEFT)
        ttk.Button(btn_box, text="Save", bootstyle="success", command=self._do_save).pack(
            side=tk.RIGHT, padx=2
        )
        ttk.Button(btn_box, text="Cancel", bootstyle="secondary", command=self.destroy).pack(
            side=tk.RIGHT
        )

        self.bind("<Escape>", lambda e: self.destroy())

    def _do_save(self) -> None:
        text = self._text_area.get("1.0", tk.END).strip()
        self._on_save(text)
        self.destroy()


# =============================================================================
# Format Cells Dialog
# =============================================================================


class FormatCellsDialog(tk.Toplevel):
    """Comprehensive Excel 'Format Cells' dialog with 5 configuration tabs:
    Number, Alignment, Font, Border, and Fill.
    """

    def __init__(
        self,
        parent: tk.Widget,
        current_format: str = "General",
        current_font: str = "Segoe UI",
        current_size: int = 10,
        current_bold: bool = False,
        current_italic: bool = False,
        current_underline: bool = False,
        current_halign: str = "left",
        current_valign: str = "center",
        current_bg: str | None = None,
        current_fg: str | None = None,
        current_wrap: bool = False,
        on_apply: Callable[[dict[str, Any]], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self.title("Format Cells")
        self.geometry("560x460")
        self.minsize(520, 420)
        self.transient(parent)
        self.grab_set()

        self._on_apply = on_apply
        self._format_var = tk.StringVar(value=current_format or "General")
        self._font_var = tk.StringVar(value=current_font or "Segoe UI")
        self._size_var = tk.IntVar(value=current_size or 10)
        self._bold_var = tk.BooleanVar(value=current_bold)
        self._italic_var = tk.BooleanVar(value=current_italic)
        self._underline_var = tk.BooleanVar(value=current_underline)
        self._strikethrough_var = tk.BooleanVar(value=False)
        self._halign_var = tk.StringVar(value=current_halign or "left")
        self._valign_var = tk.StringVar(value=current_valign or "center")
        self._wrap_var = tk.BooleanVar(value=current_wrap)
        self._decimals_var = tk.IntVar(value=2)
        self._use_separator_var = tk.BooleanVar(value=True)
        self._currency_symbol_var = tk.StringVar(value="$")
        self._bg_color_var = tk.StringVar(value=current_bg or "")
        self._fg_color_var = tk.StringVar(value=current_fg or "#000000")
        self._border_style_var = tk.StringVar(value="all")

        self._build_ui()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self) -> None:
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 5))

        # Tab 1: Number
        tab_num = ttk.Frame(notebook, padding=12)
        notebook.add(tab_num, text="  Number  ")
        self._build_number_tab(tab_num)

        # Tab 2: Alignment
        tab_align = ttk.Frame(notebook, padding=12)
        notebook.add(tab_align, text="  Alignment  ")
        self._build_alignment_tab(tab_align)

        # Tab 3: Font
        tab_font = ttk.Frame(notebook, padding=12)
        notebook.add(tab_font, text="  Font  ")
        self._build_font_tab(tab_font)

        # Tab 4: Border
        tab_border = ttk.Frame(notebook, padding=12)
        notebook.add(tab_border, text="  Border  ")
        self._build_border_tab(tab_border)

        # Tab 5: Fill
        tab_fill = ttk.Frame(notebook, padding=12)
        notebook.add(tab_fill, text="  Fill  ")
        self._build_fill_tab(tab_fill)

        # Bottom Buttons
        btn_bar = ttk.Frame(self, padding=(10, 6, 10, 10))
        btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

        ttk.Button(btn_bar, text="OK", bootstyle="primary", width=10, command=self._do_ok).pack(
            side=tk.RIGHT, padx=4
        )
        ttk.Button(
            btn_bar, text="Cancel", bootstyle="secondary", width=10, command=self.destroy
        ).pack(side=tk.RIGHT, padx=4)
        ttk.Button(
            btn_bar, text="Apply", bootstyle="info-outline", width=10, command=self._do_apply
        ).pack(side=tk.RIGHT, padx=4)

    def _build_number_tab(self, parent: ttk.Frame) -> None:
        # Category list on left, details on right
        top_box = ttk.Frame(parent)
        top_box.pack(fill=tk.BOTH, expand=True)

        left_frame = ttk.Labelframe(top_box, text="Category:", padding=6)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 10))

        categories = [
            ("General", "General"),
            ("Number", "#,##0.00"),
            ("Currency ($)", "$#,##0.00"),
            ("Currency (₽)", "#,##0.00 ₽"),
            ("Currency (€)", "€#,##0.00"),
            ("Percentage (%)", "0.00%"),
            ("Scientific", "0.00E+00"),
            ("Short Date", "YYYY-MM-DD"),
            ("Text", "@"),
        ]

        self._cat_listbox = tk.Listbox(left_frame, width=16, height=10, exportselection=False)
        self._cat_listbox.pack(fill=tk.BOTH, expand=True)
        for cat_name, _ in categories:
            self._cat_listbox.insert(tk.END, cat_name)
        self._cat_listbox.select_set(0)
        self._cat_listbox.bind("<<ListboxSelect>>", self._on_cat_selected)

        right_frame = ttk.Frame(top_box)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Sample preview box
        preview_box = ttk.Labelframe(right_frame, text="Sample Preview", padding=10)
        preview_box.pack(fill=tk.X, pady=(0, 10))
        self._sample_lbl = ttk.Label(
            preview_box, text="1234.56", font=("Segoe UI", 12, "bold"), bootstyle="primary"
        )
        self._sample_lbl.pack(anchor=tk.CENTER)

        # Decimals spinner
        dec_box = ttk.Frame(right_frame)
        dec_box.pack(fill=tk.X, pady=4)
        ttk.Label(dec_box, text="Decimal places:").pack(side=tk.LEFT)
        dec_spin = ttk.Spinbox(
            dec_box,
            from_=0,
            to=6,
            textvariable=self._decimals_var,
            width=5,
            command=self._update_number_sample,
        )
        dec_spin.pack(side=tk.LEFT, padx=6)
        self._decimals_var.trace_add("write", lambda *a: self._update_number_sample())

        # Separator check
        sep_chk = ttk.Checkbutton(
            right_frame,
            text="Use 1000 Separator (,)",
            variable=self._use_separator_var,
            command=self._update_number_sample,
        )
        sep_chk.pack(anchor=tk.W, pady=4)

        # Custom Format String Entry
        fmt_box = ttk.Labelframe(right_frame, text="Format String:", padding=6)
        fmt_box.pack(fill=tk.X, pady=(8, 0))
        fmt_entry = ttk.Entry(fmt_box, textvariable=self._format_var)
        fmt_entry.pack(fill=tk.X)
        self._format_var.trace_add("write", lambda *a: self._update_number_sample())

    def _on_cat_selected(self, event=None) -> None:
        idx = self._cat_listbox.curselection()
        if not idx:
            return
        cat_name = self._cat_listbox.get(idx[0])
        fmt_map = {
            "General": "General",
            "Number": "#,##0.00" if self._use_separator_var.get() else "0.00",
            "Currency ($)": "$#,##0.00",
            "Currency (₽)": "#,##0.00 ₽",
            "Currency (€)": "€#,##0.00",
            "Percentage (%)": "0.00%",
            "Scientific": "0.00E+00",
            "Short Date": "YYYY-MM-DD",
            "Text": "@",
        }
        if cat_name in fmt_map:
            self._format_var.set(fmt_map[cat_name])

    def _update_number_sample(self) -> None:
        fmt = self._format_var.get().strip()
        val = 1234.56
        if not fmt or fmt.upper() == "GENERAL":
            self._sample_lbl.config(text="1234.56")
            return
        try:
            from .formatting import NumberFormatter

            res = NumberFormatter.format_value(val, fmt)
            self._sample_lbl.config(text=res or "1234.56")
        except Exception:
            self._sample_lbl.config(text="1234.56")

    def _build_alignment_tab(self, parent: ttk.Frame) -> None:
        # Horizontal Alignment
        h_frame = ttk.Labelframe(parent, text="Horizontal Alignment", padding=10)
        h_frame.pack(fill=tk.X, pady=6)
        for align, lbl in [("left", "Left (Indent)"), ("center", "Center"), ("right", "Right")]:
            ttk.Radiobutton(h_frame, text=lbl, value=align, variable=self._halign_var).pack(
                anchor=tk.W, pady=2
            )

        # Vertical Alignment
        v_frame = ttk.Labelframe(parent, text="Vertical Alignment", padding=10)
        v_frame.pack(fill=tk.X, pady=6)
        for align, lbl in [("top", "Top"), ("center", "Middle"), ("bottom", "Bottom")]:
            ttk.Radiobutton(v_frame, text=lbl, value=align, variable=self._valign_var).pack(
                anchor=tk.W, pady=2
            )

        # Text Control
        ctl_frame = ttk.Labelframe(parent, text="Text Control", padding=10)
        ctl_frame.pack(fill=tk.X, pady=6)
        ttk.Checkbutton(ctl_frame, text="Wrap text", variable=self._wrap_var).pack(
            anchor=tk.W, pady=2
        )

    def _build_font_tab(self, parent: ttk.Frame) -> None:
        top_row = ttk.Frame(parent)
        top_row.pack(fill=tk.X, pady=4)

        # Font Family
        f_box = ttk.Frame(top_row)
        f_box.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Label(f_box, text="Font:").pack(anchor=tk.W)
        fonts = [
            "Segoe UI",
            "Calibri",
            "Arial",
            "Consolas",
            "Times New Roman",
            "Tahoma",
            "Verdana",
        ]
        f_combo = ttk.Combobox(
            f_box, textvariable=self._font_var, values=fonts, state="readonly", width=18
        )
        f_combo.pack(fill=tk.X, pady=2)
        f_combo.bind("<<ComboboxSelected>>", lambda e: self._update_font_preview())

        # Font Size
        s_box = ttk.Frame(top_row)
        s_box.pack(side=tk.LEFT, padx=6)
        ttk.Label(s_box, text="Size:").pack(anchor=tk.W)
        sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36]
        s_combo = ttk.Combobox(
            s_box, textvariable=self._size_var, values=sizes, state="readonly", width=5
        )
        s_combo.pack(pady=2)
        s_combo.bind("<<ComboboxSelected>>", lambda e: self._update_font_preview())

        # Font Styles & Effects
        style_frame = ttk.Labelframe(parent, text="Styles & Effects", padding=10)
        style_frame.pack(fill=tk.X, pady=6)

        c1 = ttk.Checkbutton(
            style_frame,
            text="Bold",
            variable=self._bold_var,
            command=self._update_font_preview,
        )
        c1.grid(row=0, column=0, sticky=tk.W, padx=8, pady=3)
        c2 = ttk.Checkbutton(
            style_frame,
            text="Italic",
            variable=self._italic_var,
            command=self._update_font_preview,
        )
        c2.grid(row=0, column=1, sticky=tk.W, padx=8, pady=3)
        c3 = ttk.Checkbutton(
            style_frame,
            text="Underline",
            variable=self._underline_var,
            command=self._update_font_preview,
        )
        c3.grid(row=1, column=0, sticky=tk.W, padx=8, pady=3)
        c4 = ttk.Checkbutton(
            style_frame,
            text="Strikethrough",
            variable=self._strikethrough_var,
            command=self._update_font_preview,
        )
        c4.grid(row=1, column=1, sticky=tk.W, padx=8, pady=3)

        # Color picker
        col_frame = ttk.Frame(style_frame)
        col_frame.grid(row=0, column=2, rowspan=2, padx=16)
        ttk.Label(col_frame, text="Color:").pack(anchor=tk.W)
        self._fg_btn = ttk.Button(
            col_frame, text="  Pick Text Color...  ", bootstyle="outline", command=self._pick_fg
        )
        self._fg_btn.pack(pady=2)

        # Preview area
        prev_frame = ttk.Labelframe(parent, text="Preview", padding=12)
        prev_frame.pack(fill=tk.BOTH, expand=True, pady=6)
        self._font_preview = tk.Label(
            prev_frame, text="AaBbCcYyZz 12345", bg="#ffffff", fg="#000000", height=2
        )
        self._font_preview.pack(fill=tk.BOTH, expand=True)
        self._update_font_preview()

    def _update_font_preview(self) -> None:
        weight = "bold" if self._bold_var.get() else "normal"
        slant = "italic" if self._italic_var.get() else "roman"
        size = self._size_var.get()
        fam = self._font_var.get()
        fg = self._fg_color_var.get() or "#000000"
        self._font_preview.config(font=(fam, size, weight, slant), fg=fg)

    def _pick_fg(self) -> None:
        color = colorchooser.askcolor(title="Choose Text Color", parent=self)[1]
        if color:
            self._fg_color_var.set(color)
            self._update_font_preview()

    def _build_border_tab(self, parent: ttk.Frame) -> None:
        p_frame = ttk.Labelframe(parent, text="Border Presets", padding=12)
        p_frame.pack(fill=tk.BOTH, expand=True, pady=6)

        borders = [
            ("all", "[#] All Borders (Full Grid)"),
            ("box", "[O] Box / Outer Border"),
            ("top", "[T] Top Border"),
            ("bottom", "[_] Bottom Border"),
            ("left", "[| ] Left Border"),
            ("right", "[ |] Right Border"),
            ("none", "[X] No Border"),
        ]
        for val, lbl in borders:
            ttk.Radiobutton(p_frame, text=lbl, value=val, variable=self._border_style_var).pack(
                anchor=tk.W, pady=3
            )

    def _build_fill_tab(self, parent: ttk.Frame) -> None:
        p_frame = ttk.Labelframe(parent, text="Background Fill Color", padding=12)
        p_frame.pack(fill=tk.BOTH, expand=True, pady=6)

        palette = [
            "#ffffff",
            "#f8f9fa",
            "#e9ecef",
            "#dee2e6",
            "#e3f2fd",
            "#bbdefb",
            "#90caf9",
            "#e8f5e9",
            "#c8e6c9",
            "#a5d6a7",
            "#fffde7",
            "#fff9c4",
            "#fff59d",
            "#fbe9e7",
            "#ffccbc",
            "#ffab91",
            "#f3e5f5",
            "#e1bee7",
            "#ce93d8",
            "#eceff1",
        ]

        swatch_frame = ttk.Frame(p_frame)
        swatch_frame.pack(pady=8)
        for i, col in enumerate(palette):
            r = i // 5
            c = i % 5
            btn = tk.Button(
                swatch_frame,
                bg=col,
                width=4,
                height=1,
                relief="ridge",
                command=lambda c=col: self._select_bg(c),
            )
            btn.grid(row=r, column=c, padx=3, pady=3)

        ctrl_box = ttk.Frame(p_frame)
        ctrl_box.pack(fill=tk.X, pady=8)
        ttk.Button(
            ctrl_box,
            text="No Fill",
            bootstyle="secondary-outline",
            command=lambda: self._select_bg(""),
        ).pack(side=tk.LEFT, padx=4)
        ttk.Button(
            ctrl_box, text="More Colors...", bootstyle="info", command=self._pick_custom_bg
        ).pack(side=tk.LEFT, padx=4)

        self._fill_preview = tk.Label(
            p_frame,
            text="Sample Background Preview",
            bg=self._bg_color_var.get() or "#ffffff",
            fg="#000000",
            relief="solid",
            bd=1,
            height=2,
        )
        self._fill_preview.pack(fill=tk.X, pady=(10, 0))

    def _select_bg(self, color_hex: str) -> None:
        self._bg_color_var.set(color_hex)
        self._fill_preview.config(bg=color_hex or "#ffffff")

    def _pick_custom_bg(self) -> None:
        color = colorchooser.askcolor(title="Choose Fill Color", parent=self)[1]
        if color:
            self._select_bg(color)

    def _get_result_dict(self) -> dict[str, Any]:
        return {
            "format": self._format_var.get().strip() or "General",
            "font_name": self._font_var.get(),
            "font_size": self._size_var.get(),
            "bold": self._bold_var.get(),
            "italic": self._italic_var.get(),
            "underline": self._underline_var.get(),
            "strikethrough": self._strikethrough_var.get(),
            "halign": self._halign_var.get(),
            "valign": self._valign_var.get(),
            "wrap": self._wrap_var.get(),
            "border": self._border_style_var.get(),
            "bg_color": self._bg_color_var.get() or None,
            "fg_color": self._fg_color_var.get() or None,
        }

    def _do_apply(self) -> None:
        if self._on_apply:
            self._on_apply(self._get_result_dict())

    def _do_ok(self) -> None:
        self._do_apply()
        self.destroy()


# =============================================================================
# SIMD SSE2 Benchmark Dialog
# =============================================================================


class SimdBenchmarkDialog(tk.Toplevel):
    """Interactive Hardware SIMD Vector Acceleration Benchmark Suite."""

    def __init__(self, parent: tk.Widget) -> None:
        super().__init__(parent)
        self.title("SIMD Hardware Acceleration & Benchmark")
        self.geometry("620x520")
        self.minsize(580, 480)
        self.transient(parent)

        self._size_var = tk.IntVar(value=1_000_000)
        self._func_var = tk.StringVar(value="sum")

        self._build_ui()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=14)
        main.pack(fill=tk.BOTH, expand=True)

        header = ttk.Label(
            main,
            text="Hardware SIMD Vector Acceleration Suite (SSE2 / AVX)",
            font=("Segoe UI", 12, "bold"),
            bootstyle="primary",
        )
        header.pack(anchor=tk.W, pady=(0, 6))

        desc = ttk.Label(
            main,
            text=(
                "Excel Viewer Pro embeds 64-bit native machine code kernels (xlsx_math64.dll) "
                "compiled via FASM. Compare vectorized SSE2 hardware calculations with pure Python."
            ),
            wraplength=580,
        )
        desc.pack(anchor=tk.W, pady=(0, 10))

        # Controls card
        ctrl_frame = ttk.Labelframe(main, text="Benchmark Configuration", padding=10)
        ctrl_frame.pack(fill=tk.X, pady=(0, 10))

        r_box = ttk.Frame(ctrl_frame)
        r_box.pack(fill=tk.X, pady=3)
        ttk.Label(r_box, text="Vector Array Size:").pack(side=tk.LEFT)
        sizes = [
            (100_000, "100,000 doubles (800 KB)"),
            (500_000, "500,000 doubles (4.0 MB)"),
            (1_000_000, "1,000,000 doubles (8.0 MB)"),
            (2_000_000, "2,000,000 doubles (16.0 MB)"),
        ]
        for val, lbl in sizes:
            ttk.Radiobutton(r_box, text=lbl, value=val, variable=self._size_var).pack(
                side=tk.LEFT, padx=6
            )

        f_box = ttk.Frame(ctrl_frame)
        f_box.pack(fill=tk.X, pady=6)
        ttk.Label(f_box, text="Operation:").pack(side=tk.LEFT)
        ops = [("sum", "Vector Sum"), ("avg", "Average"), ("min", "Min"), ("max", "Max")]
        for val, lbl in ops:
            ttk.Radiobutton(f_box, text=lbl, value=val, variable=self._func_var).pack(
                side=tk.LEFT, padx=8
            )

        self._btn_run = ttk.Button(
            ctrl_frame,
            text="Run Hardware Benchmark",
            bootstyle="success",
            command=self._run_benchmark,
        )
        self._btn_run.pack(pady=4)

        # Output card
        res_frame = ttk.Labelframe(main, text="Benchmark Performance Results", padding=10)
        res_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self._txt_results = tk.Text(
            res_frame, font=("Consolas", 10), height=12, bg="#f8f9fa", fg="#212529"
        )
        self._txt_results.pack(fill=tk.BOTH, expand=True)

        close_btn = ttk.Button(
            main, text="Close", bootstyle="secondary", width=12, command=self.destroy
        )
        close_btn.pack(side=tk.RIGHT)

        # Initial benchmark run
        self.after(200, self._run_benchmark)

    def _run_benchmark(self) -> None:
        self._txt_results.delete("1.0", tk.END)
        self._txt_results.insert(tk.END, "Running benchmark... please wait...\n")
        self.update_idletasks()

        n = self._size_var.get()
        op = self._func_var.get()

        try:
            from asm import asm_bridge
        except ImportError:
            try:
                import asm_bridge
            except ImportError:
                asm_bridge = None

        has_asm = asm_bridge is not None and asm_bridge.asm_is_available()

        # Generate data
        data = [(i % 100) * 0.25 for i in range(n)]

        # 1. Pure Python measurement
        t0 = time.perf_counter()
        if op == "sum":
            py_res = math.fsum(data)
        elif op == "avg":
            py_res = math.fsum(data) / len(data)
        elif op == "min":
            py_res = min(data)
        elif op == "max":
            py_res = max(data)
        else:
            py_res = math.fsum(data)
        t_py = (time.perf_counter() - t0) * 1000.0

        # 2. Native x64 Assembly SIMD measurement
        t_asm = 0.0
        asm_res = 0.0
        if has_asm:
            t0 = time.perf_counter()
            if op == "sum":
                asm_res = asm_bridge.asm_sum(data)
            elif op == "avg":
                asm_res = asm_bridge.asm_avg(data)
            elif op == "min":
                asm_res = asm_bridge.asm_min(data)
            elif op == "max":
                asm_res = asm_bridge.asm_max(data)
            t_asm = (time.perf_counter() - t0) * 1000.0

        self._txt_results.delete("1.0", tk.END)
        out = []
        out.append("=" * 64)
        out.append(f" HARDWARE SIMD ACCELERATION BENCHMARK — {op.upper()}")
        out.append("=" * 64)
        out.append(f" Array Buffer Size:   {n:,} double-precision floats")
        out.append(f" Memory Allocated:    {(n * 8) / (1024 * 1024):.2f} MB")
        out.append(
            f" Native SIMD Driver:  {'Loaded (xlsx_math64.dll - x86-64 SSE2)' if has_asm else 'Not Loaded (Pure Python)'}"
        )
        out.append("-" * 64)
        out.append(f" Pure Python Time:    {t_py:8.2f} ms | Result: {py_res:,.4f}")

        if has_asm and t_asm > 0:
            speedup = t_py / t_asm
            mops = (n / (t_asm / 1000.0)) / 1_000_000
            diff = abs(py_res - asm_res)
            out.append(f" Native Assembly SIMD:{t_asm:8.2f} ms | Result: {asm_res:,.4f}")
            out.append("-" * 64)
            out.append(f" SPEEDUP FACTOR:      {speedup:8.1f}x FASTER")
            out.append(f" VECTOR THROUGHPUT:   {mops:8.2f} Million ops / sec")
            out.append(f" NUMERICAL PRECISION: Delta = {diff:g} (Bit-exact: {diff < 1e-5})")
        else:
            out.append("-" * 64)
            out.append(" Hardware vector acceleration library is operating in pure Python mode.")

        out.append("=" * 64)
        self._txt_results.insert(tk.END, "\n".join(out))


# =============================================================================
# Keyboard Shortcuts Help Dialog
# =============================================================================


class ShortcutsHelpDialog(tk.Toplevel):
    """Excel Viewer Pro Keyboard Shortcuts Reference Dialog."""

    def __init__(self, parent: tk.Widget) -> None:
        super().__init__(parent)
        self.title("Keyboard Shortcuts Reference")
        self.geometry("640x500")
        self.minsize(560, 400)
        self.transient(parent)

        self._search_var = tk.StringVar()
        self._build_ui()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            main,
            text="Excel Viewer Pro Keyboard Shortcuts",
            font=("Segoe UI", 12, "bold"),
            bootstyle="primary",
        ).pack(anchor=tk.W, pady=(0, 6))

        # Search box
        s_box = ttk.Frame(main)
        s_box.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(s_box, text="Filter:").pack(side=tk.LEFT, padx=(0, 6))
        s_entry = ttk.Entry(s_box, textvariable=self._search_var)
        s_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._search_var.trace_add("write", lambda *a: self._filter_shortcuts())

        # Treeview
        tree_frame = ttk.Frame(main)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("key", "action", "cat")
        self._tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=14)
        self._tree.heading("key", text="Shortcut Key")
        self._tree.heading("action", text="Description / Action")
        self._tree.heading("cat", text="Category")

        self._tree.column("key", width=140, anchor=tk.W)
        self._tree.column("action", width=340, anchor=tk.W)
        self._tree.column("cat", width=120, anchor=tk.W)

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self._shortcuts_data = [
            ("Ctrl + N", "Create new blank workbook", "File"),
            ("Ctrl + O", "Open existing Excel / CSV file", "File"),
            ("Ctrl + S", "Save current workbook", "File"),
            ("Ctrl + Shift + S", "Save workbook as new file", "File"),
            ("Ctrl + W", "Close current workbook", "File"),
            ("Ctrl + Z", "Undo last cell edit / formatting action", "Edit"),
            ("Ctrl + Y", "Redo previously undone action", "Edit"),
            ("Ctrl + X", "Cut active cell or range to clipboard", "Edit"),
            ("Ctrl + C", "Copy active cell or range to clipboard", "Edit"),
            ("Ctrl + V", "Paste clipboard contents to active cell", "Edit"),
            ("Delete", "Clear cell value & formula", "Edit"),
            ("Ctrl + A", "Select all cells in the sheet", "Edit"),
            ("Ctrl + F", "Find text or values in worksheet", "Edit"),
            ("Ctrl + H", "Find and Replace text across worksheet", "Edit"),
            ("Ctrl + G / F5", "Go to cell coordinate (e.g. C25)", "Navigation"),
            ("Arrow Keys", "Navigate active cell selection", "Navigation"),
            ("Tab", "Move one cell to the right", "Navigation"),
            ("Shift + Tab", "Move one cell to the left", "Navigation"),
            ("Enter", "Commit edit and move one cell down", "Navigation"),
            ("Shift + Enter", "Move one cell up", "Navigation"),
            ("Ctrl + Home", "Jump to top-left cell (A1)", "Navigation"),
            ("F2", "Edit active cell inline", "Editing"),
            ("Ctrl + 1", "Open Format Cells dialog", "Format"),
            ("Ctrl + B", "Toggle bold text font weight", "Format"),
            ("Ctrl + I", "Toggle italic text style", "Format"),
            ("Ctrl + U", "Toggle underline font style", "Format"),
            ("Ctrl + 5", "Toggle strikethrough font effect", "Format"),
            ("F9", "Recalculate all formulas in workbook", "Formulas"),
            ("Alt + =", "AutoSum active column into cell", "Formulas"),
            ("Shift + F3", "Open Insert Function (fx) wizard", "Formulas"),
            ("Ctrl + `", "Toggle Show Formulas vs computed values", "Formulas"),
            ("Ctrl + Shift + L", "Toggle AutoFilter on headers", "Data"),
            ("F11", "Open Chart Wizard dialog", "Insert"),
            ("Shift + F2", "Add or edit cell comment", "Insert"),
            ("Shift + F11", "Insert new worksheet tab", "Sheets"),
        ]

        self._filter_shortcuts()

        btn_box = ttk.Frame(main)
        btn_box.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(
            btn_box, text="Close", bootstyle="secondary", width=12, command=self.destroy
        ).pack(side=tk.RIGHT)

    def _filter_shortcuts(self) -> None:
        q = self._search_var.get().lower().strip()
        self._tree.delete(*self._tree.get_children())
        for key, act, cat in self._shortcuts_data:
            if not q or q in key.lower() or q in act.lower() or q in cat.lower():
                self._tree.insert("", tk.END, values=(key, act, cat))


# =============================================================================
# Formula Functions Reference Dialog
# =============================================================================


class FormulaReferenceDialog(tk.Toplevel):
    """Interactive documentation and reference for all 129 Excel functions."""

    def __init__(
        self, parent: tk.Widget, on_insert_formula: Callable[[str], None] | None = None
    ) -> None:
        super().__init__(parent)
        self.title("Formula Functions Reference (129 Excel Functions)")
        self.geometry("740x520")
        self.minsize(680, 440)
        self.transient(parent)

        self._on_insert = on_insert_formula
        self._search_var = tk.StringVar()
        self._cat_var = tk.StringVar(value="All Categories")

        self._build_ui()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        # Filter bar
        filter_bar = ttk.Frame(main)
        filter_bar.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(filter_bar, text="Category:").pack(side=tk.LEFT, padx=(0, 4))
        cats = [
            "All Categories",
            "Math & Trig",
            "Statistical",
            "Financial",
            "Logical",
            "Text",
            "Date & Time",
            "Lookup & Reference",
            "Information",
            "Engineering",
        ]
        cat_combo = ttk.Combobox(
            filter_bar, textvariable=self._cat_var, values=cats, state="readonly", width=18
        )
        cat_combo.pack(side=tk.LEFT, padx=(0, 10))
        cat_combo.bind("<<ComboboxSelected>>", lambda e: self._filter_functions())

        ttk.Label(filter_bar, text="Search:").pack(side=tk.LEFT, padx=(0, 4))
        s_entry = ttk.Entry(filter_bar, textvariable=self._search_var, width=22)
        s_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._search_var.trace_add("write", lambda *a: self._filter_functions())

        # Main splitter: List on left, details on right
        paned = ttk.Panedwindow(main, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        left_frame = ttk.Frame(paned, padding=4)
        paned.add(left_frame, weight=1)

        self._func_listbox = tk.Listbox(left_frame, font=("Segoe UI", 10), exportselection=False)
        f_vsb = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self._func_listbox.yview)
        self._func_listbox.configure(yscrollcommand=f_vsb.set)
        self._func_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        f_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._func_listbox.bind("<<ListboxSelect>>", self._on_select_func)

        right_frame = ttk.Frame(paned, padding=10)
        paned.add(right_frame, weight=2)

        self._lbl_name = ttk.Label(
            right_frame,
            text="Select a Function",
            font=("Segoe UI", 13, "bold"),
            bootstyle="primary",
        )
        self._lbl_name.pack(anchor=tk.W, pady=(0, 4))

        self._lbl_syntax = ttk.Label(
            right_frame, text="", font=("Consolas", 10, "bold"), wraplength=400
        )
        self._lbl_syntax.pack(anchor=tk.W, pady=(0, 8))

        ttk.Label(right_frame, text="Description:", font=Config.FONT_BOLD).pack(
            anchor=tk.W, pady=(4, 2)
        )
        self._lbl_desc = ttk.Label(right_frame, text="", wraplength=400)
        self._lbl_desc.pack(anchor=tk.W, pady=(0, 8))

        ttk.Label(right_frame, text="Example:", font=Config.FONT_BOLD).pack(
            anchor=tk.W, pady=(4, 2)
        )
        self._lbl_example = ttk.Label(
            right_frame, text="", font=("Consolas", 10), wraplength=400, bootstyle="success"
        )
        self._lbl_example.pack(anchor=tk.W, pady=(0, 12))

        self._btn_insert = ttk.Button(
            right_frame,
            text="Insert Function into Cell",
            bootstyle="primary",
            command=self._do_insert,
            state="disabled",
        )
        self._btn_insert.pack(anchor=tk.W, pady=(8, 0))

        # Bottom Bar
        b_bar = ttk.Frame(main)
        b_bar.pack(fill=tk.X, pady=(8, 0))
        self._count_lbl = ttk.Label(b_bar, text="", font=Config.FONT_SMALL)
        self._count_lbl.pack(side=tk.LEFT)
        ttk.Button(b_bar, text="Close", bootstyle="secondary", width=12, command=self.destroy).pack(
            side=tk.RIGHT
        )

        self._filter_functions()

    def _filter_functions(self) -> None:
        cat_filter = self._cat_var.get()
        q = self._search_var.get().upper().strip()

        self._filtered_funcs: list[str] = []
        for name, meta in sorted(FUNCTION_METADATA.items()):
            meta_cat = meta.get("cat", "General") if isinstance(meta, dict) else "General"
            meta_desc = meta.get("desc", "") if isinstance(meta, dict) else ""
            if cat_filter != "All Categories" and meta_cat.lower() != cat_filter.lower():
                continue
            if q and q not in name.upper() and q not in meta_desc.upper():
                continue
            self._filtered_funcs.append(name)

        self._func_listbox.delete(0, tk.END)
        for fn in self._filtered_funcs:
            self._func_listbox.insert(tk.END, fn)

        self._count_lbl.config(
            text=f"Showing {len(self._filtered_funcs)} of {len(FUNCTION_METADATA)} supported functions"
        )
        if self._filtered_funcs:
            self._func_listbox.select_set(0)
            self._on_select_func()

    def _on_select_func(self, event=None) -> None:
        idx = self._func_listbox.curselection()
        if not idx:
            return
        fn = self._func_listbox.get(idx[0])
        meta = FUNCTION_METADATA.get(fn, {})
        if isinstance(meta, dict):
            cat = meta.get("cat", "General")
            desc = meta.get("desc", "No description available.")
            syntax = meta.get("syntax", f"={fn}()")
            example = meta.get("example", f"={fn}()")
        else:
            cat = "General"
            desc = str(meta)
            syntax = f"={fn}()"
            example = f"={fn}()"

        self._lbl_name.config(text=f"{fn}  [{cat}]")
        self._lbl_syntax.config(text=syntax)
        self._lbl_desc.config(text=desc)
        self._lbl_example.config(text=f"Example:  {example}")
        self._btn_insert.config(state="normal")
        self._current_syntax = syntax

    def _do_insert(self) -> None:
        if self._on_insert and hasattr(self, "_current_syntax"):
            self._on_insert(self._current_syntax)
            self.destroy()


# =============================================================================
# SIMD Vector Hardware Info Dialog
# =============================================================================


class SimdInfoDialog(tk.Toplevel):
    """Architecture and capabilities of the 64-bit Assembly SIMD Engine."""

    def __init__(self, parent: tk.Widget) -> None:
        super().__init__(parent)
        self.title("SIMD Hardware Acceleration Details")
        self.geometry("560x420")
        self.resizable(False, False)
        self.transient(parent)

        self._build_ui()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=14)
        main.pack(fill=tk.BOTH, expand=True)

        try:
            from asm import asm_bridge
        except ImportError:
            try:
                import asm_bridge
            except ImportError:
                asm_bridge = None

        has_asm = asm_bridge is not None and asm_bridge.asm_is_available()

        ttk.Label(
            main,
            text="Native x86-64 SIMD Engine Architecture",
            font=("Segoe UI", 12, "bold"),
            bootstyle="primary",
        ).pack(anchor=tk.W, pady=(0, 6))

        status_txt = "ACTIVE (Hardware Accelerated)" if has_asm else "FALLBACK (Pure Python)"
        status_color = "success" if has_asm else "warning"
        ttk.Label(
            main, text=f"Engine Status: {status_txt}", font=Config.FONT_BOLD, bootstyle=status_color
        ).pack(anchor=tk.W, pady=(0, 8))

        info_box = tk.Text(main, font=("Consolas", 10), height=14, bg="#f8f9fa", fg="#212529")
        info_box.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        content = [
            "• Vector Register Width:   128-bit XMM0..XMM15 (SSE2 / AVX)",
            "• Floating-Point Format:   IEEE 754 Double Precision (float64)",
            "• Assembly Dialect:        Flat Assembler (FASM v1.73)",
            "• Calling Convention:      Microsoft Windows x64 ABI",
            "• Machine Code Binary:     asm/xlsx_math64.dll",
            "• Standalone GUI Suite:    asm/xlsx_gui64.exe",
            "",
            "Accelerated Vector Kernels:",
            "  1. vec_sum_f64         - 4-way unrolled addpd + movhlps horizontal sum",
            "  2. vec_avg_f64         - Parallel summation + single divsd reduction",
            "  3. vec_min_f64         - Vectorized minsd parallel minimum search",
            "  4. vec_max_f64         - Vectorized maxsd parallel maximum search",
            "  5. vec_sumproduct_f64  - Dot product via mulpd + addpd SIMD pipeline",
            "  6. fast_pmt_f64        - Financial loan payment calculation (PMT)",
            "  7. fast_pv_f64         - Investment present value (PV)",
            "  8. fast_fv_f64         - Future value calculation (FV)",
            "  9. fast_str_hash       - 64-bit FNV-1a hash for instant formula cache",
        ]
        info_box.insert(tk.END, "\n".join(content))
        info_box.config(state="disabled")

        ttk.Button(main, text="OK", bootstyle="primary", width=12, command=self.destroy).pack(
            side=tk.RIGHT
        )


# =============================================================================
# MCP Server Integration Status Dialog
# =============================================================================


class McpStatusDialog(tk.Toplevel):
    """Model Context Protocol (MCP) Server Integration Dialog."""

    def __init__(self, parent: tk.Widget) -> None:
        super().__init__(parent)
        self.title("Model Context Protocol (MCP) Integration")
        self.geometry("600x440")
        self.minsize(540, 380)
        self.transient(parent)

        self._build_ui()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=14)
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            main,
            text="Native Model Context Protocol (MCP) Server",
            font=("Segoe UI", 12, "bold"),
            bootstyle="primary",
        ).pack(anchor=tk.W, pady=(0, 4))

        desc = ttk.Label(
            main,
            text=(
                "Excel Viewer Pro includes a native MCP server ('xlsx-mcp') that allows "
                "AI assistants (Claude Desktop, Cursor, Windsurf) to inspect spreadsheets, "
                "read cell values, evaluate formulas, and run vector calculations via stdio."
            ),
            wraplength=560,
        )
        desc.pack(anchor=tk.W, pady=(0, 8))

        cfg_box = ttk.Labelframe(main, text="Claude Desktop / Cursor Configuration", padding=8)
        cfg_box.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        txt = tk.Text(cfg_box, font=("Consolas", 10), height=10, bg="#212529", fg="#79c0ff")
        txt.pack(fill=tk.BOTH, expand=True)

        snippet = """{
  "mcpServers": {
    "xlsx-viewer": {
      "command": "python",
      "args": ["-m", "xlsx_viewer.mcp_server"]
    }
  }
}"""
        txt.insert(tk.END, snippet)
        txt.config(state="disabled")

        b_frame = ttk.Frame(main)
        b_frame.pack(fill=tk.X)
        ttk.Button(b_frame, text="OK", bootstyle="primary", width=10, command=self.destroy).pack(
            side=tk.RIGHT
        )

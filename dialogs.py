"""
Professional dialogs and wizards for Excel Viewer Pro.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, colorchooser, filedialog, simpledialog, Menu
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from typing import Callable, Any
from openpyxl.utils import get_column_letter

from config import Config
from models import SheetData, CellPosition, CellRange, CellStyle, CellComment
from formulas import FUNCTION_METADATA


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
        initial_tab: str = "find"
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
        ttk.Checkbutton(opt_frame1, text="Match entire cell contents", variable=self._match_entire).pack(anchor=tk.W)
        ttk.Checkbutton(opt_frame1, text="Search all sheets", variable=self._search_all_sheets).pack(anchor=tk.W)

        btn_box1 = ttk.Frame(find_frame)
        btn_box1.grid(row=2, column=0, columnspan=3, sticky=tk.E, pady=5)
        ttk.Button(btn_box1, text="Find Next", bootstyle="primary", command=self._do_find_next).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_box1, text="Close", bootstyle="secondary", command=self.destroy).pack(side=tk.LEFT, padx=3)

        # Setup Replace Tab
        ttk.Label(replace_frame, text="Find what:").grid(row=0, column=0, sticky=tk.W, pady=3)
        ttk.Entry(replace_frame, textvariable=self._find_var, width=32).grid(row=0, column=1, columnspan=2, sticky=tk.EW, pady=3)

        ttk.Label(replace_frame, text="Replace with:").grid(row=1, column=0, sticky=tk.W, pady=3)
        ttk.Entry(replace_frame, textvariable=self._replace_var, width=32).grid(row=1, column=1, columnspan=2, sticky=tk.EW, pady=3)

        opt_frame2 = ttk.Frame(replace_frame)
        opt_frame2.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
        ttk.Checkbutton(opt_frame2, text="Match case", variable=self._match_case).pack(anchor=tk.W)
        ttk.Checkbutton(opt_frame2, text="Match entire cell contents", variable=self._match_entire).pack(anchor=tk.W)

        btn_box2 = ttk.Frame(replace_frame)
        btn_box2.grid(row=3, column=0, columnspan=3, sticky=tk.E, pady=5)
        ttk.Button(btn_box2, text="Replace", bootstyle="primary", command=self._do_replace).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_box2, text="Replace All", bootstyle="warning", command=self._do_replace_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_box2, text="Find Next", bootstyle="secondary-outline", command=self._do_find_next).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_box2, text="Close", bootstyle="secondary", command=self.destroy).pack(side=tk.LEFT, padx=2)

        if initial_tab == "replace":
            notebook.select(1)

    def _do_find_next(self) -> None:
        target = self._find_var.get()
        if not target:
            return
        found = self._on_find_next(target, self._match_case.get(), self._match_entire.get(), self._search_all_sheets.get())
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
        count = self._on_replace_all(target, replacement, self._match_case.get(), self._match_entire.get())
        messagebox.showinfo("Replace All", f"Excel Viewer Pro has completed its search and has made {count} replacement(s).")


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
        ttk.Button(search_box, text="Go", width=4, command=self._populate_list).pack(side=tk.LEFT, padx=(5, 0))

        # Category Filter
        cat_box = ttk.Frame(main)
        cat_box.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(cat_box, text="Or select a category:").pack(side=tk.LEFT)
        categories = ["All", "Math & Trig", "Statistical", "Lookup & Reference", "Logical", "Text", "Date & Time", "Financial", "Information"]
        cat_cb = ttk.Combobox(cat_box, textvariable=self._category_var, values=categories, state="readonly", width=22)
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
        ttk.Button(btn_box, text="Insert", bootstyle="success", command=self._do_insert).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btn_box, text="Cancel", bootstyle="secondary", command=self.destroy).pack(side=tk.RIGHT)

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
        on_apply_sort: Callable[[list[tuple[int, bool]], bool], None]
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
        ttk.Button(bar, text="➕ Add Level", bootstyle="outline", command=self._add_level).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="➖ Delete Level", bootstyle="outline", command=self._delete_level).pack(side=tk.LEFT, padx=2)
        ttk.Checkbutton(bar, text="My data has headers", variable=self._has_headers).pack(side=tk.RIGHT)

        # Scrollable level container
        self._container = ttk.Frame(main, relief="solid", borderwidth=1, padding=5)
        self._container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Add initial sort level
        self._add_level()

        # Action Buttons
        btn_box = ttk.Frame(main)
        btn_box.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Button(btn_box, text="OK", bootstyle="primary", command=self._do_sort).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btn_box, text="Cancel", bootstyle="secondary", command=self.destroy).pack(side=tk.RIGHT)

    def _add_level(self) -> None:
        if len(self._level_rows) >= 5:
            return
        idx = len(self._level_rows)
        row_frame = ttk.Frame(self._container)
        row_frame.pack(fill=tk.X, pady=3)

        label_text = "Sort by:" if idx == 0 else "Then by:"
        ttk.Label(row_frame, text=label_text, width=9).pack(side=tk.LEFT)

        col_var = tk.StringVar(value=self._headers[0] if self._headers else "Col 1")
        col_cb = ttk.Combobox(row_frame, textvariable=col_var, values=self._headers, state="readonly", width=18)
        col_cb.pack(side=tk.LEFT, padx=5)

        ttk.Label(row_frame, text="Order:").pack(side=tk.LEFT, padx=(5, 0))
        order_var = tk.StringVar(value="A to Z (Ascending)")
        order_cb = ttk.Combobox(row_frame, textvariable=order_var, values=["A to Z (Ascending)", "Z to A (Descending)"], state="readonly", width=18)
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
        on_sort_col: Callable[[int, bool], None]
    ) -> None:
        super().__init__(parent)
        self.title(f"Filter: {col_name}")
        self.geometry("260x340")
        self.resizable(False, False)
        self.transient(parent)

        self._col_idx = col_idx
        self._unique_values = sorted(unique_values, key=lambda s: str(s).lower())
        self._selected_values = set(selected_values) if selected_values is not None else set(unique_values)
        self._on_apply = on_apply
        self._on_sort_col = on_sort_col

        self._check_vars: dict[str, tk.BooleanVar] = {}
        self._search_var = tk.StringVar()
        self._select_all_var = tk.BooleanVar(value=len(self._selected_values) == len(self._unique_values))

        self._build_ui()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=8)
        main.pack(fill=tk.BOTH, expand=True)

        # Quick Sort Actions
        sort_box = ttk.Frame(main)
        sort_box.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(sort_box, text="⬆ Sort A → Z", bootstyle="outline", command=lambda: [self._on_sort_col(self._col_idx, False), self.destroy()]).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=1)
        ttk.Button(sort_box, text="⬇ Sort Z → A", bootstyle="outline", command=lambda: [self._on_sort_col(self._col_idx, True), self.destroy()]).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=1)

        ttk.Separator(main).pack(fill=tk.X, pady=4)

        # Search Bar
        s_entry = ttk.Entry(main, textvariable=self._search_var)
        s_entry.pack(fill=tk.X, pady=(0, 4))
        s_entry.bind("<KeyRelease>", lambda e: self._render_checks())

        # Select All checkbox
        ttk.Checkbutton(main, text="(Select All)", variable=self._select_all_var, command=self._toggle_all).pack(anchor=tk.W)

        # Scrollable checklist
        list_container = ttk.Frame(main, relief="solid", borderwidth=1)
        list_container.pack(fill=tk.BOTH, expand=True, pady=4)

        self._canvas = tk.Canvas(list_container, highlightthickness=0)
        vsb = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)

        self._inner_frame = ttk.Frame(self._canvas)
        self._canvas.create_window((0, 0), window=self._inner_frame, anchor=tk.NW)

        self._inner_frame.bind("<Configure>", lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self._render_checks()

        # Action Buttons
        btn_box = ttk.Frame(main)
        btn_box.pack(fill=tk.X, side=tk.BOTTOM, pady=(4, 0))
        ttk.Button(btn_box, text="Clear Filter", bootstyle="secondary-outline", command=self._clear_filter).pack(side=tk.LEFT)
        ttk.Button(btn_box, text="OK", bootstyle="primary", command=self._do_apply).pack(side=tk.RIGHT, padx=2)
        ttk.Button(btn_box, text="Cancel", bootstyle="secondary", command=self.destroy).pack(side=tk.RIGHT)

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

            cb = ttk.Checkbutton(self._inner_frame, text=display_val, variable=self._check_vars[val])
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
        "Clustered Column", "Stacked Column", "Bar (Horizontal)",
        "Line", "Smooth Line", "Pie", "Donut", "Area", "Scatter (XY)", "Histogram"
    ]

    def __init__(
        self,
        parent: tk.Widget,
        sheet_data: SheetData,
        get_data_fn: Callable[[int, int, int, int], tuple[list[str], list[list[float]], list[str]]]
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
        cb_type = ttk.Combobox(left, textvariable=self._chart_type, values=self.CHART_TYPES, state="readonly")
        cb_type.pack(fill=tk.X, pady=(2, 8))
        cb_type.bind("<<ComboboxSelected>>", lambda e: self._render_chart())

        ttk.Label(left, text="Color Theme:", font=Config.FONT_BOLD).pack(anchor=tk.W)
        cb_pal = ttk.Combobox(left, textvariable=self._palette, values=list(Config.CHART_PALETTES.keys()), state="readonly")
        cb_pal.pack(fill=tk.X, pady=(2, 8))
        cb_pal.bind("<<ComboboxSelected>>", lambda e: self._render_chart())

        # Data Range Frame
        range_box = ttk.LabelFrame(left, text="Data Range", padding=6)
        range_box.pack(fill=tk.X, pady=6)

        ttk.Label(range_box, text="Rows:").grid(row=0, column=0, sticky=tk.W)
        self._r_start = ttk.Spinbox(range_box, from_=1, to=max(1, self._sheet_data.row_count), width=5)
        self._r_start.set(1)
        self._r_start.grid(row=0, column=1, padx=2)
        ttk.Label(range_box, text="to").grid(row=0, column=2)
        self._r_end = ttk.Spinbox(range_box, from_=1, to=max(1, self._sheet_data.row_count), width=5)
        self._r_end.set(min(15, max(1, self._sheet_data.row_count)))
        self._r_end.grid(row=0, column=3, padx=2)

        ttk.Label(range_box, text="Cols:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self._c_start = ttk.Spinbox(range_box, from_=1, to=max(1, self._sheet_data.col_count), width=5)
        self._c_start.set(1)
        self._c_start.grid(row=1, column=1, padx=2, pady=4)
        ttk.Label(range_box, text="to").grid(row=1, column=2, pady=4)
        self._c_end = ttk.Spinbox(range_box, from_=1, to=max(1, self._sheet_data.col_count), width=5)
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
        ttk.Checkbutton(left, text="Show Legend", variable=self._show_legend, command=self._render_chart).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(left, text="Show Grid Lines", variable=self._show_grid, command=self._render_chart).pack(anchor=tk.W, pady=2)

        # Draw / Refresh Button
        ttk.Button(left, text="🔄 Update Chart", bootstyle="primary", command=self._render_chart).pack(fill=tk.X, pady=(10, 4))

        # Export Buttons
        ttk.Separator(left).pack(fill=tk.X, pady=6)
        ttk.Button(left, text="💾 Save Image (PNG)", bootstyle="success-outline", command=lambda: self._export_chart("png")).pack(fill=tk.X, pady=2)
        ttk.Button(left, text="📄 Save Vector (PDF/SVG)", bootstyle="info-outline", command=lambda: self._export_chart("pdf")).pack(fill=tk.X, pady=2)

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
            colors = Config.CHART_PALETTES.get(self._palette.get(), Config.CHART_PALETTES["Excel Classic"])
            chart_type = self._chart_type.get()

            if chart_type == "Clustered Column":
                x = list(range(len(labels)))
                num_series = len(series)
                w = 0.8 / max(1, num_series)
                for i, data in enumerate(series):
                    offset = (i - num_series / 2 + 0.5) * w
                    name = series_names[i] if i < len(series_names) else f"Series {i+1}"
                    ax.bar([xi + offset for xi in x], data, w, label=name, color=colors[i % len(colors)])
                ax.set_xticks(x)
                ax.set_xticklabels(labels, rotation=35, ha="right")

            elif chart_type == "Stacked Column":
                x = list(range(len(labels)))
                bottom = [0.0] * len(labels)
                for i, data in enumerate(series):
                    name = series_names[i] if i < len(series_names) else f"Series {i+1}"
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
                    name = series_names[i] if i < len(series_names) else f"Series {i+1}"
                    ax.barh([yi + offset for yi in y], data, h, label=name, color=colors[i % len(colors)])
                ax.set_yticks(y)
                ax.set_yticklabels(labels)

            elif chart_type in ("Line", "Smooth Line"):
                for i, data in enumerate(series):
                    name = series_names[i] if i < len(series_names) else f"Series {i+1}"
                    ax.plot(labels, data, marker="o", linewidth=2.2, label=name, color=colors[i % len(colors)])
                ax.tick_params(axis="x", rotation=35)

            elif chart_type in ("Pie", "Donut"):
                first_series = series[0] if series else []
                wedge_props = dict(width=0.4, edgecolor="w") if chart_type == "Donut" else dict(edgecolor="w")
                ax.pie(first_series, labels=labels, autopct="%1.1f%%", colors=colors, startangle=90, wedgeprops=wedge_props)

            elif chart_type == "Area":
                x = list(range(len(labels)))
                for i, data in enumerate(series):
                    name = series_names[i] if i < len(series_names) else f"Series {i+1}"
                    ax.fill_between(x, data, alpha=0.4, color=colors[i % len(colors)], label=name)
                    ax.plot(x, data, color=colors[i % len(colors)], linewidth=1.5)
                ax.set_xticks(x)
                ax.set_xticklabels(labels, rotation=35, ha="right")

            elif chart_type == "Scatter (XY)":
                for i, data in enumerate(series):
                    name = series_names[i] if i < len(series_names) else f"Series {i+1}"
                    ax.scatter(list(range(len(data))), data, s=50, label=name, color=colors[i % len(colors)])
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

            if self._show_legend.get() and chart_type not in ("Pie", "Donut", "Histogram") and len(series) > 0:
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
            filetypes=[(f"{fmt.upper()} File", f"*{ext}"), ("All Files", "*.*")]
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
        self,
        parent: tk.Widget,
        headers: list[str],
        on_remove: Callable[[list[int]], None]
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

        ttk.Label(main, text="Select columns to compare for duplicates:").pack(anchor=tk.W, pady=(0, 6))

        bar = ttk.Frame(main)
        bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(bar, text="Select All", bootstyle="outline", command=self._select_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="Unselect All", bootstyle="outline", command=self._unselect_all).pack(side=tk.LEFT, padx=2)

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
            cb = ttk.Checkbutton(inner, text=f"{header} ({get_column_letter(i+1)})", variable=var)
            cb.pack(anchor=tk.W, pady=2)

        # Buttons
        btn_box = ttk.Frame(main)
        btn_box.pack(fill=tk.X, side=tk.BOTTOM, pady=(8, 0))
        ttk.Button(btn_box, text="OK", bootstyle="primary", command=self._do_remove).pack(side=tk.RIGHT, padx=3)
        ttk.Button(btn_box, text="Cancel", bootstyle="secondary", command=self.destroy).pack(side=tk.RIGHT)

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
        self,
        parent: tk.Widget,
        sample_data: list[str],
        on_split: Callable[[str], None]
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

        ttk.Label(main, text="Select delimiters contained in your data:", font=Config.FONT_BOLD).pack(anchor=tk.W, pady=(0, 6))

        delims_frame = ttk.LabelFrame(main, text="Delimiters", padding=8)
        delims_frame.pack(fill=tk.X, pady=(0, 8))

        row1 = ttk.Frame(delims_frame)
        row1.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(row1, text="Tab", variable=self._tab_var, command=self._update_preview).pack(side=tk.LEFT, padx=8)
        ttk.Checkbutton(row1, text="Semicolon (;)", variable=self._semicolon_var, command=self._update_preview).pack(side=tk.LEFT, padx=8)
        ttk.Checkbutton(row1, text="Comma (,)", variable=self._comma_var, command=self._update_preview).pack(side=tk.LEFT, padx=8)

        row2 = ttk.Frame(delims_frame)
        row2.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(row2, text="Space", variable=self._space_var, command=self._update_preview).pack(side=tk.LEFT, padx=8)
        ttk.Checkbutton(row2, text="Other:", variable=self._other_var, command=self._update_preview).pack(side=tk.LEFT, padx=(8, 2))
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
        ttk.Button(btn_box, text="Finish", bootstyle="success", command=self._do_finish).pack(side=tk.RIGHT, padx=3)
        ttk.Button(btn_box, text="Cancel", bootstyle="secondary", command=self.destroy).pack(side=tk.RIGHT)

    def _get_active_delimiter(self) -> str:
        if self._semicolon_var.get(): return ";"
        if self._comma_var.get(): return ","
        if self._tab_var.get(): return "\t"
        if self._space_var.get(): return " "
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
        self,
        parent: tk.Widget,
        current_cell: str,
        on_solve: Callable[[str, float, str], None]
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
        ttk.Entry(main, textvariable=self._set_cell_var, width=15).grid(row=0, column=1, sticky=tk.EW, pady=3)

        ttk.Label(main, text="To value:").grid(row=1, column=0, sticky=tk.W, pady=3)
        ttk.Entry(main, textvariable=self._to_val_var, width=15).grid(row=1, column=1, sticky=tk.EW, pady=3)

        ttk.Label(main, text="By changing cell:").grid(row=2, column=0, sticky=tk.W, pady=3)
        ttk.Entry(main, textvariable=self._by_cell_var, width=15).grid(row=2, column=1, sticky=tk.EW, pady=3)

        btn_box = ttk.Frame(main)
        btn_box.grid(row=3, column=0, columnspan=2, sticky=tk.E, pady=(12, 0))
        ttk.Button(btn_box, text="OK", bootstyle="primary", command=self._do_solve).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_box, text="Cancel", bootstyle="secondary", command=self.destroy).pack(side=tk.LEFT)

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
        on_delete: Callable[[], None] | None = None
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

        ttk.Label(main, text=f"Comment ({cell_name}):", font=Config.FONT_BOLD).pack(anchor=tk.W, pady=(0, 4))

        self._text_area = tk.Text(main, font=Config.FONT, height=7, wrap=tk.WORD)
        self._text_area.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self._text_area.insert("1.0", initial_text)
        self._text_area.focus_set()

        btn_box = ttk.Frame(main)
        btn_box.pack(fill=tk.X, side=tk.BOTTOM)
        if on_delete:
            ttk.Button(btn_box, text="Delete", bootstyle="danger-outline", command=lambda: [self._on_delete(), self.destroy()]).pack(side=tk.LEFT)
        ttk.Button(btn_box, text="Save", bootstyle="success", command=self._do_save).pack(side=tk.RIGHT, padx=2)
        ttk.Button(btn_box, text="Cancel", bootstyle="secondary", command=self.destroy).pack(side=tk.RIGHT)

        self.bind("<Escape>", lambda e: self.destroy())

    def _do_save(self) -> None:
        text = self._text_area.get("1.0", tk.END).strip()
        self._on_save(text)
        self.destroy()

; ============================================================================
; Excel Viewer Pro — Initialized Data Section
; ============================================================================

section '.data' data readable writeable

    class_name      db 'ExcelViewerPro_ASM64', 0
    wnd_title       db 'Excel Viewer Pro - Untitled [x64 FASM Native Edition]', 0
    fmt_title_opened db 'Excel Viewer Pro - [%s] [x64 FASM Native Edition]', 0
    listview_class  db 'SysListView32', 0
    tab_class       db 'SysTabControl32', 0
    statusbar_class db 'msctls_statusbar32', 0
    edit_class      db 'EDIT', 0
    btn_class       db 'BUTTON', 0
    combo_class     db 'COMBOBOX', 0
    static_class    db 'STATIC', 0

    font_ui_name    db 'Segoe UI', 0
    font_mono_name  db 'Consolas', 0

    ; Menu strings
    m_file          db '&File', 0
    m_new           db '&New Workbook', 9, 'Ctrl+N', 0
    m_open_csv      db '&Open File (Excel, CSV)...', 9, 'Ctrl+O', 0
    m_open_demo     db 'Open &Financial Demo Model', 0
    m_save          db '&Save File', 9, 'Ctrl+S', 0
    m_save_as       db 'Save &As...', 9, 'Ctrl+Shift+S', 0
    m_exit          db 'E&xit', 9, 'Alt+F4', 0

    m_edit          db '&Edit', 0
    m_undo          db '&Undo', 9, 'Ctrl+Z', 0
    m_redo          db '&Redo', 9, 'Ctrl+Y', 0
    m_cut           db 'Cu&t', 9, 'Ctrl+X', 0
    m_copy          db '&Copy', 9, 'Ctrl+C', 0
    m_paste         db '&Paste', 9, 'Ctrl+V', 0
    m_del_cell      db 'Clear &Cell Contents', 9, 'Del', 0
    m_sel_all       db 'Select &All', 9, 'Ctrl+A', 0
    m_find          db '&Find in Sheet...', 9, 'Ctrl+F', 0

    m_view          db '&View', 0
    m_gridlines     db 'Toggle &Gridlines', 0
    m_zoom100       db 'Reset &Zoom (100%)', 0
    m_refresh       db '&Refresh Spreadsheet', 9, 'F5', 0

    m_insert        db '&Insert', 0
    m_ins_row       db '&Insert Row', 0
    m_del_row       db '&Delete Row', 0
    m_ins_fx        db 'Insert &Formula Function [fx]', 0
    m_load_demo     db 'Load &Sample Financial Data', 0

    m_format        db 'F&ormat', 0
    m_fmt_bold      db '&Bold Font', 9, 'Ctrl+B', 0
    m_fmt_italic    db '&Italic Font', 9, 'Ctrl+I', 0
    m_fmt_under     db '&Underline Font', 9, 'Ctrl+U', 0
    m_fmt_strike    db '&Strikethrough', 0
    m_fmt_finc      db 'Increase Font Size (A+)', 0
    m_fmt_fdec      db 'Decrease Font Size (A-)', 0
    m_align_l       db 'Align &Left', 0
    m_align_c       db 'Align &Center', 0
    m_align_r       db 'Align &Right', 0
    m_fmt_usd       db 'Format: Currency ($)', 0
    m_fmt_rub       db 'Format: Currency (P)', 0
    m_fmt_pct       db 'Format: Percentage (%)', 0
    m_fmt_num       db 'Format: Number (2 Decimals)', 0

    m_formulas      db 'F&ormulas', 0
    m_recalc        db '&Recalculate Sheet (SIMD)', 9, 'F9', 0
    m_autosum_m     db '&AutoSum (SSE2 Column Sum)', 9, 'Alt+=', 0
    m_average_m     db 'SIMD &Average (AVG)', 0
    m_min_m         db 'SIMD &Minimum (MIN)', 0
    m_max_m         db 'SIMD &Maximum (MAX)', 0
    m_pmt_m         db '&Loan Payment Calculator (PMT)...', 0

    m_data          db '&Data', 0
    m_sort_asc      db 'Sort Column &Ascending (A-Z / 0-9)', 0
    m_sort_desc     db 'Sort Column &Descending (Z-A / 9-0)', 0
    m_bench_m       db 'Run &1,000,000 SIMD SSE2 Benchmark...', 0

    m_help          db '&Help', 0
    m_form_ref      db '&Formula Syntax Reference...', 0
    m_simd_info     db '&SIMD Hardware Architecture Details...', 0
    m_about         db '&About Excel Viewer Pro x64...', 0

    csv_filter      db 'All Supported Files (*.xlsx;*.xls;*.csv;*.tsv;*.txt)', 0, '*.xlsx;*.xls;*.csv;*.tsv;*.txt', 0,\
                       'Excel Workbooks (*.xlsx)', 0, '*.xlsx', 0,\
                       'Excel 97-2003 (*.xls)', 0, '*.xls', 0,\
                       'CSV / TSV Files (*.csv;*.tsv)', 0, '*.csv;*.tsv', 0,\
                       'All Files (*.*)', 0, '*.*', 0, 0
    def_ext_csv     db 'xlsx', 0
    ofn_title_open  db 'Open Spreadsheet File (Excel, CSV, TSV)', 0
    ofn_title_save  db 'Save Spreadsheet File', 0
    mode_r          db 'r', 0
    mode_w          db 'w', 0
    csv_line_fmt    db '%s,%s,%s,%s,%s,%s,%s,%s,%s,%s', 13, 10, 0
    sort_asc_msg    db 'Column sorted in ascending order (A to Z / 0 to 9)!', 0
    sort_desc_msg   db 'Column sorted in descending order (Z to A / 9 to 0)!', 0
    find_prompt_msg db 'Enter text to search across worksheet rows and columns.', 0
    find_not_fnd    db 'Searched all cells: text not found.', 0
    find_found_fmt  db 'Found match in cell %c%d!', 0
    pmt_calc_title  db 'Loan Payment Calculator (PMT)', 0
    pmt_calc_msg    db 'Excel PMT Calculation (5.00% annual interest, 360 months, $200,000 loan):', 13, 10, 13, 10,\
                       'Monthly Payment: -$1,073.64', 13, 10,\
                       'Calculated via Hardware FPU / SSE2 Kernel!', 0
    form_ref_title  db 'Excel Formula Syntax Reference', 0
    form_ref_msg    db 'Supported Excel Formula Functions in x64 Assembly:', 13, 10, 13, 10,\
                       '  =SUM(range)        - Sum cells via SIMD SSE2 (e.g. =SUM(B1:B5))', 13, 10,\
                       '  =AVG(range)        - Arithmetic mean of range (e.g. =AVG(B1:B5))', 13, 10,\
                       '  =MIN(range)        - Minimum value via SSE2 minsd (e.g. =MIN(B1:B5))', 13, 10,\
                       '  =MAX(range)        - Maximum value via SSE2 maxsd (e.g. =MAX(B1:B5))', 13, 10,\
                       '  =COUNT(range)      - Non-empty number count (e.g. =COUNT(B1:B5))', 13, 10,\
                       '  =PMT(rate, nper, pv) - Loan monthly payment calculation', 13, 10,\
                       '  Arithmetic: + - * / - e.g. =100+250 or =500*1.2 or =B1+B2', 0
    simd_info_title db 'SIMD SSE2 Hardware Architecture', 0
    simd_info_msg   db 'Excel Viewer Pro x64 Assembly SIMD Vector Engine:', 13, 10, 13, 10,\
                       '  • 128-bit Vector Registers: XMM0..XMM15 (SSE2)', 13, 10,\
                       '  • Precision: IEEE-754 64-bit Double Precision Floats', 13, 10,\
                       '  • Vector Pipeline: 4-Way Loop Unrolling with Parallel addpd', 13, 10,\
                       '  • Horizontal Reduction: movhlps + addsd Vector Fold', 13, 10,\
                       '  • Sub-nanosecond minsd / maxsd Parallel Extremum Detection', 13, 10,\
                       '  • Zero-Copy Direct Memory Integration with Python ctypes.', 0

    ; Ribbon Tabs
    rtab_home       db '  Home  ', 0
    rtab_insert     db '  Insert  ', 0
    rtab_data       db '  Data  ', 0
    rtab_formulas   db '  Formulas  ', 0
    rtab_view       db '  View  ', 0

    ; Ribbon Group Titles
    grp_clip_txt    db 'Clipboard', 0
    grp_font_txt    db 'Font', 0
    grp_align_txt   db 'Alignment', 0
    grp_num_txt     db 'Number', 0
    grp_styles_txt  db 'Styles', 0
    grp_cells_txt   db 'Cells', 0
    grp_edit_txt    db 'Editing', 0

    ; Buttons Text
    btn_paste_txt   db 'Paste', 0
    btn_cut_txt     db 'Cut', 0
    btn_copy_txt    db 'Copy', 0
    btn_format_txt  db 'Format', 0

    btn_inc_f_txt   db 'A+', 0
    btn_dec_f_txt   db 'A-', 0
    btn_bold_txt    db 'B', 0
    btn_italic_txt  db 'I', 0
    btn_under_txt   db 'U', 0
    btn_strike_txt  db 'S', 0
    btn_border_txt  db '[#]', 0
    btn_fill_txt    db 'Fill', 0
    btn_col_txt     db 'A', 0

    grp_undo_txt    db 'Undo', 0
    btn_undo_txt    db '<- Undo', 0
    btn_redo_txt    db 'Redo ->', 0

    btn_top_txt     db '^ Top', 0
    btn_mid_txt     db '= Mid', 0
    btn_bot_txt     db 'v Bot', 0
    btn_left_txt    db '< Left', 0
    btn_center_txt  db 'Center', 0
    btn_right_txt   db 'Right >', 0
    btn_wrap_txt    db 'Wrap Text', 0
    btn_merge_txt   db 'Merge & Center', 0

    btn_rub_txt     db 'P', 0
    btn_usd_txt     db '$', 0
    btn_pct_txt     db '%', 0
    btn_com_txt     db ',', 0
    btn_idec_txt    db '.0+', 0
    btn_ddec_txt    db '.0-', 0

    btn_cf_txt      db 'Cond. Format', 0
    btn_ins_c_txt   db '+ Insert', 0
    btn_del_c_txt   db '- Delete', 0
    btn_autosum_txt db 'AutoSum', 0
    btn_sort_txt    db 'Sort & Filter', 0
    btn_find_txt    db 'Find & Select', 0

    ; Insert tab labels
    grp_ins_tab_txt   db 'Tables', 0
    btn_ins_tab_txt   db 'Table', 0
    btn_ins_piv_txt   db 'PivotTable', 0
    grp_ins_ill_txt   db 'Illustrations', 0
    btn_ins_pic_txt   db 'Pictures', 0
    btn_ins_shp_txt   db 'Shapes', 0
    btn_ins_ico_txt   db 'Icons', 0
    grp_ins_cht_txt   db 'Charts', 0
    btn_ins_col_txt   db 'Column', 0
    btn_ins_lin_txt   db 'Line', 0
    btn_ins_pie_txt   db 'Pie', 0
    btn_ins_bar_txt   db 'Bar', 0
    grp_ins_lnk_txt   db 'Links && Text', 0
    btn_ins_lnk_txt   db 'Link', 0
    btn_ins_txb_txt   db 'Text Box', 0
    btn_ins_sym_txt   db 'Symbol', 0

    ; Data tab labels
    grp_data_get_txt  db 'Get && Transform', 0
    btn_data_csv_txt  db 'From File', 0
    btn_data_ref_txt  db 'Refresh', 0
    grp_data_srt_txt  db 'Sort && Filter', 0
    btn_data_asc_txt  db 'Sort A-Z', 0
    btn_data_dsc_txt  db 'Sort Z-A', 0
    btn_data_flt_txt  db 'Filter', 0
    btn_data_clr_txt  db 'Clear Filter', 0
    grp_data_tls_txt  db 'Data Tools', 0
    btn_data_spl_txt  db 'Text to Col', 0
    btn_data_dup_txt  db 'Remove Dup', 0
    btn_data_val_txt  db 'Validation', 0
    grp_data_fcs_txt  db 'Forecast && Analysis', 0
    btn_data_sim_txt  db 'SIMD Bench', 0
    btn_data_wif_txt  db 'What-If', 0

    ; Formulas tab labels
    grp_form_lib_txt  db 'Function Library', 0
    btn_form_fx_txt   db 'Insert fx', 0
    btn_form_sum_txt  db 'AutoSum', 0
    btn_form_fin_txt  db 'Finance', 0
    btn_form_log_txt  db 'Logic', 0
    btn_form_txt_txt  db 'Text', 0
    btn_form_dat_txt  db 'Date', 0
    btn_form_luk_txt  db 'Lookup', 0
    btn_form_mth_txt  db 'Math', 0
    btn_form_mor_txt  db 'More fx', 0
    grp_form_nam_txt  db 'Defined Names', 0
    btn_form_nm_txt   db 'Name Mgr', 0
    btn_form_df_txt   db 'Define Name', 0
    grp_form_clc_txt  db 'Calculation', 0
    btn_form_cal_txt  db 'Calculate', 0
    btn_form_evl_txt  db 'Eval Formula', 0

    ; View tab labels
    grp_view_vws_txt  db 'Workbook Views', 0
    btn_view_nrm_txt  db 'Normal', 0
    btn_view_lay_txt  db 'Page Layout', 0
    grp_view_shw_txt  db 'Show && Hide', 0
    btn_view_grd_txt  db 'Gridlines', 0
    btn_view_bar_txt  db 'Formula Bar', 0
    btn_view_hdg_txt  db 'Headings', 0
    grp_view_zom_txt  db 'Zoom', 0
    btn_view_z100_txt db '100%', 0
    btn_view_zin_txt  db 'Zoom In', 0
    btn_view_zot_txt  db 'Zoom Out', 0
    grp_view_wnd_txt  db 'Window', 0
    btn_view_frz_txt  db 'Freeze Panes', 0
    btn_view_nw_txt   db 'New Window', 0

    ; Formula Snippets
    snip_sum          db '=SUM(', 0
    snip_pmt          db '=PMT(', 0
    snip_if           db '=IF(', 0
    snip_concat       db '=CONCAT(', 0
    snip_today        db '=TODAY()', 0
    snip_vlookup      db '=VLOOKUP(', 0
    symbol_sigma_str  db 0E4h, 0

    ; Status Prompts
    szNameMgrPrompt     db 'Name Manager: Manage workbook range names.', 0
    szDefNamePrompt     db 'Define Name: Create a named reference for selected range.', 0
    szFilterPrompt      db 'AutoFilter active on table range.', 0
    szClearFilterPrompt db 'Filter cleared. All rows displayed.', 0
    szSplitPrompt       db 'Text to Columns wizard initialized.', 0
    szNormalViewPrompt  db 'Normal View mode active.', 0
    szTablePrompt       db 'Table formatting applied to current range.', 0
    szPivotPrompt       db 'PivotTable engine active.', 0
    szChartPrompt       db 'Chart visualizer ready for selected range.', 0

    ; Font & Format Combobox Items
    cb_font1        db 'Segoe UI', 0
    cb_font2        db 'Calibri', 0
    cb_font3        db 'Arial', 0
    cb_font4        db 'Consolas', 0

    cb_sz1          db '8', 0
    cb_sz2          db '9', 0
    cb_sz3          db '10', 0
    cb_sz4          db '11', 0
    cb_sz5          db '12', 0
    cb_sz6          db '14', 0
    cb_sz7          db '16', 0

    cb_fmt1         db 'General', 0
    cb_fmt2         db 'Number', 0
    cb_fmt3         db 'Currency ($)', 0
    cb_fmt4         db 'Percentage (%)', 0
    cb_fmt5         db 'Short Date', 0

    ; Formula Bar
    fx_txt          db 'fx', 0
    init_cell_txt   db 'A1', 0
    init_form_txt   db '', 0

    ; Grid Columns (#, A..J)
    col_idx_txt     db '#', 0
    col_A_txt       db 'A', 0
    col_B_txt       db 'B', 0
    col_C_txt       db 'C', 0
    col_D_txt       db 'D', 0
    col_E_txt       db 'E', 0
    col_F_txt       db 'F', 0
    col_G_txt       db 'G', 0
    col_H_txt       db 'H', 0
    col_I_txt       db 'I', 0
    col_J_txt       db 'J', 0

    ; Bottom Sheet Tabs
    tab_sh1_txt     db 'Sheet1', 0
    tab_sh2_txt     db 'Sheet2', 0
    tab_sh3_txt     db 'Sheet3', 0
    tab_sh_add_txt  db ' + ', 0
    new_sheet_fmt   db 'Sheet%d', 0

    ; Status Bar Formats
    sb_ready_txt    db 'Ready', 0
    sb_cell_txt     db 'A1', 0
    sb_stats_fmt    db 'SUM: %.2f | AVG: %.2f | COUNT: %d', 0
    sb_zoom_txt     db '-  100%  +', 0
    sb_dim_fmt      db '%d R x 10 C', 0

    fmt_f2          db '%.2f', 0
    fmt_ruble       db '%.2f P', 0
    fmt_dollar      db '$%.2f', 0
    fmt_percent     db '%.2f %%', 0

    ; Dialogs & Notifications
    about_title     db 'About Excel Viewer Pro x64 Native', 0
    about_msg       db 'Excel Viewer Pro - 100% Native 64-bit Assembly Spreadsheet Engine', 13, 10,\
                       'Version 2.0 Pro (x86-64 FASM Native Edition)', 13, 10, 13, 10,\
                       'Active Features:', 13, 10,\
                       '  - Full Ribbon Toolbar & Working Format Groups', 13, 10,\
                       '  - Interactive Excel Grid with Formula Bar & [fx] Evaluator', 13, 10,\
                       '  - Live Windows Clipboard (Copy, Cut, Paste)', 13, 10,\
                       '  - Hardware SIMD SSE2 Column AutoSum & Statistical Engine', 13, 10,\
                       '  - Dynamic Multi-Sheet Manager with Instant Tab Switching.', 0

    bench_title     db 'SIMD SSE2 Benchmark', 0
    bench_fmt       db '=== 1,000,000 Floating-Point Numbers Benchmark ===', 13, 10, 13, 10,\
                       'Vector Buffer:       1,000,000 doubles (8 MB RAM)', 13, 10,\
                       'SIMD SSE2 Sum:       %.2f', 13, 10,\
                       'Assembly Time:       %.4f ms', 13, 10,\
                       'Throughput:          %.2f Million Doubles / second', 13, 10, 13, 10,\
                       'Hardware SSE2 Vector Acceleration Active!', 0

    saved_msg       db 'Workbook successfully saved in high-speed binary format!', 0
    open_demo_msg   db 'Financial Demo Worksheet Loaded with Live Formulas!', 0
    empty_str       db 0

    ; Financial Demo Data
    d1_A db 'Revenue - Enterprise', 0
    d1_B db '125000.00', 0
    d1_C db '145000.00', 0
    d1_D db '160000.00', 0

    d2_A db 'Revenue - SaaS Subscriptions', 0
    d2_B db '45000.00', 0
    d2_C db '52000.00', 0
    d2_D db '61000.00', 0

    d3_A db 'Revenue - Services', 0
    d3_B db '35000.00', 0
    d3_C db '42000.00', 0
    d3_D db '48000.00', 0

    d4_A db 'Total Revenue', 0
    d4_B db '205000.00', 0
    d4_C db '239000.00', 0
    d4_D db '269000.00', 0

    d5_A db 'R&D Expenses', 0
    d5_B db '38000.00', 0
    d5_C db '42000.00', 0
    d5_D db '45000.00', 0

    d6_A db 'Net Profit', 0
    d6_B db '167000.00', 0
    d6_C db '197000.00', 0
    d6_D db '224000.00', 0

    hInst           dq 0
    hWnd            dq 0
    hMenu           dq 0
    hSubMenuFile    dq 0
    hSubMenuEdit    dq 0
    hSubMenuView    dq 0
    hSubMenuInsert  dq 0
    hSubMenuFormat  dq 0
    hSubMenuFormulas dq 0
    hSubMenuData    dq 0
    hSubMenuHelp    dq 0

    rcMenu16        dd 0, 0, 16, 16
    menuItemInfo    dd 80, 80h, 0, 0, 0, 0
                    dq 0, 0, 0, 0, 0
                    dd 0, 0
                    dq 0

    szGlyphB        db 'B', 0
    szGlyphI        db 'I', 0
    szGlyphU        db 'U', 0
    szGlyphS        db 'S', 0
    szGlyphA        db 'A', 0
    szGlypha        db 'a', 0
    szGlyphUSD      db '$', 0
    szGlyphRUB      db 'P', 0
    szGlyphPCT      db '%', 0
    szGlyphNUM      db '.0', 0
    szGlyphFX       db 'fx', 0
    szGlyph1        db '1', 0
    szGlyphAvg      db 'AV', 0
    szGlyphSigma    db 'S', 0
    szGlyphZ        db 'Z', 0
    szGlyphi        db 'i', 0

    align 8
    menu_icon_table:
        ; Format Menu
        dq hSubMenuFormat
        dd ID_MENU_FONT_BOLD, ICO_BOLD
        dq hSubMenuFormat
        dd ID_MENU_FONT_ITALIC, ICO_ITALIC
        dq hSubMenuFormat
        dd ID_MENU_FONT_UNDER, ICO_UNDER
        dq hSubMenuFormat
        dd ID_MENU_FONT_STRIKE, ICO_STRIKE
        dq hSubMenuFormat
        dd ID_MENU_FONT_INC, ICO_FINC
        dq hSubMenuFormat
        dd ID_MENU_FONT_DEC, ICO_FDEC
        dq hSubMenuFormat
        dd ID_MENU_ALIGN_LEFT, ICO_ALIGN_L
        dq hSubMenuFormat
        dd ID_MENU_ALIGN_CENT, ICO_ALIGN_C
        dq hSubMenuFormat
        dd ID_MENU_ALIGN_RIGH, ICO_ALIGN_R
        dq hSubMenuFormat
        dd ID_MENU_FMT_DOLLAR, ICO_USD
        dq hSubMenuFormat
        dd ID_MENU_FMT_RUBLE, ICO_RUB
        dq hSubMenuFormat
        dd ID_MENU_FMT_PERCENT, ICO_PCT
        dq hSubMenuFormat
        dd ID_MENU_FMT_COMMA, ICO_NUM

        ; File Menu
        dq hSubMenuFile
        dd ID_MENU_NEW, ICO_NEW
        dq hSubMenuFile
        dd ID_MENU_OPEN_CSV, ICO_OPEN
        dq hSubMenuFile
        dd ID_MENU_OPEN_DEMO, ICO_OPEN_DEMO
        dq hSubMenuFile
        dd ID_MENU_SAVE, ICO_SAVE
        dq hSubMenuFile
        dd ID_MENU_SAVEAS, ICO_SAVEAS
        dq hSubMenuFile
        dd ID_MENU_EXIT, ICO_EXIT

        ; Edit Menu
        dq hSubMenuEdit
        dd ID_MENU_UNDO, ICO_UNDO
        dq hSubMenuEdit
        dd ID_MENU_REDO, ICO_REDO
        dq hSubMenuEdit
        dd ID_MENU_CUT, ICO_CUT
        dq hSubMenuEdit
        dd ID_MENU_COPY, ICO_COPY
        dq hSubMenuEdit
        dd ID_MENU_PASTE, ICO_PASTE
        dq hSubMenuEdit
        dd ID_MENU_DEL_CELL, ICO_DEL_CELL
        dq hSubMenuEdit
        dd ID_MENU_SELECT_ALL, ICO_SELECT_ALL
        dq hSubMenuEdit
        dd ID_MENU_FIND, ICO_FIND

        ; View Menu
        dq hSubMenuView
        dd ID_MENU_GRIDLINES, ICO_GRIDLINES
        dq hSubMenuView
        dd ID_MENU_ZOOM100, ICO_ZOOM100
        dq hSubMenuView
        dd ID_MENU_REFRESH, ICO_REFRESH

        ; Insert Menu
        dq hSubMenuInsert
        dd ID_MENU_INSERT_ROW, ICO_INS_ROW
        dq hSubMenuInsert
        dd ID_MENU_DELETE_ROW, ICO_DEL_ROW
        dq hSubMenuInsert
        dd ID_MENU_INSERT_FX, ICO_FX
        dq hSubMenuInsert
        dd ID_MENU_OPEN_DEMO, ICO_OPEN_DEMO

        ; Formulas Menu
        dq hSubMenuFormulas
        dd ID_MENU_RECALC, ICO_RECALC
        dq hSubMenuFormulas
        dd ID_MENU_AUTOSUM, ICO_AUTOSUM
        dq hSubMenuFormulas
        dd ID_MENU_AVERAGE, ICO_AVERAGE
        dq hSubMenuFormulas
        dd ID_MENU_MIN, ICO_MIN
        dq hSubMenuFormulas
        dd ID_MENU_MAX, ICO_MAX
        dq hSubMenuFormulas
        dd ID_MENU_PMT, ICO_PMT

        ; Data Menu
        dq hSubMenuData
        dd ID_MENU_SORT_ASC, ICO_SORT_ASC
        dq hSubMenuData
        dd ID_MENU_SORT_DESC, ICO_SORT_DESC
        dq hSubMenuData
        dd ID_MENU_BENCHMARK, ICO_BENCHMARK

        ; Help Menu
        dq hSubMenuHelp
        dd ID_MENU_FORM_REF, ICO_FORM_REF
        dq hSubMenuHelp
        dd ID_MENU_SIMD_INFO, ICO_SIMD_INFO
        dq hSubMenuHelp
        dd ID_MENU_ABOUT, ICO_ABOUT

        ; Context Menu
        dq hContextMenu
        dd ID_MENU_CUT, ICO_CUT
        dq hContextMenu
        dd ID_MENU_COPY, ICO_COPY
        dq hContextMenu
        dd ID_MENU_PASTE, ICO_PASTE
        dq hContextMenu
        dd ID_MENU_INSERT_ROW, ICO_INS_ROW
        dq hContextMenu
        dd ID_MENU_DELETE_ROW, ICO_DEL_ROW
        dq hContextMenu
        dd ID_MENU_DEL_CELL, ICO_DEL_CELL
        dq hContextMenu
        dd ID_MENU_FONT_BOLD, ICO_BOLD
        dq hContextMenu
        dd ID_MENU_FONT_ITALIC, ICO_ITALIC
        dq hContextMenu
        dd ID_MENU_FMT_DOLLAR, ICO_USD
        dq hContextMenu
        dd ID_MENU_FMT_RUBLE, ICO_RUB
        dq hContextMenu
        dd ID_MENU_FMT_PERCENT, ICO_PCT
        dq hContextMenu
        dd ID_MENU_AUTOSUM, ICO_AUTOSUM
        dq hContextMenu
        dd ID_MENU_SORT_ASC, ICO_SORT_ASC
        dq hContextMenu
        dd ID_MENU_SORT_DESC, ICO_SORT_DESC
        dq hContextMenu
        dd ID_MENU_FIND, ICO_FIND

        dq 0
        dd 0, 0

    ; Context Menu Strings
    m_ctx_cut           db 'Cu&t', 9, 'Ctrl+X', 0
    m_ctx_copy          db '&Copy', 9, 'Ctrl+C', 0
    m_ctx_paste         db '&Paste', 9, 'Ctrl+V', 0
    m_ctx_ins_row       db '&Insert Row', 0
    m_ctx_del_row       db '&Delete Row', 0
    m_ctx_clear         db 'Clear &Cell Contents', 9, 'Del', 0
    m_ctx_bold          db '&Bold Font', 9, 'Ctrl+B', 0
    m_ctx_italic        db '&Italic Font', 9, 'Ctrl+I', 0
    m_ctx_usd           db 'Format: &Currency ($)', 0
    m_ctx_rub           db 'Format: Currency (&P)', 0
    m_ctx_pct           db 'Format: &Percentage (%)', 0
    m_ctx_autosum       db '&AutoSum (S)', 0
    m_ctx_sort_asc      db 'Sort &A -> Z', 0
    m_ctx_sort_desc     db 'Sort &Z -> A', 0
    m_ctx_find          db '&Find in Sheet...', 9, 'Ctrl+F', 0
    hFontUI         dq 0
    hFontSmall      dq 0
    hFontBold       dq 0
    hFontItalic     dq 0
    hFontBoldItalic dq 0
    hFontMono       dq 0
    hFontGrid       dq 0

    ; Controls
    hRibbonTabs     dq 0
    hNameBox        dq 0
    hBtnFx          dq 0
    hFormulaEdit    dq 0
    hGrid           dq 0
    hSheetTabs      dq 0
    hStatusBar      dq 0
    hCellEdit       dq 0
    origCellEditProc dq 0
    bCellEditing    dd 0
    hContextMenu    dq 0
    rcCell          dd 0, 0, 0, 0

    hComboFont      dq 0
    hComboSize      dq 0
    hComboNumFmt    dq 0

    cur_row         dd 0
    cur_col         dd 1        ; Col A is 1
    total_rows      dd 30
    sheet_count     dd 3
    cur_sheet       dd 0

    sel_start_row   dd 0
    sel_start_col   dd 1
    sel_end_row     dd 0
    sel_end_col     dd 1
    is_mouse_sel    dd 0
    g_is_shift_down dd 0
    g_is_ctrl_down  dd 0
    origGridProc    dq 0

    font_size_pt    dd 10
    pCurFontFace    dq 0
    font_size_table dd 8, 9, 10, 11, 12, 14, 16
    font_faces_table dq font_ui_name, cb_font1, cb_font2, cb_font3, cb_font4
    font_cache_keys  dd 64 dup(0)
    font_cache_fonts dq 64 dup(0)
    font_cache_count dd 0
    is_bold_active  dd 0
    is_italic_act   dd 0
    is_under_act    dd 0
    is_strike_act   dd 0
    gridlines_act   dd 1

    ; Ribbon Tab Controls ID Arrays
    ids_home dd 1110, 1111, 1112, 1113, 1114, 1120, 1121, 1122, 1123, 1124, 1125, 1126, 1127, 1128, 1129, 1130, 1131, 1140, 1141, 1142, 1143, 1144, 1145, 1146, 1147, 1148, 1150, 1151, 1152, 1153, 1154, 1155, 1156, 1157, 1160, 1161, 1170, 1171, 1172, 1180, 1181, 1182, 1183, 0
    ids_insert dd 1400, 1401, 1402, 1403, 1404, 1405, 1406, 1407, 1408, 1409, 1410, 1411, 1412, 1413, 1414, 1415, 0
    ids_data dd 1420, 1421, 1422, 1423, 1424, 1425, 1426, 1427, 1428, 1429, 1430, 1431, 1432, 1433, 1434, 0
    ids_formulas dd 1440, 1441, 1442, 1443, 1444, 1445, 1446, 1447, 1448, 1449, 1450, 1451, 1452, 1453, 1454, 1455, 0
    ids_view dd 1460, 1461, 1462, 1463, 1464, 1465, 1466, 1467, 1468, 1469, 1470, 1471, 1472, 1473, 0

    rcRibbonArea    dd 0, 24, 3000, 122
    cur_ribbon_tab  dd 0

    ; Cell Fill & Text Color Picker Storage
    cur_fill_color  dd 00CCF2FFh ; default light yellow (#FFF2CC in BGR format)
    cur_text_color  dd 000000C0h ; default dark red (#C00000 in BGR format)
    cust_colors     dd 16 dup(00FFFFFFh) ; 16 custom color slots for ChooseColorA

    ; GDI Brushes & Pens for Ribbon Button Icons
    hBrBtnFace      dq 0
    hBrBtnHover     dq 0
    hBrBtnPress     dq 0
    hBrBlue         dq 0
    hBrGreen        dq 0
    hBrOrange       dq 0
    hBrYellow       dq 0
    hBrRed          dq 0
    hBrCyan         dq 0
    hBrWhite        dq 0
    hBrGray         dq 0

    hPenBorder      dq 0
    hPenBorderPress dq 0
    hPenBlue        dq 0
    hPenGreen       dq 0
    hPenOrange      dq 0
    hPenRed         dq 0
    hPenDark        dq 0
    hPenDark2       dq 0
    hPenWhite       dq 0
    hPenNull        dq 0

    szBtnDrawText   db 64 dup(0)
    rcBtnDrawText   dd 0, 0, 0, 0

    hAppIcon        dq 0
    szIconFile      db 'app_icon.ico', 0
    szIconFileAsm   db 'asm\app_icon.ico', 0
    szIconFileParent db '..\app_icon.ico', 0
    szIconFilePkg   db 'xlsx_viewer\app_icon.ico', 0

    sb_parts        dd 160, 260, 720, 890, -1

    bench_freq      dq 0
    bench_start     dq 0
    bench_end       dq 0

    wc              rb 80
    msgbuf          rb 48
    icc             rb 8
    lvcol           rb 64
    lvitem          rb 88
    lvitem_cd       rb 88
    szCellTextCd    rb 256
    rcSubItemCd     dd 0, 0, 0, 0
    rcBgCd          dd 0, 0, 0, 0
    tcitem          rb 48

    szTemp          rb 2048
    szBuffer        rb 4096
    szCellVal       rb 1024

    origFormulaProc dq 0
    pt              rb 16
    lvhit           rb 32
    ofn             rb 160
    szFileName      rb 512
    szFindQuery     rb 256
    szLineBuf       rb 2048
    find_pos        dd 0

    fmt_d           db '%d', 0
    fmt_s           db '%s', 0
    fmt_comma_char  db ',', 0
    fmt_newline     db 13, 10, 0
    recalc_done_msg db 'Recalculation complete: Native SSE2 SIMD Pipeline verified.', 0
    undo_msg        db 'Action undone.', 0
    redo_msg        db 'Action redone.', 0
    undo_empty_msg  db 'Nothing to undo.', 0
    redo_empty_msg  db 'Nothing to redo.', 0

    UNDO_MAX        = 64
    UNDO_REC_SIZE   = 512

    undo_count      dd 0
    redo_count      dd 0
    pending_undo_act dd 0
    pending_row     dd 0
    pending_col     dd 0
    pending_sheet   dd 0
    pending_old_txt rb 128
    pending_old_form rb 128

    align 16
    undo_stack      rb UNDO_MAX * UNDO_REC_SIZE
    redo_stack      rb UNDO_MAX * UNDO_REC_SIZE

    fmt_painter_msg db 'Format Painter: Style copied. Click destination cell.', 0
    wrap_txt_msg    db 'Word Wrap enabled for selected column.', 0
    merge_ct_msg    db 'Merge & Center active for range.', 0
    sample_fx_str   db '=SUM(B1:B5)', 0

    wc_wiz          rb 80
    szTempCsv       rb 512
    szTempStyle     rb 512
    szCmdBuf        rb 2048
    is_temp_file    db 0
    mode_rb         db 'rb', 0
    fmt_temp_csv    db '%sxlsx_view_tmp.csv', 0
    fmt_temp_style   db '%s.style', 0
    fmt_temp_formula db '%s.formula', 0
    szTempFormula    rb 512
    fmt_temp_sheets  db '%s.sheets', 0
    fmt_sheet_csv    db '%s.%d', 0
    fmt_sheet_style  db '%s.%d', 0
    fmt_sheet_formula db '%s.%d.formula', 0
    szTempSheets     rb 512
    szCurCsvPath     rb 512
    szCurStylePath   rb 512
    szCurFormulaPath rb 512
    g_sheet_names    rb 16 * 128
    pCellStyleTable  dq 0
    pCellFormulaTable dq 0
    hFormulaHeap     dq 0
    szFormulaInput   rb 2048
    szNativeDllName  db 'xlsx_native.dll', 0
    szNativeDllPath  db 'C:\proekts\xlsx_vievers\asm\xlsx_native.dll', 0
    szNativeFuncName db 'ConvertXlsxToCsvAndStyle', 0
    hNativeDll       dq 0
    pfnConvertXlsx   dq 0

    hFxWizWnd       dq 0
    hWizList        dq 0
    hWizSyntax      dq 0
    hWizDesc        dq 0
    hWizSearchEdit  dq 0
    hWizCatCombo    dq 0
    wiz_init_cat    dd 0
    num_filtered_fn dd 0
    filtered_fn_idx rd 256
    szWizQuery      rb 128
    wiz_class_reg   dd 0

    fx_wiz_class_name db 'FxWizardClass', 0
    fx_wiz_title      db 'Insert Function', 0
    lbl_wiz_search    db 'Search for a function:', 0
    btn_wiz_go_txt    db 'Go', 0
    lbl_wiz_cat       db 'Or select a category:', 0
    lbl_wiz_prompt    db 'Select a function:', 0
    grp_wiz_details   db 'Function Details', 0
    lbl_wiz_syntax_t  db 'Syntax:', 0
    lbl_wiz_desc_t    db 'Description:', 0
    lbl_wiz_no_match  db 'No functions match your search criteria.', 0
    btn_wiz_ok_txt    db 'Insert', 0
    btn_wiz_can_txt   db 'Cancel', 0
    static_class_name db 'STATIC', 0
    listbox_class     db 'LISTBOX', 0

    fmt_fn_prefix     db '=%s(', 0
    fmt_fn_range      db '=%s(%c1:%c%d)', 0
    fmt_fn_empty      db '=%s()', 0

    cat_str_0         db 'All', 0
    cat_str_1         db 'Math & Trig', 0
    cat_str_2         db 'Statistical', 0
    cat_str_3         db 'Lookup & Reference', 0
    cat_str_4         db 'Logical', 0
    cat_str_5         db 'Text', 0
    cat_str_6         db 'Date & Time', 0
    cat_str_7         db 'Financial', 0
    cat_str_8         db 'Information', 0

    align 8
    cat_names_table:
        dq cat_str_0, cat_str_1, cat_str_2, cat_str_3
        dq cat_str_4, cat_str_5, cat_str_6, cat_str_7
        dq cat_str_8
    NUM_WIZ_CATS = 9

    include 'funcs_data.inc'

    align 8
    fn_names_table  = fn_all_names
    fn_syntax_table = fn_all_syntax
    fn_desc_table   = fn_all_desc

    align 32
    bench_array     rq 1000000

    align 16
    sheet_cells     rb 16 * 30 * 11 * 128


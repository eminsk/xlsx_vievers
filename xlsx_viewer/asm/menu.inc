; ============================================================================
; Excel Viewer Pro — Menu Bar & GDI Icons Engine
; ============================================================================



; =============================================================================
; Build Window Menu Bar
; =============================================================================
BuildAppMenu:
    push rbp
    mov rbp, rsp
    sub rsp, 40h

    call [CreateMenu]
    mov [hMenu], rax

    ; 1. File Menu
    call [CreatePopupMenu]
    mov [hSubMenuFile], rax
    mov rbx, rax
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_NEW
    lea r9, [m_new]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_OPEN_CSV
    lea r9, [m_open_csv]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_OPEN_DEMO
    lea r9, [m_open_demo]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_SAVE
    lea r9, [m_save]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_SAVEAS
    lea r9, [m_save_as]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_EXIT
    lea r9, [m_exit]
    call [AppendMenuA]
    mov rcx, [hMenu]
    mov edx, MF_POPUP
    mov r8, rbx
    lea r9, [m_file]
    call [AppendMenuA]

    ; 2. Edit Menu
    call [CreatePopupMenu]
    mov [hSubMenuEdit], rax
    mov rbx, rax
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_UNDO
    lea r9, [m_undo]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_REDO
    lea r9, [m_redo]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_CUT
    lea r9, [m_cut]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_COPY
    lea r9, [m_copy]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_PASTE
    lea r9, [m_paste]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_DEL_CELL
    lea r9, [m_del_cell]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_SELECT_ALL
    lea r9, [m_sel_all]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FIND
    lea r9, [m_find]
    call [AppendMenuA]
    mov rcx, [hMenu]
    mov edx, MF_POPUP
    mov r8, rbx
    lea r9, [m_edit]
    call [AppendMenuA]

    ; 3. View Menu
    call [CreatePopupMenu]
    mov [hSubMenuView], rax
    mov rbx, rax
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_GRIDLINES
    lea r9, [m_gridlines]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_ZOOM100
    lea r9, [m_zoom100]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_REFRESH
    lea r9, [m_refresh]
    call [AppendMenuA]
    mov rcx, [hMenu]
    mov edx, MF_POPUP
    mov r8, rbx
    lea r9, [m_view]
    call [AppendMenuA]

    ; 4. Insert Menu
    call [CreatePopupMenu]
    mov [hSubMenuInsert], rax
    mov rbx, rax
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_INSERT_ROW
    lea r9, [m_ins_row]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_DELETE_ROW
    lea r9, [m_del_row]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_INSERT_FX
    lea r9, [m_ins_fx]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_OPEN_DEMO
    lea r9, [m_load_demo]
    call [AppendMenuA]
    mov rcx, [hMenu]
    mov edx, MF_POPUP
    mov r8, rbx
    lea r9, [m_insert]
    call [AppendMenuA]

    ; 5. Format Menu
    call [CreatePopupMenu]
    mov [hSubMenuFormat], rax
    mov rbx, rax
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FONT_BOLD
    lea r9, [m_fmt_bold]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FONT_ITALIC
    lea r9, [m_fmt_italic]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FONT_UNDER
    lea r9, [m_fmt_under]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FONT_STRIKE
    lea r9, [m_fmt_strike]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FONT_INC
    lea r9, [m_fmt_finc]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FONT_DEC
    lea r9, [m_fmt_fdec]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_ALIGN_LEFT
    lea r9, [m_align_l]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_ALIGN_CENT
    lea r9, [m_align_c]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_ALIGN_RIGH
    lea r9, [m_align_r]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FMT_DOLLAR
    lea r9, [m_fmt_usd]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FMT_RUBLE
    lea r9, [m_fmt_rub]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FMT_PERCENT
    lea r9, [m_fmt_pct]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FMT_COMMA
    lea r9, [m_fmt_num]
    call [AppendMenuA]
    mov rcx, [hMenu]
    mov edx, MF_POPUP
    mov r8, rbx
    lea r9, [m_format]
    call [AppendMenuA]

    ; 6. Formulas Menu
    call [CreatePopupMenu]
    mov [hSubMenuFormulas], rax
    mov rbx, rax
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_RECALC
    lea r9, [m_recalc]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_AUTOSUM
    lea r9, [m_autosum_m]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_AVERAGE
    lea r9, [m_average_m]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_MIN
    lea r9, [m_min_m]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_MAX
    lea r9, [m_max_m]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_PMT
    lea r9, [m_pmt_m]
    call [AppendMenuA]
    mov rcx, [hMenu]
    mov edx, MF_POPUP
    mov r8, rbx
    lea r9, [m_formulas]
    call [AppendMenuA]

    ; 7. Data Menu
    call [CreatePopupMenu]
    mov [hSubMenuData], rax
    mov rbx, rax
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_SORT_ASC
    lea r9, [m_sort_asc]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_SORT_DESC
    lea r9, [m_sort_desc]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_BENCHMARK
    lea r9, [m_bench_m]
    call [AppendMenuA]
    mov rcx, [hMenu]
    mov edx, MF_POPUP
    mov r8, rbx
    lea r9, [m_data]
    call [AppendMenuA]

    ; 8. Help Menu
    call [CreatePopupMenu]
    mov [hSubMenuHelp], rax
    mov rbx, rax
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FORM_REF
    lea r9, [m_form_ref]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_SIMD_INFO
    lea r9, [m_simd_info]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_ABOUT
    lea r9, [m_about]
    call [AppendMenuA]
    mov rcx, [hMenu]
    mov edx, MF_POPUP
    mov r8, rbx
    lea r9, [m_help]
    call [AppendMenuA]

    mov rcx, [hWnd]
    mov rdx, [hMenu]
    call [SetMenu]

    ; 9. Right-Click Context Menu for Grid
    call [CreatePopupMenu]
    mov [hContextMenu], rax
    mov rbx, rax

    ; Cut, Copy, Paste
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_CUT
    lea r9, [m_ctx_cut]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_COPY
    lea r9, [m_ctx_copy]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_PASTE
    lea r9, [m_ctx_paste]
    call [AppendMenuA]

    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]

    ; Insert Row, Delete Row, Clear Cell
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_INSERT_ROW
    lea r9, [m_ctx_ins_row]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_DELETE_ROW
    lea r9, [m_ctx_del_row]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_DEL_CELL
    lea r9, [m_ctx_clear]
    call [AppendMenuA]

    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]

    ; Formatting
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FONT_BOLD
    lea r9, [m_ctx_bold]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FONT_ITALIC
    lea r9, [m_ctx_italic]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FMT_DOLLAR
    lea r9, [m_ctx_usd]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FMT_RUBLE
    lea r9, [m_ctx_rub]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FMT_PERCENT
    lea r9, [m_ctx_pct]
    call [AppendMenuA]

    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]

    ; Formulas & Data
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_AUTOSUM
    lea r9, [m_ctx_autosum]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_SORT_ASC
    lea r9, [m_ctx_sort_asc]
    call [AppendMenuA]
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_SORT_DESC
    lea r9, [m_ctx_sort_desc]
    call [AppendMenuA]

    mov rcx, rbx
    mov edx, MF_SEPARATOR
    xor r8d, r8d
    xor r9, r9
    call [AppendMenuA]

    ; Find
    mov rcx, rbx
    mov edx, MF_STRING
    mov r8d, ID_MENU_FIND
    lea r9, [m_ctx_find]
    call [AppendMenuA]

    call AttachMenuIcons

    leave
    ret


; =============================================================================
; AttachMenuIcons
; Dynamically creates crisp 16x16 GDI bitmaps and assigns them to menu items
; =============================================================================
AttachMenuIcons:
    push rbp
    mov rbp, rsp
    sub rsp, 80h
    mov [rbp-08h], rbx
    mov [rbp-10h], r12
    mov [rbp-18h], r13
    mov [rbp-20h], r14
    mov [rbp-28h], r15
    mov [rbp-30h], rsi

    ; 1. Get Screen DC
    xor ecx, ecx
    call [GetDC]
    mov r12, rax                ; r12 = hScreenDC

    ; 2. Create Compatible Memory DC
    mov rcx, r12
    call [CreateCompatibleDC]
    mov r13, rax                ; r13 = hMemDC

    ; 3. Create COLOR_MENU background brush
    mov ecx, 4                  ; COLOR_MENU
    call [GetSysColor]
    mov ecx, eax
    call [CreateSolidBrush]
    mov r14, rax                ; r14 = hBrMenuBg

    ; 4. Loop through menu_icon_table
    lea rsi, [menu_icon_table]

.icon_loop:
    mov rax, [rsi]              ; pointer to hSubMenu
    test rax, rax
    jz .done_loop

    mov r15, [rax]              ; r15 = hSubMenu handle
    test r15, r15
    jz .next_item

    ; Create 16x16 compatible bitmap
    mov rcx, r12
    mov edx, 16
    mov r8d, 16
    call [CreateCompatibleBitmap]
    mov rbx, rax                ; rbx = hBitmap

    ; Select bitmap into hMemDC
    mov rcx, r13
    mov rdx, rbx
    call [SelectObject]
    mov [rbp-38h], rax          ; hOldBmp

    ; Fill background with COLOR_MENU
    mov rcx, r13
    lea rdx, [rcMenu16]
    mov r8, r14
    call [FillRect]

    ; Render specific icon
    mov rcx, r13                ; hDC
    mov edx, [rsi+12]           ; icon_type
    call RenderMenuGlyph

    ; Restore old bitmap in hMemDC
    mov rcx, r13
    mov rdx, [rbp-38h]
    call [SelectObject]

    ; SetMenuItemInfoA on submenu
    lea rax, [menuItemInfo]
    mov dword [rax], 80         ; cbSize = 80
    mov dword [rax+4], 80h      ; fMask = MIIM_BITMAP (0x80)
    mov [rax+72], rbx           ; hbmpItem = rbx
    mov rcx, r15                ; hSubMenu
    mov edx, [rsi+8]            ; wID
    xor r8d, r8d                ; fByPosition = FALSE
    mov r9, rax                 ; &menuItemInfo
    call [SetMenuItemInfoA]


.next_item:
    add rsi, 16
    jmp .icon_loop

.done_loop:
    mov rcx, r14
    call [DeleteObject]

    mov rcx, r13
    call [DeleteDC]

    xor ecx, ecx
    mov rdx, r12
    call [ReleaseDC]

    mov rbx, [rbp-08h]
    mov r12, [rbp-10h]
    mov r13, [rbp-18h]
    mov r14, [rbp-20h]
    mov r15, [rbp-28h]
    mov rsi, [rbp-30h]
    leave
    ret


; =============================================================================
; RenderMenuGlyph
; Input: RCX = hDC (16x16), EDX = icon_type (ICO_*)
; =============================================================================
RenderMenuGlyph:
    push rbp
    mov rbp, rsp
    sub rsp, 40h
    mov [rbp-8], rbx
    mov [rbp-10h], r12
    mov [rbp-18h], r13

    mov r12, rcx
    mov r13d, edx

    mov rcx, r12
    mov edx, 1                  ; TRANSPARENT
    call [SetBkMode]

    cmp r13d, ICO_BOLD
    je .ico_bold
    cmp r13d, ICO_ITALIC
    je .ico_italic
    cmp r13d, ICO_UNDER
    je .ico_under
    cmp r13d, ICO_STRIKE
    je .ico_strike
    cmp r13d, ICO_FINC
    je .ico_finc
    cmp r13d, ICO_FDEC
    je .ico_fdec
    cmp r13d, ICO_ALIGN_L
    je .ico_align_l
    cmp r13d, ICO_ALIGN_C
    je .ico_align_c
    cmp r13d, ICO_ALIGN_R
    je .ico_align_r
    cmp r13d, ICO_USD
    je .ico_usd
    cmp r13d, ICO_RUB
    je .ico_rub
    cmp r13d, ICO_PCT
    je .ico_pct
    cmp r13d, ICO_NUM
    je .ico_num
    cmp r13d, ICO_NEW
    je .ico_new
    cmp r13d, ICO_OPEN
    je .ico_open
    cmp r13d, ICO_OPEN_DEMO
    je .ico_open_demo
    cmp r13d, ICO_SAVE
    je .ico_save
    cmp r13d, ICO_SAVEAS
    je .ico_saveas
    cmp r13d, ICO_EXIT
    je .ico_exit
    cmp r13d, ICO_UNDO
    je .ico_undo
    cmp r13d, ICO_REDO
    je .ico_redo
    cmp r13d, ICO_CUT
    je .ico_cut
    cmp r13d, ICO_COPY
    je .ico_copy
    cmp r13d, ICO_PASTE
    je .ico_paste
    cmp r13d, ICO_DEL_CELL
    je .ico_del_cell
    cmp r13d, ICO_SELECT_ALL
    je .ico_select_all
    cmp r13d, ICO_FIND
    je .ico_find
    cmp r13d, ICO_GRIDLINES
    je .ico_gridlines
    cmp r13d, ICO_ZOOM100
    je .ico_zoom100
    cmp r13d, ICO_REFRESH
    je .ico_refresh
    cmp r13d, ICO_INS_ROW
    je .ico_ins_row
    cmp r13d, ICO_DEL_ROW
    je .ico_del_row
    cmp r13d, ICO_FX
    je .ico_fx
    cmp r13d, ICO_RECALC
    je .ico_recalc
    cmp r13d, ICO_AUTOSUM
    je .ico_autosum
    cmp r13d, ICO_AVERAGE
    je .ico_average
    cmp r13d, ICO_MIN
    je .ico_min
    cmp r13d, ICO_MAX
    je .ico_max
    cmp r13d, ICO_PMT
    je .ico_pmt
    cmp r13d, ICO_SORT_ASC
    je .ico_sort_asc
    cmp r13d, ICO_SORT_DESC
    je .ico_sort_desc
    cmp r13d, ICO_BENCHMARK
    je .ico_benchmark
    cmp r13d, ICO_FORM_REF
    je .ico_form_ref
    cmp r13d, ICO_SIMD_INFO
    je .ico_simd_info
    cmp r13d, ICO_ABOUT
    je .ico_about
    jmp .ico_done

.ico_bold:
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]
    mov rcx, r12
    mov edx, 009A572Bh
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphB]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .ico_done

.ico_italic:
    mov rcx, r12
    mov rdx, [hFontBoldItalic]
    call [SelectObject]
    mov rcx, r12
    mov edx, 009A572Bh
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphI]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .ico_done

.ico_under:
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]
    mov rcx, r12
    mov edx, 009A572Bh
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphU]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]

    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov edx, 3
    mov r8d, 13
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 13
    mov r8d, 13
    call [LineTo]
    mov rcx, r12
    mov edx, 3
    mov r8d, 14
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 13
    mov r8d, 14
    call [LineTo]
    jmp .ico_done

.ico_strike:
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]
    mov rcx, r12
    mov edx, 009A572Bh
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphS]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]

    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 8
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 8
    call [LineTo]
    jmp .ico_done

.ico_finc:
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]
    mov rcx, r12
    mov edx, 00202020h
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphA]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]

    mov rcx, r12
    mov rdx, [hPenGreen]
    call [SelectObject]
    mov rcx, r12
    mov edx, 13
    mov r8d, 2
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 13
    mov r8d, 7
    call [LineTo]
    mov rcx, r12
    mov edx, 11
    mov r8d, 4
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 16
    mov r8d, 4
    call [LineTo]
    jmp .ico_done

.ico_fdec:
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]
    mov rcx, r12
    mov edx, 00202020h
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlypha]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]

    mov rcx, r12
    mov rdx, [hPenRed]
    call [SelectObject]
    mov rcx, r12
    mov edx, 11
    mov r8d, 4
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 15
    mov r8d, 4
    call [LineTo]
    jmp .ico_done

.ico_align_l:
    mov rcx, r12
    mov rdx, [hPenDark]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 3
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 13
    mov r8d, 3
    call [LineTo]
    mov rcx, r12
    mov edx, 2
    mov r8d, 6
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 9
    mov r8d, 6
    call [LineTo]
    mov rcx, r12
    mov edx, 2
    mov r8d, 9
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 13
    mov r8d, 9
    call [LineTo]
    mov rcx, r12
    mov edx, 2
    mov r8d, 12
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 7
    mov r8d, 12
    call [LineTo]
    jmp .ico_done

.ico_align_c:
    mov rcx, r12
    mov rdx, [hPenDark]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 3
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 3
    call [LineTo]
    mov rcx, r12
    mov edx, 4
    mov r8d, 6
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 12
    mov r8d, 6
    call [LineTo]
    mov rcx, r12
    mov edx, 2
    mov r8d, 9
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 9
    call [LineTo]
    mov rcx, r12
    mov edx, 5
    mov r8d, 12
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 11
    mov r8d, 12
    call [LineTo]
    jmp .ico_done

.ico_align_r:
    mov rcx, r12
    mov rdx, [hPenDark]
    call [SelectObject]
    mov rcx, r12
    mov edx, 3
    mov r8d, 3
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 3
    call [LineTo]
    mov rcx, r12
    mov edx, 7
    mov r8d, 6
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 6
    call [LineTo]
    mov rcx, r12
    mov edx, 3
    mov r8d, 9
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 9
    call [LineTo]
    mov rcx, r12
    mov edx, 9
    mov r8d, 12
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 12
    call [LineTo]
    jmp .ico_done

.ico_usd:
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]
    mov rcx, r12
    mov edx, 00467321h       ; Green
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphUSD]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .ico_done

.ico_rub:
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]
    mov rcx, r12
    mov edx, 009A572Bh       ; Blue
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphRUB]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]

    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov edx, 3
    mov r8d, 10
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 9
    mov r8d, 10
    call [LineTo]
    jmp .ico_done

.ico_pct:
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]
    mov rcx, r12
    mov edx, 000078D7h       ; Orange
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphPCT]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .ico_done

.ico_num:
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]
    mov rcx, r12
    mov edx, 00505050h
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphNUM]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .ico_done

.ico_new:
    mov rcx, r12
    mov rdx, [hPenDark]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, r12
    mov edx, 3
    mov r8d, 2
    mov r9d, 13
    mov dword [rsp+20h], 14
    call [Rectangle]
    mov rcx, r12
    mov edx, 9
    mov r8d, 2
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 13
    mov r8d, 6
    call [LineTo]
    jmp .ico_done

.ico_open:
    mov rcx, r12
    mov rdx, [hPenOrange]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrOrange]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 5
    mov r9d, 14
    mov dword [rsp+20h], 13
    call [Rectangle]
    mov rcx, r12
    mov edx, 2
    mov r8d, 3
    mov r9d, 7
    mov dword [rsp+20h], 6
    call [Rectangle]
    jmp .ico_done

.ico_open_demo:
    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrBlue]
    call [SelectObject]
    mov rcx, r12
    mov edx, 3
    mov r8d, 8
    mov r9d, 6
    mov dword [rsp+20h], 14
    call [Rectangle]

    mov rcx, r12
    mov rdx, [hPenGreen]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrGreen]
    call [SelectObject]
    mov rcx, r12
    mov edx, 7
    mov r8d, 5
    mov r9d, 10
    mov dword [rsp+20h], 14
    call [Rectangle]

    mov rcx, r12
    mov rdx, [hPenOrange]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrOrange]
    call [SelectObject]
    mov rcx, r12
    mov edx, 11
    mov r8d, 2
    mov r9d, 14
    mov dword [rsp+20h], 14
    call [Rectangle]
    jmp .ico_done

.ico_save:
    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrBlue]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 2
    mov r9d, 14
    mov dword [rsp+20h], 14
    call [Rectangle]

    mov rcx, r12
    mov rdx, [hPenWhite]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, r12
    mov edx, 4
    mov r8d, 2
    mov r9d, 12
    mov dword [rsp+20h], 7
    call [Rectangle]
    jmp .ico_done

.ico_saveas:
    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrBlue]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 2
    mov r9d, 14
    mov dword [rsp+20h], 14
    call [Rectangle]

    mov rcx, r12
    mov rdx, [hPenWhite]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, r12
    mov edx, 4
    mov r8d, 2
    mov r9d, 12
    mov dword [rsp+20h], 7
    call [Rectangle]

    mov rcx, r12
    mov rdx, [hPenOrange]
    call [SelectObject]
    mov rcx, r12
    mov edx, 10
    mov r8d, 10
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 14
    call [LineTo]
    jmp .ico_done

.ico_exit:
    mov rcx, r12
    mov rdx, [hPenRed]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 2
    mov r9d, 9
    mov dword [rsp+20h], 14
    call [Rectangle]

    mov rcx, r12
    mov rdx, [hPenRed]
    call [SelectObject]
    mov rcx, r12
    mov edx, 6
    mov r8d, 8
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 8
    call [LineTo]
    mov rcx, r12
    mov edx, 11
    mov r8d, 5
    call [LineTo]
    mov rcx, r12
    mov edx, 14
    mov r8d, 8
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 11
    mov r8d, 11
    call [LineTo]
    jmp .ico_done

.ico_undo:
    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov edx, 13
    mov r8d, 12
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 13
    mov r8d, 8
    call [LineTo]
    mov rcx, r12
    mov edx, 6
    mov r8d, 8
    call [LineTo]
    mov rcx, r12
    mov edx, 9
    mov r8d, 5
    call [LineTo]
    mov rcx, r12
    mov edx, 6
    mov r8d, 8
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 9
    mov r8d, 11
    call [LineTo]
    jmp .ico_done

.ico_redo:
    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov edx, 3
    mov r8d, 12
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 3
    mov r8d, 8
    call [LineTo]
    mov rcx, r12
    mov edx, 10
    mov r8d, 8
    call [LineTo]
    mov rcx, r12
    mov edx, 7
    mov r8d, 5
    call [LineTo]
    mov rcx, r12
    mov edx, 10
    mov r8d, 8
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 7
    mov r8d, 11
    call [LineTo]
    jmp .ico_done

.ico_cut:
    mov rcx, r12
    mov rdx, [hPenRed]
    call [SelectObject]
    mov rcx, r12
    mov edx, 3
    mov r8d, 3
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 13
    mov r8d, 13
    call [LineTo]
    mov rcx, r12
    mov edx, 13
    mov r8d, 3
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 3
    mov r8d, 13
    call [LineTo]
    jmp .ico_done

.ico_copy:
    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 2
    mov r9d, 10
    mov dword [rsp+20h], 11
    call [Rectangle]
    mov rcx, r12
    mov edx, 6
    mov r8d, 5
    mov r9d, 14
    mov dword [rsp+20h], 14
    call [Rectangle]
    jmp .ico_done

.ico_paste:
    mov rcx, r12
    mov rdx, [hPenOrange]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrOrange]
    call [SelectObject]
    mov rcx, r12
    mov edx, 4
    mov r8d, 3
    mov r9d, 13
    mov dword [rsp+20h], 14
    call [Rectangle]
    mov rcx, r12
    mov rdx, [hPenDark]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, r12
    mov edx, 6
    mov r8d, 1
    mov r9d, 11
    mov dword [rsp+20h], 4
    call [Rectangle]
    jmp .ico_done

.ico_del_cell:
    mov rcx, r12
    mov rdx, [hPenDark]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 2
    mov r9d, 14
    mov dword [rsp+20h], 14
    call [Rectangle]
    mov rcx, r12
    mov rdx, [hPenRed]
    call [SelectObject]
    mov rcx, r12
    mov edx, 4
    mov r8d, 8
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 12
    mov r8d, 8
    call [LineTo]
    jmp .ico_done

.ico_select_all:
    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 2
    mov r9d, 14
    mov dword [rsp+20h], 14
    call [Rectangle]
    mov rcx, r12
    mov edx, 8
    mov r8d, 2
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 8
    mov r8d, 14
    call [LineTo]
    mov rcx, r12
    mov edx, 2
    mov r8d, 8
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 8
    call [LineTo]
    jmp .ico_done

.ico_find:
    mov rcx, r12
    mov rdx, [hPenDark]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, r12
    mov edx, 3
    mov r8d, 3
    mov r9d, 11
    mov dword [rsp+20h], 11
    call [Ellipse]
    mov rcx, r12
    mov edx, 9
    mov r8d, 9
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 14
    call [LineTo]
    jmp .ico_done

.ico_gridlines:
    mov rcx, r12
    mov rdx, [hPenDark]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 2
    mov r9d, 14
    mov dword [rsp+20h], 14
    call [Rectangle]
    mov rcx, r12
    mov edx, 6
    mov r8d, 2
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 6
    mov r8d, 14
    call [LineTo]
    mov rcx, r12
    mov edx, 10
    mov r8d, 2
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 10
    mov r8d, 14
    call [LineTo]
    mov rcx, r12
    mov edx, 2
    mov r8d, 6
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 6
    call [LineTo]
    mov rcx, r12
    mov edx, 2
    mov r8d, 10
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 10
    call [LineTo]
    jmp .ico_done

.ico_zoom100:
    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 2
    mov r9d, 12
    mov dword [rsp+20h], 12
    call [Ellipse]
    mov rcx, r12
    mov edx, 10
    mov r8d, 10
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 14
    call [LineTo]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]
    mov rcx, r12
    mov edx, 009A572Bh
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyph1]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .ico_done

.ico_refresh:
    mov rcx, r12
    mov rdx, [hPenGreen]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, r12
    mov edx, 3
    mov r8d, 3
    mov r9d, 13
    mov dword [rsp+20h], 13
    call [Ellipse]
    mov rcx, r12
    mov edx, 13
    mov r8d, 8
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 15
    mov r8d, 6
    call [LineTo]
    jmp .ico_done

.ico_ins_row:
    mov rcx, r12
    mov rdx, [hPenDark]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 8
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 8
    call [LineTo]
    mov rcx, r12
    mov rdx, [hPenGreen]
    call [SelectObject]
    mov rcx, r12
    mov edx, 8
    mov r8d, 4
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 8
    mov r8d, 12
    call [LineTo]
    mov rcx, r12
    mov edx, 4
    mov r8d, 8
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 12
    mov r8d, 8
    call [LineTo]
    jmp .ico_done

.ico_del_row:
    mov rcx, r12
    mov rdx, [hPenDark]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 8
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 8
    call [LineTo]
    mov rcx, r12
    mov rdx, [hPenRed]
    call [SelectObject]
    mov rcx, r12
    mov edx, 4
    mov r8d, 8
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 12
    mov r8d, 8
    call [LineTo]
    jmp .ico_done

.ico_fx:
    mov rcx, r12
    mov rdx, [hFontBoldItalic]
    call [SelectObject]
    mov rcx, r12
    mov edx, 009A572Bh
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphFX]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .ico_done

.ico_recalc:
    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrBlue]
    call [SelectObject]
    mov rcx, r12
    mov edx, 3
    mov r8d, 2
    mov r9d, 13
    mov dword [rsp+20h], 14
    call [Rectangle]
    mov rcx, r12
    mov rdx, [hPenWhite]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, r12
    mov edx, 4
    mov r8d, 3
    mov r9d, 12
    mov dword [rsp+20h], 6
    call [Rectangle]
    jmp .ico_done

.ico_autosum:
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]
    mov rcx, r12
    mov edx, 00467321h       ; Green
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphSigma]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .ico_done

.ico_average:
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]
    mov rcx, r12
    mov edx, 009A572Bh
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphAvg]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .ico_done

.ico_min:
    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov edx, 8
    mov r8d, 3
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 8
    mov r8d, 12
    call [LineTo]
    mov rcx, r12
    mov edx, 5
    mov r8d, 9
    call [LineTo]
    mov rcx, r12
    mov edx, 8
    mov r8d, 12
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 11
    mov r8d, 9
    call [LineTo]
    mov rcx, r12
    mov edx, 3
    mov r8d, 13
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 13
    mov r8d, 13
    call [LineTo]
    jmp .ico_done

.ico_max:
    mov rcx, r12
    mov rdx, [hPenGreen]
    call [SelectObject]
    mov rcx, r12
    mov edx, 8
    mov r8d, 13
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 8
    mov r8d, 4
    call [LineTo]
    mov rcx, r12
    mov edx, 5
    mov r8d, 7
    call [LineTo]
    mov rcx, r12
    mov edx, 8
    mov r8d, 4
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 11
    mov r8d, 7
    call [LineTo]
    mov rcx, r12
    mov edx, 3
    mov r8d, 3
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 13
    mov r8d, 3
    call [LineTo]
    jmp .ico_done

.ico_pmt:
    mov rcx, r12
    mov rdx, [hPenOrange]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrOrange]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 2
    mov r9d, 14
    mov dword [rsp+20h], 14
    call [Ellipse]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]
    mov rcx, r12
    xor edx, edx
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphUSD]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .ico_done

.ico_sort_asc:
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]
    mov rcx, r12
    mov edx, 009A572Bh
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphA]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0800h
    call [DrawTextA]
    mov rcx, r12
    mov rdx, [hPenGreen]
    call [SelectObject]
    mov rcx, r12
    mov edx, 12
    mov r8d, 3
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 12
    mov r8d, 13
    call [LineTo]
    mov rcx, r12
    mov edx, 10
    mov r8d, 10
    call [LineTo]
    mov rcx, r12
    mov edx, 12
    mov r8d, 13
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 10
    call [LineTo]
    jmp .ico_done

.ico_sort_desc:
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]
    mov rcx, r12
    mov edx, 009A572Bh
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphZ]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0800h
    call [DrawTextA]
    mov rcx, r12
    mov rdx, [hPenOrange]
    call [SelectObject]
    mov rcx, r12
    mov edx, 12
    mov r8d, 3
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 12
    mov r8d, 13
    call [LineTo]
    mov rcx, r12
    mov edx, 10
    mov r8d, 10
    call [LineTo]
    mov rcx, r12
    mov edx, 12
    mov r8d, 13
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 14
    mov r8d, 10
    call [LineTo]
    jmp .ico_done

.ico_benchmark:
    mov rcx, r12
    mov rdx, [hPenOrange]
    call [SelectObject]
    mov rcx, r12
    mov edx, 9
    mov r8d, 2
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 4
    mov r8d, 8
    call [LineTo]
    mov rcx, r12
    mov edx, 8
    mov r8d, 8
    call [LineTo]
    mov rcx, r12
    mov edx, 7
    mov r8d, 14
    call [LineTo]
    jmp .ico_done

.ico_form_ref:
    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrBlue]
    call [SelectObject]
    mov rcx, r12
    mov edx, 3
    mov r8d, 2
    mov r9d, 13
    mov dword [rsp+20h], 14
    call [Rectangle]
    mov rcx, r12
    mov rdx, [hPenWhite]
    call [SelectObject]
    mov rcx, r12
    mov edx, 5
    mov r8d, 5
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 11
    mov r8d, 5
    call [LineTo]
    mov rcx, r12
    mov edx, 5
    mov r8d, 8
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 11
    mov r8d, 8
    call [LineTo]
    mov rcx, r12
    mov edx, 5
    mov r8d, 11
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 11
    mov r8d, 11
    call [LineTo]
    jmp .ico_done

.ico_simd_info:
    mov rcx, r12
    mov rdx, [hPenGreen]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrGreen]
    call [SelectObject]
    mov rcx, r12
    mov edx, 4
    mov r8d, 4
    mov r9d, 12
    mov dword [rsp+20h], 12
    call [Rectangle]
    mov rcx, r12
    mov rdx, [hPenDark]
    call [SelectObject]
    mov rcx, r12
    mov edx, 6
    mov r8d, 2
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 6
    mov r8d, 4
    call [LineTo]
    mov rcx, r12
    mov edx, 10
    mov r8d, 2
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 10
    mov r8d, 4
    call [LineTo]
    mov rcx, r12
    mov edx, 6
    mov r8d, 12
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 6
    mov r8d, 14
    call [LineTo]
    mov rcx, r12
    mov edx, 10
    mov r8d, 12
    xor r9, r9
    call [MoveToEx]
    mov rcx, r12
    mov edx, 10
    mov r8d, 14
    call [LineTo]
    jmp .ico_done

.ico_about:
    mov rcx, r12
    mov rdx, [hPenBlue]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hBrBlue]
    call [SelectObject]
    mov rcx, r12
    mov edx, 2
    mov r8d, 2
    mov r9d, 14
    mov dword [rsp+20h], 14
    call [Ellipse]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]
    mov rcx, r12
    mov edx, 00FFFFFFh
    call [SetTextColor]
    mov rcx, r12
    lea rdx, [szGlyphi]
    mov r8d, -1
    lea r9, [rcMenu16]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .ico_done

.ico_done:
    mov rbx, [rbp-8]
    mov r12, [rbp-10h]
    mov r13, [rbp-18h]
    leave
    ret

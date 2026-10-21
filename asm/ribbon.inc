; ============================================================================
; Excel Viewer Pro — Modern Ribbon Toolbar & Controls
; ============================================================================



; =============================================================================
; Build Ribbon Bar matching main.py
; =============================================================================
BuildRibbonBar:
    push rbp
    mov rbp, rsp
    sub rsp, 60h

    ; Ribbon Tabs Notebook (Home, Insert, Data, Formulas, View)
    xor ecx, ecx
    lea rdx, [tab_class]
    xor r8, r8
    mov r9d, WS_CHILD or WS_VISIBLE
    mov dword [rsp+20h], 4
    mov dword [rsp+28h], 2
    mov dword [rsp+30h], 1256
    mov dword [rsp+38h], 118
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_RIBBON_TABS
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov [hRibbonTabs], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    ; Add Ribbon Tabs
    lea rdi, [tcitem]
    xor eax, eax
    mov ecx, 48
    rep stosb
    mov dword [tcitem], TCIF_TEXT

    lea rax, [rtab_home]
    mov qword [tcitem+16], rax
    mov rcx, [hRibbonTabs]
    mov edx, TCM_INSERTITEMA
    mov r8d, 0
    lea r9, [tcitem]
    call [SendMessageA]

    lea rax, [rtab_insert]
    mov qword [tcitem+16], rax
    mov rcx, [hRibbonTabs]
    mov edx, TCM_INSERTITEMA
    mov r8d, 1
    lea r9, [tcitem]
    call [SendMessageA]

    lea rax, [rtab_data]
    mov qword [tcitem+16], rax
    mov rcx, [hRibbonTabs]
    mov edx, TCM_INSERTITEMA
    mov r8d, 2
    lea r9, [tcitem]
    call [SendMessageA]

    lea rax, [rtab_formulas]
    mov qword [tcitem+16], rax
    mov rcx, [hRibbonTabs]
    mov edx, TCM_INSERTITEMA
    mov r8d, 3
    lea r9, [tcitem]
    call [SendMessageA]

    lea rax, [rtab_view]
    mov qword [tcitem+16], rax
    mov rcx, [hRibbonTabs]
    mov edx, TCM_INSERTITEMA
    mov r8d, 4
    lea r9, [tcitem]
    call [SendMessageA]

    ; -------------------------------------------------------------------------
    ; Group 1: Undo (x=10, y=26, w=72, h=88)
    ; -------------------------------------------------------------------------
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_undo_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_GROUPBOX
    mov dword [rsp+20h], 10
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 72
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_UNDO
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Save Button (Green Outline Accent Button)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_save_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 60
    mov dword [rsp+38h], 20
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_SAVE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Undo Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_undo_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 64
    mov dword [rsp+30h], 60
    mov dword [rsp+38h], 20
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_UNDO
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Redo Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_redo_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 86
    mov dword [rsp+30h], 60
    mov dword [rsp+38h], 20
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_REDO
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; -------------------------------------------------------------------------
    ; Group 2: Clipboard (x=88, y=26, w=122, h=88)
    ; -------------------------------------------------------------------------
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_clip_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_GROUPBOX
    mov dword [rsp+20h], 88
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 122
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_CLIPBOARD
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Paste Big Button (Solid Accent Blue)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_paste_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 94
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 48
    mov dword [rsp+38h], 64
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_PASTE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; Cut Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_cut_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 146
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 58
    mov dword [rsp+38h], 20
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_CUT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Copy Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_copy_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 146
    mov dword [rsp+28h], 64
    mov dword [rsp+30h], 58
    mov dword [rsp+38h], 20
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_COPY
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Format Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_format_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 146
    mov dword [rsp+28h], 86
    mov dword [rsp+30h], 58
    mov dword [rsp+38h], 20
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FORMAT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; -------------------------------------------------------------------------
    ; Group 3: Font (x=212, y=26, w=212, h=88)
    ; -------------------------------------------------------------------------
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_font_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_GROUPBOX
    mov dword [rsp+20h], 212
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 212
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_FONT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Font Family Combo
    xor ecx, ecx
    lea rdx, [combo_class]
    xor r8, r8
    mov r9d, WS_CHILD or WS_VISIBLE or CBS_DROPDOWNLIST or WS_VSCROLL
    mov dword [rsp+20h], 218
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 96
    mov dword [rsp+38h], 160
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_COMBO_FONT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov [hComboFont], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]
    mov rcx, [hComboFont]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_font1]
    call [SendMessageA]
    mov rcx, [hComboFont]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_font2]
    call [SendMessageA]
    mov rcx, [hComboFont]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_font3]
    call [SendMessageA]
    mov rcx, [hComboFont]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_font4]
    call [SendMessageA]
    mov rcx, [hComboFont]
    mov edx, CB_SETCURSEL
    mov r8d, 0
    xor r9d, r9d
    call [SendMessageA]

    ; Font Size Combo
    xor ecx, ecx
    lea rdx, [combo_class]
    xor r8, r8
    mov r9d, WS_CHILD or WS_VISIBLE or CBS_DROPDOWNLIST or WS_VSCROLL
    mov dword [rsp+20h], 318
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 42
    mov dword [rsp+38h], 160
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_COMBO_SIZE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov [hComboSize], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]
    mov rcx, [hComboSize]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_sz1]
    call [SendMessageA]
    mov rcx, [hComboSize]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_sz2]
    call [SendMessageA]
    mov rcx, [hComboSize]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_sz3]
    call [SendMessageA]
    mov rcx, [hComboSize]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_sz4]
    call [SendMessageA]
    mov rcx, [hComboSize]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_sz5]
    call [SendMessageA]
    mov rcx, [hComboSize]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_sz6]
    call [SendMessageA]
    mov rcx, [hComboSize]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_sz7]
    call [SendMessageA]
    mov rcx, [hComboSize]
    mov edx, CB_SETCURSEL
    mov r8d, 2 ; 10
    xor r9d, r9d
    call [SendMessageA]

    ; A+ (Increase font)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_inc_f_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 364
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 26
    mov dword [rsp+38h], 22
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INC_FONT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; A- (Decrease font)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_dec_f_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 392
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 26
    mov dword [rsp+38h], 22
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DEC_FONT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; B, I, U, S, Border, Fill, Text Color row (y=70, h=24)
    ; Bold Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_bold_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 218
    mov dword [rsp+28h], 70
    mov dword [rsp+30h], 22
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_BOLD
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; Italic Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_italic_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 242
    mov dword [rsp+28h], 70
    mov dword [rsp+30h], 22
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_ITALIC
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; Underline Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_under_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 266
    mov dword [rsp+28h], 70
    mov dword [rsp+30h], 22
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_UNDERLINE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; Strike Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_strike_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 290
    mov dword [rsp+28h], 70
    mov dword [rsp+30h], 22
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_STRIKE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; Border Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_border_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 314
    mov dword [rsp+28h], 70
    mov dword [rsp+30h], 26
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_BORDER
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Fill Color Button (Split Button with Color Strip & Chevron)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_fill_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 344
    mov dword [rsp+28h], 70
    mov dword [rsp+30h], 36
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FILL_COLOR
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Text Color Button (Split Button with Color Strip & Chevron)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_col_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 384
    mov dword [rsp+28h], 70
    mov dword [rsp+30h], 36
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_TEXT_COLOR
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; -------------------------------------------------------------------------
    ; Group 4: Alignment (x=430, y=26, w=174, h=88)
    ; -------------------------------------------------------------------------
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_align_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_GROUPBOX
    mov dword [rsp+20h], 430
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 174
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_ALIGN
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Top/Mid/Bot row (y=42)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_top_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 436
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 24
    mov dword [rsp+38h], 22
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_TOP
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_mid_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 462
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 24
    mov dword [rsp+38h], 22
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_MID
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_bot_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 488
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 24
    mov dword [rsp+38h], 22
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_BOT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Wrap Text Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_wrap_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 516
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 82
    mov dword [rsp+38h], 22
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_WRAP
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Left/Center/Right row (y=68)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_left_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 436
    mov dword [rsp+28h], 68
    mov dword [rsp+30h], 24
    mov dword [rsp+38h], 22
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_LEFT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_center_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 462
    mov dword [rsp+28h], 68
    mov dword [rsp+30h], 24
    mov dword [rsp+38h], 22
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_CENTER
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_right_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 488
    mov dword [rsp+28h], 68
    mov dword [rsp+30h], 24
    mov dword [rsp+38h], 22
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_RIGHT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Merge & Ctr Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_merge_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 516
    mov dword [rsp+28h], 68
    mov dword [rsp+30h], 82
    mov dword [rsp+38h], 22
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_MERGE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; -------------------------------------------------------------------------
    ; Group 5: Number (x=610, y=26, w=170, h=88)
    ; -------------------------------------------------------------------------
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_num_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_GROUPBOX
    mov dword [rsp+20h], 610
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 170
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_NUMBER
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Number Format Combo
    xor ecx, ecx
    lea rdx, [combo_class]
    xor r8, r8
    mov r9d, WS_CHILD or WS_VISIBLE or CBS_DROPDOWNLIST or WS_VSCROLL
    mov dword [rsp+20h], 616
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 158
    mov dword [rsp+38h], 160
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_COMBO_NUM_FMT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov [hComboNumFmt], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]
    mov rcx, [hComboNumFmt]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_fmt1]
    call [SendMessageA]
    mov rcx, [hComboNumFmt]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_fmt2]
    call [SendMessageA]
    mov rcx, [hComboNumFmt]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_fmt3]
    call [SendMessageA]
    mov rcx, [hComboNumFmt]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    lea r9, [cb_fmt4]
    call [SendMessageA]
    mov rcx, [hComboNumFmt]
    mov edx, CB_SETCURSEL
    mov r8d, 0 ; General
    xor r9d, r9d
    call [SendMessageA]

    ; Currency & Decimal buttons (₽, $, %, ,, .0+, .0-) y=68, h=24
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_rub_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 616
    mov dword [rsp+28h], 68
    mov dword [rsp+30h], 24
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_RUBLE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_usd_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 642
    mov dword [rsp+28h], 68
    mov dword [rsp+30h], 24
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DOLLAR
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_pct_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 668
    mov dword [rsp+28h], 68
    mov dword [rsp+30h], 24
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_PERCENT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_com_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 694
    mov dword [rsp+28h], 68
    mov dword [rsp+30h], 24
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_COMMA
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_idec_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 720
    mov dword [rsp+28h], 68
    mov dword [rsp+30h], 26
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INC_DECIMAL
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_ddec_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 748
    mov dword [rsp+28h], 68
    mov dword [rsp+30h], 26
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DEC_DECIMAL
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; -------------------------------------------------------------------------
    ; Group 6: Styles (x=786, y=26, w=142, h=88)
    ; -------------------------------------------------------------------------
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_styles_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_GROUPBOX
    mov dword [rsp+20h], 786
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 142
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_STYLES
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Conditional Formatting Button (Teal Outline Accent Button)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_cf_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 792
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 130
    mov dword [rsp+38h], 34
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_COND_FMT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Format Cells Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_fmt_cells_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 792
    mov dword [rsp+28h], 80
    mov dword [rsp+30h], 130
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FMT_CELLS
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; -------------------------------------------------------------------------
    ; Group 7: Cells (x=934, y=26, w=88, h=88)
    ; -------------------------------------------------------------------------
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_cells_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_GROUPBOX
    mov dword [rsp+20h], 934
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 88
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_CELLS
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Insert Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_ins_c_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 940
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 76
    mov dword [rsp+38h], 20
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INSERT_CELL
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Delete Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_del_c_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 940
    mov dword [rsp+28h], 64
    mov dword [rsp+30h], 76
    mov dword [rsp+38h], 20
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DELETE_CELL
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Auto-fit Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_autofit_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 940
    mov dword [rsp+28h], 86
    mov dword [rsp+30h], 76
    mov dword [rsp+38h], 20
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_AUTOFIT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; -------------------------------------------------------------------------
    ; Group 8: Editing (x=1028, y=26, w=182, h=88)
    ; -------------------------------------------------------------------------
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_edit_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_GROUPBOX
    mov dword [rsp+20h], 1028
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 182
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_EDITING
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; AutoSum Button (Blue Outline Accent Button) - Col 1 Top
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_autosum_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 1034
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 96
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_AUTOSUM
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Sort & Filter Button - Col 1 Bottom
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_sort_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 1034
    mov dword [rsp+28h], 74
    mov dword [rsp+30h], 96
    mov dword [rsp+38h], 26
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_SORT_FILTER
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Clear Button - Col 2 Top
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_clear_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 1136
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 68
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_CLEAR
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Find & Select Button - Col 2 Bottom
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_find_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 1136
    mov dword [rsp+28h], 74
    mov dword [rsp+30h], 68
    mov dword [rsp+38h], 26
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FIND_SELECT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    call BuildInsertTabControls
    call BuildDataTabControls
    call BuildFormulasTabControls
    call BuildViewTabControls

    leave
    ret


; =============================================================================
; Ribbon GDI Resources Initialization & Cleanup
; =============================================================================
InitRibbonGdiResources:
    push rbp
    mov rbp, rsp
    sub rsp, 30h

    ; hBrBtnFace (#F7F9FA)
    mov ecx, 00FAF9F7h
    call [CreateSolidBrush]
    mov [hBrBtnFace], rax

    ; hBrBtnHover (#E5F1FB)
    mov ecx, 00FBF1E5h
    call [CreateSolidBrush]
    mov [hBrBtnHover], rax

    ; hBrBtnPress (#CCE4F7)
    mov ecx, 00F7E4CCh
    call [CreateSolidBrush]
    mov [hBrBtnPress], rax

    ; hBrBlue (#2B579A)
    mov ecx, 009A572Bh
    call [CreateSolidBrush]
    mov [hBrBlue], rax

    ; hBrGreen (#217346)
    mov ecx, 00467321h
    call [CreateSolidBrush]
    mov [hBrGreen], rax

    ; hBrOrange (#ED7D31)
    mov ecx, 00317DEDh
    call [CreateSolidBrush]
    mov [hBrOrange], rax

    ; hBrYellow (#FFC000)
    mov ecx, 0000C0FFh
    call [CreateSolidBrush]
    mov [hBrYellow], rax

    ; hBrRed (#C00000)
    mov ecx, 000000C0h
    call [CreateSolidBrush]
    mov [hBrRed], rax

    ; hBrCyan (#33A1DE)
    mov ecx, 00DEA133h
    call [CreateSolidBrush]
    mov [hBrCyan], rax

    ; hBrWhite (#FFFFFF)
    mov ecx, 00FFFFFFh
    call [CreateSolidBrush]
    mov [hBrWhite], rax

    ; hBrGray (#A6A6A6)
    mov ecx, 00A6A6A6h
    call [CreateSolidBrush]
    mov [hBrGray], rax

    ; hBrSlate (#5A6473)
    mov ecx, 0073645Ah
    call [CreateSolidBrush]
    mov [hBrSlate], rax

    ; hPenBorder (#D1D5DB)
    xor ecx, ecx
    mov edx, 1
    mov r8d, 00DBD5D1h
    call [CreatePen]
    mov [hPenBorder], rax

    ; hPenBorderPress (#005499)
    xor ecx, ecx
    mov edx, 1
    mov r8d, 00995400h
    call [CreatePen]
    mov [hPenBorderPress], rax

    ; hPenBlue (#2B579A)
    xor ecx, ecx
    mov edx, 1
    mov r8d, 009A572Bh
    call [CreatePen]
    mov [hPenBlue], rax

    ; hPenGreen (#217346)
    xor ecx, ecx
    mov edx, 1
    mov r8d, 00467321h
    call [CreatePen]
    mov [hPenGreen], rax

    ; hPenOrange (#ED7D31)
    xor ecx, ecx
    mov edx, 1
    mov r8d, 00317DEDh
    call [CreatePen]
    mov [hPenOrange], rax

    ; hPenRed (#C00000)
    xor ecx, ecx
    mov edx, 1
    mov r8d, 000000C0h
    call [CreatePen]
    mov [hPenRed], rax

    ; hPenDark (#333333)
    xor ecx, ecx
    mov edx, 1
    mov r8d, 00333333h
    call [CreatePen]
    mov [hPenDark], rax

    ; hPenDark2 (#333333, width=2)
    xor ecx, ecx
    mov edx, 2
    mov r8d, 00333333h
    call [CreatePen]
    mov [hPenDark2], rax

    ; hPenWhite (#FFFFFF)
    xor ecx, ecx
    mov edx, 1
    mov r8d, 00FFFFFFh
    call [CreatePen]
    mov [hPenWhite], rax

    ; hPenNull (PS_NULL = 5)
    mov ecx, 5
    xor edx, edx
    xor r8d, r8d
    call [CreatePen]
    mov [hPenNull], rax

    ; hBrAccentBlue (#2B66D2)
    mov ecx, 00D2662Bh
    call [CreateSolidBrush]
    mov [hBrAccentBlue], rax

    ; hBrAccentHover (#3B7CE8)
    mov ecx, 00E87C3Bh
    call [CreateSolidBrush]
    mov [hBrAccentHover], rax

    ; hPenAccentBlue (#2B66D2)
    xor ecx, ecx
    mov edx, 1
    mov r8d, 00D2662Bh
    call [CreatePen]
    mov [hPenAccentBlue], rax

    ; hPenAccentGreen (#217346)
    xor ecx, ecx
    mov edx, 1
    mov r8d, 00467321h
    call [CreatePen]
    mov [hPenAccentGreen], rax

    ; hPenAccentTeal (#0284C7)
    xor ecx, ecx
    mov edx, 1
    mov r8d, 00C78402h
    call [CreatePen]
    mov [hPenAccentTeal], rax

    ; hPenAccentRed (#D72D1E)
    xor ecx, ecx
    mov edx, 1
    mov r8d, 001E2DD7h
    call [CreatePen]
    mov [hPenAccentRed], rax

    leave
    ret

CleanupRibbonGdiResources:
    push rbp
    mov rbp, rsp
    sub rsp, 30h

    mov rcx, [hBrBtnFace]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hBrBtnHover]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hBrBtnPress]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hBrBlue]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hBrGreen]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hBrOrange]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hBrYellow]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hBrRed]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hBrCyan]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hBrWhite]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hBrGray]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hBrSlate]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hPenBorder]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hPenBorderPress]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hPenBlue]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hPenGreen]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hPenOrange]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hPenRed]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hPenDark]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hPenDark2]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hPenWhite]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hPenNull]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hBrAccentBlue]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hBrAccentHover]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hPenAccentBlue]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hPenAccentGreen]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hPenAccentTeal]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hPenAccentRed]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    ; Clean up dynamic cell font cache
    xor ebx, ebx
.cleanup_font_cache:
    cmp ebx, [font_cache_count]
    jge .done_font_cache
    mov rcx, [font_cache_fonts + rbx*8]
    test rcx, rcx
    jz .cfc_next
    call [DeleteObject]
.cfc_next:
    inc ebx
    jmp .cleanup_font_cache
.done_font_cache:
    leave
    ret

; =============================================================================
; DrawRibbonOwnerButton
; Draws modern ribbon buttons with custom colorful icons and clean labels
; Input: RSI = LPDRAWITEMSTRUCT
; =============================================================================
DrawRibbonOwnerButton:
    push rbp
    mov rbp, rsp
    sub rsp, 0A0h
    mov [rbp-08h], rbx
    mov [rbp-10h], r12
    mov [rbp-18h], r13
    mov [rbp-20h], r14
    mov [rbp-28h], r15
    mov [rbp-30h], rsi
    mov [rbp-38h], rdi

    mov r12, [rsi+32]       ; r12 = hDC
    mov r13d, [rsi+4]       ; r13d = CtlID
    mov r14d, [rsi+16]      ; r14d = itemState
    mov r15, [rsi+24]       ; r15 = hwndItem

    ; rcItem is at rsi+40
    mov eax, [rsi+40]
    mov [rbp-40h], eax      ; left
    mov [rcGlyphTemp], eax
    mov eax, [rsi+44]
    mov [rbp-44h], eax      ; top
    mov [rcGlyphTemp+4], eax
    mov eax, [rsi+48]
    mov [rbp-48h], eax      ; right
    mov [rcGlyphTemp+8], eax
    mov eax, [rsi+52]
    mov [rbp-4Ch], eax      ; bottom
    mov [rcGlyphTemp+12], eax

    ; width = right - left, height = bottom - top
    mov eax, [rbp-48h]
    sub eax, [rbp-40h]
    mov [rbp-50h], eax      ; width
    mov eax, [rbp-4Ch]
    sub eax, [rbp-44h]
    mov [rbp-54h], eax      ; height

    ; 1. Draw modern button background
    cmp r13d, ID_BTN_PASTE
    je .bg_paste_btn
    cmp r13d, ID_BTN_FX
    je .bg_fx_btn
    cmp r13d, ID_BTN_SAVE
    je .bg_save_btn
    cmp r13d, ID_BTN_COND_FMT
    je .bg_cond_fmt_btn
    cmp r13d, ID_BTN_AUTOSUM
    je .bg_autosum_btn

    test r14d, 1            ; ODS_SELECTED?
    jz .bg_not_pressed
    mov rdx, [hBrBtnPress]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBorderPress]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg

.bg_not_pressed:
    test r14d, 40h          ; ODS_HOTLIGHT?
    jz .bg_normal
    mov rdx, [hBrBtnHover]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg

.bg_normal:
    mov rdx, [hBrBtnFace]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBorder]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg

.bg_paste_btn:
    test r14d, 1
    jz @f
    mov rdx, [hBrBlue]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentBlue]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg
@@:
    test r14d, 40h
    jz @f
    mov rdx, [hBrAccentHover]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentBlue]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg
@@:
    mov rdx, [hBrAccentBlue]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentBlue]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg

.bg_fx_btn:
    test r14d, 1
    jz @f
    mov rdx, [hBrGreen]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg
@@:
    test r14d, 40h
    jz @f
    mov rdx, [hBrGreen]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBorderPress]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg
@@:
    mov rdx, [hBrGreen]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenGreen]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg

.bg_save_btn:
    test r14d, 1
    jz @f
    mov rdx, [hBrBtnPress]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentGreen]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg
@@:
    test r14d, 40h
    jz @f
    mov rdx, [hBrBtnHover]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentGreen]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg
@@:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentGreen]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg

.bg_cond_fmt_btn:
    test r14d, 1
    jz @f
    mov rdx, [hBrBtnPress]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentTeal]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg
@@:
    test r14d, 40h
    jz @f
    mov rdx, [hBrBtnHover]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentTeal]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg
@@:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentTeal]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg

.bg_autosum_btn:
    test r14d, 1
    jz @f
    mov rdx, [hBrBtnPress]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentBlue]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg
@@:
    test r14d, 40h
    jz @f
    mov rdx, [hBrBtnHover]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentBlue]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_bg
@@:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentBlue]
    mov rcx, r12
    call [SelectObject]

.draw_bg:
    mov rcx, r12
    mov edx, [rbp-40h]
    mov r8d, [rbp-44h]
    mov r9d, [rbp-48h]
    mov eax, [rbp-4Ch]
    mov [rsp+20h], eax
    mov dword [rsp+28h], 4
    mov dword [rsp+30h], 4
    call [RoundRect]

    ; 2. Determine center coordinates (ebx=cx, edx=cy)
    cmp dword [rbp-54h], 45
    jl .icon_pos_small

    ; Tall button: icon centered near top
    mov eax, [rbp-50h]
    shr eax, 1
    add eax, [rbp-40h]
    mov ebx, eax            ; cx
    mov ecx, [rbp-44h]
    add ecx, 18
    mov edx, ecx            ; cy
    jmp .draw_icon

.icon_pos_small:
    cmp dword [rbp-50h], 32
    jg .icon_left_pos
    ; Small square button (width <= 32): center icon horizontally
    mov eax, [rbp-50h]
    shr eax, 1
    add eax, [rbp-40h]
    mov ebx, eax            ; cx
    jmp .icon_vert_center
.icon_left_pos:
    ; Standard button: icon on left
    mov eax, [rbp-40h]
    add eax, 14
    mov ebx, eax            ; cx
.icon_vert_center:
    mov eax, [rbp-54h]
    shr eax, 1
    add eax, [rbp-44h]
    mov edx, eax            ; cy

.draw_icon:
    mov [rbp-58h], ebx
    mov [rbp-5Ch], edx

    cmp r13d, ID_BTN_SAVE
    je .ico_save
    cmp r13d, ID_BTN_UNDO
    je .ico_undo
    cmp r13d, ID_BTN_REDO
    je .ico_redo
    cmp r13d, ID_BTN_PASTE
    je .ico_paste
    cmp r13d, ID_BTN_CUT
    je .ico_cut
    cmp r13d, ID_BTN_COPY
    je .ico_copy
    cmp r13d, ID_BTN_FORMAT
    je .ico_format
    cmp r13d, ID_BTN_INC_FONT
    je .ico_font_inc
    cmp r13d, ID_BTN_DEC_FONT
    je .ico_font_dec
    cmp r13d, ID_BTN_BOLD
    je .ico_font_bold
    cmp r13d, ID_BTN_ITALIC
    je .ico_font_italic
    cmp r13d, ID_BTN_UNDERLINE
    je .ico_font_under
    cmp r13d, ID_BTN_STRIKE
    je .ico_font_strike
    cmp r13d, ID_BTN_BORDER
    je .ico_border
    cmp r13d, ID_BTN_FILL_COLOR
    je .ico_fill_color
    cmp r13d, ID_BTN_TEXT_COLOR
    je .ico_text_color
    cmp r13d, ID_BTN_TOP
    je .ico_align_top
    cmp r13d, ID_BTN_MID
    je .ico_align_mid
    cmp r13d, ID_BTN_BOT
    je .ico_align_bot
    cmp r13d, ID_BTN_WRAP
    je .ico_wrap_text
    cmp r13d, ID_BTN_LEFT
    je .ico_align_left
    cmp r13d, ID_BTN_CENTER
    je .ico_align_center
    cmp r13d, ID_BTN_RIGHT
    je .ico_align_right
    cmp r13d, ID_BTN_MERGE
    je .ico_merge_center
    cmp r13d, ID_BTN_RUBLE
    je .ico_ruble
    cmp r13d, ID_BTN_DOLLAR
    je .ico_dollar
    cmp r13d, ID_BTN_PERCENT
    je .ico_percent
    cmp r13d, ID_BTN_COMMA
    je .ico_comma
    cmp r13d, ID_BTN_INC_DECIMAL
    je .ico_inc_dec
    cmp r13d, ID_BTN_DEC_DECIMAL
    je .ico_dec_dec
    cmp r13d, ID_BTN_COND_FMT
    je .ico_cond_fmt
    cmp r13d, ID_BTN_FMT_CELLS
    je .ico_table
    cmp r13d, ID_BTN_INSERT_CELL
    je .ico_insert_cell
    cmp r13d, ID_BTN_DELETE_CELL
    je .ico_delete_cell
    cmp r13d, ID_BTN_AUTOFIT
    je .ico_autofit
    cmp r13d, ID_BTN_AUTOSUM
    je .ico_autosum
    cmp r13d, ID_BTN_CLEAR
    je .ico_clear_btn
    cmp r13d, ID_BTN_SORT_FILTER
    je .ico_sort_asc
    cmp r13d, ID_BTN_FIND_SELECT
    je .ico_lookup
    cmp r13d, ID_BTN_FX
    je .ico_formula_fx

    cmp r13d, ID_BTN_INS_TABLE
    je .ico_table
    cmp r13d, ID_BTN_INS_PIVOT
    je .ico_pivot
    cmp r13d, ID_BTN_INS_PICTURE
    je .ico_picture
    cmp r13d, ID_BTN_INS_SHAPES
    je .ico_shapes
    cmp r13d, ID_BTN_INS_ICONS
    je .ico_star
    cmp r13d, ID_BTN_INS_COL_CHART
    je .ico_col_chart
    cmp r13d, ID_BTN_INS_LINE_CHART
    je .ico_line_chart
    cmp r13d, ID_BTN_INS_PIE_CHART
    je .ico_pie_chart
    cmp r13d, ID_BTN_INS_BAR_CHART
    je .ico_bar_chart
    cmp r13d, ID_BTN_INS_LINK
    je .ico_link
    cmp r13d, ID_BTN_INS_TEXTBOX
    je .ico_textbox
    cmp r13d, ID_BTN_INS_SYMBOL
    je .ico_fx
    cmp r13d, ID_BTN_DATA_CSV
    je .ico_csv
    cmp r13d, ID_BTN_DATA_REFRESH
    je .ico_refresh
    cmp r13d, ID_BTN_DATA_SORT_ASC
    je .ico_sort_asc
    cmp r13d, ID_BTN_DATA_SORT_DESC
    je .ico_sort_desc
    cmp r13d, ID_BTN_DATA_FILTER
    je .ico_filter
    cmp r13d, ID_BTN_DATA_CLEAR
    je .ico_clear_filter
    cmp r13d, ID_BTN_DATA_SPLIT
    je .ico_split
    cmp r13d, ID_BTN_DATA_REMDUP
    je .ico_remdup
    cmp r13d, ID_BTN_DATA_VALID
    je .ico_valid
    cmp r13d, ID_BTN_DATA_BENCH
    je .ico_bench
    cmp r13d, ID_BTN_DATA_WHATIF
    je .ico_whatif
    cmp r13d, ID_BTN_FORM_FX
    je .ico_fx
    cmp r13d, ID_BTN_FORM_AUTOSUM
    je .ico_autosum
    cmp r13d, ID_BTN_FORM_FIN
    je .ico_finance
    cmp r13d, ID_BTN_FORM_LOGIC
    je .ico_logic
    cmp r13d, ID_BTN_FORM_TEXT
    je .ico_text_doc
    cmp r13d, ID_BTN_FORM_DATE
    je .ico_date
    cmp r13d, ID_BTN_FORM_LOOKUP
    je .ico_lookup
    cmp r13d, ID_BTN_FORM_MATH
    je .ico_math
    cmp r13d, ID_BTN_FORM_MORE
    je .ico_fx
    cmp r13d, ID_BTN_FORM_NAME_MGR
    je .ico_tag_hash
    cmp r13d, ID_BTN_FORM_DEF_NAME
    je .ico_tag_plus
    cmp r13d, ID_BTN_FORM_CALC_NOW
    je .ico_calc
    cmp r13d, ID_BTN_FORM_EVAL
    je .ico_calc
    cmp r13d, ID_BTN_VIEW_NORMAL
    je .ico_sheet_normal
    cmp r13d, ID_BTN_VIEW_PAGELAY
    je .ico_pagelayout
    cmp r13d, ID_BTN_VIEW_GRIDLINES
    je .ico_gridlines
    cmp r13d, ID_BTN_VIEW_FORMBAR
    je .ico_formbar
    cmp r13d, ID_BTN_VIEW_HEADINGS
    je .ico_headings
    cmp r13d, ID_BTN_VIEW_ZOOM100
    je .ico_zoom100
    cmp r13d, ID_BTN_VIEW_ZOOM_IN
    je .ico_zoom_in
    cmp r13d, ID_BTN_VIEW_ZOOM_OUT
    je .ico_zoom_out
    cmp r13d, ID_BTN_VIEW_FREEZE
    je .ico_freeze
    cmp r13d, ID_BTN_VIEW_NEW_WND
    je .ico_new_wnd
    jmp .ico_sheet_normal

.ico_save:
    mov rdx, [hBrGreen]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenGreen]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]

    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenWhite]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+2]
    lea eax, [eax-2]
    mov [rsp+20h], eax
    call [Rectangle]

    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    lea r9d, [ebx+4]
    lea eax, [eax+5]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_font_inc:
    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00202020h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]

    mov eax, ebx
    sub eax, 10
    mov [rcGlyphTemp], eax
    mov eax, [rbp-5Ch]
    sub eax, 8
    mov [rcGlyphTemp+4], eax
    mov eax, ebx
    mov [rcGlyphTemp+8], eax
    mov eax, [rbp-5Ch]
    add eax, 8
    mov [rcGlyphTemp+12], eax

    mov rcx, r12
    lea rdx, [szGlyphA]
    mov r8d, 1
    lea r9, [rcGlyphTemp]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]

    mov rdx, [hPenAccentGreen]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-4]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-4]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+8]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-1]
    call [LineTo]
    jmp .btn_drawn

.ico_font_dec:
    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00202020h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]

    mov eax, ebx
    sub eax, 10
    mov [rcGlyphTemp], eax
    mov eax, [rbp-5Ch]
    sub eax, 6
    mov [rcGlyphTemp+4], eax
    mov eax, ebx
    mov [rcGlyphTemp+8], eax
    mov eax, [rbp-5Ch]
    add eax, 8
    mov [rcGlyphTemp+12], eax

    mov rcx, r12
    lea rdx, [szGlyphA]
    mov r8d, 1
    lea r9, [rcGlyphTemp]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]

    mov rdx, [hPenAccentRed]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-4]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+8]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    call [LineTo]
    jmp .btn_drawn

.ico_font_bold:
    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00202020h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]

    mov rcx, r12
    lea rdx, [szGlyphB]
    mov r8d, 1
    lea r9, [rcGlyphTemp]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .btn_drawn

.ico_font_italic:
    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00202020h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBoldItalic]
    call [SelectObject]

    mov rcx, r12
    lea rdx, [szGlyphI]
    mov r8d, 1
    lea r9, [rcGlyphTemp]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .btn_drawn

.ico_font_under:
    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00202020h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]

    mov rcx, r12
    lea rdx, [szGlyphU]
    mov r8d, 1
    lea r9, [rcGlyphTemp]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]

    mov rdx, [hPenDark2]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx-5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    jmp .btn_drawn

.ico_font_strike:
    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00202020h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]

    mov rcx, r12
    lea rdx, [szGlyphS]
    mov r8d, 1
    lea r9, [rcGlyphTemp]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]

    mov rdx, [hPenAccentRed]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx-6]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+7]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    jmp .btn_drawn

.ico_border:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+3]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]

    mov rcx, r12
    lea edx, [ebx-7]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+3]
    mov r8d, [rbp-5Ch]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]

    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+10]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-1]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx+6]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+9]
    mov r8d, [rbp-5Ch]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx+7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+8]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    call [LineTo]
    jmp .btn_drawn

.ico_ruble:
    mov rdx, [hPenGreen]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-1]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-1]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx-5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx-5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+3]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+3]
    call [LineTo]
    jmp .btn_drawn

.ico_dollar:
    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00467321h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]

    mov rcx, r12
    lea rdx, [szGlyphUSD]
    mov r8d, 1
    lea r9, [rcGlyphTemp]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .btn_drawn

.ico_percent:
    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00D2662Bh
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]

    mov rcx, r12
    lea rdx, [szGlyphPCT]
    mov r8d, 1
    lea r9, [rcGlyphTemp]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .btn_drawn

.ico_comma:
    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00202020h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]

    mov eax, [rbp-40h]
    mov [rcGlyphTemp], eax
    mov eax, [rbp-44h]
    sub eax, 3
    mov [rcGlyphTemp+4], eax
    mov eax, [rbp-48h]
    mov [rcGlyphTemp+8], eax
    mov eax, [rbp-4Ch]
    sub eax, 3
    mov [rcGlyphTemp+12], eax

    mov rcx, r12
    lea rdx, [comma_str]
    mov r8d, 1
    lea r9, [rcGlyphTemp]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .btn_drawn

.ico_inc_dec:
    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00D2662Bh
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]

    mov rcx, r12
    lea rdx, [btn_idec_txt]
    mov r8d, -1
    lea r9, [rcGlyphTemp]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .btn_drawn

.ico_dec_dec:
    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00D2662Bh
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]

    mov rcx, r12
    lea rdx, [btn_ddec_txt]
    mov r8d, -1
    lea r9, [rcGlyphTemp]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .btn_drawn

.ico_insert_cell:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentGreen]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Ellipse]

    mov rcx, r12
    lea edx, [ebx-3]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+4]
    mov r8d, [rbp-5Ch]
    call [LineTo]

    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax+4]
    call [LineTo]
    jmp .render_btn_text

.ico_delete_cell:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentRed]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Ellipse]

    mov rcx, r12
    lea edx, [ebx-3]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+4]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    jmp .render_btn_text

.ico_autofit:
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]

    mov rdx, [hPenAccentBlue]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx-4]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+5]
    mov r8d, [rbp-5Ch]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-4]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+4]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    call [LineTo]
    jmp .render_btn_text

.ico_autosum:
    mov rdx, [hPenAccentBlue]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx+4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+1]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]
    jmp .render_btn_text

.ico_clear_btn:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenAccentRed]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx-5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-4]
    lea r9d, [ebx+5]
    lea eax, [eax+4]
    mov [rsp+20h], eax
    call [Rectangle]

    mov rdx, [hBrOrange]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenNull]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    mov r9d, ebx
    lea eax, [eax+3]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_formula_fx:
    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00FFFFFFh
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]

    mov rcx, r12
    lea rdx, [szGlyphFX]
    mov r8d, 2
    lea r9, [rcGlyphTemp]
    mov dword [rsp+20h], 0825h
    call [DrawTextA]
    jmp .btn_drawn

.ico_table:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+7]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrBlue]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+7]
    lea eax, [eax-2]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hPenBorder]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    jmp .render_btn_text

.ico_pivot:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+7]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrOrange]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenOrange]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+7]
    lea eax, [eax-2]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    lea r9d, [ebx-2]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_picture:
    mov rdx, [hBrCyan]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-8]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-7]
    lea r9d, [ebx+8]
    lea eax, [eax+7]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrYellow]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenNull]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    lea r9d, [ebx+6]
    lea eax, [eax-1]
    mov [rsp+20h], eax
    call [Ellipse]
    mov rdx, [hBrGreen]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-8]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    lea r9d, [ebx+8]
    lea eax, [eax+7]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_shapes:
    mov rdx, [hBrBlue]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    lea r9d, [ebx+7]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrOrange]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenOrange]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+2]
    lea eax, [eax+3]
    mov [rsp+20h], eax
    call [Ellipse]
    jmp .render_btn_text

.ico_star:
    mov rdx, [hBrYellow]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenOrange]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Ellipse]
    mov rdx, [hPenRed]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-4]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax+4]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+4]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    call [LineTo]
    jmp .render_btn_text

.ico_col_chart:
    mov rdx, [hPenNull]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hBrBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    lea r9d, [ebx-2]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrOrange]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+2]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrGreen]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx+3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_line_chart:
    mov rdx, [hPenGreen]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+4]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    call [LineTo]
    mov rdx, [hBrBlue]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-4]
    lea r9d, [ebx-1]
    lea eax, [eax-2]
    mov [rsp+20h], eax
    call [Ellipse]
    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+7]
    lea eax, [eax-4]
    mov [rsp+20h], eax
    call [Ellipse]
    jmp .render_btn_text

.ico_pie_chart:
    mov rdx, [hBrBlue]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Ellipse]
    mov rdx, [hBrOrange]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenOrange]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    mov eax, [rbp-5Ch]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_bar_chart:
    mov rdx, [hPenNull]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hBrBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    lea r9d, [ebx+3]
    lea eax, [eax-2]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrOrange]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-1]
    lea r9d, [ebx+6]
    lea eax, [eax+2]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrGreen]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+3]
    mov r9d, ebx
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_link:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-4]
    lea r9d, [ebx+1]
    lea eax, [eax+2]
    mov [rsp+20h], eax
    mov dword [rsp+28h], 2
    mov dword [rsp+30h], 2
    call [RoundRect]
    mov rcx, r12
    lea edx, [ebx-1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    lea r9d, [ebx+7]
    lea eax, [eax+4]
    mov [rsp+20h], eax
    mov dword [rsp+28h], 2
    mov dword [rsp+30h], 2
    call [RoundRect]
    jmp .render_btn_text

.ico_textbox:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+4]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+4]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    call [LineTo]
    jmp .render_btn_text

.ico_csv:
    mov rdx, [hBrGreen]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-7]
    lea r9d, [ebx+6]
    lea eax, [eax+7]
    mov [rsp+20h], eax
    mov dword [rsp+28h], 2
    mov dword [rsp+30h], 2
    call [RoundRect]
    mov rdx, [hPenWhite]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+4]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    call [LineTo]
    jmp .render_btn_text

.ico_refresh:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenGreen]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Ellipse]
    mov rdx, [hBrGreen]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-7]
    lea r9d, [ebx+6]
    lea eax, [eax-3]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_sort_asc:
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    jmp .render_btn_text

.ico_sort_desc:
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    call [LineTo]
    jmp .render_btn_text

.ico_filter:
    mov rdx, [hBrBlue]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    lea r9d, [ebx+6]
    lea eax, [eax-1]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-1]
    lea r9d, [ebx+2]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_clear_filter:
    mov rdx, [hBrBlue]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    lea r9d, [ebx+6]
    lea eax, [eax-1]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-1]
    lea r9d, [ebx+2]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hPenRed]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx+1]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+7]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    jmp .render_btn_text

.ico_split:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx-1]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rcx, r12
    lea edx, [ebx+1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+7]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_remdup:
    mov rdx, [hBrGreen]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenNull]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    lea r9d, [ebx+6]
    lea eax, [eax-1]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrGray]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    lea r9d, [ebx+6]
    lea eax, [eax+5]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hPenRed]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+3]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+3]
    call [LineTo]
    jmp .render_btn_text

.ico_valid:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hPenGreen]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+3]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    call [LineTo]
    jmp .render_btn_text

.ico_bench:
    mov rdx, [hPenOrange]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    call [LineTo]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    mov r8d, eax
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    jmp .render_btn_text

.ico_whatif:
    mov rdx, [hBrCyan]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Ellipse]
    jmp .render_btn_text

.ico_fx:
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-1]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+4]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+4]
    call [LineTo]
    jmp .render_btn_text

.ico_finance:
    mov rdx, [hBrYellow]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenOrange]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Ellipse]
    mov rdx, [hPenGreen]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-4]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax+4]
    call [LineTo]
    jmp .render_btn_text

.ico_logic:
    mov rdx, [hBrCyan]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    call [LineTo]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    call [LineTo]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    call [LineTo]
    jmp .render_btn_text

.ico_text_doc:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-7]
    lea r9d, [ebx+6]
    lea eax, [eax+7]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    call [LineTo]
    jmp .render_btn_text

.ico_date:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrRed]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenRed]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax-2]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_lookup:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+2]
    lea eax, [eax+2]
    mov [rsp+20h], eax
    call [Ellipse]
    mov rcx, r12
    lea edx, [ebx+1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    jmp .render_btn_text

.ico_math:
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-1]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+3]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+3]
    call [LineTo]
    jmp .render_btn_text

.ico_tag_hash:
    mov rdx, [hBrYellow]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenOrange]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    lea r9d, [ebx+6]
    lea eax, [eax+5]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_tag_plus:
    mov rdx, [hBrCyan]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    lea r9d, [ebx+6]
    lea eax, [eax+5]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_calc:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-7]
    lea r9d, [ebx+6]
    lea eax, [eax+7]
    mov [rsp+20h], eax
    mov dword [rsp+28h], 2
    mov dword [rsp+30h], 2
    call [RoundRect]
    mov rdx, [hBrCyan]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    lea r9d, [ebx+4]
    lea eax, [eax-2]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_sheet_normal:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+7]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrGreen]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenGreen]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+7]
    lea eax, [eax-2]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_pagelayout:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-7]
    lea r9d, [ebx+6]
    lea eax, [eax+7]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    lea r9d, [ebx+4]
    lea eax, [eax+5]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_gridlines:
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    call [LineTo]
    jmp .render_btn_text

.ico_formbar:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    lea r9d, [ebx+7]
    lea eax, [eax+3]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_headings:
    mov rdx, [hBrBlue]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    mov r9d, ebx
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrOrange]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenOrange]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    mov r8d, eax
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_zoom100:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+3]
    lea eax, [eax+3]
    mov [rsp+20h], eax
    call [Ellipse]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    jmp .render_btn_text

.ico_zoom_in:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+2]
    lea eax, [eax+2]
    mov [rsp+20h], eax
    call [Ellipse]
    mov rcx, r12
    lea edx, [ebx+1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-4]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    call [LineTo]
    jmp .render_btn_text

.ico_zoom_out:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+2]
    lea eax, [eax+2]
    mov [rsp+20h], eax
    call [Ellipse]
    mov rcx, r12
    lea edx, [ebx+1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    call [LineTo]
    jmp .render_btn_text

.ico_freeze:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    call [LineTo]
    jmp .render_btn_text

.ico_new_wnd:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+2]
    lea eax, [eax+2]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrCyan]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_paste:
    mov rdx, [hBrYellow]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenOrange]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-8]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-10]
    lea r9d, [ebx+8]
    lea eax, [eax+10]
    mov [rsp+20h], eax
    mov dword [rsp+28h], 2
    mov dword [rsp+30h], 2
    call [RoundRect]
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-7]
    lea r9d, [ebx+6]
    lea eax, [eax+8]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrGray]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-11]
    lea r9d, [ebx+4]
    lea eax, [eax-8]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+4]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    call [LineTo]
    jmp .render_btn_text

.ico_cut:
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]
    jmp .render_btn_text

.ico_copy:
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+2]
    lea eax, [eax+2]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    lea r9d, [ebx+6]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_format:
    mov rdx, [hBrYellow]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenOrange]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+5]
    mov eax, [rbp-5Ch]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrBlue]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    mov r8d, eax
    lea r9d, [ebx+2]
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_cond_fmt:
    mov rdx, [hPenNull]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hBrGreen]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    lea r9d, [ebx+6]
    lea eax, [eax-2]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrYellow]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-1]
    lea r9d, [ebx+3]
    lea eax, [eax+2]
    mov [rsp+20h], eax
    call [Rectangle]
    mov rdx, [hBrRed]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+3]
    mov r9d, ebx
    lea eax, [eax+6]
    mov [rsp+20h], eax
    call [Rectangle]
    jmp .render_btn_text

.ico_undo:
    mov rdx, [hPenAccentBlue]
    mov rcx, r12
    call [SelectObject]
    ; Arrowhead left: (cx-5, cy) to (cx-2, cy-3) and (cx-5, cy) to (cx-2, cy+3)
    mov rcx, r12
    lea edx, [ebx-5]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-5]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+3]
    call [LineTo]
    ; Curved body: (cx-4, cy) -> (cx+2, cy) -> (cx+4, cy+2) -> (cx+4, cy+4)
    mov rcx, r12
    lea edx, [ebx-4]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+2]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+4]
    call [LineTo]
    jmp .render_btn_text

.ico_redo:
    mov rdx, [hPenAccentBlue]
    mov rcx, r12
    call [SelectObject]
    ; Arrowhead right: (cx+5, cy) to (cx+2, cy-3) and (cx+5, cy) to (cx+2, cy+3)
    mov rcx, r12
    lea edx, [ebx+5]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+5]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+3]
    call [LineTo]
    ; Curved body: (cx+4, cy) -> (cx-2, cy) -> (cx-4, cy+2) -> (cx-4, cy+4)
    mov rcx, r12
    lea edx, [ebx+4]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-2]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+4]
    call [LineTo]
    jmp .render_btn_text

.ico_align_left:
    mov rdx, [hPenDark2]
    mov rcx, r12
    call [SelectObject]
    ; Line 1: (cx-6, cy-5) to (cx+3, cy-5)
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    call [LineTo]
    ; Line 2: (cx-6, cy) to (cx+6, cy)
    mov rcx, r12
    lea edx, [ebx-6]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    ; Line 3: (cx-6, cy+5) to (cx+1, cy+5)
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]
    jmp .btn_drawn

.ico_align_center:
    mov rdx, [hPenDark2]
    mov rcx, r12
    call [SelectObject]
    ; Line 1: (cx-4, cy-5) to (cx+4, cy-5)
    mov rcx, r12
    lea edx, [ebx-4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+4]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    call [LineTo]
    ; Line 2: (cx-6, cy) to (cx+6, cy)
    mov rcx, r12
    lea edx, [ebx-6]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    ; Line 3: (cx-3, cy+5) to (cx+3, cy+5)
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]
    jmp .btn_drawn

.ico_align_right:
    mov rdx, [hPenDark2]
    mov rcx, r12
    call [SelectObject]
    ; Line 1: (cx-3, cy-5) to (cx+6, cy-5)
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    call [LineTo]
    ; Line 2: (cx-6, cy) to (cx+6, cy)
    mov rcx, r12
    lea edx, [ebx-6]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    ; Line 3: (cx-1, cy+5) to (cx+6, cy+5)
    mov rcx, r12
    lea edx, [ebx-1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]
    jmp .btn_drawn

.ico_align_top:
    mov rdx, [hPenDark2]
    mov rcx, r12
    call [SelectObject]
    ; Top bar: (cx-6, cy-6) to (cx+6, cy-6)
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-6]
    call [LineTo]
    ; Line 2: (cx-4, cy) to (cx+4, cy)
    mov rcx, r12
    lea edx, [ebx-4]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+4]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    ; Line 3: (cx-2, cy+5) to (cx+2, cy+5)
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]
    jmp .btn_drawn

.ico_align_mid:
    mov rdx, [hPenDark2]
    mov rcx, r12
    call [SelectObject]
    ; Line 1: (cx-5, cy-5) to (cx+5, cy-5)
    mov rcx, r12
    lea edx, [ebx-5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    call [LineTo]
    ; Line 2: (cx-6, cy) to (cx+6, cy)
    mov rcx, r12
    lea edx, [ebx-6]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    ; Line 3: (cx-5, cy+5) to (cx+5, cy+5)
    mov rcx, r12
    lea edx, [ebx-5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]
    jmp .btn_drawn

.ico_align_bot:
    mov rdx, [hPenDark2]
    mov rcx, r12
    call [SelectObject]
    ; Line 1: (cx-2, cy-5) to (cx+2, cy-5)
    mov rcx, r12
    lea edx, [ebx-2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+2]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    call [LineTo]
    ; Line 2: (cx-4, cy) to (cx+4, cy)
    mov rcx, r12
    lea edx, [ebx-4]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+4]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    ; Bottom bar: (cx-6, cy+6) to (cx+6, cy+6)
    mov rcx, r12
    lea edx, [ebx-6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+6]
    call [LineTo]
    jmp .btn_drawn

.ico_wrap_text:
    ; 1. Top text line: (cx-7, cy-4) to (cx+1, cy-4)
    mov rdx, [hPenDark2]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-4]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-4]
    call [LineTo]

    ; 2. Bottom text line: (cx-7, cy+4) to (cx+1, cy+4)
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+4]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+4]
    call [LineTo]

    ; 3. Curved wrapping arrow in blue
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    ; Start at (cx+1, cy-2), line to (cx+5, cy-2)
    mov rcx, r12
    lea edx, [ebx+1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    call [LineTo]
    ; Downward curve: (cx+6, cy-1) -> (cx+6, cy+1) -> (cx+5, cy+2)
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-1]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+6]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+1]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+5]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    call [LineTo]
    ; Return line left to (cx-3, cy+2)
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    call [LineTo]
    ; Arrow head upper barb: to (cx-1, cy)
    mov rcx, r12
    lea edx, [ebx-1]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    ; Arrow head lower barb: (cx-3, cy+2) to (cx-1, cy+4)
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-1]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+4]
    call [LineTo]
    jmp .render_btn_text

.ico_merge_center:
    ; 1. White cell rectangle with border
    mov rdx, [hPenBorder]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hBrWhite]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-7]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    lea r9d, [ebx+7]
    lea eax, [eax+5]
    mov [rsp+20h], eax
    call [Rectangle]

    ; 2. Divider ticks on top and bottom showing cell merge boundary
    mov rdx, [hPenBorder]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-5]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax-3]
    call [LineTo]

    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax+3]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, ebx
    mov eax, [rbp-5Ch]
    lea r8d, [eax+5]
    call [LineTo]

    ; 3. Left blue arrow pointing right to center (ebx-1, cy)
    mov rdx, [hPenBlue]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    lea edx, [ebx-5]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-1]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx-1]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx-3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    call [LineTo]

    ; 4. Right blue arrow pointing left to center (ebx+1, cy)
    mov rcx, r12
    lea edx, [ebx+5]
    mov r8d, [rbp-5Ch]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+1]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax-2]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx+1]
    mov r8d, [rbp-5Ch]
    call [LineTo]
    mov rcx, r12
    lea edx, [ebx+3]
    mov eax, [rbp-5Ch]
    lea r8d, [eax+2]
    call [LineTo]
    jmp .render_btn_text

.ico_fill_color:
    ; -------------------------------------------------------------------------
    ; Modern Excel Paint Bucket (Fill Color) with Dynamic Color Bar & Chevron
    ; -------------------------------------------------------------------------
    xor eax, eax
    test r14d, 1            ; ODS_SELECTED
    jz @f
    inc eax
@@:
    mov [rbp-60h], eax

    mov eax, [rbp-40h]
    add eax, [rbp-60h]
    mov [rbp-64h], eax      ; bx = left + offset

    mov eax, [rbp-44h]
    add eax, [rbp-60h]
    mov [rbp-68h], eax      ; by = top + offset

    ; 1. Draw subtle split divider line at bx + 20
    test r14d, 41h          ; ODS_HOTLIGHT or ODS_SELECTED?
    jz .fc_divider_normal
    mov rdx, [hPenDark]
    jmp .fc_select_divider_pen
.fc_divider_normal:
    mov rdx, [hPenBorder]
.fc_select_divider_pen:
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    mov edx, [rbp-64h]
    add edx, 20
    mov r8d, [rbp-68h]
    add r8d, 3
    xor r9d, r9d
    call [MoveToEx]

    mov rcx, r12
    mov edx, [rbp-64h]
    add edx, 20
    mov r8d, [rbp-68h]
    add r8d, 21
    call [LineTo]

.fc_draw_bucket:
    ; 2. Tilted Bucket Body: Polygon [(bx+6, by+9), (bx+12, by+4), (bx+14, by+8), (bx+8, by+13)]
    mov ebx, [rbp-64h]
    mov esi, [rbp-68h]

    lea edx, [ebx + 6]
    mov [pt_bucket_poly], edx
    lea edx, [esi + 9]
    mov [pt_bucket_poly + 4], edx

    lea edx, [ebx + 12]
    mov [pt_bucket_poly + 8], edx
    lea edx, [esi + 4]
    mov [pt_bucket_poly + 12], edx

    lea edx, [ebx + 14]
    mov [pt_bucket_poly + 16], edx
    lea edx, [esi + 8]
    mov [pt_bucket_poly + 20], edx

    lea edx, [ebx + 8]
    mov [pt_bucket_poly + 24], edx
    lea edx, [esi + 13]
    mov [pt_bucket_poly + 28], edx

    mov rdx, [hBrSlate]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea rdx, [pt_bucket_poly]
    mov r8d, 4
    call [Polygon]

    ; 3. Rim opening showing liquid paint: Polygon [(bx+4, by+7), (bx+6, by+9), (bx+12, by+4), (bx+10, by+2)]
    mov ecx, [cur_fill_color]
    call [CreateSolidBrush]
    mov [rsp+40h], rax      ; save paint brush

    mov rdx, [rsp+40h]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]

    mov ebx, [rbp-64h]
    mov esi, [rbp-68h]

    lea edx, [ebx + 4]
    mov [pt_bucket_poly], edx
    lea edx, [esi + 7]
    mov [pt_bucket_poly + 4], edx

    lea edx, [ebx + 6]
    mov [pt_bucket_poly + 8], edx
    lea edx, [esi + 9]
    mov [pt_bucket_poly + 12], edx

    lea edx, [ebx + 12]
    mov [pt_bucket_poly + 16], edx
    lea edx, [esi + 4]
    mov [pt_bucket_poly + 20], edx

    lea edx, [ebx + 10]
    mov [pt_bucket_poly + 24], edx
    lea edx, [esi + 2]
    mov [pt_bucket_poly + 28], edx

    mov rcx, r12
    lea rdx, [pt_bucket_poly]
    mov r8d, 4
    call [Polygon]

    ; 4. Wire Handle
    mov ebx, [rbp-64h]
    mov esi, [rbp-68h]

    mov rcx, r12
    lea edx, [ebx + 5]
    lea r8d, [esi + 6]
    xor r9d, r9d
    call [MoveToEx]

    mov rcx, r12
    lea edx, [ebx + 7]
    lea r8d, [esi + 2]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx + 11]
    lea r8d, [esi + 2]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx + 12]
    lea r8d, [esi + 4]
    call [LineTo]

    ; 5. Pouring paint / droplet: Polygon [(bx+3, by+10), (bx+5, by+10), (bx+4, by+14), (bx+2, by+12)]
    mov ebx, [rbp-64h]
    mov esi, [rbp-68h]

    lea edx, [ebx + 3]
    mov [pt_drop_poly], edx
    lea edx, [esi + 10]
    mov [pt_drop_poly + 4], edx

    lea edx, [ebx + 5]
    mov [pt_drop_poly + 8], edx
    lea edx, [esi + 10]
    mov [pt_drop_poly + 12], edx

    lea edx, [ebx + 4]
    mov [pt_drop_poly + 16], edx
    lea edx, [esi + 14]
    mov [pt_drop_poly + 20], edx

    lea edx, [ebx + 2]
    mov [pt_drop_poly + 24], edx
    lea edx, [esi + 12]
    mov [pt_drop_poly + 28], edx

    mov rdx, [rsp+40h]      ; paint brush
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    lea rdx, [pt_drop_poly]
    mov r8d, 4
    call [Polygon]

    ; 6. Dynamic Fill Color Bar: [bx+3, by+17, bx+18, by+21] (height 4px)
    mov ebx, [rbp-64h]
    mov esi, [rbp-68h]

    lea eax, [ebx + 3]
    mov [rc_color_bar], eax
    lea eax, [esi + 17]
    mov [rc_color_bar + 4], eax
    lea eax, [ebx + 18]
    mov [rc_color_bar + 8], eax
    lea eax, [esi + 21]
    mov [rc_color_bar + 12], eax

    mov rdx, [rsp+40h]      ; paint brush
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenNull]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    mov edx, [rc_color_bar]
    mov r8d, [rc_color_bar + 4]
    mov r9d, [rc_color_bar + 8]
    mov eax, [rc_color_bar + 12]
    mov [rsp+20h], eax
    call [Rectangle]

    ; Crisp border around color bar
    mov rdx, [hPenBorder]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, [rc_color_bar]
    mov r8d, [rc_color_bar + 4]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, [rc_color_bar + 8]
    mov r8d, [rc_color_bar + 4]
    call [LineTo]
    mov rcx, r12
    mov edx, [rc_color_bar + 8]
    mov r8d, [rc_color_bar + 12]
    call [LineTo]
    mov rcx, r12
    mov edx, [rc_color_bar]
    mov r8d, [rc_color_bar + 12]
    call [LineTo]
    mov rcx, r12
    mov edx, [rc_color_bar]
    mov r8d, [rc_color_bar + 4]
    call [LineTo]

    mov rcx, [rsp+40h]
    call [DeleteObject]

    ; 7. Draw Dropdown Chevron ▼ at cx = bx + 25, cy = by + 11
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]

    mov ebx, [rbp-64h]
    add ebx, 25              ; cx
    mov esi, [rbp-68h]
    add esi, 11              ; cy

    ; Line 1 (5px wide): (cx-2, cy-1) to (cx+3, cy-1)
    mov rcx, r12
    lea edx, [ebx - 2]
    lea r8d, [esi - 1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx + 3]
    lea r8d, [esi - 1]
    call [LineTo]

    ; Line 2 (3px wide): (cx-1, cy) to (cx+2, cy)
    mov rcx, r12
    lea edx, [ebx - 1]
    mov r8d, esi
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx + 2]
    mov r8d, esi
    call [LineTo]

    ; Line 3 (1px wide): (cx, cy+1) to (cx+1, cy+1)
    mov rcx, r12
    mov edx, ebx
    lea r8d, [esi + 1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx + 1]
    lea r8d, [esi + 1]
    call [LineTo]
    jmp .btn_drawn

.ico_text_color:
    ; -------------------------------------------------------------------------
    ; Modern Excel Font Color ('A') with Dynamic Color Bar & Chevron
    ; -------------------------------------------------------------------------
    xor eax, eax
    test r14d, 1
    jz @f
    inc eax
@@:
    mov [rbp-60h], eax

    mov eax, [rbp-40h]
    add eax, [rbp-60h]
    mov [rbp-64h], eax      ; bx = left + offset

    mov eax, [rbp-44h]
    add eax, [rbp-60h]
    mov [rbp-68h], eax      ; by = top + offset

    ; 1. Draw subtle split divider line at bx + 20
    test r14d, 41h          ; ODS_HOTLIGHT or ODS_SELECTED?
    jz .tc_divider_normal
    mov rdx, [hPenDark]
    jmp .tc_select_divider_pen
.tc_divider_normal:
    mov rdx, [hPenBorder]
.tc_select_divider_pen:
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    mov edx, [rbp-64h]
    add edx, 20
    mov r8d, [rbp-68h]
    add r8d, 3
    xor r9d, r9d
    call [MoveToEx]

    mov rcx, r12
    mov edx, [rbp-64h]
    add edx, 20
    mov r8d, [rbp-68h]
    add r8d, 21
    call [LineTo]

.tc_draw_glyph:
    ; 2. Render Typographic Bold 'A' using Segoe UI Bold Font (hFontBold)
    mov rcx, r12
    mov edx, 1              ; TRANSPARENT
    call [SetBkMode]

    mov rcx, r12
    mov edx, 001E1E1Eh      ; Dark Charcoal #1E1E1E
    call [SetTextColor]

    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]

    ; Glyph rect: [bx+2, by+1, bx+19, by+16]
    mov ebx, [rbp-64h]
    mov esi, [rbp-68h]

    lea eax, [ebx + 2]
    mov [rc_glyph_a], eax
    lea eax, [esi + 1]
    mov [rc_glyph_a + 4], eax
    lea eax, [ebx + 19]
    mov [rc_glyph_a + 8], eax
    lea eax, [esi + 16]
    mov [rc_glyph_a + 12], eax

    mov rcx, r12
    lea rdx, [szGlyphA]
    mov r8d, -1
    lea r9, [rc_glyph_a]
    mov dword [rsp+20h], 0825h ; DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX
    call [DrawTextA]

    ; 3. Dynamic Text Color Bar: [bx+3, by+17, bx+18, by+21] (height 4px)
    mov ebx, [rbp-64h]
    mov esi, [rbp-68h]

    lea eax, [ebx + 3]
    mov [rc_color_bar], eax
    lea eax, [esi + 17]
    mov [rc_color_bar + 4], eax
    lea eax, [ebx + 18]
    mov [rc_color_bar + 8], eax
    lea eax, [esi + 21]
    mov [rc_color_bar + 12], eax

    mov ecx, [cur_text_color]
    call [CreateSolidBrush]
    mov [rsp+40h], rax

    mov rdx, [rsp+40h]
    mov rcx, r12
    call [SelectObject]
    mov rdx, [hPenNull]
    mov rcx, r12
    call [SelectObject]

    mov rcx, r12
    mov edx, [rc_color_bar]
    mov r8d, [rc_color_bar + 4]
    mov r9d, [rc_color_bar + 8]
    mov eax, [rc_color_bar + 12]
    mov [rsp+20h], eax
    call [Rectangle]

    ; Crisp border around color bar
    mov rdx, [hPenBorder]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, [rc_color_bar]
    mov r8d, [rc_color_bar + 4]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, [rc_color_bar + 8]
    mov r8d, [rc_color_bar + 4]
    call [LineTo]
    mov rcx, r12
    mov edx, [rc_color_bar + 8]
    mov r8d, [rc_color_bar + 12]
    call [LineTo]
    mov rcx, r12
    mov edx, [rc_color_bar]
    mov r8d, [rc_color_bar + 12]
    call [LineTo]
    mov rcx, r12
    mov edx, [rc_color_bar]
    mov r8d, [rc_color_bar + 4]
    call [LineTo]

    mov rcx, [rsp+40h]
    call [DeleteObject]

    ; 4. Draw Dropdown Chevron ▼ at cx = bx + 25, cy = by + 11
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]

    mov ebx, [rbp-64h]
    add ebx, 25              ; cx
    mov esi, [rbp-68h]
    add esi, 11              ; cy

    ; Line 1 (5px wide): (cx-2, cy-1) to (cx+3, cy-1)
    mov rcx, r12
    lea edx, [ebx - 2]
    lea r8d, [esi - 1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx + 3]
    lea r8d, [esi - 1]
    call [LineTo]

    ; Line 2 (3px wide): (cx-1, cy) to (cx+2, cy)
    mov rcx, r12
    lea edx, [ebx - 1]
    mov r8d, esi
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx + 2]
    mov r8d, esi
    call [LineTo]

    ; Line 3 (1px wide): (cx, cy+1) to (cx+1, cy+1)
    mov rcx, r12
    mov edx, ebx
    lea r8d, [esi + 1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx + 1]
    lea r8d, [esi + 1]
    call [LineTo]
    jmp .btn_drawn

.render_btn_text:
    ; Read window text from button into safe static buffer
    mov rcx, r15
    lea rdx, [szBtnDrawText]
    mov r8d, 60
    call [GetWindowTextA]

    mov rcx, r12
    mov edx, 1              ; TRANSPARENT
    call [SetBkMode]

    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]

    ; Check text color by button ID
    cmp r13d, ID_BTN_PASTE
    je .txt_accent_white
    cmp r13d, ID_BTN_FX
    je .txt_accent_white
    cmp r13d, ID_BTN_SAVE
    je .txt_accent_green
    cmp r13d, ID_BTN_COND_FMT
    je .txt_accent_teal
    cmp r13d, ID_BTN_AUTOSUM
    je .txt_accent_blue
    cmp r13d, ID_BTN_FORM_AUTOSUM
    je .txt_accent_blue

    ; Default dark text: #202020
    mov rcx, r12
    mov edx, 00202020h
    call [SetTextColor]
    jmp .txt_layout

.txt_accent_white:
    mov rcx, r12
    mov edx, 00FFFFFFh
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]
    jmp .txt_layout

.txt_accent_green:
    mov rcx, r12
    mov edx, 00467321h      ; #217346
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]
    jmp .txt_layout

.txt_accent_teal:
    mov rcx, r12
    mov edx, 00C78402h      ; #0284C7
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]
    jmp .txt_layout

.txt_accent_blue:
    mov rcx, r12
    mov edx, 00D2662Bh      ; #2B66D2
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]
    jmp .txt_layout

.txt_layout:
    cmp dword [rbp-54h], 45
    jl .txt_small_btn

    ; Tall button text rect: [left+2, top+34, right-2, bottom-2]
    mov eax, [rbp-40h]
    add eax, 2
    mov [rcBtnDrawText], eax
    mov eax, [rbp-44h]
    add eax, 34
    mov [rcBtnDrawText+4], eax
    mov eax, [rbp-48h]
    sub eax, 2
    mov [rcBtnDrawText+8], eax
    mov eax, [rbp-4Ch]
    sub eax, 2
    cmp r13d, ID_BTN_PASTE
    jne .tall_paste_noadj
    sub eax, 8
.tall_paste_noadj:
    mov [rcBtnDrawText+12], eax

    mov rcx, r12
    lea rdx, [szBtnDrawText]
    mov r8d, -1
    lea r9, [rcBtnDrawText]
    mov dword [rsp+20h], 0811h ; DT_CENTER | DT_WORDBREAK | DT_NOPREFIX
    call [DrawTextA]

    cmp r13d, ID_BTN_PASTE
    je .draw_paste_chevron
    jmp .btn_drawn

.txt_small_btn:
    cmp r13d, ID_BTN_COND_FMT
    je .txt_cond_fmt_btn

    ; Check if button needs right chevron:
    cmp r13d, ID_BTN_AUTOSUM
    je .txt_small_has_chev
    cmp r13d, ID_BTN_FORM_AUTOSUM
    je .txt_small_has_chev
    cmp r13d, ID_BTN_INSERT_CELL
    je .txt_small_has_chev
    cmp r13d, ID_BTN_DELETE_CELL
    je .txt_small_has_chev
    cmp r13d, ID_BTN_CLEAR
    je .txt_small_has_chev
    cmp r13d, ID_BTN_SORT_FILTER
    je .txt_small_has_chev
    cmp r13d, ID_BTN_FIND_SELECT
    je .txt_small_has_chev

    ; Button WITHOUT chevron: text rect is [left+24, top, right-2, bottom]
    mov eax, [rbp-40h]
    add eax, 24
    mov [rcBtnDrawText], eax
    mov eax, [rbp-44h]
    mov [rcBtnDrawText+4], eax
    mov eax, [rbp-48h]
    sub eax, 2
    mov [rcBtnDrawText+8], eax
    mov eax, [rbp-4Ch]
    mov [rcBtnDrawText+12], eax

    mov rcx, r12
    lea rdx, [szBtnDrawText]
    mov r8d, -1
    lea r9, [rcBtnDrawText]
    mov dword [rsp+20h], 0824h ; DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX
    call [DrawTextA]
    jmp .btn_drawn

.txt_small_has_chev:
    ; Small button text rect with chevron: [left+24, top, right-14, bottom]
    mov eax, [rbp-40h]
    add eax, 24
    mov [rcBtnDrawText], eax
    mov eax, [rbp-44h]
    mov [rcBtnDrawText+4], eax
    mov eax, [rbp-48h]
    sub eax, 14
    mov [rcBtnDrawText+8], eax
    mov eax, [rbp-4Ch]
    mov [rcBtnDrawText+12], eax

    mov rcx, r12
    lea rdx, [szBtnDrawText]
    mov r8d, -1
    lea r9, [rcBtnDrawText]
    mov dword [rsp+20h], 0824h ; DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX
    call [DrawTextA]

    cmp r13d, ID_BTN_AUTOSUM
    je .draw_right_chevron_blue
    cmp r13d, ID_BTN_FORM_AUTOSUM
    je .draw_right_chevron_blue
    jmp .draw_right_chevron_dark

.txt_cond_fmt_btn:
    ; 2-line Conditional Formatting: [left+24, top+2, right-14, bottom-2]
    mov eax, [rbp-40h]
    add eax, 24
    mov [rcBtnDrawText], eax
    mov eax, [rbp-44h]
    add eax, 2
    mov [rcBtnDrawText+4], eax
    mov eax, [rbp-48h]
    sub eax, 14
    mov [rcBtnDrawText+8], eax
    mov eax, [rbp-4Ch]
    sub eax, 2
    mov [rcBtnDrawText+12], eax

    mov rcx, r12
    lea rdx, [szBtnDrawText]
    mov r8d, -1
    lea r9, [rcBtnDrawText]
    mov dword [rsp+20h], 0810h ; DT_LEFT | DT_WORDBREAK | DT_NOPREFIX
    call [DrawTextA]
    jmp .draw_right_chevron_teal

.draw_paste_chevron:
    mov rdx, [hPenWhite]
    mov rcx, r12
    call [SelectObject]

    mov ebx, [rbp-58h]       ; center x
    mov esi, [rbp-4Ch]       ; bottom
    sub esi, 6               ; cy

    mov rcx, r12
    lea edx, [ebx - 2]
    lea r8d, [esi - 1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx + 3]
    lea r8d, [esi - 1]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx - 1]
    mov r8d, esi
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx + 2]
    mov r8d, esi
    call [LineTo]

    mov rcx, r12
    mov edx, ebx
    lea r8d, [esi + 1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx + 1]
    lea r8d, [esi + 1]
    call [LineTo]
    jmp .btn_drawn

.draw_right_chevron_teal:
    mov rdx, [hPenAccentTeal]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_right_chevron_core

.draw_right_chevron_blue:
    mov rdx, [hPenAccentBlue]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_right_chevron_core

.draw_right_chevron_dark:
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    jmp .draw_right_chevron_core

.draw_right_chevron_core:
    mov ebx, [rbp-48h]       ; right
    sub ebx, 8               ; cx
    mov esi, [rbp-5Ch]       ; center y

    mov rcx, r12
    lea edx, [ebx - 2]
    lea r8d, [esi - 1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx + 3]
    lea r8d, [esi - 1]
    call [LineTo]

    mov rcx, r12
    lea edx, [ebx - 1]
    mov r8d, esi
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx + 2]
    mov r8d, esi
    call [LineTo]

    mov rcx, r12
    mov edx, ebx
    lea r8d, [esi + 1]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    lea edx, [ebx + 1]
    lea r8d, [esi + 1]
    call [LineTo]
    jmp .btn_drawn

.btn_drawn:
    mov rbx, [rbp-08h]
    mov r12, [rbp-10h]
    mov r13, [rbp-18h]
    mov r14, [rbp-20h]
    mov r15, [rbp-28h]
    mov rsi, [rbp-30h]
    mov rdi, [rbp-38h]
    leave
    ret


; =============================================================================
; Build Insert Tab Controls
; =============================================================================
BuildInsertTabControls:
    push rbp
    mov rbp, rsp
    sub rsp, 60h

    ; Group 1: Tables (x=10, y=26, w=115, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_ins_tab_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 10
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 115
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_INS_TABLES
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Table button (x=16, y=42, w=103, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_ins_tab_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 103
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INS_TABLE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; PivotTable button (x=16, y=72, w=103, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_ins_piv_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 103
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INS_PIVOT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Group 2: Illustrations (x=131, y=26, w=185, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_ins_ill_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 131
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 185
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_INS_ILLUST
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Pictures (x=137, y=42, w=76, h=58)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_ins_pic_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 137
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 76
    mov dword [rsp+38h], 58
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INS_PICTURE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; Shapes (x=219, y=42, w=90, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_ins_shp_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 219
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 90
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INS_SHAPES
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Icons (x=219, y=72, w=90, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_ins_ico_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 219
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 90
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INS_ICONS
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Group 3: Charts (x=322, y=26, w=185, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_ins_cht_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 322
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 185
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_INS_CHARTS
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Column Chart (x=328, y=42, w=85, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_ins_col_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 328
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 85
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INS_COL_CHART
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Line Chart (x=417, y=42, w=84, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_ins_lin_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 417
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 84
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INS_LINE_CHART
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Pie Chart (x=328, y=72, w=85, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_ins_pie_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 328
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 85
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INS_PIE_CHART
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Bar Chart (x=417, y=72, w=84, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_ins_bar_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 417
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 84
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INS_BAR_CHART
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Group 4: Links && Text (x=513, y=26, w=175, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_ins_lnk_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 513
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 175
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_INS_LINKS
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Link (x=519, y=42, w=64, h=58)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_ins_lnk_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 519
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 64
    mov dword [rsp+38h], 58
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INS_LINK
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; Text Box (x=589, y=42, w=93, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_ins_txb_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 589
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 93
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INS_TEXTBOX
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Symbol (x=589, y=72, w=93, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_ins_sym_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 589
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 93
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_INS_SYMBOL
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    leave
    ret


; =============================================================================
; Build Data Tab Controls
; =============================================================================
BuildDataTabControls:
    push rbp
    mov rbp, rsp
    sub rsp, 60h

    ; Group 1: Get && Transform (x=10, y=26, w=125, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_data_get_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 10
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 125
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_DATA_GET
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; From CSV (x=16, y=42, w=113, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_data_csv_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 113
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DATA_CSV
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Refresh All (x=16, y=72, w=113, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_data_ref_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 113
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DATA_REFRESH
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Group 2: Sort && Filter (x=141, y=26, w=215, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_data_srt_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 141
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 215
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_DATA_SORT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Sort A-Z (x=147, y=42, w=90, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_data_asc_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 147
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 90
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DATA_SORT_ASC
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Sort Z-A (x=147, y=72, w=90, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_data_dsc_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 147
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 90
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DATA_SORT_DESC
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Filter (x=243, y=42, w=107, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_data_flt_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 243
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 107
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DATA_FILTER
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Clear Filter (x=243, y=72, w=107, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_data_clr_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 243
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 107
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DATA_CLEAR
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Group 3: Data Tools (x=362, y=26, w=220, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_data_tls_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 362
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 220
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_DATA_TOOLS
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Text to Columns (x=368, y=42, w=95, h=58)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_data_spl_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 368
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 95
    mov dword [rsp+38h], 58
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DATA_SPLIT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; Remove Duplicates (x=469, y=42, w=107, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_data_dup_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 469
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 107
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DATA_REMDUP
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Data Validation (x=469, y=72, w=107, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_data_val_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 469
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 107
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DATA_VALID
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Group 4: Forecast && Analysis (x=588, y=26, w=155, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_data_fcs_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 588
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 155
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_DATA_ANALYSIS
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; SIMD Benchmark (x=594, y=42, w=143, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_data_sim_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 594
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 143
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DATA_BENCH
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; What-If (x=594, y=72, w=143, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_data_wif_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 594
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 143
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_DATA_WHATIF
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    leave
    ret


; =============================================================================
; Build Formulas Tab Controls
; =============================================================================
BuildFormulasTabControls:
    push rbp
    mov rbp, rsp
    sub rsp, 60h

    ; Group 1: Function Library (x=10, y=26, w=530, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_form_lib_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 10
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 530
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_FORM_LIB
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Insert Function [fx] (x=16, y=42, w=76, h=58)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_form_fx_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 76
    mov dword [rsp+38h], 58
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FORM_FX
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; AutoSum (x=98, y=42, w=76, h=58)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_form_sum_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 98
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 76
    mov dword [rsp+38h], 58
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FORM_AUTOSUM
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; Financial (x=180, y=42, w=76, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_form_fin_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 180
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 76
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FORM_FIN
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Logical (x=180, y=72, w=76, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_form_log_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 180
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 76
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FORM_LOGIC
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Text (x=262, y=42, w=76, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_form_txt_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 262
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 76
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FORM_TEXT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Date/Time (x=262, y=72, w=76, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_form_dat_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 262
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 76
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FORM_DATE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Lookup (x=344, y=42, w=76, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_form_luk_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 344
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 76
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FORM_LOOKUP
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Math/Trig (x=344, y=72, w=76, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_form_mth_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 344
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 76
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FORM_MATH
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; More fx (x=426, y=42, w=108, h=58)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_form_mor_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 426
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 108
    mov dword [rsp+38h], 58
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FORM_MORE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; Group 2: Defined Names (x=546, y=26, w=125, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_form_nam_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 546
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 125
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_FORM_NAMES
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Name Mgr (x=552, y=42, w=113, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_form_nm_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 552
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 113
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FORM_NAME_MGR
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Define Name (x=552, y=72, w=113, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_form_df_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 552
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 113
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FORM_DEF_NAME
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Group 3: Calculation (x=677, y=26, w=130, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_form_clc_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 677
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 130
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_FORM_CALC
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Calculate (x=683, y=42, w=118, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_form_cal_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 683
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 118
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FORM_CALC_NOW
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Eval Formula (x=683, y=72, w=118, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_form_evl_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 683
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 118
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FORM_EVAL
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    leave
    ret


; =============================================================================
; Build View Tab Controls
; =============================================================================
BuildViewTabControls:
    push rbp
    mov rbp, rsp
    sub rsp, 60h

    ; Group 1: Workbook Views (x=10, y=26, w=125, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_view_vws_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 10
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 125
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_VIEW_VIEWS
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Normal (x=16, y=42, w=113, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_view_nrm_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 113
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_VIEW_NORMAL
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Page Layout (x=16, y=72, w=113, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_view_lay_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 113
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_VIEW_PAGELAY
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Group 2: Show && Hide (x=141, y=26, w=220, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_view_shw_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 141
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 220
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_VIEW_SHOW
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Gridlines (x=147, y=42, w=92, h=58)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_view_grd_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 147
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 92
    mov dword [rsp+38h], 58
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_VIEW_GRIDLINES
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; Formula Bar (x=245, y=42, w=110, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_view_bar_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 245
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 110
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_VIEW_FORMBAR
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Headings (x=245, y=72, w=110, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_view_hdg_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 245
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 110
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_VIEW_HEADINGS
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Group 3: Zoom (x=367, y=26, w=185, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_view_zom_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 367
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 185
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_VIEW_ZOOM
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; 100% (x=373, y=42, w=60, h=58)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_view_z100_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 373
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 60
    mov dword [rsp+38h], 58
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_VIEW_ZOOM100
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; Zoom In (+) (x=439, y=42, w=107, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_view_zin_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 439
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 107
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_VIEW_ZOOM_IN
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Zoom Out (-) (x=439, y=72, w=107, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_view_zot_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 439
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 107
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_VIEW_ZOOM_OUT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Group 4: Window (x=558, y=26, w=125, h=88)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_view_wnd_txt]
    mov r9d, WS_CHILD or BS_GROUPBOX
    mov dword [rsp+20h], 558
    mov dword [rsp+28h], 26
    mov dword [rsp+30h], 125
    mov dword [rsp+38h], 88
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRP_VIEW_WINDOW
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; Freeze Panes (x=564, y=42, w=113, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_view_frz_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 564
    mov dword [rsp+28h], 42
    mov dword [rsp+30h], 113
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_VIEW_FREEZE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    ; New Window (x=564, y=72, w=113, h=28)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_view_nw_txt]
    mov r9d, WS_CHILD or BS_OWNERDRAW
    mov dword [rsp+20h], 564
    mov dword [rsp+28h], 72
    mov dword [rsp+30h], 113
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_VIEW_NEW_WND
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontSmall]
    mov r9d, 1
    call [SendMessageA]

    leave
    ret


; =============================================================================
; Build Formula Bar (Name Box, [fx], Formula Edit)
; =============================================================================
BuildFormulaBar:
    push rbp
    mov rbp, rsp
    sub rsp, 60h

    ; Name Box (A1)
    mov ecx, WS_EX_CLIENTEDGE
    lea rdx, [edit_class]
    lea r8, [init_cell_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or WS_BORDER or ES_READONLY
    mov dword [rsp+20h], 6
    mov dword [rsp+28h], 122
    mov dword [rsp+30h], 70
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_EDIT_NAME_BOX
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov [hNameBox], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; [fx] Button
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [fx_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_OWNERDRAW
    mov dword [rsp+20h], 80
    mov dword [rsp+28h], 122
    mov dword [rsp+30h], 32
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_BTN_FX
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov [hBtnFx], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; Formula Entry Bar
    mov ecx, WS_EX_CLIENTEDGE
    lea rdx, [edit_class]
    lea r8, [init_form_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or WS_BORDER or ES_AUTOHSCROLL
    mov dword [rsp+20h], 116
    mov dword [rsp+28h], 122
    mov dword [rsp+30h], 1144
    mov dword [rsp+38h], 24
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_EDIT_FORMULA
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov [hFormulaEdit], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontMono]
    mov r9d, 1
    call [SendMessageA]

    ; Subclass hFormulaEdit to handle VK_RETURN
    mov rcx, [hFormulaEdit]
    mov edx, -4 ; GWLP_WNDPROC
    lea r8, [FormulaEditProc]
    call [SetWindowLongPtrA]
    mov [origFormulaProc], rax

    leave
    ret


; =============================================================================
; Subclass Procedure for Formula Bar Edit Control (Captures VK_RETURN)
; =============================================================================
FormulaEditProc:
    push rbp
    mov rbp, rsp
    sub rsp, 40h

    cmp edx, 0100h ; WM_KEYDOWN
    jne .call_orig
    cmp r8d, 0Dh   ; VK_RETURN
    jne .call_orig

    call DoEvalFormula

    mov rcx, [hGrid]
    call [SetFocus]
    xor eax, eax
    leave
    ret

.call_orig:
    mov rax, [origFormulaProc]
    mov [rsp+20h], r9
    mov r9, r8
    mov r8, rdx
    mov rdx, rcx
    mov rcx, rax
    call [CallWindowProcA]
    leave
    ret


; =============================================================================
; Switch Ribbon Tab
; RCX = new_tab_index (0..4: Home, Insert, Data, Formulas, View)
; =============================================================================
SwitchRibbonTab:
    push rbp
    mov rbp, rsp
    sub rsp, 30h
    mov [rbp-8], rcx

    ; Save active tab index
    mov eax, ecx
    mov [cur_ribbon_tab], eax

    ; Tab 0: Home
    mov edx, 0 ; SW_HIDE
    cmp qword [rbp-8], 0
    jne @f
    mov edx, 5 ; SW_SHOW
@@:
    lea r8, [ids_home]
    call SetTabGroupVisibility

    ; Tab 1: Insert
    mov edx, 0
    cmp qword [rbp-8], 1
    jne @f
    mov edx, 5
@@:
    lea r8, [ids_insert]
    call SetTabGroupVisibility

    ; Tab 2: Data
    mov edx, 0
    cmp qword [rbp-8], 2
    jne @f
    mov edx, 5
@@:
    lea r8, [ids_data]
    call SetTabGroupVisibility

    ; Tab 3: Formulas
    mov edx, 0
    cmp qword [rbp-8], 3
    jne @f
    mov edx, 5
@@:
    lea r8, [ids_formulas]
    call SetTabGroupVisibility

    ; Tab 4: View
    mov edx, 0
    cmp qword [rbp-8], 4
    jne @f
    mov edx, 5
@@:
    lea r8, [ids_view]
    call SetTabGroupVisibility

    ; Invalidate ribbon toolbar area to clear old buttons and draw new
    mov rcx, [hWnd]
    lea rdx, [rcRibbonArea]
    mov r8d, 1
    call [InvalidateRect]

    mov rcx, [hRibbonTabs]
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]

    mov rcx, [hWnd]
    call [UpdateWindow]

    leave
    ret


; =============================================================================
; SetTabGroupVisibility
; R8 = pointer to zero-terminated array of DWORD control IDs
; EDX = nCmdShow (0 = SW_HIDE, 5 = SW_SHOW)
; =============================================================================
SetTabGroupVisibility:
    push rbp
    mov rbp, rsp
    sub rsp, 40h
    mov [rbp-8], rdx
    mov [rbp-16], r8

.loop_vis_ids:
    mov r8, [rbp-16]
    mov edx, dword [r8]
    test edx, edx
    jz .done_vis_ids

    mov rcx, [hWnd]
    call [GetDlgItem]
    test rax, rax
    jz .skip_vis_ctrl

    mov rcx, rax
    mov edx, dword [rbp-8]
    call [ShowWindow]

    cmp dword [rbp-8], 0
    je .skip_vis_ctrl

    mov r8, [rbp-16]
    mov edx, dword [r8]
    mov rcx, [hWnd]
    call [GetDlgItem]
    test rax, rax
    jz .skip_vis_ctrl

    mov rcx, rax
    xor edx, edx ; HWND_TOP = 0
    xor r8d, r8d
    xor r9d, r9d
    sub rsp, 30h
    mov dword [rsp+20h], 0
    mov dword [rsp+28h], 0
    mov dword [rsp+30h], 0043h ; SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW
    call [SetWindowPos]
    add rsp, 30h

.skip_vis_ctrl:
    add qword [rbp-16], 4
    jmp .loop_vis_ids

.done_vis_ids:
    leave
    ret


; =============================================================================
; Show Status Message in StatusBar
; RCX = pointer to string
; =============================================================================
ShowInfoStatus:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    mov r9, rcx
    mov rcx, [hStatusBar]
    test rcx, rcx
    jz .done_sis
    mov edx, SB_SETTEXTA
    mov r8d, 2
    call [SendMessageA]
.done_sis:
    leave
    ret

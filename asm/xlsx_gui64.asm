; =============================================================================
; Excel Viewer Pro — 100% Pure Native 64-bit Assembly Spreadsheet Engine
; Modular Architecture (FASM x64) — Refactored according to svcmgr design
; =============================================================================

format PE64 GUI 5.0
entry start

include 'const.inc'
include 'data.inc'

section '.code' code readable executable


start:
    sub rsp, 28h

    ; Init Common Controls (ListView + Tabs + Status)
    mov dword [icc], 8
    mov dword [icc+4], 0FFFFh
    lea rcx, [icc]
    call [InitCommonControlsEx]

    xor ecx, ecx
    call [GetModuleHandleA]
    mov [hInst], rax

    ; Set up WNDCLASSEX
    lea rdi, [wc]
    xor eax, eax
    mov ecx, 80
    rep stosb

    mov dword [wc], 80
    mov dword [wc+4], CS_HREDRAW or CS_VREDRAW
    lea rax, [WndProc]
    mov qword [wc+8], rax
    mov rax, [hInst]
    mov qword [wc+24], rax

    ; Load Application Icon (PE embedded resource 1 or app_icon.ico)
    mov rcx, [hInst]
    mov edx, 1
    call [LoadIconA]
    test rax, rax
    jnz .icon_ready

    xor ecx, ecx
    lea rdx, [szIconFile]
    mov r8d, 1 ; IMAGE_ICON
    xor r9d, r9d
    sub rsp, 30h
    mov dword [rsp+20h], 0
    mov dword [rsp+28h], 0050h ; LR_LOADFROMFILE or LR_DEFAULTSIZE
    call [LoadImageA]
    add rsp, 30h
    test rax, rax
    jnz .icon_ready

    xor ecx, ecx
    lea rdx, [szIconFileAsm]
    mov r8d, 1
    xor r9d, r9d
    sub rsp, 30h
    mov dword [rsp+20h], 0
    mov dword [rsp+28h], 0050h
    call [LoadImageA]
    add rsp, 30h
    test rax, rax
    jnz .icon_ready

    xor ecx, ecx
    lea rdx, [szIconFileParent]
    mov r8d, 1
    xor r9d, r9d
    sub rsp, 30h
    mov dword [rsp+20h], 0
    mov dword [rsp+28h], 0050h
    call [LoadImageA]
    add rsp, 30h
    test rax, rax
    jnz .icon_ready

    xor ecx, ecx
    lea rdx, [szIconFilePkg]
    mov r8d, 1
    xor r9d, r9d
    sub rsp, 30h
    mov dword [rsp+20h], 0
    mov dword [rsp+28h], 0050h
    call [LoadImageA]
    add rsp, 30h
    test rax, rax
    jnz .icon_ready

    xor ecx, ecx
    mov edx, 32512 ; IDI_APPLICATION
    call [LoadIconA]

.icon_ready:
    mov [hAppIcon], rax
    mov qword [wc+32], rax
    mov qword [wc+72], rax

    xor ecx, ecx
    mov edx, 32512 ; IDC_ARROW
    call [LoadCursorA]
    mov qword [wc+40], rax
    mov qword [wc+48], COLOR_BTNFACE + 1
    lea rax, [class_name]
    mov qword [wc+64], rax

    lea rcx, [wc]
    call [RegisterClassExA]
    test eax, eax
    jz .exit

    ; Create Modern System Fonts
    call CreateAppFonts

    ; Allocate Cell Style Table (10000 rows * 32 cols * 16 bytes = 5,120,000 bytes)
    mov ecx, GMEM_ZEROINIT
    mov edx, 10000 * 32 * 16
    call [GlobalAlloc]
    mov [pCellStyleTable], rax

    ; Allocate Cell Formula Pointer Table (10000 rows * 32 cols * 8 bytes = 2,560,000 bytes)
    mov ecx, GMEM_ZEROINIT
    mov edx, 10000 * 32 * 8
    call [GlobalAlloc]
    mov [pCellFormulaTable], rax

    ; Create Dedicated Formula Heap
    xor ecx, ecx
    xor edx, edx
    xor r8d, r8d
    call [HeapCreate]
    mov [hFormulaHeap], rax

    ; Create Main Window (1280 x 820)
    sub rsp, 60h
    xor ecx, ecx
    lea rdx, [class_name]
    lea r8, [wnd_title]
    mov r9d, WS_OVERLAPPEDWINDOW or WS_VISIBLE
    mov dword [rsp+20h], 60
    mov dword [rsp+28h], 40
    mov dword [rsp+30h], 1280
    mov dword [rsp+38h], 820
    mov qword [rsp+40h], 0
    mov qword [rsp+48h], 0
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    add rsp, 60h
    test rax, rax
    jz .exit
    mov [hWnd], rax

    ; Set Titlebar and Taskbar icons
    mov rcx, [hWnd]
    mov edx, 0080h ; WM_SETICON
    mov r8d, 1     ; ICON_BIG
    mov r9, [hAppIcon]
    call [SendMessageA]

    mov rcx, [hWnd]
    mov edx, 0080h ; WM_SETICON
    xor r8d, r8d   ; ICON_SMALL
    mov r9, [hAppIcon]
    call [SendMessageA]

    ; Initialize Ribbon GDI resources
    call InitRibbonGdiResources

    ; Create Menu, Ribbon, Formula Bar, Grid & Status
    call BuildAppMenu
    call BuildRibbonBar
    call BuildFormulaBar
    call BuildSpreadsheetGrid
    call BuildSheetTabsAndStatus

    mov rcx, [hWnd]
    mov edx, 1 ; SW_SHOWNORMAL
    call [ShowWindow]
    mov rcx, [hWnd]
    call [UpdateWindow]

    ; Check command-line argument for spreadsheet file to open
    call [GetCommandLineA]
    test rax, rax
    jz .msg_loop

    mov rsi, rax
    ; Skip executable name
    cmp byte [rsi], '"'
    jne .skip_exe_no_quotes
    inc rsi
.skip_exe_q:
    lodsb
    test al, al
    jz .msg_loop
    cmp al, '"'
    jne .skip_exe_q
    jmp .skip_exe_ws
.skip_exe_no_quotes:
    lodsb
    test al, al
    jz .msg_loop
    cmp al, ' '
    jne .skip_exe_no_quotes
.skip_exe_ws:
    lodsb
    cmp al, ' '
    je .skip_exe_ws
    test al, al
    jz .msg_loop

    dec rsi
    ; Strip leading quote if any
    cmp byte [rsi], '"'
    jne @f
    inc rsi
@@:
    lea rdi, [szFileName]
    mov ecx, 500
.copy_cmd_arg:
    lodsb
    cmp al, '"'
    je .done_cmd_arg
    test al, al
    jz .done_cmd_arg
    stosb
    dec ecx
    jnz .copy_cmd_arg
.done_cmd_arg:
    mov byte [rdi], 0

    cmp byte [szFileName], 0
    je .msg_loop
    call DoFileOpenDirect

.msg_loop:
    lea rcx, [msgbuf]
    xor edx, edx
    xor r8d, r8d
    xor r9d, r9d
    call [GetMessageA]
    cmp eax, 0
    jle .exit

    lea rcx, [msgbuf]
    call [TranslateMessage]
    lea rcx, [msgbuf]
    call [DispatchMessageA]
    jmp .msg_loop

.exit:
    xor ecx, ecx
    call [ExitProcess]


; =============================================================================
; Create Application Fonts
; =============================================================================
CreateAppFonts:
    push rbp
    mov rbp, rsp
    sub rsp, 80h

    ; UI Normal Font
    mov ecx, 13
    xor edx, edx
    xor r8d, r8d
    xor r9d, r9d
    mov dword [rsp+20h], 400
    mov dword [rsp+28h], 0
    mov dword [rsp+30h], 0
    mov dword [rsp+38h], 0
    mov dword [rsp+40h], 1
    mov dword [rsp+48h], 0
    mov dword [rsp+50h], 0
    mov dword [rsp+58h], 0
    mov dword [rsp+60h], 0
    lea rax, [font_ui_name]
    mov qword [rsp+68h], rax
    call [CreateFontA]
    mov [hFontUI], rax

    ; UI Small Font (9pt)
    mov ecx, 11
    xor edx, edx
    xor r8d, r8d
    xor r9d, r9d
    mov dword [rsp+20h], 400
    mov dword [rsp+28h], 0
    mov dword [rsp+30h], 0
    mov dword [rsp+38h], 0
    mov dword [rsp+40h], 1
    mov dword [rsp+48h], 0
    mov dword [rsp+50h], 0
    mov dword [rsp+58h], 0
    mov dword [rsp+60h], 0
    lea rax, [font_ui_name]
    mov qword [rsp+68h], rax
    call [CreateFontA]
    mov [hFontSmall], rax

    ; UI Bold Font
    mov ecx, 14
    xor edx, edx
    xor r8d, r8d
    xor r9d, r9d
    mov dword [rsp+20h], 700
    mov dword [rsp+28h], 0
    mov dword [rsp+30h], 0
    mov dword [rsp+38h], 0
    mov dword [rsp+40h], 1
    mov dword [rsp+48h], 0
    mov dword [rsp+50h], 0
    mov dword [rsp+58h], 0
    mov dword [rsp+60h], 0
    lea rax, [font_ui_name]
    mov qword [rsp+68h], rax
    call [CreateFontA]
    mov [hFontBold], rax

    ; UI Italic Font
    mov ecx, 13
    xor edx, edx
    xor r8d, r8d
    xor r9d, r9d
    mov dword [rsp+20h], 400
    mov dword [rsp+28h], 1 ; italic
    mov dword [rsp+30h], 0
    mov dword [rsp+38h], 0
    mov dword [rsp+40h], 1
    mov dword [rsp+48h], 0
    mov dword [rsp+50h], 0
    mov dword [rsp+58h], 0
    mov dword [rsp+60h], 0
    lea rax, [font_ui_name]
    mov qword [rsp+68h], rax
    call [CreateFontA]
    mov [hFontItalic], rax

    ; UI Bold Italic Font
    mov ecx, 14
    xor edx, edx
    xor r8d, r8d
    xor r9d, r9d
    mov dword [rsp+20h], 700
    mov dword [rsp+28h], 1 ; italic
    mov dword [rsp+30h], 0
    mov dword [rsp+38h], 0
    mov dword [rsp+40h], 1
    mov dword [rsp+48h], 0
    mov dword [rsp+50h], 0
    mov dword [rsp+58h], 0
    mov dword [rsp+60h], 0
    lea rax, [font_ui_name]
    mov qword [rsp+68h], rax
    call [CreateFontA]
    mov [hFontBoldItalic], rax

    ; Monospace Formula Font
    mov ecx, 14
    xor edx, edx
    xor r8d, r8d
    xor r9d, r9d
    mov dword [rsp+20h], 400
    mov dword [rsp+28h], 0
    mov dword [rsp+30h], 0
    mov dword [rsp+38h], 0
    mov dword [rsp+40h], 1
    mov dword [rsp+48h], 0
    mov dword [rsp+50h], 0
    mov dword [rsp+58h], 0
    mov dword [rsp+60h], 0
    lea rax, [font_mono_name]
    mov qword [rsp+68h], rax
    call [CreateFontA]
    mov [hFontMono], rax

    leave
    ret


; =============================================================================
; Window Procedure
; =============================================================================
WndProc:
    push rbp
    mov rbp, rsp
    sub rsp, 0B0h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13
    mov [rbp-48], r14
    mov [rbp-56], r15

    cmp edx, WM_COMMAND
    je .command
    cmp edx, WM_NOTIFY
    je .notify
    cmp edx, WM_DRAWITEM
    je .drawitem
    cmp edx, WM_CONTEXTMENU
    je .contextmenu
    cmp edx, WM_SETFOCUS
    je .setfocus
    cmp edx, WM_KEYDOWN
    je .keydown
    cmp edx, WM_SIZE
    je .size
    cmp edx, WM_DESTROY
    je .destroy

    call [DefWindowProcA]
    jmp .done

.contextmenu:
    mov rcx, [hContextMenu]
    test rcx, rcx
    jz .ret0

    ; Check if keyboard triggered (lParam == -1)
    cmp r9d, -1
    je .cm_keyboard

    ; Mouse triggered via WM_CONTEXTMENU:
    ; lParam contains screen coordinates: LOWORD=X, HIWORD=Y
    movsx r8d, r9w
    sar r9, 16
    movsx r9d, r9w
    jmp .cm_track

.cm_keyboard:
    ; Position menu at bottom-left corner of active cell
    lea rbx, [rcCell]
    mov dword [rbx], 0          ; LVIR_BOUNDS
    mov eax, [cur_col]
    mov dword [rbx+4], eax      ; col
    mov dword [rbx+8], 0
    mov dword [rbx+12], 0
    mov rcx, [hGrid]
    mov edx, 1038h              ; LVM_GETSUBITEMRECT
    mov r8d, [cur_row]
    mov r9, rbx
    call [SendMessageA]

    mov eax, [rbx]              ; left
    mov [rbp-70h], eax
    mov eax, [rbx+12]           ; bottom
    mov [rbp-6Ch], eax
    mov rcx, [hGrid]
    lea rdx, [rbp-70h]
    call [ClientToScreen]

    mov r8d, dword [rbp-70h]
    mov r9d, dword [rbp-6Ch]

.cm_track:
    mov rcx, [hContextMenu]
    xor edx, edx
    mov dword [rsp+20h], 0
    mov rax, [hWnd]
    mov qword [rsp+28h], rax
    mov qword [rsp+30h], 0
    call [TrackPopupMenu]
    xor eax, eax
    jmp .done

.drawitem:
    mov rsi, r9             ; LPDRAWITEMSTRUCT
    call DrawRibbonOwnerButton
    mov eax, 1
    jmp .done

.setfocus:
    mov rcx, [hGrid]
    test rcx, rcx
    jz .ret0
    call [SetFocus]
    xor eax, eax
    jmp .done

.destroy:
    mov rcx, [hContextMenu]
    test rcx, rcx
    jz @f
    call [DestroyMenu]
@@:
    call CleanupRibbonGdiResources
    call CleanupTempFiles
    xor ecx, ecx
    call [PostQuitMessage]
    xor eax, eax
    jmp .done

.size:
    ; Resize controls dynamically to fit the window size
    mov r12, r9
    and r12d, 0FFFFh        ; client width cx
    mov r13, r9
    shr r13d, 16
    and r13d, 0FFFFh        ; client height cy

    cmp r12d, 100
    jl .sz_status
    cmp r13d, 100
    jl .sz_status

    ; Resize Ribbon Notebook: x=4, y=2, w=cx-8, h=118
    mov rcx, [hRibbonTabs]
    test rcx, rcx
    jz .sz_form
    mov edx, 4
    mov r8d, 2
    mov r9d, r12d
    sub r9d, 8
    mov eax, 118
    sub rsp, 30h
    mov dword [rsp+20h], eax
    mov dword [rsp+28h], 1
    call [MoveWindow]
    add rsp, 30h

.sz_form:
    ; Resize Formula Edit: x=116, y=122, w=cx-126, h=24
    mov rcx, [hFormulaEdit]
    test rcx, rcx
    jz .sz_grid
    mov edx, 116
    mov r8d, 122
    mov r9d, r12d
    sub r9d, 126
    mov eax, 24
    sub rsp, 30h
    mov dword [rsp+20h], eax
    mov dword [rsp+28h], 1
    call [MoveWindow]
    add rsp, 30h

.sz_grid:
    ; Resize Grid: x=6, y=150, w=cx-12, h=cy-208
    mov rcx, [hGrid]
    test rcx, rcx
    jz .sz_tabs
    mov edx, 6
    mov r8d, 150
    mov r9d, r12d
    sub r9d, 12
    mov eax, r13d
    sub eax, 208
    cmp eax, 60
    jge .grid_h_ok
    mov eax, 60
.grid_h_ok:
    sub rsp, 30h
    mov dword [rsp+20h], eax
    mov dword [rsp+28h], 1
    call [MoveWindow]
    add rsp, 30h

.sz_tabs:
    ; Resize Sheet Tabs: x=6, y=cy-56, w=cx-12, h=26
    mov rcx, [hSheetTabs]
    test rcx, rcx
    jz .sz_status
    mov edx, 6
    mov r8d, r13d
    sub r8d, 56
    mov r9d, r12d
    sub r9d, 12
    mov eax, 26
    sub rsp, 30h
    mov dword [rsp+20h], eax
    mov dword [rsp+28h], 1
    call [MoveWindow]
    add rsp, 30h

.sz_status:
    mov rcx, [hStatusBar]
    test rcx, rcx
    jz .ret0
    mov edx, WM_SIZE
    xor r8, r8
    mov r9, [rbp+30h]
    call [SendMessageA]
    jmp .ret0

.notify:
    mov rax, r9
    test rax, rax
    jz .ret0
    mov edx, [rax+16]
    cmp edx, TCN_SELCHANGE
    je .on_tab_change

    ; Filter out any notifications from other controls (tabs, toolbars, etc.)
    mov rcx, [rax]          ; hwndFrom
    cmp rcx, [hGrid]
    jne .ret0

    cmp edx, -12 ; NM_CUSTOMDRAW
    je .on_customdraw
    cmp edx, -108 ; LVN_COLUMNCLICK
    je .on_column_click
    cmp edx, NM_CLICK
    je .on_grid_click
    cmp edx, NM_DBLCLK
    je .on_grid_dblclick
    jmp .ret0

.on_customdraw:
    mov rsi, r9             ; rsi = pLVCD (preserved in WndProc)
    mov rcx, [rsi]          ; hdr.hwndFrom
    cmp rcx, [hGrid]
    jne .ret0

    mov edx, [rsi+24]       ; dwDrawStage
    cmp edx, 1              ; CDDS_PREPAINT
    je .cd_prepaint
    cmp edx, 10001h         ; CDDS_ITEMPREPAINT
    je .cd_item_prepaint
    cmp edx, 30001h         ; CDDS_SUBITEMPREPAINT
    je .cd_subitem_prepaint
    jmp .ret0

.cd_prepaint:
    mov eax, 20h            ; CDRF_NOTIFYITEMSPEC
    jmp .done

.cd_item_prepaint:
    mov eax, 20h            ; CDRF_NOTIFYSUBITEMDRAW
    jmp .done

.cd_subitem_prepaint:
    mov ebx, [rsi+56]      ; dwItemSpec (row index: 0..N) - non-volatile in WndProc
    mov edi, [rsi+88]      ; iSubItem (col index: 0..N) - non-volatile in WndProc

    ; Clear CDIS_SELECTED so Windows never paints its default blue selection bar
    and dword [rsi+64], not 1

    ; Col 0: Row number column styling
    test edi, edi
    jnz .check_active_cell

    ; Check if ebx is within [min_row..max_row]
    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    cmp ebx, eax
    jl .col0_normal
    cmp ebx, edx
    jg .col0_normal
    ; Selected row header: accent indicator
    mov dword [rsi+84], 00D0C4B8h ; darker accent gray
    mov dword [rsi+80], 001A1008h ; dark text
    mov rcx, [rsi+32]
    mov rdx, [hFontBold]
    call [SelectObject]
    mov eax, 2              ; CDRF_NEWFONT
    jmp .done

.col0_normal:
    mov dword [rsi+84], 00EAE6E2h ; clrTextBk (Soft light gray #E2E6EA)
    mov dword [rsi+80], 0068554Ah ; clrText (Slate gray #4A5568)
    mov rcx, [rsi+32]       ; hdc
    mov rdx, [hFontSmall]
    call [SelectObject]
    mov eax, 2              ; CDRF_NEWFONT
    jmp .done

.check_active_cell:
    ; 1. Baseline style for this cell
    test ebx, 1
    jz .even_default
    mov dword [rsi+84], 00FCFAF7h ; odd row
    jmp .set_default_fg
.even_default:
    mov dword [rsi+84], 00FFFFFFh ; even row
.set_default_fg:
    mov dword [rsi+80], 002C201Ah ; dark charcoal
    mov r12, [hFontUI]            ; default font

    mov rax, [pCellStyleTable]
    test rax, rax
    jz .apply_active_cell
    cmp ebx, 10000
    jae .apply_active_cell
    cmp edi, 32
    jae .apply_active_cell

    mov edx, ebx
    shl edx, 5
    add edx, edi
    shl rdx, 4
    add rdx, rax            ; rdx = pointer to CellStyleRecord

    mov eax, [rdx+8]        ; flags (bit0=bold, bit1=has_bk, bit2=has_fg, bit3=italic)

    test eax, 2             ; bit 1: has_bk
    jz @f
    mov ecx, [rdx]
    mov [rsi+84], ecx       ; set clrTextBk
@@:
    test eax, 4             ; bit 2: has_fg
    jz @f
    mov ecx, [rdx+4]
    mov [rsi+80], ecx       ; set clrText
@@:
    ; Resolve per-cell font (size, face, bold/italic/underline)
    movzx ecx, byte [rdx+13] ; cell font size (0 = default 10)
    movzx edx, byte [rdx+14] ; cell face index (0 = default Segoe UI)
    mov r8d, eax             ; style flags
    call GetCellFont
    mov r12, rax             ; r12 = cell HFONT

.apply_active_cell:
    ; Check if (ebx, edi) is inside selection range [min_row..max_row, min_col..max_col]
    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    cmp ebx, eax
    jl .paint_cell_now
    cmp ebx, edx
    jg .paint_cell_now

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    cmp edi, eax
    jl .paint_cell_now
    cmp edi, edx
    jg .paint_cell_now

    ; --- Cell IS in selection range! ---
    cmp ebx, [cur_row]
    jne .paint_selected_non_active
    cmp edi, [cur_col]
    jne .paint_selected_non_active

    ; Active Cell: check if cell has custom background/text color
    mov rax, [pCellStyleTable]
    test rax, rax
    jz .act_default_colors
    cmp ebx, 10000
    jae .act_default_colors
    cmp edi, 32
    jae .act_default_colors
    mov edx, ebx
    shl edx, 5
    add edx, edi
    shl rdx, 4
    add rdx, rax
    test dword [rdx+8], 2    ; has_bk?
    jnz @f
    mov dword [rsi+84], 00FFE4D0h ; soft blue #D0E4FF
@@:
    test dword [rdx+8], 4    ; has_fg?
    jnz .paint_cell_now
    mov dword [rsi+80], 0069300Ah ; Navy text #0A3069
    jmp .paint_cell_now

.act_default_colors:
    mov dword [rsi+84], 00FFE4D0h
    mov dword [rsi+80], 0069300Ah
    jmp .paint_cell_now

.paint_selected_non_active:
    ; Non-active selected cell
    mov rax, [pCellStyleTable]
    test rax, rax
    jz .sel_default_colors
    cmp ebx, 10000
    jae .sel_default_colors
    cmp edi, 32
    jae .sel_default_colors
    mov edx, ebx
    shl edx, 5
    add edx, edi
    shl rdx, 4
    add rdx, rax
    test dword [rdx+8], 2    ; has_bk?
    jnz @f
    mov dword [rsi+84], 00F8E0CCh ; Excel selection light blue
@@:
    test dword [rdx+8], 4    ; has_fg?
    jnz .paint_cell_now
    mov dword [rsi+80], 00382010h ; dark slate navy text
    jmp .paint_cell_now

.sel_default_colors:
    mov dword [rsi+84], 00F8E0CCh ; Excel selection light blue (RGB 204, 224, 248)
    mov dword [rsi+80], 00382010h ; dark slate navy text

.paint_cell_now:
    mov rcx, [rsi+32]       ; hdc
    mov rdx, r12            ; font
    call [SelectObject]

    ; Check if this cell has a custom alignment, font, or fill color in pCellStyleTable
    mov rax, [pCellStyleTable]
    test rax, rax
    jz .pcn_default
    cmp ebx, 10000
    jae .pcn_default
    cmp edi, 32
    jae .pcn_default

    mov edx, ebx
    shl edx, 5
    add edx, edi
    shl rdx, 4
    add rdx, rax            ; rdx = pointer to CellStyleRecord
    movzx r15d, byte [rdx+12] ; alignment (0=default, 1=left, 2=center, 3=right)
    cmp r15d, 2
    jae .pcn_do_draw
    cmp byte [rdx+13], 0
    jne .pcn_do_draw
    cmp byte [rdx+14], 0
    jne .pcn_do_draw
    test dword [rdx+8], 2    ; has_bk?
    jnz .pcn_do_draw
    test dword [rdx+8], 4    ; has_fg?
    jnz .pcn_do_draw
    jmp .pcn_default

.pcn_do_draw:
    ; --- Draw cell with custom alignment or custom font ---
    ; Query cell text
    mov dword [lvitem_cd], 0001h ; LVIF_TEXT
    mov dword [lvitem_cd+4], ebx
    mov dword [lvitem_cd+8], edi
    lea rax, [szCellTextCd]
    mov qword [lvitem_cd+24], rax
    mov dword [lvitem_cd+32], 256
    mov byte [szCellTextCd], 0
    mov rcx, [hGrid]
    mov edx, 102Dh ; LVM_GETITEMTEXTA (1000h + 45)
    mov r8d, ebx
    lea r9, [lvitem_cd]
    call [SendMessageA]

    ; Query subitem rectangle using LVM_GETSUBITEMRECT (1038h)
    lea r9, [rcSubItemCd]
    mov dword [r9], 0          ; left = LVIR_BOUNDS
    mov dword [r9+4], edi      ; top = col
    mov dword [r9+8], 0
    mov dword [r9+12], 0
    mov rcx, [hGrid]
    mov edx, 1038h              ; LVM_GETSUBITEMRECT
    mov r8d, ebx
    call [SendMessageA]

    ; Create solid brush for background
    mov ecx, [rsi+84]
    call [CreateSolidBrush]
    mov r13, rax ; brush

    ; Fill background
    mov eax, [rcSubItemCd]
    mov [rcBgCd], eax
    mov eax, [rcSubItemCd+4]
    mov [rcBgCd+4], eax
    mov eax, [rcSubItemCd+8]
    cmp dword [gridlines_act], 1
    jne @f
    dec eax
@@:
    mov [rcBgCd+8], eax
    mov eax, [rcSubItemCd+12]
    cmp dword [gridlines_act], 1
    jne @f
    dec eax
@@:
    mov [rcBgCd+12], eax

    mov rcx, [rsi+32] ; hdc
    lea rdx, [rcBgCd]
    mov r8, r13
    call [FillRect]

    mov rcx, r13
    call [DeleteObject]

    cmp byte [szCellTextCd], 0
    je .pcn_skip_text

    ; Set text color
    mov rcx, [rsi+32]
    mov edx, [rsi+80]
    call [SetTextColor]

    ; Transparent background
    mov rcx, [rsi+32]
    mov edx, 1 ; TRANSPARENT
    call [SetBkMode]

    ; Add 4px left/right padding to rcSubItemCd
    mov eax, [rcBgCd]
    add eax, 4
    mov [rcSubItemCd], eax
    mov eax, [rcBgCd+8]
    sub eax, 4
    mov [rcSubItemCd+8], eax
    mov eax, [rcBgCd+4]
    mov [rcSubItemCd+4], eax
    mov eax, [rcBgCd+12]
    mov [rcSubItemCd+12], eax

    ; Select cell font
    mov rcx, [rsi+32]
    mov rdx, r12
    call [SelectObject]

    ; Determine DrawText flag from r15d
    mov r14d, DT_LEFT or DT_VCENTER or DT_SINGLELINE or DT_NOPREFIX
    cmp r15d, 2 ; ALIGN_CENTER
    jne @f
    mov r14d, DT_CENTER or DT_VCENTER or DT_SINGLELINE or DT_NOPREFIX
    jmp .dt_ready
@@:
    cmp r15d, 3 ; ALIGN_RIGHT
    jne .dt_ready
    mov r14d, DT_RIGHT or DT_VCENTER or DT_SINGLELINE or DT_NOPREFIX
.dt_ready:
    mov rcx, [rsi+32]
    lea rdx, [szCellTextCd]
    mov r8d, -1
    lea r9, [rcSubItemCd]
    mov dword [rsp+20h], r14d
    call [DrawTextA]

.pcn_skip_text:
    cmp ebx, [cur_row]
    jne @f
    cmp edi, [cur_col]
    jne @f
    lea rdx, [rcBgCd]
    mov rcx, [rsi+32]
    call [DrawFocusRect]
@@:
    mov eax, 4 ; CDRF_SKIPDEFAULT
    jmp .done

.pcn_default:
    mov eax, 2              ; CDRF_NEWFONT
    jmp .done

.on_column_click:
    mov rax, r9             ; pNMListView
    mov edx, [rax+28]       ; iSubItem (column index: 0 = #, 1 = A.. 26 = Z)
    cmp edx, 0
    jne .col_click_col
    call DoSelectAll
    jmp .ret0

.col_click_col:
    cmp edx, 1
    jl .ret0
    cmp edx, 26
    jg .ret0
    mov [sel_start_col], edx
    mov [sel_end_col], edx
    mov [cur_col], edx
    mov [sel_start_row], 0
    mov eax, [total_rows]
    dec eax
    mov [sel_end_row], eax
    mov [cur_row], 0
    call UpdateCellSelectionDisplay
    call SyncListViewItemState
    jmp .ret0

.on_grid_click:
    call DoGridCellHitTest
    jmp .ret0

.on_grid_dblclick:
    call DoGridCellHitTest
    call StartInPlaceCellEdit
    jmp .ret0

.keydown:
    jmp .handle_key

.handle_key:
    cmp r8d, VK_F2
    je .key_f2
    cmp r8d, VK_F3
    je .on_insert_fx
    cmp r8d, VK_DELETE
    je .key_del
    cmp r8d, VK_F5
    je .key_f5
    cmp r8d, VK_F9
    je .key_f9

    ; Forward navigation keys to Grid if received by main window
    cmp r8d, 25h ; VK_LEFT
    je .fwd_grid_key
    cmp r8d, 26h ; VK_UP
    je .fwd_grid_key
    cmp r8d, 27h ; VK_RIGHT
    je .fwd_grid_key
    cmp r8d, 28h ; VK_DOWN
    je .fwd_grid_key
    cmp r8d, 09h ; VK_TAB
    je .fwd_grid_key
    cmp r8d, 0Dh ; VK_RETURN
    je .fwd_grid_key
    cmp r8d, 21h ; VK_PRIOR (PgUp)
    je .fwd_grid_key
    cmp r8d, 22h ; VK_NEXT (PgDn)
    je .fwd_grid_key
    cmp r8d, 24h ; VK_HOME
    je .fwd_grid_key
    cmp r8d, 23h ; VK_END
    je .fwd_grid_key

    ; Check Control key
    mov ecx, VK_CONTROL
    call [GetKeyState]
    test ax, 8000h
    jz .ret0

    ; Control Shortcuts
    cmp r8d, 'C'
    je .on_copy
    cmp r8d, 'c'
    je .on_copy
    cmp r8d, 'V'
    je .on_paste
    cmp r8d, 'v'
    je .on_paste
    cmp r8d, 'X'
    je .on_cut
    cmp r8d, 'x'
    je .on_cut
    cmp r8d, 'Z'
    je .on_undo
    cmp r8d, 'z'
    je .on_undo
    cmp r8d, 'Y'
    je .on_redo
    cmp r8d, 'y'
    je .on_redo
    cmp r8d, 'S'
    je .on_save_csv
    cmp r8d, 's'
    je .on_save_csv
    cmp r8d, 'O'
    je .on_open_csv
    cmp r8d, 'o'
    je .on_open_csv
    cmp r8d, 'N'
    je .on_clear_all
    cmp r8d, 'n'
    je .on_clear_all
    cmp r8d, 'A'
    je .on_select_all
    cmp r8d, 'a'
    je .on_select_all
    cmp r8d, 'F'
    je .on_find
    cmp r8d, 'f'
    je .on_find
    cmp r8d, 'B'
    je .on_toggle_bold
    cmp r8d, 'b'
    je .on_toggle_bold
    cmp r8d, 'I'
    je .on_toggle_italic
    cmp r8d, 'i'
    je .on_toggle_italic
    cmp r8d, 'U'
    je .on_toggle_underline
    cmp r8d, 'u'
    je .on_toggle_underline
    jmp .ret0

.fwd_grid_key:
    mov rcx, [hGrid]
    test rcx, rcx
    jz .ret0
    mov edx, WM_KEYDOWN
    call [SendMessageA]
    xor eax, eax
    jmp .done

.key_f2:
    call StartInPlaceCellEdit
    jmp .ret0

.key_del:
    call DoClearCell
    jmp .ret0

.key_f5:
    call DoRefreshGrid
    jmp .ret0

.key_f9:
    call DoRecalculateAll
    jmp .ret0

.on_tab_change:
    mov rax, r9
    mov rcx, [rax]           ; hwndFrom
    cmp rcx, [hSheetTabs]
    je .handle_sheet_switch
    cmp rcx, [hRibbonTabs]
    je .handle_ribbon_tab_switch
    jmp .ret0

.handle_ribbon_tab_switch:
    mov rcx, [hRibbonTabs]
    mov edx, TCM_GETCURSEL
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]
    mov ecx, eax
    call SwitchRibbonTab
    jmp .ret0

.handle_sheet_switch:
    mov rcx, [hSheetTabs]
    mov edx, TCM_GETCURSEL
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]
    cmp eax, [sheet_count]   ; [+] Add sheet tab
    jge .on_add_sheet

    cmp eax, [cur_sheet]
    je .ret0

    mov [cur_sheet], eax
    cmp byte [is_temp_file], 1
    jne .switch_non_temp

    mov ecx, eax
    call DoLoadSheetByIndex
    jmp .ret0

.switch_non_temp:
    call LoadSheetFromBuffer
    jmp .ret0

.on_add_sheet:
    ; Save active sheet before adding new sheet
    call SaveActiveSheetToBuffer

    mov eax, [sheet_count]
    cmp eax, 15
    jge .ret0

    inc [sheet_count]
    lea rdi, [tcitem]
    xor eax, eax
    mov ecx, 48
    rep stosb
    mov dword [tcitem], TCIF_TEXT

    lea rcx, [szTemp]
    lea rdx, [new_sheet_fmt]
    mov r8d, [sheet_count]
    call [sprintf]

    lea rax, [szTemp]
    mov qword [tcitem+16], rax
    mov rcx, [hSheetTabs]
    mov edx, TCM_INSERTITEMA
    mov r8d, [sheet_count]
    dec r8d
    lea r9, [tcitem]
    call [SendMessageA]

    mov eax, [sheet_count]
    dec eax
    mov [cur_sheet], eax
    mov rcx, [hSheetTabs]
    mov edx, TCM_SETCURSEL
    mov r8d, eax
    xor r9d, r9d
    call [SendMessageA]
    call LoadSheetFromBuffer
    jmp .ret0

.command:
    mov eax, r8d
    and eax, 0FFFFh

    ; Menu Commands - File
    cmp eax, ID_MENU_NEW
    je .on_clear_all
    cmp eax, ID_MENU_OPEN_CSV
    je .on_open_csv
    cmp eax, ID_MENU_OPEN_DEMO
    je .on_load_demo
    cmp eax, ID_MENU_SAVE
    je .on_save_csv
    cmp eax, ID_MENU_SAVEAS
    je .on_saveas_csv
    cmp eax, ID_MENU_EXIT
    je .destroy

    ; Menu Commands - Edit
    cmp eax, ID_MENU_UNDO
    je .on_undo
    cmp eax, ID_MENU_REDO
    je .on_redo
    cmp eax, ID_MENU_CUT
    je .on_cut
    cmp eax, ID_MENU_COPY
    je .on_copy
    cmp eax, ID_MENU_PASTE
    je .on_paste
    cmp eax, ID_MENU_DEL_CELL
    je .on_del_cell
    cmp eax, ID_MENU_SELECT_ALL
    je .on_select_all
    cmp eax, ID_MENU_FIND
    je .on_find

    ; Menu Commands - View
    cmp eax, ID_MENU_GRIDLINES
    je .on_toggle_gridlines
    cmp eax, ID_MENU_ZOOM100
    je .on_reset_zoom
    cmp eax, ID_MENU_REFRESH
    je .on_refresh

    ; Menu Commands - Insert
    cmp eax, ID_MENU_INSERT_ROW
    je .on_insert_row
    cmp eax, ID_MENU_DELETE_ROW
    je .on_delete_row
    cmp eax, ID_MENU_INSERT_FX
    je .on_insert_fx

    ; Menu Commands - Format
    cmp eax, ID_MENU_FONT_BOLD
    je .on_toggle_bold
    cmp eax, ID_MENU_FONT_ITALIC
    je .on_toggle_italic
    cmp eax, ID_MENU_FONT_UNDER
    je .on_toggle_underline
    cmp eax, ID_MENU_FONT_STRIKE
    je .on_toggle_strike
    cmp eax, ID_MENU_FONT_INC
    je .on_inc_font
    cmp eax, ID_MENU_FONT_DEC
    je .on_dec_font
    cmp eax, ID_MENU_ALIGN_LEFT
    je .on_align_left
    cmp eax, ID_MENU_ALIGN_CENT
    je .on_align_center
    cmp eax, ID_MENU_ALIGN_RIGH
    je .on_align_right
    cmp eax, ID_MENU_FMT_DOLLAR
    je .on_fmt_dollar
    cmp eax, ID_MENU_FMT_RUBLE
    je .on_fmt_ruble
    cmp eax, ID_MENU_FMT_PERCENT
    je .on_fmt_percent
    cmp eax, ID_MENU_FMT_COMMA
    je .on_fmt_comma

    ; Menu Commands - Formulas
    cmp eax, ID_MENU_AUTOSUM
    je .on_autosum
    cmp eax, ID_MENU_AVERAGE
    je .on_simd_avg
    cmp eax, ID_MENU_MIN
    je .on_simd_min
    cmp eax, ID_MENU_MAX
    je .on_simd_max
    cmp eax, ID_MENU_RECALC
    je .on_recalc
    cmp eax, ID_MENU_PMT
    je .on_pmt

    ; Menu Commands - Data
    cmp eax, ID_MENU_SORT_ASC
    je .on_sort_asc
    cmp eax, ID_MENU_SORT_DESC
    je .on_sort_desc
    cmp eax, ID_MENU_BENCHMARK
    je .on_benchmark

    ; Menu Commands - Help
    cmp eax, ID_MENU_FORM_REF
    je .on_formula_ref
    cmp eax, ID_MENU_SIMD_INFO
    je .on_simd_info
    cmp eax, ID_MENU_ABOUT
    je .on_about

    ; Ribbon Buttons
    cmp eax, ID_BTN_PASTE
    je .on_paste
    cmp eax, ID_BTN_CUT
    je .on_cut
    cmp eax, ID_BTN_COPY
    je .on_copy
    cmp eax, ID_BTN_FORMAT
    je .on_format_painter

    cmp eax, ID_BTN_INC_FONT
    je .on_inc_font
    cmp eax, ID_BTN_DEC_FONT
    je .on_dec_font
    cmp eax, ID_BTN_BOLD
    je .on_toggle_bold
    cmp eax, ID_BTN_ITALIC
    je .on_toggle_italic
    cmp eax, ID_BTN_UNDERLINE
    je .on_toggle_underline
    cmp eax, ID_BTN_STRIKE
    je .on_toggle_strike
    cmp eax, ID_BTN_BORDER
    je .on_toggle_gridlines
    cmp eax, ID_BTN_FILL_COLOR
    je .on_toggle_fill
    cmp eax, ID_BTN_TEXT_COLOR
    je .on_toggle_text_color

    ; Ribbon Insert Buttons
    cmp eax, ID_BTN_INS_TABLE
    je .on_ins_table
    cmp eax, ID_BTN_INS_PIVOT
    je .on_ins_pivot
    cmp eax, ID_BTN_INS_PICTURE
    je .on_ins_pic
    cmp eax, ID_BTN_INS_SHAPES
    je .on_ins_shp
    cmp eax, ID_BTN_INS_ICONS
    je .on_ins_ico
    cmp eax, ID_BTN_INS_COL_CHART
    je .on_ins_chart
    cmp eax, ID_BTN_INS_LINE_CHART
    je .on_ins_chart
    cmp eax, ID_BTN_INS_PIE_CHART
    je .on_ins_chart
    cmp eax, ID_BTN_INS_BAR_CHART
    je .on_ins_chart
    cmp eax, ID_BTN_INS_LINK
    je .on_ins_link
    cmp eax, ID_BTN_INS_TEXTBOX
    je .on_ins_textbox
    cmp eax, ID_BTN_INS_SYMBOL
    je .on_ins_symbol

    ; Ribbon Data Buttons
    cmp eax, ID_BTN_DATA_CSV
    je .on_open_csv
    cmp eax, ID_BTN_DATA_REFRESH
    je .on_recalc
    cmp eax, ID_BTN_DATA_SORT_ASC
    je .on_sort_asc
    cmp eax, ID_BTN_DATA_SORT_DESC
    je .on_sort_desc
    cmp eax, ID_BTN_DATA_FILTER
    je .on_data_filter
    cmp eax, ID_BTN_DATA_CLEAR
    je .on_data_clear
    cmp eax, ID_BTN_DATA_SPLIT
    je .on_data_split
    cmp eax, ID_BTN_DATA_REMDUP
    je .on_data_remdup
    cmp eax, ID_BTN_DATA_VALID
    je .on_data_valid
    cmp eax, ID_BTN_DATA_BENCH
    je .on_simd_info
    cmp eax, ID_BTN_DATA_WHATIF
    je .on_data_whatif

    ; Ribbon Formulas Buttons
    cmp eax, ID_BTN_FORM_FX
    je .on_insert_fx
    cmp eax, ID_BTN_FORM_AUTOSUM
    je .on_autosum
    cmp eax, ID_BTN_FORM_FIN
    je .on_form_fin
    cmp eax, ID_BTN_FORM_LOGIC
    je .on_form_logic
    cmp eax, ID_BTN_FORM_TEXT
    je .on_form_text
    cmp eax, ID_BTN_FORM_DATE
    je .on_form_date
    cmp eax, ID_BTN_FORM_LOOKUP
    je .on_form_lookup
    cmp eax, ID_BTN_FORM_MATH
    je .on_form_math
    cmp eax, ID_BTN_FORM_MORE
    je .on_insert_fx
    cmp eax, ID_BTN_FORM_NAME_MGR
    je .on_name_mgr
    cmp eax, ID_BTN_FORM_DEF_NAME
    je .on_def_name
    cmp eax, ID_BTN_FORM_CALC_NOW
    je .on_recalc
    cmp eax, ID_BTN_FORM_EVAL
    je .on_eval_formula

    ; Ribbon View Buttons
    cmp eax, ID_BTN_VIEW_NORMAL
    je .on_view_normal
    cmp eax, ID_BTN_VIEW_PAGELAY
    je .on_view_pagelay
    cmp eax, ID_BTN_VIEW_GRIDLINES
    je .on_toggle_gridlines
    cmp eax, ID_BTN_VIEW_FORMBAR
    je .on_view_toggle_formbar
    cmp eax, ID_BTN_VIEW_HEADINGS
    je .on_view_headings
    cmp eax, ID_BTN_VIEW_ZOOM100
    je .on_zoom_100
    cmp eax, ID_BTN_VIEW_ZOOM_IN
    je .on_inc_font
    cmp eax, ID_BTN_VIEW_ZOOM_OUT
    je .on_dec_font
    cmp eax, ID_BTN_VIEW_FREEZE
    je .on_view_freeze
    cmp eax, ID_BTN_VIEW_NEW_WND
    je .on_view_new_wnd

    cmp eax, ID_BTN_UNDO
    je .on_undo
    cmp eax, ID_BTN_REDO
    je .on_redo

    cmp eax, ID_BTN_TOP
    je .on_align_center
    cmp eax, ID_BTN_MID
    je .on_align_center
    cmp eax, ID_BTN_BOT
    je .on_align_center

    cmp eax, ID_BTN_LEFT
    je .on_align_left
    cmp eax, ID_BTN_CENTER
    je .on_align_center
    cmp eax, ID_BTN_RIGHT
    je .on_align_right
    cmp eax, ID_BTN_WRAP
    je .on_wrap_text
    cmp eax, ID_BTN_MERGE
    je .on_merge_center

    cmp eax, ID_BTN_DOLLAR
    je .on_fmt_dollar
    cmp eax, ID_BTN_RUBLE
    je .on_fmt_ruble
    cmp eax, ID_BTN_PERCENT
    je .on_fmt_percent
    cmp eax, ID_BTN_COMMA
    je .on_fmt_comma
    cmp eax, ID_BTN_INC_DECIMAL
    je .on_inc_decimal
    cmp eax, ID_BTN_DEC_DECIMAL
    je .on_dec_decimal

    cmp eax, ID_BTN_INSERT_CELL
    je .on_insert_row
    cmp eax, ID_BTN_DELETE_CELL
    je .on_delete_row

    cmp eax, ID_BTN_AUTOSUM
    je .on_autosum
    cmp eax, ID_BTN_SORT_FILTER
    je .on_sort_asc
    cmp eax, ID_BTN_FIND_SELECT
    je .on_find
    cmp eax, ID_BTN_COND_FMT
    je .on_cond_format
    cmp eax, ID_BTN_FX
    je .on_insert_fx

    ; Combobox Selection Changes
    cmp eax, ID_COMBO_FONT
    je .on_combo_font
    cmp eax, ID_COMBO_SIZE
    je .on_combo_size
    cmp eax, ID_COMBO_NUM_FMT
    je .on_combo_num_fmt

    jmp .ret0

; -----------------------------------------------------------------------------
; Event Handlers
; -----------------------------------------------------------------------------
.on_clear_all:
    mov rcx, [hGrid]
    mov edx, LVM_DELETEALLITEMS
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]
    mov [total_rows], 0
    call UpdateStatusStats
    jmp .ret0

.on_open_csv:
    call DoFileOpen
    jmp .ret0

.on_load_demo:
    call LoadDemoData
    jmp .ret0

.on_save_csv:
    call DoFileSave
    jmp .ret0

.on_saveas_csv:
    call DoFileSaveAs
    jmp .ret0

.on_undo:
    call DoUndo
    jmp .ret0

.on_redo:
    call DoRedo
    jmp .ret0

.on_del_cell:
    call DoClearCell
    jmp .ret0

.on_select_all:
    call DoSelectAll
    jmp .ret0

.on_find:
    call DoFindText
    jmp .ret0

.on_reset_zoom:
    call DoResetZoom
    jmp .ret0

.on_refresh:
    call DoRefreshGrid
    jmp .ret0

.on_insert_row:
    call DoInsertNewRow
    jmp .ret0

.on_delete_row:
    call DoDeleteActiveRow
    jmp .ret0

.on_insert_fx:
    mov dword [wiz_init_cat], 0
    call DoShowFxWizard
    jmp .ret0

.on_toggle_gridlines:
    mov rcx, [hGrid]
    mov edx, LVM_GETEXTENDEDLISTVIEWSTYLE
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]
    xor eax, LVS_EX_GRIDLINES
    mov r9d, eax
    mov rcx, [hGrid]
    mov edx, LVM_SETEXTENDEDLISTVIEWSTYLE
    xor r8d, r8d
    call [SendMessageA]
    jmp .ret0

.on_toggle_bold:
    call ToggleCellBold
    jmp .ret0

.on_toggle_italic:
    call ToggleCellItalic
    jmp .ret0

.on_toggle_underline:
    call ToggleCellUnderline
    jmp .ret0

.on_toggle_strike:
    call ToggleCellStrike
    jmp .ret0

.on_toggle_fill:
    call ToggleCellFill
    jmp .ret0

.on_toggle_text_color:
    call ToggleCellTextColor
    jmp .ret0

; Ribbon Action Handlers
.on_ins_table:
    lea rcx, [szTablePrompt]
    call ShowInfoStatus
    jmp .ret0

.on_ins_pivot:
    lea rcx, [szPivotPrompt]
    call ShowInfoStatus
    jmp .ret0

.on_ins_pic:
.on_ins_shp:
.on_ins_ico:
.on_ins_chart:
    lea rcx, [szChartPrompt]
    call ShowInfoStatus
    jmp .ret0

.on_ins_link:
.on_ins_textbox:
    lea rcx, [szNameMgrPrompt]
    call ShowInfoStatus
    jmp .ret0

.on_ins_symbol:
    lea rcx, [symbol_sigma_str]
    call InsertFormulaSnippet
    jmp .ret0

.on_form_fin:
    mov dword [wiz_init_cat], 7
    call DoShowFxWizard
    jmp .ret0

.on_form_logic:
    mov dword [wiz_init_cat], 4
    call DoShowFxWizard
    jmp .ret0

.on_form_text:
    mov dword [wiz_init_cat], 5
    call DoShowFxWizard
    jmp .ret0

.on_form_date:
    mov dword [wiz_init_cat], 6
    call DoShowFxWizard
    jmp .ret0

.on_form_lookup:
    mov dword [wiz_init_cat], 3
    call DoShowFxWizard
    jmp .ret0

.on_form_math:
    mov dword [wiz_init_cat], 1
    call DoShowFxWizard
    jmp .ret0

.on_name_mgr:
    lea rcx, [szNameMgrPrompt]
    call ShowInfoStatus
    jmp .ret0

.on_def_name:
    lea rcx, [szDefNamePrompt]
    call ShowInfoStatus
    jmp .ret0

.on_data_filter:
    lea rcx, [szFilterPrompt]
    call ShowInfoStatus
    jmp .ret0

.on_data_clear:
    lea rcx, [szClearFilterPrompt]
    call ShowInfoStatus
    jmp .ret0

.on_data_split:
    lea rcx, [szSplitPrompt]
    call ShowInfoStatus
    jmp .ret0

.on_data_remdup:
    lea rcx, [szClearFilterPrompt]
    call ShowInfoStatus
    jmp .ret0

.on_data_valid:
    lea rcx, [szFilterPrompt]
    call ShowInfoStatus
    jmp .ret0

.on_data_whatif:
    lea rcx, [szPivotPrompt]
    call ShowInfoStatus
    jmp .ret0

.on_view_normal:
.on_view_pagelay:
.on_view_toggle_formbar:
.on_view_headings:
.on_view_new_wnd:
    lea rcx, [szNormalViewPrompt]
    call ShowInfoStatus
    jmp .ret0

.on_zoom_100:
    mov dword [font_size_pt], 10
    call ApplyNewGridFont
    call SyncComboSize
    jmp .ret0

.on_view_freeze:
    lea rcx, [szTablePrompt]
    call ShowInfoStatus
    jmp .ret0

.on_inc_font:
    call DoIncSelectionFontSize
    jmp .ret0

.on_dec_font:
    call DoDecSelectionFontSize
    jmp .ret0

.on_align_left:
    mov r8d, ALIGN_LEFT
    call ApplySelectionAlignment
    jmp .ret0

.on_align_center:
    mov r8d, ALIGN_CENTER
    call ApplySelectionAlignment
    jmp .ret0

.on_align_right:
    mov r8d, ALIGN_RIGHT
    call ApplySelectionAlignment
    jmp .ret0

.on_fmt_dollar:
    call ApplyDollarFormat
    jmp .ret0

.on_fmt_ruble:
    call ApplyRubleFormat
    jmp .ret0

.on_fmt_percent:
    call ApplyPercentFormat
    jmp .ret0

.on_fmt_comma:
    call ApplyCommaFormat
    jmp .ret0

.on_autosum:
    call DoSimdAutoSum
    jmp .ret0

.on_simd_avg:
    call DoSimdAverage
    jmp .ret0

.on_simd_min:
    call DoSimdMin
    jmp .ret0

.on_simd_max:
    call DoSimdMax
    jmp .ret0

.on_recalc:
    call DoRecalculateAll
    jmp .ret0

.on_pmt:
    call DoPmtDialog
    jmp .ret0

.on_sort_asc:
    call DoSortAscending
    jmp .ret0

.on_sort_desc:
    call DoSortDescending
    jmp .ret0

.on_benchmark:
    call DoRunBenchmark
    jmp .ret0

.on_formula_ref:
    call DoShowFormulaRef
    jmp .ret0

.on_simd_info:
    call DoShowSimdInfo
    jmp .ret0

.on_about:
    mov rcx, [hWnd]
    lea rdx, [about_msg]
    lea r8, [about_title]
    xor r9d, r9d
    call [MessageBoxA]
    jmp .ret0

.on_copy:
    call DoClipboardCopy
    jmp .ret0

.on_cut:
    call DoClipboardCut
    jmp .ret0

.on_paste:
    call DoClipboardPaste
    jmp .ret0

.on_format_painter:
    call DoFormatPainter
    jmp .ret0

.on_wrap_text:
    call DoWrapText
    jmp .ret0

.on_merge_center:
    call DoMergeCenter
    jmp .ret0

.on_inc_decimal:
    call DoIncDecimals
    jmp .ret0

.on_dec_decimal:
    call DoDecDecimals
    jmp .ret0

.on_cond_format:
    call DoConditionalFormat
    jmp .ret0

.on_eval_formula:
    call DoEvalFormula
    jmp .ret0

.on_combo_font:
    mov rcx, [hComboFont]
    mov edx, CB_GETCURSEL
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]
    cmp eax, 0
    jl .ret0
    ; 0 -> faceIdx 1 (Segoe UI)
    ; 1 -> faceIdx 2 (Calibri)
    ; 2 -> faceIdx 3 (Arial)
    ; 3 -> faceIdx 4 (Consolas)
    inc eax
    mov r8d, eax
    call ApplySelectionFontFace
    jmp .ret0

.on_combo_size:
    mov rcx, [hComboSize]
    mov edx, CB_GETCURSEL
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]
    cmp eax, 0
    jl .ret0
    cmp eax, 6
    jg .ret0
    lea rdx, [font_size_table]
    mov r8d, dword [rdx + rax*4]
    call ApplySelectionFontSize
    jmp .ret0

.on_combo_num_fmt:
    mov rcx, [hComboNumFmt]
    mov edx, CB_GETCURSEL
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]
    cmp eax, 1
    je .on_fmt_comma
    cmp eax, 2
    je .on_fmt_dollar
    cmp eax, 3
    je .on_fmt_percent
    cmp eax, 4
    je .on_fmt_ruble
    jmp .ret0

.ret0:
    xor eax, eax

.done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    mov r14, [rbp-48]
    mov r15, [rbp-56]
    leave
    ret

; =============================================================================
; Subsystem Modules (Lexical Inclusions in Code Section)
; =============================================================================
include 'menu.inc'
include 'ribbon.inc'
include 'grid.inc'
include 'formulas.inc'
include 'file_io.inc'
include 'dialogs.inc'

include 'import.inc'
include 'rsrc.inc'

; =============================================================================
; Excel Viewer Pro — Native 64-bit Assembly Chart & Graph Visualization Engine
; Provides Excel-grade interactive visual charts:
; - Clustered Column Chart (Гистограмма)
; - Line Chart with Vertex Markers (График)
; - Pie Chart with Radial Slices & Percentages (Круговая диаграмма)
; - Horizontal Bar Chart (Линейчатая)
; - Area Chart (С областями)
; - Themes (Excel Classic, Modern Office, Vibrant, Warm Sunset)
; - Double-buffered GDI rendering (flicker-free)
; - Export & Copy to Clipboard (CF_BITMAP)
; =============================================================================

ID_CHART_BTN_COL    = 3001
ID_CHART_BTN_LINE   = 3002
ID_CHART_BTN_PIE    = 3003
ID_CHART_BTN_BAR    = 3004
ID_CHART_BTN_AREA   = 3005
ID_CHART_BTN_THEME  = 3006
ID_CHART_BTN_COPY   = 3007
ID_CHART_BTN_CLOSE  = 3008

section '.code' code readable executable

; =============================================================================
; ShowChartDialog
; Entry point: ECX = Initial Chart Type (1=Col, 2=Line, 3=Pie, 4=Bar, 5=Area)
; =============================================================================
ShowChartDialog:
    push rbp
    mov rbp, rsp
    sub rsp, 0C0h
    cld
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13
    mov [rbp-30h], r14
    mov [rbp-38h], r15

    ; Set initial chart type (clamp 1..5)
    cmp ecx, 1
    jl @f
    cmp ecx, 5
    jg @f
    mov [cur_chart_type], ecx
    jmp .chart_type_ok
@@:
    mov dword [cur_chart_type], 1
.chart_type_ok:

    ; 1. Extract data from current selection / spreadsheet
    call ExtractChartDataFromSheet

    ; 2. If chart window already exists, bring it to front and invalidate
    mov rax, [hChartWnd]
    test rax, rax
    jz .create_new_wnd

    mov rcx, rax
    xor edx, edx
    xor r8d, r8d
    mov r9d, 1
    call [InvalidateRect]

    mov rcx, [hChartWnd]
    call [SetFocus]
    jmp .scd_done

.create_new_wnd:
    ; Register window class (zero out 80 bytes first)
    lea rdi, [chart_wc]
    xor eax, eax
    mov ecx, 80
    rep stosb

    mov dword [chart_wc], 80
    mov dword [chart_wc+4], CS_HREDRAW or CS_VREDRAW
    lea rax, [ChartWndProc]
    mov qword [chart_wc+8], rax
    mov rax, [hInst]
    mov qword [chart_wc+24], rax
    mov rax, [hAppIcon]
    mov qword [chart_wc+32], rax
    mov qword [chart_wc+72], rax
    xor ecx, ecx
    mov edx, 32512 ; IDC_ARROW
    call [LoadCursorA]
    mov qword [chart_wc+40], rax
    mov qword [chart_wc+48], COLOR_WINDOW + 1
    lea rax, [chart_class_name]
    mov qword [chart_wc+64], rax

    lea rcx, [chart_wc]
    call [RegisterClassExA]

    ; Window Title: format with current chart type name
    mov eax, [cur_chart_type]
    dec eax
    lea rcx, [chart_type_names]
    mov r8, [rcx + rax*8]
    lea rcx, [chart_val_buf]
    lea rdx, [chart_title_fmt]
    call [sprintf]

    ; Create top-level window (860 x 620, centered)
    xor ecx, ecx
    lea rdx, [chart_class_name]
    lea r8, [chart_val_buf]
    mov r9d, WS_OVERLAPPEDWINDOW or WS_VISIBLE
    mov qword [rsp+20h], 120
    mov qword [rsp+28h], 80
    mov qword [rsp+30h], 860
    mov qword [rsp+38h], 620
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], 0
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov [hChartWnd], rax

.scd_done:
    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    mov r14, [rbp-30h]
    mov r15, [rbp-38h]
    leave
    ret

; =============================================================================
; ExtractChartDataFromSheet
; Reads selection coordinates and populates chart categories, series, and values.
; If range has no numbers or is single empty cell, populates rich business sample data.
; =============================================================================
ExtractChartDataFromSheet:
    push rbp
    mov rbp, rsp
    sub rsp, 0C0h
    cld
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13
    mov [rbp-30h], r14
    mov [rbp-38h], r15

    ; Check selection coordinates
    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r12d, eax       ; r1 = min_row
    mov r13d, edx       ; r2 = max_row

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov ebx, eax        ; c1 = min_col
    mov esi, edx        ; c2 = max_col

    ; Check range dimensions
    mov eax, r13d
    sub eax, r12d
    inc eax
    mov edx, esi
    sub edx, ebx
    inc edx

    cmp eax, 1
    jne .parse_range
    cmp edx, 1
    jne .parse_range

    ; Single cell selected: fallback to rich business demonstration data
    jmp .fill_sample_data

.parse_range:
    ; Limit to max 16 categories and max 6 series
    mov ecx, eax
    cmp ecx, 16
    jle @f
    mov ecx, 16
@@:
    mov [chart_num_cats], ecx

    mov ecx, edx
    cmp ecx, 6
    jle @f
    mov ecx, 6
@@:
    mov [chart_num_series], ecx

    ; Parse cells from ListView
    xor r14d, r14d      ; r14d = cat index (row offset)
.cat_loop:
    cmp r14d, [chart_num_cats]
    jge .done_parse_data

    ; Set default category name: "Cat N"
    lea rdi, [chart_cat_names]
    mov eax, r14d
    shl eax, 6
    add rdi, rax
    lea rcx, [rdi]
    lea rdx, [chart_fmt_int]
    lea r8d, [r14d+1]
    cvtsi2sd xmm2, r8d
    movq r8, xmm2
    call [sprintf]

    xor r15d, r15d      ; r15d = series index (col offset)
.series_loop:
    cmp r15d, [chart_num_series]
    jge .next_cat

    ; Read cell text at (r12 + r14, ebx + r15)
    mov eax, r12d
    add eax, r14d
    mov dword [lvitem+4], eax
    mov eax, ebx
    add eax, r15d
    mov dword [lvitem+8], eax
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 256
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, dword [lvitem+4]
    lea r9, [lvitem]
    call [SendMessageA]

    ; Parse double
    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]

    ; Store in chart_data[r14 * 8 + r15]
    mov eax, r14d
    shl eax, 3
    add eax, r15d
    movsd [chart_data + rax*8], xmm0

    inc r15d
    jmp .series_loop

.next_cat:
    inc r14d
    jmp .cat_loop

.done_parse_data:
    ; Set series names: "Series 1", "Series 2", ...
    xor ecx, ecx
.set_s_names:
    cmp ecx, [chart_num_series]
    jge .calc_bounds
    lea rdi, [chart_series_names]
    mov eax, ecx
    shl eax, 6
    add rdi, rax
    push rcx
    lea rcx, [rdi]
    lea rdx, [szSampleS1]
    call CopyStr64
    pop rcx
    inc ecx
    jmp .set_s_names

.fill_sample_data:
    ; Fallback rich business demonstration data: 4 categories, 3 series
    mov dword [chart_num_cats], 4
    mov dword [chart_num_series], 3

    ; Categories: Q1, Q2, Q3, Q4
    lea rdi, [chart_cat_names + 0*64]
    lea rsi, [szSampleQ1]
    call CopyStr64
    lea rdi, [chart_cat_names + 1*64]
    lea rsi, [szSampleQ2]
    call CopyStr64
    lea rdi, [chart_cat_names + 2*64]
    lea rsi, [szSampleQ3]
    call CopyStr64
    lea rdi, [chart_cat_names + 3*64]
    lea rsi, [szSampleQ4]
    call CopyStr64

    ; Series: Revenue, Expenses, Profit
    lea rdi, [chart_series_names + 0*64]
    lea rsi, [szSampleS1]
    call CopyStr64
    lea rdi, [chart_series_names + 1*64]
    lea rsi, [szSampleS2]
    call CopyStr64
    lea rdi, [chart_series_names + 2*64]
    lea rsi, [szSampleS3]
    call CopyStr64

    ; Data values:
    ; Q1: 150, 90, 60
    mov rax, 4062C00000000000h ; 150.0
    mov [chart_data + (0*8+0)*8], rax
    mov rax, 4056800000000000h ; 90.0
    mov [chart_data + (0*8+1)*8], rax
    mov rax, 404E000000000000h ; 60.0
    mov [chart_data + (0*8+2)*8], rax

    ; Q2: 240, 140, 100
    mov rax, 406E000000000000h ; 240.0
    mov [chart_data + (1*8+0)*8], rax
    mov rax, 4061800000000000h ; 140.0
    mov [chart_data + (1*8+1)*8], rax
    mov rax, 4059000000000000h ; 100.0
    mov [chart_data + (1*8+2)*8], rax

    ; Q3: 310, 190, 120
    mov rax, 4073600000000000h ; 310.0
    mov [chart_data + (2*8+0)*8], rax
    mov rax, 4067C00000000000h ; 190.0
    mov [chart_data + (2*8+1)*8], rax
    mov rax, 405E000000000000h ; 120.0
    mov [chart_data + (2*8+2)*8], rax

    ; Q4: 420, 230, 190
    mov rax, 407A400000000000h ; 420.0
    mov [chart_data + (3*8+0)*8], rax
    mov rax, 406CC00000000000h ; 230.0
    mov [chart_data + (3*8+1)*8], rax
    mov rax, 4067C00000000000h ; 190.0
    mov [chart_data + (3*8+2)*8], rax

.calc_bounds:
    ; Calculate min_val and max_val
    xorpd xmm0, xmm0    ; min_val = 0.0
    movsd [chart_min_val], xmm0
    mov rax, 4024000000000000h ; 10.0 default
    mov [chart_max_val], rax

    xor ecx, ecx
.bound_c:
    cmp ecx, [chart_num_cats]
    jge .bounds_done
    xor edx, edx
.bound_s:
    cmp edx, [chart_num_series]
    jge .next_bc
    mov eax, ecx
    shl eax, 3
    add eax, edx
    movsd xmm1, [chart_data + rax*8]
    movsd xmm2, [chart_max_val]
    comisd xmm1, xmm2
    jbe @f
    movsd [chart_max_val], xmm1
@@:
    inc edx
    jmp .bound_s
.next_bc:
    inc ecx
    jmp .bound_c

.bounds_done:
    ; Add 15% headroom to max_val
    movsd xmm0, [chart_max_val]
    mov rax, 3FF2666666666666h ; 1.15
    movq xmm1, rax
    mulsd xmm0, xmm1
    movsd [chart_max_val], xmm0

    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    mov r14, [rbp-30h]
    mov r15, [rbp-38h]
    leave
    ret

CopyStr64:
    cld
    ; Copies asciiz from RSI to RDI (up to 63 chars)
    push rdi
    push rsi
@@:
    lodsb
    stosb
    test al, al
    jnz @b
    pop rsi
    pop rdi
    ret

; =============================================================================
; ChartWndProc
; Window procedure for the Chart Viewer Top-Level Window
; =============================================================================
ChartWndProc:
    push rbp
    mov rbp, rsp
    sub rsp, 0C0h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13
    mov [rbp-30h], r14
    mov [rbp-38h], r15

    cmp edx, WM_CREATE
    je .on_create
    cmp edx, WM_COMMAND
    je .on_command
    cmp edx, WM_PAINT
    je .on_paint
    cmp edx, WM_SIZE
    je .on_size
    cmp edx, WM_DESTROY
    je .on_destroy
    cmp edx, WM_KEYDOWN
    je .on_keydown

    call [DefWindowProcA]
    jmp .proc_ret

.on_create:
    mov r12, rcx        ; r12 = hWndChart

    ; Create Toolbar Buttons (y=6, h=28)
    ; Button 1: Column (x=10, w=76)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [szBtnColTxt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_PUSHBUTTON
    mov dword [rsp+20h], 10
    mov dword [rsp+28h], 6
    mov dword [rsp+30h], 76
    mov dword [rsp+38h], 28
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], ID_CHART_BTN_COL
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]

    ; Button 2: Line (x=90, w=68)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [szBtnLineTxt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_PUSHBUTTON
    mov dword [rsp+20h], 90
    mov dword [rsp+28h], 6
    mov dword [rsp+30h], 68
    mov dword [rsp+38h], 28
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], ID_CHART_BTN_LINE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]

    ; Button 3: Pie (x=162, w=64)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [szBtnPieTxt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_PUSHBUTTON
    mov dword [rsp+20h], 162
    mov dword [rsp+28h], 6
    mov dword [rsp+30h], 64
    mov dword [rsp+38h], 28
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], ID_CHART_BTN_PIE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]

    ; Button 4: Bar (x=230, w=64)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [szBtnBarTxt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_PUSHBUTTON
    mov dword [rsp+20h], 230
    mov dword [rsp+28h], 6
    mov dword [rsp+30h], 64
    mov dword [rsp+38h], 28
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], ID_CHART_BTN_BAR
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]

    ; Button 5: Area (x=298, w=66)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [szBtnAreaTxt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_PUSHBUTTON
    mov dword [rsp+20h], 298
    mov dword [rsp+28h], 6
    mov dword [rsp+30h], 66
    mov dword [rsp+38h], 28
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], ID_CHART_BTN_AREA
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]

    ; Button 6: Theme (x=380, w=76)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [szBtnThemeTxt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_PUSHBUTTON
    mov dword [rsp+20h], 380
    mov dword [rsp+28h], 6
    mov dword [rsp+30h], 76
    mov dword [rsp+38h], 28
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], ID_CHART_BTN_THEME
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]

    ; Button 7: Copy Image (x=462, w=100)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [szBtnCopyTxt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_PUSHBUTTON
    mov dword [rsp+20h], 462
    mov dword [rsp+28h], 6
    mov dword [rsp+30h], 100
    mov dword [rsp+38h], 28
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], ID_CHART_BTN_COPY
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]

    ; Button 8: Close (x=568, w=68)
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [szBtnCloseTxt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_PUSHBUTTON
    mov dword [rsp+20h], 568
    mov dword [rsp+28h], 6
    mov dword [rsp+30h], 68
    mov dword [rsp+38h], 28
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], ID_CHART_BTN_CLOSE
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]

    xor eax, eax
    jmp .proc_ret

.on_command:
    movzx eax, r8w      ; button ID
    cmp eax, ID_CHART_BTN_COL
    je .cmd_col
    cmp eax, ID_CHART_BTN_LINE
    je .cmd_line
    cmp eax, ID_CHART_BTN_PIE
    je .cmd_pie
    cmp eax, ID_CHART_BTN_BAR
    je .cmd_bar
    cmp eax, ID_CHART_BTN_AREA
    je .cmd_area
    cmp eax, ID_CHART_BTN_THEME
    je .cmd_theme
    cmp eax, ID_CHART_BTN_COPY
    je .cmd_copy
    cmp eax, ID_CHART_BTN_CLOSE
    je .cmd_close
    jmp .proc_ret0

.cmd_col:
    mov dword [cur_chart_type], 1
    jmp .cmd_refresh
.cmd_line:
    mov dword [cur_chart_type], 2
    jmp .cmd_refresh
.cmd_pie:
    mov dword [cur_chart_type], 3
    jmp .cmd_refresh
.cmd_bar:
    mov dword [cur_chart_type], 4
    jmp .cmd_refresh
.cmd_area:
    mov dword [cur_chart_type], 5
    jmp .cmd_refresh
.cmd_theme:
    inc dword [cur_chart_theme]
    and dword [cur_chart_theme], 3
    jmp .cmd_refresh
.cmd_copy:
    mov rcx, [hChartWnd]
    call CopyChartImageToClipboard
    jmp .proc_ret0
.cmd_close:
    mov rcx, [hChartWnd]
    call [DestroyWindow]
    jmp .proc_ret0

.cmd_refresh:
    ; Update title
    mov eax, [cur_chart_type]
    dec eax
    lea rcx, [chart_type_names]
    mov r8, [rcx + rax*8]
    lea rcx, [chart_val_buf]
    lea rdx, [chart_title_fmt]
    call [sprintf]

    mov rcx, [hChartWnd]
    lea rdx, [chart_val_buf]
    call [SetWindowTextA]

    mov rcx, [hChartWnd]
    xor edx, edx
    xor r8d, r8d
    mov r9d, 1
    call [InvalidateRect]
    jmp .proc_ret0

.on_size:
    mov rcx, [hChartWnd]
    xor edx, edx
    xor r8d, r8d
    mov r9d, 1
    call [InvalidateRect]
    jmp .proc_ret0

.on_keydown:
    cmp r8d, 1Bh        ; VK_ESCAPE
    je .cmd_close
    cmp r8d, '1'
    je .cmd_col
    cmp r8d, '2'
    je .cmd_line
    cmp r8d, '3'
    je .cmd_pie
    cmp r8d, '4'
    je .cmd_bar
    cmp r8d, '5'
    je .cmd_area
    jmp .proc_ret0

.on_destroy:
    mov qword [hChartWnd], 0
    jmp .proc_ret0

.on_paint:
    ; Double-buffered painting using memory DC and BitBlt
    push rbp
    mov rbp, rsp
    sub rsp, 0C0h

    mov rcx, [hChartWnd]
    lea rdx, [chart_ps]
    call [BeginPaint]
    mov r12, rax        ; r12 = hdc

    ; Get client rect
    mov rcx, [hChartWnd]
    lea rdx, [rbp-40h]
    call [GetClientRect]

    mov r13d, [rbp-38h] ; width = rc.right
    mov r14d, [rbp-34h] ; height = rc.bottom

    ; Create compatible DC and Bitmap
    mov rcx, r12
    call [CreateCompatibleDC]
    mov r15, rax        ; r15 = memDC

    mov rcx, r12
    mov edx, r13d
    mov r8d, r14d
    call [CreateCompatibleBitmap]
    mov [rbp-48h], rax  ; hBitmap

    mov rcx, r15
    mov rdx, [rbp-48h]
    call [SelectObject]
    mov [rbp-50h], rax  ; oldBmp

    ; 1. Draw crisp white background
    mov rcx, r15
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, r15
    mov rdx, [hPenNull]
    call [SelectObject]
    mov rcx, r15
    xor edx, edx
    xor r8d, r8d
    mov r9d, r13d
    mov eax, r14d
    mov [rsp+20h], eax
    call [Rectangle]

    ; 2. Draw toolbar divider line at y = 40
    mov rcx, r15
    mov rdx, [hPenBorder]
    call [SelectObject]
    mov rcx, r15
    xor edx, edx
    mov r8d, 40
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r15
    mov edx, r13d
    mov r8d, 40
    call [LineTo]

    ; 3. Render Chart Elements onto memDC
    mov rcx, r15
    mov edx, r13d
    mov r8d, r14d
    call RenderChartToDC

    ; 4. BitBlt back buffer to screen DC
    mov rcx, r12
    xor edx, edx
    xor r8d, r8d
    mov r9d, r13d
    mov eax, r14d
    mov [rsp+20h], eax
    mov [rsp+28h], r15
    mov dword [rsp+30h], 0
    mov dword [rsp+38h], 0
    mov dword [rsp+40h], 00CC0020h ; SRCCOPY
    call [BitBlt]

    ; Cleanup
    mov rcx, r15
    mov rdx, [rbp-50h]
    call [SelectObject]
    mov rcx, [rbp-48h]
    call [DeleteObject]
    mov rcx, r15
    call [DeleteDC]

    mov rcx, [hChartWnd]
    lea rdx, [chart_ps]
    call [EndPaint]

    leave
.proc_ret0:
    xor eax, eax
.proc_ret:
    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    mov r14, [rbp-30h]
    mov r15, [rbp-38h]
    leave
    ret

; =============================================================================
; RenderChartToDC
; RCX = memDC, EDX = Width, R8D = Height
; =============================================================================
RenderChartToDC:
    push rbp
    mov rbp, rsp
    sub rsp, 0C0h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13

    mov r12, rcx        ; r12 = memDC

    ; Set plot area boundaries into static variables
    mov dword [plot_left], 70
    mov dword [plot_top], 85
    mov eax, edx
    sub eax, 150
    mov [plot_right], eax
    mov eax, r8d
    sub eax, 55
    mov [plot_bottom], eax

    ; 1. Draw Chart Title
    mov rcx, r12
    mov edx, 1          ; TRANSPARENT
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00222222h ; dark text
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]

    mov eax, [cur_chart_type]
    dec eax
    lea rcx, [chart_type_names]
    mov rdx, [rcx + rax*8]
    mov r13, rdx        ; title str
    mov rcx, rdx
    call [strlen]

    mov rcx, r12
    mov edx, 70
    mov r8d, 52
    mov r9, r13
    mov [rsp+20h], eax
    call [TextOutA]

    ; 2. Draw Legend on the right side
    mov rcx, r12
    call DrawChartLegend

    ; 3. Render specific chart type
    mov eax, [cur_chart_type]
    cmp eax, 1
    je .render_column
    cmp eax, 2
    je .render_line
    cmp eax, 3
    je .render_pie
    cmp eax, 4
    je .render_bar
    cmp eax, 5
    je .render_area

.render_column:
    mov rcx, r12
    call RenderColumnChart
    jmp .render_done

.render_line:
    mov rcx, r12
    call RenderLineChart
    jmp .render_done

.render_pie:
    mov rcx, r12
    call RenderPieChart
    jmp .render_done

.render_bar:
    mov rcx, r12
    call RenderBarChart
    jmp .render_done

.render_area:
    mov rcx, r12
    call RenderAreaChart
    jmp .render_done

.render_done:
    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    leave
    ret

; =============================================================================
; DrawChartAxesAndGrid
; RCX = memDC
; Draws Y-axis, X-axis, horizontal gridlines, and numeric tick values.
; =============================================================================
DrawChartAxesAndGrid:
    push rbp
    mov rbp, rsp
    sub rsp, 0C0h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13

    mov r12, rcx

    ; Horizontal gridlines: 5 intervals (steps 0..5)
    xor ebx, ebx        ; ebx = step index
.grid_loop:
    cmp ebx, 5
    jg .axes_lines

    ; y_tick = plot_bottom - (step * (plot_bottom - plot_top)) / 5
    mov eax, [plot_bottom]
    sub eax, [plot_top]
    imul eax, ebx
    cdq
    mov ecx, 5
    idiv ecx
    mov esi, [plot_bottom]
    sub esi, eax        ; esi = y_tick

    ; Draw gridline in subtle border color
    mov rdx, [hPenBorder]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, [plot_left]
    mov r8d, esi
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, [plot_right]
    mov r8d, esi
    call [LineTo]

    ; Format tick label: min_val + step * (max_val - min_val) / 5
    movsd xmm0, [chart_max_val]
    movsd xmm1, [chart_min_val]
    subsd xmm0, xmm1
    cvtsi2sd xmm2, ebx
    mulsd xmm0, xmm2
    divsd xmm0, [const_five]
    addsd xmm0, xmm1

    lea rcx, [chart_val_buf]
    lea rdx, [chart_fmt_int]
    movq r8, xmm0
    movsd xmm2, xmm0
    call [sprintf]

    ; Draw right-aligned tick label
    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00555555h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]

    lea rcx, [chart_val_buf]
    call [strlen]
    mov rcx, r12
    mov edx, [plot_left]
    sub edx, 42         ; label x
    lea r8d, [esi-7]    ; label y
    lea r9, [chart_val_buf]
    mov [rsp+20h], eax
    call [TextOutA]

    inc ebx
    jmp .grid_loop

.axes_lines:
    ; Draw solid darker X-axis baseline
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, [plot_left]
    mov r8d, [plot_bottom]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, [plot_right]
    mov r8d, [plot_bottom]
    call [LineTo]

    ; Draw solid Y-axis line
    mov rcx, r12
    mov edx, [plot_left]
    mov r8d, [plot_top]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, [plot_left]
    mov r8d, [plot_bottom]
    call [LineTo]

    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    leave
    ret

; =============================================================================
; DrawChartAxesAndGridHorizontal
; RCX = memDC
; Used for Horizontal Bar Charts: vertical gridlines and numeric ticks along X axis.
; =============================================================================
DrawChartAxesAndGridHorizontal:
    push rbp
    mov rbp, rsp
    sub rsp, 0C0h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13

    mov r12, rcx

    ; Vertical gridlines: 5 intervals (steps 0..5)
    xor ebx, ebx        ; ebx = step index
.hgrid_loop:
    cmp ebx, 5
    jg .haxes_lines

    ; x_tick = plot_left + (step * (plot_right - plot_left)) / 5
    mov eax, [plot_right]
    sub eax, [plot_left]
    imul eax, ebx
    cdq
    mov ecx, 5
    idiv ecx
    add eax, [plot_left]
    mov esi, eax        ; esi = x_tick

    ; Draw vertical gridline in subtle border color
    mov rdx, [hPenBorder]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, esi
    mov r8d, [plot_top]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, esi
    mov r8d, [plot_bottom]
    call [LineTo]

    ; Format tick label: min_val + step * (max_val - min_val) / 5
    movsd xmm0, [chart_max_val]
    movsd xmm1, [chart_min_val]
    subsd xmm0, xmm1
    cvtsi2sd xmm2, ebx
    mulsd xmm0, xmm2
    divsd xmm0, [const_five]
    addsd xmm0, xmm1

    lea rcx, [chart_val_buf]
    lea rdx, [chart_fmt_int]
    movq r8, xmm0
    movsd xmm2, xmm0
    call [sprintf]

    ; Draw tick label below bottom axis
    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00555555h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]

    lea rcx, [chart_val_buf]
    call [strlen]
    mov [rbp-70h], eax

    mov rcx, r12
    lea edx, [esi-10]   ; center label under tick
    mov eax, [plot_bottom]
    lea r8d, [eax+8]
    lea r9, [chart_val_buf]
    mov eax, [rbp-70h]
    mov [rsp+20h], eax
    call [TextOutA]

    inc ebx
    jmp .hgrid_loop

.haxes_lines:
    ; Draw solid darker X-axis baseline
    mov rdx, [hPenDark]
    mov rcx, r12
    call [SelectObject]
    mov rcx, r12
    mov edx, [plot_left]
    mov r8d, [plot_bottom]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, [plot_right]
    mov r8d, [plot_bottom]
    call [LineTo]

    ; Draw solid Y-axis line
    mov rcx, r12
    mov edx, [plot_left]
    mov r8d, [plot_top]
    xor r9d, r9d
    call [MoveToEx]
    mov rcx, r12
    mov edx, [plot_left]
    mov r8d, [plot_bottom]
    call [LineTo]

    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    leave
    ret

; =============================================================================
; GetThemeSeriesColor
; ECX = Series index (0..5), EDX = Theme (0..3)
; Returns EAX = COLORREF (BGR)
; =============================================================================
GetThemeSeriesColor:
    cmp ecx, 5
    jle @f
    mov ecx, 5
@@:
    and edx, 3
    ; Theme 0: Excel Classic
    test edx, edx
    jnz .t1
    cmp ecx, 0
    je .c_blue
    cmp ecx, 1
    je .c_orange
    cmp ecx, 2
    je .c_green
    cmp ecx, 3
    je .c_gold
    cmp ecx, 4
    je .c_azure
    mov eax, 00A03070h ; Purple
    ret
.c_blue:
    mov eax, 00BD5A18h ; #185ABD
    ret
.c_orange:
    mov eax, 00317DEDh ; #ED7D31
    ret
.c_green:
    mov eax, 0047AD70h ; #70AD47
    ret
.c_gold:
    mov eax, 0000C0FFh ; #FFC000
    ret
.c_azure:
    mov eax, 00C47244h ; #4472C4
    ret

.t1:
    cmp edx, 1
    jne .t2
    ; Theme 1: Modern Teal
    cmp ecx, 0
    je .c_teal
    cmp ecx, 1
    je .c_cyan
    cmp ecx, 2
    je .c_forest
    cmp ecx, 3
    je .c_tangerine
    cmp ecx, 4
    je .c_coral
    mov eax, 00D47800h
    ret
.c_teal:
    mov eax, 00728200h
    ret
.c_cyan:
    mov eax, 00C3B700h
    ret
.c_forest:
    mov eax, 00417C10h
    ret
.c_tangerine:
    mov eax, 00013BD8h
    ret
.c_coral:
    mov eax, 000043EAh
    ret

.t2:
    cmp edx, 2
    jne .t3
    ; Theme 2: Vibrant Palette
    cmp ecx, 0
    je .c_ruby
    cmp ecx, 1
    je .c_cobalt
    cmp ecx, 2
    je .c_emerald
    cmp ecx, 3
    je .c_amber
    cmp ecx, 4
    je .c_violet
    mov eax, 0094B200h
    ret
.c_ruby:
    mov eax, 002311E8h
    ret
.c_cobalt:
    mov eax, 00D77800h
    ret
.c_emerald:
    mov eax, 00107C10h
    ret
.c_amber:
    mov eax, 0000B9FFh
    ret
.c_violet:
    mov eax, 00981788h
    ret

.t3:
    ; Theme 3: Warm Sunset
    cmp ecx, 0
    je .c_crimson
    cmp ecx, 1
    je .c_rust
    cmp ecx, 2
    je .c_magenta
    cmp ecx, 3
    je .c_plum
    cmp ecx, 4
    je .c_indigo
    mov eax, 00107C5Ch
    ret
.c_crimson:
    mov eax, 00013BD8h
    ret
.c_rust:
    mov eax, 00008CFFh
    ret
.c_magenta:
    mov eax, 008C00E3h
    ret
.c_plum:
    mov eax, 00912D5Ch
    ret
.c_indigo:
    mov eax, 00D77800h
    ret

; =============================================================================
; RenderColumnChart
; Renders clustered vertical columns with Excel styling, borders, and value labels.
; =============================================================================
RenderColumnChart:
    push rbp
    mov rbp, rsp
    sub rsp, 0C0h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13
    mov [rbp-30h], r14
    mov [rbp-38h], r15

    mov r12, rcx

    call DrawChartAxesAndGrid

    ; Category slot width:
    mov eax, [plot_right]
    sub eax, [plot_left]
    cdq
    mov ecx, [chart_num_cats]
    cmp ecx, 1
    jge @f
    mov ecx, 1
@@:
    idiv ecx
    mov [plot_slot_w], eax
    mov r13d, eax       ; slot_w

    ; Series column width inside category:
    ; cluster_w = slot_w * 70%
    mov eax, r13d
    imul eax, 7
    cdq
    mov ecx, 10
    idiv ecx
    cdq
    mov ecx, [chart_num_series]
    cmp ecx, 1
    jge @f
    mov ecx, 1
@@:
    idiv ecx
    mov [plot_bar_w], eax
    mov r14d, eax       ; bar_w
    cmp r14d, 6
    jge @f
    mov r14d, 6
    mov [plot_bar_w], 6
@@:

    ; Render categories
    xor ebx, ebx        ; ebx = cat index
.col_cat_loop:
    cmp ebx, [chart_num_cats]
    jge .col_done

    ; Category left x = plot_left + cat * slot_w
    mov eax, ebx
    imul eax, r13d
    add eax, [plot_left]
    mov esi, eax        ; esi = cat_x

    ; Draw category label below axis:
    lea rdi, [chart_cat_names]
    mov eax, ebx
    shl eax, 6
    add rdi, rax
    mov rcx, rdi
    call [strlen]
    mov [rbp-70h], eax

    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00333333h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]

    mov rcx, r12
    mov edx, esi
    mov eax, [plot_bottom]
    lea r8d, [eax+8]
    mov r9, rdi
    mov eax, [rbp-70h]
    mov [rsp+20h], eax
    call [TextOutA]

    ; Render series columns inside category
    xor edi, edi        ; edi = series index
.col_series_loop:
    cmp edi, [chart_num_series]
    jge .next_col_cat

    ; Bar left = cat_x + (slot_w * 15%) + series * bar_w
    mov eax, r13d
    imul eax, 15
    cdq
    mov ecx, 100
    idiv ecx
    add eax, esi
    mov ecx, edi
    imul ecx, r14d
    add eax, ecx
    mov [rbp-40h], eax  ; bar_x1

    lea edx, [eax + r14d - 2]
    mov [rbp-48h], edx  ; bar_x2

    ; Compute bar height from data value:
    ; ratio = (val - min_val) / (max_val - min_val)
    mov eax, ebx
    shl eax, 3
    add eax, edi
    movsd xmm0, [chart_data + rax*8]
    movsd [rbp-58h], xmm0 ; save value

    movsd xmm1, [chart_min_val]
    subsd xmm0, xmm1
    movsd xmm2, [chart_max_val]
    subsd xmm2, xmm1
    divsd xmm0, xmm2    ; ratio

    mov eax, [plot_bottom]
    sub eax, [plot_top]
    cvtsi2sd xmm1, eax
    mulsd xmm0, xmm1
    cvttsd2si eax, xmm0

    mov edx, [plot_bottom]
    sub edx, eax
    mov [rbp-44h], edx  ; bar_y1 (top of bar)
    mov eax, [plot_bottom]
    mov [rbp-4Ch], eax  ; bar_y2 (bottom of bar)

    ; Create solid brush and pen with series theme color
    mov ecx, edi
    mov edx, [cur_chart_theme]
    call GetThemeSeriesColor
    mov [rbp-50h], eax

    mov ecx, [rbp-50h]
    call [CreateSolidBrush]
    mov [rbp-60h], rax  ; hBrush

    xor ecx, ecx
    mov edx, 1
    mov r8d, [rbp-50h]
    call [CreatePen]
    mov [rbp-68h], rax  ; hPen

    mov rcx, r12
    mov rdx, [rbp-60h]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [rbp-68h]
    call [SelectObject]

    ; Draw column rectangle
    mov rcx, r12
    mov edx, [rbp-40h]
    mov r8d, [rbp-44h]
    mov r9d, [rbp-48h]
    mov eax, [rbp-4Ch]
    mov [rsp+20h], eax
    call [Rectangle]

    ; Draw numeric value label above column
    lea rcx, [chart_val_buf]
    lea rdx, [chart_fmt_int]
    movsd xmm0, [rbp-58h]
    movq r8, xmm0
    movsd xmm2, xmm0
    call [sprintf]

    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00444444h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]

    lea rcx, [chart_val_buf]
    call [strlen]
    mov rcx, r12
    mov edx, [rbp-40h]
    sub edx, 2
    mov r8d, [rbp-44h]
    sub r8d, 15         ; 15px above bar
    lea r9, [chart_val_buf]
    mov [rsp+20h], eax
    call [TextOutA]

    ; Cleanup brush & pen
    mov rcx, [rbp-60h]
    call [DeleteObject]
    mov rcx, [rbp-68h]
    call [DeleteObject]

    inc edi
    jmp .col_series_loop

.next_col_cat:
    inc ebx
    jmp .col_cat_loop

.col_done:
    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    mov r14, [rbp-30h]
    mov r15, [rbp-38h]
    leave
    ret

; =============================================================================
; RenderLineChart
; Renders smooth line series with circular vertex markers and numeric labels.
; =============================================================================
RenderLineChart:
    push rbp
    mov rbp, rsp
    sub rsp, 0C0h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13
    mov [rbp-30h], r14
    mov [rbp-38h], r15

    mov r12, rcx

    call DrawChartAxesAndGrid

    ; Category slot width:
    mov eax, [plot_right]
    sub eax, [plot_left]
    cdq
    mov ecx, [chart_num_cats]
    cmp ecx, 1
    jge @f
    mov ecx, 1
@@:
    idiv ecx
    mov r13d, eax       ; slot_w

    ; Category labels below baseline
    xor ebx, ebx
.line_lbl_loop:
    cmp ebx, [chart_num_cats]
    jge .line_series_start

    mov eax, ebx
    imul eax, r13d
    add eax, [plot_left]
    mov esi, eax        ; cat_x

    lea rdi, [chart_cat_names]
    mov eax, ebx
    shl eax, 6
    add rdi, rax
    mov rcx, rdi
    call [strlen]
    mov [rbp-70h], eax

    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00333333h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]

    mov rcx, r12
    mov edx, esi
    mov eax, [plot_bottom]
    lea r8d, [eax+8]
    mov r9, rdi
    mov eax, [rbp-70h]
    mov [rsp+20h], eax
    call [TextOutA]

    inc ebx
    jmp .line_lbl_loop

.line_series_start:
    ; Draw each series line
    xor edi, edi        ; edi = series index
.line_series_loop:
    cmp edi, [chart_num_series]
    jge .line_done

    mov ecx, edi
    mov edx, [cur_chart_theme]
    call GetThemeSeriesColor
    mov [rbp-50h], eax  ; color

    ; Create 2px pen for line
    xor ecx, ecx
    mov edx, 2
    mov r8d, [rbp-50h]
    call [CreatePen]
    mov [rbp-60h], rax  ; hLinePen

    mov rcx, r12
    mov rdx, [rbp-60h]
    call [SelectObject]

    ; 1. Draw connecting lines
    xor ebx, ebx
.line_pt_loop:
    cmp ebx, [chart_num_cats]
    jge .draw_markers

    ; x = plot_left + cat * slot_w + slot_w / 2
    mov eax, ebx
    imul eax, r13d
    add eax, [plot_left]
    mov ecx, r13d
    shr ecx, 1
    add eax, ecx
    mov esi, eax        ; pt_x

    ; y = bottom - ratio * (bottom - top)
    mov eax, ebx
    shl eax, 3
    add eax, edi
    movsd xmm0, [chart_data + rax*8]
    movsd xmm1, [chart_min_val]
    subsd xmm0, xmm1
    movsd xmm2, [chart_max_val]
    subsd xmm2, xmm1
    divsd xmm0, xmm2

    mov eax, [plot_bottom]
    sub eax, [plot_top]
    cvtsi2sd xmm1, eax
    mulsd xmm0, xmm1
    cvttsd2si eax, xmm0
    mov edx, [plot_bottom]
    sub edx, eax
    mov r14d, edx       ; pt_y

    test ebx, ebx
    jnz .line_to
    mov rcx, r12
    mov edx, esi
    mov r8d, r14d
    xor r9d, r9d
    call [MoveToEx]
    jmp .next_lpt
.line_to:
    mov rcx, r12
    mov edx, esi
    mov r8d, r14d
    call [LineTo]

.next_lpt:
    inc ebx
    jmp .line_pt_loop

.draw_markers:
    ; 2. Draw circular vertex markers filled with white and bordered with series color
    mov rcx, r12
    mov rdx, [hBrWhite]
    call [SelectObject]

    xor ebx, ebx
.marker_loop:
    cmp ebx, [chart_num_cats]
    jge .finish_series_line

    mov eax, ebx
    imul eax, r13d
    add eax, [plot_left]
    mov ecx, r13d
    shr ecx, 1
    add eax, ecx
    mov esi, eax        ; pt_x

    mov eax, ebx
    shl eax, 3
    add eax, edi
    movsd xmm0, [chart_data + rax*8]
    movsd [rbp-58h], xmm0 ; save value
    movsd xmm1, [chart_min_val]
    subsd xmm0, xmm1
    movsd xmm2, [chart_max_val]
    subsd xmm2, xmm1
    divsd xmm0, xmm2

    mov eax, [plot_bottom]
    sub eax, [plot_top]
    cvtsi2sd xmm1, eax
    mulsd xmm0, xmm1
    cvttsd2si eax, xmm0
    mov edx, [plot_bottom]
    sub edx, eax
    mov r14d, edx       ; pt_y

    ; Marker circle: (x-4, y-4) to (x+5, y+5)
    mov rcx, r12
    lea edx, [esi-4]
    lea r8d, [r14d-4]
    lea r9d, [esi+5]
    lea eax, [r14d+5]
    mov [rsp+20h], eax
    call [Ellipse]

    ; Value label above marker
    lea rcx, [chart_val_buf]
    lea rdx, [chart_fmt_int]
    movsd xmm0, [rbp-58h]
    movq r8, xmm0
    movsd xmm2, xmm0
    call [sprintf]

    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00444444h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]

    lea rcx, [chart_val_buf]
    call [strlen]
    mov rcx, r12
    lea edx, [esi-8]
    lea r8d, [r14d-17]
    lea r9, [chart_val_buf]
    mov [rsp+20h], eax
    call [TextOutA]

    inc ebx
    jmp .marker_loop

.finish_series_line:
    mov rcx, [rbp-60h]
    call [DeleteObject]

    inc edi
    jmp .line_series_loop

.line_done:
    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    mov r14, [rbp-30h]
    mov r15, [rbp-38h]
    leave
    ret

; =============================================================================
; RenderPieChart
; Renders proportional circular pie slices using GDI Pie, with percentages and labels.
; =============================================================================
RenderPieChart:
    push rbp
    mov rbp, rsp
    sub rsp, 0C0h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13
    mov [rbp-30h], r14
    mov [rbp-38h], r15

    mov r12, rcx

    ; Compute total value for series 0
    xorpd xmm0, xmm0
    xor ebx, ebx
.pie_sum_loop:
    cmp ebx, [chart_num_cats]
    jge .pie_sum_done
    mov eax, ebx
    shl eax, 3          ; series 0
    addsd xmm0, [chart_data + rax*8]
    inc ebx
    jmp .pie_sum_loop

.pie_sum_done:
    xorpd xmm1, xmm1
    comisd xmm0, xmm1
    ja @f
    mov rax, 4000000000000000h ; 2.0 fallback
    movq xmm0, rax
@@:
    movsd [pie_total_val], xmm0

    ; Center of pie:
    ; cx = (left + right) / 2, cy = (top + bottom) / 2
    mov eax, [plot_right]
    add eax, [plot_left]
    shr eax, 1
    mov [pie_cx], eax
    mov r13d, eax       ; cx

    mov eax, [plot_bottom]
    add eax, [plot_top]
    shr eax, 1
    mov [pie_cy], eax
    mov r14d, eax       ; cy

    ; Radius R = 140
    mov dword [pie_radius], 140
    mov r15d, 140

    ; White separator pen for pie slice boundaries
    xor ecx, ecx
    mov edx, 2
    mov r8d, 00FFFFFFh
    call [CreatePen]
    mov [rbp-48h], rax

    ; Initial angle = 0.0
    xorpd xmm3, xmm3
    movsd [pie_cur_angle], xmm3

    ; Draw each slice
    xor ebx, ebx
.pie_slice_loop:
    cmp ebx, [chart_num_cats]
    jge .pie_done

    mov eax, ebx
    shl eax, 3
    movsd xmm0, [chart_data + rax*8]
    movsd [rbp-58h], xmm0 ; val

    ; Slice angle delta = (val / total) * const_two_pi
    divsd xmm0, [pie_total_val]
    mulsd xmm0, [const_two_pi]
    movsd [pie_d_angle], xmm0

    ; next_angle = cur_angle + d_angle
    movsd xmm4, [pie_cur_angle]
    addsd xmm4, xmm0
    movsd [pie_next_angle], xmm4

    ; Compute start point on circle: x1 = cx + R * cos(a0), y1 = cy - R * sin(a0)
    fld qword [pie_cur_angle]
    fsincos
    fimul dword [pie_radius]
    fiadd dword [pie_cx]
    fistp dword [pie_x_start]
    fimul dword [pie_radius]
    fchs
    fiadd dword [pie_cy]
    fistp dword [pie_y_start]

    ; Compute end point on circle: x2 = cx + R * cos(a1), y2 = cy - R * sin(a1)
    fld qword [pie_next_angle]
    fsincos
    fimul dword [pie_radius]
    fiadd dword [pie_cx]
    fistp dword [pie_x_end]
    fimul dword [pie_radius]
    fchs
    fiadd dword [pie_cy]
    fistp dword [pie_y_end]

    ; Slice theme brush:
    mov ecx, ebx
    mov edx, [cur_chart_theme]
    call GetThemeSeriesColor
    mov ecx, eax
    call [CreateSolidBrush]
    mov rsi, rax        ; hBrush

    mov rcx, r12
    mov rdx, rsi
    call [SelectObject]
    mov rcx, r12
    mov rdx, [rbp-48h]  ; white separator pen
    call [SelectObject]

    ; Draw Pie slice:
    mov rcx, r12
    mov edx, [pie_cx]
    sub edx, [pie_radius]
    mov r8d, [pie_cy]
    sub r8d, [pie_radius]
    mov r9d, [pie_cx]
    add r9d, [pie_radius]
    mov eax, [pie_cy]
    add eax, [pie_radius]
    mov [rsp+20h], eax
    mov eax, [pie_x_start]
    mov [rsp+28h], eax
    mov eax, [pie_y_start]
    mov [rsp+30h], eax
    mov eax, [pie_x_end]
    mov [rsp+38h], eax
    mov eax, [pie_y_end]
    mov [rsp+40h], eax
    call [Pie]

    mov rcx, rsi
    call [DeleteObject]

    ; Percentage label: pct = (val / total) * 100.0
    movsd xmm0, [rbp-58h]
    divsd xmm0, [pie_total_val]
    mulsd xmm0, [const_hundred]

    ; Category name
    lea r8, [chart_cat_names]
    mov eax, ebx
    shl eax, 6
    add r8, rax

    lea rcx, [chart_val_buf]
    lea rdx, [chart_fmt_pie_lbl]
    movq r9, xmm0
    movsd xmm3, xmm0
    call [sprintf]

    ; Label position: a_mid = cur_angle + d_angle * 0.5
    movsd xmm0, [pie_cur_angle]
    movsd xmm1, [pie_d_angle]
    mulsd xmm1, [const_half]
    addsd xmm0, xmm1
    movsd [pie_mid_angle], xmm0

    ; r_mid = radius * 65%
    mov dword [pie_r_mid], 90

    fld qword [pie_mid_angle]
    fsincos
    fimul dword [pie_r_mid]
    fiadd dword [pie_cx]
    fistp dword [pie_x_start]
    fimul dword [pie_r_mid]
    fchs
    fiadd dword [pie_cy]
    fistp dword [pie_y_start]

    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00FFFFFFh ; white label text
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontBold]
    call [SelectObject]

    lea rcx, [chart_val_buf]
    call [strlen]
    mov rcx, r12
    mov edx, [pie_x_start]
    sub edx, 26
    mov r8d, [pie_y_start]
    sub r8d, 7
    lea r9, [chart_val_buf]
    mov [rsp+20h], eax
    call [TextOutA]

    ; Move to next angle
    movsd xmm0, [pie_next_angle]
    movsd [pie_cur_angle], xmm0

    inc ebx
    jmp .pie_slice_loop

.pie_done:
    mov rcx, [rbp-48h]
    call [DeleteObject]

    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    mov r14, [rbp-30h]
    mov r15, [rbp-38h]
    leave
    ret

; =============================================================================
; RenderBarChart
; Renders horizontal bars extending from the Y-axis to the right.
; =============================================================================
RenderBarChart:
    push rbp
    mov rbp, rsp
    sub rsp, 0C0h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13
    mov [rbp-30h], r14
    mov [rbp-38h], r15

    mov r12, rcx

    call DrawChartAxesAndGridHorizontal

    ; Category slot height:
    mov eax, [plot_bottom]
    sub eax, [plot_top]
    cdq
    mov ecx, [chart_num_cats]
    cmp ecx, 1
    jge @f
    mov ecx, 1
@@:
    idiv ecx
    mov r13d, eax       ; slot_h

    ; Bar height inside slot:
    mov eax, r13d
    imul eax, 7
    cdq
    mov ecx, 10
    idiv ecx
    cdq
    mov ecx, [chart_num_series]
    cmp ecx, 1
    jge @f
    mov ecx, 1
@@:
    idiv ecx
    mov r14d, eax       ; bar_h
    cmp r14d, 6
    jge @f
    mov r14d, 6
@@:

    ; Render categories
    xor ebx, ebx
.bar_cat_loop:
    cmp ebx, [chart_num_cats]
    jge .bar_done

    mov eax, ebx
    imul eax, r13d
    add eax, [plot_top]
    mov esi, eax        ; cat_y

    ; Category label on the left:
    lea rdi, [chart_cat_names]
    mov eax, ebx
    shl eax, 6
    add rdi, rax

    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00333333h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]

    mov rcx, rdi
    call [strlen]

    mov rcx, r12
    mov edx, [plot_left]
    sub edx, 48         ; left of Y axis
    mov r8d, r13d
    shr r8d, 1
    add r8d, esi
    sub r8d, 7
    mov r9, rdi
    mov [rsp+20h], eax
    call [TextOutA]

    ; Render series bars horizontally
    xor edi, edi
.bar_series_loop:
    cmp edi, [chart_num_series]
    jge .next_bar_cat

    ; bar y1 = cat_y + (slot_h*15%) + series * bar_h
    mov eax, r13d
    imul eax, 15
    cdq
    mov ecx, 100
    idiv ecx
    add eax, esi
    mov ecx, edi
    imul ecx, r14d
    add eax, ecx
    mov [rbp-44h], eax  ; bar_y1
    lea edx, [eax + r14d - 2]
    mov [rbp-4Ch], edx  ; bar_y2

    ; bar x1 = plot_left
    mov eax, [plot_left]
    mov [rbp-40h], eax  ; bar_x1

    ; bar x2 = plot_left + ratio * (plot_right - plot_left)
    mov eax, ebx
    shl eax, 3
    add eax, edi
    movsd xmm0, [chart_data + rax*8]
    movsd [rbp-58h], xmm0 ; save value

    movsd xmm1, [chart_min_val]
    subsd xmm0, xmm1
    movsd xmm2, [chart_max_val]
    subsd xmm2, xmm1
    divsd xmm0, xmm2

    mov eax, [plot_right]
    sub eax, [plot_left]
    cvtsi2sd xmm1, eax
    mulsd xmm0, xmm1
    cvttsd2si eax, xmm0
    add eax, [plot_left]
    mov [rbp-48h], eax  ; bar_x2

    ; Create brush and pen
    mov ecx, edi
    mov edx, [cur_chart_theme]
    call GetThemeSeriesColor
    mov [rbp-50h], eax

    mov ecx, [rbp-50h]
    call [CreateSolidBrush]
    mov [rbp-60h], rax

    xor ecx, ecx
    mov edx, 1
    mov r8d, [rbp-50h]
    call [CreatePen]
    mov [rbp-68h], rax

    mov rcx, r12
    mov rdx, [rbp-60h]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [rbp-68h]
    call [SelectObject]

    mov rcx, r12
    mov edx, [rbp-40h]
    mov r8d, [rbp-44h]
    mov r9d, [rbp-48h]
    mov eax, [rbp-4Ch]
    mov [rsp+20h], eax
    call [Rectangle]

    ; Value label to the right of horizontal bar
    lea rcx, [chart_val_buf]
    lea rdx, [chart_fmt_int]
    movsd xmm0, [rbp-58h]
    movq r8, xmm0
    movsd xmm2, xmm0
    call [sprintf]

    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00444444h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]

    lea rcx, [chart_val_buf]
    call [strlen]
    mov rcx, r12
    mov edx, [rbp-48h]
    add edx, 6          ; 6px right of bar
    mov r8d, [rbp-44h]
    lea r9, [chart_val_buf]
    mov [rsp+20h], eax
    call [TextOutA]

    mov rcx, [rbp-60h]
    call [DeleteObject]
    mov rcx, [rbp-68h]
    call [DeleteObject]

    inc edi
    jmp .bar_series_loop

.next_bar_cat:
    inc ebx
    jmp .bar_cat_loop

.bar_done:
    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    mov r14, [rbp-30h]
    mov r15, [rbp-38h]
    leave
    ret

; =============================================================================
; RenderAreaChart
; Renders filled areas underneath lines down to zero baseline.
; =============================================================================
RenderAreaChart:
    push rbp
    mov rbp, rsp
    sub rsp, 0C0h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13
    mov [rbp-30h], r14
    mov [rbp-38h], r15

    mov r12, rcx

    call DrawChartAxesAndGrid

    ; Category slot width:
    mov eax, [plot_right]
    sub eax, [plot_left]
    cdq
    mov ecx, [chart_num_cats]
    cmp ecx, 1
    jge @f
    mov ecx, 1
@@:
    idiv ecx
    mov r13d, eax       ; slot_w

    ; Category labels below baseline
    xor ebx, ebx
.area_lbl_loop:
    cmp ebx, [chart_num_cats]
    jge .area_series_start

    mov eax, ebx
    imul eax, r13d
    add eax, [plot_left]
    mov esi, eax        ; cat_x

    lea rdi, [chart_cat_names]
    mov eax, ebx
    shl eax, 6
    add rdi, rax

    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00333333h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]

    mov rcx, rdi
    call [strlen]
    mov [rbp-70h], eax

    mov rcx, r12
    mov edx, esi
    mov eax, [plot_bottom]
    lea r8d, [eax+8]
    mov r9, rdi
    mov eax, [rbp-70h]
    mov [rsp+20h], eax
    call [TextOutA]

    inc ebx
    jmp .area_lbl_loop

.area_series_start:
    ; Draw each series from first to last (so smaller series are visible on top)
    xor edi, edi
.area_series_loop:
    cmp edi, [chart_num_series]
    jge .area_done

    mov ecx, edi
    mov edx, [cur_chart_theme]
    call GetThemeSeriesColor
    mov [rbp-50h], eax  ; color

    ; 1. Build polygon points:
    ; Point 0 = (first_x, plot_bottom)
    mov eax, [plot_left]
    mov ecx, r13d
    shr ecx, 1
    add eax, ecx
    mov [chart_poly_pts + 0*8], eax
    mov ecx, [plot_bottom]
    mov [chart_poly_pts + 0*8 + 4], ecx

    ; Points 1..N
    xor ebx, ebx
.area_pt_loop:
    cmp ebx, [chart_num_cats]
    jge .area_poly_close

    ; x = plot_left + cat * slot_w + slot_w / 2
    mov eax, ebx
    imul eax, r13d
    add eax, [plot_left]
    mov ecx, r13d
    shr ecx, 1
    add eax, ecx
    mov esi, eax        ; pt_x

    ; y = bottom - ratio * (bottom - top)
    mov eax, ebx
    shl eax, 3
    add eax, edi
    movsd xmm0, [chart_data + rax*8]
    movsd [rbp-58h], xmm0
    movsd xmm1, [chart_min_val]
    subsd xmm0, xmm1
    movsd xmm2, [chart_max_val]
    subsd xmm2, xmm1
    divsd xmm0, xmm2

    mov eax, [plot_bottom]
    sub eax, [plot_top]
    cvtsi2sd xmm1, eax
    mulsd xmm0, xmm1
    cvttsd2si eax, xmm0
    mov edx, [plot_bottom]
    sub edx, eax
    mov r14d, edx       ; pt_y

    ; Store in chart_poly_pts[(ebx + 1) * 8]
    lea edx, [ebx + 1]
    mov [chart_poly_pts + rdx*8], esi
    mov [chart_poly_pts + rdx*8 + 4], r14d

    inc ebx
    jmp .area_pt_loop

.area_poly_close:
    ; Point N+1 = (last_x, plot_bottom)
    lea edx, [ebx + 1]
    mov [chart_poly_pts + rdx*8], esi
    mov eax, [plot_bottom]
    mov [chart_poly_pts + rdx*8 + 4], eax

    ; Total points = chart_num_cats + 2
    lea r15d, [ebx + 2]

    ; Create brush and pen for area
    mov ecx, [rbp-50h]
    call [CreateSolidBrush]
    mov [rbp-60h], rax  ; hBrush

    xor ecx, ecx
    mov edx, 1
    mov r8d, [rbp-50h]
    call [CreatePen]
    mov [rbp-68h], rax  ; hPen

    mov rcx, r12
    mov rdx, [rbp-60h]
    call [SelectObject]
    mov rcx, r12
    mov rdx, [rbp-68h]
    call [SelectObject]

    ; Draw filled polygon
    mov rcx, r12
    lea rdx, [chart_poly_pts]
    mov r8d, r15d
    call [Polygon]

    mov rcx, [rbp-60h]
    call [DeleteObject]
    mov rcx, [rbp-68h]
    call [DeleteObject]

    ; Draw crisp top boundary line
    xor ecx, ecx
    mov edx, 2
    mov r8d, [rbp-50h]
    call [CreatePen]
    mov [rbp-68h], rax

    mov rcx, r12
    mov rdx, [rbp-68h]
    call [SelectObject]

    xor ebx, ebx
.area_top_line_loop:
    cmp ebx, [chart_num_cats]
    jge .area_line_done
    lea edx, [ebx + 1]
    mov esi, [chart_poly_pts + rdx*8]
    mov r14d, [chart_poly_pts + rdx*8 + 4]

    test ebx, ebx
    jnz .atl_line
    mov rcx, r12
    mov edx, esi
    mov r8d, r14d
    xor r9d, r9d
    call [MoveToEx]
    jmp .atl_next
.atl_line:
    mov rcx, r12
    mov edx, esi
    mov r8d, r14d
    call [LineTo]
.atl_next:
    inc ebx
    jmp .area_top_line_loop

.area_line_done:
    mov rcx, [rbp-68h]
    call [DeleteObject]

    inc edi
    jmp .area_series_loop

.area_done:
    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    mov r14, [rbp-30h]
    mov r15, [rbp-38h]
    leave
    ret

; =============================================================================
; DrawChartLegend
; Renders color swatches and series names at top-right of chart.
; =============================================================================
DrawChartLegend:
    push rbp
    mov rbp, rsp
    sub rsp, 0C0h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13
    mov [rbp-30h], r14

    mov eax, [plot_right]
    add eax, 25
    mov esi, eax        ; legend_x = right + 25
    mov edi, 90         ; legend_y = 90

    xor ebx, ebx
.legend_loop:
    cmp ebx, [chart_num_series]
    jge .legend_done

    ; Get series theme color
    mov ecx, ebx
    mov edx, [cur_chart_theme]
    call GetThemeSeriesColor
    mov [rbp-40h], eax

    ; Create brush
    mov ecx, [rbp-40h]
    call [CreateSolidBrush]
    mov r13, rax

    mov rcx, r12
    mov rdx, r13
    call [SelectObject]
    mov rcx, r12
    mov rdx, [hPenBorder]
    call [SelectObject]

    ; Draw 12x12 swatch rectangle
    mov rcx, r12
    mov edx, esi
    mov r8d, edi
    lea r9d, [esi+12]
    lea eax, [edi+12]
    mov [rsp+20h], eax
    call [Rectangle]

    mov rcx, r13
    call [DeleteObject]

    ; Draw series name
    lea r14, [chart_series_names]
    mov eax, ebx
    shl eax, 6
    add r14, rax

    mov rcx, r12
    mov edx, 1
    call [SetBkMode]
    mov rcx, r12
    mov edx, 00222222h
    call [SetTextColor]
    mov rcx, r12
    mov rdx, [hFontSmall]
    call [SelectObject]

    mov rcx, r14
    call [strlen]

    mov rcx, r12
    lea edx, [esi+18]
    mov r8d, edi
    mov r9, r14
    mov [rsp+20h], eax
    call [TextOutA]

    add edi, 22         ; 22px spacing between legend items
    inc ebx
    jmp .legend_loop

.legend_done:
    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    mov r14, [rbp-30h]
    leave
    ret

; =============================================================================
; CopyChartImageToClipboard
; Captures the rendered chart into a Win32 Bitmap and copies it to the clipboard.
; =============================================================================
CopyChartImageToClipboard:
    push rbp
    mov rbp, rsp
    sub rsp, 0C0h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13
    mov [rbp-30h], r14
    mov [rbp-38h], r15

    mov r12, rcx        ; r12 = hWndChart

    ; Get client rect
    lea rdx, [rbp-60h]
    call [GetClientRect]
    mov r13d, [rbp-58h] ; w
    mov r14d, [rbp-54h] ; h

    mov rcx, r12
    call [GetDC]
    mov r15, rax        ; hdcScreen

    mov rcx, r15
    call [CreateCompatibleDC]
    mov rsi, rax        ; hdcMem

    mov rcx, r15
    mov edx, r13d
    mov r8d, r14d
    call [CreateCompatibleBitmap]
    mov rdi, rax        ; hBitmap

    mov rcx, rsi
    mov rdx, rdi
    call [SelectObject]
    mov [rbp-68h], rax

    ; White fill
    mov rcx, rsi
    mov rdx, [hBrWhite]
    call [SelectObject]
    mov rcx, rsi
    mov rdx, [hPenNull]
    call [SelectObject]
    mov rcx, rsi
    xor edx, edx
    xor r8d, r8d
    mov r9d, r13d
    mov eax, r14d
    mov [rsp+20h], eax
    call [Rectangle]

    ; Render chart
    mov rcx, rsi
    mov edx, r13d
    mov r8d, r14d
    call RenderChartToDC

    ; Restore and delete memDC
    mov rcx, rsi
    mov rdx, [rbp-68h]
    call [SelectObject]
    mov rcx, rsi
    call [DeleteDC]
    mov rcx, r12
    mov rdx, r15
    call [ReleaseDC]

    ; Copy hBitmap to clipboard
    mov rcx, r12
    call [OpenClipboard]
    call [EmptyClipboard]
    mov ecx, 2          ; CF_BITMAP
    mov rdx, rdi
    call [SetClipboardData]
    call [CloseClipboard]

    ; Notification dialog
    mov rcx, r12
    lea rdx, [szChartCopiedMsg]
    lea r8, [szChartCopiedTitle]
    mov r9d, 00000040h  ; MB_OK | MB_ICONINFORMATION
    call [MessageBoxA]

    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    mov r14, [rbp-30h]
    mov r15, [rbp-38h]
    leave
    ret

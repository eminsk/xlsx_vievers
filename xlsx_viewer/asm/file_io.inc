; ============================================================================
; Excel Viewer Pro — File I/O, Multi-Sheets & Status Bar
; ============================================================================



; =============================================================================
; Build Bottom Sheet Tabs & Status Bar
; =============================================================================
BuildSheetTabsAndStatus:
    push rbp
    mov rbp, rsp
    sub rsp, 60h

    ; Sheet Tabs Control
    xor ecx, ecx
    lea rdx, [tab_class]
    xor r8, r8
    mov r9d, WS_CHILD or WS_VISIBLE
    mov dword [rsp+20h], 6
    mov dword [rsp+28h], 728
    mov dword [rsp+30h], 1254
    mov dword [rsp+38h], 28
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_SHEET_TABS
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov [hSheetTabs], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    ; Add Sheets
    lea rdi, [tcitem]
    xor eax, eax
    mov ecx, 48
    rep stosb
    mov dword [tcitem], TCIF_TEXT

    lea rax, [tab_sh1_txt]
    mov qword [tcitem+16], rax
    mov rcx, [hSheetTabs]
    mov edx, TCM_INSERTITEMA
    mov r8d, 0
    lea r9, [tcitem]
    call [SendMessageA]

    lea rax, [tab_sh2_txt]
    mov qword [tcitem+16], rax
    mov rcx, [hSheetTabs]
    mov edx, TCM_INSERTITEMA
    mov r8d, 1
    lea r9, [tcitem]
    call [SendMessageA]

    lea rax, [tab_sh3_txt]
    mov qword [tcitem+16], rax
    mov rcx, [hSheetTabs]
    mov edx, TCM_INSERTITEMA
    mov r8d, 2
    lea r9, [tcitem]
    call [SendMessageA]

    lea rax, [tab_sh_add_txt]
    mov qword [tcitem+16], rax
    mov rcx, [hSheetTabs]
    mov edx, TCM_INSERTITEMA
    mov r8d, 3
    lea r9, [tcitem]
    call [SendMessageA]

    ; Status Bar
    xor ecx, ecx
    lea rdx, [statusbar_class]
    xor r8, r8
    mov r9d, WS_CHILD or WS_VISIBLE
    mov dword [rsp+20h], 0
    mov dword [rsp+28h], 0
    mov dword [rsp+30h], 0
    mov dword [rsp+38h], 0
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_STATUSBAR
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov [hStatusBar], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    mov rcx, [hStatusBar]
    mov edx, SB_SETPARTS
    mov r8d, 5
    lea r9, [sb_parts]
    call [SendMessageA]

    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [sb_ready_txt]
    call [SendMessageA]

    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 1
    lea r9, [sb_cell_txt]
    call [SendMessageA]

    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 2
    lea r9, [sb_stats_fmt]
    call [SendMessageA]

    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 3
    lea r9, [sb_zoom_txt]
    call [SendMessageA]

    lea rcx, [szTemp]
    lea rdx, [sb_dim_fmt]
    mov r8d, [total_rows]
    call [sprintf]

    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 4
    lea r9, [szTemp]
    call [SendMessageA]

    leave
    ret


; =============================================================================
; Update Real-time SIMD Status Bar Stats (SUM, AVG, COUNT)
; =============================================================================
UpdateStatusStats:
    push rbp
    mov rbp, rsp
    sub rsp, 90h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13

    ; Local variables strictly in [rbp-30h] .. [rbp-48h]
    ; Outgoing calls use [rsp]..[rsp+28h] = [rbp-90h]..[rbp-68h]
    xor rax, rax
    mov [rbp-30h], rax       ; sum = 0.0 (qword)
    xor ebx, ebx            ; count = 0

    ; Compute min_row, max_row, min_col, max_col
    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r12d, eax           ; r12d = min_row
    mov [rbp-38h], edx      ; [rbp-38h] = max_row

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov [rbp-40h], eax      ; [rbp-40h] = min_col
    mov [rbp-48h], edx      ; [rbp-48h] = max_col

.row_loop:
    cmp r12d, [rbp-38h]
    jg .stats_done

    mov r13d, [rbp-40h]
.col_loop:
    cmp r13d, [rbp-48h]
    jg .next_row

    ; Get cell text
    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r12d
    mov dword [lvitem+8], r13d
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 1000
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r12d
    lea r9, [lvitem]
    call [SendMessageA]
    test eax, eax
    jz .next_col

    ; Parse number with strtod
    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]
    xorpd xmm1, xmm1
    ucomisd xmm0, xmm1
    jp .is_num
    jne .is_num
    cmp byte [szCellVal], '0'
    jne .next_col

.is_num:
    inc ebx
    movsd xmm2, [rbp-30h]
    addsd xmm2, xmm0
    movsd [rbp-30h], xmm2

.next_col:
    inc r13d
    jmp .col_loop

.next_row:
    inc r12d
    jmp .row_loop

.stats_done:
    movsd xmm0, [rbp-30h]   ; SUM

    ; AVG
    xorpd xmm1, xmm1
    test ebx, ebx
    jz .set_avg
    cvtsi2sd xmm2, ebx
    movsd xmm1, xmm0
    divsd xmm1, xmm2
.set_avg:

    ; Format: SUM: %.2f | AVG: %.2f | COUNT: %d
    lea rcx, [szBuffer]
    lea rdx, [sb_stats_fmt]
    movq r8, xmm0
    movq r9, xmm1
    mov [rsp+20h], rbx
    call [sprintf]

    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 2
    lea r9, [szBuffer]
    call [SendMessageA]

    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    leave
    ret


; =============================================================================
; Action: Load Demo Worksheet Data
; =============================================================================
LoadDemoData:
    push rbp
    mov rbp, rsp
    sub rsp, 60h

    ; Row 0
    mov r8d, 0
    mov r9d, 1
    lea rax, [d1_A]
    call SetCellVal
    mov r8d, 0
    mov r9d, 2
    lea rax, [d1_B]
    call SetCellVal
    mov r8d, 0
    mov r9d, 3
    lea rax, [d1_C]
    call SetCellVal
    mov r8d, 0
    mov r9d, 4
    lea rax, [d1_D]
    call SetCellVal

    ; Row 1
    mov r8d, 1
    mov r9d, 1
    lea rax, [d2_A]
    call SetCellVal
    mov r8d, 1
    mov r9d, 2
    lea rax, [d2_B]
    call SetCellVal
    mov r8d, 1
    mov r9d, 3
    lea rax, [d2_C]
    call SetCellVal
    mov r8d, 1
    mov r9d, 4
    lea rax, [d2_D]
    call SetCellVal

    ; Row 2
    mov r8d, 2
    mov r9d, 1
    lea rax, [d3_A]
    call SetCellVal
    mov r8d, 2
    mov r9d, 2
    lea rax, [d3_B]
    call SetCellVal
    mov r8d, 2
    mov r9d, 3
    lea rax, [d3_C]
    call SetCellVal
    mov r8d, 2
    mov r9d, 4
    lea rax, [d3_D]
    call SetCellVal

    ; Row 3 (Total Revenue)
    mov r8d, 3
    mov r9d, 1
    lea rax, [d4_A]
    call SetCellVal
    mov r8d, 3
    mov r9d, 2
    lea rax, [d4_B]
    call SetCellVal
    mov r8d, 3
    mov r9d, 3
    lea rax, [d4_C]
    call SetCellVal
    mov r8d, 3
    mov r9d, 4
    lea rax, [d4_D]
    call SetCellVal

    ; Row 4 (R&D)
    mov r8d, 4
    mov r9d, 1
    lea rax, [d5_A]
    call SetCellVal
    mov r8d, 4
    mov r9d, 2
    lea rax, [d5_B]
    call SetCellVal
    mov r8d, 4
    mov r9d, 3
    lea rax, [d5_C]
    call SetCellVal
    mov r8d, 4
    mov r9d, 4
    lea rax, [d5_D]
    call SetCellVal

    ; Row 5 (Net Profit)
    mov r8d, 5
    mov r9d, 1
    lea rax, [d6_A]
    call SetCellVal
    mov r8d, 5
    mov r9d, 2
    lea rax, [d6_B]
    call SetCellVal
    mov r8d, 5
    mov r9d, 3
    lea rax, [d6_C]
    call SetCellVal
    mov r8d, 5
    mov r9d, 4
    lea rax, [d6_D]
    call SetCellVal

    call UpdateStatusStats

    mov rcx, [hWnd]
    lea rdx, [open_demo_msg]
    lea r8, [wnd_title]
    xor r9d, r9d
    call [MessageBoxA]

    call SaveActiveSheetToBuffer
    leave
    ret

SetCellVal:
    push rbp
    mov rbp, rsp
    sub rsp, 40h
    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r8d
    mov dword [lvitem+8], r9d
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    lea r9, [lvitem]
    call [SendMessageA]
    leave
    ret

; =============================================================================
; Multi-Sheet Cell Persistence: Save & Load active sheet buffer
; =============================================================================
SaveActiveSheetToBuffer:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], r12
    mov [rbp-24], r13
    mov [rbp-32], r14
    mov [rbp-40], r15

    mov r12d, [cur_sheet]
    cmp r12d, 15
    ja .save_done

    xor r13d, r13d          ; row = 0..29
.save_row_loop:
    cmp r13d, [total_rows]
    jge .save_done
    cmp r13d, 30
    jge .save_done

    mov r14d, 1             ; col = 1..10
.save_col_loop:
    cmp r14d, 10
    jg .save_next_row

    ; Get cell text from ListView
    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov dword [lvitem+8], r14d
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 120
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r13d
    lea r9, [lvitem]
    call [SendMessageA]

    ; Compute offset: ((sheet * 30 + row) * 11 + col) * 128
    mov eax, r12d
    imul eax, 30
    add eax, r13d
    imul eax, 11
    add eax, r14d
    shl rax, 7
    lea rdi, [sheet_cells + rax]
    lea rsi, [szCellVal]
    mov ecx, 127
.copy_c:
    lodsb
    stosb
    test al, al
    jz .copy_end
    dec ecx
    jnz .copy_c
    mov byte [rdi], 0
.copy_end:

    inc r14d
    jmp .save_col_loop

.save_next_row:
    inc r13d
    jmp .save_row_loop

.save_done:
    mov rbx, [rbp-8]
    mov r12, [rbp-16]
    mov r13, [rbp-24]
    mov r14, [rbp-32]
    mov r15, [rbp-40]
    leave
    ret

LoadSheetFromBuffer:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], r12
    mov [rbp-24], r13
    mov [rbp-32], r14
    mov [rbp-40], r15

    mov r12d, [cur_sheet]
    cmp r12d, 15
    ja .load_done

    xor r13d, r13d          ; row = 0..29
.load_row_loop:
    cmp r13d, [total_rows]
    jge .load_finish
    cmp r13d, 30
    jge .load_finish

    mov r14d, 1             ; col = 1..10
.load_col_loop:
    cmp r14d, 10
    jg .load_next_row

    ; Compute offset: ((sheet * 30 + row) * 11 + col) * 128
    mov eax, r12d
    imul eax, 30
    add eax, r13d
    imul eax, 11
    add eax, r14d
    shl rax, 7
    lea rax, [sheet_cells + rax]

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov dword [lvitem+8], r14d
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, r13d
    lea r9, [lvitem]
    call [SendMessageA]

    inc r14d
    jmp .load_col_loop

.load_next_row:
    inc r13d
    jmp .load_row_loop

.load_finish:
    call RenumberGridRows
    call UpdateStatusStats

    ; Update Formula Bar with active cell content in this sheet
    mov dword [lvitem], LVIF_TEXT
    mov eax, [cur_row]
    mov dword [lvitem+4], eax
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 1000
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, [cur_row]
    lea r9, [lvitem]
    call [SendMessageA]

    ; Check if active cell has a formula
    mov ecx, [cur_row]
    mov edx, [cur_col]
    call GetCellFormula
    test rax, rax
    jz .use_cell_text_buf
    mov rcx, [hFormulaEdit]
    mov rdx, rax
    call [SetWindowTextA]
    jmp .load_done

.use_cell_text_buf:
    mov rcx, [hFormulaEdit]
    lea rdx, [szCellVal]
    call [SetWindowTextA]

.load_done:
    mov rbx, [rbp-8]
    mov r12, [rbp-16]
    mov r13, [rbp-24]
    mov r14, [rbp-32]
    mov r15, [rbp-40]
    leave
    ret

LoadSheetData:
    jmp LoadSheetFromBuffer


; =============================================================================
; Action: Real CSV File Open and Save
; =============================================================================
DoFileOpen:
    push rbp
    mov rbp, rsp
    sub rsp, 80h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    mov byte [szFileName], 0

    lea rdi, [ofn]
    xor eax, eax
    mov ecx, 152
    rep stosb

    mov dword [ofn], 152
    mov rax, [hWnd]
    mov qword [ofn+8], rax
    lea rax, [csv_filter]
    mov qword [ofn+24], rax
    lea rax, [szFileName]
    mov qword [ofn+48], rax
    mov dword [ofn+56], 512
    lea rax, [ofn_title_open]
    mov qword [ofn+88], rax
    mov dword [ofn+96], 00001808h ; OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR
    lea rax, [def_ext_csv]
    mov qword [ofn+104], rax

    lea rcx, [ofn]
    call [GetOpenFileNameA]
    test eax, eax
    jz .open_cancelled

    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    jmp DoFileOpenDirect

.open_cancelled:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
CopyStrHelper:
.csh_loop:
    lodsb
    stosb
    test al, al
    jnz .csh_loop
    ret

CleanupTempFiles:
    push rbp
    mov rbp, rsp
    sub rsp, 40h
    mov [rbp-8], rbx

    cmp byte [is_temp_file], 1
    jne .clean_done

    lea rcx, [szTempCsv]
    call [remove]
    lea rcx, [szTempStyle]
    call [remove]
    lea rcx, [szTempFormula]
    call [remove]
    lea rcx, [szTempSheets]
    call [remove]

    xor ebx, ebx
.clean_loop:
    inc ebx
    cmp ebx, 16
    jge .clean_done

    lea rcx, [szCurCsvPath]
    lea rdx, [fmt_sheet_csv]
    lea r8, [szTempCsv]
    mov r9d, ebx
    call [sprintf]
    lea rcx, [szCurCsvPath]
    call [remove]

    lea rcx, [szCurStylePath]
    lea rdx, [fmt_sheet_style]
    lea r8, [szTempStyle]
    mov r9d, ebx
    call [sprintf]
    lea rcx, [szCurStylePath]
    call [remove]

    lea rcx, [szCurFormulaPath]
    lea rdx, [fmt_sheet_formula]
    lea r8, [szTempCsv]
    mov r9d, ebx
    call [sprintf]
    lea rcx, [szCurFormulaPath]
    call [remove]

    jmp .clean_loop

.clean_done:
    mov rbx, [rbp-8]
    leave
    ret

SetupSheetTabsFromManifest:
    push rbp
    mov rbp, rsp
    sub rsp, 40h
    mov [rbp-8], rbx
    mov [rbp-16], r14

    ; Check if szTempSheets exists
    lea rcx, [szTempSheets]
    lea rdx, [fmt_temp_sheets]
    lea r8, [szTempCsv]
    call [sprintf]

    lea rcx, [szTempSheets]
    lea rdx, [mode_rb]
    call [fopen]
    test rax, rax
    jz .sstm_no_manifest
    mov r14, rax

    ; Read magic and count (8 bytes)
    sub rsp, 40h
    lea rcx, [rsp+20h]
    mov edx, 1
    mov r8d, 8
    mov r9, r14
    call [fread]
    cmp eax, 8
    jne .sstm_close_err
    cmp dword [rsp+20h], 53544C58h ; 'XLTS'
    jne .sstm_close_err
    mov eax, dword [rsp+24h]
    add rsp, 40h
    cmp eax, 0
    jle .sstm_close_err
    cmp eax, 16
    jle @f
    mov eax, 16
@@:
    mov [sheet_count], eax

    ; Delete all existing tabs in hSheetTabs (TCM_DELETEALLITEMS = 1309h)
    mov rcx, [hSheetTabs]
    mov edx, 1309h
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]

    xor ebx, ebx
.sstm_tab_loop:
    cmp ebx, [sheet_count]
    jge .sstm_tabs_done

    mov eax, ebx
    shl eax, 7 ; ebx * 128
    lea rsi, [g_sheet_names + rax]

    sub rsp, 40h
    mov rcx, rsi
    mov edx, 1
    mov r8d, 128
    mov r9, r14
    call [fread]
    add rsp, 40h

    ; Insert tab in hSheetTabs
    lea rdi, [tcitem]
    xor eax, eax
    mov ecx, 48
    rep stosb
    mov dword [tcitem], TCIF_TEXT
    mov qword [tcitem+16], rsi
    mov rcx, [hSheetTabs]
    mov edx, TCM_INSERTITEMA
    mov r8d, ebx
    lea r9, [tcitem]
    call [SendMessageA]

    inc ebx
    jmp .sstm_tab_loop

.sstm_tabs_done:
    ; Append [+] tab
    lea rdi, [tcitem]
    xor eax, eax
    mov ecx, 48
    rep stosb
    mov dword [tcitem], TCIF_TEXT
    lea rax, [tab_sh_add_txt]
    mov qword [tcitem+16], rax
    mov rcx, [hSheetTabs]
    mov edx, TCM_INSERTITEMA
    mov r8d, [sheet_count]
    lea r9, [tcitem]
    call [SendMessageA]

    ; Select tab 0
    mov rcx, [hSheetTabs]
    mov edx, TCM_SETCURSEL
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]
    mov dword [cur_sheet], 0

    mov rcx, r14
    call [fclose]
    jmp .sstm_ret

.sstm_close_err:
    add rsp, 40h
    mov rcx, r14
    call [fclose]

.sstm_no_manifest:
    ; Update first tab text to loaded filename
    lea rsi, [szFileName]
    mov r8, rsi
.find_fn_base_sstm:
    lodsb
    test al, al
    jz .got_fn_base_sstm
    cmp al, '\'
    cmove r8, rsi
    cmp al, '/'
    cmove r8, rsi
    jmp .find_fn_base_sstm
.got_fn_base_sstm:
    lea rdi, [tcitem]
    xor eax, eax
    mov ecx, 48
    rep stosb
    mov dword [tcitem], TCIF_TEXT
    mov qword [tcitem+16], r8
    mov rcx, [hSheetTabs]
    mov edx, 1306h ; TCM_SETITEMA
    xor r8d, r8d
    lea r9, [tcitem]
    call [SendMessageA]

.sstm_ret:
    mov rbx, [rbp-8]
    mov r14, [rbp-16]
    leave
    ret

DoLoadSheetByIndex:
    push rbp
    mov rbp, rsp
    sub rsp, 80h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13
    mov [rbp-48], r14
    mov [rbp-56], r15

    mov ebx, ecx ; ebx = sheet_idx (0, 1, 2, ...)

    cmp byte [is_temp_file], 1
    je .dls_temp_paths

    ; Non-temp (plain CSV):
    lea rsi, [szFileName]
    lea rdi, [szCurCsvPath]
    call CopyStrHelper
    mov byte [szCurStylePath], 0
    mov byte [szCurFormulaPath], 0
    jmp .dls_open_csv

.dls_temp_paths:
    test ebx, ebx
    jnz .dls_idx_gt0

    ; Sheet 0:
    lea rsi, [szTempCsv]
    lea rdi, [szCurCsvPath]
    call CopyStrHelper

    lea rsi, [szTempStyle]
    lea rdi, [szCurStylePath]
    call CopyStrHelper

    lea rcx, [szCurFormulaPath]
    lea rdx, [fmt_temp_formula]
    lea r8, [szTempCsv]
    call [sprintf]
    jmp .dls_open_csv

.dls_idx_gt0:
    lea rcx, [szCurCsvPath]
    lea rdx, [fmt_sheet_csv]
    lea r8, [szTempCsv]
    mov r9d, ebx
    call [sprintf]

    lea rcx, [szCurStylePath]
    lea rdx, [fmt_sheet_style]
    lea r8, [szTempStyle]
    mov r9d, ebx
    call [sprintf]

    lea rcx, [szCurFormulaPath]
    lea rdx, [fmt_sheet_formula]
    lea r8, [szTempCsv]
    mov r9d, ebx
    call [sprintf]

.dls_open_csv:
    lea rcx, [szCurCsvPath]
    lea rdx, [mode_r]
    call [fopen]
    test rax, rax
    jz .dls_done
    mov r12, rax ; r12 = CSV file handle

    ; 1. Clear styles table
    mov rdi, [pCellStyleTable]
    test rdi, rdi
    jz .dls_skip_st
    mov rcx, 10000 * 32 * 2
    xor eax, eax
    rep stosq
.dls_skip_st:

    ; 2. Clear formulas table and recreate heap
    mov rdi, [pCellFormulaTable]
    test rdi, rdi
    jz .dls_skip_fm
    mov rcx, 10000 * 32
    xor eax, eax
    rep stosq

    mov rcx, [hFormulaHeap]
    test rcx, rcx
    jz @f
    call [HeapDestroy]
@@:
    xor ecx, ecx
    xor edx, edx
    xor r8d, r8d
    call [HeapCreate]
    mov [hFormulaHeap], rax
.dls_skip_fm:

    ; 3. Load companion style file
    cmp byte [szCurStylePath], 0
    je .dls_style_done

    lea rcx, [szCurStylePath]
    lea rdx, [mode_rb]
    call [fopen]
    test rax, rax
    jz .dls_style_done
    mov r15, rax ; style handle

    sub rsp, 40h
    lea rcx, [rsp+20h]
    mov edx, 1
    mov r8d, 8
    mov r9, r15
    call [fread]
    cmp eax, 8
    jne .dls_close_st_hdr
    cmp dword [rsp+20h], 53534C58h ; 'XLSS'
    jne .dls_close_st_hdr
    mov r14d, dword [rsp+24h]
    add rsp, 40h

.dls_read_st_rec:
    test r14d, r14d
    jz .dls_close_st_done
    sub rsp, 40h
    lea rcx, [rsp+20h]
    mov edx, 1
    mov r8d, 12
    mov r9, r15
    call [fread]
    cmp eax, 12
    jne .dls_close_st_err

    movzx r8d, word [rsp+20h]    ; row
    movzx r9d, byte [rsp+22h]   ; col
    movzx r10d, byte [rsp+23h]  ; flags
    mov eax, dword [rsp+24h]    ; clrBk
    mov edx, dword [rsp+28h]    ; clrFg
    add rsp, 40h

    cmp r8d, 10000
    jae .dls_next_st_rec
    cmp r9d, 32
    jae .dls_next_st_rec

    mov ecx, r8d
    shl ecx, 5
    add ecx, r9d
    shl rcx, 4
    add rcx, [pCellStyleTable]

    mov dword [rcx], eax
    mov dword [rcx+4], edx
    mov dword [rcx+8], r10d

.dls_next_st_rec:
    dec r14d
    jmp .dls_read_st_rec

.dls_close_st_err:
    add rsp, 40h
    jmp .dls_close_st_done
.dls_close_st_hdr:
    add rsp, 40h
.dls_close_st_done:
    mov rcx, r15
    call [fclose]

.dls_style_done:

    ; 4. Load companion formula file
    cmp byte [szCurFormulaPath], 0
    je .dls_formula_done

    lea rcx, [szCurFormulaPath]
    lea rdx, [mode_rb]
    call [fopen]
    test rax, rax
    jz .dls_formula_done
    mov r15, rax

    sub rsp, 40h
    lea rcx, [rsp+20h]
    mov edx, 1
    mov r8d, 8
    mov r9, r15
    call [fread]
    cmp eax, 8
    jne .dls_close_fm_hdr
    cmp dword [rsp+20h], 46534C58h ; 'XLSF'
    jne .dls_close_fm_hdr
    mov r14d, dword [rsp+24h]
    add rsp, 40h

.dls_read_fm_rec:
    test r14d, r14d
    jz .dls_close_fm_done
    sub rsp, 280h
    lea rcx, [rsp+20h]
    mov edx, 1
    mov r8d, 260
    mov r9, r15
    call [fread]
    cmp eax, 260
    jne .dls_close_fm_err

    movzx r8d, word [rsp+20h]    ; row
    movzx r9d, byte [rsp+22h]    ; col
    lea r10, [rsp+24h]           ; formula string

    mov ecx, r8d
    mov edx, r9d
    mov r8, r10
    call SetCellFormula

    add rsp, 280h
    dec r14d
    jmp .dls_read_fm_rec

.dls_close_fm_err:
    add rsp, 280h
    jmp .dls_close_fm_done
.dls_close_fm_hdr:
    add rsp, 40h
.dls_close_fm_done:
    mov rcx, r15
    call [fclose]

.dls_formula_done:

    ; 5. Read CSV rows into ListView
    mov rcx, [hGrid]
    mov edx, LVM_DELETEALLITEMS
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]

    xor r13d, r13d
.dls_read_csv_loop:
    lea rcx, [szLineBuf]
    mov edx, 2040
    mov r8, r12
    call [fgets]
    test rax, rax
    jz .dls_read_csv_eof

    ; Strip newline
    lea rsi, [szLineBuf]
.dls_str_nl:
    lodsb
    test al, al
    jz .dls_str_nl_done
    cmp al, 10
    je .dls_str_nl_term
    cmp al, 13
    jne .dls_str_nl
.dls_str_nl_term:
    mov byte [rsi-1], 0
.dls_str_nl_done:

    ; Insert Row Number (col 0: 1-based "1", "2", ...)
    mov eax, r13d
    inc eax
    lea rcx, [szTemp]
    lea rdx, [fmt_d]
    mov r8d, eax
    call [sprintf]

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov dword [lvitem+8], 0
    lea rdi, [szTemp]
    mov qword [lvitem+24], rdi
    mov rcx, [hGrid]
    mov edx, LVM_INSERTITEMA
    xor r8d, r8d
    lea r9, [lvitem]
    call [SendMessageA]

    ; Split line by comma, semicolon or tab (columns 1..26)
    lea rsi, [szLineBuf]
    mov ebx, 1
.dls_next_tok:
    cmp ebx, 26
    jg .dls_tok_line_done
    cmp byte [rsi], 0
    je .dls_tok_line_done

    mov rdi, rsi
    mov byte [rsp+10h], 0 ; in_quote = 0
.dls_scan_cm:
    mov al, byte [rsi]
    test al, al
    jz .dls_tok_end
    cmp al, '"'
    jne .dls_not_q
    xor byte [rsp+10h], 1
    inc rsi
    jmp .dls_scan_cm
.dls_not_q:
    cmp byte [rsp+10h], 1
    je .dls_in_q
    cmp al, ','
    je .dls_found_cm
    cmp al, ';'
    je .dls_found_cm
    cmp al, 9
    je .dls_found_cm
.dls_in_q:
    inc rsi
    jmp .dls_scan_cm

.dls_found_cm:
    mov byte [rsi], 0
    inc rsi
.dls_tok_end:
    cmp byte [rdi], '"'
    jne .dls_check_tr
    inc rdi
.dls_check_tr:
    mov rax, rdi
.dls_find_t_end:
    cmp byte [rax], 0
    je .dls_found_t_end
    inc rax
    jmp .dls_find_t_end
.dls_found_t_end:
    cmp rax, rdi
    jle .dls_do_st_text
    cmp byte [rax-1], '"'
    jne .dls_do_st_text
    mov byte [rax-1], 0

.dls_do_st_text:
    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov dword [lvitem+8], ebx
    mov qword [lvitem+24], rdi
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, r13d
    lea r9, [lvitem]
    call [SendMessageA]

    inc ebx
    jmp .dls_next_tok

.dls_tok_line_done:
    inc r13d
    cmp r13d, 10000
    jge .dls_read_csv_eof
    jmp .dls_read_csv_loop

.dls_read_csv_eof:
    mov rcx, r12
    call [fclose]

    mov [total_rows], r13d
    call UpdateStatusStats

    ; Auto-fit columns 0..26 with minimum 85px
    mov rcx, [hGrid]
    mov edx, 101Eh ; LVM_SETCOLUMNWIDTH
    mov r8d, 0
    mov r9d, -2 ; LVSCW_AUTOSIZE_USEHEADER
    call [SendMessageA]

    mov ebx, 1
.dls_as_loop:
    cmp ebx, 26
    jg .dls_as_done
    mov rcx, [hGrid]
    mov edx, 101Eh
    mov r8d, ebx
    mov r9d, -1 ; LVSCW_AUTOSIZE
    call [SendMessageA]

    mov rcx, [hGrid]
    mov edx, 101Dh
    mov r8d, ebx
    xor r9d, r9d
    call [SendMessageA]
    cmp eax, 85
    jge @f
    mov rcx, [hGrid]
    mov edx, 101Eh
    mov r8d, ebx
    mov r9d, 85
    call [SendMessageA]
@@:
    inc ebx
    jmp .dls_as_loop
.dls_as_done:

    ; Reset selection to A1
    mov dword [cur_row], 0
    mov dword [cur_col], 1
    mov dword [sel_start_row], 0
    mov dword [sel_start_col], 1
    mov dword [sel_end_row], 0
    mov dword [sel_end_col], 1

    call UpdateCellSelectionDisplay
    call SyncListViewItemState

    ; Redraw grid
    mov rcx, [hGrid]
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]

.dls_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    mov r14, [rbp-48]
    mov r15, [rbp-56]
    leave
    ret

DoFileOpenDirect:
    push rbp
    mov rbp, rsp
    sub rsp, 90h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13
    mov [rbp-48], r14
    mov [rbp-56], r15

    ; Check extension in szFileName for .xls / .xlsx
    lea rsi, [szFileName]
    xor edx, edx
.scan_ext:
    lodsb
    test al, al
    jz .check_ext_done
    cmp al, '.'
    jne .scan_ext
    mov rdx, rsi ; rdx points after '.'
    jmp .scan_ext
.check_ext_done:
    test rdx, rdx
    jz .check_magic
    mov al, byte [rdx]
    or al, 20h
    cmp al, 'x'
    jne .check_magic
    mov al, byte [rdx+1]
    or al, 20h
    cmp al, 'l'
    jne .check_magic
    mov al, byte [rdx+2]
    or al, 20h
    cmp al, 's'
    je .is_excel_convert_no_close

.check_magic:
    ; Open in binary mode first to check magic
    lea rcx, [szFileName]
    lea rdx, [mode_rb]
    call [fopen]
    test rax, rax
    jz .direct_done
    mov r12, rax

    ; Read first 4 bytes
    lea rcx, [szTemp]
    mov edx, 1
    mov r8d, 4
    mov r9, r12
    call [fread]
    cmp eax, 4
    jl .not_excel_magic

    ; Check .xlsx (PK\x03\x04 -> 04034B50h)
    cmp dword [szTemp], 04034B50h
    je .is_excel_convert

    ; Check .xls (D0 CF 11 E0 -> 0E011CFD0h)
    cmp dword [szTemp], 0E011CFD0h
    je .is_excel_convert

.not_excel_magic:
    ; Rewind / close for normal CSV/TSV
    mov rcx, r12
    call [fclose]
    mov byte [is_temp_file], 0
    call SetupSheetTabsFromManifest
    xor ecx, ecx
    call DoLoadSheetByIndex
    jmp .finish_open_ui

.is_excel_convert:
    mov rcx, r12
    call [fclose]

.is_excel_convert_no_close:
    call CleanupTempFiles

    ; Get Temp Path
    mov ecx, 500
    lea rdx, [szTemp]
    call [GetTempPathA]

    ; Form szTempCsv: "%sxlsx_view_tmp.csv"
    lea rcx, [szTempCsv]
    lea rdx, [fmt_temp_csv]
    lea r8, [szTemp]
    call [sprintf]

    ; Form szTempStyle: "%sxlsx_view_tmp.csv.style"
    lea rcx, [szTempStyle]
    lea rdx, [fmt_temp_style]
    lea r8, [szTempCsv]
    call [sprintf]

    ; Load xlsx_native.dll dynamically if not already loaded
    mov rax, [pfnConvertXlsx]
    test rax, rax
    jnz .run_native_convert

    lea rcx, [szNativeDllName]
    call [LoadLibraryA]
    test rax, rax
    jnz .dll_loaded
    lea rcx, [szNativeDllPath]
    call [LoadLibraryA]
    test rax, rax
    jz .direct_done
.dll_loaded:
    mov [hNativeDll], rax
    mov rcx, rax
    lea rdx, [szNativeFuncName]
    call [GetProcAddress]
    test rax, rax
    jz .direct_done
    mov [pfnConvertXlsx], rax

.run_native_convert:
    ; Call native ConvertXlsxToCsvAndStyle(szFileName, szTempCsv, szTempStyle)
    lea rcx, [szFileName]
    lea rdx, [szTempCsv]
    lea r8, [szTempStyle]
    call [pfnConvertXlsx]

    mov byte [is_temp_file], 1
    call SetupSheetTabsFromManifest
    xor ecx, ecx
    call DoLoadSheetByIndex
    jmp .finish_open_ui

.open_temp_csv:

    ; Open the converted temp CSV
    lea rcx, [szTempCsv]
    lea rdx, [mode_r]
    call [fopen]
    test rax, rax
    jz .direct_done
    mov r12, rax
    mov byte [is_temp_file], 1

.start_read_grid:
    ; Clear existing styles table
    mov rdi, [pCellStyleTable]
    test rdi, rdi
    jz .skip_clear_styles
    mov rcx, 10000 * 32 * 2 ; (640,000 qwords)
    xor eax, eax
    rep stosq
.skip_clear_styles:

    ; Clear existing formulas table & heap
    mov rdi, [pCellFormulaTable]
    test rdi, rdi
    jz .skip_clear_formulas
    mov rcx, 10000 * 32 ; (320,000 qwords)
    xor eax, eax
    rep stosq

    mov rcx, [hFormulaHeap]
    test rcx, rcx
    jz .recreate_heap
    call [HeapDestroy]
.recreate_heap:
    xor ecx, ecx
    xor edx, edx
    xor r8d, r8d
    call [HeapCreate]
    mov [hFormulaHeap], rax
.skip_clear_formulas:

    ; If is_temp_file, load companion .style file
    cmp byte [is_temp_file], 1
    jne .skip_load_style

    lea rcx, [szTempStyle]
    lea rdx, [fmt_temp_style]
    lea r8, [szTempCsv]
    call [sprintf]

    lea rcx, [szTempStyle]
    lea rdx, [mode_rb]
    call [fopen]
    test rax, rax
    jz .skip_load_style
    mov r15, rax ; r15 = style file handle

    ; Read header (8 bytes: 4 bytes magic 'XLSS', 4 bytes count)
    sub rsp, 40h
    lea rcx, [rsp+20h]
    mov edx, 1
    mov r8d, 8
    mov r9, r15
    call [fread]
    cmp eax, 8
    jne .close_style_hdr
    cmp dword [rsp+20h], 53534C58h ; 'XLSS'
    jne .close_style_hdr
    mov r14d, dword [rsp+24h] ; r14d = record count
    add rsp, 40h

.read_record_loop:
    test r14d, r14d
    jz .close_style_done
    sub rsp, 40h
    lea rcx, [rsp+20h]
    mov edx, 1
    mov r8d, 12
    mov r9, r15
    call [fread]
    cmp eax, 12
    jne .close_style_err

    movzx r8d, word [rsp+20h]    ; row
    movzx r9d, byte [rsp+22h]   ; col
    movzx r10d, byte [rsp+23h]  ; flags
    mov eax, dword [rsp+24h]    ; clrBk
    mov edx, dword [rsp+28h]    ; clrFg
    add rsp, 40h

    cmp r8d, 10000
    jae .next_rec
    cmp r9d, 32
    jae .next_rec

    ; edx = (row * 32 + col) * 16
    mov ecx, r8d
    shl ecx, 5
    add ecx, r9d
    shl rcx, 4
    add rcx, [pCellStyleTable]

    mov dword [rcx], eax    ; clrBk
    mov dword [rcx+4], edx  ; clrFg
    mov dword [rcx+8], r10d ; flags

.next_rec:
    dec r14d
    jmp .read_record_loop

.close_style_err:
    add rsp, 40h
    jmp .close_style_done

.close_style_hdr:
    add rsp, 40h
.close_style_done:
    mov rcx, r15
    call [fclose]
    lea rcx, [szTempStyle]
    call [remove]

.skip_load_style:

    ; If is_temp_file, load companion .formula file
    cmp byte [is_temp_file], 1
    jne .skip_load_formula

    lea rcx, [szTempFormula]
    lea rdx, [fmt_temp_formula]
    lea r8, [szTempCsv]
    call [sprintf]

    lea rcx, [szTempFormula]
    lea rdx, [mode_rb]
    call [fopen]
    test rax, rax
    jz .skip_load_formula
    mov r15, rax ; r15 = formula file handle

    ; Read header (8 bytes: 4 bytes magic 'XLSF', 4 bytes count)
    sub rsp, 40h
    lea rcx, [rsp+20h]
    mov edx, 1
    mov r8d, 8
    mov r9, r15
    call [fread]
    cmp eax, 8
    jne .close_formula_hdr
    cmp dword [rsp+20h], 46534C58h ; 'XLSF'
    jne .close_formula_hdr
    mov r14d, dword [rsp+24h]      ; formula count
    add rsp, 40h

.read_formula_loop:
    test r14d, r14d
    jz .close_formula_done
    sub rsp, 280h
    lea rcx, [rsp+20h]
    mov edx, 1
    mov r8d, 260
    mov r9, r15
    call [fread]
    cmp eax, 260
    jne .close_formula_err

    movzx r8d, word [rsp+20h]    ; row
    movzx r9d, byte [rsp+22h]    ; col
    lea r10, [rsp+24h]           ; formula string

    mov ecx, r8d
    mov edx, r9d
    mov r8, r10
    call SetCellFormula

    add rsp, 280h
    dec r14d
    jmp .read_formula_loop

.close_formula_err:
    add rsp, 280h
    jmp .close_formula_done

.close_formula_hdr:
    add rsp, 40h
.close_formula_done:
    mov rcx, r15
    call [fclose]
    lea rcx, [szTempFormula]
    call [remove]

.skip_load_formula:
    mov rcx, [hGrid]
    mov edx, LVM_DELETEALLITEMS
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]

    xor r13d, r13d
.read_loop:
    lea rcx, [szLineBuf]
    mov edx, 2040
    mov r8, r12
    call [fgets]
    test rax, rax
    jz .eof_reached

    ; Strip newline
    lea rsi, [szLineBuf]
.strip_nl:
    lodsb
    test al, al
    jz .nl_stripped
    cmp al, 13
    je .terminate_line
    cmp al, 10
    je .terminate_line
    jmp .strip_nl
.terminate_line:
    mov byte [rsi-1], 0
.nl_stripped:

    ; Insert Row
    mov eax, r13d
    inc eax
    lea rdi, [szTemp+8]
    mov byte [rdi], 0
    mov ecx, 10
.itos_open:
    xor edx, edx
    div ecx
    add dl, '0'
    dec rdi
    mov [rdi], dl
    test eax, eax
    jnz .itos_open

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov dword [lvitem+8], 0
    mov qword [lvitem+24], rdi
    mov rcx, [hGrid]
    mov edx, LVM_INSERTITEMA
    xor r8d, r8d
    lea r9, [lvitem]
    call [SendMessageA]

    ; Split line by comma, semicolon or tab (columns 1..26)
    lea rsi, [szLineBuf]
    mov ebx, 1
.next_token:
    cmp ebx, 26
    jg .line_done
    cmp byte [rsi], 0
    je .line_done

    mov rdi, rsi
    mov byte [rsp+10h], 0 ; in_quote = 0
.scan_comma:
    mov al, byte [rsi]
    test al, al
    jz .token_end
    cmp al, '"'
    jne .not_quote
    xor byte [rsp+10h], 1 ; toggle in_quote flag
    inc rsi
    jmp .scan_comma
.not_quote:
    cmp byte [rsp+10h], 1
    je .char_in_quote
    cmp al, ','
    je .found_comma
    cmp al, ';'
    je .found_comma
    cmp al, 9
    je .found_comma
.char_in_quote:
    inc rsi
    jmp .scan_comma

.found_comma:
    mov byte [rsi], 0
    inc rsi
.token_end:
    ; Trim leading quote if present
    cmp byte [rdi], '"'
    jne .check_trail
    inc rdi
.check_trail:
    mov rax, rdi
.find_tok_end:
    cmp byte [rax], 0
    je .found_tok_end
    inc rax
    jmp .find_tok_end
.found_tok_end:
    cmp rax, rdi
    jle .do_set_text
    cmp byte [rax-1], '"'
    jne .do_set_text
    mov byte [rax-1], 0

.do_set_text:
    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov dword [lvitem+8], ebx
    mov qword [lvitem+24], rdi
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, r13d
    lea r9, [lvitem]
    call [SendMessageA]

    inc ebx
    jmp .next_token

.line_done:
    inc r13d
    cmp r13d, 10000
    jge .eof_reached
    jmp .read_loop

.eof_reached:
    mov rcx, r12
    call [fclose]

    cmp byte [is_temp_file], 1
    jne @f
    lea rcx, [szTempCsv]
    call [remove]
    lea rcx, [szTempStyle]
    call [remove]
@@:
    mov [total_rows], r13d
    call SaveActiveSheetToBuffer
    call UpdateStatusStats

    ; Auto-fit columns 0..26 with minimum 85px so cell text is never truncated
    mov rcx, [hGrid]
    mov edx, 101Eh ; LVM_SETCOLUMNWIDTH
    mov r8d, 0
    mov r9d, -2 ; LVSCW_AUTOSIZE_USEHEADER
    call [SendMessageA]

    mov ebx, 1
.autosize_loop:
    cmp ebx, 26
    jg .autosize_done
    mov rcx, [hGrid]
    mov edx, 101Eh ; LVM_SETCOLUMNWIDTH
    mov r8d, ebx
    mov r9d, -1 ; LVSCW_AUTOSIZE
    call [SendMessageA]

    mov rcx, [hGrid]
    mov edx, 101Dh ; LVM_GETCOLUMNWIDTH
    mov r8d, ebx
    xor r9d, r9d
    call [SendMessageA]
    cmp eax, 85
    jge @f
    mov rcx, [hGrid]
    mov edx, 101Eh ; LVM_SETCOLUMNWIDTH
    mov r8d, ebx
    mov r9d, 85
    call [SendMessageA]
@@:
    inc ebx
    jmp .autosize_loop
.autosize_done:

    ; Find basename for Window Title and Sheet Tab
    lea rsi, [szFileName]
    mov r8, rsi
.find_fn_base:
    lodsb
    test al, al
    jz .got_fn_base
    cmp al, '\'
    cmove r8, rsi
    cmp al, '/'
    cmove r8, rsi
    jmp .find_fn_base
.got_fn_base:
    ; Update first tab text to loaded filename
    lea rdi, [tcitem]
    xor eax, eax
    mov ecx, 48
    rep stosb
    mov dword [tcitem], TCIF_TEXT
    mov qword [tcitem+16], r8
    mov rcx, [hSheetTabs]
    mov edx, 1306h ; TCM_SETITEMA
    xor r8d, r8d
    lea r9, [tcitem]
    call [SendMessageA]

.finish_open_ui:
    ; Update Window Title to: "Excel Viewer Pro - [filename] [x64 FASM Native Edition]"
    lea rsi, [szFileName]
    mov r8, rsi
.find_fn_base2:
    lodsb
    test al, al
    jz .got_fn_base2
    cmp al, '\'
    cmove r8, rsi
    cmp al, '/'
    cmove r8, rsi
    jmp .find_fn_base2
.got_fn_base2:
    lea rcx, [szTemp]
    lea rdx, [fmt_title_opened]
    call [sprintf]

    mov rcx, [hWnd]
    lea rdx, [szTemp]
    call [SetWindowTextA]

    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [szFileName]
    call [SendMessageA]

    ; Reset active cell and selection to A1
    mov dword [cur_row], 0
    mov dword [cur_col], 1
    mov dword [sel_start_row], 0
    mov dword [sel_start_col], 1
    mov dword [sel_end_row], 0
    mov dword [sel_end_col], 1

    call UpdateCellSelectionDisplay
    call SyncListViewItemState

    ; Ensure keyboard focus is on grid so navigation keys work immediately
    mov rcx, [hGrid]
    call [SetFocus]

    ; Force full redraw of grid
    mov rcx, [hGrid]
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]

.direct_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    mov r14, [rbp-48]
    mov r15, [rbp-56]
    leave
    ret

.open_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    mov r14, [rbp-48]
    mov r15, [rbp-56]
    leave
    ret

DoFileSaveAs:
    push rbp
    mov rbp, rsp
    sub rsp, 80h

    lea rdi, [ofn]
    xor eax, eax
    mov ecx, 152
    rep stosb

    mov dword [ofn], 152
    mov rax, [hWnd]
    mov qword [ofn+8], rax
    lea rax, [csv_filter]
    mov qword [ofn+24], rax
    lea rax, [szFileName]
    mov qword [ofn+48], rax
    mov dword [ofn+56], 512
    lea rax, [ofn_title_save]
    mov qword [ofn+88], rax
    mov dword [ofn+96], 0000080Ah ; OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR
    lea rax, [def_ext_csv]
    mov qword [ofn+104], rax

    lea rcx, [ofn]
    call [GetSaveFileNameA]
    test eax, eax
    jz .save_done

    call DoFileSaveInternal

.save_done:
    leave
    ret

DoFileSave:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    cmp byte [szFileName], 0
    je .ask_name
    call DoFileSaveInternal
    leave
    ret
.ask_name:
    leave
    jmp DoFileSaveAs

DoFileSaveInternal:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    lea rcx, [szFileName]
    lea rdx, [mode_w]
    call [fopen]
    test rax, rax
    jz .err_save
    mov r12, rax

    mov r13d, 0
.save_row:
    cmp r13d, [total_rows]
    jge .save_close

    mov ebx, 1
.save_col:
    cmp ebx, 26
    jg .row_nl

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov dword [lvitem+8], ebx
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 500
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r13d
    lea r9, [lvitem]
    call [SendMessageA]

    mov rcx, r12
    lea rdx, [fmt_s]
    lea r8, [szCellVal]
    call [fprintf]

    cmp ebx, 26
    je .skip_comma
    mov rcx, r12
    lea rdx, [fmt_comma_char]
    call [fprintf]
.skip_comma:
    inc ebx
    jmp .save_col

.row_nl:
    mov rcx, r12
    lea rdx, [fmt_newline]
    call [fprintf]

    inc r13d
    jmp .save_row

.save_close:
    mov rcx, r12
    call [fclose]

    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [saved_msg]
    call [SendMessageA]

.err_save:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret

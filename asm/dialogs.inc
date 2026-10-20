; ============================================================================
; Excel Viewer Pro — Dialogs, Helpers & FX Wizard
; ============================================================================


DoConditionalFormat:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    ; Toggle color styles
    call UpdateStatusStats
    leave
    ret


; =============================================================================
; Action: Run 1,000,000 Numbers SIMD Benchmark
; =============================================================================
DoRunBenchmark:
    push rbp
    mov rbp, rsp
    sub rsp, 60h

    lea rdi, [bench_array]
    mov ecx, 1000000
    xor edx, edx
    mov rax, 03FD0000000000000h ; 0.25
    movq xmm1, rax

.fill_loop:
    mov eax, edx
    xor edx, edx
    mov r8d, 100
    div r8d
    cvtsi2sd xmm0, edx
    mulsd xmm0, xmm1
    movsd [rdi], xmm0
    add rdi, 8
    mov edx, eax
    inc edx
    dec ecx
    jnz .fill_loop

    lea rcx, [bench_freq]
    call [QueryPerformanceFrequency]
    lea rcx, [bench_start]
    call [QueryPerformanceCounter]

    lea rcx, [bench_array]
    mov rdx, 1000000
    call asm_vec_sum
    movsd [rsp+30h], xmm0   ; SIMD sum

    lea rcx, [bench_end]
    call [QueryPerformanceCounter]

    mov rax, [bench_end]
    sub rax, [bench_start]
    cvtsi2sd xmm1, rax
    cvtsi2sd xmm2, [bench_freq]
    divsd xmm1, xmm2
    mov rax, 0408F400000000000h ; 1000.0
    movq xmm3, rax
    mulsd xmm1, xmm3
    movsd [rsp+38h], xmm1   ; Time ms

    mov rax, 0408F400000000000h
    movq xmm4, rax
    divsd xmm4, [rsp+38h]
    movsd [rsp+40h], xmm4   ; M ops/sec

    lea rcx, [szBuffer]
    lea rdx, [bench_fmt]
    mov r8, [rsp+30h]
    mov r9, [rsp+38h]
    mov rax, [rsp+40h]
    mov [rsp+20h], rax
    call [sprintf]

    mov rcx, [hWnd]
    lea rdx, [szBuffer]
    lea r8, [bench_title]
    xor r9d, r9d
    call [MessageBoxA]

    leave
    ret


; =============================================================================
; Action: Find Text in Worksheet
; =============================================================================
DoFindText:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    cmp byte [szFindQuery], 0
    jne .have_fq
    ; Default query: "Revenue"
    mov dword [szFindQuery], 'Reve'
    mov dword [szFindQuery+4], 'nue'
    mov byte [szFindQuery+7], 0

.have_fq:
    mov r12d, 0
.f_row:
    cmp r12d, [total_rows]
    jge .f_not_found

    mov r13d, 1
.f_col:
    cmp r13d, 10
    jg .f_next_row

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r12d
    mov dword [lvitem+8], r13d
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 500
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r12d
    lea r9, [lvitem]
    call [SendMessageA]
    test eax, eax
    jz .f_next_col

    ; Simple substring match
    lea rsi, [szCellVal]
    lea rdi, [szFindQuery]
    call SubstrMatch
    test eax, eax
    jnz .f_match_found

.f_next_col:
    inc r13d
    jmp .f_col

.f_next_row:
    inc r12d
    jmp .f_row

.f_match_found:
    mov [cur_row], r12d
    mov [cur_col], r13d

    mov al, byte [cur_col]
    dec al
    add al, 'A'
    lea rdi, [szTemp]
    mov [rdi], al
    inc rdi

    mov eax, [cur_row]
    inc eax
    mov ecx, 10
    lea rsi, [szTemp+16]
    mov byte [rsi], 0
.itos_f:
    xor edx, edx
    div ecx
    add dl, '0'
    dec rsi
    mov [rsi], dl
    test eax, eax
    jnz .itos_f
.cpy_f_str:
    lodsb
    stosb
    test al, al
    jnz .cpy_f_str

    mov rcx, [hNameBox]
    lea rdx, [szTemp]
    call [SetWindowTextA]

    mov rcx, [hFormulaEdit]
    lea rdx, [szCellVal]
    call [SetWindowTextA]

    ; Show found message in status bar
    lea rcx, [szBuffer]
    lea rdx, [find_found_fmt]
    mov al, byte [cur_col]
    dec al
    add al, 'A'
    movzx r8d, al
    mov r9d, [cur_row]
    inc r9d
    call [sprintf]

    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [szBuffer]
    call [SendMessageA]
    jmp .f_done

.f_not_found:
    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [find_not_fnd]
    call [SendMessageA]

.f_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret

; SubstrMatch: RSI = haystack, RDI = needle
; Returns EAX = 1 if needle in haystack, 0 if not
SubstrMatch:
    push rbx
    push rsi
    push rdi
.sm_outer:
    cmp byte [rsi], 0
    je .sm_fail
    mov rbx, rsi
    mov rdx, rdi
.sm_inner:
    mov al, byte [rdx]
    test al, al
    jz .sm_success
    mov cl, byte [rbx]
    test cl, cl
    jz .sm_fail
    and al, 0DFh
    and cl, 0DFh
    cmp al, cl
    jne .sm_advance
    inc rbx
    inc rdx
    jmp .sm_inner
.sm_advance:
    inc rsi
    jmp .sm_outer
.sm_success:
    mov eax, 1
    jmp .sm_ret
.sm_fail:
    xor eax, eax
.sm_ret:
    pop rdi
    pop rsi
    pop rbx
    ret

DoPmtDialog:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    mov rcx, [hWnd]
    lea rdx, [pmt_calc_msg]
    lea r8, [pmt_calc_title]
    xor r9d, r9d
    call [MessageBoxA]
    leave
    ret

DoShowFormulaRef:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    mov rcx, [hWnd]
    lea rdx, [form_ref_msg]
    lea r8, [form_ref_title]
    xor r9d, r9d
    call [MessageBoxA]
    leave
    ret

DoShowSimdInfo:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    mov rcx, [hWnd]
    lea rdx, [simd_info_msg]
    lea r8, [simd_info_title]
    xor r9d, r9d
    call [MessageBoxA]
    leave
    ret

DoInsertFxPrompt:
DoShowFxWizard:
    push rbp
    mov rbp, rsp
    sub rsp, 80h

    ; If wizard already open, bring to front
    cmp qword [hFxWizWnd], 0
    je .register_or_create
    mov rcx, [hFxWizWnd]
    call [SetFocus]
    jmp .wiz_done

.register_or_create:
    cmp dword [wiz_class_reg], 1
    je .create_wiz_window

    ; Register FxWizardClass
    lea rdi, [wc_wiz]
    xor eax, eax
    mov ecx, 80
    rep stosb

    mov dword [wc_wiz], 80
    mov dword [wc_wiz+4], CS_HREDRAW or CS_VREDRAW
    lea rax, [FxWizardWndProc]
    mov qword [wc_wiz+8], rax
    mov rax, [hInst]
    mov qword [wc_wiz+24], rax

    mov rax, [hAppIcon]
    mov qword [wc_wiz+32], rax
    mov qword [wc_wiz+72], rax

    xor ecx, ecx
    mov edx, 32512 ; IDC_ARROW
    call [LoadCursorA]
    mov qword [wc_wiz+40], rax
    mov qword [wc_wiz+48], COLOR_BTNFACE + 1
    lea rax, [fx_wiz_class_name]
    mov qword [wc_wiz+64], rax

    lea rcx, [wc_wiz]
    call [RegisterClassExA]
    mov dword [wiz_class_reg], 1

.create_wiz_window:
    ; Disable main window (modal behavior)
    mov rcx, [hWnd]
    xor edx, edx ; FALSE
    call [EnableWindow]

    ; Create Wizard Dialog (x=240, y=140, w=520, h=485)
    sub rsp, 60h
    mov ecx, 00000001h ; WS_EX_DLGMODALFRAME
    lea rdx, [fx_wiz_class_name]
    lea r8, [fx_wiz_title]
    mov r9d, WS_POPUP or WS_CAPTION or WS_SYSMENU or WS_VISIBLE
    mov dword [rsp+20h], 240
    mov dword [rsp+28h], 140
    mov dword [rsp+30h], 520
    mov dword [rsp+38h], 485
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], 0
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    add rsp, 60h
    mov [hFxWizWnd], rax

    ; Set wizard window icons
    mov rcx, [hFxWizWnd]
    mov edx, 0080h ; WM_SETICON
    mov r8d, 1     ; ICON_BIG
    mov r9, [hAppIcon]
    call [SendMessageA]

    mov rcx, [hFxWizWnd]
    mov edx, 0080h ; WM_SETICON
    xor r8d, r8d   ; ICON_SMALL
    mov r9, [hAppIcon]
    call [SendMessageA]

    mov rcx, [hFxWizWnd]
    mov edx, 1 ; SW_SHOWNORMAL
    call [ShowWindow]
    mov rcx, [hFxWizWnd]
    call [UpdateWindow]

.wiz_done:
    leave
    ret


; =============================================================================
; RepopulateWizList: Filter 129 functions by Category and Search Query
; =============================================================================
RepopulateWizList:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13
    mov [rbp-48], r14

    ; Clear listbox
    mov rcx, [hWizList]
    test rcx, rcx
    jz .rep_ret
    mov edx, LB_RESETCONTENT
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]

    ; Get selected category
    mov rcx, [hWizCatCombo]
    test rcx, rcx
    jz .def_cat
    mov edx, CB_GETCURSEL
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]
    cmp eax, 0
    jge .cat_valid
.def_cat:
    xor eax, eax
.cat_valid:
    mov r14d, eax ; r14d = selected cat (0 = All)

    ; Get search query
    mov rcx, [hWizSearchEdit]
    test rcx, rcx
    jz .empty_query
    lea rdx, [szWizQuery]
    mov r8d, 120
    call [GetWindowTextA]
    jmp .have_query
.empty_query:
    mov byte [szWizQuery], 0
.have_query:

    mov dword [num_filtered_fn], 0
    xor ebx, ebx

.filter_loop:
    cmp ebx, NUM_WIZ_FUNCS
    jge .filter_done

    ; If r14d != 0, check category
    test r14d, r14d
    jz .check_query
    mov eax, [fn_all_cats + rbx*4]
    cmp eax, r14d
    jne .next_fn

.check_query:
    cmp byte [szWizQuery], 0
    je .matches

    ; Substring match on function name
    mov rsi, [fn_all_names + rbx*8]
    lea rdi, [szWizQuery]
    call SubstrMatch
    test eax, eax
    jnz .matches

    ; Substring match on description
    mov rsi, [fn_all_desc + rbx*8]
    lea rdi, [szWizQuery]
    call SubstrMatch
    test eax, eax
    jz .next_fn

.matches:
    ; Add to listbox
    mov rcx, [hWizList]
    mov edx, LB_ADDSTRING
    xor r8d, r8d
    mov r9, [fn_all_names + rbx*8]
    call [SendMessageA]

    ; Store function index in filtered_fn_idx
    mov eax, [num_filtered_fn]
    cmp eax, 255
    jge .next_fn
    mov [filtered_fn_idx + rax*4], ebx
    inc dword [num_filtered_fn]

.next_fn:
    inc ebx
    jmp .filter_loop

.filter_done:
    cmp dword [num_filtered_fn], 0
    jle .no_matches

    ; Select first item
    mov rcx, [hWizList]
    mov edx, LB_SETCURSEL
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]

    ; Show syntax & desc of first item
    mov eax, [filtered_fn_idx]
    mov rcx, [hWizSyntax]
    mov rdx, [fn_all_syntax + rax*8]
    call [SetWindowTextA]

    mov eax, [filtered_fn_idx]
    mov rcx, [hWizDesc]
    mov rdx, [fn_all_desc + rax*8]
    call [SetWindowTextA]
    jmp .rep_ret

.no_matches:
    mov rcx, [hWizSyntax]
    lea rdx, [empty_str]
    call [SetWindowTextA]

    mov rcx, [hWizDesc]
    lea rdx, [lbl_wiz_no_match]
    call [SetWindowTextA]

.rep_ret:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    mov r14, [rbp-48]
    leave
    ret


; =============================================================================
; FxWizardWndProc: Window Procedure for Function Wizard Dialog
; =============================================================================
FxWizardWndProc:
    push rbp
    mov rbp, rsp
    sub rsp, 70h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    cmp edx, WM_CREATE
    je .on_create
    cmp edx, WM_COMMAND
    je .on_command
    cmp edx, WM_CLOSE
    je .on_close
    cmp edx, WM_DESTROY
    je .on_destroy

    call [DefWindowProcA]
    jmp .ret_w

.on_create:
    mov r12, rcx ; hWndWizard

    ; 1. Static Label "Search for a function:" (x=16, y=12, w=480, h=18)
    sub rsp, 60h
    xor ecx, ecx
    lea rdx, [static_class_name]
    lea r8, [lbl_wiz_search]
    mov r9d, WS_CHILD or WS_VISIBLE
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 12
    mov dword [rsp+30h], 480
    mov dword [rsp+38h], 18
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], 0
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    add rsp, 60h
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    ; 2. Edit Control hWizSearchEdit (x=16, y=32, w=416, h=24)
    sub rsp, 60h
    mov ecx, WS_EX_CLIENTEDGE
    lea rdx, [edit_class]
    mov r8, 0
    mov r9d, WS_CHILD or WS_VISIBLE or WS_BORDER or ES_AUTOHSCROLL
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 32
    mov dword [rsp+30h], 416
    mov dword [rsp+38h], 24
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], ID_WIZ_SEARCH
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    add rsp, 60h
    mov [hWizSearchEdit], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    ; 3. Button "Go" (x=438, y=31, w=58, h=26)
    sub rsp, 60h
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_wiz_go_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_PUSHBUTTON
    mov dword [rsp+20h], 438
    mov dword [rsp+28h], 31
    mov dword [rsp+30h], 58
    mov dword [rsp+38h], 26
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], ID_WIZ_GO
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    add rsp, 60h
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    ; 4. Static Label "Or select a category:" (x=16, y=66, w=146, h=20)
    sub rsp, 60h
    xor ecx, ecx
    lea rdx, [static_class_name]
    lea r8, [lbl_wiz_cat]
    mov r9d, WS_CHILD or WS_VISIBLE
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 66
    mov dword [rsp+30h], 146
    mov dword [rsp+38h], 20
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], 0
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    add rsp, 60h
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    ; 5. ComboBox Category (x=166, y=64, w=330, h=240)
    sub rsp, 60h
    xor ecx, ecx
    lea rdx, [combo_class]
    mov r8, 0
    mov r9d, WS_CHILD or WS_VISIBLE or CBS_DROPDOWNLIST or WS_VSCROLL
    mov dword [rsp+20h], 166
    mov dword [rsp+28h], 64
    mov dword [rsp+30h], 330
    mov dword [rsp+38h], 240
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], ID_WIZ_CAT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    add rsp, 60h
    mov [hWizCatCombo], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    ; Populate 9 categories
    xor ebx, ebx
.populate_cats:
    mov rcx, [hWizCatCombo]
    mov edx, CB_ADDSTRING
    xor r8d, r8d
    mov r9, [cat_names_table + rbx*8]
    call [SendMessageA]
    inc ebx
    cmp ebx, NUM_WIZ_CATS
    jl .populate_cats

    ; Set initial category from wiz_init_cat
    mov eax, [wiz_init_cat]
    cmp eax, NUM_WIZ_CATS
    jl .init_cat_ok
    xor eax, eax
.init_cat_ok:
    mov rcx, [hWizCatCombo]
    mov edx, CB_SETCURSEL
    mov r8d, eax
    xor r9d, r9d
    call [SendMessageA]

    ; 6. Static Label "Select a function:" (x=16, y=98, w=480, h=18)
    sub rsp, 60h
    xor ecx, ecx
    lea rdx, [static_class_name]
    lea r8, [lbl_wiz_prompt]
    mov r9d, WS_CHILD or WS_VISIBLE
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 98
    mov dword [rsp+30h], 480
    mov dword [rsp+38h], 18
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], 0
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    add rsp, 60h
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    ; 7. ListBox ID_WIZ_LIST (x=16, y=118, w=480, h=150)
    sub rsp, 60h
    mov ecx, WS_EX_CLIENTEDGE
    lea rdx, [listbox_class]
    mov r8, 0
    mov r9d, WS_CHILD or WS_VISIBLE or WS_BORDER or WS_VSCROLL or LBS_NOTIFY
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 118
    mov dword [rsp+30h], 480
    mov dword [rsp+38h], 150
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], ID_WIZ_LIST
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    add rsp, 60h
    mov [hWizList], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    ; 8. GroupBox "Function Details" (x=16, y=278, w=480, h=120)
    sub rsp, 60h
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [grp_wiz_details]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_GROUPBOX
    mov dword [rsp+20h], 16
    mov dword [rsp+28h], 278
    mov dword [rsp+30h], 480
    mov dword [rsp+38h], 120
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], 0
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    add rsp, 60h
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; 9. Static Control hWizSyntax (x=28, y=300, w=456, h=20)
    sub rsp, 60h
    xor ecx, ecx
    lea rdx, [static_class_name]
    lea r8, [empty_str]
    mov r9d, WS_CHILD or WS_VISIBLE
    mov dword [rsp+20h], 28
    mov dword [rsp+28h], 300
    mov dword [rsp+30h], 456
    mov dword [rsp+38h], 20
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], 0
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    add rsp, 60h
    mov [hWizSyntax], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; 10. Static Control hWizDesc (x=28, y=324, w=456, h=66)
    sub rsp, 60h
    xor ecx, ecx
    lea rdx, [static_class_name]
    lea r8, [empty_str]
    mov r9d, WS_CHILD or WS_VISIBLE
    mov dword [rsp+20h], 28
    mov dword [rsp+28h], 324
    mov dword [rsp+30h], 456
    mov dword [rsp+38h], 66
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], 0
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    add rsp, 60h
    mov [hWizDesc], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    ; 11. Button Cancel [Cancel] (x=286, y=410, w=100, h=28)
    sub rsp, 60h
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_wiz_can_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_PUSHBUTTON
    mov dword [rsp+20h], 286
    mov dword [rsp+28h], 410
    mov dword [rsp+30h], 100
    mov dword [rsp+38h], 28
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], ID_WIZ_CANCEL
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    add rsp, 60h
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    ; 12. Button OK [Insert] (x=396, y=410, w=100, h=28)
    sub rsp, 60h
    xor ecx, ecx
    lea rdx, [btn_class]
    lea r8, [btn_wiz_ok_txt]
    mov r9d, WS_CHILD or WS_VISIBLE or BS_PUSHBUTTON
    mov dword [rsp+20h], 396
    mov dword [rsp+28h], 410
    mov dword [rsp+30h], 100
    mov dword [rsp+38h], 28
    mov qword [rsp+40h], r12
    mov qword [rsp+48h], ID_WIZ_OK
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    add rsp, 60h
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontBold]
    mov r9d, 1
    call [SendMessageA]

    ; Populate list initially
    call RepopulateWizList

    xor eax, eax
    jmp .ret_w

.on_command:
    mov eax, r8d
    and eax, 0FFFFh ; LOWORD(wParam)
    shr r8d, 16     ; HIWORD(wParam)

    cmp eax, ID_WIZ_SEARCH
    je .cmd_wiz_search
    cmp eax, ID_WIZ_GO
    je .cmd_wiz_go
    cmp eax, ID_WIZ_CAT
    je .cmd_wiz_cat
    cmp eax, ID_WIZ_LIST
    je .cmd_wiz_list
    cmp eax, ID_WIZ_OK
    je .cmd_wiz_ok
    cmp eax, 1 ; IDOK
    je .cmd_wiz_ok
    cmp eax, ID_WIZ_CANCEL
    je .cmd_wiz_cancel
    cmp eax, 2 ; IDCANCEL
    je .cmd_wiz_cancel
    jmp .ret0_w

.cmd_wiz_search:
    cmp r8d, EN_CHANGE
    je .do_repop
    jmp .ret0_w

.cmd_wiz_go:
    jmp .do_repop

.cmd_wiz_cat:
    cmp r8d, CBN_SELCHANGE
    je .do_repop
    jmp .ret0_w

.do_repop:
    call RepopulateWizList
    jmp .ret0_w

.cmd_wiz_list:
    cmp r8d, LBN_SELCHANGE
    je .on_wiz_selchange
    cmp r8d, LBN_DBLCLK
    je .cmd_wiz_ok
    jmp .ret0_w

.on_wiz_selchange:
    mov rcx, [hWizList]
    mov edx, LB_GETCURSEL
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]
    cmp eax, 0
    jl .ret0_w
    cmp eax, [num_filtered_fn]
    jge .ret0_w

    mov edx, [filtered_fn_idx + rax*4]
    mov rcx, [hWizSyntax]
    mov rdx, [fn_all_syntax + rdx*8]
    call [SetWindowTextA]

    mov rcx, [hWizList]
    mov edx, LB_GETCURSEL
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]
    mov edx, [filtered_fn_idx + rax*4]
    mov rcx, [hWizDesc]
    mov rdx, [fn_all_desc + rdx*8]
    call [SetWindowTextA]
    jmp .ret0_w

.cmd_wiz_ok:
    mov rcx, [hWizList]
    mov edx, LB_GETCURSEL
    xor r8d, r8d
    xor r9d, r9d
    call [SendMessageA]
    cmp eax, 0
    jl .check_has_any
    cmp eax, [num_filtered_fn]
    jl .ok_idx_valid
.check_has_any:
    cmp dword [num_filtered_fn], 0
    jle .ret0_w
    xor eax, eax
.ok_idx_valid:
    mov edx, [filtered_fn_idx + rax*4]
    mov rsi, [fn_all_names + rdx*8]

    ; Format "={NAME}(" into szTemp
    lea rcx, [szTemp]
    lea rdx, [fmt_fn_prefix]
    mov r8, rsi
    call [sprintf]

    lea rcx, [szTemp]
    call InsertFormulaSnippet

    mov rcx, [hFxWizWnd]
    call [DestroyWindow]
    jmp .ret0_w

.cmd_wiz_cancel:
    mov rcx, [hFxWizWnd]
    call [DestroyWindow]
    jmp .ret0_w

.on_close:
    mov rcx, [hFxWizWnd]
    call [DestroyWindow]
    jmp .ret0_w

.on_destroy:
    mov rcx, [hWnd]
    mov edx, 1 ; TRUE
    call [EnableWindow]
    mov rcx, [hWnd]
    call [SetFocus]
    mov qword [hFxWizWnd], 0
    xor eax, eax
    jmp .ret_w

.ret0_w:
    xor eax, eax
.ret_w:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret

; =============================================================================
; Undo / Redo Subsystem
; =============================================================================

; BeginCellChange(ECX = row, EDX = col)
BeginCellChange:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi

    mov [pending_row], ecx
    mov [pending_col], edx
    mov eax, [cur_sheet]
    mov [pending_sheet], eax

    ; Read current cell text from grid
    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], ecx
    mov dword [lvitem+8], edx
    lea rax, [pending_old_txt]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 115
    mov byte [pending_old_txt], 0
    mov rcx, [hGrid]
    test rcx, rcx
    jz .bcc_no_grid
    mov edx, LVM_GETITEMTEXTA
    mov r8d, [pending_row]
    lea r9, [lvitem]
    call [SendMessageA]
.bcc_no_grid:

    ; Read current formula from formula table
    mov ecx, [pending_row]
    mov edx, [pending_col]
    call GetCellFormula
    test rax, rax
    jz .bcc_no_form

    ; Copy formula to pending_old_form
    cld
    mov rsi, rax
    lea rdi, [pending_old_form]
    mov ecx, 115
.cpy_form:
    lodsb
    stosb
    test al, al
    jz .bcc_form_done
    dec ecx
    jnz .cpy_form
    mov byte [rdi], 0
    jmp .bcc_form_done

.bcc_no_form:
    mov byte [pending_old_form], 0

.bcc_form_done:
    mov dword [pending_undo_act], 1

    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    leave
    ret


; CommitCellChange(ECX = row, EDX = col, R8 = pNewText, R9 = pNewFormula)
CommitCellChange:
    push rbp
    mov rbp, rsp
    sub rsp, 70h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13
    mov [rbp-48], r14

    cmp dword [pending_undo_act], 1
    jne .ccc_done
    mov dword [pending_undo_act], 0

    mov r12, r8     ; pNewText
    mov r13, r9     ; pNewFormula

    ; Check if undo stack is full
    mov eax, [undo_count]
    cmp eax, UNDO_MAX
    jl .have_slot

    ; Shift stack down by 1 record (memmove undo_stack+512 to undo_stack)
    cld
    lea rdi, [undo_stack]
    lea rsi, [undo_stack + UNDO_REC_SIZE]
    mov ecx, (UNDO_MAX - 1) * UNDO_REC_SIZE
    rep movsb
    mov dword [undo_count], UNDO_MAX - 1

.have_slot:
    mov eax, [undo_count]
    imul rax, UNDO_REC_SIZE
    lea rdi, [undo_stack + rax]
    mov r14, rdi ; r14 = pointer to current record

    ; Zero out record
    cld
    xor eax, eax
    mov ecx, UNDO_REC_SIZE
    rep stosb

    ; Fill header
    mov eax, [pending_row]
    mov [r14 + 0], eax
    mov eax, [pending_col]
    mov [r14 + 4], eax
    mov eax, [pending_sheet]
    mov [r14 + 8], eax
    mov dword [r14 + 12], 1 ; cell edit

    ; Copy pending_old_txt to old_text (+16)
    cld
    lea rsi, [pending_old_txt]
    lea rdi, [r14 + 16]
    mov ecx, 115
.cp_ot:
    lodsb
    stosb
    test al, al
    jz .ot_done
    dec ecx
    jnz .cp_ot
    mov byte [rdi], 0
.ot_done:

    ; Copy pending_old_form to old_formula (+136)
    cld
    lea rsi, [pending_old_form]
    lea rdi, [r14 + 136]
    mov ecx, 115
.cp_of:
    lodsb
    stosb
    test al, al
    jz .of_done
    dec ecx
    jnz .cp_of
    mov byte [rdi], 0
.of_done:

    ; Copy pNewText to new_text (+256)
    test r12, r12
    jz .nt_done
    cld
    mov rsi, r12
    lea rdi, [r14 + 256]
    mov ecx, 115
.cp_nt:
    lodsb
    stosb
    test al, al
    jz .nt_done
    dec ecx
    jnz .cp_nt
    mov byte [rdi], 0
.nt_done:

    ; Copy pNewFormula to new_formula (+376)
    test r13, r13
    jz .nf_done
    cld
    mov rsi, r13
    lea rdi, [r14 + 376]
    mov ecx, 115
.cp_nf:
    lodsb
    stosb
    test al, al
    jz .nf_done
    dec ecx
    jnz .cp_nf
    mov byte [rdi], 0
.nf_done:

    ; Increment undo_count
    inc dword [undo_count]
    ; Clear redo stack on new action
    mov dword [redo_count], 0

.ccc_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    mov r14, [rbp-48]
    leave
    ret


; =============================================================================
; Action: Undo
; =============================================================================
DoUndo:
    push rbp
    mov rbp, rsp
    sub rsp, 80h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    cmp dword [undo_count], 0
    jle .undo_empty

    ; Pop record from undo_stack
    dec dword [undo_count]
    mov eax, [undo_count]
    imul rax, UNDO_REC_SIZE
    lea r12, [undo_stack + rax] ; r12 = undo record to apply

    ; Push copy of record to redo_stack
    mov eax, [redo_count]
    cmp eax, UNDO_MAX
    jge .skip_redo_push
    imul rax, UNDO_REC_SIZE
    lea rdi, [redo_stack + rax]
    mov rsi, r12
    cld
    mov ecx, UNDO_REC_SIZE
    rep movsb
    inc dword [redo_count]
.skip_redo_push:

    ; Extract cell coordinates
    mov ebx, [r12 + 0] ; row
    mov esi, [r12 + 4] ; col

    ; 1. Restore grid cell text from old_text (+16)
    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], ebx
    mov dword [lvitem+8], esi
    lea rax, [r12 + 16]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    test rcx, rcx
    jz .undo_no_grid
    mov edx, LVM_SETITEMTEXTA
    mov r8d, ebx
    lea r9, [lvitem]
    call [SendMessageA]
.undo_no_grid:

    ; 2. Restore formula table from old_formula (+136)
    mov ecx, ebx
    mov edx, esi
    lea r8, [r12 + 136]
    cmp byte [r8], 0
    jne .undo_set_form
    xor r8, r8
.undo_set_form:
    call SetCellFormula

    ; 3. If row == cur_row and col == cur_col, update Formula Bar
    cmp ebx, [cur_row]
    jne .undo_sync_done
    cmp esi, [cur_col]
    jne .undo_sync_done

    lea rdx, [r12 + 136]
    cmp byte [rdx], 0
    jne .undo_show_fedit
    lea rdx, [r12 + 16]
.undo_show_fedit:
    mov rcx, [hFormulaEdit]
    test rcx, rcx
    jz .undo_sync_done
    call [SetWindowTextA]

.undo_sync_done:
    call SaveActiveSheetToBuffer
    call UpdateStatusStats

    mov rcx, [hGrid]
    test rcx, rcx
    jz .undo_msg_show
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]

.undo_msg_show:
    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [undo_msg]
    call [SendMessageA]
    jmp .undo_ret

.undo_empty:
    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [undo_empty_msg]
    call [SendMessageA]

.undo_ret:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret


; =============================================================================
; Action: Redo
; =============================================================================
DoRedo:
    push rbp
    mov rbp, rsp
    sub rsp, 80h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    cmp dword [redo_count], 0
    jle .redo_empty

    ; Pop record from redo_stack
    dec dword [redo_count]
    mov eax, [redo_count]
    imul rax, UNDO_REC_SIZE
    lea r12, [redo_stack + rax] ; r12 = redo record to reapply

    ; Push copy of record to undo_stack
    mov eax, [undo_count]
    cmp eax, UNDO_MAX
    jge .skip_undo_push
    imul rax, UNDO_REC_SIZE
    lea rdi, [undo_stack + rax]
    mov rsi, r12
    cld
    mov ecx, UNDO_REC_SIZE
    rep movsb
    inc dword [undo_count]
.skip_undo_push:

    ; Extract cell coordinates
    mov ebx, [r12 + 0] ; row
    mov esi, [r12 + 4] ; col

    ; 1. Set grid cell text to new_text (+256)
    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], ebx
    mov dword [lvitem+8], esi
    lea rax, [r12 + 256]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    test rcx, rcx
    jz .redo_no_grid
    mov edx, LVM_SETITEMTEXTA
    mov r8d, ebx
    lea r9, [lvitem]
    call [SendMessageA]
.redo_no_grid:

    ; 2. Set formula table to new_formula (+376)
    mov ecx, ebx
    mov edx, esi
    lea r8, [r12 + 376]
    cmp byte [r8], 0
    jne .redo_set_form
    xor r8, r8
.redo_set_form:
    call SetCellFormula

    ; 3. If row == cur_row and col == cur_col, update Formula Bar
    cmp ebx, [cur_row]
    jne .redo_sync_done
    cmp esi, [cur_col]
    jne .redo_sync_done

    lea rdx, [r12 + 376]
    cmp byte [rdx], 0
    jne .redo_show_fedit
    lea rdx, [r12 + 256]
.redo_show_fedit:
    mov rcx, [hFormulaEdit]
    test rcx, rcx
    jz .redo_sync_done
    call [SetWindowTextA]

.redo_sync_done:
    call SaveActiveSheetToBuffer
    call UpdateStatusStats

    mov rcx, [hGrid]
    test rcx, rcx
    jz .redo_msg_show
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]

.redo_msg_show:
    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [redo_msg]
    call [SendMessageA]
    jmp .redo_ret

.redo_empty:
    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [redo_empty_msg]
    call [SendMessageA]

.redo_ret:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret

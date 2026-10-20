; ============================================================================
; Excel Viewer Pro — Vector Math & Formula Engine
; ============================================================================



; =============================================================================
; Insert Formula Snippet into Formula Bar and Grid
; RCX = pointer to snippet string
; =============================================================================
InsertFormulaSnippet:
    push rbp
    mov rbp, rsp
    sub rsp, 30h
    mov [rbp-8], rcx

    mov ecx, [cur_row]
    mov edx, [cur_col]
    call BeginCellChange

    mov rcx, [hFormulaEdit]
    test rcx, rcx
    jz .ifs_grid
    mov rdx, [rbp-8]
    call [SetWindowTextA]

    mov rcx, [hFormulaEdit]
    call [SetFocus]
    mov rcx, [hFormulaEdit]
    mov edx, EM_SETSEL
    mov r8d, 100
    mov r9d, 100
    call [SendMessageA]

    ; Register formula in formula table
    mov ecx, [cur_row]
    mov edx, [cur_col]
    mov r8, [rbp-8]
    call SetCellFormula

.ifs_grid:
    mov dword [lvitem], LVIF_TEXT
    mov eax, [cur_row]
    mov dword [lvitem+4], eax
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    mov rax, [rbp-8]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    test rcx, rcx
    jz .ifs_commit
    mov edx, LVM_SETITEMTEXTA
    mov r8d, [cur_row]
    lea r9, [lvitem]
    call [SendMessageA]

    mov rcx, [hGrid]
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]

.ifs_commit:
    mov ecx, [cur_row]
    mov edx, [cur_col]
    mov r8, [rbp-8]
    mov r9, [rbp-8]
    call CommitCellChange

    call SaveActiveSheetToBuffer
    call UpdateStatusStats

.ifs_done:
    leave
    ret


; =============================================================================
; Action: Formula Evaluator & Arithmetic Parser
; =============================================================================
DoEvalFormula:
    push rbp
    mov rbp, rsp
    sub rsp, 70h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    mov rcx, [hFormulaEdit]
    lea rdx, [szTemp]
    mov r8d, 2040
    call [GetWindowTextA]
    test eax, eax
    jz .eval_done

    ; Save a copy in szFormulaInput
    lea rsi, [szTemp]
    lea rdi, [szFormulaInput]
.copy_inp:
    lodsb
    stosb
    test al, al
    jnz .copy_inp

    ; Track undo state for this cell
    mov ecx, [cur_row]
    mov edx, [cur_col]
    call BeginCellChange

    ; If does not start with '=', write directly as plain text
    cmp byte [szTemp], '='
    jne .write_plain_cell

    ; Register formula in formula table
    mov ecx, [cur_row]
    mov edx, [cur_col]
    lea r8, [szFormulaInput]
    call SetCellFormula

    ; Check for "=SUM(" or "=SQRT("
    cmp byte [szTemp+1], 'S'
    je .check_s
    cmp byte [szTemp+1], 's'
    je .check_s

    cmp byte [szTemp+1], 'A'
    je .check_a
    cmp byte [szTemp+1], 'a'
    je .check_a

    cmp byte [szTemp+1], 'M'
    je .check_min_max
    cmp byte [szTemp+1], 'm'
    je .check_min_max

    cmp byte [szTemp+1], 'C'
    je .check_count
    cmp byte [szTemp+1], 'c'
    je .check_count

    cmp byte [szTemp+1], 'P'
    je .check_p
    cmp byte [szTemp+1], 'p'
    je .check_p

    ; Check arithmetic expressions like =100+250 or =A1+B2
    jmp .parse_arithmetic

.check_s:
    cmp byte [szTemp+2], 'Q'
    je .check_sqrt
    cmp byte [szTemp+2], 'q'
    je .check_sqrt
    ; SUM
    lea rsi, [szTemp+5]
    call ParseRangeEval
    lea rcx, [szTemp]
    lea rdx, [fmt_f2]
    movq r8, xmm0
    call [sprintf]
    jmp .write_cell

.check_sqrt:
    lea rcx, [szTemp+6]
    call ParseOperand
    sqrtsd xmm0, xmm0
    lea rcx, [szTemp]
    lea rdx, [fmt_f2]
    movq r8, xmm0
    call [sprintf]
    jmp .write_cell

.check_a:
    cmp byte [szTemp+2], 'B'
    je .check_abs
    cmp byte [szTemp+2], 'b'
    je .check_abs
    ; AVG / AVERAGE
    lea rsi, [szTemp+5]
    cmp byte [szTemp+2], 'V'
    je .do_avg
    cmp byte [szTemp+2], 'v'
    je .do_avg
    lea rsi, [szTemp+9]
.do_avg:
    call ParseRangeEval
    test ebx, ebx
    jz .avg_fmt
    cvtsi2sd xmm1, ebx
    divsd xmm0, xmm1
.avg_fmt:
    lea rcx, [szTemp]
    lea rdx, [fmt_f2]
    movq r8, xmm0
    call [sprintf]
    jmp .write_cell

.check_abs:
    lea rcx, [szTemp+5]
    call ParseOperand
    pcmpeqd xmm1, xmm1
    psrlq xmm1, 1
    andpd xmm0, xmm1
    lea rcx, [szTemp]
    lea rdx, [fmt_f2]
    movq r8, xmm0
    call [sprintf]
    jmp .write_cell

.check_p:
    cmp byte [szTemp+2], 'I'
    je .check_pi
    cmp byte [szTemp+2], 'i'
    je .check_pi
    jmp .parse_arithmetic

.check_pi:
    mov rax, 0400921FB54442D18h ; 3.141592653589793
    movq xmm0, rax
    lea rcx, [szTemp]
    lea rdx, [fmt_f2]
    movq r8, xmm0
    call [sprintf]
    jmp .write_cell

.check_min_max:
    cmp byte [szTemp+2], 'I'
    je .is_min
    cmp byte [szTemp+2], 'i'
    je .is_min
    ; MAX
    lea rsi, [szTemp+5]
    call ParseRangeEvalMax
    lea rcx, [szTemp]
    lea rdx, [fmt_f2]
    movq r8, xmm0
    call [sprintf]
    jmp .write_cell

.is_min:
    lea rsi, [szTemp+5]
    call ParseRangeEvalMin
    lea rcx, [szTemp]
    lea rdx, [fmt_f2]
    movq r8, xmm0
    call [sprintf]
    jmp .write_cell

.check_count:
    lea rsi, [szTemp+7]
    call ParseRangeEval
    lea rcx, [szTemp]
    lea rdx, [fmt_d]
    mov r8d, ebx
    call [sprintf]
    jmp .write_cell

.parse_arithmetic:
    lea rsi, [szTemp+1]
    xor ebx, ebx
.scan_op:
    lodsb
    test al, al
    jz .no_arith_op
    cmp al, '+'
    je .op_found
    cmp al, '-'
    je .op_found
    cmp al, '*'
    je .op_found
    cmp al, '/'
    je .op_found
    jmp .scan_op

.no_arith_op:
    ; Check if formula contains '!' (e.g. cross-sheet 'Sheet'!A1)
    lea rsi, [szFormulaInput]
.scan_excl:
    lodsb
    test al, al
    jz .scan_for_paren
    cmp al, '!'
    je .eval_done_keep_grid
    jmp .scan_excl

.scan_for_paren:
    lea rsi, [szFormulaInput]
.scan_p_loop:
    lodsb
    test al, al
    jz .try_cell_ref
    cmp al, '('
    je .write_formula_text_cell
    jmp .scan_p_loop

.write_formula_text_cell:
    lea rsi, [szFormulaInput]
    lea rdi, [szTemp]
.cpy_f_txt:
    lodsb
    stosb
    test al, al
    jnz .cpy_f_txt
    jmp .write_cell

.try_cell_ref:
    ; Evaluate single cell reference (e.g. =C15 or =$C$5 or number)
    lea rcx, [szTemp+1]
    call ParseOperand
    lea rcx, [szTemp]
    lea rdx, [fmt_f2]
    movq r8, xmm0
    call [sprintf]
    jmp .write_cell

.eval_done_keep_grid:
    ; Retain current grid cell text, keep formula in formula bar
    mov rcx, [hFormulaEdit]
    lea rdx, [szFormulaInput]
    call [SetWindowTextA]
    jmp .eval_done

.op_found:
    mov bl, al
    mov byte [rsi-1], 0
    mov r12, rsi

    lea rcx, [szTemp+1]
    call ParseOperand
    movsd [rsp+20h], xmm0

    mov rcx, r12
    call ParseOperand
    movsd xmm1, xmm0
    movsd xmm0, [rsp+20h]

    cmp bl, '+'
    jne .not_add
    addsd xmm0, xmm1
    jmp .arith_done
.not_add:
    cmp bl, '-'
    jne .not_sub
    subsd xmm0, xmm1
    jmp .arith_done
.not_sub:
    cmp bl, '*'
    jne .not_mul
    mulsd xmm0, xmm1
    jmp .arith_done
.not_mul:
    divsd xmm0, xmm1

.arith_done:
    lea rcx, [szTemp]
    lea rdx, [fmt_f2]
    movq r8, xmm0
    call [sprintf]

.write_cell:
    ; Commit cell change to undo stack
    mov ecx, [cur_row]
    mov edx, [cur_col]
    lea r8, [szTemp]
    lea r9, [szFormulaInput]
    call CommitCellChange

    ; Write calculated result (szTemp) to grid cell
    mov dword [lvitem], LVIF_TEXT
    mov eax, [cur_row]
    mov dword [lvitem+4], eax
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szTemp]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, [cur_row]
    lea r9, [lvitem]
    call [SendMessageA]

    ; Keep formula in Formula Bar
    mov rcx, [hFormulaEdit]
    lea rdx, [szFormulaInput]
    call [SetWindowTextA]

    call SaveActiveSheetToBuffer
    call UpdateStatusStats
    jmp .eval_done

.write_plain_cell:
    ; Commit cell change to undo stack
    mov ecx, [cur_row]
    mov edx, [cur_col]
    lea r8, [szFormulaInput]
    xor r9, r9
    call CommitCellChange

    ; Clear formula in formula table
    mov ecx, [cur_row]
    mov edx, [cur_col]
    xor r8, r8
    call SetCellFormula

    mov dword [lvitem], LVIF_TEXT
    mov eax, [cur_row]
    mov dword [lvitem+4], eax
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szFormulaInput]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, [cur_row]
    lea r9, [lvitem]
    call [SendMessageA]

    mov rcx, [hFormulaEdit]
    lea rdx, [szFormulaInput]
    call [SetWindowTextA]

    call SaveActiveSheetToBuffer
    call UpdateStatusStats

.eval_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret


; =============================================================================
; Action: SIMD AutoSum Column
; =============================================================================
DoSimdAutoSum:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], r12

    ; Sum all numbers in cur_col above cur_row
    xor rax, rax
    mov [rbp-24], rax      ; sum = 0.0

    mov r12d, 0
.sum_loop:
    cmp r12d, [cur_row]
    jge .sum_write

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r12d
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 1000
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r12d
    lea r9, [lvitem]
    call [SendMessageA]

    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]

    movsd xmm1, [rbp-24]
    addsd xmm1, xmm0
    movsd [rbp-24], xmm1

    inc r12d
    jmp .sum_loop

.sum_write:
    lea rcx, [szTemp]
    lea rdx, [fmt_f2]
    movsd xmm0, [rbp-24]
    movq r8, xmm0
    call [sprintf]

    mov dword [lvitem], LVIF_TEXT
    mov eax, [cur_row]
    mov dword [lvitem+4], eax
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szTemp]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, [cur_row]
    lea r9, [lvitem]
    call [SendMessageA]

    mov rcx, [hFormulaEdit]
    lea rdx, [szTemp]
    call [SetWindowTextA]

    call UpdateStatusStats

    mov rbx, [rbp-8]
    mov r12, [rbp-16]
    leave
    ret


; =============================================================================
; Native SIMD Vector Sum (SSE2)
; =============================================================================
asm_vec_sum:
    xorpd xmm0, xmm0
    test rcx, rcx
    jz .done
    test rdx, rdx
    jz .done

    xorpd xmm1, xmm1
    mov r8, rdx
    shr r8, 2
    jz .tail

.loop4:
    addpd xmm0, [rcx]
    addpd xmm1, [rcx+16]
    add rcx, 32
    dec r8
    jnz .loop4
    addpd xmm0, xmm1

.tail:
    and rdx, 3
    jz .hadd
.loop1:
    addsd xmm0, [rcx]
    add rcx, 8
    dec rdx
    jnz .loop1
.hadd:
    movhlps xmm2, xmm0
    addsd xmm0, xmm2
.done:
    ret


; =============================================================================
; Action: SIMD Functions & Dialogs
; =============================================================================
DoSimdAverage:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], r12

    xor rax, rax
    mov [rbp-24], rax
    xor ebx, ebx

    mov r12d, 0
.avg_loop:
    cmp r12d, [cur_row]
    jge .avg_calc

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r12d
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 1000
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r12d
    lea r9, [lvitem]
    call [SendMessageA]

    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]

    inc ebx
    movsd xmm1, [rbp-24]
    addsd xmm1, xmm0
    movsd [rbp-24], xmm1

    inc r12d
    jmp .avg_loop

.avg_calc:
    movsd xmm0, [rbp-24]
    test ebx, ebx
    jz .avg_write
    cvtsi2sd xmm1, ebx
    divsd xmm0, xmm1

.avg_write:
    lea rcx, [szTemp]
    lea rdx, [fmt_f2]
    movq r8, xmm0
    call [sprintf]

    mov dword [lvitem], LVIF_TEXT
    mov eax, [cur_row]
    mov dword [lvitem+4], eax
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szTemp]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, [cur_row]
    lea r9, [lvitem]
    call [SendMessageA]

    mov rcx, [hFormulaEdit]
    lea rdx, [szTemp]
    call [SetWindowTextA]

    call UpdateStatusStats

    mov rbx, [rbp-8]
    mov r12, [rbp-16]
    leave
    ret

DoSimdMin:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], r12

    mov rax, 07FEFFFFFFFFFFFFFh
    mov [rbp-24], rax
    xor ebx, ebx

    mov r12d, 0
.min_loop:
    cmp r12d, [cur_row]
    jge .min_write

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r12d
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 1000
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r12d
    lea r9, [lvitem]
    call [SendMessageA]

    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]

    inc ebx
    movsd xmm1, [rbp-24]
    minsd xmm1, xmm0
    movsd [rbp-24], xmm1

    inc r12d
    jmp .min_loop

.min_write:
    test ebx, ebx
    jnz .min_fmt
    xorpd xmm1, xmm1
    movsd [rbp-24], xmm1
.min_fmt:
    lea rcx, [szTemp]
    lea rdx, [fmt_f2]
    mov r8, [rbp-24]
    call [sprintf]

    mov dword [lvitem], LVIF_TEXT
    mov eax, [cur_row]
    mov dword [lvitem+4], eax
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szTemp]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, [cur_row]
    lea r9, [lvitem]
    call [SendMessageA]

    mov rcx, [hFormulaEdit]
    lea rdx, [szTemp]
    call [SetWindowTextA]

    call UpdateStatusStats

    mov rbx, [rbp-8]
    mov r12, [rbp-16]
    leave
    ret

DoSimdMax:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], r12

    mov rax, 0FFEFFFFFFFFFFFFFh
    mov [rbp-24], rax
    xor ebx, ebx

    mov r12d, 0
.max_loop:
    cmp r12d, [cur_row]
    jge .max_write

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r12d
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 1000
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r12d
    lea r9, [lvitem]
    call [SendMessageA]

    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]

    inc ebx
    movsd xmm1, [rbp-24]
    maxsd xmm1, xmm0
    movsd [rbp-24], xmm1

    inc r12d
    jmp .max_loop

.max_write:
    test ebx, ebx
    jnz .max_fmt
    xorpd xmm1, xmm1
    movsd [rbp-24], xmm1
.max_fmt:
    lea rcx, [szTemp]
    lea rdx, [fmt_f2]
    mov r8, [rbp-24]
    call [sprintf]

    mov dword [lvitem], LVIF_TEXT
    mov eax, [cur_row]
    mov dword [lvitem+4], eax
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szTemp]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, [cur_row]
    lea r9, [lvitem]
    call [SendMessageA]

    mov rcx, [hFormulaEdit]
    lea rdx, [szTemp]
    call [SetWindowTextA]

    call UpdateStatusStats

    mov rbx, [rbp-8]
    mov r12, [rbp-16]
    leave
    ret

DoRecalculateAll:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    call UpdateStatusStats
    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [recalc_done_msg]
    call [SendMessageA]
    leave
    ret


; =============================================================================
; Formula Table Access: GetCellFormula, SetCellFormula
; =============================================================================
GetCellFormula:
    cmp ecx, 10000
    jae .gcf_none
    cmp edx, 32
    jae .gcf_none
    mov rax, [pCellFormulaTable]
    test rax, rax
    jz .gcf_none

    shl ecx, 5
    add ecx, edx
    mov rax, [rax + rcx*8]
    ret

.gcf_none:
    xor eax, eax
    ret

SetCellFormula:
    push rbp
    mov rbp, rsp
    sub rsp, 40h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12

    cmp ecx, 10000
    jae .scf_done
    cmp edx, 32
    jae .scf_done
    mov rax, [pCellFormulaTable]
    test rax, rax
    jz .scf_done

    shl ecx, 5
    add ecx, edx
    mov rbx, rcx            ; rbx = cell index

    test r8, r8
    jz .scf_clear
    cmp byte [r8], 0
    jz .scf_clear

    mov rsi, r8

    ; Find string length
    mov rdi, r8
    xor eax, eax
    mov ecx, 1000
    repne scasb
    mov eax, 1000
    sub eax, ecx            ; length including null

    test eax, eax
    jle .scf_clear

    mov r12d, eax           ; size in bytes

    ; Allocate memory in hFormulaHeap
    mov rcx, [hFormulaHeap]
    test rcx, rcx
    jz .scf_done
    xor edx, edx
    mov r8d, r12d
    sub rsp, 20h
    call [HeapAlloc]
    add rsp, 20h
    test rax, rax
    jz .scf_done

    ; Copy string
    mov rdi, rax
    mov ecx, r12d
    rep movsb

    ; Store pointer in pCellFormulaTable[rbx]
    mov rdx, [pCellFormulaTable]
    mov [rdx + rbx*8], rax
    jmp .scf_done

.scf_clear:
    mov rdx, [pCellFormulaTable]
    mov qword [rdx + rbx*8], 0

.scf_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    leave
    ret


; =============================================================================
; Formula Parsing Helpers: ParseOperand, ParseRangeEval, ParseRangeEvalMin, ParseRangeEvalMax
; =============================================================================
ParseOperand:
    push rbp
    mov rbp, rsp
    sub rsp, 40h

    cmp byte [rcx], '$'
    jne .p_op_col
    inc rcx
.p_op_col:
    mov al, byte [rcx]
    and al, 0DFh
    cmp al, 'A'
    jl .op_num
    cmp al, 'Z'
    jg .op_num

    sub al, 'A' - 1
    movzx edx, al
    inc rcx
    cmp byte [rcx], '$'
    jne .p_op_row
    inc rcx
.p_op_row:
    xor eax, eax
.parse_r_num:
    movzx r8d, byte [rcx]
    cmp r8b, '0'
    jl .r_num_done
    cmp r8b, '9'
    jg .r_num_done
    sub r8d, '0'
    imul eax, 10
    add eax, r8d
    inc rcx
    jmp .parse_r_num
.r_num_done:
    dec eax
    cmp eax, 0
    jl .op_num
    cmp eax, [total_rows]
    jge .op_num

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], eax
    mov dword [lvitem+8], edx
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 500
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, dword [lvitem+4]
    lea r9, [lvitem]
    call [SendMessageA]

    lea rcx, [szCellVal]
.skip_cell_curr:
    mov al, byte [rcx]
    cmp al, '$'
    je .inc_cell_curr
    cmp al, ' '
    je .inc_cell_curr
    jmp .do_cell_strtod
.inc_cell_curr:
    inc rcx
    jmp .skip_cell_curr
.do_cell_strtod:
    xor edx, edx
    call [strtod]
    leave
    ret

.op_num:
.skip_op_curr:
    mov al, byte [rcx]
    cmp al, '$'
    je .inc_op_curr
    cmp al, ' '
    je .inc_op_curr
    jmp .do_op_strtod
.inc_op_curr:
    inc rcx
    jmp .skip_op_curr
.do_op_strtod:
    xor edx, edx
    call [strtod]
    leave
    ret

ParseRangeEval:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], r12
    mov [rbp-16], r13
    mov [rbp-24], r14
    mov [rbp-32], r15

    lodsb
    and al, 0DFh
    sub al, 'A' - 1
    movzx r12d, al
    xor eax, eax
.r1_loop:
    lodsb
    cmp al, '0'
    jl .r1_done
    cmp al, '9'
    jg .r1_done
    sub al, '0'
    movzx edx, al
    imul eax, 10
    add eax, edx
    jmp .r1_loop
.r1_done:
    dec eax
    mov r13d, eax
    lodsb
    and al, 0DFh
    sub al, 'A' - 1
    movzx r14d, al
    xor eax, eax
.r2_loop:
    lodsb
    cmp al, '0'
    jl .r2_done
    cmp al, '9'
    jg .r2_done
    sub al, '0'
    movzx edx, al
    imul eax, 10
    add eax, edx
    jmp .r2_loop
.r2_done:
    dec eax
    mov r15d, eax

    xorpd xmm0, xmm0
    movsd [rbp-40], xmm0
    xor ebx, ebx

.sum_r:
    cmp r13d, r15d
    jg .sum_range_done
    mov edx, r12d
.sum_c:
    cmp edx, r14d
    jg .sum_next_r

    mov [rsp+20h], rdx
    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov dword [lvitem+8], edx
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 500
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r13d
    lea r9, [lvitem]
    call [SendMessageA]

    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]

    inc ebx
    movsd xmm1, [rbp-40]
    addsd xmm1, xmm0
    movsd [rbp-40], xmm1

    mov rdx, [rsp+20h]
    inc edx
    jmp .sum_c

.sum_next_r:
    inc r13d
    jmp .sum_r

.sum_range_done:
    movsd xmm0, [rbp-40]
    mov r12, [rbp-8]
    mov r13, [rbp-16]
    mov r14, [rbp-24]
    mov r15, [rbp-32]
    leave
    ret

ParseRangeEvalMin:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], r12
    mov [rbp-16], r13
    mov [rbp-24], r14
    mov [rbp-32], r15

    lodsb
    and al, 0DFh
    sub al, 'A' - 1
    movzx r12d, al
    xor eax, eax
.r1_min_loop:
    lodsb
    cmp al, '0'
    jl .r1_min_done
    cmp al, '9'
    jg .r1_min_done
    sub al, '0'
    movzx edx, al
    imul eax, 10
    add eax, edx
    jmp .r1_min_loop
.r1_min_done:
    dec eax
    mov r13d, eax
    lodsb
    and al, 0DFh
    sub al, 'A' - 1
    movzx r14d, al
    xor eax, eax
.r2_min_loop:
    lodsb
    cmp al, '0'
    jl .r2_min_done
    cmp al, '9'
    jg .r2_min_done
    sub al, '0'
    movzx edx, al
    imul eax, 10
    add eax, edx
    jmp .r2_min_loop
.r2_min_done:
    dec eax
    mov r15d, eax

    mov rax, 07FEFFFFFFFFFFFFFh
    mov [rbp-40], rax
    xor ebx, ebx

.min_r:
    cmp r13d, r15d
    jg .min_range_done
    mov edx, r12d
.min_c:
    cmp edx, r14d
    jg .min_next_r

    mov [rsp+20h], rdx
    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov dword [lvitem+8], edx
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 500
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r13d
    lea r9, [lvitem]
    call [SendMessageA]

    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]

    inc ebx
    movsd xmm1, [rbp-40]
    minsd xmm1, xmm0
    movsd [rbp-40], xmm1

    mov rdx, [rsp+20h]
    inc edx
    jmp .min_c

.min_next_r:
    inc r13d
    jmp .min_r

.min_range_done:
    test ebx, ebx
    jnz .ret_min
    xorpd xmm0, xmm0
    movsd [rbp-40], xmm0
.ret_min:
    movsd xmm0, [rbp-40]
    mov r12, [rbp-8]
    mov r13, [rbp-16]
    mov r14, [rbp-24]
    mov r15, [rbp-32]
    leave
    ret

ParseRangeEvalMax:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], r12
    mov [rbp-16], r13
    mov [rbp-24], r14
    mov [rbp-32], r15

    lodsb
    and al, 0DFh
    sub al, 'A' - 1
    movzx r12d, al
    xor eax, eax
.r1_max_loop:
    lodsb
    cmp al, '0'
    jl .r1_max_done
    cmp al, '9'
    jg .r1_max_done
    sub al, '0'
    movzx edx, al
    imul eax, 10
    add eax, edx
    jmp .r1_max_loop
.r1_max_done:
    dec eax
    mov r13d, eax
    lodsb
    and al, 0DFh
    sub al, 'A' - 1
    movzx r14d, al
    xor eax, eax
.r2_max_loop:
    lodsb
    cmp al, '0'
    jl .r2_max_done
    cmp al, '9'
    jg .r2_max_done
    sub al, '0'
    movzx edx, al
    imul eax, 10
    add eax, edx
    jmp .r2_max_loop
.r2_max_done:
    dec eax
    mov r15d, eax

    mov rax, 0FFEFFFFFFFFFFFFFh
    mov [rbp-40], rax
    xor ebx, ebx

.max_r:
    cmp r13d, r15d
    jg .max_range_done
    mov edx, r12d
.max_c:
    cmp edx, r14d
    jg .max_next_r

    mov [rsp+20h], rdx
    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov dword [lvitem+8], edx
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 500
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r13d
    lea r9, [lvitem]
    call [SendMessageA]

    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]

    inc ebx
    movsd xmm1, [rbp-40]
    maxsd xmm1, xmm0
    movsd [rbp-40], xmm1

    mov rdx, [rsp+20h]
    inc edx
    jmp .max_c

.max_next_r:
    inc r13d
    jmp .max_r

.max_range_done:
    test ebx, ebx
    jnz .ret_max
    xorpd xmm0, xmm0
    movsd [rbp-40], xmm0
.ret_max:
    movsd xmm0, [rbp-40]
    mov r12, [rbp-8]
    mov r13, [rbp-16]
    mov r14, [rbp-24]
    mov r15, [rbp-32]
    leave
    ret


; =============================================================================
; Direct Imports Section (PE64 IAT)
; =============================================================================

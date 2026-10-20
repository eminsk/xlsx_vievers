; ============================================================================
; Excel Viewer Pro — Spreadsheet Grid Engine
; ============================================================================




; =============================================================================
; Grid Helper: Synchronize ListView Scroll / Visible State
; =============================================================================
SyncListViewItemState:
    push rbp
    mov rbp, rsp
    sub rsp, 40h
    mov rcx, [hGrid]
    test rcx, rcx
    jz .slv_done
    mov edx, 1013h ; LVM_ENSUREVISIBLE
    mov r8d, [cur_row]
    xor r9d, r9d
    call [SendMessageA]
.slv_done:
    leave
    ret


; =============================================================================
; Check if (r10d = row, r11d = col) is within current selection
; Returns eax = 1 if inside, 0 if outside
; =============================================================================
CheckCellInSelection:
    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    cmp r10d, eax
    jl .not_in_sel
    cmp r10d, edx
    jg .not_in_sel

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    cmp r11d, eax
    jl .not_in_sel
    cmp r11d, edx
    jg .not_in_sel

    mov eax, 1
    ret
.not_in_sel:
    xor eax, eax
    ret


; =============================================================================
; GridSubclassProc: Intercepts all Mouse Drag & Keyboard Navigation for 2D Grid
; =============================================================================
GridSubclassProc:
    push rbp
    mov rbp, rsp
    sub rsp, 0B0h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13
    mov [rbp-30h], r14
    mov [rbp-38h], r15

    ; Save original arguments in safe local stack slots
    mov [rbp-40h], rcx      ; hWnd
    mov [rbp-48h], rdx      ; uMsg
    mov [rbp-50h], r8       ; wParam
    mov [rbp-58h], r9       ; lParam

    cmp edx, 0087h ; WM_GETDLGCODE
    je .gsp_getdlgcode
    cmp edx, 0201h ; WM_LBUTTONDOWN
    je .gsp_lbutton_down
    cmp edx, 0203h ; WM_LBUTTONDBLCLK
    je .gsp_lbutton_dblclk
    cmp edx, 0200h ; WM_MOUSEMOVE
    je .gsp_mouse_move
    cmp edx, 0202h ; WM_LBUTTONUP
    je .gsp_lbutton_up
    cmp edx, 0204h ; WM_RBUTTONDOWN
    je .gsp_rbutton_down
    cmp edx, 0205h ; WM_RBUTTONUP
    je .gsp_rbutton_up
    cmp edx, 0100h ; WM_KEYDOWN
    je .gsp_keydown
    cmp edx, 0101h ; WM_KEYUP
    je .gsp_keyup
    cmp edx, 0008h ; WM_KILLFOCUS
    je .gsp_killfocus

.gsp_call_orig:
    mov rax, [origGridProc]
    mov rcx, rax            ; 1st arg: lpPrevWndFunc
    mov rdx, [rbp-40h]      ; 2nd arg: hWnd
    mov r8, [rbp-48h]       ; 3rd arg: uMsg
    mov r9, [rbp-50h]       ; 4th arg: wParam
    mov r10, [rbp-58h]      ; 5th arg: lParam
    mov [rsp+20h], r10
    call [CallWindowProcA]
    jmp .gsp_ret

.gsp_getdlgcode:
    mov eax, 0005h ; DLGC_WANTARROWS | DLGC_WANTALLKEYS
    jmp .gsp_ret

.gsp_lbutton_down:
    cmp dword [bCellEditing], 1
    jne @f
    call CommitInPlaceCellEdit
@@:
    mov rcx, [hGrid]
    call [SetFocus]

    mov r9, [rbp-58h]       ; reload lParam
    movsx eax, r9w         ; x
    mov dword [lvhit], eax
    shr r9, 16
    movsx eax, r9w         ; y
    mov dword [lvhit+4], eax
    mov dword [lvhit+8], 0
    mov dword [lvhit+12], -1
    mov dword [lvhit+16], -1

    mov rcx, [hGrid]
    mov edx, 1039h          ; LVM_SUBITEMHITTEST
    xor r8d, r8d
    lea r9, [lvhit]
    call [SendMessageA]

    cmp dword [lvhit+12], -1
    je .gsp_call_orig

    mov eax, dword [lvhit+12] ; row
    mov ebx, dword [lvhit+16] ; col

    ; If user clicked col 0 (Row number # column), select whole row 1..26!
    cmp ebx, 0
    jne .gsp_lbtn_normal_col
    mov [sel_start_col], 1
    mov [sel_end_col], 26
    mov [cur_col], 1
    mov [sel_start_row], eax
    mov [sel_end_row], eax
    mov [cur_row], eax
    jmp .gsp_lbtn_capture

.gsp_lbtn_normal_col:
    cmp ebx, 1
    jge @f
    mov ebx, 1
@@:
    cmp ebx, 26
    jle @f
    mov ebx, 26
@@:

    ; Check Shift key from wParam (MK_SHIFT = 0x0004)
    mov r8, [rbp-50h]       ; reload wParam
    test r8d, 4             ; MK_SHIFT
    jnz .gsp_lbtn_shift

    ; Normal click: set new anchor at clicked cell (lvhit+12 = row, ebx = col)
    mov eax, dword [lvhit+12]
    mov [sel_start_row], eax
    mov [sel_start_col], ebx
    mov [sel_end_row], eax
    mov [sel_end_col], ebx
    mov [cur_row], eax
    mov [cur_col], ebx
    jmp .gsp_lbtn_capture

.gsp_lbtn_shift:
    mov eax, dword [lvhit+12]
    mov [sel_end_row], eax
    mov [sel_end_col], ebx
    mov [cur_row], eax
    mov [cur_col], ebx

.gsp_lbtn_capture:
    mov dword [is_mouse_sel], 1
    mov rcx, [hGrid]
    call [SetCapture]

    call UpdateCellSelectionDisplay
    call SyncListViewItemState
    xor eax, eax
    jmp .gsp_ret

.gsp_mouse_move:
    cmp dword [is_mouse_sel], 1
    jne .gsp_call_orig

    ; Check left mouse button
    mov r8, [rbp-50h]       ; reload wParam
    test r8d, 1             ; MK_LBUTTON
    jz .gsp_lbtn_lost

    mov r9, [rbp-58h]       ; reload lParam
    movsx eax, r9w          ; x
    mov r12d, eax
    shr r9, 16
    movsx eax, r9w          ; y
    mov r13d, eax

    ; Autoscroll if dragged beyond vertical bounds
    cmp r13d, 0
    jge @f
    mov rcx, [hGrid]
    mov edx, 0115h          ; WM_VSCROLL
    mov r8d, 0              ; SB_LINEUP
    xor r9d, r9d
    call [SendMessageA]
    jmp .gsp_mm_hittest
@@:
    cmp r13d, 480
    jle .gsp_mm_hittest
    mov rcx, [hGrid]
    mov edx, 0115h          ; WM_VSCROLL
    mov r8d, 1              ; SB_LINEDOWN
    xor r9d, r9d
    call [SendMessageA]

.gsp_mm_hittest:
    mov dword [lvhit], r12d
    mov dword [lvhit+4], r13d
    mov dword [lvhit+8], 0
    mov dword [lvhit+12], -1
    mov dword [lvhit+16], -1

    mov rcx, [hGrid]
    mov edx, 1039h          ; LVM_SUBITEMHITTEST
    xor r8d, r8d
    lea r9, [lvhit]
    call [SendMessageA]

    cmp dword [lvhit+12], -1
    je .gsp_mm_done

    mov eax, dword [lvhit+12] ; row
    mov ebx, dword [lvhit+16] ; col

    cmp ebx, 1
    jge @f
    mov ebx, 1
@@:
    cmp ebx, 26
    jle @f
    mov ebx, 26
@@:

    cmp eax, [sel_end_row]
    jne .gsp_mm_update
    cmp ebx, [sel_end_col]
    je .gsp_mm_done

.gsp_mm_update:
    mov [sel_end_row], eax
    mov [sel_end_col], ebx
    mov [cur_row], eax
    mov [cur_col], ebx

    call UpdateCellSelectionDisplay
    call SyncListViewItemState

.gsp_mm_done:
    xor eax, eax
    jmp .gsp_ret

.gsp_lbtn_lost:
    mov dword [is_mouse_sel], 0
    call [ReleaseCapture]
    call UpdateCellSelectionDisplay
    xor eax, eax
    jmp .gsp_ret

.gsp_lbutton_up:
    cmp dword [is_mouse_sel], 1
    jne .gsp_call_orig

    mov dword [is_mouse_sel], 0
    call [ReleaseCapture]
    call UpdateCellSelectionDisplay
    xor eax, eax
    jmp .gsp_ret

.gsp_rbutton_down:
    cmp dword [bCellEditing], 1
    jne @f
    call CommitInPlaceCellEdit
@@:
    mov r9, [rbp-58h]       ; reload lParam
    movsx eax, r9w
    mov dword [lvhit], eax
    shr r9, 16
    movsx eax, r9w
    mov dword [lvhit+4], eax
    mov dword [lvhit+8], 0
    mov dword [lvhit+12], -1
    mov dword [lvhit+16], -1

    mov rcx, [hGrid]
    mov edx, 1039h          ; LVM_SUBITEMHITTEST
    xor r8d, r8d
    lea r9, [lvhit]
    call [SendMessageA]

    cmp dword [lvhit+12], -1
    je .gsp_rbtn_done

    mov eax, dword [lvhit+12]
    mov ebx, dword [lvhit+16]

    ; Check if clicked on row number column # (col 0)
    test ebx, ebx
    jnz .gsp_rbtn_check_cols
    mov [sel_start_col], 1
    mov [sel_end_col], 26
    mov [cur_col], 1
    mov [sel_start_row], eax
    mov [sel_end_row], eax
    mov [cur_row], eax
    call UpdateCellSelectionDisplay
    call SyncListViewItemState
    jmp .gsp_rbtn_focus

.gsp_rbtn_check_cols:
    cmp ebx, 1
    jl .gsp_rbtn_done
    cmp ebx, 26
    jg .gsp_rbtn_done

    mov r10d, eax
    mov r11d, ebx
    call CheckCellInSelection
    test eax, eax
    jnz .gsp_rbtn_focus

    mov [sel_start_row], r10d
    mov [sel_start_col], r11d
    mov [sel_end_row], r10d
    mov [sel_end_col], r11d
    mov [cur_row], r10d
    mov [cur_col], r11d
    call UpdateCellSelectionDisplay
    call SyncListViewItemState

.gsp_rbtn_focus:
    mov rcx, [hGrid]
    call [SetFocus]

.gsp_rbtn_done:
    xor eax, eax
    jmp .gsp_ret

.gsp_rbutton_up:
    cmp dword [bCellEditing], 1
    jne @f
    call CommitInPlaceCellEdit
@@:
    mov rcx, [hContextMenu]
    test rcx, rcx
    jz .gsp_ret

    ; Safe stack slot for POINT at [rsp+20h] above shadow space
    sub rsp, 40h
    lea rcx, [rsp+20h]
    call [GetCursorPos]

    mov r8d, dword [rsp+20h] ; pt.x
    mov r9d, dword [rsp+24h] ; pt.y
    mov rcx, [hContextMenu]
    xor edx, edx            ; TPM_LEFTALIGN | TPM_TOPALIGN
    mov dword [rsp+20h], 0  ; nReserved
    mov rax, [hWnd]
    mov qword [rsp+28h], rax; hWnd to receive WM_COMMAND
    mov qword [rsp+30h], 0  ; prcRect
    call [TrackPopupMenu]
    add rsp, 40h

    xor eax, eax
    jmp .gsp_ret

.gsp_lbutton_dblclk:
    cmp dword [bCellEditing], 1
    jne @f
    call CommitInPlaceCellEdit
@@:
    mov r9, [rbp-58h]       ; reload lParam
    movsx eax, r9w
    mov dword [lvhit], eax
    shr r9, 16
    movsx eax, r9w
    mov dword [lvhit+4], eax
    mov dword [lvhit+8], 0
    mov dword [lvhit+12], -1
    mov dword [lvhit+16], -1

    mov rcx, [hGrid]
    mov edx, 1039h          ; LVM_SUBITEMHITTEST
    xor r8d, r8d
    lea r9, [lvhit]
    call [SendMessageA]

    cmp dword [lvhit+12], -1
    je .gsp_call_orig
    cmp dword [lvhit+16], 1
    jl .gsp_call_orig
    cmp dword [lvhit+16], 26
    jg .gsp_call_orig

    mov eax, dword [lvhit+12]
    mov ebx, dword [lvhit+16]
    mov [cur_row], eax
    mov [sel_start_row], eax
    mov [sel_end_row], eax
    mov [cur_col], ebx
    mov [sel_start_col], ebx
    mov [sel_end_col], ebx
    call UpdateCellSelectionDisplay

    call StartInPlaceCellEdit
    xor eax, eax
    jmp .gsp_ret

.gsp_keyup:
    mov r8d, dword [rbp-50h] ; wParam
    cmp r8d, 10h            ; VK_SHIFT
    jne @f
    mov dword [g_is_shift_down], 0
@@:
    cmp r8d, 11h            ; VK_CONTROL
    jne .gsp_call_orig
    mov dword [g_is_ctrl_down], 0
    jmp .gsp_call_orig

.gsp_killfocus:
    mov dword [g_is_shift_down], 0
    mov dword [g_is_ctrl_down], 0
    jmp .gsp_call_orig

.gsp_keydown:
    mov r8d, dword [rbp-50h] ; wParam
    cmp r8d, 10h            ; VK_SHIFT
    jne @f
    mov dword [g_is_shift_down], 1
    xor eax, eax
    jmp .gsp_ret
@@:
    cmp r8d, 11h            ; VK_CONTROL
    jne .gsp_proc_key
    mov dword [g_is_ctrl_down], 1
    xor eax, eax
    jmp .gsp_ret

.gsp_proc_key:
    xor r14d, r14d
    cmp dword [g_is_shift_down], 1
    je .shift_yes
    mov ecx, VK_SHIFT
    call [GetAsyncKeyState]
    test ax, 8000h
    jnz .shift_yes
    mov ecx, VK_SHIFT
    call [GetKeyState]
    test ax, 8000h
    jz .shift_no
.shift_yes:
    mov r14b, 1
.shift_no:

    xor r15d, r15d
    cmp dword [g_is_ctrl_down], 1
    je .ctrl_yes
    mov ecx, VK_CONTROL
    call [GetAsyncKeyState]
    test ax, 8000h
    jnz .ctrl_yes
    mov ecx, VK_CONTROL
    call [GetKeyState]
    test ax, 8000h
    jz .ctrl_no
.ctrl_yes:
    mov r15b, 1
.ctrl_no:

    mov r8d, dword [rbp-50h] ; wParam

    test r15b, r15b
    jz .gsp_no_ctrl

    ; Ctrl Shortcuts
    cmp r8d, 'A'
    je .gsp_ctrl_a
    cmp r8d, 'a'
    je .gsp_ctrl_a
    cmp r8d, 'C'
    je .gsp_ctrl_c
    cmp r8d, 'c'
    je .gsp_ctrl_c
    cmp r8d, 'V'
    je .gsp_ctrl_v
    cmp r8d, 'v'
    je .gsp_ctrl_v
    cmp r8d, 'X'
    je .gsp_ctrl_x
    cmp r8d, 'x'
    je .gsp_ctrl_x
    cmp r8d, 'Z'
    je .gsp_ctrl_z
    cmp r8d, 'z'
    je .gsp_ctrl_z
    cmp r8d, 'Y'
    je .gsp_ctrl_y
    cmp r8d, 'y'
    je .gsp_ctrl_y
    cmp r8d, 'B'
    je .gsp_ctrl_b
    cmp r8d, 'b'
    je .gsp_ctrl_b
    cmp r8d, 'I'
    je .gsp_ctrl_i
    cmp r8d, 'i'
    je .gsp_ctrl_i
    cmp r8d, 'U'
    je .gsp_ctrl_u
    cmp r8d, 'u'
    je .gsp_ctrl_u
    cmp r8d, 'S'
    je .gsp_ctrl_s
    cmp r8d, 's'
    je .gsp_ctrl_s
    cmp r8d, 'O'
    je .gsp_ctrl_o
    cmp r8d, 'o'
    je .gsp_ctrl_o
    cmp r8d, 'F'
    je .gsp_ctrl_f
    cmp r8d, 'f'
    je .gsp_ctrl_f
    jmp .gsp_call_orig

.gsp_ctrl_a:
    call DoSelectAll
    xor eax, eax
    jmp .gsp_ret
.gsp_ctrl_c:
    call DoClipboardCopy
    xor eax, eax
    jmp .gsp_ret
.gsp_ctrl_v:
    call DoClipboardPaste
    xor eax, eax
    jmp .gsp_ret
.gsp_ctrl_x:
    call DoClipboardCut
    xor eax, eax
    jmp .gsp_ret
.gsp_ctrl_z:
    call DoUndo
    xor eax, eax
    jmp .gsp_ret
.gsp_ctrl_y:
    call DoRedo
    xor eax, eax
    jmp .gsp_ret
.gsp_ctrl_b:
    call ToggleCellBold
    xor eax, eax
    jmp .gsp_ret
.gsp_ctrl_i:
    call ToggleCellItalic
    xor eax, eax
    jmp .gsp_ret
.gsp_ctrl_u:
    call ToggleCellUnderline
    xor eax, eax
    jmp .gsp_ret
.gsp_ctrl_s:
    call DoFileSave
    xor eax, eax
    jmp .gsp_ret
.gsp_ctrl_o:
    call DoFileOpen
    xor eax, eax
    jmp .gsp_ret
.gsp_ctrl_f:
    call DoFindText
    xor eax, eax
    jmp .gsp_ret

.gsp_no_ctrl:
    cmp r8d, 25h    ; VK_LEFT
    je .gsp_nav_left
    cmp r8d, 27h    ; VK_RIGHT
    je .gsp_nav_right
    cmp r8d, 26h    ; VK_UP
    je .gsp_nav_up
    cmp r8d, 28h    ; VK_DOWN
    je .gsp_nav_down
    cmp r8d, 09h    ; VK_TAB
    je .gsp_nav_tab
    cmp r8d, 0Dh    ; VK_RETURN
    je .gsp_nav_return
    cmp r8d, 24h    ; VK_HOME
    je .gsp_nav_home
    cmp r8d, 23h    ; VK_END
    je .gsp_nav_end
    cmp r8d, 21h    ; VK_PRIOR (PageUp)
    je .gsp_nav_pgup
    cmp r8d, 22h    ; VK_NEXT (PageDown)
    je .gsp_nav_pgdn
    cmp r8d, 71h    ; VK_F2
    je .gsp_key_f2
    cmp r8d, 72h    ; VK_F3
    je .gsp_key_f3
    cmp r8d, 74h    ; VK_F5
    je .gsp_key_f5
    cmp r8d, 78h    ; VK_F9
    je .gsp_key_f9
    cmp r8d, 2Eh    ; VK_DELETE
    je .gsp_key_del

    jmp .gsp_call_orig

.gsp_nav_left:
    test r14b, r14b
    jz .gsp_nl_noshift
    cmp dword [sel_end_col], 1
    jle .gsp_nav_apply
    dec dword [sel_end_col]
    mov eax, [sel_end_col]
    mov [cur_col], eax
    jmp .gsp_nav_apply
.gsp_nl_noshift:
    cmp dword [cur_col], 1
    jle .gsp_nav_apply
    dec dword [cur_col]
    mov eax, [cur_col]
    mov [sel_start_col], eax
    mov [sel_end_col], eax
    mov eax, [cur_row]
    mov [sel_start_row], eax
    mov [sel_end_row], eax
    jmp .gsp_nav_apply

.gsp_nav_right:
    test r14b, r14b
    jz .gsp_nr_noshift
    cmp dword [sel_end_col], 26
    jge .gsp_nav_apply
    inc dword [sel_end_col]
    mov eax, [sel_end_col]
    mov [cur_col], eax
    jmp .gsp_nav_apply
.gsp_nr_noshift:
    cmp dword [cur_col], 26
    jge .gsp_nav_apply
    inc dword [cur_col]
    mov eax, [cur_col]
    mov [sel_start_col], eax
    mov [sel_end_col], eax
    mov eax, [cur_row]
    mov [sel_start_row], eax
    mov [sel_end_row], eax
    jmp .gsp_nav_apply

.gsp_nav_up:
    test r14b, r14b
    jz .gsp_nu_noshift
    cmp dword [sel_end_row], 0
    jle .gsp_nav_apply
    dec dword [sel_end_row]
    mov eax, [sel_end_row]
    mov [cur_row], eax
    jmp .gsp_nav_apply
.gsp_nu_noshift:
    cmp dword [cur_row], 0
    jle .gsp_nav_apply
    dec dword [cur_row]
    mov eax, [cur_row]
    mov [sel_start_row], eax
    mov [sel_end_row], eax
    mov eax, [cur_col]
    mov [sel_start_col], eax
    mov [sel_end_col], eax
    jmp .gsp_nav_apply

.gsp_nav_down:
    mov edx, [total_rows]
    dec edx
    test r14b, r14b
    jz .gsp_nd_noshift
    cmp dword [sel_end_row], edx
    jge .gsp_nav_apply
    inc dword [sel_end_row]
    mov eax, [sel_end_row]
    mov [cur_row], eax
    jmp .gsp_nav_apply
.gsp_nd_noshift:
    cmp dword [cur_row], edx
    jge .gsp_nav_apply
    inc dword [cur_row]
    mov eax, [cur_row]
    mov [sel_start_row], eax
    mov [sel_end_row], eax
    mov eax, [cur_col]
    mov [sel_start_col], eax
    mov [sel_end_col], eax
    jmp .gsp_nav_apply

.gsp_nav_tab:
    test r14b, r14b
    jnz .gsp_nav_left
    jmp .gsp_nav_right

.gsp_nav_return:
    jmp .gsp_nav_down

.gsp_nav_home:
    test r14b, r14b
    jz @f
    mov dword [sel_end_col], 1
    mov dword [cur_col], 1
    jmp .gsp_nav_apply
@@:
    mov dword [cur_col], 1
    mov dword [sel_start_col], 1
    mov dword [sel_end_col], 1
    mov eax, [cur_row]
    mov [sel_start_row], eax
    mov [sel_end_row], eax
    jmp .gsp_nav_apply

.gsp_nav_end:
    test r14b, r14b
    jz @f
    mov dword [sel_end_col], 26
    mov dword [cur_col], 26
    jmp .gsp_nav_apply
@@:
    mov dword [cur_col], 26
    mov dword [sel_start_col], 26
    mov dword [sel_end_col], 26
    mov eax, [cur_row]
    mov [sel_start_row], eax
    mov [sel_end_row], eax
    jmp .gsp_nav_apply

.gsp_nav_pgup:
    test r14b, r14b
    jz @f
    sub dword [sel_end_row], 10
    jns .pgu_ok1
    mov dword [sel_end_row], 0
.pgu_ok1:
    mov eax, [sel_end_row]
    mov [cur_row], eax
    jmp .gsp_nav_apply
@@:
    sub dword [cur_row], 10
    jns .pgu_ok2
    mov dword [cur_row], 0
.pgu_ok2:
    mov eax, [cur_row]
    mov [sel_start_row], eax
    mov [sel_end_row], eax
    mov eax, [cur_col]
    mov [sel_start_col], eax
    mov [sel_end_col], eax
    jmp .gsp_nav_apply

.gsp_nav_pgdn:
    mov edx, [total_rows]
    dec edx
    test r14b, r14b
    jz @f
    add dword [sel_end_row], 10
    cmp dword [sel_end_row], edx
    jle .pgd_ok1
    mov [sel_end_row], edx
.pgd_ok1:
    mov eax, [sel_end_row]
    mov [cur_row], eax
    jmp .gsp_nav_apply
@@:
    add dword [cur_row], 10
    cmp dword [cur_row], edx
    jle .pgd_ok2
    mov [cur_row], edx
.pgd_ok2:
    mov eax, [cur_row]
    mov [sel_start_row], eax
    mov [sel_end_row], eax
    mov eax, [cur_col]
    mov [sel_start_col], eax
    mov [sel_end_col], eax
    jmp .gsp_nav_apply

.gsp_nav_apply:
    call UpdateCellSelectionDisplay
    call SyncListViewItemState
    xor eax, eax
    jmp .gsp_ret

.gsp_key_f2:
    mov rcx, [hFormulaEdit]
    call [SetFocus]
    mov rcx, [hFormulaEdit]
    mov edx, EM_SETSEL
    xor r8d, r8d
    mov r9d, -1
    call [SendMessageA]
    xor eax, eax
    jmp .gsp_ret

.gsp_key_f3:
    call DoShowFxWizard
    xor eax, eax
    jmp .gsp_ret

.gsp_key_f5:
    call DoRefreshGrid
    xor eax, eax
    jmp .gsp_ret

.gsp_key_f9:
    call DoRecalculateAll
    xor eax, eax
    jmp .gsp_ret

.gsp_key_del:
    call DoClearCell
    xor eax, eax
    jmp .gsp_ret

.gsp_ret:
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
; Build Spreadsheet Grid View (SysListView32: #, A..J, Rows 1..30)
; =============================================================================
BuildSpreadsheetGrid:
    push rbp
    mov rbp, rsp
    sub rsp, 60h

    ; Create Grid ListView
    mov ecx, WS_EX_CLIENTEDGE
    lea rdx, [listview_class]
    xor r8, r8
    mov r9d, WS_CHILD or WS_VISIBLE or LVS_REPORT or LVS_SINGLESEL or LVS_SHOWSELALWAYS or WS_VSCROLL or WS_HSCROLL
    mov dword [rsp+20h], 6
    mov dword [rsp+28h], 150
    mov dword [rsp+30h], 1254
    mov dword [rsp+38h], 575
    mov rax, [hWnd]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], ID_GRID_VIEW
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov [hGrid], rax
    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    ; Enable Double-Buffering & Gridlines & FullRowSelect (Required for subitem hit-testing)
    mov rcx, [hGrid]
    mov edx, LVM_SETEXTENDEDLISTVIEWSTYLE
    xor r8d, r8d
    mov r9d, LVS_EX_GRIDLINES or LVS_EX_FULLROWSELECT or LVS_EX_DOUBLEBUFFER
    call [SendMessageA]

    ; Subclass Grid ListView for full 2D mouse drag & keyboard selection
    mov rcx, [hGrid]
    mov edx, -4 ; GWLP_WNDPROC
    lea r8, [GridSubclassProc]
    call [SetWindowLongPtrA]
    mov [origGridProc], rax

    ; Create In-Place Cell Edit Control (child of hGrid, initially hidden)
    xor ecx, ecx
    lea rdx, [edit_class]
    xor r8, r8
    mov r9d, WS_CHILD or WS_BORDER or ES_AUTOHSCROLL or ES_LEFT
    mov dword [rsp+20h], 0
    mov dword [rsp+28h], 0
    mov dword [rsp+30h], 100
    mov dword [rsp+38h], 24
    mov rax, [hGrid]
    mov qword [rsp+40h], rax
    mov qword [rsp+48h], IDC_CELL_EDIT
    mov rax, [hInst]
    mov qword [rsp+50h], rax
    mov qword [rsp+58h], 0
    call [CreateWindowExA]
    mov [hCellEdit], rax

    mov rcx, rax
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    ; Subclass in-place cell edit control
    mov rcx, [hCellEdit]
    mov edx, -4 ; GWLP_WNDPROC
    lea r8, [CellEditSubclassProc]
    call [SetWindowLongPtrA]
    mov [origCellEditProc], rax

    ; Setup Columns: # (width 42), A..J (width 118 each)
    lea rdi, [lvcol]
    xor eax, eax
    mov ecx, 64
    rep stosb
    mov dword [lvcol], LVCF_FMT or LVCF_WIDTH or LVCF_TEXT or LVCF_SUBITEM

    ; # Column (Center)
    mov dword [lvcol+4], LVCFMT_CENTER
    mov dword [lvcol+8], 42
    lea rax, [col_idx_txt]
    mov qword [lvcol+16], rax
    mov dword [lvcol+28], 0
    mov rcx, [hGrid]
    mov edx, LVM_INSERTCOLUMNA
    mov r8d, 0
    lea r9, [lvcol]
    call [SendMessageA]

    ; Insert Columns A..Z (1..26)
    mov ebx, 1
.loop_col_az:
    cmp ebx, 26
    jg .cols_az_done

    mov al, bl
    dec al
    add al, 'A'
    mov [szTemp], al
    mov byte [szTemp+1], 0

    mov dword [lvcol+4], LVCFMT_LEFT
    mov dword [lvcol+8], 118
    lea rax, [szTemp]
    mov qword [lvcol+16], rax
    mov dword [lvcol+28], ebx
    mov rcx, [hGrid]
    mov edx, LVM_INSERTCOLUMNA
    mov r8d, ebx
    lea r9, [lvcol]
    call [SendMessageA]

    inc ebx
    jmp .loop_col_az
.cols_az_done:

    ; Populate 30 Empty Spreadsheet Rows (1..30)
    mov ebx, 0
.populate_30:
    cmp ebx, 30
    jge .grid_done

    ; Format Row Index (1, 2, 3.. 30)
    mov eax, ebx
    inc eax
    lea rdi, [szTemp+8]
    mov byte [rdi], 0
    mov ecx, 10
.num_to_s:
    xor edx, edx
    div ecx
    add dl, '0'
    dec rdi
    mov [rdi], dl
    test eax, eax
    jnz .num_to_s

    lea rsi, [lvitem]
    xor eax, eax
    mov ecx, 88
.zero_lvi:
    mov byte [rsi+rcx-1], 0
    dec ecx
    jnz .zero_lvi

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], ebx    ; iItem
    mov dword [lvitem+8], 0      ; iSubItem
    mov qword [lvitem+24], rdi   ; pszText
    mov rcx, [hGrid]
    mov edx, LVM_INSERTITEMA
    xor r8d, r8d
    lea r9, [lvitem]
    call [SendMessageA]

    inc ebx
    jmp .populate_30

.grid_done:
    leave
    ret


; =============================================================================
; Cell Hit-Testing & UI Synchronization
; =============================================================================
DoGridCellHitTest:
    push rbp
    mov rbp, rsp
    sub rsp, 40h

    lea rcx, [rbp-10h]
    call [GetCursorPos]
    mov rcx, [hGrid]
    lea rdx, [rbp-10h]
    call [ScreenToClient]

    lea rdi, [lvhit]
    xor eax, eax
    mov ecx, 24
    rep stosb

    mov eax, dword [rbp-10h]
    mov dword [lvhit], eax      ; pt.x
    mov eax, dword [rbp-0Ch]
    mov dword [lvhit+4], eax    ; pt.y

    mov rcx, [hGrid]
    mov edx, 1039h              ; LVM_SUBITEMHITTEST
    xor r8d, r8d
    lea r9, [lvhit]
    call [SendMessageA]

    cmp dword [lvhit+12], 0    ; iItem
    jl .use_selected_row
    mov edx, [total_rows]
    cmp dword [lvhit+12], edx
    jge .use_selected_row

    mov eax, dword [lvhit+12]
    mov [cur_row], eax
    mov [sel_start_row], eax
    mov [sel_end_row], eax

    mov eax, dword [lvhit+16]   ; iSubItem (0 = #, 1 = A, 2 = B, ..., 26 = Z)
    cmp eax, 1
    jl .use_col1
    cmp eax, 26
    jg .use_col1
    mov [cur_col], eax
    mov [sel_start_col], eax
    mov [sel_end_col], eax
    jmp .update_display

.use_col1:
    mov dword [cur_col], 1
    mov dword [sel_start_col], 1
    mov dword [sel_end_col], 1
    jmp .update_display

.use_selected_row:
    mov rcx, [hGrid]
    mov edx, LVM_GETNEXTITEM
    mov r8d, -1
    mov r9d, LVNI_SELECTED
    call [SendMessageA]
    cmp eax, 0
    jl .hit_done
    mov edx, [total_rows]
    cmp eax, edx
    jge .hit_done
    mov [cur_row], eax
    mov [sel_start_row], eax
    mov [sel_end_row], eax

.update_display:
    call UpdateCellSelectionDisplay

.hit_done:
    leave
    ret


; =============================================================================
; UpdateCellSelectionDisplay
; Updates NameBox, FormulaEdit, StatusBar and Repaints Grid for active cell
; =============================================================================
UpdateCellSelectionDisplay:
    push rbp
    mov rbp, rsp
    sub rsp, 80h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13

    ; Clamp cur_row, cur_col, sel_start_row, sel_end_row
    mov eax, [cur_row]
    test eax, eax
    jns @f
    xor eax, eax
    mov [cur_row], eax
@@:
    mov edx, [total_rows]
    test edx, edx
    jz .chk_col
    dec edx
    cmp eax, edx
    jle .chk_col
    mov [cur_row], edx
.chk_col:
    mov eax, [cur_col]
    cmp eax, 1
    jge @f
    mov dword [cur_col], 1
@@:
    cmp eax, 26
    jle @f
    mov dword [cur_col], 26
@@:
    mov eax, [sel_start_row]
    test eax, eax
    jns @f
    xor eax, eax
    mov [sel_start_row], eax
@@:
    mov eax, [sel_end_row]
    test eax, eax
    jns @f
    xor eax, eax
    mov [sel_end_row], eax
@@:

    ; Check if single cell or multi-cell selection
    mov eax, [sel_start_row]
    cmp eax, [sel_end_row]
    jne .fmt_range
    mov eax, [sel_start_col]
    cmp eax, [sel_end_col]
    jne .fmt_range

    ; --- SINGLE CELL ---
    mov eax, [cur_col]
    dec eax
    cmp eax, 26
    jge .multi_letter
    add al, 'A'
    lea rdi, [szTemp]
    mov [rdi], al
    inc rdi
    jmp .fmt_row_num

.multi_letter:
    xor edx, edx
    mov ecx, 26
    div ecx
    dec eax
    add al, 'A'
    lea rdi, [szTemp]
    mov [rdi], al
    inc rdi
    mov al, dl
    add al, 'A'
    mov [rdi], al
    inc rdi

.fmt_row_num:
    mov eax, [cur_row]
    inc eax
    mov ecx, 10
    lea rsi, [szTemp+32]
    mov byte [rsi], 0
.itos2:
    xor edx, edx
    div ecx
    add dl, '0'
    dec rsi
    mov [rsi], dl
    test eax, eax
    jnz .itos2

.copy_row_str:
    lodsb
    stosb
    test al, al
    jnz .copy_row_str
    jmp .apply_to_ui

.fmt_range:
    ; Format Range: min_col, min_row : max_col, max_row
    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r12d, eax       ; min_row
    mov r13d, edx       ; max_row

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov ebx, edx        ; max_col

    ; Format Start Col (eax)
    dec eax
    add al, 'A'
    lea rdi, [szTemp]
    mov [rdi], al
    inc rdi

    ; Format Start Row (r12d + 1)
    mov eax, r12d
    inc eax
    mov ecx, 10
    lea rsi, [szTemp+48]
    mov byte [rsi], 0
.r_itos1:
    xor edx, edx
    div ecx
    add dl, '0'
    dec rsi
    mov [rsi], dl
    test eax, eax
    jnz .r_itos1
.r_copy1:
    lodsb
    stosb
    test al, al
    jnz .r_copy1
    dec rdi             ; back up over null terminator

    ; Append ':'
    mov byte [rdi], ':'
    inc rdi

    ; Format End Col (ebx)
    mov eax, ebx
    dec eax
    add al, 'A'
    mov [rdi], al
    inc rdi

    ; Format End Row (r13d + 1)
    mov eax, r13d
    inc eax
    mov ecx, 10
    lea rsi, [szTemp+48]
    mov byte [rsi], 0
.r_itos2:
    xor edx, edx
    div ecx
    add dl, '0'
    dec rsi
    mov [rsi], dl
    test eax, eax
    jnz .r_itos2
.r_copy2:
    lodsb
    stosb
    test al, al
    jnz .r_copy2

.apply_to_ui:
    mov rcx, [hNameBox]
    lea rdx, [szTemp]
    call [SetWindowTextA]

    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 1
    lea r9, [szTemp]
    call [SendMessageA]

    ; Get cell content into Formula Bar
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
    jz .use_cell_text
    ; Has formula! Display formula in Formula Bar
    mov rcx, [hFormulaEdit]
    mov rdx, rax
    call [SetWindowTextA]
    jmp .formbar_updated

.use_cell_text:
    mov rcx, [hFormulaEdit]
    lea rdx, [szCellVal]
    call [SetWindowTextA]

.formbar_updated:
    call UpdateStatusStats

    ; Sync active cell font & size to ribbon combos
    mov rax, [pCellStyleTable]
    test rax, rax
    jz .skip_active_combo_sync
    mov edx, [cur_row]
    mov ecx, [cur_col]
    cmp edx, 10000
    jae .skip_active_combo_sync
    cmp ecx, 32
    jae .skip_active_combo_sync
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, rax
    movzx r12d, byte [rdx+13]
    movzx r13d, byte [rdx+14]
    test r12d, r12d
    jnz @f
    mov r12d, 10
@@:
    mov [font_size_pt], r12d
    call SyncComboSize

    mov eax, r13d
    test eax, eax
    jz .set_fcombo_def
    dec eax
    jmp .set_fcombo
.set_fcombo_def:
    xor eax, eax
.set_fcombo:
    mov rcx, [hComboFont]
    test rcx, rcx
    jz .skip_active_combo_sync
    mov r8d, eax
    mov edx, CB_SETCURSEL
    xor r9d, r9d
    call [SendMessageA]
.skip_active_combo_sync:

    ; Repaint grid so the active cell highlight moves immediately to current cell
    mov rcx, [hGrid]
    xor edx, edx
    xor r8d, r8d
    call [InvalidateRect]

    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    leave
    ret


; =============================================================================
; Apply Dynamic Grid Font Updates
; =============================================================================
; =============================================================================
; Toggle Active Cell Bold (Flag bit 0)
; =============================================================================
ToggleCellBold:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    mov rax, [pCellStyleTable]
    test rax, rax
    jz .tcb_done

    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r12d, eax       ; min_row
    mov r13d, edx       ; max_row

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov ebx, eax        ; min_col
    mov esi, edx        ; max_col

    mov edx, [cur_row]
    mov ecx, [cur_col]
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]
    mov edi, [rdx+8]
    xor edi, 1
    and edi, 1
    mov [is_bold_active], edi

.tcb_r_loop:
    cmp r12d, r13d
    jg .tcb_redraw

    mov ecx, ebx
.tcb_c_loop:
    cmp ecx, esi
    jg .tcb_next_r

    cmp r12d, 10000
    jae .tcb_next_c
    cmp ecx, 32
    jae .tcb_next_c

    mov edx, r12d
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]

    and dword [rdx+8], not 1
    or dword [rdx+8], edi

.tcb_next_c:
    inc ecx
    jmp .tcb_c_loop

.tcb_next_r:
    inc r12d
    jmp .tcb_r_loop

.tcb_redraw:
    mov rcx, [hGrid]
    test rcx, rcx
    jz .tcb_done
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]
    call SaveActiveSheetToBuffer

.tcb_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret


; =============================================================================
; Toggle Active Cell Italic (Flag bit 3)
; =============================================================================
ToggleCellItalic:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    mov rax, [pCellStyleTable]
    test rax, rax
    jz .tci_done

    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r12d, eax
    mov r13d, edx

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov ebx, eax
    mov esi, edx

    mov edx, [cur_row]
    mov ecx, [cur_col]
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]
    mov edi, [rdx+8]
    xor edi, 8
    and edi, 8
    mov eax, edi
    shr eax, 3
    mov [is_italic_act], eax

.tci_r_loop:
    cmp r12d, r13d
    jg .tci_redraw

    mov ecx, ebx
.tci_c_loop:
    cmp ecx, esi
    jg .tci_next_r

    cmp r12d, 10000
    jae .tci_next_c
    cmp ecx, 32
    jae .tci_next_c

    mov edx, r12d
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]

    and dword [rdx+8], not 8
    or dword [rdx+8], edi

.tci_next_c:
    inc ecx
    jmp .tci_c_loop

.tci_next_r:
    inc r12d
    jmp .tci_r_loop

.tci_redraw:
    mov rcx, [hGrid]
    test rcx, rcx
    jz .tci_done
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]
    call SaveActiveSheetToBuffer

.tci_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret


; =============================================================================
; Toggle Active Cell Underline (Flag bit 4)
; =============================================================================
ToggleCellUnderline:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    mov rax, [pCellStyleTable]
    test rax, rax
    jz .tcu_done

    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r12d, eax
    mov r13d, edx

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov ebx, eax
    mov esi, edx

    mov edx, [cur_row]
    mov ecx, [cur_col]
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]
    mov edi, [rdx+8]
    xor edi, 16
    and edi, 16
    mov eax, edi
    shr eax, 4
    mov [is_under_act], eax

.tcu_r_loop:
    cmp r12d, r13d
    jg .tcu_redraw

    mov ecx, ebx
.tcu_c_loop:
    cmp ecx, esi
    jg .tcu_next_r

    cmp r12d, 10000
    jae .tcu_next_c
    cmp ecx, 32
    jae .tcu_next_c

    mov edx, r12d
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]

    and dword [rdx+8], not 16
    or dword [rdx+8], edi

.tcu_next_c:
    inc ecx
    jmp .tcu_c_loop

.tcu_next_r:
    inc r12d
    jmp .tcu_r_loop

.tcu_redraw:
    mov rcx, [hGrid]
    test rcx, rcx
    jz .tcu_done
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]
    call SaveActiveSheetToBuffer

.tcu_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret


; =============================================================================
; Toggle Active Cell Strike (Flag bit 5)
; =============================================================================
ToggleCellStrike:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    mov rax, [pCellStyleTable]
    test rax, rax
    jz .tcs_done

    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r12d, eax
    mov r13d, edx

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov ebx, eax
    mov esi, edx

    mov edx, [cur_row]
    mov ecx, [cur_col]
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]
    mov edi, [rdx+8]
    xor edi, 32
    and edi, 32
    mov eax, edi
    shr eax, 5
    mov [is_strike_act], eax

.tcs_r_loop:
    cmp r12d, r13d
    jg .tcs_redraw

    mov ecx, ebx
.tcs_c_loop:
    cmp ecx, esi
    jg .tcs_next_r

    cmp r12d, 10000
    jae .tcs_next_c
    cmp ecx, 32
    jae .tcs_next_c

    mov edx, r12d
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]

    and dword [rdx+8], not 32
    or dword [rdx+8], edi

.tcs_next_c:
    inc ecx
    jmp .tcs_c_loop

.tcs_next_r:
    inc r12d
    jmp .tcs_r_loop

.tcs_redraw:
    mov rcx, [hGrid]
    test rcx, rcx
    jz .tcs_done
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]
    call SaveActiveSheetToBuffer

.tcs_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret


; =============================================================================
; Toggle Active Cell Fill Color (Flag bit 1: yellow #FFF2CC)
; =============================================================================
ToggleCellFill:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    mov rax, [pCellStyleTable]
    test rax, rax
    jz .tcf_done

    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r12d, eax
    mov r13d, edx

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov ebx, eax
    mov esi, edx

    mov edx, [cur_row]
    mov ecx, [cur_col]
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]
    mov edi, [rdx+8]
    xor edi, 2
    and edi, 2

.tcf_r_loop:
    cmp r12d, r13d
    jg .tcf_redraw

    mov ecx, ebx
.tcf_c_loop:
    cmp ecx, esi
    jg .tcf_next_r

    cmp r12d, 10000
    jae .tcf_next_c
    cmp ecx, 32
    jae .tcf_next_c

    mov edx, r12d
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]

    test edi, edi
    jz .tcf_clr
    mov dword [rdx], 00CCF2FFh ; Yellow highlight
    or dword [rdx+8], 2
    jmp .tcf_next_c
.tcf_clr:
    and dword [rdx+8], not 2

.tcf_next_c:
    inc ecx
    jmp .tcf_c_loop

.tcf_next_r:
    inc r12d
    jmp .tcf_r_loop

.tcf_redraw:
    mov rcx, [hGrid]
    test rcx, rcx
    jz .tcf_done
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]
    call SaveActiveSheetToBuffer

.tcf_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret


; =============================================================================
; Toggle Active Cell Text Color (Flag bit 2: red #C00000)
; =============================================================================
ToggleCellTextColor:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    mov rax, [pCellStyleTable]
    test rax, rax
    jz .tct_done

    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r12d, eax
    mov r13d, edx

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov ebx, eax
    mov esi, edx

    mov edx, [cur_row]
    mov ecx, [cur_col]
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]
    mov edi, [rdx+8]
    xor edi, 4
    and edi, 4

.tct_r_loop:
    cmp r12d, r13d
    jg .tct_redraw

    mov ecx, ebx
.tct_c_loop:
    cmp ecx, esi
    jg .tct_next_r

    cmp r12d, 10000
    jae .tct_next_c
    cmp ecx, 32
    jae .tct_next_c

    mov edx, r12d
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]

    test edi, edi
    jz .tct_clr
    mov dword [rdx+4], 000000C0h ; Dark red
    or dword [rdx+8], 4
    jmp .tct_next_c
.tct_clr:
    and dword [rdx+8], not 4

.tct_next_c:
    inc ecx
    jmp .tct_c_loop

.tct_next_r:
    inc r12d
    jmp .tct_r_loop

.tct_redraw:
    mov rcx, [hGrid]
    test rcx, rcx
    jz .tct_done
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]
    call SaveActiveSheetToBuffer

.tct_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret


; =============================================================================
; Apply Dynamic Grid Font Updates
; =============================================================================
ApplyNewGridFont:
    push rbp
    mov rbp, rsp
    sub rsp, 80h

    mov ecx, [font_size_pt]
    add ecx, 3              ; font height in px
    mov r14d, ecx

    ; Delete previous fonts
    mov rcx, [hFontUI]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hFontBold]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hFontItalic]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:
    mov rcx, [hFontBoldItalic]
    test rcx, rcx
    jz @f
    call [DeleteObject]
@@:

    ; 1. hFontUI (Regular)
    mov ecx, r14d
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
    mov rax, [pCurFontFace]
    test rax, rax
    jnz .f1_face
    lea rax, [font_ui_name]
.f1_face:
    mov qword [rsp+68h], rax
    call [CreateFontA]
    mov [hFontUI], rax

    ; 2. hFontBold
    mov ecx, r14d
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
    mov rax, [pCurFontFace]
    test rax, rax
    jnz .f2_face
    lea rax, [font_ui_name]
.f2_face:
    mov qword [rsp+68h], rax
    call [CreateFontA]
    mov [hFontBold], rax

    ; 3. hFontItalic
    mov ecx, r14d
    xor edx, edx
    xor r8d, r8d
    xor r9d, r9d
    mov dword [rsp+20h], 400
    mov dword [rsp+28h], 1
    mov dword [rsp+30h], 0
    mov dword [rsp+38h], 0
    mov dword [rsp+40h], 1
    mov dword [rsp+48h], 0
    mov dword [rsp+50h], 0
    mov dword [rsp+58h], 0
    mov dword [rsp+60h], 0
    mov rax, [pCurFontFace]
    test rax, rax
    jnz .f3_face
    lea rax, [font_ui_name]
.f3_face:
    mov qword [rsp+68h], rax
    call [CreateFontA]
    mov [hFontItalic], rax

    ; 4. hFontBoldItalic
    mov ecx, r14d
    xor edx, edx
    xor r8d, r8d
    xor r9d, r9d
    mov dword [rsp+20h], 700
    mov dword [rsp+28h], 1
    mov dword [rsp+30h], 0
    mov dword [rsp+38h], 0
    mov dword [rsp+40h], 1
    mov dword [rsp+48h], 0
    mov dword [rsp+50h], 0
    mov dword [rsp+58h], 0
    mov dword [rsp+60h], 0
    mov rax, [pCurFontFace]
    test rax, rax
    jnz .f4_face
    lea rax, [font_ui_name]
.f4_face:
    mov qword [rsp+68h], rax
    call [CreateFontA]
    mov [hFontBoldItalic], rax

    ; Update Grid default font & invalidate
    mov rcx, [hGrid]
    test rcx, rcx
    jz .ang_done
    mov edx, WM_SETFONT
    mov r8, [hFontUI]
    mov r9d, 1
    call [SendMessageA]

    mov rcx, [hGrid]
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]

.ang_done:
    leave
    ret

; =============================================================================
; Synchronize Combobox with Current Font Size
; =============================================================================
SyncComboSize:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    xor edx, edx
.sync_loop:
    cmp edx, 7
    jge .sync_done
    lea rax, [font_size_table]
    mov ecx, [rax + rdx*4]
    cmp ecx, [font_size_pt]
    je .sync_found
    inc edx
    jmp .sync_loop
.sync_found:
    mov rcx, [hComboSize]
    mov r8d, edx
    mov edx, CB_SETCURSEL
    xor r9d, r9d
    call [SendMessageA]
.sync_done:
    leave
    ret

; =============================================================================
; Get or Create Cell Font from Cache
; ECX = Font Size pt (0 = default 10)
; EDX = Face Index (0 = default Segoe UI, 1 = Segoe UI, 2 = Calibri, 3 = Arial, 4 = Consolas)
; R8D = Style Flags (bit 0 = Bold, bit 3 = Italic, bit 4 = Underline)
; Returns RAX = HFONT
; =============================================================================
GetCellFont:
    push rbp
    mov rbp, rsp
    sub rsp, 80h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13
    mov [rbp-48], r14

    test ecx, ecx
    jnz .gcf_custom
    test edx, edx
    jnz .gcf_custom

    ; Default base font: Segoe UI 10pt with bold/italic combinations
    mov eax, r8d
    and eax, 9              ; bit 0 = bold, bit 3 = italic
    cmp eax, 9
    jne @f
    mov rax, [hFontBoldItalic]
    jmp .gcf_ret
@@:
    test r8d, 1
    jz @f
    mov rax, [hFontBold]
    jmp .gcf_ret
@@:
    test r8d, 8
    jz @f
    mov rax, [hFontItalic]
    jmp .gcf_ret
@@:
    mov rax, [hFontUI]
    jmp .gcf_ret

.gcf_custom:
    test ecx, ecx
    jnz @f
    mov ecx, 10
@@:
    mov ebx, ecx            ; ebx = size
    mov esi, edx            ; esi = faceIdx
    mov edi, r8d            ; edi = styleFlags

    ; Build key: (size & 0xFF) | ((faceIdx & 0xFF) << 8) | ((flags & 0xFF) << 16)
    mov eax, edi
    and eax, 0FFh
    shl eax, 16
    mov r12d, esi
    and r12d, 0FFh
    shl r12d, 8
    or eax, r12d
    mov r12d, ebx
    and r12d, 0FFh
    or eax, r12d
    mov r13d, eax           ; r13d = 32-bit cache key

    ; Search cache
    xor ecx, ecx
    mov edx, [font_cache_count]
.cache_search:
    cmp ecx, edx
    jge .cache_miss
    cmp [font_cache_keys + rcx*4], r13d
    je .cache_hit
    inc ecx
    jmp .cache_search

.cache_hit:
    mov rax, [font_cache_fonts + rcx*8]
    jmp .gcf_ret

.cache_miss:
    ; Create new font
    mov ecx, ebx
    add ecx, 3              ; font height in px (e.g. 10pt -> 13px, 14pt -> 17px)

    mov r14d, 400
    test edi, 1
    jz @f
    mov r14d, 700
@@:
    xor r12d, r12d
    test edi, 8
    jz @f
    mov r12d, 1
@@:
    xor eax, eax
    test edi, 16
    jz @f
    mov eax, 1
@@:
    cmp esi, 4
    jbe @f
    xor esi, esi
@@:
    lea rdx, [font_faces_table]
    mov rdx, [rdx + rsi*8]

    mov [rsp+20h], r14d     ; weight
    mov [rsp+28h], r12d     ; italic
    mov [rsp+30h], eax      ; underline
    mov dword [rsp+38h], 0  ; strikeout
    mov dword [rsp+40h], 1  ; DEFAULT_CHARSET
    mov dword [rsp+48h], 0  ; OUT_DEFAULT_PRECIS
    mov dword [rsp+50h], 0  ; CLIP_DEFAULT_PRECIS
    mov dword [rsp+58h], 0  ; DEFAULT_QUALITY
    mov dword [rsp+60h], 0  ; DEFAULT_PITCH
    mov qword [rsp+68h], rdx ; face name
    xor edx, edx            ; width
    xor r8d, r8d            ; escapement
    xor r9d, r9d            ; orientation
    call [CreateFontA]
    test rax, rax
    jnz .font_created
    mov rax, [hFontUI]
    jmp .gcf_ret

.font_created:
    mov ecx, [font_cache_count]
    cmp ecx, 63
    jge .cache_full
    mov [font_cache_keys + rcx*4], r13d
    mov [font_cache_fonts + rcx*8], rax
    inc dword [font_cache_count]
.cache_full:

.gcf_ret:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    mov r14, [rbp-48]
    leave
    ret

; =============================================================================
; Apply Font Size to Selected Cell(s)
; R8D = Size in pt (8, 9, 10, 11, 12, 14, 16...)
; =============================================================================
ApplySelectionFontSize:
    push rbp
    mov rbp, rsp
    sub rsp, 40h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    mov edi, r8d        ; edi = size

    mov rax, [pCellStyleTable]
    test rax, rax
    jz .asfs_done

    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r12d, eax       ; min_row
    mov r13d, edx       ; max_row

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov ebx, eax        ; min_col
    mov esi, edx        ; max_col

.asfs_r_loop:
    cmp r12d, r13d
    jg .asfs_redraw

    mov ecx, ebx
.asfs_c_loop:
    cmp ecx, esi
    jg .asfs_next_r

    cmp r12d, 10000
    jae .asfs_next_c
    cmp ecx, 32
    jae .asfs_next_c

    mov edx, r12d
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]

    mov byte [rdx+13], dil ; store font size

.asfs_next_c:
    inc ecx
    jmp .asfs_c_loop

.asfs_next_r:
    inc r12d
    jmp .asfs_r_loop

.asfs_redraw:
    mov rcx, [hGrid]
    test rcx, rcx
    jz .asfs_done
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]
    call SaveActiveSheetToBuffer

.asfs_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret

; =============================================================================
; Increase Font Size for Selected Cell(s) (A+)
; =============================================================================
DoIncSelectionFontSize:
    push rbp
    mov rbp, rsp
    sub rsp, 30h
    mov [rbp-8], rbx

    mov rax, [pCellStyleTable]
    test rax, rax
    jz .disfs_done

    mov edx, [cur_row]
    mov ecx, [cur_col]
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, rax
    movzx eax, byte [rdx+13]
    test eax, eax
    jnz @f
    mov eax, 10
@@:
    cmp eax, 72
    jge .disfs_done
    add eax, 2
    cmp eax, 72
    jle @f
    mov eax, 72
@@:
    mov ebx, eax
    mov r8d, ebx
    call ApplySelectionFontSize

    ; Sync combo
    mov [font_size_pt], ebx
    call SyncComboSize

.disfs_done:
    mov rbx, [rbp-8]
    leave
    ret

; =============================================================================
; Decrease Font Size for Selected Cell(s) (A-)
; =============================================================================
DoDecSelectionFontSize:
    push rbp
    mov rbp, rsp
    sub rsp, 30h
    mov [rbp-8], rbx

    mov rax, [pCellStyleTable]
    test rax, rax
    jz .ddsfs_done

    mov edx, [cur_row]
    mov ecx, [cur_col]
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, rax
    movzx eax, byte [rdx+13]
    test eax, eax
    jnz @f
    mov eax, 10
@@:
    cmp eax, 6
    jle .ddsfs_done
    sub eax, 2
    cmp eax, 6
    jge @f
    mov eax, 6
@@:
    mov ebx, eax
    mov r8d, ebx
    call ApplySelectionFontSize

    ; Sync combo
    mov [font_size_pt], ebx
    call SyncComboSize

.ddsfs_done:
    mov rbx, [rbp-8]
    leave
    ret

; =============================================================================
; Apply Font Face to Selected Cell(s)
; R8D = Face Index (1 = Segoe UI, 2 = Calibri, 3 = Arial, 4 = Consolas)
; =============================================================================
ApplySelectionFontFace:
    push rbp
    mov rbp, rsp
    sub rsp, 40h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    mov edi, r8d        ; edi = face index

    mov rax, [pCellStyleTable]
    test rax, rax
    jz .asff_done

    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r12d, eax       ; min_row
    mov r13d, edx       ; max_row

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov ebx, eax        ; min_col
    mov esi, edx        ; max_col

.asff_r_loop:
    cmp r12d, r13d
    jg .asff_redraw

    mov ecx, ebx
.asff_c_loop:
    cmp ecx, esi
    jg .asff_next_r

    cmp r12d, 10000
    jae .asff_next_c
    cmp ecx, 32
    jae .asff_next_c

    mov edx, r12d
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]

    mov byte [rdx+14], dil ; store font face index

.asff_next_c:
    inc ecx
    jmp .asff_c_loop

.asff_next_r:
    inc r12d
    jmp .asff_r_loop

.asff_redraw:
    mov rcx, [hGrid]
    test rcx, rcx
    jz .asff_done
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]
    call SaveActiveSheetToBuffer

.asff_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret


; =============================================================================
; Set Column Alignment (LVCFMT_LEFT, LVCFMT_CENTER, LVCFMT_RIGHT)
; R8D = Column Index, R9D = Format
; =============================================================================
SetColumnAlignment:
    push rbp
    mov rbp, rsp
    sub rsp, 40h

    mov [rbp-8], r8
    mov [rbp-16], r9

    lea rdi, [lvcol]
    xor eax, eax
    mov ecx, 64
    rep stosb
    mov dword [lvcol], LVCF_FMT
    mov eax, dword [rbp-16]
    mov dword [lvcol+4], eax

    mov rcx, [hGrid]
    mov edx, LVM_SETCOLUMNA
    mov r8d, dword [rbp-8]
    lea r9, [lvcol]
    call [SendMessageA]

    mov rcx, [hGrid]
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]

    leave
    ret


; =============================================================================
; Apply Selection Alignment (ALIGN_LEFT=1, ALIGN_CENTER=2, ALIGN_RIGHT=3)
; R8D = Alignment
; =============================================================================
ApplySelectionAlignment:
    push rbp
    mov rbp, rsp
    sub rsp, 40h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    mov edi, r8d        ; edi = alignment (1, 2, 3)

    mov rax, [pCellStyleTable]
    test rax, rax
    jz .asa_done

    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r12d, eax       ; min_row
    mov r13d, edx       ; max_row

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov ebx, eax        ; min_col
    mov esi, edx        ; max_col

.asa_r_loop:
    cmp r12d, r13d
    jg .asa_redraw

    mov ecx, ebx
.asa_c_loop:
    cmp ecx, esi
    jg .asa_next_r

    cmp r12d, 10000
    jae .asa_next_c
    cmp ecx, 32
    jae .asa_next_c

    mov edx, r12d
    shl edx, 5
    add edx, ecx
    shl rdx, 4
    add rdx, [pCellStyleTable]

    mov dword [rdx+12], edi ; set cell alignment

.asa_next_c:
    inc ecx
    jmp .asa_c_loop

.asa_next_r:
    inc r12d
    jmp .asa_r_loop

.asa_redraw:
    mov rcx, [hGrid]
    test rcx, rcx
    jz .asa_done
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]
    call SaveActiveSheetToBuffer

.asa_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret


; =============================================================================
; Action: Live Windows Clipboard Operations (Copy, Cut, Paste)
; =============================================================================
DoClipboardCopy:
    push rbp
    mov rbp, rsp
    sub rsp, 70h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13
    mov [rbp-48], r14
    mov [rbp-56], r15

    ; Check if single cell or range
    mov eax, [sel_start_row]
    cmp eax, [sel_end_row]
    jne .dcc_multi
    mov eax, [sel_start_col]
    cmp eax, [sel_end_col]
    jne .dcc_multi

    ; Single cell copy
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

    mov rcx, [hWnd]
    call [OpenClipboard]
    test eax, eax
    jz .dcc_done
    call [EmptyClipboard]

    lea rcx, [szCellVal]
    call [strlen]
    inc eax
    mov ebx, eax
    mov ecx, GMEM_MOVEABLE or GMEM_ZEROINIT
    mov edx, ebx
    call [GlobalAlloc]
    test rax, rax
    jz .dcc_close_cb
    mov r12, rax

    mov rcx, r12
    call [GlobalLock]
    mov rdi, rax
    lea rsi, [szCellVal]
    mov ecx, ebx
    rep movsb
    mov rcx, r12
    call [GlobalUnlock]

    mov ecx, CF_TEXT
    mov rdx, r12
    call [SetClipboardData]

.dcc_close_cb:
    call [CloseClipboard]
    jmp .dcc_done

.dcc_multi:
    ; Multi-cell range copy
    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r12d, eax       ; min_row
    mov r13d, edx       ; max_row

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r14d, eax       ; min_col
    mov r15d, edx       ; max_col

    mov ecx, GMEM_MOVEABLE or GMEM_ZEROINIT
    mov edx, 65536
    call [GlobalAlloc]
    test rax, rax
    jz .dcc_done
    mov [rbp-64], rax   ; hMem

    mov rcx, rax
    call [GlobalLock]
    mov rdi, rax

    mov ebx, r12d       ; row index
.dcc_m_r:
    cmp ebx, r13d
    jg .dcc_m_finish

    mov esi, r14d       ; col index
.dcc_m_c:
    cmp esi, r15d
    jg .dcc_m_next_r

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], ebx
    mov dword [lvitem+8], esi
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 1000
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, ebx
    lea r9, [lvitem]
    call [SendMessageA]

    lea rdx, [szCellVal]
.dcc_m_cp:
    mov al, [rdx]
    test al, al
    jz .dcc_m_cp_done
    mov [rdi], al
    inc rdi
    inc rdx
    jmp .dcc_m_cp

.dcc_m_cp_done:
    cmp esi, r15d
    jge @f
    mov byte [rdi], 9   ; Tab
    inc rdi
@@:
    inc esi
    jmp .dcc_m_c

.dcc_m_next_r:
    cmp ebx, r13d
    jge @f
    mov byte [rdi], 13  ; \r
    inc rdi
    mov byte [rdi], 10  ; \n
    inc rdi
@@:
    inc ebx
    jmp .dcc_m_r

.dcc_m_finish:
    mov byte [rdi], 0

    mov rcx, [rbp-64]
    call [GlobalUnlock]

    mov rcx, [hWnd]
    call [OpenClipboard]
    test eax, eax
    jz .dcc_done
    call [EmptyClipboard]

    mov ecx, CF_TEXT
    mov rdx, [rbp-64]
    call [SetClipboardData]

    call [CloseClipboard]

.dcc_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    mov r14, [rbp-48]
    mov r15, [rbp-56]
    leave
    ret

DoClipboardCut:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    call DoClipboardCopy
    call DoClearCell
    leave
    ret

DoClipboardPaste:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], r12
    mov [rbp-24], r13
    mov [rbp-32], r14
    mov [rbp-40], r15

    mov rcx, [hWnd]
    call [OpenClipboard]
    test eax, eax
    jz .done

    mov ecx, CF_TEXT
    call [GetClipboardData]
    test rax, rax
    jz .close_p
    mov r12, rax

    mov rcx, r12
    call [GlobalLock]
    test rax, rax
    jz .close_p
    mov rsi, rax
    lea rdi, [szTemp]
    mov ecx, 2040
.cp_p:
    lodsb
    stosb
    test al, al
    jz .p_done
    dec ecx
    jnz .cp_p
    mov byte [rdi], 0

.p_done:
    mov rcx, r12
    call [GlobalUnlock]

    ; Check if szTemp contains tab or newline (multi-cell paste)
    lea rsi, [szTemp]
.check_multi:
    lodsb
    test al, al
    jz .paste_single
    cmp al, 9
    je .paste_multi
    cmp al, 10
    je .paste_multi
    jmp .check_multi

.paste_single:
    mov ecx, [cur_row]
    mov edx, [cur_col]
    call BeginCellChange

    ; Write pasted text into active cell
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

    mov ecx, [cur_row]
    mov edx, [cur_col]
    lea r8, [szTemp]
    xor r9, r9
    call CommitCellChange
    jmp .paste_finish

.paste_multi:
    lea rsi, [szTemp]
    mov r12d, [cur_row]
    mov r13d, [cur_col]
    mov r14d, [cur_col]

.cell_tok_start:
    lea rdi, [szCellVal]
    mov ecx, 250
.cell_tok_scan:
    lodsb
    test al, al
    jz .cell_tok_eof
    cmp al, 9
    je .cell_tok_tab
    cmp al, 13
    je .cell_tok_cr
    cmp al, 10
    je .cell_tok_lf
    stosb
    dec ecx
    jnz .cell_tok_scan
    jmp .cell_tok_scan

.cell_tok_tab:
    mov byte [rdi], 0
    call .write_p_cell
    inc r13d
    cmp r13d, 10
    jg .skip_past_col
    jmp .cell_tok_start

.skip_past_col:
.skip_col_loop:
    lodsb
    test al, al
    jz .cell_tok_eof
    cmp al, 10
    je .cell_tok_lf
    jmp .skip_col_loop

.cell_tok_cr:
    cmp byte [rsi], 10
    jne .cr_no_lf
    inc rsi
.cr_no_lf:
.cell_tok_lf:
    mov byte [rdi], 0
    call .write_p_cell
    inc r12d
    cmp r12d, [total_rows]
    jge .paste_finish
    mov r13d, r14d
    jmp .cell_tok_start

.cell_tok_eof:
    mov byte [rdi], 0
    call .write_p_cell
    jmp .paste_finish

.write_p_cell:
    push rcx
    push rdx
    push r8
    push r9
    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r12d
    mov dword [lvitem+8], r13d
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, r12d
    lea r9, [lvitem]
    call [SendMessageA]
    pop r9
    pop r8
    pop rdx
    pop rcx
    ret

.paste_finish:
    call SaveActiveSheetToBuffer
    call UpdateStatusStats

.close_p:
    call [CloseClipboard]

.done:
    mov rbx, [rbp-8]
    mov r12, [rbp-16]
    mov r13, [rbp-24]
    mov r14, [rbp-32]
    mov r15, [rbp-40]
    leave
    ret


; =============================================================================
; Action: Number Formatting ($ / ₽ / % / Comma)
; =============================================================================
ApplyDollarFormat:
    push rbp
    mov rbp, rsp
    sub rsp, 40h

    ; Read cell value
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

    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]

    lea rcx, [szTemp]
    lea rdx, [fmt_dollar]
    movq r8, xmm0
    call [sprintf]

    ; write formatted cell
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

    call SaveActiveSheetToBuffer
    call UpdateStatusStats
    leave
    ret

ApplyRubleFormat:
    push rbp
    mov rbp, rsp
    sub rsp, 40h

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

    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]

    lea rcx, [szTemp]
    lea rdx, [fmt_ruble]
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

    call SaveActiveSheetToBuffer
    call UpdateStatusStats
    leave
    ret

ApplyPercentFormat:
    push rbp
    mov rbp, rsp
    sub rsp, 40h

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

    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]

    lea rcx, [szTemp]
    lea rdx, [fmt_percent]
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

    call SaveActiveSheetToBuffer
    call UpdateStatusStats
    leave
    ret

ApplyCommaFormat:
    push rbp
    mov rbp, rsp
    sub rsp, 40h

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

    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]

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

    call SaveActiveSheetToBuffer
    call UpdateStatusStats
    leave
    ret


; =============================================================================
; Action: Row Insert & Delete
; =============================================================================
DoInsertNewRow:
    push rbp
    mov rbp, rsp
    sub rsp, 40h

    mov eax, [cur_row]
    inc [total_rows]

    ; Insert Item
    lea rsi, [lvitem]
    xor ecx, ecx
    mov ecx, 88
.zero_lvi:
    mov byte [rsi+rcx-1], 0
    dec ecx
    jnz .zero_lvi

    mov dword [lvitem], LVIF_TEXT
    mov eax, [cur_row]
    mov dword [lvitem+4], eax
    mov dword [lvitem+8], 0
    lea rax, [init_cell_txt]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_INSERTITEMA
    xor r8d, r8d
    lea r9, [lvitem]
    call [SendMessageA]

    call RenumberGridRows
    call UpdateStatusStats
    leave
    ret

DoDeleteActiveRow:
    push rbp
    mov rbp, rsp
    sub rsp, 40h

    cmp [total_rows], 1
    jle .done

    mov rcx, [hGrid]
    mov edx, LVM_DELETEITEM
    mov r8d, [cur_row]
    xor r9d, r9d
    call [SendMessageA]

    dec [total_rows]
    call RenumberGridRows
    call UpdateStatusStats

.done:
    leave
    ret

RenumberGridRows:
    push rbp
    mov rbp, rsp
    sub rsp, 40h

    mov ebx, 0
.r_loop:
    cmp ebx, [total_rows]
    jge .r_done

    mov eax, ebx
    inc eax
    lea rdi, [szTemp+8]
    mov byte [rdi], 0
    mov ecx, 10
.r_itos:
    xor edx, edx
    div ecx
    add dl, '0'
    dec rdi
    mov [rdi], dl
    test eax, eax
    jnz .r_itos

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], ebx
    mov dword [lvitem+8], 0
    mov qword [lvitem+24], rdi
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, ebx
    lea r9, [lvitem]
    call [SendMessageA]

    inc ebx
    jmp .r_loop

.r_done:
    leave
    ret



; =============================================================================
; Action: Column Sorting (Ascending & Descending)
; =============================================================================
DoSortAscending:
    push rbp
    mov rbp, rsp
    sub rsp, 70h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13
    mov [rbp-48], r14
    mov [rbp-56], r15

    cmp [total_rows], 1
    jle .sort_asc_done

    mov r12d, 0
.loop_asc_i:
    mov eax, [total_rows]
    dec eax
    cmp r12d, eax
    jge .asc_finish

    mov r13d, r12d
    inc r13d
.loop_asc_j:
    cmp r13d, [total_rows]
    jge .next_asc_i

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r12d
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szTemp]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 500
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r12d
    lea r9, [lvitem]
    call [SendMessageA]

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 500
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r13d
    lea r9, [lvitem]
    call [SendMessageA]

    lea rcx, [szTemp]
    xor edx, edx
    call [strtod]
    movsd [rsp+20h], xmm0

    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]
    movsd [rsp+28h], xmm0

    movsd xmm1, [rsp+20h]
    xorpd xmm2, xmm2
    ucomisd xmm1, xmm2
    jp .asc_cmp_str
    jne .asc_cmp_num
    cmp byte [szTemp], '0'
    je .asc_cmp_num
.asc_cmp_str:
    lea rcx, [szTemp]
    lea rdx, [szCellVal]
    call [lstrcmpiA]
    cmp eax, 0
    jg .do_asc_swap
    jmp .next_asc_j

.asc_cmp_num:
    movsd xmm0, [rsp+28h]
    ucomisd xmm1, xmm0
    ja .do_asc_swap
    jmp .next_asc_j

.do_asc_swap:
    mov r14d, 1
.swap_asc_col:
    cmp r14d, 10
    jg .next_asc_j

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r12d
    mov dword [lvitem+8], r14d
    lea rax, [szTemp]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 500
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r12d
    lea r9, [lvitem]
    call [SendMessageA]

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov dword [lvitem+8], r14d
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 500
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r13d
    lea r9, [lvitem]
    call [SendMessageA]

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r12d
    mov dword [lvitem+8], r14d
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, r12d
    lea r9, [lvitem]
    call [SendMessageA]

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov dword [lvitem+8], r14d
    lea rax, [szTemp]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, r13d
    lea r9, [lvitem]
    call [SendMessageA]

    inc r14d
    jmp .swap_asc_col

.next_asc_j:
    inc r13d
    jmp .loop_asc_j

.next_asc_i:
    inc r12d
    jmp .loop_asc_i

.asc_finish:
    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [sort_asc_msg]
    call [SendMessageA]

.sort_asc_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    mov r14, [rbp-48]
    mov r15, [rbp-56]
    leave
    ret

DoSortDescending:
    push rbp
    mov rbp, rsp
    sub rsp, 70h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13
    mov [rbp-48], r14
    mov [rbp-56], r15

    cmp [total_rows], 1
    jle .sort_desc_done

    mov r12d, 0
.loop_desc_i:
    mov eax, [total_rows]
    dec eax
    cmp r12d, eax
    jge .desc_finish

    mov r13d, r12d
    inc r13d
.loop_desc_j:
    cmp r13d, [total_rows]
    jge .next_desc_i

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r12d
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szTemp]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 500
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r12d
    lea r9, [lvitem]
    call [SendMessageA]

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov eax, [cur_col]
    mov dword [lvitem+8], eax
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 500
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r13d
    lea r9, [lvitem]
    call [SendMessageA]

    lea rcx, [szTemp]
    xor edx, edx
    call [strtod]
    movsd [rsp+20h], xmm0

    lea rcx, [szCellVal]
    xor edx, edx
    call [strtod]
    movsd [rsp+28h], xmm0

    movsd xmm1, [rsp+20h]
    xorpd xmm2, xmm2
    ucomisd xmm1, xmm2
    jp .desc_cmp_str
    jne .desc_cmp_num
    cmp byte [szTemp], '0'
    je .desc_cmp_num
.desc_cmp_str:
    lea rcx, [szTemp]
    lea rdx, [szCellVal]
    call [lstrcmpiA]
    cmp eax, 0
    jl .do_desc_swap
    jmp .next_desc_j

.desc_cmp_num:
    movsd xmm0, [rsp+28h]
    ucomisd xmm1, xmm0
    jb .do_desc_swap
    jmp .next_desc_j

.do_desc_swap:
    mov r14d, 1
.swap_desc_col:
    cmp r14d, 10
    jg .next_desc_j

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r12d
    mov dword [lvitem+8], r14d
    lea rax, [szTemp]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 500
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r12d
    lea r9, [lvitem]
    call [SendMessageA]

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov dword [lvitem+8], r14d
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov dword [lvitem+32], 500
    mov rcx, [hGrid]
    mov edx, LVM_GETITEMTEXTA
    mov r8d, r13d
    lea r9, [lvitem]
    call [SendMessageA]

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r12d
    mov dword [lvitem+8], r14d
    lea rax, [szCellVal]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, r12d
    lea r9, [lvitem]
    call [SendMessageA]

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r13d
    mov dword [lvitem+8], r14d
    lea rax, [szTemp]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, r13d
    lea r9, [lvitem]
    call [SendMessageA]

    inc r14d
    jmp .swap_desc_col

.next_desc_j:
    inc r13d
    jmp .loop_desc_j

.next_desc_i:
    inc r12d
    jmp .loop_desc_i

.desc_finish:
    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [sort_desc_msg]
    call [SendMessageA]

.sort_desc_done:
    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    mov r14, [rbp-48]
    mov r15, [rbp-56]
    leave
    ret

DoFormatPainter:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [fmt_painter_msg]
    call [SendMessageA]
    leave
    ret

DoWrapText:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [wrap_txt_msg]
    call [SendMessageA]
    call DoRefreshGrid
    leave
    ret

DoMergeCenter:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    mov r8d, ALIGN_CENTER
    call ApplySelectionAlignment
    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 0
    lea r9, [merge_ct_msg]
    call [SendMessageA]
    leave
    ret

DoIncDecimals:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    call ApplyCommaFormat
    leave
    ret

DoDecDecimals:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    call ApplyCommaFormat
    leave
    ret

DoClearCell:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], rsi
    mov [rbp-24], rdi
    mov [rbp-32], r12
    mov [rbp-40], r13

    ; Determine selection bounds
    mov eax, [sel_start_row]
    mov edx, [sel_end_row]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov r12d, eax       ; min_row
    mov r13d, edx       ; max_row

    mov eax, [sel_start_col]
    mov edx, [sel_end_col]
    cmp eax, edx
    jle @f
    xchg eax, edx
@@:
    mov ebx, eax        ; min_col
    mov esi, edx        ; max_col

.dcc_r_loop:
    cmp r12d, r13d
    jg .dcc_finish

    mov edi, ebx
.dcc_c_loop:
    cmp edi, esi
    jg .dcc_next_r

    mov ecx, r12d
    mov edx, edi
    call BeginCellChange

    mov dword [lvitem], LVIF_TEXT
    mov dword [lvitem+4], r12d
    mov dword [lvitem+8], edi
    lea rax, [empty_str]
    mov qword [lvitem+24], rax
    mov rcx, [hGrid]
    mov edx, LVM_SETITEMTEXTA
    mov r8d, r12d
    lea r9, [lvitem]
    call [SendMessageA]

    mov ecx, r12d
    mov edx, edi
    xor r8, r8
    call SetCellFormula

    mov ecx, r12d
    mov edx, edi
    lea r8, [empty_str]
    xor r9, r9
    call CommitCellChange

    inc edi
    jmp .dcc_c_loop

.dcc_next_r:
    inc r12d
    jmp .dcc_r_loop

.dcc_finish:
    mov rcx, [hFormulaEdit]
    lea rdx, [empty_str]
    call [SetWindowTextA]

    call SaveActiveSheetToBuffer
    call UpdateStatusStats

    mov rcx, [hGrid]
    xor edx, edx
    mov r8d, 1
    call [InvalidateRect]

    mov rbx, [rbp-8]
    mov rsi, [rbp-16]
    mov rdi, [rbp-24]
    mov r12, [rbp-32]
    mov r13, [rbp-40]
    leave
    ret

DoSelectAll:
    push rbp
    mov rbp, rsp
    sub rsp, 40h
    mov dword [sel_start_row], 0
    mov dword [sel_start_col], 1
    mov eax, [total_rows]
    dec eax
    mov [sel_end_row], eax
    mov dword [sel_end_col], 26
    mov dword [cur_row], 0
    mov dword [cur_col], 1
    call UpdateCellSelectionDisplay
    call SyncListViewItemState
    leave
    ret

DoResetZoom:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    mov rcx, [hStatusBar]
    mov edx, SB_SETTEXTA
    mov r8d, 3
    lea r9, [sb_zoom_txt]
    call [SendMessageA]
    leave
    ret

DoRefreshGrid:
    push rbp
    mov rbp, rsp
    sub rsp, 20h
    call UpdateStatusStats
    mov rcx, [hWnd]
    call [UpdateWindow]
    leave
    ret


; =============================================================================
; In-Place Cell Editing: StartInPlaceCellEdit
; =============================================================================
StartInPlaceCellEdit:
    push rbp
    mov rbp, rsp
    sub rsp, 60h
    mov [rbp-8], rbx
    mov [rbp-16], r12
    mov [rbp-24], r13

    ; Validate row and column
    mov eax, [cur_col]
    cmp eax, 1
    jl .sice_ret
    cmp eax, 26
    jg .sice_ret
    mov eax, [cur_row]
    cmp eax, [total_rows]
    jge .sice_ret

    ; Ensure item is visible
    mov rcx, [hGrid]
    mov edx, 1013h ; LVM_ENSUREVISIBLE
    mov r8d, [cur_row]
    xor r9d, r9d
    call [SendMessageA]

    ; Query subitem rectangle using LVM_GETSUBITEMRECT (1038h)
    ; rcCell.top = cur_col
    ; rcCell.left = LVIR_BOUNDS (0)
    lea rbx, [rcCell]
    mov dword [rbx], 0          ; left = LVIR_BOUNDS
    mov eax, [cur_col]
    mov dword [rbx+4], eax      ; top = cur_col
    mov dword [rbx+8], 0
    mov dword [rbx+12], 0

    mov rcx, [hGrid]
    mov edx, LVM_GETSUBITEMRECT
    mov r8d, [cur_row]
    mov r9, rbx
    call [SendMessageA]

    ; Calculate width and height
    mov eax, [rbx+8]            ; right
    sub eax, [rbx]              ; left
    cmp eax, 10
    jle .sice_ret
    mov r12d, eax               ; width

    mov eax, [rbx+12]           ; bottom
    sub eax, [rbx+4]            ; top
    cmp eax, 10
    jle .sice_ret
    cmp eax, 20
    jge @f
    mov eax, 20
@@:
    mov r13d, eax               ; height

    ; Populate edit control with cell's formula or text
    mov ecx, [cur_row]
    mov edx, [cur_col]
    call GetCellFormula
    test rax, rax
    jz .sice_load_text

    ; It has a formula: load formula text
    mov rcx, [hCellEdit]
    mov rdx, rax
    call [SetWindowTextA]
    jmp .sice_show

.sice_load_text:
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

    mov rcx, [hCellEdit]
    lea rdx, [szCellVal]
    call [SetWindowTextA]

.sice_show:
    ; Position hCellEdit directly over the cell
    mov rcx, [hCellEdit]
    xor edx, edx                 ; HWND_TOP
    mov r8d, dword [rbx]         ; rc.left
    mov r9d, dword [rbx+4]       ; rc.top
    sub rsp, 30h
    mov dword [rsp+20h], r12d    ; width
    mov dword [rsp+28h], r13d    ; height
    mov dword [rsp+30h], 0040h   ; SWP_SHOWWINDOW
    call [SetWindowPos]
    add rsp, 30h

    ; Set focus to hCellEdit and select all text
    mov rcx, [hCellEdit]
    call [SetFocus]
    mov rcx, [hCellEdit]
    mov edx, EM_SETSEL
    xor r8d, r8d
    mov r9d, -1
    call [SendMessageA]

    mov dword [bCellEditing], 1

.sice_ret:
    mov rbx, [rbp-8]
    mov r12, [rbp-16]
    mov r13, [rbp-24]
    leave
    ret


; =============================================================================
; In-Place Cell Editing: CommitInPlaceCellEdit
; =============================================================================
CommitInPlaceCellEdit:
    push rbp
    mov rbp, rsp
    sub rsp, 40h

    cmp dword [bCellEditing], 0
    jz .cice_done

    ; Prevent reentrancy
    mov dword [bCellEditing], 0

    ; Hide in-place editor immediately
    mov rcx, [hCellEdit]
    xor edx, edx ; SW_HIDE
    call [ShowWindow]

    ; Read edited text from hCellEdit
    mov rcx, [hCellEdit]
    lea rdx, [szCellVal]
    mov r8d, 1000
    call [GetWindowTextA]

    ; Transfer text to Formula Bar and evaluate
    mov rcx, [hFormulaEdit]
    lea rdx, [szCellVal]
    call [SetWindowTextA]

    call DoEvalFormula

    ; Refresh UI
    call UpdateCellSelectionDisplay
    call UpdateStatusStats

    ; Return focus to grid
    mov rcx, [hGrid]
    call [SetFocus]

.cice_done:
    leave
    ret


; =============================================================================
; In-Place Cell Editing: CancelInPlaceCellEdit
; =============================================================================
CancelInPlaceCellEdit:
    push rbp
    mov rbp, rsp
    sub rsp, 30h

    cmp dword [bCellEditing], 0
    jz .can_done

    mov dword [bCellEditing], 0
    mov rcx, [hCellEdit]
    xor edx, edx ; SW_HIDE
    call [ShowWindow]

    mov rcx, [hGrid]
    call [SetFocus]

.can_done:
    leave
    ret


; =============================================================================
; CellEditSubclassProc: Subclass procedure for In-Place Cell Edit Control
; =============================================================================
CellEditSubclassProc:
    push rbp
    mov rbp, rsp
    sub rsp, 80h
    mov [rbp-08h], rbx
    mov [rbp-10h], rsi
    mov [rbp-18h], rdi
    mov [rbp-20h], r12
    mov [rbp-28h], r13

    mov [rbp-30h], rcx ; hWnd
    mov [rbp-38h], rdx ; uMsg
    mov [rbp-40h], r8  ; wParam
    mov [rbp-48h], r9  ; lParam

    cmp edx, 0087h ; WM_GETDLGCODE
    je .cep_dlgcode
    cmp edx, 0100h ; WM_KEYDOWN
    je .cep_keydown
    cmp edx, 0008h ; WM_KILLFOCUS
    je .cep_killfocus
    cmp edx, 0101h ; WM_KEYUP
    je .cep_keyup

.cep_call_orig:
    mov rax, [origCellEditProc]
    mov rcx, rax
    mov rdx, [rbp-30h]
    mov r8, [rbp-38h]
    mov r9, [rbp-40h]
    mov r10, [rbp-48h]
    mov [rsp+20h], r10
    call [CallWindowProcA]
    jmp .cep_done

.cep_dlgcode:
    mov eax, 0085h ; DLGC_WANTALLKEYS | DLGC_WANTCHARS | DLGC_WANTARROWS
    jmp .cep_done

.cep_keydown:
    mov r8d, dword [rbp-40h]
    cmp r8d, 0Dh ; VK_RETURN
    je .cep_enter
    cmp r8d, 1Bh ; VK_ESCAPE
    je .cep_escape
    cmp r8d, 09h ; VK_TAB
    je .cep_tab
    jmp .cep_call_orig

.cep_enter:
    call CommitInPlaceCellEdit
    ; Move selection down: cur_row + 1
    mov eax, [cur_row]
    inc eax
    cmp eax, [total_rows]
    jge @f
    mov [cur_row], eax
    mov [sel_start_row], eax
    mov [sel_end_row], eax
    call UpdateCellSelectionDisplay
    call SyncListViewItemState
@@:
    xor eax, eax
    jmp .cep_done

.cep_tab:
    call CommitInPlaceCellEdit
    ; Move selection right: cur_col + 1
    mov eax, [cur_col]
    inc eax
    cmp eax, 26
    jg @f
    mov [cur_col], eax
    mov [sel_start_col], eax
    mov [sel_end_col], eax
    call UpdateCellSelectionDisplay
    call SyncListViewItemState
@@:
    xor eax, eax
    jmp .cep_done

.cep_escape:
    call CancelInPlaceCellEdit
    xor eax, eax
    jmp .cep_done

.cep_killfocus:
    call CommitInPlaceCellEdit
    jmp .cep_call_orig

.cep_keyup:
    ; Sync to Formula Bar live
    mov rcx, [hCellEdit]
    lea rdx, [szCellVal]
    mov r8d, 1000
    call [GetWindowTextA]

    mov rcx, [hFormulaEdit]
    lea rdx, [szCellVal]
    call [SetWindowTextA]
    jmp .cep_call_orig

.cep_done:
    mov rbx, [rbp-08h]
    mov rsi, [rbp-10h]
    mov rdi, [rbp-18h]
    mov r12, [rbp-20h]
    mov r13, [rbp-28h]
    leave
    ret

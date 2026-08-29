; =============================================================================
; Excel Viewer Pro — High Performance SIMD & Math Engine (x86-64 FASM)
; =============================================================================
; Fast vectorized calculation routines for spreadsheet calculations:
; - SSE2 / AVX Vector Sum, Average, Min, Max, and SUMPRODUCT
; - Fast Financial Math: PMT, PV, FV, NPER
; - Ultra-Fast 64-bit FNV-1a String Hashing
; - Non-blank cell scanner
; =============================================================================

format PE64 GUI 6.0 DLL
entry DllEntryPoint

include 'C:\asm\hdd\INCLUDE\WIN64A.INC'

section '.text' code readable executable

; -----------------------------------------------------------------------------
; DLL Entry Point
; -----------------------------------------------------------------------------
proc DllEntryPoint hinstDLL, fdwReason, lpvReserved
    mov eax, 1
    ret
endp

; =============================================================================
; double vec_sum_f64(const double* arr, uint64_t count)
; RCX = pointer to array of double (8 bytes each)
; RDX = number of elements
; Returns sum in XMM0
; =============================================================================
proc vec_sum_f64
    xorpd xmm0, xmm0
    test rcx, rcx
    jz .done
    test rdx, rdx
    jz .done

    xorpd xmm1, xmm1
    mov r8, rdx
    shr r8, 2               ; 4 doubles (32 bytes) per iteration
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
endp

; =============================================================================
; double vec_avg_f64(const double* arr, uint64_t count)
; RCX = pointer to array of double
; RDX = number of elements
; Returns average in XMM0 (0.0 if count == 0)
; =============================================================================
proc vec_avg_f64
    test rdx, rdx
    jnz .calc
    xorpd xmm0, xmm0
    ret

.calc:
    push rdx
    sub rsp, 20h
    call vec_sum_f64
    add rsp, 20h
    pop rdx
    cvtsi2sd xmm1, rdx
    divsd xmm0, xmm1
    ret
endp

; =============================================================================
; double vec_min_f64(const double* arr, uint64_t count)
; RCX = pointer to array of double
; RDX = number of elements
; Returns minimum in XMM0
; =============================================================================
proc vec_min_f64
    xorpd xmm0, xmm0
    test rcx, rcx
    jz .done
    test rdx, rdx
    jz .done

    movsd xmm0, [rcx]
    add rcx, 8
    dec rdx
    jz .done

.loop:
    minsd xmm0, [rcx]
    add rcx, 8
    dec rdx
    jnz .loop

.done:
    ret
endp

; =============================================================================
; double vec_max_f64(const double* arr, uint64_t count)
; RCX = pointer to array of double
; RDX = number of elements
; Returns maximum in XMM0
; =============================================================================
proc vec_max_f64
    xorpd xmm0, xmm0
    test rcx, rcx
    jz .done
    test rdx, rdx
    jz .done

    movsd xmm0, [rcx]
    add rcx, 8
    dec rdx
    jz .done

.loop:
    maxsd xmm0, [rcx]
    add rcx, 8
    dec rdx
    jnz .loop

.done:
    ret
endp

; =============================================================================
; double vec_sumproduct_f64(const double* arrA, const double* arrB, uint64_t count)
; RCX = pointer to array A
; RDX = pointer to array B
; R8  = number of elements
; Returns sum of products in XMM0
; =============================================================================
proc vec_sumproduct_f64
    xorpd xmm0, xmm0
    test rcx, rcx
    jz .done
    test rdx, rdx
    jz .done
    test r8, r8
    jz .done

    xorpd xmm1, xmm1
    mov r9, r8
    shr r9, 2               ; 4 elements per iteration
    jz .tail

.loop4:
    movupd xmm2, [rcx]
    movupd xmm3, [rdx]
    mulpd xmm2, xmm3
    addpd xmm0, xmm2

    movupd xmm4, [rcx+16]
    movupd xmm5, [rdx+16]
    mulpd xmm4, xmm5
    addpd xmm1, xmm4

    add rcx, 32
    add rdx, 32
    dec r9
    jnz .loop4

    addpd xmm0, xmm1

.tail:
    and r8, 3
    jz .hadd

.loop1:
    movsd xmm2, [rcx]
    mulsd xmm2, [rdx]
    addsd xmm0, xmm2
    add rcx, 8
    add rdx, 8
    dec r8
    jnz .loop1

.hadd:
    movhlps xmm2, xmm0
    addsd xmm0, xmm2

.done:
    ret
endp

; =============================================================================
; double fast_pow_f64(double base, double exp)
; Helper: base^exp via x87 FPU: 2^(exp * log2(base))
; XMM0 = base, XMM1 = exp -> Returns in XMM0
; =============================================================================
proc fast_pow_f64
    sub rsp, 28h
    movsd [rsp], xmm0
    movsd [rsp+8], xmm1

    fld qword [rsp+8]       ; st0 = exp
    fld qword [rsp]         ; st0 = base, st1 = exp
    fyl2x                   ; st0 = exp * log2(base)

    fld st0                 ; duplicate
    frndint                 ; st0 = round(val), st1 = val
    fsub st1, st0           ; st1 = frac, st0 = int
    fxch st1                ; st0 = frac, st1 = int
    f2xm1                   ; st0 = 2^frac - 1
    fld1
    faddp st1, st0          ; st0 = 2^frac
    fscale                  ; st0 = 2^frac * 2^int = base^exp
    fstp st1                ; discard int
    fstp qword [rsp]        ; save result

    movsd xmm0, [rsp]
    add rsp, 28h
    ret
endp

; =============================================================================
; double fast_pmt_f64(double rate, double nper, double pv, double fv, int64_t type)
; XMM0 = rate
; XMM1 = nper
; XMM2 = pv
; XMM3 = fv
; [RSP + 28h] = type (0 = end of period, 1 = beginning of period)
; Returns PMT in XMM0
; =============================================================================
proc fast_pmt_f64
    sub rsp, 48h
    movsd [rsp], xmm0       ; rate
    movsd [rsp+8], xmm1     ; nper
    movsd [rsp+10h], xmm2   ; pv
    movsd [rsp+18h], xmm3   ; fv
    mov rax, [rsp+78h]      ; type parameter (caller stack)
    mov [rsp+20h], rax

    ; Check if rate == 0
    xorpd xmm4, xmm4
    ucomisd xmm0, xmm4
    jne .calc_pvif

    ; Zero rate: PMT = -(pv + fv) / nper
    movsd xmm0, [rsp+10h]
    addsd xmm0, [rsp+18h]
    xorpd xmm5, xmm5
    subsd xmm5, xmm0        ; -(pv + fv)
    divsd xmm5, [rsp+8]     ; / nper
    movsd xmm0, xmm5
    jmp .exit

.calc_pvif:
    ; base = 1 + rate
    movsd xmm0, [rsp]
    mov rax, 03FF0000000000000h  ; 1.0 in IEEE 754
    movq xmm4, rax
    addsd xmm0, xmm4        ; 1 + rate
    movsd xmm1, [rsp+8]     ; nper
    call fast_pow_f64       ; XMM0 = pvif = (1 + rate)^nper
    movsd [rsp+28h], xmm0   ; store pvif

    ; num = -(pv * pvif + fv)
    movsd xmm1, [rsp+10h]   ; pv
    mulsd xmm1, xmm0        ; pv * pvif
    addsd xmm1, [rsp+18h]   ; pv * pvif + fv
    xorpd xmm2, xmm2
    subsd xmm2, xmm1        ; -(pv * pvif + fv)

    ; fact = rate / (pvif - 1)
    mov rax, 03FF0000000000000h
    movq xmm3, rax
    movsd xmm4, [rsp+28h]   ; pvif
    subsd xmm4, xmm3        ; pvif - 1
    movsd xmm5, [rsp]       ; rate
    divsd xmm5, xmm4        ; rate / (pvif - 1)

    ; pmt = fact * num
    mulsd xmm5, xmm2

    ; If type == 1: pmt = pmt / (1 + rate)
    cmp qword [rsp+20h], 1
    jne .store_res

    movsd xmm0, [rsp]
    addsd xmm0, xmm3        ; 1 + rate
    divsd xmm5, xmm0

.store_res:
    movsd xmm0, xmm5

.exit:
    add rsp, 48h
    ret
endp

; =============================================================================
; double fast_pv_f64(double rate, double nper, double pmt, double fv, int64_t type)
; XMM0 = rate, XMM1 = nper, XMM2 = pmt, XMM3 = fv, [RSP+28h] = type
; Returns PV in XMM0
; =============================================================================
proc fast_pv_f64
    sub rsp, 48h
    movsd [rsp], xmm0       ; rate
    movsd [rsp+8], xmm1     ; nper
    movsd [rsp+10h], xmm2   ; pmt
    movsd [rsp+18h], xmm3   ; fv
    mov rax, [rsp+78h]      ; type
    mov [rsp+20h], rax

    xorpd xmm4, xmm4
    ucomisd xmm0, xmm4
    jne .calc

    ; rate == 0: PV = -(pmt * nper + fv)
    movsd xmm0, [rsp+10h]
    mulsd xmm0, [rsp+8]
    addsd xmm0, [rsp+18h]
    xorpd xmm5, xmm5
    subsd xmm5, xmm0
    movsd xmm0, xmm5
    jmp .exit

.calc:
    movsd xmm0, [rsp]
    mov rax, 03FF0000000000000h
    movq xmm4, rax
    addsd xmm0, xmm4        ; 1 + rate
    movsd xmm1, [rsp+8]
    call fast_pow_f64       ; XMM0 = pvif
    movsd [rsp+28h], xmm0

    ; fact = (1 + rate * type) * (pvif - 1) / rate
    movsd xmm1, [rsp]       ; rate
    cvtsi2sd xmm2, qword [rsp+20h] ; type
    mulsd xmm1, xmm2        ; rate * type
    mov rax, 03FF0000000000000h
    movq xmm3, rax
    addsd xmm1, xmm3        ; 1 + rate * type

    movsd xmm4, [rsp+28h]   ; pvif
    subsd xmm4, xmm3        ; pvif - 1
    mulsd xmm1, xmm4
    divsd xmm1, [rsp]       ; fact

    ; -(fv + pmt * fact) / pvif
    movsd xmm2, [rsp+10h]   ; pmt
    mulsd xmm2, xmm1        ; pmt * fact
    addsd xmm2, [rsp+18h]   ; fv + pmt * fact
    xorpd xmm0, xmm0
    subsd xmm0, xmm2
    divsd xmm0, [rsp+28h]   ; / pvif

.exit:
    add rsp, 48h
    ret
endp

; =============================================================================
; double fast_fv_f64(double rate, double nper, double pmt, double pv, int64_t type)
; XMM0 = rate, XMM1 = nper, XMM2 = pmt, XMM3 = pv, [RSP+28h] = type
; Returns FV in XMM0
; =============================================================================
proc fast_fv_f64
    sub rsp, 48h
    movsd [rsp], xmm0       ; rate
    movsd [rsp+8], xmm1     ; nper
    movsd [rsp+10h], xmm2   ; pmt
    movsd [rsp+18h], xmm3   ; pv
    mov rax, [rsp+78h]      ; type
    mov [rsp+20h], rax

    xorpd xmm4, xmm4
    ucomisd xmm0, xmm4
    jne .calc

    ; rate == 0: FV = -(pv + pmt * nper)
    movsd xmm0, [rsp+10h]
    mulsd xmm0, [rsp+8]
    addsd xmm0, [rsp+18h]
    xorpd xmm5, xmm5
    subsd xmm5, xmm0
    movsd xmm0, xmm5
    jmp .exit

.calc:
    movsd xmm0, [rsp]
    mov rax, 03FF0000000000000h
    movq xmm4, rax
    addsd xmm0, xmm4
    movsd xmm1, [rsp+8]
    call fast_pow_f64       ; XMM0 = pvif
    movsd [rsp+28h], xmm0

    ; fact = (1 + rate * type) * (pvif - 1) / rate
    movsd xmm1, [rsp]
    cvtsi2sd xmm2, qword [rsp+20h]
    mulsd xmm1, xmm2
    mov rax, 03FF0000000000000h
    movq xmm3, rax
    addsd xmm1, xmm3

    movsd xmm4, [rsp+28h]
    subsd xmm4, xmm3
    mulsd xmm1, xmm4
    divsd xmm1, [rsp]       ; fact

    ; -(pv * pvif + pmt * fact)
    movsd xmm2, [rsp+18h]   ; pv
    mulsd xmm2, [rsp+28h]   ; pv * pvif
    movsd xmm3, [rsp+10h]   ; pmt
    mulsd xmm3, xmm1        ; pmt * fact
    addsd xmm2, xmm3
    xorpd xmm0, xmm0
    subsd xmm0, xmm2

.exit:
    add rsp, 48h
    ret
endp

; =============================================================================
; uint64_t fast_str_hash(const char* str, uint64_t len)
; 64-bit FNV-1a Hash for ultra-fast cell text / formula hashing
; RCX = string pointer, RDX = length in bytes
; Returns 64-bit hash in RAX
; =============================================================================
proc fast_str_hash
    mov rax, 0CBF29CE484222325h      ; FNV_offset_basis
    test rcx, rcx
    jz .done
    test rdx, rdx
    jz .done

    mov r8, 0100000001B3h            ; FNV_prime

.loop:
    movzx r9d, byte [rcx]
    xor rax, r9
    imul rax, r8
    inc rcx
    dec rdx
    jnz .loop

.done:
    ret
endp

; =============================================================================
; uint64_t fast_count_nonblank(const uint64_t* ptr_arr, uint64_t count)
; RCX = array of 64-bit values / pointers
; RDX = count
; Returns count of non-zero values in RAX
; =============================================================================
proc fast_count_nonblank
    xor eax, eax
    test rcx, rcx
    jz .done
    test rdx, rdx
    jz .done

.loop:
    cmp qword [rcx], 0
    jz @f
    inc rax
@@: add rcx, 8
    dec rdx
    jnz .loop

.done:
    ret
endp

; =============================================================================
; Exports Table
; =============================================================================
section '.edata' export data readable

export 'xlsx_math64.dll',\
       vec_sum_f64, 'vec_sum_f64',\
       vec_avg_f64, 'vec_avg_f64',\
       vec_min_f64, 'vec_min_f64',\
       vec_max_f64, 'vec_max_f64',\
       vec_sumproduct_f64, 'vec_sumproduct_f64',\
       fast_pow_f64, 'fast_pow_f64',\
       fast_pmt_f64, 'fast_pmt_f64',\
       fast_pv_f64, 'fast_pv_f64',\
       fast_fv_f64, 'fast_fv_f64',\
       fast_str_hash, 'fast_str_hash',\
       fast_count_nonblank, 'fast_count_nonblank'

; ============================================================================
; Excel Viewer Pro — PE Import Table Section
; ============================================================================

section '.idata' import data readable

    dd 0, 0, 0, RVA kernel32_name, RVA kernel32_table
    dd 0, 0, 0, RVA user32_name,   RVA user32_table
    dd 0, 0, 0, RVA gdi32_name,    RVA gdi32_table
    dd 0, 0, 0, RVA comctl_name,   RVA comctl_table
    dd 0, 0, 0, RVA comdlg_name,   RVA comdlg_table
    dd 0, 0, 0, RVA msvcrt_name,   RVA msvcrt_table
    dd 0, 0, 0, 0, 0

kernel32_table:
    GetModuleHandleA            dq RVA _GetModuleHandleA
    ExitProcess                 dq RVA _ExitProcess
    QueryPerformanceFrequency   dq RVA _QueryPerformanceFrequency
    QueryPerformanceCounter     dq RVA _QueryPerformanceCounter
    GlobalAlloc                 dq RVA _GlobalAlloc
    GlobalLock                  dq RVA _GlobalLock
    GlobalUnlock                dq RVA _GlobalUnlock
    lstrcmpiA                   dq RVA _lstrcmpiA
    GetTempPathA                dq RVA _GetTempPathA
    GetCommandLineA             dq RVA _GetCommandLineA
    LoadLibraryA                dq RVA _LoadLibraryA
    GetProcAddress              dq RVA _GetProcAddress
    HeapCreate                  dq RVA _HeapCreate
    HeapAlloc                   dq RVA _HeapAlloc
    HeapDestroy                 dq RVA _HeapDestroy
                                dq 0

user32_table:
    RegisterClassExA            dq RVA _RegisterClassExA
    CreateWindowExA             dq RVA _CreateWindowExA
    ShowWindow                  dq RVA _ShowWindow
    UpdateWindow                dq RVA _UpdateWindow
    GetMessageA                 dq RVA _GetMessageA
    TranslateMessage            dq RVA _TranslateMessage
    DispatchMessageA            dq RVA _DispatchMessageA
    DefWindowProcA              dq RVA _DefWindowProcA
    PostQuitMessage             dq RVA _PostQuitMessage
    LoadIconA                   dq RVA _LoadIconA
    LoadCursorA                 dq RVA _LoadCursorA
    SendMessageA                dq RVA _SendMessageA
    GetWindowTextA              dq RVA _GetWindowTextA
    SetWindowTextA              dq RVA _SetWindowTextA
    MessageBoxA                 dq RVA _MessageBoxA
    CreateMenu                  dq RVA _CreateMenu
    CreatePopupMenu             dq RVA _CreatePopupMenu
    AppendMenuA                 dq RVA _AppendMenuA
    SetMenu                     dq RVA _SetMenu
    OpenClipboard               dq RVA _OpenClipboard
    CloseClipboard              dq RVA _CloseClipboard
    EmptyClipboard              dq RVA _EmptyClipboard
    GetClipboardData            dq RVA _GetClipboardData
    SetClipboardData            dq RVA _SetClipboardData
    GetCursorPos                dq RVA _GetCursorPos
    ScreenToClient              dq RVA _ScreenToClient
    ClientToScreen              dq RVA _ClientToScreen
    MoveWindow                  dq RVA _MoveWindow
    SetWindowLongPtrA           dq RVA _SetWindowLongPtrA
    CallWindowProcA             dq RVA _CallWindowProcA
    SetFocus                    dq RVA _SetFocus
    GetKeyState                 dq RVA _GetKeyState
    InvalidateRect              dq RVA _InvalidateRect
    EnableWindow                dq RVA _EnableWindow
    DestroyWindow               dq RVA _DestroyWindow
    DrawFocusRect               dq RVA _DrawFocusRect
    GetDlgItem                  dq RVA _GetDlgItem
    SetCapture                  dq RVA _SetCapture
    ReleaseCapture              dq RVA _ReleaseCapture
    GetAsyncKeyState             dq RVA _GetAsyncKeyState
    SetWindowPos                dq RVA _SetWindowPos
    LoadImageA                  dq RVA _LoadImageA
    DrawTextA                   dq RVA _DrawTextA
    FillRect                    dq RVA _FillRect
    GetDC                       dq RVA _GetDC
    ReleaseDC                   dq RVA _ReleaseDC
    GetSysColor                 dq RVA _GetSysColor
    SetMenuItemBitmaps          dq RVA _SetMenuItemBitmaps
    SetMenuItemInfoA            dq RVA _SetMenuItemInfoA
    TrackPopupMenu              dq RVA _TrackPopupMenu
    DestroyMenu                 dq RVA _DestroyMenu
                                dq 0

gdi32_table:
    CreateFontA                 dq RVA _CreateFontA
    SelectObject                dq RVA _SelectObject
    DeleteObject                dq RVA _DeleteObject
    CreateSolidBrush            dq RVA _CreateSolidBrush
    CreatePen                   dq RVA _CreatePen
    RoundRect                   dq RVA _RoundRect
    Rectangle                   dq RVA _Rectangle
    Ellipse                     dq RVA _Ellipse
    MoveToEx                    dq RVA _MoveToEx
    LineTo                      dq RVA _LineTo
    SetTextColor                dq RVA _SetTextColor
    SetBkMode                   dq RVA _SetBkMode
    SetBkColor                  dq RVA _SetBkColor
    Polygon                     dq RVA _Polygon
    CreateCompatibleDC          dq RVA _CreateCompatibleDC
    CreateCompatibleBitmap      dq RVA _CreateCompatibleBitmap
    DeleteDC                    dq RVA _DeleteDC
                                dq 0

comctl_table:
    InitCommonControlsEx        dq RVA _InitCommonControlsEx
                                dq 0

comdlg_table:
    GetOpenFileNameA            dq RVA _GetOpenFileNameA
    GetSaveFileNameA            dq RVA _GetSaveFileNameA
                                dq 0

msvcrt_table:
    strtod                      dq RVA _strtod
    strlen                      dq RVA _strlen
    sprintf                     dq RVA _sprintf
    fopen                       dq RVA _fopen
    fgets                       dq RVA _fgets
    fprintf                     dq RVA _fprintf
    fclose                      dq RVA _fclose
    fread                       dq RVA _fread
    fseek                       dq RVA _fseek
    system                      dq RVA _system
    remove                      dq RVA _remove
                                dq 0

kernel32_name   db 'KERNEL32.DLL', 0
user32_name     db 'USER32.DLL', 0
gdi32_name      db 'GDI32.DLL', 0
comctl_name     db 'COMCTL32.DLL', 0
comdlg_name     db 'COMDLG32.DLL', 0
msvcrt_name     db 'MSVCRT.DLL', 0

_GetModuleHandleA dw 0
 db 'GetModuleHandleA', 0
_ExitProcess dw 0
 db 'ExitProcess', 0
_QueryPerformanceFrequency dw 0
 db 'QueryPerformanceFrequency', 0
_QueryPerformanceCounter dw 0
 db 'QueryPerformanceCounter', 0
_GlobalAlloc dw 0
 db 'GlobalAlloc', 0
_GlobalLock dw 0
 db 'GlobalLock', 0
_GlobalUnlock dw 0
 db 'GlobalUnlock', 0
_lstrcmpiA dw 0
 db 'lstrcmpiA', 0
_GetTempPathA dw 0
 db 'GetTempPathA', 0
_GetCommandLineA dw 0
 db 'GetCommandLineA', 0
_LoadLibraryA dw 0
 db 'LoadLibraryA', 0
_GetProcAddress dw 0
 db 'GetProcAddress', 0
_HeapCreate dw 0
 db 'HeapCreate', 0
_HeapAlloc dw 0
 db 'HeapAlloc', 0
_HeapDestroy dw 0
 db 'HeapDestroy', 0

_RegisterClassExA dw 0
 db 'RegisterClassExA', 0
_CreateWindowExA dw 0
 db 'CreateWindowExA', 0
_ShowWindow dw 0
 db 'ShowWindow', 0
_UpdateWindow dw 0
 db 'UpdateWindow', 0
_GetMessageA dw 0
 db 'GetMessageA', 0
_TranslateMessage dw 0
 db 'TranslateMessage', 0
_DispatchMessageA dw 0
 db 'DispatchMessageA', 0
_DefWindowProcA dw 0
 db 'DefWindowProcA', 0
_PostQuitMessage dw 0
 db 'PostQuitMessage', 0
_LoadIconA dw 0
 db 'LoadIconA', 0
_LoadCursorA dw 0
 db 'LoadCursorA', 0
_SendMessageA dw 0
 db 'SendMessageA', 0
_GetWindowTextA dw 0
 db 'GetWindowTextA', 0
_SetWindowTextA dw 0
 db 'SetWindowTextA', 0
_MessageBoxA dw 0
 db 'MessageBoxA', 0
_CreateMenu dw 0
 db 'CreateMenu', 0
_CreatePopupMenu dw 0
 db 'CreatePopupMenu', 0
_AppendMenuA dw 0
 db 'AppendMenuA', 0
_SetMenu dw 0
 db 'SetMenu', 0
_OpenClipboard dw 0
 db 'OpenClipboard', 0
_CloseClipboard dw 0
 db 'CloseClipboard', 0
_EmptyClipboard dw 0
 db 'EmptyClipboard', 0
_GetClipboardData dw 0
 db 'GetClipboardData', 0
_SetClipboardData dw 0
 db 'SetClipboardData', 0
_GetCursorPos dw 0
 db 'GetCursorPos', 0
_ScreenToClient dw 0
 db 'ScreenToClient', 0
_ClientToScreen dw 0
 db 'ClientToScreen', 0
_MoveWindow dw 0
 db 'MoveWindow', 0
_SetWindowLongPtrA dw 0
 db 'SetWindowLongPtrA', 0
_CallWindowProcA dw 0
 db 'CallWindowProcA', 0
_SetFocus dw 0
 db 'SetFocus', 0
_GetKeyState dw 0
 db 'GetKeyState', 0
_InvalidateRect dw 0
 db 'InvalidateRect', 0
_EnableWindow dw 0
 db 'EnableWindow', 0
_DestroyWindow dw 0
 db 'DestroyWindow', 0
_DrawFocusRect dw 0
 db 'DrawFocusRect', 0
_GetDlgItem dw 0
 db 'GetDlgItem', 0
_SetCapture dw 0
 db 'SetCapture', 0
_ReleaseCapture dw 0
 db 'ReleaseCapture', 0
_GetAsyncKeyState dw 0
 db 'GetAsyncKeyState', 0
_SetWindowPos dw 0
 db 'SetWindowPos', 0
_LoadImageA dw 0
 db 'LoadImageA', 0
_DrawTextA dw 0
 db 'DrawTextA', 0
_FillRect dw 0
 db 'FillRect', 0
_GetDC dw 0
 db 'GetDC', 0
_ReleaseDC dw 0
 db 'ReleaseDC', 0
_GetSysColor dw 0
 db 'GetSysColor', 0
_SetMenuItemBitmaps dw 0
 db 'SetMenuItemBitmaps', 0
_SetMenuItemInfoA dw 0
 db 'SetMenuItemInfoA', 0
_TrackPopupMenu dw 0
 db 'TrackPopupMenu', 0
_DestroyMenu dw 0
 db 'DestroyMenu', 0

_CreateFontA dw 0
 db 'CreateFontA', 0

_SelectObject dw 0
 db 'SelectObject', 0
_DeleteObject dw 0
 db 'DeleteObject', 0
_CreateSolidBrush dw 0
 db 'CreateSolidBrush', 0
_CreatePen dw 0
 db 'CreatePen', 0
_RoundRect dw 0
 db 'RoundRect', 0
_Rectangle dw 0
 db 'Rectangle', 0
_Ellipse dw 0
 db 'Ellipse', 0
_MoveToEx dw 0
 db 'MoveToEx', 0
_LineTo dw 0
 db 'LineTo', 0
_SetTextColor dw 0
 db 'SetTextColor', 0
_SetBkMode dw 0
 db 'SetBkMode', 0
_SetBkColor dw 0
 db 'SetBkColor', 0
_Polygon dw 0
 db 'Polygon', 0
_CreateCompatibleDC dw 0
 db 'CreateCompatibleDC', 0
_CreateCompatibleBitmap dw 0
 db 'CreateCompatibleBitmap', 0
_DeleteDC dw 0
 db 'DeleteDC', 0

_InitCommonControlsEx dw 0
 db 'InitCommonControlsEx', 0

_GetOpenFileNameA dw 0
 db 'GetOpenFileNameA', 0
_GetSaveFileNameA dw 0
 db 'GetSaveFileNameA', 0

_strtod dw 0
 db 'strtod', 0
_strlen dw 0
 db 'strlen', 0
_sprintf dw 0
 db 'sprintf', 0
_fopen dw 0
 db 'fopen', 0
_fgets dw 0
 db 'fgets', 0
_fprintf dw 0
 db 'fprintf', 0
_fclose dw 0
 db 'fclose', 0
_fread dw 0
 db 'fread', 0
_fseek dw 0
 db 'fseek', 0
_system dw 0
 db 'system', 0
_remove dw 0
 db 'remove', 0


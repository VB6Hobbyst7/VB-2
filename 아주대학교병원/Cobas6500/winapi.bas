Attribute VB_Name = "WinAPI"

Option Explicit



Public Type POINTAPI

        X As Long

        Y As Long

End Type



Public Type MSG

    hwnd As Long

    message As Long

    wParam As Long

    lParam As Long

    time As Long

    pt As POINTAPI

End Type



Public Type SYSTEMTIME

        wYear As Integer

        wMonth As Integer

        wDayOfWeek As Integer

        wDay As Integer

        wHour As Integer

        wMinute As Integer

        wSecond As Integer

        wMilliseconds As Integer

End Type



Public Const SW_SHOWNORMAL = 1

Public Const WS_MAXIMIZE = &H1000000

Public Const WM_ACTIVATE = &H6



Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Public Declare Function lstrlen Lib "kernel32" (ByVal LPSTR&) As Integer

Public Declare Function lstrcpy Lib "kernel32" (LP1 As Any, LP2 As Any) As Long

Public Declare Function lstrcpyn Lib "kernel32" (LP1 As Any, LP2 As Any, ByVal slen As Long) As Long

Public Declare Function lstrcat Lib "Kernel" (LP1 As Any, LP2 As Any) As Long

Public Declare Function hmemcpy Lib "kernel32" (LP1 As Any, LP2 As Any, ByVal slen&) As Long

'Put binary Data into CARRAY Buffer

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nbytes As Long)

'Get binary Data From CARRAY Buffer

Public Declare Sub GetMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nbytes As Long)





Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long



Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long

Public Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long

Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long

Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Public Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long

Public Declare Function SetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long

Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long




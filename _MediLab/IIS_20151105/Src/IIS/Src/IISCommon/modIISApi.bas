Attribute VB_Name = "modIISApi"
'-----------------------------------------------------------------------------'
'   파일명  : modIISApi.bas
'   작성자  : 오세원
'   내  용  : API  클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long _
) As Long

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" ( _
    ByVal lpBuffer As String, _
    nSize As Long _
) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String _
) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String _
) As Long

Public Declare Function SetParent Lib "user32" ( _
    ByVal hWndChild As Long, _
    ByVal hWndNewParent As Long _
) As Long

Public Const WM_USER = &H400
Public Const SB_GETRECT = (WM_USER + 10)

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByRef lParam As Any _
) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const MF_STRING = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_DISABLED = &H2&
Public Const MF_SEPARATOR = &H800&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_RETURNCMD = &H100&

Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long

Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" ( _
    ByVal hMenu As Long, _
    ByVal wFlags As Long, _
    ByVal wIDNewItem As Long, _
    ByVal lpNewItem As Any _
) As Long

Public Declare Function TrackPopupMenu Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal wFlags As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nReserved As Long, _
    ByVal hwnd As Long, _
    ByVal lprc As Any _
) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As Long


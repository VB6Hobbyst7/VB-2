Attribute VB_Name = "modIISApi"
Option Explicit

Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_HWNDFIRST = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOWMAXIMIZED = 3

Public Declare Function ShowWindow Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nCmdShow As Long _
) As Long

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetWindow Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal wCmd As Long _
) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

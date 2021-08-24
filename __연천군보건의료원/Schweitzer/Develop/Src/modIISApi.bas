Attribute VB_Name = "modIISApi"
Option Explicit

Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_HWNDFIRST = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOWMAXIMIZED = 3

Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

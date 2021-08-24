Attribute VB_Name = "modIISPcxView"
'-----------------------------------------------------------------------------'
'   파일명  : modIISPcxView.cls
'   작성자  : 오세원
'   내  용  : PCX View Ini File Read (MDB Info)
'   작성일  : 2007-06-14
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const SW_HIDE = 0
Private Const SW_NORMAL = 1

Private Const WS_EX_APPWINDOW = &H40000
Private Const WS_SYSMENU As Long = &H80000

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'===============================================================================
' 함 수 명 : ShowInTaskBar()
' 설    명 : 작업 표시줄에 윈도우 표시
' 입 력 값 :
' 결 과 값 :
' 작 성 일 : 2007.06.15
' 작 성 자 : 오세원
'===============================================================================
'   Not Use
Public Sub ShowInTaskBar(hwnd As Long, Flag As Boolean)
    Dim WindowLong As Long

On Error GoTo ShowInTaskBar_Error

    ShowWindow hwnd, SW_HIDE
    WindowLong = GetWindowLong(hwnd, GWL_EXSTYLE)

    If Flag = True Then
        SetWindowLong hwnd, GWL_EXSTYLE, WindowLong Xor WS_EX_APPWINDOW
    Else
        SetWindowLong hwnd, GWL_EXSTYLE, WindowLong Or WS_EX_APPWINDOW
    End If

    ShowWindow hwnd, SW_NORMAL
 
    '  System Menu Add
    WindowLong = GetWindowLong(hwnd, GWL_STYLE)
    SetWindowLong hwnd, GWL_STYLE, WindowLong Or WS_SYSMENU

    Exit Sub
 
ShowInTaskBar_Error:
    MsgBox Err.Description, vbCritical, "ModCommon1.ShowInTaskBar()"

End Sub



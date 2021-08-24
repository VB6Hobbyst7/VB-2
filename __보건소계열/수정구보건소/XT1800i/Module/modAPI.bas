Attribute VB_Name = "modAPI"
Option Explicit

'화면 해상도 가져오기
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Public Const SM_CXSCREEN = 0
    Public Const SM_CYSCREEN = 1
'Main Form Lock(System Menu Modify)
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Const SC_SIZE = &HF000
    Public Const SC_MAXIMIZE = &HF030
    Public Const MF_BYCOMMAND = &H0
    Public Const GWL_STYLE = -16
    Public Const WS_BORDER = &H800000
    Public Const WS_MAXIMIZEBOX = &H10000
    Public Const WS_THICKFRAME = &H40000
'Window Cursour Modify
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Window Lock
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'About Box 나타내기
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
'API Timer
Declare Function KillTimer& Lib "user32" (ByVal ahWnd&, ByVal nIDEvent&) ' The KillTimer function destroys the specified timer.
Declare Function SetTimer& Lib "user32" (ByVal ahWnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc&) ' The SetTimer function creates a timer with the specified time-out value.
'AlwaysOn 관련 정의
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Const HWND_NOTOPMOST = -2   'Not Always top
    Public Const HWND_TOPMOST = -1  'Always top
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOSIZE = &H1
'폼드래그 이동
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

'--------------------------------------------------------------
' Copyright ⓒ1996-2001 VBnet, Randy Birch, All Rights Reserved.
' Terms of use http://www.mvps.org/vbnet/terms/pages/terms.htm
'--------------------------------------------------------------


Public Const LVIF_INDENT As Long = &H10
Public Const LVIF_TEXT As Long = &H1
Public Const LVS_EX_FULLROWSELECT As Long = &H20
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_GETITEM As Long = (LVM_FIRST + 5)
Public Const LVM_SETITEM As Long = (LVM_FIRST + 6)
Public Const LVM_DELETEALLITEMS As Long = (LVM_FIRST + 9)
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)
Public Const ICC_LISTVIEW_CLASSES As Long = &H1

Public Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type
 
Public Type tagINITCOMMONCONTROLSEX
    dwSize As Long
    dwICC As Long
End Type

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMiliseconds As Long)

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long


'Public Declare Function SendMessage Lib "User32" '   Alias "SendMessageA" '  (ByVal hWnd As Long, '   ByVal wMsg As Long, '   ByVal wParam As Long, '   lParam As Any) As Long
       
'Public Declare Function LockWindowUpdate '   Lib "User32" (ByVal hwndLock As Long) As Long
       
'Public Declare Function UpdateWindow Lib "User32" '   (ByVal hWnd As Long) As Long

Public Function InitComctl32(dwFlags As Long) As Boolean

  Dim icc As tagINITCOMMONCONTROLSEX
  On Error GoTo Err_OldVersion
  
  icc.dwSize = Len(icc)
  icc.dwICC = dwFlags
  
 'VB will generate error 453 "Specified
 'DLL function not found" here if the new
 'version isn't installed and it can't find
 'the function's name. We'll hopefully be
 'able to load the old version below.
  InitComctl32 = InitCommonControlsEx(icc)
  Exit Function

Err_OldVersion:
  InitCommonControls

End Function




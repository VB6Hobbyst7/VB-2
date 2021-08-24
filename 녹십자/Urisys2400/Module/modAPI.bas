Attribute VB_Name = "modAPI"
Option Explicit

'ȭ�� �ػ� ��������
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
'About Box ��Ÿ����
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
'API Timer
Declare Function KillTimer& Lib "user32" (ByVal ahWnd&, ByVal nIDEvent&) ' The KillTimer function destroys the specified timer.
Declare Function SetTimer& Lib "user32" (ByVal ahWnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc&) ' The SetTimer function creates a timer with the specified time-out value.
'AlwaysOn ���� ����
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Const HWND_NOTOPMOST = -2   'Not Always top
    Public Const HWND_TOPMOST = -1  'Always top
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOSIZE = &H1
'���巡�� �̵�
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

'--------------------------------------------------------------
' Copyright ��1996-2001 VBnet, Randy Birch, All Rights Reserved.
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

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SND_APPLICATION = &H80         '  look for application specific association
Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_FILENAME = &H20000     '  name is a file name
Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Private Const SND_PURGE = &H40               '  purge non-static events for task
Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Private Const SND_SYNC = &H0         '  play synchronously (default)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long



Public Function SoundOn(SoungFlag As Boolean)
    If SoungFlag Then
        PlaySound App.Path & "\SWIT_SOUND.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
End Function


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




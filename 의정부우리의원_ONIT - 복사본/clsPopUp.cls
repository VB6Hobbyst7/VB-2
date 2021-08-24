VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINT
    X As Long
    y As Long
End Type
'
Private Const MF_ENABLED = &H0&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const TPM_RIGHTBUTTON = &H2&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_NONOTIFY = &H80&
Private Const TPM_RETURNCMD = &H100&
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal sCaption As String) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, nIgnored As Long) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
'
Public Function Popup(ParamArray Param()) As Long
    Dim iMenu As Long
    Dim hMenu As Long
    Dim nMenus As Long
    Dim p As POINT

    GetCursorPos p

    hMenu = CreatePopupMenu()

    nMenus = 1 + UBound(Param)

    For iMenu = 1 To nMenus
        If Trim$(CStr(Param(iMenu - 1))) = "-" Then
            AppendMenu hMenu, MF_SEPARATOR, iMenu, ""
        Else
            AppendMenu hMenu, MF_STRING + MF_ENABLED, iMenu, CStr(Param(iMenu - 1))
        End If
    Next iMenu

    iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, p.X, p.y, 0, GetForegroundWindow(), 0)

    DestroyMenu hMenu

    Popup = iMenu

End Function









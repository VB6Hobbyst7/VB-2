VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISFullMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMenus As clsIISMenus        '메뉴 컬렉션

'-----------------------------------------------------------------------------'
'   기능 : 메인폼의 풀다운메뉴 구성
'-----------------------------------------------------------------------------'
Public Sub SetFullMenu()
    Dim objMain As clsIISMainInfo

    Set objMain = New clsIISMainInfo
    With objMain.Main
        .mnuPrepare.Visible = mMenus("IIS100").Visible
        .mnuOrder.Visible = mMenus("IIS101").Visible
        .mnuCancel.Visible = mMenus("IIS102").Visible
        .mnuSep01.Visible = mMenus("IISSEP01").Visible
        .mnuWorklist.Visible = mMenus("IIS103").Visible

        .mnuInterface.Visible = mMenus("IIS200").Visible
        .mnuMachine1.Visible = mMenus("IIS201").Visible
        .mnuMachine2.Visible = mMenus("IIS202").Visible
        .mnuSep02.Visible = mMenus("IISSEP02").Visible
        .mnuErrors.Visible = mMenus("IIS203").Visible
    End With
    Set objMain = Nothing
End Sub

Private Sub Class_Initialize()
    Set mMenus = New clsIISMenus
    
    '## 검사준비
    mMenus.Add "IIS100", True
    mMenus.Add "IIS101", True
    mMenus.Add "IIS102", True
    mMenus.Add "IISSEP01", False
    mMenus.Add "IIS103", False
    
    '## 장비인터페이스
    mMenus.Add "IIS200", True
    mMenus.Add "IIS201", True
    mMenus.Add "IIS202", True
    mMenus.Add "IISSEP02", True
    mMenus.Add "IIS203", True
End Sub

Private Sub Class_Terminate()
    Set mMenus = Nothing
End Sub

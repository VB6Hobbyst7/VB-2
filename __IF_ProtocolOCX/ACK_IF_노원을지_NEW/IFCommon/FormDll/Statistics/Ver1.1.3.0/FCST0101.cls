VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCST0201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMachineCd As String

Public Property Let MachineCd(ByVal sData As String)
    mMachineCd = sData
    gsMachineCd = mMachineCd
End Property

Public Property Get MachineCd() As String
    mMachineCd = gsMachineCd
    MachineCd = mMachineCd
End Property

Public Sub Init()
    On Error GoTo ErrHandler
    
    Load frmStatistics
    RegEditCurFrmTitle "Statistics", frmStatistics.Caption
    frmStatistics.Show
    
    Exit Sub
ErrHandler:
    MsgBox "Init - Err(" & Err.Description & ")"
End Sub

Public Sub ShowForm()
    On Error GoTo ErrHandler
        
    Call ShowWindow(frmStatistics.hwnd, SW_SHOWDEFAULT)     'SW_SHOWMAXIMIZED)
        
    Exit Sub
ErrHandler:
End Sub

Public Sub SetMachineInfo(ByVal sMCd$, ByVal sMNm$)
    gsMachineCd = sMCd
    gsMachineNm = sMNm
End Sub


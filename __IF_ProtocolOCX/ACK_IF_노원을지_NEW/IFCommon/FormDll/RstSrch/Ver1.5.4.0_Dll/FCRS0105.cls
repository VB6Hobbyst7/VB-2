VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCRS0205"
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

Public Sub Init(Optional ByVal sUID$, Optional ByVal sPWD$, Optional ByVal sUserNm$, Optional ByVal sUserOther$)
    On Error GoTo ErrHandler
    
    Load frmRstSrch
    RegEditCurFrmTitle "RstSrch", frmRstSrch.Caption
    RegUserInfo sUID, sPWD, sUserNm, sUserOther
    frmRstSrch.Show
    
    Exit Sub
    
ErrHandler:
    MsgBox "Init ¿À·ù - (" & Err.Description & ")"
End Sub

Public Sub ShowForm()
    On Error GoTo ErrHandler
        
    Call ShowWindow(frmRstSrch.hwnd, SW_SHOWMAXIMIZED)
        
    Exit Sub
    
ErrHandler:
End Sub

Public Sub Term()
    On Error GoTo ErrHandler
    
    RegEditCurFrmTitle "RstSrch", ""
    Unload frmRstSrch
    
    Exit Sub
    
ErrHandler:
End Sub

Public Sub SetMachineInfo(ByVal sMCd$, ByVal sMNm$)
    gsMachineCd = sMCd
    gsMachineNm = sMNm
End Sub


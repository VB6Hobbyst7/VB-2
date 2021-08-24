VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCO0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub init(ByVal iFrmSeq As Integer)
    
    On Error GoTo ErrHandler
    
    Call HidePrevFrm
    
    Select Case iFrmSeq
    Case 1
        Load FGO0101
        RegEditCurFrmTitle FGO0101.Caption
        FGO0101.Top = 2700
        FGO0101.Left = 3500
        FGO0101.Show
    Case 2
        Load FGO0201
        RegEditCurFrmTitle FGO0201.Caption
        FGO0201.Top = 2700
        FGO0201.Left = 3500
        FGO0201.Show
    Case 3
        Load FGO0301
        RegEditCurFrmTitle FGO0301.Caption
        FGO0301.Top = 2700
        FGO0301.Left = 3500
        FGO0301.Show
    End Select
    
    Call InitializePart
    
    Exit Sub
    
ErrHandler:
    
End Sub

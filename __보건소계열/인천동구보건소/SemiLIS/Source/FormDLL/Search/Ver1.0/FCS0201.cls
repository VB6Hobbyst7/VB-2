VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCS0201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub Init()
    On Error GoTo ErrHandler
    
    Call HidePrevFrm
    
    Load FGS0201
    RegEditCurFrmTitle FGS0201.Caption
    
    'Top¼³Á¤
    If fMainTop < 0 Then
        FGS0201.Top = fMainHeight + fMainTop
    Else
        FGS0201.Top = fMainHeight - fMainTop
    End If
    
    FGS0201.Show
    
    Call InitializePart
    
    Exit Sub
    
ErrHandler:
End Sub


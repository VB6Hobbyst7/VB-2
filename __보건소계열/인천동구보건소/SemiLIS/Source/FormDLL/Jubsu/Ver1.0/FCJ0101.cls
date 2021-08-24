VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCJ0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub init()

    On Error GoTo ErrHandler
    
    Call HidePrevFrm
    
    Load FGJ0101
    RegEditCurFrmTitle FGJ0101.Caption
    
    'Top¼³Á¤
    If fMainTop < 0 Then
        FGJ0101.Top = fMainHeight + fMainTop
    Else
        FGJ0101.Top = fMainHeight - fMainTop
    End If
    
    FGJ0101.Show
    
    Call InitializePart
    
Exit Sub
ErrHandler:

End Sub

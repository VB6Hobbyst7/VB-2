VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCT0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub init()

On Error GoTo ErrHandler
    
    Call HidePrevFrm
    
    Load FGT0101
    RegEditCurFrmTitle FGT0101.Caption
    
    FGT0101.Top = 800
    FGT0101.Left = 60
    
    FGT0101.Show
    
    Call InitializePart
    
    Exit Sub
    
ErrHandler:
    
End Sub

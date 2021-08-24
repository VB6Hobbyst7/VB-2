VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCR0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mInitState As Integer

Public Property Get InitState() As Integer
    InitState = mInitState
End Property

Public Sub Init()
    On Error GoTo ErrHandler
    
    Call HidePrevFrm
    
    Load FGR0101
    RegEditCurFrmTitle FGR0101.Caption
    
    'Top¼³Á¤
    If fMainTop < 0 Then
        FGR0101.Top = fMainHeight + fMainTop
    Else
        FGR0101.Top = fMainHeight - fMainTop
    End If
    
    FGR0101.Show
    
    Call InitializePart
    
    Exit Sub
    
ErrHandler:
End Sub


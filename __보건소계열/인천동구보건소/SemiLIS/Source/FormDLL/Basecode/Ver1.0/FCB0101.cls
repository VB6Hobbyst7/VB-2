VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCB0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mInitState As Integer

Public Property Get InitState() As Integer
    InitState = mInitState
End Property

Public Sub Init(ByVal iFrmSeq As Integer)
    On Error GoTo ErrHandler
    
    Call HidePrevFrm
    
    If iFrmSeq = 1 Then
        Load FGB0101
        RegEditCurFrmTitle FGB0101.Caption
        FGB0101.Show
    ElseIf iFrmSeq = 2 Then
        Load FGB0201
        RegEditCurFrmTitle FGB0201.Caption
        FGB0201.Show
    ElseIf iFrmSeq = 3 Then
        Load FGB0301
        RegEditCurFrmTitle FGB0301.Caption
        FGB0301.Show
    ElseIf iFrmSeq = 4 Then
        Load FGB0401
        RegEditCurFrmTitle FGB0401.Caption
        FGB0401.Show
    ElseIf iFrmSeq = 5 Then
        Load FGB0501
        RegEditCurFrmTitle FGB0501.Caption
        FGB0501.Show
    ElseIf iFrmSeq = 6 Then
        Load FGB0601
        RegEditCurFrmTitle FGB0601.Caption
        FGB0601.Show
    ElseIf iFrmSeq = 7 Then
        Load FGB0701
        RegEditCurFrmTitle FGB0701.Caption
        FGB0701.Show
    ElseIf iFrmSeq = 8 Then
        Load FGB0801
        RegEditCurFrmTitle FGB0801.Caption
        FGB0801.Show
    ElseIf iFrmSeq = 9 Then
        Load FGB0901
        RegEditCurFrmTitle FGB0901.Caption
        FGB0901.Show
    End If
    
'¿¡·¯¾øÀÌ Load µÊ
    mInitState = 0
    
    Call InitializePart
    
    Exit Sub
    
ErrHandler:
    mInitState = -1
End Sub

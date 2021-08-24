VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarCtrlCd As String
Private mvarLevelCd As String   '--Level(H, M, L)
Private mvarTestCd As String
Private mvarTestNm As String
Private mvarEqpCd As String

Private mvarAvalVal As String
Private mvarRstUnit As String
Private mvarPanelFg As String

'Private mvarMyDb As DrDatabase
'
'
''Database Connection
'Public Property Get MyDb() As DrDatabase
'    Set MyDb = mvarMyDb
'End Property
'Public Property Set MyDb(ByVal vNewValue As DrDatabase)
'    Set mvarMyDb = vNewValue
'End Property

'Control코드
Public Property Get CtrlCd() As String
    CtrlCd = mvarCtrlCd
End Property
Public Property Let CtrlCd(ByVal vNewValue As String)
    mvarCtrlCd = vNewValue
End Property

'Level코드(H,M,L)
Public Property Get LevelCd() As String
    LevelCd = mvarLevelCd
End Property
Public Property Let LevelCd(ByVal vNewValue As String)
    mvarLevelCd = vNewValue
End Property

'검사코드
Public Property Get TestCd() As String
    TestCd = mvarTestCd
End Property
Public Property Let TestCd(ByVal vNewValue As String)
    mvarTestCd = vNewValue
End Property

'검사명
Public Property Get TestNm() As String
    TestNm = mvarTestNm
End Property
Public Property Let TestNm(ByVal vNewValue As String)
    mvarTestNm = vNewValue
End Property

'장비코드
Public Property Get EqpCd() As String
    EqpCd = mvarEqpCd
End Property
Public Property Let EqpCd(ByVal vNewValue As String)
    mvarEqpCd = vNewValue
End Property

'유효숫자
Public Property Get AvalVal() As String
    AvalVal = mvarAvalVal
End Property
Public Property Let AvalVal(ByVal vNewValue As String)
    mvarAvalVal = vNewValue
End Property

'결과단위
Public Property Get RstUnit() As String
    RstUnit = mvarRstUnit
End Property
Public Property Let RstUnit(ByVal vNewValue As String)
    mvarRstUnit = vNewValue
End Property

'세부항목여부
Public Property Get PanelFg() As String
    PanelFg = mvarPanelFg
End Property
Public Property Let PanelFg(ByVal vNewValue As String)
    mvarPanelFg = vNewValue
End Property


Public Sub ItemSave()
    
    Dim objSqlQc As clsLISSqlQc
    Dim objTLab022 As clsTLab022
    Dim SqlStmt As String
    
    Set objSqlQc = New clsLISSqlQc
    Set objTLab022 = New clsTLab022
    Call DataMove(objTLab022)
    SqlStmt = objSqlQc.SqlItemSave(objTLab022)
    
    dbconn.Execute SqlStmt
    
End Sub

Private Sub DataMove(ByRef objTLab022 As clsTLab022)
    With objTLab022
        .CtrlCd = mvarCtrlCd
        .LevelCd = mvarLevelCd
        .TestCd = mvarTestCd
        .EqpCd = mvarEqpCd
    End With
End Sub

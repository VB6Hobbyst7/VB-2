VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQcItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarCtrlCd As String
Private mvarLevelCd As String
Private mvarLotNo As String
Private mvarTestCd As String
Private mvarTestNm As String
Private mvarAbbrNm As String
Private mvarMeanVal As Double
Private mvarSdVal As Double
Private mvarAvalVal As Integer
Private mvarRefCd As String
Private mvarRstUnit As String
Private mvarCvVal As Double
Private mvarMinVal As Double
Private mvarMaxVal As Double
Private mvarSchedFg As String
Private mvarWmSet As String
Private mvarWmFg(1 To 7) As Integer
Private mvarCalFg As String
Private mvarEntDt As String
Private mvarEntId As String

Private mvarKeyString As String
Private mvarLevelNm As String
Private mvarSelected As Boolean

Private mvarRstDiv As String
Private mvarRaDiv As String
Private mvarRstType As String
Private mvarTestDiv As String
Private mvarPanelFg As String

Private mvarOpenDt As String
Private mvarExpDt As String

Private mvarTotCnt As Long
Private mvarAccCnt As Long
Private mvarRejCnt As Long

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

'Selected Flag
Public Property Get Selected() As Boolean
    Selected = mvarSelected
End Property

Public Property Let Selected(ByVal vNewValue As Boolean)
    mvarSelected = vNewValue
End Property

'Control코드
Public Property Get CtrlCd() As String
    CtrlCd = mvarCtrlCd
End Property

Public Property Let CtrlCd(ByVal vNewValue As String)
    mvarCtrlCd = vNewValue
End Property

'Level
Public Property Get LevelCd() As String
    LevelCd = mvarLevelCd
End Property
Public Property Let LevelCd(ByVal vNewValue As String)
    mvarLevelCd = vNewValue
End Property

'Level Name
Public Property Get LevelNm() As String
    LevelNm = mvarLevelNm
End Property
Public Property Let LevelNm(ByVal vNewValue As String)
    mvarLevelNm = vNewValue
End Property

'Key String
Public Property Get KeyString() As String
    KeyString = mvarKeyString
End Property
Public Property Let KeyString(ByVal vNewValue As String)
    mvarKeyString = vNewValue
End Property

'Lot Number
Public Property Get Lotno() As String
    Lotno = mvarLotNo
End Property
Public Property Let Lotno(ByVal vNewValue As String)
    mvarLotNo = vNewValue
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

'약어명
Public Property Get AbbrNm() As String
    AbbrNm = mvarAbbrNm
End Property
Public Property Let AbbrNm(ByVal vNewValue As String)
    mvarAbbrNm = vNewValue
End Property

'평균값
Public Property Get MeanVal() As Double
    MeanVal = mvarMeanVal
End Property
Public Property Let MeanVal(ByVal vNewValue As Double)
    mvarMeanVal = vNewValue
End Property

'표준편차
Public Property Get SdVal() As Double
    SdVal = mvarSdVal
End Property
Public Property Let SdVal(ByVal vNewValue As Double)
    mvarSdVal = vNewValue
End Property

'유효숫자
Public Property Get AvalVal() As Integer
    AvalVal = mvarAvalVal
End Property
Public Property Let AvalVal(ByVal vNewValue As Integer)
    mvarAvalVal = vNewValue
End Property

'코드참고치
Public Property Get RefCd() As String
    RefCd = mvarRefCd
End Property
Public Property Let RefCd(ByVal vNewValue As String)
    mvarRefCd = vNewValue
End Property

'결과단위
Public Property Get RstUnit() As String
    RstUnit = mvarRstUnit
End Property
Public Property Let RstUnit(ByVal vNewValue As String)
    mvarRstUnit = vNewValue
End Property

'기울기
Public Property Get CvVal() As Double
    CvVal = mvarCvVal
End Property
Public Property Let CvVal(ByVal vNewValue As Double)
    mvarCvVal = vNewValue
End Property

'최소값
Public Property Get MinVal() As Double
    MinVal = mvarMinVal
End Property
Public Property Let MinVal(ByVal vNewValue As Double)
    mvarMinVal = vNewValue
End Property

'최대값
Public Property Get MaxVal() As Double
    MaxVal = mvarMaxVal
End Property
Public Property Let MaxVal(ByVal vNewValue As Double)
    mvarMaxVal = vNewValue
End Property

'스케쥴여부
Public Property Get SchedFg() As String
    SchedFg = mvarSchedFg
End Property
Public Property Let SchedFg(ByVal vNewValue As String)
    mvarSchedFg = vNewValue
End Property

'Rule Setting
Public Property Get WmSet() As String
    WmSet = mvarWmSet
End Property
Public Property Let WmSet(ByVal vNewValue As String)
    mvarWmSet = vNewValue
End Property

'Rule Setting
Public Property Get WmFg(ByVal Index As Integer) As Integer
    WmFg = mvarWmFg(Index)
End Property
Public Property Let WmFg(ByVal Index As Integer, ByVal vNewValue As Integer)
    mvarWmFg(Index) = vNewValue
End Property

'계산여부
Public Property Get CalFg() As String
    CalFg = mvarCalFg
End Property
Public Property Let CalFg(ByVal vNewValue As String)
    mvarCalFg = vNewValue
End Property

'입력일
Public Property Get EntDt() As String
    EntDt = mvarEntDt
End Property
Public Property Let EntDt(ByVal vNewValue As String)
    mvarEntDt = vNewValue
End Property

'입력자
Public Property Get EntId() As String
    EntId = mvarEntId
End Property
Public Property Let EntId(ByVal vNewValue As String)
    mvarEntId = vNewValue
End Property

'결과유형
Public Property Get RstType() As String
    RstType = mvarRstType
End Property
Public Property Let RstType(ByVal vNewValue As String)
    mvarRstType = vNewValue
End Property

'Alternative/Required구분
Public Property Get RstDiv() As String
    RstDiv = mvarRstDiv
End Property
Public Property Let RstDiv(ByVal vNewValue As String)
    mvarRstDiv = vNewValue
End Property

'검사구분
Public Property Get TestDiv() As String
    TestDiv = mvarTestDiv
End Property
Public Property Let TestDiv(ByVal vNewValue As String)
    mvarTestDiv = vNewValue
End Property

'Panel Flag
Public Property Get PanelFg() As String
    PanelFg = mvarPanelFg
End Property
Public Property Let PanelFg(ByVal vNewValue As String)
    mvarPanelFg = vNewValue
End Property

'시작일
Public Property Get OpenDt() As String
    OpenDt = mvarOpenDt
End Property
Public Property Let OpenDt(ByVal vNewValue As String)
    mvarOpenDt = vNewValue
End Property

'폐기일
Public Property Get ExpDt() As String
    ExpDt = mvarExpDt
End Property
Public Property Let ExpDt(ByVal vNewValue As String)
    mvarExpDt = vNewValue
End Property

'Total Count
Public Property Get TotCnt() As Long
    TotCnt = mvarTotCnt
End Property
Public Property Let TotCnt(ByVal vNewValue As Long)
    mvarTotCnt = vNewValue
End Property

'Accept Count
Public Property Get AccCnt() As Long
    AccCnt = mvarAccCnt
End Property
Public Property Let AccCnt(ByVal vNewValue As Long)
    mvarAccCnt = vNewValue
End Property

'Reject Count
Public Property Get RejCnt() As Long
    RejCnt = mvarRejCnt
End Property
Public Property Let RejCnt(ByVal vNewValue As Long)
    mvarRejCnt = vNewValue
End Property

Private Sub DataMove(ByRef objTLab024 As clsTLab024)
    With objTLab024
        .CtrlCd = mvarCtrlCd
        .LevelCd = mvarLevelCd
        .Lotno = mvarLotNo
        .TestCd = mvarTestCd
        .MeanVal = mvarMeanVal
        .SdVal = mvarSdVal
        .AvalVal = mvarAvalVal
        .RefCd = mvarRefCd
        .RstUnit = mvarRstUnit
        .CvVal = mvarCvVal
        .MinVal = mvarMinVal
        .MaxVal = mvarMaxVal
        .SchedFg = mvarSchedFg
        .WmSet = mvarWmSet
        .CalFg = mvarCalFg
        .EntDt = mvarEntDt
        .EntId = mvarEntId
    End With
End Sub

Public Function ItemSave() As Boolean
    
    Dim objSqlQc As clsLISSqlQc
    Dim objTLab024 As clsTLab024
    Dim SqlStmt As String
    
    Set objSqlQc = New clsLISSqlQc
    Set objTLab024 = New clsTLab024
    Call DataMove(objTLab024)
    
    SqlStmt = objSqlQc.SqlQcItemSave(objTLab024)
    
    Set objTLab024 = Nothing
    Set objSqlQc = Nothing

On Error GoTo Err_Trap

    DBConn.Execute SqlStmt
    ItemSave = True
    Exit Function
    
Err_Trap:
    MsgBox Err.Description, vbExclamation
    ItemSave = False

End Function



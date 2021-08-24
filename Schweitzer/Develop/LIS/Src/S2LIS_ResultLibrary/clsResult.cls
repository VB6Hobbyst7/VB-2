VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsResultCol"
Attribute VB_Ext_KEY = "Member1" ,"LAB014"
Option Explicit
'
'/* Data Fields Property */
Private mvarWorkArea    As String
Private mvarAccDt       As String
Private mvarAccSeq      As String
Private mvarSpcYy       As String
Private mvarSpcNo       As String
Private mvarTestCd      As String
Private mvarRstVal      As String
Private mvarRstCd       As String
Private mvarRstUnit     As String
Private mvarHLDiv       As String
Private mvarDPDiv       As String
Private mvarSpcCd       As String
Private mvarPtInfo      As String
Private mvarLastRst     As String
Private mvarLastRstVal  As String
Private mvarLastVfyDt   As String
Private mvarLastVfyTm   As String
Private mvarLastVfyId   As String
Private mvarVfyDt       As String
Private mvarVfyTm       As String
Private mvarVfyId       As String
Private mvarMfyFg       As String
Private mvarGrpFg       As String
Private mvarAutoFg      As String
Private mvarTxtFg       As String
Private mvarExcFg       As String
Private mvarRstType     As String
Private mvarRstDiv      As String
Private mvarTextRst     As String
Private mvarSuppText    As String
Private mvarColDt       As String
Private mvarTxtType     As String
Private mvarTestNm      As String
Private mvarAvalVal     As String
Private mvarPanicFg     As String
Private mvarPanicFrVal  As Double
Private mvarPanicToVal  As Double
Private mvarArletFrVal  As Double
Private mvarArletToVal  As Double
' AMR 판정추가
Private mvarAMRFrVal    As String
Private mvarAMRToVal    As String
Private mvarDeltaFg     As String
Private mvarDeltaVal1   As Long
Private mvarDeltaVal2   As Long
Private mvarRefvalFrom  As Double
Private mvarRefValTo    As Double
Private mvarRefCd       As String
Private mvarOrdDt       As String
Private mvarOrdNo       As String
Private mvarOrdSeq      As String
Private mvarDetailFg    As String
Private mvarDataFg      As String
Private mvarEqpCd       As String
Private mvarOTmpCd      As String
Private mvarEqpNm       As String
Private mvarTransDtTm   As String
Private mvarMRstVal     As String
Private mvarMRstCd      As String
Private mvarMHLDiv      As String
Private mvarMDPDiv      As String
Private mvarMfyRsn      As String
Private mvarMRstType    As String
Private mvarFootNote    As String
Private mvarReqAltCnt   As String
Private mvarReqTotCnt   As String
Private mvarWorkSeq     As String

'/* 전주예수병원 추가변수 */ ====================================
Private mvarSlipCd      As String
Private mvarRsltType    As String
Private mvarWrkDiv      As String '-- '3':종건, '4':일건
'================================================================

'
'/* Class Default Property */
Private mvarRecordCount As Long
Private mvarEOF         As Boolean
Private mvarBOF         As Boolean
Private mvarKey         As String
Private mvarErrNo       As Long
Private mvarErrText     As String
Private m_INDEX         As Long
Private mvarResultCol   As clsResultCol

'/* Class Default Property */
Public Property Get ErrNo() As Long
    mvarErrNo = OraErr.Err
    ErrNo = mvarErrNo
End Property

Public Property Get ErrText() As String
    mvarErrText = OraErr.ErrText
    ErrText = mvarErrText
End Property

Public Property Let Key(ByVal vData As String)
    'Collection Class Key Setting
    vData = mvarWorkArea & mvarAccDt & mvarAccSeq & mvarTestCd
    mvarKey = vData
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Get BOF() As Boolean
    BOF = mvarBOF
End Property

Public Property Get EOF() As Boolean
    EOF = mvarEOF
End Property

Public Sub MoveFirst()
'Copy First Data Collection to Class Property
   CopyDataToProperty 1
   '
End Sub

Public Sub MoveLast()
'Copy Last Data Collection to Class Propertys
   CopyDataToProperty Me.ResultCol.Count
   '
End Sub

Public Sub MoveNext()
'Copy Next Data Collection to Class Propertys
   With Me
      CopyDataToProperty (m_INDEX + 1)
   End With
   '
End Sub

Public Sub MovePrevious()
'Copy Previous Data Collection to Class Propertys
   With Me
      CopyDataToProperty (m_INDEX - 1)
   End With
   '
End Sub

Public Property Get ResultCol() As clsResultCol
    If mvarResultCol Is Nothing Then
        Set mvarResultCol = New clsResultCol
    End If
    Set ResultCol = mvarResultCol
End Property

Public Property Set ResultCol(vData As clsResultCol)
    Set mvarResultCol = vData
End Property

Private Sub Class_Initialize()
   '
   m_INDEX = 0
End Sub

Private Sub Class_Terminate()
  Set mvarResultCol = Nothing
End Sub

Public Property Let RecordCount(ByVal vData As Long)
   If Me.ResultCol.Count > 0 Then
      mvarRecordCount = Me.ResultCol.Count
   Else
      mvarRecordCount = vData
   End If
End Property

Public Property Get RecordCount() As Long
    RecordCount = mvarRecordCount
End Property

'/* Data Fields Property */
Public Property Let SuppText(ByVal vData As String)
    mvarSuppText = vData
End Property

Public Property Get SuppText() As String
    SuppText = mvarSuppText
End Property

Public Property Let TextRst(ByVal vData As String)
    mvarTextRst = vData
End Property

Public Property Get TextRst() As String
    TextRst = mvarTextRst
End Property

Public Property Let RstDiv(ByVal vData As String)
    mvarRstDiv = vData
End Property

Public Property Get RstDiv() As String
    RstDiv = mvarRstDiv
End Property

Public Property Let RstType(ByVal vData As String)
    mvarRstType = vData
End Property

Public Property Get RstType() As String
    RstType = mvarRstType
End Property

Public Property Let TxtFg(ByVal vData As String)
    mvarTxtFg = vData
End Property

Public Property Get TxtFg() As String
    TxtFg = mvarTxtFg
End Property

Public Property Let ExcFg(ByVal vData As String)
    mvarExcFg = vData
End Property

Public Property Get ExcFg() As String
    ExcFg = mvarExcFg
End Property

Public Property Let GrpFg(ByVal vData As String)
    mvarGrpFg = vData
End Property

Public Property Get GrpFg() As String
    GrpFg = mvarGrpFg
End Property

Public Property Let AutoFg(ByVal vData As String)
    mvarAutoFg = vData
End Property

Public Property Get AutoFg() As String
    AutoFg = mvarAutoFg
End Property

Public Property Let MfyFg(ByVal vData As String)
    mvarMfyFg = vData
End Property

Public Property Get MfyFg() As String
    MfyFg = mvarMfyFg
End Property

Public Property Let VfyId(ByVal vData As String)
    mvarVfyId = vData
End Property

Public Property Get VfyId() As String
    VfyId = mvarVfyId
End Property

Public Property Let VfyTm(ByVal vData As String)
    mvarVfyTm = vData
End Property

Public Property Get VfyTm() As String
    VfyTm = mvarVfyTm
End Property

Public Property Let VfyDt(ByVal vData As String)
    mvarVfyDt = vData
End Property

Public Property Get VfyDt() As String
    VfyDt = mvarVfyDt
End Property

Public Property Let LastVfyId(ByVal vData As String)
    mvarLastVfyId = vData
End Property

Public Property Get LastVfyId() As String
    LastVfyId = mvarLastVfyId
End Property

Public Property Let LastVfyTm(ByVal vData As String)
    mvarLastVfyTm = vData
End Property

Public Property Get LastVfyTm() As String
    LastVfyTm = mvarLastVfyTm
End Property

Public Property Let LastVfyDt(ByVal vData As String)
    mvarLastVfyDt = vData
End Property

Public Property Get LastVfyDt() As String
    LastVfyDt = mvarLastVfyDt
End Property

Public Property Let LastRst(ByVal vData As String)
    mvarLastRst = vData
End Property

Public Property Get LastRst() As String
    LastRst = mvarLastRst
End Property

Public Property Let LastRstVal(ByVal vData As String)
    mvarLastRstVal = vData
End Property

Public Property Get LastRstVal() As String
    LastRstVal = mvarLastRstVal
End Property

Public Property Let PtInfo(ByVal vData As String)
    mvarPtInfo = vData
End Property

Public Property Get PtInfo() As String
    PtInfo = mvarPtInfo
End Property

Public Property Let SpcCd(ByVal vData As String)
    mvarSpcCd = vData
End Property

Public Property Get SpcCd() As String
    SpcCd = mvarSpcCd
End Property

Public Property Let DPDiv(ByVal vData As String)
    mvarDPDiv = vData
End Property

Public Property Get DPDiv() As String
    DPDiv = mvarDPDiv
End Property

Public Property Let HLDiv(ByVal vData As String)
    mvarHLDiv = vData
End Property

Public Property Get HLDiv() As String
    HLDiv = mvarHLDiv
End Property

Public Property Let RstUnit(ByVal vData As String)
    mvarRstUnit = vData
End Property

Public Property Get RstUnit() As String
    RstUnit = mvarRstUnit
End Property

Public Property Let RstCd(ByVal vData As String)
    mvarRstCd = vData
End Property

Public Property Get RstCd() As String
    RstCd = mvarRstCd
End Property

Public Property Let RstVal(ByVal vData As String)
    mvarRstVal = vData
End Property

Public Property Get RstVal() As String
    RstVal = mvarRstVal
End Property

Public Property Let TestCd(ByVal vData As String)
    mvarTestCd = vData
End Property

Public Property Get TestCd() As String
    TestCd = mvarTestCd
End Property

Public Property Let TxtType(ByVal vData As String)
    mvarTxtType = vData
End Property

Public Property Get TxtType() As String
    TxtType = mvarTxtType
End Property

Public Property Let TestNm(ByVal vData As String)
    mvarTestNm = vData
End Property

Public Property Get TestNm() As String
    TestNm = mvarTestNm
End Property

Public Property Let ColDt(ByVal vData As String)
    mvarColDt = vData
End Property

Public Property Get ColDt() As String
    ColDt = mvarColDt
End Property

Public Property Let SpcYy(ByVal vData As String)
    mvarSpcYy = vData
End Property

Public Property Get SpcYy() As String
    SpcYy = mvarSpcYy
End Property

Public Property Let SpcNo(ByVal vData As String)
    mvarSpcNo = vData
End Property

Public Property Get SpcNo() As String
    SpcNo = mvarSpcNo
End Property

Public Property Let AccSeq(ByVal vData As String)
    mvarAccSeq = vData
End Property

Public Property Get AccSeq() As String
    AccSeq = mvarAccSeq
End Property

Public Property Let AccDt(ByVal vData As String)
    mvarAccDt = vData
End Property

Public Property Get AccDt() As String
    AccDt = mvarAccDt
End Property

Public Property Let WorkArea(ByVal vData As String)
    mvarWorkArea = vData
End Property

Public Property Get WorkArea() As String
    WorkArea = mvarWorkArea
End Property

Public Property Let RefCd(ByVal vData As String)
    mvarRefCd = vData
End Property

Public Property Get RefCd() As String
    RefCd = mvarRefCd
End Property

Public Property Let RefValTo(ByVal vData As Double)
    mvarRefValTo = vData
End Property

Public Property Get RefValTo() As Double
    RefValTo = mvarRefValTo
End Property

Public Property Let RefValFrom(ByVal vData As Double)
    mvarRefvalFrom = vData
End Property

Public Property Get RefValFrom() As Double
    RefValFrom = mvarRefvalFrom
End Property

Public Property Let DeltaVal1(ByVal vData As Long)
    mvarDeltaVal1 = vData
End Property

Public Property Get DeltaVal1() As Long
    DeltaVal1 = mvarDeltaVal1
End Property

Public Property Let DeltaVal2(ByVal vData As Long)
    mvarDeltaVal2 = vData
End Property

Public Property Get DeltaVal2() As Long
    DeltaVal2 = mvarDeltaVal2
End Property

Public Property Let DeltaFg(ByVal vData As String)
    mvarDeltaFg = vData
End Property

Public Property Get DeltaFg() As String
    DeltaFg = mvarDeltaFg
End Property

Public Property Let PanicToVal(ByVal vData As Double)
    mvarPanicToVal = vData
End Property

Public Property Get PanicToVal() As Double
    PanicToVal = mvarPanicToVal
End Property

Public Property Let PanicFrVal(ByVal vData As Double)
    mvarPanicFrVal = vData
End Property

Public Property Get PanicFrVal() As Double
    PanicFrVal = mvarPanicFrVal
End Property


Public Property Let PanicFg(ByVal vData As String)
    mvarPanicFg = vData
End Property

Public Property Get PanicFg() As String
    PanicFg = mvarPanicFg
End Property
Public Property Let ArletToVal(ByVal vData As Double)
    mvarArletToVal = vData
End Property

Public Property Get ArletToVal() As Double
    ArletToVal = mvarArletToVal
End Property

Public Property Let ArletFrVal(ByVal vData As Double)
    mvarArletFrVal = vData
End Property

Public Property Get ArletFrVal() As Double
    ArletFrVal = mvarArletFrVal
End Property
' AMR 판정 추가
Public Property Let AMRToVal(ByVal vData As String)
    mvarAMRToVal = vData
End Property
' AMR 판정 추가
Public Property Get AMRToVal() As String
    AMRToVal = mvarAMRToVal
End Property
' AMR 판정 추가
Public Property Let AMRFrVal(ByVal vData As String)
    mvarAMRFrVal = vData
End Property
' AMR 판정 추가
Public Property Get AMRFrVal() As String
    AMRFrVal = mvarAMRFrVal
End Property

Public Property Let AvalVal(ByVal vData As String)
    mvarAvalVal = vData
End Property

Public Property Get AvalVal() As String
    AvalVal = mvarAvalVal
End Property

Public Property Let OrdDt(ByVal vData As Long)
    mvarOrdDt = vData
End Property

Public Property Get OrdDt() As Long
    OrdDt = mvarOrdDt
End Property

Public Property Let OrdNo(ByVal vData As Long)
    mvarOrdNo = vData
End Property

Public Property Get OrdNo() As Long
    OrdNo = mvarOrdNo
End Property

Public Property Let OrdSeq(ByVal vData As Long)
    mvarOrdSeq = vData
End Property

Public Property Get OrdSeq() As Long
    OrdSeq = mvarOrdSeq
End Property

Public Property Let DetailFg(ByVal vData As String)
    mvarDetailFg = vData
End Property

Public Property Get DetailFg() As String
    DetailFg = mvarDetailFg
End Property

Public Property Let TransDtTm(ByVal vData As String)
    mvarTransDtTm = vData
End Property

Public Property Get TransDtTm() As String
    TransDtTm = mvarTransDtTm
End Property

Public Property Let EqpNm(ByVal vData As String)
    mvarEqpNm = vData
End Property

Public Property Get EqpNm() As String
    EqpNm = mvarEqpNm
End Property

Public Property Let EqpCd(ByVal vData As String)
    mvarEqpCd = vData
End Property

Public Property Get EqpCd() As String
    EqpCd = mvarEqpCd
End Property

Public Property Let OTmpCd(ByVal vData As String)
    mvarOTmpCd = vData
End Property

Public Property Get OTmpCd() As String
    OTmpCd = mvarOTmpCd
End Property

Public Property Let MRstType(ByVal vData As String)
    mvarMRstType = vData
End Property

Public Property Get MRstType() As String
    MRstType = mvarMRstType
End Property

Public Property Let MfyRsn(ByVal vData As String)
    mvarMfyRsn = vData
End Property

Public Property Get MfyRsn() As String
    MfyRsn = mvarMfyRsn
End Property

Public Property Let MDPDiv(ByVal vData As String)
    mvarMDPDiv = vData
End Property

Public Property Get MDPDiv() As String
    MDPDiv = mvarMDPDiv
End Property

Public Property Let MHLDiv(ByVal vData As String)
    mvarMHLDiv = vData
End Property

Public Property Get MHLDiv() As String
    MHLDiv = mvarMHLDiv
End Property

Public Property Let MRstCd(ByVal vData As String)
    mvarMRstCd = vData
End Property

Public Property Get MRstCd() As String
    MRstCd = mvarMRstCd
End Property

Public Property Let MRstVal(ByVal vData As String)
    mvarMRstVal = vData
End Property

Public Property Get MRstVal() As String
    MRstVal = mvarMRstVal
End Property

Public Property Let ReqAltCnt(ByVal vData As String)
    mvarReqAltCnt = vData
End Property

Public Property Get ReqAltCnt() As String
    ReqAltCnt = mvarReqAltCnt
End Property

Public Property Let ReqTotCnt(ByVal vData As String)
    mvarReqTotCnt = vData
End Property

Public Property Get ReqTotCnt() As String
    ReqTotCnt = mvarReqTotCnt
End Property

Public Property Let WorkSeq(ByVal vData As String)
    mvarWorkSeq = vData
End Property

Public Property Get WorkSeq() As String
    WorkSeq = mvarWorkSeq
End Property

Public Property Let FootNote(ByVal vData As String)
   mvarFootNote = vData
End Property

Public Property Get FootNote() As String
   FootNote = mvarFootNote
End Property

Public Property Let DataFg(ByVal vData As String)
   mvarDataFg = vData
End Property

Public Property Get DataFg() As String
   DataFg = mvarDataFg
End Property


'/* 전주예수병원 추가변수 */ ====================================
Public Property Let SlipCd(ByVal vData As String)
    mvarSlipCd = vData
End Property

Public Property Get SlipCd() As String
    SlipCd = mvarSlipCd
End Property

Public Property Let RsltType(ByVal vData As String)
    mvarRsltType = vData
End Property

Public Property Get RsltType() As String
    RsltType = mvarRsltType
End Property

Public Property Let WrkDiv(ByVal vData As String)
    mvarWrkDiv = vData
End Property

Public Property Get WrkDiv() As String
    WrkDiv = mvarWrkDiv
End Property
'================================================================


Private Sub CopyDataToProperty(ByVal index As Long)
   'With ResultCol.Item(index)
   If (index < 1) Or (index > Me.RecordCount) Then
         If index < 1 Then
            m_INDEX = 0
         Else
            m_INDEX = Me.RecordCount + 1
         End If
         '
         mvarWorkArea = ""
         mvarAccDt = ""
         mvarAccSeq = ""
         mvarTestCd = ""
         mvarOrdDt = ""
         mvarOrdNo = ""
         mvarOrdSeq = ""
         mvarDetailFg = ""
         mvarRstVal = ""
         mvarRstCd = ""
         mvarRstUnit = ""
         mvarHLDiv = ""
         mvarDPDiv = ""
         mvarSpcCd = ""
         mvarPtInfo = ""
         mvarLastRst = ""
         mvarLastRstVal = ""
         mvarLastVfyDt = ""
         mvarLastVfyTm = ""
         mvarLastVfyId = ""
         mvarVfyDt = ""
         mvarVfyTm = ""
         mvarVfyId = ""
         mvarMfyFg = ""
         mvarGrpFg = ""
         mvarAutoFg = ""
         mvarTxtFg = ""
         mvarExcFg = ""
         mvarRstType = ""
         mvarRstDiv = ""
         mvarTextRst = ""
         mvarSuppText = ""
         mvarColDt = ""
         mvarTxtType = ""
         mvarTestNm = ""
         mvarAvalVal = ""
         mvarPanicFg = ""
         mvarPanicFrVal = ""
         mvarPanicToVal = ""
         mvarDeltaFg = ""
         mvarDeltaVal1 = ""
         mvarDeltaVal2 = ""
         mvarRefvalFrom = ""
         mvarRefValTo = ""
         mvarRefCd = ""
         mvarEqpCd = ""
         mvarOTmpCd = ""
         mvarEqpNm = ""
         mvarTransDtTm = ""
         mvarMRstVal = ""
         mvarMRstCd = ""
         mvarMHLDiv = ""
         mvarMDPDiv = ""
         mvarMfyRsn = ""
         mvarMRstType = ""
         mvarFootNote = ""
         mvarReqAltCnt = ""
         mvarWorkSeq = ""
         mvarSlipCd = ""
         mvarRsltType = ""
         mvarWrkDiv = ""
         '*
         '
         'Collection Class Key Setting
         mvarKey = vbNullString
         mvarRecordCount = Me.ResultCol.Count
         m_INDEX = index
   Else
      With Me.ResultCol.Item(index)
         '
         mvarWorkArea = .WorkArea
         mvarAccDt = .AccDt
         mvarAccSeq = .AccSeq
         mvarTestCd = .TestCd
         mvarOrdDt = .OrdDt
         mvarOrdNo = .OrdNo
         mvarOrdSeq = .OrdSeq
         mvarDetailFg = .DetailFg
         mvarRstVal = .RstVal
         mvarRstCd = .RstCd
         mvarRstUnit = .RstUnit
         mvarHLDiv = .HLDiv
         mvarDPDiv = .DPDiv
         mvarSpcCd = .SpcCd
         mvarPtInfo = .PtInfo
         mvarLastRst = .LastRst
         mvarLastRstVal = .LastRstVal
         mvarLastVfyDt = .LastVfyDt
         mvarLastVfyTm = .LastVfyTm
         mvarLastVfyId = .LastVfyId
         mvarVfyDt = .VfyDt
         mvarVfyTm = .VfyTm
         mvarVfyId = .VfyId
         mvarMfyFg = .MfyFg
         mvarGrpFg = .GrpFg
         mvarAutoFg = .AutoFg
         mvarTxtFg = .TxtFg
         mvarExcFg = .ExcFg
         mvarRstType = .RstType
         mvarRstDiv = .RstDiv
         mvarTextRst = .TextRst
         mvarSuppText = .SuppText
         mvarColDt = .ColDt
         mvarTxtType = .TxtType
         mvarTestNm = .TestNm
         mvarAvalVal = .AvalVal
         mvarPanicFg = .PanicFg
         mvarPanicFrVal = .PanicFrVal
         mvarPanicToVal = .PanicToVal
         mvarDeltaFg = .DeltaFg
         mvarDeltaVal1 = .DeltaVal1
         mvarDeltaVal2 = .DeltaVal2
         mvarRefvalFrom = .RefValFrom
         mvarRefValTo = .RefValTo
         mvarRefCd = .RefCd
         mvarEqpCd = .EqpCd
         mvarOTmpCd = .OTmpCd
         mvarEqpNm = .EqpNm
         mvarTransDtTm = .TransDtTm
         mvarMRstVal = .MRstVal
         mvarMRstCd = .MRstCd
         mvarMHLDiv = .MHLDiv
         mvarMDPDiv = .MDPDiv
         mvarMfyRsn = .MfyRsn
         mvarMRstType = .MRstType
         mvarFootNote = .FootNote
         mvarReqAltCnt = .ReqAltCnt
         mvarWorkSeq = .WorkSeq
         mvarSlipCd = .SlipCd
         mvarRsltType = .RsltType
         mvarWrkDiv = .WrkDiv
         mvarKey = mvarWorkArea & mvarAccDt & mvarAccSeq & mvarTestCd
         mvarRecordCount = Me.ResultCol.Count
         m_INDEX = index
      End With
   End If
   '
   With Me.ResultCol
      If .Count > 0 Then
         Select Case m_INDEX
            Case Is < 1
               mvarBOF = True
               mvarEOF = False
            Case Is > .Count
               mvarEOF = True
               mvarBOF = False
            Case Else
               mvarBOF = False
               mvarEOF = False
         End Select
      Else
         mvarEOF = True
         mvarBOF = True
      End If
   End With
   '
End Sub



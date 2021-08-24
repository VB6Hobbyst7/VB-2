VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResultCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsResult"
Attribute VB_Ext_KEY = "Member0" ,"clsResult"
Option Explicit
'
Private mCol As Collection

'Public Function Add(WorkArea As String, AccDt As String, AccSeq As String, _
TestCd As String, RstVal As String, RstCd As String, RstUnit As String, HLDiv As String, _
DPDiv As String, SpcCd As String, PtInfo As String, LastRst As String, LastVfyDt As String, _
LastVfyTm As String, LastVfyId As String, VfyDt As String, VfyTm As String, VfyId As String, _
MfyFg As String, GrpFg As String, AutoFg As String, TxtFg As String, ExcFg As String, _
RstType As String, RstDiv As String, TextRst As String, SuppText As String, ColDt As String, _
TxtType As String, TestNm As String, AvalVal As String, PanicFg As String, _
PanicFrVal As Double, PanicToVal As Double, DeltaFg As String, DeltaVal As Long, _
RefValFrom As Double, RefValTo As Double, RefCd As String, OrdDt As String, _
OrdNo As String, OrdSeq As String, DetailFg As String, DataFg As String, EqpCd As String, _
OTmpCd As String, EqpNm As String, TransDtTm As String, MRstVal As String, MRstCd As String, _
MHLDiv As String, MDPDiv As String, MfyRsn As String, MRstType As String, FootNote As String, _
ReqAltCnt As String, RecordCount As Long, ResultCol As clsResultCol, Optional sKey As String) As clsResult

Public Function Add(ResultCol As clsResultCol, sKey As String, ParamArray NewVal() As Variant) As clsResult
Dim objNewMember As clsResult
   Set objNewMember = New clsResult
   '
   With objNewMember
      .WorkArea = NewVal(0)  'WorkArea
      .AccDt = NewVal(1)  'AccDt
      .AccSeq = NewVal(2)  'AccSeq
      .TestCd = NewVal(3)  'TestCd
      .RstVal = NewVal(4)  'RstVal
      .RstCd = NewVal(5)  'RstCd
      .RstUnit = NewVal(6)  'RstUnit
      .HLDiv = NewVal(7)  'HLDiv
      .DPDiv = NewVal(8)  'DPDiv
      .SpcCd = NewVal(9)  'SpcCd
      .PtInfo = NewVal(10)  'PtInfo
      .LastRst = NewVal(11)  'LastRst
      .LastVfyDt = NewVal(12)  'LastVfyDt
      .LastVfyTm = NewVal(13)  'LastVfyTm
      .LastVfyId = NewVal(14)  'LastVfyId
      .VfyDt = NewVal(15)  'VfyDt
      .VfyTm = NewVal(16)  'VfyTm
      .VfyId = NewVal(17)  'VfyId
      .MfyFg = NewVal(18)  'MfyFg
      .GrpFg = NewVal(19)  'GrpFg
      .AutoFg = NewVal(20)  'AutoFg
      .TxtFg = NewVal(21)  'TxtFg
      .ExcFg = NewVal(22)  'ExcFg
      .RstType = NewVal(23)  'RstType
      .RstDiv = NewVal(24)  'RstDiv
      .SuppText = NewVal(25)  'SuppText
      .TextRst = NewVal(26)  'TextRst
      .ColDt = NewVal(27)  'ColDt
      .TxtType = NewVal(28)  'TxtType
      .TestNm = NewVal(29)  'TestNm
      .AvalVal = NewVal(30)  'AvalVal
      .PanicFg = NewVal(31)  'PanicFg
      .PanicFrVal = NewVal(32)  'PanicFrVal
      .PanicToVal = NewVal(33)  'PanicToVal
      .DeltaFg = NewVal(34)  'DeltaFg
      .DeltaVal1 = NewVal(35)  'DeltaVal
      .DeltaVal2 = NewVal(36)
      .RefValFrom = NewVal(37)  'RefValFrom
      .RefValTo = NewVal(38)  'RefValTo
      .RefCd = NewVal(39)  'RefCd
      .OrdDt = NewVal(40)  'OrdDt
      .OrdNo = NewVal(41)  'OrdNo
      .OrdSeq = NewVal(42)  'OrdSeq
      .DetailFg = NewVal(43)  'DetailFg
      .DataFg = NewVal(44)  'DataFg
      .EqpCd = NewVal(45)  'EqpCd
      .OTmpCd = NewVal(46)  'OTmpCd
      .EqpNm = NewVal(47)  'EqpNm
      .TransDtTm = NewVal(48)  'TransDtTm
      .MRstVal = NewVal(49)  'MRstVal
      .MRstCd = NewVal(50)  'MRstCd
      .MHLDiv = NewVal(51)  'MHLDiv
      .MDPDiv = NewVal(52)  'MDPDiv
      .MfyRsn = NewVal(53)  'MfyRsn
      .MRstType = NewVal(54)  'MRstType
      .FootNote = NewVal(55)  'FootNote
      .ReqAltCnt = NewVal(56)  'ReqAltCnt
      .RecordCount = NewVal(57)  'RecordCount
      .WorkSeq = NewVal(58)
      .ArletFrVal = NewVal(59)
      .ArletToVal = NewVal(60)
      .SlipCd = NewVal(61)
      .RsltType = NewVal(62)
      .WrkDiv = NewVal(63)
      .LastRstVal = NewVal(64)
      .AMRFrVal = NewVal(65)
      .AMRToVal = NewVal(66)
   End With
   '
   Set objNewMember.ResultCol = ResultCol
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    '
    Set Add = objNewMember
    Set objNewMember = Nothing
    '
End Function

Public Property Get Item(vntIndexKey As Variant) As clsResult
Attribute Item.VB_UserMemId = 0
  On Error Resume Next
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub


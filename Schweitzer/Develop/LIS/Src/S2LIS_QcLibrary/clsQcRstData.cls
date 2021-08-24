VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQcRstData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarWorkArea As String  '--WorkArea(접수번호용)
Private mvarAccDt As String     '--접수일자(접수번호용)
Private mvarAccSeq As String    '--접수순번(접수번호용)
Private mvarTestCd As String    '--검사항목코드

Private mvarRstVal As Double    '--결과(Numeric)
Private mvarRstCd As String     '--결과(Alphabetic)
Private mvarRstCdNm As String   '--결과명(Alphabetic)
Private mvarRstUnit As String   '--결과단위
Private mvarRstType As String   '--결과유형
Private mvarRstDiv As String    '--Alternative/Required item여부
Private mvarRaDiv As String     '--Reject여부
Private mvarRaName As String    '--Reject여부
Private mvarRaColor As Long     '--Reject Color
Private mvarDetailFg As String  '--상세항목Seq

Private mvarRcvDt As String     '--접수일자
Private mvarRcvTm As String     '--접수시간
Private mvarVfyDt As String     '--결과확인일자
Private mvarVfyTm As String     '--결과확인시간
Private mvarVfyId As String     '--결과확인자
Private mvarMfyFg As String     '--수정여부
Private mvarTxtFg As String     '--Text결과여부
Private mvarAutoFg As String    '--장비전송여부
Private mvarEqpCd As String     '--장비코드
Private mvarMethod As String    '--검사방법
Private mvarMethodNm As String  '--검사방법

Private mvarRstText As String   '--결과텍스트

Private mvarTestNm As String
Private mvarAbbrNm5 As String
Private mvarAbbrNm10 As String

Private mvarLevelCd As String   '-- Level Cd
Private mvarMeanVal As Double   '-- 평균
Private mvarSdVal As Double     '-- 표준편차
Private mvarAvalVal As Integer  '-- 유효숫자
Private mvarRefCd As String     '-- Alpha결과참고치
Private mvarCvVal As Double     '-- CV값
Private mvarMinVal As Double    '-- 최소값
Private mvarMaxVal As Double    '-- 최대값

Private mvarWmSet As String     '-- Multi-Rule Setting

Private mvarRejectMsg As String
Private mvarRejectSeq As Integer

Private mvarKeyString As String

Private dbl1SD_P As Double      '--(+)1SD
Private dbl2SD_P As Double      '--(+)2SD
Private dbl3SD_P As Double      '--(+)3SD
Private dbl1SD_N As Double      '--(-)1SD
Private dbl2SD_N As Double      '--(-)2SD
Private dbl3SD_N As Double      '--(-)3SD

Private mHisCol As Collection
Private mvarRstHistory(1 To 10) As Double   '-- 과거 10회 결과
Private mvarRstHistoryRaDiv(1 To 10) As String   '-- 과거 10회 결과판정
Private mvarHistoryCnt As Integer
Private mvarHistoryFg As Boolean
Private mvarMultiFg As Boolean

'Private mvarMyDb As DrDatabase

Private Const CS_TotCnt = 10
Private Const CS_ChkCnt = 4

'Append By Legends
'2001/09/18
Private mvarCtrlCd As String
'Private mvarLevelCd As String
Private mvarLotNo As String

Public Property Let CtrlCd(ByVal vData As String)
    mvarCtrlCd = vData
End Property

Public Property Get CtrlCd() As String
    CtrlCd = mvarCtrlCd
End Property

'Public Property Let LevelCd(ByVal vData As String)
'    mvarLevelCd = vData
'End Property
'
'Public Property Get LevelCd() As String
'    LevelCd = mvarLevelCd
'End Property

Public Property Let Lotno(ByVal vData As String)
    mvarLotNo = vData
End Property

Public Property Get Lotno() As String
    Lotno = mvarLotNo
End Property

'Database Connection
'Public Property Get MyDb() As DrDatabase
'    Set MyDb = mvarMyDb
'End Property
'Public Property Set MyDb(ByVal vNewValue As DrDatabase)
'    Set mvarMyDb = vNewValue
'End Property

'WorkArea
Public Property Get WorkArea() As String
    WorkArea = mvarWorkArea
End Property
Public Property Let WorkArea(ByVal vNewValue As String)
    mvarWorkArea = vNewValue
End Property

'AccDt
Public Property Get AccDt() As String
    AccDt = mvarAccDt
End Property
Public Property Let AccDt(ByVal vNewValue As String)
    mvarAccDt = vNewValue
End Property

'AccSeq
Public Property Get AccSeq() As String
    AccSeq = mvarAccSeq
End Property
Public Property Let AccSeq(ByVal vNewValue As String)
    mvarAccSeq = vNewValue
End Property

'검사코드
Public Property Get TestCd() As String
    TestCd = mvarTestCd
End Property
Public Property Let TestCd(ByVal vNewValue As String)
    mvarTestCd = vNewValue
End Property

'결과
Public Property Get RstVal() As Double
    RstVal = mvarRstVal
End Property
Public Property Let RstVal(ByVal vNewValue As Double)
    mvarRstVal = vNewValue
End Property

'결과코드
Public Property Get RstCd() As String
    RstCd = mvarRstCd
End Property
Public Property Let RstCd(ByVal vNewValue As String)
    mvarRstCd = vNewValue
End Property

'결과명
Public Property Get RstCdNm() As String
    RstCdNm = mvarRstCdNm
End Property
Public Property Let RstCdNm(ByVal vNewValue As String)
    mvarRstCdNm = vNewValue
End Property

'단위
Public Property Get RstUnit() As String
    RstUnit = mvarRstUnit
End Property
Public Property Let RstUnit(ByVal vNewValue As String)
    mvarRstUnit = vNewValue
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

'결과유형
Public Property Get RstType() As String
    RstType = mvarRstType
End Property
Public Property Let RstType(ByVal vNewValue As String)
    mvarRstType = vNewValue
End Property

'Alternative/Required
Public Property Get RstDiv() As String
    RstDiv = mvarRstDiv
End Property
Public Property Let RstDiv(ByVal vNewValue As String)
    mvarRstDiv = vNewValue
End Property

'텍스트
Public Property Get RstText() As String
    RstText = mvarRstText
End Property
Public Property Let RstText(ByVal vNewValue As String)
    mvarRstText = vNewValue
End Property

'Reject/Accept 구분
Public Property Get RaDiv() As String
    RaDiv = mvarRaDiv
End Property
Public Property Let RaDiv(ByVal vNewValue As String)
    mvarRaDiv = vNewValue
End Property

'Reject/Accept 구분명
Public Property Get RaName() As String
    RaName = mvarRaName
End Property
Public Property Let RaName(ByVal vNewValue As String)
    mvarRaName = vNewValue
End Property

'Reject/Accept 구분 Color
Public Property Get RaColor() As Long
    RaColor = mvarRaColor
End Property
Public Property Let RaColor(ByVal vNewValue As Long)
    mvarRaColor = vNewValue
End Property

'상세항목Seq
Public Property Get DetailFg() As String
    DetailFg = mvarDetailFg
End Property
Public Property Let DetailFg(ByVal vNewValue As String)
    mvarDetailFg = vNewValue
End Property

'결과확인일
Public Property Get VfyDt() As String
    VfyDt = mvarVfyDt
End Property
Public Property Let VfyDt(ByVal vNewValue As String)
    mvarVfyDt = vNewValue
End Property

'접수일
Public Property Get RcvDt() As String
    RcvDt = mvarRcvDt
End Property
Public Property Let RcvDt(ByVal vNewValue As String)
    mvarRcvDt = vNewValue
End Property

'접수시간
Public Property Get RcvTm() As String
    RcvTm = mvarRcvTm
End Property
Public Property Let RcvTm(ByVal vNewValue As String)
    mvarRcvTm = vNewValue
End Property

'결과확인시간
Public Property Get VfyTm() As String
    VfyTm = mvarVfyTm
End Property
Public Property Let VfyTm(ByVal vNewValue As String)
    mvarVfyTm = vNewValue
End Property

'결과확인자
Public Property Get VfyId() As String
    VfyId = mvarVfyId
End Property
Public Property Let VfyId(ByVal vNewValue As String)
    mvarVfyId = vNewValue
End Property

'수정여부
Public Property Get MfyFg() As String
    MfyFg = mvarMfyFg
End Property
Public Property Let MfyFg(ByVal vNewValue As String)
    mvarMfyFg = vNewValue
End Property

'Text결과여부
Public Property Get TxtFg() As String
    TxtFg = mvarTxtFg
End Property
Public Property Let TxtFg(ByVal vNewValue As String)
    mvarTxtFg = vNewValue
End Property

'장비전송여부
Public Property Get AutoFg() As String
    AutoFg = mvarAutoFg
End Property
Public Property Let AutoFg(ByVal vNewValue As String)
    mvarAutoFg = vNewValue
End Property

'장비코드
Public Property Get EqpCd() As String
    EqpCd = mvarEqpCd
End Property
Public Property Let EqpCd(ByVal vNewValue As String)
    mvarEqpCd = vNewValue
End Property

'검사방법
Public Property Get Method() As String
    Method = mvarMethod
End Property
Public Property Let Method(ByVal vNewValue As String)
    mvarMethod = vNewValue
End Property

'검사방법명
Public Property Get MethodNm() As String
    MethodNm = mvarMethodNm
End Property
Public Property Let MethodNm(ByVal vNewValue As String)
    mvarMethodNm = vNewValue
End Property

'검사명
Public Property Get TestNm() As String
    TestNm = mvarTestNm
End Property
Public Property Let TestNm(ByVal vNewValue As String)
    mvarTestNm = vNewValue
End Property

'약어명5
Public Property Get AbbrNm5() As String
    AbbrNm5 = mvarAbbrNm5
End Property
Public Property Let AbbrNm5(ByVal vNewValue As String)
    mvarAbbrNm5 = vNewValue
End Property

'약어명10
Public Property Get AbbrNm10() As String
    AbbrNm10 = mvarAbbrNm10
End Property
Public Property Let AbbrNm10(ByVal vNewValue As String)
    mvarAbbrNm10 = vNewValue
End Property

'Level Code
Public Property Get LevelCd() As String
    LevelCd = mvarLevelCd
End Property
Public Property Let LevelCd(ByVal vNewValue As String)
    mvarLevelCd = vNewValue
End Property

'평균값
Public Property Get MeanVal() As Double
    MeanVal = mvarMeanVal
End Property
Public Property Let MeanVal(ByVal vNewValue As Double)
    mvarMeanVal = vNewValue
End Property

'SD값
Public Property Get SdVal() As Double
    SdVal = mvarSdVal
End Property
Public Property Let SdVal(ByVal vNewValue As Double)
    mvarSdVal = vNewValue
End Property

'1SD값(+)
Public Property Get P_1SdVal() As Double
    P_1SdVal = dbl1SD_P
End Property
Public Property Let P_1SdVal(ByVal vNewValue As Double)
    dbl1SD_P = vNewValue
End Property

'2SD값(+)
Public Property Get P_2SdVal() As Double
    P_2SdVal = dbl2SD_P
End Property
Public Property Let P_2SdVal(ByVal vNewValue As Double)
    dbl2SD_P = vNewValue
End Property

'3SD값(+)
Public Property Get P_3SdVal() As Double
    P_3SdVal = dbl3SD_P
End Property
Public Property Let P_3SdVal(ByVal vNewValue As Double)
    dbl3SD_P = vNewValue
End Property

'1SD값(-)
Public Property Get N_1SdVal() As Double
    N_1SdVal = dbl1SD_N
End Property
Public Property Let N_1SdVal(ByVal vNewValue As Double)
    dbl1SD_N = vNewValue
End Property

'2SD값(-)
Public Property Get N_2SdVal() As Double
    N_2SdVal = dbl2SD_N
End Property
Public Property Let N_2SdVal(ByVal vNewValue As Double)
    dbl2SD_N = vNewValue
End Property

'3SD값(-)
Public Property Get N_3SdVal() As Double
    N_3SdVal = dbl3SD_N
End Property
Public Property Let N_3SdVal(ByVal vNewValue As Double)
    dbl3SD_N = vNewValue
End Property

'CV값
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

'Multi-Rule Setting
Public Property Get WsSet() As String
    WsSet = mvarWmSet
End Property
Public Property Let WsSet(ByVal vNewValue As String)
    mvarWmSet = vNewValue
End Property

'Reject 사유
Public Property Get RejectMsg() As String
    RejectMsg = mvarRejectMsg
End Property

'Reject Reason
Public Property Get RejectSeq() As Integer
    RejectSeq = mvarRejectSeq
End Property

'과거결과
Public Property Get RstHistory(ByVal Index As Integer) As Double
    RstHistory = mvarRstHistory(Index)
End Property

Public Property Let RstHistory(ByVal Index As Integer, ByVal vNewValue As Double)
    mvarRstHistory(Index) = vNewValue
End Property

'과거 결과 판정
Public Property Get RstHistoryRaDiv(ByVal Index As Integer) As String
    RstHistoryRaDiv = mvarRstHistoryRaDiv(Index)
End Property

Public Property Let RstHistoryRaDiv(ByVal Index As Integer, ByVal vNewValue As String)
    mvarRstHistoryRaDiv(Index) = vNewValue
End Property

'과거결과Count
Public Property Get HistoryCnt() As Integer
    HistoryCnt = mvarHistoryCnt
End Property
Public Property Let HistoryCnt(ByVal vNewValue As Integer)
    mvarHistoryCnt = vNewValue
End Property

'과거결과Flag
Public Property Get HistoryFg() As Boolean
    HistoryFg = mvarHistoryFg
End Property
Public Property Let HistoryFg(ByVal vNewValue As Boolean)
    mvarHistoryFg = vNewValue
End Property

'Key String
Public Property Get KeyString() As String
    KeyString = mvarKeyString
End Property
Public Property Let KeyString(ByVal vNewValue As String)
    mvarKeyString = vNewValue
End Property


Public Sub GetRefValues()
    If SdVal <> 0 Then
        dbl1SD_P = MeanVal + SdVal
        dbl2SD_P = MeanVal + (SdVal * 2)
        dbl3SD_P = MeanVal + (SdVal * 3)
        dbl1SD_N = MeanVal - SdVal
        dbl2SD_N = MeanVal - (SdVal * 2)
        dbl3SD_N = MeanVal - (SdVal * 3)
        mvarMultiFg = True
    Else
        dbl1SD_P = 0
        dbl2SD_P = 0
        dbl3SD_P = 0
        dbl1SD_N = 0
        dbl2SD_N = 0
        dbl3SD_N = 0
        mvarMultiFg = False
    End If
End Sub

'-- Return : -1 이면 정상,
'            0-9 이면 유효숫자 오류 (소수점 이하 N자리),
'            92 이면 유효숫자 오류 (정수형만 입력),
'            93 이면 결과코드 입력오류
'            94 이면 비율결과 입력오류
'            95 이면 Free결과 입력오류
Public Function DoResultCheck(ByVal pValue As String, Optional ByVal pRuleFg As Boolean = True) As Integer
   
    DoResultCheck = 0
   
    '결과가 널이면 초기화.
    If Trim(pValue) = "" Then
        RstCd = "": RaDiv = "": RaName = ""
        VfyDt = "": VfyTm = "": VfyId = ""
        If pRuleFg Then RstText = ""
        If InStr(1, "AN ", RstType) > 0 Then RstType = ""
        Exit Function
    End If
    
    If Not (pValue Like "*+" Or pValue Like "*-") And IsNumeric(pValue) Then RstType = "N"
    
    VfyDt = "": VfyTm = "": VfyId = ""
    
    '결과가 들어왔고 결과유형이 널이 아니면 결과유형에 NUMERIC 혹은 ALPHA 셋팅.
    Select Case RstType
        Case "A"
            DoResultCheck = IsAlpha(pValue, pRuleFg)
        Case "N"
            DoResultCheck = IsValid(pValue, pRuleFg)
        Case "F"
            DoResultCheck = IsFree(pValue, pRuleFg)
        Case Else
            DoResultCheck = IsFree(pValue, pRuleFg)
    End Select

'Modify By Legends 2001/09/19
'        If InStr(1, "AN", RstType) > 0 Then
'            If Not (pValue Like "*+" Or pValue Like "*-") And IsNumeric(pValue) Then
'                DoResultCheck = IsValid(pValue, pRuleFg)
'            Else
'                DoResultCheck = IsAlpha(pValue, pRuleFg)
'            End If
'        ElseIf InStr(1, "F", RstType) > 0 Then
'            DoResultCheck = IsFree(pValue, pRuleFg)
'        End If
End Function

Public Function IsValid(ByVal pValue As String, Optional ByVal pRuleFg As Boolean = True) As Integer
    
    Dim tmpRst As String
    
    tmpRst = medGetP(pValue, 2, ".")
    
    If (AvalVal = 9) Or (Len(tmpRst) = AvalVal) Then
        IsValid = 0
        RstType = "N": RstVal = pValue: RstCd = pValue
        If pRuleFg Then Call IsAccept("N", pValue)   'Numeric 결과 참고치 체크
    Else
        If AvalVal = 0 Then
            IsValid = 92
        Else
            IsValid = AvalVal
        End If
    End If
    
End Function


Public Function IsAlpha(ByVal pValue As String, Optional ByVal pRuleFg As Boolean = True) As Integer
    
    Dim tmpSql As String
    Dim Rs As Recordset
    
    tmpSql = " select field1 from " & T_LAB031 & _
             " where " & _
                               DBW("cdindex=", LC2_ItemResult) & _
                     " and " & DBW("cdval1=", TestCd) & _
                     " and " & DBW("cdval2=", pValue)
             
    Set Rs = New Recordset
    Rs.Open tmpSql, DBConn
    
    If Not Rs.EOF Then
        IsAlpha = 0
        RstCdNm = Trim("" & Rs.Fields("Field1").Value)
        RstType = "A": RstVal = 0: RstCd = pValue
        If pRuleFg Then Call IsAccept("A", pValue)   'Alpha결과 참고치 체크
    Else
        IsAlpha = 93
    End If
    
    Set Rs = Nothing
    
End Function

Public Function IsFree(ByVal pValue As String, Optional ByVal pRuleFg As Boolean = True) As Integer
    Dim tmpSql As String
    Dim Rs As Recordset
    
    tmpSql = " select refcd from " & T_LAB024 & _
             " where " & _
                               DBW("ctrlcd=", mvarCtrlCd) & _
                     " and " & DBW("levelcd=", mvarLevelCd) & _
                     " and " & DBW("lotno=", mvarLotNo) & _
                     " and " & DBW("testcd=", mvarTestCd)

    Set Rs = New Recordset
    Rs.Open tmpSql, DBConn
    
    IsFree = 0
    
    If Rs.EOF Then
        IsFree = 95
    Else
'        If UCase(pValue) = UCase(Trim("" & rs.Fields("refcd").Value)) Then
            RstCdNm = Trim("" & Rs.Fields("refcd").Value)
            RstType = "F": RstVal = 0: RstCd = pValue
            If pRuleFg Then Call IsAccept("F", pValue)   'Alpha결과 참고치 체크
'        Else
'            IsFree = 95
'        End If
    End If
    
'    If Not rs.EOF Then
'        IsFree = 0
'        RstCdNm = Trim("" & rs.Fields("refcd").Value)
'        RstType = "F": RstVal = 0: RstCd = pValue
'        If pRuleFg Then Call IsAccept("F", pValue)   'Alpha결과 참고치 체크
'    Else
'        IsFree = 95
'    End If
    
    Set Rs = Nothing
End Function

Private Sub IsAccept(ByVal pRstTp As String, ByVal pValue As String)
    
    Dim blnCheck As Boolean
    
    Select Case pRstTp
    Case "A":   'Alpha
        If UCase(pValue) = UCase(mvarRefCd) Then
            mvarRaDiv = "A" 'Accept
            mvarRaName = "Accept"
            mvarRaColor = &HDF6A3E
            mvarRejectMsg = ""
        Else
            mvarRaDiv = "R" 'Reject
            mvarRaName = "Reject"
            mvarRaColor = &H7477EF
            mvarRejectMsg = "참고치와 결과치가 다름"
        End If
    Case "N":   'Numeric
        blnCheck = MultiRule(pValue)
        If blnCheck Then
            mvarRaDiv = "A" 'Accept
            mvarRaName = "Accept"
            mvarRaColor = &HDF6A3E
        Else
            mvarRaDiv = "R" 'Reject
            mvarRaName = "Reject"
            mvarRaColor = &H7477EF
        End If
    Case "F"
        If UCase(pValue) = UCase(RstCdNm) Then
            mvarRaDiv = "A" 'Accept
            mvarRaName = "Accept"
            mvarRaColor = &HDF6A3E
            mvarRejectMsg = ""
        Else
            mvarRaDiv = "R" 'Reject
            mvarRaName = "Reject"
            mvarRaColor = &H7477EF
            mvarRejectMsg = "참고치와 결과치가 다름"
        End If
    End Select
    
End Sub

Private Function MultiRule(ByVal pValue As String) As Boolean

    Dim pRstVal As Double
    Dim tmpVal As Double
    Dim iCnt As Integer
    Dim i As Integer

    mvarRejectMsg = ""
    MultiRule = True
    pRstVal = Val(pValue)
    
    'SD값이 설정되어 있지 않는 경우
    If Not mvarMultiFg Then Exit Function
    
'Step1 : Once 3SD
    If pRstVal > dbl3SD_P Or pRstVal < dbl3SD_N Then
        MultiRule = False
        mvarRejectMsg = mvarRejectMsg & "Step1 - Once 3SD, "
        mvarRejectSeq = 1
        'Exit Function
    End If
    
    If mvarHistoryCnt <= 0 Then Exit Function
    
'Step2 : Once 4SD
    If pRstVal > dbl2SD_P And RstHistory(1) < dbl2SD_N Then
        MultiRule = False
        mvarRejectMsg = mvarRejectMsg & "Step2 - Once 4SD,"
        mvarRejectSeq = 2
        'Exit Function
    ElseIf RstHistory(1) > dbl2SD_P And pRstVal < dbl2SD_N Then
        MultiRule = False
        mvarRejectMsg = mvarRejectMsg & "Step2 - Once 4SD, "
        mvarRejectSeq = 2
        'Exit Function
    End If
'Step3 : Twice 2SD
    If pRstVal > dbl2SD_P And RstHistory(1) > dbl2SD_P Then
        MultiRule = False
        mvarRejectMsg = mvarRejectMsg & "Step3 - Twice 2SD, "
        mvarRejectSeq = 3
        'Exit Function
    ElseIf pRstVal < dbl2SD_N And RstHistory(1) < dbl2SD_N Then
        MultiRule = False
        mvarRejectMsg = mvarRejectMsg & "Step3 - Twice 2SD, "
        mvarRejectSeq = 3
        'Exit Function
    End If
    
    '** 과거 결과가 CS_TotCnt회 만큼 충분하지 못한 경우..
    'If Not HistoryFg Then Exit Function
    If mvarHistoryCnt + 1 < CS_TotCnt Then Exit Function
        
'Step4 : 4 Times 1SD
    iCnt = 0
    If pRstVal > dbl1SD_P Then     '(+)1SD
        For i = 1 To CS_ChkCnt - 1
            If RstHistory(i) <= dbl1SD_P Then Exit For
            iCnt = iCnt + 1
        Next
    ElseIf pRstVal < dbl1SD_N Then '(-)1SD
        For i = 1 To CS_ChkCnt - 1
            If RstHistory(i) >= dbl1SD_N Then Exit For
            iCnt = iCnt + 1
        Next
    End If
    'CS_TotCnt개의 결과중 CS_ChkCnt번 (+,-)1SD를 연속해서 벗어남.
    If iCnt = CS_ChkCnt Then
        MultiRule = False
        mvarRejectMsg = mvarRejectMsg & "Step4 - 4 / 1 SD, "
        mvarRejectSeq = 4
        'Exit Function
    End If

    'Step4 : 6/10 Times 1SD
    '************************************************************************
'    iCnt1 = 0: iCnt2 = 0
'    If pRstVal > mvarMeanVal And pRstVal < dbl1SD_P Then     '(+)1SD
'        iCnt1 = iCnt1 + 1
'    ElseIf pRstVal < mvarMeanVal And pRstVal > dbl1SD_N Then '(-)1SD
'        iCnt2 = iCnt2 + 1
'    End If
'    For i = 1 To CS_TotCnt
'        If RstHistory(i) > mvarMeanVal And RstHistory(i) < dbl1SD_P Then
'            iCnt1 = iCnt1 + 1
'        ElseIf RstHistory(i) < mvarMeanVal And RstHistory(i) > dbl1SD_N Then
'            iCnt2 = iCnt2 + 1
'        End If
'    Next
'    'CS_TotCnt개의 결과중 CS_ChkCnt번이상 (+,-)1SD 내에 머무름.
'    If iCnt1 >= CS_ChkCnt Or iCnt2 >= CS_ChkCnt Then
'        MultiRule = False
'        mvarRejectMsg = mvarRejectMsg & "Step4 - 6/10 Times 1SD, "
'        mvarRejectSeq = 4
'        'Exit Function
'    End If
    '************************************************************************

'Step5 : 10 Times Trend
    iCnt = 0
    tmpVal = pRstVal
    If tmpVal > RstHistory(1) Then  '** (+)방향 Trend
        For i = 1 To CS_TotCnt
            If tmpVal < RstHistory(i) Then Exit For
            iCnt = iCnt + 1
            tmpVal = RstHistory(i)
        Next
    Else        '** (-)방향 Trend
        For i = 1 To CS_TotCnt
            If tmpVal > RstHistory(i) Then Exit For
            iCnt = iCnt + 1
            tmpVal = RstHistory(i)
        Next
    End If
    If iCnt = 10 Then
        MultiRule = False
        mvarRejectMsg = mvarRejectMsg & "Step5 - 10 Times Trend"
        mvarRejectSeq = 5
        'Exit Function
    End If

End Function


Private Sub Class_Initialize()
    mvarHistoryCnt = 0
    mvarHistoryFg = False
    Erase mvarRstHistory
    Set mHisCol = New Collection
End Sub


Public Function UpdateResult() As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim sSql As String
    
    sSql = " Update " & T_LAB026 & " Set " & _
           DBW("rstcd=", mvarRstCd, 1) & _
           DBW("rsttype=", mvarRstType, 1) & _
           DBW("radiv=", mvarRaDiv, 1) & _
           DBW("vfydt=", mvarVfyDt, 1) & _
           DBW("vfytm=", mvarVfyTm, 1) & _
           DBW("vfyid=", Val(mvarVfyId), 1) & _
           DBW("eqpcd=", mvarEqpCd, 1) & _
           DBW("method=", mvarMethod, 1) & _
           DBW("txtfg =", mvarTxtFg) & _
           " Where " & DBW("workarea=", mvarWorkArea) & _
           " and   " & DBW("accdt=", mvarAccDt) & _
           " and   " & DBW("accseq=", mvarAccSeq) & _
           " and   " & DBW("testcd=", mvarTestCd)
         
    DBConn.Execute sSql
    
    Dim Ret As Boolean
    '소견결과
    If mvarTxtFg = "1" Then
        Ret = SaveComments
        If Not Ret Then GoTo ErrorHandler
    End If
    UpdateResult = True
    
    Exit Function
    
ErrorHandler:

    UpdateResult = False
End Function

Private Function SaveComments() As Boolean
    
    Dim SqlStmt As String
    Dim Rs As Recordset
    
    On Error GoTo Err_Trap
    
    SaveComments = True
    
    SqlStmt = " select * from " & T_LAB028 & " " & _
              " where " & DBW("workarea=", mvarWorkArea) & _
              " and " & DBW("accdt=", mvarAccDt) & _
              " and " & DBW("accseq=", mvarAccSeq) & _
              " and " & DBW("testcd=", mvarTestCd) & _
              " order by seq desc"
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
        SqlStmt = "insert into " & T_LAB028 & " values " & _
                  "(" & DBV("workarea", mvarWorkArea, 1) & DBV("accdt", mvarAccDt, 1) & DBV("accseq", mvarAccSeq, 1) & _
                        DBV("testcd", mvarTestCd, 1) & DBV("seq", "0", 1) & DBV("text", mvarRstText) & ")"
    Else
        If TxtFg = "0" Then
            SqlStmt = " delete from " & T_LAB028 & " " & _
                      " where " & DBW("workarea=", mvarWorkArea) & _
                      " and " & DBW("accdt=", mvarAccDt) & _
                      " and " & DBW("accseq=", mvarAccSeq) & _
                      " and " & DBW("testcd=", mvarTestCd) & _
                      " and " & DBW("seq=", Trim("" & Rs.Fields("seq").Value))
        Else
            SqlStmt = " update " & T_LAB028 & " set " & _
                                                       DBW("text=", mvarRstText) & _
                      " where " & DBW("workarea=", mvarWorkArea) & _
                      " and " & DBW("accdt=", mvarAccDt) & _
                      " and " & DBW("accseq=", mvarAccSeq) & _
                      " and " & DBW("testcd=", mvarTestCd) & _
                      " and " & DBW("seq=", Trim("" & Rs.Fields("seq").Value))
        End If
    End If
    
    Set Rs = Nothing
    
    DBConn.Execute SqlStmt
    
    Exit Function
    
Err_Trap:
    SaveComments = False
    
    
End Function


Public Sub GetRstTxt()
    Dim strSql As String
    Dim ii As Integer
    Dim Rs As Recordset
   '
    'Text결과
    strSql = "SELECT seq, text FROM " & T_LAB028 & " " & _
            " WHERE " & DBW("workarea =", mvarWorkArea) & _
            " AND " & DBW("accdt =", mvarAccDt) & _
            " AND " & DBW("accseq = ", mvarAccSeq) & _
            " AND " & DBW("testcd = ", mvarTestCd) & _
            " ORDER BY seq desc"
    '
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    If Not Rs.EOF Then
        mvarRstText = Rs.Fields("text").Value & ""
    Else
        mvarRstText = ""
    End If
    
    Set Rs = Nothing
   '
End Sub



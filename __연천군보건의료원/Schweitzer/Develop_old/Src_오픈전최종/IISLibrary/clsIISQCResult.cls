VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISQCResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISQCResult.cls (자사 LIS와 조인할때)
'   작성자  : 이상대
'   내  용  : QC 결과내역 클래스
'   작성일  : 2004-05-21
'   버  전  :
'       1. 1.1.9: 이상대(2005-07-06)
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISQcSql

Private mIntNm      As clsIISIntNm      '장비별 검사항목 클래스
Private mQCMaster   As clsIISQCMaster   'QCMaster 클래스

Private mTestCd     As String   '검사코드
Private mTestNm10   As String   '검사명(10자리)
Private mTestNm     As String   '검사명(전체)
Private mIntResult  As String   '장비에서 전송한 결과
Private mResult     As String   '장비에서 전송한 결과 (IIS 수정한 결과)
Private mRstVal     As Single   '결과1(Numeric)
Private mRstCd      As String   '결과2(Alphabetic)
Private mUnit       As String   '결과단위
Private mRstDiv     As String   'Alternative, Require 유형(A:Alternative, R:Require)
Private mRstType    As String   '결과유형 (F:Free, N:Numeric, A:Alpha)
Private mRADiv      As String   '결과판정 (A:Accept, R:Reject)
Private mTxtFg      As String   'Text결과 유무(Null:무, 1:유)
Private mText       As String   'Text결과
Private mAutoFg     As String   '장비전송 유무(Null:무, 1:유)
Private mEqpCd      As String   '장비코드
Private mCtrlCd     As String   'Control Code
Private mLevelCd    As String   'Level Code
Private mLotNo      As String   'Lot No

Private mHisResult(8) As String '컨트롤의 과거결과 9개
Private mHisCnt     As Long     '컨트롤의 조회된 과거결과의 개수
Private mHisQuery   As Boolean  '컨트롤의 과거결과 조회유무

Private m1SDP       As Single   '1SD(+)
Private m1SDN       As Single   '1SD(-)
Private m2SDP       As Single   '2SD(+)
Private m2SDN       As Single   '2SD(-)
Private m3SDP       As Single   '3SD(+)
Private m3SDN       As Single   '3SD(-)

Private mSndChk     As Boolean  '오더전송 여부 (BN100때문에 만듬)

Public Property Get IntNm() As clsIISIntNm
    Set IntNm = mIntNm
End Property

Public Property Let IntNm(ByVal vData As clsIISIntNm)
    Set mIntNm = vData
End Property

Public Property Get QCMaster() As clsIISQCMaster
    Set QCMaster = mQCMaster
End Property

Public Property Let QCMaster(ByVal vData As clsIISQCMaster)
    Set mQCMaster = vData
End Property

Public Property Get TestCd() As String
    TestCd = mTestCd
End Property

Public Property Let TestCd(ByVal vData As String)
    mTestCd = vData
End Property

Public Property Get TestNm10() As String
    TestNm10 = mTestNm10
End Property

Public Property Let TestNm10(ByVal vData As String)
    mTestNm10 = vData
End Property

Public Property Get TestNm() As String
    TestNm = mTestNm
End Property

Public Property Let TestNm(ByVal vData As String)
    mTestNm = vData
End Property

Public Property Get IntResult() As String
    IntResult = mIntResult
End Property

Public Property Let IntResult(ByVal vData As String)
    mIntResult = vData
End Property

Public Property Get Result() As String
    Result = mResult
End Property

Public Property Let Result(ByVal vData As String)
    mResult = vData
End Property

Public Property Get RstVal() As Single
    RstVal = mRstVal
End Property

Public Property Let RstVal(ByVal vData As Single)
    mRstVal = vData
End Property

Public Property Get RstCd() As String
    RstCd = mRstCd
End Property

Public Property Let RstCd(ByVal vData As String)
    mRstCd = vData
End Property

Public Property Get Unit() As String
    Unit = mUnit
End Property

Public Property Let Unit(ByVal vData As String)
    mUnit = vData
End Property

Public Property Get RstDiv() As String

End Property

Public Property Let RstDiv(ByVal vData As String)

End Property

Public Property Get RstType() As String
    RstType = mRstType
End Property

Public Property Let RstType(ByVal vData As String)
    mRstType = vData
End Property

Public Property Get RADiv() As String
    RADiv = mRADiv
End Property

Public Property Let RADiv(ByVal vData As String)
    mRADiv = vData
End Property

Public Property Get TxtFg() As String
    TxtFg = mTxtFg
End Property

Public Property Let TxtFg(ByVal vData As String)
    mTxtFg = vData
End Property

Public Property Get Text() As String
    Text = mText
End Property

Public Property Let Text(ByVal vData As String)
    mText = vData
End Property

Public Property Get AutoFg() As String
    AutoFg = mAutoFg
End Property

Public Property Let AutoFg(ByVal vData As String)
    mAutoFg = vData
End Property

Public Property Get EqpCd() As String
    EqpCd = mEqpCd
End Property

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Get CtrlCd() As String
    CtrlCd = mCtrlCd
End Property

Public Property Let CtrlCd(ByVal vData As String)
    mCtrlCd = vData
End Property

Public Property Get LevelCd() As String
    LevelCd = mLevelCd
End Property

Public Property Let LevelCd(ByVal vData As String)
    mLevelCd = vData
End Property

Public Property Get LotNo() As String
    LotNo = mLotNo
End Property

Public Property Let LotNo(ByVal vData As String)
    mLotNo = vData
End Property

Public Property Get SndChk() As Boolean
    SndChk = mSndChk
End Property

Public Property Let SndChk(ByVal vData As Boolean)
    mSndChk = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 접수번호를 이용해 QC결과내역 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'   반환 : Recordset
'-----------------------------------------------------------------------------'
Public Function GetQCResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As ADODB.Recordset
    Dim Rs As ADODB.Recordset

On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectQCResultInfo(pWorkarea, pAccDt, pAccSeq), , adCmdText)
    Set GetQCResultInfo = Rs
    Set Rs = Nothing
    Exit Function

Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISQCResult", "GetQCResultInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스 QC 결과입력시 생성되는 쿼리 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pEqpCd    : 장비코드
'-----------------------------------------------------------------------------'
Public Function GetQCRegisterSql(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pEqpCd As String) As String
    Dim SQL As String
    
    '## 1.1.9: 이상대(2005-07-06)
    '   - QC결과내역의 autofg, eqpcd가 Null 업데이트 되는거 수정
    If mRstType = "N" Then      '## 결과유형이 Numerice 일때
        SQL = mSQL.UpdateQCResultNumeric(pWorkarea, pAccDt, pAccSeq, mTestCd, mRstVal, mRstCd, _
            mRstType, mRADiv, mTxtFg, "1", pEqpCd)
    ElseIf mRstType = "F" Then  '## 결과유형이 Free 일때
        SQL = mSQL.UpdateQCResultFree(pWorkarea, pAccDt, pAccSeq, mTestCd, mRstCd, mRstType, _
            mRADiv, mTxtFg, "1", pEqpCd)
    End If
    
    GetQCRegisterSql = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스 QC 소견결과 입력쿼리 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'-----------------------------------------------------------------------------'
Public Function GetQCCommentSql(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    Dim Rs      As ADODB.Recordset
    Dim lngSeq  As Long
    
    Set Rs = DbCon.Execute(mSQL.SelectQCComment(pWorkarea, pAccDt, pAccSeq, mTestCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then
        GetQCCommentSql = mSQL.InsertQCComment(pWorkarea, pAccDt, pAccSeq, mTestCd, 0, mText)
    Else
        lngSeq = CLng(Rs.Fields("SEQ").Value)
        GetQCCommentSql = mSQL.UpdateQCComment(pWorkarea, pAccDt, pAccSeq, mTestCd, lngSeq, mText)
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : QC검사결과 체크
'-----------------------------------------------------------------------------'
Public Sub CheckQCResult()
    '## 결과유형 체크
    '## - 수치결과: 유효숫자 체크, 멀티룰 체크
    '## - 문자결과: Alpha결과 참고치 체크
    If Not (mResult Like "*+" Or mResult Like "*-") And IsNumeric(mResult) Then
        If mQCMaster.RefCd = "" Then
            '## 유효숫자, 멀티룰 체크
            Call CheckAvalValue
            Call CheckMultiRule
            mRstType = "N"
        Else
            '## Alpha결과 참고치와 비교
            Call CheckAlphaValue
            mRstType = "F"
            mRstCd = mResult
        End If
    Else
        '## Alpha결과 참고치와 비교
        Call CheckAlphaValue
        mRstType = "F"
        mRstCd = mResult
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : QC결과의 Alpha결과 참고치와 비교
'-----------------------------------------------------------------------------'
Public Sub CheckAlphaValue()
    If QCMaster.RefCd = mResult Then
        mRADiv = "A"
    Else
        mRADiv = "R"
        mTxtFg = "1"
        mText = "참고치와 결과치가 서로 다름"
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : QC결과의 유효숫자 체크
'-----------------------------------------------------------------------------'
Public Sub CheckAvalValue()
    Dim strFormat  As String     '소수점 자리수의 포맷
    Dim strResult  As String     '임시결과
    Dim lngAvalVal As Long       '유효숫자
    
    '## 유효숫자 체크
    lngAvalVal = mQCMaster.AvalVal
    If lngAvalVal > 0 And lngAvalVal < 9 Then
        strFormat = "0." & String$(lngAvalVal, "0")
        strResult = Format$(mResult, strFormat)
    ElseIf lngAvalVal = "0" Then
        strResult = Round(mResult)
    Else
        strResult = mResult
    End If
    
    mRstVal = CSng(strResult)
    mRstCd = strResult
End Sub

'-----------------------------------------------------------------------------'
'   기능 : QC결과의 MultiRule 체크
'-----------------------------------------------------------------------------'
Public Sub CheckMultiRule()
    Dim strWMSet        As String   'Multi-Rule Setting
    Dim lng4TimeCnt     As Long     '4Times 1SD를 체크변수
    Dim lng10TimeCnt    As Long     '10Times Trend 체크변수
    Dim i               As Long
    
    If mQCMaster.SDVal = 0 Then Exit Sub
    
    '## MultiRule 체크에 필요한 정보 조회, 과거결과 9개 조회
    Call GetReference
    If mHisQuery = False Then Call GetHistoryResult
    strWMSet = mQCMaster.WMSet
    
    '## Step1 - Once 3SD
    If Mid$(strWMSet, 1, 1) = "1" Then
        If mRstVal > m3SDP Or mRstVal < m3SDN Then
            mText = "Step1 - Once 3SD"
            mRADiv = "R"
            mTxtFg = "1"
            Exit Sub
        End If
    End If
    
    '## Step2 - Once 4SD
    If Mid$(strWMSet, 2, 1) = "1" Then
        If mHisCnt > 0 Then
            If mRstVal > m2SDP And mHisResult(0) < m2SDN Then
                mText = "Step2 - Once 4SD"
                mRADiv = "R"
                mTxtFg = "1"
                Exit Sub
            ElseIf mRstVal < m2SDN And mHisResult(0) > m2SDP Then
                mText = "Step2 - Once 4SD"
                mRADiv = "R"
                mTxtFg = "1"
                Exit Sub
            End If
        End If
    End If
    
    '## Step3 - Twice 2SD
    If Mid$(strWMSet, 3, 1) = "1" Then
        If mHisCnt > 0 Then
            If mRstVal > m2SDP And mHisResult(0) > m2SDP Then
                mText = "Step3 - Twice 2SD"
                mRADiv = "R"
                mTxtFg = "1"
                Exit Sub
            ElseIf mRstVal < m2SDN And mHisResult(0) < m2SDN Then
                mText = "Step3 - Twice 2SD"
                mRADiv = "R"
                mTxtFg = "1"
                Exit Sub
            End If
        End If
    End If
    
    '## Step4 - 4 Times 1SD
    If Mid$(strWMSet, 4, 1) = "1" Then
        If mRstVal > m1SDP And mHisCnt > 3 Then
            lng4TimeCnt = 1
            For i = 0 To 2
                If mHisResult(i) > m1SDP Then
                    lng4TimeCnt = lng4TimeCnt + 1
                End If
            Next i
        ElseIf mRstVal < m1SDN And mHisCnt > 3 Then
            lng4TimeCnt = 1
            For i = 0 To 2
                If mHisResult(i) < m1SDN Then
                    lng4TimeCnt = lng4TimeCnt + 1
                End If
            Next i
        End If
        
        If lng4TimeCnt = 4 Then
            mText = "Step4 - 4 Times 1SD"
            mRADiv = "R"
            mTxtFg = "1"
            Exit Sub
        End If
    End If
    
    '## Step5 - 10 Times Trend
    If Mid$(strWMSet, 5, 1) = "1" Then
        If mRstVal > mQCMaster.MeanVal And mHisCnt >= 9 Then
            lng10TimeCnt = 1
            For i = 0 To 8
                If mHisResult(i) > mQCMaster.MeanVal Then
                    lng10TimeCnt = lng10TimeCnt + 1
                End If
            Next i
        ElseIf mRstVal < mQCMaster.MeanVal And mHisCnt >= 9 Then
            lng10TimeCnt = 1
            For i = 0 To 8
                If mHisResult(i) < mQCMaster.MeanVal Then
                    lng10TimeCnt = lng10TimeCnt + 1
                End If
            Next i
        End If
        
        If lng10TimeCnt = 10 Then
            mText = "Step5 - 10 Times Trend"
            mRADiv = "R"
            mTxtFg = "1"
            Exit Sub
        End If
    End If
    
    mRADiv = "A"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당컨트롤+LotNo+검사코드의 과거결과 9개를 조회
'-----------------------------------------------------------------------------'
Public Sub GetHistoryResult()
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    mHisQuery = True
    Set Rs = DbCon.Execute(mSQL.SelectHistoryResult(mCtrlCd, mLevelCd, mLotNo, mTestCd), , adCmdText)
    If Not (Rs.BOF Or Rs.EOF) Then
        Do Until (Rs.EOF Or mHisCnt >= 9)
            mHisResult(mHisCnt) = Rs.Fields("RSTCD").Value & ""
            mHisCnt = mHisCnt + 1
            Rs.MoveNext
        Loop
    Else
        mHisCnt = 0
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Sub
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISQCResult", "GetHistoryResult", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : MultiRule 체크에 필요한 정보조회
'-----------------------------------------------------------------------------'
Public Sub GetReference()
    Dim sngMeanVal  As Single   '평균값
    Dim sngSDVal    As Single   '표준편차
    
    With mQCMaster
        sngMeanVal = .MeanVal
        sngSDVal = .SDVal
    End With
    
    m1SDP = sngMeanVal + sngSDVal
    m2SDP = sngMeanVal + (sngSDVal * 2)
    m3SDP = sngMeanVal + (sngSDVal * 3)
    m1SDN = sngMeanVal - sngSDVal
    m2SDN = sngMeanVal - (sngSDVal * 2)
    m3SDN = sngMeanVal - (sngSDVal * 3)
End Sub

Private Sub Class_Initialize()
    Set mSQL = New clsIISQcSql
    Set mIntNm = New clsIISIntNm
    Set mQCMaster = New clsIISQCMaster
End Sub

Private Sub Class_Terminate()
    Set mQCMaster = Nothing
    Set mIntNm = Nothing
    Set mSQL = Nothing
End Sub


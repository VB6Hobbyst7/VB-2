VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------------------------'
'   파일명  : clsIISResult.cls (자사 LIS와 조인할때)
'   작성자  : 이상대
'   내  용  : 결과내역 클래스
'   작성일  : 2004-03-31
'   버  전  :
'       1. 1.1.4: 이상대(2005-04-11)
'       2. 1.1.5: 이상대(2005-04-12)
'       3. 1.1.11: 이상대(2005-07-22)
'       4. 1.1.12: 이상대(2005-07-28)
'       5. 1.1.14: 이상대(2005-10-11)
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISLibrarySql

Private mIntNm      As clsIISIntNm      '장비별 검사항목 클래스
Private mTSpc       As clsIISTSpc       '지정검체 클래스
Private mRef        As clsIISRef        '참고치 클래스

Private mTestCd     As String   '검사코드
Private mTestNm10   As String   '검사명(10자리)
Private mTestNm     As String   '검사명(전체)
Private mIntResult  As String   '장비에서 전송한 결과
Private mResult     As String   '장비에서 전송한 결과 (IIS에서 수정한결과)
Private mRstVal     As Single   '결과1(Numeric)
Private mRstCd      As String   '결과2(Alphabetic)
Private mUnit       As String   '결과단위
Private mHLDiv      As String   'H/L
Private mDPDiv      As String   'D/P
Private mLastRst    As String   '최근결과
Private mRerunFg    As String   '재검여부 (0:무, 1:유)
Private mMfyFg      As String   '수정여부 (0:무, 1:유)
Private mTxtFg      As String   '텍스트결과 (0:무, 1:유)
Private mAutoFg     As String   '장비전송여부 (0:무, 1:유)
Private mRstType    As String   '결과유형 (F:Free, N:Numeric, A:Alpha)
Private mRstDiv     As String   'Require/Alternative 여부 (R:Require, A:Alternative, *:상세모코드)
Private mDetailFg   As String   '상세항목순번
Private mEqpCd      As String   '장비코드
Private mAutoVfy    As Boolean  '자동결과등록 가능여부
Private mIntInfo    As String   '장비에서 전송된 추가정보
Private mSndChk     As Boolean  '오더전송 여부 (BN100때문에 만듬)

Public Property Get IntNm() As clsIISIntNm
    Set IntNm = mIntNm
End Property

Public Property Let IntNm(ByVal vData As clsIISIntNm)
    Set mIntNm = vData
End Property

Public Property Get TSpc() As clsIISTSpc
    Set TSpc = mTSpc
End Property

Public Property Set TSpc(ByVal vData As clsIISTSpc)
    Set mTSpc = vData
End Property

Public Property Get Ref() As clsIISRef
    Set Ref = mRef
End Property

Public Property Set Ref(ByVal vData As clsIISRef)
    Set mRef = vData
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

'## 1.1.4: 이상대(2005-04-11)
'   - mRstVal의 자료형을 Long에서 Single로 수정
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

Public Property Get HLDiv() As String
    HLDiv = mHLDiv
End Property

Public Property Let HLDiv(ByVal vData As String)
    mHLDiv = vData
End Property

Public Property Get DPDiv() As String
    DPDiv = mDPDiv
End Property

Public Property Let DPDiv(ByVal vData As String)
    mDPDiv = vData
End Property

Public Property Get LastRst() As String
    LastRst = mLastRst
End Property

Public Property Let LastRst(ByVal vData As String)
    mLastRst = vData
End Property

Public Property Get RerunFg() As String
    RerunFg = mRerunFg
End Property

Public Property Let RerunFg(ByVal vData As String)
    mRerunFg = vData
End Property

Public Property Get MfyFg() As String
    MfyFg = mMfyFg
End Property

Public Property Let MfyFg(ByVal vData As String)
    mMfyFg = vData
End Property

Public Property Get TxtFg() As String
    TxtFg = mTxtFg
End Property

Public Property Let TxtFg(ByVal vData As String)
    mTxtFg = vData
End Property

Public Property Get AutoFg() As String
    AutoFg = mAutoFg
End Property

Public Property Let AutoFg(ByVal vData As String)
    mAutoFg = vData
End Property

Public Property Get RstType() As String
    RstType = mRstType
End Property

Public Property Let RstType(ByVal vData As String)
    mRstType = vData
End Property

Public Property Get RstDiv() As String
    RstDiv = mRstDiv
End Property

Public Property Let RstDiv(ByVal vData As String)
    mRstDiv = vData
End Property

Public Property Get DetailFg() As String
    DetailFg = mDetailFg
End Property

Public Property Let DetailFg(ByVal vData As String)
    mDetailFg = vData
End Property

Public Property Get EqpCd() As String
    EqpCd = mEqpCd
End Property

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Get AutoVfy() As Boolean
    AutoVfy = mAutoVfy
End Property

Public Property Let AutoVfy(ByVal vData As Boolean)
    mAutoVfy = vData
End Property

Public Property Get IntInfo() As String
    IntInfo = mIntInfo
End Property

Public Property Let IntInfo(ByVal vData As String)
    mIntInfo = vData
End Property

Public Property Get SndChk() As Boolean
    SndChk = mSndChk
End Property

Public Property Let SndChk(ByVal vData As Boolean)
    mSndChk = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 접수번호를 이용해 결과내역 조회
'   인수 :
'       - pWorkarea  : Workarea
'       - pAccDt     : 접수일자
'       - pAccSeq    : 접수순번
'       - pAutoRerun : 자동재검 사용유무
'   반환 : Recordset
'-----------------------------------------------------------------------------'
Public Function GetResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pAutoRerun As AutoRerunEnum) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
    '## 1.1.12: 이상대(2005-07-28)
    '   - 자동재검 사용유무 인수추가
    '   - 자동재검을 사용할때는 결과가 들어있어도 보고되지 않은 항목을 조회
On Error GoTo Errors:
    If pAutoRerun = ccAutoRerun Then
        Set Rs = DbCon.Execute(mSQL.SelectAutoRerunResultInfo(pWorkarea, pAccDt, pAccSeq), , adCmdText)
    Else
        Set Rs = DbCon.Execute(mSQL.SelectResultInfo(pWorkarea, pAccDt, pAccSeq), , adCmdText)
    End If
    Set GetResultInfo = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISResult", "GetResultInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수번호를 이용해 특수검사 결과내역 조회 - 1.1.14: 이상대(2005-10-11)
'          특수검사 결과클래스를 생성해야하는게 원칙이지만 결과를 저장할 필요가
'          없어 일반검사 결과클래스에 생성함
'   인수 :
'       - pWorkarea  : Workarea
'       - pAccDt     : 접수일자
'       - pAccSeq    : 접수순번
'   반환 : Recordset
'-----------------------------------------------------------------------------'
Public Function GetSpecialResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
'    Set Rs = DbCon.Execute(mSQL.SelectSpecialResultInfo(pWorkarea, pAccDt, pAccSeq), , adCmdText)
    Set Rs = DbCon.Execute(mSQL.SelectResultInfo(pWorkarea, pAccDt, pAccSeq), , adCmdText)
    Set GetSpecialResultInfo = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISResult", "GetSpecialResultInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사결과 체크
'-----------------------------------------------------------------------------'
Public Sub CheckResult()
    If mResult = "" Then Exit Sub
    
    '## 결과유형 체크: 결과유형이 N, A, Null일경우에만 결과값을 이용해 체크
    If mRstType = "A" Or mRstType = "N" Or mRstType = "" Then
        If Not (mResult Like "*+" Or mResult Like "*-") And IsNumeric(mResult) Then
            mRstType = "N"
        Else
            mRstType = "F"
        End If
    End If
    
    '## 유효숫자 체크
    Call CheckAvalValue
    
    '## 1.1.5: 이상대(2005-04-12)
    '   - 결과유형이 Numeric일때는 Delta, Panic, High, Low 체크
    '   - 결과유형이 Free일때는 Alpha결과 참고치 체크 하도록 수정
    '## 결과가 Null이 아닐때만 H/P, D/P체크
    If mResult <> "" Then
        If mRstType = "N" Then
            Call CheckDeltaPanic
            Call CheckHighLow
        ElseIf mRstType = "F" Then
            Call CheckAlphaValue
        End If
    End If
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 검사결과의 유효숫자 체크
'-----------------------------------------------------------------------------'
Public Sub CheckAvalValue()
    Dim strFormat  As String     '소수점 자리수의 포맷
    Dim strResult  As String     '임시결과
    Dim lngAvalVal As Long       '유효숫자

On Error Resume Next
    
    '## 유효숫자 체크
    If mRstType = "N" Then
        lngAvalVal = mTSpc.AvalVal
        If lngAvalVal > 0 And lngAvalVal < 9 Then
            strFormat = "0." & String$(lngAvalVal, "0")
            strResult = Format$(mResult, strFormat)
        ElseIf lngAvalVal = "0" Then
            strResult = Round(mResult)
        Else
            strResult = mResult
        End If
    End If
    
    '## 결과유형에 따른 결과값 입력
    Select Case mRstType
        Case "N"    '## Numeric
            mRstVal = CSng(strResult)
            mRstCd = strResult
        Case "F"    '## Free
            mRstCd = mResult
    End Select
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 검사결과의 Delta/Panic 체크
'   반환 :
'       - D : Delta
'       - P : Panic
'       - DP : Delta & Panic
'-----------------------------------------------------------------------------'
Public Function CheckDeltaPanic() As String
    Dim strPanic As String  '임시 Panic
    Dim strDelta As String  '임시 Delta
    Dim lngTemp  As String

On Error Resume Next
    
    If IsNumeric(mResult) = False Then Exit Function
    
    '## Panic 체크
    If IISPANICCHECK = "IIS102" Then
        '## 지정검체 마스터에서 Panic 체크
        If mTSpc.PanicFg = "1" Then
            If Val(mResult) < mTSpc.PanicFrVal Or Val(mResult) > mTSpc.PanicToVal Then
                strPanic = "P"
            End If
        End If
    ElseIf IISPANICCHECK = "IIS103" Then
        '## 참고치 마스터에서 Panic 체크
        If mRef.PanicFg = "1" Then
            If Val(mResult) < mRef.PanicFrVal Or Val(mResult) > mRef.PanicToVal Then
                strPanic = "P"
            End If
        End If
    End If
    
    '## Delta 체크
    If mTSpc.DeltaFg = "1" And mLastRst <> "" And IsNumeric(mLastRst) Then
        '## 1.1.11: 이상대(2005-07-22)
        '   - 과거결과가 "0"일때의 처리 추가
        If Val(mLastRst) = 0 Then
            lngTemp = 0
        Else
            lngTemp = ((Val(mLastRst) - Val(mResult)) / Val(mLastRst)) * 100
        End If
        
        If lngTemp < mTSpc.DeltaFrVal * (-1) Or lngTemp > TSpc.DeltaToVal Then
            strDelta = "D"
        End If
    End If
    
    mDPDiv = strPanic & strDelta
    CheckDeltaPanic = mDPDiv
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사결과의 High/Low 체크
'   반환 :
'       - L  : Low
'       - H  : High
'       - LH : Low & High
'-----------------------------------------------------------------------------'
Public Function CheckHighLow() As String
    If IsNumeric(mResult) = False Then Exit Function
    
    If mRef.RefFg = "1" Then
        If Val(mResult) < mRef.RefFrVal Then
            mHLDiv = "L"
        ElseIf Val(mResult) > mRef.RefToVal Then
            mHLDiv = "H"
        End If
    End If
    
    CheckHighLow = mHLDiv
End Function

'-----------------------------------------------------------------------------'
'   기능 : Alpha결과 참고치 체크 - 1.1.5: 이상대(2005-04-12)
'-----------------------------------------------------------------------------'
Private Sub CheckAlphaValue()
    Dim strDecision As String
    
    strDecision = mRef.GetAlpahRef(mTestCd, mRstCd)
    Select Case strDecision
        Case "H":   mHLDiv = "H"    'High
        Case "L":   mHLDiv = "L"    'Low
        Case "P":   mDPDiv = "P"    'Panic
        Case "*":   mDPDiv = "N"    'Abnormal
    End Select
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스 결과입력시 생성되는 쿼리 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pEqpCd    : 장비코드
'   반환 : 결과내역을 업데이트 하는쿼리
'-----------------------------------------------------------------------------'
Public Function GetIISRegisterSql(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pEqpCd As String) As String
    Dim SQL As String
    
    '## 상세모코드 일때: AUTOFG, EQPCD
    '## 단일항목 일때: RSTVAL, RSTCD, HLDIV, DPDIV, AUTOFG, RSTTYPE, EQPCD,
    If mRstDiv = "*" Then
        SQL = mSQL.UpdateIISResultDetail(pWorkarea, pAccDt, pAccSeq, mTestCd, "1", pEqpCd)
    Else
        If mRstType = "N" Then
            SQL = mSQL.UpdateIISResultNumeric(pWorkarea, pAccDt, pAccSeq, mTestCd, mRstVal, mRstCd, mHLDiv, _
                    mDPDiv, "1", mRstType, pEqpCd)
        Else
            SQL = mSQL.UpdateIISResultFree(pWorkarea, pAccDt, pAccSeq, mTestCd, mRstCd, mHLDiv, _
                    mDPDiv, "1", mRstType, pEqpCd)
        End If
    End If
    
    GetIISRegisterSql = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스 자동결과등록시 생성되는 쿼리 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pEqpCd    : 장비코드
'       - pVfyDt    : 결과확인일자
'       - pVfyTm    : 결과확인시간
'   반환 : 결과내역을 업데이트 하는쿼리
'-----------------------------------------------------------------------------'
Public Function GetIISVfySql(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pEqpCd As String, ByVal pVfyDt As String, _
                    ByVal pVfyTm As String) As String
    Dim SQL      As String
    
    If mRstDiv = "*" Then
        SQL = mSQL.UpdateIISVfyDetail(pWorkarea, pAccDt, pAccSeq, mTestCd, "1", pEqpCd, pVfyDt, _
                pVfyTm, EmpId)
    Else
        SQL = mSQL.UpdateIISVfy(pWorkarea, pAccDt, pAccSeq, mTestCd, mRstVal, mRstCd, mHLDiv, _
                mDPDiv, "1", mRstType, pEqpCd, pVfyDt, pVfyTm, EmpId)
    End If
    
    GetIISVfySql = SQL
End Function

Private Sub Class_Initialize()
    Set mIntNm = New clsIISIntNm
    Set mTSpc = New clsIISTSpc
    Set mRef = New clsIISRef
    Set mSQL = New clsIISLibrarySql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
    Set mRef = Nothing
    Set mTSpc = Nothing
    Set mIntNm = Nothing
End Sub

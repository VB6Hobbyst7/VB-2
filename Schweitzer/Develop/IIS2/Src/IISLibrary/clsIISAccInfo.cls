VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISAccInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------------------------'
'   파일명  : clsIISAccInfo.cls (자사 LIS와 조인할때)
'   작성자  : 이상대
'   내  용  : 접수내역 클래스
'   작성일  : 2004-03-31
'   버  전  :
'       1. 1.0.1: 이상대(2004-11-17)
'          - mAge, mColDt, mColTm, mColId, mRcvDt, mRcvTm, mRcvId 필드추가
'          - 접수내역 조회시 추가된 필드를 같이 조회하도록 수정
'          - 미생물 결과내역 조회시 검체명을 따로 조회부분 삭제
'       2. 1.0.2: 이상대(2004-12-06)
'       3. 1.1.2: 이상대(2004-12-17)
'          - GetTargetSpcs 함수추가(검사대상조회 기능)
'       4. 1.1.8: 이상대(2005-06-30)
'          - mSentCnt 멤버변수, 프로퍼티 추가
'       5. 1.1.12: 이상대(2005-07-28)
'       6. 1.1.14: 이상대(2005-10-11)
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISLibrarySql
Private mResults    As clsIISResults        '결과내역 컬렉션 클래스
Private mQCResults  As clsIISQCResults      'QC결과내역 컬렉션 클래스
Private mMICResult  As clsIISMICResult      '미생물 결과내역 클래스

Private mSpcYy      As String   '검체연도
Private mSpcNo      As Long     '검체번호
Private mPtId       As String   '환자ID
Private mName       As String   '환자명
Private mSsn        As String   '주민번호
Private mAgeDay     As Long     '일령
Private mAge        As String   '나이
Private mSex        As String   '성별(M:남자, F:여자)
Private mDeptCd     As String   '진료과코드
Private mDeptNm     As String   '진료과명
Private mOrdDoct    As String   '처방의코드
Private mOrdDoctNm  As String   '처방의명
Private mWorkarea   As String   'Workarea
Private mAccDt      As String   '접수일자
Private mAccSeq     As Long     '접수순번
Private mStsCd      As String   '검체상태 (0:처방, 1:채혈, 2:접수, 3: 검사중, 4:중간결과, 5:완료, 6:수정, 7:접수취소)
Private mReqTotCnt  As Long     '총 검사항목 개수
Private mInputCnt   As Long     '입력된 검사항목 개수
Private mSpcCd      As String   '검체코드
Private mSpcNm      As String   '검체명
Private mWardId     As String   '병동코드
Private mWardNm     As String   '병동명
Private mTestDiv    As String   '일반, 미생물, 기타검사 구분 (0:일반검사, 1:기타검사, 2:미생물검사)
Private mQcFg       As String   'QC 여부 (0:일반처방, 1:QC처방)
Private mStatFg     As String   '응급여부 (0:비응급, 1:응급)
Private mBuildCd    As String   '건물코드
Private mHosilId    As String   '호실ID
Private mDetailCnt  As Long     '상세모코드의 개수
Private mColDt      As String   '채혈일자
Private mColTm      As String   '채혈시간
Private mColId      As String   '채혈자ID
Private mRcvDt      As String   '접수일자
Private mRcvTm      As String   '접수시간
Private mRcvId      As String   '접수자ID

Private mSendCnt    As Long     '장비로 전송할 검사항목 개수
Private mRcvCnt     As Long     '장비에서 전송받은 검사항목 개수
Private mSpcPos     As String   '장비에서 검체위치
Private mCtrlCd     As String   'Control Code
Private mLevelCd    As String   'Level Code
Private mLotNo      As String   'Lot No

'## 1.1.8: 이상대(2005-06-31)
'   - 장비로 전송한 검사항목 개수를 나타내는 변수추가
Private mSentCnt    As Long     '장비로 전송한 검사항목 개수

Public Property Get Results() As clsIISResults
    Set Results = mResults
End Property

Public Property Let Results(ByVal vData As clsIISResults)
    Set mResults = vData
End Property

Public Property Get QCResults() As clsIISQCResults
    Set QCResults = mQCResults
End Property

Public Property Let QCResults(ByVal vData As clsIISQCResults)
    Set mQCResults = vData
End Property

Public Property Get MICResult() As clsIISMICResult
    Set MICResult = mMICResult
End Property

Public Property Let MICResult(ByVal vData As clsIISMICResult)
    Set mMICResult = vData
End Property

Public Property Get SpcYy() As String
    SpcYy = mSpcYy
End Property

Public Property Let SpcYy(ByVal vData As String)
    mSpcYy = vData
End Property

Public Property Get SpcNo() As Long
    SpcNo = mSpcNo
End Property

Public Property Let SpcNo(ByVal vData As Long)
    mSpcNo = vData
End Property

Public Property Get PtId() As String
    PtId = mPtId
End Property

Public Property Let PtId(ByVal vData As String)
    mPtId = vData
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal vData As String)
    mName = vData
End Property

Public Property Get Ssn() As String
    Ssn = mSsn
End Property

Public Property Let Ssn(ByVal vData As String)
    mSsn = vData
End Property

Public Property Get AgeDay() As Long
    AgeDay = mAgeDay
End Property

Public Property Let AgeDay(ByVal vData As Long)
    mAgeDay = vData
End Property

Public Property Get Age() As String
    Age = mAge
End Property

Public Property Let Age(ByVal vData As String)
    mAge = vData
End Property

Public Property Get Sex() As String
    Sex = mSex
End Property

Public Property Let Sex(ByVal vData As String)
    mSex = vData
End Property

Public Property Get DeptCd() As String
    DeptCd = mDeptCd
End Property

Public Property Let DeptCd(ByVal vData As String)
    mDeptCd = vData
End Property

Public Property Get DeptNm() As String
    DeptNm = mDeptNm
End Property

Public Property Let DeptNm(ByVal vData As String)
    mDeptNm = vData
End Property

Public Property Get OrdDoctNm() As String
    OrdDoctNm = mOrdDoctNm
End Property

Public Property Let OrdDoctNm(ByVal vData As String)
    mOrdDoctNm = vData
End Property

Public Property Get Workarea() As String
    Workarea = mWorkarea
End Property

Public Property Let Workarea(ByVal vData As String)
    mWorkarea = vData
End Property

Public Property Get AccDt() As String
    AccDt = mAccDt
End Property

Public Property Let AccDt(ByVal vData As String)
    mAccDt = vData
End Property

Public Property Get AccSeq() As Long
    AccSeq = mAccSeq
End Property

Public Property Let AccSeq(ByVal vData As Long)
    mAccSeq = vData
End Property

Public Property Get StsCd() As String
    StsCd = mStsCd
End Property

Public Property Let StsCd(ByVal vData As String)
    mStsCd = vData
End Property

Public Property Get ReqTotCnt() As Long
    ReqTotCnt = mReqTotCnt
End Property

Public Property Let ReqTotCnt(ByVal vData As Long)
    mReqTotCnt = vData
End Property

Public Property Get InputCnt() As Long
    InputCnt = mInputCnt
End Property

Public Property Let InputCnt(ByVal vData As Long)
    mInputCnt = vData
End Property

Public Property Get SpcCd() As String
    SpcCd = mSpcCd
End Property

Public Property Let SpcCd(ByVal vData As String)
    mSpcCd = vData
End Property

Public Property Get SpcNm() As String
    SpcNm = mSpcNm
End Property

Public Property Let SpcNm(ByVal vData As String)
    mSpcNm = vData
End Property

Public Property Get WardId() As String
    WardId = mWardId
End Property

Public Property Let WardId(ByVal vData As String)
    mWardId = vData
End Property

Public Property Get WardNm() As String
    WardNm = mWardNm
End Property

Public Property Let WardNm(ByVal vData As String)
    mWardNm = vData
End Property

Public Property Get TestDiv() As String
    TestDiv = mTestDiv
End Property

Public Property Let TestDiv(ByVal vData As String)
    mTestDiv = vData
End Property

Public Property Get QcFg() As String
    QcFg = mQcFg
End Property

Public Property Let QcFg(ByVal vData As String)
    mQcFg = vData
End Property

Public Property Get StatFg() As String
    StatFg = mStatFg
End Property

Public Property Let StatFg(ByVal vData As String)
    mStatFg = vData
End Property

Public Property Get BuildCd() As String
    BuildCd = mBuildCd
End Property

Public Property Let BuildCd(ByVal vData As String)
    mBuildCd = vData
End Property

Public Property Get HosilId() As String
    HosilId = mHosilId
End Property

Public Property Let HosilId(ByVal vData As String)
    mHosilId = vData
End Property

Public Property Get ColDt() As String
    ColDt = mColDt
End Property

Public Property Let ColDt(ByVal vData As String)
    mColDt = vData
End Property

Public Property Get ColTm() As String
    ColTm = mColTm
End Property

Public Property Let ColTm(ByVal vData As String)
    mColTm = vData
End Property

Public Property Get ColId() As String
    ColId = mColId
End Property

Public Property Let ColId(ByVal vData As String)
    mColId = vData
End Property

Public Property Get RcvDt() As String
    RcvDt = mRcvDt
End Property

Public Property Let RcvDt(ByVal vData As String)
    mRcvDt = vData
End Property

Public Property Get RcvTm() As String
    RcvTm = mRcvTm
End Property

Public Property Let RcvTm(ByVal vData As String)
    mRcvTm = vData
End Property

Public Property Get RcvId() As String
    RcvId = mRcvId
End Property

Public Property Let RcvId(ByVal vData As String)
    mRcvId = vData
End Property

Public Property Get SendCnt() As Long
    SendCnt = mSendCnt
End Property

Public Property Let SendCnt(ByVal vData As Long)
    mSendCnt = vData
End Property

Public Property Get RcvCnt() As Long
    RcvCnt = mRcvCnt
End Property

Public Property Let RcvCnt(ByVal vData As Long)
    mRcvCnt = vData
End Property

Public Property Get SpcPos() As String
    SpcPos = mSpcPos
End Property

Public Property Let SpcPos(ByVal vData As String)
    mSpcPos = vData
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

'## 1.1.8: 이상대(2005-06-31)
'   - 장비로 전송한 검사항목 개수를 나타내는 프로퍼티 추가
Public Property Get SentCnt() As Long
    SentCnt = mSentCnt
End Property

Public Property Let SentCnt(ByVal vData As Long)
    mSentCnt = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 바코드번호를 이용해 접수내역을 조회
'   인수 :
'       - pSpcYy : 검체연도
'       - pSpcNo : 검체번호
'   반환 :
'       - 해당 바코드번호에 대해 접수내역이 없을경우 : -1
'-----------------------------------------------------------------------------'
Public Function GetAccInfo(ByVal pSpcYy As String, ByVal pSpcNo As Long) As Long
    Dim Rs As ADODB.Recordset
    
    '## 1.0.1: 이상대(2004-11-17)
    '   - 접수내역 조회시 검체명, 진료과코드, 병동코드, 채혈일시, 접수일시를 조회 하도록 수정
    '   - 연령(mAge)을 조회시 같이 조회하도록 수정
On Error GoTo Errors:

    Set Rs = DbCon.Execute(mSQL.SelectAccInfo(pSpcYy, CStr(pSpcNo)), , adCmdText)
    If Not (Rs.BOF Or Rs.EOF) Then
        mSpcYy = pSpcYy
        mSpcNo = pSpcNo
        mPtId = Rs.Fields("PTID").Value & ""
        mName = Rs.Fields("NAME").Value & ""
        mSsn = Rs.Fields("SSN").Value & ""
        mAgeDay = CLng(Rs.Fields("AGEDAY").Value)
        mSex = Rs.Fields("SEX").Value & ""
        mDeptCd = Rs.Fields("DEPTCD").Value & ""
        mDeptNm = Rs.Fields("DEPTNM").Value & ""
        mOrdDoctNm = Rs.Fields("DOCTNM").Value & ""
        mWorkarea = Rs.Fields("WORKAREA").Value & ""
        mAccDt = Rs.Fields("ACCDT").Value & ""
        mAccSeq = CLng(Rs.Fields("ACCSEQ").Value)
        mStsCd = Rs.Fields("STSCD").Value & ""
        mReqTotCnt = CLng(Rs.Fields("REQTOTCNT").Value)
        mInputCnt = CLng(Rs.Fields("INPUTCNT").Value)
        mColDt = Rs.Fields("COLDT").Value & ""
        mColTm = Rs.Fields("COLTM").Value & ""
        mRcvDt = Rs.Fields("RCVDT").Value & ""
        mRcvTm = Rs.Fields("RCVTM").Value & ""
        mSpcCd = Rs.Fields("SPCCD").Value & ""
        mSpcNm = Rs.Fields("SPCNM").Value & ""
        mWardId = Rs.Fields("WARDID").Value & ""
        mWardNm = Rs.Fields("WARDNM").Value & ""
        mTestDiv = Rs.Fields("TESTDIV").Value & ""
        mQcFg = Rs.Fields("QCFG").Value & ""
        mStatFg = Rs.Fields("STATFG").Value & ""
        mBuildCd = Rs.Fields("BUILDCD").Value & ""
        mHosilId = Rs.Fields("HOSILID").Value & ""
        mAge = mGetAge(Mid$(mSsn, 1, 6))
    Else
        GetAccInfo = -1
    End If
    
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    GetAccInfo = -1
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISAccInfo", "GetAccInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수번호를 이용해 결과내역 조회
'   인수 :
'       - pAutoRerun : 자동재검 사용유무
'   반환 :
'       - 결과내역이 없을경우에 -1 반환
'-----------------------------------------------------------------------------'
Public Function GetResultInfo(ByVal pAutoRerun As AutoRerunEnum) As Long
    Dim Rs          As ADODB.Recordset
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objTSpc     As clsIISTSpc       '지정검체 클래스
    
    Dim strTestCd   As String           '검사코드
    Dim strTestNm10 As String           '검사명(10자리)
    Dim strTestNm   As String           '검사명(전체)
    Dim strUnit     As String           '단위
    Dim strLastRst  As String           '최근결과
    Dim strRerunFg  As String           '재검여부
    Dim strRstType  As String           '결과유형
    Dim strRstDiv   As String           'A/R유형
    Dim strDetailFg As String           '상세항목순번
    
    '## 1.1.12: 이상대(2005-07-28)
    '   - 자동재검 사용유무 인수추가
On Error GoTo Errors
    Set objResult = New clsIISResult
    Set Rs = objResult.GetResultInfo(mWorkarea, mAccDt, mAccSeq, pAutoRerun)
    Set objResult = Nothing
    If Not (Rs.BOF Or Rs.EOF) Then
        mSpcNm = Rs.Fields("SPCNM").Value & ""
        
        Do Until Rs.EOF
            strTestCd = Rs.Fields("TESTCD").Value
            strTestNm10 = Rs.Fields("TESTNM10").Value
            strTestNm = Rs.Fields("TESTNM").Value
            strUnit = Rs.Fields("UNIT").Value & ""
            strLastRst = Rs.Fields("LASTRST").Value & ""
            strRstType = Rs.Fields("RSTTYPE").Value & ""
            strRstDiv = Rs.Fields("RSTDIV").Value & ""
            strDetailFg = Rs.Fields("DETAILFG").Value & ""
            
            If strRstDiv = "*" And strDetailFg <> "" Then
                mDetailCnt = mDetailCnt + 1
            End If
            
            Set objTSpc = New clsIISTSpc
            With objTSpc
                .AvalVal = CLng(Rs.Fields("AVALVAL").Value)
                .PanicFg = Rs.Fields("PANICFG").Value & ""
                .PanicFrVal = CSng(Rs.Fields("PANICFRVAL").Value)
                .PanicToVal = CSng(Rs.Fields("PANICTOVAL").Value)
                .DeltaFg = Rs.Fields("DELTAFG").Value & ""
                .DeltaFrVal = CLng(Rs.Fields("DELTAFRVAL").Value)
                .DeltaToVal = CLng(Rs.Fields("DELTATOVAL").Value)
            End With
            
            Call mResults.AddInter(strTestCd, strTestNm10, strTestNm, strUnit, strLastRst, strRerunFg, _
                strRstType, strRstDiv, strDetailFg, objTSpc)
            Set objTSpc = Nothing
            
            Rs.MoveNext
        Loop
    Else
        GetResultInfo = -1
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function

Errors:
    GetResultInfo = -1
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISAccInfo", "GetResultInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수번호를 이용해 특수검사 결과내역 조회 - 1.1.14: 이상대(2005-10-11)
'   반환 :
'       - 결과내역이 없을경우에 -1 반환
'-----------------------------------------------------------------------------'
Public Function GetSpecialResultInfo() As Long
    Dim Rs          As ADODB.Recordset
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objTSpc     As clsIISTSpc       '지정검체 클래스
    
    Dim strTestCd   As String           '검사코드
    Dim strTestNm10 As String           '검사명(10자리)
    Dim strTestNm   As String           '검사명(전체)
    
On Error GoTo Errors
    Set objResult = New clsIISResult
    Set Rs = objResult.GetSpecialResultInfo(mWorkarea, mAccDt, mAccSeq)
    Set objResult = Nothing
    If Not (Rs.BOF Or Rs.EOF) Then
        Do Until Rs.EOF
            strTestCd = Rs.Fields("TESTCD").Value
            strTestNm10 = Rs.Fields("TESTNM10").Value
            strTestNm = Rs.Fields("TESTNM").Value
            
            Call mResults.AddInter(strTestCd, strTestNm10, strTestNm)
            
            Rs.MoveNext
        Loop
    Else
        GetSpecialResultInfo = -1
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function

Errors:
    GetSpecialResultInfo = -1
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISAccInfo", "GetSpecialResultInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수번호를 이용해 QC결과내역 조회
'   반환 :
'       - QC결과내역이 없을경우에 -1 반환
'-----------------------------------------------------------------------------'
Public Function GetQCResultInfo() As Long
    Dim Rs          As ADODB.Recordset
    Dim objQCResult As clsIISQCResult       'QC결과내역 클래스
    Dim objQCMaster As clsIISQCMaster       'QC Master 클래스
    
    Dim strTestCd   As String   '검사코드
    Dim strTestNm10 As String   '검사명(10자리)
    Dim strTestNm   As String   '검사명(전체)
    Dim strUnit     As String   '결과단위
    Dim strRstDiv   As String   'A/R유형

On Error GoTo Errors
    Set objQCResult = New clsIISQCResult
    Set Rs = objQCResult.GetQCResultInfo(mWorkarea, mAccDt, mAccSeq)
    Set objQCResult = Nothing
    
    If Not (Rs.BOF Or Rs.EOF) Then
        mSpcNm = mSpcCd
        
        Do Until Rs.EOF
            strTestCd = Rs.Fields("TESTCD").Value
            strTestNm10 = Rs.Fields("TESTNM10").Value & ""
            strTestNm = Rs.Fields("TESTNM").Value & ""
            strUnit = Rs.Fields("UNIT").Value & ""
            strRstDiv = Rs.Fields("RSTDIV").Value & ""
            mCtrlCd = Rs.Fields("CTRLCD").Value & ""
            mLevelCd = Rs.Fields("LEVELCD").Value & ""
            mLotNo = Rs.Fields("LOTNO").Value & ""
            
            Set objQCMaster = New clsIISQCMaster
            With objQCMaster
                .CtrlCd = mCtrlCd
                .LevelCd = mLevelCd
                .LotNo = mLotNo
                .TestCd = strTestCd
                .MeanVal = CSng(Rs.Fields("MEANVAL").Value)
                .SDVal = CSng(Rs.Fields("SDVAL").Value)
                .AvalVal = CLng(Rs.Fields("AVALVAL").Value)
                .RefCd = Rs.Fields("REFCD").Value & ""
                .Unit = Rs.Fields("UNIT").Value & ""
                .CVVal = CSng(Rs.Fields("CVVAL").Value)
                .MinVal = CSng(Rs.Fields("MINVAL").Value)
                .MaxVal = CSng(Rs.Fields("MAXVAL").Value)
                .WMSet = Rs.Fields("WMSET").Value & ""
            End With
            
            Call mQCResults.AddIIS(strTestCd, strTestNm10, strTestNm, strUnit, strRstDiv, mCtrlCd, _
                    mLevelCd, mLotNo, objQCMaster)
            Set objQCMaster = Nothing
            
            Rs.MoveNext
        Loop
    Else
        GetQCResultInfo = -1
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    GetQCResultInfo = -1
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISAccInfo", "GetQCResultInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수번호를 이용해 미생물 결과내역 조회
'   인수 :
'       - pIntNms : 장비별 검사항목 컬렉션
'   반환 :
'       - 미생물 결과내역이 없을경우에 -1 반환
'-----------------------------------------------------------------------------'
Public Function GetMICResultInfo(ByVal pIntNms As clsIISIntNms) As Long
    Dim objIntNm    As clsIISIntNm  '장비별 검사항목 클래스
    Dim strTestCds  As String       '검사코드
    
    '## 1.0.2
    '   - 장비별 검사항목에 등록된 코드만 조회하도록 수정
    For Each objIntNm In pIntNms
        strTestCds = strTestCds & "'" & objIntNm.TestCd & "',"
    Next
    strTestCds = Mid$(strTestCds, 1, Len(strTestCds) - 1)
    Set objIntNm = Nothing
    
    If mMICResult.GetMICResultInfo(mWorkarea, mAccDt, mAccSeq, strTestCds) = -1 Then
        GetMICResultInfo = -1
        Exit Function
    End If
    
    '## 1.0.1: 이상대(2004-11-17)
    '   - 검체명을 따로 조회하는 코드를 삭제(접수내역 조회시 같이 조회)
'    '## 검체명 조회: IIS502, 504에 검체코드가 없어 따로 조회함
'    mSpcNm = GetSpcNm(mSpcCd)
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수번호를 이용해 미생물 결과내역 조회
'   인수 :
'       - pIntNms : 장비별 검사항목 컬렉션
'   반환 :
'       - 미생물 결과내역이 없을경우에 -1 반환
'-----------------------------------------------------------------------------'
Public Function GetMICResultInfoAFB(ByVal pIntNms As clsIISIntNms) As Long
    Dim objIntNm    As clsIISIntNm  '장비별 검사항목 클래스
    Dim strTestCds  As String       '검사코드
    
    '## 1.0.2
    '   - 장비별 검사항목에 등록된 코드만 조회하도록 수정
    For Each objIntNm In pIntNms
        strTestCds = strTestCds & "'" & objIntNm.TestCd & "',"
    Next
    strTestCds = Mid$(strTestCds, 1, Len(strTestCds) - 1)
    Set objIntNm = Nothing
    
    If mMICResult.GetMICResultInfoAFB(mWorkarea, mAccDt, mAccSeq, strTestCds) = -1 Then
        GetMICResultInfoAFB = -1
        Exit Function
    End If
    
    '## 1.0.1: 이상대(2004-11-17)
    '   - 검체명을 따로 조회하는 코드를 삭제(접수내역 조회시 같이 조회)
'    '## 검체명 조회: IIS502, 504에 검체코드가 없어 따로 조회함
'    mSpcNm = GetSpcNm(mSpcCd)
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사결과의 판정(결과유형, 유효숫자, H/L, D/P)
'-----------------------------------------------------------------------------'
Public Sub CheckResult()
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    
    '## 일반검체, QC검체일때 따로 판정
    If mQcFg = "0" Then
        For Each objResult In mResults
            objResult.CheckResult
        Next
    ElseIf mQcFg = "1" Then
        For Each objQCResult In mQCResults
            objQCResult.CheckQCResult
        Next
    End If
    
    Set objResult = Nothing
    Set objQCResult = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Server에 결과입력 (No Verify)
'   인수 :
'       - pEqpCd : 장비코드
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function IISSaveServerNoVfy(ByVal pEqpCd As String) As Boolean
    Dim objEqpTrans As clsIISEqpTrans   '장비 전송내역 클래스
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    
    Dim arySQL()    As String   '결과입력 사용되는 SQL의 배열
    Dim strDate     As String   '전송일자
    Dim strTime     As String   '전송시간
    Dim i           As Long

    '## 장비전송내역 생성
    strDate = Format$(Now, "YYYYMMDD") 'objIisInfo.SpcPos
    strTime = Format$(Time, "HHMMSS")

    '-- 장비 Position 저장용
    For Each objResult In mResults
        mSpcPos = objResult.IntInfo
        Exit For
    Next
    
    Set objEqpTrans = New clsIISEqpTrans
        ReDim arySQL(2)
        arySQL(0) = objEqpTrans.GetDelEqpTransSql(pEqpCd, mSpcYy, mSpcNo)
        arySQL(1) = objEqpTrans.GetAddEqpTransSql(pEqpCd, strDate, strTime, mSpcYy, mSpcNo, mSpcPos)
    Set objEqpTrans = Nothing
    
    '## 결과내역 업데이트
    i = 2
    If mQcFg = "0" Then     '## 일반검체
        ReDim Preserve arySQL(UBound(arySQL) + mResults.Count)
        For Each objResult In mResults
            If objResult.Result <> "" Or objResult.RstDiv = "*" Then
                arySQL(i) = objResult.GetIISRegisterSql(mWorkarea, mAccDt, mAccSeq, pEqpCd)
                i = i + 1
            End If
        Next
        Set objResult = Nothing
    ElseIf mQcFg = "1" Then '## QC검체
        ReDim Preserve arySQL(UBound(arySQL) + mQCResults.Count)
        For Each objQCResult In mQCResults
            If objQCResult.Result <> "" Then
                arySQL(i) = objQCResult.GetQCRegisterSql(mWorkarea, mAccDt, mAccSeq, pEqpCd)
                i = i + 1
                
                If objQCResult.TxtFg = "1" Then
                    ReDim Preserve arySQL(UBound(arySQL) + 1)
                    arySQL(i) = objQCResult.GetQCCommentSql(mWorkarea, mAccDt, mAccSeq)
                    i = i + 1
                End If
            End If
        Next
        Set objQCResult = Nothing
    End If
    
On Error GoTo Errors
    DbCon.BeginTrans
    
    '## DB에 적용
    For i = 0 To UBound(arySQL) - 1
        Debug.Print arySQL(i)
        If arySQL(i) = "" Then Exit For
        DbCon.Execute arySQL(i), , adCmdText + adExecuteNoRecords
    Next i
    
    DbCon.CommitTrans
    IISSaveServerNoVfy = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISAccInfo", "IISSaveServerNoVfy", Err.Description, Now
    Error.SetLog App.EXEName, "clsIISAccInfo", "IISSaveServerNoVfy", "SQL: " & arySQL(i), Now
End Function

'-----------------------------------------------------------------------------'
'   기능 : Server에 자동결과 등록
'   인수 :
'       - pEqpCd   : 장비코드
'       - pHLCheck : 자동결과 등록시 H/L결과처리 여부
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function IISSaveServerAutoVfy(ByVal pEqpCd As String, ByVal pHLCheck As HLCheckEnum) As Boolean
    Dim objEqpTrans As clsIISEqpTrans   '장비 전송내역 클래스
    Dim objResult   As clsIISResult     '결과내역 클래스
    
    Dim arySQL()    As String   '결과입력 사용되는 SQL의 배열
    Dim aryParent() As String   '상세모코드 저장배열
    Dim aryChild()  As String   '상세자코드 저장배열
    Dim strDate     As String   '전송일자
    Dim strTime     As String   '전송시간
    Dim strParentCd As String   '상세모코드
    Dim strChildCd  As String   '상세자코드
    Dim strDetailFg As String   '상세항목 순번
    Dim SQL         As String
    Dim blnDetail   As Boolean  '상세모코드의 자동결과등록 가능여부(True:가능, False:불가능)
    Dim lngOrderVfy As Long     '처방,접수내역의 결과등록 가능여부
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    
    '## 1.결과내역 업데이트
    strDate = Format$(Now, "YYYYMMDD")
    strTime = Format$(Time, "HHMMSS")
    ReDim arySQL(mResults.Count)
    ReDim aryParent(mDetailCnt)
    
    '## 1.1.단일항목에 대한 쿼리생성
    For Each objResult In mResults
        With objResult
            '## 상세모코드
            If .RstDiv = "*" Then
                aryParent(j) = .TestCd
                j = j + 1
            '## 단일항목, 상세자코드
            ElseIf .Result <> "" And .DetailFg = "" Then
                If .DPDiv <> "" Then
                    arySQL(i) = .GetIISRegisterSql(mWorkarea, mAccDt, mAccSeq, pEqpCd)
                    'Debug.Print "[" & .TestCd & "]: " & arySQL(i)
                ElseIf pHLCheck = ccCheck And .HLDiv <> "" Then
                    arySQL(i) = .GetIISRegisterSql(mWorkarea, mAccDt, mAccSeq, pEqpCd)
                    'Debug.Print "[" & .TestCd & "]: " & arySQL(i)
                Else
                    arySQL(i) = .GetIISVfySql(mWorkarea, mAccDt, mAccSeq, pEqpCd, strDate, strTime)
                    'Debug.Print "[" & .TestCd & "]: " & arySQL(i)
                    mInputCnt = mInputCnt + 1
                End If
                
                i = i + 1
            End If
        End With
    Next
    
    For j = 0 To mDetailCnt - 1
        strParentCd = aryParent(j)
        strDetailFg = mResults(strParentCd).DetailFg
        blnDetail = True
        ReDim aryChild(0)
        k = 0
        
        For Each objResult In mResults
            With objResult
                If strDetailFg = .DetailFg And strParentCd <> .TestCd Then
                    '## 상세모코드 배열에 담고
                    ReDim Preserve aryChild(k + 1)
                    aryChild(k) = .TestCd
                    
                    '## AutoVfy 체크하고
                    If .Result = "" And .RstDiv = "R" Then
                        blnDetail = False:  .AutoVfy = False
                    ElseIf .DPDiv <> "" Then
                        blnDetail = False:  .AutoVfy = False
                    ElseIf pHLCheck = ccCheck And .HLDiv <> "" Then
                        blnDetail = False:  .AutoVfy = False
                    Else
                        .AutoVfy = True
                    End If
                    k = k + 1
                End If
            End With
        Next
        
        '## AutoVfy = True일때
        '   - 1.상세모코드 Vfy Update
        '   - 2.상세자코드 배열을 돌면서 .AutoVfy을 체크하여 VfyUpdate
        '## AutoVfy = False일때
        '   - 1.상세모코드 일반 Update
        '   - 2.상세자코드 배열을 돌면서 무조건 일반 Update
        If blnDetail Then
            arySQL(i) = mResults(strParentCd).GetIISVfySql(mWorkarea, mAccDt, mAccSeq, pEqpCd, strDate, strTime)
            Debug.Print "[" & strParentCd & "]: " & arySQL(i)
            mInputCnt = mInputCnt + 1
            i = i + 1
            
            For k = 0 To UBound(aryChild) - 1
                strChildCd = aryChild(k)
                If mResults(strChildCd).AutoVfy Then
                    arySQL(i) = mResults(strChildCd).GetIISVfySql(mWorkarea, mAccDt, mAccSeq, pEqpCd, strDate, strTime)
                    Debug.Print "[" & strChildCd & "]: " & arySQL(i)
                    i = i + 1
                Else
                    arySQL(i) = mResults(strChildCd).GetIISRegisterSql(mWorkarea, mAccDt, mAccSeq, pEqpCd)
                    Debug.Print "[" & strChildCd & "]: " & arySQL(i)
                    i = i + 1
                End If
            Next k
        Else
            arySQL(i) = mResults(strParentCd).GetIISRegisterSql(mWorkarea, mAccDt, mAccSeq, pEqpCd)
            Debug.Print "[" & strParentCd & "]: " & arySQL(i)
            i = i + 1
            
            For k = 0 To UBound(aryChild) - 1
                strChildCd = aryChild(k)
                arySQL(i) = mResults(strChildCd).GetIISRegisterSql(mWorkarea, mAccDt, mAccSeq, pEqpCd)
                Debug.Print "[" & strChildCd & "]: " & arySQL(i)
                i = i + 1
            Next k
        End If
    Next j
    
    'Debug.Print "Input Count: " & CStr(mInputCnt)
    'Debug.Print "i=" & CStr(i)
    
    '## 2.1.접수내역 InputCnt 업데이트
    If mInputCnt <> 0 Then
        ReDim Preserve arySQL(i + 1)
        arySQL(i) = mSQL.UpdateInputCnt(mSpcYy, mSpcNo, mInputCnt)
        Debug.Print arySQL(i)
    End If

On Error GoTo Errors
    DbCon.BeginTrans
    
    '## DB에 적용
    For i = 0 To UBound(arySQL) - 1
        If arySQL(i) = "" Then Exit For
        DbCon.Execute arySQL(i), , adCmdText + adExecuteNoRecords
    Next i
    
    '## 2.2.TotCnt=InputCnt를 조회해서 같으면 접수내역, 처방내역 업데이트
    lngOrderVfy = CheckVfy
    If lngOrderVfy = -1 Then
        '-- 2020-11-16 수정 및 추가
        'ReDim arySQL(2)
        ReDim arySQL(3)

        '## 장비전송내역 생성
        Set objEqpTrans = New clsIISEqpTrans
        arySQL(1) = objEqpTrans.GetDelEqpTransSql(pEqpCd, mSpcYy, mSpcNo)
        arySQL(2) = objEqpTrans.GetAddEqpTransSql(pEqpCd, strDate, strTime, mSpcYy, mSpcNo, mSpcPos)
        'Debug.Print arySQL(1)
        'Debug.Print arySQL(2)
        Set objEqpTrans = Nothing
    ElseIf lngOrderVfy = 1 Then
        '-- 2020-11-16 수정 및 추가
        'ReDim arySQL(2)
        ReDim arySQL(3)
        
        '## 접수내역 업데이트
        arySQL(0) = GetAccInfoVerifySql(strDate, strTime)
        Debug.Print arySQL(0)
        
        '## 처방내역 업데이트
        arySQL(1) = GetOrderVerifySql("5", strDate, strTime)
        Debug.Print arySQL(1)
        
        '-- 2020-11-09 추가
        arySQL(2) = objEqpTrans.GetAddEqpTransSql(pEqpCd, strDate, strTime, mSpcYy, mSpcNo, mSpcPos)
    Else
        ReDim arySQL(3)
        
        '## 처방내역 업데이트
        arySQL(0) = GetOrderVerifySql("4", strDate, strTime)
        Debug.Print arySQL(0)
        
        '## 장비전송내역 생성
        Set objEqpTrans = New clsIISEqpTrans
        arySQL(1) = objEqpTrans.GetDelEqpTransSql(pEqpCd, mSpcYy, mSpcNo)
        arySQL(2) = objEqpTrans.GetAddEqpTransSql(pEqpCd, strDate, strTime, mSpcYy, mSpcNo, mSpcPos)
        'Debug.Print arySQL(1)
        'Debug.Print arySQL(2)
        Set objEqpTrans = Nothing
    End If
    
    
    '## DB에 적용
    For i = 0 To UBound(arySQL) - 1
        If arySQL(i) = "" Then Exit For
        DbCon.Execute arySQL(i), , adCmdText + adExecuteNoRecords
    Next i
    
    DbCon.CommitTrans
    Set objResult = Nothing
    IISSaveServerAutoVfy = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISAccInfo", "IISSaveServerAutoVfy", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 미생물 감수성결과 저장
'   인수 :
'       - pIntInfo : 장비에서 수신한결과 저장클래스
'-----------------------------------------------------------------------------'
Public Function MicroSaveServer(ByVal pIntInfo As clsIISIntInfo) As Boolean
    Dim arySQL(2) As String
    Dim i         As Long
    
    '## 미생물 결과내역 업데이트
    mMICResult.SenFg = "Y"
    arySQL(0) = mMICResult.SaveMicroResultSql(mWorkarea, mAccDt, mAccSeq)
    
    '## 미생물 감수성결과 입력
    arySQL(1) = mMICResult.SaveMICResultSql(mWorkarea, mAccDt, mAccSeq, pIntInfo)

On Error GoTo Errors
    DbCon.BeginTrans
    
    '## DB에 적용
    For i = 0 To 1
        DbCon.Execute arySQL(i), , adCmdText + adExecuteNoRecords
    Next i
    
    DbCon.CommitTrans
    MicroSaveServer = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISAccInfo", "MicroSaveServer", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 미생물 배양결과 저장
'   인수 :
'       - pIntInfo : 장비에서 수신한결과 저장클래스
'-----------------------------------------------------------------------------'
Public Function CultureSaveServer(ByVal pIntInfo As clsIISIntInfo) As Boolean
    Dim arySQL(2) As String
    Dim i         As Long
    Dim strTemp   As String
    
    If pIntInfo.VitekNo = "C12" Then    '-- 균동정중 결과저장 않음
        CultureSaveServer = True
        Exit Function
    End If
    
    If mMICResult.GetCultureResultInfo(mWorkarea, mAccDt, mAccSeq) = 1 Then
        If mStsCd < 5 Then
            '## 임상병리 접수내역 업데이트
            arySQL(0) = mMICResult.SaveCultureInfoSql(mWorkarea, mAccDt, mAccSeq, pIntInfo)
            '## 미생물 배양 결과내역 업데이트
            arySQL(1) = mMICResult.SaveCultureResultSql(mWorkarea, mAccDt, mAccSeq, pIntInfo)
        
        End If
    End If
    
On Error GoTo Errors
    DbCon.BeginTrans
    
    '## DB에 적용
    For i = 0 To 1
        DbCon.Execute arySQL(i), , adCmdText + adExecuteNoRecords
    Next i
    
    DbCon.CommitTrans
    CultureSaveServer = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISAccInfo", "MicroSaveServer", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 미생물 배양결과 저장
'   인수 :
'       - pIntInfo : 장비에서 수신한결과 저장클래스
'-----------------------------------------------------------------------------'
Public Function CultureSaveServerAFB(ByVal pIntInfo As clsIISIntInfo, ByVal pTestCds As String) As Boolean
    Dim arySQL(2) As String
    Dim i         As Long
    Dim strTemp   As String
    
    If mMICResult.GetCultureResultInfoAFB(mWorkarea, mAccDt, mAccSeq, pTestCds) = 1 Then
        If mStsCd < 5 Then
            '## 임상병리 접수내역 업데이트
            'arySQL(0) = mMICResult.SaveCultureInfoSqlAFB(mWorkarea, mAccDt, mAccSeq, pIntInfo)
            '## 미생물 배양 결과내역 업데이트
            arySQL(1) = mMICResult.SaveCultureResultSqlAFB(mWorkarea, mAccDt, mAccSeq, pTestCds)
        
        End If
    End If
    
On Error GoTo Errors
    DbCon.BeginTrans
    
    '## DB에 적용
    For i = 0 To 1
        If arySQL(i) <> "" Then
            DbCon.Execute arySQL(i), , adCmdText + adExecuteNoRecords
        End If
    Next i
    
    DbCon.CommitTrans
    CultureSaveServerAFB = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISAccInfo", "CultureSaveServerAFB", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 미생물 배양결과 저장
'   인수 :
'       - pIntInfo : 장비에서 수신한결과 저장클래스
'-----------------------------------------------------------------------------'
Public Function CultureSaveServerAFB_Eqp(ByVal pIntInfo As clsIISIntInfo, ByVal pTestCds As String, ByVal pEqpCd As String) As Boolean
    Dim arySQL(2) As String
    Dim i         As Long
    Dim strTemp   As String
    
    If mMICResult.GetCultureResultInfoAFB(mWorkarea, mAccDt, mAccSeq, pTestCds) = 1 Then
        If mStsCd < 5 Then
            '## 임상병리 접수내역 업데이트
            'arySQL(0) = mMICResult.SaveCultureInfoSqlAFB(mWorkarea, mAccDt, mAccSeq, pIntInfo)
            '## 미생물 배양 결과내역 업데이트
            arySQL(1) = mMICResult.SaveCultureResultSqlAFB_Eqp(mWorkarea, mAccDt, mAccSeq, pTestCds, pEqpCd)
        
        End If
    End If
    
On Error GoTo Errors
    DbCon.BeginTrans
    
    '## DB에 적용
    For i = 0 To 1
        If arySQL(i) <> "" Then
            DbCon.Execute arySQL(i), , adCmdText + adExecuteNoRecords
        End If
    Next i
    
    DbCon.CommitTrans
    CultureSaveServerAFB_Eqp = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISAccInfo", "CultureSaveServerAFB_Eqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function


Public Function CommentSaveServer(ByVal pIntInfo As clsIISIntInfo, ByVal pComment As String) As Boolean
    Dim arySQL(2) As String
    Dim i         As Long
    Dim strBarNo   As String    '바코드번호
    Dim strSpcYy   As String    '검체연도
    Dim lngSpcNo   As Long      '검체번호
    Dim lngSeq     As Long
    Dim strCmtVal  As String
    
    strBarNo = pIntInfo.BarNo
    strSpcYy = Mid$(strBarNo, 1, SPCYYLEN)
    lngSpcNo = Val(Mid$(strBarNo, SPCYYLEN + 1, SPCNOLEN))
    
    '## s2lab201 업데이트 footnotefg = 1
'    arySQL(0) = mSQL.UpdateIISComment(strSpcYy, lngSpcNo)
    
    lngSeq = GetMaxSeq(mWorkarea, mAccDt, mAccSeq)
'    strCmtVal = GetSeqCmtVal(mWorkarea, mAccDt, mAccSeq)
    
    '## s2lab304 Foot Note 내역결과 입력
    If lngSeq = 0 Then
'        arySQL(1) = mSQL.InsertIISComment(mWorkarea, mAccDt, mAccSeq, 1, "", pComment)
    Else
'        arySQL(1) = mSQL.Update_IISComment(mWorkarea, mAccDt, mAccSeq, 1, "", pComment)
    End If
    
On Error GoTo Errors
    DbCon.BeginTrans
    
    '## DB에 적용
    For i = 0 To 1
        DbCon.Execute arySQL(i), , adCmdText + adExecuteNoRecords
    Next i
    
    DbCon.CommitTrans
    CommentSaveServer = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISAccInfo", "MicroSaveServer", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
    
End Function

Public Function GetMaxSeq(ByVal pWorkarea As String, ByVal pAccDt As String, ByVal pAccSeq As Long) As Long
    Dim Rs As ADODB.Recordset
    
    GetMaxSeq = 0
    
'    Set Rs = DbCon.Execute(mSQL.SelectMaxCommentSeq(pWorkarea, pAccDt, pAccSeq), , adCmdText)
'    If Not (Rs.BOF Or Rs.EOF) Then
'        If IsNull(Rs.Fields("MAXSEQ").Value) Then
'            GetMaxSeq = 1
'        Else
'            GetMaxSeq = CLng(Rs.Fields("MAXSEQ").Value)
'        End If
'    End If
    
    If Not (Rs.BOF Or Rs.EOF) Then
        GetMaxSeq = CLng(Rs.Fields("CNT").Value)
    End If
    
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISCAccInfo", "GetMaxItemSeq", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

Public Function GetSeqCmtVal(ByVal pWorkarea As String, ByVal pAccDt As String, ByVal pAccSeq As Long) As String
    Dim Rs As ADODB.Recordset
    
'    Set Rs = DbCon.Execute(mSQL.SelectMaxCommentSeq(pWorkarea, pAccDt, pAccSeq), , adCmdText)
    If Not (Rs.BOF Or Rs.EOF) Then
        If IsNull(Rs.Fields("RSTTXT").Value) Then
            GetSeqCmtVal = ""
        Else
            GetSeqCmtVal = Trim(Rs.Fields("RSTTXT").Value)
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISCAccInfo", "GetSeqCmtVal", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수내역 TotCnt, InputCnt를 조회하여 결과등록 가능여부 조회
'   반환 :
'       - -1 : 결과입력된 항목이 없음
'       - 0  : 결과등록 불가능
'       - 1  : 결과등록 가능
'-----------------------------------------------------------------------------'
Public Function CheckVfy() As Long
    Dim Rs          As ADODB.Recordset
    Dim lngTotCnt   As Long     '총검사항목 개수
    Dim lngInputCnt As Long     '입력된 검사항목 개수
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectTotCnt(mSpcYy, mSpcNo), , adCmdText)
    If Not (Rs.BOF Or Rs.EOF) Then
        lngTotCnt = CLng(Rs.Fields("TOTCNT").Value)
        lngInputCnt = CLng(Rs.Fields("INPUTCNT").Value)
    End If
    Rs.Close
    Set Rs = Nothing
    
    '## TotCnt와 InputCnt가 같으면 Verify가 가능
    If lngInputCnt = 0 Then
        CheckVfy = -1
    ElseIf lngTotCnt = lngInputCnt Then
        CheckVfy = 1
    Else
        CheckVfy = 0
    End If
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISAccInfo", "CheckVfy", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수내역을 결과확인하는 쿼리조회
'   인수 :
'       - pVfyDt : 결과확인일자
'       - pVfyTm : 결과확인시간
'-----------------------------------------------------------------------------'
Public Function GetAccInfoVerifySql(ByVal pVfyDt As String, ByVal pVfyTm As String) As String
    Dim SQL As String
    
    SQL = mSQL.UpdateAccInfoVfy(mSpcYy, mSpcNo, "5", pVfyDt, pVfyTm, EmpId)
    
    GetAccInfoVerifySql = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 처방내역을 결과확인하는 쿼리조회
'   인수 :
'       - pStsCd  : 검체상태
'       - pVfyDt : 결과확인일자
'       - pVfyTm : 결과확인시간
'-----------------------------------------------------------------------------'
Public Function GetOrderVerifySql(ByVal pStsCd As String, ByVal pVfyDt As String, _
                    ByVal pVfyTm As String) As String
    Dim SQL As String
    
    SQL = mSQL.UpdateOrderVfy(mWorkarea, mAccDt, mAccSeq, pStsCd, pVfyDt, pVfyTm, EmpId)
    
    GetOrderVerifySql = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검체연도, 번호를 이용해 바코드번호를 조회
'   반환 :
'       - -1 : 오류
'       - 바코드번호
'-----------------------------------------------------------------------------'
Public Function GetBarNo() As String
    If mSpcYy = "" Then
        GetBarNo = "-1"
    Else
        GetBarNo = mSpcYy & Format$(mSpcNo, String$(SPCNOLEN, "0"))
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수번호를 이용해 바코드번호를 조회
'   반환 :
'       - -1 : 오류
'       - 바코드번호
'-----------------------------------------------------------------------------'
Public Function GetBarNoByAccNo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectBarNo(pWorkarea, pAccDt, pAccSeq), , adCmdText)
    If Rs.BOF Or Rs.EOF Then
        GetBarNoByAccNo = "-1"
    Else
        GetBarNoByAccNo = Rs.Fields("SPCYY").Value & Format$(Rs.Fields("SPCNO").Value, String$(SPCNOLEN, "0"))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISAccInfo", "GetBarNoByAccNo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검체코드를 이용해 검체명(약어) 조회
'   인수 :
'       - pSpcCd : 검체코드
'   반환 : 검체명
'-----------------------------------------------------------------------------'
Public Function GetSpcNm(Optional ByVal pSpcCd As String) As String
    Dim Rs          As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectSpcNm(pSpcCd), , adCmdText)
    If Not (Rs.BOF Or Rs.EOF) Then
        GetSpcNm = Rs.Fields("SPCNM").Value & ""
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISAccInfo", "GetSpcNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당장비의 검사대상리스트 조회 - 1.1.2:이상대(2004-12-17)
'   인수 :
'       - pEqpCd  : 장비코드
'       - pFromDt : 접수일자 From
'       - pToDt   : 접수일자 To
'-----------------------------------------------------------------------------'
Public Function GetTargetSpcs(ByVal pEqpCd As String, ByVal pFromDt As String, _
                              ByVal pToDt As String) As Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectTargetSpcs(pEqpCd, pFromDt, pToDt), , adCmdText)
    Set GetTargetSpcs = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISAccInfo", "GetTargetSpcs", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function


'-----------------------------------------------------------------------------'
'   기능 : 해당장비의 검사대상리스트 조회 - 1.1.2:이상대(2004-12-17)
'   인수 :
'       - pEqpCd  : 장비코드
'       - pFromDt : 접수일자 From
'       - pToDt   : 접수일자 To
'-----------------------------------------------------------------------------'
Public Function GetTargetSpcs_ABO(ByVal pEqpCd As String, ByVal pFromDt As String, _
                              ByVal pToDt As String) As Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectTargetSpcs_ABO(pEqpCd, pFromDt, pToDt), , adCmdText)
    Set GetTargetSpcs_ABO = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISAccInfo", "GetTargetSpcs_ABO", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당장비의 검사대상리스트 조회 - 1.1.2:이상대(2004-12-17)
'   인수 :
'       - pEqpCd  : 장비코드
'       - pFromDt : 접수일자 From
'       - pToDt   : 접수일자 To
'-----------------------------------------------------------------------------'
Public Function GetTargetSpcs_BUNJA(ByVal pEqpCd As String, ByVal pFromDt As String, _
                              ByVal pToDt As String, ByVal pTestCd As String) As Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectTargetSpcs_BUNJA(pEqpCd, pFromDt, pToDt, pTestCd), , adCmdText)
    Set GetTargetSpcs_BUNJA = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISAccInfo", "GetTargetSpcs_BUNJA", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function


'-----------------------------------------------------------------------------'
'   기능 : 해당장비의 검사대상리스트 조회 - 1.1.2:이상대(2004-12-17)
'   인수 :
'       - pEqpCd  : 장비코드
'       - pFromDt : 접수일자 From
'       - pToDt   : 접수일자 To
'-----------------------------------------------------------------------------'
Public Function GetTargetSpcs_Allergy(ByVal pEqpCd As String, ByVal pFromDt As String, _
                              ByVal pToDt As String) As Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectTargetSpcs_Allergy(pEqpCd, pFromDt, pToDt), , adCmdText)
    Set GetTargetSpcs_Allergy = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISAccInfo", "GetTargetSpcs_Allergy", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당장비의 검사대상리스트 조회 - 1.1.7:이상대(2005-05-23)
'   인수 :
'       - pEqpCd : 장비코드
'       - pDate  : 조회날짜
'-----------------------------------------------------------------------------'
Public Function GetTargetSpcX(ByVal pEqpCd As String, ByVal pDate As String) As Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectTargetSpcsX(pEqpCd, pDate), , adCmdText)
    Set GetTargetSpcX = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISAccInfo", "GetTargetSpcX", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사대상 리스트 조회후 조회 Flag Update - 1.1.7: 이상대(2005-05-23)
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Public Sub SetTargetUpdate(ByVal pBarNo As String)
    Dim strSpcYy As String  '검체연도
    Dim lngSpcNo As Long    '검체번호
    Dim SQL      As String
    
    pBarNo = Format$(pBarNo, String$(SPCLEN, "#"))
    strSpcYy = Mid$(pBarNo, 1, SPCYYLEN)
    lngSpcNo = Val(Mid$(pBarNo, SPCYYLEN + 1, SPCNOLEN))
    
On Error GoTo Errors
    DbCon.BeginTrans
    SQL = mSQL.UpdateTargetSpc(strSpcYy, lngSpcNo)
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    Exit Sub
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISAccInfo", "SetTargetUpdate", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로 전송할 검사를 모두 전송했는지 조회 - 1.1.8; 이상대(2005-06-30)
'          Vidas 장비에서 사용
'   반환 :
'       - True(모두전송), False(전송할 검사가 남음)
'-----------------------------------------------------------------------------'
Public Function IsAllSend() As Boolean
    If mSendCnt = mSentCnt Then
        IsAllSend = True
    Else
        IsAllSend = False
    End If
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISLibrarySql
    Set mResults = New clsIISResults
    Set mQCResults = New clsIISQCResults
    Set mMICResult = New clsIISMICResult
End Sub

Private Sub Class_Terminate()
    Set mMICResult = Nothing
    Set mQCResults = Nothing
    Set mResults = Nothing
    Set mSQL = Nothing
End Sub

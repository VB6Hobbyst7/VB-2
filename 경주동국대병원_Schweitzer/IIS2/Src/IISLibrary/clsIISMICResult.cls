VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISMICResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISMICResult.cls
'   작성자  : 이상대
'   내  용  : 미생물 결과내역 클래스
'   작성일  : 20034-09-08
'   버  전  :
'       1. 1.0.2: 이상대(2004-12-06)
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISMicroSql
Private mWSBody     As clsIISMICWSBody      '미생물 WorkSheet Body 클래스
Private mSensi      As clsIISMICSensi       '미생물 감수성결과 클래스

Private mTestCd     As String   '검사코드
Private mTestNm10   As String   '검사명(10자리)
Private mTestNm     As String   '검사명(Full)
Private mRstCd      As String   '결과
Private mSenFg      As String   '감수성 결과유무(Y:유, Null:무)
Private mStsCd      As String   '검체상태(0:처방, 1:채혈, 2:접수, 3:검사중, 4:중간결과, 5:결과, 6:수정, 7:접수취소)
Private mRstType    As String   '결과유형(S:일반감수성, M:MIC감수성, G:Gram Stain, M:AFB Culture, F:Fungus Culture, N:Fungus Stain, B:AFB Stain, A:AFB Sensitivity)
Private mMfySeq     As Long     '수정회수
Private mRstDiv     As String   '결과구분(A:Alternative, R:Require, *:결과없음)
Private mDetailFg   As String   '상세항목순번
Private mVfyDt      As String   '결과확인일자
Private mVfyTm      As String   '결과확인시간
Private mVfyId      As String   '결과확인자ID
Private mPtId       As String   '환자ID

Public Property Get WSBody() As clsIISMICWSBody
    Set WSBody = mWSBody
End Property

Public Property Let WSBody(ByVal vData As clsIISMICWSBody)
    Set mWSBody = vData
End Property

Public Property Get Sensi() As clsIISMICSensi
    Set Sensi = mSensi
End Property

Public Property Let Sensi(ByVal vData As clsIISMICSensi)
    Set mSensi = vData
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

Public Property Get RstCd() As String
    RstCd = mRstCd
End Property

Public Property Let RstCd(ByVal vData As String)
    mRstCd = vData
End Property

Public Property Get SenFg() As String
    SenFg = mSenFg
End Property

Public Property Let SenFg(ByVal vData As String)
    mSenFg = vData
End Property

Public Property Get StsCd() As String
    StsCd = mStsCd
End Property

Public Property Let StsCd(ByVal vData As String)
    mStsCd = vData
End Property

Public Property Get RstType() As String
    RstType = mRstType
End Property

Public Property Let RstType(ByVal vData As String)
    mRstType = vData
End Property

Public Property Get MfySeq() As Long
    MfySeq = mMfySeq
End Property

Public Property Let MfySeq(ByVal vData As Long)
    mMfySeq = vData
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

Public Property Get VfyDt() As String
    VfyDt = mVfyDt
End Property

Public Property Let VfyDt(ByVal vData As String)
    mVfyDt = vData
End Property

Public Property Get VfyTm() As String
    VfyTm = mVfyTm
End Property

Public Property Let VfyTm(ByVal vData As String)
    mVfyTm = vData
End Property

Public Property Get VfyId() As String
    VfyId = mVfyId
End Property

Public Property Let VfyId(ByVal vData As String)
    mVfyId = vData
End Property

Public Property Get PtId() As String
    PtId = mPtId
End Property

Public Property Let PtId(ByVal vData As String)
    mPtId = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 접수번호를 이용해 미생물 결과내역 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pTestCds  : 장비별 검사항목에 등록된 검사코드
'   반환 :
'       -1 : 결과내역이 없음
'-----------------------------------------------------------------------------'
Public Function GetMICResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCds As String) As Long
    Dim Rs As ADODB.Recordset
    
    '## 1.0.2: 이상대(2004-12-06)
    '   - 장비별 검사항목에 등록된 코드만 조회하도록 pTestCds 추가
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectMICResultInfo(pWorkarea, pAccDt, pAccSeq, pTestCds), , adCmdText)
    If Not (Rs.BOF Or Rs.EOF) Then
        mTestCd = Rs.Fields("TESTCD").Value
        mTestNm10 = Rs.Fields("TESTNM10").Value & ""
        mTestNm = Rs.Fields("TESTNM").Value & ""
        mRstType = Rs.Fields("RSTTYPE").Value & ""
        mMfySeq = CLng(Rs.Fields("MFYSEQ").Value)
        mPtId = Rs.Fields("PTID").Value & ""
        
        With mWSBody
            .WsCd = Rs.Fields("WSCD").Value
            .WsUnit = Rs.Fields("WSUNIT").Value
            .SelFg = Rs.Fields("SELFG").Value & ""
            .GenFg = Rs.Fields("GENFG").Value & ""
            .ScFg = Rs.Fields("SCFG").Value & ""
        End With
        GetMICResultInfo = 1
    Else
        GetMICResultInfo = -1
    End If
    
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    GetMICResultInfo = -1
    Error.SetLog App.EXEName, "clsIISMICResult", "GetMICResultInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수번호를 이용해 미생물 결과내역 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'   반환 :
'       -1 : 결과내역이 없음
'-----------------------------------------------------------------------------'
Public Function GetCultureResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, ByVal pAccSeq As Long) As Long
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectCultureResultInfo(pWorkarea, pAccDt, pAccSeq), , adCmdText)
    If Not (Rs.BOF Or Rs.EOF) Then
        mStsCd = Rs.Fields("STSCD").Value & ""
        mTestCd = Rs.Fields("TESTCD").Value & ""
        
        GetCultureResultInfo = 1
    Else
        GetCultureResultInfo = -1
    End If
    
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    GetCultureResultInfo = -1
    Error.SetLog App.EXEName, "clsIISMICResult", "GetCultureResultInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 미생물결과 저장쿼리 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'-----------------------------------------------------------------------------'
Public Function SaveMicroResultSql(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
                    
    mSenFg = "Y"
    SaveMicroResultSql = mSQL.UpdateMicroResult(pWorkarea, pAccDt, pAccSeq, mSenFg)
End Function

'-----------------------------------------------------------------------------'
'   기능 : 미생물 배양결과 저장쿼리 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'-----------------------------------------------------------------------------'
Public Function SaveCultureResultSql(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pIntInfo As clsIISIntInfo) As String
                    
    Dim strResult As String
    
    '-- 결과
    strResult = pIntInfo.VitekNo & "|" & mTestCd
    If pIntInfo.VitekNo <> "C12" Then
        SaveCultureResultSql = mSQL.UpdateCultureResult(pWorkarea, pAccDt, pAccSeq, strResult)
    Else
        SaveCultureResultSql = ""
    End If
    
End Function

'-----------------------------------------------------------------------------'
'   기능 : 미생물 배양 접수내역 저장쿼리 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'-----------------------------------------------------------------------------'
Public Function SaveCultureInfoSql(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pIntInfo As clsIISIntInfo) As String
                    
    
    Dim strResult As String
    
    '-- 결과
    strResult = pIntInfo.VitekNo
    If strResult <> "C12" Then  '양성은 업데이트 하지 않는다.
        SaveCultureInfoSql = mSQL.UpdateCultureInfoResult(pWorkarea, pAccDt, pAccSeq, strResult)
    Else
        SaveCultureInfoSql = ""
    End If
    
End Function
'-----------------------------------------------------------------------------'
'   기능 : 미생물 감수성결과 저장쿼리 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pIntInfo  : 장비에서 수신한결과 저장클래스
'-----------------------------------------------------------------------------'
Public Function SaveMICResultSql(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pIntInfo As clsIISIntInfo) As String
    Dim objDrug   As clsIISMICDrug  '항생제결과 클래스
    Dim strResult As String         '항생제결과
    Dim SQL       As String
    
    '## NOTE: 문자열 반드시 검증할것!
    '## 항생제결과를 DB에 입력쿼리형태로 변형
    For Each objDrug In pIntInfo.Drugs
        With objDrug
            strResult = strResult & ", '" & .DrugCd & ";" & .RstCd & ";" & .Volumn & ";;;;'"
        End With
    Next
    Set objDrug = Nothing
    
    With mSensi
        .MfySeq = 0
        .MicFg = "C"
        .MnmCd = pIntInfo.MnmCd
        .MicFg = "C"
        .MqtCd = IISMQTCD
        .SCnt = pIntInfo.Drugs.Count
        .Result = strResult
        SQL = .GetSaveMICResultSql(pWorkarea, pAccDt, pAccSeq, mTestCd)
    End With
    
    SaveMICResultSql = SQL
End Function



Private Sub Class_Initialize()
    Set mWSBody = New clsIISMICWSBody
    Set mSensi = New clsIISMICSensi
    Set mSQL = New clsIISMicroSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
    Set mSensi = Nothing
    Set mWSBody = Nothing
End Sub

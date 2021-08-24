VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISCResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISCResult.cls
'   작성자  : 오세원
'   내  용  : ClientDb의 결과내역 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISLibrarySql

Private mIntNm      As String   '화면표시 검사명
Private mTestCd     As String   '검사코드
Private mResult     As String   '결과 (장비에서 송신한 결과)
Private mHLDiv      As String   'H/L
Private mDPDiv      As String   'D/P

Public Property Get IntNm() As String
    IntNm = mIntNm
End Property

Public Property Let IntNm(ByVal vData As String)
    mIntNm = vData
End Property

Public Property Get TestCd() As String
    TestCd = mTestCd
End Property

Public Property Let TestCd(ByVal vData As String)
    mTestCd = vData
End Property

Public Property Get Result() As String
    Result = mResult
End Property

Public Property Let Result(ByVal vData As String)
    mResult = vData
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

'-----------------------------------------------------------------------------'
'   기능 : ClientDb의 결과내역에 결과저장하는 쿼리조회
'   인수 :
'       - pItemSeq  : 전송순서
'       - pIntNm    : 화면표시 검사명
'       - pTestCd   : 검사코드
'       - pResult   : 결과
'       - pHLDiv    : H/L
'       - pDPDiv    : D/P
'-----------------------------------------------------------------------------'
Public Function GetCliRegisterSql(ByVal pItemSeq As String, ByVal pIntNm As String, _
                    ByVal pTestCd As String, ByVal pResult As String, ByVal pHLDiv As String, _
                    ByVal pDPDiv As String) As String
    Dim SQL As String
    
    SQL = mSQL.InsertClientResult(pItemSeq, pIntNm, pTestCd, pResult, pHLDiv, pDPDiv)
    GetCliRegisterSql = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : ClientDb의 QC결과내역에 결과저장하는 쿼리조회
'   인수 :
'       - pItemSeq  : 전송순서
'       - pIntNm    : 화면표시 검사명
'       - pTestCd   : 검사코드
'       - pResult   : 결과
'       - pHLDiv    : H/L
'       - pDPDiv    : D/P
'       - pRADiv    : QC 결과판정(A/R)
'       - pText     : QC Reject 소견
'-----------------------------------------------------------------------------'
Public Function GetCliQCRegisterSql(ByVal pItemSeq As String, ByVal pIntNm As String, _
                    ByVal pTestCd As String, ByVal pResult As String, ByVal pRADiv As String, _
                    ByVal pText As String) As String
    Dim SQL As String
    
    SQL = mSQL.InsertClientQCResult(pItemSeq, pIntNm, pTestCd, pResult, pRADiv, pText)
    GetCliQCRegisterSql = SQL
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISLibrarySql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub

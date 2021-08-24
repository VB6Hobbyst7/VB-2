VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------------------------'
'   파일명  : clsIISResults.cls (자사 LIS와 조인할때)
'   작성자  : 이상대
'   내  용  : 결과내역 컬렉션 클래스
'   작성일  : 2004-03-31
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mResults As Collection

Public Property Get Count() As Long
    Count = mResults.Count
End Property

Public Property Get Item(ByVal pTestCd As String) As clsIISResult
Attribute Item.VB_UserMemId = 0
    Set Item = mResults(pTestCd)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mResults.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   기능 : 결과내역 추가1 (인터페이스용)
'   인수 :
'       - pTestCd   : 검사코드
'       - pTestNm10 : 검사명(10자리)
'       - pTestNm   : 검사명 (전체)
'       - pUnit     : 결과단위
'       - pRerunFg  : 재검여부 (0:무, 1:유)
'       - pRstType  : 결과유형 (F:Free, N:Numeric, A:Alpha)
'       - pRstDiv   : Require/Alternative 여부 (R:Require, A:Alternative, *:상세모코드)
'       - pDetailFg : 상세항목순번
'       - pTSpc     : 지정검체 객체
'   반환 : 결과내역 객체
'-----------------------------------------------------------------------------'
Public Function AddInter(ByVal pTestCd As String, Optional ByVal pTestNm10 As String, _
                    Optional ByVal pTestNm As String, Optional ByVal pUnit As String, _
                    Optional ByVal pLastRst As String, Optional ByVal pRerunFg As String, _
                    Optional ByVal pRstType As String, Optional ByVal pRstDiv As String, _
                    Optional ByVal pDetailFg As String, Optional ByVal pTSpc As clsIISTSpc) As clsIISResult
    Dim objResult As clsIISResult
    
    Set objResult = New clsIISResult
    With objResult
        .TestCd = pTestCd
        .TestNm10 = pTestNm10
        .TestNm = pTestNm
        .Unit = pUnit
        .LastRst = pLastRst
        .RerunFg = pRerunFg
        .RstType = pRstType
        .RstDiv = pRstDiv
        .DetailFg = pDetailFg
        Set .TSpc = pTSpc
    End With
    
    mResults.Add objResult, pTestCd
    
    Set AddInter = objResult
    Set objResult = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 결과내역 추가2 (인터페이스용)
'   인수 :
'       - pResult : 결과내역 객체
'   반환 : 결과내역 객체
'-----------------------------------------------------------------------------'
Public Sub AddInterX(ByVal pResult As clsIISResult)
    mResults.Add pResult, pResult.TestCd
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에서 해당요소 삭제
'   인수 :
'       - pTestCd : 검사코드
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pTestCd As String)
    mResults.Remove pTestCd
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션의 모든 요소삭제
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mResults.Count To 1 Step -1
        mResults.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 Index의 존재유무를 반환
'   인수 :
'       - pTestCd : 검사코드
'   반환 : True(존재) / Flase(비존재)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String) As Boolean
    Dim objResult As clsIISResult
    
    For Each objResult In mResults
        With objResult
            If pTestCd = .TestCd Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next
    Exist = False

EndLine:
    Set objResult = Nothing
End Function

Private Sub Class_Initialize()
    Set mResults = New Collection
End Sub

Private Sub Class_Terminate()
    Set mResults = Nothing
End Sub


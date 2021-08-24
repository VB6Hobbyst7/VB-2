VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISTestCdFulls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISTestCdFulls.cls

'   내  용  : 검사코드(전체) 컬렉션 클래스

'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mFTestCds As Collection      '검사코드 컬렉션

Public Property Get Count() As Long
    Count = mFTestCds.Count
End Property

Public Property Get Item(ByVal pTestCd As String, ByVal pApplyDt As String) As clsIISTestCdFull
Attribute Item.VB_UserMemId = 0
    Set Item = mFTestCds(pTestCd & DIV & pApplyDt)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mFTestCds.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에 항목추가
'   반환 : 추가한 검사코드 객체
'   인수 :
'       1.pTestCd    : 검사코드
'       2.pApplyDt   : 적용일
'       3.pExpireDt  : 폐기일
'       4.pTestNm5   : 검사명(5자리)
'       5.pTestNm10  : 검사명(10자리)
'       6.pTestNm    : 검사명(전체)
'       7.pWorkarea  : Workarea 코드
'       8.pRstType   : 결과유형(Null: 일반, F: Free, N: Numeric, A: Alpha)
'       9.pRstDiv    : Alternative, Require 유형(A: Alternative, R: Require)
'       10.pPanelFg  : 처방구분(Null: 단일항목, G: 그룹항목, D: 상세항목)
'       11.pDetailFg : Detail 항목여부(Null: 없음, *:단일항목 자코드)
'       12.pRptSeq   : 출력순서
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pTestCd As String, ByVal pApplyDt As String, _
                    ByVal pWorkarea As String, ByVal pTestNm As String, _
                    Optional ByVal pExpireDt As String, Optional ByVal pTestNm5 As String, _
                    Optional ByVal pTestNm10 As String, Optional ByVal pRstType As String, _
                    Optional ByVal pRstDiv As String, Optional ByVal pPanelFg As String, _
                    Optional ByVal pDetailFg As String, Optional ByVal pRptSeq As Long _
                    ) As clsIISTestCdFull

    Dim objFTestCd As clsIISTestCdFull
    
    Set objFTestCd = New clsIISTestCdFull
    With objFTestCd
        .TestCd = pTestCd
        .Applydt = pApplyDt
        .ExpireDt = pExpireDt
        .TestNm5 = pTestNm5
        .TestNm10 = pTestNm10
        .TestNm = pTestNm
        .Workarea = pWorkarea
        .RstType = pRstType
        .RstDiv = pRstDiv
        .PanelFg = pPanelFg
        .DetailFg = pDetailFg
        .RptSeq = pRptSeq
    End With
    
    '## 키는 검사코드, 적용일
    mFTestCds.Add objFTestCd, pTestCd & DIV & pApplyDt
    
    Set Add = objFTestCd
    Set objFTestCd = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에서 해당요소 삭제
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pTestCd As String, ByVal pApplyDt As String)
    mFTestCds.Remove pTestCd & DIV & pApplyDt
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션의 모든 요소삭제
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mFTestCds.Count To 1 Step -1
        mFTestCds.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 Index의 존재유무를 반환
'   반환 : True(존재) / Flase(비존재)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String, ByVal pApplyDt As String) As Boolean
    Dim objFTestCd As clsIISTestCdFull
    
    For Each objFTestCd In mFTestCds
        If pTestCd = objFTestCd.TestCd And pApplyDt = objFTestCd.Applydt Then
            Exist = True
            GoTo EndLine
        End If
    Next

EndLine:
    Set objFTestCd = Nothing
End Function

Public Function AddTestCd(ByRef pTestCdFull As clsIISTestCdFull)

End Function

Private Sub Class_Initialize()
    Set mFTestCds = New Collection
End Sub

Private Sub Class_Terminate()
    Set mFTestCds = Nothing
End Sub


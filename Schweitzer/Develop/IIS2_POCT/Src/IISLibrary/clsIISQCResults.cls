VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISQCResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISQCResults.cls (자사 LIS와 조인할때)
'   작성자  : 이상대
'   내  용  : QC 결과내역 컬렉션 클래스
'   작성일  : 2004-05-21
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mQCResults As Collection

Public Property Get Count() As Long
    Count = mQCResults.Count
End Property

Public Property Get Item(ByVal pTestCd As String) As clsIISQCResult
Attribute Item.VB_UserMemId = 0
    Set Item = mQCResults(pTestCd)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mQCResults.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   기능 : QC결과내역 추가1 (인터페이스용)
'   인수 :
'       - pTestCd   : 검사코드
'       - pTestNm10 : 검사명(10자리)
'       - pTestNm   : 검사명(전체)
'       - pUnit     : 결과단위
'       - pRstDiv   : Require/Alternative 여부 (R:Require, A:Alternative, *:상세모코드)
'       - pCtrlCd   : Control Code
'       - pLevelCd  : Level Code
'       - pLotNo    : Lot No
'       - pQCMaster : QC Master 객체
'   반환 : QC결가내역 객체
'-----------------------------------------------------------------------------'
Public Function AddIIS(ByVal pTestCd As String, Optional ByVal pTestNm10 As String, _
                    Optional ByVal pTestNm As String, Optional ByVal pUnit As String, _
                    Optional ByVal pRstDiv As String, Optional ByVal pCtrlCd As String, _
                    Optional ByVal pLevelCd As String, Optional ByVal pLotNo As String, _
                    Optional ByVal pQCMaster As clsIISQCMaster) As clsIISQCResult
    Dim objQCResult As clsIISQCResult
    
    Set objQCResult = New clsIISQCResult
    With objQCResult
        .TestCd = pTestCd
        .TestNm10 = pTestNm10
        .TestNm = pTestNm
        .Unit = pUnit
        .RstDiv = pRstDiv
        .CtrlCd = pCtrlCd
        .LevelCd = pLevelCd
        .LotNo = pLotNo
        .QCMaster = pQCMaster
    End With
    
    mQCResults.Add objQCResult, pTestCd
    
    Set AddIIS = objQCResult
    Set objQCResult = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : QC결과내역 추가2 (인터페이스용)
'   인수 :
'       - pQCResult : QC결과내역 객체
'   반환 : QC결과내역 객체
'-----------------------------------------------------------------------------'
Public Sub AddIISX(ByVal pQCResult As clsIISQCResult)
    mQCResults.Add pQCResult, pQCResult.TestCd
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에서 해당요소 삭제
'   인수 :
'       - pSeq : 버퍼순서
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pTestCd As String)
    mQCResults.Remove pTestCd
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션의 모든 요소삭제
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mQCResults.Count To 1 Step -1
        mQCResults.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 Index의 존재유무를 반환
'   인수 :
'       - pSeq : 버퍼순서
'   반환 : True(존재) / Flase(비존재)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String) As Boolean
    Dim objQCResult As clsIISQCResult
    
    For Each objQCResult In mQCResults
        With objQCResult
            If pTestCd = .TestCd Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next
    Exist = False

EndLine:
    Set objQCResult = Nothing
End Function

Private Sub Class_Initialize()
    Set mQCResults = New Collection
End Sub

Private Sub Class_Terminate()
    Set mQCResults = Nothing
End Sub


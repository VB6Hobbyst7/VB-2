VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISIntResults.cls
'   작성자  : 오세원
'   내  용  : 인터페이스 결과 컬렉션 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mIntResults As Collection

Public Property Get Count() As Long
    Count = mIntResults.Count
End Property

Public Property Get Item(ByVal pIntBase As String) As clsIISIntResult
Attribute Item.VB_UserMemId = 0
    Set Item = mIntResults(pIntBase)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mIntResults.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스 결과추가1
'   인수 :
'       - pIntBase   : 장비기준 검사명
'       - pIntNm     : 화면표시 검사명
'       - pIntResult : 장비에서 전송한 결과
'       - pResult    : 결과
'       - pInfo      : 추가정보
'   반환 : 인터페이스 결과객체
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pIntBase As String, ByVal pIntNm As String, ByVal pIntResult As String, _
                    ByVal pResult As String, Optional ByVal pInfo As String) As clsIISIntResult
    Dim objIntResult As clsIISIntResult
    
    Set objIntResult = New clsIISIntResult
    With objIntResult
        .IntBase = pIntBase
        .IntNm = pIntNm
        .IntResult = pIntResult
        .Result = pIntResult 'pResult
        .Info = pInfo
    End With
    mIntResults.Add objIntResult, pIntBase
    
    Set Add = objIntResult
    Set objIntResult = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 인터페이스 결과추가2
'   인수 :
'       - pIntResult : 인터페이스 결과클래스
'   반환 : QC결과내역 객체
'-----------------------------------------------------------------------------'
Public Sub AddX(ByVal pIntResult As clsIISIntResult)
    mIntResults.Add pIntResult, pIntResult.IntBase
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에서 해당요소 삭제
'   인수 :
'       - pSeq : 버퍼순서
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pIntBase As String)
    mIntResults.Remove pIntBase
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션의 모든 요소삭제
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mIntResults.Count To 1 Step -1
        mIntResults.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 Index의 존재유무를 반환
'   인수 :
'       - pSeq : 버퍼순서
'   반환 : True(존재) / Flase(비존재)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pIntBase As String) As Boolean
    Dim objIntResult As clsIISIntResult
    
    For Each objIntResult In mIntResults
        With objIntResult
            If pIntBase = .IntBase Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next
    Exist = False

EndLine:
    Set objIntResult = Nothing
End Function

Private Sub Class_Initialize()
    Set mIntResults = New Collection
End Sub

Private Sub Class_Terminate()
    Set mIntResults = Nothing
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISTestCds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISTestCds.cls
'   작성자  : 이상대
'   내  용  : 검사코드 컬렉션 클래스
'   작성일  : 2004-01-19
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mTestCds As Collection

Public Property Get Count() As Long
    Count = mTestCds.Count
End Property

Public Property Get Item(ByVal pKey As Variant) As clsIISTestCd
Attribute Item.VB_UserMemId = 0
    Set Item = mTestCds(pKey)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mTestCds.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에 항목추가
'   인수 :
'       1.pTestCd : 검사코드
'       2.pTestSeq: Sequence
'       3.pTestNm : 검사명
'   반환 : 검사코드 객체
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pTestCd As String, ByVal pTestSeq As Long, _
                    ByVal pTestNm As String) As clsIISTestCd
    Dim objTestCd As clsIISTestCd
    
    Set objTestCd = New clsIISTestCd
    With objTestCd
        .TestCd = pTestCd
        .TestSeq = pTestSeq
        .TestNm = pTestNm
    End With
    
    mTestCds.Add objTestCd, pTestCd
    
    Set Add = objTestCd
    Set objTestCd = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에 항목추가(컬렉션에 키가 없는상태로 추가)
'   인수 :
'       1.pTestCd : 컬렉션에 추가할 검사코드 클래스
'       2.pBefore : 컬렉션에서 위치
'-----------------------------------------------------------------------------'
Public Sub AddEx(ByRef pTestCd As clsIISTestCd, Optional ByVal pBefore As Long = 0)
    If pBefore = 0 Then
        mTestCds.Add pTestCd
    Else
        mTestCds.Add pTestCd, , pBefore
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 전달된 컬렉션 클래스를 키를 이용해 입력
'   인수 :
'       1.pTestCds : 키가 없는 검사코드 컬렉션 클래스
'-----------------------------------------------------------------------------'
Public Sub SortAdd(ByRef pTestCds As clsIISTestCds)
    Dim objTestCds As clsIISTestCds
    Dim objTestCd  As clsIISTestCd
    Dim i          As Long

    i = 1
    Set objTestCds = New clsIISTestCds
    For Each objTestCd In pTestCds
        objTestCds.Add objTestCd.TestCd, i, objTestCd.TestNm
        i = i + 1
    Next

    Set pTestCds = objTestCds
    Set objTestCds = Nothing
    Set objTestCd = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에서 해당요소 삭제
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pKey As Variant)
    mTestCds.Remove pKey
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션의 모든 요소삭제
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mTestCds.Count To 1 Step -1
        mTestCds.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 Index의 존재유무를 반환
'   반환 : True(존재) / Flase(비존재)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String) As Boolean
    Dim objTestCd As clsIISTestCd
    
    For Each objTestCd In mTestCds
        If pTestCd = objTestCd.TestCd Then
            Exist = True
            GoTo EndLine
        End If
    Next
    Exist = False

EndLine:
    Set objTestCd = Nothing
End Function

Private Sub Class_Initialize()
    Set mTestCds = New Collection
End Sub

Private Sub Class_Terminate()
    Set mTestCds = Nothing
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntNms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISIntNms.cls
'   작성자  : 오세원
'   내  용  : 장비별 검사항목 컬렉션 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mIntNms As Collection

Public Property Get Count() As Long
    Count = mIntNms.Count
End Property

Public Property Get Item(ByVal pTestCd As String) As clsIISIntNm
Attribute Item.VB_UserMemId = 0
    Set Item = mIntNms(pTestCd)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mIntNms.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   기능 : 검사코드 추가 (인터페이스용)
'          검사코드의 장비기준 검사명을 조회하기 위해
'   인수 :
'       - pTestCd   : 검사코드
'       - pIntNm    : 화면표시 검사명
'       - pIntBase  : 장비기준 검사명
'       - pFrVal    : BorderLine From Value
'       - ptoVal    : BorderLine To Value
'   반환 : 검사항목 클래스
'-----------------------------------------------------------------------------'
Public Function AddIIS(ByVal pTestCd As String, ByVal pIntNm As String, ByVal pIntBase As String, _
                    Optional ByVal pFrVal As String, Optional ByVal pToVal As String) As clsIISIntNm
    Dim objIntNm As clsIISIntNm
    
    Set objIntNm = New clsIISIntNm
    With objIntNm
        .TestCd = pTestCd
        .IntNm = pIntNm
        .IntBase = pIntBase
        .FrVal = pFrVal
        .ToVal = pToVal
    End With
    
    mIntNms.Add objIntNm, pTestCd
    
    Set AddIIS = objIntNm
    Set objIntNm = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에서 해당요소 삭제
'   인수 :
'       - pTestCd   : 검사코드
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pTestCd As String)
    mIntNms.Remove pTestCd
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션의 모든 요소삭제
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mIntNms.Count To 1 Step -1
        mIntNms.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 Index의 존재유무를 반환
'   인수 :
'       - pTestCd   : 검사코드
'   반환 : True(존재) / Flase(비존재)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String) As Boolean
    Dim objIntNm As clsIISIntNm     '장비별 검사항목 클래스
    
    For Each objIntNm In mIntNms
        With objIntNm
            If pTestCd = .TestCd Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next

EndLine:
    Set objIntNm = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비기준 검사명을 이용해 검사명 조회
'   인수 :
'       - pIntBase : 장비기준 검사명
'   반환 : 검사명
'-----------------------------------------------------------------------------'
Public Function GetIntNm(ByVal pIntBase As String) As String
    Dim objIntNm As clsIISIntNm     '장비별 검사항목 클래스
    
    For Each objIntNm In mIntNms
        If objIntNm.IntBase = pIntBase Then
            GetIntNm = objIntNm.IntNm
            GoTo EndLine
        End If
    Next
    
EndLine:
    Set objIntNm = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명을 이용해 장비기준 검사명 조회
'   인수 :
'       - pIntNm : 검사명
'   반환 : 검사명
'-----------------------------------------------------------------------------'
Public Function GetIntBase(ByVal pIntNm As String) As String
    Dim objIntNm As clsIISIntNm     '장비별 검사항목 클래스
    
    For Each objIntNm In mIntNms
        If objIntNm.IntNm = pIntNm Then
            GetIntBase = objIntNm.IntBase
            GoTo EndLine
        End If
    Next
    
EndLine:
    Set objIntNm = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비기준 검사명이 컬렉션에 있는지 조회
'   인수 :
'       - pIntBase : 장비기준 검사명
'   반환 : True(존재), False(비존재)
'-----------------------------------------------------------------------------'
Public Function ExistIntBase(ByVal pIntBase As String) As Boolean
    Dim objIntNm As clsIISIntNm     '장비별 검사항목 클래스
    
    For Each objIntNm In mIntNms
        If objIntNm.IntBase = pIntBase Then
            ExistIntBase = True
            GoTo EndLine
        End If
    Next
    
EndLine:
    Set objIntNm = Nothing
End Function

Private Sub Class_Initialize()
    Set mIntNms = New Collection
End Sub

Private Sub Class_Terminate()
    Set mIntNms = Nothing
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISRefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISRefs.cls
'   작성자  : 오세원
'   내  용  : 참고치 컬렉션 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mRefs As Collection

Public Property Get Count() As Long
    Count = mRefs.Count
End Property

Public Property Get Item(ByVal pTestCd As String, ByVal pSpcCd As String, _
                         ByVal pSex As String, ByVal pAgeFr As Long, _
                         ByVal pAgeTo As Long, ByVal pApplyDt As String) As clsIISRef
Attribute Item.VB_UserMemId = 0
                         
    Set Item = mRefs(pTestCd & DIV & pSpcCd & DIV & pSex & DIV & CStr(pAgeFr) & DIV & _
                    CStr(pAgeTo) & DIV & pApplyDt)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mRefs.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에 항목추가
'   반환 : 지정검체 객체
'   인수 :
'       1.pTestCd   : 검사코드
'       2.pSpcCd    : 검체코드
'       3.pSex      : 성별
'       4.pAgeFr    : From Age
'       5.pAgeTo    : To Age
'       6.pApplyDt  : 적용일
'       7.pExpireDt : 폐기일
'       8.pRefFrVal : From Reference Value
'       9.pRefToVal : To Reference Value
'       10.pRefCd   : Alpha결과 참고치
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pTestCd As String, ByVal pSpcCd As String, _
                    ByVal pSex As String, ByVal pAgeFr As Long, _
                    ByVal pAgeTo As Long, ByVal pApplyDt As String, _
                    Optional ByVal pExpireDt As String, Optional ByVal pRefFrVal As Single, _
                    Optional ByVal pRefToVal As Single, Optional ByVal pRefCd As String) As clsIISRef

    Dim objTSpc As clsIISRef
    
    Set objTSpc = New clsIISRef
    With objTSpc
        .TestCd = pTestCd
        .SpcCd = pSpcCd
        .Sex = pSex
        .AgeFr = pAgeFr
        .AgeTo = pAgeTo
        .Applydt = pApplyDt
        .ExpireDt = pExpireDt
        .RefFrVal = pRefFrVal
        .RefToVal = pRefToVal
        .RefCd = pRefCd
    End With
    
    mRefs.Add objTSpc, pTestCd & DIV & pSpcCd & DIV & pSex & DIV & CStr(pAgeFr) & DIV & _
                CStr(pAgeTo) & DIV & pApplyDt
    
    Set Add = objTSpc
    Set objTSpc = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에서 해당요소 삭제
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pTestCd As String, ByVal pSpcCd As String, _
                  ByVal pSex As String, ByVal pAgeFr As Long, _
                  ByVal pAgeTo As Long, ByVal pApplyDt As String)
                         
    mRefs.Remove pTestCd & DIV & pSpcCd & DIV & pSex & DIV & CStr(pAgeFr) & DIV & _
                    CStr(pAgeTo) & DIV & pApplyDt
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션의 모든 요소삭제
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mRefs.Count To 1 Step -1
        mRefs.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 Index의 존재유무를 반환
'   반환 : True(존재) / Flase(비존재)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String, ByVal pSpcCd As String, _
                         ByVal pSex As String, ByVal pAgeFr As Long, _
                         ByVal pAgeTo As Long, ByVal pApplyDt As String) As Boolean
    Dim objTSpc As clsIISRef
    
    For Each objTSpc In mRefs
        With objTSpc
            If pTestCd = .TestCd And pSpcCd = .SpcCd And pSex = .Sex And pAgeFr = .AgeFr And _
               pAgeTo = .AgeTo And pApplyDt = .Applydt Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next
    Exist = False

EndLine:
    Set objTSpc = Nothing
End Function

Private Sub Class_Initialize()
    Set mRefs = New Collection
End Sub

Private Sub Class_Terminate()
    Set mRefs = Nothing
End Sub


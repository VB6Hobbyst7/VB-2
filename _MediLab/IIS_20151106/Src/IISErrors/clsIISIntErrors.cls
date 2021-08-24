VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISIntErrors.cls
'   작성자  : 오세원
'   내  용  : 인터페이스중 발생하는 에러저장 컬렉션 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mIntErrs As Collection

Public Property Get Count() As Long
    Count = mIntErrs.Count
End Property

Public Property Get Item(ByVal pSeq As Long) As clsIISIntError
Attribute Item.VB_UserMemId = 0
    Set Item = mIntErrs(CStr(pSeq))
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mIntErrs.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   기능 : 에러내용 추가(장비설정, 검체관련에러)
'   인수 :
'       - pCode  : 에러코드
'       - pEqpCd : 장비코드
'       - pEqpNm : 장비명
'       - pBarNo : 바코드번호
'       - pPtId  : 환자ID
'       - pName  : 환자명
'   반환 : 에레객체
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pCode As String, Optional pEqpCd As String, Optional pEqpNm As String, _
                    Optional ByVal pBarNo As String, Optional ByVal pPtId As String, _
                    Optional ByVal pName As String) As clsIISIntError
    Dim objIntErr As clsIISIntError
    Dim lngSeq    As Long       '에러순번
    
    lngSeq = mIntErrs.Count + 1
    Set objIntErr = New clsIISIntError
    With objIntErr
        .Seq = lngSeq
        .Code = pCode
        .ErrDt = Format$(Now, "YYYY-MM-DD HH:MM:SS")
        .EqpCd = pEqpCd
        .EqpNm = pEqpNm
        .BarNo = pBarNo
        .PtId = pPtId
        .Name = pName
    End With
    
    mIntErrs.Add objIntErr, CStr(lngSeq)
    
    Set Add = objIntErr
    Set objIntErr = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 에러내용 추가(장비에서 수신한 에러)
'   인수 :
'       - pCode  : 에러코드
'       - pEqpCd : 장비코드
'       - pEqpNm : 장비명
'       - pBarNo : 바코드번호
'       - pTitle : 에러내용
'-----------------------------------------------------------------------------'
Public Function AddX(ByVal pCode As String, ByVal pEqpCd As String, ByVal pEqpNm As String, _
                     Optional pBarNo As String, Optional ByVal pTitle As String) As clsIISIntError
    Dim objIntErr As clsIISIntError
    Dim lngSeq    As Long       '에러순번
    
    lngSeq = mIntErrs.Count + 1
    Set objIntErr = New clsIISIntError
    With objIntErr
        .Seq = lngSeq
        .Code = pCode
        .ErrDt = Format$(Now, "YYYY-MM-DD HH:MM:SS")
        .EqpCd = pEqpCd
        .EqpNm = pEqpNm
        .BarNo = pBarNo
        .Title = pTitle
    End With
    
    mIntErrs.Add objIntErr, CStr(lngSeq)
    
    Set AddX = objIntErr
    Set objIntErr = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에서 해당요소 삭제
'   인수 :
'       - pTestCd   : 검사코드
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pSeq As Long)
    mIntErrs.Remove CStr(pSeq)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션의 모든 요소삭제
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mIntErrs.Count To 1 Step -1
        mIntErrs.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 Index의 존재유무를 반환
'   인수 :
'       - pSeq : 에러순번
'   반환 : True(존재) / Flase(비존재)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pSeq As Long) As Boolean
    Dim objIntErr As clsIISIntError
    
    For Each objIntErr In mIntErrs
        With objIntErr
            If pSeq = .Seq Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next
    Exist = False

EndLine:
    Set objIntErr = Nothing
End Function

Private Sub Class_Initialize()
    Set mIntErrs = New Collection
End Sub

Private Sub Class_Terminate()
    Set mIntErrs = Nothing
End Sub


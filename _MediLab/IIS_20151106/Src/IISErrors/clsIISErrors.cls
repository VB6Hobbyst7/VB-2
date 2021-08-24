VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISErrors.cls
'   작성자  : 오세원
'   내  용  : Errors 컬렉션 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mErrors     As Collection       '에러 컬렉션

Public Property Get Count() As Long
    Count = mErrors.Count
End Property

Public Property Get Item(ByVal pIndex As Long) As clsIISError
Attribute Item.VB_UserMemId = 0
    Set Item = mErrors(pIndex)
End Property

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mErrors.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에 항목추가, 로그파일 오류내용 기록
'   반환 : 에러객체
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pFileNm As String, ByVal pClassNm As String, _
                    ByVal pProcedureNm As String, ByVal pDescription As String, _
                    ByVal pErrorTime As Date, Optional ByVal pKey As String) As clsIISError
    Dim objError As clsIISError
    
    Set objError = New clsIISError
    With objError
        .FileNm = pFileNm
        .ClassNm = pClassNm
        .ProcedureNm = pProcedureNm
        .Description = Replace(pDescription, vbLf, "")
        .ErrorTime = pErrorTime
        .SetLog pFileNm, pClassNm, pProcedureNm, pDescription, pErrorTime
    End With
    
    If Len(pKey) = 0 Then
        mErrors.Add objError
    Else
        mErrors.Add objError, pKey
    End If
    
    Set Add = objError
    Set objError = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에서 해당요소 삭제
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pIndex As Long)
    mErrors.Remove pIndex
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션의 모든 요소삭제
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i   As Long
    
    For i = mErrors.Count To 1 Step -1
        mErrors.Remove i
    Next i
End Sub

Private Sub Class_Initialize()
    Set mErrors = New Collection
End Sub

Private Sub Class_Terminate()
    Set mErrors = Nothing
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISSpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISSpc.cls
'   작성자  :
'   내  용  : 검체 클래스
'   작성일  :
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISMasterSql

Private mSpcCd      As String       '검체코드
Private mSpcNm      As String       '검체명(약어)
Private mSpcFullNm  As String       '검체명(전체)
Private mSpcBarNm   As String       '검체명(바코드출력)

Public Property Get SpcCd() As String
    SpcCd = mSpcCd
End Property

Public Property Let SpcCd(ByVal vData As String)
    mSpcCd = vData
End Property

Public Property Get SpcNm() As String
    SpcNm = mSpcNm
End Property

Public Property Let SpcNm(ByVal vData As String)
    mSpcNm = vData
End Property

Public Property Get SpcFullNm() As String
    SpcFullNm = mSpcFullNm
End Property

Public Property Let SpcFullNm(ByVal vData As String)
    mSpcFullNm = vData
End Property

Public Property Get SpcBarNm() As String
    SpcBarNm = mSpcBarNm
End Property

Public Property Let SpcBarNm(ByVal vData As String)
    mSpcBarNm = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 검체추가
'   인수 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function AddSpcCd() As Boolean
    Dim SQL As String
    
    SQL = mSQL.InsertSpcCd(mSpcCd, mSpcNm, mSpcFullNm, mSpcBarNm)
    
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddSpcCd = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISSpc", "AddSpcCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검체수정
'   인수 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function ModifySpcCd() As Boolean
    Dim SQL As String
    
    SQL = mSQL.UpdateSpcCd(mSpcCd, mSpcNm, mSpcFullNm, mSpcBarNm)
    
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifySpcCd = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISSpc", "ModifySpcCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검체삭제
'   인수 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function DelSpcCd() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteSpcCd(mSpcCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelSpcCd = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISSpc", "DelSpcCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검체 리스트 조회
'   반환 : 검체 리스트 Recordset
'-----------------------------------------------------------------------------'
Public Function GetSpcCd() As ADODB.Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectSpcCd, , adCmdText)
    
    Set GetSpcCd = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISSpc", "GetSpcCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검체코드에 대한 검체명 조회
'   인수 :
'       1.pSpcCd : 컴체코드
'   반환 :
'       1.해당 검체코드가 존재하면 검체명 반환
'       2.존재하지 않으면 Null 반환
'-----------------------------------------------------------------------------'
Public Function GetSpcNm(ByVal pSpcCd As String) As String
    Dim Rs  As ADODB.Recordset

On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectSpcNm(pSpcCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    GetSpcNm = Rs.Fields("SPCNM").Value & ""
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISSpc", "GetSpcNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub

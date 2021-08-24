VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISWorkarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISWorkarea.cls

'   내  용  : WorkArea 클래스
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISMasterSql

Private mWaCd       As String       'Workarea 코드
Private mWaEngNm    As String       'Workarea 이름(영어)
Private mWaKorNm    As String       'Workarea 이름(한글)
Private mWaNumRule  As String       '접수번호 부여기준

Public Property Get WaCd() As String
    WaCd = mWaCd
End Property

Public Property Let WaCd(ByVal vData As String)
    mWaCd = vData
End Property

Public Property Get WaEngNm() As String
    WaEngNm = mWaEngNm
End Property

Public Property Let WaEngNm(ByVal vData As String)
    mWaEngNm = vData
End Property

Public Property Get WaKorNm() As String
    WaKorNm = mWaKorNm
End Property

Public Property Let WaKorNm(ByVal vData As String)
    mWaKorNm = vData
End Property

Public Property Get WaNumRule() As String
    WaNumRule = mWaNumRule
End Property

Public Property Let WaNumRule(ByVal vData As String)
    mWaNumRule = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : WorkArea 추가
'   인수 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function AddWorkarea() As Boolean
    Dim SQL As String
    
    SQL = mSQL.InsertWorkarea(mWaCd, mWaEngNm, mWaKorNm, mWaNumRule)
    
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddWorkarea = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISWorkarea", "AddWorkarea", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : WorkArea 수정
'   인수 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function ModifyWorkarea() As Boolean
    Dim SQL As String
    
    SQL = mSQL.UpdateWorkarea(mWaCd, mWaEngNm, mWaKorNm, mWaNumRule)
    
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyWorkarea = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISWorkarea", "ModifyWorkarea", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : WorkArea 삭제
'   인수 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function DelWorkarea() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteWorkarea(mWaCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelWorkarea = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISWorkarea", "DelWorkarea", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 모든 WorkArea 리스트 조회
'-----------------------------------------------------------------------------'
Public Function GetWorkarea() As ADODB.Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectWorkarea, , adCmdText)
    Set GetWorkarea = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISWorkarea", "GetWorkarea", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub

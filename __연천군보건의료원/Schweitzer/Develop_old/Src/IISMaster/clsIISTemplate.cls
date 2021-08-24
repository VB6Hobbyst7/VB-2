VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISTemplate.cls

'   내  용  : 템플릿 마스터 클래스
'   버  전  :
'-----------------------------------------------------------------------------'

'## 속성
'   - 템플릿 인덱스
'   - 템플릿 코드
'   - 템플릿 내용

'## 메소드
'   - 템플릿 입력
'   - 템플릿 수정
'   - 템플릿 삭제
'   - 템플릿 조회

Option Explicit

Private mSQL        As clsIISMasterSql

Private mCdIndex    As String   '템플릿 인덱스
Private mCode       As String   '템플릿 코드
Private mContent    As String   '템플릿 내용

Public Property Get CdIndex() As String
    CdIndex = mCdIndex
End Property

Public Property Let CdIndex(ByVal vData As String)
    mCdIndex = vData
End Property

Public Property Get Code() As String
    Code = mCode
End Property

Public Property Let Code(ByVal vData As String)
    mCode = vData
End Property

Public Property Get Content() As String
    Content = mContent
End Property

Public Property Let Content(ByVal vData As String)
    mContent = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 템플릿 입력
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function AddTemplate() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.InsertTemplate(mCdIndex, mCode, mContent), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddTemplate = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISTemlate", "AddTemplate", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 템플릿 수정
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function ModifyTemplate() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.UpdateTemplate(mCdIndex, mCode, mContent), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyTemplate = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISTemlate", "ModifyTemplate", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 템플릿 삭제
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function DelTemplate() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteTemplate(mCdIndex, mCode), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelTemplate = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISTemlate", "DelTemplate", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 템플릿 조회
'   인수 :
'       1.pCdIndex : 템플릿 인덱스
'   반환 : 템플릿 Recordset
'-----------------------------------------------------------------------------'
Public Function GetTemplates(ByVal pCdIndex As String) As ADODB.Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectTemplate(pCdIndex), , adCmdText)
    Set GetTemplates = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISTemlate", "GetTemplates", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub

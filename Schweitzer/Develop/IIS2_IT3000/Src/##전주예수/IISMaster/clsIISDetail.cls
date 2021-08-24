VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISDetail.cls
'   작성자  : 이상대
'   내  용  : 상세항목 설정 클래스
'   작성일  : 2004-02-20
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISMasterSql

Private mParentCd   As String       '모코드
Private mParentCdNm As String       '모코드명
Private mChildCd    As String       '자코드
Private mChildCdNm  As String       '자코드명

Public Property Get ParentCd() As String
    ParentCd = mParentCd
End Property

Public Property Let ParentCd(ByVal vData As String)
    mParentCd = vData
End Property

Public Property Get ParentCdNm() As String
    ParentCdNm = mParentCdNm
End Property

Public Property Let ParentCdNm(ByVal vData As String)
    mParentCdNm = vData
End Property

Public Property Get ChildCd() As String
    ChildCd = mChildCd
End Property

Public Property Let ChildCd(ByVal vData As String)
    mChildCd = vData
End Property

Public Property Get ChildCdNm() As String
    ChildCdNm = mChildCdNm
End Property

Public Property Let ChildCdNm(ByVal vData As String)
    mChildCdNm = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 상세항목 추가
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function AddDetail() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.InsertDetailCd(mParentCd, mChildCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddDetail = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISDetail", "AddDetail", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 상세항목 삭제
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function DelDetail() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteDetailCd(mParentCd, mChildCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelDetail = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISDetail", "DelDetail", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 상세항목의 모코드에 등록된 자코드를 모두삭제
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function DelDetailAll() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteDetailCdall(mParentCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelDetailAll = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISDetail", "DelDetailAll", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 모코드 리스트 조회
'   반환 : 모코드 리스트 Recordset
'-----------------------------------------------------------------------------'
Public Function GetParentList() As ADODB.Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectDetailParentCd, , adCmdText)
    Set GetParentList = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISDetail", "GetParentList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 자코드 리스트 조회
'   인수 :
'       1.pParentCd : 모코드
'   반환 : 자코드 리스트 Recordset
'-----------------------------------------------------------------------------'
Public Function GetChildList(ByVal pParentCd As String) As ADODB.Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectDetailChildCd(pParentCd), , adCmdText)
    Set GetChildList = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISDetail", "GetChildList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 모코드를 조회하여 존재하면 코드명 반환
'   인수 :
'       1.pParentCd : 모코드
'   반환 :
'       1.해당 모코드가 존재하면 코드명 반환
'       2.존재하지 않으면 Null반환
'-----------------------------------------------------------------------------'
Public Function GetParentNm(ByVal pParentCd As String) As String
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectDetailParentNm(pParentCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine

    GetParentNm = Rs.Fields("TESTNM").Value
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISDetail", "GetParentNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 자코드를 조회하여 존재하면 코드명 반환
'   인수 :
'       1.pChildCd : 모코드
'   반환 :
'       1.해당 모코드가 존재하면 코드명 반환
'       2.존재하지 않으면 Null반환
'-----------------------------------------------------------------------------'
Public Function GetChildNm(ByVal pChildCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectDatailChildNm(pChildCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine

    GetChildNm = Rs.Fields("TESTNM").Value
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISDetail", "GetChildNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : TIIS101에서 모코드 리스트를 조회
'   반환 : 조회 쿼리
'-----------------------------------------------------------------------------'
Public Function GetParentListSql() As String
    GetParentListSql = mSQL.SelectDetailParentCdX
End Function

'-----------------------------------------------------------------------------'
'   기능 : TIIS101에서 자코드 리스트를 조회
'   반환 : 조회 쿼리
'-----------------------------------------------------------------------------'
Public Function GetChildListSql() As String
    GetChildListSql = mSQL.SelectDetailChildCdX
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub

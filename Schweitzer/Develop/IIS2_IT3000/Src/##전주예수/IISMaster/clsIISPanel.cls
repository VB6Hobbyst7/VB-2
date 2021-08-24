VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISPanel.cls (자사 LIS와 조인할때)
'   작성자  : 이상대
'   내  용  : 그룹처방 설정 클래스
'   작성일  : 2004-02-21
'   버  전  :
'-----------------------------------------------------------------------------'

'## 속성
'   - 그룹코드
'   - 그룹코드명
'   - 그룹코드 SEQ
'   - 상세코드
'   - 상세코드명

'## 메소드
'   - 그룹코드 저장
'   - 그룹코드 수정
'   - 그룹코드 삭제
'   - 그룹코드 하위항목 모두 삭제
'   - 그룹코드 리스트 조회
'   - 상세항목 조회
'   - 그룹코드의 이름조회
'   - 상세코드의 이름조회
'   - 검사코드 마스터에서 그룹코드만 조회
'   - 검사코드 마스터에서 단일항목, 상세엄마만 조회

Option Explicit

Private mSQL        As clsIISMasterSql

Private mParentCd   As String       '그룹코드
Private mParentNm   As String       '그룹코드명
Private mSeq        As String       '그룹코드 SEQ
Private mChildCd    As String       '상세코드
Private mChildNm    As String       '상세코드명

Public Property Get ParentCd() As String
    ParentCd = mParentCd
End Property

Public Property Let ParentCd(ByVal vData As String)
    mParentCd = vData
End Property

Public Property Get ParentNm() As String
    ParentNm = mParentNm
End Property

Public Property Let ParentNm(ByVal vData As String)
    mParentNm = vData
End Property

Public Property Get Seq() As String
    Seq = mSeq
End Property

Public Property Let Seq(ByVal vData As String)
    mSeq = vData
End Property

Public Property Get ChildCd() As String
    ChildCd = mChildCd
End Property

Public Property Let ChildCd(ByVal vData As String)
    mChildCd = vData
End Property

Public Property Get ChildNm() As String
    ChildNm = mChildNm
End Property

Public Property Let ChildNm(ByVal vData As String)
    ChildNm = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 그룹코드 추가
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function AddPanel() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.InsertPanelCd(mParentCd, mSeq, mChildCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddPanel = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISPanel", "AddPanel", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 그룹코드 수정
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function ModifyPanel() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.UpdatePanelCd(mParentCd, mSeq, mChildCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyPanel = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISPanel", "ModifyPanel", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 그룹코드 삭제
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function DelPanel() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeletePanelCd(mParentCd, mSeq), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelPanel = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISPanel", "DelPanel", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 그룹코드 삭제
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function DelPanelAll() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeletePanelCdAll(mParentCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelPanelAll = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISPanel", "DelPanelAll", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : Panel 대표코드 리스트 조회
'   반환 : Panel 대표코드 Recordset
'-----------------------------------------------------------------------------'
Public Function GetPanelList() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectPanelList, , adCmdText)
    Set GetPanelList = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISPanel", "GetPanelList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : Panel 상세코드 리스트 조회
'   반환 : Panel 상세코드 Recordset
'-----------------------------------------------------------------------------'
Public Function GetPanelChildList(ByVal pParentCd As String) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectPanelChildList(pParentCd), , adCmdText)
    Set GetPanelChildList = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISPanel", "GetPanelChildList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 그룹코드의 코드명 조회
'   인수 :
'       1.pParentCd : 그룹코드
'   반환 :
'       1.해당 그룹코드가 존재하면 코드명 반환
'       2.존재하지 않으면 Null반환
'-----------------------------------------------------------------------------'
Public Function GetPanelNm(ByVal pParentCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectPanelNm(pParentCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
    GetPanelNm = Rs.Fields("TESTNM").Value
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISPanel", "GetPanelNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 상세코드의 코드명 조회
'   인수 :
'       1.pChildCd : 상세코드
'   반환 :
'       1.해당 상세코드가 존재하면 코드명 반환
'       2.존재하지 않으면 Null반환
'-----------------------------------------------------------------------------'
Public Function GetChildNm(ByVal pChildCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectPanelChildNm(pChildCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
    GetChildNm = Rs.Fields("TESTNM").Value
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISPanel", "GetChildNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : TIIS101에서 그룹코드 리스트를 조회
'   반환 : 조회 쿼리
'-----------------------------------------------------------------------------'
Public Function GetPanelListBySql() As String
    GetPanelListBySql = mSQL.SelectPanelParentCd
End Function

'-----------------------------------------------------------------------------'
'   기능 : TIIS101에서 단일항목, 상세모코드 리스트를 조회
'   반환 : 조회 쿼리
'-----------------------------------------------------------------------------'
Public Function GetChildListBySql() As String
    GetChildListBySql = mSQL.SelectPanelChildCd
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub

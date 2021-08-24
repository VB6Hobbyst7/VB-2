VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISRepeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISRepeat.cls (자사 LIS와 조인할때)
'   작성자  :
'   내  용  : 다빈도처방 설정 클래스
'   작성일  :
'   버  전  :
'-----------------------------------------------------------------------------'

'## 속성
'   - 다빈도처방 코드
'   - 다빈도처방 코드명
'   - 검체코드
'   - 검체명

'## 메소드
'   - 다빈도처방 추가
'   - 다빈도처방 수정
'   - 다빈도처방 삭제
'   - 모든 다빈도처방 조회
'   - 다빈도처방 코드에 대한 이름 조회(검사항목은 그룹코드, 상세모코드 이어야 한다.)
'   - 해당 검체에 대한 이름 조회(검체코드는 다빈도처방 코드에 등록된 검체이어야 한다.)
'   - TIIS101에서 그룹코드, 상세모코드 조회
'   - 해당 다빈도처방 코드에 등록된 모든검체 조회

Option Explicit

Private mSQL        As clsIISMasterSql

Private mRepeatCd   As String       '다빈도처방 코드
Private mRepeatNm   As String       '다빈도처방 코드명
Private mSpcCd      As String       '검체코드
Private mSpcNm      As String       '검체코드명

Public Property Get RepeatCd() As String
    RepeatCd = mRepeatCd
End Property

Public Property Let RepeatCd(ByVal vData As String)
    mRepeatCd = vData
End Property

Public Property Get RepeatNm() As String
    RepeatNm = mRepeatNm
End Property

Public Property Let RepeatNm(ByVal vData As String)
    mRepeatNm = vData
End Property

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

'-----------------------------------------------------------------------------'
'   기능 : 다빈도처방 코드 추가
'   인수 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function AddRepeatCd() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.InsertRepeatCd(mRepeatCd, mSpcCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddRepeatCd = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISRepeat", "AddRepeatCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 다빈도처방 코드 수정
'   인수 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function ModifyRepeatCd() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.UpdateRepeatCd(mRepeatCd, mSpcCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyRepeatCd = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISRepeat", "ModifyRepeatCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 다빈도처방 코드 삭제
'   인수 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function DelRepeatCd() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteRepeatCd(mRepeatCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelRepeatCd = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISRepeat", "ModifyRepeatCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 다빈도처방 리스트 조회
'   반환 : Recordset
'-----------------------------------------------------------------------------'
Public Function GetRepeatCd() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectRepeatCd, , adCmdText)
    Set GetRepeatCd = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISRepeat", "GetRepeatCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 다빈도처방 코드명 조회
'   인수 :
'       1.pRepeatCd : 다빈도처방 코드
'   반환 :
'       1.해당 다빈도처방 코드가 존재하면 코드명 반환
'       2.존재하지 않으면 Null반환
'-----------------------------------------------------------------------------'
Public Function GetRepeatNm(ByVal pRepeatCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectRepeatNm(pRepeatCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
    GetRepeatNm = Rs.Fields("TESTNM").Value
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISRepeat", "GetRepeatNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검체 코드명 조회 (다빈도처방 코드에 등록된 검체코드중에서 조회)
'   인수 :
'       1.pRepeatCd : 다빈도처방 코드
'       2.pSpcCd    : 검체코드
'   반환 :
'       1.해당 검체코드가 존재하면 코드명 반환
'       2.존재하지 않으면 Null반환
'-----------------------------------------------------------------------------'
Public Function GetSpcNm(ByVal pRepeatCd As String, ByVal pSpcCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectSpcNmX(pRepeatCd, pSpcCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
    GetSpcNm = Rs.Fields("SPCNM").Value
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISRepeat", "GetSpcNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : TIIS101에서 그룹코드, 상세모코드만 조회
'-----------------------------------------------------------------------------'
Public Function GetRepeatCdBySql() As String
    GetRepeatCdBySql = mSQL.SelectRepeatCdx
End Function

'-----------------------------------------------------------------------------'
'   기능 : 모든 검체코드 조회
'   인수 :
'       1.pRepeatCd : 다빈도처방 코드
'-----------------------------------------------------------------------------'
Public Function GetSpcCdBySql(ByVal pRepeatCd As String) As String
    GetSpcCdBySql = mSQL.SelectSpcCdX(pRepeatCd)
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub

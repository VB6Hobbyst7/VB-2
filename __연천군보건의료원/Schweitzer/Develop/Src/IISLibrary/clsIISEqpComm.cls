VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISEqpComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISEqpComm.cls (자사 LIS와 조인할때)

'   내  용  : 검사장비 통신설정 클래스

'   버  전  :
'-----------------------------------------------------------------------------'

'## 속성
'   - 장비코드
'   - 포트
'   - Baud Rate
'   - Data bit
'   - Stop bit
'   - Parity bit
'   - 보관일

'## 메소드
'   - 통신설정 입력
'   - 통신설정 수정
'   - 통신설정 삭제
'   - 사용중인 장비의 통신설정 리스트 조회
'   - 사용중인 장비코드, 장비명 조회
'   - 해당장비가 등록된 장비인지 조회

Option Explicit

Private mSQL        As clsIISMasterSql

Private mEqpCd      As String       '장비코드
Private mPort       As String       'Port
Private mBaudRate   As String       'Baud Rate
Private mDatabit    As String       'Data bit
Private mStopbit    As String       'Stop bit
Private mParitybit  As String       'Parity bit
Private mStoredDt   As String       '보관일자

Public Property Get EqpCd() As String
    EqpCd = mEqpCd
End Property

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Get Port() As String
    Port = mPort
End Property

Public Property Let Port(ByVal vData As String)
    mPort = vData
End Property

Public Property Get BaudRate() As String
    BaudRate = mBaudRate
End Property

Public Property Let BaudRate(ByVal vData As String)
    mBaudRate = vData
End Property

Public Property Get Databit() As String
    Databit = mDatabit
End Property

Public Property Let Databit(ByVal vData As String)
    mDatabit = vData
End Property

Public Property Get Stopbit() As String
    Stopbit = mStopbit
End Property

Public Property Let Stopbit(ByVal vData As String)
    mStopbit = vData
End Property

Public Property Get Paritybit() As String
    Paritybit = mParitybit
End Property

Public Property Let Paritybit(ByVal vData As String)
    mParitybit = vData
End Property

Public Property Get StoredDt() As String
    StoredDt = mStoredDt
End Property

Public Property Let StoredDt(ByVal vData As String)
    mStoredDt = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 장비 통신설정 입력
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function AddEqpComm() As Boolean
    Dim SQL As String
    
    SQL = mSQL.InsertEqpComm(mEqpCd, mPort, mBaudRate, mDatabit, mStopbit, mParitybit, mStoredDt)
                
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddEqpComm = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISEqpComm", "AddEqpComm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비 통신설정 수정
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function ModifyEqpComm() As Boolean
    Dim SQL As String
    
    SQL = mSQL.UpdateEqpComm(mEqpCd, mPort, mBaudRate, mDatabit, mStopbit, mParitybit, mStoredDt)
                
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyEqpComm = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISEqpComm", "ModifyEqpComm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비 통신설정 삭제
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function DelEqpComm() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteEqpComm(mEqpCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelEqpComm = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISEqpComm", "DelEqpComm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : MS Comm 컨트롤의 Settings 프로퍼티에 입력할수 문자열 조회
'   반환 : MS Comm 컨트롤의 Settings 프로퍼티에 입력할수 문자열
'-----------------------------------------------------------------------------'
Public Function GetSettings() As String
    GetSettings = mBaudRate & "," & Left$(mParitybit, 1) & "," & mDatabit & "," & mStopbit
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당장비의 통신설정 상태를 조회
'   인수 :
'       1.pEqpCd : 장비코드
'-----------------------------------------------------------------------------'
Public Sub GetEqpComm(ByVal pEqpCd As String)
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectEqpComm(pEqpCd), , adCmdText)
    If Not (Rs.BOF Or Rs.EOF) Then
        mEqpCd = Rs.Fields("EQPCD").Value
        mPort = Rs.Fields("PORT").Value & ""
        mBaudRate = Rs.Fields("BAUDRATE").Value & ""
        mDatabit = Rs.Fields("DATABIT").Value & ""
        mStopbit = Rs.Fields("STOPBIT").Value & ""
        mParitybit = Rs.Fields("PARITYBIT").Value & ""
        mStoredDt = Rs.Fields("STOREDDT").Value & ""
    End If
    
    Rs.Close
    Set Rs = Nothing
    Exit Sub
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISEqpComm", "GetEqpComm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 사용중인장비의 통신설정정보 조회
'   반환 : Recordset
'-----------------------------------------------------------------------------'
Public Function GetEqpComms() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectEqpComms, , adCmdText)
    Set GetEqpComms = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISEqpComm", "GetEqpComms", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 사용중인 장비코드, 장비명 조회
'   반환 : Recordset
'-----------------------------------------------------------------------------'
Public Function GetUsingEqp() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectUsingEqp, , adCmdText)
    Set GetUsingEqp = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISEqpComm", "GetUsingEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 장비코드의 장비명 조회
'   인수 :
'       1.pEqpCd : 장비코드
'   반환 : 존재하는 장비이면 장비명 반환, 존재하지 않으면 Null반환
'-----------------------------------------------------------------------------'
Public Function GetEqpNm(ByVal pEqpCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectEqpCd(pEqpCd), , adCmdText)
    
    If Not (Rs.BOF Or Rs.EOF) Then
        GetEqpNm = Rs.Fields("EQPNM").Value
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISEqpComm", "GetEqpNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub

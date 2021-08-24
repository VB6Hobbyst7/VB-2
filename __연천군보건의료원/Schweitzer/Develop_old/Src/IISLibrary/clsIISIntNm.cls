VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntNm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISIntNm.cls

'   내  용  : 장비별 검사항목 클래스

'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISMasterSql

Private mEqpCd      As String   '장비코드
Private mIntNm      As String   '검사명
Private mIntBase    As String   '장비기준 검사명
Private mTestCd     As String   '검사코드
Private mFrVal      As String   'BorderLine From Value
Private mToVal      As String   'BorderLine To Value

'   - 기존병원의 S2Lab702 테이블 기준에 맞추기 위해 멤버변수 추가
Private mEqpSeq     As Long     '장비순서
Private mPrtOrd     As Long     '출력순서

Public Property Get EqpCd() As String
    EqpCd = mEqpCd
End Property

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Get IntNm() As String
    IntNm = mIntNm
End Property

Public Property Let IntNm(ByVal vData As String)
    mIntNm = vData
End Property

Public Property Get IntBase() As String
    IntBase = mIntBase
End Property

Public Property Let IntBase(ByVal vData As String)
    mIntBase = vData
End Property

Public Property Get TestCd() As String
    TestCd = mTestCd
End Property

Public Property Let TestCd(ByVal vData As String)
    mTestCd = vData
End Property

Public Property Get FrVal() As String
    FrVal = mFrVal
End Property

Public Property Let FrVal(ByVal vData As String)
    mFrVal = vData
End Property

Public Property Get ToVal() As String
    ToVal = mToVal
End Property

Public Property Let ToVal(ByVal vData As String)
    mToVal = vData
End Property

'## 1.1.6: 이상대(2005-05-11)
'   - 기존병원의 S2Lab702 테이블 기준에 맞추기 위해 프로퍼티 추가
Public Property Get EqpSeq() As Long
    EqpSeq = mEqpSeq
End Property

Public Property Let EqpSeq(ByVal vData As Long)
    mEqpSeq = vData
End Property

Public Property Get PrtOrd() As Long
    PrtOrd = mPrtOrd
End Property

Public Property Let PrtOrd(ByVal vData As Long)
    mPrtOrd = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 검사명 입력
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function AddIntNm() As Boolean
    Dim SQL As String
    
On Error GoTo Errors
    DbCon.BeginTrans
    SQL = mSQL.InsertIntNm(mEqpCd, mIntNm, mIntBase, mFrVal, mToVal)
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddIntNm = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "AddIntNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명 입력(기존병원용) - 1.1.6: 이상대(2005-05-11)
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function AddIntNmX() As Boolean
    Dim SQL As String
    
On Error GoTo Errors
    DbCon.BeginTrans
    SQL = mSQL.InsertIntNmX(mEqpCd, mEqpSeq, mIntBase, mIntNm, mPrtOrd, mFrVal, mToVal)
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddIntNmX = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "AddIntNmX", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명 수정
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function ModifyIntNm() As Boolean
    Dim SQL As String
    
On Error GoTo Errors
    DbCon.BeginTrans
    SQL = mSQL.UpdateIntNm(mEqpCd, mIntNm, mIntBase, mFrVal, mToVal)
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyIntNm = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "ModifyIntNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명 수정(기존병원용) - 1.1.6: 이상대(2005-05-11)
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function ModifyIntNmX() As Boolean
    Dim SQL As String
    
On Error GoTo Errors
    DbCon.BeginTrans
    SQL = mSQL.UpdateIntNmX(mEqpCd, mEqpSeq, mIntBase, mIntNm, mPrtOrd, mFrVal, mToVal)
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyIntNmX = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "ModifyIntNmX", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명 삭제
'          해당 검사명에 등록된 모든 검사코드 삭제후 검사명 삭제
'   인수 :
'       - pEqpCd : 장비코드
'       - pIntNm : 검사명
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function DelIntNm(ByVal pEqpCd As String, ByVal pIntNm As String) As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    
    DbCon.Execute mSQL.DeleteIntTestCdAll(pEqpCd, pIntNm), , adCmdText + adExecuteNoRecords
    DbCon.Execute mSQL.DeleteIntNm(pEqpCd, pIntNm), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelIntNm = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "DelIntNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명 삭제(기존병원용) - 1.1.6: 이상대(2005-05-11)
'          해당 검사명에 등록된 모든 검사코드 삭제후 검사명 삭제
'   인수 :
'       - pEqpCd  : 장비코드
'       - pEqpSeq : 장비기준 검사명 순서
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function DelIntNmX(ByVal pEqpCd As String, ByVal pEqpSeq As Long) As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    
    DbCon.Execute mSQL.DeleteIntTestCdAllX(pEqpCd, pEqpSeq), , adCmdText + adExecuteNoRecords
    DbCon.Execute mSQL.DeleteIntNmX(pEqpCd, pEqpSeq), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelIntNmX = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "DelIntNmX", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 장비의 모든 검사명 조회
'   인수 :
'       - pEqpCd : 장비코드
'   반환 : Recordset
'-----------------------------------------------------------------------------'
Public Function GetIntNms(ByVal pEqpCd As String) As ADODB.Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectIntNms(pEqpCd), , adCmdText)
    Set GetIntNms = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISIntNm", "GetIntNms", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명에 검사코드 입력 (무조건 삭제후 입력)
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function AddTestCd() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteIntTestCd(mEqpCd, mIntNm, mTestCd), , adCmdText + adExecuteNoRecords
    DbCon.Execute mSQL.InsertIntTestCd(mEqpCd, mIntNm, mTestCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddTestCd = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "AddTestCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명에 검사코드 입력 (무조건 삭제후 입력, 기존병원용)
'          - 1.1.6: 이상대(2005-05-11)
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function AddTestCdX() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteIntTestCdX(mEqpCd, mEqpSeq, mTestCd), , adCmdText + adExecuteNoRecords
    DbCon.Execute mSQL.InsertIntTestCdX(mEqpCd, mEqpSeq, mTestCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddTestCdX = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "AddTestCdX", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명에 검사코드 삭제
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function DelTestCd() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteIntTestCd(mEqpCd, mIntNm, mTestCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelTestCd = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "DelTestCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명에 검사코드 삭제(기존병원용) - 1.1.6: 이상대(2005-05-11)
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function DelTestCdX() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteIntTestCdX(mEqpCd, mEqpSeq, mTestCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelTestCdX = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "DelTestCdX", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명에 등록된 모든 검사코드 조회
'   인수 :
'       - pEqpCd  : 장비코드
'       - pIntNm  : 검사명
'   반환 : Recordset
'-----------------------------------------------------------------------------'
Public Function GetTestCds(ByVal pEqpCd As String, ByVal pIntNm As String) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectIntTestCds(pEqpCd, pIntNm), , adCmdText)
    Set GetTestCds = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISIntNm", "GetTestCds", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사명에 등록된 모든 검사코드 조회(기존병원용) - 1.1.6: 이상대(2005-05-11)
'   인수 :
'       - pEqpCd  : 장비코드
'       - pEqpSeq : 장비기준 검사명 순서
'   반환 : Recordset
'-----------------------------------------------------------------------------'
Public Function GetTestCdsX(ByVal pEqpCd As String, ByVal pEqpSeq As Long) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectIntTestCdsX(pEqpCd, pEqpSeq), , adCmdText)
    Set GetTestCdsX = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISIntNm", "GetTestCdsX", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사코드를 기준으로 검사명 조회
'   인수 :
'       - pEqpCd  : 장비코드
'   반환 : Recordset
'-----------------------------------------------------------------------------'
Public Function GetTestCdIntNms(ByVal pEqpCd As String) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectTestCdIntNms(pEqpCd), , adCmdText)
    Set GetTestCdIntNms = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISIntNm", "GetTestCdIntNms", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 장비코드의 장비명 조회
'   인수 :
'       - pEqpCd : 장비코드
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
    Error.SetLog App.EXEName, "clsIISIntNm", "GetEqpNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 일반장비중 사용중인 장비코드, 장비명 조회
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
    Error.SetLog App.EXEName, "clsIISIntNm", "GetUsingEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사코드 마스터에서 상세자코드, 단일항목만 조회
'   반환 : Recordset
'-----------------------------------------------------------------------------'
Public Function GetTestCd() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectForIntTestCd, , adCmdText)
    Set GetTestCd = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISIntNm", "GetTestCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당검사코드의 검사명 조회(단일, 상세자코드에서만 조회)
'   인수 :
'       - pTestCd : 검사코드
'   반환 : 존재하면 이름 검사명 반환, 존재하지 않으면 Null반환
'-----------------------------------------------------------------------------'
Public Function GetTestNm(ByVal pTestCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectForIntTestNm(pTestCd), , adCmdText)
    
    If Not (Rs.BOF Or Rs.EOF) Then
        GetTestNm = Rs.Fields("TESTNM").Value
    End If
    
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISIntNm", "GetTestNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub

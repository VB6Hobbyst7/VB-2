VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISRef.cls
'   작성자  : 오세원
'   내  용  : 참고치 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISMasterSql

Private mTestCd     As String       '검사코드
Private mSpcCd      As String       '검체코드
Private mSex        As String       '성별
Private mAgeFr      As Long         'From Age
Private mAgeTo      As Long         'To Age
Private mApplyDt    As String       '적용일
Private mExpireDt   As String       '폐기일
Private mRefFg      As String       'Reference Check여부(0:No, 1:Yes)
Private mRefFrVal   As Single       'Reference From Value
Private mRefToVal   As Single       'Reference To Value
Private mRefCd      As String       'Alpha결과 참고치
Private mPanicFg    As String       'Panic Check여부(0:No, 1:Yes)
Private mPanicFrVal As Single       'Panic From Value
Private mPanicToVal As Single       'Panic To Value

Public Property Get TestCd() As String
    TestCd = mTestCd
End Property

Public Property Let TestCd(ByVal vData As String)
    mTestCd = vData
End Property

Public Property Get SpcCd() As String
    SpcCd = mSpcCd
End Property

Public Property Let SpcCd(ByVal vData As String)
    mSpcCd = vData
End Property

Public Property Get Sex() As String
    Sex = mSex
End Property

Public Property Let Sex(ByVal vData As String)
    mSex = vData
End Property

Public Property Get AgeFr() As Long
    AgeFr = mAgeFr
End Property

Public Property Let AgeFr(ByVal vData As Long)
    mAgeFr = vData
End Property

Public Property Get AgeTo() As Long
    AgeTo = mAgeTo
End Property

Public Property Let AgeTo(ByVal vData As Long)
    mAgeTo = vData
End Property

Public Property Get Applydt() As String
    Applydt = mApplyDt
End Property

Public Property Let Applydt(ByVal vData As String)
    mApplyDt = vData
End Property

Public Property Get ExpireDt() As String
    ExpireDt = mExpireDt
End Property

Public Property Let ExpireDt(ByVal vData As String)
    mExpireDt = vData
End Property

Public Property Get RefFg() As String
    RefFg = mRefFg
End Property

Public Property Let RefFg(ByVal vData As String)
    mRefFg = vData
End Property

Public Property Get RefFrVal() As Single
    RefFrVal = mRefFrVal
End Property

Public Property Let RefFrVal(ByVal vData As Single)
    mRefFrVal = vData
End Property

Public Property Get RefToVal() As Single
    RefToVal = mRefToVal
End Property

Public Property Let RefToVal(ByVal vData As Single)
    mRefToVal = vData
End Property

Public Property Get RefCd() As String
    RefCd = mRefCd
End Property

Public Property Let RefCd(ByVal vData As String)
    mRefCd = vData
End Property

Public Property Get PanicFg() As String
    PanicFg = mPanicFg
End Property

Public Property Let PanicFg(ByVal vData As String)
    mPanicFg = vData
End Property

Public Property Get PanicFrVal() As Single
    PanicFrVal = mPanicFrVal
End Property

Public Property Let PanicFrVal(ByVal vData As Single)
    mPanicFrVal = vData
End Property

Public Property Get PanicToVal() As Single
    PanicToVal = mPanicToVal
End Property

Public Property Let PanicToVal(ByVal vData As Single)
    mPanicToVal = vData
End Property

'------------------------------------------------------------------------------'
'   기능 : 참고치 추가
'   반환 : True(성공), False(실패)
'------------------------------------------------------------------------------'
Public Function AddRef() As Boolean
    Dim SQL As String
    
    SQL = mSQL.InsertReference(mTestCd, mSpcCd, mSex, mAgeFr, mAgeTo, mApplyDt, mExpireDt, mRefFrVal, _
                mRefToVal, mRefCd)
                
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddRef = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISRef", "AddRef", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'------------------------------------------------------------------------------'
'   기능 : 참고치 수정
'   반환 : True(성공), False(실패)
'------------------------------------------------------------------------------'
Public Function ModifyRef() As Boolean
    Dim SQL As String
    
    SQL = mSQL.UpdateReference(mTestCd, mSpcCd, mSex, mAgeFr, mAgeTo, mApplyDt, mExpireDt, mRefFrVal, _
                mRefToVal, mRefCd)
                
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyRef = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISRef", "ModifyRef", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'------------------------------------------------------------------------------'
'   기능 : 참고치 삭제
'   반환 : True(성공), False(실패)
'------------------------------------------------------------------------------'
Public Function DelRef() As Boolean
    Dim SQL As String
    
    SQL = mSQL.DeleteReference(mTestCd, mSpcCd, mSex, mAgeFr, mAgeTo, mApplyDt)
    
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelRef = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISRef", "DelRef", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'------------------------------------------------------------------------------'
'   기능 : 해당 검사코드, 검체코드, 성별, 나이에 해당하는 참고치 조회
'   인수 :
'       - pTestCd : 검사코드
'       - pSpcCd  : 검체코드
'       - pSex    : 성별
'       - pAge    : 일령
'   반환 : Recordset
'------------------------------------------------------------------------------'
Public Function GetRef(ByVal pTestCd As String, ByVal pSpcCd As String, _
                       ByVal pSex As String, ByVal pAge As Long) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectReferenceX(pTestCd, pSpcCd, pSex, pAge), , adCmdText)
    Set GetRef = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISRef", "GetRef", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'------------------------------------------------------------------------------'
'   기능 : 해당 검사코드, 검체코드, 성별, 나이에 해당하는 Panic수치 조회
'   인수 :
'       - pTestCd : 검사코드
'       - pSpcCd  : 검체코드
'       - pSex    : 성별
'       - pAge    : 일령
'   반환 : Recordset
'------------------------------------------------------------------------------'
Public Function GetPanic(ByVal pTestCd As String, ByVal pSpcCd As String, _
                         ByVal pSex As String, ByVal pAge As Long) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectPanic(pTestCd, pSpcCd, pSex, pAge), , adCmdText)
    Set GetPanic = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISRef", "GetPanic", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'------------------------------------------------------------------------------'
'   기능 : 해당 검사코드, 결과에 대한 Alpha결과 판정을 조회
'   인수 :
'       - pTestCd : 검사코드
'       - pResult : 결과
'   반환 : H(High), L(Low), *(Abnormal), P(Panic), Null
'------------------------------------------------------------------------------'
Public Function GetAlpahRef(ByVal pTestCd As String, ByVal pResult As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectAlphaRef(pTestCd, pResult), , adCmdText)
    If Not (Rs.BOF Or Rs.EOF) Then
        GetAlpahRef = Rs.Fields("decision").Value & ""
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISRef", "GetAlpahRef", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub

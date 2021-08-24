VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISTSpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISTSpc.cls

'   내  용  : 지정검체 클래스

'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISMasterSql

Private mTestCd     As String       '검사코드
Private mSpcCd      As String       '검체코드
Private mSpcNm      As String       '검체명
Private mApplyDt    As String       '적용일
Private mExpireDt   As String       '폐기일
Private mUnit       As String       '단위
Private mAvalVal    As Long         '유효숫자
Private mPanicFg    As String       'Panic Check (0:No, 1:Yes)
Private mPanicFrVal As Single       'Panic From Value
Private mPanicToVal As Single       'Panic To Value
Private mDeltaFg    As String       'Delta Check (0:No, 1:Yes)
Private mDeltaFrVal As Long         'Delta From Value
Private mDeltaToVal As Long         'Delta To Value

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

Public Property Get SpcNm() As String
    SpcNm = mSpcNm
End Property

Public Property Let SpcNm(ByVal vData As String)
    mSpcNm = vData
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

Public Property Get Unit() As String
    Unit = mUnit
End Property

Public Property Let Unit(ByVal vData As String)
    mUnit = vData
End Property

Public Property Get AvalVal() As Long
    AvalVal = mAvalVal
End Property

Public Property Let AvalVal(ByVal vData As Long)
    mAvalVal = vData
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

Public Property Get DeltaFg() As String
    DeltaFg = mDeltaFg
End Property

Public Property Let DeltaFg(ByVal vData As String)
    mDeltaFg = vData
End Property

Public Property Get DeltaFrVal() As Long
    DeltaFrVal = mDeltaFrVal
End Property

Public Property Let DeltaFrVal(ByVal vData As Long)
    mDeltaFrVal = vData
End Property

Public Property Get DeltaToVal() As Long
    DeltaToVal = mDeltaToVal
End Property

Public Property Let DeltaToVal(ByVal vData As Long)
    mDeltaToVal = vData
End Property

'------------------------------------------------------------------------------'
'   기능 : 지정검체 추가
'   반환 : True(성공), False(실패)
'------------------------------------------------------------------------------'
Public Function AddSpcCd() As Boolean
    Dim SQL As String
    
    SQL = mSQL.InsertTSpcCd(mTestCd, mSpcCd, mApplyDt, mExpireDt, mUnit, mAvalVal, mPanicFg, _
                mPanicFrVal, mPanicToVal, mDeltaFg, mDeltaFrVal, mDeltaToVal)
                
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddSpcCd = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISTSpc", "AddSpcCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'------------------------------------------------------------------------------'
'   기능 : 지정검체 수정
'   반환 : True(성공), False(실패)
'------------------------------------------------------------------------------'
Public Function ModifySpcCd() As Boolean
    Dim SQL As String
    
    SQL = mSQL.UpdateTSpcCd(mTestCd, mSpcCd, mApplyDt, mExpireDt, mUnit, mAvalVal, mPanicFg, _
                mPanicFrVal, mPanicToVal, mDeltaFg, mDeltaFrVal, mDeltaToVal)
    
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifySpcCd = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISTSpc", "ModifySpcCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'------------------------------------------------------------------------------'
'   기능 : 지정검체 삭제
'   반환 : True(성공), False(실패)
'------------------------------------------------------------------------------'
Public Function DelSpcCd() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteTSpcCd(mTestCd, mSpcCd, mApplyDt), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelSpcCd = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISTSpc", "DelSpcCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'------------------------------------------------------------------------------'
'   기능 : 유효숫자 조회
'   인수 :
'       - pTestCd : 검사코드
'       - pSpcCd  : 검체코드
'   반환 : 유효숫자
'------------------------------------------------------------------------------'
Public Function GetAvalValue(ByVal pTestCd As String, ByVal pSpcCd As String) As Long
'
End Function

'------------------------------------------------------------------------------'
'   기능 : Delta/Panic 체크
'   인수 :
'       - pTestCd : 검사코드
'       - pSpcCd  : 검체코드
'       - pAgeDay : 일령
'       - pSex    : 성별
'   반환 : Delta/Panic 판정결과
'------------------------------------------------------------------------------'
Public Function CheckDeltaPanic(ByVal pTestCd As String, ByVal pSpcCd As String, _
                    ByVal pAgeDay As String, ByVal pSex As String) As String
'
End Function


Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub

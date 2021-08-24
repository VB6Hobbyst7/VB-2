VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISEqpMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISEqpMaster.cls (자사 LIS와 조인할때)

'   내  용  : 검사장비 설정

'   버  전  :
'-----------------------------------------------------------------------------'

'## 속성
'   - 장비코드
'   - 장비명
'   - 장비구분
'   - 모델명(Serial-No)
'   - 건물코드
'   - 사용유무
'   - 제조회사코드
'   - 적정온도(상)
'   - 적정온도(하)
'   - 온도구분
'   - 비고

'## 메소드
'   - 건물정보 조회
'   - 제조회사 정보 조회
'   - 장비리스트 조회
'   - 검사장비 입력
'   - 검사장비 수정
'   - 검사장비 삭제

Option Explicit

Private mSQL        As clsIISMasterSql

Private mEqpCd      As String       '장비코드
Private mEqpNm      As String       '장비명
Private mEqpDiv     As String       '장비구분(E:일반장비, C:온도관리장비)
Private mModelNm    As String       '모델명, Serial-No
Private mPurchDt    As String       '구매일자
Private mLocationCd As String       '건물코드
Private mInUseFg    As String       '사용유무(1:사용, 0:사용안함)
Private mVandCd     As String       '건물코드
Private mTempLow    As Single       '적정온도(상)
Private mTempHigh   As Single       '적정온도(하)
Private mTempScale  As String       '온도구분(C:섭씨, F:화씨)
Private mRemark     As String       '비고

Public Property Get EqpCd() As String
    EqpCd = mEqpCd
End Property

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Get EqpNm() As String
    EqpNm = mEqpNm
End Property

Public Property Let EqpNm(ByVal vData As String)
    mEqpNm = vData
End Property

Public Property Get EqpDiv() As String
    EqpDiv = mEqpDiv
End Property

Public Property Let EqpDiv(ByVal vData As String)
    mEqpDiv = vData
End Property

Public Property Get ModelNm() As String
    ModelNm = mModelNm
End Property

Public Property Let ModelNm(ByVal vData As String)
    mModelNm = vData
End Property

Public Property Get PurchDt() As String
    PurchDt = mPurchDt
End Property

Public Property Let PurchDt(ByVal vData As String)
    mPurchDt = vData
End Property

Public Property Get LocationCd() As String
    LocationCd = mLocationCd
End Property

Public Property Let LocationCd(ByVal vData As String)
    mLocationCd = vData
End Property

Public Property Get InUseFg() As String
    InUseFg = mInUseFg
End Property

Public Property Let InUseFg(ByVal vData As String)
    mInUseFg = vData
End Property

Public Property Get VandCd() As String
    VandCd = mVandCd
End Property

Public Property Let VandCd(ByVal vData As String)
    mVandCd = vData
End Property

Public Property Get TempLow() As Single
    TempLow = mTempLow
End Property

Public Property Let TempLow(ByVal vData As Single)
    mTempLow = vData
End Property

Public Property Get TempHigh() As Single
    TempHigh = mTempHigh
End Property

Public Property Let TempHigh(ByVal vData As Single)
    mTempHigh = vData
End Property

Public Property Get TempScale() As String
    TempScale = mTempScale
End Property

Public Property Let TempScale(ByVal vData As String)
    mTempScale = vData
End Property

Public Property Get Remark() As String
    Remark = mRemark
End Property

Public Property Let Remark(ByVal vData As String)
    mRemark = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 건물정보 조회
'   반환 : 건물정보 Recordset (건물코드, 건물명)
'-----------------------------------------------------------------------------'
Public Function GetLocations() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectLocations, , adCmdText)
    Set GetLocations = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISEqpMaster", "GetLocations", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 제조회사 정보조회
'   반환 : 제조회사 정보 Recordset (제조회사코드, 제조회사명)
'-----------------------------------------------------------------------------'
Public Function GetVands() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectVands, , adCmdText)
    Set GetVands = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISEqpMaster", "GetVands", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사장비 조회
'   반환 : 검사장비 Recordset
'-----------------------------------------------------------------------------'
Public Function GetEqps() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectEqps, , adCmdText)
    Set GetEqps = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISEqpMaster", "GetEqps", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사장비 입력
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function AddEqp() As Boolean
    Dim SQL As String
    
    SQL = mSQL.InsertEqp(mEqpCd, mEqpNm, mEqpDiv, mModelNm, mPurchDt, mLocationCd, mInUseFg, _
                mVandCd, mTempLow, mTempHigh, mTempScale, mRemark)
                
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddEqp = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISEqpMaster", "AddEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사장비 수정
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function ModifyEqp() As Boolean
    Dim SQL As String
    
    SQL = mSQL.UpdateEqp(mEqpCd, mEqpNm, mEqpDiv, mModelNm, mPurchDt, mLocationCd, mInUseFg, _
                mVandCd, mTempLow, mTempHigh, mTempScale, mRemark)
    
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyEqp = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISEqpMaster", "ModifyEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사장비 삭제
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function DelEqp() As Boolean
    Dim SQL As String
    
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteEqp(mEqpCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelEqp = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISEqpMaster", "DelEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub

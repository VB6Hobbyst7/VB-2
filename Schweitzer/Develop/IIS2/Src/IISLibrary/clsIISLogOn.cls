VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------------------------'
'   파일명  : clsIISLogOn.cls
'   작성자  : 이상대
'   내  용  : 로그인 클래스
'   작성일  : 2003-12-07
'   버  전  :
'-----------------------------------------------------------------------------'

'## 속성
'   - 직원 아이디
'   - 직원 이름
'   - 로그인 아이디
'   - 로그인 비번

'## 메소드
'   - 사용자 정보 조회

Option Explicit

Private mSQL        As clsIISCommonSql

Private mEmpId      As String       '직원아이디
Private mEmpNm      As String       '직원이름
Private mLoginId    As String       '로그인 아이디
Private mLoginPass  As String       '로그인 비밀번호

Public Property Get EmpId() As String
    EmpId = mEmpId
End Property

Public Property Get EmpNm() As String
    EmpNm = mEmpNm
End Property

Public Property Get LoginId() As String
    LoginId = mLoginId
End Property

Public Property Get LoginPass() As String
    LoginPass = mLoginPass
End Property

'-----------------------------------------------------------------------------'
'   기능 : 사용자 정보를 로드
'   인수 :
'       1. pLoginId : 로그인 아이디
'-----------------------------------------------------------------------------'
Public Function GetEmpInfo(ByVal pLoginId As String) As Boolean
    Dim Rs      As ADODB.Recordset
    Dim SQL     As String
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectEmpInfo(pLoginId), , adCmdText)
    If Rs.BOF Or Rs.EOF Then
        GetEmpInfo = False
    Else
        mEmpId = Rs.Fields("EMPID").Value
        mEmpNm = Rs.Fields("EMPNM").Value
        mLoginId = pLoginId
        mLoginPass = Rs.Fields("LOGINPASS").Value
        GetEmpInfo = True
    End If
    
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Error.SetLog App.EXEName, "clsIISLogOn", "GetEmpInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
    Set Rs = Nothing
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISCommonSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub

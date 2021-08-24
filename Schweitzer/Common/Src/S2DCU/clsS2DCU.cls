VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsS2DCU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyForm As Form
Private mvarDeptCd As String
Private mvarEmpId As String
Private mvarProjectId As String
Private mvarTradeMark As String
Private mvarErrMsg As String
'Private mvarMyDb As DrDatabase   '지역 복사
Private mvarServer As String
Private mvarDatabase As String
Private mvarUID  As String
Private mvarPWD As String
Private mvarCanDelete As Boolean

Enum FormType
    f_TodayNote = 1
    f_ReadNote = 2
    f_WriteNote = 3
End Enum

'Can Delete
Public Property Let CanDelete(ByVal vData As Boolean)
    mvarCanDelete = vData
End Property

Public Property Get CanDelete() As Boolean
    CanDelete = mvarCanDelete
End Property

'Server
Public Property Let Server(ByVal vData As String)
    mvarServer = vData
End Property

Public Property Get Server() As String
    Server = mvarServer
End Property

'Database
Public Property Let Database(ByVal vData As String)
    mvarDatabase = vData
End Property

Public Property Get Database() As String
    Database = mvarDatabase
End Property

'UID
Public Property Let uID(ByVal vData As String)
    mvarUID = vData
End Property

Public Property Get uID() As String
    uID = mvarUID
End Property

'PWD
Public Property Let PWD(ByVal vData As String)
    mvarPWD = vData
End Property

Public Property Get PWD() As String
    PWD = mvarPWD
End Property

'DB Connection 설정
'Public Property Set MyDb(ByVal vData As Object)
'    Set mvarMyDb = vData
'    Set DCU_DBconn = vData
'End Property
'
'Public Property Get MyDb() As Object
'    Set MyDb = mvarMyDb
'End Property

'DeptCd
Public Property Let DeptCd(ByVal vData As String)
    mvarDeptCd = vData
End Property

Public Property Get DeptCd() As String
    DeptCd = mvarDeptCd
End Property

'EmpId
Public Property Let EmpId(ByVal vData As String)
    mvarEmpId = vData
End Property

Public Property Get EmpId() As String
    EmpId = mvarEmpId
End Property

'Error Message
Public Property Let ErrMsg(ByVal vData As String)
    mvarErrMsg = vData
End Property

Public Property Get ErrMsg() As String
    ErrMsg = mvarErrMsg
End Property

'Project Id
Public Property Let ProjectId(ByVal vData As String)
    mvarProjectId = vData
End Property

Public Property Get ProjectId() As String
    ProjectId = mvarProjectId
End Property

'[속성] - Trade Maek (상표)
Public Property Let TradeMark(ByVal vData As String)
    mvarTradeMark = vData
End Property
Public Property Get TradeMark() As String
    TradeMark = mvarTradeMark
End Property

Public Sub FormShow(ByVal WhichForm As FormType)
    Dim RetDbCon As Integer
    Dim blnSuccess As Boolean
    Dim strTypeCd As String
   
'    If MyDb Is Nothing Then
'        Call DCU_DBconn.DbOpen(mvarServer, mvarDatabase, mvarUID, mvarPWD)
'        If DCU_DBconn.DbConnect = False Then
'            MsgBox "Database 연결이 안됬습니다. 전산실로 문의 바랍니다."
'            Exit Sub
'        End If
'    Else
'        If Not MyDb Is Nothing Then Set DCU_DBconn = MyDb
'    End If
    ' WhichForm : 1-초기화면 공지사항, 2-공지사항읽기, 3-공지사항쓰기
    Select Case WhichForm
    Case 1:
        Set MyForm = frmDCU002
        MyForm.ProjectId = mvarProjectId
        MyForm.TradeMark = mvarTradeMark
        Call MyForm.ChkMsgExist
        If Not MyForm.DataExists Then
            Set MyForm = Nothing
            Exit Sub
        End If
    Case 2:
        Set MyForm = frmDCU003
        MyForm.ProjectId = mvarProjectId
        MyForm.TradeMark = mvarTradeMark
        MyForm.CanDelete = mvarCanDelete
    Case 3:
        Set MyForm = frmDCU001
        MyForm.EmpId = mvarEmpId
        MyForm.ProjectId = mvarProjectId
    End Select
    MyForm.WindowState = vbNormal
    MyForm.Show 1
End Sub

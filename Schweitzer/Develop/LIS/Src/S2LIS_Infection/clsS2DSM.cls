VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsS2DSM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private WithEvents MyLogOn As frmLogOn
Attribute MyLogOn.VB_VarHelpID = -1
Private MyForm As Form
Attribute MyForm.VB_VarHelpID = -1


Private mvarServer As String
Private mvarDatabase As String
Private mvarUID  As String
Private mvarPWD As String

Private mvarDeptCd As String
Private mvarEmpID As String
Private mvarLogInID As String
Private mvarLoginPwd As String
Private mvarProjectId As String

Private mvarProductName As String
Private mvarErrMsg As String
Private mvarVersion As String
Private mvarCopyright As String

Private mvarCancelIsEnd As Boolean
Private mvarLockFg As Boolean
Private mvarChangePwd As Boolean

'File Transfer ...
Private mvarHostIP As String
Private mvarHostUID As String
Private mvarHostPWD As String
Private mvarHostDIR As String
Private mvarPrinterNm As String

Private mvarSuccessLogIn As Boolean

Public Event QuitLogon()
Public Event LogonSuccess()

Enum FormType
    f_FormMaster = 1
    f_EmpMaster = 2
    f_GroupMaster = 3
    f_LoginMaster = 4
    f_ChangePwd = 5
End Enum

Private mclsMyUser As New clsDSMLogOn

'// OLD
Private mvarOldLoginId As String
Private mvarOldLogInPass As String
Private mvarOldGroupID As String


'로그온 클래스
Public Property Set MyUser(ByVal pValue As Object)
    Set mclsMyUser = pValue
End Property
Public Property Get MyUser() As Object
    Set MyUser = mclsMyUser
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

'DeptCd
Public Property Let DeptCd(ByVal vData As String)
    mvarDeptCd = vData
End Property
Public Property Get DeptCd() As String
    DeptCd = mvarDeptCd
End Property

'EmpId
Public Property Let EmpId(ByVal vData As String)
    mvarEmpID = vData
End Property
Public Property Get EmpId() As String
    EmpId = mvarEmpID
End Property

'Error Message
Public Property Let ErrMsg(ByVal vData As String)
    mvarErrMsg = vData
End Property
Public Property Get ErrMsg() As String
    ErrMsg = mvarErrMsg
End Property

'Host IP Address
Public Property Let HostIP(ByVal vData As String)
    mvarHostIP = vData
End Property
Public Property Get HostIP() As String
    HostIP = mvarHostIP
End Property

'Host Login ID
Public Property Let HostUID(ByVal vData As String)
    mvarHostUID = vData
End Property
Public Property Get HostUID() As String
    HostUID = mvarHostUID
End Property

'Host Password
Public Property Let HostPWD(ByVal vData As String)
    mvarHostPWD = vData
End Property
Public Property Get HostPWD() As String
    HostPWD = mvarHostPWD
End Property

'Host Directory
Public Property Let HostDIR(ByVal vData As String)
    mvarHostDIR = vData
End Property
Public Property Get HostDIR() As String
    HostDIR = mvarHostDIR
End Property

'Printer Name
Public Property Let PrinterNm(ByVal vData As String)
    mvarPrinterNm = vData
End Property
Public Property Get PrinterNm() As String
    PrinterNm = mvarPrinterNm
End Property

'Version
Public Property Let Version(ByVal vData As String)
    mvarVersion = vData
End Property
Public Property Get Version() As String
    Version = mvarVersion
End Property


'Project Name
Public Property Let ProductName(ByVal vData As String)
    mvarProductName = vData
End Property
Public Property Get ProductName() As String
    ProductName = mvarProductName
End Property

'Project Id
Public Property Let ProjectId(ByVal vData As String)
    mvarProjectId = vData
End Property
Public Property Get ProjectId() As String
    ProjectId = mvarProjectId
End Property

'Copyright
Public Property Let Copyright(ByVal vData As String)
    mvarCopyright = vData
End Property
Public Property Get Copyright() As String
    Copyright = mvarCopyright
End Property

'CancelIsEnd
Public Property Let CancelIsEnd(ByVal vData As Boolean)
    mvarCancelIsEnd = vData
End Property
Public Property Get CancelIsEnd() As Boolean
    CancelIsEnd = mvarCancelIsEnd
End Property

'Lock Flag
Public Property Let lockfg(ByVal vData As Boolean)
    mvarLockFg = vData
End Property
Public Property Get lockfg() As Boolean
    lockfg = mvarLockFg
End Property

'비밀번호 변경여부
Public Property Let ChangePwd(ByVal vData As Boolean)
    mvarChangePwd = vData
End Property
Public Property Get ChangePwd() As Boolean
    ChangePwd = mvarChangePwd
End Property

'로그온 성공여부
Public Property Get SuccessLogIn() As Boolean
    SuccessLogIn = mvarSuccessLogIn
End Property

Public Property Get OldLoginId() As String
    OldLoginId = mvarOldLoginId
End Property
Public Property Let OldLoginId(ByVal vData As String)
    mvarOldLoginId = vData
End Property

Public Property Get OldLogInPass() As String
    OldLogInPass = mvarOldLogInPass
End Property
Public Property Let OldLogInPass(ByVal vData As String)
    mvarOldLogInPass = vData
End Property

Public Property Get OldGroupID() As String
    OldGroupID = mvarOldGroupID
End Property
Public Property Let OldGroupID(ByVal vData As String)
    mvarOldGroupID = vData
End Property


Public Sub LoadLogOn()
    Set MyLogOn = frmLogOn
    With MyLogOn
        .lockfg = mvarLockFg
        .ProductName = mvarProductName
        .ProjectId = mvarProjectId
        .WindowState = vbNormal
        .Show 1
    End With
End Sub

Private Sub Class_Initialize()

    mvarSuccessLogIn = False
    
End Sub

Private Sub Class_Terminate()
   
   Set MyForm = Nothing
   
End Sub



Private Sub MyLogOn_LoginCanceled()
    
    Unload MyLogOn
    Set MyLogOn = Nothing
'    MsgBox "Login 취소!"
    
    'DoEvents
    RaiseEvent QuitLogon

End Sub

Private Sub MyLogOn_LoginFailed()
    
    Unload MyLogOn
    Set MyLogOn = Nothing
    MsgBox "이 시스템을 사용할 수 있는 권한이 없습니다!" & vbCrLf & "프로그램이 종료됩니다.", vbInformation, "Schweitzer2000-" & mvarProjectId
    
    'DoEvents
    RaiseEvent QuitLogon

End Sub

Private Sub MyLogOn_LoginSuccess()

    mvarSuccessLogIn = True
    
    mvarLogInID = MyLogOn.LoginId
    mvarLoginPwd = MyLogOn.LoginPwd
    mvarEmpID = MyLogOn.EmpId
    Set mclsMyUser = MyLogOn.MyUser
    
    Unload MyLogOn
    Set MyLogOn = Nothing
    
    'DoEvents
    RaiseEvent LogonSuccess
    
    
'    MsgBox "Login 성공!"
    
End Sub


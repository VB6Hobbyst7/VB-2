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
Attribute VB_Exposed = True
Option Explicit

Private WithEvents MyLogOn As frmLogOn
Attribute MyLogOn.VB_VarHelpID = -1
Private MySplash As frmSplash
Attribute MySplash.VB_VarHelpID = -1
Private MyForm As Form
Attribute MyForm.VB_VarHelpID = -1

'Private mvarMyDb As DrDatabase   '지역 복사
'Private mvarServer As String
'Private mvarDatabase As String
'Private mvarUID  As String
'Private mvarPWD As String

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

Private WithEvents objLogOn As clsLogON
Attribute objLogOn.VB_VarHelpID = -1

'로그온 클래스
Public Property Set MyUser(ByVal pValue As Object)
    Set mclsMyUser = pValue
End Property

Public Property Get MyUser() As Object
    Set MyUser = mclsMyUser
End Property

'Server
'Public Property Let Server(ByVal vData As String)
'    mvarServer = vData
'End Property
'Public Property Get Server() As String
'    Server = mvarServer
'End Property
'
''Database
'Public Property Let Database(ByVal vData As String)
'    mvarDatabase = vData
'End Property
'Public Property Get Database() As String
'    Database = mvarDatabase
'End Property
'
''UID
'Public Property Let uID(ByVal vData As String)
'    mvarUID = vData
'End Property
'Public Property Get uID() As String
'    uID = mvarUID
'End Property
'
''PWD
'Public Property Let PWD(ByVal vData As String)
'    mvarPWD = vData
'End Property
'Public Property Get PWD() As String
'    PWD = mvarPWD
'End Property

'DB Connection 설정
'Public Property Set MyDb(ByVal vData As DrDatabase)
'    Set mvarMyDb = vData
'    Set dbconn = vData
'End Property
'
'
'Public Property Get MyDb() As DrDatabase
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

Public Property Let LoginId(ByVal vData As String)
    mvarLogInID = vData
End Property

Public Property Get LoginId() As String
    LoginId = mvarLogInID
End Property

Public Property Let LoginPwd(ByVal vData As String)
    mvarLoginPwd = vData
End Property

Public Property Get LoginPwd() As String
    LoginPwd = mvarLoginPwd
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

Public Sub ProcessLogOn()
'로그인 화면을 사용하지 않고 로그인 할경우에만 사용
    Set objLogOn = New clsLogON
    With objLogOn
'        .lockfg = mvarLockFg
'        .ProductName = mvarProductName
        .ProjectId = mvarProjectId
        .LoginId = mvarLogInID
        .LoginPwd = mvarLoginPwd
'        .Show 1
        .PreLogOn
    End With
End Sub

Public Sub LoadSplash()

    Set MySplash = frmSplash
    With MySplash
        .ProductName = mvarProductName
        .Version = mvarVersion
        .Copyright = mvarCopyright
        .WindowState = vbNormal
        .Show
    End With
End Sub

Public Sub UnloadSplash()

    Unload MySplash
    Set MySplash = Nothing
    
End Sub

Public Sub SetSplashMsg(ByVal strMsg As String)
    If MySplash Is Nothing Then Exit Sub
    Call MySplash.ShowMessage(strMsg & "......")
    Call medSleep(1000)
    DoEvents
End Sub

Public Sub FormShow(ByVal WhichForm As FormType)
   
    Dim RetDbCon As Integer
    Dim blnSuccess As Boolean
    Dim strTypeCd As String
    Dim strPWD As String
   
'    If MyDb Is Nothing Then
'        Call dbconn.DbOpen(mvarServer, mvarDatabase, mvarUID, mvarPWD)
'        If dbconn.DbConnect = False Then
'            MsgBox "Database 연결이 안됬습니다. 전산실로 문의 바랍니다.", vbExclamation, "DB연결오류"
'            Exit Sub
'        End If
'    Else
'        If Not MyDb Is Nothing Then Set dbconn = MyDb
'    End If
    
    ' WhichForm : 1-폼관리, 3-그룹등록, 4-사용자관리, 2-직원마스터, 5-비밀번호변경
    
    Select Case WhichForm
        Case 1:
            Set MyForm = frmDSM001
            MyForm.ProjectId = mvarProjectId
        Case 2:
            Set MyForm = frmDSM002
        Case 3:
'            Set MyForm = frmDSM003
'실질적으로 그룹등록 폼에서 projectid를 체크할 이유가 없어서 리마크 처리함
'frmDSM003폼을 사용하려면 아래 리마크는 제거 되어야 함.
'Remark By Legends 2003/09/29
'            MyForm.ProjectId = mvarProjectId

'Modify by legends 2003/09/29
'새로운 화면으로 수정함
            Set MyForm = frmDSM003_N
        Case 4:
            Set MyForm = frmDSM004
        Case 5:
            strPWD = InputBox("변경전 암호를 입력하세요.", "암호확인", "")
            While (strPWD <> mvarLoginPwd)
                MsgBox "암호가 일치하지 않습니다.", vbExclamation + vbOKOnly, "암호입력오류"
                strPWD = InputBox("변경전 암호를 입력하세요.", "암호확인", "")
                If strPWD = "" Then Exit Sub
            Wend
            Set MyForm = frmDSM005
            MyForm.LoginId = mvarLogInID
            MyForm.ChangePwd = True
        Case 6:
            Set MyForm = frmDSM007
    End Select
    
    MyForm.WindowState = vbNormal
    
'직원정보마스터 프로그램에서 모달로 폼을 띄우고 그 위에 팝업을 띄우면 프로그램이 죽네요.
'대신 직원정보 폼에서 ClipCursor 을 사용하여 마우스가 폼 밖을 나가지 못하도록 합니다.

    If WhichForm = "2" Or WhichForm = "6" Then
        MyForm.ZOrder 0
        MyForm.Show 0
    Else
        MyForm.Show 1
    End If
    
    'MyForm.ZOrder 0
    'Call medAlwaysOn(MyForm, 1)
    
'Err_Trap:
'    Resume Next
   
End Sub

Private Sub Class_Initialize()

    mvarSuccessLogIn = False
    
End Sub

Private Sub Class_Terminate()
   
   Set MyForm = Nothing
'   Set dbconn = Nothing
'   Set MyDb = Nothing
   
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

Public Function GetBuildingList(ByVal pCdIndex As String) As String
    Dim sqlBld As String, dsBld As New Recordset, iBldCol As Integer
    Dim sBldCd As String, sBldNm As String, sBldNo As Integer

    sqlBld = "SELECT * FROM " & T_LAB032 & " WHERE " & DBW("cdindex=", pCdIndex) & _
             " ORDER BY field2 asc"
    Call dsBld.Open(sqlBld, DBConn)

    GetBuildingList = ""

    Do While (Not dsBld.EOF)
        sBldCd = "" & dsBld.Fields("cdval1").Value
        sBldNm = "" & dsBld.Fields("field1").Value
        sBldNo = Val("" & dsBld.Fields("field2").Value)
        
        GetBuildingList = GetBuildingList & sBldNo & vbTab & _
                          sBldCd & vbTab & sBldNm & vbCr
        dsBld.MoveNext
    Loop
    
    Set dsBld = Nothing
    
End Function

Private Sub objLogOn_LoginCanceled()
'    Unload MyLogOn
    Set objLogOn = Nothing
'    MsgBox "Login 취소!"
    
    'DoEvents
    RaiseEvent QuitLogon
End Sub

Private Sub objLogOn_LoginFailed()
'    Unload MyLogOn
    Set objLogOn = Nothing
    MsgBox "이 시스템을 사용할 수 있는 권한이 없습니다!" & vbCrLf & "프로그램이 종료됩니다.", vbInformation, "Schweitzer2000-" & mvarProjectId
    
    'DoEvents
    RaiseEvent QuitLogon
End Sub

Private Sub objLogOn_LoginSuccess()
    mvarSuccessLogIn = True
    
    mvarLogInID = objLogOn.LoginId
    mvarLoginPwd = objLogOn.LoginPwd
    mvarEmpID = objLogOn.EmpId
    Set mclsMyUser = objLogOn.MyUser
    
'    Unload MyLogOn
    Set objLogOn = Nothing
    
    'DoEvents
    RaiseEvent LogonSuccess
End Sub

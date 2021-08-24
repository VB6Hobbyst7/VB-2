VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Public Event LoginSuccess()
Public Event LoginCanceled()
Public Event LoginFailed()

Private mvarEmpID As String '�б�����
Private mvarEmpNm As String '�б�����
Private mvarLogInID As String '�б�,����
Private mvarLoginPwd As String '�б�,����
Private mvarProjectId As String '���� ����

Private mclsMyUser As New clsDSMLogOn

'�α׿� Ŭ����
Public Property Set MyUser(ByVal pValue As Object)
    Set mclsMyUser = pValue
End Property

Public Property Get MyUser() As Object
    Set MyUser = mclsMyUser
End Property

'EmpId
Public Property Get EmpId() As String
    EmpId = mvarEmpID
End Property

Public Property Get EmpNM() As String
    EmpNM = mvarEmpNm
End Property

'User Id
Public Property Get LoginId() As String
    LoginId = mvarLogInID
End Property

Public Property Let LoginId(ByVal vData As String)
    mvarLogInID = vData
End Property

'Password
Public Property Get LoginPwd() As String
    LoginPwd = mvarLoginPwd
End Property

Public Property Let LoginPwd(ByVal vData As String)
    mvarLoginPwd = vData
End Property

'Project Id
Public Property Let ProjectId(ByVal vData As String)
    mvarProjectId = vData
End Property

Public Function PreLogOn() As Boolean
'�α�Ƶ� üũ
    With mclsMyUser
        .ProjectId = mvarProjectId
        If CheckManagerOnInstall Then
            mvarEmpID = "-1"
            mvarLogInID = mvarLogInID
            mvarLoginPwd = mvarLoginPwd
            .IsDeveloper = True
            RaiseEvent LoginSuccess
        Else
            Call LoginInfo
        End If
    End With
End Function

Private Sub LoginInfo()
    Dim blnLoginSuccess As Boolean
    
    blnLoginSuccess = mclsMyUser.LoginInfo(mvarLogInID)
    
    With mclsMyUser
        If blnLoginSuccess Then
'            If .LogIn(mvarLoginPwd) Then
                'If mclsMyUser.EmpId <> mvarOldUid Or mclsMyUser.Password <> mvarOldPwd Then Call UnloadForms(medLogOn)
                mvarEmpID = mclsMyUser.EmpId
                mvarLogInID = mvarLogInID
                mvarLoginPwd = mvarLoginPwd
                Call .GetAuthority
                
                If .Permitted Then
                   RaiseEvent LoginSuccess
                Else
                   RaiseEvent LoginFailed
                End If
                'Unload Me
'            Else
'                MsgBox "��й�ȣ�� Ʋ���ϴ�. ��й�ȣ�� Ȯ���ϼ���.", vbOKOnly + vbExclamation, "�α���"
'                If mvarLockFg Then
'                    Me.ScaleMode = vbPixels '3
'                    Call SetMouseClip(Me.Picture1)
'                End If
'                Call txtPASSWD_GotFocus
'                txtPASSWD.SetFocus
'            End If
'        lblName.Caption = mclsMyUser.EmpLngNm
'        txtPASSWD.SetFocus
        Else
            If mclsMyUser.LoginExist Then
                MsgBox "���α׷��� ����� �� �ִ� ������ �����ϴ�.", vbExclamation, "�α��� ����"
'                If mvarLockFg Then
'                    Me.ScaleMode = vbPixels '3
'                    Call SetMouseClip(Me.Picture1)
'                End If
            Else
                MsgBox "��ϵ��� ���� ID�Դϴ�. �α��� ID�� Ȯ���ϼ���.", vbOKOnly + vbExclamation, "�α��� ����"
'                If mvarLockFg Then
'                    Me.ScaleMode = vbPixels '3
'                    Call SetMouseClip(Me.Picture1)
'                End If
            End If
'            Call txtUSERID_GotFocus
'            txtUSERID.SetFocus
        End If
    End With
End Sub

Private Function CheckManagerOnInstall() As Boolean

    Dim strIniFile As String
    Dim strBuffer As String
    
'    strIniFile = App.Path & "\Install.ini"
    strIniFile = InstallDir & "Install.ini"
    
    'UID
    strBuffer = Space$(gintMAX_SIZE)
    If GetPrivateProfileString(INIT_USER_SEC, INIT_UID_KEY, vbNullString, strBuffer, gintMAX_SIZE, strIniFile) > 0 Then
        strBuffer = RTrim$(StripTerminator(strBuffer))
    Else
        strBuffer = vbNullString
    End If
    
    CheckManagerOnInstall = IIf(mvarLogInID = strBuffer, True, False)
    
    'UNM
    strBuffer = Space$(gintMAX_SIZE)
    If GetPrivateProfileString(INIT_USER_SEC, INIT_UNM_KEY, vbNullString, strBuffer, gintMAX_SIZE, strIniFile) > 0 Then
        mvarEmpNm = RTrim$(StripTerminator(strBuffer))
    Else
        mvarEmpNm = vbNullString
    End If
    
'    If IsMissing(mvarLoginPwd) Or Not CheckManagerOnInstall Then Exit Function
    
    'PWD
    strBuffer = Space$(gintMAX_SIZE)
    If GetPrivateProfileString(INIT_USER_SEC, INIT_PWD_KEY, vbNullString, strBuffer, gintMAX_SIZE, strIniFile) > 0 Then
        strBuffer = RTrim$(StripTerminator(strBuffer))
    Else
        strBuffer = vbNullString
    End If
    
    CheckManagerOnInstall = IIf(mvarLogInID = strBuffer, True, False)
End Function

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
VERSION 5.00
Begin VB.Form frmLogOn 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   3990
   ClientLeft      =   1320
   ClientTop       =   810
   ClientWidth     =   6075
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00A4BFC3&
   FillStyle       =   0  '단색
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3990
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F7F0F0&
      Height          =   4005
      Left            =   0
      Picture         =   "medLogOn.frx":0000
      ScaleHeight     =   3945
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   -15
      Width           =   6075
      Begin VB.TextBox txtUSERID 
         Alignment       =   2  '가운데 맞춤
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3885
         TabIndex        =   2
         Top             =   2010
         Width           =   1725
      End
      Begin VB.TextBox txtPASSWD 
         Alignment       =   2  '가운데 맞춤
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  '사용 못함
         Left            =   3885
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   2790
         Width           =   1725
      End
      Begin VB.Label lblCancel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "취 소"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   180
         Left            =   4050
         MouseIcon       =   "medLogOn.frx":41592
         MousePointer    =   99  '사용자 정의
         TabIndex        =   8
         Top             =   3465
         Width           =   495
      End
      Begin VB.Shape shpCancel 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00B9B9B9&
         BorderWidth     =   2
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  '단색
         Height          =   420
         Left            =   3855
         Shape           =   4  '둥근 사각형
         Top             =   3345
         Width           =   855
      End
      Begin VB.Image lmgLock 
         Height          =   705
         Left            =   1245
         Picture         =   "medLogOn.frx":4189C
         Stretch         =   -1  'True
         Top             =   900
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblLockMsg 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE5E7&
         BackStyle       =   0  '투명
         Caption         =   "Locking..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   660
         Left            =   2220
         TabIndex        =   14
         Top             =   885
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00CDD19E&
         BorderWidth     =   3
         Height          =   330
         Left            =   3855
         Top             =   2760
         Width           =   1785
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00CDD19E&
         BorderWidth     =   3
         Height          =   330
         Left            =   3855
         Top             =   2370
         Width           =   1785
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00CDD19E&
         BorderWidth     =   3
         Height          =   330
         Left            =   3855
         Top             =   1980
         Width           =   1785
      End
      Begin VB.Label lblSysName 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Anatomic Pathology"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CDD19E&
         Height          =   555
         Index           =   0
         Left            =   630
         TabIndex        =   12
         Top             =   915
         Width           =   4635
      End
      Begin VB.Label lblSysName 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Anatomic Pathology"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   2
         Left            =   615
         TabIndex        =   13
         Top             =   900
         Width           =   4635
      End
      Begin VB.Label lblSysName 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Anatomic Pathology"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Index           =   1
         Left            =   690
         TabIndex        =   11
         Top             =   945
         Width           =   4635
      End
      Begin VB.Label lblOK 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "확 인"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   5010
         MouseIcon       =   "medLogOn.frx":42166
         MousePointer    =   99  '사용자 정의
         TabIndex        =   9
         Top             =   3465
         Width           =   495
      End
      Begin VB.Shape shpOK 
         BackColor       =   &H00F7F3F8&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00B9B9B9&
         BorderWidth     =   2
         FillColor       =   &H00F7F3F8&
         FillStyle       =   0  '단색
         Height          =   420
         Left            =   4815
         Shape           =   4  '둥근 사각형
         Top             =   3345
         Width           =   855
      End
      Begin VB.Label lblNAME 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3855
         TabIndex        =   6
         Top             =   2430
         Width           =   1785
      End
      Begin VB.Label lblUser 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "사용자ID"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2715
         TabIndex        =   5
         Top             =   2070
         Width           =   1020
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "비밀번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Top             =   2835
         Width           =   990
      End
      Begin VB.Label lblUserNm 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "사용자명"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2715
         TabIndex        =   3
         Top             =   2445
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F1F2E3&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3855
         TabIndex        =   7
         Top             =   2385
         Visible         =   0   'False
         Width           =   1785
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   2430
      TabIndex        =   10
      Top             =   1755
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Public Event LoginSuccess()
Public Event LoginCanceled()
Public Event LoginFailed()

Private mvarProductName As String
Private mvarOldUid As String
Private mvarOldPwd As String
Private mvarEmpID As String
Private mvarLogInID As String
Private mvarLoginPwd As String
Private mvarProjectId As String
Private mvarLockFg As Boolean

Private mclsMyUser As New clsDSMLogOn

'로그온 클래스
Public Property Set MyUser(ByVal pValue As Object)
    Set mclsMyUser = pValue
End Property

Public Property Get MyUser() As Object
    Set MyUser = mclsMyUser
End Property

'Product명 - Logon화면의 Title
Public Property Let ProductName(ByVal pValue As String)
    mvarProductName = pValue
End Property

'Lock Flag
Public Property Get lockfg() As Boolean
    lockfg = mvarLockFg
End Property
Public Property Let lockfg(ByVal pValue As Boolean)
    mvarLockFg = pValue
End Property

'Old User ID
Public Property Get OldUid() As String
    OldUid = mvarOldUid
End Property
Public Property Let OldUid(ByVal pValue As String)
    mvarOldUid = pValue
End Property

'Old Password
Public Property Get OldPwd() As String
    OldPwd = mvarOldPwd
End Property
Public Property Let OldPwd(ByVal pValue As String)
    mvarOldPwd = pValue
End Property

'EmpId
Public Property Get EmpId() As String
    EmpId = mvarEmpID
End Property

'User Id
Public Property Get LoginId() As String
    LoginId = mvarLogInID
End Property

'Password
Public Property Get LoginPwd() As String
    LoginPwd = mvarLoginPwd
End Property

'Project Id
Public Property Let ProjectId(ByVal vData As String)
    mvarProjectId = vData
End Property


Private Sub Form_Activate()
    If mvarLockFg Then
        Me.ScaleMode = vbPixels '3
        Call SetMouseClip(Me.Picture1)
    End If
    txtUSERID.SetFocus
End Sub

'
Private Sub Form_Load()
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    lblOK.ForeColor = DCM_Gray
    lblCancel.ForeColor = DCM_Gray
    shpOK.FillColor = DCM_LightGray
    shpCancel.FillColor = DCM_LightGray
    
    lblSysName(0).Caption = mvarProductName
    lblSysName(1).Caption = mvarProductName
    lblSysName(2).Caption = mvarProductName
    
'    Call mclsMyUser.SetDatabase(dbconn)
    
    lblLockMsg.Visible = mvarLockFg
    If mvarLockFg Then
        lblSysName(0).Visible = False
        lblSysName(1).Visible = False
        lblSysName(2).Visible = False
        lblLockMsg.Visible = True
        lblUser.ForeColor = &H80&
        lblUserNm.ForeColor = &H80&
        lblPassword.ForeColor = &H80&
        Shape1.BorderColor = DCM_LightGray
        Shape2.BorderColor = DCM_LightGray
        Shape3.BorderColor = DCM_LightGray
        lmgLock.Visible = True
        lblCancel.Visible = False
        shpCancel.Visible = False
        Picture1.Picture = LoadPicture()
        Call Dithering(Picture1)
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mvarLockFg Then
        Call ReleaseMouseClip
        Me.ScaleMode = vbTwips   '1
    End If

End Sub

'
Private Sub lblCancel_Click()

    RaiseEvent LoginCanceled

End Sub

Private Sub lblCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblCancel.ForeColor = DCM_Black
    shpCancel.FillColor = DCM_LightPink

End Sub

Private Sub lblOK_Click()

    Dim ShowAtStartup As Variant
    Dim blnRet As Boolean
    Dim strUNM As String

    With mclsMyUser
        If Trim(txtUSERID.Text) = "" Then
           MsgBox "사용자 ID를 입력하세요. ", vbOKOnly + vbExclamation, "로그인"
           If mvarLockFg Then
               Me.ScaleMode = vbPixels '3
               Call SetMouseClip(Me.Picture1)
           End If
           Exit Sub
        End If
        If CheckManagerOnInstall(txtUSERID.Text, strUNM, txtPASSWD.Text) Then
            mvarEmpID = "-1"
            mvarLogInID = txtUSERID.Text
            mvarLoginPwd = txtPASSWD.Text
            .IsDeveloper = True
            RaiseEvent LoginSuccess
        Else
            If .LogIn(txtPASSWD.Text) Then
                'If mclsMyUser.EmpId <> mvarOldUid Or mclsMyUser.Password <> mvarOldPwd Then Call UnloadForms(medLogOn)
                mvarEmpID = mclsMyUser.EmpId
                mvarLogInID = txtUSERID.Text
                mvarLoginPwd = txtPASSWD.Text
                Call .GetAuthority
                
                If .Permitted Then
                   RaiseEvent LoginSuccess
                Else
                   RaiseEvent LoginFailed
                End If
                'Unload Me
            Else
                MsgBox "비밀번호가 틀립니다. 비밀번호를 확인하세요. ", vbOKOnly + vbExclamation, "로그인"
                If mvarLockFg Then
                    Me.ScaleMode = vbPixels '3
                    Call SetMouseClip(Me.Picture1)
                End If
                Call txtPASSWD_GotFocus
                txtPASSWD.SetFocus
            End If
        End If
    End With

End Sub
'
Private Sub lblOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblOK.ForeColor = DCM_Black   '&HFF7B55
    shpOK.FillColor = DCM_LightPink      '&HDDF0F5      '&HDEEEFE
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblOK.ForeColor = DCM_Gray
    lblCancel.ForeColor = DCM_Gray
    shpOK.FillColor = DCM_LightGray
    shpCancel.FillColor = DCM_LightGray

End Sub


Private Sub txtPASSWD_GotFocus()

   With txtPASSWD
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub
'
Private Sub txtPASSWD_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If ObjSysInfo.ProjectId <> "WMN" Then
            If txtPASSWD.Text = "" Then
               MsgBox "비밀번호를 입력하세요. ", vbOKOnly + vbExclamation
               If mvarLockFg Then
                   Me.ScaleMode = vbPixels '3
                   Call SetMouseClip(Me.Picture1)
               End If
               txtPASSWD.SetFocus
               Exit Sub
            End If
        End If
        Call lblOK_Click
    End If
End Sub

Private Sub txtUSERID_Change()
   lblName.Caption = ""
End Sub

Private Sub txtUSERID_GotFocus()
   With txtUSERID
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtUSERID_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub txtUSERID_LostFocus()

    Dim blnLoginSuccess As Boolean
    Dim strUNM As String

    If ActiveControl.Name = "lblCANCEL" Then Exit Sub

    If txtUSERID.Text = "" Then
        MsgBox "로그인 ID를 입력하세요. ", vbOKOnly + vbExclamation
        If mvarLockFg Then
            Me.ScaleMode = vbPixels '3
            Call SetMouseClip(Me.Picture1)
        End If
        txtUSERID.SetFocus
        Exit Sub
    End If
    
    mclsMyUser.ProjectId = mvarProjectId
    If CheckManagerOnInstall(txtUSERID.Text, strUNM) Then
        lblName.Caption = strUNM
        txtPASSWD.SetFocus
    Else
        blnLoginSuccess = mclsMyUser.LoginInfo(txtUSERID.Text)
        
        If blnLoginSuccess Then
            lblName.Caption = mclsMyUser.EmpLngNm
            txtPASSWD.SetFocus
        Else
            If mclsMyUser.LoginExist Then
                MsgBox mvarProductName & " 프로그램을 사용할 수 있는 권한이 없습니다.", vbExclamation, "로그인 실패"
                If mvarLockFg Then
                    Me.ScaleMode = vbPixels '3
                    Call SetMouseClip(Me.Picture1)
                End If
            Else
                MsgBox "등록되지 않은 ID입니다. 로그인 ID를 확인하세요. ", vbOKOnly + vbExclamation, "로그인 실패"
                If mvarLockFg Then
                    Me.ScaleMode = vbPixels '3
                    Call SetMouseClip(Me.Picture1)
                End If
            End If
            Call txtUSERID_GotFocus
            txtUSERID.SetFocus
        End If
    End If
    
End Sub
'

Private Function CheckManagerOnInstall(ByVal strUID As String, ByRef strUNM As String, Optional ByVal strPWD As Variant) As Boolean

    Dim strIniFile As String
    Dim strBuffer As String
    
    strIniFile = App.Path & "\Install.ini"
    
    'UID
    strBuffer = Space$(gintMAX_SIZE)
    If GetPrivateProfileString(INIT_USER_SEC, INIT_UID_KEY, vbNullString, strBuffer, gintMAX_SIZE, strIniFile) > 0 Then
        strBuffer = RTrim$(StripTerminator(strBuffer))
    Else
        strBuffer = vbNullString
    End If
    
    CheckManagerOnInstall = IIf(strUID = strBuffer, True, False)
    
    'UNM
    strBuffer = Space$(gintMAX_SIZE)
    If GetPrivateProfileString(INIT_USER_SEC, INIT_UNM_KEY, vbNullString, strBuffer, gintMAX_SIZE, strIniFile) > 0 Then
        strUNM = RTrim$(StripTerminator(strBuffer))
    Else
        strUNM = vbNullString
    End If
    
    If IsMissing(strPWD) Or Not CheckManagerOnInstall Then Exit Function
    
    'PWD
    strBuffer = Space$(gintMAX_SIZE)
    If GetPrivateProfileString(INIT_USER_SEC, INIT_PWD_KEY, vbNullString, strBuffer, gintMAX_SIZE, strIniFile) > 0 Then
        strBuffer = RTrim$(StripTerminator(strBuffer))
    Else
        strBuffer = vbNullString
    End If
    
    CheckManagerOnInstall = IIf(strPWD = strBuffer, True, False)
    

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


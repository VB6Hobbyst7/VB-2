VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   3435
   ClientLeft      =   2805
   ClientTop       =   3060
   ClientWidth     =   5805
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2029.513
   ScaleMode       =   0  '사용자
   ScaleWidth      =   5450.581
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Left            =   2250
      Top             =   1410
   End
   Begin VB.TextBox txtUserID 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   270
      IMEMode         =   8  '영문
      Left            =   2715
      TabIndex        =   2
      Top             =   2160
      Width           =   1245
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   2715
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2475
      Width           =   1245
   End
   Begin VB.TextBox txtUserName 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   270
      Left            =   3990
      TabIndex        =   0
      Top             =   2160
      Width           =   1635
   End
   Begin HSCotrol.CButton cmdOK 
      Height          =   360
      Left            =   3240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2805
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   635
      BackColor       =   16777215
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      BorderStyle     =   1
      BorderColor     =   -2147483632
   End
   Begin HSCotrol.CButton cmdCancel 
      Height          =   360
      Left            =   4440
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2805
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   635
      BackColor       =   16777215
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      BorderStyle     =   1
      BorderColor     =   -2147483632
   End
   Begin VB.Label lblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "▒ 사용처 : 서대구병원"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   90
      TabIndex        =   13
      Top             =   1770
      Width           =   2835
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Laboratory Information Management Interface System."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   210
      TabIndex        =   12
      Top             =   465
      Width           =   4530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "▒ DgLog-LIMIS ▒"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   345
      Left            =   210
      TabIndex        =   11
      Top             =   105
      Width           =   2685
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   10
      Top             =   930
      Width           =   405
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   210
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "사용자 ID(&U):"
      Height          =   180
      Index           =   0
      Left            =   1515
      TabIndex        =   8
      Top             =   2205
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "암호(&P):"
      Height          =   180
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   2505
      Width           =   690
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "사용자 ID를 입력 하십시오."
      Height          =   180
      Left            =   390
      TabIndex        =   6
      Top             =   3000
      Width           =   2205
   End
   Begin VB.Image imgNet1 
      Height          =   240
      Left            =   120
      Picture         =   "frmLogin.frx":030A
      Top             =   2940
      Width           =   240
   End
   Begin VB.Image imgNet2 
      Height          =   240
      Left            =   120
      Picture         =   "frmLogin.frx":0454
      Top             =   2940
      Width           =   240
   End
   Begin VB.Image imgNet3 
      Height          =   240
      Left            =   120
      Picture         =   "frmLogin.frx":059E
      Top             =   2940
      Width           =   240
   End
   Begin VB.Image imgNet4 
      Height          =   240
      Left            =   120
      Picture         =   "frmLogin.frx":06E8
      Top             =   2940
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "헬스네트워크(www.HelNet.co.kr) Tel. 0505-832-1515"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   420
      Left            =   3000
      TabIndex        =   5
      Top             =   1560
      Width           =   2790
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image3 
      Height          =   2010
      Left            =   30
      Picture         =   "frmLogin.frx":0832
      Stretch         =   -1  'True
      Top             =   30
      Width           =   5745
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   30
      Picture         =   "frmLogin.frx":2E6F
      Stretch         =   -1  'True
      Top             =   2070
      Width           =   5745
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OldUid          As String
Private OldPwd          As String
Private MsgFg           As Boolean
Private OldUser         As UserInfo

Public CancelIsEnd      As Boolean
Public LoginSucceeded   As Boolean

Private adoRS As ADODB.Recordset

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
   Call ReleaseCapture
   Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub cmdCancel_Click()
    If MainForm Is Nothing Then
        Call Unload(Me)
        Set frmLogin = Nothing
        End
    Else
        CurrUser = OldUser
        Call Unload(Me)
        Set frmLogin = Nothing
    End If

End Sub

Private Sub cmdOk_Click()
   Dim ShowAtStartup As Variant

    Timer1.Enabled = False
    imgNet4.ZOrder
    If txtPassword = CurrUser.CuUserPW Then
        If CurrUser.CuPower = Authority.ELVEL_NOT Then
            MsgBox "실행 권한이 없읍니다. 관리자에게 문의 하세요. ", vbOKOnly + vbExclamation
            Exit Sub
        End If
        Call Unload(Me)
        
        If MainForm Is Nothing Then
            Set MainForm = New MDIMain
            MainForm.Show
            MainForm.stbMain.Panels(1).text = CurrUser.CuUserNM
        Else
            MainForm.stbMain.Panels(1).text = CurrUser.CuUserNM
        End If
      Else
         MsgBox "비밀번호가 틀립니다. 비밀번호를 확인하세요. ", vbOKOnly + vbExclamation
         txtPassword.SetFocus
         txtPassword.SelStart = 0
         txtPassword.SelLength = Len(txtPassword)
      End If

End Sub

Private Sub Form_Activate()
    txtUserID.SetFocus
End Sub

Private Sub Form_Load()

    imgNet1.ZOrder 0
    Timer1.Interval = 500
    Timer1.Enabled = True
    
    lblTitle.Caption = App.Title
    lblVersion.Caption = "Ver. " & App.major & "." & App.minor & "." & App.Revision
    lblSite.Caption = " ▒ 사용처 : " & App.CompanyName
    
    If Not MainForm Is Nothing Then
        OldUser = CurrUser
    End If
    
End Sub

Private Sub Timer1_Timer()
    DoEvents

    If imgNet2.Visible = True Then
        imgNet2.Visible = False
        imgNet3.Visible = True
        imgNet3.ZOrder
    Else
        imgNet3.Visible = False
        imgNet2.Visible = True
        imgNet2.ZOrder
    End If

End Sub

Private Sub txtPassword_GotFocus()
   With txtPassword
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        Call cmdOk_Click
        KeyAscii = 0
    End If
End Sub

Private Sub txtUserID_Change()
   txtUserName.text = ""
End Sub

Private Sub txtUserID_GotFocus()
   With txtUserID
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
        Call txtUserID_LostFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtUserID_LostFocus()
    Dim ret As Boolean

    Dim objUserInf As clsCommon
    On Error GoTo ErrorTrap

    If ActiveControl.Name = "cmdCancel" Then Exit Sub

        If txtUserID.text = "" Then
            MsgFg = True
            MsgBox "로그온 ID를 입력하세요. ", vbOKOnly + vbExclamation
            MsgFg = False
            txtUserID.SetFocus
            Exit Sub
        End If

        labMsg.Caption = "데이타 베이스에 연결중 ...."
        Screen.MousePointer = vbArrowHourglass

        Set objUserInf = New clsCommon
        With objUserInf
            .SetAdoCn AdoCn_Jet
            Set AdoRs_Jet = .Get_UserInfo(txtUserID)
            If AdoRs_Jet Is Nothing Then
                MsgBox "등록되지 않은 ID입니다. 로그인 ID를 확인하세요. ", vbOKOnly + vbExclamation
                With txtUserID
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(.text)
                End With
                Set objUserInf = Nothing
            End If
        End With

        Screen.MousePointer = vbDefault
        labMsg.Caption = "데이타 베이스에 연결 되었습니다."

        If AdoRs_Jet.EOF Then
            MsgBox "등록되지 않은 ID입니다. 로그인 ID를 확인하세요. ", vbOKOnly + vbExclamation
            Set AdoRs_Jet = Nothing
            Set objUserInf = Nothing
            With txtUserID
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.text)
            End With
        Else
            Timer1.Enabled = False
            With CurrUser
                .CuUserID = AdoRs_Jet.Fields("EMP_ID") & ""
                .CuUserNM = AdoRs_Jet.Fields("EMP_NM") & ""
                .CuUserPW = AdoRs_Jet.Fields("PASSWD") & ""
                .CuPower = AdoRs_Jet.Fields("POWERS") & ""
                txtUserName = .CuUserNM
            End With
            imgNet4.ZOrder 0
            txtPassword.SetFocus
            AdoRs_Jet.Close
        End If

ErrorTrap:
    Set AdoRs_Jet = Nothing
    Set objUserInf = Nothing
End Sub

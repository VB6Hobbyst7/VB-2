VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   " Log on"
   ClientHeight    =   3330
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5685
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1967.475
   ScaleMode       =   0  '사용자
   ScaleWidth      =   5337.907
   StartUpPosition =   2  '화면 가운데
   Begin BHButton.BHImageButton cmdCancel 
      Height          =   375
      Left            =   4365
      TabIndex        =   11
      Top             =   2790
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      Caption         =   "취소"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton cmdOK 
      Height          =   375
      Left            =   3105
      TabIndex        =   10
      Top             =   2790
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      Caption         =   "확인"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   4455
      Picture         =   "frmLogin.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   960
      TabIndex        =   9
      Top             =   450
      Width           =   1020
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   8
      Top             =   2655
      Width           =   5610
   End
   Begin VB.TextBox txtUserID 
      Appearance      =   0  '평면
      Height          =   270
      IMEMode         =   8  '영문
      Left            =   3135
      TabIndex        =   3
      Top             =   1050
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  '평면
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   3135
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1755
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   3135
      TabIndex        =   1
      Top             =   1395
      Width           =   2325
   End
   Begin VB.Timer Timer1 
      Left            =   3540
      Top             =   525
   End
   Begin VB.Image Image1 
      Height          =   2115
      Left            =   105
      Picture         =   "frmLogin.frx":1B4E
      Stretch         =   -1  'True
      Top             =   435
      Width           =   1590
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Laboratory Information Management Advanced System."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   7
      Top             =   90
      Width           =   5415
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "사용자 ID(&U):"
      Height          =   180
      Index           =   0
      Left            =   1935
      TabIndex        =   6
      Top             =   1095
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "암호(&P):"
      Height          =   180
      Index           =   1
      Left            =   2340
      TabIndex        =   5
      Top             =   1785
      Width           =   690
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "사용자 이름 :"
      Height          =   180
      Index           =   2
      Left            =   1950
      TabIndex        =   4
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "사용자 ID를 입력 하십시오."
      Height          =   180
      Left            =   2280
      TabIndex        =   0
      Top             =   2340
      Width           =   2205
   End
   Begin VB.Image imgNet1 
      Height          =   240
      Left            =   1980
      Picture         =   "frmLogin.frx":382A
      Top             =   2310
      Width           =   240
   End
   Begin VB.Image imgNet2 
      Height          =   240
      Left            =   1980
      Picture         =   "frmLogin.frx":3974
      Top             =   2310
      Width           =   240
   End
   Begin VB.Image imgNet3 
      Height          =   240
      Left            =   1980
      Picture         =   "frmLogin.frx":3ABE
      Top             =   2310
      Width           =   240
   End
   Begin VB.Image imgNet4 
      Height          =   240
      Left            =   1980
      Picture         =   "frmLogin.frx":3C08
      Top             =   2310
      Width           =   240
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
            MainForm.stbMain.Panels(1).Text = CurrUser.CuUserNM
        Else
            MainForm.stbMain.Panels(1).Text = CurrUser.CuUserNM
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
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        Call cmdOk_Click
        KeyAscii = 0
    End If
End Sub

Private Sub txtUserID_Change()
   txtUserName.Text = ""
End Sub

Private Sub txtUserID_GotFocus()
   With txtUserID
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
        Call txtUserID_LostFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtUserID_LostFocus()
    Dim Ret As Boolean

    Dim objUserInf As clsCommon
    On Error GoTo ErrorTrap

    If ActiveControl.Name = "cmdCancel" Then Exit Sub

        If txtUserID.Text = "" Then
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
                    .SelLength = Len(.Text)
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
                .SelLength = Len(.Text)
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

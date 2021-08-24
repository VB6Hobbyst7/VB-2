VERSION 5.00
Begin VB.Form frmLogOn 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLogOn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogOn.frx":030A
   ScaleHeight     =   4095
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtUserCd 
      BackColor       =   &H00FFFFF0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2340
      MaxLength       =   6
      TabIndex        =   0
      Top             =   2700
      Width           =   1005
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFF0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  '사용 못함
      Left            =   2340
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3090
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5355
      Picture         =   "frmLogOn.frx":4437
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   3180
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Caption         =   "확인"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4410
      MaskColor       =   &H000000FF&
      Picture         =   "frmLogOn.frx":4A14
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   3180
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   945
      TabIndex        =   7
      Top             =   1755
      Width           =   5100
   End
   Begin VB.Label lblUserNm 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   3360
      TabIndex        =   6
      Top             =   2700
      Width           =   2970
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "사용자 ID :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   885
      TabIndex        =   4
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label Label2 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   975
      TabIndex        =   5
      Top             =   3150
      Width           =   1290
   End
End
Attribute VB_Name = "frmLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const vbApplicationModal As Integer = 0   'Application modal; the user must respond to the message box before continuing work in the current application.
Private Const vbSystemModal As Integer = 4096     'System modal; all applications are suspended until the user responds to the message box.
Private ConnectCount As Integer
Private SOPEN As Integer

Private SqlConn As Long

Private Sub cmdcancel_Click()

    If SOPEN = QSQL_SUCCESS Then Call Qsqlclose(SqlConn, ONECLOSE)
    
    If Me.Tag = "LOGON" Then
        Unload Me
    Else
        End
    End If
    
End Sub

Private Sub cmdOk_Click()

    Dim UserCd As String
    Dim ConnectOk As Integer
    
    Dim SqlStr  As String
    
    UserCd = Trim(txtUserCd)
    If UserCd = "" Then Exit Sub
    
    ConnectOk = False
    
    SqlStr = "SELECT COUNT(USERID) FROM BAS_DB..BAS010M " _
            & " WHERE USERID = '" & UserCd & Chr$(39) _
            & "   AND USERPW = '" & txtPassword & Chr$(39)
            
    ConnectOk = G_EXIST_RECORD(SqlConn, SqlStr)
    If ConnectOk Then
        ConnectOk = Trim(txtPassword) = Get_Record(SqlConn, "USERPW", "BAS_DB..BAS010M", "USERID", UserCd + "' and OUTDATE = '")
             
        If Me.Tag = "LOGON" Then
            D0COM_USERID = UserCd
'            mdiMain.stbMsgbar.Panels(3) = lblUserNm
        Else
            D0COM_USERID = UserCd
        End If
        Unload Me
    Else
        ConnectCount = ConnectCount + 1
        If ConnectCount < 3 Then
            MsgBox "사원번호와 Password를 확인하십시요." & Chr$(13) & "다시 입력해 주십시요.", vbApplicationModal & vbExclamation
            txtUserCd.SetFocus
        Else
            MsgBox "사원번호와 Password를 확인하십시요.", vbApplicationModal & vbCritical
            End
        End If
    End If
            
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)

'    If KeyAscii = 13 Then SendKeys "{TAB}"
    
End Sub

Private Sub Form_Load()
    
    If Me.BorderStyle = 5 Then
        Me.Width = 6855
        Me.Height = 4215
    End If
    
    Me.Top = (Screen.Height - Me.Height) / 3
    Me.Left = (Screen.Width - Me.Width) / 4 + 1300
    
'    D0COM_TERMID = Get_Term_Id
    lblTitle.Caption = Title & "  장비 인터페이스"
    
    SOPEN = QSqlOpen(D0COM_SERVER01, Me.hWnd, SqlConn)
    If SOPEN <> QSQL_SUCCESS Then cmdOk.Enabled = False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If SOPEN = QSQL_SUCCESS Then Call Qsqlclose(SqlConn, ONECLOSE)
    Me.Tag = ""
    
End Sub

Private Sub txtPassword_Change()

    If txtPassword.MaxLength = txtPassword.SelStart Then SendKeys "{TAB}"
    
End Sub

Private Sub txtPassword_GotFocus()

    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
    
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub

Private Sub txtUserCd_Change()

    If txtUserCd.MaxLength = txtUserCd.SelStart Then SendKeys "{TAB}"
    
End Sub

Private Sub txtUserCd_GotFocus()
    
    txtUserCd.SelStart = 0
    txtUserCd.SelLength = Len(txtUserCd.Text)
    
    txtUserCd.Tag = txtUserCd.Text
    
End Sub

Private Sub txtUserCd_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub


Private Sub txtUserCd_LostFocus()
    
    If txtUserCd.Text = txtUserCd.Tag Then Exit Sub
    
    lblUserNm = Get_Record(SqlConn, "USERNM", "BAS_DB..BAS010M", "USERID", Trim(txtUserCd))

End Sub

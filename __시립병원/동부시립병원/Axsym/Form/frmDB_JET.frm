VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmDB_JET 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Log DataBase Setting"
   ClientHeight    =   2760
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   5655
   ControlBox      =   0   'False
   Icon            =   "frmDB_JET.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Threed.SSPanel SSPanel1 
      Height          =   375
      Left            =   930
      TabIndex        =   11
      Top             =   90
      Width           =   3915
      _Version        =   65536
      _ExtentX        =   6906
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "JET 데이타베이스 등록"
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
   End
   Begin Threed.SSCommand cmdSerch 
      Height          =   285
      Left            =   5010
      TabIndex        =   10
      Top             =   785
      Width           =   345
      _Version        =   65536
      _ExtentX        =   609
      _ExtentY        =   503
      _StockProps     =   78
      BevelWidth      =   1
      Outline         =   0   'False
      Picture         =   "frmDB_JET.frx":000C
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4290
      TabIndex        =   9
      Top             =   2310
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   3090
      TabIndex        =   8
      Top             =   2310
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   30
      TabIndex        =   7
      Top             =   2040
      Width           =   5580
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   210
      Top             =   1485
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPasswd 
      Appearance      =   0  '평면
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   1770
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "Passwd"
      Top             =   1680
      Width           =   2445
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1755
      TabIndex        =   1
      Text            =   "User"
      Top             =   1230
      Width           =   2445
   End
   Begin VB.TextBox txtFilename 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1755
      TabIndex        =   0
      Text            =   "SERVER"
      Top             =   780
      Width           =   3210
   End
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      Caption         =   "암호(&B):"
      Height          =   180
      Index           =   0
      Left            =   1035
      TabIndex        =   6
      Top             =   1725
      Width           =   690
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   90
      TabIndex        =   5
      Top             =   2505
      Width           =   60
   End
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      Caption         =   "데이타 베이스명(&S):"
      Height          =   180
      Index           =   6
      Left            =   75
      TabIndex        =   4
      Top             =   825
      Width           =   1650
   End
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      Caption         =   "사용자명(&U):"
      Height          =   180
      Index           =   4
      Left            =   675
      TabIndex        =   3
      Top             =   1275
      Width           =   1050
   End
End
Attribute VB_Name = "frmDB_JET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bConnected As Boolean

Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub cmdOk_Click()
    
    If Trim(txtFilename) = "" Then
        MsgBox " 데이타 베이스를 선택 하시오."
        Exit Sub
    ElseIf Trim(txtUser) = "" Then
        MsgBox " 사용자명을 입력 하시오."
        Exit Sub
    Else
        Call SaveString(HKEY_CURRENT_USER, REG_JETDB, REG_DATABASE, Trim(txtFilename))
        Call SaveString(HKEY_CURRENT_USER, REG_JETDB, REG_USER_ID, Trim(txtUser))
        Call SaveString(HKEY_CURRENT_USER, REG_JETDB, REG_PASSWD, Trim(txtPasswd))
        
        If DBConnect_MDS Then
            labMsg.Caption = "데이타 베이스를 찾고있습니다."
            Unload Me
        Else
            MsgBox "  연결되지 않았습니다. 다시 시도 하십시오."
            txtFilename.Enabled = True
            txtFilename.SetFocus
        End If
    End If
End Sub

Private Sub cmdSerch_Click()
    
    With CommonDialog1
      .CancelError = True
      On Error GoTo ErrHandler
      .Flags = cdlOFNHideReadOnly
      .InitDir = App.Path
      .Filter = "MS Access Files (*.MDB)|*.MDB|All Files (*.*)|*.*|"
      .FilterIndex = 1
      .FileName = "Interface.mdb"
      .ShowOpen
      txtFilename = .FileName
    End With

Exit Sub
  
ErrHandler:
  ' 사용자가 [취소] 단추를 눌렀습니다.
Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape
            Call cmdCancel_Click
        Case vbKeyReturn
            Call cmdOk_Click
        Case Else
        
    End Select
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbDefault
    txtFilename = GetString(HKEY_CURRENT_USER, REG_JETDB, REG_DATABASE)
    txtUser = GetString(HKEY_CURRENT_USER, REG_JETDB, REG_USER_ID)
    txtPasswd = GetString(HKEY_CURRENT_USER, REG_JETDB, REG_PASSWD)
    cmdOK.Enabled = True
    
End Sub

Private Sub txtFilename_LostFocus()
    KeyPreview = True
End Sub

Private Sub txtPasswd_GotFocus()
    txtPasswd.SelStart = 0
    txtPasswd.SelLength = Len(txtPasswd)
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOk_Click
        KeyAscii = 0
    End If
End Sub

Private Sub txtFilename_GotFocus()
    KeyPreview = False
    txtFilename.SelStart = 0
    txtFilename.SelLength = Len(txtFilename)
End Sub

Private Sub txtFilename_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSerch_Click
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtUser_GotFocus()
    txtUser.SelStart = 0
    txtUser.SelLength = Len(txtUser)
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

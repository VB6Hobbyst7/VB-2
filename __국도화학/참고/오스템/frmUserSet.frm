VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmUserSet 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "User Setting"
   ClientHeight    =   3270
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   5610
   ControlBox      =   0   'False
   Icon            =   "frmUserSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3570
      TabIndex        =   8
      Top             =   2430
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2010
      TabIndex        =   7
      Top             =   2430
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   30
      TabIndex        =   5
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
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  '사용 못함
      Left            =   1950
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Passwd"
      Top             =   1350
      Width           =   2865
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1935
      TabIndex        =   0
      Text            =   "User"
      Top             =   900
      Width           =   2865
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "  사용자 등록"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   5565
   End
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      Caption         =   "암호(&B):"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   1005
      TabIndex        =   4
      Top             =   1380
      Width           =   855
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   90
      TabIndex        =   3
      Top             =   2505
      Width           =   60
   End
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      Caption         =   "사용자명(&U):"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   555
      TabIndex        =   2
      Top             =   945
      Width           =   1305
   End
End
Attribute VB_Name = "frmUserSet"
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
    
    If Trim(txtUser) = "" Then
        MsgBox " 사용자명을 입력 하시오."
        Exit Sub
    Else
        Call SaveString(HKEY_CURRENT_USER, REG_POSITION, REG_USER_ID, Trim(txtUser))
        Call SaveString(HKEY_CURRENT_USER, REG_POSITION, REG_PASSWD, Trim(txtPasswd))
        
        If DBConnect_MDS Then
            Unload Me
        Else
            MsgBox "입력정보가 틀립니다. 다시 시도 하십시오."
        End If
    End If
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
    
    txtUser = ""
    txtPasswd = ""
    
    cmdOk.Enabled = True
    
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

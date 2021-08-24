VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmDB_ORACLE 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "SERVER"
   ClientHeight    =   2925
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   5190
   ControlBox      =   0   'False
   Icon            =   "frmDB_ORACLE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3840
      TabIndex        =   9
      Top             =   2460
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   2640
      TabIndex        =   8
      Top             =   2460
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   15
      TabIndex        =   5
      Top             =   2145
      Width           =   5145
   End
   Begin VB.TextBox txtPasswd 
      Appearance      =   0  '평면
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   1665
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "Passwd"
      Top             =   1755
      Width           =   2205
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1665
      TabIndex        =   3
      Text            =   "User"
      Top             =   1275
      Width           =   2205
   End
   Begin VB.TextBox txtServer 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1665
      TabIndex        =   2
      Top             =   780
      Width           =   3030
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   2325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDB_ORACLE.frx":000C
            Key             =   "Server"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDB_ORACLE.frx":0326
            Key             =   "DBase"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDB_ORACLE.frx":0640
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   375
      Left            =   660
      TabIndex        =   10
      Top             =   90
      Width           =   3915
      _Version        =   65536
      _ExtentX        =   6906
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "Oracle Server 등록"
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
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      Caption         =   "서버(&S):"
      Height          =   195
      Index           =   6
      Left            =   810
      TabIndex        =   1
      Top             =   825
      Width           =   750
   End
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      Caption         =   "암호(&B):"
      Height          =   180
      Index           =   0
      Left            =   870
      TabIndex        =   7
      Top             =   1800
      Width           =   690
   End
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      Caption         =   "사용자명(&U):"
      Height          =   180
      Index           =   1
      Left            =   510
      TabIndex        =   6
      Top             =   1320
      Width           =   1050
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   3420
      Width           =   60
   End
End
Attribute VB_Name = "frmDB_ORACLE"
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
    
    If Trim(txtServer) = "" Then
        MsgBox " ORACLE Server 이름이나 IP를 입력 하시오.", vbExclamation, "입력 오류"
        txtServer.SetFocus
        Exit Sub
    Else
        Call SaveString(HKEY_CURRENT_USER, REG_ORACLEDB, REG_SERVER, txtServer)
        Call SaveString(HKEY_CURRENT_USER, REG_ORACLEDB, REG_USER_ID, Trim(txtUser))
        Call SaveString(HKEY_CURRENT_USER, REG_ORACLEDB, REG_PASSWD, txtPasswd)

        If DbConnect_ORACLE Then
            labMsg.Caption = "Looking for the ORACLE Server & Database."
            Unload Me
        Else
            MsgBox "  Not Connected, So retry. "
            txtServer.Enabled = True
            txtServer.SetFocus
        End If
    End If
End Sub

Private Sub Form_Initialize()
    Screen.MousePointer = vbDefault
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
    txtServer = GetString(HKEY_CURRENT_USER, REG_ORACLEDB, REG_SERVER)
    txtUser = GetString(HKEY_CURRENT_USER, REG_ORACLEDB, REG_USER_ID)
    txtPasswd = GetString(HKEY_CURRENT_USER, REG_ORACLEDB, REG_PASSWD)
End Sub

Private Sub txtServer_GotFocus()
    txtServer.SelStart = 0
    txtServer.SelLength = Len(txtServer)
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

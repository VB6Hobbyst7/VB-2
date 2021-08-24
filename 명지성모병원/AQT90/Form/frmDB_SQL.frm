VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmDB_SQL 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "SERVER"
   ClientHeight    =   2925
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   5145
   ControlBox      =   0   'False
   Icon            =   "frmDB_SQL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   510
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   900
      BackColor       =   16777215
      CaptionBackColor=   16777215
      Picture         =   "frmDB_SQL.frx":000C
      Caption         =   "SQL Server 등록"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtServer 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1665
      TabIndex        =   0
      Top             =   720
      Width           =   3030
   End
   Begin VB.TextBox txtDB 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1665
      TabIndex        =   1
      Top             =   1050
      Width           =   3030
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1665
      TabIndex        =   2
      Text            =   "User"
      Top             =   1455
      Width           =   2205
   End
   Begin VB.TextBox txtPasswd 
      Appearance      =   0  '평면
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   1665
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Passwd"
      Top             =   1755
      Width           =   2205
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   15
      TabIndex        =   9
      Top             =   2145
      Width           =   5145
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
            Picture         =   "frmDB_SQL.frx":045E
            Key             =   "Server"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDB_SQL.frx":0778
            Key             =   "DBase"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDB_SQL.frx":0A92
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin BHButton.BHImageButton cmdOk 
      Height          =   375
      Left            =   2475
      TabIndex        =   11
      Top             =   2430
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      Caption         =   "Ok"
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
   Begin BHButton.BHImageButton cmdCancle 
      Height          =   375
      Left            =   3765
      TabIndex        =   12
      Top             =   2430
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      Caption         =   "Cancle"
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
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      Caption         =   "데이터베이스(&B):"
      Height          =   195
      Index           =   4
      Left            =   105
      TabIndex        =   5
      Top             =   1110
      Width           =   1455
   End
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      Caption         =   "사용자명(&U):"
      Height          =   180
      Index           =   1
      Left            =   510
      TabIndex        =   6
      Top             =   1500
      Width           =   1050
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
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   180
      Left            =   90
      TabIndex        =   8
      Top             =   3420
      Width           =   60
   End
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      Caption         =   "서버(&S):"
      Height          =   195
      Index           =   6
      Left            =   810
      TabIndex        =   4
      Top             =   765
      Width           =   750
   End
End
Attribute VB_Name = "frmDB_SQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bConnected As Boolean


Private Sub cmdCancle_Click()
    Unload Me
    End
End Sub

Private Sub cmdOk_Click()
    
    If Trim(txtServer) = "" Then
        MsgBox " SQL Server 이름이나 IP를 입력 하시오.", vbExclamation, "입력 오류"
        txtServer.SetFocus
        Exit Sub
    ElseIf Trim(txtDB) = "" Then
        MsgBox " SQL Server의 DB이름을 입력 하시오.", vbExclamation, "입력 오류"
        txtDB.SetFocus
        Exit Sub
    Else
        Call SaveString(HKEY_CURRENT_USER, REG_MSSQLDB, REG_SERVER, txtServer)
        Call SaveString(HKEY_CURRENT_USER, REG_MSSQLDB, REG_DATABASE, Trim(txtDB))
        Call SaveString(HKEY_CURRENT_USER, REG_MSSQLDB, REG_USER_ID, Trim(txtUser))
        Call SaveString(HKEY_CURRENT_USER, REG_MSSQLDB, REG_PASSWD, txtPasswd)

        If DbConnect_SQL Then
            labMsg.Caption = "Looking for the SQL Server & Database."
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
            Call cmdCancle_Click
        Case vbKeyReturn
            Call cmdOk_Click
        Case Else
        
    End Select
    
End Sub

Private Sub Form_Load()
    txtServer = GetString(HKEY_CURRENT_USER, REG_MSSQLDB, REG_SERVER)
    txtDB = GetString(HKEY_CURRENT_USER, REG_MSSQLDB, REG_DATABASE)
    txtUser = GetString(HKEY_CURRENT_USER, REG_MSSQLDB, REG_USER_ID)
    txtPasswd = GetString(HKEY_CURRENT_USER, REG_MSSQLDB, REG_PASSWD)
End Sub

Private Sub txtServer_GotFocus()
    txtServer.SelStart = 0
    txtServer.SelLength = Len(txtServer)
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

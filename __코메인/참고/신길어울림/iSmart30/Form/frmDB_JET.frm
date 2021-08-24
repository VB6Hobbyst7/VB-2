VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmDB_JET 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Log DataBase Setting"
   ClientHeight    =   2850
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   5625
   ControlBox      =   0   'False
   Icon            =   "frmDB_JET.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
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
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   1005
      BackColor       =   16777215
      CaptionBackColor=   16777215
      Picture         =   "frmDB_JET.frx":000C
      Caption         =   "JET Database 등록"
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
   Begin HSCotrol.CButton cmdSerch 
      Height          =   285
      Left            =   4995
      TabIndex        =   9
      Top             =   765
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   503
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDB_JET.frx":0CCE
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   -2147483632
   End
   Begin BHButton.BHImageButton cmdOk 
      Height          =   375
      Left            =   2970
      TabIndex        =   10
      Top             =   2340
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
      Left            =   4275
      TabIndex        =   11
      Top             =   2340
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      Caption         =   "Cancel"
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

Private Sub cmdCancle_Click()
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
            Call cmdCancle_Click
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
    cmdOk.Enabled = True
    
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

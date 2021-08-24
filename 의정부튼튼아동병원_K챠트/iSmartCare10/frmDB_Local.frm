VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDB_Local 
   BackColor       =   &H00BF8B59&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "데이터베이스 설정"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   Icon            =   "frmDB_Local.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00AE8B59&
      BorderStyle     =   0  '없음
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   8055
      TabIndex        =   9
      Top             =   0
      Width           =   8055
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   90
         Top             =   90
         Width           =   2865
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "로컬 데이터베이스 설정"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   10
         Top             =   180
         Width           =   2625
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7470
      TabIndex        =   8
      Top             =   1020
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4680
      TabIndex        =   7
      Top             =   2790
      Width           =   1545
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6330
      TabIndex        =   6
      Top             =   2790
      Width           =   1545
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2370
      TabIndex        =   1
      Top             =   1440
      Width           =   2115
   End
   Begin VB.TextBox txtPasswd 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  '사용 못함
      Left            =   2370
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1905
      Width           =   2115
   End
   Begin VB.TextBox txtFilename 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   990
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7350
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "암호 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   10
      Left            =   1635
      TabIndex        =   5
      Top             =   1965
      Width           =   615
   End
   Begin VB.Label 사용자명 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "사용자명 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   9
      Left            =   1260
      TabIndex        =   4
      Top             =   1530
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "데이터베이스 경로 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   8
      Left            =   405
      TabIndex        =   3
      Top             =   1080
      Width           =   1860
   End
End
Attribute VB_Name = "frmDB_Local"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdFind_Click()
    
    With CommonDialog1
      .CancelError = True
      On Error GoTo ErrHandler
      .Flags = cdlOFNHideReadOnly
      .InitDir = App.PATH
      .Filter = "MS Access Files (*.MDB)|*.MDB|All Files (*.*)|*.*|"
      .FilterIndex = 1
      .Filename = "Interface.mdb"
      .ShowOpen
      txtFilename = .Filename
    End With

Exit Sub
  
ErrHandler:
  ' 사용자가 [취소] 단추를 눌렀습니다.
Exit Sub

End Sub

Private Sub cmdSave_Click()
    Dim strPath As String
    Dim strUID  As String
    Dim strPWD  As String
    Dim blnLUN As Boolean
    
    Dim intYear As Integer
    Dim intMon  As Integer
    Dim intDay  As Integer
    
    If Trim(txtFilename) = "" Then
        MsgBox " 데이타 베이스를 선택 하세요"
        Exit Sub
    ElseIf Trim(txtUser) = "" Then
        MsgBox " 사용자명을 입력 하세요"
        Exit Sub
    ElseIf Trim(txtPasswd) = "" Then
        MsgBox " 비밀번호를 입력 하세요"
        Exit Sub
    Else
        strPath = txtFilename.Text
        strUID = txtUser.Text
        strPWD = txtPasswd.Text
        
        intYear = Year(Now)
        intMon = Month(Now)
        intDay = Day(Now)
        
        If Not GetSOL2LUN(intYear, intMon, intDay, strPWD) Then
            MsgBox "비밀번호가 틀립니다."
            Exit Sub
        End If
        
        'If strPWD <> Format(Now, "yyyymmdd") - 503 Then
        '    MsgBox "비밀번호가 틀립니다."
        '    Exit Sub
        'End If
        
        Call WritePrivateProfileString("DATABASE", "LOCALPATH", strPath, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("DATABASE", "LOCALUID", strUID, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("DATABASE", "LOCALPWD", strPWD, App.PATH & "\INI\" & gMACH & ".ini")
                
        'Call GetSetup
        '-- LOCAL DB GET
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "LOCALPATH", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gLocalDB.PATH = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "LOCALUID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gLocalDB.UID = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "LOCALPWD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gLocalDB.PWD = Trim(strSetUp1)

        If DbConnect_Local Then
            'labMsg.Caption = "데이타 베이스를 찾고있습니다."
            Call LetEqpMaster(gHOSP.MACHCD)
            Unload Me
        Else
            MsgBox "  연결되지 않았습니다. 다시 시도 하십시오."
            txtFilename.Enabled = True
            txtFilename.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()

    txtFilename.Text = gLocalDB.PATH
    txtUser.Text = gLocalDB.UID
    txtPasswd.Text = gLocalDB.PWD
        
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        End
    End If
End Sub


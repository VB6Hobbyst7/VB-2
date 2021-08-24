VERSION 5.00
Begin VB.Form frmDB_Oracle 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   " ◈ Oracle 데이터베이스 설정 ◈"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   Icon            =   "frmDB_Oracle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdChange 
      Caption         =   "설정 열기"
      Height          =   300
      Left            =   2940
      TabIndex        =   9
      Top             =   3210
      Width           =   2055
   End
   Begin VB.TextBox txtSID 
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
      Left            =   2910
      TabIndex        =   5
      Top             =   1680
      Width           =   2115
   End
   Begin VB.TextBox txtPWD 
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
      Left            =   2910
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2595
      Width           =   2115
   End
   Begin VB.TextBox txtUID 
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
      Left            =   2910
      TabIndex        =   3
      Top             =   2130
      Width           =   2115
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   5730
      TabIndex        =   1
      Top             =   0
      Width           =   5730
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "Oracle DB 설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   540
         Width           =   3135
      End
      Begin VB.Image Image3 
         Height          =   1065
         Left            =   0
         Picture         =   "frmDB_Oracle.frx":000C
         Top             =   0
         Width           =   12900
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  '아래 맞춤
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   5730
      TabIndex        =   0
      Top             =   3600
      Width           =   5730
      Begin VB.Image imgMenuInsert 
         Height          =   375
         Left            =   1500
         Picture         =   "frmDB_Oracle.frx":174F
         Top             =   270
         Width           =   1725
      End
      Begin VB.Image imgMenuCancel 
         Height          =   375
         Left            =   3330
         Picture         =   "frmDB_Oracle.frx":254B
         Top             =   270
         Width           =   1725
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "데이터베이스 변경 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   945
      TabIndex        =   10
      Top             =   3270
      Width           =   1860
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   4
      Left            =   690
      Picture         =   "frmDB_Oracle.frx":32A3
      Top             =   3240
      Width           =   150
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   0
      Left            =   690
      Picture         =   "frmDB_Oracle.frx":368D
      Top             =   1740
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "서버(SID) : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   8
      Left            =   1680
      TabIndex        =   8
      Top             =   1770
      Width           =   1125
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   1
      Left            =   690
      Picture         =   "frmDB_Oracle.frx":3A77
      Top             =   2190
      Width           =   150
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
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   9
      Left            =   1800
      TabIndex        =   7
      Top             =   2220
      Width           =   1005
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   2
      Left            =   690
      Picture         =   "frmDB_Oracle.frx":3E61
      Top             =   2625
      Width           =   150
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
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   10
      Left            =   2175
      TabIndex        =   6
      Top             =   2655
      Width           =   615
   End
End
Attribute VB_Name = "frmDB_Oracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChange_Click()
    Unload Me
    frmEMRInfo.Show vbModal
End Sub

Private Sub Form_Load()

    txtSID.Text = gORADB.SID
    txtUID.Text = gORADB.UID
    txtPWD.Text = gORADB.PWD
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        End
    End If
End Sub

Private Sub imgMenuCancel_Click()
    End
End Sub

Private Sub imgMenuInsert_Click()
    Dim strSID  As String
    Dim strUID  As String
    Dim strPWD  As String
    
    If Trim(txtSID) = "" Then
        MsgBox " SID를 입력 하세요"
        Exit Sub
    ElseIf Trim(txtUID) = "" Then
        MsgBox " 사용자명을 입력 하세요"
        Exit Sub
    ElseIf Trim(txtPWD) = "" Then
        MsgBox " 비밀번호를 입력 하세요"
        Exit Sub
    Else
        strSID = txtSID.Text
        strUID = txtUID.Text
        strPWD = txtPWD.Text
        
        Call WritePrivateProfileString("DATABASE", "ORACLESID", strSID, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("DATABASE", "ORACLEUID", strUID, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("DATABASE", "ORACLEPWD", strPWD, App.PATH & "\INI\" & gMACH & ".ini")
        
        'Call GetSetup
        '-- ORACLE DB SET
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "ORACLESID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gORADB.SID = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "ORACLEUID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gORADB.UID = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "ORACLEPWD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gORADB.PWD = Trim(strSetUp1)

        If DbConnect_ORACLE Then
            Unload Me
        Else
            MsgBox "  연결되지 않았습니다. 다시 시도 하십시오.", vbOKOnly, Me.Caption
            txtSID.Enabled = True
            txtSID.SetFocus
        End If
    End If
End Sub

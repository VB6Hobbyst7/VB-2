VERSION 5.00
Begin VB.Form frmHospInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "병원정보설정"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5685
   Icon            =   "frmHospInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CheckBox chkDBCon 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DB연결 확인"
      Height          =   345
      Left            =   3660
      TabIndex        =   37
      Top             =   5280
      Width           =   1365
   End
   Begin VB.TextBox txtBarLen 
      Alignment       =   2  '가운데 맞춤
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
      Height          =   300
      Left            =   2460
      TabIndex        =   35
      Top             =   3930
      Width           =   2565
   End
   Begin VB.ComboBox cboMachs 
      Height          =   300
      Left            =   2460
      TabIndex        =   33
      Top             =   720
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.CommandButton cmdLocalDBSet 
      Caption         =   "로컬 DB"
      Height          =   375
      Left            =   270
      TabIndex        =   32
      Top             =   5670
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdDBSet 
      Caption         =   "서버 DB"
      Height          =   375
      Left            =   270
      TabIndex        =   31
      Top             =   6090
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   1860
      TabIndex        =   30
      Top             =   6120
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
      Left            =   3510
      TabIndex        =   29
      Top             =   6120
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00808000&
      BorderStyle     =   0  '없음
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   5685
      TabIndex        =   27
      Top             =   0
      Width           =   5685
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   90
         Top             =   90
         Width           =   2865
      End
      Begin VB.Label lblHosp 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "병원정보 설정"
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
         Left            =   210
         TabIndex        =   28
         Top             =   180
         Width           =   2625
      End
   End
   Begin VB.TextBox txtColWidth 
      Alignment       =   2  '가운데 맞춤
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
      Height          =   300
      Left            =   2460
      TabIndex        =   10
      Top             =   3540
      Width           =   2565
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   345
      Left            =   2460
      TabIndex        =   25
      Top             =   4350
      Width           =   2565
      Begin VB.OptionButton optWorkPos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "팝업"
         Height          =   225
         Index           =   1
         Left            =   1200
         TabIndex        =   12
         Top             =   60
         Width           =   1125
      End
      Begin VB.OptionButton optWorkPos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "메인"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   60
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.CheckBox chkLog 
      BackColor       =   &H00FFFFFF&
      Caption         =   "로그기록"
      Height          =   345
      Left            =   3660
      TabIndex        =   15
      Top             =   5640
      Width           =   1365
   End
   Begin VB.TextBox txtPartNm 
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
      Height          =   300
      Left            =   3240
      TabIndex        =   5
      Top             =   1965
      Width           =   1785
   End
   Begin VB.TextBox txtLabNm 
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
      Height          =   300
      Left            =   3240
      TabIndex        =   3
      Top             =   1560
      Width           =   1785
   End
   Begin VB.TextBox txtHospNm 
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
      Height          =   300
      Left            =   3240
      TabIndex        =   1
      Top             =   1170
      Width           =   1785
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   345
      Left            =   2460
      TabIndex        =   23
      Top             =   4800
      Width           =   2565
      Begin VB.OptionButton optLoginUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "미사용"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   13
         Top             =   90
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optLoginUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         Height          =   225
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   90
         Width           =   1125
      End
   End
   Begin VB.TextBox txtUserNm 
      Alignment       =   2  '가운데 맞춤
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
      Height          =   300
      Left            =   2460
      TabIndex        =   9
      Top             =   3150
      Width           =   2565
   End
   Begin VB.TextBox txtUserID 
      Alignment       =   2  '가운데 맞춤
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
      Height          =   300
      Left            =   2460
      TabIndex        =   8
      Top             =   2745
      Width           =   2565
   End
   Begin VB.TextBox txtMachNm 
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
      Height          =   300
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2355
      Width           =   1785
   End
   Begin VB.TextBox txtLabCd 
      Alignment       =   2  '가운데 맞춤
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
      Height          =   300
      Left            =   2460
      TabIndex        =   2
      Top             =   1560
      Width           =   765
   End
   Begin VB.TextBox txtPartCd 
      Alignment       =   2  '가운데 맞춤
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
      Height          =   300
      Left            =   2460
      TabIndex        =   4
      Top             =   1965
      Width           =   765
   End
   Begin VB.TextBox txtHospCd 
      Alignment       =   2  '가운데 맞춤
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
      Height          =   300
      Left            =   2460
      TabIndex        =   0
      Top             =   1170
      Width           =   765
   End
   Begin VB.TextBox txtMachCd 
      Alignment       =   2  '가운데 맞춤
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
      Height          =   300
      Left            =   2460
      TabIndex        =   6
      Top             =   2355
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "바코드길이 "
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
      Index           =   6
      Left            =   1305
      TabIndex        =   36
      Top             =   4035
      Width           =   1050
   End
   Begin VB.Label lblMachNm 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "장비명"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   615
      TabIndex        =   34
      Top             =   780
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "검사명 넓이 "
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
      Index           =   5
      Left            =   1230
      TabIndex        =   26
      Top             =   3645
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "워크조회 "
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
      Left            =   1500
      TabIndex        =   24
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "로그인 "
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
      Index           =   4
      Left            =   1695
      TabIndex        =   22
      Top             =   4920
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "당당자 명 "
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
      Index           =   3
      Left            =   1425
      TabIndex        =   21
      Top             =   3225
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "당당자 ID "
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
      Index           =   2
      Left            =   1425
      TabIndex        =   20
      Top             =   2820
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "검사파트 코드/명 "
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
      Left            =   720
      TabIndex        =   19
      Top             =   2055
      Width           =   1620
   End
   Begin VB.Label 사용자명 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "사용부서 코드/명 "
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
      Left            =   735
      TabIndex        =   18
      Top             =   1635
      Width           =   1620
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "병원 코드/명 "
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
      Left            =   1110
      TabIndex        =   17
      Top             =   1230
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "장비 코드/명 "
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
      Index           =   0
      Left            =   1125
      TabIndex        =   16
      Top             =   2430
      Width           =   1230
   End
End
Attribute VB_Name = "frmHospInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BHImageButton1_Click()

End Sub

Private Sub cboMachs_Click()
    
    lblMachNm.Caption = cboMachs.Text

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDBSet_Click()
    
    If gDBTYPE = "99" Then
        
        Call WritePrivateProfileString("HOSP", "HOSPCD", txtHospCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Else
        If InputBox("비밀번호 입력") = "dev0503" Then
            If gDBTYPE = "1" Then
                frmDB_Oracle.Show vbModal
            ElseIf gDBTYPE = "2" Then
                frmDB_MSSQL.Show vbModal
            ElseIf gDBTYPE = "3" Then
                frmDB_PGSQL.Show vbModal
            Else
                MsgBox App.PATH & "\OKSOFT.ini 파일에서" & vbNewLine & vbNewLine & "DBTYPE을 먼저 설정하세요 ", vbOKOnly + vbInformation, "DB TYPE 설정"
            End If
        End If
    End If
    
End Sub

Private Sub cmdLocalDBSet_Click()
    
    frmDB_Local.Show vbModal

End Sub

Private Sub cmdSave_Click()
    Dim strDBType   As String
    
    Call WritePrivateProfileString("HOSP", "HOSPCD", txtHospCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "HOSPNM", txtHospNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "LABCD", txtLabCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "LABNM", txtLabNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "PARTCD", txtPartCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "PARTNM", txtPartNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "MACHCD", txtMachCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "MACHNM", txtMachNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "USERID", txtUserID.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "USERNM", txtUserNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    If optLoginUse(0).Value = True Then
        Call WritePrivateProfileString("HOSP", "LOGINYN", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        Call WritePrivateProfileString("HOSP", "LOGINYN", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    If chkLog.Value = "1" Then
        Call WritePrivateProfileString("HOSP", "LOGWRITE", "1", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        Call WritePrivateProfileString("HOSP", "LOGWRITE", "0", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    If optWorkPos(0).Value = True Then
        Call WritePrivateProfileString("VIEW", "WORKPOS", "M", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        Call WritePrivateProfileString("VIEW", "WORKPOS", "P", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    Call WritePrivateProfileString("VIEW", "COLWIDTH", txtColWidth.Text, App.PATH & "\INI\" & gMACH & ".ini")
    
    If lblMachNm.Caption <> "" Then
        Call WritePrivateProfileString("EXE", "MACH", lblMachNm.Caption, App.PATH & "\OKSOFT.ini")
    End If
    
    If chkDBCon.Value = "1" Then
        Call WritePrivateProfileString("HOSP", "DBCONCHK", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        Call WritePrivateProfileString("HOSP", "DBCONCHK", "N", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    If txtBarLen.Text <> "" And IsNumeric(txtBarLen.Text) Then
        Call WritePrivateProfileString("HOSP", "BARLEN", txtBarLen.Text, App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    If gLocalDB.PATH <> "" Then
        Call LetEqpMaster(Trim(txtMachCd.Text))
    End If
    
    SQL = ""
    SQL = SQL & "UPDATE EQPMASTER SET " & vbCrLf
    SQL = SQL & " EQUIPCD = " & STS(txtMachCd.Text)
    
    Call DBExec(AdoCn_Local, SQL)
    
    GetSetup
    
    Unload Me

    Call Main
End Sub

Private Sub Form_Load()
    
    Call CtlInitializing
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Public Sub CtlInitializing()
    Dim i As Integer
    
    txtHospCd.Text = gHOSP.HOSPCD
    txtHospNm.Text = gHOSP.HOSPNM
    txtLabCd.Text = gHOSP.LABCD
    txtLabNm.Text = gHOSP.LABNM
    txtPartCd.Text = gHOSP.PARTCD
    txtPartNm.Text = gHOSP.PARTNM
    txtMachCd.Text = gHOSP.MACHCD
    txtMachNm.Text = gMACH 'gHOSP.MACHNM
    txtUserID.Text = gHOSP.USERID
    txtUserNm.Text = gHOSP.USERNM
    txtBarLen.Text = gHOSP.BARLEN
    
    If gHOSP.DBCONCHK = "Y" Then
        chkDBCon.Value = "1"
    Else
        chkDBCon.Value = "0"
    End If
    
    If gHOSP.LOGINYN = "Y" Then
        optLoginUse(1).Value = True
    Else
        optLoginUse(0).Value = True
    End If
    If gHOSP.LOQWRITE = "1" Then
        chkLog.Value = "1"
    Else
        chkLog.Value = "0"
    End If
    
    If gWORKPOS = "P" Then
        optWorkPos(1).Value = True
    Else
        optWorkPos(0).Value = True
    End If
    
    If gCOLWIDTH = "" Then
        txtColWidth.Text = "10"
    Else
        txtColWidth.Text = gCOLWIDTH
    End If
    
    lblMachNm.Caption = ""
    cboMachs.Clear
    If IsNumeric(gMACHCOUNT) Then
        For i = 1 To gMACHCOUNT
            cboMachs.AddItem gMACHS(i)
        Next
        cboMachs.ListIndex = 0
    End If
    
End Sub


Private Sub lblHosp_DblClick()
    
    If cboMachs.Visible = False Then
        If InputBox("비밀번호 입력" & Space(5) & "hint:개발자oyh") = "dev0503" Then
            cboMachs.Visible = True
            lblMachNm.Visible = True
            cmdLocalDBSet.Visible = True
            cmdDBSet.Visible = True
        End If
    Else
        cboMachs.Visible = False
        lblMachNm.Visible = False
        cmdLocalDBSet.Visible = False
        cmdDBSet.Visible = False
    End If

End Sub

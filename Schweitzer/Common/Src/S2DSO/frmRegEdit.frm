VERSION 5.00
Begin VB.Form frmRegEdit 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Registry 등록"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows 기본값
   WindowState     =   2  '최대화
   Begin VB.Frame fraBldInfo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Building Information"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   4545
      TabIndex        =   31
      Top             =   3465
      Width           =   4395
      Begin VB.TextBox txtBldNm 
         BorderStyle     =   0  '없음
         Height          =   270
         Left            =   1485
         TabIndex        =   15
         Top             =   1260
         Width           =   2565
      End
      Begin VB.TextBox txtBldCd 
         BorderStyle     =   0  '없음
         Height          =   270
         Left            =   1485
         TabIndex        =   14
         Top             =   825
         Width           =   2565
      End
      Begin VB.TextBox txtBldNo 
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         Height          =   255
         Left            =   1485
         TabIndex        =   13
         Top             =   420
         Width           =   2565
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0FF&
         BorderWidth     =   2
         Height          =   330
         Index           =   10
         Left            =   1455
         Shape           =   4  '둥근 사각형
         Top             =   1245
         Width           =   2640
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0FF&
         BorderWidth     =   2
         Height          =   330
         Index           =   9
         Left            =   1455
         Shape           =   4  '둥근 사각형
         Top             =   810
         Width           =   2640
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0FF&
         BorderWidth     =   2
         Height          =   330
         Index           =   8
         Left            =   1455
         Shape           =   4  '둥근 사각형
         Top             =   390
         Width           =   2640
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Name  :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   11
         Left            =   255
         TabIndex        =   34
         Top             =   1305
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Code  :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   10
         Left            =   240
         TabIndex        =   33
         Top             =   885
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "No.  :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   9
         Left            =   255
         TabIndex        =   32
         Top             =   465
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장"
      Height          =   450
      Left            =   6600
      TabIndex        =   16
      Top             =   5565
      Width           =   1155
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "종료"
      Height          =   450
      Left            =   7830
      TabIndex        =   17
      Top             =   5550
      Width           =   1155
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   660
      Left            =   450
      TabIndex        =   29
      Top             =   330
      Width           =   8520
      Begin VB.ComboBox cboProject 
         Height          =   300
         ItemData        =   "frmRegEdit.frx":0000
         Left            =   1395
         List            =   "frmRegEdit.frx":000A
         Style           =   2  '드롭다운 목록
         TabIndex        =   0
         Top             =   240
         Width           =   2700
      End
      Begin VB.TextBox txtExeName 
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         Height          =   255
         Left            =   5595
         TabIndex        =   1
         Top             =   255
         Width           =   2565
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00CFB4CF&
         BorderWidth     =   2
         Height          =   330
         Index           =   11
         Left            =   5565
         Shape           =   4  '둥근 사각형
         Top             =   225
         Width           =   2640
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "EXE Name :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   12
         Left            =   4260
         TabIndex        =   35
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Project   : "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   30
         Top             =   315
         Width           =   1080
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Option"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   435
      TabIndex        =   27
      Top             =   3465
      Width           =   4095
      Begin VB.CheckBox chkSplash 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Splash화면 띄우기"
         Height          =   270
         Left            =   225
         TabIndex        =   10
         Top             =   390
         Width           =   3120
      End
      Begin VB.CheckBox chkShowAtStart 
         BackColor       =   &H00E0E0E0&
         Caption         =   "프로그램 시작 시 공지사항 조회"
         Height          =   270
         Left            =   225
         TabIndex        =   12
         Top             =   1065
         Width           =   3120
      End
      Begin VB.CheckBox chkUseBldFg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "건물정보 사용"
         Height          =   270
         Left            =   225
         TabIndex        =   11
         Top             =   720
         Width           =   3120
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   4560
      TabIndex        =   23
      Top             =   1110
      Width           =   4365
      Begin VB.TextBox txtAppPath 
         BorderStyle     =   0  '없음
         Height          =   270
         Left            =   1500
         TabIndex        =   9
         Top             =   1695
         Width           =   2565
      End
      Begin VB.TextBox txtHospital 
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         Height          =   255
         Left            =   1500
         TabIndex        =   6
         Top             =   435
         Width           =   2565
      End
      Begin VB.TextBox txtHelpLine 
         BorderStyle     =   0  '없음
         Height          =   270
         Left            =   1500
         TabIndex        =   7
         Top             =   840
         Width           =   2565
      End
      Begin VB.TextBox txtFileServer 
         BorderStyle     =   0  '없음
         Height          =   270
         Left            =   1500
         TabIndex        =   8
         Top             =   1275
         Width           =   2565
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "App Path   :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   5
         Left            =   165
         TabIndex        =   28
         Top             =   1740
         Width           =   1080
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   2
         Height          =   330
         Index           =   4
         Left            =   1470
         Shape           =   4  '둥근 사각형
         Top             =   1680
         Width           =   2640
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Hospital   :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   8
         Left            =   165
         TabIndex        =   26
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Help Line  :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   7
         Left            =   150
         TabIndex        =   25
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "File Server:"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   6
         Left            =   165
         TabIndex        =   24
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   2
         Height          =   330
         Index           =   7
         Left            =   1470
         Shape           =   4  '둥근 사각형
         Top             =   405
         Width           =   2640
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   2
         Height          =   330
         Index           =   6
         Left            =   1470
         Shape           =   4  '둥근 사각형
         Top             =   825
         Width           =   2640
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   2
         Height          =   330
         Index           =   5
         Left            =   1470
         Shape           =   4  '둥근 사각형
         Top             =   1260
         Width           =   2640
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Database"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   450
      TabIndex        =   18
      Top             =   1110
      Width           =   4095
      Begin VB.ComboBox cboDBType 
         Height          =   300
         ItemData        =   "frmRegEdit.frx":0048
         Left            =   1380
         List            =   "frmRegEdit.frx":0055
         Style           =   2  '드롭다운 목록
         TabIndex        =   37
         Top             =   240
         Width           =   2475
      End
      Begin VB.TextBox txtServer 
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         Height          =   255
         Left            =   1410
         TabIndex        =   2
         Top             =   645
         Width           =   2415
      End
      Begin VB.TextBox txtDbNm 
         BorderStyle     =   0  '없음
         Height          =   270
         Left            =   1410
         TabIndex        =   3
         Top             =   1005
         Width           =   2415
      End
      Begin VB.TextBox txtLogin 
         BorderStyle     =   0  '없음
         Height          =   270
         Left            =   1410
         TabIndex        =   4
         Top             =   1380
         Width           =   2415
      End
      Begin VB.TextBox txtPwd 
         BorderStyle     =   0  '없음
         Height          =   270
         IMEMode         =   3  '사용 못함
         Left            =   1410
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1755
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "※"
         Height          =   180
         Left            =   3870
         TabIndex        =   38
         Top             =   675
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Type     :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   13
         Left            =   240
         TabIndex        =   36
         Top             =   315
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Server   :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   4
         Left            =   225
         TabIndex        =   22
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Database :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   21
         Top             =   1065
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Login    :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   2
         Left            =   225
         TabIndex        =   20
         Top             =   1425
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   3
         Left            =   210
         TabIndex        =   19
         Top             =   1800
         Width           =   900
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FBC2A4&
         BorderWidth     =   2
         Height          =   330
         Index           =   0
         Left            =   1380
         Shape           =   4  '둥근 사각형
         Top             =   615
         Width           =   2490
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FBC2A4&
         BorderWidth     =   2
         Height          =   330
         Index           =   1
         Left            =   1380
         Shape           =   4  '둥근 사각형
         Top             =   990
         Width           =   2490
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FBC2A4&
         BorderWidth     =   2
         Height          =   330
         Index           =   2
         Left            =   1380
         Shape           =   4  '둥근 사각형
         Top             =   1365
         Width           =   2490
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FBC2A4&
         BorderWidth     =   2
         Height          =   330
         Index           =   3
         Left            =   1380
         Shape           =   4  '둥근 사각형
         Top             =   1740
         Width           =   2490
      End
   End
End
Attribute VB_Name = "frmRegEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mvarTradeMark As String

Dim strAppName As String

Public Property Let TradeMark(ByVal vData As String)
    mvarTradeMark = vData
End Property
Public Property Get TradeMark() As String
    TradeMark = mvarTradeMark
End Property


Private Sub ReadRegistry()
    Dim strDbType As String
    
    On Error Resume Next
    
    '서버, 데이타베이스, 로그인아뒤, 비밀번호
    txtServer.Text = GetSetting(AppName:=strAppName, Section:=RegSsSvr, Key:=RegK1Svr, Default:="")
    txtDbNm.Text = GetSetting(AppName:=strAppName, Section:=RegSsSvr, Key:=RegK2Svr, Default:="")
    txtLogin.Text = GetSetting(AppName:=strAppName, Section:=RegSsSvr, Key:=RegK3Svr, Default:="")
    txtPwd.Text = GetSetting(AppName:=strAppName, Section:=RegSsSvr, Key:=RegK4Svr, Default:="")
    strDbType = GetSetting(AppName:=strAppName, Section:=RegSsSvr, Key:=RegK5Svr, Default:="")
    If txtServer = "" Or txtDbNm = "" Then txtServer = "(Undefined)"
    If strDbType >= 0 And strDbType <= cboDBType.ListCount - 1 Then
        cboDBType.ListIndex = strDbType
    Else
        cboDBType.ListIndex = -1
    End If
    
    
    'File Server
    txtFileServer.Text = GetSetting(AppName:=strAppName, Section:=RegSsSet, Key:=RegK1Set, Default:="")
    '임시...개발자를 위해..^^;
    'txtFileServer.Text = GetSetting(AppName:="Schweitzer2000", Section:="Develop", Key:="ServerIP", Default:="")
    
    '병원명
    txtHospital.Text = GetSetting(AppName:=strAppName, Section:=RegSsSet, Key:=RegK2Set, Default:="")
    txtHelpLine.Text = GetSetting(AppName:=strAppName, Section:=RegSsSet, Key:=RegK3Set, Default:="")
    '건물정보 사용여부
    chkUseBldFg.Value = GetSetting(AppName:=strAppName, Section:=RegSsBld, Key:=RegK0Bld, Default:="0")
    '건물정보(코드,명,번호)
    txtBldCd.Text = GetSetting(AppName:=strAppName, Section:=RegSsBld, Key:=RegK1Bld, Default:="")
    txtBldNm.Text = GetSetting(AppName:=strAppName, Section:=RegSsBld, Key:=RegK2Bld, Default:="(건물정보 누락)")
    txtBldNo.Text = GetSetting(AppName:=strAppName, Section:=RegSsBld, Key:=RegK3Bld, Default:=0)
'
    '공지사항 조회여부
    chkShowAtStart.Value = GetSetting(strAppName, RegSsOpt, RegK1Opt, "1")
    chkSplash.Value = GetSetting(strAppName, RegSsOpt, RegK2Opt, "1")
    
    '임시...
    txtAppPath.Text = GetSetting(strAppName, RegSsApp, RegK1App, "")
    txtExeName.Text = GetSetting(strAppName, RegSsApp, RegK2App, "")

End Sub

Private Sub cboProject_Click()
    strAppName = mvarTradeMark & " " & Mid(cboProject.Text, 1, 3)
    Call ReadRegistry
End Sub

Private Sub chkUseBldFg_Click()
    fraBldInfo.Visible = Choose(chkUseBldFg.Value + 1, False, True)
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmRegEdit = Nothing
End Sub

Private Sub cmdSave_Click()

    If cboProject.ListIndex = -1 Then
        Exit Sub
    End If
    
    If cboDBType.ListIndex = -1 Then
        Exit Sub
    End If
    
    SaveSetting strAppName, RegSsSvr, RegK1Svr, txtServer.Text
    SaveSetting strAppName, RegSsSvr, RegK2Svr, txtDbNm.Text
    SaveSetting strAppName, RegSsSvr, RegK3Svr, txtLogin.Text
    SaveSetting strAppName, RegSsSvr, RegK4Svr, txtPwd.Text
    SaveSetting strAppName, RegSsSvr, RegK5Svr, cboDBType.ListIndex
    
    '임시로...
    SaveSetting strAppName, RegSsApp, RegK1App, txtAppPath.Text
    SaveSetting strAppName, RegSsApp, RegK2App, txtExeName.Text
    SaveSetting strAppName, RegSsSet, RegK1Set, txtFileServer.Text
'    SaveSetting "Schweitzer2000", "Develop", "DllPath", txtAppPath.Text
'    SaveSetting "Schweitzer2000", "Develop", "ServerIP", txtFileServer.Text
    '--여기까지..
    SaveSetting strAppName, RegSsSet, RegK2Set, txtHospital.Text
    SaveSetting strAppName, RegSsSet, RegK3Set, txtHelpLine.Text
    
    
    SaveSetting strAppName, RegSsBld, RegK0Bld, chkUseBldFg.Value
    SaveSetting strAppName, RegSsBld, RegK1Bld, txtBldCd.Text
    SaveSetting strAppName, RegSsBld, RegK2Bld, txtBldNm.Text
    SaveSetting strAppName, RegSsBld, RegK3Bld, txtBldNo.Text
    
    SaveSetting strAppName, RegSsOpt, RegK1Opt, chkShowAtStart.Value
    SaveSetting strAppName, RegSsOpt, RegK2Opt, chkSplash.Value
    
    MsgBox "정상적으로 Registry에 등록되었습니다.", vbInformation, "메세지"
    
End Sub

Private Sub txtDbNm_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtLogin_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtPwd_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtServer_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

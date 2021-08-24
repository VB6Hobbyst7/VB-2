VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm459MAccCnt 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  '없음
   Caption         =   "미생물 통계"
   ClientHeight    =   10215
   ClientLeft      =   585
   ClientTop       =   915
   ClientWidth     =   14565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   14565
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00F4F0F2&
      Caption         =   "조회(&S)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   66
      Tag             =   "158"
      Top             =   45
      Width           =   1320
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   1770
      TabIndex        =   61
      Top             =   -30
      Width           =   6735
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   315
         Left            =   765
         TabIndex        =   62
         Top             =   195
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   85131264
         CurrentDate     =   36238
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   315
         Left            =   3780
         TabIndex        =   63
         Top             =   195
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   85131264
         CurrentDate     =   36391
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   65
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3405
         TabIndex        =   64
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Refresh"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   60
      Tag             =   "158"
      Top             =   45
      Width           =   1320
   End
   Begin VB.Frame fraInOut 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   9465
      TabIndex        =   56
      Top             =   -30
      Width           =   2355
      Begin VB.OptionButton optInOut 
         BackColor       =   &H00DBE6E6&
         Caption         =   "외래"
         Height          =   225
         Index           =   1
         Left            =   765
         TabIndex        =   58
         Top             =   255
         Width           =   765
      End
      Begin VB.OptionButton optInOut 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체"
         Height          =   225
         Index           =   2
         Left            =   1500
         TabIndex        =   57
         Top             =   255
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton optInOut 
         BackColor       =   &H00DBE6E6&
         Caption         =   "입원"
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   59
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox txtDeptCd 
      Height          =   315
      Left            =   7770
      TabIndex        =   49
      Top             =   1245
      Width           =   1050
   End
   Begin VB.TextBox txtTestCd 
      Height          =   345
      Left            =   11250
      TabIndex        =   48
      Top             =   1245
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      Caption         =   "검사항목"
      ForeColor       =   &H00864B24&
      Height          =   1110
      Left            =   11070
      TabIndex        =   42
      Top             =   600
      Width           =   3390
      Begin VB.ComboBox cboSort 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Index           =   3
         ItemData        =   "Lis459.frx":0000
         Left            =   2340
         List            =   "Lis459.frx":0016
         Style           =   2  '드롭다운 목록
         TabIndex        =   47
         Tag             =   "검 사 항 목"
         Top             =   255
         Width           =   780
      End
      Begin VB.CheckBox chkSubTot 
         BackColor       =   &H00E0E0E0&
         Caption         =   "소 계"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   4
         Left            =   1365
         TabIndex        =   46
         Top             =   285
         Width           =   750
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00E0E0E0&
         Caption         =   "전  체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   180
         TabIndex        =   45
         Top             =   300
         Width           =   1035
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1215
         MaskColor       =   &H00F4F0F2&
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   44
         Tag             =   "DeptCd"
         Top             =   645
         Width           =   285
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   345
         Left            =   1515
         TabIndex        =   43
         Top             =   645
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   609
         BackColor       =   15463405
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
   End
   Begin VB.Frame fraCond 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      Caption         =   "의뢰과"
      ForeColor       =   &H00864B24&
      Height          =   1110
      Index           =   3
      Left            =   7635
      TabIndex        =   36
      Top             =   600
      Width           =   3390
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00E0E0E0&
         Caption         =   "전  체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   150
         TabIndex        =   41
         Top             =   330
         Width           =   1035
      End
      Begin VB.CheckBox chkSubTot 
         BackColor       =   &H00E0E0E0&
         Caption         =   "소 계"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   3
         Left            =   1215
         TabIndex        =   40
         Top             =   330
         Width           =   750
      End
      Begin VB.ComboBox cboSort 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Index           =   2
         ItemData        =   "Lis459.frx":002F
         Left            =   2430
         List            =   "Lis459.frx":0045
         Style           =   2  '드롭다운 목록
         TabIndex        =   39
         Tag             =   "의뢰과"
         Top             =   285
         Width           =   780
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1200
         MaskColor       =   &H00F4F0F2&
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   38
         Tag             =   "DeptCd"
         Top             =   645
         Width           =   285
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   330
         Left            =   1515
         TabIndex        =   37
         Top             =   645
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         BackColor       =   15463405
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   210
      Left            =   1815
      TabIndex        =   34
      Top             =   1815
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   370
      BackColor       =   12648447
      ForeColor       =   8801060
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "참고 : 설정되는 기간에 따라 검색 소요시간이 길어질 수 있습니다. "
      RightGab        =   0
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFF8EE&
      Caption         =   "화면지움(&C)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   20
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame fraCond 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      Caption         =   "균  주"
      ForeColor       =   &H00864B24&
      Height          =   1110
      Index           =   2
      Left            =   3675
      TabIndex        =   13
      Top             =   600
      Width           =   3945
      Begin VB.CheckBox chkSubTot 
         BackColor       =   &H00E0E0E0&
         Caption         =   "소 계"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   2
         Left            =   1770
         TabIndex        =   24
         Top             =   360
         Width           =   750
      End
      Begin VB.ComboBox cboSort 
         BackColor       =   &H00F7FFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "Lis459.frx":005E
         Left            =   2865
         List            =   "Lis459.frx":0074
         Style           =   2  '드롭다운 목록
         TabIndex        =   22
         Tag             =   "검사장비"
         Top             =   315
         Width           =   795
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00E0E0E0&
         Caption         =   "전  체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1380
      End
      Begin VB.ComboBox cboBacSpc 
         Height          =   300
         Left            =   240
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   645
         Width           =   3435
      End
   End
   Begin VB.Frame fraCond 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      Caption         =   "검체군"
      ForeColor       =   &H00864B24&
      Height          =   1110
      Index           =   1
      Left            =   60
      TabIndex        =   12
      Top             =   600
      Width           =   3585
      Begin VB.CheckBox chkSubTot 
         BackColor       =   &H00E0E0E0&
         Caption         =   "소 계"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   1
         Left            =   1290
         TabIndex        =   23
         Top             =   345
         Width           =   750
      End
      Begin VB.ComboBox cboSort 
         BackColor       =   &H00F7FFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "Lis459.frx":008D
         Left            =   2655
         List            =   "Lis459.frx":00A3
         Style           =   2  '드롭다운 목록
         TabIndex        =   21
         Tag             =   "WorkArea"
         Top             =   300
         Width           =   780
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00E0E0E0&
         Caption         =   "전  체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   255
         TabIndex        =   18
         Top             =   345
         Width           =   1035
      End
      Begin VB.ComboBox cboBacGrp 
         Height          =   300
         Left            =   255
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   645
         Width           =   3195
      End
   End
   Begin VB.Frame frmPrgBar 
      BackColor       =   &H00AFBCC5&
      BorderStyle     =   0  '없음
      Caption         =   "                                                                                    "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000F5386&
      Height          =   1035
      Left            =   4560
      TabIndex        =   9
      Top             =   4260
      Visible         =   0   'False
      Width           =   6525
      Begin MSComctlLib.ProgressBar Prgbar 
         Height          =   225
         Left            =   60
         TabIndex        =   10
         Top             =   720
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00A9B4BA&
         BackStyle       =   0  '투명
         Caption         =   "데이터를 로드중 입니다."
         Height          =   180
         Left            =   2355
         TabIndex        =   11
         Top             =   300
         Width           =   1980
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00A9B4BA&
         Height          =   1035
         Left            =   0
         Top             =   0
         Width           =   6525
      End
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   8580
      Top             =   8595
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EAE7E3&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00EAE7E3&
      Caption         =   "출력(&P)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "132"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00EAE7E3&
      Caption         =   "Excel(&E)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "127"
      Top             =   8535
      Width           =   1320
   End
   Begin TabDlg.SSTab tabView 
      Height          =   6705
      Left            =   75
      TabIndex        =   5
      Top             =   1740
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   11827
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "테이블"
      TabPicture(0)   =   "Lis459.frx":00BC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTotalCnt"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ssDataBuf"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkAcc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "spdStat"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkGrowth"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CheckBox chkGrowth 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Growth된 건수"
         Height          =   240
         Left            =   2355
         TabIndex        =   69
         Top             =   420
         Value           =   1  '확인
         Width           =   1905
      End
      Begin FPSpread.vaSpread spdStat 
         Height          =   5775
         Left            =   315
         TabIndex        =   6
         Top             =   705
         Width           =   13380
         _Version        =   196608
         _ExtentX        =   23601
         _ExtentY        =   10186
         _StockProps     =   64
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   7
         MaxRows         =   23
         OperationMode   =   1
         ShadowColor     =   13818331
         ShadowDark      =   13818331
         SpreadDesigner  =   "Lis459.frx":00D8
         TextTip         =   4
      End
      Begin VB.CheckBox chkAcc 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수이후 모든건수"
         Height          =   240
         Left            =   330
         TabIndex        =   55
         Top             =   420
         Width           =   1905
      End
      Begin VB.ListBox lstSort 
         Height          =   240
         Left            =   -73305
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   60
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.CommandButton cmdGrpShow 
         BackColor       =   &H00D1DCD7&
         Caption         =   "Sho&w"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -62280
         Style           =   1  '그래픽
         TabIndex        =   28
         Tag             =   "158"
         Top             =   495
         Width           =   990
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   585
         Left            =   -74730
         TabIndex        =   15
         Top             =   405
         Width           =   12375
         Begin VB.ComboBox cboSeries 
            BackColor       =   &H00F7FFFF&
            Enabled         =   0   'False
            Height          =   300
            ItemData        =   "Lis459.frx":07E3
            Left            =   7875
            List            =   "Lis459.frx":07E5
            Style           =   2  '드롭다운 목록
            TabIndex        =   35
            Tag             =   "검사실"
            Top             =   195
            Width           =   780
         End
         Begin VB.CheckBox chkTable 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Show Data Table"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   10110
            TabIndex        =   33
            Top             =   240
            Width           =   2145
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Name"
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   8850
            TabIndex        =   32
            Top             =   240
            Width           =   840
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Count of"
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   6810
            TabIndex        =   31
            Top             =   240
            Width           =   1035
         End
         Begin VB.ComboBox cboXVal 
            Enabled         =   0   'False
            Height          =   300
            ItemData        =   "Lis459.frx":07E7
            Left            =   3390
            List            =   "Lis459.frx":07E9
            Style           =   2  '드롭다운 목록
            TabIndex        =   26
            Top             =   195
            Width           =   1875
         End
         Begin VB.ComboBox cboYVal 
            Enabled         =   0   'False
            Height          =   300
            ItemData        =   "Lis459.frx":07EB
            Left            =   795
            List            =   "Lis459.frx":07FB
            Style           =   2  '드롭다운 목록
            TabIndex        =   7
            Top             =   195
            Width           =   1875
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "Sort By : "
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5790
            TabIndex        =   30
            Top             =   270
            Width           =   1020
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "가로 : "
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2805
            TabIndex        =   27
            Top             =   255
            Width           =   660
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "세로 : "
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   210
            TabIndex        =   25
            Top             =   255
            Width           =   660
         End
      End
      Begin ChartfxLibCtl.ChartFX cfxStat 
         Height          =   5700
         Left            =   -74730
         TabIndex        =   8
         Top             =   975
         Width           =   13455
         _cx             =   1710382261
         _cy             =   1710368582
         Build           =   7
         TypeMask        =   101187586
         Style           =   -1179655
         RightGap        =   23
         TopGap          =   33
         AngleX          =   4
         AngleY          =   69
         View3DDepth     =   20
         MarkerShape     =   5
         MarkerSize      =   2
         Axis(0).MinorStep=   -10
         Axis(0).Max     =   90
         Axis(0).Decimals=   0
         Axis(0).TickMark=   -32767
         Axis(0).MinorTickMark=   -32766
         Axis(2).MinorStep=   -1
         Axis(2).Min     =   0
         Axis(2).Max     =   100
         RGBBk           =   14411494
         RGB2DBk         =   16777216
         nColors         =   10
         Colors          =   "Lis459.frx":081E
         TopFontMask     =   268435464
         BeginProperty TopFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BottomFontMask  =   268435464
         BeginProperty BottomFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LegendFontMask  =   268435464
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         nPts            =   10
         nSer            =   10
         NumPoint        =   10
         NumSer          =   10
         _Data_          =   "Lis459.frx":088E
      End
      Begin FPSpread.vaSpread ssDataBuf 
         Height          =   5115
         Left            =   315
         TabIndex        =   16
         Top             =   705
         Visible         =   0   'False
         Width           =   13335
         _Version        =   196608
         _ExtentX        =   23521
         _ExtentY        =   9022
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         MaxRows         =   22
         OperationMode   =   1
         ShadowColor     =   13818331
         ShadowDark      =   13818331
         SpreadDesigner  =   "Lis459.frx":08E2
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00DBE6E6&
         BackStyle       =   1  '투명하지 않음
         Height          =   6390
         Left            =   -74985
         Top             =   315
         Width           =   14370
      End
      Begin VB.Label lblTotalCnt 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   12240
         TabIndex        =   17
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label Label4 
         BackColor       =   &H00DBE6E6&
         Caption         =   "ToTal Count : "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   10455
         TabIndex        =   14
         Top             =   375
         Width           =   1725
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00DBE6E6&
         BackStyle       =   1  '투명하지 않음
         Height          =   6390
         Left            =   0
         Top             =   315
         Width           =   14385
      End
   End
   Begin VB.Frame fraCond 
      Appearance      =   0  '평면
      BackColor       =   &H00DBE6E6&
      Caption         =   "검사실"
      ForeColor       =   &H00864B24&
      Height          =   1110
      Index           =   0
      Left            =   -45
      TabIndex        =   50
      Top             =   1620
      Visible         =   0   'False
      Width           =   3255
      Begin VB.ComboBox cboBuilding 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Left            =   270
         Style           =   2  '드롭다운 목록
         TabIndex        =   54
         Top             =   660
         Width           =   2895
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전  체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   53
         Top             =   345
         Width           =   1035
      End
      Begin VB.ComboBox cboSort 
         BackColor       =   &H00F7FFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "Lis459.frx":0F00
         Left            =   2385
         List            =   "Lis459.frx":0F16
         Style           =   2  '드롭다운 목록
         TabIndex        =   52
         Tag             =   "검사실"
         Top             =   315
         Width           =   780
      End
      Begin VB.CheckBox chkSubTot 
         BackColor       =   &H00DBE6E6&
         Caption         =   "소 계"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   0
         Left            =   1590
         TabIndex        =   51
         Top             =   345
         Width           =   795
      End
   End
   Begin MedControls1.LisLabel lblCondition 
      Height          =   510
      Left            =   75
      TabIndex        =   67
      Top             =   45
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   900
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "조회기간"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   510
      Left            =   8520
      TabIndex        =   68
      Top             =   60
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   900
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "조회유형"
      Appearance      =   0
   End
End
Attribute VB_Name = "frm459MAccCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_BUILDING = 5
Private Const COL_BACGRP = 1
Private Const COL_BACSPC = 2
Private Const COL_DEPTNM = 3
Private Const COL_TESTNM = 4
Private Const COL_COUNT = 6
Private Const COL_SERIES = 7
Private Const COL_POINTS = 8

Dim rsDeptStat As Recordset
Dim rsTestStat As Recordset
Dim rsDeptTestStat  As Recordset
Dim QueryFlag As Boolean
Dim MsgFg As Boolean

Dim iPrgbarCount As Long

Dim SortKeys(6) As Integer
Dim SubTot(6) As Long
Dim ColWid(6) As Double
Dim totCnt As Long
Dim GrpColor(10) As Long

Public Event LastFormUnload()

Private objSQL  As New clsLISSqlStatistic
Private WithEvents objMyList    As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1

Private Sub cboBuilding_Click()
    'Call LoadEqpList
End Sub

Private Sub cboXVal_Click()
    cboSeries.Clear

End Sub

Private Sub cboYVal_Click()

    With cboXVal
        .Clear
        Select Case cboYVal.ListIndex
            Case 0: '전체
'                .AddItem "검사실":    .ItemData(0) = COL_BUILDING
                .AddItem "검체군":    .ItemData(1) = COL_BACGRP
                .AddItem "균  주":    .ItemData(2) = COL_BACSPC
                .AddItem "의뢰과":    .ItemData(3) = COL_DEPTNM
                .AddItem "검사항목":  .ItemData(4) = COL_TESTNM
            Case 1: '검사실
'                .AddItem "검체군":    .ItemData(0) = COL_BACGRP
                .AddItem "균  주":    .ItemData(1) = COL_BACSPC
                .AddItem "의뢰과":    .ItemData(2) = COL_DEPTNM
                .AddItem "검사항목":  .ItemData(3) = COL_TESTNM
            Case 2: '검체군
                .AddItem "검사항목":  .ItemData(0) = COL_TESTNM
            Case 3: '의뢰과
                .AddItem "검사항목":  .ItemData(0) = COL_TESTNM
        End Select
    End With
    cboSeries.Clear

End Sub

Private Sub cfxStat_LButtonUp(ByVal X As Integer, ByVal Y As Integer, nRes As Integer)
'MsgBox nRes
End Sub

Private Sub chkAll_Click(Index As Integer)
    Dim ChkValue As Boolean
    
    ChkValue = IIf(chkAll(Index).Value = 0, True, False)
    Select Case Index
    Case 0:
        cboBuilding.Enabled = ChkValue
    Case 1:
        cboBacGrp.Enabled = ChkValue
    Case 2:
        cboBacSpc.Enabled = ChkValue
    Case 3:
        txtDeptCd.Text = ""
        txtDeptCd.Enabled = ChkValue
        cmdHelpList(0).Enabled = ChkValue
        lblDeptNm.Caption = ""
    Case 4:
        txtTestCd.Text = ""
        txtTestCd.Enabled = ChkValue
        cmdHelpList(1).Enabled = ChkValue
        lblTestNm.Caption = ""
    End Select
End Sub

Private Sub chkTable_Click()
    If chkTable.Value = 1 Then
        cfxStat.DataEditor = True
        cfxStat.DataEditorObj.BkColor = &HE0E0E0
        cfxStat.DataEditorObj.Docked = 258  'TGFP_BOTTOM
        cfxStat.DataEditorObj.AutoSize = True
        cfxStat.DataEditorObj.Font = "돋움"
        cfxStat.DataEditorObj.SizeToFit
    Else
        cfxStat.DataEditor = False
    End If
End Sub

Private Sub cmdClear_Click()
    Call ClearRtn
    dtpStart.Value = Now
    dtpEnd.Value = Now
    dtpStart.SetFocus
    chkAcc.Enabled = True
    chkAcc.Value = 0
End Sub

Private Sub cmdExcel_Click()
    
    Dim strTitle As String

    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.ShowSave
    
    With spdStat
        .ReDraw = False
        .Row = 0: .Row2 = 0
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strTitle = .ClipValue
        .BlockMode = False
        .Row = 1
        .Action = ActionInsertRow
        .Row = 1: .Row2 = 1
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .ClipValue = strTitle
        .BlockMode = False
        .SaveTabFile (DlgSave.FileName)
        .Row = 1
        .Action = ActionDeleteRow
        .ReDraw = True
    End With
          
End Sub

Private Sub cmdExit_Click()

    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload

End Sub

Private Sub cmdGrpShow_Click()
    MsgBox "그래프 기능은 제공하지 않습니다.", vbExclamation, "확  인"
    
'    If cboXVal.ListIndex < 0 Or cboYVal.ListIndex < 0 Then Exit Sub
'    Call clearcfx(cfxStat)
'    DoEvents
'    Call ShowGraph
End Sub



Private Sub cmdRefresh_Click()
    Call clearcfx(cfxStat)
    Call ShowData
End Sub

Private Sub cmdStart_Click()
    Dim sStartDate As String, sEndDate As String
    
    If dtpStart.Value > dtpEnd.Value Then
        MsgBox "Duration input Error"
        Exit Sub
    End If
    
    sStartDate = Format(dtpStart.Value, CS_DateDbFormat)
    sEndDate = Format(dtpEnd.Value, CS_DateDbFormat)
    
    Screen.MousePointer = vbArrowHourglass
    QueryFlag = ReadData  ' True 이면 조회가 이루어 졌음을 의미
    Screen.MousePointer = vbDefault
    
    If QueryFlag Then
        dtpStart.Enabled = False
        dtpEnd.Enabled = False
        cmdStart.Enabled = False
        
        cmdRefresh.Enabled = True
        cmdGrpShow.Enabled = True
        cmdPrint.Enabled = True
        cmdExcel.Enabled = True
        Call cmdRefresh_Click
        chkAcc.Enabled = False
    Else
        frmPrgBar.Visible = False
        MsgBox "해당 자료가 없습니다...", vbInformation
    End If
End Sub


Private Sub dtpEnd_Validate(Cancel As Boolean)
    Call clearspdStat
    Call clearcfx(cfxStat)

End Sub

Private Sub dtpStart_Validate(Cancel As Boolean)
    Call clearspdStat
    Call clearcfx(cfxStat)


End Sub

Private Sub Form_Activate()
    MainFrm.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
'    Me.Show
    Call ClearRtn
    
    DoEvents
    
    Call LoadBuildingList
    Call LoadBacGrpList
    Call LoadBacSpcList
    
    dtpStart.Value = Format(Now, "yyyy-mm-dd")
    dtpEnd.Value = Format(Now, "yyyy-mm-dd")
     
    chkAll(0).Value = 1
    chkAll(1).Value = 1
    chkAll(2).Value = 1
    chkAll(3).Value = 1
    chkAll(4).Value = 1
    
    ColWid(1) = 10
    ColWid(2) = 36
    ColWid(3) = 14
    ColWid(4) = 30
    ColWid(5) = 0

    GrpColor(0) = &HCC99FF
    GrpColor(1) = &HFF99CC
    GrpColor(2) = &H8080FF
    GrpColor(3) = &HFFCC00
    GrpColor(4) = &HDF6A3E     '&H864B24
    GrpColor(5) = &HFFFF
    GrpColor(6) = &H808066
    GrpColor(7) = &HFF9999
    GrpColor(8) = &H663399
    GrpColor(9) = &H0
    
    MsgFg = True
    For i = 1 To cboSort.Count
        cboSort(i - 1).ListIndex = i
    Next
    MsgFg = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    QueryFlag = False
    
    Set rsDeptStat = Nothing
    Set rsTestStat = Nothing
    Set rsDeptTestStat = Nothing
    
End Sub

Private Sub clearspdStat()
    With spdStat
        .Col = 1: .COL2 = .MaxCols
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .MaxRows = 0
    End With
End Sub

Private Sub clearcfx(Ccfx As ChartFX)
    With Ccfx
        .ClearData CD_VALUES
        .ClearLegend CHART_LEGEND
    End With
End Sub

Public Sub LoadBuildingList()

    Dim i As Integer
    Dim SqlStmt As String
    Dim tmpRs As Recordset
   
    SqlStmt = "SELECT cdval1 as BuildCd, field1 as BuildNm FROM " & T_LAB032 & _
             " WHERE cdindex = '" & LC3_Buildings & "' ORDER BY BuildCd "
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
   
    cboBuilding.Clear
    For i = 1 To tmpRs.RecordCount
        cboBuilding.AddItem Trim("" & tmpRs.Fields("BuildCd").Value) & "   " & _
                            Trim("" & tmpRs.Fields("BuildNm").Value)
        tmpRs.MoveNext
    Next
   
    Set tmpRs = Nothing
   
    If cboBuilding.ListCount > 0 Then cboBuilding.ListIndex = 0 'medComboFind(cboBuilding, BuildingCd)
   
End Sub

Private Sub LoadBacGrpList()
    Dim sSqlGetBG As String
    Dim rsGetBG As Recordset
    Dim i%
    
    sSqlGetBG = " SELECT cdval1 as BGCd , field1 as BGNm " & _
                " FROM " & T_LAB032 & _
                " WHERE cdindex = '" & LC3_SGroup & "'"
    Set rsGetBG = New Recordset
    rsGetBG.Open sSqlGetBG, DBConn
    
    cboBacGrp.Clear
    For i = 1 To rsGetBG.RecordCount
        cboBacGrp.AddItem rsGetBG.Fields("BGCd").Value & "   " & _
                            rsGetBG.Fields("BGNm").Value
        rsGetBG.MoveNext
    Next i
    
    Set rsGetBG = Nothing
        
End Sub

Private Sub LoadBacSpcList()
    Dim sSqlGetBS As String
    Dim rsGetBS As Recordset
    Dim i%
    
    sSqlGetBS = " SELECT cdval1 as BSCd , text1 as BSNm " & _
                " FROM " & T_LAB032 & _
                " WHERE cdindex = '" & LC3_Microbe & "'"
    Set rsGetBS = New Recordset
    rsGetBS.Open sSqlGetBS, DBConn
    
    cboBacSpc.Clear
    For i = 1 To rsGetBS.RecordCount
        cboBacSpc.AddItem rsGetBS.Fields("BSCd").Value & "   " & _
                            rsGetBS.Fields("BSNm").Value
        rsGetBS.MoveNext
    Next i
    
    Set rsGetBS = Nothing
End Sub

Private Function ReadData() As Boolean
    Dim SqlStmt As String
    Dim RS As Recordset
    Dim i As Integer
    Dim sSqlGetBS As String
    Dim rsGetBS As Recordset
    
    Dim Inout As String
    Dim SqlDept As String
    Dim Table1 As String
    
    Dim SQLStmt1 As String
    Dim SQLStmt2 As String
    
    Table1 = T_HIS003
    
    If optInOut(0).Value Then
        Inout = "2"     '입원
    ElseIf optInOut(1).Value Then
        Inout = "1"     '외래
    Else
        Inout = "3"     '전체
    End If
    
    ReadData = False
    lblMsg = "입력된 기간동안의 검사건수를 집계하고 있습니다..."
    Prgbar.Max = 10
    Prgbar.Value = 1
    frmPrgBar.Visible = True
    DoEvents
    
    Select Case Inout
        Case "1": SqlDept = SqlDept & " AND (a.wardid is null or a.wardid=' ') "
                  SqlDept = SqlDept & "AND    " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & " "
        Case "2": SqlDept = SqlDept & " AND (a.wardid <>' ' or a.wardid is not null) "
                  SqlDept = SqlDept & "AND    " & DBJ("c." & F_DEPTCD & " =* a.wardid") & " "
        Case Else
                  SqlDept = SqlDept & "AND    " & DBJ("c." & F_DEPTCD & " =* a.deptcd") & " "
    End Select
    
    If chkAcc.Value = 0 Then
        SqlStmt = "SELECT a.buildcd as BuildCd , f.field2 as SpcGrp, e.field1 as SpcBac, "  'b.rstcd as SpcBac, "
        SqlStmt = SqlStmt & "c." & F_DEPTNM & " as DeptNm, b.testcd as TestCd, d.testnm,a.ptid, count(*) as Cnt "
        SqlStmt = SqlStmt & "FROM " & T_LAB032 & " f, " & T_LAB031 & " e, " & Table1 & " c, " & T_LAB001 & " d, "
        SqlStmt = SqlStmt & T_LAB404 & " b, " & T_LAB201 & " a "
        SqlStmt = SqlStmt & "WHERE  a.workarea >= '" & MIC_WorkArea & "' AND a.workarea <= '" & MIC_WorkArea & "' "
        SqlStmt = SqlStmt & "AND    a.rcvdt >= '" & Format(dtpStart.Value, CS_DateDbFormat) & "' "
        SqlStmt = SqlStmt & "AND    a.rcvdt <= '" & Format(dtpEnd.Value, CS_DateDbFormat) & "' "
        SqlStmt = SqlStmt & "AND    a.stscd<'7' "
        SqlStmt = SqlStmt & "AND    a.workarea = b.workarea "
        SqlStmt = SqlStmt & "AND    a.accdt = b.accdt "
        SqlStmt = SqlStmt & "AND    a.accseq = b.accseq "
        SqlStmt = SqlStmt & "  " & SqlDept
        SqlStmt = SqlStmt & "AND    d.testcd = b.testcd "
        SqlStmt = SqlStmt & "AND    d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                                "WHERE  testcd = d.testcd) "
        SqlStmt = SqlStmt & "AND    b.rsttype in ('S', 'C') "
        SqlStmt = SqlStmt & "AND   (b.senfg = '' or b.senfg is null) "
        SqlStmt = SqlStmt & "AND    e.cdindex = '" & LC2_ItemResult & "' AND e.cdval1 = b.testcd AND e.cdval2 = b.rstcd "
        SqlStmt = SqlStmt & "AND    f.cdindex = '" & LC3_Specimen & "' AND f.cdval1 = a.spccd "
        SqlStmt = SqlStmt & "GROUP BY a.buildcd, f.field2 , e.field1, c." & F_DEPTNM & ", b.testcd, d.testnm,a.ptid "
    Else
        SqlStmt = "SELECT a.buildcd as BuildCd , f.field2 as SpcGrp, '' as SpcBac, "  'b.rstcd as SpcBac, "
        SqlStmt = SqlStmt & "c." & F_DEPTNM & " as DeptNm, b.testcd as TestCd, d.testnm,a.ptid, count(*) as Cnt "
        SqlStmt = SqlStmt & "FROM " & T_LAB032 & " f,  " & Table1 & " c, " & T_LAB001 & " d, "
        SqlStmt = SqlStmt & T_LAB404 & " b, " & T_LAB201 & " a "
        SqlStmt = SqlStmt & "WHERE  a.workarea >= '" & MIC_WorkArea & "' AND a.workarea <= '" & MIC_WorkArea & "' "
        SqlStmt = SqlStmt & "AND    a.rcvdt >= '" & Format(dtpStart.Value, CS_DateDbFormat) & "' "
        SqlStmt = SqlStmt & "AND    a.rcvdt <= '" & Format(dtpEnd.Value, CS_DateDbFormat) & "' "
        SqlStmt = SqlStmt & "AND    a.stscd<'7' "
        SqlStmt = SqlStmt & "AND    a.workarea = b.workarea "
        SqlStmt = SqlStmt & "AND    a.accdt = b.accdt "
        SqlStmt = SqlStmt & "AND    a.accseq = b.accseq "
        SqlStmt = SqlStmt & "  " & SqlDept
        SqlStmt = SqlStmt & "AND    d.testcd = b.testcd "
        SqlStmt = SqlStmt & "AND    d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                                "WHERE  testcd = d.testcd) "

        SqlStmt = SqlStmt & "AND    f.cdindex = '" & LC3_Specimen & "' AND f.cdval1 = a.spccd "
        SqlStmt = SqlStmt & "GROUP BY a.buildcd, f.field2 , c." & F_DEPTNM & ", b.testcd, d.testnm,a.ptid "
    End If
    
    
    '감수성결과
    SQLStmt1 = "SELECT a.buildcd as BuildCd , f.field2 as SpcGrp, 'Growth' as SpcBac, "  'b.rstcd as SpcBac, "
    SQLStmt1 = SQLStmt1 & "c." & F_DEPTNM & " as DeptNm, b.testcd as TestCd, d.testnm,a.ptid, count(*) as Cnt "
    SQLStmt1 = SQLStmt1 & "FROM " & T_LAB032 & " f, " & Table1 & " c, " & T_LAB001 & " d, "
    SQLStmt1 = SQLStmt1 & T_LAB404 & " b, " & T_LAB201 & " a "
    SQLStmt1 = SQLStmt1 & "WHERE  a.workarea >= '" & MIC_WorkArea & "' AND a.workarea <= '" & MIC_WorkArea & "' "
    SQLStmt1 = SQLStmt1 & "AND    a.rcvdt >= '" & Format(dtpStart.Value, CS_DateDbFormat) & "' "
    SQLStmt1 = SQLStmt1 & "AND    a.rcvdt <= '" & Format(dtpEnd.Value, CS_DateDbFormat) & "' "
    SQLStmt1 = SQLStmt1 & "AND    a.stscd<'7' "
    SQLStmt1 = SQLStmt1 & "AND    a.workarea = b.workarea "
    SQLStmt1 = SQLStmt1 & "AND    a.accdt = b.accdt "
    SQLStmt1 = SQLStmt1 & "AND    a.accseq = b.accseq "
    SQLStmt1 = SQLStmt1 & "  " & SqlDept
    SQLStmt1 = SQLStmt1 & "AND    d.testcd = b.testcd "
    SQLStmt1 = SQLStmt1 & "AND    d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                            "WHERE  testcd = d.testcd) "
    SQLStmt1 = SQLStmt1 & "AND    b.rsttype in ('S', 'C') "
    
    SQLStmt1 = SQLStmt1 & "AND    b.senfg = 'Y' AND  b.senfg <> ' ' "
    SQLStmt1 = SQLStmt1 & "AND    f.cdindex = '" & LC3_Specimen & "' AND f.cdval1 = a.spccd "
    SQLStmt1 = SQLStmt1 & "AND    NOT EXISTS (SELECT * FROM " & T_LAB405 & _
                                          " WHERE workarea = b.workarea " & _
                                          " AND   accdt    = b.accdt " & _
                                          " AND   accseq   = b.accseq " & _
                                          " AND   testcd   = b.testcd) "
    SQLStmt1 = SQLStmt1 & "AND    f.cdindex = '" & LC3_Specimen & "' AND f.cdval1 = a.spccd "
    SQLStmt1 = SQLStmt1 & "GROUP BY a.buildcd, f.field2 , c." & F_DEPTNM & ", b.testcd, d.testnm,a.ptid "
    
    '감수성결과
    SQLStmt2 = "SELECT a.buildcd as BuildCd , f.field2 as SpcGrp, b.mnmcd as SpcBac, "
    SQLStmt2 = SQLStmt2 & "c." & F_DEPTNM & " as DeptNm, b.testcd as TestCd, d.testnm,a.ptid, count(*) as Cnt "
    SQLStmt2 = SQLStmt2 & "FROM " & T_LAB032 & " f, " & T_LAB032 & " e, " & Table1 & " c, " & T_LAB001 & " d, "
    SQLStmt2 = SQLStmt2 & T_LAB405 & " b, " & T_LAB201 & " a "
    SQLStmt2 = SQLStmt2 & "WHERE  a.workarea >= '" & MIC_WorkArea & "' AND a.workarea <= '" & MIC_WorkArea & "' "
    SQLStmt2 = SQLStmt2 & "AND    a.rcvdt >= '" & Format(dtpStart.Value, CS_DateDbFormat) & "' "
    SQLStmt2 = SQLStmt2 & "AND    a.rcvdt <= '" & Format(dtpEnd.Value, CS_DateDbFormat) & "' "
    SQLStmt2 = SQLStmt2 & "AND    a.stscd<'7' "
    SQLStmt2 = SQLStmt2 & "AND    a.workarea = b.workarea "
    SQLStmt2 = SQLStmt2 & "AND    a.accdt = b.accdt "
    SQLStmt2 = SQLStmt2 & "AND    a.accseq = b.accseq "
    SQLStmt2 = SQLStmt2 & "  " & SqlDept
    SQLStmt2 = SQLStmt2 & "AND    d.testcd = b.testcd "
    SQLStmt2 = SQLStmt2 & "AND    d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                            "WHERE  testcd = d.testcd) "
    SQLStmt2 = SQLStmt2 & "AND    e.cdindex = '" & LC3_Microbe & "' AND e.cdval1 = b.mnmcd "
    SQLStmt2 = SQLStmt2 & "AND    f.cdindex = '" & LC3_Specimen & "' AND f.cdval1 = a.spccd "
    SQLStmt2 = SQLStmt2 & "GROUP BY a.buildcd, f.field2 , b.mnmcd , c." & F_DEPTNM & ", b.testcd, d.testnm,a.ptid "
    
    If chkAcc.Value = 0 Then
        If chkGrowth.Value = 1 Then
            SqlStmt = SQLStmt2
        Else
            SqlStmt = SqlStmt & " union all " & SQLStmt1 & " union all " & SQLStmt2
        End If
    Else
        SqlStmt = SqlStmt
    End If

    SqlStmt = SqlStmt & "ORDER BY BuildCd, SpcGrp, SpcBac, DeptNm, testcd "
    
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    If RS.EOF Then Exit Function
    
    Screen.MousePointer = vbArrowHourglass
    Prgbar.Max = RS.RecordCount
    Prgbar.Min = 0
    Prgbar.Value = 0
    DoEvents
    
    With ssDataBuf
        .MaxRows = 0
        Do Until RS.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1
                i = medComboFind(cboBacGrp, RS.Fields("SpcGrp").Value & "")
                .Value = cboBacGrp.List(i)
            .Col = 2
                i = medComboFind(cboBacSpc, RS.Fields("SpcBac").Value & "")
                If i = -1 Then
                    .Value = RS.Fields("SpcBac").Value

                    sSqlGetBS = " SELECT field1 as BSNm " & _
                                " FROM " & T_LAB031 & _
                                " WHERE cdindex = 'C110' AND cdval1 = '" & RS.Fields("TestCd").Value & _
                                " '  AND cdval2 = '" & .Value & "'"
                    Set rsGetBS = Nothing
                    Set rsGetBS = New Recordset
                    rsGetBS.Open sSqlGetBS, DBConn

                    If Not rsGetBS.EOF Then
                        .Value = .Value & "   " & rsGetBS.Fields("BSNm").Value
                    End If

                    Set rsGetBS = Nothing
                Else
                    .Value = cboBacSpc.List(i)
                End If
            .Col = 3: .Value = "" & RS.Fields("DeptNm").Value
            .Col = 4: .Value = Format("" & RS.Fields("TestCd").Value, "!@@@@@@@@@@") & RS.Fields("TestNm").Value
            .Col = 6: .Value = RS.Fields("Cnt").Value & ""
            .Col = 7: .Value = RS.Fields("ptid").Value & ""
            Prgbar.Value = .MaxRows
            DoEvents
            RS.MoveNext
        Loop
    End With
    
    Screen.MousePointer = vbDefault
    frmPrgBar.Visible = False
    
    Set RS = Nothing
    ReadData = True
End Function

Private Sub cboSort_Click(Index As Integer)

    Dim i As Integer
    Dim j As Integer
    
    j = Val(cboSort(Index).Tag)
    If cboSort(Index).ListIndex = 0 Then
        chkSubTot(Index).Value = 0
        Exit Sub
    End If
    
    cboSort(Index).Tag = cboSort(Index).ListIndex
    SortKeys(cboSort(Index).ListIndex) = Index + 1
    
    If MsgFg Then Exit Sub
    MsgFg = True
    For i = 0 To cboSort.Count - 1
        If i <> Index Then
            If Val(cboSort(i).Tag) = cboSort(Index).ListIndex Then
                If cboSort(i).ListIndex > 0 Then
                    cboSort(i).ListIndex = j
                Else
                    cboSort(i).Tag = j
                    SortKeys(j) = i + 1
                End If
            End If
        End If
    Next
    MsgFg = False

End Sub

Private Sub ShowData()
    Dim i As Integer
    Dim K(7) As String
    Dim FirstFg As Boolean
    
    FirstFg = True
    K(1) = "": K(2) = "": K(3) = ""
    K(4) = "": K(5) = "": K(6) = ""
    K(7) = ""
    SubTot(1) = 0: SubTot(2) = 0
    SubTot(3) = 0: SubTot(4) = 0
    totCnt = 0
    
    With ssDataBuf
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .SortKey(1) = SortKeys(1)
        .SortKeyOrder(1) = SortKeyOrderAscending
        .SortKey(2) = SortKeys(2)
        .SortKeyOrder(2) = SortKeyOrderAscending
        .SortKey(3) = SortKeys(3)
        .SortKeyOrder(3) = SortKeyOrderAscending
        .SortKey(4) = SortKeys(4)
        .SortKeyOrder(4) = SortKeyOrderAscending
        .SortKey(5) = SortKeys(5)
        .SortKeyOrder(5) = SortKeyOrderAscending
        .SortKey(6) = SortKeys(6)
        .SortKeyOrder(6) = SortKeyOrderAscending
        .SortBy = SortByRow
        .Action = ActionSort
        .BlockMode = False
        
        spdStat.MaxRows = 0
        spdStat.Row = 0
        
        For i = 0 To cboSort.Count - 1
            spdStat.Col = Val(cboSort(i).Tag)
'            Debug.Print spdStat.Col
            If cboSort(i).ListIndex = 0 Then
                spdStat.ColHidden = True
            Else
                spdStat.ColHidden = False
            End If
            
           spdStat.ColWidth(spdStat.Col) = ColWid(i + 1)
        Next
        
        .Row = 0
        Call SetValue(1, K(1))
        Call SetValue(2, K(2))
        Call SetValue(3, K(3))
        Call SetValue(4, K(4))
        Call SetValue(5, K(5))
        
        For i = 1 To .MaxRows
            .Row = i
            
            If i > 0 Then
                If chkAll(0).Value = 0 Then
                    .Col = 5
                    If .Value <> cboBuilding.Text Then GoTo Skip
                End If
                If chkAll(1).Value = 0 Then
                    .Col = 1
                    If .Value <> cboBacGrp.Text Then GoTo Skip
                End If
                If chkAll(2).Value = 0 Then
                    .Col = 2
                    If .Value <> cboBacSpc.Text Then GoTo Skip
                End If
                
                If chkAll(3).Value = 0 Then
                    .Col = 3
                    If .Value <> Trim(lblDeptNm.Caption) Then GoTo Skip
                End If
                
                If chkAll(4).Value = 0 Then
                    .Col = 4
                    If Trim(Mid(.Value, 1, 9)) <> txtTestCd.Text Then GoTo Skip
                End If
            End If

            .Col = SortKeys(1)
            If K(1) <> .Value Then
                If Not FirstFg Then
                    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
                    If chkSubTot(SortKeys(3)).Value = 1 Then Call SetSubTot(3)
                    If chkSubTot(SortKeys(2)).Value = 1 Then Call SetSubTot(2)
                    If chkSubTot(SortKeys(1)).Value = 1 Then Call SetSubTot(1)
                End If
                If cboSort(SortKeys(1) - 1).ListIndex > 0 Then
                    spdStat.MaxRows = spdStat.MaxRows + 1
                    spdStat.Row = spdStat.MaxRows
                    Call SetValue(1, K(1))
                    Call SetValue(2, K(2))
                    Call SetValue(3, K(3))
                    Call SetValue(4, K(4))
                    Call SetValue(5, K(5))
                End If
            End If
            .Col = SortKeys(2)
            If K(2) <> .Value Then
                If Not FirstFg Then
                    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
                    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
                    If chkSubTot(SortKeys(3)).Value = 1 Then Call SetSubTot(3)
                    If chkSubTot(SortKeys(2)).Value = 1 Then Call SetSubTot(2)
                End If
                If cboSort(SortKeys(2) - 1).ListIndex > 0 Then
                    spdStat.MaxRows = spdStat.MaxRows + 1
                    spdStat.Row = spdStat.MaxRows
                    Call SetValue(2, K(2))
                    Call SetValue(3, K(3))
                    Call SetValue(4, K(4))
                    Call SetValue(5, K(5))
                End If
            End If
            .Col = SortKeys(3)
            If K(3) <> .Value Then
                If Not FirstFg Then
                    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
                    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
                    If chkSubTot(SortKeys(3)).Value = 1 Then Call SetSubTot(3)
                End If
                If cboSort(SortKeys(3) - 1).ListIndex > 0 Then
                    spdStat.MaxRows = spdStat.MaxRows + 1
                    spdStat.Row = spdStat.MaxRows
                    Call SetValue(3, K(3))
                    Call SetValue(4, K(4))
                    Call SetValue(5, K(5))
                End If
            End If
            .Col = SortKeys(4)
            If K(4) <> .Value Then
                If Not FirstFg Then
                    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
                    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
                End If
                If cboSort(.Col - 1).ListIndex > 0 Then
                    spdStat.MaxRows = spdStat.MaxRows + 1
                    spdStat.Row = spdStat.MaxRows
                    Call SetValue(4, K(4))
                    Call SetValue(5, K(5))
                End If
            End If

            .Col = SortKeys(5)
            If K(5) <> .Value Then
                'If chkSubTot(.Col - 1).Value = 1 Then Call SetSubTot(5)
                If cboSort(.Col - 1).ListIndex > 0 Then
                    spdStat.MaxRows = spdStat.MaxRows + 1
                    spdStat.Row = spdStat.MaxRows
                    Call SetValue(5, K(5))
                End If
            End If
'            .Col = SortKeys(1)
'            If K(1) <> .Value Then
'                If Not FirstFg Then
''                    If chkSubTot(SortKeys(4) - 1).Value = 1 Then Call SetSubTot(4)
''                    If chkSubTot(SortKeys(3) - 1).Value = 1 Then Call SetSubTot(3)
''                    If chkSubTot(SortKeys(2) - 1).Value = 1 Then Call SetSubTot(2)
''                    If chkSubTot(SortKeys(1) - 1).Value = 1 Then Call SetSubTot(1)
''                If Not FirstFg Then
'                    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
'                    If chkSubTot(SortKeys(3)).Value = 1 Then Call SetSubTot(3)
'                    If chkSubTot(SortKeys(2)).Value = 1 Then Call SetSubTot(2)
'                    If chkSubTot(SortKeys(1)).Value = 1 Then Call SetSubTot(1)
''                End If
'
'                End If
'                If cboSort(SortKeys(1) - 1).ListIndex > 0 Then
'                    spdStat.MaxRows = spdStat.MaxRows + 1
'                    spdStat.Row = spdStat.MaxRows
'                    Call SetValue(1, K(1))
'                    Call SetValue(2, K(2))
'                    Call SetValue(3, K(3))
'                    Call SetValue(4, K(4))
'                    Call SetValue(5, K(5))
'                End If
'            End If
'            .Col = SortKeys(2)
'            If K(2) <> .Value Then
'                If Not FirstFg Then
'                    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
'                    If chkSubTot(SortKeys(4) - 1).Value = 1 Then Call SetSubTot(4)
'                    If chkSubTot(SortKeys(3) - 1).Value = 1 Then Call SetSubTot(3)
'                    If chkSubTot(SortKeys(2) - 1).Value = 1 Then Call SetSubTot(2)
'                End If
'                If cboSort(SortKeys(2) - 1).ListIndex > 0 Then
'                    spdStat.MaxRows = spdStat.MaxRows + 1
'                    spdStat.Row = spdStat.MaxRows
'                    Call SetValue(2, K(2))
'                    Call SetValue(3, K(3))
'                    Call SetValue(4, K(4))
'                    Call SetValue(5, K(5))
'                End If
'            End If
'            .Col = SortKeys(3)
'            If K(3) <> .Value Then
'                If Not FirstFg Then
'                    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
'                    If chkSubTot(SortKeys(4) - 1).Value = 1 Then Call SetSubTot(4)
'                    If chkSubTot(SortKeys(3) - 1).Value = 1 Then Call SetSubTot(3)
'                End If
'                If cboSort(SortKeys(3) - 1).ListIndex > 0 Then
'                    spdStat.MaxRows = spdStat.MaxRows + 1
'                    spdStat.Row = spdStat.MaxRows
'                    Call SetValue(3, K(3))
'                    Call SetValue(4, K(4))
'                    Call SetValue(5, K(5))
'                End If
'            End If
'            .Col = SortKeys(4)
'            If K(4) <> .Value Then
'                If Not FirstFg Then
'                    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
'                    If chkSubTot(SortKeys(4) - 1).Value = 1 Then Call SetSubTot(4)
'                End If
'                If cboSort(.Col - 1).ListIndex > 0 Then
'                    spdStat.MaxRows = spdStat.MaxRows + 1
'                    spdStat.Row = spdStat.MaxRows
'                    Call SetValue(4, K(4))
'                    Call SetValue(5, K(5))
'                End If
'            End If
'
'            .Col = SortKeys(5)
'            If K(5) <> .Value Then
'                'If chkSubTot(.Col - 1).Value = 1 Then Call SetSubTot(5)
'                If cboSort(.Col - 1).ListIndex > 0 Then
'                    spdStat.MaxRows = spdStat.MaxRows + 1
'                    spdStat.Row = spdStat.MaxRows
'                    Call SetValue(5, K(5))
'                End If
'            End If
'
            .Col = 6: spdStat.Col = 6
            spdStat.Value = Val(spdStat.Value) + Val(.Value)
            
            SubTot(1) = SubTot(1) + Val(.Value)
            SubTot(2) = SubTot(2) + Val(.Value)
            SubTot(3) = SubTot(3) + Val(.Value)
            SubTot(4) = SubTot(4) + Val(.Value)
            SubTot(5) = SubTot(5) + Val(.Value)
            FirstFg = False
            
            
            totCnt = totCnt + Val(.Value)
            .Col = 7: spdStat.Col = 7
            .Value = spdStat.Value
Skip:
        Next
        
    End With
    
    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
    If chkSubTot(SortKeys(4) - 1).Value = 1 Then Call SetSubTot(4)
    If chkSubTot(SortKeys(3) - 1).Value = 1 Then Call SetSubTot(3)
    If chkSubTot(SortKeys(2) - 1).Value = 1 Then Call SetSubTot(2)
    If chkSubTot(SortKeys(1) - 1).Value = 1 Then Call SetSubTot(1)
    
    lblTotalCnt.Caption = Format(totCnt, "###,###,###,###")
    tabView.Tab = 0
    spdStat.SetFocus
End Sub


Private Sub SetValue(ByVal Col As Integer, ByRef SvVal As String)
    With ssDataBuf
        .Col = SortKeys(Col)
        spdStat.Col = Col
        spdStat.Value = .Value
        SvVal = .Value
    End With
End Sub

Private Sub SetSubTot(ByVal Col As Integer)
    Dim lngColor As Long
    With spdStat
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = Col
        .Value = "소  계"
        lngColor = .BackColor
        .Col = 6
        .Value = SubTot(Col)
        .Col = Col: .COL2 = .MaxCols
        .Row = .Row: .Row2 = .Row
        .BlockMode = True
        .BackColor = &HEEEEEE        'lngColor
        .ForeColor = &HB9602F
        .CellBorderStyle = CellBorderStyleDot
        .CellBorderType = 8  '16
        .Action = ActionSetCellBorder
        '.FontBold = True
        .BlockMode = False
        SubTot(Col) = 0
    End With
End Sub

Private Sub optSort_Click(Index As Integer)
    If Index = 0 Then
        cboSeries.Enabled = True
    Else
        cboSeries.Enabled = False
    End If
End Sub

Private Sub spdStat_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    If Row = 0 Then Exit Sub
    If Col = Val(cboSort(4).Tag) Then
        spdStat.Row = Row
        spdStat.Col = Col
        If spdStat.Value = "소  계" Or Trim(spdStat.Value) = "" Then
            ShowTip = False
            Exit Sub
        End If
        MultiLine = 1
        TipText = "  " & spdStat.Value
        TipWidth = 3000
        spdStat.TextTipDelay = 200
        'Call spdStat.SetTextTipAppearance("굴림", 9, False, False, &HEEFDF2, vbBlue)    '&H996666)
        Call spdStat.SetTextTipAppearance("Arial", 11, False, False, vbWhite, vbBlue)    '&H996666)
        ShowTip = True
    Else
        ShowTip = False
    End If
End Sub

Private Sub ShowGraph()

    Dim i As Integer, j As Integer
    Dim K(2) As String
    Dim FirstFg As Boolean
    Dim iSeries As Integer, iPoints As Integer
    Dim iSS As Integer, iPT As Integer, iCnt As Long, iVal As Long
    Dim tmpStr As String
    
    FirstFg = True
    K(1) = "": K(2) = ""
    iSeries = 0: iPoints = 0    ': iMaxValue = 100
    
    iSeries = GetCount(cboYVal.ItemData(cboYVal.ListIndex), 7)
    iPoints = GetCount(cboXVal.ItemData(cboXVal.ListIndex), 8)
    
    Call InitDraw(iSeries, iPoints)
    Call AddToSortList(iPoints)
    
    With ssDataBuf
        cfxStat.Title(CHART_TOPTIT) = cboYVal.Text & "  :  " & cboXVal.Text
        cfxStat.ClearData CD_VALUES
        cfxStat.ClearLegend CHART_LEGEND
        cfxStat.OpenDataEx COD_VALUES, iSeries, iPoints
        
        cfxStat.BottomGap = 20
        If iSeries = 1 Then
            cfxStat.FixedGap = 28 * iSeries
        Else
            cfxStat.FixedGap = 15 * iSeries
        End If
        'cfxStat.Axis(AXIS_Y).Max = iMaxValue
         
        cfxStat.Scrollable = True
        cfxStat.PointLabels = True
        cfxStat.Axis(AXIS_X).STEP = 1
        cfxStat.Axis(AXIS_X).Decimals = 0
        'cfxStat.PointLabelsFont.Bold = False
        
        Call SetSerLeg
        Call SetLegend
        Call chkTable_Click
        
        For i = 0 To iSeries - 1
            cfxStat.Series(i).Color = GrpColor(i)
        Next
            
        For i = lstSort.ListCount To 1 Step -1
            tmpStr = lstSort.List(i - 1)
            iPT = Val(medGetP(tmpStr, 3, ":"))
            
            If optSort(0).Value Then
                cfxStat.Axis(AXIS_X).Label(lstSort.ListCount - i) = medGetP(tmpStr, 2, ":")
            Else
                cfxStat.Axis(AXIS_X).Label(lstSort.ListCount - i) = medGetP(tmpStr, 1, ":")
            End If
            cfxStat.Legend(lstSort.ListCount - i) = cfxStat.Axis(AXIS_X).Label(lstSort.ListCount - i)
            
            For j = 1 To .MaxRows
                .Row = j
                .Col = COL_POINTS
                If Val(.Value) - 1 = iPT Then
                    .Col = COL_SERIES:  iSS = Val(.Value)
                    .Col = COL_COUNT:   iCnt = Val(.Value)
                    iVal = cfxStat.ValueEx(iSS - 1, lstSort.ListCount - i)
                    cfxStat.ValueEx(iSS - 1, lstSort.ListCount - i) = iVal + iCnt
                    
                    .Col = cboYVal.ItemData(cboYVal.ListIndex)
                    cfxStat.SerLeg(iSS - 1) = .Value
                End If
            Next
        Next
        
        cboSeries.Clear
        For i = 1 To iSeries
            cboSeries.AddItem cfxStat.SerLeg(i - 1)
        Next
        
        'cfxStat.Axis(AXIS_Y).Max = iMaxValue + 1
        cfxStat.CloseData COD_VALUES + COD_SCROLLLEGEND
            
    End With

End Sub


Private Sub InitDraw(ByVal nSeries As Integer, ByVal nPoints As Integer)

    Dim iMaxValue As Long
    Dim iSS As Integer, iPT As Integer, iCnt As Long, iVal As Long
    Dim i As Integer
    
    With ssDataBuf
        
        cfxStat.ClearData CD_VALUES
        cfxStat.OpenDataEx COD_VALUES, nSeries, nPoints
        
        For i = 0 To .MaxRows - 1
            .Row = i + 1
            .Col = COL_SERIES:  iSS = Val(.Value)
            .Col = COL_POINTS:  iPT = Val(.Value)
            .Col = COL_COUNT:   iCnt = Val(.Value)
            
            iVal = cfxStat.ValueEx(iSS - 1, iPT - 1)
            cfxStat.ValueEx(iSS - 1, iPT - 1) = iVal + iCnt
            
            iVal = cfxStat.ValueEx(iSS - 1, iPT - 1)
            If iMaxValue < iVal Then iMaxValue = iVal
            
            .Col = cboXVal.ItemData(cboXVal.ListIndex)
            'cfxStat.Axis(AXIS_X).Label(iPT - 1) = .Value
            cfxStat.Legend(iPT - 1) = .Value
        Next i
        
        cfxStat.Axis(AXIS_Y).Max = iMaxValue + 1
    
    End With
    
End Sub

Private Sub AddToSortList(ByVal nPoints As Integer)

    Dim tmpStr As String
    Dim i As Integer
    Dim nSeries As Integer
    
    lstSort.Clear
    If cboSeries.ListCount = 0 Or cboSeries.ListIndex < 0 Then
        nSeries = 0
    Else
        nSeries = cboSeries.ListIndex
    End If
    With cfxStat
        For i = 0 To nPoints - 1
            If optSort(0).Value Then
                tmpStr = Format(.ValueEx(nSeries, i), "0#####")
                tmpStr = tmpStr & ":" & .Legend(i)
                'tmpStr = tmpStr & ":" & .Axis(AXIS_X).Label(i)
            Else
                'tmpStr = .Axis(AXIS_X).Label(i)
                tmpStr = .Legend(i)
                tmpStr = tmpStr & ":" & Format(.ValueEx(nSeries, i), "0#####")
            End If
            tmpStr = tmpStr & ":" & Format(i, "0####")
            lstSort.AddItem tmpStr
        Next
    End With
    
End Sub

Private Sub SetSerLeg()

    With cfxStat
        .SerLegBox = True
        .SerLegBoxObj.Docked = 256  'TGFP_TOP
        .SerLegBoxObj.Height = 100
        .SerLegBoxObj.Sizeable = 3  'BAS_ALWAYS
        .SerLegBoxObj.BkColor = &HE0E0E0   '&HD2D9DB   '&HD1DCD7
        .SerLegBoxObj.SizeToFit
    End With
    
End Sub

Private Sub SetLegend()

    With cfxStat
        .LegendBox = True
        .LegendBoxObj.AutoSize = True
        .LegendBoxObj.Moveable = True
        .LegendBoxObj.Docked = 515  'TGFP_RIGHT
        .LegendBoxObj.Width = 100
        .LegendBoxObj.Sizeable = 3  'BAS_ALWAYS
        '.LegendBoxObj.FontMask = CF_SMALLFONTS
        .LegendBoxObj.BkColor = &HE0E0E0   '&HD2D9DB  '&HD1DCD7
        .LegendBoxObj.Font = "smallfonts"
        '.Axis(AXIS_X).PixPerUnit = 30
        'cfxStat.LegendBoxObj.SizeToFit
    End With
    
End Sub

Private Function GetCount(ByVal iCol As Integer, ByVal iCol1 As Integer) As Integer

    Dim i As Integer
    Dim iCount As Integer
    Dim K As String
    
    If iCol = 0 Then
        GetCount = 1
        ssDataBuf.Row = 1: ssDataBuf.Row2 = ssDataBuf.MaxRows
        ssDataBuf.Col = iCol1: ssDataBuf.COL2 = iCol1
        ssDataBuf.BlockMode = True
        ssDataBuf.Value = 1
        ssDataBuf.BlockMode = False
        Exit Function
    End If
    
    iCount = 0: K = ""
    With ssDataBuf
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .SortKey(1) = iCol
        .SortKeyOrder(1) = SortKeyOrderAscending
        .SortBy = SortByRow
        .Action = ActionSort
        .BlockMode = False
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = iCol
            If K <> .Value Then
                iCount = iCount + 1
                K = .Value
            End If
            .Col = iCol1
            .Value = iCount
        Next
    End With
    
    GetCount = iCount
    
End Function

Private Sub ClearRtn()
    
    Dim i As Integer
     
    ssDataBuf.MaxRows = 0
    spdStat.MaxRows = 0
    lblTotalCnt.Caption = ""
    
    For i = 0 To chkAll.Count - 1
        chkAll(i).Value = 1
        chkSubTot(i).Value = 0
        cboSort(i).ListIndex = i + 1
    Next
    
    Call clearcfx(cfxStat)
    
    tabView.Tab = 0
    optSort(0).Value = True
    
    dtpStart.Enabled = True
    dtpEnd.Enabled = True
    cmdStart.Enabled = True
    cmdRefresh.Enabled = False
    cmdGrpShow.Enabled = False
    cmdPrint.Enabled = False
    cmdExcel.Enabled = False

End Sub

Private Sub txtDeptCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub txtTestCd_Change()
    If Len(txtTestCd.Text) = 0 Then
        chkAll(4).Value = 1
    Else
        chkAll(4).Value = 0
    End If
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
    
End Sub

Private Sub txtDeptCd_LostFocus()

    Dim strDeptCd   As String
'    Dim objDept As clsBasisData
    Dim strDept As String
    
'    Set objDept = New clsBasisData
    
    strDeptCd = Trim(txtDeptCd.Text)
    
    If strDeptCd <> "" Then
        strDept = GetDeptNm(strDeptCd)
        If strDept = "" Then
            MsgBox "등록되지 않은 진료과입니다. 진료과코드를 확인하십시요!", vbCritical, "입력오류"
            txtDeptCd.Text = ""
            txtDeptCd.SetFocus
        Else
            lblDeptNm.Caption = strDept 'ObjLISComCode.DeptCd.Fields("deptnm")
        End If
        
        
'        If ObjLISComCode.DeptCd.Exists(strDeptCd) = False Then
'            MsgBox "등록되지 않은 진료과입니다. 진료과코드를 확인하십시요!", vbCritical, "입력오류"
'            txtDeptCd.Text = ""
'            txtDeptCd.SetFocus
'        Else
'            ObjLISComCode.DeptCd.KeyChange strDeptCd
'            lblDeptNm.Caption = ObjLISComCode.DeptCd.Fields("deptnm")
'        End If
    End If
'    Set objDept = Nothing
End Sub

Private Sub txtTestCd_LostFocus()

    Dim RS As New Recordset
    
    If Trim(txtTestCd.Text) = "" Then Exit Sub
    
    RS.Open objSQL.GetAccTest(Trim(txtTestCd.Text)), DBConn
    If RS.RecordCount > 0 Then
        lblTestNm.Caption = RS.Fields("abbrnm10").Value & ""
    Else
        MsgBox "등록되지 않은 검사코드입니다. 검사코드를 확인하십시요!", vbCritical, "입력오류"
        txtTestCd.Text = ""
        txtTestCd.SetFocus
    End If
    Set RS = Nothing
    
End Sub

Private Sub cmdHelpList_Click(Index As Integer)
'    Dim objData As clsBasisData
    
'    Set objData = New clsBasisData
    Set objMyList = New clsPopUpList
    objMyList.Connection = DBConn
    With objMyList
        Select Case Index
            Case 0:
                If optInOut(0).Value Then
                    .FormCaption = "병동 조회"
                    .ColumnHeaderText = "병동코드;병동명"
                    Call .LoadPopUp(GetSQLWardList) ', 3400, 6500) ', ObjLISComCode.WardId)
                Else
                    .FormCaption = "진료과 조회"
                    .ColumnHeaderText = "진료과코드;진료과명"
                    Call .LoadPopUp(GetSQLDeptList) ', 3400, 6500) ', ObjLISComCode.DeptCd)
                End If
                
                txtDeptCd.Text = medGetP(.SelectedString, 1, ";")
                lblDeptNm.Caption = medGetP(.SelectedString, 2, ";")
            Case 1:
                .FormCaption = "검사항목 조회"
                .ColumnHeaderText = "검사항목코드;검사명"
                Call .LoadPopUp(objSQL.GetAccTest) ', 3400, 9800)
                txtTestCd.Text = medGetP(.SelectedString, 1, ";")
                lblTestNm.Caption = medGetP(.SelectedString, 2, ";")
        End Select
    End With
'    Set objData = Nothing
    Set objMyList = Nothing
End Sub

Private Sub cmdPrint_Click()
    Call MicCount
    Exit Sub
End Sub


Private Sub MicCountHead()
    Dim strTmp  As String
    Dim ii      As Integer
    
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    lngCurYPos = 10

    Printer.FontSize = 20: Printer.FontBold = True
    Call Print_Setting("미생물 검사건수 통계", PrtLeft, LineSpace * 3, Printer.ScaleWidth - PrtLeft, "C", "C", True)
    Printer.FontSize = 9: Printer.FontBold = False
    
    strTmp = Format(dtpStart.Value, "YYYY년 MM월 DD일") & " ~ " & Format(dtpEnd.Value, "YYYY년 MM월 DD일")
    
    Call Print_Setting("조회기간 : " & strTmp, PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C", False)
    If optInOut(0).Value Then strTmp = "[ 입 원 ]"
    If optInOut(1).Value Then strTmp = "[ 외 래 ]"
    If optInOut(2).Value Then strTmp = "[ 전 체 ]"
    Call Print_Setting("조회유형 : " & strTmp, 110, LineSpace, Printer.ScaleWidth, "L", "C")
    
    strTmp = "[ 전 체 ]"
    If chkAll(1).Value = 0 Then strTmp = cboBacGrp.Text
    Call Print_Setting("검 체 군 : " & strTmp, PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C", False)
    strTmp = "[ 전 체 ]"
    If chkAll(2).Value = 0 Then strTmp = cboBacSpc.Text
    Call Print_Setting("균    주 : " & strTmp, 110, LineSpace, Printer.ScaleWidth, "L", "C")
    strTmp = "[ 전 체 ]"
    If chkAll(3).Value = 0 Then strTmp = "[ " & txtDeptCd.Text & " ] " & lblDeptNm.Caption
    Call Print_Setting("의 뢰 과 : " & strTmp, PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C", False)
    strTmp = "[ 전 체 ]"
    If chkAll(4).Value = 0 Then strTmp = "[ " & txtTestCd.Text & " ] " & lblTestNm.Caption
    Call Print_Setting("검사항목 : " & strTmp, 110, LineSpace, Printer.ScaleWidth, "L", "C")
    strTmp = Format(GetSystemDate, "YYYY년 MM월 DD일")
    Call Print_Setting("출 력 일 : " & strTmp, PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting("조회건수 : " & lblTotalCnt.Caption, 110, LineSpace, Printer.ScaleWidth, "L", "C")
    
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
    
    Call MicCountBody("검체군", "균주", "의뢰과", "검사항목", "건수")
    
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
End Sub
Private Sub MicCountBody(ByVal sGrp As String, ByVal sSUS As String, ByVal sDept As String, _
                         ByVal sTest As String, ByVal sCnt As String)
                           
    If lngCurYPos > Printer.ScaleHeight - 6 Then
        Printer.NewPage
        Call MicCountHead
    End If
   
    Call Print_Setting(sGrp, 5, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sSUS, 30, LineSpace, 50, "L", "C", False)
    Call Print_Setting(sDept, 100, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sTest, 125, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sCnt, 180, LineSpace, 30, "L", "C")
End Sub

Private Sub MicCount()
    Dim sGrp    As String
    Dim sSUS    As String
    Dim sDept   As String
    Dim sTest   As String
    Dim sCnt    As String
    
    
    Dim ii          As Integer
    
    If spdStat.DataRowCnt < 1 Then Exit Sub
    
    Call P_PrtSet
    Call MicCountHead
    
    With spdStat
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 1:   sGrp = .Value
            .Col = 2:   sSUS = .Value
            .Col = 3:   sDept = .Value
            .Col = 4:   sTest = Trim(Mid(.Value, 11))
            .Col = 6:   sCnt = .Value
            Call MicCountBody(sGrp, sSUS, sDept, sTest, sCnt)
        Next
    End With
    
    Printer.EndDoc
End Sub

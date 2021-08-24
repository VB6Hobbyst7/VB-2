VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Main 
   Caption         =   "Bioneer 인터페이스 프로그램"
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10620
   ScaleWidth      =   15120
   Begin VB.CommandButton cmdLoad 
      Caption         =   "데이타불러오기"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   133
      Top             =   930
      Width           =   2175
   End
   Begin VB.TextBox txtTestID 
      Height          =   375
      Left            =   10260
      TabIndex        =   130
      Top             =   990
      Width           =   2265
   End
   Begin VB.CommandButton le 
      Caption         =   "Command6"
      Height          =   405
      Left            =   12990
      TabIndex        =   129
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtCol 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2370
      TabIndex        =   127
      Text            =   "A"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtRow 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      TabIndex        =   126
      Text            =   "2"
      Top             =   960
      Width           =   375
   End
   Begin FPSpread.vaSpread vasOrder 
      Height          =   3795
      Left            =   6300
      TabIndex        =   87
      Top             =   5070
      Visible         =   0   'False
      Width           =   7290
      _Version        =   393216
      _ExtentX        =   12859
      _ExtentY        =   6694
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Form_Main.frx":0442
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   735
      Left            =   7590
      TabIndex        =   125
      Top             =   3990
      Width           =   1545
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "종 료"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13920
      TabIndex        =   124
      Top             =   120
      Width           =   1065
   End
   Begin VB.CommandButton cmdSear 
      Caption         =   "조 회"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12840
      TabIndex        =   123
      Top             =   120
      Width           =   1065
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "화면정리"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11760
      TabIndex        =   122
      Top             =   120
      Width           =   1065
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "코드설정"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10680
      TabIndex        =   121
      Top             =   120
      Width           =   1065
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1080
      TabIndex        =   8
      Top             =   3030
      Width           =   6990
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   390
      Left            =   5580
      TabIndex        =   86
      Top             =   3810
      Width           =   1320
   End
   Begin IF_Bioneer_칠곡경북대학병원.MDButton Command_close 
      Height          =   735
      Left            =   13890
      TabIndex        =   119
      Top             =   120
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin IF_Bioneer_칠곡경북대학병원.MDButton Command3_Search_1 
      Height          =   735
      Left            =   12810
      TabIndex        =   118
      Top             =   120
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin IF_Bioneer_칠곡경북대학병원.MDButton cmd_Clear 
      Height          =   735
      Left            =   11730
      TabIndex        =   117
      Top             =   120
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin IF_Bioneer_칠곡경북대학병원.MDButton Command_setup 
      Height          =   735
      Left            =   10650
      TabIndex        =   116
      Top             =   120
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Check1"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   600
      TabIndex        =   80
      Top             =   1470
      Width           =   195
   End
   Begin FPSpread.vaSpread vasExam 
      Height          =   7875
      Left            =   60
      TabIndex        =   6
      Top             =   1380
      Width           =   14895
      _Version        =   393216
      _ExtentX        =   26273
      _ExtentY        =   13891
      _StockProps     =   64
      ColHeaderDisplay=   0
      ColsFrozen      =   8
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   25
      RetainSelBlock  =   0   'False
      RowsFrozen      =   5
      SelectBlockOptions=   0
      SpreadDesigner  =   "Form_Main.frx":493D
      UserResize      =   2
   End
   Begin VB.CommandButton Command4 
      Caption         =   "TEST"
      Height          =   465
      Left            =   8370
      TabIndex        =   112
      Top             =   960
      Width           =   1275
   End
   Begin IF_Bioneer_칠곡경북대학병원.MDFrame MDFrame1 
      Height          =   525
      Left            =   420
      Top             =   2130
      Visible         =   0   'False
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   926
      EdgeOUTER       =   0
      EdgeSpacing     =   0
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   2820
         TabIndex        =   108
         Top             =   90
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   661
         _Version        =   393216
         Format          =   84738049
         CurrentDate     =   40218
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1110
         TabIndex        =   107
         Top             =   90
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   661
         _Version        =   393216
         Format          =   84738049
         CurrentDate     =   40218
      End
      Begin VB.TextBox txtChkCnt 
         Appearance      =   0  '평면
         BackColor       =   &H0080FFFF&
         Height          =   375
         Left            =   6660
         TabIndex        =   113
         Top             =   75
         Width           =   1425
      End
      Begin IF_Bioneer_칠곡경북대학병원.MDButton btnSearch 
         Height          =   405
         Left            =   4530
         TabIndex        =   104
         Top             =   60
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin IF_Bioneer_칠곡경북대학병원.MDButton btnWork 
         Height          =   405
         Left            =   6420
         TabIndex        =   105
         Top             =   30
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin VB.Label Label29 
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   105
         Left            =   2640
         TabIndex        =   110
         Top             =   180
         Width           =   135
      End
      Begin VB.Label Label28 
         BackStyle       =   0  '투명
         Caption         =   "접수일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   106
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Frame Frame7 
      Height          =   1035
      Left            =   90
      TabIndex        =   92
      Top             =   9210
      Width           =   9615
      Begin VB.ComboBox cboExam 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   132
         Top             =   390
         Width           =   2865
      End
      Begin VB.CommandButton cmdOrderSend 
         Caption         =   "Order 전송"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7470
         TabIndex        =   131
         Top             =   300
         Width           =   1965
      End
      Begin VB.TextBox txtPID_1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         TabIndex        =   99
         Top             =   1230
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.TextBox txtPName_1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         TabIndex        =   98
         Top             =   1620
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtPSex_1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         TabIndex        =   97
         Top             =   2010
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPAge_1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   96
         Top             =   2010
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtBarCode 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4650
         TabIndex        =   94
         Top             =   360
         Width           =   2415
      End
      Begin FPSpread.vaSpread vasRes3 
         Height          =   1545
         Left            =   4470
         TabIndex        =   93
         Top             =   1380
         Visible         =   0   'False
         Width           =   4785
         _Version        =   393216
         _ExtentX        =   8440
         _ExtentY        =   2725
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         MaxRows         =   20
         RowHeaderDisplay=   0
         ScrollBars      =   0
         SpreadDesigner  =   "Form_Main.frx":6B8C
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "환자번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   103
         Top             =   1275
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "환자이름"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   102
         Top             =   1665
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "성별"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   101
         Top             =   2055
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "나이"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   100
         Top             =   2055
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3300
         TabIndex        =   95
         Top             =   405
         Width           =   1260
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1035
      Left            =   9720
      TabIndex        =   25
      Top             =   9210
      Width           =   5265
      Begin VB.CommandButton cmd_Send 
         Caption         =   "서버 결과 전송"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   270
         TabIndex        =   114
         Top             =   270
         Width           =   1965
      End
      Begin VB.CommandButton Command1 
         Caption         =   "서버 결과 전송"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3090
         TabIndex        =   84
         Top             =   -120
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.CommandButton Command_Print 
         Caption         =   "인쇄"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3330
         TabIndex        =   83
         Top             =   270
         Width           =   1065
      End
      Begin VB.CommandButton Command_Delete 
         Caption         =   "삭제"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2250
         TabIndex        =   82
         Top             =   270
         Width           =   1065
      End
      Begin VB.CommandButton cmdQCSch 
         Caption         =   "QC조회"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   270
         TabIndex        =   66
         Top             =   3420
         Width           =   2385
      End
      Begin VB.OptionButton Option2 
         Caption         =   "내림차순"
         Height          =   255
         Left            =   300
         TabIndex        =   43
         Top             =   2430
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         Caption         =   "오른차순"
         Height          =   255
         Left            =   300
         TabIndex        =   42
         Top             =   2160
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "조   회"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   270
         TabIndex        =   41
         Top             =   2880
         Width           =   2385
      End
      Begin VB.Frame Frame6 
         Height          =   75
         Left            =   120
         TabIndex        =   40
         Top             =   2760
         Width           =   2685
      End
      Begin VB.Frame Frame5 
         Height          =   75
         Left            =   150
         TabIndex        =   37
         Top             =   2070
         Width           =   2685
      End
      Begin VB.TextBox txtStart 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   28
         Top             =   270
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtEnd 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2190
         TabIndex        =   27
         Top             =   270
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdPreRes 
         Caption         =   "환자 결과 조회"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1500
         TabIndex        =   26
         Top             =   2100
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   " - "
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1770
         TabIndex        =   36
         Top             =   330
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "순번"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   35
         Top             =   300
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin FPSpread.vaSpread vasResTemp 
      Height          =   1335
      Left            =   1170
      TabIndex        =   91
      Top             =   6300
      Visible         =   0   'False
      Width           =   2430
      _Version        =   393216
      _ExtentX        =   4286
      _ExtentY        =   2355
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Form_Main.frx":7219
   End
   Begin FPSpread.vaSpread vasOrdBuff 
      Height          =   1965
      Left            =   5505
      TabIndex        =   90
      Top             =   4140
      Visible         =   0   'False
      Width           =   4560
      _Version        =   393216
      _ExtentX        =   8043
      _ExtentY        =   3466
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Form_Main.frx":7480
   End
   Begin VB.OptionButton optOption 
      Caption         =   "Auto"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   9195
      TabIndex        =   89
      Top             =   180
      Width           =   1245
   End
   Begin VB.OptionButton optOption 
      Caption         =   "Manual"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   9195
      TabIndex        =   88
      Top             =   540
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.CommandButton Command_Search 
      Caption         =   "검색"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2250
      TabIndex        =   79
      Top             =   -90
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   5850
      TabIndex        =   77
      Text            =   "Text1"
      Top             =   5310
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Height          =   1035
      Left            =   4860
      TabIndex        =   29
      Top             =   9210
      Width           =   10125
      Begin VB.CommandButton cmdDataDel 
         Caption         =   "이전 데이타 삭제"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2730
         TabIndex        =   30
         Top             =   1050
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpDel1 
         Height          =   300
         Left            =   480
         TabIndex        =   31
         Top             =   1050
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   84738049
         CurrentDate     =   36892
      End
      Begin MSComCtl2.DTPicker dtpDel2 
         Height          =   300
         Left            =   480
         TabIndex        =   32
         Top             =   1470
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   84738049
         CurrentDate     =   36892
      End
      Begin VB.Label Label24 
         BackStyle       =   0  '투명
         Caption         =   ": Delta Check"
         Height          =   195
         Left            =   8580
         TabIndex        =   75
         Top             =   450
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label23 
         BackStyle       =   0  '투명
         Caption         =   ": Panic Low"
         Height          =   195
         Left            =   6600
         TabIndex        =   74
         Top             =   450
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label22 
         BackStyle       =   0  '투명
         Caption         =   ": Panic High"
         Height          =   195
         Left            =   4470
         TabIndex        =   73
         Top             =   450
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label21 
         BackStyle       =   0  '투명
         Caption         =   ": Negative"
         Height          =   195
         Left            =   2640
         TabIndex        =   72
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label20 
         BackStyle       =   0  '투명
         Caption         =   ": Positive"
         Height          =   195
         Left            =   630
         TabIndex        =   71
         Top             =   450
         Width           =   1305
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H000080FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  '단색
         Height          =   225
         Left            =   8070
         Top             =   435
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H000080FF&
         FillColor       =   &H0078B83C&
         FillStyle       =   0  '단색
         Height          =   225
         Left            =   6090
         Top             =   435
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H000080FF&
         FillColor       =   &H004F6CF2&
         FillStyle       =   0  '단색
         Height          =   225
         Left            =   3960
         Top             =   420
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         FillColor       =   &H0068FEFF&
         FillStyle       =   0  '단색
         Height          =   225
         Left            =   2130
         Top             =   435
         Width           =   435
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000080FF&
         FillColor       =   &H007996F6&
         FillStyle       =   0  '단색
         Height          =   225
         Left            =   120
         Top             =   420
         Width           =   435
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "까지"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2190
         TabIndex        =   39
         Top             =   1500
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "부터"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2190
         TabIndex        =   38
         Top             =   1110
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "- 삭제할 검사 일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "◈ 이전 데이타 삭제"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   120
         TabIndex        =   33
         Top             =   300
         Visible         =   0   'False
         Width           =   2445
      End
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   3045
      Left            =   8370
      TabIndex        =   1
      Top             =   3570
      Visible         =   0   'False
      Width           =   2715
      _Version        =   393216
      _ExtentX        =   4789
      _ExtentY        =   5371
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Form_Main.frx":B97B
   End
   Begin VB.TextBox txtErr 
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   10920
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   0
      Top             =   7560
      Visible         =   0   'False
      Width           =   1245
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8865
      _Version        =   65536
      _ExtentX        =   15637
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   " BIONEER  INTERFACE"
      ForeColor       =   12582912
      BackColor       =   16056319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1680
         Top             =   150
      End
      Begin VB.TextBox txtUID 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   4605
         TabIndex        =   69
         Top             =   180
         Width           =   1275
      End
      Begin VB.TextBox Text_Today 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   7215
         TabIndex        =   4
         Text            =   "2002/02/18"
         Top             =   180
         Width           =   1515
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   1080
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label lblCnt 
         Caption         =   "0"
         Height          =   285
         Left            =   3510
         TabIndex        =   115
         Top             =   60
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblUID 
         BackStyle       =   0  '투명
         Caption         =   "보고자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3735
         TabIndex        =   70
         Top             =   225
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   6105
         TabIndex        =   5
         Top             =   240
         Width           =   1020
      End
   End
   Begin FPSpread.vaSpread vasCode 
      Height          =   555
      Left            =   8910
      TabIndex        =   2
      Top             =   7530
      Visible         =   0   'False
      Width           =   1545
      _Version        =   393216
      _ExtentX        =   2725
      _ExtentY        =   979
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Form_Main.frx":BBE2
   End
   Begin VB.Frame Frame1 
      Height          =   2835
      Left            =   -720
      TabIndex        =   7
      Top             =   7380
      Width           =   14895
      Begin FPSpread.vaSpread vasList 
         Height          =   1935
         Left            =   1980
         TabIndex        =   21
         Top             =   -360
         Width           =   5055
         _Version        =   393216
         _ExtentX        =   8916
         _ExtentY        =   3413
         _StockProps     =   64
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   16777215
         MaxCols         =   20
         ShadowColor     =   16773849
         ShadowDark      =   16773849
         SpreadDesigner  =   "Form_Main.frx":10177
      End
      Begin VB.TextBox txtReceDate 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4980
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   1635
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   390
         Left            =   360
         TabIndex        =   23
         Top             =   240
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "환자 이전 결과"
         ForeColor       =   12582912
         BackColor       =   16773849
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   240
         TabIndex        =   22
         Top             =   750
         Width           =   14505
      End
      Begin VB.TextBox txtPAge 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13290
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   270
         Width           =   585
      End
      Begin VB.TextBox txtPSex 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   12360
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   270
         Width           =   585
      End
      Begin VB.TextBox txtJumin2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9690
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   270
         Width           =   1125
      End
      Begin VB.TextBox txtJumin1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   1125
      End
      Begin VB.TextBox txtPName 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5070
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   1635
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   390
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "등록번호"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtPID 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   9
         Top             =   270
         Visible         =   0   'False
         Width           =   1635
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   390
         Left            =   3750
         TabIndex        =   12
         Top             =   240
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "환자성명"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   390
         Left            =   6960
         TabIndex        =   14
         Top             =   240
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "주민번호"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   390
         Left            =   11040
         TabIndex        =   18
         Top             =   240
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "성별/나이"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "/"
         Height          =   195
         Left            =   13080
         TabIndex        =   20
         Top             =   345
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   9480
         TabIndex        =   16
         Top             =   345
         Width           =   105
      End
   End
   Begin VB.TextBox Text_ini 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8940
      TabIndex        =   68
      Top             =   7830
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "출력"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6420
      TabIndex        =   44
      Top             =   180
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdQCSet 
      Caption         =   "QC설정"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8250
      TabIndex        =   65
      Top             =   180
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSta 
      Caption         =   "통계"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7410
      TabIndex        =   67
      Top             =   180
      Visible         =   0   'False
      Width           =   1035
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   4845
      Left            =   1290
      TabIndex        =   78
      Top             =   2610
      Visible         =   0   'False
      Width           =   9345
      _Version        =   393216
      _ExtentX        =   16484
      _ExtentY        =   8546
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   13
      RowHeaderDisplay=   0
      ScrollBars      =   2
      SpreadDesigner  =   "Form_Main.frx":20E86
   End
   Begin FPSpread.vaSpread vasT 
      Height          =   915
      Left            =   5280
      TabIndex        =   85
      Top             =   7530
      Visible         =   0   'False
      Width           =   945
      _Version        =   393216
      _ExtentX        =   1667
      _ExtentY        =   1614
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Form_Main.frx":24E3B
   End
   Begin FPSpread.vaSpread vasTemp1 
      Height          =   3105
      Left            =   10530
      TabIndex        =   81
      Top             =   3870
      Width           =   3615
      _Version        =   393216
      _ExtentX        =   6376
      _ExtentY        =   5477
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Form_Main.frx":250A2
   End
   Begin Threed.SSPanel sspDetail 
      Height          =   6735
      Left            =   870
      TabIndex        =   45
      Top             =   1560
      Visible         =   0   'False
      Width           =   10785
      _Version        =   65536
      _ExtentX        =   19024
      _ExtentY        =   11880
      _StockProps     =   15
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      BevelInner      =   2
      Begin VB.CommandButton cmdUnvisible 
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   780
         TabIndex        =   64
         Top             =   5880
         Width           =   1755
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   53
         Top             =   4425
         Width           =   615
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   52
         Top             =   3930
         Width           =   615
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   51
         Top             =   3450
         Width           =   975
      End
      Begin VB.TextBox txtReceNo 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   50
         Top             =   2970
         Width           =   2085
      End
      Begin VB.TextBox txtPos 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   49
         Top             =   1845
         Width           =   1005
      End
      Begin VB.TextBox txtRack 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   48
         Top             =   1365
         Width           =   1005
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   47
         Top             =   870
         Width           =   1005
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   46
         Top             =   390
         Width           =   1725
      End
      Begin FPSpread.vaSpread vasRes1 
         Height          =   6015
         Left            =   3780
         TabIndex        =   62
         Top             =   390
         Width           =   3255
         _Version        =   393216
         _ExtentX        =   5741
         _ExtentY        =   10610
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         MaxRows         =   20
         RowHeaderDisplay=   0
         ScrollBars      =   0
         SelectBlockOptions=   0
         SpreadDesigner  =   "Form_Main.frx":25309
      End
      Begin FPSpread.vaSpread vasRes2 
         Height          =   6015
         Left            =   7200
         TabIndex        =   63
         Top             =   390
         Width           =   3255
         _Version        =   393216
         _ExtentX        =   5741
         _ExtentY        =   10610
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         MaxRows         =   20
         RowHeaderDisplay=   0
         ScrollBars      =   0
         SelectBlockOptions=   0
         SpreadDesigner  =   "Form_Main.frx":25986
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "나이"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   61
         Top             =   4470
         Width           =   450
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "성별"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   60
         Top             =   3975
         Width           =   450
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "환자이름"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   59
         Top             =   3495
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "환자번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   58
         Top             =   3015
         Width           =   900
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Pos"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   57
         Top             =   1890
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Rack"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   56
         Top             =   1410
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Seq #"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   55
         Top             =   915
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   54
         Top             =   435
         Width           =   900
      End
   End
   Begin FPSpread.vaSpread vasTmp 
      Height          =   795
      Left            =   10650
      TabIndex        =   109
      Top             =   2940
      Visible         =   0   'False
      Width           =   1185
      _Version        =   393216
      _ExtentX        =   2090
      _ExtentY        =   1402
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Form_Main.frx":26003
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   111
      Top             =   10245
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   6165
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5467
            MinWidth        =   5467
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2011-12-19"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오전 10:56"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "메디메이트 ☎(051)462-1751"
            TextSave        =   "메디메이트 ☎(051)462-1751"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   585
      Left            =   3870
      TabIndex        =   76
      Top             =   5250
      Width           =   1665
   End
   Begin VB.CommandButton cmdWork 
      Caption         =   "WorkList"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5820
      TabIndex        =   120
      Top             =   960
      Width           =   1605
   End
   Begin VB.Label Label30 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "Sample Position"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   128
      Top             =   1020
      Width           =   2055
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'2010.09.06 이상은 - 타이머로 소켓연결 확인
'타이머 - 3000

Dim sBarcode As String
Dim sSeqNo As String
Dim sDiskNo As String
Dim sPosNo As String
Dim sSampleType As String
Dim lRow As Long
Dim sOrder As String
Dim sResult As String
Dim sResDateTime As String

Const colID = 2
Const colSeqNo = 3
Const colDiskNo = 4
Const colPosNo = 5
Const colReceDate = 6
Const colReceNo = 7         '환자번호
Const colPName = 8
Const colSex = 9
Const colAge = 10
Const colJumin1 = 11
Const colJumin2 = 12
'Const colReceDate = 12     '처방일자
Const colPID = 13
Const colBun = 14
Const colSampleType = 15
Const colState = 16
Const colResult = 17

Dim colResult1      As Long
Dim colHospital     As Long
Dim lsHeaderDate    As String

Public gSpecID As String
Public gPreSpecID As String

Public gOrdRow      As Long
Public gPreRow      As Long
Public gPreMsg      As String
Public gSndState    As String

Public gENQFlag As Integer
Public gNAKFlag As Integer
Public gMsgFlag As String


Public gVersion As String
Public gDateTime As String
Public gPatFlag As Integer

Dim gResCol As Long
Dim gCurRow As Long
Dim gMaxCol As Long

Public gCup As String
Public gPos As String

Public gRecodeType As String

Dim gAllData1       As String

Dim lsIPCHBVRes As String
Dim lsIPCHBVCt As String
    
Function Get_Sample_Info(ByVal asRow As Long) As Integer
    Dim lsBarcode As String

    '샘플 환자 정보 가져오기
    lsBarcode = Trim(GetText(vasExam, asRow, colID))
    
    '환자번호, 환자이름, 주민번호, 성별, 나이
    SQL = "select a.PTNO, b.SNAME, c.AGE, c.SEX" & vbCrLf & _
          "from TWEXAM_RESULTC a, TWBAS_PATIENT b,TWEXAM_SPECMST c " & vbCrLf & _
          "Where a.SPECNO = '" & lsBarcode & "' " & vbCrLf & _
          "  and a.SUBCODE In (" & gAllExam & ") " & vbCrLf & _
          "  --and a.STATUS in ('2','3') " & vbCrLf & _
          "  and a.PTNO = b.PTNO " & vbCrLf & _
          "  and a.PTNO = c.PTNO " & vbCrLf & _
          "  and b.PTNO = c.PTNO " & vbCrLf & _
          "  and a.SPECNO =c.SPECNO" & vbCrLf & _
          " Group By a.SPECNO, a.PTNO, b.SNAME, c.AGE, c.SEX  "
    res = db_select_Col(gServer, SQL)
    If gReadBuf(0) <> "" Then
        SetText vasExam, Trim(gReadBuf(0)), asRow, colReceNo
        SetText vasExam, Trim(gReadBuf(1)), asRow, colPName
        SetText vasExam, Trim(gReadBuf(2)), asRow, colAge
        SetText vasExam, Trim(gReadBuf(3)), asRow, colSex
    Else
        SetText vasExam, 0, asRow, colAge
        SetText vasExam, "", asRow, colState
    End If
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    gReadBuf(2) = ""
    gReadBuf(3) = ""
End Function




Private Sub cmdLoad_Click()
    Dim iRow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim y As Integer
    Dim sResFlag As String
    Dim sRes As String
    
    Dim sResult As String
    
    ClearSpread vasExam

    SQL = " select barcode, seqno, diskno, posno, pid, pname, psex, page, '', '', '', '' " & _
          " from pat_res " & _
          " where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "   and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " group by barcode, seqno, diskno, posno, pid, pname, psex, page " '& vbCrLf & _
          " order by diskno,posno"
    res = db_select_Vas(gLocal, SQL, vasExam, vasExam.DataRowCnt + 1, 2)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    'vasSort vasExam, colRack, colTube
    
    For iRow = 1 To vasExam.DataRowCnt
        Select Case Trim(GetText(vasExam, iRow, colState))
        Case "B"
            SetBackColor vasExam, iRow, iRow, 1, colState, 255, 250, 205
            SetText vasExam, "결과", iRow, colState
        Case "C"
            SetBackColor vasExam, iRow, iRow, 1, colState, 202, 255, 112
            SetText vasExam, "완료", iRow, colState
        Case Else
            SetBackColor vasExam, iRow, iRow, 1, colState, 255, 255, 255
            SetText vasExam, "", iRow, colState
        End Select
    
        '결과 불러오기
        ClearSpread vasTemp
        
        SQL = " Select equipcode, result From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasExam, iRow, colID)) & "' " '& vbCrLf & _
              " And pid = '" & Trim(GetText(vasExam, iRow, colPID)) & "' "
        res = db_select_Vas(gLocal, SQL, vasTemp)
        
        For i = 1 To vasTemp.DataRowCnt
            For j = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp, i, 1)) = gArrEquip(j, 2) Then
'                    y = colResult + ((gArrEquip(j, 1)) - 1)
'
'                    k = gArrEquip(j, 1)
                    y = (gArrEquip(j, 1) - 1)

                    sResult = Trim(GetText(vasTemp, i, 2))
                    'SetText vasExam, sResult, iRow, y
                    
                    SetText vasExam, sResult, iRow, colResult + y * 4
                    SetText vasExam, sResult, iRow, colResult1 + y
                    
                    Exit For
                End If
            Next j




'            If j <= UBound(gArrEquip) Then
'                If gArrEquip(j, 2) = "HBV-HPS" Then
'                    sResFlag = ""
'                    sRes = ""
'                    If Mid(sResult, 1, 1) = ">" Or Mid(sResult, 1, 1) = "<" Then
'                        sResFlag = Mid(sResult, 1, 1)
'                        sRes = Mid(sResult, 2)
'                    Else
'                        sRes = sResult
'                    End If
'                    If IsNumeric(sRes) = True Then
'                        sRes = CCur(sRes) * 5.82
'                        sRes = Format(sRes, "###,###,###,###")
'                        If Right(sRes, 1) = "." Then
'                            sRes = Mid(sRes, 1, Len(sRes) - 1)
'                        End If
'                        vasExam.SetText vasExam.MaxCols, iRow, Trim(sResFlag & sRes)
'                    End If
'                End If
'            End If
        Next i
    Next iRow

End Sub

Private Sub cmdOrderSend_Click()
    Dim i, j, k, m, lCol As Integer
    Dim sPName_E As String
    Dim sRow As Integer
    Dim sCol As String
    
    Dim lsKit As String
    Dim db_tmp As String * 100
    Dim lsExamCode As String
    
    
'    db_tmp = ""
'    Call GetPrivateProfileString("KIT", cboExam.Text, "", db_tmp, 100, App.Path & "\Interface.ini")
'    Form_Main.Text_ini = Trim(db_tmp)
'    lsKit = Trim(Form_Main.Text_ini)

    ClearSpread vasTemp
    SQL = "SELECT Unitcode, examcode   " & CR & _
          "  From equipexam " & CR & _
          " WHERE equipno = '" & gEquip & "' " & CR & _
          "   AND Unitcode = '" & Trim(cboExam.Text) & "' "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    lsExamCode = ""
    
    For m = 1 To vasTemp.DataRowCnt
        
        If Trim(GetText(vasTemp, m, 2)) <> "" Then
            lsExamCode = lsExamCode & Trim(GetText(vasTemp, m, 2)) & ","
        End If

    Next m
    
    If Len(lsExamCode) > 0 Then
        lsExamCode = Left(lsExamCode, Len(lsExamCode) - 1)
    End If
    
    vasOrder.MaxRows = 20
    ClearSpread vasOrder
    
    gRecodeType = "Q"
    gHeader = """"
    gPatient = ""
    gOrder = ""
    gMsgEnd = ""
'    gHeader = "H|\^&||||||||||P|1" & chrCR & chrETX
    i = 0
    
    sRow = CCur(txtRow)
    sCol = txtCol
    
    For k = 1 To vasExam.DataRowCnt
        vasExam.Row = k
        vasExam.Col = 1
        
        If vasExam.Value = 1 Then
            
            
            If k = 1 Then
                i = i + 1
                If i = 8 Then
                    i = 0
                End If
                gHeader = "H|\^&" & chrCR & chrETX
                gHeader = chrSTX & CCur(i) & gHeader & CheckSum(CStr(1) & gHeader) & chrCR & chrLF
                
                vasOrder.MaxRows = vasOrder.MaxRows + 1
                
                SetText vasOrder, gHeader, vasOrder.DataRowCnt + 1, 1
                
                
            End If
            
            i = i + 1
            If i = 8 Then
                i = 0
            End If
            
            sPName_E = Trim(GetText(vasExam, k, colPName))
            sPName_E = UCase(Conv_Kor_Eng(Trim(GetText(vasExam, k, colPName))))
            
'                       P|1||PID0001||Lee^Chang Yeop^^^^|[CR]
            
            gPatient = "P|" & k & "||" & Trim(GetText(vasExam, k, colReceNo)) & "||" & sPName_E & "^^^^^|" & chrCR & chrETX
            gPatient = chrSTX & CCur(i) & gPatient & CheckSum(CStr(i) & gPatient) & chrCR & chrLF
            
            vasOrder.MaxRows = vasOrder.MaxRows + 1
            
            SetText vasOrder, gPatient, vasOrder.DataRowCnt + 1, 1
            
            i = i + 1
            If i = 8 Then
                i = 0
            End If
'                     O|1|SID0001^HBV Sample Kit^Lee Chang Yeop blood sample^B^1||^^^TID00_HBV^HBV|||||||N|[CR]
            gOrder = "O|1|" & Trim(GetText(vasExam, k, colID)) & "^" & cboExam.Text & "^" & sPName_E & "^" & sCol & "^" & sRow & "||^^^" & lsExamCode & "^" & cboExam.Text & "|||||||N|" & _
                          chrCR & chrETX
                          
            gOrder = chrSTX & CCur(i) & gOrder & CheckSum(CStr(i) & gOrder) & chrCR & chrLF
            
            vasOrder.MaxRows = vasOrder.MaxRows + 1
            
            SetText vasOrder, gOrder, vasOrder.DataRowCnt + 1, 1
            
'            sRow = sRow + 1
'            If sRow = 10 Then
'                sRow = 1
'            End If
            
            sCol = Chr(Asc(sCol) + 1)
            If sCol = "I" Then
                sRow = Chr(Asc(sRow) + 1)
                sCol = "A"
            End If
            
        End If
        vasExam.Row = k
        vasExam.Col = 1
        vasExam.Value = 0
        
    Next k
    
    i = i + 1
    If i = 8 Then
        i = 0
    End If
    gMsgEnd = "L|1" & chrCR & chrETX
    gMsgEnd = Chr(2) & CCur(i) & gMsgEnd & CheckSum(CStr(i) & gMsgEnd) & chrCR & chrLF

    vasOrder.MaxRows = vasOrder.MaxRows + 1
    
    SetText vasOrder, gMsgEnd, vasOrder.DataRowCnt + 1, 1
    
    i = i + 1
    If i = 8 Then
        i = 0
    End If
    
    vasOrder.MaxRows = vasOrder.MaxRows + 1
    
    SetText vasOrder, chrEOT, vasOrder.DataRowCnt + 1, 1
    gOrdRow = 0
    Winsock1.SendData chrENQ
    
End Sub

Private Sub cmdWork_Click()
    Dim sSch1   As String
    Dim sSch2   As String
    Dim iRow    As Integer
    Dim ii      As Integer
    Dim jj      As Integer
    
    sSch1 = Format(dtpSDate, "yymmdd") & "0001"
    sSch2 = Format(dtpEDate, "yymmdd") & "9999"
    
    txtChkCnt = 0
    
    ClearSpread vasExam

    '지원부서처리여부 - 0:저장, 1:전송, 2:접수, 3:결과
'
'    SQL = "select a.PTNO, b.SNAME, c.AGE, c.SEX" & vbCrLf & _
'          "from TWEXAM_RESULTC a, TWBAS_PATIENT b,TWEXAM_SPECMST c " & vbCrLf & _
'          "Where a.SPECNO = '" & lsBarcode & "' " & vbCrLf & _
'          "  and a.SUBCODE In (" & gAllExam & ") " & vbCrLf & _
'          "  --and a.STATUS in ('2','3') " & vbCrLf & _
'          "  and a.PTNO = b.PTNO " & vbCrLf & _
'          "  and a.PTNO = c.PTNO " & vbCrLf & _
'          "  and b.PTNO = c.PTNO " & vbCrLf & _
'          "  and a.SPECNO =c.SPECNO" & vbCrLf & _
'          " Group By a.SPECNO, a.PTNO, b.SNAME, c.AGE, c.SEX  "
    
    '외래
    SQL = " Select a.SPECNO, '', '', '', '', a.PTNO, a.SNAME, a.SEX, a.AGE, '', '', '','' " & vbCrLf
    SQL = SQL & " From TWEXAM_SPECMST a, TWEXAM_RESULTC b " & vbCrLf
    SQL = SQL & "WHERE a.SPECNO >= '" & sSch1 & "' " & vbCrLf
    SQL = SQL & "  AND a.SPECNO <= '" & sSch2 & "' " & vbCrLf
    SQL = SQL & "  AND b.SPECNO = a.SPECNO " & vbCrLf
    SQL = SQL & "  AND b.SUBCODE In (" & gAllExam & ") " & vbCrLf
    SQL = SQL & "  AND b.STATUS in ('2','3') " & vbCrLf
    SQL = SQL & "Group by a.PTNO, " & vbCrLf
    SQL = SQL & " a.SNAME, a.SEX, a.AGE, '', " & vbCrLf
    SQL = SQL & " '20' || substr(a.SPECNO, 1, 6), substr(a.SPECNO, 7, 4), a.SPECNO "
'          " Where a.SPECNO >= '" & sSch1 & "' " & vbCrLf
'          "   AND a.SPECNO <= '" & sSch2 & "' " & vbCrLf
'          " And a.WD_CODE IN ( " & gAllExam & " ) " & CR & _
'          " And a.WD_END_DEP <> '3' " & CR & _
'          " And a.WD_JDATE <> '' " & CR & _
'          " And a.WD_CANCEL = '0' " & CR & _
'          " And a.WD_CHART = b.PE_CHART " & CR & _
'          " Group By a.WD_DATE, a.WD_CHART, b.PE_SUJINJA, substring(b.PE_JUMIN,1,6), substring(b.PE_JUMIN,8,7),A.WD_STEP "
    res = db_select_Vas(gServer, SQL, vasExam, , 2)
    
   
        
    '입원
'    SQL = " Select a.ID_CHART, '', '', '', a.ID_DATE, a.ID_CHART, b.PE_SUJINJA, '', '', substring(b.PE_JUMIN,1,6), substring(b.PE_JUMIN,8,7),  A.ID_STEP, '입원' " & CR & _
'          " From ICHDAT a, PERSON b " & CR & _
'          " Where a.ID_DATE >= '" & Trim(sSch1) & "' " & CR & _
'          " And a.ID_DATE <= '" & Trim(sSch2) & "' " & CR & _
'          " And a.ID_CODE IN ( " & gAllExam & " ) " & CR & _
'          " And a.ID_END_DEP <> '3' " & CR & _
'          " And a.ID_JDATE <> '' " & CR & _
'          " And a.ID_CANCEL = '0' " & CR & _
'          " And a.ID_CHART = b.PE_CHART " & CR & _
'          " Group By a.ID_DATE,a.ID_CHART, b.PE_SUJINJA, substring(b.PE_JUMIN,1,6), substring(b.PE_JUMIN,8,7),A.ID_STEP "
'    res = db_select_Vas(gServer, SQL, vasExam, vasExam.DataRowCnt + 1, 2)
    'vasSort vasExam, 6, 7
    
    For iRow = 1 To vasExam.DataRowCnt
'        CalAgeSex Trim(GetText(vasExam, iRow, colJumin1)) & Trim(GetText(vasExam, iRow, colJumin2)), Trim(GetDateFull)
'
'        SetText vasExam, gPatGen.Sex, iRow, colSex
'        SetText vasExam, gPatGen.Age, iRow, colAge
        
        vasExam.Row = iRow
        vasExam.Col = 1
        vasExam.Value = 1
        
        '처방불러오기-2010.02.12 이상은
        ClearSpread vasOrder
        
        SQL = "SELECT SUBCODE " & vbCrLf
        SQL = SQL & "From TWEXAM_RESULTC  " & vbCrLf
        SQL = SQL & "WHERE SPECNO = '" & Trim(GetText(vasExam, iRow, 2)) & "' " & vbCrLf
        SQL = SQL & "  AND SUBCODE In (" & gAllExam & ") " & vbCrLf
        SQL = SQL & "  AND STATUS in ('2','3') "
        res = db_select_Vas(gServer, SQL, vasOrder)
        vasSort vasOrder, 1
        
        If vasOrder.DataRowCnt > 0 Then
            For ii = 1 To vasOrder.DataRowCnt
                For jj = 1 To UBound(gArrEquip)
                    If Trim(gArrEquip(jj, 3)) = Trim(GetText(vasOrder, ii, 1)) Then
                        SetText vasExam, "*", iRow, colResult + (gArrEquip(jj, 1) - 1) * 4
                    End If
                Next jj
            Next ii
        End If
    
    
'        Select Case Trim(GetText(vasExam, iRow, colBun))
'        Case "외래"
'            SQL = " Select WD_CODE, WD_Name From WCHDAT " & CR & _
'                  " Where WD_DATE = '" & Trim(GetText(vasExam, iRow, colReceDate)) & "' " & CR & _
'                  " And WD_CHART = '" & Trim(GetText(vasExam, iRow, colReceNo)) & "' " & CR & _
'                  " And WD_CODE in (" & gAllExam & ") " & CR & _
'                  " --And WD_END_DEP <> '3' "
'            res = db_select_Vas(gServer, SQL, vasOrder)
'
'            If vasOrder.DataRowCnt > 0 Then
'                For ii = 1 To vasOrder.DataRowCnt
'                    For jj = 1 To UBound(gArrEquip)
'                        If Trim(gArrEquip(jj, 3)) = Trim(GetText(vasOrder, ii, 1)) Then
'                            SetText vasExam, "*", iRow, colResult + (gArrEquip(jj, 1) - 1) * 4
'                        End If
'                    Next jj
'                Next ii
'            End If
'        Case "입원"
'            SQL = " Select ID_CODE, ID_Name From ICHDAT " & CR & _
'                  " Where ID_DATE = '" & Trim(GetText(vasExam, iRow, colReceDate)) & "' " & CR & _
'                  " And ID_CHART = '" & Trim(GetText(vasExam, iRow, colReceNo)) & "' " & CR & _
'                  " And ID_CODE in (" & gAllExam & ") " & CR & _
'                  " --And ID_END_DEP <> '3' "
'            res = db_select_Vas(gServer, SQL, vasOrder)
'
'            If vasOrder.DataRowCnt > 0 Then
'                For ii = 1 To vasOrder.DataRowCnt
'                    For jj = 1 To UBound(gArrEquip)
'                        If Trim(gArrEquip(jj, 3)) = Trim(GetText(vasOrder, ii, 1)) Then
'                            SetText vasExam, "*", iRow, colResult + (gArrEquip(jj, 1) - 1) * 4
'                        End If
'                    Next jj
'                Next ii
'            End If
'        End Select
        
        txtChkCnt = txtChkCnt + 1
    Next iRow
End Sub

Private Sub btnWork_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasExam.DataRowCnt
        vasExam.Row = lRow
        vasExam.Col = 1
        
        If vasExam.Value = 1 Then
            
        End If
    Next lRow
    
    
End Sub

Private Sub chkAll_Click()
    Dim iRow As Integer
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasExam.DataRowCnt
            vasExam.Row = iRow
            vasExam.Col = 1
            
            vasExam.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasExam.DataRowCnt
            vasExam.Row = iRow
            vasExam.Col = 1
            
            vasExam.Value = 0
        Next iRow
    End If

End Sub

Private Sub cmd_Send_Click()
    Dim iStart  As Long
    Dim iEnd    As Long
    Dim i, j    As Long
    
    Dim lsExamCode  As String
    Dim lsResult    As String
    Dim lsRefFlag   As String
    Dim lsReceNo    As String
    Dim lsReceDate  As String
    Dim lsBun       As String
    
    Dim sCnt        As String
    
    Dim sExamUID    As String
    
    Dim sResEnd     As String
    Dim lsBarcode   As String
    Dim lsPID       As String
    Dim sSeqNo      As String
    
    Dim sEditDate   As String
    Dim lsPanicFlag As String
    Dim lsDeltaFlag As String
    
'    If txtUID.Text = "" Then
'        sExamUID = InputBox("검사자를 입력하세요", "알림")
'        If sExamUID = "" Then
'            Exit Sub
'        End If
'        txtUID = sExamUID
'    End If
    
'    lblUID.Visible = True
'    txtUID.Visible = True
    
  
'    If Not Connect_Server Then
'        MsgBox "서버에 연결 되지 않았습니다 " & vbCrLf & "네트워크 확인 바랍니다", vbInformation, "알림"
'        cn_Server_Flag = False
'        Exit Sub
'    Else
'        cn_Server_Flag = True
'    End If
    
    Me.MousePointer = 11
    
    For i = 1 To vasExam.DataRowCnt
        vasExam.Row = i
        vasExam.Col = 1
        
        If vasExam.Value = 1 Then '체크된것만
            
            res = To_Server_1(i)
                    
            'Local
'            SQL = "Update pat_res set " & vbCrLf & _
'                  " resflag = 'B', examuid = '" & sExamUID & "' " & vbCrLf & _
'                  "where examdate ='" & SeperatorCls(Text_Today) & "' " & vbCrLf & _
'                  "  and barcode = '" & Trim(GetText(vasExam, i, colID)) & "' " & vbCrLf & _
'                  "  and examcode = '" & lsExamCode & "' "
'            res = SendQuery(gLocal, SQL)
'            If res = -1 Then
'                Me.MousePointer = 0
'                SaveQuery SQL
'                'Exit Sub
'            End If
                
                'db_Commit gServer
                            
            SetText vasExam, "전송", i, colState
            'SetForeColor vasExam, i, i, 1, vasExam.MaxCols, 202, 202, 202
            
            vasExam.Row = i
            vasExam.Col = 1
            vasExam.Value = 0
        
        End If
    Next i
    
    
    
'    If cn_Server_Flag Then DisConnect_Server
    
'    lblUID.Visible = False
'    txtUID.Visible = False
    Me.MousePointer = 0
End Sub

Private Sub cmdClear_Click()
Dim i As Integer

    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    'txtUID = ""
    
    If chkAll.Value = 1 Then
            For i = 1 To vasExam.DataRowCnt
                vasExam.Row = i
                vasExam.Col = 1
                
                If vasExam.Value = 1 Then
                    DeleteRow vasExam, i, i
                    i = i - 1
                End If
            Next i
            
            chkAll.Value = 0
    Else
        vasExam.Row = 1
        vasExam.Row2 = vasExam.MaxRows
        vasExam.Col = 1
        vasExam.Col2 = vasExam.MaxCols
        vasExam.BlockMode = True
        vasExam.BackColor = RGB(255, 255, 255)
        vasExam.ForeColor = RGB(0, 0, 0)
        vasExam.Action = 3
        vasExam.BlockMode = False
    End If
    
    vasPrint.BlockMode = False
    vasPrint.Row = 1
    vasPrint.Row2 = vasPrint.MaxRows
    vasPrint.Col = 1
    vasPrint.Col2 = vasPrint.MaxCols
    vasPrint.BlockMode = True
    vasPrint.BackColor = RGB(255, 255, 255)
    vasPrint.ForeColor = RGB(0, 0, 0)
    vasPrint.Action = ActionDeleteRow
    vasPrint.BlockMode = False
    
    ClearSpread vasList
   
    'ClearGraph ChartFX1
    
    '검사일자
    Text_Today = Format(CDate(Date), "yyyy/mm/dd")
    
    txtBarCode.Text = ""
    txtPID_1.Text = ""
    txtPName_1.Text = ""
    txtPSex_1.Text = ""
    txtPAge_1.Text = ""
    
    ClearSpread vasRes3
End Sub

Private Sub cmdDataDel_Click()
    Dim sDate1 As String
    Dim sDate2 As String
    
    sDate1 = Format(dtpDel1.Value, "yyyymmdd")
    sDate2 = Format(dtpDel2.Value, "yyyymmdd")
    
    If MsgBox(dtpDel1.Value & " 부터 " & dtpDel2.Value & " 까지 검사한 데이타를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
        Exit Sub
    End If
    
    Me.MousePointer = 11
    DoSleep 1
    
    db_BeginTran (gLocal)
    
    SQL = "delete From pat_res " & vbCrLf & _
          "Where examdate <= '" & sDate1 & "' and dr_wk_date >= '" & sDate2 & "' "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        db_RollBack (gLocal)
        SaveQuery SQL
        Me.MousePointer = 0
        DoSleep 1
        Exit Sub
    End If
        
    db_Commit (gLocal)
    
    Me.MousePointer = 0
    DoSleep 1
    
    MsgBox "삭제되었습니다!", vbInformation, "알림"
        
    sDate1 = ""
    SQL = "select min(examdate) From pat_res "
    res = db_select_Var(gLocal, SQL, sDate1)
    If IsDate(sDate1) Then
        dtpDel1.Value = Format(sDate1, "yyyy-mm-dd")
    End If
    
End Sub

Private Sub cmdPreRes_Click()
'    Dim i As Long
'
'    ClearSpread vasList
'
'    'ChartFX의 Data를 Clear한다
'    ChartFX1.OpenDataEx COD_VALUES, 1, 1
'    ChartFX1.Axis(AXIS_Y).Max = 1
'    ChartFX1.Axis(AXIS_Y).Min = 0
'    ChartFX1.ThisSerie = 0
'    'ChartFX.Value(1) = CHART_HIDDEN
'    ChartFX1.CloseData COD_VALUES
'
'    i = vasExam.ActiveRow
'    If i < 1 Or i > vasExam.DataRowCnt Then
'        Exit Sub
'    End If
'
'    txtReceDate = Trim(GetText(vasExam, i, colReceDate))
'    txtPID = Trim(GetText(vasExam, i, colPID))
'    txtPName = Trim(GetText(vasExam, i, colPName))
'    txtJumin1 = Trim(GetText(vasExam, i, colJumin1))
'    txtJumin2 = Trim(GetText(vasExam, i, colJumin2))
'    txtPAge = Trim(GetText(vasExam, i, colAge))
'    Select Case Trim(GetText(vasExam, i, colSex))
'    Case "1", "3", "5", "7", "9", "남", "M"
'        txtPSex = "남"
'    Case "2", "4", "6", "8", "여", "F", "W"
'        txtPSex = "여"
'    Case Else
'        txtPSex = Trim(GetText(vasExam, i, colSex))
'    End Select
'
'    Display_Data
'
'    Display_Graph
End Sub

Sub Display_Data()
    Dim sExamCode As String
    Dim i, j, k As Long
    Dim X, y As Long
    
    Dim sPreDate As String
    
    If IsNumeric(txtJumin1) = False Or IsNumeric(txtJumin2) = False Then
        Exit Sub
    End If
    
'    If Not Connect_Server Then
'        cn_Server_Flag = False
'        Exit Sub
'    Else
'        cn_Server_Flag = True
'    End If
    
    sExamCode = ""
    For i = 1 To UBound(gArrEquip)
        If sExamCode = "" Then
            sExamCode = "'" & gArrEquip(i, 3) & "'"
        Else
            sExamCode = sExamCode & ", '" & gArrEquip(i, 3) & "'"
        End If
    Next i
    
    SQL = "Select a.접수일자, b.검사코드, b.결과치, b.판정 from 접수 a, 결과_검사 b " & vbCrLf & _
          "where a.접수일자 < '" & Trim(txtReceDate.Text) & "' " & vbCrLf & _
          "  and a.주민번호1 = '" & Trim(txtJumin1.Text) & "' " & vbCrLf & _
          "  and a.주민번호2 = '" & Trim(txtJumin2.Text) & "' " & vbCrLf & _
          "  and b.검사분류 = a.검사분류 " & vbCrLf & _
          "  and b.접수번호 = a.접수번호 " & vbCrLf & _
          "  and b.접수일자 = a.접수일자 " & vbCrLf & _
          "  and b.검사코드 in (" & sExamCode & ") " & vbCrLf & _
          "Order by a.접수일자 "
    If Option1.Value = True Then
        SQL = SQL & " asc "
    End If
    If Option2.Value = True Then
        SQL = SQL & " desc "
    End If
    res = db_select_Vas(gServer, SQL, vasTemp)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
'    If cn_Server_Flag Then DisConnect_Server
    
    If vasTemp.DataRowCnt < 1 Then
        Exit Sub
    End If
    
    X = 2
    sPreDate = Trim(GetText(vasTemp, 1, 1))
    SetText vasList, sPreDate, X, 1
    For j = 1 To UBound(gArrEquip)
        If Trim(GetText(vasTemp, 1, 2)) = gArrEquip(j, 3) Then
            y = gArrEquip(j, 1) + 1
            Exit For
        End If
    Next j
    SetText vasList, Trim(GetText(vasTemp, 1, 3)), X, y
    Select Case Trim(GetText(vasTemp, 1, 4))
    Case "H"
        SetBackColor vasList, X, X, y, y, 255, 149, 149
    Case "L"
        SetBackColor vasList, X, X, y, y, 149, 149, 255
    Case Else
        SetBackColor vasList, X, X, y, y, 255, 255, 255
    End Select
    For i = 2 To vasTemp.DataRowCnt
        If Trim(GetText(vasTemp, i, 1)) = sPreDate Then
            For j = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp, i, 2)) = gArrEquip(j, 3) Then
                    y = gArrEquip(j, 1) + 1
                    Exit For
                End If
            Next j
            SetText vasList, Trim(GetText(vasTemp, i, 3)), X, y
            Select Case Trim(GetText(vasTemp, i, 4))
            Case "H"
                SetBackColor vasList, X, X, y, y, 255, 149, 149
            Case "L"
                SetBackColor vasList, X, X, y, y, 149, 149, 255
            Case Else
                SetBackColor vasList, X, X, y, y, 255, 255, 255
            End Select
        Else
            X = X + 1
            
            sPreDate = Trim(GetText(vasTemp, i, 1))
            SetText vasList, sPreDate, X, 1
            For j = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp, i, 2)) = gArrEquip(j, 3) Then
                    y = gArrEquip(j, 1) + 1
                    Exit For
                End If
            Next j
            SetText vasList, Trim(GetText(vasTemp, i, 3)), X, y
            Select Case Trim(GetText(vasTemp, i, 4))
            Case "H"
                SetBackColor vasList, X, X, y, y, 255, 149, 149
            Case "L"
                SetBackColor vasList, X, X, y, y, 149, 149, 255
            Case Else
                SetBackColor vasList, X, X, y, y, 255, 255, 255
            End Select
        End If
    Next i
    
    For i = 1 To UBound(gArrEquip)
        vasList.Row = 1
        vasList.Col = i + 1
        vasList.Value = 1
    Next i
End Sub

Sub Display_Graph()
'    Dim i, j, k, l, m, n As Long
'    Dim DataMax, DataMin
'
'    If vasList.DataRowCnt < 1 Then
'        Exit Sub
'    End If
'
'    k = 0
'    For j = 2 To vasList.MaxCols
'        vasList.Row = 1
'        vasList.Col = j
'        If vasList.Value = 1 Then
'            k = k + 1
'        End If
'    Next j
'
'    m = 0
'    For i = 2 To vasList.DataRowCnt
'        n = -1
'        For j = 2 To vasList.MaxCols
'            vasList.Row = 1
'            vasList.Col = j
'            If vasList.Value = 1 Then
'                If Trim(GetText(vasList, i, j)) <> "" Then
'                    n = 1
'                    Exit For
'                End If
'            End If
'        Next j
'        If n = 1 Then
'            m = m + 1
'        End If
'    Next i
'
'    ChartFX1.Fonts(CHART_TOPTIT) = CF_BOLD Or CF_ITALIC Or 12
'    ChartFX1.RGBFont(CHART_TOPTIT) = RGB(0, 0, 0)
'
'    ChartFX1.OpenDataEx COD_VALUES Or COD_REMOVE, k, m  '/Graph 그리기
'    'ChartFX1.OpenDataEx COD_CONSTANTS, 2, 0             '/Constant Line 그리기
'
'    ChartFX1.ThisSerie = 0
'
'    For i = 0 To vasList.DataRowCnt - 1
'        ChartFX1.Value(i) = Trim(GetText(vasList, i + 1, 2))
'        If DataMax < ChartFX1.Value(i) Then
'            DataMax = ChartFX1.Value(i)
'        End If
'        If DataMin > ChartFX1.Value(i) Then
'            DataMin = ChartFX1.Value(i)
'        End If
'    Next i
'
'    m = -1
'    For i = 2 To vasList.DataRowCnt
'        n = -1
'        l = 0
'        For j = 2 To vasList.MaxCols
'            vasList.Row = 1
'            vasList.Col = j
'            If vasList.Value = 1 Then
'                n = n + 1
'                ChartFX1.ThisSerie = n
'                If Trim(GetText(vasList, i, j)) <> "" Then
'                    l = l + 1
'                    If l = 1 Then
'                        m = m + 1
'                    End If
'                    ChartFX1.Value(m) = Trim(GetText(vasList, i, j))
'
'                    If DataMax < ChartFX1.Value(m) Then
'                        DataMax = ChartFX1.Value(m)
'                    End If
'                    If DataMin > ChartFX1.Value(m) Then
'                        DataMin = ChartFX1.Value(m)
'                    End If
'                End If
'            End If
'        Next j
'    Next i
'
'    ChartFX1.Axis(AXIS_Y).Max = DataMax + 0.1
'    ChartFX1.Axis(AXIS_Y).Min = DataMin - 0.1
'
'    ChartFX1.Axis(AXIS_Y).AutoScale = True
'    ChartFX1.CloseData COD_VALUES
End Sub

Private Sub cmdQC_Click()

End Sub

Private Sub cmdPrint_Click()
    Dim llRow As Long
    Dim llCol As Long
    
    Dim sHead As String
    Dim sFoot As String
    Dim sCurDate As String
    
    Me.MousePointer = 11
    
    ClearSpread vasPrint
    For llRow = 1 To vasExam.DataRowCnt
        For llCol = 2 To vasExam.MaxCols
            SetText vasPrint, Trim(GetText(vasExam, llRow, llCol)), llRow, llCol
        Next llCol
'        '결과 셀 색깔 변화=========================================================
'        Select Case Trim(GetText(vasExam, llRow, colRefFlag))
'        Case "H"
'            SetForeColor vasPrint, llRow, llRow, colRefFlag - 1, colRefFlag - 1, 255, 0, 0
'        Case "L"
'            SetForeColor vasPrint, llRow, llRow, colRefFlag - 1, colRefFlag - 1, 0, 0, 255
'        Case Else
'            SetForeColor vasPrint, llRow, llRow, colRefFlag - 1, colRefFlag - 1, 0, 0, 0
'        End Select
'
'        Select Case Trim(GetText(vasExam, llRow, colPanicFlag))
'        Case "H"
'            SetForeColor vasPrint, llRow, llRow, colPanicFlag - 1, colPanicFlag - 1, 255, 0, 0
'        Case "L"
'            SetForeColor vasPrint, llRow, llRow, colPanicFlag - 1, colPanicFlag - 1, 0, 0, 255
'        Case Else
'            SetForeColor vasPrint, llRow, llRow, colPanicFlag - 1, colPanicFlag - 1, 0, 0, 0
'        End Select
'
'        Select Case Trim(GetText(vasExam, llRow, colDeltaFlag))
'        Case Is <> ""
'            SetForeColor vasPrint, llRow, llRow, colDeltaFlag - 1, colDeltaFlag - 1, 255, 0, 0
'        Case Else
'            SetForeColor vasPrint, llRow, llRow, colDeltaFlag - 1, colDeltaFlag - 1, 0, 0, 0
'        End Select
    Next llRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Me.MousePointer = 0
        Exit Sub
    End If

    sCurDate = GetDateFull
    

    sHead = "검사 일자 : " & Trim(Text_Today.Text)
    
    sHead = "/fn""궁서체"" /fz""15"" /fb1 /fi0 /fu0 " & "/c" & "▣ Elecsys 검사 결과 ▣" & "/n/n " & _
                "/fn""굴림체"" /fz""11"" /fb0 /fi0 /fu0 " & "/c" & sHead & "/n" & _
                "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/l검사자 : " & Trim(txtUID) & "/rPage /p" & "/n"
    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & "SCL 부산"
    vasPrint.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
    vasPrint.PrintAbortMsg = "인쇄중 입니다 ..."
    vasPrint.PrintJobName = "Elecsys 검사 현황"
    vasPrint.PrintHeader = sHead
    vasPrint.PrintFooter = sFoot
    vasPrint.PrintMarginTop = 720
    vasPrint.PrintMarginBottom = 720
'현재 SS가 비대칭으로 출력함
'    vasprint.PrintMarginLeft = 720
    vasPrint.PrintMarginLeft = 0
    vasPrint.PrintMarginRight = 0
    
    vasPrint.PrintColor = True
    vasPrint.PrintGrid = True
'Set printing range
    vasPrint.PrintType = 0  'SS_PRINT_ALL(default)

    vasPrint.PrintShadows = True

    vasPrint.Action = 13 'SS_ACTION_PRINT
    
    Me.MousePointer = 0
End Sub

Private Sub cmdQCSch_Click()
    'frmQCResSch.Show
End Sub

Private Sub cmdQCSet_Click()
    'frmQCSet.Show
End Sub

Private Sub cmdSearch_Click()
    frmSearch.Show
End Sub

Private Sub cmdSta_Click()
    frmSta.Show
End Sub

Private Sub cmdUnvisible_Click()
    txtID = ""
    txtSeq = ""
    txtRack = ""
    txtPos = ""
    
    txtReceNo = ""
    txtName = ""
    txtSex = ""
    txtAge = ""
    
    ClearSpread vasRes1, 1, 0
    ClearSpread vasRes1, 2, 0
    
    glRow = -1
    sspDetail.Visible = False
End Sub

Private Sub cmdClose_Click()
    'MSComm1.PortOpen = False
    Unload Me
End Sub

Private Sub Command_config_Click()
    Form_config.Show 1
End Sub

Private Sub Command_Delete_Click()
'로컬 데이터와 Spread도 삭제되게
    Dim i As Long
    Dim j As Long
    Dim lsPID As String
    
'    If Not Connect_Server Then
'        MsgBox "서버에 연결 되지 않았습니다 " & vbCrLf & "네트워크 확인 바랍니다", vbInformation, "알림"
'        cn_Server_Flag = False
'        Exit Sub
'    Else
'        cn_Server_Flag = True
'    End If
    
    Me.MousePointer = 11
    
    If vasExam.Row < 1 Then
        Me.MousePointer = 0
        Exit Sub
    End If
    
    db_BeginTran gLocal
    
    For i = 1 To vasExam.DataRowCnt
        vasExam.Row = i
        vasExam.Col = 1
            
        If vasExam.Value = 1 Then '체크된것만
            'Local 지우기
            lsPID = Trim(GetText(vasExam, i, colID))
            
            SQL = " Delete From pat_res " & CR & _
                  " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & CR & _
                  " And barcode = '" & Trim(lsPID) & "'"
            
            res = SendQuery(gLocal, SQL)
            
            If res = -1 Then
                SaveQuery SQL
                db_RollBack gLocal
                Exit Sub
            End If
            
            db_Commit gLocal
            
            vasExam.Row = i
            vasExam.Col = 1
            vasExam.Row2 = i
            vasExam.Col2 = vasExam.DataColCnt
            vasExam.BlockMode = True
            vasExam.Action = 3
            vasExam.BlockMode = False
        End If
                    
    Next i
    chkAll.Value = 0
    
'    '스프레드 지우기
'    For j = 1 To vasExam.DataRowCnt
'        vasExam.Row = j
'        vasExam.Col = 1
'
'        If vasExam.Value = 1 Then
'            DeleteRow vasExam, j, j
'            j = 1
'        End If
'
'    Next j
    
'    If cn_Server_Flag Then DisConnect_Server
    
    Me.MousePointer = 0

End Sub

Private Sub Command_Print_Click()
    Dim llRow As Long
    Dim llCol As Long
    Dim iRow As Integer
    Dim i As Integer
    Dim j As Integer
    
    Dim sBarcode As String
    
    Dim sCurDate As String
    
    ClearSpread vasPrint
    
    sCurDate = Trim(GetDateFull)
    
    If MsgBox("출력하시겠습니까", vbInformation + vbYesNo, "알림") = vbNo Then
        Exit Sub
    End If
    
    Me.MousePointer = 11
    
    i = 1
    j = 1
    
    For iRow = 1 To vasExam.DataRowCnt
        vasExam.Row = iRow
        vasExam.Col = 1
        
        If vasExam.Value = 1 Then '체크된것만
            sBarcode = Trim(GetText(vasExam, iRow, 2))
            SetText vasPrint, sBarcode, i, 1
            
            SetText vasPrint, Trim(GetText(vasExam, iRow, 3)), i, 2     'SeqNo
            SetText vasPrint, Trim(GetText(vasExam, iRow, 4)), i, 3     'Rack
            SetText vasPrint, Trim(GetText(vasExam, iRow, 5)), i, 4     'Pos
            SetText vasPrint, Trim(GetText(vasExam, iRow, 6)), i, 5     'PID
            SetText vasPrint, Trim(GetText(vasExam, iRow, 7)), i, 6     'PName
            SetText vasPrint, Trim(GetText(vasExam, iRow, 8)), i, 7     'Sex
            SetText vasPrint, Trim(GetText(vasExam, iRow, 9)), i, 8     'Age
            
            For llCol = colResult To colResult1 - 1 Step 4
                If Trim(GetText(vasExam, iRow, llCol)) <> "" Then
                    '검사명
                    SetText vasPrint, Trim(GetText(vasExam, 0, llCol)), j, 9
                    '검사결과
                    SetText vasPrint, Trim(GetText(vasExam, iRow, llCol)), j, 10
                    '판정
                    SetText vasPrint, Trim(GetText(vasExam, iRow, llCol + 1)), j, 11
                    'Panic
                    SetText vasPrint, Trim(GetText(vasExam, iRow, llCol + 2)), j, 12
                    'Delta
                    SetText vasPrint, Trim(GetText(vasExam, iRow, llCol + 3)), j, 13
                    
                    j = j + 1
                End If
                    
            Next llCol
            
            i = j
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Me.MousePointer = 0
        Exit Sub
    End If
    
    vasPrint.PrintOrientation = 1
    
    vasPrint.PrintAbortMsg = "인쇄중입니다..."
    vasPrint.PrintJobName = "Exicycler"
    
    vasPrint.PrintFooter = sCurDate
    
    vasPrint.PrintMarginTop = 720
    vasPrint.PrintMarginBottom = 720
    
    vasPrint.PrintMarginLeft = 0
    vasPrint.PrintMarginRight = 0
    
    vasPrint.PrintColor = True
    vasPrint.PrintGrid = True
    
    'Set printing range
    vasPrint.PrintType = 0  'SS_PRINT_ALL(default)
    
    vasPrint.PrintShadows = False
    
    vasPrint.Action = 13

    Me.MousePointer = 0
End Sub

Private Sub Command_Search_Click()
'    frmSearch.Show 1
'서버에 전송안된것 불러오기
    
    Dim i, j, k, m As Long
    Dim X, y As Long
    Dim sPreID As String
    Dim sBarcode As String
    Dim sResult, sResult1 As String
    Dim iPos As Integer
    
    Dim sExecDate As String
    Dim sExecDate1 As String
    
    ClearSpread vasTemp1

    sResult = ""
    sResult1 = ""
    
    sExecDate = Format(Trim(Text_Today), "yyyymmdd")
    sExecDate1 = Format(DateAdd("d", "1", Trim(Text_Today)), "yyyymmdd")
    
'Local이 아닌 서버에서 워크리스트 조회하게
    SQL = "Select barcode, seqno, diskno, posno, pid, " & _
          "       pname, psex, page, jumin1, jumin2, " & _
          "       recedate, pid, '', resflag, examcode, " & _
          "       result, refflag, panicflag, deltaflag, examuid " & vbCrLf & _
          "From pat_res " & CR & _
          "Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & CR & _
          "And ResFlag <> 'B' " & CR & _
          "And SampleType <> 'Q' Order by seqno "

'2005/07/26 이상은 - 코드설정에 등록된 항목만 불러오기
'    SQL = " Select distinct(W.SPCID), '', '', '', W.PATNO, " & CR & _
'          "        a.PATNAME, a.SEX, to_char(trunc(months_between(SYSDATE,A.BIRTDATE)/12)), '', '', " & CR & _
'          "        '', W.PATNO, '', '', '', " & CR & _
'          "        '', '', '', '', '', W.WORKNO " & CR & _
'          " From APPATBAT A, SLXWORKT W " & CR & _
'          " Where W.EXECDATE >= to_date('" & Trim(sExecDate) & "','yyyymmdd') " & CR & _
'          " And W.EXECDATE <  to_date('" & Trim(sExecDate1) & "', 'yyyymmdd') " & CR & _
'          " And W.ROOMCODE = 'LEB' " & CR & _
'          " And W.WORKCODE = 'LHCM' " & CR & _
'          " And W.PROCSTAT = 'E' " & CR & _
'          " And W.WORKNO between to_number('1') And to_number('9999') " & CR & _
'          " And W.EXAMCODE IN (" & gAllExam & ") " & CR & _
'          " And A.PATNO = W.PATNO " & CR & _
'          " Order by W.WORKNO "
'     Debug.Print SQL

    res = db_select_Vas(gLocal, SQL, vasTemp1)
    If vasTemp1.DataRowCnt < 1 Then
        cmdClear_Click
        Exit Sub
    End If

    X = 1
    sPreID = Trim(GetText(vasTemp1, 1, 1))
    SetText vasExam, sPreID, 1, 2
    For j = 1 To 14
        SetText vasExam, Trim(GetText(vasTemp1, 1, j)), X, j + 1
    Next j

    For i = 2 To vasTemp1.DataRowCnt
        If Trim(GetText(vasTemp1, i, 1)) = sPreID Then

        Else
            X = X + 1

            If X > vasExam.MaxRows Then
                vasExam.MaxRows = X
            End If

            sPreID = Trim(GetText(vasTemp1, i, 1))
            SetText vasExam, sPreID, i, 2
            For j = 1 To 14
                SetText vasExam, Trim(GetText(vasTemp1, i, j)), X, j + 1
            Next j
        End If
    Next i
        
    For j = 1 To vasExam.DataRowCnt
        ClearSpread vasTemp1
        
        SQL = " Select W.ExamCode From SLXWORKT W " & CR & _
              " Where W.SPCID = '" & Trim(GetText(vasExam, j, 2)) & "' " & CR & _
              " And W.EXECDATE >= to_date('" & Trim(sExecDate) & "','yyyymmdd') " & CR & _
              " And W.EXECDATE <  to_date('" & Trim(sExecDate1) & "', 'yyyymmdd') " & CR & _
              " And W.ROOMCODE = 'LBB' " & CR & _
              " And W.WORKCODE = 'LBCOA' " & CR & _
              " And W.PROCSTAT = 'E' " & CR & _
              " And W.WORKNO between to_number('1') And to_number('9999') "
'              Debug.Print SQL
        res = db_select_Vas(gServer, SQL, vasTemp1)
        
        For m = 1 To vasTemp1.DataRowCnt
            For k = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp1, m, 1)) = gArrEquip(k, 3) Then
                    y = 16 + gArrEquip(k, 1) * 4 - 3
                    SetText vasExam, "*", j, y
                    Exit For
                End If
            Next k
        Next m
    Next j
    
End Sub

Private Sub cmdsetup_Click()
    ComState = False
    Form_OrderCode.Show 1
    ComState = True
    GetExamCode
End Sub

Private Sub Command1_Click()
    Dim iStart As Long
    Dim iEnd As Long
    Dim i, j As Long
    
    Dim lsExamCode As String
    Dim lsResult As String
    Dim lsRefFlag As String
    Dim lsReceNo As String
    Dim lsReceDate As String
    Dim lsBun As String
    
    Dim sCnt As String
    
    Dim sExamUID As String
    
    Dim sResEnd As String
    Dim lsPID As String
    Dim sSeqNo As String
    
    Dim sEditDate As String
    Dim lsPanicFlag As String
    Dim lsDeltaFlag As String
    
    Dim iPos As Integer

    
    sEditDate = Format(GetDateFull, "yyyy-mm-dd hh:mm:ss")
    
    If txtUID.Text = "" Then
        sExamUID = InputBox("검사자를 입력하세요", "알림")
        If sExamUID = "" Then
            Exit Sub
        End If
        txtUID = sExamUID
    Else
        sExamUID = txtUID
    End If
    
'    If Not Connect_Server Then
'        MsgBox "서버에 연결 되지 않았습니다 " & vbCrLf & "네트워크 확인 바랍니다", vbInformation, "알림"
'        cn_Server_Flag = False
'        Exit Sub
'    Else
'        cn_Server_Flag = True
'    End If
    
    Me.MousePointer = 11
    
    'db_BeginTran gServer
    
    For i = 1 To vasExam.DataRowCnt
        vasExam.Row = i
        vasExam.Col = 1
        
        If vasExam.Value = 1 Then '체크된것만
            lsPID = Trim(GetText(vasExam, i, colReceNo))

            If lsPID <> "" Then
                For j = 1 To UBound(gArrEquip)
                    lsResult = Trim(GetText(vasExam, i, colResult + (gArrEquip(j, 1) - 1) * 4))
'                    lsRefFlag = Trim(GetText(vasExam, i, colResult + (gArrEquip(j, 1) - 1) * 4 + 1))
'                    lsPanicFlag = Trim(GetText(vasExam, i, colResult + (gArrEquip(j, 1) - 1) * 4 + 2))
'                    lsDeltaFlag = Trim(GetText(vasExam, i, colResult + (gArrEquip(j, 1) - 1) * 4 + 3))

                    lsExamCode = gArrEquip(j, 3)

                    '최종결과
'                        If lsRefFlag = "" Then
                        sResEnd = lsResult
'                        Else
'                            sResEnd = lsRefFlag & "/" & lsResult
'                        End If
                If Trim(GetText(vasExam, i, colBun)) = "외래" Then
                    '외래
                    SQL = " Update WCHDAT Set " & CR & _
                          " WD_RESULT = '" & sResEnd & "', " & CR & _
                          " WD_END_DEP ='3' " & CR & _
                          " Where WD_DATE = '" & Trim(GetText(vasExam, i, colReceDate)) & "' " & CR & _
                          " And WD_CHART = '" & Trim(GetText(vasExam, i, colReceNo)) & "' " & CR & _
                          " And WD_CODE = '" & lsExamCode & "' "
                          
                    res = SendQuery(gServer, SQL)
                    
                    If res = -1 Then
                        SaveQuery SQL
                        'db_RollBack gServer
                        Exit Sub
                    End If
        
                ElseIf Trim(GetText(vasExam, i, colBun)) = "입원" Then
                    '입원
                    SQL = " Update ICHDAT Set " & CR & _
                          " ID_RESULT = '" & sResult & "', " & CR & _
                          " ID_END_DEP ='3' " & CR & _
                          " Where ID_DATE = '" & Trim(GetText(vasExam, i, colReceDate)) & "' " & CR & _
                          " And ID_CHART = '" & Trim(GetText(vasExam, i, colReceNo)) & "' " & CR & _
                          " And ID_CODE = '" & lsExamCode & "' "
                          
                    res = SendQuery(gServer, SQL)
                    
                    If res = -1 Then
                        SaveQuery SQL
                        'db_RollBack gServer
                        Exit Sub
                    End If
                Else
                End If
                    'Local
                    SQL = "Update pat_res set " & vbCrLf & _
                          " resflag = 'B', examuid = '" & sExamUID & "' " & vbCrLf & _
                          "where examdate ='" & SeperatorCls(Text_Today) & "' " & vbCrLf & _
                          "  and barcode = '" & Trim(GetText(vasExam, i, colID)) & "' " & vbCrLf & _
                          "  and examcode = '" & lsExamCode & "' "
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then
                        Me.MousePointer = 0
                        SaveQuery SQL
                        Exit Sub
                    End If
                Next j
                
                'db_Commit gServer
                            
                SetText vasExam, "전송", i, colState
                'SetForeColor vasExam, i, i, 1, vasExam.MaxCols, 202, 202, 202
                
                vasExam.Value = 0
            End If
        End If
    Next i
    
'    If cn_Server_Flag Then DisConnect_Server
    
'    lblUID.Visible = False
'    txtUID.Visible = False
    Me.MousePointer = 0
    
    'db_Commit gServer
End Sub



Private Sub Command2_Click()
'    If Not Connect_Server Then
'        MsgBox "서버에 연결 되지 않았습니다 " & vbCrLf & "네트워크 확인 바랍니다", vbInformation, "알림"
'        cn_Server_Flag = False
'        Exit Sub
'    Else
'        cn_Server_Flag = True
'    End If
    
'     SQL = "Select ordseqno From slxworkt" & CR & _
'           " Where SPCID = '2030000888'" & CR & _
'           "And ExamCode = 'C4802'"
'     res = db_select_Col(gServer, SQL)
'
'    If Trim(gReadBuf(0)) <> "" Then
'       Text1 = Trim(gReadBuf(0))
'    End If


    SQL = " Update slxworkt " & CR & _
          "Set rslttext =  '/0.468', " & CR & _
     " execid = ''," & CR & _
     " deltayn = '' ," & CR & _
     " panicyn = '', " & CR & _
     " eqipcode = '52'," & CR & _
     " editid = '', " & CR & _
     " editip = '116.3.40.178', " & CR & _
     " editdate = to_date('2003-07-07 14:35:09','yyyy-mm-dd hh24:mi:ss')" & CR & _
 " Where SPCID = '2030000888'" & CR & _
 " And ordseqno  = '2001' "
 
 res = SendQuery(gServer, SQL)
 
 
 
End Sub

Private Sub Command3_Click()
    'CheckSum txtData
    
    Bioneer txtData
     
    txtData = ""
End Sub

Private Sub cmdSear_Click()
'서버에 전송안된것 불러오기
    
    Dim i, j, k As Long
    Dim X, y As Long
    Dim sPreID As String
    Dim sResult, sResult1 As String
    Dim iPos As Integer
    
    ClearSpread vasTemp1
    
    sResult = ""
    sResult1 = ""

    SQL = "Select barcode, seqno, diskno, posno, recedate, pid, " & _
          "       pname, psex, page, jumin1, jumin2, " & _
          "       pid, examgubun, resflag, examcode, " & _
          "       result, refflag, panicflag, deltaflag, examuid,equipcode " & vbCrLf & _
          "From pat_res " & CR & _
          "Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & CR & _
          "And ResFlag <> 'B' " & CR & _
          "And SampleType <> 'Q' Order by seqno "
          
    res = db_select_Vas(gLocal, SQL, vasTemp1)
    If res = -1 Then
        SaveQuery SQL
    End If
    If vasTemp1.DataRowCnt < 1 Then
        cmdClear_Click
        Exit Sub
    End If
    
    X = 1
    sPreID = Trim(GetText(vasTemp1, 1, 1))
    For j = 1 To 14
        SetText vasExam, Trim(GetText(vasTemp1, 1, j)), X, j + 1
    Next j
    For k = 1 To UBound(gArrEquip)
        If Trim(GetText(vasTemp1, 1, 15)) = gArrEquip(k, 3) Then
            y = 16 + (gArrEquip(k, 1)) * 4 - 3
            Exit For
        End If
    Next k
    
    sResult = Trim(GetText(vasTemp1, 1, 16))
    iPos = InStr(1, sResult, "/")
    sResult1 = Mid(sResult, iPos + 1)
    
    If y > 0 Then
        SetText vasExam, sResult1, X, y
        'SetText vasExam, Trim(GetText(vasTemp1, 1, 16)), x, y
        SetText vasExam, Trim(GetText(vasTemp1, 1, 17)), X, y + 1
        SetText vasExam, Trim(GetText(vasTemp1, 1, 18)), X, y + 2
        SetText vasExam, Trim(GetText(vasTemp1, 1, 19)), X, y + 3
    End If
    
    Select Case Trim(GetText(vasTemp1, 1, 17))
    Case "Pos"  '"H", "P"
        SetBackColor vasExam, X, X, y, y, 255, 149, 149
    Case "Neg"  '"L"
        SetBackColor vasExam, X, X, y, y, 149, 149, 255
    Case Else
        SetBackColor vasExam, X, X, y, y, 255, 255, 255
    End Select

    For i = 2 To vasTemp1.DataRowCnt
        If Trim(GetText(vasTemp1, i, 1)) = sPreID Then
            For k = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp1, i, 15)) = gArrEquip(k, 3) Then
                    y = 16 + gArrEquip(k, 1) * 4 - 3
                    Exit For
                End If
            Next k
            
            sResult = ""
            sResult1 = ""
            
            sResult = Trim(GetText(vasTemp1, i, 16))
            iPos = InStr(1, sResult, "/")
            sResult1 = Mid(sResult, iPos + 1)
    
            SetText vasExam, sResult1, X, y
            'SetText vasExam, Trim(GetText(vasTemp1, i, 16)), x, y
            SetText vasExam, Trim(GetText(vasTemp1, i, 17)), X, y + 1
            SetText vasExam, Trim(GetText(vasTemp1, i, 18)), X, y + 2
            SetText vasExam, Trim(GetText(vasTemp1, i, 19)), X, y + 3

            Select Case Trim(GetText(vasTemp1, i, 17))
            Case "Pos"  '"H", "P"
                SetBackColor vasExam, X, X, y, y, 255, 149, 149
            Case "Neg"  '"L"
                SetBackColor vasExam, X, X, y, y, 149, 149, 255
            Case Else
                SetBackColor vasExam, X, X, y, y, 255, 255, 255
            End Select

        Else
            X = X + 1
            
            If X > vasExam.MaxRows Then
                vasExam.MaxRows = X
            End If
            
            sPreID = Trim(GetText(vasTemp1, i, 1))
            For j = 1 To 14
                SetText vasExam, Trim(GetText(vasTemp1, i, j)), X, j + 1
            Next j
            For k = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp1, i, 15)) = gArrEquip(k, 3) Then
                    y = 16 + gArrEquip(k, 1) * 4 - 3
                    Exit For
                End If
            Next k
            
            
            sResult = ""
            sResult1 = ""
            
            sResult = Trim(GetText(vasTemp1, i, 16))
            iPos = InStr(1, sResult, "/")
            sResult1 = Mid(sResult, iPos + 1)
    
            SetText vasExam, sResult1, X, y
            'SetText vasExam, Trim(GetText(vasTemp1, i, 16)), x, y
            SetText vasExam, Trim(GetText(vasTemp1, i, 17)), X, y + 1
            SetText vasExam, Trim(GetText(vasTemp1, i, 18)), X, y + 2
            SetText vasExam, Trim(GetText(vasTemp1, i, 19)), X, y + 3

            Select Case Trim(GetText(vasTemp1, i, 17))
            Case "Pos"  '"H", "P"
                SetBackColor vasExam, X, X, y, y, 255, 149, 149
            Case "Neg"  '"L"
                SetBackColor vasExam, X, X, y, y, 149, 149, 255
            Case Else
                SetBackColor vasExam, X, X, y, y, 255, 255, 255
            End Select

        End If

    Next i
End Sub

Private Sub Command4_Click()
    Dim sSch1   As String
    Dim sSch2   As String
    Dim iRow    As Integer
    
    sSch1 = dtpSDate.Value
    sSch2 = dtpEDate.Value
    
    txtChkCnt = 0
    
    ClearSpread vasExam

    '지원부서처리여부 - 0:저장, 1:전송, 2:접수, 3:결과
    
    '외래
    SQL = " Select a.WD_DATE, a.WD_CHART, b.PE_SUJINJA, '', '', substring(b.PE_JUMIN,1,6), substring(b.PE_JUMIN,8,7), '','외래' " & CR & _
          " From WCHDAT a, PERSON b " & CR & _
          " Where a.WD_DATE >= '" & Trim(sSch1) & "' " & CR & _
          " And a.WD_DATE <= '" & Trim(sSch2) & "' " & CR & _
          " And a.WD_CODE IN ( " & gAllExam & " ) " & CR & _
          " And a.WD_END_DEP = '1' " & CR & _
          " And a.WD_CHART = b.PE_CHART " & CR & _
          " Group By a.WD_DATE, a.WD_CHART, b.PE_SUJINJA, substring(b.PE_JUMIN,1,6), substring(b.PE_JUMIN,8,7) "
    res = db_select_Vas(gServer, SQL, vasExam, , 6)
    
    '입원
    SQL = " Select a.ID_DATE, a.ID_CHART, b.PE_SUJINJA, '', '', substring(b.PE_JUMIN,1,6), substring(b.PE_JUMIN,8,7),  '', '입원' " & CR & _
          " From ICHDAT a, PERSON b " & CR & _
          " Where a.ID_DATE >= '" & Trim(sSch1) & "' " & CR & _
          " And a.ID_DATE <= '" & Trim(sSch2) & "' " & CR & _
          " And a.ID_CODE IN ( " & gAllExam & " ) " & CR & _
          " --And a.ID_END_DEP = '1' " & CR & _
          " And a.ID_CHART = b.PE_CHART " & CR & _
          " Group By a.ID_DATE,a.ID_CHART, b.PE_SUJINJA, substring(b.PE_JUMIN,1,6), substring(b.PE_JUMIN,8,7) "
    res = db_select_Vas(gServer, SQL, vasExam, vasExam.DataRowCnt + 1, 6)
    vasSort vasExam, 6, 7
    
    For iRow = 1 To vasExam.DataRowCnt
        CalAgeSex Trim(GetText(vasExam, iRow, colJumin1)) & Trim(GetText(vasExam, iRow, colJumin2)), Trim(GetDateFull)
        
        SetText vasExam, gPatGen.Sex, iRow, colSex
        SetText vasExam, gPatGen.Age, iRow, colAge
        
        vasExam.Row = iRow
        vasExam.Col = 1
        vasExam.Value = 1
        
        txtChkCnt = txtChkCnt + 1
    Next iRow
End Sub



Private Sub Command6_Click()
    Dim lRow As Long
    
    gOrder_Select.ok = 0
    
    giIndex = -1
    ReDim gOrder_List(0)
    
    kbnu_Order_Request txtTestID, gHPEQUIP
    
    If gOrder_Select.ok = 1 Then
        lRow = vasExam.DataRowCnt + 1
        If lRow > vasExam.MaxRows Then
            vasExam.MaxRows = lRow
        End If
        
        vasExam.SetText 2, lRow, txtTestID
        vasExam.SetText 7, lRow, gOrder_Select.PT_NO
        vasExam.SetText 8, lRow, gOrder_Select.PT_NM
        If InStr(1, gOrder_Select.Sex, "/") > 0 Then
            vasExam.SetText 9, lRow, Left(gOrder_Select.Sex, InStr(1, gOrder_Select.Sex, "/") - 1)
            vasExam.SetText 10, lRow, Mid(gOrder_Select.Sex, InStr(1, gOrder_Select.Sex, "/") + 1)
        End If
    End If
End Sub

Private Sub Command5_Click()
    Winsock1.SendData chrENQ
End Sub

Private Sub Form_Activate()
    Text_Today.SetFocus
End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim sRet As String
    
     
    Me.Left = 0
    Me.Top = 0
    Me.Height = 11190
    Me.Width = 15360

    cmdClear_Click
    
    'ini파일에서 정보가져오기
    GetSetup

'    Winsock1.LocalPort = gRemote.RemotePort
    Call Winsock1.Connect(gRemote.RemoteHost, gRemote.RemotePort)
    
   Timer1.Enabled = True
    
    If Not Connect_Local Then
        MsgBox "로컬에 연결되지 않습니다. 종료합니다."
        cn_Local_Flag = False
        End
    Else
        cn_Local_Flag = True
    End If
    
'    sRet = kbnu_Server_Connect(gHPEQUIP)
    
    If Not Connect_Server Then
        MsgBox "서버에 연결되지 않습니다. 종료합니다."
        cn_Server_Flag = False
        End
    Else
        cn_Server_Flag = True
    End If
    
    '검사일자
'    Text_Today = Format(CDate(GetDateFull), "yyyy/mm/dd")
    dtpSDate.Value = Text_Today.Text
    dtpEDate.Value = Text_Today.Text
    
    txtUID.Text = Trim(gUID)
    'txtUID.Text = ""
    
    '순서 컬럼 추가함
    SQL = " Select seqno From equipexam "
    res = db_select_Col(gLocal, SQL)
    
    If res = -1 Then
        SQL = " alter table equipexam " & vbCrLf & _
              " Add seqno varchar(10) "
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
        End If
    End If
    
    '검사코드 불러오기
    GetExamCode
    
'2005/07/26 이상은 -
'    Command_Search_Click

    sDate = ""
    SQL = "select min(examdate) From pat_res "
    res = db_select_Var(gLocal, SQL, sDate)
    If IsDate(sDate) Then
        dtpDel1.Value = Format(sDate, "yyyy-mm-dd")
    End If
    
    sDate = ""
    SQL = "select max(examdate) From pat_res "
    res = db_select_Var(gLocal, SQL, sDate)
    If IsDate(sDate) Then
        dtpDel2.Value = Format(sDate, "yyyy-mm-dd")
    End If
    
    '로컬데이타 지우기
    sDate = Format(DateAdd("y", CDate(Text_Today.Text), -60), "yyyymmdd")
    
    SQL = "Delete from pat_res where examdate < '" & sDate & "' "
    SendQuery gLocal, SQL
        
    SQL = " Select examgubun From pat_res "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table pat_res Add Column examgubun text(10) "
        res = SendQuery(gLocal, SQL)
    End If
End Sub

Function GetExamCode() As Integer
    Dim i As Long
    Dim j As Long
    
    gAllExam = ""
    
    ClearSpread vasTemp
    GetExamCode = -1
    
'    SQL = "Select equipcode, examcode, examname, resprec, reflow, refhigh, paniclow, panichigh, deltavalue " & vbCrLf & _
'          "From equipexam " & vbCrLf & _
'          "Where equipno = '" & gEquip & "' " & vbCrLf & _
'          "order by  examcode "
    
    cboExam.Clear
    SQL = "Select unitcode, max(equipcode) " & vbCrLf & _
          "From equipexam " & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "   " & vbCrLf & _
          "group by  unitcode "
    res = db_select_Combo(gLocal, SQL, cboExam)
    If cboExam.ListCount > 0 Then cboExam.ListIndex = 0
    
    SQL = "Select equipcode, examcode, examname, resprec, reflow, refhigh, " & vbCrLf & _
          "paniclow, panichigh, deltavalue, unitcode, OrdGubun " & vbCrLf & _
          "From equipexam " & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "order by  seqno "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    If res > 0 Then
        ReDim gArrEquip(1 To vasTemp.DataRowCnt, 1 To 12)
    Else
        SaveQuery SQL
        Exit Function
    End If
    vasList.MaxCols = UBound(gArrEquip) + 1
    vasExam.MaxCols = vasTemp.DataRowCnt * 5 + colResult - 1 + 1
    
    colResult1 = vasTemp.DataRowCnt * 4 + colResult
    colHospital = vasTemp.DataRowCnt * 5 + colResult - 1 + 1
    vasExam.ColWidth(colHospital) = 0
    
    For i = 1 To vasTemp.DataRowCnt
        gArrEquip(i, 1) = i
        For j = 1 To 11
            gArrEquip(i, j + 1) = Trim(GetText(vasTemp, i, j))
        Next j
        
        SetText vasExam, gArrEquip(i, 4), 0, colResult + (i - 1) * 4
'        vasExam.ColWidth(colResult + (i - 1) * 4) = 6.13
        vasExam.ColWidth(colResult + (i - 1) * 4 + 1) = 0
        vasExam.ColWidth(colResult + (i - 1) * 4 + 2) = 0
        vasExam.ColWidth(colResult + (i - 1) * 4 + 3) = 0
        
        vasExam.ColWidth(colResult1 + i - 1) = 0
        
        SetText vasList, gArrEquip(i, 4), 0, i + 1

        SetText vasList, gArrEquip(i, 4), 0, i + 1
        
        
        vasExam.ColWidth(colResult + (i - 1) * 4) = 13.3
        
        '2005/07/26 이상은 - 추가
        If Trim(GetText(vasTemp, i, 2)) <> "" Then
            If gAllExam = "" Then
                gAllExam = "'" & Trim(GetText(vasTemp, i, 2)) & "'"
            Else
                gAllExam = gAllExam & ", '" & Trim(GetText(vasTemp, i, 2)) & "'"
            End If
        End If
            
    Next i
    
    GetExamCode = 1
End Function

Private Sub Form_Unload(Cancel As Integer)
    DisConnect_Local
    DisConnect_Server
    
    If Form_Main.optOption(0).Value = True Then
        WritePrivateProfileString "CONFIG", "MODE", "1", App.Path & "\Interface.ini"
    Else
        WritePrivateProfileString "CONFIG", "MODE", "0", App.Path & "\Interface.ini"
    End If
    
    End
End Sub

Private Sub ACL9000(asData As String)
    Dim lsData As String
    Dim lsOrder() As String
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim z As Integer
    
    Dim iCnt As Integer
    
    Dim ResultTbl(1 To 40) As String
    Dim TablePtr As Integer
    Dim sTmp As String
    
    Dim sDec As String
    Dim sCnt As String
    
    Dim lCol As Long
    
    Dim sDate As String
    Dim sTmpStr As String
    Dim sPoint As String
    Dim sRefFlag As String
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sPanicFlag As String
    Dim sDeltaFlag As String
    Dim sReceNo As String
    Dim sReceDate  As String
    Dim sPID As String
    Dim sPname As String
    Dim sJumin1 As String
    Dim sJumin2 As String
    Dim sPSex As String
    Dim sPage As String
    Dim sBun As String
    Dim sTestID As String
    Dim sFlag As String
    Dim sExamCode As String
    Dim sExamCode1 As String
    Dim sResult As String
    Dim sExamDate As String
    
    Dim sLevelNo As String
    
    Dim sResEnd As String
    
    Dim sExecDate As String
    Dim sExecDate1 As String
    
    sExecDate = Format(Trim(Text_Today.Text), "yyyymmdd")
    sExecDate1 = Format(DateAdd("d", "1", Trim(Text_Today.Text)), "yyyymmdd")
    
    TablePtr = 1
' ----- for start
    For j = 1 To Len(asData)
        If (Mid(asData, j, 1) = "|") Then
            TablePtr = TablePtr + 1
            ResultTbl(TablePtr) = " "
        Else
            ResultTbl(TablePtr) = ResultTbl(TablePtr) + Mid(asData, j, 1)
        End If
    Next j
' ------- for end
    
    If Mid(ResultTbl(1), 2, 1) = "H" Then     'Header Record
        Var_Clear
        
        iCnt = 0
                
        i = InStr(1, lsData, "|")
        Do While i > 0
            iCnt = iCnt + 1
            lsData = Mid(lsData, i + 1)
            i = InStr(1, lsData, "|")
        Loop
        
        lsHeaderDate = lsData
        
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "Q" Then     'Request Information Record
        gOrderMessage = "Q"
        
        sTmp = ResultTbl(3)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sBarcode = Left(sTmp, i - 1)
        
        gOrdRow = 0
        
        lRow = -1
        For i = 1 To vasExam.DataRowCnt
            If sBarcode <> "" Then  '메뉴얼일경우 바코드 없음
                If Trim(GetText(vasExam, i, colID)) = sBarcode Then
                    lRow = i
                    Exit For
                End If
            End If
        Next i

        If lRow = -1 Then
            lRow = vasExam.DataRowCnt + 1
            If vasExam.MaxRows < lRow Then
                vasExam.MaxRows = lRow
            End If
        End If
        
        SetText vasExam, sBarcode, lRow, colID
        
        res = MakeOrder(lRow, sBarcode)     'Order 만들기
                
        If res = 1 Then
            'Order 만들기
            vasOrder.MaxRows = 20
            ClearSpread vasOrder
            
            ReDim lsOrder(0)
            z = 0
            
            For i = 1 To vasOrdBuff.DataRowCnt
                gReadBuf(0) = ""
                gReadBuf(1) = ""
                
                SQL = "Select ExamCode, EquipCode from EquipExam " & vbCrLf & _
                      "where EquipNo = '" & gEquip & "' " & vbCrLf & _
                      "  and ExamCode = '" & Trim(GetText(vasOrdBuff, i, 4)) & "' " & vbCrLf & _
                      "  and UseFlag = '1' " & vbCrLf & _
                      "  and OrdGubun = 'O' "
                res = db_select_Col(gLocal, SQL)
                'Debug.Print Trim(GetText(vasOrdBuff, i, 4))
                
                If Trim(gReadBuf(0)) = Trim(GetText(vasOrdBuff, i, 4)) Then
                    k = -1
                    For j = LBound(lsOrder) To UBound(lsOrder)
                        If Trim(lsOrder(j)) = Trim(gReadBuf(1)) Then
                            k = 1
                            Exit For
                        End If
                    Next j
                    If k = -1 Then
                        z = z + 1
                        ReDim Preserve lsOrder(z)
                        lsOrder(z) = Trim(gReadBuf(1))
                    End If
                End If
            Next i
            
            i = 0
            'Head
            i = i + 1
            If i = 8 Then
                i = 0
            End If
            gHeader = "H|\^&||||||||ACL9000||P|1|" & lsHeaderDate & chrCR & chrETX
            gHeader = chrSTX & CCur(i) & gHeader & ASTM_CSum(CStr(1) & gHeader) & chrCR & chrLF
            
            SetText vasOrder, gHeader, vasOrder.DataRowCnt + 1, 1
            
            'Patient
            i = i + 1
            If i = 8 Then
                i = 0
            End If
            gPatient = "P|1|||||||U||||||||||||||||||||||||||" & chrCR & chrETX
            gPatient = chrSTX & CCur(i) & gPatient & ASTM_CSum(CStr(2) & gPatient) & chrCR & chrLF
            
            SetText vasOrder, gPatient, vasOrder.DataRowCnt + 1, 1
            
            'Order
            For j = 1 To UBound(lsOrder)
                i = i + 1
                If i = 8 Then
                    i = 0
                End If
                gOrder = "O|" & j & "|" & sBarcode & "||^^^" & lsOrder(j) & "|||" & _
                                          "||||||||^||||||||||O||||||" & chrCR & chrETX
                gOrder = chrSTX & CCur(i) & gOrder & ASTM_CSum(CStr(i) & gOrder) & chrCR & chrLF
            
                SetText vasOrder, gOrder, vasOrder.DataRowCnt + 1, 1
            Next j

            If i = 2 Then
                i = i + 1
                If i = 8 Then
                    i = 0
                End If
                gOrder = "O|1|" & sBarcode & "||^^^ALL|||" & _
                                          "||||||||||||||||||O||||" & chrCR & chrETX
                gOrder = chrSTX & CCur(i) & gOrder & ASTM_CSum(CStr(i) & gOrder) & chrCR & chrLF
    
                SetText vasOrder, gOrder, vasOrder.DataRowCnt + 1, 1
    
            End If
            
            'Msg End
            i = i + 1
            If i = 8 Then
                i = 0
            End If
            gMsgEnd = "L|1|N" & chrCR & chrETX
            gMsgEnd = Chr(2) & CCur(i) & gMsgEnd & ASTM_CSum(CStr(i) & gMsgEnd) & chrCR & chrLF
        
            SetText vasOrder, gMsgEnd, vasOrder.DataRowCnt + 1, 1
            
            i = i + 1
            If i = 8 Then
                i = 0
            End If
            SetText vasOrder, chrEOT, vasOrder.DataRowCnt + 1, 1
            
        End If
    End If
    

    If (Mid(ResultTbl(1), 2, 1) = "O") Then          'Test Order Record
        sSampleType = "P"
        
        sBarcode = Trim(ResultTbl(3))

        lRow = -1
        For i = 1 To vasExam.DataRowCnt
            'If sSampleType = "P" Then
                If Trim(GetText(vasExam, i, colID)) = sBarcode Then
                    lRow = i
                    Exit For
                End If
            'End If
        Next i
        
        If lRow < 0 Then
            lRow = vasExam.DataRowCnt + 1
            If vasExam.MaxRows < lRow Then
                vasExam.MaxRows = lRow
            End If
        End If
        
        SetText vasExam, sBarcode, lRow, colID

'        If sSampleType = "Q" Then
'            SetText vasExam, "QC", lRow, colState
'            SetText vasExam, sResDateTime, lRow, colReceDate
'        Else
            If Trim(GetText(vasExam, lRow, colReceNo)) = "" Then
                res = GetPatientInfo(lRow)
                If res = -1 Then
                    SetText vasExam, "", lRow, colState
                ElseIf res = 0 Then
                    SetText vasExam, "미접수", lRow, colState
                End If
            End If
'        End If
    End If
    
    If (Mid(ResultTbl(1), 2, 1) = "R") Then     'Result
        gOrderMessage = "R"
        
        sTmp = ResultTbl(3)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        sTestID = Trim(sTmp)
        
        sResult = Trim(ResultTbl(4))
        
        sFlag = Trim(ResultTbl(5))      '단위정보


        '검사코드 불러오기
        ClearSpread vasTemp
        SQL = "SELECT EquipCode, ExamCode, ExamName, UnitCode " & CR & _
              "  From EquipExam " & CR & _
              " WHERE Equipno = '" & gEquip & "' " & CR & _
              "   and EquipCode = '" & sTestID & "' " & vbCrLf & _
              "   and UnitCode = '" & sFlag & "' " & vbCrLf & _
              " Order by seqno "
        res = db_select_Vas(gLocal, SQL, vasTemp)
        
        sExamCode1 = ""
        For i = 1 To vasTemp.DataRowCnt
            If Trim(GetText(vasTemp, i, 2)) <> "" Then
                If sExamCode1 = "" Then
                    sExamCode1 = "'" & Trim(GetText(vasTemp, i, 2)) & "'"
                Else
                    sExamCode1 = sExamCode1 & ",'" & Trim(GetText(vasTemp, i, 2)) & "' "
                End If
            End If
        Next i
          
        gReadBuf(0) = ""
        sExamCode = ""
        
        '2008/02/26 이상은 - 응급검사가 걸릴때도 있고 안 걸릴때도 있음
'        SQL = " Select W.EXAMCODE " & CR & _
'              " From SLXWORKT W, SLSPCMDT S " & CR & _
'              " Where W.SPCID = '" & sBarCode & "' " & CR & _
'              " And W.EXECDATE >= to_date('" & Trim(sExecDate) & "', 'yyyymmdd') " & CR & _
'              " And W.EXECDATE <  to_date('" & Trim(sExecDate1) & "', 'yyyymmdd') " & CR & _
'              " And W.EXAMCODE IN (" & sExamCode1 & ") " & CR & _
'              " And W.WORKNO between to_number('1') And to_number('9999') " & CR & _
'              " And S.SPCID = W.SPCID "

        SQL = " Select W.EXAMCODE " & CR & _
              " From SLXWORKT W, SLSPCMDT S " & CR & _
              " Where W.SPCID = '" & sBarcode & "' " & CR & _
              " And W.EXAMCODE IN (" & sExamCode1 & ") " & CR & _
              " And W.WORKNO between to_number('1') And to_number('9999') " & CR & _
              " And S.SPCID = W.SPCID "
        res = db_select_Col(gServer, SQL)
        
        If res = 1 And gReadBuf(0) <> "" Then
            sExamCode = Trim(gReadBuf(0))
        End If
                
        For i = 1 To UBound(gArrEquip)
            If Trim(sExamCode) = gArrEquip(i, 3) Then
                k = gArrEquip(i, 1)
                lCol = (gArrEquip(i, 1) - 1)
                
                Exit For
            End If
        Next i

        If k = 0 Then
            Exit Sub
        End If
        
        sRefFlag = ""
        sPanicFlag = ""
        sDeltaFlag = ""
       If IsNumeric(sResult) Then
            '소수점 처리
            sPoint = gArrEquip(k, 5)
            
            If IsNumeric(sPoint) Then
                If CInt(sPoint) > 0 Then
                    sTmpStr = "#0."
                    For i = 1 To CInt(sPoint)
                        sTmpStr = sTmpStr & "0"
                    Next i
                Else
                    sTmpStr = "#0"
                End If
                sResult = Format(sResult, sTmpStr)

                SetText vasExam, sResult, lRow, colResult + lCol * 4
                SetText vasExam, sResult, lRow, colResult1 + lCol
            End If
        
            '참고치 체크================================================================
            sRefLow = gArrEquip(k, 6)
            sRefHigh = gArrEquip(k, 7)
            If Not IsNumeric(sRefLow) Then
                sRefLow = "0"
            End If
            If Not IsNumeric(sRefHigh) Then
                sRefHigh = "0"
            End If
            If CCur(sRefLow) = 0 And CCur(sRefHigh) = 0 Then
                sRefFlag = ""
            Else
                If CCur(sResult) < CCur(sRefLow) Then
                    sRefFlag = "Neg"    'Low
                End If
                If CCur(sResult) > CCur(sRefHigh) Then
                    sRefFlag = "Pos"    'High
                End If
            End If
            SetText vasExam, sRefFlag, lRow, colResult + lCol * 4 + 1

            'Panic 체크================================================================
            If Not IsNumeric(gArrEquip(k, 8)) Then
                gArrEquip(k, 8) = "0"
            End If
            If Not IsNumeric(gArrEquip(k, 9)) Then
                gArrEquip(k, 9) = "0"
            End If
            If CCur(gArrEquip(k, 8)) = 0 And CCur(gArrEquip(k, 9)) = 0 Then
                sPanicFlag = "X"
            Else
                If CCur(sResult) < CCur(gArrEquip(k, 8)) Then
                    sPanicFlag = "Neg"
                End If
                If CCur(sResult) > CCur(gArrEquip(k, 9)) Then
                    sPanicFlag = "Pos"
                End If
            End If
            
            If sPanicFlag = "" Then
                sPanicFlag = "X"
            End If
            
            SetText vasExam, sPanicFlag, lRow, colResult + lCol * 4 + 2
            'Delta 체크================================================================
            sDeltaFlag = Delta_Check(lRow, sResult, k)
            If sDeltaFlag = "" Then
                sDeltaFlag = "X"
            End If
            
            SetText vasExam, sDeltaFlag, lRow, colResult + lCol * 4 + 3
        End If
        
        '결과 셀 색깔 변화=========================================================
        Select Case sRefFlag
        Case "Pos"   'Positive, 'High
            SetBackColor vasExam, lRow, lRow, colResult + lCol * 4, colResult + lCol * 4 + 1, 246, 150, 121
        Case "Neg"  'Negative, Low
            SetBackColor vasExam, lRow, lRow, colResult + lCol * 4, colResult + lCol * 4 + 1, 255, 245, 104
        Case Else   'Normal
        End Select
        Select Case sPanicFlag
        Case "Pos"   'High
            SetBackColor vasExam, lRow, lRow, colResult + lCol * 4, colResult + lCol * 4 + 2, 242, 108, 79
        Case "Neg"
            SetBackColor vasExam, lRow, lRow, colResult + lCol * 4, colResult + lCol * 4 + 2, 60, 184, 120
        Case Else   'Normal
        End Select
        If sDeltaFlag = "D" Then
            SetBackColor vasExam, lRow, lRow, colResult + lCol * 4, colResult + lCol * 4 + 3, 255, 0, 0
        End If
        '변수에 담기 ====================================================================================
        sDate = GetDateFull
        sExamDate = SeperatorCls(Text_Today.Text)
        sBarcode = Trim(GetText(vasExam, lRow, colID))
        sReceNo = Trim(GetText(vasExam, lRow, colReceNo))
        sSeqNo = Trim(GetText(vasExam, lRow, colSeqNo))
        sDiskNo = Trim(GetText(vasExam, lRow, colDiskNo))
        sPosNo = Trim(GetText(vasExam, lRow, colPosNo))
        sPname = Trim(GetText(vasExam, lRow, colPName))
        sPSex = Trim(GetText(vasExam, lRow, colSex))
        sPage = Trim(GetText(vasExam, lRow, colAge))
        sJumin1 = Trim(GetText(vasExam, lRow, colJumin1))
        sJumin2 = Trim(GetText(vasExam, lRow, colJumin2))
        sReceDate = Trim(GetText(vasExam, lRow, colReceDate))
        sPID = Trim(GetText(vasExam, lRow, colPID))
        sBun = Trim(GetText(vasExam, lRow, colBun))
        sSampleType = Trim(GetText(vasExam, lRow, colSampleType))

        'Local Table Insert
        '환자 데이타 ====================================================================================
        db_BeginTran gLocal
        
        If sSampleType = "Q" Then   'QC Data Local에 저장
            
'            SQL = "Select t_mean, t_sd from qcexam " & vbCrLf & _
'                  "where equipno = '" & gEquip & "' " & vbCrLf & _
'                  "  and validstart >= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'                  "  and valiend <= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'                  "  and levelname = '" & sBarCode & "' " & vbCrLf & _
'                  "  and equipcode = '" & sTestID & "' "
'            res = db_select_Col(gLocal, SQL)
'            If res > 0 Then
'                If IsNumeric(gReadBuf(0)) And IsNumeric(gReadBuf(1)) Then
'                    sRefLow = CCur(gReadBuf(0)) - CCur(gReadBuf(1))
'                    sRefHigh = CCur(gReadBuf(0)) + CCur(gReadBuf(1))
'                    If CCur(sRefHigh) < CCur(sResult) Then
'                        sRefFlag = "H"
'                    End If
'                    If CCur(sRefLow) > CCur(sResult) Then
'                        sRefFlag = "L"
'                    End If
'                End If
'            End If
'
'            sCnt = ""
'            SQL = "Select count(*) from qc_res " & vbCrLf & _
'                  "where equipno = '" & gEquip & "' " & vbCrLf & _
'                  "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'                  "  and examtime = '" & Mid(sResDateTime, 10, 6) & "' " & vbCrLf & _
'                  "  and levelname = '" & sBarCode & "' " & vbCrLf & _
'                  "  and equipcode = '" & sTestID & "' "
'            res = db_select_Var(gLocal, SQL, sCnt)
'            If res <= 0 Then
'                SaveQuery SQL
'                db_RollBack gLocal
'                Exit Sub
'            End If
'            res = db_select_Var(gLocal, SQL, sCnt)
'            If res <= 0 Then
'                SaveQuery SQL
'                db_RollBack gLocal
'                Exit Sub
'            End If
'            If Not IsNumeric(sPage) Then
'                sPage = "0"
'            End If
'
'            If CInt(sCnt) > 0 Then
'                SQL = "delete from qc_res " & vbCrLf & _
'                      "where equipno = '" & gEquip & "' " & vbCrLf & _
'                      "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'                      "  and examtime = '" & Mid(sResDateTime, 9, 4) & "' " & vbCrLf & _
'                      "  and levelname = '" & sBarCode & "' " & vbCrLf & _
'                      "  and equipcode = '" & sTestID & "' "
'                res = SendQuery(gLocal, SQL)
'                If res = -1 Then
'                    db_RollBack gLocal
'                    SaveQuery SQL
'                    Exit Sub
'                End If
'            End If
'            SQL = "Insert into qc_res (equipno, examdate, examtime, levelname, equipcode, result, resflag, remark, examuid) " & vbCrLf & _
'                  "values ('" & gEquip & "', '" & Left(sResDateTime, 8) & "', '" & Mid(sResDateTime, 9, 4) & "', '" & sBarCode & "', '" & sTestID & "', '" & sResult & "', '" & sRefFlag & "','','') "
'            res = SendQuery(gLocal, SQL)
'            If res = -1 Then
'                db_RollBack gLocal
'                SaveQuery SQL
'                Exit Sub
'            End If
                
        Else    'Sample Data Local에 저장
            sCnt = ""
            SQL = "Select count(*) from pat_res " & vbCrLf & _
                  "where barcode = '" & sBarcode & "' and equipcode = '" & sTestID & "'  and examcode = '" & sExamCode & "' "
            res = db_select_Var(gLocal, SQL, sCnt)
            If res <= 0 Then
                SaveQuery SQL
                db_RollBack gLocal
                Exit Sub
            End If
            If Not IsNumeric(sPage) Then
                sPage = "0"
            End If
            
            If sRefFlag = "" Then
                sResEnd = sResult
            Else
                sResEnd = sRefFlag & "/" & sResult
            End If
            
            If CInt(sCnt) = 0 Then
                '입력
                SQL = "insert into pat_res (examdate, equipno, seqno, diskno, posno, " & vbCrLf & _
                      " barcode, equipcode, examcode, result, refflag,  " & vbCrLf & _
                      " refvalue,receno, recedate, pid, pname,  " & vbCrLf & _
                      " jumin1, jumin2,psex, page, resflag,  " & vbCrLf & _
                      " resdate, panicflag, deltaflag, sampletype ) " & vbCrLf & _
                      "values ('" & sExamDate & "', '" & gEquip & "', '" & sSeqNo & "', '" & sDiskNo & "', '" & sPosNo & "', " & vbCrLf & _
                      " '" & sBarcode & "', '" & sTestID & "', '" & sExamCode & "', '" & sResEnd & "', '" & sRefFlag & "', " & vbCrLf & _
                      " '" & sRefLow & " - " & sRefHigh & "', '', '" & sReceDate & "', '" & sReceNo & "', '" & sPname & "', " & vbCrLf & _
                      " '" & sJumin1 & "', '" & sJumin2 & "', '" & sPSex & "', " & sPage & ", 'A', " & vbCrLf & _
                      " '" & sDate & "', '" & sPanicFlag & "', '" & sDeltaFlag & "', '" & sSampleType & "' ) "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    db_RollBack gLocal
                    Exit Sub
                End If
            ElseIf CInt(sCnt) > 0 Then
                '수정
                SQL = "Update pat_res set " & vbCrLf & _
                      " seqno =  '" & sSeqNo & "', " & vbCrLf & _
                      " diskno =  '" & sDiskNo & "', " & vbCrLf & _
                      " posno =  '" & sPosNo & "', " & vbCrLf & _
                      " examcode =  '" & sExamCode & "', " & vbCrLf & _
                      " result =  '" & sResEnd & "' , " & vbCrLf & _
                      " refflag =  '" & sRefFlag & "', " & vbCrLf & _
                      " refvalue =  '" & sRefLow & " - " & sRefHigh & "', " & vbCrLf & _
                      " receno =  '', " & vbCrLf & _
                      " recedate =  '" & sReceDate & "', " & vbCrLf & _
                      " pid =  '" & sReceNo & "', " & vbCrLf & _
                      " pname =  '" & sPname & "', " & vbCrLf & _
                      " jumin1 =  '" & sJumin1 & "', " & vbCrLf & _
                      " jumin2 =  '" & sJumin2 & "', " & vbCrLf & _
                      " psex =  '" & sPSex & "', " & vbCrLf & _
                      " page =  " & sPage & ", " & vbCrLf & _
                      " resflag =  'A', " & vbCrLf & _
                      " resdate =  '" & sDate & "', " & vbCrLf & _
                      " panicflag = '" & sPanicFlag & "', " & vbCrLf & _
                      " deltaflag = '" & sDeltaFlag & "', " & vbCrLf & _
                      " sampletype = '" & sSampleType & "' " & vbCrLf & _
                      "where examdate ='" & sExamDate & "' " & vbCrLf & _
                      "  and barcode = '" & sBarcode & "' " & vbCrLf & _
                      "  and equipcode = '" & sTestID & "' " & vbCrLf & _
                      "  and examcode = '" & sExamCode & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    db_RollBack gLocal
                    Exit Sub
                End If
            End If
        End If
        
        db_Commit gLocal
        
        '서버에 바로 전송하기
        SetText vasList, "결과", lRow, colState
        
        If optOption(0).Value = True Then
            To_Server lRow
        End If
        '==============================================================================================

    End If
End Sub


Function MakeOrder(argRow As Long, argID As String) As Integer
    Dim i, j As Integer
    Dim sCnt As String
    Dim iGet As Integer
    
    Dim sExecDate As String
    Dim sExecDate1 As String
    
    MakeOrder = -1
    sOrder = ""

    
    sExecDate = Format(Trim(Text_Today), "yyyymmdd")
    sExecDate1 = Format(DateAdd("d", "1", Trim(Text_Today)), "yyyymmdd")
    
    iGet = 1
    
    sCnt = ""
    SQL = "Select count(*) from pat_res " & vbCrLf & _
          "where barcode = '" & argID & "'  "
    res = db_select_Var(gLocal, SQL, sCnt)
    If res > 0 Then
        If Not IsNumeric(sCnt) Then
            sCnt = "0"
        End If
        If CInt(sCnt) > 0 Then
            iGet = 2
        End If
    ElseIf res = -1 Then
        SaveQuery SQL
    End If
    
    ClearSpread vasOrdBuff
    
    If iGet = 2 Then    '재검
        SQL = "select equipcode from pat_res where barcode = '" & argID & "' and result = '*' "
        res = db_select_Vas(gLocal, SQL, vasOrdBuff)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        If vasOrdBuff.DataRowCnt > 0 Then
'            For i = 1 To vasOrdBuff.DataRowCnt
'                If Trim(GetText(vasOrdBuff, i, 1)) <> "" Then
'                    If sOrder = "" Then
'                        sOrder = "^^^" & Trim(GetText(vasOrdBuff, i, 1)) & "^0"
'                    Else
'                        sOrder = sOrder & "\^^^" & Trim(GetText(vasOrdBuff, i, 1)) & "^0"
'                    End If
'
'                End If
'            Next i
            MakeOrder = 1
            SetText vasExam, "재검", argRow, colState
            Exit Function
        Else
            iGet = 1
        End If
    End If
    If iGet = 1 Then    'Server에서 검사항목 불러오기
'        If Not cn_Server_Flag Then
'            If Not Connect_Server Then
'                cn_Server_Flag = False
'                Exit Function
'            Else
'                cn_Server_Flag = True
'            End If
'        End If
    
        res = GetPatientInfo(argRow)    '환자정보 조회
        If res < 1 Then
            MakeOrder = 0
            SetText vasExam, "미접수", argRow, colState
            Exit Function
        End If
        
'        If Trim(GetText(vasExam, argRow, colReceNo)) = "" Then
'            Exit Function
'        End If
        
'        If Not cn_Server_Flag Then
'            If Not Connect_Server Then
'                cn_Server_Flag = False
'                Exit Function
'            Else
'                cn_Server_Flag = True
'            End If
'        End If

'2008/02/26 이상은 - 응급검사가 걸릴때도 있고 안 걸릴때도 있음
'        SQL = " Select W.SPCID, to_char(W.ORDDATE, 'yyyy-mm-dd'), W.ORDSEQNO, W.EXAMCODE, to_char(W.WORKNO), W.RSLTTEXT, s.SPCCODE" & CR & _
'              " From SLXWORKT W, SLSPCMDT S " & CR & _
'              " Where W.SPCID = '" & argID & "' " & CR & _
'              " And W.EXECDATE >= to_date('" & Trim(sExecDate) & "', 'yyyymmdd') " & CR & _
'              " And W.EXECDATE <  to_date('" & Trim(sExecDate1) & "', 'yyyymmdd') " & CR & _
'              " And W.PROCSTAT = 'E' " & CR & _
'              " And W.WORKNO between to_number('1') And to_number('9999') " & CR & _
'              " And S.SPCID    = W.SPCID " & CR & _
'              " Order by W.WORKNO, W.EXAMCODE "

        SQL = " Select W.SPCID, to_char(W.ORDDATE, 'yyyy-mm-dd'), W.ORDSEQNO, W.EXAMCODE, to_char(W.WORKNO), W.RSLTTEXT, s.SPCCODE" & CR & _
              " From SLXWORKT W, SLSPCMDT S " & CR & _
              " Where W.SPCID = '" & argID & "' " & CR & _
              " And W.PROCSTAT = 'E' " & CR & _
              " And W.WORKNO between to_number('1') And to_number('9999') " & CR & _
              " And S.SPCID    = W.SPCID " & CR & _
              " Order by W.WORKNO, W.EXAMCODE "
        res = db_select_Vas(gServer, SQL, vasOrdBuff)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        If vasOrdBuff.DataRowCnt > 0 Then
            For i = 1 To vasOrdBuff.DataRowCnt
                For j = 1 To UBound(gArrEquip)
                    'If Trim(gArrEquip(j, 3)) = Trim(GetText(vasOrdBuff, i, 4)) And Trim(gArrEquip(j, 12)) = "O" Then
                    If Trim(gArrEquip(j, 3)) = Trim(GetText(vasOrdBuff, i, 4)) Then
                        SetText vasExam, "*", argRow, colResult + (gArrEquip(j, 1) - 1) * 4
                    
                        SetText vasOrdBuff, gArrEquip(j, 2), i, 2
'                        If sOrder = "" Then
'                            sOrder = "^^^" & gArrEquip(j, 2) & "^0"
'                        Else
'                            sOrder = sOrder & "\^^^" & gArrEquip(j, 2) & "^0"
'                        End If
'
'                        Exit For
                    End If
                Next j
            Next i
            MakeOrder = 1
        Else
            MakeOrder = 0
            SetText vasExam, "없음", argRow, colState
            Exit Function
        End If
        
'        DisConnect_Server
    End If
      
End Function

Function To_Server_1(ByVal asRow As Long) As Integer
    Dim iStart As Long
    Dim iEnd As Long
    Dim i, j As Long
    
    Dim lsExamCode As String
    Dim lsResult As String
    Dim lsRefFlag As String
    Dim lsReceNo As String
    Dim lsReceDate As String
    Dim lsBun As String
    
    Dim sCnt As String
    
    Dim sExamUID As String
    
    Dim sResEnd As String
    Dim lsPID As String
    Dim sSeqNo As String
    
    Dim sEditDate As String
    Dim lsPanicFlag As String
    Dim lsDeltaFlag As String
    
    Dim lsSenddata As String
    
    lsSenddata = ""
    
    To_Server_1 = -1
    
    sEditDate = Format(GetDateFull, "yyyy-mm-dd hh:mm:ss")
    
    If txtUID.Text = "" Then
        sExamUID = InputBox("검사자를 입력하세요", "알림")
        If sExamUID = "" Then
            Exit Function
        End If
        txtUID = sExamUID
    End If
    
'    lblUID.Visible = True
'    txtUID.Visible = True
    
  
'    If Not Connect_Server Then
'        MsgBox "서버에 연결 되지 않았습니다 " & vbCrLf & "네트워크 확인 바랍니다", vbInformation, "알림"
'        cn_Server_Flag = False
'        Exit Sub
'    Else
'        cn_Server_Flag = True
'    End If
    
    Me.MousePointer = 11
    
    
    If Len(Trim(GetText(vasExam, asRow, colID))) > 5 Then   '수작업인것은 바코드(10자리)보다 작으므로 서버전송안됨
'        vasExam.Row = asRow
'        vasExam.Col = 1
        
        'If vasExam.Value = 1 Then '체크된것만
'            If Trim(GetText(vasExam, asRow, colID)) = "" Then
'                res = GetPatientInfo(asRow)
'                If res = -1 Then
'                    SetText vasExam, "", asRow, colState
'                ElseIf res = 0 Then
'                    SetText vasExam, "미접수", asRow, colState
'                End If
'
'            End If
            
            lsPID = Trim(GetText(vasExam, asRow, colID))

            If lsPID <> "" Then
                ClearSpread vasResTemp
                
                SQL = "select examcode, equipcode, result from pat_res " & vbCrLf & _
                      "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                      "  and equipno = '" & gEquip & "' " & vbCrLf & _
                      "  and barcode = '" & lsPID & "'  "
                res = db_select_Vas(gLocal, SQL, vasResTemp)
                For j = 1 To vasResTemp.DataRowCnt
                    lsExamCode = ""
                    lsResult = ""
                    
                    lsExamCode = Trim(GetText(vasResTemp, j, 1))
                    lsResult = Trim(GetText(vasResTemp, j, 3))
                    
                    'LPD19403
                    If Left(lsExamCode, 1) <> "L" Then
                        lsExamCode = ""
                        lsResult = ""
                    End If
                    If Len(lsExamCode) < 5 Then
                        lsExamCode = ""
                        lsResult = ""
                    End If
                    If lsExamCode = "*" Then
                        lsExamCode = ""
                        lsResult = ""
                    End If
                    
                    If lsResult = "*" Then
                        lsResult = ""
                    End If
                    
                    '최종결과
                    sResEnd = lsResult
                    
                    'http://his031.knu.ac.kr/himed/webapps/com/commonweb/xrw/.live?
                    'submit_id=TXLII00101&business_id=lis&ex_interface=93367|031&bcno=E145Z0050&
                    'result=LHC1020190.420101228LHC102021.0620101228LHC1020411.620101228LHC1030229.920101228&instcd=031&eqmtcd=H05&userid=93367&paste=Y&
                    
                    If lsExamCode <> "" And lsResult <> "" Then
                        lsSenddata = lsSenddata & lsExamCode & "" & lsResult & "" & Format(Now, "yyyymmdd") & ""
                        'Local
                        SQL = "Update pat_res set " & vbCrLf & _
                              " resflag = 'B', examuid = '" & sExamUID & "' " & vbCrLf & _
                              "where examdate ='" & SeperatorCls(Text_Today) & "' " & vbCrLf & _
                              "  and barcode = '" & Trim(GetText(vasExam, i, colID)) & "' " & vbCrLf & _
                              "  and equipcode = '" & lsExamCode & "' "
                        res = SendQuery(gLocal, SQL)
                        If res = -1 Then
                            Me.MousePointer = 0
                            SaveQuery SQL
                            Exit Function
                        End If
                    End If
                Next j

                'db_Commit gServer
                Save_Raw_Data lsPID & " : " & lsSenddata
                
                kbnu_sendresult lsPID, Trim(txtUID.Text), gHPEQUIP, lsSenddata
                            
                SetText vasExam, "전송", asRow, colState
                'SetForeColor vasExam, asRow, asRow, 1, vasExam.MaxCols, 202, 202, 202
                
                SetBackColor vasExam, asRow, asRow, 1, colState, 202, 255, 112

            
                vasExam.Value = 0
            End If
'        End If
    End If
    
'    If cn_Server_Flag Then DisConnect_Server
    
'    lblUID.Visible = False
'    txtUID.Visible = False
    Me.MousePointer = 0
    
    To_Server_1 = 1
    'db_Commit gServer

End Function

Function To_Server(ByVal asRow As Long) As Integer
    Dim iStart As Long
    Dim iEnd As Long
    Dim i, j As Long
    
    Dim lsExamCode As String
    Dim lsResult As String
    Dim lsRefFlag As String
    Dim lsReceNo As String
    Dim lsReceDate As String
    Dim lsBun As String
    
    Dim sCnt As String
    
    Dim sExamUID As String
    
    Dim sResEnd As String
    Dim lsBarcode As String
    Dim lsPID As String
    
    Dim sEditDate As String
    Dim lsPanicFlag As String
    Dim lsDeltaFlag As String
    
    To_Server = -1
    
    sEditDate = Format(GetDateFull, "yyyy-mm-dd hh:mm:ss")
    
'    If Not Connect_Server Then
'        MsgBox "서버에 연결 되지 않았습니다 " & vbCrLf & "네트워크 확인 바랍니다", vbInformation, "알림"
'        cn_Server_Flag = False
'        Exit Sub
'    Else
'        cn_Server_Flag = True
'    End If
    
    Me.MousePointer = 11
    
    'db_BeginTran gServer
    
    lsBarcode = Trim(GetText(vasExam, asRow, colID))
    lsPID = Trim(GetText(vasExam, asRow, colReceNo))
    
    If lsBarcode <> "" Then
        For j = 1 To UBound(gArrEquip)
            lsResult = Trim(GetText(vasExam, asRow, colResult + (gArrEquip(j, 1) - 1) * 4))
            lsRefFlag = Trim(GetText(vasExam, asRow, colResult + (gArrEquip(j, 1) - 1) * 4 + 1))
            lsPanicFlag = Trim(GetText(vasExam, asRow, colResult + (gArrEquip(j, 1) - 1) * 4 + 2))
            lsDeltaFlag = Trim(GetText(vasExam, asRow, colResult + (gArrEquip(j, 1) - 1) * 4 + 3))
            
            lsExamCode = gArrEquip(j, 3)
            
            If lsResult <> "" Then
                '검사결과 테이블 업데이트
                SQL = " Select STATUS From TWEXAM_RESULTC " & vbCrLf & _
                      "Where SPECNO  = '" & lsBarcode & "' " & vbCrLf & _
                      "And PTNO = '" & lsPID & "' " & vbCrLf & _
                      "And SUBCODE = '" & Trim(lsExamCode) & "' "
                res = db_select_Col(gServer, SQL)
                SaveQuery Trim(gReadBuf(0)), 1
                
                If Trim(gReadBuf(0)) <> "5" Then      '결과전송된 것은 업데이트 안 되도록
                    SQL = "Update TWEXAM_RESULTC Set " & vbCrLf & _
                          "   RESULT = '" & lsResult & "', " & vbCrLf & _
                          "   RESULT2 = '" & lsResult & "', " & vbCrLf & _
                          "   EQUCODE = 'IC01', " & vbCrLf & _
                          "   STATUS = '4', " & vbCrLf & _
                          "   INTERFACEDATE = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss') " & vbCrLf & _
                          "Where SPECNO  =  '" & lsBarcode & "' " & vbCrLf & _
                          "And PTNO = '" & lsPID & "' " & vbCrLf & _
                          "And SUBCODE = '" & Trim(lsExamCode) & "' "
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        db_RollBack gServer
                        SaveQuery SQL, 1
                        Exit Function
                    End If
                End If

                'Local
                SQL = "Update pat_res set " & vbCrLf & _
                      " resflag = 'B', examuid = '" & sExamUID & "' " & vbCrLf & _
                      "where examdate ='" & SeperatorCls(Text_Today) & "' " & vbCrLf & _
                      "  and barcode = '" & Trim(GetText(vasExam, asRow, colID)) & "' " & vbCrLf & _
                      "  and equipcode = '" & lsExamCode & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    Me.MousePointer = 0
                    SaveQuery SQL
                    Exit Function
                End If
            End If
        Next j

        'db_Commit gServer
                    
        SQL = " Update TWEXAM_SPECMST Set " & vbCrLf & _
              " STATUS = '3' " & vbCrLf & _
              "Where SPECNO  =  '" & lsBarcode & "' " & vbCrLf & _
              "And PTNO = '" & lsPID & "' "
        res = SendQuery(gServer, SQL)
        If res = -1 Then
            db_RollBack gServer
            SaveQuery SQL, 1
            Exit Function
        End If
    End If
    
    SetText vasExam, "전송", asRow, colState

    vasExam.Row = asRow
    vasExam.Col = 1
    vasExam.Value = 0
    
    
    Me.MousePointer = 0
    
    To_Server = 1
    'db_Commit gServer

End Function

Function Delta_Check(ByVal asRow As Long, asRet As String, ByVal asK As Integer) As String
    Dim sPreRet As String
    Dim sExecDate As String
    
    Delta_Check = ""
    
'    If Not cn_Server_Flag Then
'        If Not Connect_Server Then
'            cn_Server_Flag = False
'            Exit Function
'        Else
'            cn_Server_Flag = True
'        End If
'    End If
    
    If Not IsNumeric(gArrEquip(asK, 9)) Then
        Exit Function
    End If
    
    If CCur(gArrEquip(asK, 9)) = 0 Then
        Exit Function
    End If
    
    ClearSpread vasTemp
    

'SELECT W.SPCID,
'                to_char(W.ORDDATE, 'yyyy-mm-dd'),
'                W.ORDSEQNO,
'                W.EXAMCODE,
'                to_char(W.WORKNO),
'                W.RSLTTEXT,
'                S.SPCCODE,
'                w.deltayn,
'                w.panicyn,
'                w.normflag
'
'           FROM SLXWORKT W, SLSPCMDT S
'          WHERE W.SPCID    in('2030166218', '1030083459')
'            AND W.EXECDATE >= to_date('20030703', 'yyyymmdd')
'            AND W.EXECDATE <  to_date('20030707', 'yyyymmdd')+1
'            AND W.PROCSTAT = 'E'
'            AND W.WORKNO between to_number('1')
'                             and to_number('9999')
'            AND S.SPCID    = W.SPCID
'          ORDER BY W.WORKNO, W.EXAMCODE;

    sExecDate = Format(Trim(Text_Today), "yyyymmdd")
    
    SQL = " Select W.RSLTTEXT " & CR & _
          " From SLXWORKT W " & CR & _
          " Where W.PATNO = '" & Trim(GetText(vasExam, asRow, colReceNo)) & "' " & CR & _
          " And W.EXAMCODE = '" & gArrEquip(asK, 2) & "' " & CR & _
          " And W.EXECDATE < to_date('" & Trim(sExecDate) & "', 'yyyymmdd') " & CR & _
          " And W.REPTDATE is not null " & CR & _
          " And W.RSLTTYPE <> 'T' " & CR & _
          " Order by execdate desc "
          
    res = db_select_Vas(gServer, SQL, vasTemp)
    
'    If cn_Server_Flag Then DisConnect_Server
    
    If vasTemp.DataRowCnt > 0 Then
        sPreRet = Trim(GetText(vasTemp, 1, 1))
        If IsNumeric(sPreRet) And IsNumeric(asRet) Then
            If CCur(sPreRet) <> 0 Then
                sPreRet = (Abs(CCur(sPreRet) - CCur(asRet)) / sPreRet) * 100
                If CCur(sPreRet) > CCur(gArrEquip(asK, 9)) Then
                    Delta_Check = "D"
                    Exit Function
                End If
            End If
        End If
    Else
        Exit Function
    End If
End Function

Function GetPatientInfo(ByVal asRow As Long) As Integer
    Dim sID As String
    Dim i As Integer

    GetPatientInfo = -1
    
    For i = 0 To 10
        gReadBuf(i) = ""
    Next i
    
    If Not cn_Server_Flag Then
        If Not Connect_Server Then
            cn_Server_Flag = False
            Exit Function
        Else
            cn_Server_Flag = True
        End If
    End If

    sID = Trim(GetText(vasExam, asRow, colID))
        
    SQL = "select a.SPECNO, a.PTNO, b.SNAME, c.AGE, c.SEX" & vbCrLf & _
      "from TWEXAM_RESULTC a, TWBAS_PATIENT b,TWEXAM_SPECMST c " & vbCrLf & _
      "Where a.SPECNO = '" & sID & "' " & vbCrLf & _
      "  and a.SUBCODE In (" & sID & ") " & vbCrLf & _
      "  and a.STATUS in ('2','3') " & vbCrLf & _
      "  and a.PTNO = b.PTNO " & vbCrLf & _
      "  and a.PTNO = c.PTNO " & vbCrLf & _
      "  and b.PTNO = c.PTNO " & vbCrLf & _
      "  and a.SPECNO =c.SPECNO" & vbCrLf & _
      " Group By a.SPECNO, a.PTNO, b.SNAME, c.AGE, c.SEX  "
    res = db_select_Col(gServer, SQL)

    If res = -1 Then
        GetPatientInfo = -1
        SaveQuery "[-1]" & SQL
    ElseIf res = 0 Then
        GetPatientInfo = 0
        SaveQuery "[0]" & SQL
    End If

    If Trim(gReadBuf(0)) = sID Then
        SetText vasExam, Trim(gReadBuf(1)), asRow, colReceNo    'colRece = colPID
        SetText vasExam, Trim(gReadBuf(2)), asRow, colPName
        SetText vasExam, Trim(gReadBuf(3)), asRow, colSex
        SetText vasExam, Trim(gReadBuf(4)), asRow, colAge

        GetPatientInfo = 1
    Else
        SaveQuery SQL
        GetPatientInfo = 0
    End If

End Function

Sub Var_Clear()
    gOrderMessage = ""
    
    sBarcode = ""
    sSeqNo = ""
    sDiskNo = ""
    sPosNo = ""
    sSampleType = ""
    lRow = -1
    sOrder = ""
    sResult = ""
    sResDateTime = ""
End Sub

Private Sub SSPanel1_DblClick()
    If Command3.Visible = False Then
        Command3.Visible = True
        txtData.Visible = True
        vasOrder.Visible = True
    Else
        Command3.Visible = False
        txtData.Visible = False
        vasOrder.Visible = False
    End If
End Sub

Private Sub Text_Today_GotFocus()
    SelectFocus Text_Today
End Sub

Private Sub Text_Today_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Command_Search_Click
    End If
End Sub

Private Sub Timer1_Timer()
'2010.08.27 이상은 추가

On Error GoTo errFind

    lblCnt.Caption = CStr(CInt(lblCnt.Caption) + 1)
    If lblCnt.Caption = "3" Then
        lblCnt.Caption = "0"
        
        If Winsock1.State = 0 Then
            Winsock1.Close
            
            Winsock1.LocalPort = gRemote.RemotePort
            Winsock1.Listen
        End If
        
        'Timer1.Enabled = False
    End If
    
errFind:

End Sub

Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        'zz
        Get_Order_barcode Trim(txtBarCode.Text)
        txtBarCode = ""
    End If
End Sub

Private Sub txtEnd_GotFocus()
    SelectFocus txtEnd
End Sub

Private Sub txtEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsNumeric(txtEnd) = False Then
            txtEnd.SetFocus
            Exit Sub
        End If
        Command1.SetFocus
    End If
End Sub

Private Sub txtStart_GotFocus()
    SelectFocus txtStart
End Sub

Private Sub txtStart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsNumeric(txtStart) = False Then
            txtStart.SetFocus
            Exit Sub
        End If
        txtEnd.SetFocus
    End If
End Sub

'Private Sub vasExam_Click(ByVal Col As Long, ByVal Row As Long)
'    Dim i As Long
'
'    If Row < 1 Or Row > vasExam.DataRowCnt Then
'        Exit Sub
'    End If
'
'    txtBarCode.Text = ""
'    txtPID_1.Text = ""
'    txtPName_1.Text = ""
'    txtPSex_1.Text = ""
'    txtPAge_1.Text = ""
'
'    ClearSpread vasRes3
'
'    txtBarCode.Text = Trim(GetText(vasExam, Row, colID))
'    txtPID_1.Text = Trim(GetText(vasExam, Row, colReceNo))
'    txtPName_1.Text = Trim(GetText(vasExam, Row, colPName))
'    txtPSex_1.Text = Trim(GetText(vasExam, Row, colSex))
'    txtPAge_1.Text = Trim(GetText(vasExam, Row, colAge))
'
'    SQL = " Select a.Examcode, '',  a.Result From pat_res a " & vbCrLf & _
'          " Where a.equipno = '" & gEquip & "' " & vbCrLf & _
'          " And a.examdate = '" & Format(Text_Today.Text, "yyyymmdd") & "' " & vbCrLf & _
'          " And a.barcode = '" & Trim(txtBarCode.Text) & "' "
'    res = db_select_Vas(gLocal, SQL, vasRes3, , 0)
'
'    For i = 1 To vasRes3.DataRowCnt
'        SQL = " Select examname From equipexam " & vbCrLf & _
'              " Where equipno = '" & gEquip & "' " & vbCrLf & _
'              " And examcode = '" & Trim(GetText(vasRes3, i, 0)) & "' "
'        res = db_select_Col(gLocal, SQL)
'
'        If gReadBuf(0) <> "" Then
'            SetText vasRes3, Trim(gReadBuf(0)), i, 1
'        End If
'    Next i
'End Sub
'
'Private Sub vasExam_DblClick(ByVal Col As Long, ByVal Row As Long)
'    Dim i, j, k As Long
'    Dim vasObj As vaSpread
'    Dim shp As String
'
'    If Row < 1 Or Row > vasExam.DataRowCnt Then
'        Exit Sub
'    End If
'
'    ClearSpread vasRes1, 1, 0
'    ClearSpread vasRes2, 1, 0
'
'    txtID = Trim(GetText(vasExam, Row, colID))
'    txtSeq = Trim(GetText(vasExam, Row, colSeqNo))
'    txtRack = Trim(GetText(vasExam, Row, colDiskNo))
'    txtPos = Trim(GetText(vasExam, Row, colPosNo))
'    txtReceNo = Trim(GetText(vasExam, Row, colReceNo))
'    txtName = Trim(GetText(vasExam, Row, colPName))
'    txtSex = Trim(GetText(vasExam, Row, colSex))
'    txtAge = Trim(GetText(vasExam, Row, colAge))
'    shp = Trim(GetText(vasExam, Row, colHospital))
'
'    txtBarCode.Text = ""
'    txtPID_1.Text = ""
'    txtPName_1.Text = ""
'    txtPSex_1.Text = ""
'    txtPAge_1.Text = ""
'
'    ClearSpread vasRes3
'
'    txtBarCode.Text = txtID.Text
'    txtPID_1.Text = txtReceNo.Text
'    txtPName_1.Text = txtName.Text
'    txtPSex_1.Text = txtSex.Text
'    txtPAge_1.Text = txtAge.Text
'
'    SQL = " Select a.ExamCode, b.ExamName, a.Result From pat_res a, equipexam b " & vbCrLf & _
'          " Where a.equipno = '" & gEquip & "' " & vbCrLf & _
'          " And a.examdate = '" & Format(Text_Today.Text, "yyyymmdd") & "' " & vbCrLf & _
'          " And a.equipno = b.equipno And a.ExamCode = b.ExamCode " & vbCrLf & _
'          " And a.barcode = '" & Trim(txtBarCode.Text) & "' "
'    res = db_select_Vas(gLocal, SQL, vasRes3, , 0)
'
''    If Not Connect_Server Then
''        MsgBox "서버에 연결 되지 않았습니다 " & vbCrLf & "네트워크 확인 바랍니다", vbInformation, "알림"
''        cn_Server_Flag = False
''        Exit Sub
''    Else
''        cn_Server_Flag = True
''    End If
'
'
'    j = 1
'    Set vasObj = vasRes1
'
'    k = 0
'
'    For i = colResult To colResult1 - 1 Step 4
'        If Trim(GetText(vasExam, Row, i)) <> "" Then
'            '검사명
'            SetText vasObj, Trim(GetText(vasExam, 0, i)), j, 0
'            '결과
'            SetText vasObj, Trim(GetText(vasExam, Row, i)), j, 1
'            '참고치 체크
'            SetText vasObj, Trim(GetText(vasExam, Row, i + 1)), j, 2
'            Select Case Trim(GetText(vasExam, Row, i + 1))
'            Case "P", "H"
'                SetForeColor vasObj, j, j, 2, 2, 255, 0, 0
'            Case "L"
'                SetForeColor vasObj, j, j, 2, 2, 0, 0, 255
'            End Select
'            '패닉 체크
'            SetText vasObj, Trim(GetText(vasExam, Row, i + 2)), j, 3
'            Select Case Trim(GetText(vasExam, Row, i + 2))
'            Case "H"
'                SetForeColor vasObj, j, j, 3, 3, 255, 0, 0
'            Case "L"
'                SetForeColor vasObj, j, j, 3, 3, 0, 0, 255
'            End Select
'            '델타 체크
'            SetText vasObj, Trim(GetText(vasExam, Row, i + 3)), j, 4
'            Select Case Trim(GetText(vasExam, Row, i + 3))
'            Case Is <> ""
'                SetForeColor vasObj, j, j, 4, 4, 255, 0, 0
'            End Select
'
'            '결과 원본
'            SetText vasObj, Trim(GetText(vasExam, Row, colResult1 + k)), j, 5
'            SetText vasObj, k, j, 6
'
'            j = j + 1
'            If j = 21 Then
'                Set vasObj = vasRes2
'                j = 1
'            End If
'        End If
'        k = k + 1
'    Next i
'
'    glRow = Row
'    sspDetail.Visible = True
'End Sub

Private Sub vasExam_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i, j As Long
    Dim k, m As Integer
    Dim sDec As String
    Dim sCnt As String
    Dim sID As String

    Dim sRefFlag As String
    Dim sPanicFlag As String
    Dim sDeltaFlag As String
    Dim sRefLow As String
    Dim sRefHigh As String
    
    Dim sResDateTime As String
    
    i = vasExam.ActiveRow
    j = vasExam.ActiveCol
    
    If KeyCode = vbKeyReturn Then
        If i < 1 Or i > vasExam.DataRowCnt Then
            Exit Sub
        End If
        
        If j < colResult Or j >= colResult1 Or (j Mod 4 <> 1) Then
            If j = colID Then
                sID = Trim(GetText(vasExam, i, colID))
                
            End If
            Exit Sub
        End If
        
        k = (j \ 4) - 4
        
        If Trim(GetText(vasExam, i, j)) <> Trim(GetText(vasExam, i, colResult1 + k)) Then
            If MsgBox("결과를 " & Trim(GetText(vasExam, i, colResult1 + k)) & " 에서 " & Trim(GetText(vasExam, i, j)) & "로 수정하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
                SetText vasExam, Trim(GetText(vasExam, i, colResult1 + k)), i, j
                Exit Sub
            End If
            
            SetForeColor vasExam, i, i, 1, 20, 0, 0, 0

            sResult = Trim(GetText(vasExam, i, j))
            
            sRefFlag = ""
            sPanicFlag = ""
            sDeltaFlag = ""
            If IsNumeric(sResult) Then
            
                SetText vasExam, sResult, i, colResult1 + k
    
                '소수자리 처리
    '            If IsNumeric(gArrEquip(k + 1, 5)) Then
    '                sDec = ""
    '                For m = 1 To CInt(gArrEquip(k + 1, 5))
    '                    sDec = sDec & "0"
    '                Next m
    '                If CInt(gArrEquip(k + 1, 5)) = 0 Then
    '                    sDec = "#0"
    '                Else
    '                    sDec = "#0." & sDec
    '                End If
    '                sResult = Format(sResult, sDec)
    '                SetText vasExam, sResult, i, j
    '            End If
    

                Select Case gArrEquip(k + 1, 2)
                Case 400    'HBsAg
                    If CCur(sResult) <= 1 Then
                        sRefFlag = "N"
                    Else
                        sRefFlag = "P"
                    End If
                    SetText vasExam, sRefFlag, i, j + 1
                Case 410    'Anti-HBs
                    If CCur(sResult) <= 10 Then
                        sRefFlag = "N"
                    Else
                        sRefFlag = "P"
                    End If
                    SetText vasExam, sRefFlag, i, j + 1
                Case 430    'Anti-Hbe
                    If CCur(sResult) >= 1 Then
                        sRefFlag = "Neg"
                    Else
                        sRefFlag = "Pos"
                    End If
                    SetText vasExam, sRefFlag, i, j + 1
                Case 440    'HBeAg
                    If CCur(sResult) <= 1 Then
                        sRefFlag = "Neg"
                    Else
                        sRefFlag = "Pos"
                    End If
                    SetText vasExam, sRefFlag, i, j + 1
                Case 460    'HBcIgM
                    If CCur(sResult) <= 1 Then
                        sRefFlag = "Neg"
                    Else
                        sRefFlag = "Pos"
                    End If
                    SetText vasExam, sRefFlag, i, j + 1
                Case Else
                    '참고치 체크================================================================
                    sRefLow = gArrEquip(k + 1, 6)
                    sRefHigh = gArrEquip(k + 1, 7)
                    If Not IsNumeric(sRefLow) Then
                        sRefLow = "0"
                    End If
                    If Not IsNumeric(sRefHigh) Then
                        sRefHigh = "0"
                    End If
                    If CCur(sRefLow) = 0 And CCur(sRefHigh) = 0 Then
                        sRefFlag = ""
                    Else
                        If CCur(sResult) < CCur(sRefLow) Then
                            sRefFlag = "Neg"
                        End If
                        If CCur(sResult) > CCur(sRefHigh) Then
                            sRefFlag = "Pos"
                        End If
                    End If
                    SetText vasExam, sRefFlag, i, j + 1
                    'Panic 체크================================================================
                    
                    If Not IsNumeric(gArrEquip(k + 1, 8)) Then
                        gArrEquip(k + 1, 8) = "0"
                    End If
                    If Not IsNumeric(gArrEquip(k + 1, 9)) Then
                        gArrEquip(k + 1, 9) = "0"
                    End If
                    If CCur(gArrEquip(k + 1, 8)) = 0 And CCur(gArrEquip(k + 1, 9)) = 0 Then
                        sPanicFlag = ""
                    Else
                        If CCur(sResult) < CCur(gArrEquip(k + 1, 8)) Then
                            sPanicFlag = "L"
                        End If
                        If CCur(sResult) > CCur(gArrEquip(k + 1, 9)) Then
                            sPanicFlag = "H"
                        End If
                    End If
                    SetText vasExam, sPanicFlag, i, j + 2
                    'Delta 체크================================================================
                    sDeltaFlag = Delta_Check(lRow, sResult, k + 1)
                    SetText vasExam, sDeltaFlag, i, j + 3
                End Select

            End If

            '결과 셀 색깔 변화=========================================================
            Select Case sRefFlag
            Case "Pos"  '"P", "H"   Positive, 'High
                SetBackColor vasExam, i, i, j, j, 246, 150, 121
            Case "Neg" ' "L"
                SetBackColor vasExam, i, i, j, j, 255, 245, 104
            Case Else   'Normal
                SetBackColor vasExam, i, i, j, j, 255, 255, 255
            End Select
            Select Case sPanicFlag
            Case "H"   'High
                SetBackColor vasExam, i, i, j, j, 242, 108, 79
            Case "L"
                SetBackColor vasExam, i, i, j, j, 60, 184, 120
            Case Else   'Normal
                SetBackColor vasExam, i, i, j, j, 255, 255, 255
            End Select
            If sDeltaFlag <> "" Then
                SetBackColor vasExam, i, i, j, j, 255, 0, 0
            Else
                SetBackColor vasExam, i, i, j, j, 255, 255, 255
            End If

'            '참고치 체크================================================================
'            sRefFlag = ""
'            sRefLow = gArrEquip(k + 1, 6)
'            sRefHigh = gArrEquip(k + 1, 7)
'            If Not IsNumeric(sRefLow) Then
'                sRefLow = "0"
'            End If
'            If Not IsNumeric(sRefHigh) Then
'                sRefHigh = "0"
'            End If
'            If CCur(sRefLow) = 0 And CCur(sRefHigh) = 0 Then
'                sRefFlag = ""
'            Else
'                If CCur(sResult) < CCur(sRefLow) Then
'                    sRefFlag = "L"
'                End If
'                If CCur(sResult) > CCur(sRefHigh) Then
'                    sRefFlag = "H"
'                End If
'            End If
'            SetText vasExam, sRefFlag, i, j + 1
'            'Panic 체크================================================================
'            sPanicFlag = ""
'            If Not IsNumeric(gArrEquip(k + 1, 8)) Then
'                gArrEquip(k + 1, 8) = "0"
'            End If
'            If Not IsNumeric(gArrEquip(k + 1, 9)) Then
'                gArrEquip(k + 1, 9) = "0"
'            End If
'            If CCur(gArrEquip(k + 1, 8)) = 0 And CCur(gArrEquip(k + 1, 9)) = 0 Then
'                sPanicFlag = ""
'            Else
'                If CCur(sResult) < CCur(gArrEquip(k + 1, 8)) Then
'                    sPanicFlag = "L"
'                End If
'                If CCur(sResult) > CCur(gArrEquip(k + 1, 9)) Then
'                    sPanicFlag = "H"
'                End If
'            End If
'            SetText vasExam, sPanicFlag, i, j + 2
'            'Delta 체크================================================================
'            sDeltaFlag = Delta_Check(i, sResult, k + 1)
'            SetText vasExam, sDeltaFlag, i, j + 3
'
'            Select Case sRefFlag
'            Case "H"
'                SetForeColor vasExam, i, i, j + 1, j + 1, 255, 0, 0
'            Case "L"
'                SetForeColor vasExam, i, i, j + 1, j + 1, 0, 0, 255
'            Case Else
'                SetForeColor vasExam, i, i, j + 1, j + 1, 0, 0, 0
'            End Select
'
'            Select Case sPanicFlag
'            Case "H"
'                SetForeColor vasExam, i, i, j + 2, j + 2, 255, 0, 0
'            Case "L"
'                SetForeColor vasExam, i, i, j + 2, j + 2, 0, 0, 255
'            Case Else
'                SetForeColor vasExam, i, i, j + 2, j + 2, 0, 0, 0
'            End Select
'
'            Select Case sDeltaFlag
'            Case Is <> ""
'                SetForeColor vasExam, i, i, j + 3, j + 3, 255, 0, 0
'            Case Else
'                SetForeColor vasExam, i, i, j + 3, j + 3, 0, 0, 0
'            End Select

            SetText vasExam, "수정", i, colState

            'Local Table Insert
            '환자 데이타 ====================================================================================
            db_BeginTran gLocal

            sCnt = ""
            SQL = "Select count(*) from pat_res " & vbCrLf & _
                  "where barcode = '" & Trim(GetText(vasExam, i, colID)) & "' and equipcode = '" & gArrEquip(k + 1, 2) & "' "
            res = db_select_Var(gLocal, SQL, sCnt)
            If res <= 0 Then
                SaveQuery SQL
                db_RollBack gLocal
                Exit Sub
            End If
            If CInt(sCnt) = 0 Then
                res = GetPatientInfo(i)
                If Not IsNumeric(GetText(vasExam, i, colAge)) Then
                    SetText vasExam, "0", i, colAge
                End If
                '입력
                SQL = "insert into pat_res (examdate, equipno, seqno, diskno, posno, " & vbCrLf & _
                      " barcode, equipcode, examcode, " & vbCrLf & _
                      " result, refflag, refvalue, " & vbCrLf & _
                      " receno, recedate, pid, pname, jumin1, jumin2, " & _
                      " psex, page, resflag, resdate, panicflag, deltaflag, " & _
                      " ) " & vbCrLf & _
                      "values ('" & SeperatorCls(Text_Today) & "', '" & gEquip & "', '" & Trim(GetText(vasExam, i, colSeqNo)) & "', '" & Trim(GetText(vasExam, i, colDiskNo)) & "', '" & Trim(GetText(vasExam, i, colPosNo)) & "', " & _
                      " '" & Trim(GetText(vasExam, i, colID)) & "', '" & gArrEquip(k + 1, 2) & "', '" & gArrEquip(k + 1, 3) & "', " & _
                      " '" & Trim(GetText(vasExam, i, j)) & "', '" & Trim(GetText(vasExam, i, j + 1)) & "', '" & sRefLow & " - " & sRefHigh & "', " & _
                      " '" & Trim(GetText(vasExam, i, colReceNo)) & "', '" & Trim(GetText(vasExam, i, colReceDate)) & "', '" & Trim(GetText(vasExam, i, colPID)) & "', '" & Trim(GetText(vasExam, i, colPName)) & "', '" & Trim(GetText(vasExam, i, colJumin1)) & "', '" & Trim(GetText(vasExam, i, colJumin2)) & "', " & _
                      " '" & Trim(GetText(vasExam, i, colSex)) & "', " & Trim(GetText(vasExam, i, colAge)) & ", 'A', '" & GetDateFull & "', '" & Trim(GetText(vasExam, i, j + 2)) & "', '" & Trim(GetText(vasExam, i, j + 3)) & "' )"
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    db_RollBack gLocal
                    Exit Sub
                End If
            ElseIf CInt(sCnt) > 0 Then
                '수정
                SQL = "Update pat_res set " & vbCrLf & _
                      " result =  '" & sResult & "', " & vbCrLf & _
                      " refflag =  '" & sRefFlag & "', " & vbCrLf & _
                      " panicflag = '" & Trim(GetText(vasExam, i, j + 2)) & "', " & vbCrLf & _
                      " deltaflag = '" & Trim(GetText(vasExam, i, j + 3)) & "', " & vbCrLf & _
                      " resflag = 'A' " & vbCrLf & _
                      "where examdate ='" & SeperatorCls(Text_Today) & "' " & vbCrLf & _
                      "  and barcode = '" & Trim(GetText(vasExam, i, colID)) & "' " & vbCrLf & _
                      "  and equipcode = '" & gArrEquip(k + 1, 2) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    db_RollBack gLocal
                    Exit Sub
                End If
            End If

            db_Commit gLocal

        End If
    End If
    If KeyCode = vbKeyDelete Then
        If i < 1 Or i > vasExam.DataRowCnt Then
            Exit Sub
        End If
        
        If j < colResult Or j >= colResult1 Or (j Mod 4 <> 1) Then
            Exit Sub
        End If
        
        k = (j \ 4) - 4
        sResult = GetText(vasExam, i, j)
        sResDateTime = Trim(GetText(vasExam, i, colReceDate))
        
        If MsgBox("선택하신 QC 결과를 삭제하시겠습니까?", vbOKCancel, "알림") = vbCancel Then
            Exit Sub
        End If
        
        If Trim(GetText(vasExam, i, colSampleType)) = "Q" Then
            SQL = "delete from qc_res " & vbCrLf & _
                  "where equipno = '" & gEquip & "' " & vbCrLf & _
                  "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
                  "  and examtime = '" & Mid(sResDateTime, 9, 6) & "' " & vbCrLf & _
                  "  and levelname = '" & Trim(GetText(vasExam, i, colID)) & "' " & vbCrLf & _
                  "  and equipcode = '" & gArrEquip(k + 1, 2) & "' "
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                db_RollBack gLocal
                SaveQuery SQL
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub vasRes1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i, j As Long
    Dim k, m As Integer
    
    Dim sDec As String
    Dim sCnt As String
    
    Dim sRefFlag As String
    Dim sPanicFlag As String
    Dim sDeltaFlag As String
    Dim sRefLow As String
    Dim sRefHigh As String
    
    i = vasRes1.ActiveRow
    j = vasRes1.ActiveCol
    
    If KeyCode = vbKeyReturn Then
        If i < 1 Or i > vasRes1.DataRowCnt Then
            Exit Sub
        End If
        
        k = CInt(GetText(vasRes1, i, 6))
        
        If Trim(GetText(vasRes1, i, j)) <> Trim(GetText(vasRes1, i, 5)) Then
            If MsgBox("결과를 " & Trim(GetText(vasRes1, i, 5)) & " 에서 " & Trim(GetText(vasRes1, i, j)) & "로 수정하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
                SetText vasRes1, Trim(GetText(vasRes1, i, 5)), i, j
                Exit Sub
            End If
            
            sResult = Trim(GetText(vasRes1, i, j))
            
            If IsNumeric(sResult) Then

                SetText vasRes1, sResult, i, 5
    
                '소수자리 처리
    '            If IsNumeric(gArrEquip(k, 5)) Then
    '                sDec = ""
    '                For m = 1 To CInt(gArrEquip(k, 5))
    '                    sDec = sDec & "0"
    '                Next m
    '                If CInt(gArrEquip(k, 5)) = 0 Then
    '                    sDec = "#0"
    '                Else
    '                    sDec = "#0." & sDec
    '                End If
    '                sResult = Format(sResult, sDec)
    '                SetText vasRes1, sResult, i, j
    '            End If
    
                sRefFlag = ""
                sPanicFlag = ""
                sDeltaFlag = ""
                Select Case gArrEquip(k, 2)
                Case 400    'HBsAg
                    If CCur(sResult) <= 1 Then
                        sRefFlag = "N"
                    Else
                        sRefFlag = "P"
                    End If
                    SetText vasRes1, sRefFlag, i, j + 1
                Case 410    'Anti-HBs
                    If CCur(sResult) <= 10 Then
                        sRefFlag = "N"
                    Else
                        sRefFlag = "P"
                    End If
                    SetText vasRes1, sRefFlag, i, j + 1
                Case 430    'Anti-Hbe
                    If CCur(sResult) >= 1 Then
                        sRefFlag = "N"
                    Else
                        sRefFlag = "P"
                    End If
                    SetText vasRes1, sRefFlag, i, j + 1
                Case 440    'HBeAg
                    If CCur(sResult) <= 1 Then
                        sRefFlag = "N"
                    Else
                        sRefFlag = "P"
                    End If
                    SetText vasRes1, sRefFlag, i, j + 1
                Case Else
                    '참고치 체크================================================================
                    sRefLow = gArrEquip(k, 6)
                    sRefHigh = gArrEquip(k, 7)
                    If Not IsNumeric(sRefLow) Then
                        sRefLow = "0"
                    End If
                    If Not IsNumeric(sRefHigh) Then
                        sRefHigh = "0"
                    End If
                    If CCur(sRefLow) = 0 And CCur(sRefHigh) = 0 Then
                        sRefFlag = ""
                    Else
                        If CCur(sResult) < CCur(sRefLow) Then
                            sRefFlag = "L"
                        End If
                        If CCur(sResult) > CCur(sRefHigh) Then
                            sRefFlag = "H"
                        End If
                    End If
                    SetText vasRes1, sRefFlag, i, j + 1
                    'Panic 체크================================================================
                    If Not IsNumeric(gArrEquip(k, 8)) Then
                        gArrEquip(k, 8) = "0"
                    End If
                    If Not IsNumeric(gArrEquip(k, 9)) Then
                        gArrEquip(k, 9) = "0"
                    End If
                    If CCur(gArrEquip(k, 8)) = 0 And CCur(gArrEquip(k, 9)) = 0 Then
                        sPanicFlag = ""
                    Else
                        If CCur(sResult) < CCur(gArrEquip(k, 8)) Then
                            sPanicFlag = "L"
                        End If
                        If CCur(sResult) > CCur(gArrEquip(k, 9)) Then
                            sPanicFlag = "H"
                        End If
                    End If
                    SetText vasRes1, sPanicFlag, i, j + 2
                    'Delta 체크================================================================
                    sDeltaFlag = Delta_Check(i, sResult, k)
                    SetText vasRes1, sDeltaFlag, i, j + 3
                End Select
                
                SetText vasExam, "수정", glRow, colState
                
                SetText vasExam, sResult, glRow, colResult + 4 * k
                SetText vasExam, sRefFlag, glRow, colResult + (4 * k) + 1
                SetText vasExam, sPanicFlag, glRow, colResult + (4 * k) + 2
                SetText vasExam, sDeltaFlag, glRow, colResult + (4 * k) + 3
                SetText vasExam, sResult, glRow, colResult1 + k
                
                '결과 셀 색깔 변화=========================================================
                Select Case sRefFlag
                Case "P", "H"   'Positive, 'High
                    SetBackColor vasExam, glRow, glRow, colResult + (4 * k) + 1, colResult + (4 * k) + 1, 246, 150, 121
                Case "L"
                    SetBackColor vasExam, glRow, glRow, colResult + (4 * k) + 1, colResult + (4 * k) + 1, 255, 245, 104
                Case Else   'Normal
                End Select
                Select Case sPanicFlag
                Case "H"   'High
                    SetBackColor vasExam, glRow, glRow, colResult + (4 * k) + 2, colResult + (4 * k) + 2, 242, 108, 79
                Case "L"
                    SetBackColor vasExam, glRow, glRow, colResult + (4 * k) + 2, colResult + (4 * k) + 2, 60, 184, 120
                Case Else   'Normal
                End Select
                If sDeltaFlag <> "" Then
                    SetBackColor vasExam, glRow, glRow, colResult + (4 * k) + 3, colResult + (4 * k) + 3, 255, 0, 0
                End If
            End If
            'Local Table Insert
            '환자 데이타 ====================================================================================
            db_BeginTran gLocal

            sCnt = ""
            SQL = "Select count(*) from pat_res " & vbCrLf & _
                  "where barcode = '" & Trim(txtID.Text) & "' "
            res = db_select_Var(gLocal, SQL, sCnt)
            If res <= 0 Then
                SaveQuery SQL
                db_RollBack gLocal
                Exit Sub
            End If
            If CInt(sCnt) = 0 Then
                res = GetPatientInfo(glRow)
                If Not IsNumeric(GetText(vasExam, glRow, colAge)) Then
                    SetText vasExam, "0", glRow, colAge
                End If
                '입력
                SQL = "insert into pat_res (examdate, equipno, seqno, diskno, posno, " & vbCrLf & _
                      " barcode, equipcode, examcode, " & vbCrLf & _
                      " result, refflag, refvalue, " & vbCrLf & _
                      " receno, recedate, pid, pname, jumin1, jumin2, " & _
                      " psex, page, resflag, resdate, panicflag, deltaflag, " & _
                      " ) " & vbCrLf & _
                      "values ('" & SeperatorCls(Text_Today) & "', '" & gEquip & "', '" & Trim(GetText(vasExam, glRow, colSeqNo)) & "', '" & Trim(GetText(vasExam, glRow, colDiskNo)) & "', '" & Trim(GetText(vasExam, glRow, colPosNo)) & "', " & _
                      " '" & Trim(GetText(vasExam, glRow, colID)) & "', '" & gArrEquip(k, 2) & "', '" & gArrEquip(k, 3) & "', " & _
                      " '" & sResult & "', '" & sRefFlag & "', '" & sRefLow & " - " & sRefHigh & "', " & _
                      " '" & Trim(GetText(vasExam, glRow, colReceNo)) & "', '" & Trim(GetText(vasExam, glRow, colReceDate)) & "', '" & Trim(GetText(vasExam, glRow, colPID)) & "', '" & Trim(GetText(vasExam, glRow, colPName)) & "', '" & Trim(GetText(vasExam, glRow, colJumin1)) & "', '" & Trim(GetText(vasExam, glRow, colJumin2)) & "', " & _
                      " '" & Trim(GetText(vasExam, glRow, colSex)) & "', " & Trim(GetText(vasExam, glRow, colAge)) & ", 'A', '" & GetDateFull & "', '" & sPanicFlag & "', '" & sDeltaFlag & "' )"
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    db_RollBack gLocal
                    Exit Sub
                End If
            ElseIf CInt(sCnt) > 0 Then
                '수정
                SQL = "Update pat_res set " & vbCrLf & _
                      " result =  '" & sResult & "', " & vbCrLf & _
                      " refflag =  '" & sRefFlag & "', " & vbCrLf & _
                      " panicflag = '" & sPanicFlag & "', " & vbCrLf & _
                      " deltaflag = '" & sDeltaFlag & "', " & vbCrLf & _
                      " resflag = 'A' " & vbCrLf & _
                      "where examdate ='" & SeperatorCls(Text_Today) & "' " & vbCrLf & _
                      "  and barcode = '" & Trim(GetText(vasExam, glRow, colID)) & "' " & vbCrLf & _
                      "  and equipcode = '" & gArrEquip(k, 2) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    db_RollBack gLocal
                    Exit Sub
                End If
            End If

            db_Commit gLocal

        End If
    End If
End Sub

Private Sub vasRes2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i, j As Long
    Dim k, m As Integer
    
    Dim sDec As String
    Dim sCnt As String
    
    Dim sRefFlag As String
    Dim sPanicFlag As String
    Dim sDeltaFlag As String
    Dim sRefLow As String
    Dim sRefHigh As String
    
    i = vasRes2.ActiveRow
    j = vasRes2.ActiveCol
    
    If KeyCode = vbKeyReturn Then
        If i < 1 Or i > vasRes2.DataRowCnt Then
            Exit Sub
        End If
        
        k = CInt(GetText(vasRes2, i, 6))
        
        If Trim(GetText(vasRes2, i, j)) <> Trim(GetText(vasRes2, i, 5)) Then
            If MsgBox("결과를 " & Trim(GetText(vasRes2, i, 5)) & " 에서 " & Trim(GetText(vasRes2, i, j)) & "로 수정하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
                SetText vasRes2, Trim(GetText(vasRes2, i, 5)), i, j
                Exit Sub
            End If
            
            sResult = Trim(GetText(vasRes2, i, j))
            
            If Not IsNumeric(sResult) Then

                SetText vasRes2, sResult, i, 5
    
                '소수자리 처리
    '            If IsNumeric(gArrEquip(k, 5)) Then
    '                sDec = ""
    '                For m = 1 To CInt(gArrEquip(k, 5))
    '                    sDec = sDec & "0"
    '                Next m
    '                If CInt(gArrEquip(k, 5)) = 0 Then
    '                    sDec = "#0"
    '                Else
    '                    sDec = "#0." & sDec
    '                End If
    '                sResult = Format(sResult, sDec)
    '                SetText vasRes2, sResult, i, j
    '            End If
    
                sRefFlag = ""
                sPanicFlag = ""
                sDeltaFlag = ""
                Select Case gArrEquip(k, 2)
                Case 400    'HBsAg
                    If CCur(sResult) <= 1 Then
                        sRefFlag = "N"
                    Else
                        sRefFlag = "P"
                    End If
                    SetText vasRes2, sRefFlag, i, j + 1
                Case 410    'Anti-HBs
                    If CCur(sResult) <= 10 Then
                        sRefFlag = "N"
                    Else
                        sRefFlag = "P"
                    End If
                    SetText vasRes2, sRefFlag, i, j + 1
                Case 430    'Anti-Hbe
                    If CCur(sResult) >= 1 Then
                        sRefFlag = "N"
                    Else
                        sRefFlag = "P"
                    End If
                    SetText vasRes2, sRefFlag, i, j + 1
                Case 440    'HBeAg
                    If CCur(sResult) <= 1 Then
                        sRefFlag = "N"
                    Else
                        sRefFlag = "P"
                    End If
                    SetText vasRes2, sRefFlag, i, j + 1
                Case Else
                    '참고치 체크================================================================
                    sRefLow = gArrEquip(k, 6)
                    sRefHigh = gArrEquip(k, 7)
                    If Not IsNumeric(sRefLow) Then
                        sRefLow = "0"
                    End If
                    If Not IsNumeric(sRefHigh) Then
                        sRefHigh = "0"
                    End If
                    If CCur(sRefLow) = 0 And CCur(sRefHigh) = 0 Then
                        sRefFlag = ""
                    Else
                        If CCur(sResult) < CCur(sRefLow) Then
                            sRefFlag = "L"
                        End If
                        If CCur(sResult) > CCur(sRefHigh) Then
                            sRefFlag = "H"
                        End If
                    End If
                    SetText vasRes2, sRefFlag, i, j + 1
                    'Panic 체크================================================================
                    If Not IsNumeric(gArrEquip(k, 8)) Then
                        gArrEquip(k, 8) = "0"
                    End If
                    If Not IsNumeric(gArrEquip(k, 9)) Then
                        gArrEquip(k, 9) = "0"
                    End If
                    If CCur(gArrEquip(k, 8)) = 0 And CCur(gArrEquip(k, 9)) = 0 Then
                        sPanicFlag = ""
                    Else
                        If CCur(sResult) < CCur(gArrEquip(k, 8)) Then
                            sPanicFlag = "L"
                        End If
                        If CCur(sResult) > CCur(gArrEquip(k, 9)) Then
                            sPanicFlag = "H"
                        End If
                    End If
                    SetText vasRes2, sPanicFlag, i, j + 2
                    'Delta 체크================================================================
                    sDeltaFlag = Delta_Check(i, sResult, k)
                    SetText vasRes2, sDeltaFlag, i, j + 3
                End Select
                
                SetText vasExam, "수정", glRow, colState
                
                SetText vasExam, sResult, glRow, colResult + 4 * k
                SetText vasExam, sRefFlag, glRow, colResult + (4 * k) + 1
                SetText vasExam, sPanicFlag, glRow, colResult + (4 * k) + 2
                SetText vasExam, sDeltaFlag, glRow, colResult + (4 * k) + 3
                SetText vasExam, sResult, glRow, colResult1 + k
                
                '결과 셀 색깔 변화=========================================================
                Select Case sRefFlag
                Case "P", "H"   'Positive, 'High
                    SetBackColor vasExam, glRow, glRow, colResult + (4 * k) + 1, colResult + (4 * k) + 1, 246, 150, 121
                Case "L"
                    SetBackColor vasExam, glRow, glRow, colResult + (4 * k) + 1, colResult + (4 * k) + 1, 255, 245, 104
                Case Else   'Normal
                End Select
                Select Case sPanicFlag
                Case "H"   'High
                    SetBackColor vasExam, glRow, glRow, colResult + (4 * k) + 2, colResult + (4 * k) + 2, 242, 108, 79
                Case "L"
                    SetBackColor vasExam, glRow, glRow, colResult + (4 * k) + 2, colResult + (4 * k) + 2, 60, 184, 120
                Case Else   'Normal
                End Select
                If sDeltaFlag <> "" Then
                    SetBackColor vasExam, glRow, glRow, colResult + (4 * k) + 3, colResult + (4 * k) + 3, 255, 0, 0
                End If
            End If
            
            'Local Table Insert
            '환자 데이타 ====================================================================================
            db_BeginTran gLocal

            sCnt = ""
            SQL = "Select count(*) from pat_res " & vbCrLf & _
                  "where barcode = '" & Trim(txtID.Text) & "' "
            res = db_select_Var(gLocal, SQL, sCnt)
            If res <= 0 Then
                SaveQuery SQL
                db_RollBack gLocal
                Exit Sub
            End If
            If CInt(sCnt) = 0 Then
                res = GetPatientInfo(glRow)
                If Not IsNumeric(GetText(vasExam, glRow, colAge)) Then
                    SetText vasExam, "0", glRow, colAge
                End If
                '입력
                SQL = "insert into pat_res (examdate, equipno, seqno, diskno, posno, " & vbCrLf & _
                      " barcode, equipcode, examcode, " & vbCrLf & _
                      " result, refflag, refvalue, " & vbCrLf & _
                      " receno, recedate, pid, pname, jumin1, jumin2, " & _
                      " psex, page, resflag, resdate, panicflag, deltaflag, " & _
                      " ) " & vbCrLf & _
                      "values ('" & SeperatorCls(Text_Today) & "', '" & gEquip & "', '" & Trim(GetText(vasExam, glRow, colSeqNo)) & "', '" & Trim(GetText(vasExam, glRow, colDiskNo)) & "', '" & Trim(GetText(vasExam, glRow, colPosNo)) & "', " & _
                      " '" & Trim(GetText(vasExam, glRow, colID)) & "', '" & gArrEquip(k, 2) & "', '" & gArrEquip(k, 3) & "', " & _
                      " '" & sResult & "', '" & sRefFlag & "', '" & sRefLow & " - " & sRefHigh & "', " & _
                      " '" & Trim(GetText(vasExam, glRow, colReceNo)) & "', '" & Trim(GetText(vasExam, glRow, colReceDate)) & "', '" & Trim(GetText(vasExam, glRow, colPID)) & "', '" & Trim(GetText(vasExam, glRow, colPName)) & "', '" & Trim(GetText(vasExam, glRow, colJumin1)) & "', '" & Trim(GetText(vasExam, glRow, colJumin2)) & "', " & _
                      " '" & Trim(GetText(vasExam, glRow, colSex)) & "', " & Trim(GetText(vasExam, glRow, colAge)) & ", 'A', '" & GetDateFull & "', '" & sPanicFlag & "', '" & sDeltaFlag & "' )"
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    db_RollBack gLocal
                    Exit Sub
                End If
            ElseIf CInt(sCnt) > 0 Then
                '수정
                SQL = "Update pat_res set " & vbCrLf & _
                      " result =  '" & sResult & "', " & vbCrLf & _
                      " refflag =  '" & sRefFlag & "', " & vbCrLf & _
                      " panicflag = '" & sPanicFlag & "', " & vbCrLf & _
                      " deltaflag = '" & sDeltaFlag & "', " & vbCrLf & _
                      " resflag = 'A' " & vbCrLf & _
                      "where examdate ='" & SeperatorCls(Text_Today) & "' " & vbCrLf & _
                      "  and barcode = '" & Trim(GetText(vasExam, glRow, colID)) & "' " & vbCrLf & _
                      "  and equipcode = '" & gArrEquip(k, 2) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    db_RollBack gLocal
                    Exit Sub
                End If
            End If

            db_Commit gLocal

        End If
    End If

End Sub

Private Sub Winsock1_Close()
    StatusBar1.Panels(1).Text = "장비와의 연결이 끊어졌습니다"
End Sub

Private Sub Winsock1_Connect()
    StatusBar1.Panels(1).Text = "포트로 연결되었습니다"
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Close
    Winsock1.Accept requestID
    
    'Timer1.Enabled = True       '2010.08.27 이상은 추가
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim sTmp As String
    
    Dim sPID As String
    Dim sSendData As String
    Dim sSndMessage As String
    Dim i As Integer
    Dim iRow As Integer
    Dim lResRow As Long
    
    Dim sExamCode As String
    Dim sExamName As String
    Dim sResult As String
    
    On Error GoTo ErrHandle:
    
    'sAck = "MSH|^~\&|LIS-Server|Wellness Hospital|Mindray|BS-300|20090131111349||ACK^R01|3|P|2.3.1|MSA|AA|3||||0|" & chrCR
    'sAck = "MSH|^~\&|||Mindray|BS-380|" & Format(GetDateFull, "YYYYMMDDHHMMSS") & "||ACK^R01|1|P|2.3.1|MSA|AA|3||||0|" & chrCR
        
'    If Text_Today <> Format(GetDateFull, "YYYY-MM-DD") Then
'        Text_Today = Format(GetDateFull, "YYYY-MM-DD")
'    End If
    
    Winsock1.GetData sTmp

'    Save_Raw_Data "[RX:" & Format(Time, "hh:nn:ss") & "]" & sTmp

'    i = InStr(1, sTmp, chrCR)
'    If i > 0 Then
'        gAllData1 = gAllData1 & sTmp
'        txtData = gAllData1
'
''        BS400 gAllData1
'        Bioneer gAllData1
'        gAllData1 = ""
'
''        Winsock1.SendData sACK
''
''        Save_Raw_Data "[TX:" & Format(Time, "hh:nn:ss") & "]" & sAck
'    End If
'
''    i = InStr(1, sTmp, Chr(6))
''    If i > 0 Then
''        If vasTmp.DataRowCnt > 0 Then
''            Winsock1.SendData GetText(vasTmp, 1, 1)
''            Save_Raw_Data "[TX:" & Format(Time, "hh:nn:ss") & "]" & GetText(vasTmp, 1, 1)
''            DeleteRow vasTmp, 1, 1
''        End If
''    End If
'
'    Exit Sub

    For i = 1 To Len(sTmp)
    
        Select Case Mid(sTmp, i, 1)
      
        Case chrENQ
            Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & chrENQ
                    
            gSndState = ""
            gENQFlag = 9
            
            gRecodeType = ""
    
            Winsock1.SendData chrACK
            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
            
            gPreSpecID = ""
            gPreRow = 0
            
            
        Case chrACK
            'WorkList(Rack) 형식인 경우=================================================
            
            SaveData "[RX" & Format(Time, "hh:mm:ss") & "]" & chrACK
            gOrdRow = gOrdRow + 1
    
            If GetText(vasOrder, gOrdRow, 1) = "" Then
                Exit Sub
            End If
            
            If gOrdRow <= vasOrder.DataRowCnt Then
                
                sSendData = Trim(GetText(vasOrder, gOrdRow, 1))
                        
                Winsock1.SendData sSendData
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & sSendData
                
                If gOrdRow = vasOrder.DataRowCnt Then
                    
                    ClearSpread vasOrder
                    
                    Me.MousePointer = 0
    
                End If
            End If
                    
        Case chrSTX     '자료수신 시작
            txtData.Text = Mid(sTmp, i, 1)
            
        Case chrETX
            txtData.Text = txtData.Text & Mid(sTmp, i, 1)
        
        Case chrLF
            txtData.Text = txtData.Text & Mid(sTmp, i, 1)
            Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & txtData.Text
            
            Bioneer txtData.Text
            
            
            Winsock1.SendData chrACK
            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
            
        Case chrEOT     '자료수신 완료
            If gRecodeType = "R" Then
    '            If gPreRow > 0 And gPreRow <= vaslist.DataRowCnt Then
    '
    '                If chkMode.Value = 1 Then
    '                    'res = Insert_Data(gPreRow)
    '                    Res = ToServer(gPreRow)
    '                    If Res = 1 Then
    '                        SetBackColor vaslist, gPreRow, gPreRow, colCheckBox, colCheckBox, 202, 255, 112
    '                        SetText vaslist, "완료", gPreRow, colState
    '                    Else
    '                        SetBackColor vaslist, gPreRow, gPreRow, colCheckBox, colCheckBox, 255, 0, 0
    '                        SetText vaslist, "실패", gPreRow, colState
    '                    End If
    '                End If
    '            End If
                gSndState = "R"
            ElseIf gRecodeType = "Q" Then
                gOrdRow = 0
                gPreMsg = chrENQ
                
                Winsock1.SendData chrENQ
'                ARCHITECT "[TX" & Format(Time, "hh:mm:ss") & "]" & chrENQ
                        
                'gENQFlag = 1
                gSndState = "Q"
                gPreMsg = chrENQ
            End If
    
    
            gMsgFlag = ""
            'gHeadRecode = ""
            txtData.Text = ""
                
        Case Else
            txtData.Text = txtData.Text & Mid(sTmp, i, 1)
        End Select
    Next i

ErrHandle:
    Exit Sub

'Dim GetDat As String
'Dim i As Integer
'
'On Error GoTo ErrHandle:
'
'    Winsock1.GetData GetDat
'
'    txtData.Text = txtData.Text & GetDat
'
'    Save_Raw_Data "[RX:" & Format(Time, "hh:nn:ss") & "]" & txtData
'
'    Call BS380(txtData.Text)
'
'
'ErrHandle:
'    Exit Sub
End Sub

Sub Bioneer(argData As String)
    Dim i, j, k, z, m As Integer
    Dim iCnt As Integer
    Dim jCnt As Integer
    Dim aCnt As Integer
    Dim bCnt As Integer
    Dim sCnt As String
    
    Dim lsTmp As String

    Dim sDate As String '필요없을 수도 있음
    Dim sGubun As String
    Dim sPID As String
    Dim sReceNo As String
    Dim sSpecID As String
    Dim lsEquipCode As String
    Dim lsEquip As String
    
    Dim sExamCode As String
    Dim sExamName As String
    Dim sResClassCode As String
    Dim sFlag As String
    Dim lsResult As String
    Dim lsResFlag As String
    Dim sQC As String
    
    Dim sRefFlag As String
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sPanicFlag As String
    Dim sDeltaFlag As String
    
    Dim sExamDate   As String
    Dim sBarcode    As String
    Dim sPname      As String
    Dim sPSex       As String
    Dim sPage       As String
    
    Dim sAg_Res As String
    Dim sAb_Res As String
    
    Dim sGiho As String
    
    Dim sExamCode_All As String
    
    Dim lRow As Long
    Dim lCol As Long
    
    Dim lResRow As Long
    
    Dim slen, sLen2 As String
    Dim iRCnt As Integer
    
    Dim liEquipCode
    Dim mExam As Variant
    Dim lsOrder() As String
    Dim sPName_E As String
        

    
    Select Case Mid(argData, 3, 1)
    Case "H"    'Header
'        If chkMode.Value = 1 And gPreRow > 0 And gPreRow <= vaslist.DataRowCnt Then
'            'res = Insert_Data(gPreRow)
'            Res = ToServer(gPreRow)
'            If Res = 1 Then
'                SetBackColor vaslist, gPreRow, gPreRow, colCheckBox, colCheckBox, 202, 255, 112
'                SetText vaslist, "완료", gPreRow, colState
'            Else
'                SetBackColor vaslist, gPreRow, gPreRow, colCheckBox, colCheckBox, 255, 0, 0
'                SetText vaslist, "실패", gPreRow, colState
'            End If
'        End If
        
        
        gPreRow = -1
        glRow = -1
        iCnt = 0
'
'        i = InStr(1, argData, "|")
'        Do While i > 0
'            iCnt = iCnt + 1
'            Select Case iCnt
'            Case 5  'Equip Version
'                gVersion = Left(argData, i - 1)
'            Case 15 'DateTime
'                gDateTime = Left(argData, i - 1)
'                Exit Do
'            End Select
'            argData = Mid(argData, i + 1)
'            i = InStr(1, argData, "|")
'        Loop
        
        
    Case "P"    'Patient
        gPatFlag = -1
        
        If optOption(0).Value = True Then
            If glRow > 0 And glRow <= vasExam.DataRowCnt Then
                To_Server_1 glRow
            End If
        End If
        
        'ClearSpread vasRes
    Case "O"    'Order
        iCnt = 0
            
            
        lsEquipCode = ""
        lsResult = ""
        
        lsIPCHBVRes = ""
        lsIPCHBVCt = ""
        
    
        i = InStr(1, argData, "|")
        Do While i > 0
            iCnt = iCnt + 1
            Select Case iCnt
            Case 3  '검체번호
                lsTmp = Left(argData, i - 1)
                j = InStr(1, lsTmp, "^")
                sPID = Left(lsTmp, j - 1)
                sSpecID = sPID
                gSpecID = sSpecID
                
'            Case 4  'sample position
'                lsTmp = Left(argData, i - 1)
'                j = InStr(1, lsTmp, "^")
'                'gCup = Left(lsTmp, j - 1)
'                lsTmp = Mid(lsTmp, j + 1)
'                j = InStr(1, lsTmp, "^")
'                'gPos = Mid(lsTmp, j - 1)
'                If j > 0 Then
'                    gCup = Left(lsTmp, j - 1)
'                    gPos = Mid(lsTmp, j + 1)
'                End If
'
'            Case 5  'TestID
''                lsTmp = Left(ArgData, i - 1)
''                lsTmp = Mid(lsTmp, 4)
''                j = InStr(1, lsTmp, "^")
''                lsEquipCode = Left(lsTmp, j - 1)
''                gTestID = lsEquipCode
'
'            Case 12
'                lsTmp = Left(argData, i - 1)
'                sQC = lsTmp
'                gQC = sQC
                
                Exit Do
                
            End Select
            
            argData = Mid(argData, i + 1)
            i = InStr(1, argData, "|")
        Loop
                
        glRow = -1
        For i = 1 To vasExam.DataRowCnt
            If Trim(GetText(vasExam, i, colID)) = gSpecID Then
                glRow = i
                
                If gPatFlag = -1 Then
                    vasActiveCell vasExam, glRow, 2
                    'vasList_Click colBarCode, glRow
                    'SetText vasList, vasRes.DataRowCnt, glRow, colRCnt
'                    SetText vasExam, gCup, glRow, 6
'                    SetText vasExam, gPos, glRow, 7
                    
                    gPatFlag = 1
                End If
                
                Exit For
            End If
        Next i
        
        If glRow = -1 Then  ' vaslist에 없는 검체의 결과가 나올 때 데이터 추가
            glRow = vasExam.DataRowCnt + 1
            If glRow > vasExam.MaxRows Then
                vasExam.MaxRows = glRow + 1
            End If
            vasActiveCell vasExam, glRow, 2
            SetText vasExam, gSpecID, glRow, 2
'            SetText vasExam, gCup, glRow, 6
'            SetText vasExam, gPos, glRow, 7
        End If
        gOrder_Select.ok = 0
        
        giIndex = -1
        ReDim gOrder_List(0)
        
        kbnu_Order_Request gSpecID, gHPEQUIP
        
        If gOrder_Select.ok = 1 Then
'            lRow = vasExam.DataRowCnt + 1
'            If lRow > vasExam.MaxRows Then
'                vasExam.MaxRows = lRow
'            End If
'
            vasExam.SetText 2, glRow, gSpecID
            vasExam.SetText 7, glRow, gOrder_Select.PT_NO
            vasExam.SetText 8, glRow, gOrder_Select.PT_NM
            If InStr(1, gOrder_Select.Sex, "/") > 0 Then
                vasExam.SetText 9, glRow, Left(gOrder_Select.Sex, InStr(1, gOrder_Select.Sex, "/") - 1)
                vasExam.SetText 10, glRow, Mid(gOrder_Select.Sex, InStr(1, gOrder_Select.Sex, "/") + 1)
            End If
            
            vasExam.Row = glRow
            vasExam.Col = 1
            vasExam.Value = 1
        Else
            vasExam.Row = glRow
            vasExam.Col = 1
            vasExam.Value = 0
        End If
        
'        mExam = Get_OrderBody(gSpecID)
'        If Not IsNull(mExam) Then
'            SetText vasList, Trim(mExam(1, LBound(mExam, 2))), glRow, 3
'            SetText vasList, Trim(mExam(2, LBound(mExam, 2))), glRow, 4
'
'            SetText vasList, Trim(mExam(5, LBound(mExam, 2))) & "-" & Trim(mExam(6, LBound(mExam, 2))), glRow, 5
'
'            For i = LBound(mExam, 2) To UBound(mExam, 2)
'                'SetText vasExam, mExam(3, i), lRow, 1
'                'SetText vasExam, mExam(4, i), lRow, 2
'
'                gReadBuf(0) = ""
'                gReadBuf(1) = ""
'
'                SQL = "Select ExamCode, EquipCode from EquipExam " & vbCrLf & _
'                      "where Equip = '" & gEquip & "' " & vbCrLf & _
'                      "  and ExamCode = '" & Trim(mExam(3, i)) & "' " & vbCrLf & _
'                      "  and UseFlag = 1 "
'                res = db_select_Col(gLocal, SQL)
'                If Trim(gReadBuf(0)) = Trim(mExam(3, i)) Then
'
'
'                    For j = 1 To UBound(gArrEquip)
'                        If Trim(gArrEquip(j, 1)) = Trim(gReadBuf(1)) Then
'                            SetText vasList, "*", glRow, gResCol + j
'
'                            SQL = "select result FROM pat_res " & vbCrLf & _
'                                  "WHERE examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'                                  "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'                                  "  AND equipcode = '" & Trim(Trim(gReadBuf(1))) & "'" & vbCrLf & _
'                                  "  AND barcode = '" & gSpecID & "' "
'                            res = db_select_Col(gLocal, SQL)
'                            If Trim(gReadBuf(0)) <> "" Then
'                                SetText vasList, Trim(gReadBuf(0)), glRow, gResCol + j
'                            End If
'
'                            Exit For
'                        End If
'                    Next j
'                End If
'            Next i
'
'        Else
'            vasList.Row = glRow
'            vasList.Col = 1
'            vasList.Value = 1
'            vasList.SetText 8, glRow, "오류"
'        End If
        
        '==========================================================================
                        
'        If Trim(GetText(vaslist, glRow, colPID)) = "" Then
'            PatInfo sSpecID, glRow
'        End If
        
        gPreSpecID = gSpecID
        gPreRow = glRow
        
    Case "R"    'Result
        gRecodeType = "R"

        'SetText vasList, "결과", glRow, 8
        'SetBackColor vasList, glRow, glRow, 1, 1, 255, 250, 205
        
        sExamCode = ""
        sResClassCode = ""
        sExamName = ""
        'lsResult = ""
        
        iCnt = 0
        i = InStr(1, argData, "|")
        Do While i > 0
            iCnt = iCnt + 1
            lsTmp = Left(argData, i - 1)
            Select Case iCnt
            Case 3
                lsTmp = Mid(lsTmp, 4)
                j = InStr(1, lsTmp, "^")
                lsTmp = Mid(lsTmp, j + 1)
                j = InStr(1, lsTmp, "^")
                lsEquip = Left(lsTmp, j - 1)
                lsEquipCode = Mid(lsTmp, j + 1)
                
'                sFlag = Right(lsTmp, 1)
                                
            Case 4
                lsResult = lsTmp
                
                '5R|2|^^^TID00^HBV^IPC_HBV Result|Invalid|||N||F||||||
                If lsEquipCode = "IPC_HBV Result" Then
                    lsIPCHBVRes = lsResult
                End If
                '6R|3|^^^TID00^HBV^HBV Ct|Undetermined|||N||F||||||
                If lsEquipCode = "HBV Ct" Then
                    lsIPCHBVCt = lsResult
                End If
                
                
                If IsNumeric(lsResult) = True Then
                    lsResult = Format(lsResult, "###########.####")
                
                    If Right(lsResult, 1) = "." Then
                        lsResult = Left(lsResult, Len(lsResult) - 1)
                    End If
                End If
                
                If lsIPCHBVCt = "Undetermined" Then
                    lsResult = "Undetermined"
                End If
                If lsIPCHBVRes = "Invalid" Then
                    lsResult = "Invalid"
                End If
'                                검사코드 불러오기
                ClearSpread vasTemp
                gReadBuf(0) = ""
                gReadBuf(1) = ""
                SQL = "SELECT EquipCode, ExamCode, ExamName " & CR & _
                      "  From EquipExam " & CR & _
                      " WHERE Equipno = '" & gEquip & "' " & CR & _
                      "   and EquipCode = '" & lsEquipCode & "' " & vbCrLf & _
                      " Order by seqno "
'                SQL = "SELECT EquipCode, ExamCode, ExamName " & CR & _
'                      "  From EquipExam " & CR & _
'                      " WHERE Equipno = '" & gEquip & "' " & CR & _
'                      "   and EquipCode = '" & lsEquip & "' " & vbCrLf & _
'                      "   and UnitCode = '" & lsEquipCode & "' " & vbCrLf & _
'                      " Order by seqno "
                res = db_select_Col(gLocal, SQL)
                
                If res > 0 Then
                    sExamCode = gReadBuf(1)
                Else
                    Exit Sub
                End If
                
                For m = 1 To UBound(gArrEquip)
                    If Trim(lsEquipCode) = gArrEquip(m, 2) Then


                        k = gArrEquip(m, 1)
                        lCol = (gArrEquip(m, 1) - 1)

                        SetText vasExam, lsResult, glRow, colResult + lCol * 4
                        SetText vasExam, lsResult, glRow, colResult1 + lCol

  
                        sExamDate = SeperatorCls(Text_Today.Text)
                        sBarcode = Trim(GetText(vasExam, glRow, colID))
                        sReceNo = Trim(GetText(vasExam, glRow, colReceNo))
                        sSeqNo = Trim(GetText(vasExam, glRow, colSeqNo))
                        sDiskNo = Trim(GetText(vasExam, glRow, colDiskNo))
                        sPosNo = Trim(GetText(vasExam, glRow, colPosNo))
                        sPname = Trim(GetText(vasExam, glRow, colPName))
                        sPSex = Trim(GetText(vasExam, glRow, colSex))
                        sPage = Trim(GetText(vasExam, glRow, colAge))
'                        sJumin1 = Trim(GetText(vasExam, glRow, colJumin1))
'                        sJumin2 = Trim(GetText(vasExam, glRow, colJumin2))
'                        sReceDate = Trim(GetText(vasExam, glRow, colReceDate))
                        sPID = sReceNo
'                        sBun = Trim(GetText(vasExam, glRow, colBun))
'                        sWorNo = Trim(GetText(vasExam, glRow, colPID))

                        'Local Table Insert
                        '환자 데이타 ====================================================================================
                        db_BeginTran gLocal

                        sCnt = ""
                        SQL = "Select count(*) from pat_res " & vbCrLf & _
                              "where examdate = '" & sExamDate & "' and barcode = '" & sBarcode & "' and equipcode = '" & lsEquipCode & "'  and examcode = '" & sExamCode & "' "
                        res = db_select_Var(gLocal, SQL, sCnt)
                        'res = db_select_Col(gLocal, SQL)
                        If res <= 0 Then
                            SaveQuery SQL
                            db_RollBack gLocal
                            Exit Sub
                        End If
                        If Not IsNumeric(sPage) Then
                            sPage = "0"
                        End If

                        
                        sDate = Format(Now, "yyyy/mm/dd hh:nn:ss")
                        
                        If CInt(sCnt) = 0 Then
                            '입력
                            SQL = "insert into pat_res (examdate, equipno, seqno, diskno, posno, " & vbCrLf & _
                                  " barcode, equipcode, examcode, result, refflag,  " & vbCrLf & _
                                  " refvalue,receno, recedate, pid, pname,  " & vbCrLf & _
                                  " psex, page, resflag,  " & vbCrLf & _
                                  " resdate, panicflag, deltaflag) " & vbCrLf & _
                                  "values ('" & sExamDate & "', '" & gEquip & "', '" & sSeqNo & "', '" & sDiskNo & "', '" & sPosNo & "', " & vbCrLf & _
                                  " '" & sBarcode & "', '" & lsEquipCode & "', '" & sExamCode & "', '" & lsResult & "', '" & sRefFlag & "', " & vbCrLf & _
                                  " '" & sRefLow & " - " & sRefHigh & "', '', '" & sExamDate & "', '" & sReceNo & "', '" & sPname & "', " & vbCrLf & _
                                  " '" & sPSex & "', " & sPage & ", 'A', " & vbCrLf & _
                                  " '" & sDate & "', '" & sPanicFlag & "', '" & sDeltaFlag & "' ) "
                            res = SendQuery(gLocal, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                                db_RollBack gLocal
                                'Exit Sub
                            End If
                        ElseIf CInt(sCnt) > 0 Then
                            '수정
                            SQL = "Update pat_res set " & vbCrLf & _
                                  " seqno =  '" & sSeqNo & "', " & vbCrLf & _
                                  " diskno =  '" & sDiskNo & "', " & vbCrLf & _
                                  " posno =  '" & sPosNo & "', " & vbCrLf & _
                                  " examcode =  '" & sExamCode & "', " & vbCrLf & _
                                  " result =  '" & lsResult & "' , " & vbCrLf & _
                                  " refflag =  '" & sRefFlag & "', " & vbCrLf & _
                                  " refvalue =  '" & sRefLow & " - " & sRefHigh & "', " & vbCrLf & _
                                  " receno =  '', " & vbCrLf & _
                                  " recedate =  '" & sExamDate & "', " & vbCrLf & _
                                  " pid =  '" & sReceNo & "', " & vbCrLf & _
                                  " pname =  '" & sPname & "', " & vbCrLf & _
                                  " psex =  '" & sPSex & "', " & vbCrLf & _
                                  " page =  " & sPage & ", " & vbCrLf & _
                                  " resflag =  'A', "

                            SQL = SQL & vbCrLf & _
                                  " resdate =  '" & sDate & "', " & vbCrLf & _
                                  " panicflag = '" & sPanicFlag & "', " & vbCrLf & _
                                  " deltaflag = '" & sDeltaFlag & "' " & vbCrLf & _
                                  "where examdate ='" & sExamDate & "' " & vbCrLf & _
                                  "  and barcode = '" & sBarcode & "' " & vbCrLf & _
                                  "  and equipcode = '" & lsEquipCode & "' " & vbCrLf & _
                                  "  and examcode = '" & sExamCode & "' "
                            res = SendQuery(gLocal, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                                db_RollBack gLocal
                                'Exit Sub
                            End If
                        End If

                        db_Commit gLocal

                        '서버에 바로 전송하기
                        SetText vasExam, "결과", glRow, colState

    '                    If optOption(0).Value = True Then
    '                        To_Server glRow
    '                    End If
                        '==============================================================================================
                        
                        If lsEquipCode = "HBV (copy/ml)" Then
                            
                            lsEquipCode = "HBV(pg/ml)"
                            
                            If IsNumeric(lsResult) = True Then
                                lsResult = Format(lsResult / 283000, "######0.0000")
                            End If
                            
                            For z = 1 To UBound(gArrEquip)
                                
                                
                                ClearSpread vasTemp
                                gReadBuf(0) = ""
                                gReadBuf(1) = ""
                                SQL = "SELECT EquipCode, ExamCode, ExamName " & CR & _
                                      "  From EquipExam " & CR & _
                                      " WHERE Equipno = '" & gEquip & "' " & CR & _
                                      "   and EquipCode = '" & lsEquipCode & "' " & vbCrLf & _
                                      " Order by seqno "
                                res = db_select_Col(gLocal, SQL)
                                
                                If res > 0 Then
                                    sExamCode = gReadBuf(1)
                                Else
                                    Exit Sub
                                End If
                                
                                If Trim(lsEquipCode) = gArrEquip(z, 2) Then
            
            
                                    k = gArrEquip(z, 1)
                                    lCol = (gArrEquip(z, 1) - 1)
            
                                    SetText vasExam, lsResult, glRow, colResult + lCol * 4
                                    SetText vasExam, lsResult, glRow, colResult1 + lCol
                                    
                                    If CInt(sCnt) = 0 Then
                                        '입력
                                        SQL = "insert into pat_res (examdate, equipno, seqno, diskno, posno, " & vbCrLf & _
                                              " barcode, equipcode, examcode, result, refflag,  " & vbCrLf & _
                                              " refvalue,receno, recedate, pid, pname,  " & vbCrLf & _
                                              " psex, page, resflag,  " & vbCrLf & _
                                              " resdate, panicflag, deltaflag) " & vbCrLf & _
                                              "values ('" & sExamDate & "', '" & gEquip & "', '" & sSeqNo & "', '" & sDiskNo & "', '" & sPosNo & "', " & vbCrLf & _
                                              " '" & sBarcode & "', '" & lsEquipCode & "', '" & sExamCode & "', '" & lsResult & "', '" & sRefFlag & "', " & vbCrLf & _
                                              " '" & sRefLow & " - " & sRefHigh & "', '', '" & sExamDate & "', '" & sReceNo & "', '" & sPname & "', " & vbCrLf & _
                                              " '" & sPSex & "', " & sPage & ", 'A', " & vbCrLf & _
                                              " '" & sDate & "', '" & sPanicFlag & "', '" & sDeltaFlag & "' ) "
                                        res = SendQuery(gLocal, SQL)
                                        If res = -1 Then
                                            SaveQuery SQL
                                            db_RollBack gLocal
                                            'Exit Sub
                                        End If
                                    ElseIf CInt(sCnt) > 0 Then
                                        '수정
                                        SQL = "Update pat_res set " & vbCrLf & _
                                              " seqno =  '" & sSeqNo & "', " & vbCrLf & _
                                              " diskno =  '" & sDiskNo & "', " & vbCrLf & _
                                              " posno =  '" & sPosNo & "', " & vbCrLf & _
                                              " examcode =  '" & sExamCode & "', " & vbCrLf & _
                                              " result =  '" & lsResult & "' , " & vbCrLf & _
                                              " refflag =  '" & sRefFlag & "', " & vbCrLf & _
                                              " refvalue =  '" & sRefLow & " - " & sRefHigh & "', " & vbCrLf & _
                                              " receno =  '', " & vbCrLf & _
                                              " recedate =  '" & sExamDate & "', " & vbCrLf & _
                                              " pid =  '" & sReceNo & "', " & vbCrLf & _
                                              " pname =  '" & sPname & "', " & vbCrLf & _
                                              " psex =  '" & sPSex & "', " & vbCrLf & _
                                              " page =  " & sPage & ", " & vbCrLf & _
                                              " resflag =  'A', "
            
                                        SQL = SQL & vbCrLf & _
                                              " resdate =  '" & sDate & "', " & vbCrLf & _
                                              " panicflag = '" & sPanicFlag & "', " & vbCrLf & _
                                              " deltaflag = '" & sDeltaFlag & "' " & vbCrLf & _
                                              "where examdate ='" & sExamDate & "' " & vbCrLf & _
                                              "  and barcode = '" & sBarcode & "' " & vbCrLf & _
                                              "  and equipcode = '" & lsEquipCode & "' " & vbCrLf & _
                                              "  and examcode = '" & sExamCode & "' "
                                        res = SendQuery(gLocal, SQL)
                                        If res = -1 Then
                                            SaveQuery SQL
                                            db_RollBack gLocal
                                            'Exit Sub
                                        End If
                                    End If
                                    Exit For
                                End If
                            Next z
                        End If

                        Exit For
                    End If
                Next m

            End Select
            
            
            argData = Mid(argData, i + 1)
            i = InStr(1, argData, "|")
        Loop

        
        gMsgFlag = ""
        'gHeadRecode = ""
        txtData.Text = ""
        
    Case "Q"    'Request
        If Text_Today.Text <> Format(Date, "yyyy-mm-dd") Then
            Text_Today.Text = Format(Date, "yyyy-mm-dd")
        End If
        
        gRecodeType = "Q"
        
        vasOrder.MaxRows = 5
        
        ClearSpread vasTemp
        ClearSpread vasOrder
        
        gHeader = """"
        gPatient = ""
        gOrder = ""
        gMsgEnd = ""
        'lsOrder = ""
        
        slen = InStr(1, argData, "|")
        argData = Mid(argData, slen + 1)
        
        slen = InStr(1, argData, "|")
        argData = Mid(argData, slen + 1)
        
        slen = InStr(1, argData, "|")
        gSpecID = Mid(argData, 1, slen - 1)     '검체번호
        slen = InStr(1, gSpecID, "^")
        gSpecID = Mid(gSpecID, slen + 1)   '검체번호
        
        glRow = vasList.DataRowCnt + 1
        If vasList.MaxRows < glRow + 1 Then
            vasList.MaxRows = glRow + 1
        End If
        
        glRow = -1
        For i = 1 To vasList.DataRowCnt
            If Trim(GetText(vasList, i, 2)) = gSpecID Then
                glRow = i
                
                'SetText vasList, vasRes.DataRowCnt, glRow, colRCnt
                
                Exit For
            End If
        Next i
        
        If glRow = -1 Then  ' vaslist에 없는 검체의 결과가 나올 때 데이터 추가
            glRow = vasList.DataRowCnt + 1
            If glRow > vasList.MaxRows Then
                vasList.MaxRows = glRow + 1
            End If
            vasActiveCell vasList, glRow, 2
            SetText vasList, gSpecID, glRow, 2
            
        End If
                        
        'If Trim(GetText(vaslist, glRow, colPID)) = "" Then
        '    PatInfo gSpecID, glRow
        'End If
        'lsOrder = ""
        
        ReDim lsOrder(0)
        z = 0
        
        gOrder_Select.ok = 0
        
        giIndex = -1
        ReDim gOrder_List(0)
        
        kbnu_Order_Request gSpecID, gHPEQUIP
    
        If gOrder_Select.ok = 1 Then
            lRow = vasExam.DataRowCnt + 1
            If lRow > vasExam.MaxRows Then
                vasExam.MaxRows = lRow
            End If
            
            vasExam.SetText 2, lRow, gSpecID
            vasExam.SetText 7, lRow, gOrder_Select.PT_NO
            vasExam.SetText 8, lRow, gOrder_Select.PT_NM
            If InStr(1, gOrder_Select.Sex, "/") > 0 Then
                vasExam.SetText 9, lRow, Left(gOrder_Select.Sex, InStr(1, gOrder_Select.Sex, "/") - 1)
                vasExam.SetText 10, lRow, Mid(gOrder_Select.Sex, InStr(1, gOrder_Select.Sex, "/") + 1)
            End If
        
            For i = 1 To UBound(gOrder_List)
                'SetText vasExam, mExam(3, i), lRow, 1
                'SetText vasExam, mExam(4, i), lRow, 2
                
                gReadBuf(0) = ""
                gReadBuf(1) = ""
                
                SQL = "Select ExamCode, UnitCode, EquipCode from EquipExam " & vbCrLf & _
                      "where Equip = '" & gEquip & "' " & vbCrLf & _
                      "  and ExamCode = '" & Trim(gOrder_List(i).TST_CD) & "' " & vbCrLf & _
                      "  and UseFlag = 1 "
                res = db_select_Col(gLocal, SQL)
                If Trim(gReadBuf(0)) = Trim(mExam(3, i)) Then
                    '^^^CKMB\^^^Myoglob\^^^Tpn-I
                    
                    'lsOrder = lsOrder & "^^^" & Trim(gReadBuf(1)) & "\"
                    k = -1
                    For j = LBound(lsOrder) To UBound(lsOrder)
                        If Trim(lsOrder(j)) = Trim(gReadBuf(1)) Then
                            k = 1
                            Exit For
                        End If
                    Next j
                    
                    If k = -1 Then
                        z = z + 1
                        ReDim Preserve lsOrder(z)
                        lsOrder(z) = Trim(gReadBuf(1))
                        

                    End If
                    
                    For j = 1 To UBound(gArrEquip)
                        If Trim(gArrEquip(j, 1)) = Trim(gReadBuf(2)) Then
                            SetText vasList, "*", glRow, gResCol + j
    
    '                        Save_Local_One glRow, i, "A"
                            Exit For
                        End If
                    Next j
                    
                End If
            Next i
'            If Len(lsOrder) > 0 Then
'                lsOrder = Left(lsOrder, Len(lsOrder) - 1)
'            End If
        End If
        
        'Order 만들기
        vasOrder.MaxRows = 20
        ClearSpread vasOrder
        
        i = 0
        'Head
        i = i + 1
        If i = 8 Then
            i = 0
        End If
        'gHeader = "H|\^&||||||||" & gVersion & "||P|1|" & chrCR & chrETX
        'gHeader = "H|\^&||||||||||P|1" & chrCR & chrETX
        gHeader = "H|\^&" & chrCR & chrETX
        gHeader = chrSTX & CCur(i) & gHeader & CheckSum(CStr(1) & gHeader) & chrCR & chrLF
        
        SetText vasOrder, gHeader, vasOrder.DataRowCnt + 1, 1
        
        'Patient
        i = i + 1
        If i = 8 Then
            i = 0
        End If
        
        sPName_E = UCase(Conv_Kor_Eng(Trim(GetText(vasList, glRow, 8))))
        '3P|2||PID0003||Hong Gil dong2^^^|29
        'gPatient = "P|1||||" & Trim(GetText(vasList, glRow, 3)) & "||||||||||||||" & chrCR & chrETX
        gPatient = "P|1||" & Trim(GetText(vasList, glRow, 7)) & "||" & sPName_E & "^^^|" & chrCR & chrETX
        gPatient = chrSTX & CCur(i) & gPatient & CheckSum(CStr(2) & gPatient) & chrCR & chrLF
        
        SetText vasOrder, gPatient, vasOrder.DataRowCnt + 1, 1
        
        'Order
        For j = 1 To UBound(lsOrder)
            i = i + 1
            If i = 8 Then
                i = 0
            End If
            'gOrder = "O|" & j & "|" & gSpecID & "||^^^" & lsOrder(j) & "|R||" & _
                                      "||||A||||Serum" & chrCR & chrETX
            'gOrder = "O|" & j & "|" & gSpecID & "||^^^" & lsOrder(j) & chrCR & chrETX
            '3O|1|SID0002^HBV^Hong gil dong1 blood sample^B^2||^^^TID00_HBV^HBV|||||||||||||||||||||F||||||06
'            If j = 1 Then
'                gOrder = "O|" & j & "|" & gSpecID & "||^^^" & lsOrder(j) & "|R||" & _
'                                          "||||N||||||||||||||Q|" & chrCR & chrETX
'            Else
'                gOrder = "O|" & j & "|" & gSpecID & "||^^^" & lsOrder(j) & "|R||" & _
'                                          "||||A||||||||||||||Q|" & chrCR & chrETX
'            End If
            If j = 1 Then
                gOrder = "O|" & j & "|" & gSpecID & "^" & lsOrder(j) & "^^" & Trim(txtCol.Text) & "^" & Trim(txtRow.Text) & _
                                          "|||||||||||||||||||||F||||||" & chrCR & chrETX
            Else
                gOrder = "O|" & j & "|" & gSpecID & "^" & lsOrder(j) & "^^" & Trim(txtCol.Text) & "^" & Trim(txtRow.Text) & _
                                          "|||||||||||||||||||||F||||||" & chrCR & chrETX
            End If
            gOrder = chrSTX & CCur(i) & gOrder & CheckSum(CStr(i) & gOrder) & chrCR & chrLF
            
            SetText vasOrder, gOrder, vasOrder.DataRowCnt + 1, 1
        Next j
        If UBound(lsOrder) < 1 Then
            i = i + 1
            If i = 8 Then
                i = 0
            End If
            'gOrder = "O|1|" & gSpecID & "||^^^ALL|R||" & _
                                      "||||N||||||||||||||Q|" & chrCR & chrETX
            gOrder = "O|" & j & "|" & gSpecID & "^ALL^^" & Trim(txtCol.Text) & "^" & Trim(txtRow.Text) & _
                                                      "|||||||||||||||||||||F||||||" & chrCR & chrETX
            gOrder = chrSTX & CCur(i) & gOrder & CheckSum(CStr(i) & gOrder) & chrCR & chrLF
            
            SetText vasOrder, gOrder, vasOrder.DataRowCnt + 1, 1
        
        End If
        
        'Msg End
        i = i + 1
        If i = 8 Then
            i = 0
        End If
        gMsgEnd = "L|1|N" & chrCR & chrETX
        gMsgEnd = Chr(2) & CCur(i) & gMsgEnd & CheckSum(CStr(i) & gMsgEnd) & chrCR & chrLF
    
        SetText vasOrder, gMsgEnd, vasOrder.DataRowCnt + 1, 1
        
        i = i + 1
        If i = 8 Then
            i = 0
        End If
        SetText vasOrder, chrEOT, vasOrder.DataRowCnt + 1, 1
        
        
        'Make_Order gSpecID, glRow
        
    Case "L"    '자료수신 완료
    
        If optOption(0).Value = True Then
            If glRow > 0 And glRow <= vasExam.DataRowCnt Then
                To_Server_1 glRow
            End If
        End If
        
    End Select


End Sub

Sub BS400(argData As String)
    Dim i       As Integer
    Dim j       As Integer
    Dim k       As Integer
    Dim iPos    As Integer
    
    Dim iCnt    As Integer
    
    Dim sTmp    As String
    Dim sMSH    As String
    Dim sQRD    As String
    Dim sQRF    As String
    Dim sMsgID  As String
    Dim sACK
    
    Dim sMsgType As String
    
    Dim lRow    As Long
    
    Dim iOCnt   As Integer
    
    Dim iRow    As Integer
    Dim ii      As Integer
    Dim jj      As Integer
    
    Dim iStr    As String
    
    
    If argData = "" Then
        Exit Sub
    End If

    i = InStr(1, argData, "QRY^Q02")        'Request************************
    If i > 0 Then
        j = InStr(1, argData, chrCR)
        If j > 0 Then
            'QCK============================================================
            sTmp = ""
            sTmp = Mid(argData, 1, j - 1)

            sMSH = ""
            iCnt = 0

            k = InStr(1, sTmp, "|")
            Do While k > 0
                iCnt = iCnt + 1
                Select Case iCnt
                Case 3
                    sMSH = sMSH & "|"
                Case 4
                    sMSH = sMSH & "|"
                Case 5
                    sMSH = sMSH & "Mindray|"
                Case 6
                    sMSH = sMSH & "BS-200|"
                Case 7
                    sMSH = sMSH & Format(GetDateFull, "YYYYMMDDHHMMSS") & "|"
                Case 9
                    sMSH = sMSH & "QCK^Q02|"
                Case 10
                    sMsgID = "1|"
                    sMSH = sMSH & "1|"
                Case Else
                    sMSH = sMSH & Mid(sTmp, 1, k)
                End Select

                sTmp = Mid(sTmp, k + 1)
                k = InStr(1, sTmp, "|")
            Loop

            sACK = ""
            sACK = sMSH & chrCR
            sACK = sACK & "MSA|AA|" & sMsgID & "|||0|" & chrCR
            sACK = sACK & "ERR|0|" & chrCR & "QAK|SR|OK|" & chrCR & chrFS & chrCR

            Winsock1.SendData sACK
            Save_Raw_Data "[TX:" & Format(Time, "hh:nn:ss") & "]" & sACK
            
            'Order==========================================================
            iOCnt = 0
            
            For lRow = 1 To vasExam.DataRowCnt
                vasExam.Row = lRow
                vasExam.Col = 1
                If vasExam.Value = 1 Then
                    sACK = ""
                    sTmp = ""
                    
                    iPos = InStr(1, argData, chrCR)
                    sTmp = Mid(argData, 1, iPos - 1)
            
                    sMSH = ""
                    iCnt = 0
                    
                    iOCnt = iOCnt + 1
                    
                    k = InStr(1, sTmp, "|")
                    Do While k > 0
                        iCnt = iCnt + 1
                        Select Case iCnt
                        Case 3
                            sMSH = sMSH & "|"
                        Case 4
                            sMSH = sMSH & "|"
                        Case 5
                            sMSH = sMSH & "Mindray|"
                        Case 6
                            sMSH = sMSH & "BS-200|"
                        Case 7
                            sMSH = sMSH & Format(GetDateFull, "YYYYMMDDHHMMSS") & "|"
                        Case 9
                            sMSH = sMSH & "DSR^Q03|"
                        Case 10
                            'sMsgID = "1|"
                            sMsgID = iOCnt & "|"
                            sMSH = sMSH & sMsgID
                        Case Else
                            sMSH = sMSH & Mid(sTmp, 1, k)
                        End Select
        
                        sTmp = Mid(sTmp, k + 1)
                        k = InStr(1, sTmp, "|")
                    Loop
        
                    sACK = ""
                    sACK = sMSH & chrCR
                    sACK = sACK & "MSA|AA|" & sMsgID & "|||0|" & chrCR
                    sACK = sACK & "ERR|0|" & chrCR & "QAK|SR|OK|" & chrCR
        
                    iPos = InStr(1, argData, "QRD")
                    If iPos > 0 Then
                        j = InStr(iPos, argData, chrCR)
                        If j > 0 Then
                            sTmp = ""
                            sTmp = Mid(argData, iPos, j - iPos)
        
                            sQRD = ""
                            iCnt = 0
        
                            k = InStr(1, sTmp, "|")
                            Do While k > 0
                                iCnt = iCnt + 1
                                Select Case iCnt
                                Case 2
                                    sQRD = sQRD & Format(GetDateFull, "YYYYMMDDHHMMSS") & "|"
                                Case 5
                                    'sMsgID = "1|"
                                    
                                    sMsgID = iOCnt & "|"
                                    sQRD = sQRD & sMsgID
                                Case Else
                                    sQRD = sQRD & Mid(sTmp, 1, k)
                                End Select
        
                                sTmp = Mid(sTmp, k + 1)
                                k = InStr(1, sTmp, "|")
                            Loop
        
                            sACK = sACK & sQRD & chrCR
                        End If
                    End If
        
                    iPos = InStr(1, argData, "QRF")
                    If iPos > 0 Then
                        j = InStr(iPos, argData, chrCR)
                        If j > 0 Then
                            sTmp = ""
                            sTmp = Mid(argData, iPos, j - iPos)
        
                            sQRF = ""
                            iCnt = 0
        
                            k = InStr(1, sTmp, "|")
                            Do While k > 0
                                iCnt = iCnt + 1
                                Select Case iCnt
                                Case 2
                                    sQRF = sQRF & "BS-200|"
        '                        Case 3
        '                            sQRF = sQRF & Format(GetDateFull, "YYYYMMDDHHMMSS") & "|"
        '                        Case 4
        '                            sQRF = sQRF & Format(GetDateFull, "YYYYMMDDHHMMSS") & "|"
                                Case Else
                                    sQRF = sQRF & Mid(sTmp, 1, k)
                                End Select
        
                                sTmp = Mid(sTmp, k + 1)
                                k = InStr(1, sTmp, "|")
                            Loop
        
                            sACK = sACK & sQRF & chrCR
        
                        End If
                    End If
        
                    sACK = sACK & "DSP|1|||||" & chrCR
                    sACK = sACK & "DSP|2|||||" & chrCR
                    sACK = sACK & "DSP|3||" & Trim(GetText(vasExam, lRow, colPName)) & "|||" & chrCR
                    sACK = sACK & "DSP|4|||||" & chrCR
                    sACK = sACK & "DSP|5||" & Trim(GetText(vasExam, lRow, colSex)) & "|||" & chrCR
                    sACK = sACK & "DSP|6|||||" & chrCR
                    sACK = sACK & "DSP|7|||||" & chrCR
                    sACK = sACK & "DSP|8|||||" & chrCR
                    sACK = sACK & "DSP|9|||||" & chrCR
                    sACK = sACK & "DSP|10|||||" & chrCR
                    sACK = sACK & "DSP|11|||||" & chrCR
                    sACK = sACK & "DSP|12|||||" & chrCR
                    sACK = sACK & "DSP|13|||||" & chrCR
                    sACK = sACK & "DSP|14|||||" & chrCR
                    sACK = sACK & "DSP|15||outpatient|||" & chrCR
                    sACK = sACK & "DSP|16|||||" & chrCR
                    sACK = sACK & "DSP|17|||||" & chrCR
                    sACK = sACK & "DSP|18|||||" & chrCR
                    sACK = sACK & "DSP|19|||||" & chrCR
                    sACK = sACK & "DSP|20|||||" & chrCR
                    sACK = sACK & "DSP|21||" & Trim(GetText(vasExam, lRow, colID)) & "|||" & chrCR
                    sACK = sACK & "DSP|22||" & Trim(GetText(vasExam, lRow, colSeqNo)) & "|||" & chrCR
                    sACK = sACK & "DSP|23|||||" & chrCR
                    sACK = sACK & "DSP|24||N|||" & chrCR        '응급여부
                    sACK = sACK & "DSP|25|||||" & chrCR
                    sACK = sACK & "DSP|26||Serum|||" & chrCR    '검체정보
                    sACK = sACK & "DSP|27|||||" & chrCR
                    sACK = sACK & "DSP|28|||||" & chrCR
                    
'                    sACK = sACK & "DSP|29||TP^^^|||" & chrCR
'        '            sACK = sACK & "DSP|30|ALB^ALB^^|||" & chrCR
                    
                
                    ClearSpread vasOrder
                        
                    SQL = "SELECT SUBCODE " & vbCrLf
                    SQL = SQL & "From TWEXAM_RESULTC  " & vbCrLf
                    SQL = SQL & "WHERE SPECNO = '" & Trim(GetText(vasExam, lRow, 2)) & "' " & vbCrLf
                    SQL = SQL & "  AND SUBCODE In (" & gAllExam & ") " & vbCrLf
                    SQL = SQL & "  AND STATUS in ('2','3') "
                    
                    res = db_select_Vas(gServer, SQL, vasOrder)
                    
                    If res = -1 Then
                        SaveQuery "[-1]" & SQL
                    ElseIf res = 0 Then
                        SaveQuery "[0]" & SQL
                    End If
                        
                    iStr = 29
                    If vasOrder.DataRowCnt > 0 Then
                        For ii = 1 To vasOrder.DataRowCnt
                            For jj = 1 To UBound(gArrEquip)
                                If Trim(gArrEquip(jj, 3)) = Trim(GetText(vasOrder, ii, 1)) Then
                                    SetText vasExam, "*", lRow, colResult + (gArrEquip(jj, 1) - 1) * 4
                                
                                    SetText vasOrdBuff, gArrEquip(jj, 2), i, 2
                                    
                                    sACK = sACK & "DSP|" & iStr & "||" & Trim(gArrEquip(jj, 2)) & "^^^|||" & chrCR
                                    iStr = iStr + 1
                                End If
                            Next jj
                        Next ii
                    End If
                        
                   
                    
                    If iOCnt = txtChkCnt Then
                        sACK = sACK & "DSC||" & chrCR & chrFS & chrCR
                    Else
                        sACK = sACK & "DSC|" & iOCnt & "|" & chrCR & chrFS & chrCR
                    End If
                    
                    Winsock1.SendData sACK
                    Save_Raw_Data "[TX:" & Format(Time, "hh:nn:ss") & "]" & sACK
                End If
            Next lRow
        End If
      
    End If
    '***********************************************************************
    
    i = InStr(1, argData, "ORU^R01")        'Result*************************
    If i > 0 Then
        j = InStr(1, argData, chrCR)
        If j > 0 Then
            sTmp = ""
            sTmp = Mid(argData, 1, j - 1)
            
            sMSH = ""
            iCnt = 0
            
            k = InStr(1, sTmp, "|")
            Do While k > 0
                iCnt = iCnt + 1
                Select Case iCnt
                Case 3
                    sMSH = sMSH & "|"
                Case 4
                    sMSH = sMSH & "|"
                Case 5
                    sMSH = sMSH & "Mindray|"
                Case 6
                    sMSH = sMSH & "BS-200|"
                Case 7
                    sMSH = sMSH & Format(GetDateFull, "YYYYMMDDHHMMSS") & "|"
                Case 9
                    sMSH = sMSH & "ACK^R01|"
                Case 10
                    sMsgID = Mid(sTmp, 1, k)
                    iPos = InStr(1, sMsgID, "|")
                    If iPos > 0 Then
                        sMsgID = sMsgID
                    End If
'
                    sMSH = sMSH & sMsgID
                    
                    'sMsgID = "1|"
                    'sMSH = sMSH & "1|"
                Case Else
                    sMSH = sMSH & Mid(sTmp, 1, k)
                End Select
                
                sTmp = Mid(sTmp, k + 1)
                k = InStr(1, sTmp, "|")
            Loop
                                                     
            sACK = ""
            sACK = sMSH & chrCR
            sACK = sACK & "MSA|AA|" & sMsgID & "|||0|" & chrCR & chrFS & chrCR
            
            Winsock1.SendData sACK
            Save_Raw_Data "[TX:" & Format(Time, "hh:nn:ss") & "]" & sACK
        End If
        
        '결과처리
        Proc_Result argData
    End If
    
End Sub

Sub Proc_Result(argData As String)
    Dim i       As Integer
    Dim j       As Integer
    Dim k       As Integer
    Dim m       As Integer
    Dim z       As Integer
    
    Dim iPos    As Integer
    Dim iCnt    As Integer
    Dim iRow    As Integer
    Dim lCol    As Long
    
    Dim sTmp        As String
   

    Dim sDate As String
    Dim sTmpStr As String
    Dim sPoint As String
    Dim sRefFlag As String
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sPanicFlag As String
    Dim sDeltaFlag As String
    Dim sReceNo As String
    Dim sReceDate  As String
    Dim sPID As String
    Dim sPname As String
    Dim sJumin1 As String
    Dim sJumin2 As String
    Dim sPSex As String
    Dim sPage As String
    Dim sBun As String
    Dim sTestID As String
    Dim sFlag As String
    Dim sExamCode As String
    Dim sExamCode1 As String
    Dim sResult As String
    Dim sExamDate As String
    Dim sBarcode    As String
    Dim sWorNo As String
    
    Dim sCnt As String
    
    If argData = "" Then
        Exit Sub
    End If
    
    'OBR|12|80487|1|Mindray^BS-380|N||20100211101116|||||||20100211095454|Serum|||||||||||||||||||||||||||||||||
    i = InStr(1, argData, "OBR")
    If i > 0 Then
        iPos = InStr(i, argData, chrCR)

        sTmp = ""
        sTmp = Mid(argData, i, iPos - i)

        iCnt = 0

        j = InStr(1, sTmp, "|")
        Do While j > 0
            iCnt = iCnt + 1

            Select Case iCnt
            Case 3
                sBarcode = Mid(sTmp, 1, j - 1)
                
                glRow = -1
                For iRow = 1 To vasExam.DataRowCnt
                    If sBarcode <> "" Then  '메뉴얼일경우 바코드 없음
                        If Trim(GetText(vasExam, iRow, colID)) = sBarcode Then
                            glRow = iRow
                            Exit For
                        End If
                    End If
                Next iRow

                If glRow = -1 Then
                    glRow = vasExam.DataRowCnt + 1
                    If vasExam.MaxRows < glRow Then
                        vasExam.MaxRows = glRow
                    End If
                End If

                SetText vasExam, sBarcode, glRow, colID

                '환자정보 불러오기
                If sBarcode <> "" And Trim(GetText(vasExam, glRow, colPName)) = "" Then
                    Get_Sample_Info glRow
                End If
            End Select

            sTmp = Mid(sTmp, j + 1)
            j = InStr(1, sTmp, "|")
        Loop
    End If

    'OBX|1|NM|BUN|BUN|14.725409|mg/dL|-|N|||F||14.725409|20100211101116||||
    i = InStr(1, argData, "OBX")
    Do While i > 0
        iPos = InStr(i, argData, chrCR)

        sTmp = ""
        sTmp = Mid(argData, i, iPos - i)

        iCnt = 0

        j = InStr(1, sTmp, "|")
        Do While j > 0
            iCnt = iCnt + 1

            Select Case iCnt
            Case 4
                sTestID = Mid(sTmp, 1, j - 1)

            Case 6
                sResult = Mid(sTmp, 1, j - 1)
                
                '검사코드 불러오기
                ClearSpread vasTemp
                gReadBuf(0) = ""
                gReadBuf(1) = ""
                SQL = "SELECT EquipCode, ExamCode, ExamName " & CR & _
                      "  From EquipExam " & CR & _
                      " WHERE Equipno = '" & gEquip & "' " & CR & _
                      "   and EquipCode = '" & sTestID & "' " & vbCrLf & _
                      " Order by seqno "
                res = db_select_Col(gLocal, SQL)
                
                sExamCode = gReadBuf(1)
                
                For m = 1 To UBound(gArrEquip)
                    If Trim(sTestID) = gArrEquip(m, 2) Then
                        If IsNumeric(sResult) Then
                             '소수점 처리
                             sPoint = gArrEquip(m, 5)
                             
                             If IsNumeric(sPoint) Then
                                 If CInt(sPoint) > 0 Then
                                     sTmpStr = "#0."
                                     For z = 1 To CInt(sPoint)
                                         sTmpStr = sTmpStr & "0"
                                     Next z
                                 Else
                                     sTmpStr = "#0"
                                 End If
                                 sResult = Format(sResult, sTmpStr)
                             End If
                         End If
                
                        k = gArrEquip(m, 1)
                        lCol = (gArrEquip(m, 1) - 1)
                        
                        SetText vasExam, sResult, glRow, colResult + lCol * 4
                        SetText vasExam, sResult, glRow, colResult1 + lCol
                        
                        '참고치 체크================================================================
                        sRefLow = gArrEquip(k, 6)
                        sRefHigh = gArrEquip(k, 7)
                        If Not IsNumeric(sRefLow) Then
                            sRefLow = "0"
                        End If
                        If Not IsNumeric(sRefHigh) Then
                            sRefHigh = "0"
                        End If
                        If CCur(sRefLow) = 0 And CCur(sRefHigh) = 0 Then
                            sRefFlag = ""
                        Else
                            If CCur(sResult) < CCur(sRefLow) Then
                                sRefFlag = "Neg"    'Low
                            End If
                            If CCur(sResult) > CCur(sRefHigh) Then
                                sRefFlag = "Pos"    'High
                            End If
                        End If
                        SetText vasExam, sRefFlag, glRow, colResult + lCol * 4 + 1
            
                        'Panic 체크================================================================
                        If Not IsNumeric(gArrEquip(k, 8)) Then
                            gArrEquip(k, 8) = "0"
                        End If
                        If Not IsNumeric(gArrEquip(k, 9)) Then
                            gArrEquip(k, 9) = "0"
                        End If
                        If CCur(gArrEquip(k, 8)) = 0 And CCur(gArrEquip(k, 9)) = 0 Then
                            sPanicFlag = "X"
                        Else
                            If CCur(sResult) < CCur(gArrEquip(k, 8)) Then
                                sPanicFlag = "Neg"
                            End If
                            If CCur(sResult) > CCur(gArrEquip(k, 9)) Then
                                sPanicFlag = "Pos"
                            End If
                        End If
                        
                        If sPanicFlag = "" Then
                            sPanicFlag = "X"
                        End If
                        
                        SetText vasExam, sPanicFlag, glRow, colResult + lCol * 4 + 2
                        'Delta 체크================================================================
'                        sDeltaFlag = Delta_Check(glRow, sResult, k)
'                        If sDeltaFlag = "" Then
'                            sDeltaFlag = "X"
'                        End If
                        
                        SetText vasExam, sDeltaFlag, glRow, colResult + lCol * 4 + 3
                        
                        '결과 셀 색깔 변화=========================================================
                        Select Case sRefFlag
                        Case "Pos"   'Positive, 'High
                            SetBackColor vasExam, glRow, glRow, colResult + lCol * 4, colResult + lCol * 4 + 1, 246, 150, 121
                        Case "Neg"  'Negative, Low
                            SetBackColor vasExam, glRow, glRow, colResult + lCol * 4, colResult + lCol * 4 + 1, 255, 245, 104
                        Case Else   'Normal
                        End Select
                        Select Case sPanicFlag
                        Case "Pos"   'High
                            SetBackColor vasExam, glRow, glRow, colResult + lCol * 4, colResult + lCol * 4 + 2, 242, 108, 79
                        Case "Neg"
                            SetBackColor vasExam, glRow, glRow, colResult + lCol * 4, colResult + lCol * 4 + 2, 60, 184, 120
                        Case Else   'Normal
                        End Select
                        If sDeltaFlag = "D" Then
                            SetBackColor vasExam, glRow, glRow, colResult + lCol * 4, colResult + lCol * 4 + 3, 255, 0, 0
                        End If
                        '변수에 담기 ====================================================================================
                        sDate = GetDateFull
                        sExamDate = SeperatorCls(Text_Today.Text)
                        sBarcode = Trim(GetText(vasExam, glRow, colID))
                        sReceNo = Trim(GetText(vasExam, glRow, colReceNo))
                        sSeqNo = Trim(GetText(vasExam, glRow, colSeqNo))
                        sDiskNo = Trim(GetText(vasExam, glRow, colDiskNo))
                        sPosNo = Trim(GetText(vasExam, glRow, colPosNo))
                        sPname = Trim(GetText(vasExam, glRow, colPName))
                        sPSex = Trim(GetText(vasExam, glRow, colSex))
                        sPage = Trim(GetText(vasExam, glRow, colAge))
                        sJumin1 = Trim(GetText(vasExam, glRow, colJumin1))
                        sJumin2 = Trim(GetText(vasExam, glRow, colJumin2))
                        sReceDate = Trim(GetText(vasExam, glRow, colReceDate))
                        sPID = sReceNo
                        sBun = Trim(GetText(vasExam, glRow, colBun))
                        sWorNo = Trim(GetText(vasExam, glRow, colPID))
                
                        'Local Table Insert
                        '환자 데이타 ====================================================================================
                        db_BeginTran gLocal
                        
                        sCnt = ""
                        SQL = "Select count(*) from pat_res " & vbCrLf & _
                              "where barcode = '" & sBarcode & "' and equipcode = '" & sTestID & "'  and examcode = '" & sExamCode & "' "
                        res = db_select_Var(gLocal, SQL, sCnt)
                        If res <= 0 Then
                            SaveQuery SQL
                            db_RollBack gLocal
                            Exit Sub
                        End If
                        If Not IsNumeric(sPage) Then
                            sPage = "0"
                        End If
                        
                        
                        If CInt(sCnt) = 0 Then
                            '입력
                            SQL = "insert into pat_res (examdate, equipno, seqno, diskno, posno, " & vbCrLf & _
                                  " barcode, equipcode, examcode, result, refflag,  " & vbCrLf & _
                                  " refvalue,receno, recedate, pid, pname,  " & vbCrLf & _
                                  " jumin1, jumin2,psex, page, resflag,  " & vbCrLf & _
                                  " resdate, panicflag, deltaflag, sampletype, examgubun ) " & vbCrLf & _
                                  "values ('" & sExamDate & "', '" & gEquip & "', '" & sSeqNo & "', '" & sDiskNo & "', '" & sPosNo & "', " & vbCrLf & _
                                  " '" & sBarcode & "', '" & sTestID & "', '" & sExamCode & "', '" & sResult & "', '" & sRefFlag & "', " & vbCrLf & _
                                  " '" & sRefLow & " - " & sRefHigh & "', '', '" & sReceDate & "', '" & sReceNo & "', '" & sPname & "', " & vbCrLf & _
                                  " '" & sJumin1 & "', '" & sJumin2 & "', '" & sPSex & "', " & sPage & ", 'A', " & vbCrLf & _
                                  " '" & sDate & "', '" & sPanicFlag & "', '" & sDeltaFlag & "', '" & sWorNo & "', '" & sBun & "' ) "
                            res = SendQuery(gLocal, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                                db_RollBack gLocal
                                'Exit Sub
                            End If
                        ElseIf CInt(sCnt) > 0 Then
                            '수정
                            SQL = "Update pat_res set " & vbCrLf & _
                                  " seqno =  '" & sSeqNo & "', " & vbCrLf & _
                                  " diskno =  '" & sDiskNo & "', " & vbCrLf & _
                                  " posno =  '" & sPosNo & "', " & vbCrLf & _
                                  " examcode =  '" & sExamCode & "', " & vbCrLf & _
                                  " result =  '" & sResult & "' , " & vbCrLf & _
                                  " refflag =  '" & sRefFlag & "', " & vbCrLf & _
                                  " refvalue =  '" & sRefLow & " - " & sRefHigh & "', " & vbCrLf & _
                                  " receno =  '', " & vbCrLf & _
                                  " recedate =  '" & sReceDate & "', " & vbCrLf & _
                                  " pid =  '" & sReceNo & "', " & vbCrLf & _
                                  " pname =  '" & sPname & "', " & vbCrLf & _
                                  " jumin1 =  '" & sJumin1 & "', " & vbCrLf & _
                                  " jumin2 =  '" & sJumin2 & "', " & vbCrLf & _
                                  " psex =  '" & sPSex & "', " & vbCrLf & _
                                  " page =  " & sPage & ", " & vbCrLf & _
                                  " resflag =  'A', "
                                  
                            SQL = SQL & vbCrLf & _
                                  " resdate =  '" & sDate & "', " & vbCrLf & _
                                  " panicflag = '" & sPanicFlag & "', " & vbCrLf & _
                                  " deltaflag = '" & sDeltaFlag & "', " & vbCrLf & _
                                  " sampletype = '" & sWorNo & "', " & vbCrLf & _
                                  " examgubun = '" & sBun & "' " & vbCrLf & _
                                  "where examdate ='" & sExamDate & "' " & vbCrLf & _
                                  "  and barcode = '" & sBarcode & "' " & vbCrLf & _
                                  "  and equipcode = '" & sTestID & "' " & vbCrLf & _
                                  "  and examcode = '" & sExamCode & "' "
                            res = SendQuery(gLocal, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                                db_RollBack gLocal
                                'Exit Sub
                            End If
                        End If
                        
                        db_Commit gLocal
                        
                        '서버에 바로 전송하기
                        SetText vasExam, "결과", glRow, colState
                        
    '                    If optOption(0).Value = True Then
    '                        To_Server glRow
    '                    End If
                        '==============================================================================================
                        
                        Exit For
                    End If
                Next m
        
                If k = 0 Then
                    Exit Sub
                End If
        
            End Select

            sTmp = Mid(sTmp, j + 1)
            j = InStr(1, sTmp, "|")
        Loop
        
        argData = Mid(argData, i + 1)
        i = InStr(1, argData, "OBX")
    Loop
    
    If optOption(0).Value = True Then
        To_Server glRow
    End If

End Sub

Function Save_Local_One(ByVal asRow As Long, ByVal aiIndex As Integer, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = GetDateFull
    
    sCnt = ""
    SQL = "Delete FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND equipcode = '" & gArrExamRes(aiIndex).EquipCode & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasList, asRow, 2)) & "' "
    
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = "INSERT INTO pat_res (examdate, equipno, " & _
            "barcode, examtype, receno, " & _
            "pid, pname, pjumin, page, psex, " & _
            "resdate, seqno, diskno, posno, " & _
            "equipcode, examcode, " & _
            "result, sendflag, examname) " & vbCrLf & _
          "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
          "'" & Trim(GetText(vasList, asRow, 2)) & "','" & Trim(GetText(vasList, asRow, 5)) & "', '', " & _
          "'" & Trim(GetText(vasList, asRow, 3)) & "', '" & Trim(GetText(vasList, asRow, 4)) & "', '', 0, '', " & _
          "'" & sExamDate & "', '" & gArrExamRes(aiIndex).SeqNo & "', '" & Trim(GetText(vasList, asRow, 6)) & "', '" & Trim(GetText(vasList, asRow, 7)) & "', " & vbCrLf & _
          "'" & gArrExamRes(aiIndex).EquipCode & "', '" & gArrExamRes(aiIndex).ExamCode & "', " & _
          "'" & gArrExamRes(aiIndex).res & "', '" & asSend & "', '" & gArrExamRes(aiIndex).ExamName & "', " & vbCrLf & _
          "'" & gArrExamRes(aiIndex).RefFlag & "', '', '', '', " & _
          "'" & gArrExamRes(aiIndex).RefLow & " ~ " & gArrExamRes(aiIndex).RefHigh & "', '' ) "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Sub BS380_TEST(argData As String)
    Dim i       As Integer
    Dim j       As Integer
    Dim k       As Integer
    Dim iPos    As Integer
    
    Dim iCnt    As Integer
    
    Dim sTmp    As String
    Dim sMSH    As String
    Dim sQRD    As String
    Dim sQRF    As String
    Dim sMsgID  As String
    Dim sACK
    
    Dim sMsgType As String
    
    Dim lRow    As Long
    
    If argData = "" Then
        Exit Sub
    End If
    
    i = InStr(1, argData, "QRY^Q02")        'Request************************
    If i > 0 Then
        j = InStr(1, argData, chrCR)
        If j > 0 Then
            'QCK============================================================
            sTmp = ""
            sTmp = Mid(argData, 1, j - 1)

            sMSH = ""
            iCnt = 0

            k = InStr(1, sTmp, "|")
            Do While k > 0
                iCnt = iCnt + 1
                Select Case iCnt
                Case 3
                    sMSH = sMSH & "|"
                Case 4
                    sMSH = sMSH & "|"
                Case 5
                    sMSH = sMSH & "Mindray|"
                Case 6
                    sMSH = sMSH & "BS-380|"
                Case 7
                    sMSH = sMSH & Format(GetDateFull, "YYYYMMDDHHMMSS") & "|"
                Case 9
                    sMSH = sMSH & "QCK^Q02|"
                Case 10
                    sMsgID = "1|"
                    sMSH = sMSH & "1|"
                Case Else
                    sMSH = sMSH & Mid(sTmp, 1, k)
                End Select

                sTmp = Mid(sTmp, k + 1)
                k = InStr(1, sTmp, "|")
            Loop

            sACK = ""
            sACK = sMSH & chrCR
            sACK = sACK & "MSA|AA|" & sMsgID & "|||0|" & chrCR
            sACK = sACK & "ERR|0|" & chrCR & "QAK|SR|OK|" & chrCR & chrFS & chrCR

            Winsock1.SendData sACK
            Save_Raw_Data "[TX:" & Format(Time, "hh:nn:ss") & "]" & sACK
            
            'DoSleep 500
            
            'Order==========================================================
            sACK = ""
            sTmp = ""
            sTmp = Mid(argData, 1, j - 1)

            sMSH = ""
            iCnt = 0

            k = InStr(1, sTmp, "|")
            Do While k > 0
                iCnt = iCnt + 1
                Select Case iCnt
                Case 3
                    sMSH = sMSH & "|"
                Case 4
                    sMSH = sMSH & "|"
                Case 5
                    sMSH = sMSH & "Mindray|"
                Case 6
                    sMSH = sMSH & "BS-380|"
                Case 7
                    sMSH = sMSH & Format(GetDateFull, "YYYYMMDDHHMMSS") & "|"
                Case 9
                    sMSH = sMSH & "DSR^Q03|"
                Case 10
                    sMsgID = "1|"
                    sMSH = sMSH & "1|"
                Case Else
                    sMSH = sMSH & Mid(sTmp, 1, k)
                End Select

                sTmp = Mid(sTmp, k + 1)
                k = InStr(1, sTmp, "|")
            Loop

            sACK = ""
            sACK = sMSH & chrCR
            sACK = sACK & "MSA|AA|" & sMsgID & "|||0|" & chrCR
            sACK = sACK & "ERR|0|" & chrCR & "QAK|SR|OK|" & chrCR

            iPos = InStr(1, argData, "QRD")
            If iPos > 0 Then
                j = InStr(iPos, argData, chrCR)
                If j > 0 Then
                    sTmp = ""
                    sTmp = Mid(argData, iPos, j - iPos)

                    sQRD = ""
                    iCnt = 0

                    k = InStr(1, sTmp, "|")
                    Do While k > 0
                        iCnt = iCnt + 1
                        Select Case iCnt
                        Case 2
                            sQRD = sQRD & Format(GetDateFull, "YYYYMMDDHHMMSS") & "|"
                        Case 5
                            sMsgID = "1|"
                            sQRD = sQRD & "1|"
                        Case Else
                            sQRD = sQRD & Mid(sTmp, 1, k)
                        End Select

                        sTmp = Mid(sTmp, k + 1)
                        k = InStr(1, sTmp, "|")
                    Loop

                    sACK = sACK & sQRD & chrCR
                End If
            End If

            iPos = InStr(1, argData, "QRF")
            If iPos > 0 Then
                j = InStr(iPos, argData, chrCR)
                If j > 0 Then
                    sTmp = ""
                    sTmp = Mid(argData, iPos, j - iPos)

                    sQRF = ""
                    iCnt = 0

                    k = InStr(1, sTmp, "|")
                    Do While k > 0
                        iCnt = iCnt + 1
                        Select Case iCnt
                        Case 2
                            sQRF = sQRF & "BS-380|"
'                        Case 3
'                            sQRF = sQRF & Format(GetDateFull, "YYYYMMDDHHMMSS") & "|"
'                        Case 4
'                            sQRF = sQRF & Format(GetDateFull, "YYYYMMDDHHMMSS") & "|"
                        Case Else
                            sQRF = sQRF & Mid(sTmp, 1, k)
                        End Select

                        sTmp = Mid(sTmp, k + 1)
                        k = InStr(1, sTmp, "|")
                    Loop

                    sACK = sACK & sQRF & chrCR

                End If
            End If

            sACK = sACK & "DSP|1|||||" & chrCR
            sACK = sACK & "DSP|2|||||" & chrCR
            sACK = sACK & "DSP|3||이상은|||" & chrCR
            sACK = sACK & "DSP|4|||||" & chrCR
            sACK = sACK & "DSP|5||F|||" & chrCR
            sACK = sACK & "DSP|6|||||" & chrCR
            sACK = sACK & "DSP|7|||||" & chrCR
            sACK = sACK & "DSP|8|||||" & chrCR
            sACK = sACK & "DSP|9|||||" & chrCR
            sACK = sACK & "DSP|10|||||" & chrCR
            sACK = sACK & "DSP|11|||||" & chrCR
            sACK = sACK & "DSP|12|||||" & chrCR
            sACK = sACK & "DSP|13|||||" & chrCR
            sACK = sACK & "DSP|14|||||" & chrCR
            sACK = sACK & "DSP|15|||||" & chrCR
            sACK = sACK & "DSP|16|||||" & chrCR
            sACK = sACK & "DSP|17|||||" & chrCR
            sACK = sACK & "DSP|18|||||" & chrCR
            sACK = sACK & "DSP|19|||||" & chrCR
            sACK = sACK & "DSP|20|||||" & chrCR
            sACK = sACK & "DSP|21||12345|||" & chrCR
            sACK = sACK & "DSP|22||12345|||" & chrCR
            sACK = sACK & "DSP|23|||||" & chrCR
            sACK = sACK & "DSP|24||N|||" & chrCR
            sACK = sACK & "DSP|25|||||" & chrCR
            sACK = sACK & "DSP|26||Serum|||" & chrCR
            sACK = sACK & "DSP|27|||||" & chrCR
            sACK = sACK & "DSP|28|||||" & chrCR
            sACK = sACK & "DSP|29||TP^^^|||" & chrCR
'            sACK = sACK & "DSP|30|ALB^ALB^^|||" & chrCR
            sACK = sACK & "DSC||" & chrCR & chrFS & chrCR
            'sACK = sACK & chrFS & chrCR

            Winsock1.SendData sACK
            Save_Raw_Data "[TX:" & Format(Time, "hh:nn:ss") & "]" & sACK
        End If
      
    End If
    
    'QCK
'    If i > 0 Then
'        j = InStr(1, argData, chrCR)
'        If j > 0 Then
''            For lRow = 1 To vasExam.DataRowCnt
''                vasExam.Row = lRow
''                vasExam.Col = 1
''                If vasExam.Value = 1 Then
''
''                End If
''            Next lRow
'
'            sTmp = ""
'            sTmp = Mid(argData, 1, j - 1)
'
'            sMSH = ""
'            iCnt = 0
'
'            k = InStr(1, sTmp, "|")
'            Do While k > 0
'                iCnt = iCnt + 1
'                Select Case iCnt
'                Case 3
'                    sMSH = sMSH & "|"
'                Case 4
'                    sMSH = sMSH & "|"
'                Case 5
'                    sMSH = sMSH & "Mindray|"
'                Case 6
'                    sMSH = sMSH & "BS-380|"
'                Case 7
'                    sMSH = sMSH & Format(GetDateFull, "YYYYMMDDHHMMSS") & "|"
'                Case 9
'                    sMSH = sMSH & "QCK^Q02|"
'                Case 10
''                    sMsgID = Mid(sTmp, 1, k)
'''                    iPos = InStr(1, sMsgID, "|")
'''                    If iPos > 0 Then
'''                        sMsgID = sMsgID + 1 & "|"
'''                    End If
''
''                    sMSH = sMSH & sMsgID
'
'                    sMsgID = "1|"
'                    sMSH = sMSH & "1|"
'                Case Else
'                    sMSH = sMSH & Mid(sTmp, 1, k)
'                End Select
'
'                sTmp = Mid(sTmp, k + 1)
'                k = InStr(1, sTmp, "|")
'            Loop
'
'            sACK = ""
'            sACK = sMSH & chrCR
'            sACK = sACK & "MSA|AA|" & sMsgID & "|||0|" & chrCR
'            sACK = sACK & "ERR|0|" & chrCR & "QAK|SR|OK|" & chrCR

    '***********************************************************************
    
    i = InStr(1, argData, "ORU^R01")        'Result*************************
    If i > 0 Then
        j = InStr(1, argData, chrCR)
        If j > 0 Then
            sTmp = ""
            sTmp = Mid(argData, 1, j - 1)
            
            sMSH = ""
            iCnt = 0
            
            k = InStr(1, sTmp, "|")
            Do While k > 0
                iCnt = iCnt + 1
                Select Case iCnt
                Case 3
                    sMSH = sMSH & "|"
                Case 4
                    sMSH = sMSH & "|"
                Case 5
                    sMSH = sMSH & "Mindray|"
                Case 6
                    sMSH = sMSH & "BS-380|"
                Case 7
                    sMSH = sMSH & Format(GetDateFull, "YYYYMMDDHHMMSS") & "|"
                Case 9
                    sMSH = sMSH & "ACK^R01|"
                Case 10
'                    sMsgID = Mid(sTmp, 1, k)
''                    iPos = InStr(1, sMsgID, "|")
''                    If iPos > 0 Then
''                        sMsgID = sMsgID + 1 & "|"
''                    End If
'
'                    sMSH = sMSH & sMsgID
                    
                    sMsgID = "1|"
                    sMSH = sMSH & "1|"
                Case Else
                    sMSH = sMSH & Mid(sTmp, 1, k)
                End Select
                
                sTmp = Mid(sTmp, k + 1)
                k = InStr(1, sTmp, "|")
            Loop
                                                     
            sACK = ""
            sACK = sMSH & chrCR
            'sACK = sACK & "MSA|AA|" & sMsgID & "Message accepted|||0|" & chrCR & chrFS & chrCR
            sACK = sACK & "MSA|AA|" & sMsgID & "|||0|" & chrCR & chrFS & chrCR
        End If
        Winsock1.SendData sACK
        Save_Raw_Data "[TX:" & Format(Time, "hh:nn:ss") & "]" & sACK
    End If
End Sub

Sub zz()

 Dim i, j, k As Long
    Dim X, y As Long
    Dim sPreID As String
    Dim sResult, sResult1 As String
    Dim iPos As Integer
    Dim sSID As String
    
    
    
   sSID = Format(Trim(txtBarCode.Text), "000000####")
   
    ClearSpread vasTemp1
    
    sResult = ""
    sResult1 = ""

    SQL = "Select barcode, seqno, diskno, posno, recedate, pid, " & _
          "       pname, psex, page, jumin1, jumin2, " & _
          "       pid, examgubun, resflag, examcode, " & _
          "       result, refflag, panicflag, deltaflag, examuid,equipcode " & vbCrLf & _
          "From pat_res " & CR & _
          "Where barcode = '" & Trim(sSID) & "' " & CR & _
          "And ResFlag <> 'B' " & CR & _
          "And SampleType <> 'Q' Order by seqno "
          
    res = db_select_Vas(gLocal, SQL, vasTemp1)
    If res = -1 Then
        SaveQuery SQL
    End If
    If vasTemp1.DataRowCnt < 1 Then
        cmdClear_Click
        Exit Sub
    End If
    
    X = 1
    sPreID = Trim(GetText(vasTemp1, 1, 1))
    For j = 1 To 14
        SetText vasExam, Trim(GetText(vasTemp1, 1, j)), X, j + 1
    Next j
    For k = 1 To UBound(gArrEquip)
        If Trim(GetText(vasTemp1, 1, 15)) = gArrEquip(k, 3) Then
            y = 16 + (gArrEquip(k, 1)) * 4 - 3
            Exit For
        End If
    Next k
    
    sResult = Trim(GetText(vasTemp1, 1, 16))
    iPos = InStr(1, sResult, "/")
    sResult1 = Mid(sResult, iPos + 1)
    
    If y > 0 Then
        SetText vasExam, sResult1, X, y
        'SetText vasExam, Trim(GetText(vasTemp1, 1, 16)), x, y
        SetText vasExam, Trim(GetText(vasTemp1, 1, 17)), X, y + 1
        SetText vasExam, Trim(GetText(vasTemp1, 1, 18)), X, y + 2
        SetText vasExam, Trim(GetText(vasTemp1, 1, 19)), X, y + 3
    End If
    
    Select Case Trim(GetText(vasTemp1, 1, 17))
    Case "Pos"  '"H", "P"
        SetBackColor vasExam, X, X, y, y, 255, 149, 149
    Case "Neg"  '"L"
        SetBackColor vasExam, X, X, y, y, 149, 149, 255
    Case Else
        SetBackColor vasExam, X, X, y, y, 255, 255, 255
    End Select

    For i = 2 To vasTemp1.DataRowCnt
        If Trim(GetText(vasTemp1, i, 1)) = sPreID Then
            For k = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp1, i, 15)) = gArrEquip(k, 3) Then
                    y = 16 + gArrEquip(k, 1) * 4 - 3
                    Exit For
                End If
            Next k
            
            sResult = ""
            sResult1 = ""
            
            sResult = Trim(GetText(vasTemp1, i, 16))
            iPos = InStr(1, sResult, "/")
            sResult1 = Mid(sResult, iPos + 1)
    
            SetText vasExam, sResult1, X, y
            'SetText vasExam, Trim(GetText(vasTemp1, i, 16)), x, y
            SetText vasExam, Trim(GetText(vasTemp1, i, 17)), X, y + 1
            SetText vasExam, Trim(GetText(vasTemp1, i, 18)), X, y + 2
            SetText vasExam, Trim(GetText(vasTemp1, i, 19)), X, y + 3

            Select Case Trim(GetText(vasTemp1, i, 17))
            Case "Pos"  '"H", "P"
                SetBackColor vasExam, X, X, y, y, 255, 149, 149
            Case "Neg"  '"L"
                SetBackColor vasExam, X, X, y, y, 149, 149, 255
            Case Else
                SetBackColor vasExam, X, X, y, y, 255, 255, 255
            End Select

        Else
            X = X + 1
            
            If X > vasExam.MaxRows Then
                vasExam.MaxRows = X
            End If
            
            sPreID = Trim(GetText(vasTemp1, i, 1))
            For j = 1 To 14
                SetText vasExam, Trim(GetText(vasTemp1, i, j)), X, j + 1
            Next j
            For k = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp1, i, 15)) = gArrEquip(k, 3) Then
                    y = 16 + gArrEquip(k, 1) * 4 - 3
                    Exit For
                End If
            Next k
            
            
            sResult = ""
            sResult1 = ""
            
            sResult = Trim(GetText(vasTemp1, i, 16))
            iPos = InStr(1, sResult, "/")
            sResult1 = Mid(sResult, iPos + 1)
    
            SetText vasExam, sResult1, X, y
            'SetText vasExam, Trim(GetText(vasTemp1, i, 16)), x, y
            SetText vasExam, Trim(GetText(vasTemp1, i, 17)), X, y + 1
            SetText vasExam, Trim(GetText(vasTemp1, i, 18)), X, y + 2
            SetText vasExam, Trim(GetText(vasTemp1, i, 19)), X, y + 3

            Select Case Trim(GetText(vasTemp1, i, 17))
            Case "Pos"  '"H", "P"
                SetBackColor vasExam, X, X, y, y, 255, 149, 149
            Case "Neg"  '"L"
                SetBackColor vasExam, X, X, y, y, 149, 149, 255
            Case Else
                SetBackColor vasExam, X, X, y, y, 255, 255, 255
            End Select

        End If

    Next i

End Sub

Function Get_Order_barcode(ByVal asBarcode As String, Optional asFlag As Integer = 0)
    Dim lRow, lCol As Long
    Dim i, j, k, z As Integer
    Dim lsID As String
    
    lsID = asBarcode
    
    
        ReDim lsOrder(0)
        
        
        gOrder_Select.ok = 0
        
        giIndex = -1
        ReDim gOrder_List(0)
        
        kbnu_Order_Request lsID, gHPEQUIP
    
        If gOrder_Select.ok = 1 Then
            lRow = -1
            If asFlag = 1 Then
                For i = vasExam.DataRowCnt To 1 Step -1
                    If Trim(GetText(vasExam, i, colID)) = lsID Then
                        lRow = i
                        Exit For
                    End If
                Next i
            End If
            
            If lRow = -1 Then
                lRow = vasExam.DataRowCnt + 1
                If lRow > vasExam.MaxRows Then
                    vasExam.MaxRows = lRow
                End If
            End If
            
            vasExam.Row = lRow
            vasExam.Col = 1
            vasExam.Value = 1
            
            vasExam.SetText colID, lRow, lsID
            vasExam.SetText colReceNo, lRow, gOrder_Select.PT_NO
            vasExam.SetText colPID, lRow, gOrder_Select.PT_NO
            vasExam.SetText colPName, lRow, gOrder_Select.PT_NM
            If InStr(1, gOrder_Select.Sex, "/") > 0 Then
                vasExam.SetText colSex, lRow, Left(gOrder_Select.Sex, InStr(1, gOrder_Select.Sex, "/") - 1)
                vasExam.SetText colAge, lRow, Mid(gOrder_Select.Sex, InStr(1, gOrder_Select.Sex, "/") + 1)
            End If
        
            For i = 1 To UBound(gOrder_List)
                'SetText vasExam, mExam(3, i), lRow, 1
                'SetText vasExam, mExam(4, i), lRow, 2
                
                gReadBuf(0) = ""
                gReadBuf(1) = ""
                
                SQL = "Select ExamCode, UnitCode, EquipCode from EquipExam " & vbCrLf & _
                      "where Equip = '" & gEquip & "' " & vbCrLf & _
                      "  and ExamCode = '" & Trim(gOrder_List(i).TST_CD) & "' " & vbCrLf & _
                      "  and UseFlag = 1 "
                res = db_select_Col(gLocal, SQL)
                If Trim(gReadBuf(0)) = Trim(gOrder_List(i).TST_CD) Then
                    '^^^CKMB\^^^Myoglob\^^^Tpn-I
                    
                    'lsOrder = lsOrder & "^^^" & Trim(gReadBuf(1)) & "\"
                    k = -1
                    For j = LBound(lsOrder) To UBound(lsOrder)
                        If Trim(lsOrder(j)) = Trim(gReadBuf(1)) Then
                            k = 1
                            Exit For
                        End If
                    Next j
                    
                    If k = -1 Then
                        z = z + 1
                        ReDim Preserve lsOrder(z)
                        lsOrder(z) = Trim(gReadBuf(1))
                        

                    End If
                        
                    For j = 1 To UBound(gArrEquip)
                        If Trim(gArrEquip(j, 2)) = Trim(gReadBuf(2)) Then
                            'SetText vasList, "*", glRow, gResCol + j
    
                            lCol = (gArrEquip(j, 1) - 1)

                            SetText vasExam, "*", glRow, colResult + lCol * 4
                            'SetText vasExam, lsResult, glRow, colResult1 + lCol
                        
    '                        Save_Local_One glRow, i, "A"
                            Exit For
                        End If
                    Next j
                    
                End If
            Next i
'            If Len(lsOrder) > 0 Then
'                lsOrder = Left(lsOrder, Len(lsOrder) - 1)
'            End If
        End If

End Function

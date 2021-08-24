VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmInterface 
   Caption         =   "Sysmex XE-2100 IPU Interface Program [Service Center ☎(02)6205-1751]"
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15240
   Icon            =   "frmInterface.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   15240
   Begin VB.CommandButton Command4 
      Caption         =   "Command3"
      Height          =   345
      Left            =   8010
      TabIndex        =   75
      Top             =   810
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   345
      Left            =   7920
      TabIndex        =   74
      Top             =   360
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   345
      Left            =   6780
      TabIndex        =   73
      Top             =   810
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   6690
      TabIndex        =   72
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin FPSpread.vaSpread vasTemp2 
      Height          =   1575
      Left            =   8790
      TabIndex        =   61
      Top             =   1170
      Visible         =   0   'False
      Width           =   2265
      _Version        =   393216
      _ExtentX        =   3995
      _ExtentY        =   2778
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
      SpreadDesigner  =   "frmInterface.frx":0442
   End
   Begin FPSpread.vaSpread vasTemp1 
      Height          =   1575
      Left            =   6960
      TabIndex        =   60
      Top             =   3030
      Visible         =   0   'False
      Width           =   2265
      _Version        =   393216
      _ExtentX        =   3995
      _ExtentY        =   2778
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
      SpreadDesigner  =   "frmInterface.frx":0647
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   1545
      Left            =   3900
      TabIndex        =   71
      Top             =   2730
      Visible         =   0   'False
      Width           =   2955
      _Version        =   393216
      _ExtentX        =   5212
      _ExtentY        =   2725
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
      SpreadDesigner  =   "frmInterface.frx":084C
   End
   Begin FPSpread.vaSpread vasTemp_2 
      Height          =   1545
      Left            =   30
      TabIndex        =   67
      Top             =   5700
      Visible         =   0   'False
      Width           =   2955
      _Version        =   393216
      _ExtentX        =   5212
      _ExtentY        =   2725
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
      SpreadDesigner  =   "frmInterface.frx":0A51
   End
   Begin FPSpread.vaSpread vasRes_2 
      Height          =   1245
      Left            =   60
      TabIndex        =   68
      Top             =   7500
      Visible         =   0   'False
      Width           =   2925
      _Version        =   393216
      _ExtentX        =   5159
      _ExtentY        =   2196
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
      SpreadDesigner  =   "frmInterface.frx":0C56
   End
   Begin FPSpread.vaSpread vasRes_1 
      Height          =   1245
      Left            =   30
      TabIndex        =   69
      Top             =   4350
      Visible         =   0   'False
      Width           =   2925
      _Version        =   393216
      _ExtentX        =   5159
      _ExtentY        =   2196
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
      SpreadDesigner  =   "frmInterface.frx":51A5
   End
   Begin FPSpread.vaSpread vasTemp_1 
      Height          =   1545
      Left            =   60
      TabIndex        =   70
      Top             =   2910
      Visible         =   0   'False
      Width           =   2955
      _Version        =   393216
      _ExtentX        =   5212
      _ExtentY        =   2725
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
      SpreadDesigner  =   "frmInterface.frx":96F4
   End
   Begin VB.TextBox txtBuff2 
      Height          =   345
      Left            =   600
      TabIndex        =   6
      Top             =   810
      Visible         =   0   'False
      Width           =   6195
   End
   Begin VB.TextBox txtBuff1 
      Height          =   345
      Left            =   750
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   6195
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   90
      Top             =   2010
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   8096
      InputLen        =   1
      RThreshold      =   1
      RTSEnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   90
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   8096
      InputLen        =   1
      RThreshold      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   6435
      Left            =   3030
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   10995
      Begin VB.TextBox txtWorkListNo 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   2580
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtWardRoom 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   2160
         Width           =   1545
      End
      Begin VB.TextBox txtSexAge 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1740
         Width           =   1545
      End
      Begin VB.TextBox txtFlag 
         Appearance      =   0  '평면
         Height          =   1365
         Left            =   6870
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   44
         Top             =   4830
         Width           =   3885
      End
      Begin VB.TextBox txtTube 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   4740
         Width           =   1545
      End
      Begin VB.TextBox txtRack 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   4320
         Width           =   1545
      End
      Begin VB.TextBox txtEquip 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   3900
         Width           =   1545
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "▶"
         Height          =   465
         Left            =   900
         TabIndex        =   37
         Top             =   5640
         Width           =   645
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "◀"
         Height          =   465
         Left            =   240
         TabIndex        =   36
         Top             =   5640
         Width           =   645
      End
      Begin Threed.SSCommand cmdCloseDetail 
         Height          =   495
         Left            =   1560
         TabIndex        =   22
         Top             =   5625
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "닫기"
      End
      Begin VB.TextBox txtResDate 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   3420
         Width           =   2475
      End
      Begin VB.TextBox txtPName 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1275
         Width           =   1545
      End
      Begin VB.TextBox txtPID 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   825
         Width           =   1545
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   390
         Width           =   1545
      End
      Begin FPSpread.vaSpread vasRes1 
         Height          =   5865
         Left            =   2940
         TabIndex        =   12
         Top             =   330
         Width           =   3885
         _Version        =   393216
         _ExtentX        =   6853
         _ExtentY        =   10345
         _StockProps     =   64
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   20
         RowHeaderDisplay=   0
         ScrollBars      =   0
         SpreadDesigner  =   "frmInterface.frx":98F9
      End
      Begin FPSpread.vaSpread vasRes2 
         Height          =   4485
         Left            =   6870
         TabIndex        =   13
         Top             =   330
         Width           =   3885
         _Version        =   393216
         _ExtentX        =   6853
         _ExtentY        =   7911
         _StockProps     =   64
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   20
         RowHeaderDisplay=   0
         ScrollBars      =   0
         SpreadDesigner  =   "frmInterface.frx":9EFA
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "W/L No."
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   51
         Top             =   2640
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "병동병실"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   49
         Top             =   2220
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "성별나이"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   47
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Tube"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   43
         Top             =   4800
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Rack"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   41
         Top             =   4380
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사장비"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   39
         Top             =   3960
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과시간"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   20
         Top             =   3135
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "환자이름"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   18
         Top             =   1335
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "등록번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   16
         Top             =   885
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   14
         Top             =   450
         Width           =   840
      End
   End
   Begin VB.Frame frameSch 
      Height          =   8805
      Left            =   120
      TabIndex        =   33
      Top             =   930
      Visible         =   0   'False
      Width           =   15075
      Begin FPSpread.vaSpread vasSchESR 
         Height          =   7515
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Visible         =   0   'False
         Width           =   14835
         _Version        =   393216
         _ExtentX        =   26167
         _ExtentY        =   13256
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   12
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":A4FB
         UserResize      =   2
      End
      Begin VB.CommandButton cmdCol2 
         Caption         =   "<"
         Height          =   315
         Left            =   2010
         TabIndex        =   54
         Top             =   240
         Width           =   225
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check1"
         Height          =   225
         Left            =   750
         TabIndex        =   66
         Top             =   300
         Width           =   165
      End
      Begin FPSpread.vaSpread vasSch 
         Height          =   7845
         Left            =   120
         TabIndex        =   45
         Top             =   270
         Visible         =   0   'False
         Width           =   14835
         _Version        =   393216
         _ExtentX        =   26167
         _ExtentY        =   13838
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   0
         ColsFrozen      =   13
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
         MaxCols         =   52
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":E407
         UserResize      =   2
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "전송"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   11880
         TabIndex        =   56
         Top             =   8220
         Width           =   1485
      End
      Begin VB.CommandButton cmdSchClose 
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   13440
         TabIndex        =   34
         Top             =   8220
         Width           =   1485
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   225
      Left            =   750
      TabIndex        =   65
      Top             =   960
      Width           =   165
   End
   Begin FPSpread.vaSpread vasResTemp 
      Height          =   1155
      Left            =   30
      TabIndex        =   64
      Top             =   6420
      Visible         =   0   'False
      Width           =   2055
      _Version        =   393216
      _ExtentX        =   3625
      _ExtentY        =   2037
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
      SpreadDesigner  =   "frmInterface.frx":12F08
   End
   Begin FPSpread.vaSpread vasExam2 
      Height          =   4665
      Left            =   7770
      TabIndex        =   63
      Top             =   4650
      Visible         =   0   'False
      Width           =   2385
      _Version        =   393216
      _ExtentX        =   4207
      _ExtentY        =   8229
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
      SpreadDesigner  =   "frmInterface.frx":1310D
   End
   Begin FPSpread.vaSpread vasExam1 
      Height          =   4575
      Left            =   3780
      TabIndex        =   62
      Top             =   4860
      Visible         =   0   'False
      Width           =   2655
      _Version        =   393216
      _ExtentX        =   4683
      _ExtentY        =   8070
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
      SpreadDesigner  =   "frmInterface.frx":13312
   End
   Begin VB.CommandButton cmdCol1 
      Caption         =   "<"
      Height          =   315
      Left            =   1980
      TabIndex        =   55
      Top             =   930
      Width           =   225
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   30
      Top             =   4650
   End
   Begin VB.TextBox txtTemp 
      Height          =   270
      Left            =   30
      TabIndex        =   8
      Top             =   870
      Visible         =   0   'False
      Width           =   705
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   585
      Left            =   90
      TabIndex        =   1
      Top             =   9780
      Width           =   15075
      _Version        =   65536
      _ExtentX        =   26591
      _ExtentY        =   1032
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelOuter      =   1
      Begin VB.CommandButton cmdClose 
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   13500
         TabIndex        =   30
         Top             =   60
         Width           =   1485
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   11910
         TabIndex        =   29
         Top             =   60
         Width           =   1485
      End
      Begin VB.CommandButton cmd_Trans 
         Caption         =   "전송"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   10290
         TabIndex        =   28
         Top             =   60
         Width           =   1485
      End
      Begin Threed.SSPanel sspPort 
         Height          =   465
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7064
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "   IPU1"
         ForeColor       =   0
         BackColor       =   16777215
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
         Alignment       =   1
         Begin VB.Label lblIPU1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "연결"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   1530
            TabIndex        =   10
            Top             =   120
            Width           =   360
         End
         Begin VB.Label lblIPU1Com 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "[COM1]9600,n,8,1"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   2010
            TabIndex        =   9
            Top             =   120
            Width           =   1530
         End
      End
      Begin Threed.SSPanel sspState 
         Height          =   465
         Left            =   4080
         TabIndex        =   25
         Top             =   60
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7064
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "   IPU2"
         ForeColor       =   0
         BackColor       =   16777215
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
         Alignment       =   1
         Begin VB.Label lblIPU2Com 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "[COM2]9600,n,8,1"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   1920
            TabIndex        =   27
            Top             =   120
            Width           =   1530
         End
         Begin VB.Label lblIPU2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "연결"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   1440
            TabIndex        =   26
            Top             =   120
            Width           =   360
         End
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   90
      TabIndex        =   3
      Top             =   60
      Width           =   15075
      _Version        =   65536
      _ExtentX        =   26591
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "     Sysmex XE-2100 IPU Interface"
      ForeColor       =   16056319
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
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   435
         Left            =   3990
         TabIndex        =   80
         Top             =   60
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox Text1 
         Height          =   420
         Left            =   420
         TabIndex        =   79
         Top             =   120
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.CommandButton cmdChangeUser 
         Caption         =   "사용자변경"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   13410
         TabIndex        =   78
         Top             =   30
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   11130
         Picture         =   "frmInterface.frx":13517
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   76
         Top             =   30
         Width           =   315
      End
      Begin VB.CheckBox chkRange 
         BackColor       =   &H00F4FFFF&
         Caption         =   "구간출력"
         Height          =   180
         Left            =   9630
         TabIndex        =   59
         Top             =   270
         Width           =   1065
      End
      Begin VB.TextBox txtBarcode 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4980
         TabIndex        =   57
         Top             =   195
         Width           =   1545
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "출력"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   8220
         TabIndex        =   53
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton cmdSch 
         Caption         =   "검색"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6840
         TabIndex        =   35
         Top             =   120
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker dtpExamDate 
         Height          =   345
         Left            =   1890
         TabIndex        =   24
         Top             =   210
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   97452033
         CurrentDate     =   38584
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   585
         Left            =   10350
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   -390
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtRemark 
         Height          =   435
         Left            =   2220
         TabIndex        =   4
         Top             =   1230
         Width           =   1545
      End
      Begin VB.Label lblUser 
         BackColor       =   &H00F4FFFF&
         Caption         =   "사용자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11550
         TabIndex        =   77
         Top             =   60
         Width           =   2295
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "바코드번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3870
         TabIndex        =   58
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   31
         Top             =   270
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "국립암센터 진단검사의학과"
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
         Left            =   11130
         TabIndex        =   23
         Top             =   450
         Width           =   2820
      End
   End
   Begin FPSpread.vaSpread vasComList 
      Height          =   1125
      Left            =   3060
      TabIndex        =   32
      Top             =   1710
      Visible         =   0   'False
      Width           =   7215
      _Version        =   393216
      _ExtentX        =   12726
      _ExtentY        =   1984
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   11
      MaxRows         =   3
      RowHeaderDisplay=   0
      ScrollBars      =   0
      SpreadDesigner  =   "frmInterface.frx":13AA1
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   8805
      Left            =   120
      TabIndex        =   0
      Top             =   930
      Width           =   15075
      _Version        =   393216
      _ExtentX        =   26591
      _ExtentY        =   15531
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   0
      ColsFrozen      =   13
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
      MaxCols         =   100
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmInterface.frx":1414B
      UserResize      =   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   "파일"
      Begin VB.Menu subChangeUser 
         Caption         =   "사용자변경"
         Visible         =   0   'False
      End
      Begin VB.Menu subN11 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu subSendQC 
         Caption         =   "QC 결과 전송"
         Visible         =   0   'False
      End
      Begin VB.Menu subN12 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu subClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu subN1 
         Caption         =   "-"
      End
      Begin VB.Menu subClose 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnuPort 
      Caption         =   "연결"
      Begin VB.Menu subComConnect 
         Caption         =   "장비연결"
      End
      Begin VB.Menu subN2 
         Caption         =   "-"
      End
      Begin VB.Menu subSendMode 
         Caption         =   "서버 결과 전송"
         Begin VB.Menu subSend1 
            Caption         =   "Auto"
         End
         Begin VB.Menu subSend2 
            Caption         =   "Manual"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "설정"
      Begin VB.Menu subCodeSet 
         Caption         =   "검사코드설정"
      End
      Begin VB.Menu subComSetup 
         Caption         =   "통신설정"
      End
   End
   Begin VB.Menu mnuWorkList 
      Caption         =   "WorkList"
      Visible         =   0   'False
      Begin VB.Menu subWorkList 
         Caption         =   "WorkList - Socket"
      End
      Begin VB.Menu subLASCOrder 
         Caption         =   "WorkList - SP"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "검색"
      Begin VB.Menu subSearch 
         Caption         =   "날짜별 검사내역"
      End
      Begin VB.Menu subSchESR 
         Caption         =   "ESR 검체조회"
      End
      Begin VB.Menu subProficiency 
         Caption         =   "기기간 DATA 비교"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Individual Order Format

Dim gResCol As Long
Dim gMaxCol As Long

Dim gCurRow As Long
Dim gCurRow1 As Long

Dim gID As String
Dim gID1 As String
Dim gRack As String
Dim gRack1 As String
Dim gPos As String
Dim gPos1 As String
Dim gFlag As String
Dim gFlag1 As String

Dim iRow1 As Long
Dim iRow2 As Long
Dim iCol1 As Long
Dim iCol2 As Long

Dim flagWBC As Long
Dim flagRBC As Long
Dim flagPLT As Long

Dim SelVas As Integer

Dim DelaySP As String

Dim lsWBC1 As String
Dim lsNEUT1 As String
Dim lsWBC2 As String
Dim lsNEUT2 As String

Dim gsVersion As String

Dim iMsgCnt1 As Integer
Dim gOrder1 As String

Dim iMsgCnt2 As Integer
Dim gOrder2 As String

Sub GetOPtion()
    Dim db_tmp As String * 20

    db_tmp = ""
    Call GetPrivateProfileString("OPTION", "Delay_SP", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    DelaySP = Trim(txtTemp)
    If Not IsNumeric(DelaySP) Then DelaySP = "10"
    
End Sub

Private Sub Check1_Click()
    vasList.Row = -1
    vasList.Col = 1
    vasList.Value = Check1.Value
End Sub

Private Sub Check2_Click()
    vasSch.Row = -1
    vasSch.Col = 1
    vasSch.Value = Check2.Value
End Sub

Private Sub chkMode_Click()
    If chkMode.Value = 1 Then
        chkMode.Caption = "Auto"
        SaveSetting "MEDIMATE", "XE2100", "SendMode", "1"
    Else
        chkMode.Caption = "Manual"
        SaveSetting "MEDIMATE", "XE2100", "SendMode", "0"
    End If
End Sub

Private Sub chkRange_Click()
    Dim asSpread As vaSpread
    
    If chkRange.Value = 1 Then
        If vasSch.Visible = True Then
            Set asSpread = vasSch
        Else
            Set asSpread = vasList
        End If
        
        asSpread.ColWidth(1) = 0
        asSpread.ColWidth(3) = 8
        asSpread.ColWidth(4) = 8
        asSpread.ColWidth(5) = 0
        asSpread.ColWidth(6) = 0
        asSpread.ColWidth(7) = 0
        asSpread.ColWidth(8) = 0
        asSpread.ColWidth(9) = 0
        asSpread.ColWidth(10) = 0
        asSpread.ColWidth(11) = 0
        asSpread.ColWidth(12) = 0
        asSpread.ColWidth(13) = 0
    Else
        If vasSch.Visible = True Then
            Set asSpread = vasSch
        Else
            Set asSpread = vasList
        End If
        
        asSpread.ColWidth(1) = 2.38
        asSpread.ColWidth(3) = 8
        asSpread.ColWidth(4) = 8
        asSpread.ColWidth(5) = 0
        asSpread.ColWidth(6) = 2.88
        asSpread.ColWidth(7) = 5.25
        asSpread.ColWidth(8) = 9
        asSpread.ColWidth(9) = 6.38
        asSpread.ColWidth(10) = 4.38
        asSpread.ColWidth(11) = 7
        asSpread.ColWidth(12) = 5.13
        asSpread.ColWidth(13) = 7.5
    
    End If
End Sub

Private Sub cmd_Trans_Click()
    Dim lRow, i, liEquipCode As Long
    Dim lsID As String
    Dim lsResult As String
    Dim lsInscode As String
    Dim lsWBC As String
    Dim lsNRBC As String
    Dim lsEOSIN As String
    Dim lsC_WBC As String
    Dim liRet, j, liMustBe As Integer
    Dim liCheck1, liCheck2 As Integer
    
    Dim lsSmearCode As String
    
    Dim mExam
    
    If MsgBox(" " & vbCrLf & "검사 결과를 전송하시겠습니까?" & vbCrLf & " ", vbInformation + vbYesNo + vbDefaultButton2, "결과 전송 알림") = vbNo Then
        Exit Sub
    End If
    
    For lRow = 1 To vasList.DataRowCnt
        vasList.Row = lRow
        vasList.Col = 1
        
        lsWBC = ""
        lsNRBC = ""
        lsEOSIN = ""
        
        If vasList.Value = 1 Then
            lsID = Trim(GetText(vasList, lRow, 2))
            
            If Trim(GetText(vasList, lRow, 5)) = "IPU1" Then
'                lsInscode = "02"
                lsInscode = IPU1.UseEquip
            Else
'                lsInscode = "01"
                lsInscode = IPU2.UseEquip
            End If
            
            'lsInscode = "01"
            
            res = ToServer(lRow, vasList)
            If res = 1 Then
                vasList.Row = lRow
                vasList.Col = 1
                vasList.Value = 0

                SetText vasList, "완료", lRow, gResCol
                SetBackColor vasList, lRow, lRow, 1, 1, 202, 255, 112
                'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
            ElseIf res = 2 Then
                SetText vasList, "결과", lRow, gResCol
            Else
                SetText vasList, "실패", lRow, gResCol
                SetForeColor vasList, lRow, lRow, 1, 1, 255, 0, 0
                'SetBackColor vasID, gPreRow, gPreRow, colCheckBox, vasID.MaxCols, 255, 255, 255
                'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
            End If
        
        
        End If
    Next lRow
End Sub

Private Sub cmdChangeUser_Click()
    frmUserChange.Show 1
End Sub

Private Sub cmdClear_Click()
'vsSpread의 내용을 Clear 한다.
    vasList.Row = 1
    vasList.Col = 1
    vasList.Row2 = vasList.MaxRows
    vasList.Col2 = vasList.MaxCols
    vasList.BlockMode = True
    vasList.Action = 3
    vasList.BackColor = RGB(255, 255, 255)
    vasList.ForeColor = RGB(0, 0, 0)
    vasList.BlockMode = False

    vasList.Row = 1
    vasList.Col = 1
    vasList.Row2 = vasList.MaxRows
    vasList.Col2 = 1
    vasList.BlockMode = True
    vasList.Value = 0
    vasList.BlockMode = False

    txtBuff1 = ""
    txtBuff2 = ""
    
    gCurRow = -1
    gCurRow1 = -1
    
    ReDim gArrExamRes(0)
    GetExamCode
    
    GetOPtion
End Sub

Private Sub cmdCode_Click()
    frmCode.Show 1
End Sub

Private Sub cmdComSetup_Click()
    frmConfig.Show 1
End Sub

Private Sub cmdConnect_Click()
    frmConnect.Show 1
    
    If IPU1.ConnectFlag Then
        If MSComm1.PortOpen = False Then
            MSComm1.CommPort = IPU1.ComPort
            MSComm1.Settings = IPU1.Speed & "," & IPU1.Parity & "," & IPU1.DataBit & "," & IPU1.StartBit
            If IPU1.RTSEnable = "1" Then
                MSComm1.RTSEnable = True
            Else
                MSComm1.RTSEnable = False
            End If
            If IPU1.DTREnable = "1" Then
                MSComm1.DTREnable = True
            Else
                MSComm1.DTREnable = False
            End If
            MSComm1.PortOpen = True
            
            lblIPU1.ForeColor = RGB(0, 255, 0)
        End If
    Else
        If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
        lblIPU1.ForeColor = RGB(255, 0, 0)
    End If

    If IPU2.ConnectFlag Then
        If MSComm2.PortOpen = False Then
            MSComm2.CommPort = IPU2.ComPort
            MSComm2.Settings = IPU2.Speed & "," & IPU2.Parity & "," & IPU2.DataBit & "," & IPU2.StartBit
            If IPU2.RTSEnable = "1" Then
                MSComm2.RTSEnable = True
            Else
                MSComm2.RTSEnable = False
            End If
            If IPU2.DTREnable = "1" Then
                MSComm2.DTREnable = True
            Else
                MSComm2.DTREnable = False
            End If
            MSComm2.PortOpen = True
        
            lblIPU2.ForeColor = RGB(0, 255, 0)
        End If
    Else
        If MSComm2.PortOpen = True Then MSComm2.PortOpen = False
        lblIPU2.ForeColor = RGB(255, 0, 0)
    End If

End Sub

Private Sub cmdClose_Click()
    subClose_Click
End Sub

Private Sub cmdCloseDetail_Click()
    Frame1.Visible = False
End Sub

Private Sub cmdCol1_Click()
    If cmdCol1.Caption = ">" Then
        vasList.ColWidth(3) = 8
        vasList.ColWidth(4) = 8
        vasList.ColWidth(5) = 0
        vasList.ColWidth(6) = 2.88
        vasList.ColWidth(7) = 5.25
        vasList.ColWidth(8) = 9
        
        'vasList.ColWidth(10) = 4.38
        vasList.ColWidth(10) = 8
        vasList.ColWidth(11) = 7
        vasList.ColWidth(12) = 5.13
        vasList.ColWidth(13) = 7.5
        
        cmdCol1.Caption = "<"
    Else
        vasList.ColWidth(3) = 0
        vasList.ColWidth(4) = 0
        vasList.ColWidth(5) = 0
        vasList.ColWidth(6) = 0
        vasList.ColWidth(7) = 0
        vasList.ColWidth(8) = 0
        
        vasList.ColWidth(10) = 0
        vasList.ColWidth(11) = 0
        vasList.ColWidth(12) = 0
        vasList.ColWidth(13) = 0
        
        cmdCol1.Caption = ">"
    End If

End Sub

Private Sub cmdCol2_Click()
    If cmdCol2.Caption = ">" Then
        vasSch.ColWidth(3) = 8
        vasSch.ColWidth(4) = 8
        vasSch.ColWidth(5) = 0
        vasSch.ColWidth(6) = 2.88
        vasSch.ColWidth(7) = 5.25
        vasSch.ColWidth(8) = 9
        
        'vasSch.ColWidth(10) = 4.38
        vasList.ColWidth(10) = 8
        vasSch.ColWidth(11) = 7
        vasSch.ColWidth(12) = 5.13
        vasSch.ColWidth(13) = 7.5
        
        cmdCol2.Caption = "<"
    Else
        vasSch.ColWidth(3) = 0
        vasSch.ColWidth(4) = 0
        vasSch.ColWidth(5) = 0
        vasSch.ColWidth(6) = 0
        vasSch.ColWidth(7) = 0
        vasSch.ColWidth(8) = 0
        
        vasSch.ColWidth(10) = 0
        vasSch.ColWidth(11) = 0
        vasSch.ColWidth(12) = 0
        vasSch.ColWidth(13) = 0
        
        cmdCol2.Caption = ">"
    End If
End Sub

Private Sub cmdNext_Click()
    Dim argSpread As vaSpread
    Dim argRes As vaSpread
    
    Dim lRow1, lRow, lCol, i As Long
    
    If SelVas = 1 Then
        Set argSpread = vasList
    ElseIf SelVas = 2 Then
        Set argSpread = vasSch
    End If
    lRow1 = argSpread.ActiveRow
    lRow1 = lRow1 + 1
    
    vasActiveCell argSpread, lRow1, 2
    
    If lRow1 = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
    ElseIf lRow1 = argSpread.DataRowCnt Then
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
    Else
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
    End If
    
    If argSpread.DataRowCnt = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
    End If
    
    
    txtID = ""
    txtPID = ""
    txtPName = ""
    txtResDate = ""
    txtEquip = ""
    ClearSpread vasRes1
    ClearSpread vasRes2
    
    FlagComment Trim(GetText(argSpread, lRow1, gMaxCol)), Trim(GetText(argSpread, lRow1, gMaxCol + 1)), Trim(GetText(argSpread, lRow1, gMaxCol + 2)), Trim(GetText(argSpread, lRow1, gMaxCol + 3))
    
    txtID = Trim(GetText(argSpread, lRow1, 2))
    txtPID = Trim(GetText(argSpread, lRow1, 3))
    txtPName = Trim(GetText(argSpread, lRow1, 4))
    
    txtSexAge = Trim(GetText(argSpread, lRow1, 6)) & "/" & Trim(GetText(argSpread, lRow1, 7))
    txtWardRoom = Trim(GetText(argSpread, lRow1, 8))
    txtWorkListNo = Trim(GetText(argSpread, lRow1, 9))
    txtRack = Trim(GetText(argSpread, lRow1, 11))
    txtTube = Trim(GetText(argSpread, lRow1, 12))
    
    Select Case Trim(GetText(argSpread, lRow1, 10))
    Case "IPU1"
        txtEquip = "XE2100-1"
    Case "IPU2"
        txtEquip = "XE2100-2"
    End Select
    
    SQL = "Select resdate from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND barcode = '" & Trim(txtID) & "' "
    res = db_select_Text(gLocal, SQL, txtResDate)
    
    'lCol = gResCol
    Set argRes = vasRes1
    lRow = 0
    For lCol = gResCol + 1 To gMaxCol - 1
        If Trim(GetText(argSpread, lRow1, lCol)) <> "" Then
            lRow = lRow + 1
            If lRow = 21 Then
                lRow = 1
                Set argRes = vasRes2
            End If
            
            For i = LBound(gArrExam) To UBound(gArrExam)
                If lCol = gArrExam(i, 10) Then
                    SetText argRes, gArrExam(i, 1), lRow, 1
                    Exit For
                End If
            Next i
            
            'SetText argRes, gArrExam(lCol - gResCol, 1), lRow, 1
            SetText argRes, Trim(GetText(argSpread, lRow1, lCol)), lRow, 3
            SetText argRes, Trim(GetText(argSpread, 0, lCol)), lRow, 2
            
            argSpread.Row = lRow1
            argSpread.Col = lCol
            Select Case argSpread.ForeColor
            Case RGB(255, 127, 0)
                SetForeColor argRes, lRow, lRow, 4, 4, 255, 127, 0
                SetText argRes, "▲", lRow, 4
            Case RGB(0, 127, 255)
                SetForeColor argRes, lRow, lRow, 4, 4, 0, 127, 255
                SetText argRes, "▼", lRow, 4
            Case Else
                SetText argRes, "", lRow, 4
            End Select
        
        End If
    Next lCol
    
    SQL = " Select message From pat_resmemo " & vbCrLf & _
          " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(txtID) & "' "
    res = db_select_Col(gLocal, SQL)
    If Trim(txtFlag) = "" Then
        txtFlag = Trim(gReadBuf(0))
    Else
        txtFlag = txtFlag & vbCrLf & Trim(gReadBuf(0))
    End If
    
End Sub

Private Sub cmdPrev_Click()
    Dim argSpread As vaSpread
    Dim argRes As vaSpread
    
    Dim lRow1, lRow, lCol, i As Long
    
    If SelVas = 1 Then
        Set argSpread = vasList
    ElseIf SelVas = 2 Then
        Set argSpread = vasSch
    End If
    lRow1 = argSpread.ActiveRow
    lRow1 = lRow1 - 1
    
    vasActiveCell argSpread, lRow1, 2
    
    If lRow1 = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
    ElseIf lRow1 = argSpread.DataRowCnt Then
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
    Else
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
    End If
    
    If argSpread.DataRowCnt = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
    End If
    
    txtID = ""
    txtPID = ""
    txtPName = ""
    txtResDate = ""
    txtEquip = ""
    ClearSpread vasRes1
    ClearSpread vasRes2
    
    FlagComment Trim(GetText(argSpread, lRow1, gMaxCol)), Trim(GetText(argSpread, lRow1, gMaxCol + 1)), Trim(GetText(argSpread, lRow1, gMaxCol + 2)), Trim(GetText(argSpread, lRow1, gMaxCol + 3))
    
    txtID = Trim(GetText(argSpread, lRow1, 2))
    txtPID = Trim(GetText(argSpread, lRow1, 3))
    txtPName = Trim(GetText(argSpread, lRow1, 4))
    
    txtSexAge = Trim(GetText(argSpread, lRow1, 6)) & "/" & Trim(GetText(argSpread, lRow1, 7))
    txtWardRoom = Trim(GetText(argSpread, lRow1, 8))
    txtWorkListNo = Trim(GetText(argSpread, lRow1, 9))
    txtRack = Trim(GetText(argSpread, lRow1, 11))
    txtTube = Trim(GetText(argSpread, lRow1, 12))
    Select Case Trim(GetText(argSpread, lRow1, 10))
    Case "IPU1"
        txtEquip = "XE2100-1"
    Case "IPU2"
        txtEquip = "XE2100-2"
    End Select
    
    SQL = "Select resdate from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND barcode = '" & Trim(txtID) & "' "
    res = db_select_Text(gLocal, SQL, txtResDate)
    
    'lCol = gResCol
    Set argRes = vasRes1
    lRow = 0
    For lCol = gResCol + 1 To gMaxCol - 1
        If Trim(GetText(argSpread, lRow1, lCol)) <> "" Then
            lRow = lRow + 1
            If lRow = 21 Then
                lRow = 1
                Set argRes = vasRes2
            End If
            
            For i = LBound(gArrExam) To UBound(gArrExam)
                If lCol = gArrExam(i, 10) Then
                    SetText argRes, gArrExam(i, 1), lRow, 1
                    Exit For
                End If
            Next i
            
            'SetText argRes, gArrExam(lCol - gResCol, 1), lRow, 1
            SetText argRes, Trim(GetText(argSpread, lRow1, lCol)), lRow, 3
            SetText argRes, Trim(GetText(argSpread, 0, lCol)), lRow, 2
            
            argSpread.Row = lRow1
            argSpread.Col = lCol
            Select Case argSpread.ForeColor
            Case RGB(255, 127, 0)
                SetForeColor argRes, lRow, lRow, 4, 4, 255, 127, 0
                SetText argRes, "▲", lRow, 4
            Case RGB(0, 127, 255)
                SetForeColor argRes, lRow, lRow, 4, 4, 0, 127, 255
                SetText argRes, "▼", lRow, 4
            Case Else
                SetText argRes, "", lRow, 4
            End Select
        
        End If
    Next lCol
    
    SQL = " Select message From pat_resmemo " & vbCrLf & _
          " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(txtID) & "' "
    res = db_select_Col(gLocal, SQL)
    If Trim(txtFlag) = "" Then
        txtFlag = Trim(gReadBuf(0))
    Else
        txtFlag = txtFlag & vbCrLf & Trim(gReadBuf(0))
    End If
    
End Sub

Private Sub cmdPrint_Click()
    Dim sHead As String
    Dim sFoot As String
    Dim sCurDate As String
    
    If vasSch.Visible = True Then
        If vasSch.DataRowCnt < 1 Then
            MsgBox "출력할 자료가 없습니다.", , "알 림"
            Exit Sub
        End If
    
        sCurDate = GetDateFull
        
        sHead = "/fn""궁서체"" /fz""15"" /fb1 /fi0 /fu0 " & "/c" & "▣ XE2100 검사현황 ▣" & "/n/n " & _
                    "/fn""굴림체"" /fz""11"" /fb0 /fi0 /fu0 " & "/c" & "조회 일자 : " & dtpExamDate.Value & "/n" & _
                    "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/rPage /p" & "/n"
        sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & "부산백병원 진단검사의학과"
        
        vasSch.PrintOrientation = 2  ' SS_PRINTORIENT_PORTRAIT
        vasSch.PrintAbortMsg = "인쇄중 입니다 ..."
        vasSch.PrintJobName = "XE2100 - 검사현황"
        vasSch.PrintHeader = sHead
        vasSch.PrintFooter = sFoot
        vasSch.PrintMarginTop = 720
        vasSch.PrintMarginBottom = 720
    '현재 SS가 비대칭으로 출력함
    '    vassch.PrintMarginLeft = 720
        vasSch.PrintMarginLeft = 300
        vasSch.PrintMarginRight = 300
        
        vasSch.PrintColor = False
        vasSch.PrintGrid = True
    'Set printing range
        If chkRange.Value = 1 Then
            vasSch.Row = iRow1
            vasSch.Row2 = iRow2
            vasSch.Col = iCol1
            vasSch.Col2 = iCol2
            vasSch.PrintType = PrintTypeCellRange
        Else
            vasSch.PrintType = 0  'SS_PRINT_ALL(default)
        End If
        
        vasSch.PrintShadows = True
    
        vasSch.Action = 13 'SS_ACTION_PRINT
    
    End If
    
    If vasSchESR.Visible = True Then
        If vasSchESR.DataRowCnt < 1 Then
            MsgBox "출력할 자료가 없습니다.", , "알 림"
            Exit Sub
        End If
    
        sCurDate = GetDateFull
        
        sHead = "/fn""궁서체"" /fz""15"" /fb1 /fi0 /fu0 " & "/c" & "▣ ESR 검체현황 ▣" & "/n/n " & _
                    "/fn""굴림체"" /fz""11"" /fb0 /fi0 /fu0 " & "/c" & "조회 일자 : " & dtpExamDate.Value & "/n" & _
                    "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/rPage /p" & "/n"
        sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & "부산백병원 진단검사의학과"
        
        vasSchESR.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
        vasSchESR.PrintAbortMsg = "인쇄중 입니다 ..."
        vasSchESR.PrintJobName = "XE2100 - ESR 검체현황"
        vasSchESR.PrintHeader = sHead
        vasSchESR.PrintFooter = sFoot
        vasSchESR.PrintMarginTop = 720
        vasSchESR.PrintMarginBottom = 720
    '현재 SS가 비대칭으로 출력함
    '    vasschesr.PrintMarginLeft = 720
        vasSchESR.PrintMarginLeft = 300
        vasSchESR.PrintMarginRight = 300
        
        vasSchESR.PrintColor = False
        vasSchESR.PrintGrid = True
    'Set printing range
        If chkRange.Value = 1 Then
            vasSchESR.PrintType = PrintTypeCellRange
        Else
            vasSchESR.PrintType = 0  'SS_PRINT_ALL(default)
        End If
        vasSchESR.PrintShadows = True
    
        vasSchESR.Action = 13 'SS_ACTION_PRINT
    End If
    
    If vasSch.Visible = False And vasSchESR.Visible = False Then
        If vasList.DataRowCnt < 1 Then
            MsgBox "출력할 자료가 없습니다.", , "알 림"
            Exit Sub
        End If
    
        sCurDate = GetDateFull
        
        sHead = "/fn""궁서체"" /fz""15"" /fb1 /fi0 /fu0 " & "/c" & "▣ XE2100 검사현황 ▣" & "/n/n " & _
                    "/fn""굴림체"" /fz""11"" /fb0 /fi0 /fu0 " & "/c" & "검사 일자 : " & dtpExamDate.Value & "/n" & _
                    "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/rPage /p" & "/n"
        sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & "국립암센터 진단검사의학과"
        
        vasList.PrintOrientation = 2  ' SS_PRINTORIENT_PORTRAIT
        vasList.PrintAbortMsg = "인쇄중 입니다 ..."
        vasList.PrintJobName = "XE2100 - 검사현황"
        vasList.PrintHeader = sHead
        vasList.PrintFooter = sFoot
        vasList.PrintMarginTop = 720
        vasList.PrintMarginBottom = 720
    '현재 SS가 비대칭으로 출력함
    '    vaslist.PrintMarginLeft = 720
        vasList.PrintMarginLeft = 300
        vasList.PrintMarginRight = 300
        
        vasList.PrintColor = False
        vasList.PrintGrid = True
    'Set printing range
        If chkRange.Value = 1 Then
            vasList.Row = iRow1
            vasList.Row2 = iRow2
            vasList.Col = iCol1
            vasList.Col2 = iCol2
            vasList.PrintType = PrintTypeCellRange
        Else
            vasList.PrintType = 0  'SS_PRINT_ALL(default)
        End If
        
        vasList.PrintShadows = True
    
        vasList.Action = 13 'SS_ACTION_PRINT
    
    End If
    
    chkRange.Value = 0
End Sub

Private Sub cmdSch_Click()
    Dim lRow, lCol As Long
    Dim lsID, lsType As String
    Dim liEquipCode As Integer
    Dim i As Integer
    
    Dim rs_Res As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    vasSchESR.Visible = False
    vasSch.Visible = True
    cmdCol2.Visible = True
    
    ClearSpread vasSch, 0, 2
    
    Me.MousePointer = 11
    
'    For lCol = 1 To vasList.MaxCols
'        SetText vasSch, Trim(GetText(vasList, 0, lCol)), 0, lCol
'    Next lCol

    vasSch.MaxCols = vasList.MaxCols
    For lCol = 1 To vasList.MaxCols
        SetText vasSch, Trim(GetText(vasList, 0, lCol)), 0, lCol
        vasSch.ColWidth(lCol) = vasList.ColWidth(lCol)
    Next lCol
    
    cmdPrint.Visible = True
    frameSch.Visible = True
    
    
    SQL = "Select a.barcode, a.pid, a.pname, a.pjumin, a.psex, a.page1, a.WardRoom, a.ReceNo, a.examtype, a.diskno, " & _
            "a.posno, a.sendflag, a.equipcode, a.examcode, a.result, " & _
            "a.refflag, a.resdate, a.seqno, b.WBCSusp, b.RBCSusp, " & _
            "b.PLTSusp, b.SampleJudg, b.PBSFlag  " & vbCrLf & _
          "from pat_res a,res_flag b" & vbCrLf & _
          "where a.examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and b.barcode = a.barcode " & vbCrLf & _
          "Order by a.ReceNo, a.barcode, a.equipcode "
    
    SQL = "Select distinct a.barcode, a.pid, a.pname, a.pjumin, a.psex, a.page1, a.WardRoom, a.ReceNo, a.examtype, a.diskno, " & _
            "a.posno, a.sendflag, a.equipcode, '', a.result, " & _
            "a.refflag, a.resdate, a.seqno, b.WBCSusp, b.RBCSusp, " & _
            "b.PLTSusp, b.SampleJudg, b.PBSFlag  " & vbCrLf & _
          "from pat_res a,res_flag b" & vbCrLf & _
          "where a.examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and b.barcode = a.barcode " & vbCrLf & _
          "Order by a.ReceNo, a.barcode, a.examtype, a.equipcode "
    
    Set rs_Res = db_select_rs(gLocal, SQL)
       
    If rs_Res Is Nothing Then GoTo ErrHandle
    
    lsID = ""
    lsType = ""
    lRow = 0
    Do While Not rs_Res.EOF
        'MsgBox Trim(CStr(rs_Res.Fields.Item(0).Value)) & " : " & Trim(CStr(rs_Res.Fields.Item(8).Value))
        
        If Trim(CStr(rs_Res.Fields.Item(0).Value)) <> lsID Or Trim(CStr(rs_Res.Fields.Item(8).Value)) <> lsType Then
            lRow = lRow + 1
            
            If lRow > vasSch.MaxRows Then
                vasSch.MaxRows = lRow
                
                vasSch.RowHeight(lRow) = 12.6
            End If
            
            For lCol = 2 To 13
                If IsNull(rs_Res.Fields.Item(lCol - 2).Value) Then
                    SetText vasSch, "", lRow, lCol
                Else
                    SetText vasSch, Trim(CStr(rs_Res.Fields.Item(lCol - 2).Value)), lRow, lCol
                End If
            Next lCol
            SetText vasSch, Trim(CStr(rs_Res.Fields.Item(0).Value)), lRow, gMaxCol + 4
            
            If IsNumeric(Trim(GetText(vasSch, lRow, 11))) Then  'Rack
                vasSch.SetText 11, lRow, Format(CDbl(GetText(vasSch, lRow, 11)), "0000")
            End If
            If IsNumeric(Trim(GetText(vasSch, lRow, 12))) Then  'Pos
                vasSch.SetText 12, lRow, Format(CDbl(GetText(vasSch, lRow, 12)), "00")
            End If
            
            For i = 0 To 3
                If Not IsNull(rs_Res.Fields.Item(18 + i).Value) Then
                    SetText vasSch, Trim(CStr(rs_Res.Fields.Item(18 + i).Value)), lRow, gMaxCol + i
                End If
            Next i
            If Not IsNull(rs_Res.Fields.Item(22).Value) Then
                SetText vasSch, Trim(CStr(rs_Res.Fields.Item(22).Value)), lRow, gMaxCol + 5
            End If
            
            Select Case Trim(GetText(vasSch, lRow, gResCol))
            Case "B"
                SetText vasSch, "완료", lRow, gResCol
                vasSch.Row = lRow
                vasSch.Col = 1
                vasSch.Value = 0
            Case "E"
                SetText vasSch, "실패", lRow, gResCol
                vasSch.Row = lRow
                vasSch.Col = 1
                vasSch.Value = 1
            Case Else
                SetText vasSch, "수신", lRow, gResCol
                vasSch.Row = lRow
                vasSch.Col = 1
                vasSch.Value = 1
            End Select
        
            Select Case Trim(GetText(vasSch, lRow, gMaxCol + 3))
            Case "0"
                SetText vasSch, "Negative", lRow, gMaxCol + 3
                SetBackColor vasSch, lRow, lRow, 2, 2, 255, 255, 255
            Case "1"
                SetText vasSch, "Positive", lRow, gMaxCol + 3
                SetBackColor vasSch, lRow, lRow, 2, 2, 255, 224, 193
            Case "2"
                SetText vasSch, "Error", lRow, gMaxCol + 3
                SetBackColor vasSch, lRow, lRow, 2, 2, 255, 224, 193
            Case "3"
                SetText vasSch, "Potive+Error", lRow, gMaxCol + 3
                SetBackColor vasSch, lRow, lRow, 2, 2, 255, 224, 193
            Case "4"
                SetText vasSch, "QC Sample", lRow, gMaxCol + 3
                SetBackColor vasSch, lRow, lRow, 2, 2, 255, 255, 255
            End Select
            
            If Trim(GetText(vasSch, lRow, gMaxCol)) = "" Then
                SetBackColor vasSch, lRow, lRow, flagWBC, flagWBC, 255, 255, 255
            ElseIf Trim(GetText(vasSch, lRow, gMaxCol)) <> "0" Then
                SetBackColor vasSch, lRow, lRow, flagWBC, flagWBC, 193, 234, 255
            Else
                SetBackColor vasSch, lRow, lRow, flagWBC, flagWBC, 255, 255, 255
            End If
            If Trim(GetText(vasSch, lRow, gMaxCol + 1)) = "" Then
                SetBackColor vasSch, lRow, lRow, flagRBC, flagRBC, 255, 255, 255
            ElseIf Trim(GetText(vasSch, lRow, gMaxCol + 1)) <> "0" Then
                SetBackColor vasSch, lRow, lRow, flagRBC, flagRBC, 193, 234, 255
            Else
                SetBackColor vasSch, lRow, lRow, flagRBC, flagRBC, 255, 255, 255
            End If
            If Trim(GetText(vasSch, lRow, gMaxCol + 2)) = "" Then
                SetBackColor vasSch, lRow, lRow, flagPLT, flagPLT, 255, 255, 255
            ElseIf Trim(GetText(vasSch, lRow, gMaxCol + 2)) <> "0" Then
                SetBackColor vasSch, lRow, lRow, flagPLT, flagPLT, 193, 234, 255
            Else
                SetBackColor vasSch, lRow, lRow, flagPLT, flagPLT, 255, 255, 255
            End If
            
'            If Trim(GetText(vasSch, lRow, gMaxCol + 5)) = "1" Then
'                SetBackColor vasSch, lRow, lRow, 2, 2, 255, 224, 193
'            End If
            
            If Trim(GetText(vasSch, lRow, gMaxCol + 5)) = 1 Then
                SetBackColor vasSch, lRow, lRow, 9, 9, 255, 224, 193
'                SetForeColor vasList, lRow, lRow, 9, 9, 0, 0, 160
'                vasList.Row = lRow
'                vasList.Col = 9
'                vasList.FontBold = True
            Else
                SetBackColor vasSch, lRow, lRow, 9, 9, 255, 255, 255
'                SetForeColor vasList, lRow, lRow, 9, 9, 0, 0, 0
'                vasList.Row = lRow
'                vasList.Col = 9
'                vasList.FontBold = False
            End If
            
        End If
        
        For liEquipCode = 1 To UBound(gArrExam)
            If CInt(gArrExam(liEquipCode, 1)) = Trim(rs_Res.Fields.Item(12).Value) Then
                'lCol = gResCol + liEquipCode
                lCol = gArrExam(liEquipCode, 10)
                'lCol = liEquipCode - 1
'                If liEquipCode = 1 Then
'                    MsgBox ""
'                End If
                SetText vasSch, Trim(CStr(rs_Res.Fields.Item(14).Value)), lRow, lCol
                Select Case Trim(CStr(rs_Res.Fields.Item(15).Value))
                Case "H"
                    SetForeColor vasSch, lRow, lRow, lCol, lCol, 255, 127, 0
                Case "L"
                    SetForeColor vasSch, lRow, lRow, lCol, lCol, 0, 127, 255
                Case Else
                    SetForeColor vasSch, lRow, lRow, lCol, lCol, 0, 0, 0
                End Select
                
                Exit For
            End If
        Next liEquipCode
        
        lsID = Trim(CStr(rs_Res.Fields.Item(0).Value))
        lsType = Trim(CStr(rs_Res.Fields.Item(8).Value))
        
        rs_Res.MoveNext
    Loop
    
    rs_Res.Close
    
    Me.MousePointer = 0
    
    vasSch.MaxRows = vasSch.DataRowCnt
    vasSch.RowHeight(-1) = 12.6
    
    vasActiveCell vasSch, 1, 2
    
    frameSch.Visible = True
    'EquipCode,ExamCode, ExamName, Seqno, PointSize, RefLow, RefHigh
    
    Exit Sub
ErrHandle:
    Me.MousePointer = 0
    Exit Sub
End Sub

Private Sub cmdSchClose_Click()
    dtpExamDate.Value = Format(CDate(GetDateFull), "yyyy-mm-dd")
    frameSch.Visible = False
    'cmdPrint.Visible = False
    cmdSch.Visible = True
End Sub

Private Sub cmdSend_Click()
    Dim lRow, i, liEquipCode As Long
    Dim lsID As String
    Dim lsResult As String
    Dim lsInscode As String
    Dim lsWBC As String
    Dim lsNRBC As String
    Dim lsEOSIN As String
    Dim lsC_WBC As String
    Dim liRet, j, liMustBe As Integer
    Dim liCheck1, liCheck2 As Integer
    Dim lsSmearCode As String
    
    Dim mExam
    
    If MsgBox(" " & vbCrLf & "검사 결과를 전송하시겠습니까?" & vbCrLf & " ", vbInformation + vbYesNo + vbDefaultButton2, "결과 전송 알림") = vbNo Then
        Exit Sub
    End If
    
    For lRow = 1 To vasSch.DataRowCnt
        vasSch.Row = lRow
        vasSch.Col = 1
        
        lsWBC = ""
        lsNRBC = ""
        lsEOSIN = ""
        
        If vasSch.Value = 1 Then
            lsID = Trim(GetText(vasSch, lRow, 2))
            
'            If Trim(GetText(vasSch, lRow, 5)) = "IPU1" Then
''                lsInscode = "02"
'                lsInscode = IPU1.UseEquip
'            Else
''                lsInscode = "01"
'                lsInscode = IPU2.UseEquip
'            End If
            
            'lsInscode = "01"
            
            res = ToServer(lRow, vasSch)
            If res = 1 Then
                vasSch.Row = lRow
                vasSch.Col = 1
                vasSch.Value = 0

                SetText vasSch, "완료", lRow, gResCol
                SetBackColor vasSch, lRow, lRow, 1, 1, 202, 255, 112
                'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
                
                SQL = "update pat_res set sendflag = 'B' " & vbCrLf & _
                      "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                      "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                      "  AND barcode = '" & Trim(GetText(vasSch, lRow, 1)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then SaveQuery SQL
                
                SQL = "Update worklist set OrdFlag = 'D' where barcode = '" & Trim(GetText(vasSch, lRow, 1)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then SaveQuery SQL
                
            ElseIf res = 2 Then
                SetText vasSch, "결과", lRow, gResCol
            Else
                SetText vasSch, "실패", lRow, gResCol
                SetForeColor vasList, lRow, lRow, 1, 1, 255, 0, 0
                'SetBackColor vasID, gPreRow, gPreRow, colCheckBox, vasID.MaxCols, 255, 255, 255
                'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
            End If
        End If
    Next lRow

End Sub

Private Sub Command1_Click()
    XE2100_ASTM_1
    txtBuff1 = ""
End Sub

Private Sub Command2_Click()
    XE2100_ASTM_2
    txtBuff2 = ""
End Sub

Private Sub Command3_Click()
    giLevel = giLevel + 1
    
    txtBuff1 = SendOrder1(giLevel)
End Sub

Private Sub Command4_Click()
'    Get_QCOrder "99909170435", "1"
'
'    Exit Sub
    
    giLevel1 = giLevel1 + 1
    
    txtBuff2 = SendOrder2(giLevel1)
    
End Sub

Private Sub Command5_Click()
    Dim lsChar As String
    Dim lsSend As String
    Dim i As Integer
    
    For i = 1 To Len(Text1)
        lsChar = Mid(Text1, i, 1)
    
        Select Case lsChar
        Case chrENQ
'            SaveData "[1:RX]" & lsChar
            MSComm1.Output = chrACK
'            SaveData "[1:TX]" & chrACK
        Case chrEOT
'            SaveData "[1:RX]" & chrEOT
            
            If gMsgFlag = "Q" Then
                giLevel = 0
                
                MSComm1.Output = chrENQ
'                SaveData "[1:TX]" & chrENQ
            End If
            
        Case chrSTX
            txtBuff1.Text = ""
            
        Case chrETX
'            'SaveData "[1:RX]" & txtBuff1.Text
            If IPU1.Protocol = "ASTM" Then
    '            XE2100_ASTM
    '
    '            COM_OUTPUT1
            Else
'                SaveData "[1:RX]" & txtBuff1.Text
                
                If Mid(txtBuff1.Text, 1, 1) = "D" Then   '항상 "D"임
                    XE2100_1
                End If
                
                If IPU1.Protocol = "ClassB" Then
                    COM_OUTPUT1
                End If
            End If
            
            If IPU1.Protocol = "B" Then
                COM_OUTPUT1
    '            DoSleep 3000
    '            MSComm1.Output = Chr(6)
'    '            SaveData "[1:TX]" & chrACK
            End If
            
            'If Mid(txtBuff1.Text, 1, 1) = "D" Then   '항상 "D"임
                XE2100_1
            'End If
    
            If Mid(txtBuff1.Text, 1, 1) = "R" And Trim(gOrdMSG1_1) <> "" Then
                MSComm1.Output = gOrdMSG1_1
'                SaveData "[1:TX]" & gOrdMSG1_1
                
                gOrdMSG1_1 = ""
            End If
            
        Case chrNACK
    
        Case chrACK
'            SaveData "[1:RX]" & chrACK
            
            If IPU1.Protocol = "B" Then
                If Trim(gOrdMSG2_1) <> "" Then
                    MSComm1.Output = gOrdMSG2_1
'                    SaveData "[1:TX]" & gOrdMSG2_1
                    
                    gOrdMSG2_1 = ""
                End If
            ElseIf IPU1.Protocol = "ASTM" Then
                giLevel = giLevel + 1
                
                lsSend = SendOrder1(giLevel)
                MSComm1.Output = lsSend
'                SaveData "[1:TX]" & lsSend
            End If
        Case chrCR
        Case chrLF
            If IPU1.Protocol = "ASTM" Then
'                SaveData "[1:RX]" & txtBuff1.Text
                
                XE2100_ASTM_1
                
                COM_OUTPUT1
            End If
        Case Else
            txtBuff1.Text = txtBuff1.Text & lsChar
        End Select
    Next i
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    
    IPU1.ConnectFlag = True
    IPU2.ConnectFlag = True
    
    
    gResCol = 13
    
    'frmConnect.Show 1
    GetComSetup
    GetOPtion

    
    If IPU1.ConnectFlag Then
        MSComm1.CommPort = IPU1.ComPort
        'MSComm1.CommPort = 1
        MSComm1.Settings = IPU1.Speed & "," & IPU1.Parity & "," & IPU1.DataBit & "," & IPU1.StartBit
        If IPU1.RTSEnable = "1" Then
            MSComm1.RTSEnable = True
        Else
            MSComm1.RTSEnable = False
        End If
        If IPU1.DTREnable = "1" Then
            MSComm1.DTREnable = True
        Else
            MSComm1.DTREnable = False
        End If
        If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
        End If
        MSComm1.PortOpen = True
        
        lblIPU1Com.Caption = "[COM" & IPU1.ComPort & "]" & MSComm1.Settings
        lblIPU1.ForeColor = RGB(0, 0, 255)
    Else
        lblIPU1.ForeColor = RGB(255, 0, 0)
    End If
    
    If IPU2.ConnectFlag Then
        MSComm2.CommPort = IPU2.ComPort
        MSComm2.Settings = IPU2.Speed & "," & IPU2.Parity & "," & IPU2.DataBit & "," & IPU2.StartBit
        If IPU2.RTSEnable = "1" Then
            MSComm2.RTSEnable = True
        Else
            MSComm2.RTSEnable = False
        End If
        If IPU2.DTREnable = "1" Then
            MSComm2.DTREnable = True
        Else
            MSComm2.DTREnable = False
        End If
        MSComm2.PortOpen = True

        lblIPU2Com.Caption = "[COM" & IPU2.ComPort & "]" & MSComm2.Settings
        lblIPU2.ForeColor = RGB(0, 0, 255)
    Else
        lblIPU2.ForeColor = RGB(255, 0, 0)
    End If
    
    cn_Local_Flag = False
    
    GetSetup    'ini파일에서 정보가져오기
    
    lblUser.Caption = Trim(gIFUser)
    
    If Connect_Local Then
        cn_Local_Flag = True
    End If
    
    dtpExamDate.Value = Format(CDate(GetDateFull), "yyyy-mm-dd")
    
    If Trim(GetSetting("MEDIMATE", "XE2100", "SendMode", "0")) = "1" Then
        chkMode.Value = 1
        subSend1.Checked = True
        subSend2.Checked = False
    Else
        chkMode.Value = 0
        subSend1.Checked = False
        subSend2.Checked = True
    End If
        
    SQL = "Select WardRoom from pat_res "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "Alter table pat_res add column WardRoom varchar(20) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = "Select PBSFlag from res_flag "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "Alter table res_flag add column PBSFlag varchar(1) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = "Select SlideOrd from res_flag "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "Alter table SlideOrd add column SlideOrd varchar(2) "
        res = SendQuery(gLocal, SQL)
    End If
    
    GetExamCode
    
    txtBuff1 = ""
    txtBuff2 = ""
    
    If Not IsNumeric(gDays) Then
        gDays = 7
        
        WritePrivateProfileString "Data", "Days", gDays, App.Path & "\interface.ini"
        
    End If
    
    SQL = "Delete from worklist where recedate < '" & DateAdd("d", 0 - CInt(gDays), dtpExamDate.Value) & "' "
    SendQuery gLocal, SQL
    
    'DoSleep 500
    
    '2009.11.24 이상은 - 로그인 속도가 느리다고 하셔서 equipno 조건 추가
    SQL = "Delete from pat_res where examdate < '" & DateAdd("d", 0 - CInt(gDays), dtpExamDate.Value) & "' "
    SQL = SQL & CR & " And equipno = '" & gEquip & "' And barcode <> '' "
    SendQuery gLocal, SQL

    'DoSleep 500
    
    SQL = "Delete from res_flag where examdate < '" & DateAdd("d", 0 - CInt(gDays), dtpExamDate.Value) & "' "
    SQL = SQL & CR & " And barcode <> '' "
    SendQuery gLocal, SQL
        
    'DoSleep 500
    
    cmdCol1_Click
End Sub

Sub GetExamCode()
    Dim AdoRs_Exam As ADODB.Recordset
    Dim lCol, lCol1 As Long
    Dim i As Integer
    Dim PreEquip As String
    
    Dim sCnt As String
    
'    ReDim gArrExam(0)
'    gArrExam(0) = ""
    
    sCnt = ""
    SQL = " Select count(equipcode) From equipexam Where Equip = '" & gEquip & "' "
    res = db_select_Var(gLocal, SQL, sCnt)
    If sCnt = "" Then sCnt = "0"
    
    SQL = "SELECT EquipCode,ExamCode, ExamName, Seqno, PointSize, RefLow, RefHigh, RSGubun, ExamNo  " & vbCrLf & _
          "  From EquipExam " & vbCrLf & _
          " WHERE Equip = '" & gEquip & "' " & vbCrLf & _
          " Order by seqno "
    Set AdoRs_Exam = db_select_rs(gLocal, SQL)
    If AdoRs_Exam Is Nothing Then
        ClearSpread vasList, 1, 1
    Else
        ClearSpread vasList, 0, 1
        
        AdoRs_Exam.MoveFirst
        lCol = gResCol
        
        
        Do Until AdoRs_Exam.EOF
            If Not IsNull(AdoRs_Exam.Fields(0).Value) Then
                If IsNumeric(AdoRs_Exam.Fields(0).Value) Then
                    'If CInt(AdoRs_Exam.Fields(0).Value) <= 36 Then
                    If CInt(AdoRs_Exam.Fields(0).Value) <= CCur(sCnt) Then
                        lCol = lCol + 1
                    End If
                End If
            End If
            
'            ReDim Preserve gArrExam(lCol - gResCol, 5)
'
'            gArrExam(lCol - gResCol, 1) = AdoRs_Exam.Fields(0).Value
'            gArrExam(lCol - gResCol, 2) = AdoRs_Exam.Fields(1).Value
'            gArrExam(lCol - gResCol, 3) = AdoRs_Exam.Fields(2).Value
'            gArrExam(lCol - gResCol, 4) = AdoRs_Exam.Fields(3).Value
'            gArrExam(lCol - gResCol, 5) = AdoRs_Exam.Fields(4).Value
'
'            SetText vasList, AdoRs_Exam.Fields(2).Value, 0, lCol
'
            AdoRs_Exam.MoveNext
        Loop
        
        ReDim gArrExam(lCol - gResCol, 10)
        
        AdoRs_Exam.MoveFirst
        lCol = gResCol
        lCol1 = gResCol
        
        PreEquip = ""
        
        Do Until AdoRs_Exam.EOF
            If Not IsNull(AdoRs_Exam.Fields(0).Value) Then
                If IsNumeric(AdoRs_Exam.Fields(0).Value) Then
                    'If CInt(AdoRs_Exam.Fields(0).Value) <= 36 Then
                    If CInt(AdoRs_Exam.Fields(0).Value) <= CCur(sCnt) Then
                    
                        lCol = lCol + 1
                                                
                        'ReDim Preserve gArrExam(lCol - gResCol, 5)
                        For i = 0 To 8
                            If IsNull(AdoRs_Exam.Fields(i).Value) Then
                                gArrExam(lCol - gResCol, i + 1) = ""
                            Else
                                gArrExam(lCol - gResCol, i + 1) = AdoRs_Exam.Fields(i).Value
                            End If
                        Next i
                        
'                        If IsNumeric(AdoRs_Exam.Fields(0).Value) Then
'                            Select Case CInt(AdoRs_Exam.Fields(0).Value)
'                            Case 1
'                                flagWBC = lCol
'                            Case 2
'                                flagRBC = lCol
'                            Case 8
'                                flagPLT = lCol
'                            End Select
'                        End If
                        If PreEquip <> Trim(AdoRs_Exam.Fields(2).Value) Then
                            lCol1 = lCol1 + 1
                            SetText vasList, AdoRs_Exam.Fields(2).Value, 0, lCol1
                            'SetText vasList, lCol1, 1, lCol1
                            
                            If IsNumeric(AdoRs_Exam.Fields(0).Value) Then
                                Select Case CInt(AdoRs_Exam.Fields(0).Value)
                                Case 1
                                    flagWBC = lCol1
                                Case 2
                                    flagRBC = lCol1
                                Case 8
                                    flagPLT = lCol1
                                End Select
                            End If
                        End If
                        PreEquip = Trim(AdoRs_Exam.Fields(2).Value)
                        gArrExam(lCol - gResCol, 10) = lCol1
                    End If
                End If
            End If
           
            
            AdoRs_Exam.MoveNext
        Loop
    End If
    
'    lCol = lCol + 1
'    gMaxCol = lCol
    lCol1 = lCol1 + 1
    gMaxCol = lCol1

    vasList.MaxCols = gMaxCol + 5
    
    SetText vasList, "WBC(S)", 0, gMaxCol
    SetText vasList, "RBC(S)", 0, gMaxCol + 1
    SetText vasList, "PLT(S)", 0, gMaxCol + 2
    SetText vasList, "Sample", 0, gMaxCol + 3
    SetText vasList, "PBS", 0, gMaxCol + 5
    
    vasList.ColWidth(gMaxCol + 4) = 0      '2010.04.14 이상은 수정
    
    vasList.ColWidth(gMaxCol + 5) = 4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DisConnect_Local
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    If MSComm2.PortOpen = True Then MSComm2.PortOpen = False
    
    End
End Sub

Private Sub MSComm1_OnComm()
    Dim lsChar As String
    Dim lsSend As String
    
    lsChar = MSComm1.Input
    
'    raw_data = raw_data & lsChar
    
    Select Case lsChar
    Case chrENQ
'        SaveData "[1:RX]" & lsChar
        MSComm1.Output = chrACK
'        SaveData "[1:TX]" & chrACK
    Case chrEOT
'        SaveData "[1:RX]" & chrEOT
        
        If gMsgFlag = "Q" Then
            giLevel = 0
            
            MSComm1.Output = chrENQ
'            SaveData "[1:TX]" & chrENQ
        End If
        
    Case chrSTX
        txtBuff1.Text = ""
        
    Case chrETX
        'SaveData "[1:RX]" & txtBuff1.Text
        If IPU1.Protocol = "ASTM" Then
'            XE2100_ASTM
'
'            COM_OUTPUT1
        Else
'            SaveData "[1:RX]" & txtBuff1.Text
            
            If Mid(txtBuff1.Text, 1, 1) = "D" Then   '항상 "D"임
                XE2100_1
            End If
            
            If IPU1.Protocol = "ClassB" Then
                COM_OUTPUT1
            End If
        End If
        
        If IPU1.Protocol = "B" Then
            COM_OUTPUT1
'            DoSleep 3000
'            MSComm1.Output = Chr(6)
'            SaveData "[1:TX]" & chrACK
        End If
        
        'If Mid(txtBuff1.Text, 1, 1) = "D" Then   '항상 "D"임
            XE2100_1
        'End If

        If Mid(txtBuff1.Text, 1, 1) = "R" And Trim(gOrdMSG1_1) <> "" Then
            MSComm1.Output = gOrdMSG1_1
'            SaveData "[1:TX]" & gOrdMSG1_1
            
            gOrdMSG1_1 = ""
        End If
        
    Case chrNACK

    Case chrACK
'        SaveData "[1:RX]" & chrACK
        
        If IPU1.Protocol = "B" Then
            If Trim(gOrdMSG2_1) <> "" Then
                MSComm1.Output = gOrdMSG2_1
'                SaveData "[1:TX]" & gOrdMSG2_1
                
                gOrdMSG2_1 = ""
            End If
        ElseIf IPU1.Protocol = "ASTM" Then
            giLevel = giLevel + 1
            
            lsSend = SendOrder1(giLevel)
            MSComm1.Output = lsSend
'            SaveData "[1:TX]" & lsSend
        End If
    Case chrCR
    Case chrLF
        If IPU1.Protocol = "ASTM" Then
'            SaveData "[1:RX]" & txtBuff1.Text
            
            XE2100_ASTM_1
            
            COM_OUTPUT1
        End If
    Case Else
        txtBuff1.Text = txtBuff1.Text & lsChar
    End Select
    
End Sub

Sub COM_OUTPUT1()
'    SaveData "[1:TX]" & Chr(6)
    MSComm1.Output = Chr(6)
End Sub

Sub COM_OUTPUT2()
    MSComm2.Output = Chr(6)
End Sub

Sub XE2100_1()
    Dim myVar As String
    Dim lsTmp As String

    Dim iPoint As Integer

    Dim lsID As String
    Dim lsDate As String
    Dim lsRack As String
    Dim lsTube As String
    Dim SampleJudg, PosDif, PosMor, PosCnt, ErrFun, ErrRes As String
    Dim InfoOrd, InfoSample, InfoUnit, InfoWBC, InfoPLT As String

    Dim liEquipCode As Integer
    Dim lsResult As String

    Dim lRow As Long
    Dim lCol As Long
    Dim i, j, k As Long
    Dim z As Integer

    Dim mExam As Variant


    Dim lsInscode As String
    Dim lsWBC As String
    Dim lsNRBC As String
    Dim lsEOSIN As String
    Dim lsC_WBC As String
    Dim liEOSIN As Integer  'CP0132 Eosinophil Diff.count 처방 유무
    Dim lsSmearCode As String
    Dim lsSlideOrd As String
    Dim liMustBe As Integer

    Dim liRet As Integer

    Dim lsCurDate As String

    lsCurDate = GetDateFull

    If Trim(txtBuff1) = "" Then Exit Sub

    Select Case Left(txtBuff1, 2)
    Case "R1"   'Real-time Inquiry
        gOrdMSG1_1 = ""
        gOrdMSG2_1 = ""

        lsID = Trim(Mid(txtBuff1, 6, 15))
        lsRack = Trim(Mid(txtBuff1, 23, 6))
        lsTube = Trim(Mid(txtBuff1, 29, 2))

        OrderEntry_1 lsID, lsRack, lsTube

        'CBC & DIFF & CDC & Reserve & "Reti" & Reserve & "CBC" & NRBC & Reserve

    Case "R2"   'Batch Inquiry


    Case "S1"    'Analysis Information Format 1
    Case "S2"    'Analysis Information Format 2
    Case "D1"
        If Mid(txtBuff1, 3, 1) = "U" Then '샘플
            liMustBe = 0

            myVar = Mid(txtBuff1, 4)

            lRow = vasList.DataRowCnt + 1
            If lRow > vasList.MaxRows Then
                vasList.MaxRows = lRow
                vasList.RowHeight(lRow) = 12.6
            End If

            lsTmp = Left(myVar, 16)     'Instrument Name
            lsTmp = Mid(myVar, 17, 10)  'Sequence No
            lsID = Trim(Mid(myVar, 30, 15))   'Sample ID No.
            If InStr(1, lsID, "QC") > 0 Then
                lsID = lsID = "-1"
            End If

            'FormatA
            'lsDate = Mid(myVar, 45, 2) & "-" & Mid(myVar, 47, 2) & "-" & Mid(myVar, 49, 2) & " " & Mid(myVar, 51, 2) & "-" & Mid(myVar, 53, 2)
            'myVar = Mid(myVar, 55)
            'FormatB
            lsDate = Mid(myVar, 45, 4) & "-" & Mid(myVar, 49, 2) & "-" & Mid(myVar, 51, 2) & " " & Mid(myVar, 53, 2) & "-" & Mid(myVar, 55, 2)
            myVar = Mid(myVar, 57)

            lsRack = Mid(myVar, 3, 6)
            If IsNumeric(lsRack) Then
                lsRack = CStr(CCur(lsRack))
            End If
            lsTube = Mid(myVar, 9, 2)
            lsTmp = Mid(myVar, 13, 16) 'Patient ID
            myVar = Mid(myVar, 29)
            SampleJudg = Mid(myVar, 2, 1)
            PosDif = Mid(myVar, 3, 1)
            PosMor = Mid(myVar, 4, 1)
            PosCnt = Mid(myVar, 5, 1)
            ErrFun = Mid(myVar, 6, 1)
            ErrRes = Mid(myVar, 7, 1)
            InfoOrd = Mid(myVar, 8, 1)
            InfoSample = Mid(myVar, 9, 6)
            InfoUnit = Mid(myVar, 15, 1)
            InfoWBC = Mid(myVar, 16, 1)
            InfoPLT = Mid(myVar, 17, 1)

            SetText vasList, lsID, lRow, 2
            SetText vasList, lsRack, lRow, 11
            SetText vasList, lsTube, lRow, 12
            SetText vasList, lsID, lRow, gMaxCol + 4

            SetText vasList, Mid(InfoSample, 2, 1), lRow, gMaxCol
            SetText vasList, Mid(InfoSample, 4, 1), lRow, gMaxCol + 1
            SetText vasList, Mid(InfoSample, 6, 1), lRow, gMaxCol + 2


            SQL = "Select barcode, SlideOrd from res_flag " & vbCrLf & _
                  "where examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "'  " & vbCrLf & _
                  "  and barcode = '" & lsID & "' "
            res = db_select_Col(gLocal, SQL)
            If Trim(gReadBuf(0)) = lsID Then
                lsSlideOrd = Trim(gReadBuf(1))
            Else
                lsSlideOrd = ""
            End If

            Select Case SampleJudg
            Case "0"
                SetText vasList, "Negative", lRow, gMaxCol + 3
                SetBackColor vasList, lRow, lRow, 2, 2, 255, 255, 255
            Case "1"
                SetText vasList, "Positive", lRow, gMaxCol + 3
                SetBackColor vasList, lRow, lRow, 2, 2, 255, 224, 193
                If lsSlideOrd = "SC" Then
                    liMustBe = 1
                End If
            Case "2"
                SetText vasList, "Error", lRow, gMaxCol + 3
                SetBackColor vasList, lRow, lRow, 2, 2, 255, 224, 193
            Case "3"
                SetText vasList, "Potive+Error", lRow, gMaxCol + 3
                SetBackColor vasList, lRow, lRow, 2, 2, 255, 224, 193
                If lsSlideOrd = "SC" Then
                    liMustBe = 1
                End If
            Case "4"
                SetText vasList, "QC Sample", lRow, gMaxCol + 3
                SetBackColor vasList, lRow, lRow, 2, 2, 255, 255, 255
            End Select

            If lsSlideOrd = "SP" Then
                liMustBe = 1
            End If

            If Mid(InfoSample, 2, 1) <> "0" Then
                SetBackColor vasList, lRow, lRow, flagWBC, flagWBC, 193, 234, 255
            Else
                SetBackColor vasList, lRow, lRow, flagWBC, flagWBC, 255, 255, 255
            End If
            If Mid(InfoSample, 4, 1) <> "0" Then
                SetBackColor vasList, lRow, lRow, flagRBC, flagRBC, 193, 234, 255
            Else
                SetBackColor vasList, lRow, lRow, flagRBC, flagRBC, 255, 255, 255
            End If
            If Mid(InfoSample, 6, 1) <> "0" Then
                SetBackColor vasList, lRow, lRow, flagPLT, flagPLT, 193, 234, 255
            Else
                SetBackColor vasList, lRow, lRow, flagPLT, flagPLT, 255, 255, 255
            End If

            SetText vasList, "IPU1", lRow, 10

            vasActiveCell vasList, lRow, 2      '2009.10.29 이상은

            '환자정보
            GetPatientInfo1 lsID, lRow

            If liMustBe = 1 Then
                SetText vasList, liMustBe, lRow, gMaxCol + 5
                SetBackColor vasList, lRow, lRow, 9, 9, 255, 224, 193
                'SetForeColor vasList, lRow, lRow, 9, 9, 0, 0, 160
'                vasList.Row = lRow
'                vasList.Col = 9
'                vasList.FontBold = True
            Else
                SetText vasList, "", lRow, gMaxCol + 5
                SetBackColor vasList, lRow, lRow, 9, 9, 255, 255, 255
                'SetForeColor vasList, lRow, lRow, 9, 9, 0, 0, 0
'                vasList.Row = lRow
'                vasList.Col = 9
'                vasList.FontBold = False
            End If

            gCurRow = lRow

            SQL = "Delete from res_flag " & vbCrLf & _
                  "where examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "'  " & vbCrLf & _
                  "  and barcode = '" & lsID & "' "
            res = SendQuery(gLocal, SQL)

            SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
                  "Values ('" & Format(CDate(lsCurDate), "yyyymmdd") & "', '" & lsID & "', '" & SampleJudg & "', '" & PosDif & "', '" & PosMor & "', '" & PosCnt & "', " & _
                  "'" & ErrFun & "', '" & ErrRes & "', '" & Left(InfoSample, 1) & "', '" & Mid(InfoSample, 2, 1) & "', '" & Mid(InfoSample, 3, 1) & "', '" & Mid(InfoSample, 4, 1) & "', " & _
                  "'" & Mid(InfoSample, 5, 1) & "', '" & Mid(InfoSample, 6, 1) & "', '" & InfoWBC & "', '" & InfoPLT & "', '" & liMustBe & "', '" & lsSlideOrd & "' ) "
            res = SendQuery(gLocal, SQL)

        ElseIf Mid(txtBuff1, 3, 1) = "C" Then 'QC
            myVar = ""
        End If
    Case "D2"
        lsWBC = ""
        lsNRBC = ""
        lsEOSIN = ""
        liEOSIN = -1

        If Mid(txtBuff1, 3, 1) = "U" Then '샘플
            myVar = Mid(txtBuff1, 4)

            lsTmp = Left(myVar, 16)     'Instrument Name
            lsTmp = Mid(myVar, 17, 10)  'Sequence No
            lsID = Trim(Mid(myVar, 30, 15))   'Sample ID No.
            If InStr(1, lsID, "QC") > 0 Then
                lsID = lsID = "-1"
            End If


'            lRow = -1
'            For i = vasList.DataRowCnt To 1 Step -1
'                If Trim(GetText(vasList, i, 2)) = lsID And Trim(GetText(vasList, i, 10)) = "IPU1" Then
'                    lRow = i
'                    Exit For
'                End If
'            Next i
'
'            If lRow = -1 Then
'                lRow = vasList.DataRowCnt + 1
'                If lRow > vasList.MaxRows Then
'                    vasList.MaxRows = lRow
'                    vasList.RowHeight(lRow) = 12.6
'                End If
'            End If

            If gCurRow > 0 And gCurRow <= vasList.DataRowCnt And Trim(GetText(vasList, gCurRow, 2)) = lsID Then
                lRow = gCurRow
            Else
                lRow = vasList.DataRowCnt + 1
                If lRow > vasList.MaxRows Then
                    vasList.MaxRows = lRow
                    vasList.RowHeight(lRow) = 12.6
                End If
            End If

            If Trim(GetText(vasList, lRow, 3)) = "" Then
                SetText vasList, lsID, lRow, 2
                SetText vasList, lsID, lRow, gMaxCol + 4

                '환자정보
                GetPatientInfo1 lsID, lRow

                SetText vasList, "IPU1", lRow, 10
            End If

            vasActiveCell vasList, lRow, 2      '2009.10.29 이상은

            ClearSpread vasTemp1
            ClearSpread vasExam1

'            res = Get_Order(lsID)
'            For i = 0 To UBound(gOrder_List)
'                vasExam1.SetText 1, i + 1, gOrder_List(i).TST_CD
'                Select Case Trim(GetText(vasExam1, i, 1))
'                Case "CP0112"
'                    liEOSIN = 1
'                Case "CP0107"   'ESR
'                    Save_Local_ESR lRow, "CP0107", "99", "A", "ESR"
'                End Select
'            Next i

            res = Online_XML(gXml_S07, Trim(lsID))
            For i = 0 To UBound(gExam_Select)
                vasExam1.SetText 1, i + 1, gExam_Select(i).TST_CD
                Select Case Trim(GetText(vasExam1, i, 1))
                Case "CP0112"
                    liEOSIN = 1
                Case "CP0107"   'ESR
                    Save_Local_ESR lRow, "CP0107", "99", "A", "ESR"
                End Select
            Next i
            
            myVar = Mid(myVar, 45)


            ReDim gArrExamRes1(1 To 35)

            For liEquipCode = 1 To 35
                Select Case liEquipCode
                Case 1, 14, 15, 16, 17, 18, 31, 32, 33, 35 '(31 은 FORMAT B 일때)
                    iPoint = 6
                Case Else
                    iPoint = 5
                End Select

                lsResult = Left(myVar, iPoint)
                If InStr(1, lsResult, "*") > 0 Then
                    lsResult = "----"
                End If

                myVar = Mid(myVar, iPoint + 1)

                i = 1
                For i = 1 To UBound(gArrExam)
                    If CInt(gArrExam(i, 1)) = liEquipCode Then
                        z = -1
                        If Trim(GetText(vasList, lRow, gResCol)) = "미접수" Then
                            z = 1

                            gArrExamRes1(liEquipCode).EquipCode = liEquipCode
                            gArrExamRes1(liEquipCode).ExamCode = gArrExam(i, 2)
                            gArrExamRes1(liEquipCode).ExamNo = gArrExam(i, 9)
                            gArrExamRes1(liEquipCode).ExamName = gArrExam(i, 3)
                            gArrExamRes1(liEquipCode).SeqNo = gArrExam(i, 5)
                            gArrExamRes1(liEquipCode).RefLow = gArrExam(i, 6)
                            gArrExamRes1(liEquipCode).RefHigh = gArrExam(i, 7)
                            gArrExamRes1(liEquipCode).RefFlag = ""
                            gArrExamRes1(liEquipCode).res = lsResult
                            gArrExamRes1(liEquipCode).EquipGubun = "IPU1"

                            SetResult1 liEquipCode, i

                            SetText vasList, gArrExamRes1(liEquipCode).res, lRow, gArrExam(i, 10)
                            If gArrExamRes1(liEquipCode).RefFlag = "H" Then
                                SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 255, 127, 0
                            ElseIf gArrExamRes1(liEquipCode).RefFlag = "L" Then
                                SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 127, 255
                            Else
                                SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 0, 0
                            End If

                            Save_Local_One1 lRow, liEquipCode, "A"
                        End If

                        For k = 1 To vasExam1.DataRowCnt
                            If Trim(GetText(vasExam1, k, 1)) = gArrExam(i, 2) Then
                                z = 1

                                gArrExamRes1(liEquipCode).EquipCode = liEquipCode
                                gArrExamRes1(liEquipCode).ExamCode = gArrExam(i, 2)
                                gArrExamRes1(liEquipCode).ExamNo = gArrExam(i, 9)
                                gArrExamRes1(liEquipCode).ExamName = gArrExam(i, 3)
                                gArrExamRes1(liEquipCode).SeqNo = gArrExam(i, 5)
                                gArrExamRes1(liEquipCode).RefLow = gArrExam(i, 6)
                                gArrExamRes1(liEquipCode).RefHigh = gArrExam(i, 7)
                                gArrExamRes1(liEquipCode).RefFlag = ""
                                gArrExamRes1(liEquipCode).res = lsResult
                                gArrExamRes1(liEquipCode).EquipGubun = "IPU1"

                                SetResult1 liEquipCode, i

                                SetText vasList, gArrExamRes1(liEquipCode).res, lRow, gArrExam(i, 10)
                                If gArrExamRes1(liEquipCode).RefFlag = "H" Then
                                    SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 255, 127, 0
                                ElseIf gArrExamRes1(liEquipCode).RefFlag = "L" Then
                                    SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 127, 255
                                Else
                                    SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 0, 0
                                End If

                                Save_Local_One1 lRow, liEquipCode, "A"

                                DeleteRow vasExam1, k, k

                                Exit For
                            End If
                        Next k
                        If z = 1 Then
                            Exit For
                        End If
                    End If
                Next i


                Select Case liEquipCode
                Case 1
                    lsWBC = gArrExamRes1(liEquipCode).res
                Case 12
                    lsEOSIN = gArrExamRes1(liEquipCode).res
                Case 32
                    lsNRBC = gArrExamRes1(liEquipCode).res
                End Select


                'SetText vasList, lsResult, lRow, grescol + liEquipCode
            Next liEquipCode

'            If liEOSIN = 1 Then
'                lsEOSIN = Format(CCur(lsEOSIN) * CCur(lsWBC) * 10, "#0")
'                Save_Local_ESR lRow, "CP0112", "98", "A", "Eos.Diif.Cnt", lsEOSIN
'            End If

            vasList.Row = lRow
            vasList.Col = 1
            If vasList.Value = 0 Then
                SetText vasList, "수신", lRow, gResCol
            End If

            If lRow < 1 Or lRow > vasList.DataRowCnt Then Exit Sub

'            vasList.Row = lRow
'            vasList.Col = 1
            If chkMode.Value = 1 Then
                res = 1
                'lsInscode = "01"
                lsInscode = IPU1.UseEquip


'                SaveData lRow & " : " & lsID
                res = ToServer(lRow, vasList)
                If res = 1 Then
                    vasList.Row = lRow
                    vasList.Col = 1
                    vasList.Value = 0

                    SetText vasList, "완료", lRow, gResCol
                    SetBackColor vasList, lRow, lRow, 1, 1, 202, 255, 112
                    'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid

                    SQL = "update pat_res set sendflag = 'B' " & vbCrLf & _
                          "WHERE examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "' " & vbCrLf & _
                          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                          "  AND barcode = '" & lsID & "' "
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then SaveQuery SQL

                    SQL = "Update worklist set OrdFlag = 'D' where barcode = '" & lsID & "' "
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then SaveQuery SQL


                ElseIf res = 2 Then
                    vasList.Row = lRow
                    vasList.Col = 1
                    vasList.Value = 1

                    'SetText vasList, "결과", lRow, gResCol
                Else
                    vasList.Row = lRow
                    vasList.Col = 1
                    vasList.Value = 1

                    SQL = "Update worklist set OrdFlag = 'E' where barcode = '" & lsID & "' "
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then SaveQuery SQL

                    SQL = "update pat_res set sendflag = 'E' " & vbCrLf & _
                          "WHERE examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "' " & vbCrLf & _
                          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                          "  AND barcode = '" & lsID & "' "
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then SaveQuery SQL

                    SetText vasList, "실패", lRow, gResCol
                    SetForeColor vasList, lRow, lRow, 1, 1, 255, 0, 0
                    'SetBackColor vasID, gPreRow, gPreRow, colCheckBox, vasID.MaxCols, 255, 255, 255
                    'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
                End If
            Else
                SQL = "Update worklist set OrdFlag = 'C' where barcode = '" & lsID & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then SaveQuery SQL

            End If
        ElseIf Mid(txtBuff1, 3, 1) = "C" Then 'QC
            myVar = ""
        End If
    End Select
End Sub

Function OrderEntry_1(ByVal asID As String, ByVal asRack As String, ByVal asTube As String) As Integer
    Dim lsID As String
    Dim lsRack As String
    Dim lsTube As String
    
    Dim i, j, k As Integer
    Dim AdoRs_Exam As ADODB.Recordset
    Dim Ord(7) As String
    Dim Item(1 To 35)  As String
    Dim lsOrder As String
    Dim lRow, lRow1 As Long
    Dim lsDate As String
    
    Dim lsWkNo As String
    Dim lsPID As String
    Dim lsPName As String
    Dim lsPEName As String
    Dim lsPAge As String
    Dim lsPSex As String
    Dim lsPBirth As String
    Dim lsWard As String
    
    Dim lsSlideOrd As String
    Dim lsExamDate As String
    
    Dim lsSlideName As String
    Dim iMOR As Integer
    Dim iPBS As Integer
    
    iMOR = -1
    iPBS = -1
    
    OrderEntry_1 = -1
    
    lsExamDate = Format(CDate(GetDateFull), "yyyymmdd")
    
    SQL = "select equipcode, examcode, examname, OrdGubun, examno from equipexam where equip = '" & gEquip & "' and OrdGubun <> 'N' order by 2, 5 "
    Set AdoRs_Exam = db_select_rs(gLocal, SQL)
        
    lsSlideName = ""
        
    lsOrder = ""
    For i = 1 To 7
        Ord(i) = "0"
    Next i
    For i = 1 To 35
        Item(i) = "0"
    Next i
    
    lsSlideOrd = ""
        
    lsID = Trim(asID)
    lsRack = asRack
    lsTube = asTube
    
    If Trim(lsID) = "" Or InStr(1, lsID, "ERR") > 0 Then
        OrderEntry_1 = -1
    End If
    
    
    lsPID = ""
    lsPName = ""
    lsSlideName = ""
    lsPEName = ""
    lsPSex = ""
    lsPAge = ""
    lsWard = ""
    lsPBirth = ""
    lsPSex = ""

    
    SQL = "Select wnifsmyr || LPAD(to_char(wnifsmsn), 7, '0') || to_char(wnifsms1), " & _
          " wnifwkno, wnifidno, wnifname, WNIFRSEX, WNIFRRNF, '', wnifward" & vbCrLf & _
          "from arcwnifh a " & vbCrLf & _
          "Where wnifsmyr = '" & Mid(lsID, 1, 2) & "' " & vbCrLf & _
          "  AND wnifsmsn = " & Mid(lsID, 3, 7) & vbCrLf & _
          "  AND wnifstat <> 'X' "
                
    res = db_select_Col(gServer, SQL)
    If Trim(gReadBuf(0)) <> "" Then
    
        lsPID = Trim(gReadBuf(2))
        lsPName = Trim(gReadBuf(3))
        
        lsSlideName = Trim(gReadBuf(3))
        lsSlideName = Conv_Kor_Eng(lsSlideName)
        lsPEName = lsSlideName
        CalSexAge Trim(gReadBuf(5)) & Trim(gReadBuf(4)) & "000000", Format(Date, "yyyy/mm/dd")
        lsPSex = gPatGen.Sex
        lsPAge = gPatGen.Age
        lsWard = Trim(gReadBuf(7))
        lsPBirth = Format(CDate(gPatGen.Birth), "yyyymmdd")
        
        lsDate = GetDateFull
        
        ClearSpread vasTemp1
        
        SQL = "SELECT b.cpnwcode, c.coifabbr " & CR
        SQL = SQL & "From arccpnwh b, arcwnifh a, ABCCOIFM c" & CR
        SQL = SQL & "WHERE a.wnifdpcd = 'CP' " & vbCrLf
        SQL = SQL & "  AND a.wnifsmyr = '" & Mid(lsID, 1, 2) & "' " & CR
        SQL = SQL & "  AND a.wnifsmsn = '" & Mid(lsID, 3, 7) & "' " & CR
        SQL = SQL & "  And a.wnifstat <> 'X' " & CR
        SQL = SQL & "  and b.cpnwdpcd = a.wnifdpcd " & CR
        SQL = SQL & "  and b.cpnwdate = a.wnifdate " & CR
        SQL = SQL & "  and b.cpnwslip = a.wnifslip " & CR
        SQL = SQL & "  and b.cpnwitem = a.wnifitem " & CR
        SQL = SQL & "  and b.cpnwoitp = a.wnifoitp " & CR
        SQL = SQL & "  and b.cpnwwkno = a.wnifwkno " & CR
        'SQL = SQL & "  and b.cpnwcode In (" & sExamCode & ") " & CR
        SQL = SQL & "  and b.cpnwstat <> 'X' "
        SQL = SQL & "  and c.coifcode = b.cpnwcode  "
        res = db_select_Vas(gServer, SQL, vasTemp1)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
                
        
        k = 1
        For j = 1 To vasTemp1.DataRowCnt
            Debug.Print (GetText(vasTemp1, j, 1)) & "  " & Trim(GetText(vasTemp1, j, 2))
            If Not AdoRs_Exam Is Nothing Then
                AdoRs_Exam.MoveFirst
                Do Until AdoRs_Exam.EOF
                    Debug.Print Trim(AdoRs_Exam("examcode")) & "  " & Trim(AdoRs_Exam("examno"))
                    'If Trim(AdoRs_Exam("examcode")) = Trim(GetText(vastemp1, j, 1)) Then
                    If Trim(AdoRs_Exam("examcode")) = Trim(GetText(vasTemp1, j, 1)) Then
                        Select Case Trim(AdoRs_Exam("OrdGubun"))
                        Case "C": Ord(1) = "1"
                        Case "D": Ord(2) = "1"
                        Case "R"
                            Ord(3) = "1"
                            Ord(1) = "1"
                        Case "P"
                            Ord(4) = "1"
                            lsSlideOrd = "SP"
                        Case "S"
                            Ord(5) = "1"
                            lsSlideOrd = "SC"
                        'Case "P": Ord(4) = "1"
                        'Case "S": Ord(5) = "1"
                        Case "X": Ord(6) = "1"
                        Case "B": Ord(7) = "1"
                        End Select
                        
                        If Trim(GetText(vasTemp1, j, 1)) = "CP0106" And Trim(GetText(vasTemp1, j, 3)) = "00" Then
                            iMOR = 1
                        End If
                        If Trim(GetText(vasTemp1, j, 1)) = "CP0131" And Trim(GetText(vasTemp1, j, 3)) = "00" Then
                            iPBS = 1
                        End If
                        
                        Exit Do
                    End If
                    
                    AdoRs_Exam.MoveNext
                Loop
            End If
        Next j
        
        
    End If
    
    If iPBS = 1 Then
        lsSlideName = "PBS." & lsSlideName
    Else
        If iMOR = 1 Then
            lsSlideName = "MOR." & lsSlideName
        End If
    End If
    If Len(lsSlideName) > 40 Then
        lsSlideName = Left(lsSlideName, 40)
    End If
    
    lsOrder = ""
    For i = 1 To 7
        lsOrder = lsOrder & Ord(i)
        If Ord(i) = "1" Then
            Select Case i
            Case 1  'CBC
                For j = 1 To 8
                    Item(j) = 1
                Next j
                For j = 19 To 21
                    Item(j) = 1
                Next j
                Item(33) = 1
            Case 2  'Diff
                For j = 9 To 18
                    Item(j) = 1
                Next j
            Case 3  'Reti
                For j = 26 To 31
                    Item(j) = 1
                Next j
            Case 7
                Item(34) = 1
                Item(35) = 1
            End Select
        End If
    Next i
            
    If lsOrder = "0000000" Or lsOrder <> "" Then      'Default : CBC+Diff
        lsOrder = "11111111" & "1111111111" & _
                 "11111" & "00" & "000000" & "0" & "1" & "00" & "000000000000000"

        OrderEntry_1 = 0

    Else
        lsOrder = ""

        For i = 1 To 35
            lsOrder = lsOrder & Item(i)
        Next i
        lsOrder = lsOrder & "000000000000000"

        OrderEntry_1 = 1
    End If

    'CBC Only
'    lsOrder = "11111111" & "0000000000" & _
'             "11111" & "00" & "000000" & "0" & "1" & "00" & "000000000000000"
             
    
    Select Case lsPSex
    Case "M"
        lsPSex = "1"
    Case "F"
        lsPSex = "2"
    Case Else
        lsPSex = "3"
    End Select
    If Len(lsPBirth) <> 8 Then
        lsPBirth = Space(8)
    End If
    
    gOrdMSG1_1 = chrSTX & _
                    "S11" & lsExamDate & "000" & _
                    SetChar(Trim(lsID), 15, 1, " ") & "00" & _
                    SetChar(lsRack, 6, 1, " ") & SetChar(lsTube, 2, 1, " ") & "1" & _
                    SetChar(lsPID, 16, 1, " ") & _
                    SetChar(lsSlideName, 40, 2, " ") & _
                    lsPSex & SetSpace(lsPBirth, 8) & Space(20) & _
                    SetChar(lsWard, 20, 2, " ") & Space(40) & _
                    Space(18) & _
                    lsOrder & _
                    chrETX
                    
    gOrdMSG2_1 = chrSTX & _
                    "S21" & lsExamDate & Space(3) & _
                    SetChar(Trim(lsID), 15, 1, " ") & Space(2) & _
                    SetChar(lsRack, 6, 1, " ") & SetChar(lsTube, 2, 1, " ") & "1" & _
                    SetChar(lsPID, 16, 1, " ") & Space(100) & Space(97) & _
                    chrETX
                        
End Function

Function OrderEntry_2(ByVal asID As String, ByVal asRack As String, ByVal asTube As String) As Integer
    Dim lsID As String
    Dim lsRack As String
    Dim lsTube As String
    
    Dim i, j, k As Integer
    Dim AdoRs_Exam As ADODB.Recordset
    Dim Ord(7) As String
    Dim Item(1 To 35)  As String
    Dim lsOrder As String
    Dim lRow, lRow1 As Long
    Dim lsDate As String
    
    Dim lsWkNo As String
    Dim lsPID As String
    Dim lsPName As String
    Dim lsPEName As String
    Dim lsPAge As String
    Dim lsPSex As String
    Dim lsPBirth As String
    Dim lsWard As String
    
    Dim lsSlideOrd As String
    Dim lsExamDate As String
    
    
    Dim lsSlideName As String
    Dim iMOR As Integer
    Dim iPBS As Integer
    
    iMOR = -1
    iPBS = -1
        
    OrderEntry_2 = -1
    
    lsExamDate = Format(CDate(GetDateFull), "yyyymmdd")
    
    SQL = "select equipcode, examcode, examname, OrdGubun, examno from equipexam where equip = '" & gEquip & "' and OrdGubun <> 'N' order by 2, 5 "
    Set AdoRs_Exam = db_select_rs(gLocal, SQL)
        
    lsSlideName = ""
        
    lsOrder = ""
    For i = 1 To 7
        Ord(i) = "0"
    Next i
    For i = 1 To 35
        Item(i) = "0"
    Next i
    
    lsSlideOrd = ""
        
    lsID = Trim(asID)
    lsRack = asRack
    lsTube = asTube
    
    If Trim(lsID) = "" Or InStr(1, lsID, "ERR") > 0 Then
        OrderEntry_2 = -1
    End If
    
    '접수여부확인
    lsPID = ""
    lsPName = ""
    lsSlideName = ""
    lsPEName = ""
    lsPSex = ""
    lsPAge = ""
    lsWard = ""
    lsPBirth = ""
    lsPSex = ""
    
    
    SQL = "Select wnifsmyr || LPAD(to_char(wnifsmsn), 7, '0') || to_char(wnifsms1), " & _
          " wnifwkno, wnifidno, wnifname, WNIFRSEX, WNIFRRNF, '', wnifward" & vbCrLf & _
          "from arcwnifh a " & vbCrLf & _
          "Where wnifsmyr = '" & Mid(lsID, 1, 2) & "' " & vbCrLf & _
          "  AND wnifsmsn = " & Mid(lsID, 3, 7) & vbCrLf & _
          "  AND wnifstat <> 'X' "
                
    res = db_select_Col(gServer, SQL)
    If Trim(gReadBuf(0)) <> "" Then
    
        lsPID = Trim(gReadBuf(2))
        lsPName = Trim(gReadBuf(3))
        
        lsSlideName = Trim(gReadBuf(3))
        lsSlideName = Conv_Kor_Eng(lsSlideName)
        lsPEName = lsSlideName
        CalSexAge Trim(gReadBuf(5)) & Trim(gReadBuf(4)) & "000000", Format(Date, "yyyy/mm/dd")
        lsPSex = gPatGen.Sex
        lsPAge = gPatGen.Age
        lsWard = Trim(gReadBuf(7))
        lsPBirth = Format(CDate(gPatGen.Birth), "yyyymmdd")
        
        lsDate = GetDateFull
        
        ClearSpread vasTemp1
        
        SQL = "SELECT b.cpnwcode, c.coifabbr " & CR
        SQL = SQL & "From arccpnwh b, arcwnifh a, ABCCOIFM c" & CR
        SQL = SQL & "WHERE a.wnifdpcd = 'CP' " & vbCrLf
        SQL = SQL & "  AND a.wnifsmyr = '" & Mid(lsID, 1, 2) & "' " & CR
        SQL = SQL & "  AND a.wnifsmsn = '" & Mid(lsID, 3, 7) & "' " & CR
        SQL = SQL & "  And a.wnifstat <> 'X' " & CR
        SQL = SQL & "  and b.cpnwdpcd = a.wnifdpcd " & CR
        SQL = SQL & "  and b.cpnwdate = a.wnifdate " & CR
        SQL = SQL & "  and b.cpnwslip = a.wnifslip " & CR
        SQL = SQL & "  and b.cpnwitem = a.wnifitem " & CR
        SQL = SQL & "  and b.cpnwoitp = a.wnifoitp " & CR
        SQL = SQL & "  and b.cpnwwkno = a.wnifwkno " & CR
        'SQL = SQL & "  and b.cpnwcode In (" & sExamCode & ") " & CR
        SQL = SQL & "  and b.cpnwstat <> 'X' "
        SQL = SQL & "  and c.coifcode = b.cpnwcode  "
        res = db_select_Vas(gServer, SQL, vasTemp2)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
                
        
        k = 1
        For j = 1 To vasTemp2.DataRowCnt
            Debug.Print (GetText(vasTemp2, j, 1)) & "  " & Trim(GetText(vasTemp2, j, 2))
            If Not AdoRs_Exam Is Nothing Then
                AdoRs_Exam.MoveFirst
                Do Until AdoRs_Exam.EOF
                    Debug.Print Trim(AdoRs_Exam("examcode")) & "  " & Trim(AdoRs_Exam("examno"))
                    'If Trim(AdoRs_Exam("examcode")) = Trim(GetText(vastemp2, j, 1)) Then
                    If Trim(AdoRs_Exam("examcode")) = Trim(GetText(vasTemp2, j, 1)) And _
                       Trim(AdoRs_Exam("examno")) = Trim(GetText(vasTemp2, j, 2)) Then
                        Select Case Trim(AdoRs_Exam("OrdGubun"))
                        Case "C": Ord(1) = "1"
                        Case "D": Ord(2) = "1"
                        Case "R"
                            Ord(3) = "1"
                            Ord(1) = "1"
                        Case "P"
                            Ord(4) = "1"
                            lsSlideOrd = "SP"
                        Case "S"
                            Ord(5) = "1"
                            lsSlideOrd = "SC"
                        'Case "P": Ord(4) = "1"
                        'Case "S": Ord(5) = "1"
                        Case "X": Ord(6) = "1"
                        Case "B": Ord(7) = "1"
                        End Select
                        
                        If Trim(GetText(vasTemp2, j, 1)) = "CP0106" And Trim(GetText(vasTemp2, j, 3)) = "00" Then
                            iMOR = 1
                        End If
                        If Trim(GetText(vasTemp2, j, 1)) = "CP0131" And Trim(GetText(vasTemp2, j, 3)) = "00" Then
                            iPBS = 1
                        End If
                        
                        Exit Do
                    End If
                    
                    AdoRs_Exam.MoveNext
                Loop
            End If
        Next j
        
        
    End If
    
    If iPBS = 1 Then
        lsSlideName = "PBS." & lsSlideName
    Else
        If iMOR = 1 Then
            lsSlideName = "MOR." & lsSlideName
        End If
    End If
    If Len(lsSlideName) > 40 Then
        lsSlideName = Left(lsSlideName, 40)
    End If
    
    lsOrder = ""
    For i = 1 To 7
        lsOrder = lsOrder & Ord(i)
        If Ord(i) = "1" Then
            Select Case i
            Case 1  'CBC
                For j = 1 To 8
                    Item(j) = 1
                Next j
                For j = 19 To 21
                    Item(j) = 1
                Next j
                Item(33) = 1
            Case 2  'Diff
                For j = 9 To 18
                    Item(j) = 1
                Next j
            Case 3  'Reti
                For j = 26 To 31
                    Item(j) = 1
                Next j
            Case 7
                Item(34) = 1
                Item(35) = 1
            End Select
        End If
    Next i
            
    If lsOrder = "0000000" Or lsOrder <> "" Then       'Default : CBC+Diff
        lsOrder = "11111111" & "1111111111" & _
                 "11111" & "00" & "000000" & "0" & "1" & "00" & "000000000000000"
                 
        OrderEntry_2 = 0
        
    Else
        lsOrder = ""
        
        For i = 1 To 35
            lsOrder = lsOrder & Item(i)
        Next i
        lsOrder = lsOrder & "000000000000000"
        
        OrderEntry_2 = 1
    End If

    Select Case lsPSex
    Case "M"
        lsPSex = "1"
    Case "F"
        lsPSex = "2"
    Case Else
        lsPSex = "3"
    End Select
    If Len(lsPBirth) <> 8 Then
        lsPBirth = Space(8)
    End If
    
    gOrdMSG1_2 = chrSTX & _
                    "S11" & lsExamDate & "000" & _
                    SetChar(Trim(lsID), 15, 1, " ") & "00" & _
                    SetChar(lsRack, 6, 1, " ") & SetChar(lsTube, 2, 1, " ") & "1" & _
                    SetChar(lsPID, 16, 1, " ") & _
                    SetChar(lsSlideName, 40, 2, " ") & _
                    lsPSex & SetSpace(lsPBirth, 8) & Space(20) & _
                    SetChar(lsWard, 20, 2, " ") & Space(40) & _
                    Space(18) & _
                    lsOrder & _
                    chrETX
                    
    gOrdMSG2_2 = chrSTX & _
                    "S21" & lsExamDate & Space(3) & _
                    SetChar(Trim(lsID), 15, 1, " ") & Space(2) & _
                    SetChar(lsRack, 6, 1, " ") & SetChar(lsTube, 2, 1, " ") & "1" & _
                    SetChar(lsPID, 16, 1, " ") & Space(100) & Space(97) & _
                    chrETX
                        
End Function

Sub SetResult1(ByVal aiRow As Integer, ByVal aiItem As Integer)
    Dim iFloat As Integer
    Dim sFormat As String
    
    gArrExamRes1(aiRow).RefFlag = ""
    
    If Not IsNumeric(gArrExamRes1(aiRow).res) Then
        Exit Sub
    End If

    iFloat = gArrExam(aiItem, 5)
    
    If IPU1.Protocol <> "ASTM" Then
        If iFloat = 0 Then
            gArrExamRes1(aiRow).res = CStr(CCur(gArrExamRes1(aiRow).res))
        Else
            gArrExamRes1(aiRow).res = CCur(CStr(CCur(Left(gArrExamRes1(aiRow).res, Len(gArrExamRes1(aiRow).res) - iFloat)) & "." & Right(gArrExamRes1(aiRow).res, iFloat)))
            'If aiItem = 1 Or aiItem = 14 Or aiItem = 15 Or aiItem = 16 Or aiItem = 17 Or aiItem = 18 Then
            '    gArrExamRes1(aiRow).Res = CStr(CCur(Left(gArrExamRes1(aiRow).Res, 5 - iFloat)) & "." & Right(gArrExamRes1(aiRow).Res, iFloat))
            'Else
            '    gArrExamRes1(aiRow).Res = CStr(CCur(Left(gArrExamRes1(aiRow).Res, 4 - iFloat)) & "." & Right(gArrExamRes1(aiRow).Res, iFloat))
            'End If
        End If
    End If
    
    If IsNumeric(gArrExamRes1(aiRow).res) Then
        If IsNumeric(gArrExam(aiItem, 6)) Then
            If CCur(gArrExam(aiItem, 6)) > CCur(gArrExamRes1(aiRow).res) Then
                gArrExamRes1(aiRow).RefFlag = "L"
            End If
        End If
        If IsNumeric(gArrExam(aiItem, 7)) Then
            If CCur(gArrExam(aiItem, 7)) < CCur(gArrExamRes1(aiRow).res) Then
                gArrExamRes1(aiRow).RefFlag = "H"
            End If
        End If
    End If

    iFloat = gArrExam(aiItem, 8)
    If IsNumeric(iFloat) Then
        If CInt(iFloat) = 0 Then
            sFormat = "#0"
        ElseIf CInt(iFloat) > 0 Then
            sFormat = ""
            sFormat = SetChar(sFormat, CInt(iFloat), 1, "0")
            sFormat = "0." & sFormat
        End If
        If IsNumeric(gArrExamRes1(aiRow).res) Then
            gArrExamRes1(aiRow).res = Format(CCur(gArrExamRes1(aiRow).res), sFormat)
        End If
    End If
    
End Sub

Sub SetResult2(ByVal aiRow As Integer, ByVal aiItem As Integer)
    Dim iFloat As Integer
    Dim sFormat As String
    
    gArrExamRes2(aiRow).RefFlag = ""
    
    If Not IsNumeric(gArrExamRes2(aiRow).res) Then
        Exit Sub
    End If

    iFloat = gArrExam(aiItem, 5)
    
    If IPU2.Protocol <> "ASTM" Then
        If iFloat = 0 Then
            gArrExamRes2(aiRow).res = CStr(CCur(gArrExamRes2(aiRow).res))
        Else
            gArrExamRes2(aiRow).res = CCur(CStr(CCur(Left(gArrExamRes2(aiRow).res, Len(gArrExamRes2(aiRow).res) - iFloat)) & "." & Right(gArrExamRes2(aiRow).res, iFloat)))
            'If aiItem = 1 Or aiItem = 14 Or aiItem = 15 Or aiItem = 16 Or aiItem = 17 Or aiItem = 18 Then
            '    gArrExamRes2(aiRow).Res = CStr(CCur(Left(gArrExamRes2(aiRow).Res, 5 - iFloat)) & "." & Right(gArrExamRes2(aiRow).Res, iFloat))
            'Else
            '    gArrExamRes2(aiRow).Res = CStr(CCur(Left(gArrExamRes2(aiRow).Res, 4 - iFloat)) & "." & Right(gArrExamRes2(aiRow).Res, iFloat))
            'End If
        End If
    End If
    If IsNumeric(gArrExamRes2(aiRow).res) Then
        If IsNumeric(gArrExam(aiItem, 6)) Then
            If CCur(gArrExam(aiItem, 6)) > CCur(gArrExamRes2(aiRow).res) Then
                gArrExamRes2(aiRow).RefFlag = "L"
            End If
        End If
        If IsNumeric(gArrExam(aiItem, 7)) Then
            If CCur(gArrExam(aiItem, 7)) < CCur(gArrExamRes2(aiRow).res) Then
                gArrExamRes2(aiRow).RefFlag = "H"
            End If
        End If
    End If

    iFloat = gArrExam(aiItem, 8)
    If IsNumeric(iFloat) Then
        If CInt(iFloat) = 0 Then
            sFormat = "#0"
        ElseIf CInt(iFloat) > 0 Then
            sFormat = ""
            sFormat = SetChar(sFormat, CInt(iFloat), 1, "0")
            sFormat = "0." & sFormat
        End If
        If IsNumeric(gArrExamRes2(aiRow).res) Then
            gArrExamRes2(aiRow).res = Format(CCur(gArrExamRes2(aiRow).res), sFormat)
        End If
    End If
    
End Sub

Function Save_Local_One1(ByVal asRow As Long, ByVal aiIndex As Integer, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = GetDateFull
    
'    SQL = "select examno from pat_res "
'    res = db_select_Col(gLocal, SQL)
'    If res = -1 Then
'        SQL = "Alter table pat_res add examno varchar(10) "
'        res = SendQuery(gLocal, SQL)
'    End If
    
    sCnt = ""
    SQL = "Delete FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(sExamDate), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND equipcode = '" & gArrExamRes1(aiIndex).EquipCode & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasList, asRow, 2)) & "' " '& vbCrLf & _
          "  and examtype = '" & Trim(GetText(vasList, asRow, 10)) & "' "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = "INSERT INTO pat_res (examdate, equipno, " & _
            "barcode, examtype, " & _
            "receno, pid, " & _
            "pname, pjumin, page, " & _
            "psex, page1, " & _
            "WardRoom, resdate, seqno, " & _
            "diskno, posno, " & _
            "equipcode, examcode, examno, " & _
            "result, sendflag, examname, " & _
            "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
          "VALUES ('" & Format(CDate(sExamDate), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
          "'" & Trim(GetText(vasList, asRow, 2)) & "','" & Trim(GetText(vasList, asRow, 10)) & "', " & _
          "'" & Trim(GetText(vasList, asRow, 9)) & "', '" & Trim(GetText(vasList, asRow, 3)) & "', " & _
          "'" & Trim(GetText(vasList, asRow, 4)) & "', '" & Trim(GetText(vasList, asRow, 5)) & "', 0, " & _
          "'" & Trim(GetText(vasList, asRow, 6)) & "', '" & Trim(GetText(vasList, asRow, 7)) & "'," & _
          "'" & Trim(GetText(vasList, asRow, 8)) & "', '" & sExamDate & "', '" & gArrExamRes1(aiIndex).SeqNo & "', " & _
          "'" & Trim(GetText(vasList, asRow, 11)) & "', '" & Trim(GetText(vasList, asRow, 12)) & "', " & vbCrLf & _
          "'" & gArrExamRes1(aiIndex).EquipCode & "', '" & gArrExamRes1(aiIndex).ExamCode & "', '" & gArrExamRes1(aiIndex).ExamNo & "', " & _
          "'" & gArrExamRes1(aiIndex).res & "', '" & asSend & "', '" & gArrExamRes1(aiIndex).ExamName & "', " & vbCrLf & _
          "'" & gArrExamRes1(aiIndex).RefFlag & "', '', '', '', " & _
          "'" & gArrExamRes1(aiIndex).RefLow & " ~ " & gArrExamRes1(aiIndex).RefHigh & "', '' ) "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Function Save_Local_One2(ByVal asRow As Long, ByVal aiIndex As Integer, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = GetDateFull
    
    sCnt = ""
    SQL = "Delete FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(sExamDate), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND equipcode = '" & gArrExamRes2(aiIndex).EquipCode & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasList, asRow, 2)) & "' " '& vbCrLf & _
          "  and examtype = '" & Trim(GetText(vasList, asRow, 5)) & "' "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = "INSERT INTO pat_res (examdate, equipno, " & _
            "barcode, examtype, " & _
            "receno, pid, " & _
            "pname, pjumin, page, " & _
            "psex, page1, " & _
            "WardRoom, resdate, seqno, " & _
            "diskno, posno, " & _
            "equipcode, examcode, " & _
            "result, sendflag, examname, " & _
            "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
          "VALUES ('" & Format(CDate(sExamDate), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
          "'" & Trim(GetText(vasList, asRow, 2)) & "','" & Trim(GetText(vasList, asRow, 10)) & "', " & _
          "'" & Trim(GetText(vasList, asRow, 9)) & "', '" & Trim(GetText(vasList, asRow, 3)) & "', " & _
          "'" & Trim(GetText(vasList, asRow, 4)) & "', '" & Trim(GetText(vasList, asRow, 5)) & "', 0, " & _
          "'" & Trim(GetText(vasList, asRow, 6)) & "', '" & Trim(GetText(vasList, asRow, 7)) & "'," & _
          "'" & Trim(GetText(vasList, asRow, 8)) & "', '" & sExamDate & "', '" & gArrExamRes2(aiIndex).SeqNo & "', " & _
          "'" & Trim(GetText(vasList, asRow, 11)) & "', '" & Trim(GetText(vasList, asRow, 12)) & "', " & vbCrLf & _
          "'" & gArrExamRes2(aiIndex).EquipCode & "', '" & gArrExamRes2(aiIndex).ExamCode & "', " & _
          "'" & gArrExamRes2(aiIndex).res & "', '" & asSend & "', '" & gArrExamRes2(aiIndex).ExamName & "', " & vbCrLf & _
          "'" & gArrExamRes2(aiIndex).RefFlag & "', '', '', '', " & _
          "'" & gArrExamRes2(aiIndex).RefLow & " ~ " & gArrExamRes2(aiIndex).RefHigh & "', '' ) "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Function Save_Local_One_1(ByVal asRow As Long, ByVal aiIndex As Integer, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = GetDateFull
    
    sCnt = ""
    SQL = "Delete FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(sExamDate), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND equipcode = '" & gArrExam(aiIndex, 1) & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasList, asRow, 2)) & "' "
    
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = "INSERT INTO pat_res (examdate, equipno, barcode, examtype, receno, pid, " & _
          "pname, pjumin, page, psex, resdate, seqno, diskno, posno, " & _
          "equipcode, examcode, examtype, result, sendflag, examname, " & _
          "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
          "VALUES ('" & Format(CDate(sExamDate), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
          "'" & Trim(GetText(vasList, asRow, 2)) & "','" & Trim(GetText(vasList, asRow, 5)) & "', '', " & _
          "'" & Trim(GetText(vasList, asRow, 3)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasList, asRow, 4)) & "', '', " & _
          "0, '', " & _
          "'" & sExamDate & "', '" & gArrExam(aiIndex, 4) & "', '" & Trim(GetText(vasList, asRow, 6)) & "', '" & Trim(GetText(vasList, asRow, 7)) & "', " & vbCrLf & _
          "'" & gArrExam(aiIndex, 1) & "', '" & gArrExam(aiIndex, 2) & "', '', " & _
          "'" & Trim(GetText(vasList, asRow, gResCol + aiIndex)) & "', '" & asSend & "', '" & gArrExam(aiIndex, 3) & "', " & vbCrLf & _
          "'', '', " & _
          "'', '', " & _
          "'', '') "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Function Save_Local_ESR(ByVal asRow As Long, asExamCode As String, asEquipCode As String, asSend As String, Optional asExamName As String = "", Optional asRes As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = GetDateFull
    
    sCnt = ""
    SQL = "Delete FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(sExamDate), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND examcode = '" & asExamCode & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasList, asRow, 2)) & "' " & vbCrLf & _
          "  and examtype = '" & Trim(GetText(vasList, asRow, 5)) & "' "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = "INSERT INTO pat_res (examdate, equipno, " & _
            "barcode, examtype, " & _
            "receno, pid, " & _
            "pname, pjumin, page, " & _
            "psex, page1, " & _
            "WardRoom, resdate, seqno, " & _
            "diskno, posno, " & _
            "equipcode, examcode, " & _
            "result, sendflag, examname, " & _
            "refflag,panicflag, deltaflag, unit, refvalue, panicvalue, examno ) " & vbCrLf & _
          "VALUES ('" & Format(CDate(sExamDate), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
          "'" & Trim(GetText(vasList, asRow, 2)) & "','" & Trim(GetText(vasList, asRow, 10)) & "', " & _
          "'" & Trim(GetText(vasList, asRow, 9)) & "', '" & Trim(GetText(vasList, asRow, 3)) & "', " & _
          "'" & Trim(GetText(vasList, asRow, 4)) & "', '" & Trim(GetText(vasList, asRow, 5)) & "', 0, " & _
          "'" & Trim(GetText(vasList, asRow, 6)) & "', '" & Trim(GetText(vasList, asRow, 7)) & "'," & _
          "'" & Trim(GetText(vasList, asRow, 8)) & "', '" & sExamDate & "', '99', " & _
          "'" & Trim(GetText(vasList, asRow, 11)) & "', '" & Trim(GetText(vasList, asRow, 12)) & "', " & vbCrLf & _
          "'" & asEquipCode & "', '" & asExamCode & "', " & _
          "'" & asRes & "', '" & asSend & "', '" & asExamName & "', " & vbCrLf & _
          "'', '', '', '', " & _
          "'', '', '00' ) "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function


Function Update_Sample(ByVal asID As String)
    Dim sExamDate As String
    
    sExamDate = GetDateFull
    
    SQL = "Update pat_res set sendflag = 'B' " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(sExamDate), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND barcode = '" & asID & "' "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function

Function DeleteWorkList(ByVal asID As String)
    SQL = "Delete from WorkList where Barcode ='" & asID & "'"
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function

Public Function Set_EqpResultsql1(ByVal Testcd As String, ByVal EqpRst As String, ByVal ErrDes As String, ByVal SPCID As String, ByVal INS_CODE As String) As Boolean
On Error GoTo errtrap
    'Set cmdSQL = New ADODB.Command
    Dim sDate As String
    Dim sRes As String
    
    sRes = EqpRst
    If Not IsNumeric(sRes) Then
        If sRes <> "pbs" Then
            sRes = "*"
        End If
    End If
    
'    SaveData "InterfaceResult_INSERT_sp " & SPCID & "," & Trim(Testcd) & "," & Left(Trim(sDate), 19) & ", " & Trim(sRes) & "," & Trim(INS_CODE) & "," & Trim(ErrDes)

    sDate = GetDateFull
    
    'DoSleep CLng(DelaySP)
    
    With cmdSQL
        .ActiveConnection = cn_Ser
        .CommandType = adCmdStoredProc
        .CommandText = "InterfaceResult_INSERT_sp"
        .Parameters.Append .CreateParameter("retval", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@i_barcodeNumber", adChar, adParamInput, 11, Trim(SPCID))
        .Parameters.Append .CreateParameter("@i_itemCode", adVarChar, adParamInput, 10, Trim(Testcd))
        .Parameters.Append .CreateParameter("@i_transTimestamp", adChar, adParamInput, 19, Trim(Left(sDate, 19)))
'        .Parameters.Append .CreateParameter("@i_itemResultValue", adVarChar, adParamInput, 1000, Trim(EqpRst))
        .Parameters.Append .CreateParameter("@i_itemResultValue", adVarChar, adParamInput, 1000, Trim(sRes))
        .Parameters.Append .CreateParameter("@i_instrumentCode", adChar, adParamInput, 2, Trim(INS_CODE))
        .Parameters.Append .CreateParameter("@i_errorDescription", adVarChar, adParamInput, 100, Trim(ErrDes))
        
        .Execute
    End With
    
    If cmdSQL("retval").Value = 2 Then
        Set_EqpResultsql1 = False
        MsgBox "결과전송 실패", vbInformation, "알림"
        'Set cmdSQL = Nothing
        Exit Function
    End If
    
    Set_EqpResultsql1 = True
    'Set cmdSQL = Nothing
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing

    Exit Function
    
errtrap:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing
    'Err.Raise Err.Number, Err.Description
End Function

Public Function Set_EqpResultsql2(ByVal Testcd As String, ByVal EqpRst As String, ByVal ErrDes As String, ByVal SPCID As String, ByVal INS_CODE As String) As Boolean
On Error GoTo errtrap
    'Set cmdSQL = New ADODB.Command
    Dim sDate As String
    Dim sRes As String
    
    sRes = EqpRst
    If Not IsNumeric(sRes) Then
        If sRes <> "pbs" Then
            sRes = "*"
        End If
    End If
    
'    SaveData "InterfaceResult_INSERT_sp " & SPCID & "," & Trim(Testcd) & "," & Left(Trim(sDate), 19) & ", " & Trim(sRes) & "," & Trim(INS_CODE) & "," & Trim(ErrDes)

    sDate = GetDateFull
    
    'DoSleep CLng(DelaySP)
    
    With cmdSQL
        .ActiveConnection = cn_Ser
        .CommandType = adCmdStoredProc
        .CommandText = "InterfaceResult_INSERT_sp"
        .Parameters.Append .CreateParameter("retval", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@i_barcodeNumber", adChar, adParamInput, 11, Trim(SPCID))
        .Parameters.Append .CreateParameter("@i_itemCode", adVarChar, adParamInput, 10, Trim(Testcd))
        .Parameters.Append .CreateParameter("@i_transTimestamp", adChar, adParamInput, 19, Trim(Left(sDate, 19)))
'        .Parameters.Append .CreateParameter("@i_itemResultValue", adVarChar, adParamInput, 1000, Trim(EqpRst))
        .Parameters.Append .CreateParameter("@i_itemResultValue", adVarChar, adParamInput, 1000, Trim(sRes))
        .Parameters.Append .CreateParameter("@i_instrumentCode", adChar, adParamInput, 2, Trim(INS_CODE))
        .Parameters.Append .CreateParameter("@i_errorDescription", adVarChar, adParamInput, 100, Trim(ErrDes))
        
        .Execute
    End With
    
    If cmdSQL("retval").Value = 2 Then
        Set_EqpResultsql2 = False
        MsgBox "결과전송 실패", vbInformation, "알림"
        'Set cmdSQL = Nothing
        Exit Function
    End If
    
    Set_EqpResultsql2 = True
    'Set cmdSQL = Nothing
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing

    Exit Function
    
errtrap:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing
    'Err.Raise Err.Number, Err.Description
End Function

Public Function Set_EqpResultsql(ByVal Testcd As String, ByVal EqpRst As String, ByVal ErrDes As String, ByVal SPCID As String, ByVal INS_CODE As String) As Boolean
On Error GoTo errtrap
    'Set cmdSQL = New ADODB.Command
    Dim sDate As String
    Dim sRes As String
    
    sRes = EqpRst
    If Not IsNumeric(sRes) Then
        If sRes <> "pbs" Then
            sRes = "*"
        End If
    End If
'    SaveData "InterfaceResult_INSERT_sp " & SPCID & "," & Trim(Testcd) & "," & Left(Trim(sDate), 19) & ", " & Trim(sRes) & "," & Trim(INS_CODE) & "," & Trim(ErrDes)

    sDate = GetDateFull
    
    'DoSleep CLng(DelaySP)
    
    With cmdSQL
        .ActiveConnection = cn_Ser
        .CommandType = adCmdStoredProc
        .CommandText = "InterfaceResult_INSERT_sp"
        .Parameters.Append .CreateParameter("retval", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@i_barcodeNumber", adChar, adParamInput, 11, Trim(SPCID))
        .Parameters.Append .CreateParameter("@i_itemCode", adVarChar, adParamInput, 10, Trim(Testcd))
        .Parameters.Append .CreateParameter("@i_transTimestamp", adChar, adParamInput, 19, Trim(Left(sDate, 19)))
'        .Parameters.Append .CreateParameter("@i_itemResultValue", adVarChar, adParamInput, 1000, Trim(EqpRst))
        .Parameters.Append .CreateParameter("@i_itemResultValue", adVarChar, adParamInput, 1000, Trim(sRes))
        .Parameters.Append .CreateParameter("@i_instrumentCode", adChar, adParamInput, 2, Trim(INS_CODE))
        .Parameters.Append .CreateParameter("@i_errorDescription", adVarChar, adParamInput, 100, Trim(ErrDes))
        
        .Execute
    End With
    
    If cmdSQL("retval").Value = 2 Then
        Set_EqpResultsql = False
        MsgBox "결과전송 실패", vbInformation, "알림"
        'Set cmdSQL = Nothing
        Exit Function
    End If
    
    Set_EqpResultsql = True
    'Set cmdSQL = Nothing
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing

    Exit Function
    
errtrap:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing
    'Err.Raise Err.Number, Err.Description
End Function

Private Sub MSComm2_OnComm()
    Dim lsChar As String
    Dim lsSend   As String
    
    lsChar = MSComm2.Input
    
'    raw_data = raw_data & lsChar
    
    Select Case lsChar
    Case chrENQ
'        SaveData "[2:RX]" & lsChar
        MSComm2.Output = chrACK
'        SaveData "[2:TX]" & chrACK
    Case chrEOT
'        SaveData "[2:RX]" & chrEOT
        
        If gMsgFlag1 = "Q" Then
            giLevel1 = 0
            
            MSComm2.Output = chrENQ
'            SaveData "[2:TX]" & chrENQ
        End If
    Case chrSTX
        txtBuff2.Text = ""
        
    Case chrETX
        If IPU2.Protocol = "ASTM" Then
'            XE2100_ASTM
'            COM_OUTPUT1
        Else
'            SaveData "[2:RX]" & txtBuff2.Text
            
            If Mid(txtBuff2.Text, 1, 1) = "D" Then   '항상 "D"임
                XE2100_2
            End If
            
            If IPU2.Protocol = "ClassB" Then
                COM_OUTPUT2
            End If
            
            If Mid(txtBuff2.Text, 1, 1) = "R" And Trim(gOrdMSG1_2) <> "" Then
                MSComm2.Output = gOrdMSG1_2
'                SaveData "[2:TX]" & gOrdMSG1_2
                gOrdMSG1_2 = ""
            End If
        
        End If
        
    Case chrNACK

    Case chrACK
'        SaveData "[2:RX]" & chrACK
        If IPU1.Protocol = "B" Then
            If Trim(gOrdMSG2_2) <> "" Then
                MSComm2.Output = gOrdMSG2_2
'                SaveData "[2:TX]" & gOrdMSG2_2
                gOrdMSG2_2 = ""
            End If
        ElseIf IPU1.Protocol = "ASTM" Then
            giLevel1 = giLevel1 + 1
            
            lsSend = SendOrder2(giLevel1)
            MSComm2.Output = lsSend
'            SaveData "[2:TX]" & lsSend
        End If
        
    Case chrCR
    Case chrLF
        If IPU2.Protocol = "ASTM" Then
'            SaveData "[2:RX]" & txtBuff2.Text
            
            XE2100_ASTM_2
            
            COM_OUTPUT2
        End If
    Case Else
        txtBuff2.Text = txtBuff2.Text & lsChar
    End Select
End Sub

Sub XE2100_2()
    Dim myVar As String
    Dim lsTmp As String

    Dim iPoint As Integer

    Dim lsID As String
    Dim lsDate As String
    Dim lsRack As String
    Dim lsTube As String
    Dim SampleJudg, PosDif, PosMor, PosCnt, ErrFun, ErrRes As String
    Dim InfoOrd, InfoSample, InfoUnit, InfoWBC, InfoPLT As String

    Dim liEquipCode As Integer
    Dim lsResult As String

    Dim lRow As Long
    Dim lCol As Long
    Dim i, j, k As Long
    Dim z As Integer

    Dim mExam As Variant


    Dim lsInscode As String
    Dim lsWBC As String
    Dim lsNRBC As String
    Dim lsEOSIN As String
    Dim lsC_WBC As String
    Dim liEOSIN As Integer  'CP0132 Eosinophil Diff.count 처방 유무
    Dim lsSmearCode As String
    Dim lsSlideOrd As String
    Dim liMustBe As Integer

    Dim liRet As Integer

    Dim lsCurDate As String

    lsCurDate = GetDateFull

    If Trim(txtBuff2) = "" Then Exit Sub

    Select Case Left(txtBuff2, 2)
    Case "R1"   'Real-time Inquiry
        gOrdMSG1_2 = ""
        gOrdMSG2_2 = ""
        lsID = Trim(Mid(txtBuff2, 6, 15))
        lsRack = Trim(Mid(txtBuff2, 23, 6))
        lsTube = Trim(Mid(txtBuff2, 29, 2))

        OrderEntry_2 lsID, lsRack, lsTube

    Case "R2"   'Batch Inquiry
    Case "S1"    'Analysis Information Format 1
    Case "S2"    'Analysis Information Format 2
    Case "D1"
        If Mid(txtBuff2, 3, 1) = "U" Then '샘플
            liMustBe = 0

            myVar = Mid(txtBuff2, 4)

            lRow = vasList.DataRowCnt + 1
            If lRow > vasList.MaxRows Then
                vasList.MaxRows = lRow
                vasList.RowHeight(lRow) = 12.6
            End If

            lsTmp = Left(myVar, 16)     'Instrument Name
            lsTmp = Mid(myVar, 17, 10)  'Sequence No
            lsID = Trim(Mid(myVar, 30, 15))   'Sample ID No.
            If InStr(1, lsID, "QC") > 0 Then
                lsID = lsID = "-2"
            End If

            'FormatA
            'lsDate = Mid(myVar, 45, 2) & "-" & Mid(myVar, 47, 2) & "-" & Mid(myVar, 49, 2) & " " & Mid(myVar, 51, 2) & "-" & Mid(myVar, 53, 2)
            'myVar = Mid(myVar, 55)
            'FormatB
            lsDate = Mid(myVar, 45, 4) & "-" & Mid(myVar, 49, 2) & "-" & Mid(myVar, 51, 2) & " " & Mid(myVar, 53, 2) & "-" & Mid(myVar, 55, 2)
            myVar = Mid(myVar, 57)

            lsRack = Mid(myVar, 3, 6)
            If IsNumeric(lsRack) Then
                lsRack = CStr(CCur(lsRack))
            End If
            lsTube = Mid(myVar, 9, 2)
            lsTmp = Mid(myVar, 13, 16) 'Patient ID
            myVar = Mid(myVar, 29)
            SampleJudg = Mid(myVar, 2, 1)
            PosDif = Mid(myVar, 3, 1)
            PosMor = Mid(myVar, 4, 1)
            PosCnt = Mid(myVar, 5, 1)
            ErrFun = Mid(myVar, 6, 1)
            ErrRes = Mid(myVar, 7, 1)
            InfoOrd = Mid(myVar, 8, 1)
            InfoSample = Mid(myVar, 9, 6)
            InfoUnit = Mid(myVar, 15, 1)
            InfoWBC = Mid(myVar, 16, 1)
            InfoPLT = Mid(myVar, 17, 1)

            vasActiveCell vasList, lRow, 2      '2009.10.29 이상은

            SetText vasList, lsID, lRow, 2
            SetText vasList, lsRack, lRow, 11
            SetText vasList, lsTube, lRow, 12
            SetText vasList, lsID, lRow, gMaxCol + 4

            SetText vasList, Mid(InfoSample, 2, 1), lRow, gMaxCol
            SetText vasList, Mid(InfoSample, 4, 1), lRow, gMaxCol + 1
            SetText vasList, Mid(InfoSample, 6, 1), lRow, gMaxCol + 2


            SQL = "Select barcode, SlideOrd from res_flag " & vbCrLf & _
                  "where examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "'  " & vbCrLf & _
                  "  and barcode = '" & lsID & "' "
            res = db_select_Col(gLocal, SQL)
            If Trim(gReadBuf(0)) = lsID Then
                lsSlideOrd = Trim(gReadBuf(1))
            Else
                lsSlideOrd = ""
            End If

            Select Case SampleJudg
            Case "0"
                SetText vasList, "Negative", lRow, gMaxCol + 3
                SetBackColor vasList, lRow, lRow, 2, 2, 255, 255, 255
            Case "1"
                SetText vasList, "Positive", lRow, gMaxCol + 3
                SetBackColor vasList, lRow, lRow, 2, 2, 255, 224, 193
                If lsSlideOrd = "SC" Then
                    liMustBe = 1
                End If
            Case "2"
                SetText vasList, "Error", lRow, gMaxCol + 3
                SetBackColor vasList, lRow, lRow, 2, 2, 255, 224, 193
            Case "3"
                SetText vasList, "Potive+Error", lRow, gMaxCol + 3
                SetBackColor vasList, lRow, lRow, 2, 2, 255, 224, 193
                If lsSlideOrd = "SC" Then
                    liMustBe = 1
                End If
            Case "4"
                SetText vasList, "QC Sample", lRow, gMaxCol + 3
                SetBackColor vasList, lRow, lRow, 2, 2, 255, 255, 255
            End Select

            If lsSlideOrd = "SP" Then
                liMustBe = 1
            End If

            If Mid(InfoSample, 2, 1) <> "0" Then
                SetBackColor vasList, lRow, lRow, flagWBC, flagWBC, 193, 234, 255
            Else
                SetBackColor vasList, lRow, lRow, flagWBC, flagWBC, 255, 255, 255
            End If
            If Mid(InfoSample, 4, 1) <> "0" Then
                SetBackColor vasList, lRow, lRow, flagRBC, flagRBC, 193, 234, 255
            Else
                SetBackColor vasList, lRow, lRow, flagRBC, flagRBC, 255, 255, 255
            End If
            If Mid(InfoSample, 6, 1) <> "0" Then
                SetBackColor vasList, lRow, lRow, flagPLT, flagPLT, 193, 234, 255
            Else
                SetBackColor vasList, lRow, lRow, flagPLT, flagPLT, 255, 255, 255
            End If

            SetText vasList, "IPU2", lRow, 10

            '접수여부확인


            '환자정보
            GetPatientInfo2 lsID, lRow

            If liMustBe = 1 Then
                SetText vasList, liMustBe, lRow, gMaxCol + 5
                SetBackColor vasList, lRow, lRow, 9, 9, 255, 224, 193
            Else
                SetText vasList, "", lRow, gMaxCol + 5
                SetBackColor vasList, lRow, lRow, 9, 9, 255, 255, 255
            End If

            gCurRow1 = lRow

            SQL = "Delete from res_flag " & vbCrLf & _
                  "where examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "'  " & vbCrLf & _
                  "  and barcode = '" & lsID & "' "
            res = SendQuery(gLocal, SQL)

            SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
                  "Values ('" & Format(CDate(lsCurDate), "yyyymmdd") & "', '" & lsID & "', '" & SampleJudg & "', '" & PosDif & "', '" & PosMor & "', '" & PosCnt & "', " & _
                  "'" & ErrFun & "', '" & ErrRes & "', '" & Left(InfoSample, 1) & "', '" & Mid(InfoSample, 2, 1) & "', '" & Mid(InfoSample, 3, 1) & "', '" & Mid(InfoSample, 4, 1) & "', " & _
                  "'" & Mid(InfoSample, 5, 1) & "', '" & Mid(InfoSample, 6, 1) & "', '" & InfoWBC & "', '" & InfoPLT & "', '" & liMustBe & "', '" & lsSlideOrd & "' ) "
            res = SendQuery(gLocal, SQL)

        ElseIf Mid(txtBuff2, 3, 1) = "C" Then 'QC
            myVar = ""
        End If
    Case "D2"
        lsWBC = ""
        lsNRBC = ""
        lsEOSIN = ""
        liEOSIN = -1

        If Mid(txtBuff2, 3, 1) = "U" Then '샘플
            myVar = Mid(txtBuff2, 4)

            lsTmp = Left(myVar, 16)     'Instrument Name
            lsTmp = Mid(myVar, 17, 10)  'Sequence No
            lsID = Trim(Mid(myVar, 30, 15))   'Sample ID No.

            If InStr(1, lsID, "QC") > 0 Then
                lsID = lsID = "-2"
            End If


'            lRow = -1
'            For i = vasList.DataRowCnt To 1 Step -1
'                If Trim(GetText(vasList, i, 2)) = lsID And Trim(GetText(vasList, i, 10)) = "IPU2" Then
'                    lRow = i
'                    Exit For
'                End If
'            Next i
'
'            If lRow = -1 Then
'                lRow = vasList.DataRowCnt + 1
'                If lRow > vasList.MaxRows Then
'                    vasList.MaxRows = lRow
'                    vasList.RowHeight(lRow) = 12.6
'                End If
'            End If

            If gCurRow1 > 0 And gCurRow1 <= vasList.DataRowCnt And Trim(GetText(vasList, gCurRow1, 2)) = lsID Then
                lRow = gCurRow1
            Else
                lRow = vasList.DataRowCnt + 1
                If lRow > vasList.MaxRows Then
                    vasList.MaxRows = lRow
                    vasList.RowHeight(lRow) = 12.6
                End If
            End If

            If Trim(GetText(vasList, lRow, 3)) = "" Then
                SetText vasList, lsID, lRow, 2
                SetText vasList, lsID, lRow, gMaxCol + 4

                '접수여부확인
                '환자정보
                GetPatientInfo2 lsID, lRow

                SetText vasList, "IPU2", lRow, 10
            End If

            vasActiveCell vasList, lRow, 2      '2009.10.29 이상은

            ClearSpread vasTemp2
            ClearSpread vasExam2

'            res = Get_Order(lsID)
'            For i = 0 To UBound(gOrder_List)
'                vasExam2.SetText 1, i + 1, gOrder_List(i).TST_CD
'                Select Case Trim(GetText(vasExam2, i, 1))
'                Case "CP0112"
'                    liEOSIN = 1
'                Case "CP0107"   'ESR
'                    Save_Local_ESR lRow, "CP0107", "99", "A", "ESR"
'                End Select
'            Next i

            res = Online_XML(gXml_S07, Trim(lsID))
            For i = 0 To UBound(gExam_Select)
                vasExam2.SetText 1, i + 1, gExam_Select(i).TST_CD
                Select Case Trim(GetText(vasExam2, i, 1))
                Case "CP0112"
                    liEOSIN = 1
                Case "CP0107"   'ESR
                    Save_Local_ESR lRow, "CP0107", "99", "A", "ESR"
                End Select
            Next i
        
            myVar = Mid(myVar, 45)


            ReDim gArrExamRes2(1 To 35)

            For liEquipCode = 1 To 35
                Select Case liEquipCode
                Case 1, 14, 15, 16, 17, 18, 31, 32, 33, 35 '(31 은 FORMAT B 일때)
                    iPoint = 6
                Case Else
                    iPoint = 5
                End Select

                lsResult = Left(myVar, iPoint)
                If InStr(1, lsResult, "*") > 0 Then
                    lsResult = "----"
                End If

                myVar = Mid(myVar, iPoint + 1)

                i = 1
                For i = 1 To UBound(gArrExam)
                    If CInt(gArrExam(i, 1)) = liEquipCode Then
                        z = -1
                        If Trim(GetText(vasList, lRow, gResCol)) = "미접수" Then
                            z = 1

                            gArrExamRes2(liEquipCode).EquipCode = liEquipCode
                            gArrExamRes2(liEquipCode).ExamCode = gArrExam(i, 2)
                            gArrExamRes2(liEquipCode).ExamNo = gArrExam(i, 9)
                            gArrExamRes2(liEquipCode).ExamName = gArrExam(i, 3)
                            gArrExamRes2(liEquipCode).SeqNo = gArrExam(i, 5)
                            gArrExamRes2(liEquipCode).RefLow = gArrExam(i, 6)
                            gArrExamRes2(liEquipCode).RefHigh = gArrExam(i, 7)
                            gArrExamRes2(liEquipCode).RefFlag = ""
                            gArrExamRes2(liEquipCode).res = lsResult
                            gArrExamRes2(liEquipCode).EquipGubun = "IPU2"

                            SetResult2 liEquipCode, i

                            SetText vasList, gArrExamRes2(liEquipCode).res, lRow, gArrExam(i, 10)
                            If gArrExamRes2(liEquipCode).RefFlag = "H" Then
                                SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 255, 127, 0
                            ElseIf gArrExamRes2(liEquipCode).RefFlag = "L" Then
                                SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 127, 255
                            Else
                                SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 0, 0
                            End If

                            Save_Local_One2 lRow, liEquipCode, "A"
                        End If

                        For k = 1 To vasExam2.DataRowCnt
                            If Trim(GetText(vasExam2, k, 1)) = gArrExam(i, 2) Then
                                z = 1

                                gArrExamRes2(liEquipCode).EquipCode = liEquipCode
                                gArrExamRes2(liEquipCode).ExamCode = gArrExam(i, 2)
                                gArrExamRes2(liEquipCode).ExamNo = gArrExam(i, 9)
                                gArrExamRes2(liEquipCode).ExamName = gArrExam(i, 3)
                                gArrExamRes2(liEquipCode).SeqNo = gArrExam(i, 5)
                                gArrExamRes2(liEquipCode).RefLow = gArrExam(i, 6)
                                gArrExamRes2(liEquipCode).RefHigh = gArrExam(i, 7)
                                gArrExamRes2(liEquipCode).RefFlag = ""
                                gArrExamRes2(liEquipCode).res = lsResult
                                gArrExamRes2(liEquipCode).EquipGubun = "IPU2"

                                SetResult2 liEquipCode, i

                                SetText vasList, gArrExamRes2(liEquipCode).res, lRow, gArrExam(i, 10)
                                If gArrExamRes2(liEquipCode).RefFlag = "H" Then
                                    SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 255, 127, 0
                                ElseIf gArrExamRes2(liEquipCode).RefFlag = "L" Then
                                    SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 127, 255
                                Else
                                    SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 0, 0
                                End If

                                Save_Local_One2 lRow, liEquipCode, "A"

                                DeleteRow vasExam2, k, k

                                Exit For
                            End If
                        Next k
                        If z = 1 Then
                            Exit For
                        End If
                    End If
                Next i

                Select Case liEquipCode
                Case 1
                    lsWBC = gArrExamRes2(liEquipCode).res
                Case 12
                    lsEOSIN = gArrExamRes2(liEquipCode).res
                Case 32
                    lsNRBC = gArrExamRes2(liEquipCode).res
                End Select


                'SetText vasList, lsResult, lRow, grescol + liEquipCode
            Next liEquipCode

'            If liEOSIN = 1 Then
'                lsEOSIN = Format(CCur(lsEOSIN) * CCur(lsWBC) * 10, "#0")
'                Save_Local_ESR lRow, "CP0112", "98", "A", "Eos.Diif.Cnt", lsEOSIN
'            End If

            vasList.Row = lRow
            vasList.Col = 1
            If vasList.Value = 0 Then
                SetText vasList, "수신", lRow, gResCol
            End If

            If lRow < 1 Or lRow > vasList.DataRowCnt Then Exit Sub

'            vasList.Row = lRow
'            vasList.Col = 1
            If chkMode.Value = 1 Then
                liRet = 1
                'lsInscode = "01"
                lsInscode = IPU2.UseEquip


                'debug.print lRow & " : " & lsID
                res = ToServer(lRow, vasList)
                If res = 1 Then
                    vasList.Row = lRow
                    vasList.Col = 1
                    vasList.Value = 0

                    SQL = "update pat_res set sendflag = 'B' " & vbCrLf & _
                          "WHERE examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "' " & vbCrLf & _
                          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                          "  AND barcode = '" & lsID & "' "
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then SaveQuery SQL

                    SetText vasList, "완료", lRow, gResCol
                    SetBackColor vasList, lRow, lRow, 1, 1, 202, 255, 112
                    'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid


                ElseIf res = 2 Then
                    vasList.Row = lRow
                    vasList.Col = 1
                    vasList.Value = 1

                    'SetText vasList, "결과", lRow, gResCol
                Else
                    SQL = "update pat_res set sendflag = 'E' " & vbCrLf & _
                          "WHERE examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "' " & vbCrLf & _
                          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                          "  AND barcode = '" & lsID & "' "
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then SaveQuery SQL

                    vasList.Row = lRow
                    vasList.Col = 1
                    vasList.Value = 1

                    SetText vasList, "실패", lRow, gResCol
                    SetForeColor vasList, lRow, lRow, 1, 1, 255, 0, 0
                    'SetBackColor vasID, gPreRow, gPreRow, colCheckBox, vasID.MaxCols, 255, 255, 255
                    'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
                End If

            End If
        ElseIf Mid(txtBuff2, 3, 1) = "C" Then 'QC
            myVar = ""
        End If
    End Select

End Sub

Sub XE2100_ASTM_1()
    Dim lsID    As String
    Dim i, j, iCnt As Integer
    Dim lRow, lResRow As Long
    Dim lsData As String
    Dim lsTmp As String
    Dim lsEquip, lsResult, lsUnit, lsFlag As String
    Dim liRet As Integer
    
    Dim mExam As Variant
    
    Dim SampleJudg, PosDif, PosMor, PosCnt, ErrFun, ErrRes As String
    Dim InfoOrd, InfoSample, InfoUnit, InfoWBC, InfoPLT As String
        
    Dim lsInscode As String
    Dim lsWBC As String
    Dim lsNRBC As String
    Dim lsEOSIN As String
    Dim lsC_WBC As String
    Dim lsNEUT As String
    
    Dim liC_WBC As Integer
    Dim liEOSIN As Integer
    

    Dim lsSmearCode As String
    Dim lsSlideOrd As String
    Dim liMustBe As Integer
    
    Dim lsCurDate As String
    
    Dim liEquipCode As Integer
    
    lsCurDate = GetDateFull
        
    lsData = Trim(txtBuff1)
    Select Case Mid(lsData, 2, 2)
    Case "H|"
        dtpExamDate.Value = Format(CDate(GetDateFull), "yyyy-mm-dd")
        
        ClearSpread vasTemp_1
        ClearSpread vasRes_1
        gID = ""
        gRack = ""
        gPos = ""
        
        liMustBe = 0
        
        ReDim gArrExamRes1(0)
            
        iCnt = 0
        i = InStr(1, lsData, "|")
        Do While i > 0
            iCnt = iCnt + 1
            lsTmp = Left(lsData, i - 1)
            lsData = Mid(lsData, i + 1)
            
            If iCnt = 5 Then
                j = InStr(1, lsTmp, "^")
                If j > 0 Then
                    lsTmp = Mid(lsTmp, j + 1)
                End If
                
                j = InStr(1, lsTmp, "^")
                If j > 0 Then
                    lsTmp = Mid(lsTmp, j + 1)
                End If
                
                j = InStr(1, lsTmp, "^")
                If j > 0 Then
                    lsTmp = Trim(Mid(lsTmp, 1, j - 1))
                    
                    Select Case lsTmp
                    Case "A1292"
                        gsVersion = "XE2100A"
                    Case "F5188"
                        gsVersion = "XE2100B"
                    End Select
                    
                    Exit Do
                End If
                
'            ElseIf iCnt = 13 Then
'                gsVersion = lsTmp
'
'                Exit Do
            End If
            
            i = InStr(1, lsData, "|")
        Loop
        
        liMustBe = 0
        
        ReDim gArrExamRes1(0)
            
    Case "P|"
    Case "O|"
        
        lsWBC1 = ""
        lsNEUT1 = ""


        iCnt = 0
        i = InStr(1, lsData, "|")
        Do While i > 0
            iCnt = iCnt + 1
            lsTmp = Left(lsData, i - 1)
            lsData = Mid(lsData, i + 1)
            
            If iCnt = 4 Then
                i = InStr(1, lsTmp, "^")
                If i > 0 Then
                    gRack = Trim(Left(lsTmp, i - 1))
                    lsTmp = Mid(lsTmp, i + 1)
                    
                    i = InStr(1, lsTmp, "^")
                    If i > 0 Then
                        gPos = Trim(Left(lsTmp, i - 1))
                        lsTmp = Mid(lsTmp, i + 1)
                        
                        i = InStr(1, lsTmp, "^")
                        If i > 0 Then
                            gID = Trim(Left(lsTmp, i - 1))
'                            If IsNumeric(gID) = True And Len(gID) = 9 Then
'                                gID = "0" & gID
'                            End If
                        Else
                            gID = Trim(lsTmp)
                        End If
                    End If
                End If
                    
                Exit Do
            End If
            
            i = InStr(1, lsData, "|")
        Loop
    
        If InStr(1, gID, "QC") > 0 Then
            gID = gID & "-1"
        End If
            
        ClearSpread vasTemp_1
        
        lRow = vasList.DataRowCnt + 1
        If lRow > vasList.MaxRows Then
            vasList.MaxRows = lRow
        End If
        
        gRow = lRow
        
        vasList.SetText 2, lRow, gID
        vasList.SetText 11, lRow, gRack
        vasList.SetText 12, lRow, gPos
        vasList.SetText gMaxCol + 4, lRow, gID
        
        vasActiveCell vasList, lRow, 2      '2009.10.29 이상은
        
        SetBackColor vasList, lRow, lRow, 1, 1, 255, 250, 205
        
'        If IsNumeric(gID) = True And Len(gID) = 10 And InStr(1, gID, "ERR") < 1 Then
'            SQL = " Select TestCd " & vbCrLf & _
'                  " From LIMAS301 " & vbCrLf & _
'                  " Where SpcNo = '" & gID & "'"
'            res = db_select_Vas(gServer, SQL, vasTemp_1)
'
'            Get_Sample_Info lRow
'
'            vasList.Row = lRow
'            vasList.Col = 1
'            vasList.Value = 0
'
'        Else
'            vasList.Row = lRow
'            vasList.Col = 1
'            vasList.Value = 1
'        End If
                
'        SampleJudg = Mid(myVar, 2, 1)
'        PosDif = Mid(myVar, 3, 1)
'        PosMor = Mid(myVar, 4, 1)
'        PosCnt = Mid(myVar, 5, 1)
'        ErrFun = Mid(myVar, 6, 1)
'        ErrRes = Mid(myVar, 7, 1)
'        InfoOrd = Mid(myVar, 8, 1)
'        InfoSample = Mid(myVar, 9, 6)
'        InfoUnit = Mid(myVar, 15, 1)
'        InfoWBC = Mid(myVar, 16, 1)
'        InfoPLT = Mid(myVar, 17, 1)
        
        SetText vasList, gID, lRow, 2
        SetText vasList, gRack, lRow, 11
        SetText vasList, gPos, lRow, 12
        SetText vasList, gID, lRow, gMaxCol + 4
        
'        SetText vasList, Mid(InfoSample, 2, 1), lRow, gMaxCol
'        SetText vasList, Mid(InfoSample, 4, 1), lRow, gMaxCol + 1
'        SetText vasList, Mid(InfoSample, 6, 1), lRow, gMaxCol + 2
        
        SQL = " delete from pat_resmemo  " & vbCrLf & _
              " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & gID & "' "
        res = SendQuery(gLocal, SQL)
        
        SQL = "Select barcode, SlideOrd from res_flag " & vbCrLf & _
              "where examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "'  " & vbCrLf & _
              "  and barcode = '" & gID & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = gID Then
            lsSlideOrd = Trim(gReadBuf(1))
        Else
            lsSlideOrd = ""
        End If
        
        Select Case SampleJudg
        Case "0"
            SetText vasList, "Negative", lRow, gMaxCol + 3
            SetBackColor vasList, lRow, lRow, 2, 2, 255, 255, 255
        Case "1"
            SetText vasList, "Positive", lRow, gMaxCol + 3
            SetBackColor vasList, lRow, lRow, 2, 2, 255, 224, 193
            If lsSlideOrd = "SC" Then
                liMustBe = 1
            End If
        Case "2"
            SetText vasList, "Error", lRow, gMaxCol + 3
            SetBackColor vasList, lRow, lRow, 2, 2, 255, 224, 193
        Case "3"
            SetText vasList, "Potive+Error", lRow, gMaxCol + 3
            SetBackColor vasList, lRow, lRow, 2, 2, 255, 224, 193
            If lsSlideOrd = "SC" Then
                liMustBe = 1
            End If
        Case "4"
            SetText vasList, "QC Sample", lRow, gMaxCol + 3
            SetBackColor vasList, lRow, lRow, 2, 2, 255, 255, 255
        End Select
        
        If lsSlideOrd = "SP" Then
            liMustBe = 1
        End If
        
'        If Mid(InfoSample, 2, 1) <> "0" Then
'            SetBackColor vasList, lRow, lRow, flagWBC, flagWBC, 193, 234, 255
'        Else
'            SetBackColor vasList, lRow, lRow, flagWBC, flagWBC, 255, 255, 255
'        End If
'        If Mid(InfoSample, 4, 1) <> "0" Then
'            SetBackColor vasList, lRow, lRow, flagRBC, flagRBC, 193, 234, 255
'        Else
'            SetBackColor vasList, lRow, lRow, flagRBC, flagRBC, 255, 255, 255
'        End If
'        If Mid(InfoSample, 6, 1) <> "0" Then
'            SetBackColor vasList, lRow, lRow, flagPLT, flagPLT, 193, 234, 255
'        Else
'            SetBackColor vasList, lRow, lRow, flagPLT, flagPLT, 255, 255, 255
'        End If
                
        SetText vasList, gsVersion, lRow, 10
        
        '환자정보
        GetPatientInfo1 gID, lRow
        
        '검사코드 불러오기
        ClearSpread vasExam1
    
        If IsNumeric(gID) = False Or Len(gID) < 11 Then
        Else
            'res = Online_XML(gXml_S07, Trim(gID))
            res = Online_XML(gXml_S08, Trim(gID))
            For i = 0 To UBound(gExam_Select)
                vasExam1.SetText 1, i + 1, gExam_Select(i).TST_CD
            Next i
        End If
        
        If liMustBe = 1 Then
            SetText vasList, liMustBe, lRow, gMaxCol + 5
            SetBackColor vasList, lRow, lRow, 9, 9, 255, 224, 193
            'SetForeColor vasList, lRow, lRow, 9, 9, 0, 0, 160
'                vasList.Row = lRow
'                vasList.Col = 9
'                vasList.FontBold = True
        Else
            SetText vasList, "", lRow, gMaxCol + 5
            SetBackColor vasList, lRow, lRow, 9, 9, 255, 255, 255
            'SetForeColor vasList, lRow, lRow, 9, 9, 0, 0, 0
'                vasList.Row = lRow
'                vasList.Col = 9
'                vasList.FontBold = False
        End If
        
        gCurRow = lRow
        
        SQL = "Delete from res_flag " & vbCrLf & _
              "where examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "'  " & vbCrLf & _
              "  and barcode = '" & gID & "' "
        res = SendQuery(gLocal, SQL)
        
        SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
              "Values ('" & Format(CDate(lsCurDate), "yyyymmdd") & "', '" & gID & "', '" & SampleJudg & "', '" & PosDif & "', '" & PosMor & "', '" & PosCnt & "', " & _
              "'" & ErrFun & "', '" & ErrRes & "', '" & Left(InfoSample, 1) & "', '" & Mid(InfoSample, 2, 1) & "', '" & Mid(InfoSample, 3, 1) & "', '" & Mid(InfoSample, 4, 1) & "', " & _
              "'" & Mid(InfoSample, 5, 1) & "', '" & Mid(InfoSample, 6, 1) & "', '" & InfoWBC & "', '" & InfoPLT & "', '" & liMustBe & "', '" & lsSlideOrd & "' ) "
        res = SendQuery(gLocal, SQL)
                
        SQL = " delete From pat_resmemo " & vbCrLf & _
              " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' "
        res = SendQuery(gLocal, SQL)
                
        gRow = lRow
        
    Case "R|"
        gMsgFlag = "R"
        
        iCnt = 0
        i = InStr(1, lsData, "|")
        Do While i > 0
            iCnt = iCnt + 1
            lsTmp = Left(lsData, i - 1)
            lsData = Mid(lsData, i + 1)
            
            Select Case iCnt
            Case 3
                lsTmp = Mid(lsTmp, 5)
                j = InStr(1, lsTmp, "^")
                If j > 0 Then
                    lsEquip = Trim(Left(lsTmp, j - 1))
                Else
                    lsEquip = Trim(lsTmp)
                End If
            Case 4
                lsResult = Trim(lsTmp)
            Case 5
                lsUnit = Trim(lsTmp)
            Case 6
            Case 7
                lsFlag = Trim(lsTmp)
                
                If lsEquip = "WBC" Then lsWBC1 = lsResult
                If lsEquip = "NEUT%" Then lsNEUT1 = lsResult
                    
                
                lResRow = vasRes_1.DataRowCnt + 1
                vasRes_1.MaxRows = lResRow
                vasRes_1.SetText 1, lResRow, lsEquip
                vasRes_1.SetText 2, lResRow, lsResult
                vasRes_1.SetText 3, lResRow, lsFlag
                vasRes_1.SetText 4, lResRow, lsUnit
                
                Exit Do
            End Select
            
            i = InStr(1, lsData, "|")
        Loop
        
        Set_Sample_Result_1 lResRow
        
        If IsNumeric(lsWBC1) And IsNumeric(lsNEUT1) Then
            lsWBC1 = Format(CCur(lsWBC1) * CCur(lsNEUT1) * 0.01, "#0.00")
            
            lResRow = vasRes_1.DataRowCnt + 1
            vasRes_1.MaxRows = lResRow
            vasRes_1.SetText 1, lResRow, "ANC"
            vasRes_1.SetText 2, lResRow, lsWBC1
            vasRes_1.SetText 3, lResRow, ""
            vasRes_1.SetText 4, lResRow, ""
            
            Set_Sample_Result_1 lResRow
            
            lsWBC1 = ""
            lsNEUT1 = ""
            
        End If
        
        
        vasList.Row = gRow
        vasList.Col = 1
        If vasList.Value = 0 Then
            SetText vasList, "수신", gRow, gResCol
        End If
        
    Case "L|"
        
        liC_WBC = -1
        liEOSIN = -1
        
        vasList.Row = gRow
        vasList.Col = 1
        If chkMode.Value = 1 And gMsgFlag = "R" Then
'            SaveData gRow & " : " & gID
            res = ToServer(gRow, vasList)
            If res = 1 Then
                vasList.Row = gRow
                vasList.Col = 1
                vasList.Value = 0

                SetText vasList, "완료", gRow, gResCol
                SetBackColor vasList, gRow, gRow, 1, 1, 202, 255, 112
                'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
            
                SQL = "update pat_res set sendflag = 'B' " & vbCrLf & _
                      "WHERE examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "' " & vbCrLf & _
                      "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                      "  AND barcode = '" & gID & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then SaveQuery SQL
                
                SQL = "Update worklist set OrdFlag = 'D' where barcode = '" & gID & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then SaveQuery SQL
                
'                SQL = "select sid from cpslipuse where sid = '" & gID & "' "
'                res = db_select_Col(gServer, SQL)
'                If Trim(gReadBuf(0)) <> Trim(gID) Then
'                    SetBackColor vasList, gRow, gRow, 1, 1, 255, 0, 0
'                    SetForeColor vasList, gRow, gRow, gResCol - 3, gResCol, 255, 0, 0
'                    vasList.SetText gResCol, lRow, "미접수"
'                End If
            ElseIf res = 2 Then
                vasList.Row = gRow
                vasList.Col = 1
                vasList.Value = 1
                
                'SetText vasList, "결과", lRow, gResCol
            Else
                vasList.Row = gRow
                vasList.Col = 1
                vasList.Value = 1
            
                SQL = "Update worklist set OrdFlag = 'E' where barcode = '" & gID & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then SaveQuery SQL
                
                SQL = "update pat_res set sendflag = 'E' " & vbCrLf & _
                      "WHERE examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "' " & vbCrLf & _
                      "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                      "  AND barcode = '" & gID & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then SaveQuery SQL
                
                SetText vasList, "실패", gRow, gResCol
                SetForeColor vasList, gRow, gRow, 1, 1, 255, 0, 0
                'SetBackColor vasID, gPreRow, gPreRow, colCheckBox, vasID.MaxCols, 255, 255, 255
                'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
            End If
        Else
            SQL = "Update worklist set OrdFlag = 'C' where barcode = '" & gID & "' "
            res = SendQuery(gLocal, SQL)
            If res = -1 Then SaveQuery SQL
        End If
    Case "Q|"
        gMsgFlag = "Q"
        
        iCnt = 0
        i = InStr(1, lsData, "|")
        Do While i > 0
            iCnt = iCnt + 1
            lsTmp = Left(lsData, i - 1)
            lsData = Mid(lsData, i + 1)
            
            If iCnt = 3 Then
                i = InStr(1, lsTmp, "^")
                If i > 0 Then
                    gRack = Trim(Left(lsTmp, i - 1))
                    lsTmp = Mid(lsTmp, i + 1)
                    
                    i = InStr(1, lsTmp, "^")
                    If i > 0 Then
                        gPos = Trim(Left(lsTmp, i - 1))
                        lsTmp = Mid(lsTmp, i + 1)
                        
                        i = InStr(1, lsTmp, "^")
                        If i > 0 Then
                            gID = Trim(Left(lsTmp, i - 1))
                            gFlag = Mid(lsTmp, i + 1)
'                            If IsNumeric(gID) = True And Len(gID) = 9 Then
'                                gID = "0" & gID
'                            End If
                        Else
                            gID = Trim(lsTmp)
                            gFlag = ""
                        End If
                    End If
                End If
                
                Exit Do
            End If
            
            i = InStr(1, lsData, "|")
        Loop
        
    End Select
End Sub

Function Set_Sample_Result_1(asRow As Long) As Long
    Dim lRow As Long
    Dim i, j, k, z As Long
    Dim lsEquip As String
    Dim lsResult As String
    Dim lsFlagRes As String
    
    Dim lsFlag As String
    Dim lsUnit As String
    Dim liEquipCode As Integer
    
    Dim lsMsg As String
    
    Set_Sample_Result_1 = -1
    
    If asRow < 1 Or asRow > vasRes_1.DataRowCnt Then Exit Function
    
    lRow = gRow
    
    lsEquip = Trim(GetText(vasRes_1, asRow, 1))
    lsResult = Trim(GetText(vasRes_1, asRow, 2))
    
    lsFlag = Trim(GetText(vasRes_1, asRow, 3))
    lsUnit = Trim(GetText(vasRes_1, asRow, 4))

    If lsFlag = "A" Then
        lsFlagRes = ""
        If lsResult <> "" Then
            lsFlagRes = "[" & lsResult & "]"
        End If
        
        lsMsg = ""
        
        Select Case lsEquip
        'WBC*************************
        Case "WBC_Abn_Scattergram"
            lsMsg = "WBC abnormal scatergram" & lsFlagRes
        Case "NRBC_Abn_Scattergram"
            lsMsg = "NRBC Abn Scg" & lsFlagRes
        Case "Neutropenia"
            lsMsg = "Neutro-" & lsFlagRes
        Case "Neutrophilia"
            lsMsg = "Neutro+" & lsFlagRes
        Case "Lymphopenia"
            lsMsg = "Lympho-" & lsFlagRes
        Case "Lymphocytosis"
            lsMsg = "Lympho+" & lsFlagRes
        Case "Monocytosis"
            lsMsg = "Mono+" & lsFlagRes
        Case "Eosinophilia"
            lsMsg = "Eo+" & lsFlagRes
        Case "Basophilia"
            lsMsg = "Baso+" & lsFlagRes
        Case "Leukocytopenia"
            lsMsg = "Leuko-" & lsFlagRes
        Case "Leukocytosis"
            lsMsg = "Leuko+" & lsFlagRes
        Case "NRBC_Present"
            lsMsg = lsEquip & lsFlagRes
        Case "Blasts?"
            lsMsg = lsEquip & lsFlagRes
        Case "Left_Shift?"
            lsMsg = lsEquip & lsFlagRes
        Case "NRBC?"
            lsMsg = lsEquip & lsFlagRes
        Case "Immature_Gran?"
            lsMsg = "Immature granules" & lsFlagRes
        Case "Atypical_Lympho?"
            lsMsg = "Atypical lymphocytes" & lsFlagRes
        Case "Abn_Lympho/L_Blasts?"
            lsMsg = "Abn Ly/L_Bl?" & lsFlagRes
        Case "RBC_Lyse Resistance?"
            lsMsg = "RBC Lyse resistance" & lsFlagRes

        'RBC*************************
        Case "RBC_Abn_Distribution"
            lsMsg = "RBC Abn Dst" & lsFlagRes
        Case "Dimorphic_Population"
            lsMsg = "Dimorph Pop" & lsFlagRes
        Case "RET_Abn_Scattergram"
            lsMsg = "RET Abn Scg" & lsFlagRes
        Case "Reticulocytosis"
            lsMsg = "Reticulo" & lsFlagRes
        Case "Anisocytosis"
            lsMsg = "Aniso" & lsFlagRes
        Case "Microcytosis"
            lsMsg = "Micro" & lsFlagRes
        Case "Macrocytosis"
            lsMsg = "Macro" & lsFlagRes

        Case "Hypochromia", "Anemia", "HGB_Defect?", "Fragments?"
            lsMsg = lsEquip & lsFlagRes

        Case "Erythrocytosis"
            lsMsg = "Erythro+" & lsFlagRes
'
        Case "RBC_Agglutination?"
            lsMsg = "RBC agglutination" & lsFlagRes
        Case "Turbidity/HGB Interference?"
            lsMsg = "Turbidty/Hb interface" & lsFlagRes
        Case "Iron_Deficiency?"
            lsMsg = "Iron Def?" & lsFlagRes

        'PLT*************************
        Case "PLT_Abn_Scattergram"
            lsMsg = "PLT Abn Scg" & lsFlagRes
        Case "PLT_Abn_Distribution"
            lsMsg = "PLT Abn Dst" & lsFlagRes
        Case "Thrombocytopenia"
            lsMsg = "Thrombo-" & lsFlagRes
        Case "Thrombocytosis"
            lsMsg = "Thrombo+" & lsFlagRes
        Case "PLT_Clumps?"
            lsMsg = "PLT Clumps" & lsFlagRes
        Case "PLT_Clumps(S)?"
            lsMsg = "PLT Clumps(S)" & lsFlagRes
            
        Case "Positive_Morph", "Positive_Count"
            lsMsg = lsEquip & lsFlagRes
            
            SQL = "select barcode from res_flag " & vbCrLf & _
                  "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "'  " & vbCrLf & _
                  "  and barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' "
            res = db_select_Col(gLocal, SQL)
            If Trim(gReadBuf(0)) = Trim(GetText(vasList, lRow, 2)) Then
                SQL = "Update res_flag set SampleJudg = '1'  " & vbCrLf & _
                      "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "'  " & vbCrLf & _
                      "  and barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' "
                res = SendQuery(gLocal, SQL)
            Else
                SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
                      "Values ('" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "', '" & Trim(GetText(vasList, lRow, 2)) & "', '1', '', '', '', " & _
                      "'', '', '', '', '', '', " & _
                      "'', '', '', '', '', '' ) "
                res = SendQuery(gLocal, SQL)
            End If
        
        Case Else   '2010.03.20 이상은 추가
            lsMsg = lsEquip & lsFlagRes
        End Select
        
        '메모결과 입력
        If Trim(lsMsg) <> "" Then Save_ResMemo_1 lRow, lsMsg
        
    End If
    
    SQL = "Update res_flag set SampleJudg = '1'  " & vbCrLf & _
          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "'  " & vbCrLf & _
          "  and barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' " & vbCrLf & _
          "  and PBSFlag = '1'"
    res = SendQuery(gLocal, SQL)
    
    'gArrExam
    'EquipCode,ExamCode, ExamName, Seqno, PointSize,
    'RefLow, RefHigh, RSGubun
    
    k = UBound(gArrExamRes1)
    k = k + 1
    ReDim Preserve gArrExamRes1(k)
                        
    i = 1
    For i = 1 To UBound(gArrExam)
        If Trim(gArrExam(i, 3)) = lsEquip Then
            z = -1
            If Trim(GetText(vasList, lRow, gResCol)) = "미접수" Then
                z = 1
                
                gArrExamRes1(k).EquipCode = gArrExam(i, 1)
                gArrExamRes1(k).ExamCode = gArrExam(i, 2)
                gArrExamRes1(k).ExamNo = gArrExam(i, 9)
                gArrExamRes1(k).ExamName = gArrExam(i, 3)
                gArrExamRes1(k).SeqNo = gArrExam(i, 5)
                gArrExamRes1(k).RefLow = gArrExam(i, 6)
                gArrExamRes1(k).RefHigh = gArrExam(i, 7)
                gArrExamRes1(k).RefFlag = ""
                gArrExamRes1(k).res = lsResult
                gArrExamRes1(k).EquipGubun = "IPU1"
            
                SetResult1 k, i
                                                
                SetText vasList, gArrExamRes1(k).res, lRow, gArrExam(i, 10)
                'SetText vasList, gArrExamRes1(k).res, lRow, gResCol + i
                
                Select Case gArrExamRes1(k).RefFlag
                Case "H", ">", "A"
                    SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 255, 127, 0
                Case "L", "W"
                    SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 127, 255
                Case Else
                    SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 0, 0
                End Select
                
                Save_Local_One1 lRow, k, "A"
                
            End If
            
            For j = 1 To vasExam1.DataRowCnt
                If Trim(GetText(vasExam1, j, 1)) = gArrExam(i, 2) Then
                    z = 1
                    
                    gArrExamRes1(k).EquipCode = gArrExam(i, 1)
                    gArrExamRes1(k).ExamCode = gArrExam(i, 2)
                    gArrExamRes1(k).ExamNo = gArrExam(i, 9)
                    gArrExamRes1(k).ExamName = gArrExam(i, 3)
                    gArrExamRes1(k).SeqNo = gArrExam(i, 5)
                    gArrExamRes1(k).RefLow = gArrExam(i, 6)
                    gArrExamRes1(k).RefHigh = gArrExam(i, 7)
                    gArrExamRes1(k).RefFlag = ""
                    gArrExamRes1(k).res = lsResult
                    gArrExamRes1(k).EquipGubun = "IPU1"
                    
                    SetResult1 k, i
                                                    
                    SetText vasList, gArrExamRes1(k).res, lRow, gArrExam(i, 10)
                    Select Case gArrExamRes1(k).RefFlag
                    Case "H", ">", "A"
                        SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 255, 127, 0
                    Case "L", "W"
                        SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 127, 255
                    Case Else
                        SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 0, 0
                    End Select
                    
                    Save_Local_One1 lRow, k, "A"
                                           
                    DeleteRow vasExam1, j, j
                    
                    Exit For
                End If
            Next j
            If z = 1 Then
                Exit For
            End If
        End If
    Next i

    If z = -1 Then
        i = 1
        For i = 1 To UBound(gArrExam)
            If Trim(gArrExam(i, 3)) = lsEquip Then
                    z = 1
                    
                    gArrExamRes1(k).EquipCode = gArrExam(i, 1)
                    gArrExamRes1(k).ExamCode = gArrExam(i, 2)
                    gArrExamRes1(k).ExamNo = gArrExam(i, 9)
                    gArrExamRes1(k).ExamName = gArrExam(i, 3)
                    gArrExamRes1(k).SeqNo = gArrExam(i, 5)
                    gArrExamRes1(k).RefLow = gArrExam(i, 6)
                    gArrExamRes1(k).RefHigh = gArrExam(i, 7)
                    gArrExamRes1(k).RefFlag = ""
                    gArrExamRes1(k).res = lsResult
                    gArrExamRes1(k).EquipGubun = "IPU1"
                
                    SetResult1 k, i
                                                    
                    SetText vasList, gArrExamRes1(k).res, lRow, gArrExam(i, 10)
                    'SetText vasList, gArrExamRes1(k).res, lRow, gResCol + i
                    
                    Select Case gArrExamRes1(k).RefFlag
                    Case "H", ">", "A"
                        SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 255, 127, 0
                    Case "L", "W"
                        SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 127, 255
                    Case Else
                        SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 0, 0
                    End Select
                    
                    Save_Local_One1 lRow, k, "A"
                    
                    Exit For
            End If
        Next i
    End If

End Function

Function Save_ResMemo_1(ByVal asRow As Long, asMessage As String)
'메시지 저장하기
    Dim sMessage As String
    
    If asMessage = "" Then
        Exit Function
    End If
    
    sMessage = ""
    
    SQL = " Select message From pat_resmemo " & vbCrLf & _
          " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(vasList, asRow, 2)) & "' "
    res = db_select_Col(gLocal, SQL)
    
    sMessage = Trim(gReadBuf(0))
    
    If sMessage = "" Then
        SQL = " Insert Into pat_resmemo (examdate, equipno, barcode, message) " & vbCrLf & _
              " VALUES ('" & Format(dtpExamDate.Value, "yyyymmdd") & "', '" & gEquip & "', " & vbCrLf & _
              "         '" & Trim(GetText(vasList, asRow, 2)) & "', '" & asMessage & "') "
    Else
        'sMessage = sMessage & vbCrLf & asMessage
        sMessage = sMessage & ", " & asMessage

        SQL = " Update pat_resmemo Set " & vbCrLf & _
              " message = '" & Trim(sMessage) & "' " & vbCrLf & _
              " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasList, asRow, 2)) & "' "
    End If
    
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function


Sub XE2100_ASTM_2()
    Dim lsID As String
    Dim i, j, iCnt As Integer
    Dim lRow, lResRow As Long
    Dim lsData As String
    Dim lsTmp As String
    Dim lsEquip, lsResult, lsUnit, lsFlag As String
    Dim liRet As Integer
    
    Dim mExam As Variant
    
    Dim SampleJudg, PosDif, PosMor, PosCnt, ErrFun, ErrRes As String
    Dim InfoOrd, InfoSample, InfoUnit, InfoWBC, InfoPLT As String
        
    Dim lsInscode As String
    Dim lsNRBC As String
    Dim lsEOSIN As String
    Dim lsC_WBC As String
    
    
    Dim liC_WBC As Integer
    Dim liEOSIN As Integer
    

    Dim lsSmearCode As String
    Dim lsSlideOrd As String
    Dim liMustBe As Integer
    
    Dim lsCurDate As String
    
    Dim liEquipCode As Integer
    
    lsCurDate = GetDateFull
        
    lsData = Trim(txtBuff2)
    Select Case Mid(lsData, 2, 2)
    Case "H|"
        dtpExamDate.Value = Format(CDate(GetDateFull), "yyyy-mm-dd")
        
        ClearSpread vasTemp_2
        ClearSpread vasRes_2
        gID1 = ""
        gRack1 = ""
        gPos1 = ""
        
        liMustBe = 0
        
        iCnt = 0
        i = InStr(1, lsData, "|")
        Do While i > 0
            iCnt = iCnt + 1
            lsTmp = Left(lsData, i - 1)
            lsData = Mid(lsData, i + 1)
            
            If iCnt = 5 Then
                j = InStr(1, lsTmp, "^")
                If j > 0 Then
                    lsTmp = Mid(lsTmp, j + 1)
                End If
                
                j = InStr(1, lsTmp, "^")
                If j > 0 Then
                    lsTmp = Mid(lsTmp, j + 1)
                End If
                
                j = InStr(1, lsTmp, "^")
                If j > 0 Then
                    lsTmp = Trim(Mid(lsTmp, 1, j - 1))
                    
                    Select Case lsTmp
                    Case "A1292"
                        gsVersion = "XE2100A"
                    Case "F5188"
                        gsVersion = "XE2100B"
                    End Select
                    
                    Exit Do
                End If
                
'            ElseIf iCnt = 13 Then
'                gsVersion = lsTmp
'
'                Exit Do
            End If
            
            i = InStr(1, lsData, "|")
        Loop
        
        
        ReDim gArrExamRes2(0)
            
    Case "P|"
    Case "O|"
        lsWBC2 = ""
        lsNEUT2 = ""
        
        iCnt = 0
        i = InStr(1, lsData, "|")
        Do While i > 0
            iCnt = iCnt + 1
            lsTmp = Left(lsData, i - 1)
            lsData = Mid(lsData, i + 1)
            
            If iCnt = 4 Then
                i = InStr(1, lsTmp, "^")
                If i > 0 Then
                    gRack1 = Trim(Left(lsTmp, i - 1))
                    lsTmp = Mid(lsTmp, i + 1)
                    
                    i = InStr(1, lsTmp, "^")
                    If i > 0 Then
                        gPos1 = Trim(Left(lsTmp, i - 1))
                        lsTmp = Mid(lsTmp, i + 1)
                        
                        i = InStr(1, lsTmp, "^")
                        If i > 0 Then
                            gID1 = Trim(Left(lsTmp, i - 1))
'                            If IsNumeric(gID1) = True And Len(gID1) = 9 Then
'                                gID1 = "0" & gID1
'                            End If
                        Else
                            gID1 = Trim(lsTmp)
                        End If
                    End If
                End If
                    
                Exit Do
            End If
            
            i = InStr(1, lsData, "|")
        Loop
    
        If InStr(1, gID1, "QC") > 0 Then
            gID1 = gID1 & "-2"
        End If
            
        ClearSpread vasTemp_2
        
        lRow = vasList.DataRowCnt + 1
        If lRow > vasList.MaxRows Then
            vasList.MaxRows = lRow
        End If
        
        gRow1 = lRow
        
        vasList.SetText 2, lRow, gID1
        vasList.SetText 11, lRow, gRack1
        vasList.SetText 12, lRow, gPos1
        vasList.SetText gMaxCol + 4, lRow, gID1
        
        vasActiveCell vasList, lRow, 2      '2009.10.29 이상은
        
        SetBackColor vasList, lRow, lRow, 1, 1, 255, 250, 205
        
'        If IsNumeric(gID1) = True And Len(gID1) = 10 And InStr(1, gID1, "ERR") < 1 Then
'            SQL = " Select TestCd " & vbCrLf & _
'                  " From LIMAS301 " & vbCrLf & _
'                  " Where SpcNo = '" & gID1 & "'"
'            res = db_select_Vas(gServer, SQL, vasTemp_2)
'
'            Get_Sample_Info lRow
'
'            vasList.Row = lRow
'            vasList.Col = 1
'            vasList.Value = 0
'
'        Else
'            vasList.Row = lRow
'            vasList.Col = 1
'            vasList.Value = 1
'        End If
                
'        SampleJudg = Mid(myVar, 2, 1)
'        PosDif = Mid(myVar, 3, 1)
'        PosMor = Mid(myVar, 4, 1)
'        PosCnt = Mid(myVar, 5, 1)
'        ErrFun = Mid(myVar, 6, 1)
'        ErrRes = Mid(myVar, 7, 1)
'        InfoOrd = Mid(myVar, 8, 1)
'        InfoSample = Mid(myVar, 9, 6)
'        InfoUnit = Mid(myVar, 15, 1)
'        InfoWBC = Mid(myVar, 16, 1)
'        InfoPLT = Mid(myVar, 17, 1)
        
        SetText vasList, gID1, lRow, 2
        SetText vasList, gRack1, lRow, 11
        SetText vasList, gPos1, lRow, 12
        SetText vasList, gID1, lRow, gMaxCol + 4
        
'        SetText vasList, Mid(InfoSample, 2, 1), lRow, gMaxCol
'        SetText vasList, Mid(InfoSample, 4, 1), lRow, gMaxCol + 1
'        SetText vasList, Mid(InfoSample, 6, 1), lRow, gMaxCol + 2
        
        SQL = " delete from pat_resmemo  " & vbCrLf & _
              " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & gID1 & "' "
        res = SendQuery(gLocal, SQL)
        
        SQL = "Select barcode, SlideOrd from res_flag " & vbCrLf & _
              "where examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "'  " & vbCrLf & _
              "  and barcode = '" & gID1 & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = gID1 Then
            lsSlideOrd = Trim(gReadBuf(1))
        Else
            lsSlideOrd = ""
        End If
        
        Select Case SampleJudg
        Case "0"
            SetText vasList, "Negative", lRow, gMaxCol + 3
            SetBackColor vasList, lRow, lRow, 2, 2, 255, 255, 255
        Case "1"
            SetText vasList, "Positive", lRow, gMaxCol + 3
            SetBackColor vasList, lRow, lRow, 2, 2, 255, 224, 193
            If lsSlideOrd = "SC" Then
                liMustBe = 1
            End If
        Case "2"
            SetText vasList, "Error", lRow, gMaxCol + 3
            SetBackColor vasList, lRow, lRow, 2, 2, 255, 224, 193
        Case "3"
            SetText vasList, "Potive+Error", lRow, gMaxCol + 3
            SetBackColor vasList, lRow, lRow, 2, 2, 255, 224, 193
            If lsSlideOrd = "SC" Then
                liMustBe = 1
            End If
        Case "4"
            SetText vasList, "QC Sample", lRow, gMaxCol + 3
            SetBackColor vasList, lRow, lRow, 2, 2, 255, 255, 255
        End Select
        
        If lsSlideOrd = "SP" Then
            liMustBe = 1
        End If
        
'        If Mid(InfoSample, 2, 1) <> "0" Then
'            SetBackColor vasList, lRow, lRow, flagWBC, flagWBC, 193, 234, 255
'        Else
'            SetBackColor vasList, lRow, lRow, flagWBC, flagWBC, 255, 255, 255
'        End If
'        If Mid(InfoSample, 4, 1) <> "0" Then
'            SetBackColor vasList, lRow, lRow, flagRBC, flagRBC, 193, 234, 255
'        Else
'            SetBackColor vasList, lRow, lRow, flagRBC, flagRBC, 255, 255, 255
'        End If
'        If Mid(InfoSample, 6, 1) <> "0" Then
'            SetBackColor vasList, lRow, lRow, flagPLT, flagPLT, 193, 234, 255
'        Else
'            SetBackColor vasList, lRow, lRow, flagPLT, flagPLT, 255, 255, 255
'        End If
                
        SetText vasList, gsVersion, lRow, 10
        
        '환자정보 불러오기
        GetPatientInfo2 gID1, lRow
        
        ClearSpread vasExam2
        
        If IsNumeric(gID1) = False Or Len(gID1) < 11 Then
        Else
            '검사항목 불러오기
'            res = Get_Order(gID1)
'            For i = 0 To UBound(gOrder_List)
'                vasExam2.SetText 1, i + 1, gOrder_List(i).TST_CD
'            Next i
            'res = Online_XML1(gXml_S07, Trim(gID1))
            res = Online_XML1(gXml_S08, Trim(gID1))
            For i = 0 To UBound(gExam_Select1)
                vasExam2.SetText 1, i + 1, gExam_Select1(i).TST_CD
            Next i
        End If
        
        If liMustBe = 1 Then
            SetText vasList, liMustBe, lRow, gMaxCol + 5
            SetBackColor vasList, lRow, lRow, 9, 9, 255, 224, 193
            'SetForeColor vasList, lRow, lRow, 9, 9, 0, 0, 160
'                vasList.Row = lRow
'                vasList.Col = 9
'                vasList.FontBold = True
        Else
            SetText vasList, "", lRow, gMaxCol + 5
            SetBackColor vasList, lRow, lRow, 9, 9, 255, 255, 255
            'SetForeColor vasList, lRow, lRow, 9, 9, 0, 0, 0
'                vasList.Row = lRow
'                vasList.Col = 9
'                vasList.FontBold = False
        End If
        
        gCurRow1 = lRow
        
        SQL = "Delete from res_flag " & vbCrLf & _
              "where examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "'  " & vbCrLf & _
              "  and barcode = '" & gID1 & "' "
        res = SendQuery(gLocal, SQL)
        
        SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
              "Values ('" & Format(CDate(lsCurDate), "yyyymmdd") & "', '" & gID1 & "', '" & SampleJudg & "', '" & PosDif & "', '" & PosMor & "', '" & PosCnt & "', " & _
              "'" & ErrFun & "', '" & ErrRes & "', '" & Left(InfoSample, 1) & "', '" & Mid(InfoSample, 2, 1) & "', '" & Mid(InfoSample, 3, 1) & "', '" & Mid(InfoSample, 4, 1) & "', " & _
              "'" & Mid(InfoSample, 5, 1) & "', '" & Mid(InfoSample, 6, 1) & "', '" & InfoWBC & "', '" & InfoPLT & "', '" & liMustBe & "', '" & lsSlideOrd & "' ) "
        res = SendQuery(gLocal, SQL)
                
        SQL = " delete From pat_resmemo " & vbCrLf & _
              " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' "
        res = SendQuery(gLocal, SQL)
                
        gRow1 = lRow
        
    Case "R|"
        gMsgFlag1 = "R"
        
        iCnt = 0
        i = InStr(1, lsData, "|")
        Do While i > 0
            iCnt = iCnt + 1
            lsTmp = Left(lsData, i - 1)
            lsData = Mid(lsData, i + 1)
            
            Select Case iCnt
            Case 3
                lsTmp = Mid(lsTmp, 5)
                j = InStr(1, lsTmp, "^")
                If j > 0 Then
                    lsEquip = Trim(Left(lsTmp, j - 1))
                Else
                    lsEquip = Trim(lsTmp)
                End If
            Case 4
                lsResult = Trim(lsTmp)
            Case 5
                lsUnit = Trim(lsTmp)
            Case 6
            Case 7
                lsFlag = Trim(lsTmp)
                
                If lsEquip = "WBC" Then lsWBC2 = lsResult
                If lsEquip = "NEUT%" Then lsNEUT2 = lsResult
                
                lResRow = vasRes_2.DataRowCnt + 1
                vasRes_2.MaxRows = lResRow
                vasRes_2.SetText 1, lResRow, lsEquip
                vasRes_2.SetText 2, lResRow, lsResult
                vasRes_2.SetText 3, lResRow, lsFlag
                vasRes_2.SetText 4, lResRow, lsUnit
                
                Exit Do
            End Select
            
            i = InStr(1, lsData, "|")
        Loop
        
        Set_Sample_Result_2 lResRow
        
        If IsNumeric(lsWBC2) And IsNumeric(lsNEUT2) Then
            lsWBC2 = Format(CCur(lsWBC2) * CCur(lsNEUT2) * 0.01, "#0.00")
            
            lResRow = vasRes_2.DataRowCnt + 1
            vasRes_2.MaxRows = lResRow
            vasRes_2.SetText 1, lResRow, "ANC"
            vasRes_2.SetText 2, lResRow, lsWBC2
            vasRes_2.SetText 3, lResRow, ""
            vasRes_2.SetText 4, lResRow, ""
            
            Set_Sample_Result_2 lResRow
            
            lsWBC2 = ""
            lsNEUT2 = ""
        End If
        
        vasList.Row = gRow1
        vasList.Col = 1
        If vasList.Value = 0 Then
            SetText vasList, "수신", gRow1, gResCol
        End If
        
    Case "L|"
        
        liC_WBC = -1
        liEOSIN = -1
        
        vasList.Row = gRow1
        vasList.Col = 1
        If chkMode.Value = 1 And gMsgFlag1 = "R" Then
'            SaveData gRow1 & " : " & gID1
            res = ToServer1(gRow1, vasList)
            If res = 1 Then
                vasList.Row = gRow1
                vasList.Col = 1
                vasList.Value = 0

                SetText vasList, "완료", gRow1, gResCol
                SetBackColor vasList, gRow1, gRow1, 1, 1, 202, 255, 112
                'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
            
                SQL = "update pat_res set sendflag = 'B' " & vbCrLf & _
                      "WHERE examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "' " & vbCrLf & _
                      "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                      "  AND barcode = '" & gID1 & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then SaveQuery SQL
                
                SQL = "Update worklist set OrdFlag = 'D' where barcode = '" & gID1 & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then SaveQuery SQL
                
'                SQL = "select sid from cpslipuse where sid = '" & gID1 & "' "
'                res = db_select_Col(gServer, SQL)
'                If Trim(gReadBuf(0)) <> Trim(gID1) Then
'                    SetBackColor vasList, gRow1, gRow1, 1, 1, 255, 0, 0
'                    SetForeColor vasList, gRow1, gRow1, gResCol - 3, gResCol, 255, 0, 0
'                    vasList.SetText gResCol, lRow, "미접수"
'                End If
            ElseIf res = 2 Then
                vasList.Row = gRow1
                vasList.Col = 1
                vasList.Value = 1
                
                'SetText vasList, "결과", lRow, gResCol
            Else
                vasList.Row = gRow1
                vasList.Col = 1
                vasList.Value = 1
            
                SQL = "Update worklist set OrdFlag = 'E' where barcode = '" & gID1 & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then SaveQuery SQL
                
                SQL = "update pat_res set sendflag = 'E' " & vbCrLf & _
                      "WHERE examdate = '" & Format(CDate(lsCurDate), "yyyymmdd") & "' " & vbCrLf & _
                      "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                      "  AND barcode = '" & gID1 & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then SaveQuery SQL
                
                SetText vasList, "실패", gRow1, gResCol
                SetForeColor vasList, gRow1, gRow1, 1, 1, 255, 0, 0
                'SetBackColor vasID, gPreRow, gPreRow, colCheckBox, vasID.MaxCols, 255, 255, 255
                'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
            End If
        Else
            SQL = "Update worklist set OrdFlag = 'C' where barcode = '" & gID1 & "' "
            res = SendQuery(gLocal, SQL)
            If res = -1 Then SaveQuery SQL
        End If
        
    Case "Q|"
        gMsgFlag1 = "Q"
        
        iCnt = 0
        i = InStr(1, lsData, "|")
        Do While i > 0
            iCnt = iCnt + 1
            lsTmp = Left(lsData, i - 1)
            lsData = Mid(lsData, i + 1)
            
            If iCnt = 3 Then
                i = InStr(1, lsTmp, "^")
                If i > 0 Then
                    gRack1 = Trim(Left(lsTmp, i - 1))
                    lsTmp = Mid(lsTmp, i + 1)
                    
                    i = InStr(1, lsTmp, "^")
                    If i > 0 Then
                        gPos1 = Trim(Left(lsTmp, i - 1))
                        lsTmp = Mid(lsTmp, i + 1)
                        
                        i = InStr(1, lsTmp, "^")
                        If i > 0 Then
                            gID1 = Trim(Left(lsTmp, i - 1))
                            gFlag1 = Mid(lsTmp, i + 1)
'                            If IsNumeric(gID) = True And Len(gID) = 9 Then
'                                gID = "0" & gID
'                            End If
                        Else
                            gID1 = Trim(lsTmp)
                            gFlag1 = ""
                        End If
                    End If
                End If
                
                Exit Do
            End If
            
            i = InStr(1, lsData, "|")
        Loop
        
        
    End Select
End Sub

Function Set_Sample_Result_2(asRow As Long) As Long
    Dim lRow As Long
    Dim i, j, k, z As Long
    Dim lsEquip As String
    Dim lsResult As String
    Dim lsFlagRes As String
    
    Dim lsFlag As String
    Dim lsUnit As String
    Dim liEquipCode As Integer
    
    Dim lsMsg As String
    
    Set_Sample_Result_2 = -1
    
    If asRow < 1 Or asRow > vasRes_2.DataRowCnt Then Exit Function
    
    lRow = gRow1
    
    lsEquip = Trim(GetText(vasRes_2, asRow, 1))
    lsResult = Trim(GetText(vasRes_2, asRow, 2))
    
    lsFlag = Trim(GetText(vasRes_2, asRow, 3))
    lsUnit = Trim(GetText(vasRes_2, asRow, 4))
    
    If lsFlag = "A" Then
        lsFlagRes = ""
        If lsResult <> "" Then
            lsFlagRes = "[" & lsResult & "]"
        End If
    
        lsMsg = ""
        
        Select Case lsEquip
        'WBC*************************
        Case "WBC_Abn_Scattergram"
            lsMsg = "WBC abnormal scatergram" & lsFlagRes
        Case "NRBC_Abn_Scattergram"
            lsMsg = "NRBC Abn Scg" & lsFlagRes
        Case "Neutropenia"
            lsMsg = "Neutro-" & lsFlagRes
        Case "Neutrophilia"
            lsMsg = "Neutro+" & lsFlagRes
        Case "Lymphopenia"
            lsMsg = "Lympho-" & lsFlagRes
        Case "Lymphocytosis"
            lsMsg = "Lympho+" & lsFlagRes
        Case "Monocytosis"
            lsMsg = "Mono+" & lsFlagRes
        Case "Eosinophilia"
            lsMsg = "Eo+" & lsFlagRes
        Case "Basophilia"
            lsMsg = "Baso+" & lsFlagRes
        Case "Leukocytopenia"
            lsMsg = "Leuko-" & lsFlagRes
        Case "Leukocytosis"
            lsMsg = "Leuko+" & lsFlagRes
        Case "NRBC_Present"
            lsMsg = lsEquip & lsFlagRes
        Case "Blasts?"
            lsMsg = lsEquip & lsFlagRes
        Case "Left_Shift?"
            lsMsg = lsEquip & lsFlagRes
        Case "NRBC?"
            lsMsg = lsEquip & lsFlagRes
        Case "Immature_Gran?"
            lsMsg = "Immature granules" & lsFlagRes
        Case "Atypical_Lympho?"
            lsMsg = "Atypical lymphocytes" & lsFlagRes
        Case "Abn_Lympho/L_Blasts?"
            lsMsg = "Abn Ly/L_Bl?" & lsFlagRes
        Case "RBC_Lyse Resistance?"
            lsMsg = "RBC Lyse resistance" & lsFlagRes

        'RBC*************************
        Case "RBC_Abn_Distribution"
            lsMsg = "RBC Abn Dst" & lsFlagRes
        Case "Dimorphic_Population"
            lsMsg = "Dimorph Pop" & lsFlagRes
        Case "RET_Abn_Scattergram"
            lsMsg = "RET Abn Scg" & lsFlagRes
        Case "Reticulocytosis"
            lsMsg = "Reticulo" & lsFlagRes
        Case "Anisocytosis"
            lsMsg = "Aniso" & lsFlagRes
        Case "Microcytosis"
            lsMsg = "Micro" & lsFlagRes
        Case "Macrocytosis"
            lsMsg = "Macro" & lsFlagRes

        Case "Hypochromia", "Anemia", "HGB_Defect?", "Fragments?"
            lsMsg = lsEquip & lsFlagRes

        Case "Erythrocytosis"
            lsMsg = "Erythro+" & lsFlagRes
'
        Case "RBC_Agglutination?"
            lsMsg = "RBC agglutination" & lsFlagRes
        Case "Turbidity/HGB Interference?"
            lsMsg = "Turbidty/Hb interface" & lsFlagRes
        Case "Iron_Deficiency?"
            lsMsg = "Iron Def?" & lsFlagRes

        'PLT*************************
        Case "PLT_Abn_Scattergram"
            lsMsg = "PLT Abn Scg" & lsFlagRes
        Case "PLT_Abn_Distribution"
            lsMsg = "PLT Abn Dst" & lsFlagRes
        Case "Thrombocytopenia"
            lsMsg = "Thrombo-" & lsFlagRes
        Case "Thrombocytosis"
            lsMsg = "Thrombo+" & lsFlagRes
        Case "PLT_Clumps?"
            lsMsg = "PLT Clumps" & lsFlagRes
        Case "PLT_Clumps(S)?"
            lsMsg = "PLT Clumps(S)" & lsFlagRes
            
        Case "Positive_Morph", "Positive_Count"
            lsMsg = lsEquip & lsFlagRes
            
            SQL = "select barcode from res_flag " & vbCrLf & _
                  "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "'  " & vbCrLf & _
                  "  and barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' "
            res = db_select_Col(gLocal, SQL)
            If Trim(gReadBuf(0)) = Trim(GetText(vasList, lRow, 2)) Then
                SQL = "Update res_flag set SampleJudg = '1'  " & vbCrLf & _
                      "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "'  " & vbCrLf & _
                      "  and barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' "
                res = SendQuery(gLocal, SQL)
            Else
                SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
                      "Values ('" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "', '" & Trim(GetText(vasList, lRow, 2)) & "', '1', '', '', '', " & _
                      "'', '', '', '', '', '', " & _
                      "'', '', '', '', '', '' ) "
                res = SendQuery(gLocal, SQL)
            End If
        Case Else   '2010.03.20 이상은 추가
            lsMsg = lsEquip
            
        End Select
        
        '메모결과 입력
        If lsMsg <> "" Then Save_ResMemo_2 lRow, lsMsg
        
    End If
    
    SQL = "Update res_flag set SampleJudg = '1'  " & vbCrLf & _
          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "'  " & vbCrLf & _
          "  and barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' " & vbCrLf & _
          "  and PBSFlag = '1' "
    res = SendQuery(gLocal, SQL)
    'gArrExam
    'EquipCode,ExamCode, ExamName, Seqno, PointSize,
    'RefLow, RefHigh, RSGubun
    
    k = UBound(gArrExamRes2)
    k = k + 1
    ReDim Preserve gArrExamRes2(k)
    
    i = 1
    For i = 1 To UBound(gArrExam)
        If Trim(gArrExam(i, 3)) = lsEquip Then
            z = -1
            If Trim(GetText(vasList, lRow, gResCol)) = "미접수" Then
                z = 1
                
                gArrExamRes2(k).EquipCode = gArrExam(i, 1)
                gArrExamRes2(k).ExamCode = gArrExam(i, 2)
                gArrExamRes2(k).ExamNo = gArrExam(i, 9)
                gArrExamRes2(k).ExamName = gArrExam(i, 3)
                gArrExamRes2(k).SeqNo = gArrExam(i, 5)
                gArrExamRes2(k).RefLow = gArrExam(i, 6)
                gArrExamRes2(k).RefHigh = gArrExam(i, 7)
                gArrExamRes2(k).RefFlag = ""
                gArrExamRes2(k).res = lsResult
                gArrExamRes2(k).EquipGubun = "IPU2"
            
                SetResult2 k, i
                                                
                SetText vasList, gArrExamRes2(k).res, lRow, gArrExam(i, 10)
                'SetText vasList, gArrExamRes1(k).res, lRow, gResCol + i
                
                Select Case gArrExamRes2(k).RefFlag
                Case "H", ">", "A"
                    SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 255, 127, 0
                Case "L", "W"
                    SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 127, 255
                Case Else
                    SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 0, 0
                End Select
                
                Save_Local_One2 lRow, k, "A"
            End If
            
            For j = 1 To vasExam2.DataRowCnt
                If Trim(GetText(vasExam2, j, 1)) = gArrExam(i, 2) Then
                    z = 1
                    
                    gArrExamRes2(k).EquipCode = gArrExam(i, 1)
                    gArrExamRes2(k).ExamCode = gArrExam(i, 2)
                    gArrExamRes2(k).ExamNo = gArrExam(i, 9)
                    gArrExamRes2(k).ExamName = gArrExam(i, 3)
                    gArrExamRes2(k).SeqNo = gArrExam(i, 5)
                    gArrExamRes2(k).RefLow = gArrExam(i, 6)
                    gArrExamRes2(k).RefHigh = gArrExam(i, 7)
                    gArrExamRes2(k).RefFlag = ""
                    gArrExamRes2(k).res = lsResult
                    gArrExamRes2(k).EquipGubun = "IPU2"
                    
                    SetResult2 k, i
                                                    
                    SetText vasList, gArrExamRes2(k).res, lRow, gArrExam(i, 10)
                    Select Case gArrExamRes2(k).RefFlag
                    Case "H", ">", "A"
                        SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 255, 127, 0
                    Case "L", "W"
                        SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 127, 255
                    Case Else
                        SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 0, 0
                    End Select
                    
                    Save_Local_One2 lRow, k, "A"
                                           
                    DeleteRow vasExam2, j, j
                    
                    Exit For
                End If
            Next j
            If z = 1 Then
                Exit For
            End If
        End If
    Next i
                        
    If z = -1 Then
        z = 1
        
        For i = 1 To UBound(gArrExam)
            If Trim(gArrExam(i, 3)) = lsEquip Then
                    z = 1
                    
                    gArrExamRes2(k).EquipCode = gArrExam(i, 1)
                    gArrExamRes2(k).ExamCode = gArrExam(i, 2)
                    gArrExamRes2(k).ExamNo = gArrExam(i, 9)
                    gArrExamRes2(k).ExamName = gArrExam(i, 3)
                    gArrExamRes2(k).SeqNo = gArrExam(i, 5)
                    gArrExamRes2(k).RefLow = gArrExam(i, 6)
                    gArrExamRes2(k).RefHigh = gArrExam(i, 7)
                    gArrExamRes2(k).RefFlag = ""
                    gArrExamRes2(k).res = lsResult
                    gArrExamRes2(k).EquipGubun = "IPU2"
                
                    SetResult2 k, i
                                                    
                    SetText vasList, gArrExamRes2(k).res, lRow, gArrExam(i, 10)
                    'SetText vasList, gArrExamRes1(k).res, lRow, gResCol + i
                    
                    Select Case gArrExamRes2(k).RefFlag
                    Case "H", ">", "A"
                        SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 255, 127, 0
                    Case "L", "W"
                        SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 127, 255
                    Case Else
                        SetForeColor vasList, lRow, lRow, gArrExam(i, 10), gArrExam(i, 10), 0, 0, 0
                    End Select
                    
                    Save_Local_One2 lRow, k, "A"
                    
                    Exit For
            End If
        Next i
    End If
                        
End Function



Function Save_ResMemo_2(ByVal asRow As Long, asMessage As String)
'메시지 저장하기
    Dim sMessage As String
    
    If asMessage = "" Then
        Exit Function
    End If
    
    sMessage = ""
    
    SQL = " Select message From pat_resmemo " & vbCrLf & _
          " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(vasList, asRow, 2)) & "' "
    res = db_select_Col(gLocal, SQL)
    
    sMessage = Trim(gReadBuf(0))
    
    If sMessage = "" Then
        SQL = " Insert Into pat_resmemo (examdate, equipno, barcode, message) " & vbCrLf & _
              " VALUES ('" & Format(dtpExamDate.Value, "yyyymmdd") & "', '" & gEquip & "', " & vbCrLf & _
              "         '" & Trim(GetText(vasList, asRow, 2)) & "', '" & asMessage & "') "
    Else
        'sMessage = sMessage & vbCrLf & asMessage
        sMessage = sMessage & ", " & asMessage

        SQL = " Update pat_resmemo Set " & vbCrLf & _
              " message = '" & Trim(sMessage) & "' " & vbCrLf & _
              " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasList, asRow, 2)) & "' "
    End If
    
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function


Public Sub SaveRes(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    
    FilNum = FreeFile
    
    Open App.Path & "\Log\" & Format(dtpExamDate.Value, "yyyymmdd") & " .res" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

Private Sub Picture1_Click()
    frmUser.Show 0
End Sub

Private Sub subChangeUser_Click()
    frmUserChange.Show 1
End Sub

Private Sub subClear_Click()
'vsSpread의 내용을 Clear 한다.
    vasList.Row = 1
    vasList.Col = 1
    vasList.Row2 = vasList.MaxRows
    vasList.Col2 = vasList.MaxCols
    vasList.BlockMode = True
    vasList.Action = 3
    vasList.BackColor = RGB(255, 255, 255)
    vasList.ForeColor = RGB(0, 0, 0)
    vasList.BlockMode = False

    vasList.Row = 1
    vasList.Col = 1
    vasList.Row2 = vasList.MaxRows
    vasList.Col2 = 1
    vasList.BlockMode = True
    vasList.Value = 1
    vasList.BlockMode = False

    txtBuff1 = ""
    txtBuff2 = ""
    
    gCurRow = -1
    ReDim gArrExamRes(0)
    GetExamCode
    
    GetOPtion
End Sub

Private Sub subClose_Click()
    Unload Me
End Sub

Private Sub subCodeSet_Click()
    frmCode.Show 1
End Sub

Private Sub subComConnect_Click()
    frmConnect.Show 1
    
    If IPU1.ConnectFlag Then
        If MSComm1.PortOpen = False Then
            MSComm1.CommPort = IPU1.ComPort
            MSComm1.Settings = IPU1.Speed & "," & IPU1.Parity & "," & IPU1.DataBit & "," & IPU1.StartBit
            If IPU1.RTSEnable = "1" Then
                MSComm1.RTSEnable = True
            Else
                MSComm1.RTSEnable = False
            End If
            If IPU1.DTREnable = "1" Then
                MSComm1.DTREnable = True
            Else
                MSComm1.DTREnable = False
            End If
            MSComm1.PortOpen = True
            
            If MSComm1.CTSHolding = True Then
                lblIPU1.ForeColor = RGB(0, 255, 0)
            Else
                lblIPU1.ForeColor = RGB(0, 0, 255)
            End If
        End If
    Else
        If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
        lblIPU1.ForeColor = RGB(255, 0, 0)
    End If

    If IPU2.ConnectFlag Then
        If MSComm2.PortOpen = False Then
            MSComm2.CommPort = IPU2.ComPort
            MSComm2.Settings = IPU2.Speed & "," & IPU2.Parity & "," & IPU2.DataBit & "," & IPU2.StartBit
            If IPU2.RTSEnable = "1" Then
                MSComm2.RTSEnable = True
            Else
                MSComm2.RTSEnable = False
            End If
            If IPU2.DTREnable = "1" Then
                MSComm2.DTREnable = True
            Else
                MSComm2.DTREnable = False
            End If
            MSComm2.PortOpen = True
        
            If MSComm2.CTSHolding = True Then
                lblIPU2.ForeColor = RGB(0, 255, 0)
            Else
                lblIPU2.ForeColor = RGB(0, 0, 255)
            End If
        End If
    Else
        If MSComm2.PortOpen = True Then MSComm2.PortOpen = False
        lblIPU2.ForeColor = RGB(255, 0, 0)
    End If

End Sub

Private Sub subComparison_Click()

End Sub

Private Sub subComSetup_Click()
    frmConfig.Show 1
End Sub

Private Sub subLASCOrder_Click()
    Shell App.Path & "\LASCOrder.exe"
End Sub

Private Sub subProficiency_Click()
    frmProficiency.Show
End Sub

Private Sub subSchESR_Click()
    vasSch.Visible = False
    cmdCol2.Visible = False
    vasSchESR.Visible = True
    cmdPrint.Visible = True
    cmdSch.Visible = False
    ESR_Search
End Sub

Private Sub subSearch_Click()
    vasSch.Visible = True
    cmdCol2.Visible = True
    vasSchESR.Visible = False
    cmdSch.Visible = True
    
    cmdSch_Click
End Sub

Private Sub subSend1_Click()
    subSend1.Checked = True
    subSend2.Checked = False
    
    chkMode.Value = 1
    SaveSetting "MEDIMATE", "XE2100", "SendMode", "1"
End Sub

Private Sub subSend2_Click()
    subSend1.Checked = False
    subSend2.Checked = True
    
    chkMode.Value = 0
    SaveSetting "MEDIMATE", "XE2100", "SendMode", "0"

End Sub

Private Sub subSendQC_Click()
'    frmResult.Show 1
End Sub

Private Sub subWorkList_Click()
    'Shell App.Path & "\workList.exe"
End Sub


Private Sub Timer1_Timer()
'    dtpExamDate.Value = Format(Date, "yyyy-mm-dd")
'    sspTime.Caption = Format(Time, "hh:nn")
    
    If IPU1.ConnectFlag Then
        If MSComm1.CTSHolding = True Then
            lblIPU1.ForeColor = RGB(0, 0, 255)
        Else
            lblIPU1.ForeColor = RGB(0, 255, 0)
        End If
    End If
    
    If IPU2.ConnectFlag Then
        If MSComm2.CTSHolding = True Then
            lblIPU2.ForeColor = RGB(0, 0, 255)
        Else
            lblIPU2.ForeColor = RGB(0, 255, 0)
        End If
    End If
End Sub

Private Sub txtBarcode_GotFocus()
    SelectFocus txtBarcode
End Sub

Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    Dim iSch As Integer
    
    If KeyCode = vbKeyReturn Then
        iSch = -1
        For lRow = 1 To vasList.DataRowCnt
            If Trim(GetText(vasList, lRow, 2)) = Trim(txtBarcode) Then
                vasActiveCell vasList, lRow, 2
                iSch = 1
                Exit For
            End If
        Next lRow
        If iSch = -1 Then
            SearchSample Trim(txtBarcode)
        End If
    End If
End Sub

Private Sub txtBuff1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        XE2100_ASTM_1
        txtBuff1 = ""
    End If
End Sub

Private Sub txtBuff2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        XE2100_2
    End If
End Sub

Private Sub vasList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    iRow1 = BlockRow
    iRow2 = BlockRow2
    iCol1 = BlockCol
    iCol2 = BlockCol2
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        iRow1 = 0
        iRow2 = 0

        Select Case Col
        Case 0, 1
            Exit Sub
        Case 2
            vasSort vasList, Col
        Case 3
            vasSort vasList, Col
        Case 4
            vasSort vasList, Col, 2
        Case 5
            vasSort vasSch, Col, 6
        Case 9
            vasSort vasList, Col
        Case 10
            vasSort vasSch, 10, 11, 12, 2
        Case 11, 12
            vasSort vasSch, 11, 12, 2
        'Case 6, 7, 8, gMaxCol, gMaxCol + 1, gMaxCol + 2, gMaxCol + 3, gMaxCol + 5
        '    vasSort vasList, Col, 9
        Case Else
            vasSort vasList, Col
        End Select

    Else
        iRow1 = Row
        iRow2 = Row
    End If
    
End Sub

Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim argSpread As vaSpread
    Dim lRow As Long
    Dim lCol, i As Long
    
    If Row < 1 Or Row > vasList.DataRowCnt Then Exit Sub
    
    SelVas = 1
    
    txtID = ""
    txtPID = ""
    txtPName = ""
    txtResDate = ""
    txtEquip = ""
    
    ClearSpread vasRes1
    ClearSpread vasRes2
    
    txtID = Trim(GetText(vasList, Row, 2))
    txtPID = Trim(GetText(vasList, Row, 3))
    txtPName = Trim(GetText(vasList, Row, 4))
    txtSexAge = Trim(GetText(vasList, Row, 6)) & "/" & Trim(GetText(vasList, Row, 7))
    txtWardRoom = Trim(GetText(vasList, Row, 8))
    txtWorkListNo = Trim(GetText(vasList, Row, 9))
    txtRack = Trim(GetText(vasList, Row, 11))
    txtTube = Trim(GetText(vasList, Row, 12))
    Select Case Trim(GetText(vasList, Row, 10))
    Case "IPU1"
        txtEquip = "XE2100-1"
    Case "IPU2"
        txtEquip = "XE2100-2"
    End Select
    
    FlagComment Trim(GetText(vasList, Row, gMaxCol)), Trim(GetText(vasList, Row, gMaxCol + 1)), Trim(GetText(vasList, Row, gMaxCol + 2)), Trim(GetText(vasSch, Row, gMaxCol + 3))
    
    SQL = "Select resdate from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND barcode = '" & Trim(txtID) & "' "
    res = db_select_Text(gLocal, SQL, txtResDate)
    
    'lCol = gResCol
    lRow = 0
    Set argSpread = vasRes1
    For lCol = gResCol + 1 To gMaxCol - 1
        If Trim(GetText(vasList, Row, lCol)) <> "" Then
            lRow = lRow + 1
            
            If lRow = 21 Then
                lRow = 1
                Set argSpread = vasRes2
            End If
            
            For i = LBound(gArrExam) To UBound(gArrExam)
                If lCol = gArrExam(i, 10) Then
                    SetText argSpread, gArrExam(i, 1), lRow, 1
                    Exit For
                End If
            Next i
            
            'SetText argSpread, gArrExam(lCol - gResCol, 1), lRow, 1
            SetText argSpread, Trim(GetText(vasList, Row, lCol)), lRow, 3
            SetText argSpread, Trim(GetText(vasList, 0, lCol)), lRow, 2
            
            vasList.Row = Row
            vasList.Col = lCol
            Select Case vasList.ForeColor
            Case RGB(255, 127, 0)
                SetForeColor argSpread, lRow, lRow, 4, 4, 255, 127, 0
                SetText argSpread, "▲", lRow, 4
            Case RGB(0, 127, 255)
                SetForeColor argSpread, lRow, lRow, 4, 4, 0, 127, 255
                SetText argSpread, "▼", lRow, 4
            Case Else
                SetText argSpread, "", lRow, 4
            End Select
        
        End If
    Next lCol
    
    SQL = " Select message From pat_resmemo " & vbCrLf & _
          " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(txtID) & "' "
    res = db_select_Col(gLocal, SQL)
    If Trim(txtFlag) = "" Then
        txtFlag = Trim(gReadBuf(0))
    Else
        txtFlag = txtFlag & vbCrLf & Trim(gReadBuf(0))
    End If
    
'    lCol = gResCol
'    For lRow = 1 To 20
'        lCol = lCol + 1
'
'        If Trim(GetText(vasList, Row, lCol)) <> "" Then
'
'            SetText vasRes1, gArrExam(lCol - gResCol, 1), lRow, 1
'            SetText vasRes1, Trim(GetText(vasList, Row, lCol)), lRow, 3
'            SetText vasRes1, Trim(GetText(vasList, 0, lCol)), lRow, 2
'
'            vasList.Row = Row
'            vasList.Col = lCol
'            Select Case vasList.ForeColor
'            Case RGB(255, 127, 0)
'                SetForeColor vasRes1, lRow, lRow, lCol, lCol, 255, 127, 0
'                SetText vasRes1, "▲", lRow, 4
'            Case RGB(0, 127, 255)
'                SetForeColor vasRes1, lRow, lRow, lCol, lCol, 0, 127, 255
'                SetText vasRes1, "▼", lRow, 4
'            Case Else
'                SetText vasRes1, "", lRow, 4
'            End Select
'        End If
'    Next lRow
'
'    For lRow = 1 To 15
'        lCol = lCol + 1
'
'        If Trim(GetText(vasList, Row, lCol)) <> "" Then
'
'            SetText vasRes2, gArrExam(lCol - gResCol, 1), lRow, 1
'            SetText vasRes2, Trim(GetText(vasList, Row, lCol)), lRow, 3
'            SetText vasRes2, Trim(GetText(vasList, 0, lCol)), lRow, 2
'
'            vasList.Row = Row
'            vasList.Col = lCol
'            Select Case vasList.ForeColor
'            Case RGB(255, 127, 0)
'                SetForeColor vasRes2, lRow, lRow, lCol, lCol, 255, 127, 0
'                SetText vasRes2, "▲", lRow, 4
'            Case RGB(0, 127, 255)
'                SetForeColor vasRes2, lRow, lRow, lCol, lCol, 0, 127, 255
'                SetText vasRes2, "▼", lRow, 4
'            Case Else
'                SetText vasRes2, "", lRow, 4
'            End Select
'        End If
'    Next lRow
    
    Frame1.Visible = True
    
End Sub

Sub GetComSetup1()
    Dim db_tmp As String * 20
    Dim i As Integer
    Dim lRow As Long
       
    lRow = 0
    For i = 1 To 4
        db_tmp = ""
        Call GetPrivateProfileString("COM " & CStr(i), "Use", "", db_tmp, 20, App.Path & "\Interface.ini")
        txtTemp = Trim(db_tmp)
        If Trim(txtTemp) <> "" Then
'            lRow = lRow + 1
'
'            vasComList.Row = lRow
'            vasComList.Col = 1
'            If Trim(txtTemp) = "1" Then
'                vasComList.Value = 1
'            Else
'                vasComList.Value = 0
'            End If
            
            db_tmp = ""
            Call GetPrivateProfileString("COM " & CStr(i), "Gubun", "", db_tmp, 20, App.Path & "\Interface.ini")
            txtTemp = Trim(db_tmp)
            
            If Left(Trim(txtTemp), 3) = "IPU" Then
                lRow = lRow + 1
                
                SetText vasComList, Trim(txtTemp), lRow, 2
                
                vasComList.Row = lRow
                vasComList.Col = 1
                If Trim(txtTemp) = "IPU1" Then
                    If frmInterface.MSComm1.PortOpen = True Then
                        vasComList.Value = 1
                    Else
                        vasComList.Value = 0
                    End If
                ElseIf Trim(txtTemp) = "IPU2" Then
                    If frmInterface.MSComm2.PortOpen = True Then
                        vasComList.Value = 1
                    Else
                        vasComList.Value = 0
                    End If
                End If
                
                SetText vasComList, CStr(i), lRow, 3
                
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "Speed", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 4
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "Parity", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 5
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "DataBit", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 6
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "StopBit", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 7
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "RTSEnable", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 8
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "DTREnable", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 9
            End If
        End If
    Next i

    For i = 1 To vasComList.DataRowCnt
        If Trim(GetText(vasComList, i, 2)) = "IPU1" Then
            IPU1.ComPort = Trim(GetText(vasComList, i, 3))
            IPU1.Speed = Trim(GetText(vasComList, i, 4))
            IPU1.Parity = Trim(GetText(vasComList, i, 5))
            IPU1.DataBit = Trim(GetText(vasComList, i, 6))
            IPU1.StopBit = Trim(GetText(vasComList, i, 7))
            IPU1.RTSEnable = Trim(GetText(vasComList, i, 8))
            IPU1.DTREnable = Trim(GetText(vasComList, i, 9))
            IPU1.ConnectFlag = True
        ElseIf Trim(GetText(vasComList, i, 2)) = "IPU2" Then
            IPU2.ComPort = Trim(GetText(vasComList, i, 3))
            IPU2.Speed = Trim(GetText(vasComList, i, 4))
            IPU2.Parity = Trim(GetText(vasComList, i, 5))
            IPU2.DataBit = Trim(GetText(vasComList, i, 6))
            IPU2.StopBit = Trim(GetText(vasComList, i, 7))
            IPU2.RTSEnable = Trim(GetText(vasComList, i, 8))
            IPU2.DTREnable = Trim(GetText(vasComList, i, 9))
            IPU2.ConnectFlag = True
        End If
    Next i

End Sub

Sub GetComSetup()
    Dim db_tmp As String * 20
    Dim i As Integer
    Dim lRow As Long
       
    lRow = 0
    For i = 1 To 10
'        If Trim(txtTemp) <> "" Then
            
            db_tmp = ""
            Call GetPrivateProfileString("COM " & CStr(i), "Gubun", "", db_tmp, 20, App.Path & "\Interface.ini")
            txtTemp = Trim(db_tmp)
            
            If Left(Trim(txtTemp), 3) = "IPU" Then
                lRow = lRow + 1
                
                SetText vasComList, Trim(txtTemp), lRow, 2
                
                vasComList.Row = lRow
                vasComList.Col = 1
                If Trim(txtTemp) = "IPU1" Then
                    If frmInterface.MSComm1.PortOpen = True Then
                        vasComList.Value = 1
                    Else
                        vasComList.Value = 0
                    End If
                ElseIf Trim(txtTemp) = "IPU2" Then
                    If frmInterface.MSComm2.PortOpen = True Then
                        vasComList.Value = 1
                    Else
                        vasComList.Value = 0
                    End If
                End If
                
                SetText vasComList, CStr(i), lRow, 3
                
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "Use", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 1
                
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "Speed", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 4
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "Parity", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 5
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "DataBit", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 6
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "StopBit", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 7
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "RTSEnable", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 8
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "DTREnable", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 9
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "Protocol", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 10
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "Equip_CD", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 11
            
            End If
'        End If
    Next i
    
    For i = 1 To vasComList.DataRowCnt
        If Trim(GetText(vasComList, i, 2)) = "IPU1" Then
            IPU1.UseEquip = Trim(GetText(vasComList, i, 1))
            IPU1.ComPort = Trim(GetText(vasComList, i, 3))
            IPU1.Speed = Trim(GetText(vasComList, i, 4))
            IPU1.Parity = Trim(GetText(vasComList, i, 5))
            IPU1.DataBit = Trim(GetText(vasComList, i, 6))
            IPU1.StopBit = Trim(GetText(vasComList, i, 7))
            IPU1.RTSEnable = Trim(GetText(vasComList, i, 8))
            IPU1.DTREnable = Trim(GetText(vasComList, i, 9))
            IPU1.ConnectFlag = True
            IPU1.Protocol = Trim(GetText(vasComList, i, 10))
            IPU1.Equip_CD = Trim(GetText(vasComList, i, 11))
        ElseIf Trim(GetText(vasComList, i, 2)) = "IPU2" Then
            IPU2.UseEquip = Trim(GetText(vasComList, i, 1))
            IPU2.ComPort = Trim(GetText(vasComList, i, 3))
            IPU2.Speed = Trim(GetText(vasComList, i, 4))
            IPU2.Parity = Trim(GetText(vasComList, i, 5))
            IPU2.DataBit = Trim(GetText(vasComList, i, 6))
            IPU2.StopBit = Trim(GetText(vasComList, i, 7))
            IPU2.RTSEnable = Trim(GetText(vasComList, i, 8))
            IPU2.DTREnable = Trim(GetText(vasComList, i, 9))
            IPU2.Protocol = Trim(GetText(vasComList, i, 10))
            IPU2.Equip_CD = Trim(GetText(vasComList, i, 11))
            IPU2.ConnectFlag = True
        End If
    Next i

End Sub

Private Sub vasList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow, i, liEquipCode As Long
    Dim lsID As String
    Dim liRet As Integer
    'Dim lsID As String
    Dim lsResult As String
    Dim mExam
    
    If KeyCode = vbKeyReturn Then
        lRow = vasList.ActiveRow
        
        SQL = "Select barcode, diskno, posno, examtype from pat_res where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' and barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = Trim(GetText(vasList, lRow, 2)) Then
            If MsgBox("입력하신 검체 [" & Trim(GetText(vasList, lRow, 2)) & "]는 장비 " & Trim(gReadBuf(3)) & "의 " & Trim(gReadBuf(1)) & " Rack " & Trim(gReadBuf(2)) & " Position 에서 검사한 것입니다 " & vbCrLf & _
                      " " & vbCrLf & _
                      "저장하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, "알림") = vbNo Then
                SetText vasList, Trim(GetText(vasList, lRow, gMaxCol + 4)), lRow, 2
                Exit Sub
            End If
        End If
        
        If MsgBox("결과를 전송하시겠습니까? ", vbCritical + vbYesNo + vbDefaultButton2, "알림") = vbYes Then
            
            lsID = Trim(GetText(vasList, lRow, 2))
            
            If GetPatientInfo1(lsID, lRow) = 1 Then
                
                SQL = "Update pat_res set " & vbCrLf & _
                      "  barcode = '" & lsID & "', " & vbCrLf & _
                      "  pid = '" & Trim(GetText(vasList, lRow, 3)) & "', " & vbCrLf & _
                      "  pname = '" & Trim(GetText(vasList, lRow, 4)) & "', " & vbCrLf & _
                      "  pjumin = '" & Trim(GetText(vasList, lRow, 5)) & "', " & vbCrLf & _
                      "  psex = '" & Trim(GetText(vasList, lRow, 6)) & "', " & vbCrLf & _
                      "  page1 = '" & Trim(GetText(vasList, lRow, 7)) & "', " & vbCrLf & _
                      "  WardRoom = '" & Trim(GetText(vasList, lRow, 8)) & "', " & vbCrLf & _
                      "  receno = '" & Trim(GetText(vasList, lRow, 9)) & "' " & vbCrLf & _
                      "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                      "  and barcode = '" & Trim(GetText(vasList, lRow, gMaxCol + 4)) & "'"
                res = SendQuery(gLocal, SQL)
                
'                SaveData lRow & " : " & lsID
                res = ToServer(lRow, vasList, 1)
                If res = 1 Then
                    vasList.Row = lRow
                    vasList.Col = 1
                    vasList.Value = 0
    
                    SetText vasList, "완료", lRow, gResCol
                    SetBackColor vasList, lRow, lRow, 1, 1, 202, 255, 112
                    'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
                
                    SQL = "Update worklist set OrdFlag = 'D' where barcode = '" & lsID & "' "
                    res = SendQuery(gLocal, SQL)
                
                ElseIf res = 2 Then
                    SetText vasList, "결과", lRow, gResCol
                Else
                    SetText vasList, "실패", lRow, gResCol
                    SetForeColor vasList, lRow, lRow, 1, 1, 255, 0, 0
                    'SetBackColor vasID, gPreRow, gPreRow, colCheckBox, vasID.MaxCols, 255, 255, 255
                    'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
                End If
            
            Else
                vasList.Row = lRow
                vasList.Col = 1
                vasList.Value = 1
                SetBackColor vasList, lRow, lRow, 1, 1, 255, 0, 0
                SetText vasList, "미접수", lRow, gResCol
                
                Exit Sub
            End If
            
        End If
    End If
End Sub

Private Sub vasList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim i As Long
    
    If iRow1 < 0 And iRow2 < 0 Then
        iRow1 = Row
        iRow2 = Row
    End If
    
    For i = iRow1 To iRow2
        vasList.Row = i
        vasList.Col = 1
        vasList.Value = 0
    Next i
    vasList.BlockMode = False

End Sub

Private Sub vasSch_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    iRow1 = BlockRow
    iRow2 = BlockRow2
    iCol1 = BlockCol
    iCol2 = BlockCol2
End Sub

Private Sub vasSch_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        iRow1 = 0
        iRow2 = 0
        
        'If Col > gResCol Then Exit Sub
        
        Select Case Col
        Case 0, 1
            Exit Sub
        Case 2
            vasSort vasSch, Col
        Case 3
            vasSort vasSch, Col
        Case 4
            vasSort vasSch, Col, 2
        'Case 5
        '    vasSort vasSch, Col, 6
        Case 9
            vasSort vasSch, Col
        Case 10
            vasSort vasSch, 10, 11, 12, 2
        Case 11, 12
            vasSort vasSch, 11, 12, 2
        'Case 1, 6, 7, 8, 10, gMaxCol, gMaxCol + 1, gMaxCol + 2, gMaxCol + 3, gMaxCol + 5
        '    vasSort vasSch, Col, 9
        Case Else
            vasSort vasSch, Col
        End Select
    End If
End Sub

Private Sub vasSch_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lRow, lCol, i As Long
    Dim argSpread As vaSpread
    
    If Row < 1 Or Row > vasSch.DataRowCnt Then Exit Sub
    
    SelVas = 2
    
    If Row = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
    ElseIf Row = vasSch.DataRowCnt Then
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
    Else
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
    End If
    
    If vasSch.DataRowCnt = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
    End If
    
    
    txtID = ""
    txtPID = ""
    txtPName = ""
    txtResDate = ""
    txtEquip = ""
    
    ClearSpread vasRes1
    ClearSpread vasRes2
    
    txtID = Trim(GetText(vasSch, Row, 2))
    txtPID = Trim(GetText(vasSch, Row, 3))
    txtPName = Trim(GetText(vasSch, Row, 4))
    txtSexAge = Trim(GetText(vasSch, Row, 6)) & "/" & Trim(GetText(vasSch, Row, 7))
    txtWardRoom = Trim(GetText(vasSch, Row, 8))
    txtWorkListNo = Trim(GetText(vasSch, Row, 9))
    txtRack = Trim(GetText(vasSch, Row, 11))
    txtTube = Trim(GetText(vasSch, Row, 12))
    Select Case Trim(GetText(vasSch, Row, 10))
    Case "IPU1"
        txtEquip = "XE2100-1"
    Case "IPU2"
        txtEquip = "XE2100-2"
    End Select
    
    FlagComment Trim(GetText(vasSch, Row, gMaxCol)), Trim(GetText(vasSch, Row, gMaxCol + 1)), Trim(GetText(vasSch, Row, gMaxCol + 2)), Trim(GetText(vasSch, Row, gMaxCol + 3))
    
    SQL = "Select resdate from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND barcode = '" & Trim(txtID) & "' "
    res = db_select_Text(gLocal, SQL, txtResDate)
    
    
    
    'lCol = gResCol
    lRow = 0
    Set argSpread = vasRes1
    For lCol = gResCol + 1 To gMaxCol - 1
        If Trim(GetText(vasSch, Row, lCol)) <> "" Then
            lRow = lRow + 1
'            If lRow <= 20 Then
'                Set argSpread = vasRes1
'            Else
'                Set argSpread = vasRes2
'            End If
            If lRow = 21 Then
                lRow = 1
                Set argSpread = vasRes2
            End If
            
            For i = LBound(gArrExam) To UBound(gArrExam)
                If lCol = gArrExam(i, 10) Then
                    SetText argSpread, gArrExam(i, 1), lRow, 1
                    Exit For
                End If
            Next i
            
            'SetText argSpread, gArrExam(lCol - gResCol, 1), lRow, 1
            SetText argSpread, Trim(GetText(vasSch, Row, lCol)), lRow, 3
            SetText argSpread, Trim(GetText(vasSch, 0, lCol)), lRow, 2
            
            vasSch.Row = Row
            vasSch.Col = lCol
            Select Case vasSch.ForeColor
            Case RGB(255, 127, 0)
                SetForeColor argSpread, lRow, lRow, 4, 4, 255, 127, 0
                SetText argSpread, "▲", lRow, 4
            Case RGB(0, 127, 255)
                SetForeColor argSpread, lRow, lRow, 4, 4, 0, 127, 255
                SetText argSpread, "▼", lRow, 4
            Case Else
                SetText argSpread, "", lRow, 4
            End Select
        
        End If
    Next lCol
    
    SQL = " Select message From pat_resmemo " & vbCrLf & _
          " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(txtID) & "' "
    res = db_select_Col(gLocal, SQL)
    If Trim(txtFlag) = "" Then
        txtFlag = Trim(gReadBuf(0))
    Else
        txtFlag = txtFlag & vbCrLf & Trim(gReadBuf(0))
    End If
    
'    For lRow = 1 To 20
'        lCol = lCol + 1
'
'        If Trim(GetText(vasSch, Row, lCol)) <> "" Then
'
'            SetText argSpread, gArrExam(lCol - gResCol, 1), lRow, 1
'            vasActiveCell vasSch, lRow, lCol
'            SetText argSpread, Trim(GetText(vasSch, Row, lCol)), lRow, 3
'            SetText argSpread, Trim(GetText(vasSch, 0, lCol)), lRow, 2
'
'            vasSch.Row = Row
'            vasSch.Col = lCol
'            Select Case vasSch.ForeColor
'            Case RGB(255, 127, 0)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 255, 127, 0
'                SetText argSpread, "▲", lRow, 4
'            Case RGB(0, 127, 255)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 0, 127, 255
'                SetText argSpread, "▼", lRow, 4
'            Case Else
'                SetText argSpread, "", lRow, 4
'            End Select
'        End If
'    Next lRow
'
'    For lRow = 1 To 15
'        lCol = lCol + 1
'
'        If Trim(GetText(vasSch, lRow, lCol)) <> "" Then
'
'            SetText argSpread, gArrExam(lCol - gResCol, 1), lRow, 1
'            SetText argSpread, Trim(GetText(vasSch, Row, lCol)), lRow, 3
'            SetText argSpread, Trim(GetText(vasSch, 0, lCol)), lRow, 2
'
'            vasSch.Row = Row
'            vasSch.Col = lCol
'            Select Case vasSch.ForeColor
'            Case RGB(255, 127, 0)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 255, 127, 0
'                SetText argSpread, "▲", lRow, 4
'            Case RGB(0, 127, 255)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 0, 127, 255
'                SetText argSpread, "▼", lRow, 4
'            Case Else
'                SetText argSpread, "", lRow, 4
'            End Select
'        End If
'    Next lRow
    
    Frame1.Visible = True

End Sub

Sub FlagComment(asWBC As String, asRBC As String, asPLT As String, Optional asJudge As String)
    Dim lsComment As String
    
    lsComment = ""
    
    Select Case asJudge
    Case "0"
        lsComment = "Negative" & vbCrLf & vbCrLf
    Case "1"
        lsComment = "Positive" & vbCrLf & vbCrLf
    Case "2"
        lsComment = "Error" & vbCrLf & vbCrLf
    Case "3"
        lsComment = "Positive+Error" & vbCrLf & vbCrLf
    Case "4"
        lsComment = "QC Sample" & vbCrLf & vbCrLf
    Case Else
        lsComment = "Sample : " & asJudge & vbCrLf & vbCrLf
    End Select
        
    If asWBC <> "0" And Trim(asWBC) <> "" Then lsComment = lsComment & "WBC SUSPECT Existing " & vbCrLf
'    Select Case asWBC
'    Case "1"
'        lsComment = lsComment & "Blasts?"
'    Case "2"
'        lsComment = lsComment & "Immature Grain?"
'    Case "3"
'        lsComment = lsComment & "Left Shift?"
'    Case "4"
'        lsComment = lsComment & "Atypical Lympho?"
'    Case "5"
'        lsComment = lsComment & "Abn Lympho/L_Blasts?"
'    Case "6"
'        lsComment = lsComment & "NRBC?"
'    Case "7"
'        lsComment = lsComment & "RBC Lyse Resistance?"
'    End Select
    
    If asRBC <> "0" And Trim(asRBC) <> "" Then lsComment = lsComment & "RBC SUSPECT Existing " & vbCrLf
'    Select Case asRBC
'    Case "1"
'        lsComment = lsComment & "RBC Agglutination?"
'    Case "2"
'        lsComment = lsComment & "Turbidity/HGB Interference?"
'    Case "3"
'        lsComment = lsComment & "Iron Deficiency?"
'    Case "4"
'        lsComment = lsComment & "HCG Defect?"
'    Case "5"
'        lsComment = lsComment & "Fragments?"
'    End Select
    
    If asPLT <> "0" And Trim(asPLT) <> "" Then lsComment = lsComment & "PLT SUSPECT Existing " & vbCrLf
'    Select Case asPLT
'    Case "1"
'        lsComment = lsComment & "PLT Clumps?"
'    Case "2"
'        lsComment = lsComment & "PLT Clumps(S)?"
'    End Select

    txtFlag = lsComment
End Sub

Sub ESR_Search()
    Dim lRow, lCol As Long
    Dim lsID As String
    Dim liEquipCode As Integer
    Dim i As Integer
    
    On Error GoTo ErrHandle
    
    ClearSpread vasSch, 0, 2
    
    Me.MousePointer = 11
        
    
'    SQL = "INSERT INTO pat_res (examdate, equipno, barcode, examtype, receno, pid, " & _
'          "pname, pjumin, page, psex, resdate, seqno, diskno, posno, " & _
'          "equipcode, examcode, examtype, result, sendflag, examname, " & _
'          "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
'          "VALUES ('" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'          "'" & Trim(GetText(vasList, asRow, 2)) & "','" & Trim(GetText(vasList, asRow, 5)) & "', '', " & _
'          "'" & Trim(GetText(vasList, asRow, 3)) & "', " & vbCrLf & _
'          "'" & Trim(GetText(vasList, asRow, 4)) & "', '', " & _
'          "0, '', " & _
'          "'" & sExamDate & "', '" & gArrExamRes(aiIndex).SeqNo & "', '" & Trim(GetText(vasList, asRow, 6)) & "', '" & Trim(GetText(vasList, asRow, 7)) & "', " & vbCrLf & _
'          "'" & gArrExamRes(aiIndex).EquipCode & "', '" & gArrExamRes(aiIndex).ExamCode & "', '', " & _
'          "'" & gArrExamRes(aiIndex).res & "', '" & asSend & "', '" & gArrExamRes(aiIndex).ExamName & "', " & vbCrLf & _
'          "'" & gArrExamRes(aiIndex).RefFlag & "', '', " & _
'          "'', '', " & _
'          "'" & gArrExamRes(aiIndex).RefLow & " ~ " & gArrExamRes(aiIndex).RefHigh & "', '' ) "
    
    SQL = "Select distinct a.barcode, a.pid, a.pname, a.pjumin, a.psex, a.page1, a.WardRoom, a.ReceNo, a.examtype, a.diskno, " & _
            "a.posno " & vbCrLf & _
          "from pat_res a " & vbCrLf & _
          "where a.examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.examcode = '99' " & vbCrLf & _
          "Order by a.barcode "
    
    'SQL = "Select a.barcode, a.pid, a.pname, a.examtype, a.diskno, " & _
            "a.posno, a.sendflag, a.equipcode, a.examcode, a.result, " & _
            "a.refflag, a.resdate, a.seqno, b.WBCSusp, b.RBCSusp, " & _
            "b.PLTSusp, b.SampleJudg  " & vbCrLf & _
          "from pat_res a,res_flag b" & vbCrLf & _
          "where a.examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and b.barcode = a.barcode " & vbCrLf & _
          "Order by a.barcode, a.equipcode "
    
    res = db_select_Vas(gLocal, SQL, vasSchESR, 1, 1)
       
    
    Me.MousePointer = 0
    
    vasSchESR.MaxRows = vasSchESR.DataRowCnt
    vasSchESR.RowHeight(-1) = 12.6
    
    vasActiveCell vasSchESR, 1, 2
    
    frameSch.Visible = True
    'EquipCode,ExamCode, ExamName, Seqno, PointSize, RefLow, RefHigh
    
    Exit Sub
ErrHandle:
    Me.MousePointer = 0
    Exit Sub
End Sub

Private Sub vasSch_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow, i, liEquipCode As Long
    Dim lsID As String
    Dim liRet As Integer
    'Dim lsID As String
    Dim lsResult As String
    Dim mExam
    
    If KeyCode = vbKeyReturn Then
        lRow = vasSch.ActiveRow
        
        SQL = "Select barcode, diskno, posno, examtype from pat_res where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' and barcode = '" & Trim(GetText(vasSch, lRow, 2)) & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = Trim(GetText(vasSch, lRow, 2)) Then
            If MsgBox("입력하신 검체 [" & Trim(GetText(vasSch, lRow, 2)) & "]는 장비 " & Trim(gReadBuf(3)) & "의 " & Trim(gReadBuf(1)) & " Rack " & Trim(gReadBuf(2)) & " Position 에서 검사한 것입니다 " & vbCrLf & _
                      " " & vbCrLf & _
                      "저장하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, "알림") = vbNo Then
                SetText vasSch, Trim(GetText(vasSch, lRow, gMaxCol + 4)), lRow, 2
                Exit Sub
            End If
        End If
        
        If MsgBox("결과를 전송하시겠습니까? ", vbCritical + vbYesNo + vbDefaultButton2, "알림") = vbYes Then
            
            lsID = Trim(GetText(vasSch, lRow, 2))
            
            With vasSch
                'res = Get_PatInfo(lsID)
                res = Online_XML(gXml_S03, lsID)
                
                If res > 0 Then
'                    .SetText 3, lRow, gPatient_Info.PTNO
'                    .SetText 4, lRow, gPatient_Info.PATNAME
'                    .SetText 6, lRow, gPatient_Info.Sex
'                    .SetText 7, lRow, gPatient_Info.Age
'                    .SetText 8, lRow, gPatient_Info.WD_NO
'                    .SetText 9, lRow, gPatient_Info.ACPT_NO
                    
                    .SetText 3, lRow, gPat_Info_Select.PT_NO
                    .SetText 4, lRow, gPat_Info_Select.PT_NM
                    .SetText 6, lRow, gPat_Info_Select.Sex
                    .SetText 7, lRow, gPat_Info_Select.Age
                    .SetText 8, lRow, gPat_Info_Select.ORD_SITE
                    .SetText 9, lRow, gPat_Info_Select.ACPTNO_1
            
                    SQL = "Update pat_res set " & vbCrLf & _
                          "  barcode = '" & lsID & "', " & vbCrLf & _
                          "  pid = '" & Trim(GetText(vasSch, lRow, 3)) & "', " & vbCrLf & _
                          "  pname = '" & Trim(GetText(vasSch, lRow, 4)) & "', " & vbCrLf & _
                          "  pjumin = '" & Trim(GetText(vasSch, lRow, 5)) & "', " & vbCrLf & _
                          "  psex = '" & Trim(GetText(vasSch, lRow, 6)) & "', " & vbCrLf & _
                          "  page1 = '" & Trim(GetText(vasSch, lRow, 7)) & "', " & vbCrLf & _
                          "  WardRoom = '" & Trim(GetText(vasSch, lRow, 8)) & "', " & vbCrLf & _
                          "  receno = '" & Trim(GetText(vasSch, lRow, 9)) & "' " & vbCrLf & _
                          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                          "  and barcode = '" & Trim(GetText(vasSch, lRow, gMaxCol + 4)) & "'"
                    res = SendQuery(gLocal, SQL)
                    
                    SQL = " Update pat_resmemo Set " & vbCrLf & _
                          " barcode = '" & lsID & "' " & vbCrLf & _
                          " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
                          " And equipno = '" & gEquip & "' " & vbCrLf & _
                          " And barcode = '" & Trim(GetText(vasSch, lRow, gMaxCol + 4)) & "' "
                    res = SendQuery(gLocal, SQL)
                    
                    SQL = " Update resflag Set " & vbCrLf & _
                          " barcode = '" & lsID & "' " & vbCrLf & _
                          " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
                          " And barcode = '" & Trim(GetText(vasSch, lRow, gMaxCol + 4)) & "' "
                    res = SendQuery(gLocal, SQL)
                    
'                    SaveData lRow & " : " & lsID
                    res = ToServer(lRow, vasSch, 1)
                    If res = 1 Then
                        vasSch.Row = lRow
                        vasSch.Col = 1
                        vasSch.Value = 0
        
                        SetText vasSch, "완료", lRow, gResCol
                        SetBackColor vasSch, lRow, lRow, 1, 1, 202, 255, 112
                        'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
                    
                        SQL = "Update worklist set OrdFlag = 'D' where barcode = '" & lsID & "' "
                        res = SendQuery(gLocal, SQL)
                    
                    ElseIf res = 2 Then
                        SetText vasSch, "결과", lRow, gResCol
                    Else
                        SetText vasSch, "실패", lRow, gResCol
                        SetForeColor vasSch, lRow, lRow, 1, 1, 255, 0, 0
                        'SetBackColor vasID, gPreRow, gPreRow, colCheckBox, vasID.MaxCols, 255, 255, 255
                        'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
                    End If
            
                Else
                    SetBackColor vasList, lRow, lRow, 1, 1, 255, 0, 0
                    SetText vasList, "미접수", lRow, gResCol
                    
                End If
            End With
                        
        End If
    End If

End Sub

Private Sub vasSchESR_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        Select Case Col
        Case 2
            vasSort vasSchESR, Col
        Case 3
            vasSort vasSchESR, Col
        Case 4
            vasSort vasSchESR, Col, 2
        'Case 5
        '    vasSort vasSch, Col, 6
        Case 9
            vasSort vasSchESR, Col
        Case 11, 12
            vasSort vasSchESR, 11, 12, 9
        Case 1, 6, 7, 8, 10
            vasSort vasSchESR, Col, 9
        End Select
    End If
End Sub

Sub SearchSample(ByVal asBarcode As String)
    Dim lRow, lCol As Long
    Dim lsID, lsType As String
    Dim liEquipCode As Integer
    Dim i As Integer
    
    Dim rs_Res As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    vasSchESR.Visible = False
    vasSch.Visible = True
    cmdCol2.Visible = True
    
    ClearSpread vasSch, 0, 2
    
    Me.MousePointer = 11
    
'    For lCol = 1 To vasList.MaxCols
'        SetText vasSch, Trim(GetText(vasList, 0, lCol)), 0, lCol
'    Next lCol

    vasSch.MaxCols = vasList.MaxCols
    For lCol = 1 To vasList.MaxCols
        SetText vasSch, Trim(GetText(vasList, 0, lCol)), 0, lCol
        vasSch.ColWidth(lCol) = vasList.ColWidth(lCol)
    Next lCol
    
    cmdPrint.Visible = True
    frameSch.Visible = True
    
    lsID = Trim(asBarcode)
    
    SQL = "Select a.barcode, a.pid, a.pname, a.pjumin, a.psex, a.page1, a.WardRoom, a.ReceNo, a.examtype, a.diskno, " & _
            "a.posno, a.sendflag, a.equipcode, a.examcode, a.result, " & _
            "a.refflag, a.resdate, a.seqno, b.WBCSusp, b.RBCSusp, " & _
            "b.PLTSusp, b.SampleJudg, b.PBSFlag  " & vbCrLf & _
          "from pat_res a,res_flag b" & vbCrLf & _
          "where a.examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
          "  and b.barcode = a.barcode " & vbCrLf & _
          "Order by a.ReceNo, a.barcode, a.equipcode "
    
    SQL = "Select distinct a.barcode, a.pid, a.pname, a.pjumin, a.psex, a.page1, a.WardRoom, a.ReceNo, a.examtype, a.diskno, " & _
            "a.posno, a.sendflag, a.equipcode, '', a.result, " & _
            "a.refflag, a.resdate, a.seqno, b.WBCSusp, b.RBCSusp, " & _
            "b.PLTSusp, b.SampleJudg, b.PBSFlag  " & vbCrLf & _
          "from pat_res a,res_flag b" & vbCrLf & _
          "where a.examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
          "  and b.barcode = a.barcode " & vbCrLf & _
          "Order by a.ReceNo, a.barcode, a.examtype, a.equipcode "
    
    Set rs_Res = db_select_rs(gLocal, SQL)
       
    If rs_Res Is Nothing Then GoTo ErrHandle
    
    lsID = ""
    lsType = ""
    lRow = 0
    Do While Not rs_Res.EOF
        'MsgBox Trim(CStr(rs_Res.Fields.Item(0).Value)) & " : " & Trim(CStr(rs_Res.Fields.Item(8).Value))
        
        If Trim(CStr(rs_Res.Fields.Item(0).Value)) <> lsID Or Trim(CStr(rs_Res.Fields.Item(8).Value)) <> lsType Then
            lRow = lRow + 1
            
            If lRow > vasSch.MaxRows Then
                vasSch.MaxRows = lRow
                
                vasSch.RowHeight(lRow) = 12.6
            End If
            
            For lCol = 2 To 13
                If IsNull(rs_Res.Fields.Item(lCol - 2).Value) Then
                    SetText vasSch, "", lRow, lCol
                Else
                    SetText vasSch, Trim(CStr(rs_Res.Fields.Item(lCol - 2).Value)), lRow, lCol
                End If
            Next lCol
            SetText vasSch, Trim(CStr(rs_Res.Fields.Item(0).Value)), lRow, gMaxCol + 4
            
            If IsNumeric(Trim(GetText(vasSch, lRow, 11))) Then  'Rack
                vasSch.SetText 11, lRow, Format(CDbl(GetText(vasSch, lRow, 11)), "0000")
            End If
            If IsNumeric(Trim(GetText(vasSch, lRow, 12))) Then  'Pos
                vasSch.SetText 12, lRow, Format(CDbl(GetText(vasSch, lRow, 12)), "00")
            End If
            
            For i = 0 To 3
                If Not IsNull(rs_Res.Fields.Item(18 + i).Value) Then
                    SetText vasSch, Trim(CStr(rs_Res.Fields.Item(18 + i).Value)), lRow, gMaxCol + i
                End If
            Next i
            If Not IsNull(rs_Res.Fields.Item(22).Value) Then
                SetText vasSch, Trim(CStr(rs_Res.Fields.Item(22).Value)), lRow, gMaxCol + 5
            End If
            
            Select Case Trim(GetText(vasSch, lRow, gResCol))
            Case "B"
                SetText vasSch, "완료", lRow, gResCol
                vasSch.Row = lRow
                vasSch.Col = 1
                vasSch.Value = 0
            Case "E"
                SetText vasSch, "실패", lRow, gResCol
                vasSch.Row = lRow
                vasSch.Col = 1
                vasSch.Value = 1
            Case Else
                SetText vasSch, "수신", lRow, gResCol
                vasSch.Row = lRow
                vasSch.Col = 1
                vasSch.Value = 1
            End Select
        
            Select Case Trim(GetText(vasSch, lRow, gMaxCol + 3))
            Case "0"
                SetText vasSch, "Negative", lRow, gMaxCol + 3
                SetBackColor vasSch, lRow, lRow, 2, 2, 255, 255, 255
            Case "1"
                SetText vasSch, "Positive", lRow, gMaxCol + 3
                SetBackColor vasSch, lRow, lRow, 2, 2, 255, 224, 193
            Case "2"
                SetText vasSch, "Error", lRow, gMaxCol + 3
                SetBackColor vasSch, lRow, lRow, 2, 2, 255, 224, 193
            Case "3"
                SetText vasSch, "Potive+Error", lRow, gMaxCol + 3
                SetBackColor vasSch, lRow, lRow, 2, 2, 255, 224, 193
            Case "4"
                SetText vasSch, "QC Sample", lRow, gMaxCol + 3
                SetBackColor vasSch, lRow, lRow, 2, 2, 255, 255, 255
            End Select
            
            If Trim(GetText(vasSch, lRow, gMaxCol)) = "" Then
                SetBackColor vasSch, lRow, lRow, flagWBC, flagWBC, 255, 255, 255
            ElseIf Trim(GetText(vasSch, lRow, gMaxCol)) <> "0" Then
                SetBackColor vasSch, lRow, lRow, flagWBC, flagWBC, 193, 234, 255
            Else
                SetBackColor vasSch, lRow, lRow, flagWBC, flagWBC, 255, 255, 255
            End If
            If Trim(GetText(vasSch, lRow, gMaxCol + 1)) = "" Then
                SetBackColor vasSch, lRow, lRow, flagRBC, flagRBC, 255, 255, 255
            ElseIf Trim(GetText(vasSch, lRow, gMaxCol + 1)) <> "0" Then
                SetBackColor vasSch, lRow, lRow, flagRBC, flagRBC, 193, 234, 255
            Else
                SetBackColor vasSch, lRow, lRow, flagRBC, flagRBC, 255, 255, 255
            End If
            If Trim(GetText(vasSch, lRow, gMaxCol + 2)) = "" Then
                SetBackColor vasSch, lRow, lRow, flagPLT, flagPLT, 255, 255, 255
            ElseIf Trim(GetText(vasSch, lRow, gMaxCol + 2)) <> "0" Then
                SetBackColor vasSch, lRow, lRow, flagPLT, flagPLT, 193, 234, 255
            Else
                SetBackColor vasSch, lRow, lRow, flagPLT, flagPLT, 255, 255, 255
            End If
            
'            If Trim(GetText(vasSch, lRow, gMaxCol + 5)) = "1" Then
'                SetBackColor vasSch, lRow, lRow, 2, 2, 255, 224, 193
'            End If
            
            If Trim(GetText(vasSch, lRow, gMaxCol + 5)) = 1 Then
                SetBackColor vasSch, lRow, lRow, 9, 9, 255, 224, 193
'                SetForeColor vasList, lRow, lRow, 9, 9, 0, 0, 160
'                vasList.Row = lRow
'                vasList.Col = 9
'                vasList.FontBold = True
            Else
                SetBackColor vasSch, lRow, lRow, 9, 9, 255, 255, 255
'                SetForeColor vasList, lRow, lRow, 9, 9, 0, 0, 0
'                vasList.Row = lRow
'                vasList.Col = 9
'                vasList.FontBold = False
            End If
            
        End If
        
        For liEquipCode = 1 To UBound(gArrExam)
            If CInt(gArrExam(liEquipCode, 1)) = Trim(rs_Res.Fields.Item(12).Value) Then
                'lCol = gResCol + liEquipCode
                lCol = gArrExam(liEquipCode, 10)
                'lCol = liEquipCode - 1
'                If liEquipCode = 1 Then
'                    MsgBox ""
'                End If
                SetText vasSch, Trim(CStr(rs_Res.Fields.Item(14).Value)), lRow, lCol
                Select Case Trim(CStr(rs_Res.Fields.Item(15).Value))
                Case "H"
                    SetForeColor vasSch, lRow, lRow, lCol, lCol, 255, 127, 0
                Case "L"
                    SetForeColor vasSch, lRow, lRow, lCol, lCol, 0, 127, 255
                Case Else
                    SetForeColor vasSch, lRow, lRow, lCol, lCol, 0, 0, 0
                End Select
                
                Exit For
            End If
        Next liEquipCode
        
        lsID = Trim(CStr(rs_Res.Fields.Item(0).Value))
        lsType = Trim(CStr(rs_Res.Fields.Item(8).Value))
        
        rs_Res.MoveNext
    Loop
    
    rs_Res.Close
    
    Me.MousePointer = 0
    
    vasSch.MaxRows = vasSch.DataRowCnt
    vasSch.RowHeight(-1) = 12.6
    
    vasActiveCell vasSch, 1, 2
    
    frameSch.Visible = True
    'EquipCode,ExamCode, ExamName, Seqno, PointSize, RefLow, RefHigh
    
    Exit Sub
ErrHandle:
    Me.MousePointer = 0
    Exit Sub

End Sub

Function Insert_Data(ByVal argSpcRow As Integer, argSpread As vaSpread, Optional asFlag As Integer = 0) As Integer
'서버의 데이타 베이스에 저장
    Dim iRow As Integer
    Dim i, j, k, z As Integer
    
    Dim lCol As Long
    Dim lsID As String
    Dim lsRet As String
    Dim lsSeqNo As String
    Dim lsFinalFlag, lsFinishDate, lsFinishTime As String
    Dim lsCnt, lsCnt1 As String
    Dim lsTestDate As String
    Dim lsTestTime As String
    Dim lsStatus As String
    
    Dim lsRTestcd As String
    Dim lsRemark As String
    
    Dim QCFlag As Integer
    
    Dim wDpcd, wDate, wSlip, wItem, wIoro, wWkno, wIdno, wSmyy, wSmsq, wSmsb, wSms1, wStat As String
    Dim aSeqNo, aSpcl, aGubun, aRtcd, aState, aVal, aRes As String
    
    Dim sDate As String
    Dim sExamCode As String
    Dim sResult As String
    Dim sResultN As String
        
    Dim iRetCnt As Integer
    
    Dim ResultCnt As String
    
    Dim lsPanic, lsDelta As String
    '검사코드의 검사마스터에서 정도관리 기본 정보 가져오기
    Dim t_coifpaup, t_coifpalo, t_coifdeup, t_coifdelo, t_coifdech, t_coifditv, t_coifcove As String
    '과거 결과 가져오기
    Dim hh_cpnartdt, hh_cpnaritm, hh_cpnarslt As String
    Dim hh_term '과거검사일과 현재 검사일의 간격
    Dim gijun_val   'delta check 방법에 따른 기준값
    
    Dim lsQCOn, lsNState, tSmsn, lsQCChk As String
    Dim iDP, iNone As Integer
    
    Insert_Data = -1
    
    'sDate = Format(GetDateFull, "yyyymmddhhnnss")
    sDate = Format(GetDateFull, "yyyymmddhhnn")
    
    QCFlag = -1
    
    If argSpcRow < 1 Or argSpcRow > vasList.DataRowCnt Then
        Insert_Data = 0
        Exit Function
    End If
    
    lsTestDate = Format(Date, "yyyymmdd")
    lsTestTime = Format(Time, "hhnnss")
    
'    iRow = argRow
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread vasResTemp
    
    lsID = Trim(GetText(argSpread, argSpcRow, 2))
    
    If Not IsNumeric(lsID) Then Exit Function
    
    lsSeqNo = Trim(GetText(argSpread, argSpcRow, 9))
    
    If Trim(GetText(vasList, argSpcRow, 10)) = "IPU1" Then
        lsRTestcd = "HST1"
    Else
        lsRTestcd = "HST2"
    End If
    If IsNumeric(GetText(vasList, argSpcRow, 11)) Then
        vasList.SetText 11, argSpcRow, Format(CCur(GetText(vasList, argSpcRow, 11)), "000")
    Else
        vasList.SetText 11, argSpcRow, "000"
    End If
    
    lsRemark = lsRTestcd & Trim(GetText(vasList, argSpcRow, 11)) & "-" & Trim(GetText(vasList, argSpcRow, 12))
    
    SQL = " Select equipcode, examcode, result, refflag, panicflag, deltaflag, examno " & vbCrLf & _
          " From pat_res " & vbCrLf & _
          " Where examdate = '" & Format(Trim(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(argSpread, argSpcRow, 2)) & "' "
    res = db_select_Vas(gLocal, SQL, vasResTemp)
    If res = -1 Then SaveQuery SQL
    
    Debug.Print res & " : " & SQL
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    '장부에서 읽어오기
    SQL = "Select wnifdpcd, wnifdate, wnifslip, wnifitem, wnifoitp, " & vbCrLf & _
          "       wnifwkno, wnifidno, wnifsmyr, wnifsmsn, wnifsms1, " & vbCrLf & _
          "       wnifstat, wnifrpdt, wnifrptm, wnifname, wnifrsex, wnifsms1" & vbCrLf & _
          "From arcwnifh " & vbCrLf & _
          "Where wnifdpcd = 'CP' "
    SQL = SQL & vbCrLf & _
      "  AND wnifitem = '00' " & vbCrLf & _
      "  AND wnifsmyr = '" & Mid(lsID, 1, 2) & "' " & vbCrLf & _
      "  ANd wnifsmsn = " & Mid(lsID, 3, 7) & vbCrLf & _
      "  And wnifstat <> 'X' "
'          "  And wnifsmsb = " & Mid(sSpec, 10, 1) & vbCrLf & _
'          "  And wnifsms1 = " & Mid(sSpec, 11) & vbCrLf
    res = db_select_Col(gServer, SQL)
    If Trim(gReadBuf(0)) = "" Then
        Exit Function
    End If
    
    wDpcd = Trim(gReadBuf(0))   '검사부터
    wDate = Trim(gReadBuf(1))   '검사일자
    wSlip = Trim(gReadBuf(2))   '슬립번호
    wItem = Trim(gReadBuf(3))   '검사항목
    wIoro = Trim(gReadBuf(4))   '입원외래검사실구분
    wWkno = Trim(gReadBuf(5))   '시리얼 No
    wIdno = Trim(gReadBuf(6))   '등록번호
    wSmyy = Trim(gReadBuf(7))   '검체번호(년도)
    wSmsq = Trim(gReadBuf(8))   '검체번호(일련번호)
    wSmsb = Trim(gReadBuf(9))   '검체번호(Sub sq.)    If asFlag = 1 Then
    
    If asFlag = 1 Then
        
        ClearSpread vasComList
        ClearSpread vasExam1
        
        SQL = "SELECT b.cpnwcode, c.coifabbr " & CR
        SQL = SQL & "From arccpnwh b, arcwnifh a, ABCCOIFM c" & CR
        SQL = SQL & "WHERE a.wnifdpcd = 'CP' " & vbCrLf
        SQL = SQL & "  AND a.wnifsmyr = '" & Mid(lsID, 1, 2) & "' " & CR
        SQL = SQL & "  AND a.wnifsmsn = '" & Mid(lsID, 3, 7) & "' " & CR
        SQL = SQL & "  And a.wnifstat <> 'X' " & CR
        SQL = SQL & "  and b.cpnwdpcd = a.wnifdpcd " & CR
        SQL = SQL & "  and b.cpnwdate = a.wnifdate " & CR
        SQL = SQL & "  and b.cpnwslip = a.wnifslip " & CR
        SQL = SQL & "  and b.cpnwitem = a.wnifitem " & CR
        SQL = SQL & "  and b.cpnwoitp = a.wnifoitp " & CR
        SQL = SQL & "  and b.cpnwwkno = a.wnifwkno " & CR
        'SQL = SQL & "  and b.cpnwcode In (" & sExamCode & ") " & CR
        SQL = SQL & "  and b.cpnwstat <> 'X' "
        SQL = SQL & "  and c.coifcode = b.cpnwcode  "
    
        res = db_select_Vas(gServer, SQL, vasComList)
        
        If vasComList.DataRowCnt > 0 Then
            For i = 1 To vasResTemp.DataRowCnt
            
                For j = 1 To UBound(gArrExam)
                    If CInt(gArrExam(j, 1)) = Trim(GetText(vasResTemp, i, 1)) Then
                        z = -1
                        
                        For k = 1 To vasComList.DataRowCnt
                        
                            If Trim(GetText(vasComList, k, 1)) = gArrExam(j, 2) Then
                                z = 1
                                
                                SQL = " update pat_res set  " & vbCrLf & _
                                      " examcode = '" & gArrExam(j, 2) & "', examno = '" & gArrExam(j, 9) & "' " & vbCrLf & _
                                      " Where examdate = '" & Format(Trim(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                                      " And equipno = '" & gEquip & "' " & vbCrLf & _
                                      " And barcode = '" & Trim(GetText(argSpread, argSpcRow, 2)) & "' " & vbCrLf & _
                                      " and equipcode = '" & Trim(GetText(vasResTemp, i, 1)) & "' "
                                res = SendQuery(gLocal, SQL)
                                If res = -1 Then SaveQuery SQL
                                
                                vasResTemp.SetText 2, i, gArrExam(j, 2)
                                vasResTemp.SetText 7, i, gArrExam(j, 9)
                                
                                Debug.Print "장비 : " & Trim(GetText(vasResTemp, i, 1)) & ", 검사코드 : " & gArrExam(j, 2) & ", 검사번호 : " & gArrExam(j, 9)
                                DeleteRow vasComList, k, k
                                
                                Exit For
                            End If
                        Next k
                        If z = 1 Then
                            Exit For
                        End If
                    End If
                Next j
                                    
            Next i
        End If
        
        ClearSpread vasComList
    End If

    wSms1 = Trim(gReadBuf(15))   '검체번호(Sub sq.)
    wStat = Trim(gReadBuf(10))
    
    If wSms1 = 9 Then lsQCOn = "Y"          '/* QC 검사이다        */
    
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    gReadBuf(2) = ""
    gReadBuf(3) = ""
    gReadBuf(4) = ""
        
    
    lsStatus = "D"
        
'    SaveData lsID & " : 저장 시작전 : " & vasResTemp.DataRowCnt
    
    '서버로 결과값 저장하기
    For i = 1 To vasResTemp.DataRowCnt
        sResult = Trim(GetText(vasResTemp, i, 3))
        sExamCode = Trim(GetText(vasResTemp, i, 2))
        sResultN = sResult
        
        Debug.Print lsID & " : 쿼리 전 : " & Trim(GetText(vasResTemp, i, 2)) & " : " & sResultN
        
        lsPanic = "N"
        lsDelta = "N"
        lsNState = "1"
        
        If sResult <> "" And sExamCode <> "" And IsNumeric(sResultN) = True Then
            iRetCnt = iRetCnt + 1
                        
            If lsQCOn = "Y" Then
                sExamCode = "QC" & sExamCode
            End If
                        
            SQL = "Select  NVL(rtrim(coifpaup),' '), NVL(rtrim(coifpalo),' '), " & _
                    "NVL(rtrim(coifdeup),' '), NVL(rtrim(coifdelo),' '), " & _
                    "NVL(rtrim(coifdech),' '), NVL(coifditv, 1 ), " & _
                    "NVL(rtrim(coifcove),' ')" & vbCrLf & _
                  "From abccoifm " & vbCrLf & _
                  "where coifcode = '" & sExamCode & "' "
            res = db_select_Col(gServer, SQL)
            
            If lsQCOn = "Y" Then
                sExamCode = Mid(sExamCode, 3)
            End If
                        
            t_coifpaup = Trim(gReadBuf(0))
            t_coifpalo = Trim(gReadBuf(1))
            t_coifdeup = Trim(gReadBuf(2))
            t_coifdelo = Trim(gReadBuf(3))
            t_coifdech = Trim(gReadBuf(4))
            t_coifditv = Trim(gReadBuf(5))
            t_coifcove = Trim(gReadBuf(6))
            
            'Panic Check=================================================================
            If t_coifdech = "X" Then
                lsPanic = "N"
            Else
                If IsNumeric(sResult) = True Then
                    If IsNumeric(t_coifpaup) Then
                        If CCur(sResult) > CCur(t_coifpaup) Then
                            lsPanic = "Y"
                            iDP = iDP + 1
                            lsNState = "0"
                        End If
                    End If
                    If IsNumeric(t_coifpalo) Then
                        If CCur(sResult) < CCur(t_coifpalo) Then
                            lsPanic = "Y"
                            iDP = iDP + 1
                            lsNState = "0"
                        End If
                    End If
                End If
            End If
                        
            If t_coifdech = "X" Then
                lsDelta = "N"
            Else
                'Delta Check=================================================================
                '과거 데이타 가져오기
                If wSmyy < 98 Then
                    tSmsn = "20"
                Else
                    tSmsn = "19"
                End If
                tSmsn = tSmsn & wSmyy & wSmsq & wSms1

                If InStr(1, wIdno, "GNUHQC") > 0 Then
                    SQL = "SELECT NVL(SUBSTR(cpnwupdt,1,8),'-'), NVL(cpnwrslt,' ') " & vbCrLf & _
                          "From arcqcwnh, arccpnwh " & vbCrLf & _
                          "WHERE  qcwnsmsn < '" & Mid(tSmsn, 3) & "' " & vbCrLf & _
                          "and    qcwnidno = '" & wIdno & "' " & vbCrLf & _
                          "and    qcwndate = cpnwdate " & vbCrLf & _
                          "and    qcwnaslp = cpnwslip " & vbCrLf & _
                          "and    qcwndpcd = cpnwdpcd " & vbCrLf & _
                          "and    qcwnwkno = cpnwwkno " & vbCrLf & _
                          "and    cpnwcode = '" & sExamCode & "' "
                Else
                    '  ':h_cpnwupdt,:h_cpnwrslt
                    SQL = "SELECT cpnwidno, NVL(SUBSTR(cpnwupdt,1,8),'-'), NVL(cpnwrslt,' ') , DECODE(wnifsmyr,'98','19','99','19','20')||wnifsmyr||wnifsmsn||wnifsms1 " & vbCrLf & _
                          "  From arccpnwh, arcwnifh " & vbCrLf & _
                          " WHERE cpnwidno  = '" & wIdno & "' " & vbCrLf & _
                          "   AND cpnwdpcd  = '" & wDpcd & "' " & vbCrLf & _
                          "   AND cpnwcode  = '" & sExamCode & "' " & vbCrLf & _
                          "   AND cpnwslip = wnifslip " & vbCrLf & _
                          "   and cpnwdpcd  = wnifdpcd " & vbCrLf & _
                          "   and cpnwdate  = wnifdate " & vbCrLf & _
                          "   and cpnwwkno  = wnifwkno " & vbCrLf & _
                          "   and cpnwoitp  = wnifoitp " & vbCrLf & _
                          "   and cpnwitem  = wnifitem " & vbCrLf & _
                          "   and cpnwidno  = wnifidno " & vbCrLf & _
                          "   and wnifstat  = '2' " & vbCrLf & _
                          "   and DECODE(wnifsmyr,'98','19','99','19','20')||wnifsmyr||wnifsmsn||wnifsms1 < '" & tSmsn & "' " & vbCrLf & _
                          "   and cpnwrslt IS NOT NULL " & vbCrLf & _
                          "   and cpnwupdt IS NOT NULL " & vbCrLf & _
                          "ORDER BY cpnwupdt DESC"
                End If
                SQL = "select  cpnartdt, cpnaritm, NVL(RTRIM(cpnarslt), ' ')" & vbCrLf & _
                      "from    arccpnah" & vbCrLf & _
                      "where   cpnaidno = '" & wIdno & "' " & vbCrLf & _
                      "and     cpnadpcd = '" & wDpcd & "' " & vbCrLf & _
                      "and     cpnacode = '" & sExamCode & "' " & vbCrLf & _
                      "and    (cpnartdt || cpnaritm) < '" & Left(sDate, 12) & "' " & vbCrLf & _
                      "and   ( cpnadate != '" & wDate & "' or " & vbCrLf & _
                      "        cpnaslip != '" & wSlip & "' or " & vbCrLf & _
                      "        cpnaitem != '" & wItem & "' or " & vbCrLf & _
                      "        cpnaioro != '" & wIoro & "') " & vbCrLf & _
                      "and     nvl(rtrim(cpnarslt), '!') != '!' " & vbCrLf & _
                      "order by cpnartdt desc , cpnaritm desc "
                res = db_select_Col(gServer, SQL)
                hh_cpnartdt = Trim(gReadBuf(0))
                hh_cpnaritm = Trim(gReadBuf(1))
                hh_cpnarslt = Trim(gReadBuf(2))
                If IsNumeric(hh_cpnarslt) = True And IsNumeric(sResult) = True And Len(hh_cpnartdt) >= 8 Then
                    hh_term = Abs(DateDiff("d", Left(hh_cpnartdt, 4) & "/" & Mid(hh_cpnartdt, 5, 2) & "/" & Mid(hh_cpnartdt, 7, 2), Left(sDate, 4) & "/" & Mid(sDate, 5, 2) & "/" & Mid(sDate, 7, 2)))
                    
                    If IsNumeric(t_coifditv) = True And IsNumeric(hh_term) = True Then
                        If CCur(hh_term) = 0 Then hh_term = 1
                        If CCur(hh_cpnarslt) = 0 Then hh_cpnarslt = 1
                        
                        If CCur(t_coifditv) >= CCur(hh_term) Then   '유효기간 내인 경우
                            Select Case t_coifdech
                            Case "1": gijun_val = CCur(sResult) - CCur(hh_cpnarslt)
                            Case "2": gijun_val = (CCur(sResult) - CCur(hh_cpnarslt)) / CCur(hh_term)
                            Case "3": gijun_val = ((CCur(sResult) - CCur(hh_cpnarslt)) / CCur(hh_cpnarslt)) * 100
                            Case "4": gijun_val = (((CCur(sResult) - CCur(hh_cpnarslt)) / CCur(hh_cpnarslt)) / CCur(hh_term)) * 100
                            Case "5": gijun_val = 0
                            Case "P": gijun_val = 0
                            Case "X": gijun_val = 0
                            End Select
                            
                            If IsNumeric(t_coifdeup) Then
                                If CCur(gijun_val) > CCur(t_coifdeup) Then
                                    lsDelta = "Y"
                                    iDP = iDP + 1
                                    lsNState = "0"
                                End If
                            End If
                            If IsNumeric(t_coifdelo) Then
                                If CCur(gijun_val) < CCur(t_coifdelo) Then
                                    lsDelta = "Y"
                                    iDP = iDP + 1
                                    lsNState = "0"
                                End If
                            End If
                        End If
                    End If
                End If
                '=====================================================================================
            End If
            
            
            SQL = "SELECT cpnwdpcd, cpnwdate, cpnwslip, cpnwitem, cpnwoitp, " & _
                  " cpnwwkno, cpnwseqn, cpnwcode, cpnwidno, cpnwrtdt, " & _
                  " cpnwritm, cpnwstat, cpnwspcl, cpnwgubn, cpnwrtcd, " & _
                  " cpnwnval, cpnwrslt" & vbCrLf & _
                  "From arccpnwh " & vbCrLf & _
                  "WHERE cpnwdpcd = '" & wDpcd & "' " & vbCrLf & _
                  "  and cpnwdate = '" & wDate & "' " & vbCrLf & _
                  "  and cpnwslip = '" & wSlip & "' " & vbCrLf & _
                  "  and cpnwitem = '" & wItem & "' " & vbCrLf & _
                  "  and cpnwoitp = '" & wIoro & "' " & vbCrLf & _
                  "  and cpnwwkno = '" & wWkno & "' " & vbCrLf & _
                  "  and cpnwcode = '" & sExamCode & "' " & vbCrLf & _
                  "  and cpnwstat <> 'X' "
            res = db_select_Col(gServer, SQL)
            If res > 0 Then
                aSeqNo = Trim(gReadBuf(6))
                aState = Trim(gReadBuf(11))
                aSpcl = Trim(gReadBuf(12))
                aGubun = Trim(gReadBuf(13))
                aRtcd = Trim(gReadBuf(14))
                aVal = Trim(gReadBuf(15))
                aRes = Trim(gReadBuf(16))
            End If
            
            Select Case aState
            Case "1"    '결과입력
            Case "2"    '결과수정
                aState = "2"
            Case "3"    '재검입력
                aState = "3"
            Case "4"    '재검수정
                aState = "4"
            Case "5", "6"
                aState = "3"
            Case "R"    '재검요구
                aState = "3"
            Case "0"    '접수완료
                aState = "1"
            Case Else
                aState = "1"
            End Select
            
            '확정에 저장
            SQL = "Select count(*) From arccpnah " & vbCrLf & _
                  "WHERE  cpnaidno = '" & wIdno & "'  " & vbCrLf & _
                  "  and cpnadpcd = '" & wDpcd & "' " & vbCrLf & _
                  "  and cpnadate = '" & wDate & "' " & vbCrLf & _
                  "  and cpnaslip = '" & wSlip & "' " & vbCrLf & _
                  "  and cpnaitem = '" & wItem & "' " & vbCrLf & _
                  "  and cpnaioro = '" & wIoro & "'  " & vbCrLf & _
                  "  and cpnawkno = '" & wWkno & "'  " & vbCrLf & _
                  "  and cpnacode = '" & sExamCode & "' "
            res = db_select_Col(gServer, SQL)
            If res <= 0 Then
                Exit Function
            End If
            
            If IsNumeric(Trim(gReadBuf(0))) = False Then
                Exit Function
            End If
            
            '워킹에 저장
'            SQL = "Update arccpnwh" & vbCrLf & _
                  " set cpnwdltf ='" & lsDelta & "', " & vbCrLf & _
                  "    cpnwpncf='" & lsPanic & "', " & vbCrLf & _
                  "    cpnwrtdt='" & Left(sDate, 8) & "', " & vbCrLf & _
                  "    cpnwritm='" & Mid(sDate, 9, 4) & "', " & vbCrLf & _
                  "    cpnwstat='" & aState & "', " & vbCrLf & _
                  "    cpnwnval='" & sResultN & "', " & vbCrLf & _
                  "    cpnwrslt='" & sResultN & "'" & vbCrLf & _
                  "WHERE cpnwidno = '" & wIdno & "'  " & vbCrLf & _
                  "  And cpnwdpcd='" & wDpcd & "'" & vbCrLf & _
                  "  And cpnwdate='" & wDate & "'" & vbCrLf & _
                  "  And cpnwslip ='" & wSlip & "'" & vbCrLf & _
                  "  And cpnwitem='" & wItem & "'" & vbCrLf & _
                  "  And cpnwoitp='" & wIoro & "'" & vbCrLf & _
                  "  And cpnwwkno='" & wWkno & "'" & vbCrLf & _
                  "  And cpnwcode = '" & sExamCode & "' "
            SQL = "Update arccpnwh" & vbCrLf & _
                  " set cpnwdltf ='" & lsDelta & "', " & vbCrLf & _
                  "    cpnwpncf='" & lsPanic & "', " & vbCrLf & _
                  "    cpnwstat='" & lsNState & "', " & vbCrLf & _
                  "    cpnwnval='" & sResultN & "', " & vbCrLf & _
                  "    cpnwrslt='" & sResultN & "'," & vbCrLf & _
                  "    cpnwupdt = '" & sDate & "', " & vbCrLf & _
                  "    cpnwrpdt = to_char(sysdate, 'yyyymmddhh24mi'), " & vbCrLf & _
                  "    arccpnwh_d = '" & Left(sDate, 8) & "', " & vbCrLf & _
                  "    arccpnwh_t = '" & Mid(sDate, 9) & "' " & vbCrLf & _
                  "WHERE cpnwidno = '" & wIdno & "'  " & vbCrLf & _
                  "  And cpnwdpcd='" & wDpcd & "'" & vbCrLf & _
                  "  And cpnwdate='" & wDate & "'" & vbCrLf & _
                  "  And cpnwslip ='" & wSlip & "'" & vbCrLf & _
                  "  And cpnwitem='" & wItem & "'" & vbCrLf & _
                  "  And cpnwoitp='" & wIoro & "'" & vbCrLf & _
                  "  And cpnwwkno='" & wWkno & "'" & vbCrLf & _
                  "  And cpnwcode = '" & sExamCode & "' "
            res = SendQuery(gServer, SQL)
            'res = 1
            If res = -1 Then
                db_RollBack gServer
                SaveQuery SQL
                Exit Function
            End If
            
            'Save_PatRes wDate, asRow, "", Trim(sExamCode), "", sResultN, 1, 1
        End If

    Next i
    
    ResultCnt = ""
    SQL = "select count(*) from arccpnwh" & vbCrLf & _
          "WHERE cpnwidno  = '" & wIdno & "'  " & vbCrLf & _
          "  And cpnwdpcd  = '" & wDpcd & "' " & vbCrLf & _
          "  And cpnwdate  = '" & wDate & "' " & vbCrLf & _
          "  And cpnwslip  = '" & wSlip & "' " & vbCrLf & _
          "  And cpnwitem  = '" & wItem & "' " & vbCrLf & _
          "  And cpnwoitp  = '" & wIoro & "' " & vbCrLf & _
          "  And cpnwwkno  = '" & wWkno & "' " & vbCrLf & _
          "  And cpnwstat  = '0' "
    res = db_select_Var(gServer, SQL, ResultCnt)
    If res <= 0 Then
        SaveQuery SQL
    End If
    If Not IsNumeric(Trim(ResultCnt)) Then
        ResultCnt = "0"
    End If
    
    SQL = ""
    Select Case wStat
    Case "5", "6"
        wStat = "3"
    Case "R"    '재검요구
        wStat = "3"
    Case "0"    '접수완료
        wStat = "1"
    Case "3"
        wStat = "3"
    Case Else
        wStat = "1"
    End Select
    If CInt(ResultCnt) = 0 Then '검사결과 입력 완료
        Select Case wStat
        Case "1"
            wStat = "2"
        Case "3"
            wStat = "4"
        End Select
    End If
    
    If iDP = 0 Then
        lsQCChk = "N"
    Else
        lsQCChk = "Y"
    End If
    
    If iNone = 0 And iDP = 0 Then
        wStat = "2"
    Else
        wStat = "1"
    End If
    
    
    '장부 상태 Update
    SQL = "Update arcwnifh " & vbCrLf & _
          "Set wnifrpdt = '" & Left(sDate, 8) & "', " & vbCrLf & _
          "    wnifrptm = '" & Mid(sDate, 9, 4) & "', " & vbCrLf & _
          "    wnifstat = '" & wStat & "', " & vbCrLf & _
          "    wnifqchk = '" & lsQCChk & "' " & vbCrLf & _
          "Where wnifdpcd = '" & wDpcd & "' " & vbCrLf & _
          "  AND wnifitem = '" & wItem & "' " & vbCrLf & _
          "  AND wnifsmyr = '" & wSmyy & "' " & vbCrLf & _
          "  ANd wnifsmsn = '" & wSmsq & "' " & vbCrLf & _
          "  And wnifsmsb = '" & wSmsb & "' " & vbCrLf & _
          "  And wnifsms1 = " & wSms1 & vbCrLf & _
          "  And wnifstat <> 'X' "
    res = SendQuery(gServer, SQL)
    'res = 1
    If res = -1 Then
        db_RollBack gServer
        SaveQuery SQL
        Exit Function
    End If
    
    
    db_Commit gServer
    
    Insert_Data = 1
    
End Function

Function ToServer(ByVal argSpcRow As Integer, argSpread As vaSpread, Optional asFlag As Integer = 0) As Integer
    Dim i           As Integer
    Dim j           As Integer
    
    Dim lsID        As String
    
    Dim lRow        As Long
    Dim lsMsg       As String
    Dim lsEqFlag    As String
    
    Dim lsExamCode  As String
    
    Dim sRet        As String
    
    Dim sParam As String
    
    Dim sResRow As Long
    
    Dim sEquip As String
    
    ToServer = -1
    
    lRow = argSpcRow
    
    If lRow < 1 Or lRow > argSpread.DataRowCnt Then Exit Function
    
    lsID = Trim(GetText(argSpread, lRow, 2))
    
    If lsID = "" Then Exit Function
    
    If IsNumeric(lsID) = False Or Len(lsID) < 11 Then Exit Function
    
'    lsExamCode = ""
'    res = Online_XML(gXml_S07, lsID)
'    For i = 0 To UBound(gExam_Select)
'        If lsExamCode = "" Then
'            lsExamCode = "'" & gExam_Select(i).TST_CD & "'"
'        Else
'            lsExamCode = lsExamCode & ",'" & gExam_Select(i).TST_CD & "'"
'        End If
'    Next i

    ClearSpread vasTemp

'    SQL = "Select equipcode, examcode, examname, result, result, pid, pname, receno, examtype " & vbCrLf & _
'          "from pat_res " & vbCrLf & _
'          "where  examdate = '" & Format(Trim(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
'          "  and equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and barcode = '" & lsID & "' " & vbCrLf & _
'          "  and examcode in (" & lsExamCode & ") " & vbCrLf & _
'          "  and result <> '' "

    SQL = "Select equipcode, examcode, examname, result, result, pid, pname, receno, examtype " & vbCrLf & _
          "from pat_res " & vbCrLf & _
          "where  examdate = '" & Format(Trim(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & gEquip & "' " & vbCrLf & _
          "  and barcode = '" & lsID & "' " & vbCrLf & _
          "  and result <> '' "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    lsMsg = ""
    lsEqFlag = ""
    
    sEquip = ""
    For i = 1 To vasTemp.DataRowCnt
        sEquip = Trim(GetText(vasTemp, i, 9))
        
        If sEquip <> "" Then
            Exit For
        End If
    Next i
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    SQL = " Select message From pat_resmemo " & vbCrLf & _
          " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(lsID) & "' "
    res = db_select_Col(gLocal, SQL)
    If res > 0 Then
        'lsMsg = "XE2100A : " & Trim(gReadBuf(0))
        lsMsg = sEquip & " : " & Trim(gReadBuf(0))
    Else
        lsMsg = sEquip
    End If
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
    'vasTemp.Visible = True
    If vasTemp.DataRowCnt < 1 Then Exit Function
      
'    Save_Raw_Data lsID & " : 서버 결과 전송 시작"

    On Error GoTo ErrHandle
    
    sParam = ""
    
    For sResRow = 1 To vasTemp.DataRowCnt
        If Trim(GetText(vasTemp, sResRow, 2)) <> "" And IsNumeric(Trim(GetText(vasTemp, sResRow, 5))) = True Then
            Debug.Print Trim(GetText(vasTemp, sResRow, 2)) & " : " & Trim(GetText(vasTemp, sResRow, 5))
            
            sParam = sParam & "<Table>" & _
                    "<QID><![CDATA[PG_SRL.SLP91_P03]]></QID>" & _
                    "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                    "<USERID><![CDATA[" & gServerID & "]]></USERID>" & _
                    "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                    "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                    "<P0><![CDATA[" & lsID & "]]></P0>" & _
                    "<P1><![CDATA[" & Trim(GetText(vasTemp, sResRow, 2)) & "]]></P1>" & _
                    "<P2><![CDATA[" & Trim(GetText(vasTemp, sResRow, 5)) & "]]></P2>" & _
                    "<P3><![CDATA[]]></P3>" & _
                    "<P4><![CDATA[" & sEquip & "]]></P4>" & _
                    "<P5><![CDATA[" & gIFUser & "]]></P5>" & _
                    "<P6><![CDATA[]]></P6>" & _
                    "<P7><![CDATA[" & lsMsg & "]]></P7>" & _
                    "<P8><![CDATA[]]></P8>" & _
                    "<P9><![CDATA[]]></P9>" & _
                    "</Table>"
'            SQL = "Update pat_res set sendflag = 'C' " & vbCrLf & _
'                  "where equipno = '" & gEquip & "' " & vbCrLf & _
'                  "  and barcode = '" & lsID & "' and examcode = '" & Trim(GetText(vasTemp, sResRow, 2)) & "'"
'            res = SendQuery(gLocal, SQL)
'            If res = -1 Then
'                SaveQuery SQL
'                Exit Function
'            End If
        End If
    Next
    
    If sParam <> "" Then
        sParam = "<NewDataSet>" & sParam & "</NewDataSet>"
    
        Online_Result_Qry sParam
    
        ToServer = 1

'        Save_Raw_Data lsID & " : 서버 결과 전송 완료!"
    End If
    
    Exit Function

ErrHandle:
'    Save_Raw_Data Err.Number & " : " & Err.Description & vbCrLf & _
                  SQL
    Resume Next
End Function

Function ToServer1(ByVal argSpcRow As Integer, argSpread As vaSpread, Optional asFlag As Integer = 0) As Integer
    Dim i           As Integer
    Dim j           As Integer
    
    Dim lsID        As String
    
    Dim lRow        As Long
    Dim lsMsg       As String
    Dim lsEqFlag    As String
    
    Dim sRet        As String
    
    Dim sParam As String
    
    Dim sResRow As Long
    
    Dim sEquip As String
    
    ToServer1 = -1
    
    lRow = argSpcRow
    
    If lRow < 1 Or lRow > argSpread.DataRowCnt Then Exit Function
    
    lsID = Trim(GetText(argSpread, lRow, 2))
    
    If lsID = "" Then Exit Function
    
    If IsNumeric(lsID) = False Or Len(lsID) < 11 Then Exit Function
    
    ClearSpread vasTemp

    SQL = "Select equipcode, examcode, examname, result, result, pid, pname, receno, , examtype " & vbCrLf & _
          "from pat_res " & vbCrLf & _
          "where  examdate = '" & Format(Trim(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & gEquip & "' " & vbCrLf & _
          "  and barcode = '" & lsID & "' " & vbCrLf & _
          "  and result <> '' "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    sEquip = ""
    For i = 1 To vasTemp.DataRowCnt
        sEquip = Trim(GetText(vasTemp, i, 9))
        
        If sEquip <> "" Then
            Exit For
        End If
    Next i
    
    lsMsg = ""
    lsEqFlag = ""
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    SQL = " Select message From pat_resmemo " & vbCrLf & _
          " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(lsID) & "' "
    res = db_select_Col(gLocal, SQL)
    If res > 0 Then
        'lsMsg = "XE2100B : " & Trim(gReadBuf(0))
        lsMsg = sEquip & " : " & Trim(gReadBuf(0))
    Else
        lsMsg = sEquip
    End If
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
    'vasTemp.Visible = True
    If vasTemp.DataRowCnt < 1 Then Exit Function
      
'    Save_Raw_Data lsID & " : 서버 결과 전송 시작"
'    Save_Raw_Data lsID & " : 장부 정보 가져오기"

    On Error GoTo ErrHandle
    
    sParam = ""
    
    For sResRow = 1 To vasTemp.DataRowCnt
        If Trim(GetText(vasTemp, sResRow, 2)) <> "" And IsNumeric(Trim(GetText(vasTemp, sResRow, 5))) = True Then
            sParam = sParam & "<Table>" & _
                    "<QID><![CDATA[PG_SRL.SLP91_P03]]></QID>" & _
                    "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                    "<USERID><![CDATA[" & gServerID & "]]></USERID>" & _
                    "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                    "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                    "<P0><![CDATA[" & lsID & "]]></P0>" & _
                    "<P1><![CDATA[" & Trim(GetText(vasTemp, sResRow, 2)) & "]]></P1>" & _
                    "<P2><![CDATA[" & Trim(GetText(vasTemp, sResRow, 5)) & "]]></P2>" & _
                    "<P3><![CDATA[]]></P3>" & _
                    "<P4><![CDATA[" & sEquip & "]]></P4>" & _
                    "<P5><![CDATA[" & gIFUser & "]]></P5>" & _
                    "<P6><![CDATA[]]></P6>" & _
                    "<P7><![CDATA[" & lsMsg & "]]></P7>" & _
                    "<P8><![CDATA[]]></P8>" & _
                    "<P9><![CDATA[]]></P9>" & _
                    "</Table>"
'            SQL = "Update pat_res set sendflag = 'C' " & vbCrLf & _
'                  "where equipno = '" & gEquip & "' " & vbCrLf & _
'                  "  and barcode = '" & lsID & "' and examcode = '" & Trim(GetText(vasTemp, sResRow, 2)) & "'"
'            res = SendQuery(gLocal, SQL)
'            If res = -1 Then
'                SaveQuery SQL
'                Exit Function
'            End If
        End If
    Next
    
    sParam = "<NewDataSet>" & sParam & "</NewDataSet>"
    
    Online_Result_Qry sParam
    
    ToServer1 = 1

'    Save_Raw_Data lsID & " : 서버 결과 전송 완료!"

    Exit Function

ErrHandle:
'    Save_Raw_Data Err.Number & " : " & Err.Description & vbCrLf & _
                  SQL
    Resume Next
End Function


Function ToServer_이전(ByVal argSpcRow As Integer, argSpread As vaSpread, Optional asFlag As Integer = 0) As Integer
'    Dim lsSpcNo, lsEquip, lsExamCode, lsResult, lsEqFlag As String
'
'    Dim iNone As Integer
'
'    Dim i, j As Integer
'
'    Dim lsID As String
'
'    Dim lRow As Long
'    Dim lsQCOn As String
'    Dim lsMsg As String
'
'    Dim sRet As String
'
'    Dim lsLotNo As String
'    Dim lsResDate As String
'
'    ToServer_이전 = -1
'
'    lsQCOn = ""
'
'    lRow = argSpcRow
'
'    If lRow < 1 Or lRow > argSpread.DataRowCnt Then Exit Function
'
'    lsID = Trim(GetText(argSpread, lRow, 2))
'    lsResDate = Format(Date, "yyyymmdd") & Format(Time, "yynnss")
'
'    If lsID = "" Then Exit Function
'
'    If IsNumeric(lsID) = False Or Len(lsID) < 10 Then Exit Function
'
'    lsLotNo = ""
'
'
'    ClearSpread vasTemp
'    ClearSpread vasTemp1
'
'    iNone = 0
'
'    SQL = "Select equipcode, examcode, examname, result, result, pid, pname, receno " & vbCrLf & _
'          "from pat_res " & vbCrLf & _
'          "where  examdate = '" & Format(Trim(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
'          "  and equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and barcode = '" & lsID & "' " & vbCrLf & _
'          "  and result <> '' "
'    res = db_select_Vas(gLocal, SQL, vasTemp)
'
'    lsMsg = ""
'    lsEqFlag = ""
'''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    SQL = " Select message From pat_resmemo " & vbCrLf & _
'          " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
'          " And equipno = '" & gEquip & "' " & vbCrLf & _
'          " And barcode = '" & Trim(lsID) & "' "
'    res = db_select_Col(gLocal, SQL)
'    If res > 0 Then
'        lsMsg = Trim(gReadBuf(0))
'    End If
'''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'    'vasTemp.Visible = True
'    If vasTemp.DataRowCnt < 1 Then Exit Function
'
'    lsSpcNo = ""
'    lsEquip = ""
'    lsExamCode = ""
'    lsResult = ""
'
'    lsLotNo = Trim(GetText(vasTemp, 1, 7))
'
'    iNone = 0
'
'    For i = 1 To vasTemp.DataRowCnt
'        If Trim(GetText(vasTemp, i, 2)) <> "" And Trim(GetText(vasTemp, i, 4)) <> "" Then
'            iNone = iNone + 1
'
'            lsSpcNo = lsSpcNo & lsID & vbTab
'
'            lsEquip = lsEquip & IPU1.Equip_CD & vbTab
'
'            lsExamCode = lsExamCode & Trim(GetText(vasTemp, i, 2)) & vbTab
'
'            lsResult = lsResult & Trim(GetText(vasTemp, i, 4)) & vbTab
'
''            If Trim(lsMsg) <> "" Then
''                lsEqFlag = lsEqFlag & lsMsg & vbTab
''            End If
'
'        End If
'    Next i
'
'    If iNone > 0 Then
'
'        lsSpcNo = vbTab & lsSpcNo
'
'        lsEquip = vbTab & lsEquip
'
'        lsExamCode = vbTab & lsExamCode
'
'        lsResult = vbTab & lsResult
'
'        If Trim(lsMsg) <> "" Then
'            iNone = iNone + 1
'
'            lsSpcNo = lsSpcNo & lsID & vbTab
'
'            lsEquip = lsEquip & IPU1.Equip_CD & vbTab
'
'            lsExamCode = vbTab & "REMARK" & lsExamCode
'
'            lsResult = vbTab & "※장비비고:" & lsMsg & lsResult
'
'
'            lsEqFlag = ""
''
''            j = 0
''            i = InStr(1, lsMsg, vbCrLf)
''            Do While i > 0
''                j = j + 1
''
''                lsEqFlag = lsEqFlag & Left(lsMsg, i - 1) & vbTab
''
''                lsMsg = Mid(lsMsg, i + 2)
''
''                i = InStr(1, lsMsg, vbCrLf)
''            Loop
''
''            If Trim(lsMsg) <> "" Then
''                j = j + 1
''                lsEqFlag = lsEqFlag & lsMsg & vbTab
''            End If
''
''            For i = 1 To iNone - j
''                lsEqFlag = lsEqFlag & "" & vbTab
''            Next i
''
''            lsEqFlag = vbTab & lsEqFlag
'        End If
'
'
'        If Left(lsSpcNo, 1) <> "9" Then
'            If gWorker_Info.WK_ID = "" Then
'                sRet = Online_Result(lsSpcNo, lsExamCode, lsResult, lsEquip, CStr(iNone), lsEqFlag)
'            Else
'                sRet = Online_Result_New(lsSpcNo, lsExamCode, lsResult, lsEquip, CStr(iNone), lsEqFlag, gWorker_Info.WK_ID)
'            End If
'        Else
'            sRet = QCOnline_Result(lsSpcNo, lsLotNo, lsExamCode, lsResult, IPU1.Equip_CD, CStr(iNone), lsResDate)
'        End If
'
'        SaveData lsID & vbTab & gOnline_Ret
'
'        If sRet = "N" Then
'            SQL = "Update pat_res set sendflag = 'B' " & vbCrLf & _
'                  "where equipno = '" & gEquip & "' " & vbCrLf & _
'                  "  and barcode = '" & lsID & "' "
'            res = SendQuery(gLocal, SQL)
'            If res = -1 Then
'                SaveQuery SQL
'                Exit Function
'            End If
'
'            ToServer_이전 = 1
'        Else
'            ToServer_이전 = -1
'        End If
'    Else
'        ToServer_이전 = 0
'    End If
    
End Function

Function ToServer1_원본(ByVal argSpcRow As Integer, argSpread As vaSpread, Optional asFlag As Integer = 0) As Integer
'    Dim lsSpcNo, lsEquip, lsExamCode, lsResult, lsEqFlag As String
'
'    Dim iNone As Integer
'
'
'    Dim lsID As String
'
'    Dim lRow As Long
'    Dim lsQCOn As String
'    Dim lsMsg As String
'
'    Dim sRet As String
'
'    Dim i As Integer
'
'    Dim lsLotNo, lsResDate As String
'
'    ToServer1_원본 = -1
'
'    lsQCOn = ""
'
'    lRow = argSpcRow
'
'    If lRow < 1 Or lRow > argSpread.DataRowCnt Then Exit Function
'
'    lsID = Trim(GetText(argSpread, lRow, 2))
'    lsResDate = Format(Date, "yyyymmdd") & Format(Time, "yynnss")
'
'    If lsID = "" Then Exit Function
'
'    If IsNumeric(lsID) = False Or Len(lsID) < 10 Then Exit Function
'
'    ClearSpread vasTemp
'    ClearSpread vasTemp1
'
'    iNone = 0
'
'    SQL = "Select equipcode, examcode, examname, result, result, pid, pname, receno " & vbCrLf & _
'          "from pat_res " & vbCrLf & _
'          "where  examdate = '" & Format(Trim(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
'          "  and equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and barcode = '" & lsID & "' " & vbCrLf & _
'          "  and result <> '' "
'    res = db_select_Vas(gLocal, SQL, vasTemp)
'    'vasTemp.Visible = True
'    If vasTemp.DataRowCnt < 1 Then Exit Function
'
'''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    SQL = " Select message From pat_resmemo " & vbCrLf & _
'          " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
'          " And equipno = '" & gEquip & "' " & vbCrLf & _
'          " And barcode = '" & Trim(lsID) & "' "
'    res = db_select_Col(gLocal, SQL)
'    If res > 0 Then
'        lsMsg = Trim(gReadBuf(0))
'    End If
'''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'    'vasTemp.Visible = True
'    If vasTemp.DataRowCnt < 1 Then Exit Function
'
'    lsLotNo = Trim(GetText(vasTemp, 1, 7))
'    lsSpcNo = ""
'    lsEquip = ""
'    lsExamCode = ""
'    lsResult = ""
'
'    iNone = 0
'
'    For i = 1 To vasTemp.DataRowCnt
'        If Trim(GetText(vasTemp, i, 2)) <> "" And Trim(GetText(vasTemp, i, 4)) <> "" Then
'            iNone = iNone + 1
'
'            lsSpcNo = lsSpcNo & lsID & vbTab
'
'            lsEquip = lsEquip & IPU2.Equip_CD & vbTab
'
'            lsExamCode = lsExamCode & Trim(GetText(vasTemp, i, 2)) & vbTab
'
'            lsResult = lsResult & Trim(GetText(vasTemp, i, 4)) & vbTab
'        End If
'    Next i
'
'    If iNone > 0 Then
'        lsSpcNo = vbTab & lsSpcNo
'
'        lsEquip = vbTab & lsEquip
'
'        lsExamCode = vbTab & lsExamCode
'
'        lsResult = vbTab & lsResult
'
'
'        If Trim(lsMsg) <> "" Then
'            iNone = iNone + 1
'
'            lsSpcNo = lsSpcNo & lsID & vbTab
'
'            lsEquip = lsEquip & IPU1.Equip_CD & vbTab
'
'            lsExamCode = vbTab & "REMARK" & lsExamCode
'
'            lsResult = vbTab & "※장비비고:" & lsMsg & lsResult
'
'
'            lsEqFlag = ""
'
''
''            j = 0
''            i = InStr(1, lsMsg, vbCrLf)
''            Do While i > 0
''                j = j + 1
''
''                lsEqFlag = lsEqFlag & Left(lsMsg, i - 1) & vbTab
''
''                lsMsg = Mid(lsMsg, i + 2)
''
''                i = InStr(1, lsMsg, vbCrLf)
''            Loop
''
''            If Trim(lsMsg) <> "" Then
''                j = j + 1
''                lsEqFlag = lsEqFlag & lsMsg & vbTab
''            End If
''
''            For i = 1 To iNone - j
''                lsEqFlag = lsEqFlag & "" & vbTab
''            Next i
''
''            lsEqFlag = vbTab & lsEqFlag
'        End If
'
'        'sRet = Online_Result(lsSpcNo, lsExamCode, lsResult, lsEquip, CStr(iNone), lsEqFlag)
'
'        If Left(lsSpcNo, 1) <> "9" Then
'            If gWorker_Info.WK_ID = "" Then
'                sRet = Online_Result(lsSpcNo, lsExamCode, lsResult, lsEquip, CStr(iNone), lsEqFlag)
'            Else
'                sRet = Online_Result_New1(lsSpcNo, lsExamCode, lsResult, lsEquip, CStr(iNone), lsEqFlag, gWorker_Info.WK_ID)
'            End If
'        Else
'            sRet = QCOnline_Result1(lsSpcNo, lsLotNo, lsExamCode, lsResult, IPU2.Equip_CD, CStr(iNone), lsResDate)
'        End If
'        SaveData lsID & vbTab & gOnline_Ret1
'        If sRet = "N" Then
'            SQL = "Update pat_res set sendflag = 'B' " & vbCrLf & _
'                  "where equipno = '" & gEquip & "' " & vbCrLf & _
'                  "  and barcode = '" & lsID & "' "
'            res = SendQuery(gLocal, SQL)
'            If res = -1 Then
'                SaveQuery SQL
'                Exit Function
'            End If
'
'            ToServer1_원본 = 1
'        Else
'            ToServer1_원본 = -1
'        End If
'    Else
'        ToServer1_원본 = 0
'    End If

End Function


Function GetPatientInfo1(ByVal asID As String, ByVal asRow As Long) As Integer

    Dim lsID As String
    Dim lRow As Long
    
    If IsNumeric(asID) = False Or Len(asID) < 11 Then Exit Function
    
    lsID = Trim(asID)
    lRow = asRow
    
    With vasList
        res = Online_XML(gXml_S03, lsID)
        If res = 1 Then
            .SetText 3, lRow, gPat_Info_Select.PT_NO
            .SetText 4, lRow, gPat_Info_Select.PT_NM
            .SetText 6, lRow, gPat_Info_Select.Sex
            .SetText 7, lRow, gPat_Info_Select.Age
            .SetText 8, lRow, gPat_Info_Select.ORD_SITE
            .SetText 9, lRow, gPat_Info_Select.ACPTNO_1
            
            GetPatientInfo1 = 1
        Else
            SetBackColor vasList, lRow, lRow, 1, 1, 255, 0, 0
            'SetText vasList, "미접수", lRow, gResCol       '2010.04.19 이상은 - 이 부분때문에 항상 WBC 못 가져옴
            
            GetPatientInfo1 = 0
        End If
    End With
End Function

Function GetPatientInfo2(ByVal asID As String, ByVal asRow As Long) As Integer
    Dim lsID As String
    Dim lRow As Long
    
    If IsNumeric(asID) = False Or Len(asID) < 11 Then Exit Function
    
    lsID = Trim(asID)
    lRow = asRow
    
    With vasList
        res = Online_XML1(gXml_S03, lsID)
        If res = 1 Then
            .SetText 3, lRow, gPat_Info_Select1.PT_NO
            .SetText 4, lRow, gPat_Info_Select1.PT_NM
            .SetText 6, lRow, gPat_Info_Select1.Sex
            .SetText 7, lRow, gPat_Info_Select1.Age
            .SetText 8, lRow, gPat_Info_Select1.ORD_SITE
            .SetText 9, lRow, gPat_Info_Select1.ACPTNO_1
            
            GetPatientInfo2 = 1
        Else
            SetBackColor vasList, lRow, lRow, 1, 1, 255, 0, 0
            'SetText vasList, "미접수", lRow, gResCol       '2010.04.19 이상은 - 이 부분때문에 항상 WBC 못 가져옴
            
            GetPatientInfo2 = 0
        End If
    End With
End Function

Function Get_Order1(ByVal asID As String) As Integer
    Dim lsID As String
    Dim lRow As Long
    
    ClearSpread vasTemp1
    
    lsID = Trim(asID)
    
'    res = Get_Order(lsID)
'    For lRow = 0 To UBound(gOrder_List)
'        vasTemp1.SetText 1, lRow + 1, gOrder_List(lRow).TST_CD
'
'        SQL = "select equipcode, examname, OrdGubun from equipexam " & vbCrLf & _
'              "Where Equip = '" & gEquip & "' " & CR & _
'              "  and examcode = '" & Trim(GetText(vasTemp1, lRow + 1, 1)) & "' "
'        res = db_select_Col(gLocal, SQL)
'        vasTemp1.SetText 4, lRow + 1, Trim(gReadBuf(0))
'        vasTemp1.SetText 5, lRow + 1, Trim(gReadBuf(1))
'        vasTemp1.SetText 6, lRow + 1, Trim(gReadBuf(2))
'    Next lRow
    
    res = Online_XML(gXml_S07, Trim(lsID))
    For lRow = 0 To UBound(gExam_Select)
        vasTemp1.SetText 1, lRow + 1, gExam_Select(lRow).TST_CD

        SQL = "select equipcode, examname, OrdGubun from equipexam " & vbCrLf & _
              "Where Equip = '" & gEquip & "' " & CR & _
              "  and examcode = '" & Trim(GetText(vasTemp1, lRow + 1, 1)) & "' "
        res = db_select_Col(gLocal, SQL)
        vasTemp1.SetText 4, lRow + 1, Trim(gReadBuf(0))
        vasTemp1.SetText 5, lRow + 1, Trim(gReadBuf(1))
        vasTemp1.SetText 6, lRow + 1, Trim(gReadBuf(2))
    Next lRow
    
End Function

Function Get_Order2(ByVal asID As String) As Integer
    Dim lsID As String
    Dim lRow As Long
    
    ClearSpread vasTemp2
    
    lsID = Trim(asID)
            
'    res = Get_Order_1(lsID)
'    For lRow = 0 To UBound(gOrder_List1)
'        vasTemp2.SetText 1, lRow + 1, gOrder_List1(lRow).TST_CD
'
'        SQL = "select equipcode, examname, OrdGubun from equipexam " & vbCrLf & _
'              "Where Equip = '" & gEquip & "' " & CR & _
'              "  and examcode = '" & Trim(GetText(vasTemp2, lRow + 1, 1)) & "' "
'        res = db_select_Col(gLocal, SQL)
'        vasTemp2.SetText 4, lRow + 1, Trim(gReadBuf(0))
'        vasTemp2.SetText 5, lRow + 1, Trim(gReadBuf(1))
'        vasTemp2.SetText 6, lRow + 1, Trim(gReadBuf(2))
'    Next lRow
            
    res = Online_XML1(gXml_S07, Trim(lsID))
    For lRow = 0 To UBound(gExam_Select1)
        vasTemp1.SetText 1, lRow + 1, gExam_Select1(lRow).TST_CD

        SQL = "select equipcode, examname, OrdGubun from equipexam " & vbCrLf & _
              "Where Equip = '" & gEquip & "' " & CR & _
              "  and examcode = '" & Trim(GetText(vasTemp1, lRow + 1, 1)) & "' "
        res = db_select_Col(gLocal, SQL)
        vasTemp1.SetText 4, lRow + 1, Trim(gReadBuf(0))
        vasTemp1.SetText 5, lRow + 1, Trim(gReadBuf(1))
        vasTemp1.SetText 6, lRow + 1, Trim(gReadBuf(2))
    Next lRow
End Function

Function SendOrder1(ByVal aiLevel As Integer) As String
    Dim lsOrder As String
    Dim lsExam As String
    Dim lRow, i, j, k As Long
    Dim iCBC As Integer
    Dim iDiff As Integer
    Dim iReti As Integer
    Dim iNRBC As Integer
    Dim lsPID, lsPName, lsAcpNo As String
    
    lsOrder = ""
    
    Select Case aiLevel
    Case 1  'Header
        lsOrder = "H|\^&|||||||||||" & gsVersion & chrCR & chrETX
        lsOrder = chrSTX & CCur(aiLevel) & lsOrder & CheckSum(CStr(1) & lsOrder) & chrCR & chrLF
    Case 2  'Patient
        res = Online_XML(gXml_S03, gID)
        If res > 0 Then
            lsPID = gPat_Info_Select.PT_NO
            lsAcpNo = gPat_Info_Select.ACPTNO_1
            lsPName = gPat_Info_Select.PT_NM
            lsPName = Trim(UCase(Conv_Kor_Eng_1(lsPName)))
            gPatGen.Birth = Format(DateAdd("yyyy", 0 - gPat_Info_Select.Age, Date), "yyyymmdd")
           
            lsOrder = "P|1|||" & lsPID & "|^" & lsPName & "^||" & _
                        gPatGen.Birth & "|" & gPatGen.Sex & "|||||||||||||||||^^^|" & chrCR & chrETX
        Else
            lsOrder = "P|1|" & chrCR & chrETX
        End If
        
        lsOrder = chrSTX & CCur(aiLevel) & lsOrder & CheckSum(CStr(2) & lsOrder) & chrCR & chrLF
        
        iMsgCnt1 = 2
        
        gOrder1 = ""
    
    Case 3  'Order
        
        iMsgCnt1 = iMsgCnt1 + 1
        If iMsgCnt1 = 8 Then
            iMsgCnt1 = 0
        End If
        
        If Trim(gOrder1) <> "" Then
            lsOrder = CStr(iMsgCnt1) & gOrder1 & chrCR & chrETX
            lsOrder = chrSTX & lsOrder & CheckSum(lsOrder) & chrCR & chrLF
            
            gOrder1 = ""
            
            SendOrder1 = lsOrder
            Exit Function
        End If
        
        lsExam = ""
        
        lRow = vasList.DataRowCnt + 1
        If lRow > vasList.MaxRows Then vasList.MaxRows = lRow
        
        vasList.SetText 2, lRow, gID
        vasList.SetText 11, lRow, gRack
        vasList.SetText 12, lRow, gPos
        
        If InStr(1, gID, "ERR") < 1 And IsNumeric(gID) = True Then
            GetPatientInfo1 gID, lRow
            Get_Order1 gID
        End If
         
        DeleteRow vasList, lRow, lRow
        
        j = 0
        iCBC = 0
        iDiff = 0
        iReti = 0
        iNRBC = 0
        
        If vasTemp1.DataRowCnt > 0 Then
            For i = 1 To vasTemp1.DataRowCnt
                Select Case Trim(GetText(vasTemp1, i, 6))
                Case "C"
                    iCBC = 1
                Case "D"
                    iDiff = 1
                Case "R"
                    iReti = 1
                Case "B"
                    iNRBC = 1
                End Select
'                If Trim(GetText(vasTemp1, i, 14)) <> "N" And Trim(GetText(vasTemp1, i, 13)) <> "" Then
'                    lsExam = lsExam & "^^^" & Trim(GetText(vasTemp1, i, 13)) & "\"
'                    j = 1
'                End If
            Next i
        End If
           
'        SaveData iCBC & iDiff & iReti & iNRBC
        
        If iCBC = 0 And iDiff = 0 Then
            iCBC = 1
            iDiff = 1
        ElseIf iCBC = 0 And iDiff = 1 Then
            iCBC = 1
            iDiff = 1
        End If
        lsExam = ""
        If iCBC = 1 Then
            lsExam = lsExam & "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\"
            lsExam = lsExam & "^^^^RDW-CV\^^^^RDW-SD\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\"
        End If
        If iDiff = 1 Then
            lsExam = lsExam & "^^^^NEUT%\^^^^LYMPH%\^^^^MONO%\^^^^EO%\^^^^BASO%\"
                              
            lsExam = lsExam & "^^^^NEUT#\^^^^LYMPH#\^^^^MONO#\^^^^EO#\^^^^BASO#\"
        End If
        If iReti = 1 Then
            lsExam = lsExam & "^^^^RET%\^^^^RET#\^^^^IRF\^^^^LFR\^^^^MFR\^^^^HFR\^^^^RET-HE\^^^^IPF\"
        End If
        If iNRBC = 1 Then
            lsExam = lsExam & "^^^^NRBC#\^^^^NRBC%\"
        End If

        
'        SaveData lsExam
        
        lsExam = Left(lsExam, Len(lsExam) - 1)
        lsOrder = "O|1|^^" & SetSpace(gID, 15) & "^" & gFlag & "||" & lsExam & "||" & Format(Date, "yyyymmdd") & Format(Time, "hhnnss") & "|||||N||||||||||||||Q"
        If Len(lsOrder) > 240 Then
            For k = 240 To 1 Step -1
                If Mid(lsOrder, k, 1) = "\" Then
                    Exit For
                End If
            Next k
            
            gOrder1 = Mid(lsOrder, k + 1)
            giLevel = giLevel - 1
            
            lsOrder = Left(lsOrder, k)
            lsOrder = iMsgCnt1 & lsOrder & chrCR & chrETB
        Else
            lsOrder = iMsgCnt1 & lsOrder & chrCR & chrETX
            gOrder1 = ""
        End If
        
        'lsOrder = "O|" & imsgflag & "|^^" & gID & "^||" & lsExam & "||" & Format(Date, "yyyymmdd") & Format(Time, "hhnnss") & "|||||N||||||||||||||Q" & chrCR & chrETX
        lsOrder = chrSTX & lsOrder & CheckSum(lsOrder) & chrCR & chrLF
        
        
    Case 4  'Comment
        iMsgCnt1 = iMsgCnt1 + 1
        If iMsgCnt1 = 8 Then
            iMsgCnt1 = 0
        End If
        
        lsOrder = "C|1||" & gPat_Info_Select.ACPTNO_1 & chrCR & chrETX
        lsOrder = chrSTX & CCur(iMsgCnt1) & lsOrder & CheckSum(CStr(iMsgCnt1) & lsOrder) & chrCR & chrLF
    Case 5  'Message End
        iMsgCnt1 = iMsgCnt1 + 1
        If iMsgCnt1 = 8 Then
            iMsgCnt1 = 0
        End If
        
        lsOrder = "L|1|N" & chrCR & chrETX
        lsOrder = chrSTX & CCur(iMsgCnt1) & lsOrder & CheckSum(CStr(iMsgCnt1) & lsOrder) & chrCR & chrLF
    Case 6  'EOT
        lsOrder = chrEOT
        giLevel = 0
    End Select
    
    SendOrder1 = lsOrder
    
End Function


Function SendOrder2(ByVal aiLevel As Integer) As String
    Dim lsOrder As String
    Dim lsExam As String
    Dim lRow, i, j, k As Long
    Dim iCBC As Integer
    Dim iDiff As Integer
    Dim iReti As Integer
    Dim iNRBC As Integer
    Dim lsPID, lsAcpNo, lsPName As String
    
    lsOrder = ""
    
    Select Case aiLevel
    Case 1  'Header
        lsOrder = "H|\^&|||||||||||" & gsVersion & chrCR & chrETX
        lsOrder = chrSTX & CCur(aiLevel) & lsOrder & CheckSum(CStr(1) & lsOrder) & chrCR & chrLF
    Case 2  'Patient
'        res = Get_PatInfo1(gID1)
'        If res > 0 Then
'            lsPID = gPatient_Info1.PTNO
'            lsAcpNo = gPatient_Info1.ACPT_NO
'            lsPName = gPatient_Info1.PATNAME
'            lsPName = Trim(UCase(Conv_Kor_Eng_1(lsPName)))
'            gPatGen.Birth = Format(DateAdd("yyyy", 0 - gPatient_Info1.Age, Date), "yyyymmdd")
            
        res = Online_XML1(gXml_S03, gID1)
        If res > 0 Then
            lsPID = gPat_Info_Select1.PT_NO
            lsAcpNo = gPat_Info_Select1.ACPTNO_1
            lsPName = gPat_Info_Select1.PT_NM
            lsPName = Trim(UCase(Conv_Kor_Eng_1(lsPName)))
            gPatGen.Birth = Format(DateAdd("yyyy", 0 - gPat_Info_Select1.Age, Date), "yyyymmdd")
            
            lsOrder = "P|1|||" & lsPID & "|^" & lsPName & "^||" & _
                        gPatGen.Birth & "|" & gPatGen.Sex & "|||||||||||||||||^^^|" & chrCR & chrETX
        Else
            lsOrder = "P|1|" & chrCR & chrETX
        End If
        
        lsOrder = chrSTX & CCur(aiLevel) & lsOrder & CheckSum(CStr(2) & lsOrder) & chrCR & chrLF
        iMsgCnt2 = 2
        gOrder2 = ""
    Case 3  'Order
        
        iMsgCnt2 = iMsgCnt2 + 1
        If iMsgCnt2 = 8 Then
            iMsgCnt2 = 0
        End If
        
        If Trim(gOrder2) <> "" Then
            lsOrder = CStr(iMsgCnt2) & gOrder2 & chrCR & chrETX
            lsOrder = chrSTX & lsOrder & CheckSum(lsOrder) & chrCR & chrLF
            
            gOrder2 = ""
            
            SendOrder2 = lsOrder
            Exit Function
        End If
        
        lsExam = ""
        
        lRow = vasList.DataRowCnt + 1
        If lRow > vasList.MaxRows Then vasList.MaxRows = lRow
        
        vasList.SetText 2, lRow, gID1
        vasList.SetText 11, lRow, gRack1
        vasList.SetText 12, lRow, gPos1
        
        If InStr(1, gID1, "ERR") < 1 And IsNumeric(gID1) = True Then
            GetPatientInfo2 gID1, lRow
            Get_Order2 gID1
        End If
         
        DeleteRow vasList, lRow, lRow
        
        j = 0
        iCBC = 0
        iDiff = 0
        iReti = 0
        iNRBC = 0
        
        If vasTemp2.DataRowCnt > 0 Then
            For i = 1 To vasTemp2.DataRowCnt
                Select Case Trim(GetText(vasTemp2, i, 6))
                Case "C"
                    iCBC = 1
                Case "D"
                    iDiff = 1
                Case "R"
                    iReti = 1
                Case "B"
                    iNRBC = 1
                End Select
'                If Trim(GetText(vasTemp2, i, 14)) <> "N" And Trim(GetText(vasTemp2, i, 13)) <> "" Then
'                    lsExam = lsExam & "^^^" & Trim(GetText(vasTemp2, i, 13)) & "\"
'                    j = 1
'                End If
            Next i
        End If
            
'        SaveData iCBC & iDiff & iReti & iNRBC
        
        If iCBC = 0 And iDiff = 0 Then
            iCBC = 1
            iDiff = 1
        ElseIf iCBC = 0 And iDiff = 1 Then
            iCBC = 1
            iDiff = 1
        End If
        lsExam = ""
        If iCBC = 1 Then
            lsExam = lsExam & "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\"
            lsExam = lsExam & "^^^^RDW-CV\^^^^RDW-SD\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\"
        End If
        If iDiff = 1 Then
            lsExam = lsExam & "^^^^NEUT%\^^^^LYMPH%\^^^^MONO%\^^^^EO%\^^^^BASO%\"
                              
            lsExam = lsExam & "^^^^NEUT#\^^^^LYMPH#\^^^^MONO#\^^^^EO#\^^^^BASO#\"
        End If
        If iReti = 1 Then
            lsExam = lsExam & "^^^^RET%\^^^^RET#\^^^^MFR\^^^^HFR\^^^^LFR\^^^^IRF\^^^^RET-HE\^^^^IPF\"
        End If
        If iNRBC = 1 Then
            lsExam = lsExam & "^^^^NRBC#\^^^^NRBC%\"
        End If

        
'        SaveData lsExam
        
        lsExam = Left(lsExam, Len(lsExam) - 1)
        lsOrder = "O|1|^^" & SetSpace(gID1, 15) & "^" & gFlag1 & "||" & lsExam & "||" & Format(Date, "yyyymmdd") & Format(Time, "hhnnss") & "|||||N||||||||||||||Q"
        If Len(lsOrder) > 240 Then
'            gOrder2 = Mid(lsOrder, 241)
'            giLevel1 = giLevel1 - 1
'
'            lsOrder = Left(lsOrder, 240)
            
            For k = 240 To 1 Step -1
                If Mid(lsOrder, k, 1) = "\" Then
                    Exit For
                End If
            Next k
            
            gOrder2 = Mid(lsOrder, k + 1)
            giLevel1 = giLevel1 - 1
            
            lsOrder = Left(lsOrder, k)
            
            lsOrder = iMsgCnt2 & lsOrder & chrCR & chrETB
        Else
            lsOrder = iMsgCnt2 & lsOrder & chrCR & chrETX
            gOrder2 = ""
        End If
        
        'lsOrder = "O|" & imsgflag1 & "|^^" & gID1 & "^||" & lsExam & "||" & Format(Date, "yyyymmdd") & Format(Time, "hhnnss") & "|||||N||||||||||||||Q" & chrCR & chrETX
        lsOrder = chrSTX & lsOrder & CheckSum(lsOrder) & chrCR & chrLF
        
    Case 4  'Comment
        iMsgCnt2 = iMsgCnt2 + 1
        If iMsgCnt2 = 8 Then
            iMsgCnt2 = 0
        End If
        
        lsOrder = "C|1||" & gPat_Info_Select1.ACPTNO_1 & chrCR & chrETX
        lsOrder = chrSTX & CCur(iMsgCnt2) & lsOrder & CheckSum(CStr(iMsgCnt2) & lsOrder) & chrCR & chrLF
        
    Case 5  'Message End
        iMsgCnt2 = iMsgCnt2 + 1
        If iMsgCnt2 = 8 Then
            iMsgCnt2 = 0
        End If
        
        lsOrder = "L|1|N" & chrCR & chrETX
        lsOrder = chrSTX & CCur(iMsgCnt2) & lsOrder & CheckSum(CStr(iMsgCnt2) & lsOrder) & chrCR & chrLF
        
    Case 6  'EOT
        lsOrder = chrEOT
        giLevel1 = 0
        
    End Select
    
    SendOrder2 = lsOrder
    
End Function



Private Sub vasSchESR_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow, i, liEquipCode As Long
    Dim lsID As String
    Dim liRet As Integer
    'Dim lsID As String
    Dim lsResult As String
    Dim mExam
    
    If KeyCode = vbKeyReturn Then
        lRow = vasSch.ActiveRow
        
        SQL = "Select barcode, diskno, posno, examtype from pat_res where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' and barcode = '" & Trim(GetText(vasSch, lRow, 2)) & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = Trim(GetText(vasSch, lRow, 2)) Then
            If MsgBox("입력하신 검체 [" & Trim(GetText(vasSch, lRow, 2)) & "]는 장비 " & Trim(gReadBuf(3)) & "의 " & Trim(gReadBuf(1)) & " Rack " & Trim(gReadBuf(2)) & " Position 에서 검사한 것입니다 " & vbCrLf & _
                      " " & vbCrLf & _
                      "저장하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, "알림") = vbNo Then
                SetText vasSch, Trim(GetText(vasSch, lRow, gMaxCol + 4)), lRow, 2
                Exit Sub
            End If
        End If
        
        If MsgBox("결과를 전송하시겠습니까? ", vbCritical + vbYesNo + vbDefaultButton2, "알림") = vbYes Then
            
            lsID = Trim(GetText(vasSch, lRow, 2))
            
            If GetPatientInfo1(lsID, lRow) = 1 Then
                
                SQL = "Update pat_res set " & vbCrLf & _
                      "  barcode = '" & lsID & "', " & vbCrLf & _
                      "  pid = '" & Trim(GetText(vasSch, lRow, 3)) & "', " & vbCrLf & _
                      "  pname = '" & Trim(GetText(vasSch, lRow, 4)) & "', " & vbCrLf & _
                      "  pjumin = '" & Trim(GetText(vasSch, lRow, 5)) & "', " & vbCrLf & _
                      "  psex = '" & Trim(GetText(vasSch, lRow, 6)) & "', " & vbCrLf & _
                      "  page1 = '" & Trim(GetText(vasSch, lRow, 7)) & "', " & vbCrLf & _
                      "  WardRoom = '" & Trim(GetText(vasSch, lRow, 8)) & "', " & vbCrLf & _
                      "  receno = '" & Trim(GetText(vasSch, lRow, 9)) & "' " & vbCrLf & _
                      "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                      "  and barcode = '" & Trim(GetText(vasSch, lRow, gMaxCol + 4)) & "'"
                res = SendQuery(gLocal, SQL)
                
'                SaveData lRow & " : " & lsID
                res = ToServer(lRow, vasSch, 1)
                If res = 1 Then
                    vasSch.Row = lRow
                    vasSch.Col = 1
                    vasSch.Value = 0
    
                    SetText vasSch, "완료", lRow, gResCol
                    SetBackColor vasSch, lRow, lRow, 1, 1, 202, 255, 112
                    'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
                
                    SQL = "Update worklist set OrdFlag = 'D' where barcode = '" & lsID & "' "
                    res = SendQuery(gLocal, SQL)
                
                ElseIf res = 2 Then
                    SetText vasSch, "결과", lRow, gResCol
                Else
                    SetText vasSch, "실패", lRow, gResCol
                    SetForeColor vasSch, lRow, lRow, 1, 1, 255, 0, 0
                    'SetBackColor vasID, gPreRow, gPreRow, colCheckBox, vasID.MaxCols, 255, 255, 255
                    'vasID.SetCellBorder 1, gPreRow, vasID.MaxCols, gPreRow, 15, &H8000000F, CellBorderStyleSolid
                End If
            
            Else
                vasSch.Row = lRow
                vasSch.Col = 1
                vasSch.Value = 1
                SetBackColor vasSch, lRow, lRow, 1, 1, 255, 0, 0
                SetText vasSch, "미접수", lRow, gResCol
                
                Exit Sub
            End If
            
        End If
    End If

End Sub

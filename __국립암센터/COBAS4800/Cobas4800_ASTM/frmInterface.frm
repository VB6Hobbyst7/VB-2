VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   Caption         =   " Cobas4800 Interface"
   ClientHeight    =   10650
   ClientLeft      =   1680
   ClientTop       =   750
   ClientWidth     =   18255
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInterface.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmInterface.frx":113A
   ScaleHeight     =   10650
   ScaleWidth      =   18255
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture2 
      Align           =   1  '위 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   18195
      TabIndex        =   84
      Top             =   0
      Width           =   18255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1110
         Picture         =   "frmInterface.frx":13BD
         ScaleHeight     =   255
         ScaleWidth      =   315
         TabIndex        =   87
         Top             =   450
         Width           =   345
      End
      Begin VB.TextBox Text_Today 
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1110
         TabIndex        =   86
         Text            =   "2002/02/18"
         Top             =   90
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5700
         TabIndex        =   100
         Top             =   480
         Width           =   285
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5700
         TabIndex        =   99
         Top             =   150
         Width           =   285
      End
      Begin VB.Label lblIFState 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "연결대기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4650
         TabIndex        =   98
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblConnectState 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "연결대기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4620
         TabIndex        =   97
         Top             =   150
         Width           =   1035
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과통신상태 ["
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   3330
         TabIndex        =   96
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "오더통신상태 ["
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3330
         TabIndex        =   95
         Top             =   150
         Width           =   1260
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   615
         Left            =   15090
         TabIndex        =   94
         Top             =   90
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         _Version        =   131072
         Font3D          =   1
         ForeColor       =   12583104
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":1947
         Caption         =   "결과출력    "
         PictureAlignment=   4
      End
      Begin Threed.SSCommand cmdSend 
         Height          =   615
         Left            =   13560
         TabIndex        =   93
         Top             =   90
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         _Version        =   131072
         Font3D          =   3
         ForeColor       =   12582912
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":1C61
         Caption         =   "결과전송    "
         PictureAlignment=   4
      End
      Begin Threed.SSCommand cmdReset 
         Height          =   615
         Left            =   12030
         TabIndex        =   92
         Top             =   90
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         _Version        =   131072
         Font3D          =   1
         ForeColor       =   16512
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":1F7B
         Caption         =   "초기화    "
         PictureAlignment=   4
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   615
         Left            =   16620
         TabIndex        =   91
         Top             =   90
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         _Version        =   131072
         Font3D          =   3
         ForeColor       =   16512
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":2295
         Caption         =   "종  료    "
         ButtonStyle     =   2
         PictureAlignment=   4
      End
      Begin Threed.SSCommand cmdCall 
         Height          =   615
         Left            =   10500
         TabIndex        =   90
         Top             =   90
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         _Version        =   131072
         Font3D          =   3
         ForeColor       =   192
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":25AF
         Caption         =   "결과조회    "
         PictureAlignment=   4
      End
      Begin Threed.SSCommand cmdResProc 
         Height          =   615
         Left            =   9030
         TabIndex        =   89
         Top             =   90
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         _Version        =   131072
         Font3D          =   3
         ForeColor       =   192
         BackColor       =   16777215
         PictureFrames   =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":28C9
         Caption         =   "결과처리    "
         ButtonStyle     =   2
         PictureAlignment=   4
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '투명
         Caption         =   "사용자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1590
         TabIndex        =   88
         Top             =   510
         Width           =   675
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
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
         Left            =   120
         TabIndex        =   85
         Top             =   150
         Width           =   780
      End
   End
   Begin VB.Frame FrmTempBox 
      Caption         =   "TempBox"
      Height          =   9525
      Left            =   19320
      TabIndex        =   50
      Top             =   1500
      Visible         =   0   'False
      Width           =   8985
      Begin VB.TextBox txtXMLRes 
         Height          =   645
         Left            =   180
         TabIndex        =   102
         Top             =   2370
         Width           =   8055
      End
      Begin VB.TextBox Text6 
         Height          =   675
         Left            =   6840
         TabIndex        =   80
         Top             =   5070
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.TextBox txtDataTmp 
         Height          =   615
         Left            =   5100
         TabIndex        =   79
         Top             =   5070
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   1635
         Left            =   330
         TabIndex        =   73
         Top             =   5880
         Visible         =   0   'False
         Width           =   6285
         Begin VB.CommandButton Command18 
            Caption         =   "s07"
            Height          =   495
            Left            =   2670
            TabIndex        =   78
            Top             =   300
            Width           =   675
         End
         Begin VB.TextBox Text4 
            Height          =   495
            Left            =   120
            TabIndex        =   77
            Text            =   "barcode"
            Top             =   300
            Width           =   2505
         End
         Begin VB.CommandButton Command19 
            Caption         =   "S03"
            Height          =   495
            Left            =   3390
            TabIndex        =   76
            Top             =   300
            Width           =   675
         End
         Begin VB.TextBox txtXML 
            Height          =   585
            Left            =   120
            TabIndex        =   75
            Top             =   900
            Width           =   6045
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Command20"
            Height          =   435
            Left            =   4920
            TabIndex        =   74
            Top             =   300
            Width           =   915
         End
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Command21"
         Height          =   465
         Left            =   3810
         TabIndex        =   72
         Top             =   5100
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Command17"
         Height          =   465
         Left            =   2640
         TabIndex        =   71
         Top             =   5100
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Command16"
         Height          =   465
         Left            =   1380
         TabIndex        =   70
         Top             =   5100
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   825
         Left            =   210
         TabIndex        =   64
         Top             =   4080
         Width           =   8595
      End
      Begin VB.TextBox txtBuff 
         Height          =   765
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   63
         Top             =   3240
         Width           =   8625
      End
      Begin FPSpread.vaSpread vasIDTmp 
         Height          =   1275
         Left            =   3210
         TabIndex        =   60
         Top             =   990
         Width           =   1455
         _Version        =   393216
         _ExtentX        =   2566
         _ExtentY        =   2249
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
         SpreadDesigner  =   "frmInterface.frx":2BE3
      End
      Begin VB.CommandButton cmdQC 
         Caption         =   "QC"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   4530
         TabIndex        =   58
         Top             =   240
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton Command14 
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
         Height          =   615
         Left            =   7290
         TabIndex        =   57
         Top             =   210
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmdResCall 
         Caption         =   "QC 결과전송"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   5700
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Command15"
         Height          =   465
         Left            =   240
         TabIndex        =   55
         Top             =   5100
         Width           =   1095
      End
      Begin VB.CommandButton Command_setup 
         Caption         =   "코드설정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2310
         TabIndex        =   54
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton Command_close 
         Caption         =   "종료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3420
         TabIndex        =   53
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton Command_Config 
         Caption         =   "통신설정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1200
         TabIndex        =   52
         Top             =   240
         Width           =   1065
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   585
         Left            =   90
         Style           =   1  '그래픽
         TabIndex        =   51
         Top             =   240
         Value           =   1  '확인
         Width           =   1065
      End
      Begin FPSpread.vaSpread vasOrderTmp 
         Height          =   1275
         Left            =   180
         TabIndex        =   59
         Top             =   990
         Width           =   1455
         _Version        =   393216
         _ExtentX        =   2566
         _ExtentY        =   2249
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
         SpreadDesigner  =   "frmInterface.frx":2E09
      End
      Begin FPSpread.vaSpread vasOrder 
         Height          =   1260
         Left            =   1650
         TabIndex        =   61
         Top             =   990
         Visible         =   0   'False
         Width           =   1485
         _Version        =   393216
         _ExtentX        =   2619
         _ExtentY        =   2222
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
         MaxCols         =   2
         SpreadDesigner  =   "frmInterface.frx":302F
      End
      Begin FPSpread.vaSpread vasOrderBuf 
         Height          =   1305
         Left            =   4710
         TabIndex        =   62
         Top             =   990
         Visible         =   0   'False
         Width           =   1305
         _Version        =   393216
         _ExtentX        =   2302
         _ExtentY        =   2302
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
         MaxCols         =   2
         SpreadDesigner  =   "frmInterface.frx":6B94
      End
      Begin FPSpread.vaSpread vasASTM 
         Height          =   1305
         Left            =   6090
         TabIndex        =   81
         Top             =   990
         Visible         =   0   'False
         Width           =   1305
         _Version        =   393216
         _ExtentX        =   2302
         _ExtentY        =   2302
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
         MaxCols         =   1
         SpreadDesigner  =   "frmInterface.frx":A704
      End
      Begin FPSpread.vaSpread vasWork 
         Height          =   1305
         Left            =   7440
         TabIndex        =   82
         Top             =   990
         Visible         =   0   'False
         Width           =   1365
         _Version        =   393216
         _ExtentX        =   2408
         _ExtentY        =   2302
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
         SpreadDesigner  =   "frmInterface.frx":BF02
      End
      Begin FPSpread.vaSpread vasPrint 
         Height          =   6705
         Left            =   240
         TabIndex        =   83
         Top             =   7680
         Visible         =   0   'False
         Width           =   12435
         _Version        =   393216
         _ExtentX        =   21934
         _ExtentY        =   11827
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   9
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frmInterface.frx":103A2
      End
   End
   Begin VB.Frame FrmUseControl 
      Caption         =   "UseControl"
      Height          =   1155
      Left            =   19140
      TabIndex        =   49
      Top             =   480
      Visible         =   0   'False
      Width           =   4785
      Begin MSWinsockLib.Winsock Winsock2 
         Left            =   2850
         Top             =   390
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer Timer2 
         Left            =   1680
         Top             =   360
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   2250
         Top             =   420
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   990
         Top             =   360
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   150
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         InputLen        =   1
         RThreshold      =   1
         RTSEnable       =   -1  'True
         EOFEnable       =   -1  'True
      End
      Begin MSComDlg.CommonDialog cdResProc 
         Left            =   3420
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10275
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   15002
            MinWidth        =   14995
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   11817
            MinWidth        =   11817
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2018-12-04"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오후 3:38"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   9315
      Left            =   60
      TabIndex        =   1
      Top             =   900
      Width           =   18105
      Begin Threed.SSPanel SSPanel3 
         Height          =   315
         Left            =   210
         TabIndex        =   107
         Top             =   360
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         _Version        =   131072
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "번호"
         BevelOuter      =   0
      End
      Begin VB.ComboBox cboTestIdList 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13860
         Sorted          =   -1  'True
         TabIndex        =   105
         Text            =   "Combo1"
         Top             =   630
         Width           =   2895
      End
      Begin VB.TextBox txteGFRCmnt 
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
         Height          =   6180
         Left            =   8640
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   103
         Top             =   3000
         Width           =   9315
      End
      Begin VB.TextBox txtTestWay 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9660
         TabIndex        =   67
         Top             =   270
         Width           =   7065
      End
      Begin VB.TextBox txtTestIdName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9660
         TabIndex        =   66
         Top             =   630
         Width           =   3615
      End
      Begin FPSpread.vaSpread vasXML 
         Height          =   3255
         Left            =   180
         TabIndex        =   65
         Top             =   9600
         Visible         =   0   'False
         Width           =   12975
         _Version        =   393216
         _ExtentX        =   22886
         _ExtentY        =   5741
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
         SpreadDesigner  =   "frmInterface.frx":12004
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   1545
         Left            =   8640
         TabIndex        =   48
         Top             =   1080
         Width           =   9315
         _Version        =   393216
         _ExtentX        =   16431
         _ExtentY        =   2725
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":1222A
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "Check1"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   690
         TabIndex        =   46
         Top             =   330
         Width           =   195
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   8955
         Left            =   150
         TabIndex        =   47
         Top             =   270
         Width           =   8385
         _Version        =   393216
         _ExtentX        =   14790
         _ExtentY        =   15796
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   17
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":13D3D
         UserResize      =   2
      End
      Begin Threed.SSCommand cmdDel 
         Height          =   345
         Left            =   13320
         TabIndex        =   106
         Top             =   630
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         _Version        =   131072
         Font3D          =   3
         ForeColor       =   12582912
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Del"
         PictureAlignment=   4
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "[ eGFR 코멘트 ]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   8640
         TabIndex        =   104
         Top             =   2730
         Width           =   1530
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   735
         Left            =   16770
         TabIndex        =   101
         Top             =   240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1296
         _Version        =   131072
         Font3D          =   3
         ForeColor       =   12582912
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":15C7F
         Caption         =   "저장     "
         PictureAlignment=   4
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '투명
         Caption         =   "검사방법"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   8730
         TabIndex        =   69
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label12 
         BackStyle       =   0  '투명
         Caption         =   "보 고 자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   8730
         TabIndex        =   68
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   4455
      Left            =   1470
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   13095
      Begin FPSpread.vaSpread vasList 
         Height          =   1125
         Left            =   4260
         TabIndex        =   25
         Top             =   660
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":15F99
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   1575
         Left            =   3720
         TabIndex        =   31
         Top             =   2790
         Width           =   9285
         Begin VB.TextBox txtEquipID 
            Height          =   345
            Left            =   3600
            TabIndex        =   42
            Text            =   "10"
            Top             =   1140
            Width           =   1875
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Rack Pos"
            Height          =   375
            Left            =   7560
            TabIndex        =   41
            Top             =   1110
            Width           =   1635
         End
         Begin VB.CommandButton Command10 
            Caption         =   "결과입력"
            Height          =   375
            Left            =   5880
            TabIndex        =   40
            Top             =   1110
            Width           =   1635
         End
         Begin VB.TextBox txtEquipCode 
            Height          =   345
            Left            =   1710
            TabIndex        =   39
            Text            =   "0ADVI120"
            Top             =   1125
            Width           =   1875
         End
         Begin VB.CommandButton Command9 
            Caption         =   "장비ID조회"
            Height          =   375
            Left            =   60
            TabIndex        =   38
            Top             =   1110
            Width           =   1635
         End
         Begin VB.CommandButton Command8 
            Caption         =   "미검사상세목록"
            Height          =   375
            Left            =   5010
            TabIndex        =   37
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command7 
            Caption         =   "미검사목록"
            Height          =   375
            Left            =   3360
            TabIndex        =   36
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command6 
            Caption         =   "검사상세목록"
            Height          =   375
            Left            =   1710
            TabIndex        =   35
            Top             =   690
            Width           =   1635
         End
         Begin VB.TextBox txtID 
            Height          =   345
            Left            =   6660
            TabIndex        =   34
            Text            =   "05111000003"
            Top             =   720
            Width           =   1875
         End
         Begin VB.CommandButton Command5 
            Caption         =   "검사목록"
            Height          =   375
            Left            =   60
            TabIndex        =   33
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command4 
            Caption         =   "서버시간"
            Height          =   375
            Left            =   60
            TabIndex        =   32
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label lblDate2 
            AutoSize        =   -1  'True
            Caption         =   "서버시간1"
            Height          =   195
            Left            =   1920
            TabIndex        =   44
            Top             =   330
            Width           =   945
         End
         Begin VB.Label lblDate1 
            AutoSize        =   -1  'True
            Caption         =   "서버시간1"
            Height          =   195
            Left            =   3150
            TabIndex        =   43
            Top             =   330
            Width           =   945
         End
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   210
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   3360
         Width           =   945
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
         Height          =   615
         Left            =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   28
         Top             =   300
         Visible         =   0   'False
         Width           =   5835
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   285
         Left            =   60
         TabIndex        =   27
         Top             =   555
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   240
         TabIndex        =   26
         Top             =   1380
         Width           =   3045
      End
      Begin VB.Frame Frame3 
         Height          =   585
         Left            =   60
         TabIndex        =   19
         Top             =   3780
         Visible         =   0   'False
         Width           =   3675
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
            Left            =   1950
            TabIndex        =   22
            Top             =   180
            Width           =   885
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
            Left            =   630
            TabIndex        =   21
            Top             =   180
            Width           =   885
         End
         Begin VB.CommandButton cmdDelete 
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
            Height          =   315
            Left            =   2880
            TabIndex        =   20
            Top             =   180
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "번호"
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
            Left            =   60
            TabIndex        =   24
            Top             =   240
            Width           =   450
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
            Left            =   1530
            TabIndex        =   23
            Top             =   240
            Width           =   360
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
         Height          =   345
         Left            =   240
         TabIndex        =   18
         Top             =   1875
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   10260
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   17
         Top             =   1950
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdUp 
         Height          =   525
         Left            =   1260
         Picture         =   "frmInterface.frx":161BF
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   3240
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton cmdDown 
         Height          =   525
         Left            =   2010
         Picture         =   "frmInterface.frx":162EE
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   3240
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Command13"
         Height          =   285
         Left            =   1710
         TabIndex        =   14
         Top             =   900
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Command12"
         Height          =   285
         Left            =   1710
         TabIndex        =   13
         Top             =   570
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox Text3 
         Height          =   345
         Left            =   240
         TabIndex        =   11
         Top             =   2850
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.CommandButton cmdResSave 
         Caption         =   "결과저장"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5970
         TabIndex        =   10
         Top             =   1500
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   285
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtBarcode 
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
         Left            =   3480
         TabIndex        =   7
         Top             =   1500
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "WorkList 작성"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   30
         TabIndex        =   6
         Top             =   930
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox Text2 
         Height          =   345
         Left            =   240
         TabIndex        =   5
         Top             =   2355
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   285
         Left            =   1710
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1635
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   1125
         Left            =   10740
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":16420
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1125
         Left            =   12000
         TabIndex        =   9
         Top             =   7000
         Visible         =   0   'False
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":1A8C8
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   1125
         Left            =   6300
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":1ED8A
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1125
         Left            =   8925
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":1EFB0
      End
      Begin VB.Label lblMT 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "0"
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
         Left            =   9750
         TabIndex        =   45
         Top             =   2370
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin VB.Menu MnMain 
      Caption         =   "파일"
      Begin VB.Menu MnExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "설정"
      Begin VB.Menu MnTConfig 
         Caption         =   "포트설정"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "코드설정"
      End
      Begin VB.Menu MnComment 
         Caption         =   "EGFR 코멘트설정"
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "전송"
      Begin VB.Menu MnTransAuto 
         Caption         =   "자동"
      End
      Begin VB.Menu MnTransManual 
         Caption         =   "수동"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const colCheckBox = 1
Const colBarCode = 2
Const colSeqNo = 3
Const colReceno = 4
Const colRack = 5
Const colPos = 6
Const colPID = 7
Const colPName = 8
Const colPSex = 9
Const colPAge = 10
Const colPJumin = 11
Const colState = 12

Const colOrd = 13
Const colRes = 14
Const colDate = 15
Const colTime = 16
Const colTestType = 17
Const colSampleNo = 17

Const colEquipCode = 1
Const colExamCode = 2
Const colExamName = 3
Const colResult = 4
Const colSeq = 5
Const colRCheck = 6

'2004/10/21 이상은
'Const colRefLow = 7
Const colResult1 = 7        '장비결과

Const colRefHigh = 8

Dim gRow            As Long

Dim gsBarCode       As String
Dim gsPID           As String
Dim gsRackNo        As String
Dim gsPosNo         As String
Dim gsResDateTime   As String
Dim gsSeqNo         As String
Dim gsTestID        As String
Dim gsExamCode      As String
Dim gsExamName      As String
Dim gsOrder         As String
Dim gsResult        As String

Dim sSampleType     As String

Dim sResult         As String
Dim sResultT        As String

Dim gMT             As String
Dim gComState       As Long
Dim gErrState       As Long

Public gENQFlag     As Integer
Public gNAKFlag     As Integer

Public gPatFlag     As Integer

Public gAttribute   As String

Dim gWBCRes         As String
Dim gNetRes         As String
Dim gLucRes         As String
Dim gLymphRes       As String

Public gRCnt        As Integer

Dim gSelExam        As String


Function Advia_IDSet(asID As String) As String
    Advia_IDSet = "000" & asID
End Function

Function Advia_Init() As String
    Dim lsData As String
    
    gMT = "0"
    gErrState = 0
    
    lsData = gMT & "I " & chrCR & chrLF
    lsData = chrSTX & lsData & LRC(lsData) & chrETX
    
    gComState = 0
    
    MSComm1.Output = lsData
    Timer1.Enabled = True
    SaveData "[Tx]" & lsData
End Function

Function Advia_NoOrder(asID As String) As String
    Dim lsData As String
    
    gMT = Chr(Asc(gMT) + 1)
    If gMT > "Z" Then gMT = "0"
    
    lsData = gMT & "N R " & Advia_IDSet(asID) & chrCR & chrLF
    lsData = chrSTX & lsData & LRC(lsData) & chrETX
    
    gComState = 3
    
    MSComm1.Output = lsData
    Timer1.Enabled = True
    SaveData "[TX]" & lsData
End Function

Function Advia_ResValid() As String
    Dim lsData As String
    
    gMT = Chr(Asc(gMT) + 1)
    If gMT > "Z" Then gMT = "0"
    
    lsData = gMT & "Z   " & Space(6) & " " & Space(6) & " " & " 0" & chrCR & chrLF
    lsData = chrSTX & lsData & LRC(lsData) & chrETX
    
    gComState = 4
    
    MSComm1.Output = lsData
    Timer1.Enabled = True
    SaveData "[TX]" & lsData
    
End Function

Function Advia_Token() As String
    Dim lsData As String
    
    gMT = Chr(Asc(gMT) + 1)
    If gMT > "Z" Then gMT = "0"     'After the last message toggle code, 5Ah(Z), is used, the codes are recycled beginning with 30h.
    
    lsData = gMT & "S          " & chrCR & chrLF
    lsData = chrSTX & lsData & LRC(lsData) & chrETX
    
    gComState = 1
    
    lblMT.Caption = gMT
    DoSleep 1
    
    MSComm1.Output = lsData
    Timer1.Enabled = True
    SaveData "[Tx]" & lsData
End Function

Function Advia_Token_1() As String
    Dim lsData As String
    
    gMT = Chr(Asc(gMT) + 1)
    If gMT > "Z" Then gMT = "0"

    lsData = "S          " & chrCR & chrLF
    lsData = chrSTX & lsData & LRC(lsData) & chrETX
    
    gComState = 1
    
    lblMT.Caption = gMT
    DoSleep 1
    
    MSComm1.Output = lsData
    Timer1.Enabled = True
    SaveData "[Tx]" & lsData
End Function

Function LRC(ByVal asData As String) As String
    Dim i As Integer
    Dim a
    
    a = Asc(Left(asData, 1))
    
    For i = 2 To Len(asData)
        a = a Xor Asc(Mid(asData, i, 1))
    Next i
    
    If a = 3 Then a = 127
    
    LRC = Chr(a)
End Function


Function Result_Set(ByVal asTest As String, ByVal asRes As String) As Integer
    Dim sGiho As String
    Dim sRes As String
    Dim sRes1 As String
    Dim sFormat As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim sValFlag As String
    
    Dim iRCnt
    
    Dim i As Integer
    Dim lResRow As Integer
    
    Result_Set = -1
    
    If Trim(asTest) = "" Then Exit Function
    
    SQL = "Select EquipCode, ExamCode, ExamName, ResGubun, Range, CutOffFlag, " & vbCrLf & _
          " NegValue, NegEqual, PosValue, PosEqual, cutoff" & vbCrLf & _
          "from EquipExam " & vbCrLf & _
          "where Equip = '" & gEquip & "' " & vbCrLf & _
          "  and EquipCode = '" & Trim(asTest) & "' "
    res = db_select_Col(gLocal, SQL)
    If res < 1 Then Exit Function
    If Trim(gReadBuf(0)) <> Trim(asTest) Then Exit Function
    
    sGiho = ""
    sRes = ""
    sRes1 = ""
    
    sExamCode = Trim(gReadBuf(1))
    sExamName = Trim(gReadBuf(2))
    sValFlag = Trim(gReadBuf(10))
    
    If Trim(sExamCode) = "" Then Exit Function
    
    For i = 1 To Len(asRes)
        If IsNumeric(Mid(asRes, i, 1)) = True Or Mid(asRes, i, 1) = "." Then
            sRes = sRes & Mid(asRes, i, 1)
        Else
            sGiho = sGiho & Mid(asRes, i, 1)
        End If
    Next i
    
    Select Case Trim(gReadBuf(3))
    Case "I"
        sRes1 = Format(CCur(sRes), "#0")
        sRes1 = sGiho & sRes1
    
    Case "F"
        sFormat = ""
        For i = 1 To CInt(gReadBuf(4))
            sFormat = sFormat & "0"
        Next i
        sFormat = "0." & sFormat
        sRes1 = Format(CCur(sRes), sFormat)
        
        sRes1 = sGiho & sRes1
    
    Case "T"
'        'CuttOff
        If Trim(gReadBuf(5)) = "1" Then     '크다
            If Trim(gReadBuf(7)) = "1" And Trim(gReadBuf(9)) = "1" Then
                If CCur(sRes) <= CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) >= CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
            ElseIf Trim(gReadBuf(7)) = "1" And Trim(gReadBuf(9)) = "0" Then
                If CCur(sRes) <= CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) > CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
            ElseIf Trim(gReadBuf(7)) = "0" And Trim(gReadBuf(9)) = "1" Then
                If CCur(sRes) < CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) >= CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
            ElseIf Trim(gReadBuf(7)) = "0" And Trim(gReadBuf(9)) = "0" Then
                If CCur(sRes) < CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) > CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
            End If
        ElseIf Trim(gReadBuf(5)) = "2" Then      '작다
            If Trim(gReadBuf(7)) = "1" And Trim(gReadBuf(9)) = "1" Then
                If CCur(sRes) >= CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) <= CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
                
            ElseIf Trim(gReadBuf(7)) = "1" And Trim(gReadBuf(9)) = "0" Then
                If CCur(sRes) >= CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) < CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
                
            ElseIf Trim(gReadBuf(7)) = "0" And Trim(gReadBuf(9)) = "1" Then
                If CCur(sRes) > CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) <= CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
            ElseIf Trim(gReadBuf(7)) = "0" And Trim(gReadBuf(9)) = "0" Then
                If CCur(sRes) > CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) < CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
            End If
        End If
    
    End Select
    
    lResRow = -1
    For i = 1 To vasRes.DataRowCnt
        If Trim(asTest) = Trim(GetText(vasRes, i, colEquipCode)) Then
            lResRow = i
            Exit For
        End If
    Next i
    
    If lResRow = -1 Then
        lResRow = vasRes.DataRowCnt + 1
        If lResRow > vasRes.MaxRows Then
            vasRes.MaxRows = lResRow
        End If
    End If
    
    SetText vasRes, gsBarCode, lResRow, colBarCode      '검체번호
    SetText vasRes, asTest, lResRow, colEquipCode       '장비코드
    SetText vasRes, sExamCode, lResRow, colExamCode     '검사코드
    SetText vasRes, sExamName, lResRow, colExamName     '검사명
    SetText vasRes, sRes1, lResRow, colResult           '검사결과
    SetText vasRes, asRes, lResRow, colResult1          '장비결과
    
    SetText vasID, vasRes.DataRowCnt, glRow, colRes
    If InStr(1, Trim(GetText(vasRes, lResRow, colResult)), "POS") > 0 Then
        vasRes.Row = lResRow
        vasRes.Col = colResult
        vasRes.ForeColor = RGB(205, 55, 0)
    ElseIf InStr(1, Trim(GetText(vasRes, lResRow, colResult)), "Weak-POS") > 0 Then
        vasRes.Row = lResRow
        vasRes.Col = colResult
        vasRes.ForeColor = RGB(55, 0, 205)
    Else
        vasRes.Row = lResRow
        vasRes.Col = colResult
        vasRes.ForeColor = RGB(0, 0, 0)
    End If
    
    Result_Set = lResRow
End Function


Private Sub cboTestIdList_Click()

    txtTestIdName.Text = cboTestIdList.Text

End Sub

Private Sub chkAll_Click()
    Dim iRow As Long
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 0
        Next iRow
    End If

End Sub

Private Sub chkMode_Click()
    If chkMode.Value = 1 Then
        chkMode.Caption = "Auto"
    Else
        chkMode.Caption = "Manual"
    End If
End Sub

Private Sub cmdCall_Click()
    Dim iRow As Long

    ClearSpread vasID
    ClearSpread vasRes
    ClearSpread vasPrint
    
    SQL = "select distinct levelname, '', '', '0', '0', examtime, '', '', '', 'F' " & vbCrLf & _
          "from qc_res " & vbCrLf & _
          "where equipno  = '" & Trim(gEquip) & "' " & vbCrLf & _
          "  and examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' "
    res = db_select_Vas(gLocal, SQL, vasID, 1, 2)
    
    SQL = "select barcode, seqno, receno, diskno, posno, pid, pname, psex, page, jumin, sendflag, count(*), count(*), max(recedate)" & _
          " from pat_res " & _
          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          "  and sendflag in ('B','C') " & vbCrLf & _
          "group by diskno, posno, barcode, seqno, receno, pid, pname, psex, page, jumin, sendflag "
    SQL = SQL & vbCrLf & " Union " & vbCrLf
    SQL = SQL & vbCrLf & _
          "select barcode, seqno, receno, diskno, posno, pid, pname, psex, page, jumin, sendflag, count(*), '0',  max(recedate)" & _
          " from pat_res " & _
          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          "  and sendflag not in ('B','C') " & vbCrLf & _
          "group by diskno, posno, barcode, seqno, receno,  pid, pname, psex, page, jumin, sendflag " & vbCrLf & _
          "order by posno,diskno"
    res = db_select_Vas(gLocal, SQL, vasID, vasID.DataRowCnt + 1, 2)
   
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    vasSort vasID, colDate, colReceno 'colRack, colPos
    
    For iRow = 1 To vasID.DataRowCnt
        Select Case Trim(GetText(vasID, iRow, colState))
        Case "B", "C"
            SetBackColor vasID, iRow, iRow, 1, colState, 202, 255, 112
            SetText vasID, "완료", iRow, colState
        Case "O"
            SetText vasID, "오더", iRow, colState
         Case "A"
            SetText vasID, "결과", iRow, colState
        End Select
    Next iRow

End Sub

Private Sub cmdDel_Click()
    Dim i       As Integer
    Dim strList As String
    
    strList = ""
    For i = 1 To cboTestIdList.ListCount - 1
        If txtTestIdName.Text <> cboTestIdList.List(i) Then
            strList = strList & cboTestIdList.List(i) & ","
        End If
    Next
    If strList <> "" Then
        strList = Mid(strList, 1, Len(strList) - 1)
    End If
    
    Call WritePrivateProfileString("config", "gTestIdList", strList, App.Path & "\Interface.ini")
End Sub

Private Sub cmdDelete_Click()
    Dim lRow As Long
    Dim lsPID As String
    Dim lsReceNo1 As String
    Dim lsReceNo2 As String
    
    Dim sStart As String
    Dim send As String
    
    sStart = Trim(txtStart.Text)
    send = Trim(txtEnd.Text)
    
    If sStart <> "" And send <> "" Then
        For lRow = sStart To send
            lsPID = Trim(GetText(vasID, lRow, 5))
            lsReceNo1 = Trim(GetText(vasID, lRow, 11))
            lsReceNo2 = Trim(GetText(vasID, lRow, 12))
            SQL = "Delete from pat_res " & vbCrLf & _
                  "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                  "  and equipno = '" & gEquip & "' " & vbCrLf & _
                  "  and pid = '" & lsPID & "' " & vbCrLf & _
                  "  and receno = '" & lsReceNo1 & "' " & vbCrLf & _
                  "  and receno1 = '" & lsReceNo2 & "' "
            res = SendQuery(gLocal, SQL)
            
            DeleteRow vasID, lRow, lRow
        Next lRow
    Else
        lRow = 1
        Do While lRow <= vasID.DataRowCnt
            vasID.Row = lRow
            vasID.Col = 1
            If vasID.Value = 1 Then
                lsPID = Trim(GetText(vasID, lRow, 5))
                lsReceNo1 = Trim(GetText(vasID, lRow, 11))
                lsReceNo2 = Trim(GetText(vasID, lRow, 12))
                SQL = "Delete from pat_res " & vbCrLf & _
                      "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                      "  and equipno = '" & gEquip & "' " & vbCrLf & _
                      "  and pid = '" & lsPID & "' " & vbCrLf & _
                      "  and receno = '" & lsReceNo1 & "' " & vbCrLf & _
                      "  and receno1 = '" & lsReceNo2 & "' "
                res = SendQuery(gLocal, SQL)
                
                DeleteRow vasID, lRow, lRow
            Else
                lRow = lRow + 1
            End If
        Loop
    End If
    
    MsgBox "삭제 완료"
    chkAll.Value = 0
End Sub

Private Sub cmdDown_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow + 1
    vasActiveCell vasID, lRow + 1, 2
    vasID_Click 2, lRow + 1
End Sub

Private Sub cmdOrder_Click()

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
    
Private Sub cmdPrint_Click()
    Dim i As Integer
    Dim iRow As Integer
    
    ClearSpread vasPrint
    
    iRow = 0
    
    For i = 1 To vasID.DataRowCnt
        If GetText(vasID, i, 1) = "1" Then
            iRow = iRow + 1
            If vasPrint.MaxRows <= iRow Then
                vasPrint.MaxRows = iRow + 1
            End If
            SetText vasPrint, Trim(GetText(vasID, i, colBarCode)), iRow, 3
            
            SQL = "select result" & _
                  " from pat_res " & _
                  "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                  "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
                  "  and sendflag in ('B','C') " & vbCrLf & _
                  "  AND BARCODE = '" & Trim(GetText(vasID, i, colBarCode)) & "'" & vbCrLf & _
                  "group by result"
            res = db_select_Col(gLocal, SQL)
            SetText vasPrint, Trim(gReadBuf(0)), iRow, 9
        End If
    Next i
    
    For i = 1 To vasPrint.DataRowCnt
        Call Online_XML(gXml_S03, GetText(vasPrint, i, 3))
        SetText vasPrint, gPat_Info_Select.ACPTNO_1, i, 1
        SetText vasPrint, Mid(gPat_Info_Select.ACPT_DTE, 1, 10), i, 2
        
        Rem SetText vasPrint, Mid(gPat_Info_Select.ACPT_DTETM, 1, 10), i, 2
        
        SetText vasPrint, gPat_Info_Select.PT_NO, i, 4
        SetText vasPrint, gPat_Info_Select.PT_NM, i, 5
        SetText vasPrint, gPat_Info_Select.MEDDEPT, i, 6
        
        '검체코드 로 검체명 불러오기
        SQL = "SELECT SPCNAME"
        SQL = SQL & vbCrLf & "  FROM SPCCONFIG"
        SQL = SQL & vbCrLf & " WHERE SPCCODE = '" & gPat_Info_Select.SPC_CD_1 & "'"
        res = db_select_Col(gLocal, SQL)
        
        If res < 1 Then
            SetText vasPrint, gPat_Info_Select.SPC_CD_1, i, 7
        Else
            SetText vasPrint, gReadBuf(0), i, 7
        End If
        
        If Trim(GetText(vasPrint, i, 7)) = "1SWP" Then
            SetText vasPrint, "SWP", i, 7
        End If
        
        SetText vasPrint, gPat_Info_Select.TST_NM, i, 8
    Next i
    
    Dim sTitle As String
    Dim sHead As String
    Dim sFoot As String
    
    If vasPrint.DataRowCnt > 0 Then
        'vasPrint.PrintSmartPrint
        sTitle = "Cobas4800 RESULT"
    
        sHead = "/fn""굴림체"" /fz""15"" /fb1 /fi0 /fu0 " & "/l" & "                          " & "" & sTitle & "" & "/n/n/n " & _
                "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & "  " & "/fn""굴림체"" /fz""11"" /fb1 /fi0 /fu0 /r" & "검사일자 : " & Text_Today & "   " & "/n" '& "/n/n"
        
        sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & "  " & Text_Today & "/fn""굴림체"" /fz""11"" /fb1 /fi0 /fu0 /r" & "        국립암센터   "
        
        vasPrint.PrintHeader = sHead
        vasPrint.PrintFooter = sFoot
        vasPrint.BorderStyle = BorderStyleFixedSingle
        vasPrint.PrintBorder = True
        vasPrint.PrintGrid = True
        
        vasPrint.PrintMarginTop = 1000
        vasPrint.PrintMarginLeft = 300
        'vasPrint.PrintSmartPrint = True
        vasPrint.Action = ActionPrint
        MsgBox "출력완료"
    Else
        MsgBox "출력할 데이터가 없습니다." & vbCrLf & "조회 후 출력버튼을 눌러 주세요."
    End If
    
End Sub

Private Sub cmdQC_Click()
    'frmQCResSch.Show
End Sub

Private Sub cmdResCall_Click()
    'frmResult.Show 0
End Sub

Private Sub cmdReset_Click()
    Dim i As Integer

    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
    txteGFRCmnt.Text = ""
    
    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    ClearSpread vasID
    ClearSpread vasRes
    
    ClearSpread vasPrint
    Text_Today = Format(CDate(Date), "yyyy/mm/dd")
    
    StatusBar1.Panels(1).Text = ""
    StatusBar1.Panels(2).Text = ""
    
    gRow = 0
    
End Sub

Private Sub cmdResProc_Click()
    Dim strFileName As String
    Dim strFilePath As String
    Dim i As Integer
    Dim j As Integer
    Dim iRow As Integer
    Dim strResult As String
    Dim strBarcode As String
    Dim strPos As String
    Dim strType16Res As String
    Dim strType18Res As String
    Dim strTypeOtherRes As String
    Dim strType16ct As String
    Dim strType18ct As String
    Dim strTypeOtherct As String
    Dim strEquipCode As String
    Dim lResRow As String
    Dim strExamCode As String
    Dim strExamName As String
    Dim strSeqNo As String
    Dim strEquipRes As String
    Dim strEquipResult As String
    
    Dim strTestWay As String
    Dim strTestIdName As String
    
    'Call Proc_Order_New
     
       strTestWay = "- 검사방법 : " & gSetup.gTestWay
    strTestIdName = "- 보 고 자 : " & gSetup.gTestIdName
    
    cdResProc.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
    cdResProc.ShowOpen
    
    strFileName = cdResProc.FileName
    
    ClearSpread vasXML
    
    Cobas4800_Xml strFileName
    
    For i = 1 To vasXML.DataRowCnt
        strResult = ""
        
        strBarcode = Trim(GetText(vasXML, i, 1))
        strPos = Trim(GetText(vasXML, i, 2))
        strType16ct = Trim(GetText(vasXML, i, 3))
        strType16Res = Trim(GetText(vasXML, i, 4))
        strType18ct = Trim(GetText(vasXML, i, 5))
        strType18Res = Trim(GetText(vasXML, i, 6))
        strTypeOtherct = Trim(GetText(vasXML, i, 7))
        strTypeOtherRes = Trim(GetText(vasXML, i, 8))
        strEquipResult = Trim(GetText(vasXML, i, 9))
        
        If UCase(strType16Res) = "POSITIVE" Then
            strResult = "Positive :" & vbCrLf & "High risk type 16 (Ct : " & strType16ct & ")" & vbCrLf & strTestWay & vbCrLf & strTestIdName
        End If
                
        If UCase(strType18Res) = "POSITIVE" Then
            If strResult = "" Then
                strResult = "Positive :" & vbCrLf & "High risk type 18 (Ct : " & strType18ct & ")" & vbCrLf & strTestWay & vbCrLf & strTestIdName
            Else
                strResult = strResult & vbCrLf & "High risk type 18 (Ct : " & strType18ct & ")" & vbCrLf & strTestWay & vbCrLf & strTestIdName
            End If
            
        End If
        
        If UCase(strTypeOtherRes) = "POSITIVE" Then
            If strResult = "" Then
                strResult = "Positive :" & vbCrLf & "Other high risk (Ct : " & strTypeOtherct & ")" & vbCrLf & strTestWay & vbCrLf & strTestIdName
            Else
                strResult = strResult & vbCrLf & "Other high risk (Ct : " & strTypeOtherct & ")" & vbCrLf & strTestWay & vbCrLf & strTestIdName
            End If

        End If
         
        If strResult = "" Then
            If strEquipResult = "NEG Other HR HPV; NEG HPV16; NEG HPV18" Or strEquipResult = "Valid" Then
                strResult = "Negative" & vbCrLf & strTestWay & vbCrLf & strTestIdName
            Else
                strResult = "Invalid" & vbCrLf & strTestWay & vbCrLf & strTestIdName
            End If
        End If
        
        iRow = -1
        
        For j = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, j, colBarCode)) = strBarcode Then
                iRow = j
                Exit For
            End If
        Next
        
        If iRow = -1 Then
            iRow = vasID.DataRowCnt + 1
            If vasID.MaxRows < iRow Then
                vasID.MaxRows = iRow
            End If
            
            SetText vasID, strBarcode, iRow, colBarCode
            
        End If
        
        SetText vasID, strPos, iRow, colPos
        
        vasID.SetText colState, iRow, "Result"
        
        If Trim(GetText(vasID, iRow, colPName)) = "" Then
            Get_Sample_Info iRow
        End If
        
        
        strEquipCode = "HPV"
        
        ClearSpread vasRes
        
        
        lResRow = vasRes.DataRowCnt + 1
        If vasRes.MaxRows < lResRow Then vasRes.MaxRows = lResRow
        
        vasRes.SetText colEquipCode, lResRow, strEquipCode
        
        
        SQL = "select examcode, examname, seqno, resprec  from equipexam where equipcode = '" & strEquipCode & "'"
        res = db_select_Col(gLocal, SQL)
        
        If res > 0 Then
        
            strExamCode = Trim(gReadBuf(0))
            strExamName = Trim(gReadBuf(1))
            strSeqNo = Trim(gReadBuf(2))
            strEquipRes = strResult
            
            
            vasRes.SetText colExamCode, lResRow, Trim(strExamCode)
            vasRes.SetText colExamName, lResRow, Trim(strExamName)
            vasRes.SetText colSeqNo, lResRow, Trim(strSeqNo)
            
            vasRes.SetText colResult, lResRow, strEquipRes
            vasRes.SetText colResult1, lResRow, strEquipRes
            
            Save_Local_One_1 iRow, lResRow, "B"

        End If

        If MnTransAuto.Checked = True Then
            res = Insert_Data(iRow)
            
            If res = -1 Then
                SetForeColor vasID, iRow, iRow, 1, colState, 255, 0, 0
                SetText vasID, "Failed", iRow, colState
            Else
               
                SetBackColor vasID, iRow, iRow, 1, colState, 202, 255, 112
                SetText vasID, "Trans", iRow, colState
                
                SQL = " Update pat_res Set " & vbCrLf & _
                      " sendflag = 'C' " & vbCrLf & _
                      " Where equipno = '" & gEquip & "' " & vbCrLf & _
                      " And barcode = '" & Trim(GetText(vasID, iRow, colBarCode)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
        Else
'''            SetBackColor vasID, iRow, iRow, 1, vasID.MaxCols, 202, 255, 112
            SetText vasID, "Result", iRow, colState
        End If
    
    
    Next
    
End Sub


Private Sub Proc_Auto_res(argData As String)
    Dim strFileName As String
    Dim strFilePath As String
    Dim i As Integer
    Dim j As Integer
    Dim iRow As Integer
    Dim strResult As String
    Dim strBarcode As String
    Dim strPos As String
    Dim strType16Res As String
    Dim strType18Res As String
    Dim strTypeOtherRes As String
    Dim strType16ct As String
    Dim strType18ct As String
    Dim strTypeOtherct As String
    Dim strEquipCode As String
    Dim lResRow As String
    Dim strExamCode As String
    Dim strExamName As String
    Dim strSeqNo As String
    Dim strEquipRes As String
    Dim strEquipResult As String
    Dim strAData() As String
    Dim strSData() As String
    
  
'''    strFileName = App.Path & "\ResultXML\" & argFileName & ".xml"
    
    ClearSpread vasXML
    
    strAData = Split(argData, chrCR)
    For i = 1 To UBound(strAData)
        strSData = Split(strAData(i - 1), ",")
        For j = 1 To UBound(strSData)
            SetText vasXML, strSData(j - 1), i, j
        Next
    Next
'''    Cobas4800_Xml strFileName
    
    For i = 1 To vasXML.DataRowCnt
        strResult = ""
        
        strBarcode = Trim(GetText(vasXML, i, 1))
        'If strBarcode = "18082902160" Then
        '    Stop
        'End If
        strPos = Trim(GetText(vasXML, i, 2))
        strType16ct = Trim(GetText(vasXML, i, 3))
        strType16Res = Trim(GetText(vasXML, i, 4))
        strType18ct = Trim(GetText(vasXML, i, 5))
        strType18Res = Trim(GetText(vasXML, i, 6))
        strTypeOtherct = Trim(GetText(vasXML, i, 7))
        strTypeOtherRes = Trim(GetText(vasXML, i, 8))
        strEquipResult = Trim(GetText(vasXML, i, 9))
        
        If UCase(strType16Res) = "POSITIVE" Then
            strResult = "Positive :" & vbCrLf & "High risk type 16 (Ct : " & strType16ct & ")"
        End If
                
        If UCase(strType18Res) = "POSITIVE" Then
            If strResult = "" Then
                strResult = "Positive :" & vbCrLf & "High risk type 18 (Ct : " & strType18ct & ")"
            Else
                strResult = strResult & vbCrLf & "High risk type 18 (Ct : " & strType18ct & ")"
            End If
            
        End If
        
        If UCase(strTypeOtherRes) = "POSITIVE" Then
            If strResult = "" Then
                strResult = "Positive :" & vbCrLf & "Other high risk (Ct : " & strTypeOtherct & ")"
            Else
                strResult = strResult & vbCrLf & "Other high risk (Ct : " & strTypeOtherct & ")"
            End If

        End If
         
        If strResult = "" Then
            If strEquipResult = "NEG Other HR HPV; NEG HPV16; NEG HPV18" Or strEquipResult = "Valid" Then
                strResult = "Negative"
            Else
                strResult = "Invalid"
            End If
        End If
        
        iRow = -1
        
        For j = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, j, colBarCode)) = strBarcode Then
                iRow = j
                Exit For
            End If
        Next
        
        If iRow = -1 Then
            iRow = vasID.DataRowCnt + 1
            If vasID.MaxRows < iRow Then
                vasID.MaxRows = iRow
            End If
            
            SetText vasID, strBarcode, iRow, colBarCode
            
        End If
        
        SetText vasID, strPos, iRow, colPos
        SetText vasID, "Result", iRow, colState
        'vasID.SetText colState, iRow, "Result"
        
        gOrderExam = ""
        
        If Trim(GetText(vasID, iRow, colPName)) = "" Then
            '테스트용
            Get_Sample_Info iRow
        End If
        
        
        strEquipCode = "HPV"
        'strEquipCode = "02HPVGEN^^FULL"
        
        ClearSpread vasRes
        
        
        lResRow = vasRes.DataRowCnt + 1
        If vasRes.MaxRows < lResRow Then vasRes.MaxRows = lResRow
        
        vasRes.SetText colEquipCode, lResRow, strEquipCode
        
        '해당검사코드 불러와서 Examcode 랑 매칭해야함 . 2014-06-30 이지성=========================================
        'gOrderExam = ""
    
        
        '테스트용
        If gOrderExam = "" Then
            Online_XML gXml_S07, strBarcode
        End If
        
        SQL = ""
        SQL = SQL & "select examcode, examname, seqno, resprec  " & vbCr
        SQL = SQL & "  from equipexam " & vbCr
        SQL = SQL & " where equipcode = '" & strEquipCode & "'" & vbCr
        '테스트용
        SQL = SQL & "   and examcode in (" & gOrderExam & ")"
        
        Save_Raw_Data "[검사코드조회]" & SQL
        res = db_select_Col(gLocal, SQL)
        
        If res > 0 Then
        
            strExamCode = Trim(gReadBuf(0))
            strExamName = Trim(gReadBuf(1))
            strSeqNo = Trim(gReadBuf(2))
            strEquipRes = strResult
            
            
            vasRes.SetText colExamCode, lResRow, Trim(strExamCode)
            vasRes.SetText colExamName, lResRow, Trim(strExamName)
            vasRes.SetText colSeq, lResRow, Trim(strSeqNo)
            
            vasRes.SetText colResult, lResRow, strEquipRes
            vasRes.SetText colResult1, lResRow, strEquipRes
            
            Save_Local_One_1 iRow, lResRow, "B"

        End If

        vasID.RowHeight(-1) = 13
        
        If MnTransAuto.Checked = True Then
            res = Insert_Data(iRow)
            
            If res = -1 Then
                SetForeColor vasID, iRow, iRow, 1, colState, 255, 0, 0
                SetText vasID, "Failed", iRow, colState
            Else
               
                SetBackColor vasID, iRow, iRow, 1, colState, 202, 255, 112
                SetText vasID, "Trans", iRow, colState
                
                SQL = " Update pat_res Set " & vbCrLf & _
                      " sendflag = 'C' " & vbCrLf & _
                      " Where equipno = '" & gEquip & "' " & vbCrLf & _
                      " And barcode = '" & Trim(GetText(vasID, iRow, colBarCode)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
        Else
'''            SetBackColor vasID, iRow, iRow, 1, vasID.MaxCols, 202, 255, 112
            SetText vasID, "Result", iRow, colState
        End If
    
    
    Next
    
End Sub


Function getMutationDetected() As String

    Dim TextLine
    Dim strBuffer
    
    Open App.Path & "\eGFR_MD코멘트.txt" For Input As #1 ' 파일을 엽니다.
    
    Do While Not EOF(1) ' 파일의 끝을 만날 때까지 반복합니다.
        Line Input #1, TextLine ' 변수로 데이터 행을 읽어들입니다.
        strBuffer = strBuffer & TextLine & vbNewLine
    Loop
    
    Close #1 ' 파일을 닫습니다

    getMutationDetected = strBuffer


End Function

Function getMutationNotDetected()

    Dim TextLine
    Dim strBuffer

    Open App.Path & "\eGFR_ND코멘트.txt" For Input As #1 ' 파일을 엽니다.
    
    Do While Not EOF(1) ' 파일의 끝을 만날 때까지 반복합니다.
        Line Input #1, TextLine ' 변수로 데이터 행을 읽어들입니다.
        strBuffer = strBuffer & TextLine & vbNewLine
    Loop
    
    Close #1 ' 파일을 닫습니다

    getMutationNotDetected = strBuffer

End Function

Function getMutationInvalid()

    Dim TextLine
    Dim strBuffer

    Open App.Path & "\eGFR_IV코멘트.txt" For Input As #1 ' 파일을 엽니다.
    
    Do While Not EOF(1) ' 파일의 끝을 만날 때까지 반복합니다.
        Line Input #1, TextLine ' 변수로 데이터 행을 읽어들입니다.
        strBuffer = strBuffer & TextLine & vbNewLine
    Loop
    
    Close #1 ' 파일을 닫습니다

    getMutationInvalid = strBuffer

End Function


Private Sub Proc_Auto_res_eGFR(argData As String)
    Dim strFileName As String
    Dim strFilePath As String
    Dim i As Integer
    Dim j As Integer
    Dim iRow As Integer
    Dim strResult As String
    Dim strBarcode As String
    Dim strPos As String
    
    Dim strType16Res As String
    Dim strType18Res As String
    Dim strTypeOtherRes As String
    Dim strType16ct As String
    Dim strType18ct As String
    Dim strTypeOtherct As String
    
    Dim strTypeRes1 As String
    Dim strTypeLISRes1 As String
    Dim strTypeRes2 As String
    Dim strTypeLISRes2 As String
    Dim strTypeRes3 As String
    Dim strTypeLISRes3 As String
    
    
    Dim strEquipCode As String
    Dim lResRow As String
    Dim strExamCode As String
    Dim strExamName As String
    Dim strSeqNo As String
    Dim strEquipRes As String
    Dim strEquipResult As String
    Dim strAData() As String
    Dim strSData() As String
    
  
'''    strFileName = App.Path & "\ResultXML\" & argFileName & ".xml"
    
    ClearSpread vasXML
    
    strAData = Split(argData, chrCR)
    For i = 1 To UBound(strAData)
        strSData = Split(strAData(i - 1), ",")
        For j = 1 To UBound(strSData)
            SetText vasXML, strSData(j - 1), i, j
        Next
    Next
'''    Cobas4800_Xml strFileName
    
    '18072303081,C03,,,,,,,Mutation Detected,faba5872-93bc-11e8-aa01-0023242e6652,
    '18072404048,E03,,,,,,,No Mutation Detected,faba589a-93bc-11e8-aa01-0023242e6652,
    
    For i = 1 To vasXML.DataRowCnt
        strResult = ""
        
        strBarcode = Trim(GetText(vasXML, i, 1))
        strPos = Trim(GetText(vasXML, i, 2))
        strTypeRes1 = Trim(GetText(vasXML, i, 3))
        strTypeLISRes1 = Trim(GetText(vasXML, i, 4))
        strTypeRes2 = Trim(GetText(vasXML, i, 5))
        strTypeLISRes2 = Trim(GetText(vasXML, i, 6))
        strTypeRes3 = Trim(GetText(vasXML, i, 7))
        strTypeLISRes3 = Trim(GetText(vasXML, i, 8))
        strEquipResult = Trim(GetText(vasXML, i, 9))
        
'        If UCase(strType16Res) = "POSITIVE" Then
'            strResult = "Positive :" & vbCrLf & "High risk type 16 (Ct : " & strType16ct & ")"
'        End If
'
'        If UCase(strType18Res) = "POSITIVE" Then
'            If strResult = "" Then
'                strResult = "Positive :" & vbCrLf & "High risk type 18 (Ct : " & strType18ct & ")"
'            Else
'                strResult = strResult & vbCrLf & "High risk type 18 (Ct : " & strType18ct & ")"
'            End If
'
'        End If
'
'        If UCase(strTypeOtherRes) = "POSITIVE" Then
'            If strResult = "" Then
'                strResult = "Positive :" & vbCrLf & "Other high risk (Ct : " & strTypeOtherct & ")"
'            Else
'                strResult = strResult & vbCrLf & "Other high risk (Ct : " & strTypeOtherct & ")"
'            End If
'
'        End If
'
'        If strResult = "" Then
'            If strEquipResult = "NEG Other HR HPV; NEG HPV16; NEG HPV18" Or strEquipResult = "Valid" Then
'                strResult = "Negative"
'            Else
'                strResult = "Invalid"
'            End If
'        End If
        
        If strEquipResult = "Mutation Detected" Then
            'strResult = getMutationDetected
            strResult = strEquipResult & vbNewLine & strTypeRes3
        Else
            If strTypeRes1 = "Mutation Detected" Then
                strResult = strTypeRes1 & vbNewLine & strTypeRes3
            
            '2018.12.04 추가
            ElseIf strTypeRes1 = "Invalid" Then
                strResult = strTypeRes1 & vbNewLine & strTypeRes3
            Else
                'strResult = getMutationNotDetected
                strResult = "Not detected" & vbNewLine & strTypeRes3
            End If
        End If
    
        
        iRow = -1
        
        For j = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, j, colBarCode)) = strBarcode Then
                iRow = j
                Exit For
            End If
        Next
        
        If iRow = -1 Then
            iRow = vasID.DataRowCnt + 1
            If vasID.MaxRows < iRow Then
                vasID.MaxRows = iRow
            End If
            
            SetText vasID, strBarcode, iRow, colBarCode
            
        End If
        
        SetText vasID, strPos, iRow, colPos
        SetText vasID, "Result", iRow, colState
        'vasID.SetText colState, iRow, "Result"
        
        gOrderExam = ""
        
        If Trim(GetText(vasID, iRow, colPName)) = "" Then
            '테스트용
            Get_Sample_Info iRow
        End If
        
        
        strEquipCode = "EGFR"
        'strEquipCode = "08EGFR^^AnD"
        
        ClearSpread vasRes
        
        
        lResRow = vasRes.DataRowCnt + 1
        If vasRes.MaxRows < lResRow Then vasRes.MaxRows = lResRow
        
        vasRes.SetText colEquipCode, lResRow, strEquipCode
        
        '해당검사코드 불러와서 Examcode 랑 매칭해야함 . 2014-06-30 이지성=========================================
    
        If gOrderExam = "" Then
            '테스트용
            Online_XML gXml_S07, strBarcode
        End If
        
        SQL = ""
        SQL = SQL & "select examcode, examname, seqno, resprec  " & vbCr
        SQL = SQL & "  from equipexam " & vbCr
        SQL = SQL & " where equipcode = '" & strEquipCode & "'" & vbCr
        
        '테스트용
        SQL = SQL & "   and examcode in (" & gOrderExam & ")"
        res = db_select_Col(gLocal, SQL)
        
        If res > 0 Then
        
            strExamCode = Trim(gReadBuf(0))
            strExamName = Trim(gReadBuf(1))
            strSeqNo = Trim(gReadBuf(2))
            strEquipRes = strResult
            
            
            vasRes.SetText colExamCode, lResRow, Trim(strExamCode)
            vasRes.SetText colExamName, lResRow, Trim(strExamName)
            vasRes.SetText colSeq, lResRow, Trim(strSeqNo)
            
            vasRes.SetText colResult, lResRow, strEquipRes
            vasRes.SetText colResult1, lResRow, strEquipRes
            'vasRes.RowHeight(-1) = 350
            Save_Local_One_1 iRow, lResRow, "B"

        End If

        If MnTransAuto.Checked = True Then
            res = Insert_Data(iRow)
            
            If res = -1 Then
                SetForeColor vasID, iRow, iRow, 1, colState, 255, 0, 0
                SetText vasID, "Failed", iRow, colState
            Else
               
                SetBackColor vasID, iRow, iRow, 1, colState, 202, 255, 112
                SetText vasID, "Trans", iRow, colState
                
                SQL = " Update pat_res Set " & vbCrLf & _
                      " sendflag = 'C' " & vbCrLf & _
                      " Where equipno = '" & gEquip & "' " & vbCrLf & _
                      " And barcode = '" & Trim(GetText(vasID, iRow, colBarCode)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
        Else
'''            SetBackColor vasID, iRow, iRow, 1, vasID.MaxCols, 202, 255, 112
            SetText vasID, "Result", iRow, colState
        End If
    
    
    Next
    
End Sub

Private Sub Proc_Res(argData As String)
    
End Sub

Private Sub cmdResSave_Click()
    'Proc_Result txtBarcode
End Sub

Private Sub cmdSave_Click()
    Dim i       As Integer
    Dim strList As String
    
    Call WritePrivateProfileString("config", "gTestWay", txtTestWay.Text, App.Path & "\Interface.ini")
    Call WritePrivateProfileString("config", "gTestIdName", txtTestIdName.Text, App.Path & "\Interface.ini")
    
    strList = ""
    For i = 1 To cboTestIdList.ListCount - 1
        If txtTestIdName.Text <> cboTestIdList.List(i) Then
            strList = strList & cboTestIdList.List(i) & ","
        End If
    Next
    If strList <> "" Then
        strList = strList & txtTestIdName.Text
    End If
    
    Call WritePrivateProfileString("config", "gTestIdList", strList, App.Path & "\Interface.ini")
    
End Sub

Private Sub cmdSend_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
            res = Insert_Data(lRow)
        
            If res = -1 Then
                SetForeColor vasID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID, "실패", lRow, colState
            Else
                vasID.Row = lRow
                vasID.Col = 1
                vasID.Value = 1
                
                SetBackColor vasID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID, "완료", lRow, colState
                
                SQL = " Update pat_res Set " & vbCrLf & _
                      " sendflag = 'C' " & vbCrLf & _
                      " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                      " And equipno = '" & gEquip & "' " & vbCrLf & _
                      " And barcode = '" & Trim(GetText(vasID, lRow, colBarCode)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasID.Row = lRow
            vasID.Col = 1
            vasID.Value = 0
        End If
    Next lRow
End Sub

Private Sub cmdUp_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow - 1
    vasActiveCell vasID, lRow - 1, 2
    vasID_Click 2, lRow - 1
End Sub

Private Sub Command_close_Click()
    Unload Me
End Sub

Private Sub Command_config_Click()
    frmConfig.Show 1
End Sub


Private Sub Command_setup_Click()
    frmOrderCode.Show 1
    GetExamCode
End Sub


Private Sub Command13_Click()
    Dim i As Integer
    
    SQL = "select item_code, item_name, m_stype_code, disp_seq, m_item_code from tbl_item"
    res = db_select_Vas(gLocal_1, SQL, vaSpread1)
    
    SQL = "delete from equipexam"
    res = SendQuery(gLocal, SQL)
    
    For i = 1 To vaSpread1.DataRowCnt
    
        SQL = "insert into equipexam(equipno, examcode, equipcode, examname, examtype, seqno, resprec, examflag) " & vbCrLf & _
              "values('C064','" & Trim(GetText(vaSpread1, i, 1)) & "','" & Trim(GetText(vaSpread1, i, 5)) & "','" & Trim(GetText(vaSpread1, i, 2)) & "','" & Trim(GetText(vaSpread1, i, 3)) & "','" & Trim(GetText(vaSpread1, i, 4)) & "','1','1')"
        res = SendQuery(gLocal, SQL)
    Next
    
End Sub

Private Sub Command14_Click()
'    frmUserChange.Show 0
    
End Sub

Private Sub Command15_Click()
    Dim ss As String
    Dim i As Integer
    
    
    For i = 1 To Len(Text5.Text)
        ss = Mid(Text5.Text, i, 1)
        
        Select Case ss
        Case Chr(5)
            Save_Raw_Data "[Rx]" & ss
            
            If comSend = "stENQ" Then
                If gENQCnt <= 3 Then
                    gENQCnt = gENQCnt + 1
                    comSend = "stENQ"
                    gPreMsg = Chr(5)
                    Save_Raw_Data "[Tx]" & gPreMsg
                    MSComm1.Output = Chr(5)
                    Exit Sub
                End If
            End If
            
            gNACKCnt = 0
            
            comState = "stRx"
            comsignal = "stGen"
            
            txtBuff = ""
            
            gPreMsg = Chr(6)
            Save_Raw_Data "[Tx]" & gPreMsg
            MSComm1.Output = Chr(6)
        
        Case Chr(6)
            Save_Raw_Data "[Rx]" & ss
            If comSend = "stENQ" Then
                SendOrder
            ElseIf comSend = "stOrder" Then
                comSend = "stEOT"
                
                gPreMsg = Chr(4)
                Save_Raw_Data "[Tx]" & gPreMsg
                MSComm1.Output = Chr(4)
            ElseIf comSend = "stEOT" Then
                comState = "stRX"
            End If
            
        Case chrNACK
        
            Save_Raw_Data "[Rx]" & ss
            
            gNACKCnt = gNACKCnt + 1
            If gNACKCnt < 3 Then
                Save_Raw_Data "[Tx]" & gPreMsg
                MSComm1.Output = gPreMsg
            Else
                gNACKCnt = 0
                gPreMsg = Chr(4)
                Save_Raw_Data "[Tx]" & gPreMsg
                MSComm1.Output = Chr(4)
            End If
            
        Case chrEOT
        
            Save_Raw_Data "[Rx]" & txtBuff
            
            gENQCnt = 0
            
            Modular txtBuff
    
            If Trim(gOrderMessage) <> "" Then
                comSend = "stENQ"
                gPreMsg = Chr(5)
                Save_Raw_Data "[Tx]" & gPreMsg
                
                MSComm1.Output = Chr(5)
                
                gENQCnt = gENQCnt + 1
            End If
            
        Case chrSTX
            comsignal = "stSTX"
        Case chrETB
            comsignal = "stETB"
        Case chrETX
            comsignal = "stETX"
        Case chrCR
            If comsignal = "stETB" Then
            
            ElseIf comsignal = "stGen" Then
                txtBuff = txtBuff & ss
            End If
        Case chrLF
            If comsignal = "stETB" Then
            
            ElseIf comsignal = "stETX" Then
            
            ElseIf comsignal = "stGen" Then
                txtBuff = txtBuff & ss
            End If
            
            gPreMsg = Chr(6)
            Save_Raw_Data "[Tx]" & gPreMsg
            MSComm1.Output = Chr(6)
        Case Else
            If comsignal = "stSTX" Then
                comsignal = "stGen"
                Exit Sub
            ElseIf comsignal = "stETB" Then
                Exit Sub
            ElseIf comsignal = "stETX" Then
                Exit Sub
            End If
            
            txtBuff = txtBuff & ss
        End Select
    
    Next
    
    
End Sub

Private Sub Command16_Click()

    If Text_Today <> Format(Date, "yyyy/mm/dd") Then
        Text_Today = Format(Date, "yyyy/mm/dd")
        cmdReset_Click
    End If

End Sub

Private Sub Command17_Click()
    XE2100_ASTM txtDataTmp
    
    txtDataTmp = ""
End Sub

Private Sub Command18_Click()
    Online_XML gXml_S07, Text4
End Sub

Private Sub Command19_Click()
    Online_XML gXml_S03, Text4
End Sub

Private Sub Command20_Click()
        Dim ss As String
    Dim i As Long
    
    For i = 1 To Len(txtXML)
    
    ss = Mid(txtXML, i, 1)
    
    Select Case ss
    Case Chr(5)
        Save_Raw_Data "[Rx]" & ss
        
        If comSend = "stENQ" Then
            If gENQCnt <= 3 Then
                gENQCnt = gENQCnt + 1
                comSend = "stENQ"
                gPreMsg = Chr(5)
                Save_Raw_Data "[Tx]" & gPreMsg
                'MSComm1.Output = Chr(5)
'''                Exit Sub
            End If
        End If
        
        gNACKCnt = 0
        
        comState = "stRx"
        comsignal = "stGen"
        
        txtBuff = ""
        
        gPreMsg = Chr(6)
        Save_Raw_Data "[Tx]" & gPreMsg
        'MSComm1.Output = Chr(6)
    
    Case Chr(6)
        Save_Raw_Data "[Rx]" & ss
        If comSend = "stENQ" Then
            SendOrder
        ElseIf comSend = "stOrder" Then
            comSend = "stEOT"
            
            gPreMsg = Chr(4)
            Save_Raw_Data "[Tx]" & gPreMsg
            'MSComm1.Output = Chr(4)
        ElseIf comSend = "stEOT" Then
            comState = "stRX"
        End If
        
    Case chrNACK
    
        Save_Raw_Data "[Rx]" & ss
        
        gNACKCnt = gNACKCnt + 1
        If gNACKCnt < 3 Then
            Save_Raw_Data "[Tx]" & gPreMsg
            'MSComm1.Output = gPreMsg
        Else
            gNACKCnt = 0
            gPreMsg = Chr(4)
            Save_Raw_Data "[Tx]" & gPreMsg
            'MSComm1.Output = Chr(4)
        End If
        
    Case chrEOT
    
        Save_Raw_Data "[Rx]" & txtBuff
        
        gENQCnt = 0
        
        Modular txtBuff

        If Trim(gOrderMessage) <> "" Then
            comSend = "stENQ"
            gPreMsg = Chr(5)
            Save_Raw_Data "[Tx]" & gPreMsg
            
            'MSComm1.Output = Chr(5)
            
            gENQCnt = gENQCnt + 1
        End If
        
    Case chrSTX
        comsignal = "stSTX"
    Case chrETB
        comsignal = "stETB"
    Case chrETX
        comsignal = "stETX"
    Case chrCR
        If comsignal = "stETB" Then
        
        ElseIf comsignal = "stGen" Then
            txtBuff = txtBuff & ss
        End If
    Case chrLF
        If comsignal = "stETB" Then
        
        ElseIf comsignal = "stETX" Then
        
        ElseIf comsignal = "stGen" Then
            txtBuff = txtBuff & ss
        End If
        
        gPreMsg = Chr(6)
        Save_Raw_Data "[Tx]" & gPreMsg
        'MSComm1.Output = Chr(6)
    Case Else
        If comsignal = "stSTX" Then
            comsignal = "stGen"
''            Exit Sub
        ElseIf comsignal = "stETB" Then
'''            Exit Sub
        ElseIf comsignal = "stETX" Then
'''            Exit Sub
        End If
        
        txtBuff = txtBuff & ss
    End Select
    Next
End Sub

Private Sub Command21_Click()
    Dim sTmp As String
    Dim i As Integer
    Dim strResFlag As String
    
    
    sTmp = Text6.Text
    
    
    txtBuff.Text = txtBuff.Text & sTmp
    
    Save_Raw_Data "[RX " & Format(Time, "hh:mm:ss") & "]" & sTmp
    
    If InStr(1, sTmp, chrENQ) > 0 Then
        Save_Raw_Data "[RX " & Format(Time, "hh:mm:ss") & "]" & chrACK
'''        Winsock1.SendData chrACK
        
    End If
    
    If InStr(1, sTmp, chrLF) > 0 Then
        Save_Raw_Data "[RX " & Format(Time, "hh:mm:ss") & "]" & chrACK
'''        Winsock1.SendData chrACK
    End If
    
    If InStr(1, sTmp, chrEOT) > 0 Then
        strResFlag = Cobas4800All(txtBuff.Text)
        
    End If
    If InStr(1, sTmp, chrACK) > 0 Then
        If Trim(GetText(vasASTM, 1, 1)) <> "" Then
            Save_Raw_Data "[TX]" & Trim(GetText(vasASTM, 1, 1))
'''            Winsock1.SendData Trim(GetText(vasASTM, 1, 1))
            DeleteRow vasASTM, 1, 1
            
        End If
    End If
    
End Sub

Private Sub Command3_Click()
    SQL = "CREATE INDEX resindex1 ON pat_res (examdate,equipno,barcode,equipcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex1 created"
    Else
        MsgBox "resindex1 failed"
    End If
    SQL = "CREATE INDEX resindex2 ON pat_res (examdate,equipno,barcode,examcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex2 created"
    Else
        MsgBox "resindex2 failed"
    End If
    
    SQL = "CREATE INDEX resindex3 ON pat_res (barcode,examcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex3 created"
    Else
        MsgBox "resindex3 failed"
    End If
    
    SQL = "CREATE INDEX resindex4 ON pat_res (barcode,equipcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex4 created"
    Else
        MsgBox "resindex4 failed"
    End If
End Sub


Private Sub Form_Load()
    Dim sDate   As String
    Dim varTmp  As Variant
    Dim i       As Integer
    
    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
    
    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    
    cmdReset_Click
    
    GetSetup
    
    varTmp = Split(gSetup.gTestIdList, ",")
    
    cboTestIdList.Clear
    cboTestIdList.AddItem " == 선택 == "
    For i = 0 To UBound(varTmp)
        cboTestIdList.AddItem varTmp(i)
    Next
    cboTestIdList.ListIndex = 0
    
    txtTestWay.Text = gSetup.gTestWay
    txtTestIdName.Text = gSetup.gTestIdName
    
    
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
        
    Text_Today = Format(Date, "yyyy/mm/dd")

    GetExamCode
        
    sDate = Format(DateAdd("y", CDate(Text_Today.Text), -30), "yyyymmdd")
    SQL = "delete from pat_res where examdate < '" & sDate & "'"
    res = SendQuery(gLocal, SQL)
    
    '******************************************************
    
    SQL = " Select cutoff From equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table equipexam Add Column cutoff text(20) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Select cutoffflag From equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table equipexam Add Column cutoffflag long "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Select negvalue From equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table equipexam Add Column negvalue text(10) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Select posvalue From equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table equipexam Add Column posvalue text(10) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Select posequal From equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table equipexam Add Column posequal long "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Select negequal From equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table equipexam Add Column negequal long "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Select ordgubun From equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table equipexam Add column ordgubun text(1) "
        res = SendQuery(gLocal, SQL)
    End If
    
'    SQL = " Alter Table pat_res modify column posno text(20) "
'    res = SendQuery(gLocal, SQL)
    
    lblUser.Caption = gIFUser
    
    WinSock_Listen Winsock1
    WinSock_Listen2 Winsock2
'    SQL = " Alter table equipexam Alter column seqno text(3) "
'    res = SendQuery(gLocal, SQL)
    '******************************************************
End Sub

Function Get_Sample_Info(ByVal asRow As Long) As Integer
    Dim lsbarcode As String
    Dim lsPID As String
    Dim lsReceNo As String
    Dim sRes As String

    Get_Sample_Info = -1
    
    '샘플 환자 정보 가져오기
    lsbarcode = Trim(GetText(vasID, asRow, colBarCode))   '샘플 바코드 번호
    
'    If Trim(lsbarcode) = "" Then: Exit Function
     sRes = Online_XML(gXml_S03, lsbarcode)
'    If sRes = 1 Then
        SetText vasID, gPat_Info_Select.PT_NO, asRow, colPID
        SetText vasID, gPat_Info_Select.PT_NM, asRow, colPName
        SetText vasID, gPat_Info_Select.SEX, asRow, colPSex
        SetText vasID, gPat_Info_Select.AGE, asRow, colPAge
        SetText vasID, gPat_Info_Select.SPC_CD_1, asRow, colSeqNo
        SetText vasID, Mid(gPat_Info_Select.ACPT_DTETM, 1, 10), asRow, colDate
        SetText vasID, gPat_Info_Select.ACPTNO_1, asRow, colReceno

'''        vasID.RowHeight(asRow) = 20
        
        Get_Sample_Info = 1
'    End If

End Function

Function Get_Sample_Info_Local(ByVal asRow As Long) As Integer
    Dim lsbarcode As String
    Dim lsPID As String
    Dim lsReceNo As String
    Dim sRes As String

    Get_Sample_Info_Local = -1
    
    '샘플 환자 정보 가져오기
    lsbarcode = Trim(GetText(vasID, asRow, colBarCode))   '샘플 바코드 번호
    
    SQL = " Select pid, pname, psex, page, seqno, recedate, receno From pat_res " & CR & _
          " Where equipno = '" & gEquip & "' " & CR & _
          " And examdate = '" & Format(Text_Today.Text, "YYYYMMDD") & "' "
    res = db_select_Col(gLocal, SQL)
    
    If res = 1 Then
        SetText vasID, Trim(gReadBuf(0)), asRow, colPID
        SetText vasID, Trim(gReadBuf(1)), asRow, colPName
        SetText vasID, Trim(gReadBuf(2)), asRow, colPSex
        SetText vasID, Trim(gReadBuf(3)), asRow, colPAge
        SetText vasID, Trim(gReadBuf(4)), asRow, colSeqNo
        SetText vasID, Trim(gReadBuf(5)), asRow, colDate
        SetText vasID, Trim(gReadBuf(6)), asRow, colReceno
        
'''        vasID.RowHeight(asRow) = 20
        
        Get_Sample_Info_Local = 1
    End If
End Function


Function EquipExamCode(argEquipCode As String, argPID As String, argSENO As String, argSEQN As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String

    EquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    ClearSpread vasTemp1
    sExamCode = ""
    
    SQL = " Select examcode From EquipExam " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(argEquipCode) & "' "
    res = db_select_Vas(gLocal, SQL, vasTemp1)
    
    If vasTemp1.DataRowCnt < 1 Then
        Exit Function
    End If
    
    For i = 1 To vasTemp1.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & ",'" & Trim(GetText(vasTemp1, i, 1)) & "'"
        Else
            sExamCode = "'" & Trim(GetText(vasTemp1, i, 1)) & "'"
        End If
    Next i

    SQL = " Select SUCD From LRESULT " & CR & _
          " Where PAID = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and SENO = " & argSENO & vbCrLf & _
          "   and SEQN = " & argSEQN & vbCrLf & _
          "   and SUCD in ( " & sExamCode & ")  "
          
    res = db_select_Col(gServer, SQL)
  
    If gReadBuf(0) <> "" Then
        EquipExamCode = Trim(gReadBuf(0))
    End If
    
End Function

Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    
    SQL = "Select equipcode, examcode, examname, reflow, refhigh,ordgubun " & vbCrLf & _
          "From equipexam " & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "order by  examcode "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    If res > 0 Then
        ReDim gArrEquip(1 To vasTemp.DataRowCnt, 1 To 7)
    Else
        SaveQuery SQL
        Exit Function
    End If
        
    For i = 1 To vasTemp.DataRowCnt
        gArrEquip(i, 1) = i
        For j = 1 To 6
            gArrEquip(i, j + 1) = Trim(GetText(vasTemp, i, j))
        Next j
    Next i
    
    GetExamCode = 1
End Function

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

'    Call dce_close_env      ' Server와 연결을 끊는 곳
    DisConnect_Local
    
    Unload Me
    
    End
    
End Sub
'
'Private Sub Label1_DblClick()
'    If Command21.Visible = False Then
'        Command21.Visible = True
'        Text6.Visible = True
'    Else
'        Command21.Visible = False
'        Text6.Visible = False
'    End If
'End Sub

Private Sub MnComment_Click()
    frmRemark.Show 1
End Sub

Private Sub MnExamConfig_Click()
    frmOrderCode.Show 1
    GetExamCode
End Sub

Private Sub MnExit_Click()
    Unload Me
End Sub

Private Sub MnTConfig_Click()
    frmConfig.Show 1
End Sub

Private Sub MnTransAuto_Click()
    chkMode.Caption = "Auto"
    MnTransAuto.Checked = True
    MnTransManual.Checked = False
    chkMode.Value = 1
    
End Sub

Private Sub MnTransManual_Click()
    chkMode.Caption = "Manual"
    MnTransAuto.Checked = False
    MnTransManual.Checked = True
    chkMode.Value = 0
End Sub

Private Sub MSComm1_OnComm()
    Dim ss As String
    
    ss = MSComm1.Input
    
    Select Case ss
    Case Chr(5)
        Save_Raw_Data "[Rx]" & ss
        
        If comSend = "stENQ" Then
            If gENQCnt <= 3 Then
                gENQCnt = gENQCnt + 1
                comSend = "stENQ"
                gPreMsg = Chr(5)
                Save_Raw_Data "[Tx]" & gPreMsg
                MSComm1.Output = Chr(5)
                Exit Sub
            End If
        End If
        
        gNACKCnt = 0
        
        comState = "stRx"
        comsignal = "stGen"
        
        txtBuff = ""
        
        gPreMsg = Chr(6)
        Save_Raw_Data "[Tx]" & gPreMsg
        MSComm1.Output = Chr(6)
    
    Case Chr(6)
        Save_Raw_Data "[Rx]" & ss
        If comSend = "stENQ" Then
            SendOrder
        ElseIf comSend = "stOrder" Then
            comSend = "stEOT"
            
            gPreMsg = Chr(4)
            Save_Raw_Data "[Tx]" & gPreMsg
            MSComm1.Output = Chr(4)
        ElseIf comSend = "stEOT" Then
            comState = "stRX"
        End If
        
    Case chrNACK
    
        Save_Raw_Data "[Rx]" & ss
        
        gNACKCnt = gNACKCnt + 1
        If gNACKCnt < 3 Then
            Save_Raw_Data "[Tx]" & gPreMsg
            MSComm1.Output = gPreMsg
        Else
            gNACKCnt = 0
            gPreMsg = Chr(4)
            Save_Raw_Data "[Tx]" & gPreMsg
            MSComm1.Output = Chr(4)
        End If
        
    Case chrEOT
    
        Save_Raw_Data "[Rx]" & txtBuff
        
        gENQCnt = 0
        
        Modular txtBuff

        If Trim(gOrderMessage) <> "" Then
            comSend = "stENQ"
            gPreMsg = Chr(5)
            Save_Raw_Data "[Tx]" & gPreMsg
            
            MSComm1.Output = Chr(5)
            
            gENQCnt = gENQCnt + 1
        End If
        
    Case chrSTX
        comsignal = "stSTX"
    Case chrETB
        comsignal = "stETB"
    Case chrETX
        comsignal = "stETX"
    Case chrCR
        If comsignal = "stETB" Then
        
        ElseIf comsignal = "stGen" Then
            txtBuff = txtBuff & ss
        End If
    Case chrLF
        If comsignal = "stETB" Then
        
        ElseIf comsignal = "stETX" Then
        
        ElseIf comsignal = "stGen" Then
            txtBuff = txtBuff & ss
        End If
        
        gPreMsg = Chr(6)
        Save_Raw_Data "[Tx]" & gPreMsg
        MSComm1.Output = Chr(6)
    Case Else
        If comsignal = "stSTX" Then
            comsignal = "stGen"
            Exit Sub
        ElseIf comsignal = "stETB" Then
            Exit Sub
        ElseIf comsignal = "stETX" Then
            Exit Sub
        End If
        
        txtBuff = txtBuff & ss
    End Select
End Sub

Sub Modular(asVar As String)
    Dim i As Integer
    Dim iIndex As Integer
    
    Dim lsData As String
    Dim lsTemp As String
    
    Dim lsHead As String
    Dim lsPatient As String
    Dim lsRequest As String
    Dim lsOrder As String
    'Dim lsResult() As String
    'Dim lsComment() As String
    Dim lsMessage As String
    
    Dim lsMSGflag As String
    
    lsMessage = ""
    
    If asVar = "" Then
        Exit Sub
    End If
    
    ClearSpread vasRes
    ClearSpread vasResTemp
    
    iIndex = 0
    lsData = asVar
    
    i = InStr(1, lsData, Chr(13))
    Do While i > 0
        lsTemp = Mid(lsData, 1, i - 1)
        lsData = Mid(lsData, i + 1)
        
        Select Case Left(lsTemp, 1)
        Case "H"
            lsHead = lsTemp
        Case "P"
            lsPatient = lsTemp
        Case "O"
            lsOrder = lsTemp
        Case "Q"
            lsRequest = lsTemp
            lsMSGflag = "Q"
        Case "R"
'            iIndex = iIndex + 1
'            If iIndex > vasRes.MaxRows Then vasRes.MaxRows = iIndex
'
'            'ReDim lsResult(0 To iindex)
'            'ReDim lsComment(0 To iindex)
'            SetText vasRes, lsTemp, iIndex, 1
            
            iIndex = iIndex + 1
            If iIndex > vasResTemp.MaxRows Then vasResTemp.MaxRows = iIndex
            
            'ReDim lsResult(0 To iindex)
            'ReDim lsComment(0 To iindex)
            SetText vasResTemp, lsTemp, iIndex, 1
            
            
            'lsResult(iindex) = lsTemp
            lsMSGflag = "R"
        Case "C"
            'lsComment(iindex) = lsTemp
            SetText vasRes, lsTemp, iIndex, 2
        Case "L"
            lsMessage = lsTemp
        End Select
        
        i = InStr(1, lsData, chrCR)
    Loop
    
    If lsMSGflag = "R" Then
        res = Proc_Result(lsOrder, vasResTemp)
        
    ElseIf lsMSGflag = "Q" Then
        res = Proc_Order(lsRequest)
    End If
End Sub

Function Proc_Order(asReq As String) As Integer
    Dim i As Integer
    Dim iStr As Integer
    Dim iCnt As String
    
    Dim OKFlag As Integer
    
    Dim lsData As String
    Dim lsTemp As String
    
    Dim lsSampleNo As String
    Dim lsID As String
    Dim lsSampleType As String
    Dim lsRackID As String
    Dim lsPosNO As String
    Dim lsKind As String
    Dim lsPriority As String
    
    Dim lsCurDate As String
    
    Dim iRow As Integer
    
    lsData = asReq
    
'    lsCurDate = SeperatorCls(GetDateFull)
    lsCurDate = Format(Date, "yyyymmdd") & Format(Time, "hhnnss")
    OKFlag = -1
    Proc_Order = -1
    
    gOrd.OrderCnt = 0
    gOrd.OrderText = ""
    gOrd.ExamCode = ""
    
    i = 0
    iStr = 1
    iCnt = 0
    
    i = InStr(iStr, lsData, "|")
    Do While i > 0
        iCnt = iCnt + 1
        
        lsTemp = Mid(lsData, iStr, i - iStr)
        lsData = Mid(lsData, i + 1)
        
        If iCnt = 3 Then
            OKFlag = 1
            Exit Do
        End If
        lsTemp = ""
        i = InStr(iStr, lsData, "|")
    Loop
    If OKFlag = 1 Then
        lsData = lsTemp
        
        i = InStr(1, lsData, "/")
        If i > 2 Then
            lsSampleNo = Mid(lsData, 3, i - 3)
            lsData = Mid(lsData, i + 1)
        End If
        
        i = InStr(1, lsData, "/")
        If i > 0 Then
            lsID = Mid(lsData, 1, i - 1)
            lsData = Mid(lsData, i + 1)
        End If
        
        i = InStr(1, lsData, "/")
        If i > 0 Then
            lsSampleType = Mid(lsData, 1, i - 1)
            lsData = Mid(lsData, i + 1)
        End If
        
        i = InStr(1, lsData, "/")
        If i > 0 Then
            lsRackID = Mid(lsData, 1, i - 1)
            lsData = Mid(lsData, i + 1)
        End If
        
        i = InStr(1, lsData, "/")
        If i > 0 Then
            lsPosNO = Mid(lsData, 1, i - 1)
            lsData = Mid(lsData, i + 1)
        End If
        
        i = InStr(1, lsData, "/")
        If i > 0 Then
            lsKind = Mid(lsData, 1, i - 1)
            lsData = Mid(lsData, i + 1)
        End If
        lsPriority = Trim(lsData)
        
        iRow = vasID.DataRowCnt + 1
        If iRow > vasID.MaxRows Then
            vasID.MaxRows = iRow + 1
        End If
        
        
        vasID.SetText colBarCode, iRow, Trim(lsID)
        vasID.SetText colRack, iRow, Trim(lsRackID)
        vasID.SetText colPos, iRow, Trim(lsPosNO)
'''        vasID.SetText colSampleNo, iRow, lsSampleNo
'''        vasID.SetText colSampleType, iRow, lsSampleType
'''        vasID.SetText colKind, iRow, lsKind
'''        vasID.SetText colPriority, iRow, lsPriority
'''        vasID.SetText colOCnt, iRow, "0"
        
        gOrd.SampleType1 = lsSampleType
        
        If Trim(GetText(vasID, iRow, colPName)) = "" Then
            Get_Sample_Info iRow
        End If
        
        res = MakeOrderRecode(lsID, lsPriority, Trim(lsRackID) & "-" & Trim(lsPosNO), lsKind, iRow)
        
        If res > 0 Then
'''            vasID.SetText colOCnt, iRow, gOrd.OrderCnt
            vasID.SetText colState, iRow, "오더"
            Proc_Order = 1
        Else
            
            Proc_Order = 0
        End If
        
        If gOrd.SampleType1 = "0" Then
            Select Case gOrd.SampleType2
            Case "1", "2", "3", "4", "5"
                lsSampleType = gOrd.SampleType2
            Case Else
                lsSampleType = "1"
            End Select
        End If
        
        gOrderCnt = 1
        'lsSampleType = "1"
        If gOrd.OrderCnt > 0 Then
            gOrderMessage = "H|\^&|||host^2|||||H7600|TSDWN^BATCH|P|1" & chrCR & _
                            "P|1" & chrCR & _
                            "O|1|" & lsSampleNo & "^" & SetSpace(lsID, 13, 1) & "^" & lsSampleType & "^" & lsRackID & "^" & lsPosNO & "|" & lsKind & "|" & gOrd.OrderText & "|" & lsPriority & "||" & lsCurDate & "||||N||^^||||||^^^^||||||O" & chrCR & _
                            "L|1|N" & chrCR
                            '& chrETX
            'gOrderMessage = chrSTX & gOrderMessage & CheckSum(gOrderMessage) & chrCR & chrLF
    
            comState = "stTX"
        Else
            gOrderMessage = "H|\^&|||host^2|||||H7600|TSDWN^BATCH|P|1" & chrCR & _
                            "P|1" & chrCR & _
                            "O|1|" & lsSampleNo & "^" & SetSpace(lsID, 13, 1) & "^" & lsSampleType & "^" & lsRackID & "^" & lsPosNO & "|" & lsKind & "||" & lsPriority & "||" & lsCurDate & "||||N||^^||||||^^^^||||||O" & chrCR & _
                            "L|1|N" & chrCR
                            '& chrETX
            'gOrderMessage = chrSTX & gOrderMessage & CheckSum(gOrderMessage) & chrCR & chrLF
    
            comState = "stTX"

        End If
        
        
        'SetFont vasExam, iRow, iRow, 1, vasExam.MaxCols, 9, False
        
        vasActiveCell vasID, iRow, colBarCode
    Else
        Proc_Order = 0
    End If
End Function

Public Function MakeOrderRecode(argCode As String, asEM As String, asRackPos As String, asKind As String, ByVal asRow As Long) As Integer
Dim i, j As Integer
Dim iCnt As Integer

Dim retOrder As String
Dim lsID As String
Dim lsEquipCode As String
Dim lsExamCode As String
Dim lsExamName As String
Dim lsSeqNo As String
Dim lsSampleType As String

Dim iISE As Integer
Dim iISE_r As String

Dim eDate As String

Dim sCnt As String
Dim sRv As String
Dim lsReceCode As String


    ClearSpread vasRes
    
    iCnt = 0
    MakeOrderRecode = -1
    
    gOrd.OrderCnt = 0
    gOrd.OrderText = ""
    gOrd.ExamCode = ""
    gOrd.SampleType2 = ""
    
    retOrder = ""
    ClearSpread vasTemp
    
    If argCode = "" Then
        MakeOrderRecode = -1
        Exit Function
    End If
    
    eDate = Trim(Text_Today.Text)
    'argCode = Trim(argCode)
    lsID = Trim(argCode)

'    '처음 검사 샘플
    
'''    SQL = "SELECT  b.wd_code ,max(b.wd_date) ,'W' ,a.pe_sujinja , a.pe_jumin  " & vbCrLf & _
'''          "From person a, wchdat b " & vbCrLf & _
'''          "WHERE a.pe_chart = '" & lsID & "' " & vbCrLf & _
'''          "  and a.pe_chart = b.wd_chart " & vbCrLf & _
'''          "  and b.wd_code in (" & gAllExam & ") " & vbCrLf & _
'''          "  and b.wd_end_dep = '2' and wd_cancel = '0' " & vbCrLf & _
'''          "group by b.wd_code ,b.wd_date ,a.pe_sujinja , a.pe_jumin "
'''
'''    SQL = SQL & vbCrLf & "union SELECT  b.id_code ,max(b.id_date) ,'I' ,a.pe_sujinja , a.pe_jumin  " & vbCrLf & _
'''          "From person a, ichdat b " & vbCrLf & _
'''          "WHERE a.pe_chart = '" & lsID & "' " & vbCrLf & _
'''          "  and a.pe_chart = b.id_chart " & vbCrLf & _
'''          "  and b.id_code in (" & gAllExam & ") " & vbCrLf & _
'''          "  and b.id_end_dep = '2' and id_cancel = '0' " & vbCrLf & _
'''          "group by b.id_code ,b.id_date ,a.pe_sujinja , a.pe_jumin "
    
    Clear_XML_Exam
    sRv = Online_XML(gXml_S07, Trim(lsID))
    lsReceCode = ""
    
   
    
    
    For i = 0 To UBound(gExam_Select)

        If lsReceCode = "" Then
            lsReceCode = "'" & Trim(gExam_Select(i).TST_CD) & "'"
        Else
            lsReceCode = lsReceCode & ",'" & Trim(gExam_Select(i).TST_CD) & "'"
        End If
        
    Next i
   
    If lsReceCode = "" Then
        lsReceCode = "''"
    End If
    
    ClearSpread vasTemp
    
    SQL = "select examcode, equipcode, examname, seqno from equipexam where equipno = '" & gEquip & "' and examcode in (" & lsReceCode & ")"
    res = db_select_Vas(gLocal, SQL, vasTemp)
'''    res = db_select_Vas(gServer, SQL, vasTemp)
    If res = -1 Then
        SaveQuery SQL
        'Exit Function
    End If


    iISE = -1
    If vasTemp.DataRowCnt > 0 Then

        retOrder = ""
        ClearSpread vasRes
        
        For i = 1 To vasTemp.DataRowCnt
            
            
            lsExamCode = Trim(GetText(vasTemp, i, 1))
            lsEquipCode = Trim(GetText(vasTemp, i, 2))
            lsExamName = Trim(GetText(vasTemp, i, 3))
            lsSeqNo = Trim(GetText(vasTemp, i, 4))
            
            'Serum 만 검사.
            lsSampleType = gOrd.SampleType1
            
            retOrder = retOrder & "^^^" & lsEquipCode & "/\"
            
            If vasRes.MaxRows < i Then vasRes.MaxRows = i
                    
            SetText vasRes, lsEquipCode, i, colEquipCode
            SetText vasRes, lsExamCode, i, colExamCode
            SetText vasRes, lsExamName, i, colExamName
                    
            Save_Local_One_1 asRow, i, "A"
    
        Next i
    Else

        MakeOrderRecode = 0
    End If
    If Len(retOrder) > 0 Then
        gOrd.OrderText = Mid(retOrder, 1, Len(retOrder) - 1)
    Else
        gOrd.OrderText = ""
    End If
    
    gOrd.OrderCnt = i
    gOrd.ExamCode = lsExamCode
    
    MakeOrderRecode = 1

End Function

Function Proc_Result(asOrd As String, ByVal argSpread As vaSpread) As Integer
    Dim i, j, k, iArr, lResRow As Long
    Dim iStr As Integer
    Dim iCnt As Integer
    
    Dim lsData As String
    Dim lsTemp As String
    Dim lsSampleType As String
    Dim lsSpecimenID As String
    Dim lsOrder As String
    Dim lsID As String
    Dim lsRackID As String
    Dim lsPosNO As String
    Dim lsPriority As String
    
    Dim lsExamCode As String
    Dim lsExamDate As String
    Dim lsEquipCode As String
    Dim lsResult As String
    Dim lsUnit As String
    Dim lsRef As String
    Dim lsState As String
    Dim lsComment As String
    
    Dim iRow As Integer
    
    Dim sCnt As String
    
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim strEquipRes As String
    
    Proc_Result = -1
    
    gOrd.OrderCnt = 0
    gOrd.OrderText = ""
    lsData = asOrd
    i = 0
    iStr = 1
    iCnt = 0
    lsID = ""
    i = InStr(iStr, lsData, "|")
    Do While i > 0
        iCnt = iCnt + 1
        
        lsTemp = Mid(lsData, iStr, i - iStr)
        lsData = Mid(lsData, i + 1)
        
        Select Case iCnt
        Case 3
            lsSpecimenID = lsTemp
        Case 5
            lsOrder = lsTemp
        Case 6
            lsPriority = lsTemp
            Exit Do
        Case 23
            lsExamDate = lsTemp
            Exit Do
        Case Else
        End Select
        
        lsTemp = ""
        i = InStr(iStr, lsData, "|")
    Loop
    
    'lsExamDate = Left(lsExamDate, 4) & "-" & Mid(lsExamDate, 5, 2) & "-" & Mid(lsExamDate, 7, 2) & " " & Mid(lsExamDate, 9, 2) & ":" & Mid(lsExamDate, 11, 2) & ":" & Mid(lsExamDate, 13, 2)
    lsExamDate = Format(CDate(GetDateFull), "yyyy-mm-dd hh:nn:ss")
    
    i = InStr(1, lsSpecimenID, "^")
    If i > 0 Then
        lsSpecimenID = Mid(lsSpecimenID, i + 1)
        'lsID = Trim(Left(lsSpecimenID, 13))
        'lsSpecimenID = Mid(lsSpecimenID, 14)
        i = InStr(1, lsSpecimenID, "^")
        If i > 0 Then
            lsID = Trim(Left(lsSpecimenID, i - 1))
            lsSpecimenID = Mid(lsSpecimenID, i + 1)
            i = InStr(1, lsSpecimenID, "^")
            If i > 0 Then
                lsSampleType = Trim(Left(lsSpecimenID, i - 1))
                lsSpecimenID = Mid(lsSpecimenID, i + 1)
                'lsRackID = Mid(lsSpecimenID, 1, i - 1)
                i = InStr(1, lsSpecimenID, "^")
                If i > 0 Then
                    lsRackID = Left(lsSpecimenID, i - 1)
                    lsPosNO = Trim(Mid(lsSpecimenID, i + 1))
                End If
            End If
        End If
    End If
    iRow = -1
    For i = vasID.DataRowCnt To 1 Step -1
        If Trim(GetText(vasID, i, colBarCode)) = lsID Then
            iRow = i
            Exit For
        End If
    Next i
    If iRow = -1 Then
        iCnt = 0
        i = InStr(1, lsOrder, "/")
        Do While i > 0
            iCnt = iCnt + 1
            lsOrder = Mid(lsOrder, i + 1)
            i = InStr(1, lsOrder, "/")
        Loop
        
        iRow = vasID.DataRowCnt + 1
        If iRow > vasID.MaxRows Then
            vasID.MaxRows = iRow + 1
        End If

        vasID.SetText colBarCode, iRow, Trim(lsID)
        vasID.SetText colRack, iRow, Trim(lsRackID)
        vasID.SetText colPos, iRow, Trim(lsPosNO)
        
        
'''        vasID.SetText colPriority, iRow, lsPriority
'''        vasID.SetText colSampleType, iRow, lsSampleType
    End If
    
    vasID.SetText colState, iRow, "Result"
    
    If Trim(GetText(vasID, iRow, colPName)) = "" Then
        Get_Sample_Info iRow
    End If
    
    'vasID_Click colBarCode, iRow
    SetForeColor vasID, iRow, iRow, colBarCode, colState, 0, 0, 0
    
    For iArr = 1 To argSpread.DataRowCnt
        iStr = 1
        iCnt = 0
        lsData = GetText(argSpread, iArr, 1)
        If lsData <> "" Then
            i = InStr(iStr, lsData, "|")
            Do While i > 0
                iCnt = iCnt + 1
                lsTemp = Mid(lsData, iStr, i - iStr)
                lsData = Mid(lsData, i + 1)
                
                Select Case iCnt
                Case 3
                    lsEquipCode = lsTemp
                    j = InStr(1, lsEquipCode, "/")
                    If j > 0 Then
                        lsEquipCode = Mid(lsEquipCode, 4, j - 4)
                    Else
                    lsEquipCode = Mid(lsEquipCode, 4, Len(lsEquipCode) - 4)
                    End If
                Case 4
                    lsResult = lsTemp
                    'Exit Do
                Case 5
                    lsUnit = lsTemp
                Case 7
                    lsRef = lsTemp
                    If UCase(lsRef) = "N" Then lsRef = ""
                    'If UCase(lsRef) = "H" Then lsRef = ""
                    If UCase(lsRef) = "L" Then lsRef = ""
                    
                    Exit Do
                Case 9
                    lsState = lsTemp
                    Exit Do
                Case Else
                End Select
                
                lsTemp = ""
                i = InStr(iStr, lsData, "|")
            
            Loop
            
            lResRow = iArr
                        
            lsExamCode = ""
            
            If vasRes.MaxRows < lResRow Then vasRes.MaxRows = lResRow
            
            vasRes.SetText colEquipCode, lResRow, lsEquipCode
'''            vasRes.SetText colEquipRes, lResRow, lsResult
'''            vasRes.SetText colBarCode, lResRow, lsID
            SQL = "select examcode, examname, seqno, resprec  from equipexam where equipcode = '" & lsEquipCode & "'"
            res = db_select_Col(gLocal, SQL)
            
            If res > 0 Then
            
                lsExamCode = Trim(gReadBuf(0))
                lsExamName = Trim(gReadBuf(1))
                lsSeqNo = Trim(gReadBuf(2))
                strEquipRes = lsResult
                Dim sGiho As String
                sGiho = ""
                If Mid(sResult, 1, 1) = ">" Or Mid(sResult, 1, 1) = "<" Then
                    sGiho = Mid(sResult, 1, 1)
                    sResult = Trim(Mid(sResult, 2))
                End If
                
                If IsNumeric(lsResult) = True Then
                    If Trim(gReadBuf(3)) = "0" Then
                        lsResult = Format(lsResult, "0")
                    ElseIf Trim(gReadBuf(3)) = "1" Then
                        lsResult = Format(lsResult, "0.0")
                    ElseIf Trim(gReadBuf(3)) = "2" Then
                        lsResult = Format(lsResult, "0.00")
                    ElseIf Trim(gReadBuf(3)) = "3" Then
                        lsResult = Format(lsResult, "0.000")
                    End If
                
                End If
                lsResult = sGiho & lsResult
                
                vasRes.SetText colExamCode, lResRow, Trim(lsExamCode)
                vasRes.SetText colExamName, lResRow, Trim(lsExamName)
                vasRes.SetText colSeqNo, lResRow, Trim(lsSeqNo)
                
'''                lsResult = SetResult(lResRow, k)
                vasRes.SetText colResult, lResRow, lsResult
                vasRes.SetText colResult1, lResRow, strEquipRes
                
                Save_Local_One_1 iRow, lResRow, "B"
'''            k = -1
'''            For i = LBound(gArr_ExamCode) To UBound(gArr_ExamCode)
''''                Debug.Print lsEquipCode & " : " & Trim(gArr_ExamCode(i, 1))
''''                Debug.Print lsExamCode & " : " & Trim(gArr_ExamCode(i, 2))
'''                If lsEquipCode = Trim(gArr_ExamCode(i, 1)) Then
'''                    lsExamName = Trim(gArr_ExamCode(i, 3))
'''                    lsSeqNo = Trim(gArr_ExamCode(i, 6))
'''                    For j = 1 To vasTemp.DataRowCnt
'''                        If Trim(gArr_ExamCode(i, 2)) = Trim(GetText(vasTemp, j, 1)) _
'''                            And Trim(gArr_ExamCode(i, 12)) = lsSampleType Then
'''                            k = i
'''                            lsExamCode = Trim(gArr_ExamCode(i, 2))
'''                            Exit For
'''                        End If
'''                    Next j
'''                    If k > 0 Then Exit For
'''                End If
'''            Next i
            
'''                vasRes.SetText colExamName, lResRow, lsExamName
'''                vasRes.SetText colSeqNo, lResRow, lsSeqNo
            End If
                                    
            
            
            
        End If
    Next iArr
    
    
'''    SetText vasID, argSpread.DataRowCnt, iRow, colRCnt
    
    If MnTransAuto.Checked = True Then
        res = Insert_Data(iRow)
        
        If res = -1 Then
            SetForeColor vasID, iRow, iRow, 1, colState, 255, 0, 0
            SetText vasID, "Failed", iRow, colState
        Else
           
            SetBackColor vasID, iRow, iRow, 1, colState, 202, 255, 112
            SetText vasID, "Trans", iRow, colState
            
            SQL = " Update pat_res Set " & vbCrLf & _
                  " sendflag = 'C' " & vbCrLf & _
                  " Where equipno = '" & gEquip & "' " & vbCrLf & _
                  " And barcode = '" & Trim(GetText(vasID, iRow, colBarCode)) & "' "
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL
                Exit Function
            End If
            
        End If
    Else
        SetBackColor vasID, iRow, iRow, 1, vasID.MaxCols, 202, 255, 112
        SetText vasID, "Result", iRow, colState
    End If
'''        res = ToServer(iRow)
'''        If res = 1 Then
'''            SetText vasID, "완료", iRow, colState
'''            SetForeColor vasID, iRow, iRow, colBarCode, colState, 0, 0, 0
'''            SetBackColor vasID, iRow, iRow, colBarCode, colState, 202, 255, 112
'''        Else
'''            SetText vasID, "실패", iRow, colState
'''            SetForeColor vasID, iRow, iRow, colBarCode, colState, 255, 0, 0
'''            SetBackColor vasID, iRow, iRow, colBarCode, colState, 255, 255, 255
'''        End If
   
    
End Function

Sub SendOrder()
Dim sSendOrder As String
    
    If Len(gOrderMessage) > 240 Then
        
        If gOrderCnt = 8 Then
            gOrderCnt = 0
        End If
        
        sSendOrder = CStr(gOrderCnt) & Left(gOrderMessage, 240) & chrETB
        gOrderMessage = Mid(gOrderMessage, 241)
        
        sSendOrder = chrSTX & sSendOrder & CheckSum(sSendOrder) & chrCR & chrLF
        SaveQuery sSendOrder, 1
        
        gOrderCnt = gOrderCnt + 1
        comSend = "stENQ"
        
        gPreMsg = sSendOrder
        
        Save_Raw_Data "[Tx]" & gPreMsg
        Winsock1.SendData sSendOrder
        
    Else
        If gOrderCnt = 8 Then
            gOrderCnt = 0
        End If
        
        sSendOrder = CStr(gOrderCnt) & gOrderMessage & chrETX
        sSendOrder = chrSTX & sSendOrder & CheckSum(sSendOrder) & chrCR & chrLF
                
        gOrderMessage = ""
        comSend = "stOrder"
        
        gPreMsg = sSendOrder
        
        Save_Raw_Data "[Tx]" & gPreMsg
        Winsock1.SendData sSendOrder
    End If
End Sub

'''Sub SendOrder()
'''    If gOrderMessage <> "" Then
'''        gPreData = gOrderMessage
'''        gOrderMessage = ""
'''
'''        Save_Raw_Data "[TX]" & gPreData
'''        MSComm1.Output = gPreData
'''    End If
'''End Sub

Sub XE2100_ASTM(asData As String)
'ASTM

    Dim MyVar As String
    Dim MyRet As String
    
    Dim i As Integer
    Dim j As Integer
    Dim iCnt As Integer
    Dim jCnt As Integer
    Dim aCnt As Integer
    Dim bCnt As Integer
    Dim ii  As Integer
    
    Dim iRow As Integer
    Dim lRow As Integer
    Dim liRet As Integer
    
    Dim lsDistinctII As String
    Dim lsInqueryMode As String
    Dim lsDate      As String
    Dim lsTime      As String
    Dim lsRack      As String
    Dim lsTube      As String
    Dim lsID        As String
    Dim lsIDInfo    As String
    Dim lsPName     As String
    
    Dim lsTemp      As String
    Dim lsData      As String
    
    Dim lsTestID    As String
    Dim lsResult    As String
    Dim lsFlag      As String
    
    Dim lsTemp1     As String
    Dim lsMessage   As String
    
    Dim lsExamCode  As String
    Dim lsRsCode    As String
    Dim lsExamName  As String
    Dim lsPoint     As String
    Dim sTmpStr     As String
    Dim lsSelExam   As String
    
    Dim sDate       As String
    Dim iExamCnt    As Integer
    
    Dim sLen, sLen2 As String
    
    
    sDate = Format(Text_Today.Text, "yyyymmdd")

    j = 1
    
    Select Case Mid(asData, 3, 1)
    Case "H"    'Header
        gPreRow = -1
        
        lsMessage = ""
        
        ClearSpread vasRes
        
    Case "P"    'Patient
        gPatFlag = -1
        
    Case "Q"    'Request
        gRecodeType = "Q"
        
        ClearSpread vasTemp
        ClearSpread vasOrder
        ClearSpread vasOrderBuf
        
        sLen = InStr(1, asData, "|")
        asData = Mid(asData, sLen + 1)
        
        sLen = InStr(1, asData, "|")
        asData = Mid(asData, sLen + 1)
        
        sLen = InStr(1, asData, "|")
        gsBarCode = Mid(asData, 1, sLen - 1)
        
        
        sLen = InStr(1, gsBarCode, "^")
        gsRackNo = Trim(Mid(gsBarCode, 1, sLen - 1))      'Rack
        gsBarCode = Mid(gsBarCode, sLen + 1)
        
        sLen = InStr(1, gsBarCode, "^")
        gsPosNo = Trim(Mid(gsBarCode, 1, sLen - 1))        'Tube
        If Len(gsPosNo) = 1 Then
            gsPosNo = Format(gsPosNo, "0#")
        End If
        
        gsBarCode = Mid(gsBarCode, sLen + 1)
        
        sLen = InStr(1, gsBarCode, "^")
        gsBarCode = Trim(Mid(gsBarCode, 1, sLen - 1))    '검체번호
        gAttribute = Mid(gsBarCode, sLen + 1)           'Attribute
        
'        If Len(gBarcode) = 12 Then
'
'        Else
'            If UCase(Left(gBarcode, 3)) = "ERR" Then    '바코드리딩 에러
'                gBarcode = CInt(gRack) & gTube
'            Else                                        '메뉴얼
'                gRack = Mid(gBarcode, 1, 1)
'                gTube = Mid(gBarcode, 2, 2)
'            End If
'        End If
    
        glRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < glRow + 1 Then
            vasID.MaxRows = glRow + 1
        End If
        
        glRow = -1
        For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colBarCode)) = gsBarCode Then
                glRow = i
                
                vasID_DblClick 2, glRow
                
                Exit For
            End If
        Next i
        
        '2004/06/16 이상은========================================================
        'Order 전송뒤 Clear시 다시 바코드 스캔 안하고 결과 넘어오도록 수정
        If glRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
            glRow = vasID.DataRowCnt + 1
            If glRow > vasID.MaxRows Then
                vasID.MaxRows = glRow + 1
            End If
            vasActiveCell vasID, glRow, colBarCode
            SetText vasID, gsBarCode, glRow, colBarCode
            SetText vasID, gsRackNo, glRow, colRack
            SetText vasID, gsPosNo, glRow, colPos
        End If
        '==========================================================================
                   
        '환자정보 가져오기
        If Trim(GetText(vasID, glRow, colPID)) = "" Then
            Get_Sample_Info glRow
        End If
        
        'Order 만들기
        Make_Order_ASTM gsBarCode, glRow
                     
    Case "O"    'Test Order
        sLen = InStr(1, asData, "|")
        asData = Mid(asData, sLen + 1)
        
        sLen = InStr(1, asData, "|")
        asData = Mid(asData, sLen + 1)
        
        sLen = InStr(1, asData, "|")
        asData = Mid(asData, sLen + 1)
        
        sLen = InStr(1, asData, "|")
        gsBarCode = Mid(asData, 1, sLen - 1)
        
        sLen = InStr(1, gsBarCode, "^")
        gsRackNo = Trim(Mid(gsBarCode, 1, sLen - 1))      'Rack
        gsBarCode = Mid(gsBarCode, sLen + 1)
        
        sLen = InStr(1, gsBarCode, "^")
        gsPosNo = Trim(Mid(gsBarCode, 1, sLen - 1))        'Tube
        If Len(gsPosNo) = 1 Then
            gsPosNo = Format(gsPosNo, "0#")
        End If
        
        gsBarCode = Mid(gsBarCode, sLen + 1)
        
        sLen = InStr(1, gsBarCode, "^")
        gsBarCode = Trim(Mid(gsBarCode, 1, sLen - 1))    '검체번호
        
        
'        If Len(gsBarCode) = 12 Then
'
'        Else
'            If UCase(Left(gsBarCode, 3)) = "ERR" Then    '바코드리딩 에러
'                gsBarCode = CInt(gsRackNo) & gsPosNo
'            Else                                        '메뉴얼
'                gsRackNo = Mid(gsBarCode, 1, 1)
'                gsPosNo = Mid(gsBarCode, 2, 2)
'            End If
'        End If
        
        glRow = -1
        For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colBarCode)) = gsBarCode Then
                glRow = i
                
                If gPatFlag = -1 Then
                    vasID_DblClick 2, glRow
                    gPatFlag = 1
                    vasActiveCell vasID, glRow, 2
                End If
                
                Exit For
            End If
        Next i
        
        '2004/06/16 이상은========================================================
        'Order 전송뒤 Clear시 다시 바코드 스캔 안하고 결과 넘어오도록 수정
        If glRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
            glRow = vasID.DataRowCnt + 1
            If glRow > vasID.MaxRows Then
                vasID.MaxRows = glRow + 1
            End If
            vasActiveCell vasID, glRow, colBarCode
            SetText vasID, gsBarCode, glRow, colBarCode
            SetText vasID, gsRackNo, glRow, colRack
            SetText vasID, gsPosNo, glRow, colPos
        End If
        '==========================================================================
        
        '환자정보 가져오기
        If Trim(GetText(vasID, glRow, colPID)) = "" Then
            Get_Sample_Info glRow
        End If
     
        '2010.03.11 이상은*********************************
        res = Online_XML(gXml_S07, Trim(gsBarCode))
        
        ClearSpread vasTemp
        
        lsSelExam = ""
        
        gSelExam = ""
        
        
        For ii = 0 To UBound(gExam_Select)
            vasTemp.SetText 1, ii + 1, gExam_Select(ii).TST_CD
            If lsSelExam = "" Then
                lsSelExam = "'" & Trim(GetText(vasTemp, ii + 1, 1)) & "'"
            Else
                lsSelExam = lsSelExam & ",'" & Trim(GetText(vasTemp, ii + 1, 1)) & "'"
            End If
        Next ii
        
        gSelExam = lsSelExam
        '**************************************************
        
    Case "R"    'Result
        gRecodeType = "R"

        SetText vasID, "Result", glRow, colState
        
        If vasRes.MaxRows = 0 Then
            vasRes.MaxRows = 1
            iRow = vasRes.MaxRows
        Else
            vasRes.MaxRows = vasRes.MaxRows + 1
            iRow = vasRes.MaxRows
        End If
        
        iCnt = 0
        i = InStr(1, asData, "|")
        Do While i > 0
            iCnt = iCnt + 1
            
            lsTemp = Mid(asData, 1, i - 1)
            asData = Mid(asData, i + 1)
            
            Select Case iCnt
            Case 2  '결과갯수
                gRCnt = Trim(lsTemp)
                
            Case 3  'TestID
                If lsTemp <> "" Then
                    lsTemp = Mid(lsTemp, 5)
                    j = InStr(1, lsTemp, "^")
                    
                    If j > 0 Then
                        lsTestID = Left(lsTemp, j - 1)
                        gsTestID = lsTestID
                    Else
                        lsTestID = Trim(lsTemp)
                        gsTestID = lsTestID
                    End If
                    
                    If lsTestID = "" Then
                        Exit Sub
                    End If
                End If
    
            Case 4 '결과 Data
                gReadBuf(0) = "0"
                SQL = "Select ExamCode, ExamName, resprec From EquipExam" & vbCrLf & _
                      " Where Equipno = '" & gEquip & "' " & vbCrLf & _
                      "  And  EquipCode = '" & Trim(gsTestID) & "'" & vbCrLf & _
                      "  And  ExamCode in (" & gSelExam & ") "
                res = db_select_Col(gLocal, SQL)
                
                If res = 1 And gReadBuf(0) <> "" Then
                    lsExamCode = Trim(gReadBuf(0))
                    lsExamName = Trim(gReadBuf(1))
                    lsPoint = Trim(gReadBuf(2))
                    
                    j = vasRes.DataRowCnt + 1
                    
                    lsResult = Trim(lsTemp)
                    
'                    If gsTestID = "WBC" Then
'                        lsResult = Format(lsResult, "#0.0")
'                    End If
                    
                    If IsNumeric(lsResult) Then
                        '소수점처리
                        If IsNumeric(lsPoint) Then
                            If CInt(lsPoint) > 0 Then
                                sTmpStr = "#0."
                                For i = 1 To CInt(lsPoint)
                                    sTmpStr = sTmpStr & "0"
                                Next i
                            Else
                                sTmpStr = "#0"
                            End If
                            lsResult = Format(lsResult, sTmpStr)
                        End If
                    
                        SetText vasRes, gsBarCode, j, colBarCode                '검체번호
                        SetText vasRes, gsTestID, j, colEquipCode               '장비코드
                        SetText vasRes, lsExamCode, j, colExamCode              '검사코드
                        SetText vasRes, lsExamName, j, colExamName              '검사명
                        SetText vasRes, lsResult, j, colResult                  '검사결과
                        SetText vasRes, lsResult, j, colResult1                 '검사결과
                        
                        Save_Local_One_1 glRow, j, "A"
                    Else
                        '2004/06/09 이상은
                        'SetText vasRes, "", j, colResult
                        '================================================================
                        '결과값 없어도 항목 디스플레이 되도록
                        SetText vasRes, gsBarCode, j, colBarCode                '검체번호
                        SetText vasRes, gsTestID, j, colEquipCode               '장비코드
                        SetText vasRes, lsExamCode, j, colExamCode              '검사코드
                        SetText vasRes, lsExamName, j, colExamName              '검사명
                        SetText vasRes, "", j, colResult                        '검사결과
                        SetText vasRes, "", j, colResult1                       '검사결과
                            
                        Save_Local_One_1 glRow, j, "A"
                        '================================================================
                    End If
                Else
                    gReadBuf(0) = "0"
                    SQL = "Select ExamCode, ExamName From EquipExam" & vbCrLf & _
                          " Where Equipno = '" & gEquip & "' " & vbCrLf & _
                          "  And  EquipCode = '" & Trim(gsTestID) & "'"
                    res = db_select_Col(gLocal, SQL)
                    If res = 1 Then
                        j = vasRes.DataRowCnt + 1
                        
                        lsResult = Trim(lsTemp)
                    
                        SetText vasRes, gsBarCode, j, colBarCode                '검체번호
                        SetText vasRes, gsTestID, j, colEquipCode               '장비코드
                        SetText vasRes, "", j, colExamCode              '검사코드
                        SetText vasRes, Trim(gReadBuf(1)), j, colExamName              '검사명
                        SetText vasRes, lsResult, j, colResult                  '검사결과
                        SetText vasRes, lsResult, j, colResult1                 '검사결과
                        
                        Save_Local_One_1 glRow, j, "A"
                    End If
                End If
            
            Case "7"        'Flag
                lsFlag = Trim(lsTemp)

                If lsFlag = "A" Then
                    lsTemp1 = ""

                    Select Case gsTestID
                    'WBC*************************
                    Case "WBC_Abn_Scattergram"
                        lsTemp1 = "WBC Abn Scg"
                    Case "NRBC_Abn_Scattergram"
                        lsTemp1 = "NRBC Abn Scg"
                    Case "Neutropenia"
                        lsTemp1 = "Neutro-"
                    Case "Neutrophilia"
                        lsTemp1 = "Neutro+"
                    Case "Lymphopenia"
                        lsTemp1 = "Lympho-"
                    Case "Lymphocytosis"
                        lsTemp1 = "Lympho+"
                    Case "Monocytosis"
                        lsTemp1 = "Mono+"
                    Case "Eosinophilia"
                        lsTemp1 = "Eo+"
                    Case "Basophilia"
                        lsTemp1 = "Baso+"
                    Case "Leukocytopenia"
                        lsTemp1 = "Leuko-"
                    Case "Leukocytosis"
                        lsTemp1 = "Leuko+"
                    Case "NRBC_Present", "Blasts?", "Left_Shift?", "NRBC?"
                        lsTemp1 = gsTestID

                    Case "Immulature_Gran?"
                        lsTemp1 = "Imm Gran?"
                    Case "Atypical_Lympho?"
                        lsTemp1 = "Atypical Ly?"
                    Case "Abn_Lympho/L_Blasts?"
                        lsTemp1 = "Abn Ly/L_Bl?"
                    Case "RBC_Lyse Resistance?"
                        lsTemp1 = "RBC Lyse Res?"

                    'RBC*************************
                    Case "RBC_Abn_Distribution"
                        lsTemp1 = "RBC Abn Dst"
                    Case "Dimorphic_Population"
                        lsTemp1 = "Dimorph Pop"
                    Case "RET_Abn_Scattergram"
                        lsTemp1 = "RET Abn Scg"
                    Case "Reticulocytosis"
                        lsTemp1 = "Reticulo"
                    Case "Anisocytosis"
                        lsTemp1 = "Aniso"
                    Case "Microcytosis"
                        lsTemp1 = "Micro"
                    Case "Macrocytosis"
                        lsTemp1 = "Macro"

                    Case "Hypochromia", "Anemia", "HGB_Defect?", "Fragments?"
                        lsTemp1 = gsTestID

                    Case "Erythrocytosis"
                        lsTemp1 = "Erythro+"

                    Case "RBC_Agglutination?"
                        lsTemp1 = "RBC Agglut?"
                    Case "Turbidity/HGB Interference?"
                        lsTemp1 = "Turb/HGB?"
                    Case "Iron_Deficiency?"
                        lsTemp1 = "Iron Def?"

                    'PLT*************************
                    Case "PLT_Abn_Scattergram"
                        lsTemp1 = "PLT Abn Scg"
                    Case "PLT_Abn_Distribution"
                        lsTemp1 = "PLT Abn Dst"
                    Case "Thrombocytopenia"
                        lsTemp1 = "Thrombo-"
                    Case "Thrombocytosis"
                        lsTemp1 = "Thrombo+"
                    Case "PLT_Clumps?"
                        lsTemp1 = gsTestID
                    Case "PLT_Clumps(S)?"
                        lsTemp1 = "PLT C(S)?"
                    Case Else
                        lsTemp1 = gsTestID
                    End Select

'                    If lsMessage = "" Then
'                        lsMessage = lsTemp1
'                    Else
'                        lsMessage = lsMessage & "," & lsTemp1
'                    End If

                    '메모결과 입력
'                    Save_ResMemo glRow, lsMessage

'                    lsTemp1 = gsTestID
                    
                    Save_ResMemo glRow, lsTemp1

                End If
            End Select
     
            lsTemp = ""
            lsTestID = ""
            lsResult = ""
            i = InStr(1, asData, "|")
        Loop
        
    Case "L"
        If glRow <> -1 And gRecodeType = "R" Then
            If chkMode.Value = 1 Then
                vasID.Col = 1
                vasID.Row = glRow
                vasID.Value = 1
            
                res = Insert_Data(glRow)
                If res = 1 Then
                    SetBackColor vasID, glRow, glRow, colCheckBox, colState, 202, 255, 112
                    SetText vasID, "완료", glRow, colState
                        SQL = "update pat_res set sendflag = 'C' where examdate = '" & Format(Text_Today.Text, "YYYYMMDD") & "' " & CR & _
                              "and equipno = '" & gEquip & "' And barcode = '" & Trim(GetText(vasID, glRow, colBarCode)) & "' "
                        res = SendQuery(gLocal, SQL)
                        
                        vasID.Row = glRow
                        vasID.Col = 1
                        
                        vasID.Value = 0
                ElseIf res = -1 Then
                    SetForeColor vasID, glRow, glRow, colCheckBox, colState, 255, 0, 0
                    SetText vasID, "실패", glRow, colState
                End If
            End If
        End If
    End Select
    
End Sub


Function Make_Order_ASTM(argNo As String, argRow As Long) As String
'Order Text 만들기
    Dim sRetOrder(3) As String    'Order Text넣을 변수
    Dim sOrder      As String
    Dim sOrder1     As String
    
    Dim sOrderGubun As String
    
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
        
    Dim lsExamCode  As String
    Dim sExamCode   As String     '검사코드
    Dim sEquipCode  As String
    Dim sOrdGubun   As String     '오더구분
    Dim sPsex       As String
    
    Dim sDate       As String

    Dim llrow       As Long
    Dim llRow_Order As Long
    
    Dim iCnt_Ord    As Integer    'Order conut
    Dim sOCnt       As Long
    
    Dim sReceNo     As String
    Dim sBarcode    As String
    
    Dim sRet        As String
    
    If argNo = "" Then
        Exit Function
    End If
    
    sDate = SeperatorCls(Text_Today.Text)
    
    sReceNo = Trim(GetText(vasID, argRow, colPID))
    sBarcode = Trim(GetText(vasID, argRow, colBarCode))
    
    sPsex = Trim(GetText(vasID, argRow, colPSex))
    If sPsex = "" Then sPsex = "U"

    sRet = Online_XML(gXml_S07, Trim(sBarcode))

    lsExamCode = ""

    ClearSpread vasTemp

    For i = 0 To UBound(gExam_Select)
        vasTemp.SetText 1, i + 1, gExam_Select(i).TST_CD
        If lsExamCode = "" Then
            lsExamCode = "'" & Trim(GetText(vasTemp, i + 1, 1)) & "'"
        Else
            lsExamCode = lsExamCode & ",'" & Trim(GetText(vasTemp, i + 1, 1)) & "'"
        End If
    Next i

        
    For i = 1 To 3
        sRetOrder(i) = "0"
    Next i
    
    If vasTemp.DataRowCnt > 0 Then
        llRow_Order = 1
    
        gCurMsgCnt = 1
        'Head
        'gHeader = "H|\^&|||XE-2100^00-32^A4349^^^^98313519||||||||E1394-97" & chrCR & chrETX   '2010.01.19 이상은
        gHeader = "H|\^&|||XE-2100^00-32^F5187^^^^98313519||||||||E1394-97" & chrCR & chrETX
        gHeader = chrSTX & CCur(gCurMsgCnt) & gHeader & CheckSum(CStr(gCurMsgCnt) & gHeader) & chrCR & chrLF
        
        SetText frmInterface.vasOrder, gHeader, llRow_Order, 1
        
        gCurMsgCnt = gCurMsgCnt + 1
        If gCurMsgCnt = 8 Then
            gCurMsgCnt = 0
        End If
        
        llRow_Order = llRow_Order + 1
        If llRow_Order > frmInterface.vasOrder.MaxRows Then
            frmInterface.vasOrder.MaxRows = llRow_Order
        End If
        
        'Patient
        gPatient = "P|1|||" & Trim(GetText(vasID, argRow, colPID)) & "||||" & sPsex & chrCR & chrETX
        gPatient = chrSTX & CCur(gCurMsgCnt) & gPatient & CheckSum(CStr(gCurMsgCnt) & gPatient) & chrCR & chrLF
        
        SetText frmInterface.vasOrder, gPatient, llRow_Order, 1
        
        gCurMsgCnt = gCurMsgCnt + 1
        If gCurMsgCnt = 8 Then
            gCurMsgCnt = 0
        End If
        
        llRow_Order = llRow_Order + 1
        If llRow_Order > frmInterface.vasOrder.MaxRows Then
            frmInterface.vasOrder.MaxRows = llRow_Order
        End If
        
        'Order
        sOCnt = 1
        
        For i = 1 To 3
            sRetOrder(i) = "0"
        Next i

        k = 1
        Do While k <= vasTemp.DataRowCnt
            sExamCode = Trim(GetText(vasTemp, k, 1))
            For j = 1 To UBound(gArrEquip())
                If sExamCode = gArrEquip(j, 3) Then
                    Select Case gArrEquip(j, 7)
                    Case "C"
                        sRetOrder(1) = "1"
                    Case "D"
                        sRetOrder(2) = "1"
                    Case "R"
                        sRetOrder(3) = "1"
                    End Select
                    
                    Exit For
                End If
            Next j
            
            k = k + 1
        Loop
        
        sOrder = ""
        sOrder1 = ""
        
        For i = 1 To 3
            sOrder = sOrder & sRetOrder(i)
        Next i
        
        sOrderGubun = sOrder
        
        If sOrder <> "" And sOrderGubun = "100" Then      'CBC
            sOrder = "O|" & sOCnt & "|" & gsRackNo & "^" & gsPosNo & "^" & argNo & "^" & "C||"
            sOrder = sOrder & "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\"
            sOrder = sOrder & "^^^^RDW-CV\^^^^RDW-SD\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT"
            sOrder = sOrder & "|||||||N||||||||||||||Q" & chrCR & chrETX
        ElseIf sOrder <> "" And sOrderGubun = "110" Then  'CBC + Diff
            sOrder = "O|" & sOCnt & "|" & gsRackNo & "^" & gsPosNo & "^" & argNo & "^" & "C||"
            sOrder = sOrder & "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\"
            sOrder = sOrder & "^^^^NEUT%\^^^^LYMPH%\^^^^MONO%\^^^^EO%\^^^^BASO%\"
            sOrder = sOrder & "^^^^LYMPH#\^^^^MONO#\^^^^NEUT#\^^^^EO#\^^^^BASO#\" & chrETB
        End If
        
        If sOrder <> "" Then
            sOrder1 = chrSTX & CCur(gCurMsgCnt) & sOrder & CheckSum(CStr(gCurMsgCnt) & sOrder) & chrCR & chrLF
            SetText frmInterface.vasOrder, sOrder1, llRow_Order, 1
            
            sOCnt = sOCnt + 1
            
            gCurMsgCnt = gCurMsgCnt + 1
            If gCurMsgCnt = 8 Then
                gCurMsgCnt = 0
            End If
    
            llRow_Order = llRow_Order + 1
            If llRow_Order > frmInterface.vasOrder.MaxRows Then
                frmInterface.vasOrder.MaxRows = llRow_Order
            End If
        End If
        
        For i = 1 To 3
            sOrder = sOrder & sRetOrder(i)
        Next i
        
        If sOrder <> "" Then
            Select Case sOrderGubun
            Case "110"  'CBC+DIFF
                sOrder = "^^^^RDW-CV\^^^^RDW-SD\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT"
                sOrder = sOrder & "|||||||N||||||||||||||Q" & chrCR & chrETX
            Case "111"  'CBC+DIFF+RET
                sOrder = "^^^^RDW-CV\^^^^RDW-SD\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\"
                sOrder = sOrder & "^^^^RET%\^^^^RET#\^^^^IRF\^^^^LFR\^^^^MFR\^^^^HFR"
                sOrder = sOrder & "|||||||N||||||||||||||Q" & chrCR & chrETX
            End Select
            
            If sOrderGubun = "110" Or sOrderGubun = "111" Then
                sOrder1 = chrSTX & CCur(gCurMsgCnt) & sOrder & CheckSum(CStr(gCurMsgCnt) & sOrder) & chrCR & chrLF
                SetText frmInterface.vasOrder, sOrder1, llRow_Order, 1
                
                sOCnt = sOCnt + 1
                
                gCurMsgCnt = gCurMsgCnt + 1
                If gCurMsgCnt = 8 Then
                    gCurMsgCnt = 0
                End If
        
                llRow_Order = llRow_Order + 1
                If llRow_Order > frmInterface.vasOrder.MaxRows Then
                    frmInterface.vasOrder.MaxRows = llRow_Order
                End If
            End If
        End If
        SetText frmInterface.vasID, "Order", glRow, colState
        
        'Order 전송하기==============================================
        gMsgEnd = "L|1|N" & chrCR & chrETX
        gMsgEnd = Chr(2) & CCur(gCurMsgCnt) & gMsgEnd & CheckSum(CStr(gCurMsgCnt) & gMsgEnd) & chrCR & chrLF
        
        gCurMsgCnt = gCurMsgCnt + 1
        If gCurMsgCnt = 8 Then
            gCurMsgCnt = 1
        End If
        
        llRow_Order = frmInterface.vasOrder.DataRowCnt + 1
        If llRow_Order + 1 > frmInterface.vasOrder.MaxRows Then
            frmInterface.vasOrder.MaxRows = llRow_Order + 1
        End If
        
        SetText frmInterface.vasOrder, gMsgEnd, llRow_Order, 1
        SetText frmInterface.vasOrder, chrEOT, llRow_Order + 1, 1
    Else    '오더가 없다면 CBC+Diff만 검사하도록 강제셋팅함(장비에서 에러 발생하므로)
        llRow_Order = 1
    
        gCurMsgCnt = 1
        'Head
        gHeader = "H|\^&|||XE-2100^00-32^A4349^^^^98313519||||||||E1394-97" & chrCR & chrETX
        gHeader = chrSTX & CCur(gCurMsgCnt) & gHeader & CheckSum(CStr(gCurMsgCnt) & gHeader) & chrCR & chrLF
        
        SetText frmInterface.vasOrder, gHeader, llRow_Order, 1
        
        gCurMsgCnt = gCurMsgCnt + 1
        If gCurMsgCnt = 8 Then
            gCurMsgCnt = 0
        End If
        
        llRow_Order = llRow_Order + 1
        If llRow_Order > frmInterface.vasOrder.MaxRows Then
            frmInterface.vasOrder.MaxRows = llRow_Order
        End If
        
        'Patient
        gPatient = "P|1|||" & Trim(GetText(vasID, argRow, colPID)) & "||||" & sPsex & chrCR & chrETX
        gPatient = chrSTX & CCur(gCurMsgCnt) & gPatient & CheckSum(CStr(gCurMsgCnt) & gPatient) & chrCR & chrLF
        
        SetText frmInterface.vasOrder, gPatient, llRow_Order, 1
        
        gCurMsgCnt = gCurMsgCnt + 1
        If gCurMsgCnt = 8 Then
            gCurMsgCnt = 0
        End If
        
        llRow_Order = llRow_Order + 1
        If llRow_Order > frmInterface.vasOrder.MaxRows Then
            frmInterface.vasOrder.MaxRows = llRow_Order
        End If
        
        'Order
        sOCnt = 1
        
        For i = 1 To 3
            sRetOrder(i) = "0"
        Next i
        
        'CBC + Diff
        sRetOrder(1) = "1"
        sRetOrder(2) = "1"
        sRetOrder(3) = "0"
        
        sOrder = ""
        sOrder1 = ""
        
        For i = 1 To 3
            sOrder = sOrder & sRetOrder(i)
        Next i
        
        sOrderGubun = sOrder
        
        If sOrder <> "" And sOrderGubun = "100" Then      'CBC
            sOrder = "O|" & sOCnt & "|" & gsRackNo & "^" & gsPosNo & "^" & argNo & "^" & "C||"
            sOrder = sOrder & "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\"
            sOrder = sOrder & "^^^^RDW-CV\^^^^RDW-SD\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT"
            sOrder = sOrder & "|||||||N||||||||||||||Q" & chrCR & chrETX
        ElseIf sOrder <> "" And sOrderGubun = "110" Then  'CBC + Diff
            sOrder = "O|" & sOCnt & "|" & gsRackNo & "^" & gsPosNo & "^" & argNo & "^" & "C||"
            sOrder = sOrder & "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\"
            sOrder = sOrder & "^^^^NEUT%\^^^^LYMPH%\^^^^MONO%\^^^^EO%\^^^^BASO%\"
            sOrder = sOrder & "^^^^LYMPH#\^^^^MONO#\^^^^NEUT#\^^^^EO#\^^^^BASO#\" & chrETB
        ElseIf sOrder <> "" And sOrderGubun = "111" Then  'CBC + Diff + RET
            sOrder = "O|" & sOCnt & "|" & gsRackNo & "^" & gsPosNo & "^" & argNo & "^" & "C||"
            sOrder = sOrder & "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\"
            sOrder = sOrder & "^^^^NEUT%\^^^^LYMPH%\^^^^MONO%\^^^^EO%\^^^^BASO%\"
            sOrder = sOrder & "^^^^LYMPH#\^^^^MONO#\^^^^NEUT#\^^^^EO#\^^^^BASO#\" & chrETB
        End If
        
        If sOrder <> "" Then
            sOrder1 = chrSTX & CCur(gCurMsgCnt) & sOrder & CheckSum(CStr(gCurMsgCnt) & sOrder) & chrCR & chrLF
            SetText frmInterface.vasOrder, sOrder1, llRow_Order, 1
            
            sOCnt = sOCnt + 1
            
            gCurMsgCnt = gCurMsgCnt + 1
            If gCurMsgCnt = 8 Then
                gCurMsgCnt = 0
            End If
    
            llRow_Order = llRow_Order + 1
            If llRow_Order > frmInterface.vasOrder.MaxRows Then
                frmInterface.vasOrder.MaxRows = llRow_Order
            End If
        End If
        
        For i = 1 To 3
            sOrder = sOrder & sRetOrder(i)
        Next i
        
        If sOrder <> "" Then
            Select Case sOrderGubun
            Case "110"  'CBC+DIFF
                sOrder = "^^^^RDW-CV\^^^^RDW-SD\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT"
                sOrder = sOrder & "|||||||N||||||||||||||Q" & chrCR & chrETX
            Case "111"  'CBC+DIFF+RET
                sOrder = "^^^^RDW-CV\^^^^RDW-SD\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\"
                sOrder = sOrder & "^^^^RET%\^^^^RET#\^^^^IRF\^^^^LFR\^^^^MFR\^^^^HFR"
                sOrder = sOrder & "|||||||N||||||||||||||Q" & chrCR & chrETX
            End Select
            
            If sOrderGubun = "110" Or sOrderGubun = "111" Then
                sOrder1 = chrSTX & CCur(gCurMsgCnt) & sOrder & CheckSum(CStr(gCurMsgCnt) & sOrder) & chrCR & chrLF
                SetText frmInterface.vasOrder, sOrder1, llRow_Order, 1
                
                sOCnt = sOCnt + 1
                
                gCurMsgCnt = gCurMsgCnt + 1
                If gCurMsgCnt = 8 Then
                    gCurMsgCnt = 0
                End If
        
                llRow_Order = llRow_Order + 1
                If llRow_Order > frmInterface.vasOrder.MaxRows Then
                    frmInterface.vasOrder.MaxRows = llRow_Order
                End If
            End If
        End If
        SetText frmInterface.vasID, "Order", glRow, colState
        
        'Order 전송하기==============================================
        gMsgEnd = "L|1|N" & chrCR & chrETX
        gMsgEnd = Chr(2) & CCur(gCurMsgCnt) & gMsgEnd & CheckSum(CStr(gCurMsgCnt) & gMsgEnd) & chrCR & chrLF
        
        gCurMsgCnt = gCurMsgCnt + 1
        If gCurMsgCnt = 8 Then
            gCurMsgCnt = 1
        End If
        
        llRow_Order = frmInterface.vasOrder.DataRowCnt + 1
        If llRow_Order + 1 > frmInterface.vasOrder.MaxRows Then
            frmInterface.vasOrder.MaxRows = llRow_Order + 1
        End If
        
        SetText frmInterface.vasOrder, gMsgEnd, llRow_Order, 1
        SetText frmInterface.vasOrder, chrEOT, llRow_Order + 1, 1
    End If

End Function

Function Make_Order_테스트(asSpecID As String, asRow As Long) As Integer
    Dim sCnt As String
    
    Dim sOCnt As Long
    Dim sRetOrder As String
    Dim sOrder As String
    
    Dim iRow As Long
    Dim llrow As Long
    Dim llRow_Order As Long
    
    Dim sReceDate As String
    Dim sPID As String
    Dim sPName As String
    Dim sPName_E As String
    Dim sSex As String
    Dim sAge As String
    Dim sEmgFlag As String
    Dim sReceNo As String
    
    Dim lsID As String
    
    Dim sExamCode As String
    Dim sExamName As String
    Dim sEquipCode As String
    Dim sAllExam As String
    Dim rv As String
    
    Dim GluFlag As Integer
    Dim li_cnt
    Dim i
    
    Dim RetVal As String
                    
    sReceDate = Format(CDate(Text_Today), "yyyymmdd")
    
    lsID = asSpecID
        
    ClearSpread vasTemp
    ClearSpread vasOrder
        
    sEmgFlag = "R"
    
    '// Order 찾기
    sOCnt = 0
                        
    llRow_Order = 1

    gCurMsgCnt = 1
    
    'Head
    gHeader = "H|\^&||||||||" & gVersion & "||P|1|" & chrCR & chrETX
    gHeader = chrSTX & CCur(gCurMsgCnt) & gHeader & CheckSum(CStr(gCurMsgCnt) & gHeader) & chrCR & chrLF
    
    SetText frmInterface.vasOrder, gHeader, llRow_Order, 1
    
    gCurMsgCnt = gCurMsgCnt + 1
    If gCurMsgCnt = 8 Then
        gCurMsgCnt = 0
    End If
    
    sOCnt = 1
        
'    '처음 검사 샘플
    res = Get_Sample_Info(asRow)
    
    gReceCode = ""
    
    rv = Online_XML(gXml_S07, lsID)
    
    If res < 1 Or rv < 1 Then
        llRow_Order = llRow_Order + 1
        If llRow_Order > frmInterface.vasOrder.MaxRows Then
            frmInterface.vasOrder.MaxRows = llRow_Order
        End If
        
        'Patient
        gPatient = "P|1||||||||||||||||||" & chrCR & chrETX
        gPatient = chrSTX & CCur(gCurMsgCnt) & gPatient & CheckSum(CStr(gCurMsgCnt) & gPatient) & chrCR & chrLF
        
        SetText frmInterface.vasOrder, gPatient, llRow_Order, 1
        
        gCurMsgCnt = gCurMsgCnt + 1
        If gCurMsgCnt = 8 Then
            gCurMsgCnt = 0
        End If
        
        llRow_Order = llRow_Order + 1
        If llRow_Order > frmInterface.vasOrder.MaxRows Then
            frmInterface.vasOrder.MaxRows = llRow_Order
        End If
    
        sRetOrder = ""
        
        sRetOrder = "O|" & sOCnt & "|" & asSpecID & "||^^^ALL|" & sEmgFlag & "||" & _
                                  "||||N||||||||||||||Q|" & chrCR & chrETX
        If sRetOrder <> "" Then
            sOrder = chrSTX & CCur(gCurMsgCnt) & sRetOrder & CheckSum(CStr(gCurMsgCnt) & sRetOrder) & chrCR & chrLF
            SetText frmInterface.vasOrder, sOrder, llRow_Order, 1
            
            sOCnt = sOCnt + 1
            
            gCurMsgCnt = gCurMsgCnt + 1
            If gCurMsgCnt = 8 Then
                gCurMsgCnt = 0
            End If
    
            llRow_Order = llRow_Order + 1
            If llRow_Order > frmInterface.vasOrder.MaxRows Then
                frmInterface.vasOrder.MaxRows = llRow_Order
            End If
        End If
    
        SetText frmInterface.vasID, 0, glRow, colOrd
        SetText frmInterface.vasID, 0, glRow, colRes
        SetText frmInterface.vasID, "검체확인", glRow, colState
   Else
   
   
        sAllExam = gReceCode
        
        sSex = Trim(GetText(vasID, glRow, colPSex))
        sAge = Trim(GetText(vasID, glRow, colPAge))
        sPID = Trim(GetText(vasID, glRow, colPID))
        sPName = Trim(GetText(vasID, glRow, colPName))
        sReceNo = Trim(GetText(vasID, glRow, colReceno))
        sPName_E = UCase(Conv_Kor_Eng(sPName))
                
        llRow_Order = llRow_Order + 1
        If llRow_Order > frmInterface.vasOrder.MaxRows Then
            frmInterface.vasOrder.MaxRows = llRow_Order
        End If
        
        If Len(sReceNo & "_" & sPName_E) > 20 Then
            sPName_E = Left(sPName_E, 20 - (Len(sReceNo) + 1) - 1) & "-"
        End If
        'sPName_E = ""
        
        'Patient
        gPatient = "P|1||||" & sReceNo & "_" & sPName_E & "||||||||||||||" & chrCR & chrETX
        gPatient = chrSTX & CCur(gCurMsgCnt) & gPatient & CheckSum(CStr(gCurMsgCnt) & gPatient) & chrCR & chrLF
        
        SetText frmInterface.vasOrder, gPatient, llRow_Order, 1
        
        gCurMsgCnt = gCurMsgCnt + 1
        If gCurMsgCnt = 8 Then
            gCurMsgCnt = 0
        End If
        
        llRow_Order = llRow_Order + 1
        If llRow_Order > frmInterface.vasOrder.MaxRows Then
            frmInterface.vasOrder.MaxRows = llRow_Order
        End If
    
        i = 1
        sExamCode = ""
        
        
        SQL = "Select ExamCode, EquipCode, Examname from EquipExam " & _
              "where Equip = '" & gEquip & "' and ExamCode in (" & sAllExam & ")  "
        res = db_select_Vas(gLocal, SQL, vasTemp)
        
            
        
        Do While i <= vasTemp.DataRowCnt
            sEquipCode = ""
            sExamName = ""
            
            sExamCode = Trim(GetText(vasTemp, i, 1))
            sEquipCode = Trim(GetText(vasTemp, i, 2))
            sExamName = Trim(GetText(vasTemp, i, 3))
            
'            Res = GetEquip(sExamCode)
'            If Res > 0 Then
'                sEquipCode = Trim(gReadBuf(0))
'                sExamName = Trim(gReadBuf(1))
'            End If
                
            If sEquipCode <> "" Then
                sCnt = ""
                sCnt = "0"
                sRetOrder = ""
                
                If sCnt = "0" Then
                    If sOCnt = 1 Then
                        sRetOrder = "O|" & sOCnt & "|" & asSpecID & "||^^^" & sEquipCode & "|" & sEmgFlag & "||" & _
                                                  "||||N||||||||||||||Q|" & chrCR & chrETX
                    Else
                        sRetOrder = "O|" & sOCnt & "|" & asSpecID & "||^^^" & sEquipCode & "|" & sEmgFlag & "||" & _
                                                  "||||A||||||||||||||Q|" & chrCR & chrETX
                    End If
                End If
                    
                If sRetOrder <> "" Then
                    sOrder = chrSTX & CCur(gCurMsgCnt) & sRetOrder & CheckSum(CStr(gCurMsgCnt) & sRetOrder) & chrCR & chrLF
                    SetText frmInterface.vasOrder, sOrder, llRow_Order, 1
                    
                    sOCnt = sOCnt + 1
                    
                    gCurMsgCnt = gCurMsgCnt + 1
                    If gCurMsgCnt = 8 Then
                        gCurMsgCnt = 0
                    End If
            
                    llRow_Order = llRow_Order + 1
                    If llRow_Order > frmInterface.vasOrder.MaxRows Then
                        frmInterface.vasOrder.MaxRows = llRow_Order
                    End If
                End If
                
                
                SQL = "Select examcode from pat_res " & vbCrLf & _
                      "where examdate = '" & Trim(sReceDate) & "' " & vbCrLf & _
                      "  and equipno = '" & gEquip & "' " & vbCrLf & _
                      "  and barcode = '" & Trim(lsID) & "' " & vbCrLf & _
                      "  and equipcode = '" & Trim(sEquipCode) & "' " & vbCrLf & _
                      "  and examcode = '" & sExamCode & "' "
                res = db_select_Col(gLocal, SQL)
                If Trim(gReadBuf(0)) = sExamCode Then
                    SQL = "delete from pat_res " & vbCrLf & _
                          "where examdate = '" & Trim(sReceDate) & "' " & vbCrLf & _
                          "  and equipno = '" & gEquip & "' " & vbCrLf & _
                          "  and barcode = '" & Trim(lsID) & "' " & vbCrLf & _
                          "  and equipcode = '" & Trim(sEquipCode) & "' " & vbCrLf & _
                          "  and examcode = '" & sExamCode & "' "
                    res = SendQuery(gLocal, SQL)
                End If
                SQL = " Insert Into pat_res(examdate, equipno, barcode, equipcode,  " & vbCrLf & _
                      " examcode, examname, pid, pname, psex, page, diskno, posno, resdate, sendflag, receno)  " & vbCrLf & _
                      " Values ( '" & Trim(sReceDate) & "', '" & gEquip & "',  '" & Trim(lsID) & "', '" & sEquipCode & "', " & vbCrLf & _
                      " '" & sExamCode & "', '" & sExamName & "', '" & sPID & "', " & vbCrLf & _
                      " '" & sPName & "', '" & sSex & "', " & sAge & ", '', '', " & vbCrLf & _
                      " '" & Trim(GetDateFull) & "', '0', '" & sReceNo & "' ) "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                End If
            End If
    
            i = i + 1
        Loop
    
        SetText frmInterface.vasID, sOCnt - 1, glRow, colOrd
        SetText frmInterface.vasID, 0, glRow, colRes
        SetText frmInterface.vasID, "Order", glRow, colState
        
    End If
        
    gMsgEnd = "L|1|N" & chrCR & chrETX
    gMsgEnd = Chr(2) & CCur(gCurMsgCnt) & gMsgEnd & CheckSum(CStr(gCurMsgCnt) & gMsgEnd) & chrCR & chrLF
    
    gCurMsgCnt = gCurMsgCnt + 1
    If gCurMsgCnt = 8 Then
        gCurMsgCnt = 1
    End If
    
    llRow_Order = frmInterface.vasOrder.DataRowCnt + 1
    If llRow_Order + 1 > frmInterface.vasOrder.MaxRows Then
        frmInterface.vasOrder.MaxRows = llRow_Order + 1
    End If
    
    SetText frmInterface.vasOrder, gMsgEnd, llRow_Order, 1
    SetText frmInterface.vasOrder, chrEOT, llRow_Order + 1, 1
        
End Function

Function Make_Order_Local(asSpecID As String, asRow As Long) As Integer
    Dim sCnt As String
    
    Dim sOCnt As Long
    Dim sRetOrder As String
    Dim sOrder As String
    
    Dim iRow As Long
    Dim llrow As Long
    Dim llRow_Order As Long
    
    Dim sReceDate As String
    Dim sPID As String
    Dim sPName As String
    Dim sPName_E As String
    Dim sSex As String
    Dim sAge As String
    Dim sEmgFlag As String
    Dim sReceNo As String
    
    Dim lsID As String
    
    Dim sExamCode As String
    Dim sExamName As String
    Dim sEquipCode As String
    Dim sAllExam As String
    Dim rv As Integer
    
    Dim GluFlag As Integer
    Dim li_cnt
    Dim i
    
    Dim RetVal As String
                    
    sReceDate = Format(CDate(Text_Today), "yyyymmdd")
    
    lsID = asSpecID
        
    ClearSpread vasTemp
    ClearSpread vasOrder
        
    sEmgFlag = "R"
    
    '// Order 찾기
    sOCnt = 0
                        
    llRow_Order = 1

    gCurMsgCnt = 1
    
    'Head
    gHeader = "H|\^&||||||||" & gVersion & "||P|1|" & chrCR & chrETX
    gHeader = chrSTX & CCur(gCurMsgCnt) & gHeader & CheckSum(CStr(gCurMsgCnt) & gHeader) & chrCR & chrLF
    
    SetText frmInterface.vasOrder, gHeader, llRow_Order, 1
    
    gCurMsgCnt = gCurMsgCnt + 1
    If gCurMsgCnt = 8 Then
        gCurMsgCnt = 0
    End If
    
    sOCnt = 1
        
'    '처음 검사 샘플
    res = Get_Sample_Info_Local(asRow)
    
    '오더확인
    sAllExam = ""
    
    ClearSpread vasTemp
    
    SQL = " Select ExamCode from pat_res " & CR & _
          " Where equipno = '" & gEquip & "' " & CR & _
          " And barcode= '" & Trim(asSpecID) & "' " & CR & _
          " And sendflag = '0' "
    rv = db_select_Vas(gLocal, SQL, vasTemp)
    For i = 1 To vasTemp.DataRowCnt
        If sAllExam = "" Then
            sAllExam = "'" & Trim(GetText(vasTemp, i, 1)) & "'"
        Else
            sAllExam = sAllExam & ",'" & Trim(GetText(vasTemp, i, 1)) & "'"
        End If
    Next i
    
    If res < 1 Or rv < 1 Then
        llRow_Order = llRow_Order + 1
        If llRow_Order > frmInterface.vasOrder.MaxRows Then
            frmInterface.vasOrder.MaxRows = llRow_Order
        End If
        
        'Patient
        gPatient = "P|1||||||||||||||||||" & chrCR & chrETX
        gPatient = chrSTX & CCur(gCurMsgCnt) & gPatient & CheckSum(CStr(gCurMsgCnt) & gPatient) & chrCR & chrLF
        
        SetText frmInterface.vasOrder, gPatient, llRow_Order, 1
        
        gCurMsgCnt = gCurMsgCnt + 1
        If gCurMsgCnt = 8 Then
            gCurMsgCnt = 0
        End If
        
        llRow_Order = llRow_Order + 1
        If llRow_Order > frmInterface.vasOrder.MaxRows Then
            frmInterface.vasOrder.MaxRows = llRow_Order
        End If
    
        sRetOrder = ""
        
        sRetOrder = "O|" & sOCnt & "|" & asSpecID & "||^^^ALL|" & sEmgFlag & "||" & _
                                  "||||N||||||||||||||Q|" & chrCR & chrETX
        If sRetOrder <> "" Then
            sOrder = chrSTX & CCur(gCurMsgCnt) & sRetOrder & CheckSum(CStr(gCurMsgCnt) & sRetOrder) & chrCR & chrLF
            SetText frmInterface.vasOrder, sOrder, llRow_Order, 1
            
            sOCnt = sOCnt + 1
            
            gCurMsgCnt = gCurMsgCnt + 1
            If gCurMsgCnt = 8 Then
                gCurMsgCnt = 0
            End If
    
            llRow_Order = llRow_Order + 1
            If llRow_Order > frmInterface.vasOrder.MaxRows Then
                frmInterface.vasOrder.MaxRows = llRow_Order
            End If
        End If
    
        SetText frmInterface.vasID, 0, glRow, colOrd
        SetText frmInterface.vasID, 0, glRow, colRes
        SetText frmInterface.vasID, "검체확인", glRow, colState
   Else
        'sAllExam = gReceCode
        
        sSex = Trim(GetText(vasID, glRow, colPSex))
        sAge = Trim(GetText(vasID, glRow, colPAge))
        sPID = Trim(GetText(vasID, glRow, colPID))
        sPName = Trim(GetText(vasID, glRow, colPName))
        sReceNo = Trim(GetText(vasID, glRow, colReceno))
        sPName_E = UCase(Conv_Kor_Eng(sPName))
                
        llRow_Order = llRow_Order + 1
        If llRow_Order > frmInterface.vasOrder.MaxRows Then
            frmInterface.vasOrder.MaxRows = llRow_Order
        End If
        
        If Len(sReceNo & "_" & sPName_E) > 20 Then
            sPName_E = Left(sPName_E, 20 - (Len(sReceNo) + 1) - 1) & "-"
        End If
        'sPName_E = ""
        
        'Patient
        gPatient = "P|1||||" & sReceNo & "_" & sPName_E & "||||||||||||||" & chrCR & chrETX
        gPatient = chrSTX & CCur(gCurMsgCnt) & gPatient & CheckSum(CStr(gCurMsgCnt) & gPatient) & chrCR & chrLF
        
        SetText frmInterface.vasOrder, gPatient, llRow_Order, 1
        
        gCurMsgCnt = gCurMsgCnt + 1
        If gCurMsgCnt = 8 Then
            gCurMsgCnt = 0
        End If
        
        llRow_Order = llRow_Order + 1
        If llRow_Order > frmInterface.vasOrder.MaxRows Then
            frmInterface.vasOrder.MaxRows = llRow_Order
        End If
    
        i = 1
        sExamCode = ""
        
        
        SQL = "Select ExamCode, EquipCode, Examname from EquipExam " & _
              "where Equipno = '" & gEquip & "' and ExamCode in (" & sAllExam & ")  "
        res = db_select_Vas(gLocal, SQL, vasTemp)
        
            
        
        Do While i <= vasTemp.DataRowCnt
            sEquipCode = ""
            sExamName = ""
            
            sExamCode = Trim(GetText(vasTemp, i, 1))
            sEquipCode = Trim(GetText(vasTemp, i, 2))
            sExamName = Trim(GetText(vasTemp, i, 3))
            
'            Res = GetEquip(sExamCode)
'            If Res > 0 Then
'                sEquipCode = Trim(gReadBuf(0))
'                sExamName = Trim(gReadBuf(1))
'            End If
                
            If sEquipCode <> "" Then
                sCnt = ""
                sCnt = "0"
                sRetOrder = ""
                
                If sCnt = "0" Then
                    If sOCnt = 1 Then
                        sRetOrder = "O|" & sOCnt & "|" & asSpecID & "||^^^" & sEquipCode & "|" & sEmgFlag & "||" & _
                                                  "||||N||||||||||||||Q|" & chrCR & chrETX
                    Else
                        sRetOrder = "O|" & sOCnt & "|" & asSpecID & "||^^^" & sEquipCode & "|" & sEmgFlag & "||" & _
                                                  "||||A||||||||||||||Q|" & chrCR & chrETX
                    End If
                End If
                    
                If sRetOrder <> "" Then
                    sOrder = chrSTX & CCur(gCurMsgCnt) & sRetOrder & CheckSum(CStr(gCurMsgCnt) & sRetOrder) & chrCR & chrLF
                    SetText frmInterface.vasOrder, sOrder, llRow_Order, 1
                    
                    sOCnt = sOCnt + 1
                    
                    gCurMsgCnt = gCurMsgCnt + 1
                    If gCurMsgCnt = 8 Then
                        gCurMsgCnt = 0
                    End If
            
                    llRow_Order = llRow_Order + 1
                    If llRow_Order > frmInterface.vasOrder.MaxRows Then
                        frmInterface.vasOrder.MaxRows = llRow_Order
                    End If
                End If
                
                
                SQL = "Select examcode from pat_res " & vbCrLf & _
                      "where examdate = '" & Trim(sReceDate) & "' " & vbCrLf & _
                      "  and equipno = '" & gEquip & "' " & vbCrLf & _
                      "  and barcode = '" & Trim(lsID) & "' " & vbCrLf & _
                      "  and equipcode = '" & Trim(sEquipCode) & "' " & vbCrLf & _
                      "  and examcode = '" & sExamCode & "' "
                res = db_select_Col(gLocal, SQL)
                If Trim(gReadBuf(0)) = sExamCode Then
'                    SQL = "delete from pat_res " & vbCrLf & _
'                          "where examdate = '" & Trim(sReceDate) & "' " & vbCrLf & _
'                          "  and equipno = '" & gEquip & "' " & vbCrLf & _
'                          "  and barcode = '" & Trim(lsID) & "' " & vbCrLf & _
'                          "  and equipcode = '" & Trim(sEquipCode) & "' " & vbCrLf & _
'                          "  and examcode = '" & sExamCode & "' "
'                    res = SendQuery(gLocal, SQL)
                End If
'                SQL = " Insert Into pat_res(examdate, equipno, barcode, equipcode,  " & vbCrLf & _
'                      " examcode, examname, pid, pname, psex, page, diskno, posno, resdate, sendflag, receno)  " & vbCrLf & _
'                      " Values ( '" & Trim(sReceDate) & "', '" & gEquip & "',  '" & Trim(lsID) & "', '" & sEquipCode & "', " & vbCrLf & _
'                      " '" & sExamCode & "', '" & sExamName & "', '" & sPID & "', " & vbCrLf & _
'                      " '" & sPName & "', '" & sSex & "', " & sAge & ", '', '', " & vbCrLf & _
'                      " '" & Trim(GetDateFull) & "', 'O', '" & sReceNo & "' ) "
'                res = SendQuery(gLocal, SQL)
'                If res = -1 Then
'                    SaveQuery SQL
'                End If
            End If
    
            i = i + 1
        Loop
    
        SetText frmInterface.vasID, sOCnt - 1, glRow, colOrd
        SetText frmInterface.vasID, 0, glRow, colRes
        SetText frmInterface.vasID, "Order", glRow, colState
        
    End If
        
    gMsgEnd = "L|1|N" & chrCR & chrETX
    gMsgEnd = Chr(2) & CCur(gCurMsgCnt) & gMsgEnd & CheckSum(CStr(gCurMsgCnt) & gMsgEnd) & chrCR & chrLF
    
    gCurMsgCnt = gCurMsgCnt + 1
    If gCurMsgCnt = 8 Then
        gCurMsgCnt = 1
    End If
    
    llRow_Order = frmInterface.vasOrder.DataRowCnt + 1
    If llRow_Order + 1 > frmInterface.vasOrder.MaxRows Then
        frmInterface.vasOrder.MaxRows = llRow_Order + 1
    End If
    
    SetText frmInterface.vasOrder, gMsgEnd, llRow_Order, 1
    SetText frmInterface.vasOrder, chrEOT, llRow_Order + 1, 1
        
End Function

Function Proc_Order_LX(asReq As String) As Integer
    Dim i, j As Integer
    Dim sCnt As String
    Dim iCnt As Integer

    Dim lsData As String

    Dim lsFunc As String
    Dim lsSampleNo As String
    Dim lsDisk As String
    Dim lsPosNO As String
    Dim lsID As String
    Dim lsExamCode As String

    Dim lsSpcCode As String
    
    Dim retOrder As String
    Dim retHead As String
    Dim retMiddle As String
    
    Dim lsEquipCode As String
    Dim iISE As Integer

    Dim lsClass As String

    Dim eDate As String
    Dim llrow As Long

    Dim lsOrder As String
    
    Dim rv As Integer
    Dim lsSex As String
    Dim lsAge As String
    
    Dim vTemp As String
    
    Dim iCCR As Integer
    
    Dim iTIBC As Integer

On Error GoTo ErrHandle

    lsID = asReq
    
    retOrder = ""
    lsOrder = ""
    gOrderMessage = ""
    
    eDate = Format(CDate(Text_Today.Text), "yyyymmdd")

    Proc_Order_LX = -1
    
    llrow = vasID.DataRowCnt + 1
    If llrow > vasID.MaxRows Then
        vasID.MaxRows = llrow + 1
    End If

    If Trim(lsID) = "" Then
        Exit Function
    End If

    vasActiveCell vasID, llrow, colPID

    ClearSpread vasRes, 1, 1
    
    SetForeColor vasID, llrow, llrow, 1, colState, 0, 0, 0

    iCnt = 0

    retOrder = ""
    lsExamCode = ""
                                    
    'rv = Get_Order(lsID)
    
    If rv < 1 Then

        SetText vasID, "없음", llrow, colState

        Exit Function
    Else

        SetText vasID, lsID, llrow, colBarCode
        
        Get_Sample_Info llrow
        

        lsExamCode = gReceCode
        
    End If

    iCnt = 0
    j = 0
    Proc_Order_LX = 0
    If lsExamCode <> "" Then
        ClearSpread vasTemp
        
        SQL = "Select EquipCode, ExamCode, Examname, examflag from EquipExam " & _
              "where EquipNo = '" & gEquip & "' and ExamCode in (" & lsExamCode & ")  "
        res = db_select_Vas(gLocal, SQL, vasTemp)
        For i = 1 To vasTemp.DataRowCnt
            If Trim(GetText(vasTemp, i, 1)) <> "" Then
                If Trim(GetText(vasTemp, i, 4)) = "1" Then
                    lsEquipCode = Trim(GetText(vasTemp, i, 1))
                    
                    If lsEquipCode = "07D" And lsSpcCode = "1CSF" Then
                        lsEquipCode = "07B"
                    End If
                    
                    If Trim(lsOrder) = "" Then
                        lsOrder = SetSpace(lsEquipCode, 4, 2) & ",0"
                    Else
                        lsOrder = lsOrder & "," & SetSpace(lsEquipCode, 4, 2) & ",0"
                    End If
                        
                    iCnt = iCnt + 1
                    
                    If vasRes.MaxRows < iCnt Then
                        vasRes.MaxRows = iCnt
                    End If
                    
                    SetText vasRes, lsEquipCode, iCnt, colEquipCode
                    SetText vasRes, Trim(GetText(vasTemp, i, 2)), iCnt, colExamCode
                    SetText vasRes, Trim(GetText(vasTemp, i, 3)), iCnt, colExamName
                    
                    Save_Local_One_1 llrow, iCnt, "O"
                Else
                
                    iCnt = iCnt + 1
                    
                    If vasRes.MaxRows < iCnt Then
                        vasRes.MaxRows = iCnt
                    End If
                    
                    SetText vasRes, Trim(GetText(vasTemp, i, 1)), iCnt, colEquipCode
                    SetText vasRes, Trim(GetText(vasTemp, i, 2)), iCnt, colExamCode
                    SetText vasRes, Trim(GetText(vasTemp, i, 3)), iCnt, colExamName
                    
                    Save_Local_One_1 llrow, iCnt, "O"
                
                End If
            End If
        Next i
    End If
    
    '=======================================================================
    'SampleType에 가져오는 부분
    SQL = "select examtype from equipexam where examcode in (" & lsExamCode & ")"
    res = db_select_Col(gLocal, SQL)
    
    
'    Select Case lsSpcCode
'    Case "124U", "16h", "8hr"  '24h UR, 16h UR, 8h UR
'        lsClass = "TU"
'    Case "1RUR"    'Random UR
'        lsClass = "UR"
'    Case "1URC"    'UR catheter
'        lsClass = "UR"
'    Case "1CSF"    'CSF
'        lsClass = "SF"
'    Case Else
'        lsClass = "SE"
'    End Select
    
    lsClass = gReadBuf(0)
    

    '=======================================================================
    
    lsSex = Trim(GetText(vasID, llrow, colPSex))
    If lsSex <> "M" And lsSex <> "F" Then
        lsSex = "M"
    End If
    lsAge = Trim(GetText(vasID, llrow, colPAge))
    If Not IsNumeric(lsAge) Then
        lsAge = 5
    End If
    'lsAge = Format(lsAge, "000")
    
    retOrder = ""
    retHead = " 0,801,01,0000,00,0,RO," & lsClass & "," & SetSpace(lsID, 15, 2) & ","
    retHead = retHead & Space(20) & ","
    retHead = retHead & Space(12) & ","
    retHead = retHead & Space(25) & ","
    retHead = retHead & Space(18) & ","
    retMiddle = Space(15) & "," & Space(1) & ","
    retMiddle = retMiddle & SetSpace(lsID, 15, 2) & ","
    retMiddle = retMiddle & Space(18) & ","
    retMiddle = retMiddle & Format(Date, "ddmmyyyy") & ","
    retMiddle = retMiddle & Format(Time, "hhmm") & ","
    retMiddle = retMiddle & Space(20) & ","
    retMiddle = retMiddle & Space(3) & ",5," & Space(8) & ",M,"
    retMiddle = retMiddle & Space(45) & ","
    retMiddle = retMiddle & Space(7) & "," & Space(4) & "," & Space(4) & ","
    retMiddle = retMiddle & Space(2) & "," & Space(6) & ","
    retOrder = retHead & retMiddle & Format(iCnt, "000") & "," & lsOrder
    
    'retOrder = retOrder & "020,09A ,0,43B ,0,06A ,0,05A ,0,41A ,0,44A ,0,07A ,0,08A ,0,11A ,0,35A ,0,30A ,0,31A ,0,03A ,0,12A ,0,10A ,0,01A ,0,01B ,0,04A ,0,02A ,0,50A ,0"
    retOrder = "[" & retOrder & "]"
    
    gOrderMessage = retOrder & CS(retOrder) & Chr(13) & Chr(10)
    
    vasTemp1.MaxRows = vasTemp1.DataRowCnt + 1
    vasTemp1.SetText 1, vasTemp1.DataRowCnt + 1, gOrderMessage
    
    SetText vasID, iCnt, llrow, colOrd
    If iCnt = 0 Then
        SetText vasID, "없음", llrow, colState
        SetForeColor vasID, llrow, llrow, 2, 2, 255, 0, 0
    Else
        SetText vasID, iCnt, llrow, colOrd
        SetText vasID, "오더", llrow, colState
        SetForeColor vasID, llrow, llrow, 2, 2, 0, 0, 0
    End If
    SetFont vasID, llrow, llrow, 1, vasID.MaxCols, 9, False

    vasActiveCell vasID, llrow, 1

        
    Proc_Order_LX = 1

    Exit Function

ErrHandle:
    Proc_Order_LX = -1
    SaveQuery SQL
    Resume Next
End Function

Function Save_Local_One_1(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String

    sExamDate = GetDateFull
    
    sCnt = ""
    If Trim(GetText(vasRes, asRow2, colEquipCode)) = "" Then Exit Function
    
    SQL = "select count(*) from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & gEquip & "' " & vbCrLf & _
          "  and barcode = '" & Trim(GetText(vasID, asRow1, colBarCode)) & "' " & vbCrLf & _
          "  and equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
          "  and examcode= '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
    res = db_select_Col(gLocal, SQL)
    sCnt = Trim(gReadBuf(0))
    If res = -1 Then
        SaveQuery SQL, 1
        Exit Function
    End If
    
    If Not IsNumeric(sCnt) Then
        sCnt = "0"
    End If
    
    If Not IsNumeric(GetText(vasID, asRow1, colPAge)) Then
        SetText vasID, "0", asRow1, colPAge
    End If
'    If Not IsDate(Trim(GetText(vasExam, asRow, colExamDate))) Then
'        SetText vasExam, "1900-01-01", asRow, colExamDate
'    End If

    If sCnt = "0" Then
        SQL = "INSERT INTO pat_res (examdate, equipno, barcode, seqno, diskno, posno, " & _
              "pid, pname, jumin, page, psex, resdate, receno, " & _
              "equipcode, examcode, result, result1, sendflag, examname, " & _
              "refflag, refvalue, panicvalue, recedate ) " & vbCrLf & _
              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
              "'" & Trim(GetText(vasID, asRow1, colBarCode)) & "', '" & Trim(GetText(vasID, asRow1, colSeqNo)) & "'," & _
              "'" & Trim(GetText(vasID, asRow1, colRack)) & "', '" & Trim(GetText(vasID, asRow1, colPos)) & "', " & _
              "'" & Trim(GetText(vasID, asRow1, colPID)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasID, asRow1, colPName)) & "', '" & Trim(GetText(vasID, asRow1, colPJumin)) & "', " & _
              "'" & Trim(GetText(vasID, asRow1, colPAge)) & "', '" & Trim(GetText(vasID, asRow1, colPSex)) & "', " & _
              "'" & sExamDate & "', '" & Trim(GetText(vasID, asRow1, colReceno)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "',  " & _
              "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', '" & Trim(GetText(vasRes, asRow2, colResult1)) & "', '" & asSend & "', '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colRCheck)) & "',  " & _
              "'" & Trim(GetText(vasID, asRow1, colOrd)) & "', '" & Trim(GetText(vasID, asRow1, colRes)) & "', '" & Trim(GetText(vasID, asRow1, colDate)) & "') "
        
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
    Else
        SQL = " Update pat_res Set " & vbCrLf & _
              " diskno = '" & Trim(GetText(vasID, asRow1, colRack)) & "', " & vbCrLf & _
              " posno  = '" & Trim(GetText(vasID, asRow1, colPos)) & "', " & vbCrLf & _
              " result = '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf & _
              " result1 = '" & Trim(GetText(vasRes, asRow2, colResult1)) & "', " & vbCrLf & _
              " refflag = '" & Trim(GetText(vasRes, asRow2, colRCheck)) & "', " & vbCrLf & _
              " refvalue = '" & Trim(GetText(vasID, asRow1, colOrd)) & "', " & vbCrLf & _
              " panicvalue = '" & Trim(GetText(vasID, asRow1, colRes)) & "', " & vbCrLf & _
              " resdate = '" & sExamDate & "' " & vbCrLf & _
              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasID, asRow1, colBarCode)) & "' " & vbCrLf & _
              " And equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "' " & vbCrLf & _
              " And examcode = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "' "
        
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If

    End If
    
End Function

Function Save_Local_One_2(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String

    sExamDate = GetDateFull
    
    sCnt = ""
    'If Trim(GetText(vasOrderTmp, asRow2, colEquipCode)) = "" Then Exit Function
    
    SQL = "select count(*) from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & gEquip & "' " & vbCrLf & _
          "  and barcode = '" & Trim(GetText(vasIDTmp, asRow1, colBarCode)) & "' " & vbCrLf & _
          "  and equipcode = '" & Trim(GetText(vasOrderTmp, asRow2, colEquipCode)) & "'" & vbCrLf & _
          "  and examcode= '" & Trim(GetText(vasOrderTmp, asRow2, colExamCode)) & "'"
    res = db_select_Col(gLocal, SQL)
    sCnt = Trim(gReadBuf(0))
    If res = -1 Then
        SaveQuery SQL, 1
        Exit Function
    End If
    
    If Not IsNumeric(sCnt) Then
        sCnt = "0"
    End If
    
    If Not IsNumeric(GetText(vasIDTmp, asRow1, colPAge)) Then
        SetText vasIDTmp, "0", asRow1, colPAge
    End If
'    If Not IsDate(Trim(GetText(vasExam, asRow, colExamDate))) Then
'        SetText vasExam, "1900-01-01", asRow, colExamDate
'    End If

    If sCnt = "0" Then
        SQL = "INSERT INTO pat_res (examdate, equipno, barcode, seqno, diskno, posno, " & _
              "pid, pname, jumin, page, psex, resdate, receno, " & _
              "equipcode, examcode, result, result1, sendflag, examname, " & _
              "refflag, refvalue, panicvalue, recedate ) " & vbCrLf & _
              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
              "'" & Trim(GetText(vasIDTmp, asRow1, colBarCode)) & "', '" & Trim(GetText(vasIDTmp, asRow1, colSeqNo)) & "'," & _
              "'" & Trim(GetText(vasIDTmp, asRow1, colRack)) & "', '" & Trim(GetText(vasIDTmp, asRow1, colPos)) & "', " & _
              "'" & Trim(GetText(vasIDTmp, asRow1, colPID)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasIDTmp, asRow1, colPName)) & "', '" & Trim(GetText(vasIDTmp, asRow1, colPJumin)) & "', " & _
              "'" & Trim(GetText(vasIDTmp, asRow1, colPAge)) & "', '" & Trim(GetText(vasIDTmp, asRow1, colPSex)) & "', " & _
              "'" & sExamDate & "', '" & Trim(GetText(vasIDTmp, asRow1, colReceno)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasOrderTmp, asRow2, colEquipCode)) & "', '" & Trim(GetText(vasOrderTmp, asRow2, colExamCode)) & "',  " & _
              "'" & Trim(GetText(vasOrderTmp, asRow2, colResult)) & "', '" & Trim(GetText(vasOrderTmp, asRow2, colResult1)) & "', '" & asSend & "', '" & Trim(GetText(vasOrderTmp, asRow2, colExamName)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasOrderTmp, asRow2, colRCheck)) & "',  " & _
              "'" & Trim(GetText(vasIDTmp, asRow1, colOrd)) & "', '" & Trim(GetText(vasIDTmp, asRow1, colRes)) & "', '" & Trim(GetText(vasIDTmp, asRow1, colDate)) & "') "
        
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
    Else
        SQL = " Update pat_res Set " & vbCrLf & _
              " diskno = '" & Trim(GetText(vasIDTmp, asRow1, colRack)) & "', " & vbCrLf & _
              " posno  = '" & Trim(GetText(vasIDTmp, asRow1, colPos)) & "', " & vbCrLf & _
              " result = '" & Trim(GetText(vasOrderTmp, asRow2, colResult)) & "', " & vbCrLf & _
              " result1 = '" & Trim(GetText(vasOrderTmp, asRow2, colResult1)) & "', " & vbCrLf & _
              " refflag = '" & Trim(GetText(vasOrderTmp, asRow2, colRCheck)) & "', " & vbCrLf & _
              " refvalue = '" & Trim(GetText(vasIDTmp, asRow1, colOrd)) & "', " & vbCrLf & _
              " panicvalue = '" & Trim(GetText(vasIDTmp, asRow1, colRes)) & "', " & vbCrLf & _
              " resdate = '" & sExamDate & "' " & vbCrLf & _
              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasIDTmp, asRow1, colBarCode)) & "' " & vbCrLf & _
              " And equipcode = '" & Trim(GetText(vasOrderTmp, asRow2, colEquipCode)) & "' " & vbCrLf & _
              " And examcode = '" & Trim(GetText(vasOrderTmp, asRow2, colExamCode)) & "' "
        
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If

    End If
    
End Function

Function Save_ResMemo(ByVal asRow As Long, asMessage As String)
'메시지 저장하기
    Dim sMessage As String
    
    If asMessage = "" Then
        Exit Function
    End If
    
    sMessage = ""
    
    SQL = " Select message From pat_resmemo " & vbCrLf & _
          " Where examdate = '" & Format(Text_Today.Text, "yyyymmdd") & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(vasID, asRow, colBarCode)) & "' "
    res = db_select_Col(gLocal, SQL)
    
    sMessage = Trim(gReadBuf(0))
    
    If sMessage = "" Then
        SQL = " Insert Into pat_resmemo (examdate, equipno, barcode, message) " & vbCrLf & _
              " VALUES ('" & Format(Text_Today.Text, "yyyymmdd") & "', '" & gEquip & "', " & vbCrLf & _
              "         '" & Trim(GetText(vasID, asRow, colBarCode)) & "', '" & asMessage & "') "
    Else
        'sMessage = sMessage & vbCrLf & asMessage
        sMessage = sMessage & ", " & asMessage

        SQL = " Update pat_resmemo Set " & vbCrLf & _
              " message = '" & Trim(sMessage) & "' " & vbCrLf & _
              " Where examdate = '" & Format(Text_Today.Text, "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasID, asRow, colBarCode)) & "' "
    End If
    
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function

Function Insert_Data(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
'서버의 데이타 베이스에 저장
    Dim sDpcd, sDate1, sSlip, sItem, sOitp, sWkno As String
    Dim sIDNo, sSmyr, sSmsn, sSms1 As String
    Dim tSmsn As String
    Dim lsExamCode, lsResult As String
    Dim lPanicLow, lPanicHigh As Currency
    Dim lDeltaLow, lDeltaHigh, lDeltaMeth, lDeltaGap
    Dim lsPanic, lsDelta As String
    Dim lsPreDate, lsPreResult As String
    Dim lsNState, lsWState As String
    Dim lStdVal
    Dim lTerm As Long
    Dim lsQCChk As String

    Dim iNone, iDP

    Dim sResDate As String
    Dim sRDate As String
    Dim sRTime As String

    Dim lsID As String

    Dim i, j As Long
    Dim lRow As Long
    Dim lsQCOn As String
    
    Dim sResult As String
    Dim sExamCode As String
    Dim sBarcode As String
    Dim sEquipCode As String
    Dim sResStr As String
    Dim sResRow As Long
    Dim sResCnt As String
    Dim sEquipRes As String
    Dim sParam As String
    Dim X As Integer
    
    Dim lsMsg       As String
    Dim lsEqFlag    As String
    
    Dim varTmp      As Variant
    Dim varTmp1     As Variant
    
    Insert_Data = -1

    lsQCOn = ""

    lRow = argSpcRow

    If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Function
    
    If Trim(GetText(vasID, lRow, colPName)) = "" Then Exit Function
    

    lsID = Trim(GetText(vasID, lRow, colBarCode))
    sBarcode = ""
    sEquipCode = ""
    sResult = ""
    sExamCode = ""
    
    If lsID = "" Then Exit Function

    ClearSpread vasTemp
    ClearSpread vasTemp1

    iNone = 0
    iDP = 0

    SQL = "Select equipcode, examcode, examname, result, result1 " & vbCrLf & _
          "from pat_res " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and examdate = '" & Format(Text_Today.Text, "YYYYMMDD") & "' " & vbCrLf & _
          "  and barcode = '" & lsID & "' " & vbCrLf & _
          "  and examcode <> '' " & vbCrLf & _
          "  and result <> '' "
    If asSend = 0 Then
'        SQL = SQL & vbCrLf & _
'          "  and sendflag <> 'C' "
    End If
    
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    If vasTemp.DataRowCnt < 1 Then Exit Function

    lsMsg = ""
    lsEqFlag = ""
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'''    SQL = " Select message From pat_resmemo " & vbCrLf & _
'''          " Where examdate = '" & Format(Text_Today.Text, "yyyymmdd") & "' " & vbCrLf & _
'''          " And equipno = '" & gEquip & "' " & vbCrLf & _
'''          " And barcode = '" & Trim(lsID) & "' "
'''    res = db_select_Col(gLocal, SQL)
'''    If res > 0 Then
'''        lsMsg = "XE2100 : " & Trim(gReadBuf(0))
'''    End If
'''
'''    If Trim(lsMsg) = "" Then
'''        lsMsg = "XE2100"
'''    End If
    
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    Save_Raw_Data lsID & " : 서버 결과 전송 시작"
    Save_Raw_Data lsID & " : 장부 정보 가져오기"

    On Error GoTo ErrHandle
    
    sParam = ""
    
    
    Dim strTestWay As String
    Dim strTestIdName As String
    Dim strResult     As String
    Dim strEGFRResult As String
    
       strTestWay = "- 검사방법 : " & txtTestWay.Text
    strTestIdName = "- 보 고 자 : " & txtTestIdName.Text
    
    For sResRow = 1 To vasTemp.DataRowCnt
        If Trim(GetText(vasTemp, sResRow, 2)) <> "" And Trim(GetText(vasTemp, sResRow, 5)) <> "" Then
            If UCase(Trim(GetText(vasTemp, sResRow, 1))) = "EGFR" Then
                If InStr(Trim(GetText(vasTemp, sResRow, 4)), "Mutation Detected") > 0 Then 'Mutation Detected & vbcr Ex19Del: 7.98
                    varTmp = Split(Trim(GetText(vasTemp, sResRow, 4)), vbNewLine)
                    varTmp1 = Split(varTmp(1), ";")
                    varTmp(1) = Replace(varTmp(1), ": ", "(SQI:")
                    varTmp(1) = Replace(varTmp(1), ";", ") ")
                    varTmp(1) = varTmp(1) & ")"
                    strResult = Replace(getMutationDetected, "XXXXX", varTmp(1))
                    
                    For i = 0 To UBound(varTmp1)
                        'Debug.Print varTmp1(i)
                        Select Case mGetP(varTmp1(i), 1, ":")
                            Case "G719X":       strResult = Replace(strResult, "  18         G719X              -", "  18         G719X              +" & Space(9) & mGetP(varTmp1(i), 2, ":"))
                            Case "Ex19Del":     strResult = Replace(strResult, "  19         Ex19Del            -", "  19         Ex19Del            +" & Space(9) & mGetP(varTmp1(i), 2, ":"))
                            Case "S768I":       strResult = Replace(strResult, "  20         S768I              -", "  20         S768I              +" & Space(9) & mGetP(varTmp1(i), 2, ":"))
                            Case "T790M":       strResult = Replace(strResult, "  20         T790M              -", "  20         T790M              +" & Space(9) & mGetP(varTmp1(i), 2, ":"))
                            Case "Ex20Ins":     strResult = Replace(strResult, "  20         Ex20Ins            -", "  20         Ex20Ins            +" & Space(9) & mGetP(varTmp1(i), 2, ":"))
                            Case "L858R":       strResult = Replace(strResult, "  21         L858R              -", "  21         L858R              +" & Space(9) & mGetP(varTmp1(i), 2, ":"))
                            Case "L861Q":       strResult = Replace(strResult, "  21         L861Q              -", "  21         L861Q              +" & Space(9) & mGetP(varTmp1(i), 2, ":"))
                        End Select
                        
                    Next
                    
                    strResult = Replace(strResult, "검사방법:", strTestWay)
                    strResult = Replace(strResult, "- 보 고 자:", strTestIdName)
                
                ElseIf InStr(Trim(GetText(vasTemp, sResRow, 4)), "Invalid") > 0 Then 'Invalid
                    strResult = getMutationInvalid
                    
                    strResult = Replace(strResult, "검사방법:", strTestWay)
                    strResult = Replace(strResult, "- 보 고 자:", strTestIdName)
                Else
                    strResult = getMutationNotDetected
                    
                    strResult = Replace(strResult, "검사방법:", strTestWay)
                    strResult = Replace(strResult, "- 보 고 자:", strTestIdName)
                End If
                
                'strResult = Trim(GetText(vasTemp, sResRow, 4)) & vbCrLf & vbCrLf & strResult
                
                strEGFRResult = "-검사결과 : "
                If InStr(Trim(GetText(vasTemp, sResRow, 4)), "Mutation") > 0 Then
                    strEGFRResult = strEGFRResult & "Mutation detected"
                Else
                    strEGFRResult = strEGFRResult & "Not detected"
                End If
                
                strResult = strEGFRResult & vbCrLf & vbCrLf & strResult
            Else
                strResult = ""
                strResult = Trim(GetText(vasTemp, sResRow, 4)) & vbCrLf & vbCrLf & strTestWay & vbCrLf & vbCrLf & strTestIdName
            End If
            
            'getMutationDetected
        
            sParam = sParam & "<Table>" & _
                    "<QID><![CDATA[PG_SRL.SLP91_P03]]></QID>" & _
                    "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                    "<USERID><![CDATA[LIA]]></USERID>" & _
                    "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                    "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                    "<P0><![CDATA[" & lsID & "]]></P0>" & _
                    "<P1><![CDATA[" & Trim(GetText(vasTemp, sResRow, 2)) & "]]></P1>" & _
                    "<P2><![CDATA[" & strResult & "]]></P2>" & _
                    "<P3><![CDATA[]]></P3>" & _
                    "<P4><![CDATA[" & gEquip & "]]></P4>" & _
                    "<P5><![CDATA[" & gIFUser & "]]></P5>" & _
                    "<P6><![CDATA[]]></P6>" & _
                    "<P7><![CDATA[]]></P7>" & _
                    "<P8><![CDATA[]]></P8>" & _
                    "<P9><![CDATA[]]></P9>" & _
                    "</Table>"
                    
            SQL = "Update pat_res set sendflag = 'C' " & vbCrLf & _
                  "where equipno = '" & gEquip & "' " & vbCrLf & _
                  "  and barcode = '" & lsID & "' and examcode = '" & Trim(GetText(vasTemp, sResRow, 2)) & "'"
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL
                Exit Function
            End If
        End If
    Next
    
    If sParam = "" Then
        Exit Function
    End If
    
    
    sParam = "<NewDataSet>" & sParam & "</NewDataSet>"
    
    Online_Result_Qry sParam
    
    Insert_Data = 1

    Save_Raw_Data lsID & " : 서버 결과 전송 완료!"

    Exit Function

ErrHandle:
    Save_Raw_Data Err.Number & " : " & Err.Description & vbCrLf & _
                  SQL
    Resume Next
    
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열을 구분자를 이용해 구분해 지정한 위치의 문자열을 구함
'   인수 :
'       1.pText      : 구분자로 구성된 문자열
'       2.pPosiion   : 위치
'       3.pDelimiter : 구분자
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition 인수가 1인 경우 For문 Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function

Sub Var_Clear()
    gsBarCode = ""
    gsPID = ""
    gsRackNo = ""
    gsPosNo = ""
    gsResDateTime = ""
    gsSeqNo = ""
    gsExamCode = ""
    gsExamName = ""
    gsOrder = ""
    gsResult = ""
End Sub

Private Sub Picture1_Click()
    frmUser.Show 0
    
End Sub

Private Sub Text_Today_GotFocus()
    SelectFocus Text_Today
End Sub

Private Sub Text_Today_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCall_Click
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        LX20 Text2
'        'MsgBox CS(Text2)
'
'        'Hitachi747 Mid(Text2.Text, 2)
'        'txtData = ""
'    End If
End Sub

Private Sub Timer1_Timer()
    If Winsock1.State = 7 Then
        lblConnectState = "연결성공"
        lblConnectState.ForeColor = RGB(0, 0, 255)
    Else
        lblConnectState = "연결대기"
        lblConnectState.ForeColor = RGB(255, 0, 0)
    End If
    
    If Winsock2.State = 7 Then
        If lblIFState.Caption = "수신중.." Then
        Else
            lblIFState = "연결성공"
            lblIFState.ForeColor = RGB(0, 0, 255)
        End If
    Else
        lblIFState = "연결대기"
        lblIFState.ForeColor = RGB(255, 0, 0)
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
        cmdSend.SetFocus
    End If
End Sub

Private Sub txtHelp_Change()

End Sub

Private Sub txtID_GotFocus()
    SelectFocus txtID
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


Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        vasSort vasID, colDate, colReceno
'''        If Col = colRack Or Col = colPos Then
'''            vasSort vasID, colRack, colPos
'''        Else
'''            vasSort vasID, Col
'''        End If
    End If
    
'''    If Row < 0 Or Row > vasID.DataRowCnt Then
'''        cmdUp.Enabled = False
'''        cmdDown.Enabled = False
'''    End If
'''
'''    If Row = 1 Then
'''        cmdUp.Enabled = False
'''        cmdDown.Enabled = True
'''    ElseIf Row = vasID.DataRowCnt Then
'''        cmdUp.Enabled = True
'''        cmdDown.Enabled = False
'''    Else
'''        cmdUp.Enabled = True
'''        cmdDown.Enabled = True
'''    End If
    
    vasID_DblClick Col, Row
    
    vasRes_Click Col, 1

End Sub

Private Sub vasID_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lsCnt As String
    Dim lsID As String
    Dim lsDate As String
    Dim lsTime As String
    Dim lsState As String
    
    
    Dim iRow As Long
    
    'cmdCall_Click
    ClearSpread vasRes
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    
    lsID = Trim(GetText(vasID, Row, colBarCode))
    
    StatusBar1.Panels(1).Text = lsID & " : " & Trim(GetText(vasID, Row, colPName)) & " [" & Trim(GetText(vasID, Row, colPName)) & "]"
    
'    If Trim(GetText(vasID, Row, colState)) = "결과" Then
'        lsState = "A"
'    ElseIf Trim(GetText(vasID, Row, colState)) = "완료" Then
'        lsState = "C"
'    End If
    'Local에서 불러오기
    ClearSpread vasRes
    
    If Trim(GetText(vasID, Row, colPJumin)) = "F" Then
        lsTime = Trim(GetText(vasID, Row, colPID))
        If Len(lsTime) = 4 Then
        Else
            lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
        End If
        SQL = "select a.equipcode, min(b.examcode), min(b.examname), a.result, b.seqno, a.resflag, a.result " & vbCrLf & _
              " From qc_res a, equipexam b " & vbCrLf & _
              "where a.equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
              "  and a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
              "  and a.examtime = '" & lsTime & "' " & vbCrLf & _
              "  and a.levelname = '" & lsID & "' " & vbCrLf & _
              "  and b.equipno = a.equipno " & vbCrLf & _
              "  and b.equipcode = a.equipcode " & vbCrLf & _
              "group by a.equipcode, a.result, b.seqno, a.resflag, a.result "
        res = db_select_Vas(gLocal, SQL, vasRes)
    End If
    

    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "Select a.equipcode, a.examcode, b.examname, a.result, b.seqno, a.refflag, a.result1 " & vbCrLf & _
          "from pat_res a, equipexam b " & vbCrLf & _
          "where a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
          "  and a.examcode <> a.equipcode " & vbCrLf & _
          "  and b.equipno = a.equipno " & vbCrLf & _
          "  and b.equipcode = a.equipcode " & vbCrLf & _
          "  and b.examcode = a.examcode"
    res = db_select_Vas(gLocal, SQL, vasRes)
    SQL = "Select a.equipcode, a.examcode, max(b.examname), a.result, b.seqno, a.refflag, a.result1 " & vbCrLf & _
          "from pat_res a, equipexam b " & vbCrLf & _
          "where a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
          "  and a.examcode = a.equipcode " & vbCrLf & _
          "  and b.equipno = a.equipno " & vbCrLf & _
          "  and b.equipcode = a.equipcode " & vbCrLf & _
          "group by a.equipcode, a.examcode, a.result, b.seqno, a.refflag, a.result1 "
    res = db_select_Vas(gLocal, SQL, vasRes, vasRes.DataRowCnt + 1, 1)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    For iRow = 1 To vasRes.DataRowCnt
        If Trim(GetText(vasRes, iRow, colRCheck)) <> "" Then
            SetForeColor vasRes, iRow, iRow, colResult, colResult, 255, 0, 0
        Else
            SetForeColor vasRes, iRow, iRow, colResult, colResult, 0, 0, 0
        End If
    Next iRow
    vasRes.MaxRows = vasRes.DataRowCnt
    'vasSort vasRes, 5, 2
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Long
    Dim lsID As String
    Dim lsTime As String
    
    iRow = vasID.ActiveRow
    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasID.DataRowCnt Then
            Exit Sub
        End If
        
        lsID = Trim(GetText(vasID, iRow, colBarCode))
        
        If Trim(GetText(vasID, iRow, colPJumin)) = "F" Then
            If MsgBox("해당 QC 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
                Exit Sub
            End If
            
            lsTime = Trim(GetText(vasID, iRow, colPID))
            If Len(lsTime) = 4 Then
            Else
                lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
            End If
            
            SQL = "Delete From qc_res a " & vbCrLf & _
                  "where a.equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
                  "  and a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
                  "  and a.examtime = '" & lsTime & "' " & vbCrLf & _
                  "  and a.levelname = '" & lsID & "' "
            res = SendQuery(gLocal, SQL)
                
            Exit Sub
        End If
            
        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
            
        SQL = " Delete From pat_res " & vbCrLf & _
              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & lsID & "' "
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
            
        DeleteRow vasID, iRow, iRow
        ClearSpread vasRes
    End If
End Sub

Private Sub vasID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasID.ActiveRow
        If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Sub
            
        vasID_DblClick colBarCode, lRow
    End If
End Sub

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'Dim iRow As Long
'Dim lsID As String
'
'    If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'        Exit Sub
'    End If
'
'    iRow = Row
'
'    lsID = Trim(GetText(vasID, iRow, colBarcode))
'
'    SQL = " Delete From pat_res " & vbCrLf & _
'          " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'          " And equipno = '" & gEquip & "' " & vbCrLf & _
'          " And barcode = '" & lsID & "' "
'    res = SendQuery(gLocal, SQL)
'
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
'
'    DeleteRow vasID, iRow, iRow
End Sub

Private Sub vasRes_Click(ByVal Col As Long, ByVal Row As Long)
    Dim varTmp      As Variant
    Dim varTmp1     As Variant
    Dim strResult   As String
    Dim i           As Integer
    Dim strTestWay  As String
    Dim strTestIdName   As String
    
    
    strTestWay = "- 검사방법 : " & gSetup.gTestWay
    strTestIdName = "- 보 고 자 : " & gSetup.gTestIdName


    If Trim(GetText(vasRes, Row, 1)) = "EGFR" Then
        If InStr(GetText(vasRes, Row, 4), "Invalid") > 0 Then
            strResult = getMutationInvalid
            
            strResult = Replace(strResult, "검사방법:", strTestWay)
            strResult = Replace(strResult, "- 보 고 자:", strTestIdName)
        
        ElseIf InStr(GetText(vasRes, Row, 4), "No Mutation Detected") > 0 Or InStr(GetText(vasRes, Row, 4), "N/A") > 0 Then
            
            strResult = getMutationNotDetected
            
            strResult = Replace(strResult, "검사방법:", strTestWay)
            strResult = Replace(strResult, "- 보 고 자:", strTestIdName)
            
        Else
            varTmp = Split(Trim(GetText(vasRes, Row, 4)), vbNewLine)
            varTmp1 = Split(varTmp(1), ";")
            varTmp(1) = Replace(varTmp(1), ": ", "(SQI:")
            varTmp(1) = Replace(varTmp(1), ";", ") ")
            varTmp(1) = varTmp(1) & ")"
            strResult = Replace(getMutationDetected, "XXXXX", varTmp(1))
            
            For i = 0 To UBound(varTmp1)
                'Debug.Print varTmp1(i)
                Select Case mGetP(varTmp1(i), 1, ":")
                    Case "G719X":       strResult = Replace(strResult, "  18         G719X              -", "  18         G719X              +" & Space(9) & mGetP(varTmp1(i), 2, ":"))
                    Case "Ex19Del":     strResult = Replace(strResult, "  19         Ex19Del            -", "  19         Ex19Del            +" & Space(9) & mGetP(varTmp1(i), 2, ":"))
                    Case "S768I":       strResult = Replace(strResult, "  20         S768I              -", "  20         S768I              +" & Space(9) & mGetP(varTmp1(i), 2, ":"))
                    Case "T790M":       strResult = Replace(strResult, "  20         T790M              -", "  20         T790M              +" & Space(9) & mGetP(varTmp1(i), 2, ":"))
                    Case "Ex20Ins":     strResult = Replace(strResult, "  20         Ex20Ins            -", "  20         Ex20Ins            +" & Space(9) & mGetP(varTmp1(i), 2, ":"))
                    Case "L858R":       strResult = Replace(strResult, "  21         L858R              -", "  21         L858R              +" & Space(9) & mGetP(varTmp1(i), 2, ":"))
                    Case "L861Q":       strResult = Replace(strResult, "  21         L861Q              -", "  21         L861Q              +" & Space(9) & mGetP(varTmp1(i), 2, ":"))
                End Select
                
            Next
            
            strResult = Replace(strResult, "검사방법:", strTestWay)
            strResult = Replace(strResult, "- 보 고 자:", strTestIdName)
        End If
        
        txteGFRCmnt.Text = strResult
    Else
        txteGFRCmnt.Text = ""
    End If
    
    StatusBar1.Panels(2).Text = Trim(GetText(vasRes, Row, 1))
    
End Sub

Private Sub vasRes_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vasResRow As Long
    Dim vasResCol As Long
    Dim vasIDRow As Long
        
    Dim lCCR, lM_C_ratio, lP_C_ratio As Long
    Dim sCCR, sCrea_S, sCrea_U, sM_ALB_U, sTP_U As String
    
    Dim sResult As String
    Dim sResult1 As String
    
    Dim i As Integer
    
    Dim sTotalVol As String
    
    Dim lsTime As String
    
    vasIDRow = vasID.ActiveRow
    vasResRow = vasRes.ActiveRow
    vasResCol = vasRes.ActiveCol
    
    If KeyCode = vbKeyReturn Then

        If vasResCol = colResult Then
            
            If Trim(GetText(vasRes, vasResRow, colEquipCode)) = "88888" Then
                sTotalVol = Trim(GetText(vasRes, vasResRow, colResult))
                SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
                Save_Local_One_1 vasIDRow, vasResRow, "A"
            
            ElseIf Trim(GetText(vasRes, vasResRow, colEquipCode)) = "99999" Then
                sTotalVol = Trim(GetText(vasRes, vasResRow, colResult))
                SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
                Save_Local_One_1 vasIDRow, vasResRow, "A"
                
                If IsNumeric(sTotalVol) Then
                    lCCR = -1
                    sCCR = ""
                    sCrea_S = ""
                    sCrea_U = ""
                    sM_ALB_U = ""
                    sTP_U = ""
                    
                    i = 1
                    Do While i <= vasRes.DataRowCnt
                        Select Case Trim(GetText(vasRes, i, colExamCode))
                        Case "L3117", "L3101", "L3102", "L3103"  'Microalbumun(24hr),Na,K,Cl
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 1000, "0.00")
                                SetText vasRes, sResult, i, colResult
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                            
                        Case "L3104", "L3106", "L3107", "L3109" 'Ca,Pi,UA,Protein(24hr)
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100, "0.00")
                                SetText vasRes, sResult, i, colResult
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                        Case "L31094", "L31095" 'Protein 16hr, 8hr
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100, "0.00")
                                SetText vasRes, sResult, i, colResult
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                        Case "L31111", "L31112", "L31123", "L3113" 'Creatinie 16hr, 8hr,24hr, BUN(24hr UR)
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            sCrea_U = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, "L31123", i, colExamCode
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100 / 1000, "0.00")
                                SetText vasRes, sResult, i, colResult
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                        Case "L3041", "88888"   'Serum Creatinine
                            sCrea_S = Trim(GetText(vasRes, i, colResult1))
                            
                            'Save_Local_One_1 vasIDRow, i, "A"
                        Case "L31121"   'CCR
                            sCCR = Trim(GetText(vasRes, i, colResult1))
                            lCCR = i
                        Case "L31171"   'Microalbumin(random)
                            sM_ALB_U = Trim(GetText(vasRes, i, colResult1))
                        Case "L31110"  'Creatinine(random)
                            sCrea_U = Trim(GetText(vasRes, i, colResult1))
                        Case "L31090"   'Protein(random)
                            sTP_U = Trim(GetText(vasRes, i, colResult1))
                        Case "L31172"   'Microalbumin / creatinine (random urine)
                            lM_C_ratio = i
                        Case "L31172"   'protein / creatinie (random)
                            lP_C_ratio = i
                        End Select
                        i = i + 1
                    Loop
                    
                    If lCCR > 0 And lCCR <= vasRes.DataRowCnt And IsNumeric(sCrea_U) = True And IsNumeric(sCrea_S) = True Then
                        sResult = Format(CCur(sCrea_U) * CCur(sTotalVol) / 1440 / CCur(sCrea_S), "0.000")
                        SetText vasRes, sResult, lCCR, colResult
                        SetText vasRes, sResult, lCCR, colResult1
                        Save_Local_One_1 vasIDRow, i, "A"
                    End If
                    
'                    If IsNumeric(sM_ALB_U) = True And IsNumeric(sCrea_U) = True Then
'                        sResult = Format(CCur(sM_ALB_U) / CCur(sCrea_U), "0.00") * 100
'                        If lM_C_ratio > 0 And lM_C_ratio <= vasRes.DataRowCnt Then
'                            SetText vasRes, sResult, lM_C_ratio, colResult
'                        Else
'                            i = vasRes.DataRowCnt + 1
'                            If i > vasRes.maxrows Then
'                                vasRes.maxrows = i
'                            End If
'
'                            SetText vasRes, "101", i, colEquipCode
'                            SetText vasRes, "L31172", i, colExamCode
'                            SetText vasRes, "Microalbumin / Urine Creatinine", i, colExamName
'                            SetText vasRes, sResult, i, colResult
'                            SetText vasRes, sResult, i, colResult1
'                        End If
'
'                        Save_Local_One_1 vasIDRow, i, "A"
'                    End If
'
'                    If IsNumeric(sTP_U) = True And IsNumeric(sCrea_U) = True Then
'                        sResult = Format(CCur(sTP_U) / CCur(sCrea_U), "0.00") * 1000
'                        If lP_C_ratio > 0 And lP_C_ratio <= vasRes.DataRowCnt Then
'                            SetText vasRes, sResult, lM_C_ratio, colResult
'                        Else
'                            i = vasRes.DataRowCnt + 1
'                            If i > vasRes.maxrows Then
'                                vasRes.maxrows = i
'                            End If
'
'                            SetText vasRes, "102", i, colEquipCode
'                            SetText vasRes, "L31201", i, colExamCode
'                            SetText vasRes, "Urine Protein / Urine Creatinine", i, colExamName
'                            SetText vasRes, sResult, i, colResult
'                            SetText vasRes, sResult, i, colResult1
'                        End If
'
'                        Save_Local_One_1 vasIDRow, i, "A"
'                    End If
                End If
            Else
                
                If Trim(GetText(vasRes, vasIDRow, colPJumin)) = "F" Then
                
                    If MsgBox("해당 QC의 " & Trim(GetText(vasRes, vasResRow, colExamName)) & " 결과를 수정 하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
                        Exit Sub
                    End If
                
                    lsTime = Trim(GetText(vasID, vasIDRow, colPID))
                    If Len(lsTime) = 4 Then
                    Else
                        lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
                    End If
                    
                    SQL = "update qc_res set result = '" & sResult & "' " & vbCrLf & _
                          "where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
                          "  and examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
                          "  and examtime = '" & lsTime & "' " & vbCrLf & _
                          "  and levelname = '" & Trim(GetText(vasID, vasIDRow, colBarCode)) & "' " & vbCrLf & _
                          "  and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' "
                    res = SendQuery(gLocal, SQL)
                
                    Exit Sub
                Else
                
                
                    sResult = Trim(GetText(vasRes, vasResRow, colResult))
                    If MsgBox("저장하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "주의!!!  확인!!!") = vbYes Then
                        sResult = Trim(GetText(vasRes, vasResRow, colResult))
                        
                        SQL = " update pat_res set " & vbCrLf & _
                              " Result = '" & sResult & "' " & vbCrLf & _
                              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                              " And equipno = '" & gEquip & "' " & vbCrLf & _
                              " And barcode = '" & Trim(GetText(vasID, vasIDRow, colBarCode)) & "' " & vbCrLf & _
                              " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' "
                        res = SendQuery(gLocal, SQL)
                        
                        If res = -1 Then
                            SaveQuery SQL
                            Exit Sub
                        End If
        
                        'SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
        
                    End If
                End If
            End If
            
            
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If Trim(GetText(vasID, vasIDRow, colPJumin)) = "F" Then
        
            If MsgBox("해당 QC의 " & Trim(GetText(vasRes, vasResRow, colExamName)) & " 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
                Exit Sub
            End If
        
            lsTime = Trim(GetText(vasID, vasIDRow, colPID))
            If Len(lsTime) = 4 Then
            Else
                lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
            End If
            
            SQL = "Delete From qc_res a " & vbCrLf & _
                  "where a.equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
                  "  and a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
                  "  and a.examtime = '" & lsTime & "' " & vbCrLf & _
                  "  and a.levelname = '" & Trim(GetText(vasID, vasIDRow, colBarCode)) & "' " & vbCrLf & _
                  " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' "
            res = SendQuery(gLocal, SQL)
        
            Exit Sub
        End If
        If MsgBox("해당 환자의 " & Trim(GetText(vasRes, vasResRow, colExamName)) & " 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        SQL = " Delete From pat_res " & vbCrLf & _
              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasID, vasIDRow, colBarCode)) & "' " & vbCrLf & _
              " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' " & vbCrLf & _
              " and examcode =  '" & Trim(GetText(vasRes, vasResRow, colExamCode)) & "' "
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
            
        DeleteRow vasRes, vasResRow, vasResRow
    
    End If
End Sub

Function Save_Local_QC(asExamDate As String, asBarcode As String, asExamCode As String, asRes1 As String, asRes2 As String)
    Dim sResDateTime As String
    Dim sControl As String
    Dim sLotNo As String
    
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sRefFlag As String
    
    Dim sCnt As String
    
    sResDateTime = Format(CDate(asExamDate), "yyyymmdd hhnnss")
    'sControl = Trim(Left(asBarcode, 2))
    'sLotNo = Trim(Mid(asBarcode, 3))
    sControl = asBarcode
    sRefFlag = ""
    
    SQL = "Select t_mean, t_sd from qcexam " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and validstart >= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and valiend <= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and levelname = '" & sControl & "' " & vbCrLf & _
          "  and equipcode = '" & asExamCode & "' "
    res = db_select_Col(gLocal, SQL)
    If res > 0 Then
        If IsNumeric(gReadBuf(0)) And IsNumeric(gReadBuf(1)) Then
            sRefLow = CCur(gReadBuf(0)) - CCur(gReadBuf(1))
            sRefHigh = CCur(gReadBuf(0)) + CCur(gReadBuf(1))
            If CCur(sRefHigh) < CCur(asRes2) Then
                sRefFlag = "H"
            End If
            If CCur(sRefLow) > CCur(asRes2) Then
                sRefFlag = "L"
            End If
        End If
    End If
    
    sCnt = ""
    SQL = "Select count(*) from qc_res " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and examtime = '" & Mid(sResDateTime, 10, 6) & "' " & vbCrLf & _
          "  and levelname = '" & sControl & "' " & vbCrLf & _
          "  and equipcode = '" & asExamCode & "' "
    res = db_select_Var(gLocal, SQL, sCnt)
    If res <= 0 Then
        SaveQuery SQL
        db_RollBack gLocal
        Exit Function
    End If
    res = db_select_Var(gLocal, SQL, sCnt)
    If res <= 0 Then
        SaveQuery SQL
        Exit Function
    End If
    If Not IsNumeric(sCnt) Then sCnt = "0"
    
    If CInt(sCnt) > 0 Then
        SQL = "delete from qc_res " & vbCrLf & _
              "where equipno = '" & gEquip & "' " & vbCrLf & _
              "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
              "  and examtime = '" & Mid(sResDateTime, 9, 4) & "' " & vbCrLf & _
              "  and levelname = '" & sControl & "' " & vbCrLf & _
              "  and equipcode = '" & asExamCode & "' "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            'db_RollBack gLocal
            SaveQuery SQL
            Exit Function
        End If
    End If
    SQL = "Insert into qc_res (equipno, examdate, examtime, levelname, equipcode, sresult, result, resflag, remark, examuid, lotno) " & vbCrLf & _
          "values ('" & gEquip & "', '" & Left(sResDateTime, 8) & "', '" & Mid(sResDateTime, 10, 4) & "', '" & sControl & "', '" & asExamCode & "', '" & asRes1 & "', '" & asRes2 & "', '" & sRefFlag & "','','', '" & sLotNo & "') "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        'db_RollBack gLocal
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Public Sub WinSock_Listen(argWinSock As Winsock)
    Dim sWinSockPort As String
    
    
    sWinSockPort = gSetup.gPort
    
    
    If sWinSockPort = "0" Or IsNumeric(sWinSockPort) = False Then
        Exit Sub
    End If
    
    If argWinSock.State <> sckClosed Then
        argWinSock.Close
    End If
    
    argWinSock.LocalPort = sWinSockPort
    argWinSock.Listen
    

    
End Sub

Private Sub Winsock1_Close()
    
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    Winsock1.LocalPort = gSetup.gPort
    Winsock1.Listen
    
    
'''    lblConnect1.Caption = "연결 대기중..."
    
End Sub


Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    
    Winsock1.Accept requestID
'''    lblConnect1.Caption = "연결[" & requestID & "]" & Winsock1.RemoteHostIP
End Sub

Public Function HL7_Ack(argMSH As String) As String
'''    Dim strMSH As String
'''    Dim strACK As String
'''    Dim strDateTime As String
'''    Dim strSplit() As String
'''    Dim strSigNum As String
'''
'''    Dim i As Integer
'''    Dim j As Integer
'''
'''    strMSH = argMSH
'''    strSplit = Split(strMSH, "|")
'''
''''''    MSH|^~\&|cobas 8000||host||20130104114005||OUL^R22^REAL|31777||2.5||||AA||UNICODE UTF-8|
'''    strDateTime = Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
'''
'''    strACK = Chr(11)
'''    strACK = strACK & "MSH|^~\&|" & Trim(strSplit(5)) & "||" & Trim(strSplit(3)) & "||" & strDateTime & "||ACK|" & CStr(gMSGSeq) & "||" & Trim(strSplit(3)) & "||||AA||" & Trim(strSplit(18)) & "|" & vbCr
'''    strACK = strACK & "MSA|AA|" & Trim(strSplit(10)) & "||" & vbCr
'''    strACK = strACK & Chr(28) & vbCr
'''    gMSGSeq = gMSGSeq + 1

End Function

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim sTmp As String
    Dim i As Integer
    Dim strResFlag As String
    
    
    Winsock1.GetData sTmp
    
    txtBuff.Text = txtBuff.Text & sTmp
    
    Save_Raw_Data "[RX " & Format(Time, "hh:mm:ss") & "]" & sTmp
    
    If InStr(1, sTmp, chrENQ) > 0 Then
        Save_Raw_Data "[RX " & Format(Time, "hh:mm:ss") & "]" & chrACK
        Winsock1.SendData chrACK
        txtBuff.Text = sTmp
    End If
    
    If InStr(1, sTmp, chrLF) > 0 Then
        Save_Raw_Data "[RX " & Format(Time, "hh:mm:ss") & "]" & chrACK
        Winsock1.SendData chrACK
    End If
    
    If InStr(1, sTmp, chrEOT) > 0 Then
        strResFlag = Cobas4800All(txtBuff.Text)
        gOrderCnt = 1
    End If
    
    If InStr(1, sTmp, chrACK) > 0 Then
        If gOrderMessage <> "" Then
            SendOrder
            
        Else
            Winsock1.SendData chrEOT
        End If
    End If
        
End Sub


Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'''    lblConnect1.Caption = "[Error]" & Number & " : " & Description
End Sub

Public Sub WinSock_Listen2(argWinSock As Winsock)
    Dim sWinSockPort As String
    
    
    sWinSockPort = gSetup.gPort2
    
    
    If sWinSockPort = "0" Or IsNumeric(sWinSockPort) = False Then
        Exit Sub
    End If
    
    If argWinSock.State <> sckClosed Then
        argWinSock.Close
    End If
    
    argWinSock.LocalPort = sWinSockPort
    argWinSock.Listen
    

    
End Sub

Private Sub Winsock2_Close()
    
    If Winsock2.State <> sckClosed Then
        Winsock2.Close
    End If
    Winsock2.LocalPort = gSetup.gPort2
    Winsock2.Listen
    
    
'''    lblConnect1.Caption = "연결 대기중..."
    
End Sub


Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
    If Winsock2.State <> sckClosed Then
        Winsock2.Close
    End If
    
    Winsock2.Accept requestID
'''    lblConnect1.Caption = "연결[" & requestID & "]" & Winsock1.RemoteHostIP
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)

    Dim sTmp As String
    Dim i As Integer
    Dim strResFlag As String
    
    
    Winsock2.GetData sTmp
    
    txtXMLRes.Text = txtXMLRes.Text & sTmp
    
    
    If InStr(1, sTmp, chrENQ) > 0 Then
        txtXMLRes.Text = Mid(sTmp, InStr(1, sTmp, chrENQ) + 1)
        lblIFState.Caption = "수신중.."
    ElseIf InStr(1, sTmp, chrEOT) > 0 Then
        If InStr(txtXMLRes.Text, "HPV_Result_Data") > 0 Then
            Proc_Auto_res txtXMLRes.Text
        ElseIf InStr(txtXMLRes.Text, "eGFR_Result_Data") > 0 Then
            Proc_Auto_res_eGFR txtXMLRes.Text
        End If
        
        txtXMLRes.Text = ""
        lblIFState.Caption = "연결완료"

    End If
    
End Sub


Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'''    lblConnect1.Caption = "[Error]" & Number & " : " & Description
End Sub


Private Function Cobas4800All(argData As String, Optional argTest As Integer = 0) As String

    Dim strData As String
    Dim strSub() As String
    Dim strBarcode() As String
    Dim sID As String
    
    Dim i As Integer
    Dim j As Integer
    Dim strSubData As String
    Dim strRes As String
    
    Cobas4800All = "R"
    
    strData = argData
    
    strData = Replace(strData, chrENQ, "")
    strData = Replace(strData, chrEOT, "")
    
    i = InStr(1, strData, chrSTX)
    While i > 0
        strData = Mid(strData, 1, i - 1) & Mid(strData, i + 2)
        i = InStr(1, strData, chrSTX)
    Wend
    
    i = InStr(1, strData, chrLF)
    While i > 0
        strData = Mid(strData, 1, i - 5) & Mid(strData, i + 1)
        i = InStr(1, strData, chrLF)
    Wend
    
    strSub = Split(strData, chrCR)
    
    ClearSpread vasWork
    
    For i = 0 To UBound(strSub)
        strSubData = Trim(strSub(i)) & "|"
        Dim strDel() As String
        strDel = Split(strSubData, "|")
        If UBound(strDel) > 0 Then
            If strDel(0) = "Q" Then
                Cobas4800All = "Q"
                j = vasWork.DataRowCnt + 1
                If j > vasWork.MaxRows Then
                    vasWork.MaxRows = j
                End If
                
                sID = Replace(strDel(2), "^", "")
                SetText vasWork, sID, j, 1
                
            ElseIf strDel(0) = "H" Then
                ClearSpread vasWork
            End If
        End If
    Next
    
'''    Save_Raw_Data
    If Cobas4800All = "Q" Then
        gOrderMessage = ""
        strRes = Proc_Order_New()
        
        If gOrderMessage <> "" Then
            Save_Raw_Data "[TX]" & chrENQ
            Winsock1.SendData chrENQ
        End If
    ElseIf Cobas4800All = "R" Then
'''        res = Proc_Result()
    End If
    
End Function


Function Proc_Order_New() As String
    Dim strOrder As String
    Dim i As Integer
    Dim j As Integer
    Dim strBarcode As String
    Dim iRow As Integer
    Dim strDateTime As String
'''    Dim iRow As Integer
    Dim iHeadCnt As Integer
    Dim strEquipCode As String
    
    
'''H|\^&|||LIS^00000000-0000-0000-0000-000000001114^UserID^0.0.0.0^1394.LIS2|||||cobas4800|TSDWN^REAL|P|1|YYYYMMDDhhmmss
'''P|1
'''O|1|BarcodeID1||^^^LISOrderCode1^^RunType|||YYYYMMDDhhmmss||||ActionCode|||YYYYMMDDhhmmss|SpecimenType^P|UserID|||||||||ReportType
'''P|2
'''O|1|BarcodeID1||^^^LISOrderCode2^^RunType|||YYYYMMDDhhmmss||||ActionCode|||YYYYMMDDhhmmss|SpecimenType^P|UserID|||||||||ReportType
'''L|1|N

    
    'ASTM 신호 만들기
    
    ClearSpread vasASTM
    
'    strEquipCode = "02HPVGEN^^FULL"
'    strEquipCode = "08EGFR^^AnD"
    
    strDateTime = Format(Now, "yyyymmddhhmmss")
    gOrderMessage = ""
    

    strBarcode = Trim(GetText(vasWork, 1, 1))
    iRow = -1
    
    For j = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, j, colBarCode)) = strBarcode Then
            iRow = j
            Exit For
        End If
    Next
    
    If iRow = -1 Then
        iRow = vasID.DataRowCnt + 1
        If iRow > vasID.MaxRows Then
            vasID.MaxRows = iRow
        End If
    End If
    
    SetText vasID, strBarcode, iRow, colBarCode
    SetText vasID, "Order", iRow, colState
    
    If Trim(GetText(vasID, iRow, colPName)) = "" Then
        Get_Sample_Info iRow
    End If
    
'    gPat_Info_Select.TST_CD

'    strEquipCode = "02HPVGEN^^FULL"
'    strEquipCode = "08EGFR^^AnD"

    For j = 1 To UBound(gArrEquip())
        If gPat_Info_Select.TST_CD = gArrEquip(j, 3) Then
            strEquipCode = gArrEquip(j, 2)
            Exit For
        End If
    Next j
    
    If strEquipCode = "HPV" Then
        strEquipCode = "02HPVGEN^^FULL"
    Else
        strEquipCode = "08EGFR^^AnD"
    End If
    
    strOrder = ""
    strOrder = strOrder & "H|\^&|||LIS^" & Format(Time, "HHMMSS") & "^LIS^47.11^1394.LIS2|||||cobas 4800|TSDWN^REAL|P|1|" & strDateTime & vbCr
    strOrder = strOrder & "P|1" & vbCr
    'strOrder = strOrder & "O|1|" & Trim(GetText(vasWork, 1, 1)) & "^^||^^^" & strEquipCode & "^^FULL|||" & strDateTime & "||||N|||" & strDateTime & "|PCYT^P|admin|||||||||O" & vbCr
    strOrder = strOrder & "O|1|" & Trim(GetText(vasWork, 1, 1)) & "^^||^^^" & strEquipCode & "|||" & strDateTime & "||||N|||" & strDateTime & "|PCYT^P|admin|||||||||O" & vbCr
    strOrder = strOrder & "L|1|N" & vbCr
    
    gOrderMessage = strOrder
    Proc_Order_New = gOrderMessage

End Function

'''Public Function MakeOrderRecode(argCode As String, asEM As String, asRackPos As String, asKind As String, ByVal asRow As Long) As Integer
'''Dim i, j As Integer
'''Dim iCnt As Integer
'''
'''Dim retOrder As String
'''Dim lsID As String
'''Dim lsEquipCode As String
'''Dim lsExamCode As String
'''Dim lsExamName As String
'''Dim lsSeqNo As String
'''Dim lsSampleType As String
'''
'''Dim iISE As Integer
'''Dim iISE_r As String
'''
'''Dim eDate As String
'''
'''Dim sCnt As String
'''Dim sRv As String
'''Dim lsReceCode As String
'''
'''
'''    ClearSpread vasRes
'''
'''    iCnt = 0
'''    MakeOrderRecode = -1
'''
'''    gOrd.OrderCnt = 0
'''    gOrd.OrderText = ""
'''    gOrd.ExamCode = ""
'''    gOrd.SampleType2 = ""
'''
'''    retOrder = ""
'''    ClearSpread vasTemp
'''
'''    If argCode = "" Then
'''        MakeOrderRecode = -1
'''        Exit Function
'''    End If
'''
'''    eDate = Trim(Text_Today.Text)
'''    'argCode = Trim(argCode)
'''    lsID = Trim(argCode)
'''
''''    '처음 검사 샘플
'''
''''''    SQL = "SELECT  b.wd_code ,max(b.wd_date) ,'W' ,a.pe_sujinja , a.pe_jumin  " & vbCrLf & _
''''''          "From person a, wchdat b " & vbCrLf & _
''''''          "WHERE a.pe_chart = '" & lsID & "' " & vbCrLf & _
''''''          "  and a.pe_chart = b.wd_chart " & vbCrLf & _
''''''          "  and b.wd_code in (" & gAllExam & ") " & vbCrLf & _
''''''          "  and b.wd_end_dep = '2' and wd_cancel = '0' " & vbCrLf & _
''''''          "group by b.wd_code ,b.wd_date ,a.pe_sujinja , a.pe_jumin "
''''''
''''''    SQL = SQL & vbCrLf & "union SELECT  b.id_code ,max(b.id_date) ,'I' ,a.pe_sujinja , a.pe_jumin  " & vbCrLf & _
''''''          "From person a, ichdat b " & vbCrLf & _
''''''          "WHERE a.pe_chart = '" & lsID & "' " & vbCrLf & _
''''''          "  and a.pe_chart = b.id_chart " & vbCrLf & _
''''''          "  and b.id_code in (" & gAllExam & ") " & vbCrLf & _
''''''          "  and b.id_end_dep = '2' and id_cancel = '0' " & vbCrLf & _
''''''          "group by b.id_code ,b.id_date ,a.pe_sujinja , a.pe_jumin "
'''
'''    Clear_XML_Exam
'''    sRv = Online_XML(gXml_S07, Trim(lsID))
'''    lsReceCode = ""
'''
'''
'''
'''
'''    For i = 0 To UBound(gExam_Select)
'''
'''        If lsReceCode = "" Then
'''            lsReceCode = "'" & Trim(gExam_Select(i).TST_CD) & "'"
'''        Else
'''            lsReceCode = lsReceCode & ",'" & Trim(gExam_Select(i).TST_CD) & "'"
'''        End If
'''
'''    Next i
'''
'''    If lsReceCode = "" Then
'''        lsReceCode = "''"
'''    End If
'''
'''    ClearSpread vasTemp
'''
'''    SQL = "select examcode, equipcode, examname, seqno from equipexam where equipno = '" & gEquip & "' and examcode in (" & lsReceCode & ")"
'''    res = db_select_Vas(gLocal, SQL, vasTemp)
''''''    res = db_select_Vas(gServer, SQL, vasTemp)
'''    If res = -1 Then
'''        SaveQuery SQL
'''        'Exit Function
'''    End If
'''
'''
'''    iISE = -1
'''    If vasTemp.DataRowCnt > 0 Then
'''
'''        retOrder = ""
'''        ClearSpread vasRes
'''
'''        For i = 1 To vasTemp.DataRowCnt
'''
'''
'''            lsExamCode = Trim(GetText(vasTemp, i, 1))
'''            lsEquipCode = Trim(GetText(vasTemp, i, 2))
'''            lsExamName = Trim(GetText(vasTemp, i, 3))
'''            lsSeqNo = Trim(GetText(vasTemp, i, 4))
'''
'''            'Serum 만 검사.
'''            lsSampleType = gOrd.SampleType1
'''
'''            retOrder = retOrder & "^^^" & lsEquipCode & "/\"
'''
'''            If vasRes.MaxRows < i Then vasRes.MaxRows = i
'''
'''            SetText vasRes, lsEquipCode, i, colEquipCode
'''            SetText vasRes, lsExamCode, i, colExamCode
'''            SetText vasRes, lsExamName, i, colExamName
'''
'''            Save_Local_One_1 asRow, i, "A"
'''
'''        Next i
'''    Else
'''
'''        MakeOrderRecode = 0
'''    End If
'''    If Len(retOrder) > 0 Then
'''        gOrd.OrderText = Mid(retOrder, 1, Len(retOrder) - 1)
'''    Else
'''        gOrd.OrderText = ""
'''    End If
'''
'''    gOrd.OrderCnt = i
'''    gOrd.ExamCode = lsExamCode
'''
'''    MakeOrderRecode = 1
'''
'''End Function



VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F4EDE462-638B-421B-9D98-84223E09E2A0}#1.0#0"; "ACKMAXMAT0100.ocx"
Begin VB.Form frmInterface 
   Caption         =   "인터페이스 화면"
   ClientHeight    =   10620
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15180
   ClipControls    =   0   'False
   Icon            =   "INTERFACE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10620
   ScaleWidth      =   15180
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   255
      TabIndex        =   41
      Top             =   8115
      Width           =   2490
      Begin Threed.SSCommand cmdClear1 
         Height          =   465
         Left            =   90
         TabIndex        =   42
         Top             =   630
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   820
         _StockProps     =   78
         Caption         =   "블럭리스트 화면삭제"
         ForeColor       =   16576
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "INTERFACE.frx":08CA
      End
      Begin Threed.SSCommand cmdClear2 
         Height          =   450
         Left            =   90
         TabIndex        =   43
         Top             =   1110
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   794
         _StockProps     =   78
         Caption         =   "전체리스트 화면삭제"
         ForeColor       =   8388736
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "INTERFACE.frx":08E6
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   465
         Left            =   90
         TabIndex        =   44
         Top             =   1590
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   820
         _StockProps     =   78
         Caption         =   "인터페이스 종료"
         ForeColor       =   128
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "INTERFACE.frx":0902
      End
      Begin Threed.SSCommand cmdInitial 
         Height          =   345
         Left            =   210
         TabIndex        =   45
         Top             =   880
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "Comm. Initialize"
         ForeColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdSendOrder 
         Height          =   465
         Left            =   90
         TabIndex        =   49
         Top             =   135
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   820
         _StockProps     =   78
         Caption         =   "오더 전송"
         ForeColor       =   32768
      End
   End
   Begin Threed.SSFrame fraBarCd 
      Height          =   945
      Left            =   0
      TabIndex        =   33
      Top             =   7275
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   1667
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtBarCd 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  '사용 못함
         Left            =   720
         TabIndex        =   0
         Top             =   495
         Width           =   1775
      End
      Begin Threed.SSPanel pnlJNo 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   34
         Top             =   195
         Width           =   2145
         _Version        =   65536
         _ExtentX        =   3784
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "검체번호 입력"
         ForeColor       =   12648447
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodColor      =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   35
         Top             =   525
         Width           =   330
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1320
      Left            =   15270
      TabIndex        =   47
      Top             =   5235
      Width           =   2505
      Begin Threed.SSCommand cmdWorkList 
         Height          =   465
         Left            =   90
         TabIndex        =   48
         Top             =   210
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   820
         _StockProps     =   78
         Caption         =   "WorkList 조회"
         ForeColor       =   8388608
         BevelWidth      =   1
         RoundedCorners  =   0   'False
      End
   End
   Begin ACKMAXMAT01.MAXMAT MAXMAT1 
      Height          =   2190
      Left            =   6255
      TabIndex        =   46
      Top             =   6915
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   3863
   End
   Begin FPSpread.vaSpread spdIntList 
      Height          =   5775
      Left            =   90
      TabIndex        =   38
      Top             =   60
      Width           =   15060
      _Version        =   393216
      _ExtentX        =   26564
      _ExtentY        =   10186
      _StockProps     =   64
      ColHeaderDisplay=   0
      ColsFrozen      =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   236
      MaxRows         =   25
      NoBeep          =   -1  'True
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "INTERFACE.frx":091E
      UserResize      =   0
      VisibleCols     =   236
      VisibleRows     =   25
      TextTip         =   1
   End
   Begin VB.ListBox listTest 
      Height          =   240
      Left            =   1740
      TabIndex        =   17
      Top             =   10500
      Visible         =   0   'False
      Width           =   6105
   End
   Begin VB.ListBox listNoOrd 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      Height          =   2550
      ItemData        =   "INTERFACE.frx":44F0
      Left            =   3030
      List            =   "INTERFACE.frx":44F2
      TabIndex        =   25
      Top             =   6750
      Width           =   4965
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   11670
      TabIndex        =   14
      Top             =   10440
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   11000
      Left            =   2670
      Top             =   1140
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3060
      Top             =   1140
   End
   Begin Threed.SSPanel SSPanel10 
      Height          =   285
      Left            =   3030
      TabIndex        =   15
      Top             =   5970
      Width           =   1965
      _Version        =   65536
      _ExtentX        =   3466
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "Interface Result ....."
      ForeColor       =   12648447
      BackColor       =   16512
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodColor      =   0
      Alignment       =   8
   End
   Begin Threed.SSPanel lblCSelList 
      Height          =   285
      Left            =   5010
      TabIndex        =   16
      Top             =   5970
      Width           =   2985
      _Version        =   65536
      _ExtentX        =   5265
      _ExtentY        =   503
      _StockProps     =   15
      ForeColor       =   0
      BackColor       =   12640511
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Alignment       =   8
   End
   Begin FPSpread.vaSpread spdRst2 
      Height          =   4140
      Left            =   11520
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5970
      Width           =   3615
      _Version        =   393216
      _ExtentX        =   6376
      _ExtentY        =   7303
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      DisplayRowHeaders=   0   'False
      EditEnterAction =   8
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   85
      NoBeep          =   -1  'True
      ProcessTab      =   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "INTERFACE.frx":44F4
      UserResize      =   0
      VisibleCols     =   4
      VisibleRows     =   85
      TextTip         =   1
   End
   Begin FPSpread.vaSpread spdRst 
      Height          =   4140
      Left            =   8130
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5970
      Width           =   3360
      _Version        =   393216
      _ExtentX        =   5927
      _ExtentY        =   7303
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      DisplayRowHeaders=   0   'False
      EditEnterAction =   8
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   15
      NoBeep          =   -1  'True
      ProcessTab      =   -1  'True
      ScrollBars      =   0
      SpreadDesigner  =   "INTERFACE.frx":4E3C
      UserResize      =   0
      VisibleCols     =   4
      VisibleRows     =   15
      TextTip         =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   285
      Left            =   3030
      TabIndex        =   20
      Top             =   6450
      Width           =   4965
      _Version        =   65536
      _ExtentX        =   8758
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "Warning Event Log (삭제 = F2)"
      ForeColor       =   12648447
      BackColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodColor      =   16576
   End
   Begin Threed.SSPanel pnlOrder 
      Height          =   285
      Left            =   3030
      TabIndex        =   21
      Top             =   9495
      Width           =   1965
      _Version        =   65536
      _ExtentX        =   3466
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "최근 전송 Order"
      ForeColor       =   0
      BackColor       =   8454143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Alignment       =   8
   End
   Begin Threed.SSPanel lblOrder 
      Height          =   285
      Left            =   5010
      TabIndex        =   22
      Top             =   9495
      Width           =   2985
      _Version        =   65536
      _ExtentX        =   5265
      _ExtentY        =   503
      _StockProps     =   15
      ForeColor       =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
   End
   Begin Threed.SSPanel pnlResult 
      Height          =   285
      Left            =   3030
      TabIndex        =   23
      Top             =   9795
      Width           =   1965
      _Version        =   65536
      _ExtentX        =   3466
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "최근 수신 Result"
      ForeColor       =   0
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Alignment       =   8
   End
   Begin Threed.SSPanel lblResult 
      Height          =   285
      Left            =   5010
      TabIndex        =   24
      Top             =   9795
      Width           =   2985
      _Version        =   65536
      _ExtentX        =   5265
      _ExtentY        =   503
      _StockProps     =   15
      ForeColor       =   0
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
   End
   Begin Threed.SSCommand cmdMerge 
      Height          =   270
      Left            =   8475
      TabIndex        =   37
      Top             =   11460
      Visible         =   0   'False
      Width           =   1590
      _Version        =   65536
      _ExtentX        =   2805
      _ExtentY        =   476
      _StockProps     =   78
      Caption         =   "블럭리스트 병합"
      ForeColor       =   8388608
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "INTERFACE.frx":5482
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   1545
      Left            =   0
      TabIndex        =   30
      Top             =   5850
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   2725
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.OptionButton optRegOpt 
         Caption         =   "자동등록"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   6
         Top             =   1260
         Width           =   1155
      End
      Begin VB.OptionButton optRegOpt 
         Caption         =   "일괄등록"
         Height          =   180
         Index           =   1
         Left            =   1500
         TabIndex        =   7
         Top             =   1260
         Width           =   1155
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   285
         Left            =   450
         TabIndex        =   31
         Top             =   240
         Width           =   1995
         _Version        =   65536
         _ExtentX        =   3519
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Interface 작업일자"
         ForeColor       =   12648447
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodColor      =   0
      End
      Begin MSComCtl2.DTPicker dtpLabDate 
         Height          =   315
         Left            =   690
         TabIndex        =   5
         Top             =   540
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   64815107
         CurrentDate     =   36737
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   285
         Left            =   450
         TabIndex        =   32
         Top             =   960
         Width           =   1995
         _Version        =   65536
         _ExtentX        =   3519
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "등록 Option"
         ForeColor       =   12648447
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
   End
   Begin Threed.SSFrame fraSendOrd 
      Height          =   1395
      Left            =   0
      TabIndex        =   26
      Top             =   7560
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   2461
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin Threed.SSCommand cmdSendOrd 
         Height          =   585
         Left            =   1890
         TabIndex        =   4
         Top             =   390
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "Send"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtOrdNo 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   990
         TabIndex        =   3
         Top             =   990
         Width           =   1755
      End
      Begin VB.TextBox txtPos 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   990
         MaxLength       =   2
         TabIndex        =   2
         Top             =   690
         Width           =   525
      End
      Begin VB.TextBox txtRack 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   990
         TabIndex        =   1
         Top             =   390
         Width           =   855
      End
      Begin Threed.SSPanel pnlJNo 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   990
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "검체번호"
         ForeColor       =   12648447
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRackTray 
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   390
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Rack"
         ForeColor       =   12648447
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlPosCup 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   690
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Pos"
         ForeColor       =   12648447
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "<One-By-One>  Order 전송"
         Height          =   180
         Left            =   255
         TabIndex        =   36
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.CheckBox chkOExist 
      Caption         =   "chkOExist"
      Height          =   225
      Left            =   660
      TabIndex        =   13
      Top             =   210
      Value           =   1  '확인
      Width           =   1185
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1080
      Left            =   14205
      TabIndex        =   8
      Top             =   4260
      Visible         =   0   'False
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   1905
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSCommLib.MSComm Comm1 
         Left            =   2055
         Top             =   1140
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin Threed.SSFrame fraOrdOpt 
         Height          =   915
         Left            =   6780
         TabIndex        =   9
         Top             =   2010
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   1614
         _StockProps     =   14
         Caption         =   "Option 구분"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Begin VB.OptionButton optOrdOpt 
            Caption         =   "STAT"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   660
            Width           =   915
         End
         Begin VB.OptionButton optOrdOpt 
            Caption         =   "Active"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1005
         End
         Begin VB.OptionButton optOrdOpt 
            Caption         =   "Passive"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   450
            Width           =   975
         End
      End
   End
   Begin Threed.SSPanel SSPanel7 
      Align           =   2  '아래 맞춤
      Height          =   390
      Left            =   0
      TabIndex        =   39
      Top             =   10230
      Width           =   15180
      _Version        =   65536
      _ExtentX        =   26776
      _ExtentY        =   688
      _StockProps     =   15
      Caption         =   "  Interface Board"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Alignment       =   1
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   1830
         TabIndex        =   40
         Top             =   -15
         Width           =   13470
         _ExtentX        =   23760
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   23707
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   555
      Picture         =   "INTERFACE.frx":549E
      Top             =   7845
      Visible         =   0   'False
      Width           =   1950
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim miHlpClick%

Dim miPhase As Integer
Dim msWkbuf As String
Dim msRcvBuffer As String

Dim miIdleFlag As Integer
Dim miPendingFlag As Integer
Dim miOrderFlag As Integer
Dim miResultFlag As Integer
Dim miTimerFlag As Integer

Dim msBeforeLabDate As String
Dim miSpaceCnt%
Dim miTimerCnt%
Dim miLeaveCell%

Dim msRcvState As String
Dim msSndState As String
Dim msSndPacket As String
Dim msBarCdQryState As String
Dim msEmptyOrder As String

Public Sub DisPlayInit_Site()

'    If gsIFMode = "0" Then
'        Frame1.Top = 8400
'        cmdClear1.Top = 530
'        cmdClear2.Top = 910
'        cmdWorkList.Top = 120
'
'        fraBarCd.Visible = True
'        fraBarCd.Top = 7300
'        fraBarCd.Height = 1050
'        pnlJNo.Item(1).Top = 250
'        txtBarCd.Top = 670
'
'        spdIntList.ColWidth(6) = 0
'        spdIntList.ColWidth(7) = 0
'    End If
'
'        cmdClear1.Top = 530
'        cmdClear2.Top = 910
'        cmdWorkList.Top = 120
    
End Sub


Public Sub Disp_SampleOrder(Optional ByVal sGbn As String)
    On Error GoTo ErrHandler
    
    Dim i%, iOrdCnt%
    Dim sTmp$, sBuf$, sOrdList$, sIFSeq$, sTIFSeq$
    Dim sTIFOrdCd$
    Dim objOrd As Object
    
    'Order Dll을 Call하여 서버쪽에 Order를 가져옴
    sBuf = gOrdCfg.sComponent
    
    If sBuf = "" Then
        ViewMsg "오더 Dll 파일이 존재하지 않습니다!!"
        
        Exit Sub
    End If
    
    'Empty Check
    If txtBarCd = "" And txtOrdNo = "" Then Exit Sub
    
    Set objOrd = CreateObject(sBuf)
    Call objOrd.SetMachineInfo(gsMachineCd, gsMachineNm)
    sOrdList = objOrd.FetchOrder(gsMachineCd, "", "", "", gOrderTable.sSampID)
    Set objOrd = Nothing
    
    If sOrdList = "" Then
        ViewMsgLog "검체 ERR : " & gOrderTable.sRack & " " & gOrderTable.sPos & " " & gOrderTable.sSampID
        
        Exit Sub
    Else
        'sOrdList 구성
        '환자번호 | 이름 | n | IFSeq 1 | IFSeq 2 | ... | IFSeq n |
        gOrderTable.sWDate = Format(dtpLabDate.Value, "YYYYMMDD")
        gOrderTable.sJDate = ""
        gOrderTable.sJGbn = ""
        gOrderTable.sJNo = gOrderTable.sSampID
        gOrderTable.sRegNo = GetByOne(sOrdList, sOrdList)
        gOrderTable.sName = GetByOne(sOrdList, sOrdList)
        gOrderTable.sSex = GetByOne(sOrdList, sOrdList)
        gOrderTable.sEmer = GetByOne(sOrdList, sOrdList)
        gOrderTable.iOrdCnt = Val(GetByOne(sOrdList, sOrdList))
        gOrderTable.sOrdOpt = ""
        
        If gOrderTable.iOrdCnt = 0 Then
            ViewMsgLog "오더 ERR : " & gOrderTable.sRack & " " & gOrderTable.sPos & " " & gOrderTable.sSampID
            
            Exit Sub
        Else
            For i = 1 To gOrderTable.iOrdCnt
                sIFSeq = GetByOne(sOrdList, sOrdList)
                
                sTmp = sIFSeq
                
                'IFOrdCd로 변환
                sTmp = ConvertIFItemInfo(6, sTmp)
                
                If sTmp = "" Then
                Else
                    iOrdCnt = iOrdCnt + 1
                    
                    'IFSeq를 합친다
                    sTIFSeq = sTIFSeq & sIFSeq & Chr(124)
                End If
            Next
            
            'IFSeq 순서로 재구성
            sTIFSeq = ReOrder_IFSeq_And_RealOrdCnt(sTIFSeq, iOrdCnt)
            
            gOrderTable.iOrdCnt = iOrdCnt
            ReDim gOrderTable.sIFSeq(iOrdCnt)
            
            For i = 1 To iOrdCnt
                gOrderTable.sIFSeq(i) = GetByOne(sTIFSeq, sTIFSeq)
            Next
        End If
                   
        'ORDER 내역만 Display
        Call DisplayOrderOK("DISPLAY")
    End If
    
    Exit Sub
ErrHandler:
End Sub

Public Function Get_CurOrderInfo() As String
    On Error GoTo ErrHandler
    
    Dim i%, iCRow%, iOrdCnt%
    Dim vChk, vIFCnt, vTmp
    Dim sBuf$, sTmp$, sIFSeq$, sTIFOrdCd$
    
    iCRow = 0
    
    With spdIntList
        For i = 1 To .MaxRows
            Call .GetText(2, i, vChk)
            
            If vChk = "1" Then
                iCRow = i
                
                Exit For
            End If
        Next
    End With
    
    If iCRow = 0 Then
        Get_CurOrderInfo = "NONE"
        Exit Function
    End If
    
    'Order Count
    With spdIntList
    '--- 사이트 + 장비 특성 ----------------------------------------------------------------------------------------------
        Call .GetText(5, iCRow, vTmp)
        gOrderTable.sSampID = CStr(vTmp)
        gOrderTable.iCRow = iCRow
    '---------------------------------------------------------------------------------------------------------------------
        
        Call .GetText(16, iCRow, vIFCnt)
        
        sTIFOrdCd = ""
        iOrdCnt = 0
        
        If Val(vIFCnt) = 0 Then
            Get_CurOrderInfo = "NONE"
            Exit Function
        End If
        
        For i = 1 To CInt(Val(vIFCnt))
            Call .GetText(16 + i, iCRow, vTmp)
            
            sTmp = CStr(vTmp)
            sIFSeq = GetByOne(sTmp, sTmp)
            
            'IFOrdCd로 변환
            sTmp = ConvertIFItemInfo(6, sIFSeq)
            
            If sTmp = "" Then
            Else
                sTIFOrdCd = sTIFOrdCd & sTmp & ","
            End If
        Next
        
        '--- 계산식 등을 위한 오더로 재구성 S -----------------------------------------------------------------------------
        sTmp = sTIFOrdCd
        
        sTIFOrdCd = RemoveDuplicatedOrder(sTmp, iOrdCnt)
        '--- 계산식 등을 위한 오더로 재구성 E -----------------------------------------------------------------------------
    End With
       
    msSndPacket = ""
    msSndPacket = Chr(2)
    msSndPacket = msSndPacket & "S" & gOrderTable.sSampNo
    msSndPacket = msSndPacket & String(8 - Len(gOrderTable.sSampID), " ") & gOrderTable.sSampID
    
    sBuf = ""
    
    For i = 1 To iOrdCnt
        sTmp = GetByOneUserSymbol(sTIFOrdCd, sTIFOrdCd, ",")
        
        If sTmp = "" Then
        Else
            sBuf = sBuf & sTmp
        End If
    Next
    
    msSndPacket = msSndPacket & sBuf & Chr(3)
    
    Exit Function
    
ErrHandler:
    Get_CurOrderInfo = "NONE"
End Function

Private Sub Get_OrderString(Optional ByVal sSendYN$)
    On Error GoTo ErrRtn
    
    Dim i%, iOrdCnt%
    Dim vIFCnt, vTmp
    Dim sTmp$, sOrdList$, sIFSeq$, sBuf$, sTIFSeq$
    Dim sTIFOrdCd$
    Dim objOrd  As Object
    
    Dim sSpcGbn$, sTmpSpc$
    
    sBuf = ""
    msEmptyOrder = ""
    iOrdCnt = 0
    sOrdList = ""
    
    'Order Dll을 Call하여 서버쪽에 Order를 가져옴
    sBuf = gOrdCfg.sComponent
    
    If sBuf = "" Then
        ViewMsg "오더 Dll 파일이 존재하지 않습니다!!"
        Exit Sub
    End If
    
    If Trim(gOrderTable.sSampID) <> "" Then
        Set objOrd = CreateObject(sBuf)
        Call objOrd.SetMachineInfo(gsMachineCd, gsMachineNm)
        sOrdList = objOrd.FetchOrder(gsMachineCd, "", "", "", gOrderTable.sSampID)
        Set objOrd = Nothing
    End If
    
'    sOrdList = "12345|TEST|M|10|001|002|003|004|005|006|007|008|009|010|"
    If sOrdList = "" Then
        ViewMsgLog "검체 ERR : " & Trim(gOrderTable.sRack & " " & gOrderTable.sPos & " " & gOrderTable.sSampID)
        
        If sSendYN = "N" Then
            Exit Sub
        End If
    End If
    
    'sOrdList 구성
    '환자번호 | 이름 | 성별 | n | IFSeq 1 | IFSeq 2 | ... | IFSeq n |
    gOrderTable.sWDate = Format(dtpLabDate.Value, "YYYYMMDD")
    gOrderTable.sJDate = Format(dtpLabDate.Value, "YYYYMMDD")
    gOrderTable.sJGbn = ""
    gOrderTable.sJNo = gOrderTable.sSampID
    gOrderTable.sRegNo = GetByOne(sOrdList, sOrdList)
    gOrderTable.sName = GetByOne(sOrdList, sOrdList)
    gOrderTable.sSex = GetByOne(sOrdList, sOrdList)
    gOrderTable.sEmer = ""
    gOrderTable.sReRun = ""
    gOrderTable.sOrdOpt = ""
    gOrderTable.sOther = ""
    gOrderTable.iOrdCnt = Val(GetByOne(sOrdList, sOrdList))
        
    If gOrderTable.iOrdCnt = 0 Then
        ViewMsgLog "오더 ERR : " & Trim(gOrderTable.sSampNo & " " & gOrderTable.sSampID)
        
        'NO QUERY 방식일 때
        Exit Sub
    Else
        For i = 1 To gOrderTable.iOrdCnt
            sIFSeq = GetByOne(sOrdList, sOrdList)
            
            sTmp = sIFSeq
            
            'IFOrdCd로 변환
            sTmp = ConvertIFItemInfo(6, sTmp)
            
            If sTmp = "" Then
            Else
                iOrdCnt = iOrdCnt + 1
                
                'IFSeq를 합친다
                sTIFSeq = sTIFSeq & sIFSeq & Chr(124)
            End If
        Next i
        
        'IFSeq 순서로 재구성
        sTIFSeq = ReOrder_IFSeq_And_RealOrdCnt(sTIFSeq, iOrdCnt)
        
        gOrderTable.iOrdCnt = iOrdCnt
        ReDim gOrderTable.sIFSeq(iOrdCnt)
        
        For i = 1 To iOrdCnt
            gOrderTable.sIFSeq(i) = GetByOne(sTIFSeq, sTIFSeq)
        Next i
    End If
                            
    'Order 전송없이 내용만 Display - DisplayResultOK 에서 사용
    If sSendYN = "N" And iOrdCnt > 0 Then
        Call DisplayOrderOK
        
        Exit Sub
    End If
    
'    '--- 계산식 등을 위한 오더로 재구성 S --------------------------------------------------
'    sTIFOrdCd = ""
'
'    For i = 1 To iOrdCnt
'        sTmp = ConvertIFItemInfo(6, gOrderTable.sIFSeq(i))
'
'        If sTmp = "" Then
'        Else
'            sTIFOrdCd = sTIFOrdCd & sTmp & ","
'
'            'For 검체구분
'            sTmpSpc = ConvertIFItemInfo(11, gOrderTable.sIFSeq(i))
'            If sTmpSpc <> "" And sTmpSpc <> "1" Then
'                sSpcGbn = sTmpSpc
'            End If
'        End If
'    Next
'
'    sTmp = sTIFOrdCd
'
'    'For 검체구분
'    gOrderTable.sIFSpcCd = sSpcGbn
'
'
'    sTIFOrdCd = RemoveDuplicatedOrder(sTmp, iOrdCnt)
'    '--- 계산식 등을 위한 오더로 재구성 E --------------------------------------------------
'
'    ReDim gOrderTable.sIFOrdCd(iOrdCnt)
'
'    For i = 1 To iOrdCnt
'        gOrderTable.sIFOrdCd(i) = GetByOneUserSymbol(sTIFOrdCd, sTIFOrdCd, ",")
'    Next i
'    gOrderTable.iOrdCnt = iOrdCnt
    
ErrRtn:
    If Err <> 0 Then
        ViewMsg "Get_OrderString 오류 - (" & Err.Description & ")"
    End If
End Sub

Private Function GetNowOrderList() As String
    On Error GoTo ErrRtn
    
    Dim i%, iOrdCnt%
    Dim vIFCnt, vTmp, vChk, vChk2, vCnt
    Dim sTmp$, sOrdList$, sIFSeq$, sBuf$, sTIFSeq$
    Dim sTIFOrdCd$
    Dim objOrd  As Object
    Dim sLabDate$, sLabNo$, sSpcSeq$, sResDate$
    Dim tmpData()   As String
    
    Dim sJDate$, sJNo$, sRegNo$, sNm$, sSex$, sRack$, sPos$, sReRun$, sOther$
    Dim i1stRow%
    
    i1stRow = 0
    
    With spdIntList
        For i = 1 To .MaxRows
            Call .GetText(2, i, vChk)
            Call .GetText(15, i, vChk2)
            
            If Trim(vChk) = "1" And Trim(vChk2) = "N" Then
                i1stRow = i
                Exit For
            End If
        Next i
    
        If i1stRow = 0 Then
            ViewMsg "Now, Order 전송을 할 데이터가 없습니다..."
            miOrderFlag = 0
            GetNowOrderList = "NONE"
            Exit Function
        End If
        
        GetNowOrderList = "OK"
        
        gOrderTable.iCRow = i1stRow
        
        Call .GetText(14, gOrderTable.iCRow, vCnt)
        
        sOrdList = ""
        For i = 1 To Val(vCnt)
            Call .GetText(16 + i, gOrderTable.iCRow, vTmp)
            If Trim(vTmp) <> "" Then
                tmpData() = Split(Trim(vTmp), Chr(124))
                sOrdList = sOrdList & Trim(tmpData(0) & "") & Chr(124)
            End If
        Next i
        sOrdList = Trim(vCnt) & Chr(124) & sOrdList
        
        Call .GetText(3, gOrderTable.iCRow, vTmp): sJDate = Trim(vTmp)
        Call .GetText(5, gOrderTable.iCRow, vTmp): sJNo = Trim(vTmp)
        Call .GetText(6, gOrderTable.iCRow, vTmp): sRack = Trim(vTmp)
        Call .GetText(7, gOrderTable.iCRow, vTmp): sPos = Trim(vTmp)
        Call .GetText(8, gOrderTable.iCRow, vTmp): sRegNo = Trim(vTmp)
        Call .GetText(9, gOrderTable.iCRow, vTmp): sNm = Trim(vTmp)
        Call .GetText(10, gOrderTable.iCRow, vTmp): sSex = Trim(vTmp)
        Call .GetText(12, gOrderTable.iCRow, vTmp): sReRun = Trim(vTmp)
        Call .GetText(13, gOrderTable.iCRow, vTmp): sOther = Trim(vTmp)
    End With
    
    gOrderTable.sWDate = Format(dtpLabDate.Value, "YYYYMMDD")
    gOrderTable.sJDate = sJDate
    gOrderTable.sJGbn = ""
    gOrderTable.sJNo = sJNo
    
    gOrderTable.sSampID = gOrderTable.sJNo
'    gOrderTable.sSampID = Mid(gOrderTable.sJDate, 3) & "-" & gOrderTable.sJNo
    
    gOrderTable.sRegNo = sRegNo
    gOrderTable.sName = sNm
    gOrderTable.sSex = sSex
    gOrderTable.sEmer = ""
    gOrderTable.sReRun = sReRun
    gOrderTable.sOther = sOther
    gOrderTable.sRack = sRack
    gOrderTable.sPos = sPos
    
    gOrderTable.iOrdCnt = Val(GetByOne(sOrdList, sOrdList))
    gOrderTable.sOrdOpt = ""
    
    If gOrderTable.sRegNo <> "" And gOrderTable.iOrdCnt = 0 Then
        ViewMsgLog "오더 ERR : " & Trim(gOrderTable.sSampNo & " " & gOrderTable.sSampID)
        
        'NO QUERY 방식일 때
        Exit Function
    Else
        For i = 1 To gOrderTable.iOrdCnt
            sIFSeq = GetByOne(sOrdList, sOrdList)
            
            sTmp = sIFSeq
            
            'IFOrdCd로 변환
            sTmp = ConvertIFItemInfo(6, sTmp)
            
            If sTmp = "" Then
            Else
                iOrdCnt = iOrdCnt + 1
                
                'IFSeq를 합친다
                sTIFSeq = sTIFSeq & sIFSeq & Chr(124)
            End If
        Next i
        
        'IFSeq 순서로 재구성
        sTIFSeq = ReOrder_IFSeq_And_RealOrdCnt(sTIFSeq, iOrdCnt)
        
        gOrderTable.iOrdCnt = iOrdCnt
        ReDim gOrderTable.sIFSeq(iOrdCnt)
        
        For i = 1 To iOrdCnt
            gOrderTable.sIFSeq(i) = GetByOne(sTIFSeq, sTIFSeq)
        Next i
    End If
    
    '--- 계산식 등을 위한 오더로 재구성 S --------------------------------------------------
    sTIFOrdCd = ""
    
    For i = 1 To iOrdCnt
        sTmp = ConvertIFItemInfo(6, gOrderTable.sIFSeq(i))
        
        If sTmp = "" Then
        Else
            sTIFOrdCd = sTIFOrdCd & sTmp & ","
        End If
    Next
    
    sTmp = sTIFOrdCd
    
    sTIFOrdCd = RemoveDuplicatedOrder(sTmp, iOrdCnt)
    '--- 계산식 등을 위한 오더로 재구성 E --------------------------------------------------
    
    ReDim gOrderTable.sIFOrdCd(iOrdCnt)

    For i = 1 To iOrdCnt
        gOrderTable.sIFOrdCd(i) = GetByOneUserSymbol(sTIFOrdCd, sTIFOrdCd, ",")
    Next i

    gOrderTable.iOrdCnt = iOrdCnt
    
    lblOrder = gOrderTable.sJNo
    
ErrRtn:
    If Err <> 0 Then
        ViewMsg "Get_OrderString 오류 - (" & Err.Description & ")"
    End If
End Function

Public Sub Order_Input(Optional ByVal sSendYN$)
    On Error GoTo ErrHandler
    
    Dim i%, iOrdCnt%
    Dim vIFCnt, vTmp
    Dim sTmp$, sOrdList$, sIFSeq$, sBuf$, sTIFSeq$
    Dim sTIFOrdCd$
    Dim objOrd As Object
    
    sBuf = ""
    msEmptyOrder = ""
    iOrdCnt = 0
    sOrdList = ""
    
    'Order Dll을 Call하여 서버쪽에 Order를 가져옴
    sBuf = gOrdCfg.sComponent
    
    If sBuf = "" Then
        ViewMsg "오더 Dll 파일이 존재하지 않습니다!!"
        Exit Sub
    End If
    
    If Trim(gOrderTable.sSampID) <> "" Then
        Set objOrd = CreateObject(sBuf)
        Call objOrd.SetMachineInfo(gsMachineCd, gsMachineNm)
        sOrdList = objOrd.FetchOrder(gsMachineCd, "", "", "", gOrderTable.sSampID)
        Set objOrd = Nothing
    End If
    
    If sOrdList = "" Then
        ViewMsgLog "검체 ERR : " & Trim(gOrderTable.sRack & " " & gOrderTable.sPos & " " & gOrderTable.sSampID)
        
        If sSendYN = "N" Then
            Exit Sub
        End If
    End If
    
    'sOrdList 구성
    '환자번호 | 이름 | 성별 | n | IFSeq 1 | IFSeq 2 | ... | IFSeq n |
    gOrderTable.sWDate = Format(dtpLabDate.Value, "YYYYMMDD")
    gOrderTable.sJDate = Format(dtpLabDate.Value, "YYYYMMDD")
    gOrderTable.sJGbn = ""
    gOrderTable.sJNo = gOrderTable.sSampID
    gOrderTable.sRegNo = GetByOne(sOrdList, sOrdList)
    gOrderTable.sName = GetByOne(sOrdList, sOrdList)
    gOrderTable.sSex = GetByOne(sOrdList, sOrdList)
    gOrderTable.sEmer = ""
    gOrderTable.sReRun = ""
    gOrderTable.sOrdOpt = ""
    gOrderTable.sOther = ""
    gOrderTable.iOrdCnt = Val(GetByOne(sOrdList, sOrdList))
    
    If gOrderTable.iOrdCnt = 0 Then
        ViewMsgLog "오더 ERR : " & Trim(gOrderTable.sSampNo & " " & gOrderTable.sSampID)
        
        'NO QUERY 방식일 때
        Exit Sub
    Else
        For i = 1 To gOrderTable.iOrdCnt
            sIFSeq = GetByOne(sOrdList, sOrdList)
            
            sTmp = sIFSeq
            
            'IFOrdCd로 변환
            sTmp = ConvertIFItemInfo(6, sTmp)
            
            If sTmp = "" Then
            Else
                iOrdCnt = iOrdCnt + 1
                
                'IFSeq를 합친다
                sTIFSeq = sTIFSeq & sIFSeq & Chr(124)
            End If
        Next
        
        'IFSeq 순서로 재구성
        sTIFSeq = ReOrder_IFSeq_And_RealOrdCnt(sTIFSeq, iOrdCnt)
        
        gOrderTable.iOrdCnt = iOrdCnt
        ReDim gOrderTable.sIFSeq(iOrdCnt)
        
        For i = 1 To iOrdCnt
            gOrderTable.sIFSeq(i) = GetByOne(sTIFSeq, sTIFSeq)
        Next
    End If
                            
    'Order 전송없이 내용만 Display - DisplayResultOK 에서 사용
    If sSendYN = "N" And iOrdCnt > 0 Then
        Call DisplayOrderOK
        
        Exit Sub
    End If
    
    'Order 전송없이 내용만 Display - '바코드 정보 수정'에서 사용
    If sSendYN = "B" And iOrdCnt > 0 Then
        '일단 기존 정보 화면삭제...
        Dim vCurSeq
        With spdIntList
            Call .GetText(1, .ActiveRow, vCurSeq)
        
            .BlockMode = True
            .Col = 3: .Col2 = .MaxCols
            .Row = .ActiveRow: .Row2 = .ActiveRow
            .Action = ActionClearText
            .BlockMode = False
        End With
        
        Call DisplayOrderInfo(Trim(vCurSeq), spdIntList.ActiveRow)
        
        Exit Sub
    End If
    
    '--- 계산식 등을 위한 오더로 재구성 S --------------------------------------------------
    sTIFOrdCd = ""
    
    For i = 1 To iOrdCnt
        sTmp = ConvertIFItemInfo(6, gOrderTable.sIFSeq(i))
        
        If sTmp = "" Then
        Else
            sTIFOrdCd = sTIFOrdCd & sTmp & ","
        End If
    Next
    
    sTmp = sTIFOrdCd
    
    sTIFOrdCd = RemoveDuplicatedOrder(sTmp, iOrdCnt)
    '--- 계산식 등을 위한 오더로 재구성 E --------------------------------------------------
    
    Exit Sub
ErrHandler:
    ViewMsg "Order_Input 오류 - (" & Err.Description & ")"
End Sub

Private Function Order_Check() As String
    On Error GoTo ErrRtn

    Dim ii      As Integer
    Dim sTmp    As String
    Dim sBuf    As String
    Dim tmpData()   As String
        
    With gOrderTable
        sTmp = ""
        For ii = 1 To .iOrdCnt
            sTmp = sTmp & Trim(.sIFOrdCd(ii)) & ","
        Next ii
    End With
    
    '중복된 Order를 제거(계산식 등에 관련)
    sBuf = RemoveDuplicatedOrder(sTmp, gOrderTable.iOrdCnt)
    
    '중복이 제거된 순수 Order 문자열을 가지고 재조합
    tmpData() = Split(sBuf, ",")
    sTmp = ""
    For ii = 1 To gOrderTable.iOrdCnt
        sTmp = sTmp & tmpData(ii - 1) & Chr(124)
    Next ii

    Order_Check = sTmp

ErrRtn:
    If Err <> 0 Then
        Order_Check = ""
        ViewMsg Err.Description
    End If
End Function


Public Function SpecificProcessResult(ByVal sIFRstCd$, sSpRst1$, sSpRst2$, Optional ByVal sIFSeq$, Optional ByVal sSex$) As Integer
    On Error GoTo ErrHandler
    
    Dim vIFItemCnt, vTmp
    Dim sFlag$, sBuf$
    
    SpecificProcessResult = 0
    
'    If IsNumeric(sSpRst1) = False Then
'        If Left(sSpRst1, 1) = ">" Then
'            sFlag = ">"
'
'            If Len(sSpRst1) > 1 Then
'                sSpRst1 = Trim(Mid(sSpRst1, 2))
'            End If
'        ElseIf Left(sSpRst1, 1) = "<" Then
'            sFlag = "<"
'
'            If Len(sSpRst1) > 1 Then
'                sSpRst1 = Trim(Mid(sSpRst1, 2))
'            End If
'        End If
'    Else
'        sFlag = ""
'    End If
    
    If sIFSeq <> "" Then
        sIFRstCd = ConvertIFItemInfo(8, sIFSeq)
    Else
        sIFSeq = ConvertIFItemInfo(7, sIFRstCd)
    End If
    
    sSpRst1 = ConvertResult1("", "", sSpRst1, sIFRstCd, sIFSeq)
    
    sSpRst1 = JudgeRstBySex(sIFSeq, sSpRst1, sSex, sSpRst2)
    
'    Select Case Left(sSpRst1, 1)
'        Case "N"
'            sSpRst1 = "Negative"
'        Case "P"
'            sSpRst1 = "Positive"
'        Case "T"
'            sSpRst1 = "Weakly Positive"
'    End Select
'    Select Case Left(sSpRst2, 1)
'        Case "N"
'            sSpRst1 = "Neg(" & Trim(sSpRst1) & ")"
'        Case "P"
'            sSpRst1 = "Pos(" & Trim(sSpRst1) & ")"
'        Case "T"
'            sSpRst1 = "W.Pos(" & Trim(sSpRst1) & ")"
'    End Select
    
'    sSpRst2 = ""
    
    Exit Function
ErrHandler:
    ViewMsg "SpecificProcessResult(" & Err.Description & ")"
End Function
Private Sub ConvertResultData(ByVal iCnt As Integer, ByVal sTIFCd As String, _
                            ByRef sTRst1 As String, ByRef sTRst2 As String, ByRef sTFlag As String)
    On Error GoTo ErrRtn
    
    Dim ii      As Integer
    Dim sIFCd() As String
    Dim sRst1() As String
    Dim sRst2() As String
    Dim sFlag() As String
    Dim tmpIFCd As String
    Dim tmpRst1 As String
    Dim tmpRst2 As String
    Dim tmpFlag As String
    
    sIFCd() = Split(sTIFCd, Chr(124))
    sRst1() = Split(sTRst1, Chr(124))
    sRst2() = Split(sTRst2, Chr(124))
    sFlag() = Split(sTFlag, Chr(124))
    
    sTRst1 = ""
    sTRst2 = ""
    sTFlag = ""
    
    For ii = 0 To iCnt - 1
        tmpIFCd = sIFCd(ii)
        tmpRst1 = sRst1(ii)
        tmpRst2 = sRst2(ii)
        tmpFlag = sFlag(ii)

'        If InStr(tmpFlag, "r") > 0 Then       'Flag가 'Data Transfer to Host'인 경우는 무시(AU-5400/2700)...2007/5/29 yk
'            tmpFlag = Replace(tmpFlag, "r", "")
'        End If
        
        If Trim(tmpRst1) <> "" Then
            '결과자릿수 편집/정성결과 표시
            Call SpecificProcessResult(tmpIFCd, tmpRst1, tmpRst2, , "M")
        End If
                
        sTRst1 = sTRst1 & tmpRst1 & Chr(124)
        sTRst2 = sTRst2 & tmpRst2 & Chr(124)
        sTFlag = sTFlag & tmpFlag & Chr(124)
     Next ii
    
ErrRtn:
    If Err <> 0 Then
        ViewMsg Err.Description
    End If
End Sub
Private Sub Init_gOrderTable()

    'gOrderTable 초기화
    With gOrderTable
        .iCRow = 0
        .iOrdCnt = 0
        .sEmer = ""
        Erase .sIFOrdCd
        Erase .sIFRstCd
        Erase .sIFSeq
        .sIFSpcCd = ""
        .sJDate = ""
        .sJGbn = ""
        .sJNo = ""
        .sName = ""
        .sOrdOpt = ""
        .sOther = ""
        .sPos = ""
        .sRack = ""
        .sRegNo = ""
        .sReRun = ""
        .sSampID = ""
        Erase .sServerCd
        .sSex = ""
        .sWDate = ""
        .sWSeq = ""
    End With

End Sub

Private Sub cmdClear1_Click()
    Call ClearBlockedList
End Sub

Private Sub cmdClear2_Click()
    If spdIntList.MaxRows < 1 Then Exit Sub
    
    Call ClearAllList
End Sub

Private Sub cmdInitial_Click()

'    Call INTEGRA1.ConnectionMsg
    
End Sub

Private Sub cmdMerge_Click()

    Call WMerge_SPD

End Sub

Private Sub cmdSendOrd_Click()
'    On Error GoTo ErrHandler
'
'    Dim sRackFlag$
'
'    If miTimerFlag = 1 Then
'        ViewMsg "이전 데이터 처리 중입니다. 잠시 후 작업하십시요!!"
'        Exit Sub
'    End If
'
'    If txtOrdNo = "" Or txtRack = "" Or txtPos = "" Then
'        ViewMsg "필수항목이 비어있습니다!!"
'        Exit Sub
'    End If
'
'    gOrderTable.sRack = txtRack
'    gOrderTable.sPos = txtPos
'    gOrderTable.sSampID = txtOrdNo
'
'    Me.MousePointer = vbHourglass
'    If gOrderTable.sSampID <> "" Then
'        'Manual Rack, Pos Order 내리는 동안 Timer 가동 중지!!
'        miTimerFlag = 1
'        '--- Rack/Pos 별로 오더전송
'        Call DPC1_RequestCurOrder(gOrderTable.sSampID, gOrderTable.sRack, gOrderTable.sPos)
'
'        DPC1.Send_Chr (5)
'        DPC1.iPhase = 3
'        '-------------------------
''        Call Order_Input("R")
'        miTimerFlag = 0
'    End If
'    Me.MousePointer = vbDefault
'
'    sRackFlag = gIFRack.sPosSetting
'
'    Call DisplayNextRackPos(GetByOne(sRackFlag, sRackFlag))
'    txtOrdNo.SetFocus
'
'    Exit Sub
'ErrHandler:
'    miTimerFlag = 0
End Sub

Private Sub cmdSendOrder_Click()
    Dim sBuf$
    Dim sTestCd$

    Me.MousePointer = vbHourglass
    
    sBuf = GetNowOrderList
    
    If sBuf <> "OK" Then
        Me.MousePointer = vbDefault
        msSndState = ""
        miPhase = 1
        Exit Sub
    Else
        '검사항목 편집
        sTestCd = ""
        
        '<S--- 계산식등에 의해 오더와 결과 갯수가 다른 경우 처리...2007/5/25 yk
        Dim sTmp$, sTIFOrdCd$
        Dim iOrdCnt%, i%, ii%
        Dim aIFOrdCd()    As String
        '--- 계산식 등을 위한 오더로 재구성 S ---------------------
        sTIFOrdCd = ""
        iOrdCnt = gOrderTable.iOrdCnt
        
        For i = 1 To iOrdCnt
            sTmp = ConvertIFItemInfo(6, gOrderTable.sIFSeq(i))
            
            If sTmp = "" Then
            Else
                sTIFOrdCd = sTIFOrdCd & sTmp & ","
            End If
        Next
        
        sTmp = sTIFOrdCd
        
        sTIFOrdCd = RemoveDuplicatedOrder(sTmp, iOrdCnt)
        '--- 계산식 등을 위한 오더로 재구성 E ---------------------
        
        ReDim aIFOrdCd(iOrdCnt)
    
        For i = 1 To iOrdCnt
            aIFOrdCd(i) = GetByOneUserSymbol(sTIFOrdCd, sTIFOrdCd, ",")
        Next i
        '>E----------------------------------------------------
        
        For ii = 1 To iOrdCnt
            sTestCd = sTestCd & Trim(aIFOrdCd(ii)) & Chr(124)
        Next ii
    
        With MAXMAT1
            .p_sID = gOrderTable.sSampID
            .p_sSeq = ""
            .p_sRack = gOrderTable.sRack
            .p_sPos = gOrderTable.sPos
            .p_iOrdCnt = gOrderTable.iOrdCnt
            .p_sTIFCd = sTestCd
                            
            .iFrameN = 1
            .iSendPhase = 1
            .iPhase = 3
            .Send_Chr (5)
        End With
    End If
    
    Me.MousePointer = vbDefault
    
    Exit Sub
End Sub

Private Sub cmdWorkList_Click()
    Load frmWorkList
    frmWorkList.Left = Me.Left + Me.Width - frmWorkList.Width - 700
    frmWorkList.Top = Me.Top + 500
    frmWorkList.Show 1
End Sub

Public Function Chk_spdIntList(ByVal sSampleID As String) As Boolean
        
    Dim vTmp    As Variant
    Dim ii      As Integer
    
    Chk_spdIntList = False
    
    With spdIntList
        For ii = 1 To .MaxRows
            .Col = 1: .Row = ii
            If .BackColor = vbWhite Then
                Call .GetText(5, ii, vTmp)
                If Trim(vTmp) = sSampleID Then
                    Chk_spdIntList = True
                    Exit For
                End If
            End If
        Next ii
    End With
End Function

Private Sub dtpLabDate_Change()
    If MsgBox("Interface 작업일자를 바꾸시겠습니까?" & vbCrLf & _
                "화면 List가 Clear되는 것을 유의하시기 바랍니다.", vbYesNo, "Interface 작업일자 전환 여부") = vbYes Then
        
        Call cmdClear2.DoClick
        
        Call GetLastWorkSeq(Format(dtpLabDate.Value, "YYYYMMDD"))
    Else
        dtpLabDate.Value = Left(msBeforeLabDate, 4) & "-" & Mid(msBeforeLabDate, 5, 2) & "-" & Right(msBeforeLabDate, 2)
    End If
End Sub

Private Sub dtpLabDate_Click()
    msBeforeLabDate = Format(dtpLabDate.Value, "YYYYMMDD")
End Sub

Private Sub dtpLabDate_GotFocus()
    msBeforeLabDate = Format(dtpLabDate.Value, "YYYYMMDD")
End Sub

Private Sub Form_Activate()
    If miHlpClick = 0 Then
    ElseIf miHlpClick = 1 Then
        If MsgBox("환경설정을 새로 바꾸셨다면 프로그램을 다시 시작해야 합니다. " & _
            vbCrLf & "다시 시작하시겠습니까?", vbYesNo, "프로그램 재시작 여부") = vbYes Then
            
            Unload Me
        End If
        
        miHlpClick = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        listNoOrd.Clear
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim sUseYN$
    Dim bRetVal As Boolean
    
    Me.Top = 0: Me.Left = 0
    
    miHlpClick = 0
    miSpaceCnt = 0
    miTimerCnt = 0
    
    Call RegViewMsgHwnd(Me.StatusBar1.hwnd)
    Call GetMachineInfo
    
'    Call GetTestItem
    Call GetTestItem_Flag
    Call GetOrdRstCfg
    Call GetTestCdSeq
    Call GetTestMode
    Call GetCSMode
    
    Call DisplayInit
    Call DisplayInitItem
    Call DisPlayInit_Site
    
    'Server등록, Client등록 옵션의 디폴트값 설정
    Call SetDefaultRegOption
    
'    Call PortOpen
    Call LogFileOpen
    
    miPhase = 1
    miTimerFlag = 0
    miIdleFlag = 0
    miPendingFlag = 0
    miOrderFlag = 0
    miResultFlag = 0
    
    ViewMsg "Interface Program Ready..."
    
    dtpLabDate.Value = Format(Now, "YYYY-MM-DD")
    
    Call GetLastWorkSeq(Format(dtpLabDate.Value, "YYYYMMDD"))
    
    If gsINITMode = "1" Then
        Call cmdInitial_Click
    End If
    
    '--- 각 장비별 설정
    If giTestMode = 77 Then
        MAXMAT1.Visible = True
    Else
        MAXMAT1.Visible = False: MAXMAT1.TabStop = False
    End If
    
    Frame2.Visible = False
    cmdSendOrder.Visible = True
    cmdWorkList.Visible = False
    fraBarCd.Visible = True
    
    '장비 OCX 초기화
    With MAXMAT1
        .OpenPW = "ACK"
        .EditPW = "MEDI@CK"
        .EqName = "MAXMAT"
        .bUseBarcode = True
        .iPhase = 1
        .iSendPhase = 1
        .iFrameN = 1
        .sTestMode = Trim(giTestMode)
        .CommPort = gCommInfo.sPort
        .Settings = gCommInfo.sBaudRate & "," & gCommInfo.sParity & "," & gCommInfo.sDataBit & "," & gCommInfo.sStopBIt
        .RTSEnable = True
        .RThreshold = 1
        .PortOpen = True
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If miHlpClick = 1 Then
    Else
        If MsgBox("[" & UCase$(gsMachineNm) & "]" & " Interface Program을 종료하시겠습니까?" & vbCrLf & vbCrLf & _
                "Interface 작업 도중에 종료할 경우 전송데이터가 손실이 됩니다.", vbYesNo + vbQuestion, _
                "Interface 종료 확인") = vbYes Then
                
            Unload Me
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call PortClose
    Call LogFileClose
    
    RegEditCurFrmTitle "IF", ""
End Sub

Private Sub cmdExit_Click()
    If MsgBox("[" & UCase$(gsMachineNm) & "]" & " Interface Program을 종료하시겠습니까?" & vbCrLf & vbCrLf & _
            "Interface 작업 도중에 종료할 경우 전송데이터가 손실이 됩니다.", vbYesNo + vbQuestion, _
            "Interface 종료 확인") = vbYes Then
            
        miHlpClick = 1
        Unload Me
    End If
End Sub

Private Sub MAXMAT1_AppendData(sID As String, sSeq As String, sRack As String, sPos As String, iRstCnt As Integer, sTIFCd As String, sTRst1 As String, sTRst2 As String, sTUnit As String, sTFlag As String, sTAlarmCd As String, sKind As String, sTRstDT As String, sTOther As String)

    With gOrderTable
        .sSampID = sID
        .sSampNo = sSeq
        .sRack = sRack
        .sPos = sPos
    End With
    
    '----- 결과 저장 및 화면표시
'    If gOrderTable.sSampid <> "" Then
        '결과값 편집(소숫점 변경 처리)
        Call ConvertResultData(iRstCnt, sTIFCd, sTRst1, sTRst2, sTFlag)
        
        Call DisplayResultOK(3, Format(dtpLabDate.Value, "YYYYMMDD"), "", _
                            "", "", gOrderTable.sSampID, gOrderTable.sRack, gOrderTable.sPos, _
                            "", "", "", "", "", "", _
                            iRstCnt, sTIFCd, sTRst1, sTRst2, _
                            "", "", sTFlag)
'    End If
    
End Sub

Private Sub MAXMAT1_DispMsg(sMsg As String)
    ViewMsg sMsg
End Sub


Private Sub MAXMAT1_PrintRcvLog(sLog As String)
    Print #1, sLog;
End Sub


Private Sub MAXMAT1_PrintSendLog(sLog As String)
    Print #2, sLog;
End Sub


Private Sub MAXMAT1_RaiseError(sError As String)
    MsgBox sError, vbCritical, Me.Caption
    
    Unload Me
End Sub


Private Sub MAXMAT1_RequestCurOrder(sID As String, sSeq As String, sRack As String, sPos As String)

    Dim ii      As Integer
    Dim sTestCd As String
    Dim tmpID$, tmpSeq$, tmpRack$, tmpPos$
    
    tmpID = sID
    tmpSeq = ""
    tmpRack = sRack
    tmpPos = sPos
    
    Call Init_gOrderTable
    
    With gOrderTable
        .sSampID = tmpID
        .sSampNo = tmpSeq
        .sRack = tmpRack
        .sPos = tmpPos
        .iOrdCnt = 0
    End With
    
    '----- 검사항목 조회
    Call Get_OrderString
    
    If gOrderTable.iOrdCnt = 0 Then
        ViewMsg "인터페이스 오더 항목이 존재하지 않습니다!!"
        
        With MAXMAT1
            .p_sID = sID
            .p_sSeq = ""
            .p_sRack = sRack
            .p_sPos = sPos
            .p_iOrdCnt = 0
            .p_sTIFCd = ""
        End With
    
        Exit Sub
    End If

    '검사항목 편집
    sTestCd = ""
    
    '<S--- 계산식등에 의해 오더와 결과 갯수가 다른 경우 처리...2007/5/25 yk
    Dim sTmp$, sTIFOrdCd$
    Dim iOrdCnt%, i%
    Dim aIFOrdCd()    As String
    '--- 계산식 등을 위한 오더로 재구성 S ---------------------
    sTIFOrdCd = ""
    iOrdCnt = gOrderTable.iOrdCnt
    
    For i = 1 To iOrdCnt
        sTmp = ConvertIFItemInfo(6, gOrderTable.sIFSeq(i))
        
        If sTmp = "" Then
        Else
            sTIFOrdCd = sTIFOrdCd & sTmp & ","
        End If
    Next
    
    sTmp = sTIFOrdCd
    
    sTIFOrdCd = RemoveDuplicatedOrder(sTmp, iOrdCnt)
    '--- 계산식 등을 위한 오더로 재구성 E ---------------------
    
    ReDim aIFOrdCd(iOrdCnt)

    For i = 1 To iOrdCnt
        aIFOrdCd(i) = GetByOneUserSymbol(sTIFOrdCd, sTIFOrdCd, ",")
    Next i
    '>E----------------------------------------------------
    
    For ii = 1 To iOrdCnt
        sTestCd = sTestCd & Trim(aIFOrdCd(ii)) & Chr(124)
    Next ii
    
'    With gOrderTable
'        For ii = 1 To gOrderTable.iOrdCnt
'            sTestCd = sTestCd & Trim(.sIFOrdCd(ii)) & Chr(124)
'        Next ii
'    End With
    
    With MAXMAT1
        .p_sID = sID
        .p_sSeq = ""
        .p_sRack = sRack
        .p_sPos = sPos
        .p_iOrdCnt = iOrdCnt    'gOrderTable.iOrdCnt
        .p_sTIFCd = sTestCd
    End With
    
End Sub

Private Sub MAXMAT1_RequestNextOrder()

    Dim sRetVal As String
    Dim sTest   As String
    
    '--- SENDORDER OK
    'Order 전송 OK 이므로 Order 성공을 화면에 표시
    Call SpdForeBack(gfIFDisplayForm.spdIntList, 3, 15, gOrderTable.iCRow, gOrderTable.iCRow, _
                RGB(0, 0, 0), 연노랑)
    '작업일련번호를 구함
'    gOrderTable.sWSeq = Format(Val(GetCurLastWSeq) + 1, "0000")
'    Call spdIntList.SetText(1, gOrderTable.iCRow, gOrderTable.sWSeq)
    Call spdIntList.SetText(2, gOrderTable.iCRow, "0")
'    'Order 내역 Local MDB에 Insert
'    Call RegOrder(1)
    '-----------------
    
    sRetVal = GetNowOrderList

    '다시 Order List를 찾는다
    If sRetVal = "NONE" Then
        ViewMsg "Order 전송이 완료되었습니다."
    ElseIf sRetVal = "OK" Then
        With gOrderTable
            MAXMAT1.p_sID = .sSampID
            MAXMAT1.p_sSeq = .sSampNo
            MAXMAT1.p_sRack = .sRack
            MAXMAT1.p_sPos = .sPos
            
            'ORDER 편집(계산식 or 중복 오더 체크...)
            sTest = Order_Check
            
            MAXMAT1.p_iOrdCnt = .iOrdCnt
            
            MAXMAT1.p_sTIFCd = sTest
        End With
    
        With MAXMAT1
            .Send_Chr (5)
            .iPhase = 3
        End With
    End If
    
End Sub

Private Sub MAXMAT1_SendOrderOK(sID As String, sRack As String, sPos As String)

    With gOrderTable
        .sSampID = sID
        .sSampNo = ""
        .sRack = sRack
        .sPos = sPos
    End With
    
    If Trim(sID) <> "" Then
        Dim ii%
        Dim vTmp1, vTmp2
        
        With spdIntList
            If .MaxRows > 0 Then
                For ii = 1 To .MaxRows
                    Call .GetText(15, ii, vTmp1)    'R
                    Call .GetText(5, ii, vTmp2)
                    
                    If Trim(vTmp1) = "N" And Trim(vTmp2) = Trim(sID) Then
                        Exit Sub
                    End If
                Next ii
            End If
        End With
        
        gOrderTable.sOther = Format(Now, "HH:MM")   'Order 전송 시간 표시(담당자 요청)...2007/5/29 yk
        
        '전송된 오더가 있는 경우 화면표시
        Call DisplayOrderOK
    Else
        '조회된 내용이 없는 경우 환자정보 구조체 초기화
        Call Init_gOrderTable
    End If
    
End Sub


Private Sub listNoOrd_Click()
    listNoOrd.ToolTipText = listNoOrd.List(listNoOrd.ListIndex)
End Sub

Private Sub listNoOrd_DblClick()
    listNoOrd.RemoveItem (listNoOrd.ListIndex)
End Sub

Private Sub listNoOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        listNoOrd.Clear
    End If
End Sub

Private Sub spdIntList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    If BlockRow = -1 And BlockRow2 = -1 Then
        giBSRow = 1
        giBERow = spdIntList.MaxRows
    Else
        giBSRow = CInt(BlockRow)
        giBERow = CInt(BlockRow2)
    End If
End Sub

Private Sub spdIntList_Change(ByVal Col As Long, ByVal Row As Long)
    Dim vRack
    Dim vPos
    
    'Rack
    If Col = 6 Then
        With spdIntList
            If giAddKey = 1 Then
                giAddKey = 0
            Else
                Call .GetText(Col, Row, vRack)
                
                If Len(vRack) <= gIFRack.sRackDigit Then
                    Call .SetText(Col, Row, Format(vRack, RackFormat(gIFRack.sRackDigit)))
                    Call .GetText(7, Row, vPos)
                    Call DisplayRackPos(Row)
                ElseIf Len(vRack) > gIFRack.sRackDigit Then
                    ViewMsgLog "위치 ERR : Rack(Tray) is over!!"
                    Exit Sub
                End If
            End If
        End With
    End If
    
    'Pos
    If Col = 7 Then
        With spdIntList
            If giAddKey = 1 Then
                giAddKey = 0
            Else
                Call .GetText(Col, Row, vPos)
                
                If IsNumeric(vPos) = False Then
                    ViewMsgLog "위치 ERR : Pos(Cup) is not number!!"
                    Exit Sub
                End If
                
                If LenH(vPos) <= gIFRack.sPosDigit Then
                    Call .SetText(Col, Row, Format(vPos, RackFormat(gIFRack.sPosDigit)))
                    Call .GetText(6, Row, vRack)
                    Call DisplayRackPos(Row)
                ElseIf LenH(vPos) > gIFRack.sPosDigit Then
                    ViewMsgLog "위치 ERR : Pos(Cup) is over!!"
                    Exit Sub
                End If
            End If
        End With
    End If
End Sub

Private Sub spdIntList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    If miLeaveCell = 1 Then
        miLeaveCell = 0
        
        Exit Sub
    End If
    
    miLeaveCell = miLeaveCell - 1
    
    Call DisplayResult2(CInt(Row))
End Sub

Private Sub spdIntList_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row = 0 Then Exit Sub

    If Col <> 5 Then
        If MsgBox("해당 Interface List를 삭제하시겠습니까?" & vbCrLf & _
            "삭제된 Interface List는 결과를 받을 수 없습니다. 계속 하시겠습니까?", _
            vbYesNo + vbQuestion, "해당 Interface List 삭제 확인") = vbYes Then
            
            With spdIntList
                .ReDraw = False
                .Row = Row
                .Action = SS_ACTION_DELETE_ROW
                .MaxRows = spdIntList.MaxRows - 1
                .ReDraw = True
            End With
        End If
    Else
        'Barcode 정보수정...2007/4/5 yk
        Dim vTmp, vRstCnt
        Dim sBarCd$, sTmp$
        
        With spdIntList
            Call .GetText(15, Row, vRstCnt)
            If Trim(vRstCnt) = "N" Then
                Exit Sub
            End If
            
            Call .SetText(0, Row, "▶" & Trim(Row))
            Call .GetText(5, Row, vTmp)
            
            sTmp = InputBox("수정할 BARCODE를 입력해 주십시요.", "BARCODE 수정", Trim(vTmp), 2800, 2600)
            
            If Len(sTmp) <> Val(gOrdCfg.sFSize(3)) Then
                Call .SetText(0, Row, Trim(Row))
                Exit Sub
            End If
            
            MousePointer = vbHourglass

            gOrderTable.sSampID = sTmp
            
            Call AppendEditBarCdResult(sTmp, Row)
            
            Call .SetText(0, Row, Trim(Row))
            
            ViewMsg "BARCODE 정보 수정이 정상적으로 완료되었습니다."
            MousePointer = vbDefault
        End With
    End If

End Sub

Private Sub AppendEditBarCdResult(ByVal sBarCd As String, ByVal iEditRow As Integer)
    On Error GoTo ErrEdit
    
    Dim vTmp
    Dim iOldCnt%, ii%, iTmpCnt%, iTmpOrdCnt%
    Dim sOldRack$, sOldPos$, sOldSeq$, sOldRerun$
    Dim sTmpIFSeq$, sTmpRstCd$
    Dim sTmpIFCd$, sTmpRst1$, sTmpRst2$, sTmpFlag$
    Dim sTmp$
    Dim aTmpData()  As String
    
    If iEditRow = 0 Then
        Exit Sub
    End If
    
    With spdIntList
        Call .GetText(6, iEditRow, vTmp): sOldRack = Trim(vTmp)
        Call .GetText(7, iEditRow, vTmp): sOldPos = Trim(vTmp)
        
        Call .GetText(16, iEditRow, vTmp): iOldCnt = Val(vTmp)
        
        For ii = 1 To iOldCnt
            Call .GetText(16 + ii, iEditRow, vTmp)
            sTmp = Trim(vTmp)
            If InStr(sTmp, Chr(124)) > 0 Then
                Erase aTmpData()
                aTmpData() = Split(sTmp & "|||", Chr(124))
                
                sTmpIFSeq = Trim(aTmpData(0))
                sTmpRstCd = ConvertIFItemInfo(8, sTmpIFSeq)     'IFSeq -> IFRstCd
                
                If sTmpRstCd <> "" Then
                    iTmpCnt = iTmpCnt + 1
                    
                    sTmpIFCd = sTmpIFCd & sTmpRstCd & Chr(124)
                    sTmpRst1 = sTmpRst1 & Trim(aTmpData(1)) & Chr(124)
                    sTmpRst2 = sTmpRst2 & Trim(aTmpData(2)) & Chr(124)
                    sTmpFlag = sTmpFlag & Trim(aTmpData(3)) & Chr(124)
                End If
            End If
        Next ii
        
        Call Order_Input("B")
        
        Call .GetText(14, iEditRow, vTmp)
        iTmpOrdCnt = Val(vTmp)
    End With
    
    If iTmpOrdCnt > 0 Then
'        '결과값 편집(소숫점 변경 처리)
'        Call ConvertResultData_BCNoEdit(iTmpCnt, sTmpIFCd, sTmpRst1, sTmpRst2)
        
        Call DisplayResultOK(3, Format(dtpLabDate.Value, "YYYYMMDD"), "", _
                            "", "", sBarCd, sOldRack, sOldPos, _
                            "", "", "", "", "", "", _
                            iTmpCnt, sTmpIFCd, sTmpRst1, sTmpRst2, _
                            "", "", sTmpFlag)
    End If
    
ErrEdit:
    If Err <> 0 Then
        ViewMsg "AppendEditBarCdResult Err - " & Err.Description
    End If
End Sub
Private Sub spdIntList_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If Row = NewRow Then Exit Sub
    If Row < 0 Then Exit Sub
    If NewRow < 0 Then Exit Sub
    
    miLeaveCell = 2
        
    Call spdIntList_Click(1, NewRow)
End Sub

Private Sub spdRst_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    
    Dim vTmp
    Dim sDisp$
    
    If Col <> 5 Or Row = 0 Then Exit Sub
    
    With spdRst
        Call .GetText(5, Row, vTmp)
        If Trim(vTmp) <> "" Then
'            sDisp = GetIFFlagInfo2(Trim(vTmp))
            sDisp = GetIFFlagInfo(Trim(vTmp))
            
            If Trim(sDisp) <> "" Then
                TipText = Trim(sDisp)
                If InStr(sDisp, vbCrLf) > 0 Then
                    MultiLine = 1
                Else
                    MultiLine = 2
                End If
                ShowTip = True
            End If
        End If
    End With
    
End Sub


Private Sub spdRst2_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

    Dim vTmp
    Dim sDisp$
    
    If Col <> 5 Or Row = 0 Then Exit Sub
    
    With spdRst2
        Call .GetText(5, Row, vTmp)
        If Trim(vTmp) <> "" Then
'            sDisp = GetIFFlagInfo2(Trim(vTmp))
            sDisp = GetIFFlagInfo(Trim(vTmp))
            
            If Trim(sDisp) <> "" Then
                TipText = Trim(sDisp)
                If InStr(sDisp, vbCrLf) > 0 Then
                    MultiLine = 1
                Else
                    MultiLine = 2
                End If
                ShowTip = True
            End If
        End If
    End With
    
End Sub

Private Sub Timer1_Timer()
'    ViewMsg ""
'    miTimerFlag = 1
End Sub

Private Sub Timer2_Timer()
'    ViewMsg ""
'    miTimerFlag = 1
End Sub

Private Sub txtBarCd_GotFocus()
    Call Txt_Highlight(txtBarCd)
End Sub

Private Sub txtBarCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtBarCd = "" Then Exit Sub
    
    If KeyCode = vbKeyReturn Then
        Screen.MousePointer = vbHourglass
        
        gOrderTable.sSampID = txtBarCd
        
        Call Get_OrderString("N")
        
        Screen.MousePointer = vbDefault
        
        Call Txt_Highlight(txtBarCd)
    End If
End Sub

Private Sub txtBarCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtOrdNo_Click()
    Call Txt_Highlight(txtOrdNo)
End Sub

Private Sub txtOrdNo_GotFocus()
    Call Txt_Highlight(txtOrdNo)
End Sub

Private Sub txtOrdNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdSendOrd.SetFocus
    End If
End Sub

Private Sub txtOrdNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRack_GotFocus()
    Call Txt_Highlight(txtRack)
End Sub

Private Sub txtRack_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtPos.SetFocus
    End If
End Sub

Private Sub txtRack_LostFocus()
    'Integra400, Integra700
    If txtRack.MaxLength = 3 Then
        If Len(txtRack) < 3 Then
            txtRack = Format(txtRack, "000")
        End If
    'Integra800
    ElseIf txtRack.MaxLength = 4 Then
        If Len(txtRack) < 4 Then
            txtRack = Format(txtRack, "0000")
        End If
    End If
End Sub

Private Sub txtRack_Validate(Cancel As Boolean)
    'Integra400, Integra700
    If txtRack.MaxLength = 3 Then
        If Len(txtRack) < 3 Then
            txtRack = Format(txtRack, "000")
        End If
    'Integra800
    ElseIf txtRack.MaxLength = 4 Then
        If Len(txtRack) < 4 Then
            txtRack = Format(txtRack, "0000")
        End If
    End If
End Sub

Private Sub txtPos_GotFocus()
    Call Txt_Highlight(txtPos)
End Sub

Private Sub txtPos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtOrdNo.SetFocus
    End If
End Sub

Private Sub txtPos_KeyPress(KeyAscii As Integer)
    Call TxtTypeOnlyNumeric(txtPos, KeyAscii)
End Sub

Private Sub txtPos_LostFocus()
    If Len(txtPos) < 2 Then
        txtPos = Format(txtPos, "00")
    End If
End Sub

Private Sub txtPos_Validate(Cancel As Boolean)
    If Len(txtPos) < 2 Then
        txtPos = Format(txtPos, "00")
    End If
End Sub

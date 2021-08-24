VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmComm 
   Caption         =   "인터페이스"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   FillStyle       =   0  '단색
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   15210
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdStartNo 
      Caption         =   "시작번호변경"
      Height          =   375
      Left            =   7020
      TabIndex        =   40
      Top             =   45
      Width           =   2100
   End
   Begin VB.CommandButton WorkList 
      Caption         =   "WorkList 불러오기"
      Height          =   375
      Left            =   4905
      TabIndex        =   25
      Top             =   45
      Width           =   2100
   End
   Begin VB.TextBox txtBarcode 
      Height          =   375
      Left            =   2700
      TabIndex        =   24
      Top             =   45
      Width           =   2175
   End
   Begin FPSpreadADO.fpSpread spdResult1 
      Height          =   4875
      Left            =   90
      TabIndex        =   19
      Top             =   480
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   8599
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   17
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      ShadowColor     =   14737632
      SpreadDesigner  =   "frmComm.frx":0000
   End
   Begin VB.CommandButton cmdACK 
      Caption         =   "ACK"
      Height          =   315
      Left            =   8820
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fraCmdBar 
      Height          =   705
      Left            =   60
      TabIndex        =   5
      Top             =   9240
      Width           =   15075
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Run"
         Height          =   375
         Index           =   0
         Left            =   9540
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   9
         Top             =   210
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Stop"
         Height          =   375
         Index           =   1
         Left            =   10860
         TabIndex        =   8
         Top             =   210
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clear"
         Height          =   375
         Index           =   2
         Left            =   12180
         TabIndex        =   7
         Top             =   210
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exit"
         Height          =   375
         Index           =   3
         Left            =   13500
         TabIndex        =   6
         Top             =   210
         Width           =   1245
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "작업대기 중.."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   900
         TabIndex        =   11
         Top             =   315
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   " 상태 :"
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
         Left            =   150
         TabIndex        =   10
         Top             =   315
         Width           =   615
      End
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   14070
      Top             =   9750
   End
   Begin VB.Timer tmrReceive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   13650
      Top             =   9750
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   12510
      Top             =   8940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":07B3
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":0811
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":086F
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":08CD
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":092B
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":0989
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm comEQP 
      Left            =   14490
      Top             =   8850
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin MSComctlLib.ImageList imlStatus 
      Left            =   13080
      Top             =   9750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":09E7
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":0F81
            Key             =   "NOT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":151B
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":1AB5
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":2347
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":24A1
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":25FB
            Key             =   "NOF"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwCuData 
      Height          =   3420
      Left            =   4740
      TabIndex        =   3
      Top             =   435
      Visible         =   0   'False
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   6033
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin Threed.SSFrame FrameResult 
      Height          =   3525
      Left            =   90
      TabIndex        =   2
      Top             =   5460
      Width           =   7995
      _Version        =   65536
      _ExtentX        =   14102
      _ExtentY        =   6218
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
      Begin FPSpreadADO.fpSpread spdRstDetail 
         Height          =   3195
         Left            =   150
         TabIndex        =   20
         Top             =   210
         Width           =   7695
         _Version        =   393216
         _ExtentX        =   13573
         _ExtentY        =   5636
         _StockProps     =   64
         BackColorStyle  =   1
         EditModePermanent=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   6
         MaxRows         =   10
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   11589855
         SpreadDesigner  =   "frmComm.frx":2755
         ScrollBarTrack  =   1
      End
   End
   Begin Threed.SSFrame FrameInterface 
      Height          =   3525
      Left            =   90
      TabIndex        =   0
      Top             =   5460
      Width           =   7995
      _Version        =   65536
      _ExtentX        =   14102
      _ExtentY        =   6218
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
      Begin VB.TextBox txtCom 
         Height          =   2895
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "frmComm.frx":2C18
         Top             =   480
         Width           =   7725
      End
      Begin VB.Label Label4 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "인터페이스 내역"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   150
         TabIndex        =   4
         Top             =   210
         Width           =   4785
      End
   End
   Begin Threed.SSFrame FrameError 
      Height          =   3525
      Left            =   8220
      TabIndex        =   21
      Top             =   5490
      Width           =   6885
      _Version        =   65536
      _ExtentX        =   12144
      _ExtentY        =   6218
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
      Begin VB.ListBox lstErr 
         Height          =   2940
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   6675
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "오류 내역"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   150
         TabIndex        =   23
         Top             =   240
         Width           =   4785
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   6810
      Left            =   7020
      TabIndex        =   26
      Top             =   45
      Visible         =   0   'False
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
      _ExtentY        =   12012
      _StockProps     =   15
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel4 
         Height          =   285
         Left            =   90
         TabIndex        =   36
         Top             =   90
         Width           =   5595
         _Version        =   65536
         _ExtentX        =   9869
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "업무나열서 불러오기"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   555
         Left            =   90
         TabIndex        =   27
         Top             =   405
         Width           =   5595
         _Version        =   65536
         _ExtentX        =   9869
         _ExtentY        =   979
         _StockProps     =   15
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSCommand cmdSelect 
            Height          =   375
            Left            =   4365
            TabIndex        =   35
            Top             =   90
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "조회"
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   315
            Left            =   1080
            TabIndex        =   32
            Top             =   135
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   56688641
            CurrentDate     =   37285
         End
         Begin MSComCtl2.DTPicker dtpDate1 
            Height          =   315
            Left            =   2790
            TabIndex        =   33
            Top             =   135
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   56688641
            CurrentDate     =   37285
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "-"
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
            Left            =   2610
            TabIndex        =   39
            Top             =   180
            Width           =   120
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "조회일자 :"
            Height          =   180
            Index           =   6
            Left            =   135
            TabIndex        =   34
            Top             =   180
            Width           =   840
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   5280
         Left            =   90
         TabIndex        =   28
         Top             =   990
         Width           =   5595
         _Version        =   65536
         _ExtentX        =   9869
         _ExtentY        =   9313
         _StockProps     =   15
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSCommand cmdSel 
            Height          =   360
            Index           =   1
            Left            =   270
            TabIndex        =   30
            Top             =   0
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   644
            _StockProps     =   78
            BevelWidth      =   1
            Picture         =   "frmComm.frx":2C1E
         End
         Begin Threed.SSCommand cmdSel 
            Height          =   360
            Index           =   0
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   644
            _StockProps     =   78
            BevelWidth      =   1
            Picture         =   "frmComm.frx":30A0
         End
         Begin FPSpreadADO.fpSpread spdWorkList 
            Height          =   5235
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   5595
            _Version        =   393216
            _ExtentX        =   9869
            _ExtentY        =   9234
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridShowHoriz   =   0   'False
            GridSolid       =   0   'False
            MaxCols         =   5
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmComm.frx":350E
         End
      End
      Begin Threed.SSCommand cmdInsert 
         Height          =   375
         Left            =   3420
         TabIndex        =   37
         Top             =   6345
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "등록"
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   375
         Left            =   4590
         TabIndex        =   38
         Top             =   6345
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "취소"
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Receive :"
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
      Index           =   0
      Left            =   13740
      TabIndex        =   18
      Top             =   120
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Send : "
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
      Left            =   12705
      TabIndex        =   17
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Port : "
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
      Left            =   11670
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgReceive 
      Height          =   240
      Left            =   14730
      Picture         =   "frmComm.frx":39AF
      Top             =   90
      Width           =   240
   End
   Begin VB.Image imgSend 
      Height          =   240
      Left            =   13410
      Picture         =   "frmComm.frx":3F39
      Top             =   90
      Width           =   240
   End
   Begin VB.Image imgPort 
      Height          =   240
      Left            =   12270
      Picture         =   "frmComm.frx":44C3
      Top             =   90
      Width           =   240
   End
   Begin VB.Label Label6 
      Appearance      =   0  '평면
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  '단일 고정
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11520
      TabIndex        =   15
      Top             =   30
      Width           =   3615
   End
   Begin VB.Label lblSubMenu 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "VitrosEci  Interface"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   135
      TabIndex        =   13
      Top             =   90
      Width           =   2460
   End
   Begin VB.Label picSubMenu 
      Appearance      =   0  '평면
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  '단일 고정
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   90
      TabIndex        =   14
      Top             =   30
      Width           =   2565
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const COL_KEY       As String = "K"
Private Const COL_EQP_NUM   As String = "EQP_ID"

Private Const KEY_SEQ       As String = "KEY_SEQ"   ' "순서"
Private Const KEY_PTID      As String = "KEY_PTID"  ' "등록번호"
Private Const KEY_PTNM      As String = "KEY_PTNM"  ' "성  명"
Private Const KEY_SPCNO     As String = "KEY_SPCNO" ' "검체번호"
Private Const KEY_EQPNO     As String = "KEY_EQPNO" ' "검체번호"
Private Const KEY_STAT      As String = "KEY_STAT"  ' "상 태"
Private Const KEY_TEST      As String = "KEY_TEST"  ' "검사항목"

Private Const TEST_NM_EQP   As String = "EQP_NM"    '장비 코드
Private Const TEST_CD_LIS   As String = "LIS_CD"    '검사실 코드
Private Const TEST_NM_LIS   As String = "LIS_NM"    '검사실 이름
Private Const TEST_VALUES   As String = "VALUES"    '결과

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const FS  As String = ""
Const RS  As String = ""

Public WithEvents Result As clsMsg_Result
Attribute Result.VB_VarHelpID = -1
Public WithEvents Order  As clsMsg_Query
Attribute Order.VB_VarHelpID = -1
Public Result1 As clsResult
Attribute Result1.VB_VarHelpID = -1

Private mAdoRs      As ADODB.Recordset
Private mAdoRs1     As ADODB.Recordset
Private CallForm    As String
Private IS_SET      As Boolean

Private f_strBuffer     As String
Private f_strJOB_FLAG   As String
Private f_strOrdList    As String
Private f_intSampleNo   As Integer

Private f_blnWorkList   As Boolean
Private f_lngWork_Row   As Long
Dim ReceiveData      As String
Dim SendFlg          As Boolean
Dim Patiant_Recevid As Integer

Private MSG_STX     As String
Private MSG_ETX     As String
Private MSG_ENQ     As String
Private MSG_EOT     As String
Private MSG_ACK     As String
Private MSG_NAK     As String
Private MSG_CR      As String
Private MSG_LF      As String
Private MSG_CRLF    As String

Private Type typeNOVA
    TestDate      As String
    TestTime      As String
    RunType       As String 'N, E, R, C, S, B
    SampleNo      As String
    SID           As String 'Sid
    SampleTy      As String '1~5
    RackNo        As String
    Position      As String '1~5
    Priority      As String
    TestId(100)   As String
    Result(100)   As String
    Status(100)   As String
    Rerun(100)    As String
End Type

Dim NOVA As typeNOVA
Dim flgETB As Boolean
Dim fAxsym(100) As String
Dim fAxsymCfg(100) As Integer
Dim fAxsymSize(100, 1) As Integer

Dim fChannel() As String

Dim SeqNo As String

Private Type TYPE_CD
    strEqpCd    As String
    intCnt      As Integer
    strTestcd(100) As String
End Type

Private f_typCode() As TYPE_CD

Dim OrderCnt As Integer
Dim SendCount As Integer

Dim CountTest As Integer, sErrorFlag As Boolean
Dim cntCheckSum      As Integer
Dim flgETX           As Boolean

Private Type typeAXSYM
    TestDate      As String
    TestTime      As String
    RunType       As String 'N, E, R, C, S, B
    SampleNo      As String
    SID           As String 'Sid
    SampleTy      As String '1~5
    RackNo        As String
    Position      As String '1~5
    Priority      As String
    TestId(100)   As String
    Result(100)   As String
    Status(100)   As String
    Rerun(100)    As String
End Type

Dim AXSYM As typeAXSYM

Const Field_      As String = "|"
Const Repeat_     As String = "\"
Const Component_  As String = "^"
Const Escape_     As String = "&"
Const Slash_      As String = "/"
Dim cntField_     As Integer '|
Dim cntRepeat_    As Integer '\
Dim cntComponent_ As Integer '^
Dim cntEscape_    As Integer '&
Dim cntSlash_     As Integer '/

Dim fAXSYM_1(100) As String
'Dim SendData(10)     As String
Dim SendData     As String
Dim HostOutput       As String

Dim phase  As Integer
Dim bufcnt As Integer
Dim State  As String
Dim SndPhase As Integer
Dim FrameNo As Integer

Private strRcvbufR As String

'------------------
Dim cInterface As New clsIInterface ' Interface Class

Private Function f_funGet_ConvertResult(ByVal strRstval As String) As String

    Dim intPos  As Integer
    Dim strTmp1 As String, strTmp2  As String
    
    intPos = InStr(strRstval, "E")
    If intPos > 0 Then
        strTmp1 = Mid$(strRstval, 1, intPos - 1)
        strTmp2 = Mid$(strRstval, intPos + 1)
        
        If Mid$(strTmp2, 1, 1) = "-" Then
            f_funGet_ConvertResult = Round(Val(strTmp1) * (0.1 ^ Val(Mid$(strTmp2, 2))), 2)
        Else
            f_funGet_ConvertResult = Round(Val(strTmp1) * (10 ^ Val(Mid$(strTmp2, 2))), 2)
        End If
    Else
        f_funGet_ConvertResult = strRstval
    End If
    
End Function

Private Function MakeCS(Source As String) As String
    Dim X      As Long
    Dim ChkCS  As String
    Dim SumCS  As String
    Dim AddCS  As Long
    
    For X = 1 To Len(Source)
        AddCS = AddCS + Asc(Mid(Source, X, 1))
    Next X
    AddCS = AddCS + Asc(Chr(13)) + Asc(ETX)
    AddCS = AddCS Mod &H100
    SumCS = Hex(AddCS)
    If Len(SumCS) = 1 Then
        ChkCS = "0" & SumCS
    Else
        ChkCS = Mid(SumCS, Len(SumCS) - 1, 1)
        ChkCS = ChkCS & Right(SumCS, 1)
    End If
    MakeCS = ChkCS
End Function

Private Function f_funGet_SpreadRow(ByVal objSpd As fpSpread, ByVal intCol As Integer, _
                                    ByVal strPara As String) As Integer

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    f_funGet_SpreadRow = 0
    
    With objSpd
        For intRow = 1 To .MaxRows
            .GetText intCol, intRow, varTmp
            If Trim$(varTmp) = strPara Then
                f_funGet_SpreadRow = intRow
                Exit For
            End If
        Next
    End With
    
End Function

Private Sub f_subSet_ComCharacter()

    MSG_STX = Chr(COM_STX)
    MSG_ETX = Chr(COM_ETX)
    MSG_ENQ = Chr(COM_ENQ)
    MSG_EOT = Chr(COM_EOT)
    MSG_ACK = Chr(COM_ACK)
    MSG_NAK = Chr(COM_NACK)
    MSG_CR = Chr(COM_CR)
    MSG_LF = Chr(COM_LF)
    MSG_CRLF = Chr(COM_CR) & Chr(COM_LF)
    
End Sub


Private Sub f_subSet_ItemList()

    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strTest As String, intPos   As Integer
    Dim strTmp  As String, intCol   As Integer, intCol2   As Integer, intRow  As Integer, intCnt  As Integer
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear
    
    intRow = 1
    intCol = 9
    intCol2 = 1
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdRstDetail
        .MaxRows = 10
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFLM, REFHM, REFLF, REFHF, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") " & _
             "   and ((TESTCD <> '') and (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
             
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst:        ReDim fChannel(adoRS.RecordCount + intCol)
    
    Do While Not adoRS.EOF
        Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TEST_EQP") & ""), , "LST")
            itemX.SubItems(1) = Trim(adoRS.Fields("TESTCD") & "")
            itemX.SubItems(2) = Trim(adoRS.Fields("TESTNM") & "")
            itemX.SubItems(3) = ""
            itemX.SubItems(4) = Trim(adoRS.Fields("DELTA") & "")
            itemX.SubItems(5) = Trim(adoRS.Fields("DELTAGBN") & "")
            itemX.SubItems(6) = Trim(adoRS.Fields("PANICL") & "")
            itemX.SubItems(7) = Trim(adoRS.Fields("PANICH") & "")
            itemX.SubItems(8) = Trim(adoRS.Fields("REFLM") & "")
            itemX.SubItems(9) = Trim(adoRS.Fields("REFHM") & "")
            itemX.SubItems(10) = Trim(adoRS.Fields("REFLF") & "")
            itemX.SubItems(11) = Trim(adoRS.Fields("REFHF") & "")
            itemX.SubItems(12) = Trim(adoRS.Fields("AUTOVERIFY") & "")
            itemX.SubItems(13) = Trim(adoRS.Fields("REMARK") & "")
            itemX.Tag = Trim(adoRS.Fields("TEST_EQP") & "")
            itemX.Text = Trim(adoRS.Fields("TESTCD") & "")
        Set itemX = Nothing
        
        With spdResult1
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        With spdRstDetail
            If intRow > .MaxRows Then
                intRow = 1
                intCol2 = intCol2 + 2
            End If
            
            .SetText intCol2, intRow, Trim$(adoRS("TESTNM") & "")
            intRow = intRow + 1
            
        End With
        
        intCol = intCol + 1
        
        adoRS.MoveNext
    Loop
    
    Set adoRS = Nothing

Exit Sub
ErrRoutine:
    Set adoRS = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub cmdACK_Click()
    Dim RecData As String
    Dim sRs As Object
    
    'Call f_subSet_WorkList
    'Call f_subSet_TestList("10124")
    
    Set sRs = f_subSet_TestList("10124")
    
    If Not sRs.EOF Then
        Debug.Print Trim(sRs("검체번호")) & ""
        Debug.Print Trim(sRs("품목코드")) & ""
    End If
    
    Exit Sub
    'Call COM_OUTPUT(ACK)
          RecData = "|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|2004031714235384"
RecData = RecData & "1H|\^&|||AxSYM^5.00^18795^H1P1O1R1C1Q1L1M1|||||||P|1|2004031714255083"
RecData = RecData & "2Q|1|^9030||^^^ALL||||||||O00"
RecData = RecData & "3L|13C"
RecData = RecData & ""


    Call ComReceive(RecData)
    
End Sub

Private Sub cmdAction_Click(Index As Integer)
    
    Select Case Index
        Case 0:     Call cmdRun
        Case 1:     Call cmdStop
        Case 2:     Call cmdClear
        Case 3:     Call cmdExit
        Case Else
    End Select

End Sub

Private Sub cmdClear()
    
    txtCom.Text = ""
    lstErr.Clear
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With

End Sub

Private Sub cmdExit()
    
    If frmComm.comEQP.PortOpen = True Then
        If MsgBox("인터페이스중입니다." & Chr(10) & _
               "작업을 종료하면 받고있거나 검사중인 데이터를 잃게 됩니다" & Chr(10) & _
               "종료하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
            
            Unload Me
            
        End If
    Else
        Unload Me
    End If

End Sub

Private Sub cmdRun()
    
    Dim itemX As ListItem
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub cmdRun()"
    
    If Not comEQP.PortOpen Then comEQP.PortOpen = True
    If comEQP.PortOpen Then
        Call ShowMessage("연결 되었습니다.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "작업중.."
    Else
        Call ShowMessage("연결 되지 않았습니다.")
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "작업 대기중.."
    End If
        
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
End Sub

Private Sub cmdStop()
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub cmdRun()"
    
    If comEQP.PortOpen Then comEQP.PortOpen = False
    If comEQP.PortOpen Then
        Call ShowMessage("중지 되지 않았습니다.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "작업중.."
    Else
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "작업 대기중.."
    End If
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
End Sub

Private Sub cmdCancel_Click()
    SSPanel1.Visible = False
End Sub

Private Sub cmdInsert_Click()

    Dim varTmp  As Variant
    Dim introw1 As Integer, intRow2 As Integer
    Dim intIdx  As Integer
    Dim Rev     As Long
    Dim Test_Cd() As String, strPid()   As String, strPnm() As String
    Dim itemX As ListItem
    Dim blnFlag As Boolean
    Dim strBarno    As String, strSPid  As String, strSPnm   As String
    
    Dim strEqpCd    As String
       
    blnFlag = False
    With spdWorkList
        For introw1 = 1 To .MaxRows
            .GetText 1, introw1, varTmp
            If Trim$(varTmp) = "1" Then
                .GetText 2, introw1, varTmp:    strSPnm = Trim$(varTmp)
                .GetText 3, introw1, varTmp:    strSPid = Trim$(varTmp)
                .GetText 4, introw1, varTmp:    strBarno = Trim$(varTmp)
                
                intRow2 = f_funGet_SpreadRow(spdResult1, 2, strBarno)
                If intRow2 < 1 Then
                    intRow2 = f_funGet_SpreadRow(spdResult1, 2, "")
                    If intRow2 < 1 Then
                        spdResult1.MaxRows = spdResult1.MaxRows + 1
                        spdResult1.RowHeight(spdResult1.MaxRows) = 12
                        intRow2 = spdResult1.MaxRows
                    End If
                    
                    blnFlag = False
                    Set mAdoRs = f_subSet_TestList(strBarno)
                    If Len(strBarno) = 12 Then
                        Do Until mAdoRs.EOF
                            strEqpCd = f_funGet_CODE(mAdoRs("g15_worknm"))
                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then
                                blnFlag = True
                                spdResult1.Row = intRow2
                                spdResult1.Col = itemX.Index + 6
                                spdResult1.BackColor = &HC6FEFF '&H80C0FF
                                DoEvents
                            End If
                            mAdoRs.MoveNext
                        Loop
                    End If
                    If blnFlag = True Then
                        spdResult1.SetText 2, intRow2, strBarno
                        spdResult1.SetText 3, intRow2, strSPnm
                        spdResult1.SetText 4, intRow2, strSPid
                    Else
                        spdResult1.MaxRows = spdResult1.MaxRows - 1
                    End If
                End If
                spdResult1.SetText 1, intRow2, "1"
                spdResult1.MaxRows = intRow2

                .SetText 1, introw1, ""
            End If
        Next
    End With
    
    Dim aROW    As Integer, aCOL   As Integer
    Dim varChk  As Variant, varBar As Variant, varNum As Variant
    Dim iRow    As Integer, iCnt   As Integer
    
    With spdResult1
        iCnt = 0
        .GetText 1, 1, varChk
        .GetText 2, 1, varBar
        varNum = 1
        If Trim(varChk) = "1" And Trim(varBar) <> "" Then
            For iRow = 1 To .MaxRows
                .SetText 5, iRow, varNum
                .SetText 6, iRow, ((iCnt Mod 10) + 1) - 1
                iCnt = iCnt + 1
                If (iCnt Mod 10) = 0 Then varNum = varNum + 1
            Next
        End If
    End With
    
End Sub

Private Sub cmdSelect_Click()

    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Privete sub cmdWorkQuery_Click()"
    
    Dim strKeyno    As String
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String, strBarno()   As String
    Dim strTestcd() As String, strTPid()    As String, strTPnm() As String
    Dim strEqpCd    As String
    Dim intRow  As String, intIdx  As Integer, intCol   As Integer
    Dim itemX   As ListItem
       
    '-- WorkList조회
    Set mAdoRs = f_subSet_WorkList(mskOrdDate.Text, mskOrdDate1.Text)
    
    With spdWorkList
        .MaxRows = 0
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    intRow = 0
    Do Until mAdoRs.EOF
        intIdx = 0
        With spdWorkList
            If strKeyno <> mAdoRs.Fields("g13_sample") Then
                intRow = intRow + 1
                If intRow > .MaxRows Then .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
                
                .SetText 1, intRow, "1"
                If optSpc1.Value = True Then
                    .SetText 2, intRow, mAdoRs("H11_sjjnam")
                Else
                    .SetText 2, intRow, mAdoRs("psn_sjjnam")
                End If
                
                .SetText 3, intRow, mAdoRs("g13_chtno")
                .SetText 4, intRow, mAdoRs("g13_sample")

                '-- 검사항목조회
                Set mAdoRs1 = New Recordset
                Set mAdoRs1 = f_subSet_TestList(mAdoRs("g13_sample"))
                
                Do Until mAdoRs1.EOF
                    strEqpCd = f_funGet_CODE(mAdoRs1("g15_worknm"))
                    
                    Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                    If Not itemX Is Nothing Then .SetText 4 + itemX.Index, intRow, "V"
                    Set itemX = Nothing
                    mAdoRs1.MoveNext
                Loop
            End If
            strKeyno = mAdoRs("g13_sample")
        End With
        intIdx = intIdx + 1
        mAdoRs.MoveNext
    Loop
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)

End Sub

Private Sub comEQP_OnComm()
    
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim strDta      As String
    Dim Arr()       As Byte
    Dim strdata     As String
    Dim fRcvString As String
    
    Dim Buffer As String
    Dim iBufLen As Integer
    Dim BufChar As String
    Dim i%

    On Error GoTo ErrorTrap
    
    CallForm = "frmInterface - Private Sub comEQP_OnComm()"
    
    Select Case comEQP.CommEvent
        Case comEvReceive
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If

            
            Buffer = comEQP.Input
            iBufLen = Len(Buffer)
            For i = 1 To iBufLen
                BufChar = Mid(Buffer, i, 1)
                Select Case cInterface.phase
                    Case 1                  ' ENQ 대기
                        Select Case Asc(BufChar)
                            Case 5          ' ENQ
                                'comEQP.Output = Chr(6)
                                Call COM_OUTPUT(Chr(6))
                                cInterface.phase = 2
                                cInterface.bufcnt = 1
                       End Select
                    Case 2                  ' LF 대기
                        Select Case Asc(BufChar)
                            Case 2          ' STX
                                Call cInterface.clearRcvbuf
                            Case 4          ' EOT
                                If cInterface.State = "Q" Then
                                    'comEQP.Output = Chr(5)
                                    Call COM_OUTPUT(Chr(5))
                                    cInterface.Snd_Phase = 1
                                    cInterface.FrameN = 1
                                End If
                                cInterface.phase = 3
                            Case 10        ' LF
                                Call psDataDefine(Buffer, fChannel(), spdResult1)
                                'comEQP.Output = Chr(6)
                                Call COM_OUTPUT(Chr(6))
                                cInterface.phase = 2
                            Case Else
                                Call cInterface.addRcvbuf(BufChar)
                        End Select
                    Case 3                  ' ACK 대기
                        Select Case Asc(BufChar)
                            Case 6
                                If cInterface.State = "Q" Then
                                    Call SendOrdData
                                End If
                            Case 5
                                'comEQP.Output = Chr(6)
                                Call COM_OUTPUT(Chr(6))
                                cInterface.phase = 2
                            Case 21
                                'comEQP.Output = Chr(5)
                                Call COM_OUTPUT(Chr(5))
                            Case 4
                                cInterface.phase = 1
                        End Select

                End Select
            Next
        
        Case comEvSend
        
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
        Case comEvCTS
            strEVMsg = " CTS(Clear to Send) 변경 감지"
        Case comEvDSR
            strEVMsg = " DSR(Data Set Read) 변경 감지"
        Case comEvCD
            strEVMsg = " CD(Carrier Detecr) 변경 감지"
        Case comEvRing
            strEVMsg = " 전화 벨이 울리는 중"
        Case comEvEOF
            strEVMsg = " EOF(End Of File) 감지"

        ' 오류 메시지
        Case comBreak
            strERMsg = " 중단 신호 수신"
        Case comCDTO
            strERMsg = " 반송파 검출 시간 초과"
        Case comCTSTO
            strERMsg = " CTS(Clear to Send) 시간 초과"
        Case comDCB
            strERMsg = " 포트에 대한 장치 제어 블록(DCB) 검색 중 예기치 못한 오류"
        Case comDSRTO
            strERMsg = " DSR(Data Set Read) 시간 초과"
        Case comFrame
            strERMsg = " 프레이밍 오류"
        Case comOverrun
            strERMsg = " 패리티 오류"
        Case comRxOver
            strERMsg = " 수신 버퍼 초과"
        Case comRxParity
            strERMsg = " 패리티 오류"
        Case comTxFull
            strERMsg = " 전송 버퍼에 여유가 없음"
        Case Else
            strERMsg = " 알 수 없는 오류 또는 이벤트"
    End Select
    If Len(strERMsg) > 0 Then Call ShowMessage(strERMsg)
    
    Exit Sub
ErrorTrap:
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub SendOrdData()
    Dim tmp     As String
    Dim ChkS    As String
    Dim LabDate As String
    Dim TestDat As String
    Dim i       As Integer
    Dim sSampleNo As String
    Dim sRs As Object
    Dim intRow  As Integer
    Dim itemX As ListItem
    Dim strEqpCd As String
    Dim strOrder As String
    Dim sndOrder As String
    
    On Error GoTo ErrorTrap
    
    CallForm = "frmInterface - Private Sub SendOrdData()"
    
    Select Case cInterface.Snd_Phase
        Case 1      ' Header Record
            Debug.Print "----> H"
            tmp = cInterface.FrameN & "H|\^&||||||||||P|1|" & vbCr & Chr(3)
            cInterface.Snd_Phase = 2
        Case 2      ' Patient Record
            Debug.Print "----> P"
            tmp = cInterface.FrameN & "P|1||" & Trim$(AXSYM.SampleNo) & "||" & vbCr & Chr(3)
            cInterface.Snd_Phase = 3
        Case 3      ' Order Record
            Debug.Print "----> O"
            sSampleNo = AXSYM.SampleNo
        
            With spdResult1
                Set sRs = f_subSet_TestList(Trim(sSampleNo))
                If Not sRs.EOF Then
                    .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
                    intRow = SeqNullSearch(spdResult1, sRs("검체번호"), 1)
                    
                    .SetText 1, intRow, Trim(sRs("검체번호")) & ""
                    .SetText 4, intRow, Trim(sRs("성명")) & ""
                    .SetText 5, intRow, IIf(Trim(sRs("성별코드")) = "M", "남", "여") & "/" & Format(Now, "yyyy") - Format(Trim(sRs("생년월일")), "yyyy") - 1
                    If sRs("입외구분") = "1" Then
                        .SetText 6, intRow, "입원"
                    ElseIf sRs("입외구분") = "2" Then
                        .SetText 6, intRow, "외래"
                    Else
                        .SetText 6, intRow, "퇴원"
                    End If
                    .SetText 7, intRow, Trim(sRs("병록번호")) & ""
                    .SetText 8, intRow, Trim(sRs("접수일자")) & ""
                    '-- 검사항목조회
                    Do Until sRs.EOF
                        strEqpCd = Trim(sRs("품목코드"))
                        Set itemX = lvwCuData.FindItem(strEqpCd, lvwSubItem, , lvwWhole)
        
                        If Not itemX Is Nothing Then
                            spdResult1.Row = intRow
                            spdResult1.Col = itemX.Index + 8
                            spdResult1.BackColor = &HC6FEFF '&H80C0FF
                            DoEvents
                        
                            If strOrder = "" Then
                                strOrder = "^^^" & Trim(itemX.Tag)
                            Else
                                strOrder = strOrder & "\^^^" & Trim(itemX.Tag)
                            End If
                        
                        End If
                        
                        sRs.MoveNext
                    Loop
                    tmp = cInterface.FrameN & "O|1|" & Trim$(AXSYM.SampleNo) & "||" & strOrder & "|R||||||N||||||||||||||Q" & vbCr & Chr(3)
                    'Debug.Print tmp
                End If
            End With
            cInterface.Snd_Phase = 4
        Case 4      'Terminator Record
            Debug.Print "----> L"
            tmp = cInterface.FrameN & "L|1" & vbCr & Chr(3)
            cInterface.Snd_Phase = 5
            Debug.Print tmp
        Case 5      ' EOT
            Debug.Print "----> EOT"
            'comEQP.Output = Chr(4)   'EOT
            Call COM_OUTPUT(Chr(4))
            cInterface.FrameN = 1
            cInterface.phase = 1
            cInterface.Snd_Phase = 1
            cInterface.State = ""
            Exit Sub
    End Select
    
    ChkS = getChkSum(tmp)
    sndOrder = Chr(2) & tmp & ChkS & vbCr & vbLf
    Call COM_OUTPUT(sndOrder)
    cInterface.addFrameN
        
    Exit Sub
ErrorTrap:
    Call ErrMsgProc(CallForm)

End Sub

Private Sub ComReceive(ByRef RecData As String)

    Dim sStxCheck As Integer, sEnqCheck As Integer, sEtxCheck As Integer
    Dim sLfCheck As Integer, sCrcheck As Integer, ii As Integer
    Dim MHead As String, Pinfo As String, OutputData As String, com_sTemp As String
    Dim strRec  As String, strBuff  As String
    Dim strTmp  As String, intIdx   As Integer
    
    Static OrgMsg As String

    strRec = RecData
    Print #1, strRec;
    Call COM_INPUT(strRec)
'    Debug.Print strRec
    
    For ii = 1 To Len(strRec)
        Select Case Mid(strRec, ii, 1)
            Case STX:
                    ii = ii + 1 'Frame Number
                    If sErrorFlag Then
                        sErrorFlag = False
                        cntCheckSum = 2
                    Else
                        cntCheckSum = 0
                    End If
            Case ETX:
                    cntCheckSum = cntCheckSum + 1
                    Call COM_OUTPUT(ACK)
'                        Debug.Print "[HOST] " & ACK
                    flgETX = True
            Case ETB:
                    If Mid(ReceiveData, ii, 2) = vbCr & vbLf Then
                        ReceiveData = left(ReceiveData, Len(ReceiveData) - 2) 'Remove CR & LF
                    End If
                    cntCheckSum = cntCheckSum + 1
                    Call COM_OUTPUT(ACK)
'                        Debug.Print "[HOST] " & ACK
                    flgETB = True
                    sErrorFlag = True
            Case vbCr:
                    If flgETB = True Then
                        flgETB = False
                    Else
'                            ReceiveTheDataAXSYM
                        '---------------------------------------------
'                        Dim sTxfile As String
'
'                        sTxfile = App.Path & "\" & Format(Now, "yyyyMMdd") & ".LOG"
'                        If Len(Dir(sTxfile)) = 0 Then
'                            Open sTxfile For Output As #1
'                            Close #1
'                        End If
'                        Open sTxfile For Append As #1
'                            Print #1, "RCV=> "; ReceiveData
'                        Close #1
                        '---------------------------------------------
                        Call psDataDefine(ReceiveData, fChannel(), spdResult1)
                        GoSub ClearReceiveData
                    End If
            Case vbLf:
                '
            Case ENQ:
                    Call COM_OUTPUT(ACK)
            Case ACK:
                    If SendFlg = True Then
'                            SendTest
                        Call SendTest(ReceiveData, fChannel(), spdResult1)
                    Else
                        Call COM_OUTPUT(EOT)
'                           Debug.Print "[HOST] " & EOT
                        AXSYM.SID = ""
                    End If
            Case NAK:
                    If AXSYM.SID <> "" Then
                        SendCount = SendCount - 1
                        Call SendTest(ReceiveData, fChannel(), spdResult1)
                    Else
                        Call COM_OUTPUT(EOT)
                    End If
            Case EOT:
                    Call COM_OUTPUT(ENQ)
                    cntCheckSum = 0
                    GoSub ClearReceiveData
            Case Else:
                Select Case cntCheckSum
                    Case 1:
                        cntCheckSum = cntCheckSum + 1
                    Case 2:
                        cntCheckSum = 0
                    Case Else:
                        ReceiveData = ReceiveData & Mid(strRec, ii, 1)
                End Select
                
        End Select
    Next ii
    
    Exit Sub
    
ClearReceiveData:
    ReceiveData = ""
    cntField_ = 0
    cntRepeat_ = 0
    cntComponent_ = 0
    cntEscape_ = 0
    cntSlash_ = 0
    Return
    
errOnComm:
    Exit Sub
            
End Sub


Private Sub SendTest(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object)
    If SendCount <= 0 Then
'        SendTheSample
        Call SendTheSample(ReceiveData, fChannel(), spdResult1)
    Else
        Call COM_OUTPUT(SendData)
        SendCount = SendCount - 1
        If SendCount = 0 Then
           SendFlg = False
        Else
           SendFlg = True
        End If
    End If
End Sub

Private Sub SendTheSample(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object)
    On Error GoTo SendTheSampleSub_
    Dim sRs As Object
    Dim Loop_Count As Integer
    Dim FunStr1 As String
    Dim PatientID As String
    Dim PatientNo As String
    Dim ii As Integer, sDeCnt    As Integer
    Dim Testcd As String
    Dim OutputData As String, sOrderLst As String
    Dim EndStr, strEqpCd, sChannel, strOrdLst As String
    Dim intRow  As Integer
    Dim itemX As ListItem
    Dim sHead As String, sPInfo As String, sOrder As String, sLast As String
    Dim sTmp  As String
    Dim sSampleNo  As String
    Dim strOrder As String
    Dim ChkS    As String
    
    On Error GoTo errSend
    
    strOrder = ""
    
    Select Case SendCount
        Case 0      ' Header Record
'            tmp = objInt.FrameNo & "H|\^&|||AxSYM^3.60^1180^H1P1O1R1C1Q1L1M1|||||||P|1|" & Format(DBConn.getSysDate, "yyyyMMddhhMMss") & Chr(13) & Chr(3)
            sTmp = "1H|\^&||||||||||P|1|" & vbCr & Chr(3)
            SendCount = 1
        Case 1      ' Patient Record
'            tmp = objInt.FrameNo & "P|1||" & Trim$(objAxsym.SpcYY & objAxsym.SpcNo) & "||" & vbCr & Chr(3)
            sTmp = "1P|1||" & Trim$(AXSYM.SampleNo) & "||" & Chr(13) & Chr(3)
            SendCount = 2
        Case 2      ' Order Record
            sSampleNo = AXSYM.SampleNo
        
            With spdResult1
                Set sRs = f_subSet_TestList(Trim(sSampleNo))
                If Not sRs.EOF Then
                    .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
                    intRow = SeqNullSearch(spdResult1, sRs("검체번호"), 1)
                    
                    .SetText 1, intRow, Trim(sRs("검체번호")) & ""
                    .SetText 4, intRow, Trim(sRs("성명")) & ""
                    .SetText 5, intRow, IIf(Trim(sRs("성별코드")) = "M", "남", "여") & "/" & Format(Now, "yyyy") - Format(Trim(sRs("생년월일")), "yyyy") - 1
                    If sRs("입외구분") = "1" Then
                        .SetText 6, intRow, "입원"
                    ElseIf sRs("입외구분") = "2" Then
                        .SetText 6, intRow, "외래"
                    Else
                        .SetText 6, intRow, "퇴원"
                    End If
                    .SetText 7, intRow, Trim(sRs("병록번호")) & ""
                    .SetText 8, intRow, Trim(sRs("접수일자")) & ""
                    '-- 검사항목조회
                    Do Until sRs.EOF
                        strEqpCd = Trim(sRs("품목코드"))
                        Set itemX = lvwCuData.FindItem(strEqpCd, lvwSubItem, , lvwWhole)
        
                        If Not itemX Is Nothing Then
                            spdResult1.Row = intRow
                            spdResult1.Col = itemX.Index + 8
                            spdResult1.BackColor = &HC6FEFF '&H80C0FF
                            DoEvents
                        End If
                        
                        If strOrder = "" Then
                            strOrder = "^^^" & Trim(itemX.Tag)
                        Else
                            strOrder = strOrder & "\^^^" & Trim(itemX.Tag)
                        End If
                        sRs.MoveNext
                    Loop
                End If
            End With
    
            'sOrder = "1O|1|" & mvarSpcYY & mvarSpcNo & "||" & strOrder & "|R||||||N||||||||||||||Q" & vbCr & Chr(3)
            sTmp = "1O|1|" & AXSYM.SampleNo & "||" & strOrder & "|R||||||N||||||||||||||Q" & Chr(13) & Chr(3)
            SendCount = 3
        Case 3      'Terminator Record
            sTmp = "1L|1" & vbCr & Chr(3)
            SendCount = 4
        Case 4      ' EOT
            comEQP.Output = Chr(4)   'EOT
            SendCount = 0
            Exit Sub
    End Select
    
    
    lblStatus.Caption = "Order 전송 중.."
    ChkS = getChkSum(sTmp)
    'comEQP.Output = Chr(2) & tmp & ChkS & Chr(13) & Chr(10)
    
    SendData = Chr(2) & sTmp & ChkS & Chr(13) & Chr(10)
    
    'SendCount = Int((Len(SendData) / 230)) + 1
    
    Call COM_OUTPUT(SendData)
    Debug.Print SendData
    
    'SendCount = SendCount - 1
    If SendCount = 0 Then
       SendFlg = False
    Else
       SendFlg = True
    End If
    Exit Sub
    
SendTheSampleSub_:
    Call COM_OUTPUT(EOT)
    Exit Sub
    
errSend:

End Sub

Public Function getChkSum(sMsg As String) As String
    Dim i%
    Dim iChkSum As Integer
    
    iChkSum = 0
    For i = 1 To Len(sMsg)
        iChkSum = (iChkSum + Asc(Mid(sMsg, i, 1)))
    Next
    iChkSum = iChkSum Mod 256
    getChkSum = Right("0" & Hex(iChkSum), 2)
    
End Function

Private Function SeqSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long

Dim sCnt As Long

    SeqSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Val(Trim(.Text)) = brSeq Then
                SeqSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With
    
End Function

Private Sub psDataDefine(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object) ', ByVal brOst As String) ' ByRef brItemdeci() As String)

    Dim ii As Long
    Dim jj As Long
    Dim KK As Long
    Dim Found As Boolean
    Dim SID As String
    Dim SEX As String
    Dim AGE As Long
    Dim sql As String
    Dim OutputData As String, sOrderLst As String
    Dim sTemp       As String       ' On Com으로부터 넘겨받은 Receive Data
    Dim Channel_No  As String       ' 문자형 변수
    Dim Patiant_No  As String       ' 환자번호
    Dim pGrid_Point As Integer      ' 해당 검사자 Point
    Dim Max_Arary_Cnt As Integer    ' 검사 항목수
    '-------------------------------' 임시 변수들.....
    Dim sDeCnt      As Integer
    Dim pDoCount    As Integer
    Dim Loop_Count  As Integer
    Dim sRtn As Integer, sChannel As String, sRstText As String, sRstValue As Single, sUnit As String
    Dim sPatiant_No As Long
    
    Dim sSeq As String
    Dim sCol As Integer
    Dim iCnt As Integer
    Dim intRow As Integer, intCol As Integer, intIdx As Integer
    Dim varTmp
    Dim strRstval As String, strRefVal As String
    Dim itemX  As ListItem
    Dim sqlDoc As String
    Dim sqlRet As Integer
    Dim strTestcd As String
    Dim varBarno, varSPnm, varSPid, varZation, varRegNo, varRegDt
    Dim varSex As String, varAge As String, varRef As String
    Dim blnFlag As Boolean
    Dim strBarno As String
    Dim sRs As Object
    Dim CntsTxt As Integer
    Dim sTxt As String, tTxt As String

    On Error GoTo ErrorTrap
    
    CallForm = "frmInterface - Private Sub psDataDefine()"
    
    ReceiveData = strdata
    
    ' 결과버퍼 통합
    cInterface.addBufs
    
    Dim sIDCode$, rcvbufs$, sendbufs$, tmp$
    
    ReceiveData = cInterface.getrcvbufs
    
    'sIDCode = Mid(ReceiveData, 2, 1)
    
    Debug.Print "받은데이타:" & ReceiveData
    Call COM_INPUT(ReceiveData)
    On Error GoTo errReceive
    For iCnt = 1 To Len(ReceiveData)
        Select Case Mid(ReceiveData, iCnt, 1)
            
            Case "H" 'Message Header
                GoSub Clear_AXSYM_
                Exit For
            Case "P" 'Patient Informatioin
                GoSub Clear_AXSYM_
                Exit For
            Case "O" 'Test Order
'                For ii = 1 To Len(ReceiveData)
'
'                    Select Case Mid(ReceiveData, ii, 1)
'                        Case Field_ '|
'                            cntField_ = cntField_ + 1
'                            cntRepeat_ = 0
'                            cntComponent_ = 0
'                        Case Repeat_ '\
'                            cntRepeat_ = cntRepeat_ + 1
'                            cntComponent_ = 0
'                        Case Component_ '^
'                            cntComponent_ = cntComponent_ + 1
'                        Case Slash_ '/
'                            '
'                        Case Else
'                            Select Case cntField_ '|
'                                Case 2
'                                    Select Case cntComponent_ '^
'                                        Case 0
'                                            AXSYM.SampleNo = AXSYM.SampleNo & Mid(ReceiveData, ii, 1)
'                                        Case 1
'                                            AXSYM.SID = AXSYM.SID & Mid(ReceiveData, ii, 1)
'                                        Case 2
'                                            AXSYM.SampleTy = AXSYM.SampleTy & Mid(ReceiveData, ii, 1)
'                                        Case 3
'                                            AXSYM.RackNo = AXSYM.RackNo & Mid(ReceiveData, ii, 1)
'                                        Case 4
'                                            AXSYM.Position = Mid(ReceiveData, ii, 1)
'                                        Case 6
'                                            AXSYM.Priority = Mid(ReceiveData, ii, 1)
'                                    End Select
'                                Case 4
'                                    Select Case cntComponent_ Mod 3 '^
'                                        Case 0
'                                            If IsNumeric(Mid(ReceiveData, ii, 1)) Then
'                                                AXSYM.TestId(cntRepeat_) = AXSYM.TestId(cntRepeat_) & Mid(ReceiveData, ii, 1)
'                                            End If
'                                    End Select
'                                Case 22
'                                    AXSYM.TestDate = AXSYM.TestDate & Mid(ReceiveData, ii, 1)
'
'                            End Select
'                    End Select
'                Next ii
'                AXSYM.TestTime = Mid(AXSYM.TestDate, 9, 6)
'                AXSYM.TestDate = left(AXSYM.TestDate, 8)
                
                sSeq = Trim(AXSYM.SID)
                sCol = 2
                
                'Patiant_Recevid = SeqSearch(spdResult1, sSeq, sCol)
                Erase fAxsym
            
                Do While InStr(ReceiveData, Chr$(124)) > 0 '--Chr(29)
                    pDoCount = pDoCount + 1
                    fAxsym(pDoCount) = Text_Redefine(ReceiveData, Chr$(124))
                    ReceiveData = Mid$(ReceiveData, InStr(ReceiveData, Chr$(124)) + 1)
                    If pDoCount > 99 Then
                        ReceiveData = ""
                        Exit Do
                    End If
                    If pDoCount < 10 Then fAxsym(pDoCount) = Text_Change(fAxsym(pDoCount), Chr$(29), "")
'                    Debug.Print fAxsym(pDoCount)
                Loop
                
                SeqNo = Mid(fAxsym(4), 1, InStr(fAxsym(4), "^") - 1)
                Exit For
    
            Case "R" 'Test Result
                pDoCount = 0
            
                Erase fAxsym
            
                Do While InStr(ReceiveData, Chr$(124)) > 0 '--Chr(29)
                    pDoCount = pDoCount + 1
                    fAxsym(pDoCount) = Text_Redefine(ReceiveData, Chr$(124))
                    ReceiveData = Mid$(ReceiveData, InStr(ReceiveData, Chr$(124)) + 1)
                    If pDoCount > 99 Then
                        ReceiveData = ""
                        Exit Do
                    End If
                    If pDoCount < 10 Then fAxsym(pDoCount) = Text_Change(fAxsym(pDoCount), Chr$(29), "")
                Loop
                
                Channel_No = Mid(fAxsym(3), 4, 3)
                '-- 숫자형 결과
                If Len(SeqNo) > 0 And Right(fAxsym(3), 1) = "F" Then 'And (Channel_No <> "106" And Channel_No <> "118" And Channel_No <> "841" And Channel_No <> "817") Then
                    intRow = 0
                    With spdResult1
                        intRow = SeqSearch(brspread, SeqNo, 1)
                        '-- 해당번호 찾음
                        If intRow > 0 Then
                            For intCol = 9 To .MaxCols
                                .GetText intCol, 0, varTmp
                                Channel_No = Mid(fAxsym(3), 4, 3)
                                If Right(Channel_No, 1) = "^" Then Channel_No = Mid(Channel_No, 1, 2)
                                
                                If Right(fAxsym(3), 1) = "P" Then Exit Sub
                                
                                strRstval = fAxsym(4)
                                Select Case Channel_No
                                    Case "106" '-- Hbs Ag
                                        If Val(strRstval) >= 2 And Val(strRstval) <= 10 Then '-- WeeklyPositive
                                            strRstval = "WeekPositive(" & strRstval & ")"
                                        ElseIf Val(strRstval) > 10 Then '-- Positive
                                            strRstval = "Positive(" & strRstval & ")"
                                        ElseIf Val(strRstval) < 2 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case "118" '-- Hbs Ab
                                        If Val(strRstval) >= 10 And Val(strRstval) <= 20 Then '-- WeeklyPositive
                                            strRstval = "WeekPositive(" & strRstval & ")"
                                        ElseIf Val(strRstval) > 20 Then '-- Positive
                                            strRstval = "Positive(" & strRstval & ")"
                                        ElseIf Val(strRstval) < 10 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case "841" '-- HCV
                                        If Val(strRstval) > 1 Then '-- Positive
                                            strRstval = "" '"Positive(" & strRstval & ")"
                                        ElseIf Val(strRstval) < 1 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case "817" '-- Hiv
                                        If Val(strRstval) > 1 Then '-- Positive
                                            strRstval = ""
                                        ElseIf Val(strRstval) < 1 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case Else
                                
                                End Select
                                
                                strRstval = Mid(strRstval, 1, 20)
                                                                
                                Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                If Not itemX Is Nothing Then
                                    If Channel_No = itemX.Tag Then
                                        .SetText intCol, intRow, strRstval
                                    
                                        strTestcd = itemX.ListSubItems(1)
                                        
                                        .GetText 1, intRow, varBarno
                                        .GetText 4, intRow, varSPnm
                                        .GetText 5, intRow, varSPid '--성별/나이
                                        .GetText 6, intRow, varZation
                                        If varZation = "입원" Then
                                            varZation = "1"
                                        ElseIf varZation = "외래" Then
                                            varZation = "2"
                                        Else
                                            varZation = "3"
                                        End If
                                        .GetText 7, intRow, varRegNo
                                        .GetText 8, intRow, varRegDt
                                        varSex = Mid(varSPid, 1, 1)
                                        varAge = Mid(varSPid, 3)
                                        varRef = ""
                                        
                                        If Channel_No <> "106" And Channel_No <> "118" And Channel_No <> "841" And Channel_No <> "817" Then
                                            If varSex = "남" Then
                                                If fAxsym(4) < Val(itemX.ListSubItems(8)) Then
                                                    varRef = "L"
                                                ElseIf fAxsym(4) > Val(itemX.ListSubItems(9)) Then
                                                    varRef = "H"
                                                End If
                                            Else
                                                If fAxsym(4) < Val(itemX.ListSubItems(10)) Then
                                                    varRef = "L"
                                                ElseIf fAxsym(4) > Val(itemX.ListSubItems(11)) Then
                                                    varRef = "H"
                                                End If
                                            End If
                                        End If
                                        
                                        spdResult1.Col = intCol
                                        spdResult1.Row = intRow
                                        spdResult1.ForeColor = IIf(varRef <> "", vbRed, vbBlack)

                                        sqlDoc = "Update INTERFACE003" & _
                                                 "   set RESULT1  = '" & strRstval & "', REFERENCE = '" & varRef & "'" & _
                                                 " where SPCNO   = '" & varBarno & "'" & _
                                                 "   and TESTCD  = '" & itemX.Text & "'" & _
                                                 "   and REGDATE = '" & varRegDt & "'"
            
                                        AdoCn_Jet.Execute sqlDoc, sqlRet
                    
                                        If sqlRet = 0 Then
                                            sqlDoc = "insert into INTERFACE003(" & _
                                                     "            SPCNO, LACKNO, POSNO, TESTCD, REGDATE, REGNO, HOSPNO, NAME, SEX, AGE, HOSPZATION, RESULT1, RESULT2, RSTGBN, REFERENCE, DELTA, PANIC, TRANSDT, TRANSTM, SERVERGBN)" & _
                                                     "    values( '" & varBarno & "','','', '" & itemX.Text & "', '" & varRegDt & "','','" & varRegNo & "', " & _
                                                     "            '" & varSPnm & "', '" & varSex & "','" & varAge & "','" & varZation & "'," & _
                                                     "            '" & strRstval & "','','', '" & varRef & "','',''," & _
                                                     "            '" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','')"
                                            AdoCn_Jet.Execute sqlDoc
                                        End If
                                    End If
                                End If
                                                        
                                Set itemX = Nothing
                            Next
                        '-- 해당번호 못찾음
                        Else
                            intRow = SeqNullSearch(spdResult1, "", 1)
                            If intRow = 0 Then
                                .MaxRows = .MaxRows + 1
                                intRow = .MaxRows
                                lstErr.AddItem "검체번호 " & SeqNo & "은(는) 오더리스트에 없는 결과입니다."
                                .SetText 1, .MaxRows, SeqNo
                                .SetText 8, .MaxRows, Format(Now, "yyyy-mm-dd")
                            Else
                                .SetText 1, intRow, SeqNo
                                .SetText 8, intRow, Format(Now, "yyyy-mm-dd")
                                lstErr.AddItem "검체번호 " & SeqNo & "은(는) 오더리스트에 없는 결과입니다."
                            End If
                            For intCol = 9 To .MaxCols
                                .GetText intCol, 0, varTmp
                                Channel_No = Mid(fAxsym(3), 4, 3)
                                If Right(Channel_No, 1) = "^" Then Channel_No = Mid(Channel_No, 1, 2)

                                If Right(fAxsym(3), 1) = "P" Then Exit Sub

                                strRstval = fAxsym(4)
                                Select Case Channel_No
                                    Case "106" '-- Hbs Ag
                                        If Val(strRstval) >= 2 And Val(strRstval) <= 10 Then '-- WeeklyPositive
                                            strRstval = "WeekPositive(" & strRstval & ")"
                                        ElseIf Val(strRstval) > 10 Then '-- Positive
                                            strRstval = "Positive(" & strRstval & ")"
                                        ElseIf Val(strRstval) < 2 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case "118" '-- Hbs Ab
                                        If Val(strRstval) >= 10 And Val(strRstval) <= 20 Then '-- WeeklyPositive
                                            strRstval = "WeekPositive(" & strRstval & ")"
                                        ElseIf Val(strRstval) > 20 Then '-- Positive
                                            strRstval = "Positive(" & strRstval & ")"
                                        ElseIf Val(strRstval) < 10 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case "841" '-- HCV
                                        If Val(strRstval) > 1 Then '-- Positive
                                            strRstval = "Positive(" & strRstval & ")"
                                        ElseIf Val(strRstval) < 1 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case "817" '-- Hiv
                                        If Val(strRstval) > 1 Then '-- Positive
                                            strRstval = ""
                                        ElseIf Val(strRstval) < 1 Then  '-- Negative
                                            strRstval = "Negative"
                                        End If
                                    Case Else
                                
                                End Select
                                
                                strRstval = Mid(strRstval, 1, 20)
                                
                                Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                If Not itemX Is Nothing Then
                                    If Channel_No = itemX.Tag Then
                                        .SetText intCol, intRow, strRstval
                                    
                                        strTestcd = itemX.ListSubItems(1)
                                        
                                        .GetText 1, intRow, varBarno
                                        .GetText 4, intRow, varSPnm
                                        .GetText 5, intRow, varSPid '--성별/나이
                                        .GetText 6, intRow, varZation
                                        If varZation = "입원" Then
                                            varZation = "1"
                                        ElseIf varZation = "외래" Then
                                            varZation = "2"
                                        Else
                                            varZation = "3"
                                        End If
                                        .GetText 7, intRow, varRegNo
                                        .GetText 8, intRow, varRegDt
                                        varSex = Mid(varSPid, 1, 1)
                                        varAge = Mid(varSPid, 3)
                                        varRef = ""
                                        
                                        If Channel_No <> "106" And Channel_No <> "118" And Channel_No <> "841" And Channel_No <> "817" Then
                                            If varSex = "남" Then
                                                If fAxsym(4) < Val(itemX.ListSubItems(8)) Then
                                                    varRef = "L"
                                                ElseIf fAxsym(4) > Val(itemX.ListSubItems(9)) Then
                                                    varRef = "H"
                                                End If
                                            Else
                                                If fAxsym(4) < Val(itemX.ListSubItems(10)) Then
                                                    varRef = "L"
                                                ElseIf fAxsym(4) > Val(itemX.ListSubItems(11)) Then
                                                    varRef = "H"
                                                End If
                                            End If
                                        End If
                                        
                                        spdResult1.Col = intCol
                                        spdResult1.Row = intRow
                                        spdResult1.ForeColor = IIf(varRef <> "", vbRed, vbBlack)

                                        sqlDoc = "Update INTERFACE003" & _
                                                 "   set RESULT1  = '" & strRstval & "', REFERENCE = '" & varRef & "'" & _
                                                 " where SPCNO   = '" & varBarno & "'" & _
                                                 "   and TESTCD  = '" & itemX.Text & "'" & _
                                                 "   and REGDATE = '" & varRegDt & "'"
            
                                        AdoCn_Jet.Execute sqlDoc, sqlRet
                    
                                        If sqlRet = 0 Then
                                            sqlDoc = "insert into INTERFACE003(" & _
                                                     "            SPCNO, LACKNO, POSNO, TESTCD, REGDATE, REGNO, HOSPNO, NAME, SEX, AGE, HOSPZATION, RESULT1, RESULT2, RSTGBN, REFERENCE, DELTA, PANIC, TRANSDT, TRANSTM, SERVERGBN)" & _
                                                     "    values( '" & varBarno & "','','', '" & itemX.Text & "', '" & varRegDt & "','','" & varRegNo & "', " & _
                                                     "            '" & varSPnm & "', '" & varSex & "','" & varAge & "','" & varZation & "'," & _
                                                     "            '" & strRstval & "','','', '" & varRef & "','',''," & _
                                                     "            '" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','')"
                                            AdoCn_Jet.Execute sqlDoc
                                        End If
                                    End If
                                End If
                                                        
                                Set itemX = Nothing
                            Next
                            
                        End If
                    End With
'                '-- 판정형 결과
'                ElseIf Len(SeqNo) > 0 And Right(fAxsym(3), 1) = "I" And (Channel_No = "106" Or Channel_No = "118" Or Channel_No = "841" Or Channel_No = "817") Then
'                    intRow = 0
'                    With spdResult1
'                        intRow = SeqSearch(brspread, SeqNo, 1)
'                        If intRow > 0 Then
'                            For intCol = 9 To .MaxCols
'                                .GetText intCol, 0, varTmp
'                                Channel_No = Mid(fAxsym(3), 4, 3)
'
'                                If Right(Channel_No, 1) = "^" Then Channel_No = Mid(Channel_No, 1, 2)
'                                strRstval = fAxsym(4)
'                                Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
'                                If Not itemX Is Nothing Then
'                                    For intIdx = 1 To .MaxCols
'                                        If Len(Channel_No) > 0 Then
'                                            If Channel_No = itemX.Tag Then
'                                                .SetText intCol, intRow, strRstval
'
'                                                strTestcd = itemX.ListSubItems(1)
'
'                                                .GetText 1, intRow, varBarno
'                                                .GetText 4, intRow, varSPnm
'                                                .GetText 5, intRow, varSPid '--성별/나이
'                                                .GetText 6, intRow, varZation
'                                                If varZation = "입원" Then
'                                                    varZation = "1"
'                                                ElseIf varZation = "외래" Then
'                                                    varZation = "2"
'                                                Else
'                                                    varZation = "3"
'                                                End If
'                                                .GetText 7, intRow, varRegNo
'                                                .GetText 8, intRow, varRegDt
'                                                varSex = Mid(varSPid, 1, 1)
'                                                varAge = Mid(varSPid, 3)
'                                                varRef = ""
''                                                If varSex = "남" Then
''                                                    If fAxsym(4) < Val(itemX.ListSubItems(8)) Then
''                                                        varRef = "L"
''                                                    ElseIf fAxsym(4) > Val(itemX.ListSubItems(9)) Then
''                                                        varRef = "H"
''                                                    End If
''                                                Else
''                                                    If fAxsym(4) < Val(itemX.ListSubItems(10)) Then
''                                                        varRef = "L"
''                                                    ElseIf fAxsym(4) > Val(itemX.ListSubItems(11)) Then
''                                                        varRef = "H"
''                                                    End If
''                                                End If
'                                                spdResult1.Col = intCol
'                                                spdResult1.Row = intRow
'                                                spdResult1.ForeColor = IIf(varRef <> "", vbRed, vbBlack)
'
'                                                sqlDoc = "Update INTERFACE003" & _
'                                                         "   set RESULT1  = '" & fAxsym(4) & "', REFERENCE = '" & varRef & "'" & _
'                                                         " where SPCNO   = '" & varBarno & "'" & _
'                                                         "   and TESTCD  = '" & itemX.Text & "'" & _
'                                                         "   and REGDATE = '" & varRegDt & "'"
'
'                                                AdoCn_Jet.Execute sqlDoc, sqlRet
'
'                                                If sqlRet = 0 Then
'                                                    sqlDoc = "insert into INTERFACE003(" & _
'                                                             "            SPCNO, LACKNO, POSNO, TESTCD, REGDATE, REGNO, HOSPNO, NAME, SEX, AGE, HOSPZATION, RESULT1, RESULT2, RSTGBN, REFERENCE, DELTA, PANIC, TRANSDT, TRANSTM, SERVERGBN)" & _
'                                                             "    values( '" & varBarno & "','','', '" & itemX.Text & "', '" & varRegDt & "','','" & varRegNo & "', " & _
'                                                             "            '" & varSPnm & "', '" & varSex & "','" & varAge & "','" & varZation & "'," & _
'                                                             "            '" & fAxsym(4) & "','','', '" & varRef & "','',''," & _
'                                                             "            '" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','')"
'                                                    AdoCn_Jet.Execute sqlDoc
'                                                End If
'                                            End If
'                                        End If
'                                        Exit For
'                                    Next intIdx
'                                End If
'                                Set itemX = Nothing
'                            Next
'                        '-- 해당번호 못찾음
'                        Else
'                            intRow = SeqNullSearch(spdResult1, "", 1)
'                            If intRow = 0 Then
'                                .MaxRows = .MaxRows + 1
'                                intRow = .MaxRows
'                                lstErr.AddItem "검체번호 " & intRow & "은(는) 오더리스트에 없는 결과입니다."
'                                .SetText 1, .MaxRows, SeqNo
'                                .SetText 8, .MaxRows, Format(Now, "yyyy-mm-dd")
'                            Else
'                                .SetText 1, intRow, SeqNo
'                                .SetText 8, intRow, Format(Now, "yyyy-mm-dd")
'                                lstErr.AddItem "검체번호 " & intRow & "은(는) 오더리스트에 없는 결과입니다."
'                            End If
'
'                            For intCol = 9 To .MaxCols
'                                .GetText intCol, 0, varTmp
'                                Channel_No = Mid(fAxsym(3), 4, 3)
'                                If Right(Channel_No, 1) = "^" Then Channel_No = Mid(Channel_No, 1, 2)
'
'                                If Right(fAxsym(3), 1) = "P" Then Exit Sub
'
'                                strRstval = fAxsym(4)
'                                Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
'                                If Not itemX Is Nothing Then
'                                    If Channel_No = itemX.Tag Then
'                                        .SetText intCol, intRow, strRstval
'
'                                        strTestcd = itemX.ListSubItems(1)
'
'                                        .GetText 1, intRow, varBarno
'                                        .GetText 4, intRow, varSPnm
'                                        .GetText 5, intRow, varSPid '--성별/나이
'                                        .GetText 6, intRow, varZation
'                                        If varZation = "입원" Then
'                                            varZation = "1"
'                                        ElseIf varZation = "외래" Then
'                                            varZation = "2"
'                                        Else
'                                            varZation = "3"
'                                        End If
'                                        .GetText 7, intRow, varRegNo
'                                        .GetText 8, intRow, varRegDt
'                                        varSex = Mid(varSPid, 1, 1)
'                                        varAge = Mid(varSPid, 3)
'                                        varRef = ""
''                                        If varSex = "남" Then
''                                            If fAxsym(4) < Val(itemX.ListSubItems(8)) Then
''                                                varRef = "L"
''                                            ElseIf fAxsym(4) > Val(itemX.ListSubItems(9)) Then
''                                                varRef = "H"
''                                            End If
''                                        Else
''                                            If fAxsym(4) < Val(itemX.ListSubItems(10)) Then
''                                                varRef = "L"
''                                            ElseIf fAxsym(4) > Val(itemX.ListSubItems(11)) Then
''                                                varRef = "H"
''                                            End If
''                                        End If
'
'                                        spdResult1.Col = intCol
'                                        spdResult1.Row = intRow
'                                        spdResult1.ForeColor = IIf(varRef <> "", vbRed, vbBlack)
'
'                                        sqlDoc = "Update INTERFACE003" & _
'                                                 "   set RESULT1  = '" & fAxsym(4) & "', REFERENCE = '" & varRef & "'" & _
'                                                 " where SPCNO   = '" & varBarno & "'" & _
'                                                 "   and TESTCD  = '" & itemX.Text & "'" & _
'                                                 "   and REGDATE = '" & varRegDt & "'"
'
'                                        AdoCn_Jet.Execute sqlDoc, sqlRet
'
'                                        If sqlRet = 0 Then
'                                            sqlDoc = "insert into INTERFACE003(" & _
'                                                     "            SPCNO, LACKNO, POSNO, TESTCD, REGDATE, REGNO, HOSPNO, NAME, SEX, AGE, HOSPZATION, RESULT1, RESULT2, RSTGBN, REFERENCE, DELTA, PANIC, TRANSDT, TRANSTM, SERVERGBN)" & _
'                                                     "    values( '" & varBarno & "','','', '" & itemX.Text & "', '" & varRegDt & "','','" & varRegNo & "', " & _
'                                                     "            '" & varSPnm & "', '" & varSex & "','" & varAge & "','" & varZation & "'," & _
'                                                     "            '" & fAxsym(4) & "','','', '" & varRef & "','',''," & _
'                                                     "            '" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','')"
'                                            AdoCn_Jet.Execute sqlDoc
'                                        End If
'                                    End If
'                                End If
'
'                                Set itemX = Nothing
'                            Next
'                        End If
'                    End With
                End If

                
            Case "C" 'Comment
                Exit For
                
            Case "L" 'Message Termination
                If SendFlg = True Then Exit Sub
                AXSYM.SID = Right(Trim(AXSYM.SID), 10)
                If AXSYM.SID = "" Then Exit Sub
                If AXSYM.SampleNo = "" Then AXSYM.SampleNo = "0"
                If AXSYM.SampleTy = "" Then AXSYM.SampleTy = "1"
                Exit For
               
            Case "Q" 'Request Information
                SendFlg = True
                SendCount = 0
                cInterface.State = "Q"
                If Len(ReceiveData) > 0 Then
                    AXSYM.SampleNo = Mid(ReceiveData, InStr(ReceiveData, "^") + 1)
                    AXSYM.SampleNo = Mid(AXSYM.SampleNo, 1, InStr(AXSYM.SampleNo, "|") - 1)
                End If
'                For ii = 1 To Len(ReceiveData)
'                    Select Case Mid(ReceiveData, ii, 1)
'                        Case Field_ '|
'                            cntField_ = cntField_ + 1
'                            cntRepeat_ = 0
'                            cntComponent_ = 0
'                            cntSlash_ = 0
'                        Case Repeat_ '\
'                            cntRepeat_ = cntRepeat_ + 1
'                            cntComponent_ = 0
'                            cntSlash_ = 0
'                        Case Component_ '^
'                            cntComponent_ = cntComponent_ + 1
'                            cntSlash_ = 0
'                        Case Slash_ '/
'                            cntSlash_ = cntSlash_ + 1
'                        Case Else
'                            Select Case cntComponent_ '^
'                                Case 2 'Sid
'                                    Select Case cntSlash_ '/
'                                        Case 0
'                                            AXSYM.SampleNo = Trim(AXSYM.SampleNo) & Mid(ReceiveData, ii, 1)
'                                        Case 1
'                                            AXSYM.SID = Trim(AXSYM.SID) & Mid(ReceiveData, ii, 1)
'                                        Case 2
'                                            AXSYM.SampleTy = Trim(AXSYM.SampleTy) & Mid(ReceiveData, ii, 1)
'                                        Case 3
'                                            AXSYM.RackNo = Trim(AXSYM.RackNo) & Mid(ReceiveData, ii, 1)
'                                        Case 4
'                                            AXSYM.Position = Trim(AXSYM.Position) & Mid(ReceiveData, ii, 1)
'                                        Case 6
'                                            AXSYM.Priority = Mid(ReceiveData, ii, 1)
'                                    End Select
'                            End Select
'                    End Select
'                Next ii
                
                If AXSYM.SampleNo = "" Then AXSYM.SampleNo = "0"
                If AXSYM.SampleTy = "" Then AXSYM.SampleTy = "1"
                Exit For
            Case Else
                
        End Select
    Next iCnt
    Exit Sub
    
Clear_AXSYM_:
    AXSYM.TestDate = ""
    AXSYM.TestTime = ""
    AXSYM.SampleNo = ""
    AXSYM.SID = ""
    AXSYM.SampleTy = ""
    AXSYM.RackNo = ""
    AXSYM.Position = ""
    AXSYM.Priority = ""
    For ii = 0 To 100
        AXSYM.TestId(ii) = ""
        AXSYM.Result(ii) = ""
        AXSYM.Status(ii) = ""
        AXSYM.Rerun(ii) = ""
    Next ii
    flgETB = False
    SendCount = 0
    SendFlg = False
    Return
    
    Exit Sub
ErrorTrap:
    Call ErrMsgProc(CallForm)

errReceive:


End Sub


Public Function Text_Redefine(FSend_Str As String, FCheck_Char As String) As String
    
    If InStr(FSend_Str, FCheck_Char) > 0 Then
        Text_Redefine = left$(FSend_Str, InStr(FSend_Str, FCheck_Char) - 1)
    Else
        Text_Redefine = FSend_Str
    End If
    
End Function

Public Function Text_Change(FSend_Str As String, FCheck_Char As String, FChange_Char As String) As String
Dim Pos_point As Integer
    Do
        Pos_point = InStr(FSend_Str, FCheck_Char)
        If Pos_point < 1 Then
            Exit Do
        ElseIf Pos_point = 1 Then
            FSend_Str = FChange_Char + Mid$(FSend_Str, 2)
        Else
            FSend_Str = Mid$(FSend_Str, 1, Pos_point - 1) + FChange_Char + Mid$(FSend_Str, Pos_point + 1)
        End If
    Loop
    Text_Change = FSend_Str
    
End Function

Private Function SeqNullSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqNullSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = "" Then
                SeqNullSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function



Private Sub cmdStartNo_Click()
Dim sNo As String, sCnt As Integer, sAdd As Integer

AgainInput:
    sNo = InputBox("시작 번호를 입력하세요 !")
    If Len(sNo) > 0 And spdResult1.MaxRows > 0 Then
        If Not IsNumeric(sNo) Then
            MsgBox "숫자만 입력하세요.!", vbCritical
            GoTo AgainInput
        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .MaxRows
                .Row = sCnt
                .Col = 0:       .Text = Trim(sAdd + Val(sNo))
                sAdd = sAdd + 1
            Next sCnt
        
            .StartingRowNumber = Val(sNo)
        End With
    End If

End Sub

Private Sub Form_Activate()

    If IS_SET = False Then Unload Me
    
    Call cmdRun           ' 실행

End Sub

Private Sub Form_Load()
    
'    Me.Show
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    Call cmdClear               ' 초기화
    Call f_subSet_ItemHeader    ' 리스트해더
    Call f_subSet_ItemList      ' 검사항목
    
    Call f_subSet_ComCharacter  ' 통신문자
    Call f_subGet_Setting       ' 통신설정
    
    'Call cmdRun           ' 실행
    
    dtpDate.Value = Now
    dtpDate1.Value = Now
    
    Open App.Path + "\" + "Axsym.log" For Append As #1

    Print #1, Chr(13) + Chr(10);
    
    cInterface.phase = 1
    bufcnt = 0
    
    DoEvents

End Sub

Private Sub f_subSet_ItemHeader()
    
    '검사코드 테이블
    With lvwCuData
        .View = lvwReport
        Set .ColumnHeaderIcons = imlList
        Set .SmallIcons = imlList
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HideColumnHeaders = True
        With .ColumnHeaders
            .Clear
            Call .Add(, TEST_NM_EQP, "ID", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_CD_LIS, "검사코드", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_NM_LIS, "검 사 명", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_VALUES, "검사결과", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTA", "DELTA", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTAGBN", "DELTAGBN", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICL", "PANIC(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICH", "PANIC(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFLM", "참고치남(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFHM", "참고치남(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFLF", "참고치여(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFHF", "참고치여(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "AUTOVERIFY", "재검", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, "REMARK", "검체코드", (lvwCuData.Width - 310) * 0.1)
        End With
        .HideColumnHeaders = False
    End With

End Sub

Private Function f_subSet_WorkList()
    Dim sqlRet      As Integer
    Dim gSql        As String

On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"

    Set AdoRs_ORACLE = New ADODB.Recordset

    gSql = "select 처방전코드, 처방전명, 검체번호, 검체명, 품목코드, 품목명, 접수일자, 입외구분, 병록번호, 성명, 생년월일, 나이, 성별코드, 검체코드, 과코드, 특기사항, 처리구분코드 " & _
           "  from cli.검사검체1v " & _
           "   where (처리구분코드 <> 'N' or 처리구분코드 <> 'R') " & _
           "   and 접수일자 = '2004-03-17' " & _
           "   and 검체번호 = '9135' " & _
           " order by 검체번호 "

'           " where 처방전코드 = '250' " & _

'    gSql = "select * from cli.검사기품목 "
    
    With DataRs(gSql)
        Dim ii As Integer
        
        Do Until .EOF
        
            'Debug.Print .Fields("처방전코드")
            If .Fields("처방전코드") = "249" Then
'                Stop
'                Debug.Print .Fields("처방전코드")
'                Debug.Print .Fields("처방전명")
                Debug.Print .Fields("품목코드")
                Debug.Print .Fields("품목명")
            End If
            'Debug.Print .Fields("처방전명")
            
            .MoveNext
        Loop
        .Close
    End With

Exit Function

ErrorTrap:
    'Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)


End Function


Private Function f_subSet_TestList(ByVal strRecei As String)
    Dim sqlRet      As Integer
    Dim gSql        As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_DateOrder() As ADODB.Recordset"
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    
    gSql = "select * " & _
           "  from cli.검사검체1v " & _
           " where 처방전코드 = '250' " & _
           "   and 검체번호 = '" & strRecei & "'" & _
           "   and (처리구분코드 <> 'N' or 처리구분코드 <> 'R') " & _
           " order by 품목코드 "
    
    Set f_subSet_TestList = DataRs(gSql)

Exit Function

ErrorTrap:
'    Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)

    
End Function


Private Sub f_subGet_Setting()
    
    Dim objComSetting As clsCommon
    Dim Baudratio As String
    Dim Paritybit As String
    Dim Databit As String
    Dim Stopbit As String
    
    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subGet_Setting()"
    Set objComSetting = New clsCommon
    
    With objComSetting
        .SetAdoCn AdoCn_Jet
        Set mAdoRs = .Get_EqpProperty(INS_CODE)
    End With
    Set objComSetting = Nothing
    
    If mAdoRs Is Nothing Then
        IS_SET = False
        MsgBox INS_CODE & " 에 대한 장비 통신 구성이 없습니다. 통신 설정후 다시 시도 하십시오.", vbExclamation
        Exit Sub
    Else
        If mAdoRs.EOF Then
            IS_SET = False
            MsgBox INS_CODE & " 에 대한 장비 통신 구성이 없습니다. 통신 설정후 다시 시도 하십시오.", vbExclamation
            Set mAdoRs = Nothing
            Exit Sub
        Else
            IS_SET = True
            Baudratio = Trim(mAdoRs.Fields("COM_SPEED") & "")
            Paritybit = Trim(mAdoRs.Fields("COM_PARITYBIT") & "")
            Databit = Trim(mAdoRs.Fields("COM_DATABIT") & "")
            Stopbit = Trim(mAdoRs.Fields("COM_STOPBIT") & "")
            
            With comEQP
                .CommPort = Trim(mAdoRs.Fields("COM_PORT") & "")
'                .Handshaking = Trim(mAdoRs.Fields("COM_HANDSHAK") & "")
'                .InputMode = Trim(mAdoRs.Fields("COM_INPUTMOD") & "")
'                .DTREnable = Trim(mAdoRs.Fields("COM_DTR") & "")
'                .EOFEnable = Trim(mAdoRs.Fields("COM_EOF") & "")
'                .NullDiscard = Trim(mAdoRs.Fields("COM_NULDIS") & "")
'                .RTSEnable = Trim(mAdoRs.Fields("COM_RTS") & "")
'                .InBufferSize = Trim(mAdoRs.Fields("COM_IBS") & "")
'                .InputLen = Trim(mAdoRs.Fields("COM_INLEN") & "")
'                .OutBufferSize = Trim(mAdoRs.Fields("COM_OBS") & "")
'                .ParityReplace = Trim(mAdoRs.Fields("COM_PTR") & "")
'                .RThreshold = Trim(mAdoRs.Fields("COM_RTH") & "")
'                .SThreshold = Trim(mAdoRs.Fields("COM_STH") & "")
                .Settings = Baudratio & "," & Paritybit & "," & Databit & "," & Stopbit
            End With
            Call Del_OldData
        End If
    End If
    
    Set mAdoRs = Nothing
Exit Sub

ErrRoutine:
    Set objComSetting = Nothing
    Set mAdoRs = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    
    Call cmdStop
    Set Result = Nothing
    
    Close #1

End Sub


Private Sub imgReceive_DblClick()

    If FrameResult.Visible = False Then
        FrameResult.Visible = True
    Else
        FrameInterface.Visible = True
        FrameInterface.ZOrder 0
    End If

End Sub

Private Sub imgSend_DblClick()
    
    If FrameResult.Visible = True Then
        FrameResult.Visible = False
    Else
        FrameInterface.Visible = True
        FrameInterface.ZOrder 0
    End If

End Sub

Private Sub Order_Ready(ByVal ACK As String)

    Static msgIndex As Long
    
    Select Case ACK
        Case Chr(COM_ENQ)
            msgIndex = 1
        Case Chr(COM_ACK)
            msgIndex = msgIndex + 1
        Case Chr(COM_NACK)
            msgIndex = msgIndex
        Case Chr(COM_EOT)
            msgIndex = 7
            Set Order = Nothing
        Case Else
        
    End Select
    
    Select Case msgIndex
        Case 1
            Call COM_OUTPUT(Order.MSG_ENQ)
        Case 2
            Call COM_OUTPUT(Order.MSG_HEADER)
        Case 3
            Call COM_OUTPUT(Order.MSG_PATIENT)
        Case 4
            Call COM_OUTPUT(Order.MSG_ORDER)
        Case 5
            Call COM_OUTPUT(Order.MSG_TERMINATION)
        Case 6
            Call COM_OUTPUT(Order.MSG_EOT)
        Case Else
    End Select
    
End Sub

Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol1 As Integer
    Dim intCol2 As Integer
    Dim introw1 As Integer
    Dim intRow2 As Integer
    Dim iCnt    As Integer
    
    If Row = 0 Then Exit Sub
    
    intCol1 = 9
    intCol2 = 2
    introw1 = 1
    
    With spdResult1
        For iCnt = intCol1 To .MaxCols
            .Row = Row
            .Col = intCol1
            
            spdRstDetail.Row = introw1
            spdRstDetail.Col = intCol2
            spdRstDetail.Text = .Text
            
            introw1 = introw1 + 1
            intCol1 = intCol1 + 1
            
            If introw1 > spdRstDetail.MaxRows Then
                introw1 = 1
                intCol2 = intCol2 + 2
            End If
        
        Next
    End With
    
End Sub

'Private Sub spdResult1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Dim intCol1 As Integer
'    Dim intCol2 As Integer
'    Dim intRow1 As Integer
'    Dim intRow2 As Integer
'    Dim iCnt    As Integer
'
'    intCol1 = 9
'    intCol2 = 2
'    intRow1 = 1
'
'    With spdResult1
'        For iCnt = intCol1 To .MaxCols
'            .Row = NewRow
'            .Col = intCol1
'
'            spdRstDetail.Row = intRow1
'            spdRstDetail.Col = intCol2
'            spdRstDetail.Text = .Text
'
'            intRow1 = intRow1 + 1
'            intCol1 = intCol1 + 1
'
'            If intRow1 > spdRstDetail.MaxRows Then
'                intRow1 = 1
'                intCol2 = intCol2 + 2
'            End If
'
'        Next
'    End With
'
'End Sub

Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For i = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
End Sub


Private Sub WorkList_Click()
    SSPanel1.Visible = True
End Sub

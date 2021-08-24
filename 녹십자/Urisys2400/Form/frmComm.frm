VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmComm 
   Caption         =   "Interface"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15420
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10080
   ScaleWidth      =   15420
   WindowState     =   2  '최대화
   Begin HSCotrol.UserPanel uplUrineMicro 
      Height          =   7065
      Left            =   7260
      TabIndex        =   83
      Top             =   1440
      Visible         =   0   'False
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   12462
      BorderStyle     =   1
      Bevel           =   2
      CaptionVisible  =   -1  'True
      Caption         =   "::: Urine Micro Order List"
      Moveble         =   -1  'True
      CloseEnabled    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComDlg.CommonDialog cdgExport 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin BHButton.BHImageButton cmd닫기 
         Height          =   375
         Left            =   2460
         TabIndex        =   86
         Top             =   6570
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   661
         Caption         =   "닫기"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdExcel_Export 
         Height          =   375
         Left            =   60
         TabIndex        =   85
         Top             =   6570
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   661
         Caption         =   "엑셀 내보내기"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin FPUSpreadADO.fpSpread spdMicro 
         Height          =   6225
         Left            =   60
         TabIndex        =   84
         Top             =   300
         Width           =   4755
         _Version        =   524288
         _ExtentX        =   8387
         _ExtentY        =   10980
         _StockProps     =   64
         ColsFrozen      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   0
         ScrollBars      =   2
         SpreadDesigner  =   "frmComm.frx":0000
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1500
      Top             =   2610
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSCommLib.MSComm comEQP 
      Left            =   3000
      Top             =   5190
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer tmrReceive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4905
      Top             =   5130
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5310
      Top             =   5130
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   3735
      Top             =   5130
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
            Picture         =   "frmComm.frx":0571
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":0B0B
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":10A5
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":163F
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":1BD9
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":2173
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlStatus 
      Left            =   5715
      Top             =   5130
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
            Picture         =   "frmComm.frx":270D
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":2CA7
            Key             =   "NOT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":3241
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":37DB
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":406D
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":41C7
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":4321
            Key             =   "NOF"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   45
      TabIndex        =   1
      Top             =   9015
      Width           =   15315
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   0
         Left            =   6615
         TabIndex        =   21
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Run"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   1
         Left            =   7920
         TabIndex        =   22
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Stop"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   2
         Left            =   9225
         TabIndex        =   23
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Clear"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   3
         Left            =   10530
         TabIndex        =   24
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Close"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "frmComm.frx":447B
         ImgOutLineSize  =   3
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
         Left            =   960
         TabIndex        =   3
         Top             =   225
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
         Left            =   210
         TabIndex        =   2
         Top             =   225
         Width           =   615
      End
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15420
      _ExtentX        =   27199
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Caption         =   " Communication"
      SubCaption      =   "검사 장비와 통신하여 결과를 저장 합니다."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel pnlPort 
         Height          =   285
         Left            =   11940
         TabIndex        =   70
         Top             =   120
         Width           =   3285
         _Version        =   65536
         _ExtentX        =   5794
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Begin VB.Image imgPort 
            Height          =   240
            Left            =   660
            Picture         =   "frmComm.frx":5D05
            Top             =   30
            Width           =   240
         End
         Begin VB.Image imgSend 
            Height          =   240
            Left            =   1800
            Picture         =   "frmComm.frx":628F
            Top             =   30
            Width           =   240
         End
         Begin VB.Image imgReceive 
            Height          =   240
            Left            =   3030
            Picture         =   "frmComm.frx":6819
            Top             =   30
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "Port : "
            Height          =   180
            Left            =   30
            TabIndex        =   73
            Top             =   60
            Width           =   510
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "Send : "
            Height          =   180
            Left            =   1125
            TabIndex        =   72
            Top             =   60
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "Receive : "
            Height          =   180
            Left            =   2160
            TabIndex        =   71
            Top             =   60
            Width           =   855
         End
      End
   End
   Begin TabDlg.SSTab tabWork 
      Height          =   8370
      Left            =   30
      TabIndex        =   4
      Top             =   600
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   14764
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " ▒     WorkList      "
      TabPicture(0)   =   "frmComm.frx":6DA3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "spdRstview"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtHelp"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "spdResult1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ssInformation"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "SSPanel4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "pnlCom2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "List1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lvwCuData"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdPrint"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdStartNo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdPosNo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdRackNo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdLongSelect"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdSearch"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtBarCode"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Command1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "SSPanel2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "pnlCom"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "uplBarcode"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Frame3"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "chkJubsu"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "chkSound"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "chkAuto"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "spd저장체크"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "chkOrderCheck"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "  ▒     받은 결과      "
      TabPicture(1)   =   "frmComm.frx":6DBF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdSel(2)"
      Tab(1).Control(1)=   "cmdSel(3)"
      Tab(1).Control(2)=   "spdResult2"
      Tab(1).Control(3)=   "SSPanel3"
      Tab(1).Control(4)=   "cmdAppend(1)"
      Tab(1).Control(5)=   "cmdRstQuery"
      Tab(1).ControlCount=   6
      Begin VB.CheckBox chkOrderCheck 
         Caption         =   "오더만 보기"
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
         Height          =   200
         Left            =   7320
         TabIndex        =   89
         Top             =   60
         Width           =   1350
      End
      Begin FPUSpreadADO.fpSpread spd저장체크 
         Height          =   4605
         Left            =   390
         TabIndex        =   87
         Top             =   2520
         Visible         =   0   'False
         Width           =   6435
         _Version        =   524288
         _ExtentX        =   11351
         _ExtentY        =   8123
         _StockProps     =   64
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   12
         RowsFrozen      =   1
         SpreadDesigner  =   "frmComm.frx":6DDB
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "자동통보"
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
         Height          =   200
         Left            =   8700
         TabIndex        =   79
         Top             =   60
         Value           =   1  '확인
         Width           =   1110
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "소리사용"
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
         Height          =   200
         Left            =   9870
         TabIndex        =   78
         Top             =   60
         Value           =   1  '확인
         Width           =   1155
      End
      Begin VB.CheckBox chkJubsu 
         Caption         =   "자동접수"
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
         Height          =   200
         Left            =   6690
         TabIndex        =   77
         Top             =   60
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Frame Frame3 
         Height          =   405
         Left            =   90
         TabIndex        =   61
         Top             =   900
         Width           =   675
         Begin Threed.SSCommand cmdSel 
            Height          =   390
            Index           =   0
            Left            =   0
            TabIndex        =   63
            Top             =   0
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   688
            _StockProps     =   78
            BevelWidth      =   1
            Picture         =   "frmComm.frx":75BC
         End
         Begin Threed.SSCommand cmdSel 
            Height          =   390
            Index           =   1
            Left            =   330
            TabIndex        =   62
            Top             =   0
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   688
            _StockProps     =   78
            BevelWidth      =   1
            Picture         =   "frmComm.frx":7A2A
         End
      End
      Begin HSCotrol.UserPanel uplBarcode 
         Height          =   1725
         Left            =   4470
         TabIndex        =   55
         Top             =   1350
         Visible         =   0   'False
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   3043
         BorderStyle     =   1
         Bevel           =   2
         CaptionVisible  =   -1  'True
         Caption         =   "::::: 바코드 번호를 입력하세요.   "
         Moveble         =   -1  'True
         CloseEnabled    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin BHButton.BHImageButton cmdBarcode_Close 
            Height          =   375
            Left            =   90
            TabIndex        =   75
            Top             =   1200
            Width           =   4320
            _ExtentX        =   7620
            _ExtentY        =   661
            Caption         =   "닫기"
            CaptionChecked  =   "BHImageButton1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ImgOutLineSize  =   3
         End
         Begin VB.TextBox txt접수번호 
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
            Height          =   315
            Left            =   3420
            TabIndex        =   58
            Top             =   300
            Width           =   975
         End
         Begin VB.TextBox txt순번 
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
            Height          =   315
            Left            =   1560
            TabIndex        =   57
            Top             =   300
            Width           =   675
         End
         Begin VB.TextBox txtBarcode_Keyin 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   18
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   90
            TabIndex        =   56
            Text            =   "123456789012"
            Top             =   660
            Width           =   4305
         End
         Begin VB.Label Label16 
            Caption         =   "::: 접수번호 :"
            Height          =   255
            Left            =   2340
            TabIndex        =   60
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label Label15 
            Caption         =   "::: Display 순번 :"
            Height          =   255
            Left            =   90
            TabIndex        =   59
            Top             =   360
            Width           =   1575
         End
      End
      Begin HSCotrol.UserPanel pnlCom 
         Height          =   4755
         Left            =   90
         TabIndex        =   27
         Top             =   3270
         Visible         =   0   'False
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   8387
         BorderStyle     =   1
         CloseEnabled    =   -1  'True
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
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3720
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   28
            Top             =   330
            Width           =   10905
         End
         Begin VB.Frame Frame1 
            Height          =   645
            Left            =   45
            TabIndex        =   29
            Top             =   4020
            Width           =   10920
            Begin HSCotrol.CButton cmdCOMSave 
               Height          =   360
               Left            =   9795
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "File Save"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdCOMOutput 
               Height          =   360
               Left            =   1155
               TabIndex        =   31
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Send"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMClear 
               Height          =   360
               Left            =   8730
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Clear"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdCOMInput 
               Height          =   360
               Left            =   90
               TabIndex        =   33
               Top             =   180
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Receive"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   285
         Left            =   4290
         TabIndex        =   35
         Top             =   -60
         Visible         =   0   'False
         Width           =   3075
         _Version        =   65536
         _ExtentX        =   5424
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Enabled         =   0   'False
         Begin VB.OptionButton optBar 
            BackColor       =   &H00FFC0C0&
            Caption         =   "병록번호"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1710
            TabIndex        =   37
            Top             =   -30
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optSeq 
            BackColor       =   &H00FFC0C0&
            Caption         =   "검사번호"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   270
            TabIndex        =   36
            Top             =   60
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "TEST"
         Height          =   285
         Left            =   3300
         TabIndex        =   20
         Top             =   30
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.TextBox txtBarCode 
         Height          =   300
         Left            =   4410
         MaxLength       =   12
         TabIndex        =   5
         Top             =   -30
         Visible         =   0   'False
         Width           =   1500
      End
      Begin BHButton.BHImageButton cmdRstQuery 
         Height          =   375
         Left            =   -69240
         TabIndex        =   26
         Top             =   450
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   661
         Caption         =   "조회"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAppend 
         Height          =   375
         Index           =   1
         Left            =   -61140
         TabIndex        =   25
         Top             =   450
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
         Caption         =   "서버등록"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdSearch 
         Height          =   375
         Left            =   1890
         TabIndex        =   34
         Top             =   -30
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
         Caption         =   "조회"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdLongSelect 
         Height          =   420
         Left            =   4770
         TabIndex        =   38
         Top             =   -150
         Visible         =   0   'False
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   741
         Caption         =   "▽"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdRackNo 
         Height          =   420
         Left            =   7980
         TabIndex        =   39
         Top             =   5370
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Rack변경"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdPosNo 
         Height          =   420
         Left            =   9270
         TabIndex        =   40
         Top             =   5370
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Pos변경"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdStartNo 
         Height          =   375
         Left            =   4950
         TabIndex        =   41
         Top             =   -240
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         Caption         =   "시작번호변경"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   465
         Left            =   -74910
         TabIndex        =   42
         Top             =   405
         Width           =   5625
         _Version        =   65536
         _ExtentX        =   9922
         _ExtentY        =   820
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
         BorderWidth     =   0
         BevelInner      =   1
         Begin VB.ComboBox cboRstgbn 
            Height          =   300
            Index           =   1
            ItemData        =   "frmComm.frx":7EAC
            Left            =   3735
            List            =   "frmComm.frx":7EB9
            Style           =   2  '드롭다운 목록
            TabIndex        =   47
            Top             =   90
            Width           =   1770
         End
         Begin MSMask.MaskEdBox mskRstDate 
            Height          =   300
            Left            =   1305
            TabIndex        =   43
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskRstDate1 
            Height          =   300
            Left            =   2565
            TabIndex        =   44
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "검사결과일 :"
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
            Left            =   90
            TabIndex        =   46
            Top             =   165
            Width           =   1125
         End
         Begin VB.Label Label8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "-"
            Height          =   285
            Left            =   2430
            TabIndex        =   45
            Top             =   135
            Width           =   195
         End
      End
      Begin BHButton.BHImageButton cmdPrint 
         Height          =   420
         Left            =   6510
         TabIndex        =   48
         Top             =   90
         Visible         =   0   'False
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   741
         Caption         =   "결과 Print(&P)"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   4710
         Left            =   90
         TabIndex        =   17
         Top             =   3330
         Visible         =   0   'False
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   8308
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
      Begin VB.ListBox List1 
         Height          =   2400
         ItemData        =   "frmComm.frx":7EE3
         Left            =   90
         List            =   "frmComm.frx":7EE5
         TabIndex        =   49
         Top             =   5640
         Visible         =   0   'False
         Width           =   11055
      End
      Begin HSCotrol.UserPanel pnlCom2 
         Height          =   5415
         Left            =   1560
         TabIndex        =   7
         Top             =   2520
         Visible         =   0   'False
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   9551
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtCOM2 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4395
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   16
            Top             =   270
            Width           =   5700
         End
         Begin VB.Frame Frame2 
            Height          =   645
            Left            =   90
            TabIndex        =   8
            Top             =   4635
            Width           =   5760
            Begin MSComDlg.CommonDialog cdlFile 
               Left            =   5265
               Top             =   60
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin HSCotrol.CButton cmdChksum 
               Height          =   360
               Left            =   2205
               TabIndex        =   9
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "SUM"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMOutput2 
               Height          =   360
               Left            =   1155
               TabIndex        =   10
               Top             =   180
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Send"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMClear2 
               Height          =   360
               Left            =   3600
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Clear"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdCOMInput2 
               Height          =   360
               Left            =   90
               TabIndex        =   12
               Top             =   180
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Receive"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMLoad 
               Height          =   360
               Left            =   4635
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "File Load"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdACK 
               Height          =   360
               Left            =   3105
               TabIndex        =   14
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "ACK"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdENQ 
               Height          =   360
               Left            =   2655
               TabIndex        =   15
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "ENQ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
         End
      End
      Begin FPSpreadADO.fpSpread spdResult2 
         Height          =   7365
         Left            =   -74910
         TabIndex        =   52
         Top             =   900
         Width           =   15075
         _Version        =   524288
         _ExtentX        =   26591
         _ExtentY        =   12991
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   11
         SpreadDesigner  =   "frmComm.frx":7EE7
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   495
         Left            =   90
         TabIndex        =   64
         Top             =   360
         Width           =   11295
         _Version        =   65536
         _ExtentX        =   19923
         _ExtentY        =   873
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Begin VB.CheckBox chkUF100 
            Caption         =   "UF1000 오더 감추기"
            ForeColor       =   &H00C00000&
            Height          =   345
            Left            =   7200
            Style           =   1  '그래픽
            TabIndex        =   88
            Top             =   90
            Value           =   2  '연회색
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DptDate 
            Height          =   345
            Left            =   930
            TabIndex        =   80
            Top             =   90
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   609
            _Version        =   393216
            Format          =   121700353
            CurrentDate     =   41800
         End
         Begin BHButton.BHImageButton cmdFind_Barcode 
            Height          =   375
            Left            =   3840
            TabIndex        =   76
            Top             =   60
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   661
            Caption         =   "바코드 찾기"
            CaptionChecked  =   "BHImageButton1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ImgOutLineSize  =   3
         End
         Begin VB.CheckBox chkOption 
            Caption         =   "환자추가정보 보기"
            ForeColor       =   &H00C00000&
            Height          =   345
            Left            =   5160
            Style           =   1  '그래픽
            TabIndex        =   69
            Top             =   90
            Value           =   1  '확인
            Width           =   1935
         End
         Begin BHButton.BHImageButton cmdView 
            Height          =   375
            Left            =   8610
            TabIndex        =   65
            Top             =   60
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   661
            Caption         =   "화면확대"
            CaptionChecked  =   "화면축소"
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
            ForeColor       =   255
            ImgOutLineSize  =   3
         End
         Begin BHButton.BHImageButton cmdAppend 
            Height          =   375
            Index           =   0
            Left            =   9990
            TabIndex        =   66
            Top             =   60
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   661
            Caption         =   "결과통보"
            CaptionChecked  =   "화면축소"
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
            ForeColor       =   12582912
            ImgOutLineSize  =   3
         End
         Begin BHButton.BHImageButton cmdMachine 
            Height          =   375
            Left            =   6600
            TabIndex        =   68
            Top             =   60
            Visible         =   0   'False
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   661
            Caption         =   "장비(사용자)코드설정"
            CaptionChecked  =   "화면축소"
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
            ImgOutLineSize  =   3
         End
         Begin BHButton.BHImageButton cmdBarcode 
            Height          =   375
            Left            =   2550
            TabIndex        =   74
            Top             =   60
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   661
            Caption         =   "바코드 수정"
            CaptionChecked  =   "BHImageButton1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ImgOutLineSize  =   3
         End
         Begin VB.Label Label6 
            Caption         =   "작업일자"
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
            Height          =   255
            Left            =   90
            TabIndex        =   81
            Top             =   180
            Width           =   825
         End
      End
      Begin Threed.SSPanel ssInformation 
         Height          =   495
         Left            =   11400
         TabIndex        =   67
         Top             =   360
         Width           =   3765
         _Version        =   65536
         _ExtentX        =   6641
         _ExtentY        =   873
         _StockProps     =   15
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
      End
      Begin FPUSpreadADO.fpSpread spdResult1 
         Height          =   7395
         Left            =   90
         TabIndex        =   53
         Top             =   900
         Width           =   11295
         _Version        =   524288
         _ExtentX        =   19923
         _ExtentY        =   13044
         _StockProps     =   64
         ColsFrozen      =   10
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   14
         MaxRows         =   10
         RowsFrozen      =   1
         SpreadDesigner  =   "frmComm.frx":875A
         CellNoteIndicatorColor=   16777215
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   3
         Left            =   -74640
         TabIndex        =   18
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":9113
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   -74910
         TabIndex        =   19
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":9595
      End
      Begin VB.TextBox txtHelp 
         Height          =   3225
         Left            =   11370
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   82
         Text            =   "frmComm.frx":9A03
         Top             =   5070
         Width           =   3765
      End
      Begin FPUSpreadADO.fpSpread spdRstview 
         Height          =   7395
         Left            =   11400
         TabIndex        =   54
         Top             =   900
         Width           =   3765
         _Version        =   524288
         _ExtentX        =   6641
         _ExtentY        =   13044
         _StockProps     =   64
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   10
         RowsFrozen      =   1
         SpreadDesigner  =   "frmComm.frx":9A09
      End
      Begin VB.Label Label7 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "::::: Happy Call Center : 0505-831-1515 :::::"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   10920
         TabIndex        =   51
         Top             =   60
         Width           =   4290
      End
      Begin VB.Label Label10 
         Caption         =   "● Sample Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   7950
         TabIndex        =   50
         Top             =   5790
         Width           =   2085
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "재검/QC :"
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
         Left            =   8040
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents fSocket  As clsFTP
Attribute fSocket.VB_VarHelpID = -1

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

Private Const sTo1 As String = "G1"

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const fs  As String = ""
Const RS  As String = ""

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
Dim Patiant_Recevid As Integer
Dim sStxCheck As Integer
Dim sEtxCheck As Integer


' --------------------------------------------------------------
Private Type Type_Urisys2400
    BarCode_NO    As String
    Chart_No      As String
    
    SEQ           As String
    RackNo        As String
    Position      As String '1~5
    Type          As String '1~5
    
    PID           As String
    PNM           As String
    
    TestDate      As String
    TestTime      As String
    RunType       As String 'N, E, R, C, S, B
    
    Order_Str     As String

End Type

Dim Urisys2400 As Type_Urisys2400


Dim strOrdLst As String

Dim fUrisys2400()   As String
Dim fUrisys2400_SUB()   As String

Dim fUrisys2400Size(100, 1) As Integer
Dim Immulite2000_1   As Variant
Dim SendData(10)     As String
Dim SendCount        As String
Dim Or_Seq           As Integer
Dim SendBuffW           As String
Dim SendBuffT           As String
Dim intRow          As Integer
Dim brStr           As String

Dim cntCheckSum      As Integer
Dim ReceiveData      As String

Dim SendFlg          As Boolean
Dim HostOutput       As String

Public WithEvents Result As clsMsg_Result
Attribute Result.VB_VarHelpID = -1
Public WithEvents Order  As clsMsg_Query
Attribute Order.VB_VarHelpID = -1
Public Result1 As clsResult

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

Private MSG_STX     As String
Private MSG_ETX     As String
Private MSG_ENQ     As String
Private MSG_EOT     As String
Private MSG_ACK     As String
Private MSG_NAK     As String
Private MSG_CR      As String
Private MSG_LF      As String
Private MSG_CRLF    As String

Dim fCellDynCfg(100) As Integer
Dim fCellDynSize(100, 1) As Integer
Dim fChannel() As String
Dim pname   As String
Dim pNo     As String
Dim chkEnq  As Integer

Dim flgETB           As Boolean
Dim flgETX           As Boolean

Private Type SugaMatch
    TestId          As String
    Sugacd          As String
    Testcd          As String
    DecPoint        As Long
End Type

Dim SMatch() As SugaMatch
Dim CountTest As Integer, sErrorFlag As Boolean
Dim strBarno As String

Private Type TYPE_CD
    strEqpCd    As String
    intCnt      As Integer
    strTestcd(100) As String
    strSEQ(100)   As String
    strACK(100)   As String
End Type

Private f_typCode() As TYPE_CD

Dim RecordChk As Boolean

Dim fRcvString As String
Dim PatientID  As String    'Q Message Pattern Check
Dim PatientChart As String    'Q Message Pattern Check

Dim PatientSeq As String
Dim PatientDisk As String
Dim PatientPos As String

Dim sPatiant_No As String

'
' 읽어 올때..
'
Dim tDR_TO  As String
Dim tDR_DATE   As String
Dim tDR_ATTEND   As String
Dim tDR_SERIAL   As String
Dim tDR_CHART   As String
Dim tDR_ORIGIN   As String
Dim tDR_YU_NO   As String
Dim tDR_YUHYUNG   As String
Dim tDR_GAMOK   As String
Dim tDR_ISERIAL  As String
 
'
' 결과 저장 할때..
'
Dim stGumDat_TO  As String
Dim stGumDat_DATE As String
Dim stGumDat_ATTEND As String
Dim stGumDat_ISERIAL As String
Dim stGumDat_CHART As String

Dim stGumDat_YUHYUNG As String
Dim stGumDat_ORIGIN  As String
Dim stGumDat_YUNO As String
Dim stGumDat_GAMOK As String
Dim stGumDat_CODE As String
Dim stGumDat_NAME As String
Dim stGumDat_RESULT As String
Dim stGumDat_1DATE As String
Dim stGumDat_1TIME As String
Dim stGumDat_1ID  As String
Dim stGumDat_1COM     As String
Dim tRCNT  As Boolean

Private fHostIP     As String
Private fPort       As String
Private fID         As String
Private fPW         As String
Private fDir        As String

Private 검사일자 As String
Private 검사순   As String
Private 장비코드  As String

Const gItemStartPos As Integer = 14

Dim gWideOption As Boolean
Private adoRS               As ADODB.Recordset
Private AdoCmd              As ADODB.Command

'''''
Dim strRecvData()   As String
Dim intPhase        As Integer
Dim strState        As String
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer
'''''


'===== User Define
'인터페이스에서 사용
Dim RcvBuffer       As String
Dim wkBuf           As String
Dim sState          As String
Dim sReqStatusCd    As String

Dim bEndChk As Boolean
Dim bSTXChk As Boolean
Dim sNextSend   As String
Dim RstEnd      As String

Dim m_iPhase      As Integer
Dim m_iSendPhase  As Integer
Dim m_p_sRerunGbn As String
Dim m_iFrameN     As Integer

Dim gLocalIP      As String
Dim gLocalNm      As String

Dim pNIT_Check As Boolean
Dim pWBC_Check As Boolean
Dim pRBC_Check As Boolean

Dim miLeaveCell%

Dim OldRow As Long


Private Sub chkOption_Click()
With spdResult1
    If chkOption.Value = 1 Then
        chkOption.Caption = "환자추가정보 보기"
        .Col = 4: .ColHidden = True
        .Col = 6: .ColHidden = True
        .Col = 7: .ColHidden = True
        .Col = 8: .ColHidden = True

        
    Else
        chkOption.Caption = "환자추가정보 감추기"
        .Col = 4: .ColHidden = False
        .Col = 6: .ColHidden = False
        .Col = 7: .ColHidden = False
        .Col = 8: .ColHidden = False

        
    End If
End With
End Sub

Private Sub chkUF100_Click()

    If chkUF100.Value = 1 Then
        chkUF100.Caption = "UF1000 오더 보기"
        uplUrineMicro.Visible = True

        
    Else
        chkUF100.Caption = "UF1000 오더 감추기"
        uplUrineMicro.Visible = False

        
    End If

End Sub

Private Sub cmdBarcode_Click()
    uplBarcode.Visible = True
    txtBarcode_Keyin.text = ""
    txt순번.text = spdResult1.ActiveRow
    txtBarcode_Keyin.SetFocus
    
End Sub

Private Sub cmdBarcode_Close_Click()
    uplBarcode.Visible = False
End Sub

Private Sub cmdExcel_Export_Click()
Dim sFile As String

    MousePointer = vbHourglass
    On Error GoTo errExit
    With cdgExport
        .CancelError = True
        .DialogTitle = "UF100 검사현황"
        .DefaultExt = "CSV"
        .Filter = "Excel(*.CSV)|*.csv|모든파일(*.*)|*.*"
                .FileName = Format(DptDate.Value, "YYYY-MM-DD") & "_UF100 검사현황"
        .ShowSave
      '  .FileName = DptDate.Value & "UF100 검사현황"
        sFile = .FileName
        If Len(sFile) > 0 Then
            If spdMicro.ExportToTextFile(sFile, "", ",", Chr(13), ExportToTextFileColHeaders + ExportToTextFileCreateNewFile, "Export.Log") Then
                MsgBox "UF100 검사현황을 " & sFile & "로 변환하였습니다.!", vbInformation
            Else
                MsgBox Err.Description, vbCritical
            End If
        End If
    End With
errExit:
    MousePointer = vbDefault

End Sub

Private Sub cmdFind_Barcode_Click()
Dim p바코드번호 As String
Dim pGrid_Point As Integer

    p바코드번호 = InputBox("바코드번호를 입력하세요 !", "바코드번호 찿기")
    If Len(p바코드번호) > 0 Then
        pGrid_Point = SeqSearch(spdResult1, p바코드번호, 3)
        If pGrid_Point > 0 Then
            spdResult1.Row = pGrid_Point
        Else
            MsgBox (vbCrLf & "【 " & p바코드번호 & " 】" & "는 없는 바코드번호 입니다.        " & vbCrLf & vbCrLf)
        End If
    End If
End Sub

'Private Sub cmdMachine_Click()
'Dim t검사순 As String
'
'    t검사순 = InputBox("Seq No 를 입력 하십시요..", "Seq No설정", 검사순)
'
'    If Len(Trim(t검사순)) <> 0 Then
'        Call SaveString(HKEY_CURRENT_USER, REG_Machine, REG_Code, Trim(t검사순))
'        MsgBox "저장 장비코드가 설정되었습니다.", vbInformation + vbOKOnly, App.Title
'        장비코드 = t장비코드
'    End If
'End Sub

Private Sub cmdView_Click()

    If gWideOption = False Then
        spdResult1.Width = spdResult1.Width + spdRstview.Width
        gWideOption = True
    Else
        spdResult1.Width = spdResult1.Width - spdRstview.Width
        gWideOption = False
    End If
    
End Sub

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
    SumCS = Hex(AddCS)
    ChkCS = Mid(SumCS, Len(SumCS) - 1, 1)
    ChkCS = ChkCS & Right(SumCS, 1)
    MakeCS = ChkCS
End Function

Private Function f_funGet_SpreadRow(ByVal objSpd As vaSpread, ByVal intCol As Integer, _
                                    ByVal strPara As String) As Integer

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    f_funGet_SpreadRow = 0
    
    With objSpd
        For intRow = 1 To .maxrows
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
            Call .Add(, "REFL", "참고치(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFH", "참고치(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "AUTOVERIFY", "재검", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, "REMARK", "검체코드", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, "OUTSEQ", "OUTSEQ", (lvwCuData.Width - 310) * 0.1)
           
        End With
        .HideColumnHeaders = False
    End With
    
   
End Sub

Private Function f_subSet_WorkList(ByVal strDATE As String)
    Dim sqlRet      As Integer
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
    Set AdoRs_SQL = New ADODB.Recordset
       
    gSql = "       Select a.*, b.수진자명,b.챠트번호,b.주민등록번호 from TB_검사항목 a, TB_인적사항 b"
    gSql = gSql & " Where a.진료년+a.진료월+a.진료일 = '" & strDATE & "'"
    gSql = gSql & "   and a.진료지원상태 = '1' "
    gSql = gSql & "   and a.처방코드     in('C3360','C4802','C4812','C4712','C3520','C3290','C3340') "
    gSql = gSql & "   and a.챠트번호     = b.챠트번호"
    gSql = gSql & "   and a.진료지원상태 < 5  "
    gSql = gSql & "   and a.처방번호     > 0  "
    gSql = gSql & " Order by a.챠트번호"
    
    AdoRs_SQL.Open gSql, AdoCn_SQL, adOpenStatic, adLockReadOnly
   
    If AdoRs_SQL.RecordCount = 0 Then
        Set f_subSet_WorkList = Nothing
        tRCNT = False
    Else
        Set f_subSet_WorkList = AdoRs_SQL
        tRCNT = True
    End If

    Set AdoRs_SQL = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    Call ErrMsgProc(CallForm)

    
End Function

Private Function f_Barcode_ReqNo(ByVal strBarno As String) As String
Dim pDateCount As Double
Dim pBartmp    As Double
Dim pDateVal   As Double
Dim pYYYYMMDD  As String

'
' 녹십자 ReqNo 만들기
'
    
    pDateCount = DateDiff("d", "2000/01/01", Now)
    pBartmp = Mid(strBarno, 1, 4)
    
    pDateVal = pBartmp - pDateCount
    
    pYYYYMMDD = Format(Now + pDateVal, "YYYYMMDD")
    f_Barcode_ReqNo = pYYYYMMDD & Mid(strBarno, 5, 8)


End Function

Private Function f_DeleTe_MSSQL(ByVal strSEQ As String, ByVal strDATE As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim strYY As String
    Dim strMM As String
    Dim strDD As String
    Dim strChart As String
    Dim strNo As String
    Dim pReqNo As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList_Barcode() As ADODB.Recordset"
    
        sqlDoc = ""
        sqlDoc = sqlDoc + " DELETE from IFRESULT01                     " & vbCrLf
        sqlDoc = sqlDoc + "  Where WDATE =  '" & strDATE & "'     " & vbCrLf
        sqlDoc = sqlDoc + "    and WSEQ  =  '" & strSEQ & "'      "
        
        Call DBExec(AdoCn_SQL, sqlDoc, CallForm)
        
Exit Function

ErrorTrap:

   
End Function



Private Function f_subSet_WorkList_Barcode(ByVal strBarno As String, Optional ByVal strINS_CODE As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim strYY As String
    Dim strMM As String
    Dim strDD As String
    Dim strChart As String
    Dim strNo As String
    Dim pReqNo As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList_Barcode() As ADODB.Recordset"
    
        If Not IsNumeric(strBarno) Then
            Set f_subSet_WorkList_Barcode = Nothing
            RecordChk = False
            Set AdoRs_ORACLE = Nothing
            Exit Function
        End If
        
        pReqNo = f_Barcode_ReqNo(strBarno)
        pReqNo = Mid(pReqNo, 1, Len(pReqNo) - 1)
                
        sqlDoc = ""
        sqlDoc = sqlDoc + " Select *                              " & vbCrLf
        sqlDoc = sqlDoc + "   From MCHORDER                       " & vbCrLf
        sqlDoc = sqlDoc + "  Where REQNO =  '" & pReqNo & "'      "
        
        Set AdoRs_ORACLE = New ADODB.Recordset
        AdoRs_ORACLE.CursorLocation = adUseClient
        AdoRs_ORACLE.Open sqlDoc, AdoCn_ORACLE
        
        If AdoRs_ORACLE.RecordCount = 0 Then
            Set f_subSet_WorkList_Barcode = Nothing
            RecordChk = False
            Set AdoRs_ORACLE = Nothing
            Exit Function
        Else
            Set f_subSet_WorkList_Barcode = AdoRs_ORACLE
            RecordChk = True
        End If
    
        Set AdoRs_ORACLE = Nothing
           
Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    
    Call ErrMsgProc(CallForm)

    
End Function

Private Function f_subSet_SearchOrder(ByVal strBarcode As String, ByVal strTestcd As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
    
    Set AdoRs_SQL = New ADODB.Recordset
    
    sqlDoc = ""
    sqlDoc = sqlDoc + " Select *                              " & vbCrLf
    sqlDoc = sqlDoc + "   From MCHORDER                       " & vbCrLf
    sqlDoc = sqlDoc + "  Where 1=1                            " & vbCrLf
    sqlDoc = sqlDoc + "    And REQNO  =  '" & strBarcode & "' " & vbCrLf
    sqlDoc = sqlDoc + "    And ITEMCD =  '" & strTestcd & "'  "
    
    AdoRs_ORACLE.Open sqlDoc, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
    
    If AdoRs_ORACLE.RecordCount = 0 Then
        Set f_subSet_SearchOrder = Nothing
    Else
        Set f_subSet_SearchOrder = AdoRs_ORACLE
    End If

    Set AdoRs_ORACLE = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    Call ErrMsgProc(CallForm)

    
End Function

Private Sub f_subSet_ItemList()

    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strTest As String, intPos   As Integer
    Dim strTmp  As String, intCol   As Integer, intCol2   As Integer, intCnt  As Integer, intRow  As Integer
    Dim strTmpSeq As String
    Dim strtmpACK As String

    
'On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    'lvwCuData.ListItems.Clear:
    f_strOrdList = ""
    
    intCol = gItemStartPos
    intCol2 = 1
    intRow = 1
    

    
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 2
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    With spdResult2
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
   
    Set adoRS = Nothing
        
    sqlDoc = "Select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH,TESTCD_EQP " & _
             "  From INTERFACE002" & _
             " Where (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " Order by OUT_SEQ, TESTCD_EQP"

             
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst: ReDim fChannel(adoRS.RecordCount)
    Do While Not adoRS.EOF
        Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TEST_EQP") & ""), , "LST")
            itemX.SubItems(1) = Trim(adoRS.Fields("TESTCD") & "")
            itemX.SubItems(2) = Trim(adoRS.Fields("TESTNM") & "")
            itemX.SubItems(3) = Trim(adoRS.Fields("TESTCD_EQP") & "")
            itemX.SubItems(4) = Trim(adoRS.Fields("DELTA") & "")
            itemX.SubItems(5) = Trim(adoRS.Fields("DELTAGBN") & "")
            itemX.SubItems(6) = Trim(adoRS.Fields("PANICL") & "")
            itemX.SubItems(7) = Trim(adoRS.Fields("PANICH") & "")
            itemX.SubItems(8) = Trim(adoRS.Fields("REFL") & "")
            itemX.SubItems(9) = Trim(adoRS.Fields("REFH") & "")
            itemX.SubItems(10) = Trim(adoRS.Fields("AUTOVERIFY") & "")
            itemX.SubItems(11) = Trim(adoRS.Fields("REMARK") & "")
            itemX.SubItems(12) = Trim(adoRS.Fields("OUT_SEQ") & "")
            
            itemX.tag = Trim(adoRS.Fields("TEST_EQP") & "")
            itemX.text = Trim(adoRS.Fields("TESTCD") & "")
            
            
            Debug.Print Trim(adoRS.Fields("TEST_EQP") & "") & " | " & Trim(adoRS.Fields("REMARK") & "")
        Set itemX = Nothing
        
        
        With spdResult1
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
                .ColWidth(intCol) = 8.5
            End If
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")

        End With
        
        With spdRstview
            If intRow > .maxrows Then
                intRow = 1
                intCol2 = intCol2 + 2
            End If
            
            .SetText intCol2, intRow, Trim$(adoRS("TESTNM") & "")
            intRow = intRow + 1
            
        End With
        
        With spdResult2
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
                .ColWidth(intCol) = 8.5
            End If
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        fChannel(intCol - gItemStartPos) = adoRS.Fields("TEST_EQP")
        
        intCnt = intCnt + 1
        ReDim Preserve f_typCode(1 To intCnt) As TYPE_CD
        
        f_typCode(intCnt).strEqpCd = Trim$(adoRS.Fields("TEST_EQP"))
        'f_typCode(intCnt).strSEQ = Trim$(adoRS.Fields("OUT_SEQ"))
        
        f_typCode(intCnt).intCnt = 0
        
        strTmp = Trim$(adoRS.Fields("TESTCD"))
        strTmpSeq = Trim$(adoRS.Fields("OUT_SEQ"))
        strtmpACK = Trim$(adoRS.Fields("REMARK"))
        
        intPos = InStr(strTmp, ",")
        Do While intPos > 0
            f_strOrdList = f_strOrdList + "'" + Mid$(strTmp, 1, intPos - 1) + "',"
            
            f_typCode(intCnt).intCnt = f_typCode(intCnt).intCnt + 1
            f_typCode(intCnt).strTestcd(f_typCode(intCnt).intCnt) = Mid$(strTmp, 1, intPos - 1)
            f_typCode(intCnt).strSEQ(f_typCode(intCnt).intCnt) = strTmpSeq
            f_typCode(intCnt).strACK(f_typCode(intCnt).intCnt) = strtmpACK
           
            
            strTmp = Mid$(strTmp, intPos + 1)
            
            intPos = InStr(strTmp, ",")
        Loop
        f_strOrdList = f_strOrdList + "'" + strTmp + "',"
        f_typCode(intCnt).intCnt = f_typCode(intCnt).intCnt + 1
        f_typCode(intCnt).strTestcd(f_typCode(intCnt).intCnt) = strTmp
        
        intCol = intCol + 1
        
        adoRS.MoveNext
    Loop
    Set adoRS = Nothing
    
     With spdResult2
        If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
        .SetText intCol, 0, ""
        .Col = intCol:  .ColHidden = True
    End With
    
'    sqlDoc = "Select max(wseq) as MaxNo      " & _
'             "  From dbo.IFRESULT01          " & _
'             " Where WDATE = " & Format(Now, "YYYYMMDD") & " "
'
'    adoRS.CursorLocation = adUseClient
'    adoRS.Open sqlDoc, AdoCn_SQL
'    adoRS.MoveFirst
'    pSeq = Work_seq
    

Exit Sub
ErrRoutine:
    Set adoRS = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub

Private Function f_funGet_CODE_ACK(ByVal strOrdcd As String) As String
    Dim intIdx1 As Integer, intIdx2 As Integer
    
    f_funGet_CODE_ACK = 0
    
    For intIdx1 = 1 To UBound(f_typCode)
        For intIdx2 = 1 To f_typCode(intIdx1).intCnt
            If Trim$(strOrdcd) = Trim$(f_typCode(intIdx1).strTestcd(intIdx2)) Then
                f_funGet_CODE_ACK = f_typCode(intIdx1).strACK(intIdx2)
                Exit Function
            End If
        Next
    Next
    
End Function

Private Function f_funGet_CODE_SUB(ByVal strOrdcd As String) As Integer
    Dim intIdx1 As Integer, intIdx2 As Integer
    
    f_funGet_CODE_SUB = 0
    
    For intIdx1 = 1 To UBound(f_typCode)
        For intIdx2 = 1 To f_typCode(intIdx1).intCnt
            If Trim$(strOrdcd) = Trim$(f_typCode(intIdx1).strTestcd(intIdx2)) Then
                f_funGet_CODE_SUB = f_typCode(intIdx1).strSEQ(intIdx2)
                Exit Function
            End If
        Next
    Next
    
End Function

Private Function f_funGet_CODE(ByVal strOrdcd As String) As String

    Dim intIdx1 As Integer, intIdx2 As Integer
    
    f_funGet_CODE = ""
    
    For intIdx1 = 1 To UBound(f_typCode)
        For intIdx2 = 1 To f_typCode(intIdx1).intCnt
            If Trim$(strOrdcd) = Trim$(f_typCode(intIdx1).strTestcd(intIdx2)) Then
                f_funGet_CODE = f_typCode(intIdx1).strEqpCd
                Exit Function
            End If
        Next
    Next
    
End Function

Private Sub cmdLongSelect_Click()
    If cmdLongSelect.Caption = "▽" Then
        cmdLongSelect.Caption = "△"
        spdResult1.Height = 7305
    Else
        cmdLongSelect.Caption = "▽"
        spdResult1.Height = 4335
    End If
End Sub

Private Sub cmdPosNo_Click()
Dim sNo As String, sCnt As Integer, sAdd As Integer

AgainInput:
    sNo = InputBox("시작 번호를 입력하세요 !")
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
        If Not IsNumeric(sNo) Then
            MsgBox "숫자만 입력하세요.!", vbCritical
            GoTo AgainInput
        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .maxrows
                .Row = sCnt
                .Col = 7:       .text = Trim(sAdd + Val(sNo))
                sAdd = sAdd + 1
            Next sCnt
        End With
    End If
End Sub

Private Sub cmdPrint_Click()

Dim objclsCommon As New clsCommon

Dim Tmp_Testnm As String
Dim Row_cnt As Integer, Col_cnt As Integer, TmpPrintline As Integer
Dim vTmp As Variant
Dim vTestNm  As Variant
Dim stragesex As String

Const TmpLine = "────────────────────────────────────────────────────────────"

    If spdResult1.maxrows >= 1 Then
        With objclsCommon
            .PrintText 15, 3, Format(Date, "yyyy/mm/dd") & "  Result Report..( " & App.EXEName & " )", "Arial", 12
            
            .PrintText 0.5, 5, TmpLine
            .PrintText 0.5, 6, "순", , 9
            .PrintText 2, 6, "접수일자", , 9
            .PrintText 6, 6, "차트번호", , 9
            .PrintText 10, 6, "이  름", , 9
            .PrintText 14, 6, "구 분", , 9
            .PrintText 17, 6, "검사종목[결과]", , 9
            .PrintText 0.5, 7, TmpLine
            
            TmpPrintline = 8
        
        For Row_cnt = 1 To spdResult1.maxrows
            spdResult1.Row = Row_cnt
            
            If (Row_cnt Mod 34) <> 0 Then
                                    .PrintText 0.5, TmpPrintline, Row_cnt, , 9                                                          ' 순
            
                spdResult1.Col = 2: .PrintText 2, TmpPrintline, Mid(spdResult1.text, 3), , 9                                                    ' 처방일자
                spdResult1.Col = 3: .PrintText 6, TmpPrintline, Trim(spdResult1.text), , 9                                             ' 이    름
                spdResult1.Col = 4: .PrintText 10, TmpPrintline, Trim(spdResult1.text), , 9                                             ' 이    름
                spdResult1.Col = 5: .PrintText 14, TmpPrintline, Trim(spdResult1.text), , 9                                             ' 이    름
                
                For Col_cnt = 10 To spdResult1.MaxCols
            
                    spdResult1.Col = Col_cnt
                    
                    If Trim(spdResult1.text) <> "" Then
                    
                        spdResult1.GetText Col_cnt, 0, vTestNm
                        spdResult1.GetText Col_cnt, Row_cnt, vTmp

                        
                        Tmp_Testnm = Tmp_Testnm & vTestNm & "[" & vTmp & "]" & " / "
                        
                        If Len(Tmp_Testnm) Mod 60 < 0 Then
                            Tmp_Testnm = Tmp_Testnm & vbCrLf
                        End If
                    End If
                    
                Next Col_cnt
                
                spdResult1.Col = 5: .PrintText 18, TmpPrintline, Trim(Tmp_Testnm), , 8
                
                TmpPrintline = TmpPrintline + 2
                Tmp_Testnm = ""
            Else
            
                '-------------------------------------------------------
            
                                    .PrintText 0.5, TmpPrintline, Row_cnt, , 9                                                         ' 순
            
                spdResult1.Col = 2: .PrintText 2, TmpPrintline, Mid(spdResult1.text, 3), , 9                                           ' 처방일자
                spdResult1.Col = 3: .PrintText 6, TmpPrintline, Trim(spdResult1.text), , 9                                             ' 이    름
                spdResult1.Col = 4: .PrintText 10, TmpPrintline, Trim(spdResult1.text), , 9                                            ' 이    름
                spdResult1.Col = 5: .PrintText 14, TmpPrintline, Trim(spdResult1.text), , 9                                            ' 이    름
                
                For Col_cnt = 10 To spdResult1.MaxCols
            
                    spdResult1.Col = Col_cnt
                    
                    If Trim(spdResult1.text) <> "" Then
                    
                        spdResult1.GetText Col_cnt, 0, vTestNm
                        spdResult1.GetText Col_cnt, Row_cnt, vTmp

                        
                        Tmp_Testnm = Tmp_Testnm & vTestNm & "[" & vTmp & "]" & " / "
                        
                        If Len(Tmp_Testnm) Mod 60 < 0 Then
                            Tmp_Testnm = Tmp_Testnm & vbCrLf
                        End If
                    End If
                    
                Next Col_cnt
                
                spdResult1.Col = 5: .PrintText 18, TmpPrintline, Trim(Tmp_Testnm), , 8
                
                TmpPrintline = TmpPrintline + 2
                Tmp_Testnm = ""
                
                '-------------------------------------------------------
            
                    .PrintText 0.5, TmpPrintline, TmpLine
                    .PrintText 1, TmpPrintline + 1, "── Next  Report ──", , 9, True
                    Printer.NewPage
                    
                    .PrintText 0.5, 5, TmpLine
                    .PrintText 0.5, 6, "순", , 9
                    .PrintText 2, 6, "접수일자", , 9
                    .PrintText 6, 6, "차트번호", , 9
                    .PrintText 10, 6, "이  름", , 9
                    .PrintText 14, 6, "구 분", , 9
                    .PrintText 17, 6, "검사종목[결과]", , 9
                    .PrintText 0.5, 7, TmpLine
                    
                    TmpPrintline = 9
            End If
        
        Next Row_cnt
        .PrintText 0.5, TmpPrintline, TmpLine
        .PrintText 1, TmpPrintline + 1, "── End  of  Report ──", , 9, True
        
        End With
        Printer.NewPage
        Printer.EndDoc
        
        MsgBox Format(Date, "yyyy/mm/dd") & "일자의 " & App.EXEName & "의 장비 검사 Result가 Print되었습니다..       " & vbCrLf & vbCrLf & "다음 작업을 진행하십시요..", vbInformation + vbOKOnly, App.EXEName
    Else
        MsgBox Format(Date, "yyyy/mm/dd") & "일자의 " & App.EXEName & "의 장비 검사 Result가  Load 되어 있지 않습니다..       " & vbCrLf & vbCrLf & "자료를 확인 하십시요..", vbInformation + vbOKOnly, App.EXEName
    End If
End Sub

Private Sub cmdRackNo_Click()
    Dim sNo As String, sCnt As Integer, sAdd As Integer
    Dim aROW    As Integer, aCOL   As Integer
    Dim varChk  As Variant, varBar As Variant, varNum As Variant
    Dim iRow    As Integer, iCnt   As Integer
    Dim strRack_tmp As String
        

AgainInput:
    sNo = InputBox("시작 번호를 입력하세요 !")
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
        If Not IsNumeric(sNo) Then
            MsgBox "숫자만 입력하세요.!", vbCritical
            GoTo AgainInput
        End If
               
        With spdResult1
            iCnt = 1
            .GetText 1, 1, varChk
            .GetText 2, 1, varBar
            varNum = sNo
            If Trim(varChk) = "1" And Trim(varBar) <> "" Then
                For iRow = 1 To .maxrows
                    .SetText 6, iRow, varNum
                    .SetText 7, iRow, ((iCnt Mod 101) + 1) - 1
                    iCnt = iCnt + 1
                    If (iCnt Mod 101) = 1 Then varNum = varNum + 1
                Next
            End If
        End With
    End If
End Sub

Private Sub cmdACK_Click()
    Call COM_OUTPUT(Chr(1))
End Sub

Private Sub cmdAction_Click(Index As Integer)
    
    Select Case Index
        Case 0:     Call cmdRun
        Case 1:     Call cmdStop
        Case 2:     Call cmdClear
        Case 3:     Call cmdExit
        Case Else
    End Select
    
    intRow = 0
    
End Sub

Private Sub cmdClear()
        
    List1.Clear
    
    f_strJOB_FLAG = "1"
    f_intSampleNo = 0
    Or_Seq = 1
    
    ssInformation.Caption = ""
    txtHelp.text = ""
    
    Dim iiresult As Integer
    
    With spd저장체크
        For iiresult = 1 To .maxrows
            .Row = iiresult
            .Col = 1: .text = ""
            .Col = 2: .text = ""
            .Col = 3: .text = ""
            .Col = 4: .text = ""
            .Col = 5: .text = ""
        Next
    End With
    
    With spdMicro
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 0
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    With spdResult1
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
        .RowHeight(-1) = 14
        
    End With

    With spdResult2
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    With spdRstview
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With

End Sub

Private Sub cmdExit()
    
    Unload Me

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

Private Sub cmdAppend_Click(Index As Integer)
    Dim adoRS           As New ADODB.Recordset
    Dim sqlDoc          As String

    Dim varTmp          As Variant
    Dim strErrMsg       As String
    Dim strSampleno()   As String, strBarno     As String, strTime      As String
    Dim strOrdcd()      As String, strRstval    As String, intCnt       As Integer
    Dim strHL           As String
    Dim strTmp1()       As String, strTmp2()    As String
    Dim intPos          As String, strTestcd    As String, strTestRst   As String
    Dim strTestnm       As String
    Dim strRef          As String
    Dim strUnit         As String
    Dim strOrdLst()     As String, strPid()    As String, strPnm() As String

    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim intIdx      As Integer
    Dim blnFlag     As Boolean
    Dim itemX       As ListItem
    Dim objSpd      As vaSpread
    Dim sqlRet      As Integer
    Dim flgSave     As Boolean
    Dim SaveGbn     As Integer
    Dim strDATE     As String
    Dim strInDate   As String
    Dim pname       As String
    Dim pNo         As String

    Dim strRefVal   As String
    Dim strDelVal   As String
    Dim strPenVal   As String
    Dim strEqpCd    As String

    
    Dim pRstString  As String ' 서버 만들기 위한 변수
    Dim gRstString  As String

    Dim pSeqNo As Integer   '' DB 저장용 SEQ
    Dim strRnt As Boolean
    
    CallForm = "frmComm - Private Sub cmdAppend_Click()"

On Error Resume Next

    Me.MousePointer = 11
    
    With spdResult1
    
        For intRow = 1 To .maxrows
        
            Dim iiresult As Integer
    
            With spd저장체크
                For iiresult = 0 To .maxrows
                    .Row = iiresult
                    .Col = 1: .text = ""
                    .Col = 2: .text = ""
                    .Col = 3: .text = ""
                    .Col = 4: .text = ""
                    .Col = 5: .text = ""
                Next
            End With
        
            .GetText 3, intRow, varTmp:   strBarno = Trim$(varTmp)
            .GetText 4, intRow, varTmp:   pNo = Trim$(varTmp)
            .GetText 5, intRow, varTmp:   pname = Trim$(varTmp)
            .GetText 8, intRow, varTmp:   strDATE = Trim$(varTmp)

            .GetText 1, intRow, varTmp

            strTime = Format(Now, "HHMMSS")
            strInDate = Format(Now, "YYYYMMDD")

            If strBarno <> "" Then

                intCnt = 0: Erase strOrdcd
                
                If Trim$(varTmp) = "1" Then
                        
                    For intCol = 14 To .MaxCols
                        .GetText intCol, intRow, varTmp
                        .GetText intCol, 0, varTmp

                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                        
                        If Not itemX Is Nothing Then
                            
                           .GetText intCol, intRow, varTmp
                           .Row = intRow: .Col = intCol
                            
                            strRstval = Trim(varTmp)
                            
                            strTestcd = itemX.ListSubItems(1)
                            intPos = InStr(strTestcd, ",")
                            
                            If intPos >= 0 Then
                                strTestcd = Mid$(strTestcd, 1, intPos - 1)
                            End If
                            
                            If intPos >= 0 Then

                                Dim pOrder As String
                                
                                With spdResult1
                                    .Row = intRow
                                    .Col = intCol
                                    pOrder = .CellTag
                                    
                                    If intCol = 14 Then
                                        .CellTag = "Micro"
                                        pOrder = .CellTag
                                    End If
                                
                                End With
                                
                                If Mid(pname, 1, 2) = "QC" Then
                                    pOrder = "QC"
                                End If

                                .Col = 13: .text = "전송"
                                
                                Debug.Print pOrder & " | " & itemX.SubItems(2)
                                strRnt = Result_Upload_MSSQL(intRow, intCol, itemX.SubItems(11), Format(DptDate.Value, "YYYYMMDD"), True)
                                
                                ' 저장 된 검사를 넣어 둔다..
                                Dim pACK As Integer
                                pACK = f_funGet_CODE_ACK(strTestcd)
                                
                                With spd저장체크
                                    .Col = 1
                                    .Row = itemX.SubItems(3)
                                    .Col = 1: .text = itemX.SubItems(12)
                                    .Col = 3: .text = itemX.SubItems(2)
                                    .Col = 4: .text = strRstval
                                    
                                    If Len(pOrder) <> 0 Then  ' 오더가 있는 경우
                                        .Col = 2: .text = f_funGet_CODE_ACK(pOrder)
                                        .Col = 5: .text = "*"
                                    Else ' 오더가 없는 경우
                                        .Col = 2: .text = itemX.SubItems(11)
                                    End If
                                        
                                End With
                                    
                                '
                                ' 오더가 있는 SID 화면 색상변경..
                                '
                                If strRnt = True Then
                                    .Row = intRow
                                    .Col = 1: .Value = 0
                                    If pname <> "오더확인" Then
                                        .Col = 1: .BackColor = vbCyan
                                        .Col = 2: .BackColor = vbCyan
                                        .Col = 3: .BackColor = vbCyan
                                        .Col = 4: .BackColor = vbCyan
                                        .Col = 5: .BackColor = vbCyan
                                        .Col = 6: .BackColor = vbCyan
                                        .Col = 7: .BackColor = vbCyan
                                        .Col = 8: .BackColor = vbCyan
                                        .Col = 9: .BackColor = vbCyan
                                        .Col = 10: .BackColor = vbCyan
                                        .Col = 11: .BackColor = vbCyan
                                        .Col = 12: .BackColor = vbCyan
                                        .Col = 13: .BackColor = vbCyan
                                    End If
                                End If
                                
                            End If
                            
                        End If
    
                        Set itemX = Nothing
                        
                    Next intCol
                    
                    '
                    ' 모든 SID 저장 후 한번 만..
                    '
                    If strErrMsg = "" Then
                        .GetText 3, intRow, varTmp:   strBarno = Trim$(varTmp)
                        
                        sqlDoc = "Update INTERFACE003                        " & _
                                 "   Set SERVERGBN = 'Y'                     " & _
                                 " Where SPCNO   = '" & strBarno & "'        " & _
                                 "   And TRANSDT = '" & mskRstDate.text & "' "
                        AdoCn_Jet.Execute sqlDoc
                    Else
                        MsgBox strErrMsg, vbInformation, Me.Caption
                    End If
                    
                    Dim iresult As Integer
                    Dim m채널  As String
                    Dim pCOL   As String
                    Dim pACKCODE As String
                    
                    '
                    ' 저장된 검사 이외에 저장하는 루틴..
                    '
                    With spd저장체크
                        For iresult = 0 To spd저장체크.maxrows
                            .Row = iresult
                            .Col = 5

                            If .text = "" Then
                                With spd저장체크
                                    .Col = 1
                                    .Row = iresult
                                    .Col = 1:  pCOL = .text
                                    .Col = 2:  pACKCODE = .text
                                End With

                                strRnt = Result_Upload_MSSQL(intRow, pCOL, pACKCODE, Format(DptDate.Value, "YYYYMMDD"), False)
                            
                            End If
                        Next iresult
                    End With
                    
        
                End If
            End If
            

            
            
            
        Next intRow
        
        

        
        '
        ' 검체 트래킹
        
        '
        '
        
        
    End With
    
    Me.MousePointer = 0
    
    Exit Sub

End Sub

Public Function Result_Upload_MSSQL(ByVal strRow As String, _
                                    ByVal strCol As String, _
                                    ByVal strIFCD As String, _
                                    ByVal strWdate As String, _
                                    Optional ByVal Order_Check As Boolean) As Boolean
    Dim sqlRet  As Integer
    Dim adoRS   As New ADODB.Recordset
    
    Dim pjno   As String     ' 바코드번호
    Dim pifseq  As String    ' ACK ID
    Dim pRack  As String     ' RACK
    Dim pPos  As String      ' POS
    Dim pwdate  As String    ' 검사일자
    Dim pwseq  As String     ' SEQ
    Dim peqgbn  As String    ' 장비명
    Dim pjdate  As String    ' 처방일자
    Dim pjgbn  As String     ' WRKDTE,WORKNO,LABEMP,
    Dim pregno  As String    ' 병원명
    Dim pname  As String     ' 환자명
    Dim psex  As String      ' 성별
    Dim pemer  As String     ' REQNO
    Dim prerun  As String    ' ""
    Dim pOTHER  As String    ' 시간
    Dim pResult1  As String  ' 결과
    Dim presult2  As String  ' H/L
    Dim pregstate  As String ' "0"
    Dim pflag  As String     ' Flag

    Dim plstupddt  As String ' 저장시간
    Dim pUSERID  As String   ' User id
    Dim pwrkgbn  As String   ' "E"
    Dim pcmt1  As String     ' NULL
    Dim phold1  As String    ' NULL
    Dim sqlDoc  As String
   ' Dim pOTHER   As String
   ' Dim pwdate   As String
    Dim pOrderCode As String
    Dim strACK     As String
    Dim p칼라      As String
    
    On Error GoTo ErrRoutine
    
    With spdResult1
        .Row = strRow
        .Col = 2: pwseq = .text
        .Col = 3: pjno = .text
        .Col = 4: pregno = .text
        .Col = 5: pname = .text
        .Col = 6: pjdate = .text:
        .Col = 7: pemer = .text
        .Col = 8: psex = .text
        .Col = 10: pRack = .text
        .Col = 11: pPos = .text
        .Col = 12: pjgbn = .text
        
        .Col = strCol: pResult1 = .text: p칼라 = .BackColor
        
        presult2 = medGetP(.CellNote, 1, "/")
        pflag = medGetP(.CellNote, 2, "/")
        
        pOrderCode = .CellTag
        
        If Order_Check = True Then
            If pOrderCode = "Micro" Then
                strACK = "001"
            Else
                strACK = f_funGet_CODE_ACK(pOrderCode)
            End If
        Else
            pOrderCode = strIFCD
            strACK = strIFCD
        End If
        
       ' Debug.Print .CellNote & " - " & presult2 & " - " & pflag
        
    End With
    
    pifseq = strIFCD
    pwdate = Format(Now, "YYYYMMDD")
    prerun = ""
    pwrkgbn = "E"
    peqgbn = "AX"
    pregstate = "0"
    pOTHER = Format(Now, "HH:MM")
    pwdate = strWdate
    
    If UCase(Trim(pResult1)) = "X" Then
        pResult1 = ""
    End If

    If pjno <> "" And strACK <> "" Then
    
        If pOrderCode <> "" Then

            sqlDoc = ""
            sqlDoc = sqlDoc + " Select *                              " & vbCrLf
            sqlDoc = sqlDoc + "   From dbo.IFRESULT01                 " & vbCrLf
            sqlDoc = sqlDoc + "  Where 1 = 1                          " & vbCrLf
            sqlDoc = sqlDoc + "    And WDATE  =  '" & pwdate & "'     " & vbCrLf
            sqlDoc = sqlDoc + "    And IFSEQ  =  '" & strACK & "'     " & vbCrLf
            sqlDoc = sqlDoc + "    And WSEQ   =  '" & pwseq & "'      "
             
            adoRS.CursorLocation = adUseClient
            adoRS.Open sqlDoc, AdoCn_SQL
            
              
             If adoRS.RecordCount >= 0 Then
             
    
                sqlDoc = ""
                sqlDoc = sqlDoc & " Update IFRESULT01                        " & vbCrLf
                sqlDoc = sqlDoc & "    Set eqgbn    = '" & peqgbn & "',      " & vbCrLf
                sqlDoc = sqlDoc & "        WSEQ     = '" & pwseq & "',       " & vbCrLf
                sqlDoc = sqlDoc & "        WDATE    = '" & pwdate & "',      " & vbCrLf
                sqlDoc = sqlDoc & "        JDATE    = '" & pjdate & "',      " & vbCrLf
                sqlDoc = sqlDoc & "        JGBN     = '" & pjgbn & "',       " & vbCrLf
                sqlDoc = sqlDoc & "        RACK     = '" & pRack & "',       " & vbCrLf
                sqlDoc = sqlDoc & "        POS      = '" & pPos & "',        " & vbCrLf
                sqlDoc = sqlDoc & "        REGNO    = '" & pregno & "',      " & vbCrLf
                sqlDoc = sqlDoc & "        NAME     = '" & pname & "',       " & vbCrLf
                sqlDoc = sqlDoc & "        SEX      = '" & psex & "',        " & vbCrLf
                sqlDoc = sqlDoc & "        EMER     = '" & pemer & "',       " & vbCrLf
                sqlDoc = sqlDoc & "        RERUN    = '',                    " & vbCrLf
                sqlDoc = sqlDoc & "        OTHER    = '" & pOTHER & "',      " & vbCrLf
                sqlDoc = sqlDoc & "        RESULT1  = '" & pResult1 & "',    " & vbCrLf
                sqlDoc = sqlDoc & "        RESULT2  = '" & presult2 & "',    " & vbCrLf
                sqlDoc = sqlDoc & "        REGSTATE = '0',                        " & vbCrLf
                sqlDoc = sqlDoc & "        lstupddt = getdate(),                  " & vbCrLf
                sqlDoc = sqlDoc & "        wrkgbn   = '" & pwrkgbn & "',          " & vbCrLf
                sqlDoc = sqlDoc & "        userid   = '" & CurrUser.CuUserID & "' " & vbCrLf
                sqlDoc = sqlDoc + "  Where WDATE    =  '" & pwdate & "'           " & vbCrLf
                sqlDoc = sqlDoc + "    And IFSEQ    =  '" & strACK & "'           " & vbCrLf
                sqlDoc = sqlDoc + "    And WSEQ     =  '" & pwseq & "'            " & vbCrLf
                AdoCn_SQL.Execute sqlDoc, sqlRet
                Result_Upload_MSSQL = True
                
                If sqlRet <> 0 Then
                    Call SaveLog("::::: UPDATE ; " & Order_Check & vbCrLf & sqlDoc)
                End If

                
                If sqlRet = 0 Then

                
                    sqlDoc = ""
                    sqlDoc = sqlDoc & "Insert Into IFRESULT01(jno,ifseq,rack,pos,wdate,wseq,eqgbn,jdate,jgbn,regno,name,sex,emer,rerun,other,result1,result2,regstate,flag,lstupddt,userid,wrkgbn,cmt1,hold1) " & vbCrLf
                    sqlDoc = sqlDoc & "     Values('" & pjno & "', "
                    sqlDoc = sqlDoc & "            '" & strACK & "', "
                    sqlDoc = sqlDoc & "            '" & pRack & "', "
                    sqlDoc = sqlDoc & "            '" & pPos & "', "
                    sqlDoc = sqlDoc & "            '" & pwdate & "', "
                    sqlDoc = sqlDoc & "            '" & pwseq & "', "
                    sqlDoc = sqlDoc & "            '" & peqgbn & "', "
                    sqlDoc = sqlDoc & "            '" & pjdate & "', "
                    sqlDoc = sqlDoc & "            '" & pjgbn & "', "
                    sqlDoc = sqlDoc & "            '" & pregno & "', "
                    sqlDoc = sqlDoc & "            '" & pname & "', "
                    sqlDoc = sqlDoc & "            '" & psex & "', "
                    sqlDoc = sqlDoc & "            '" & pemer & "', "
                    sqlDoc = sqlDoc & "            'N', "
                    sqlDoc = sqlDoc & "            '" & pOTHER & "', "
                    sqlDoc = sqlDoc & "            '" & pResult1 & "', "
                    sqlDoc = sqlDoc & "            '" & presult2 & "', "
                    sqlDoc = sqlDoc & "            '0', "
                    sqlDoc = sqlDoc & "            '" & pflag & "', "
                    sqlDoc = sqlDoc & "            getdate(), "
                    sqlDoc = sqlDoc & "            '" & CurrUser.CuUserID & "', "
                    sqlDoc = sqlDoc & "            '" & pwrkgbn & "', "
                    sqlDoc = sqlDoc & "            '', "
                    sqlDoc = sqlDoc & "            '') " & vbCrLf
    
                    AdoCn_SQL.Execute sqlDoc
                    Result_Upload_MSSQL = True
                    Call SaveLog("::::: INSERT ; " & Order_Check & vbCrLf & sqlDoc)
                    Call SaveLog("결과         ; " & pjno & " | " & strACK & " | " & pResult1 & vbCrLf)
                    Call SaveLog("---------------------------------------------------------------------" & vbCrLf)
                    
                    
                    Debug.Print pjno & " | " & strACK & " | " & pResult1 & vbCrLf
                    
                End If
            End If

        End If

    End If

Exit Function

ErrRoutine:
    Result_Upload_MSSQL = False

End Function

Public Function CheckSum_ECi_Tx(ByVal strPrmValue As String)

    Dim I                   As Integer
    Dim intValueLength      As Integer
    Dim intCheck            As Integer
    Dim strCheck            As String
    
    intCheck = 0
    
    intValueLength = LenA(strPrmValue)
    
    For I = 1 To intValueLength
        intCheck = intCheck + Asc(Mid(strPrmValue, I, 1))
    Next
    
    strCheck = Hex(intCheck)
    
    If Len(strCheck) = 1 Then
        CheckSum_ECi_Tx = "0" & strCheck
    Else
        CheckSum_ECi_Tx = Right(strCheck, 2)
    End If

End Function

Public Function LenA(strPrmString As String) As Integer

    Dim I                   As Integer
    Dim intStrLen           As Integer
    Dim intAnsiStrLen       As Integer
    Dim strTemp             As String
    
    intStrLen = Len(strPrmString)
    For I = 1 To intStrLen
        strTemp = Mid(strPrmString, I, 1)
        
        Select Case AscW(strTemp)
        Case 0 To 255
            intAnsiStrLen = intAnsiStrLen + 1
        
        Case Else
            intAnsiStrLen = intAnsiStrLen + 2
        
        End Select
    Next
    
    LenA = intAnsiStrLen

End Function

Private Sub cmdENQ_Click()
    
    Call COM_OUTPUT(charCOM_Convert(COM_ENQ))

End Sub

Private Sub cmdRstQuery_Click()

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String, intRet   As Integer
    
    Dim strSpcno    As String
    Dim intRow      As Integer, intCol  As Integer
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
    
    Dim itemX       As ListItem

    intRow = 0
    
    With spdResult2
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    sqlDoc = "Select SPCNO, TESTCD, EQUIPCD, TRANSDT, RSTVAL, REFVAL, TRANSDT, EQPNUM, NAME, PNO,SEQ" & _
             "  from INTERFACE003" & _
             " where TRANSDT between '" & mskRstDate.text & "' and '" & mskRstDate1.text & "'" & _
             "   and EQUIPCD = '" & INS_CODE & "'"
    If cboRstgbn(1).ListIndex = 0 Then
        sqlDoc = sqlDoc & "   and SERVERGBN = ''"
    ElseIf cboRstgbn(1).ListIndex = 1 Then
        sqlDoc = sqlDoc & "   and SERVERGBN = 'Y'"
    End If
    sqlDoc = sqlDoc & " Order by SPCNO, TRANSTM"
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        With spdResult2
        If strSpcno <> Trim$(adoRS(0) & "") + Trim$(adoRS(6) & "") Then
                intRow = intRow + 1
                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                .SetText 1, intRow, "1"
                .SetText 2, intRow, Trim$(adoRS(3) & "")
                .SetText 3, intRow, Trim$(adoRS(9) & "")
                .SetText 4, intRow, Trim$(adoRS(8) & "")
                .SetText 8, intRow, Trim$(adoRS(0) & "")
                .SetText 9, intRow, Trim$(adoRS(10) & "")
               
            End If
                strSpcno = Trim$(adoRS(0) & "") + Trim$(adoRS(6) & "")
                Set itemX = lvwCuData.FindItem(Trim$(adoRS(7) & ""), lvwTag, , lvwWhole)
                If Not itemX Is Nothing Then
                    intCol = itemX.Index + 11
                    .SetText intCol, intRow, Trim$(adoRS(4)) & ""
                    .Col = intCol:  .Row = intRow:  .ForeColor = vbBlack ' IIf(Trim$(adoRS(5) & "") <> "", vbRed, vbBlack)
                End If
        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub cmdSel_Click(Index As Integer)

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    With spdResult1
        For intRow = 1 To .maxrows
            .GetText 2, intRow, varTmp
            If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 0, "1", "")
        Next
    End With
    
End Sub

Private Sub cmdStartNo_Click()
Dim sNo As String, sCnt As Integer, sAdd As Integer

AgainInput:
    sNo = InputBox("시작 번호를 입력하세요 !")
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
        If Not IsNumeric(sNo) Then
            MsgBox "숫자만 입력하세요.!", vbCritical
            GoTo AgainInput
        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .maxrows
                .Row = sCnt
                .Col = 0:       .text = Trim(sAdd + Val(sNo))
                sAdd = sAdd + 1
            Next sCnt
        
            .StartingRowNumber = Val(sNo)
        End With
    End If

End Sub

Private Function f_subSet_TestList(ByVal strYY As String, ByVal strMM As String, ByVal strDD As String, ByVal strChart As String)
    Dim sqlRet      As Integer
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
    
    Set AdoRs_SQL = New ADODB.Recordset

    gSql = "       Select * From TB_검사항목 "
    gSql = gSql & " Where 진료년 = '" & strYY & "'"
    gSql = gSql & "   and 진료월 = '" & strMM & "'"
    gSql = gSql & "   and 진료일 = '" & strDD & "'"
    gSql = gSql & "   and 챠트번호 = '" & strChart & "'"
    
    AdoRs_SQL.Open gSql, AdoCn_SQL, adOpenStatic, adLockReadOnly
    
    If AdoRs_SQL.RecordCount = 0 Then
        Set f_subSet_TestList = Nothing
    Else
        Set f_subSet_TestList = AdoRs_SQL
    End If

    Set AdoRs_SQL = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    Call ErrMsgProc(CallForm)

    
End Function

Private Sub Protocol_Call()
    
    Dim wkDat   As String
    Dim ix1 As Integer
    Dim I   As Integer
    
    Call SaveLog(wkBuf)
    
    Debug.Print wkBuf

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
        

        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        m_iPhase = 2
                        RstEnd = "Y"
                        bEndChk = True: bSTXChk = False

                        comEQP.Output = ACK

                    Case Else
                        'm_iPhase = 1
                End Select

            Case 2
                Select Case Asc(wkDat)
                    Case 2      'STX
                        If bEndChk = True Then
                            RcvBuffer = ""
                        Else
                            bSTXChk = True
                        End If
                        bEndChk = True

                    Case 10     '<LF>
                        If bEndChk = True Then

                            Call EditRcvData

                            RcvBuffer = ""
                        End If
                        comEQP.Output = ACK

                    Case 13     'CR
                        If bEndChk = True Then

                            Call EditRcvData

                            RcvBuffer = ""
                        End If
                    Case 3      'ETX
                        comEQP.Output = ACK
                    Case 4      'EOT
                        RcvBuffer = ""
                        m_iPhase = 1

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = True
                        comEQP.Output = ACK   'Send ACK

                    Case 21     'NAK
                        Call EditRcvData
                        
                        m_iSendPhase = 1
                        m_iFrameN = 1

                        comEQP.Output = ENQ   'Send ENQ

                    Case 23     ' ETB
                        bEndChk = False

                    Case Else
                        If bEndChk = True Then
                            If bSTXChk = True Then
                                bSTXChk = False
                            Else
                                RcvBuffer = RcvBuffer & wkDat
                            End If
                        End If

                End Select
            
            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK

                    
                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        comEQP.Output = ACK
                        m_iPhase = 2
                    
                    Case 21     'NAK
                        comEQP.Output = ENQ   'ENQ
                        m_iPhase = 3
                        
                    Case 4      'EOT
                        m_iPhase = 1
                        
                End Select

        End Select
    Next ix1
    
End Sub

Private Sub cmd닫기_Click()
    If uplUrineMicro.Visible = True Then
        uplUrineMicro.Visible = False
    Else
        uplUrineMicro.Visible = True
    End If
End Sub

Private Sub comEQP_OnComm()
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim Arr()       As Byte
    Dim strdata     As String
    Dim brStr As String
    Dim sStxCheck As Integer, sEtxCheck As Integer, sCrcheck As Integer
    Dim com_sTemp As String
    Dim ii As Integer, jj As Integer
    Dim MHead  As String, pInfo As String
    Dim PatientID As String
    
    Dim Orderoutput As String
    Dim OutPutData  As String
    Dim Rev As Long
    Dim Test_Cd() As String, strPid()    As String, strPnm() As String
    Dim sRow As Integer
    Dim oPatNo As String
    Dim oRackNo As String
    Dim oPosNo As String
    Dim oIdNo As String
    
    Dim adoRS As ADODB.Recordset
    Dim sqlDoc As String
    Dim itemX As ListItem
    Dim strEqpCd1 As String
    
    Dim varTmp  As Variant
    Dim intCol  As Integer
    Dim strLevel() As String

    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim I           As Long
    
    Select Case comEQP.CommEvent
        Case comEvReceive
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If
            
            wkBuf = comEQP.Input
            
''''''Update IFRESULT01
''''''SET eqgbn = 'UF', WSEQ = '0002',  Wdate = '20131012',         JDATE = '',         JGBN = '',         RACK = '',         POS = '',         REGNO = '',         NAME = '',         SEX = '',         EMER = '',         RERUN = '',         OTHER = '',         RESULT1 = 'many',         RESULT2 = '',         FLAG = '',         REGSTATE = '0',         lstupddt = getdate(),         wrkgbn = 'E',         userid = '513201'
''''''where jno = 'QC14'    AND IFSEQ = '016' )
''''''
''''''INSERT INTO IFRESULT01 (eqgbn, WDATE, WSEQ, IFSEQ, JDATE, JGBN, JNO, RACK, POS, REGNO, [NAME], SEX, EMER, RERUN, OTHER, RESULT1, RESULT2, REGSTATE, FLAG, lstupddt, wrkgbn, userid)
''''''VALUES ( 'UF', '20131012', '0002', '016', '', '', 'QC14', '', '', '', '', '', '', '', '', 'many', '', '0', '',  getdate(), 'E', '513201') )
''''''
''''''
             
            Call Protocol_Call
        
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
        
End Sub


Private Function Work_seq(ByVal brDate As String, ByVal brType As String) As Integer
Dim adoRS  As New ADODB.Recordset
Dim sqlDoc As String
Dim pSeq   As Integer
    
    If brType = "CONTROL" Then

        sqlDoc = "Select max(wseq) as MaxNo      " & _
                 "  From dbo.IFRESULT01          " & _
                 " Where WDATE = " & Format(brDate, "########") & " "
        
        adoRS.CursorLocation = adUseClient
        adoRS.Open sqlDoc, AdoCn_SQL
        adoRS.MoveFirst
        
        If Len(adoRS.Fields("MaxNo") & "") = 0 Then
             pSeq = 0
        Else
             pSeq = adoRS.Fields("MaxNo")
        End If

        
        If pSeq = 0 Then
            Work_seq = 1

        Else
            Work_seq = adoRS.Fields("MaxNo") + 1
        End If
    Else
        sqlDoc = "Select max(wseq) as MaxNo      " & _
                 "  From dbo.IFRESULT01          " & _
                 " Where WDATE = " & Format(brDate, "########") & " "
        
        adoRS.CursorLocation = adUseClient
        adoRS.Open sqlDoc, AdoCn_SQL
        adoRS.MoveFirst
        Work_seq = adoRS.Fields("MaxNo") + 1
    
    End If
    
   
    Set adoRS = Nothing

End Function
'
'1H|\^&|||2338004|||||||P|6146010-10-05
'P|1
'O|1|527714102405|890^50032^5^1^SAMPLE||R||||||X|||20140613232233
'R|1|^^^1|1: SG|||||
'C|1|I|*|I
'R|2|^^^2|2: pH|||||
'C|2|I||I
'R|3|^^^3|3: LEU(WBC)|||||
'C|3|I|#|I
'R|4|^^^4|4: NIT|||||
'C|4|I||I
'R|5|^^^5|5: PRO|||||
'C|5|I80
'2|#|I
'R|6|^^^6|6: GLU|mg/dL||||
'C|6|I|*^#|I
'R|7|^^^7|7: KET|mg/dL||||
'C|7|I|*^#|I
'R|8|^^^8|8: UBG|||||
'C|8|I||I
'R|9|^^^9|9: BIL|||||
'C|9|I||I
'R|10|^^^10|10: ERY(RBC)|||||
'C|10|I|#|I
'R|11|^^^11|11: COL|||||
'C|11|I||I
'R|12|^^^12|12: CLA|||||
'C|12|I||I
'L|1|
'03
'

Private Sub EditRcvData()

    Dim sTemp         As String
    Dim Channel_No    As String
    Dim Patiant_No    As String
    Dim pGrid_Point   As Integer
    Dim Max_Arary_Cnt As Integer
    Dim sDeCnt        As Integer
    Dim pDoCount      As Integer
    Dim Loop_count    As Integer
    Dim sRtn          As Integer
    Dim sChannel      As String
    Dim sRstText      As String
    Dim sRstValue     As Single
    Dim sUnit         As String
    Dim intIdx        As Integer
    Dim strEqpCd      As String
    Dim itemX         As ListItem
    Dim sCol          As Integer
    Dim varTmp        As Variant

    Dim strTmp        As String
    Dim strBarno      As String
    Dim strDATE       As String
    Dim strTime       As String
    Dim strDate1      As String

    Dim intCol        As Integer
    Dim strRstval     As String
    Dim strRefVal     As String
    Dim pRstval       As String ' 검사결과
    Dim pDBTestcd     As String ' 디비검사항목코드
    Dim sqlDoc        As String
    Dim iDataChk      As Integer
    Dim BooData       As Boolean

    Dim valEqpcd      As Variant
    Dim intCnt        As Integer
    Dim pJubSu        As Boolean
    
    Dim strRcvBuf    As String   '수신한 Data
    Dim RecType As String
'
    On Error Resume Next
    Dim ii As Integer
    Dim Ord_Cnt As Integer

    Debug.Print RcvBuffer


    ii = InStr(1, RcvBuffer, "|")
    If ii <> 0 Then
        RecType = Mid$(RcvBuffer, ii - 1, 1)
    Else
        Exit Sub
    End If
    
    Select Case RecType

        Case "H"
        
            sState = ""
            With Urisys2400
                .BarCode_NO = ""
                .SEQ = ""
                .Position = ""
                .RackNo = ""
                .Type = ""
                .Order_Str = ""
            End With
            
            pNIT_Check = False
            pWBC_Check = False
            pRBC_Check = False
 
        Case "P"

        Case "C"

        Case "Q"

        Case "O"
            BooData = False
            
            Dim sSeq   As String
            Dim sRack  As String
            Dim sPos   As String
            Dim sType   As String
            Dim pSeq    As Integer
            
            
            fUrisys2400 = Split(RcvBuffer, "|")
    
            sPatiant_No = fUrisys2400(2)

            fUrisys2400_SUB = Split(fUrisys2400(3), "^")
            
            With Urisys2400
                .BarCode_NO = sPatiant_No
                .SEQ = fUrisys2400_SUB(0)
                .RackNo = fUrisys2400_SUB(1)
                .Position = fUrisys2400_SUB(2)
                .Type = fUrisys2400_SUB(4)
            End With
            
            If Urisys2400.Type = "CONTROL" Then     '2008/1/8 yk
                Urisys2400.BarCode_NO = fUrisys2400_SUB(0) & "(" & Urisys2400.RackNo & "-" & Urisys2400.Position & ")"
                Urisys2400.Type = "CONTROL"
                Urisys2400.SEQ = "1"
            Else
                Urisys2400.Type = "SAMPLE"
            End If
            
            pSeq = Work_seq(Format(DptDate.Value, "YYYYMMDD"), Urisys2400.Type)
    
            For iDataChk = 1 To spdResult1.maxrows
                spdResult1.GetText 3, iDataChk, varTmp:   strBarno = Trim$(varTmp)
                If sPatiant_No = strBarno Then
                    BooData = False
                End If
            Next
    
            If BooData = False Then
                
                If Len(sPatiant_No) <> 0 Then
    
                    Set mAdoRs = f_subSet_WorkList_Barcode(Urisys2400.BarCode_NO, INS_CODE)
                Else
                    RecordChk = False
                End If
    
                If RecordChk = False Then
                    List1.AddItem ("샘플번호 " & sPatiant_No & " 는 등록되지 않은 검체입니다. 확인하세요.")
    
'                    pGrid_Point = SeqSearch(spdResult1, Urisys2400.BarCode_NO, 3)
'
'                    If pGrid_Point = 0 Then
'                        pGrid_Point = SeqNullSearch(spdResult1, Urisys2400.BarCode_NO, 3)
'                        If pGrid_Point = 0 Then spdResult1.maxrows = spdResult1.maxrows + 1: pGrid_Point = spdResult1.maxrows
'                    End If
                    
                    spdResult1.Row = 1
                    spdResult1.Col = 13
                    
                    If spdResult1.text = "" Then
                        pGrid_Point = spdResult1.maxrows
                    Else
                    
                        spdResult1.maxrows = spdResult1.maxrows + 1
                        pGrid_Point = spdResult1.maxrows
                    End If

    
                    With spdResult1
                        .SetText 1, pGrid_Point, "0"
                        .SetText 2, pGrid_Point, Format(pSeq, "0000")                   '-- 검체번호
                        .SetText 3, pGrid_Point, Format(pSeq, "0000") & "(" & Urisys2400.SEQ & "-" & Urisys2400.RackNo & "-" & Urisys2400.Position & ")"
                        If Urisys2400.Type = "CONTROL" Then
                            .SetText 5, pGrid_Point, "QC" & Urisys2400.SEQ
                            .SetText 13, pGrid_Point, "검사"
                        Else
                            .SetText 5, pGrid_Point, "오더확인"
                            .SetText 13, pGrid_Point, "검사"
                        End If
                        .SetText 7, pGrid_Point, Urisys2400.BarCode_NO
                        .SetText 8, pGrid_Point, Urisys2400.BarCode_NO
                        .SetText 9, pGrid_Point, Urisys2400.SEQ
                        .SetText 10, pGrid_Point, Urisys2400.RackNo
                        .SetText 11, pGrid_Point, Urisys2400.Position
                        .SetText 12, pGrid_Point, Urisys2400.Position

                        .Row = pGrid_Point
                        If OldRow > 0 Then
                            .Row = OldRow
                        End If
                        .Action = ActionActiveCell
                    
                    End With
                Else
                    With spdResult1
                    Do Until mAdoRs.EOF
                        intIdx = 0
                        If strBarno <> mAdoRs.Fields("REQNO") Then
                            pGrid_Point = SeqSearch(spdResult1, mAdoRs.Fields("REQNO"), 7)
                            
                            If pGrid_Point >= 1 Then
                            
                                Dim pMessage As String
                            
                                pMessage = ""
                                .Row = pGrid_Point
                                .Col = 2: pMessage = pMessage & .text
                                .Col = 3: pMessage = pMessage & " | " & .text
                                .Col = 5: pMessage = pMessage & " | " & .text
                                txtHelp.text = txtHelp.text & "중복 ▷" & pMessage & vbCrLf
                            
                            End If
                            
'
'                            If pGrid_Point = 0 Then
'                                pGrid_Point = SeqNullSearch(spdResult1, mAdoRs.Fields("REQNO"), 7)
'                                If pGrid_Point = 0 Then .maxrows = .maxrows + 1: pGrid_Point = .maxrows
'                            End If

                            spdResult1.Row = 1
                            spdResult1.Col = 13
                            
                            If spdResult1.text = "" Then
                                pGrid_Point = spdResult1.maxrows
                            Else
                                spdResult1.maxrows = spdResult1.maxrows + 1
                                pGrid_Point = spdResult1.maxrows
                            End If
    
                            .SetText 1, pGrid_Point, "0"
   
                            .SetText 2, pGrid_Point, Format(pSeq, "0000")                   '-- 검체번호
                            .SetText 3, pGrid_Point, Urisys2400.BarCode_NO
                            .SetText 4, pGrid_Point, Trim(mAdoRs("CSTNM")) & ""
                            .SetText 5, pGrid_Point, Trim(mAdoRs("PATNM")) & ""
                            
                            .SetText 6, pGrid_Point, Trim(mAdoRs("REQDTE"))
                            .SetText 7, pGrid_Point, Trim(mAdoRs("REQNO"))
                            .SetText 8, pGrid_Point, Trim(mAdoRs("IDNO"))
                            .SetText 12, pGrid_Point, Trim(mAdoRs("WRKDTE")) & "/" & Trim(mAdoRs("WORKNO")) & "/" & Trim(mAdoRs("LABEMP"))
                            .SetText 13, pGrid_Point, "검사"
                            .Row = pGrid_Point
                            .Action = ActionActiveCell
                            '.Action = ActionGotoCell
                            '.SearchRow
                        End If
    
                        .SetText 9, pGrid_Point, Urisys2400.SEQ
                        .SetText 10, pGrid_Point, Urisys2400.RackNo
                        .SetText 11, pGrid_Point, Urisys2400.Position

    
                        If Len(Trim(mAdoRs.Fields("ITEMCD"))) > 0 Then
                            strEqpCd = f_funGet_CODE_SUB(Trim(mAdoRs.Fields("ITEMCD")))
'                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
'                            If Not itemX Is Nothing Then
                            If strEqpCd > 0 Then
                                spdResult1.SetText 1, pGrid_Point, "0"
                                spdResult1.Col = strEqpCd
                                spdResult1.Row = pGrid_Point
                                spdResult1.BackColor = &HC6FEFF
                                spdResult1.CellTag = Trim(mAdoRs.Fields("ITEMCD"))
                            End If
                            'End If
                            strBarno = mAdoRs.Fields("REQNO")
                        End If
    
                        mAdoRs.MoveNext
                    Loop
                    End With

                End If
            End If
        Case "R"
            Dim mType      As String
            Dim pHighLow   As String
            Dim pflag      As String

            Dim p오더체크  As String
            Dim strSEQ     As String
            Dim sqlRet     As Integer

            fUrisys2400 = Split(RcvBuffer, "|")

            mType = Right(fUrisys2400(8), 1)
            fUrisys2400(2) = medGetP(fUrisys2400(2), 4, "^")

            Channel_No = fUrisys2400(2) ' channel

            intRow = 0
            pGrid_Point = 0

            With spdResult1
                sCol = 13

                pGrid_Point = SeqSearch(spdResult1, "검사", sCol)

                If pGrid_Point = 0 Then
                    pGrid_Point = SeqSearch(spdResult1, Urisys2400.SEQ, 9)
                End If

                If pGrid_Point = 0 Then Exit Sub
                
                .GetText 2, pGrid_Point, varTmp:    strSEQ = Trim$(varTmp)
                .GetText 3, pGrid_Point, varTmp:    strBarno = Trim$(varTmp)
                .GetText 5, pGrid_Point, varTmp:    pname = Trim$(varTmp)
                .GetText 6, pGrid_Point, varTmp:    strDATE = Trim$(varTmp)
                .GetText 8, pGrid_Point, varTmp:    pNo = Trim$(varTmp)

                strDate1 = Format$(Now, "YYYYMMDD"): strTime = Format$(Now, "MMSS")

                If pGrid_Point > 0 Then
                
                    .Row = pGrid_Point
                    .Action = ActionActiveCell
                    
                    For intCol = 13 To .MaxCols
                        strRstval = ""
                        strEqpCd = ""
                        .GetText intCol, 0, varTmp
                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                        If Not itemX Is Nothing Then
                            For intIdx = 13 To .MaxCols
                                If Len(Trim(fUrisys2400(3))) > 0 Or fUrisys2400(8) = "X" Then
                                    If Channel_No = itemX.tag Then

                                        strRefVal = ""
                                        
                                        If fUrisys2400(3) = "" Then
                                            strRstval = Trim(fUrisys2400(8))
                                        Else
                                            strRstval = Trim(fUrisys2400(3))
                                        End If
                                        
                                        pflag = Trim(fUrisys2400(3))

                                        strRstval = Trim(Replace(strRstval, "TR", ""))

                                        Select Case Channel_No
                                            '1: SG
                                            Case 1: strRstval = strRstval: pHighLow = "": pflag = ""
                                            '2: pH
                                            Case 2: strRstval = Format(strRstval, "###0.0"): pHighLow = "": pflag = ""
                                            '3: LEU(WBC)
                                            Case 3
                                                Select Case strRstval
                                                    Case "NEG":       strRstval = "Negative":   pHighLow = "":  pflag = ""
                                                    Case "1+ 25":     strRstval = "1 Positive": pHighLow = "":  pflag = "*": pWBC_Check = True
                                                    Case "2+100":     strRstval = "2 Positive": pHighLow = "":  pflag = "*": pWBC_Check = True
                                                    Case "3+500":     strRstval = "3 Positive": pHighLow = "":  pflag = "*": pWBC_Check = True
                                                    Case Else:        strRstval = strRstval:    pHighLow = "":  pflag = ""
                                                End Select
                                            '4: NIT
                                            Case 4
                                                Select Case strRstval
                                                    Case "NEG":       strRstval = "Negative":   pHighLow = "":  pflag = ""
                                                    Case "POS":       strRstval = "Positive":   pHighLow = "H": pflag = "*": pNIT_Check = True
                                                    Case Else:        strRstval = strRstval:    pHighLow = "":  pflag = ""
                                                End Select
                                            '5: PRO
                                            Case 5
                                                Select Case strRstval
                                                    Case "NEG":       strRstval = "Negative":   pHighLow = "":  pflag = ""
                                                    Case "25":        strRstval = "Trace":      pHighLow = "H": pflag = ""
                                                    Case "1+ 75":     strRstval = "1 Positive": pHighLow = "H": pflag = "*"
                                                    Case "2+150":     strRstval = "2 Positive": pHighLow = "H": pflag = "*"
                                                    Case "3+500":     strRstval = "3 Positive": pHighLow = "H": pflag = "*"
                                                    Case Else:        strRstval = strRstval:    pHighLow = "":  pflag = ""
                                                End Select
                                            '6: GLU
                                            Case 6
                                                Select Case strRstval
                                                    Case "NORM":      strRstval = "Negative":   pHighLow = "":  pflag = ""
                                                    Case "50":        strRstval = "Trace":      pHighLow = "H": pflag = ""
                                                    Case "1+100":     strRstval = "1 Positive": pHighLow = "H": pflag = "*"
                                                    Case "2+300":     strRstval = "2 Positive": pHighLow = "H": pflag = "*"
                                                    Case "3+999":     strRstval = "3 Positive": pHighLow = "H": pflag = "*"
                                                    Case Else:        strRstval = strRstval:    pHighLow = "":  pflag = ""
                                                End Select
                                            '7: KET
                                            Case 7
                                                Select Case strRstval
                                                    Case "NEG":        strRstval = "Negative":   pHighLow = "":  pflag = ""
                                                    Case "5":          strRstval = "Trace":      pHighLow = "H": pflag = ""
                                                    Case "1+ 15":      strRstval = "1 Positive": pHighLow = "H": pflag = "*"
                                                    Case "2+ 50":      strRstval = "2 Positive": pHighLow = "H": pflag = "*"
                                                    Case "3+150":      strRstval = "3 Positive": pHighLow = "H": pflag = "*"
                                                    Case Else:         strRstval = strRstval:    pHighLow = "":  pflag = ""
                                                End Select
                                            '8: UBG
                                            Case 8
                                                Select Case strRstval
                                                    Case "NORM":       strRstval = "Negative":   pHighLow = "":  pflag = ""
                                                    Case "1":          strRstval = "Trace":      pHighLow = "H": pflag = ""
                                                    Case "1+ 4":       strRstval = "1 Positive": pHighLow = "H": pflag = "*"
                                                    Case "2+ 8":       strRstval = "2 Positive": pHighLow = "H": pflag = "*"
                                                    Case "3+ 12":      strRstval = "3 Positive": pHighLow = "H": pflag = "*"
                                                    Case Else:         strRstval = strRstval:    pHighLow = "":  pflag = ""
                                                End Select
                                            '9: BIL
                                            Case 9
                                                Select Case strRstval
                                                    Case "NEG":        strRstval = "Negative":   pHighLow = "":  pflag = ""
                                                    Case "1+ 1":       strRstval = "1 Positive": pHighLow = "H": pflag = "*"
                                                    Case "2+ 3":       strRstval = "2 Positive": pHighLow = "H": pflag = "*"
                                                    Case "3+ 6":       strRstval = "3 Positive": pHighLow = "H": pflag = "*"
                                                    Case Else:         strRstval = strRstval:    pHighLow = "":  pflag = ""
                                                End Select
                                            '10: ERY(RBC)
                                            Case 10
                                                Select Case strRstval
                                                    Case "NEG":        strRstval = "Negative":    pHighLow = "":  pflag = ""
                                                    Case "10":         strRstval = "Trace":       pHighLow = "":  pflag = ""
                                                    Case "1+ 25":      strRstval = "1 Positive":  pHighLow = "":  pflag = "*": pRBC_Check = True
                                                    Case "2+ 50":      strRstval = "2 Positive":  pHighLow = "":  pflag = "*": pRBC_Check = True
                                                    Case "3+150":      strRstval = "3 Positive":  pHighLow = "":  pflag = "*": pRBC_Check = True
                                                    Case "4+250":      strRstval = "4 Positive":  pHighLow = "":  pflag = "*": pRBC_Check = True
                                                    Case Else:         strRstval = strRstval:     pHighLow = "":  pflag = ""
                                                End Select
                                            '11: COL
                                            Case 11
                                                Select Case strRstval
                                                    Case "P.YEL":     strRstval = "Pale yellow":    pHighLow = "": pflag = ""
                                                    Case "YELLO":     strRstval = "Yellow":         pHighLow = "": pflag = ""
                                                    Case "AMBER":     strRstval = "Amber":          pHighLow = "": pflag = ""
                                                    Case "BROWN":     strRstval = "Brown":          pHighLow = "": pflag = ""
                                                    Case "ORANG":     strRstval = "Orange":         pHighLow = "": pflag = ""
                                                    Case "RED":       strRstval = "Red":            pHighLow = "": pflag = ""
                                                    Case "GREEN":     strRstval = "Green":          pHighLow = "": pflag = ""
                                                    Case "OTHER":     strRstval = "Other":          pHighLow = "": pflag = ""

                                                    Case Else:      strRstval = strRstval:       pHighLow = "":  pflag = ""
                                                End Select
                                            '12: CLA
                                            'P.YEL','YELLO','AMBER','BROWN','ORANGE','RED','GREEN','OTHER'
                                            Case 12:
                                                strRstval = strRstval: pHighLow = "": pflag = ""
                                        End Select

                                        .Col = intCol: .Row = pGrid_Point

                                        .TypeHAlign = TypeVAlignCenter
                                        .TypeVAlign = TypeVAlignCenter

                                        .text = strRstval
                                        .ForeColor = vbBlack
                                        .CellNote = pHighLow & "/" & pflag

                                        Select Case Trim(pflag)
                                            Case "*": .ForeColor = vbRed
                                           ' Case Else: .ForeColor = vbBlack
                                        End Select
                                        
                                        .SetText 13, pGrid_Point, "검사"
                                         
                                        '
                                        ' 오더 체크 / Comment 처리
                                        '
                                        .Col = 14
                                        p오더체크 = .BackColor
                                        .TypeHAlign = TypeHAlignLeft
                                        .TypeVAlign = TypeVAlignCenter
                                        
                                        If .BackColor = &HC6FEFF Then
                                            If pNIT_Check = True Then
                                                .text = "WBC: 0~3 " & vbCrLf & "RBC: 0~3 " & vbCrLf & "E.P cell: 0~3 " & vbCrLf & "Others: Bacteria are seen."
                                            Else
                                                .text = "WBC: 0~3 " & vbCrLf & "RBC: 0~3 " & vbCrLf & "E.P cell: 0~3 " & vbCrLf & "Others: None"
                                                
                                                ' UF1000 order check
                                                If pWBC_Check = True Or pRBC_Check = True Then
                                                    
                                                    Dim UF_SEQ As String
                                                    Dim UF_BARCODE As String
                                                    Dim UF_NAME As String
                                                    Dim UF_Grid_Point As String
                                                    .Col = 2: UF_SEQ = .text
                                                    .Col = 3: UF_BARCODE = .text
                                                    .Col = 5: UF_NAME = .text
                                                    
                                                    UF_Grid_Point = Barcode_Search(spdMicro, UF_BARCODE, 2, 0)
                                                    
                                                    If UF_Grid_Point = 0 Then
                                                        With spdMicro
                                                            .maxrows = .maxrows + 1
                                                            .Row = .maxrows
                                                            .Col = 1: .text = UF_SEQ
                                                            .Col = 2: .text = UF_BARCODE
                                                            .Col = 3: .text = UF_NAME
                                                        End With
                                                    End If
                                                    
                                                    
                                                
                                                End If
                                                
                                            End If
                                        Else
                                            .text = "오더없음"
                                        End If

                                        If pGrid_Point <> 0 Then
                                        
                                            sqlDoc = "Update INTERFACE003" & _
                                                     "   set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRstval & "' , SERVERGBN = 'Y'" & _
                                                     " where SPCNO   = '" & strBarno & "'" & _
                                                     "   and EQPNUM  = '" & itemX.tag & "'" & _
                                                     "   and TESTCD  = '" & itemX.text & "'" & _
                                                     "   and TRANSDT = '" & strDATE & "'" & _
                                                     "   and TRANSTM = '" & strTime & "'"
                                            AdoCn_Jet.Execute sqlDoc, sqlRet
                                            
                                            If sqlRet = 0 Then
                                               sqlDoc = "insert into INTERFACE003(" & _
                                                        "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO,Seq)" & _
                                                        "    values( '" & strBarno & "', '" & itemX.text & "', '" & itemX.tag & "'," & _
                                                        "            '" & strDATE & "', '" & strTime & "'," & _
                                                        "            '" & strRstval & "', '" & strRefVal & "'," & _
                                                        "            '" & INS_CODE & "', '', '" & pname & "', '" & pNo & "', '" & strSEQ & "')"
                                               AdoCn_Jet.Execute sqlDoc
                                            End If

                                        End If
                                        .Col = 1:  .Value = 1

                                    End If
                                    Exit For
                                End If
                            Next intIdx
                        End If
                        Set itemX = Nothing
                    Next
                End If
            End With
        Case "L"
            ' 저장..
            '
            If OldRow > 0 Then
                spdResult1.Row = OldRow
            End If
    
            spdResult1.Action = ActionActiveCell

            If chkAuto.Value = "1" Then
                Call cmdAppend_Click(0)
            End If
            
            pNIT_Check = False
            pWBC_Check = False
            pRBC_Check = False
            
    End Select

    Exit Sub

errDefine:

    Call ErrMsgProc(CallForm)

End Sub


'
'   환자 Order 전송
'
Private Sub SendOrder()
    On Error GoTo Err_Rtn

    Dim sSendBuff   As String
    Dim iCnt    As Integer
    Dim ChkSum  As String
    Dim sStat   As String
    
    m_iFrameN = 1

    Select Case m_iSendPhase
        Case 1
            sSendBuff = m_iFrameN & "H|\^&|||HOST^2|||||E7600^1|TSDWN^REPLY|P|1" & vbCr
            sSendBuff = sSendBuff & "P|1" & vbCr
            
            sSendBuff = sSendBuff & "O|1|" & left(Trim(Urisys2400.BarCode_NO) & Space(22), 22) & Chr(124)
            sSendBuff = sSendBuff & Urisys2400.SEQ & "^" & Trim(Urisys2400.RackNo) & "^" & Trim(Urisys2400.RackNo) & "^^S1^SC" & Chr(124)
            
            If Len(Urisys2400.Order_Str) <> 0 Then
                sSendBuff = sSendBuff & Urisys2400.Order_Str
                sSendBuff = left(sSendBuff, Len(sSendBuff) - 1)      '"\" Cutting
            End If

            If left(Urisys2400.RackNo, 1) = "4" Then
                sStat = "S"
            Else
                sStat = "R"
            End If

            sSendBuff = sSendBuff & "|" & sStat & "||" & Format(Now, "YYYYMMDDHHNNSS") & "||||A||||1||||||||||O" & vbCr
            
            If Trim(Urisys2400.Chart_No) <> "" Then
                sSendBuff = sSendBuff & "C|1|L|" & Trim(Urisys2400.Chart_No) & "^^^^|G" & vbCr
            End If
            
            sSendBuff = sSendBuff & "L|1|N"

            If Len(sSendBuff) >= 241 Then
                sNextSend = Mid(sSendBuff, 241)
                sSendBuff = left(sSendBuff, 240)
                sSendBuff = sSendBuff & Chr(23)

                m_iFrameN = m_iFrameN + 1
                m_iSendPhase = 2
            Else
                sSendBuff = sSendBuff & Chr(13) & Chr(3)
                GoTo Send_Terminate
            End If

        Case 2
            sSendBuff = m_iFrameN & sNextSend & Chr(13) & Chr(3)
            sNextSend = ""

Send_Terminate:
            m_iSendPhase = 3

        Case 3      'EOT
            comEQP.Output = Chr(4)   'EOT
            m_iFrameN = 1
            m_iPhase = 3
            m_iSendPhase = 1

            sState = "": sReqStatusCd = ""

            Exit Sub
    End Select

    ChkSum = ChkSum_ASTM(sSendBuff)
    sSendBuff = sSendBuff & ChkSum
    
    'Debug.Print "PC->Urisys2400 :" & Chr(2) & sSendBuff & Chr(13) & Chr(10)
    comEQP.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)
    
    With Urisys2400
        .BarCode_NO = ""
        .Chart_No = ""
        .SEQ = ""
        .Position = ""
        .RackNo = ""
        .Order_Str = ""
    End With

Err_Rtn:

End Sub

Private Function SeqNullSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqNullSearch = 0
    If brspread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .maxrows
            .Row = sCnt
            .Col = brCol
            If Trim(.text) = "" Then
                SeqNullSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Function f_funAdd_Server(ByVal strBarno As String, ByVal strTestcd As String, _
                                 ByVal strTestval As String, ByRef strOrdLst() As String) As Boolean
                                 
    Dim strErrMsg       As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval()  As String
    Dim strTmp1()       As String, strTmp2()    As String, strTmp   As String
    Dim intPos          As Integer, intIdx      As Integer
    Dim blnFlag         As Boolean
    
    blnFlag = False
    f_funAdd_Server = False
    
    strTmp = strTestcd: intPos = InStr(strTmp, ",")
    Do While intPos > 0
        blnFlag = False
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = Mid$(strTmp, 1, intPos - 1) Then
                blnFlag = True
                strTmp = Mid$(strTmp, 1, intPos - 1)
                Exit Do
            End If
        Next
        
        strTmp = Mid$(strTmp, intPos + 1)
        intPos = InStr(strTmp, ",")
    Loop
    
    If Not blnFlag Then
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = strTmp Then blnFlag = True: Exit For
        Next
    End If
    
    If blnFlag Then
        ReDim Preserve strSampleno(1 To 1) As String
        ReDim Preserve strOrdcd(1 To 1) As String
        ReDim Preserve strRstval(1 To 1) As String
        ReDim Preserve strTmp1(1 To 1) As String
        ReDim Preserve strTmp2(1 To 1) As String
        
        strSampleno(1) = strBarno
        strOrdcd(1) = strTmp
        strRstval(1) = strTestval
        strTmp2(1) = INS_CODE
        
        Call sl_online_result_ul_4&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
        If strErrMsg = "0" Then
            f_funAdd_Server = True
        Else
            Call ErrMsgProc("", strErrMsg)
        End If
    End If
                                
End Function

Private Function f_funAdd_QcServer(ByVal strBarno As String, ByVal strTestcd As String, _
                                 ByVal strTestval As String, ByRef strOrdLst() As String) As Boolean
                                 
    Dim strErrMsg       As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval()  As String
    Dim strTmp1()       As String, strTmp2()    As String, strTmp   As String
    Dim intPos          As Integer, intIdx      As Integer
    Dim blnFlag         As Boolean
    
    blnFlag = False
    f_funAdd_QcServer = False
    
    strTmp = strTestcd: intPos = InStr(strTmp, ",")
    Do While intPos > 0
        blnFlag = False
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = Mid$(strTmp, 1, intPos - 1) Then
                blnFlag = True
                strTmp = Mid$(strTmp, 1, intPos - 1)
                Exit Do
            End If
        Next
        
        strTmp = Mid$(strTmp, intPos + 1)
        intPos = InStr(strTmp, ",")
    Loop
    
    If Not blnFlag Then
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = strTmp Then blnFlag = True: Exit For
        Next
    End If
    
    If blnFlag Then
        ReDim Preserve strSampleno(1 To 1) As String
        ReDim Preserve strOrdcd(1 To 1) As String
        ReDim Preserve strRstval(1 To 1) As String
        ReDim Preserve strTmp1(1 To 1) As String
        ReDim Preserve strTmp2(1 To 1) As String
        
        strSampleno(1) = strBarno
        strOrdcd(1) = strTmp
        strRstval(1) = strTestval
        strTmp2(1) = INS_CODE
        
        Call sl_online_pc_98&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
        If strErrMsg = "" Then
            f_funAdd_QcServer = True
        Else
            Call ErrMsgProc("", strErrMsg)
        End If
    End If
                                
End Function

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

Private Function SeqSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch = 0
    If brspread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brspread
        If optSeq.Value = False Then
            For sCnt = 1 To .maxrows
                .Row = sCnt
                .Col = brCol
                If Trim(.text) = brSeq Then
                    SeqSearch = sCnt
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            Next sCnt
        Else
            For sCnt = 1 To .maxrows
                .Row = sCnt
                .Col = brCol
                If Val(spdResult1.StartingRowNumber + (Val(sCnt) - 1)) = Val(brSeq) Then
                    SeqSearch = sCnt 'brSeq
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            Next sCnt
        End If
    End With

End Function

Private Sub Command1_Click()

   
    Dim Arr()   As Byte
    Dim strDta  As String
  


    wkBuf = ""
    wkBuf = wkBuf & ENQ & vbCrLf
    wkBuf = wkBuf & "1H|\^&|||2338004|||||||P|6146010-10-05" & vbCrLf
    wkBuf = wkBuf & "P|1" & vbCrLf
    wkBuf = wkBuf & "O|1|525915450074|4^50033^2^1^SAMPLE||R||||||X|||20140530204028" & vbCrLf
    wkBuf = wkBuf & "R|1|^^^1|1.013|||||" & vbCrLf
    wkBuf = wkBuf & "C|1|I||I" & vbCrLf
    wkBuf = wkBuf & "R|2|^^^2|5|||||" & vbCrLf
    wkBuf = wkBuf & "C|2|I||I" & vbCrLf
    wkBuf = wkBuf & "R|3|^^^3|NEG|||||" & vbCrLf
    wkBuf = wkBuf & "C|3|I|#|I" & vbCrLf
    wkBuf = wkBuf & "R|4|^^^4|NEG|||||" & vbCrLf
    wkBuf = wkBuf & "C|4|I||I" & vbCrLf
    wkBuf = wkBuf & "R|5|^^^5|TR 25|mg/dL||||" & vbCrLf
    wkBuf = wkBuf & "C|5|I|S^*^#|I" & vbCrLf
    wkBuf = wkBuf & "R|6|^^^6|TR 50|mg/dL||||" & vbCrLf
    wkBuf = wkBuf & "C|6|I|*^#|I" & vbCrLf
    wkBuf = wkBuf & "R|7|^^^7|TR 5|mg/dL||||" & vbCrLf
    wkBuf = wkBuf & "C|7|I|*^#|I" & vbCrLf
    wkBuf = wkBuf & "R|8|^^^8|NORM|||||" & vbCrLf
    wkBuf = wkBuf & "C|8|I||I" & vbCrLf
    wkBuf = wkBuf & "R|9|^^^9|NEG|||||" & vbCrLf
    wkBuf = wkBuf & "C|9|I||I" & vbCrLf
    wkBuf = wkBuf & "R|10|^^^10|NEG|||||" & vbCrLf
    wkBuf = wkBuf & "C|10|I|#|I" & vbCrLf
    wkBuf = wkBuf & "R|11|^^^11|YELLO|||||" & vbCrLf
    wkBuf = wkBuf & "C|11|I||I" & vbCrLf
    wkBuf = wkBuf & "R|12|^^^12|CLEAR|||||" & vbCrLf
    wkBuf = wkBuf & "C|12|I||I" & vbCrLf
    wkBuf = wkBuf & "L|1|" & vbCrLf

'
'
'wkBuf = ""
'wkBuf = wkBuf & "" & vbLf
'wkBuf = wkBuf & "1H|\^&|||H7600" & vbLf
'wkBuf = wkBuf & "^1|||||host|TSR" & vbLf
'wkBuf = wkBuf & "EQ^REAL|P|1" & vbLf
'wkBuf = wkBuf & "Q|1" & vbLf
'wkBuf = wkBuf & "1H|\^&|||H7600^1|||||host|TSREQ^REAL|P|1" & vbLf
'wkBuf = wkBuf & "|^^   4051251759^0^5007^3^^S1^SC||AL" & vbLf
''wkBuf = wkBuf & "|^^   4041101029^0^5007^3^^S1^SC||AL" & vbLf
'wkBuf = wkBuf & "L||||||||O" & vbLf
'wkBuf = wkBuf & "" & vbLf
'wkBuf = wkBuf & "L|1|N" & vbLf
'wkBuf = wkBuf & "11" & vbLf
'wkBuf = wkBuf & "" & vbLf
'wkBuf = wkBuf & "Q|1|^^   4051251759^0^5007^3^^S1^SC||ALL||||||||O" & vbLf
'wkBuf = wkBuf & " | 89074 | 2" & vbLf
'wkBuf = wkBuf & "L|1|N" & vbLf
'wkBuf = wkBuf & "" & vbLf
'
''2Q|1|^3013002327||^^^ALL||||||||O
    Call Protocol_Call
    
'    RcvBuffer = wkBuf
'    Call EditRcvData
End Sub

Private Sub Form_Activate()
    
    If IS_SET = False Then Unload Me
    gWideOption = False
    Call cmdView_Click
End Sub

Private Sub Form_Load()
    
    
    
'    Me.Show
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    CaptionBar1.Caption = INS_NAME & " Communication"
    
    장비코드 = GetString(HKEY_CURRENT_USER, REG_Machine, REG_Code)
    
    Call cmdClear               ' 초기화
    Call f_subSet_ItemHeader    ' 리스트해더
    Call f_subSet_ItemList      ' 검사항목
    
    Call f_subSet_ComCharacter  ' 통신문자
    Call f_subGet_Setting       ' 통신설정
    
    Call cmdRun                 ' 실행
    
    mskRstDate.text = Format$(Now, "YYYYMMDD")
    mskRstDate1.text = Format$(Now, "YYYYMMDD")
    
    DptDate.Value = Now
    
    gLocalIP = Winsock1.LocalIP
    gLocalNm = Winsock1.LocalHostName
    
    Open App.Path + "\" + REG_INSNAME + ".Log" For Append As #50

    Print #50, Chr(13) + Chr(10);
   
    f_strJOB_FLAG = "1":    f_intSampleNo = 0
    tabWork.Tab = 0
    Or_Seq = 1
    intRow = 0
    chkEnq = 0
    SendCount = 0
    
    '
    ' 스프레드 소트 기능 사용
    '
    Dim sCol As Integer
    With spdResult1
        .UserColAction = UserColActionSort
        .ColUserSortIndicator(3) = ColUserSortIndicatorAscending
    End With
     
    With spdResult2
        .UserColAction = UserColActionSort
        .ColUserSortIndicator(3) = ColUserSortIndicatorAscending
    End With
     
    Call cmdClear
    
    
    '==============================
    m_iPhase = 1
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1

    '==============================

End Sub

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
                .Handshaking = Trim(mAdoRs.Fields("COM_HANDSHAK") & "")
                .RTSEnable = Trim(mAdoRs.Fields("COM_RTS") & "")
                
                
                
                
'                .InputMode = Trim(mAdoRs.Fields("COM_INPUTMOD") & "")
'                .DTREnable = Trim(mAdoRs.Fields("COM_DTR") & "")
'                .EOFEnable = Trim(mAdoRs.Fields("COM_EOF") & "")
'                .NullDiscard = Trim(mAdoRs.Fields("COM_NULDIS") & "")
'                .InBufferSize = Trim(mAdoRs.Fields("COM_IBS") & "")
'                .InputLen = Trim(mAdoRs.Fields("COM_INLEN") & "")
'                .OutBufferSize = Trim(mAdoRs.Fields("COM_OBS") & "")
'                .ParityReplace = Trim(mAdoRs.Fields("COM_PTR") & "")





                .RThreshold = Trim(mAdoRs.Fields("COM_RTH") & "")
                .SThreshold = Trim(mAdoRs.Fields("COM_STH") & "")
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
    
    Close #50

End Sub

Private Sub FrameError_Click()
    List1.Visible = False
End Sub

Private Sub imgPort_DblClick()
    
    If lvwCuData.Visible Then
        lvwCuData.Visible = False
    Else
        lvwCuData.Visible = True
        lvwCuData.ZOrder 0
    End If
    
End Sub

Private Sub imgReceive_DblClick()

    If pnlCom2.Visible = True Then
        pnlCom2.Visible = False
    Else
        pnlCom2.Visible = True
        pnlCom2.ZOrder 0
    End If
    
End Sub

Private Sub imgSend_DblClick()
    
    If pnlCom.Visible = True Then
        pnlCom.Visible = False
    Else
        pnlCom.Visible = True
        pnlCom.ZOrder 0
    End If

End Sub

Private Sub Label6_Click()
    If Command1.Visible = False Then
        Command1.Visible = True
    Else
        Command1.Visible = False
    End If
End Sub

Private Sub Label7_Click()
    If List1.Visible = True Then
        List1.Visible = False
        spd저장체크.Visible = False
        
    Else
        List1.Visible = True
        spd저장체크.Visible = True
    End If
    
End Sub

Private Sub Label9_DblClick()

    If COM_MODE = "1" Then
        COM_MODE = "0"
        ShowMessage "인터페이스 내용을 화면에 출력하지 않습니다."
    Else
        COM_MODE = "1"
        ShowMessage "인터페이스 내용을 화면에 출력합니다."
    End If
End Sub





Private Sub mskRstDate_GotFocus()

    With mskRstDate
        .SelStart = 0
        .SelLength = Len(.text) + 2
    End With '
    
End Sub


Private Sub mskRstDate_KeyPress(KeyAscii As Integer)

    If Not KeyAscii = vbKeyBack Then mskRstDate.SelLength = 1
    
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

Private Sub spdResult1_Change(ByVal Col As Long, ByVal Row As Long)
Dim pBarcode As String
Dim pGrid_Point As Integer
Dim mBarcode As String
Dim pMessage As String
Dim vRow As Integer

    vRow = Row
    
    If Col <> 3 Then Exit Sub
    
    With spdResult1
        .Row = Row
        .Col = 3
        
        pBarcode = Mid(.text, 1, 12)
        
        If (pBarcode) = "" Or Len(pBarcode) < 12 Then Exit Sub
        
        With spdResult1
        
            .Col = 3
            .Row = Row
            mBarcode = .text
            

            pGrid_Point = Barcode_Search(spdResult1, mBarcode, 3, vRow)

        
            If pGrid_Point >= 1 Then

            
                pMessage = ""
                .Row = pGrid_Point
                .Col = 2: pMessage = pMessage & .text
                .Col = 3: pMessage = pMessage & " | " & .text
                .Col = 5: pMessage = pMessage & " | " & .text
                txtHelp.text = txtHelp.text & "중복 ▶" & pMessage & vbCrLf
                
                MsgBox "☞ 중복 ☜ " & pMessage & " | 중복 등록된 검체번호 입니다. "
                
                .Row = vRow
                .Col = 3
                .text = "::중복검체::"
            
            End If
        End With
        
        If pGrid_Point = 0 Then
        
            Call OrderByRead(vRow, pBarcode)
            
        End If

        
    End With
    
    
End Sub

Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)
   Dim intCol1 As Integer
    Dim intCol2 As Integer
    Dim intRow1 As Integer
    Dim intRow2 As Integer
    Dim iCnt    As Integer
    
    Dim varTmp  As Variant
    Dim pResult   As String
    Dim pTestName As String
    Dim mTestName As Boolean
    Dim pBackColor As String
    
    ssInformation.Caption = ""
    
    If Row = 0 Then Exit Sub
    
    With spdRstview
    
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
        
    End With
    
    mTestName = False
    
    With spdResult1
    
        Dim pname    As String
        Dim pBarcode As String
        Dim pChartno As String
        
        .Row = Row
        
        .Col = 3: pBarcode = .text
        .Col = 4: pChartno = .text
        .Col = 5: pname = .text
        
        If pChartno <> "" And pBarcode <> "" Then
            ssInformation.Caption = "[ " & pname & " ]*[" & pBarcode & "]*[" & pChartno & "]"
        End If
        
        For iCnt = gItemStartPos To .MaxCols
            .Row = Row: .Col = iCnt

            If chkOrderCheck.Value = 1 Then
                If .BackColor = &HC6FEFF Then
                    .GetText .Col, .Row, varTmp:        pResult = Trim$(varTmp)
                    .GetText .Col, 0, varTmp:           pTestName = Trim$(varTmp): pBackColor = .BackColor
                    
                    If spdRstview.maxrows = 1 And mTestName = False Then
                        mTestName = True
                        spdRstview.SetText 1, spdRstview.maxrows, pTestName
                        spdRstview.SetText 2, spdRstview.maxrows, pResult
                        
                        spdRstview.Row = spdRstview.maxrows
                        spdRstview.Col = 2
                        
                        spdRstview.BackColor = &HC6FEFF
                        
                    Else
                        spdRstview.maxrows = spdRstview.maxrows + 1
                        spdRstview.SetText 1, spdRstview.maxrows, pTestName
                        spdRstview.SetText 2, spdRstview.maxrows, pResult
                        
                        spdRstview.Row = spdRstview.maxrows
                        spdRstview.Col = 2
                        
                        spdRstview.BackColor = &HC6FEFF
                    
                    End If
                
                End If
            Else
                If .BackColor = &HC6FEFF Or .text <> "" Then
                    .GetText .Col, .Row, varTmp:        pResult = Trim$(varTmp)
                    .GetText .Col, 0, varTmp:           pTestName = Trim$(varTmp): pBackColor = .BackColor
                    
                    If spdRstview.maxrows = 1 And mTestName = False Then
                        mTestName = True
                        spdRstview.SetText 1, spdRstview.maxrows, pTestName
                        spdRstview.SetText 2, spdRstview.maxrows, pResult
                        
                        spdRstview.Row = spdRstview.maxrows
                        spdRstview.Col = 2
                        
                        spdRstview.BackColor = &HC6FEFF
                        
                    Else
                        spdRstview.maxrows = spdRstview.maxrows + 1
                        spdRstview.SetText 1, spdRstview.maxrows, pTestName
                        spdRstview.SetText 2, spdRstview.maxrows, pResult
                        
                        spdRstview.Row = spdRstview.maxrows
                        spdRstview.Col = 2
                        
                        spdRstview.BackColor = &HC6FEFF
                    
                    End If
                
                End If
            
            End If
            

        Next
    End With
End Sub

Private Sub spdResult1_KeyPress(KeyAscii As Integer)
    Dim pBarcode As String
    
    Select Case KeyAscii
        Case vbKeyReturn
            With spdResult1
                .Col = .ActiveCol
                .Row = .ActiveRow
                pBarcode = .text
                
                If .Col = 3 And Len(pBarcode) = 10 Then
                    
                    Call OrderByRead(.Row, pBarcode)
                End If
                
            End With
    End Select
    
End Sub

Private Sub spdResult1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    OldRow = Row
    If Row = NewRow Then Exit Sub
    If Row < 0 Then Exit Sub
    If NewRow < 0 Then Exit Sub
    
    miLeaveCell = 2
        
    Call spdIntList_Click(1, NewRow)
End Sub

Private Sub spdIntList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    If miLeaveCell = 1 Then
        miLeaveCell = 0
        
        Exit Sub
    End If
    
    miLeaveCell = miLeaveCell - 1
    
    'Call DisplayResult2(CInt(Row))
End Sub


Private Sub spdResult1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

Dim oMenu As cPopupMenu
Dim lMenuChosen As Long

    Set oMenu = New cPopupMenu

    lMenuChosen = oMenu.Popup(" ▒ 바코드등록", "-", " ▒ Row 삭제", "-", " ▒ 결과 수정", "-", " ▒ 서버 저장")

    Select Case lMenuChosen
        Case 1
            uplBarcode.Visible = True
            
            txtBarcode_Keyin.text = ""
            txt접수번호.text = ""
            txt순번.text = ""
            
            If Row = 0 Then
                uplBarcode.Visible = False
                MsgBox " ::: 바코드 환자결과를  선택하십시요..!!", vbExclamation
                Exit Sub
            End If
            
            txt순번.text = Row
            
            With spdResult1
                .Row = Row
                .Col = 9
                txt접수번호.text = .text
            End With
            
            txtBarcode_Keyin.SetFocus
        Case 3
            With spdResult1
                .Col = Col
                .Row = Row
                .Action = ActionDeleteRow
                .maxrows = .maxrows - 1
            End With
        Case 5
            Dim NEW_Result As String
            Dim OLD_Result As String
            
            With spdResult1
                .Row = Row: .Col = Col: OLD_Result = .text
            End With
            
            NEW_Result = InputBox("::: 수정할 결과를 입력하십시요." & vbCrLf & vbCrLf & "::: 수정 전 결과 : " & OLD_Result, "■ 결과수정", OLD_Result)
            
            If NEW_Result <> OLD_Result Then
                With spdResult1
                    .Row = Row: .Col = Col: .text = NEW_Result
                End With
            End If

        Case 7
            Call cmdAppend_Click(0)
    End Select
End Sub


Private Sub Timer1_Timer()

    Call COM_OUTPUT(ENQ)
'    Debug.Print ENQ

End Sub

Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub

Private Sub Form_Resize()
    Dim I As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 4300 Then Exit Sub
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For I = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(I).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - I)) + (70 * (cmdAction.UBound - I)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
    
    Label7.left = Me.Width - Label7.Width - 100
    chkSound.left = Me.Width - Label7.Width - chkSound.Width - 200
    chkAuto.left = Me.Width - Label7.Width - chkSound.Width - chkAuto.Width - 400
    chkJubsu.left = Me.Width - Label7.Width - chkSound.Width - chkAuto.Width - chkJubsu.Width - 1000
    chkOrderCheck.left = Me.Width - Label7.Width - chkSound.Width - chkAuto.Width - chkJubsu.Width - chkOrderCheck.Width
    
    
    spdRstview.left = Me.Width - spdRstview.Width - 300
    ssInformation.left = Me.Width - ssInformation.Width - 300
    pnlPort.left = Me.Width - pnlPort.Width - 300
    
    spdResult1.Width = Me.Width - spdRstview.Width - 400
    spdResult2.Width = Me.Width - 400
    
    SSPanel4.Width = Me.Width - ssInformation.Width - 420
    
    tabWork.Width = Me.Width - 200
    
    Call cmdAppend(0).Move(SSPanel4.Width - cmdAppend(0).Width - 50)
    Call cmdView.Move(SSPanel4.Width - cmdView.Width - 100 - cmdAppend(0).Width)
    Call cmdMachine.Move(SSPanel4.Width - cmdMachine.Width - 150 - cmdAppend(0).Width - cmdView.Width)
    
'    Call chkSound.Move(SSPanel4.Width - chkSound.Width - 200 - cmdAppend(0).Width - cmdView.Width - cmdMachine.Width)
'    Call chkAuto.Move(SSPanel4.Width - chkAuto.Width - 250 - cmdAppend(0).Width - cmdView.Width - cmdMachine.Width - chkSound.Width)
'    Call chkJubsu.Move(SSPanel4.Width - chkAuto.Width - 300 - cmdAppend(0).Width - cmdView.Width - cmdMachine.Width - chkSound.Width - chkJubsu.Width)
    
    tabWork.Height = ScaleHeight - 1250
    
    spdResult1.Height = ScaleHeight - 2250
    spdResult2.Height = ScaleHeight - 2250
    spdRstview.Height = ScaleHeight - 2250
    gWideOption = False
    
    
    txtHelp.Top = tabWork.Height - txtHelp.Height - 50
    txtHelp.left = Me.Width - txtHelp.Width - 300

    spdRstview.Height = spdRstview.Height - txtHelp.Height - 2

End Sub

Private Sub txtBarCode_Change()

    If txtBarCode.SelStart = txtBarCode.MaxLength Then SendKeys "{TAB}"
    
End Sub

Private Sub txtBarCode_GotFocus()

    With txtBarCode
        .SelStart = 0
        .SelLength = Len(.text)
    End With
    
End Sub


Private Sub txtBarcode_Keyin_Change()
Dim mResult As String
    If Len(txtBarcode_Keyin.text) >= 12 Then
        With spdResult1
            .Row = txt순번.text
            .Col = 7:            .text = txtBarcode_Keyin.text
            .Col = 1:            .text = 1
            .Col = 11:           mResult = .text
        End With
        
        Call OrderByRead(txt순번.text, txtBarcode_Keyin.text)

        txtBarcode_Keyin.text = ""
        
        uplBarcode.Visible = False
        
'        txt접수번호.text = ""
'        txt순번.text = txt순번.text + 1
'
'        spdResult1.Row = Val(txt순번.text)
'
'        txtBarcode_Keyin.SetFocus

    End If
End Sub




Private Function OrderByRead(pROW As Integer, pBarcode As String)

    Dim pCOL   As Integer
    Dim varChk  As Variant, varBar As Variant, varNum As Variant
    Dim iRow    As Integer, iCnt   As Integer
    Dim strEqpCd As String
    Dim itemX   As ListItem
    Dim blnFlag As Boolean
    Dim pGrid_Point As Integer
    Dim pSeq        As String
       
    If Len(pBarcode) >= 12 Then
        strBarno = ""
        
        With spdResult1
            .Col = 2: .Row = pROW
            pSeq = .text
        End With
        
        pSeq = Format(pSeq, "0000")
        
        Call f_DeleTe_MSSQL(pSeq, Format(DptDate.Value, "YYYYMMDD"))

        Set mAdoRs = f_subSet_WorkList_Barcode(pBarcode, INS_CODE)
        If RecordChk = False Then
            List1.AddItem ("샘플번호 " & PatientID & " 는 등록되지 않은 검체입니다. 확인하세요. - ReceiveTheData")
        Else
            With spdResult1
                Do Until mAdoRs.EOF

                    If strBarno <> mAdoRs.Fields("REQNO") Then
'
                        .SetText 1, pROW, "1"
                        .SetText 3, pROW, pBarcode
                        .SetText 4, pROW, Trim(mAdoRs("CSTNM")) & ""
                        .SetText 5, pROW, Trim(mAdoRs("PATNM")) & ""
                        
                        .SetText 6, pROW, Trim(mAdoRs("REQDTE"))
                        .SetText 7, pROW, Trim(mAdoRs("REQNO"))
                        .SetText 8, pROW, Trim(mAdoRs("IDNO"))
                        .SetText 12, pROW, Trim(mAdoRs("WRKDTE")) & "/" & Trim(mAdoRs("WORKNO")) & "/" & Trim(mAdoRs("LABEMP"))
                        
                        .SetText 14, pROW, "오더없음"
                        
                        
                    End If

    
                    If Len(Trim(mAdoRs.Fields("ITEMCD"))) > 0 Then
                        strEqpCd = f_funGet_CODE_SUB(Trim(mAdoRs.Fields("ITEMCD")))
'                        Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
'                        If Not itemX Is Nothing Then
                        If strEqpCd > 0 Then
                            spdResult1.SetText 1, pGrid_Point, "0"
                            spdResult1.Col = strEqpCd
                            spdResult1.Row = pROW
                            spdResult1.BackColor = &HC6FEFF
                            spdResult1.CellTag = Trim(mAdoRs.Fields("ITEMCD"))
                            
                            If Trim(mAdoRs.Fields("ITEMCD")) & "" = "U111" Then
                                    .SetText 14, pROW, "WBC: 0~3 " & vbCrLf & "RBC: 0~3 " & vbCrLf & "E.P cell: 0~3 " & vbCrLf & "Others: None"
                            End If
                            
                            
                            If medGetP(.CellNote, 2, "/") = "*" Then
                                spdResult1.BackColor = &HC0C0FF
                            End If
                        End If
'                        End If
                        strBarno = mAdoRs.Fields("REQNO")
                    End If
    
                    mAdoRs.MoveNext
                Loop
            End With
        End If
    End If
    
    Call cmdAppend_Click(0)
    
End Function

' 통신상태 확인 관련이벤트
' ------------------------------------------------------------------------
Private Sub txtCom_Change()
    txtCom.SelStart = Len(txtCom.text)
End Sub

Private Sub cmdCOMLoad_Click()
    Dim I               As Long
    Dim lngFIleNum      As Long
    Dim strTemp         As String
    Dim strTemp2        As String
    Dim bteBuffer()     As Byte
    
On Error GoTo ErrorRoutine
    
    With cdlFile
        .CancelError = True
        .FileName = App.Path & "\comm.txt"
        .ShowOpen
        lngFIleNum = FreeFile
        
        Open .FileName For Binary Access Read As #lngFIleNum
        
        txtCOM2.text = ""
        ReDim bteBuffer(LOF(lngFIleNum))
        Get #lngFIleNum, , bteBuffer

        strTemp = StrConv(bteBuffer, vbUnicode)
        txtCOM2.text = strTemp
                
        Close #lngFIleNum
    End With
    Exit Sub
    
ErrorRoutine:
    Close #lngFIleNum
        
End Sub

Private Sub cmdCOMSave_Click()
    Dim lngFIleNum      As Long
    
On Error GoTo ErrorRoutine

    With cdlFile
        .CancelError = True
        .FileName = App.Path & "\comm.txt"
        .ShowSave
        lngFIleNum = FreeFile
        
        Open .FileName For Append As #lngFIleNum
        Print #lngFIleNum, _
              Format(Date, "YYYY년 MM월 DD일") & "  "; Time & vbNewLine & _
              "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbNewLine & _
              txtCom.text & _
              "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbNewLine
    Close #lngFIleNum
    End With
    Exit Sub
    
ErrorRoutine:
    Close #lngFIleNum

End Sub

Private Sub cmdCOMOutput_Click()
    'Call COM_OUTPUT(StrConv(charCOM_Convert(txtCom.SelText), vbFromUnicode))
    Call COM_OUTPUT(charCOM_Convert(txtCom.SelText))
End Sub

Private Sub cmdCOMClear_Click()
    mlngRecLen = 0
    txtCom.text = ""
End Sub

Private Sub cmdCOMClear2_Click()
    txtCOM2.text = ""
End Sub

Private Sub cmdCOMInput_Click()

    Dim bytTemp() As Byte
    
    bytTemp = StrConv(charCOM_Convert(txtCom.SelText), vbFromUnicode)

   ' Call ComReceive(txtCom.SelText)
    
End Sub

Private Sub cmdCOMInput2_Click()
    
    Dim bytTemp() As Byte
    
    If txtCOM2.SelLength = 0 Then
        bytTemp = StrConv(charCOM_Convert(txtCOM2.text), vbFromUnicode)
    Else
        bytTemp = StrConv(charCOM_Convert(txtCOM2.SelText), vbFromUnicode)
    End If

   ' Call ComReceive(txtCOM2.SelText)

End Sub

Private Function Barcode_Search(ByVal brspread As Object, ByVal brBarcode As String, ByVal brCol As Integer, ByVal brRow As Integer) As Long
Dim sCnt As Long
Dim mBarcode As String
    If brspread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brspread

        For sCnt = 1 To .maxrows
            .Row = sCnt
            .Col = brCol
            mBarcode = .text
            If mBarcode = brBarcode Then
                If sCnt <> brRow Then
                    Barcode_Search = sCnt 'brSeq
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            End If
        Next sCnt

    End With

End Function

Private Sub cmdCOMOutput2_Click()
    
    If txtCOM2.SelLength = 0 Then
        Call COM_OUTPUT(charCOM_Convert(txtCOM2.text))
    Else
        Call COM_OUTPUT(charCOM_Convert(txtCOM2.SelText))
    End If
    
End Sub
' ------------------------------------------------------------------------
' 통신상태 확인 관련이벤트


Private Sub txtResult_DblClick()
'    txtResult.text = ""
    List1.text = ""
    
'    If txtResult.Visible Then txtResult.Visible = False
    List1.Visible = True
End Sub

Private Sub uplBarcode_CloseMe()
    uplBarcode.Visible = False
End Sub

Private Sub uplUrineMicro_CloseMe()
    Call cmd닫기_Click
End Sub

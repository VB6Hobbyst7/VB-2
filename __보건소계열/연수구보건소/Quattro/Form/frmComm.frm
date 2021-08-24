VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmComm 
   Caption         =   "Interface"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   15360
   WindowState     =   2  '최대화
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
      Left            =   4365
      Top             =   4545
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
            Picture         =   "frmComm.frx":0000
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":059A
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":0B34
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":10CE
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":1668
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":1C02
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm comEQP 
      Left            =   4320
      Top             =   5130
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   1
      RThreshold      =   1
      SThreshold      =   1
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
            Picture         =   "frmComm.frx":219C
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":2736
            Key             =   "NOT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":2CD0
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":326A
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":3AFC
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":3C56
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":3DB0
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
         TabIndex        =   16
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
         TabIndex        =   17
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
         TabIndex        =   18
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
         TabIndex        =   19
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
         TransparentPicture=   "frmComm.frx":3F0A
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
         TabIndex        =   6
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
         TabIndex        =   5
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
      Width           =   15360
      _ExtentX        =   27093
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Receive : "
         Height          =   180
         Left            =   14055
         TabIndex        =   4
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send : "
         Height          =   180
         Left            =   13020
         TabIndex        =   3
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port : "
         Height          =   180
         Left            =   11925
         TabIndex        =   2
         Top             =   285
         Width           =   510
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   14925
         Picture         =   "frmComm.frx":5794
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   13695
         Picture         =   "frmComm.frx":5D1E
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12555
         Picture         =   "frmComm.frx":62A8
         Top             =   255
         Width           =   240
      End
   End
   Begin TabDlg.SSTab tabWork 
      Height          =   8370
      Left            =   45
      TabIndex        =   7
      Top             =   585
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   14764
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " WorkList"
      TabPicture(0)   =   "frmComm.frx":6832
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdResultQuery"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdWorkList1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdStartNo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdWordQuery"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdEot"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSearch"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdAppend(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdWorkList"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "spdWorkList"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkAuto"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "spdResult1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkReTest"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdSel(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdSel(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "SSPanel1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "SSFrame1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "FrameError"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "SSPanel2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "optBar"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "optSeq"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdOrder"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "FileBep2000"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   " 받은 결과"
      TabPicture(1)   =   "frmComm.frx":684E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSPanel4"
      Tab(1).Control(1)=   "chkServer"
      Tab(1).Control(2)=   "SSPanel3"
      Tab(1).Control(3)=   "cmdAppend(1)"
      Tab(1).Control(4)=   "lvwCuData"
      Tab(1).Control(5)=   "cmdRstQuery"
      Tab(1).Control(6)=   "cmdSel(2)"
      Tab(1).Control(7)=   "cmdSel(3)"
      Tab(1).Control(8)=   "spdResult2"
      Tab(1).ControlCount=   9
      Begin VB.FileListBox FileBep2000 
         Height          =   990
         Left            =   7515
         TabIndex        =   59
         Top             =   450
         Visible         =   0   'False
         Width           =   2775
      End
      Begin BHButton.BHImageButton cmdOrder 
         Height          =   375
         Left            =   5085
         TabIndex        =   38
         Top             =   450
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         Caption         =   "업무나열서"
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
      Begin VB.OptionButton optSeq 
         BackColor       =   &H80000004&
         Caption         =   "Seq"
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
         Left            =   6525
         TabIndex        =   48
         Top             =   0
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton optBar 
         BackColor       =   &H80000004&
         Caption         =   "Bar"
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
         Left            =   7245
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin FPSpread.vaSpread spdResult2 
         Height          =   7350
         Left            =   -74910
         TabIndex        =   40
         Top             =   900
         Width           =   15015
         _Version        =   196608
         _ExtentX        =   26485
         _ExtentY        =   12965
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   4
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   4
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ShadowColor     =   16761024
         SpreadDesigner  =   "frmComm.frx":686A
         UserResize      =   0
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   465
         Left            =   11475
         TabIndex        =   37
         Top             =   405
         Visible         =   0   'False
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   33023
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
         Begin VB.OptionButton optSpc1 
            BackColor       =   &H000080FF&
            Caption         =   "종합검진"
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
            Height          =   285
            Left            =   270
            TabIndex        =   44
            Top             =   90
            Value           =   -1  'True
            Width           =   1365
         End
         Begin VB.OptionButton optSpc2 
            BackColor       =   &H000080FF&
            Caption         =   "계약자/고계약자"
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
            Height          =   285
            Left            =   1620
            TabIndex        =   43
            Top             =   90
            Width           =   1860
         End
      End
      Begin Threed.SSFrame FrameError 
         Height          =   2490
         Left            =   8130
         TabIndex        =   24
         Top             =   5805
         Width           =   6930
         _Version        =   65536
         _ExtentX        =   12224
         _ExtentY        =   4392
         _StockProps     =   14
         Caption         =   "오류내역"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ListBox List1 
            Height          =   2040
            Left            =   135
            TabIndex        =   29
            Top             =   225
            Width           =   6630
         End
         Begin VB.TextBox txtResult 
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   2040
            Left            =   135
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   28
            Top             =   225
            Visible         =   0   'False
            Width           =   6630
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   3030
         Left            =   90
         TabIndex        =   25
         Top             =   5310
         Width           =   7995
         _Version        =   65536
         _ExtentX        =   14102
         _ExtentY        =   5345
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
         Begin FPSpread.vaSpread spdRstview 
            Height          =   2880
            Left            =   45
            TabIndex        =   26
            Top             =   135
            Width           =   7875
            _Version        =   196608
            _ExtentX        =   13891
            _ExtentY        =   5080
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   0
            ColsFrozen      =   4
            EditEnterAction =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridShowVert    =   0   'False
            GridSolid       =   0   'False
            MaxCols         =   8
            MaxRows         =   10
            RetainSelBlock  =   0   'False
            ScrollBarMaxAlign=   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmComm.frx":6C2D
            UserResize      =   0
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   465
         Left            =   90
         TabIndex        =   31
         Top             =   390
         Width           =   3690
         _Version        =   65536
         _ExtentX        =   6509
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   33023
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
         Begin VB.TextBox txtBarCode 
            Height          =   300
            Left            =   3690
            MaxLength       =   12
            TabIndex        =   45
            Top             =   90
            Visible         =   0   'False
            Width           =   1500
         End
         Begin MSMask.MaskEdBox mskOrdDate1 
            Height          =   300
            Left            =   2475
            TabIndex        =   32
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
         Begin MSMask.MaskEdBox mskOrdDate 
            Height          =   300
            Left            =   1170
            TabIndex        =   33
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
         Begin VB.Label Label7 
            BackColor       =   &H000080FF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2310
            TabIndex        =   35
            Top             =   150
            Width           =   315
         End
         Begin VB.Label Label6 
            BackColor       =   &H000080FF&
            Caption         =   "처방일자 :"
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
            Height          =   225
            Left            =   120
            TabIndex        =   34
            Top             =   150
            Width           =   1095
         End
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   556
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":71D5
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   556
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":7657
      End
      Begin VB.CheckBox chkReTest 
         Caption         =   "재검/QC"
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
         Left            =   10890
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   1380
      End
      Begin FPSpread.vaSpread spdResult1 
         Height          =   4380
         Left            =   3465
         TabIndex        =   23
         Top             =   885
         Width           =   11640
         _Version        =   196608
         _ExtentX        =   20532
         _ExtentY        =   7726
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   4
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   16761024
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   4
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   2
         ShadowColor     =   16761024
         SpreadDesigner  =   "frmComm.frx":7AC5
         UserResize      =   0
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "Auto(서버)"
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
         Left            =   13590
         TabIndex        =   15
         Top             =   45
         Width           =   1410
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TEST"
         Height          =   285
         Left            =   9780
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   990
      End
      Begin FPSpread.vaSpread spdWorkList 
         Height          =   3975
         Left            =   90
         TabIndex        =   8
         Top             =   885
         Width           =   3345
         _Version        =   196608
         _ExtentX        =   5900
         _ExtentY        =   7011
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   16761024
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   7
         MaxRows         =   1
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   2
         ShadowColor     =   16761024
         SpreadDesigner  =   "frmComm.frx":7E52
         UserResize      =   0
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   3
         Left            =   -74640
         TabIndex        =   12
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":8216
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   -74910
         TabIndex        =   13
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":8698
      End
      Begin BHButton.BHImageButton cmdRstQuery 
         Height          =   375
         Left            =   -69195
         TabIndex        =   22
         Top             =   480
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
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   4830
         Left            =   -67980
         TabIndex        =   11
         Top             =   900
         Visible         =   0   'False
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   8520
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
      Begin BHButton.BHImageButton cmdAppend 
         Height          =   375
         Index           =   1
         Left            =   -67935
         TabIndex        =   21
         Top             =   480
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
      Begin BHButton.BHImageButton cmdWorkList 
         Height          =   390
         Left            =   90
         TabIndex        =   20
         Top             =   4905
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   688
         Caption         =   "WorkList 등록"
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
         BackColor       =   14737632
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAppend 
         Height          =   420
         Index           =   0
         Left            =   13860
         TabIndex        =   30
         Top             =   5355
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
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
         Left            =   3855
         TabIndex        =   36
         Top             =   450
         Width           =   1170
         _ExtentX        =   2064
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
      Begin BHButton.BHImageButton cmdEot 
         Height          =   375
         Left            =   11430
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         Caption         =   "초기화"
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
      Begin BHButton.BHImageButton cmdWordQuery 
         Height          =   420
         Left            =   10395
         TabIndex        =   41
         Top             =   5355
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
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
      Begin BHButton.BHImageButton cmdStartNo 
         Height          =   375
         Left            =   13815
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
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
      Begin BHButton.BHImageButton cmdWorkList1 
         Height          =   390
         Left            =   90
         TabIndex        =   46
         Top             =   4905
         Visible         =   0   'False
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   688
         Caption         =   "불러오기"
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
         TabIndex        =   49
         Top             =   405
         Width           =   5625
         _Version        =   65536
         _ExtentX        =   9922
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   33023
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
            ItemData        =   "frmComm.frx":8B06
            Left            =   3735
            List            =   "frmComm.frx":8B13
            Style           =   2  '드롭다운 목록
            TabIndex        =   54
            Top             =   90
            Width           =   1770
         End
         Begin MSMask.MaskEdBox mskRstDate 
            Height          =   300
            Left            =   1260
            TabIndex        =   52
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
            TabIndex        =   53
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
            BackColor       =   &H000080FF&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   90
            TabIndex        =   51
            Top             =   135
            Width           =   1125
         End
         Begin VB.Label Label8 
            BackColor       =   &H000080FF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2400
            TabIndex        =   50
            Top             =   150
            Width           =   315
         End
      End
      Begin Threed.SSCheck chkServer 
         Height          =   165
         Left            =   -65550
         TabIndex        =   55
         Top             =   630
         Visible         =   0   'False
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   291
         _StockProps     =   78
         Caption         =   " SERVER DATA"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   465
         Left            =   -63525
         TabIndex        =   56
         Top             =   405
         Visible         =   0   'False
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   33023
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
         Begin VB.OptionButton optGae 
            BackColor       =   &H000080FF&
            Caption         =   "계약자/고계약자"
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
            Height          =   285
            Left            =   1620
            TabIndex        =   58
            Top             =   90
            Width           =   1860
         End
         Begin VB.OptionButton optJong 
            BackColor       =   &H000080FF&
            Caption         =   "종합검진"
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
            Height          =   285
            Left            =   270
            TabIndex        =   57
            Top             =   90
            Value           =   -1  'True
            Width           =   1365
         End
      End
      Begin BHButton.BHImageButton cmdResultQuery 
         Height          =   375
         Left            =   6300
         TabIndex        =   60
         Top             =   450
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         Caption         =   "결과조회"
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
Dim Patiant_Recevid As Boolean
Dim sStxCheck As Integer
Dim sEtxCheck As Integer
' --------------------------------------------------------------
Private Type typeElecsys2010
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
Dim Elecsys2010 As typeElecsys2010
Dim strOrdLst As String

Dim fQuattro() As String
Dim fQuattro_1  As Variant
Dim SendData(10)     As String
Dim SendCount        As String
Dim Or_Seq           As Integer
Dim SendBuffW           As String
Dim SendBuffT           As String
Dim intRow          As Integer
Dim brStr           As String

Dim fRcvString As String
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
Private mAdoRs2     As ADODB.Recordset
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
Dim pName   As String
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

Private Type TYPE_CD
    strEqpCd    As String
    intCnt      As Integer
    strTestCd(100) As String
End Type

Private f_typCode() As TYPE_CD

Dim PatientID As String    'Q Message Pattern Check
Dim PatientSeq As String
Dim PatientDisk As String
Dim PatientRack As String
Dim PatientPos As String

Dim SeqNo As String
Dim RecordChk   As Boolean

Dim G_CLVALU    As String
Dim G_CHVALU    As String
Dim G_EVALUATE  As String
Dim G_PANIC     As String
Dim G_DELTA     As String
Dim strFrameNo  As Integer
Dim OrderCnt As Integer
Dim vRow As Integer
Dim sPatiant_No As String
Dim tmpOrder       As String
Dim sBEP2000(10) As String
Dim mBEP2000(10) As String

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
    Dim x      As Long
    Dim ChkCS  As String
    Dim SumCS  As String
    Dim AddCS  As Long
    For x = 1 To Len(Source)
        AddCS = AddCS + Asc(Mid(Source, x, 1))
    Next x
    SumCS = Hex(AddCS)
    ChkCS = Mid(SumCS, Len(SumCS) - 1, 1)
    ChkCS = ChkCS & Right(SumCS, 1)
    MakeCS = ChkCS
End Function

Public Function GetChkSum(ByVal pMsg As String) As String

    Dim lngChkSum   As Long
    Dim i           As Long

    For i = 1 To Len(pMsg)
        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
    Next

    If lngChkSum = 0 Then
        GetChkSum = "00"
    Else
        GetChkSum = LCase(Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2))
    End If

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

Private Sub f_subGet_JobList(ByVal strKeyno As String, ByRef strOrder As String, _
                             ByRef intOrdCnt As Integer, ByRef strSpec As String, _
                             ByRef strPcFlag As String)

    Dim adoRS1  As New ADODB.Recordset
    Dim adoRS2  As New ADODB.Recordset
    Dim sqlDoc  As String
    
    strOrder = "":  strPcFlag = "  ":   strSpec = "SE": intOrdCnt = 0
    sqlDoc = "select ORD_CODE, CHART_NO From L3A01" & _
             " where SAMPLE_DATE = '" & Mid$(strKeyno, 1, 8) & "'" & _
             "   and SAMPLE_SEQ  = " & Format(Mid$(strKeyno, 9, 3), "##0") & "" & _
             "   and PART        = '" & Mid$(strKeyno, 12, 2) & "'"
    adoRS1.CursorLocation = adUseClient
    adoRS1.Open sqlDoc, AdoCn_SQL
    If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
    
    sqlDoc = "select TESTCD_EQP, TESTCD, REMARK, AUTOVERIFY from INTERFACE002 where (EQP_CD = " & STS(INS_CODE) & ") AND (TESTCD <> '')"
    adoRS2.CursorLocation = adUseClient
    adoRS2.Open sqlDoc, AdoCn_Jet
    If adoRS2.RecordCount > 0 Then adoRS2.MoveFirst
    Do While Not adoRS2.EOF
        If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
        adoRS1.Find "ORD_CODE = " & STS(Trim(adoRS2("TESTCD") & ""))
        If Not adoRS1.EOF Then
            Select Case Trim(adoRS2(2) & "")
                Case "128": strSpec = "PL"
                Case Else:  strSpec = "SE"
            End Select
            
            If Trim(adoRS2("TESTCD_EQP") & "") = "XXX" Then
                strOrder = strOrder + "06A ," + Trim$(adoRS2("AUTOVERIFY") & "") + ",": strPcFlag = "PC"
            Else
                strOrder = strOrder + Trim(adoRS2("TESTCD_EQP") & "") + " ," + Trim$(adoRS2("AUTOVERIFY") & "") + ","
            End If
            intOrdCnt = intOrdCnt + 1
        End If
        adoRS2.MoveNext
    Loop
    adoRS2.Close:   Set adoRS2 = Nothing
    adoRS1.Close:   Set adoRS1 = Nothing
    
    If strOrder <> "" Then strOrder = Mid$(strOrder, 1, Len(strOrder) - 1)
    
End Sub

Private Sub f_subGet_WorkList(ByRef strOrder As String, ByRef intOrdCnt As Integer, _
                              ByRef strSpec As String, ByRef strPcFlag As String, _
                              ByVal intRow As Integer)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String

    Dim varTmp  As Variant
    Dim intCol  As Integer
    
    Dim itemX   As ListItems
    
    Set itemX = lvwCuData.ListItems
    
    strOrder = "":  strPcFlag = "  ": strSpec = "SE":   intOrdCnt = 0
    With spdWorkList
        For intCol = 5 To .MaxCols
            .Row = intRow:  .Col = intCol
            If .BackColor = &HC6FEFF Then
                Select Case itemX.Item(intCol - 4).SubItems(11)
                    Case "128": strSpec = "PL"
                    Case Else:  strSpec = "SE"
                End Select
                .GetText intCol, 0, varTmp
                
                If itemX.Item(intCol - 4).tag = "XXX" Then
                    strOrder = strOrder + "06A ," + itemX.Item(intCol - 4).SubItems(10) + ",": strPcFlag = "PC"
                Else
                    strOrder = strOrder + itemX.Item(intCol - 4).tag + " ," & itemX.Item(intCol - 4).SubItems(10) + ","
                End If
                intOrdCnt = intOrdCnt + 1
            End If
        Next
    End With
    
    If strOrder <> "" Then strOrder = Mid$(strOrder, 1, Len(strOrder) - 1)
   
End Sub

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
        End With
        .HideColumnHeaders = False
    End With
    
   
End Sub

Private Sub f_subSet_ItemComplete(lvw As Listview)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim itemH           As ColumnHeader
    Dim objHeadeItem    As clsCommon
    
    Dim intCol  As Integer
    
    lvw.ColumnHeaders.Clear
    Call lvw.ColumnHeaders.Add(, "EQP_ID", "검체 번호")
    
    intCol = 4
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) AS TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = '" & INS_CODE & "') AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        With lvw
            .Enabled = True
            Set itemH = .ColumnHeaders.Add
            With itemH
                '컬럽 헤더키를 장비검사 코드로
                .Key = COL_KEY & Trim(adoRS.Fields("TESTCD_EQP") & "")
                '컬럽명은 검사 항목 이름
                .Text = Trim(adoRS.Fields("TESTNM") & "")
                '테그는 검사 코드로
                .tag = Trim(adoRS.Fields("TESTCD") & "")
                .Width = 700
                .Alignment = lvwColumnCenter
            End With
            Set itemH = Nothing
        End With
        
        With spdWorkList
            intCol = intCol + 1
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1:  .ColWidth(.MaxCols) = 6.5
            
            .SetText intCol, 0, adoRS.Fields("TESTNM")
        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Function f_subSet_WorkList(ByVal strDate As String, ByVal strDate1 As String)
    Dim sqlRet      As Integer
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
    Set AdoRs_ORACLE = New ADODB.Recordset
'    gSql = "select a.RE_RCID,b.JU_NAME,b.JU_PERID from EXAM_TOC A,JUMN_TMA B,RECE_TJU C" _
'            & " Where c.RE_DATE = '" & strDate & "' And a.IN_CODE like 'HE%' And a.EX_INST < '2' And a.RE_RCID = c.RE_RCID And b.JU_PERID = c.JU_PERID"

    gSql = "select a.RE_RCID,b.JU_NAME,b.JU_PERID from EXAM_TOC A,JUMN_TMA B,RECE_TJU C" _
            & " Where c.RE_DATE >= '" & strDate & "' And c.RE_DATE <= '" & strDate1 & "' " _
            & " And a.IN_CODE IN ('SE00125') And a.EX_INST = '1' " _
            & " And a.RE_RCID = c.RE_RCID And b.JU_PERID = c.JU_PERID order by c.RE_DATE, c.RE_BUSU, a.RE_RCID"
            
    AdoRs_ORACLE.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
   
    If AdoRs_ORACLE.RecordCount = 0 Then
        Set f_subSet_WorkList = Nothing
        RecordChk = False
    Else
        Set f_subSet_WorkList = AdoRs_ORACLE
        RecordChk = True
    End If

    Set AdoRs_ORACLE = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)
    
End Function

Private Function f_subSet_ResultList(ByVal strDate As String, ByVal strDate1 As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_ResultList() As ADODB.Recordset"
    
    Set AdoRs_ORACLE = New ADODB.Recordset

    If optJong = True Then
        If D0COM_CENTERCOD = "10" Then
            sqlDoc = "         SELECT center_code, name, resdnt, health_num, sample_num"
            sqlDoc = sqlDoc & "     , interface_day"
            sqlDoc = sqlDoc & "  FROM interfaceTB "
            sqlDoc = sqlDoc & " WHERE interface_day = " + Chr(39) + strDate + Chr(39)
            sqlDoc = sqlDoc & "   AND center_code = " + Chr(39) + D0COM_CENTERCOD + Chr(39)
            sqlDoc = sqlDoc & "   AND NOT SUBSTR(health_num, 1,1) = " + Chr(39) + "7" + Chr(39)
            
            sqlDoc = sqlDoc & "   AND NOT sample_num = 0"
            sqlDoc = sqlDoc & " ORDER BY sample_num ASC"
        Else
            '각 총국에 종합건강진단 검사
            sqlDoc = "         SELECT center_code, name, resdnt, health_num, serial_num"
            sqlDoc = sqlDoc & "     , health_day"
            sqlDoc = sqlDoc & "  FROM interfaceTB "
            sqlDoc = sqlDoc & " WHERE health_day = " + Chr(39) + strDate + Chr(39)
            sqlDoc = sqlDoc & "   AND center_code = " + Chr(39) + D0COM_CENTERCOD + Chr(39)
            sqlDoc = sqlDoc & "   AND NOT SUBSTR(health_num, 1,1) = " + Chr(39) + "7" + Chr(39)
            
            sqlDoc = sqlDoc & " ORDER BY serial_num ASC"
        End If
    ElseIf optGae = True Then
        If D0COM_CENTERCOD = "10" Then
            sqlDoc = "SELECT DISTINCT center_code, name, resdnt, health_num, sample_num"
            sqlDoc = sqlDoc & "     , interface_day"
            sqlDoc = sqlDoc & "  FROM interfaceTB "
            sqlDoc = sqlDoc & " WHERE interface_day = " + Chr(39) + strDate + Chr(39)
            '-- Query 추가 본부일경우 본부/남대문/인천데이터만 조회(종양표시자/갑상선의 경우..)
            sqlDoc = sqlDoc & "   AND center_code in (" + Chr(39) + "10" + Chr(39) + "," + Chr(39) + "12" + Chr(39) + "," + Chr(39) + "14" + Chr(39) + ")"
            sqlDoc = sqlDoc & "   AND SUBSTR(health_num, 1,1) = " + Chr(39) + "7" + Chr(39)
            sqlDoc = sqlDoc & "   AND NOT sample_num = 0"
            sqlDoc = sqlDoc & " UNION ALL "
                
            sqlDoc = sqlDoc & "SELECT DISTINCT center_code, name, resdnt, health_num, sample_num"
            sqlDoc = sqlDoc & "     , interface_day"
            sqlDoc = sqlDoc & "  FROM interfaceTB "
            sqlDoc = sqlDoc & " WHERE interface_day = " + Chr(39) + strDate + Chr(39)
            '-- Query 추가 본부일경우 본부/남대문/인천데이터만 조회(종양표시자/갑상선의 경우..)
            sqlDoc = sqlDoc & "   AND NOT center_code = " + Chr(39) + "10" + Chr(39)
            sqlDoc = sqlDoc & "   AND NOT SUBSTR(health_num, 1,1) = " + Chr(39) + "7" + Chr(39)
            sqlDoc = sqlDoc & "   AND not center_code in (" + Chr(39) + "15" + Chr(39) + "," + Chr(39) + "16" + Chr(39) + "," + Chr(39) + "17" + Chr(39) + "," + Chr(39) + "18" + Chr(39) + ")"
            sqlDoc = sqlDoc & "   AND NOT sample_num = 0"
            sqlDoc = sqlDoc & " ORDER BY sample_num ASC"
        Else
            '각 총국 계약자서비스
            sqlDoc = "         SELECT DISTINCT center_code, name, resdnt, health_num, serial_num"
            sqlDoc = sqlDoc & "     , health_day"
            sqlDoc = sqlDoc & "  FROM interfaceTB "
            sqlDoc = sqlDoc & " WHERE health_day = " + Chr(39) + strDate + Chr(39)
            sqlDoc = sqlDoc & "   AND center_code = " + Chr(39) + D0COM_CENTERCOD + Chr(39)
            sqlDoc = sqlDoc & "   AND SUBSTR(health_num, 1,1) = " + Chr(39) + "7" + Chr(39)
            sqlDoc = sqlDoc & " ORDER BY serial_num ASC"
        End If
    End If
    
    AdoRs_ORACLE.Open sqlDoc, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
   
    If AdoRs_ORACLE.EOF = True Then
        Set f_subSet_ResultList = Nothing
        RecordChk = False
    Else
        Set f_subSet_ResultList = AdoRs_ORACLE
        RecordChk = True
    End If

    Set AdoRs_ORACLE = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)
    
End Function

Private Function f_subSet_WorkList_Barcode(ByVal strBarno As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
   
        Set AdoRs_SQL = New ADODB.Recordset
                       
        If Len(strBarno) > 8 Then
            sqlDoc = " SELECT a.per_gumjin_date, a.per_gum_num, a.edpscode, a.result, a.send_date, a.per_name " & _
                    " FROM mdck..gumjin_interface a, mdck..bag_interfacecode b " & _
                    " WHERE substring(a.per_gumjin_date,3,8) = '" & Mid(strBarno, 1, 6) & "'" & _
                    " AND a.per_gum_num = '" & Val(Mid(strBarno, 7)) & "' " & _
                    " AND a.result = '' " & _
                    " AND substring(b.kind,1,1) = 'C' " & _
                    " AND a.edpscode=b.meditem " & _
                    " ORDER BY a.per_gumjin_date, a.per_gum_num "
        Else
            sqlDoc = " SELECT a.EnterDate, b.Status, b.waitseqno, b.MAP2SEQNO, b.DispDesc, b.RVALUEKIND, b.NORMLOW, b.NORMHIGH, b.NORMALVALUE, b.RVALUEKIND , " & _
                    " a.ChartNo, b.GumsaKind, c.sujinname, b.status " & _
                    " FROM medicom..WaitPrsnp a, medicom..jun370_resulttb b, medicom..pewprsnp c, medicom..BAGMAP2PREF d " & _
                    " WHERE a.Chartno = '" & strBarno & "' " & _
                    " AND a.WaitSeqNo = b.WaitSeqNo " & _
                    " AND a.status = '1' " & _
                    " AND d.labno = 4 " & _
                    " AND b.jun370no = d.map2seqno " & _
                    " AND b.status = '0' " & _
                    " AND a.chartno = c.chartno " & _
                    " ORDER BY a.chartno "
        End If
        
        Set AdoRs_SQL = New ADODB.Recordset
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
        
        If AdoRs_SQL.RecordCount = 0 Then
            Set f_subSet_WorkList_Barcode = Nothing
            RecordChk = False
            Set AdoRs_SQL = Nothing
            Exit Function
        Else
            Set f_subSet_WorkList_Barcode = AdoRs_SQL
            RecordChk = True
        End If
    
        Set AdoRs_SQL = Nothing
    
Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    
    Call ErrMsgProc(CallForm)

    
End Function

Private Function f_subSet_SearchList(ByVal strBarcode As String)
    Dim sqlRet      As Integer
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
    
    Set AdoRs_SQL = New ADODB.Recordset
'    gSql = "select IN_CODE from EXAM_TOC Where RE_RCID = '" & strBarcode & "'"
    
    gSql = "select a.IN_CODE,a.RE_RCID,b.JU_NAME,b.JU_PERID from EXAM_TOC A,JUMN_TMA B,RECE_TJU C" _
            & " Where a.RE_RCID = '" & strBarcode & "'" _
            & " And b.HE_UNID = 'HC-46101'" _
            & " And a.RE_RCID = c.RE_RCID" _
            & " And b.JU_PERID = c.JU_PERID " _
            & " And a.EX_INST = '2'" _
            & " And a.IN_CODE like 'SE%'"

    AdoRs_SQL.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
    
    If AdoRs_SQL.RecordCount = 0 Then
        Set f_subSet_SearchList = Nothing
    Else
        Set f_subSet_SearchList = AdoRs_SQL
    End If

    Set AdoRs_SQL = Nothing

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
    Dim intCol3  As Integer
'On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = ""
    
    intCol = 5
    intCol3 = 9
    intCol2 = 1
    intRow = 1
    With spdWorkList
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 0
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 0
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdResult2
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 0
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
'             " order by TESTCD_EQP, TESTCD"
             
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst: ReDim fChannel(adoRS.RecordCount)
    Do While Not adoRS.EOF
        Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TEST_EQP") & ""), , "LST")
            itemX.SubItems(1) = Trim(adoRS.Fields("TESTCD") & "")
            itemX.SubItems(2) = Trim(adoRS.Fields("TESTNM") & "")
            itemX.SubItems(3) = ""
            itemX.SubItems(4) = Trim(adoRS.Fields("DELTA") & "")
            itemX.SubItems(5) = Trim(adoRS.Fields("DELTAGBN") & "")
            itemX.SubItems(6) = Trim(adoRS.Fields("PANICL") & "")
            itemX.SubItems(7) = Trim(adoRS.Fields("PANICH") & "")
            itemX.SubItems(8) = Trim(adoRS.Fields("REFL") & "")
            itemX.SubItems(9) = Trim(adoRS.Fields("REFH") & "")
            itemX.SubItems(10) = Trim(adoRS.Fields("AUTOVERIFY") & "")
            itemX.SubItems(11) = Trim(adoRS.Fields("REMARK") & "")
            itemX.tag = Trim(adoRS.Fields("TEST_EQP") & "")
            itemX.Text = Trim(adoRS.Fields("TESTCD") & "")
        Set itemX = Nothing
        
        With spdWorkList
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
            .Col = intCol:  .ColHidden = True
        End With
        
        With spdResult1
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
                .ColWidth(intCol) = 14.5
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
                .ColWidth(intCol) = 14.5
            End If
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        fChannel(intCol - 5) = adoRS.Fields("TEST_EQP")
        
        intCnt = intCnt + 1
        ReDim Preserve f_typCode(1 To intCnt) As TYPE_CD
        
        f_typCode(intCnt).strEqpCd = Trim$(adoRS.Fields("TEST_EQP"))
        f_typCode(intCnt).intCnt = 0
        
        strTmp = Trim$(adoRS.Fields("TESTCD"))
        intPos = InStr(strTmp, ",")
        Do While intPos > 0
            f_strOrdList = f_strOrdList + "'" + Mid$(strTmp, 1, intPos - 1) + "',"
            
            f_typCode(intCnt).intCnt = f_typCode(intCnt).intCnt + 1
            f_typCode(intCnt).strTestCd(f_typCode(intCnt).intCnt) = Mid$(strTmp, 1, intPos - 1)
            
            strTmp = Mid$(strTmp, intPos + 1)
            
            intPos = InStr(strTmp, ",")
        Loop
        f_strOrdList = f_strOrdList + "'" + strTmp + "',"
        f_typCode(intCnt).intCnt = f_typCode(intCnt).intCnt + 1
        f_typCode(intCnt).strTestCd(f_typCode(intCnt).intCnt) = strTmp
        
        intCol = intCol + 1
        
        adoRS.MoveNext
    Loop
    Set adoRS = Nothing
    
    With spdResult2
        If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
        .SetText intCol, 0, ""
        .Col = intCol:  .ColHidden = True
    End With

Exit Sub
ErrRoutine:
    Set adoRS = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub
Private Function f_funGet_CODE(ByVal strOrdcd As String) As String

    Dim intIdx1 As Integer, intIdx2 As Integer
    
    f_funGet_CODE = ""
    
    For intIdx1 = 1 To UBound(f_typCode)
        For intIdx2 = 1 To f_typCode(intIdx1).intCnt
            If Trim$(strOrdcd) = Trim$(f_typCode(intIdx1).strTestCd(intIdx2)) Then
                f_funGet_CODE = f_typCode(intIdx1).strEqpCd
                Exit Function
            End If
        Next
    Next
    
End Function

Private Sub cmdEot_Click()
    comEQP.Output = EOT
End Sub

Private Sub cmdOrder_Click()
'Dim ii As Integer
'Dim chkRackNo As Variant
'Dim strMsg As String
'
'    spdResult1.GetText 5, 1, chkRackNo
'    strMsg = "오더전송 준비가 되었습니다." & vbCrLf & chkRackNo & "을 사용하시겠습니까?"
'    If vbYes = MsgBox(strMsg, vbCritical + vbYesNo) Then
'        OrderCnt = 0
'
'        With spdResult1
'            For ii = 1 To .maxrows
'                .Col = 1: .Row = ii
'                If .Value = 1 Then
'                    .Col = 2
'                    If Len(Trim(.Text)) > 0 And .BackColor = vbWhite Then
'                        comEQP.Output = ENQ
'                        Debug.Print "[HOST] " & ENQ
'                        txtResult.Text = txtResult.Text + "[HOST] " & ENQ
'                        lblStatus = "오더전송중.."
'                        .BackColor = vbCyan
'                        OrderCnt = ii
'                        .Col = 3
'                        .BackColor = vbCyan
'                        .Col = 4
'                        .BackColor = vbCyan
'                        .Col = 5
'                        .BackColor = vbCyan
'                        .Col = 6
'                        .BackColor = vbCyan
'                        Exit For
'                    End If
'                End If
'            Next
'        End With
'    Else
'        Exit Sub
'    End If
        
    Dim iRow  As Integer
    Dim icol  As Integer
    Dim sTxt  As String
    Dim sName As String
    Dim varTmp  As Variant
    Dim strSeq As String
        
   Open ImportPath & "\" & Format(Now, "yymmdd") & "ipt.txt" For Output As #2
    
    sTxt = "Patient ID,"
    For icol = 5 To spdResult1.MaxCols
        If icol = spdResult1.MaxCols Then
            sTxt = sTxt + "Test name"
        Else
            sTxt = sTxt + "Test name,"
        End If
    Next
    
    Print #2, sTxt + Chr(13) & Chr(10);
    With spdResult1
        For iRow = 1 To .maxrows
            sTxt = ""
            strSeq = iRow
            .Row = iRow: .Col = 1
            If .Value <> 1 Then GoTo rst2
            For icol = 3 To .MaxCols
                .Col = icol
                .GetText icol, 0, varTmp
                If icol = 4 Then GoTo rst1
                If icol < 4 Then
                    .Col = 4
                    sTxt = sTxt + Format(strSeq, "00") + ","
                Else
                    If Trim$(varTmp) <> "" Then
                        sTxt = sTxt + Trim$(varTmp) + ","
                    Else
                        sTxt = sTxt + ","
                    End If
                End If
rst1:
            Next icol
            If Right(sTxt, 1) = "," Then
                sTxt = Mid(sTxt, 1, Len(sTxt) - 1)
            End If
            
            Print #2, sTxt + Chr(13) & Chr(10);
rst2:
        Next iRow
    End With

    Close #2
    MsgBox "업무나열서가 생성되었습니다", vbInformation, Me.Caption
    
'    Call cmdInitial_Click
    Exit Sub

ErrLst:
    MsgBox Err.Number & ": " & Err.Description, vbCritical, Me.Caption
        
End Sub

'Private Sub cmdRackNo_Click()
'Dim sNo As String, sCnt As Integer, sAdd As Integer
'Dim fNum1 As Integer, fNum2 As Integer
'Dim intRow1 As Integer
'
'AgainInput:
'    fNum1 = 1: fNum2 = 0
'    sNo = InputBox("시작 번호를 입력하세요 !")
'    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
'        If Not IsNumeric(sNo) Then
'            MsgBox "숫자만 입력하세요.!", vbCritical
'            GoTo AgainInput
'        End If
'
'        With spdResult1
'            sAdd = 0
'            For sCnt = .ActiveRow To .maxrows
'                intRow1 = intRow1 + 1
'                .Row = sCnt
'                .Col = 1
'                If .Value >= 1 Then
'                    'If .ActiveCol = 3 Then
'                        .Col = 5 '.ActiveCol
'                        If intRow1 = (5 * fNum1) + 1 Then fNum1 = fNum1 + 1: fNum2 = 0
'                        fNum2 = fNum2 + 1
'                        .Text = Format(Trim((fNum1 + Val(sNo)) - 1), "00000")
'                        .Col = 6 '.ActiveCol + 1
'                        .Text = fNum2
'                    'End If
'                End If
'            Next sCnt
'        End With
'    End If
'End Sub

Private Sub cmdResultQuery_Click()
    FileBep2000.Path = ExportPath
    If FileBep2000.Visible = True Then
        FileBep2000.Visible = False
    Else
        FileBep2000.Visible = True
    End If
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Privete sub cmdWorkQuery_Click()"
    
    Dim strKeyno    As String
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String, strBarno()   As String
    Dim strTestCd() As String, strTPid()    As String, strTPnm() As String
    Dim strEqpCd    As String
    Dim intRow  As String, intIdx  As Integer, intCol   As Integer
    Dim itemX   As ListItem
    Dim strOld  As Integer
       
    '-- WorkList조회
    Set mAdoRs = f_subSet_WorkList(mskOrdDate.Text, mskOrdDate1.Text)
    If RecordChk = False Then
            MsgBox "조회 된 대상자가 없습니다." & vbCrLf & "검진일자를 확인하세요.", vbInformation, Me.Caption
        Exit Sub
    End If
    With spdWorkList
        .maxrows = 0
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    intRow = 0
    Do Until mAdoRs.EOF
        intIdx = 0
        With spdWorkList
            If strKeyno <> mAdoRs.Fields("RE_RCID") Then
                intRow = intRow + 1
                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                
                .SetText 1, intRow, 1
                .SetText 2, intRow, mAdoRs("JU_NAME")
                .SetText 3, intRow, mAdoRs("JU_PERID")
                .SetText 4, intRow, mAdoRs("RE_RCID")
            '-- 검사항목조회
                Set mAdoRs1 = New Recordset
                Set mAdoRs1 = f_subSet_TestList(mAdoRs("RE_RCID"))
                
                Do Until mAdoRs1.EOF
                    strEqpCd = f_funGet_CODE(mAdoRs1("IN_CODE"))
                    
                    Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                    If Not itemX Is Nothing Then .SetText 7 + itemX.Index, intRow, "V"
                    Set itemX = Nothing
                    mAdoRs1.MoveNext
                Loop
            End If
            strKeyno = mAdoRs("RE_RCID")
        End With
        intIdx = intIdx + 1
        mAdoRs.MoveNext
    Loop
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub cmdACK_Click()

    comEQP.Output = ACK

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
    
    f_strJOB_FLAG = "1"
    f_intSampleNo = 0
    Or_Seq = 1
    
    With spdWorkList
        .maxrows = 0
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdResult1
        .maxrows = 0
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
        .RowHeight(-1) = 12
    End With

    With spdResult2
        .maxrows = 0
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
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
   
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
        
    Dim varTmp  As Variant, strErrMsg   As String
    Dim strSampleno()   As String, strBarno     As String, strTime      As String
    Dim strOrdcd()      As String, strRstval()  As String, intCnt       As Integer
    Dim strTmp1()       As String, strTmp2()    As String
    Dim intPos          As String, strTestCd    As String, strTestRst   As String
    Dim strName         As String
    Dim strChartNo      As String
    Dim strHealth       As String
    Dim strOrdLst()     As String, strPid()    As String, strPnm() As String
    
    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
    Dim itemX   As ListItem
    Dim objSpd  As vaSpread
    Dim sqlRet  As Integer
    Dim flgSave As Boolean
    Dim strResult   As String
    Dim WK_SLEKWA   As String
    Dim WK_SJKEY    As String
    
    CallForm = "frmComm - Private Sub cmdAppend_Click()"
    
'On Error GoTo ErrorRoutine
    
    Me.MousePointer = 11
    
    If Index = 0 Then
        Set objSpd = spdResult1
    Else
        Set objSpd = spdResult2
    End If
    
    With objSpd
        For intRow = 1 To .maxrows
            .GetText 2, intRow, varTmp:         strBarno = Trim$(varTmp)
            .GetText 3, intRow, varTmp:         strName = Trim$(varTmp)
            .GetText 4, intRow, varTmp:         strChartNo = Trim$(varTmp)
            
            .GetText 1, intRow, varTmp
            
            If strChartNo = "" Then Exit For
            
            intCnt = 0: Erase strOrdcd: Erase strRstval
            If Trim$(varTmp) = "1" Then
                For intCol = 5 To .MaxCols
                    .GetText intCol, intRow, varTmp
                    If Trim$(varTmp) <> "" Then
                        .GetText intCol, 0, varTmp
                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                        If Not itemX Is Nothing Then
                            .GetText intCol, intRow, varTmp:    strResult = Trim(varTmp)
                            strTestCd = itemX.ListSubItems(1)
                            
                            intPos = InStr(strTestCd, ",")
                            If intPos > 0 Then
                                Exit Sub
                            Else
                                sqlDoc = "update EXAM_TOC set EX_INRV = '" & Trim(strResult) & "',EX_INST = '2',EX_DATE = '" & Format$(Now, "YYYYMMDD") & "',EX_INEM='" & Trim(CurrUser.CuUserID) & "'" _
                                       & " where RE_RCID ='" & strChartNo & "' And IN_CODE='" & itemX.Text & "'"

                                AdoCn_ORACLE.Execute (sqlDoc)
                                lblStatus.Caption = "저장 성공!!"
                                AdoCn_ORACLE.Execute sqlDoc
                                'Debug.Print sqlDoc
                            End If
                        End If
                                                
                        Set itemX = Nothing
                    End If
                Next
                
                spdResult1.Row = intRow
                spdResult1.Col = 1
                spdResult1.BackColor = vbWhite
                spdResult1.Row = intRow
                spdResult1.Col = 2
                spdResult1.BackColor = vbWhite
                spdResult1.Col = 3
                spdResult1.BackColor = vbWhite
                spdResult1.Col = 4

                spdResult1.Col = 1: spdResult1.Value = 0

                If strErrMsg = "" Then
                    sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
                             " where SPCNO   = '" & strChartNo & "'" & _
                             "   and TRANSDT = '" & mskRstDate.Text & "'"
                    AdoCn_Jet.Execute sqlDoc
                Else
                    MsgBox strErrMsg, vbInformation, Me.Caption
                End If
            End If
        Next
    End With
    Me.MousePointer = 0
    MsgBox "작업이 완료되었습니다.", vbInformation, Me.Caption
 
End Sub

Private Function f_delta_chk(ByVal WK_SAMPLE As String, ByVal WK_WORKNM As String, ByVal WK_VALUE As String)
Dim WK_CLVALUE As String, WK_CHVALUE As String, WK_MLVALUE As String, WK_MHVALUE As String
Dim WK_FLVALUE As String, WK_FHVALUE As String, WK_DLVALUE As String, WK_DHVALUE As String, WK_PLVALUE As String, WK_PHVALUE As String
Dim WK_METHOD As String, S_GMSEX As String
Dim WK_STAND As Integer

On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_BarcodeSp() As ADODB.Recordset"

    If Mid(WK_VALUE, 1, 1) = "p" Or Mid(WK_VALUE, 1, 1) = "n" Then
        WK_VALUE = ""
    End If

    Set AdoRs_ORACLE = New ADODB.Recordset

    gSql = "SELECT G13_GMSEX " _
         & "  From GUMSA013 " _
         & " WHERE GUMSA013.G13_SAMPLE = '" & WK_SAMPLE & "' "

    AdoRs_ORACLE.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly

    If AdoRs_ORACLE.RecordCount = 0 Then
    Else
        S_GMSEX = AdoRs_ORACLE("G13_GMSEX")
    End If


'    If AdoRs_ORACLE("WK_STAND") = 0 Then
'        MsgBox "작업 품목인  " + WK_WORKNM + " 의 정상치값이 입력되어 있지 않습니다.!!!", "'F_DELTA_CHECK MESSAGE!'", vbExclamation
'        Exit Function
'    End If
    AdoRs_ORACLE.Close

    
    gSql = "SELECT NVL(GUMSA010.G10_CLVALU,-99999999) WK_CLVALUE,NVL(GUMSA010.G10_CHVALU,-99999999) WK_CHVALUE, " _
         & "       NVL(GUMSA010.G10_MLVALU,-99999999) WK_MLVALUE,NVL(GUMSA010.G10_MHVALU,-99999999) WK_MHVALUE, " _
         & "       NVL(GUMSA010.G10_FLVALU,-99999999) WK_FLVALUE,NVL(GUMSA010.G10_FHVALU,-99999999) WK_FHVALUE, " _
         & "       NVL(GUMSA010.G10_DLVALU,-99999999) WK_DLVALUE,NVL(GUMSA010.G10_DHVALU,-99999999) WK_DHVALUE, " _
         & "       NVL(GUMSA010.G10_PLVALU,-99999999) WK_PLVALUE,NVL(GUMSA010.G10_PHVALU,-99999999) WK_PHVALUE, " _
         & "       NVL(GUMSA010.G10_METHOD ,' ') WK_METHOD" _
         & "  From GUMSA010" _
         & " WHERE GUMSA010.G10_GMPART = '13' " _
         & "   AND GUMSA010.G10_WORKNM = '" & WK_WORKNM & "'  "
    
    AdoRs_ORACLE.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
    
    If AdoRs_ORACLE.RecordCount = 0 Then
'        MsgBox "검사 기준값을 정의한 TABLE을 읽지 못했습니다..!!!", vbExclamation, "CAUTION!"
        Exit Function
    End If
    
    If AdoRs_ORACLE.EOF Then
    Else
        WK_CLVALUE = AdoRs_ORACLE("WK_CLVALUE"):    WK_CHVALUE = AdoRs_ORACLE("WK_CHVALUE")
        WK_MLVALUE = AdoRs_ORACLE("WK_MLVALUE"):    WK_MHVALUE = AdoRs_ORACLE("WK_MHVALUE")
        WK_FLVALUE = AdoRs_ORACLE("WK_FLVALUE"):    WK_FHVALUE = AdoRs_ORACLE("WK_FHVALUE")
        WK_DLVALUE = AdoRs_ORACLE("WK_DLVALUE"):    WK_DHVALUE = AdoRs_ORACLE("WK_DHVALUE")
        WK_PLVALUE = AdoRs_ORACLE("WK_PLVALUE"):    WK_PHVALUE = AdoRs_ORACLE("WK_PHVALUE")
        WK_METHOD = AdoRs_ORACLE("WK_METHOD")
        
        '----------------------------------------
        '-- COMMON EVALUATE CHECKING
        '----------------------------------------
        If WK_CLVALUE > -99999999 Then
            G_CLVALU = WK_CLVALUE
            G_CHVALU = WK_CHVALUE
            
            If WK_VALUE = "" Then
                G_EVALUATE = ""
            Else
                If WK_VALUE < WK_CLVALUE Then
                    G_EVALUATE = "L"
                ElseIf WK_VALUE > WK_CHVALUE Then
                    G_EVALUATE = "H"
                Else
                    G_EVALUATE = ""
                End If
            End If
        End If
        
        '----------------------------------------
        '-- MAN EVALUATE CHECKING
        '----------------------------------------
        If WK_MLVALUE > -99999999 And S_GMSEX = "M" Then
            G_CLVALU = WK_MLVALUE
            G_CHVALU = WK_MHVALUE
            
            If WK_VALUE = "" Then
                G_EVALUATE = ""
            Else
                If WK_VALUE < WK_MLVALUE Then
                    G_EVALUATE = "L"
                ElseIf WK_VALUE > WK_MHVALUE Then
                    G_EVALUATE = "H"
                Else
                    G_EVALUATE = ""
                End If
            End If
        End If
        
        '----------------------------------------
        '-- FEMALE EVALUATE CHECKING
        '----------------------------------------
        If WK_FLVALUE > -99999999 And S_GMSEX = "F" Then
            G_CLVALU = WK_FLVALUE
            G_CHVALU = WK_FHVALUE
            
            If WK_VALUE = "" Then
                G_EVALUATE = ""
            Else
                If WK_VALUE < WK_FLVALUE Then
                    G_EVALUATE = "L"
                ElseIf WK_VALUE > WK_FHVALUE Then
                    G_EVALUATE = "H"
                Else
                    G_EVALUATE = ""
                End If
            End If
        End If
        
        '----------------------------------------
        '-- PANIC CHECKING
        '----------------------------------------
        If WK_PLVALUE > -99999999 And S_GMSEX = "F" Then
            
            If WK_VALUE = "" Then
                G_PANIC = ""
            Else
                If WK_VALUE < WK_PLVALUE Then
                    G_PANIC = "L"
                ElseIf WK_VALUE > WK_PHVALUE Then
                    G_PANIC = "H"
                Else
                    G_PANIC = ""
                End If
            End If
        End If
        
        '----------------------------------------
        '-- DELTA CHECKING
        '----------------------------------------
        If WK_DLVALUE > -99999999 And S_GMSEX = "F" Then
            If WK_VALUE = "" Then
                G_DELTA = ""
            Else
                If WK_VALUE < WK_DLVALUE Then
                    G_DELTA = "L"
                ElseIf WK_VALUE > WK_DHVALUE Then
                    G_DELTA = "H"
                Else
                    G_DELTA = ""
                End If
            End If
        End If
    End If
    
    Set AdoRs_ORACLE = Nothing
    
Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)

End Function

Private Sub F_HEL30203_UPDATE(ByVal WK_SJKEY As String, ByVal WK_WORKNM As String, ByVal WK_VALUE As String)
Dim WK_CNT As Integer
'------------------------------------------------------------
'-- 종합/일반/채용/암환자에 대한 자료 UPDATE (면역혈청 검사)
'------------------------------------------------------------

'-- WK_SJKEY/WK_WORKNM/WK_VALUE

On Error GoTo HEL323_UPDATE

    CallForm = "frmComm - Private Sub F_HEL30203_UPDATE()"

    Set Ado323 = New ADODB.Recordset
    
    gSql = "SELECT COUNT(H33_SJKEY) WK_CNT" _
         & "  From HEL30203 " _
         & " Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    
    Ado323.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
    
    If Ado323.RecordCount = 0 Then
            MsgBox "건강검진 환자에 대한 종합검진 면역혈청 자료를 찾지 못했습니다.!!!", vbExclamation, "HEL30203"
            Return
    Else
        WK_CNT = Ado323("WK_CNT")
    End If
    
    Ado323.Close
    Set Ado323 = Nothing
    
    If WK_CNT < 1 Then
        MsgBox "건강검진 환자에 대한 종합검진 면역혈청 자료를 찾지 못했습니다.!!!", vbExclamation, "HEL30203"
        Return
    End If
    
    If WK_WORKNM = "uTSH" Then                         '-- TSH
        gSql = " Update HEL30203 " _
             & "    Set H33_TSHTSH = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
             
    ElseIf WK_WORKNM = "FreeT4" Then                   '-- FREE T4
        gSql = " Update HEL30203 " _
             & "    Set H33_FREET4 = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
       
    ElseIf WK_WORKNM = "HBeAg 정" Then                 '-- B형 간염E항원(HBe-Ag)
        Select Case UCase(WK_VALUE)
            Case "POSITIVE"
                WK_VALUE = "양성"
            Case "NEGATIVE"
                WK_VALUE = "음성"
        End Select
        gSql = " Update HEL30203 " _
             & "    Set H33_HBEAG = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "HBeAb 정" Then                '-- Anti-HBe  B형 간염E항체(HEe-Ag)
        Select Case UCase(WK_VALUE)
            Case "POSITIVE"
                WK_VALUE = "양성"
            Case "NEGATIVE"
                WK_VALUE = "음성"
        End Select
        
        gSql = " Update HEL30203 " _
             & "    Set H33_ANHBE = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "HBsAg 정" Then           '-- B형 간염항원(HBs-Ag)
        Select Case UCase(WK_VALUE)
            Case "POSITIVE"
                WK_VALUE = "양성"
            Case "NEGATIVE"
                WK_VALUE = "음성"
        End Select
        
        gSql = " Update HEL30203 " _
             & "    Set H33_HBSAG = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
        
    ElseIf WK_WORKNM = "HBsAb 정" Then           '-- B형 간염항체(Anti-HBs)
        Select Case UCase(WK_VALUE)
            Case "POSITIVE"
                WK_VALUE = "양성"
            Case "NEGATIVE"
                WK_VALUE = "음성"
        End Select
        
        gSql = " Update HEL30203 " _
             & "    Set H33_ANHBS = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "HCV Ab 정" Then         '-- C형 간염항체(HCV-Ab)
        Select Case UCase(WK_VALUE)
            Case "POSITIVE"
                WK_VALUE = "양성"
            Case "NEGATIVE"
                WK_VALUE = "음성"
        End Select
        
        gSql = " Update HEL30203 " _
             & "    Set H33_HCVAB = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "VDRL(Quan)" Then          '-- VDRL 매독
        If WK_VALUE = "Non-Reactive" Then
            WK_VALUE = "음성"
        End If
        
        gSql = " Update HEL30203 " _
             & "    Set H33_VDRL = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "VDRL(Qual)" Then          '-- VDRL 매독
        If WK_VALUE = "Non-Reactive" Then
            WK_VALUE = "음성"
        End If
        
        gSql = " Update HEL30203 " _
             & "    Set H33_VDRL = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "AIDS" Then                '-- AIDS
        gSql = " Update HEL30203 " _
             & "    Set H33_AIDS = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "HIV(AIDS)" Then               '-- AIDS
        gSql = " Update HEL30203 " _
             & "    Set H33_AIDS = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "AFP(정밀)" Then            '-- AFP
        gSql = " Update HEL30203 " _
             & "    Set H33_AFPAFP = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "CEA" Then                  '-- CEA
        gSql = " Update HEL30203 " _
             & "    Set H33_CEACEA = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "PSA" Then                  '-- PSA
        gSql = " Update HEL30203 " _
             & "    Set H33_PSA = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "ASO(Quan)" Then              '-- ASO
        If WK_VALUE >= "220" Then
           WK_VALUE = "양성"
        Else
           WK_VALUE = "음성"
        End If
        gSql = " Update HEL30203 " _
             & "    Set H33_ASO = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "RA(Quan)" Then             '-- RA-FACTOR
        If WK_VALUE >= "20" Then
           WK_VALUE = "양성"
        Else
           WK_VALUE = "음성"
        End If
        gSql = " Update HEL30203 " _
             & "    Set H33_RAFACT = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "CRP(Quan)" Then         '-- CRP
        If WK_VALUE >= "0.5" Then
           WK_VALUE = "양성"
        Else
           WK_VALUE = "음성"
        End If
        gSql = " Update HEL30203 " _
             & "    Set H33_CRP = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "CA125" Then         '-- CA125
        gSql = " Update HEL30203 " _
             & "    Set H33_CA125 = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "CA19-9" Then            '-- CA19-9
        gSql = " Update HEL30203 " _
             & "    Set H33_CA199 = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
    ElseIf WK_WORKNM = "THCY" Then          '-- homocysteine
        gSql = " Update HEL30203 " _
             & "    Set H33_HOMO = '" & WK_VALUE & "' " _
             & "  Where HEL30203.H33_SJKEY = '" & WK_SJKEY & "' "
     End If

    AdoCn_ORACLE.Execute (gSql)
        
HEL323_UPDATE:
    Set Ado323 = Nothing
    Call ErrMsgProc(CallForm)

End Sub

Public Function CheckSum_ECi_Tx(ByVal strPrmValue As String)

    Dim i                   As Integer
    Dim intValueLength      As Integer
    Dim intCheck            As Integer
    Dim strCheck            As String
    
    intCheck = 0
    
    intValueLength = LenA(strPrmValue)
    
    For i = 1 To intValueLength
        intCheck = intCheck + Asc(Mid(strPrmValue, i, 1))
    Next
    
    strCheck = Hex(intCheck)
    
    If Len(strCheck) = 1 Then
        CheckSum_ECi_Tx = "0" & strCheck
    Else
        CheckSum_ECi_Tx = Right(strCheck, 2)
    End If

End Function

Public Function LenA(strPrmString As String) As Integer

    Dim i                   As Integer
    Dim intStrLen           As Integer
    Dim intAnsiStrLen       As Integer
    Dim strTemp             As String
    
    intStrLen = Len(strPrmString)
    For i = 1 To intStrLen
        strTemp = Mid(strPrmString, i, 1)
        
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
    
    comEQP.Output = ENQ

End Sub

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                               *
'*   생년월일로 나이를 계산                      *
'*   passport_id   :  생년월일 변환대상 data     *
'*                                               *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function D0SUB_BIRTHDAY(ByVal PassPort_Id As String) As String

    Dim yy       As String
    Dim age      As Integer

    On Error GoTo D0SUB_BIRTHDAY
    
    Select Case Len(left$(PassPort_Id, 6))
        Case 2, 3, 4, 5
            yy = left$(PassPort_Id, 2) & "-01-01"
            age = DateDiff("yyyy", yy, Now)

        Case 6
            yy = Format$(PassPort_Id, "##-##-##")
            age = DateDiff("yyyy", yy, Now)
    End Select
        
    D0SUB_BIRTHDAY = Trim$(Str$(age))
    
    On Error GoTo 0
    Exit Function
D0SUB_BIRTHDAY:
    Resume Next
        
End Function

Function D0SUB_SETCENTER(para As String) As Variant

    Select Case Trim$(para)
        Case "10": D0SUB_SETCENTER = " 본사 "
        Case "12": D0SUB_SETCENTER = "남대문"
        Case "14": D0SUB_SETCENTER = " 인천 "
        Case "15": D0SUB_SETCENTER = " 대전 "
        Case "16": D0SUB_SETCENTER = " 광주 "
        Case "17": D0SUB_SETCENTER = " 대구 "
        Case "18": D0SUB_SETCENTER = " 부산 "
        Case "본사": D0SUB_SETCENTER = "10"
        Case "남대문": D0SUB_SETCENTER = "12"
        Case "인천": D0SUB_SETCENTER = "14"
        Case "대전": D0SUB_SETCENTER = "15"
        Case "광주": D0SUB_SETCENTER = "16"
        Case "대구": D0SUB_SETCENTER = "17"
        Case "부산": D0SUB_SETCENTER = "18"
    End Select

End Function

Private Sub cmdRstQuery_Click()

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String, intRet   As Integer
    
    Dim strSpcno    As String
    Dim intRow      As Integer, intCol  As Integer
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
    
    Dim itemX       As ListItem

    intRow = 0
    With spdResult2
        .maxrows = 25
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    sqlDoc = "select SPCNO, TESTCD, EQUIPCD, TRANSDT, RSTVAL, REFVAL, TRANSDT, EQPNUM, NAME, PNO" & _
             "  from INTERFACE003" & _
             " where TRANSDT >= '" & mskRstDate.Text & "'" & _
             "   and EQUIPCD = '" & INS_CODE & "'"
    If cboRstgbn(1).ListIndex = 0 Then
        sqlDoc = sqlDoc & "   and SERVERGBN = ''"
    ElseIf cboRstgbn(1).ListIndex = 1 Then
        sqlDoc = sqlDoc & "   and SERVERGBN = 'Y'"
    End If
    sqlDoc = sqlDoc & " order by SPCNO, TRANSTM"
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        With spdResult2
        If strSpcno <> Trim$(adoRS(0) & "") Then
                intRow = intRow + 1
                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                .SetText 1, intRow, "1"
                .SetText 2, intRow, Trim$(adoRS(9) & "")
                .SetText 3, intRow, Trim$(adoRS(8) & "")
                .SetText 4, intRow, Trim$(adoRS(0) & "")
            End If
                strSpcno = Trim$(adoRS(0) & "")
                Set itemX = lvwCuData.FindItem(Trim$(adoRS(7) & ""), lvwTag, , lvwWhole)
                If Not itemX Is Nothing Then
                    intCol = itemX.Index + 4
                    .SetText intCol, intRow, Trim$(adoRS(4)) & ""
                    .Col = intCol:  .Row = intRow:  .ForeColor = IIf(Trim$(adoRS(5) & "") <> "", vbRed, vbBlack)
                End If
        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub cmdSel_Click(Index As Integer)

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
'    If Index = 2 Or Index = 3 Then
        With spdWorkList
            For intRow = 1 To .maxrows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 0, "1", "")
            Next
        End With
'    Else
'        With spdWorkList
'            For intRow = 1 To .maxrows
'                .GetText 2, intRow, varTmp
'                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 0, "1", "")
'            Next
'        End With
'    End If
    
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
                .Col = 6:       .Text = Trim(sAdd + Val(sNo))
                sAdd = sAdd + 1
            Next sCnt
        
            .StartingRowNumber = Val(sNo)
        End With
    End If

End Sub

Private Function f_subSet_TestList(ByVal strBarcode As String)
    Dim sqlRet      As Integer
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    gSql = "select IN_CODE from EXAM_TOC Where RE_RCID = '" & strBarcode & "' and EX_INST = '1' "
    AdoRs_ORACLE.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
    
    If AdoRs_ORACLE.RecordCount = 0 Then
        Set f_subSet_TestList = Nothing
    Else
        Set f_subSet_TestList = AdoRs_ORACLE
    End If

    Set AdoRs_ORACLE = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)

    
End Function

Private Function f_subSet_TestList1(ByVal strBarcode As String, ByVal strOld As Integer, ByVal strSex As Integer)
    Dim sqlRet      As Integer
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
    
    Set AdoRs_ORACLE = New ADODB.Recordset
   If optJong.Value = True Then
        If strOld >= 40 And strSex = 1 Then
                   gSql = "SELECT * from tb_msmedhsa" & vbLf
            gSql = gSql & " WHERE medi_cust_id = (select medi_cust_id from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
            gSql = gSql & "   AND medi_type_dvsn = (select medi_type_dvsn from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
            gSql = gSql & "   AND chex_objt_sqno = (select chex_objt_sqno from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
            gSql = gSql & "   AND medi_item_code in ('AFA01','AFA02','ABA01','ABA02','AFA06')" & vbLf
        Else
                   gSql = "SELECT * from tb_msmedhsa" & vbLf
            gSql = gSql & " WHERE medi_cust_id = (select medi_cust_id from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
            gSql = gSql & "   AND medi_type_dvsn = (select medi_type_dvsn from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
            gSql = gSql & "   AND chex_objt_sqno = (select chex_objt_sqno from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
            gSql = gSql & "   AND medi_item_code in ('AFA01','AFA02','ABA01','ABA02')" & vbLf
        End If
   Else
        If Mid(strBarcode, 1, 1) = "7" Then
                    gSql = "SELECT * from tb_msmedhsa" & vbLf
             gSql = gSql & " WHERE medi_cust_id = (select medi_cust_id from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
             gSql = gSql & "   AND medi_type_dvsn = (select medi_type_dvsn from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
             gSql = gSql & "   AND chex_objt_sqno = (select chex_objt_sqno from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
             gSql = gSql & "   AND medi_item_code in ('AFA01','ABA02','ACA01')" & vbLf
         ElseIf Mid(strBarcode, 1, 1) = "1" Then
            If strOld >= 40 And strSex = 1 Then
                       gSql = "SELECT * from tb_msmedhsa" & vbLf
                gSql = gSql & " WHERE medi_cust_id = (select medi_cust_id from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
                gSql = gSql & "   AND medi_type_dvsn = (select medi_type_dvsn from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
                gSql = gSql & "   AND chex_objt_sqno = (select chex_objt_sqno from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
                gSql = gSql & "   AND medi_item_code in ('AFA01','AFA02','ABA02','AFA06','ABA01')" & vbLf
            Else
                       gSql = "SELECT * from tb_msmedhsa" & vbLf
                gSql = gSql & " WHERE medi_cust_id = (select medi_cust_id from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
                gSql = gSql & "   AND medi_type_dvsn = (select medi_type_dvsn from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
                gSql = gSql & "   AND chex_objt_sqno = (select chex_objt_sqno from tb_ncmedcmr where intf_chex_no = '" & strBarcode & "')" & vbLf
                gSql = gSql & "   AND medi_item_code in ('AFA01','AFA02','ABA02','ABA01')" & vbLf
            End If
         End If
    End If
    
    AdoRs_ORACLE.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
    
    If AdoRs_ORACLE.EOF = True Then
        Set f_subSet_TestList1 = Nothing
        RecordChk = False
        Exit Function
    Else
        Set f_subSet_TestList1 = AdoRs_ORACLE
        RecordChk = True
    End If

    Set AdoRs_ORACLE = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)

    
End Function
Private Sub cmdWorkList_Click()

    Dim varTmp  As Variant
    Dim intRow1 As Integer, intRow2 As Integer
    Dim intIdx  As Integer
    Dim Rev     As Long
    Dim Test_Cd() As String, strPid()   As String, strPnm() As String
    Dim itemX As ListItem
    Dim blnFlag As Boolean
    Dim strBarno    As String, strSPid  As String, strSPnm   As String
    Dim strSex      As String, strOld   As String, strArea   As String
    Dim strTmpSex   As Integer
    
    Dim strEqpCd    As String
       
    blnFlag = False
    With spdWorkList
        For intRow1 = 1 To .maxrows
            .GetText 1, intRow1, varTmp
            If Trim$(varTmp) = "1" Then
                .GetText 2, intRow1, varTmp:    strSPnm = Trim$(varTmp)
                .GetText 3, intRow1, varTmp:    strSPid = Trim$(varTmp)
                .GetText 4, intRow1, varTmp:    strBarno = Trim$(varTmp)
                
                intRow2 = f_funGet_SpreadRow(spdResult1, 2, strBarno)
                If intRow2 < 1 Then
                    intRow2 = f_funGet_SpreadRow(spdResult1, 2, "")
                    If intRow2 < 1 Then
                        spdResult1.maxrows = spdResult1.maxrows + 1
                        spdResult1.RowHeight(spdResult1.maxrows) = 12
                        intRow2 = spdResult1.maxrows
                    End If
                    
                    blnFlag = False
                    Set mAdoRs = f_subSet_TestList(strBarno)
                    If Len(strBarno) > 0 Then
                        Do Until mAdoRs.EOF
                            strEqpCd = f_funGet_CODE(mAdoRs("IN_CODE"))
                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then
                                blnFlag = True
                                spdResult1.Row = intRow2
                                spdResult1.Col = itemX.Index + 9
                                spdResult1.BackColor = &HC6FEFF '&H80C0FF
                                DoEvents
                            End If
                            mAdoRs.MoveNext
                        Loop
                    End If
                    If blnFlag = True Then
                        spdResult1.SetText 2, intRow2, strSPid
                        spdResult1.SetText 3, intRow2, strSPnm
                        spdResult1.SetText 4, intRow2, strBarno
                    Else
                        spdResult1.maxrows = spdResult1.maxrows - 1
                    End If
                End If
                spdResult1.SetText 1, intRow2, "1"
                spdResult1.maxrows = intRow2

                .SetText 1, intRow1, ""
            End If
        Next
    End With
    
'    Dim aROW    As Integer, aCOL   As Integer
'    Dim varChk  As Variant, varBar As Variant, varNum As Variant
'    Dim iRow    As Integer, iCnt   As Integer
'    Dim strRack_tmp As String
'
'    With spdResult1
'        iCnt = 1
'        .GetText 1, 1, varChk
'        .GetText 2, 1, varBar
'        varNum = 1
'        If Trim(varChk) = "1" And Trim(varBar) <> "" Then
'            For iRow = 1 To .maxrows
'                strRack_tmp = Format(varNum, "00000")
'                .SetText 5, iRow, strRack_tmp
'                .SetText 6, iRow, ((iCnt Mod 6) + 1) - 1
'                iCnt = iCnt + 1
'                If (iCnt Mod 6) = 0 Then
'                    varNum = varNum + 1
'                    iCnt = 1
'                End If
'            Next
'        End If
'    End With

       
End Sub

Private Sub cmdWorkList1_Click()
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String, intRet   As Integer
    
    Dim strSpcno    As String
    Dim intRow      As Integer, intCol  As Integer
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
    
    sqlDoc = "select * from Worklist" _
              & " where workdate = '" & mskOrdDate.Text & "'"
                 
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        With spdWorkList
            If strSpcno <> Trim$(adoRS(1) & "") Then
                intRow = intRow + 1
                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                .SetText 1, intRow, "1"
                .SetText 2, intRow, Trim$(adoRS(2) & "")
                .SetText 3, intRow, Trim$(adoRS(3) & "")
                .SetText 4, intRow, Trim$(adoRS(1) & "")
                
                spdWorkList.Row = intRow
                spdWorkList.Col = 2
                spdWorkList.BackColor = &HEDD3CD
                spdWorkList.Col = 3
                spdWorkList.BackColor = &HEDD3CD
                spdWorkList.Col = 4
                spdWorkList.BackColor = &HEDD3CD
                
                .BlockMode = True
                .Row = intRow
                .Col = 2
                .BackColor = &HEDD3CD
                .Col = 3
                .BackColor = &HEDD3CD
                .Col = 4
                .BackColor = &HEDD3CD
                .Col = 1
                .Position = PositionUpperLeft
                .Action = ActionActiveCell
                .Action = ActionGotoCell
                 If .maxrows > 1 Then
                    .Row = 1:       .Row2 = .maxrows - 1
                    .Col = 1:       .Col2 = .MaxCols
                    .BackColor = vbWhite
                End If
                .BlockMode = False
            End If
        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing

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
    Dim MHead  As String, Pinfo As String
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
    
    Select Case comEQP.CommEvent
        Case comEvReceive
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If
            brStr = ""
            brStr = comEQP.Input
            
            txtResult.Text = txtResult.Text + brStr
                       
            If left$(brStr, 1) = ENQ Or Mid(brStr, 2) = ENQ Then '광주추가(Mid(brStr, 2) = ENQ)
                Debug.Print "[ELEC] " & ENQ
                comEQP.Output = ACK
                Debug.Print "[HOST] " & ACK
                fRcvString = ""
                Exit Sub
            End If
            
            For ii = 1 To Len(brStr)
                fRcvString = fRcvString + Mid(brStr, ii, 1)
            Next ii
            
            spdResult1.Col = 1: spdResult1.Row = OrderCnt
            If Len(Trim(spdResult1.Text)) > 0 Then
                '-- Ordering
                If left$(brStr, 1) = ACK Then
                    Debug.Print "[ELEC] " & ACK
                    Select Case SendCount
                        Case 0      'Message Header
                            MHead = "1H|\^&||||||||||P||" & vbCr & ETX
                            comEQP.Output = STX & MHead & MakeCS(MHead) & vbCr & vbLf
                            Debug.Print "[HOST] " & STX & MHead & MakeCS(MHead) & vbCr & vbLf
                            SendCount = SendCount + 1
                            MHead = ""
                        Case 1      'patient information
                            Pinfo = "2P|1|||||||||||||||||||||||||||||||||" & vbCr & ETX
                            comEQP.Output = STX & Pinfo & MakeCS(Pinfo) & vbCr & vbLf
                            Debug.Print "[HOST] " & STX & Pinfo & MakeCS(Pinfo) & vbCr & vbLf
                            SendCount = SendCount + 1
                            Pinfo = ""
                        Case 2      'Test Order
                            SendCount = SendCount + 1
                            spdResult1.Row = OrderCnt
                            spdResult1.Col = 2
                            
                            oPatNo = spdResult1.Text
                                                       
                            spdResult1.Col = 4: oIdNo = spdResult1.Text
                            spdResult1.Col = 5: oRackNo = spdResult1.Text
                            spdResult1.Col = 6: oPosNo = spdResult1.Text
                                                        
                            Orderoutput = ""
                            
                            With spdResult1
                                For intCol = 7 To .MaxCols
                                    spdResult1.GetText intCol, 0, varTmp
                                    If Trim$(varTmp) = "" Then Exit For
                                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                    If Not itemX Is Nothing Then
                                        spdResult1.Col = intCol:
                                        If spdResult1.BackColor = &HC6FEFF Then
                                            If Orderoutput = "" Then
                                                Orderoutput = "^^^" & Trim(itemX.tag) & "^0"
                                            Else
                                                Orderoutput = Orderoutput & "\" & "^^^" & Trim(itemX.tag) & "^0"
                                            End If
                                        End If
                                    End If
                                    Set itemX = Nothing
                                Next intCol
                            End With

'                            Orderoutput = Mid(Orderoutput, 1, Len(Orderoutput) - 1)
                            
                            Orderoutput = "3O|1|" + oIdNo + "|^" + oRackNo + "^" + oPosNo + "^^SAMPLE^NORMAL|" + Orderoutput + "|R||||||N||||||||||||||O" & vbCr & ETX
                            OutPutData = STX & Orderoutput & MakeCS(Orderoutput) & vbCr & vbLf
                            comEQP.Output = OutPutData
                            'MsgBox OutPutData
                            Debug.Print "[HOST] " & OutPutData
                        Case 3      'Message Terminator
                            SendCount = SendCount + 1
                            comEQP.Output = STX & "4L|1|N" & vbCr & ETX & "07" & vbCr & vbLf
                            Debug.Print "[HOST] " & STX & "4L|1|N" & vbCr & ETX & "07" & vbCr & vbLf
                        Case Else
                            comEQP.Output = EOT
                            Debug.Print "[HOST] " & EOT
                            '-- 오더내역이 남아있는지 체크
                            With spdResult1
                                For ii = OrderCnt To .maxrows
                                    .Col = 1: .Row = ii
                                    If .Value = 1 Then
                                        .Col = 2
                                        If .BackColor = vbWhite Then
                                            .BackColor = vbCyan
                                            .Col = 3
                                            .BackColor = vbCyan
                                            .Col = 4
                                            .BackColor = vbCyan
                                            .Col = 5
                                            .BackColor = vbCyan
                                            .Col = 6
                                            .BackColor = vbCyan
                                            OrderCnt = OrderCnt + 1
                                            .Col = 2
                                            If Len(Trim(.Text)) > 0 Then
                                                '-- osw edit
                                                .Row = ii '+ 1
                                                If Len(Trim(.Text)) > 0 Then
                                                    comEQP.Output = ENQ
                                                    Debug.Print "[HOST] " & ENQ
                                                End If
                                                SendCount = 0
                                                Exit For
                                            End If
                                        End If
                                    End If
                                    Me.Enabled = True
                                Next
                            End With
                            SendCount = 0
                    End Select
                End If
            End If
                        
            sStxCheck = InStr(fRcvString, STX)
            sEtxCheck = InStr(fRcvString, ETX)
            sCrcheck = InStr(fRcvString, vbCr)
            
            If sStxCheck <> 0 And sEtxCheck <> 0 And sCrcheck <> 0 Then
                Call psDataDefine(fRcvString, fChannel(), spdResult1)
                Debug.Print fRcvString
                fRcvString = ""
            End If
            
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

Private Sub ComReceive(ByRef RecData As String)
    
'------------------------------세미양방향일 경우 사용함---------------------------------------------
'---------------------------------------------------------------------------------------------------
Dim strRec  As String, strBuff  As String
    Dim strTmp  As String, intIdx   As Integer
    Dim intPos0 As Integer, intPos1 As Integer, intPos2 As Integer

    Dim age As String
    Dim i As Integer
    Dim tt As Boolean
    Dim sHead       As String
    Dim sPInfo      As String
    Dim sRtypeId    As String  'Record Type ID(1)
    Dim sSNumber    As String  'Sequence Number(6)
    '---[Specimen ID]----------------------------
    Dim sSampleNo   As String  'Sample No(5)
    Dim sSampleId   As String  'Sample ID(13)
    Dim sSampleType As String  'Sample Type(1)
    Dim sRackId     As String  'Rack Id(5)
    Dim sPositionNo As String  'Position No(1)
    '--------------------------------------------
    Dim sSpecimenID As String  'Specimen ID(2)
    '---[Universal Test Id]----------------------
    Dim sAppCode    As String  'Application Code(3)
    Dim sIdc        As String  'Inc,Dec or Cir(3)
    '--------------------------------------------
    Dim sPriority   As String  'Priority(1)
    Dim sRDateTime  As String  'Requested/Ordered Date and Time
    Dim sSDateTime  As String  'Specimen Collection Date and Time(14)
    Dim sCEndTime   As String  'Collection End Time
    Dim sCvolume    As String  'Collection Volume
    Dim sCId        As String  'Collection Id
    Dim sACode      As String  'Action Code(1)
    Dim sDCode      As String  'Danger Code
    Dim sRcinfo     As String  'Relevant Clinical Information(7)
    Dim sDtSpeR     As String  'Date/Time Specimen Received
    Dim sSpeDesc    As String  'Specimen Descriptor(2)
    Dim sOrderPh    As String  'Ordering Physician
    Dim sPtNum      As String  'Physician's Telephone Number
    Dim sUserF1     As String  'User Field No1(6)
    Dim sUserF2     As String  'User Field No2(104)
    Dim sLaboF1     As String  'Laboratory Field No.1
    Dim sLaboF2     As String  'Laboratory Field No.2
    Dim sDtRr       As String  'Date/Time Result(14)
    Dim sIccs       As String  'Instrument Charge to Computer System
    Dim sIsId       As String  'Instrument Section ID
    Dim sReportT    As String  'Report Types(1)
    Dim ii As Integer
    Dim sTempid As String
    Dim Orderoutput As String
    Dim OutPutData As String
    Dim Testcd As String, sOrderLst As String
    Dim Loop_count As Integer, pDoCount, pChnoCount As Integer
    Dim SEX As String
    Dim intldx As Integer
    Dim sCrLfCheck As Integer
    Dim strOrder As String
        
    Dim SendBuffD           As String           'data
    
    Dim Lencheck
    Dim Specimen            As String

    Static OrgMsg As String

    strRec = RecData
    Print #1, strRec;

    Debug.Print strRec
    Call psDataDefine(strRec, fChannel(), spdResult1)

End Sub

Private Sub psDataDefine(ByVal brbarcd As String, ByRef brChannel() As String, ByVal brspread As Object)

    Dim sTemp       As String       ' On Com으로부터 넘겨받은 Receive Data
    Dim Channel_No  As String       ' 문자형 변수
    Dim Patiant_No  As String       ' 환자번호
    Dim pGrid_Point As Integer      ' 해당 검사자 Point
    Dim Max_Arary_Cnt As Integer    ' 검사 항목수
    '-------------------------------' 임시 변수들.....
    Dim sDeCnt      As Integer
    Dim pDoCount    As Integer
    Dim Loop_count  As Integer
    Dim sRtn As Integer, sChannel As String, sRstText As String, sRstValue As Single, sUnit As String
    Dim itemX As ListItem
    Dim strRstval(1 To 19) As String, strRefVal(1 To 19)  As String
    Dim FunStr As String
    Dim sqlDoc  As String
    Dim intCol As Integer
    Dim Test_Cd() As String, strPid()    As String, strPnm() As String
    Dim Rev As Long
    Dim ii As Integer
    Dim tmpTstCd As String
    Dim strLevel() As String
    Dim chkPos  As Variant
    Dim strResult As String
    Dim strBarno    As String, strSPid  As String, strSPnm   As String
    Dim strSex      As String, strOld   As String, strArea   As String
    Dim varTmp  As Variant
    Dim strDate As String, strTime  As String, sqlRet   As Integer
    Dim strResultTmp As String
    Dim sLimit As String
    Dim sSeq, strTmp, strDate1, sCol As String
    
    On Error GoTo errDefine
    sRstText = brbarcd
    
    Erase fQuattro
       
    pDoCount = 0: sLimit = ""
    Do While InStr(sRstText, Chr$(13)) > 0
        If pDoCount = 0 And InStr(Text_Redefine(sRstText, Chr$(13)), "[") > 0 And InStr(Text_Redefine(sRstText, Chr$(13)), "]") > 0 Then
            '-- pDoCount = 0 : 검사명
            ReDim Preserve fQuattro(pDoCount)
            fQuattro(pDoCount) = Text_Redefine(sRstText, Chr$(13))
            fQuattro(pDoCount) = Trim(Mid(fQuattro(pDoCount), InStr(Text_Redefine(sRstText, Chr$(13)), "[") + 1, InStr(Text_Redefine(sRstText, Chr$(13)), "]") - 2))
'                Debug.Print fQuattro(pDoCount)
            pDoCount = pDoCount + 1
        ElseIf pDoCount = 1 And InStr(Text_Redefine(sRstText, Chr$(13)), "[") > 0 And InStr(Text_Redefine(sRstText, Chr$(13)), "]") > 0 Then
            ReDim Preserve fQuattro(pDoCount)
            fQuattro(pDoCount) = Text_Redefine(sRstText, Chr$(13))
'                Debug.Print fQuattro(pDoCount)
            pDoCount = pDoCount + 1
        ElseIf pDoCount = 2 Then
            '-- pDoCount = 2 : Header
            ReDim Preserve fQuattro(pDoCount)
            fQuattro(pDoCount) = Text_Redefine(sRstText, Chr$(13))
'                Debug.Print fQuattro(pDoCount)
            pDoCount = pDoCount + 1
        ElseIf pDoCount > 2 Then
            '-- pDoCount > 2 : Result
            If Mid(Trim(Text_Redefine(sRstText, Chr$(13))), 2, 1) <> """" And InStr(Text_Redefine(sRstText, Chr$(13)), "[") = 0 And InStr(Text_Redefine(sRstText, Chr$(13)), "]") = 0 Then
                ReDim Preserve fQuattro(pDoCount)
                fQuattro(pDoCount) = Text_Redefine(sRstText, Chr$(13))
                'Debug.Print fQuattro(pDoCount)
                pDoCount = pDoCount + 1
            End If
        End If

        If UCase(Mid(sRstText, 1, 6)) = "CUTOFF" Or UCase(Mid(sRstText, 1, 7)) = "CUT_OFF" Or UCase(Mid(sRstText, 1, 7)) = "CUT-OFF" Then
            sLimit = Text_Redefine(sRstText, Chr$(13))
            sLimit = Right(Trim(sLimit), 5)
        End If
        
        sRstText = Mid$(sRstText, InStr(sRstText, Chr$(13)) + 2)
    Loop
   
    sRstText = ""

    With spdResult1
        For Loop_count = 3 To UBound(fQuattro())
            fQuattro_1 = Split(fQuattro(Loop_count), ",")
            sSeq = Trim(Replace(fQuattro_1(0), Chr(34), ""))
            sCol = 0
            pGrid_Point = SeqSearch(spdResult1, sSeq, sCol)
            Channel_No = UCase(Replace(fQuattro_1(1), Chr(34), ""))
            If pGrid_Point > 0 Then
                For pDoCount = 5 To .MaxCols
                    .Row = pGrid_Point
                    .Col = pDoCount
                    .GetText 2, pGrid_Point, varTmp:    strBarno = Trim$(varTmp)
                    .GetText 3, pGrid_Point, varTmp:    strSPnm = Trim$(varTmp)
                    .GetText 4, pGrid_Point, varTmp:    strSPid = Trim$(varTmp)
                    .GetText pDoCount, 0, varTmp
                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                    If Len(Channel_No) > 0 And Channel_No = UCase(Trim$(varTmp)) Then
                        If Trim(fQuattro_1(3)) <> "" Then
                            strResultTmp = Replace(Trim(fQuattro_1(3)), Chr(34), "")
                            Select Case UCase(strResultTmp)
                                Case UCase("neg")
                                    strResult = "음성"
                                Case UCase("pos")
                                    strResult = "양성"
                            End Select
                             .Text = strResult
                        Else
                            .Text = ""
                        End If
        
                        If strResult <> "" Then
                            strDate = Format$(Now, "YYYYMMDD"): strTime = Format$(Now, "MMSS")
        
                                    sqlDoc = "Update INTERFACE003" & _
                                             "   set RSTVAL  = '" & strResult & "', REFVAL = ''" & _
                                             " where SPCNO   = '" & strSPid & "'" & _
                                             "   and EQPNUM  = '" & itemX.tag & "'" & _
                                             "   and TRANSDT = '" & strDate & "'" & _
                                             "   and TRANSTM = '" & strTime & "'"
                                    AdoCn_Jet.Execute sqlDoc
        
                                    sqlDoc = "insert into INTERFACE003(" & _
                                             "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                             "    values( '" & strSPid & "', '" & itemX.Text & "', '" & itemX.tag & "'," & _
                                             "            '" & strDate & "', '" & strTime & "'," & _
                                             "            '" & strResult & "', ''," & _
                                             "            '" & INS_CODE & "', '', '" & strSPnm & "', '" & strBarno & "')"
                                    AdoCn_Jet.Execute sqlDoc
                            
                            '-- 서버결과등록
                            If chkAuto.Value = "1" Then
                                sqlDoc = "update EXAM_TOC set EX_INRV = '" & Trim(strResult) & "',EX_INST = '2',EX_DATE = '" & Format$(Now, "YYYYMMDD") & "',EX_INEM='" & Trim(CurrUser.CuUserID) & "'" _
                                       & " where RE_RCID ='" & strSPid & "' And IN_CODE='" & itemX.Text & "'"

                                AdoCn_ORACLE.Execute (sqlDoc)
                                lblStatus.Caption = "저장 성공!!"
                                AdoCn_ORACLE.Execute sqlDoc
                            End If
        
                            Set itemX = Nothing
                        End If
                    End If
                    .SetText 1, vRow, 0
                Next pDoCount
            End If
        Next Loop_count
    End With



    Exit Sub
errDefine:

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
            If Trim(.Text) = "" Then
                SeqNullSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Function f_funAdd_Server(ByVal strBarno As String, ByVal strTestCd As String, _
                                 ByVal strTestval As String, ByRef strOrdLst() As String) As Boolean
                                 
    Dim strErrMsg       As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval()  As String
    Dim strTmp1()       As String, strTmp2()    As String, strTmp   As String
    Dim intPos          As Integer, intIdx      As Integer
    Dim blnFlag         As Boolean
    
    blnFlag = False
    f_funAdd_Server = False
    
    strTmp = strTestCd: intPos = InStr(strTmp, ",")
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

Private Function f_funAdd_QcServer(ByVal strBarno As String, ByVal strTestCd As String, _
                                 ByVal strTestval As String, ByRef strOrdLst() As String) As Boolean
                                 
    Dim strErrMsg       As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval()  As String
    Dim strTmp1()       As String, strTmp2()    As String, strTmp   As String
    Dim intPos          As Integer, intIdx      As Integer
    Dim blnFlag         As Boolean
    
    blnFlag = False
    f_funAdd_QcServer = False
    
    strTmp = strTestCd: intPos = InStr(strTmp, ",")
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

    If brspread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brspread
        If optSeq.Value = False Then
            For sCnt = 1 To .maxrows
                .Row = sCnt
                .Col = brCol
                If Trim(.Text) = brSeq Then
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
                If Trim(spdResult1.StartingRowNumber + (Val(sCnt) - 1)) = Val(brSeq) Then
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

   
    strDta = ";A1               3300001   3 0330051941" & vbCrLf
'    strDta = ":A1            0503300001   3 0330051941 32  1    35   2    43   3    58   4   160   5   225   6   373   7    47   8   7.9   9   4.9  10   1.0  11   0.2  12   243  13  13.5  14   1.0  15   3.2  16    51  17   372  18   104  19   129  20   164  22   123  23   7.5  24   3.2  25   4.7  87   141  88   4.3  89   100  93   1.6  94   287  95   3.0  96    14  97     5 " & vbCrLf

    Call ComReceive(strDta)
    
End Sub

Private Sub FileBep2000_Click()
    Dim wkbuf As String
    
    'Test_OpenFlag = 1
'    ExportPath = App.Path & "\document\exportfiles"
    
    Open ExportPath & "\" & FileBep2000.FileName For Input As #3
    
    'Test_OpenFlag = 2
    wkbuf = ""
    
    Do While Not EOF(3)
        wkbuf = wkbuf & Input(1, #3)
    Loop

'    Debug.Print wkbuf
    Close #3
    
    Call ComReceive(wkbuf)
    
    FileBep2000.Visible = False
End Sub

Private Sub Form_Activate()
    
    If IS_SET = False Then Unload Me

End Sub

Private Sub Form_Load()
    
'    Me.Show
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    CaptionBar1.Caption = INS_NAME & " Communication"
    
    Call cmdClear               ' 초기화
    Call f_subSet_ItemHeader    ' 리스트해더
    Call f_subSet_ItemList      ' 검사항목
    
    Call f_subSet_ComCharacter  ' 통신문자
    Call f_subGet_Setting       ' 통신설정
    
    Call cmdRun           ' 실행
    
    mskRstDate.Text = Format$(Now, "YYYYMMDD")
    mskRstDate1.Text = Format$(Now, "YYYYMMDD")
    mskOrdDate.Text = Format$(Now, "YYYYMMDD")
    mskOrdDate1.Text = Format$(Now, "YYYYMMDD")
    Open App.Path + "\" + "Quattro.log" For Append As #1

    Print #1, Chr(13) + Chr(10);
    
    f_strJOB_FLAG = "1":    f_intSampleNo = 0
    cboRstgbn(1).ListIndex = 2
    tabWork.Tab = 0
    Or_Seq = 1
    strFrameNo = 1
    SendCount = 0
    
'    If D0COM_CENTERCOD = "10" Then
'        cmdStartNo.Visible = False
'        cmdRackNo.Visible = True
'    Else
'        cmdStartNo.Visible = True
'        cmdRackNo.Visible = False
'    End If
    
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
'                .InputMode = Trim(mAdoRs.Fields("COM_INPUTMOD") & "")
'                .DTREnable = Trim(mAdoRs.Fields("COM_DTR") & "")
'                .EOFEnable = Trim(mAdoRs.Fields("COM_EOF") & "")
'                .NullDiscard = Trim(mAdoRs.Fields("COM_NULDIS") & "")
                .RTSEnable = Trim(mAdoRs.Fields("COM_RTS") & "")
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
    
    Close #1

End Sub

Private Sub FrameError_Click()
    txtResult.Visible = True
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

Private Sub Label9_DblClick()

    If COM_MODE = "1" Then
        COM_MODE = "0"
        ShowMessage "인터페이스 내용을 화면에 출력하지 않습니다."
    Else
        COM_MODE = "1"
        ShowMessage "인터페이스 내용을 화면에 출력합니다."
    End If
End Sub

Private Sub mskOrdDate_GotFocus()

    With mskOrdDate
        .SelStart = 8
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub mskOrdDate_KeyPress(KeyAscii As Integer)

    If Not KeyAscii = vbKeyBack Then mskOrdDate.SelLength = 1
    
End Sub


Private Sub mskRstDate_GotFocus()

    With mskRstDate
        .SelStart = 0
        .SelLength = Len(.Text) + 2
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
            comEQP.Output = Order.MSG_ENQ
        Case 2
            comEQP.Output = Order.MSG_HEADER
        Case 3
            comEQP.Output = Order.MSG_PATIENT
        Case 4
            comEQP.Output = Order.MSG_ORDER
        Case 5
            comEQP.Output = Order.MSG_TERMINATION
        Case 6
            comEQP.Output = Order.MSG_EOT
        Case Else
    End Select
    
End Sub

Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol1 As Integer
    Dim intCol2 As Integer
    Dim intRow1 As Integer
    Dim intRow2 As Integer
    Dim iCnt    As Integer
    
    If Row = 0 Then Exit Sub
    
    intCol1 = 10
    intCol2 = 2
    intRow1 = 1
    
    With spdResult1
        For iCnt = intCol1 To .MaxCols
            .Row = Row
            .Col = intCol1
            
            spdRstview.Row = intRow1
            spdRstview.Col = intCol2
            spdRstview.Text = .Text
            
            intRow1 = intRow1 + 1
            intCol1 = intCol1 + 1
            
            If intRow1 > spdRstview.maxrows Then
                intRow1 = 1
                intCol2 = intCol2 + 2
            End If
        
        Next
    End With
End Sub

Private Sub spdResult1_KeyPress(KeyAscii As Integer)

    Dim aROW    As Integer, aCOL   As Integer
    Dim varChk  As Variant, varBar As Variant, varNum As Variant
    Dim iRow    As Integer, iCnt   As Integer
    
    'Debug.Print Col & NewCol & Row & NewRow
       
    If KeyAscii = vbKeyReturn Then
        With spdResult1
            aCOL = .ActiveCol
            aROW = .ActiveRow
            If aCOL = 4 Then
                iCnt = 0
                For iRow = aROW To .maxrows
                    .GetText 1, iRow, varChk
                    .GetText 3, iRow, varBar
                    .GetText aCOL, aROW, varNum
                    If Trim(varChk) = "1" And Trim(varBar) <> "" Then
                        .SetText aCOL, iRow, varNum
                        .SetText aCOL + 1, iRow, ((iCnt Mod 40) + 1) + (40 * (varNum - 1))
                        iCnt = iCnt + 1
                        If (iCnt Mod 40) = 0 Then varNum = varNum + 1
                    End If
                Next
'            ElseIf aCOL = 5 Then
'                iCnt = 0
'                For iRow = aROW To .maxrows
'                    .GetText 1, iRow, varChk
'                    .GetText 3, iRow, varBar
'                    .GetText aCOL, aROW, varNum
'                    If Trim(varChk) = "1" And Trim(varBar) <> "" Then
'                        .SetText aCOL, iRow, ((iCnt Mod 40) + varNum) '+ (40 * (varNum - 1))
'                        '.SetText aCOL - 1, iRow, varNum
'                        iCnt = iCnt + 1
'                        If (iCnt Mod 40) = 0 Then varNum = varNum + 1
'                    End If
'                Next
            
            End If
        End With
    End If
    
End Sub

Private Sub spdWorkList_Click(ByVal Col As Long, ByVal Row As Long)

    If Col < 3 Then Exit Sub
    
    Dim varTmp  As Variant
    
    With spdWorkList
        If Col = 1 Then
            .GetText 2, Row, varTmp
            If Trim$(varTmp) = "" Then Exit Sub
            
            .SetText 1, Row, IIf(Trim$(varTmp) = "1", "", "1")
        ElseIf Col > 4 Then
            .GetText Col, 0, varTmp
            If Trim$(varTmp) = "" Then Exit Sub
            
            .Row = Row: .Col = Col
            If .BackColor = vbWhite Then
                .BackColor = &HC6FEFF
            Else
                .BackColor = vbWhite
            End If
        End If
    End With
    
End Sub

Private Sub Timer1_Timer()

    comEQP.Output = ENQ

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
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For i = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
End Sub

Private Sub txtBarCode_Change()

    If txtBarCode.SelStart = txtBarCode.MaxLength Then SendKeys "{TAB}"
    
End Sub

Private Sub txtBarCode_GotFocus()

    With txtBarCode
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txtBarCode_KeyPress(KeyAscii As Integer)

    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Privete sub txtBarCode_LostFocus()"
    
    Dim varTmp  As Variant, strEqpCd    As String
    Dim intRow  As Integer, intCol  As Integer, blnFlag As Boolean
    Dim strOrdcd() As String, strPid()  As String, strPnm() As String
    Dim strPexzm() As String, strPeqpcd() As String
    Dim strEqcode() As String, strExamname() As String, strAcptno() As String

    Dim itemX   As ListItem
    
    If txtBarCode.Text = "" Then Exit Sub
    
    blnFlag = False
    If KeyAscii = vbKeyReturn Then
        intCol = sl_examdata_select&(txtBarCode.Text, INS_CODE, strEqcode, strExamname, strOrdcd, strPid, strPnm, strAcptno)
        
        For intCol = 0 To UBound(strOrdcd)
            If strOrdcd(intCol) <> "" Then
                strEqpCd = f_funGet_CODE(strOrdcd(intCol))
                Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                If Not itemX Is Nothing Then
                    If Not blnFlag Then
                        intRow = f_funGet_SpreadRow(spdResult1, 2, txtBarCode.Text)
                        If intRow < 1 Then
                            intRow = f_funGet_SpreadRow(spdResult1, 2, "")
                            If intRow < 1 Then
                                spdResult1.maxrows = spdResult1.maxrows + 1
                                spdResult1.RowHeight(spdResult1.maxrows) = 13
                                intRow = spdWorkList.maxrows
                            End If
                            spdResult1.SetText 2, intRow, txtBarCode.Text
                            spdResult1.SetText 3, intRow, strPnm(0)
                            spdResult1.SetText 4, intRow, strPid(0)
                        End If
                        spdResult1.SetText 1, intRow, "1"
                    End If
                        
                    'spdResult1.SetText itemX.Index + 6, intRow, "V"
                    spdResult1.Col = itemX.Index + 6
                    spdResult1.Row = intRow
                    spdResult1.BackColor = &HC6FEFF
                    
                    blnFlag = True
                End If
            End If
        Next
    
        If Not blnFlag Then MsgBox "해당 검사항목이 존재하지 않은 검체입니다.", vbInformation, Me.Caption
        
        txtBarCode.Text = "":   txtBarCode.SetFocus
        Exit Sub
    
    End If
    
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)

End Sub

Private Function psDataExists() As Boolean
Dim sCnt As Long
    
    psDataExists = False
    With spdWorkList
        For sCnt = 1 To .maxrows
            .Row = sCnt:    .Col = 2
            If Trim(.Text) = Mid(txtBarCode.Text, 1, 11) Then
                psDataExists = True
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Sub txtResult_DblClick()
    txtResult.Text = ""
    List1.Text = ""
    
    If txtResult.Visible Then txtResult.Visible = False
    List1.Visible = True
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmComm 
   Caption         =   "Interface"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16035
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9960
   ScaleWidth      =   16035
   WindowState     =   2  '최대화
   Begin VB.Timer tmrReceive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6015
      Top             =   5130
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6420
      Top             =   5130
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   4845
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
      Left            =   5460
      Top             =   5100
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   1
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin MSComctlLib.ImageList imlStatus 
      Left            =   6825
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
      Top             =   9225
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
         Left            =   9285
         TabIndex        =   26
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
         Left            =   10560
         TabIndex        =   27
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
         Left            =   11865
         TabIndex        =   28
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
         Left            =   13170
         TabIndex        =   29
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
      Begin BHButton.BHImageButton cmdInit 
         Height          =   405
         Left            =   8040
         TabIndex        =   54
         Top             =   90
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   714
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
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16035
      _ExtentX        =   28284
      _ExtentY        =   609
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
         Top             =   75
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send : "
         Height          =   180
         Left            =   13020
         TabIndex        =   3
         Top             =   75
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port : "
         Height          =   180
         Left            =   11925
         TabIndex        =   2
         Top             =   75
         Width           =   510
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   14925
         Picture         =   "frmComm.frx":5794
         Top             =   45
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   13695
         Picture         =   "frmComm.frx":5D1E
         Top             =   45
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12555
         Picture         =   "frmComm.frx":62A8
         Top             =   45
         Width           =   240
      End
   End
   Begin TabDlg.SSTab tabWork 
      Height          =   8610
      Left            =   60
      TabIndex        =   7
      Top             =   375
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   15187
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   16711680
      TabCaption(0)   =   " ▒    WorkList "
      TabPicture(0)   =   "frmComm.frx":6832
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSPanel1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameResult"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "spdWorklist"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdWorkList"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "pnlCom2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "pnlCom"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdStartNo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdPosNo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdRackNo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdWordQuery"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdOrder"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdAppend(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtBarCode"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkReTest"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "fraError"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "SSPanel2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdSel(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdSel(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "UserPanel1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "chkAuto"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "spdResult1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "spdRstview"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lvwCuData"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Timer2"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   " ▒   받은 결과  "
      TabPicture(1)   =   "frmComm.frx":684E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "spdResult2"
      Tab(1).Control(1)=   "SSPanel3"
      Tab(1).Control(2)=   "cmdAppend(1)"
      Tab(1).Control(3)=   "cmdExcel"
      Tab(1).Control(4)=   "cmdSel(2)"
      Tab(1).Control(5)=   "cmdSel(3)"
      Tab(1).ControlCount=   6
      Begin VB.Timer Timer2 
         Interval        =   30000
         Left            =   7500
         Top             =   4800
      End
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   4515
         Left            =   2400
         TabIndex        =   53
         Top             =   2490
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   7964
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
      Begin FPSpread.vaSpread spdRstview 
         Height          =   7995
         Left            =   12750
         TabIndex        =   71
         Top             =   360
         Width           =   2325
         _Version        =   393216
         _ExtentX        =   4101
         _ExtentY        =   14102
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
         MaxCols         =   6
         MaxRows         =   30
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmComm.frx":686A
         UserResize      =   0
      End
      Begin FPSpread.vaSpread spdResult1 
         Height          =   8160
         Left            =   30
         TabIndex        =   72
         Top             =   360
         Width           =   12675
         _Version        =   393216
         _ExtentX        =   22357
         _ExtentY        =   14393
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   5
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
         MaxCols         =   9
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         SpreadDesigner  =   "frmComm.frx":70DF
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
         Left            =   11925
         TabIndex        =   25
         Top             =   90
         Value           =   1  '확인
         Width           =   1410
      End
      Begin HSCotrol.UserPanel UserPanel1 
         Height          =   30
         Left            =   3420
         TabIndex        =   52
         Top             =   390
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   53
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   345
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   609
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":76AE
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   609
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":7B30
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   465
         Left            =   9570
         TabIndex        =   44
         Top             =   90
         Visible         =   0   'False
         Width           =   3075
         _Version        =   65536
         _ExtentX        =   5424
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   16761024
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
            Left            =   1650
            TabIndex        =   46
            Top             =   90
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
            Height          =   285
            Left            =   210
            TabIndex        =   45
            Top             =   90
            Width           =   1455
         End
      End
      Begin Threed.SSFrame fraError 
         Height          =   2355
         Left            =   8160
         TabIndex        =   39
         Top             =   5760
         Width           =   6915
         _Version        =   65536
         _ExtentX        =   12197
         _ExtentY        =   4154
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
         Begin VB.ListBox LstErr 
            Height          =   2040
            Left            =   135
            TabIndex        =   42
            Top             =   225
            Width           =   6675
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
            TabIndex        =   41
            Top             =   225
            Visible         =   0   'False
            Width           =   6630
         End
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
         Left            =   3750
         TabIndex        =   40
         Top             =   90
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TEST"
         Height          =   285
         Left            =   11130
         TabIndex        =   24
         Top             =   60
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.TextBox txtBarCode 
         Height          =   300
         Left            =   9120
         MaxLength       =   12
         TabIndex        =   8
         Top             =   1845
         Visible         =   0   'False
         Width           =   1500
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   3
         Left            =   -74640
         TabIndex        =   22
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":7F9E
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   -74910
         TabIndex        =   23
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":8420
      End
      Begin BHButton.BHImageButton cmdAppend 
         Height          =   300
         Index           =   0
         Left            =   13440
         TabIndex        =   43
         Top             =   0
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
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
      Begin BHButton.BHImageButton cmdOrder 
         Height          =   375
         Left            =   6390
         TabIndex        =   47
         Top             =   450
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
         Caption         =   "오더전송"
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
         Height          =   375
         Left            =   8670
         TabIndex        =   48
         Top             =   30
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
      Begin BHButton.BHImageButton cmdRackNo 
         Height          =   375
         Left            =   6150
         TabIndex        =   49
         Top             =   30
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
         Caption         =   "Track변경"
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
         Height          =   375
         Left            =   5130
         TabIndex        =   50
         Top             =   450
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
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
         Left            =   7410
         TabIndex        =   51
         Top             =   30
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
      Begin HSCotrol.UserPanel pnlCom 
         Height          =   1155
         Left            =   9720
         TabIndex        =   31
         Top             =   4620
         Visible         =   0   'False
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   2037
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
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   32
            Top             =   315
            Visible         =   0   'False
            Width           =   11595
         End
         Begin VB.Frame Frame1 
            Height          =   645
            Left            =   45
            TabIndex        =   33
            Top             =   3660
            Visible         =   0   'False
            Width           =   11610
            Begin HSCotrol.CButton cmdCOMSave 
               Height          =   360
               Left            =   10515
               TabIndex        =   34
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
               TabIndex        =   35
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
               Left            =   9450
               TabIndex        =   36
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
               TabIndex        =   37
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
      Begin HSCotrol.UserPanel pnlCom2 
         Height          =   5340
         Left            =   5880
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   5910
         _ExtentX        =   10425
         _ExtentY        =   9419
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Frame Frame2 
            Height          =   645
            Left            =   90
            TabIndex        =   13
            Top             =   4635
            Width           =   5760
            Begin MSComDlg.CommonDialog cdlFile 
               Left            =   4290
               Top             =   90
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin HSCotrol.CButton cmdChksum 
               Height          =   360
               Left            =   2205
               TabIndex        =   14
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
               TabIndex        =   15
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
               TabIndex        =   16
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
               TabIndex        =   17
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
               TabIndex        =   18
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
               TabIndex        =   19
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
               TabIndex        =   20
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
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   21
            Top             =   1290
            Width           =   5730
         End
      End
      Begin BHButton.BHImageButton cmdWorkList 
         Height          =   480
         Left            =   90
         TabIndex        =   30
         Top             =   5175
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   847
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
         ImgOutLineSize  =   3
      End
      Begin FPSpread.vaSpread spdWorklist 
         Height          =   4230
         Left            =   90
         TabIndex        =   55
         Top             =   900
         Width           =   4755
         _Version        =   393216
         _ExtentX        =   8387
         _ExtentY        =   7461
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   1
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
         MaxCols         =   10
         MaxRows         =   5
         Protect         =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmComm.frx":888E
         UserResize      =   2
      End
      Begin Threed.SSFrame FrameResult 
         Height          =   3525
         Left            =   2700
         TabIndex        =   38
         Top             =   900
         Visible         =   0   'False
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
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   525
         Left            =   90
         TabIndex        =   56
         Top             =   360
         Width           =   4755
         _Version        =   65536
         _ExtentX        =   8387
         _ExtentY        =   926
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
         Begin MSMask.MaskEdBox mskOrdDate1 
            Height          =   300
            Left            =   2415
            TabIndex        =   57
            Top             =   120
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
            TabIndex        =   58
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin BHButton.BHImageButton cmdSearch 
            Height          =   390
            Left            =   3570
            TabIndex        =   59
            Top             =   60
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   688
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
         Begin VB.Label Label8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "접수일자 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   120
            TabIndex        =   62
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E0E0E0&
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
            Left            =   2280
            TabIndex        =   61
            Top             =   180
            Width           =   315
         End
         Begin VB.Label Label6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "분 접수까지."
            Height          =   255
            Left            =   5520
            TabIndex        =   60
            Top             =   840
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin BHButton.BHImageButton cmdExcel 
         Height          =   390
         Left            =   -66060
         TabIndex        =   63
         Top             =   330
         Visible         =   0   'False
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   688
         Caption         =   "Excel 파일 생성 / 출력"
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
         Left            =   -62340
         TabIndex        =   64
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   525
         Left            =   -74910
         TabIndex        =   65
         Top             =   360
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   926
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
            ItemData        =   "frmComm.frx":8F6B
            Left            =   2235
            List            =   "frmComm.frx":8F78
            Style           =   2  '드롭다운 목록
            TabIndex        =   66
            Top             =   135
            Width           =   1410
         End
         Begin MSMask.MaskEdBox mskRstDate 
            Height          =   300
            Left            =   1110
            TabIndex        =   67
            Top             =   135
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin BHButton.BHImageButton cmdRstQuery 
            Height          =   375
            Left            =   3735
            TabIndex        =   68
            Top             =   90
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
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "분 접수까지."
            Height          =   255
            Left            =   5520
            TabIndex        =   70
            Top             =   840
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "검사일자 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   120
            TabIndex        =   69
            Top             =   180
            Width           =   1095
         End
      End
      Begin FPSpread.vaSpread spdResult2 
         Height          =   7350
         Left            =   -74910
         TabIndex        =   73
         Top             =   900
         Width           =   15015
         _Version        =   393216
         _ExtentX        =   26485
         _ExtentY        =   12965
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   7
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
         MaxCols         =   8
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         SpreadDesigner  =   "frmComm.frx":8FA2
         UserResize      =   0
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
         TabIndex        =   9
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

Const colBANO = 1   '바코드번호
Const colORDT = 2   '처방일자
Const colORQN = 3   '처방번호
Const colPANM = 4   '환자명
Const colPAID = 5   '병록번호
Const colOIFL = 6   '입/외구분
Const colSENO = 7   '일련번호
Const colSEXS = 8   '성별
Const colAGES = 9   '나이
Const colNWNO = 10  '내원번호
Const colSQNO = 11  'SeqNo

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
Dim Patiant_Recevid As Integer
Dim sStxCheck As Integer
Dim sEtxCheck As Integer
' --------------------------------------------------------------
Dim strOrdLst As String

Dim fAdvia2120(100)       As String
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

Dim fAdvia2120Cfg(100) As Integer
Dim fAdvia2120Size(100, 1) As Integer
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
    strTestcd(100) As String
End Type

Private f_typCode() As TYPE_CD

Dim RecordChk As Boolean

Dim strJinCd As String

Private f_strBarno()    As String, f_strTest()  As String, f_strEtc()   As String
Private f_strRack()     As String, f_strCup()   As String
Private f_intCnt        As Integer, f_intIdx    As Integer

Dim objIntPhase     As Integer
Dim objIntBufCnt    As Integer
Dim objIntstate     As String
Dim objDicBuf       As String

Private objRst      As New clsCommon

Dim crCnt   As Integer

Private mORDER      As clsIISIntOrder             '오더정보 클래스

Dim mIntLibPhase    As Integer
Dim mIntLibstate    As String
Dim mIntLibBuffer   As String

Const SPCLEN As Integer = 12
Dim strGumCd        As String
Dim OrderSort_Flag  As Integer

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
    With spdWorklist
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
        
        With spdWorklist
            intCol = intCol + 1
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1:  .ColWidth(.MaxCols) = 6.5
            
            .SetText intCol, 0, adoRS.Fields("TESTNM")
        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Function f_subSet_WorkList(ByVal strDate As String, ByVal strDate1 As String, Optional ByVal strTime As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim pFrDt As String
    Dim pToDt As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
   
    pFrDt = mskOrdDate.Text
    pToDt = mskOrdDate1.Text
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    
                '-- 처방일자,처방일련번호,환자명,환자번호,입외구분,일련번호,성별,나이,내원번호,처방코드
             sqlDoc = "Select a.ORDT,a.ORQN ,b.PANM,a.PAID,a.OIFL,a.SENO,b.SEXS,b.AGES,a.NWNO,a.ORCD "
    sqlDoc = sqlDoc & "  From LRESULT a, APATINF b"
'    sqlDoc = sqlDoc & " Where a.ORDT between  '" & mskOrdDate.Text & "' and '" & mskOrdDate1.Text & "'"
    sqlDoc = sqlDoc & " Where a.ETDT between TO_DATE(" & Format(pFrDt, "########") & ",'yyyymmdd') + 0.000000 "
    sqlDoc = sqlDoc & "    and TO_DATE(" & Format(pToDt, "########") & ",'yyyymmdd') + 0.999999 " & vbCrLf
    sqlDoc = sqlDoc & "   And a.PAID = b.PAID "
    sqlDoc = sqlDoc & "   And a.ORCD in (" & strGumCd & ")"
    sqlDoc = sqlDoc & "   And a.OKFL <> 'Y' "   '-- 결과확정유무
    sqlDoc = sqlDoc & " Order By a.ORDT,a.PAID,a.ORQN,b.PANM,a.SENO"
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    
    AdoRs_ORACLE.CursorLocation = adUseClient
    AdoRs_ORACLE.Open sqlDoc, AdoCn_ORACLE
    
    If AdoRs_ORACLE.RecordCount = 0 Then
        Set f_subSet_WorkList = Nothing
        RecordChk = False
        Set AdoRs_ORACLE = Nothing
        Exit Function
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

Private Function f_subSet_WorkList_Barcode(ByVal strORDT As String, Optional ByVal strPAID As String, Optional ByVal strSENO As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim stryy, strmm, strdd, strDate  As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList_Barcode() As ADODB.Recordset"
    
    
    Set AdoRs_ORACLE = New ADODB.Recordset

             sqlDoc = "SELECT DiSTINCT b.SCP42IDNOA, a.SCP41NAME, a.SCP41JDATE, b.SCP42SUGACD "
    sqlDoc = sqlDoc & vbCrLf & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b "
    sqlDoc = sqlDoc & vbCrLf & " WHERE a.SCP41PCODE = b.SCP42PCODE"
    sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41JDATE = b.SCP42JDATE"
    sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41SID   = b.SCP42SID"
    sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41SPMNO2 = '" & strORDT & "'"
'    sqlDoc = sqlDoc & vbCrLf & "   AND b.SCP42SUGACD in (" & strGumCd & ")"
    sqlDoc = sqlDoc & vbCrLf & "   AND b.SCP42SUGACD in (" & strGumCd & "," & strJinCd & ")"


'    sqlDoc = sqlDoc & vbCrLf & "   AND (b.SCP42RSTCD <> 'N' OR b.SCP42RSTCD IS null)"
'    sqlDoc = sqlDoc & vbCrLf & "   AND (b.SCP42RSTCD <> 'N' OR b.SCP42RSTCD IS null OR b.SCP42PROFLG  <> 'M')"

    '-- 2012.04.13 수정
'    sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41SNDYN = 'N' "
'    sqlDoc = sqlDoc & vbCrLf & "   AND b.SCP42RESULT IS NULL "

    sqlDoc = sqlDoc & vbCrLf & "   AND b.SCP42RESULT IS NULL "

'    sqlDoc = sqlDoc & vbCrLf & "   AND b.SCP42RSTCD <> 'N'"
    
'    sqlDoc = sqlDoc & vbCrLf & "   AND b.SCP42RSTCD  <> 'N' "
'    sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41SNDYN  = 'N'"
    '-- 2012.04.03 추가
    'sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41SNDYN  <> 'N' " '--고정값:         'N'"
'    sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41RSTYN  <> 'Y' " '--고정값:         'Y'"
    'sqlDoc = sqlDoc & vbCrLf & "   AND b.SCP42RSTCD  = '' " '-- 결과형식 => 숫자 : 'N', 문자 : 'X', 장문 : 'R'"
    'sqlDoc = sqlDoc & vbCrLf & "   AND b.SCP42RESULT = ''   " '-- 결과값"

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


Private Function SeqSearch_PAID(ByVal brspread As Object, ByVal brQry1 As String, ByVal brQry2 As String, ByVal brQry3 As String, ByVal brCol As Integer) As Long
Dim sCnt As Long
Dim sCnt1 As Long
Dim sCnt2 As Long
Dim sCnt3 As Long

    SeqSearch_PAID = 0
    If brspread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt1 = 1 To .maxrows
            .Row = sCnt1
            .Col = 2
            If Trim(.Text) = brQry1 Then
                .Col = 5
                If Trim(.Text) = brQry2 Then
                    .Col = 3
                    If Trim(.Text) = brQry3 Then
                        SeqSearch_PAID = sCnt1
                        .Action = ActionActiveCell
                        .Refresh
                        Exit For
                    End If
                End If
            End If
        Next sCnt1
    End With

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
    
    Dim intPos1 As Integer
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = ""
    
    intCol = 12
    intCol2 = 1
    intRow = 1
    With spdWorklist
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    With spdResult2
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
'             " order by TESTCD_EQP, TESTCD"
             
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then
        adoRS.MoveFirst
        ReDim fChannel(adoRS.RecordCount)
        strJinCd = ""
        strGumCd = ""
    End If
    
    Do While Not adoRS.EOF
        If Trim(adoRS.Fields("TESTCD")) <> "" Then
            intPos1 = InStr(Trim(adoRS.Fields("TESTCD")), ",")
            If intPos1 = 0 Then
                strGumCd = strGumCd & "'" & Trim(adoRS.Fields("TESTCD")) & "',"
            Else
                strGumCd = strGumCd & "'" & Mid(Trim(adoRS.Fields("TESTCD")), 1, intPos1 - 1) & "',"
                strJinCd = strJinCd & "'" & Mid(Trim(adoRS.Fields("TESTCD")), intPos1 + 1) & "',"
            End If
        End If
        
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
        
        With spdWorklist
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
            .Col = intCol:  .ColHidden = True
        End With
        
        With spdResult1
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
                .ColWidth(intCol) = 6.5
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
                .ColWidth(intCol) = 6.5
            End If
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        fChannel(intCol - colSQNO) = adoRS.Fields("TEST_EQP")
        
        intCnt = intCnt + 1
        ReDim Preserve f_typCode(1 To intCnt) As TYPE_CD
        
        f_typCode(intCnt).strEqpCd = Trim$(adoRS.Fields("TEST_EQP"))
        f_typCode(intCnt).intCnt = 0
        
        strTmp = Trim$(adoRS.Fields("TESTCD"))
        intPos = InStr(strTmp, ",")
        Do While intPos > 0
            f_strOrdList = f_strOrdList + "'" + Mid$(strTmp, 1, intPos - 1) + "',"
            
            f_typCode(intCnt).intCnt = f_typCode(intCnt).intCnt + 1
            f_typCode(intCnt).strTestcd(f_typCode(intCnt).intCnt) = Mid$(strTmp, 1, intPos - 1)
            
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
    
    If Trim(strGumCd) <> "" Then strGumCd = Mid(strGumCd, 1, Len(strGumCd) - 1)
    If Trim(strJinCd) <> "" Then strJinCd = Mid(strJinCd, 1, Len(strJinCd) - 1)
    
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
            If Trim$(strOrdcd) = Trim$(f_typCode(intIdx1).strTestcd(intIdx2)) Then
                f_funGet_CODE = f_typCode(intIdx1).strEqpCd
                Exit Function
            End If
        Next
    Next
    
End Function

'Private Function f_subSet_ComList()
'    Dim sqlRet      As Integer
'    Dim sqlDoc      As String
'
'On Error GoTo ErrorTrap
'    CallForm = "clsCommon - Public Function f_subSet_ComList() As ADODB.Recordset"
'
'
'        Set AdoRs_SQL = New ADODB.Recordset
'
'        sqlDoc = "         SELECT B.COM_CODE, B.COM_NAME " & vbCr
'        sqlDoc = sqlDoc & "  FROM MDCK..GUMJIN_INTERFACE A, MDCK..TB_COMPANY B, MDCK..BAG_INTERFACECODE C " & vbCr
'        sqlDoc = sqlDoc & " WHERE A.Per_COM_Code = B.COM_CODE " & vbCr
'        sqlDoc = sqlDoc & "   AND A.per_gumjin_date BETWEEN '" & Trim(mskOrdDate.Text) & "' AND '" & Trim(mskOrdDate1.Text) & "'" & vbCr
'        sqlDoc = sqlDoc & "   AND SUBSTRING(C.KIND, 1, 1) = 'I' " & vbCr
'        sqlDoc = sqlDoc & "   AND A.EDPSCODE = C.MEDITEM " & vbCr
'        sqlDoc = sqlDoc & " GROUP BY B.COM_CODE, B.COM_NAME " & vbCr
'
'        Set AdoRs_SQL = New ADODB.Recordset
'        AdoRs_SQL.CursorLocation = adUseClient
'        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
'
'        If AdoRs_SQL.RecordCount > 0 Then
'            AdoRs_SQL.MoveFirst
'            cboComNm.Clear
'            Do Until AdoRs_SQL.EOF
'                cboComNm.AddItem AdoRs_SQL.Fields("COM_NAME") & ""
'                AdoRs_SQL.MoveNext
'            Loop
'            cboComNm.ListIndex = 0
'        End If
'
'        AdoRs_SQL.Close:  Set AdoRs_SQL = Nothing
'
'Exit Function
'
'ErrorTrap:
'    Set AdoRs_SQL = Nothing
'
'    Call ErrMsgProc(CallForm)
'
'
'End Function

'Private Sub cboChk_Click()
'    If Trim(cboChk.Text) = "검진" Then
'        Call f_subSet_ComList
'    Else
'        cboComNm.Clear
'    End If
'End Sub


'------------------------------------------------------------------------------'
'   기능 : Initialization Message 조회
'   반환 : Initialization Message
'------------------------------------------------------------------------------'
'Public Function GetInit() As String
'    Dim strInit As String   'Initialization 문자열
'
'    mMT = &H30
'    strInit = MT & "I " & vbCrLf
'    strInit = STX & strInit & GetLRC(strInit) & ETX
'    GetInit = strInit
'End Function

Private Sub cmdInit_Click()
    Dim strOutput As String     '송신할 데이터
    

    '## 포트가 오픈되어 있지 않으면 에러표시
    If comEQP.PortOpen = False Then
        MsgBox "포트가 열려있지 않습니다.", vbCritical, "오류"
        Exit Sub
    End If
    
    '## Initialization Message 전송
    strOutput = mORDER.GetInit
    comEQP.Output = strOutput
'    Print #1, "[INIT] " & strOutput & vbNewLine;
    Print #1, vbNewLine & Format(Now, "yyyymmdd hh:mm:ss") & "[INIT]" & strOutput & vbNewLine;
    
    mIntLibstate = "I"
'    mIntLibstate = ""
    
    Timer2.Enabled = True
    
End Sub

'Private Sub cmdOrder_Click()
'
'    Call SendOrder
'
'    Exit Sub
'
'    Dim varTmp      As Variant
'    Dim intRow      As Integer, intCol  As Integer
'    Dim strBarNo    As String
'
'    Dim itemX       As ListItem
'
'    f_intIdx = 0
'
'    With spdResult1
'        For intRow = 1 To .maxrows
'            .GetText 2, intRow, varTmp:    strBarNo = Trim$(varTmp) '검체번호
'            .GetText 1, intRow, varTmp
'
'            If strBarNo = "" Then Exit Sub
'
'            If Trim$(varTmp) = "1" Then
'                f_intIdx = f_intIdx + 1
'                ReDim Preserve f_strBarno(1 To f_intIdx) As String
'                ReDim Preserve f_strTest(1 To f_intIdx) As String
'                ReDim Preserve f_strEtc(1 To f_intIdx) As String
'
'                f_strBarno(f_intIdx) = strBarNo
'
'                f_strTest(f_intIdx) = ""
'                For intCol = 8 To .MaxCols
'                    .GetText intCol, 0, varTmp
'                    If Trim$(varTmp) = "" Then Exit For
'
'                    .Col = intCol:  .Row = intRow
'                    If .BackColor = &HC6FEFF Then
'                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
'                        If Not itemX Is Nothing Then
'                            f_strTest(f_intIdx) = f_strTest(f_intIdx) + IIf(Mid$(itemX.tag, 1, 1) = "X", "13", itemX.tag)
'
'                            If itemX.tag = "X1" Then
'                                f_strEtc(f_intIdx) = "P"
'                            ElseIf itemX.tag = "X2" Then
'                                f_strEtc(f_intIdx) = "F"
'                            End If
'                        End If
'                        Set itemX = Nothing
'                    End If
'                Next
'                '-- 추가 : 한번보낸검체는 체크를 풀어줌
'                .SetText 1, intRow, "0"
'            End If
'        Next
'    End With
'
'    f_intCnt = 0
'
'End Sub


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
                .Col = 9:       .Text = Trim(sAdd + Val(sNo))
                'If (sNo Mod 11) = 1 Then varNum = varNum + 1
                sAdd = sAdd + 1
            Next sCnt
        End With
    End If
End Sub

Private Sub cmdRackNo_Click()
    Dim sNo As String, sCnt As Integer, sAdd As Integer
    Dim aRow    As Integer, aCOL   As Integer
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
                    .SetText 8, iRow, varNum
                    .SetText 9, iRow, ((iCnt Mod 11) + 1) - 1
                    iCnt = iCnt + 1
                    If (iCnt Mod 11) = 1 Then varNum = varNum + 1
                Next
            End If
        End With
    End If
End Sub

Private Sub cmdSearch_Click()
    Dim strDoc As String
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim intRow      As Integer
    Dim pGrid_Point As Integer
    Dim intCnt      As Integer
    Dim strBarno As String
    Dim itemX As ListItem
    Dim strEqpCd As String
    Dim strBartmpNo As String
    Dim blt As Boolean
    
    With spdWorklist
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    blt = True
    
On Error GoTo ErrorTrap

    CallForm = "clsCommon - Public Function cmdSearch_Click() As ADODB.Recordset"
        
    Set mAdoRs = f_subSet_WorkList(mskOrdDate.Text, mskOrdDate1.Text)
    
    If RecordChk = False Then
        MsgBox Format(mskOrdDate.Text, "####-##-##") & "일 에서  " & Format(mskOrdDate1.Text, "####-##-##") & "일까지의 검사 대상자가 없습니다.", vbOKOnly + vbInformation, App.Title
        Exit Sub
    Else
        strBarno = ""
        mAdoRs.MoveFirst

        With spdWorklist
            For intCnt = 0 To mAdoRs.RecordCount - 1
                '-- 처방일자,처방일련번호,환자명,환자번호,입외구분,일련번호,성별,나이,내원번호,처방코드
                If strBarno <> Trim(mAdoRs.Fields("ORDT").Value & "") & Trim(mAdoRs.Fields("PAID").Value & "") Then
                    pGrid_Point = SeqSearch_PAID(spdWorklist, Trim(mAdoRs.Fields("ORDT").Value & ""), "", Trim(mAdoRs.Fields("PAID")), colORQN)
'                    pGrid_Point = SeqSearch_PAID(spdWorklist, Trim(mAdoRs.Fields("ORDT").Value & ""), Trim(mAdoRs.Fields("PAID")), Trim(mAdoRs.Fields("ORQN")), colORQN)

                    If pGrid_Point = 0 Then
                        pGrid_Point = SeqNullSearch(spdWorklist, Trim(mAdoRs.Fields("ORDT")), colORDT)
                        If pGrid_Point = 0 Then .maxrows = .maxrows + 1: pGrid_Point = .maxrows
                    End If
                    
                    .SetText colBANO, pGrid_Point, "0"
                    .SetText colORDT, pGrid_Point, mAdoRs("ORDT").Value & ""
                    .SetText colORQN, pGrid_Point, mAdoRs("ORQN").Value & ""
                    .SetText colPANM, pGrid_Point, Trim(mAdoRs("PANM").Value & "")
                    .SetText colPAID, pGrid_Point, Trim(mAdoRs("PAID").Value & "")
                    .SetText colOIFL, pGrid_Point, Trim(mAdoRs("OIFL").Value & "")
                    .SetText colSENO, pGrid_Point, Trim(mAdoRs("SENO").Value & "")
                    .SetText colSEXS, pGrid_Point, Trim(mAdoRs("SEXS").Value & "")
                    .SetText colAGES, pGrid_Point, Trim(mAdoRs("AGES").Value & "")
                    .SetText colNWNO, pGrid_Point, Trim(mAdoRs("NWNO").Value & "")
                    
                    .Row = pGrid_Point: .Col = colBANO: .ForeColor = HNC_Black
                                        .Col = colORDT: .ForeColor = HNC_Black
                                        .Col = colORQN: .ForeColor = HNC_Black
                                        .Col = colPANM: .ForeColor = HNC_Black
                                        .Col = colPAID: .ForeColor = HNC_Black
                                        .Col = colOIFL: .ForeColor = HNC_Black
                                        .Col = colSENO: .ForeColor = HNC_Black
                                        .Col = colSEXS: .ForeColor = HNC_Black
                                        .Col = colAGES: .ForeColor = HNC_Black
                                        .Col = colNWNO: .ForeColor = HNC_Black
                    
                End If

                strBarno = Trim(mAdoRs.Fields("ORDT").Value & "") & Trim(mAdoRs.Fields("PAID").Value & "")
                'strBarno = Trim(mAdoRs.Fields("ORDT").Value & "") & Trim(mAdoRs.Fields("PAID").Value & "") & Trim(mAdoRs.Fields("ORQN").Value & "")
                mAdoRs.MoveNext
            Next
            
            If blt = False Then
                .Row = pGrid_Point
                .Action = ActionDeleteRow
                .maxrows = .maxrows - 1
            End If
        End With
    End If
    
    Set mAdoRs = Nothing
    
    spdWorklist.Row = 1
    spdWorklist.Col = 1
    spdWorklist.Action = ActionActiveCell
        
Exit Sub

ErrorTrap:
    Set mAdoRs = Nothing
    Call ErrMsgProc(CallForm)
End Sub

Private Sub cmdACK_Click()
'
'    Call COM_OUTPUT(charCOM_Convert(COM_ACK))
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
    
    f_strJOB_FLAG = "1"
    f_intSampleNo = 0
    Or_Seq = 1
    
    Erase f_strBarno:   Erase f_strTest
    f_intCnt = 0:       f_intIdx = 0
        
    With spdWorklist
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

Private Function f_subSet_RefVal(ByVal strORCD As String, Optional ByVal strRSLT As String, Optional ByVal strSex As String, Optional ByVal strAGE As String) As String
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim stryy, strmm, strdd, strDate  As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_RefVal() As ADODB.Recordset"
    
    strRSLT = Replace(strRSLT, "<", "")
    strRSLT = Replace(strRSLT, ">", "")
    f_subSet_RefVal = " "
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    f_subSet_RefVal = ""
    If strAGE <> "" Then
        If strAGE <= 7 Then
            sqlDoc = "Select YMAX as MAX, YMIN as MIN "
        Else
            If strSex = "M" Then
                     sqlDoc = "Select MMAX as MAX, MMIN as MIN "
            Else
                     sqlDoc = "Select WMAX as MAX, WMIN as MIN "
            End If
        End If
    Else
        sqlDoc = "Select MMAX as MAX, MMIN as MIN "
    End If
    
    sqlDoc = sqlDoc & "  From LABMAST"
    sqlDoc = sqlDoc & " Where ORCD =  '" & strORCD & "'"

    Set AdoRs_ORACLE = New ADODB.Recordset
    
    AdoRs_ORACLE.CursorLocation = adUseClient
    AdoRs_ORACLE.Open sqlDoc, AdoCn_ORACLE
    
    If AdoRs_ORACLE.RecordCount = 0 Then
        f_subSet_RefVal = " "
        Set AdoRs_ORACLE = Nothing
        Exit Function
    Else
        If IsNumeric(strRSLT) And IsNumeric(AdoRs_ORACLE.Fields("MAX")) And IsNumeric(AdoRs_ORACLE.Fields("MIN")) Then
            If Val(strRSLT) > Val(AdoRs_ORACLE.Fields("MAX")) Then
                f_subSet_RefVal = "H"
            ElseIf Val(strRSLT) < Val(AdoRs_ORACLE.Fields("MIN")) Then
                f_subSet_RefVal = "L"
            Else
                f_subSet_RefVal = " "
            End If
        Else
            f_subSet_RefVal = " "
        End If
    End If

    Set AdoRs_ORACLE = Nothing
    
Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    
    Call ErrMsgProc(CallForm)
     
End Function


Private Sub cmdAppend_Click(Index As Integer)
'''
'''    Dim adoRS   As New ADODB.Recordset
'''    Dim sqlDoc  As String
'''
'''    Dim varTmp  As Variant, strErrMsg   As String
'''    Dim strSampleno()   As String
'''    Dim strOrdcd()      As String, strRstval  As String, intCnt       As Integer
'''    Dim strTmp1()       As String, strTmp2()    As String
'''    Dim intPos          As String, strTestcd    As String, strTestRst   As String
'''    Dim strTestnm       As String
'''    Dim strRef          As String
'''    Dim strUnit         As String
'''    Dim strOrdLst()     As String, strPid()    As String, strPnm() As String
'''
'''    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
'''    Dim itemX   As ListItem
'''    Dim objSpd  As vaSpread
'''    Dim sqlRet  As Integer
'''    Dim flgSave As Boolean
'''    Dim SaveGbn As Integer
'''
'''    Dim strDate As String
'''    Dim strBarno As String
'''    Dim strSPnm As String
'''    Dim strSPid As String
'''    Dim strChartNo As String
'''    Dim strEqpCd As String
'''    Dim strORDT, strORQN, strPANM, strPAID, strOIFL, strSENO, strSEXS, strAGES, strNWNO, strORCD As String
'''    Dim strRefVal As String
'''
'''    CallForm = "frmComm - Private Sub cmdAppend_Click()"
'''
'''On Error GoTo ErrorRoutine
'''
'''    Me.MousePointer = 11
'''
'''    If Index = 0 Then
'''        Set objSpd = spdResult1
'''    Else
'''        Set objSpd = spdResult2
'''    End If
'''
'''    With objSpd
'''        For intRow = 1 To .maxrows
'''            .GetText colORDT, intRow, varTmp:    strORDT = Trim$(varTmp)
'''            .GetText colORQN, intRow, varTmp:    strORQN = Trim$(varTmp)
'''            .GetText colPANM, intRow, varTmp:    strPANM = Trim$(varTmp)
'''            .GetText colPAID, intRow, varTmp:    strPAID = Trim$(varTmp)
'''            .GetText colOIFL, intRow, varTmp:    strOIFL = Trim$(varTmp)
'''            .GetText colSENO, intRow, varTmp:    strSENO = Trim$(varTmp)
'''            .GetText colSEXS, intRow, varTmp:    strSEXS = Trim$(varTmp)
'''            .GetText colAGES, intRow, varTmp:    strAGES = Trim$(varTmp)
'''            .GetText colNWNO, intRow, varTmp:    strNWNO = Trim$(varTmp)
'''
'''            .GetText colBANO, intRow, varTmp
'''
'''            If strORDT = "" Then Exit For
'''
'''            intCnt = 0: Erase strOrdcd ': Erase strRstval
'''
'''            If Trim$(varTmp) = "1" Then
'''                For intCol = colSQNO + 1 To .MaxCols
'''                    .GetText intCol, intRow, varTmp
'''                        If Trim$(varTmp) <> "" Then
'''                            .GetText intCol, 0, varTmp
'''                            Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
'''                            If Not itemX Is Nothing Then
'''                                .GetText intCol, intRow, varTmp: strRstval = varTmp
'''                                strTestcd = itemX.ListSubItems(1)
'''                                intPos = InStr(strTestcd, ",")
'''                                strEqpCd = itemX.Text
'''
'''                                If strEqpCd <> "" Then
'''                                    '-- H/L 판정
'''                                    strRefVal = f_subSet_RefVal(strEqpCd, strRstval, strSEXS, strAGES)
'''
'''                                    '-- 최근 접수번호[ORQN] 찾기
'''                                    sqlDoc = "Select ORQN,TRANSDT,TRANSTM from INTERFACE004"
'''                                    sqlDoc = sqlDoc & " Where ORDT = '" & strORDT & "'"
'''                                    sqlDoc = sqlDoc & "   And PAID = '" & strPAID & "'"
'''                                    sqlDoc = sqlDoc & "   And OIFL = '" & strOIFL & "'"
'''                                    sqlDoc = sqlDoc & "   And EQPCD = '" & strEqpCd & "'"
'''                                    sqlDoc = sqlDoc & " Order By TRANSDT, TRANSTM desc "
'''
'''                                    adoRS.CursorLocation = adUseClient
'''                                    adoRS.Open sqlDoc, AdoCn_Jet
'''                                    If adoRS.RecordCount > 0 Then
'''                                        adoRS.MoveFirst
'''                                    End If
'''
'''                                    Do While Not adoRS.EOF
'''                                        Debug.Print adoRS.Fields("TRANSDT") & adoRS.Fields("TRANSTM")
'''                                        If Trim(adoRS.Fields("ORQN")) <> "" Then
'''                                            strORQN = Trim(adoRS.Fields("ORQN"))
'''                                            Exit Do
'''                                        End If
'''                                        adoRS.MoveNext
'''                                    Loop
'''
'''                                    Set adoRS = Nothing
'''
'''                                    '-- 서버저장
'''                                    sqlDoc = " Update LRESULT"
'''                                    sqlDoc = sqlDoc & "   Set RSFL = 'Y',"
'''                                    sqlDoc = sqlDoc & "       RSLT = '" & strRstval & "',"
'''                                    sqlDoc = sqlDoc & "       HLFL = '" & strRefVal & "',"
'''                                    sqlDoc = sqlDoc & "       RSDT = '" & Format(Now, "YYYYMMDD") & "',"
'''                                    sqlDoc = sqlDoc & "       RSID = '" & CurrUser.CuUserID & "'"
'''                                    sqlDoc = sqlDoc & " Where PAID = '" & strPAID & "'"
'''                                    sqlDoc = sqlDoc & "   And NWNO = " & strNWNO
'''                                    sqlDoc = sqlDoc & "   And ORDT = '" & strORDT & "'"
'''                                    sqlDoc = sqlDoc & "   And ORQN = " & strORQN
'''                                    sqlDoc = sqlDoc & "   And OIFL = '" & strOIFL & "'"
'''                                    sqlDoc = sqlDoc & "   And ORCD = '" & strEqpCd & "'"
'''                                    sqlDoc = sqlDoc & "   And OKFL <> 'Y' "   '-- 결과확정유무
'''
'''                                    AdoCn_ORACLE.Execute sqlDoc
'''
'''                                    '-- 로컬 업데이트
'''                                    sqlDoc = "Update INTERFACE004" & _
'''                                             "   Set SERVERGBN = 'Y'" & _
'''                                             " Where PAID = '" & strPAID & "'" & _
'''                                             "   And NWNO = " & strNWNO & _
'''                                             "   And ORDT = '" & strORDT & "'" & _
'''                                             "   And ORQN = " & strORQN & _
'''                                             "   And OIFL = '" & strOIFL & "'" & _
'''                                             "   And EQPCD = '" & strEqpCd & "'"
'''
'''                                    AdoCn_Jet.Execute sqlDoc
'''
'''                                    Debug.Print sqlDoc
'''
'''
'''                                    lblStatus.Caption = "저장 성공!!"
'''
'''                                    .Row = intRow: .Col = colBANO: .Value = 0
'''                                                   .Col = colORDT: .BackColor = HNC_Cyan
'''                                                   .Col = colORQN: .BackColor = HNC_Cyan
'''                                                   .Col = colPANM: .BackColor = HNC_Cyan
'''                                                   .Col = colPAID: .BackColor = HNC_Cyan
'''                                                   .Col = colOIFL: .BackColor = HNC_Cyan
'''                                                   .Col = colSENO: .BackColor = HNC_Cyan
'''                                                   .Col = colSEXS: .BackColor = HNC_Cyan
'''                                                   .Col = colAGES: .BackColor = HNC_Cyan
'''                                                   .Col = colNWNO: .BackColor = HNC_Cyan
'''
'''                                End If
'''                            Set itemX = Nothing
'''                        End If
'''                    End If
'''                Next
'''            End If
'''        Next
'''    End With
'''
'''    Me.MousePointer = 0
'''
'''    If lblStatus.Caption = "저장 성공!!" Then
'''        MsgBox "▒ EMR SERVER에 결과를 Upload 완료되었습니다. ▒      " & vbCrLf & vbCrLf & "     LIS 결과조회 화면에서 결과를 확인 하십시요..  ", vbInformation, App.Title
'''    End If
'''
'''    Exit Sub
'''ErrorRoutine:
'''
'''    Set AdoRs_SQL = Nothing
'''
'''    Set itemX = Nothing
'''
'''    Me.MousePointer = 0
'''    Call ErrMsgProc(CallForm)
    
   
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String

    Dim varTmp  As Variant, strErrMsg   As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval  As String, intCnt       As Integer
    Dim strTmp1()       As String, strTmp2()    As String
    Dim intPos          As String, strTestcd    As String, strTestRst   As String
    Dim strTestnm       As String
    Dim strRef          As String
    Dim strUnit         As String
    Dim strOrdLst()     As String, strPid()    As String, strPnm() As String

    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
    Dim itemX   As ListItem
    Dim objSpd  As vaSpread
    Dim sqlRet  As Integer
    Dim flgSave As Boolean
    Dim SaveGbn As Integer
    
    Dim strDate As String
    Dim strTime As String
    
    Dim strBarno As String
    Dim strSPnm As String
    Dim strSPid As String
    Dim strChartNo As String
    Dim strEqpCd As String
    Dim strORDT, strORQN, strPANM, strPAID, strOIFL, strSENO, strSEXS, strAGES, strNWNO, strORCD As String
    Dim strRefVal As String
    Dim pName As String
    Dim pNo As String
    
    CallForm = "frmComm - Private Sub cmdAppend_Click()"

On Error GoTo ErrorRoutine

    Me.MousePointer = 11

    If Index = 0 Then
        Set objSpd = spdResult1
    Else
        Set objSpd = spdResult2
    End If

    With objSpd
        For intRow = 1 To .maxrows
            .GetText colORDT, intRow, varTmp:    strORDT = Trim$(varTmp)
            .GetText colORQN, intRow, varTmp:    strORQN = Trim$(varTmp)
            .GetText colPANM, intRow, varTmp:    strPANM = Trim$(varTmp)
            .GetText colPAID, intRow, varTmp:    strPAID = Trim$(varTmp): strBarno = strPAID
            .GetText colOIFL, intRow, varTmp:    strOIFL = Trim$(varTmp)
            .GetText colSENO, intRow, varTmp:    strSENO = Trim$(varTmp)
            .GetText colSEXS, intRow, varTmp:    strSEXS = Trim$(varTmp)
            .GetText colAGES, intRow, varTmp:    strAGES = Trim$(varTmp)
            .GetText colNWNO, intRow, varTmp:    strNWNO = Trim$(varTmp)

            .GetText colBANO, intRow, varTmp

            If strPAID = "" Then Exit For

            intCnt = 0: Erase strOrdcd ': Erase strRstval
            
            If Trim$(varTmp) = "1" Then
                For intCol = 10 To .MaxCols
                    strDate = Format$(Now, "YYYYMMDD"):    strTime = Format$(Now, "HHMMSS")
                    .GetText intCol, intRow, varTmp
                        If Trim$(varTmp) <> "" Then
                            .GetText intCol, 0, varTmp
                            Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                            If Not itemX Is Nothing Then
                                .GetText intCol, intRow, varTmp: strRstval = varTmp
                                strTestcd = itemX.ListSubItems(1)
                                intPos = InStr(strTestcd, ",")
                                strEqpCd = itemX.Text
                                                    
                                If strEqpCd <> "" Then
                                    '-- 로컬저장
                                    sqlDoc = "Update INTERFACE003" & _
                                             "   set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRefVal & "'" & _
                                             " where SPCNO   = '" & strBarno & "'" & _
                                             "   and EQPNUM  = '" & itemX.tag & "'" & _
                                             "   and TRANSDT = '" & strDate & "'" & _
                                             "   and TRANSTM = '" & strTime & "'"
                                    AdoCn_Jet.Execute sqlDoc
                                    
                                    sqlDoc = "insert into INTERFACE003(" & _
                                             "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                             "    values( '" & strBarno & "', '" & strEqpCd & "', '" & itemX.tag & "'," & _
                                             "            '" & strDate & "', '" & strTime & "'," & _
                                             "            '" & strRstval & "', '" & strRefVal & "'," & _
                                             "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
                                    AdoCn_Jet.Execute sqlDoc
                                    
                                    '   3-1. 검사정보 MASTER
                                             sqlDoc = "UPDATE JAIN_SCP.SCPRST41 SET "
                                    sqlDoc = sqlDoc & "       SCP41TSTDAT = '" & Format(Now, "YYYYMMDD") & "'," '결과일자 => YYYYMMDD"
                                    sqlDoc = sqlDoc & "       SCP41SNDYN  = 'N',"                               '고정값 : 'N'
                                    sqlDoc = sqlDoc & "       SCP41RSTYN  = 'Y',"                               '고정값 : 'Y'
                                    sqlDoc = sqlDoc & "       SCP41TSTUID = '" & CurrUser.CuUserID & "'"        '검사자사번
                                    sqlDoc = sqlDoc & " WHERE SCP41SPMNO2 = '" & strBarno & "'"                 '바코드번호
                                    
                                    AdoCn_ORACLE.Execute sqlDoc
                                    
                                    '   3-2. 검사정보 DETAIL
                                             sqlDoc = "UPDATE JAIN_SCP.SCPRST42 SET "
                                    sqlDoc = sqlDoc & "       SCP42TSTDAT = '" & Format(Now, "YYYYMMDD") & "'," '결과일자 => YYYYMMDD"
                                    sqlDoc = sqlDoc & "       SCP42RSTCD  = 'N',"                               '결과형식 => 숫자 : 'N', 문자 : 'X', 장문 : 'R'
                                    sqlDoc = sqlDoc & "       SCP42RESULT = '" & strRstval & "'"                '결과값
                                    sqlDoc = sqlDoc & " WHERE SCP42SPMNO2 = '" & strBarno & "'"                 '바코드번호
                                    sqlDoc = sqlDoc & "   AND SCP42SUGACD = '" & strEqpCd & "'"              '수가코드

                                    AdoCn_ORACLE.Execute sqlDoc
                                    
                                    
                                    lblStatus.Caption = "저장 성공!!"
                                                                                                                    
                                    .Row = intRow: .Col = colBANO: .Value = 0
                                                   .Col = colORDT: .BackColor = HNC_Cyan
                                                   .Col = colORQN: .BackColor = HNC_Cyan
                                                   .Col = colPANM: .BackColor = HNC_Cyan
                                                   .Col = colPAID: .BackColor = HNC_Cyan
                                                   .Col = colOIFL: .BackColor = HNC_Cyan
                                                   .Col = colSENO: .BackColor = HNC_Cyan
                                                   '.Col = colSEXS: .BackColor = HNC_Cyan
                                                   '.Col = colAGES: .BackColor = HNC_Cyan
                                                   '.Col = colNWNO: .BackColor = HNC_Cyan
                                                                        
                                End If
                            Set itemX = Nothing
                        End If
                    End If
                Next
            End If
        Next
    End With
    
    Me.MousePointer = 0
    
    If lblStatus.Caption = "저장 성공!!" Then
        MsgBox "▒ EMR SERVER에 결과를 Upload 완료되었습니다. ▒      " & vbCrLf & vbCrLf & "     LIS 결과조회 화면에서 결과를 확인 하십시요..  ", vbInformation, App.Title
    End If
    
    Exit Sub
ErrorRoutine:

    Set AdoRs_SQL = Nothing

    Set itemX = Nothing

    Me.MousePointer = 0
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
    
    Call COM_OUTPUT(charCOM_Convert(COM_ENQ))

End Sub

Private Sub cmdRstQuery_Click()
'
'    Dim adoRS   As New ADODB.Recordset
'    Dim sqlDoc  As String, intRet   As Integer
'
'    Dim strSpcno    As String
'    Dim intRow      As Integer, intCol  As Integer
'    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
'
'    Dim itemX       As ListItem
'
'    intRow = 0
'    With spdResult2
'        .maxrows = 1
'        .Col = 1:   .Col2 = .MaxCols
'        .Row = 1:   .Row2 = .maxrows
'        .BlockMode = True
'        .Action = ActionClearText
'        .BlockMode = False
'    End With
'
'    sqlDoc = "Select ORDT, ORQN, PAID, OIFL, SENO, NWNO, ORCD, TRANSDT, TRANSTM, EqpCD, RSTVAL, REFVAL, SERVERGBN, PANM, SEX, AGE,EQPNUM" & _
'             "  From INTERFACE004" & _
'             " Where TRANSDT >= '" & mskRstDate.Text & "'"
'
'    If cboRstgbn(1).ListIndex = 0 Then
'        sqlDoc = sqlDoc & "   And (SERVERGBN = '' or SERVERGBN = 'N')"
'    ElseIf cboRstgbn(1).ListIndex = 1 Then
'        sqlDoc = sqlDoc & "   And SERVERGBN = 'Y'"
'    End If
'    sqlDoc = sqlDoc & " Order By ORDT, TRANSDT,TRANSTM"
'
'    adoRS.CursorLocation = adUseClient
'    adoRS.Open sqlDoc, AdoCn_Jet
'    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
'    Do While Not adoRS.EOF
'        With spdResult2
'            If strSpcno <> Trim$(adoRS("ORDT") & "") & Trim$(adoRS("ORQN") & "") & Trim$(adoRS("PAID") & "") Then
'                intRow = intRow + 1
'                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
'                .SetText 1, intRow, "1"
'                .SetText 2, intRow, adoRS("ORDT").Value & ""
'                .SetText 3, intRow, adoRS("ORQN").Value & ""
'                .SetText 4, intRow, Trim(adoRS("PANM").Value & "")
'                .SetText 5, intRow, Trim(adoRS("PAID").Value & "")
'                .SetText 6, intRow, Trim(adoRS("OIFL").Value & "")
'                .SetText 7, intRow, Trim(adoRS("SENO").Value & "")
'                .SetText 8, intRow, Trim(adoRS("SEX").Value & "")
'                .SetText 9, intRow, Trim(adoRS("AGE").Value & "")
'                .SetText 10, intRow, Trim(adoRS("NWNO").Value & "")
'                .SetText 11, intRow, Format(Trim(adoRS("TRANSDT").Value & ""), "####-##-##")
'            End If
'            strSpcno = Trim$(adoRS("ORDT") & "") & Trim$(adoRS("ORQN") & "") & Trim$(adoRS("PAID") & "")
'            Set itemX = lvwCuData.FindItem(Trim$(adoRS("EQPNUM") & ""), lvwTag, , lvwWhole)
'            If Not itemX Is Nothing Then
'                intCol = itemX.Index + 11
'                .SetText intCol, intRow, Trim$(adoRS("RSTVAL")) & ""
'                .Col = intCol:  .Row = intRow:  .ForeColor = IIf(Trim$(adoRS("REFVAL") & "") <> "", vbRed, vbBlack)
'            End If
'        End With
'        adoRS.MoveNext
'    Loop
'    adoRS.Close:    Set adoRS = Nothing
    

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
        If strSpcno <> Trim$(adoRS(0) & "") + Trim$(adoRS(6) & "") Then
                intRow = intRow + 1
                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                .SetText 1, intRow, "1"
                .SetText 2, intRow, Trim$(adoRS(3) & "")
                .SetText 3, intRow, Trim$(adoRS(0) & "")
                .SetText 6, intRow, Trim$(adoRS(8) & "")
                .SetText 7, intRow, Trim$(adoRS(9) & "")
                '.SetText .MaxCols, intRow, Trim$(adoRS(6) & "")
            End If
                strSpcno = Trim$(adoRS(0) & "") + Trim$(adoRS(6) & "")
                Set itemX = lvwCuData.FindItem(Trim$(adoRS(7) & ""), lvwTag, , lvwWhole)
                If Not itemX Is Nothing Then
                    intCol = itemX.Index + 8
                    .SetText intCol, intRow, Trim$(adoRS(4)) & ""
                    .Col = intCol:  .Row = intRow:  .ForeColor = IIf(Trim$(adoRS(5) & "") <> "", vbRed, vbBlack)
                End If
        End With
        adoRS.MoveNext
    Loop
'    spdResult2.MaxCols = spdResult2.MaxCols - 1
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub cmdSel_Click(Index As Integer)

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    If Index = 2 Or Index = 3 Then
        With spdResult1
            For intRow = 1 To .maxrows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 0, "1", "")
            Next
        End With
    Else
        With spdWorklist
            For intRow = 1 To .maxrows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 0, "1", "")
            Next
        End With
    End If
    
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
                .Col = 0:       .Text = Trim(sAdd + Val(sNo))
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
    
    Set AdoRs_SQL = New ADODB.Recordset
    AdoRs_SQL.CursorLocation = adUseClient
    AdoRs_SQL.Open AdoCn_SQL.Execute("Exec AP_INF_S_GetCoda '" & INS_CODE & "', '" & strBarcode & "'", sqlRet)
    
    If sqlRet = 0 Then
        Set f_subSet_TestList = Nothing
        MsgBox strBarcode & "-해당검체는 검사가 완료되었습니다.", vbOKOnly + vbExclamation
        RecordChk = False
        Exit Function
    Else
        Set f_subSet_TestList = AdoRs_SQL
        RecordChk = True
    End If

    Set AdoRs_SQL = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    Call ErrMsgProc(CallForm)
    
End Function

Private Sub cmdWordQuery_Click()
'    On Error GoTo ErrRoutine
'    CallForm = "frmInterface - Privete sub cmdWorkQuery_Click()"
'
'    Dim strKeyno    As String
'    Dim strOrdcd()  As String, strPid() As String, strPnm() As String, strBarno()   As String
'    Dim strTestcd() As String, strTPid()    As String, strTPnm() As String
'    Dim strEqpCd    As String
'    Dim intRow  As String, intIdx  As Integer, intCol   As Integer
'    Dim itemX   As ListItem
'
'    '-- WorkList조회
'    Set mAdoRs = f_subSet_WorkList(mskOrdDate.Text)
'
'    If RecordChk = False Then
'        Exit Sub
'    End If
'
''    With spdWorkList
''        .maxrows = 14
''        .Col = 1:   .Col2 = .MaxCols
''        .Row = 1:   .Row2 = .maxrows
''        .BlockMode = True
''        .Action = ActionClearText
''        .BlockMode = False
''        .RowHeight(-1) = 12
''    End With
'
'    intRow = 0
'    Do Until mAdoRs.EOF
'        intIdx = 0
'        With spdResult1
'            If strKeyno <> mAdoRs.Fields("EXAM_NO") Then
''                intRow = SeqNullSearch(spdResult1, "", 1)
''                If intRow = "0" Then
''                    .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
''                    intRow = .maxrows
''                Else
''                    intRow = intRow + 1
''                End If
'                intRow = intRow + 1
'                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
'
'                .SetText 1, intRow, "1"
'                .SetText 2, intRow, Trim(mAdoRs("REQUEST_DATE")) & ""
'                '.SetText 3, intRow, Trim(cboRegGbn.Text) & ""
'                .SetText 4, intRow, Trim(mAdoRs("PERSON_NAME")) & ""
'                .SetText 5, intRow, Trim(mAdoRs("EXAM_NO")) & ""
'                .SetText 6, intRow, Trim(mAdoRs("CHART_NO")) & ""
'                '.SetText 6, intRow, Trim(mAdoRs("COMPANY_NAME"))
'
'                '-- 검사항목조회
''                blnFlag = False
'                Set mAdoRs1 = f_subSet_TestList(mAdoRs.Fields("EXAM_NO"))
'                If Len(mAdoRs.Fields("EXAM_NO")) > 0 Then
'                    Do Until mAdoRs1.EOF
'                        strEqpCd = f_funGet_CODE(Trim(mAdoRs1("EXAM_CODE")))
'                        Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
'                        If Not itemX Is Nothing Then
''                            blnFlag = True
'                            spdResult1.Row = intRow
'                            spdResult1.Col = itemX.Index + 6
'                            spdResult1.BackColor = &HC6FEFF '&H80C0FF
'                            DoEvents
'                        End If
'                        mAdoRs1.MoveNext
'                    Loop
'                End If
'
'            End If
'            strKeyno = mAdoRs("EXAM_NO")
'        End With
'        intIdx = intIdx + 1
'        mAdoRs.MoveNext
'    Loop
'    Exit Sub
'
'ErrRoutine:
'
'    Call ErrMsgProc(CallForm)

End Sub

Private Sub cmdWorkList_Click()

    Dim varTmp  As Variant
    Dim intRow1 As Integer, intRow2 As Integer
    Dim intIdx  As Integer
    Dim Rev     As Long
    Dim Test_Cd() As String, strPid()   As String, strPnm() As String
    Dim itemX As ListItem
    Dim blnFlag As Boolean
    Dim strBarno    As String, strSPid  As String, strSPnm   As String, strChartNo As String, strSex As String
    Dim strWDate As String
    Dim strEqpCd    As String
    Dim tmpDate     As String
    Dim strORDT, strORQN, strPANM, strPAID, strOIFL, strSENO, strSEXS, strAGES, strNWNO, strORCD, strSQNO As String
    
    blnFlag = False
    
    With spdWorklist
        For intRow1 = 1 To .maxrows
            .GetText 1, intRow1, varTmp
            If Trim$(varTmp) = "1" Then
                    '-- 처방일자,처방일련번호,환자명,환자번호,입외구분,일련번호,성별,나이,내원번호,처방코드
                .GetText colORDT, intRow1, varTmp:    strORDT = Trim$(varTmp)
                .GetText colORQN, intRow1, varTmp:    strORQN = Trim$(varTmp)
                .GetText colPANM, intRow1, varTmp:    strPANM = Trim$(varTmp)
                .GetText colPAID, intRow1, varTmp:    strPAID = Trim$(varTmp)
                .GetText colOIFL, intRow1, varTmp:    strOIFL = Trim$(varTmp)
                .GetText colSENO, intRow1, varTmp:    strSENO = Trim$(varTmp)
                .GetText colSEXS, intRow1, varTmp:    strSEXS = Trim$(varTmp)
                .GetText colAGES, intRow1, varTmp:    strAGES = Trim$(varTmp)
                .GetText colNWNO, intRow1, varTmp:    strNWNO = Trim$(varTmp)
                .GetText colSQNO, intRow1, varTmp:    strSQNO = Trim$(varTmp)
                
                .Row = intRow1: .Col = colBANO: .ForeColor = HNC_Red
                                .Col = colORDT: .ForeColor = HNC_Red
                                .Col = colORQN: .ForeColor = HNC_Red
                                .Col = colPANM: .ForeColor = HNC_Red
                                .Col = colPAID: .ForeColor = HNC_Red
                                .Col = colOIFL: .ForeColor = HNC_Red
                                .Col = colSENO: .ForeColor = HNC_Red
                                .Col = colSEXS: .ForeColor = HNC_Red
                                .Col = colAGES: .ForeColor = HNC_Red
                                .Col = colNWNO: .ForeColor = HNC_Red
                
                intRow2 = f_funGet_SpreadRow_PAID(spdResult1, colORQN, strORDT, strORQN, strPAID)
                If intRow2 < 1 Then
                    intRow2 = f_funGet_SpreadRow(spdResult1, colORDT, "")
                    If intRow2 < 1 Then
                        spdResult1.maxrows = spdResult1.maxrows + 1
                        spdResult1.RowHeight(spdResult1.maxrows) = 13
                        intRow2 = spdResult1.maxrows
                    End If

                    blnFlag = False
                    
                    Set mAdoRs = f_subSet_WorkList_Barcode(strORDT, strPAID, strORQN)

                    If Len(strPAID) > 0 And Not mAdoRs Is Nothing Then
                        Do Until mAdoRs.EOF
                            strEqpCd = f_funGet_CODE(Trim(mAdoRs("ORCD").Value))
                            
                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then
                                blnFlag = True
                                spdResult1.Row = intRow2
                                spdResult1.Col = itemX.Index + colSQNO
                                spdResult1.BackColor = &HC6FEFF '&H80C0FF
                                spdResult1.ForeColor = vbWhite '&H80C0FF
                                spdResult1.Text = Trim(mAdoRs("ORQN").Value) & ""
                    
                                DoEvents
                            End If
                            mAdoRs.MoveNext
                        Loop
                    End If
                    If blnFlag = True Then
                        spdResult1.SetText colORDT, intRow2, strORDT
                        spdResult1.SetText colORQN, intRow2, strORQN
                        spdResult1.SetText colPANM, intRow2, strPANM
                        spdResult1.SetText colPAID, intRow2, strPAID
                        spdResult1.SetText colOIFL, intRow2, strOIFL
                        spdResult1.SetText colSENO, intRow2, strSENO
                        spdResult1.SetText colSEXS, intRow2, strSEXS
                        spdResult1.SetText colAGES, intRow2, strAGES
                        spdResult1.SetText colNWNO, intRow2, strNWNO
                        spdResult1.SetText colSQNO, intRow2, strSQNO
                        
                        spdResult1.Row = intRow2:
                        spdResult1.Col = colSQNO:
                        spdResult1.ForeColor = HNC_Red
                    Else
                        spdResult1.maxrows = spdResult1.maxrows - 1
                    End If
                End If
                
                .SetText 1, intRow1, ""
            End If
        Next
    End With
                
End Sub

Private Function f_funGet_SpreadRow_PAID(ByVal objSpd As vaSpread, ByVal intCol As Integer, _
                                    ByVal strPara1 As String, ByVal strPara2 As String, ByVal strPara3 As String) As Integer

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    f_funGet_SpreadRow_PAID = 0
    
    With objSpd
        For intRow = 1 To .maxrows
            .GetText 2, intRow, varTmp
            If Trim$(varTmp) = strPara1 Then
                .GetText 5, intRow, varTmp
                If Trim$(varTmp) = strPara3 Then
                    f_funGet_SpreadRow_PAID = intRow
                    Exit For
                End If
            End If
        Next
    End With
    
End Function

Private Sub comEQP_OnComm()
    
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim strDta      As String
   
    Select Case comEQP.CommEvent
        Case comEvReceive
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If

            strDta = comEQP.Input
            Debug.Print strDta
            Call ComReceive(strDta)
                                        
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

'-----------------------------------------------------------------------------'
'   기능 : Session start, Exchange parameter 메시지 조회
'-----------------------------------------------------------------------------'
Public Function GetParameter() As String
    GetParameter = "Y~R @-#N1"
End Function

'-----------------------------------------------------------------------------'
'   기능 : Vitros 장비 CheckSum을 조회
'-----------------------------------------------------------------------------'
Public Function GetCheckSum(ByVal pMsg) As String
    Dim lngChkSum   As Long
    Dim lngTemp     As Long
    Dim i           As Long
    
    For i = 1 To Len(pMsg)
        lngChkSum = lngChkSum + Asc(Mid$(pMsg, i, 1))
    Next i
    
    lngTemp = (lngChkSum And 192) / 64
    lngChkSum = (lngChkSum + lngTemp) And 63
    GetCheckSum = Chr(lngChkSum + 32)
End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비로 ACK 메시지를 전송
'-----------------------------------------------------------------------------'
Private Sub SendAckPacket()
    Dim strRcvBuf   As String   '수신한 Data
    Dim strSeq      As String   '수신한 Seq
    Dim strType     As String   '수신한 Data Tytpe
    Dim strOutput   As String   '송신할 Data
    Dim blnFirst    As Boolean
    
    strRcvBuf = f_strBuffer 'objInt.objDicBuf.Fields("bufchar")
    strSeq = Mid$(strRcvBuf, 2, 1)
    strType = Mid$(strRcvBuf, 3, 1)

    Select Case strType
        Case "S"    'Send initiate packet (session start, exchange parameters)
            strOutput = strSeq & GetParameter
            strOutput = Chr(33 + Len(strOutput)) & strOutput
            strOutput = Chr(COM_SOH) & strOutput & GetCheckSum(strOutput) & vbCr
            objIntBufCnt = objIntBufCnt + 1
        
        Case "F"    'File header packet
            objIntstate = "F"
            strOutput = "#" & strSeq & "Y"
            strOutput = Chr(COM_SOH) & strOutput & GetCheckSum(strOutput) & vbCr
            objIntBufCnt = objIntBufCnt + 1
        
        Case "D"    'Data packet
            strOutput = "#" & strSeq & "Y"
            strOutput = Chr(COM_SOH) & strOutput & GetCheckSum(strOutput) & vbCr
            
            If objDicBuf <> "" Then
                objDicBuf = objDicBuf & Mid(strRcvBuf, 4)
            Else
                objDicBuf = objDicBuf & strRcvBuf
            End If
            'objDicBuf = objDicBuf & strRcvBuf
'            If objIntstate = "F" Then
'                objIntBufCnt = objIntBufCnt + 1
'                objIntstate = "D"
'            Else
'            End If
'                objDicBuf = ""
'                objDicBuf = objDicBuf & strRcvBuf
'            End If
                       
        Case "Z"    'End of file packet (EOF)
            strOutput = "#" & strSeq & "Y"
            strOutput = Chr(COM_SOH) & strOutput & GetCheckSum(strOutput) & vbCr
                   
            'Call EditRcvData
            Call ReceiveTheData(objDicBuf, fChannel(), spdResult1)
            objDicBuf = ""
'            Call objInt.ClearBuffer
            f_strBuffer = ""
            objIntBufCnt = 0
            
        Case "B"    'Break transmission packet (EOT.end session)
            strOutput = "#" & strSeq & "Y"
            strOutput = Chr(COM_SOH) & strOutput & GetCheckSum(strOutput) & vbCr
            
'            Call objInt.ClearBuffer
            f_strBuffer = ""
            objIntBufCnt = 0
    End Select
    
    comEQP.Output = strOutput
    Print #1, "[PC] " & strOutput
    f_strBuffer = ""
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로 오더정보 전송
'   인수
'       - pFlag : True(ACK수신), False(NAK수신)
'-----------------------------------------------------------------------------'
'Private Sub SendOrder(Optional ByVal pFlag As Boolean = True)
'    Dim strOutput As String     '송신할 데이터
'
'    '## NAK수신시 바로전에 보낸 오더정보를 재전송
'    'If pFlag = True Then
'        strOutput = GetOrder(spdResult1)
'    'Else
'    '    strOutput = objOrder.Order
'    'End If
'
'    comEQP.Output = strOutput
'    Debug.Print strOutput
'    Print #1, "[PC] " & strOutput
'End Sub


'-----------------------------------------------------------------------------'
'   기능 : 오더문자열 조회
'   인수 :
'       - pWorklist : tblWorklist
'   반환 : 오더문자열
'-----------------------------------------------------------------------------'
'Public Function GetOrder(ByRef pWorklist As Object) As String
'    Dim vBarNo      As Variant  'Spread의 바코드번호
'    Dim vIntBase    As Variant  'Spread의 장비기준 검사명
'    Dim strIntBase  As String   '장비기준 검사명
'    Dim strItems    As String   '송신할 검사명 문자열
'    Dim strBarNo    As String   '송신할 바코드번호
'    Dim strPtId     As String   '송신할 환자ID
'    Dim i           As Long
'    Dim j           As Long
'    Dim varTmp      As Variant
'
'    '## 검사항목 문자열 생성
'    'strBarNo = Format$(objRst.mBarNo, "0" & String$(IIS_SPCYY_LEN + IIS_SPCNO_LEN - 1, "#"))
'    '.Col = 5
'    'strBarNo = Format$(.Text, "0" & String$(2 + 9 - 1, "#"))
'    'objRst.mBarNo = strBarNo
'Dim itemX   As ListItem
'
'    With pWorklist
'        For i = 1 To .DataRowCnt
'            Call .GetText(5, i, vBarNo)
'            'If CStr(vBarNo) = strBarNo Then
'                Call .SetText(1, i, "1")
'
'                For j = 10 To .MaxCols
'
'                    .GetText j, 0, varTmp
''                    varTmp = "r-GPT"
'
'                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
'                    If Not itemX Is Nothing Then
'                        'If itemX.tag = strIntBase Then
'
'
'                        'Call .GetText(j, I, vIntBase)
'                        If Trim$(itemX.tag) = "" Then Exit For
'
'                        strIntBase = Format$(itemX.tag, "@@@@") & "1"
'                        strItems = strItems & strIntBase
'                        'End If
'                    End If
'                Next j
'                Exit For
'            'End If
'        Next i
'    End With
'
'    If strItems = "" Then   '## 장비로 검사할 항목이 없는경우
'        objRst.mOrder = COM_STX & "O " & objRst.mBarNo & objRst.mDiskNo & objRst.mPos & COM_ETB & COM_ETB & COM_ETX
'    Else '10!01.000
'        objRst.mOrder = Chr(COM_STX) & "O " & Format(vBarNo, "000000000000000") & "10!01.000" & strItems & Chr(COM_ETB) & _
'                 String$(12, "0") & Space(30) & Space(1) & String$(8, "0") & Space(60) & _
'                 Chr(COM_ETB) & Chr(COM_ETX)
'        GetOrder = objRst.mOrder
'    End If
'    GetOrder = objRst.mOrder
'
'
'End Function

Private Sub ComReceive(ByRef RecData As String)

    Dim EVMsg As String
    Dim ERMsg As String
    Dim ret   As Long
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
        
    Buffer = RecData
    Print #1, Buffer;
'    Print #1, vbNewLine & Format(Now, "yyyymmdd hh:mm:ss") & "[RX]" & Buffer & vbNewLine;
        
    lngBufLen = Len(Buffer)
    For i = 1 To lngBufLen
        Timer2.Enabled = False
        BufChar = Mid$(Buffer, i, 1)
        
        Select Case mIntLibPhase
            Case 1      '## STX 대기
                Select Case BufChar
                    Case STX
                        mIntLibBuffer = ""
                        mIntLibPhase = 2
                    Case mORDER.MT
                        If mIntLibstate = "I" Then
                            Call SendToken
                        End If
                    Case NAK
                        Call cmdInit_Click
                End Select
            Case 2      '## ETX 대기
                Select Case BufChar
                    Case ETX
                    
                        'Timer2.Enabled = False
                        
                        Call EditRcvData
                        
                        'Timer2.Interval = 60000
                        'Timer2.Enabled = True
                        
                        mIntLibPhase = 1
                    
                    Case NAK
                        Call cmdInit_Click
                    
                    Case Else
                        mIntLibBuffer = mIntLibBuffer + BufChar
                End Select
        End Select
    Next i
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Message Toggle 전송
'-----------------------------------------------------------------------------'
Private Sub SendMT()
    Dim strOutput As String     '송신할 데이터
    
    strOutput = mORDER.MT
    comEQP.Output = strOutput
'    Print #1, "[SendMT]" & strOutput & vbNewLine;
    Print #1, vbNewLine & Format(Now, "yyyymmdd hh:mm:ss") & "[SendMT]" & strOutput & vbNewLine;

    Timer2.Enabled = True
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
'    Dim objIntInfo   As clsIISIntInfo    '인터페이스 검체정보 클래스
'    Dim objIntNms    As clsIISIntNms     '장비별 검사항목 컬렉션 클래스


    Dim sqlDoc  As String, sqlRet   As Integer
    
    Dim varTmp      As Variant
    Dim strTmp      As String
    Dim intRow      As Long, intCol As Integer, intIdx  As Integer
    Dim strRstval   As String, strRefVal    As String
    Dim sSeq, sCol, strBarno As String
    Dim strTime  As String, strDate  As String, strDate1 As String
    Dim strSeqno    As String
    Dim ii          As Integer
    Dim jj          As Integer
    
    Dim strOrdLst() As String, strPid() As String, strPnm() As String
    Dim intRet      As Integer
    Dim Channel_No  As String       ' 문자형 변수
    Dim Result As String
    Dim itemX   As ListItem
    Dim iRsltcnt As Integer
    Dim pDoCount As Integer
    Dim tmpChannel As String
    Dim tmpSeqno1 As Variant
    
    Dim pName   As String
    Dim pNo     As String
    Dim strID, strSampleno, strDept, strDept1, strPidNo  As String
    Dim strSeqTmpno, strOrdDate    As String
    Dim strSerial As String
    Dim strRorder As String
    
    Dim strRcvBuf    As String   '수신한 Data
    Dim strMT        As String   '수신한 MT
    Dim strIDCode    As String   '수신한 ID Code
'    Dim strBarno     As String   '수신한 BarNo
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strIntResult As String   '수신한 검사결과
    Dim strErrCd     As String   '수신한 에러정보
    Dim strResult    As String   'LIS 결과
    Dim strOutput    As String   '송신할 데이터
    Dim strTemp      As String
    Dim i            As Long
    Dim pGrid_Point As Integer
    Dim strORDT, strORQN, strPANM, strPAID, strOIFL, strSENO, strSEXS, strAGES, strNWNO, strORCD As String
    Dim strEqpCd As String
    Dim strState As String
    Dim intOrdCnt As Integer
    Dim strOrd(20) As String
    Dim intCnt As Integer
    Dim s As Integer
    Dim strEqpCds As String
    
    intCnt = 0
    strRcvBuf = mIntLibBuffer
    mORDER.MT = Mid$(strRcvBuf, 1, 1)
    strIDCode = Mid$(strRcvBuf, 2, 1)
    strState = ""
    
    Select Case strIDCode
        Case "S"    '## Token Transfer
            '## MT 전송
            Call SendMT
            
            '## Token 전송
            Call SendToken
            
        Case "E"    '## Workorder Validation
            strErrCd = Mid$(strRcvBuf, 11, 2)
            If strErrCd = "14" Then
                '## NOTE: 메시지의 Error 코드를 판단하여 사용자에게 표시해야함
            End If
            
            '## MT 전송
            Call SendMT
            
        Case "Q"    '## Query
            '## MT 전송
            Call SendMT
            
            strBarno = Trim(Format$(Mid$(strRcvBuf, 4, 14), String$(SPCLEN, "#")))
            With mORDER
                .ClsClear
                .BarNo = strBarno
            End With
            strSeqno = strBarno
            
            '## 오더전송
            Call GetOrder(strBarno)
            
        Case "R"    '## Result
            '## MT 전송
            Call SendMT
            
            '## 바코드번호, Tube Position 조회
            strBarno = Trim(Format$(Mid$(strRcvBuf, 4, 14), String$(SPCLEN, "#")))
            strTubePos = Trim$(Mid$(strRcvBuf, 20, 6))
            
            If strBarno <> "" Then
                pGrid_Point = SeqSearch(spdResult1, strBarno, colPAID)
                If pGrid_Point = 0 Then
                    pGrid_Point = SeqNullSearch(spdResult1, strBarno, colPAID)
                End If
            Else
                pGrid_Point = SeqNullSearch(spdResult1, strBarno, colPAID)
            End If
            
            If pGrid_Point = 0 Then
                spdResult1.maxrows = spdResult1.maxrows + 1
                pGrid_Point = spdResult1.maxrows
            End If

            'strBarno = "2012004611"
            intOrdCnt = 0
            
            '바코드번호로 오더찾아오기
            Set mAdoRs = f_subSet_WorkList_Barcode(strBarno)
            
            If RecordChk = True Then
                Do Until mAdoRs.EOF
                    intIdx = 0
                    With spdResult1
                        sCol = 5
                        pGrid_Point = SeqSearch(spdResult1, strBarno, sCol)
                        If pGrid_Point = 0 Then
                            pGrid_Point = SeqNullSearch(spdResult1, strBarno, sCol)
                            If pGrid_Point = 0 Then
                                .maxrows = .maxrows + 1
                                .RowHeight(.maxrows) = 13
                            End If
                            pGrid_Point = .maxrows
                        End If
                        
                        If intOrdCnt = 0 Then
                            .SetText 1, pGrid_Point, "1"
                            strSeqno = mAdoRs("SCP42IDNOA")
                            .SetText 2, pGrid_Point, strSeqno 'mAdoRs("SCP42IDNOA") & ""
                            .SetText 3, pGrid_Point, mAdoRs("SCP41NAME") & ""
                            .SetText 4, pGrid_Point, mAdoRs("SCP41JDATE") & ""
                            .SetText 5, pGrid_Point, strBarno
                            .SetText 6, pGrid_Point, mAdoRs("SCP42SUGACD") & ""
                        End If
                        
                        
                        intOrdCnt = intOrdCnt + 1
                        For intCol = 10 To .MaxCols
                            .GetText intCol, 0, varTmp
                            
                            
                            '-- 추가
                            strEqpCds = mAdoRs("SCP42SUGACD") & ""
                            Select Case strEqpCds
                                Case "L07007502", "L80004702": strEqpCds = "L07007502,L80004702"
                                Case "L07007503", "L80004703": strEqpCds = "L07007503,L80004703"
                                Case "L07007504", "L80004704": strEqpCds = "L07007504,L80004704"
                                Case "L07007505", "L80004705": strEqpCds = "L07007505,L80004705"
                                Case "L07007506", "L80004706": strEqpCds = "L07007506,L80004706"
                                'Case Else:                     strEqpCds = "L07007507,L80004707"
                            End Select
                            
                            Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                            If Not itemX Is Nothing Then
                                If strEqpCds = Trim(itemX.SubItems(1)) Then
                                'If InStr(Trim(itemX.SubItems(1)), mAdoRs("SCP42SUGACD") & "") > 0 Then
                                    intOrdCnt = intOrdCnt + 1
                                    spdResult1.Col = itemX.Index + 9
                                    spdResult1.BackColor = &HC6FEFF   '&HC6FEFF
                                    strOrd(intCnt) = mAdoRs("SCP42SUGACD") & ""
                                    intCnt = intCnt + 1
                                    Exit For
                                End If
                            End If
                        Next
                        
                    End With
                    mAdoRs.MoveNext
                Loop
            Else
                spdResult1.SetText 1, pGrid_Point, "0"
                spdResult1.SetText 2, pGrid_Point, ""
                spdResult1.SetText 3, pGrid_Point, ""
                spdResult1.SetText 4, pGrid_Point, ""
                spdResult1.SetText 5, pGrid_Point, strBarno
                spdResult1.SetText 6, pGrid_Point, ""
                
                lblStatus.Caption = "바코드 번호 " & strBarno & " 는 검사대상이 아닙니다"
            End If
                                                                                            
            Set mAdoRs = Nothing
            
            strTemp = mGetP(strRcvBuf, 2, vbLf)
            For i = 1 To Len(strTemp) Step 9
                strIntBase = Format$(Mid$(strTemp, i, 3), "000")
                strRstval = Trim$(Mid$(strTemp, i + 3, 5))
                
                intRow = 0
                With spdResult1
                    'If strBarno <> "" Then
                    '    pGrid_Point = SeqSearch(spdResult1, strBarno, colPAID)
                    '    If pGrid_Point = 0 Then
                    '        pGrid_Point = SeqNullSearch(spdResult1, strSeqno, colSQNO)
                    '    End If
                    'Else
                    '    pGrid_Point = SeqNullSearch(spdResult1, strSeqno, colSQNO)
                    'End If
                    'If pGrid_Point = 0 Then
                    '    spdResult1.maxrows = spdResult1.maxrows + 1
                    '    pGrid_Point = spdResult1.maxrows
                    'End If
                    
                    If pGrid_Point > 0 Then
                        spdResult1.GetText colORDT, pGrid_Point, varTmp:    strORDT = Trim$(varTmp)
                        spdResult1.GetText colORQN, pGrid_Point, varTmp:    strORQN = Trim$(varTmp)
                        spdResult1.GetText colPANM, pGrid_Point, varTmp:    strPANM = Trim$(varTmp)
                        spdResult1.GetText colPAID, pGrid_Point, varTmp:    strPAID = Trim$(varTmp)
                        spdResult1.GetText colOIFL, pGrid_Point, varTmp:    strOIFL = Trim$(varTmp)
                        spdResult1.GetText colSENO, pGrid_Point, varTmp:    strSENO = Trim$(varTmp)
                        spdResult1.GetText colSEXS, pGrid_Point, varTmp:    strSEXS = Trim$(varTmp)
                        spdResult1.GetText colAGES, pGrid_Point, varTmp:    strAGES = Trim$(varTmp)
                        spdResult1.GetText colNWNO, pGrid_Point, varTmp:    strNWNO = Trim$(varTmp)
                
                        For intCol = 10 To .MaxCols
                            .GetText intCol, 0, varTmp
                            Channel_No = strIntBase
                            Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                            If Not itemX Is Nothing Then
                                For intIdx = 1 To .MaxCols
                                    If Len(Channel_No) > 0 Then
                                        If Val(Channel_No) = Val(itemX.tag) Then
                                            strRefVal = ""

                                            strEqpCd = itemX.Text
                                            For s = 0 To UBound(strOrd)
                                                If InStr(Trim(itemX.Text), strOrd(s)) > 0 Then
                                                    strEqpCd = itemX.Text
                                                    Exit For
                                                End If
                                            Next
                                            
                                            
                                            'strRstval = strResult
                                            strState = "R"
                                            strRefVal = ""
                                            
                                            strDate = Format$(Now, "YYYYMMDD"):    strTime = Format$(Now, "HHMMSS")
                                            
                                            '-- 처방번호 찾기
                                            spdResult1.GetText intCol, pGrid_Point, varTmp: strORQN = varTmp
                                            spdResult1.Col = intCol
                                            spdResult1.ForeColor = vbBlack
                                            spdResult1.SetText intCol, pGrid_Point, strRstval
                    
                                            'If strORDT <> "" And Len(strEqpCd) <> 0 And strORQN <> "" Then
                                                '-- H/L 판정
                                               ' strRefVal = f_subSet_RefVal(strEqpCd, strRstval, strSEXS, strAGES)
        
                                               ' strSENO = 0
        
                                                '-- 로컬저장
                                            sqlDoc = "Update INTERFACE003" & _
                                                     "   set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRefVal & "'" & _
                                                     " where SPCNO   = '" & strBarno & "'" & _
                                                     "   and EQPNUM  = '" & itemX.tag & "'" & _
                                                     "   and TRANSDT = '" & strDate & "'" & _
                                                     "   and TRANSTM = '" & strTime & "'"
                                            AdoCn_Jet.Execute sqlDoc
                                            
                                            sqlDoc = "insert into INTERFACE003(" & _
                                                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                                     "    values( '" & strBarno & "', '" & strEqpCd & "', '" & itemX.tag & "'," & _
                                                     "            '" & strDate & "', '" & strTime & "'," & _
                                                     "            '" & strRstval & "', '" & strRefVal & "'," & _
                                                     "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
                                            AdoCn_Jet.Execute sqlDoc
                                            
                                            If chkAuto.Value = "1" And Len(strEqpCd) <> 0 Then
                                                '   3-1. 검사정보 MASTER
                                                         sqlDoc = "UPDATE JAIN_SCP.SCPRST41 SET "
                                                sqlDoc = sqlDoc & "       SCP41TSTDAT = '" & Format(Now, "YYYYMMDD") & "'," '결과일자 => YYYYMMDD"
                                                sqlDoc = sqlDoc & "       SCP41SNDYN  = 'N',"                               '고정값 : 'N'
                                                sqlDoc = sqlDoc & "       SCP41RSTYN  = 'Y',"                               '고정값 : 'Y'
                                                sqlDoc = sqlDoc & "       SCP41TSTUID = '" & CurrUser.CuUserID & "'"        '검사자사번
                                                sqlDoc = sqlDoc & " WHERE SCP41SPMNO2 = '" & strBarno & "'"                 '바코드번호
                                                
                                                AdoCn_ORACLE.Execute sqlDoc
                                                
                                                '   3-2. 검사정보 DETAIL
                                                         sqlDoc = "UPDATE JAIN_SCP.SCPRST42 SET "
                                                sqlDoc = sqlDoc & "       SCP42TSTDAT = '" & Format(Now, "YYYYMMDD") & "'," '결과일자 => YYYYMMDD"
                                                sqlDoc = sqlDoc & "       SCP42RSTCD  = 'N',"                               '결과형식 => 숫자 : 'N', 문자 : 'X', 장문 : 'R'
                                                sqlDoc = sqlDoc & "       SCP42RESULT = '" & strRstval & "'"                '결과값
                                                sqlDoc = sqlDoc & " WHERE SCP42SPMNO2 = '" & strBarno & "'"                 '바코드번호
                                                sqlDoc = sqlDoc & "   AND SCP42SUGACD = '" & strEqpCd & "'"              '수가코드
            
                                                AdoCn_ORACLE.Execute sqlDoc
                                                
                                                spdResult1.Row = pGrid_Point
                                                
                                                spdResult1.Col = colORDT:   spdResult1.BackColor = vbCyan
                                                spdResult1.Col = colORQN:   spdResult1.BackColor = vbCyan
                                                spdResult1.Col = colPANM:   spdResult1.BackColor = vbCyan
                                                spdResult1.Col = colPAID:   spdResult1.BackColor = vbCyan
                                                spdResult1.Col = colOIFL:   spdResult1.BackColor = vbCyan
                                                spdResult1.Col = colSENO:   spdResult1.BackColor = vbCyan
                                                spdResult1.Col = colSEXS:   spdResult1.BackColor = vbCyan
                                                'spdResult1.Col = colSENO:   spdResult1.BackColor = vbCyan
                                                'spdResult1.Col = colAGES:   spdResult1.BackColor = vbCyan
                                                'spdResult1.Col = colNWNO:   spdResult1.BackColor = vbCyan
                                                spdResult1.Col = colBANO:   spdResult1.Value = 0
                                                
            
                                            End If
                                            Set itemX = Nothing

                                            'End If
                                        End If
                                        Exit For
                                        Set itemX = Nothing
                                    End If
                                Next
                            End If
                        Next
                    End If
                End With
            Next i
            
            '## 결과저장
'            Call SaveServer(objIntInfo)
            
'            If strState = "R" Then
'                spdResult1.Col = colSQNO: spdResult1.ForeColor = vbRed: spdResult1.BackColor = vbCyan
'                spdResult1.Text = strBarno
'                spdResult1.SetText colBANO, pGrid_Point, "1"
'            End If
            
            
'            Set objIntNms = Nothing
'            Set objIntInfo = Nothing
            
            '## Result Validation 전송
            strOutput = mORDER.GetResultValid
            comEQP.Output = strOutput
            Print #1, "[ResultValidation]" & strOutput & vbNewLine;
    
    End Select
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 접수정보 조회, tblReady, tblResult에 표시
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)
'    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
    Dim strOutput  As String            '송신한 데이터
    Dim strBarno   As String
    Dim pGrid_Point As Integer
    Dim strEqpCd  As String
    Dim itemX   As ListItem
    Dim sOrderLst As String
    Dim blt As Boolean
    Dim intIdx As Integer
    Dim intOrdCnt As Integer
    Dim strSeqno As String
    
    intIdx = 0
    intOrdCnt = 0
    
    mORDER.BarNo = pBarNo
   ' mBarNo = pBarNo
    
    Set mAdoRs = f_subSet_WorkList_Barcode(pBarNo, INS_CODE, 0)
    If RecordChk = False Then
        'LstErr.AddItem ("샘플번호 " & fAdvia2120(4) & " 는 등록되지 않은 검체입니다. 확인하세요.")
        '## No Order 전송
        strOutput = mORDER.GetNoOrder
    Else
        mAdoRs.MoveFirst
        With spdResult1
            Do Until mAdoRs.EOF
                intIdx = 0
                
                pGrid_Point = SeqSearch(spdResult1, strBarno, 5)
                If pGrid_Point = 0 Then
                    pGrid_Point = SeqNullSearch(spdResult1, strBarno, 5)
                    If pGrid_Point = 0 Then
                        .maxrows = .maxrows + 1
                        .RowHeight(.maxrows) = 13
                        pGrid_Point = .maxrows
                    End If
                    'pGrid_Point = .maxrows
                End If
                
                strSeqno = mAdoRs("SCP42IDNOA")
                
                If intOrdCnt = 0 Then
                    .SetText 1, pGrid_Point, "1"
                    .SetText 2, pGrid_Point, strSeqno 'mAdoRs("SCP42IDNOA") & ""
                    .SetText 3, pGrid_Point, mAdoRs("SCP41NAME") & ""
                    .SetText 4, pGrid_Point, mAdoRs("SCP41JDATE") & ""
                    .SetText 5, pGrid_Point, pBarNo
                    .SetText 6, pGrid_Point, mAdoRs("SCP42SUGACD") & ""
                End If
                intOrdCnt = intOrdCnt + 1
                
                strEqpCd = mAdoRs("SCP42SUGACD") & "" 'f_funGet_CODE(Trim(mAdoRs.Fields("Coda")) & "/" & Trim(mAdoRs.Fields("SubCoda")) & "")
                                
                '-- 추가
                Select Case strEqpCd
                    Case "L07007502", "L80004702": strEqpCd = "L07007502,L80004702"
                    Case "L07007503", "L80004703": strEqpCd = "L07007503,L80004703"
                    Case "L07007504", "L80004704": strEqpCd = "L07007504,L80004704"
                    Case "L07007505", "L80004705": strEqpCd = "L07007505,L80004705"
                    Case "L07007506", "L80004706": strEqpCd = "L07007506,L80004706"
                    'Case Else:                     strEqpCd = "L07007507,L80004707"
                End Select
                
'                        Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                
'                Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                Set itemX = lvwCuData.FindItem(strEqpCd, lvwSubItem, , lvwWhole)
                
                
                If Not itemX Is Nothing Then
                    spdResult1.SetText 1, pGrid_Point, "1"
                    spdResult1.Col = itemX.Index + 9
                    spdResult1.Row = pGrid_Point
                    spdResult1.BackColor = &HC6FEFF   '&HC6FEFF
                    sOrderLst = sOrderLst & Format(itemX.tag, "000")
                    blt = True
                End If
                'strBarno = Trim(mAdoRs.Fields("LID"))
                mAdoRs.MoveNext
            Loop
                
            If blt = False Then
                .Row = pGrid_Point
                .Action = ActionDeleteRow
            End If
        
        End With
                           
        '## 오더정보 전송
        strOutput = mORDER.GetOrder(sOrderLst)
            
    End If
    
'    mAdoRs.Close
    Set mAdoRs = Nothing
    
    comEQP.Output = strOutput
'    Print #1, "[SendOrder]" & strOutput & vbNewLine;
    Print #1, vbNewLine & Format(Now, "yyyymmdd hh:mm:ss") & "[SendOrder]" & strOutput & vbNewLine;
    
    Debug.Print strOutput

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열을 구분자를 이용해 구분해 지정한 위치의 문자열을 구함
'   인수 :
'       1.pText      : 구분자로 구성된 문자열
'       2.pPosiion   : 위치
'       3.pDelimiter : 구분자
'-----------------------------------------------------------------------------'
Private Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
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

'-----------------------------------------------------------------------------'
'   기능 : Token Transfer Message 전송
'-----------------------------------------------------------------------------'
Private Sub SendToken()
    Dim strOutput As String     '송신할 데이터
    
    strOutput = mORDER.GetToken
    comEQP.Output = strOutput
    Print #1, vbNewLine & Format(Now, "yyyymmdd hh:mm:ss") & "[SendToken]" & strOutput & vbNewLine;

    mIntLibstate = ""
End Sub

Private Sub f_subSet_Result(ByVal strdata As String)

    On Error GoTo ErrRoutine
    
    CallForm = "frmInterface - Privete sub f_subSet_Result()"
    
    Dim sqlDoc  As String, sqlRet   As Integer
    Dim itemX   As ListItem
        
    Dim strSampleno As String
    Dim strBarno    As String, strEtc       As String, strDate      As String
    Dim strRstval() As String, strRefVal()  As String, strEqpCd()   As String
    Dim intIdx      As Integer, intCnt      As Integer, intRow  As Integer
    Dim strTmp      As String, strTime      As String, intCol   As Integer
    
    Dim strOrdLst() As String, intRet   As Integer
    Dim strPid()    As String, strPnm() As String
    
    Dim strOrdcd()  As String, strBarno1() As String
    Dim strLevel()  As String
    
    Dim strSerial As String
    Dim strRorder As String
    Dim varTmp
    
    strDate = Format$(Now, "YYYYMMDD"): strTime = Format$(Now, "MMSS")
    strBarno = Trim$(Mid$(strdata, 14, 20))
    strEtc = Trim$(Mid$(strdata, 25, 4))
    If strBarno = "" Then Exit Sub
    
    intIdx = InStr(strdata, "E")
    strTmp = Mid$(strdata, intIdx + 1)
    
    intCnt = 0
    Do While Len(strTmp) >= 10
        intCnt = intCnt + 1
        ReDim Preserve strEqpCd(1 To intCnt) As String
        ReDim Preserve strRstval(1 To intCnt) As String
        ReDim Preserve strRefVal(1 To intCnt) As String
        
        If Mid$(strTmp, 1, 2) = "13" Then
            Select Case Mid$(strEtc, 2, 1)
                Case "P":   strEqpCd(intCnt) = Mid$(strEtc, 1, 1) + "X1"
                Case "F":   strEqpCd(intCnt) = Mid$(strEtc, 1, 1) + "X2"
                Case Else:  strEqpCd(intCnt) = Mid$(strEtc, 1, 1) + Mid$(strTmp, 1, 2)
            End Select
        Else
            strEqpCd(intCnt) = Mid$(strEtc, 1, 1) + Mid$(strTmp, 1, 2)
        End If
        strRstval(intCnt) = Mid$(strTmp, 3, 6)
        strRefVal(intCnt) = Mid$(strTmp, 9, 1)
        
        strTmp = Mid$(strTmp, 11)
    Loop
    
    
    For intIdx = 1 To intCnt
        Set itemX = lvwCuData.FindItem(strEqpCd(intIdx), lvwTag, , lvwWhole)
        If Not itemX Is Nothing Then
            intCol = itemX.Index
            intRow = f_funGet_SpreadRow(spdResult1, strBarno, 2)
            If intRow < 1 Then
                intRow = f_funGet_SpreadRow(spdResult1, "", 2)
                If intRow < 1 Then
                    spdResult1.maxrows = spdResult2.maxrows + 1
                    spdResult1.RowHeight(spdResult2.maxrows) = 13
                    intRow = spdResult1.maxrows
                End If
                spdResult1.SetText 2, intRow, strBarno
            End If
            
            spdResult1.SetText intCol + 8, intRow, strRstval(intIdx)
            spdResult1.Col = intCol + 8
            spdResult1.Row = intRow:    spdResult1.ForeColor = IIf(Trim$(strRefVal(intIdx)) <> "", vbRed, vbBlack)
            
            spdResult1.GetText 2, intRow, varTmp:   strDate = Trim$(varTmp)
            spdResult1.GetText 3, intRow, varTmp:   strBarno = Trim$(varTmp)
            spdResult1.GetText 4, intRow, varTmp:   pName = Trim$(varTmp)
            spdResult1.GetText 5, intRow, varTmp:   pNo = Trim$(varTmp)
            spdResult1.GetText 6, intRow, varTmp:   strSerial = Trim$(varTmp)
            spdResult1.GetText 7, intRow, varTmp:   strRorder = Trim$(varTmp)

'            sqlDoc = "Update INTERFACE003" & _
'                     "   set RSTVAL  = '" & strRstval(intIdx) & "', REFVAL = '" & strRefVal(intIdx) & "'" & _
'                     " where SPCNO   = '" & strBarno & "'" & _
'                     "   and EQPNUM  = '" & itemX.tag & "'" & _
'                     "   and TRANSDT = '" & strDate & "'" & _
'                     "   and TRANSTM = '" & strTime & "'"
'            AdoCn_Jet.Execute sqlDoc, sqlRet
'            If sqlRet = 0 Then
'                sqlDoc = "insert into INTERFACE003(" & _
'                         "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN)" & _
'                         "    values( '" & strBarno & "', '" & itemX.ListSubItems(1) & "', '" & itemX.tag & "'," & _
'                         "            '" & strDate & "', '" & strTime & "'," & _
'                         "            '" & strRstval(intIdx) & "', '" & strRefVal(intIdx) & "'," & _
'                         "            '" & INS_CODE & "', '')"
'                AdoCn_Jet.Execute sqlDoc
'            End If
        
            sqlDoc = "Update INTERFACE003" & _
                     "   set RSTVAL  = '" & strRstval(intIdx) & "', REFVAL = '" & strRefVal(intIdx) & "'" & _
                     " where SPCNO   = '" & strBarno & "'" & _
                     "   and EQPNUM  = '" & itemX.tag & "'" & _
                     "   and TRANSDT = '" & strDate & "'" & _
                     "   and TRANSTM = '" & strTime & "'"
            AdoCn_Jet.Execute sqlDoc

            sqlDoc = "insert into INTERFACE003(" & _
                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                     "    values( '" & strBarno & "', '" & itemX.Text & "', '" & itemX.tag & "'," & _
                     "            '" & strDate & "', '" & strTime & "'," & _
                     "            '" & strRstval(intIdx) & "', '" & strRefVal(intIdx) & "'," & _
                     "            '" & INS_CODE & "', '', '" & pName & "', '" & strSerial & "')"

            AdoCn_Jet.Execute sqlDoc

            If chkAuto.Value = "1" Then
                sqlDoc = "AP_INF_Update '" & strBarno & "', '" & strBarno & "', '" & strSerial & "',"
                sqlDoc = sqlDoc & " '" & strRorder & "', '" & Mid(itemX.Text, 1, 5) & "', '" & Mid(itemX.Text, 6, Len(itemX.Text)) & "',"
                sqlDoc = sqlDoc & " '" & strRstval(intIdx) & "', '" & CurrUser.CuUserID & "'"

                AdoCn_SQL.Execute sqlDoc
                
                spdResult1.Row1 = intRow: spdResult1.Col1 = 2:
                spdResult1.Row2 = intRow: spdResult1.Col2 = 6
                spdResult1.BlockMode = True
                spdResult1.BackColor = vbCyan
                spdResult1.BlockMode = False
                
                spdResult1.Col = 1: spdResult1.Value = 0
            End If
        End If
        
        Set itemX = Nothing
    Next
    
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부터 전송된 Datalog를 편집
'-----------------------------------------------------------------------------'
'Private Sub EditRcvData()
'    Dim strRcvBuf   As String   '수신한 Data
'    Dim strSeq      As String   '수신한 Seq
'    Dim strType     As String   '수신한 Data Tytpe
'    Dim strBarNo    As String   '수신한 BarNo
'    Dim strPos      As String   '수신한 Tube Position
'    Dim strIntBase  As String   '수신한 장비기준 검사명
'    Dim strResult   As String   '수신한 결과
'    Dim strOutput   As String   '송신할 Data
'    Dim aryTemp()   As String
'
'    Dim strTemp     As String
'    Dim I           As Long
'
'    objInt.objDicBuf.MoveFirst
'    Do Until objInt.objDicBuf.EOF
'        strRcvBuf = objInt.objDicBuf.Fields("bufchar")
'        strType = Mid$(strRcvBuf, 3, 1)
'
'        If strType = "D" Then
'            '## 바코드번호, Cup Position 조회
'            strBarNo = Trim$(Mid$(strRcvBuf, 29, 15))
'            If strBarNo = "" Then Exit Sub
'
'            strPos = CStr(Asc(Mid$(strRcvBuf, 46, 1)) - 32)
'
'            With objRst.objPreRst
'                .DeleteAll
'                .AddNew "1", String(FieldCnt + 5, COL_DIV)
'
'                .KeyChange "1"
'                .Fields("devid") = strBarNo
'                .Fields("devseq") = strPos
'                .Fields("devinfo1") = ""
'                .Fields("devinfo2") = ""
'            End With
'
'            '## 장비기준 검사명, 결과조회
'            If strPos = "3" Then
'                strTemp = Mid$(strRcvBuf, 54)
'            Else
'                strTemp = Mid$(strRcvBuf, 53)
'            End If
'
'            strTemp = Mid$(strTemp, 1, InStr(strTemp, "|") - 1)
'            aryTemp = Split(strTemp, "}")
'
'            For I = LBound(aryTemp) To UBound(aryTemp)
'                strIntBase = Mid$(aryTemp(I), 1, 1)
'                strIntBase = IIf(strIntBase = Space(1), "N1", strIntBase)
'                strResult = Trim$(Mid$(aryTemp(I), 2, 9))
'
'                If IsNumeric(strResult) = False Then
'                    strResult = CS_EqpError
'                End If
'
'                With objRst.objDicIntBase
'                    If .Exists(strIntBase) Then
'                        .KeyChange (strIntBase)
'
'                        objRst.objPreRst.KeyChange "1"
'                        objRst.objPreRst.Fields("rst" & .Fields("eqpseq")) = strResult & vbTab & _
'                                                                .Fields("testcd") & vbTab & _
'                                                                .Fields("intnm")
'                    End If
'                End With
'            Next I
'
'            '## 결과저장
'            Call SaveServer(strBarNo, strPos)
'        End If
'
'        objInt.objDicBuf.MoveNext
'    Loop
'End Sub

Private Sub ReceiveTheData(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object)
    Dim sTemp      As String
    Dim Channel_No As String       ' 검사항목 번호 : Channel No
    Dim pGrid_Point As Integer
    Dim pDoCount   As Integer
    Dim Loop_count As Integer
    Dim FunStr As String
    Dim Max_Arary_Cnt As Integer    ' 검사 항목수
    Dim sAdd As Integer, sPosition As Integer
    Dim itemX As ListItem
    Dim strRstval As String, strRefVal  As String
    Dim sqlDoc  As String
    Dim intCol As Integer
    Dim Gnum   As String
    Dim ii As Integer, jj As Integer, kk As Integer
    Dim Test_Cd() As String
    Dim Rev As Long
    Dim tmpTstCd As String
    Dim tmpMXD As Variant
    Dim sSeq, strTmp, varTmp, strBarno, strDate, strTime, strDate1 As String
    Dim SaveGbn As Integer
    Dim sCol As Integer
    Dim sDecnt As Integer
    Dim Float_rate1 As String
    Dim Float_rate2 As String
    Dim Float_rate  As String
    Dim intRow, intIdx As Integer
    Dim chrChk As Boolean
    Dim sSeqTmp As Variant
    Dim sOrderLst   As String
    Dim ssTemp1, ssTemp2 As String
    Dim fTest(88) As String
    Dim TestStatus(88)  As String * 1
    Dim xx              As Integer
    Dim intCnt As Integer
    Dim strEqpCd As String
    Dim blt As Boolean
    Dim sItemCode As Variant
    Dim intItemCode As Integer
    Dim strRorder  As String
    
    On Error Resume Next

    CallForm = "frmInterface - Privete sub psDataDefine()"
    tmpMXD = "0"
    Erase fAdvia2120
    
    sTemp = Mid(strdata, 2)

    fAdvia2120(0) = Str$(sDecnt)                   ' 총 검사항목 갯수를 넣는다. 나중에 사용한다.
    fAdvia2120(1) = Mid$(sTemp, 1, 1)              ' 항목 1 ":"   '
                                               
    pGrid_Point = 0
    
    Select Case fAdvia2120(1)
        Case ""
            comEQP.Output = STX & ">" & ETX & vbCr & vbLf
        Case "?"
            comEQP.Output = STX & ">" & ETX & vbCr & vbLf
        Case ">"
            comEQP.Output = STX & ">" & ETX & vbCr & vbLf
        Case ";"
            fAdvia2120(3) = Mid(sTemp, 2, 2)           ' Spare
            fAdvia2120(4) = Mid(sTemp, 5, 5)           ' RackNumber
            fAdvia2120(5) = Mid(sTemp, 11, 3)           ' SEQNO
'            fAdvia2120(6) = Mid(sTemp, 13, 1)          ' PositionNumber
            fAdvia2120(7) = Mid(sTemp, 32, 10)         ' IDNumber
'            strBarno = Trim(fAdvia2120(7))
            Set mAdoRs = f_subSet_WorkList_Barcode(Trim(fAdvia2120(7)), INS_CODE, 0)
            If RecordChk = False Then
                LstErr.AddItem ("샘플번호 " & fAdvia2120(4) & " 는 등록되지 않은 검체입니다. 확인하세요.")
                comEQP.Output = STX & ";A1" & Space(12) & Trim(fAdvia2120(7)) & "   3 0330051941 " & "88000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" & ETX & vbCr & vbLf
            Else
                mAdoRs.MoveFirst
                With spdResult1
                    Do Until mAdoRs.EOF
                        If strBarno <> Trim(mAdoRs.Fields("LID")) & "" Then
                            pGrid_Point = SeqSearch(spdResult1, mAdoRs.Fields("LID"), 3)
                
                            If pGrid_Point = 0 Then
                                pGrid_Point = SeqNullSearch(spdResult1, mAdoRs.Fields("LID") & "", 3)
                                If pGrid_Point = 0 Then .maxrows = .maxrows + 1: pGrid_Point = .maxrows
                            End If
                            
'                            r.LID: 검체번호
'                            r.Hcode     : 병록(차트)번호
'                            r.Serial    : 일련번호(처방접수시 생성되는 번호)
'                            c.PtName: 환자이름
'                            r.Orderdate: 처방일
'                            r.Coda: 검사항목 코드
'                            r.SubCoda: 검사항목 서브코드
'                            r.ROrder    : 오더번호(결과저장시 중복 피하기 위해 필수)
'                            r.ErYn      : 응급여부(0:일반, 1:응급)

                            .SetText 1, pGrid_Point, "1"
                            .SetText 2, pGrid_Point, mAdoRs("Orderdate")
                            .SetText 3, pGrid_Point, mAdoRs("LID")
                            .SetText 4, pGrid_Point, mAdoRs("PtName")
                            .SetText 5, pGrid_Point, mAdoRs("Serial")
                            .SetText 6, pGrid_Point, mAdoRs("ROrder")
                        End If
                        
                        strEqpCd = f_funGet_CODE(Trim(mAdoRs.Fields("Coda")) & "/" & Trim(mAdoRs.Fields("SubCoda")) & "")
                        Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                        If Not itemX Is Nothing Then
                            spdResult1.SetText 1, pGrid_Point, "1"
                            spdResult1.Col = itemX.Index + 7
                            spdResult1.Row = pGrid_Point
                            spdResult1.BackColor = &HC6FEFF   '&HC6FEFF
                            sOrderLst = sOrderLst & itemX.tag & ","
                            blt = True
                        End If
                        strBarno = Trim(mAdoRs.Fields("LID"))
                        mAdoRs.MoveNext
                    Loop
'                    If blt = False Then
'                        .Row = pGrid_Point
'                        .Action = ActionDeleteRow
'                    End If
                End With
                           
                For xx = 1 To 88
                    TestStatus(xx) = "0"
                Next xx
            ';N1  2221 221104149035831      0530042157 880000000000000000000000011110001000000000000000000000000000000000000000000000000000000000100000104149035831                  75
            '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
                    ssTemp2 = ";" & fAdvia2120(3)
                    ssTemp2 = ssTemp2 & Space(1) & fAdvia2120(4) & Space(1) & fAdvia2120(5) & Space(18) & fAdvia2120(7)
                    ssTemp2 = ssTemp2 & Space(1) & "88"
    '
                    For Loop_count = 1 To 88: fTest(Loop_count) = "": Next Loop_count
                    
                    pDoCount = 0
                    Do While InStr(sOrderLst, ",") > 0
                        pDoCount = pDoCount + 1
                        fTest(pDoCount) = Text_Redefine(sOrderLst, ",")
                        sOrderLst = Mid$(sOrderLst, InStr(sOrderLst, ",") + 1)
                        TestStatus(Val(fTest(pDoCount))) = "1"
                        If pDoCount > 99 Then
                            sOrderLst = ""
                            Exit Do
                        End If
                    Loop
                    
                    For xx = 1 To 87
                        ssTemp2 = ssTemp2 & TestStatus(xx)
                    Next xx
                    
                    ssTemp2 = ssTemp2 & "000000" 'Comment
                    Debug.Print "[HOST]" & STX & ssTemp2 & ETX & vbCr & vbLf
                    comEQP.Output = STX & ssTemp2 & ETX & vbCr & vbLf
            '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
            
            End If
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
            comEQP.Output = STX & ">" & ETX & vbCr & vbLf
        Case ":"
            If fAdvia2120(1) <> "@" And fAdvia2120(1) <> "?" And fAdvia2120(1) <> "L" And Len(sTemp) > 1 Then
                sDecnt = Mid$(sTemp, 49, 2)                                        ' channel count
                fAdvia2120(2) = Trim(Mid(sTemp, 31, 10))                         ' IDNumber
                                             
                For pDoCount = 1 To sDecnt
                    ssTemp1 = (pDoCount - 1) * 10 + 51           ' 첫번째 Channel 및 검사결과 위치 확인
                    ssTemp2 = Mid$(sTemp, ssTemp1, 9)
                    fAdvia2120(((pDoCount - 1) * 2) + 8) = Trim(Mid$(ssTemp2, 1, 3))   ' channel
                    fAdvia2120(((pDoCount - 1) * 2) + 9) = Trim(Mid$(ssTemp2, 4, 6))   ' result
                Next pDoCount
                                
                
                intRow = 0
                pGrid_Point = 0
                With spdResult1
                    sSeq = Trim(fAdvia2120(2))
                    sCol = 3
                    pGrid_Point = SeqSearch(spdResult1, sSeq, sCol)
        
                    .GetText 2, pGrid_Point, varTmp:   strDate = Trim$(varTmp)
                    .GetText 4, pGrid_Point, varTmp:   pName = Trim$(varTmp)
                    .GetText 3, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)
                    .GetText 5, pGrid_Point, varTmp:   pNo = Trim$(varTmp)
                    .GetText 6, pGrid_Point, varTmp:   strRorder = Trim$(varTmp)
        
                    If pGrid_Point > 0 Then
                        For intCol = 9 To .MaxCols
                            strRstval = ""
                            .GetText intCol, 0, varTmp
                            Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                            If Not itemX Is Nothing Then
                                For intIdx = 1 To .MaxCols
                                    If Len(fAdvia2120(2)) > 0 Then
                                        Channel_No = Val(fAdvia2120(((intIdx - 1) * 2) + 8))
                                        If Channel_No = itemX.tag Then
                                            strRstval = Trim(fAdvia2120(((intIdx - 1) * 2) + 9))
                                            strRefVal = ""
                                            strDate1 = Format$(Now, "YYYYMMDD"):     strTime = Format$(Now, "MMSS")
                                            .SetText intCol, pGrid_Point, strRstval
                                            .Col = intCol:  .Row = pGrid_Point
                                                            .ForeColor = IIf(Trim$(strRefVal) <> "", vbRed, vbBlack)
        
'                                            sqlDoc = "Update INTERFACE003" & _
'                                                     "   set RSTVAL  = '" & strResult & "', REFVAL = '" & strRefVal & "'" & _
'                                                     " where SPCNO   = '" & pNo & "'" & _
'                                                     "   and EQPNUM  = '" & itemX.tag & "'" & _
'                                                     "   and TRANSDT = '" & strDate1 & "'" & _
'                                                     "   and TRANSTM = '" & strTime & "'"
'                                            AdoCn_Jet.Execute sqlDoc
                
'                                            sqlDoc = "insert into INTERFACE003(" & _
'                                                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
'                                                     "    values( '" & pNo & "', '" & itemX.Text & "', '" & itemX.tag & "'," & _
'                                                     "            '" & strDate1 & "', '" & strTime & "'," & _
'                                                     "            '" & strResult & "', '" & strRefVal & "'," & _
'                                                     "            '" & INS_CODE & "', '', '" & pName & "', '" & strSerial & "')"
'
'                                            AdoCn_Jet.Execute sqlDoc
'
'                                            If chkAuto.Value = "1" Then
'                                                sqlDoc = "AP_INF_Update '" & strBarno & "', '" & strBarno & "', '" & strSerial & "',"
'                                                sqlDoc = sqlDoc & " '" & strRorder & "', '" & Mid(itemX.Text, 1, 5) & "', '" & Mid(itemX.Text, 6, Len(itemX.Text)) & "',"
'                                                sqlDoc = sqlDoc & " '" & strRstval & "', '" & CurrUser.CuUserID & "'"
'
'                                                AdoCn_SQL.Execute sqlDoc
'
'                                                spdResult1.Row1 = pGrid_Point: spdResult1.Col1 = 2:
'                                                spdResult1.Row2 = pGrid_Point: spdResult1.Col2 = 6
'                                                spdResult1.BlockMode = True
'                                                spdResult1.BackColor = vbCyan
'                                                spdResult1.BlockMode = False
'
'                                                spdResult1.Col = 1: spdResult1.Value = 0
'                                            End If
'
'                                            spdResult1.Row = pGrid_Point
'                                            spdResult1.Col = 2
'                                            spdResult1.BackColor = vbCyan
'                                            spdResult1.Col = 3
'                                            spdResult1.BackColor = vbCyan
'                                            spdResult1.Col = 4
'                                            spdResult1.BackColor = vbCyan
'                                            spdResult1.Col = 6
'                                            spdResult1.BackColor = vbCyan
'                                            spdResult1.Col = 7
'                                            spdResult1.BackColor = vbCyan
'                                            spdResult1.Col = 1: spdResult1.Value = 0
                                            Exit For
                                        End If
                                    End If
                                Next intIdx
                            End If
                            Set itemX = Nothing
                        Next
                    End If
                End With
                'comEQP.Output = STX & "A" & ETX & vbCr & vbLf
                comEQP.Output = STX & ">" & ETX & vbCr & vbLf
            End If
        Case Else
        'comEQP.Output = STX & "A" & ETX & vbCr & vbLf
        comEQP.Output = STX & ">" & ETX & vbCr & vbLf
    End Select
    
    Exit Sub

ErrRoutine:

    Call ErrMsgProc(CallForm)
      
End Sub

Private Sub psDataDefine(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object) ', ByVal brOst As String) ' ByRef brItemdeci() As String)

    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Privete sub psDataDefine()"

    Dim sqlDoc  As String, sqlRet   As Integer

    Dim varTmp      As Variant
    Dim strTmp      As String
    Dim intRow      As Long, intCol As Integer, intIdx  As Integer
    Dim strRstval(4) As String
    Dim strRstval1  As String
    Dim strRstval2  As String
    Dim strRstval3  As String
    Dim strRstval4  As String
    Dim strRefVal    As String
    Dim sSeq, sCol, strBarno As String
    Dim strTime  As String, strDate  As String
    Dim strSeqno    As String
    Dim ii          As Integer

    Dim strOrdLst() As String, strPid() As String, strPnm() As String
    Dim intRet      As Integer
    Dim Channel_No  As String       ' 문자형 변수
    Dim Result As String
    Dim itemX   As ListItem

    ReceiveData = strdata
    If Len(ReceiveData) < 2 Then
        Exit Sub
    End If

    If InStr(ReceiveData, ":") > 0 Then
        crCnt = 1
    End If

    Select Case crCnt
        Case 1
            crCnt = crCnt + 1
        Case 2
            crCnt = crCnt + 1
        Case 3
            crCnt = crCnt + 1
        Case 4
            Print #1, "ReceiveData >>" & ReceiveData;
            sSeq = Right(ReceiveData, 2)
            Print #1, "sSeq >>" & sSeq;
            If sSeq = Patiant_Recevid Then
                strRstval(1) = ""
                strRstval(2) = ""
                strRstval(3) = ""
                strRstval(4) = Mid(ReceiveData, 1, 4)       'APTT
            Else
                strRstval(1) = Mid(ReceiveData, 1, 4)       'PT(Sec)
                strRstval(3) = Mid(ReceiveData, 5, 4)       'PT(%)
                strRstval(2) = Trim(Mid(ReceiveData, 9, 5)) 'PT(INR)
                strRstval(4) = ""
            End If
            
            sCol = 1
            Patiant_Recevid = SeqSearch(spdResult1, sSeq, sCol)

            If Len(sSeq) > 0 Then
                intRow = 0
                With spdResult1

                    .GetText 3, Patiant_Recevid, varTmp:   strBarno = Trim$(varTmp)

                    If Patiant_Recevid > 0 Then
                        For intCol = 7 To .MaxCols
                            .GetText intCol, 0, varTmp
                            Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                            If Not itemX Is Nothing Then
                                For intIdx = 1 To 4
                                    Channel_No = intIdx
                                    If Channel_No = itemX.tag And strRstval(intIdx) <> "" Then
                                        If strRstval(intIdx) = "S" Or strRstval(intIdx) = "N" Then
                                            strRstval(intIdx) = ""
                                        End If

                                        strDate = Format$(Now, "YYYYMMDD"):     strTime = Format$(Now, "MMSS")
                                        
                                        .SetText intCol, Patiant_Recevid, strRstval(intIdx)
                                        
                                        .Col = intCol:  .Row = Patiant_Recevid
                                                        .ForeColor = IIf(Trim$(strRefVal) <> "", vbRed, vbBlack)

                                            sqlDoc = "Update INTERFACE003" & _
                                                     "   set RSTVAL  = '" & strRstval(intIdx) & "', REFVAL = '" & strRefVal & "'" & _
                                                     " where SPCNO   = '" & strBarno & "'" & _
                                                     "   and EQPNUM  = '" & itemX.tag & "'" & _
                                                     "   and TRANSDT = '" & strDate & "'" & _
                                                     "   and TRANSTM = '" & strTime & "'"
                                            AdoCn_Jet.Execute sqlDoc, sqlRet

                                            If sqlRet = 0 And chkAuto.Value = vbChecked Then
                                                AdoCn_SQL.Execute "Exec InterfaceResult_INSERT_sp '" & strBarno & "','" & itemX & "','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & strRstval(intIdx) & "','" & INS_CODE & "','' ", sqlRet

                                                If sqlRet = 1 Then
                                                    lblStatus.Caption = "저장 성공!!"
                                                    spdResult1.SetText 1, Patiant_Recevid, "0"
                                                    spdResult1.Row = Patiant_Recevid
                                                    spdResult1.Col = -1:    spdResult1.BackColor = &HFFF8F0
                                                    sqlDoc = "insert into INTERFACE003(" & _
                                                             "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN)" & _
                                                             "    values( '" & strBarno & "', '" & itemX.Text & "', '" & itemX.tag & "'," & _
                                                             "            '" & strDate & "', '" & strTime & "'," & _
                                                             "            '" & strRstval(intIdx) & "', '" & strRefVal & "'," & _
                                                             "            '" & INS_CODE & "', '')"
                                                    AdoCn_Jet.Execute sqlDoc
                                                Else
                                                    lblStatus.Caption = "저장 실패!!"
                                                End If

                                            End If
                                            Exit For
                                    End If

                                Next intIdx
                            End If
                            Set itemX = Nothing
                        Next
                        '-----------------------------------------------------------------------
                    End If
                End With
            End If

        Case Else

    End Select

    Exit Sub

ErrRoutine:

    Call ErrMsgProc(CallForm)

End Sub

Private Function GetContent(ByVal pFlags As String) As String
    Dim strFlag     As String   'Abnormal Flag
    Dim strContent  As String   'Abnormal Content
    Dim strTemp     As String
    Dim i           As Long
    
    If pFlags = "" Then Exit Function
    
    For i = 1 To Len(Trim(pFlags))
        strFlag = Mid$(pFlags, i, 1)
        If i = 1 Then
            strTemp = Space(2) & "[Abnormality flag]: " & strFlag & vbCrLf
        Else
            strTemp = strTemp & vbCrLf & Space(2) & "[Abnormality flag]: " & strFlag & vbCrLf
        End If
        
        strContent = ""
        Select Case strFlag
            Case "/": strContent = "Test no performed: test has been requisitioned but not performed due to any reason."
            Case "S": strContent = "Result extracted for repeat run"
            Case "?": strContent = "Calculation unable due to abnormal photometric data. UNIT in STOP mode (Incl. Lamp OFF), etc."
            Case "n": strContent = "8087 error"
            Case "R": strContent = "Reagent level detection error"
            Case "#": strContent = "Sample level detection error"
            Case "!": strContent = "A/D error of photometry"
            Case ">": strContent = "The absolute OD value is over 2.665."
            Case "<": strContent = "The absolute OD value is under 0.99."
            Case "-": strContent = "The final result is negative."
            Case "U": strContent = "Reagent absorbance value at P0 of Reagent Blank run, is smaller than the lower limit of the Parameter."
            Case "u": strContent = "Reagent absorbance value at P0 or p8 is lower than the lower limit specified in the Parameters in routine run."
            Case "Y": strContent = "Reagent absorbance value at P16 of Reagent Blank run, is greater than the upper limit of the Parameter."
            Case "y": strContent = "Reagent absorbance value at P0 or p8 is higher than the upper limit specified in the Parameters in routine run."
            Case "@": strContent = "Abnormally high result: absorbance of every wavelength is more than 2.5."
            Case "$": strContent = "No linearity validation conducted because less than 3 data obtained in the kinetics."
            Case "D": strContent = "Too quick reaction slope in increasing kinetics, absorbance at P-START is higher than MAX. OD in increasing FIXED assay, or too slow reaction slope in decreasing kinetics (=no reaction observed)"
            Case "B": strContent = "Too quick reaction slope in increasing kinetics, or absorbance at P-END is lower than MIN. OD in increasing FIXED assay."
            Case "*": strContent = "Linearity error in kinetics"
            Case "P": strContent = "Result higher than DECIDE RANGE designated in parameters."
            Case "N": strContent = "Result lower than DECIDE RANGE designated in parameters."
            Case "&": strContent = "Data check 2 error"
            Case "Z": strContent = "Data check 1 error"
            Case "F": strContent = "Result higher than the dynamic range specified in the Parameters"
            Case "G": strContent = "Result lower than the dynamic range specified in the Parameters"
            Case "p": strContent = "Result beyond the panic value specified in the Parameters"
            Case "T": strContent = "Abnormality found in the Inter-Item Check"
            Case "H": strContent = "Result higher than the normal value range specified in the Parameters"
            Case "L": strContent = "Result lower than the normal value range specified in the Parameters"
            Case "W": strContent = "Abnormality in WB data. Photocal has not been performed."
            Case "J": strContent = "Result higher than the repeat run range specified in the Parameters"
            Case "K": strContent = "Result higher than the repeat run range specified in the Parameters"
        End Select
        
        If strContent <> "" Then
            strTemp = strTemp & Space(2) & "[Content]: " & strContent & vbCrLf
        End If
    Next i
    
    GetContent = vbCrLf & strTemp

End Function


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
'        If optSeq.Value = False Then
            For sCnt = 1 To .maxrows
                .Row = sCnt
                .Col = brCol
                If Trim(.Text) = Trim(brSeq) Then
                    SeqSearch = sCnt
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            Next sCnt
'        Else
'            For sCnt = 1 To .maxrows
'                .Row = sCnt
'                .Col = brCol
'                If Val(spdResult1.StartingRowNumber + (Val(sCnt) - 1)) = Val(brSeq) Then
'                    SeqSearch = sCnt 'brSeq
'                    .Action = ActionActiveCell
'                    .Refresh
'                    Exit For
'                End If
'            Next sCnt
'        End If
    End With

End Function

Private Sub Command1_Click()

   
    Dim Arr()   As Byte
    Dim strDta  As String

    strDta = ";A1               3300001   3  1080200065" & vbCrLf
'    strDta = ";N1  2011  11                  108020006598" & vbCrLf
'    strDta = ":n1  6011  11  04149035831    10802000658        4 24  5092  25   625  26   195  27    61 6C" & vbCrLf

strDta = "ZQ 1080200008" & vbCrLf & "/"
    
strDta = "ZR 00001080200008 002-01           02/16/08 15:00:56   " & vbCrLf
strDta = strDta & "1 6.76  22  4.4  23  0.9  24  1.0  25  2.2  28*****  26 1.57  27  1.0 " & vbCrLf
strDta = strDta & ","

    Call ComReceive(strDta)
    
End Sub

Private Sub Form_Activate()
    
    If IS_SET = False Then Unload Me

End Sub

Private Sub Form_Load()
    
    Call GetResultValid
    
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
    
    Call cmdRun                 ' 실행
    
    mskRstDate.Text = Format$(Now, "YYYYMMDD")
    mskOrdDate.Text = Format$(Now, "YYYYMMDD")
    mskOrdDate1.Text = Format$(Now, "YYYYMMDD")
    Open App.Path + "\Log\" & Format(Now, "yyyymmdd") & "_Advia2120.log" For Append As #1

    Print #1, Chr(13) + Chr(10);
   
    f_strJOB_FLAG = "1"
    f_intSampleNo = 0
    tabWork.Tab = 0
    Or_Seq = 1
    intRow = 0
    chkEnq = 0
'    cboChk.ListIndex = 1
    objIntPhase = 1
    
    mIntLibPhase = 1
    
    Set mORDER = New clsIISIntOrder
    
'    Timer2.Interval = 60000
'    Timer2.Interval = 10000
    
'    Timer2.Interval = 10000
'    Timer2.Enabled = True
    
    Call cmdInit_Click

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
'            Call Del_OldData
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
    LstErr.Visible = False
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

''    Dim aROW    As Integer, aCOL   As Integer
''    Dim varChk  As Variant, varBar As Variant, varNum As Variant
''    Dim iRow    As Integer, iCnt   As Integer
''
''    'Debug.Print Col & NewCol & Row & NewRow
''
''    If KeyAscii = vbKeyReturn Then
''        With spdResult1
''            aCOL = .ActiveCol
''            aROW = .ActiveRow
''            If aCOL = 4 Then
''                iCnt = 0
''                For iRow = aROW To .maxrows
''                    .GetText 1, iRow, varChk
''                    .GetText 3, iRow, varBar
''                    .GetText aCOL, aROW, varNum
''                    If Trim(varChk) = "1" And Trim(varBar) <> "" Then
''                        .SetText aCOL, iRow, varNum
''                        .SetText aCOL + 1, iRow, ((iCnt Mod 40) + 1) + (40 * (varNum - 1))
''                        iCnt = iCnt + 1
''                        If (iCnt Mod 40) = 0 Then varNum = varNum + 1
''                    End If
''                Next
'''            ElseIf aCOL = 5 Then
'''                iCnt = 0
'''                For iRow = aROW To .maxrows
'''                    .GetText 1, iRow, varChk
'''                    .GetText 3, iRow, varBar
'''                    .GetText aCOL, aROW, varNum
'''                    If Trim(varChk) = "1" And Trim(varBar) <> "" Then
'''                        .SetText aCOL, iRow, ((iCnt Mod 40) + varNum) '+ (40 * (varNum - 1))
'''                        '.SetText aCOL - 1, iRow, varNum
'''                        iCnt = iCnt + 1
'''                        If (iCnt Mod 40) = 0 Then varNum = varNum + 1
'''                    End If
'''                Next
''
''            End If
''        End With
''    End If
    

    Dim aRow    As Integer, aCOL   As Integer
    Dim varChk  As Variant, varBar As Variant, varNum As Variant
    Dim iRow    As Integer, iCnt   As Integer
    
    'Debug.Print Col & NewCol & Row & NewRow
       
    If KeyAscii = vbKeyReturn Then
        With spdResult1
            .Row = .ActiveRow: aRow = .ActiveRow
            .Col = .ActiveCol: aCOL = .ActiveCol
                               varBar = Trim(.Text)
                               
            If aCOL = 5 And Len(varBar) = 10 Then
                '바코드번호로 오더찾아오기
                Set mAdoRs = f_subSet_WorkList_Barcode(varBar)
                
                If RecordChk = True Then
                    Do Until mAdoRs.EOF
                        .SetText 1, aRow, "1"
                        .SetText 2, aRow, mAdoRs("SCP42IDNOA") & ""
                        .SetText 3, aRow, mAdoRs("SCP41NAME") & ""
                        .SetText 4, aRow, mAdoRs("SCP41JDATE") & ""
                        '.SetText 5, aRow, varBar
                        .SetText 6, aRow, mAdoRs("SCP42SUGACD") & ""
                        mAdoRs.MoveNext
                    Loop
                Else
                    .SetText 1, aRow, "0"
                    .SetText 2, aRow, ""
                    .SetText 3, aRow, ""
                    .SetText 4, aRow, ""
                    .SetText 5, aRow, varBar
                    .SetText 6, aRow, ""
                    
                    lblStatus.Caption = "바코드 번호 " & varBar & " 는 검사대상이 아닙니다"
                End If
                                                                                                
                Set mAdoRs = Nothing
            End If
        End With
    End If
    
End Sub

'Private Sub spdWorkList_Click(ByVal Col As Long, ByVal Row As Long)
'
'    If Col < 3 Then Exit Sub
'
'    Dim varTmp  As Variant
'
'    With spdWorklist
'        If Col = 1 Then
'            .GetText 2, Row, varTmp
'            If Trim$(varTmp) = "" Then Exit Sub
'
'            .SetText 1, Row, IIf(Trim$(varTmp) = "1", "", "1")
'        ElseIf Col > 4 Then
'            .GetText Col, 0, varTmp
'            If Trim$(varTmp) = "" Then Exit Sub
'
'            .Row = Row: .Col = Col
'            If .BackColor = vbWhite Then
'                .BackColor = &HC6FEFF
'            Else
'                .BackColor = vbWhite
'            End If
'        End If
'    End With
'
'End Sub

Private Sub SpreadSheetSort(ByRef Spread As vaSpread, ByVal Col As Integer, Optional ByVal SortType As Integer = 1)
    Dim intCount As Integer
    Dim strDataField As String
    'SortType
    ' 0 : none
    ' 1 : ascending
    ' 2 : descending

    With Spread
        .Col = 1: .Col2 = .MaxCols
        .Row = 1: .Row2 = .DataRowCnt
        .SortBy = 0
        .SortKey(1) = Col       '정렬키 열번호

        If SortType = 0 Then
            .SortKeyOrder(1) = SortKeyOrderNone
        ElseIf SortType = 1 Then
            .SortKeyOrder(1) = SortKeyOrderAscending
        ElseIf SortType = 2 Then
            .SortKeyOrder(1) = SortKeyOrderDescending
        Else
            .SortKeyOrder(1) = SortKeyOrderAscending
        End If

        .Action = ActionSort
    End With

End Sub


Private Sub spdWorklist_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim varTmp  As Variant
    
    If Row = 0 Then
        If Col = 1 Then
            Col = 2
        End If
        
        If OrderSort_Flag = 1 Then
            Call SpreadSheetSort(spdWorklist, Col, 2)
            OrderSort_Flag = 2
        Else
            Call SpreadSheetSort(spdWorklist, Col, 1)
            OrderSort_Flag = 1
        End If
    Else
        With spdWorklist
            .GetText 2, Row, varTmp
            If Trim$(varTmp) = "" Then Exit Sub
    
            .SetText 1, Row, IIf(Trim$(varTmp) = "1", "", "1")
            cmdWorkList_Click
        End With
    End If
End Sub

Private Sub Timer1_Timer()

    Call COM_OUTPUT(ENQ)
'    Debug.Print ENQ

End Sub

Private Sub Timer2_Timer()
    
'    f_strJOB_FLAG = "1"
'    f_intSampleNo = 0
'    tabWork.Tab = 0
'    Or_Seq = 1
'    intRow = 0
'    chkEnq = 0
'    objIntPhase = 1
'    mIntLibPhase = 1
    
    Call cmdInit_Click

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
                                intRow = spdWorklist.maxrows
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
    With spdWorklist
        For sCnt = 1 To .maxrows
            .Row = sCnt:    .Col = 2
            If Trim(.Text) = Mid(txtBarCode.Text, 1, 11) Then
                psDataExists = True
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Sub txtBarCode_LostFocus()

'    Dim intRow      As Integer
'    Dim strOrdcd(1 To 100) As String
'
'    Call sl_spcid_tstcd_select&(txtBarCode.Text, strOrdcd)
'    If strOrdcd(1) = "" Then
'        MsgBox "해당 검사항목이 존재하지 않은 검체입니다.", vbInformation, Me.Caption
'        Exit Sub
'    End If
'
'    intRow = f_funGet_SpreadRow(spdWorkList, 2, txtBarCode.Text)
'    If intRow < 1 Then
'        intRow = f_funGet_SpreadRow(spdWorkList, 2, "")
'        If intRow < 1 Then
'            spdWorkList.maxrows = spdWorkList.maxrows + 1
'            spdWorkList.RowHeight(spdWorkList.maxrows) = 13
'            intRow = spdWorkList.maxrows
'        End If
'        spdWorkList.SetText 2, intRow, txtBarCode.Text
'    End If
'    spdWorkList.SetText 1, intRow, "1"
    
End Sub


' 통신상태 확인 관련이벤트
' ------------------------------------------------------------------------
Private Sub txtCom_Change()
    txtCom.SelStart = Len(txtCom.Text)
End Sub

Private Sub cmdCOMLoad_Click()
    Dim i               As Long
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
        
        txtCOM2.Text = ""
        ReDim bteBuffer(LOF(lngFIleNum))
        Get #lngFIleNum, , bteBuffer

        strTemp = StrConv(bteBuffer, vbUnicode)
        txtCOM2.Text = strTemp
                
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
              txtCom.Text & _
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
    txtCom.Text = ""
End Sub

Private Sub cmdCOMClear2_Click()
    txtCOM2.Text = ""
End Sub

Private Sub cmdCOMInput_Click()

    Dim bytTemp() As Byte
    
    bytTemp = StrConv(charCOM_Convert(txtCom.SelText), vbFromUnicode)

    Call ComReceive(txtCom.SelText)
    
End Sub

Private Sub cmdCOMInput2_Click()
    
    Dim bytTemp() As Byte
    
    If txtCOM2.SelLength = 0 Then
        bytTemp = StrConv(charCOM_Convert(txtCOM2.Text), vbFromUnicode)
    Else
        bytTemp = StrConv(charCOM_Convert(txtCOM2.SelText), vbFromUnicode)
    End If

    Call ComReceive(txtCOM2.SelText)

End Sub

Private Sub cmdCOMOutput2_Click()
    
    If txtCOM2.SelLength = 0 Then
        Call COM_OUTPUT(charCOM_Convert(txtCOM2.Text))
    Else
        Call COM_OUTPUT(charCOM_Convert(txtCOM2.SelText))
    End If
    
End Sub
' ------------------------------------------------------------------------
' 통신상태 확인 관련이벤트


Private Sub txtResult_DblClick()
    txtResult.Text = ""
    LstErr.Text = ""
    
    If txtResult.Visible Then txtResult.Visible = False
    LstErr.Visible = True
End Sub

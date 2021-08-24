VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmComm 
   Caption         =   "Interface"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   15360
   WindowState     =   2  '최대화
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
      Begin VB.Timer tmrSend 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5475
         Top             =   60
      End
      Begin VB.Timer tmrReceive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5070
         Top             =   60
      End
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
         TabIndex        =   23
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
         TabIndex        =   24
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
         TabIndex        =   25
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
         TabIndex        =   26
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
         TransparentPicture=   "frmComm.frx":0000
         ImgOutLineSize  =   3
      End
      Begin MSComctlLib.ImageList imlList 
         Left            =   3900
         Top             =   60
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
               Picture         =   "frmComm.frx":188A
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmComm.frx":1E24
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmComm.frx":23BE
               Key             =   "NOF"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmComm.frx":2958
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmComm.frx":2EF2
               Key             =   "LSE"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmComm.frx":348C
               Key             =   "LSN"
            EndProperty
         EndProperty
      End
      Begin MSCommLib.MSComm comEQP 
         Left            =   4485
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         Handshaking     =   1
         RThreshold      =   1
         SThreshold      =   1
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   5880
         Top             =   60
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
               Picture         =   "frmComm.frx":3A26
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmComm.frx":3FC0
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmComm.frx":455A
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmComm.frx":4AF4
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmComm.frx":5386
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmComm.frx":54E0
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmComm.frx":563A
               Key             =   "NOF"
            EndProperty
         EndProperty
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
      Left            =   60
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
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSPanel1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lvwCuData"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "pnlCom2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "pnlCom"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdPosNo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdEot"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdAppend(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "FrameResult"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkAuto"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "SSFrame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fraError"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "spdResult1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   " 받은 결과"
      TabPicture(1)   =   "frmComm.frx":684E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdSel(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdSel(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdRstQuery"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdAppend(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "spdResult2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "SSPanel3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin FPSpread.vaSpread spdResult1 
         Height          =   4800
         Left            =   105
         TabIndex        =   44
         Top             =   420
         Width           =   14955
         _Version        =   393216
         _ExtentX        =   26379
         _ExtentY        =   8467
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
         MaxCols         =   5
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         SpreadDesigner  =   "frmComm.frx":686A
         UserResize      =   0
      End
      Begin Threed.SSFrame fraError 
         Height          =   2355
         Left            =   8160
         TabIndex        =   37
         Top             =   5850
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
            TabIndex        =   40
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
            TabIndex        =   39
            Top             =   225
            Visible         =   0   'False
            Width           =   6630
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   3075
         Left            =   105
         TabIndex        =   38
         Top             =   5175
         Width           =   8040
         _Version        =   65536
         _ExtentX        =   14182
         _ExtentY        =   5424
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
            Left            =   90
            TabIndex        =   47
            Top             =   135
            Width           =   7875
            _Version        =   393216
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
            ScrollBarMaxAlign=   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmComm.frx":6C99
            UserResize      =   0
         End
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
         Left            =   13650
         TabIndex        =   22
         Top             =   60
         Value           =   1  '확인
         Width           =   1410
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TEST"
         Height          =   285
         Left            =   11130
         TabIndex        =   21
         Top             =   60
         Visible         =   0   'False
         Width           =   990
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   3
         Left            =   -74640
         TabIndex        =   19
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":72E6
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   -74910
         TabIndex        =   20
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":7768
      End
      Begin BHButton.BHImageButton cmdRstQuery 
         Height          =   375
         Left            =   -69060
         TabIndex        =   28
         Top             =   450
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
      Begin BHButton.BHImageButton cmdAppend 
         Height          =   375
         Index           =   1
         Left            =   -61140
         TabIndex        =   27
         Top             =   480
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
      Begin Threed.SSFrame FrameResult 
         Height          =   3525
         Left            =   285
         TabIndex        =   36
         Top             =   945
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
      Begin BHButton.BHImageButton cmdAppend 
         Height          =   510
         Index           =   0
         Left            =   13500
         TabIndex        =   41
         Top             =   5310
         Visible         =   0   'False
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   900
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
      Begin BHButton.BHImageButton cmdEot 
         Height          =   375
         Left            =   9900
         TabIndex        =   42
         Top             =   30
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
      Begin BHButton.BHImageButton cmdPosNo 
         Height          =   375
         Left            =   5460
         TabIndex        =   45
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
      Begin HSCotrol.UserPanel pnlCom 
         Height          =   4875
         Left            =   90
         TabIndex        =   29
         Top             =   900
         Visible         =   0   'False
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   8599
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
            TabIndex        =   30
            Top             =   315
            Visible         =   0   'False
            Width           =   11595
         End
         Begin VB.Frame Frame1 
            Height          =   645
            Left            =   45
            TabIndex        =   31
            Top             =   3660
            Visible         =   0   'False
            Width           =   11610
            Begin HSCotrol.CButton cmdCOMSave 
               Height          =   360
               Left            =   10515
               TabIndex        =   32
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
               TabIndex        =   33
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
               TabIndex        =   34
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
               TabIndex        =   35
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
         Left            =   5895
         TabIndex        =   9
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
            TabIndex        =   10
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
               TabIndex        =   11
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
               TabIndex        =   12
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
               TabIndex        =   13
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
               TabIndex        =   14
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
               TabIndex        =   15
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
               TabIndex        =   16
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
               TabIndex        =   17
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
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   18
            Top             =   780
            Width           =   5730
         End
      End
      Begin FPSpread.vaSpread spdResult2 
         Height          =   7350
         Left            =   -74910
         TabIndex        =   43
         Top             =   900
         Width           =   15015
         _Version        =   393216
         _ExtentX        =   26485
         _ExtentY        =   12965
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   6
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
         MaxCols         =   6
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         SpreadDesigner  =   "frmComm.frx":7BD6
         UserResize      =   0
      End
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   4515
         Left            =   10470
         TabIndex        =   46
         Top             =   735
         Visible         =   0   'False
         Width           =   4635
         _ExtentX        =   8176
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
      Begin Threed.SSPanel SSPanel1 
         Height          =   465
         Left            =   90
         TabIndex        =   48
         Top             =   390
         Visible         =   0   'False
         Width           =   4080
         _Version        =   65536
         _ExtentX        =   7197
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
         Begin VB.Label Label6 
            BackColor       =   &H00FFC0C0&
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
            Height          =   225
            Left            =   120
            TabIndex        =   50
            Top             =   150
            Width           =   1095
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFC0C0&
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
            Left            =   2520
            TabIndex        =   49
            Top             =   135
            Width           =   165
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   465
         Left            =   -74910
         TabIndex        =   51
         Top             =   390
         Width           =   5805
         _Version        =   65536
         _ExtentX        =   10239
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
         Begin VB.ComboBox cboRstgbn 
            Height          =   300
            Index           =   1
            ItemData        =   "frmComm.frx":8080
            Left            =   3990
            List            =   "frmComm.frx":808D
            Style           =   2  '드롭다운 목록
            TabIndex        =   53
            Top             =   75
            Width           =   1770
         End
         Begin MSComCtl2.DTPicker mskRstDate 
            Height          =   285
            Left            =   1260
            TabIndex        =   52
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            _Version        =   393216
            Format          =   87359489
            CurrentDate     =   40646
         End
         Begin MSComCtl2.DTPicker mskRstDate1 
            Height          =   285
            Left            =   2700
            TabIndex        =   54
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            _Version        =   393216
            Format          =   87359489
            CurrentDate     =   40646
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
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
            TabIndex        =   56
            Top             =   165
            Width           =   1125
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFC0C0&
            Caption         =   "-"
            Height          =   165
            Left            =   2580
            TabIndex        =   55
            Top             =   150
            Width           =   195
         End
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
         TabIndex        =   8
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

Private Const TEST_NM_EQP   As String = "EQP_NM"    '장비 코드
Private Const TEST_CD_LIS   As String = "LIS_CD"    '검사실 코드
Private Const TEST_NM_LIS   As String = "LIS_NM"    '검사실 이름
Private Const TEST_VALUES   As String = "VALUES"    '결과

Const OrderColor As String = &HC6FEFF          '    '오더 배경색
Const pIntCol   As Integer = 6

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""

Dim fHitachi7180(100)       As String

Private mAdoRs      As ADODB.Recordset
Private mAdoRs1     As ADODB.Recordset
Private CallForm    As String
Private IS_SET      As Boolean

Private f_strBuffer     As String

Dim fChannel() As String
Dim pName   As String
Dim pNo     As String

Private Type TYPE_CD
    strEqpCd    As String
    intCnt      As Integer
    strTestcd(100) As String
End Type
Private f_typCode() As TYPE_CD

Dim RecordChk As Boolean

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
        End With
        .HideColumnHeaders = False
    End With
    
End Sub

Private Function f_subSet_WorkList_Barcode(ByVal psBarno As String, ByVal psMCode As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList_Barcode() As ADODB.Recordset"
    
   
        Set AdoRs_SQL = New ADODB.Recordset
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open AdoCn_SQL.Execute("Exec neolis..AP_INF_Bar_Order_Coda '" & psMCode & "','" & psBarno & "'", sqlRet)
        
        If sqlRet = 0 Then
            Set f_subSet_WorkList_Barcode = Nothing
            RecordChk = False
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

Private Sub f_subSet_ItemList()

    Dim itemX   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strTest As String, intPos   As Integer
    Dim strTmp  As String, intCol   As Integer, intCol2   As Integer, intCnt  As Integer, intRow  As Integer
    
'On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear
    
    intCol = pIntCol
    intCol2 = 1
    intRow = 1
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .MaxRows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdResult2
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .MaxRows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
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
    End If
    
    Do While Not adoRS.EOF
        Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TEST_EQP") & ""), , "LST")
            itemX.SubItems(1) = Trim(adoRS.Fields("TESTCD_EQP") & "")
            itemX.SubItems(2) = Trim(adoRS.Fields("TESTNM_EQP") & "")
            itemX.Tag = Trim(adoRS.Fields("TEST_EQP") & "")
            itemX.Text = Trim(adoRS.Fields("TESTCD") & "")
        Set itemX = Nothing
        
        With spdResult1
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
                .ColWidth(intCol) = 6.5
            End If
            .SetText intCol, 0, Trim$(adoRS("TESTNM_EQP") & "")
        End With
        
        With spdRstview
            If intRow > .MaxRows Then
                intRow = 1
                intCol2 = intCol2 + 2
            End If
            
            .SetText intCol2, intRow, Trim$(adoRS("TESTNM_EQP") & "")
            intRow = intRow + 1
            
        End With
        
        With spdResult2
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
                .ColWidth(intCol) = 6.5
            End If
            .SetText intCol, 0, Trim$(adoRS("TESTNM_EQP") & "")
        End With
        
        fChannel(intCol - (pIntCol - 1)) = adoRS.Fields("TEST_EQP")
        
        intCnt = intCnt + 1
        ReDim Preserve f_typCode(1 To intCnt) As TYPE_CD
        
        f_typCode(intCnt).strEqpCd = Trim$(adoRS.Fields("TEST_EQP"))
        f_typCode(intCnt).intCnt = 0
        
        strTmp = Trim$(adoRS.Fields("TESTCD"))
        intPos = InStr(strTmp, ",")
        Do While intPos > 0
            f_typCode(intCnt).intCnt = f_typCode(intCnt).intCnt + 1
            f_typCode(intCnt).strTestcd(f_typCode(intCnt).intCnt) = Mid$(strTmp, 1, intPos - 1)
            
            strTmp = Mid$(strTmp, intPos + 1)
            
            intPos = InStr(strTmp, ",")
        Loop
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

Private Sub chkAuto_Click()
    If chkAuto.Value = 1 Then
        cmdAppend(0).Visible = False
    Else
        cmdAppend(0).Visible = True
    End If
End Sub

Private Sub cmdEot_Click()
    Call COM_OUTPUT(EOT)
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
    
End Sub

Private Sub cmdClear()
    
    With spdResult1
        .MaxRows = 0
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
        .RowHeight(-1) = 12
    End With

    With spdResult2
        .MaxRows = 0
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With

    LstErr.Clear

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

Dim varTmp      As Variant, strErrMsg   As String
Dim strBarno    As String
Dim strRstval   As String

Dim intRow  As Integer, intCol  As Integer
Dim itemX   As ListItem
Dim objSpd  As vaSpread
Dim intColtmp As Integer
    
    CallForm = "frmComm - Private Sub cmdAppend_Click()"

On Error GoTo ErrorRoutine

    Me.MousePointer = 11

    If Index = 0 Then
        Set objSpd = spdResult1
        intColtmp = pIntCol       'WorkList와 받은결과의 컬럼수가 틀릴때..
    Else
        Set objSpd = spdResult2
        intColtmp = 6
    End If

    With objSpd
    
        For intRow = 1 To .MaxRows
        
            .GetText 3, intRow, varTmp:   strBarno = Trim$(varTmp)
            .GetText 4, intRow, varTmp:   pName = Trim$(varTmp)
            .GetText 5, intRow, varTmp:   pNo = Trim$(varTmp)

            If strBarno = "" Then Exit For

            .GetText 1, intRow, varTmp
            If Trim$(varTmp) = "1" Then
                For intCol = intColtmp To .MaxCols
                    .GetText intCol, intRow, varTmp
                    If Trim$(varTmp) <> "" Then
                        .GetText intCol, 0, varTmp
                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                        If Not itemX Is Nothing Then
                            .GetText intCol, intRow, varTmp
                            strRstval = Trim(varTmp)
                                
                            Dim tmpTestcd   As Variant
                            Dim tmpTestSeq1 As Integer
                            Dim tmpTestCd1  As String

                            tmpTestcd = Split(itemX.Text, ",")
                            Set mAdoRs = f_subSet_TestList(strBarno)
                            For tmpTestSeq1 = 0 To UBound(tmpTestcd)
                                tmpTestCd1 = tmpTestcd(tmpTestSeq1)
                                mAdoRs.MoveFirst
                                Do Until mAdoRs.EOF
                                    If Trim(mAdoRs("CODA") & "/" & mAdoRs("SUBCODA")) = tmpTestCd1 Then
                                        Exit For
                                    End If
                                    mAdoRs.MoveNext
                                Loop
                            Next
                            
                            sqlDoc = "AP_INF_Bar_Result '" & strBarno & "', "
                            sqlDoc = sqlDoc & " '" & INS_CODE & "', '" & Mid(tmpTestCd1, 1, InStr(tmpTestCd1, "/") - 1) & "', '" & Mid(tmpTestCd1, InStr(tmpTestCd1, "/") + 1) & "',"
                            sqlDoc = sqlDoc & " '" & strRstval & "'"
                            AdoCn_SQL.Execute sqlDoc
                            lblStatus.Caption = "저장 성공!!"
                                
                            Set adoRS = Nothing:    mAdoRs.Close
                            
                            .Row = intRow
                            .Col = 2: .BackColor = vbCyan
                            .Col = 3: .BackColor = vbCyan
                            .Col = 4: .BackColor = vbCyan
                            .Col = 5: .BackColor = vbCyan
                            .Col = 1: .Value = 0
            
                            If strErrMsg = "" Then
                                sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y' Where SPCNO   = '" & strBarno & "'"
                                AdoCn_Jet.Execute sqlDoc
                            Else
                                MsgBox strErrMsg, vbInformation, INS_NAME
                            End If
                        End If

                        Set itemX = Nothing
                    End If
                Next
            End If
        Next
    End With
    Me.MousePointer = 0
    MsgBox "작업이 완료되었습니다.", vbInformation, Me.Caption

    Exit Sub
ErrorRoutine:
    Set itemX = Nothing

    Me.MousePointer = 0
    Call ErrMsgProc(CallForm)
End Sub

Private Sub cmdENQ_Click()
    
    Call COM_OUTPUT(charCOM_Convert(COM_ENQ))

End Sub

Private Sub cmdRstQuery_Click()
Dim adoRS   As New ADODB.Recordset
Dim sqlDoc  As String

Dim strSpcno    As String
Dim intRow      As Integer, intCol  As Integer

Dim itemX       As ListItem

    intRow = 0
    With spdResult2
        .MaxRows = 0
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    sqlDoc = "select SPCNO, TESTCD, EQUIPCD, TRANSDT, RSTVAL, REFVAL, TRANSDT, EQPNUM, NAME, PNO" & _
             "  from INTERFACE003" & _
             " where TRANSDT >= '" & Format(mskRstDate.Value, "YYYYMMDD") & "'" & _
             "   and TRANSDT <= '" & Format(mskRstDate1.Value, "YYYYMMDD") & "'" & _
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
                If intRow > .MaxRows Then .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
                .SetText 1, intRow, "1"
                .SetText 2, intRow, Trim$(adoRS(3) & "")
                .SetText 3, intRow, Trim$(adoRS(0) & "")
                .SetText 4, intRow, Trim$(adoRS(8) & "")
                .SetText 5, intRow, Trim$(adoRS(9) & "")
            End If
            strSpcno = Trim$(adoRS(0) & "") + Trim$(adoRS(6) & "")
            Set itemX = lvwCuData.FindItem(Trim$(adoRS(7) & ""), lvwTag, , lvwWhole)
            If Not itemX Is Nothing Then
                intCol = itemX.Index + 5
                .SetText intCol, intRow, Trim$(adoRS(4)) & ""
            End If
        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub cmdSel_Click(Index As Integer)

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    If Index = 2 Or Index = 3 Then
        With spdResult2
            For intRow = 1 To .MaxRows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 0, "1", "")
            Next
        End With
    End If
    
End Sub

Private Function f_subSet_TestList(ByVal strBarcode As String)
   Dim sqlRet      As Integer
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
    
    Set AdoRs_SQL = New ADODB.Recordset
    AdoRs_SQL.CursorLocation = adUseClient
    AdoRs_SQL.Open AdoCn_SQL.Execute("Exec AP_INF_Bar_Order_Coda '" & INS_CODE & "', '" & strBarcode & "'", sqlRet)
    
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

Private Sub ComReceive(ByRef RecData As String)
Dim strRec      As String, strBuff  As String
Dim intIdx      As Integer
Dim sStartCheck As Integer
Dim sEndCheck   As Integer

    strRec = RecData

    Print #1, strRec;
    Call COM_INPUT(strRec)
    
    For intIdx = 1 To Len(strRec)
        strBuff = Mid$(strRec, intIdx, 1)
        f_strBuffer = f_strBuffer + strBuff
        sStartCheck = InStr(f_strBuffer, STX)
        sEndCheck = InStr(f_strBuffer, vbCrLf)
        If sStartCheck <> 0 And sEndCheck <> 0 Then
            f_strBuffer = f_strBuffer + strBuff
            Call ReceiveTheData(f_strBuffer, fChannel(), spdResult1)
            f_strBuffer = ""
        End If
     Next
    
End Sub

Private Sub ReceiveTheData(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object)
Dim pGrid_Point As Integer
Dim pDoCount    As Integer
Dim Loop_count  As Integer
Dim sDecnt      As Integer
Dim xx          As Integer
Dim intIdx      As Integer
Dim intCol      As Integer
Dim sTemp       As String
Dim Channel_No  As String       ' 검사항목 번호 : Channel No
Dim strRstval   As String
Dim sqlDoc      As String
Dim strEqpCd    As String
Dim sSeq, strBarno, strDate, strTime As String
Dim sOrderLst   As String
Dim ssTemp1, ssTemp2 As String
Dim fTest(88)   As String
Dim TestStatus(88)  As String * 1
Dim strROrder   As String
Dim itemX       As ListItem
Dim varTmp      As Variant

    On Error Resume Next

    CallForm = "frmInterface - Privete sub ReceiveTheData()"
    Erase fHitachi7180
    
    sTemp = Mid(strdata, 2)

    fHitachi7180(0) = Str$(sDecnt)                   ' 총 검사항목 갯수를 넣는다. 나중에 사용한다.
    fHitachi7180(1) = Mid$(sTemp, 1, 1)              ' 항목 1 ":"   '
                                               
    pGrid_Point = 0
    
    Select Case fHitachi7180(1)
        Case ""
            comEQP.Output = STX & ">" & ETX & vbCr & vbLf
        Case "?"
            comEQP.Output = STX & ">" & ETX & vbCr & vbLf
        Case ">"
            comEQP.Output = STX & ">" & ETX & vbCr & vbLf
        Case ";"
            fHitachi7180(3) = Mid(sTemp, 2, 2)           ' Spare
            fHitachi7180(4) = Mid(sTemp, 5, 5)           ' RackNumber
            fHitachi7180(5) = Mid(sTemp, 11, 3)           ' SEQNO
            fHitachi7180(6) = Mid(sTemp, 4, 28)          ' PositionNumber
            fHitachi7180(7) = Trim(Mid(sTemp, 16, 13))         ' IDNumber
            strBarno = ""

            Set mAdoRs = f_subSet_WorkList_Barcode(Trim(fHitachi7180(7)), INS_CODE)
            If RecordChk = False Then
                LstErr.AddItem ("샘플번호 " & fHitachi7180(7) & " 는 등록되지 않은 검체입니다. 확인하세요.")
                comEQP.Output = STX & ";A1" & Space(12) & Trim(fHitachi7180(7)) & "   3 00216080951 " & "88000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" & ETX & vbCr & vbLf
            Else
                mAdoRs.MoveFirst
                With spdResult1
                    Do Until mAdoRs.EOF
                        If strBarno <> Trim(mAdoRs.Fields("BCID")) & "" Then
                            pGrid_Point = SeqSearch(spdResult1, Trim(mAdoRs.Fields("BCID")), 3)

                            If pGrid_Point = 0 Then .MaxRows = .MaxRows + 1: pGrid_Point = .MaxRows

                            .SetText 1, pGrid_Point, "1"
                            .SetText 2, pGrid_Point, mAdoRs("Orderdate")
                            .SetText 3, pGrid_Point, Trim(mAdoRs("BCID"))
                            .SetText 4, pGrid_Point, mAdoRs("PtName")
                            .SetText 5, pGrid_Point, mAdoRs("HCode")
                        End If

                        strEqpCd = f_funGet_CODE(Trim(mAdoRs.Fields("Coda")) & "/" & Trim(mAdoRs.Fields("SubCoda")) & "")
                        Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                        If Not itemX Is Nothing Then
                            .SetText 1, pGrid_Point, "1"
                            .Col = itemX.Index + (pIntCol - 1)
                            .Row = pGrid_Point
                            .BackColor = OrderColor
                            sOrderLst = sOrderLst & itemX.Tag & ","
                        End If
                        strBarno = Trim(mAdoRs.Fields("BCID"))
                        mAdoRs.MoveNext
                    Loop
                End With

                For xx = 1 To 88
                    TestStatus(xx) = "0"
                Next xx
                ';A1     0  1100216081106       1080200063 88111101010100110011000000000000000000000000000000000000000000000000000000000000000000000000000
                ';N1  2221 221104149035831      0530042157 880000000000000000000000011110001000000000000000000000000000000000000000000000000000000000100000104149035831                  75
                ';A1     0  11   1080200063    00216080951
                ';A1     0  11   1080200063    00216081045 88011101010100110011000000000000000000000000000000000000000000000000000000000000000000000000000
                ';A1     0  2100216081110       1080200062 88111101010100110011000000000000000000000000000000000000000000000000000000000000000000000000000
                ';A1     0  11                  1080200063 88000001110011111001000000101000000000000000000000000000000000000000000000000000000000000000000
                ';A1     0  11   1080200063    00216080951 88111101010100110011000000000000000000000000000000000000000000000000000000000000000000000000000
                '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
                    ssTemp2 = ";" & fHitachi7180(3)
                    ssTemp2 = ssTemp2 & fHitachi7180(6) & Format(Now, "mmddyyhhmm") & Space(1)
                    ssTemp2 = ssTemp2 & "88"
    '
                    For Loop_count = 1 To 88: fTest(Loop_count) = "": Next Loop_count

                    Dim strISE As Variant

                    strISE = Split(sOrderLst, ",")

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

                    For xx = 1 To 88
                        ssTemp2 = ssTemp2 & TestStatus(xx)
                    Next xx

                    ssTemp2 = ssTemp2 & "000000" 'Comment
                    'Debug.Print "[HOST]" & STX & ssTemp2 & ETX & vbCr & vbLf
                    comEQP.Output = STX & ssTemp2 & ETX & vbCr & vbLf
            '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-

            End If
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
            comEQP.Output = STX & ">" & ETX & vbCr & vbLf
        Case ":"
            If fHitachi7180(1) <> "@" And fHitachi7180(1) <> "?" And fHitachi7180(1) <> "L" And Len(sTemp) > 1 Then
                sDecnt = Mid$(sTemp, 49, 2)                                        ' channel count
                fHitachi7180(2) = Trim(Mid(sTemp, 16, 13))                         ' IDNumber
                                             
                For pDoCount = 1 To sDecnt
                    ssTemp1 = (pDoCount - 1) * 10 + 51           ' 첫번째 Channel 및 검사결과 위치 확인
                    ssTemp2 = Mid$(sTemp, ssTemp1, 10)
                    fHitachi7180(((pDoCount - 1) * 2) + 8) = Trim(Mid$(ssTemp2, 1, 3))   ' channel
                    fHitachi7180(((pDoCount - 1) * 2) + 9) = Trim(Mid$(ssTemp2, 4, 7))   ' result
                Next pDoCount
                                
                pGrid_Point = 0
                With spdResult1
                    .ReDraw = False
                    sSeq = Trim(fHitachi7180(2))
                    pGrid_Point = SeqSearch(spdResult1, sSeq, 3)
        
                    .GetText 3, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)
                    .GetText 4, pGrid_Point, varTmp:   pName = Trim$(varTmp)
                    .GetText 5, pGrid_Point, varTmp:   pNo = Trim$(varTmp)
        
                    If pGrid_Point > 0 Then
                        For intCol = pIntCol To .MaxCols
                            strRstval = ""
                            .GetText intCol, 0, varTmp
                            Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                            If Not itemX Is Nothing Then
                                For intIdx = 1 To .MaxCols
                                    If Len(fHitachi7180(2)) > 0 Then
                                        Channel_No = Val(fHitachi7180(((intIdx - 1) * 2) + 8))
                                        If Channel_No = itemX.Tag Then
                                            strRstval = Trim(fHitachi7180(((intIdx - 1) * 2) + 9))
                                            
                                            strDate = Format$(Now, "YYYYMMDD"):     strTime = Format$(Now, "hhmmss")
                                            
                                            '20101019 sog 정량 결과에 문자가 포함된 결과는 저장하지 않고 표시만
                                            '20110308 sog 모든 알파벳 대분자 중에 하나가 포함되어 있는지 체크
                                            Dim liCheckCnt      As Integer
                                            Dim lbAlphaExist    As Boolean
                                            lbAlphaExist = False
                                            For liCheckCnt = Asc("A") To Asc("Z")
                                                If InStr(UCase(strRstval), Chr(liCheckCnt)) > 0 Then
                                                    lbAlphaExist = True
                                                    liCheckCnt = Asc("Z") + 1
                                                End If
                                            Next liCheckCnt
                                            
                                            If lbAlphaExist = True Then
                                                .SetText intCol, pGrid_Point, strRstval
                                                .Row = pGrid_Point: .Col = intCol: .BackColor = &H80FF&

                                                .Col = 1: .Value = 0
                                            Else
                                                .SetText intCol, pGrid_Point, strRstval
                                                .Col = intCol:  .Row = pGrid_Point: .ForeColor = vbBlack
            
                                                sqlDoc = "Update INTERFACE003" & _
                                                         "   set RSTVAL  = '" & strRstval & "', REFVAL = ''" & _
                                                         " where SPCNO   = '" & pNo & "'" & _
                                                         "   and EQPNUM  = '" & itemX.Tag & "'" & _
                                                         "   and TRANSDT = '" & strDate & "'" & _
                                                         "   and TRANSTM = '" & strTime & "'"
                                                AdoCn_Jet.Execute sqlDoc
    
                                                sqlDoc = "insert into INTERFACE003(" & _
                                                         "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                                         "    values( '" & strBarno & "', '" & itemX.Text & "', '" & itemX.Tag & "'," & _
                                                         "            '" & strDate & "', '" & strTime & "'," & _
                                                         "            '" & strRstval & "', ''," & _
                                                         "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
                                                AdoCn_Jet.Execute sqlDoc
    
                                                If chkAuto.Value = "1" Then
                                                    Dim tmpTestcd   As Variant
                                                    Dim tmpTestSeq1 As Integer
                                                    Dim tmpTestCd1  As String
                                                    
                                                    tmpTestcd = Split(itemX.Text, ",")
                                                    Set mAdoRs = f_subSet_TestList(strBarno)
                                                    For tmpTestSeq1 = 0 To UBound(tmpTestcd)
                                                        tmpTestCd1 = tmpTestcd(tmpTestSeq1)
                                                        mAdoRs.MoveFirst
                                                        Do Until mAdoRs.EOF
                                                            If Trim(mAdoRs("CODA") & "/" & mAdoRs("SUBCODA")) = tmpTestCd1 Then
                                                                Exit For
                                                            End If
                                                            mAdoRs.MoveNext
                                                        Loop
                                                    Next
                                                    
                                                    sqlDoc = "exec neolis..AP_INF_Bar_Result '" & strBarno & "', '" & INS_CODE & "', '" & Mid(tmpTestCd1, 1, InStr(tmpTestCd1, "/") - 1) & "', '" & Mid(tmpTestCd1, InStr(tmpTestCd1, "/") + 1) & "', '" & strRstval & "'"
                                                    AdoCn_SQL.Execute sqlDoc
                                                    mAdoRs.Close:   Set mAdoRs = Nothing
    
                                                    .Row1 = pGrid_Point: .Col1 = 2:
                                                    .Row2 = pGrid_Point: .Col2 = 5
                                                    .BlockMode = True
                                                    .BackColor = vbCyan
                                                    .BlockMode = False
    
                                                    .Col = 1: .Value = 0
                                                    .Row = pGrid_Point
                                                    .Col = 2: .BackColor = vbCyan
                                                    .Col = 3: .BackColor = vbCyan
                                                    .Col = 4: .BackColor = vbCyan
                                                    .Col = 5: .BackColor = vbCyan
                                                End If
                                            End If
                                            Exit For
                                        End If
                                    End If
                                Next intIdx
                            End If
                            Set itemX = Nothing
                        Next intCol
                    End If
                    .ReDraw = True
                End With
                comEQP.Output = STX & ">" & ETX & vbCr & vbLf
            End If
        Case Else
            comEQP.Output = STX & ">" & ETX & vbCr & vbLf
    End Select
    
    Exit Sub

ErrRoutine:

    Call ErrMsgProc(CallForm)
      
End Sub

Private Function Text_Redefine(FSend_Str As String, FCheck_Char As String) As String
    
    If InStr(FSend_Str, FCheck_Char) > 0 Then
        Text_Redefine = left$(FSend_Str, InStr(FSend_Str, FCheck_Char) - 1)
    Else
        Text_Redefine = FSend_Str
    End If
    
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
            If Trim(.Text) = Trim(brSeq) Then
                SeqSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Sub Command1_Click()

   
    Dim Arr()   As Byte
    Dim strDta  As String

'    strDta = ";A1               3300001   3  1080200065" & vbCrLf
'    strDta = ";A1     0  11   09091400001    09091400001" & vbCrLf
     strDta = ":A1     0  11   09091400001    09091400001      11  1   7.2   2   4.9   3   0.6   4   0.2   6    59  10   129  13   170  14   230  17   100  19  11.7  93   2.1 " & vbCrLf
'    strDta = ";N1  2011  11                  108020006598" & vbCrLf
'    strDta = ":n1  6011  11  04149035831    10802000658        4 24  5092  25   625  26   195  27    61 6C" & vbCrLf

    Call ComReceive(strDta)
    
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
    
    Call f_subGet_Setting       ' 통신설정
    
    Call cmdRun                 ' 실행
    
    mskRstDate.Value = Format$(Now, "YYYY-MM-DD")
    mskRstDate1.Value = Format$(Now, "YYYY-MM-DD")
    Open App.Path + "\" + "Hitachi7180.log" For Append As #1

    Print #1, Chr(13) + Chr(10);
   
    tabWork.Tab = 0
    spdResult1.MaxRows = 0
    
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

Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol1 As Integer
    Dim intCol2 As Integer
    Dim intRow1 As Integer
    Dim intRow2 As Integer
    Dim iCnt    As Integer
    
    If Row = 0 Then Exit Sub
    
    intCol1 = pIntCol
    intCol2 = 2
    intRow1 = 1
    
    With spdResult1
        .ReDraw = False
        For iCnt = intCol1 To .MaxCols
            .Row = Row
            .Col = intCol1
            
            spdRstview.Row = intRow1
            spdRstview.Col = intCol2
            spdRstview.BackColor = vbWhite
            If .BackColor = OrderColor Then
                spdRstview.BackColor = OrderColor
            End If
            spdRstview.Text = .Text
            
            intRow1 = intRow1 + 1
            intCol1 = intCol1 + 1
            
            If intRow1 > spdRstview.MaxRows Then
                intRow1 = 1
                intCol2 = intCol2 + 2
            End If
        Next
        .ReDraw = True
    End With
End Sub

Private Sub spdResult1_KeyPress(KeyAscii As Integer)

    Dim aROW    As Integer, aCOL   As Integer
    Dim varChk  As Variant, varBar As Variant, varNum As Variant
    Dim iRow    As Integer, iCnt   As Integer
    
    If KeyAscii = vbKeyReturn Then
        With spdResult1
            aCOL = .ActiveCol
            aROW = .ActiveRow
            If aCOL = 4 Then
                iCnt = 0
                For iRow = aROW To .MaxRows
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
            End If
        End With
    End If
    
End Sub

Private Sub Timer1_Timer()

    Call COM_OUTPUT(ENQ)

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

Private Sub txtResult_DblClick()
    txtResult.Text = ""
    LstErr.Text = ""
    
    If txtResult.Visible Then txtResult.Visible = False
    LstErr.Visible = True
End Sub

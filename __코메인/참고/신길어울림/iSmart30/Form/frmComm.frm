VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
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
         TabIndex        =   25
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Print"
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
         Index           =   2
         Left            =   9225
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
         Index           =   3
         Left            =   10530
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
         TransparentPicture=   "frmComm.frx":3F0A
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   4
         Left            =   11820
         TabIndex        =   61
         Top             =   90
         Width           =   1275
         _ExtentX        =   2249
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
      Left            =   60
      TabIndex        =   7
      Top             =   600
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
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdDown"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdUpper"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdWorkList"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "SSPanel1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "pnlCom"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdSearch"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdAppend(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "pnlCom2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtBarCode"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkAuto"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "SSFrame1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "FrameError"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "UserPanel1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Crpt"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "spdResult1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdSel(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdSel(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   " 받은 결과"
      TabPicture(1)   =   "frmComm.frx":684E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "spdResult2"
      Tab(1).Control(1)=   "cmdAppend(1)"
      Tab(1).Control(2)=   "SSPanel3"
      Tab(1).Control(3)=   "cmdRstQuery"
      Tab(1).Control(4)=   "lvwCuData"
      Tab(1).Control(5)=   "cmdSel(2)"
      Tab(1).Control(6)=   "cmdSel(3)"
      Tab(1).ControlCount=   7
      Begin Threed.SSCommand cmdSel 
         Height          =   405
         Index           =   1
         Left            =   420
         TabIndex        =   10
         Top             =   810
         Width           =   345
         _Version        =   65536
         _ExtentX        =   609
         _ExtentY        =   714
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":686A
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   405
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   810
         Width           =   345
         _Version        =   65536
         _ExtentX        =   609
         _ExtentY        =   714
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":6CEC
      End
      Begin FPSpread.vaSpread spdResult1 
         Height          =   4470
         Left            =   90
         TabIndex        =   65
         Top             =   810
         Width           =   15015
         _Version        =   393216
         _ExtentX        =   26485
         _ExtentY        =   7885
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
         MaxCols         =   6
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         SpreadDesigner  =   "frmComm.frx":715A
      End
      Begin Crystal.CrystalReport Crpt 
         Left            =   6240
         Top             =   4530
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   3
         Left            =   -74640
         TabIndex        =   48
         Top             =   840
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":7630
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   -74910
         TabIndex        =   49
         Top             =   840
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":7AB2
      End
      Begin HSCotrol.UserPanel UserPanel1 
         Height          =   30
         Left            =   4320
         TabIndex        =   42
         Top             =   585
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
      Begin Threed.SSFrame FrameError 
         Height          =   2445
         Left            =   8130
         TabIndex        =   36
         Top             =   5850
         Width           =   6975
         _Version        =   65536
         _ExtentX        =   12303
         _ExtentY        =   4313
         _StockProps     =   14
         Caption         =   "Message"
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
            TabIndex        =   39
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
            TabIndex        =   38
            Top             =   225
            Visible         =   0   'False
            Width           =   6630
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   3030
         Left            =   90
         TabIndex        =   37
         Top             =   5265
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
            Height          =   2835
            Left            =   120
            TabIndex        =   59
            Top             =   120
            Width           =   7755
            _Version        =   393216
            _ExtentX        =   13679
            _ExtentY        =   5001
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
            MaxRows         =   10
            RetainSelBlock  =   0   'False
            RowsFrozen      =   10
            ScrollBarMaxAlign=   0   'False
            ScrollBars      =   0
            SpreadDesigner  =   "frmComm.frx":7F20
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
         Left            =   13710
         TabIndex        =   24
         Top             =   480
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TEST"
         Height          =   285
         Left            =   9780
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.TextBox txtBarCode 
         Height          =   300
         Left            =   1020
         MaxLength       =   12
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   1500
      End
      Begin HSCotrol.UserPanel pnlCom2 
         Height          =   4035
         Left            =   6960
         TabIndex        =   12
         Top             =   1050
         Visible         =   0   'False
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   7117
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
            Left            =   30
            TabIndex        =   13
            Top             =   3270
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
            Height          =   2955
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   21
            Top             =   300
            Width           =   5730
         End
      End
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   4830
         Left            =   -67980
         TabIndex        =   22
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
         Height          =   420
         Index           =   0
         Left            =   13860
         TabIndex        =   40
         Top             =   5400
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
         Left            =   4380
         TabIndex        =   41
         Top             =   390
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
      Begin HSCotrol.UserPanel pnlCom 
         Height          =   3885
         Left            =   1500
         TabIndex        =   29
         Top             =   1380
         Visible         =   0   'False
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   6853
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
         Begin VB.Frame Frame1 
            Height          =   645
            Left            =   90
            TabIndex        =   31
            Top             =   3180
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
            Height          =   2880
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   30
            Top             =   270
            Width           =   11595
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   465
         Left            =   90
         TabIndex        =   43
         Top             =   330
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
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
         Begin MSComCtl2.DTPicker mskOrdDate 
            Height          =   285
            Left            =   1080
            TabIndex        =   44
            Top             =   90
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            _Version        =   393216
            Format          =   83427329
            CurrentDate     =   39037
         End
         Begin MSComCtl2.DTPicker mskOrdDate1 
            Height          =   285
            Left            =   2730
            TabIndex        =   45
            Top             =   90
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            _Version        =   393216
            Format          =   83427329
            CurrentDate     =   39037
         End
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
            TabIndex        =   47
            Top             =   150
            Width           =   1095
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFC0C0&
            Caption         =   "-"
            Height          =   165
            Left            =   2550
            TabIndex        =   46
            Top             =   150
            Width           =   135
         End
      End
      Begin BHButton.BHImageButton cmdRstQuery 
         Height          =   375
         Left            =   -69240
         TabIndex        =   50
         Top             =   390
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   465
         Left            =   -74910
         TabIndex        =   51
         Top             =   330
         Width           =   5625
         _Version        =   65536
         _ExtentX        =   9922
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
            ItemData        =   "frmComm.frx":8487
            Left            =   3720
            List            =   "frmComm.frx":8494
            Style           =   2  '드롭다운 목록
            TabIndex        =   52
            Top             =   90
            Width           =   1770
         End
         Begin MSMask.MaskEdBox mskRstDate 
            Height          =   300
            Left            =   1305
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
         Begin MSMask.MaskEdBox mskRstDate1 
            Height          =   300
            Left            =   2565
            TabIndex        =   54
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
         Begin VB.Label Label8 
            BackColor       =   &H00FFC0C0&
            Caption         =   "-"
            Height          =   285
            Left            =   2430
            TabIndex        =   56
            Top             =   135
            Width           =   195
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
            TabIndex        =   55
            Top             =   165
            Width           =   1125
         End
      End
      Begin BHButton.BHImageButton cmdAppend 
         Height          =   375
         Index           =   1
         Left            =   -61170
         TabIndex        =   57
         Top             =   390
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
         Height          =   375
         Left            =   5730
         TabIndex        =   58
         Top             =   390
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
         Caption         =   "W/L 작성"
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
      Begin FPSpread.vaSpread spdResult2 
         Height          =   7410
         Left            =   -74910
         TabIndex        =   60
         Top             =   810
         Width           =   15015
         _Version        =   393216
         _ExtentX        =   26485
         _ExtentY        =   13070
         _StockProps     =   64
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
         SpreadDesigner  =   "frmComm.frx":84BE
         UserResize      =   0
      End
      Begin BHButton.BHImageButton cmdUpper 
         Height          =   375
         Left            =   7965
         TabIndex        =   62
         ToolTipText     =   "위로 이동"
         Top             =   390
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   661
         Caption         =   ""
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
         Picture         =   "frmComm.frx":893E
         PictureAlignment=   5
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdDown 
         Height          =   375
         Left            =   8550
         TabIndex        =   63
         ToolTipText     =   "아래로 이동"
         Top             =   390
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   661
         Caption         =   ""
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
         Picture         =   "frmComm.frx":8A98
         PictureAlignment=   5
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label10 
         Caption         =   "순서변경 :"
         Height          =   195
         Left            =   7035
         TabIndex        =   64
         Top             =   495
         Width           =   855
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

Private Const TEST_NM_EQP   As String = "EQP_NM"    '장비 코드
Private Const TEST_CD_LIS   As String = "LIS_CD"    '검사실 코드
Private Const TEST_NM_LIS   As String = "LIS_NM"    '검사실 이름

Const OrderColor As String = &H80FF80          '"&HC6FEFF"    '오더 배경색

Const pIntcol   As Integer = 7

Const STX As String = ""
Const ENQ As String = ""
Const ACK As String = ""

Private iSmart30(100)   As String
Private fiSmart30       As Variant

Private mAdoRs      As ADODB.Recordset
Private mAdoRs1     As ADODB.Recordset
Private CallForm    As String
Private IS_SET      As Boolean

Private f_strBuffer     As String

Private Type TYPE_CD
    strEqpCd    As String
    intCnt      As Integer
    strTestcd(100) As String
End Type
Private f_typCode() As TYPE_CD

Private RecordChk   As Boolean
Private msSeq       As String


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
            Call .Add(, TEST_CD_LIS, "검사코드", (lvwCuData.Width - 310) * 0.4)
            Call .Add(, TEST_NM_LIS, "검 사 명", (lvwCuData.Width - 310) * 0.4)
        End With
        .HideColumnHeaders = False
    End With
   
End Sub

Private Function f_subSet_WorkList(ByVal InstNo As String, ByVal ConfirmYN As String, ByVal strDate As String, ByVal strDate1 As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
        
        Set AdoRs_SQL = New ADODB.Recordset
        AdoRs_SQL.CursorLocation = adUseClient
        
        If gLisVer = "EMRLIS2" Then
            'r.BCID, r.Hcode, r.Serial, c.PtName, r.Orderdate, ErYn
            AdoRs_SQL.Open AdoCn_SQL.Execute("Exec AP_INF_Bar_Order '" & InstNo & "','" & strDate & "','" & strDate1 & "'", sqlRet)
        Else
            'r.LID, r.Hcode, r.Serial, c.PtName, r.ROrder, ErYn
            AdoRs_SQL.Open AdoCn_SQL.Execute("Exec AP_INF_S_Order '" & InstNo & "','" & ConfirmYN & "','" & strDate & "','" & strDate1 & "'", sqlRet)
        End If
        
        If sqlRet = 0 Then
            Set f_subSet_WorkList = Nothing
            RecordChk = False
            Exit Function
        Else
            Set f_subSet_WorkList = AdoRs_SQL
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
    Dim itemA   As ListItem
    
    Dim AdoRs   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strTest As String, intPos   As Integer
    Dim strTmp  As String, intCol   As Integer, intCol2   As Integer, intCnt  As Integer, intRow  As Integer
    
'On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear
    
    intCol = pIntcol
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
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"

    AdoRs.CursorLocation = adUseClient
    AdoRs.Open sqlDoc, AdoCn_Jet
    If AdoRs.RecordCount > 0 Then AdoRs.MoveFirst
    Do While Not AdoRs.EOF
        Set itemX = lvwCuData.ListItems.Add(, , Trim(AdoRs.Fields("TEST_EQP") & ""), , "LST")
            itemX.SubItems(1) = Trim(AdoRs.Fields("TESTCD") & "")
            itemX.SubItems(2) = Trim(AdoRs.Fields("TESTNM_EQP") & "")
            itemX.Tag = Trim(AdoRs.Fields("TEST_EQP") & "")
            itemX.Text = Trim(AdoRs.Fields("TESTCD") & "")
        Set itemX = Nothing

        With spdResult1
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
                .ColWidth(intCol) = 8
            End If
            .SetText intCol, 0, Trim$(AdoRs("TESTNM_EQP") & "")
            .Row = 0: .Col = intCol
            .CellTag = Trim(AdoRs.Fields("TEST_EQP") & "")
        End With
        
        With spdRstview
            If intRow > .MaxRows Then
                intRow = 1
                intCol2 = intCol2 + 2
            End If
            
            .SetText intCol2, intRow, Trim$(AdoRs("TESTNM_EQP") & "")
            intRow = intRow + 1
            
        End With
        
        With spdResult2
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
                .ColWidth(intCol) = 8
            End If
            .SetText intCol, 0, Trim$(AdoRs("TESTNM_EQP") & "")
            .Row = 0: .Col = intCol
            .CellTag = Trim(AdoRs.Fields("TEST_EQP") & "")
        End With
        
        intCnt = intCnt + 1
        ReDim Preserve f_typCode(1 To intCnt) As TYPE_CD
        
        f_typCode(intCnt).strEqpCd = Trim$(AdoRs.Fields("TEST_EQP"))
        f_typCode(intCnt).intCnt = 0
        
        strTmp = Trim$(AdoRs.Fields("TESTCD"))
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
        
        AdoRs.MoveNext
    Loop
    Set AdoRs = Nothing
    
    With spdResult2
        If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
        .SetText intCol, 0, ""
        .Col = intCol: .ColHidden = True
    End With

Exit Sub
ErrRoutine:
    Set AdoRs = Nothing
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


'워크리스트에서 필요없는 환자 삭제
Private Sub cmdWorkList_Click()
    Dim i As Long
    Dim CHK As Variant
    
    With spdResult1
        If .MaxRows < 1 Then Exit Sub
        For i = .MaxRows To 1 Step -1
            .GetText 1, i, CHK
            If CHK = 0 Or CHK = "" Then
                .DeleteRows i, 1
                .MaxRows = .MaxRows - 1
            End If
        Next i
    End With
    
    List1.AddItem ("WorkList 작성이 완료 되었습니다.")
    
End Sub

Private Sub cmdSearch_Click()
Dim intRow      As Integer, i      As Integer
Dim strBarno    As String
Dim itemX       As ListItem
Dim strEqpCd    As String
Dim intIdx      As Integer
Dim varTmp      As Variant
Dim barnochk()  As String
Dim barnoflag   As String

Dim lsOrderCoda As String
    
    Erase barnochk
   
    Screen.MousePointer = vbHourglass
    
    With spdResult1
        .ReDraw = False
        For intRow = 1 To .MaxRows
            .Row = intRow: .Col = 3
            If Trim(.Text) = "" Then
                .DeleteRows intRow, 1
                .MaxRows = .MaxRows - 1
            End If
        Next intRow
        ReDim barnochk(.MaxRows + 1)
        For intRow = 1 To .MaxRows
            .Col = 3: .Row = intRow
            barnochk(intRow) = .Text
        Next intRow
        .ReDraw = True
        intRow = .MaxRows
    End With
    
On Error GoTo ErrorTrap
    '-- WorkList조회
    Set mAdoRs = f_subSet_WorkList(INS_CODE, "0", mskOrdDate.Value, mskOrdDate1.Value)
    
    If RecordChk = False Then
'        MsgBox "해당일의 검사 대상자가 없습니다.", vbOKOnly + vbInformation, Me.Caption
        List1.AddItem ("해당일의 검사 대상자가 없습니다.")
    Else
        With spdResult1
            .ReDraw = False
            strBarno = ""
            intRow = UBound(barnochk)
            If gLisVer = "EMRLIS2" Then
                Do Until mAdoRs.EOF
                    intIdx = 0
                    
                    For i = 1 To UBound(barnochk)   '검사중 환자 체크
                        If barnochk(i) = mAdoRs.Fields("BCID") Then
                            barnoflag = "0"
                            Exit For
                        Else
                            barnoflag = "1"
                        End If
                    Next i
    
                    If strBarno <> mAdoRs.Fields("BCID") And barnoflag = "1" Then
                        intRow = intRow + 1
                        ReDim Preserve barnochk(intRow)
                        barnochk(intRow) = mAdoRs.Fields("BCID")
                        
                        If intRow > .MaxRows Then .MaxRows = .MaxRows + 1: .RowHeight(.MaxRows) = 13
                        
                        .SetText 1, intRow, "1"     '환자조회시 체크박스 체크
                        .SetText 2, intRow, mAdoRs("ORDERDATE")
                        .SetText 3, intRow, mAdoRs("BCID")
                        
                        .Row = intRow
                        If mAdoRs("ErYn") = "1" Then    '응급 환자 체크
                            .SetText 4, intRow, "ⓔ" & mAdoRs("PtName")
                            .Col = 6
                            .BackColor = &HC0E0FF
                        Else
                            .SetText 4, intRow, mAdoRs("PtName")
                            .Col = 6
                            .BackColor = &HC0FFFF
                        End If
                        .SetText 5, intRow, mAdoRs.Fields("Hcode")
    
                        '-- 검사항목조회
                        Set mAdoRs1 = New Recordset
                        Set mAdoRs1 = f_subSet_TestList(mAdoRs("BCID"))
                        
                        Do Until mAdoRs1.EOF
                            lsOrderCoda = mAdoRs1("CODA") & "/" & mAdoRs1("SUBCODA")
                            strEqpCd = f_funGet_CODE(lsOrderCoda)
                            
                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then
                                .SetText (pIntcol - 1) + itemX.Index, intRow, ""
                                .Col = (pIntcol - 1) + itemX.Index
                                .Row = intRow
                                .CellNote = lsOrderCoda & "^^"
                                .BackColor = OrderColor
                                       
                            End If
                            Set itemX = Nothing
                            mAdoRs1.MoveNext
                        Loop
                        mAdoRs1.Close: Set mAdoRs1 = Nothing
                    End If
                    strBarno = Trim(mAdoRs("BCID"))
    
                    intIdx = intIdx + 1
                    mAdoRs.MoveNext
                Loop
            Else
                Do Until mAdoRs.EOF
                    intIdx = 0
                    For i = 1 To UBound(barnochk)   '검사중 환자 체크
                        If barnochk(i) = mAdoRs.Fields("LID") Then
                            barnoflag = "0"
                            Exit For
                        Else
                            barnoflag = "1"
                        End If
                    Next i
                    
                    If strBarno <> mAdoRs.Fields("LID") And barnoflag = "1" Then
                        intRow = intRow + 1
                        ReDim Preserve barnochk(intRow)
                        barnochk(intRow) = mAdoRs.Fields("LID")
                        
                        If intRow > .MaxRows Then .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
                        
                        .SetText 1, intRow, "1"     '환자조회시 체크박스 체크
                        .SetText 2, intRow, mAdoRs("ORDERDATE")
                        .SetText 3, intRow, mAdoRs("LID")
                        
                        If mAdoRs("ErYn") = "1" Then    '응급 환자 체크
                            .SetText 4, intRow, "ⓔ" & mAdoRs("PtName")
                        Else
                            .SetText 4, intRow, mAdoRs("PtName")
                        End If
                        .SetText 5, intRow, mAdoRs.Fields("Hcode")
                        '-- 검사항목조회
                        Set mAdoRs1 = New Recordset
                        Set mAdoRs1 = f_subSet_TestList(mAdoRs("LID"))
                        
                        Do Until mAdoRs1.EOF
                            lsOrderCoda = mAdoRs1("CODA") & "/" & mAdoRs1("SUBCODA")
                            strEqpCd = f_funGet_CODE(lsOrderCoda)
                            
                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then
                                .SetText (pIntcol - 1) + itemX.Index, intRow, ""
                                .Row = intRow
                                .Col = (pIntcol - 1) + itemX.Index
                                .CellNote = lsOrderCoda & "^" & mAdoRs1("ROrder") & "^" & mAdoRs("Serial") & ""
                                .BackColor = OrderColor
                            End If
                            Set itemX = Nothing
                            mAdoRs1.MoveNext
                        Loop
                        mAdoRs1.Close: Set mAdoRs1 = Nothing
                    End If
                    strBarno = Trim(mAdoRs("Lid"))
                    
                    intIdx = intIdx + 1
                    mAdoRs.MoveNext
                Loop
            End If
            .ReDraw = True
        End With
        mAdoRs.Close: Set mAdoRs = Nothing
    End If
    Screen.MousePointer = vbDefault
Exit Sub

ErrorTrap:
    Set AdoRs_SQL = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub cmdACK_Click()

    Call COM_OUTPUT(Chr(1))
    
End Sub

Private Sub cmdAction_Click(Index As Integer)
    
    Select Case Index
        Case 0:     Call cmdRPrint
        Case 1:     Call cmdRun
        Case 2:     Call cmdStop
        Case 3:     Call cmdClear
        Case 4:     Call cmdExit
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
    
    mskOrdDate.Value = Format(Now, "YYYY-MM-DD")
    mskOrdDate1.Value = Format(Now, "YYYY-MM-DD")

    List1.Clear
    
End Sub

Private Sub cmdExit()
    
    Unload Me

End Sub

Private Sub cmdRPrint()

    Dim msgChk  As Boolean
    Dim CrptPath As String
    Dim varTmp As Variant
    Dim Paratmp As Variant
    Dim iRow As Integer
    CrptPath = DirPath & "Database\"
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
       
    UpdateODBCMDB CrptPath & "Interface.mdb"
       
    If tabWork.Tab = 1 Then
        
        If vbYes = MsgBox("검체 검사 결과를 인쇄하시겠습니까?", vbYesNo + vbInformation, INS_NAME) Then
            For iRow = 1 To spdResult2.MaxRows
                
                spdResult2.GetText 1, iRow, varTmp
                
                If Trim(varTmp) = "1" Then

                    spdResult2.GetText 3, iRow, Paratmp
                    
                    With Crpt
                        .Reset
                        .ReportFileName = CrptPath & "Resultprint.rpt"
                        .WindowState = crptMaximized
                        .Destination = crptToPrinter    '미리보기 없이 바로인쇄
    '                    .Destination = crptToWindow    '미리보기
    
                        .WindowTitle = "[" & Paratmp & "] 검체 검사 결과"
        
                        .ParameterFields(0) = "[sNo]; " & Trim(Paratmp) & "; True"
                        .ParameterFields(10) = "장비명; " & INS_NAME & "; True"
                        .ParameterFields(20) = "병원명; " & HOS_NAME & "; True"
                        
                        .Action = 1
                        .ReportFileName = ""
                    End With
                End If
            Next iRow
        Else
            Exit Sub
        End If
    Else
        If vbYes = MsgBox("WorkList를 인쇄하시겠습니까?", vbYesNo + vbInformation, INS_NAME) Then
            Call PrintFrom(spdResult1)
        End If
    End If
    Screen.MousePointer = vbDefault
    
    If Err Then
        Call ErrMsgProc(CallForm)
        On Error GoTo 0
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

Private Sub cmdAppend_Click(Index As Integer)
'    Dim AdoRs   As New ADODB.Recordset
'    Dim sqlDoc  As String
'
'    Dim varTmp  As Variant, strErrMsg   As String
'    Dim strSampleno()   As String, strBarno     As String, strTime      As String
'    Dim strRstval       As String
'    Dim intPos          As String, strTestcd    As String, strTestRst   As String
'
'    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
'    Dim itemX   As ListItem
'    Dim objSpd  As vaSpread
'    Dim strRorder As String
'    Dim intColtmp As Integer
'    Dim ListYN As Boolean
'
'    CallForm = "frmComm - Private Sub cmdAppend_Click()"
'
'On Error GoTo ErrorRoutine
'
'    Me.MousePointer = 11
'
'    If Index = 0 Then
'        Set objSpd = spdResult1
'        intColtmp = pIntcol       'WorkList와 받은결과의 걸럼수가 틀릴때..
'    Else
'        Set objSpd = spdResult2
'        intColtmp = 6
'    End If
'
'    With objSpd
'
'        For intRow = 1 To .MaxRows
'
'            .GetText 3, intRow, varTmp:   strBarno = Trim$(varTmp)
'            .GetText 4, intRow, varTmp:   pName = Trim$(varTmp)
'            .GetText 5, intRow, varTmp:   pNo = Trim$(varTmp)
'
'            If strBarno = "" Then Exit For
'
'            .GetText 1, intRow, varTmp
'            If Trim$(varTmp) = "1" Then
'                ListYN = False
'                For intCol = intColtmp To .MaxCols
'                    .GetText intCol, intRow, varTmp
'                    If Trim$(varTmp) <> "" Then
'                        .GetText intCol, 0, varTmp
'                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
'                        If Not itemX Is Nothing Then
'                            .GetText intCol, intRow, varTmp
'                            strRstval = Trim(varTmp)
'                            strTestcd = itemX.ListSubItems(1)
'                            intPos = InStr(strTestcd, ",")
'                            If intPos > 0 Then
'                                Dim tmpTestCd   As Variant
'                                Dim tmpTestSeq1 As Integer
'                                Dim tmpTestCd1  As String
'
'                                tmpTestCd = Split(strTestcd, ",")
'                                blnFlag = False
'                                Set mAdoRs = f_subSet_TestList(strBarno)
'                                For tmpTestSeq1 = 1 To UBound(tmpTestCd) + 1
'                                    tmpTestCd1 = tmpTestCd(tmpTestSeq1 - 1)
'                                    mAdoRs.MoveFirst
'                                    Do Until mAdoRs.EOF
'                                        If Trim(mAdoRs("CODA") & "/" & mAdoRs("SUBCODA")) = tmpTestCd1 Then
'                                            blnFlag = True
'                                            strRorder = Trim(mAdoRs("RORDER"))
'                                            txtResult.Text = txtResult.Text & strRorder
'                                            Exit For
'                                        End If
'                                        mAdoRs.MoveNext
'                                    Loop
'                                Next
'                                If gLisVer = "EMRLIS2" Then
'                                    sqlDoc = "AP_INF_Bar_Result '" & strBarno & "', "
'                                    sqlDoc = sqlDoc & " '" & INS_CODE & "', '" & Mid(tmpTestCd1, 1, InStr(tmpTestCd1, "/") - 1) & "', '" & Mid(tmpTestCd1, InStr(tmpTestCd1, "/") + 1) & "',"
'                                    sqlDoc = sqlDoc & " '" & strRstval & "'"
'                                Else
'                                    sqlDoc = "AP_INF_S_Update " & strBarno & ", " & Trim(mAdoRs("Serial")) & ", " & strRorder & ","
'                                    sqlDoc = sqlDoc & " '" & INS_CODE & "', '" & Mid(tmpTestCd1, 1, InStr(tmpTestCd1, "/") - 1) & "', '" & Mid(tmpTestCd1, InStr(tmpTestCd1, "/") + 1) & "',"
'                                    sqlDoc = sqlDoc & " '" & strRstval & "'"
'                                End If
'                                AdoCn_SQL.Execute sqlDoc
'                                lblStatus.Caption = "저장 성공!!"
'
'                                Set AdoRs = Nothing:    mAdoRs.Close
'                            Else
'                                blnFlag = False
'                                Set mAdoRs = f_subSet_TestList(strBarno)
'                                Do Until mAdoRs.EOF
'                                    If Trim(mAdoRs("CODA") & "/" & mAdoRs("SUBCODA")) = strTestcd Then
'                                        blnFlag = True
'                                        strRorder = Trim(mAdoRs("RORDER"))
'                                        Exit Do
'                                    End If
'                                    mAdoRs.MoveNext
'                                Loop
'                                If blnFlag Then
'                                    If gLisVer = "EMRLIS2" Then
'                                        sqlDoc = "AP_INF_Bar_Result '" & strBarno & "', "
'                                        sqlDoc = sqlDoc & " '" & INS_CODE & "', '" & Mid(strTestcd, 1, InStr(strTestcd, "/") - 1) & "', '" & Mid(strTestcd, InStr(strTestcd, "/") + 1) & "',"
'                                        sqlDoc = sqlDoc & " '" & strRstval & "'"
'                                    Else
'                                        sqlDoc = "AP_INF_S_Update " & strBarno & ", " & Trim(mAdoRs("Serial")) & ", " & strRorder & ","
'                                        sqlDoc = sqlDoc & " '" & INS_CODE & "', '" & Mid(strTestcd, 1, InStr(strTestcd, "/") - 1) & "', '" & Mid(strTestcd, InStr(strTestcd, "/") + 1) & "',"
'                                        sqlDoc = sqlDoc & " '" & strRstval & "'"
'                                    End If
'                                    AdoCn_SQL.Execute sqlDoc
'                                    lblStatus.Caption = "저장 성공!!"
'                                    Set AdoRs = Nothing:    mAdoRs.Close
'                                End If
'                            End If
'                            spdResult1.Row = intRow
'                            spdResult1.Col = 2
'                            spdResult1.BackColor = vbCyan
'                            spdResult1.Col = 3
'                            spdResult1.BackColor = vbCyan
'                            spdResult1.Col = 4
'                            spdResult1.BackColor = vbCyan
'                            spdResult1.Col = 5
'                            spdResult1.BackColor = vbCyan
'                            spdResult1.Col = 1: spdResult1.Value = 0
'
'                            .SetText 0, intRow, "완료"
'                            ListYN = True
'
'                            If strErrMsg = "" Then
'                                sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
'                                         " where SPCNO   = '" & strBarno & "'"
''                                         "   and TRANSDT = '" & strDate & "'"
'                                AdoCn_Jet.Execute sqlDoc
'                            Else
'                                MsgBox strErrMsg, vbInformation, INS_NAME
'                            End If
'                        End If
'
'                        Set itemX = Nothing
'                    End If
'                Next
'                If ListYN = True Then List1.AddItem ("검체번호 : " & strBarno & "의 결과값 등록이 완료되었습니다.")
'            End If
'        Next
'
'    End With
'    Me.MousePointer = 0
'
''    MsgBox "작업이 완료되었습니다.", vbInformation, Me.Caption
'
'    Exit Sub
'ErrorRoutine:
'    Set itemX = Nothing
'
'    Me.MousePointer = 0
'    Call ErrMsgProc(CallForm)
End Sub

Private Sub cmdENQ_Click()
    
    Call COM_OUTPUT(charCOM_Convert(COM_ENQ))

End Sub

Private Sub cmdRstQuery_Click()

    Dim AdoRs   As New ADODB.Recordset
    Dim sqlDoc  As String, intRet   As Integer
    
    Dim strSpcno    As String
    Dim intRow      As Integer, intCol  As Integer
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
    
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
             " where TRANSDT >= '" & mskRstDate.Text & "'" & _
             "   and EQUIPCD = '" & INS_CODE & "'"
    If cboRstgbn(1).ListIndex = 0 Then
        sqlDoc = sqlDoc & "   and SERVERGBN = ''"
    ElseIf cboRstgbn(1).ListIndex = 1 Then
        sqlDoc = sqlDoc & "   and SERVERGBN = 'Y'"
    End If
    sqlDoc = sqlDoc & " order by SPCNO, TRANSTM"
    
    AdoRs.CursorLocation = adUseClient
    AdoRs.Open sqlDoc, AdoCn_Jet
    If AdoRs.RecordCount > 0 Then AdoRs.MoveFirst
    Do While Not AdoRs.EOF
        With spdResult2
        If strSpcno <> Trim$(AdoRs(0) & "") + Trim$(AdoRs(6) & "") Then
                intRow = intRow + 1
                If intRow > .MaxRows Then .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
'                    .SetText 1, intRow, "1"
                    .SetText 2, intRow, Trim$(AdoRs(3) & "")
                    .SetText 3, intRow, Trim$(AdoRs(0) & "")
                    .SetText 4, intRow, Trim$(AdoRs(8) & "")
                    .SetText 5, intRow, Trim$(AdoRs(9) & "")
        End If
                strSpcno = Trim$(AdoRs(0) & "") + Trim$(AdoRs(6) & "")
                Set itemX = lvwCuData.FindItem(Trim$(AdoRs(7) & ""), lvwTag, , lvwWhole)
                If Not itemX Is Nothing Then
                    intCol = itemX.Index + 5
                    .SetText intCol, intRow, Trim$(AdoRs(4)) & ""
'                    .Col = intCol:  .Row = intRow:  .ForeColor = IIf(Trim$(adoRS(5) & "") <> "", vbRed, vbBlack)
                End If
        End With
        AdoRs.MoveNext
    Loop
    AdoRs.Close:    Set AdoRs = Nothing
    
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
    Else
        With spdResult1
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
    CallForm = "clsCommon - Public Function f_subSet_DateOrder() As ADODB.Recordset"
    
    Set AdoRs_SQL = New ADODB.Recordset
    AdoRs_SQL.CursorLocation = adUseClient
    
    If gLisVer = "EMRLIS2" Then
        AdoRs_SQL.Open AdoCn_SQL.Execute("Exec AP_INF_Bar_Order_Coda '" & INS_CODE & "', '" & strBarcode & "'", sqlRet)
    Else
        AdoRs_SQL.Open AdoCn_SQL.Execute("Exec AP_INF_S_GetCoda '" & INS_CODE & "', '" & strBarcode & "'", sqlRet)
    End If
    
    If sqlRet = 0 Then
        Set f_subSet_TestList = Nothing
'        MsgBox strBarcode & "-해당검체는 검사가 완료되었습니다.", vbOKOnly + vbExclamation
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

Private Sub cmdUpper_Click()
    With spdResult1
    
        .SetFocus
        If .ActiveRow = 0 Or .ActiveRow = 1 Then Exit Sub
        
        .SwapRowRange .ActiveRow, .ActiveRow, .ActiveRow - 1
        .SetActiveCell .ActiveCol, .ActiveRow - 1

    End With
End Sub

Private Sub cmdDown_Click()
    With spdResult1
        
        .SetFocus
        If .ActiveRow = 0 Or .ActiveRow = .MaxRows Then Exit Sub
        
        .SwapRowRange .ActiveRow, .ActiveRow, .ActiveRow + 1
        .SetActiveCell .ActiveCol, .ActiveRow + 1

    End With
End Sub

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
            strDta = ""
                                        
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
Dim liStart         As Long     'iSmart30 -> STX
Dim liEnd           As Long     'iSmart30 -> CR + LF
Dim lsBuffer        As String
Dim f_strBuffer     As String

Static lsRecData    As String
    
    lsRecData = lsRecData & RecData

    Print #1, RecData;
    
    Call COM_INPUT(RecData)
'    Debug.Print strTmp
    
    liStart = InStr(lsRecData, STX)
    liEnd = InStr(lsRecData, vbCrLf) + 1
    If liStart > 0 And liEnd > 0 And liStart < liEnd Then
        f_strBuffer = Mid(lsRecData, liStart, liEnd - liStart)
        lsRecData = Mid(lsRecData, liEnd + 1)
        Call ReceiveTheData(f_strBuffer)
        Call COM_OUTPUT(ACK)
    End If
    
    If InStr(lsRecData, ENQ) > 0 Then
'        Debug.Print lsRecData
        Call COM_OUTPUT(ACK)
    End If
End Sub

Private Sub ReceiveTheData(ByVal strdata As String)
Dim strRstval   As String
Dim sqlDoc      As String
Dim varTmp, strBarno, strDate, strTime As String

Dim Channel_No  As String       ' 검사항목 번호 : 시작번호
Dim pGrid_Point As Integer
Dim sTemp       As String
Dim lvData      As Variant
Dim lvSeq       As Variant
Dim lvResult    As Variant
Dim lvKey       As Variant
Dim lsOrderCoda As String
Dim lsRorder    As String
Dim lsSerial    As String
Dim lsEQCode    As String
Dim lsPNo       As String
Dim lsPName     As String
Dim liRow       As Integer
Dim liCol       As Integer
    
    On Error Resume Next
       
    CallForm = "frmInterface - Privete sub psDataDefine()"
    sTemp = strdata
    
'    Debug.Print strdata
    Erase fiSmart30
    
    fiSmart30 = Split(sTemp, "|")
    
'    Debug.Print UBound(fiSmart30)
    pGrid_Point = 0
    Select Case Mid(fiSmart30(0), 3, 1)
        Case "H"
            msSeq = ""
        Case "P"
            msSeq = Val(fiSmart30(3))
        Case "R"
            If Len(msSeq) > 0 Then 'fiSmart30(3) = ID
                With spdResult1
                    pGrid_Point = SeqSearch(spdResult1, "", 6)
                    If pGrid_Point > 0 Then
                        .GetText 3, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)
                        .GetText 4, pGrid_Point, varTmp:   lsPName = Trim$(varTmp)
                        .GetText 5, pGrid_Point, varTmp:   lsPNo = Trim$(varTmp)
                        
                        For liCol = pIntcol To .MaxCols
                            .Row = 0: .Col = liCol: lsEQCode = .CellTag
                            strRstval = ""
                            Channel_No = Replace(Replace(Trim(fiSmart30(2)), "^^^", ""), "^M", "")
                            If Trim(UCase(Channel_No)) = Trim(UCase(lsEQCode)) Then
                                strRstval = Trim(fiSmart30(3))
                                        
                                strDate = Format$(Now, "YYYYMMDD"):     strTime = Format$(Now, "hhmmss")
                                .SetText liCol, pGrid_Point, strRstval
                                .Col = liCol:  .Row = pGrid_Point
                                If Trim(.CellNote) = "" Then
                                    lsOrderCoda = ""
                                    lsRorder = ""
                                    lsSerial = ""
                                Else
                                    lvKey = Split(.CellNote, "^")
                                    lsOrderCoda = lvKey(0)
                                    lsRorder = lvKey(1)
                                    lsSerial = lvKey(2)
                                End If
        '                        Debug.Print lsOrderCoda & " - " & lsRorder & " - " & lsSerial & vbCrLf
                                
                                sqlDoc = "insert into INTERFACE003(SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                         " values( '" & strBarno & "', '" & lsOrderCoda & "', '" & lsEQCode & "'," & _
                                         "         '" & strDate & "', '" & strTime & "','" & strRstval & "', ''," & _
                                         "         '" & INS_CODE & "', '', '" & lsPName & "', '" & lsPNo & "')"
                                AdoCn_Jet.Execute sqlDoc
    
                                If chkAuto.Value = vbChecked And lsOrderCoda <> "" Then
                                    If gLisVer = "EMRLIS2" Then
                                        sqlDoc = "AP_INF_Bar_Result '" & strBarno & "', "
                                        sqlDoc = sqlDoc & " '" & INS_CODE & "', '" & Mid(lsOrderCoda, 1, InStr(lsOrderCoda, "/") - 1) & "', '" & Mid(lsOrderCoda, InStr(lsOrderCoda, "/") + 1) & "',"
                                        sqlDoc = sqlDoc & " '" & strRstval & "'"
                                    Else
                                        sqlDoc = "AP_INF_S_Update " & strBarno & ", " & lsSerial & ", " & lsRorder & ","
                                        sqlDoc = sqlDoc & " '" & INS_CODE & "', '" & Mid(lsOrderCoda, 1, InStr(lsOrderCoda, "/") - 1) & "', '" & Mid(lsOrderCoda, InStr(lsOrderCoda, "/") + 1) & "',"
                                        sqlDoc = sqlDoc & " '" & strRstval & "'"
                                    End If
                                    AdoCn_SQL.Execute sqlDoc
                                End If
                                Exit For
                            End If
                        Next liCol
                    End If
                
                End With
            End If
        Case "L"
            With spdResult1
                pGrid_Point = SeqSearch(spdResult1, "", 6)
                .GetText 3, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)
                .SetText 6, pGrid_Point, msSeq
                If chkAuto.Value = "1" Then
                    .Row = pGrid_Point
                    .Col = 2: .BackColor = vbCyan
                    .Col = 3: .BackColor = vbCyan
                    .Col = 4: .BackColor = vbCyan
                    .Col = 5: .BackColor = vbCyan
                    .Col = 1: .Value = 0
        
                    sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y' where SPCNO = '" & strBarno & "'"
        
                    AdoCn_Jet.Execute sqlDoc
                    List1.AddItem ("검체번호 : " & strBarno & "의 결과값 등록이 완료되었습니다.")
                End If
            End With
            msSeq = ""
    End Select
    Exit Sub

ErrRoutine:

    Call ErrMsgProc(CallForm)

End Sub

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
            If Trim(.Text) = brSeq Then
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

    strDta = strDta & vbLf
    '1H|\^&|||iSmart30^iSmart30343^-^1.0.4.1 EX R3||||||||1394-97|2011040711384615
    '2P|1|||||||||||||||||||||||||||||||||2E
    '3O|1||014690-S8||||||||||||Sample|||||||||||||||8C
    '1H|\^&|||iSmart30^iSmart30343^-^1.0.4.1 EX R3||||||||1394-97|2011040711390612
    '2P|1|||||||||||||||||||||||||||||||||2E
    '3O|1||014690-S8||||||||||||Sample|||||||||||||||8C
    '4R|1|^^^Na+^M|142|mmol/L||N||||||20110407105615|7B
    '5R|2|^^^K+^M|4.7|mmol/L||N|||||||5A
    '6R|3|^^^Cl-^M|98|mmol/L||N|||||||9A
    '7R|4|^^^Hct^M|Out of Range(L)|%||L|||||||38
    '0L|1|NF6
    '
    
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
    
    mskRstDate.Text = Format$(Now, "YYYYMMDD")
    mskOrdDate.Value = Format$(Now, "YYYY-MM-DD")
    mskOrdDate1.Value = Format$(Now, "YYYY-MM-DD")
    mskRstDate1.Text = Format$(Now, "YYYYMMDD")
    cboRstgbn(1).ListIndex = 0
    
    Open App.Path + "\" + "iSmart30.log" For Append As #1

    Print #1, Chr(13) + Chr(10);
   
    tabWork.Tab = 0
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
        MsgBox INS_CODE & " 에 대한 장비 통신 구성이 없습니다. 통신 설정후 다시 시도 하십시오.", vbExclamation, INS_NAME
        Exit Sub
    Else
        If mAdoRs.EOF Then
            IS_SET = False
            MsgBox INS_CODE & " 에 대한 장비 통신 구성이 없습니다. 통신 설정후 다시 시도 하십시오.", vbExclamation, INS_NAME
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
        .SelLength = Len(.Text) + 2
    End With '
    
End Sub


Private Sub mskRstDate_KeyPress(KeyAscii As Integer)

    If Not KeyAscii = vbKeyBack Then mskRstDate.SelLength = 1
    
End Sub

Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol1 As Integer
    Dim intCol2 As Integer
    Dim intRow1 As Integer
    Dim intRow2 As Integer
    Dim iCnt    As Integer
    
    If Row = 0 Then Exit Sub
    
    intCol1 = pIntcol
    intCol2 = 2
    intRow1 = 1
    
    With spdResult1
        
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
    End With
End Sub

Private Sub spdRstview_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    Dim ordcol As Integer, ordrow As Integer, i As Integer
    Dim colChk As Variant
    
    With spdRstview
        .Col = .ActiveCol: .Row = .ActiveRow
        ordcol = .Col:  ordrow = .Row
        If .BackColor = OrderColor Then
            .BackColor = vbWhite
        ElseIf .BackColor = vbWhite Then
            .BackColor = OrderColor
        End If
    End With
    
    With spdResult1
       
        For i = .ActiveCol To .MaxCols
            .GetText i, 0, colChk
            spdRstview.Col = ordcol - 1: spdRstview.Row = ordrow
            If colChk = spdRstview.Text Then
                .Col = i:   .Row = .ActiveRow
                If .BackColor = OrderColor Then
                    .BackColor = vbWhite
                ElseIf .BackColor = vbWhite Then
                    .BackColor = OrderColor
                End If
                Exit For
            End If
        Next i
        
    End With

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
    List1.Text = ""
    
    If txtResult.Visible Then txtResult.Visible = False
    List1.Visible = True
End Sub

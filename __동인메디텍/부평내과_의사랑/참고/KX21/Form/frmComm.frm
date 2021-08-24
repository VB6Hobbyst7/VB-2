VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
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
         TabIndex        =   27
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
         TabIndex        =   28
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
         TabIndex        =   29
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
         TabIndex        =   30
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
      Tab(0).Control(1)=   "cmdStartNo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdPosNo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdRackNo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "pnlCom"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdWordQuery"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdEot"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdOrder"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdSearch"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FrameResult"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "pnlCom2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtBarCode"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkAuto"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chkReTest"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "SSPanel1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "SSFrame1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "FrameError"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "SSPanel2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "spdResult1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdWorkList"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "spdWorklist"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdSel(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdSel(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdAppend(0)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lvwCuData"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdSel(3)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdSel(2)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   " 받은 결과"
      TabPicture(1)   =   "frmComm.frx":684E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboRstgbn(1)"
      Tab(1).Control(1)=   "mskRstDate"
      Tab(1).Control(2)=   "cmdRstQuery"
      Tab(1).Control(3)=   "cmdAppend(1)"
      Tab(1).Control(4)=   "spdResult2"
      Tab(1).Control(5)=   "Label4"
      Tab(1).ControlCount=   6
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   4830
         TabIndex        =   71
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":686A
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   3
         Left            =   5100
         TabIndex        =   70
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":6CD8
      End
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   4830
         Left            =   8670
         TabIndex        =   69
         Top             =   2070
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
         Left            =   13800
         TabIndex        =   48
         Top             =   420
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
      Begin Threed.SSCommand cmdSel 
         Height          =   345
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   609
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":715A
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   609
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":75DC
      End
      Begin FPSpread.vaSpread spdWorklist 
         Height          =   3930
         Left            =   90
         TabIndex        =   68
         Top             =   900
         Width           =   4755
         _Version        =   196608
         _ExtentX        =   8387
         _ExtentY        =   6932
         _StockProps     =   64
         BackColorStyle  =   1
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
         OperationMode   =   1
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmComm.frx":7A4A
         UserResize      =   0
      End
      Begin BHButton.BHImageButton cmdWorkList 
         Height          =   390
         Left            =   90
         TabIndex        =   31
         Top             =   4875
         Width           =   4710
         _ExtentX        =   8308
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
         ImgOutLineSize  =   3
      End
      Begin FPSpread.vaSpread spdResult1 
         Height          =   4380
         Left            =   4830
         TabIndex        =   64
         Top             =   900
         Width           =   10185
         _Version        =   196608
         _ExtentX        =   17965
         _ExtentY        =   7726
         _StockProps     =   64
         BackColorStyle  =   1
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
         SpreadDesigner  =   "frmComm.frx":7EF2
         UserResize      =   0
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   465
         Left            =   6360
         TabIndex        =   55
         Top             =   360
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
            TabIndex        =   57
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
            TabIndex        =   56
            Top             =   90
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin Threed.SSFrame FrameError 
         Height          =   2940
         Left            =   8130
         TabIndex        =   42
         Top             =   5280
         Width           =   6975
         _Version        =   65536
         _ExtentX        =   12303
         _ExtentY        =   5186
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
            Height          =   2580
            Left            =   150
            TabIndex        =   47
            Top             =   210
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
            Height          =   2580
            Left            =   135
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   46
            Top             =   225
            Visible         =   0   'False
            Width           =   6630
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   3030
         Left            =   90
         TabIndex        =   43
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
            Height          =   2880
            Left            =   45
            TabIndex        =   44
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
            SpreadDesigner  =   "frmComm.frx":8378
            UserResize      =   0
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   465
         Left            =   90
         TabIndex        =   49
         Top             =   390
         Width           =   3540
         _Version        =   65536
         _ExtentX        =   6244
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
         Begin VB.ComboBox cboChk 
            Height          =   300
            ItemData        =   "frmComm.frx":88FF
            Left            =   3600
            List            =   "frmComm.frx":8909
            TabIndex        =   63
            Top             =   90
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.ComboBox cboComNm 
            Height          =   300
            ItemData        =   "frmComm.frx":8919
            Left            =   4755
            List            =   "frmComm.frx":891B
            TabIndex        =   62
            Top             =   90
            Visible         =   0   'False
            Width           =   1545
         End
         Begin MSMask.MaskEdBox mskOrdDate1 
            Height          =   300
            Left            =   2355
            TabIndex        =   50
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
            Left            =   1110
            TabIndex        =   51
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
            Left            =   2220
            TabIndex        =   53
            Top             =   150
            Width           =   315
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
            TabIndex        =   52
            Top             =   150
            Width           =   1095
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
         Left            =   10890
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   1380
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
         Left            =   12270
         TabIndex        =   26
         Top             =   540
         Width           =   1410
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TEST"
         Height          =   285
         Left            =   9780
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.ComboBox cboRstgbn 
         Height          =   300
         Index           =   1
         ItemData        =   "frmComm.frx":891D
         Left            =   -72570
         List            =   "frmComm.frx":892A
         Style           =   2  '드롭다운 목록
         TabIndex        =   9
         Top             =   495
         Width           =   1770
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
      Begin MSMask.MaskEdBox mskRstDate 
         Height          =   300
         Left            =   -73695
         TabIndex        =   10
         Top             =   495
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin HSCotrol.UserPanel pnlCom2 
         Height          =   3615
         Left            =   5895
         TabIndex        =   15
         Top             =   1005
         Visible         =   0   'False
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   6376
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
            Left            =   -450
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   24
            Top             =   -90
            Width           =   5730
         End
         Begin VB.Frame Frame2 
            Height          =   645
            Left            =   90
            TabIndex        =   16
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
               TabIndex        =   17
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
               TabIndex        =   18
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
               TabIndex        =   19
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
               TabIndex        =   20
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
               TabIndex        =   21
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
               TabIndex        =   22
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
               TabIndex        =   23
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
      Begin BHButton.BHImageButton cmdRstQuery 
         Height          =   375
         Left            =   -61095
         TabIndex        =   33
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
      Begin BHButton.BHImageButton cmdAppend 
         Height          =   375
         Index           =   1
         Left            =   -62355
         TabIndex        =   32
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
      Begin Threed.SSFrame FrameResult 
         Height          =   3525
         Left            =   285
         TabIndex        =   41
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
      Begin BHButton.BHImageButton cmdSearch 
         Height          =   375
         Left            =   3705
         TabIndex        =   54
         Top             =   450
         Width           =   1140
         _ExtentX        =   2011
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
      Begin BHButton.BHImageButton cmdOrder 
         Height          =   375
         Left            =   8550
         TabIndex        =   58
         Top             =   -45
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
      Begin BHButton.BHImageButton cmdEot 
         Height          =   375
         Left            =   12285
         TabIndex        =   59
         Top             =   -90
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
         Left            =   8145
         TabIndex        =   61
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
      Begin HSCotrol.UserPanel pnlCom 
         Height          =   2565
         Left            =   -2025
         TabIndex        =   34
         Top             =   1440
         Visible         =   0   'False
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   4524
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
            TabIndex        =   35
            Top             =   315
            Visible         =   0   'False
            Width           =   11595
         End
         Begin VB.Frame Frame1 
            Height          =   645
            Left            =   45
            TabIndex        =   36
            Top             =   4020
            Visible         =   0   'False
            Width           =   11610
            Begin HSCotrol.CButton cmdCOMSave 
               Height          =   360
               Left            =   10515
               TabIndex        =   37
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
               TabIndex        =   38
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
               TabIndex        =   39
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
               TabIndex        =   40
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
      Begin BHButton.BHImageButton cmdRackNo 
         Height          =   375
         Left            =   6435
         TabIndex        =   65
         Top             =   -45
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
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
         Height          =   375
         Left            =   7695
         TabIndex        =   66
         Top             =   -45
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
         Left            =   4890
         TabIndex        =   67
         Top             =   450
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
      Begin FPSpread.vaSpread spdResult2 
         Height          =   7350
         Left            =   -74910
         TabIndex        =   60
         Top             =   930
         Width           =   15015
         _Version        =   196608
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
         SpreadDesigner  =   "frmComm.frx":8954
         UserResize      =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Left            =   -74910
         TabIndex        =   12
         Top             =   570
         Width           =   1125
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
         TabIndex        =   11
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

Dim AVL9180(100)   As String
Dim fAVL9180       As Variant
Dim fAVL9180_1     As Variant
Dim fAVL9180_2     As Variant
Dim fAVL9180_3     As Variant
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


Dim fKX21(100) As String
Dim fKX21Cfg(100) As Integer
Dim fKX21Size(100, 1) As Integer
Dim fRcvString As String
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
Dim strGumCd As String


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

Private Function f_subSet_WorkList(ByVal strDate As String, ByVal strDate1 As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
   
        Set AdoRs_SQL = New ADODB.Recordset
            
                 sqlDoc = "SELECT a.ORDDT, a.SPCNO, a.PTID,a.PTNM,b.TESTCD"
        sqlDoc = sqlDoc & "  FROM TC201 a, TC301 b"
        sqlDoc = sqlDoc & " WHERE a.ORDDT BETWEEN '" & mskOrdDate.Text & "' AND '" & mskOrdDate1.Text & "'"
        sqlDoc = sqlDoc & "   AND a.SPCNO = b.SPCNO"
        sqlDoc = sqlDoc & "   AND a.PTID = b.PTID"
        sqlDoc = sqlDoc & "   AND b.TESTCD in (" & strGumCd & ")" & vbLf
        sqlDoc = sqlDoc & "   AND b.RSTVAL = ' '" & vbLf
        sqlDoc = sqlDoc & "   AND a.STSCD > '1'" & vbLf
        sqlDoc = sqlDoc & "   AND a.CANCELFG = '0'" & vbLf
        sqlDoc = sqlDoc & " ORDER BY a.ORDDT, a.SPCNO, b.TESTCD"
        
        Set AdoRs_SQL = New ADODB.Recordset
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
        
        If AdoRs_SQL.RecordCount = 0 Then
            Set f_subSet_WorkList = Nothing
            RecordChk = False
            Set AdoRs_SQL = Nothing
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

Private Function f_subSet_WorkList_Barcode(ByVal strBarno As String, Optional ByVal strStatus As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
   
        Set AdoRs_SQL = New ADODB.Recordset
                       
                 sqlDoc = "SELECT a.ORDDT, a.SPCNO, a.PTID,a.PTNM,b.TESTCD"
        sqlDoc = sqlDoc & "  FROM TC201 a, TC301 b"
        sqlDoc = sqlDoc & " WHERE a.ORDDT BETWEEN '" & mskOrdDate.Text & "' AND '" & mskOrdDate1.Text & "'"
        sqlDoc = sqlDoc & "   AND a.SPCNO = b.SPCNO"
        sqlDoc = sqlDoc & "   AND a.PTID = b.PTID"
        sqlDoc = sqlDoc & "   AND b.TESTCD in (" & strGumCd & ")" & vbLf
        sqlDoc = sqlDoc & "   AND b.SPCNO = '" & strBarno & "'" & vbLf
        sqlDoc = sqlDoc & " ORDER BY a.ORDDT, a.SPCNO, b.TESTCD"

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
'Private Function f_subSet_WorkList_Barcode(ByVal strBarno As String)
'    Dim sqlRet      As Integer
'    Dim sqlDoc      As String
'
'On Error GoTo ErrorTrap
'    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
'
'
'        Set AdoRs_SQL = New ADODB.Recordset
'
'        If Len(strBarno) > 8 Then
'            sqlDoc = " SELECT a.per_gumjin_date, a.per_gum_num, a.edpscode, a.result, a.send_date, a.per_name " & _
'                    " FROM mdck..gumjin_interface a, mdck..bag_interfacecode b " & _
'                    " WHERE substring(a.per_gumjin_date,3,8) = '" & Mid(strBarno, 1, 6) & "'" & _
'                    " AND a.per_gum_num = '" & Val(Mid(strBarno, 7)) & "' " & _
'                    " AND a.result = '' " & _
'                    " AND substring(b.kind,1,1) = 'C' " & _
'                    " AND a.edpscode=b.meditem " & _
'                    " ORDER BY a.per_gumjin_date, a.per_gum_num "
'        Else
'            sqlDoc = " SELECT a.EnterDate, b.Status, b.waitseqno, b.MAP2SEQNO, b.DispDesc, b.RVALUEKIND, b.NORMLOW, b.NORMHIGH, b.NORMALVALUE, b.RVALUEKIND , " & _
'                    " a.ChartNo, b.GumsaKind, c.sujinname, b.status " & _
'                    " FROM medicom..WaitPrsnp a, medicom..jun370_resulttb b, medicom..pewprsnp c, medicom..BAGMAP2PREF d " & _
'                    " WHERE a.Chartno = '" & strBarno & "' " & _
'                    " AND a.WaitSeqNo = b.WaitSeqNo " & _
'                    " AND a.status = '1' " & _
'                    " AND d.labno = 4 " & _
'                    " AND b.jun370no = d.map2seqno " & _
'                    " AND b.status = '0' " & _
'                    " AND a.chartno = c.chartno " & _
'                    " ORDER BY a.chartno "
'        End If
'
'        Set AdoRs_SQL = New ADODB.Recordset
'        AdoRs_SQL.CursorLocation = adUseClient
'        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
'
'        If AdoRs_SQL.RecordCount = 0 Then
'            Set f_subSet_WorkList_Barcode = Nothing
'            RecordChk = False
'            Set AdoRs_SQL = Nothing
'            Exit Function
'        Else
'            Set f_subSet_WorkList_Barcode = AdoRs_SQL
'            RecordChk = True
'        End If
'
'        Set AdoRs_SQL = Nothing
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
        
'On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = ""
    
    intCol = 8
    intCol2 = 1
    intRow = 1
    With spdWorklist
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdResult2
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
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
                strJinCd = strJinCd & "" & Mid(Trim(adoRS.Fields("TESTCD")), intPos1 + 1) & ","
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
        
        fChannel(intCol - 7) = adoRS.Fields("TEST_EQP")
        
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

Private Function f_subSet_ComList()
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_ComList() As ADODB.Recordset"
    
   
        Set AdoRs_SQL = New ADODB.Recordset
        
        sqlDoc = "         SELECT B.COM_CODE, B.COM_NAME " & vbCr
        sqlDoc = sqlDoc & "  FROM MDCK..GUMJIN_INTERFACE A, MDCK..TB_COMPANY B, MDCK..BAG_INTERFACECODE C " & vbCr
        sqlDoc = sqlDoc & " WHERE A.Per_COM_Code = B.COM_CODE " & vbCr
        sqlDoc = sqlDoc & "   AND A.per_gumjin_date BETWEEN '" & Trim(mskOrdDate.Text) & "' AND '" & Trim(mskOrdDate1.Text) & "'" & vbCr
        sqlDoc = sqlDoc & "   AND SUBSTRING(C.KIND, 1, 1) = 'I' " & vbCr
        sqlDoc = sqlDoc & "   AND A.EDPSCODE = C.MEDITEM " & vbCr
        sqlDoc = sqlDoc & " GROUP BY B.COM_CODE, B.COM_NAME " & vbCr
        
        Set AdoRs_SQL = New ADODB.Recordset
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
        
        If AdoRs_SQL.RecordCount > 0 Then
            AdoRs_SQL.MoveFirst
            cboComNm.Clear
            Do Until AdoRs_SQL.EOF
                cboComNm.AddItem AdoRs_SQL.Fields("COM_NAME") & ""
                AdoRs_SQL.MoveNext
            Loop
            cboComNm.ListIndex = 0
        End If
        
        AdoRs_SQL.Close:  Set AdoRs_SQL = Nothing
    
Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    
    Call ErrMsgProc(CallForm)

    
End Function

Private Sub cboChk_Click()
    If Trim(cboChk.Text) = "검진" Then
        Call f_subSet_ComList
    Else
        cboComNm.Clear
    End If
End Sub

Private Sub cmdEot_Click()
    Call COM_OUTPUT(EOT)
End Sub

Private Sub cmdOrder_Click()
    Dim varTmp      As Variant
    Dim intRow      As Integer, intCol  As Integer
    Dim strBarno    As String, strTest  As String
    Dim strName     As String
    Dim strRack     As String, strCup   As String
    Dim intCnt      As Integer
    Dim itemX       As ListItem
    Dim strOrdList  As String
    Dim TestStatus(32)  As String * 1
    Dim xx              As Integer
    Dim intchannel As Integer
    Dim ssTemp2    As String
    Dim Loop_Count, pDoCount As Integer

    For xx = 1 To 32
        TestStatus(xx) = "0"
    Next xx
    
    With spdResult1
        For intRow = 1 To .maxrows
            .Row = intRow
            .Col = 2
            If .BackColor = vbWhite Then
                intCnt = 0
                .GetText 3, intRow, varTmp: strBarno = Trim$(varTmp)
                .GetText 6, intRow, varTmp: strName = Trim$(varTmp)
                .GetText 7, intRow, varTmp: strRack = Trim$(varTmp)
                For intCol = 8 To .MaxCols
                    spdResult1.GetText intCol, 0, varTmp
                    If Trim$(varTmp) = "" Then Exit For
                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                    If Not itemX Is Nothing Then
                        spdResult1.Col = intCol:    'spdResult1.Row = OrderCnt
                        If spdResult1.BackColor = &HC6FEFF Then
                            intCnt = intCnt + 1
                            Select Case Trim(itemX.tag)
                                Case "GOT":  intchannel = 1
                                Case "GPT":  intchannel = 3
                                Case "GGTS": intchannel = 10
                                Case "LDLC": intchannel = 21
                                Case "HDLC": intchannel = 20
                                Case "BUN":  intchannel = 16
                                Case "CHOL": intchannel = 5
                                Case "TRIG": intchannel = 8
                                Case "U-A":  intchannel = 17
                                Case "TP":   intchannel = 14
                                Case "ALB":  intchannel = 7
                                Case "TBIL": intchannel = 9
                                Case "ALP":  intchannel = 6
                                Case "LPDC": intchannel = 11
                                Case "CREA": intchannel = 13
                                Case "LDH":  intchannel = 18
                            End Select
                       
                            TestStatus(intchannel) = "1"
        
                        End If
                    End If
                    Set itemX = Nothing
                Next intCol
                
                For xx = 1 To 32
                    ssTemp2 = ssTemp2 & TestStatus(xx)
                Next xx
                
                strOrdList = ""
                strOrdList = STX
                strOrdList = strOrdList & "{"
                strOrdList = strOrdList & "Q" & ";"
                strOrdList = strOrdList & strRack & Space(12 - Len(strRack)) & ";"
                strOrdList = strOrdList & "N" & ";"
                strOrdList = strOrdList & strName & Space(20 - Len(strName)) & ";"
                strOrdList = strOrdList & ";"
                strOrdList = strOrdList & "P" & ";"
                strOrdList = strOrdList & ssTemp2 & ";"
                strOrdList = strOrdList & "}"
                strOrdList = strOrdList & ETX
                strOrdList = strOrdList & vbCrLf
                
                Debug.Print strOrdList
                comEQP.Output = strOrdList
                               
                .Row = intRow
                .Col = 2: .BackColor = vbCyan
                .Col = 3: .BackColor = vbCyan
                .Col = 4: .BackColor = vbCyan
                .Col = 5: .BackColor = vbCyan
                
                ssTemp2 = ""
                strOrdList = ""
                For xx = 1 To 32
                    TestStatus(xx) = "0"
                Next xx
                Sleep (500)
            End If
        Next intRow
    End With
    
    
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
                .Col = 7:       .Text = Trim(sAdd + Val(sNo))
                sAdd = sAdd + 1
            Next sCnt
        End With
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
        .RowHeight(-1) = 12
    End With
    
    blt = True
    
    If cboChk.Text = "" Then
        MsgBox "검진유형을 선택하세요.", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
        Set AdoRs_ORACLE = New ADODB.Recordset
       
    '-- WorkList조회
    Set mAdoRs = f_subSet_WorkList(mskOrdDate.Text, mskOrdDate1.Text)
    
    If RecordChk = False Then
        MsgBox mskOrdDate1.Text & "일의 검사 대상자가 없습니다.", vbOKOnly + vbInformation, Me.Caption
    Else
        strBarno = ""
        mAdoRs.MoveFirst
        With spdWorklist
            'If cboChk.ListIndex = 0 Then
                For intCnt = 0 To mAdoRs.RecordCount - 1
                    If strBarno <> mAdoRs.Fields("SPCNO") & "" Then
                        'optBar.Value = True
                        pGrid_Point = SeqSearch(spdWorklist, mAdoRs.Fields("SPCNO") & "", 3)
            
                        If pGrid_Point = 0 Then
                            pGrid_Point = SeqNullSearch(spdWorklist, mAdoRs.Fields("SPCNO") & "", 3)
                            If pGrid_Point = 0 Then .maxrows = .maxrows + 1: pGrid_Point = .maxrows
                        End If
                            
                        .SetText 1, pGrid_Point, "0"
                        .SetText 2, pGrid_Point, mAdoRs("ORDDT") & ""
                        .SetText 3, pGrid_Point, mAdoRs.Fields("SPCNO") & ""
                        .SetText 4, pGrid_Point, mAdoRs.Fields("PTNM") & ""
                        .SetText 5, pGrid_Point, mAdoRs.Fields("PTID") & ""
                        blt = False
                    End If
                    
                    strEqpCd = f_funGet_CODE(Trim(mAdoRs.Fields("TESTCD")) & "")
                    Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                    If Not itemX Is Nothing Then
                        'spdWorklist.SetText 1, pGrid_Point, "1"
                        spdWorklist.Col = itemX.Index + 7
                        spdWorklist.Row = pGrid_Point
                        spdWorklist.BackColor = &HC6FEFF   '&HC6FEFF
                    End If
                    strBarno = mAdoRs.Fields("SPCNO") & ""
                    mAdoRs.MoveNext
                Next

'                If blt = False Then
'                    .Row = pGrid_Point
'                    .Action = ActionDeleteRow
'                    .maxrows = .maxrows - 1
'                End If
            'End If
        End With
    End If
    
    Set AdoRs_SQL = Nothing
    spdWorklist.Row = 1
    spdWorklist.Col = 1
    spdWorklist.Action = ActionActiveCell
    
'    Dim aROW    As Integer, aCOL   As Integer
'    Dim varChk  As Variant, varBar As Variant, varNum As Variant
'    Dim iRow    As Integer, iCnt   As Integer
'    Dim strRack_tmp As String
'
'    With spdWorklist
'        iCnt = 1
'        .GetText 1, 1, varChk
'        .GetText 2, 1, varBar
'        varNum = 0
'        If Trim(varChk) = "1" And Trim(varBar) <> "" Then
'            For iRow = 1 To .maxrows
'                .SetText 6, iRow, varNum
'                .SetText 7, iRow, ((iCnt Mod 101) + 1) - 1
'                iCnt = iCnt + 1
'                If (iCnt Mod 101) = 1 Then varNum = varNum + 1
'            Next
'        End If
'    End With
'
'    optSeq.Value = True
Exit Sub

ErrorTrap:
    Set AdoRs_SQL = Nothing
    Set AdoRs_ORACLE = Nothing
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
    
    With spdWorklist
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdResult1
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
        .RowHeight(-1) = 12
    End With

    With spdResult2
        .maxrows = 1
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

            .GetText 2, intRow, varTmp:         strDate = Trim$(varTmp)
            .GetText 3, intRow, varTmp:         strBarno = Trim$(varTmp)
            .GetText 4, intRow, varTmp:   pName = Trim$(varTmp)
            .GetText 5, intRow, varTmp:   pNo = Trim$(varTmp)

            .GetText 1, intRow, varTmp

            If strBarno = "" Then Exit For

            intCnt = 0: Erase strOrdcd: Erase strRstval
            If Trim$(varTmp) = "1" Then
                For intCol = 8 To .MaxCols
                    .GetText intCol, intRow, varTmp
                    If Trim$(varTmp) <> "" Then
                        .GetText intCol, 0, varTmp
                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                        If Not itemX Is Nothing Then
                            .GetText intCol, intRow, varTmp
                            strTestcd = itemX.ListSubItems(1)
                            'intPos = InStr(strTestcd, ",")

                            blnFlag = False
                            Set mAdoRs = f_subSet_WorkList_Barcode(strBarno)
                            Do Until mAdoRs.EOF
                                If mAdoRs("TESTCD") = Trim(strTestcd) Then
                                    blnFlag = True
                                    Exit Do
                                End If
                                mAdoRs.MoveNext
                            Loop
                            'strTestcd = Mid$(strTestcd, intPos + 1)
                            'intPos = InStr(strTestcd, ",")

                            'AdoCn_ORACLE.BeginTrans

                            'sqlDoc = "insert into lab_result" & _
                                     "   (RESULTNO, LCODE, LSEQ, LNAME, LRESULT, UNIT, REFV,LTYPE,RESULT_DATE, REPORTER, PT_SEQ) " & _
                                     " values('" & mAdoRs("resultno") & "', '" & strTestcd & "'," & _
                                     "       '1', '" & Trim(itemX.ListSubItems(2)) & "'," & _
                                     "       '" & Trim(varTmp) & "', '" & Trim(itemX.ListSubItems(9)) & "'," & _
                                     "       '" & Trim(itemX.ListSubItems(8)) & "','0',sysdate," & _
                                     "       '70001','" & mAdoRs("pt_seq") & "')"
'                            AdoCn_ORACLE.Execute sqlDoc

'                            sqlDoc = ""
'                            sqlDoc = sqlDoc & "update ipd_order_date set req_result2 = '*'"
'                            sqlDoc = sqlDoc & " where patient_no = '" & mAdoRs("patient_no") & "'"
'                            sqlDoc = sqlDoc & "   and order_date = to_date('" & strDate & "','yyyy-mm-dd')"
'                            AdoCn_ORACLE.Execute sqlDoc
'
'                            sqlDoc = ""
'                            sqlDoc = sqlDoc & "update lab_order set r_flag='1'"
'                            sqlDoc = sqlDoc & " where resultno = '" & mAdoRs("resultno") & "'"
'                            sqlDoc = sqlDoc & "   and patient_no = '" & mAdoRs("patient_no") & "'"
'                            sqlDoc = sqlDoc & "   and lcode = '" & mAdoRs("lcode") & "'"
'                            sqlDoc = sqlDoc & "   and pt_seq = '" & mAdoRs("pt_seq") & "'"
'
'                            AdoCn_ORACLE.Execute sqlDoc
'
'                            AdoCn_ORACLE.CommitTrans
                            If blnFlag = True Then
                                         sqlDoc = "INSERT INTO TC206 "
                                sqlDoc = sqlDoc & "(SpcNo, TestCd, TransDt, TransTm, RstVal, EquipCd, SpcDiv, ErrDes) "
                                sqlDoc = sqlDoc & " Values  ("
                                sqlDoc = sqlDoc & "'" & strBarno & "',"
                                sqlDoc = sqlDoc & "'" & strTestcd & "',"
                                sqlDoc = sqlDoc & "'" & Format$(Now, "YYYYMMDD") & "',"
                                sqlDoc = sqlDoc & "'" & Format$(Now, "HHMMSS") & "',"
                                sqlDoc = sqlDoc & "'" & CStr(varTmp) & "',"
                                sqlDoc = sqlDoc & "'H01',"
                                sqlDoc = sqlDoc & "'',"
                                sqlDoc = sqlDoc & "'')"
                                'Barcode,ExamCode,Date,Tiem,Result,EquipCode(병원내 장비지정 코드),'','')"
    
                                AdoCn_SQL.Execute sqlDoc
    
                                lblStatus.Caption = "저장 성공!!"
                            End If
                            Set adoRS = Nothing:    mAdoRs.Close
                        End If

                        Set itemX = Nothing
                    End If
                Next
                spdResult1.Row = intRow
                spdResult1.Col = 2
                spdResult1.BackColor = vbCyan
                spdResult1.Col = 3
                spdResult1.BackColor = vbCyan
                spdResult1.Col = 4
                spdResult1.BackColor = vbCyan
                spdResult1.Col = 1: spdResult1.Value = 0

                If strErrMsg = "" Then
                    sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
                             " where SPCNO   = '" & strBarno & "'" & _
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

    Exit Sub
ErrorRoutine:
    Set itemX = Nothing

    Me.MousePointer = 0
    Call ErrMsgProc(CallForm)
End Sub

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
'    spdResult2.maxrows = adoRS.RecordCount
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
                    intCol = itemX.Index + 7
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
    
    If Index = 2 Or Index = 3 Then
        With spdResult1
            For intRow = 1 To .maxrows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 2, "1", "")
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

Private Function f_subSet_TestList(ByVal strDate As String, ByVal strSeq As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
        Set AdoRs_SQL = New ADODB.Recordset
               
                 sqlDoc = "select EDPSCODE from GUMJIN_INTERFACE"
        sqlDoc = sqlDoc & " where PER_GUMJIN_DATE = '" & strDate & "'"
        sqlDoc = sqlDoc & "   and PER_GUM_NUM = '" & strSeq & "'"
        sqlDoc = sqlDoc & "   and EDPSCODE in ('0208','0226','0225','0227','0207','0206','0205','0221','0222','0223','0224','0209')"
        sqlDoc = sqlDoc & "   and RESULT=''"
        
        Set AdoRs_SQL = New ADODB.Recordset
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
        
        If AdoRs_SQL.RecordCount = 0 Then
            Set f_subSet_TestList = Nothing
        Else
            Set f_subSet_TestList = AdoRs_SQL
        End If
    
        Set AdoRs_SQL = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
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
    Dim introw1 As Integer, intRow2 As Integer
    Dim intIdx  As Integer
    Dim Rev     As Long
    Dim Test_Cd() As String, strPid()   As String, strPnm() As String
    Dim itemX As ListItem
    Dim blnFlag As Boolean
    Dim strBarno    As String, strSPid  As String, strSPnm   As String
    Dim strWDate As String
    Dim strEqpCd    As String

    blnFlag = False
    With spdWorklist
        For introw1 = 1 To .maxrows
            .GetText 1, introw1, varTmp
            If Trim$(varTmp) = "1" Then
                .GetText 2, introw1, varTmp:    strWDate = Trim$(varTmp)
                .GetText 3, introw1, varTmp:    strBarno = Trim$(varTmp)
                .GetText 4, introw1, varTmp:    strSPnm = Trim$(varTmp)
                .GetText 5, introw1, varTmp:    strSPid = Trim$(varTmp)
                                
                intRow2 = f_funGet_SpreadRow(spdResult1, 3, strBarno)
                If intRow2 < 1 Then
                    intRow2 = f_funGet_SpreadRow(spdResult1, 2, "")
                    If intRow2 < 1 Then
                        spdResult1.maxrows = spdResult1.maxrows + 1
                        spdResult1.RowHeight(spdResult1.maxrows) = 12
                        intRow2 = spdResult1.maxrows
                    End If

                    blnFlag = False
                    
                    Set mAdoRs = f_subSet_WorkList_Barcode(strBarno, Mid(strSPnm, 1, 2))

                    If Len(strBarno) > 0 Then
                        Do Until mAdoRs.EOF
                            strEqpCd = Trim(mAdoRs("TESTCD") & "")
'                            If cboChk.ListIndex = 0 Then
'                                strEqpCd = f_funGet_CODE(Trim(mAdoRs("edpscode")))
'                            Else
'                                strEqpCd = f_funGet_CODE(Trim(mAdoRs("map2seqno")))
'                            End If
                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwSubItem, , lvwWhole)
                            If Not itemX Is Nothing Then
                                blnFlag = True
                                spdResult1.Row = intRow2
                                spdResult1.Col = itemX.Index + 7
                                spdResult1.BackColor = &HC6FEFF '&H80C0FF
                                DoEvents
                            End If
                            mAdoRs.MoveNext
                        Loop
                    End If
                    If blnFlag = True Then
                        spdResult1.SetText 2, intRow2, strWDate
                        spdResult1.SetText 4, intRow2, strSPnm
                        spdResult1.SetText 3, intRow2, strBarno
                        spdResult1.SetText 5, intRow2, strSPid
                    Else
                        spdResult1.maxrows = spdResult1.maxrows - 1
                    End If
                End If
                spdResult1.SetText 1, intRow2, "0"
'                spdResult1.maxrows = intRow2

                .SetText 1, introw1, ""
            End If
        Next
    End With
                
End Sub

Private Sub comEQP_OnComm()
    
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim strDta      As String
    Dim Arr()       As Byte
    Dim sStxCheck As Integer, sEnqCheck As Integer, sEtxCheck As Integer
    Dim sLfCheck As Integer, sCrcheck As Integer, ii As Integer
    Dim MHead As String, Pinfo As String, OutPutData As String, com_sTemp As String
    Dim strRec  As String, strBuff  As String
    
    Dim intIdx     As Integer, intIdx1    As Integer, intIdx2     As Integer
    Dim strTmp1     As String, strTmp2      As String
    Dim intPos1     As Integer, intPos2     As Integer
    Dim intCnt       As Integer
   
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
    
    Dim strRec  As String, strBuff  As String
    Dim strTmp  As String, intIdx   As Integer
    Dim intPos1 As Integer, intPos2 As Integer
    
    Dim strdata()   As String, intCnt   As Integer
    
    Static OrgMsg As String
    strRec = RecData ' StrConv(RecData, vbUnicode)
    
    Print #1, strRec;
    
    strTmp = strRec
    Call COM_INPUT(strTmp)
    
    For intIdx = 1 To Len(strRec)
        strBuff = Mid$(strRec, intIdx, 1)
        Select Case Asc(strBuff)
            Case 2  '-- STX
                    f_strBuffer = strBuff
                    
            Case 3  '-- ETX
                    f_strBuffer = f_strBuffer + strBuff
                    intCnt = 0
                    strTmp = f_strBuffer
                    If InStr(strTmp, "D1U") > 0 Then
                        Call psDataDefine(strTmp, fChannel(), spdResult1, IIf(optSeq.Value, "0", "1"))
                    End If
'                    comEQP.Output = ACK
            Case Else
                    f_strBuffer = f_strBuffer + strBuff
        End Select
     Next
End Sub


Private Sub psDataDefine(ByVal brbarcd As String, ByRef brChannel() As String, ByVal brspread As Object, ByVal brOst As String) ' ByRef brItemdeci() As String)

Dim sTemp      As String
Dim Channel_No As Integer       ' 검사항목 번호 : Channel No
Dim pGrid_Point As Integer
Dim pDoCount   As Integer
Dim Loop_Count As Integer
Dim FunStr As String
Dim Max_Arary_Cnt As Integer    ' 검사 항목수
Dim sAdd As Integer, sPosition As Integer
Dim itemX As ListItem
Dim strRstval As String, strRefVal(1 To 19)  As String, strOrdLst() As String
Dim sqlDoc  As String
Dim intCol As Integer
Dim Gnum   As String
Dim ii As Integer
Dim Rev As Long
Dim tmpTstCd As String
Dim tmpMXD As Variant
 
Dim strTime As String, sqlRet   As Integer
Dim pQCDoCount   As Integer
Dim chkGbn As Variant
Dim varTmp
Dim sSeq, strTmp, strBarno, strDate, strDate1 As String
Dim sCol As Integer
Dim strMidval As String
Dim strEqpCd As String

    'On Error GoTo ErrRoutine
    
    CallForm = "frmInterface - Privete sub psDataDefine()"
    sTemp = brbarcd
    '------------------------------<<< fKX21() 배열 Clear 한다.         >>>----------
    'For Loop_Count = 1 To 100: fKX21(Loop_Count) = "": Next Loop_Count
    '------------------------------<<< fKX21() 배열에 구분하여 넣는다.  >>>----------
    Erase fKX21
    
    If Mid$(sTemp, 2, 3) = "D1U" Then
        strTime = Format$(Now, "MMSS")  '-- 같은 Row이면 시간을 동일하게 처리
        
        fKX21(0) = Mid(sTemp, 11, 13)                   ' 항목 1 "0000000073" 일련번호
        sSeq = Val(fKX21(0))
        sPosition = 31
         fKX21(1) = Mid(sTemp, sPosition, 4)      '-- WBC
         fKX21(2) = Mid(sTemp, sPosition + 5, 4)  '-- RBC
         fKX21(3) = Mid(sTemp, sPosition + 10, 4) '-- HGB
         fKX21(4) = Mid(sTemp, sPosition + 15, 4) '-- HCT
         fKX21(5) = Mid(sTemp, sPosition + 20, 4) '-- MCV
         fKX21(6) = Mid(sTemp, sPosition + 25, 4) '-- MCH
         fKX21(7) = Mid(sTemp, sPosition + 30, 4) '-- MCHC
         fKX21(8) = Mid(sTemp, sPosition + 35, 4) '-- PLT
         
         fKX21(9) = Mid(sTemp, sPosition + 40, 4) '-- LYM%
        fKX21(10) = Mid(sTemp, sPosition + 45, 4) '-- MXD%
        fKX21(11) = Mid(sTemp, sPosition + 50, 4) '-- NEUT%
        fKX21(12) = Mid(sTemp, sPosition + 55, 4) '-- LYM#
        fKX21(13) = Mid(sTemp, sPosition + 60, 4) '-- MXD#
        fKX21(14) = Mid(sTemp, sPosition + 65, 4) '-- NEUT#
        fKX21(15) = Mid(sTemp, sPosition + 70, 4) '-- RDW
        fKX21(16) = Mid(sTemp, sPosition + 75, 4) '-- PDW
        fKX21(17) = Mid(sTemp, sPosition + 80, 4) '-- MPV
        fKX21(18) = Mid(sTemp, sPosition + 85, 4) '-- P-LCR

        '-------------------------------------------<<< 해당검사결과와 해당환자를 찿는다.       >>>----------

        Max_Arary_Cnt = brspread.MaxCols - 6  ' 앞에서부터 2까지는 환자 정보 이기때문에.... -2를 한다.
                                               ' 해당 배열은  brItem(),brChannel() 이다.
        pGrid_Point = 0
        
        With brspread
            pGrid_Point = SeqSearch(spdResult1, sSeq, sCol)
            
            .GetText 2, pGrid_Point, varTmp:   strDate = Trim$(varTmp)
            .GetText 3, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)
            .GetText 4, pGrid_Point, varTmp:   pName = Trim$(varTmp)
            .GetText 5, pGrid_Point, varTmp:   pNo = Trim$(varTmp)
            chkGbn = Split(pName, ":")
            
            If pGrid_Point > 0 Then                ' 해당 대상자를 찿으면 ....
                Set mAdoRs = f_subSet_WorkList_Barcode(strBarno)
                pQCDoCount = 0
                For pDoCount = 1 To Max_Arary_Cnt - 1 '-------------------------------<<<<<<<<<,  세부검사항목을 찿는다.  >>>>>>>---------
                    .Row = pGrid_Point
                    .Col = pDoCount + 7
                    Channel_No = Val(brChannel(pDoCount))              '  Channel이 숫자이기 때문에 숫자로 치환한다.
                    
                    If Len(fKX21(Channel_No)) > 0 Or Channel_No = 19 Or Channel_No = 20 Or Channel_No = 21 Then
                        If fKX21Size(Channel_No, 1) = fKX21Size(Channel_No, 0) Then
                            FunStr = Trim(Val(fKX21(Channel_No)))
                        Else
                            FunStr = Trim(Val(Mid$(fKX21(Channel_No), 1, fKX21Size(Channel_No, 1)))) + "." + _
                                           Mid$(fKX21(Channel_No), fKX21Size(Channel_No, 1) + 1)
                        End If
                        
                        If FunStr <> "" Then
                            Select Case Channel_No
                                'Case 1: FunStr = FunStr * 1000 'WBC
                                'Case 2: FunStr = FunStr * 100 'WBC
                                Case 8: strMidval = FunStr
                                        If Right(FunStr, 1) = "." Then
                                            FunStr = Replace(FunStr, ".", "")
                                        End If
                                Case 9: If FunStr = "0.0" Then FunStr = ""
                                Case 10: If FunStr = "0.0" Then FunStr = ""
                                Case 11: If FunStr = "0.0" Then FunStr = ""
                            End Select
                            
                            .Text = FunStr
                            
                            Set itemX = lvwCuData.FindItem(Channel_No, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then '-- itemX : 검사코드
                                strEqpCd = ""
                                Do Until mAdoRs.EOF
                                    If Mid(itemX.Text, InStr(itemX.Text, ",") + 1) = Trim(mAdoRs.Fields("TESTCD")) Then
                                        strEqpCd = Trim(mAdoRs.Fields("TESTCD"))
                                        Exit Do
                                    End If
                                    mAdoRs.MoveNext
                                Loop
                                mAdoRs.MoveFirst
                                        
                                If strEqpCd <> "" Then
                                    strDate1 = Format$(Now, "YYYYMMDD"):     strTime = Format$(Now, "MMSS")
                                    sqlDoc = "Update INTERFACE003" & _
                                             "   set RSTVAL  = '" & FunStr & "', REFVAL = ''" & _
                                             " where SPCNO   = '" & CStr(strBarno) & "'" & _
                                             "   and EQPNUM  = '" & Channel_No & "'" & _
                                             "   and TRANSDT = '" & Format$(Now, "YYYYMMDD") & "'" & _
                                             "   and TRANSTM = '" & strTime & "'"
                                    AdoCn_Jet.Execute sqlDoc, sqlRet
                                    
                                    If sqlRet = 0 Then
                                        sqlDoc = "insert into INTERFACE003(" & _
                                                 "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN,NAME,PNO)" & _
                                                 "    values( '" & CStr(strBarno) & "', '" & strEqpCd & "', '" & Channel_No & "'," & _
                                                 "            '" & Format$(Now, "YYYYMMDD") & "', '" & strTime & "'," & _
                                                 "            '" & FunStr & "', ''," & _
                                                 "            '" & INS_CODE & "', '','" & pName & "','" & pNo & "')"
                                        
                                        AdoCn_Jet.Execute sqlDoc
                                    End If
                                    
                                    If chkAuto.Value = "1" Then
'                                        If Mid(pName, 1, 2) = "검진" Then
'                                            sqlDoc = "Update GUMJIN_INTERFACE" & _
'                                                     "   set RESULT = '" & strRstval & "'," & _
'                                                     "       ACT_RETURN_DATE = '" & strDate1 & "'" & _
'                                                     " where PER_GUMJIN_DATE = '" & strDate & "'" & _
'                                                     "   and PER_GUM_NUM = " & pNo & "" & _
'                                                     "   and EDPSCODE = '" & strEqpCd & "'"
'                                        Else
'                                            sqlDoc = "Update MEDICOM..jun370_resulttb" _
'                                                    & "   Set Result = '" & FunStr & "', status='1'" _
'                                                    & " Where WaitSeqNo = '" & pNo & "'" _
'                                                    & "   and map2seqno = '" & strEqpCd & "'"
'                                        End If
                                        
                                    
                                                 sqlDoc = "INSERT INTO TC206 "
                                        sqlDoc = sqlDoc & "(SpcNo, TestCd, TransDt, TransTm, RstVal, EquipCd, SpcDiv, ErrDes) "
                                        sqlDoc = sqlDoc & " Values  ("
                                        sqlDoc = sqlDoc & "'" & strBarno & "',"
                                        sqlDoc = sqlDoc & "'" & strEqpCd & "',"
                                        sqlDoc = sqlDoc & "'" & Format$(Now, "YYYYMMDD") & "',"
                                        sqlDoc = sqlDoc & "'" & Format$(Now, "HHMMSS") & "',"
                                        sqlDoc = sqlDoc & "'" & FunStr & "',"
                                        sqlDoc = sqlDoc & "'H01',"
                                        sqlDoc = sqlDoc & "'',"
                                        sqlDoc = sqlDoc & "'')"
                                        'Barcode,ExamCode,Date,Tiem,Result,EquipCode(병원내 장비지정 코드),'','')"

                                        AdoCn_SQL.Execute sqlDoc
                                    End If
                                End If
                            End If
                            

                            Set itemX = Nothing
                            
                        End If
                    End If
                Next pDoCount
                spdResult1.Row = pGrid_Point
                spdResult1.Col = 2
                spdResult1.BackColor = vbCyan
                spdResult1.Col = 3
                spdResult1.BackColor = vbCyan
                spdResult1.Col = 4
                spdResult1.BackColor = vbCyan
                spdResult1.Col = 5
                spdResult1.BackColor = vbCyan
                spdResult1.Col = 6
                spdResult1.BackColor = vbCyan
                spdResult1.Col = 7
                spdResult1.BackColor = vbCyan
                spdResult1.Col = 1: spdResult1.Value = 1
                
                '-----------------------------------------------------------------------
            End If
        End With
    End If
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)

End Sub

Private Sub ReceiveTheData(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object) ', ByVal brOst As String) ' ByRef brItemdeci() As String)
    
    
    Dim sTemp      As String
    Dim Channel_No As String       ' 검사항목 번호 : Channel No
    Dim pGrid_Point As Integer
    Dim pDoCount   As Integer
    Dim Loop_Count As Integer
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
    Dim sSeq, strTmp, varTmp, strBarno, strDate, strDate1, strTime As String
    Dim sCol As Integer
    Dim sDecnt As Integer
    Dim Float_rate1 As String
    Dim Float_rate2 As String
    Dim Float_rate  As String
    Dim intRow, intIdx As Integer
    Dim chrChk As Boolean
    Dim seqChk As Variant
    Dim chkGbn As Variant

    On Error Resume Next
       
    CallForm = "frmInterface - Privete sub psDataDefine()"
    tmpMXD = "0"
    sTemp = strdata
    Erase fAVL9180
    Erase fAVL9180_1
    Erase fAVL9180_2
    Erase fAVL9180_3
    
    fAVL9180 = Split(sTemp, vbLf)
    
    pGrid_Point = 0
    strTmp = ""
    If Len(fAVL9180(9)) > 0 Then
        intRow = 0
        With spdResult1
            fAVL9180_1 = Split(fAVL9180(9), ".")
            sSeq = Trim(fAVL9180_1(1))
            sCol = 8
'            pGrid_Point = SeqSearch(spdResult1, sSeq, sCol)
            pGrid_Point = SeqNullSearch(spdResult1, sSeq, sCol)
    
            .GetText 2, pGrid_Point, varTmp:   strDate = Trim$(varTmp)
            .GetText 3, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)
            .GetText 4, pGrid_Point, varTmp:   pName = Trim$(varTmp)
            .GetText 5, pGrid_Point, varTmp:   pNo = Trim$(varTmp)
            chkGbn = Split(pName, ":")

            .GetText 2, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)

            If pGrid_Point > 0 Then
                For intCol = 8 To .MaxCols
                    strRstval = ""
                    .GetText intCol, 0, varTmp
                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                    If Not itemX Is Nothing Then
                        For intIdx = 1 To .MaxCols
                            If Len(fAVL9180(9)) > 0 Then
                                fAVL9180_2 = Split(Trim(fAVL9180(intIdx + 10)), "=")
'                                fAVL9180_2 = Split(Trim(fAVL9180(intIdx + 11)), "=")
                                Channel_No = Trim(fAVL9180_2(0))
                                If UCase(Channel_No) = UCase(itemX.tag) Then
                                    fAVL9180_3 = Split(Trim(fAVL9180_2(1)), " ")
'                                    strRstval = fAVL9180_3(0)
                                    If UCase(Channel_No) = "CL" Then
                                        strRstval = ""
                                    Else
                                        'strRstval = Mid(fAVL9180_3(0), 1, 3)
                                        strRstval = Trim(Mid(fAVL9180_2(1), 3, 3))
                                    End If
                                     strDate1 = Format$(Now, "YYYYMMDD"):     strTime = Format$(Now, "MMSS")
                                    .SetText intCol, pGrid_Point, strRstval
                                    .Col = intCol:  .Row = pGrid_Point
                                                    .ForeColor = IIf(Trim$(strRefVal) <> "", vbRed, vbBlack)

                                    sqlDoc = "Update INTERFACE003" & _
                                             "   set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRefVal & "'" & _
                                             " where SPCNO   = '" & strBarno & "'" & _
                                             "   and EQPNUM  = '" & itemX.tag & "'" & _
                                             "   and TRANSDT = '" & strDate1 & "'" & _
                                             "   and TRANSTM = '" & strTime & "'"
                                    AdoCn_Jet.Execute sqlDoc

                                    If cboChk.ListIndex = 0 Then
                                        sqlDoc = "insert into INTERFACE003(" & _
                                                 "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                                 "    values( '" & strBarno & "', '" & itemX.Text & "', '" & itemX.tag & "'," & _
                                                 "            '" & strDate1 & "', '" & strTime & "'," & _
                                                 "            '" & strRstval & "', '" & strRefVal & "'," & _
                                                 "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
                                    Else
                                        sqlDoc = "insert into INTERFACE003(" & _
                                                 "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                                 "    values( '" & strBarno & "', '" & Mid(itemX.Text, 6) & "', '" & itemX.tag & "'," & _
                                                 "            '" & strDate1 & "', '" & strTime & "'," & _
                                                 "            '" & strRstval & "', '" & strRefVal & "'," & _
                                                 "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
                                    End If

                                    AdoCn_Jet.Execute sqlDoc

                                    If chkAuto.Value = "1" Then
                                        If chkGbn(0) = "검진" Then
                                            sqlDoc = "Update GUMJIN_INTERFACE" & _
                                                     "   set RESULT = '" & strRstval & "'," & _
                                                     "       ACT_RETURN_DATE = '" & strDate1 & "'" & _
                                                     " where PER_GUMJIN_DATE = '" & strDate & "'" & _
                                                     "   and PER_GUM_NUM = " & pNo & "" & _
                                                     "   and EDPSCODE = '" & Mid(itemX.Text, 1, 4) & "'"
                                        Else
                                            sqlDoc = "Update MEDICOM..jun370_resulttb" _
                                                    & "   Set Result = '" & strRstval & "', status='1'" _
                                                    & " Where WaitSeqNo = '" & pNo & "'" _
                                                    & "   and map2seqno = '" & itemX.Text & "'"
                                        End If
                                        AdoCn_SQL.Execute sqlDoc
                                    End If

                                    spdResult1.Row = pGrid_Point
                                    spdResult1.Col = 2
                                    spdResult1.BackColor = vbCyan
                                    spdResult1.Col = 3
                                    spdResult1.BackColor = vbCyan
                                    spdResult1.Col = 4
                                    spdResult1.BackColor = vbCyan
                                    spdResult1.Col = 5
                                    spdResult1.BackColor = vbCyan
                                    spdResult1.Col = 6
                                    spdResult1.BackColor = vbCyan
                                    spdResult1.Col = 7
                                    spdResult1.BackColor = vbCyan
                                    spdResult1.Col = 1: spdResult1.Value = 0
                                    Exit For

                                End If
                            End If
                        Next intIdx
                    End If
                    Set itemX = Nothing
                Next
            End If
        End With
    End If
    Exit Sub

ErrRoutine:

    Call ErrMsgProc(CallForm)

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
                If Val(.Text) = brSeq Then
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

   
    strDta = "D1U0509090000000000295000000S006500388001150035200907002960032700099001450003800817000090000200054004700016300115003730"


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
    
    Call f_subSet_ComCharacter  ' 통신문자
    Call f_subGet_Setting       ' 통신설정
    
    Call cmdRun                 ' 실행
    
    mskRstDate.Text = Format$(Now, "YYYYMMDD")
    mskOrdDate.Text = Format$(Now, "YYYYMMDD")
    mskOrdDate1.Text = Format$(Now, "YYYYMMDD")
    Open App.Path + "\" + "AVL9180.log" For Append As #1

    Print #1, Chr(13) + Chr(10);
   
    f_strJOB_FLAG = "1":    f_intSampleNo = 0
    tabWork.Tab = 0
    Or_Seq = 1
    intRow = 0
    chkEnq = 0
    cboChk.ListIndex = 1
    
    fKX21Size(1, 0) = 5:      fKX21Size(1, 1) = 3       ' WBC
    fKX21Size(2, 0) = 5:      fKX21Size(2, 1) = 2       ' RBC
    fKX21Size(3, 0) = 5:      fKX21Size(3, 1) = 3       ' HGB
    fKX21Size(4, 0) = 5:      fKX21Size(4, 1) = 3       ' HCT
    fKX21Size(5, 0) = 5:      fKX21Size(5, 1) = 3       ' MCV
    fKX21Size(6, 0) = 5:      fKX21Size(6, 1) = 3       ' MCH
    fKX21Size(7, 0) = 5:      fKX21Size(7, 1) = 3       ' MCHC
    fKX21Size(8, 0) = 5:      fKX21Size(8, 1) = 4      ' PLT
    fKX21Size(9, 0) = 5:      fKX21Size(9, 1) = 3      ' LYMP%
    fKX21Size(10, 0) = 5:     fKX21Size(10, 1) = 3      ' MXD%
    fKX21Size(11, 0) = 5:     fKX21Size(11, 1) = 3      ' NEUT%
    fKX21Size(12, 0) = 5:     fKX21Size(12, 1) = 3      ' LYMPH#
    fKX21Size(13, 0) = 5:     fKX21Size(13, 1) = 3      ' MONO#
    fKX21Size(14, 0) = 5:     fKX21Size(14, 1) = 3      ' NEUT#
    fKX21Size(15, 0) = 5:     fKX21Size(15, 1) = 3      ' RDW
    fKX21Size(16, 0) = 5:     fKX21Size(16, 1) = 3      ' PDW
    fKX21Size(17, 0) = 5:     fKX21Size(17, 1) = 3      ' MPV
    fKX21Size(18, 0) = 5:     fKX21Size(18, 1) = 3      ' P-LCR
    
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
    Dim introw1 As Integer
    Dim intRow2 As Integer
    Dim iCnt    As Integer
    
    If Row = 0 Then Exit Sub
    
    intCol1 = 8
    intCol2 = 2
    introw1 = 1
    
    With spdResult1
        For iCnt = intCol1 To .MaxCols
            .Row = Row
            .Col = intCol1
            
            spdRstview.Row = introw1
            spdRstview.Col = intCol2
            spdRstview.Text = .Text
            
            introw1 = introw1 + 1
            intCol1 = intCol1 + 1
            
            If introw1 > spdRstview.maxrows Then
                introw1 = 1
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

'    If Col < 3 Then Exit Sub
    
    Dim varTmp  As Variant
    
    With spdWorklist
        If Col = 1 Then
            .GetText 2, Row, varTmp
            If Trim$(varTmp) = "" Then Exit Sub
            .GetText 1, Row, varTmp
            
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


Private Sub spdWorklist_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    With spdWorklist
        .Row = Row
        .Action = ActionDeleteRow
        .maxrows = .maxrows - 1
    End With
    
End Sub

Private Sub tabWork_Click(PreviousTab As Integer)
    cboRstgbn(1).ListIndex = 2
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
    If ScaleWidth < 60 Then Exit Sub
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For I = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(I).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - I)) + (70 * (cmdAction.UBound - I)) + 100), _
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

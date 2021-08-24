VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "GTCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmComm 
   Caption         =   "Interface"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   11985
   WindowState     =   2  '최대화
   Begin VB.Timer tmrReceive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4695
      Top             =   6510
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5175
      Top             =   6510
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   2955
      Top             =   6510
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
      Left            =   3810
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   4800
      SThreshold      =   1
   End
   Begin MSComctlLib.ImageList imlStatus 
      Left            =   5640
      Top             =   6510
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
      Left            =   30
      TabIndex        =   1
      Top             =   6495
      Width           =   11940
      Begin VB.Timer Timer1 
         Left            =   2580
         Top             =   90
      End
      Begin VB.Timer tmrInitial 
         Left            =   2160
         Top             =   120
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   0
         Left            =   6375
         TabIndex        =   2
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Run"
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
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   1
         Left            =   7740
         TabIndex        =   3
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Stop"
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
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   2
         Left            =   9120
         TabIndex        =   4
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
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
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   3
         Left            =   10485
         TabIndex        =   5
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Close"
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
         TabIndex        =   10
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
         TabIndex        =   9
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
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmComm.frx":3F0A
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
         Left            =   10140
         TabIndex        =   8
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send : "
         Height          =   180
         Left            =   9105
         TabIndex        =   7
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port : "
         Height          =   180
         Left            =   8040
         TabIndex        =   6
         Top             =   285
         Width           =   510
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   11010
         Picture         =   "frmComm.frx":518C
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   9780
         Picture         =   "frmComm.frx":5716
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   8640
         Picture         =   "frmComm.frx":5CA0
         Top             =   255
         Width           =   240
      End
   End
   Begin TabDlg.SSTab tabWork 
      Height          =   5850
      Left            =   30
      TabIndex        =   11
      Top             =   630
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   10319
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " WorkList"
      TabPicture(0)   =   "frmComm.frx":622A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "pnlCom"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdInitial"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "spdWorkList"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdWorkList"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdStartNo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdWordQuery"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "pnlCom2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdAppend(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "mskOrdDate"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboRstgbn(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "optBar"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "optSeq"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkAuto"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboOrder"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "spdResult1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdSel(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdSel(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   " 받은 결과"
      TabPicture(1)   =   "frmComm.frx":6246
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(2)=   "mskRstDate1"
      Tab(1).Control(3)=   "lvwCuData"
      Tab(1).Control(4)=   "cmdSel(2)"
      Tab(1).Control(5)=   "cmdSel(3)"
      Tab(1).Control(6)=   "cmdAppend(1)"
      Tab(1).Control(7)=   "mskRstDate"
      Tab(1).Control(8)=   "cboRstgbn(1)"
      Tab(1).Control(9)=   "spdResult2"
      Tab(1).Control(10)=   "cmdRstQuery"
      Tab(1).ControlCount=   11
      Begin HSCotrol.CButton cmdRstQuery 
         Height          =   300
         Left            =   -64380
         TabIndex        =   16
         Top             =   495
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         Caption         =   "조 회"
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
      Begin FPSpreadADO.fpSpread spdResult2 
         Height          =   4875
         Left            =   -74865
         TabIndex        =   57
         Top             =   810
         Width           =   11565
         _Version        =   393216
         _ExtentX        =   20399
         _ExtentY        =   8599
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
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
         SpreadDesigner  =   "frmComm.frx":6262
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   0
         Left            =   135
         TabIndex        =   22
         Top             =   840
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":6781
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   1
         Left            =   435
         TabIndex        =   21
         Top             =   840
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":6BEF
      End
      Begin FPSpreadADO.fpSpread spdResult1 
         Height          =   4920
         Left            =   120
         TabIndex        =   49
         Top             =   840
         Width           =   11565
         _Version        =   393216
         _ExtentX        =   20399
         _ExtentY        =   8678
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   5
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
         SpreadDesigner  =   "frmComm.frx":7071
      End
      Begin VB.ComboBox cboOrder 
         Height          =   300
         ItemData        =   "frmComm.frx":759E
         Left            =   2970
         List            =   "frmComm.frx":75A0
         Style           =   2  '드롭다운 목록
         TabIndex        =   48
         Top             =   525
         Visible         =   0   'False
         Width           =   780
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
         Left            =   8685
         TabIndex        =   46
         Top             =   30
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Test"
         Height          =   285
         Left            =   6390
         TabIndex        =   45
         Top             =   360
         Visible         =   0   'False
         Width           =   960
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
         Left            =   10170
         TabIndex        =   44
         Top             =   0
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
         Left            =   11100
         TabIndex        =   43
         Top             =   0
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cboRstgbn 
         Height          =   300
         Index           =   1
         ItemData        =   "frmComm.frx":75A2
         Left            =   -70080
         List            =   "frmComm.frx":75AF
         Style           =   2  '드롭다운 목록
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.ComboBox cboRstgbn 
         Height          =   300
         Index           =   0
         ItemData        =   "frmComm.frx":75D9
         Left            =   3930
         List            =   "frmComm.frx":75E6
         Style           =   2  '드롭다운 목록
         TabIndex        =   12
         Top             =   15
         Visible         =   0   'False
         Width           =   1500
      End
      Begin MSMask.MaskEdBox mskRstDate 
         Height          =   300
         Left            =   -73695
         TabIndex        =   14
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
      Begin HSCotrol.CButton cmdAppend 
         Height          =   300
         Index           =   1
         Left            =   -64380
         TabIndex        =   15
         Top             =   495
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         Caption         =   "서버등록"
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
      Begin MSMask.MaskEdBox mskOrdDate 
         Height          =   300
         Left            =   1275
         TabIndex        =   17
         Top             =   525
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin HSCotrol.CButton cmdAppend 
         Height          =   300
         Index           =   0
         Left            =   9450
         TabIndex        =   18
         Top             =   510
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         Caption         =   "서버등록"
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
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   3
         Left            =   -74640
         TabIndex        =   41
         Top             =   810
         Visible         =   0   'False
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":7610
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   -74910
         TabIndex        =   42
         Top             =   810
         Visible         =   0   'False
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":7A92
      End
      Begin HSCotrol.UserPanel pnlCom2 
         Height          =   4785
         Left            =   5895
         TabIndex        =   30
         Top             =   1005
         Visible         =   0   'False
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   8440
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
            TabIndex        =   31
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
               TabIndex        =   32
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
               TabIndex        =   33
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
            Begin HSCotrol.CButton cmdCOMInput2 
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
            Begin HSCotrol.CButton cmdCOMLoad 
               Height          =   360
               Left            =   4635
               TabIndex        =   36
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
               TabIndex        =   37
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
               TabIndex        =   38
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
            Left            =   45
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   39
            Top             =   270
            Width           =   5730
         End
      End
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   4920
         Left            =   -67980
         TabIndex        =   40
         Top             =   810
         Visible         =   0   'False
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   8678
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
      Begin HSCotrol.CButton cmdWordQuery 
         Height          =   300
         Left            =   3825
         TabIndex        =   50
         Top             =   495
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         Caption         =   "조 회"
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
      Begin HSCotrol.CButton cmdStartNo 
         Height          =   300
         Left            =   4890
         TabIndex        =   51
         Top             =   495
         Visible         =   0   'False
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "시작번호변경"
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
      Begin HSCotrol.CButton cmdWorkList 
         Height          =   300
         Left            =   135
         TabIndex        =   52
         Top             =   5400
         Visible         =   0   'False
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   529
         Caption         =   "WorkList 작성"
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
      Begin FPSpreadADO.fpSpread spdWorkList 
         Height          =   4560
         Left            =   135
         TabIndex        =   53
         Top             =   810
         Visible         =   0   'False
         Width           =   2490
         _Version        =   393216
         _ExtentX        =   4392
         _ExtentY        =   8043
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
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
         SpreadDesigner  =   "frmComm.frx":7F00
      End
      Begin HSCotrol.CButton cmdInitial 
         Height          =   300
         Left            =   10590
         TabIndex        =   54
         Top             =   510
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         Caption         =   "초기화"
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
      Begin HSCotrol.UserPanel pnlCom 
         Height          =   5355
         Left            =   90
         TabIndex        =   23
         Top             =   990
         Visible         =   0   'False
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   9446
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
            Left            =   45
            TabIndex        =   24
            Top             =   4650
            Width           =   11610
            Begin HSCotrol.CButton cmdCOMSave 
               Height          =   360
               Left            =   10515
               TabIndex        =   25
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
               TabIndex        =   26
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
               TabIndex        =   27
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
               TabIndex        =   28
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
            Height          =   4395
            Left            =   45
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   29
            Top             =   270
            Width           =   11595
         End
      End
      Begin MSMask.MaskEdBox mskRstDate1 
         Height          =   300
         Left            =   -72360
         TabIndex        =   55
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
      Begin VB.Label Label7 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -72540
         TabIndex        =   56
         Top             =   495
         Width           =   150
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "차수"
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
         Left            =   2490
         TabIndex        =   47
         Top             =   585
         Visible         =   0   'False
         Width           =   390
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
         TabIndex        =   20
         Top             =   570
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "검체접수일 :"
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
         TabIndex        =   19
         Top             =   555
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

Private Const ST_ID = " "
Private Const ST_UI = " "

Private mSID            As String * 14
Private mPTID           As String * 14
Private mPTNM           As String * 30
Private mDOB            As String * 10
Private mSEX            As String * 1
Private mCLD            As String * 8
Private mCLT            As String * 4
Private mLCT            As String * 6
Private mDOC            As String * 6

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

Private MSG_STX     As String
Private MSG_ETX     As String
Private MSG_ENQ     As String
Private MSG_EOT     As String
Private MSG_ACK     As String
Private MSG_NAK     As String
Private MSG_CR      As String
Private MSG_LF      As String
Private MSG_CRLF    As String

Dim fAdvia120(100) As String
Dim fAdvia120Cfg(100) As Integer
Dim fAdvia120Size(100, 1) As Integer
Dim fChannel() As String
Dim RecordChk As Boolean
Dim sndMsg As String
Dim stxchk, etxchk As Integer
Dim sOrderLst As String
Dim PatientID As String
Dim sChannel As String
Dim CHR_MT       As String

Private Type TYPE_CD
    strEqpCd    As String
    intCnt      As Integer
    strTestcd(100) As String
End Type
Private f_typCode() As TYPE_CD

Dim pName    As String
Dim pNo      As String


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
                .Tag = Trim(adoRS.Fields("TESTCD") & "")
                .Width = 700
                .Alignment = lvwColumnCenter
            End With
            Set itemH = Nothing
        End With
        
'        With spdWorkList
'            intCol = intCol + 1
'            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1:  .ColWidth(.MaxCols) = 6.5
'
'            .SetText intCol, 0, adoRS.Fields("TESTNM")
'        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub f_subSet_ItemList()

    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strTest As String, intPos   As Integer
    Dim strTmp  As String, intCol   As Integer, intCnt  As Integer
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = ""
    
    intCol = 6
'    With spdResult1
'        .Col = 1:   .Col2 = .MaxCols
'        .Row = 0:   .Row2 = .MaxRows
'        .MaxRows = 15
'        .BlockMode = True
'        .Action = ActionClearText
'        .BlockMode = False
'        .RowHeight(-1) = 12
'    End With
    
    With spdResult2
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .MaxRows = 15
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
            itemX.Tag = Trim(adoRS.Fields("TEST_EQP") & "")
            itemX.Text = Trim(adoRS.Fields("TESTCD") & "")
        Set itemX = Nothing
        
        With spdResult1
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
            .ColWidth(intCol) = 6.5
        End With
        
        With spdResult2
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
            .ColWidth(intCol) = 6.5
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
    f_strOrdList = Mid$(f_strOrdList, 1, Len(f_strOrdList) - 1)
    
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

Private Sub cboOrder_DropDown()
    Call f_subSet_DateOrder
End Sub


Private Sub cmdACK_Click()
    
    Call Com_Output(Chr(1))

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
    
    f_strJOB_FLAG = "1"
    f_intSampleNo = 0
    
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
        .MaxRows = 15
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
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
    
    If Not comEQP.PortOpen Then
        comEQP.PortOpen = True
        comEQP.RTSEnable = True
        comEQP.RThreshold = 1
    End If
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
    
    Dim strOrdLst()     As String, strPid()    As String, strPnm() As String
    
    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
    Dim itemX   As ListItem
    Dim objSpd  As fpSpread
    'Dim  As Integer
    Dim flgSave As Boolean
    Dim sqlRet As Integer
    
    CallForm = "frmComm - Private Sub cmdAppend_Click()"
    
On Error GoTo ErrorRoutine
    
    Me.MousePointer = 11
    
    If Index = 0 Then
        Set objSpd = spdResult1
    Else
        Set objSpd = spdResult2
    End If
    
    With objSpd
        For intRow = 1 To .MaxRows
            .GetText 2, intRow, varTmp:         strBarno = Trim$(varTmp)
'            .GetText .MaxCols, intRow, varTmp:  strTime = Trim$(varTmp)
'            strTime = Format$(Now, "MMSS")
            .GetText 1, intRow, varTmp
            
            If strBarno = "" Then Exit For
            
            intCnt = 0: Erase strOrdcd: Erase strRstval
            If Trim$(varTmp) = "1" Then
                For intCol = 5 To .MaxCols
                    .GetText intCol, intRow, varTmp
                    If Trim$(varTmp) <> "" Then
                        .GetText intCol, 0, varTmp
                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                        If Not itemX Is Nothing Then
                            .GetText intCol, intRow, varTmp
                            strTestcd = itemX.ListSubItems(1)
                            intPos = InStr(strTestcd, ",")
                            If intPos > 0 Then
                                Do While intPos > 0
                                    
                                    blnFlag = False
                                    Set mAdoRs = f_subSet_TestList(strBarno)
                                    Do Until mAdoRs.EOF
                                        If mAdoRs("itemCode") = Mid$(strTestcd, 1, intPos - 1) Then blnFlag = True: Exit Do
                                        mAdoRs.MoveNext
                                    Loop
                                    Set adoRS = Nothing: mAdoRs.Close
                                    
                                    strTestcd = Mid$(strTestcd, intPos + 1)
                                    intPos = InStr(strTestcd, ",")
                                    
                                    AdoCn_SQL.Execute "Exec InterfaceResult_INSERT_sp '" & strBarno & "','" & strTestcd & "','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & Trim$(varTmp) & "','" & INS_CODE & "','' ", sqlRet
                                
                                    If sqlRet = 1 Then
                                        lblStatus.Caption = "저장 성공!!"
                                    Else
                                        lblStatus.Caption = "저장 실패!!"
                                    End If
                                Loop
                            Else
                                blnFlag = False
                                Set mAdoRs = f_subSet_TestList(strBarno)
                                Do Until mAdoRs.EOF
                                    If mAdoRs("itemCode") = strTestcd Then blnFlag = True: Exit Do
                                    mAdoRs.MoveNext
                                Loop
                                Set adoRS = Nothing: mAdoRs.Close
                                
                                If blnFlag Then
                                    AdoCn_SQL.Execute "Exec InterfaceResult_INSERT_sp '" & strBarno & "','" & strTestcd & "','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & Trim$(varTmp) & "','" & INS_CODE & "','' ", sqlRet
                                
                                    If sqlRet = 1 Then
                                        lblStatus.Caption = "저장 성공!!"
                                    Else
                                        lblStatus.Caption = "저장 실패!!"
                                    End If
                                End If
                            End If
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

                'If intCnt > 0 Then
                    If strErrMsg = "" Then
                        If sqlRet = 1 Then
                            sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
                                     " where SPCNO   = '" & strBarno & "'" & _
                                     "   and TRANSDT = '" & mskRstDate.Text & "'" '& _
                                     '"   and TRANSTM = '" & strTime & "'"
                            AdoCn_Jet.Execute sqlDoc
                        End If
                    Else
                        MsgBox strErrMsg, vbInformation, Me.Caption
                    End If
                'Else
                '    MsgBox "검체번호 [" + strBarno + "]를 저장하지 못했습니다.", vbInformation, Me.Caption
                'End If
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
    
    Call Com_Output(charCOM_Convert(COM_ENQ))

End Sub

Private Sub cmdInitial_Click()
    
    'comEQP.PortOpen = False
    'Call cmdRun           ' 실행
    
    If comEQP.PortOpen = True Then
        comEQP.Output = STX & "0I" & Space(1) & vbCr & vbLf & "^" & ETX
        
        tmrInitial.Enabled = True
        tmrInitial.Interval = 100
        Timer1.Enabled = True
        Timer1.Interval = 5000
    End If
    
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
        .MaxRows = 15
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    sqlDoc = "select SPCNO, TESTCD, EQUIPCD, TRANSTM, RSTVAL, REFVAL, TRANSDT, EQPNUM, NAME, PNO" & _
             "  from INTERFACE003" & _
             " where TRANSDT >= '" & mskRstDate.Text & "'" & _
             "   and TRANSDT <= '" & mskRstDate1.Text & "'" & _
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
'    Do While Not adoRS.EOF
'        With spdResult2
'            If strSpcno <> Trim$(adoRS(0) & "") + Trim$(adoRS(6) & "") Then
''                intRow = intRow + 1
''                If intRow > .MaxRows Then .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
'
''                .SetText 1, intRow, "1"
''                .SetText 2, intRow, Trim$(adoRS(0) & "")
'
'                Set mAdoRs = f_subSet_TestList(Trim$(adoRS(0)))
'                If RecordChk = False Then
'                    adoRS.MoveNext
'                Else
'                    intRow = intRow + 1
'                    If intRow > .MaxRows Then .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
'                    .SetText 1, intRow, "1"
'                    .SetText 2, intRow, Trim$(adoRS(0) & "")
'
'                    .SetText 3, intRow, mAdoRs("patientname")
'                    .SetText 4, intRow, mAdoRs("patientnumber")
'                    .SetText .MaxCols, intRow, Trim$(adoRS(6) & "")
'                End If
'            End If
'            If RecordChk = False Then
'                adoRS.MoveNext
'            Else
'                strSpcno = Trim$(adoRS(0) & "") + Trim$(adoRS(6) & "")
'                Set itemX = lvwCuData.FindItem(Trim$(adoRS(7) & ""), lvwTag, , lvwWhole)
'                If Not itemX Is Nothing Then
'                    intCol = itemX.Index + 5
'                    .SetText intCol, intRow, Trim$(adoRS(4)) & ""
'                    .Col = intCol:  .Row = intRow:  .ForeColor = IIf(Trim$(adoRS(5) & "") <> "", vbRed, vbBlack)
'                End If
'            End If
'        End With
'        adoRS.MoveNext
'    Loop
    
    Do While Not adoRS.EOF
        With spdResult2
            If strSpcno <> Trim$(adoRS(0) & "") + Trim$(adoRS(6) & "") Then
                intRow = intRow + 1
                If intRow > .MaxRows Then .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
                .SetText 1, intRow, "1"
                .SetText 2, intRow, Trim$(adoRS(0) & "")
                .SetText 3, intRow, Trim$(adoRS(8) & "")
                .SetText 4, intRow, Trim$(adoRS(9) & "")
                .SetText .MaxCols, intRow, Trim$(adoRS(6) & "")
            End If
            strSpcno = Trim$(adoRS(0) & "") + Trim$(adoRS(6) & "")
            Set itemX = lvwCuData.FindItem(Trim$(adoRS(7) & ""), lvwTag, , lvwWhole)
            If Not itemX Is Nothing Then
                intCol = itemX.Index + 5
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
        With spdResult2
            For intRow = 1 To .MaxRows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 2, "1", "")
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

Private Sub cmdWordQuery_Click()

    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Privete sub cmdWorkQuery_Click()"
    
    Dim strKeyno    As String
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String, strBarno()   As String
    Dim strTestcd() As String, strTPid()    As String, strTPnm() As String
    Dim strEqpCd    As String
    Dim intRow  As String, intIdx  As Integer, intCol   As Integer
    Dim itemX   As ListItem
    
    If Trim(cboOrder.Text) = "" Then
        lblStatus.Caption = "조회차수가 없습니다"
        Exit Sub
    End If
    
    '-- WorkList조회
    Set mAdoRs = f_subSet_WorkList
    
    If RecordChk = False Then
        Exit Sub
    End If
    
    With spdWorkList
        .MaxRows = 14
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
            If strKeyno <> mAdoRs.Fields("Barcodenumber") Then
                intRow = intRow + 1
                If intRow > .MaxRows Then .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
                
                .SetText 1, intRow, "1"
                .SetText 2, intRow, mAdoRs("Barcodenumber")
                .SetText 3, intRow, mAdoRs("PatientName")
                .SetText 4, intRow, Trim(cboOrder.Text) + "-" + CStr(mAdoRs("seq_Number"))
                .SetText 5, intRow, mAdoRs("PatientNumber")

                '-- 검사항목조회
                Set mAdoRs1 = New Recordset
                Set mAdoRs1 = f_subSet_TestList(mAdoRs("Barcodenumber"))
                
                Do Until mAdoRs1.EOF
                    strEqpCd = f_funGet_CODE(mAdoRs1("itemCode"))
                    
                    Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                    If Not itemX Is Nothing Then .SetText 4 + itemX.Index, intRow, "V"
                    Set itemX = Nothing
                    mAdoRs1.MoveNext
                Loop
            End If
            strKeyno = mAdoRs("Barcodenumber")
        End With
        intIdx = intIdx + 1
        mAdoRs.MoveNext
    Loop
    Call cmdWorkList_Click
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)

End Sub

Private Sub cmdWorkList_Click()

    Dim varTmp  As Variant
    Dim intRow1 As Integer, intRow2 As Integer
    Dim intIdx  As Integer
    Dim Rev     As Long
    Dim Test_Cd() As String, strPid()   As String, strPnm() As String
    Dim itemX As ListItem
    Dim blnFlag As Boolean
    Dim strBarno    As String, strSPid  As String, strSPnm   As String
    
    Dim strEqpCd    As String
    
    blnFlag = False
    With spdWorkList
        For intRow1 = 1 To .MaxRows
            .GetText 1, intRow1, varTmp
            If Trim$(varTmp) = "1" Then
                .GetText 3, intRow1, varTmp:    strSPnm = Trim$(varTmp)
                .GetText 4, intRow1, varTmp:    strSPid = Trim$(varTmp)
                .GetText 2, intRow1, varTmp:    strBarno = Trim$(varTmp)
                
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
                    If Len(strBarno) = 11 Then
                        Do Until mAdoRs.EOF
                            strEqpCd = f_funGet_CODE(mAdoRs("itemCode"))
                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then
                                blnFlag = True
                                spdResult1.Row = intRow2
                                spdResult1.Col = itemX.Index + 4
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
                        MsgBox strBarno & "- 해당검체의 검사는 완료되었습니다.", vbOKOnly + vbExclamation
                    End If
                End If
                spdResult1.SetText 1, intRow2, "1"
                spdResult1.MaxRows = intRow2

                .SetText 1, intRow1, ""
            End If
        Next
    End With
End Sub
Private Sub comEQP_OnComm()
    
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim strDta      As String
    Dim Arr()       As Byte
    Dim strdata     As String


'    strDta = "2R 000A0406151941 003-03           06/15/04 10:47:11     1 8.57   2 3.45   3 10.7   4 31.2   5 90.3   6 30.9   7 34.2  51 35.0  72 31.5   8 13.4   9 2.66  10  568  11  6.7  20 88.0  21  7.4  22  3.7  23  0.4  24  0.2  25  0.3  14 7.54  15 0.63  16 0.32  17 0.03  18 0.02  19 0.03  50 9.62  52 8.57 19163.90 19266.82  83 27.9  84 10.9  42    +  12 46.3  13 0.38  26 2.03  27 -0.8  73 3.87 "
'    strDta = strDta & "  1 7.52   2 4.56   3 14.2   4 41.9   5 91.9   6 31.1   7 33.8  51 34.0  72 31.1   8 11.5   9 2.50  10  334  11  6.8  20 57.3  21 25.5  22  5.2  23 10.0  24  0.7  25  1.3  14 4.31  15 1.92  16 0.39  17 0.75  18 0.05  19 0.10  50 7.59  52 7.52 19168.35 19281.91  83 40.3  84 16.0  12 46.9  13 0.23  26 2.18  27  6.9  73 3.19"
'    strDta = strDta & "4"
'    GoTo rst
    
'    strDta = "" & "1Q" & Space(1) & "000A0407015451" & vbCr & vbLf & "" & ""
'    strDta = "2Q 000A04070154512"
'    strDta = "2Q 0000041123T001"
    strDta = "2Q 0000041123T010"
'2Q 0000041126T010
'!
    
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
'rst:
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
    Dim intIdx1     As Integer, intIdx2     As Integer
    Dim strTmp1     As String, strTmp2      As String
    Dim intPos1     As Integer, intPos2     As Integer
    Dim strDta()    As String, intCnt       As Integer
    Dim strRec      As String, strbuff      As String
    Dim Com_sTemp   As String
    Dim SendData As String
    
    strRec = RecData
    Debug.Print "Advia120==> " & strRec
    
    For intIdx1 = 1 To Len(strRec)
        strbuff = Mid$(strRec, intIdx1, 1)
        
        f_strBuffer = f_strBuffer + strbuff
        stxchk = InStr(f_strBuffer, Chr(2))
        etxchk = InStr(f_strBuffer, Chr(3))
        If stxchk <> 0 And etxchk <> 0 Then
            Com_sTemp = Mid$(f_strBuffer, stxchk, etxchk)
            f_strBuffer = Mid$(f_strBuffer, etxchk + 1)
            If InStr(Mid$(Com_sTemp, 1, 7), "R ") > 0 Then
                Timer1.Enabled = False
                Call psDataDefine(Com_sTemp, fChannel(), spdResult1)
                comEQP.Output = Mid(Com_sTemp, 2, 1)
                sndMsg = "3Z" & Space(18) & "0" & vbCr & vbLf
                comEQP.Output = STX & sndMsg & Get_CRC(sndMsg) & ETX

                Timer1.Enabled = True
                Timer1.Interval = 15000

            ElseIf InStr(Mid(Com_sTemp, 1, 10), "Q") > 0 Then
                Timer1.Enabled = False
                Call RequestDefine(Com_sTemp, fChannel(), spdResult1)
                comEQP.Output = Mid(Com_sTemp, 2, 1)
                CHR_MT = Mid(Com_sTemp, 2, 1)
                
                sndMsg = Chr(Asc(CHR_MT) + 1) & "Y" & Space(2) & _
                         "0000" & PatientID & Space(114) & _
                         vbCr & vbLf & _
                         sChannel & _
                         vbCr & vbLf
                
                sOrderLst = STX & sndMsg & Get_CRC(sndMsg) & ETX
                comEQP.Output = sOrderLst
                Timer1.Enabled = True
                Timer1.Interval = 15000
                Debug.Print "Host==> " & sOrderLst
                sndMsg = ""
                sOrderLst = ""
                sChannel = ""
            ElseIf InStr(Mid(Com_sTemp, 1, 10), "S") > 0 Then
                comEQP.Output = Mid(Com_sTemp, 2, 1)
                sndMsg = "0I" & Space(1) & vbCr & vbLf
                comEQP.Output = STX & sndMsg & Get_CRC(sndMsg) & ETX
                tmrInitial.Enabled = True
                tmrInitial.Interval = 100
            ElseIf InStr(Mid(Com_sTemp, 1, 10), "E") > 0 Then
                Timer1.Enabled = False
                comEQP.Output = Mid(Com_sTemp, 2, 1)
                Timer1.Enabled = True
                Timer1.Interval = 15000
            End If
            f_strBuffer = ""
            Com_sTemp = ""
        End If
    Next

End Sub

Private Sub RequestDefine(ByVal strdata As String, ByRef brChannel() As String, ByVal brSpread As Object) ', ByVal brOst As String) ' ByRef brItemdeci() As String)
    
    Dim Loop_Count, pDoCount As Integer
    Dim FunStr1, FunStr2 As String
    Dim PatientNo As String
    Dim ii As Integer
    Dim Testcd As String
    Dim OutputData As String
    Dim EndStr, strEqpCd As String
    Dim intRow  As Integer
    Dim itemX As ListItem
    Dim sTemp, ssTemp1, ssTemp2 As String
    Dim sETBTemp    As String
    Dim chkWbc      As Integer

    On Error GoTo errRequest
    
    PatientID = Mid(strdata, 9, 10)
    chkWbc = 0
    
    With spdResult1
        Set mAdoRs = f_subSet_TestList(Trim(PatientID))
        If RecordChk = True Then
            If Not mAdoRs.EOF Then
                .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
                intRow = .MaxRows
                .SetText 1, intRow, "1"
'                    .SetText 2, intRow, mAdoRs("labdate") & mAdoRs("numgbn") & mAdoRs("labsqno")
                .SetText 2, intRow, PatientID
                .SetText 3, intRow, Trim(mAdoRs("patnm"))
                .SetText 5, intRow, Trim(mAdoRs("idno"))
                Do Until mAdoRs.EOF
                    strEqpCd = f_funGet_CODE(Trim(mAdoRs("ordcd"))) ' & Trim(mAdoRs("subcd")))
                    If strEqpCd = "01" Then
                        chkWbc = chkWbc + 1
                    End If
                    Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                    If Not itemX Is Nothing Then
                        If strEqpCd = itemX.Tag Then
                            spdResult1.Row = intRow
                            spdResult1.Col = itemX.Index + 5
                            spdResult1.BackColor = &HC6FEFF
                            DoEvents
                            sChannel = sChannel & Format(itemX.Tag, "000")
                        End If
                        'sChannel = sChannel & Format(itemX.Tag, "000")
                    End If
                    mAdoRs.MoveNext
                    Debug.Print mAdoRs("ordcd")
                Loop
'                If chkWbc > 0 Then
'                    sChannel = sChannel & "020" & "021" & "022" & "023" & "024" & "025"
'                    chkWbc = 0
'                End If
                sChannel = "001" & sChannel
            End If
        End If
    End With
    
    Exit Sub
    
errRequest:

End Sub

Public Function Get_CRC(ByVal StrCRC As String) As String
    Dim bytTemp() As Byte
    Dim i As Integer
    Dim XOR_Rst As Byte
    
    bytTemp = StrConv(StrCRC, vbFromUnicode)
    XOR_Rst = bytTemp(LBound(bytTemp))
    
    For i = (LBound(bytTemp) + 1) To UBound(bytTemp)
        XOR_Rst = XOR_Rst Xor bytTemp(i)
    Next i
    
    If XOR_Rst = COM_ETX Then
        XOR_Rst = &H7F
    End If
    
    Get_CRC = Chr(XOR_Rst)
End Function

Private Sub psDataDefine(ByVal strdata As String, ByRef brChannel() As String, ByVal brSpread As Object) ', ByVal brOst As String) ' ByRef brItemdeci() As String)

    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Privete sub psDataDefine()"
    
    Dim ssTemp1 As String
    Dim ssTemp2 As String
    Dim ssTemp3 As String
    
    Dim sqlDoc  As String, sqlRet   As Integer
    Dim sTemp      As String
    Dim Channel_No As Integer       ' 검사항목 번호 : Channel No
    Dim pGrid_Point As Integer
    Dim pDoCount   As Integer
    Dim Loop_Count As Integer
    Dim varTmp      As Variant
    Dim strTmp, ssTemp     As String
    Dim intRow, intRow1, intRow2, sRow   As Long, intCol As Integer, intIdx  As Integer
    Dim Max_Arary_Cnt As Integer
    Dim Gnum   As String
    Dim strRstval As String, strRefval As String
    Dim strBarno    As String, strTime  As String, strDate  As String
    Dim strSeqno    As String, strRackno As String

    Dim strOrdLst As String, strPid() As String, strPnm() As String
    Dim intRet      As Integer
    Dim ii As Integer
    Dim itemX   As ListItem
    Dim sDeCnt   As Integer, strEqpCd    As String, sChannel As String
    Dim strKeyno    As String, sCheck As Boolean
    Dim FunStr1 As String
    Dim intChannel As Integer
    Dim strSubcd, strOrdcd   As String
    
    sTemp = strdata
    
    For Loop_Count = 1 To 100: fAdvia120(Loop_Count) = "": Next Loop_Count
    
    sTemp = Text_Change(sTemp, Chr$(10), "")
    sTemp = Text_Change(sTemp, Chr$(13), "")
    
    fAdvia120(0) = Mid$(sTemp, 9, 10)                     ' 항목 2 "A0403310061"   환자번호
    
    sTemp = Mid(sTemp, 57)
    intChannel = 1
    For pDoCount = 1 To Len(sTemp) Step 9
        fAdvia120(intChannel) = Mid$(sTemp, pDoCount, 9)   ' channel result
        Debug.Print Format(intChannel, "00") & ":" & fAdvia120(intChannel)
        intChannel = intChannel + 1
    Next pDoCount
 
    Max_Arary_Cnt = spdResult1.MaxCols
      
    pGrid_Point = 0
      
    Dim sSeq As String
    Dim sCol As Integer
      
    strTmp = ""
    If Len(fAdvia120(0)) > 0 Then
        intRow = 0
        With spdResult1
            If optSeq.Value = True Then
                sSeq = Trim(fAdvia120(0))
            ElseIf optBar.Value = True Then
                If Len(Trim(fAdvia120(0))) < 10 Then
                    Exit Sub
                Else
                    sSeq = Trim(fAdvia120(0))
                End If
            End If
            
            sCol = 2
            pGrid_Point = SeqSearch(spdResult1, sSeq, sCol)
            
            If pGrid_Point = 0 Then
                pGrid_Point = SeqNullSearch(spdResult1, sSeq, sCol)
                If pGrid_Point = 0 Then .MaxRows = .MaxRows + 1: pGrid_Point = .MaxRows
            End If

            .SetText 2, pGrid_Point, sSeq

'            -- 검사항목조회
            Set mAdoRs1 = New Recordset
            Set mAdoRs1 = f_subSet_TestList(sSeq)

            If RecordChk = False Then
                Exit Sub
            End If
            Do Until mAdoRs1.EOF
                .SetText 1, pGrid_Point, "1"
'                .SetText 2, pGrid_Point, mAdoRs1("Barcodenumber")
'                .SetText 3, pGrid_Point, mAdoRs1("PatientName"):        pName = mAdoRs1("PatientName")
'                '.SetText 4, pGrid_Point, Trim(cboOrder.Text) + "-" + CStr(mAdoRs1("seq_Number"))
'                .SetText 5, pGrid_Point, mAdoRs1("PatientNumber"):      pNo = mAdoRs1("PatientNumber")

'                .SetText 2, intRow, mAdoRs("labdate") & mAdoRs("numgbn") & mAdoRs("labsqno")
                .SetText 2, pGrid_Point, sSeq
                .SetText 3, pGrid_Point, Trim(mAdoRs1("patnm")): pName = Trim(mAdoRs1("patnm"))
                .SetText 5, pGrid_Point, Trim(mAdoRs1("idno")): pNo = Trim(mAdoRs1("idno"))

                strEqpCd = f_funGet_CODE(mAdoRs1("ordcd"))
                mAdoRs1.MoveNext
            Loop
           
            If pGrid_Point > 0 Then
                For intCol = 6 To .MaxCols
                    strRstval = ""
                    .GetText intCol, 0, varTmp
                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                    If Not itemX Is Nothing Then
                        For intIdx = 1 To 50 'Max_Arary_Cnt
                            If Len(fAdvia120(intIdx)) > 0 Then
                                If Val(Mid(fAdvia120(intIdx), 1, 3)) = Val(itemX.Tag) Then
                                    strRstval = Mid(fAdvia120(intIdx), 4)
                                    
                                    strDate = Format$(Now, "YYYYMMDD"):     strTime = Format$(Now, "MMSS")
                                    .SetText intCol, pGrid_Point, strRstval
'                                    strRefval = ""
'                                    If CLng(strRstval) < CLng(itemX.SubItems(8)) Then
'                                        strRefval = "L"
'                                    ElseIf CLng(strRstval) > CLng(itemX.SubItems(9)) Then
'                                        strRefval = "H"
'                                    Else
'                                        strRefval = ""
'                                    End If
                                    .Col = intCol:  .Row = pGrid_Point
                                                    .ForeColor = IIf(Trim$(strRefval) <> "", vbRed, vbBlack)
                                    
                                    sqlDoc = "Update INTERFACE003" & _
                                             "   set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRefval & "'" & _
                                             " where SPCNO   = '" & sSeq & "'" & _
                                             "   and EQPNUM  = '" & itemX.Tag & "'" & _
                                             "   and TESTCD  = '" & itemX.Text & "'" & _
                                             "   and TRANSDT = '" & strDate & "'" & _
                                             "   and TRANSTM = '" & strTime & "'"
                                    AdoCn_Jet.Execute sqlDoc, sqlRet
                                    
                                    sqlDoc = "insert into INTERFACE003(" & _
                                             "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                             "    values( '" & sSeq & "', '" & itemX.Text & "', '" & itemX.Tag & "'," & _
                                             "            '" & strDate & "', '" & strTime & "'," & _
                                             "            '" & strRstval & "', '" & strRefval & "'," & _
                                             "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
                                    AdoCn_Jet.Execute sqlDoc
                                    
                                    If sqlRet = 0 And chkAuto.Value = vbChecked And .BackColor = &HC6FEFF Then
                                        If Len(itemX) > 3 Then
                                            strSubcd = Mid(itemX, 4, 2)
                                        Else
                                            strSubcd = ""
                                        End If
                                        strOrdcd = Mid(itemX, 1, 3)
                                        
                                        mAdoRs1.MoveFirst
                                        
                                        AdoCn_SQL.Execute "Exec LAB_DB..STR_YEJ_LABResult_I '" & Trim(mAdoRs1("labdate")) & "','" & Trim(mAdoRs1("numgbn")) & "','" & Trim(mAdoRs1("labsqno")) & "','" & Trim(mAdoRs1("slipcd")) & "'" & _
                                        ", '" & strOrdcd & "', '" & Trim(mAdoRs1("spccd")) & "', '" & strSubcd & "', '" & strRstval & "', '" & strRefval & "', '', '',  '" & strDate & "', '', '" & Trim(mAdoRs1("deptcd")) & "', '" & Trim(mAdoRs1("orddate")) & "'" & _
                                        ", '" & Trim(mAdoRs1("seqno")) & "', '" & Trim(mAdoRs1("rslipcd")) & "', '" & strOrdcd & "', '" & Trim(mAdoRs1("spccd")) & "', '" & Trim(mAdoRs1("idno")) & "', '', '', '' , '', '', ''", sqlRet
                                    
                                        If sqlRet = 1 Then
                                            lblStatus.Caption = "저장 성공!!"
                                            spdResult1.SetText 1, pGrid_Point, "0"
                                            spdResult1.Row = pGrid_Point
'                                            spdResult1.Col = -1:    spdResult1.BackColor = &HFFF8F0
                                            
                                            spdResult1.Col = 2
                                            spdResult1.BackColor = vbCyan
                                            spdResult1.Col = 3
                                            spdResult1.BackColor = vbCyan
                                            spdResult1.Col = 4
                                            spdResult1.BackColor = vbCyan
                                            spdResult1.Col = 5
                                            spdResult1.BackColor = vbCyan
                                            
                                            sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
                                                     " where SPCNO   = '" & strBarno & "'" & _
                                                     "   and TRANSDT = '" & mskRstDate.Text & "'" & _
                                                     "   and TRANSTM = '" & strTime & "'"
                                            AdoCn_Jet.Execute sqlDoc
                                        Else
                                            'lblStatus.Caption = "저장 실패!!"
                                        End If
                                                        
                                    End If
                                    Exit For
                                End If
                            End If
                        Next intIdx
                    End If
                    Set itemX = Nothing
                Next
                '-----------------------------------------------------------------------
            End If
        End With
    End If
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)

End Sub

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

Private Function SeqNullSearch(ByVal brSpread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqNullSearch = 0
    If brSpread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brSpread
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

Private Function SeqSearch(ByVal brSpread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long
Dim varTmp As String

    SeqSearch = 0
    If brSpread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brSpread
        If optBar.Value = True Then
            For sCnt = 1 To .MaxRows
                .Row = sCnt
                .Col = brCol
'                .GetText 0, sCnt, varTmp
                If Trim(.Text) = brSeq Then
                    SeqSearch = sCnt
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            Next sCnt
        Else
            For sCnt = 1 To .MaxRows
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

    Dim strDta(1 To 2)  As String
    Dim intIdx          As Integer
    
    Dim byeTemp()       As Byte
    
    Call comEQP_OnComm
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
    mskOrdDate.Text = Format$(Now, "YYYYMMDD")
    mskRstDate1.Text = Format$(Now, "YYYYMMDD")
    Open App.Path + "\" + "Advia120.log" For Append As #1

    Print #1, Chr(13) + Chr(10);
    
    f_strJOB_FLAG = "1":    f_intSampleNo = 0
    cboRstgbn(0).ListIndex = 0: cboRstgbn(1).ListIndex = 2
    tabWork.Tab = 0
    
'    comEQP.Output = STX & "0I" & Space(1) & vbCr & vbLf & "^" & ETX
    
    DoEvents
    
'    If optSeq.Value = True Then
'        Label6.Visible = True
'        cboOrder.Visible = True
'        cmdWordQuery.Visible = True
'        cmdStartNo.Visible = True
'        '-- 해당일자의 차수불러오기
'        Call f_subSet_DateOrder
'    ElseIf optBar.Value = True Then
'        Label6.Visible = False
'        cboOrder.Visible = False
'        cmdWordQuery.Visible = False
'        cmdStartNo.Visible = False
'    End If
'
'    '-- 해당일자의 차수불러오기
'    Call f_subSet_DateOrder
    
    stxchk = 0
    
End Sub

Private Function f_subSet_WorkList()
    Dim sqlRet      As Integer
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
    Set AdoRs_SQL = New ADODB.Recordset
    AdoRs_SQL.CursorLocation = adUseClient
    AdoRs_SQL.Open AdoCn_SQL.Execute("Exec interface_WorkList_SELECT_sp '" & INS_CODE & "','" & mskOrdDate.FormattedText & "','" & cboOrder.Text & "'", sqlRet)
    
    If sqlRet = 0 Then
        Set f_subSet_WorkList = Nothing
        MsgBox "해당차수의 검사는 완료되었습니다. 차수를 확인하세요.", vbOKOnly + vbExclamation
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

Private Function f_subSet_TestList(ByVal strBarcode As String)
   Dim sqlRet      As Integer
   Dim gSql        As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
    
    Set AdoRs_SQL = New ADODB.Recordset
    AdoRs_SQL.CursorLocation = adUseClient
    
    gSql = ""
    gSql = "select a.*,b.patnm,b.idno from lab_db..lab030m a, pat_db..pat010m b"
    gSql = gSql & " where a.labdate = '" & "20" & Mid(strBarcode, 1, 6) & "'"
    gSql = gSql & "   and a.numgbn = '" & Mid(strBarcode, 7, 1) & "'"
    gSql = gSql & "   and a.labsqno = '" & "00" & Mid(strBarcode, 8, 3) & "'"
    gSql = gSql & "   and a.idno = b.idno"
    gSql = gSql & "   and a.slipcd = 'HB'"
    
    AdoRs_SQL.Open gSql, AdoCn_SQL, adOpenStatic, adLockReadOnly
   
    If AdoRs_SQL.RecordCount = 0 Then
        Set f_subSet_TestList = Nothing
        lblStatus.Caption = strBarcode & "-해당검체는 검사가 완료되었습니다."
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

Private Sub f_subSet_DateOrder()
    Dim sqlRet      As Integer
    Dim iCnt        As Integer
    Dim itemCnt     As Integer
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_DateOrder() As ADODB.Recordset"
    
    Set AdoRs_SQL = New ADODB.Recordset
    AdoRs_SQL.CursorLocation = adUseClient
    AdoRs_SQL.Open AdoCn_SQL.Execute("Exec Interface_WorkListStep_SELECT_sp '" & INS_CODE & "','" & mskOrdDate.FormattedText & "'", sqlRet)

    cboOrder.Clear
    If sqlRet <> 0 Then
        iCnt = 0
        If IsNull(AdoRs_SQL.Fields(0)) Then
            lblStatus.Caption = "조회차수가 없습니다"
            Exit Sub
        Else
            Do Until AdoRs_SQL.EOF
                If AdoRs_SQL.Fields(0) > 1 Then
                    For itemCnt = 1 To AdoRs_SQL.Fields(0)
                        cboOrder.AddItem itemCnt
                        iCnt = iCnt + 1
                    Next
                Else
                    cboOrder.AddItem AdoRs_SQL.Fields(0)
                    iCnt = iCnt + 1
                End If
                
    '            cboOrder.AddItem AdoRs_SQL.Fields(0)
    '            iCnt = iCnt + 1
                AdoRs_SQL.MoveNext
            Loop
            cboOrder.ListIndex = iCnt - 1
        End If
    End If

    Set AdoRs_SQL = Nothing

Exit Sub

ErrorTrap:
    Set AdoRs_SQL = Nothing
    Call ErrMsgProc(CallForm)

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
    
    If KeyAscii = vbKeyReturn Then Call f_subSet_DateOrder
    
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


Private Sub optBar_Click()
    Label6.Visible = False
    cboOrder.Visible = False
    cmdWordQuery.Visible = False
    cmdStartNo.Visible = False
End Sub

Private Sub optSeq_Click()
    Label6.Visible = True
    cboOrder.Visible = True
    cmdWordQuery.Visible = True
    cmdStartNo.Visible = True
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
            Call Com_Output(Order.MSG_ENQ)
        Case 2
            Call Com_Output(Order.MSG_HEADER)
        Case 3
            Call Com_Output(Order.MSG_PATIENT)
        Case 4
            Call Com_Output(Order.MSG_ORDER)
        Case 5
            Call Com_Output(Order.MSG_TERMINATION)
        Case 6
            Call Com_Output(Order.MSG_EOT)
        Case Else
    End Select
    
End Sub

Private Sub Timer1_Timer()
    
    If comEQP.PortOpen = True Then
        comEQP.Output = STX & "0I" & Space(1) & vbCr & vbLf & "^" & ETX
        tmrInitial.Enabled = True
    End If
End Sub

Private Sub tmrInitial_Timer()
    
'    comEQP.Output = STX & "0I" & Space(1) & vbCr & vbLf & "^" & ETX
    comEQP.Output = STX & "1S" & Space(10) & vbCr & vbLf & "e" & ETX
    tmrInitial.Enabled = False

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
    Call Com_Output(charCOM_Convert(txtCom.SelText))
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
        Call Com_Output(charCOM_Convert(txtCOM2.Text))
    Else
        Call Com_Output(charCOM_Convert(txtCOM2.SelText))
    End If
    
End Sub
' ------------------------------------------------------------------------
' 통신상태 확인 관련이벤트



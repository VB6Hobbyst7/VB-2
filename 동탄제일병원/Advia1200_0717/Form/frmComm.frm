VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
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
   ClientWidth     =   15390
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   15390
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  '�ִ�ȭ
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   8850
      Top             =   5160
   End
   Begin MSCommLib.MSComm comEQP 
      Left            =   6690
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin VB.Timer tmrReceive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7860
      Top             =   5130
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8340
      Top             =   5130
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   7230
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
         Name            =   "����"
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
      Begin VB.Timer tmrWorking 
         Interval        =   100
         Left            =   5520
         Top             =   60
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1440
         Top             =   300
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
         Caption         =   "Run"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Caption         =   "Stop"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Caption         =   "Clear"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Caption         =   "Close"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin VB.Image imgBack 
         BorderStyle     =   1  '���� ����
         Height          =   1050
         Index           =   0
         Left            =   4020
         Picture         =   "frmComm.frx":5794
         Stretch         =   -1  'True
         Top             =   -210
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Image imgLogo 
         Height          =   240
         Index           =   0
         Left            =   3630
         Picture         =   "frmComm.frx":6F67
         Top             =   180
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "�˾��� ==>"
         Height          =   225
         Index           =   1
         Left            =   2700
         TabIndex        =   73
         Top             =   210
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "�۾���� ��.."
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   " ���� :"
         BeginProperty Font 
            Name            =   "����"
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
      Align           =   1  '�� ����
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Caption         =   " Communication"
      SubCaption      =   "�˻� ���� ����Ͽ� ����� ���� �մϴ�."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Receive : "
         Height          =   180
         Left            =   13995
         TabIndex        =   4
         Top             =   195
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Send : "
         Height          =   180
         Left            =   12960
         TabIndex        =   3
         Top             =   195
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Port : "
         Height          =   180
         Index           =   0
         Left            =   11865
         TabIndex        =   2
         Top             =   195
         Width           =   510
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   14865
         Picture         =   "frmComm.frx":74F1
         Top             =   165
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   13575
         Picture         =   "frmComm.frx":7A7B
         Top             =   165
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12375
         Picture         =   "frmComm.frx":8005
         Top             =   165
         Width           =   240
      End
   End
   Begin TabDlg.SSTab tabWork 
      Height          =   8370
      Left            =   30
      TabIndex        =   7
      Top             =   570
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   14764
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   16711680
      TabCaption(0)   =   " ��    WorkList     "
      TabPicture(0)   =   "frmComm.frx":858F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "pnlCom"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SSPanel1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdPrint"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdStartNo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdSearch"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "spdWorklist"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtBarCode"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "pnlCom2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdRequist(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkAuto"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtResult"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "spdRstview"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdRackNo"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdWordQuery"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdEot"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdAppend(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Command1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "SSPanel2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdWorkList"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "List1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Frame3"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdOrder"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdPosNo"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Command2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdNext"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdPrevious"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtSeqNo"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "spdResult1"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      TabCaption(1)   =   " ��   ���� ���     "
      TabPicture(1)   =   "frmComm.frx":85AB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkExcel"
      Tab(1).Control(1)=   "spdResult2"
      Tab(1).Control(2)=   "lvwCuData"
      Tab(1).Control(3)=   "cmdAppend(1)"
      Tab(1).Control(4)=   "CommonDialog1"
      Tab(1).Control(5)=   "cmdSel(3)"
      Tab(1).Control(6)=   "cmdSel(2)"
      Tab(1).Control(7)=   "cmdExcel"
      Tab(1).Control(8)=   "cmdRstQuery"
      Tab(1).Control(9)=   "SSPanel"
      Tab(1).Control(10)=   "tblexcel"
      Tab(1).ControlCount=   11
      Begin FPSpread.vaSpread spdResult1 
         Height          =   4995
         Left            =   90
         TabIndex        =   85
         Top             =   360
         Width           =   15105
         _Version        =   196608
         _ExtentX        =   26644
         _ExtentY        =   8811
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   7
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   9
         MaxRows         =   5
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ShadowColor     =   14735309
         SpreadDesigner  =   "frmComm.frx":85C7
         UserResize      =   0
      End
      Begin VB.TextBox txtSeqNo 
         Alignment       =   1  '������ ����
         Height          =   300
         Left            =   12660
         MaxLength       =   12
         TabIndex        =   84
         Text            =   "0"
         Top             =   480
         Width           =   960
      End
      Begin BHButton.BHImageButton cmdPrevious 
         Height          =   330
         Left            =   90
         TabIndex        =   59
         Top             =   5400
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         Caption         =   "��"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         BackColor       =   16711680
         AlphaColor      =   255
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdNext 
         Height          =   330
         Left            =   330
         TabIndex        =   60
         Top             =   5400
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         Caption         =   "��"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "frmComm.frx":8B6A
         ForeColor       =   16711680
         BackColor       =   255
         AlphaColor      =   255
         ImgOutLineSize  =   3
      End
      Begin VB.CommandButton Command2 
         Caption         =   "TEST"
         Height          =   375
         Left            =   7620
         TabIndex        =   58
         Top             =   -60
         Visible         =   0   'False
         Width           =   1230
      End
      Begin BHButton.BHImageButton cmdPosNo 
         Height          =   375
         Left            =   8610
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         Caption         =   "Pos����"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Height          =   420
         Left            =   10830
         TabIndex        =   42
         Top             =   420
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "��������"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Frame Frame3 
         Height          =   315
         Left            =   90
         TabIndex        =   54
         Top             =   900
         Width           =   555
         Begin Threed.SSCommand cmdSel 
            Height          =   345
            Index           =   1
            Left            =   270
            TabIndex        =   56
            Top             =   0
            Visible         =   0   'False
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   609
            _StockProps     =   78
            BevelWidth      =   1
            Picture         =   "frmComm.frx":8FDC
         End
         Begin Threed.SSCommand cmdSel 
            Height          =   345
            Index           =   0
            Left            =   0
            TabIndex        =   55
            Top             =   0
            Visible         =   0   'False
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   609
            _StockProps     =   78
            ForeColor       =   14735310
            BevelWidth      =   1
            Picture         =   "frmComm.frx":945E
         End
      End
      Begin VB.CheckBox chkExcel 
         Appearance      =   0  '���
         BackColor       =   &H80000004&
         Caption         =   "Excel ����"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   -61080
         TabIndex        =   53
         Top             =   30
         Value           =   1  'Ȯ��
         Width           =   1245
      End
      Begin VB.ListBox List1 
         Height          =   2220
         ItemData        =   "frmComm.frx":98CC
         Left            =   7950
         List            =   "frmComm.frx":98CE
         TabIndex        =   49
         Top             =   6060
         Width           =   7215
      End
      Begin BHButton.BHImageButton cmdWorkList 
         Height          =   435
         Left            =   4140
         TabIndex        =   29
         Top             =   4170
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   767
         Caption         =   "WorkList ���"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         TabIndex        =   44
         Top             =   900
         Width           =   15015
         _Version        =   196608
         _ExtentX        =   26485
         _ExtentY        =   12965
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   7
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
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
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmComm.frx":98D0
         UserResize      =   0
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   465
         Left            =   12120
         TabIndex        =   39
         Top             =   5400
         Visible         =   0   'False
         Width           =   3075
         _Version        =   65536
         _ExtentX        =   5424
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
            Appearance      =   0  '���
            BackColor       =   &H00E0E0E0&
            Caption         =   "���Ϲ�ȣ"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1650
            TabIndex        =   41
            Top             =   90
            Width           =   1335
         End
         Begin VB.OptionButton optSeq 
            Appearance      =   0  '���
            BackColor       =   &H00E0E0E0&
            Caption         =   "�˻��ȣ"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   210
            TabIndex        =   40
            Top             =   90
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TEST"
         Height          =   375
         Left            =   6270
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   1230
      End
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   4830
         Left            =   -67980
         TabIndex        =   20
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
         Left            =   -62355
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
         Caption         =   "�������"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Height          =   420
         Index           =   0
         Left            =   13650
         TabIndex        =   38
         Top             =   405
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   741
         Caption         =   "�������"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Left            =   12420
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         Caption         =   "�ʱ�ȭ"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Height          =   390
         Left            =   9330
         TabIndex        =   45
         Top             =   5400
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   688
         Caption         =   "��ȸ"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Left            =   11130
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
         Caption         =   "Rack����"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin FPSpread.vaSpread spdRstview 
         Height          =   2865
         Left            =   90
         TabIndex        =   48
         Top             =   5400
         Width           =   7785
         _Version        =   196608
         _ExtentX        =   13732
         _ExtentY        =   5054
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   4
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
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
         MaxRows         =   8
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmComm.frx":9D44
         UserResize      =   0
      End
      Begin VB.TextBox txtResult 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1500
         Left            =   8190
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   50
         Top             =   6540
         Visible         =   0   'False
         Width           =   6600
      End
      Begin VB.CheckBox chkAuto 
         Appearance      =   0  '���
         Caption         =   "Auto Server"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   13830
         TabIndex        =   24
         Top             =   60
         Visible         =   0   'False
         Width           =   1320
      End
      Begin BHButton.BHImageButton cmdRequist 
         Height          =   390
         Index           =   2
         Left            =   7950
         TabIndex        =   52
         Top             =   5400
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   688
         Caption         =   "Last Order.."
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -65490
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin HSCotrol.UserPanel pnlCom2 
         Height          =   5385
         Left            =   8460
         TabIndex        =   10
         Top             =   8520
         Visible         =   0   'False
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   9499
         Bevel           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtCOM2 
            BeginProperty Font 
               Name            =   "����ü"
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
            ScrollBars      =   2  '����
            TabIndex        =   19
            Top             =   300
            Width           =   5730
         End
         Begin VB.Frame Frame2 
            Height          =   645
            Left            =   90
            TabIndex        =   11
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
               TabIndex        =   12
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "SUM"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               TabIndex        =   13
               Top             =   180
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Send"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Clear"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               TabIndex        =   15
               Top             =   180
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Receive"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "File Load"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               TabIndex        =   17
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "ACK"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               TabIndex        =   18
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "ENQ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
      Begin VB.TextBox txtBarCode 
         Height          =   300
         Left            =   11910
         MaxLength       =   12
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   750
      End
      Begin FPSpread.vaSpread spdWorklist 
         Height          =   3960
         Left            =   90
         TabIndex        =   57
         Top             =   900
         Width           =   3705
         _Version        =   196608
         _ExtentX        =   6535
         _ExtentY        =   6985
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   7
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
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
         MaxRows         =   5
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmComm.frx":A5C5
         UserResize      =   2
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   3
         Left            =   -74640
         TabIndex        =   21
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":AB40
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   -74910
         TabIndex        =   22
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":AFC2
      End
      Begin BHButton.BHImageButton cmdSearch 
         Height          =   420
         Left            =   90
         TabIndex        =   61
         Top             =   4890
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   741
         Caption         =   "��ȸ"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Height          =   420
         Left            =   7920
         TabIndex        =   62
         Top             =   420
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "���۹�ȣ����"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdPrint 
         Height          =   420
         Left            =   9210
         TabIndex        =   63
         Top             =   420
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   741
         Caption         =   "WorkSheet Print"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   465
         Left            =   90
         TabIndex        =   64
         Top             =   390
         Width           =   6465
         _Version        =   65536
         _ExtentX        =   11404
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
            ItemData        =   "frmComm.frx":B430
            Left            =   2850
            List            =   "frmComm.frx":B43A
            TabIndex        =   86
            Top             =   90
            Width           =   855
         End
         Begin VB.TextBox txtChart 
            Alignment       =   2  '��� ����
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   300
            Left            =   4980
            MaxLength       =   12
            TabIndex        =   66
            Top             =   90
            Width           =   1395
         End
         Begin VB.ComboBox cboComNm 
            Height          =   300
            ItemData        =   "frmComm.frx":B44A
            Left            =   4590
            List            =   "frmComm.frx":B44C
            TabIndex        =   65
            Top             =   480
            Visible         =   0   'False
            Width           =   1725
         End
         Begin MSMask.MaskEdBox mskOrdtime 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "H:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   4
            EndProperty
            Height          =   300
            Left            =   4560
            TabIndex        =   67
            Top             =   450
            Visible         =   0   'False
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   315
            Left            =   1500
            TabIndex        =   68
            Top             =   90
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   43646977
            CurrentDate     =   40248
         End
         Begin MSComCtl2.DTPicker dtpStartDt 
            Height          =   315
            Left            =   60
            TabIndex        =   69
            Top             =   90
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   43646977
            CurrentDate     =   40248
         End
         Begin VB.Label Label10 
            BackColor       =   &H00E0E0E0&
            Caption         =   "�� ��������."
            Height          =   255
            Left            =   5520
            TabIndex        =   72
            Top             =   840
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E0E0E0&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1350
            TabIndex        =   71
            Top             =   150
            Width           =   315
         End
         Begin VB.Label Label6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "���� :"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   70
            Top             =   150
            Width           =   1095
         End
      End
      Begin BHButton.BHImageButton cmdExcel 
         Height          =   420
         Left            =   -68940
         TabIndex        =   74
         Top             =   420
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   741
         Caption         =   "Excel ���� ���� / ���"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdRstQuery 
         Height          =   420
         Left            =   -70260
         TabIndex        =   75
         Top             =   420
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "��ȸ"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   465
         Left            =   -74910
         TabIndex        =   76
         Top             =   390
         Width           =   4545
         _Version        =   65536
         _ExtentX        =   8017
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
            ItemData        =   "frmComm.frx":B44E
            Left            =   2640
            List            =   "frmComm.frx":B45B
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   78
            Top             =   105
            Width           =   1770
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            ItemData        =   "frmComm.frx":B485
            Left            =   4590
            List            =   "frmComm.frx":B487
            TabIndex        =   77
            Top             =   480
            Visible         =   0   'False
            Width           =   1725
         End
         Begin MSMask.MaskEdBox MaskEdBox 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "H:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   4
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   4560
            TabIndex        =   79
            Top             =   450
            Visible         =   0   'False
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker dtpRsltDay 
            Height          =   315
            Left            =   1290
            TabIndex        =   80
            Top             =   90
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   43646977
            CurrentDate     =   40248
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "�˻����� :"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   90
            TabIndex        =   82
            Top             =   150
            Width           =   1125
         End
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "�� ��������."
            Height          =   255
            Left            =   5520
            TabIndex        =   81
            Top             =   840
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin FPSpread.vaSpread tblexcel 
         Height          =   675
         Left            =   -66420
         TabIndex        =   83
         Top             =   300
         Visible         =   0   'False
         Width           =   675
         _Version        =   196608
         _ExtentX        =   1191
         _ExtentY        =   1191
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
         SpreadDesigner  =   "frmComm.frx":B489
      End
      Begin HSCotrol.UserPanel pnlCom 
         Height          =   4725
         Left            =   2340
         TabIndex        =   31
         Top             =   1860
         Visible         =   0   'False
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   8334
         Bevel           =   1
         Moveble         =   -1  'True
         CloseEnabled    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Timer tmrOrder 
            Left            =   4110
            Top             =   1740
         End
         Begin VB.TextBox txtCom 
            BeginProperty Font 
               Name            =   "����ü"
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
            ScrollBars      =   2  '����
            TabIndex        =   32
            Top             =   315
            Width           =   11595
         End
         Begin VB.Frame Frame1 
            Height          =   645
            Left            =   45
            TabIndex        =   33
            Top             =   4020
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         DrawMode        =   5  'ī�� ���� �ƴ�
         X1              =   9480
         X2              =   15180
         Y1              =   5940
         Y2              =   5940
      End
      Begin VB.Label Label8 
         Caption         =   "�� Information List"
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
         TabIndex        =   51
         Top             =   5790
         Width           =   1755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "���/QC :"
         BeginProperty Font 
            Name            =   "����"
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

Private Const KEY_SEQ       As String = "KEY_SEQ"   ' "����"
Private Const KEY_PTID      As String = "KEY_PTID"  ' "��Ϲ�ȣ"
Private Const KEY_PTNM      As String = "KEY_PTNM"  ' "��  ��"
Private Const KEY_SPCNO     As String = "KEY_SPCNO" ' "��ü��ȣ"
Private Const KEY_EQPNO     As String = "KEY_EQPNO" ' "��ü��ȣ"
Private Const KEY_STAT      As String = "KEY_STAT"  ' "�� ��"
Private Const KEY_TEST      As String = "KEY_TEST"  ' "�˻��׸�"

Private Const TEST_NM_EQP   As String = "EQP_NM"    '��� �ڵ�
Private Const TEST_CD_LIS   As String = "LIS_CD"    '�˻�� �ڵ�
Private Const TEST_NM_LIS   As String = "LIS_NM"    '�˻�� �̸�
Private Const TEST_VALUES   As String = "VALUES"    '���

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const fs  As String = ""
Const Rs  As String = ""

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



Dim sStxCheck As Integer
Dim sEtxCheck As Integer
Dim sLfCheck  As Integer
Dim sCrcheck  As Integer
' --------------------------------------------------------------
Dim strOrdLst(100) As String

Dim ELEC1010(100)   As String
Dim fELEC1010       As Variant
Dim fELEC1010_1     As Variant
Dim fELEC1010_2     As Variant
Dim fELEC1010_3     As Variant
Dim SendData(10)     As String
Dim SendCount        As Integer
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
Private f_strOR_FLAG   As String

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

Dim fTBA40FR(50) As String
Dim fCellDynSize(50, 1) As Integer
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
    strTestCd(50) As String
End Type

Private f_typCode() As TYPE_CD

Dim RecordChk As Boolean

Dim strGumCd As String
Dim strJinCd As String
Dim fRcvString As String

Dim PatientID As String    'Q Message Pattern Check
Dim PatientSeq As String
Dim PatientDisk As String
Dim PatientRack As String
Dim PatientPos As String

Dim SeqNo As String
'Dim RecordChk   As Boolean

Dim G_CLVALU    As String
Dim G_CHVALU    As String
Dim G_EVALUATE  As String
Dim G_PANIC     As String
Dim G_DELTA     As String
Dim strFrameNo  As Integer
Dim OrderCnt As Integer
Dim vRow As Integer
Dim sPatiant_No As String


Private Type typeCobasE411
    TestDate      As String
    TestTime      As String
    RunType       As String 'N, E, R, C, S, B
    SampleNo      As String
    SID           As String 'Sid
    SampleTy      As String '1~5
    RackNo        As String
    Position      As String '1~5
    Priority      As String
    TestId(50)   As String
    Result(50)   As String
    Status(50)   As String
    Rerun(50)    As String
End Type

Dim CobasE411 As typeCobasE411
Dim fCobasE411(100) As String
Dim fCobasE411_1(100) As String

Dim OrderSort_Flag As Integer
Dim Patiant_Recevid As Boolean

Dim gspdResultRow  As Integer

'-- 2010.03.11 osw �߰� : �˻��� �˾��޼���
Private WithEvents mobjPopups   As PopUpMessages
Attribute mobjPopups.VB_VarHelpID = -1

Private mobjDefault             As PopUpMessage

'-- Interface Class
Private cInterface              As New clsInterface
Private objIntInfo              As clsIntInfo           '��ü���� Ŭ����
Private objOrder                As clsIntOrder          '�������� Ŭ����
Private objResult               As clsIntResults        '�������
Private objIntNm                As New clsIntTest       '�˻�����

Const SPCLEN As Integer = "11"

Dim SndCount As Integer
Dim sDeCnt   As Integer

Dim pDoCount1 As Integer
Dim fAdvia1650(100) As String
Dim fAdvia1650Temp(100) As String
Dim fAdvia1650Cfg(100) As Integer
Dim fAdvia1650Size(100, 1) As Integer

Dim fAdvia1200_CH(100) As String
Dim fAdvia1200_RS(100) As String


'-- 2010.03.11 osw �߰� : �˻��� �˾��޼���
Private Sub AddPopup(ByVal strSPnm As String, ByVal strSPid As String)
Dim objPopUp    As PopUpMessage
    
    Set objPopUp = New PopUpMessage
    With objPopUp
        .Caption = INS_NAME
        .Message = strSPnm & "(" & strSPid & ") ��" & vbCrLf & vbCrLf & " �˻��� ���ۼ���" & vbCrLf & ""
        .Clickable = False
        .Sticky = False
        Set .Background = imgBack.Item(0)
        Set .Logo = imgLogo.Item(0)
        .WavFile = App.Path & "\sounds\type.wav"
    End With
    mobjPopups.Show objPopUp
    
End Sub


'-- 2010.03.11 osw �߰� : �˻��� �˾��޼���
Private Sub Form_Unload(Cancel As Integer)
    Set mobjPopups = Nothing
    Set mobjDefault = Nothing
End Sub

'-- 2010.03.11 osw �߰� : �˻��� �˾��޼���
Private Sub SetupDefaultPopup()
    Set mobjDefault = New PopUpMessage
    With mobjDefault
        Set .Background = imgBack.Item(1)
        .ForeColor = vbWhite
        Set .Logo = imgLogo.Item(1)
        .WavFile = App.Path & "\newemail.wav"
        .Caption = "New Email"
        .Message = "You have received" & vbCrLf & "4 new emails." & vbCrLf & "Downloading..."
        .Clickable = True
        .ProgressBar = True
    End With
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
    
    '�˻��ڵ� ���̺�
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
            Call .Add(, TEST_CD_LIS, "�˻��ڵ�", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_NM_LIS, "�� �� ��", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_VALUES, "�˻���", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTA", "DELTA", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTAGBN", "DELTAGBN", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICL", "PANIC(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICH", "PANIC(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFL", "����ġ(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFH", "����ġ(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "AUTOVERIFY", "���", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, "REMARK", "��ü�ڵ�", (lvwCuData.Width - 310) * 0.1)
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
    Call lvw.ColumnHeaders.Add(, "EQP_ID", "��ü ��ȣ")
    
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
                '�÷� ���Ű�� ���˻� �ڵ��
                .Key = COL_KEY & Trim(adoRS.Fields("TESTCD_EQP") & "")
                '�÷����� �˻� �׸� �̸�
                .text = Trim(adoRS.Fields("TESTNM") & "")
                '�ױ״� �˻� �ڵ��
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
    Dim strTest     As String
    Dim strTestCd   As String
    Dim varTestCd   As Variant
    Dim tmpTestCd   As String
    Dim intCnt      As Integer
    Dim adoRS       As New Recordset
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
' �˻��׸� �ϰ� ó�� �ϱ� ���� ó��
        sqlDoc = ""
        sqlDoc = sqlDoc + vbLf + " SELECT TESTCD        "
        sqlDoc = sqlDoc + vbLf + "   FROM INTERFACE002  "
        sqlDoc = sqlDoc + vbLf + "  WHERE (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL)) "
        sqlDoc = sqlDoc + vbLf + "  ORDER BY OUT_SEQ, TESTCD_EQP"

        adoRS.CursorLocation = adUseClient
        adoRS.Open sqlDoc, AdoCn_Jet

        strTestCd = ""
        tmpTestCd = ""

        If adoRS.RecordCount > 0 Then
            adoRS.MoveFirst
            Do Until adoRS.EOF
                tmpTestCd = tmpTestCd & adoRS.Fields("TESTCD") & ""
                adoRS.MoveNext
            Loop
        End If
        
        adoRS.Close
        
        varTestCd = Split(tmpTestCd, ",")
        
        For intCnt = 0 To UBound(varTestCd) - 1
            strTestCd = strTestCd & "'" & varTestCd(intCnt) & "'" & ","
        Next
   
        Set AdoRs_SQL = New ADODB.Recordset
        
        If cboChk.ListIndex = 0 Then
            sqlDoc = ""
            sqlDoc = sqlDoc + vbLf + "SELECT PER_GUMJIN_DATE, PER_SSN, PER_NAME, CHARTNO, PER_GUM_NUM, SEQNO, meditem "
            sqlDoc = sqlDoc + vbLf + "  FROM ONIT..GUMJIN_INTERFACE                                                    "
            sqlDoc = sqlDoc + vbLf + " WHERE PER_GUMJIN_DATE BETWEEN '" & strDate & "' AND '" & strDate1 & "'          "
            sqlDoc = sqlDoc + vbLf + "   AND meditem IN (" & Mid(strTestCd, 1, Len(strTestCd) - 1) & ")                  "
            sqlDoc = sqlDoc + vbLf + "   AND STATUS = '0'                                                              "
            sqlDoc = sqlDoc + vbLf + "   AND RESULT = ''  OR RESULT IS NULL"
        Else
            sqlDoc = ""
                     sqlDoc = " SELECT a.EnterDate, b.Status, b.waitseqno, b.MAP2SEQNO, b.DispDesc, b.RVALUEKIND, b.NORMLOW, b.NORMHIGH, b.NORMALVALUE, b.RVALUEKIND , " & vbLf
            sqlDoc = sqlDoc & " a.ChartNo, b.GumsaKind, c.sujinname, b.status, c.PassNo " & vbLf
            sqlDoc = sqlDoc & "   FROM onit_out..WaitPrsnp a, onit_out..jun370_resulttb b, onit_out..pewprsnp c, onit_out..BAGMAP2PREF d " & vbLf
            sqlDoc = sqlDoc & "  WHERE a.EnterDate BETWEEN  '" & strDate & "' AND '" & strDate1 & "' " & vbLf
            sqlDoc = sqlDoc & "    AND a.WaitSeqNo = b.WaitSeqNo " & vbLf
            sqlDoc = sqlDoc & "    AND d.labno in (4) " & vbLf
            sqlDoc = sqlDoc & "    AND b.map2seqno=d.map2seqno " & vbLf
            sqlDoc = sqlDoc & "    AND b.Result= '' " & vbLf
            sqlDoc = sqlDoc & "    AND a.chartno=c.chartno " & vbLf
            'sqlDoc = sqlDoc & "    AND Substring(a.EnterDate,1,8) + b.rinputtime1 <= '" & strTime & "' " & vbLf
            sqlDoc = sqlDoc & "    AND (substring(a.sujinpart,1,2) <> '62') " & vbLf
            
            sqlDoc = sqlDoc & "  ORDER BY  a.EnterDate,a.entertime "
        End If
        
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

'-- ���ڵ��ȣ�� ȯ�������� �����´�.
Private Function f_subGet_PatInfo(ByVal strBarCd As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_PatInfo() As ADODB.Recordset"
    
        Set AdoRs_SQL = New ADODB.Recordset
        
        If cboChk.ListIndex = 0 Then
'            sqlDoc = ""
'            sqlDoc = "         Select a.PRSNVSDT as ��������,Rtrim(a.PRSNCODE) + lTrim(a.PRSNSUBC) as ó���ڵ� , b.ABHJNAME as �����ڸ� ,b.ABHJMRNO as íƮ��ȣ ,b.ABHJSCBT + b.ABHJSCID as �ֹε�Ϲ�ȣ , a.PRSNRSLT as �˻���, b.ABHJPSEX as ���� "
'            sqlDoc = sqlDoc & "  From MediEHE..PRSNUMBM a, MediEHE..ABHJMSTM b"
'            sqlDoc = sqlDoc & " Where a.PRSNVSDT between  '" & strDate & "' and '" & strDate1 & "'"
'            sqlDoc = sqlDoc & "   And a.PRSNRSLT = '' "
'            sqlDoc = sqlDoc & "   And a.PRSNCODE in('21AC','21AN','21AD','21AE','21AF','21AK','21AF1')"
'            sqlDoc = sqlDoc & "   And a.PRSNSUBC In('','001','002','003','004','005')"
'            sqlDoc = sqlDoc & "   And a.PRSNMRNO = b.ABHJMRNO"
'            sqlDoc = sqlDoc & " Order By a.PRSNMRNO"

        Else
'             sqlDoc = "         Select a.*, b.�����ڸ�,b.íƮ��ȣ,b.�ֹε�Ϲ�ȣ,  b.�ֹε�Ϲ�ȣ as ���� from TB_�˻��׸� a, TB_�������� b"
'            sqlDoc = sqlDoc & " Where a.�����+a.�����+a.������ between '" & strDate & "' and '" & strDate1 & "'"
'            sqlDoc = sqlDoc & "   And a.������������ < 5"
'            sqlDoc = sqlDoc & "   And a.������������ <> 5"
'            sqlDoc = sqlDoc & "   and ó���ڵ� in('C4802','C4812','C2243') "
'            sqlDoc = sqlDoc & "   and �����ڵ� in('','001','002','003','004','005','006','007') "
'            sqlDoc = sqlDoc & "   And a.íƮ��ȣ = b.íƮ��ȣ"
'            sqlDoc = sqlDoc & " Order By a.íƮ��ȣ"

        End If
        
        Set AdoRs_SQL = New ADODB.Recordset
        
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
        
        If AdoRs_SQL.RecordCount = 0 Then
            Set f_subGet_PatInfo = Nothing
            RecordChk = False
            Set AdoRs_SQL = Nothing
            Exit Function
        Else
            Set f_subGet_PatInfo = AdoRs_SQL
            RecordChk = True
        End If
    
        Set AdoRs_SQL = Nothing
    
Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    
    Call ErrMsgProc(CallForm)
    
End Function

Private Function f_subSet_WorkList_Barcode(ByVal strDate As String, Optional ByVal strPid As String, Optional ByVal strName As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim stryy, strmm, strdd
    
    Dim strTest     As String
    Dim strTestCd   As String
    Dim varTestCd   As Variant
    Dim tmpTestCd   As String
    Dim intCnt      As Integer
    Dim strDate1    As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
        strDate1 = Val(strDate)
        Set AdoRs_SQL = New ADODB.Recordset
        
        
        sqlDoc = ""
        sqlDoc = sqlDoc + vbLf + " SELECT TESTCD        "
        sqlDoc = sqlDoc + vbLf + "   FROM INTERFACE002  "
        sqlDoc = sqlDoc + vbLf + "  WHERE (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL)) "
        sqlDoc = sqlDoc + vbLf + "  ORDER BY OUT_SEQ, TESTCD_EQP"

        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_Jet

        strTestCd = ""
        tmpTestCd = ""

        If AdoRs_SQL.RecordCount > 0 Then
            AdoRs_SQL.MoveFirst
            Do Until AdoRs_SQL.EOF
                tmpTestCd = tmpTestCd & AdoRs_SQL.Fields("TESTCD") & ""
                AdoRs_SQL.MoveNext
            Loop
        End If
        
        AdoRs_SQL.Close
        
        varTestCd = Split(tmpTestCd, ",")
        
        For intCnt = 0 To UBound(varTestCd) - 1
            strTestCd = strTestCd & "'" & varTestCd(intCnt) & "'" & ","
        Next
        
        
        
        
        If Len(strDate1) > 6 Then
            sqlDoc = ""
            sqlDoc = sqlDoc + vbLf + "SELECT PER_GUMJIN_DATE, PER_SSN, PER_NAME, CHARTNO, PER_GUM_NUM, SEQNO, meditem, INTERFACECODE "
            sqlDoc = sqlDoc + vbLf + "  FROM ONIT..GUMJIN_INTERFACE                                                    "
            sqlDoc = sqlDoc + vbLf + " WHERE PER_GUMJIN_DATE = '" & Mid(strDate1, 1, 8) & "' "
            sqlDoc = sqlDoc + vbLf + "   AND PER_GUM_NUM = '" & Mid(strDate1, 9) & "' "
            sqlDoc = sqlDoc + vbLf + "   AND INTERFACECODE IN (" & Mid(strTestCd, 1, Len(strTestCd) - 1) & ")                  "
            sqlDoc = sqlDoc + vbLf + "   AND STATUS = '0'                                                              "
            sqlDoc = sqlDoc + vbLf + "   AND RESULT = ''  OR RESULT IS NULL"
            
        Else
            sqlDoc = ""
            sqlDoc = " SELECT a.EnterDate, b.Status, b.waitseqno, b.MAP2SEQNO, b.DispDesc, b.RVALUEKIND, b.NORMLOW, b.NORMHIGH, b.NORMALVALUE, b.RVALUEKIND , " & _
                     " a.ChartNo, b.GumsaKind, c.sujinname, b.status " & _
                     " FROM onit_out..WaitPrsnp a, onit_out..jun370_resulttb b, onit_out..pewprsnp c, onit_out..BAGMAP2PREF d " & _
                     " WHERE a.WaitSeqNo = '" & Val(strDate) & "' " & _
                     " AND a.WaitSeqNo = b.WaitSeqNo " & _
                     " AND d.labno in (4) " & _
                     " AND B.Result = '' " & _
                     " AND b.map2seqno = d.map2seqno " & _
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
    
    Dim intPos1 As Integer
    
    Dim mIntNms As clsIntTest


'On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = ""
    
    intCol = 10
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
            itemX.text = Trim(adoRS.Fields("TESTCD") & "")
        Set itemX = Nothing
        
        '-------�߰�-------------------------------------
'        With objIntNm
'            .TestCd = Trim(adoRS.Fields("TESTCD") & "")
'            .TestNm = Trim(adoRS.Fields("TESTNM") & "")
'            .McTestCd = Trim(adoRS.Fields("TEST_EQP") & "")
'            .McTestNm = Trim(adoRS.Fields("TESTNM") & "")
'            .FrVal = Trim(adoRS.Fields("REFL") & "")
'            .ToVal = Trim(adoRS.Fields("REFH") & "")
'        End With
        
        Dim strTestKey, strTestData As String
        Dim varTestCd   As Variant
        Dim intTstCnt   As Integer
        Dim strItemData As String
        
'        Set objIntNm = New clsIntTest
    '    Set objIntNm = Nothing
        
        strTestKey = Trim(adoRS.Fields("TEST_EQP")) & ""
                
        varTestCd = Split(Trim(adoRS.Fields("TESTCD")), ",")
        
        For intTstCnt = 0 To UBound(varTestCd)
            If varTestCd(intTstCnt) = "" Then Exit For
            strTestData = varTestCd(intTstCnt)
        
            If objIntNm.Exists(strTestKey) = False Then
                objIntNm.AddNew strTestKey, strTestData
            End If
        Next intTstCnt
        
'        strTestData = Trim(adoRS.Fields("TESTCD")) & "|" & Trim(adoRS.Fields("TESTNM") & "") & "|" & Trim(adoRS.Fields("TEST_EQP")) & "|" '& Trim(adoRS.Fields("TESTNM") & "") & "|"
        
'                    MsgBox "�����Ͱ� �ߺ��Դϴ�.", vbCritical, "����Ȯ��"
'                    GoTo ErrMsg
        'objIntNm.GetString
        strItemData = objIntNm.GetIntNm(strTestKey)
        '-------�߰�-------------------------------------
        
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
        
        fChannel(intCol - 9) = adoRS.Fields("TEST_EQP")
        
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
            If Trim$(strOrdcd) = Trim$(f_typCode(intIdx1).strTestCd(intIdx2)) Then
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
        sqlDoc = sqlDoc & " WHERE A.Per_com_Code = B.COM_CODE " & vbCr
        sqlDoc = sqlDoc & "   AND A.per_gumjin_date BETWEEN '" & Format(dtpStartDt.Value, "yyyymmdd") & "' AND '" & Format(dtpStopDt.Value, "yyyymmdd") & "'" & vbCr
        sqlDoc = sqlDoc & "  AND SUBSTRING(C.KIND, 1, 1) = 'C' " & vbCr
        sqlDoc = sqlDoc & "   AND A.EDPSCODE = C.MEDITEM " & vbCr
        sqlDoc = sqlDoc & " GROUP BY B.COM_CODE, B.COM_NAME " & vbCr
        
        Set AdoRs_SQL = New ADODB.Recordset
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
        
        If AdoRs_SQL.RecordCount > 0 Then
            AdoRs_SQL.MoveFirst
            cboComNm.Clear
            cboComNm.AddItem "��ü"
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
    If Trim(cboChk.text) = "����" Then
'        cboComNm.Visible = True
'        mskOrdtime.Visible = False
'        Label10.Visible = False
'        Call f_subSet_ComList
    Else
        cboComNm.Visible = False
'        mskOrdtime.Visible = True
'        Label10.Visible = True
        cboComNm.Clear
    End If
End Sub

Private Sub cboComNm_DropDown()
        Call f_subSet_ComList
End Sub

Private Sub cmdAppend_Click(Index As Integer)
   
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String

    Dim varTmp  As Variant, strErrMsg   As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval()  As String, intCnt       As Integer
    Dim strTmp1()       As String, strTmp2()    As String
    Dim intPos          As String, strTestCd    As String, strTestRst   As String
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
    Dim strBarno As String
    Dim strSPnm As String
    Dim strSPid As String
    Dim strChartNo As String
    Dim strEqpCd As String
    Dim valEqpcd As Variant

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

            .GetText 2, intRow, varTmp:    strDate = Trim$(varTmp)
            .GetText 3, intRow, varTmp:    strBarno = Trim$(varTmp)
            .GetText 4, intRow, varTmp:    strSPnm = Trim$(varTmp)
            .GetText 6, intRow, varTmp:    strChartNo = Trim$(varTmp)
            .GetText 7, intRow, varTmp:    strSPid = Trim$(varTmp)

            .GetText 1, intRow, varTmp
            
'            strDate = Mid(strDate, 1, 4) & Mid(strDate, 6, 2) & Mid(strDate, 9, 2)

            If strSPid = "" Then Exit For

            intCnt = 0: Erase strOrdcd: Erase strRstval
            
            If Trim$(varTmp) = "1" Then
                For intCol = 8 To .MaxCols
                    .GetText intCol, intRow, varTmp
                        If Trim$(varTmp) <> "" Then
                            .GetText intCol, 0, varTmp
                            Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                            If Not itemX Is Nothing Then
                                .GetText intCol, intRow, varTmp
                                strTestCd = itemX.ListSubItems(1)
                                intPos = InStr(strTestCd, ",")
                                strEqpCd = ""
                
                                blnFlag = False
                                
                                If cboChk.ListIndex = 0 Then
                                    Set mAdoRs = f_subSet_WorkList_Barcode(strDate, strSPid)
                                Else
                                    Set mAdoRs = f_subSet_WorkList_Barcode(strDate, strBarno)
                                End If
                                
                                If RecordChk = True Then
                                    
                                   strEqpCd = ""

                                    Do Until mAdoRs.EOF
                                        If cboChk.ListIndex = 0 Then
                                            If InStr(itemX.text, Trim(mAdoRs.Fields("meditem") & ",")) > 0 Then
                                                strEqpCd = Trim(mAdoRs.Fields("meditem"))
                                                Exit Do
                                            End If
                                        Else
                                            If InStr(itemX.text, Trim(mAdoRs.Fields("MAP2SEQNO") & ",")) > 0 Then
                                                strEqpCd = Trim(mAdoRs.Fields("MAP2SEQNO"))
                                                Exit Do
                                            End If
                                        End If
                                        mAdoRs.MoveNext
                                    Loop
                                    
                                    If strEqpCd <> "" Then
                                        Dim stryy, strmm, strdd, tmpDate, strEMRID As String
                                        Dim tmpREF As String
                                        
                                        If cboChk.ListIndex = 0 Then
                                            strEqpCd = Replace(strEqpCd, ",", "")
                                            sqlDoc = ""
                                            sqlDoc = sqlDoc + "UPDATE GUMJIN_INTERFACE"
                                            sqlDoc = sqlDoc + "   SET RESULT = '" & Trim$(varTmp) & "',"
                                           ' sqlDoc = sqlDoc + "       ACT_TEST_DATE = '" & Format(Now, "yyyymmdd") & "',"
                                            sqlDoc = sqlDoc + "       STATUS = '1'"
                                            sqlDoc = sqlDoc + " WHERE PER_GUMJIN_DATE = '" & strDate & "'"
                                            sqlDoc = sqlDoc + "   AND PER_SSN = '" & strSPid & "'"
                                            sqlDoc = sqlDoc + "   AND MEDITEM = '" & strEqpCd & "'"
                                            
                                        Else
                                            sqlDoc = "Update onit_out..Jun370_resulttb" _
                                                    & "   Set Result = '" & varTmp & "' " _
                                                    & " Where WaitSeqNo = '" & strBarno & "'" _
                                                    & "   and Map2seqno = '" & strEqpCd & "'"
                                                    
                                        Debug.Print sqlDoc
                                        End If

                                        AdoCn_SQL.Execute sqlDoc
    
                                        lblStatus.Caption = "���� ����!!"
                                        
                                        Set adoRS = Nothing:    mAdoRs.Close
                                        
                                        spdResult1.Row = intRow
                                        spdResult1.Col = 2: spdResult1.BackColor = vbCyan
                                        spdResult1.Col = 3: spdResult1.BackColor = vbCyan
                                        spdResult1.Col = 4: spdResult1.BackColor = vbCyan
                                        spdResult1.Col = 5: spdResult1.BackColor = vbCyan
                                        spdResult1.Col = 6: spdResult1.BackColor = vbCyan
                                        spdResult1.Col = 7: spdResult1.BackColor = vbCyan
                                        spdResult1.Col = 1: spdResult1.Value = 0
                                        
'                                        If strErrMsg = "" Then
'                                            sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
'                                                     " where SPCNO   = '" & strSPid & "'" & _
'                                                     "   and TRANSDT = '" & Format(Now, "yyyymmdd") & "'"
'                                            AdoCn_Jet.Execute sqlDoc
'                                        Else
'                                            MsgBox strErrMsg, vbInformation, App.Title
'                                        End If
                                    End If  ' strEqpCd <> ""
                                End If ' RecordChk =  true
                            Set itemX = Nothing
                        End If ' Not itemX
                    End If ' Trim$(varTmp) <> ""
                Next ' intCol
            End If ' Trim$(varTmp) = "1"
        Next ' intRow
    End With
    Me.MousePointer = 0
    MsgBox "�� SERVER�� ����� Upload �Ϸ�Ǿ����ϴ�. ��      " & vbCrLf & vbCrLf & "     OCS/EMR �����ȸ ȭ�鿡�� ����� Ȯ�� �Ͻʽÿ�..  ", vbInformation, App.Title

    Exit Sub
ErrorRoutine:

    Set AdoRs_SQL = Nothing

    Set itemX = Nothing

    Me.MousePointer = 0
    Call ErrMsgProc(CallForm)
End Sub

Private Sub cmdEot_Click()
    Call COM_OUTPUT(EOT)
End Sub

Private Sub cmdExcel_Click()
'Dim sRow As Integer, sCol As Integer, sCnt As Integer
'Dim sSave As Boolean
'Dim fName As String
'
'    If chkExcel.Value = 1 Then
'        With CommonDialog1
'             .FileName = App.Path & "\" & fName & ".xls"
'             .DialogTitle = "Save As New Excel Spread"
'             .FileName = REG_INSNAME & "  " & Format(mskRstDate, "####-##-##") & " �˻���Ȳ����"
'             .Filter = "New Excel file(*.xls)"
'             .ShowSave
'            sSave = spdResult2.ExportToExcel(.FileName, Format(mskRstDate, "####-##-##") & " TBA20FR", "\log.txt")
'        End With
'    Else
'        Call gsp_SetSpdTExcelExport(spdResult2, True)
'    End If

    Dim strTmp As String
    Dim lngRows As Long
    
    If spdResult2.DataRowCnt = 0 And spdResult2.DataRowCnt = 0 Then Exit Sub
    
    With spdResult2
        .Row = 0: .Row2 = .maxrows
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .maxrows
    End With
 
    With tblexcel
        .maxrows = spdResult2.maxrows + 1
        .MaxCols = spdResult2.MaxCols
        .Row = 1: .Row2 = .maxrows
        .Col = 1: .Col2 = spdResult2.MaxCols
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    CommonDialog1.InitDir = "C:\"
    CommonDialog1.Filter = "ExCelFile(*.XLS)|*.XLS"
    CommonDialog1.FileName = REG_INSNAME & "  " & Format(dtpRsltDay, "####-##-##") & " �˻���Ȳ����"
    CommonDialog1.ShowSave

    tblexcel.SaveTabFile (CommonDialog1.FileName)

End Sub

Private Sub cmdOrder_Click()
Dim ii As Integer
Dim chkRackNo As Variant
Dim strMsg As String

    spdResult1.GetText 7, 1, chkRackNo
    strMsg = "�������� �غ� �Ǿ����ϴ�." & vbCrLf & chkRackNo & "�� ����Ͻðڽ��ϱ�?"
    If vbYes = MsgBox(strMsg, vbCritical + vbYesNo) Then
        OrderCnt = 0
        comEQP.Output = ENQ
        tmrOrder.Enabled = False
        
        With spdResult1
            For ii = 1 To .maxrows
                .Col = 1: .Row = ii
                If .Value = 1 Then
                    .Col = 2
                    If Len(Trim(.text)) > 0 And .BackColor <> vbCyan Then
                        comEQP.Output = ENQ
'                        Debug.Print "[HOST] " & ENQ
                        SendCount = 0
                        txtResult.text = txtResult.text + "[HOST] " & ENQ
                        lblStatus = "����������.."
                        .BackColor = vbCyan
                        OrderCnt = ii
                        .Col = 3
                        .BackColor = vbCyan
                        .Col = 4
                        .BackColor = vbCyan
                        .Col = 5
                        .BackColor = vbCyan
                        .Col = 6
                        .BackColor = vbCyan
                        Exit For
                    End If
                End If
            Next
        End With
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdPosNo_Click()
'Dim sNo As String, sCnt As Integer, sAdd As Integer
'
'AgainInput:
'    sNo = InputBox("���� ��ȣ�� �Է��ϼ��� !")
'    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
'        If Not IsNumeric(sNo) Then
'            MsgBox "���ڸ� �Է��ϼ���.!", vbCritical
'            GoTo AgainInput
'        End If
'
'        With spdResult1
'            sAdd = 0
'            For sCnt = .ActiveRow To .maxrows
'                .Row = sCnt
'                .Col = 7:       .Text = Trim(sAdd + Val(sNo))
'                sAdd = sAdd + 1
'            Next sCnt
'        End With
'    End If
Dim sNo As String, sCnt As Integer, sAdd As Integer

AgainInput:
    sNo = InputBox("���� ��ȣ�� �Է��ϼ��� !")
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
        If Not IsNumeric(sNo) Then
            MsgBox "���ڸ� �Է��ϼ���.!", vbCritical
            GoTo AgainInput
        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .maxrows
                .Row = sCnt
                .Col = 7:       .text = Trim(sAdd + Val(sNo))
                If Trim(sAdd + Val(sNo)) = 14 Then sNo = 0
                sAdd = sAdd + 1
            Next sCnt
        
            .StartingRowNumber = Val(sNo)
        End With
    End If
End Sub

Private Sub cmdPrint_Click()
Dim objclsCommon As New clsCommon

Dim Tmp_Testnm As String
Dim Row_cnt As Integer, Col_cnt As Integer, TmpPrintline As Integer
Dim vTmp As Variant
Dim stragesex As String

Const TmpLine = "������������������������������������������������������������������������������������������������������������������������"

    If spdResult1.maxrows >= 1 Then
        With objclsCommon
            .PrintText 15, 3, Format(Date, "yyyy/mm/dd") & "  WorkList Report..( " & App.EXEName & " )", "Arial", 12
            
            .PrintText 0.5, 5, TmpLine
            .PrintText 0.5, 6, "��", , 9
            .PrintText 2, 6, "ó������", , 9
            .PrintText 7, 6, "ȯ�ڼ���", , 9
            .PrintText 12, 6, "���Ϲ�ȣ", , 9
            .PrintText 16, 6, "���˻�����", , 9
            .PrintText 0.5, 7, TmpLine
            
            TmpPrintline = 8
        
        For Row_cnt = 1 To spdResult1.maxrows
            spdResult1.Row = Row_cnt
            
            If (Row_cnt Mod 34) <> 0 Then
                                    .PrintText 0.5, TmpPrintline, Row_cnt, , 9                          ' ��
                spdResult1.Col = 2: .PrintText 2, TmpPrintline, Mid(spdResult1.text, 3), , 9                    ' ó������
                spdResult1.Col = 4: .PrintText 7, TmpPrintline, Trim(spdResult1.text), 9              ' ��ü��ȣ
                spdResult1.Col = 6: .PrintText 12, TmpPrintline, Trim(spdResult1.text), , 9             ' ��    ��
               ' spdResult1.Col = 2: .PrintText 16, TmpPrintline, Trim(spdResult1.text), , 9             ' ������
                
                
                For Col_cnt = 8 To spdResult1.MaxCols
            
                    spdResult1.Row = Row_cnt:            spdResult1.Col = Col_cnt
                    
                    If spdResult1.BackColor = &HC6FEFF Then
                        spdResult1.GetText Col_cnt, 0, vTmp
                        Tmp_Testnm = Tmp_Testnm & ", " & vTmp
                    End If
                    
                Next Col_cnt
                
                spdResult1.Col = 5: .PrintText 16, TmpPrintline, Mid(Trim(Tmp_Testnm), 2), , 7.5
                
                TmpPrintline = TmpPrintline + 2
                Tmp_Testnm = ""
            Else
            
                '-------------------------------------------------------
            
                                    .PrintText 0.5, TmpPrintline, Row_cnt, , 9                          ' ��
                spdResult1.Col = 2: .PrintText 2, TmpPrintline, Mid(spdResult1.text, 3), , 9                   ' ó������
                spdResult1.Col = 4: .PrintText 6, TmpPrintline, Trim(spdResult1.text), 9              ' ��ü��ȣ
                spdResult1.Col = 6: .PrintText 12, TmpPrintline, Trim(spdResult1.text), , 9             ' ��    ��
                
                
                For Col_cnt = 8 To spdResult1.MaxCols
            
                    spdResult1.Row = Row_cnt:            spdResult1.Col = Col_cnt
                    
                    If Trim(spdResult1.text) <> "" Then
                        spdResult1.GetText Col_cnt, 0, vTmp
                        Tmp_Testnm = Tmp_Testnm & ", " & vTmp
                    End If
                    
                Next Col_cnt
                
                spdResult1.Col = 5: .PrintText 16, TmpPrintline, Mid(Trim(Tmp_Testnm), 2), , 7.5
                
                TmpPrintline = TmpPrintline + 2
                Tmp_Testnm = ""
                
                '-------------------------------------------------------
            
                    .PrintText 0.5, TmpPrintline, TmpLine
                    .PrintText 1, TmpPrintline + 1, "���� Next Report ����", , 9, True
                    Printer.NewPage
                    
                    .PrintText 0.5, 5, TmpLine
                    .PrintText 0.5, 6, "��", , 9
                    .PrintText 2, 6, "������ȣ", , 9
                    .PrintText 6, 6, "ȯ�ڼ���", , 9
                    .PrintText 12, 6, "���Ϲ�ȣ", , 9
                    .PrintText 16, 6, "ó������", , 9
                    .PrintText 20, 6, "���˻�����", , 9
                    .PrintText 0.5, 7, TmpLine
                    
                    TmpPrintline = 9
            End If
        
        Next Row_cnt
        .PrintText 0.5, TmpPrintline, TmpLine
        .PrintText 1, TmpPrintline + 1, "���� End of Report ����", , 9, True
        
        End With
        Printer.NewPage
        Printer.EndDoc
        
        MsgBox Format(Date, "yyyy/mm/dd") & "������ " & App.EXEName & "�� ��� �˻� WorkList�� Print�Ǿ����ϴ�..       " & vbCrLf & vbCrLf & "���� �۾��� �����Ͻʽÿ�..", vbInformation + vbOKOnly, App.Title
    Else
        MsgBox Format(Date, "yyyy/mm/dd") & "������ " & App.EXEName & "�� ��� �˻� WorkList��  Load �Ǿ� ���� �ʽ��ϴ�..       " & vbCrLf & vbCrLf & "�ڷḦ Ȯ�� �Ͻʽÿ�..", vbInformation + vbOKOnly, App.Title
    End If
    
    '
    ' ������ ����
    '
    spdResult1.SaveTabFile App.Path & "\" & REG_INSNAME & "_Request.txt"
    

End Sub

Private Sub cmdRackNo_Click()
'    Dim sNo As String, sCnt As Integer, sAdd As Integer
'    Dim aROW    As Integer, aCOL   As Integer
'    Dim varChk  As Variant, varBar As Variant, varNum As Variant
'    Dim iRow    As Integer, iCnt   As Integer
'    Dim strRack_tmp As String
'
'
'AgainInput:
'    sNo = InputBox("���� ��ȣ�� �Է��ϼ��� !")
'    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
'        If Not IsNumeric(sNo) Then
'            MsgBox "���ڸ� �Է��ϼ���.!", vbCritical
'            GoTo AgainInput
'        End If
'
'        With spdResult1
'            iCnt = 1
'            .GetText 1, 1, varChk
'            .GetText 2, 1, varBar
'            varNum = sNo
'            If Trim(varChk) = "1" And Trim(varBar) <> "" Then
'                For iRow = 1 To .maxrows
'                    .SetText 6, iRow, varNum
'                    .SetText 7, iRow, ((iCnt Mod 101) + 1) - 1
'                    iCnt = iCnt + 1
'                    If (iCnt Mod 101) = 1 Then varNum = varNum + 1
'                Next
'            End If
'        End With
'    End If

Dim sNo As String, sCnt As Integer, sAdd As Integer
Dim fNum1 As Integer, fNum2 As Integer
Dim intRow1 As Integer

AgainInput:
    fNum1 = 1: fNum2 = 0
    sNo = InputBox("���� ����ȣ�� �Է��ϼ���!")
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
'        sNo = UCase(sNo)
        
'        If Asc(sNo) < 65 Or Asc(sNo) > 70 Then
'            MsgBox "a~f������ ���ڸ� �Է��ϼ���.!", vbCritical
'            GoTo AgainInput
'        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .maxrows
                intRow1 = intRow1 + 1
                .Row = sCnt
                .Col = 1
                If .Value >= 1 Then
                    .Col = 6
                    If intRow1 = (14 * fNum1) + 1 Then
                        fNum1 = fNum1 + 1: fNum2 = 0
                    End If
                    fNum2 = fNum2 + 1
                    .text = Chr(fNum1 + Asc(sNo) - 1)
                    
                    .Col = 7
                    .text = fNum2
                End If
            Next sCnt
        End With
    End If

End Sub

Private Sub cmdRequist_Click(Index As Integer)
    Dim ret As Integer
    
    ret = spdResult1.LoadFromFile(App.Path & "\" & REG_INSNAME & "_Request.txt")

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
    
    If cboChk.text = "" Then
        MsgBox " �˻������� �����ϼ���.", vbOKOnly + vbInformation, App.Title
        Exit Sub
    End If

On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
        Set AdoRs_ORACLE = New ADODB.Recordset
       
    '-- WorkList��ȸ
    Dim strTime As String
    
    strTime = dtpStopDt.Value
    Set mAdoRs = f_subSet_WorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"), strTime)
    
    If RecordChk = False Then
        MsgBox dtpStartDt.Value & "�� ����  " & dtpStopDt.Value & "�ϱ����� �˻� ����ڰ� �����ϴ�.", vbOKOnly + vbInformation, App.Title
        Exit Sub
    Else
        strBarno = ""
        mAdoRs.MoveFirst

        With spdWorklist
            If cboChk.ListIndex = 0 Then
                For intCnt = 0 To mAdoRs.RecordCount - 1
                    If strBarno <> mAdoRs.Fields("PER_SSN") Then
                        optBar.Value = True
                        pGrid_Point = SeqSearch(spdWorklist, mAdoRs.Fields("PER_SSN"), 5)
    
                        If pGrid_Point = 0 Then
                            pGrid_Point = SeqNullSearch(spdWorklist, mAdoRs.Fields("PER_SSN"), 5)
                            If pGrid_Point = 0 Then .maxrows = .maxrows + 1: pGrid_Point = .maxrows
                        End If
    
                        .SetText 1, pGrid_Point, "0"
                        .SetText 2, pGrid_Point, mAdoRs("PER_GUMJIN_DATE")
                        .SetText 4, pGrid_Point, mAdoRs("PER_NAME")
                        .SetText 5, pGrid_Point, mAdoRs("CHARTNO")
                        .SetText 6, pGrid_Point, mAdoRs("PER_SSN")
                        
                        Dim mSex As String
                        
'                        If Mid(mAdoRs("PER_SSN") & "", 8, 1) = " " Then
'                            mSex = ""
'                        Else
'                            mSex = Mid(mAdoRs("PER_SSN") & "", 7, 1)
'                            Select Case mSex
'                                Case 1, 3
'                                    .SetText 6, pGrid_Point, "M"
'                                Case 2, 4
'                                    .SetText 6, pGrid_Point, "F"
'                            End Select
'                        End If
                        
                        .Row = pGrid_Point: .Col = 1: .ForeColor = HNC_Black
                                            .Col = 2: .ForeColor = HNC_Black
                                            .Col = 4: .ForeColor = HNC_Black
                                            .Col = 5: .ForeColor = HNC_Black
                                            .Col = 6: .ForeColor = HNC_Black
                                            
                        
                        If blt = False Then
                            .Row = pGrid_Point - 1
                            .Action = ActionDeleteRow
                            .maxrows = .maxrows - 1
                        Else
                            blt = False
                        End If
                    End If
    
                    strEqpCd = f_funGet_CODE(Trim(mAdoRs.Fields("meditem")))
                    
                    Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                    If Not itemX Is Nothing Then
                        spdWorklist.SetText 1, pGrid_Point, "0"
                        spdWorklist.Col = itemX.Index + 7
                        spdWorklist.Row = pGrid_Point
                        spdWorklist.BackColor = &HC6FEFF   '&HC6FEFF
                        blt = True
                    End If
                    strBarno = mAdoRs.Fields("PER_SSN")
                    mAdoRs.MoveNext
                Next
            Else
                For intCnt = 0 To mAdoRs.RecordCount - 1
                    If strBarno <> mAdoRs.Fields("ENTERDATE") & Format(mAdoRs("WAITSEQNO"), "0000") & "" Then
                        pGrid_Point = SeqSearch(spdWorklist, mAdoRs.Fields("ENTERDATE") & Format(mAdoRs("WAITSEQNO"), "0000"), 3)
            
                        If pGrid_Point = 0 Then
                            pGrid_Point = SeqNullSearch(spdWorklist, mAdoRs.Fields("ENTERDATE") & Format(mAdoRs("WAITSEQNO"), "0000") & "", 3)
                            If pGrid_Point = 0 Then .maxrows = .maxrows + 1: pGrid_Point = .maxrows
                        End If
                        .SetText 1, pGrid_Point, "0"
                        .SetText 2, pGrid_Point, mAdoRs("ENTERDATE")
                        .SetText 3, pGrid_Point, mAdoRs("WAITSEQNO")
                        .SetText 4, pGrid_Point, Trim(mAdoRs("SUJINNAME"))
                        .SetText 5, pGrid_Point, mAdoRs("Chartno")
                        .SetText 6, pGrid_Point, mAdoRs("PassNo")
                        
                        
'                        If Mid(mAdoRs("PassNo") & "", 8, 1) = " " Then
'                            mSex = ""
'                        Else
'                            mSex = Mid(mAdoRs("PassNo") & "", 7, 1)
'                            Select Case mSex
'                                Case 1, 3
'                                    .SetText 6, pGrid_Point, "M"
'                                Case 2, 4
'                                    .SetText 6, pGrid_Point, "F"
'                            End Select
'                        End If
                        If blt = False Then
                            .Row = pGrid_Point - 1
                            .Action = ActionDeleteRow
                            .maxrows = .maxrows - 1
                        Else
                            blt = False
                        End If
                        
                    End If
                    
                    strEqpCd = f_funGet_CODE(Trim(mAdoRs.Fields("MAP2SEQNO")) & "")
                    Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                    If Not itemX Is Nothing Then
                        spdWorklist.SetText 1, pGrid_Point, "0"
                        spdWorklist.Col = itemX.Index + 7
                        spdWorklist.Row = pGrid_Point
                        spdWorklist.BackColor = &HC6FEFF   '&HC6FEFF
                        blt = True
                    End If
                    strBarno = mAdoRs.Fields("ENTERDATE") & Format(mAdoRs("WAITSEQNO"), "0000")
                    mAdoRs.MoveNext
                Next
            End If
            If blt = False Then
                .Row = pGrid_Point
                .Action = ActionDeleteRow
                .maxrows = .maxrows - 1
            End If
        End With
    End If
    
    Set AdoRs_SQL = Nothing
    spdWorklist.Row = 1
    spdWorklist.Col = 1
    spdWorklist.Action = ActionActiveCell
    
    Dim arow    As Integer, aCOL   As Integer
    Dim varChk  As Variant, varBar As Variant, varNum As Variant
    Dim iRow    As Integer, iCnt   As Integer
    Dim strRack_tmp As String
        
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
    
    optSeq.Value = True
    
    txtChart.ForeColor = &HFFC0C0
    txtChart.text = "��Ʈ��ȣ �Է�"
    
    Rem txtChart.SetFocus
    
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
    List1.Clear
    txtChart.text = ""
    
    With spdWorklist

        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
        .maxrows = 1
        
    End With
    
    With spdResult1
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
        .RowHeight(-1) = 13
    End With

    With spdResult2
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    Dim Rowcnt As Integer
    Dim Colcnt As Integer

    With spdRstview
        For Rowcnt = 1 To 8
            For Colcnt = 2 To 6 Step 2
                .Row = Rowcnt
                .Col = Colcnt
                .BackColor = &HFFFFFF
                .text = ""
            Next Colcnt
        Next Rowcnt
    End With

    SndCount = 0
    sDeCnt = 0
    Erase strOrdLst

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
        Call ShowMessage("���� �Ǿ����ϴ�.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "�۾���.."
    Else
        Call ShowMessage("���� ���� �ʾҽ��ϴ�.")
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "�۾� �����.."
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
        Call ShowMessage("���� ���� �ʾҽ��ϴ�.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "�۾���.."
    Else
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "�۾� �����.."
    End If
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
End Sub

'Private Sub cmdAppend_Click(Index As Integer)
'
'    Dim adoRS   As New ADODB.Recordset
'    Dim sqlDoc  As String
'
'    Dim varTmp  As Variant, strErrMsg   As String
'    Dim strSampleno()   As String, strBarno     As String, strTime      As String
'    Dim strOrdcd()      As String, strRstval()  As String, intCnt       As Integer
'    Dim strTmp1()       As String, strTmp2()    As String
'    Dim intPos          As String, strTestcd    As String, strTestRst   As String
'    Dim strTestnm       As String
'    Dim strRef          As String
'    Dim strUnit         As String
'    Dim strOrdLst()     As String, strPid()    As String, strPnm() As String
'
'    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
'    Dim itemX   As ListItem
'    Dim objSpd  As vaSpread
'    Dim sqlRet  As Integer
'    Dim flgSave As Boolean
'    Dim SaveGbn As Integer
'    Dim strDate As String
'
'    CallForm = "frmComm - Private Sub cmdAppend_Click()"
'
'On Error GoTo ErrorRoutine
'
'    Me.MousePointer = 11
'
'    If Index = 0 Then
'        Set objSpd = spdResult1
'    Else
'        Set objSpd = spdResult2
'    End If
'
'    With objSpd
'        For intRow = 1 To .maxrows
'
''            .GetText 2, intRow, varTmp:         strDate = Trim$(varTmp)
''            .GetText 3, intRow, varTmp:         strBarno = Trim$(varTmp)
''            .GetText .MaxCols, intRow, varTmp:  strTime = Trim$(varTmp)
'
'            .GetText 2, pGrid_Point, varTmp:   strDate = Trim$(varTmp)
'            .GetText 3, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)
'            .GetText 4, pGrid_Point, varTmp:   pName = Trim$(varTmp)
'            .GetText 5, pGrid_Point, varTmp:   pNo = Trim$(varTmp)
'
'
'            .GetText 1, intRow, varTmp
'
'            If strBarno = "" Then Exit For
'
'            intCnt = 0: Erase strOrdcd: Erase strRstval
'            If Trim$(varTmp) = "1" Then
'                For intCol = 6 To .MaxCols
'                    .GetText intCol, intRow, varTmp
'                    If Trim$(varTmp) <> "" Then
'                        .GetText intCol, 0, varTmp
'                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
'                        If Not itemX Is Nothing Then
'                            .GetText intCol, intRow, varTmp
'                            strTestcd = itemX.ListSubItems(1)
'                            intPos = InStr(strTestcd, ",")
'                            If intPos > 0 Then
'                                Do While intPos > 0
'
'                                    blnFlag = False
'                                    Set mAdoRs = f_subSet_TestList(strBarno)
'                                    Do Until mAdoRs.EOF
'                                        If mAdoRs("LCODE") = Mid$(strTestcd, 1, intPos - 1) Then blnFlag = True: Exit Do
'                                        mAdoRs.MoveNext
'                                    Loop
'                                    strTestcd = Mid$(strTestcd, intPos + 1)
'                                    intPos = InStr(strTestcd, ",")
'
'                                    AdoCn_ORACLE.BeginTrans
'
'                                    sqlDoc = "insert into lab_result" & _
'                                             "   (RESULTNO, LCODE, LSEQ, LNAME, LRESULT, UNIT, REFV,LTYPE,RESULT_DATE, REPORTER, PT_SEQ) " & _
'                                             " values('" & mAdoRs("resultno") & "', '" & strTestcd & "'," & _
'                                             "       '1', '" & Trim(itemX.ListSubItems(2)) & "'," & _
'                                             "       '" & Trim(varTmp) & "', '" & Trim(itemX.ListSubItems(9)) & "'," & _
'                                             "       '" & Trim(itemX.ListSubItems(8)) & "','0',sysdate," & _
'                                             "       '70001','" & mAdoRs("pt_seq") & "')"
'                                    AdoCn_ORACLE.Execute sqlDoc
'
'                                    sqlDoc = ""
'                                    sqlDoc = sqlDoc & "update ipd_order_date set req_result2 = '*'"
'                                    sqlDoc = sqlDoc & " where patient_no = '" & mAdoRs("patient_no") & "'"
'                                    sqlDoc = sqlDoc & "   and order_date = to_date('" & strDate & "','yyyy-mm-dd')"
'                                    AdoCn_ORACLE.Execute sqlDoc
'
'                                    sqlDoc = ""
'                                    sqlDoc = sqlDoc & "update lab_order set r_flag='1'"
'                                    sqlDoc = sqlDoc & " where resultno = '" & mAdoRs("resultno") & "'"
'                                    sqlDoc = sqlDoc & "   and patient_no = '" & mAdoRs("patient_no") & "'"
'                                    sqlDoc = sqlDoc & "   and lcode = '" & mAdoRs("lcode") & "'"
'                                    sqlDoc = sqlDoc & "   and pt_seq = '" & mAdoRs("pt_seq") & "'"
'
'                                    AdoCn_ORACLE.Execute sqlDoc
'
'                                    AdoCn_ORACLE.CommitTrans
'
'                                    lblStatus.Caption = "���� ����!!"
'                                    Set adoRS = Nothing:    mAdoRs.Close
'                                Loop
'                            Else
'                                blnFlag = False
'                                Set mAdoRs = f_subSet_TestList(strBarno)
'                                Do Until mAdoRs.EOF
'                                    If Trim(mAdoRs("LCODE")) = strTestcd Then blnFlag = True: Exit Do
'                                    mAdoRs.MoveNext
'                                Loop
'                                If blnFlag Then
'                                    AdoCn_SQL.BeginTrans
'
'                                    If chkAuto.Value = "1" Then
'                                           If Mid(pName, 1, 2) = "����" Then
'                                               sqlDoc = "Update MDCK..GUMJIN_INTERFACE" & _
'                                                        "   set RESULT = '" & strRstval & "'," & _
'                                                        "       ACT_RETURN_DATE = '" & strDate & "'" & _
'                                                        " where PER_GUMJIN_DATE = '" & strDate2 & "'" & _
'                                                        "   and PER_GUM_NUM = " & pNo & "" & _
'                                                        "   and EDPSCODE = '" & Mid(itemX.text, 1, 4) & "'"
'                                           Else
'                                               sqlDoc = "Update MEDICOM..jun370_resulttb" _
'                                                       & "   Set Result = '" & strRstval & "', status='1'" _
'                                                       & " Where WaitSeqNo = '" & pNo & "'" _
'                                                       & "   and map2seqno = '" & strEqpCd & "'"
'
'                                           End If
'                                           AdoCn_SQL.Execute sqlDoc
'                                    End If
'
'                                    AdoCn_SQL.CommitTrans
'
'                                    lblStatus.Caption = "���� ����!!"
'                                    Set adoRS = Nothing:    mAdoRs.Close
'                                End If
'                            End If
'                        End If
'
'                        Set itemX = Nothing
'                    End If
'                Next
'                spdResult1.Row = intRow
'                spdResult1.Col = 2
'                spdResult1.BackColor = vbCyan
'                spdResult1.Col = 3
'                spdResult1.BackColor = vbCyan
'                spdResult1.Col = 4
'                spdResult1.BackColor = vbCyan
'                spdResult1.Col = 1: spdResult1.Value = 0
'
'                If strErrMsg = "" Then
'                    sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
'                             " where SPCNO   = '" & strBarno & "'" & _
'                             "   and TRANSDT = '" & mskRstDate.text & "'"
'                    AdoCn_Jet.Execute sqlDoc
'                Else
'                    MsgBox strErrMsg, vbInformation, Me.Caption
'                End If
'            End If
'        Next
'    End With
'    Me.MousePointer = 0
'    MsgBox "�۾��� �Ϸ�Ǿ����ϴ�.", vbInformation, Me.Caption
'
'    Exit Sub
'ErrorRoutine:
'    Set itemX = Nothing
'
'    Me.MousePointer = 0
'    Call ErrMsgProc(CallForm)
'End Sub

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
    End With
    
    sqlDoc = "Select SPCNO, TESTCD, EQUIPCD, TRANSDT, RSTVAL, REFVAL, TRANSDT, EQPNUM, NAME, PNO" & _
             "  From INTERFACE003" & _
             " Where TRANSDT >= '" & Format(dtpRsltDay.Value, "yyyymmdd") & "'" & _
             "   And EQUIPCD = '" & INS_CODE & "'"
    If cboRstgbn(1).ListIndex = 0 Then
        sqlDoc = sqlDoc & "   And SERVERGBN = ''"
    ElseIf cboRstgbn(1).ListIndex = 1 Then
        sqlDoc = sqlDoc & "   And SERVERGBN = 'Y'"
    End If
    sqlDoc = sqlDoc & " Order By SPCNO, TRANSTM"
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        With spdResult2
        If strSpcno <> Trim$(adoRS(0) & "") + Trim$(adoRS(9) & "") Then
                intRow = intRow + 1
                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                .SetText 1, intRow, "1"
                .SetText 2, intRow, Trim$(adoRS(3) & "")
                .SetText 3, intRow, Trim$(adoRS(0) & "")
                .SetText 6, intRow, Trim$(adoRS(8) & "")
                .SetText 7, intRow, Trim$(adoRS(9) & "")
                '.SetText .MaxCols, intRow, Trim$(adoRS(6) & "")
            End If
            strSpcno = Trim$(adoRS(0) & "") + Trim$(adoRS(9) & "")
            Set itemX = lvwCuData.FindItem(Trim$(adoRS(7) & ""), lvwTag, , lvwWhole)
            If Not itemX Is Nothing Then
                intCol = itemX.Index + 8
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
    
    sNo = InputBox("���� ��ȣ�� �Է��ϼ��� !")
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
        If Not IsNumeric(sNo) Then
            MsgBox "���ڸ� �Է��ϼ���.!", vbCritical
            GoTo AgainInput
        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .maxrows
                .Row = sCnt
                .Col = 9:       .text = Trim(sAdd + Val(sNo))
                txtSeqNo.text = Trim(sAdd + Val(sNo))
                sAdd = sAdd + 1
            Next sCnt
        
            .StartingRowNumber = Val(sNo)
        End With
    End If
    
End Sub

Private Function f_subSet_TestList(ByVal strDate As String, Optional ByVal strSeq As String)
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
'    '-- WorkList��ȸ
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
'                '-- �˻��׸���ȸ
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
    Dim strRack_tmp As String

    blnFlag = False
    strRack_tmp = 1
    
    With spdWorklist
        For intRow1 = 1 To .maxrows
            .GetText 1, intRow1, varTmp
            If Trim$(varTmp) = "1" Then
                .GetText 2, intRow1, varTmp:    strWDate = Trim$(varTmp)
                .GetText 3, intRow1, varTmp:    strBarno = Trim$(varTmp)
                .GetText 4, intRow1, varTmp:    strSPnm = Trim$(varTmp)
                .GetText 5, intRow1, varTmp:    strChartNo = Trim$(varTmp)
                .GetText 6, intRow1, varTmp:    strSPid = Trim$(varTmp)
                               
                .Row = intRow1:
                
                .Col = 1: .ForeColor = HNC_Red
                .Col = 2: .ForeColor = HNC_Red
                .Col = 4: .ForeColor = HNC_Red
                .Col = 5: .ForeColor = HNC_Red
                .Col = 6: .ForeColor = HNC_Red

                intRow2 = f_funGet_SpreadRow(spdResult1, 6, strSPid)
                If intRow2 < 1 Then
                    intRow2 = f_funGet_SpreadRow(spdResult1, 2, "")
                    If intRow2 < 1 Then
                        spdResult1.maxrows = spdResult1.maxrows + 1
                        spdResult1.RowHeight(spdResult1.maxrows) = 13
                        intRow2 = spdResult1.maxrows
                    End If

                    blnFlag = False
                    
                    tmpDate = strWDate
                    
                    If cboChk.ListIndex = 0 Then
                        Set mAdoRs = f_subSet_WorkList_Barcode(tmpDate, strSPid, strSPnm)
                    Else
                        Set mAdoRs = f_subSet_WorkList_Barcode(tmpDate, strBarno, strSPnm)
                    End If

                    If Len(strSPid) > 0 Then
                        Do Until mAdoRs.EOF
                        
                            If cboChk.ListIndex = 0 Then
                                strEqpCd = f_funGet_CODE(Trim(mAdoRs("meditem")))
                            Else
                                strEqpCd = f_funGet_CODE(Trim(mAdoRs("map2seqno")))
                            End If
'
'                            strEqpCd = f_funGet_CODE(Trim(mAdoRs("EDPSCODE")))
                            
                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then
                                blnFlag = True
                                spdResult1.Row = intRow2
                                spdResult1.Col = itemX.Index + 9
                                spdResult1.BackColor = &HC6FEFF '&H80C0FF
                                spdResult1.text = " "
                                
                                DoEvents
                            End If
                            mAdoRs.MoveNext
                        Loop
                    End If
                    If blnFlag = True Then
                    
                        Dim tmpSeq As String

                        tmpSeq = txtSeqNo.text + 1
                        
                        spdResult1.SetText 2, intRow2, strWDate
                        spdResult1.SetText 3, intRow2, strBarno
                        spdResult1.SetText 4, intRow2, strSPnm
                        spdResult1.SetText 6, intRow2, strChartNo
                        spdResult1.SetText 7, intRow2, strSPid
                        
                        If (tmpSeq Mod 31) = 0 Then
                            strRack_tmp = strRack_tmp + 1
                            spdResult1.SetText 8, intRow2, strRack_tmp
                            tmpSeq = 1
                        Else
                            spdResult1.SetText 8, intRow2, strRack_tmp
                        End If
                        
                        spdResult1.Row = intRow2:
                        spdResult1.Col = 9:
                        spdResult1.ForeColor = HNC_Red
                        spdResult1.SetText 9, intRow2, tmpSeq
                    Else
                        spdResult1.maxrows = spdResult1.maxrows - 1
                    End If
                End If
                 spdResult1.SetText 1, intRow2, "1"

                .SetText 1, intRow1, ""
                If tmpSeq <> "" Then
                    txtSeqNo.text = tmpSeq
                End If
            End If
        Next
    End With
    
'    With spdResult1
'        varNum = 1
'        For iRow = 1 To .maxrows
'            If (iCnt Mod 31) = 0 Then
'                varNum = varNum + 1
'                iCnt = 1
'            End If
'            strRack_tmp = Format(varNum, "0")
'            .SetText 7, iRow, strRack_tmp
'            .SetText 8, iRow, ((iCnt Mod 31))
''                txtSeqNo.text = txtSeqNo.text + 1
'            iCnt = iCnt + 1
'        Next
'    End With
                
End Sub

Private Sub spdResult1_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim TmpYesno As String
Dim Tmpptno, TmpPtnm As String

    If Row = 0 Then
    
        If Col = 1 Then
            Col = 2
        End If
        
        If OrderSort_Flag = 1 Then
            Call SpreadSheetSort(spdResult1, Col, 2)
            OrderSort_Flag = 2
        Else
            Call SpreadSheetSort(spdResult1, Col, 1)
            OrderSort_Flag = 1
        End If
        
        Exit Sub
    End If


    If Col = 4 Or Col = 6 Then
        With spdResult1
            .Row = Row
            
            ' ���Ϲ�ȣ �ҷ�����
            .Col = 6
            Tmpptno = .text
            
            ' ȯ���̸� �ҷ�����
            .Col = 4
            TmpPtnm = .text
        End With
        
        If Len(Trim(Tmpptno)) >= 1 And Len(Trim(TmpPtnm)) >= 1 Then
             TmpYesno = MsgBox(Tmpptno & " (  " & TmpPtnm & "  ) " & " ȯ�ڸ� ���� �ϼ̽��ϴ�..    " & vbCrLf & vbCrLf & "�˻縦 ���� �Ͻðڽ��ϱ�..??", vbCritical + vbYesNo, App.Title)
        
             If TmpYesno = vbYes Then
                spdResult1.Action = ActionDeleteRow
                spdResult1.maxrows = spdResult1.maxrows - 1
             End If
        End If
    End If
        
End Sub

Private Sub spdResult1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
    
    Set oMenu = New cPopupMenu
    
    lMenuChosen = oMenu.Popup(" �� �˻��� �߰�", "-", " �� �˻��� ����", "-", " �� ���۹�ȣ����", "-", " �� ���� ����")

    Select Case lMenuChosen
        Case 1
            With spdResult1
                .maxrows = .maxrows + 1
                .Col = Col
                .Row = Row
                .Action = ActionInsertRow
            End With
        Case 3
            With spdResult1
                .Col = Col
                .Row = Row
                .Action = ActionDeleteRow
                .maxrows = .maxrows - 1
            End With
        Case 5
            Call cmdStartNo_Click
        Case 7
            Call cmdAppend_Click(0)
    End Select
End Sub


Private Sub ComReceive(ByRef RecData As String)
    
    Dim intIdx1     As Integer, intIdx2     As Integer, Loop_Count      As Integer
    Dim strTmp1     As String, strTmp2      As String
    Dim intPos1     As Integer, intPos2     As Integer
    Dim strDta()    As String, intCnt       As Integer
    Dim strRec      As String, strbuff      As String
    Dim pDoCount    As Integer

    strRec = RecData
    Print #1, strRec;
    Call COM_INPUT(strRec)
    'Debug.Print "1200 >>" & strRec
    
    For intIdx1 = 1 To Len(strRec)
        strbuff = Mid$(strRec, intIdx1, 1)
        
        Select Case strbuff
            Case ACK
                If SndCount > 0 Then
                    Call COM_OUTPUT(strOrdLst(SndCount))
                    lblStatus.Caption = SndCount - 1 & " ��° ���� ���� �Ϸ�"
                    Debug.Print SndCount - 1 & " ��° ���� ���� �Ϸ�"
                    Debug.Print "Advia1200 ==>" & strOrdLst(SndCount)
                    
                    Debug.Print "aSndCount>>" & SndCount
                    Debug.Print "asDeCnt>>" & sDeCnt
                    
                    If SndCount = sDeCnt Then
                        Timer1.Enabled = True
                        Timer1.Interval = 3000
                    End If
                    SndCount = SndCount + 1
'                    If sDeCnt = 1 Then
'                        Call COM_OUTPUT(EOT)
'                    End If
                ElseIf SndCount > sDeCnt Then
                    Call COM_OUTPUT(EOT)
                    Debug.Print "Advia1200>>EOT"
                    
                    Debug.Print "bSndCount>>" & SndCount
                    Debug.Print "bsDeCnt>>" & sDeCnt

                    SndCount = 0
                    sDeCnt = 0
                End If
            Case ETB:
                    f_strBuffer = f_strBuffer + strbuff
                    comEQP.Output = ACK
            Case EOT
                    If f_strOR_FLAG = "Q" Then
                        comEQP.Output = ENQ
                    Else
                        comEQP.Output = ACK
                    End If
                    f_strBuffer = ""
                    Exit Sub
            Case ENQ
                    comEQP.Output = ACK
                    f_strBuffer = ""
                    Exit Sub
            Case NAK
                    comEQP.Output = ACK
                    f_strBuffer = ""
                    Exit Sub
            Case STX
                    f_strBuffer = ""
                    f_strBuffer = f_strBuffer + strbuff
            Case ETX
                    If Mid$(f_strBuffer, 3, 1) = "Q" Then
                        f_strOR_FLAG = "Q"
                        Call RequestDefine(f_strBuffer, fChannel(), spdResult1)
                        Debug.Print "Advia1200 ==>" & f_strBuffer
                    ElseIf Mid$(f_strBuffer, 3, 1) = "R" Then
                        'Call COM_OUTPUT(ACK)
                        f_strOR_FLAG = "R"
                        Call psDataDefine(f_strBuffer, fChannel(), spdResult1)
                        Call COM_OUTPUT(ACK)
                        Debug.Print "Advia1200 ==>" & f_strBuffer
                    End If
                    f_strBuffer = ""
            Case Else
                    f_strBuffer = f_strBuffer + strbuff
        End Select
    Next
    
End Sub

Private Sub RequestDefine(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object) ', ByVal brOst As String) ' ByRef brItemdeci() As String)
    
    Dim Loop_Count, pDoCount As Integer
    Dim FunStr1, FunStr2 As String
    Dim PatientID As String
    Dim PatientNo As String
    Dim ii As Integer
    Dim Testcd As String
    Dim OutputData As String, sOrderLst As String
    Dim EndStr, strEqpCd, sChannel
    Dim intRow  As Integer
    Dim itemX As ListItem
    Dim sTemp, ssTemp1, ssTemp2 As String
    Dim sETBTemp    As String
    Dim strBarVal As String
    
    On Error GoTo errRequest
    
    sTemp = strdata
    If InStr(sTemp, "") <> 0 Then
        sETBTemp = Replace(sTemp, Mid(sTemp, InStr(sTemp, "") - 1, 17), "")
        'Debug.Print sETBTemp
    Else
        sETBTemp = sTemp
        'Debug.Print sETBTemp
    End If
    If InStr(sETBTemp, "") <> 0 Then
        sETBTemp = Replace(sETBTemp, Mid(sETBTemp, InStr(sETBTemp, "") - 1, 17), "")
    Else
        sETBTemp = sETBTemp
        'Debug.Print sETBTemp
    End If
   
    For Loop_Count = 1 To 100: fAdvia1650(Loop_Count) = "": Next Loop_Count
    
    For Loop_Count = 1 To 100: fAdvia1650Temp(Loop_Count) = "": Next Loop_Count
    
    sDeCnt = (Len(sETBTemp) - 12) / 13
    Debug.Print sDeCnt
    fAdvia1650(0) = Str$(sDeCnt)

    For pDoCount = 1 To sDeCnt
        ssTemp1 = (pDoCount - 1) * 13 + 12
        ssTemp2 = Trim(Mid$(sETBTemp, ssTemp1, 13))
        fAdvia1650(pDoCount) = Mid$(ssTemp2, 1, 13)
    
'        ssTemp1 = (pDoCount - 1) * 13 + 11
'        ssTemp2 = Trim(Mid$(sETBTemp, ssTemp1, 13))
'        fAdvia1650(pDoCount) = Mid$(ssTemp2, 1, 13)
    
    Next pDoCount
       
'    1O 0101  8N146           01-46                                  M            1.011 16M 19M 22M 25M 28M 73M 34M 64M E5
'    1Q 010101101-01          08
'    1Q 0101010A0404010021
    
    With spdResult1
        For pDoCount = 1 To sDeCnt
            Set mAdoRs = f_subSet_WorkList_Barcode(Trim(fAdvia1650(pDoCount)))
            
            If RecordChk = True Then
                If Not mAdoRs.EOF Then
                    strBarVal = Val(Trim(fAdvia1650(pDoCount)))
                    If Len(strBarVal) > 6 Then
                    'If cboChk.ListIndex = 0 Then
                        intRow = SeqSearch(spdResult1, mAdoRs.Fields("ChartNo"), 5)
                
                        If intRow = 0 Then
                            intRow = SeqNullSearch(spdResult1, mAdoRs.Fields("ChartNo"), 5)
                            If intRow = 0 Then .maxrows = .maxrows + 1: intRow = .maxrows
                        End If
                        
                        '.maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                        'intRow = .maxrows
                        .SetText 1, intRow, "1"
                        .SetText 2, intRow, mAdoRs("PER_GUMJIN_DATE")
                        .SetText 3, intRow, strBarVal
                        .SetText 4, intRow, mAdoRs("PER_NAME")
                        .SetText 5, intRow, mAdoRs("PER_GUM_NUM")
                        .SetText 7, intRow, mAdoRs("PER_SSN")
                        
                        
                        
                        Do Until mAdoRs.EOF
                            strEqpCd = f_funGet_CODE(mAdoRs("meditem"))
                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then
                                spdResult1.Row = intRow
                                spdResult1.Col = itemX.Index + 9
                                spdResult1.BackColor = &HC6FEFF
                                DoEvents
                            End If
                            If Len(strEqpCd) = 2 Then
                                strEqpCd = Space(1) & Val(strEqpCd)
                            End If
                            If Len(strEqpCd) = 1 Then
                                strEqpCd = Space(2) & Val(strEqpCd)
                            End If
                            sChannel = sChannel & strEqpCd & "M"
                            mAdoRs.MoveNext
                        Loop
                    Else
                        intRow = SeqSearch(spdResult1, mAdoRs.Fields("WAITSEQNO"), 3)
                
                        If intRow = 0 Then
                            intRow = SeqNullSearch(spdResult1, mAdoRs.Fields("WAITSEQNO"), 3)
                            If intRow = 0 Then .maxrows = .maxrows + 1: intRow = .maxrows
                        End If
                        
                        .SetText 1, intRow, "1"
                        .SetText 2, intRow, mAdoRs("ENTERDATE")
                        .SetText 3, intRow, mAdoRs("WAITSEQNO")
                        .SetText 4, intRow, Trim(mAdoRs("SUJINNAME"))
                        .SetText 5, intRow, mAdoRs("Chartno")
                        .SetText 6, intRow, mAdoRs("Chartno")
                        
                        Do Until mAdoRs.EOF
                            strEqpCd = f_funGet_CODE(mAdoRs("MAP2SEQNO"))
                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then
                                spdResult1.Row = intRow
                                spdResult1.Col = itemX.Index + 9
                                spdResult1.BackColor = &HC6FEFF
                                DoEvents
                            End If
                            If Len(strEqpCd) = 2 Then
                                strEqpCd = Space(1) & Val(strEqpCd)
                            End If
                            If Len(strEqpCd) = 1 Then
                                strEqpCd = Space(2) & Val(strEqpCd)
                            End If
                            sChannel = sChannel & strEqpCd & "M"
                            mAdoRs.MoveNext
                        Loop
                    
                    End If
                    
                    fAdvia1650Temp(1) = STX
                    fAdvia1650Temp(2) = CStr(Val(pDoCount) Mod 8) & "O" & Space(1)
                    'fAdvia1650Temp(3) = Mid$(sETBTemp, 5, 4)
                    fAdvia1650Temp(3) = "0101"
                    fAdvia1650Temp(4) = Format(Len(sChannel) / 4, "000") & "N0"
                    
                    '���ҽ�
'                    fAdvia1650Temp(5) = Trim(fAdvia1650(pDoCount))
'                    fAdvia1650Temp(6) = Space(40) & "M" & "   " & Format$(Now, "YYYYMMDD") & Space(1)
'                    fAdvia1650Temp(7) = "1.011"
'                    fAdvia1650Temp(8) = sChannel & Space(1)
                    
                    '�����ҽ�
                    fAdvia1650Temp(5) = Trim(fAdvia1650(pDoCount)) & " "                '13�ڸ� >> ��ź���Ϻ��� ���ڵ� 12�ڸ��̹Ƿ� ���� Space�� ���ڸ� ä��
                    fAdvia1650Temp(6) = Space(7)                                        'Position Number (7)
                    fAdvia1650Temp(6) = fAdvia1650Temp(6) & Space(16)                   'Comment1 (16)
                    fAdvia1650Temp(6) = fAdvia1650Temp(6) & Space(16)                   'Comment2 (16)
                    fAdvia1650Temp(6) = fAdvia1650Temp(6) & "M"                         'Sex (1)
                    fAdvia1650Temp(6) = fAdvia1650Temp(6) & "   "                       'Age (3)
                    fAdvia1650Temp(6) = fAdvia1650Temp(6) & Format$(Now, "YYYYMMDD")    'Blood sampling date (8)
                    fAdvia1650Temp(7) = " 1.0"                                          'Dilution coefficient (4)
                    fAdvia1650Temp(7) = fAdvia1650Temp(7) & "1"                         'Sample classification (1)
                    fAdvia1650Temp(7) = fAdvia1650Temp(7) & "1"                         'Container classification (1)
                    fAdvia1650Temp(8) = sChannel                                        'Request item  (xxxM : xxx ä��(����Space), M:fix)
                    fAdvia1650Temp(8) = fAdvia1650Temp(8) & Space(1)                    'Spare (1)

                    fAdvia1650Temp(9) = ETX
                    
                    FunStr1 = ""
                    FunStr2 = ""
                    
                    '���� ���� (STX ����)
                    For Loop_Count = 1 To 9
                        FunStr1 = FunStr1 + fAdvia1650Temp(Loop_Count)
                    Next Loop_Count
                    
                    'Checksum ��(STX ����)
                    For Loop_Count = 2 To 9
                        FunStr2 = FunStr2 + fAdvia1650Temp(Loop_Count)
                    Next Loop_Count
                    
                    '-----------------------------------------------------------------
        '            FunStr1 = ""
        '            FunStr2 = "1O 0101008N0A0404200101  " & Space(39) & "M" & "000" & "20040420" & Space(1) & "1.011 16M 19M 22M 25M 28M 73M 34M 64M "
        '            FunStr2 = "1O 0201015N0A0406191031                                         M00020040619 1.011  1M  4M  7M 22M 25M 28M 73M 34M 64M 40M 43M 49M 55M 67M 61MM 0A
                    
                    strOrdLst(pDoCount) = FunStr1 & MakeCS(FunStr2) & vbCr & vbLf
                    sChannel = ""
                End If
            End If
        Next pDoCount
    End With
    
    SndCount = 1
    Call COM_OUTPUT(ACK)
    mAdoRs.Close
    
    
    Exit Sub
    
errRequest:

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
    Dim OutputData  As String
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
           'strDta = ReceiveData
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
            strEVMsg = " CTS(Clear to Send) ���� ����"
        Case comEvDSR
            strEVMsg = " DSR(Data Set Read) ���� ����"
        Case comEvCD
            strEVMsg = " CD(Carrier Detecr) ���� ����"
        Case comEvRing
            strEVMsg = " ��ȭ ���� �︮�� ��"
        Case comEvEOF
            strEVMsg = " EOF(End Of File) ����"

        ' ���� �޽���
        Case comBreak
            strERMsg = " �ߴ� ��ȣ ����"
        Case comCDTO
            strERMsg = " �ݼ��� ���� �ð� �ʰ�"
        Case comCTSTO
            strERMsg = " CTS(Clear to Send) �ð� �ʰ�"
        Case comDCB
            strERMsg = " ��Ʈ�� ���� ��ġ ���� ���(DCB) �˻� �� ����ġ ���� ����"
        Case comDSRTO
            strERMsg = " DSR(Data Set Read) �ð� �ʰ�"
        Case comFrame
            strERMsg = " �����̹� ����"
        Case comOverrun
            strERMsg = " �и�Ƽ ����"
        Case comRxOver
            strERMsg = " ���� ���� �ʰ�"
        Case comRxParity
            strERMsg = " �и�Ƽ ����"
        Case comTxFull
            strERMsg = " ���� ���ۿ� ������ ����"
        Case Else
            strERMsg = " �� �� ���� ���� �Ǵ� �̺�Ʈ"
    End Select
    If Len(strERMsg) > 0 Then Call ShowMessage(strERMsg)
                
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڿ��� �����ڸ� �̿��� ������ ������ ��ġ�� ���ڿ��� ����
'   �μ� :
'       1.pText      : �����ڷ� ������ ���ڿ�
'       2.pPosiion   : ��ġ
'       3.pDelimiter : ������
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition �μ��� 1�� ��� For�� Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '�ش� �÷�
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function


'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڵ��ȣ�� ���� �������� ��ȸ, �������忡 ǥ��
'   �μ� :
'       - pBarNo : ���ڵ��ȣ
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)
    Dim intRow      As Integer
    Dim strEqpCd    As String
    
    If pBarNo = "" Then Exit Sub
    
    With spdResult1
        intRow = SeqSearch(spdResult1, pBarNo, 2)
        
        '-- ȯ������ ��������
        Set mAdoRs = f_subGet_PatInfo(pBarNo)
        Do Until mAdoRs.EOF
            .SetText 1, intRow, "üũ��ư��"
            .SetText 2, intRow, mAdoRs.Fields("ó������") & ""
            .SetText 3, intRow, mAdoRs.Fields("��ü��ȣ") & ""
            .SetText 4, intRow, mAdoRs.Fields("�̸�") & ""
            .SetText 5, intRow, mAdoRs.Fields("��Ϲ�ȣ") & ""
            .SetText 6, intRow, mAdoRs.Fields("����") & ""
            .SetText 7, intRow, mAdoRs.Fields("POS") & ""
            
            mAdoRs.MoveNext
        Loop
        
        mAdoRs.Close
        Set mAdoRs = Nothing
        
        '-- �˻��������� ��������
        Set mAdoRs1 = f_subSet_WorkList_Barcode(pBarNo)
    
        Do Until mAdoRs1.EOF
    
'            strEqpCd = f_funGet_CODE(Trim(mAdoRs1("ó���ڵ�")))
'
'            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
'            If Not itemX Is Nothing Then
'                blnFlag = True
'                .Row = intRow2
'                .Col = itemX.Index + 7
'                .BackColor = &HC6FEFF '&H80C0FF
'                .text = " "
'
'                DoEvents
'            End If
            mAdoRs.MoveNext
        Loop
        
        mAdoRs1.Close
        Set mAdoRs1 = Nothing
        
        .Row = intRow:
        .Col = 7:
        .ForeColor = HNC_Red
    
    End With

End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
'    Dim objIntNms    As clsIISIntNms     '��� �˻��׸� �÷��� Ŭ����
'    Dim objBuffer    As clsIISBuffer     '����Ŭ����

    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strBarno     As String   '������ ���ڵ��ȣ
    Dim strSeq       As String   '������ Sequence
    Dim strRackno    As String   '������ Rack Or Disk No
    Dim strTubePos   As String   '������ Tube Position
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   '������ ���
    Dim strFlag      As String   '������ Abnormal Flag
    Dim strComm      As String   '������ Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    
    Dim intCnt       As Integer
    
'    Set objIntNm = New clsIntTest
    Set objResult = New clsIntResults
    
    With cInterface
        For intCnt = 1 To .bufcnt
            strRcvBuf = .getrcvbuf(intCnt)
            strType = Mid$(strRcvBuf, 2, 1)
            
            Select Case strType
                Case "H"    '## Header
                Case "P"    '## Patient
                Case "Q"    '## Request Information
                    '## ���ڵ��ȣ, SEQ, Disk No, Tube Position ��ȸ
                    If mGetP(strRcvBuf, 13, "|") = "A" Then Exit Sub
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = Trim$(mGetP(strTemp1, 2, "^"))
    
                    Set objOrder = New clsIntOrder
                    With objOrder
                        .ClsClear
                        .BarNo = strBarno
                        .Seq = mGetP(strTemp1, 3, "^")
                        .RackNo = mGetP(strTemp1, 4, "^")
                        .TubePos = mGetP(strTemp1, 5, "^")
                    End With
                    Call GetOrder(strBarno)
                    cInterface.state = "Q"
                    
                Case "O"    '## Order
                    strBarno = Format$(mGetP(strRcvBuf, 3, "|"), String$(SPCLEN, "#"))
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strSeq = mGetP(strTemp1, 1, "^")
                    strRackno = Format$(mGetP(strTemp1, 2, "^"), "####")
                    strTubePos = Format$(mGetP(strTemp1, 3, "^"), "##")
    
                    Set objIntInfo = New clsIntInfo
                    With objIntInfo
                        .BarNo = strBarno
                        .SeqNo = strSeq
                        .SpcPos = strTubePos & "/" & strRackno
                    End With
                    
                Case "R"    '## Result
                    '## ������ �˻��, ���, Abnormal Flag
                    strTemp1 = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strTemp2 = mGetP(strRcvBuf, 4, "|")
                    strFlag = mGetP(strRcvBuf, 7, "|")
                    
                    '   - ����, ��������� ���ÿ� �����Ҽ� �ֵ��� ����
                    '## ����, ������ �������ó��, ����� "^"�� ���ԵǸ� �������
                    If InStr(strTemp2, "^") > 0 Then
                        strIntBase = strTemp1 & "C"
                        Select Case mGetP(strTemp2, 1, "^")
                            Case "-1":  strResult = "NEGATIVE"
                            Case "0":   strResult = "GRAYZONE"
                            Case "1":   strResult = "POSITIVE"
                        End Select
                        '## ������� ����
'                        If objIntNms.ExistIntBase(strIntBase) Then
'                            Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
'                                 strResult, strResult)
'                        End If
'
'                        '## ������� ����
'                        strIntBase = strTemp1 & "N"
'                        strResult = mGetP(strTemp2, 2, "^")
'                        If objIntNms.ExistIntBase(strIntBase) Then
'                            Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
'                                 strResult, strResult)
'                        End If
                    Else
                        '## ������� ����
                        strIntBase = strTemp1 & "N"
                        strResult = strTemp2
    
                        If strResult <> "" Then
                            Call objResult.Add(strIntBase, objIntNm.GetIntNm(strIntBase), strResult, strResult)
                        End If
                    End If
                    .state = "R"
                    
                    '-- 2010.03.11 osw �߰� : �˻��� �˾��޼��� �߰�
                    Call AddPopup("������", "12345789")

                Case "C"    '## Comment
                    '## Abnormal ����϶� Comment ����
                    If strFlag <> "N" Then
                        strTemp1 = mGetP(strRcvBuf, 4, "|")
                        strComm = "[Flag]: " & mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")
                        
                        '   - �������̽� ��� �÷����� �ش� ������ �˻���� �����Ҷ��� Comment�� �Է�
                        '     �ϵ��� ����
    '                    If objIntInfo.IntResults.Exist(strIntBase) Then
    '                        objIntInfo.IntResults(strIntBase).Info = strComm
    '                    End If
                    End If
                    
                Case "L"    '## Terminator
                    '## DB�� �������
    '                If mIntLib.State = "R" Then
    '                    Call SaveServer(objIntInfo)
    '                    Set objIntInfo = Nothing
    '                    mIntLib.State = ""
    '                End If
            End Select
        Next
    End With
'    Set objIntNms = Nothing
'    Set objBuffer = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
'Private Sub SendOrder()
'    Dim strOutput As String     '�۽��� ������
'    Dim mIntLib As clsInterface
'    Dim objOrder As clsIntOrder
'
'
'    Select Case mIntLib.phase
'        Case 1  '## Header
'            strOutput = mIntLib.FrameN & "H|\^&||||||||||P|1" & vbCr & ETX
'            mIntLib.phase = 2
'
'        Case 2  '## Patient
'            strOutput = mIntLib.FrameN & "P|1" & vbCr & ETX
'            mIntLib.phase = 4
'
'        Case 3  '## No Order
'
'        Case 4  '## Order
'            With objOrder
'                If .NoOrder = True Then
'                    '## ���������� �������
'                    strOutput = mIntLib.FrameN & "O|1|" & .BarNo & "|" & .Seq & "^" & .RackNo & _
'                                "^" & .TubePos & "^^SAMPLE^NORMAL|ALL" & _
'                                "|R||||||C||||||||||||||Q" & vbCr & ETX
'                    mIntLib.phase = 5
'                Else
'                    If .IsSending = False Then  '## ���� ������
'                        strOutput = "O|1|" & .BarNo & "|" & .Seq & "^" & .RackNo & "^" & .TubePos & _
'                                    "^^SAMPLE^NORMAL|" & .GetOrder & "|R||||||N||||||||||||||Q"
'                        If Len(strOutput) > 230 Then
'                            .IsSending = True
'                            .Order = Mid$(strOutput, 231)
'                            strOutput = mIntLib.FrameN & Mid$(strOutput, 1, 230) & vbCr & ETB
'                            mIntLib.phase = 4
'                        Else
'                            strOutput = mIntLib.FrameN & strOutput & vbCr & ETX
'                            mIntLib.phase = 5
'                        End If
'                    Else                        '## ���� ���ڿ��� ������
'                        strOutput = .Order
'                        If Len(strOutput) > 230 Then
'                            .Order = Mid$(strOutput, 231)
'                            strOutput = mIntLib.FrameN & Mid$(strOutput, 1, 230) & vbCr & ETB
'                            mIntLib.phase = 4
'                        Else
'                            .IsSending = False
'                            strOutput = mIntLib.FrameN & strOutput & vbCr & ETX
'                            mIntLib.phase = 5
'                        End If
'                    End If
'                End If
'            End With
'
'        Case 5  '## Termianator
'            strOutput = mIntLib.FrameN & "L|1" & vbCr & ETX
'            mIntLib.phase = 6
'
'        Case 6  '## EOT
'            mIntLib.state = ""
'            comEQP.Output = EOT
''            Call mIntLib.WriteLog(EOT, ccPCLog)
'            Exit Sub
'    End Select
'
'    strOutput = STX & strOutput & objOrder.GetChkSum(strOutput) & vbCrLf
'    comEQP.Output = strOutput
''    Call mIntLib.WriteLog(strOutput, ccPCLog)
'End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷����� ��� ��һ���
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    cInterface.clearRcvbuf
    
'    For i = mBuffers.Count To 1 Step -1
'        mBuffers.Remove i
'    Next i
End Sub


''''Private Sub psDataDefine(ByVal brbarcd As String, ByRef brChannel() As String, ByVal brspread As Object)
''''
''''Dim sTemp       As String       ' On Com���κ��� �Ѱܹ��� Receive Data
''''Dim Channel_No  As String       ' ������ ����
''''Dim Patiant_No  As String       ' ȯ�ڹ�ȣ
''''Dim pGrid_Point As Integer      ' �ش� �˻��� Point
''''Dim Max_Arary_Cnt As Integer    ' �˻� �׸��
'''''-------------------------------' �ӽ� ������.....
''''Dim sDeCnt      As Integer
''''Dim pDoCount    As Integer
''''Dim Loop_Count  As Integer
''''Dim sRtn As Integer, sChannel As String, sRstText As String, sRstValue As Single, sUnit As String
''''Dim itemX As ListItem
''''Dim strRstval(1 To 19) As String, strRefVal(1 To 19)  As String
''''Dim FunStr As String
''''Dim sqlDoc  As String
''''Dim intCol As Integer
''''Dim Test_Cd() As String, strPid()    As String, strPnm() As String
''''Dim Rev As Long
''''Dim ii As Integer
''''Dim tmpTstCd As String
''''Dim strLevel() As String
''''Dim chkPos  As Variant
''''Dim strResult As String
''''Dim strBarno    As String, strSPid  As String, strSPnm   As String
''''Dim strSex      As String, strOld   As String, strArea   As String
''''Dim varTmp  As Variant
''''Dim strDate As String, strTime  As String, sqlRet   As Integer
''''Dim strResultTmp As String
''''
''''    On Error GoTo errDefine
''''    sRstText = brbarcd
'''''    Debug.Print "sRstText : " & sRstText
''''    '------------------------------<<< fCobasE411() �迭 Clear �Ѵ�.         >>>----------
''''    For Loop_Count = 1 To 100: fCobasE411(Loop_Count) = "": Next Loop_Count
''''    '------------------------------<<< fCobasE411() �迭�� �����Ͽ� �ִ´�.  >>>----------
''''
''''    pDoCount = 0
'''''    sRstText = Mid(sRstText, STX)
''''    sRstText = Mid(sRstText, InStr(fRcvString, STX))
''''    Do While InStr(sRstText, "|") > 0
''''        pDoCount = pDoCount + 1
''''        fCobasE411(pDoCount) = Text_Redefine(sRstText, "|")
''''        sRstText = Mid$(sRstText, InStr(sRstText, "|") + 1)   ' �����ڰ� "|" �̴�....
''''        If pDoCount > 99 Then
''''            sRstText = ""
''''            Exit Do
''''        End If
''''    Loop
''''
''''    sRstText = ""
''''    If Mid$(fCobasE411(1), 3, 1) = "H" Then          ' "H" Head Message Display
''''        comEQP.Output = ACK
''''        Debug.Print "H [HOST] " & ACK
''''    ElseIf Mid$(fCobasE411(1), 3, 1) = "P" Then      ' "P" Patiant Information Data Process
''''        comEQP.Output = ACK
''''        Debug.Print "P [HOST] " & ACK
''''    ElseIf Mid$(fCobasE411(1), 3, 1) = "C" Then
''''        comEQP.Output = ACK
''''        Debug.Print "C [HOST] " & ACK
''''    ElseIf Mid$(fCobasE411(1), 3, 1) = "Q" Then      ' "Q" Patiant Information Data Process
''''        comEQP.Output = ACK
''''        Debug.Print "Q [HOST] " & ACK
''''    ElseIf Mid$(fCobasE411(1), 3, 1) = "O" Then      ' "O" Order Data Process
''''        comEQP.Output = ACK
''''        Debug.Print "O [HOST] " & ACK
''''        PatientID = fCobasE411(3)
''''        pDoCount = 0
''''        Do While InStr(fCobasE411(4), "^") > 0
''''            pDoCount = pDoCount + 1
''''            Select Case pDoCount
''''                Case 1:    PatientSeq = Text_Redefine(fCobasE411(4), "^")
''''                Case 2:    PatientRack = Text_Redefine(fCobasE411(4), "^")
''''                Case 3:    PatientPos = Text_Redefine(fCobasE411(4), "^")
''''                Case Else: Exit Do
''''            End Select
''''            fCobasE411(4) = Mid$(fCobasE411(4), InStr(fCobasE411(4), "^") + 1)   ' �����ڰ� "^" �̴�....
''''        Loop
''''
''''        Patiant_Recevid = False        ' ȯ�ڹ�ȣ Flag
''''        sPatiant_No = PatientSeq ' ȯ�ڹ�ȣ
''''        '-------------------------------------------<<< �ش�˻����� �ش�ȯ�ڸ� �O�´�.       >>>----------
''''        With brspread
''''            For pDoCount = 1 To .maxrows
''''                .Row = pDoCount: .Col = 6
''''                If Trim$(.text) = Trim$(PatientID) Then
''''                    vRow = pDoCount
''''                    Patiant_Recevid = True
''''                    Exit For
''''                End If
''''            Next pDoCount
''''        End With
''''
''''    ElseIf Mid$(fCobasE411(1), 3, 1) = "R" Then
''''        comEQP.Output = ACK
''''        Debug.Print "R [HOST] " & ACK
''''        Dim strChannel_No1 As String
''''        Dim strChannel_No2 As String
''''
''''        If Patiant_Recevid = True Then
''''            strChannel_No1 = Mid(fCobasE411(3), InStr(fCobasE411(3), "^^^") + 3)
''''            strChannel_No2 = left(strChannel_No1, InStr(strChannel_No1, "^^") - 1)
''''            Channel_No = strChannel_No2
''''            With spdResult1
''''                For pDoCount = 10 To .MaxCols
''''                    .Row = vRow
''''                    .Col = pDoCount
''''                    .GetText 6, vRow, varTmp:    strBarno = Trim$(varTmp)
''''                    .GetText 4, vRow, varTmp:    strSPnm = Trim$(varTmp)
''''                    .GetText 7, vRow, varTmp:    strSPid = Trim$(varTmp)
''''
''''                    .GetText pDoCount, 0, varTmp
''''                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
''''                    If Channel_No = itemX.tag Then
''''                        If Trim(fCobasE411(4)) <> "" Then
''''
''''                            If InStr(fCobasE411(4), "^") > 0 Then
''''                                strResult = Trim(Mid$(fCobasE411(4), InStr(fCobasE411(4), "^") + 1))
''''                            Else
''''                                strResult = Trim(fCobasE411(4))
''''                            End If
''''
'''''                            Select Case Channel_No
'''''                                Case "900"
'''''                                    strResult = Mid(fCobasE411(4), InStr(fCobasE411(4), "^") + 1)
'''''                                Case Else
'''''                                    strResult = Trim(fCobasE411(4))
'''''                            End Select
''''                             .text = strResult
''''                        Else
''''                            .text = ""
''''                        End If
''''
''''
''''                        If strResult <> "" Then
''''                            strDate = Format$(Now, "YYYYMMDD"): strTime = Format$(Now, "MMSS")
''''
'''''                            sqlDoc = "Update INTERFACE003" & _
'''''                                     "   set RSTVAL  = '" & strResult & "', REFVAL = ''" & _
'''''                                     " where SPCNO   = '" & strBarno & "'" & _
'''''                                     "   and EQPNUM  = '" & itemX.tag & "'" & _
'''''                                     "   and TRANSDT = '" & strDate & "'" & _
'''''                                     "   and TRANSTM = '" & strTime & "'"
'''''                            AdoCn_Jet.Execute sqlDoc
'''''
'''''                            sqlDoc = "insert into INTERFACE003(" & _
'''''                                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
'''''                                     "    values( '" & strBarno & "', '" & itemX.text & "', '" & itemX.tag & "'," & _
'''''                                     "            '" & strDate & "', '" & strTime & "'," & _
'''''                                     "            '" & strResult & "', ''," & _
'''''                                     "            '" & INS_CODE & "', '', '" & strSPnm & "', '" & strSPid & "')"
'''''                            AdoCn_Jet.Execute sqlDoc
''''
''''                            '-- ����������
'''''                            If chkAuto.Value = "1" Then
'''''                                sqlDoc = "Update EXAM_TOC set EX_INRV = '" & Trim(strResult) & "',EX_INST = '2',EX_DATE = '" & Format$(Now, "YYYYMMDD") & "',EX_INEM='1271'" _
'''''                                       & " where RE_RCID ='" & strSPid & "' And IN_CODE='" & itemX.text & "'"
'''''
'''''                                AdoCn_ORACLE.Execute (sqlDoc)
'''''                                lblStatus.Caption = "���� ����!!"
'''''                                AdoCn_ORACLE.Execute sqlDoc
'''''                            End If
''''
''''                            Set itemX = Nothing
''''                        End If
''''                    End If
''''                    .Col = 7: .ForeColor = vbRed: .BackColor = vbCyan
''''                    .SetText 1, vRow, 1
''''                Next pDoCount
''''            End With
''''        End If
''''    ElseIf Mid$(fCobasE411(1), 3, 1) = "L" Then      ' "L" Data Last
''''        comEQP.Output = ACK
''''        Debug.Print "L [HOST] " & ACK
''''        Patiant_Recevid = False                        ' ȯ�� ��ȣ  Flag
''''    Else
''''        comEQP.Output = ACK
''''    End If
''''
''''    Exit Sub
''''errDefine:
''''
''''End Sub



'Private Sub ComReceive(ByRef RecData As String)
'    Dim intIdx1     As Integer, intIdx2     As Integer
'    Dim strTmp1     As String, strTmp2      As String
'    Dim intPos1     As Integer, intPos2     As Integer
'    Dim strDta()    As String, intCnt       As Integer
'    Dim strRec      As String, strbuff      As String
'
'    Debug.Print RecData
'    strRec = RecData
'    Print #1, strRec;
'    Call COM_INPUT(strRec)
'    Debug.Print strRec
'
'    For intIdx1 = 1 To Len(strRec)
'        strbuff = Mid$(strRec, intIdx1, 1)
'
'        Select Case Asc(strbuff)
'            Case 2 '-- STX
'                        f_strBuffer = strbuff
'            Case 3 '-- ETX
'                        If Mid$(f_strBuffer, 2, 2) = "R " Or Mid$(f_strBuffer, 2, 2) = "RH" Then
''                            Call RequestDefine(f_strBuffer, fChannel(), spdResult1)
'                        ElseIf Mid$(f_strBuffer, 2, 2) = "D " Or Mid$(f_strBuffer, 2, 2) = "DH" Then
'                            Call psDataDefine(f_strBuffer, fChannel(), spdResult1)
'                        End If
'                        f_strBuffer = ""
'            Case Else
'                        f_strBuffer = f_strBuffer + strbuff
'        End Select
'    Next
'End Sub


Private Sub psDataDefine(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object) ', ByVal brOst As String) ' ByRef brItemdeci() As String)

'    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Privete sub psDataDefine()"
    
    Dim ssTemp1 As String
    Dim ssTemp2 As String
    Dim ssTemp3 As String
    
    Dim sqlDoc  As String, sqlRet   As Integer
    Dim sTemp      As String
    Dim Channel_No As Integer       ' �˻��׸� ��ȣ : Channel No
    Dim pGrid_Point As Integer
    Dim Loop_Count, pDoCount As Integer
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
    Dim strDate1 As String
    Dim chkGbn As String
    Dim intChNum As Integer
    Dim strBarVal As String
    
    sTemp = strdata
    
    For Loop_Count = 1 To 100: fAdvia1650(Loop_Count) = "": Next Loop_Count
    
    Erase fAdvia1200_CH
    Erase fAdvia1200_RS
    
    sDeCnt = (Len(sTemp) - 155) / 19                         ' �� �˻��׸� ������ �O�´�.
    fAdvia1650(0) = Str$(sDeCnt)                            ' �� �˻��׸� ������ �ִ´�. ���߿� ����Ѵ�.
    fAdvia1650(1) = Trim(Mid$(sTemp, 29, 13))
'fAdvia1650(1) = "201207075022"
    strBarVal = Val(fAdvia1650(1))

    For pDoCount = 1 To sDeCnt
        ssTemp1 = (pDoCount - 1) * 19 + 98              ' ù��° Channel �� �˻��� ��ġ Ȯ��
        ssTemp2 = Mid$(sTemp, ssTemp1, 19)
'        FunStr1 = Mid$(ssTemp2, 3, 7)
        fAdvia1650(((pDoCount - 1) * 2) + 9 + 1) = Mid$(ssTemp2, 1, 3)   ' channel
        fAdvia1650(((pDoCount - 1) * 2) + 9 + 2) = Trim(Mid$(ssTemp2, 5, 11))  ' result
    
        fAdvia1200_CH(pDoCount) = Mid$(ssTemp2, 1, 3)    ' channel
        fAdvia1200_RS(pDoCount) = Trim(Mid$(ssTemp2, 5, 11))  ' result
    
    Next pDoCount
     
    Max_Arary_Cnt = spdResult1.MaxCols
      
    pGrid_Point = 0
      
    Dim sSeq As String
    Dim sCol As Integer
      
    strTmp = ""
    If Len(fAdvia1650(1)) > 0 Then
        intRow = 0
        With spdResult1
            sSeq = Val(Trim(fAdvia1650(1)))
            sCol = 3
            pGrid_Point = SeqSearch(spdResult1, sSeq, sCol)
            
            If pGrid_Point = 0 Then
                Set mAdoRs = f_subSet_WorkList_Barcode(sSeq)
                
                If RecordChk = True Then
                    If Not mAdoRs.EOF Then
                        'If cboChk.ListIndex = 0 Then
                        If Len(strBarVal) > 6 Then
                            pGrid_Point = SeqSearch(spdResult1, mAdoRs.Fields("ChartNo"), 5)
                    
                            If pGrid_Point = 0 Then
                                pGrid_Point = SeqNullSearch(spdResult1, mAdoRs.Fields("ChartNo"), 5)
                                If pGrid_Point = 0 Then .maxrows = .maxrows + 1: pGrid_Point = .maxrows
                            End If
                            
'                        .SetText 2, intRow, mAdoRs("PER_GUMJIN_DATE")
'                        .SetText 3, intRow, strBarVal
'                        .SetText 4, intRow, mAdoRs("PER_NAME")
'                        .SetText 5, intRow, mAdoRs("CHARTNO")
'                        .SetText 7, intRow, mAdoRs("PER_SSN")
                        
                            .SetText 1, pGrid_Point, "1"
                            .SetText 2, pGrid_Point, mAdoRs("PER_GUMJIN_DATE")
                            .SetText 3, pGrid_Point, strBarVal
                            .SetText 4, pGrid_Point, mAdoRs("PER_NAME")
                            .SetText 5, pGrid_Point, mAdoRs("PER_GUM_NUM")
                            .SetText 7, pGrid_Point, mAdoRs("PER_SSN")
                            
                            mAdoRs.MoveNext
                        Else
                            pGrid_Point = SeqSearch(spdResult1, mAdoRs.Fields("WAITSEQNO"), 3)
                    
                            If pGrid_Point = 0 Then
                                pGrid_Point = SeqNullSearch(spdResult1, mAdoRs.Fields("WAITSEQNO"), 3)
                                If pGrid_Point = 0 Then .maxrows = .maxrows + 1: pGrid_Point = .maxrows
                            End If
                            
                            .SetText 1, pGrid_Point, "1"
                            .SetText 2, pGrid_Point, mAdoRs("ENTERDATE")
                            .SetText 3, pGrid_Point, mAdoRs("WAITSEQNO")
                            .SetText 4, pGrid_Point, Trim(mAdoRs("SUJINNAME"))
                            .SetText 5, pGrid_Point, mAdoRs("Chartno")
                            .SetText 6, pGrid_Point, mAdoRs("Chartno")
                            
                            mAdoRs.MoveNext
                            
                        
                        End If
                    End If
                End If
            End If

            .GetText 2, pGrid_Point, varTmp:   strDate1 = Trim$(varTmp)
            .GetText 3, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)
            .GetText 4, pGrid_Point, varTmp:   pName = Trim$(varTmp)
            .GetText 5, pGrid_Point, varTmp:   pNo = Trim$(varTmp)
            'chkGbn = Split(pName, ":")

            '.GetText 2, pGrid_Point, varTmp ':   strBarno = Trim$(varTmp)

            If pGrid_Point > 0 Then
                Set mAdoRs = f_subSet_WorkList_Barcode(strBarno, Mid(pName, 1, 2))
                For intCol = 10 To .MaxCols
                    strRstval = ""
                    .GetText intCol, 0, varTmp
                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                    If Not itemX Is Nothing Then
                        intChNum = 0
                        For intIdx = 1 To .MaxCols
                            If Len(fAdvia1650(intCol)) > 0 Then
                                strEqpCd = ""
                                Do Until mAdoRs.EOF
                                    'If cboChk.ListIndex = 0 Then
                                    
                                    If Len(strBarVal) > 6 Then
                                    
                                        If InStr(itemX.text, Trim(mAdoRs.Fields("meditem"))) > 0 Then
                                            strEqpCd = Trim(mAdoRs.Fields("meditem"))
                                            Exit Do
                                        End If
                                    Else
                                    
                                        If InStr(itemX.text, Trim(mAdoRs.Fields("MAP2SEQNO"))) > 0 Then
                                            strEqpCd = Trim(mAdoRs.Fields("MAP2SEQNO"))
                                            Exit Do
                                        End If
                                    End If
                                    mAdoRs.MoveNext
                                Loop
                                mAdoRs.MoveFirst

                                If Trim(strEqpCd) <> "" Then
'                                    If intCol = 10 Then
'                                        intChNum = intCol
'                                    Else
'                                        If intChNum = 0 Then
'                                            intChNum = 10
'                                        End If
'                                        intChNum = intChNum + 1
'                                    End If
                                    
                                    intChNum = intChNum + 1
                                    
                                    'Channel_No = Trim(fAdvia1650(intChNum))
                                    If Trim(fAdvia1200_CH(intChNum)) <> "" Then
                                        Channel_No = Trim(fAdvia1200_CH(intChNum))
                                    End If
                                    If UCase(Channel_No) = UCase(itemX.tag) Then
                                        'strRstval = Trim(fAdvia1650(intChNum + 1))
                                        strRstval = Trim(fAdvia1200_RS(intChNum))
                                        
                                         strDate1 = Format$(Now, "YYYYMMDD"):     strTime = Format$(Now, "MMSS")
                                        .SetText intCol, pGrid_Point, strRstval
                                        .Col = intCol:  .Row = pGrid_Point
                                                        .ForeColor = IIf(Trim$(strRefval) <> "", vbRed, vbBlack)

                                        sqlDoc = "Update INTERFACE003" & _
                                                 "   set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRefval & "'" & _
                                                 " where SPCNO   = '" & strBarno & "'" & _
                                                 "   and EQPNUM  = '" & itemX.tag & "'" & _
                                                 "   and TRANSDT = '" & strDate1 & "'" & _
                                                 "   and TRANSTM = '" & strTime & "'"
                                        AdoCn_Jet.Execute sqlDoc

                                        'If cboChk.ListIndex = 0 Then
                                        If Len(strBarVal) > 6 Then
                                            sqlDoc = "insert into INTERFACE003(" & _
                                                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                                     "    values( '" & strBarno & "', '" & strEqpCd & "', '" & itemX.tag & "'," & _
                                                     "            '" & strDate1 & "', '" & strTime & "'," & _
                                                     "            '" & strRstval & "', '" & strRefval & "'," & _
                                                     "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
                                        Else
                                            sqlDoc = "insert into INTERFACE003(" & _
                                                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                                     "    values( '" & strBarno & "', '" & strEqpCd & "', '" & itemX.tag & "'," & _
                                                     "            '" & strDate1 & "', '" & strTime & "'," & _
                                                     "            '" & strRstval & "', '" & strRefval & "'," & _
                                                     "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
                                        End If

                                        AdoCn_Jet.Execute sqlDoc

                                        'If chkAuto.Value = "1" Then
                                            'If cboChk.ListIndex = 0 Then
                                            If Len(strBarVal) > 6 Then
                                                sqlDoc = "Update ONIT..GUMJIN_INTERFACE" & _
                                                         "   set RESULT = '" & strRstval & "'," & _
                                                         "       ACT_RETURN_DATE = '" & strDate1 & "'" & _
                                                         " where PER_GUMJIN_DATE = '" & Mid(strBarVal, 1, 8) & "'" & _
                                                         "   and PER_GUM_NUM = " & pNo & "" & _
                                                         "   and EDPSCODE = '" & strEqpCd & "'"
                                            Else
                                                'sqlDoc = "Update onit_out..jun370_resulttb" _
                                                        & "   Set Result = '" & strRstval & "', status='1'" _
                                                        & " Where WaitSeqNo = '" & strBarno & "'" _
                                                        & "   and map2seqno = '" & strEqpCd & "'"
                                                        
                                                sqlDoc = "Update onit_out..jun370_resulttb" _
                                                        & "   Set Result = '" & strRstval & "'" _
                                                        & " Where WaitSeqNo = '" & strBarno & "'" _
                                                        & "   and map2seqno = '" & strEqpCd & "'"
                                            End If
                                            AdoCn_SQL.Execute sqlDoc
                                        'End If

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
                                        intChNum = 0
                                        Exit For
                                        
                                    End If
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

'''
'''Private Sub ReceiveTheData(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object) ', ByVal brOst As String) ' ByRef brItemdeci() As String)
'''
'''
'''    Dim sTemp      As String
'''    Dim Channel_No As String        ' �˻��׸� ��ȣ : Channel No
'''    Dim pGrid_Point As Integer
'''    Dim pDoCount   As Integer
'''    Dim Loop_Count As Integer
'''    Dim FunStr As String
'''    Dim Max_Arary_Cnt As Integer    ' �˻� �׸��
'''    Dim sAdd As Integer, sPosition As Integer
'''    Dim itemX As ListItem
'''    Dim strRstval As String, strRefval  As String
'''    Dim sqlDoc  As String
'''    Dim intCol As Integer
'''    Dim Gnum   As String
'''    Dim ii As Integer, jj As Integer, kk As Integer
'''    Dim Test_Cd() As String
'''    Dim Rev As Long
'''    Dim tmpTstCd As String
'''    Dim tmpMXD As Variant
'''    Dim sSeq, strTmp, varTmp, strBarno, strDate, strDate1, strTime As String
'''    Dim sCol As Integer
'''    Dim sDeCnt As Integer
'''    Dim Float_rate1 As String
'''    Dim Float_r
'''    Dim intIdx1     As Integer, intIdx2     As Integer, Loop_Count      As Integer
'''    Dim strTmp1     As String, strTmp2      As String
'''    Dim intPos1     As Integer, intPos2     As Integer
'''    Dim strDta()    As String, intCnt       As Integer
'''    Dim strRec      As String, strbuff      As String
'''    Dim pDoCount    As Integer
'''
'''    strRec = RecData
'''    Print #1, strRec;
'''    Call COM_INPUT(strRec)
'''    Debug.Print "1650 >>" & strRec
'''
'''    For intIdx1 = 1 To Len(strRec)
'''        strbuff = Mid$(strRec, intIdx1, 1)
'''
'''        Select Case strbuff
'''            Case ACK
'''                If SndCount > 0 Then
'''                    Call COM_OUTPUT(strOrdLst(SndCount))
'''                    lblStatus.Caption = SndCount - 1 & " ��° ���� ���� �Ϸ�"
'''                    Debug.Print "Advia1650 ==>" & strOrdLst(SndCount)
'''                    Debug.Print SndCount
'''                    If SndCount = sDeCnt Then
'''                        Timer1.Enabled = True
'''                        Timer1.Interval = 3000
'''                    End If
'''                    SndCount = SndCount + 1
''''                    If sDeCnt = 1 Then
''''                        Call COM_OUTPUT(EOT)
''''                    End If
'''                ElseIf SndCount > sDeCnt Then
'''                    Call COM_OUTPUT(EOT)
'''                    Debug.Print EOT
'''                    SndCount = 0
'''                    sDeCnt = 0
'''                End If
'''            Case ETB:
'''                    f_strBuffer = f_strBuffer + strbuff
'''                    comEQP.Output = ACK
'''            Case EOT
'''                    comEQP.Output = ENQ
'''                    f_strBuffer = ""
'''                    Exit Sub
'''            Case ENQ
'''                    comEQP.Output = ACK
'''                    f_strBuffer = ""
'''                    Exit Sub
'''            Case NAK
'''                    comEQP.Output = ACK
'''                    f_strBuffer = ""
'''                    Exit Sub
'''            Case STX
'''                   f_strBuffer = f_strBuffer + strbuff
'''            Case ETX
'''                    If Mid$(f_strBuffer, 3, 1) = "Q" Then
'''                        Call RequestDefine(f_strBuffer, fChannel(), spdResult1)
'''                        Debug.Print "Advia1650 ==>" & f_strBuffer
'''                    ElseIf Mid$(f_strBuffer, 3, 1) = "R" Then
'''                        Call psDataDefine(f_strBuffer, fChannel(), spdResult1)
'''                        Call COM_OUTPUT(ACK)
'''                        Debug.Print "Advia1650 ==>" & f_strBuffer
'''                    End If
'''                    f_strBuffer = ""
'''            Case Else
'''                    f_strBuffer = f_strBuffer + strbuff
'''        End Select
'''    Nextate2 As String
'''    Dim Float_rate  As String
'''    Dim intRow, intIdx As Integer
'''    Dim chrChk As Boolean
'''    Dim seqChk As Variant
'''    Dim chkGbn As Variant
'''    Dim strEqpCd As String
'''
'''    On Error Resume Next
'''
'''    CallForm = "frmInterface - Privete sub psDataDefine()"
'''
'''    pDoCount = 0
'''    Do While InStr(strdata, "|") > 0
'''        pDoCount = pDoCount + 1
'''        fTBA40FR(pDoCount) = Text_Redefine(strdata, "|")
'''        strdata = Mid$(strdata, InStr(strdata, "|") + 1)   ' �����ڰ� "|" �̴�....
'''        If pDoCount > 99 Then
'''            strdata = ""
'''            Exit Do
'''        End If
'''    Loop
'''
'''    pGrid_Point = 0
'''    strTmp = ""
'''
'''    If Mid$(fTBA40FR(1), 3, 1) = "H" Then          ' "H" Head Message Display
'''        comEQP.Output = ACK
'''        Debug.Print "[HOST] " & ACK
'''    ElseIf Mid$(fTBA40FR(1), 3, 1) = "P" Then      ' "P" Patiant Information Data Process
'''        comEQP.Output = ACK
'''        Debug.Print "[HOST] " & ACK
'''    ElseIf Mid$(fTBA40FR(1), 3, 1) = "C" Then
'''        comEQP.Output = ACK
'''        Debug.Print "[HOST] " & ACK
'''    ElseIf Mid$(fTBA40FR(1), 3, 1) = "Q" Then      ' "Q" Patiant Information Data Process
'''        comEQP.Output = ACK
'''    ElseIf Mid$(fTBA40FR(1), 3, 1) = "O" Then      ' "O" Order Data Process
'''        comEQP.Output = ACK
'''        Debug.Print "[HOST] " & ACK
'''        Patiant_Recevid = False                        ' ȯ�� ��ȣ  Flag
'''        strBarno = Val(Text_Redefine(fTBA40FR(4), "^"))  '' ȯ�ڹ�ȣ  "5450^0^57"
'''        '-------------------------------------------<<< �ش�˻����� �ش�ȯ�ڸ� �O�´�.       >>>----------
'''        If optSeq.Value = 1 Then
'''            sCol = 7
'''        Else
'''            sCol = 3
'''        End If
'''        pGrid_Point = SeqSearch(brspread, strBarno, sCol)
'''        Patiant_Recevid = (pGrid_Point > 0)
'''    ElseIf Mid$(fTBA40FR(1), 3, 1) = "R" Then      ' "R" Result Data Process
'''        Dim ssChannel() As String
'''        comEQP.Output = ACK
'''
'''        If Patiant_Recevid = True Then
'''            ssChannel = Split(fTBA40FR(3), "^")
'''            If UBound(ssChannel) > 3 Then
'''                fTBA40FR(3) = ssChannel(3)
'''                Channel_No = fTBA40FR(3)
'''            Else
'''                Channel_No = 0
'''            End If
''''            fTBA40FR(3) = fclsFunc.Text_Change(fTBA40FR(3), "^", "")    ' channel
''''            Channel_No = Val(fTBA40FR(3) / 10)                                   ' channel
'''            '-------------------------------------------<<< �ش�˻����� �O�´�.       >>>----------
'''            Max_Arary_Cnt = brspread.MaxCols - 6   ' �տ������� 5������ ȯ�� ���� �̱⶧����.... -6�� �Ѵ�.
'''                                                   ' �ش� �迭��  brItem(),brChannel() �̴�.
'''            With brspread
'''                '----------------------------------------------<<<<<<<<<,  ���ΰ˻��׸��� �O�´�.  >>>>>>>----------
'''
'''                For pDoCount = 1 To Max_Arary_Cnt
'''                    .Col = pDoCount + 6
'''                    If Channel_No > 0 And Channel_No = Val(brChannel(pDoCount)) Then          ' �˻����� ������...
'''                        If Trim(fTBA40FR(4)) <> "" Then
'''                            fTBA40FR(4) = Text_Change(fTBA40FR(4), ">", "")
'''                            fTBA40FR(4) = Text_Change(fTBA40FR(4), "<", "")
'''
'''                            If InStr(fTBA40FR(4), "^") > 0 Then
'''                                .text = Trim(Mid$(fTBA40FR(4), InStr(fTBA40FR(4), "^") + 1))
'''                            Else
'''                                .text = Trim(fTBA40FR(4))
'''                            End If
'''                        Else
'''                            .text = ""
'''                        End If
'''
'''                    End If
'''
'''                Next pDoCount
'''            End With
'''        End If
'''
'''        intRow = 0
'''        With spdResult1
'''            sCol = 8
'''            pGrid_Point = SeqNullSearch(spdResult1, sSeq, sCol)
'''
'''
'''
'''            .GetText 2, pGrid_Point, varTmp:   strDate1 = Trim$(varTmp)
'''            .GetText 3, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)
'''            .GetText 4, pGrid_Point, varTmp:   pName = Trim$(varTmp)
'''            .GetText 5, pGrid_Point, varTmp:   pNo = Trim$(varTmp)
'''            chkGbn = Split(pName, ":")
'''
'''            .GetText 2, pGrid_Point, varTmp ':   strBarno = Trim$(varTmp)
'''
'''            If pGrid_Point > 0 Then
'''                Set mAdoRs = f_subSet_WorkList_Barcode(strBarno, Mid(pName, 1, 2))
'''                For intCol = 8 To .MaxCols
'''                    strRstval = ""
'''                    .GetText intCol, 0, varTmp
'''                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
'''                    If Not itemX Is Nothing Then
'''                        For intIdx = 1 To .MaxCols
'''                            If Len(fELEC1010(9)) > 0 Then
'''                                strEqpCd = ""
'''                                Do Until mAdoRs.EOF
'''                                    If Mid(pName, 1, 2) = "����" Then
'''                                        If Mid(itemX.text, InStr(itemX.text, ",") + 1) = Trim(mAdoRs.Fields("EDPSCODE")) Then
'''                                            strEqpCd = Trim(mAdoRs.Fields("EDPSCODE"))
'''                                            Exit Do
'''                                        End If
'''                                    Else
'''                                        If Mid(itemX.text, InStr(itemX.text, ",") + 1) = Trim(mAdoRs.Fields("MAP2SEQNO")) Then
'''                                            strEqpCd = Trim(mAdoRs.Fields("MAP2SEQNO"))
'''                                            Exit Do
'''                                        End If
'''                                    End If
'''                                    mAdoRs.MoveNext
'''                                Loop
'''                                mAdoRs.MoveFirst
'''
'''                                If Trim(strEqpCd) <> "" Then
'''                                    fELEC1010_2 = Split(Trim(fELEC1010(intIdx + 10)), "=")
'''                                    Channel_No = Trim(fELEC1010_2(0))
'''                                    If UCase(Channel_No) = UCase(itemX.tag) Then
'''                                        fELEC1010_3 = Split(Trim(fELEC1010_2(1)), " ")
'''                                        If UCase(Channel_No) = "CL" Then
'''                                            strRstval = ""
'''                                        Else
'''                                            strRstval = Trim(Mid(fELEC1010_2(1), 3, 3))
'''                                        End If
'''                                         strDate1 = Format$(Now, "YYYYMMDD"):     strTime = Format$(Now, "MMSS")
'''                                        .SetText intCol, pGrid_Point, strRstval
'''                                        .Col = intCol:  .Row = pGrid_Point
'''                                                        .ForeColor = IIf(Trim$(strRefval) <> "", vbRed, vbBlack)
'''
'''                                        sqlDoc = "Update INTERFACE003" & _
'''                                                 "   set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRefval & "'" & _
'''                                                 " where SPCNO   = '" & strBarno & "'" & _
'''                                                 "   and EQPNUM  = '" & itemX.tag & "'" & _
'''                                                 "   and TRANSDT = '" & strDate1 & "'" & _
'''                                                 "   and TRANSTM = '" & strTime & "'"
'''                                        AdoCn_Jet.Execute sqlDoc
'''
'''                                        If cboChk.ListIndex = 0 Then
'''                                            sqlDoc = "insert into INTERFACE003(" & _
'''                                                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
'''                                                     "    values( '" & strBarno & "', '" & strEqpCd & "', '" & itemX.tag & "'," & _
'''                                                     "            '" & strDate1 & "', '" & strTime & "'," & _
'''                                                     "            '" & strRstval & "', '" & strRefval & "'," & _
'''                                                     "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
'''                                        Else
'''                                            sqlDoc = "insert into INTERFACE003(" & _
'''                                                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
'''                                                     "    values( '" & strBarno & "', '" & strEqpCd & "', '" & itemX.tag & "'," & _
'''                                                     "            '" & strDate1 & "', '" & strTime & "'," & _
'''                                                     "            '" & strRstval & "', '" & strRefval & "'," & _
'''                                                     "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
'''                                        End If
'''
'''                                        AdoCn_Jet.Execute sqlDoc
'''
'''                                        If chkAuto.Value = "1" Then
'''                                            If Mid(pName, 1, 2) = "����" Then
'''                                                sqlDoc = "Update MDCK..GUMJIN_INTERFACE" & _
'''                                                         "   set RESULT = '" & strRstval & "'," & _
'''                                                         "       ACT_RETURN_DATE = '" & strDate1 & "'" & _
'''                                                         " where PER_GUMJIN_DATE = '" & strDate & "'" & _
'''                                                         "   and PER_GUM_NUM = " & pNo & "" & _
'''                                                         "   and EDPSCODE = '" & strEqpCd & "'"
'''                                            Else
'''                                                sqlDoc = "Update MEDICOM..jun370_resulttb" _
'''                                                        & "   Set Result = '" & strRstval & "', status='1'" _
'''                                                        & " Where WaitSeqNo = '" & pNo & "'" _
'''                                                        & "   and map2seqno = '" & strEqpCd & "'"
'''                                            End If
'''                                            AdoCn_SQL.Execute sqlDoc
'''                                        End If
'''
'''                                        spdResult1.Row = pGrid_Point
'''                                        spdResult1.Col = 2
'''                                        spdResult1.BackColor = vbCyan
'''                                        spdResult1.Col = 3
'''                                        spdResult1.BackColor = vbCyan
'''                                        spdResult1.Col = 4
'''                                        spdResult1.BackColor = vbCyan
'''                                        spdResult1.Col = 5
'''                                        spdResult1.BackColor = vbCyan
'''                                        spdResult1.Col = 6
'''                                        spdResult1.BackColor = vbCyan
'''                                        spdResult1.Col = 7
'''                                        spdResult1.BackColor = vbCyan
'''                                        spdResult1.Col = 1: spdResult1.Value = 0
'''                                        Exit For
'''
'''                                    End If
'''                                End If
'''                            End If
'''                        Next intIdx
'''                    End If
'''                    Set itemX = Nothing
'''                Next
'''            End If
'''        End With
'''
'''        Set mAdoRs = Nothing
'''
'''    ElseIf Mid$(fTBA40FR(1), 3, 1) = "L" Then      ' "L" Data Last
'''        comEQP.Output = ACK
'''        Debug.Print "[HOST] " & ACK
'''        Patiant_Recevid = False                        ' ȯ�� ��ȣ  Flag
'''    End If
'''
'''    Exit Sub
'''
'''ErrRoutine:
'''
'''    Call ErrMsgProc(CallForm)
'''
'''End Sub

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

    SeqSearch = 0
    If brspread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brspread
        If optSeq.Value = False Then
            For sCnt = 1 To .maxrows
                .Row = sCnt
                .Col = brCol
                If Val(.text) = brSeq Then
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
                If Val(.text) = Val(brSeq) Then
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
    Dim strTmp  As String
   
                  ReceiveData = ENQ & "1H|\^&|||Roche^OMNI-C^1.45^1^1758||||||Meas|P|1394-97|20100301235824" & vbCr
    ReceiveData = ReceiveData & "92"
    ReceiveData = ReceiveData & "2P|1||||^^||00000|U||||||||0^cm|0.0^kg||||||00000000000000/00000000000000|||||||||||" & vbCr
    ReceiveData = ReceiveData & "C7"
    ReceiveData = ReceiveData & "3O|1||MEASUREMENT^33942||||||||||||blood^arterial^" & vbCr
    ReceiveData = ReceiveData & "D0"
    ReceiveData = ReceiveData & "4C|1|I||G" & vbCr
    ReceiveData = ReceiveData & "38"
    ReceiveData = ReceiveData & "5R|1|^^^pH^^^M^1|7.513||7.350^7.450^reference\7.200^7.600^critical|H||F||||20100301175358|" & vbCr
    ReceiveData = ReceiveData & "E7"
    ReceiveData = ReceiveData & "6R|2|^^^PCO2^^^M^4|24.7|mmHg|35.0^45.0^reference\20.0^60.0^critical|L||F|||||" & vbCr
    ReceiveData = ReceiveData & "02"
    ReceiveData = ReceiveData & "7R|3|^^^PO2^^^M^3|123.0|mmHg|80.0^100.0^reference\60.0^800.0^critical|H||F|||||" & vbCr
    ReceiveData = ReceiveData & "43"
    ReceiveData = ReceiveData & "0R|4|^^^Na^^^M^6|139.3|mmol/L|135.0^148.0^reference\125.0^160.0^critical|N||F|||||" & vbCr
    ReceiveData = ReceiveData & "43"
    ReceiveData = ReceiveData & "1R|5|^^^K^^^M^7|4.07|mmol/L|3.50^4.50^reference\2.80^6.00^critical|N||F|||||" & vbCr
    ReceiveData = ReceiveData & "E9"
    ReceiveData = ReceiveData & "2R|6|^^^Cl^^^M^9|105.8|mmol/L|98.0^107.0^reference\80.0^115.0^critical|N||F|||||" & vbCr
    ReceiveData = ReceiveData & "EB"
    ReceiveData = ReceiveData & "3R|7|^^^iCa^^^M^8|-|mmol/L|1.120^1.320^reference\1.050^1.500^critical|A||X|||||" & vbCr
    ReceiveData = ReceiveData & "CE"
    ReceiveData = ReceiveData & "4R|8|^^^tHb^^^M^10|-|g/dL|11.5^17.4^reference\8.0^23.0^critical|A||X|||||" & vbCr
    ReceiveData = ReceiveData & "3A"
    ReceiveData = ReceiveData & "5R|9|^^^SO2^^^M^11|-|%|75.0^99.0^reference\60.0^100.0^critical|A||X|||||" & vbCr
    ReceiveData = ReceiveData & "37"
    ReceiveData = ReceiveData & "6R|10|^^^Hct^^^M^5|-|%|35.0^50.0^reference\25.0^65.0^critical|A||X|||||" & vbCr
    ReceiveData = ReceiveData & "48"
    ReceiveData = ReceiveData & "7R|11|^^^Temperature^^^I^155|37.0|�C||||F|||||" & vbCr
    ReceiveData = ReceiveData & "4C" & vbCrLf
    ReceiveData = ReceiveData & "0R|12|^^^Baro^^^M^31|754.2|mmHg||||F|||||" & vbCr
    ReceiveData = ReceiveData & "D7" & vbCrLf
    ReceiveData = ReceiveData & "1R|13|^^^cHCO3^^^C^51|19.5|mmol/L||||F|||||" & vbCr
    ReceiveData = ReceiveData & "31" & vbCrLf
    ReceiveData = ReceiveData & "2R|14|^^^ctCO2(P)^^^C^52|20.2|mmol/L||||F|||||" & vbCr
    ReceiveData = ReceiveData & "F5" & vbCrLf
    ReceiveData = ReceiveData & "3R|15|^^^SO2(c)^^^C^58|99.1|%||||F|||||" & vbCr
    ReceiveData = ReceiveData & "4D" & vbCrLf
    ReceiveData = ReceiveData & "4R|16|^^^BE^^^C^53|-1.6|mmol/L||||F|||||" & vbCr
    ReceiveData = ReceiveData & "45" & vbCrLf
    ReceiveData = ReceiveData & "5R|17|^^^BEecf^^^C^55|-3.5|mmol/L||||F|||||" & vbCr
    ReceiveData = ReceiveData & "78" & vbCrLf
    ReceiveData = ReceiveData & "6R|18|^^^BB^^^C^56|46.4|mmol/L||||F|||||" & vbCr
    ReceiveData = ReceiveData & "53" & vbCrLf
    ReceiveData = ReceiveData & "7R|19|^^^ctO2^^^C^60|-|Vol%||A||X|||||" & vbCr
    ReceiveData = ReceiveData & "FE" & vbCrLf
    ReceiveData = ReceiveData & "0R|20|^^^ctCO2(B)^^^C^61|-|mmol/L||A||X|||||" & vbCr
    ReceiveData = ReceiveData & "A0" & vbCrLf
    ReceiveData = ReceiveData & "1R|21|^^^pHst^^^C^62|7.384|||||F|||||" & vbCr
    ReceiveData = ReceiveData & "68" & vbCrLf
    ReceiveData = ReceiveData & "2R|22|^^^cHCO3st^^^C^63|23.0|mmol/L||||F|||||" & vbCr
    ReceiveData = ReceiveData & "12" & vbCrLf
    ReceiveData = ReceiveData & "3R|23|^^^HbI^^^C^200|-|||A||X|||||" & vbCr
    ReceiveData = ReceiveData & "66" & vbCrLf
    ReceiveData = ReceiveData & "4R|24|^^^PAO2^^^C^64|123.0|mmHg||||F|||||" & vbCr
    ReceiveData = ReceiveData & "5C" & vbCrLf
    ReceiveData = ReceiveData & "5R|25|^^^AaDO2^^^C^65|0.0|mmHg||||F|||||" & vbCr
    ReceiveData = ReceiveData & "4E" & vbCrLf
    ReceiveData = ReceiveData & "6R|26|^^^a/AO2^^^C^66|100.0|%||||F|||||" & vbCr
    ReceiveData = ReceiveData & "39" & vbCrLf
    ReceiveData = ReceiveData & "7R|27|^^^avDO2^^^C^67|-|%||A||X|||||" & vbCr
    ReceiveData = ReceiveData & "17" & vbCrLf
    ReceiveData = ReceiveData & "0R|28|^^^RI^^^C^68|0|%||||F|||||" & vbCr
    ReceiveData = ReceiveData & "C1" & vbCrLf
    ReceiveData = ReceiveData & "1R|29|^^^niCa^^^C^70|-|mmol/L||A||X|||||" & vbCr
    ReceiveData = ReceiveData & "F7" & vbCrLf
    ReceiveData = ReceiveData & "2R|30|^^^AG^^^C^71|18.1|mmol/L||||F|||||" & vbCr
    ReceiveData = ReceiveData & "46" & vbCrLf
    ReceiveData = ReceiveData & "3R|31|^^^pHt^^^C^72|7.513|||||F|||||" & vbCr
    ReceiveData = ReceiveData & "F3" & vbCrLf
    ReceiveData = ReceiveData & "4R|32|^^^H+t^^^C^73|30.658|nmol/L||||F|||||" & vbCr
    ReceiveData = ReceiveData & "18" & vbCrLf
    ReceiveData = ReceiveData & "5R|33|^^^PCO2t^^^C^74|24.7|mmHg||||F|||||" & vbCr
    ReceiveData = ReceiveData & "AB" & vbCrLf
    ReceiveData = ReceiveData & "6R|34|^^^PO2t^^^C^75|123.0|mmHg||||F|||||" & vbCr
    ReceiveData = ReceiveData & "94" & vbCrLf
    ReceiveData = ReceiveData & "7R|35|^^^PAO2t^^^C^76|123.0|mmHg||||F|||||" & vbCr
    ReceiveData = ReceiveData & "D8" & vbCrLf
    ReceiveData = ReceiveData & "0R|36|^^^AaDO2t^^^C^77|0.0|mmHg||||F|||||" & vbCr
    ReceiveData = ReceiveData & "C2" & vbCrLf
    ReceiveData = ReceiveData & "1R|37|^^^a/AO2t^^^C^78|100.0|%||||F|||||" & vbCr
    ReceiveData = ReceiveData & "AD" & vbCrLf
    ReceiveData = ReceiveData & "2R|38|^^^RIt^^^C^79|0|%||||F|||||" & vbCr
    ReceiveData = ReceiveData & "3A" & vbCrLf
    ReceiveData = ReceiveData & "3R|39|^^^Hct(c)^^^C^80|-|%||A||X|||||" & vbCr
    ReceiveData = ReceiveData & "48" & vbCrLf
    ReceiveData = ReceiveData & "4R|40|^^^MCHC^^^C^81|-|g/dL||A||X|||||" & vbCr
    ReceiveData = ReceiveData & "AB" & vbCrLf
    ReceiveData = ReceiveData & "5R|41|^^^BO2^^^C^84|-|||A||X|||||" & vbCr
    ReceiveData = ReceiveData & "12" & vbCrLf
    ReceiveData = ReceiveData & "6R|42|^^^BEact^^^C^54|-|mmol/L||A||X|||||" & vbCr
    ReceiveData = ReceiveData & "3D" & vbCrLf
    ReceiveData = ReceiveData & "7R|43|^^^Osm^^^C^82|277.6|mOsm/kg||||F|||||" & vbCr
    ReceiveData = ReceiveData & "A1" & vbCrLf
    ReceiveData = ReceiveData & "0R|44|^^^OER^^^C^83|-|%||A||X|||||" & vbCr
    ReceiveData = ReceiveData & "57" & vbCrLf
    ReceiveData = ReceiveData & "1R|45|^^^Qs/Qt^^^C^69|-|%||A||X|||||" & vbCr
    ReceiveData = ReceiveData & "2F" & vbCrLf
    ReceiveData = ReceiveData & "2R|46|^^^Qt^^^C^86|-|%||A||X|||||" & vbCr
    ReceiveData = ReceiveData & "3D" & vbCrLf
    ReceiveData = ReceiveData & "3R|47|^^^P/F^^^C^88|585.6|mmHg||||F|||||" & vbCr
    ReceiveData = ReceiveData & "2B" & vbCrLf
    ReceiveData = ReceiveData & "4R|48|^^^ALLEN test^^^I^152|unknown|||||F|||||" & vbCr
    ReceiveData = ReceiveData & "63" & vbCrLf
    ReceiveData = ReceiveData & "5R|49|^^^A/F^^^I^154|adult|||||F|||||" & vbCr
    ReceiveData = ReceiveData & "DB" & vbCrLf
    ReceiveData = ReceiveData & "6R|50|^^^P50^^^I^156|26.7|mmHg||||F|||||" & vbCr
    ReceiveData = ReceiveData & "11" & vbCrLf
    ReceiveData = ReceiveData & "7R|51|^^^R^^^I^157|0.840|||||F|||||" & vbCr
    ReceiveData = ReceiveData & "55" & vbCrLf
    ReceiveData = ReceiveData & "0R|52|^^^FIO2^^^I^158|0.210|||||F|||||" & vbCr
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "1L|1|N" & vbCr
    ReceiveData = ReceiveData & "04" & vbCrLf
    ReceiveData = ReceiveData & ""

    ReceiveData = ENQ
    ReceiveData = ReceiveData & "1H|\^&||||||||||P||" & vbCr
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "2P|1|||||||||||||||||||||||||||||||||" & vbCr
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "3O|1|0001|1^0001^1^^SAMPLE^NORMAL|ALL|R|20030722194828|||||X||||||||||||||O|||||" & vbCr
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "4R|1|^^^250^^0|22.30|ng/ml|25.00^72.00|L||F|||20030722195530|20030722200528|" & vbCr
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "5C|1|I|48^Below expected value range|I  DA"
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "6R|2|^^^10^^0|0.058|mIU/l|0.270^4.20|L||F|||20030722195448|20030722201310|" & vbCr
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "7C|1|I|48^Below expected value range|I" & vbCr
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "0L|1" & vbCr
    ReceiveData = ReceiveData & "04" & vbCrLf
    ReceiveData = ReceiveData & ""
 
'    Call ComReceive(strTmp)
    ReceiveData = "1Q 0101010201207075022  " & vbCrLf

    ReceiveData = "1R 010101420120713165625N0000000038269                                        M  020120713 1.011  3M       7    bay  4M     0.5    bay  5M    -0.4    bay  8M     3.2    bay  9M      15    bay 10M       8    bay 11M      10    bay 14M     256    bay 15M     245    bay 16M      59    bay 61M     148    bay 78M       -    bay 79M       -    bay 80M       -    bay " & vbCrLf

'1R 010101220120713165710N0000000038773                                        M  020120713 1.011  3M      16    bay  4M     1.4    bay  8M     5.2    bay  9M      38    bay 10M      53    bay 14M     181    bay 15M     321    bay 16M      36    bay 61M      81    bay 78M       -    bay 79M       -    bay 80M       -    bay 
'1R 010100220120713170359N0000000038807                                        M  020120713 1.011  4M     0.9    bay 12M      27    bay 
'1R 010101420120713165755N0000000038462                                        M  020120713 1.011  3M       6    bay  4M     0.6    bay  5M    -0.6    bay  8M     4.6    bay  9M      23    bay 10M      14    bay 11M      13    bay 14M     194    bay 15M     212    bay 16M      56    bay 61M      96    bay 78M       -    bay 79M       -    bay 80M       -    bay 
'1R 010101320120713165840N0000000038771                                        M  020120713 1.011  3M      12    bay  4M     1.0    bay  5M     0.0    bay  8M     6.3    bay 10M
    Call comEQP_OnComm
    
End Sub

Private Sub Form_Activate()

    If IS_SET = False Then Unload Me

End Sub

Private Sub Form_Load()
    
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    CaptionBar1.Caption = INS_NAME & " Communication"
    
    Call cmdClear               ' �ʱ�ȭ
    Call f_subSet_ItemHeader    ' ����Ʈ�ش�
    Call f_subSet_ItemList      ' �˻��׸�
    
    Call f_subSet_ComCharacter  ' ��Ź���
    Call f_subGet_Setting       ' ��ż���
    
    Call cmdRun                 ' ����
    
    dtpRsltDay.Value = Now
    dtpStartDt.Value = Now
    dtpStopDt.Value = Now
    mskOrdtime.text = Format$(Now, "HHMM")
    
    Open App.Path + "\Log\" + REG_INSNAME & Format(Now, "yyyymmdd") & ".log" For Append As #1

    Print #1, Chr(13) + Chr(10);
    
    Open App.Path + "\ErrorLog\" + REG_INSNAME + "_" + Format(Now, "YYYYMMDD") + ".sql" For Append As #2

    Print #2, Chr(13) + Chr(10);
   
    f_strJOB_FLAG = "1":    f_intSampleNo = 0
    tabWork.Tab = 0
    Or_Seq = 1
    intRow = 0
    chkEnq = 0
    cboChk.ListIndex = 1
    
    gspdResultRow = 0
    
    '-- 2010.03.11 osw �߰� : �˻��� �˾��޼���
'    Set mobjPopups = New PopUpMessages
'    With mobjPopups
'       ' .XPos = Screen.Width / 2
'       ' .YPos = 0
'       ' .PopUpDirection = vbPopDown
'        .ShowDelay = 3000
'        .MovementIndex = 5
'        .ScrollDelay = 30
'
'    End With
'
'    SetupDefaultPopup
    
    COM_MODE = "1"
    
    SndCount = 0
    sDeCnt = 0
    Erase strOrdLst
    
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
        MsgBox INS_CODE & " �� ���� ��� ��� ������ �����ϴ�. ��� ������ �ٽ� �õ� �Ͻʽÿ�.", vbExclamation
        Exit Sub
    Else
        If mAdoRs.EOF Then
            IS_SET = False
            MsgBox INS_CODE & " �� ���� ��� ��� ������ �����ϴ�. ��� ������ �ٽ� �õ� �Ͻʽÿ�.", vbExclamation
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
    Close #2
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
    
        tmrWorking.Interval = 20000
        tmrWorking.Enabled = True
    End If

End Sub

Private Sub Label6_DblClick()
    If Command1.Visible = False Then
        Command1.Visible = True
    Else
        Command1.Visible = False
    End If
End Sub

Private Sub Label9_DblClick()

    If COM_MODE = "1" Then
        COM_MODE = "0"
        ShowMessage "�������̽� ������ ȭ�鿡 ������� �ʽ��ϴ�."
    Else
        COM_MODE = "1"
        ShowMessage "�������̽� ������ ȭ�鿡 ����մϴ�."
    End If
End Sub

Private Sub tabWork_Click(PreviousTab As Integer)
    cboRstgbn(1).ListIndex = 2
'    spdResult2.maxrows = 0
End Sub

Private Sub tmrOrder_Timer()
    Dim OrderCnt    As Integer
    Dim ii          As Integer
    
    '-- ���������� �����ִ��� üũ
    With spdResult1
        For ii = 1 To .maxrows
            .Col = 1: .Row = ii
            If .Value = 1 Then
                .Col = 2
                If .BackColor <> vbCyan Then
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
                    If Len(Trim(.text)) > 0 Then
                        '-- osw edit
                        .Row = ii '+ 1
                        If Len(Trim(.text)) > 0 Then
                            comEQP.Output = ENQ
                            Debug.Print "[HOST] " & ENQ
                        End If
                        SendCount = 0
                        Exit For
                    End If
                End If
            End If
        Next
    End With
    
    tmrOrder.Enabled = False

End Sub

Private Sub tmrWorking_Timer()
    pnlCom.Visible = False
End Sub

'Private Sub mskOrdDate_GotFocus()
'
'    With mskOrdDate
'        .SelStart = 8
'        .SelLength = Len(.text)
'    End With
'
'End Sub


'Private Sub mskOrdDate_KeyPress(KeyAscii As Integer)
'
'    If Not KeyAscii = vbKeyBack Then mskOrdDate.SelLength = 1
'
'End Sub


'Private Sub mskRstDate_GotFocus()
'
'    With mskRstDate
'        .SelStart = 0
'        .SelLength = Len(.text) + 2
'    End With '
'
'End Sub
'
'
'Private Sub mskRstDate_KeyPress(KeyAscii As Integer)
'
'    If Not KeyAscii = vbKeyBack Then mskRstDate.SelLength = 1
'
'End Sub

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

Private Sub spdRstview_Click(ByVal Col As Long, ByVal Row As Long)
Dim iCnt, rCnt As Integer
Dim intCol, intRow As Integer
Dim tCol As Integer
Dim iresult As String
'
' ��� ���� Position
'
Const sResultPos As Integer = 8
    With spdRstview
        For iCnt = 2 To .MaxCols Step 2
            For rCnt = 1 To .maxrows
                .Row = rCnt: .Col = iCnt
                iresult = Trim(.text)
                
                With spdResult1
                    .Row = gspdResultRow:  .Col = sResultPos + tCol
                    If Len(Trim(iresult)) <> 0 Then
                        .text = iresult
                    End If
                    DoEvents
                End With
                tCol = tCol + 1
                
            Next rCnt
            rCnt = 0
        Next iCnt
    End With
End Sub

Private Sub spdRstview_EnterRow(ByVal Row As Long, ByVal RowIsLast As Long)
    Call spdRstview_Click(Row, RowIsLast)
End Sub

'
'
'

Private Sub spdRstview_KeyPress(KeyAscii As Integer)

Dim iCnt, rCnt As Integer
Dim intCol, intRow As Integer
Dim tCol As Integer
Dim iresult As String

'
' ��� ���� Position
'
Const sResultPos As Integer = 8
     
    ' ó�� ���� ���� Ȯ��..
    With spdRstview
        .Row = .ActiveRow: .Col = .ActiveCol
        If .BackColor <> &HC6FEFF And Len(.text) >= 1 Then
            .text = ""
            MsgBox "�� OCS/EMR�� �˻� ó���� ���� �׸� �Դϴ�.." & Space(5), vbOKOnly + vbInformation, App.Title
            spdRstview.SetFocus
            Exit Sub
        End If
    End With
    
    ' Enter Key ����..
    If KeyAscii = vbKeyReturn Then
    
        If gspdResultRow < 1 Then
            With spdRstview
                .Row = .ActiveRow:  .Col = .ActiveCol
                .text = ""
            End With
            
            MsgBox "�� ������ ���ϴ� �˻� Sample�� ���� �� ���� �Ͻʽÿ�.." & Space(5), vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        
        ' ������ ��� �� Spread�� �ű��..
        With spdRstview
            For iCnt = 2 To .MaxCols Step 2
                For rCnt = 1 To .maxrows
                    .Row = rCnt: .Col = iCnt
                    iresult = .text
                    
                    With spdResult1
                        .Row = gspdResultRow:  .Col = sResultPos + tCol
                        If Len(Trim(iresult)) <> 0 Then
                            .text = iresult
                        End If
                    End With
                    tCol = tCol + 1
                Next rCnt
            Next iCnt
        End With
    End If

End Sub


Private Sub spdRstview_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
   Dim objResult As clsResult
   Dim lngCol As Long
   
   If gspdResultRow = 0 Then Exit Sub
   
   If 2280 >= X And X >= 1410 Then
      lngCol = 2
   ElseIf 4125 >= X And X >= 3210 Then
      lngCol = 4
   ElseIf 5055 >= X And X >= 5955 Then
      lngCol = 8
   ElseIf 6885 >= X And X >= 7755 Then
      lngCol = 8
   Else
      lngCol = 9
   End If

   If y < 330 Then Exit Sub

   Select Case lngCol
      Case 2, 4, 6, 8
        spdRstview_TextTipFetch lngCol, gspdResultRow, 1, 6500, "", True
      Case Else
        Exit Sub
   End Select
   
End Sub


Private Sub spdRstview_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    
    Dim pDate, pPtnm, pPtno, pSex, pPos As String
    
    With spdResult1
        .Row = gspdResultRow
        .Col = 2: pDate = .text
        .Col = 4: pPtnm = .text
        .Col = 5: pSex = .text
        .Col = 6: pPtno = .text
        .Col = 7: pPos = .text
    End With
            
'    Debug.Print pDate, pPtnm, pPtno, pSex, pPos
            
    With spdRstview
        .Row = Row
        .Col = Col
         MultiLine = 1
         TipWidth = 3000
         .SetTextTipAppearance "����ü", 9, False, False, &HEEFDF2, vbBlack
         .TextTip = TextTipFloating
         
    
         .SetTextTipAppearance "����ü", 9, False, False, &HEEFDF2, vbBlue
         
         TipText = "" & vbNewLine & _
                   "   �� ó������ ; " & pDate & vbNewLine & _
                   "   �� ȯ �� �� ; " & pPtnm & vbNewLine & _
                   "   �� ���Ϲ�ȣ ; " & pPtno & vbNewLine & _
                   "   �� ��    �� ; " & pSex & vbNewLine & vbNewLine & _
                   "   �� �˻� POS ; " & pPos & vbNewLine
                   
         ShowTip = True
       
    End With
End Sub


Private Sub spdRstview_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
    
    Set oMenu = New cPopupMenu
    
    lMenuChosen = oMenu.Popup(" �� ���� ȯ��", "-", " �� ���� ȯ��")

    Select Case lMenuChosen
        Case 1
            With spdResult1
                Col = .ActiveCol
                Row = .ActiveRow
            End With
            
            If gspdResultRow >= 1 Then
                Call spdResult1_Click(Col, gspdResultRow - 1)
            ElseIf gspdResultRow = 0 Then
                MsgBox "�� ó�� �ڷ��Դϴ�." & Space(5), vbOKOnly + vbInformation, App.Title
            Else: Exit Sub
            End If
            
        Case 3
            With spdResult1
                Col = .ActiveCol
                Row = .ActiveRow
            End With
            
            If gspdResultRow < spdResult1.maxrows Then
                Call spdResult1_Click(Col, gspdResultRow + 1)
            ElseIf gspdResultRow = spdResult1.maxrows Then
                MsgBox "�� ������ �ڷ��Դϴ�." & Space(5), vbOKOnly + vbInformation, App.Title
            Else: Exit Sub
    
            End If
    
    End Select


End Sub

Private Sub cmdNext_Click()
Dim Col, Row As Integer
    
    With spdResult1
        Col = .ActiveCol
        Row = .ActiveRow
    End With
    
    If gspdResultRow < spdResult1.maxrows Then
        Call spdResult1_Click(Col, gspdResultRow + 1)
    ElseIf gspdResultRow = spdResult1.maxrows Then
        MsgBox "�� ������ �ڷ��Դϴ�." & Space(5), vbOKOnly + vbInformation, App.Title
    Else: Exit Sub
    
    End If
    
End Sub

Private Sub cmdPrevious_Click()
Dim Col, Row As Integer
    With spdResult1
        Col = .ActiveCol
        Row = .ActiveRow
    End With
    
    If gspdResultRow >= 1 Then
        Call spdResult1_Click(Col, gspdResultRow - 1)
    ElseIf gspdResultRow = 0 Then
        MsgBox "�� ó�� �ڷ��Դϴ�." & Space(5), vbOKOnly + vbInformation, App.Title
    Else: Exit Sub
    End If
End Sub

Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol1 As Integer
    Dim intCol2 As Integer
    Dim intRow1 As Integer
    Dim intRow2 As Integer
    Dim iCnt    As Integer
    
    If Row = 0 Then
        gspdResultRow = 0:        Exit Sub
    Else
        gspdResultRow = Row
    End If
    
    intCol1 = 10
    intCol2 = 2
    intRow1 = 1
    
    With spdResult1
        For iCnt = intCol1 To .MaxCols
            .Row = Row
            .Col = intCol1
            
            spdRstview.Row = intRow1
            spdRstview.Col = intCol2
            spdRstview.BackColor = vbWhite
            
            If .BackColor = &HC6FEFF Then
                spdRstview.BackColor = &HC6FEFF
            Else
                spdRstview.BackColor = &H80000005
            End If
            
            spdRstview.text = .text
            
            intRow1 = intRow1 + 1
            intCol1 = intCol1 + 1
            
            If intRow1 > spdRstview.maxrows Then
                intRow1 = 1
                intCol2 = intCol2 + 2
            End If

        Next
    End With
    

End Sub

'
' END
'

Private Sub spdResult1_KeyPress(KeyAscii As Integer)

    Dim arow    As Integer, aCOL   As Integer
    Dim varChk  As Variant, varBar As Variant, varNum As Variant
    Dim iRow    As Integer, iCnt   As Integer
    
    'Debug.Print Col & NewCol & Row & NewRow
       
    If KeyAscii = vbKeyReturn Then
        With spdResult1
            aCOL = .ActiveCol
            arow = .ActiveRow
            If aCOL = 4 Then
                iCnt = 0
                For iRow = arow To .maxrows
                    .GetText 1, iRow, varChk
                    .GetText 3, iRow, varBar
                    .GetText aCOL, arow, varNum
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
        .SortKey(1) = Col       '����Ű ����ȣ

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


Private Sub Timer1_Timer()

    comEQP.Output = EOT
    Timer1.Enabled = False
    Debug.Print "1200>>" & EOT

End Sub

'Private Sub Timer2_Timer()
'    comEQP.Output = ACK
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

Private Sub txtBarCode_Change()

    If txtBarCode.SelStart = txtBarCode.MaxLength Then SendKeys "{TAB}"
    
End Sub

Private Sub txtBarCode_GotFocus()

    With txtBarCode
        .SelStart = 0
        .SelLength = Len(.text)
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
    
    If txtBarCode.text = "" Then Exit Sub
    
    blnFlag = False
    If KeyAscii = vbKeyReturn Then
        intCol = sl_examdata_select&(txtBarCode.text, INS_CODE, strEqcode, strExamname, strOrdcd, strPid, strPnm, strAcptno)
        
        For intCol = 0 To UBound(strOrdcd)
            If strOrdcd(intCol) <> "" Then
                strEqpCd = f_funGet_CODE(strOrdcd(intCol))
                Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                If Not itemX Is Nothing Then
                    If Not blnFlag Then
                        intRow = f_funGet_SpreadRow(spdResult1, 2, txtBarCode.text)
                        If intRow < 1 Then
                            intRow = f_funGet_SpreadRow(spdResult1, 2, "")
                            If intRow < 1 Then
                                spdResult1.maxrows = spdResult1.maxrows + 1
                                spdResult1.RowHeight(spdResult1.maxrows) = 13
                                intRow = spdWorklist.maxrows
                            End If
                            spdResult1.SetText 2, intRow, txtBarCode.text
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
    
        If Not blnFlag Then MsgBox "�ش� �˻��׸��� �������� ���� ��ü�Դϴ�.", vbInformation, App.Title
        
        txtBarCode.text = "":   txtBarCode.SetFocus
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
            If Trim(.text) = Mid(txtBarCode.text, 1, 11) Then
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
'        MsgBox "�ش� �˻��׸��� �������� ���� ��ü�Դϴ�.", vbInformation, Me.Caption
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

Private Sub txtChart_GotFocus()
'
' Focus ������ ���
'
    txtChart.ForeColor = &HFF&
    txtChart.text = ""
End Sub

Private Sub txtChart_LostFocus()
'
' Focus �� ���� ���
'
    txtChart.ForeColor = &HFFC0C0
    txtChart.text = "��Ʈ��ȣ �Է�"
End Sub

Private Sub txtChart_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim intRow2 As Integer
    
    Dim tBlood As Boolean
        
    If Len(Trim(txtChart)) > 0 Then
        If KeyCode = 13 Then
        
          tBlood = False
          
          Rem txtChart = Format(txtChart, "0000000")
          
          intRow2 = f_funGet_SpreadRow(spdWorklist, 5, txtChart)
          
          If intRow2 >= 1 Then
              
              With spdWorklist
                .SetText 1, intRow2, "1"
                cmdWorkList_Click
                txtChart.text = ""
                tBlood = True
              End With
          End If
          
          If tBlood = False Then
            MsgBox txtChart.text & " �ش� ȯ���� ó���� �����ϴ�.     ", vbInformation + vbOKOnly, App.Title
            txtChart.text = ""
          End If
        
         End If
    End If

End Sub

' ------------------------------------------------------------------------
' ��Ż��� Ȯ�� �����̺�Ʈ
' ------------------------------------------------------------------------
Private Sub txtCom_Change()
    txtCom.SelStart = Len(txtCom.text)
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
              Format(Date, "YYYY�� MM�� DD��") & "  "; Time & vbNewLine & _
              "������������������������������������������������������������" & vbNewLine & _
              txtCom.text & _
              "������������������������������������������������������������" & vbNewLine
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

    Call ComReceive(txtCom.SelText)
    
End Sub

Private Sub cmdCOMInput2_Click()
    
    Dim bytTemp() As Byte
    
    If txtCOM2.SelLength = 0 Then
        bytTemp = StrConv(charCOM_Convert(txtCOM2.text), vbFromUnicode)
    Else
        bytTemp = StrConv(charCOM_Convert(txtCOM2.SelText), vbFromUnicode)
    End If

    Call ComReceive(txtCOM2.SelText)

End Sub

Private Sub cmdCOMOutput2_Click()
    
    If txtCOM2.SelLength = 0 Then
        Call COM_OUTPUT(charCOM_Convert(txtCOM2.text))
    Else
        Call COM_OUTPUT(charCOM_Convert(txtCOM2.SelText))
    End If
    
End Sub
' ------------------------------------------------------------------------
' ��Ż��� Ȯ�� �����̺�Ʈ


Private Sub txtResult_DblClick()
    txtResult.text = ""
    List1.text = ""
    
    If txtResult.Visible Then txtResult.Visible = False
    List1.Visible = True
End Sub



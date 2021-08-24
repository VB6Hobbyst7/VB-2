VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmComm 
   Caption         =   "Interface"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   11985
   WindowState     =   2  '최대화
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   90
      Top             =   6570
   End
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
            Picture         =   "frmComm_YEJ.frx":0000
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_YEJ.frx":059A
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_YEJ.frx":0B34
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_YEJ.frx":10CE
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_YEJ.frx":1668
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_YEJ.frx":1C02
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm comEQP 
      Left            =   3795
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
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
            Picture         =   "frmComm_YEJ.frx":219C
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_YEJ.frx":2736
            Key             =   "NOT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_YEJ.frx":2CD0
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_YEJ.frx":326A
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_YEJ.frx":3AFC
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_YEJ.frx":3C56
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_YEJ.frx":3DB0
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
      Begin HSCotrol.CButton CButton1 
         Height          =   360
         Left            =   2295
         TabIndex        =   57
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
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
         BorderColor     =   8421504
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
         Picture         =   "frmComm_YEJ.frx":3F0A
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   9780
         Picture         =   "frmComm_YEJ.frx":4494
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   8640
         Picture         =   "frmComm_YEJ.frx":4A1E
         Top             =   255
         Width           =   240
      End
   End
   Begin TabDlg.SSTab tabWork 
      Height          =   5850
      Left            =   45
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
      TabPicture(0)   =   "frmComm_YEJ.frx":4FA8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "pnlCom2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "pnlCom"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "spdWorkList"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSel(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdWorkList"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdWordQuery"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdAppend(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "mskOrdDate"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboRstgbn(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtBarCode"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdSel(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "optBar"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "optSeq"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdStartNo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkAuto"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "spdResult1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "mskOrdDate1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "SSPanel1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdOrder"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   " 받은 결과"
      TabPicture(1)   =   "frmComm_YEJ.frx":4FC4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwCuData"
      Tab(1).Control(1)=   "cboRstgbn(1)"
      Tab(1).Control(2)=   "mskRstDate"
      Tab(1).Control(3)=   "cmdAppend(1)"
      Tab(1).Control(4)=   "cmdRstQuery"
      Tab(1).Control(5)=   "cmdSel(3)"
      Tab(1).Control(6)=   "cmdSel(2)"
      Tab(1).Control(7)=   "spdResult2"
      Tab(1).Control(8)=   "Label4"
      Tab(1).ControlCount=   9
      Begin HSCotrol.CButton cmdOrder 
         Height          =   300
         Left            =   8370
         TabIndex        =   56
         Top             =   495
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "오더전송"
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
      Begin Threed.SSPanel SSPanel1 
         Height          =   195
         Left            =   2430
         TabIndex        =   54
         Top             =   540
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   344
         _StockProps     =   15
         Caption         =   "-"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
      End
      Begin MSMask.MaskEdBox mskOrdDate1 
         Height          =   300
         Left            =   2655
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
      Begin FPSpread.vaSpread spdResult1 
         Height          =   4875
         Left            =   2925
         TabIndex        =   52
         Top             =   900
         Width           =   8925
         _Version        =   196608
         _ExtentX        =   15743
         _ExtentY        =   8599
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   3
         EditEnterAction =   2
         EditModePermanent=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmComm_YEJ.frx":4FE0
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
         Left            =   6840
         TabIndex        =   51
         Top             =   0
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TEST"
         Height          =   285
         Left            =   5580
         TabIndex        =   50
         Top             =   495
         Visible         =   0   'False
         Width           =   960
      End
      Begin HSCotrol.CButton cmdStartNo 
         Height          =   300
         Left            =   7125
         TabIndex        =   49
         Top             =   495
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
         Left            =   11100
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   4920
         Left            =   -67980
         TabIndex        =   44
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
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   1
         Left            =   360
         TabIndex        =   25
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_YEJ.frx":53E6
      End
      Begin VB.ComboBox cboRstgbn 
         Height          =   300
         Index           =   1
         ItemData        =   "frmComm_YEJ.frx":5868
         Left            =   -72570
         List            =   "frmComm_YEJ.frx":5875
         Style           =   2  '드롭다운 목록
         TabIndex        =   15
         Top             =   495
         Width           =   1770
      End
      Begin VB.TextBox txtBarCode 
         Height          =   300
         Left            =   3720
         MaxLength       =   12
         TabIndex        =   13
         Top             =   -45
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.ComboBox cboRstgbn 
         Height          =   300
         Index           =   0
         ItemData        =   "frmComm_YEJ.frx":589F
         Left            =   5235
         List            =   "frmComm_YEJ.frx":58AC
         Style           =   2  '드롭다운 목록
         TabIndex        =   12
         Top             =   -45
         Visible         =   0   'False
         Width           =   1500
      End
      Begin MSMask.MaskEdBox mskRstDate 
         Height          =   300
         Left            =   -73695
         TabIndex        =   16
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
         TabIndex        =   17
         Top             =   495
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
      Begin HSCotrol.CButton cmdRstQuery 
         Height          =   300
         Left            =   -65415
         TabIndex        =   18
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
      Begin MSMask.MaskEdBox mskOrdDate 
         Height          =   300
         Left            =   1305
         TabIndex        =   19
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
         Index           =   0
         Left            =   10620
         TabIndex        =   20
         Top             =   495
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
      Begin HSCotrol.CButton cmdWordQuery 
         Height          =   300
         Left            =   9585
         TabIndex        =   21
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
      Begin HSCotrol.CButton cmdWorkList 
         Height          =   300
         Left            =   105
         TabIndex        =   24
         Top             =   5460
         Width           =   2805
         _ExtentX        =   4948
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
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   0
         Left            =   90
         TabIndex        =   26
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_YEJ.frx":58D6
      End
      Begin FPSpread.vaSpread spdWorkList 
         Height          =   4515
         Left            =   90
         TabIndex        =   14
         Top             =   900
         Width           =   2850
         _Version        =   196608
         _ExtentX        =   5027
         _ExtentY        =   7964
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
         MaxCols         =   4
         MaxRows         =   1
         ScrollBarMaxAlign=   0   'False
         SpreadDesigner  =   "frmComm_YEJ.frx":5D44
         UserResize      =   0
      End
      Begin HSCotrol.UserPanel pnlCom 
         Height          =   5355
         Left            =   45
         TabIndex        =   27
         Top             =   495
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
            TabIndex        =   33
            Top             =   270
            Width           =   11595
         End
         Begin VB.Frame Frame1 
            Height          =   645
            Left            =   45
            TabIndex        =   28
            Top             =   4650
            Width           =   11610
            Begin HSCotrol.CButton cmdCOMSave 
               Height          =   360
               Left            =   10515
               TabIndex        =   29
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
               TabIndex        =   30
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
               TabIndex        =   31
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
               TabIndex        =   32
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
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   3
         Left            =   -74640
         TabIndex        =   45
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_YEJ.frx":609C
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   -74910
         TabIndex        =   46
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_YEJ.frx":651E
      End
      Begin HSCotrol.UserPanel pnlCom2 
         Height          =   4785
         Left            =   5895
         TabIndex        =   34
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
            TabIndex        =   43
            Top             =   270
            Width           =   5730
         End
         Begin VB.Frame Frame2 
            Height          =   645
            Left            =   90
            TabIndex        =   35
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
               TabIndex        =   36
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
               TabIndex        =   37
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
               TabIndex        =   38
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
               TabIndex        =   39
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
               TabIndex        =   40
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
               TabIndex        =   41
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
               TabIndex        =   42
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
      Begin FPSpread.vaSpread spdResult2 
         Height          =   4830
         Left            =   -74910
         TabIndex        =   53
         Top             =   900
         Width           =   11670
         _Version        =   196608
         _ExtentX        =   20585
         _ExtentY        =   8520
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   7
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         SpreadDesigner  =   "frmComm_YEJ.frx":698C
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   570
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
Private Type typeVECI
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

Dim VECI As typeVECI
Dim fVECI(100) As String
Dim fVECI_1(100) As String
Dim SendData(10)     As String
Dim SendCount        As String
Dim Or_Seq           As Integer
Dim SendBuffW           As String
Dim SendBuffT           As String

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

'Dim fVECI(100) As String
'Dim fVECI_1(100) As String
Dim fVECICfg(100) As Integer
Dim fVECISize(100, 1) As Integer
Dim fChannel() As String

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
'            & " Where c.RE_DATE = '" & strDate & "' And a.IN_CODE like 'SE%' And a.EX_INST < '2' And a.RE_RCID = c.RE_RCID And b.JU_PERID = c.JU_PERID"

    '----용인시 수지출장소(출장소코드 추가) HE_UNID = 'HS-44906' ----
    gSql = "select a.RE_RCID,b.JU_NAME,b.JU_PERID from EXAM_TOC A,JUMN_TMA B,RECE_TJU C" _
            & " Where c.RE_DATE >= '" & strDate & "' And c.RE_DATE <= '" & strDate1 & "'" _
            & " And b.HE_UNID = 'HS-44906' And a.IN_CODE = 'SE00125' And a.EX_INST = '1'" _
            & " And a.RE_RCID = c.RE_RCID And b.JU_PERID = c.JU_PERID order by a.RE_RCID"
    
    AdoRs_ORACLE.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
   
    If AdoRs_ORACLE.RecordCount = 0 Then
        Set f_subSet_WorkList = Nothing
    Else
        Set f_subSet_WorkList = AdoRs_ORACLE
    End If

    Set AdoRs_ORACLE = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)

    
End Function

Private Sub f_subSet_ItemList()

    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strTest As String, intPos   As Integer
    Dim strTmp  As String, intCol   As Integer, intCnt  As Integer
    
'On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = ""
    
    intCol = 7
    With spdWorkList
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 15
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 15
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdResult2
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 15
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
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        With spdResult2
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        fChannel(intCol - 6) = adoRS.Fields("TEST_EQP")
        
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

Private Sub CButton1_Click()

    Call COM_OUTPUT(EOT)
    
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

End Sub

Private Sub cmdClear()
    
    f_strJOB_FLAG = "1"
    f_intSampleNo = 0
    Or_Seq = 1
    
    With spdWorkList
        .maxrows = 14
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdResult1
        .maxrows = 15
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
        .RowHeight(-1) = 12
    End With

    With spdResult2
        .maxrows = 15
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
    
    Dim strOrdLst()     As String, strPid()    As String, strPnm() As String
    
    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
    Dim itemX   As ListItem
    Dim objSpd  As vaSpread
    Dim sqlRet  As Integer
    Dim flgSave As Boolean
    
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
            .GetText 2, intRow, varTmp:         strBarno = Trim$(varTmp)
'            .GetText .MaxCols, intRow, varTmp:  strTime = Trim$(varTmp)
            
            .GetText 1, intRow, varTmp
            
            If strBarno = "" Then Exit For
            
            intCnt = 0: Erase strOrdcd: Erase strRstval
            If Trim$(varTmp) = "1" Then
                For intCol = 7 To .MaxCols
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
                                    
                                    gSql = "update EXAM_TOC set EX_INRV = '" & Trim$(varTmp) & "',EX_INST = '2',EX_DATE = '" & Format$(Now, "YYYYMMDD") & "',EX_INEM='1271'" _
                                           & " where RE_RCID ='" & strBarno & "' And IN_CODE='" & strTestcd & "'"
                                    
                                    AdoCn_ORACLE.Execute (gSql)
                                    lblStatus.Caption = "저장 성공!!"

                                Loop
                            Else
                                blnFlag = False
                                Set mAdoRs = f_subSet_TestList(strBarno)
                                Do Until mAdoRs.EOF
                                    If mAdoRs("in_code") = strTestcd Then blnFlag = True: Exit Do
                                    mAdoRs.MoveNext
                                Loop
                                Set adoRS = Nothing: mAdoRs.Close
                                
                                If blnFlag Then
                                    gSql = "update EXAM_TOC set EX_INRV = '" & Trim$(varTmp) & "',EX_INST = '2',EX_DATE = '" & Format$(Now, "YYYYMMDD") & "',EX_INEM='1271'" _
                                           & " where RE_RCID ='" & strBarno & "' And IN_CODE='" & strTestcd & "'"
                                    
                                    AdoCn_ORACLE.Execute (gSql)
                                    lblStatus.Caption = "저장 성공!!"
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

Private Sub cmdOrder_Click()

    Call COM_OUTPUT(ENQ)
    
End Sub

Private Sub cmdRstQuery_Click()

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String, intRet   As Integer
    
    Dim strSpcno, strBarno, strEqpCd   As String
    Dim intRow      As Integer, intCol  As Integer
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
    Dim introw1, intRow2 As Integer
    Dim itemX       As ListItem
    
    intRow = 0
    With spdResult2
        .maxrows = 15
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    sqlDoc = "select SPCNO, TESTCD, EQUIPCD, TRANSTM, RSTVAL, REFVAL, TRANSDT, EQPNUM" & _
             "  from INTERFACE003" & _
             " where TRANSDT = '" & mskRstDate.Text & "'" & _
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
                    .SetText 2, intRow, Trim$(adoRS(0) & "")
                    
                    Set mAdoRs = f_subSet_SearchList(Trim$(adoRS(0)))
                    .SetText 3, intRow, mAdoRs("ju_name")
                    .SetText 4, intRow, mAdoRs("ju_perid")
                    .SetText .MaxCols, intRow, Trim$(adoRS(6) & "")
                End If
                
                strSpcno = Trim$(adoRS(0) & "")
                Set itemX = lvwCuData.FindItem(Trim$(adoRS(7) & ""), lvwTag, , lvwWhole)
                If Not itemX Is Nothing Then
                    intCol = itemX.Index + 6
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
            For intRow = 1 To .maxrows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 2, "1", "")
            Next
        End With
    Else
        With spdWorkList
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
    
    Set AdoRs_ORACLE = New ADODB.Recordset
'    gSql = "select IN_CODE from EXAM_TOC Where RE_RCID = '" & strBarcode & "'"

    gSql = "select a.IN_CODE,a.RE_RCID,b.JU_NAME,b.JU_PERID from EXAM_TOC A,JUMN_TMA B,RECE_TJU C" _
            & " Where a.RE_RCID = '" & strBarcode & "' And b.HE_UNID = 'HS-44906'" _
            & " And a.IN_CODE like 'SE%' And a.RE_RCID = c.RE_RCID" _
            & " And b.JU_PERID = c.JU_PERID and a.EX_INST = '1'"
            
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

Private Function f_subSet_SearchList(ByVal strBarcode As String)
    Dim sqlRet      As Integer
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
    
    Set AdoRs_ORACLE = New ADODB.Recordset
'    gSql = "select IN_CODE from EXAM_TOC Where RE_RCID = '" & strBarcode & "'"
    
    gSql = "select a.IN_CODE,a.RE_RCID,b.JU_NAME,b.JU_PERID from EXAM_TOC A,JUMN_TMA B,RECE_TJU C" _
            & " Where a.RE_RCID = '" & strBarcode & "'" _
            & " And b.HE_UNID = 'HS-44906'" _
            & " And a.RE_RCID = c.RE_RCID" _
            & " And b.JU_PERID = c.JU_PERID " _
            & " And a.EX_INST = '2'" _
            & " And a.IN_CODE like 'SE%'"

    AdoRs_ORACLE.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
    
    If AdoRs_ORACLE.RecordCount = 0 Then
        Set f_subSet_SearchList = Nothing
    Else
        Set f_subSet_SearchList = AdoRs_ORACLE
    End If

    Set AdoRs_ORACLE = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)

    
End Function
Private Sub cmdWordQuery_Click()
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
        .maxrows = 14
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
                
                .SetText 1, intRow, "1"
                .SetText 2, intRow, mAdoRs("JU_NAME")
                .SetText 3, intRow, mAdoRs("JU_PERID")
                .SetText 4, intRow, mAdoRs("RE_RCID")

                '-- 검사항목조회
                Set mAdoRs1 = New Recordset
                Set mAdoRs1 = f_subSet_TestList(mAdoRs("RE_RCID"))
                
                Do Until mAdoRs1.EOF
                    strEqpCd = f_funGet_CODE(mAdoRs1("IN_CODE"))
                    
                    Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                    If Not itemX Is Nothing Then .SetText 4 + itemX.Index, intRow, "V"
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

Private Sub cmdWorkList_Click()

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
        For introw1 = 1 To .maxrows
            .GetText 1, introw1, varTmp
            If Trim$(varTmp) = "1" Then
                .GetText 2, introw1, varTmp:    strSPnm = Trim$(varTmp)
                .GetText 3, introw1, varTmp:    strSPid = Trim$(varTmp)
                .GetText 4, introw1, varTmp:    strBarno = Trim$(varTmp)
                
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
                    If Len(strBarno) = 16 Then
                        Do Until mAdoRs.EOF
                            strEqpCd = f_funGet_CODE(mAdoRs("in_code"))
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
                        spdResult1.maxrows = spdResult1.maxrows - 1
                    End If
                End If
                spdResult1.SetText 1, intRow2, "1"
                spdResult1.maxrows = intRow2

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
            For iRow = 1 To .maxrows
                .SetText 5, iRow, varNum
                .SetText 6, iRow, ((iCnt Mod 10) + 1) - 1
                iCnt = iCnt + 1
                If (iCnt Mod 10) = 0 Then varNum = varNum + 1
            Next
        End If
    End With
    
End Sub

Private Sub comEQP_OnComm()
    
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim strDta      As String
    Dim Arr()       As Byte
   
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
    
'------------------------------세미양방향일 경우 사용함---------------------------------------------
'---------------------------------------------------------------------------------------------------
Dim strRec  As String, strBuff  As String
    Dim strTmp  As String, intIdx   As Integer
    Dim intPos0 As Integer, intPos1 As Integer, intPos2 As Integer

    Dim AGE As String
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
'    Dim sStxCheck As Integer
'    Dim sEtxCheck As Integer
    Dim SendBuffD           As String           'data
    
    Dim Lencheck
    Dim Specimen            As String

    Static OrgMsg As String

    strRec = RecData
    Print #1, strRec;

    Debug.Print strRec

    For intIdx = 1 To Len(strRec)
        strBuff = Mid$(strRec, intIdx, 1)
        Select Case strBuff
            Case STX
                    sStxCheck = InStr(strBuff, STX)
            Case ETX
                    sEtxCheck = InStr(strBuff, ETX)
                    If sStxCheck <> 0 And sEtxCheck <> 0 Then
                        Call psDataDefine(f_strBuffer, fChannel(), spdResult1)
                        Call COM_OUTPUT(ACK)
                        GoSub ClearReceiveData
                     End If
            Case ETB
                    If Mid(f_strBuffer, intIdx, 2) = vbCrLf Then
                        f_strBuffer = left(f_strBuffer, Len(f_strBuffer) - 2) 'Remove CR & LF
                    End If
                    cntCheckSum = cntCheckSum + 1
                    flgETB = True
            Case vbCr
'                    f_strBuffer = f_strBuffer + strBuff
'                    Call COM_OUTPUT(ACK)
            Case vbLf
'                    Call COM_OUTPUT(ACK)
'                    f_strBuffer = f_strBuffer + strBuff
            Case ENQ
                    Call COM_OUTPUT(ACK)
            Case ACK
                    Dim varTmp      As Variant
                    Dim intRow      As Integer, intCol  As Integer
                    Dim strBarno    As String, strTest  As String
                    Dim strRack     As String, strCup   As String
                    Dim intCnt      As Integer
                    Dim itemX       As ListItem

                    With spdResult1
                        For intRow = 1 To .maxrows
                            .Row = intRow
                            .Col = 2
                            If .BackColor = vbWhite Then
                                sAppCode = ""
                                intCnt = 0
                                .GetText 2, intRow, varTmp: strBarno = Trim$(varTmp)
                                .GetText 5, intRow, varTmp: strRack = Trim$(varTmp)
                                .GetText 6, intRow, varTmp: strCup = Trim$(varTmp)
                                sRackId = Format$(strRack, "00")
'                                .GetText 1, intRow, varTmp
                                For intCol = 7 To .MaxCols
                                    spdResult1.GetText intCol, 0, varTmp
                                    If Trim$(varTmp) = "" Then Exit For
                                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                    If Not itemX Is Nothing Then
                                        spdResult1.Col = intCol:    'spdResult1.Row = OrderCnt
                                        If spdResult1.BackColor = &HC6FEFF Then
                                            intCnt = intCnt + 1
                                            SendBuffD = SendBuffD & "^^^1.000000+" & Trim(itemX.tag) & "+1" & "\"
                                        End If
                                    End If
                                    Set itemX = Nothing
                                Next intCol
                                If Or_Seq = 5 And SendBuffD <> "" Then
                                    .Row = intRow
                                    .Col = 2: .BackColor = vbCyan
                                    .Col = 3: .BackColor = vbCyan
                                    .Col = 4: .BackColor = vbCyan
                                End If
                                Exit For
                            End If
                        Next intRow
                        
                        If intRow >= .maxrows Then
                            Timer1.Enabled = False
                        End If
                    End With
                    
                    Select Case Or_Seq
                           Case 1   ' Send Header
                                    sSDateTime = Format(Now, "YYYYMMDDHHMMSS")
                                    SendBuffW = Or_Seq & "H|\^&||||||||||||" & sSDateTime & vbCr & ETX
                                    SendBuffT = STX & SendBuffW & CheckSum_ECi_Tx(SendBuffW) & vbCr & vbLf
                                    Call COM_OUTPUT(SendBuffT)
                                    Debug.Print SendBuffT
                                    Or_Seq = Or_Seq + 1

                           Case 2   ' Send Patient Information
                                    SendBuffW = Or_Seq & "P|1||||||" & "19000000" & vbCr & ETX
                                    SendBuffT = STX & SendBuffW & CheckSum_ECi_Tx(SendBuffW) & vbCr & vbLf
                                    Call COM_OUTPUT(SendBuffT)
                                    Debug.Print SendBuffT

                                    Or_Seq = Or_Seq + 1

                           Case 3   ' Send Order Record
                                    SendBuffW = Or_Seq & "O|1|" & intRow & "^" & sRackId & "^" & strCup & "||" & SendBuffD & "|" & "R" & "||||||N||||" & "4" & "||||||||||F" & vbCr & ETX

                                    SendBuffT = STX & SendBuffW & CheckSum_ECi_Tx(SendBuffW) & vbCr & vbLf
 
                                    Call COM_OUTPUT(SendBuffT)
                                    Debug.Print SendBuffT
                                    Or_Seq = Or_Seq + 1

                           Case 4   ' Send Message Terminator
                                    SendBuffW = Or_Seq & "L|1|N" & vbCr & ETX
                                    SendBuffT = STX & SendBuffW & CheckSum_ECi_Tx(SendBuffW) & vbCr & vbLf
                                    Call COM_OUTPUT(SendBuffT)
                                    Debug.Print SendBuffT
                                    Or_Seq = Or_Seq + 1
                           Case 5   ' Send EOT
                                    Or_Seq = 1
                                    SendBuffD = ""
                                    Call COM_OUTPUT(EOT)
                                    Debug.Print EOT
                                    If intRow <= spdResult1.maxrows Then
                                        Timer1.Interval = 100
                                        Timer1.Enabled = True
                                    End If

                    End Select
            
            Case NAK
'                    Call COM_OUTPUT(ACK)
            Case EOT
'                    Call COM_OUTPUT(ENQ)
                    GoSub ClearReceiveData
            Case Else
                    f_strBuffer = f_strBuffer + strBuff
        End Select
    Next

    Exit Sub
ClearReceiveData:
    sStxCheck = 0
    sEtxCheck = 0
    ReceiveData = ""
    cntField_ = 0
    cntRepeat_ = 0
    cntComponent_ = 0
    cntEscape_ = 0
    cntSlash_ = 0
    f_strBuffer = ""
'    Return

''------------------------------단방향일 경우 사용함---------------------------------------------
''-----------------------------------------------------------------------------------------------
'    Dim strRec  As String, strBuff  As String
'    Dim strTmp  As String, intIdx   As Integer
'    Dim intPos0 As Integer, intPos1 As Integer, intPos2 As Integer
'
'    Dim sStxCheck As Integer
'    Dim sEtxCheck As Integer
'    Dim sStartCheck As Integer
'    Dim sEndCheck As Integer
'
'    Static OrgMsg As String
'    strRec = RecData ' StrConv(RecData, vbUnicode)
'    f_strBuffer = ""
'
'    Print #1, strRec;
'    strTmp = strRec
'    Call COM_INPUT(strTmp)
'    Debug.Print strRec
'    For intIdx = 1 To Len(strRec)
'        strBuff = Mid$(strRec, intIdx, 1)
'        f_strBuffer = f_strBuffer + strBuff
'        sStartCheck = InStr(f_strBuffer, "!000a02")
'        sEndCheck = InStr(f_strBuffer, "h")
'        If sStartCheck <> 0 And sEndCheck <> 0 Then
'            f_strBuffer = f_strBuffer + strBuff
'            strTmp = f_strBuffer
'            Call psDataDefine(strTmp, fChannel(), spdResult1)
'            f_strBuffer = ""
'        End If
'     Next
            
End Sub

Private Sub psDataDefine(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object) ', ByVal brOst As String) ' ByRef brItemdeci() As String)

    Dim ii As Long
    Dim jj As Long
    Dim KK As Long
    Dim Found As Boolean
    Dim SID As String
    Dim SEX As String
    Dim AGE As Long
    Dim sql As String
    Dim OutPutData As String, sOrderLst As String
    Dim sTemp       As String       ' On Com으로부터 넘겨받은 Receive Data
    Dim Channel_No  As String       ' 문자형 변수
    Dim Patiant_No  As String       ' 환자번호
    Dim pGrid_Point As Integer      ' 해당 검사자 Point
    Dim Max_Arary_Cnt As Integer    ' 검사 항목수
    '-------------------------------' 임시 변수들.....
    Dim itemX As ListItem
    Dim sDeCnt      As Integer
    Dim strRstval As String, strRefVal  As String
    Dim sqlDoc  As String
    Dim sSeq, strTmp, varTmp, strBarno, strDate, strTime As String
    Dim pDoCount    As Integer
    Dim intRow, intIdx As Integer
    Dim Loop_count  As Integer
    Dim intCol, sCol As Integer
    Dim sRtn As Integer, sChannel As String, sRstText As String, sRstValue As Single, sUnit As String
    Dim sPatiant_No As Long
    
    
'    Debug.Print "받은데이타:" & ReceiveData
    On Error GoTo errReceive
    ReceiveData = strdata
    Select Case Mid(ReceiveData, 2, 1)
        
        Case "H" 'Message Header
            GoSub Clear_VECI_
        Case "P" 'Patient Informatioin
            GoSub Clear_VECI_
            '
        Case "O" 'Test Order
            For ii = 1 To Len(ReceiveData)
                
                Select Case Mid(ReceiveData, ii, 1)
                    Case Field_ '|
                        cntField_ = cntField_ + 1
                        cntRepeat_ = 0
                        cntComponent_ = 0
                    Case Repeat_ '\
                        cntRepeat_ = cntRepeat_ + 1
                        cntComponent_ = 0
                    Case Component_ '^
                        cntComponent_ = cntComponent_ + 1
                    Case Slash_ '/
                        '
                    Case Else
                        Select Case cntField_ '|
                            Case 2
                                Select Case cntComponent_ '^
                                    Case 0
                                        VECI.SampleNo = VECI.SampleNo & Mid(ReceiveData, ii, 1)
                                    Case 1
                                        VECI.SID = VECI.SID & Mid(ReceiveData, ii, 1)
                                    Case 2
                                        VECI.SampleTy = VECI.SampleTy & Mid(ReceiveData, ii, 1)
                                    Case 3
                                        VECI.RackNo = VECI.RackNo & Mid(ReceiveData, ii, 1)
                                    Case 4
                                        VECI.Position = Mid(ReceiveData, ii, 1)
                                    Case 6
                                        VECI.Priority = Mid(ReceiveData, ii, 1)
                                End Select
                            Case 4
                                Select Case cntComponent_ Mod 3 '^
                                    Case 0
                                        If IsNumeric(Mid(ReceiveData, ii, 1)) Then
                                            VECI.TestId(cntRepeat_) = VECI.TestId(cntRepeat_) & Mid(ReceiveData, ii, 1)
                                        End If
                                End Select
                            Case 22
                                VECI.TestDate = VECI.TestDate & Mid(ReceiveData, ii, 1)
    
                        End Select
                End Select
            Next ii
            VECI.TestTime = Mid(VECI.TestDate, 9, 6)
            VECI.TestDate = left(VECI.TestDate, 8)

            sSeq = VECI.SampleNo
            sCol = 2
            Patiant_Recevid = SeqSearch(spdResult1, sSeq, sCol)
'=============================================================================================================================================================================
        
        Case "R" 'Test Result
            Dim TestId As String
            Dim Result As String
            Dim Status As String
            Dim Rerun As String
            For ii = 1 To Len(ReceiveData)
                Select Case Mid(ReceiveData, ii, 1)
                    Case Field_ '|
                        cntField_ = cntField_ + 1
                        cntRepeat_ = 0
                        cntComponent_ = 0
                    Case Repeat_ '\
                        cntRepeat_ = cntRepeat_ + 1
                        cntComponent_ = 0
                    Case Component_ '^
                        cntComponent_ = cntComponent_ + 1
                    Case Slash_ '/
                        '
                    Case Else
                        Select Case cntField_ '|
                            Case 2
                                Select Case cntComponent_
                                    Case 3
                                        If IsNumeric(Mid(ReceiveData, ii, 1)) Then 'Rerun : inc, dec, clr
                                            TestId = TestId & Mid(ReceiveData, ii, 1)
'                                            TestId = Mid(TestId, 8, 3)
                                        End If
                                        Result = ""
                                        Status = ""
                                        Rerun = ""
                                End Select
                            Case 3
                                Result = Result & Mid(ReceiveData, ii, 1)
                            Case 6
                                Status = Status & Mid(ReceiveData, ii, 1)
                            Case 8
                                Rerun = Mid(ReceiveData, ii, 1) 'F:First, C:Rerun
                        End Select
                End Select
            Next ii
            
            Channel_No = Mid(TestId, 8, 3)
            Max_Arary_Cnt = spdResult1.MaxCols - 6
            With spdResult1
                If Patiant_Recevid > 0 Then
                    For intCol = 7 To .MaxCols
                        .GetText intCol, 0, varTmp
                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                        If Not itemX Is Nothing Then
                            For intIdx = 1 To Max_Arary_Cnt
                                If Len(Channel_No) > 0 Then
                                    If Trim(Channel_No) = itemX.tag Then
                                        strRstval = Result
                                        If IsNumeric(Result) Then
                                            strRstval = Result
                                        Else
                                            strRstval = ""
                                        End If
                                        If strRstval >= 1 Then
                                            strRstval = "양성"
                                        Else
                                            strRstval = "음성"
                                        End If

                                        strDate = Format$(Now, "YYYYMMDD"):     strTime = Format$(Now, "MMSS")
                                        .SetText intCol, Patiant_Recevid, strRstval
                                        .Col = intCol:  .Row = Patiant_Recevid
                                                        .ForeColor = IIf(Trim$(strRefVal) <> "", vbRed, vbBlack)
                                        .GetText 2, Patiant_Recevid, varTmp:   strBarno = Trim$(varTmp)
'                                        .SetText 1, Patiant_Recevid, "0"

                                        sqlDoc = "Update INTERFACE003" & _
                                                 "   set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRefVal & "'" & _
                                                 " where SPCNO   = '" & strBarno & "'" & _
                                                 "   and EQPNUM  = '" & itemX.tag & "'" & _
                                                 "   and TRANSDT = '" & strDate & "'" & _
                                                 "   and TRANSTM = '" & strTime & "'"
                                        AdoCn_Jet.Execute sqlDoc

                                        sqlDoc = "insert into INTERFACE003(" & _
                                                 "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN)" & _
                                                 "    values( '" & strBarno & "', '" & itemX.Text & "', '" & itemX.tag & "'," & _
                                                 "            '" & strDate & "', '" & strTime & "'," & _
                                                 "            '" & strRstval & "', '" & strRefVal & "'," & _
                                                 "            '" & INS_CODE & "', '')"
                                        AdoCn_Jet.Execute sqlDoc
    
                                       Exit For
                                    End If
                                End If
                            Next intIdx
                        End If
                        Set itemX = Nothing
                    Next
                End If
           End With
           
            cntRepeat_ = -1
            For ii = 0 To 100
                If VECI.TestId(ii) = TestId Then
                    cntRepeat_ = ii
                    Exit For
                End If
            Next ii
            If cntRepeat_ <> -1 Then
                VECI.TestId(cntRepeat_) = TestId
                VECI.Result(cntRepeat_) = Result
                VECI.Status(cntRepeat_) = Status
                VECI.Rerun(cntRepeat_) = Rerun
            End If
            GoSub Clear_VECI_
        Case "C" 'Comment
        
        Case "L" 'Message Termination
            GoSub Clear_VECI_
        Case "Q" 'Request Information
        
        Case Else
            '
    End Select
    Exit Sub
    
Clear_VECI_:
    VECI.TestDate = ""
    VECI.TestTime = ""
    VECI.SampleNo = ""
    VECI.SID = ""
    VECI.SampleTy = ""
    VECI.RackNo = ""
    VECI.Position = ""
    VECI.Priority = ""
    Channel_No = ""
    Result = ""
    For ii = 0 To 100
        VECI.TestId(ii) = ""
        VECI.Result(ii) = ""
        VECI.Status(ii) = ""
        VECI.Rerun(ii) = ""
    Next ii
    flgETB = False
    SendCount = 0
    SendFlg = False
    Return
    
errReceive:

'Exception_:
'    On Error Resume Next
'    OraDataBase.RollbackTrans
'    Exit Sub

    
'-----------------단방향일경우 사용-----------------------------------------------------------
'    Dim sTemp      As String
'    Dim Channel_No As String       ' 검사항목 번호 : Channel No
'    Dim pGrid_Point As Integer
'    Dim pDoCount   As Integer
'    Dim Loop_count As Integer
'    Dim FunStr As String
'    Dim Max_Arary_Cnt As Integer    ' 검사 항목수
'    Dim sAdd As Integer, sPosition As Integer
'    Dim itemX As ListItem
'    Dim strRstval As String, strRefVal  As String
'    Dim sqlDoc  As String
'    Dim intCol As Integer
'    Dim Gnum   As String
'    Dim ii As Integer, jj As Integer, kk As Integer
'    Dim Test_Cd() As String
'    Dim Rev As Long
'    Dim tmpTstCd As String
'    Dim tmpMXD As Variant
'    Dim sSeq, strTmp, varTmp, strBarno, strDate, strTime As String
'    Dim sCol As Integer
'    Dim sDecnt As Integer
'    Dim Float_rate1 As String
'    Dim Float_rate2 As String
'    Dim Float_rate  As String
'    Dim intRow, intIdx As Integer
'    On Error Resume Next
'
'    CallForm = "frmInterface - Privete sub psDataDefine()"
'    tmpMXD = "0"
'    sTemp = strdata
'    Erase fVitors
'    '------------------------------<<< fVitors() 배열 Clear 한다.         >>>----------
'    For Loop_count = 1 To 100: fVitors(Loop_count) = "": Next Loop_count
'    '------------------------------<<< fVitors() 배열에 구분하여 넣는다.  >>>----------
'
'    sDecnt = (Len(sTemp) - 282) / 39                                   ' 총 검사항목 갯수를 찿는다.
'    fVitors(0) = Str$(sDecnt)                                          ' 총 검사항목 갯수를 넣는다. 나중에 사용한다.
'    fVitors(1) = Mid$(sTemp, 40, 3)                                    ' 항목 2 "1"   일련번호
'
'    Dim ssTemp1 As Integer
'    Dim ssTemp2 As String
'    For pDoCount = 1 To sDecnt
'        ssTemp1 = (pDoCount - 1) * 39 + 282                            ' 첫번째 Channel 및 검사결과 위치 확인
'        ssTemp2 = Mid$(sTemp, ssTemp1, 39)
'        fVitors(((pDoCount - 1) * 2) + 4 + 1) = Mid$(ssTemp2, 6, 5)    ' channel
'        fVitors(((pDoCount - 1) * 2) + 4 + 2) = Mid$(ssTemp2, 15, 5)   ' result
'    Next pDoCount
'
'    Max_Arary_Cnt = brspread.MaxCols - 4
'    pGrid_Point = 0
'    strTmp = ""
'    If Len(fVitors(1)) > 0 Then
'        intRow = 0
'        With spdResult1
'
'            sSeq = Trim(fVitors(1))
'            sCol = 2
'            pGrid_Point = SeqSearch(spdResult1, sSeq, sCol)
'
'            .GetText 2, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)
'
'            If pGrid_Point > 0 Then
'                For intCol = 5 To .MaxCols
'                    strRstval = ""
'                    .GetText intCol, 0, varTmp
'                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
'                    If Not itemX Is Nothing Then
'                        For intIdx = 1 To Max_Arary_Cnt
'                            If Len(fVitors(((intIdx - 1) * 2) + 4 + 1)) > 0 Then
'                                If Trim(fVitors(((intIdx - 1) * 2) + 4 + 1)) = itemX.tag Then
'                                    strRstval = Trim(fVitors((intIdx - 1) * 2 + 4 + 2))
'                                    If strRstval = 99.99 Then
'                                        strRstval = ""
'                                    End If
'                                    strDate = Format$(Now, "YYYYMMDD"):     strTime = Format$(Now, "MMSS")
'                                    .SetText intCol, pGrid_Point, strRstval
'                                    .Col = intCol:  .Row = pGrid_Point
'                                                    .ForeColor = IIf(Trim$(strRefVal) <> "", vbRed, vbBlack)
'
'                                    sqlDoc = "Update INTERFACE003" & _
'                                             "   set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRefVal & "'" & _
'                                             " where SPCNO   = '" & strBarno & "'" & _
'                                             "   and EQPNUM  = '" & itemX.tag & "'" & _
'                                             "   and TRANSDT = '" & strDate & "'" & _
'                                             "   and TRANSTM = '" & strTime & "'"
'                                    AdoCn_Jet.Execute sqlDoc
'
'                                    sqlDoc = "insert into INTERFACE003(" & _
'                                             "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN)" & _
'                                             "    values( '" & strBarno & "', '" & itemX.Text & "', '" & itemX.tag & "'," & _
'                                             "            '" & strDate & "', '" & strTime & "'," & _
'                                             "            '" & strRstval & "', '" & strRefVal & "'," & _
'                                             "            '" & INS_CODE & "', '')"
'                                    AdoCn_Jet.Execute sqlDoc
'
'                                   Exit For
'                                End If
'                            End If
'                        Next intIdx
'                    End If
'                    Set itemX = Nothing
'                Next
'            End If
'        End With
'    End If
'    Exit Sub
'
'ErrRoutine:
'
'    Call ErrMsgProc(CallForm)

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
        
        Call sl_online_result_ul_4&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
        If strErrMsg = "" Then
            f_funAdd_Server = True
        Else
            Call ErrMsgProc("", strErrMsg)
        End If
'    Else
'        Call ErrMsgProc("", "검체번호 [" + strBarno + "]를 저장하지 못했습니다.")
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

'    Dim intIdx          As Integer
'    Dim strData(1 To 5)         As Byte

    
    Dim Arr()   As Byte
    Dim strTmp  As String
                      
                      '         1         2         3         4         5         6         7         9         10
                      '123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789
                      
             strTmp = "!000a02 02HSHC  1212020304091              4 0101             FF01.000  66"
    strTmp = strTmp & "!001c                                                                                              D5"
    strTmp = strTmp & "!002d                                                                            97"
    strTmp = strTmp & "!003fTSH      0.664mIU/L       008D"
    strTmp = strTmp & "!004fFT4       12.5pmol/L      00E2"
    strTmp = strTmp & "!005fTT3       1.93nmol/L      00F3"
    strTmp = strTmp & "!006fTT4       84.8nmol/L      00FC"
    strTmp = strTmp & "!007fFerr      73.6ng/mL       0038"
    strTmp = strTmp & "!008fPSA      0.338ng/mL       00BC"
    strTmp = strTmp & "!009fAFP       1.68IU/mL       006A"
    strTmp = strTmp & "!010fCA199      0.0U/ml        004A"
    strTmp = strTmp & "!011fCEA       1.42ng/mL       0084"
    strTmp = strTmp & "!012faHCV      0.09            Q0E4"
    strTmp = strTmp & "!013gT3/T4    0.023            07D"
    strTmp = strTmp & "!014h  02C0"
    
    Call ComReceive(strTmp)
    
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
    mskOrdDate1.Text = Format$(Now, "YYYYMMDD")
    Open App.Path + "\" + "VitrosEci.log" For Append As #1

    Print #1, Chr(13) + Chr(10);
    
    f_strJOB_FLAG = "1":    f_intSampleNo = 0
    cboRstgbn(0).ListIndex = 0: cboRstgbn(1).ListIndex = 2
    tabWork.Tab = 0
    Or_Seq = 1
    
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

Private Sub Result_MsgBegin(ByVal SID As String)
'
'    Dim itemX As ListItem
'
'    Set itemX = lvwComplete.FindItem(Trim(SID), lvwTag, , lvwWhole)
'    If itemX Is Nothing Then
'        Set itemX = lvwComplete.ListItems.Add(, , Trim(SID))
'        If Not itemX Is Nothing Then
'            With itemX
'                .Key = COL_KEY & Trim(SID)
'                .tag = Trim(SID)
'                .SmallIcon = "LST"
'            End With
'        End If
'    End If
'
End Sub

Private Sub Result_MsgSplit(ByVal Result As clsResult)

'On Error GoTo ErrorRoutine
'
'    Dim sqlDoc  As String, sqlRet   As Integer
'
'    Dim strTime As String
'    Dim itemX   As ListItem
'    Dim itemH   As ListItem
'    Dim itemS   As ListSubItem
'
'    CallForm = "frmComm - Private Sub Result_MsgSplit()"
'
'    '메치 테이블에서 검사코드를 가져옴
'    Set itemX = lvwCuData.FindItem(Trim(Result.Rst_Test), lvwTag, , lvwWhole)
'    If Not itemX Is Nothing Then
'        If Mid$(Result.Rst_Sid, 10, 2) = "PC" And Trim(Result.Rst_Test) = "06A" Then
'            Result.Rst_Sid = Mid$(Result.Rst_Sid, 1, 9)
'            Result.Rst_Test = "XXX"
'            Result.Rst_Tag = ""
'        Else
'            Result.Rst_Sid = Mid$(Result.Rst_Sid, 1, 9)
'            Result.Rst_Tag = Trim(itemX.SubItems(1))
'        End If
'
'        sqlDoc = "Update INTERFACE003 set RSTVAL = '" & Result.Rst_Values & "', REFVAL = '" & Result.Rst_Eid & "'" & _
'                 " where SPCNO  = '" & Result.Rst_Sid & "'" & _
'                 "   and TESTCD = '" & Result.Rst_Test & "'" & _
'                 "   and TRANSDT = '" & Format$(Now, "YYYYMMDD") & "'" & _
'                 "   and TRANSTM = '" & Format$(Now, "MMSS") & "'"
'        AdoCn_Jet.Execute sqlDoc, sqlRet
'        If sqlRet = 0 Then
'            sqlDoc = "insert into INTERFACE003(" & _
'                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD)" & _
'                     "    values( '" & Result.Rst_Sid & "', '" & Result.Rst_Test & "'," & _
'                     "            '" & Result.Rst_Eid & "', '" & Format$(Now, "YYYYMMDD") & "'," & _
'                     "            '" & Format$(Now, "MMSS") & "', '" & Result.Rst_Values & "'," & _
'                     "            '" & Result.Rst_Eid & "', '" & INS_CODE & "')"
'            AdoCn_Jet.Execute sqlDoc
'        End If
'
'        '결과 표시
'        Set itemH = lvwComplete.FindItem(Result.Rst_Sid, lvwText, , lvwWhole)
'        If itemH Is Nothing Then
'            Set itemH = lvwComplete.ListItems.Add()
'            With itemH
'                .Key = COL_KEY & Result.Rst_Sid '아이템 키에 검체번호
'                .Text = Result.Rst_Sid          '아이템 에 검체번호
'                .tag = Result.Rst_Type          '테그에 결과 타입
'                .SmallIcon = "LSE"
'            End With
'        End If
'        '결과값 등록
'        itemH.SubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex) = Result.Rst_Values
'
'        '--- 판정
'        itemH.ListSubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex).ForeColor = vbBlack
'        If Val(itemX.SubItems(7)) < Val(Result.Rst_Values) Or Val(itemX.SubItems(8)) > Val(Result.Rst_Values) Then
'            itemH.ListSubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex).ForeColor = vbRed
'        End If
'
'        Set itemS = itemH.ListSubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex)
'
'        itemS.tag = Result.Rst_Error '서브아이템 테그에 에러 메시지
'        Set itemS = Nothing
'        Set itemX = Nothing
'        Set itemX = Nothing
'    End If
'    '검사코드가 없는것은 등록 하지 않음
'    Exit Sub
'ErrorRoutine:
'
'    Set itemS = Nothing
'    Set itemX = Nothing
'    Set itemX = Nothing
'
'    Call ErrMsgProc(CallForm)
'    Err.Clear
'
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
            If aCOL = 5 Then
                iCnt = 0
                .GetText 1, aROW, varChk
                .GetText 2, aROW, varBar
                .GetText aCOL, aROW, varNum
                If Trim(varChk) = "1" And Trim(varBar) <> "" Then
                    For iRow = aROW To .maxrows
                        .SetText aCOL, iRow, varNum
                        .SetText aCOL + 1, iRow, ((iCnt Mod 10) + 1) - 1
                        iCnt = iCnt + 1
                        If (iCnt Mod 10) = 0 Then varNum = varNum + 1
                    Next
                End If
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

Private Sub txtBarCode_Change()

    If txtBarCode.SelStart = txtBarCode.MaxLength Then SendKeys "{TAB}"
    
End Sub

Private Sub txtBarCode_GotFocus()

    With txtBarCode
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


'Private Sub txtBarCode_KeyPress(KeyAscii As Integer)
'
'    Dim tst_no() As String, strPid()    As String, strPnm() As String
'    Dim TMP() As String
'    Dim rv As Long
'    Dim samChk As Boolean
'    Dim ii As Integer
'    Dim bgetWork As Boolean
'    Dim itemX As ListItem
'
'    samChk = False
'    If KeyAscii = vbKeyReturn Then
'        rv = sl_spcid_tstcd_select(Trim(txtBarCode.Text), tst_no, strPid, strPnm)
'        If (rv = 0) Then
'            MsgBox "미접수 검체입니다.!", vbCritical
'        Else
'            If psDataExists Then
'                MsgBox "이미 등록된 검체입니다.!", vbCritical
'                txtBarCode.Text = ""
'                Exit Sub
'            End If
'
'            bgetWork = False
'            For ii = 0 To rv - 1
'                Set itemX = lvwCuData.FindItem(tst_no(ii), lvwText, , lvwWhole)
'                If Not itemX Is Nothing Then
'                    bgetWork = True
'                End If
'            Next
'
'             With spdWorkList
'                If bgetWork = True Then
'                    .Col = 2
'                    For ii = 1 To .maxrows
'                        .Row = ii
'                        If Trim(.Text) = "" Then
'                            .Text = txtBarCode.Text
'                            .SetText 3, ii, strPnm(0)
'                            .SetText 4, ii, strPid(0)
'                            txtBarCode.Text = ""
'                            .Col = 1
'                            .Value = 1
'                            samChk = True
'                            Exit For
'                        End If
'                    Next
'                    If samChk = False Then
'                         .maxrows = .maxrows + 1
'                         .Row = .maxrows
'                         .Text = txtBarCode.Text
'                         .SetText 3, .maxrows, strPnm(0)
'                         .SetText 3, .maxrows, strPid(0)
'                         .RowHeight(.maxrows) = 13
'                         txtBarCode.Text = ""
'                    End If
'                Else
'                   MsgBox "해당검사항목이 존재하지 않는 검체입니다.", vbOKOnly + vbInformation, Me.Caption
'                End If
'             End With
'        End If
'    End If
'End Sub

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



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6660
      ScaleHeight     =   345
      ScaleWidth      =   5205
      TabIndex        =   55
      Top             =   570
      Width           =   5235
      Begin VB.OptionButton optJobgbn 
         BackColor       =   &H00FFFFC0&
         Caption         =   "WorkList"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1410
         TabIndex        =   59
         Top             =   90
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.OptionButton optJobgbn 
         BackColor       =   &H00FFFFC0&
         Caption         =   "BarCode"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   58
         Top             =   90
         Width           =   1230
      End
      Begin VB.CheckBox chkQC 
         BackColor       =   &H00FFFFC0&
         Caption         =   "QC"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3090
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   60
         Width           =   720
      End
      Begin VB.ComboBox cboLevel 
         Height          =   300
         Left            =   4020
         TabIndex        =   56
         Text            =   "Combo1"
         Top             =   30
         Width           =   1065
      End
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
            Picture         =   "frmComm_2.frx":0000
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":059A
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":0B34
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":10CE
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":1668
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":1C02
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm comEQP 
      Left            =   3780
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
      EOFEnable       =   -1  'True
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
            Picture         =   "frmComm_2.frx":219C
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":2736
            Key             =   "NOT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":2CD0
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":326A
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":3AFC
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":3C56
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":3DB0
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
      TabIndex        =   3
      Top             =   6495
      Width           =   11940
      Begin HSCotrol.CButton cmdOrder 
         Height          =   360
         Left            =   5160
         TabIndex        =   53
         Top             =   100
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   635
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
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   0
         Left            =   6375
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   7
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
      Begin MSCommLib.MSComm comOrder 
         Left            =   2340
         Top             =   -30
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         RTSEnable       =   -1  'True
         SThreshold      =   1
         EOFEnable       =   -1  'True
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   225
         Width           =   615
      End
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   555
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmComm_2.frx":3F0A
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
         TabIndex        =   10
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send : "
         Height          =   180
         Left            =   9105
         TabIndex        =   9
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port : "
         Height          =   180
         Left            =   8040
         TabIndex        =   8
         Top             =   285
         Width           =   510
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   11010
         Picture         =   "frmComm_2.frx":518C
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   9780
         Picture         =   "frmComm_2.frx":5716
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   8640
         Picture         =   "frmComm_2.frx":5CA0
         Top             =   255
         Width           =   240
      End
   End
   Begin TabDlg.SSTab tabWork 
      Height          =   5850
      Left            =   30
      TabIndex        =   13
      Top             =   660
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   10319
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " WorkList"
      TabPicture(0)   =   "frmComm_2.frx":622A
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
      Tab(0).Control(12)=   "Command1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "optBar"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "optSeq"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdStartNo"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkAuto"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "spdResult1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   " 받은 결과"
      TabPicture(1)   =   "frmComm_2.frx":6246
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
      Begin FPSpread.vaSpread spdResult1 
         Height          =   4830
         Left            =   2130
         TabIndex        =   54
         Top             =   900
         Width           =   9555
         _Version        =   196608
         _ExtentX        =   16854
         _ExtentY        =   8520
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   3
         EditEnterAction =   2
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   24
         MaxRows         =   14
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmComm_2.frx":6262
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
         Left            =   6870
         TabIndex        =   52
         Top             =   555
         Value           =   1  '확인
         Width           =   1410
      End
      Begin HSCotrol.CButton cmdStartNo 
         Height          =   300
         Left            =   8340
         TabIndex        =   51
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
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   50
         Top             =   -60
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton optBar 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   49
         Top             =   -60
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   4920
         Left            =   -67980
         TabIndex        =   46
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
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   180
         Left            =   2745
         TabIndex        =   45
         Top             =   315
         Visible         =   0   'False
         Width           =   420
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   1
         Left            =   360
         TabIndex        =   26
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_2.frx":6A20
      End
      Begin VB.ComboBox cboRstgbn 
         Height          =   300
         Index           =   1
         ItemData        =   "frmComm_2.frx":6EA2
         Left            =   -72570
         List            =   "frmComm_2.frx":6EAF
         Style           =   2  '드롭다운 목록
         TabIndex        =   16
         Top             =   495
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.TextBox txtBarCode 
         Height          =   300
         Left            =   2415
         MaxLength       =   11
         TabIndex        =   0
         Top             =   495
         Width           =   1500
      End
      Begin VB.ComboBox cboRstgbn 
         Height          =   300
         Index           =   0
         ItemData        =   "frmComm_2.frx":6ED9
         Left            =   3930
         List            =   "frmComm_2.frx":6EE6
         Style           =   2  '드롭다운 목록
         TabIndex        =   14
         Top             =   495
         Visible         =   0   'False
         Width           =   1500
      End
      Begin MSMask.MaskEdBox mskRstDate 
         Height          =   300
         Left            =   -73695
         TabIndex        =   17
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
         TabIndex        =   18
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
         Left            =   -65460
         TabIndex        =   19
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
         TabIndex        =   20
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
         TabIndex        =   21
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
         Left            =   9570
         TabIndex        =   22
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
         Left            =   90
         TabIndex        =   25
         Top             =   5490
         Width           =   1980
         _ExtentX        =   3493
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
         TabIndex        =   27
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_2.frx":6F10
      End
      Begin FPSpread.vaSpread spdWorkList 
         Height          =   4560
         Left            =   90
         TabIndex        =   1
         Top             =   900
         Width           =   2025
         _Version        =   196608
         _ExtentX        =   3572
         _ExtentY        =   8043
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   6
         MaxRows         =   1
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmComm_2.frx":737E
      End
      Begin HSCotrol.UserPanel pnlCom 
         Height          =   5355
         Left            =   45
         TabIndex        =   28
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
            TabIndex        =   34
            Top             =   270
            Width           =   11595
         End
         Begin VB.Frame Frame1 
            Height          =   645
            Left            =   45
            TabIndex        =   29
            Top             =   4650
            Width           =   11610
            Begin HSCotrol.CButton cmdCOMSave 
               Height          =   360
               Left            =   10515
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
               Left            =   9450
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
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   3
         Left            =   -74640
         TabIndex        =   47
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_2.frx":76F8
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   -74910
         TabIndex        =   48
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_2.frx":7B7A
      End
      Begin FPSpread.vaSpread spdResult2 
         Height          =   4830
         Left            =   -74910
         TabIndex        =   15
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
         MaxCols         =   20
         MaxRows         =   1
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmComm_2.frx":7FE8
      End
      Begin HSCotrol.UserPanel pnlCom2 
         Height          =   4785
         Left            =   5895
         TabIndex        =   35
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
            TabIndex        =   44
            Top             =   270
            Width           =   5730
         End
         Begin VB.Frame Frame2 
            Height          =   645
            Left            =   90
            TabIndex        =   36
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
               TabIndex        =   37
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
               TabIndex        =   38
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
            Begin HSCotrol.CButton cmdCOMInput2 
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
            Begin HSCotrol.CButton cmdCOMLoad 
               Height          =   360
               Left            =   4635
               TabIndex        =   41
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
               TabIndex        =   42
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
               TabIndex        =   43
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
         TabIndex        =   24
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
         TabIndex        =   23
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

Public WithEvents Result As clsMsg_Result
Attribute Result.VB_VarHelpID = -1
Public WithEvents Order  As clsMsg_Query
Attribute Order.VB_VarHelpID = -1
Public Result1 As clsResult
Attribute Result1.VB_VarHelpID = -1

Private mAdoRs      As ADODB.Recordset
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

Dim fXE2100(100) As String
Dim fXE2100Cfg(100) As Integer
Dim fXe2100Size(100, 1) As Integer
Dim fRcvString As String
Dim fChannel() As String

Dim fXE2100Order(100) As String
Dim chkCbcAll() As Boolean
Dim chkRetAll() As Boolean

Private Type TYPE_CD
    strEqpCd    As String
    intCnt      As Integer
    strTestcd(100) As String
End Type
Private f_typCode() As TYPE_CD


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

Private Sub f_subSet_ItemList()

    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim intCol  As Integer, intPos  As Integer, intCnt  As Integer
    Dim strTmp  As String
    
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = ""
    
    intCol = 3
    With spdWorkList
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 14
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
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " order by TESTCD "
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
            If intCol + 2 > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol + 2, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        With spdResult1
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol + 4, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        With spdResult2
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        fChannel(intCol - 2) = adoRS.Fields("TEST_EQP")
        
        intCol = intCol + 1
        strTmp = Trim$(adoRS.Fields("TESTCD"))
        intPos = InStr(strTmp, ",")
        Do While intPos > 0
            f_strOrdList = f_strOrdList + "'" + Mid$(strTmp, 1, intPos - 1) + "',"
            strTmp = Mid$(strTmp, intPos + 1)
            intPos = InStr(strTmp, ",")
        Loop
        f_strOrdList = f_strOrdList + "'" + strTmp + "',"

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
        
        adoRS.MoveNext
    Loop
    Set adoRS = Nothing
    
    f_strOrdList = Mid$(f_strOrdList, 1, Len(f_strOrdList) - 1)
    
Exit Sub
ErrRoutine:
    Set adoRS = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub f_subSet_Result(ByVal strdata As String)

    Dim sqlDoc  As String
    Dim varTmp  As Variant

    Dim strSampleno As String
    Dim strTestcd(1 To 19) As String, strEquipCd(1 To 19) As String
    Dim strRstval(1 To 19) As String, strRefVal(1 To 19)  As String
    Dim intIdx      As Integer, intCol  As Integer
    
    Dim itemX   As ListItem
    
    On Error GoTo ErrRoutine
    
    CallForm = "frmInterface - Privete sub f_subSet_Result()"
    
    If Mid$(strdata, 1, 1) <> "D" Or Mid$(strdata, 3, 1) <> "U" Then Exit Sub
    
    f_intSampleNo = f_intSampleNo + 1
    
    spdResult1.SetText 1, f_intSampleNo, "1"
    spdResult1.GetText 2, f_intSampleNo, varTmp
    If Trim$(varTmp) = "" Then
        strSampleno = Format$(Now, "YYYYMMDD") + "X" + Format$(f_intSampleNo, "000")
        spdResult1.GetText 2, f_intSampleNo, varTmp
    Else
        strSampleno = Trim$(varTmp)
    End If
    
    strRstval(1) = Mid$(strdata, 54, 6):    strRstval(1) = Format$(Mid$(strRstval(1), 1, 3), "###") + "." + Mid$(strRstval(1), 4, 3)   '-- 1. WBC
    strRstval(2) = Mid$(strdata, 60, 5):    strRstval(2) = Format$(Mid$(strRstval(2), 1, 2), "##") + "." + Mid$(strRstval(2), 3, 3)    '-- 2. RBC
    strRstval(3) = Mid$(strdata, 65, 5):    strRstval(3) = Format$(Mid$(strRstval(3), 1, 3), "###") + "." + Mid$(strRstval(3), 4, 2)     '-- 3. HGB
    strRstval(4) = Mid$(strdata, 70, 5):    strRstval(4) = Format$(Mid$(strRstval(4), 1, 3), "###") + "." + Mid$(strRstval(4), 4, 2)              '-- 4. HCT
    strRstval(5) = Mid$(strdata, 75, 5):    strRstval(5) = Format$(Mid$(strRstval(5), 1, 3), "###") + "." + Mid$(strRstval(5), 4, 2)              '-- 5. MCV
    
    strRstval(6) = Mid$(strdata, 80, 5):    strRstval(6) = Format$(Mid$(strRstval(6), 1, 3), "###") + "." + Mid$(strRstval(6), 4, 2)             '-- 6. MCH
    strRstval(7) = Mid$(strdata, 85, 5):    strRstval(7) = Format$(Mid$(strRstval(7), 1, 3), "###") + "." + Mid$(strRstval(7), 4, 2)             '-- 7. MCHC
    strRstval(8) = Mid$(strdata, 90, 5):    strRstval(8) = Format$(Mid$(strRstval(8), 1, 3), "###") + "." + Mid$(strRstval(8), 4, 2)             '-- 8. PLT
    strRstval(9) = Mid$(strdata, 95, 5):    strRstval(9) = Format$(Mid$(strRstval(9), 1, 3), "###") + "." + Mid$(strRstval(9), 4, 2)             '-- 9. W-SCR
    strRstval(10) = Mid$(strdata, 100, 5):   strRstval(10) = Format$(Mid$(strRstval(10), 1, 3), "###") + "." + Mid$(strRstval(10), 4, 2)           '-- 10. W-MCR
    strRstval(11) = Mid$(strdata, 105, 5):   strRstval(11) = Format$(Mid$(strRstval(11), 1, 3), "###") + "." + Mid$(strRstval(11), 4, 2)            '-- 11. W-LCR
    
    strRstval(12) = Mid$(strdata, 115, 6):   strRstval(12) = Format$(Mid$(strRstval(12), 1, 3), "###") + "." + Mid$(strRstval(12), 4, 3)             '-- 12. W-SCC
    strRstval(13) = Mid$(strdata, 121, 6):   strRstval(13) = Format$(Mid$(strRstval(13), 1, 3), "###") + "." + Mid$(strRstval(13), 4, 3)             '-- 13. W-MCC
    strRstval(14) = Mid$(strdata, 127, 7):   strRstval(14) = Format$(Mid$(strRstval(14), 1, 3), "###") + "." + Mid$(strRstval(14), 4, 3)             '-- 14. W-LCC
    strRstval(15) = Mid$(strdata, 139, 5):   strRstval(15) = Format$(Mid$(strRstval(15), 1, 3), "###") + "." + Mid$(strRstval(15), 4, 2)             '-- 15. RDW-CV
    strRstval(16) = Mid$(strdata, 144, 5):   strRstval(16) = Format$(Mid$(strRstval(16), 1, 3), "###") + "." + Mid$(strRstval(16), 4, 2)             '-- 16. RDW-SD
    strRstval(17) = Mid$(strdata, 149, 5):   strRstval(17) = Format$(Mid$(strRstval(17), 1, 3), "###") + "." + Mid$(strRstval(17), 4, 2)             '-- 17. PDW
    strRstval(18) = Mid$(strdata, 154, 5):   strRstval(18) = Format$(Mid$(strRstval(18), 1, 3), "###") + "." + Mid$(strRstval(18), 4, 2)             '-- 18. MPV
    strRstval(19) = Mid$(strdata, 159, 5):   strRstval(19) = Format$(Mid$(strRstval(19), 1, 3), "###") + "." + Mid$(strRstval(19), 4, 2)             '-- 19. P-LCR
    
    For intIdx = 1 To UBound(strRstval)
        If strRstval(intIdx) <> "" Then
            
            Set itemX = lvwCuData.FindItem(Format$(intIdx, "00"), lvwTag, , lvwWhole)
            If Not itemX Is Nothing Then
                
                If itemX.ListSubItems(8) <> "" And itemX.ListSubItems(9) <> "" Then
                    If Val(strRefVal(intIdx)) < itemX.ListSubItems(8) Then
                        strRefVal(intIdx) = "L"
                    ElseIf Val(strRefVal(intIdx)) > itemX.ListSubItems(9) Then
                        strRefVal(intIdx) = "H"
                    End If
                End If
                
                sqlDoc = "insert into INTERFACE003(" & _
                         "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN)" & _
                         "    values( '" & strSampleno & "', '" & itemX.ListSubItems(1) & "', ''," & _
                         "            '" & Format$(Now, "YYYYMMDD") & "', '" & Format$(Now, "MMSS") & "'," & _
                         "            '" & strRstval(intIdx) & "', '" & strRefVal(intIdx) & "'," & _
                         "            '" & itemX.tag & "', '')"
                AdoCn_Jet.Execute sqlDoc
                
                intCol = itemX.Index
                spdResult1.SetText intCol + 4, f_intSampleNo, strRstval(intIdx)
                spdResult1.Col = intCol + 4
                spdResult1.Row = f_intSampleNo
                spdResult1.ForeColor = IIf(strRefVal(intIdx) <> "", vbRed, vbBlack)
            End If
            Set itemX = Nothing
            
        End If
    Next
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub cmdACK_Click()
'
'    Call COM_OUTPUT(charCOM_Convert(COM_ACK))
Call COM_OUTPUT(Chr(1))
End Sub

Private Sub cmdAction_Click(Index As Integer)
    
    Select Case Index
        Case 0
            Call cmdRun
        Case 1
            Call cmdStop
        Case 2
            Call cmdClear
        Case 3 'cmd close
            Call cmdExit
        Case Else
    End Select

End Sub

Private Sub cmdClear()
    
    f_strJOB_FLAG = "1"
    f_intSampleNo = 0
    
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
    
    Dim strOrdLst()     As String, strPid()     As String, strPnm()     As String
    
    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
    Dim itemX   As ListItem
    Dim objSpd  As vaSpread
    
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
            .GetText .MaxCols, intRow, varTmp:  strTime = Trim$(varTmp)
            .GetText 1, intRow, varTmp
            
            If strBarno = "" Then Exit For
            If chkQC.Value = 0 Then
                Call sl_spcid_tstcd_select&(strBarno, strOrdLst, strPid, strPnm)
            Else
                Call sl_spcid_tstcd_select_qc&(INS_CODE, strBarno, strOrdLst)
            End If
            
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
                            Do While intPos > 0
                                
                                blnFlag = False
                                For intIdx = 0 To UBound(strOrdLst)
                                    If strOrdLst(intIdx) = Mid$(strTestcd, 1, intPos - 1) Then blnFlag = True:  Exit For
                                Next
                                
                                If blnFlag Then
                                    intCnt = intCnt + 1
                                    ReDim Preserve strSampleno(1 To intCnt) As String
                                    ReDim Preserve strOrdcd(1 To intCnt) As String
                                    ReDim Preserve strRstval(1 To intCnt) As String
                                    ReDim Preserve strTmp1(1 To intCnt) As String
                                    ReDim Preserve strTmp2(1 To intCnt) As String
                                    
                                    strSampleno(intCnt) = strBarno
                                    strOrdcd(intCnt) = Mid$(strTestcd, 1, intPos - 1)
                                    strRstval(intCnt) = Trim$(varTmp)
                                End If
                                
                                strTestcd = Mid$(strTestcd, intPos + 1)
                                intPos = InStr(strTestcd, ",")
                            Loop
                            
                            blnFlag = False
                            For intIdx = 0 To UBound(strOrdLst)
                                If strOrdLst(intIdx) = strTestcd Then blnFlag = True: Exit For
                            Next
                            
                            If blnFlag Then
                                intCnt = intCnt + 1
                                ReDim Preserve strSampleno(1 To intCnt) As String
                                ReDim Preserve strOrdcd(1 To intCnt) As String
                                ReDim Preserve strRstval(1 To intCnt) As String
                                ReDim Preserve strTmp1(1 To intCnt) As String
                                ReDim Preserve strTmp2(1 To intCnt) As String
                                
                                strSampleno(intCnt) = strBarno
                                strOrdcd(intCnt) = strTestcd
                                strRstval(intCnt) = Trim$(varTmp)
                            End If
                        End If
                        Set itemX = Nothing
                        .SetText 1, intRow, "0"
                    End If
                Next
                
                If intCnt > 0 Then
                    If chkQC.Value = 0 Then
                        Call sl_online_result_ul_4&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
                    Else
                        Call sl_online_pc_98&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
                    End If
                    If strErrMsg = "" Then
                        spdResult2.Row = intRow
                        spdResult2.Col = -1:    spdResult2.BackColor = &HFFF8F0
                    
                        sqlDoc = "Update INTERFACE003 set SERVERGBN  = 'Y'" & _
                                 " where SPCNO   = '" & strBarno & "'" & _
                                 "   and TRANSDT = '" & mskRstDate.Text & "'" & _
                                 "   and TRANSTM = '" & strTime & "'"
                        AdoCn_Jet.Execute sqlDoc
                    Else
                        MsgBox strErrMsg, vbInformation, Me.Caption
                    End If
                Else
                    lblStatus.Caption = "검체번호 [" + strBarno + "]를 저장하지 못했습니다." ', vbInformation, Me.Caption
                End If
            End If
        Next
    End With
    Me.MousePointer = 0
    lblStatus.Caption = "작업이 완료되었습니다." ', vbInformation, Me.Caption
    
    Exit Sub
ErrorRoutine:
    Set itemX = Nothing
    
    Me.MousePointer = 0
    Call ErrMsgProc(CallForm)
   
  
End Sub

Private Sub cmdENQ_Click()
    
    Call COM_OUTPUT(charCOM_Convert(COM_ENQ))

End Sub

Private Sub cmdOrder_Click()
    Dim varTmp  As Variant
    Dim intRow1 As Integer, intRow2 As Integer
    Dim intIdx  As Integer
    Dim fNum1   As Integer, fNum2  As Integer
    Dim Rev     As Long
    Dim Test_Cd() As String
    Dim ii As Integer, kk As Long
    Dim itemX As ListItem
    Dim adoRS As ADODB.Recordset
    Dim sqlDoc As String
    Dim bgetWork As Boolean
    Dim strEqpCd    As String
    Dim RetVal As Long
    Dim sBuf As String
    Dim com_port As String
    Dim Deletime As String
    Dim set_ting As String
    Dim strOrder As String
    Dim tmpChannel As String, tmpChannel1 As String
    Dim strErrMsg As String
    Dim tmpBarnum() As String
    Dim reChk As Boolean
    Dim strWorkNum As String
    Dim chkCBC As Boolean
    Dim chkDIFF As Boolean
    Dim chkRET As Boolean
    
    
    reChk = False
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("OrderPort", "CommPort", "No", sBuf, 255, App.Path & "\orderport.ini")
    com_port = left(sBuf, RetVal)
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("OrderPort", "Settings", "No", sBuf, 255, App.Path & "\orderport.ini")
    set_ting = left(sBuf, RetVal)
    'Exit Sub
    With comOrder
        If .PortOpen = False Then
            .CommPort = com_port
            '속도,페리티,테이타비트,stop bit
            .Settings = set_ting
            .RThreshold = 1
            .SThreshold = 1
            '-- 임시 : 노트북이 com1만 있으므로...
            If comEQP.PortOpen = True Then comEQP.PortOpen = False
            '--------------------------------------------
            If .PortOpen = True Then .PortOpen = False
            .PortOpen = True
        End If
    End With
    
    'If reFlag = False Then
    bgetWork = False
    fNum1 = 1: fNum2 = 0
    With spdResult1 'spdWorkList
        For intRow1 = 1 To .maxrows
            chkCBC = False: chkRET = False: chkDIFF = False
            .GetText 1, intRow1, varTmp
            If Trim$(varTmp) = "1" Then
                .GetText 2, intRow1, varTmp
                If Len(varTmp) > 0 Then
                    reChk = True

                    ReDim Preserve tmpBarnum(intRow1 - 1)
                    tmpBarnum(intRow1 - 1) = varTmp
                    
                    lblStatus.Caption = "오더전송중.."
                    'S0000000100 0003080407403 000001 01 101 111111111111111111111110011111111000000000000000011111000000000000000000000000000000000000000000****************************************
                    tmpChannel = "00000000000000000000000000000000000" '-- 35
                                                
                    strOrder = "S00000000"
                    .GetText 2, intRow1, varTmp
                    
                    If Len(varTmp) > 0 Then
                        '-- Order Data 가공
                        strOrder = strOrder + "00" + Format(varTmp, "0000000000000") + "00000000" + "000"
                        
                        Dim iCnt1 As Integer
                        Dim oNum As Integer
                        
                        For iCnt1 = 5 To .MaxCols
                            .Col = iCnt1: .Row = intRow1
                            If .BackColor = &HC6FEFF Then
                                .Row = 0
                                Set adoRS = New ADODB.Recordset
                                sqlDoc = "select TESTCD_EQP, PANICH from INTERFACE002" & _
                                         " where (EQP_CD = '" & INS_CODE & "') AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
                                         "   and TESTNM_EQP = '" & Trim(.Text) & "' "
                                adoRS.CursorLocation = adUseClient
                                adoRS.Open sqlDoc, AdoCn_Jet
                                If adoRS.RecordCount > 0 Then adoRS.MoveFirst
                                oNum = CInt(adoRS.Fields(0))
                                
                                If adoRS.Fields(1) = "CBC" Then chkCBC = True
                                If adoRS.Fields(1) = "DIFF" Then chkDIFF = True
                                If adoRS.Fields(1) = "RET" Then chkRET = True
                                                                        
                                If chkCBC = True And chkDIFF = True And chkRET = True Then    '-- CBC + DIFF + RET
                                    tmpChannel = "11111111111111111111111001111111100"
                                ElseIf chkCBC = True And chkDIFF = True And chkRET = False Then '-- CBC + DIFF
                                    tmpChannel = "11111111111111111111111000000000100"
                                ElseIf chkCBC = True And chkDIFF = False And chkRET = True Then '-- CBC + RET
                                    tmpChannel = "11111111000001111111111001111111100"
                                ElseIf chkCBC = True And chkDIFF = False And chkRET = False Then '-- CBC
                                    tmpChannel = "11111111000001111111111000000000100"
                                ElseIf chkCBC = False And chkDIFF = True And chkRET = False Then '-- DIFF
                                    tmpChannel = "00000000111110000000000000000000000"
                                ElseIf chkCBC = False And chkDIFF = False And chkRET = True Then '-- RET
                                    tmpChannel = "00000000000000000000000001111111000"
                                ElseIf chkCBC = False And chkDIFF = True And chkRET = True Then '-- DIFF + RET
                                    tmpChannel = "00000000111110000000000001111111000"
                                'ElseIf chkCBC = False And chkDIFF = False And chkRET = True Then
                                '    tmpChannel = "00000000000000000000000001111110000"
                                Else
                                    If adoRS.Fields(0) = 25 Then
                                        oNum = 26
                                    Else
                                        If adoRS.Fields(0) >= 24 Then oNum = CInt(adoRS.Fields(0)) + 2
                                    End If
                                    tmpChannel = Mid(tmpChannel, 1, oNum - 1) & "1" & Mid(tmpChannel, oNum + 1)
                                End If
                                adoRS.Close: Set adoRS = Nothing
                            End If
                        Next iCnt1
                        
                        strOrder = strOrder + tmpChannel + "00000000000000" + "00000" + "0000000000000" + "0000000000000" + "0000000000000" + "000" + "****************************************"
                        
                        comOrder.Output = strOrder
                        'Debug.Print "[HOST ]" + strOrder
                        strOrder = ""
                    End If
                 End If
            End If
            .GetText 2, intRow1, varTmp
            If Len(Trim(varTmp)) > 0 Then
                .Col = 2: .Row = intRow1
                .BackColor = vbCyan
            End If
            
            'tmrFlag = True
            sBuf = String(255, 0)
            RetVal = GetPrivateProfileString("OrderPort", "Deletime", "No", sBuf, 255, App.Path & "\orderport.ini")
            Deletime = left(sBuf, RetVal)
            'If fXE2100Order(intRow1) <> "" Then
            For kk = 0 To CLng(Deletime): Next kk
            'End If
        Next
'        Timer1.Enabled = True
'        Timer1.Interval = 20000
    End With
    'End If
'    If reChk = True Then
'        Call sl_p_61_flagup(strErrMsg, tmpBarnum)
'        If strErrMsg <> "0" Then MsgBox strErrMsg, vbInformation, Me.Caption
'        reChk = False
'    End If
    'Timer1.Enabled = True
'    reFlag = True
    f_strJOB_FLAG = "1"
    f_intSampleNo = 0

End Sub

Private Sub cmdRstQuery_Click()

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strSpcno    As String
    Dim intRow      As Integer, intCol  As Integer
    
    Dim itemX       As ListItem

    intRow = 0
    With spdResult2
        .maxrows = 14
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
                
                '.SetText 1, intRow, "1"
                .SetText 2, intRow, Trim$(adoRS(0) & "")
            End If
            strSpcno = Trim$(adoRS(0) & "") + Trim$(adoRS(6) & "")
            Set itemX = lvwCuData.FindItem(Trim$(adoRS(7) & ""), lvwTag, , lvwWhole)
            If Not itemX Is Nothing Then
                intCol = itemX.Index + 2
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
'            For sCnt = .ActiveRow To .maxrows
'                .Row = sCnt
'                .Col = 2:       .Text = Trim(sAdd + Val(sNo))
'                sAdd = sAdd + 1
'            Next sCnt
            .StartingRowNumber = Val(sNo)
        End With
    End If

End Sub

Private Sub cmdWordQuery_Click()

    If optJobgbn(1).Value = False Then Exit Sub

    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Privete sub cmdWorkQuery_Click()"
    
    Dim strKeyno    As String
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String, strBarno()   As String
    Dim strTestcd() As String, strTPid()    As String, strTPnm() As String
    Dim strEqpCd    As String
    Dim intRow  As String, intIdx  As Integer, intCol   As Integer
    Dim itemX   As ListItem
    Dim ii As Integer
    Dim strLevel()  As String
    
'    intIdx = sl_tstcd_spcid_select&(mskOrdDate.Text, "(" + f_strOrdList + ")", strBarno, strPid, strPnm, strOrdcd)
    If chkQC.Value = 1 Then
        If Trim(cboLevel.Text) <> "" Then
            intIdx = sl_spcid_tstcd_select_qc1&(INS_CODE, cboLevel.Text, strBarno, strOrdcd, strLevel)
        Else
            lblStatus.Caption = "QC Level을 선택하세요!"
            cboLevel.SetFocus
            Exit Sub
        End If
    Else
        intIdx = sl_tstcd_spcid_select&(mskOrdDate.Text, "(" + f_strOrdList + ")", strBarno, strPid, strPnm, strOrdcd)
    End If
    
    If intIdx = 0 Then Exit Sub
    intRow = 0
    For intIdx = 0 To UBound(strOrdcd) - 1
        With spdWorkList
            If strKeyno <> strBarno(intIdx) Then
                intRow = intRow + 1
                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                
                If chkQC.Value = 1 Then
                    .SetText 2, intRow, strBarno(intIdx)
                    .SetText 3, intRow, strOrdcd(intIdx)
                    .SetText 4, intRow, strLevel(intIdx)
                Else
                    .SetText 2, intRow, strBarno(intIdx)
                    .SetText 3, intRow, strPnm(intIdx)
                    .SetText 4, intRow, strPid(intIdx)
                End If
                
'                intCol = sl_spcid_tstcd_select(strBarno(intIdx), strTestcd, strTPid, strTPnm)
'                '-- OSW EDIT
'                'strEqpCd = f_funGet_CODE(strTestcd(intIdx))
'                For ii = 0 To intCol
'                    strEqpCd = f_funGet_CODE(strTestcd(ii))
'                    Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
'                    If Not itemX Is Nothing Then .SetText 6 + itemX.Index, intRow, "V"
'                    Set itemX = Nothing
'                Next
            End If
            If chkQC.Value = 1 Then
                intCol = sl_spcid_tstcd_select_qc&(INS_CODE, strBarno(intIdx), strTestcd)
            Else
                intCol = sl_spcid_tstcd_select(strBarno(intIdx), strTestcd, strTPid, strTPnm)
            End If

            strEqpCd = f_funGet_CODE(strOrdcd(intIdx))
            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
            If Not itemX Is Nothing Then .SetText 6 + itemX.Index, intRow, "V"
            Set itemX = Nothing
            strKeyno = strBarno(intIdx)
        End With
    Next
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)

End Sub

Public Function Text_Redefine(FSend_Str As String, FCheck_Char As String) As String
    If InStr(FSend_Str, FCheck_Char) > 0 Then
        Text_Redefine = left$(FSend_Str, InStr(FSend_Str, FCheck_Char) - 1)
    Else
        Text_Redefine = FSend_Str
    End If
    
End Function

Private Sub cmdWorkList_Click()
    Dim varTmp  As Variant
    Dim intRow1 As Integer, intRow2 As Integer
    Dim intIdx  As Integer
    Dim intCol  As Integer
    ReDim strDta(1 To spdWorkList.MaxCols) As String
    
    Dim itemX   As ListItem
    
    With spdWorkList
        For intRow1 = 1 To .maxrows
            For intCol = 1 To .MaxCols
                .GetText intCol, intRow1, varTmp:   strDta(intCol) = Trim$(varTmp)
            Next
            
            If strDta(2) = "" Then Exit For
            If strDta(1) = "1" Then
                intRow2 = f_funGet_SpreadRow(spdResult1, 2, strDta(2))
                If intRow2 < 1 Then
                    intRow2 = f_funGet_SpreadRow(spdResult1, 2, "")
                    If intRow2 < 1 Then
                        spdResult1.maxrows = spdResult1.maxrows + 1
                        spdResult1.RowHeight(spdResult1.maxrows) = 13
                        intRow2 = spdResult1.maxrows
                    End If
                    spdResult1.SetText 2, intRow2, strDta(2)
                    spdResult1.SetText 3, intRow2, strDta(3)
                    spdResult1.SetText 4, intRow2, strDta(4)
                End If
                .SetText 1, intRow1, ""
                
                spdResult1.SetText 1, intRow2, "1"
                
                For intCol = 7 To UBound(strDta) - 1
                    .GetText intCol, 0, varTmp
                    If Trim$(varTmp) = "" Then Exit For
                    
                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                    If Not itemX Is Nothing Then
                        spdResult1.Col = intCol:  spdResult1.Row = intRow2
                        spdResult1.BackColor = IIf(strDta(intCol) = "V", &HC6FEFF, vbWhite)
                        'spdResult1.BackColor = IIf(strDta(intCol) = "", &HC6FEFF, vbWhite)
                    End If
                    Set itemX = Nothing
                Next
                
                .Row = intRow1
                .Action = ActionDeleteRow
                .RowHeight(.maxrows) = 13
                If .maxrows > 14 Then .maxrows = .maxrows - 1
                If intRow1 > 0 Then intRow1 = intRow1 - 1
            End If
        Next
    End With
    
    f_strJOB_FLAG = "1"
'    f_intSampleNo = 0
End Sub

Private Sub comEQP_OnComm()
    
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim Arr()       As Byte
    Dim strdata     As String
    Dim tmpData     As String
    
''              strdata = "D1U   XE-2100^A10930000000073000             830107231544000000090302                0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093"
'              strdata = "D1U   XE-2100^A22030000000024000           1004200307221027000000060442              0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^98313616^A2203"
''    strdata = strdata & "D2U   XE-2100^A10930000000073000             8300461004120013800400009710033500345001950         011900424001250010700304000000000000000000000000000000000210           00000000000000000XE-2100^99337319^A1093"
'    strdata = strdata & "D2U   XE-2100^A22030000000024000           10040078300409001300037600919003180034600213002590006600649000230000300020300005200050800001800000200137004560012300109003130                              00230                             000000000000000000000"
'    Arr = strdata

    Select Case comEQP.CommEvent
        Case comEvReceive
        
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If
            strdata = comEQP.Input
            Call ComReceive(strdata)
            
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

Private Sub ComReceive(ByVal RecData As String)
    
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
                    intPos1 = InStr(strTmp, "D2U")
                    Do While intPos1 > 0
                        intPos2 = InStr(strTmp, Chr(3))
                        intCnt = intCnt + 1
                        ReDim Preserve strdata(1 To intCnt) As String
                        
                        If intPos2 > 0 Then
                            strdata(intCnt) = Mid$(strTmp, intPos1 - 1, intPos2)
                        Else
                            strdata(intCnt) = Mid$(strTmp, intPos1 - 1)
                        End If
                    
                        strTmp = Mid$(strTmp, intPos2 + 1)
                        intPos1 = InStr(strTmp, "D2U")
                    Loop
                    
                    For intPos1 = 1 To intCnt
                        Call psDataDefine(strdata(intCnt), fChannel(), spdResult1, IIf(optSeq.Value, "0", "1"))  ', brSpread, brChannel(), brItemdeci(), brOpt
                    Next
                    
'                    Call COM_OUTPUT(Chr(6))
                    comEQP.Output = Chr(6)
                    f_strBuffer = ""
                    f_strJOB_FLAG = "1"
                            
            Case Else
                    f_strBuffer = f_strBuffer + strBuff
        End Select
     Next
End Sub

Public Function SeqNullSearch(ByVal brSpread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqNullSearch = 0
    If brSpread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brSpread
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

Private Sub psDataDefine(ByVal brbarcd As String, ByRef brChannel() As String, ByVal brSpread As Object, ByVal brOst As String) ' ByRef brItemdeci() As String)

Dim sTemp      As String
Dim Channel_No As Integer       ' 검사항목 번호 : Channel No
Dim pGrid_Point As Integer
Dim pDoCount   As Integer
Dim Loop_Count As Integer
Dim FunStr As String
Dim Max_Arary_Cnt As Integer    ' 검사 항목수
Dim sAdd As Integer, sPosition As Integer
Dim itemX As ListItem
Dim strRstval(1 To 19) As String, strRefVal(1 To 19)  As String, strOrdLst() As String
Dim sqlDoc  As String
Dim intCol As Integer
Dim Gnum   As String
Dim ii As Integer
Dim Test_Cd() As String, strTPid()   As String, strTPnm() As String
Dim Rev As Long
Dim tmpTstCd As String
Dim tmpMXD As Variant
 
Dim strTime As String, sqlRet   As Integer
Dim QCcode() As String
Dim errmsg As String, QCispcid() As String, QCiexamcode() As String, QCiresult() As String, QCierrflag() As String, QCiequipcd() As String, QCigubun As String
Dim pQCDoCount   As Integer
Dim strLevel()  As String

    On Error GoTo ErrRoutine
    
    CallForm = "frmInterface - Privete sub psDataDefine()"
    tmpMXD = "0"
    sTemp = brbarcd
    '------------------------------<<< fXE2100() 배열 Clear 한다.         >>>----------
    For Loop_Count = 1 To 100: fXE2100(Loop_Count) = "": Next Loop_Count
    '------------------------------<<< fXE2100() 배열에 구분하여 넣는다.  >>>----------
        
    If Mid$(sTemp, 2, 3) = "D1U" Then
    
        '  order sending
        
    ElseIf Mid$(sTemp, 2, 3) = "D2U" Then
    
        strTime = Format$(Now, "MMSS")  '-- 같은 Row이면 시간을 동일하게 처리
        
        fXE2100(1) = Mid(sTemp, 21, 10)                   ' 항목 1 "0000000073" 일련번호
        fXE2100(2) = Mid(sTemp, 34, 15)                   ' 항목 2 "             83"   ID 번호(Barcode)
        sPosition = 49
        For sAdd = 3 To 10
            fXE2100(sAdd) = Mid(sTemp, sPosition, fXe2100Size(sAdd, 0))
            sPosition = sPosition + fXe2100Size(sAdd, 0) + 1
        Next sAdd
        Select Case Len(sTemp)
            Case Is <= 209
                        fXE2100(11) = ""
                        fXE2100(12) = ""
                        fXE2100(13) = ""
                        fXE2100(14) = ""
                        fXE2100(15) = ""
                        fXE2100(16) = ""
                        fXE2100(17) = ""
                        fXE2100(18) = ""
                        fXE2100(19) = ""
                        fXE2100(20) = ""
                        fXE2100(21) = Mid(sTemp, 99, fXe2100Size(21, 0))
                        fXE2100(22) = Mid(sTemp, 104, fXe2100Size(22, 0))
                        fXE2100(23) = Mid(sTemp, 109, fXe2100Size(23, 0))
                        fXE2100(24) = Mid(sTemp, 114, fXe2100Size(24, 0))
                        fXE2100(25) = Mid(sTemp, 119, fXe2100Size(25, 0))
                        fXE2100(26) = Mid(sTemp, 154, fXe2100Size(26, 0))
                        fXE2100(27) = ""
                        fXE2100(28) = ""
            Case Is >= 244
                        fXE2100(11) = Mid(sTemp, 90, fXe2100Size(11, 0))  '-- LYMPH%
                        fXE2100(12) = Mid(sTemp, 95, fXe2100Size(12, 0))  '-- MONO%
                        fXE2100(13) = Mid(sTemp, 100, fXe2100Size(13, 0)) '-- NEUT%
                        fXE2100(14) = Mid(sTemp, 105, fXe2100Size(14, 0)) '-- EO%
                        fXE2100(15) = Mid(sTemp, 110, fXe2100Size(15, 0)) '-- BASO%
                        
                        fXE2100(16) = Mid(sTemp, 115, fXe2100Size(16, 0)) '-- LYMPH
                        fXE2100(17) = Mid(sTemp, 121, fXe2100Size(17, 0)) '-- MONO
                        fXE2100(18) = Mid(sTemp, 127, fXe2100Size(18, 0)) '-- NEUT
                        fXE2100(19) = Mid(sTemp, 133, fXe2100Size(19, 0)) '-- EO
                        fXE2100(20) = Mid(sTemp, 139, fXe2100Size(20, 0)) '-- BASO
                        
                        fXE2100(21) = Mid(sTemp, 145, fXe2100Size(21, 0)) '-- RDW-CV
                        fXE2100(22) = Mid(sTemp, 150, fXe2100Size(22, 0)) '-- RDW-SD
                        fXE2100(23) = Mid(sTemp, 155, fXe2100Size(23, 0)) '-- PDW
                        fXE2100(24) = Mid(sTemp, 160, fXe2100Size(24, 0)) '-- MPV
                        fXE2100(25) = Mid(sTemp, 165, fXe2100Size(25, 0)) '-- P-LCR
                        
                        fXE2100(26) = Mid(sTemp, 200, fXe2100Size(26, 0)) '-- PCT
                        fXE2100(27) = Mid(sTemp, 170, fXe2100Size(27, 0)) '-- Ret(%)
                        If Len(sTemp) = 244 Then
                            fXE2100(28) = ""
                            fXE2100(29) = ""
                        Else
                            fXE2100(28) = Mid(sTemp, 210, fXe2100Size(28, 0))
                            fXE2100(29) = Mid(sTemp, 301, fXe2100Size(29, 0))
                        End If
                        
                        fXE2100(28) = Mid(sTemp, 180, fXe2100Size(29, 0)) '-- IRF
                        
                        fXE2100(34) = Val(fXE2100(12)) + Val(fXE2100(14)) + Val(fXE2100(15))
                       
            Case Else
        End Select
        
        '-------------------------------------------<<< 해당검사결과와 해당환자를 찿는다.       >>>----------

        Max_Arary_Cnt = brSpread.MaxCols - 6  ' 앞에서부터 2까지는 환자 정보 이기때문에.... -2를 한다.
                                               ' 해당 배열은  brItem(),brChannel() 이다.
        pGrid_Point = 0
        Dim sSeq As String
        Dim sCol As Integer
        
        With brSpread
            If brOst = 0 Then
                sSeq = Val(fXE2100(1))
                sCol = 0
                pGrid_Point = SeqSearch(brSpread, sSeq, sCol)
            Else
                sSeq = Trim(fXE2100(2))
                sCol = 2
                pGrid_Point = SeqSearch(brSpread, sSeq, sCol)
            End If
            
            If UCase(Mid(sSeq, 1, 2)) <> "QC" And pGrid_Point = 0 Then
                sSeq = IIf(left(sSeq, 1) <> "0", "0" + sSeq, sSeq)
                pGrid_Point = SeqSearch(spdResult1, sSeq, sCol)
            End If
            
            If UCase(Mid(sSeq, 1, 2)) = "QC" And pGrid_Point = 0 Then
                pGrid_Point = SeqNullSearch(spdResult1, sSeq, 2)
            End If
            .Col = 2: .Row = pGrid_Point: Gnum = .Text
            
            If UCase(Mid(sSeq, 1, 2)) = "QC" And Trim$(Gnum) = "" Then
                Gnum = sSeq
                spdResult1.SetText 2, pGrid_Point, sSeq
                spdResult1.SetText 1, pGrid_Point, "1"
            ElseIf Trim$(Gnum) = "" Then
                Gnum = Format$(Now, "YYMMDDHHMMSS")
            End If
            
            If pGrid_Point > 0 Then                ' 해당 대상자를 찿으면 ....
                pQCDoCount = 0
                For pDoCount = 1 To Max_Arary_Cnt + 1 '-------------------------------<<<<<<<<<,  세부검사항목을 찿는다.  >>>>>>>---------
                    .Row = pGrid_Point
                    .Col = pDoCount + 6
                    Channel_No = Val(brChannel(pDoCount))              '  Channel이 숫자이기 때문에 숫자로 치환한다.
                    If Len(fXE2100(Channel_No + 2)) > 0 Then
                        pQCDoCount = pQCDoCount + 1
                        If fXe2100Size(Channel_No + 2, 1) = fXe2100Size(Channel_No + 2, 0) Then
                            FunStr = Trim(Val(fXE2100(Channel_No + 2)))
                        Else
                            FunStr = Trim(Val(Mid$(fXE2100(Channel_No + 2), 1, fXe2100Size(Channel_No + 2, 1)))) + "." + _
                                           Mid$(fXE2100(Channel_No + 2), fXe2100Size(Channel_No + 2, 1) + 1)
                        End If
                        
                        If Val(Channel_No) = 5 Or (Val(Channel_No) >= 9 And Val(Channel_No) <= 13) Then
                            FunStr = Format$(FunStr, "##0")
                        End If
                        
                        If Channel_No = 1 Then FunStr = Format$(FunStr, "##0.0")
                        
                        If FunStr <> "" And Mid$(FunStr, 1, 1) = "." Then FunStr = "0" + FunStr
                                                
                        If Val(FunStr) = 0 Then
                            'FunStr = ""
                        ElseIf Val(FunStr) <> 0 And Val(Channel_No) = 32 Then
                            FunStr = FunStr / 10
                            FunStr = Round(FunStr)
                        End If
                        
                        If Val(Channel_No) = 25 Then
                            FunStr = Format(FunStr, "0.00")
                        End If
                        
                        If IsNumeric(FunStr) Then
                            .Text = IIf(FunStr = ".0", "", FunStr)
                        Else
                            .Text = ""
                        End If
                        
                        If FunStr <> "" Then
                            Set itemX = lvwCuData.FindItem(Channel_No, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then '-- itemX : 검사코드
                                If itemX.ListSubItems(8) <> "" And itemX.ListSubItems(9) <> "" Then
                                    If Val(FunStr) < itemX.ListSubItems(8) Then
                                        strRefVal(pDoCount) = "L"
                                    ElseIf Val(FunStr) > itemX.ListSubItems(9) Then
                                        strRefVal(pDoCount) = "H"
                                    End If
                                End If
                                
                                If chkQC.Value = 1 Then
                                    Rev = sl_spcid_tstcd_select_qc1&(INS_CODE, cboLevel.Text, strTPid, Test_Cd, strLevel)
                                Else
                                    Rev = sl_spcid_tstcd_select&(Trim(Gnum), Test_Cd, strTPid, strTPnm)
                                End If
                                
                                For ii = 0 To Rev - 1
                                    If InStr(itemX.ListSubItems(1), Trim(Test_Cd(ii))) > 0 Then
                                          tmpTstCd = "" & Trim(Test_Cd(ii))
                                          Exit For
                                    End If
                                Next ii
                                            
                                sqlDoc = "Update INTERFACE003" & _
                                         "   set RSTVAL  = '" & FunStr & "', REFVAL = '" & strRefVal(pDoCount) & "'" & _
                                         " where SPCNO   = '" & CStr(Gnum) & "'" & _
                                         "   and EQPNUM  = '" & Channel_No & "'" & _
                                         "   and TRANSDT = '" & Format$(Now, "YYYYMMDD") & "'" & _
                                         "   and TRANSTM = '" & strTime & "'"
                                AdoCn_Jet.Execute sqlDoc, sqlRet
                                
                                If sqlRet = 0 Then
                                    sqlDoc = "insert into INTERFACE003(" & _
                                             "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN)" & _
                                             "    values( '" & CStr(Gnum) & "', '" & itemX.ListSubItems(1) & "', '" & Channel_No & "'," & _
                                             "            '" & Format$(Now, "YYYYMMDD") & "', '" & strTime & "'," & _
                                             "            '" & FunStr & "', '" & strRefVal(pDoCount) & "'," & _
                                             "            '" & INS_CODE & "', '')"
                                    
                                    AdoCn_Jet.Execute sqlDoc
                                End If
                                
                                intCol = itemX.Index
                                '.SetText intCol + 2, sSeq, strRstval(pDoCount)
                                .Col = intCol + 4
'                                .Row = pDoCount
                                .ForeColor = IIf(strRefVal(pDoCount) <> "", vbRed, vbBlack)
                            End If
                            
                                '-- 서버결과등록
                            If Rev > 0 And chkAuto.Value = vbChecked Then
                                If pDoCount = Max_Arary_Cnt + 1 Then Call cmdAppend_Click(0)
                                    spdResult1.Row = f_intSampleNo
                                    spdResult1.Col = -1:    spdResult1.BackColor = &HFFF8F0
                                    
                                    sqlDoc = "Update INTERFACE003 set SERVERGBN  = 'Y'" & _
                                             " where SPCNO   = '" & sSeq & "'" & _
                                             "   and EQPNUM  = '" & itemX.tag & "'" & _
                                             "   and TRANSDT = '" & Format(Now, "YYYYMMDD") & "'" & _
                                             "   and TRANSTM = '" & strTime & "'"
                                    AdoCn_Jet.Execute sqlDoc
                                End If
                            End If
                            
                            Set itemX = Nothing
                            
                        End If
                Next pDoCount
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
        
        Call sl_online_result_ul_4&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
        If strErrMsg = "" Then
            f_funAdd_Server = True
        Else
            Call ErrMsgProc(strErrMsg)
        End If
'    Else
'        Call ErrMsgProc("", "검체번호 [" + strBarno + "]를 저장하지 못했습니다.")
    End If
                                
End Function


Public Function SeqSearch(ByVal brSpread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch = 0
    If brSpread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brSpread
        If brCol = 0 Then
            For sCnt = 1 To .maxrows
                .Row = sCnt
                .Col = brCol
                If Val(.StartingRowNumber + (sCnt - 1)) = brSeq Then
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
'                If Val(.Text) = brSeq Then
                If .Text = brSeq Then
                    SeqSearch = sCnt
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            Next sCnt
        End If
    End With

End Function

Public Function f_funGet_CheckSum(ByVal strPara As String) As String

    Dim intIdx      As Integer
    Dim intChkSum   As Integer
    
    intChkSum = 0
    For intIdx = 1 To Len(strPara)
        intChkSum = intChkSum + (0 Xor Asc(Mid$(strPara, intIdx, 1)))
    Next
    
    f_funGet_CheckSum = Chr(intChkSum) '-Format$(Hex(intChkSum), "00")
        
End Function

Private Sub Command1_Click()

'    Dim intIdx          As Integer
'    Dim strData(1 To 5)         As Byte
'
'                 strData(1) = "D1U   XE-2100^A10930000000073000             830107231544000000090302                0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093"
'    strData(1) = strData(1) & "D2U   XE-2100^A10930000000073000             8300461004120013800400009710033500345001950         011900424001250010700304000000000000000000000000000000000210           00000000000000000XE-2100^99337319^A1093"
'D1U   XE-2100^A10930000000072000             820107231543000000090202                0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000072000             8200510004290013700412009600031900333002500         013000452001150010100264000000000000000000000000000000000250           00000000000000000XE-2100^99337319^A1093
'D1U   XE-2100^A10930000000071000             810107231543000000090102                0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000071000             8100422004160012300378009090029600325001690         012800423001450011800401000000000000000000000000000000000200           00000000000000000XE-2100^99337319^A1093
'D1U   XE-2100^A10930000000070000             800107231543000000081002            0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000070000             8000492004340013700403009290031600340001740         01250042500108000970022000000000000000000000000000000000017000000000000000000XE-2100^99337319^A1093
'D1U   XE-2100^A10930000000073000             830107231544000000090302                0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000073000             8300461004120013800400009710033500345001950         011900424001250010700304000000000000000000000000000000000210           00000000000000000XE-2100^99337319^A1093
'D1U   XE-2100^A10930000000072000             820107231543000000090202                0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000072000             8200510004290013700412009600031900333002500         013000452001150010100264000000000000000000000000000000000250           00000000000000000XE-2100^99337319^A1093
'D1U   XE-2100^A10930000000071000             810107231543000000090102                0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000071000             8100422004160012300378009090029600325001690     012800423001450011800401000000000000000000000000000000000200           00000000000000000XE-2100^99337319^A1093
'D1U   XE-2100^A10930000000070000             800107231543000000081002            0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000070000             8000492004340013700403009290031600340001740         01250042500108000970022000000000000000000000000000000000017000000000000000000XE-2100^99337319^A1093
    
    Dim Arr()   As Byte
    Dim strTmp  As String
    
    strTmp = "D1U   XE-2100^A22030000000013000       Oh SeJoo20030731100200000004010203-6951 ER M35  00000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^98313616^A2203D2U   XE-2100^A22030000000013000       Oh SeJoo0061900502001570044000876003130035700088001232002400849100020000200007620001500052600000100000100132004230014100107003210                              00092                             000000000000000000000D1U   XE-2100^A22030000000012000  Kim KwangSoon20030731100000      000103-6766 6W M82  01110000110000000000000000000000000000000000000000000000000000000000000000000000XE-2100^98313616^A2203D2U   XE-2100^A22030000000012000  Kim KwangSoon008040031000087002582083220281003370021900121401034073540039000020000974000834005914000310000020015100462001060009900233000720022300079009210007500004000220            00001000010      000000000000000000000"
    
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
    Open App.Path + "\" + "dump_job.log" For Append As #1

    f_strJOB_FLAG = "1":    f_intSampleNo = 0
    cboRstgbn(0).ListIndex = 0: cboRstgbn(1).ListIndex = 2
    
    fXe2100Size(3, 0) = 5:      fXe2100Size(3, 1) = 3       ' WBC
    fXe2100Size(4, 0) = 4:      fXe2100Size(4, 1) = 2       ' RBC
    fXe2100Size(5, 0) = 4:      fXe2100Size(5, 1) = 3       ' HGB
    fXe2100Size(6, 0) = 4:      fXe2100Size(6, 1) = 3       ' HCT
    fXe2100Size(7, 0) = 4:      fXe2100Size(7, 1) = 3       ' MCV
    fXe2100Size(8, 0) = 4:      fXe2100Size(8, 1) = 3       ' MCH
    fXe2100Size(9, 0) = 4:      fXe2100Size(9, 1) = 3       ' MCHC
    fXe2100Size(10, 0) = 4:     fXe2100Size(10, 1) = 4      ' PLT
    fXe2100Size(11, 0) = 4:     fXe2100Size(11, 1) = 3      ' LYMP%
    fXe2100Size(12, 0) = 4:     fXe2100Size(12, 1) = 3      ' MONO%
    fXe2100Size(13, 0) = 4:     fXe2100Size(13, 1) = 3      ' NEUT%
    fXe2100Size(14, 0) = 4:     fXe2100Size(14, 1) = 3      ' EO%
    fXe2100Size(15, 0) = 4:     fXe2100Size(15, 1) = 3      ' BASO%
    fXe2100Size(16, 0) = 5:     fXe2100Size(16, 1) = 3      ' LYMPH#
    fXe2100Size(17, 0) = 5:     fXe2100Size(17, 1) = 3      ' MONO#
    fXe2100Size(18, 0) = 5:     fXe2100Size(18, 1) = 3      ' NEUT#
    fXe2100Size(19, 0) = 5:     fXe2100Size(19, 1) = 3      ' EO#
    fXe2100Size(20, 0) = 5:     fXe2100Size(20, 1) = 3      ' BASO#
    fXe2100Size(21, 0) = 4:     fXe2100Size(21, 1) = 3      ' RDW-CV
    fXe2100Size(22, 0) = 4:     fXe2100Size(22, 1) = 3      ' RDW-SD
    fXe2100Size(23, 0) = 4:     fXe2100Size(23, 1) = 3      ' PDW
    fXe2100Size(24, 0) = 4:     fXe2100Size(24, 1) = 3      ' MPV
    fXe2100Size(25, 0) = 4:     fXe2100Size(25, 1) = 3      ' P-LCR
    fXe2100Size(26, 0) = 4:     fXe2100Size(26, 1) = 2      ' PCT
    fXe2100Size(27, 0) = 5:     fXe2100Size(27, 1) = 2      ' NRBC% ?NEUT?
    fXe2100Size(28, 0) = 5:     fXe2100Size(28, 1) = 3      ' NRBC#
    
    fXe2100Size(29, 0) = 4:     fXe2100Size(29, 1) = 3      ' IRF
    
'    fXe2100Size(26, 0) = 4:      fXe2100Size(26, 1) = 2      ' RET(%)
'    fXe2100Size(27, 0) = 4:      fXe2100Size(27, 1) = 2      ' RET(#)
'    fXe2100Size(28, 0) = 4:      fXe2100Size(28, 1) = 2      ' IRF(%)
'    fXe2100Size(29, 0) = 4:      fXe2100Size(29, 1) = 2      ' LFR(%)
'    fXe2100Size(30, 0) = 4:      fXe2100Size(30, 1) = 2      ' MFR(%)
'    fXe2100Size(31, 0) = 4:      fXe2100Size(31, 1) = 2      ' HFR(%)
'    fXe2100Size(32, 0) = 4:      fXe2100Size(32, 1) = 2      ' PCT
'    fXe2100Size(33, 0) = 5:      fXe2100Size(33, 1) = 2      ' NRBC%
'    fXe2100Size(34, 0) = 5:      fXe2100Size(34, 1) = 3      ' NRBC#

    cboLevel.Clear
    cboLevel.AddItem "H"
    cboLevel.AddItem "M"
    cboLevel.AddItem "L"
    
End Sub

Private Sub f_subGet_Setting()
    
    Dim objComSetting As clsCommon
    Dim Baudratio As String
    Dim Paritybit As String
    Dim Databit As String
    Dim Stopbit As String
    
    Dim RetVal As Long
    Dim sBuf As String
    Dim com_port As String
    Dim set_ting As String
    
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
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("OrderPort", "CommPort", "No", sBuf, 255, App.Path & "\orderport.ini")
    com_port = left(sBuf, RetVal)
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("OrderPort", "Settings", "No", sBuf, 255, App.Path & "\orderport.ini")
    set_ting = left(sBuf, RetVal)
    'Exit Sub
    With comOrder
        If .PortOpen = False Then
            .CommPort = com_port
            '속도,페리티,테이타비트,stop bit
            .Settings = set_ting
            .RThreshold = 1
            .SThreshold = 1
            '-- 임시 : 노트북이 com1만 있으므로...
            'If comEQP.PortOpen = True Then comEQP.PortOpen = False
            '--------------------------------------------
            If .PortOpen = True Then .PortOpen = False
            .PortOpen = True
        End If
    End With
    
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
    
End Sub

Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Col > 4 Then
        spdResult1.Col = Col
        spdResult1.Row = Row
        If spdResult1.BackColor = &HC6FEFF Then
            spdResult1.BackColor = vbWhite
        Else
            spdResult1.BackColor = &HC6FEFF
        End If
    End If

End Sub

Private Sub spdResult1_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Col > 4 Then
        spdResult1.Col = Col
        spdResult1.Row = Row
        spdResult1.CellType = CellTypeEdit
'        spdResult1.EditMode = True
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
    Dim tst_no() As String, strTPid() As String, strTPnm() As String
    Dim TMP() As String
    Dim rv As Long
    Dim samChk As Boolean
    Dim ii As Integer
    Dim bgetWork As Boolean
    Dim itemX As ListItem
    
    Dim strEqpCd    As String
    Dim strOrdcd() As String
    Exit Sub
    samChk = False
    If KeyAscii = vbKeyReturn Then
'        rv = sl_spcid_tstcd_select(Trim(txtBarCode.Text), tst_no, strTPid, strTPnm)
        If chkQC.Value = 1 Then
            rv = sl_spcid_tstcd_select_qc&(Trim(txtBarCode.Text), INS_CODE, strOrdcd)
        Else
            rv = sl_spcid_tstcd_select(Trim(txtBarCode.Text), tst_no, strTPid, strTPnm)
        End If
        
        If (rv = 0) Then
            MsgBox "미접수 검체입니다.!", vbCritical
        Else
            If psDataExists Then
                MsgBox "이미 등록된 검체입니다.!", vbCritical
                txtBarCode.Text = ""
                Exit Sub
            End If
            
            bgetWork = False
            For ii = 0 To rv - 1
                strEqpCd = f_funGet_CODE(tst_no(ii))
                Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                If Not itemX Is Nothing Then
                    bgetWork = True
                End If
            Next
                    
             With spdWorkList
                If bgetWork = True Then
                    .Col = 2
                    For ii = 1 To .maxrows
                        .Row = ii
                        If Trim(.Text) = "" Then
                            .Text = txtBarCode.Text
                            txtBarCode.Text = ""
                            .Col = 1
                            .Value = 1
                            samChk = True
                            Exit For
                        End If
                    Next
                    If samChk = False Then
                         .maxrows = .maxrows + 1
                         .Row = .maxrows
                         .Text = txtBarCode.Text
                         txtBarCode.Text = ""
                    End If
                Else
                   MsgBox "해당검사항목이 존재하지 않는 검체입니다.", vbOKOnly + vbInformation, Me.Caption
                End If
             End With
        End If
    End If
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

Private Sub txtBarCode_LostFocus()
    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Privete sub txtBarCode_LostFocus()"
    
    Dim varTmp  As Variant, strEqpCd    As String
    Dim intRow  As Integer, intCol  As Integer, blnFlag As Boolean
    Dim strOrdcd() As String, strPid()  As String, strPnm() As String
    
    Dim itemX   As ListItem
    
    If txtBarCode.Text = "" Then Exit Sub
    
    blnFlag = False
    intCol = sl_spcid_tstcd_select&(txtBarCode.Text, strOrdcd, strPid, strPnm)
    
    For intCol = 0 To UBound(strOrdcd)
        If strOrdcd(intCol) <> "" Then
            strEqpCd = f_funGet_CODE(strOrdcd(intCol))
            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
            If Not itemX Is Nothing Then
                If Not blnFlag Then
                    intRow = f_funGet_SpreadRow(spdWorkList, 2, txtBarCode.Text)
                    If intRow < 1 Then
                        intRow = f_funGet_SpreadRow(spdWorkList, 2, "")
                        If intRow < 1 Then
                            spdWorkList.maxrows = spdWorkList.maxrows + 1
                            spdWorkList.RowHeight(spdWorkList.maxrows) = 13
                            intRow = spdWorkList.maxrows
                        End If
                        spdWorkList.SetText 2, intRow, txtBarCode.Text
                        spdWorkList.SetText 3, intRow, strPnm(0)
                        spdWorkList.SetText 4, intRow, strPid(0)
                    End If
                    spdWorkList.SetText 1, intRow, "1"
                End If
                    
                spdWorkList.SetText itemX.Index + 6, intRow, "V"
                spdWorkList.Col = itemX.Index + 6
                spdWorkList.Row = intRow
                spdWorkList.BackColor = &HC6FEFF
                
                blnFlag = True
            End If
        End If
    Next
    
    If Not blnFlag Then MsgBox "해당 검사항목이 존재하지 않은 검체입니다.", vbInformation, Me.Caption
    
    txtBarCode.Text = "":   txtBarCode.SetFocus
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)
    
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

    Call ComReceive(bytTemp)
End Sub

Private Sub cmdCOMInput2_Click()
    
    Dim bytTemp() As Byte
    
    If txtCOM2.SelLength = 0 Then
        bytTemp = StrConv(charCOM_Convert(txtCOM2.Text), vbFromUnicode)
    Else
        bytTemp = StrConv(charCOM_Convert(txtCOM2.SelText), vbFromUnicode)
    End If

    Call ComReceive(bytTemp)

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



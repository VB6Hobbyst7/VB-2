VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
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
      Left            =   9660
      ScaleHeight     =   345
      ScaleWidth      =   2265
      TabIndex        =   50
      Top             =   540
      Width           =   2295
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
         Left            =   180
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   60
         Width           =   720
      End
      Begin VB.ComboBox cboLevel 
         Height          =   300
         Left            =   1110
         TabIndex        =   51
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
      Left            =   3825
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
      Left            =   90
      TabIndex        =   1
      Top             =   6480
      Width           =   11940
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   0
         Top             =   0
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
      Left            =   90
      TabIndex        =   11
      Top             =   615
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
      Tab(0).Control(1)=   "pnlCom"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cboRstgbn(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdReceive"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "pnlCom2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdWorkList"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdAppend(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "spdResult1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "spdWorkList"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdSel(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdSel(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdWorkQuery"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdWKSend"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "mskOrdDate"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtBarCode"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkAuto"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   " 받은 결과"
      TabPicture(1)   =   "frmComm_2.frx":6246
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "spdResult2"
      Tab(1).Control(2)=   "cmdQuery"
      Tab(1).Control(3)=   "cmdAppend(1)"
      Tab(1).Control(4)=   "mskRstDate"
      Tab(1).Control(5)=   "lvwCuData"
      Tab(1).Control(6)=   "cboRstgbn(1)"
      Tab(1).ControlCount=   7
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
         Left            =   5940
         TabIndex        =   49
         Top             =   495
         Value           =   1  '확인
         Width           =   1410
      End
      Begin VB.ComboBox cboRstgbn 
         Height          =   300
         Index           =   1
         ItemData        =   "frmComm_2.frx":6262
         Left            =   -72525
         List            =   "frmComm_2.frx":626F
         Style           =   2  '드롭다운 목록
         TabIndex        =   17
         Top             =   495
         Width           =   2085
      End
      Begin VB.TextBox txtBarCode 
         Height          =   300
         Left            =   2475
         MaxLength       =   11
         TabIndex        =   14
         Top             =   450
         Width           =   2085
      End
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   5415
         Left            =   -71265
         TabIndex        =   13
         Top             =   405
         Visible         =   0   'False
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   9551
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
      Begin MSMask.MaskEdBox mskRstDate 
         Height          =   300
         Left            =   -73650
         TabIndex        =   18
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
         TabIndex        =   19
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
      Begin HSCotrol.CButton cmdQuery 
         Height          =   300
         Left            =   -65460
         TabIndex        =   20
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
         Left            =   1350
         TabIndex        =   21
         Top             =   450
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin HSCotrol.CButton cmdWKSend 
         Height          =   300
         Left            =   7425
         TabIndex        =   22
         Top             =   450
         Width           =   1080
         _ExtentX        =   1905
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
      Begin HSCotrol.CButton cmdWorkQuery 
         Height          =   300
         Left            =   9585
         TabIndex        =   23
         Top             =   450
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
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   1
         Left            =   360
         TabIndex        =   26
         Top             =   855
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_2.frx":6299
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   0
         Left            =   90
         TabIndex        =   27
         Top             =   855
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_2.frx":671B
      End
      Begin FPSpread.vaSpread spdWorkList 
         Height          =   4515
         Left            =   90
         TabIndex        =   16
         Top             =   855
         Width           =   2310
         _Version        =   196608
         _ExtentX        =   4075
         _ExtentY        =   7964
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditEnterAction =   2
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
         MaxCols         =   22
         MaxRows         =   14
         ScrollBarMaxAlign=   0   'False
         SpreadDesigner  =   "frmComm_2.frx":6B89
         UserResize      =   0
      End
      Begin FPSpread.vaSpread spdResult1 
         Height          =   4830
         Left            =   2430
         TabIndex        =   28
         Top             =   855
         Width           =   9330
         _Version        =   196608
         _ExtentX        =   16457
         _ExtentY        =   8520
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   4
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   22
         MaxRows         =   14
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmComm_2.frx":72D3
         UserResize      =   0
      End
      Begin HSCotrol.CButton cmdAppend 
         Height          =   300
         Index           =   0
         Left            =   10665
         TabIndex        =   29
         Top             =   450
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
      Begin HSCotrol.CButton cmdWorkList 
         Height          =   300
         Left            =   90
         TabIndex        =   30
         Top             =   5400
         Width           =   2310
         _ExtentX        =   4075
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
      Begin HSCotrol.UserPanel pnlCom2 
         Height          =   5340
         Left            =   5940
         TabIndex        =   38
         Top             =   450
         Visible         =   0   'False
         Width           =   5880
         _ExtentX        =   10372
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
            TabIndex        =   47
            Top             =   270
            Width           =   5730
         End
         Begin VB.Frame Frame2 
            Height          =   645
            Left            =   60
            TabIndex        =   39
            Top             =   4650
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
               TabIndex        =   40
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
               TabIndex        =   41
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
               TabIndex        =   42
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
               TabIndex        =   43
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
               TabIndex        =   44
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
               TabIndex        =   45
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
               TabIndex        =   46
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
      Begin HSCotrol.CButton cmdReceive 
         Height          =   300
         Left            =   8505
         TabIndex        =   48
         Top             =   450
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         Caption         =   "결과재송신"
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
      Begin VB.ComboBox cboRstgbn 
         Height          =   300
         Index           =   0
         ItemData        =   "frmComm_2.frx":7A16
         Left            =   2475
         List            =   "frmComm_2.frx":7A23
         Style           =   2  '드롭다운 목록
         TabIndex        =   12
         Top             =   450
         Visible         =   0   'False
         Width           =   2085
      End
      Begin HSCotrol.UserPanel pnlCom 
         Height          =   5310
         Left            =   45
         TabIndex        =   31
         Top             =   450
         Visible         =   0   'False
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   9366
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
            TabIndex        =   33
            Top             =   4650
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
            TabIndex        =   32
            Top             =   270
            Width           =   11595
         End
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
         ColsFrozen      =   4
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   22
         MaxRows         =   14
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmComm_2.frx":7A4D
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
         Left            =   -74865
         TabIndex        =   25
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
         Left            =   135
         TabIndex        =   24
         Top             =   525
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

Private f_strBuffer         As String
Private f_intPendFlag       As Integer
Private f_intIdleFlag       As Integer
Private f_intOrderFlag      As Integer
Private f_intReceiveFlag    As Boolean
Private f_strOrdList        As String
Private f_strBarno()        As String, f_strTest()  As String
Private f_strRack()         As String, f_strCup()   As String
Private f_intCnt            As Integer

Private MSG_STX     As String
Private MSG_ETX     As String
Private MSG_ENQ     As String
Private MSG_EOT     As String
Private MSG_ACK     As String
Private MSG_NAK     As String
Private MSG_CR      As String
Private MSG_LF      As String
Private MSG_CRLF    As String

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

Private Sub f_subSet_Result(ByVal strdata As String)

    Dim sqlDoc      As String, sqlRet   As Integer

    Dim strSendBuff As String
    Dim strSampleno As String, strDate      As String, strTime      As String
    Dim strRack     As String, strCup       As String, strEqpCd     As String, strTestcd    As String
    Dim strRstval   As String, strRefVal    As String, strRstEtc1   As String, strRstEtc2   As String
    Dim intRow      As Integer, intCol      As Integer
    
    Dim intBcPos    As Integer, intLcPos    As Integer
    Dim strBC       As String, strLC        As String, intErrCode As Integer
    Dim intResultFlag   As Integer
    
    Dim itemX   As ListItem
    
    Dim strOrdLst() As String, strPid()  As String, strPnm() As String
    Dim intRet   As Integer
    
    Dim strOrdcd()  As String, strBarno() As String
    Dim strLevel()  As String
    
    On Error GoTo ErrRoutine
    
    CallForm = "frmInterface - Privete sub f_subSet_Result()"
    
    intBcPos = 22
    strBC = Mid$(strdata, intBcPos, 2)
    
    'Order manipulation response
    If strBC = "19" Then intErrCode = 99
    
    If strBC = "99" Then
        strSendBuff = ""
        strSendBuff = Chr(1) & Chr(10) + _
                      "06" & " " & "HOSTNAMESIXTEENX" & " " & "00" & Chr(10) + _
                      Chr(2) & Chr(10) + _
                      Chr(3) & Chr(10) + _
                      Chr(4) & Chr(10)
    
        Call COM_OUTPUT(strSendBuff)
        Exit Sub
    End If
    
    intLcPos = intBcPos + 5

    Do
        If Asc(Mid$(strdata, intLcPos, 1)) = 3 Then Exit Do  'ETX(END OF DATA BLOCK)
        
        strLC = Mid$(strdata, intLcPos, 2)
        Select Case strLC
            Case "00"   '-- FINAL RESULT
                        Me.MousePointer = 11
                        strRstval = Mid$(strdata, intLcPos + 3, 9)
                        
'                        If Format(Val(Mid$(strRstVal, 3)), "####0.##") = 0 Then
                        strRstval = Mid$(strRstval, 1, 1) + Format$(Val(Mid$(strRstval, 3)), "####0.0#")
'                        Else
'                            strRstVal = Mid$(strRstVal, 1, 1) + Format$(Val(Mid$(strRstVal, 3)), "####0.0")
'                        End If
                        
'                        If Trim$(strRstEtc2) = "qal" Then
'                            If Val(strRstVal) > 100 Then
'                                strRstVal = "positive"
'                            Else
'                                strRstVal = "negative"
'                            End If
'                        End If
                        
                        intLcPos = intLcPos + 22
                        intResultFlag = 1
                        Exit Do
            Case "01"   '-- BLANK RESULT
                        Exit Do
            Case "02"   '-- Absorbance value
                        Exit Do
            Case "03"   '-- CAL/KC IDENTIFICATION
                        Exit Do
            Case "04"   '-- CAL/KC NUMBER
                        Exit Do
            Case "05"   '-- CURVE DATA
                        Exit Do
            Case "06":  Exit Do
            Case "08":  Exit Do
            Case "09"   '-- RESULT SPEC. - SHORT/RAW, QUAL/QUAN
                        strRstEtc1 = Mid$(strdata, intLcPos + 3, 1)
                        strRstEtc2 = Mid$(strdata, intLcPos + 5, 3)
                        intLcPos = intLcPos + 9
            Case "11"   '-- SAMPLE IDENTIFICATION
                        Me.MousePointer = 11
                        strSampleno = Trim$(Mid$(strdata, intLcPos + 17, 20))
                        strRack = Mid$(strdata, intLcPos + 10, 3)
                        strCup = Mid$(strdata, intLcPos + 14, 2)
                        intLcPos = intLcPos + 38
            Case "12"   '-- TEST NUMBER
                        strEqpCd = Mid$(strdata, intLcPos + 3, 2)
                        intLcPos = intLcPos + 6
            Case "14"   '-- Error
                        Exit Do
            Case "42"   '-- Smaple Processing에서 펜딩 SAMPLE ID
            Case "59"   '-- ERROR CHECK
                        If Mid$(strdata, intLcPos + 3, 2) = "00" Then
                            intErrCode = 0     'Order Input Accepted
'                            If f_intCnt > 0 And UBound(f_strBarno) <= f_intCnt Then
'                                intRow = f_funGet_SpreadRow(spdResult1, 2, f_strBarno(f_intCnt))
'                                If intRow > 0 Then
'                                    spdResult1.Row = intRow:  spdResult1.Row2 = intRow
'                                    spdResult1.Col = 3:       spdResult1.Col2 = 4
'                                    spdResult1.BlockMode = True
'                                    spdResult1.BackColor = vbCyan
'                                    spdResult1.BlockMode = False
'                                End If
'                            End If
                            Exit Do
                        Else
                            If Mid$(strdata, intLcPos + 9, 2) = "04" Then
                                intErrCode = 1     'Order already available
                                Exit Do
                            Else
                                intErrCode = 2     '기타 검사항목 초과, 검사 중 등의 에러
                                Exit Do
                            End If
                        End If
            Case Else:  Exit Do
            
        End Select
    Loop
    
'### Pending Sample Request ##############################################
    If f_intPendFlag = 1 And strBC = "62" Then
        strSendBuff = ""
        strSendBuff = Chr(1) & Chr(10) + _
                      "06" & " " & "HOSTNAMESIXTEENX" & " " & "60" & Chr(10) + _
                      Chr(2) & Chr(10) + _
                      "40" & " " & "1" & Chr(10) + _
                      Chr(3) & Chr(10) + _
                      Chr(4) & Chr(10)
    
        Call COM_OUTPUT(strSendBuff)
    End If
    
'### CONNECTION CHECK ##########################################################
    If f_intIdleFlag = 1 And strBC = "00" Then
        f_intIdleFlag = 0
        Timer1.Enabled = True
    End If
    
'### NO MORE PENDING SAMPLE #####################################################
    If f_intPendFlag = 1 And strBC = "00" Then f_intPendFlag = 0
    
'### ORDER INPUT RESPONSE ################################################################
    'f_intOrderFlag = 1 --> From Host To CORE : Sample Order 내린 상태
    'f_intOrderFlag = 2 --> From Host To CORE : Order Delete를 요청한 상태
    'f_intOrderFlag = 0 --> Order 전송이 제대로 끝난 상태
    
    If strBC = "19" And intErrCode = 0 Then
        If f_intOrderFlag = 1 Then
            f_intOrderFlag = 0      '-- Order 전송이 제대로 끝난 상태
            f_intCnt = f_intCnt + 1
            If f_intCnt <= UBound(f_strBarno) And f_strTest(1) <> "" Then
                strSendBuff = Chr(1) & Chr(10) + _
                              "06 HOSTNAMESIXTEENX 10" & Chr(10) + _
                              Chr(2) & Chr(10) + _
                              "50 R " & f_strRack(f_intCnt) & "/" & f_strCup(f_intCnt) & Chr(10) + _
                              "51 " & f_strBarno(f_intCnt) + Space(20 - Len(f_strBarno(f_intCnt))) & Chr(10) + _
                              "52 " & f_strTest(f_intCnt) + Chr(10) + _
                              Chr(3) & Chr(10) + _
                              Chr(4) & Chr(10)
                Call COM_OUTPUT(strSendBuff)
                f_intOrderFlag = 1
            End If
        ElseIf f_intOrderFlag = 2 Then
            Call cmdWKSend_Click
        End If
    ElseIf strBC = "19" And intErrCode = 1 Then
        
        If MsgBox("지금 Rack/Pos에 Order가 이미 존재합니다. 이번 Order로 바꾸시겠습니까?" & Chr(10) & _
                  "바꾸시면 이전 Order는 삭제됩니다!!", 4, "Old Order 삭제 & New Order 전송 여부") = 7 Then Exit Sub

        Call f_subDel_Order(f_strBarno(f_intCnt), f_strRack(f_intCnt), f_strCup(f_intCnt))
        
    ElseIf strBC = "19" And intErrCode = 2 Then
        
        MsgBox "Order 거부!! 검사항목수가 많거나 검사항목이 잘못되었습니다."
        Exit Sub
        
    End If
    
'### SAMPLE RESULT 보기& 등록 #####################################################
    If strSampleno <> "" And strEqpCd <> "" Then
        If intResultFlag = 1 And strBC = "04" Then     'BLOCKCODE 04 : SAMPLE RESULT
            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
            If Not itemX Is Nothing Then
                strTestcd = itemX.ListSubItems(1)
                
                If Trim$(strRstEtc2) = "qal" Then
                    If itemX.ListSubItems(8) <> "" Then
                        If Val(strRstval) < Val(itemX.ListSubItems(8)) Then
                            strRstval = "positive"
                        Else
                            strRstval = "negative"
                        End If
                    ElseIf itemX.ListSubItems(9) <> "" Then
                        If Val(strRstval) > Val(itemX.ListSubItems(9)) Then
                            strRstval = "positive"
                        Else
                            strRstval = "negative"
                        End If
                    End If
                End If
                
                intCol = itemX.Index
                intRow = f_funGet_SpreadRow(spdResult1, 2, strSampleno)
                If intRow < 1 Then
                    intRow = f_funGet_SpreadRow(spdResult1, 2, "")
                    If intRow < 1 Then
                        spdResult1.maxrows = spdResult1.maxrows + 1
                        spdResult1.RowHeight(spdResult1.maxrows) = 13
                        intRow = spdResult1.maxrows
                    End If
                    spdResult1.SetText 2, intRow, strSampleno
                    
                End If
                
                '-- osw 추가
'                If InStr(strRstval, "<") > 0 Then
'                    spdResult1.SetText intCol + 6, intRow, "<0.01"
'                Else
'                    spdResult1.SetText intCol + 6, intRow, strRstval
'                End If
                '-- osw 재수정
                If InStr(strRstval, "<") > 0 Then strRstval = "<0.01"
                
                spdResult1.SetText intCol + 6, intRow, strRstval
                
                spdResult1.Col = intCol + 6
                spdResult1.Row = intRow
                spdResult1.ForeColor = IIf(strRefVal <> "", vbRed, vbBlack)
            
                strDate = Format$(Now, "YYYYMMDD"): strTime = Format$(Now, "MMSS")
                sqlDoc = "Update INTERFACE003" & _
                         "   set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRefVal & "'" & _
                         " where SPCNO   = '" & strSampleno & "'" & _
                         "   and EQPNUM  = '" & itemX.tag & "'" & _
                         "   and TRANSDT = '" & strDate & "'" & _
                         "   and TRANSTM = '" & strTime & "'"
                AdoCn_Jet.Execute sqlDoc, sqlRet
                If sqlRet = 0 Then
                    sqlDoc = "insert into INTERFACE003(" & _
                             "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN)" & _
                             "    values( '" & strSampleno & "', '" & strTestcd & "', '" & itemX.tag & "'," & _
                             "            '" & strDate & "', '" & strTime & "'," & _
                             "            '" & strRstval & "', '" & strRefVal & "'," & _
                             "            '" & INS_CODE & "', '')"
                    AdoCn_Jet.Execute sqlDoc
                End If
                
                '-- 서버결과등록
                'intRet = sl_spcid_tstcd_select&(strSampleno, strOrdLst, strPid, strPnm)
                
                If chkQC.Value = 1 Then
                    intRet = sl_spcid_tstcd_select_qc1&(INS_CODE, cboLevel.Text, strBarno, strOrdcd, strLevel)
                Else
                    intRet = sl_spcid_tstcd_select&(strSampleno, strOrdLst, strPid, strPnm)
                End If
                
'                If intRet > 0 And chkAuto.Value = vbChecked Then
'                    If f_funAdd_Server(strSampleno, strTestcd, strRstval, strOrdLst) Then
'                        spdResult1.Row = intRow
'                        spdResult1.Col = -1:    spdResult1.BackColor = &HFFF8F0
'
'                        sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
'                                 " where SPCNO   = '" & strSampleno & "'" & _
'                                 "   and EQPNUM  = '" & itemX.tag & "'" & _
'                                 "   and TRANSDT = '" & strDate & "'" & _
'                                 "   and TRANSTM = '" & strTime & "'"
'                        AdoCn_Jet.Execute sqlDoc, sqlRet
'                    End If
'                End If
                
                If intRet > 0 And chkAuto.Value = vbChecked And chkQC.Value = 0 Then
                    If f_funAdd_Server(strSampleno, strTestcd, strRstval, strOrdLst) Then
                        spdResult1.Row = intRow
                        spdResult1.Col = -1:    spdResult1.BackColor = &HFFF8F0
                    
                        sqlDoc = "Update INTERFACE003" & _
                                 "   set SERVERGBN  = 'Y'" & _
                                 " where SPCNO   = '" & strSampleno & "'" & _
                                 "   and EQPNUM  = '" & itemX.tag & "'" & _
                                 "   and TRANSDT = '" & strDate & "'" & _
                                 "   and TRANSTM = '" & strTime & "'"
                        AdoCn_Jet.Execute sqlDoc
                    End If
                Else
                    If f_funAdd_QcServer(strSampleno, strTestcd, strRstval, strOrdLst) Then
                        spdResult1.Row = intRow
                        spdResult1.Col = -1:    spdResult1.BackColor = &HFFF8F0
                    
                        sqlDoc = "Update INTERFACE003" & _
                                 "   set SERVERGBN  = 'Y'" & _
                                 " where SPCNO   = '" & strSampleno & "'" & _
                                 "   and EQPNUM  = '" & itemX.tag & "'" & _
                                 "   and TRANSDT = '" & strDate & "'" & _
                                 "   and TRANSTM = '" & strTime & "'"
                        AdoCn_Jet.Execute sqlDoc
                    End If
                End If
            End If
            Set itemX = Nothing
            Timer1.Enabled = True
        End If
    Else
        If intResultFlag = 1 And strBC = "04" Then
            intResultFlag = 0
            Timer1.Enabled = True
        End If
    End If
    
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)
    
End Sub
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



Private Sub f_subDel_Order(ByVal strBarno As String, ByVal strRack As String, ByVal strCup As String)

    Dim strSendBuff As String
    
    strSendBuff = Chr(1) & Chr(10) + _
                  "06 HOSTNAMESIXTEENX 11" & Chr(10) + _
                  Chr(2) & Chr(10) + _
                  "54 " & strRack + "/" + strCup & Chr(10) + _
                  Chr(3) & Chr(10) + _
                  Chr(4) & Chr(10)                    ' EOT & LF
    
    Call COM_OUTPUT(strSendBuff)
    
    f_intOrderFlag = 2   'Order Delete 요구를 전송했을 때
    
End Sub

Private Sub cmdQuery_Click()

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String, intRet   As Integer
    
    Dim strSpcno    As String
    Dim intRow      As Integer, intCol  As Integer
    
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
    Dim itemX       As ListItem

    intRow = 0
    With spdResult2
        .maxrows = 14
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
    End With
    
    sqlDoc = "select SPCNO, TESTCD, EQUIPCD, TRANSTM, RSTVAL, REFVAL" & _
             "  from INTERFACE003" & _
             " where TRANSDT = '" & mskRstDate.Text & "'"
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
            If strSpcno <> Trim$(adoRS(0) & "") + Trim$(adoRS(3) & "") Then
                intRow = intRow + 1
                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                
                intRet = sl_spcid_tstcd_select&(Trim$(adoRS(0) & ""), strOrdcd, strPid, strPnm)

                .SetText 1, intRow, "1"
                .SetText 2, intRow, Trim$(adoRS(0) & "")
                .SetText 5, intRow, strPnm(0)
                .SetText 6, intRow, strPid(0)
                .SetText .MaxCols, intRow, Trim$(adoRS(3) & "")
            End If
            strSpcno = Trim$(adoRS(0) & "") + Trim$(adoRS(3) & "")
            Set itemX = lvwCuData.FindItem(Trim$(adoRS(1) & ""), lvwSubItem, , lvwWhole)
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

Private Sub cmdReceive_Click()

    Dim varTmp      As Variant
    Dim strBuffer   As String
    Dim strBarno    As String, strRack  As String, strCup   As String
    Dim intRow      As Integer, intCol  As Integer

    Dim intCnt      As Integer
    Dim itemX   As ListItem
    
    '-- 02/03/12/13/10/11/14/15
    
    intCnt = 0
    Erase f_strBarno:   Erase f_strRack:    Erase f_strCup: Erase f_strTest

    ReDim Preserve f_strBarno(1 To 1) As String
    ReDim Preserve f_strRack(1 To 1) As String
    ReDim Preserve f_strCup(1 To 1) As String
    ReDim Preserve f_strTest(1 To 1) As String
    
    f_strTest(1) = ""
    
    For intRow = 1 To spdResult1.maxrows
        spdResult1.GetText 2, intRow, varTmp:  strBarno = Trim$(varTmp)
        spdResult1.GetText 3, intRow, varTmp:  strRack = Format$(Trim$(varTmp), "000")
        spdResult1.GetText 4, intRow, varTmp:  strCup = Format$(Trim$(varTmp), "00")
        spdResult1.GetText 1, intRow, varTmp
        
        If Trim$(varTmp) = "1" And strRack <> "" Then
            For intCol = 5 To spdResult1.MaxCols
                spdResult1.Col = intCol:    spdResult1.Row = intRow
                If spdResult1.BackColor = &HC6FEFF Then
                    spdResult1.GetText intCol, 0, varTmp
                    
                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                    If Not itemX Is Nothing Then
                        intCnt = intCnt + 1
            
                        ReDim Preserve f_strBarno(1 To intCnt) As String
                        ReDim Preserve f_strRack(1 To intCnt) As String
                        ReDim Preserve f_strCup(1 To intCnt) As String
                        ReDim Preserve f_strTest(1 To intCnt) As String
                        
                        f_strBarno(intCnt) = strBarno
                        f_strRack(intCnt) = strRack:    f_strCup(intCnt) = strCup
                                    
                        f_strTest(intCnt) = itemX.tag + " "
                    End If
                    Set itemX = Nothing
                End If
            Next
        End If
    Next
    
    f_intCnt = 0
    f_intReceiveFlag = True
    
End Sub

Private Sub cmdSel_Click(Index As Integer)

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    With spdWorkList
        For intRow = 1 To .maxrows
            .GetText 2, intRow, varTmp
            
            If Trim$(varTmp) = "" Then Exit For
            
            .SetText 1, intRow, IIf(Index = 0, "1", "")
        Next
    End With
    
            
End Sub

Private Sub cmdWorkQuery_Click()

    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Privete sub cmdWorkQuery_Click()"
    
    Dim strKeyno    As String
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String, strBarno()   As String
    Dim strTestcd() As String, strTPid()    As String, strTPnm() As String
    Dim strEqpCd    As String
    Dim intRow  As String, intIdx  As Integer, intCol   As Integer
    Dim itemX   As ListItem
    
    Dim strLevel() As String
    
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
'                If chkQC.Value = 1 Then
'                    intCol = sl_spcid_tstcd_select_qc&(INS_CODE, strBarno(intIdx), strTestcd)
'                Else
'                    intCol = sl_spcid_tstcd_select(strBarno(intIdx), strTestcd, strTPid, strTPnm)
'                End If
'
'                strEqpCd = f_funGet_Code(strTestcd(intIdx))
'                Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
'                If Not itemX Is Nothing Then .SetText 6 + itemX.Index, intRow, "V"
'                Set itemX = Nothing
            End If
            If chkQC.Value = 1 Then
                intCol = sl_spcid_tstcd_select_qc&(INS_CODE, strBarno(intIdx), strTestcd)
            Else
                intCol = sl_spcid_tstcd_select(strBarno(intIdx), strTestcd, strTPid, strTPnm)
            End If

            strEqpCd = f_funGet_CODE(strTestcd(intIdx))
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

Private Sub spdResult1_Change(ByVal Col As Long, ByVal Row As Long)
    
    Dim varTmp  As Variant
    Dim intRack As String, intCup   As String
    Dim intRow  As Integer
    
    Select Case Col
        Case 3  '-- Rack
        Case 4  '-- Cup
                spdResult1.GetText 3, Row, varTmp:     intRack = Val(varTmp)
                spdResult1.GetText Col, Row, varTmp:   intCup = Val(varTmp)
                
                For intRow = Row + 1 To spdResult1.maxrows
                    spdResult1.GetText 2, intRow, varTmp
                    If Trim$(varTmp) = "" Then Exit For
                    
                    intCup = intCup + 1
                    If intCup <= 30 Then
                        spdResult1.SetText 3, intRow, CStr(intRack)
                        spdResult1.SetText 4, intRow, intCup
                    End If
                Next
    End Select

End Sub

Private Sub spdResult1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Col < 5 Then Exit Sub
    
    With spdResult1
        .Col = Col: .Row = Row
        .CellType = CellTypeEdit
        .TypeVAlign = True
        .TypeHAlign = True
    End With
    
End Sub


Private Sub Timer1_Timer()
    
    Dim strSendBuff    As String
    
 '########### ALL TYPES OF FINAL RESULTS ARE TRANSFFERD TO THE HOST ######################
    
    strSendBuff = ""
    If f_intReceiveFlag Then
        f_intCnt = f_intCnt + 1
    
        If f_intCnt <= UBound(f_strRack) And f_strTest(1) <> "" Then
            strSendBuff = Chr(1) & Chr(10) + _
                          "06" & " " & "HOSTNAMESIXTEENX" & " " & "09" & Chr(10) + _
                          Chr(2) & Chr(10) + _
                          "10" & " " & "08" & Chr(10) + _
                          "11" & " " & f_strRack(f_intCnt) + "/" + f_strCup(f_intCnt) & Chr(10) + "12" & " " & f_strTest(f_intCnt) + Chr(10) + _
                          Chr(3) & Chr(10) + _
                          Chr(4) & Chr(10)
        Else
            f_intReceiveFlag = False
        End If
   
    Else
        strSendBuff = Chr(1) & Chr(10) + _
                      "06" & " " & "HOSTNAMESIXTEENX" & " " & "09" & Chr(10) + _
                      Chr(2) & Chr(10) + _
                      "10" & " " & "09" & Chr(10) + _
                      Chr(3) & Chr(10) + _
                      Chr(4) & Chr(10)
    End If
   
   Call COM_OUTPUT(strSendBuff)

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
    
    Dim strTest As String, intPos   As Integer
    Dim strTmp  As String, intCol   As Integer, intCnt  As Integer
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = "":  Erase f_typCode
    
    intCol = 7
    With spdWorkList
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 14
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 14
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) AS TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TESTCD_EQP") & ""), , "LST")
        
            strTest = ""
            strTmp = Trim(adoRS.Fields("TESTCD") & "")
            intPos = InStr(strTmp, ",")
            Do While intPos > 0
                strTest = strTest + "[" + Mid$(strTmp, 1, intPos - 1) + "]"
                strTmp = Mid$(strTmp, intPos + 1)
                intPos = InStr(strTmp, ",")
            Loop
            strTest = strTest + "[" + strTmp + "]"
        
            itemX.SubItems(1) = Trim(adoRS.Fields("TESTCD") & "")
            itemX.SubItems(2) = Trim(adoRS.Fields("TESTNM") & "")
            itemX.SubItems(3) = strTest
            itemX.SubItems(4) = Trim(adoRS.Fields("DELTA") & "")
            itemX.SubItems(5) = Trim(adoRS.Fields("DELTAGBN") & "")
            itemX.SubItems(6) = Trim(adoRS.Fields("PANICL") & "")
            itemX.SubItems(7) = Trim(adoRS.Fields("PANICH") & "")
            itemX.SubItems(8) = Trim(adoRS.Fields("REFL") & "")
            itemX.SubItems(9) = Trim(adoRS.Fields("REFH") & "")
            itemX.SubItems(10) = Trim(adoRS.Fields("AUTOVERIFY") & "")
            itemX.SubItems(11) = Trim(adoRS.Fields("REMARK") & "")
            itemX.tag = Trim(adoRS.Fields("TESTCD_EQP") & "")
            itemX.Text = Trim(adoRS.Fields("TESTCD") & "")
        
        Set itemX = Nothing
        
        With spdWorkList
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
            .Col = intCol: .ColHidden = True
        End With
        
        With spdResult1
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        With spdResult2
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
'        strTmp = Trim$(adoRS.Fields("TESTCD"))
'        intPos = InStr(strTmp, ",")
'        Do While intPos > 0
'            f_strOrdList = f_strOrdList + Mid$(strTmp, 1, intPos - 1) + "|"
'            strTmp = Mid$(strTmp, intPos + 1)
'
'            intPos = InStr(strTmp, ",")
'        Loop
'        f_strOrdList =  f_strOrdList + strTmp + "|"
        
        intCnt = intCnt + 1
        ReDim Preserve f_typCode(1 To intCnt) As TYPE_CD
        
        f_typCode(intCnt).strEqpCd = Trim$(adoRS.Fields("TESTCD_EQP"))
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

'    f_strOrdList = "|" +f_strOrdList
    f_strOrdList = Mid$(f_strOrdList, 1, Len(f_strOrdList) - 1)
    
Exit Sub
ErrRoutine:
    Set adoRS = Nothing
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

End Sub

Private Sub cmdClear()
    
    With spdWorkList
        .maxrows = 14
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    With spdResult1
        .maxrows = 14
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
    End With
    
    With spdResult1
        .maxrows = 14
        .Col = 5:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
'        .CellType = CellTypeStaticText
'        .TypeVAlign = True
'        .TypeHAlign = True
        .BlockMode = False
    End With
    
    With spdResult2
        .maxrows = 14
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    f_intReceiveFlag = False
    
End Sub

Private Sub cmdExit()
    
    Unload Me

End Sub

Private Sub cmdRun()
    
    Dim itemX       As ListItem
    Dim strSendBuff As String
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub cmdRun()"
    
    If Not comEQP.PortOpen Then comEQP.PortOpen = True
    If comEQP.PortOpen Then
        Call ShowMessage("연결 되었습니다.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "작업중.."
        
        Timer1.Enabled = True
        
        '-- 통신모드(장비)
        '########### CONNECTION ESTABLISH ######################
        strSendBuff = ""
        strSendBuff = Chr(1) & Chr(10) + _
                      "06" & " " & "HOSTNAMESIXTEENX" & " " & "00" & Chr(10) + _
                      Chr(2) & Chr(10) + _
                      Chr(3) & Chr(10) + _
                      Chr(4) & Chr(10)
    
        Call COM_OUTPUT(strSendBuff)
        
        f_intIdleFlag = 1
        
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
    
    Dim strOrdLst()     As String, strPid() As String, strPnm() As String
    
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
                    End If
                Next
                
                If intCnt > 0 Then
                    If chkQC.Value = 0 Then
                        Call sl_online_result_ul_4&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
                    Else
                        Call sl_online_pc_98&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
                    End If
                    If strErrMsg = "" Then
                        If Index = 1 Then
                            sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
                                     " where SPCNO   = '" & strBarno & "'" & _
                                     "   and TRANSDT = '" & mskRstDate.Text & "'" & _
                                     "   and TRANSTM = '" & strTime & "'"
                            AdoCn_Jet.Execute sqlDoc
                        End If
                    Else
                        MsgBox strErrMsg, vbInformation, Me.Caption
                    End If
                Else
                    MsgBox "검체번호 [" + strBarno + "]를 저장하지 못했습니다.", vbInformation, Me.Caption
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

Private Sub cmdENQ_Click()
    
    Call COM_OUTPUT(charCOM_Convert(COM_ENQ))

End Sub

Private Sub cmdWKSend_Click()

    Dim varTmp      As Variant
    Dim strBuffer   As String
    Dim strBarno    As String, strRack  As String, strCup   As String
    Dim intRow      As Integer, intCol  As Integer

    Dim intCnt      As Integer
    Dim itemX   As ListItem
    
    '-- 02/03/12/13/10/11/14/15
    
    intCnt = 0
    Erase f_strBarno:   Erase f_strRack:    Erase f_strCup: Erase f_strTest

    ReDim Preserve f_strBarno(1 To 1) As String
    ReDim Preserve f_strRack(1 To 1) As String
    ReDim Preserve f_strCup(1 To 1) As String
    ReDim Preserve f_strTest(1 To 1) As String
    
    f_strTest(1) = ""
    
    For intRow = 1 To spdResult1.maxrows
        spdResult1.GetText 2, intRow, varTmp:  strBarno = Trim$(varTmp)
        spdResult1.GetText 3, intRow, varTmp:  strRack = Format$(Trim$(varTmp), "000")
        spdResult1.GetText 4, intRow, varTmp:  strCup = Format$(Trim$(varTmp), "00")
        spdResult1.GetText 1, intRow, varTmp
        
        If Trim$(varTmp) = "1" Then
            If strBarno = "000" Or strCup = "00" Then
                MsgBox IIf(strBarno = "", "Rack", "Cup") + " 번호가 없습니다."
                Exit Sub
            End If
                
            intCnt = intCnt + 1
            ReDim Preserve f_strBarno(1 To intCnt) As String
            ReDim Preserve f_strRack(1 To intCnt) As String
            ReDim Preserve f_strCup(1 To intCnt) As String
            ReDim Preserve f_strTest(1 To intCnt) As String
            
            f_strBarno(intCnt) = strBarno
            f_strRack(intCnt) = strRack:    f_strCup(intCnt) = strCup
            f_strTest(intCnt) = ""
            For intCol = 5 To spdResult1.MaxCols
                spdResult1.Col = intCol:    spdResult1.Row = intRow
                If spdResult1.BackColor = &HC6FEFF Then
                    spdResult1.GetText intCol, 0, varTmp
                    
                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                    If Not itemX Is Nothing Then
                        f_strTest(intCnt) = f_strTest(intCnt) + itemX.tag + " "
                    End If
                    Set itemX = Nothing
                End If
            Next
            f_strTest(intCnt) = Trim$(f_strTest(intCnt))
        End If
        spdResult1.SetText 1, intRow, "0"
    Next
    
    f_intCnt = 1
    strBuffer = Chr(1) & Chr(10) + _
                "06 HOSTNAMESIXTEENX 10" & Chr(10) + _
                Chr(2) & Chr(10) + _
                "50 R " & f_strRack(f_intCnt) & "/" & f_strCup(f_intCnt) & Chr(10) + _
                "51 " & f_strBarno(f_intCnt) + Space(20 - Len(f_strBarno(f_intCnt))) & Chr(10) + _
                "52 " & f_strTest(f_intCnt) + Chr(10) + _
                Chr(3) & Chr(10) + _
                Chr(4) & Chr(10)
                        
    Call COM_OUTPUT(strBuffer)
    f_intOrderFlag = 1
    
End Sub

Private Sub cmdWorkList_Click()

    Dim varTmp      As Variant
    Dim intRow1     As Integer, intRow2 As Integer
    Dim intIdx      As Integer, intCol  As Integer
    
    ReDim strDta(1 To spdWorkList.MaxCols) As String
    
    With spdWorkList
        For intRow1 = 1 To .maxrows
            For intCol = 1 To .MaxCols
                .GetText intCol, intRow1, varTmp:   strDta(intCol) = Trim$(varTmp)
            Next
            
            If strDta(1) = "1" Then
                intRow2 = f_funGet_SpreadRow(spdResult1, 2, Trim$(strDta(2)))
                If intRow2 < 1 Then
                    intRow2 = f_funGet_SpreadRow(spdResult1, 2, "")
                    If intRow2 < 1 Then
                        spdResult1.maxrows = spdResult1.maxrows + 1
                        spdResult1.RowHeight(spdResult1.maxrows) = 13
                        intRow2 = spdResult1.maxrows
                    End If
                    spdResult1.SetText 2, intRow2, strDta(2)
                End If
                
'                spdResult1.SetText 1, intRow2, "1"
'                spdResult1.SetText 2, intRow2, strDta(2)
'                spdResult1.SetText 5, intRow2, strDta(5)
'                spdResult1.SetText 6, intRow2, strDta(6)
                
                If chkQC.Value = 0 Then
                    spdResult1.SetText 1, intRow2, "1"
                    spdResult1.SetText 2, intRow2, strDta(2)
                    spdResult1.SetText 5, intRow2, strDta(5)
                    spdResult1.SetText 6, intRow2, strDta(6)
                Else
                    spdResult1.SetText 1, intRow2, "1"
                    spdResult1.SetText 2, intRow2, strDta(2)
                    spdResult1.SetText 5, intRow2, "QC검사"
                    spdResult1.SetText 6, intRow2, "Level : " & strDta(4)
                End If
                
                For intCol = 7 To UBound(strDta)
                    spdResult1.Col = intCol:  spdResult1.Row = intRow2
                    spdResult1.BackColor = IIf(strDta(intCol) = "V", &HC6FEFF, vbWhite)
                Next
                
                intRow2 = intRow2 + 1
                .Row = intRow1
                .Action = ActionDeleteRow
                .RowHeight(.maxrows) = 13
                If .maxrows > 14 Then .maxrows = .maxrows - 1
                If intRow1 > 0 Then intRow1 = intRow1 - 1
            End If
        Next
    End With

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
            Call ErrMsgProc("", strErrMsg)
        End If
'    Else
'        Call ErrMsgProc("", "검체번호 [" + strBarno + "]를 저장하지 못했습니다.")
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
        
        Call sl_online_pc_98&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
        If strErrMsg = "" Then
            f_funAdd_QcServer = True
        Else
            Call ErrMsgProc("", strErrMsg)
        End If
'    Else
'        Call ErrMsgProc("", "검체번호 [" + strBarno + "]를 저장하지 못했습니다.")
    End If
                                
End Function

Private Sub comEQP_OnComm()
    
    Dim strEVMsg    As String
    Dim strERMsg    As String
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
            Arr = comEQP.Input
            Call ComReceive(Arr)
            
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

Private Sub ComReceive(ByRef RecData() As Byte)
    
    
    Dim strTmp1 As String, strTmp2  As String
    Dim strRec  As String, strBuff  As String
    
    Dim varTmp      As Variant
    Dim intIdx1     As Integer, intIdx2 As Integer, intCol      As Integer
    Dim strOrder    As String, strSpec  As String, strPcFlag    As String
    Dim strSampleno As String, intPos   As Integer, intOrdCnt   As Integer
    Dim strSector   As String, strCup   As String
    
    Static OrgMsg As String
    strRec = StrConv(RecData, vbUnicode)
    
    Print #1, strRec;
    
    Call COM_INPUT(strRec)
    
    For intIdx1 = 1 To Len(strRec)
        strBuff = Mid$(strRec, intIdx1, 1)
        Select Case Asc(strBuff)
            Case "1"    '-- SOH
                        f_strBuffer = ""
            Case "4"    '--  EOT
                        Call f_subSet_Result(f_strBuffer)
            Case Else
                        f_strBuffer = f_strBuffer + strBuff
        End Select
     Next

End Sub



Public Function f_funGet_CheckSum(ByVal strPara As String) As String

    Dim intIdx      As Integer
    Dim intChkSum   As Integer
    
    intChkSum = 0
    For intIdx = 1 To Len(strPara)
        intChkSum = intChkSum + (0 Xor Asc(Mid$(strPara, intIdx, 1)))
    Next
    
    f_funGet_CheckSum = Chr(intChkSum) '-Format$(Hex(intChkSum), "00")
        
End Function

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
    Open App.Path + "\" + "dump_job.log" For Output As #1
    
    cboRstgbn(0).ListIndex = 0: cboRstgbn(1).ListIndex = 2
    f_intReceiveFlag = False
    tabWork.Tab = 0
    
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
                .InputMode = Trim(mAdoRs.Fields("COM_INPUTMOD") & "")
                .DTREnable = Trim(mAdoRs.Fields("COM_DTR") & "")
                .EOFEnable = Trim(mAdoRs.Fields("COM_EOF") & "")
                .NullDiscard = Trim(mAdoRs.Fields("COM_NULDIS") & "")
                .RTSEnable = Trim(mAdoRs.Fields("COM_RTS") & "")
                .InBufferSize = Trim(mAdoRs.Fields("COM_IBS") & "")
                .InputLen = Trim(mAdoRs.Fields("COM_INLEN") & "")
                .OutBufferSize = Trim(mAdoRs.Fields("COM_OBS") & "")
                .ParityReplace = Trim(mAdoRs.Fields("COM_PTR") & "")
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

Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)

    If Col < 5 Then Exit Sub
    
    Dim varTmp  As Variant
    
    With spdResult1
        .GetText Col, 0, varTmp
        If Trim$(varTmp) = "" Then Exit Sub
        
        .Row = Row: .Col = Col
        If .BackColor = vbWhite Then
            .BackColor = &HC6FEFF
        Else
            .BackColor = vbWhite
        End If
    End With

End Sub

Private Sub spdWorkList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    Dim varTmp      As Variant
    Dim intRow      As Integer
    Dim intStartRow As Integer, intEndRow   As Integer
    
    If BlockRow > BlockRow2 Then
        intStartRow = BlockRow2
        intEndRow = BlockRow
    Else
        intStartRow = BlockRow
        intEndRow = BlockRow2
    End If
    
    For intRow = intStartRow To intEndRow
        
        spdWorkList.GetText 2, intRow, varTmp
        If Trim$(varTmp) <> "" Then
            spdWorkList.GetText 1, intRow, varTmp
            spdWorkList.SetText 1, intRow, IIf(Trim$(varTmp) = "1", "", "1")
        End If
    Next

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
                        spdWorkList.SetText 5, intRow, strPnm(0)
                        spdWorkList.SetText 6, intRow, strPid(0)
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



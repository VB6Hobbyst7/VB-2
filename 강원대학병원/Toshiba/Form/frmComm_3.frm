VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   11985
   WindowState     =   2  '최대화
   Begin VB.Timer tmrReceive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3420
      Top             =   6630
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3450
      Top             =   6690
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   2250
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
            Picture         =   "frmComm_3.frx":0000
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_3.frx":059A
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_3.frx":0B34
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_3.frx":10CE
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_3.frx":1668
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_3.frx":1C02
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm comEQP 
      Left            =   2820
      Top             =   6510
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   2
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin MSComctlLib.ImageList imlStatus 
      Left            =   2730
      Top             =   6540
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
            Picture         =   "frmComm_3.frx":219C
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_3.frx":2736
            Key             =   "NOT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_3.frx":2CD0
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_3.frx":326A
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_3.frx":3AFC
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_3.frx":3C56
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_3.frx":3DB0
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
      Begin VB.Timer Timer3 
         Left            =   1200
         Top             =   120
      End
      Begin VB.Timer Timer2 
         Left            =   570
         Top             =   90
      End
      Begin VB.Timer Timer1 
         Left            =   90
         Top             =   90
      End
      Begin HSCotrol.CButton cmdInit 
         Height          =   360
         Left            =   3930
         TabIndex        =   55
         Top             =   135
         Width           =   1170
         _ExtentX        =   2064
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
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdOrder 
         Height          =   360
         Left            =   5160
         TabIndex        =   52
         Top             =   135
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
      Picture         =   "frmComm_3.frx":3F0A
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
         Picture         =   "frmComm_3.frx":518C
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   9780
         Picture         =   "frmComm_3.frx":5716
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   8640
         Picture         =   "frmComm_3.frx":5CA0
         Top             =   255
         Width           =   240
      End
   End
   Begin TabDlg.SSTab tabWork 
      Height          =   5850
      Left            =   45
      TabIndex        =   13
      Top             =   585
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   10319
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " WorkList"
      TabPicture(0)   =   "frmComm_3.frx":622A
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
      Tab(0).Control(6)=   "spdResult1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdWordQuery"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdAppend(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "mskOrdDate"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboRstgbn(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtBarCode"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdSel(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "optBar"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "optSeq"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdRackNo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdStartNo"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cboTest"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "chkMan"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdSel(5)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdSel(4)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "chkAuto"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   " 받은 결과"
      TabPicture(1)   =   "frmComm_3.frx":6246
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
         Left            =   6885
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   540
         Value           =   1  '확인
         Width           =   1410
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   4
         Left            =   2820
         TabIndex        =   58
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_3.frx":6262
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   5
         Left            =   2550
         TabIndex        =   59
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_3.frx":66E4
      End
      Begin VB.CheckBox chkMan 
         Caption         =   "Manual"
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
         Left            =   10710
         TabIndex        =   57
         Top             =   30
         Width           =   1065
      End
      Begin VB.ComboBox cboTest 
         Height          =   300
         ItemData        =   "frmComm_3.frx":6B52
         Left            =   5460
         List            =   "frmComm_3.frx":6B5F
         Style           =   2  '드롭다운 목록
         TabIndex        =   56
         Top             =   480
         Visible         =   0   'False
         Width           =   1485
      End
      Begin HSCotrol.CButton cmdStartNo 
         Height          =   300
         Left            =   7080
         TabIndex        =   54
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
      Begin HSCotrol.CButton cmdRackNo 
         Height          =   300
         Left            =   8310
         TabIndex        =   53
         Top             =   495
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "Rack번호변경"
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
         Left            =   9000
         TabIndex        =   51
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
         Left            =   9930
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   4920
         Left            =   -69105
         TabIndex        =   47
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
         TabIndex        =   46
         Top             =   315
         Visible         =   0   'False
         Width           =   420
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   1
         Left            =   360
         TabIndex        =   27
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_3.frx":6B7F
      End
      Begin VB.ComboBox cboRstgbn 
         Height          =   300
         Index           =   1
         ItemData        =   "frmComm_3.frx":7001
         Left            =   -72570
         List            =   "frmComm_3.frx":700E
         Style           =   2  '드롭다운 목록
         TabIndex        =   17
         Top             =   495
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.TextBox txtBarCode 
         Height          =   300
         Left            =   2415
         MaxLength       =   12
         TabIndex        =   1
         Top             =   480
         Width           =   1485
      End
      Begin VB.ComboBox cboRstgbn 
         Height          =   300
         Index           =   0
         ItemData        =   "frmComm_3.frx":7038
         Left            =   3900
         List            =   "frmComm_3.frx":7045
         Style           =   2  '드롭다운 목록
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   1485
      End
      Begin MSMask.MaskEdBox mskRstDate 
         Height          =   300
         Left            =   -73695
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
      Begin HSCotrol.CButton cmdRstQuery 
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
         Left            =   1305
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   480
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
         Left            =   9555
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
      Begin FPSpread.vaSpread spdResult1 
         Height          =   4830
         Left            =   2565
         TabIndex        =   25
         Top             =   900
         Width           =   9090
         _Version        =   196608
         _ExtentX        =   16034
         _ExtentY        =   8520
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   4
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
         MaxCols         =   14
         MaxRows         =   14
         ScrollBarMaxAlign=   0   'False
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmComm_3.frx":706F
         UserResize      =   1
      End
      Begin HSCotrol.CButton cmdWorkList 
         Height          =   300
         Left            =   90
         TabIndex        =   26
         Top             =   5445
         Width           =   2460
         _ExtentX        =   4339
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
         TabIndex        =   28
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_3.frx":76B0
      End
      Begin FPSpread.vaSpread spdWorkList 
         Height          =   4515
         Left            =   90
         TabIndex        =   16
         Top             =   900
         Width           =   2475
         _Version        =   196608
         _ExtentX        =   4366
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
         MaxRows         =   14
         ScrollBarMaxAlign=   0   'False
         SpreadDesigner  =   "frmComm_3.frx":7B1E
         UserResize      =   0
      End
      Begin HSCotrol.UserPanel pnlCom 
         Height          =   5355
         Left            =   45
         TabIndex        =   29
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
            TabIndex        =   35
            Top             =   270
            Width           =   11595
         End
         Begin VB.Frame Frame1 
            Height          =   645
            Left            =   45
            TabIndex        =   30
            Top             =   4650
            Width           =   11610
            Begin HSCotrol.CButton cmdCOMSave 
               Height          =   360
               Left            =   10515
               TabIndex        =   31
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
               TabIndex        =   32
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
               TabIndex        =   33
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
               TabIndex        =   34
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
         TabIndex        =   48
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_3.frx":7FDD
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   -74910
         TabIndex        =   49
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_3.frx":845F
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   15
         MaxRows         =   14
         ScrollBarMaxAlign=   0   'False
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmComm_3.frx":88CD
      End
      Begin HSCotrol.UserPanel pnlCom2 
         Height          =   3975
         Left            =   5880
         TabIndex        =   36
         Top             =   1815
         Visible         =   0   'False
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   7011
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
            TabIndex        =   45
            Top             =   270
            Width           =   5730
         End
         Begin VB.Frame Frame2 
            Height          =   645
            Left            =   90
            TabIndex        =   37
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
               TabIndex        =   38
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
               TabIndex        =   39
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
               TabIndex        =   40
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
               TabIndex        =   41
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
               TabIndex        =   42
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
               TabIndex        =   43
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
               TabIndex        =   44
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

Private f_intCnt        As Integer
Private f_strBarno()    As String, f_strTest()  As String
Private f_strRack()     As String, f_strCup()   As String

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

Dim fRcvString As String
Dim fChannel() As String
Dim fTBA200(100) As String
Dim fTBA200Cfg(100) As Integer
Dim Flag_HQL As String
Dim Patiant_Recevid As Boolean

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""

Dim PatientID As String    'Q Message Pattern Check
Dim PatientSeq As String
Dim PatientDisk As String
Dim PatientRack As String
Dim PatientPos As String
Dim SendCount As Integer

Dim OrderCnt As Integer
Dim TbaStat As Boolean
Dim TTT As Integer

Private Type ITEMCODE
    MachCode As String
    MachCnt As Integer
    MachTst(1 To 99) As String
End Type

Private Channel() As ITEMCODE
Private t_no As String

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
        For intCol = 7 To .MaxCols
            .Row = intRow:  .Col = intCol
            If .BackColor = &HC6FEFF Then
                Select Case itemX.Item(intCol - 6).SubItems(11)
                    Case "128": strSpec = "PL"
                    Case Else:  strSpec = "SE"
                End Select
                .GetText intCol, 0, varTmp
                
                If itemX.Item(intCol - 6).tag = "XXX" Then
                    strOrder = strOrder + "06A ," + itemX.Item(intCol - 6).SubItems(10) + ",": strPcFlag = "PC"
                Else
                    strOrder = strOrder + itemX.Item(intCol - 6).tag + " ," & itemX.Item(intCol - 6).SubItems(10) + ","
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
            Call .Add(, "ETC", "비고", (lvwCuData.Width - 310) * 0.1)
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
             " order by OUT_SEQ, TESTCD"
             
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
    
    Dim intCol  As Integer, intPos  As Integer
    Dim strTmp  As String, strTest  As String
    Dim intCnt  As Integer
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = ""
    
    intCol = 7
    With spdWorkList
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 14
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .Col = 5:   .ColHidden = True
        .Col = 6:   .ColHidden = True
    End With
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 15
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
    End With
    
    With spdResult2
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 15
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))"
             
    sqlDoc = sqlDoc + " order by OUT_SEQ, TESTCD "
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst: ReDim fChannel(adoRS.RecordCount)
    Do While Not adoRS.EOF
        Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TEST_EQP") & ""), , "LST")
            
            intCnt = intCnt + 1
            ReDim Preserve Channel(1 To intCnt) As ITEMCODE
            
            Channel(intCnt).MachCode = Trim$(adoRS.Fields("TEST_EQP"))
            Channel(intCnt).MachCnt = 0
            
            strTmp = Trim$(adoRS.Fields("TESTCD"))
            intPos = InStr(strTmp, ",")
            Do While intPos > 0
                strTest = strTest + "'" + Mid$(strTmp, 1, intPos - 1) + "',"
                
                Channel(intCnt).MachCnt = Channel(intCnt).MachCnt + 1
                Channel(intCnt).MachTst(Channel(intCnt).MachCnt) = Mid$(strTmp, 1, intPos - 1)
                strTmp = Mid$(strTmp, intPos + 1)
                
                intPos = InStr(strTmp, ",")
            Loop
            strTest = strTest + "'" + strTmp + "',"
            Channel(intCnt).MachCnt = Channel(intCnt).MachCnt + 1
            Channel(intCnt).MachTst(Channel(intCnt).MachCnt) = strTmp
            
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
            itemX.SubItems(12) = strTest
            itemX.tag = Trim(adoRS.Fields("TEST_EQP") & "")
            'itemX.Text = Trim(adoRS.Fields("TESTCD") & "")
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
        
        intCol = intCol + 1
        f_strOrdList = f_strOrdList + ", '" & Trim$(adoRS.Fields("TESTCD")) & "'"
        
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
    With spdResult2
        If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
        .SetText intCol, 0, ""
        .Col = intCol:  .ColHidden = True
    End With
    
    f_strOrdList = Mid$(f_strOrdList, 3)
    
Exit Sub
ErrRoutine:
    Set adoRS = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub


Private Sub cboTest_Click()
    
    If Timer1.Enabled = False Then
        spdResult1.MaxCols = 6
        Call f_subSet_ItemList      ' 검사항목
    End If
    
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
    Dim strSampleno()   As String, strBarno     As String
    Dim strOrdcd()      As String, strRstval()  As String, intCnt       As Integer
    Dim strTmp1()       As String, strTmp2()    As String
    Dim intPos          As String, strTestcd    As String, strTestRst   As String
    
    Dim strOrdLst()     As String
    
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
            .GetText 2, intRow, varTmp: strBarno = Trim$(varTmp)
            .GetText 1, intRow, varTmp
            
            If strBarno = "" Then Exit For
            Call sl_spcid_tstcd_select&(strBarno, strOrdLst)
            
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
                    Call sl_online_result_ul_4&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
                    If strErrMsg <> "" Then MsgBox strErrMsg, vbInformation, Me.Caption
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
   
'    Dim adoRS   As New ADODB.Recordset
'    Dim sqlDoc  As String
'
'    Dim varTmp  As Variant, strErrMsg   As String
'    Dim strSampleno()   As String, strBarno     As String
'    Dim strOrdcd()      As String, strRstVal()  As String, intCnt       As Integer
'    Dim strTmp1()       As String, strTmp2()    As String
'    Dim intPos          As String, strTestCd    As String, strTestRst   As String
'
'    Dim strOrdLst()     As String
'
'    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
'    Dim itemX   As ListItem
'    Dim objSpd  As vaSpread
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
'            .GetText 2, intRow, varTmp: strBarno = Trim$(varTmp)
'            .GetText 1, intRow, varTmp
'
'            If strBarno = "" Then Exit For
'            Call sl_spcid_tstcd_select&(strBarno, strOrdLst)
'
'            intCnt = 0: Erase strOrdcd: Erase strRstVal
'            If Trim$(varTmp) = "1" Then
'                For intCol = 5 To .MaxCols
'                    .GetText intCol, intRow, varTmp
'                    If Trim$(varTmp) <> "" Then
'                        .GetText intCol, 0, varTmp
'                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
'                        If Not itemX Is Nothing Then
'                            .GetText intCol, intRow, varTmp
'                            strTestCd = itemX.ListSubItems(1)
'                            intPos = InStr(strTestCd, ",")
'                            Do While intPos > 0
'
'                                blnFlag = False
'                                For intIdx = 0 To UBound(strOrdLst)
'                                    If strOrdLst(intIdx) = Mid$(strTestCd, 1, intPos - 1) Then blnFlag = True:  Exit For
'                                Next
'
'                                If blnFlag Then
'                                    intCnt = intCnt + 1
'                                    ReDim Preserve strSampleno(1 To intCnt) As String
'                                    ReDim Preserve strOrdcd(1 To intCnt) As String
'                                    ReDim Preserve strRstVal(1 To intCnt) As String
'                                    ReDim Preserve strTmp1(1 To intCnt) As String
'                                    ReDim Preserve strTmp2(1 To intCnt) As String
'
'                                    strSampleno(intCnt) = strBarno
'                                    strOrdcd(intCnt) = Mid$(strTestCd, 1, intPos - 1)
'                                    strRstVal(intCnt) = Trim$(varTmp)
'                                End If
'
'                                strTestCd = Mid$(strTestCd, intPos + 1)
'                                intPos = InStr(strTestCd, ",")
'                            Loop
'
'                            blnFlag = False
'                            For intIdx = 0 To UBound(strOrdLst)
'                                If strOrdLst(intIdx) = strTestCd Then blnFlag = True: Exit For
'                            Next
'
'                            If blnFlag Then
'                                intCnt = intCnt + 1
'                                ReDim Preserve strSampleno(1 To intCnt) As String
'                                ReDim Preserve strOrdcd(1 To intCnt) As String
'                                ReDim Preserve strRstVal(1 To intCnt) As String
'                                ReDim Preserve strTmp1(1 To intCnt) As String
'                                ReDim Preserve strTmp2(1 To intCnt) As String
'
'                                strSampleno(intCnt) = strBarno
'                                strOrdcd(intCnt) = strTestCd
'                                strRstVal(intCnt) = Trim$(varTmp)
'                            End If
'                        End If
'                        Set itemX = Nothing
'                    End If
'                Next
'
'                If intCnt > 0 Then
'                    Call sl_online_result_ul_4&(strErrMsg, strSampleno, strOrdcd, strRstVal, strTmp1, strTmp2, Chr(0))
'                    If strErrMsg <> "" Then MsgBox strErrMsg, vbInformation, Me.Caption
'                Else
'                    MsgBox "검체번호 [" + strBarno + "]를 저장하지 못했습니다.", vbInformation, Me.Caption
'                End If
'            End If
'        Next
'    End With
'    Me.MousePointer = 0
'    MsgBox "작업이 완료되었습니다.", vbInformation, Me.Caption
'
'    Exit Sub
'ErrorRoutine:
'    Set itemX = Nothing
'
'    Me.MousePointer = 0
'    Call ErrMsgProc(CallForm)
'
''    Dim adoRS   As New ADODB.Recordset
''    Dim sqlDoc  As String
''
''    Dim varTmp  As Variant, strErrMsg   As String
''    Dim strSampleno()   As String
''    Dim strOrdcd()      As String, strRstval()  As String, intCnt       As Integer
''    Dim strTmp1()       As String, strTmp2()    As String, strTmp3()    As String
''    Dim intRow  As Integer, intCol  As Integer
''    Dim itemX   As ListItem
''    Dim objSPD  As vaSpread
''    Dim Test_Cd() As String
''    Dim Rev As Long
''    Dim ii As Integer
''    Dim strEqpCd As String
''
''    CallForm = "frmComm - Private Sub cmdAppend_Click()"
''
''On Error GoTo ErrorRoutine
''
''    Me.MousePointer = 11
''
''    If Index = 0 Then
''        Set objSPD = spdResult1
''    Else
''        Set objSPD = spdResult2
''    End If
''
''    With objSPD
''        For intRow = 1 To .maxrows
''            .GetText 1, intRow, varTmp
''
''            intCnt = 0: Erase strOrdcd: Erase strRstval
''            If Trim$(varTmp) = "1" Then
''                For intCol = 5 To .MaxCols
''                    .GetText intCol, 0, varTmp
''                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
''                    If Not itemX Is Nothing Then
''
'''                        intCnt = intCnt + 1
'''                        ReDim Preserve strSampleno(1 To intCnt) As String
'''                        ReDim Preserve strOrdcd(1 To intCnt) As String
'''                        ReDim Preserve strRstval(1 To intCnt) As String
'''                        ReDim Preserve strTmp1(1 To intCnt) As String
'''                        ReDim Preserve strTmp2(1 To intCnt) As String
'''                        ReDim Preserve strTmp3(1 To intCnt) As String
'''
'''                        .GetText 2, intRow, varTmp: strSampleno(intCnt) = Trim$(varTmp)
'''                        strOrdcd(intCnt) = itemX.ListSubItems(1)
'''                        .GetText intCol, intRow, varTmp:    strRstval(intCnt) = Trim$(varTmp)
''
''                        .GetText intCol, intRow, varTmp
''                        If Len(Trim(varTmp)) > 0 Then
''                            intCnt = intCnt + 1
''                            ReDim Preserve strSampleno(1 To intCnt) As String
''                            ReDim Preserve strOrdcd(1 To intCnt) As String
''                            ReDim Preserve strRstval(1 To intCnt) As String
''                            ReDim Preserve strTmp1(1 To intCnt) As String
''                            ReDim Preserve strTmp2(1 To intCnt) As String
''                            ReDim Preserve strTmp3(1 To intCnt) As String
''
''                            'strSampleno(intCnt) = Trim$(varTmp)
''                            .GetText 2, intRow, varTmp
''                            Rev = sl_spcid_tstcd_select(Trim(varTmp), Test_Cd)
''                            For ii = intCnt To Rev - 1
''                                strEqpCd = f_funget_code(Test_Cd(ii))
''                                Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
''                                If Not itemX Is Nothing Then
''                                    strSampleno(intCnt) = Trim$(varTmp)
''                                    strOrdcd(intCnt) = Test_Cd(ii)
''                                    .GetText intCol, intRow, varTmp:    strRstval(intCnt) = Trim$(varTmp)
''                                    'intCnt = intCnt + 1
''                                    Exit For
''                                End If
''                            Next
''
''                            'strOrdcd(intCnt) = itemX.ListSubItems(1)
'''                            .GetText intCol, intRow, varTmp:    strRstval(intCnt) = Trim$(varTmp)
''                        End If
''                    End If
''                    Set itemX = Nothing
''                Next
''
''                Call sl_online_result_ul_4&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, "")
''                If strErrMsg <> "" Then MsgBox strErrMsg, vbInformation, Me.Caption
''            End If
''        Next
''    End With
''    Me.MousePointer = 0
''    MsgBox "작업이 완료되었습니다.", vbInformation, Me.Caption
''
''    Exit Sub
''ErrorRoutine:
''    Set itemX = Nothing
''
''    Me.MousePointer = 0
''    Call ErrMsgProc(CallForm)

End Sub

Private Sub cmdENQ_Click()
    
    Call COM_OUTPUT(charCOM_Convert(COM_ENQ))

End Sub

Private Sub cmdInit_Click()
    
    comEQP.Output = STX + "I " + ETX
    Debug.Print "[HOST] " + STX + "I " + ETX
    fRcvString = ""
    SendCount = 0
    f_intCnt = 0

    TbaStat = False
    OrderCnt = 0
    TTT = 0
    If Timer1.Enabled = True Then Timer1.Enabled = False

End Sub

Private Sub cmdOrder_Click()

    Dim varTmp      As Variant
    Dim intRow      As Integer, intCol  As Integer
    Dim strBarno    As String, strTest  As String
    Dim strRack     As String, strCup   As String
    Dim intCnt      As Integer
    
    Dim itemX       As ListItem
    
    intCnt = 0
    Erase f_strBarno:   Erase f_strTest:    Erase f_strRack:    Erase f_strCup
    ReDim Preserve f_strBarno(1 To 1) As String
    ReDim Preserve f_strRack(1 To 1) As String
    ReDim Preserve f_strCup(1 To 1) As String
    ReDim Preserve f_strTest(1 To 1) As String
    
    f_strBarno(1) = "":  f_strRack(1) = "":    f_strCup(1) = ""
    f_strTest(1) = ""
    
    If f_intCnt <> 0 And f_strBarno(1) <> "" Then
        MsgBox "Order 전송 중입니다.  잠시후에 다시 실행하세요"
        Exit Sub
    End If
    
    With spdResult1
        For intRow = 1 To .maxrows
            .GetText 2, intRow, varTmp: strBarno = Trim$(varTmp)
            .GetText 3, intRow, varTmp: strRack = Format$(varTmp, "00")
            .GetText 4, intRow, varTmp: strCup = Format$(varTmp, "00")
            .GetText 1, intRow, varTmp
            
            If Trim$(varTmp) = "1" Then
                intCnt = intCnt + 1
                ReDim Preserve f_strBarno(1 To intCnt) As String
                ReDim Preserve f_strRack(1 To intCnt) As String
                ReDim Preserve f_strCup(1 To intCnt) As String
                ReDim Preserve f_strTest(1 To intCnt) As String
                
                f_strBarno(intCnt) = strBarno:  f_strRack(intCnt) = strRack:    f_strCup(intCnt) = strCup
                f_strTest(intCnt) = ""
                For intCol = 7 To .MaxCols
                    .GetText intCol, 0, varTmp
                    If Trim$(varTmp) = "" Then Exit For
                    
                    .Row = intRow:  .Col = intCol
                    If .BackColor = &HC6FEFF Then
                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                        If Not itemX Is Nothing Then
                            f_strTest(intCnt) = f_strTest(intCnt) + Space$(4 - Len(itemX.tag)) + itemX.tag + "1"
                            
                            If itemX.tag = "50" Then _
                               f_strTest(intCnt) = f_strTest(intCnt) + "  291  301"
                        End If
                        Set itemX = Nothing
                    End If
                Next
                TbaStat = True
            End If
        Next
    End With
    
    If TbaStat Then
       
        comEQP.Output = STX + "M     " + ETX
        Debug.Print "[HOST] " & STX + "M     " + ETX
        SendCount = SendCount + 1
        OrderCnt = 1
        lblStatus = "오더전송중.."
        f_intCnt = 0
    Else
        MsgBox "장비측과의 통신상태를 확인하세요", vbCritical, Me.Caption
        OrderCnt = 1
        SendCount = 0
        fRcvString = ""
        Call cmdInit_Click
    End If

'Dim ii As Integer
'Dim jj As Integer
'
'    If chkMan.Value = 1 Then
'        spdResult1.Col = 2
'        For jj = 1 To spdResult1.maxrows
'            spdResult1.Row = jj
'            If Trim(spdResult1.Text) = "" Then
'                spdResult1.maxrows = jj - 1
'                Exit For
'            End If
'        Next jj
'    End If
'
'    If TbaStat = True Then
'        With spdResult1
'            'OrderCnt = 1
'            For ii = OrderCnt To .maxrows
'                .Col = 1: .Row = ii
'                If .Value = "1" Then
'                    .Col = 2
''                    If Len(Trim(.Text)) > 0 And SendCount = 1 Then
'                    If Len(Trim(.Text)) > 0 Then
'                        comEQP.Output = STX + "M     " + ETX
'                        Debug.Print "[HOST] " & STX + "M     " + ETX
'                        SendCount = SendCount + 1
'                        OrderCnt = ii
'                        lblStatus.Caption = "오더전송중.."
'                        'Me.Enabled = False
'                        Exit For
'                    End If
'                End If
'            Next ii
'        End With
'    Else
'        MsgBox "장비측과의 통신상태를 확인하세요", vbCritical, Me.Caption
'        OrderCnt = 0
'        SendCount = 0
'        fRcvString = ""
'        Call cmdInit_Click
'    End If
    
End Sub

Private Sub cmdRackNo_Click()

Dim sNo As String, sCnt As Integer, sAdd As Integer
Dim fNum1 As Integer, fNum2 As Integer
Dim intRow1 As Integer

AgainInput:
    fNum1 = 1: fNum2 = 0
    If t_no = "" Then
        sNo = InputBox("시작 번호를 입력하세요 !")
    Else
        sNo = t_no
    End If
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
        If Not IsNumeric(sNo) Then
            MsgBox "숫자만 입력하세요.!", vbCritical
            GoTo AgainInput
        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .maxrows
                intRow1 = intRow1 + 1
                .Row = sCnt
                .Col = 1
                If .Value >= 1 Then
                    'If .ActiveCol = 3 Then
                        .Col = 3 '.ActiveCol
                        If intRow1 = (5 * fNum1) + 1 Then fNum1 = fNum1 + 1: fNum2 = 0
                        fNum2 = fNum2 + 1
                        .Text = Format(Trim((fNum1 + Val(sNo)) - 1), "00")
                        .Col = 4 '.ActiveCol + 1
                        .Text = fNum2
                    'End If
                End If
            Next sCnt
        End With
'        spdResult1.StartingRowNumber = Val(sNo)
    End If
End Sub

Private Sub cmdRstQuery_Click()

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strSpcno    As String
    Dim intRow      As Integer, intCol  As Integer
    
    Dim itemX       As ListItem
    Dim jj As Integer
    
    intRow = 0
    With spdResult2
        .maxrows = 14
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    sqlDoc = "select SPCNO, TESTCD, EQPNUM, TRANSTM, RSTVAL, REFVAL, TRANSDT " & _
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
            If strSpcno <> Trim$(adoRS(0) & "") + Trim$(adoRS(6) & "") Then
                intRow = intRow + 1
                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                
                .SetText 1, intRow, "1"
                .SetText 2, intRow, Trim$(adoRS(0) & "")
            End If
            strSpcno = Trim$(adoRS(0) & "") + Trim$(adoRS(6) & "")
            Set itemX = lvwCuData.FindItem(Trim$(adoRS(1) & ""), lvwSubItem, , lvwWhole)
            
            For jj = 1 To spdResult1.MaxCols - 6
                If Len(Channel(jj).MachCode) > 0 Then
                    'spdResult1.Col = jj + 4
                    'If spdResult1.BackColor = &HC6FEFF Then
                        'Testoutput = Testoutput + Space$(4 - Len(Channel(jj).MachCode)) + Channel(jj).MachCode + "1"
                        'If Channel(jj).MachCode = "50" Then
                        '    Testoutput = Testoutput + "  291  301"
                        'End If
                    'End If
                    If Channel(jj).MachCode = adoRS(2) Then
                        intCol = jj + 6
                        .SetText intCol, intRow, Trim$(adoRS(4)) & ""
                        .Col = intCol:  .Row = intRow:  .ForeColor = IIf(Trim$(adoRS(5) & "") <> "", vbRed, vbBlack)
                        Exit For
                    End If
                End If
            Next jj
            
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
    ElseIf Index = 4 Or Index = 5 Then
        With spdResult1
            For intRow = 1 To .maxrows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 5, "1", "")
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

Private Sub cmdWorkList_Click()

    Dim varTmp  As Variant
    Dim intRow1 As Integer, intRow2 As Integer
    Dim intIdx  As Integer
    Dim intRack   As Integer, intCup  As Integer
    Dim Rev     As Long
    Dim Test_Cd() As String
    Dim ii As Integer
    Dim itemX As ListItem
'    Dim adoRS As ADODB.Recordset
'    Dim sqlDoc As String
    Dim bgetWork As Boolean
'    Dim iCol As Integer, iRow As Integer
    Dim strEqpCd As String
    
    bgetWork = False
    intRack = 1: intCup = 0
    With spdWorkList
        For intRow1 = 1 To .maxrows
            .GetText 1, intRow1, varTmp
            If Trim$(varTmp) = "1" Then
                .GetText 2, intRow1, varTmp
                intRow2 = f_funGet_SpreadRow(spdResult1, 2, Trim$(varTmp))
                If intRow2 < 1 Then
                    intRow2 = f_funGet_SpreadRow(spdResult1, 2, "")
                    If intRow2 < 1 Then
                        spdResult1.maxrows = spdResult1.maxrows + 1
                        spdResult1.RowHeight(spdResult1.maxrows) = 12
                        intRow2 = spdResult1.maxrows
                    End If
                    
                    Rev = sl_spcid_tstcd_select(Trim(varTmp), Test_Cd)
                    For ii = 0 To Rev - 1
                        strEqpCd = f_funget_code(Test_Cd(ii))
                        Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                        If Not itemX Is Nothing Then
                            bgetWork = True
                            spdResult1.Row = intRow2
                            spdResult1.Col = itemX.Index + 6
                            spdResult1.BackColor = &HC6FEFF
                            'spdResult1.CellBorderColor = vbRed
                            DoEvents
                        End If
                    Next
                    
                    If bgetWork = True Then
                        spdResult1.SetText 2, intRow2, Trim(varTmp)
                        If intRow2 = 1 Then
                            If intRow1 = (5 * intRack) + 1 Then intRack = intRack + 1: intCup = 0
                            intCup = intCup + 1
                        Else
                            spdResult1.GetText 3, intRow2 - 1, varTmp:  intRack = Format(varTmp, "#0")
                            spdResult1.GetText 4, intRow2 - 1, varTmp:  intCup = Val(varTmp)
                            
                            intCup = intCup + 1
                            If intCup > 5 Then intRack = intRack + 1:   intCup = 1
                        End If
                        spdResult1.SetText 1, intRow2, "1"
                        spdResult1.SetText 3, intRow2, Format(Val(intRack), "00")
                        spdResult1.SetText 4, intRow2, Trim(intCup)
                    Else
                        spdResult1.maxrows = spdResult1.maxrows - 1
                    End If
                    bgetWork = False

                End If
                'spdResult1.SetText 1, intRow2, "1"
                spdResult1.maxrows = intRow2
                
                .SetText 1, intRow1, ""
'                .Row = intRow1
'                .Action = ActionDeleteRow
'                If .maxrows > 14 Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                
                If intRow1 > 0 Then intRow1 = intRow1 - 1
                
            End If
        Next
    End With


'    Dim varTmp  As Variant
'    Dim intRow1 As Integer, intRow2 As Integer
'    Dim intIdx  As Integer
'    Dim fNum1   As Integer, fNum2  As Integer
'    Dim Rev     As Long
'    Dim Test_Cd() As String
'    Dim ii As Integer
'    Dim itemX As ListItem
''    Dim adoRS As ADODB.Recordset
''    Dim sqlDoc As String
'    Dim bgetWork As Boolean
''    Dim iCol As Integer, iRow As Integer
'    Dim strEqpCd As String
'
''    Call cmdInit_Click
''
''    comEQP.Output = STX + "I " + ETX
''    Debug.Print "[HOST] " + STX + "I " + ETX
''    fRcvString = ""
''    SendCount = 0
''    TbaStat = False
''    OrderCnt = 0
''    TTT = 0
''    If Timer1.Enabled = True Then Timer1.Enabled = False
'
'    bgetWork = False
'    fNum1 = 1: fNum2 = 0
'    With spdWorkList
'        For intRow1 = 1 To .maxrows
'            .GetText 1, intRow1, varTmp
'            If Trim$(varTmp) = "1" Then
'                .GetText 2, intRow1, varTmp
'                intRow2 = f_funGet_SpreadRow(spdResult1, 2, Trim$(varTmp))
'                If intRow2 < 1 Then
'                    intRow2 = f_funGet_SpreadRow(spdResult1, 2, "")
'                    If intRow2 < 1 Then
'                        spdResult1.maxrows = spdResult1.maxrows + 1
'                        spdResult1.RowHeight(spdResult1.maxrows) = 12
'                        intRow2 = spdResult1.maxrows
'                    End If
'
'                    If chkMan.Value = 0 Then
'                        Rev = sl_spcid_tstcd_select(Trim(varTmp), Test_Cd)
'                        For ii = 0 To Rev - 1
'                            strEqpCd = f_funget_code(Test_Cd(ii))
'                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
'                            If Not itemX Is Nothing Then
'                                bgetWork = True
'                                spdResult1.Row = intRow2
'                                spdResult1.Col = itemX.Index + 4
'                                spdResult1.BackColor = &HC6FEFF
'                                'spdResult1.CellBorderColor = vbRed
'                                DoEvents
'                            End If
'                        Next
'                    End If
'
'                    If chkMan.Value = 0 Then
'                        If bgetWork = True Then
'                            spdResult1.SetText 2, intRow2, Trim(varTmp)
''                            If intRow1 = (5 * fNum1) + 1 Then fNum1 = fNum1 + 1: fNum2 = 0
''                            fNum2 = fNum2 + 1
'                            If intRow1 > 1 Then
'                                spdResult1.Row = intRow1 - 1
'                                spdResult1.Col = 4
'                                If spdResult1.Text = "5" Then
'                                    fNum1 = Val(spdResult1.Text) + 1
'                                    'fNum1 = fNum1 + 1
'                                    fNum2 = 1
'                                Else
'                                    fNum2 = Val(spdResult1.Text) + 1
'                                    spdResult1.Col = 3
'                                    fNum1 = spdResult1.Text
'                                End If
'                            Else
'                                spdResult1.SetText 2, intRow2, Trim(varTmp)
'                                If intRow1 = (5 * fNum1) + 1 Then fNum1 = fNum1 + 1: fNum2 = 0
'                                fNum2 = fNum2 + 1
'                            End If
'
'                            spdResult1.SetText 1, intRow2, "1"
'                            spdResult1.SetText 3, intRow2, Format(Val(fNum1), "00")
'                            spdResult1.SetText 4, intRow2, Trim(fNum2)
'                        Else
'                            spdResult1.maxrows = spdResult1.maxrows - 1
'                        End If
'                        bgetWork = False
'                    Else
'                        spdResult1.SetText 2, intRow2, Trim(varTmp)
'                        If intRow1 > 1 Then
'                            spdResult1.Row = intRow1 - 1
'                            spdResult1.Col = 4
'                            If spdResult1.Text = "5" Then
'                                spdResult1.Col = 3
'                                fNum1 = Val(spdResult1.Text) + 1
'                                'fNum1 = fNum1 + 1
'                                fNum2 = 1
'                            Else
'                                fNum2 = Val(spdResult1.Text) + 1
'                                spdResult1.Col = 3
'                                fNum1 = spdResult1.Text
'                            End If
'                        Else
'                            spdResult1.SetText 2, intRow2, Trim(varTmp)
'                            If intRow1 = (5 * fNum1) + 1 Then fNum1 = fNum1 + 1: fNum2 = 0
'                            fNum2 = fNum2 + 1
'                        End If
'                        spdResult1.SetText 1, intRow2, "1"
'                        spdResult1.SetText 3, intRow2, Format(Val(fNum1), "00")
'                        spdResult1.SetText 4, intRow2, Trim(fNum2)
'                    End If
'                End If
'                'spdResult1.SetText 1, intRow2, "1"
'                spdResult1.maxrows = intRow2
'
'                .Row = intRow1: .Col = 1
'                .Value = 0
'                '.Row = intRow1
'                '.Action = ActionDeleteRow
'                'If .maxrows > 14 Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
'
'                'If intRow1 > 0 Then intRow1 = intRow1 - 1
'
'            End If
'        Next
'    End With
'
'    '-- 임시
'    'comEQP.Output = STX + "M     " + ETX
'
'    f_strJOB_FLAG = "1"
'    f_intSampleNo = 0
    
End Sub

Private Function f_funget_code(ByVal Testcd As String)

    Dim intIdx1 As Integer, intIdx2 As Integer
    
    f_funget_code = ""
    For intIdx1 = 1 To UBound(Channel)
        For intIdx2 = 1 To Channel(intIdx1).MachCnt
            If Channel(intIdx1).MachTst(intIdx2) = Testcd Then
                f_funget_code = Channel(intIdx1).MachCode
                Exit Function
            End If
        Next
    Next
    
End Function
Private Sub comEQP_OnComm()
    
    On Error Resume Next
    
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim Arr()       As Byte
    Dim strdata     As String
    Dim brStr       As String
    Dim sStxCheck As Integer, sEtxCheck As Integer, sEtbcheck As Integer, sRstcheck As Integer
    Dim com_sTemp As String
    Dim ii As Integer, jj As Integer, kk As Integer
    Dim MHead  As String, Pinfo As String
    Dim PatientID As String
    
    Dim Orderoutput As String
    Dim Testoutput As String
    Dim OutPutData  As String
    Dim Rev As Long
    Dim Test_Cd() As String
    Dim sRow As Integer
    Dim oPatNo As String
    Dim oRackNo As String
    Dim oPosNo As String
    
    Dim intRow  As Integer
    
    Dim adoRS As ADODB.Recordset
    Dim sqlDoc As String
    Dim itemX As ListItem
    Dim strEqpCd1 As String
    
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
'            Debug.Print fRcvString
            For ii = 1 To Len(brStr)
                Select Case Mid(brStr, ii, 1)
                    Case STX
                        fRcvString = ""
                        fRcvString = fRcvString + Mid(brStr, ii, 1)
                    Case ETX
                        fRcvString = fRcvString + Mid(brStr, ii, 1)
                        If Mid(fRcvString, 1, 3) = STX + ACK + ETX And SendCount = 0 Then
                            Debug.Print "[TBA2] " & STX + ACK + ETX
                            OrderCnt = OrderCnt + 1
                            TbaStat = True
                            fRcvString = ""
                        ElseIf Mid(fRcvString, 1, 3) = STX + NAK + ETX And SendCount > 0 Then
                            Debug.Print "[TBA2] " & STX + NAK + ETX
                            Timer2.Enabled = True
                            Timer2.Interval = 20000
                            fRcvString = ""
                        ElseIf Mid(fRcvString, 1, 8) = STX + "M 0001" + ETX Then 'And SendCount > 0 And SendCount <= spdResult1.maxrows Then
                            Debug.Print "[TBA2] " & STX + "M 0001" + ETX
                            '-- 오더송신
                            '1       10        20        30        40        50        60        70        80        90       100       110       120       130
                            '1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
                            
                            'O 20030724C0100105      0201  1   81   31   41   51   11   21  101   61   71  171  731  141  161   91  111  121  241200307241735                                                                                                   M     
                            f_intCnt = f_intCnt + 1
                            If f_intCnt <= UBound(f_strBarno) And f_strBarno(1) <> "" Then
                                comEQP.Output = STX & ACK & ETX
                                Debug.Print "[HOST] " & STX & ACK & ETX
                                
                                Orderoutput = STX + "O "
                                Orderoutput = Orderoutput + f_strBarno(f_intCnt) + Space$(20 - Len(Trim(f_strBarno(f_intCnt))))
                                Orderoutput = Orderoutput + "  " + f_strRack(f_intCnt)  '-- Rack
                                Orderoutput = Orderoutput + f_strCup(f_intCnt)  '-- cup
                                '-- Manual Dil
                                Orderoutput = Orderoutput + "  1"
                                
                                If Len(f_strTest(f_intCnt)) > 0 Then
                                    Orderoutput = Orderoutput + f_strTest(f_intCnt) + ETB
                                Else
                                    MsgBox oPatNo & "의 해당 검사항목이 없습니다.!", vbCritical
                                    comEQP.Output = STX + "I " + ETX
                                    Debug.Print "[HOST] " & STX + "I " + ETX
                                    OrderCnt = 0
                                    SendCount = 0
                                    Exit Sub
                                End If
                                '-- Order Date/Time
                                Orderoutput = Orderoutput + Format(Now, "yyyymmddhhmm")
                                '-- Name
                                Orderoutput = Orderoutput + Space$(30)
                                '-- Sex
                                Orderoutput = Orderoutput + Space$(1)
                                '-- Birth Date
                                Orderoutput = Orderoutput + Space$(8)
                                '-- Location
                                Orderoutput = Orderoutput + Space$(20)
                                '-- Doctor
                                Orderoutput = Orderoutput + Space$(20)
                                '-- Comment
                                Orderoutput = Orderoutput + Space$(20)
                                Orderoutput = Orderoutput + ETB + ETX
                                
                                comEQP.Output = Orderoutput
                                
                                Debug.Print "[HOST] " + Orderoutput
                                OrderCnt = OrderCnt + 1
                                fRcvString = ""
                                
                                lblStatus.Caption = "오더전송중.."
                                intRow = f_funGet_SpreadRow(spdResult1, 2, f_strBarno(f_intCnt))
                                If intRow > 0 Then
                                    spdResult1.Row = intRow:    spdResult1.Row2 = intRow
                                    spdResult1.Col = 2:         spdResult1.Col2 = 4
                                    spdResult1.BlockMode = True
                                    spdResult1.BackColor = vbCyan
                                    spdResult1.BlockMode = False
                                    spdResult1.SetText 1, intRow, ""
                                End If
                                
                                If f_intCnt >= UBound(f_strBarno) Then
                                    Erase f_strBarno:   Erase f_strTest:    Erase f_strRack:    Erase f_strCup
                                    ReDim Preserve f_strBarno(1 To 1) As String
                                    ReDim Preserve f_strRack(1 To 1) As String
                                    ReDim Preserve f_strCup(1 To 1) As String
                                    ReDim Preserve f_strTest(1 To 1) As String
                                    
                                    f_strBarno(1) = "":  f_strRack(1) = "":    f_strCup(1) = ""
                                    f_strTest(1) = ""
                                    
                                    TbaStat = False
                                    f_intCnt = 0
                                    
                                    Timer3.Enabled = True
                                    Timer3.Interval = 2000
                                Else
                                    SendCount = SendCount + 1 '-- SendCount = 2
                                    Timer2.Enabled = False
                                    
                                    comEQP.Output = STX & "M     " & ETX
                                    
                                    Debug.Print "[HOST] " & STX & "M     " & ETX
                                End If
                            Else
                                TTT = TTT + 1
                                
                                Timer3.Enabled = True
                                Timer3.Interval = 20000
                                
                                lblStatus.Caption = "결과대기중.."
                            End If
                            fRcvString = ""
                            
'                        ElseIf Mid(fRcvString, 1, 8) = STX + "M 0001" + ETX And SendCount > 0 And SendCount > spdResult1.maxrows Then
'                            Debug.Print "[TBA2] " & STX + "M 0001" + ETX
'                            TTT = TTT + 1
'                            Timer3.Enabled = True
'                            Timer3.Interval = 20000
                        '-- 결과대기
                        ElseIf Mid(fRcvString, 1, 3) = STX + NAK + ETX Then ' And SendCount > 0 Then
                            '-- Communication un unSuccess
                            Debug.Print "[TBA2] " & STX + NAK + ETX
                            comEQP.Output = STX & "M     " & ETX
                            Debug.Print "[HOST] " & STX & "M     " & ETX
                            'TbaStat = False
                            'SendCount = 0
                            fRcvString = ""
                            lblStatus = "결과대기중.."
                            'Me.Enabled = False
                        '-- 결과수신
                        ElseIf Mid(fRcvString, 1, 3) = STX + "R " Then
                            Debug.Print "[TBA2] " & fRcvString
                            sStxCheck = InStr(fRcvString, STX)
                            sEtxCheck = InStr(fRcvString, ETX)
                            sRstcheck = InStr(fRcvString, "R")
                            
                            If UCase(Mid(fRcvString, 2, 1)) = "R" And sStxCheck <> 0 And sEtxCheck <> 0 And sRstcheck <> 0 Then
                                lblStatus = "결과수신중.."
                                Me.Enabled = True
                                Call psDataDefine(fRcvString, fChannel(), spdResult1)
                                'comEQP.Output = STX + ACK + ETX
                                'Debug.Print "[HOST] " & STX + ACK + ETX
                                'comEQP.Output = STX + "M     " + ETX
                                'Debug.Print "[HOST] " & STX + "M     " + ETX
                            End If
                        
                        End If
                    
                    Case Else
                        fRcvString = fRcvString + Mid(brStr, ii, 1)
                End Select
            Next ii
            
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

Public Sub OrderingTheDataElecsys(ByVal brCom As Object, ByVal brbarcd As String, ByVal brSpread As Object, ByRef brChannel() As String) ', ByRef brItemdeci() As String)
Dim Orderoutput As String
Dim OutPutData As String
Dim Testcd As String, sOrderLst As String
Dim ii As Integer, Loop_count As Integer, pDoCount As Integer
Dim rv As Long
Dim tst_cd() As String

    For Loop_count = 1 To 100: fTBA200(Loop_count) = "": Next Loop_count

    pDoCount = 0
    '-- 날짜와 검체번호로 오더를 조회한다.
'    rv = sl_spcid_tstcd_select(PatientID, tst_cd)
'    If rv = 0 Then
'        MsgBox "해당검체 없음"
'    Else
'        ReDim tst_cd(rv)
'        For II = 1 To rv
'            sOrderLst = sOrderLst + "," + tst_cd(II)
'        Next
'    End If
'    sOrderLst = sOrderLst + ","
'
'    Do While InStr(sOrderLst, ",") > 0
'        pDoCount = pDoCount + 1
'        fTBA200(pDoCount) = Text_Redefine(sOrderLst, ",")
'        sOrderLst = Mid$(sOrderLst, InStr(sOrderLst, ",") + 1)
'        If pDoCount > 99 Then
'            sOrderLst = ""
'            Exit Do
'        End If
'    Loop
    
'    For II = 1 To pDoCount
'        If Val(fTBA200(II)) > 0 Then
'            OutPutData = OutPutData + "^^^" & fTBA200(II) & "^" & "0"
'            If II = pDoCount Then
'                Exit For
'            Else
'                OutPutData = OutPutData + "\"
'            End If
'        End If
'    Next II
    '-- 임시 검사항목
    OutPutData = "^^^250^0\^^^10^0"
    
'    Orderoutput = "3O" & "|1|" & PatientID & "|" & PatientSeq & "|" & OutPutData & "|R|" & Format(Now, "YYYYMMDDHHMMSS") & "|||||N||||||||||||||Q"
'    OutPutData = STX & Orderoutput & vbCr & ETX & MakeCS(Orderoutput) & vbCr & vbLf
    
    '-- 임시
    Orderoutput = "3O" & "|1|" & "0001" & "|" & "011" & "|" & OutPutData & "|R|" & Format(Now, "YYYYMMDDHHMMSS") & "|||||N||||||||||||||Q"
    OutPutData = STX & Orderoutput & vbCr & ETX & MakeCS(Orderoutput) & vbCr & vbLf
    
    comEQP.Output = OutPutData
    Debug.Print "[HOST] " & OutPutData
    OutPutData = ""
    PatientID = ""
    PatientSeq = ""
    PatientDisk = ""
    PatientPos = ""
    
End Sub

Private Function MakeCS(Source As String) As String
    Dim x      As Long
    Dim ChkCS  As String
    Dim SumCS  As String
    Dim AddCS  As Long
    
    For x = 1 To Len(Source)
        AddCS = AddCS + Asc(Mid(Source, x, 1))
    Next x
    AddCS = AddCS + Asc(Chr(13)) + Asc(ETX)
    AddCS = AddCS Mod &H100
    SumCS = Hex(AddCS)
    If Len(SumCS) = 1 Then
        ChkCS = "0" & SumCS
    Else
        ChkCS = Mid(SumCS, Len(SumCS) - 1, 1)
        ChkCS = ChkCS & Right(SumCS, 1)
    End If
    MakeCS = ChkCS
End Function

Private Sub ComReceive(ByRef RecData() As Byte)
    
    Dim strRec  As String, strBuff  As String
    
    Dim strTmp  As String, strDta() As String
    Dim intIdx  As Integer, intCnt  As Integer, intPos  As Integer
    Dim sStxCheck As Integer, sEtxCheck As Integer
    Dim com_sTemp As String
    Dim fOpt        As String

    Static OrgMsg As String
    
'    strRec = StrConv(RecData, vbUnicode)
    strRec = RecData
    Print #1, strRec;
    
    Call COM_INPUT(strRec)
    
    For intIdx = 1 To Len(strRec)
        strBuff = Mid$(strRec, intIdx, 1)
        Select Case f_strJOB_FLAG
            Case "1"    '-- 대기
                        Select Case Asc(strBuff)
                            Case 2  '-- STX
                                    f_strJOB_FLAG = "2"
                        End Select
                        f_strBuffer = f_strBuffer + strBuff
            Case "2"    '--  받기
                        Select Case Asc(strBuff)
                            Case 2      '-- STX
                                        f_strJOB_FLAG = "2"
                            Case 3      '-- ETX
                                        f_strBuffer = f_strBuffer + strBuff
                                        'strTmp = f_strBuffer
                                        sStxCheck = InStr(f_strBuffer, Chr(2))
                                        sEtxCheck = InStr(f_strBuffer, Chr(3))
                                        If sStxCheck <> 0 And sEtxCheck <> 0 Then
                                            com_sTemp = Mid$(f_strBuffer, sStxCheck, sEtxCheck)
                                            f_strBuffer = Mid$(f_strBuffer, sEtxCheck + 1)
                                            If optSeq.Value = True Then
                                                fOpt = "0"
                                            Else
                                                fOpt = "1"
                                            End If
                                            'Call psDataDefine(com_sTemp, fChannel(), spdResult1, fOpt) ', brSpread, brChannel(), brItemdeci(), brOpt
                                        End If
                                        
'                                        intPos = InStr(strTmp, "D0U"):  intCnt = 0: Erase strDta
'                                        Do While intPos > 0
'                                            intCnt = intCnt + 1
'                                            ReDim strDta(1 To intCnt) As String
'
'                                            strDta(intCnt) = Mid$(strTmp, 1, intPos - 1)
'                                            strTmp = Mid$(strTmp, intPos + 1)
'                                            intPos = InStr(strTmp, "D0U")
'                                        Loop
'                                        If strTmp <> "" Then
'                                            intCnt = intCnt + 1
'                                            ReDim strDta(1 To intCnt) As String
'
'                                            strDta(intCnt) = strTmp
'                                        End If
                                        
'                                        For intPos = 1 To intCnt
'                                            Call f_subSet_Result(strDta(intPos))
'                                        Next
                                        
                                        Call COM_OUTPUT(Chr(6))
                                        f_strBuffer = ""
                                        f_strJOB_FLAG = "1"
                            
                            Case Else
                                        f_strBuffer = f_strBuffer + strBuff
                        End Select
        End Select
     Next
End Sub

Private Sub psDataDefine(ByVal brbarcd As String, ByRef brChannel() As String, ByVal brSpread As Object)
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
Dim sPatiant_No As Long
Dim itemX As ListItem
Dim strRstval(1 To 19) As String, strRefVal(1 To 19)  As String
Dim FunStr As String
Dim sqlDoc  As String
Dim intCol As Integer
Dim ii As Integer, jj As Integer, kk As Integer
Dim Test_Cd() As String
Dim Rev As Long
Dim tmpTstCd As String
Dim tmT As String
Dim iRow As Integer

Dim strTime As String

'R000120030724C0100105       2^ 1^^
'200307241808^  1^^
'   1   6.1 1A
'   2   3.7 1A   3  16.5 1A   4   1.0 1A   5   5.4 1A   6   0.8 1A   7  0.18 1A   8   141 1A   9   181 1A  10   247 1A  11    68 1A  12   177 1A  14   8.9 1A  16   2.7 1A  17    26 1A  24  10.0 1E  73   109 1A200307241735                               
'sRstText = "R 000120030724C0100105       2 1200307241808  1   1   6.1 1A   2   3.7 1A   3  16.5 1A   4   1.0 1A   5   5.4 1A   6   0.8 1A   7  0.18 1A   8   141 1A   9   181 1A  10   247 1A  11    68 1A  12   177 1A  14   8.9 1A  16   2.7 1A  17    26 1A  24  10.0 1E  73   109 1A200307241735                               "

    'On Error GoTo errDefine
    On Error Resume Next
    
    sRstText = brbarcd
    '------------------------------<<< fTBA200() 배열 Clear 한다.         >>>----------
    'For Loop_count = 1 To 100: fTBA200(Loop_count) = "": Next Loop_count
    '------------------------------<<< fTBA200() 배열에 구분하여 넣는다.  >>>----------
        
    PatientID = Trim(Mid(sRstText, 8, 20))
    PatientRack = Trim(Mid(sRstText, 28, 4))
    PatientPos = Trim(Mid(sRstText, 32, 2))
'    Patiant_Recevid = False        ' 환자번호 Flag
'    sPatiant_No = fTBA200(3)  ' 환자번호
    
    pDoCount = 0
    For ii = 49 To Len(sRstText) Step 13
        pDoCount = pDoCount + 1
        fTBA200(pDoCount) = Mid(sRstText, ii, 13)
'        Debug.Print fTBA200(pDoCount)
        If pDoCount > 99 Or Mid(sRstText, ii + 13, 1) = ETB Then
            'pDoCount = pDoCount - 1
            sRstText = ""
            Exit For
        End If
    Next
    
    '----<<< 해당검사결과와 해당환자를 찿는다. >>>----------
    Patiant_Recevid = False
    With brSpread
        For ii = 1 To .maxrows
            .Row = ii: .Col = 2
            If Trim(.Text) = Trim(PatientID) Then
                iRow = ii
                If ii = .maxrows Then
                    lblStatus = "결과수신완료"
'                    Timer3.Enabled = False
                    Me.Enabled = True
                End If
                Patiant_Recevid = True
                Exit For
            End If
        Next
    End With
    If Patiant_Recevid = False Then Exit Sub
    
    If Patiant_Recevid = True Then
        '-- 장비검사수
        For ii = 1 To pDoCount
            '-- 검사한 장비번호
            Channel_No = Trim(Mid(fTBA200(ii), 1, 4))  ' channel
            'Max_Arary_Cnt = brSpread.MaxCols - 2   ' 앞에서부터 5까지는 환자 정보 이기때문에.... -6를 한다.
            For jj = 1 To spdResult1.MaxCols - 6 '37 '100 UBOUND(brChannel)
            With brSpread
                '----<<<<<<<<<,  세부검사항목을 찿는다.  >>>>>>>----------
                '.Col = pDoCount + 4
                'If brChannel(jj) = "1" Then Stop
                If Len(Channel_No) > 0 And Channel_No = brChannel(jj) Then          ' 검사결과가 있으면...
                    .Col = jj + 6
                    If Trim(fTBA200(ii)) <> "" Then
                        .Text = Mid(fTBA200(ii), 5, 6)
                    Else
                        .Text = ""
                    End If
                    FunStr = .Text
                    
                    If FunStr <> "" Then
                        'Set itemX = lvwCuData.FindItem(brChannel(jj), lvwTag, , lvwWhole)
                        Set itemX = lvwCuData.FindItem(brChannel(jj), lvwTag, , lvwWhole)
                        If Not itemX Is Nothing Then '-- itemX : 검사코드
                            If itemX.ListSubItems(8) <> "" And itemX.ListSubItems(9) <> "" Then
                                If Val(.Text) < itemX.ListSubItems(8) Then
                                    strRefVal(pDoCount) = "L"
                                ElseIf Val(.Text) > itemX.ListSubItems(9) Then
                                    strRefVal(pDoCount) = "H"
                                End If
                            End If
                                                        
                            Rev = sl_spcid_tstcd_select(Trim(PatientID), Test_Cd)
                            tmpTstCd = ""
                            For kk = 0 To Rev - 1
                                If InStr(itemX.ListSubItems(1), Trim(Test_Cd(kk))) > 0 Then
                                      tmpTstCd = "" & Trim(Test_Cd(kk))
                                      Exit For
                                End If
                            Next kk
                            If UBound(strRefVal) < pDoCount Then
                                tmT = ""
                            Else
                                tmT = "" 'strRefVal(pDoCount)
                            End If
                            
                            strTime = Format$(Now, "MMSS")
                            sqlDoc = "insert into INTERFACE003(" & _
                                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN)" & _
                                     "    values( '" & Trim(PatientID) & "', '" & tmpTstCd & "', '" & brChannel(jj) & "'," & _
                                     "            '" & Format$(Now, "YYYYMMDD") & "', '" & strTime & "'," & _
                                     "            '" & Trim(FunStr) & "', '" & tmT & "'," & _
                                     "            '" & INS_CODE & "', '')"
                            
                            AdoCn_Jet.Execute sqlDoc
                            
                            intCol = itemX.Index + 6
                            .Col = intCol
                            .Row = iRow
                            .ForeColor = IIf(strRefVal(ii) <> "", vbRed, vbBlack)
                            
                            '-- 서버결과등록
                            If Rev > 0 And chkAuto.Value = vbChecked Then
                                If f_funAdd_Server(Trim(PatientID), tmpTstCd, Trim(FunStr), Test_Cd) Then
                                    spdResult1.Row = iRow
                                    spdResult1.Col = intCol:    spdResult1.BackColor = &HFFF8F0
                                
                                    sqlDoc = "Update INTERFACE003" & _
                                             "   set SERVERGBN  = 'Y'" & _
                                             " where SPCNO   = '" & Trim(PatientID) & "'" & _
                                             "   and EQPNUM  = '" & brChannel(jj) & "'" & _
                                             "   and TRANSDT = '" & Format$(Now, "YYYYMMDD") & "'" & _
                                             "   and TRANSTM = '" & strTime & "'"
                                    AdoCn_Jet.Execute sqlDoc
                                End If
                            End If
                            
                            Exit For
                        End If
                        Set itemX = Nothing
                    End If
                    
                End If
            End With
            Next jj
        Next ii
    End If
    Timer2.Enabled = False
    Exit Sub
errDefine:

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

Public Function SeqSearch(ByVal brSpread As Object, ByVal brSeq As Long, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch = 0
    If brSpread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brSpread
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

    'Dim intIdx          As Integer
    'Dim strdata(1 To 5)         As Byte
    
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
    
    Call comEQP_OnComm
    
'    comEQP.Output = STX & "M     " & ETX
'    Debug.Print "[HOST] " & STX & "M     " & ETX

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
    
    cboTest.ListIndex = 0
    
    Call cmdClear               ' 초기화
    Call f_subSet_ItemHeader    ' 리스트해더
    Call f_subSet_ItemList      ' 검사항목
    
    Call f_subSet_ComCharacter  ' 통신문자
    Call f_subGet_Setting       ' 통신설정
    
    Call cmdRun           ' 실행
    
    mskRstDate.Text = Format$(Now, "YYYYMMDD")
    mskOrdDate.Text = Format$(Now, "YYYYMMDD")
    tabWork.tag = 0
    
    Open App.Path + "\" + "Toshiba.log" For Append As #1

    f_strJOB_FLAG = "1":    f_intSampleNo = 0
    cboRstgbn(0).ListIndex = 0: cboRstgbn(1).ListIndex = 2
    'cboTest.ListIndex = 0
    
    SendCount = 0
    TbaStat = False
    
    Timer1.Enabled = True
    Timer1.Interval = 2000
    f_intCnt = 0
    
    t_no = ""
    
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
                .RThreshold = 1
                .SThreshold = 1
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
    
End Sub


Private Sub spdResult1_Change(ByVal Col As Long, ByVal Row As Long)

    Dim varTmp  As Variant
    Dim intRow  As Integer
    Dim intRack As Integer, intCupno    As Integer
    
    With spdResult1
        .GetText 3, Row, varTmp: intRack = Val(varTmp)
        .GetText 4, Row, varTmp: intCupno = Val(varTmp)
        
        Select Case Col
            Case 3  '-- Rack변호 변경시
                    For intRow = Row + 1 To .maxrows
                        .GetText 2, intRow, varTmp
                        If Trim$(varTmp) = "" Then Exit For
                        
                        .SetText 3, intRow, Format$(intRack, "00")
                    Next
            Case 4  '-- Cup번호 변경시
                    For intRow = Row + 1 To .maxrows
                        .GetText 2, intRow, varTmp
                        If Trim$(varTmp) = "" Then Exit For
                        
                        intCupno = intCupno + 1
                        If intCupno > 5 Then
                            intCupno = 1
                            intRack = intRack + 1
                        End If
                        .SetText 3, intRow, Format$(intRack, "00")
                        .SetText 4, intRow, CStr(intCupno)
                    Next
        End Select
    End With
    
End Sub

Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Col > 6 Then
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

Private Sub spdResult1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If spdResult1.maxrows = spdResult1.maxrows Then
            spdResult1.maxrows = spdResult1.maxrows + 1
        End If
    End If
    
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

Private Sub spdWorkList_Click(ByVal Col As Long, ByVal Row As Long)

    If Col < 3 Then Exit Sub
    
    Dim varTmp  As Variant
    
    With spdWorkList
        If Col = 1 Then
            .GetText 2, Row, varTmp
            If Trim$(varTmp) = "" Then Exit Sub
            
            .SetText 1, Row, IIf(Trim$(varTmp) = "1", "", "1")
        ElseIf Col > 6 Then
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
    
    Call cmdInit_Click

End Sub

Private Sub Timer2_Timer()
    
    If TTT > 0 Then
        comEQP.Output = STX + "M     " + ETX
        Debug.Print "[HOST] " & STX + "M     " + ETX
        Exit Sub
    End If
    If SendCount > 0 Then
'        Call cmdOrder_Click
        'If chkMan.Value <> 1 Then
            Timer2.Enabled = False
        'End If
'    Else
'        Timer2.Interval = 4000
'        comEQP.Output = STX + "M     " + ETX
'        Debug.Print "[HOST] " & STX + "M     " + ETX
    End If
    
End Sub

Private Sub Timer3_Timer()
    
    comEQP.Output = STX + "M     " + ETX
    Debug.Print "[HOST] " & STX + "M     " + ETX

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
    
    Dim tst_no() As String
    Dim TMP() As String
    Dim rv As Long
    Dim samChk As Boolean
    Dim ii As Integer
    Dim bgetWork As Boolean
    Dim itemX As ListItem
    
    Dim strEqpCd   As String
    
    samChk = False
    If KeyAscii = vbKeyReturn Then
        rv = sl_spcid_tstcd_select(Trim(txtBarCode.Text), tst_no)
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
                strEqpCd = f_funget_code(tst_no(ii))
                Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                If Not itemX Is Nothing Then
                    bgetWork = True
                    Exit For
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
                        .Col = 1
                        .Row = .maxrows + 1
                        '.Value = 1
                        .Col = 2
                        .maxrows = .maxrows + 1
                        .Row = .maxrows
                        .Action = ActionActiveCell
                        .Text = txtBarCode.Text
                        
                        .Col = 1
                        .Value = 1
                        
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
'
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
'
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



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
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
   Begin VB.PictureBox picWork 
      Height          =   5910
      Left            =   30
      ScaleHeight     =   5850
      ScaleWidth      =   11865
      TabIndex        =   11
      Top             =   570
      Width           =   11925
      Begin TabDlg.SSTab tabWork 
         Height          =   5850
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   10319
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   " WorkList"
         TabPicture(0)   =   "frmComm_2.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label5"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdWordlist"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdWKSend"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "mskOrdDate"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "chkSpdSel"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "spdWorkList"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtBarCode"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "optSeq"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "optBar"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   " 받은 결과"
         TabPicture(1)   =   "frmComm_2.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "spdResult"
         Tab(1).Control(1)=   "chkSel"
         Tab(1).Control(2)=   "cboRstgbn(1)"
         Tab(1).Control(3)=   "lvwData"
         Tab(1).Control(4)=   "mskRstDate"
         Tab(1).Control(5)=   "cmdAppend"
         Tab(1).Control(6)=   "cmdQuery"
         Tab(1).Control(7)=   "Label4"
         Tab(1).ControlCount=   8
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
            Left            =   10980
            TabIndex        =   29
            Top             =   0
            Width           =   735
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
            Left            =   10050
            TabIndex        =   28
            Top             =   0
            Value           =   -1  'True
            Width           =   735
         End
         Begin FPSpread.vaSpread spdResult 
            Height          =   4920
            Left            =   -74910
            TabIndex        =   26
            Top             =   840
            Width           =   11670
            _Version        =   196608
            _ExtentX        =   20585
            _ExtentY        =   8678
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
            MaxCols         =   4
            MaxRows         =   15
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmComm_2.frx":0038
         End
         Begin VB.TextBox txtBarCode 
            Height          =   300
            Left            =   4680
            MaxLength       =   11
            TabIndex        =   27
            Top             =   495
            Width           =   2085
         End
         Begin FPSpread.vaSpread spdWorkList 
            Height          =   4920
            Left            =   90
            TabIndex        =   25
            Top             =   840
            Width           =   11640
            _Version        =   196608
            _ExtentX        =   20532
            _ExtentY        =   8678
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   0
            ColsFrozen      =   4
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
            MaxCols         =   4
            MaxRows         =   15
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "frmComm_2.frx":050D
         End
         Begin VB.CheckBox chkSpdSel 
            Caption         =   "All"
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
            Height          =   300
            Left            =   105
            MaskColor       =   &H00400000&
            TabIndex        =   20
            Top             =   495
            Width           =   645
         End
         Begin VB.CheckBox chkSel 
            Caption         =   "All"
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
            Height          =   300
            Left            =   -74910
            MaskColor       =   &H00400000&
            TabIndex        =   15
            Top             =   495
            Width           =   645
         End
         Begin VB.ComboBox cboRstgbn 
            Height          =   300
            Index           =   1
            ItemData        =   "frmComm_2.frx":09EA
            Left            =   -70305
            List            =   "frmComm_2.frx":09F7
            Style           =   2  '드롭다운 목록
            TabIndex        =   14
            Top             =   495
            Width           =   2085
         End
         Begin MSComctlLib.ListView lvwData 
            Height          =   4965
            Left            =   -68745
            TabIndex        =   13
            Top             =   810
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   8758
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FlatScrollBar   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSMask.MaskEdBox mskRstDate 
            Height          =   300
            Left            =   -71430
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
         Begin HSCotrol.CButton cmdQuery 
            Height          =   300
            Left            =   -65460
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
            Left            =   3555
            TabIndex        =   21
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
         Begin HSCotrol.CButton cmdWKSend 
            Height          =   300
            Left            =   10665
            TabIndex        =   22
            Top             =   495
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "삭  제"
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
         Begin HSCotrol.CButton cmdWordlist 
            Height          =   300
            Left            =   9585
            TabIndex        =   23
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
            Left            =   2340
            TabIndex        =   24
            Top             =   570
            Width           =   1125
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
            Left            =   -72660
            TabIndex        =   19
            Top             =   570
            Width           =   1125
         End
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
            Picture         =   "frmComm_2.frx":0A21
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":0FBB
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":1555
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":1AEF
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":2089
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":2623
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
            Picture         =   "frmComm_2.frx":2BBD
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":3157
            Key             =   "NOT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":36F1
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":3C8B
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":451D
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":4677
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":47D1
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
      Height          =   585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   1032
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmComm_2.frx":492B
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
         Picture         =   "frmComm_2.frx":5BAD
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   9780
         Picture         =   "frmComm_2.frx":6137
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   8640
         Picture         =   "frmComm_2.frx":66C1
         Top             =   255
         Width           =   240
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

Private f_strBuffer As String
Private f_strSend   As String, f_strSendChr As String
Private f_strPCFlag As String
Private f_intTestNo As Integer
Private f_strJOB_FLAG   As String, f_intIdx As Integer
Private f_strSample()   As String, f_intCnt As Integer
Private f_strJOB_ACKETC As String
Private f_blnWorkList   As Boolean
Private f_blnJOB_Conent As Boolean
Private f_strOrdList    As String
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

Private Function f_funGet_SpreadRow(ByVal strPara As String, ByVal intCol As Long) As Long

    Dim intRow  As Integer
    Dim varTmp  As Variant
    
    f_funGet_SpreadRow = 0
    With spdWorkList
        For intRow = 1 To .maxrows
            .GetText intCol, intRow, varTmp
            If Trim$(varTmp) = strPara Then f_funGet_SpreadRow = intRow: Exit For
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
    '-- WorkList
    Call f_subSet_ItemComplete(spdWorkList)
    '-- Complete
    Call f_subSet_ItemComplete(spdResult)
    
End Sub

Private Sub f_subSet_ItemComplete(Spd As vaSpread)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim itemH           As ColumnHeader
    Dim objHeadeItem    As clsCommon
    
    Dim intCol  As Integer
    
    'lvw.ColumnHeaders.Clear
    'Call lvw.ColumnHeaders.Add(, "EQP_ID", "검체 번호")
    
    intCol = 2
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) AS TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = '" & INS_CODE & "') AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst: ReDim fChannel(adoRS.RecordCount)
    Do While Not adoRS.EOF
        With Spd
            intCol = intCol + 1
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1:  .ColWidth(.MaxCols) = 6.5
            
            .SetText intCol, 0, adoRS.Fields("TESTNM")
            fChannel(intCol - 2) = adoRS.Fields("TESTCD_EQP")
            Debug.Print fChannel(intCol - 2)
        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub f_subSet_ItemList()

'    Dim itemX   As ListItem
'    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    Dim ii      As Integer
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = ""
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) AS TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst: spdWorkList.Row = 1: ii = 0
    Do While Not adoRS.EOF
        With spdWorkList
            .Col = ii + 3
            .Text = Trim(adoRS.Fields("TESTNM") & "")
'        Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TESTCD_EQP") & ""), , "LST")
'            itemX.SubItems(1) = Trim(adoRS.Fields("TESTCD") & "")
'            itemX.SubItems(2) = Trim(adoRS.Fields("TESTNM") & "")
'            itemX.SubItems(3) = ""
'            itemX.SubItems(4) = Trim(adoRS.Fields("DELTA") & "")
'            itemX.SubItems(5) = Trim(adoRS.Fields("DELTAGBN") & "")
'            itemX.SubItems(6) = Trim(adoRS.Fields("PANICL") & "")
'            itemX.SubItems(7) = Trim(adoRS.Fields("PANICH") & "")
'            itemX.SubItems(8) = Trim(adoRS.Fields("REFL") & "")
'            itemX.SubItems(9) = Trim(adoRS.Fields("REFH") & "")
'            itemX.SubItems(10) = Trim(adoRS.Fields("AUTOVERIFY") & "")
'            itemX.SubItems(11) = Trim(adoRS.Fields("REMARK") & "")
'            itemX.Tag = Trim(adoRS.Fields("TESTCD_EQP") & "")
'            itemX.Text = Trim(adoRS.Fields("TESTCD") & "")
'        Set itemX = Nothing
        
'        Set itemA = lvwData.ListItems.Add(, , Trim(adoRS.Fields("TESTCD_EQP") & ""), , "LST")
'            itemA.SubItems(1) = Trim(adoRS.Fields("TESTNM") & "")
'            itemA.SubItems(6) = Trim(adoRS.Fields("REFL") & "") + " ~ " + Trim(adoRS.Fields("REFH"))
'            itemA.Tag = Trim(adoRS.Fields("TESTCD_EQP") & "")
'        Set itemA = Nothing
        
'        f_strOrdList = f_strOrdList + ", '" & Trim$(adoRS.Fields("TESTCD")) & "'"
        End With
        adoRS.MoveNext
    Loop
    Set adoRS = Nothing
    
    f_strOrdList = Mid$(f_strOrdList, 3)
    
Exit Sub
ErrRoutine:
    Set adoRS = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub chkSel_Click()

    Dim itemX   As ListItem
    
    For Each itemX In lvwComplete.ListItems
        itemX.SmallIcon = IIf(chkSel.Value = vbChecked, "LSE", "ITM")
    Next
    
End Sub

Private Sub chkSpdSel_Click()

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    For intRow = 1 To spdWorkList.maxrows
        spdWorkList.GetText 2, intRow, varTmp
        If Trim$(varTmp) = "" Then Exit For
        spdWorkList.SetText 1, intRow, IIf(chkSpdSel.Value = vbChecked, "1", "")
    Next
    
End Sub

Private Sub cmdAppend_Click()
    
    Dim itemLX  As ListItem:    Dim itemSX  As ListSubItem
    Dim itemLA  As ListItem
    Dim objSave As clsEqpResult
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strChatno   As String, strOrdcd As String
    
    CallForm = "frmComm - Private Sub cmdServer_Click()"
On Error GoTo ErrorRoutine
    
    Me.MousePointer = 11
    For Each itemLX In lvwComplete.ListItems
        If itemLX.SmallIcon = "LSE" Then
            Set objSave = New clsEqpResult
            With objSave
                .EQPNUM = itemLX.Text
                .SPCID = itemLX.Text
                
                .SPCTYPE = itemLX.tag
                For Each itemSX In itemLX.ListSubItems
                    '서브아이템에 검사 결과 가 있으면
                    If Trim(itemSX.Text) <> "" Then
                        
                        strChatno = "": strOrdcd = ""
                        strOrdcd = lvwComplete.ColumnHeaders(itemSX.Index + 1).tag
                        
                        sqlDoc = "select distinct CHART_NO, ORD_CODE from L3A01" & _
                                 " where SAMPLE_DATE = '" & Mid$(mskRstDate.Text, 1, 4) + Mid$(itemLX.Text, 1, 4) & "'" & _
                                 "   and SAMPLE_SEQ  = '" & Mid$(itemLX.Text, 5, 3) & "'" & _
                                 "   and PART        = '" & Mid$(itemLX.Text, 8, 2) & "'" & _
                                 "   and ORD_CODE   in ( '" & strOrdcd & "')"
                        adoRS.CursorLocation = adUseClient
                        adoRS.Open sqlDoc, AdoCn_SQL
                        If adoRS.RecordCount > 0 Then adoRS.MoveFirst
                        If Not adoRS.EOF Then
                            strChatno = Trim$(adoRS(0) & "") ': strOrdcd = Trim$(adoRS(1) & "")
                            
                            sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
                                     " where SPCNO  = '" & Mid$(itemLX.Text, 5, 3) & "'" & _
                                     "   and TESTCD = '" & lvwComplete.ColumnHeaders(itemSX.Index + 1).tag & "'" & _
                                     "   and TRANSDT = '" & mskRstDate.Text & "'"
                            If Mid$(itemLX.Text, 10, 4) <> "" Then
                               sqlDoc = sqlDoc & "   and TRANSTM = '" & Mid$(itemLX.Text, 10, 4) & "'"
                            AdoCn_Jet.Execute sqlDoc
                        End If
                        adoRS.Close:    Set adoRS = Nothing
                        
                        sqlDoc = "exec p_l3a01interface" & _
                                "      'U'," & _
                                "      '" & Mid$(mskRstDate.Text, 1, 4) + Mid$(itemLX.Text, 1, 4) & "'," & _
                                "      '" & Mid$(itemLX.Text, 8, 2) & "'," & _
                                "      '" & strChatno & "'," & _
                                "      '" & Mid$(itemLX.Text, 5, 3) & "'," & _
                                "      '" & strOrdcd & "'," & _
                                "      '0', '" & itemSX.Text & "'"
                        AdoCn_SQL.Execute sqlDoc
                        
                        
                    End If
                Next
                
                Set itemLA = lvwComplete.FindItem(itemLX.Text, lvwText, , lvwWhole)
                If Not itemLA Is Nothing Then itemLA.SmallIcon = "ITM"
                Set itemLA = Nothing
            End With
        
            Set itemSX = Nothing
            Set objSave = Nothing
        End If
    Next
    Set itemLX = Nothing
    
    For Each itemLX In lvwData.ListItems
        itemLX.SubItems(2) = ""
        itemLX.SubItems(3) = ""
        itemLX.SubItems(4) = ""
        itemLX.SubItems(5) = ""
    Next
    Set itemLX = Nothing

    Me.MousePointer = 0
    MsgBox "작업이 완료되었습니다.", vbInformation, Me.Caption
    
Exit Sub
ErrorRoutine:
    Set itemLX = Nothing
    Set itemLA = Nothing
    Set itemSX = Nothing
    Set objSave = Nothing
    Me.MousePointer = 0
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
    
    Dim itemX As ListItem
    Dim itemS As ListSubItem
    
'    For Each itemX In lvwCuData.ListItems
'        itemX.SubItems(3) = ""
'    Next
'    Set itemX = Nothing
'
'    Do
'        Set itemX = lvwComplete.SelectedItem
'        If itemX Is Nothing Then Exit Do
'        lvwComplete.ListItems.Remove (itemX.Index)
'    Loop
'    Set itemX = Nothing
'
'    For Each itemX In lvwData.ListItems
'        itemX.SubItems(2) = ""
'        itemX.SubItems(3) = ""
'        itemX.SubItems(4) = ""
'        itemX.SubItems(5) = ""
'    Next
'    Set itemX = Nothing
    
    f_strJOB_FLAG = "1":     f_strJOB_ACKETC = "1": f_blnJOB_Conent = False
    f_blnWorkList = False
    
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

Private Sub cmdENQ_Click()
    
    Call COM_OUTPUT(charCOM_Convert(COM_ENQ))

End Sub

Private Sub cmdQuery_Click()

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    CallForm = "frmComm - Private Sub cmdQuery_Click()"
    
On Error GoTo ErrorRoutine
    Me.MousePointer = 11
    
    '-- CLEAR
    Do
        Set itemX = lvwComplete.SelectedItem
        If itemX Is Nothing Then Exit Do
        lvwComplete.ListItems.Remove (itemX.Index)
    Loop
    Set itemX = Nothing
    
    sqlDoc = "select SPCNO, TESTCD, RSTVAL, REFVAL, TRANSDT" & _
             "  from INTERFACE003" & _
             " where TRANSDT = '" & mskRstDate.Text & "'"
    If cboRstgbn(1).ListIndex = 0 Then
        sqlDoc = sqlDoc & "   and SERVERGBN = ''"
    ElseIf cboRstgbn(1).ListIndex = 1 Then
        sqlDoc = sqlDoc & "   and SERVERGBN = 'Y'"
    End If
    
    sqlDoc = sqlDoc & " order by TRANSTM"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
    
        Set itemX = lvwComplete.FindItem(Trim$(adoRS(0) & ""), lvwText, , lvwWhole)
        If itemX Is Nothing Then
            Set itemX = lvwComplete.ListItems.Add(, , Trim$(adoRS(0) & "") + Trim$(adoRS(4)))
            If Not itemX Is Nothing Then
                With itemX
                    .Key = COL_KEY & Trim$(adoRS(0) & "") + Trim$(adoRS(4))
                    .tag = "G"
                    .Text = Trim$(adoRS(0) & "") + Trim$(adoRS(4))
                    .SmallIcon = "LST"
                End With
            End If
        End If
        
        itemX.SubItems(lvwComplete.ColumnHeaders(COL_KEY & Trim$(adoRS(1) & "")).SubItemIndex) = Trim$(adoRS(2) & "")
        
        itemX.ListSubItems(lvwComplete.ColumnHeaders(COL_KEY & Trim$(adoRS(1) & "")).SubItemIndex).ForeColor = vbBlack
        '-- 참고치판정
        Set itemA = lvwCuData.FindItem(Trim$(adoRS(1) & ""), lvwTag, , lvwWhole)
        If Not itemA Is Nothing Then
            If Val(adoRS(2) & "") < Val(itemA.SubItems(8)) Or Val(adoRS(2) & "") > Val(itemA.SubItems(9)) Then
               itemX.ListSubItems(lvwComplete.ColumnHeaders(COL_KEY & Trim$(adoRS(1) & "")).SubItemIndex).ForeColor = vbRed
            End If
        End If
        Set itemA = Nothing
        
        Set itemX = Nothing
        
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
    Me.MousePointer = 0
    Exit Sub
ErrorRoutine:
    Set itemX = Nothing
    
    Call ErrMsgProc(CallForm)
    Me.MousePointer = 0
    
End Sub

Private Sub cmdWKSend_Click()

    f_blnWorkList = True
    f_strJOB_FLAG = 5
    f_intIdx = 1
    Call COM_OUTPUT(Chr(4) + Chr(1))
'
'    Dim intRow  As Integer, intIdx2 As Integer
'    Dim varTmp  As Variant
'    Dim strSampleno As String, strSector As String, strCup As String
'    Dim strPcFlag   As String, strSpec   As String
'    Dim strOrder    As String, intOrdcnt As Integer
'
'    f_intIdx = 1
'    For intIdx2 = f_intIdx To spdWorkList.MaxRows
'        With spdWorkList
'            .GetText 2, intIdx2, varTmp:    strSampleno = Trim$(varTmp)
'            .GetText 3, intIdx2, varTmp:    strSector = Format$(varTmp, "@@")
'            .GetText 4, intIdx2, varTmp:    strCup = Format$(varTmp, "@@")
'            .GetText 1, intIdx2, varTmp
'            If Trim$(varTmp) = "1" And strSampleno <> "" And strSector <> "" And strCup <> "" Then
'                strSampleno = strSampleno + Space(11 - Len(strSampleno))
'                Call f_subGet_WorkList(strOrder, intOrdcnt, strSpec, strPcFlag, intIdx2)
'                f_strSend = "[ 0,701,01," + strSector + "," + strCup + ",0,ST," + strSpec & "," + _
'                            strSampleno + "," + String(20, " ") + "," + _
'                            String(25, " ") + "," + String(25, " ") + "," + _
'                            String(18, " ") + "," + String(15, " ") + ", ," + _
'                            strPcFlag + Space(10) + "," + String(18, " ") + "," + _
'                            Format(Now, "ddmmyy") + "," + Format(Now, "hhmm") + "," + _
'                            String(20, " ") + ",000,4," + String(6, " ") + ",F," + _
'                            String(25, " ") + "," + String(7, " ") + "," + String(4, " ") + "," + _
'                            String(4, " ") + "," + String(6, " ") + "," + Format$(intOrdcnt, "000") & "," + strOrder + "]"
'                f_strSend = f_strSend + f_funGet_CheckSum(f_strSend) + Chr(13) + Chr(10)
'
''                Call COM_OUTPUT(f_strSend)
'                f_strSend = ""
'            End If
'            f_intIdx = f_intIdx + 1
'        End With
'    Next
    
End Sub

Private Sub cmdWordlist_Click()
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim itemX       As ListItem
    Dim strSeqno    As String
    Dim lngRow      As Long, lngPos  As Long
    
    Dim rv          As Long
    Dim tst_no() As String
    
    CallForm = "frmComm - Private Sub cmdQuery_Click()"
    
On Error GoTo ErrorRoutine
    Me.MousePointer = 11
    
    With spdWorkList
        .maxrows = 15
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
    End With
    lngRow = 0
    
    'f_intWork_Row = 0
    
    strSeqno = mskOrdDate.Text
    rv = sl_spcid_tstcd_select("03071800007", tst_no)
    If (rv = 0) Then
    Else
         rv = dce_error("msg")
    End If
    
    
'    sqlDoc = "select SAMPLE_DATE, SAMPLE_SEQ, PART, ORD_CODE from L3A01" & _
'             " where SAMPLE_DATE = '" & mskOrdDate.Text & "'" & _
'             "   and ORD_CODE in (" & f_strOrdList & ")"
'    If cboRstgbn(0).ListIndex = 0 Then
'        sqlDoc = sqlDoc & "   and RESULT = ''"
'    ElseIf cboRstgbn(0).ListIndex = 1 Then
'        sqlDoc = sqlDoc & "   and RESULT <> ''"
'    End If
'    sqlDoc = sqlDoc & " order by 1, 2"
'    adoRS.CursorLocation = adUseClient
'    adoRS.Open sqlDoc, AdoCn_SQL
'    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
'    Do While Not adoRS.EOF
'        With spdWorkList
'            strSeqno = Mid$(adoRS(0) & "", 5, 4) + Format$(adoRS(1) & "", "000") + Trim$(adoRS(2) & "")
'
'            lngPos = f_funGet_SpreadRow(strSeqno, 2)
'            If lngPos < 1 Then
'                lngRow = lngRow + 1
'                lngPos = lngRow
'            End If
'
'            If lngPos > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
'
'            .SetText 1, lngPos, "1"
'            .SetText 2, lngPos, strSeqno
'
'            Set itemX = lvwCuData.FindItem(Trim$(adoRS(3) & ""), lvwText, , lvwWhole)
'            If Not itemX Is Nothing Then
'                .Row = lngPos:  .Col = itemX.Index + 4
'                .BackColor = &HC6FEFF
'            End If
'            Set itemX = Nothing
'        End With
'
'        adoRS.MoveNext
'    Loop
'    adoRS.Close:    Set adoRS = Nothing
    
    Me.MousePointer = 0
    Exit Sub
ErrorRoutine:
    Set itemX = Nothing
    
    Call ErrMsgProc(CallForm)
    Me.MousePointer = 0

End Sub

Private Sub comEQP_OnComm()
    Dim sStxCheck As Integer
    Dim sEnqCheck As Integer
    Dim sEtxCheck As Integer
    Dim com_sTemp As String
    Dim sLfCheck As Integer
    Dim sCrcheck As Integer
    
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim brStr       As String
    Dim fOpt        As String
'    Dim Arr()       As Byte

    Select Case comEQP.CommEvent
        Case comEvReceive

            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If
            brStr = comEQP.Input
'            Call ComReceive(Arr)
            fRcvString = fRcvString + brStr
            sStxCheck = InStr(fRcvString, Chr(2))
            sEtxCheck = InStr(fRcvString, Chr(3))
            If sStxCheck <> 0 And sEtxCheck <> 0 Then
                com_sTemp = Mid$(fRcvString, sStxCheck, sEtxCheck)
                fRcvString = Mid$(fRcvString, sEtxCheck + 1)
                'brCom.Output = Chr(6)
                If optSeq.Value = True Then
                    fOpt = "0"
                Else
                    fOpt = "1"
                End If
                Call psDataDefine(com_sTemp, fChannel(), spdWorkList, fOpt) ', brSpread, brChannel(), brItemdeci(), brOpt
            End If

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
        Select Case f_strJOB_FLAG
            Case "1"    '-- 대기
                        Select Case Asc(strBuff)
                            Case 2:     f_strBuffer = f_strBuffer + strBuff
                                        f_strJOB_FLAG = "2"     '-- 받기
                        End Select
                        
            Case "2"    '--  받기
                        Select Case Asc(strBuff)
                            Case 2      '-- STX 수신
                                        f_strBuffer = f_strBuffer + strBuff
                                        f_strJOB_FLAG = 2
                            Case 3      '-- ETX 수신
                                        f_strBuffer = f_strBuffer + strBuff
                                        Set Result1 = New clsResult
                                        With Result1
                                            .Rst_Sid = strSampleno + f_strPCFlag
                                            .Rst_Eid = Mid$(f_strBuffer, 102, 1)
                                            .Rst_Type = "G"
                                            .Rst_Test = Mid$(f_strBuffer, 60, 4)
                                            .Rst_Values = IIf(Mid$(f_strBuffer, 82, 9) = "#########", "", Mid$(f_strBuffer, 82, 9))
                                            .Rst_Tag = Mid$(f_strBuffer, 60, 4)
                                            .Rst_Error = ""
                                        End With
                                        Call Result_MsgSplit(Result1)
                                        Set Result1 = Nothing

                                        Call COM_OUTPUT(Chr(6))
                                        f_strBuffer = ""
                                        f_strJOB_FLAG = 1
                            Case Else
                                        f_strBuffer = f_strBuffer + strBuff
                        End Select
            Case "3"    '-- WorkList
                        Select Case Asc(strBuff)
                            Case 6  '-- ACK 수신시
                                    If f_lngWork_Row > f_intCnt Then
                                        f_strJOB_FLAG = 1
                                    Else
                                        Call COM_OUTPUT(f_strSample(f_lngWork_Row))
                                        f_lngWork_Row = f_lngWork_Row + 1
                                        f_strJOB_FLAG = 3
                                    End If
                                    
                            Case 21 '-- NAK 수신시
                                    Call COM_OUTPUT(f_strSample(f_lngWork_Row))
                                    f_strJOB_FLAG = 3
                        End Select
            
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
    
    f_funGet_CheckSum = Chr(intChkSum) - -Format$(Hex(intChkSum), "00")
        
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
    
    Call cmdClear         ' 초기화
    Erase fChannel
    Call f_subSet_ItemHeader          ' 리스트해더
    Call f_subGet_Setting            ' 통신설정
    'Call f_subSet_ItemList           ' 검사항목
    Call f_subSet_ComCharacter       ' 통신문자
    
    Call cmdRun           ' 실행
    
    f_intTestNo = 0
    mskRstDate.Text = Format$(Now, "YYYYMMDD")
    mskOrdDate.Text = Format$(Now, "YYYYMMDD")
    Open App.Path + "\" + "dump_job.log" For Append As #1

    f_strJOB_FLAG = "1":    f_strJOB_ACKETC = "1":  f_blnJOB_Conent = False
    f_blnWorkList = False
'    cboRstgbn(0).ListIndex = 0: cboRstgbn(1).ListIndex = 2
    
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
    fXe2100Size(27, 0) = 5:     fXe2100Size(27, 1) = 3      ' NRBC%
    fXe2100Size(28, 0) = 5:     fXe2100Size(28, 1) = 3      ' NRBC#
    
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

Private Sub lvwComplete_Click()

    Dim itemX   As ListItem
    
    Set itemX = lvwComplete.SelectedItem
    If Not itemX Is Nothing Then
        If itemX.SmallIcon = "LSE" Then
            itemX.SmallIcon = "LST"
        Else
            itemX.SmallIcon = "LSE"
        End If
    End If
    Set itemX = Nothing
    
End Sub


Private Sub lvwComplete_DblClick()

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim itemX   As ListItem:    Dim itemA   As ListItem:    Dim itemL   As ListItem
    Dim itemS   As ListSubItem: Dim itemSA  As ListSubItem
    
    Dim strSample_dt    As String, strSample_no As String, strPart      As String
    Dim strOrd_cd       As String, strChart_no  As String, strResult    As String
    Dim strValue1       As String, strValue2    As String
    
    Me.MousePointer = 11
    For Each itemX In lvwData.ListItems
        itemX.SubItems(2) = ""
        itemX.SubItems(3) = ""
        itemX.SubItems(4) = ""
        itemX.SubItems(5) = ""
    Next
    Set itemX = Nothing

    Set itemX = lvwComplete.SelectedItem
    If Not itemX Is Nothing Then
        For Each itemS In itemX.ListSubItems
            If Trim(itemS.Text) <> "" Then
                strSample_dt = Mid$(mskRstDate.Text, 1, 4) + Mid$(itemX.Text, 1, 4)
                strSample_no = Format$(Mid$(itemX.Text, 5, 3), "##0")
                strPart = Mid$(itemX.Text, 8, 2)
                strOrd_cd = lvwComplete.ColumnHeaders(itemS.Index + 1).tag
                
                Set itemL = lvwData.FindItem(Mid$(lvwComplete.ColumnHeaders(itemS.Index + 1).Key, 2), lvwTag, , lvwWhole)
                Set itemA = lvwCuData.FindItem(Mid$(lvwComplete.ColumnHeaders(itemS.Index + 1).Key, 2), lvwTag, , lvwWhole)
                If Not itemL Is Nothing Then
                    itemL.SubItems(3) = itemS.Text
                    itemL.ListSubItems(3).ForeColor = vbBlack
                    If Val(itemS.Text) < Val(itemA.SubItems(8)) Then
                        itemL.SubItems(3) = itemS.Text + " [L]"
                        itemL.ListSubItems(3).ForeColor = vbRed
                    ElseIf Val(itemS.Text) > Val(itemA.SubItems(9)) Then
                        itemL.SubItems(3) = itemS.Text + " [H]"
                        itemL.ListSubItems(3).ForeColor = vbRed
                    End If
                    sqlDoc = "   set rowcount 1 " & _
                             "select c.SAMPLE_DATE, c.SAMPLE_SEQ, c.RESULT" & _
                             "  from (select CHART_NO FROM L3A01" & _
                             "        where  SAMPLE_DATE = '" & strSample_dt & "'" & _
                             "        and    SAMPLE_SEQ  = " & strSample_no & "" & _
                             "        and    PART        = '" & strPart & "'" & _
                             "        and    ORD_CODE    = '" & strOrd_cd & "'" & _
                             "        group  by CHART_NO) as a," & _
                             "       (select b1.CHART_NO, max(b1.SAMPLE_DATE) SAMPLE_DATE from L3A01 b1," & _
                             "               (select CHART_NO from L3A01" & _
                             "                where  SAMPLE_DATE = '" & strSample_dt & "'" & _
                             "                and    SAMPLE_SEQ  =  " & strSample_no & "" & _
                             "                and    PART        = '" & strPart & "'" & _
                             "                and    ORD_CODE    = '" & strOrd_cd & "'" & _
                             "                group  by CHART_NO) as b2" & _
                             "        where  b1.SAMPLE_DATE < '" & strSample_dt & "'" & _
                             "        and    b1.ORD_CODE    = '" & strOrd_cd & "'" & _
                             "        and    b1.CHART_NO    = b2.CHART_NO" & _
                             "        group  by b1.CHART_NO) AS b,"
                    sqlDoc = sqlDoc & _
                             "       (select c1.SAMPLE_DATE, c1.SAMPLE_SEQ, c1.CHART_NO, c1.RESULT from L3A01 c1," & _
                             "               (select CHART_NO from L3A01" & _
                             "                where  SAMPLE_DATE = '" & strSample_dt & "'" & _
                             "                and    SAMPLE_SEQ  =  " & strSample_no & "" & _
                             "                and    PART        = '" & strPart & "'" & _
                             "                and    ORD_CODE    = '" & strOrd_cd & "'" & _
                             "                group  by CHART_NO) as c2" & _
                             "        where  c1.ORD_CODE  =  '" & strOrd_cd & "'" & _
                             "          and  c1.CHART_NO = c2.CHART_NO" & _
                             "        group  by c1.SAMPLE_DATE, c1.SAMPLE_SEQ, c1.CHART_NO, c1.RESULT) c" & _
                             " Where a.CHART_NO = b.CHART_NO" & _
                             "   and b.CHART_NO = c.CHART_NO" & _
                             "   and b.SAMPLE_DATE = c.SAMPLE_DATE" & _
                             " order by c.SAMPLE_DATE desc, c.SAMPLE_SEQ desc" & _
                             "   set rowcount 0"
                    adoRS.CursorLocation = adUseClient
                    adoRS.Open sqlDoc, AdoCn_SQL
                    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
                    If Not adoRS.EOF Then itemL.SubItems(2) = Trim$(adoRS(2) & ""): strResult = Trim$(adoRS(2) & "")
                    adoRS.Close:    Set adoRS = Nothing
                    If strResult <> "" Then
                        If itemA.SubItems(4) <> "" Then
                            '-- DELTA
                            strValue1 = Abs(Val(itemS.Text) - Val(strResult))
                            strValue2 = (strValue1 / Val(strResult)) * 100
                            Select Case itemA.SubItems(5)
                                Case "1":   If strValue1 > Val(itemA.SubItems(4)) Then itemL.SubItems(4) = "D"
                                Case "2":   If strValue2 > Val(itemA.SubItems(4)) Then itemL.SubItems(4) = "D"
                                Case "3":   If Val(strValue2) / 30 > Val(itemA.SubItems(4)) Then itemL.SubItems(4) = "D"
                                Case "4":   If Val(strValue1) / Val(strValue2) > Val(itemA.SubItems(4)) Then itemL.SubItems(4) = "D"
                            End Select
                            '-- PANIC
                            If Val(itemS.Text) < Val(itemA.SubItems(6)) Or Val(itemS.Text) > Val(itemA.SubItems(7)) Then itemL.SubItems(5) = "P"
                        End If
                    End If
                End If
                Set itemL = Nothing
            End If
        Next
    End If
    Set itemX = Nothing
    Me.MousePointer = 0

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
    
    Dim itemX As ListItem
    
    Set itemX = lvwComplete.FindItem(Trim(SID), lvwTag, , lvwWhole)
    If itemX Is Nothing Then
        Set itemX = lvwComplete.ListItems.Add(, , Trim(SID))
        If Not itemX Is Nothing Then
            With itemX
                .Key = COL_KEY & Trim(SID)
                .tag = Trim(SID)
                .SmallIcon = "LST"
            End With
        End If
    End If
    
End Sub

Private Sub Result_MsgSplit(ByVal Result As clsResult)

On Error GoTo ErrorRoutine
    
    Dim sqlDoc  As String, sqlRet   As Integer
    
    Dim strTime As String
    Dim itemX   As ListItem
    Dim itemH   As ListItem
    Dim itemS   As ListSubItem
    
    CallForm = "frmComm - Private Sub Result_MsgSplit()"

    '메치 테이블에서 검사코드를 가져옴
    Set itemX = lvwCuData.FindItem(Trim(Result.Rst_Test), lvwTag, , lvwWhole)
    If Not itemX Is Nothing Then
        If Mid$(Result.Rst_Sid, 10, 2) = "PC" And Trim(Result.Rst_Test) = "06A" Then
            Result.Rst_Sid = Mid$(Result.Rst_Sid, 1, 9)
            Result.Rst_Test = "XXX"
            Result.Rst_Tag = ""
        Else
            Result.Rst_Sid = Mid$(Result.Rst_Sid, 1, 9)
            Result.Rst_Tag = Trim(itemX.SubItems(1))
        End If
        
        sqlDoc = "Update INTERFACE003 set RSTVAL = '" & Result.Rst_Values & "', REFVAL = '" & Result.Rst_Eid & "'" & _
                 " where SPCNO  = '" & Result.Rst_Sid & "'" & _
                 "   and TESTCD = '" & Result.Rst_Test & "'" & _
                 "   and TRANSDT = '" & Format$(Now, "YYYYMMDD") & "'" & _
                 "   and TRANSTM = '" & Format$(Now, "MMSS") & "'"
        AdoCn_Jet.Execute sqlDoc, sqlRet
        If sqlRet = 0 Then
            sqlDoc = "insert into INTERFACE003(" & _
                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD)" & _
                     "    values( '" & Result.Rst_Sid & "', '" & Result.Rst_Test & "'," & _
                     "            '" & Result.Rst_Eid & "', '" & Format$(Now, "YYYYMMDD") & "'," & _
                     "            '" & Format$(Now, "MMSS") & "', '" & Result.Rst_Values & "'," & _
                     "            '" & Result.Rst_Eid & "', '" & INS_CODE & "')"
            AdoCn_Jet.Execute sqlDoc
        End If
        
        '결과 표시
        Set itemH = lvwComplete.FindItem(Result.Rst_Sid, lvwText, , lvwWhole)
        If itemH Is Nothing Then
            Set itemH = lvwComplete.ListItems.Add()
            With itemH
                .Key = COL_KEY & Result.Rst_Sid '아이템 키에 검체번호
                .Text = Result.Rst_Sid          '아이템 에 검체번호
                .tag = Result.Rst_Type          '테그에 결과 타입
                .SmallIcon = "LSE"
            End With
        End If
        '결과값 등록
        itemH.SubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex) = Result.Rst_Values
        
        '--- 판정
        itemH.ListSubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex).ForeColor = vbBlack
        If Val(itemX.SubItems(7)) < Val(Result.Rst_Values) Or Val(itemX.SubItems(8)) > Val(Result.Rst_Values) Then
            itemH.ListSubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex).ForeColor = vbRed
        End If
        
        Set itemS = itemH.ListSubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex)
        
        itemS.tag = Result.Rst_Error '서브아이템 테그에 에러 메시지
        Set itemS = Nothing
        Set itemX = Nothing
        Set itemX = Nothing
    End If
    '검사코드가 없는것은 등록 하지 않음
    Exit Sub
ErrorRoutine:

    Set itemS = Nothing
    Set itemX = Nothing
    Set itemX = Nothing
    
    Call ErrMsgProc(CallForm)
    Err.Clear
    
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

Private Sub psDataDefine(ByVal brbarcd As String, ByRef brChannel() As String, ByVal brSpread As Object, ByVal brOst As String) ' ByRef brItemdeci() As String)

Dim sTemp      As String
Dim Channel_No As Integer       ' 검사항목 번호 : Channel No
Dim Patiant_No As String
Dim sDeCnt     As Integer
Dim pStart_Point, pEnd_Point As Integer
Dim pGrid_Point As Integer
Dim pDoCount   As Integer
Dim Loop_Count As Integer
Dim FunStr1 As String, FunStr2 As String, FunStr3 As String, FunStr4 As String
Dim Max_Arary_Cnt As Integer    ' 검사 항목수
Dim ssTemp1 As Integer, ssTemp2 As String, sAdd As Integer, sPosition As Integer
    
    On Error GoTo errDefine
    sTemp = brbarcd
    '------------------------------<<< fXE2100() 배열 Clear 한다.         >>>----------
    For Loop_Count = 1 To 100: fXE2100(Loop_Count) = "": Next Loop_Count
    '------------------------------<<< fXE2100() 배열에 구분하여 넣는다.  >>>----------
        
    If Mid$(sTemp, 2, 3) = "D1U" Then
    
        '  order sending
        
    ElseIf Mid$(sTemp, 2, 3) = "D2U" Then
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
                        fXE2100(11) = Mid(sTemp, 90, fXe2100Size(11, 0))
                        fXE2100(12) = Mid(sTemp, 95, fXe2100Size(12, 0))
                        fXE2100(13) = Mid(sTemp, 100, fXe2100Size(13, 0))
                        fXE2100(14) = Mid(sTemp, 105, fXe2100Size(14, 0))
                        fXE2100(15) = Mid(sTemp, 110, fXe2100Size(15, 0))
                        fXE2100(16) = Mid(sTemp, 115, fXe2100Size(16, 0))
                        fXE2100(17) = Mid(sTemp, 121, fXe2100Size(17, 0))
                        fXE2100(18) = Mid(sTemp, 127, fXe2100Size(18, 0))
                        fXE2100(19) = Mid(sTemp, 133, fXe2100Size(19, 0))
                        fXE2100(20) = Mid(sTemp, 139, fXe2100Size(20, 0))
                        fXE2100(21) = Mid(sTemp, 145, fXe2100Size(21, 0))
                        fXE2100(22) = Mid(sTemp, 150, fXe2100Size(22, 0))
                        fXE2100(23) = Mid(sTemp, 155, fXe2100Size(23, 0))
                        fXE2100(24) = Mid(sTemp, 160, fXe2100Size(24, 0))
                        fXE2100(25) = Mid(sTemp, 165, fXe2100Size(25, 0))
                        fXE2100(26) = Mid(sTemp, 200, fXe2100Size(26, 0))
                        If Len(sTemp) = 244 Then
                            fXE2100(27) = ""
                            fXE2100(28) = ""
                        Else
                            fXE2100(27) = Mid(sTemp, 205, fXe2100Size(27, 0))
                            fXE2100(28) = Mid(sTemp, 301, fXe2100Size(28, 0))
                        End If
            Case Else
        End Select
        
        '-------------------------------------------<<< 해당검사결과와 해당환자를 찿는다.       >>>----------

        Max_Arary_Cnt = brSpread.MaxCols - 6   ' 앞에서부터 5까지는 환자 정보 이기때문에.... -5를 한다.
                                               ' 해당 배열은  brItem(),brChannel() 이다.
        pGrid_Point = 0
        Dim sSEq As Long
        Dim sCol As Integer
        
        With brSpread
            If brOst = 1 Then
                sSEq = Val(fXE2100(1))
                sCol = 0
            Else
                sSEq = Val(fXE2100(2))
                sCol = 2
                'Call sclsinfo.EmpList(Trim(fXE2100(2)), fMachCd)       ' 화면에 수검자 등록
            End If
            pGrid_Point = SeqSearch(brSpread, sSEq, sCol)
            
            If pGrid_Point > 0 Then                ' 해당 대상자를 찿으면 ....
                For pDoCount = 1 To Max_Arary_Cnt    '-------------------------------<<<<<<<<<,  세부검사항목을 찿는다.  >>>>>>>---------
                    .Row = pGrid_Point
                    .Col = pDoCount + 2
                    Channel_No = Val(brChannel(pDoCount))               '  Channel이 숫자이기 때문에 숫자로 치환한다.
                    If Len(fXE2100(Channel_No)) > 0 Then
                        If fXe2100Size(Channel_No, 1) = fXe2100Size(Channel_No, 0) Then
                            FunStr3 = Trim(Val(fXE2100(Channel_No)))
                        Else
                            FunStr3 = Trim(Val(Mid$(fXE2100(Channel_No), 1, fXe2100Size(Channel_No, 1)))) + "." + _
                                           Mid$(fXE2100(Channel_No), fXe2100Size(Channel_No, 1) + 1)
                        End If
                        If IsNumeric(FunStr3) Then
                            .Text = FunStr3
                        Else
                            .Text = ""
                        End If
                    End If
                Next pDoCount
                '-----------------------------------------------------------------------
            End If
        End With
    End If
    Exit Sub
    
errDefine:

End Sub

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

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'         1         2         3         4         5         6         7         8         9         10        1         2         3         4         5         6         7         8         9         20        1         2         3         4
'123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789  123456789
'                    SAMPLE SEQ
'                              RESERVED
'                                 SAMPLE ID
'                                                WBC
'                                                     RBC
'D1U   XE-2100^A10930000000073000             830107231544000000090302                0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000073000             8300461004120013800400009710033500345001950         011900424001250010700304000000000000000000000000000000000210           00000000000000000XE-2100^99337319^A1093
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
'         1         2         3         4         5         6         7         8         9         10        1         2         3         4         5         6         7         8         9         20        1         2         3         4
'123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789  123456789

'D2U   XE-2100^A10930000000053000             7900392004200013700394009380032600348001850047440051404394003300003000186400020400172400013000001001200040300109000980022000000000000000000000000000000000018000000000000000000XE-2100^99337319^A1093
'000000000000000000000000000000
'D2U   XE-2100^A10930000000053000             79003920042000137003940093800326003480018500474400514043940033000030001864000204001724000130000010012000403001090009800220000000000000000000000000000000000180           00000000000000000XE-2100^99337319^A1093



Private Sub txtBarCode_Change()

End Sub

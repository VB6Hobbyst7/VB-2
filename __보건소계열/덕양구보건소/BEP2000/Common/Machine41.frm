VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form INTface41 
   BorderStyle     =   0  '없음
   Caption         =   "해당일의 검사결과 받아 보기"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   1200
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7500
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame fraWork 
      Height          =   6075
      Left            =   0
      TabIndex        =   16
      Top             =   1035
      Width           =   11805
      _Version        =   65536
      _ExtentX        =   20823
      _ExtentY        =   10716
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spdList 
         Height          =   4710
         Left            =   60
         TabIndex        =   3
         Top             =   1200
         Width           =   4005
         _Version        =   196608
         _ExtentX        =   7064
         _ExtentY        =   8308
         _StockProps     =   64
         BackColorStyle  =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         MaxCols         =   3
         MaxRows         =   0
         ScrollBars      =   2
         SelectBlockOptions=   6
         SpreadDesigner  =   "Machine41.frx":0000
         UserResize      =   0
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin FPSpread.vaSpread spdWorkList 
         Height          =   5730
         Left            =   6150
         TabIndex        =   23
         Top             =   210
         Width           =   5505
         _Version        =   196608
         _ExtentX        =   9710
         _ExtentY        =   10107
         _StockProps     =   64
         BackColorStyle  =   1
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         MaxCols         =   4
         MaxRows         =   21
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "Machine41.frx":10FC
         UserResize      =   0
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   975
         Left            =   60
         TabIndex        =   17
         Top             =   180
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7064
         _ExtentY        =   1720
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
         BevelInner      =   2
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboSelect 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1110
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   555
            Width           =   1605
         End
         Begin MSMask.MaskEdBox mskDate 
            Height          =   345
            Index           =   0
            Left            =   1110
            TabIndex        =   0
            Top             =   135
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-mm-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   390
            Left            =   60
            TabIndex        =   18
            Top             =   120
            Width           =   1020
            _Version        =   65536
            _ExtentX        =   1799
            _ExtentY        =   688
            _StockProps     =   15
            Caption         =   "접수일자"
            ForeColor       =   65535
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   390
            Left            =   60
            TabIndex        =   19
            Top             =   510
            Width           =   1020
            _Version        =   65536
            _ExtentX        =   1799
            _ExtentY        =   688
            _StockProps     =   15
            Caption         =   "조회조건"
            ForeColor       =   65535
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            RoundedCorners  =   0   'False
            MouseIcon       =   "Machine41.frx":1537
         End
         Begin MSMask.MaskEdBox mskDate 
            Height          =   345
            Index           =   1
            Left            =   2610
            TabIndex        =   1
            Top             =   150
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-mm-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2460
            TabIndex        =   26
            Top             =   210
            Width           =   105
         End
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   930
         Left            =   4170
         TabIndex        =   20
         Top             =   2550
         Width           =   1725
         _Version        =   65536
         _ExtentX        =   3043
         _ExtentY        =   1640
         _StockProps     =   78
         Caption         =   "Work List 작성"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
         RoundedCorners  =   0   'False
         Picture         =   "Machine41.frx":1E11
      End
      Begin Threed.SSCommand cmdOrder 
         Height          =   870
         Left            =   5085
         TabIndex        =   4
         Top             =   1620
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "전송"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
         Picture         =   "Machine41.frx":1E2D
      End
      Begin Threed.SSCommand cmdQuery 
         Height          =   870
         Left            =   4170
         TabIndex        =   25
         Top             =   1620
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "자료"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
         RoundedCorners  =   0   'False
         Picture         =   "Machine41.frx":227F
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   390
         Left            =   4140
         TabIndex        =   29
         Top             =   300
         Width           =   1080
         _Version        =   65536
         _ExtentX        =   1905
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "Order No"
         ForeColor       =   65535
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         RoundedCorners  =   0   'False
         MouseIcon       =   "Machine41.frx":2B59
      End
      Begin MSMask.MaskEdBox mskOrderno 
         Height          =   345
         Left            =   5250
         TabIndex        =   30
         Top             =   315
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-mm-dd"
         Mask            =   "###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskRackno 
         Height          =   345
         Left            =   5250
         TabIndex        =   31
         Top             =   735
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   390
         Left            =   4140
         TabIndex        =   32
         Top             =   720
         Width           =   1080
         _Version        =   65536
         _ExtentX        =   1905
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "Rack No"
         ForeColor       =   65535
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   390
         Left            =   4140
         TabIndex        =   33
         Top             =   1140
         Width           =   1080
         _Version        =   65536
         _ExtentX        =   1905
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "Position"
         ForeColor       =   65535
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         RoundedCorners  =   0   'False
         MouseIcon       =   "Machine41.frx":3433
      End
      Begin MSMask.MaskEdBox mskPosition 
         Height          =   345
         Left            =   5250
         TabIndex        =   34
         Top             =   1155
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Label2"
         Height          =   5775
         Left            =   6000
         TabIndex        =   24
         Top             =   180
         Width           =   75
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5310
      Top             =   7110
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   555
      Left            =   4020
      TabIndex        =   35
      Top             =   -15
      Width           =   2325
      _Version        =   65536
      _ExtentX        =   4101
      _ExtentY        =   979
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSMask.MaskEdBox mskOrdDate 
         Height          =   315
         Left            =   930
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-mm-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "검사일자"
         Height          =   180
         Left            =   90
         TabIndex        =   36
         Top             =   240
         Width           =   810
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   12000
      Left            =   4830
      Top             =   7140
   End
   Begin Threed.SSFrame fmeSlipSeq 
      Height          =   555
      Left            =   4020
      TabIndex        =   7
      Top             =   480
      Width           =   2325
      _Version        =   65536
      _ExtentX        =   4101
      _ExtentY        =   979
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txsapno 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   930
         TabIndex        =   8
         Top             =   150
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "일련번호"
         Height          =   180
         Left            =   60
         TabIndex        =   9
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   7950
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1035
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4005
      _Version        =   65536
      _ExtentX        =   7064
      _ExtentY        =   1826
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "검사 인터페이스 작업을 수행합니다!!"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   225
         TabIndex        =   12
         Top             =   675
         Width           =   3675
      End
      Begin VB.Label LblMMDD 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   11
         Top             =   225
         Width           =   1080
      End
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   870
      Left            =   10080
      TabIndex        =   15
      Top             =   150
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "삭   제"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Font3D          =   2
      Picture         =   "Machine41.frx":3D0D
   End
   Begin MSCommLib.MSComm Comm1 
      Left            =   4245
      Top             =   6795
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   512
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
   End
   Begin Threed.SSCommand cmdclose 
      Height          =   870
      Left            =   10920
      TabIndex        =   5
      Top             =   150
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "종   료"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Font3D          =   2
      Picture         =   "Machine41.frx":56AF
   End
   Begin Threed.SSFrame fmeGuide 
      Height          =   525
      Left            =   6405
      TabIndex        =   13
      Top             =   0
      Width           =   1995
      _Version        =   65536
      _ExtentX        =   3519
      _ExtentY        =   926
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtguide 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   135
         Width           =   1845
      End
   End
   Begin Threed.SSCommand cmdInitial 
      Height          =   870
      Left            =   9240
      TabIndex        =   22
      Top             =   150
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "초기화"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Font3D          =   2
      Picture         =   "Machine41.frx":7051
   End
   Begin Threed.SSCommand cmdDown 
      Height          =   870
      Left            =   8385
      TabIndex        =   28
      Top             =   150
      Visible         =   0   'False
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "받기"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Font3D          =   2
      Picture         =   "Machine41.frx":74A3
   End
   Begin VB.TextBox txtStatus 
      Height          =   420
      Left            =   6405
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   27
      Top             =   570
      Visible         =   0   'False
      Width           =   2835
   End
   Begin FPSpread.vaSpread spdface 
      Height          =   5985
      Left            =   15
      TabIndex        =   6
      Top             =   1110
      Width           =   11775
      _Version        =   196608
      _ExtentX        =   20770
      _ExtentY        =   10557
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      ArrowsExitEditMode=   -1  'True
      BackColorStyle  =   1
      ColsFrozen      =   1
      EditEnterAction =   2
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      MaxRows         =   14
      SelectBlockOptions=   2
      SpreadDesigner  =   "Machine41.frx":78F5
      UserResize      =   0
      VisibleCols     =   500
      VisibleRows     =   500
   End
End
Attribute VB_Name = "INTface41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private transcfg        As commset
Dim errfound            As Integer
Dim Porttag             As Integer
Dim RcvBuffer           As String
Dim TestNameTable(30)   As TestNameTbl
Dim nowcount            As Integer

Dim EmerCnt             As Integer
Dim SampleCnt           As Integer
Dim EditCnt             As Integer  'Express처럼 한 슬립이 두 프레임 이상 거쳐 나올 경우 사용
Dim PrevCnt             As Integer
Dim CurSampCnt          As Integer
Dim Startslip           As String

Dim StartBCol           As Integer
Dim EndBCol             As Integer
Dim StartBRow           As Integer
Dim EndBRow             As Integer
Dim identbOpenKey       As Integer
Dim IdList              As String
Dim BlockKey            As Integer
Dim Errkey              As Integer

'public에 slipno As String 선언
Dim phase               As Integer
Dim bufcnt              As Integer
Dim wkbuf               As String
Dim ix1                 As Integer
Dim PrevReq             As Integer
Dim OldRow              As Long
Dim long_slip1          As Long
Dim long_slip2          As Long

Dim Test_OpenFlag       As Integer
Dim F_iComm_Cnt         As Integer  '99.11.23 YEJ

Dim f_iWork_Row         As Integer
Dim f_adoCn             As ADODB.Connection

Private Type TYPE_TESTID
    sPatid      As String
    sPatnm      As String
    sOrderId    As String
    sRackno     As String
    sPosition   As String
    sTestID(1 To 100)   As String
    iTestCnt            As Integer
End Type
Dim f_tpTestList()  As TYPE_TESTID
Dim f_iTestCnt     As Integer

Dim AxOrder(5)     As String
Dim AxResult(13)   As String
Dim tmpOrder       As String
Sub f_subGet_검사항목(ByVal sKeyno As String, _
                      ByRef sTestID() As String, ByRef iCnt As Integer)

    Dim adoRs   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim labdate As String, numgbn As String, labsqno As String
    Dim iIdx    As Integer
    
    If InStr(sKeyno, "-") > 0 Then
        labdate = Mid$(sKeyno, 1, 8)
        numgbn = Mid$(sKeyno, 10, 1)
        labsqno = Mid$(sKeyno, 12, 5)
    Else
        labdate = Mid$(sKeyno, 1, 8)
        numgbn = Mid$(sKeyno, 9, 1)
        labsqno = Mid$(sKeyno, 10, 5)
    End If
    
    sqlDoc = "select SLIPCD+ORDCD+SPCCD from LAB_DB..LAB030M" _
           & " where LABDATE = '" & labdate & "'" _
           & "   and NUMGBN  = '" & numgbn & "'" _
           & "   and LABSQNO = '" & labsqno & "'" _
           & "   and SUBCD   = ''" _
           & "   and SLIPCD + ORDCD + SPCCD in ("
           
    sqlDoc = sqlDoc + "''"
    
    Set dbcode = OpenDatabase(filename & codestr)
    Set tbcode = dbcode.OpenRecordset("cdtable")
    
    tbcode.MoveLast
    If tbcode.RecordCount > 0 Then tbcode.MoveFirst
    
    Do While Not tbcode.EOF
        sqlDoc = sqlDoc + ",'" + tbcode!code & "" & "'"
        tbcode.MoveNext
    Loop
    tbcode.Close:   dbcode.Close

    sqlDoc = sqlDoc + ")"
    
    adoRs.CursorLocation = adUseClient
    adoRs.Open sqlDoc, f_adoCn, adOpenStatic, adLockReadOnly
    
    If adoRs.RecordCount > 0 Then adoRs.MoveFirst
    
    iCnt = 0
    Do While Not adoRs.EOF
    
        For iIdx = 1 To 30
            If InStr(TestNameTable(iIdx).code, Trim$(adoRs(0) & "")) > 0 Then
                iCnt = iCnt + 1
                sTestID(iCnt) = Trim(TestNameTable(iIdx).eqno)
            End If
        Next
        
        adoRs.MoveNext
    Loop
    adoRs.Close:    Set adoRs = Nothing
    
End Sub

'
'   참고치 판정
'
Private Function Chk_Ref(sOrdCd As String, sSubNo As String, sRes As String, _
                        sex As String) As String

    Dim sStr    As String
    Dim sData() As String
    Dim iRet_Cd As Integer
    
    Dim LowVal  As Single
    Dim HighVal As Single
    Dim RefVal  As Single
    Dim RefChar As String

    Chk_Ref = ""

    If sex = "M" Then
        sStr = " Select REFLOM, REFHIM, REFCHAR, REFCHK "
    ElseIf sex = "F" Then
        sStr = " Select REFLOF, REFHIF, REFCHAR, REFCHK "
    End If
    
    sStr = sStr & "  From LAB01_DB..DJA060M " _
            & " where ORDCD = '" & sOrdCd & "'" _
            & "   and SUBCD = '" & sSubNo & "'"
    
    If QSqlDBExec(sStr, QsqlConn) = QSQL_SUCCESS Then
        If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
            QSqlGetField 4, sStr, sData()
            
            If Trim(sData(4)) = "C" Then    '참고치 문자
                RefChar = Trim(sData(3))
                If RefChar <> Trim(sRes) Then
                    Chk_Ref = "*"
                End If
            ElseIf Trim(sData(4)) = "N" Then        '숫자
                RefVal = CSng(Val(Trim(sRes)))
                LowVal = CSng(sData(1)): HighVal = CSng(sData(2))
            
                If RefVal > HighVal Then
                    Chk_Ref = "H"
                ElseIf RefVal < LowVal Then
                    Chk_Ref = "L"
                End If
            End If
            
        End If
    End If
    iRet_Cd = QSqlSelectFree(QsqlConn)
    
End Function


Sub f_subClear_Form()

    f_iWork_Row = 0
    
    mskOrderno.Text = 1
    mskRackno.Text = "A"
    mskPosition.Text = 1
    
    txtStatus.Text = ""
    txtStatus.Visible = False
    fraWork.Visible = True
    
    spdList.MaxRows = 0
    
    With spdWorkList
        .MaxRows = 21
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = 21
        .BlockMode = True
        .Action = SS_ACTION_CLEAR_TEXT
        .BlockMode = False
    End With
    
    With spdface
        .MaxRows = 14
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = 14
        .BlockMode = True
        .Action = SS_ACTION_CLEAR_TEXT
        .BlockMode = False
    End With
    
End Sub

Private Sub Update_DJC020M(OrdSqNo As String)

    Dim iRet_Cd As Integer
    Dim sStr    As String
    Dim tData() As String
    Dim sqlDoc  As String
    
    sStr = " Select count(*) from LAB03_DB..DJC050M " _
            & " where ORDDATE = '" & Left(OrdSqNo, 8) & "'" _
            & "   and DEPTCD = '" & Mid(OrdSqNo, 10, 2) & "'" _
            & "   and SEQNO = '" & Right(OrdSqNo, 5) & "'" _
            & "   and RSTGBN = '' "
    
    If QSqlDBExec(sStr, QsqlConn) = QSQL_SUCCESS Then
        If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
            QSqlGetField 1, sStr, tData()
            
            If Val(tData(1)) = 0 Then
                sqlDoc = " Update LAB03_DB..DJC020M " _
                        & "   set ORDSTAT = '1' " _
                        & " where ORDDATE = '" & Left(OrdSqNo, 8) & "'" _
                        & "   and DEPTCD = '" & Mid(OrdSqNo, 10, 2) & "'" _
                        & "   and SEQNO = '" & Right(OrdSqNo, 5) & "'" _
                        & "   AND ORDSTAT IN ('0','6') "
                iRet_Cd = QSqlDBExec(sqlDoc, QsqlCode)
            End If
        End If
    End If
    iRet_Cd = QSqlSelectFree(QsqlConn)
                
End Sub


Private Function Append_To_Server(P_Key As String, iCnt As Integer, sOrdNo As String, RtnCd As String) As Integer
                
    Dim iRet    As Integer
    Dim sStr    As String
    Dim rData() As String
    Dim sLabNo    As String
    Dim II      As Integer
    Dim sqlDoc  As String
    
    'sLabNo = Left(P_Key, 8) & Mid(P_Key, 10, 1) & Right(P_Key, 5)
    sLabNo = P_Key
    
    Append_To_Server = True
    
    '----- Server결과등록
    With Insert_Server(iCnt)
        '--- 검사 Order 내역 Table Update
        sStr = " Update LAB03_DB..DJC050M " _
                & "   set RSTGBN = 'Y' " _
                & " where ORDDATE = '" & Left(sOrdNo, 8) & "'" _
                & "   and DEPTCD = '" & Mid(sOrdNo, 9, 2) & "'" _
                & "   and SEQNO = '" & Right(sOrdNo, 5) & "'"
        If Mid(RtnCd, 4, 1) = "0" Then
            sStr = sStr & "   and ORDCD = '" & RtnCd & "'"
        Else
            sStr = sStr & "   and ORDCD = '" & .ordcd & "'"
        End If
        
        If QSqlDBExec(sStr, QsqlConn) = QSQL_SUCCESS Then
            '--- Update
            sStr = " Update LAB04_DB..DJD010M " _
                    & "   set RSTVAL = '" & .Result & "', " _
                    & "       REFVAL = '" & .Ref & "', " _
                    & "       RSTID  = '" & D0COM_USERID & "', " _
                    & "       RSTDATE = '" & Format(Now, "YYYYMMDD") & "'" _
                    & " where LABDATE = '" & Left(sLabNo, 8) & "'" _
                    & "   and NUMGBN  = '" & Mid(sLabNo, 9, 1) & "'" _
                    & "   and LABSQNO = '" & Right(sLabNo, 5) & "'" _
                    & "   and ORDCD = '" & .ordcd & "'" _
                    & "   and SUBCD = '" & .SubNo & "'" _
                    & "   and ORDERNO = '" & sOrdNo & "'"
            
            If QSqlDBExec(sStr, QsqlConn) <> QSQL_SUCCESS Then
                '--- Insert(Sub 검사항목인 경우-조회 후 입력처리)
                ReDim INSDATA(1 To 5) As String
                
                '--- Insert할 항목 조회
                sqlDoc = " Select DISTINCT REQGBN, SPCGBN, RETGBN, RTNCD, IDNO " _
                        & "  from LAB04_DB..DJD010M " _
                        & " where LABDATE = '" & Left(sLabNo, 8) & "'" _
                        & "   and NUMGBN = '" & Mid(sLabNo, 9, 1) & "'" _
                        & "   and LABSQNO = '" & Right(sLabNo, 5) & "'" _
                        & "   and ORDCD = '" & .ordcd & "'" _
                        & "   and ORDERNO = '" & sOrdNo & "'"
                        
                If QSqlDBExec(sqlDoc, QsqlConn) = QSQL_SUCCESS Then
                    If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
                        QSqlGetField 5, sStr, rData()
                        
                        For II = 1 To 5
                            INSDATA(II) = Trim(rData(II))
                        Next II
                    Else
                        iRet = QSqlSelectFree(QsqlConn)
    '                    Append_To_Server = False
                        Exit Function
                    End If
                Else
                    iRet = QSqlSelectFree(QsqlConn)
    '                Append_To_Server = False
                    Exit Function
                End If
                iRet = QSqlSelectFree(QsqlConn)
                '--- 조회된 자료로 Insert처리
                sStr = " Insert into LAB04_DB..DJD010M ( " _
                        & " LABDATE, NUMGBN,  LABSQNO, ORDCD,  SUBCD, " _
                        & " RSTVAL,  REFVAL,  PANVAL,  DELVAL, REQGBN, " _
                        & " SPCGBN,  RETGBN,  RTNCD,   IDNO,   RSTID, " _
                        & " RSTDATE, ORDERNO, SYSDATE, SYSTIME ) values ( " _
                        & "'" & Left(sLabNo, 8) & "', " _
                        & "'" & Mid(sLabNo, 9, 1) & "', " _
                        & "'" & Right(sLabNo, 5) & "', " _
                        & "'" & .ordcd & "', " _
                        & "'" & .SubNo & "', " _
                        & "'" & .Result & "', " _
                        & "'" & .Ref & "', " _
                        & "'', '',"
                For II = 1 To 5
                    sStr = sStr & "'" & INSDATA(II) & "', "
                Next II
                sStr = sStr & "'" & D0COM_USERID & "', " _
                        & "'" & Format(Now, "YYYYMMDD") & "', " _
                        & "'" & sOrdNo & "', " _
                        & "'" & Format(Now, "YYYYMMDD") & "', " _
                        & "'" & Format(Now, "HHMMSS") & "') "
                        
                If QSqlDBExec(sStr, QsqlConn) <> QSQL_SUCCESS Then
                    Append_To_Server = False
                    Exit Function
                End If
                '---------------
            End If
        End If
        
        '--- 초기화
        .ordcd = ""
        .SubNo = ""
        .Result = ""
        .Ref = ""
        '-----------
    End With
    
End Function


Sub add_db_identb(sample As String, slip As String)
   
   If RecordExist(identb, "PrimaryKey", sample) Then
      identb.Edit
   Else
      identb.AddNew
      identb!seq_no = sample
   End If
   
   identb!slip_no = slip
   identb.Update
   identb.MoveLast

End Sub

Sub add_db_resulttb(sample2 As String, tcd As String, trt As String)
   
   tcd = Right$(tcd, 2)
   If RecordExists(resulttb, "PrimaryKey", sample2, tcd) Then
      resulttb.Edit
   Else
      resulttb.AddNew
      resulttb!seq_no = sample2
      resulttb!TestCode = tcd
   End If
   resulttb!TestResult = trt
   resulttb.Update
   resulttb.MoveLast

End Sub

Function FindIdList(position As Integer) As Integer
     
     FindIdList = InStr(position + 1, RcvBuffer, Chr(13))
     IdList = Mid$(RcvBuffer, FindIdList + 1, 1)

End Function

Sub PhaseCfg_Protocol()
    
    Dim wkdat          As String
    Dim ix1            As Integer
    Dim ir             As Integer
    
    Erase AxResult
    Erase AxOrder
    
    wkbuf = txtStatus.Text
    tmpOrder = ""
'    SampleCnt = 0
    
    For ix1 = 1 To Len(wkbuf)
        wkdat = Mid$(wkbuf, ix1, 1)
        If Trim(wkdat) = "" Then Exit For
        Select Case Asc(wkdat)
        Case 5 'ENQ
        
        Case 21 'NAK
        
        Case 2  'STX
            phase = 2
            
            Select Case Mid$(wkbuf, ix1 + 2, 1)
'                Case "H" '-- Header
'                Case "P" '-- Patient
'                Case "C" '-- Comment
'                Case "Q" '-- Request
'                Case "L" '-- Terminator
                Case "O" '-- Order
                    'If InStr(wkbuf, AxOrder(1)) <> 0 Then
                    
                    For ir = 1 To 4
                        AxOrder(ir) = GetByOne(Mid(wkbuf, ix1, 100), wkbuf)
                    Next

                    For ir = 1 To 3
                        AxOrder(ir) = GetByOne1(AxOrder(4), AxOrder(4))
                    Next
                    
                    If tmpOrder = "" Then
                        tmpOrder = AxOrder(1)
                        SampleCnt = SampleCnt + 1
                    Else
                        If tmpOrder <> AxOrder(1) Then
                            tmpOrder = AxOrder(1)
                            SampleCnt = SampleCnt + 1
                        End If
                    End If
                    'Order    번호 : AxOrder(1) ex) 200109250002
                    'Lec      번호 : AxOrder(2) ex) E
                    'Lec별일련번호 : AxOrder(4) ex)
                Case "R" '-- Result
                    For ir = 1 To 12
                        AxResult(ir) = GetByOne(Mid(wkbuf, ix1, 100), wkbuf)
                    Next
                    '결과순서 : AxResult(2)
                    '검 사 명 : AxResult(3)
                    '결    과 : AxResult(4)
                    '단    위 : AxResult(5)
                    
                    '-- 2001.11.10 판정치가 있는경우 판정치를 넣어준다.
                    '-- 1|...........^^F|결과값 ==> 결과치
                    '-- 2|...........^^P|참고값 ==> 참고치
                    '-- 3|...........^^I|판  정 ==> 판정치(NONREACTIVE & NEGATIVE & POSITIVE)
                    If AxResult(2) = 1 And Right(Trim(AxResult(3)), 1) = "F" Then
                        AxResult(3) = Mid(AxResult(3), 4, Len(AxResult(3)) - 4)
                        For ir = 1 To 3
                            AxResult(ir) = GetByOne1(AxResult(3), AxResult(3))
                        Next
                        '검사명 : AxResult(2)
                        Call edit_data
                    
                    ElseIf AxResult(2) = 3 And Right(Trim(AxResult(3)), 1) = "I" Then
                            AxResult(3) = Mid(AxResult(3), 4, Len(AxResult(3)) - 4)
                            For ir = 1 To 3
                                AxResult(ir) = GetByOne1(AxResult(3), AxResult(3))
                            Next
                            '검사명 : AxResult(2)
                            
                            Call edit_data
                        
                        'End If
                    End If
                Case Else
                    'Exit Sub
            End Select
        Case Else
        
        End Select
                   
    Next
             
    txtguide.Text = "Data 전송 완료!!"

End Sub

Private Function RecordExist(Tb As Recordset, IndexName As String, Samp As String) As Integer
         
         Dim CurrRecord As Variant

         If Tb.RecordCount < 1 Or Tb.BOF Or Tb.EOF Then
            RecordExist = False
            Exit Function
         End If

         '''CurrRecord = Tb.Bookmark
         Tb.MoveFirst
         Tb.Index = IndexName
         Tb.Seek "=", Samp

         If Tb.NoMatch Then
            '''Tb.Bookmark = CurrRecord
            RecordExist = False
         Else
            RecordExist = True
         End If
         
End Function
Private Function RecordExists(Tb As Recordset, IndexName As String, samp2 As String, tcd2 As String) As Integer
         Dim CurrRecord As Variant

         If Tb.RecordCount < 1 Then
            RecordExists = False
            Exit Function
         End If

         '''CurrRecord = Tb.Bookmark
         Tb.MoveFirst
         Tb.Index = IndexName
         Tb.Seek "=", samp2, tcd2

         If Tb.NoMatch Then
            '''Tb.Bookmark = CurrRecord
            RecordExists = False
         Else
             RecordExists = True
         End If

End Function
Sub edit_data()
        
    Dim seqno       As String
    Dim tresult(1 To 30) As String
    Dim tcode       As String
    Dim i           As Integer
    Dim a           As Integer
    Dim ix1         As Integer
    Dim pos         As Integer
    Dim tmpbuff     As String
    Dim NextPos     As Integer
    Dim tmpbuffer   As String
    Dim StartPos    As Integer
    Dim temp
    Dim no_tmp, no_tmp1, no_tmp2
    Dim iC          As Integer
    Dim chk_Id      As Boolean
    Dim tmp_iC      As Integer
    
'    SampleCnt = SampleCnt + 1
    
    '-- sampleNo.얻기->slipno에 해당
    Call spdface.GetText(1, SampleCnt, no_tmp)
    Call spdWorkList.GetText(3, SampleCnt, no_tmp1)
    Call spdWorkList.GetText(4, SampleCnt, no_tmp2)

    If Trim(no_tmp) = "" Then
'        MsgBox "Work List를 먼저 등록 하십시요", vbInformation, "Work List 등록"
        Exit Sub
    End If

'---검사결과값 얻기 ------------------------------------------------------------
    Erase tresult
    chk_Id = False
    
    For iC = 1 To spdWorkList.MaxRows
        Call spdWorkList.GetText(3, iC, no_tmp1)
        Call spdWorkList.GetText(4, iC, no_tmp2)
        
        If Trim(no_tmp1) = "" And Trim(no_tmp2) = "" Then Exit For
        
        If AxOrder(2) = no_tmp1 And AxOrder(4) = Format(no_tmp2, "00") Then
            tmp_iC = SampleCnt
            SampleCnt = iC
            chk_Id = True
            Exit For
        Else
'            If AxOrder(2) = no_tmp1 And AxOrder(4) = no_tmp2 Then
'                SampleCnt = iC
'                chk_Id = False
'            End If
        End If
    Next
    
    If chk_Id = False Then Exit Sub
    
    For ix1 = 1 To 30
        If Trim(TestNameTable(ix1).Name) = "" Then Exit For
        If InStr(AxResult(2), Trim(TestNameTable(ix1).Name)) <> 0 Then
            If Not IsNumeric(AxResult(4)) Then
                If UCase(Left(Trim(AxResult(4)), 1)) = "N" Then
                    tresult(ix1) = "음성"
                Else
                    tresult(ix1) = "양성"
                End If
            Else
                tresult(ix1) = AxResult(4)
            End If
            
            Exit For
        End If
    Next ix1

'---검사명, 검사결과값 spread에 뿌리기------------------------------------------------------------
'    txslipno.Text = slipno
'    txsapno.Text = Format(PrevCnt + SampleCnt, "0000")
    txsapno.Text = Format(SampleCnt, "0000")
    txtguide.Text = "TX Data!!"

    If SampleCnt = 1 Then
        spdface.Row = 1
    Else
'        Call Row_Plus(spdface)
        If SampleCnt >= spdface.MaxRows Then
            spdface.MaxRows = spdface.MaxRows + 1
            spdface.Row = SampleCnt
        Else
            spdface.Row = SampleCnt
        End If
    End If
'---검사명, 검사결과값 db에 등록 ------------------------------------------------------------
    tcode = ""
    For i = 1 To 30
        If Trim(TestNameTable(i).code) <> "" Then
            tcode = Format$(i, "00")
            If tcode <> "" Then
                add_db_identb Format(SampleCnt + PrevCnt, "0000"), CStr(no_tmp)
                If tresult(i) <> "" Then
                    Call spdface.GetText(1, SampleCnt, temp)
                
                    If temp <> "" Then
                        add_db_resulttb Format(PrevCnt + SampleCnt, "0000"), tcode, tresult(i)
                    End If
                    Call spdsettext(spdface, TestNameTable(i).col_cnt, SampleCnt, tresult(i))
                End If
            End If
        End If
    Next
    
    identbOpenKey = True   'DB에 등록이 되었으므로 결과 등록이 가능한 조건이 되었음을 나타내는 키
    SampleCnt = tmp_iC
    
End Sub


Sub edit_data1()
        
    Dim tresult     As String
    Dim tcode       As String
    Dim i           As Integer
    Dim sTmpbuff    As String
    Dim sTestID     As String, sTestVal As String
    
    Dim slip_no     As String   '접수일자+구분+작업번호
    Dim seq_no      As String   'Sample 순서
    Dim iPos1       As Integer, iPos2   As Integer, iPos3   As Integer
    Dim iRow        As Integer
    Dim vTmp        As Variant
    Dim sDate       As String, iEtc As Integer
    Dim j           As Integer
    Dim tmpAssay(7) As String
    
    sDate = Mid$(mskOrdDate, 1, 4) + Mid$(mskOrdDate, 5, 2) + Mid$(mskOrdDate, 7, 2)
    
'임시Remark    If Not sDate = Mid$(AxResult(13), 10) Then Exit Sub
    
    seq_no = AxResult(2)
    
    iRow = D0SUB_SPREADGETROW(spdface, spdface.MaxCols, seq_no)
    If iRow < 1 Then Exit Sub
    
    spdface.GetText 1, iRow, vTmp:  slip_no = Trim$(vTmp)
    seq_no = Format$(nowcount + iRow, "0000")

'---검사명, 검사결과값 spread에 뿌리기------------------------------------------------------------
    txsapno.Text = Format(seq_no, "0000")
    txtguide.Text = "TX Data!!"
    
    Call add_db_identb(seq_no, slip_no)
    
    For j = 1 To 7
        tmpAssay(j) = GetByOne1(AxResult(3), AxResult(3))
    Next
    
    sTmpbuff = tmpAssay(5)

    Do While seq_no > 0
        sTestID = tmpAssay(4)
        sTestVal = AxResult(4) '& AxResult(5)
        
'        iEtc = 0
'        For i = 1 To spdface.MaxCols - 2
'            If Trim$(TestNameTable(i).eqno) = sTestID Then
'                iEtc = Val(TestNameTable(i).etc)
'                Exit For
'            End If
'        Next
'
'        sTestVal = Round(Val(sTestVal), iEtc)
        
        If Not sTestID = "" Then
            For i = 1 To spdface.MaxCols - 2
                If TestNameTable(i).eqno = sTestID Then
                    
                    If sTestVal <> "" Then
                        add_db_resulttb seq_no, sTestID, sTestVal
                        spdface.SetText TestNameTable(i).col_cnt, iRow, sTestVal
                    End If
                End If
            Next
        End If
        
        sTmpbuff = Mid$(sTmpbuff, iPos3 + 55)
        seq_no = Format$(nowcount + iRow, "0000")
    Loop
    
'    spdface.SetText spdface.MaxCols, iRow, seq_no
        
    F_iComm_Cnt = F_iComm_Cnt + 1
    
    identbOpenKey = True   'DB에 등록이 되었으므로 결과 등록이 가능한 조건이 되었음을 나타내는 키

End Sub
Sub Test()
    
    Dim rv%
    
    Test_OpenFlag = 1

    Open App.Path & "\dump_axsym.dat" For Input As #3
'    Open App.Path & "\axsym.log" For Input As #3
    Test_OpenFlag = 2
    wkbuf = ""
    Do While Not EOF(3)
        wkbuf = wkbuf & Input(1, #3)
    Loop

    Close #3

    Call PhaseCfg_Protocol
    
End Sub

Private Sub cboSelect_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub


Private Sub cmdAdd_Click()

    Dim iRow1   As Integer, iRow2   As Integer
    Dim vTmp    As Variant
    Dim sKeyno  As String, sTestID  As String, sPatnm   As String
    Dim iRack   As String, iPos    As Integer
    Dim sSendBuf    As String
    
    iRack = UCase(mskRackno.Text):    iPos = Val(mskPosition.Text)
    For iRow1 = 1 To spdList.MaxRows
        
        spdList.GetText 3, iRow1, vTmp:  sPatnm = Trim$(vTmp)
        spdList.GetText 2, iRow1, vTmp:  sKeyno = Trim$(vTmp)
        spdList.GetText 1, iRow1, vTmp
        
        If Trim(vTmp) = "1" And Not sKeyno = "" Then
        
            If Asc(iRack) > 70 Then
                MsgBox "Rack No 범위가 넘었습니다.  확인하세요.", vbCritical, Me.Caption
                Exit For
            End If
                
            With spdWorkList
                f_iWork_Row = f_iWork_Row + 1
                
                If f_iWork_Row > .MaxRows Then .MaxRows = .MaxRows + 1
                
                .SetText 1, f_iWork_Row, sKeyno
                .SetText 2, f_iWork_Row, sPatnm
                .SetText 3, f_iWork_Row, iRack
                .SetText 4, f_iWork_Row, CStr(iPos)
                
'                Call add_db_identb(Format(f_iWork_Row, "0000"), Trim(sKeyno))

            End With
                    
            '-- Rack, Position 증가
            iPos = iPos + 1
            If iPos > 15 Then
                iRack = Chr(Asc(iRack) + 1):  mskRackno.Text = iRack
                iPos = 1
            End If
            mskPosition.Text = CStr(iPos)
            
        End If
    Next

    With spdList
        For iRow1 = 1 To .MaxRows
            If iRow1 > .MaxRows Then Exit For
            .GetText 1, iRow1, vTmp
            
            If Trim(vTmp) = "1" Then
                .Row = iRow1
                .Action = SS_ACTION_DELETE_ROW
                
                .MaxRows = .MaxRows - 1
                iRow1 = iRow1 - 1
            End If
        Next
    End With

End Sub

Private Sub cmdclose_Click()

        Unload Me
        FrmFlag = 0
End Sub

Private Sub cmdDelete_Click()

    Dim rv As Integer
    Dim i  As Integer
    Dim CurrentTbRows As Integer
    Dim ExistTxtKey As Integer
    Dim tmpSlip
    Dim seq_no      As String
    
    If StartBRow = -1 And EndBRow = -1 Then
        StartBRow = 1
        EndBRow = nowcount - PrevCnt
    End If
    
    For i = StartBRow To EndBRow
        rv = spdface.GetText(1, i, tmpSlip)
        If tmpSlip = "" Then
            ExistTxtKey = False
            Exit For
        Else
            ExistTxtKey = True
        End If
    Next
    
    If identbOpenKey = True And ExistTxtKey = True Then
        If StartBCol = -1 And EndBCol = -1 And BlockKey = True Then
            
            rv = MsgBox("블록으로 지정된 Slip을 삭제하시겠습니까?", 4, Title & "  " & "Slip No. 삭제 확인!!")
            If rv = 7 Then
                BlockKey = False
                spdface.EditMode = True
                spdface.EditMode = False
                cmdclose.SetFocus
                Exit Sub
            End If
            
            identb.Index = "primarykey"
            resulttb.Index = "Seq_No"
            
'            identb.MoveLast
'            CurrentTbRows = nowcount - PrevCnt
            
            For i = Val(StartBRow) To Val(EndBRow)
                With spdface
                    .Row = i
                    .Col = .MaxCols:    seq_no = .Text
                End With
                
                identb.Seek "=", seq_no
                If Not identb.NoMatch Then identb.Delete
                
                SampleCnt = SampleCnt - 1
                
                resulttb.Seek "=", seq_no
                If resulttb.NoMatch = False Then
                   Do Until resulttb.EOF
                       If resulttb!seq_no <> seq_no Then Exit Do
                       
                       resulttb.Delete
                            
                       resulttb.MoveNext
                   Loop
                End If
            Next
            
        '삭제하는 Spread 라인의 텍스트를 지움.
            spdface.BlockMode = True
            spdface.Col = -1
            spdface.Col2 = -1
            spdface.Row = StartBRow
            spdface.Row2 = EndBRow
            spdface.Action = SS_ACTION_DELETE_ROW
            spdface.BlockMode = False

        '1st Column(SlipNo)의 색깔을 노란색
            spdface.BlockMode = True
            spdface.Col = 1
            spdface.Col2 = 1
            spdface.Row = -1
            spdface.Row2 = -1
            spdface.BackColor = &HC0FFFF
            spdface.BlockMode = False
            
            txsapno = ""
            txtguide = "Data 삭제!!"
            
        Else
        
            MsgBox "잘못된 삭제 방법입니다." & Chr(10) & "왼쪽의 회색빛 헤더부분을 클릭하거나 끌어서 해당줄의 전체가 어두워지게 한 후," & Chr(10) & "삭제를 하십시요!!"
        
        End If
   Else
   
        MsgBox "데이터가 없거나 검사 결과 전송을 받지 않으셨습니다!!"
        
   End If

   BlockKey = False
   spdface.EditMode = True
   spdface.EditMode = False
   
'현재의 Row를 점검
   spdface.Row = nowcount - PrevCnt
   
End Sub

Private Sub cmdInitial_Click()
    
    Dim SendBuff    As String
    
    Call f_subClear_Form
    
    fraWork.Visible = True
    Timer1.Enabled = False
    
 '########### CONNECTION ESTABLISH ######################
    
    SendBuff = ""
    
    SendBuff = Chr(1) & Chr(10)     '<SOH><LF>
 
 '--- HEADER BLOCK ---------------------------------------------------------------------
'''    SendBuff = SendBuff & "06" & " " & "HOSTNAMESIXTEENX" & " " & "00" & Chr(10)
'''    'IC(2)^ID(16)^BC(2)<LF> -- COBASCORE IC^ID^IDLE BLOCK^LF
    
    SendBuff = SendBuff & "09" & " " & "COBAS INTEGRA   " & " " & "00" & Chr(10)
    'IC(2)^ID(16)^BC(2)<LF> -- COBASCORE IC^ID^IDLE BLOCK^LF
    
    SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>
    
    SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>
    
    SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>

    Comm1.Output = SendBuff
    
End Sub

Private Sub cmdOrder_Click()

    Dim iRow1   As Integer, iRow2   As Integer
    Dim iCnt    As Integer
    Dim vTmp    As Variant
    Dim sKeyno  As String, sTestID  As String, sPatnm   As String
    Dim sRack   As String, sPos     As String
    Dim sSendBuf    As String
    
    iRow2 = 0
    For iRow1 = 1 To spdWorkList.MaxRows
        spdWorkList.GetText 1, iRow1, vTmp:  sKeyno = Trim$(vTmp)
        spdWorkList.GetText 2, iRow1, vTmp:  sPatnm = Trim$(vTmp)
        spdWorkList.GetText 3, iRow1, vTmp:  sRack = Trim$(vTmp)
        spdWorkList.GetText 4, iRow1, vTmp:  sPos = Trim$(vTmp)
        
        If sKeyno = "" Then Exit For
        
        iRow2 = iRow2 + 1
        ReDim Preserve f_tpTestList(1 To iRow2) As TYPE_TESTID
        
        With f_tpTestList(iRow2)
            Call f_subGet_검사항목(sKeyno, .sTestID, .iTestCnt)
        
            .sPatnm = sPatnm
            .sPatid = Mid$(sKeyno, 1, 8) + Mid$(sKeyno, 10, 1) + Mid$(sKeyno, 12, 5)
'            .sOrderId = "1" + Format(iRow2 + Val(mskOrderno.Text) - 1, "000") + Space(11) & " 00/00/0000"
            .sOrderId = iRow2
            .sRackno = "  " + CStr(sRack)
            .sPosition = IIf(Val(sPos) > 9, "", " ") + CStr(sPos)
            
        End With
        
    Next
    
    For iRow1 = 1 To iRow2
    
        With f_tpTestList(iRow1)
            For iCnt = 1 To .iTestCnt
'                sSendBuf = sSendBuf & "55 " & .sTestID(iCnt) & Chr(10)
            Next
                    
            '-- 결과 보기
            If iRow1 > spdface.MaxRows Then spdface.MaxRows = spdface.MaxRows + 1

            spdface.SetText 1, iRow1, Mid(.sPatid, 1, 8) + "-" + Mid$(.sPatid, 9, 1) + "-" + Mid$(.sPatid, 10, 5)
            spdface.SetText 2, iRow1, .sPatnm
            spdface.SetText spdface.MaxCols, iRow1, Mid$(.sOrderId, 1, 4)
    
        End With
    Next
    
    'Call Test
    Comm1.Output = Chr$(5) 'ENQ
    txtguide.Text = "Data 수신대기"
    
    CurSampCnt = iRow2
    
    If iCnt = 0 Then
        MsgBox "Work List에 등록할 자료가 없습니다. 조회를 실행해 주십시오.", vbCritical, Me.Caption
    Else
        fraWork.Visible = False
        
        txtStatus.Text = ""
        txtStatus.Visible = True
        Timer1.Enabled = True
    End If
    
End Sub

Private Sub cmdQuery_Click()

    Dim adoRs   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim iRow    As Integer
    Dim vTmp    As Variant
    
    Dim iRet    As Integer
    Dim sStr    As String
    Dim tData() As String
        
    '--- 조회조건 체크
    If Not IsDate(Format(mskDate(0), "####-##-##")) Then
        MsgBox "조회를 원하는 접수일자를 입력해 주십시오.", vbExclamation
        mskDate(0).SetFocus
        Exit Sub
    End If
    
    If Not IsDate(Format(mskDate(1), "####-##-##")) Then
        MsgBox "조회를 원하는 접수일자를 입력해 주십시오.", vbExclamation
        mskDate(1).SetFocus
        Exit Sub
    End If
    
    Me.MousePointer = 11
    spdList.MaxRows = 0
    
    sqlDoc = "select distinct" _
           & "       a.LABDATE,  a.NUMGBN, a.LABSQNO, b.PATNM" _
           & "  from LAB_DB..LAB030M a, PAT_DB..PAT010M b" _
           & " where a.LABDATE between '" & mskDate(0).Text & "' and '" & mskDate(1).Text & "'" _
           & "   and a.SUBCD   = ''" _
           & "   and a.SLIPCD + a.ORDCD + a.SPCCD in ("
    
    sqlDoc = sqlDoc + "''"
    Set dbcode = OpenDatabase(filename & codestr)
    Set tbcode = dbcode.OpenRecordset("cdtable")

    tbcode.MoveLast
    If tbcode.RecordCount > 0 Then tbcode.MoveFirst
    
    Do While Not tbcode.EOF
            
        sqlDoc = sqlDoc & ",'" + tbcode!code & "" & "'"
        tbcode.MoveNext
    
    Loop
    tbcode.Close:   dbcode.Close
    sqlDoc = sqlDoc & ")"
    
    If cboSelect.ListIndex = 0 Then
        sqlDoc = sqlDoc & "   and (a.RSTVAL = '' or a.RSTVAL is null) "
    Else
        sqlDoc = sqlDoc & "   and not (a.RSTVAL = '' or a.RSTVAL is null)"
    End If
    sqlDoc = sqlDoc _
           & "   and a.IDNO = b.IDNO " _
           & " order by a.NUMGBN, a.LABSQNO "
                
    adoRs.CursorLocation = adUseClient
    adoRs.Open sqlDoc, f_adoCn, adOpenStatic, adLockReadOnly
    
    If adoRs.RecordCount > 0 Then adoRs.MoveFirst
    
    iRow = 0
    Do While Not adoRs.EOF
    
        iRow = iRow + 1
        
        With spdList
        
            If iRow > .MaxRows Then .MaxRows = .MaxRows + 1
        
            .SetText 2, iRow, adoRs(0) & "" + "-" + adoRs(1) & "" + "-" & adoRs(2) & ""
            .SetText 3, iRow, Trim$(adoRs(3) & "")
        End With
        
        adoRs.MoveNext
    Loop
    adoRs.Close:    Set adoRs = Nothing
    
    If spdList.MaxRows = 0 Then _
        MsgBox "해당자료가 존재하지 않습니다.", vbInformation, Me.Caption
    
    Me.MousePointer = 0
    
End Sub


Private Sub Comm1_OnComm()
    Dim wkdat   As String
    Dim pnlDump As String
    
    Screen.MousePointer = 11
    
    Select Case Comm1.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                               ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            
            'txtguide.Text = "Data 수신준비"
            wkbuf = Comm1.Input
            Select Case Asc(wkbuf)
                Case 3  '-- ETX
'                    Timer1.Enabled = False
'                    Timer2.Enabled = False
                Case 21 '-- NAK
                    
                    Timer1.Enabled = True
                    Timer1.Interval = 12000
                    Timer2.Enabled = False
                Case 4  '-- EOT
                    Timer1.Enabled = False
                    Timer2.Enabled = False
'                    Call Test
                    Call PhaseCfg_Protocol
                    txtguide.Text = "Data 전송 완료!!"
                    txtStatus.Text = ""
                    wkbuf = ""
                Case 5  '-- ENQ
                    txtguide.Text = "Data 수신 중!!"
                    Timer2.Enabled = True
                    Timer2.Interval = 1000
                    Timer1.Enabled = False
'                    txtguide.Text = "Data 수신 중!!"
                    DoEvents
                Case Else

            End Select
            
            txtStatus.Text = txtStatus.Text + wkbuf

'            Print #1, wkbuf;    'Test
'            wkbuf = wkbuf + wkbuf
            
        
        Case MSCOMM_EV_CTS      'j
        Case MSCOMM_EV_DSR      ' Change in the DSR line.
        Case MSCOMM_EV_CD       ' Change in the CD line.
        Case MSCOMM_EV_RING     ' Change in the Ring Indicator.
        ' Errors
        Case MSCOMM_ER_BREAK    ' A Break was received.
        ' Code to handle a BREAK goes here, and so on.
        Case MSCOMM_ER_CTSTO    ' CTS Timeout.
        Case MSCOMM_ER_DSRTO    ' DSR Timeout.
        Case MSCOMM_ER_FRAME    ' Framing Error.
        Case MSCOMM_ER_OVERRUN  ' Data Lost.
        Case MSCOMM_ER_CDTO     ' CD (RLSD) Timeout.
        Case MSCOMM_ER_RXOVER   ' Receive buffer overflow.
        Case MSCOMM_ER_RXPARITY ' Parity Error.
        Case MSCOMM_ER_TXFULL   ' Transmit buffer full.
    End Select
           
    Screen.MousePointer = 0

End Sub
Private Sub Command1_Click()
 
    Call Test

End Sub

Private Sub Form_Load()
    Dim iCol    As Integer
    
    'form을 가운데에 위치
    Me.Top = 0
    Me.Left = 0
    Me.Height = INTmain00.Height - INTmain00.pnlMain.Height - 500
    Me.Width = INTmain00.Width - 200
    
    If UCase(D0COM_USERID) = "SUPER" Then Command1.Visible = True
    
    fraWork.ZOrder 0
    mskRackno.Text = "A"
    mskPosition.Text = "1"
    mskOrdDate.Text = Format(Now, "YYYYMMDD")
    
    Dim tablerows As Integer
    Dim iRow As Integer
    Dim i As Integer
    Dim TestItemNo As Integer
    
    Set f_adoCn = New ADODB.Connection
    f_adoCn.Open p_adoCnStr_1
    
    Set dbcode = OpenDatabase(filename & codestr)
    Set tbcode = dbcode.OpenRecordset("cdtable")
    
    tbcode.MoveLast
    tablerows = tbcode.RecordCount
        
    tbcode.MoveFirst
   
    iRow = 0
    Do While Not tbcode.EOF
        
        iRow = iRow + 1
        
        TestNameTable(iRow).eqno = tbcode!EQIPNO & ""
        TestNameTable(iRow).code = tbcode!code & ""
        TestNameTable(iRow).Name = tbcode!Name & ""
        TestNameTable(iRow).etc = tbcode!etc & ""
        
        If TestNameTable(iRow).code <> "" Then
            
            TestItemNo = TestItemNo + 1
            TestNameTable(iRow).col_cnt = TestItemNo + 2
            spdface.MaxCols = TestNameTable(iRow).col_cnt
            
            '-- 2001.11.12추가
            spdface.Font = "굴림체"
            spdface.FontSize = 9
            
            For iCol = 3 To spdface.MaxCols
                spdface.ColWidth(iCol) = 7
            Next
            
            Call spdsettext(spdface, TestNameTable(iRow).col_cnt, 0, TestNameTable(iRow).Name)
            
        End If
        
        tbcode.MoveNext
    Loop

    With spdface
        .MaxCols = .MaxCols + 1
        .Col = .MaxCols
        .ColHidden = True
    End With

    SampleCnt = 0
    F_iComm_Cnt = 1 '99.11.24 YEJ 추가
    
    mskDate(0).Text = Format(Now, "yyyymmdd")
    mskDate(1).Text = Format(Now, "yyyymmdd")
    
    With cboSelect
        .AddItem ("미등록 자료")
        .AddItem ("등록된 자료")
        .ListIndex = 0
    End With
    
'1st Column(SlipNo)의 색깔을 노란색
    spdface.BlockMode = True
    spdface.Col = 1
    spdface.Col2 = 1
    spdface.Row = -1
    spdface.Row2 = -1
    spdface.BackColor = &HC0FFFF
    spdface.BlockMode = False

'Interface Result를 일단 Lock
    spdface.BlockMode = True
    spdface.Col = 1
    spdface.Col2 = spdface.MaxCols
    spdface.Row = 1
    spdface.Row2 = spdface.MaxRows
    spdface.Lock = True
    spdface.BlockMode = False

'Spread.Row Initialization
    spdface.Row = 0
    
    tbcode.Close
    dbcode.Close
    
    LblMMDD.Caption = Val(Left$(textmmdd, 2)) & "월" & " " & Val(Right$(textmmdd, 2)) & "일"
    
    txtguide.Text = "RX Results!!"

    
    On Error GoTo PortOpenErr:
    errfound = False
    
    Set dbcomm = OpenDatabase(filename & commstr)
    Set tbcomm = dbcomm.OpenRecordset("cfgcomm")

    tbcomm.MoveFirst
        
    With transcfg
        .Port = tbcomm!Port
        .data_bit = tbcomm!data_bit
        .stop_bit = tbcomm!stop_bit
        .baud_rate = tbcomm!baud_rate
        .parity = tbcomm!parity
        .blocksize = tbcomm!blocksize
    End With
    
    tbcomm.Close
    dbcomm.Close
    
    With Comm1
        .CommPort = transcfg.Port
        .Settings = transcfg.baud_rate & "," & transcfg.parity & "," & transcfg.data_bit & "," & transcfg.stop_bit
        .PortOpen = True
        .RTSEnable = True
        .RThreshold = 1
    End With
          
    If errfound = True Then
        Me.MousePointer = 0
        interfacfrm.Show
        Unload interfacfrm
        Exit Sub
    End If
    
    Porttag = 1     'PortOpen에 성공을 나타냄
    identbOpenKey = False
    
    Set Db = OpenDatabase(filename & "comm\" & strmmdd + ".mdb")
    Set identb = Db.OpenRecordset("sp_identify")
    Set resulttb = Db.OpenRecordset("sp_result")
    
    identbOpenKey = True
    Porttag = 2     'OpenDB에 성공을 나타냄

'해당일의 이전까지의 샘플의 갯수를 db.identb에서 읽어 옴
    If identb.RecordCount > 0 Then
        identb.MoveLast
        nowcount = CInt(identb!seq_no)
    Else
        nowcount = 0
    End If
    
    PrevCnt = nowcount
        
    phase = 1
    EditCnt = 0
    SampleCnt = 0
    Test_OpenFlag = 0
    long_slip1 = 0
    long_slip2 = 0
    
    Open filename & machstr & ".log" For Output As #1
    
    If errfound = True Then
        Me.MousePointer = 0
        interfacfrm.Show
        Unload interfacfrm
        Exit Sub
    End If
    
    Open filename & machstr & "Buff.log" For Output As #2
          
    Porttag = 3     'OpenSamFile에 성공을 나타냄
    FrmFlag = 41
          
    If errfound = True Then
        Me.MousePointer = 0
        interfacfrm.Show
        Unload interfacfrm
        Exit Sub
    End If
          
    Timer1.Enabled = False
    Timer2.Enabled = False
        
    cmdInitial_Click
    
    If UCase(D0COM_USERID) = "SUPER" Then Command1.Visible = True


Exit Sub

PortOpenErr:
    
    errfound = True
    
    MsgBox "통신 구성 에러!!  통신구성을 다시 설정해 주십시요!!"
    
    If Porttag = 2 Then     'OpenDB까지 성공, OpenSamFile에서 실패
        identb.Close
        resulttb.Close
        Db.Close
        Close #1
    End If
    
    Porttag = 0
    
    If Test_OpenFlag = 1 Then    'Sub Test에서 Open시 에러발생하면 Test_OpenFlag = 1가 되고,
        Close #3                   '완전히 Open되면 Test_OpenFlag = 2
        Porttag = 3
    End If
    
    Resume Next
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If Porttag = 3 Then
        Comm1.PortOpen = False
        identb.Close
        resulttb.Close
        Db.Close
        identbOpenKey = False
        Close #1
        Close #2
    End If
    
    If Not f_adoCn.State = adStateClosed Then
        f_adoCn.Close:  Set f_adoCn = Nothing
    End If
    
End Sub

Private Sub mskDate_GotFocus(Index As Integer)
    
    With mskDate(Index)
        .SelStart = 0
        .SelLength = .MaxLength
    End With

End Sub

Private Sub mskDate_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        KeyAscii = 0
    ElseIf Not KeyAscii = vbKeyBack Then
        mskDate(Index).SelLength = 1
    End If
    
End Sub

Private Sub mskOrderno_GotFocus()

    mskOrderno.SelStart = 0
    mskOrderno.SelLength = Len(mskOrderno.Text)

End Sub

Private Sub mskPosition_GotFocus()
    
    mskPosition.SelStart = 0
    mskPosition.SelLength = Len(mskPosition.Text)

End Sub

Private Sub mskPosition_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If

End Sub


Private Sub mskRackno_GotFocus()

    mskRackno.SelStart = 0
    mskRackno.SelLength = Len(mskRackno.Text)

End Sub

Private Sub mskRackno_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub


Private Sub spdface_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    StartBCol = CInt(BlockCol)
    StartBRow = CInt(BlockRow)
    EndBCol = CInt(BlockCol2)
    EndBRow = CInt(BlockRow2)
    BlockKey = True
End Sub

Private Sub spdList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    Dim sNo As Long
    Dim eNo As Long
    Dim iRow    As Integer
    Dim Tmp As Variant
    
    If BlockRow = 0 Or BlockRow2 = 0 Or BlockRow = BlockRow2 Then Exit Sub
    
    If BlockRow < BlockRow2 Then
        sNo = BlockRow: eNo = BlockRow2
    Else
        sNo = BlockRow2: eNo = BlockRow
    End If

    For iRow = sNo To eNo
        With spdList
            Call .GetText(1, iRow, Tmp)
            If Tmp = True Then
                Call .SetText(1, iRow, "0")
            Else
                Call .SetText(1, iRow, "1")
            End If
        End With
    Next iRow
    
End Sub

Private Sub spdList_GotFocus()
    With spdList
        If OldRow <> 0 Then
            .Row = OldRow
            .Col = -1
            .BackColor = &H80000005
        End If
        If .ActiveRow = 1 Then
            .Row = .ActiveRow
            .Col = -1
    
            If .Lock = False Then
                .BackColor = &HEEEFFF
                OldRow = .ActiveRow
            End If
        End If
    End With
End Sub

Private Sub spdList_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
      If Row <> NewRow Then
        If OldRow = 0 Then OldRow = Row
        
        With spdList
            If NewRow <> -1 Then
                .Row = OldRow
                .Col = -1
                .BackColor = &H80000005
                
                .Row = NewRow
                .Col = -1

                .BackColor = &HEEEFFF
                OldRow = NewRow
                'Call Disp_Data(OldRow)     '해당 Sample의 Order표시
            End If
        End With
    End If
End Sub

Private Sub cmdDown_Click()

    fraWork.Visible = False
    
    txtStatus.Text = ""
    txtStatus.Visible = True
    
    Comm1.Output = Chr$(5) 'ENQ
'    Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
    
    Comm1.Output = Chr$(5) 'ENQ

End Sub


Private Sub Timer2_Timer()

    Comm1.Output = Chr$(6) 'ACK
    
End Sub

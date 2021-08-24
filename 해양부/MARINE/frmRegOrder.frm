VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmJoui 
   BackColor       =   &H00FFFFFF&
   Caption         =   "조위관측소"
   ClientHeight    =   12720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22140
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12720
   ScaleWidth      =   22140
   WindowState     =   2  '최대화
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   3375
      Left            =   17070
      TabIndex        =   34
      Top             =   9120
      Visible         =   0   'False
      Width           =   4785
      Begin VB.ComboBox cboVPNList 
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmRegOrder.frx":0000
         Left            =   1350
         List            =   "frmRegOrder.frx":0002
         Style           =   2  '드롭다운 목록
         TabIndex        =   35
         Top             =   510
         Width           =   2415
      End
      Begin FPSpread.vaSpread spdBUOYList 
         Height          =   1695
         Left            =   300
         TabIndex        =   37
         Top             =   1260
         Width           =   6525
         _Version        =   393216
         _ExtentX        =   11509
         _ExtentY        =   2990
         _StockProps     =   64
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   15921919
         GridShowVert    =   0   'False
         MaxCols         =   3
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16773087
         SpreadDesigner  =   "frmRegOrder.frx":0004
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "관측소"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   510
         TabIndex        =   36
         Top             =   570
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   60
      TabIndex        =   2
      Top             =   10380
      Width           =   20355
      Begin VB.Timer tmrResult 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   9000
         Top             =   120
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "설정 저장"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5850
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   210
         Width           =   1485
      End
      Begin VB.TextBox txtInterval60 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7830
         MaxLength       =   10
         TabIndex        =   7
         Top             =   180
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.ComboBox cboIntervalGrade 
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmRegOrder.frx":053A
         Left            =   2940
         List            =   "frmRegOrder.frx":053C
         Style           =   2  '드롭다운 목록
         TabIndex        =   5
         Top             =   210
         Width           =   1005
      End
      Begin VB.CheckBox chkRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "자동갱신"
         Height          =   255
         Left            =   420
         TabIndex        =   4
         Top             =   270
         Width           =   1065
      End
      Begin VB.TextBox txtInterval 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   3
         Top             =   210
         Width           =   750
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "데이터 갱신"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4290
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   210
         Width           =   1485
      End
   End
   Begin VB.Frame fra1 
      BackColor       =   &H00FFFFFF&
      Height          =   9675
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   20355
      Begin VB.Frame Frame2 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   6540
         TabIndex        =   24
         Top             =   120
         Width           =   13695
         Begin VB.ComboBox cboFromHour 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmRegOrder.frx":053E
            Left            =   2340
            List            =   "frmRegOrder.frx":0540
            Style           =   2  '드롭다운 목록
            TabIndex        =   28
            Top             =   210
            Width           =   795
         End
         Begin VB.ComboBox cboSearchCount 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmRegOrder.frx":0542
            Left            =   11490
            List            =   "frmRegOrder.frx":0544
            Style           =   2  '드롭다운 목록
            TabIndex        =   27
            Top             =   240
            Width           =   1005
         End
         Begin VB.CommandButton cmdSearch2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "조회"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5730
            Style           =   1  '그래픽
            TabIndex        =   26
            Top             =   180
            Width           =   1095
         End
         Begin VB.ComboBox cboToHour 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmRegOrder.frx":0546
            Left            =   4830
            List            =   "frmRegOrder.frx":0548
            Style           =   2  '드롭다운 목록
            TabIndex        =   25
            Top             =   210
            Width           =   795
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   345
            Left            =   900
            TabIndex        =   29
            Top             =   210
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   137625601
            CurrentDate     =   43884
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   345
            Left            =   3420
            TabIndex        =   30
            Top             =   210
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   137625601
            CurrentDate     =   43884
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "기간"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   300
            TabIndex        =   33
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '투명
            Caption         =   "출력건수"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   10620
            TabIndex        =   32
            Top             =   330
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   3180
            TabIndex        =   31
            Top             =   270
            Width           =   195
         End
      End
      Begin FPSpread.vaSpread spdVPNList 
         Height          =   8985
         Left            =   120
         TabIndex        =   6
         Top             =   210
         Width           =   6345
         _Version        =   393216
         _ExtentX        =   11192
         _ExtentY        =   15849
         _StockProps     =   64
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   15921919
         GridShowVert    =   0   'False
         MaxCols         =   2
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16773087
         SpreadDesigner  =   "frmRegOrder.frx":054A
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
      Begin FPSpread.vaSpread spdJOUYList 
         Height          =   3825
         Left            =   6540
         TabIndex        =   38
         Top             =   960
         Width           =   13695
         _Version        =   393216
         _ExtentX        =   24156
         _ExtentY        =   6747
         _StockProps     =   64
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   15921919
         GridShowVert    =   0   'False
         MaxCols         =   2
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16773087
         SpreadDesigner  =   "frmRegOrder.frx":0994
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3825
         Left            =   6510
         TabIndex        =   39
         Top             =   5370
         Width           =   13695
         _Version        =   393216
         _ExtentX        =   24156
         _ExtentY        =   6747
         _StockProps     =   64
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   15921919
         GridShowVert    =   0   'False
         MaxCols         =   2
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16773087
         SpreadDesigner  =   "frmRegOrder.frx":0DDE
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "관측소 자료"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6780
         TabIndex        =   40
         Top             =   5100
         Width           =   1935
      End
   End
   Begin VB.Frame fra3 
      BackColor       =   &H00FFFFFF&
      Height          =   9675
      Left            =   60
      TabIndex        =   10
      Top             =   2520
      Width           =   20355
      Begin VB.Frame Frame3 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   " 조위관측소 자료수집 로그 검색조건"
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   330
         TabIndex        =   12
         Top             =   180
         Width           =   19695
         Begin VB.ComboBox cboVPNList3 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmRegOrder.frx":1228
            Left            =   1110
            List            =   "frmRegOrder.frx":122A
            Style           =   2  '드롭다운 목록
            TabIndex        =   16
            Top             =   360
            Width           =   2415
         End
         Begin VB.ComboBox cboFromHour3 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmRegOrder.frx":122C
            Left            =   6060
            List            =   "frmRegOrder.frx":122E
            Style           =   2  '드롭다운 목록
            TabIndex        =   15
            Top             =   330
            Width           =   795
         End
         Begin VB.CommandButton cmdSearch3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "검색"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9510
            Style           =   1  '그래픽
            TabIndex        =   14
            Top             =   300
            Width           =   1095
         End
         Begin VB.ComboBox cboToHour3 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmRegOrder.frx":1230
            Left            =   8610
            List            =   "frmRegOrder.frx":1232
            Style           =   2  '드롭다운 목록
            TabIndex        =   13
            Top             =   330
            Width           =   795
         End
         Begin MSComCtl2.DTPicker dtpFromDate3 
            Height          =   345
            Left            =   4620
            TabIndex        =   17
            Top             =   330
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   137625601
            CurrentDate     =   43884
         End
         Begin MSComCtl2.DTPicker dtpToDate3 
            Height          =   345
            Left            =   7200
            TabIndex        =   18
            Top             =   330
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   137625601
            CurrentDate     =   43884
         End
         Begin VB.Label lblSerchSatus 
            BackStyle       =   0  '투명
            Caption         =   "검색상태 :"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   12810
            TabIndex        =   23
            Top             =   390
            Width           =   975
         End
         Begin VB.Label Label9 
            BackStyle       =   0  '투명
            Caption         =   "관측소"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   270
            TabIndex        =   22
            Top             =   420
            Width           =   975
         End
         Begin VB.Label Label8 
            BackStyle       =   0  '투명
            Caption         =   "기간"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4020
            TabIndex        =   21
            Top             =   420
            Width           =   975
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '투명
            Caption         =   "검색상태 :"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   11760
            TabIndex        =   20
            Top             =   390
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   6960
            TabIndex        =   19
            Top             =   390
            Width           =   195
         End
      End
      Begin FPSpread.vaSpread spdBUOYViewList 
         Height          =   8535
         Left            =   330
         TabIndex        =   11
         Top             =   1050
         Width           =   19695
         _Version        =   393216
         _ExtentX        =   34740
         _ExtentY        =   15055
         _StockProps     =   64
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   15921919
         GridShowVert    =   0   'False
         MaxCols         =   2
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16773087
         SpreadDesigner  =   "frmRegOrder.frx":1234
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
   Begin VB.Frame fra2 
      BackColor       =   &H00FFFFFF&
      Height          =   9675
      Left            =   60
      TabIndex        =   9
      Top             =   720
      Width           =   20355
   End
End
Attribute VB_Name = "frmJoui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intOneMinute As Integer

'-- 컨트롤초기화
Private Sub CtlInitializing()
    Dim iHour   As Integer
    
    With spdVPNList
        .MaxCols = 3
        .MaxRows = 0
        
        Call SetText(spdVPNList, "관측소", 0, 1):           .ColWidth(1) = 20
        Call SetText(spdVPNList, "관측시간", 0, 2):         .ColWidth(2) = 30
        Call SetText(spdVPNList, "관측소ID", 0, 3):         .ColWidth(3) = 0
    End With
    
    With spdBUOYList
        .MaxCols = 4
        .MaxRows = 0
        
        Call SetText(spdBUOYList, "관측소", 0, 1):          .ColWidth(1) = 20
        Call SetText(spdBUOYList, "관측시간", 0, 2):        .ColWidth(2) = 30
        Call SetText(spdBUOYList, "업체", 0, 3):            .ColWidth(3) = 20
        Call SetText(spdBUOYList, "관측소ID", 0, 4):        .ColWidth(4) = 0
    End With
            
    '------------------------------------------------------------------------
    With spdJOUYList
        .MaxCols = 5
        .MaxRows = 0
        
        Call SetText(spdJOUYList, "관측소ID", 0, 1):         .ColWidth(1) = 20
        Call SetText(spdJOUYList, "관측소명", 0, 2):         .ColWidth(2) = 30
        Call SetText(spdJOUYList, "관측시간", 0, 3):         .ColWidth(3) = 30
        Call SetText(spdJOUYList, "로그기록시간", 0, 4):     .ColWidth(4) = 30
        Call SetText(spdJOUYList, "로그내용", 0, 5):         .ColWidth(5) = 50
    End With
    
    '------------------------------------------------------------------------
    With spdBUOYViewList
        .MaxCols = 9
        .MaxRows = 0
        
        Call SetText(spdBUOYViewList, "경도", 0, 1):            .ColWidth(1) = 15
        Call SetText(spdBUOYViewList, "위도", 0, 2):            .ColWidth(2) = 15
        Call SetText(spdBUOYViewList, "관측시간", 0, 3):        .ColWidth(3) = 30
        Call SetText(spdBUOYViewList, "풍속", 0, 4):            .ColWidth(4) = 15
        Call SetText(spdBUOYViewList, "풍향", 0, 5):            .ColWidth(5) = 15
        Call SetText(spdBUOYViewList, "돌풍(최대풍속)", 0, 6):  .ColWidth(6) = 20
        Call SetText(spdBUOYViewList, "기온", 0, 7):            .ColWidth(7) = 15
        Call SetText(spdBUOYViewList, "기압", 0, 8):            .ColWidth(8) = 15
        Call SetText(spdBUOYViewList, "부이방향", 0, 9):        .ColWidth(9) = 20
    End With
    
    
    cboIntervalGrade.Clear
    cboIntervalGrade.AddItem "초"
    cboIntervalGrade.AddItem "분"
    cboIntervalGrade.ListIndex = 0
    
    dtpFromDate.Value = Now
    dtpFromDate3.Value = Now
    dtpToDate.Value = Now
    dtpToDate3.Value = Now

    For iHour = 1 To 24
        cboFromHour.AddItem iHour
        cboFromHour3.AddItem iHour
        cboToHour.AddItem iHour
        cboToHour3.AddItem iHour
    Next
    
    cboFromHour.ListIndex = 0
    cboFromHour3.ListIndex = 0
    cboToHour.ListIndex = 0
    cboToHour3.ListIndex = 0
    
    cboSearchCount.Clear
    For iHour = 1 To 10
        cboSearchCount.AddItem iHour * 10
    Next
    cboSearchCount.ListIndex = 0
    
    lblSerchSatus.Caption = ""
    
    gSORT = 0

End Sub

Private Sub chkRefresh_Click()

    If chkRefresh.Value = "1" Then
        tmrResult.Interval = 1000
        tmrResult.Enabled = True
    Else
        tmrResult.Enabled = False
    End If
    
End Sub

Private Sub cmdSave_Click()
    
    If chkRefresh.Value = "1" Then
        Call WritePrivateProfileString("USER", "AUTOREFREH", "1", App.PATH & "\MARINE.ini")
    Else
        Call WritePrivateProfileString("USER", "AUTOREFREH", "0", App.PATH & "\MARINE.ini")
    End If
    
    Call WritePrivateProfileString("USER", "INTERVAL", txtInterval.Text, App.PATH & "\MARINE.ini")
    Call WritePrivateProfileString("USER", "INTERGBN", cboIntervalGrade.Text, App.PATH & "\MARINE.ini")

End Sub

Private Sub cmdSearch_Click()
    
    Call GetDataSearch

End Sub

Private Sub cmdSearch2_Click()
        
    '조위관측소 가져오기
    Call GetJOUYList(mGetP(cboVPNList.Text, 2, "|"), dtpFromDate.Value, cboFromHour.Text, dtpToDate.Value, cboToHour.Text, cboSearchCount.Text)

End Sub

Private Sub cmdSearch3_Click()
    
    Call GetBOUYViewList(mGetP(cboVPNList.Text, 2, "|"), dtpFromDate.Value, cboFromHour.Text, dtpToDate.Value, cboToHour.Text)

End Sub

Private Sub cmdView_Click(Index As Integer)
    Dim i   As Integer
    
    For i = 0 To 2
        cmdView(i).BackColor = vbWhite
    Next
        
    fra1.Visible = False
    fra2.Visible = False
    fra3.Visible = False
        
    cmdView(Index).BackColor = &HC0FFFF
    
    If Index = 0 Then
        fra1.Visible = True
    ElseIf Index = 1 Then
        fra2.Visible = True
        
        '조위관측소-VPN 리스트 가져오기
        Call GetVPNList_Combo(cboVPNList, "")

        '조위관측소 가져오기
        Call GetJOUYList("", "", "", "", "", 0)

    ElseIf Index = 2 Then
        fra3.Visible = True
    
        '조위관측소-VPN 리스트 가져오기
        Call GetVPNList_Combo(cboVPNList3, "")
        
        '해양부이VIEW
        Call GetBOUYViewList("", "", "", "", "")
    End If

End Sub

Private Sub Form_Load()

    Call CtlInitializing
    
    fra1.Visible = True
    fra1.ZOrder 0

    txtInterval.Text = gInterVal
    cboIntervalGrade.Text = gInterGbn
    If gAutoRefresh = "1" Then
        chkRefresh.Value = "1"
    Else
        chkRefresh.Value = "0"
    End If
    
'    If gInterGbn = "분" Then
'        txtInterval60.Visible = True
'    Else
'        txtInterval60.Visible = False
'    End If
    
    tmrResult.Interval = 1000
    tmrResult.Enabled = True

    intOneMinute = 0
    
    If cn_Server_Flag = True Then
        Call GetDataSearch
    End If
    
End Sub

Private Sub GetDataSearch()

    '조위관측소-VPN 리스트 가져오기
    Call GetVPNList("")

    '종합해양관측부이 가져오기
    Call GetBUOYList("")

End Sub

Private Sub GetVPNList(Optional ByVal pDT_TS_ID As String)
    Dim pAdoRS  As ADODB.Recordset
    Dim intRow  As Integer
    
    Set pAdoRS = Get_VPNList(pDT_TS_ID)
    
    spdVPNList.MaxRows = 0
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        With spdVPNList
            Do Until pAdoRS.EOF
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
                Call SetText(spdVPNList, pAdoRS.Fields("TS_NAME").Value & "", intRow, 1)
                Call SetText(spdVPNList, pAdoRS.Fields("DT_TIME").Value & "", intRow, 2)
                Call SetText(spdVPNList, pAdoRS.Fields("DT_TS_ID").Value & "", intRow, 3)
                pAdoRS.MoveNext
            Loop
        End With
    End If
    
    pAdoRS.Close

End Sub

Private Sub GetVPNList_Combo(ByVal obj As Object, Optional ByVal pDT_TS_ID As String)
    Dim pAdoRS  As ADODB.Recordset
    
    Set pAdoRS = Get_VPNList(pDT_TS_ID)
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        obj.Clear
        Do Until pAdoRS.EOF
            obj.AddItem pAdoRS.Fields("TS_NAME").Value & Space(30) & "|" & pAdoRS.Fields("DT_TS_ID").Value
            pAdoRS.MoveNext
        Loop
        obj.ListIndex = 0
    End If
    
    pAdoRS.Close

End Sub

'종합해양관측부이
Private Sub GetBUOYList(Optional ByVal pSTATION_ID As String)
    Dim pAdoRS  As ADODB.Recordset
    Dim intRow  As Integer
    
    Set pAdoRS = Get_BUOYList(pSTATION_ID)
    
    spdBUOYList.MaxRows = 0
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        With spdBUOYList
            Do Until pAdoRS.EOF
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
                Call SetText(spdBUOYList, pAdoRS.Fields("STATION_NAME").Value & "", intRow, 1)
                Call SetText(spdBUOYList, pAdoRS.Fields("OBS_TIME").Value & "", intRow, 2)
                Call SetText(spdBUOYList, pAdoRS.Fields("EQUIP_ID").Value & "", intRow, 3)
                Call SetText(spdBUOYList, pAdoRS.Fields("STATION_ID").Value & "", intRow, 4)
                pAdoRS.MoveNext
            Loop
        End With
    End If
    
    pAdoRS.Close

End Sub

Private Sub GetJOUYList(Optional ByVal pDT_TS_ID As String, _
                        Optional ByVal pDT_F_TIME As String, _
                        Optional ByVal pDT_F_HOUR As String, _
                        Optional ByVal pDT_T_TIME As String, _
                        Optional ByVal pDT_T_HOUR As String, _
                        Optional ByVal pCount As Integer)
    
    Dim pAdoRS  As ADODB.Recordset
    Dim intRow  As Integer
    
    Set pAdoRS = Get_JOUYList(pDT_TS_ID, pDT_F_TIME, pDT_F_HOUR, pDT_T_TIME, pDT_T_HOUR, pCount)
    
    spdBUOYList.MaxRows = 0
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        With spdJOUYList
            Do Until pAdoRS.EOF
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
                Call SetText(spdJOUYList, pAdoRS.Fields("DT_TS_ID").Value & "", intRow, 1)      '관측소ID
                Call SetText(spdJOUYList, pAdoRS.Fields("TS_NAME").Value & "", intRow, 2)       '관측소명
                Call SetText(spdJOUYList, pAdoRS.Fields("DT_TIME").Value & "", intRow, 3)       '관측시간
                Call SetText(spdJOUYList, pAdoRS.Fields("REG_DATE").Value & "", intRow, 4)      '로그기록시간
                Call SetText(spdJOUYList, pAdoRS.Fields("LOG_CONTENT").Value & "", intRow, 5)   '로그내용
                pAdoRS.MoveNext
            Loop
        End With
    End If
    
    pAdoRS.Close

End Sub

Private Sub GetBOUYViewList(Optional ByVal pDT_TS_ID As String, _
                            Optional ByVal pDT_F_TIME As String, _
                            Optional ByVal pDT_F_HOUR As String, _
                            Optional ByVal pDT_T_TIME As String, _
                            Optional ByVal pDT_T_HOUR As String)
    
    Dim pAdoRS  As ADODB.Recordset
    Dim intRow  As Integer
    
    Set pAdoRS = Get_BOUYViewList(pDT_TS_ID, pDT_F_TIME, pDT_F_HOUR, pDT_T_TIME, pDT_T_HOUR)
    
    spdBUOYList.MaxRows = 0
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        With spdBUOYViewList
            Do Until pAdoRS.EOF
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
                Call SetText(spdBUOYViewList, pAdoRS.Fields("DT_TS_ID").Value & "", intRow, 1)      '관측소ID
                Call SetText(spdBUOYViewList, pAdoRS.Fields("TS_NAME").Value & "", intRow, 2)       '관측소명
                Call SetText(spdBUOYViewList, pAdoRS.Fields("DT_TIME").Value & "", intRow, 3)       '관측시간
                Call SetText(spdBUOYViewList, pAdoRS.Fields("REG_DATE").Value & "", intRow, 4)      '로그기록시간
                Call SetText(spdBUOYViewList, pAdoRS.Fields("LOG_CONTENT").Value & "", intRow, 5)   '로그내용
                pAdoRS.MoveNext
            Loop
        End With
    End If
    
    pAdoRS.Close

End Sub




Private Sub tmrResult_Timer()
    
    If chkRefresh.Value = "1" Then
        If cboIntervalGrade.Text = "초" Then
            txtInterval.Text = txtInterval.Text - 1
            If txtInterval.Text = "0" Then
                '자동갱신
                If chkRefresh.Value = "1" Then
                    '조회추가
                    Call GetDataSearch
                End If
                
                txtInterval.Enabled = False
                txtInterval.Text = gInterVal
            End If
        Else
            intOneMinute = intOneMinute + 1
            txtInterval60.Text = intOneMinute
            If intOneMinute = 60 Then
                intOneMinute = 0
                txtInterval.Text = txtInterval.Text - 1
                If txtInterval.Text = "0" Then
                    '자동갱신
                    If chkRefresh.Value = "1" Then
                        '조회추가
                        Call GetDataSearch
                    End If
                    
                    txtInterval.Enabled = False
                    txtInterval.Text = gInterVal
                End If
            End If
        End If
    End If

End Sub

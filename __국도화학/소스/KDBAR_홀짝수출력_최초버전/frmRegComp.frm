VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmRegProd 
   BackColor       =   &H00FFFFFF&
   Caption         =   "제품코드 등록"
   ClientHeight    =   12075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18390
   BeginProperty Font 
      Name            =   "맑은 고딕"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12075
   ScaleWidth      =   18390
   Tag             =   "LBL_M_PROD"
   WindowState     =   2  '최대화
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 제품등록 "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8235
      Left            =   90
      TabIndex        =   25
      Top             =   1050
      Width           =   17865
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7725
         Left            =   4830
         TabIndex        =   26
         Top             =   240
         Width           =   12765
         Begin VB.TextBox txtProdPrtNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   8670
            MaxLength       =   50
            TabIndex        =   52
            Text            =   "화성사업소"
            Top             =   780
            Width           =   3705
         End
         Begin VB.TextBox txtProdMatCd1 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   9450
            MaxLength       =   20
            TabIndex        =   50
            Text            =   "화성사업소"
            Top             =   2190
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.CommandButton cmdMakeCode 
            BackColor       =   &H00E0E0E0&
            Caption         =   "신규코드발번"
            Height          =   375
            Left            =   6420
            Style           =   1  '그래픽
            TabIndex        =   49
            Top             =   360
            Width           =   1395
         End
         Begin VB.ComboBox cboProdPcnNo 
            Height          =   375
            ItemData        =   "frmRegComp.frx":0000
            Left            =   2670
            List            =   "frmRegComp.frx":0007
            Style           =   2  '드롭다운 목록
            TabIndex        =   18
            Top             =   6060
            Width           =   3705
         End
         Begin VB.ComboBox cboProdCtrlYN 
            Height          =   375
            ItemData        =   "frmRegComp.frx":000E
            Left            =   2670
            List            =   "frmRegComp.frx":0018
            Style           =   2  '드롭다운 목록
            TabIndex        =   17
            Top             =   5640
            Width           =   3705
         End
         Begin VB.ComboBox cboProdSlitFA 
            Height          =   375
            ItemData        =   "frmRegComp.frx":0038
            Left            =   2670
            List            =   "frmRegComp.frx":003F
            Style           =   2  '드롭다운 목록
            TabIndex        =   16
            Top             =   5220
            Width           =   3705
         End
         Begin VB.ComboBox cboProdLineFA 
            Height          =   375
            ItemData        =   "frmRegComp.frx":0055
            Left            =   2670
            List            =   "frmRegComp.frx":005C
            Style           =   2  '드롭다운 목록
            TabIndex        =   15
            Top             =   4800
            Width           =   3705
         End
         Begin VB.CheckBox chkUsedYN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "사용"
            Height          =   255
            Left            =   2730
            TabIndex        =   20
            Top             =   6540
            Width           =   795
         End
         Begin VB.ComboBox cboProdMatCd 
            Height          =   375
            ItemData        =   "frmRegComp.frx":006E
            Left            =   6000
            List            =   "frmRegComp.frx":0070
            Style           =   2  '드롭다운 목록
            TabIndex        =   9
            Top             =   2310
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   8820
            TabIndex        =   44
            Top             =   6780
            Width           =   3525
            Begin VB.CommandButton cmdDelete 
               BackColor       =   &H00E0E0E0&
               Caption         =   "삭제"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   1200
               Style           =   1  '그래픽
               TabIndex        =   22
               Top             =   150
               Width           =   1095
            End
            Begin VB.CommandButton cmdOK 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Caption         =   "저장"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   60
               Style           =   1  '그래픽
               TabIndex        =   21
               Top             =   150
               Width           =   1095
            End
            Begin VB.CommandButton cmdClose 
               BackColor       =   &H00E0E0E0&
               Caption         =   "닫기"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   2340
               Style           =   1  '그래픽
               TabIndex        =   23
               Top             =   150
               Width           =   1095
            End
         End
         Begin VB.TextBox txtProdChimPN 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2670
            MaxLength       =   20
            TabIndex        =   13
            Text            =   "화성사업소"
            Top             =   3960
            Width           =   3705
         End
         Begin VB.TextBox txtProdSize 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2670
            MaxLength       =   20
            TabIndex        =   12
            Text            =   "화성사업소"
            Top             =   3540
            Width           =   3705
         End
         Begin VB.TextBox txtProdStorTemp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2670
            MaxLength       =   20
            TabIndex        =   11
            Text            =   "화성사업소"
            Top             =   3120
            Width           =   3705
         End
         Begin VB.TextBox txtExprMon 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2670
            MaxLength       =   2
            TabIndex        =   10
            Text            =   "화성사업소"
            Top             =   2700
            Width           =   2370
         End
         Begin VB.TextBox txtProdMatCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2670
            MaxLength       =   30
            TabIndex        =   33
            Text            =   "1235"
            Top             =   2280
            Width           =   2355
         End
         Begin VB.TextBox txtProdLen 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2670
            MaxLength       =   10
            TabIndex        =   8
            Text            =   "화성사업소"
            Top             =   1860
            Width           =   2370
         End
         Begin VB.ComboBox cboCompCd 
            Height          =   375
            Left            =   2670
            Style           =   2  '드롭다운 목록
            TabIndex        =   7
            Top             =   1200
            Width           =   3705
         End
         Begin VB.TextBox txtCompCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   6420
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   30
            Text            =   "화성사업소"
            Top             =   1200
            Width           =   1365
         End
         Begin VB.TextBox txtProdNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   2670
            MaxLength       =   50
            TabIndex        =   6
            Text            =   "화성사업소"
            Top             =   780
            Width           =   3705
         End
         Begin VB.TextBox txtProdCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   2670
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   5
            Text            =   "P0001"
            Top             =   360
            Width           =   3705
         End
         Begin VB.TextBox txtVenCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2670
            MaxLength       =   20
            TabIndex        =   14
            Text            =   "화성사업소"
            Top             =   4380
            Width           =   3705
         End
         Begin VB.TextBox txtItemBar 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   10020
            MaxLength       =   20
            TabIndex        =   19
            Text            =   "화성사업소"
            Top             =   4740
            Visible         =   0   'False
            Width           =   3705
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "제품명(출력용)"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   15
            Left            =   6420
            TabIndex        =   53
            Top             =   780
            Width           =   2205
         End
         Begin VB.Label lblMatInfo 
            BackStyle       =   0  '투명
            Caption         =   "INNOLUX 릴라벨에서는 MATERIAL 코드로 표현됨"
            Height          =   465
            Left            =   6480
            TabIndex        =   51
            Top             =   3960
            Width           =   4995
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "바코드"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   14
            Left            =   7770
            TabIndex        =   43
            Top             =   4740
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "달(Month)"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   5160
            TabIndex        =   47
            Top             =   2760
            Width           =   1125
         End
         Begin VB.Label lblWorkDate 
            BackStyle       =   0  '투명
            Caption         =   "미터(M)"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   5160
            TabIndex        =   46
            Top             =   1920
            Width           =   1155
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Height          =   30
            Left            =   420
            Top             =   1710
            Width           =   8145
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "제품코드"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   420
            TabIndex        =   27
            Top             =   360
            Width           =   2200
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "제품명"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   420
            TabIndex        =   28
            Top             =   780
            Width           =   2200
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "고객사"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   420
            TabIndex        =   29
            Top             =   1200
            Width           =   2200
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "제품길이"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   420
            TabIndex        =   31
            Top             =   1860
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "자재코드"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   420
            TabIndex        =   32
            Top             =   2280
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "보관온도"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   420
            TabIndex        =   35
            Top             =   3120
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "Size"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   420
            TabIndex        =   36
            Top             =   3540
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "Chimei P/N"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   8
            Left            =   420
            TabIndex        =   37
            Top             =   3960
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "유효기간"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   420
            TabIndex        =   34
            Top             =   2700
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "PCN 차수"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   13
            Left            =   420
            TabIndex        =   42
            Top             =   6060
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "관리선 이탈여부"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   12
            Left            =   420
            TabIndex        =   41
            Top             =   5640
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "Slitting 공장"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   11
            Left            =   420
            TabIndex        =   40
            Top             =   5220
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "제조라인공장"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   10
            Left            =   420
            TabIndex        =   39
            Top             =   4800
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "Vendor코드"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   9
            Left            =   420
            TabIndex        =   38
            Top             =   4380
            Width           =   2205
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "사용여부"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   420
            TabIndex        =   45
            Top             =   6480
            Width           =   2205
         End
      End
      Begin FPSpread.vaSpread spdRegProd 
         Height          =   7605
         Left            =   150
         TabIndex        =   4
         Top             =   330
         Width           =   4605
         _Version        =   393216
         _ExtentX        =   8123
         _ExtentY        =   13414
         _StockProps     =   64
         ColsFrozen      =   8
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
         MaxCols         =   21
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   16773087
         SpreadDesigner  =   "frmRegComp.frx":0072
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   17895
      Begin VB.TextBox txtComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5790
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   48
         Top             =   330
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "화면정리"
         Height          =   375
         Left            =   8340
         Style           =   1  '그래픽
         TabIndex        =   3
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   300
         Width           =   1245
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "조회"
         Height          =   375
         Left            =   7050
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   300
         Width           =   1245
      End
      Begin VB.ComboBox cboComp 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1350
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   330
         Width           =   4395
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "▶ 고객사 "
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
         Left            =   330
         TabIndex        =   24
         Top             =   390
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmRegProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------'
'   파일명  : frmRegProd.frm
'   작성자  : 오세원
'   내  용  : 제품코드 등록
'   작성일  : 2020-02-06
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'


Private Sub cboComp_Click()
    Dim intCnt As Integer
    
    txtComp.Text = mGetP(cboComp.Text, 2, "|")
    
    For intCnt = 0 To cboCompCd.ListCount
        If mGetP(cboCompCd.List(intCnt), 2, "|") = txtComp.Text Then
            cboCompCd.ListIndex = intCnt
        End If
    Next
End Sub

Private Sub cboCompCd_Click()
    
    txtCompCd.Text = mGetP(cboCompCd.Text, 2, "|")
    
End Sub

Private Sub cboProdMatCd_Click()
    
    txtProdMatCd.Text = mGetP(cboProdMatCd.Text, 2, "|")

End Sub

Private Sub cmdClear_Click()
        
    spdRegProd.MaxRows = 0
    
    '-- 1 Line
    'cboComp.ListIndex = 0
    txtProdCd.Text = ""
    txtCompCd.Text = ""
    txtProdNm.Text = ""
    cboCompCd.ListIndex = 0
    
    '-- 2 Line
    txtProdLen.Text = ""
    'cboProdMatCd.ListIndex = 0
    txtProdMatCd.Text = ""
    txtExprMon.Text = ""
    txtProdStorTemp.Text = ""
    txtProdSize.Text = ""
    txtProdChimPN.Text = ""
    
    '-- 3 Line
    txtVenCd.Text = "E1B4" '(KUKDO)
    cboProdLineFA.ListIndex = 0
    cboProdSlitFA.ListIndex = 0
    cboProdCtrlYN.ListIndex = 0
    cboProdPcnNo.ListIndex = 0
    txtItemBar.Text = ""
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdDelete_Click()
    
    If MsgBox("삭제하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton1, Me.Caption) = vbNo Then
        Exit Sub
    End If
    
    '필수입력 체크
    If txtProdCd.Text = "" Then
        MsgBox "제품코드를 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtProdCd.SetFocus
        Exit Sub
    End If
        
    If txtProdNm.Text = "" Then
        MsgBox "제품명을 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtProdNm.SetFocus
        Exit Sub
    End If
        
    If txtCompCd.Text = "" Then
        MsgBox "고객사를 선택하세요", vbOKOnly + vbCritical, Me.Caption
        cboCompCd.SetFocus
        Exit Sub
    End If
        
    If txtProdLen.Text = "" Then
        MsgBox "제품길이를 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtProdLen.SetFocus
        Exit Sub
    End If
        
    If txtProdMatCd.Text = "" Then
        MsgBox "자재코드를 선택하세요", vbOKOnly + vbCritical, Me.Caption
        cboProdMatCd.SetFocus
        Exit Sub
    End If
        
        
    '-- 담기
    gProd.CD = txtProdCd.Text
    gProd.NAME = txtProdNm.Text
    gProd.COMPCD = txtCompCd.Text
    gProd.LEN = txtProdLen.Text
    gProd.METCD = txtProdMatCd.Text
    gProd.MONTH = txtExprMon.Text
    gProd.TEMP = txtProdStorTemp.Text
    gProd.SIZE = txtProdSize.Text
    gProd.CHPN = txtProdChimPN.Text
    gProd.VDCD = txtVenCd.Text
    gProd.LINEFA = Trim(mGetP(cboProdLineFA.Text, 1, "|"))
    gProd.SLITFA = Trim(mGetP(cboProdSlitFA.Text, 1, "|"))
    gProd.CTYN = Trim(mGetP(cboProdCtrlYN.Text, 1, "|"))
    gProd.PCNNO = cboProdPcnNo.Text
    gProd.BAR = txtItemBar
    If chkUsedYN.Value = "1" Then
        gProd.YN = "Y"
    Else
        gProd.YN = "N"
    End If
                
    '-- Insert / Update 찾아오기
    Set AdoRs = Get_ProdList(txtProdCd.Text, txtCompCd.Text)
        
    '-- 저장
    If AdoRs.RecordCount > 0 Then
        'DELETE
        
        If Set_Prod("DEL") Then
            Call CtlInitializing
            Call GetProdList
        End If
    End If
End Sub

Private Sub cmdMakeCode_Click()
    
    Dim strMaxNum   As String
    
    'Call cmdClear_Click
    
    txtProdCd.Text = ""
    txtCompCd.Text = ""
    txtProdNm.Text = ""
    cboCompCd.ListIndex = 0
    
    txtProdLen.Text = ""
    cboProdMatCd.ListIndex = 0
    txtProdMatCd.Text = ""
    txtExprMon.Text = ""
    txtProdStorTemp.Text = ""
    txtProdSize.Text = ""
    txtProdChimPN.Text = ""
    
    txtVenCd.Text = "E1B4" '(KUKDO)
    cboProdLineFA.ListIndex = 0
    cboProdSlitFA.ListIndex = 0
    cboProdCtrlYN.ListIndex = 0
    cboProdPcnNo.ListIndex = 0
    txtItemBar.Text = ""
    
    
    strMaxNum = Get_MaxProdCode
    txtProdCd.Text = strMaxNum
    

End Sub

Private Sub cmdOK_Click()

    Call SetProd

End Sub

Private Sub SetProd()
    
    '필수입력 체크
    If txtProdCd.Text = "" Then
        MsgBox "제품코드를 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtProdCd.SetFocus
        Exit Sub
    End If
        
    If txtProdNm.Text = "" Then
        MsgBox "제품명을 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtProdNm.SetFocus
        Exit Sub
    End If
        
    If txtCompCd.Text = "" Then
        MsgBox "고객사를 선택하세요", vbOKOnly + vbCritical, Me.Caption
        cboCompCd.SetFocus
        Exit Sub
    End If
        
    If txtProdLen.Text = "" Then
        MsgBox "제품길이를 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtProdLen.SetFocus
        Exit Sub
    End If
        
    If txtProdMatCd.Text = "" Then
        MsgBox "자재코드를 선택하세요", vbOKOnly + vbCritical, Me.Caption
        cboProdMatCd.SetFocus
        Exit Sub
    End If
        
        
    '-- 담기
    gProd.CD = txtProdCd.Text
    gProd.NAME = txtProdNm.Text
    gProd.PRTNAME = txtProdPrtNm.Text
    gProd.COMPCD = txtCompCd.Text
    gProd.LEN = txtProdLen.Text
    gProd.METCD = txtProdMatCd.Text
    gProd.MONTH = txtExprMon.Text
    gProd.TEMP = txtProdStorTemp.Text
    gProd.SIZE = txtProdSize.Text
    gProd.CHPN = txtProdChimPN.Text
    gProd.VDCD = txtVenCd.Text
    gProd.LINEFA = Trim(mGetP(cboProdLineFA.Text, 1, "|"))
    gProd.SLITFA = Trim(mGetP(cboProdSlitFA.Text, 1, "|"))
    gProd.CTYN = Trim(mGetP(cboProdCtrlYN.Text, 1, "|"))
    gProd.PCNNO = cboProdPcnNo.Text
    gProd.BAR = txtItemBar
    If chkUsedYN.Value = "1" Then
        gProd.YN = "Y"
    Else
        gProd.YN = "N"
    End If
                
    '-- Insert / Update 찾아오기
    Set AdoRs = Get_ProdList(txtProdCd.Text, txtCompCd.Text)
        
    '-- 저장
    If AdoRs.RecordCount = 0 Then
        'INSERT
        If Set_Prod("IN") Then
            Call CtlInitializing
            Call GetProdList
        End If
    Else
        'UPDATE
        If Set_Prod("UP") Then
            Call CtlInitializing
            Call GetProdList
        End If
    End If
    
End Sub
   
    
Private Sub GetProdList(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
    
    Dim strCompNm       As String
    
    Set AdoRs = Get_ProdList(pProdCd, pCompCd)
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdRegProd
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdRegProd, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 1)
                Call SetText(spdRegProd, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 2)
                strCompNm = GetCompList_ViewName(AdoRs.Fields("COMP_CD").Value & "")
                Call SetText(spdRegProd, strCompNm, .MaxRows, 3)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_NAME").Value & "", .MaxRows, 4)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_LENGTH").Value & "", .MaxRows, 5)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_MATERIAL_CD").Value & "", .MaxRows, 6)
                Call SetText(spdRegProd, AdoRs.Fields("EXPIR_MONTH").Value & "", .MaxRows, 7)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_STOR_TEMP").Value & "", .MaxRows, 8)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_SIZE").Value & "", .MaxRows, 9)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_CHIMEI_PN").Value & "", .MaxRows, 10)
                Call SetText(spdRegProd, AdoRs.Fields("VENDER_CD").Value & "", .MaxRows, 11)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_LINE_FA").Value & "", .MaxRows, 12)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_SLIT_FA").Value & "", .MaxRows, 13)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_CONTROL_YN").Value & "", .MaxRows, 14)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_PCN_NO").Value & "", .MaxRows, 15)
                Call SetText(spdRegProd, AdoRs.Fields("ITEM_BARCODE").Value & "", .MaxRows, 16)
                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
                    Call SetText(spdRegProd, "1", .MaxRows, 17)
                Else
                    Call SetText(spdRegProd, "0", .MaxRows, 17)
                End If
                Call SetText(spdRegProd, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 18)
                Call SetText(spdRegProd, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 19)
                Call SetText(spdRegProd, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 20)
                Call SetText(spdRegProd, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 21)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_PRT_NAME").Value & "", .MaxRows, 22)
            End With
            
            AdoRs.MoveNext
        Loop
        
    End If
    
    AdoRs.Close

End Sub


Private Sub cmdSearch_Click()

    Call cmdClear_Click
    
    Call GetProdList("", txtComp.Text)
    
End Sub

Private Sub Form_Load()

    Call CtlInitializing
    
    '고객사 리스트 가져오기
    Call GetCompList_CodeName
    
    '자재코드 가져오기
    'Call Get_Material_CodeName
    
End Sub

Private Sub Get_Material_CodeName()

    Set AdoRs = Get_Material
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        cboProdMatCd.Clear
        Do Until AdoRs.EOF
            cboProdMatCd.AddItem AdoRs.Fields("MAT_NAME").Value & Space(20) & "|" & AdoRs.Fields("MAT_CD").Value & ""
            
            AdoRs.MoveNext
        Loop
        
        cboProdMatCd.ListIndex = 0
    End If
    
    AdoRs.Close

End Sub

Private Function GetCompList_Name(Optional ByVal pCompCd As String) As String
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_CompList_Name(pCompCd)

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until pAdoRS.EOF
            GetCompList_Name = pAdoRS.Fields("COMP_NAME").Value & ""

            pAdoRS.MoveNext
        Loop

    End If

    pAdoRS.Close

End Function

Private Function GetCompList_ViewName(Optional ByVal pCompCd As String) As String
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_CompList_ViewName(pCompCd)

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until pAdoRS.EOF
            GetCompList_ViewName = pAdoRS.Fields("COMP_VIEW").Value & "-" & pAdoRS.Fields("COMP_NAME").Value & ""

            pAdoRS.MoveNext
        Loop

    End If

    pAdoRS.Close

End Function


Private Sub GetCompList_CodeName()

    
    Set AdoRs = Get_CompList_CodeName
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        cboComp.Clear
        cboCompCd.Clear
        
        cboComp.AddItem "전체" & Space(30) & "|" & "전체"
        
        Do Until AdoRs.EOF
            cboComp.AddItem AdoRs.Fields("COMP_NAME").Value & Space(15 - Len(AdoRs.Fields("COMP_NAME").Value)) & AdoRs.Fields("COMP_LINE").Value & Space(30) & "|" & AdoRs.Fields("COMP_CD").Value & ""
            cboCompCd.AddItem AdoRs.Fields("COMP_NAME").Value & Space(15 - Len(AdoRs.Fields("COMP_NAME").Value)) & AdoRs.Fields("COMP_LINE").Value & Space(30) & "|" & AdoRs.Fields("COMP_CD").Value & ""
            
            AdoRs.MoveNext
        Loop
        
        cboComp.ListIndex = 0
        cboCompCd.ListIndex = 0
        
    End If
    
    AdoRs.Close
    
End Sub


'-- 컨트롤초기화
Private Sub CtlInitializing()
    
    Dim i   As Integer
    
    With spdRegProd
        .MaxCols = 22
        Call SetText(spdRegProd, "제품코드", 0, 1):         .ColWidth(1) = 0
        Call SetText(spdRegProd, "고객사코드", 0, 2):       .ColWidth(2) = 0
        Call SetText(spdRegProd, "고객사명", 0, 3):         .ColWidth(3) = 12
        Call SetText(spdRegProd, "제품명", 0, 4):           .ColWidth(4) = 12
        Call SetText(spdRegProd, "제품길이", 0, 5):         .ColWidth(5) = 10
        Call SetText(spdRegProd, "자재코드", 0, 6):         .ColWidth(6) = 10
        Call SetText(spdRegProd, "유효기간", 0, 7):         .ColWidth(7) = 10
        Call SetText(spdRegProd, "보관온도", 0, 8):         .ColWidth(8) = 10
        Call SetText(spdRegProd, "사이즈", 0, 9):           .ColWidth(9) = 8
        Call SetText(spdRegProd, "CHIMEI코드", 0, 10):      .ColWidth(10) = 8
        Call SetText(spdRegProd, "VENDOR코드", 0, 11):      .ColWidth(11) = 8
        Call SetText(spdRegProd, "제조라인공장", 0, 12):    .ColWidth(12) = 10
        Call SetText(spdRegProd, "SLITTING공장", 0, 13):    .ColWidth(13) = 10
        Call SetText(spdRegProd, "관리선이탈여부", 0, 14):  .ColWidth(14) = 8
        Call SetText(spdRegProd, "PCN차수", 0, 15):         .ColWidth(15) = 8
        Call SetText(spdRegProd, "바코드", 0, 16):          .ColWidth(16) = 8
        Call SetText(spdRegProd, "사용여부", 0, 17):        .ColWidth(17) = 10
        Call SetText(spdRegProd, "입력자", 0, 18):          .ColWidth(18) = 10
        Call SetText(spdRegProd, "입력일시", 0, 19):        .ColWidth(19) = 10
        Call SetText(spdRegProd, "수정자", 0, 20):          .ColWidth(20) = 10
        Call SetText(spdRegProd, "수정일시", 0, 21):        .ColWidth(21) = 10
        Call SetText(spdRegProd, "제품명(출력용)", 0, 22):        .ColWidth(22) = 0
    
        .MaxRows = 0
    End With
    
    '-- 1 Line
    'cboComp.ListIndex = 0
    txtProdCd.Text = ""
    txtProdNm.Text = ""
    txtProdPrtNm.Text = ""
    'cboCompCd.ListIndex = 0
'    txtComp.Text = ""
    txtCompCd.Text = ""
    
    '-- 2 Line
    txtProdLen.Text = ""
    'cboProdMatCd.ListIndex = 0
    txtProdMatCd.Text = ""
    txtExprMon.Text = ""
    txtProdStorTemp.Text = ""
    txtProdSize.Text = ""
    txtProdChimPN.Text = ""
    
    '-- 3 Line
    txtVenCd.Text = "" '"E1B4" '(KUKDO)
    cboProdLineFA.ListIndex = 0
    cboProdSlitFA.ListIndex = 0
    cboProdCtrlYN.ListIndex = 0
    cboProdPcnNo.Clear
    For i = 1 To 9
        cboProdPcnNo.AddItem i
    Next
    cboProdPcnNo.ListIndex = 0
    txtItemBar.Text = ""
    
    chkUsedYN.Value = "1"
    If gKUKDO.USERGRD = "1" Then
        cmdDelete.Visible = True
    Else
        cmdDelete.Visible = False
    End If
    
    gSORT = 0

End Sub


Private Sub spdRegProd_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i       As Integer

    If Row = 0 Then
        Call SetSpreadSort(spdRegProd)
        Exit Sub
    End If
    
    txtProdCd.Text = GetText(spdRegProd, Row, 1)
    For i = 0 To cboCompCd.ListCount
        If Trim(mGetP(cboCompCd.List(i), 2, "|")) = GetText(spdRegProd, Row, 2) Then
            cboCompCd.ListIndex = i
            txtCompCd.Text = Trim(mGetP(cboCompCd.List(i), 2, "|"))
            Exit For
        End If
    Next
    
    'txtCompCd.Text = GetText(spdRegProd, Row, 3)
    txtProdNm.Text = GetText(spdRegProd, Row, 4)
    txtProdPrtNm.Text = GetText(spdRegProd, Row, 22)
    txtProdLen.Text = GetText(spdRegProd, Row, 5)
'    For i = 0 To cboProdMatCd.ListCount
'        If Trim(mGetP(cboProdMatCd.List(i), 2, "|")) = GetText(spdRegProd, Row, 6) Then
'            cboProdMatCd.ListIndex = i
'            '?? 이유를 모르겠음 : index = 0 일 경우만 발생함
'            txtProdMatCd.Text = mGetP(cboProdMatCd.Text, 2, "|")
'            Exit For
'        End If
'    Next
    txtProdMatCd.Text = GetText(spdRegProd, Row, 6)
    txtExprMon.Text = GetText(spdRegProd, Row, 7)
    txtProdStorTemp.Text = GetText(spdRegProd, Row, 8)
    txtProdSize.Text = GetText(spdRegProd, Row, 9)
    txtProdChimPN.Text = GetText(spdRegProd, Row, 10)
    txtVenCd.Text = GetText(spdRegProd, Row, 11)
    '제조라인
    For i = 0 To cboProdLineFA.ListCount
        If Trim(mGetP(cboProdLineFA.List(i), 1, "|")) = GetText(spdRegProd, Row, 12) Then
            cboProdLineFA.ListIndex = i
            Exit For
        End If
    Next
    'Slitting
    For i = 0 To cboProdSlitFA.ListCount
        If Trim(mGetP(cboProdSlitFA.List(i), 1, "|")) = GetText(spdRegProd, Row, 13) Then
            cboProdSlitFA.ListIndex = i
            Exit For
        End If
    Next
    '관리선이탈
    For i = 0 To cboProdCtrlYN.ListCount
        If Trim(mGetP(cboProdCtrlYN.List(i), 1, "|")) = GetText(spdRegProd, Row, 14) Then
            cboProdCtrlYN.ListIndex = i
            Exit For
        End If
    Next
    'PCN차수
    For i = 0 To cboProdPcnNo.ListCount
        If Trim(mGetP(cboProdPcnNo.List(i), 1, "|")) = GetText(spdRegProd, Row, 15) Then
            cboProdPcnNo.ListIndex = i
            Exit For
        End If
    Next
    txtItemBar.Text = GetText(spdRegProd, Row, 16)
    
    If GetText(spdRegProd, Row, 17) = "1" Then
        chkUsedYN.Value = "1"
    Else
        chkUsedYN.Value = "0"
    End If
End Sub

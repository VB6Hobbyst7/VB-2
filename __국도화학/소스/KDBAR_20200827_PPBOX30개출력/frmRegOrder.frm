VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmRegOrder 
   BackColor       =   &H00FFFFFF&
   Caption         =   "작업지시서 등록"
   ClientHeight    =   12165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12165
   ScaleWidth      =   20580
   WindowState     =   2  '최대화
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   60
      TabIndex        =   30
      Top             =   60
      Width           =   18765
      Begin VB.CommandButton cmdSearch 
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
         Left            =   5610
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "화면정리"
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
         Left            =   6750
         Style           =   1  '그래픽
         TabIndex        =   4
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   360
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   375
         Left            =   1650
         TabIndex        =   1
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   127926273
         CurrentDate     =   43884
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   375
         Left            =   3750
         TabIndex        =   2
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   127926273
         CurrentDate     =   43884
      End
      Begin VB.Label Label1 
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
         Left            =   3450
         TabIndex        =   34
         Top             =   420
         Width           =   195
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "▶ 생산일자 "
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
         TabIndex        =   31
         Top             =   390
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   9675
      Left            =   60
      TabIndex        =   0
      Top             =   1050
      Width           =   18765
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   8745
         Left            =   8160
         TabIndex        =   17
         Top             =   300
         Width           =   10335
         Begin VB.CommandButton cmdAdd 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "(+) 추가"
            Height          =   405
            Left            =   1050
            Style           =   1  '그래픽
            TabIndex        =   43
            Top             =   5220
            Width           =   1395
         End
         Begin VB.CommandButton cmdRemove 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Caption         =   "(-) 제거"
            Height          =   405
            Left            =   1050
            Style           =   1  '그래픽
            TabIndex        =   42
            Top             =   5670
            Width           =   1395
         End
         Begin VB.TextBox txtNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            Enabled         =   0   'False
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
            Left            =   7320
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   39
            Top             =   810
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox txtLotNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2490
            MaxLength       =   10
            TabIndex        =   37
            Top             =   360
            Width           =   3720
         End
         Begin VB.TextBox txtCompNm 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   6270
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   36
            Top             =   2610
            Width           =   3720
         End
         Begin VB.TextBox txtPackInfo 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   6270
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   35
            Top             =   2160
            Width           =   3720
         End
         Begin VB.ComboBox cboSlittingNo 
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
            ItemData        =   "frmRegOrder.frx":0000
            Left            =   2490
            List            =   "frmRegOrder.frx":0002
            Style           =   2  '드롭다운 목록
            TabIndex        =   11
            Top             =   3060
            Width           =   3735
         End
         Begin VB.TextBox txtReelQTY 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
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
            Left            =   2490
            MaxLength       =   10
            TabIndex        =   13
            Top             =   7440
            Width           =   2310
         End
         Begin VB.TextBox txtOrderMemo 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1185
            Left            =   2490
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   3510
            Width           =   7470
         End
         Begin MSComCtl2.DTPicker dtpProdOrderDt 
            Height          =   375
            Left            =   2490
            TabIndex        =   6
            Top             =   810
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   127926273
            CurrentDate     =   43884
         End
         Begin VB.ComboBox cboCompCd 
            BeginProperty Font 
               Name            =   "돋움체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2490
            Style           =   2  '드롭다운 목록
            TabIndex        =   10
            Top             =   2610
            Width           =   3735
         End
         Begin VB.ComboBox cboProdCd 
            BackColor       =   &H00C0FFFF&
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
            ItemData        =   "frmRegOrder.frx":0004
            Left            =   2490
            List            =   "frmRegOrder.frx":0006
            Style           =   2  '드롭다운 목록
            TabIndex        =   7
            Top             =   1260
            Width           =   2085
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   795
            Left            =   6510
            TabIndex        =   20
            Top             =   7830
            Width           =   3525
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
               TabIndex        =   16
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
               TabIndex        =   14
               Top             =   150
               Width           =   1095
            End
            Begin VB.CommandButton cmdDelete 
               BackColor       =   &H00E0E0E0&
               Caption         =   "삭제"
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
               TabIndex        =   15
               Top             =   150
               Width           =   1095
            End
         End
         Begin VB.ComboBox cboPackCd 
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
            ItemData        =   "frmRegOrder.frx":0008
            Left            =   2490
            List            =   "frmRegOrder.frx":000A
            Style           =   2  '드롭다운 목록
            TabIndex        =   9
            Top             =   2160
            Width           =   3735
         End
         Begin VB.ComboBox cboProdPosNo1 
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
            ItemData        =   "frmRegOrder.frx":000C
            Left            =   2490
            List            =   "frmRegOrder.frx":000E
            Style           =   2  '드롭다운 목록
            TabIndex        =   8
            Top             =   1710
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.TextBox txtProdCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
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
            Left            =   4590
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   19
            Top             =   1260
            Width           =   1605
         End
         Begin VB.TextBox txtProdLen 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
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
            Left            =   7320
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   18
            Top             =   1260
            Width           =   1140
         End
         Begin FPSpread.vaSpread spdRegOrderDetail 
            Height          =   2475
            Left            =   2490
            TabIndex        =   41
            Top             =   4770
            Width           =   7455
            _Version        =   393216
            _ExtentX        =   13150
            _ExtentY        =   4366
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
            MaxCols         =   19
            MaxRows         =   20
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            ScrollBars      =   2
            ShadowColor     =   16774120
            SpreadDesigner  =   "frmRegOrder.frx":0010
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin VB.Label Label2 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "개 (ea)"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   4770
            TabIndex        =   44
            Top             =   7500
            Width           =   975
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "일련번호"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   6270
            TabIndex        =   40
            Top             =   810
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "투입Roll"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   38
            Top             =   4770
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "Reel 수량"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   33
            Top             =   7440
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "제품명"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   32
            Top             =   1260
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "고객사"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   29
            Top             =   2610
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "Slitting 작업번호"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   15
            Left            =   240
            TabIndex        =   28
            Top             =   3060
            Width           =   2205
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "공정 No"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   1710
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "제품길이"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   6270
            TabIndex        =   26
            Top             =   1260
            Width           =   1020
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "메모"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   25
            Top             =   3510
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "생산LotNo"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "제조일자"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   810
            Width           =   2205
         End
         Begin VB.Label lblWorkDate 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "미터(M)"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   8520
            TabIndex        =   22
            Top             =   1350
            Width           =   975
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "포장코드"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   21
            Top             =   2160
            Width           =   2205
         End
      End
      Begin FPSpread.vaSpread spdRegOrder 
         Height          =   8625
         Left            =   210
         TabIndex        =   5
         Top             =   390
         Width           =   7875
         _Version        =   393216
         _ExtentX        =   13891
         _ExtentY        =   15214
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
         MaxCols         =   19
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16774120
         SpreadDesigner  =   "frmRegOrder.frx":0F8B
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
End
Attribute VB_Name = "frmRegOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------'
'   파일명  : frmRegBar.frm
'   작성자  : 오세원
'   내  용  : 작업지시서 등록
'   작성일  : 2020-02-23
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'

Private Sub cboCompCd_Click()
    Dim strCompCd   As String
    Dim strLotNo As String
    Dim strDate  As String


    strCompCd = Trim(mGetP(cboCompCd.Text, 2, "|"))

   ' Call GetProdList_CodeName_Reg("", strCompCd)
    
    txtCompNm.Text = Trim(mGetP(mGetP(cboCompCd.Text, 1, "|"), 2, ":"))

    strLotNo = GetLotNo(Mid(dtpProdOrderDt.Value, 1, 10), cboSlittingNo.Text, Mid(cboPackCd, 1, 2), Mid(cboCompCd.Text, 1, 2))
    txtLotNo.Text = strLotNo

End Sub

Private Sub cboPackCd_Click()
    Dim strLotNo As String
    Dim strDate  As String

    txtPackInfo.Text = Trim(mGetP(cboPackCd.Text, 2, Space(3)))
    
    strLotNo = GetLotNo(Mid(dtpProdOrderDt.Value, 1, 10), cboSlittingNo.Text, Mid(cboPackCd, 1, 2), Mid(cboCompCd.Text, 1, 2))
    txtLotNo.Text = strLotNo
    
End Sub

Private Sub cboProdCd_Click()
    Dim strCompCd    As String
    
    txtProdCd.Text = Trim(mGetP(cboProdCd.Text, 2, "|"))
    strCompCd = Trim(mGetP(cboCompCd.Text, 2, "|"))
    
    Call GetComp_CodeName(txtProdCd.Text)
    
    spdRegOrderDetail.MaxRows = 0
    
End Sub

Private Sub cboSlittingNo_Click()
    Dim strLotNo As String
    Dim strDate  As String
    
    strLotNo = GetLotNo(Mid(dtpProdOrderDt.Value, 1, 10), cboSlittingNo.Text, Mid(cboPackCd, 1, 2), Mid(cboCompCd.Text, 1, 2))
    txtLotNo.Text = strLotNo
    
End Sub

Private Sub cmdAdd_Click()
    Dim pAdoRS      As ADODB.Recordset
    Dim intRow      As Integer
    Dim intNum      As Integer
    Dim intMaxNum   As Integer
    
    With spdRegOrderDetail
        .MaxRows = .MaxRows + 1
        Call SetText(spdRegOrderDetail, dtpProdOrderDt.Value, .MaxRows, 1)
        Call SetText(spdRegOrderDetail, cboSlittingNo.Text, .MaxRows, 2)
        Call SetText(spdRegOrderDetail, txtProdCd.Text, .MaxRows, 3)
        Call SetText(spdRegOrderDetail, cboSlittingNo.Text, .MaxRows, 4)
        Call SetText(spdRegOrderDetail, CStr(.MaxRows), .MaxRows, 5)
        
        Call SetText(spdRegOrderDetail, "", .MaxRows, 6)
        .Row = .MaxRows
        .Col = 6
        .CellType = CellTypeEdit
        .TypeMaxEditLen = 300
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter

'        Call SetText(spdRegOrderDetail, "P" & CStr(.MaxRows), .MaxRows, 7)
'        .Row = .MaxRows
'        .Col = 7
'        .CellType = CellTypeEdit
'        .TypeMaxEditLen = 2
'        .TypeHAlign = TypeHAlignCenter
'        .TypeVAlign = TypeVAlignCenter

        Call SetText(spdRegOrderDetail, "", .MaxRows, 7)
        .Row = .MaxRows
        .Col = 7
        .CellType = CellTypeEdit
        .TypeMaxEditLen = 4
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        
        Call SetText(spdRegOrderDetail, "", .MaxRows, 8)
        .Row = .MaxRows
        .Col = 8
        .CellType = CellTypeEdit
        .TypeMaxEditLen = 4
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
    
        '-- 2020-03-17 추가
        .Action = ActionActiveCell
    End With
    
    


End Sub

Private Sub cmdClear_Click()
    Dim i   As Integer
    
    spdRegOrder.MaxRows = 0
    spdRegOrderDetail.MaxRows = 0

    dtpFromDate.Value = Now - 1
    dtpToDate.Value = Now

'    txtRoolInfo.Text = ""
    dtpProdOrderDt.Value = Now
    txtNo.Text = "1"
    
    'cboProdPosNo.Clear
    cboSlittingNo.Clear
    For i = 1 To 10
    '    cboProdPosNo.AddItem CStr(i)
        cboSlittingNo.AddItem CStr(i)
    Next
    'cboProdPosNo.ListIndex = 0
    cboSlittingNo.ListIndex = 0
    
    txtPackInfo.Text = ""
    txtCompNm.Text = ""
    
    txtOrderMemo.Text = ""
    txtLotNo.Text = ""
    txtReelQTY.Text = "0"
    
    
    '고객사 리스트 가져오기
    Call GetCompList_CodeName
    
    '제품 리스트 가져오기
    Call GetProdList_CodeName("", "")
    
    ' 포장코드 리스트 가져오기
    Call GetPackList
    
    txtLotNo.Text = ""
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdDelete_Click()
    '필수입력 체크

    If MsgBox(txtLotNo.Text & " 항목을 삭제하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton1, Me.Caption) = vbNo Then
        Exit Sub
    End If
    
    If txtLotNo.Text = "" Then
        MsgBox "제품을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        txtLotNo.SetFocus
        Exit Sub
    End If
    
    If txtProdCd.Text = "" Or txtProdCd.Text = "전체" Then
        MsgBox "제품코드를 선택하세요", vbOKOnly + vbCritical, Me.Caption
        cboProdCd.SetFocus
        Exit Sub
    End If

    '-- 담기
    gOrder.ORDDATE = Format(dtpProdOrderDt.Value, "yyyymmdd")   'Key
    'gOrder.PRODPOSNO = cboProdPosNo.Text                        'Key
    gOrder.PRODCD = txtProdCd.Text                              'Key
    gOrder.SLITINGNO = cboSlittingNo.Text                       'Key
    
    gOrderDetail.ORDDATE = Format(dtpProdOrderDt.Value, "yyyymmdd")     'Key
'    gOrderDetail.PRODPOSNO = cboProdPosNo.Text                          'Key
    gOrderDetail.PRODCD = txtProdCd.Text                                'Key
    gOrderDetail.SLITINGNO = cboSlittingNo.Text                         'Key
    
    'INSERT
    If Set_Order("DEL") Then
        If Set_Order_Detail("DEL") Then
            Call cmdSearch_Click
        End If
    End If

End Sub

Private Sub cmdOK_Click()

    Call SetOrder
    
End Sub

Private Sub SetOrder()
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim intItemNo   As Integer
    Dim strLotNo    As String
    
    '필수입력 체크
    If txtLotNo.Text = "" Then
        MsgBox "LotNo를 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtLotNo.SetFocus
        Exit Sub
    End If

    If spdRegOrderDetail.MaxRows = 0 Then
        MsgBox "투입Roll 정보를 등록하세요", vbOKOnly + vbCritical, Me.Caption
        spdRegOrderDetail.SetFocus
        Exit Sub
    End If
    
    '-- 담기
    gOrder.ORDDATE = Format(dtpProdOrderDt.Value, "yyyymmdd")   'Key
'    gOrder.PRODPOSNO = cboProdPosNo.Text                        'Key
    gOrder.PRODCD = txtProdCd.Text                              'Key
    gOrder.SLITINGNO = cboSlittingNo.Text                       'Key
    
    'gOrder.NO = txtNo.Text
    gOrder.COMPCD = Trim(mGetP(cboCompCd.Text, 2, "|"))
    gOrder.PRODNAME = Trim(mGetP(cboProdCd.Text, 1, "|"))
    gOrder.PACKCD = Mid(cboPackCd.Text, 1, 2)
    gOrder.REELQTY = txtReelQTY.Text
    gOrder.ORDERMEMO = txtOrderMemo.Text

    'strLotNo = GetLotNo(gOrder.ORDDATE, gOrder.SLITINGNO, gOrder.PACKCD, Mid(cboCompCd.Text, 1, 2))
    
    gOrder.LOTNO = txtLotNo.Text
    gOrder.CLOSEYN = "N"
    
    With spdRegOrderDetail
        gOrderDetail.ORDDATE = Format(dtpProdOrderDt.Value, "yyyymmdd")    'Key
'        gOrderDetail.PRODPOSNO = cboProdPosNo.Text      'Key
        gOrderDetail.PRODCD = txtProdCd.Text            'Key
        gOrderDetail.SLITINGNO = cboSlittingNo.Text     'Key
        ReDim gOrderDetail.NO(.MaxRows) As String       'Key
        ReDim gOrderDetail.SLTINFO(.MaxRows) As String
        ReDim gOrderDetail.PFROMNO(.MaxRows) As String
        ReDim gOrderDetail.PTONO(.MaxRows) As String
        
        For intRow = 1 To .DataRowCnt
            gOrderDetail.NO(intRow) = GetText(spdRegOrderDetail, intRow, 5)
            gOrderDetail.SLTINFO(intRow) = GetText(spdRegOrderDetail, intRow, 6)
            gOrderDetail.PFROMNO(intRow) = GetText(spdRegOrderDetail, intRow, 7)
            gOrderDetail.PTONO(intRow) = GetText(spdRegOrderDetail, intRow, 8)
        Next
    End With
    
    '-- Insert / Update 찾아오기
    Set AdoRs = Get_Order(gOrder.ORDDATE, gOrder.PRODCD, gOrder.SLITINGNO)
        
    If AdoRs.RecordCount = 0 Then
        'INSERT
        If Set_Order("IN") Then
            '상세내용 저장
            For intRow = 1 To spdRegOrderDetail.DataRowCnt
                If Set_Order_Detail("IN", intRow) Then
                End If
            Next
        End If
        Call cmdSearch_Click
    Else
        'UPDATE
        If Set_Order("UP") Then
            '상세내용 저장
            If Set_Order_Detail("DEL", intRow) Then
                '상세내용 저장
                For intRow = 1 To spdRegOrderDetail.DataRowCnt
                    If Set_Order_Detail("IN", intRow) Then
                    End If
                Next
            End If
        End If
        Call cmdSearch_Click
    End If
    
End Sub


    
'제품 리스트 가져오기(조회용)
Private Sub GetProdList_CodeName(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
    Dim pAdoRS      As ADODB.Recordset

    Set pAdoRS = New ADODB.Recordset

    Set pAdoRS = Get_ProdList_CodeName(pProdCd, pCompCd)

    cboProdCd.Clear

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        cboProdCd.AddItem "전체" & Space(50) & "|전체"

        Do Until pAdoRS.EOF
            cboProdCd.AddItem pAdoRS.Fields("PROD_NAME").Value & "-" & pAdoRS.Fields("PROD_LENGTH").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
            pAdoRS.MoveNext
        Loop

        If pAdoRS.RecordCount > 0 Then
            cboProdCd.ListIndex = 0
        End If
    End If

    pAdoRS.Close

End Sub
    
    
'제품 리스트 가져오기(등록용)
Private Sub GetProdList_CodeName_Reg(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_ProdList_CodeName(pProdCd, pCompCd)
    
    cboProdCd.Clear
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until pAdoRS.EOF
            cboProdCd.AddItem pAdoRS.Fields("PROD_NAME").Value & "-" & pAdoRS.Fields("PROD_LENGTH").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
            pAdoRS.MoveNext
        Loop
                            
        If pAdoRS.RecordCount > 0 Then
            cboProdCd.ListIndex = 0
        End If
    End If
    
    pAdoRS.Close
    
End Sub
    
' 작업지시서 리스트 가져옴
Private Sub GetOrderList(ByVal pOrderFromDate As String, ByVal pOrderToDate As String, Optional ByVal pProdCd As String, Optional ByVal pOrderNo As String)
    
    Dim strLabelType    As String
    
    Set AdoRs = Get_OrderList(pOrderFromDate, pOrderToDate, pProdCd, pOrderNo, "R")
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdRegOrder
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdRegOrder, AdoRs.Fields("LOT_NO").Value & "", .MaxRows, 1)
                Call SetText(spdRegOrder, AdoRs.Fields("PROD_ORDER_DT").Value & "", .MaxRows, 2)
'                Call SetText(spdRegOrder, AdoRs.Fields("PROD_POS_NO").Value & "", .MaxRows, 3)
                Call SetText(spdRegOrder, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 4)
                Call SetText(spdRegOrder, AdoRs.Fields("PROD_NAME").Value & "", .MaxRows, 5)
                Call SetText(spdRegOrder, AdoRs.Fields("PACK_CD").Value & "", .MaxRows, 6)
                Call SetText(spdRegOrder, AdoRs.Fields("ORDER_MEMO").Value & "", .MaxRows, 7)
                'Call SetText(spdRegOrder, AdoRs.Fields("JOB_INFO").Value & "", .MaxRows, 8)
                Call SetText(spdRegOrder, AdoRs.Fields("PROD_LENGTH").Value & "", .MaxRows, 9)
                Call SetText(spdRegOrder, AdoRs.Fields("SLITING_NO").Value & "", .MaxRows, 10)
                Call SetText(spdRegOrder, AdoRs.Fields("REEL_QTY").Value & "", .MaxRows, 11)
                'Call SetText(spdRegOrder, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 12)
                Call SetText(spdRegOrder, AdoRs.Fields("COMP_VIEW").Value & Space(10) & "|" & AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 12)
                'Call SetText(spdRegOrder,  "", .MaxRows, 13)
                Call SetText(spdRegOrder, AdoRs.Fields("CLOSE_YN").Value & "", .MaxRows, 14)
                Call SetText(spdRegOrder, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 16)
                Call SetText(spdRegOrder, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 17)
                Call SetText(spdRegOrder, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 18)
                Call SetText(spdRegOrder, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 19)
            
                .Row = .MaxRows
                .Col = 12
                .CellType = CellTypeStaticText
                .TypeMaxEditLen = 10
                .TypeHAlign = TypeHAlignLeft
                .TypeVAlign = TypeVAlignCenter
            End With
            AdoRs.MoveNext
        Loop
    End If
    AdoRs.Close
    
End Sub


Private Sub cmdRemove_Click()
    
    If spdRegOrderDetail.MaxRows <= 0 Then
        Exit Sub
    End If
    
    If MsgBox(GetText(spdRegOrderDetail, spdRegOrderDetail.ActiveRow, 6) & " 항목을 삭제하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton1, Me.Caption) = vbYes Then
    
        If spdRegOrderDetail.MaxRows > 0 Then
            Call DeleteRow(spdRegOrderDetail, spdRegOrderDetail.ActiveRow, spdRegOrderDetail.ActiveRow)
            spdRegOrderDetail.MaxRows = spdRegOrderDetail.MaxRows - 1
        End If
    End If
    
    DoEvents
    
    Call spdRegOrderDetail_KeyPress(vbKeyReturn)


End Sub

Private Sub cmdSearch_Click()
    Dim strFromDt    As String
    Dim strToDt      As String
    
    strFromDt = Format(dtpFromDate, "yyyymmdd")
    strToDt = Format(dtpToDate, "yyyymmdd")
    
    Call cmdClear_Click
    
    Call GetOrderList(strFromDt, strToDt)
    
    
    '고객사 리스트 가져오기
    Call GetCompList_CodeName
    
    '제품 리스트 가져오기
    Call GetProdList_CodeName("", "")
    
    ' 포장코드 리스트 가져오기
    Call GetPackList

    
End Sub



Private Sub dtpProdOrderDt_Change()
    Dim strLotNo As String
    Dim strDate  As String
    
    strLotNo = GetLotNo(Mid(dtpProdOrderDt.Value, 1, 10), cboSlittingNo.Text, Mid(cboPackCd, 1, 2), Mid(cboCompCd.Text, 1, 2))
    txtLotNo.Text = strLotNo

End Sub


Private Sub Form_Load()

    Call CtlInitializing
    
    '고객사 리스트 가져오기
    Call GetCompList_CodeName
    
    '제품 리스트 가져오기
    Call GetProdList_CodeName("", "")
    
    ' 포장코드 리스트 가져오기
    Call GetPackList
    
'    txtLotNo.Text = ""
    
End Sub

Private Sub GetPackList()
    Dim pAdoRS      As ADODB.Recordset
    Dim strPackInfo As String
    
    Set pAdoRS = Get_PackList
    
    cboPackCd.Clear
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until pAdoRS.EOF
            ' PACK_CAT_WIDTH,PACK_PRO_WIDTH,PACK_PRO_LENGTH
            strPackInfo = pAdoRS.Fields("PACK_CORE").Value & "x" & pAdoRS.Fields("PACK_DIA").Value & " " & pAdoRS.Fields("PACK_CAT_WIDTH").Value & " " & pAdoRS.Fields("PACK_PRO_WIDTH").Value
            
            cboPackCd.AddItem pAdoRS.Fields("PACK_NAME").Value & Space(3) & strPackInfo & Space(20) & "|" & pAdoRS.Fields("PACK_CD").Value & Space(3)
            pAdoRS.MoveNext
        Loop
        
    End If
    
    If pAdoRS.RecordCount > 0 Then
        cboPackCd.ListIndex = 0
    End If
    
    pAdoRS.Close

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

'-- 상단 고객사리스트 가져오기
Private Sub GetCompList_CodeName()
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_CompList_CodeName
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        cboCompCd.Clear
        
        Do Until pAdoRS.EOF
            cboCompCd.AddItem pAdoRS.Fields("COMP_VIEW").Value & Space(1) & ":" & Space(1) & pAdoRS.Fields("COMP_NAME").Value & Space(15 - Len(pAdoRS.Fields("COMP_NAME").Value)) & pAdoRS.Fields("COMP_LINE").Value & Space(20) & "|" & pAdoRS.Fields("COMP_CD").Value & ""
            
            pAdoRS.MoveNext
        Loop
        
        If pAdoRS.RecordCount > 0 Then
            cboCompCd.ListIndex = 0
        End If
    End If
    
    pAdoRS.Close
    
End Sub

'-- 제품선택했을때 해당 고객사 가져오기
Private Sub GetComp_CodeName(ByVal pProdCd As String)
    Dim pAdoRS      As ADODB.Recordset
    Dim i           As Integer
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_Comp_CodeName(pProdCd)
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        txtProdLen.Text = ""
        
        Do Until pAdoRS.EOF
            txtProdLen.Text = pAdoRS.Fields("PROD_LENGTH").Value & ""
            For i = 0 To cboCompCd.ListCount
                If pAdoRS.Fields("COMP_CD").Value & "" = mGetP(cboCompCd.List(i), 2, "|") Then
                    cboCompCd.ListIndex = i
                    Exit For
                End If
            Next
            pAdoRS.MoveNext
        Loop
    
        pAdoRS.Close
        
    End If
    
    
End Sub


'-- 컨트롤초기화
Private Sub CtlInitializing()
    Dim i           As Integer
    
    With spdRegOrder
        Call SetText(spdRegOrder, "Lot No", 0, 1):            .ColWidth(1) = 11 '생산LOT
        Call SetText(spdRegOrder, "제조일자", 0, 2):          .ColWidth(2) = 10 '생산일자
        Call SetText(spdRegOrder, "공정No", 0, 3):            .ColWidth(3) = 0
        Call SetText(spdRegOrder, "제품코드", 0, 4):          .ColWidth(4) = 0
        Call SetText(spdRegOrder, "제품명", 0, 5):            .ColWidth(5) = 11 '제품명
        Call SetText(spdRegOrder, "포장코드", 0, 6):          .ColWidth(6) = 0
        Call SetText(spdRegOrder, "메모", 0, 7):              .ColWidth(7) = 0
        Call SetText(spdRegOrder, "작업내용설명", 0, 8):      .ColWidth(8) = 0 'Roll정보
        Call SetText(spdRegOrder, "길이", 0, 9):              .ColWidth(9) = 8
        Call SetText(spdRegOrder, "SLT No", 0, 10):           .ColWidth(10) = 7
        Call SetText(spdRegOrder, "작업수량", 0, 11):         .ColWidth(11) = 7
        Call SetText(spdRegOrder, "고객사", 0, 12):           .ColWidth(12) = 7
        Call SetText(spdRegOrder, "일련번호", 0, 13):         .ColWidth(13) = 0
        Call SetText(spdRegOrder, "작업완료여부", 0, 14):     .ColWidth(14) = 0
        Call SetText(spdRegOrder, "사용여부", 0, 15):         .ColWidth(15) = 0
        Call SetText(spdRegOrder, "입력자", 0, 16):           .ColWidth(16) = 0
        Call SetText(spdRegOrder, "입력일시", 0, 17):         .ColWidth(17) = 0
        Call SetText(spdRegOrder, "수정자", 0, 18):           .ColWidth(18) = 0
        Call SetText(spdRegOrder, "수정일시", 0, 19):         .ColWidth(19) = 0
    
        .MaxRows = 0
    End With
    
'        Call SetText(spdRegOrder, "Lot No", 0, 13):           .ColWidth(13) = 10 '생산LOT
'        Call SetText(spdRegOrder, "제조일자", 0, 1):          .ColWidth(1) = 10 '생산일자
'        Call SetText(spdRegOrder, "공정No", 0, 6):            .ColWidth(6) = 0
'        Call SetText(spdRegOrder, "제품코드", 0, 2):          .ColWidth(2) = 0
'        Call SetText(spdRegOrder, "제품명", 0, 3):            .ColWidth(3) = 10 '제품명
'        Call SetText(spdRegOrder, "포장코드", 0, 7):          .ColWidth(7) = 0
'        Call SetText(spdRegOrder, "메모", 0, 12):             .ColWidth(12) = 0
'        Call SetText(spdRegOrder, "작업내용설명", 0, 9):      .ColWidth(9) = 0 'Roll정보
'        Call SetText(spdRegOrder, "길이", 0, 4):              .ColWidth(4) = 8
'        Call SetText(spdRegOrder, "SLT No", 0, 10):           .ColWidth(10) = 8
'        Call SetText(spdRegOrder, "작업수량", 0, 8):          .ColWidth(8) = 8
'        Call SetText(spdRegOrder, "고객사", 0, 11):           .ColWidth(11) = 8
'        Call SetText(spdRegOrder, "일련번호", 0, 5):          .ColWidth(5) = 0
'        Call SetText(spdRegOrder, "작업완료여부", 0, 14):     .ColWidth(14) = 0
'        Call SetText(spdRegOrder, "사용여부", 0, 15):         .ColWidth(15) = 0
'        Call SetText(spdRegOrder, "입력자", 0, 16):           .ColWidth(16) = 0
'        Call SetText(spdRegOrder, "입력일시", 0, 17):         .ColWidth(17) = 0
'        Call SetText(spdRegOrder, "수정자", 0, 18):           .ColWidth(18) = 0
'        Call SetText(spdRegOrder, "수정일시", 0, 19):         .ColWidth(19) = 0
    
    With spdRegOrderDetail
        Call SetText(spdRegOrderDetail, "제조일자", 0, 1):        .ColWidth(1) = 0
        Call SetText(spdRegOrderDetail, "순번", 0, 2):            .ColWidth(2) = 0
        Call SetText(spdRegOrderDetail, "제품코드", 0, 3):        .ColWidth(3) = 0
        Call SetText(spdRegOrderDetail, "SLT No", 0, 4):          .ColWidth(4) = 0
        Call SetText(spdRegOrderDetail, "일련번호", 0, 5):        .ColWidth(5) = 8
        Call SetText(spdRegOrderDetail, "SLT내용", 0, 6):         .ColWidth(6) = 28
'        Call SetText(spdRegOrderDetail, "P No", 0, 7):            .ColWidth(7) = 4
        Call SetText(spdRegOrderDetail, "시작번호", 0, 7):        .ColWidth(7) = 10
        Call SetText(spdRegOrderDetail, "끝번호", 0, 8):          .ColWidth(8) = 10
        Call SetText(spdRegOrderDetail, "", 0, 9):                .ColWidth(9) = 0
        Call SetText(spdRegOrderDetail, "", 0, 10):               .ColWidth(10) = 0
        Call SetText(spdRegOrderDetail, "", 0, 11):               .ColWidth(11) = 0
        Call SetText(spdRegOrderDetail, "", 0, 12):               .ColWidth(12) = 0
        Call SetText(spdRegOrderDetail, "", 0, 13):               .ColWidth(13) = 0
        Call SetText(spdRegOrderDetail, "", 0, 14):               .ColWidth(14) = 0
        Call SetText(spdRegOrderDetail, "사용여부", 0, 15):       .ColWidth(15) = 0
        Call SetText(spdRegOrderDetail, "입력자", 0, 16):         .ColWidth(16) = 0
        Call SetText(spdRegOrderDetail, "입력일시", 0, 17):       .ColWidth(17) = 0
        Call SetText(spdRegOrderDetail, "수정자", 0, 18):         .ColWidth(18) = 0
        Call SetText(spdRegOrderDetail, "수정일시", 0, 19):       .ColWidth(19) = 0
    
        .MaxRows = 0
    End With
    
    
    dtpFromDate.Value = Now '- 1
    dtpToDate.Value = Now

'    txtRoolInfo.Text = ""
    dtpProdOrderDt.Value = Now
    txtNo.Text = "1"

    'cboProdPosNo.Clear
    cboSlittingNo.Clear
    For i = 1 To 10
        'cboProdPosNo.AddItem CStr(i)
        cboSlittingNo.AddItem CStr(i)
    Next
    'cboProdPosNo.ListIndex = 0
    cboSlittingNo.ListIndex = 0
    
    txtPackInfo.Text = ""
    txtCompNm.Text = ""
    
    txtOrderMemo.Text = ""
    txtLotNo.Text = ""
    txtReelQTY.Text = "0"
    
    gSORT = 0

End Sub

Private Sub spdRegOrder_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i               As Integer
    Dim strDate         As String
    Dim strLotNo        As String
    Dim strProdPosNo    As String
    Dim strProdCd       As String
    Dim strSltNo        As String
    
    If Row = 0 Then
        Call SetSpreadSort(spdRegOrder)
        Exit Sub
    End If
    
    spdRegOrderDetail.MaxRows = 0
    
    txtLotNo.Text = GetText(spdRegOrder, Row, 1)
    strDate = GetText(spdRegOrder, Row, 2)
    dtpProdOrderDt.Value = Format(strDate, "####-##-##")
    'cboProdPosNo.Text = GetText(spdRegOrder, Row, 3)
'    strProdPosNo = cboProdPosNo.Text
    
    For i = 0 To cboCompCd.ListCount
        If Trim(mGetP(cboProdCd.List(i), 2, "|")) = GetText(spdRegOrder, Row, 4) Then
            cboProdCd.ListIndex = i
            strProdCd = Trim(mGetP(cboProdCd.List(i), 2, "|"))
            Exit For
        End If
    Next
    For i = 0 To cboPackCd.ListCount
        If Mid(cboPackCd.List(i), 1, 2) = GetText(spdRegOrder, Row, 6) Then
            cboPackCd.ListIndex = i
            Exit For
        End If
    Next
    txtOrderMemo.Text = GetText(spdRegOrder, Row, 7)
    txtProdLen.Text = GetText(spdRegOrder, Row, 9)
    cboSlittingNo.Text = GetText(spdRegOrder, Row, 10)
    strSltNo = cboSlittingNo.Text
    txtReelQTY.Text = GetText(spdRegOrder, Row, 11)
    
    '고객사
    For i = 0 To cboCompCd.ListCount
        If Trim(mGetP(cboCompCd.List(i), 2, "|")) = Trim(mGetP(GetText(spdRegOrder, Row, 12), 2, "|")) Then
            cboCompCd.ListIndex = i
            Exit For
        End If
    Next
    
    Call GetOrderDetail(strDate, strProdCd, strSltNo)
    
    
End Sub


' 작업지시서 리스트 가져옴 'strDate, cboProdPosNo.Text, cboProdCd.Text, cboSlittingNo.Text
Private Sub GetOrderDetail(ByVal pDate As String, ByVal pProCd As String, ByVal pSltNo As String)
    
    Set AdoRs = Get_OrderDetail(pDate, pProCd, pSltNo)
            
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        spdRegOrderDetail.MaxRows = 0
        
        Do Until AdoRs.EOF
            With spdRegOrderDetail
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_ORDER_DT").Value & "", .MaxRows, 1)
                'Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_POS_NO").Value & "", .MaxRows, 2)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 3)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("SLITING_NO").Value & "", .MaxRows, 4)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("SEQ_NO").Value & "", .MaxRows, 5)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("SLITING_INFO").Value & "", .MaxRows, 6)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("P_NO_F").Value & "", .MaxRows, 7)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("P_NO_T").Value & "", .MaxRows, 8)
  
            End With
            
            AdoRs.MoveNext
        Loop
        AdoRs.Close
    End If

End Sub

Private Sub spdRegOrderDetail_KeyPress(KeyAscii As Integer)
    Dim i           As Integer
    Dim strColNum   As String
    Dim intFrNum    As Integer
    Dim intToNum    As Integer
    Dim intLineSum  As Integer
    Dim intSum  As Integer
    
    With spdRegOrderDetail
        If KeyAscii = vbKeyReturn Then
            If .ActiveCol = 7 Or .ActiveCol = 8 Then
                For i = 1 To .MaxRows
                    intLineSum = 0
                    intFrNum = 0
                    intToNum = 0
                    .Row = i
                    .Col = 7
                    strColNum = .Text
                    If strColNum <> "" Then
                        If IsNumeric(Mid(strColNum, 3)) Then
                            intFrNum = Mid(strColNum, 3)
                        End If
                    End If
                    
                    .Row = i
                    .Col = 8
                    strColNum = .Text
                    If strColNum <> "" Then
                        If IsNumeric(Mid(strColNum, 3)) Then
                            intToNum = Mid(strColNum, 3)
                        End If
                    End If
                    intLineSum = (intToNum - intFrNum) + 1
                    intSum = intSum + intLineSum
                Next
            End If
        End If
    End With
    
    txtReelQTY.Text = intSum
End Sub

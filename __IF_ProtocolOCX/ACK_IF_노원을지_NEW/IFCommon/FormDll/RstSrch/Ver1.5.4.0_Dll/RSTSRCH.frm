VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmRstSrch 
   Caption         =   "인터페이스 결과 조회 및 등록"
   ClientHeight    =   9930
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15405
   Icon            =   "RSTSRCH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   15405
   StartUpPosition =   2  '화면 가운데
   Begin Threed.SSCommand cmdReg 
      Height          =   900
      Left            =   12375
      TabIndex        =   50
      Top             =   105
      Width           =   2955
      _Version        =   65536
      _ExtentX        =   5212
      _ExtentY        =   1587
      _StockProps     =   78
      Caption         =   "변경사항 저장  F12"
      ForeColor       =   16512
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "RSTSRCH.frx":1272
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   855
      Left            =   12375
      TabIndex        =   46
      Top             =   990
      Width           =   2955
      _Version        =   65536
      _ExtentX        =   5212
      _ExtentY        =   1508
      _StockProps     =   78
      Caption         =   "닫기  ESC"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "RSTSRCH.frx":128E
   End
   Begin Threed.SSFrame SSFrame6 
      Height          =   2115
      Left            =   30
      TabIndex        =   13
      Top             =   30
      Width           =   5115
      _Version        =   65536
      _ExtentX        =   9022
      _ExtentY        =   3731
      _StockProps     =   14
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel5 
         Height          =   1455
         Left            =   150
         TabIndex        =   27
         Top             =   510
         Width           =   4785
         _Version        =   65536
         _ExtentX        =   8440
         _ExtentY        =   2566
         _StockProps     =   15
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
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Begin VB.TextBox txtJGbn 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   31
            Top             =   930
            Width           =   2565
         End
         Begin VB.ComboBox cmbGbn 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "RSTSRCH.frx":12AA
            Left            =   1140
            List            =   "RSTSRCH.frx":12AC
            Style           =   2  '드롭다운 목록
            TabIndex        =   30
            Top             =   570
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtpWDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   1140
            TabIndex        =   0
            Top             =   210
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   52822019
            CurrentDate     =   36361
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   315
            Left            =   150
            TabIndex        =   28
            Top             =   210
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "작업일자"
            ForeColor       =   8454143
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.01
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSCommand cmdList 
            Height          =   1035
            Left            =   2940
            TabIndex        =   32
            Top             =   210
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   1826
            _StockProps     =   78
            Caption         =   "LIST 조회(&L)"
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Picture         =   "RSTSRCH.frx":12AE
         End
         Begin Threed.SSPanel SSPanel16 
            Height          =   285
            Left            =   150
            TabIndex        =   33
            Top             =   570
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "LIST 옵션"
            ForeColor       =   8454143
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.01
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel17 
         Height          =   285
         Left            =   150
         TabIndex        =   29
         Top             =   210
         Width           =   4780
         _Version        =   65536
         _ExtentX        =   8431
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "  해당 작업일 인터페이스 LIST 조회 조건"
         ForeColor       =   12648447
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   1
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   1575
      Left            =   30
      TabIndex        =   36
      Top             =   2040
      Width           =   5115
      _Version        =   65536
      _ExtentX        =   9022
      _ExtentY        =   2778
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel21 
         Height          =   915
         Left            =   150
         TabIndex        =   37
         Top             =   510
         Width           =   2235
         _Version        =   65536
         _ExtentX        =   3942
         _ExtentY        =   1614
         _StockProps     =   15
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
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Begin VB.TextBox txtSSeq 
            Height          =   315
            Left            =   1320
            TabIndex        =   38
            Top             =   150
            Width           =   825
         End
         Begin VB.TextBox txtESeq 
            Height          =   315
            Left            =   1320
            TabIndex        =   39
            Top             =   480
            Width           =   825
         End
         Begin Threed.SSPanel SSPanel22 
            Height          =   315
            Left            =   90
            TabIndex        =   40
            Top             =   150
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Start 라인번호"
            ForeColor       =   0
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.01
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   1
         End
         Begin Threed.SSPanel SSPanel23 
            Height          =   315
            Left            =   90
            TabIndex        =   41
            Top             =   480
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   " End 라인번호"
            ForeColor       =   0
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.01
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   1
         End
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   915
         Left            =   3660
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   510
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   1614
         _StockProps     =   78
         Caption         =   "삭제(&D)"
         ForeColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "RSTSRCH.frx":1700
      End
      Begin Threed.SSCommand cmdServer 
         Height          =   915
         Left            =   2400
         TabIndex        =   43
         Top             =   510
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   1614
         _StockProps     =   78
         Caption         =   "서버 등록(&S)"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "RSTSRCH.frx":1FDA
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   285
         Left            =   150
         TabIndex        =   47
         Top             =   210
         Width           =   4780
         _Version        =   65536
         _ExtentX        =   8431
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "  DATA 서버 등록 및 DATA 삭제"
         ForeColor       =   12648447
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodColor      =   0
         Alignment       =   1
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   6345
      Left            =   30
      TabIndex        =   44
      Top             =   3510
      Width           =   5115
      _Version        =   65536
      _ExtentX        =   9022
      _ExtentY        =   11192
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spdList 
         Height          =   5685
         Left            =   150
         TabIndex        =   45
         Top             =   510
         Width           =   4785
         _Version        =   393216
         _ExtentX        =   8440
         _ExtentY        =   10028
         _StockProps     =   64
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   8421504
         MaxCols         =   123
         MaxRows         =   20
         NoBeep          =   -1  'True
         RetainSelBlock  =   0   'False
         ShadowDark      =   8421504
         SpreadDesigner  =   "RSTSRCH.frx":22F4
         UserResize      =   0
         VirtualMaxRows  =   9999
         VisibleCols     =   123
         VisibleRows     =   20
         TextTip         =   1
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   285
         Left            =   150
         TabIndex        =   48
         Top             =   210
         Width           =   4780
         _Version        =   65536
         _ExtentX        =   8440
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "   조회된 LIST"
         ForeColor       =   12648447
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodColor      =   0
         Alignment       =   1
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1845
      Left            =   5100
      TabIndex        =   14
      Top             =   30
      Width           =   7275
      _Version        =   65536
      _ExtentX        =   12832
      _ExtentY        =   3254
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtPos 
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   5760
         MaxLength       =   7
         TabIndex        =   10
         Top             =   810
         Width           =   1335
      End
      Begin VB.TextBox txtRack 
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   5760
         MaxLength       =   7
         TabIndex        =   9
         Top             =   510
         Width           =   1335
      End
      Begin VB.TextBox txtOther 
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   5760
         MaxLength       =   7
         TabIndex        =   11
         Top             =   1110
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         MaxLength       =   20
         TabIndex        =   25
         Text            =   "12345678901234567890"
         Top             =   -300
         Width           =   2235
      End
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   3870
         MaxLength       =   20
         TabIndex        =   6
         Text            =   "남"
         Top             =   810
         Width           =   915
      End
      Begin VB.TextBox txtReRun 
         Height          =   285
         Left            =   3870
         MaxLength       =   20
         TabIndex        =   8
         Text            =   "N"
         Top             =   1410
         Width           =   915
      End
      Begin VB.TextBox txtLabDate 
         Height          =   285
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   2
         Text            =   "20010501"
         Top             =   810
         Width           =   1665
      End
      Begin VB.TextBox txtNo 
         Height          =   285
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   1
         Text            =   "12345678901234567890"
         Top             =   510
         Width           =   1665
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   4
         Text            =   "김태윤"
         Top             =   1410
         Width           =   1665
      End
      Begin VB.TextBox txtEmer 
         Height          =   285
         Left            =   3870
         MaxLength       =   20
         TabIndex        =   7
         Text            =   "Y"
         Top             =   1110
         Width           =   915
      End
      Begin VB.TextBox txtRegNo 
         Height          =   285
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   3
         Text            =   "00001"
         Top             =   1110
         Width           =   1665
      End
      Begin VB.TextBox txtGbn 
         Height          =   285
         Left            =   3870
         MaxLength       =   20
         TabIndex        =   5
         Text            =   "19990720"
         Top             =   510
         Width           =   915
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   285
         Left            =   180
         TabIndex        =   16
         Top             =   210
         Width           =   6915
         _Version        =   65536
         _ExtentX        =   12197
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "   환자 정보 조회 및 수정, 변경사항 저장(F12)"
         ForeColor       =   12648447
         BackColor       =   32896
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   1
      End
      Begin Threed.SSPanel pnlGbn 
         Height          =   285
         Left            =   2910
         TabIndex        =   17
         Top             =   510
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "구 분"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlRegNo 
         Height          =   285
         Left            =   180
         TabIndex        =   18
         Top             =   1110
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "등록번호"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlSex 
         Height          =   285
         Left            =   2910
         TabIndex        =   19
         Top             =   810
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "성 별"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlRerun 
         Height          =   285
         Left            =   2910
         TabIndex        =   20
         Top             =   1410
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "재 검"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel pnlLabDate 
         Height          =   285
         Left            =   180
         TabIndex        =   21
         Top             =   810
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "접수일자"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel pnlNo 
         Height          =   285
         Left            =   180
         TabIndex        =   22
         Top             =   510
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "검체번호"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel pnlName 
         Height          =   285
         Left            =   180
         TabIndex        =   23
         Top             =   1410
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "이  름"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel pnlEmer 
         Height          =   285
         Left            =   2910
         TabIndex        =   24
         Top             =   1110
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "응 급"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel pnlOther 
         Height          =   285
         Left            =   4890
         TabIndex        =   26
         Top             =   1110
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "기 타"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel pnlRack 
         Height          =   285
         Left            =   4890
         TabIndex        =   34
         Top             =   510
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Tray"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel pnlPos 
         Height          =   285
         Left            =   4890
         TabIndex        =   35
         Top             =   810
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Pos"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   12582912
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   8085
      Left            =   5100
      TabIndex        =   15
      Top             =   1770
      Width           =   10245
      _Version        =   65536
      _ExtentX        =   18071
      _ExtentY        =   14261
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spdRst1 
         Height          =   7080
         Left            =   180
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   510
         Width           =   9885
         _Version        =   393216
         _ExtentX        =   17436
         _ExtentY        =   12488
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         ColHeaderDisplay=   0
         ColsFrozen      =   8
         EditEnterAction =   8
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
         FormulaSync     =   0   'False
         GridColor       =   8421504
         MaxCols         =   19
         MaxRows         =   25
         NoBeep          =   -1  'True
         SpreadDesigner  =   "RSTSRCH.frx":4430
         UserResize      =   0
         VisibleCols     =   4
         VisibleRows     =   20
         TextTip         =   1
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   285
         Left            =   180
         TabIndex        =   49
         Top             =   210
         Width           =   9885
         _Version        =   65536
         _ExtentX        =   17436
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "  검사항목 결과값 조회 및 수정, 변경사항 저장(F12)"
         ForeColor       =   12648447
         BackColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   1
         Begin VB.CheckBox chkBlockMode 
            BackColor       =   &H00000080&
            Caption         =   "결과 블럭설정 기능"
            ForeColor       =   &H00C0FFFF&
            Height          =   210
            Left            =   7950
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   45
            Visible         =   0   'False
            Width           =   1890
         End
      End
   End
End
Attribute VB_Name = "frmRstSrch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gsWSeq       As String
Public giCurRow     As Integer
Public giRegServer  As Integer

Dim miLeaveCell%
Dim msTotIFSeq$

Private Sub Disp_Abnormal_Info(ByVal iRow%)
    Dim vIFSeq, vRst, vPrevRst, vPrevWNo
    Dim sSex$, sDateDiff$, sRefBuf$, sPanBuf$, sDelBuf$, sReturnVal$
    
    With spdRst1
        Call .GetText(2, iRow, vIFSeq)
        Call .GetText(3, iRow, vRst)
        Call .GetText(7, iRow, vPrevRst)
        Call .GetText(8, iRow, vPrevWNo)
        
        If txtSex = "F" Or txtSex = "여" Or txtSex = "2" Then
            sSex = "F"
        Else
            sSex = "M"
        End If
            
        If vIFSeq = "" Then
        Else
            If vPrevWNo = "" Then
                sDateDiff = ""
            Else
                sDateDiff = CStr(DateDiff("d", CDate(Left(Format(vPrevWNo, "@@@@-@@-@@"), 10)), CDate(Format(dtpWDate.Value, "YYYY-MM-DD"))))
                    
                If sDateDiff = "0" Then
                    sDateDiff = "1"
                End If
            End If
            
            sReturnVal = JudgeResultBySex(CStr(vIFSeq), CStr(vRst), sSex, CStr(vPrevRst), sDateDiff, sRefBuf, sPanBuf, sDelBuf)
            
            Call .SetText(3, iRow, CVar(sReturnVal & ""))
            Call .SetText(4, iRow, CVar(sRefBuf & ""))
            Call .SetText(5, iRow, CVar(sPanBuf & ""))
            Call .SetText(6, iRow, CVar(sDelBuf & ""))
        End If
    End With
End Sub

Private Sub Disp_ItemInfo(ByVal iRow%, ByVal sIFSeq$)
    Dim i%
    
    With spdRst1
        If Len(sIFSeq) = 3 Then
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s01 = sIFSeq Then
                    '하한(남)
                    Call .SetText(9, iRow, CVar(gIFItem(i).s10 & ""))
                    '상한(남)
                    Call .SetText(10, iRow, CVar(gIFItem(i).s11 & ""))
                    '하한(여)
                    Call .SetText(11, iRow, CVar(gIFItem(i).s12 & ""))
                    '상한(여)
                    Call .SetText(12, iRow, CVar(gIFItem(i).s13 & ""))
                    'PANIC LOW
                    Call .SetText(13, iRow, CVar(gIFItem(i).s14 & ""))
                    'PANIC HIGH
                    Call .SetText(14, iRow, CVar(gIFItem(i).s15 & ""))
                    'DELTA GBN
                    Select Case gIFItem(i).s16
                        Case "", "0"
                            Call .SetText(15, iRow, CVar("사용안함"))
                        Case "1"
                            Call .SetText(15, iRow, CVar("변화차"))
                        Case "2"
                            Call .SetText(15, iRow, CVar("변화비율"))
                        Case "3"
                            Call .SetText(15, iRow, CVar("기간당 변화차"))
                        Case "4"
                            Call .SetText(15, iRow, CVar("기간당 변화비율"))
                        Case "5"
                            Call .SetText(15, iRow, CVar("절대변화비율"))
                    End Select
                    'DELTA LOW
                    Call .SetText(16, iRow, CVar(gIFItem(i).s17 & ""))
                    'DELTA HIGH
                    Call .SetText(17, iRow, CVar(gIFItem(i).s18 & ""))
                    
                    Exit For
                End If
            Next
        Else
            For i = 1 To giOriginCalItemCnt
                If gCalItem(i).s01 = sIFSeq Then
                    '하한(남)
                    Call .SetText(9, iRow, CVar(gCalItem(i).s08 & ""))
                    '상한(남)
                    Call .SetText(10, iRow, CVar(gIFItem(i).s09 & ""))
                    '하한(여)
                    Call .SetText(11, iRow, CVar(gIFItem(i).s10 & ""))
                    '상한(여)
                    Call .SetText(12, iRow, CVar(gIFItem(i).s11 & ""))
                    'PANIC LOW
                    Call .SetText(13, iRow, CVar(gIFItem(i).s12 & ""))
                    'PANIC HIGH
                    Call .SetText(14, iRow, CVar(gIFItem(i).s13 & ""))
                    'DELTA GBN
                    Select Case gIFItem(i).s14
                        Case "", "0"
                            Call .SetText(15, iRow, CVar("사용안함"))
                        Case "1"
                            Call .SetText(15, iRow, CVar("변화차"))
                        Case "2"
                            Call .SetText(15, iRow, CVar("변화비율"))
                        Case "3"
                            Call .SetText(15, iRow, CVar("기간당 변화차"))
                        Case "4"
                            Call .SetText(15, iRow, CVar("기간당 변화비율"))
                        Case "5"
                            Call .SetText(15, iRow, CVar("절대변화비율"))
                    End Select
                    'DELTA LOW
                    Call .SetText(16, iRow, CVar(gIFItem(i).s15 & ""))
                    'DELTA HIGH
                    Call .SetText(17, iRow, CVar(gIFItem(i).s16 & ""))
                    
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub Disp_PrevResult(ByVal sWDate$, ByVal sWSeq$)
    Dim sBuf$, sOneRow$
    Dim objRST As Object
    Dim i%, iCnt%, iCur%, j%
    Dim vIFSeq, vHeader
    Dim aIFSeq$()
    Dim aPrevRst$()
    Dim aPrevWNo$()
    Dim sPRstBuf$, sPWNoBuf$
    
    If txtRegNo <> "" Then
        Set objRST = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
        
        sBuf = objRST.Get_PreviousResult(gsMachineCd, txtRegNo, sWDate, sWSeq, msTotIFSeq)
        
        Set objRST = Nothing
    End If
    
    iCnt = 0
    
    Erase aIFSeq
    Erase aPrevRst
    Erase aPrevWNo
    
    If sBuf = "" Or sBuf = "NONE" Then
        MsgBox "작업일자-SEQ (" & sWDate & "-" & sWSeq & ") 이전의 해당환자의 결과내역이 없습니다!!", vbInformation
        
        Exit Sub
    End If
    
    Do
        sOneRow = GetByOneUserSymbol(sBuf, sBuf, Chr(3))
        
        If sOneRow = "" Then Exit Do
        
        sOneRow = sOneRow & Chr(124)
        
        iCnt = iCnt + 1
        
        ReDim Preserve aIFSeq(iCnt)
        ReDim Preserve aPrevRst(iCnt)
        ReDim Preserve aPrevWNo(iCnt)
        
        aIFSeq(iCnt) = GetByOne(sOneRow, sOneRow)
        aPrevRst(iCnt) = GetByOne(sOneRow, sOneRow)
        aPrevWNo(iCnt) = GetByOne(sOneRow, sOneRow)
    Loop
    
    With spdRst1
        Call .GetText(7, 0, vHeader)
        
        sBuf = CStr(vHeader)
        
        sBuf = GetByOneUserSymbol(sBuf, sBuf, "번")
        
        sBuf = CStr(Val(sBuf) + 1)
        
        Call .SetText(7, 0, CVar(sBuf & "번전 결과"))
        
        For i = 1 To .MaxRows
            sPRstBuf = "": sPWNoBuf = ""
            
            Call .GetText(2, i, vIFSeq)
            
            If vIFSeq = "" Then Exit For
            
            iCur = 0
            
            If iCnt > 0 Then
                For j = 1 To iCnt
                    If aIFSeq(j) = vIFSeq Then
                        iCur = j
                        
                        Exit For
                    End If
                Next
            End If
            
            If iCur > 0 Then
                sPRstBuf = aPrevRst(iCur)
                sPWNoBuf = aPrevWNo(iCur)
                
                Call .SetText(7, i, CVar(sPRstBuf & ""))
                Call .SetText(8, i, CVar(sPWNoBuf & ""))
            Else
                Call .SetText(7, i, CVar(sPRstBuf & ""))
                Call .SetText(8, i, CVar(sPWNoBuf & ""))
            End If
        Next
    End With
End Sub

Private Sub Disp_PrevResult_Abnormal_Info(ByVal sTotIFSeq$)
    Dim sBuf$, sOneRow$, sDateDiff$
    Dim objRST As Object
    Dim i%, iCnt%, iCur%, j%
    Dim vIFSeq, vRst
    Dim aIFSeq$()
    Dim aPrevRst$()
    Dim aPrevWNo$()
    Dim sRefBuf$, sPanBuf$, sDelBuf$, sPRstBuf$, sPWNoBuf$
    
    Dim sSex$
    
    If txtRegNo <> "" Then
        Set objRST = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
        
        sBuf = objRST.Get_PreviousResult(gsMachineCd, txtRegNo, Format(dtpWDate, "YYYYMMDD"), gsWSeq, sTotIFSeq)
        
        Set objRST = Nothing
    End If
    
    iCnt = 0
    
    Erase aIFSeq
    Erase aPrevRst
    Erase aPrevWNo
    
    Do
        sOneRow = GetByOneUserSymbol(sBuf, sBuf, Chr(3))
        
        If sOneRow = "" Then Exit Do
        
        sOneRow = sOneRow & Chr(124)
        
        iCnt = iCnt + 1
        
        ReDim Preserve aIFSeq(iCnt)
        ReDim Preserve aPrevRst(iCnt)
        ReDim Preserve aPrevWNo(iCnt)
        
        aIFSeq(iCnt) = GetByOne(sOneRow, sOneRow)
        aPrevRst(iCnt) = GetByOne(sOneRow, sOneRow)
        aPrevWNo(iCnt) = GetByOne(sOneRow, sOneRow)
    Loop
    
    With spdRst1
        For i = 1 To .MaxRows
            sRefBuf = "": sPanBuf = "": sDelBuf = "": sPRstBuf = "": sPWNoBuf = ""
            sDateDiff = ""
            
            Call .GetText(2, i, vIFSeq)
            Call .GetText(3, i, vRst)
            
            If txtSex = "F" Or txtSex = "여" Or txtSex = "2" Then
                sSex = "F"
            Else
                sSex = "M"
            End If
            
            If vIFSeq = "" Then Exit For
            
            iCur = 0
            
            If iCnt > 0 Then
                For j = 1 To iCnt
                    If aIFSeq(j) = vIFSeq Then
                        iCur = j
                        
                        Exit For
                    End If
                Next
            End If
            
            If iCur > 0 Then
                sPRstBuf = aPrevRst(iCur)
                sPWNoBuf = aPrevWNo(iCur)
                
                sDateDiff = CStr(DateDiff("d", CDate(Left(Format(sPWNoBuf, "@@@@-@@-@@"), 10)), CDate(Format(dtpWDate.Value, "YYYY-MM-DD"))))
                
                If sDateDiff = "0" Then
                    sDateDiff = "1"
                End If
                
                Call JudgeResultBySex(aIFSeq(iCur), CStr(vRst), sSex, sPRstBuf, sDateDiff, sRefBuf, sPanBuf, sDelBuf)
            Else
                Call JudgeResultBySex(CStr(vIFSeq), CStr(vRst), sSex, sPRstBuf, sDateDiff, sRefBuf, sPanBuf, sDelBuf)
            End If
            
            Call .SetText(4, i, CVar(sRefBuf & ""))
            Call .SetText(5, i, CVar(sPanBuf & ""))
            Call .SetText(6, i, CVar(sDelBuf & ""))
            Call .SetText(7, i, CVar(sPRstBuf & ""))
            Call .SetText(8, i, CVar(sPWNoBuf & ""))
        Next
    End With
End Sub

Private Sub DisplayInit()
    Dim iCnt        As Integer
    Dim i%
    
    Me.Caption = "[" & gsMachineNm & "]  " & Me.Caption
    dtpWDate.Value = Format(Now, "YYYY-MM-DD")
    
    txtLabDate = ""
    txtGbn = ""
    txtNo = ""
    txtRegNo = ""
    txtSex = ""
    txtName = ""
    txtEmer = ""
    txtReRun = ""
    txtOther = ""

    Call GetOrdRstCfg
    
    'Spread는 gRstCfg에 따라서
    For i = 1 To 9
        If gRstcfg.sFUse(i) = "0" Then
            If i < 4 Then
                spdList.ColWidth(i + 1) = 0
            Else
                spdList.ColWidth(i + 3) = 0
            End If
        ElseIf gRstcfg.sFUse(i) = "1" Then
            If i < 4 Then
                Call spdList.SetText(i + 1, 0, CVar(gRstcfg.sFName(i) & ""))
            Else
                Call spdList.SetText(i + 3, 0, CVar(gRstcfg.sFName(i) & ""))
            End If
        End If
    Next
    
    
    If gOrdCfg.sFUse(1) = "0" Then
        pnlLabDate = ""
        txtLabDate.Enabled = False
    ElseIf gOrdCfg.sFUse(1) = "1" Then
        pnlLabDate = gOrdCfg.sFName(1)
    End If
    
    If gOrdCfg.sFUse(2) = "0" Then
        pnlGbn = ""
        txtGbn.Enabled = False
    ElseIf gOrdCfg.sFUse(2) = "1" Then
        pnlGbn = gOrdCfg.sFName(2)
    End If

    If gOrdCfg.sFUse(3) = "0" Then
        pnlNo = ""
        txtNo.Enabled = False
    ElseIf gOrdCfg.sFUse(3) = "1" Then
        pnlNo = gOrdCfg.sFName(3)
    End If

    If gOrdCfg.sFUse(4) = "0" Then
        pnlRegNo = ""
        txtRegNo.Enabled = False
    ElseIf gOrdCfg.sFUse(4) = "1" Then
        pnlRegNo = gOrdCfg.sFName(4)
    End If

    If gOrdCfg.sFUse(5) = "0" Then
        pnlName = ""
        txtName.Enabled = False
    ElseIf gOrdCfg.sFUse(5) = "1" Then
        pnlName = gOrdCfg.sFName(5)
    End If

    If gOrdCfg.sFUse(6) = "0" Then
        pnlSex = ""
        txtSex.Enabled = False
    ElseIf gOrdCfg.sFUse(6) = "1" Then
        pnlSex = gOrdCfg.sFName(6)
    End If

    If gOrdCfg.sFUse(7) = "0" Then
        pnlEmer = ""
        txtEmer.Enabled = False
    ElseIf gOrdCfg.sFUse(7) = "1" Then
        pnlEmer = gOrdCfg.sFName(7)
    End If

    If gOrdCfg.sFUse(8) = "0" Then
        pnlRerun = ""
        txtReRun.Enabled = False
    ElseIf gOrdCfg.sFUse(8) = "1" Then
        pnlRerun = gOrdCfg.sFName(8)
    End If

    If gOrdCfg.sFUse(9) = "0" Then
        pnlOther = ""
        txtOther.Enabled = False
    ElseIf gOrdCfg.sFUse(9) = "1" Then
        pnlOther = gOrdCfg.sFName(9)
    End If
    
    'Rack, Pos 설정
    If Val(gIFRack.sMaxRack) = 0 Then
        spdList.ColWidth(5) = 0
        spdList.ColWidth(6) = 0
        
        pnlRack = ""
        pnlPos = ""
        txtRack.Enabled = False
        txtPos.Enabled = False
    Else
        If gsIFMode = "1" Then
        'Rack Or Tray 방식 지원안함, But Rack/Pos 표시
            pnlRack = "Rack"
            pnlPos = "Pos"
            
            Call spdList.SetText(5, 0, CVar("Rack"))
            Call spdList.SetText(6, 0, CVar("Pos"))
        ElseIf gsIFMode = "2" Then
        'Rack Or Tray 방식 지원안함, But Tray/Pos 표시
            pnlRack = "Tray"
            pnlPos = "Pos"
            
            Call spdList.SetText(5, 0, CVar("Tray"))
            Call spdList.SetText(6, 0, CVar("Pos"))
        ElseIf gsIFMode = "3" Then
        'Rack Or Tray 방식 지원안함, But Tray/Cup 표시
            pnlRack = "Tray"
            pnlPos = "Cup"
            
            Call spdList.SetText(5, 0, CVar("Tray"))
            Call spdList.SetText(6, 0, CVar("Cup"))
        ElseIf gsIFMode = "4" Then
        'Rack/Pos 방식 지원
            pnlRack = "Rack"
            pnlPos = "Pos"
            
            Call spdList.SetText(5, 0, CVar("Rack"))
            Call spdList.SetText(6, 0, CVar("Pos"))
        ElseIf gsIFMode = "5" Then
        'Tray/Pos 방식 지원
            pnlRack = "Tray"
            pnlPos = "Pos"
            
            Call spdList.SetText(5, 0, CVar("Tray"))
            Call spdList.SetText(6, 0, CVar("Pos"))
        ElseIf gsIFMode = "6" Then
        'Tray/Cup 방식 지원
            pnlRack = "Tray"
            pnlPos = "Cup"
            
            Call spdList.SetText(5, 0, CVar("Tray"))
            Call spdList.SetText(6, 0, CVar("Cup"))
        End If
    End If
               
    With spdList
        .MaxRows = 0
        .Row = -1
        .Row2 = -1
        .Col = -1
        .Col2 = -1
        .BlockMode = True
        .BackColor = 연노랑
        .Lock = True
        .BlockMode = False
    End With
    
    With spdRst1
        .EditModePermanent = True
    End With
    
    With cmbGbn
        Call .AddItem("ALL", 0)
        Call .AddItem("등록", 1)
        Call .AddItem("미등록", 2)
        Call .AddItem(pnlNo & "(A)", 3)
        Call .AddItem(pnlNo & "(L)", 4)
        
        .ListIndex = 2
    End With
    
End Sub

Private Sub chkBlockMode_Click()
    
    With spdRst1
        If chkBlockMode.Value = vbChecked Then
            .EditModePermanent = False
        Else
            .EditModePermanent = True
        End If
    End With
    
End Sub

Private Sub cmbGbn_Click()
    txtJGbn = ""
End Sub

Private Sub cmbGbn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        txtJGbn.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim i%, iDelFail%
    Dim objRST      As Object
    Dim vWSeq
    
    If txtSSeq = "" Then
        MsgBox "Start SeqNo를 입력해 주십시요!!"
        txtSSeq.SetFocus
        Exit Sub
    End If
    
    If txtESeq = "" Then
        MsgBox "End SeqNo를 입력해 주십시요!!"
        txtESeq.SetFocus
        Exit Sub
    End If
    
    iDelFail = 0
    
    With spdList
        If .MaxRows = 0 Then Exit Sub
        
        If MsgBox("라인 " & txtSSeq & " 번에서 라인 " & txtESeq & " 번까지 로컬 데이터에서 삭제하시겠습니까?", vbYesNo + vbQuestion, "로컬 삭제 여부") = vbYes Then
            Set objRST = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
                    
            For i = CInt(Val(txtSSeq)) To CInt(Val(txtESeq))
                Call .GetText(1, i, vWSeq)
                    
                If objRST.Del_IFResult(gsMachineCd, 0, Format(dtpWDate, "YYYYMMDD"), CStr(vWSeq)) = False Then
                    MsgBox "삭제에 실패하였습니다."
                    iDelFail = 1
                End If
            Next
        
            Set objRST = Nothing
            
            .Row = CInt(Val(txtSSeq))
            .Row2 = CInt(Val(txtESeq))
            .BlockMode = True
            .Action = ActionDeleteRow
            .BlockMode = False
            
            .MaxRows = .MaxRows - (.Row2 - .Row + 1)
        End If
    End With
    
    spdRst1.MaxRows = 0
    spdRst1.MaxRows = 110
    
    If iDelFail = 1 Then
        Call cmdList_Click
    Else
        txtSSeq = "": txtESeq = ""
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdList_Click()
    Dim objRST      As Object
    Dim sRstList    As String
    Dim sRstOneRow  As String
    Dim sRst        As String
    
    Dim sWSeq       As String
    Dim sIFSeq      As String
    Dim sRst1       As String
    Dim sRst2       As String
    Dim sFlag       As String
    Dim sTmp        As String
    Dim iIFCnt      As Integer
    
    Dim iFldCnt     As Integer
    
    Dim sPrevWSeq   As String
    Dim iNew        As Integer
    
    Dim vReturnVal  As Variant
    Dim lngRowCnt   As Long
    
    Dim sSpcNo      As String
    
    '초기화
    spdList.MaxRows = 0
    
    txtLabDate = ""
    txtGbn = ""
    txtNo = ""
    txtRegNo = ""
    txtSex = ""
    txtName = ""
    txtEmer = ""
    txtReRun = ""
    txtOther = ""
    
    txtSSeq = "": txtESeq = ""
    
    With spdRst1
        .Col = 1
        .Col2 = .MaxCols
        .Row = 1
        .Row2 = .MaxRows
        .BlockMode = True
        .Action = SS_ACTION_CLEAR_TEXT
        .BlockMode = False
    End With
    
    Screen.MousePointer = vbHourglass
    
    Set objRST = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
'    Call cmbGbn.AddItem("ALL", 0)
'    Call cmbGbn.AddItem("등록", 1)
'    Call cmbGbn.AddItem("미등록", 2)
'    Call cmbGbn.AddItem("검체번호(A)", 3)
'    Call cmbGbn.AddItem("검체번호(L)", 4)
    
    'Get_IFResult의 iGbn
    '0 - All
    '1 - Register
    '2 - No Register
    '3 - JNo(All)
    '4 - JNo(Last)
    Select Case cmbGbn.ListIndex
        Case 3, 4
            sRstList = objRST.Get_IFResult(gsMachineCd, Format(dtpWDate, "YYYYMMDD"), "", cmbGbn.ListIndex, txtJGbn, vReturnVal)
        Case Else
            sRstList = objRST.Get_IFResult(gsMachineCd, Format(dtpWDate, "YYYYMMDD"), "", cmbGbn.ListIndex, , vReturnVal)
    End Select
    
    Set objRST = Nothing
    
    If sRstList = "NONE" Then
        Screen.MousePointer = vbDefault
        
        If giRegServer = 0 Then
            MsgBox "조회된 데이터가 없습니다!!", vbInformation, Me.Caption
        Else
        End If
        
        Exit Sub
    End If

    iIFCnt = 1
    
    spdList.ReDraw = False
    
    With spdList
        For lngRowCnt = 0 To UBound(vReturnVal, 2)
            'WSEQ 저장
            sWSeq = vReturnVal(0, lngRowCnt)
            
            If .MaxRows = 0 Then
                iNew = 1
            Else
                If sWSeq = sPrevWSeq Then
                    iNew = 0
                Else
                    iNew = 1
                End If
            End If
            
            If iNew = 1 Then
                .MaxRows = .MaxRows + 1
                    
                Call .SetText(1, .MaxRows, CVar(sWSeq))
    
                For iFldCnt = 2 To 12
                    sRst = vReturnVal(iFldCnt - 1, lngRowCnt)
                    Call .SetText(iFldCnt, .MaxRows, CVar(sRst))
                Next
                
                iIFCnt = 1
            Else
                iIFCnt = iIFCnt + 1
            End If
            
            sSpcNo = vReturnVal(4 - 1, lngRowCnt) '-- 검체번호
            
            sIFSeq = vReturnVal(12, lngRowCnt)
            sRst1 = vReturnVal(13, lngRowCnt)
            sRst2 = vReturnVal(14, lngRowCnt)
            sFlag = vReturnVal(15, lngRowCnt)
            
            Call .SetText(13, .MaxRows, CVar(iIFCnt))
            Call .SetText(13 + iIFCnt, .MaxRows, CVar(sIFSeq & Chr(124) & sRst1 & Chr(124) & sRst2 & Chr(124) & sFlag & Chr(124) & sSpcNo & Chr(124)))
            
            sPrevWSeq = sWSeq
        Next
    End With
        
    spdList.ReDraw = True
    
    If spdList.MaxRows > 0 Then
        Call spdList_Click(1, 1)
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdReg_Click()
    Dim objRST      As Object
    Dim sIFSeq      As String
    Dim sResult1    As String
    Dim sResult2    As String
    Dim sFlag       As String
    Dim sSpcNo      As String
    
    Dim iCnt        As Integer
    
    Dim vIFSeq
    Dim vResult1
    Dim vResult2
    Dim vFlag
    Dim vSpcNo
        
    Dim iItemCnt    As Integer
    
    Dim iMerge%
    
    If spdList.MaxRows = 0 Or gsWSeq = "" Then
        Exit Sub
    End If
    
    iItemCnt = 0
    
    For iCnt = 1 To spdRst1.MaxRows
        Call spdRst1.GetText(2, iCnt, vIFSeq)
        Call spdRst1.GetText(3, iCnt, vResult1)
        Call spdRst1.GetText(4, iCnt, vResult2)
        Call spdRst1.GetText(18, iCnt, vSpcNo)
        Call spdRst1.GetText(19, iCnt, vFlag)
        
        If vIFSeq = "" Then Exit For
        
        If vIFSeq <> "" Then
            sIFSeq = sIFSeq & vIFSeq & Chr(124)
            sResult1 = sResult1 & vResult1 & Chr(124)
            sResult2 = sResult2 & vResult2 & Chr(124)
            sFlag = sFlag & Trim(vFlag) & Chr(124)
            sSpcNo = sSpcNo & vSpcNo & Chr(124)
            
            iItemCnt = iItemCnt + 1
        End If
    Next
    
    Screen.MousePointer = vbHourglass
    
    Set objRST = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))

    If objRST.Edit_IFResult(gsMachineCd, 1, Format(dtpWDate, "YYYYMMDD"), gsWSeq, sIFSeq, txtLabDate, txtGbn, txtNo, _
                            txtRegNo, txtName, txtSex, txtEmer, txtReRun, txtOther, sResult1, sResult2, sFlag, "0", iItemCnt) = False Then
        
        Screen.MousePointer = vbDefault
        
        MsgBox "로컬 저장에 실패하였습니다."
        
        Set objRST = Nothing
        
        Exit Sub
    End If

    Set objRST = Nothing

    Screen.MousePointer = vbDefault
    
    '스프레드 내용 재설정
    If iMerge = 0 Then
        With spdList
            Call .SetText(2, giCurRow, CVar(txtLabDate))
            Call .SetText(3, giCurRow, CVar(txtGbn))
            Call .SetText(4, giCurRow, CVar(txtNo))
            Call .SetText(7, giCurRow, CVar(txtRegNo))
            Call .SetText(8, giCurRow, CVar(txtName))
            Call .SetText(9, giCurRow, CVar(txtSex))
            Call .SetText(10, giCurRow, CVar(txtEmer))
            Call .SetText(11, giCurRow, CVar(txtReRun))
            Call .SetText(12, giCurRow, CVar(txtOther))
    
            'MERGE 관련 sIFSeq & Chr(124) & sRst1 & Chr(124) & sRst2 & Chr(124) & sSpcNo & Chr(124)
            For iCnt = 1 To iItemCnt
                Call .SetText(13 + iCnt, giCurRow, CVar(GetByOne(sIFSeq, sIFSeq) & Chr(124) _
                            & GetByOne(sResult1, sResult1) & Chr(124) & GetByOne(sResult2, sResult2) & Chr(124) & GetByOne(sFlag, sFlag) & Chr(124) & txtNo & Chr(124)))
                Call spdRst1.SetText(18, iCnt, CVar(txtNo & ""))
            Next
        End With
    ElseIf iMerge = 1 Then
        For iCnt = 1 To iItemCnt
            Call spdList.SetText(13 + iCnt, giCurRow, CVar(GetByOne(sIFSeq, sIFSeq) _
                                & Chr(124) & GetByOne(sResult1, sResult1) & Chr(124) & GetByOne(sResult2, sResult2) & Chr(124) & GetByOne(sFlag, sFlag) & Chr(124) & GetByOne(sSpcNo, sSpcNo) & Chr(124)))
        Next
    End If
End Sub

Private Sub cmdServer_Click()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    Dim i%, j%, iPersonCnt%
    Dim sWDate$, sTWSeq$, sTJDate$, sTJGbn$, sTJNo$, sTRack$, sTPos$, sTRegNo$, sTName$, sTSex$, sTEmer$, sTReRun$, sTOther$
    Dim sItemCnt$, sTIFSeq$, sTServerCd$, sTRst1$, sTRst2$, sTFlag$, sCd$, sRst1$, sRst2$, sFlag$
    Dim vTmp
    Dim objRegRst As Object
    
    Dim sTmp         As String
    Dim sSpcNo       As String
    Dim sOtherSpcNo  As String
    
    Dim sTIFSeq_1    As String
    Dim sTServerCd_1 As String
    Dim sTRst1_1     As String
    Dim sTRst2_1     As String
    Dim sTFlag_1     As String
    Dim iTCnt_1      As Integer
    
    Dim sTIFSeq_2    As String
    Dim sTServerCd_2 As String
    Dim sTRst1_2     As String
    Dim sTRst2_2     As String
    Dim sTFlag_2     As String
    Dim iTCnt_2      As Integer
    
    Dim sTItemCnt$
    
    If txtSSeq = "" Then
        MsgBox "Start 라인번호를 입력해 주십시요!!", vbInformation
        txtSSeq.SetFocus
        Exit Sub
    End If
    
    If txtESeq = "" Then
        MsgBox "End 라인번호를 입력해 주십시요!!", vbInformation
        txtESeq.SetFocus
        Exit Sub
    End If
    
    '초기화
    sWDate = "": sTWSeq = "": sTJDate = "": sTJGbn = "": sTJNo = "": sTRack = "": sTPos = ""
    sTRegNo = "": sTName = "": sTSex = "": sTEmer = "": sTReRun = "": sTOther = ""
    sTIFSeq = "": sTServerCd = "": sTRst1 = "": sTRst2 = ""
    iPersonCnt = 0
    
    sWDate = Format(dtpWDate.Value, "YYYYMMDD")
        
    With spdList
        If .MaxRows = 0 Then Exit Sub
        
        For i = Val(txtSSeq) To Val(txtESeq)
            iPersonCnt = iPersonCnt + 1
            
            Call .GetText(1, i, vTmp)
            sTWSeq = sTWSeq & CStr(vTmp) & Chr(124)
            
            Call .GetText(2, i, vTmp)
            sTJDate = sTJDate & CStr(vTmp) & Chr(124)
            
            Call .GetText(3, i, vTmp)
            sTJGbn = sTJGbn & CStr(vTmp) & Chr(124)
            
            Call .GetText(4, i, vTmp)
            sSpcNo = CStr(vTmp)
            sTJNo = sTJNo & CStr(vTmp) & Chr(124)
            
            Call .GetText(5, i, vTmp)
            sTRack = sTRack & CStr(vTmp) & Chr(124)
            
            Call .GetText(6, i, vTmp)
            sTPos = sTPos & CStr(vTmp) & Chr(124)
            
            Call .GetText(7, i, vTmp)
            sTRegNo = sTRegNo & CStr(vTmp) & Chr(124)
            
            Call .GetText(8, i, vTmp)
            sTName = sTName & CStr(vTmp) & Chr(124)
            
            Call .GetText(9, i, vTmp)
            sTSex = sTSex & CStr(vTmp) & Chr(124)
            
            Call .GetText(10, i, vTmp)
            sTEmer = sTEmer & CStr(vTmp) & Chr(124)
            
            Call .GetText(11, i, vTmp)
            sTReRun = sTReRun & CStr(vTmp) & Chr(124)
            
            'Call .GetText(12, i, vTmp)
            sTOther = sTOther & txtOther & Chr(124)
            
            '-- 검사항목 갯수
            Call .GetText(13, i, vTmp)
            
            sTIFSeq_1 = "":    sTIFSeq_2 = ""
            sTServerCd_1 = "": sTServerCd_2 = ""
            sTRst1_1 = "":     sTRst1_2 = ""
            sTRst2_1 = "":     sTRst2_2 = ""
            sTFlag_1 = "":     sTFlag_2 = ""
            iTCnt_1 = 0:      iTCnt_2 = 0
            
            For j = 1 To Val(vTmp)
                Call .GetText(13 + j, i, vTmp)
                sBuf = CStr(vTmp)
                
                sCd = GetByOne(sBuf, sBuf)
                sRst1 = GetByOne(sBuf, sBuf)
                sRst2 = GetByOne(sBuf, sBuf)
                sFlag = GetByOne(sBuf, sBuf)
                
                sTmp = GetByOne(sBuf, sBuf)
                
                If sSpcNo <> sTmp Then
                    sOtherSpcNo = sTmp
                     
                    sTIFSeq_2 = sTIFSeq_2 & sCd & Chr(124)
                    
                    '서버쪽 코드로 바꿈
                    sCd = ConvertIFItemInfo(2, sCd)
                    
                    sTServerCd_2 = sTServerCd_2 & sCd & Chr(124)
                    sTRst1_2 = sTRst1_2 & sRst1 & Chr(124)
                    sTRst2_2 = sTRst2_2 & sRst2 & Chr(124)
                    sTFlag_2 = sTFlag_2 & sFlag & Chr(124)
                    iTCnt_2 = iTCnt_2 + 1
                    
                Else
                    sTIFSeq_1 = sTIFSeq_1 & sCd & Chr(124)
                    
                    '서버쪽 코드로 바꿈
                    sCd = ConvertIFItemInfo(2, sCd)
                    
                    sTServerCd_1 = sTServerCd_1 & sCd & Chr(124)
                    sTRst1_1 = sTRst1_1 & sRst1 & Chr(124)
                    sTRst2_1 = sTRst2_1 & sRst2 & Chr(124)
                    sTFlag_1 = sTFlag_1 & sFlag & Chr(124)
                    iTCnt_1 = iTCnt_1 + 1
                
                End If
            Next
            
            If iTCnt_2 <> 0 Then
                iPersonCnt = iPersonCnt + 1
                
                Call .GetText(1, i, vTmp)
                sTWSeq = sTWSeq & CStr(vTmp) & Chr(124)
                
                Call .GetText(2, i, vTmp)
                sTJDate = sTJDate & CStr(vTmp) & Chr(124)
                
                Call .GetText(3, i, vTmp)
                sTJGbn = sTJGbn & CStr(vTmp) & Chr(124)
                
'                Call .GetText(4, i, vTmp)
                sTJNo = sTJNo & sOtherSpcNo & Chr(124)
                
                Call .GetText(5, i, vTmp)
                sTRack = sTRack & CStr(vTmp) & Chr(124)
                
                Call .GetText(6, i, vTmp)
                sTPos = sTPos & CStr(vTmp) & Chr(124)
                
                Call .GetText(7, i, vTmp)
                sTRegNo = sTRegNo & CStr(vTmp) & Chr(124)
                
                Call .GetText(8, i, vTmp)
                sTName = sTName & CStr(vTmp) & Chr(124)
                
                Call .GetText(9, i, vTmp)
                sTSex = sTSex & CStr(vTmp) & Chr(124)
                
                Call .GetText(10, i, vTmp)
                sTEmer = sTEmer & CStr(vTmp) & Chr(124)
                
                Call .GetText(11, i, vTmp)
                sTReRun = sTReRun & CStr(vTmp) & Chr(124)
                
'                'Call .GetText(12, i, vTmp)
'                sTOther = sTOther & txtOther & Chr(124)
                
                
                '-- 검사항목 갯수
                sTItemCnt = sTItemCnt & CStr(iTCnt_1) & Chr(124) & CStr(iTCnt_2) & Chr(124)
                
                sTIFSeq = sTIFSeq & sTIFSeq_1 & Chr(3) & sTIFSeq_2 & Chr(3)
                sTServerCd = sTServerCd & sTServerCd_1 & Chr(3) & sTServerCd_2 & Chr(3)
                sTRst1 = sTRst1 & sTRst1_1 & Chr(3) & sTRst1_2 & Chr(3)
                sTRst2 = sTRst2 & sTRst2_1 & Chr(3) & sTRst2_2 & Chr(3)
                sTFlag = sTFlag & sTFlag_1 & Chr(3) & sTFlag_2 & Chr(3)
                
            Else
                '-- 검사항목 갯수
                sTItemCnt = sTItemCnt & CStr(iTCnt_1) & Chr(124)
                
                sTIFSeq = sTIFSeq & sTIFSeq_1 & Chr(3)
                sTServerCd = sTServerCd & sTServerCd_1 & Chr(3)
                sTRst1 = sTRst1 & sTRst1_1 & Chr(3)
                sTRst2 = sTRst2 & sTRst2_1 & Chr(3)
                sTFlag = sTFlag & sTFlag_1 & Chr(3)
                
            End If
        Next
    End With
    
    'Register Result Current Object
    sBuf = gRstcfg.sComponent
    
    Me.MousePointer = vbHourglass
    
    Set objRegRst = CreateObject(sBuf)
    
    Call objRegRst.SetMachineInfo(gsMachineCd, gsMachineNm)
    Call objRegRst.RegServer(iPersonCnt, sWDate, sTWSeq, _
                             sTJDate, sTJGbn, sTJNo, _
                             sTRack, sTPos, _
                             sTRegNo, sTName, sTSex, _
                             sTEmer, sTReRun, sTOther, _
                             sTItemCnt, sTIFSeq, sTServerCd, _
                             sTRst1, sTRst2, sTFlag)
                
    Set objRegRst = Nothing
    
    Me.MousePointer = vbDefault
    
    giRegServer = 1
    
    If cmbGbn.ListIndex = 2 Then
        Call cmdList_Click
    End If
    
    giRegServer = 0
    
    Exit Sub
ErrHandler:
    Set objRegRst = Nothing
    Me.MousePointer = vbDefault
    
    MsgBox "cmdServer_Click 오류 (" & Err.Description & ")"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i%
    
    Select Case KeyCode
        Case vbKeyF4
            If MsgBox("조회된 LIST에서 삭제하시겠습니까?", vbYesNo, "조회된 LIST 삭제 여부") = vbYes Then
                With spdList
                    .Row = giCurRow
                    .Action = ActionDeleteRow
                    .MaxRows = .MaxRows - 1
                    
                    Call spdList_Click(1, giCurRow)
                End With
            End If
        
        Case vbKeyF12
            Me.MousePointer = vbHourglass
            
            Call cmdReg_Click
            
            DoEvents
            spdRst1.Action = ActionSelModeClear
                
            For i = 1 To spdList.MaxRows
                spdList.Col = 1
                spdList.Row = i
                    
                If spdList.BackColor = 연빨강 Then
                    Call spdList_Click(1, i)
                    spdList.Col = 1
                    spdList.Row = i
                    spdList.Action = ActionActiveCell
                    spdList.SetFocus
                    Exit For
                End If
            Next
            
            Me.MousePointer = vbDefault
            
        Case vbKeyEscape
            Call cmdExit_Click
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    giRegServer = 0
    miLeaveCell = 0
    
    Call GetMachineInfo
    Call DisplayInit
    Call GetTestItem
    
    Exit Sub
ErrHandler:
    MsgBox "Form_Load 오류(FRS01.Dll) - (" & Err.Description & ")"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call RegEditCurFrmTitle("RstSrch", "")
    ViewMsg ""
End Sub

Private Sub spdList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    If BlockCol = -1 Then
        If BlockRow = -1 Then
            txtSSeq = 1
            txtESeq = spdList.MaxRows
        Else
            txtSSeq = CStr(BlockRow)
            txtESeq = CStr(BlockRow2)
        End If
        
        spdList.Action = ActionDeselectBlock
    End If
End Sub

Private Sub spdList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vTmp
    Dim sBuf        As String
    
    Dim sIFSeq      As String
    Dim sRst        As String
    Dim sRst2       As String
    Dim sFlag       As String
    Dim sSpcNo      As String
    Dim sIFNm       As String
    
    Dim iCnt        As Integer
    Dim iCnt2       As Integer
    Dim iIFCnt      As Integer
    Dim iRealCnt    As Integer
    Dim iExist      As Integer
    
    Dim iPos%
    Dim arrTmp
    
    If Row = 0 Then
        Exit Sub
    End If
    
    If miLeaveCell = 1 Then
        miLeaveCell = 0
        
        Exit Sub
    End If
    
    miLeaveCell = miLeaveCell - 1
    
    spdRst1.MaxRows = 0
    spdRst1.MaxRows = 110
    
    Call spdReverse(spdList, -1, -1, Row, Row, 연빨강, 3)
    giCurRow = Row

    '작업번호
    Call spdList.GetText(1, Row, vTmp)
    gsWSeq = vTmp

    '접수일자
    Call spdList.GetText(2, Row, vTmp)
    txtLabDate = vTmp
    
    '구분
    Call spdList.GetText(3, Row, vTmp)
    txtGbn = vTmp
    
    '번호
    Call spdList.GetText(4, Row, vTmp)
    txtNo = vTmp
    
    'Rack
    Call spdList.GetText(5, Row, vTmp)
    txtRack = vTmp
    
    'Pos
    Call spdList.GetText(6, Row, vTmp)
    txtPos = vTmp
    
    '등록번호
    Call spdList.GetText(7, Row, vTmp)
    txtRegNo = vTmp
    
    '이름
    Call spdList.GetText(8, Row, vTmp)
    txtName = vTmp
    
    '성별
    Call spdList.GetText(9, Row, vTmp)
    txtSex = vTmp
    
'    With spdRst1
'        If txtSex = "F" Or txtSex = "여" Or txtSex = "2" Then
'            .ColWidth(9) = 0
'            .ColWidth(10) = 0
'            .ColWidth(11) = 6.5
'            .ColWidth(12) = 6.5
'        Else
'            .ColWidth(9) = 6.5
'            .ColWidth(10) = 6.5
'            .ColWidth(11) = 0
'            .ColWidth(12) = 0
'        End If
'    End With
    
    '응급
    Call spdList.GetText(10, Row, vTmp)
    txtEmer = vTmp
    
    '재검
    Call spdList.GetText(11, Row, vTmp)
    txtReRun = vTmp
    
    '기타
    Call spdList.GetText(12, Row, vTmp)
    txtOther = vTmp

    'IFCOUNT
    Call spdList.GetText(13, Row, vTmp)
    iIFCnt = Val(vTmp)
    
    Me.MousePointer = vbHourglass
    
    msTotIFSeq = ""
    
    Call spdRst1.SetText(7, 0, CVar("1번전 결과"))
    
    For iCnt = 1 To iIFCnt
        iExist = 0
        
        Call spdList.GetText(13 + iCnt, Row, vTmp)
        sBuf = vTmp
        
        sIFSeq = GetByOne(sBuf, sBuf)
        sRst = GetByOne(sBuf, sBuf)
        sRst2 = GetByOne(sBuf, sBuf)
        sFlag = GetByOne(sBuf, sBuf)
        sSpcNo = GetByOne(sBuf, sBuf)
        
        msTotIFSeq = msTotIFSeq & sIFSeq & "','"
        
        If LenH(sIFSeq) = 3 Then
            For iCnt2 = 1 To giOriginIFItemCnt
                If gIFItem(iCnt2).s01 = sIFSeq Then
                    sIFNm = gIFItem(iCnt2).s02
                    iCnt2 = giOriginIFItemCnt + 1
                    iRealCnt = iRealCnt + 1
                    
                    iExist = 1
                    
                    Exit For
                End If
            Next
        ElseIf LenH(sIFSeq) = 2 Then
            For iCnt2 = 1 To giOriginCalItemCnt
                If gCalItem(iCnt2).s01 = sIFSeq Then
                    sIFNm = gCalItem(iCnt2).s02
                    iCnt2 = giOriginCalItemCnt + 1
                    iRealCnt = iRealCnt + 1
                    
                    iExist = 1
                    
                    Exit For
                End If
            Next
        End If

        If iExist = 1 And iRealCnt < spdRst1.MaxRows + 1 Then
            Call spdRst1.SetText(1, iRealCnt, sIFNm)
            Call spdRst1.SetText(2, iRealCnt, sIFSeq)
            Call spdRst1.SetText(3, iRealCnt, sRst)
            Call spdRst1.SetText(18, iRealCnt, sSpcNo)
            Call spdRst1.SetText(19, iRealCnt, sFlag)
            
'            '-- 2002-05-26 JJH 추가
'            '   특정항목인경우 ComboBox입력으로....
'            iPos = InStr(1, gsComboBox_InputItems, sIFSeq)
'            If iPos > 0 Then
'
'                '-- 해당항목 위치설정
'                arrTmp = Split(gsComboBox_InputItems, "|")
'                For iPos = 0 To UBound(arrTmp)
'                    If arrTmp(iPos) = sIFSeq Then
'                        Exit For
'                    End If
'                Next
'
'                '-- 해당항목의 결과목록
'                arrTmp = Split(gsComboBox_InputResults, "|")
'                arrTmp(iPos) = Replace(arrTmp(iPos), Chr$(1), Chr$(9))
'
'                With spdRst1
'                    .Row = iRealCnt
'                    .Col = 3
'                    .CellType = CellTypeComboBox
'                    .TypeComboBoxList = arrTmp(iPos)
'                    '.TypeComboBoxList = "0-1" & Chr$(9) & "1-4" & Chr$(9) + "5-9" & Chr$(9) + "10-29" & Chr$(9) + "many"
'                End With
'            End If
'            '-----------------

            If sIFSeq <> "" Then
                Call Disp_ItemInfo(iRealCnt, sIFSeq)
            End If
        End If
    Next
    
    If iIFCnt > 0 Then
        msTotIFSeq = "'" & Left(msTotIFSeq, Len(msTotIFSeq) - 2)
        
        Call Disp_PrevResult_Abnormal_Info(msTotIFSeq)
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub spdList_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    If MsgBox("조회된 LIST에서 화면 삭제하시겠습니까?", vbYesNo, "조회된 LIST 화면 삭제 여부") = vbYes Then
        With spdList
            .Row = Row
            .Action = ActionDeleteRow
            .MaxRows = .MaxRows - 1
            
            Call spdList_Click(1, Row)
        End With
    End If
End Sub

Private Sub spdList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        With spdRst1
            .SetFocus
            .Row = 1
            .Col = 3
            .Action = ActionActiveCell
        End With
    End If
End Sub

Private Sub spdList_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If Row = NewRow Then Exit Sub
    If Row < 0 Then Exit Sub
    If NewRow < 0 Then Exit Sub
    
    If NewRow = -1 Then
    Else
        miLeaveCell = 2
        
        Call spdList_Click(1, NewRow)
    End If
End Sub

Private Sub spdRst1_Change(ByVal Col As Long, ByVal Row As Long)
    With spdRst1
        If Col = 3 Then
            Call Disp_Abnormal_Info(CInt(Row))
        End If
    End With
End Sub

Private Sub spdRst1_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim vWNo
    
    If Row < 1 Then Exit Sub
    
    If Col = 8 Then
        With spdRst1
            Call .GetText(8, Row, vWNo)
        End With
        
        If vWNo <> "" Then
            Me.MousePointer = vbHourglass
            Call Disp_PrevResult(Left(CStr(vWNo), 8), Right(CStr(vWNo), 4))
            Me.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub spdRst1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i%
    
    If KeyCode = vbKeyF12 Then
        Me.MousePointer = 11
        
        Call cmdReg_Click
        
        DoEvents
        spdRst1.Action = ActionSelModeClear
            
        For i = 1 To spdList.MaxRows
            spdList.Col = 1
            spdList.Row = i
                
            If spdList.BackColor = 연빨강 Then
                Call spdList_Click(1, i)
                spdList.Col = 1
                spdList.Row = i
                spdList.Action = ActionActiveCell
                spdList.SetFocus
                Exit For
            End If
        Next
        
        Me.MousePointer = 0
    End If
End Sub

Private Sub spdRst1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim vOrigin, vTmp, vIFSeq
    Dim sRst1$, sRst2$, sIFSeq$, sBuf$, sCal$
    
    With spdRst1
        If Col = 2 Then
            Call .GetText(2, Row, vTmp)
            Call .GetText(4, Row, vIFSeq)
            
            If vTmp = "" Or vIFSeq = "" Then
            Else
                sRst1 = CStr(vTmp)
            End If
        End If
    End With
End Sub

Private Sub txtESeq_GotFocus()
    Call Txt_Highlight(txtESeq)
End Sub

Private Sub txtESeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtESeq_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        cmdServer.SetFocus
'    End If
    
    Call TxtTypeOnlyNumeric(txtESeq, KeyAscii)
End Sub

Private Sub txtJGbn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtJGbn <> "" Then cmdList_Click
        Call Txt_Highlight(txtJGbn)
    End If
End Sub

Private Sub txtNo_GotFocus()
    Call Txt_Highlight(txtNo)
End Sub

Private Sub txtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHandler
    
    Dim sBuf$
    Dim objOrd As Object
    
    If KeyCode = vbKeyReturn Then
        If Len(txtNo) < Val(gOrdCfg.sFSize(3)) Then
            txtNo = Format(txtNo, RackFormat(CStr(Val(gOrdCfg.sFSize(3)))))
        End If
        
        If Len(txtNo) = Val(gOrdCfg.sFSize(3)) Then
            sBuf = gOrdCfg.sComponent
            
            If sBuf = "" Then
                ViewMsg "오더 Dll 파일이 존재하지 않습니다!!"
                Exit Sub
            End If
            
            Screen.MousePointer = vbHourglass
            
            Set objOrd = CreateObject(sBuf)
            Call objOrd.SetMachineInfo(gsMachineCd, gsMachineNm)
            sBuf = objOrd.FetchPatInfo(gsMachineCd, txtNo)
            Set objOrd = Nothing
            
            Screen.MousePointer = vbDefault
            
            If sBuf = "" Then
                ViewMsg "입력한 바코드 번호에 대한 정보가 존재하지 않습니다!!"
                Exit Sub
            Else
                txtLabDate = GetByOne(sBuf, sBuf)
                txtJGbn = ""        'GetByOne(sBuf, sBuf)
                txtNo = GetByOne(sBuf, sBuf)
                txtRegNo = GetByOne(sBuf, sBuf)
                txtName = GetByOne(sBuf, sBuf)
                txtSex = GetByOne(sBuf, sBuf)
                txtEmer = ""        'GetByOne(sBuf, sBuf)
                txtReRun = ""       'GetByOne(sBuf, sBuf)
                txtOther = ""       'GetByOne(sBuf, sBuf)
            End If
        End If
    End If
    
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    MsgBox "txtNo_KeyDown 오류 - (" & Err.Description & ")"
End Sub

Private Sub txtSSeq_GotFocus()
    Call Txt_Highlight(txtSSeq)
End Sub

Private Sub txtSSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSSeq_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        txtESeq.SetFocus
'    End If
    
    Call TxtTypeOnlyNumeric(txtSSeq, KeyAscii)
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmInterface 
   Caption         =   "VISION"
   ClientHeight    =   11580
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   22320
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInterface.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmInterface.frx":1272
   ScaleHeight     =   11580
   ScaleWidth      =   22320
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame7 
      Height          =   9825
      Left            =   13290
      TabIndex        =   79
      Top             =   1290
      Width           =   6585
      Begin VB.CommandButton cmdIFTrans3 
         Caption         =   "선택저장"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   97
         Top             =   600
         Width           =   1155
      End
      Begin VB.CommandButton cmdClear3 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   88
         Top             =   600
         Width           =   1155
      End
      Begin FPSpread.vaSpread vasRes3 
         Height          =   1860
         Left            =   120
         TabIndex        =   85
         Top             =   7350
         Visible         =   0   'False
         Width           =   6645
         _Version        =   393216
         _ExtentX        =   11721
         _ExtentY        =   3281
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   8
         MaxRows         =   10
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":14F5
      End
      Begin VB.TextBox txtBarNum3 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1620
         TabIndex        =   80
         Top             =   600
         Width           =   2235
      End
      Begin FPSpread.vaSpread vasID3 
         Height          =   8715
         Left            =   60
         TabIndex        =   81
         Top             =   1020
         Width           =   6375
         _Version        =   393216
         _ExtentX        =   11245
         _ExtentY        =   15372
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   16
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   17
         MaxRows         =   20
         MoveActiveOnFocus=   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":1B44
      End
      Begin VB.Label lblCom3 
         Caption         =   "1번 기기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   150
         TabIndex        =   91
         Top             =   180
         Width           =   6285
      End
      Begin VB.Label Label7 
         Caption         =   "바코드번호 :"
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
         Left            =   120
         TabIndex        =   82
         Top             =   690
         Width           =   1380
      End
   End
   Begin VB.Frame Frame5 
      Height          =   9825
      Left            =   6690
      TabIndex        =   75
      Top             =   1290
      Width           =   6585
      Begin VB.CommandButton cmdIFTrans2 
         Caption         =   "선택저장"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   96
         Top             =   570
         Width           =   1155
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   87
         Top             =   570
         Width           =   1155
      End
      Begin FPSpread.vaSpread vasRes2 
         Height          =   1860
         Left            =   690
         TabIndex        =   84
         Top             =   7380
         Visible         =   0   'False
         Width           =   6645
         _Version        =   393216
         _ExtentX        =   11721
         _ExtentY        =   3281
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   8
         MaxRows         =   10
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":27B5
      End
      Begin VB.TextBox txtBarNum2 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1620
         TabIndex        =   76
         Top             =   600
         Width           =   2235
      End
      Begin FPSpread.vaSpread vasID2 
         Height          =   8715
         Left            =   60
         TabIndex        =   78
         Top             =   1020
         Width           =   6375
         _Version        =   393216
         _ExtentX        =   11245
         _ExtentY        =   15372
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   16
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   17
         MaxRows         =   20
         MoveActiveOnFocus=   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":2E04
      End
      Begin VB.Label lblCom2 
         Caption         =   "2번 기기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   120
         TabIndex        =   90
         Top             =   210
         Width           =   6285
      End
      Begin VB.Label Label4 
         Caption         =   "바코드번호 :"
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
         Left            =   120
         TabIndex        =   77
         Top             =   690
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '위 맞춤
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   22260
      TabIndex        =   25
      Top             =   525
      Width           =   22320
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   90
         TabIndex        =   37
         Top             =   -60
         Width           =   19755
         Begin VB.PictureBox picLogin 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   10530
            Picture         =   "frmInterface.frx":3A75
            ScaleHeight     =   285
            ScaleWidth      =   285
            TabIndex        =   93
            Top             =   240
            Width           =   315
         End
         Begin VB.CheckBox chkMemoSave 
            Caption         =   "학부메모저장"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   8520
            TabIndex        =   92
            Top             =   270
            Value           =   1  '확인
            Width           =   1635
         End
         Begin VB.CommandButton cmdIFTrans 
            Caption         =   "선택저장"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5370
            TabIndex        =   55
            Top             =   210
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.CommandButton cmdIFClear 
            Caption         =   "전체Clear"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6600
            TabIndex        =   54
            Top             =   210
            Width           =   1365
         End
         Begin VB.CommandButton cmdRsltSearch 
            Caption         =   "결과조회"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4140
            TabIndex        =   53
            Top             =   210
            Width           =   1155
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "워크조회"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   17100
            TabIndex        =   38
            Top             =   300
            Visible         =   0   'False
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   345
            Left            =   2610
            TabIndex        =   39
            Top             =   240
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   136577025
            CurrentDate     =   40248
         End
         Begin MSComCtl2.DTPicker dtpStartDt 
            Height          =   345
            Left            =   1080
            TabIndex        =   40
            Top             =   240
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   136577025
            CurrentDate     =   40248
         End
         Begin VB.Label lblUser 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   11010
            TabIndex        =   94
            Top             =   330
            Width           =   1485
         End
         Begin VB.Label Label12 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2490
            TabIndex        =   42
            Top             =   330
            Width           =   105
         End
         Begin VB.Label Label20 
            Caption         =   "조회일자"
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
            Left            =   150
            TabIndex        =   41
            Top             =   300
            Width           =   915
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9825
      Left            =   90
      TabIndex        =   24
      Top             =   1290
      Width           =   6585
      Begin VB.CommandButton cmdIFTrans1 
         Caption         =   "선택저장"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4110
         TabIndex        =   95
         Top             =   570
         Width           =   1155
      End
      Begin VB.CommandButton cmdClear1 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5310
         TabIndex        =   86
         Top             =   570
         Width           =   1155
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   1860
         Left            =   510
         TabIndex        =   83
         Top             =   7020
         Visible         =   0   'False
         Width           =   6645
         _Version        =   393216
         _ExtentX        =   11721
         _ExtentY        =   3281
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   8
         MaxRows         =   10
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":3FFF
      End
      Begin VB.TextBox txtBarNum 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1620
         TabIndex        =   73
         Top             =   600
         Width           =   2235
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   8715
         Left            =   90
         TabIndex        =   60
         Top             =   1020
         Width           =   6375
         _Version        =   393216
         _ExtentX        =   11245
         _ExtentY        =   15372
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   16
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   17
         MaxRows         =   20
         MoveActiveOnFocus=   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":464E
      End
      Begin VB.Label lblCom1 
         Caption         =   "1번 기기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   150
         TabIndex        =   89
         Top             =   180
         Width           =   6315
      End
      Begin VB.Label Label8 
         Caption         =   "바코드번호 :"
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
         Left            =   120
         TabIndex        =   74
         Top             =   690
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   22260
      TabIndex        =   16
      Top             =   0
      Width           =   22320
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   14460
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   1050
         TabIndex        =   22
         Top             =   90
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   136577024
         CurrentDate     =   40457
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   10080
         Top             =   -30
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         RTSEnable       =   -1  'True
         EOFEnable       =   -1  'True
      End
      Begin MSCommLib.MSComm comEqp2 
         Left            =   10890
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         RTSEnable       =   -1  'True
         EOFEnable       =   -1  'True
      End
      Begin MSCommLib.MSComm comEqp3 
         Left            =   11730
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         RTSEnable       =   -1  'True
         EOFEnable       =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "RapidLab 744"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   4
         Left            =   3780
         TabIndex        =   31
         Top             =   90
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   150
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "RapidLab 744"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   2
         Left            =   3810
         TabIndex        =   20
         Top             =   120
         Width           =   1815
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   6750
         Picture         =   "frmInterface.frx":52BF
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   7905
         Picture         =   "frmInterface.frx":5849
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   9060
         Picture         =   "frmInterface.frx":5DD3
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "포트"
         Height          =   195
         Index           =   0
         Left            =   6240
         TabIndex        =   19
         Top             =   150
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "송신"
         Height          =   195
         Left            =   7425
         TabIndex        =   18
         Top             =   150
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "수신"
         Height          =   195
         Left            =   8550
         TabIndex        =   17
         Top             =   150
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   11175
      Width           =   22320
      _ExtentX        =   39370
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12348
            MinWidth        =   12348
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "2019-07-01"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "오후 7:33"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   9945
      Left            =   16320
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   16305
      Begin VB.Frame Frame6 
         Height          =   585
         Left            =   2850
         TabIndex        =   67
         Top             =   7320
         Visible         =   0   'False
         Width           =   6675
         Begin VB.CommandButton cmdBarInput 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   68
            Top             =   180
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblBarcode 
            Caption         =   "12345"
            Height          =   165
            Index           =   0
            Left            =   1905
            TabIndex        =   72
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label6 
            Caption         =   "환자명 :"
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
            Left            =   4080
            TabIndex        =   71
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lblPname 
            Caption         =   "1234567890ab"
            Height          =   225
            Index           =   0
            Left            =   5130
            TabIndex        =   70
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000008&
            ForeColor       =   &H8000000E&
            Height          =   315
            Left            =   180
            TabIndex        =   69
            Top             =   720
            Width           =   1155
         End
      End
      Begin VB.CheckBox chkWAll 
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1230
         TabIndex        =   66
         Top             =   9420
         Width           =   225
      End
      Begin VB.CommandButton cmdSL 
         Caption         =   "▶"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   660
         TabIndex        =   65
         Top             =   9360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdGetRslt 
         Appearance      =   0  '평면
         Caption         =   "결과받기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10050
         TabIndex        =   63
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txtRCnt 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9570
         TabIndex        =   62
         Text            =   "0"
         Top             =   690
         Width           =   405
      End
      Begin FPSpread.vaSpread vasID1 
         Height          =   1125
         Left            =   4020
         TabIndex        =   61
         Top             =   3900
         Visible         =   0   'False
         Width           =   4995
         _Version        =   393216
         _ExtentX        =   8811
         _ExtentY        =   1984
         _StockProps     =   64
         ButtonDrawMode  =   4
         ColHeaderDisplay=   0
         ColsFrozen      =   16
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   17
         MaxRows         =   20
         MoveActiveOnFocus=   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":635D
      End
      Begin VB.CommandButton cmdPatDelete 
         Caption         =   "선택제외"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdResult 
         Appearance      =   0  '평면
         Caption         =   "엑셀열기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4650
         TabIndex        =   58
         Top             =   0
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox txtPos 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6660
         TabIndex        =   56
         Text            =   "A"
         Top             =   6240
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.ComboBox cboTest 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         ItemData        =   "frmInterface.frx":6F6B
         Left            =   5220
         List            =   "frmInterface.frx":6F6D
         Style           =   2  '드롭다운 목록
         TabIndex        =   52
         Top             =   7080
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CheckBox chkSaveAll 
         Caption         =   "저장포함"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4320
         TabIndex        =   51
         Top             =   7350
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cboChk 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmInterface.frx":6F6F
         Left            =   5250
         List            =   "frmInterface.frx":6F7C
         TabIndex        =   50
         Top             =   5250
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtStopNum 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4410
         TabIndex        =   49
         Text            =   "009999"
         Top             =   6030
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtStartNum 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3600
         TabIndex        =   48
         Text            =   "000001"
         Top             =   6030
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtRack 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6060
         TabIndex        =   47
         Text            =   "1"
         Top             =   6240
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.CommandButton cmdWorkPrint 
         Caption         =   "워크출력"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5010
         TabIndex        =   46
         Top             =   6210
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdExcelExport 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   4740
         TabIndex        =   45
         Top             =   2970
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton Command16 
         Caption         =   "전송테스트"
         Height          =   435
         Left            =   4980
         TabIndex        =   44
         Top             =   270
         Width           =   1215
      End
      Begin VB.TextBox txtTest 
         Height          =   1785
         Left            =   4530
         MultiLine       =   -1  'True
         TabIndex        =   43
         Top             =   750
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.OptionButton optSaveResult 
         Caption         =   "장비"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1110
         TabIndex        =   33
         Top             =   5160
         Width           =   735
      End
      Begin VB.OptionButton optSaveResult 
         Caption         =   "수정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1890
         TabIndex        =   32
         Top             =   5160
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Error Log"
         Height          =   945
         Left            =   180
         TabIndex        =   29
         Top             =   8190
         Width           =   4530
         Begin VB.TextBox txtErrLog 
            Appearance      =   0  '평면
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   30
            Top             =   240
            Width           =   4275
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Print"
         Height          =   2415
         Left            =   180
         TabIndex        =   26
         Top             =   5670
         Width           =   3045
         Begin FPSpread.vaSpread vasPrint 
            Height          =   1035
            Left            =   120
            TabIndex        =   27
            Top             =   1290
            Width           =   2760
            _Version        =   393216
            _ExtentX        =   4868
            _ExtentY        =   1826
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   9
            SpreadDesigner  =   "frmInterface.frx":6F94
         End
         Begin FPSpread.vaSpread vasPrintBuf 
            Height          =   975
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   2715
            _Version        =   393216
            _ExtentX        =   4789
            _ExtentY        =   1720
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpreadDesigner  =   "frmInterface.frx":8A0D
         End
      End
      Begin VB.CheckBox chkBar 
         Caption         =   "BARCODE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   465
         Left            =   3090
         Style           =   1  '그래픽
         TabIndex        =   21
         Top             =   3210
         Value           =   1  '확인
         Width           =   1065
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   945
         Left            =   120
         TabIndex        =   15
         Top             =   2250
         Width           =   1665
         _Version        =   393216
         _ExtentX        =   2937
         _ExtentY        =   1667
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmInterface.frx":8C25
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   945
         Left            =   1860
         TabIndex        =   2
         Top             =   1290
         Width           =   2535
         _Version        =   393216
         _ExtentX        =   4471
         _ExtentY        =   1667
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmInterface.frx":8E3D
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   180
         TabIndex        =   14
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   8
         Top             =   3240
         Width           =   1665
      End
      Begin VB.TextBox txtTemp 
         Height          =   435
         Left            =   2730
         TabIndex        =   7
         Top             =   3690
         Width           =   645
      End
      Begin VB.TextBox Text_ini 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2070
         TabIndex        =   6
         Top             =   3705
         Width           =   645
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   5
         Top             =   3840
         Width           =   1635
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   465
         Left            =   1980
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   3180
         Value           =   1  '확인
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   870
         Left            =   1860
         TabIndex        =   3
         Top             =   2310
         Width           =   2835
         Begin VB.Timer tmrSend 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   2220
            Top             =   300
         End
         Begin VB.Timer tmrReceive 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   1740
            Top             =   300
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   720
            Top             =   270
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList imlStatus 
            Left            =   1140
            Top             =   180
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
                  Picture         =   "frmInterface.frx":9055
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":95EF
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":9B89
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":A123
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":A9B5
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":AB0F
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":AC69
                  Key             =   "NOF"
               EndProperty
            EndProperty
         End
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   1695
         _Version        =   393216
         _ExtentX        =   2990
         _ExtentY        =   1720
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmInterface.frx":ADC3
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1035
         Left            =   1860
         TabIndex        =   10
         Top             =   240
         Width           =   2505
         _Version        =   393216
         _ExtentX        =   4419
         _ExtentY        =   1826
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmInterface.frx":AFDB
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   1260
         Width           =   1695
         _Version        =   393216
         _ExtentX        =   2990
         _ExtentY        =   1720
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmInterface.frx":B1F3
      End
      Begin FPSpread.vaSpread vasExcel 
         Height          =   1155
         Left            =   5010
         TabIndex        =   57
         Top             =   8040
         Visible         =   0   'False
         Width           =   2085
         _Version        =   393216
         _ExtentX        =   3678
         _ExtentY        =   2037
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmInterface.frx":B40B
      End
      Begin VB.Label Label2 
         Caption         =   "결과갯수"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9060
         TabIndex        =   64
         Top             =   690
         Width           =   465
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과적용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   36
         Top             =   5250
         Width           =   780
      End
      Begin VB.Label lblSaveSeq 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "99999"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   2790
         TabIndex        =   35
         Top             =   5250
         Width           =   615
      End
      Begin VB.Label lblExamDate 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "20160202"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   3600
         TabIndex        =   34
         Top             =   5250
         Width           =   1005
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   2880
         TabIndex        =   13
         Top             =   4650
         Width           =   465
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   3390
         TabIndex        =   12
         Top             =   4650
         Width           =   435
      End
   End
   Begin VB.Menu MnMain 
      Caption         =   "Main"
      Begin VB.Menu MnPrint 
         Caption         =   "인쇄"
         Begin VB.Menu MnPrintLand 
            Caption         =   "가로인쇄"
         End
         Begin VB.Menu MnPrintPort 
            Caption         =   "세로인쇄"
         End
      End
      Begin VB.Menu MnExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "Setting"
      Begin VB.Menu MnTConfig 
         Caption         =   "통신설정"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "코드설정"
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "Send"
      Begin VB.Menu MnTransAuto 
         Caption         =   "Auto"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnTransManual 
         Caption         =   "Manual"
      End
   End
   Begin VB.Menu MnMode 
      Caption         =   "Mode"
      Visible         =   0   'False
      Begin VB.Menu MnModeBarcode 
         Caption         =   "Barcode"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnModeWorkList 
         Caption         =   "WorkList"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Const colSpecNo = 0     '미사용
'Const colCheckBox = 1
'Const colSAVESEQ = 2    '저장순번(날짜별)
'Const colEXAMDATE = 3   '검사일자
'Const colHOSPDATE = 4   '병원접수일자
'Const colBARCODE = 5
'Const colCHARTNO = 6
'Const colPID = 7        '병록번호(내원번호)
'Const colINOUT = 8      '입원/외래
'Const colDISKNO = 9
'Const colPOSNO = 10
'Const colPNAME = 11
'Const colPSEX = 12
'Const colPAGE = 13
'Const colOCNT = 14
'Const colRCNT = 15
'Const colState = 16

'sendflag
'0: Order
'1: Result
'2: Trans
'vasres, vasrres colum
'Const colEQUIPCODE = 1
'Const colEXAMCODE = 2
'Const colEXAMNAME = 3
'Const colMachResult = 4
'Const colRESULT = 5
'Const colSeq = 6
'Const colFLAG = 7
'Const colSubCode = 8

Dim gRow As Long

Dim gsBarCode       As String
Dim gsSampleType    As String
Dim gsPID           As String
Dim gsRackNo        As String
Dim gsPosNo         As String
Dim gsResDateTime   As String
Dim gsSeqNo         As String
Dim gsExamCode      As String
Dim gsExamName      As String
Dim gsOrder         As String
Dim gsResult        As String
Dim gsFlag          As String

Dim gMT             As String
Dim gComState       As Long
Dim gErrState       As Long

Dim strBuffer       As String
Dim strBuffer1       As String
Dim strBuffer2       As String
Dim strBuffer3       As String

Dim strORQN         As String


'===============================
Const SPCLEN As Integer = 10

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const FS  As String = ""
Const RS  As String = ""
Const GS  As String = ""


Dim strRecvData()   As String
Dim intPhase        As Integer
Dim strState        As String
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer
Dim intCrCnt        As Integer
'===============================

'===============================
Dim AckOn   As Boolean
Dim Sample_Seq  As String
Dim aMod    As String
Dim iIID    As String


Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


Dim OFName As OPENFILENAME

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Dim ReceiveData As String


Private Sub chkMode_Click()
    If chkMode.Value = 1 Then
        chkMode.Caption = "Auto"
    Else
        chkMode.Caption = "Manual"
    End If
End Sub

Private Sub chkWAll_Click()
    Dim iRow As Long
    
    With vasID
        If chkWAll.Value = 1 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = colCheckBox
                .Value = 1
            Next iRow
        ElseIf chkWAll.Value = 0 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = colCheckBox
                .Value = 0
            Next iRow
        End If
    End With
    
End Sub

Private Sub cmdBarInput_Click()
    If cmdBarInput.Caption = "+" Then
        cmdBarInput.Caption = "-"
        txtBarNum.Visible = True
        txtBarNum.SetFocus
    Else
        cmdBarInput.Caption = "+"
        txtBarNum.Visible = False
    End If
End Sub


Sub SaveExcel(Filename As String, argSpread As vaSpread)

On Error Resume Next

' Excel Object Library 와 연결합니다.
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim iRow As Integer
Dim iCol As Integer
Dim i As Integer

    Set xlApp = CreateObject("Excel.Application")
    
    xlApp.DisplayAlerts = False
    
    Set xlBook = xlApp.Workbooks.Add
    
    Set xlSheet = xlBook.Worksheets(1)
     
    For iRow = 0 To argSpread.DataRowCnt
        For iCol = 1 To argSpread.DataColCnt
            argSpread.Row = iRow
            argSpread.Col = iCol
            xlSheet.Cells(iRow + 1, iCol) = argSpread.Text
        Next iCol
    Next iRow
    
    xlBook.SaveAs (Filename)
    xlApp.Quit


End Sub

Private Sub cmdClear1_Click()
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasID.MaxRows = 0
    
    vasRes.MaxRows = 0
    
    txtBarNum.Text = ""

End Sub

Private Sub cmdClear2_Click()
    
    SetForeColor vasID2, 1, vasID2.MaxRows, 1, vasID2.MaxCols, 0, 0, 0
    
    SetForeColor vasRes2, 1, vasRes2.MaxRows, 1, vasRes2.MaxCols, 0, 0, 0
    
    vasID2.MaxRows = 0
    
    vasRes2.MaxRows = 0
    
    txtBarNum2.Text = ""

End Sub

Private Sub cmdClear3_Click()
    
    SetForeColor vasID3, 1, vasID3.MaxRows, 1, vasID3.MaxCols, 0, 0, 0
    
    SetForeColor vasRes3, 1, vasRes3.MaxRows, 1, vasRes3.MaxCols, 0, 0, 0
    
    vasID3.MaxRows = 0
    
    vasRes3.MaxRows = 0
    
    txtBarNum3.Text = ""

End Sub

Private Sub cmdExcelExport_Click()

    Dim iRow As Integer
    Dim j As Integer
    
    Dim sCurDate As String
    Dim sSerDate As String
    Dim sHead As String
    Dim sFoot As String
    Dim sFileName As String
    
    Dim sA1c As String
    Dim sIFCC As String
    Dim seAg As String
    Dim blnWrite As Variant
    
    ClearSpread vasPrint

    blnWrite = False
    vasPrint.MaxRows = vasID.MaxRows
    vasPrint.MaxCols = vasID.MaxCols
    
    For iRow = 1 To vasID.DataRowCnt
        vasID.Row = iRow
        vasID.Col = 1
            
        If vasID.Value = 1 Then
            If blnWrite = False Then
                For j = 1 To vasID.MaxCols
                    SetText vasPrint, Trim(GetText(vasID, 0, j)), 0, j
                Next
            End If
            
            For j = 1 To vasID.MaxCols
                SetText vasPrint, Trim(GetText(vasID, iRow, j)), iRow, j
            Next
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "저장할 자료가 없습니다.", , "알 림"
        Exit Sub
    Else
        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
        CommonDialog1.ShowSave
        sFileName = CommonDialog1.Filename
        SaveExcel sFileName, vasPrint
        MsgBox "엑셀 저장완료", vbOKOnly + vbInformation, Me.Caption
    End If
    
End Sub

Private Sub cmdGetRslt_Click()
    Dim strSendData As String
    
    strSendData = "0" & vbTab & "GET" & vbTab & "0" & vbTab & Trim(txtRCnt.Text) - 1
    
    Winsock1.SendData strSendData & vbLf
    SetRawData "[Tx]" & strSendData & vbLf
End Sub

Private Sub cmdIFClear_Click()
    Dim i As Integer

    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasID2, 1, vasID2.MaxRows, 1, vasID2.MaxCols, 0, 0, 0
    SetForeColor vasID3, 1, vasID3.MaxRows, 1, vasID3.MaxCols, 0, 0, 0
    
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    SetForeColor vasRes2, 1, vasRes2.MaxRows, 1, vasRes2.MaxCols, 0, 0, 0
    SetForeColor vasRes3, 1, vasRes3.MaxRows, 1, vasRes3.MaxCols, 0, 0, 0
    
    vasID.MaxRows = 0
    vasID2.MaxRows = 0
    vasID3.MaxRows = 0
    
    vasRes.MaxRows = 0
    vasRes2.MaxRows = 0
    vasRes3.MaxRows = 0
    
    gRow = 0
    
    txtRack.Text = "1"
    txtPos.Text = "A"
    
    txtBarNum.Text = ""
    txtBarNum2.Text = ""
    txtBarNum3.Text = ""
    
End Sub

Private Sub cmdIFTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
            
            Res = SaveTransDataW(lRow, vasID)
        
            If Res = -1 Then
                SetForeColor vasID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID, "Failed", lRow, colState
            Else
                vasID.Row = lRow
                vasID.Col = 1
                vasID.Value = 1
                
                SetBackColor vasID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID, "Trans", lRow, colState
                
                      SQL = " UPDATE PATRESULT SET " & vbCrLf
                SQL = SQL & "  SENDFLAG = '2' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & Trim(GetText(vasID, lRow, colBARCODE)) & "' "
                
                Res = SendQuery(gLocal, SQL)
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasID.Row = lRow
            vasID.Col = 1
            vasID.Value = 0
        End If
    Next lRow
End Sub

Private Sub cmdIFTrans1_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
            
            Res = SaveTransDataW(lRow, vasID)
        
            If Res = -1 Then
                SetForeColor vasID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID, "Failed", lRow, colState
            Else
                vasID.Row = lRow
                vasID.Col = 1
                vasID.Value = 1
                
                SetBackColor vasID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID, "Trans", lRow, colState
                
                      SQL = " UPDATE PATRESULT SET " & vbCrLf
                SQL = SQL & "  SENDFLAG = '2' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & Trim(GetText(vasID, lRow, colBARCODE)) & "' "
                
                Res = SendQuery(gLocal, SQL)
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasID.Row = lRow
            vasID.Col = 1
            vasID.Value = 0
        End If
    Next lRow
End Sub

Private Sub cmdIFTrans2_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID2.DataRowCnt
        vasID2.Row = lRow
        vasID2.Col = 1
        If vasID2.Value = 1 Then
            
            Res = SaveTransDataW(lRow, vasID2)
        
            If Res = -1 Then
                SetForeColor vasID2, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID2, "Failed", lRow, colState
            Else
                vasID2.Row = lRow
                vasID2.Col = 1
                vasID2.Value = 1
                
                SetBackColor vasID2, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID2, "Trans", lRow, colState
                
                      SQL = " UPDATE PATRESULT SET " & vbCrLf
                SQL = SQL & "  SENDFLAG = '2' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & Trim(GetText(vasID2, lRow, colBARCODE)) & "' "
                
                Res = SendQuery(gLocal, SQL)
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasID2.Row = lRow
            vasID2.Col = 1
            vasID2.Value = 0
        End If
    Next lRow
End Sub

Private Sub cmdIFTrans3_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID3.DataRowCnt
        vasID3.Row = lRow
        vasID3.Col = 1
        If vasID3.Value = 1 Then
            
            Res = SaveTransDataW(lRow, vasID3)
        
            If Res = -1 Then
                SetForeColor vasID3, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID3, "Failed", lRow, colState
            Else
                vasID3.Row = lRow
                vasID3.Col = 1
                vasID3.Value = 1
                
                SetBackColor vasID3, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID3, "Trans", lRow, colState
                
                      SQL = " UPDATE PATRESULT SET " & vbCrLf
                SQL = SQL & "  SENDFLAG = '2' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & Trim(GetText(vasID3, lRow, colBARCODE)) & "' "
                
                Res = SendQuery(gLocal, SQL)
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasID3.Row = lRow
            vasID3.Col = 1
            vasID3.Value = 0
        End If
    Next lRow
End Sub

Private Sub cmdPatDelete_Click()
    Dim i As Integer
    Dim j As Integer
    
    j = 0
    With vasID
        For i = .DataRowCnt To 1 Step -1
            .Row = i
            .Col = colCheckBox
            If .Value = "1" Then
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
                j = j + 1
            End If
        Next
    End With
    
End Sub


'엑셀 파일을 그리드에 넣기
Private Sub Excel_Open()
    Dim xlApp   As New Excel.Application
    Dim XLappWS As Worksheet
    Dim lngSCnt As Long
    Dim lngSColCnt(100) As Long
    Dim dummy       As String
    Dim strRowData  As Variant
    Dim lngRowCnt   As Long
    Dim chk_str     As String
    Dim dummy_max   As Long
    Dim lngTotColCnt   As Long
    Dim lngTotRowCnt   As Long
    Dim i, j, k     As Long

    
    
'Dim xlapp As New Excel.Application
'Dim xlapp_worksheet As Worksheet
'Dim sheet_count As Long
'Dim sheet_col_count(100) As Long
'Dim i, j, k As Long
'Dim dummy As String
'Dim row_data As Variant
'Dim row_cnt As Long
'Dim chk_str As String
'Dim dummy_max As Long
'Dim tot_col_count As Long
'Dim tot_row_count As Long
    
    lngTotColCnt = 0
    lngTotRowCnt = 0
    
    
    '엑셀 열기
    CommonDialog1.Filter = "Excel(*.xlsx)|*.xlsx|Excel(*.xls)|*.xls"
    CommonDialog1.Action = 1
    
    
    If CommonDialog1.FileTitle = "" Then
        Exit Sub
    End If
    
    xlApp.Workbooks.Open (Trim(CommonDialog1.Filename))
    
    lngSCnt = xlApp.Worksheets.Count
    
    '-- 전체 워크시트 불러오기와서 '임시.txt' 파일로 저장
    For i = 1 To lngSCnt
        Set XLappWS = xlApp.Worksheets(i)
        XLappWS.Activate
        lngSColCnt(i) = XLappWS.UsedRange.Columns.Count
        xlApp.DisplayAlerts = False
    
        '''xlApp.ActiveWorkbook.SaveAs App.Path & "\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 공용
        xlApp.ActiveWorkbook.SaveAs "C:\CFX_EXCEL\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 공용
        
        
        'XLappWS.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False ==>엑셀 2000용
        'ActiveWorkbook.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False  ===>엑셀 2003용
    Next i
    
    xlApp.Quit
    Set XLappWS = Nothing
    Set xlApp = Nothing
    
    '-- 전체 엑셀의 MAX cols값 추출
    dummy_max = 0
    For i = 1 To lngSCnt
        If lngSColCnt(i) >= dummy_max Then
            dummy_max = lngSColCnt(i)
        End If
    Next i
    lngTotColCnt = dummy_max
    
    '전체 row값 추출
    For i = 1 To lngSCnt
'''        Open (App.Path & "\" & Trim(i) & ".txt") For Input As #1
        Open ("C:\CFX_EXCEL\" & Trim(i) & ".txt") For Input As #1
        While Not EOF(1)
            Line Input #1, dummy
            strRowData = Split(Trim(dummy), Chr(9))
            chk_str = ""
            For j = 0 To UBound(strRowData)
                chk_str = chk_str & strRowData(j)
            Next j
            If Len(Trim(dummy)) > 0 Then
                lngTotRowCnt = lngTotRowCnt + 1
            End If
        Wend
        Close #1
    Next i
    
    '-- 그리드 초기화
    vasExcel.MaxRows = 0
    vasExcel.MaxRows = lngTotRowCnt
    vasExcel.MaxCols = lngTotColCnt
    
    '-- 그리드에 출력
    For i = 1 To lngSCnt
        '''Open (App.Path & "\" & Trim(i) & ".txt") For Input As #1
        Open ("C:\CFX_EXCEL\" & Trim(i) & ".txt") For Input As #1
        While Not EOF(1)
            Line Input #1, dummy
            strRowData = Split(Trim(dummy), Chr(9))
            chk_str = ""
            For j = 0 To UBound(strRowData)
                chk_str = chk_str & strRowData(j)
            Next j
            If Len(chk_str) > 0 Then
                lngRowCnt = lngRowCnt + 1
                For j = 0 To UBound(strRowData)
                    Call vasExcel.SetText(j + 1, lngRowCnt, CStr(strRowData(j)))
                Next j
            End If
        Wend
        Close #1
    Next i

'    Call SpreadSheetSort(vasExcel, 6, 2)
    With vasExcel
        .Col = 1: .Col2 = .MaxCols
        .Row = 2: .Row2 = .DataRowCnt
        .SortBy = 0
        .SortKey(1) = 2       '정렬키 열번호
        .SortKey(2) = 6       '정렬키 열번호

        .SortKeyOrder(1) = SortKeyOrderAscending
        .SortKeyOrder(2) = SortKeyOrderAscending

        .Action = ActionSort
    End With


End Sub



Private Function SeqSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = brSeq Then
                SeqSearch = sCnt 'brSeq
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function



'Private Sub cmdResult_Click()
'    Dim sSeq As String
'    Dim sBarcode As String
'    Dim strEqpResult As String
'    Dim strLisResult As String
'    Dim strIntBase As String
'    Dim lsExamCode As String
'    Dim lsExamName As String
'    Dim lsSeqNo As String
'    Dim lsResRow    As String
'    Dim lsEquipRes As String
'    Dim lsResult_Buff As String
'
'    Dim lRow As Integer
'    Dim lRow1 As Integer
'    Dim intRow As Integer
'    Dim sWellOld As String
'    Dim sWell As String
'    Dim sExamCode As String
'    Dim sExamName As String
'    Dim sEquipCode As String
'    Dim sItemCode As String
'    Dim strAge As String
'    Dim strSex As String
'    Dim strPtno As String
'    Dim strPtname As String
'    Dim varTmp As Variant
'    Dim intTstCnt As Integer
'    Dim intCol   As Integer
'
'
'
'    ReceiveData = ENQ
'    ReceiveData = ReceiveData & "1H|\^&|||AQT90 FLEX^AQT90 FLEX||||||||1|20100901160446" & vbCr
'    ReceiveData = ReceiveData & "56" & vbCr
'    ReceiveData = ReceiveData & "2P|1||1020135856||^|||||||||^|^^^|^|^||||||||" & vbCr
'    ReceiveData = ReceiveData & "83" & vbCr
'    ReceiveData = ReceiveData & "3O|1||Sample #^250|^^^|||||||||||Whole Blood^||||||||||F" & vbCr
'    ReceiveData = ReceiveData & "5A" & vbCr
'    ReceiveData = ReceiveData & "4C|1|I|1196^A default Hct value was used|I" & vbCr
'    ReceiveData = ReceiveData & "7B" & vbCr
'    ReceiveData = ReceiveData & "5R|1|^^^Hct^I|0.420000|||||F|||20100901100549|20100901100549" & vbCr
'    ReceiveData = ReceiveData & "D6" & vbCr
'    ReceiveData = ReceiveData & "6R|2|^^^TnI^M|<0.010|ug/L||N||F||||" & vbCr
'    ReceiveData = ReceiveData & "94" & vbCr
'    ReceiveData = ReceiveData & "7R|3|^^^CKMB^M|<2.0|ug/L||N||F||||" & vbCr
'    ReceiveData = ReceiveData & "49" & vbCr
'    ReceiveData = ReceiveData & "0R|4|^^^D-dimer^M|67.716374|ng/L||N||F||||" & vbCr
'    ReceiveData = ReceiveData & "2A" & vbCr
'    ReceiveData = ReceiveData & "1L|1|N" & vbCr
'    ReceiveData = ReceiveData & "04" & vbCr
'    ReceiveData = ReceiveData & ""
'
'
'
'    Screen.MousePointer = 11
'
'    vasExcel.MaxRows = 0
'
'    Call Excel_Open
'
'    intTstCnt = 0
'
'    With vasExcel
'        For intRow = 2 To .DataRowCnt
'
'            .GetText 6, intRow, varTmp: sSeq = varTmp
'            .GetText 2, intRow, varTmp: sWell = varTmp
'            If sSeq <> "" Then
'                With mResult
'                    .BarNo = sSeq
'                    .RsltDate = Format(Now, "yyyymmddhhmmss")
'                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
'                    .RackNo = Val(Mid(sWell, 2))
'                    .TubePos = Mid(sWell, 1, 1)
'                End With
'
'                .GetText 3, intRow, varTmp: strIntBase = varTmp
'                '.GetText 10, intRow, varTmp: strEqpResult = varTmp
'
'                If strIntBase = "FAM" Then
'                    Call SetPatInfo(sSeq)
'
'                    vasID.GetText colBARCODE, gRow, varTmp: sBarcode = varTmp
'                    SetText vasID, "Result", gRow, colState
'
'                    '-- 채널
'                    .GetText 3, intRow, varTmp: strIntBase = varTmp
'
'                    '-- 결과
'                    .GetText 10, intRow, varTmp: strEqpResult = varTmp
'
'                    If Val(strEqpResult) = 0 Then
'                        strLisResult = "Not-Detected"
'                    Else
'                       ' strLisResult = CSng(strEqpResult)
'                        strLisResult = Convert2EXP(strEqpResult, "")
'                    End If
'
'
'                    If strLisResult <> "" Then
'                              SQL = "Select examcode, examname, seqno "
'                        SQL = SQL & "  From EQPMASTER"
'                        SQL = SQL & " Where equipno = '" & gEquip & "' "
'                        SQL = SQL & "   AND EXAMNAME = '" & Trim(mGetP(cboTest.Text, 1, "|")) & "'"
'                        SQL = SQL & "   and equipcode = '" & strIntBase & "' "
'                        SQL = SQL & "   and examcode in (" & gOrderExam & ") "      '"'36721','36722','36723','36724'"
'
'                        Res = GetDBSelectColumn(gLocal, SQL)
'
'                        If Res > 0 Then
'                            lsExamCode = Trim(gReadBuf(0))
'                            lsExamName = Trim(gReadBuf(1))
'                            lsSeqNo = Trim(gReadBuf(2))
'
'                            lsResRow = vasRes.DataRowCnt + 1
'                            If vasRes.MaxRows < lsResRow Then
'                                vasRes.MaxRows = lsResRow
'                            End If
'
'                            '소수점 처리, 결과 형태 처리
'                            'lsEquipRes = strLisResult
'                            'strLisResult = SetResult(strLisResult, strIntBase)
'                            'lsResult_Buff = strLisResult
'
'                            For intCol = colState + 1 To vasID.MaxCols
'                                If lsExamCode = gArrEquip(intCol - colState, 3) Then
'                                    SetText vasID, strLisResult, gRow, intCol
'                                    Exit For
'                                End If
'                            Next
'
'                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
'                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
'                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
'                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
'                            SetText vasRes, strLisResult, lsResRow, colRESULT          '결과
'                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
''                            SetText vasRes, strComm, lsResRow, 7                    'Flag
'                            '-- 로컬 저장
'                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
'
'                            lsResult_Buff = ""
'
'                        Else
'                            '-- 오더 없을 경우
'                                  SQL = "Select examcode, examname, seqno "
'                            SQL = SQL & "  From EQPMASTER"
'                            SQL = SQL & " Where equipno = '" & gEquip & "' "
'                            SQL = SQL & "   AND EXAMNAME = '" & Trim(mGetP(cboTest.Text, 1, "|")) & "'"
'                            SQL = SQL & "   and equipcode = '" & strIntBase & "' "
'                            Res = GetDBSelectColumn(gLocal, SQL)
'
'                            If Res > 0 Then
'                                lsExamCode = Trim(gReadBuf(0))
'                                lsExamName = Trim(gReadBuf(1))
'                                lsSeqNo = Trim(gReadBuf(2))
'
'                                lsResRow = vasRes.DataRowCnt + 1
'                                If vasRes.MaxRows < lsResRow Then
'                                    vasRes.MaxRows = lsResRow
'                                End If
'
'                                '소수점 처리, 결과 형태 처리
'                                'lsEquipRes = strLisResult
'                                'strLisResult = SetResult(strLisResult, strIntBase)
'                                'lsResult_Buff = strLisResult
'
'                                For intCol = colState + 1 To vasID.MaxCols
'                                    If lsExamCode = gArrEquip(intCol - colState, 3) Then
'                                        SetText vasID, strLisResult, gRow, intCol
'                                        Exit For
'                                    End If
'                                Next
''
'                                SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
'                                SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
'                                SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
'                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
'                                SetText vasRes, strLisResult, lsResRow, colRESULT          '결과
'                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
''                                SetText vasRes, strComm, lsResRow, 7                    'Flag
'                                '-- 로컬 저장
'                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
'
'
'                                lsResult_Buff = ""
'                                strState = ""
'                            End If
'                        End If
'                    End If
'
'                    strState = "R"
'                End If
'            End If
'        Next
'    End With
'
'    Screen.MousePointer = 0
'
'End Sub


Function Convert2EXP(ByVal srcV#, Optional fmt$) As String
    Dim mul%, dat#, sign$
    
    If srcV# = 0 Then
        Convert2EXP = "0E+00"
        Exit Function
    End If
    
    If srcV# < 0 Then
        sign$ = "-"
        srcV# = Abs(srcV#)
    Else
        sign$ = ""
    End If
    
    mul% = Int(Log(srcV#) / Log(10))
    dat# = srcV# * 10 ^ (mul% * -1)
        
    If fmt$ = "" Then
        Convert2EXP = sign$ & dat# & "E" & Format$(mul%, "+00;-00")
    Else
        Convert2EXP = sign$ & Format$(dat#, fmt$) & "E" & Format$(mul%, "+00;-00")
    End If
    
    If Right(Convert2EXP, 2) = "01" Then    '9.704507E+01
        Convert2EXP = Mid(Convert2EXP, 1, InStr(Convert2EXP, "E") - 1)
        Convert2EXP = Convert2EXP * 10
        Convert2EXP = Format(Convert2EXP, "#.#0")
    ElseIf Right(Convert2EXP, 2) = "02" Then
        Convert2EXP = Mid(Convert2EXP, 1, InStr(Convert2EXP, "E") - 1)
        Convert2EXP = Convert2EXP * 100
        Convert2EXP = Format(Convert2EXP, "#.#0")
    Else
        Convert2EXP = Format(Mid(Convert2EXP, 1, InStr(Convert2EXP, "E") - 1), "#.#0") & "X10^" & Val(Mid(Convert2EXP, InStr(Convert2EXP, "E") + 2))
        'Convert2EXP = Convert2EXP * 100
        'Convert2EXP = Format(Convert2EXP, "#.#0")
    End If
    
End Function



Private Sub cmdRsltSearch_Click()
    Dim iRow As Long
    Dim strDate As String
    Dim strSaveSeq As String
    Dim strChart As String
    Dim RS          As ADODB.Recordset
    Dim i As Integer
    Dim blnSame As Boolean
    Dim intCol As Integer
''
''    ReceiveData = ENQ
''    ReceiveData = ReceiveData & "1H|\^&|||AQT90 FLEX^AQT90 FLEX||||||||1|20100901160446" & vbCr
''    ReceiveData = ReceiveData & "56" & vbCr
''    ReceiveData = ReceiveData & "2P|1||1020135856||^|||||||||^|^^^|^|^||||||||" & vbCr
''    ReceiveData = ReceiveData & "83" & vbCr
''    ReceiveData = ReceiveData & "3O|1||Sample #^250|^^^|||||||||||Whole Blood^||||||||||F" & vbCr
''    ReceiveData = ReceiveData & "5A" & vbCr
''    ReceiveData = ReceiveData & "4C|1|I|1196^A default Hct value was used|I" & vbCr
''    ReceiveData = ReceiveData & "7B" & vbCr
''    ReceiveData = ReceiveData & "5R|1|^^^Hct^I|0.420000|||||F|||20100901100549|20100901100549" & vbCr
''    ReceiveData = ReceiveData & "D6" & vbCr
''    ReceiveData = ReceiveData & "6R|2|^^^TnI^M|<0.010|ug/L||N||F||||" & vbCr
''    ReceiveData = ReceiveData & "94" & vbCr
''    ReceiveData = ReceiveData & "7R|3|^^^CKMB^M|<2.0|ug/L||N||F||||" & vbCr
''    ReceiveData = ReceiveData & "49" & vbCr
''    ReceiveData = ReceiveData & "0R|4|^^^D-dimer^M|67.716374|ng/L||N||F||||" & vbCr
''    ReceiveData = ReceiveData & "2A" & vbCr
''    ReceiveData = ReceiveData & "1L|1|N" & vbCr
''    ReceiveData = ReceiveData & "04" & vbCr
''    ReceiveData = ReceiveData & ""
''
''    Call comEqp_OnComm
''
''    Exit Sub
    
    
    ClearSpread vasID
    ClearSpread vasRes

    vasID.MaxRows = 0
    vasRes.MaxRows = 0
          
          SQL = " SELECT '', SAVESEQ, MID(EXAMDATE,1,8) AS EXAMDATE, HOSPDATE AS 접수일자, BARCODE AS 바코드번호, CHARTNO AS 차트번호, PID AS 내원번호, PNAME AS 이름,PSEX AS 성별, PAGE AS 나이, DISKNO, POSNO, EXAMCODE, RESULT, REFFLAG, SENDFLAG,INOUT " & vbCrLf
    SQL = SQL & "   FROM PATRESULT " & vbCrLf
    SQL = SQL & "  WHERE MID(EXAMDATE,1,8) Between '" & Format(dtpStartDt, "YYYYMMDD") & "' AND '" & Format(dtpStopDt, "YYYYMMDD") & "'" & vbCrLf
    SQL = SQL & "    AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    SQL = SQL & " ORDER BY EXAMDATE,SAVESEQ,HOSPDATE,BARCODE "
    
    Set RS = cn.Execute(SQL, , 1)

    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With vasID
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    strSaveSeq = GetText(vasID, i, colSAVESEQ)
                    
                    If Trim(RS("접수일자")) = strDate And Trim(RS("SAVESEQ")) = strSaveSeq And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
                    
                    If blnSame = True Then
                        For intCol = colState + 1 To vasID.MaxCols
                            If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                                SetText vasID, Trim(RS.Fields("RESULT")) & "", .MaxRows, intCol
                                If Trim(RS.Fields("REFFLAG")) = "H" Then
                                    .Row = .MaxRows
                                    .Col = intCol
                                    .ForeColor = vbRed
                                ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
                                    .Row = .MaxRows
                                    .Col = intCol
                                    .ForeColor = vbBlue
                                End If
                                Exit For
                            End If
                        Next
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1

                    SetText vasID, "0", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("SAVESEQ")) & "", .MaxRows, colSAVESEQ
                    SetText vasID, Trim(RS.Fields("EXAMDATE")) & "", .MaxRows, colEXAMDATE
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    SetText vasID, Trim(RS.Fields("INOUT")) & "", .MaxRows, colINOUT
                    SetText vasID, Trim(RS.Fields("DISKNO")) & "", .MaxRows, colDISKNO
                    SetText vasID, Trim(RS.Fields("POSNO")) & "", .MaxRows, colPOSNO
                    
                    Select Case Trim(RS.Fields("SENDFLAG")) & ""
                        Case "0": SetText vasID, "에러", .MaxRows, colState
                                  SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 201, 112
                        Case "1": SetText vasID, "결과", .MaxRows, colState
                        Case "2": SetText vasID, "완료", .MaxRows, colState
                                  SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 255, 112
                        Case "3": SetText vasID, "수정", .MaxRows, colState
                                  SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 245, 112
                    End Select
                    
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                            SetText vasID, Trim(RS.Fields("RESULT")) & "", .MaxRows, intCol
                            If Trim(RS.Fields("REFFLAG")) = "H" Then
                                .Row = .MaxRows
                                .Col = intCol
                                .ForeColor = vbRed
                            ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
                                .Row = .MaxRows
                                .Col = intCol
                                .ForeColor = vbBlue
                            End If
                            Exit For
                        End If
                    Next

                End If

                blnSame = False

            End With

            RS.MoveNext
        Loop
    End If
    
    RS.Close
    
    vasID.RowHeight(-1) = 12
    
End Sub

Private Sub GetWorkList(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
          SQL = " SELECT DISTINCT '1', '' AS SN ,'' AS 결과일시, REQ_DT AS 접수일자" & vbCrLf
    SQL = SQL & ", QC_BAR_NO AS 바코드번호, LOT_NO AS 차트번호, REQ_SEQ AS 내원번호, '입원' AS 입외" & vbCrLf
    SQL = SQL & ", '' AS R, '' AS P, REQ_SEQ AS 이름, '남자' AS 성별, REQ_SEQ AS 나이, ITEM_CD AS ITEM " & vbCrLf
    SQL = SQL & "  FROM S2QCS101 " & vbCrLf
    SQL = SQL & " WHERE 1=1 " & vbCrLf
    If pBarNo <> "" Then
        SQL = SQL & "   AND QC_BAR_NO = '" & pBarNo & "'" & vbCrLf
    Else
        SQL = SQL & "   AND REQ_DT BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCrLf
    End If
    'SQL = SQL & "   AND ITEM_CD IN (" & gAllExam & ")" & vbCrLf
    SQL = SQL & " ORDER BY 접수일자, 바코드번호, 차트번호, 내원번호"
    
'    If pBarNo <> "" Then
'        Res = GetDBSelectVas(gServer, SQL, vasID, vasID.MaxRows + 1)
'    Else
'        Res = GetDBSelectVas(gServer, SQL, vasID)
'    End If
    
    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount
                
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub

Private Sub GetWorkList_DADESOFT(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
'''          SQL = " SELECT DISTINCT '1', '' AS SN ,'' AS 결과일시, '' AS 접수일자" & vbCrLf
'''    SQL = SQL & ", '' AS 바코드번호, '' AS 차트번호, '' AS 내원번호, '' AS 입외" & vbCrLf
'''    SQL = SQL & ", '' AS R, '' AS P, '' AS 이름, '' AS 성별, '' AS 나이, '' AS ITEM " & vbCrLf
'''    SQL = SQL & "  FROM S2QCS101 " & vbCrLf
'''    SQL = SQL & " WHERE 1=1 " & vbCrLf
'''    If pBarNo <> "" Then
'''        SQL = SQL & "   AND QC_BAR_NO = '" & pBarNo & "'" & vbCrLf
'''    Else
'''        SQL = SQL & "   AND REQ_DT BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCrLf
'''    End If
'''    'SQL = SQL & "   AND ITEM_CD IN (" & gAllExam & ")" & vbCrLf
'''    SQL = SQL & " ORDER BY 접수일자, 바코드번호, 차트번호, 내원번호"
    
          SQL = " SELECT DISTINCT '1', '' AS SN, '' AS 결과일시, J.접수일자 AS 접수일자," & vbCrLf
    SQL = SQL & "        L.검체번호 AS 바코드번호, A.챠트번호 AS 차트번호, J.접수번호 AS 내원번호,'입원' AS 입외, " & vbCrLf
    SQL = SQL & "        J.진료검사ID AS R, L.진료지원ID AS P,  A.환자이름 AS 이름, A.환자성별 AS 성별, A.환자나이  AS 나이, L.처방코드 + L.서브코드 AS ITEM " & vbCrLf
    SQL = SQL & "   FROM TB_진료검사 L " & vbCrLf
    SQL = SQL & "  INNER JOIN TB_진료지원 J ON (L.진료지원ID=J.진료지원ID) " & vbCrLf
    SQL = SQL & "  INNER JOIN TB_진료일반 A ON (J.진료일자=A.진료일자 AND J.챠트번호=A.챠트번호 AND J.진료번호=A.진료번호) " & vbCrLf
    SQL = SQL & "  Where 1 = 1 " & vbCrLf
    SQL = SQL & "    AND J.접수일자 Between '" & pFrDt & "' and '" & pToDt & "'" & vbCrLf
    SQL = SQL & "    AND L.검사종류 = '" & gDept_Code & "'" & vbCrLf
    SQL = SQL & "    AND L.검사상태 < 5 " & vbCrLf
    If chkSaveAll.Value = "0" Then
        SQL = SQL & "  AND (L.검사결과 = '' OR L.검사결과 IS NULL)"
    End If
    SQL = SQL & "  ORDER BY J.접수일자, J.접수번호"
    
    
'          SQL = " SELECT DISTINCT '1', '' AS SN, '' AS 결과일시, L.접수일자 AS 접수일자," & vbCrLf
'    SQL = SQL & "        L.검체번호 AS 바코드번호, L.챠트번호 AS 차트번호, '55555' AS 내원번호,'입원' AS 입외, " & vbCrLf
'    SQL = SQL & "        L.진료검사ID AS R, L.진료지원ID AS P,  '홍길동' AS 이름, '남자' AS 성별, '35'  AS 나이, L.처방코드 + L.서브코드 AS ITEM " & vbCrLf
'    SQL = SQL & "   FROM TB_진료검사 L " & vbCrLf
'    SQL = SQL & "  Where 1 = 1 " & vbCrLf
'    SQL = SQL & "    AND L.접수일자 Between convert(datetime,'" & pFrDt & "') and convert(datetime,'" & pToDt & "')" & vbCrLf
'    SQL = SQL & "    AND L.검사종류 = '" & gDept_Code & "'" & vbCrLf
'    SQL = SQL & "    AND L.검사상태 < 5 " & vbCrLf
'    If chkSaveAll.Value = "0" Then
'        SQL = SQL & "  AND (검사결과 = '' OR 검사결과 IS NULL)"
'    End If
'    SQL = SQL & "  ORDER BY L.접수일자"
    
    
    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
                
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    
                    SetText vasID, Trim(RS.Fields("R")) & "", .MaxRows, colDISKNO
                    SetText vasID, Trim(RS.Fields("P")) & "", .MaxRows, colPOSNO

                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub

Private Sub GetWorkList_TWIN(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
'             SQL = "Select C.SPECNO , C.SNAME, C.DEPTCODE, DECODE(C.GBIO,'I','입 원 ','O','외 래 ') as GBIO, B.EXAMNAME,  B.MASTERCODE, B.CHANNEL "
          SQL = " SELECT DISTINCT '1', '' AS SN ,'' AS 결과일시, B.JOBDATE AS 접수일자" & vbCrLf
    SQL = SQL & ",       C.SPECNO AS 바코드번호, C.PTNO AS 차트번호, C.JOBNO AS 내원번호, DECODE(C.GBIO,'I','입원','O','외래') AS 입외" & vbCrLf
    SQL = SQL & ", '' AS R, '' AS P, C.SNAME AS 이름, C.SEX AS 성별, C.AGE AS 나이, A.MASTERCODE AS ITEM " & vbCrLf
    SQL = SQL & "  From TW_HSP_OCS.TWEXAM_RESULTC A,"
    SQL = SQL & "       TW_HSP_OCS.TWEXAM_MASTER  B,"
    SQL = SQL & "       TW_HSP_OCS.TWEXAM_SPECMST C"
    SQL = SQL & " Where B.JOBDATE BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCrLf '작업일자
    SQL = SQL & "   And B.EQUCODE1 = '" & gEquipCode & "'" & vbCrLf                     ' 장비코드
    SQL = SQL & "   AND C.STATUS   = '3' " & vbCrLf                                     ' 검사상태
    SQL = SQL & "   And (C.SPECNO  = A.SPECNO) " & vbCrLf
    SQL = SQL & "   And (A.MASTERCODE = B.MASTERCODE)"
    SQL = SQL & " ORDER BY 접수일자, 바코드번호, 차트번호, 내원번호"

    SetRawData "[Sql]" & SQL

    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
                
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub


Private Sub GetWorkList_BIT(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
    '-- BIT
          SQL = " SELECT DISTINCT '1', '' AS SN ,'' AS 결과일시, SUBSTRING(O.OCMACPDTM,1,8) AS 접수일자," & vbCrLf
    SQL = SQL & "        R.RESSPMNUM AS 바코드번호, O.OCMCHTNUM AS 차트번호,R.RESOCMNUM AS 내원번호, '' AS 입외," & vbCrLf
    SQL = SQL & "        '' AS R, '' AS P, P.PBSPATNAM AS 이름, P.PBSSEXTYP AS 성별,'' AS 나이, '' AS ITEM" & vbCrLf
    SQL = SQL & "   FROM DRBITPACK..RESINF AS R, DRBITPACK..OCMINF AS O, DRBITPACK..PBSINF AS P, DRBITPACK..LABMST AS E, DRBITPACK..ODRINF AS W" & vbCrLf
    SQL = SQL & " WHERE O.OCMACPDTM BETWEEN '" & pFrDt & "000000" & "' AND '" & pToDt & "235959" & "'" & vbCrLf
    SQL = SQL & "   AND O.OCMCOMSTT NOT IN ('CN', 'CR', 'VC')" & vbCrLf
    SQL = SQL & "   AND R.RESLABCOD IN (" & gAllExam & ")" & vbCrLf
    SQL = SQL & "   AND R.RESOCMNUM = O.OCMNUM" & vbCrLf
    SQL = SQL & "   AND O.OCMCHTNUM = P.PBSCHTNUM" & vbCrLf
    SQL = SQL & "   AND R.RESOCMNUM = W.ODROCMNUM" & vbCrLf
    SQL = SQL & "   AND R.RESLABCOD = W.ODRCOD" & vbCrLf
    SQL = SQL & "   AND R.RESLABCOD = E.LABCOD" & vbCrLf
    '-- 저장미포함
    If chkSaveAll.Value = "0" Then
        SQL = SQL & "   AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F') " & vbCrLf         '--  'I':중간 'F' 완료"
        SQL = SQL & "   AND W.ODRDELFLG = 'N'" & vbCrLf
        SQL = SQL & "   AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)" & vbCrLf
    End If
    SQL = SQL & " ORDER BY 접수일시, 차트번호, 내원번호"


    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
                
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub

Private Sub GetWorkList_GINUSDLL(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    '-- 지누스
    Dim strRequest  As String
    Dim strResponse As String
    Dim varResponse As Variant
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
    '-- 검사대상자 가져오기
                 strRequest = "jobs" + vbTab + "L" + vbTab
    strRequest = strRequest & "hos_org_no" + vbTab + gGINUS_Parm.HCD + vbTab
    strRequest = strRequest & "fr_ymd" + vbTab + pFrDt + vbTab
    strRequest = strRequest & "to_ymd" + vbTab + pToDt + vbTab
    strRequest = strRequest & "mach_cd" + vbTab + gGINUS_Parm.HCD + vbTab
    strRequest = strRequest & "smp_no" + vbTab + "%" + vbTab + vbCr
    
    strResponse = W2ACALL2("SCC0191A", strRequest, gGINUS_Parm.URL) '-- 바코드로 검사대상 조회(https://211.172.17.66)
    
    strResponse = Mid(strResponse, 90)
    varResponse = Split(strResponse, vbLf)
    
    If UBound(varResponse) > 0 Then
        chkWAll.Value = 1
    Else
        chkWAll.Value = 0
    End If
    
    For i = 0 To UBound(varResponse) - 1
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = UBound(varResponse) - 1
        With vasID
            If .MaxRows = 0 Then
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
                
                SetText vasID, "1", intRow, colCheckBox
                SetText vasID, Mid(mGetP(varResponse(i), 5, vbTab), 1, 8), intRow, colHOSPDATE  '-- 접수일자
                SetText vasID, mGetP(varResponse(i), 2, vbTab), intRow, colBARCODE              '-- 바코드번호
                SetText vasID, mGetP(varResponse(i), 6, vbTab), intRow, colPID                  '-- 내원번호
                SetText vasID, mGetP(varResponse(i), 7, vbTab), intRow, colPNAME                '-- 이름
                Select Case mGetP(varResponse(i), 13, vbTab)                                    '-- 입/외
                    Case "O": SetText vasID, "외래", intRow, colINOUT
                    Case "E": SetText vasID, "응급", intRow, colINOUT
                    Case "I": SetText vasID, "입원", intRow, colINOUT
                End Select
                Call SetOrderColor(mGetP(varResponse(i), 2, vbTab), intRow)
            Else
                '-- 같은 바코드 번호가 있는지 체크..
                intRow = GetSameRowNum(Trim(mGetP(varResponse(i), 2, vbTab)))
                If intRow = 0 Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                    
                    SetText vasID, "1", intRow, colCheckBox
                    SetText vasID, Mid(mGetP(varResponse(i), 5, vbTab), 1, 8), intRow, colHOSPDATE  '-- 접수일자
                    SetText vasID, mGetP(varResponse(i), 2, vbTab), intRow, colBARCODE              '-- 바코드번호
                    SetText vasID, mGetP(varResponse(i), 6, vbTab), intRow, colPID                  '-- 내원번호
                    SetText vasID, mGetP(varResponse(i), 7, vbTab), intRow, colPNAME                '-- 이름
                    Select Case mGetP(varResponse(i), 13, vbTab)                                    '-- 입/외
                        Case "O": SetText vasID, "외래", intRow, colINOUT
                        Case "E": SetText vasID, "응급", intRow, colINOUT
                        Case "I": SetText vasID, "입원", intRow, colINOUT
                    End Select
                    Call SetOrderColor(mGetP(varResponse(i), 2, vbTab), intRow)
                End If
            End If
        End With
        
        '-- 프로그레스바 진행
        frmProgress.Xprog.Value = i + 1
        DoEvents
        
    Next
    
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub


Private Sub SetOrderColor(ByVal pBarNo As String, ByVal pRow As Integer)
    Dim i       As Integer
    Dim intCol  As Integer
    Dim strItem As String
    
    '-- 지누스
    Dim strRequest  As String
    Dim strResponse As String
    Dim varResponse As Variant
    
    
    '-- 검사ITEM 가져오기
                 strRequest = "jobs" + vbTab + "Q" + vbTab
    strRequest = strRequest & "hos_org_no" + vbTab + gGINUS_Parm.HCD + vbTab
    strRequest = strRequest & "smp_no" + vbTab + pBarNo + vbTab
    strRequest = strRequest & "mach_cd" + vbTab + gGINUS_Parm.MCD + vbTab + vbCr
    
    strResponse = W2ACALL2("SCC0191A", strRequest, gGINUS_Parm.URL) '-- 바코드로 검사대상 조회(https://211.172.17.66)
    strResponse = Mid(strResponse, 90)
    varResponse = Split(strResponse, vbLf)
    
    If UBound(varResponse) > 0 Then
        For i = 0 To UBound(varResponse) - 1
            For intCol = colState + 1 To vasID.MaxCols
                If mGetP(varResponse(i), 6, vbTab) = gArrEquip(intCol - colState, 3) Then
                    vasID.Row = pRow
                    vasID.Col = intCol
                    vasID.BackColor = vbYellow
                    '-- 결과저장용 SEQ
                    gArrEquip(intCol - colState, 7) = mGetP(varResponse(i), 3, vbTab) & "|" & mGetP(varResponse(i), 4, vbTab) & "|" & mGetP(varResponse(i), 5, vbTab)
                    Exit For
                End If
            Next intCol
        Next i
    Else
        SetText vasID, "No Order", pRow, colState
    End If
    
End Sub

Private Sub cmdSearch_Click()
                
    Select Case gOCS
        Case "BIT":         Call GetWorkList_BIT(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "TWIN":        Call GetWorkList_TWIN(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "DADESOFT":    Call GetWorkList_DADESOFT(Format(dtpStartDt.Value, "yyyy-mm-dd"), Format(dtpStopDt.Value, "yyyy-mm-dd"))
        Case "GINUSDLL":    Call GetWorkList_GINUSDLL(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "GINUSDB":     Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "BITSMALL":    Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "BITLARGE":    Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "MEDICHART":   Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "JBUNIV":      Call GetWorkList_JBUNIV(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "MSINFOTEC":   Call GetWorkList_MSINFOTEC(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "HMHOSP":      Call GetWorkList_HMHOSP(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "JWINFO":      Call GetWorkList_JWINFO(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "PHILL":       Call GetWorkList_PHILL(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    End Select
    
    vasID.RowHeight(-1) = 12
    vasRes.MaxRows = 0
    
End Sub


Private Sub GetWorkList_JBUNIV(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
    '-- 전북대병원  r010m.SPCCD
    SQL = ""
    SQL = SQL & " SELECT '1', '' AS SN ,'' AS 결과일시, j011m.colldt AS 접수일자, j011m.bcno AS 바코드번호, j010m.bcprtno AS 차트번호" & vbCr
    SQL = SQL & "       , r010m.WKYMD||r010m.WKGRPCD||r010m.WKNO FLWKNO " & vbCr
    SQL = SQL & "       , r010m.WKNO AS 접수번호 " & vbCr
    SQL = SQL & "       , j011m.regno AS 내원번호 " & vbCr
    SQL = SQL & "       , j010m.patnm AS 이름 " & vbCr
    SQL = SQL & "       , j010m.age AS 나이 " & vbCr
    SQL = SQL & "       , j010m.sex AS 성별 " & vbCr
    SQL = SQL & "       , j011m.IOGBN  " & vbCr
    SQL = SQL & "       , j010m.DEPTCD " & vbCr
    SQL = SQL & "       , j010m.WARDNO " & vbCr
    SQL = SQL & "       , j010m.ROOMNO " & vbCr
    SQL = SQL & "       , f72m.testcd AS ITEM " & vbCr
    SQL = SQL & "       , r010m.SPCCD AS SPCCD " & vbCr
    SQL = SQL & "  FROM LJ011M j011m                                     " & vbCr
    SQL = SQL & "       INNER JOIN LJ010M j010m                          " & vbCr
    SQL = SQL & "               ON j011m.bcno  = j010m.bcno              " & vbCr
    SQL = SQL & "              AND j011m.regno = j010m.regno             " & vbCr
    SQL = SQL & "       INNER JOIN LR010M r010m                          " & vbCr
    SQL = SQL & "               ON j011m.bcno   = r010m.bcno             " & vbCr
    SQL = SQL & "              AND j011m.regno  = r010m.regno            " & vbCr
    SQL = SQL & "              AND NVL(r010m.rstflg,'0') = '0'       " & vbCr
    SQL = SQL & "       INNER JOIN LF072M f72m                           " & vbCr
    SQL = SQL & "               ON f72m.eqcd    = '" & gEquipCode & "' " & vbCr
    SQL = SQL & "              AND f72m.testcd  = '" & mGetP(cboTest.Text, 2, "|") & "'   " & vbCr
    SQL = SQL & "              AND r010m.testcd = f72m.testcd            " & vbCr
    SQL = SQL & " WHERE j011m.colldt BETWEEN '" & pFrDt & "000000" & "' AND '" & pToDt & "235959" & "'  " & vbCr
    SQL = SQL & "   and r010m.wkno between '" & txtStartNum.Text & "' AND '" & txtStopNum.Text & "' " & vbCr
    SQL = SQL & "   AND j011m.spcflg  = '4'                        " & vbCr
    SQL = SQL & "   AND NVL(j011m.rstflg, '0')  = '0'            " & vbCr
    SQL = SQL & " UNION                                              " & vbCr
    SQL = SQL & " SELECT '1', '' AS SN ,'' AS 결과일시, j011m.colldt AS 접수일자, j011m.bcno AS 바코드번호, j010m.bcprtno AS 차트번호 " & vbCr
    SQL = SQL & "        , r010m.FLWKNO " & vbCr
    SQL = SQL & "        , r010m.WKNO AS 접수번호 " & vbCr
    SQL = SQL & "        , j011m.regno AS 내원번호 " & vbCr
    SQL = SQL & "        , j010m.patnm AS 이름 " & vbCr
    SQL = SQL & "        , j010m.age AS 나이 " & vbCr
    SQL = SQL & "        , j010m.sex AS 성별 " & vbCr
    SQL = SQL & "        , j011m.IOGBN " & vbCr
    SQL = SQL & "        , j010m.DEPTCD " & vbCr
    SQL = SQL & "        , j010m.WARDNO " & vbCr
    SQL = SQL & "        , j010m.ROOMNO " & vbCr
    SQL = SQL & "       , f72m.testcd AS ITEM " & vbCr
    SQL = SQL & "       , r010m.SPCCD AS SPCCD " & vbCr
    SQL = SQL & "   FROM LJ011M j011m                                " & vbCr
    SQL = SQL & "        INNER JOIN LJ010M j010m                     " & vbCr
    SQL = SQL & "                ON j011m.bcno  = j010m.bcno         " & vbCr
    SQL = SQL & "               AND j011m.regno = j010m.regno        " & vbCr
    SQL = SQL & "        INNER JOIN LM010M r010m                     " & vbCr
    SQL = SQL & "                ON j011m.bcno   = r010m.bcno        " & vbCr
    SQL = SQL & "               AND j011m.regno  = r010m.regno       " & vbCr
    SQL = SQL & "               AND NVL(r010m.rstflg,'0') = '0'  " & vbCr
    SQL = SQL & "        INNER JOIN LF072M f72m                      " & vbCr
    SQL = SQL & "                ON f72m.eqcd    = '" & gEquipCode & "' " & vbCr
    SQL = SQL & "                AND f72m.testcd  = '" & mGetP(cboTest.Text, 2, "|") & "'  " & vbCr
    SQL = SQL & "               AND r010m.testcd = f72m.testcd       " & vbCr
    SQL = SQL & "  WHERE j011m.colldt BETWEEN '" & pFrDt & "000000" & "' AND '" & pToDt & "235959" & "'  " & vbCr
    SQL = SQL & "   and r010m.wkno BETWEEN '" & txtStartNum.Text & "' AND '" & txtStopNum.Text & "' " & vbCr
    SQL = SQL & "    AND j011m.spcflg  = '4'               " & vbCr
    SQL = SQL & "    AND NVL(j011m.rstflg, '0')  = '0'     " & vbCr
    SQL = SQL & "    ORDER BY FLWKNO  " & vbCr

 '   SetRawData "[SQL]" & SQL

    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
        
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    SetText vasID, Trim(RS.Fields("SPCCD")) & "", .MaxRows, colINOUT
                    
                    '.MaxRows = .MaxRows + 1
                    
                    SetText vasID, txtRack.Text, .MaxRows, colDISKNO
                    SetText vasID, txtPos.Text, .MaxRows, colPOSNO
                    
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            txtPos.Text = Chr(Asc(txtPos.Text) + 1)
            If txtPos.Text = "I" Then
                txtPos.Text = "A"
                txtRack.Text = txtRack.Text - 1
            End If
            
            If txtRack.Text = "1" And txtPos.Text = "H" Then
                txtRack.Text = "5"
                txtPos.Text = "A"
            End If
            
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub



Private Sub GetWorkList_MSINFOTEC(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
                '-- 처방일자,처방일련번호,환자명,검체번호,입외구분,일련번호,성별,나이,내원번호,처방코드
    SQL = ""
    SQL = SQL & "Select DISTINCT a.ORDT as 접수일자,'0',b.PANM as 이름,a.SPNO as 바코드번호,a.PAID as 챠트번호, a.OIFL,'0',b.SEXS as 성별,b.AGES as 나이,a.NWNO as 내원번호,a.ORCD as ITEM " & vbCr
    SQL = SQL & "  From LRESULT a, APATINF b" & vbCr
    SQL = SQL & " Where a.ORDT between  '" & pFrDt & "' and '" & pToDt & "'" & vbCr
    SQL = SQL & "   And a.PAID = b.PAID " & vbCr
    SQL = SQL & "   And a.ORCD in (" & gAllExam & ")" & vbCr
    SQL = SQL & "   And a.OKFL <> 'Y' "   '-- 결과확정유무

    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
        
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
'                    For intCol = colState + 1 To vasID.MaxCols
'                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
'                            vasID.Row = .MaxRows
'                            vasID.Col = intCol
'                            vasID.BackColor = vbYellow
'                            Exit For
'                        End If
'                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("챠트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    'SetText vasID, Trim(RS.Fields("SPCCD")) & "", .MaxRows, colINOUT
                    SetText vasID, txtRack.Text, .MaxRows, colDISKNO
                    
                    '.MaxRows = .MaxRows + 1
                    
                    SetText vasID, txtRack.Text, .MaxRows, colDISKNO
                    SetText vasID, txtPos.Text, .MaxRows, colPOSNO
                    
                    txtRack.Text = txtRack.Text + 1
                    
                    If txtRack.Text = "31" Then
                        txtRack.Text = "1"
                    End If
                    
                    
'                    For intCol = colState + 1 To vasID.MaxCols
'                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
'                            vasID.Row = .MaxRows
'                            vasID.Col = intCol
'                            vasID.BackColor = vbYellow
'                            Exit For
'                        End If
'                    Next
                End If
                blnSame = False
            End With
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = iCnt
            DoEvents
                        
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub




Private Sub GetWorkList_HMHOSP(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
                '-- 처방일자,처방일련번호,환자명,검체번호,입외구분,일련번호,성별,나이,내원번호,처방코드
    SQL = ""
    SQL = SQL & "Select DISTINCT a.ORDT as 접수일자,'0',b.PANM as 이름,a.SPNO as 바코드번호,a.PAID as 챠트번호, a.OIFL,'0',b.SEXS as 성별,b.AGES as 나이,a.NWNO as 내원번호,a.ORCD as ITEM " & vbCr
    SQL = SQL & "  From LRESULT a, APATINF b" & vbCr
    SQL = SQL & " Where a.ORDT between  '" & pFrDt & "' and '" & pToDt & "'" & vbCr
    SQL = SQL & "   And a.PAID = b.PAID " & vbCr
    SQL = SQL & "   And a.ORCD in (" & gAllExam & ")" & vbCr
    SQL = SQL & "   And a.OKFL <> 'Y' "   '-- 결과확정유무


    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
        
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
'                    For intCol = colState + 1 To vasID.MaxCols
'                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
'                            vasID.Row = .MaxRows
'                            vasID.Col = intCol
'                            vasID.BackColor = vbYellow
'                            Exit For
'                        End If
'                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("챠트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    'SetText vasID, Trim(RS.Fields("SPCCD")) & "", .MaxRows, colINOUT
                    SetText vasID, txtRack.Text, .MaxRows, colDISKNO
                    
                    '.MaxRows = .MaxRows + 1
                    
                    SetText vasID, txtRack.Text, .MaxRows, colDISKNO
                    SetText vasID, txtPos.Text, .MaxRows, colPOSNO
                    
                    txtRack.Text = txtRack.Text + 1
                    
                    If txtRack.Text = "31" Then
                        txtRack.Text = "1"
                    End If
                    
                    
'                    For intCol = colState + 1 To vasID.MaxCols
'                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
'                            vasID.Row = .MaxRows
'                            vasID.Col = intCol
'                            vasID.BackColor = vbYellow
'                            Exit For
'                        End If
'                    Next
                End If
                blnSame = False
            End With
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = iCnt
            DoEvents
                        
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub

Private Sub GetWorkList_JWINFO(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    

          SQL = "SELECT DISTINCT RECEIPTDATE as 접수일자, SPECIMENNUM as 바코드번호, RECEIPTNO as 챠트번호, IPDOPD, PTNO as 내원번호, SNAME as 이름, LABCODE as ITEM,ORDERCODE"
    SQL = SQL & vbCrLf & "  FROM SLA_LabMaster "
    SQL = SQL & vbCrLf & " WHERE RECEIPTDATE between '" & Format(pFrDt, "####-##-##") & "' and '" & Format(pToDt, "####-##-##") & "'"
    SQL = SQL & vbCrLf & "   AND LABCODE IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND JSTATUS < '3'" & vbLf
    '입원
    If cboChk.ListIndex = 1 Then
        SQL = SQL & "    AND IpdOpd = 1 "
    '외래
    ElseIf cboChk.ListIndex = 2 Then
        SQL = SQL & "    AND IpdOpd = 0 "
    End If

    SQL = SQL & "  ORDER BY RECEIPTDATE "

    Call SetSQLData("워크조회", SQL)

    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
        
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colCHARTNO)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("챠트번호")) = strChart Then
                        blnSame = True
                    End If
                    
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    If Trim(RS.Fields("바코드번호")) & "" = "0" Then
                        SetText vasID, Trim(RS.Fields("챠트번호")) & "", .MaxRows, colBARCODE
                    Else
                        SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    End If
                    SetText vasID, Trim(RS.Fields("챠트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, IIf(Trim(RS.Fields("IPDOPD")) = 1, "입원", "외래"), .MaxRows, colINOUT
                    SetText vasID, Trim(RS.Fields("ORDERCODE")) & "", .MaxRows, colPSEX
                    
                  
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                
                End If
                
                blnSame = False
            End With
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = iCnt
            DoEvents
                        
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub

Private Sub GetWorkList_PHILL(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    Dim sqlRet      As Integer
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    StatusBar1.Panels(3).Text = ""

'    SQL = ""
'    SQL = SQL & "SELECT DISTINCT P.request_date AS 접수일자, P.exam_no AS 내원번호, P.company_code AS 의뢰처, P.chart_no AS 차트번호, p.personal_id, p.person_name AS 이름, " & vbCr
'    SQL = SQL & "       P.worker_code, P.patient_kind, P.person_sex AS 성별, P.person_age AS 나이, R.pro_code AS 처방코드 " & vbCr
'    SQL = SQL & "  FROM trust P, trures R " & vbCr
'    SQL = SQL & " WHERE P.request_date BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCr
'    SQL = SQL & "   AND R.pro_code IN (" & gAllExam & ") " & vbCr
'    SQL = SQL & "   AND R.exam_code <> 'X999' " & vbCr
'    SQL = SQL & "   AND P.request_date = R.request_date " & vbCr
'    SQL = SQL & "   AND P.exam_no = R.exam_no " & vbCr
'    SQL = SQL & " ORDER BY P.request_date, P.exam_no "


    SQL = ""
    SQL = SQL & "SELECT DISTINCT CONVERT(NVARCHAR(50),M.접수일자,112) AS 접수일자, M.접수번호 AS 내원번호,'' AS 의뢰처, M.차트번호 AS 차트번호, M.성명 AS 이름," & vbCr
    SQL = SQL & "       M.성별 AS 성별,M.나이 AS 나이, E.검사코드 AS 처방코드 " & vbCr
    SQL = SQL & "  FROM VW_검사접수 M, VW_검사결과 R, VW_검사코드 E  " & vbCr
    SQL = SQL & " WHERE M.접수일자 = CONVERT(DATE, '" & pFrDt & "')" & vbCr
    SQL = SQL & "   AND E.학부코드 = 'C2' " & vbCr
    SQL = SQL & "   AND E.검사코드 IN (" & gAllExam & ") " & vbCr
    SQL = SQL & "   AND isnull(R.보고여부, 'N') <> 'Y' " & vbCr
    SQL = SQL & "   AND (R.결과값 is null or R.결과값 = '') " & vbCr
    SQL = SQL & "   AND M.접수일자 = R.접수일자 " & vbCr
    SQL = SQL & "   AND M.접수번호 = R.접수번호 " & vbCr
    SQL = SQL & " ORDER BY 접수일자, M.접수번호 "
    
    Call SetSQLData("워크조회", SQL)

    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
                
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("접수일자")) = strDate And Mid(Trim(RS.Fields("접수일자")), 1, 8) & PedLeftStr(Trim(RS.Fields("내원번호")), 5, "0") = strChart Then
                        blnSame = True
                    End If
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("접수일자")) & PedLeftStr(Trim(RS.Fields("내원번호")), 5, "0"), .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("의뢰처")) & "", .MaxRows, colDISKNO
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    
'                    Select Case Trim(RS.Fields("처방코드")) & ""        '처방코드 ??
'                        Case gAssayNM.INHALANT_CD: SetText vasID, "INHALANT", .MaxRows, colINOUT
'                        Case gAssayNM.FOOD_CD:     SetText vasID, "FOOD", .MaxRows, colINOUT
'                        Case gAssayNM.ATOPY_CD:    SetText vasID, "ATOPY", .MaxRows, colINOUT
'                        Case Else:
'                                                   SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 255, 112
'                                                   SetText vasID, "처방오류", .MaxRows, colINOUT
'
'                    End Select
                End If
                blnSame = False
            End With
            
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub


'Private Sub cmdSL_Click()
'    If cmdSL.Caption = "▶" Then
'        cmdSL.Caption = "◀"
'        'vasID.Width = 15225
'        vasID.Width = Frame1.Width - 200
'    Else
'        cmdSL.Caption = "▶"
'        'vasID.Width = 8475
'        vasID.Width = Me.Width - Frame6.Width - 710
'    End If
'
''    Call Form_Resize
'End Sub

Private Sub cmdWorkPrint_Click()
Dim iRow As Integer
Dim j As Integer

Dim sCurDate As String
Dim sSerDate As String
Dim sHead As String
Dim sFoot As String
    
    ClearSpread vasPrint

    j = 1

    vasPrint.RowHeight(-1) = 25.9
    
    For iRow = 1 To vasID.DataRowCnt
        vasID.Row = iRow
        vasID.Col = colCheckBox

        If vasID.Value = 1 Then
            SetText vasPrint, Trim(GetText(vasID, iRow, colBARCODE)), j, 1     '검체번호
            SetText vasPrint, Trim(GetText(vasID, iRow, colCHARTNO)), j, 2     '환자번호
            SetText vasPrint, Trim(GetText(vasID, iRow, colPNAME)), j, 3     '환자이름

            SetText vasPrint, Trim(GetText(vasID, iRow, colPSEX)), j, 4     '성별
            SetText vasPrint, Trim(GetText(vasID, iRow, colPAGE)), j, 5     '나이
            SetText vasPrint, Trim(GetText(vasID, iRow, colHOSPDATE)), j, 7     '처방일자
            SetText vasPrint, Trim(GetText(vasID, iRow, colHOSPDATE)), j, 8     '처방일자
            
            j = j + 1
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    End If
    
    sCurDate = GetDateFull
    
    sSerDate = Trim(dtpStartDt.Value) & " - " & Trim(dtpStopDt.Value)
    
    vasPrint.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
    vasPrint.PrintAbortMsg = "인쇄중 입니다 ..."
    vasPrint.PrintJobName = "WorkList 출력"
    

    sHead = "/fn""궁서체"" /fz""12"" /fb1 /fi0 /fu0 " & "/c" & "▣ WorkList ▣" & "/n/n " & _
            "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/c" & "처방일자 : " & dtpStartDt & " ~ " & dtpStopDt

    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & " 검사실"
    
    vasPrint.PrintHeader = sHead
    vasPrint.PrintFooter = sFoot

    vasPrint.PrintMarginTop = 680
    vasPrint.PrintMarginBottom = 680
'현재 SS가 비대칭으로 출력함
'    vaslist.PrintMarginLeft = 720
    vasPrint.PrintMarginLeft = 0
    vasPrint.PrintMarginRight = 0
    
    vasPrint.PrintColor = True
    vasPrint.PrintGrid = True
    
'Set printing range
    vasPrint.PrintType = 0  'SS_PRINT_ALL(default)

    vasPrint.PrintShadows = True

    vasPrint.Action = 13 'SS_ACTION_PRINT
End Sub

Private Sub comEqp2_OnComm()
    Dim EVMsg       As String
    Dim ERMsg       As String
    Dim Ret         As Long
    Dim strDate     As String
    
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    
On Error GoTo Rst

    Select Case comEqp2.CommEvent
        Case comEvReceive

            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If

            Buffer = comEqp2.Input
            
            SetRawData "[Rx]" & Buffer
            StatusBar1.Panels(3).Text = Buffer
            
            lngBufLen = Len(Buffer)
            
                For i = 1 To lngBufLen
                    BufChar = Mid$(Buffer, i, 1)
                    Select Case intPhase
                        Case 1
                            Select Case BufChar
                                Case "="
                                    strBuffer2 = ""
                                    intCrCnt = 0
                                    intPhase = 2
                                Case Else
                                    strBuffer2 = strBuffer2 & BufChar
                            End Select
                        Case 2
                            Select Case BufChar
                                Case vbCr '## CR 대기
                                    intCrCnt = intCrCnt + 1
                                    If intCrCnt >= 23 Then
                                        Call EditRcvDataASTM2(vasID2)
                                        strBuffer2 = ""
                                        intPhase = 1
                                        intCrCnt = 0
                                    End If
                                    strBuffer2 = strBuffer2 & BufChar
                                Case Else
                                    strBuffer2 = strBuffer2 & BufChar
                            
                            End Select
                            
                            
                            
                    End Select
                Next i
            
        Case comEvSend
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
        
        Case comEvCTS
            EVMsg$ = "CTS 변경 감지"
        Case comEvDSR
            EVMsg$ = "DSR 변경 감지"
        Case comEvCD
            EVMsg$ = "CD 변경 감지"
        Case comEvRing
            EVMsg$ = "전화 벨이 울리는 중"
        Case comEvEOF
            EVMsg$ = "EOF 감지"

        '오류 메시지
        Case comBreak
            ERMsg$ = "중단 신호 수신"
        Case comCDTO
            ERMsg$ = "반송파 검출 시간 초과"
        Case comCTSTO
            ERMsg$ = "CTS 시간 초과"
        Case comDCB
            ERMsg$ = "DCB 검색 오류"
        Case comDSRTO
            ERMsg$ = "DSR 시간 초과"
        Case comFrame
            ERMsg$ = "프레이밍 오류"
        Case comOverrun
            ERMsg$ = "패리티 오류"
        Case comRxOver
            ERMsg$ = "수신 버퍼 초과"
        Case comRxParity
            ERMsg$ = "패리티 오류"
        Case comTxFull
            ERMsg$ = "전송 버퍼에 여유가 없음"
        Case Else
            ERMsg$ = "알 수 없는 오류 또는 이벤트"
    End Select

Exit Sub

Rst:
    StatusBar1.Panels(2).Text = Err.Description
    SetRawData "[Error_2]" & Err.Number & ":" & Err.Description

End Sub

Private Sub comEqp3_OnComm()
    Dim EVMsg       As String
    Dim ERMsg       As String
    Dim Ret         As Long
    Dim strDate     As String
    
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    
On Error GoTo Rst
    
    Select Case comEqp3.CommEvent
        Case comEvReceive

            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If

            Buffer = comEqp3.Input
            
            SetRawData "[Rx]" & Buffer
            StatusBar1.Panels(3).Text = Buffer
            
            lngBufLen = Len(Buffer)
            
                For i = 1 To lngBufLen
                    BufChar = Mid$(Buffer, i, 1)
                    Select Case intPhase
                        Case 1
                            Select Case BufChar
                                Case "="
                                    strBuffer3 = ""
                                    intCrCnt = 0
                                    intPhase = 2
                                Case Else
                                    strBuffer3 = strBuffer3 & BufChar
                            End Select
                        Case 2
                            Select Case BufChar
                                Case vbCr '## CR 대기
                                    intCrCnt = intCrCnt + 1
                                    If intCrCnt >= 23 Then
                                        Call EditRcvDataASTM3(vasID3)
                                        strBuffer3 = ""
                                        intPhase = 1
                                        intCrCnt = 0
                                    End If
                                    strBuffer3 = strBuffer3 & BufChar
                                Case Else
                                    strBuffer3 = strBuffer3 & BufChar
                            
                            End Select
                            
                            
                            
                    End Select
                Next i
            
        Case comEvSend
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
        
        Case comEvCTS
            EVMsg$ = "CTS 변경 감지"
        Case comEvDSR
            EVMsg$ = "DSR 변경 감지"
        Case comEvCD
            EVMsg$ = "CD 변경 감지"
        Case comEvRing
            EVMsg$ = "전화 벨이 울리는 중"
        Case comEvEOF
            EVMsg$ = "EOF 감지"

        '오류 메시지
        Case comBreak
            ERMsg$ = "중단 신호 수신"
        Case comCDTO
            ERMsg$ = "반송파 검출 시간 초과"
        Case comCTSTO
            ERMsg$ = "CTS 시간 초과"
        Case comDCB
            ERMsg$ = "DCB 검색 오류"
        Case comDSRTO
            ERMsg$ = "DSR 시간 초과"
        Case comFrame
            ERMsg$ = "프레이밍 오류"
        Case comOverrun
            ERMsg$ = "패리티 오류"
        Case comRxOver
            ERMsg$ = "수신 버퍼 초과"
        Case comRxParity
            ERMsg$ = "패리티 오류"
        Case comTxFull
            ERMsg$ = "전송 버퍼에 여유가 없음"
        Case Else
            ERMsg$ = "알 수 없는 오류 또는 이벤트"
    End Select

Exit Sub

Rst:
    StatusBar1.Panels(2).Text = Err.Description
    SetRawData "[Error_3]" & Err.Number & ":" & Err.Description

End Sub


Private Sub Command1_Click()

End Sub

'Private Sub Form_Resize()
'    On Error Resume Next
'
'    If frmInterface.ScaleHeight = 0 Then Exit Sub
'
'
'    If cmdSL.Caption = "▶" Then
'        Frame1.Height = frmInterface.ScaleHeight - (Picture2.Top) - 1200
'        vasID.Height = Frame1.Height - 300
'
'        Frame1.Width = frmInterface.ScaleWidth - 200
'        vasID.Width = frmInterface.ScaleWidth - 7300
'
'
'        Frame6.Left = vasID.Width + 300
'        vasRes.Height = vasID.Height - 550
'        vasRes.Left = Frame6.Left
'    Else
'        Frame1.Height = frmInterface.ScaleHeight - (Picture2.Top) - 1200
'        vasID.Height = Frame1.Height - 300
'
'        Frame1.Width = frmInterface.ScaleWidth - 200
'        vasID.Width = frmInterface.ScaleWidth - 300
'
'        'Frame6.Left = frmInterface.ScaleWidth - vasID.Width
'        'vasRes.Height = vasID.Height - 550
'        'vasRes.Left = Frame6.Left
'
'    End If
'
'    Picture2.Width = Frame1.Width
'
'    StatusBar1.Panels(3).Width = Frame1.Width - 8500
'End Sub


Private Sub imgPort_DblClick()
    
    '-- 개발시에만 Remark 풀어서 테스트진행
    If FrmHideControl.Visible = True Then
        Me.Width = 16545
        FrmHideControl.Visible = False
    Else
        Me.Width = 22000
        FrmHideControl.Visible = True
    End If

End Sub





'Private Sub Label1_Click(Index As Integer)
'
'    Call comEqp_OnComm
'
'End Sub

Private Sub lblclear_Click()
    lblChangePID.Caption = ""
    lblChangeBar.Caption = ""
    lblBarcode(0).Caption = ""
    lblPname(0).Caption = ""
    lblSaveSeq.Caption = ""
    lblExamDate.Caption = ""
End Sub

Private Sub Command16_Click()
    
    strBuffer = ":N1    80 81                 00620141422      15 1   7.0  2   4.1  3   0.5  4   4.5  5    34  6    20  7   417  8   239  9    97 14    85 15    14 16   0.7 18    93 19      T54     1 "
    
    strBuffer = txtTest.Text
    
    Call comEqp_OnComm
        

End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
On Error GoTo Rst
    
    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
    
    'Me.Height = 11520
    Me.Width = 20160 '16545
    
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    cmdIFClear_Click
    lblclear_Click
    
    GetSetup
    
    If gSave = "True" Then
        chkMode.Caption = "Auto"
        MnTransAuto.Checked = True
        MnTransManual.Checked = False
        chkMode.Value = 1
    Else
        chkMode.Caption = "Manual"
        MnTransAuto.Checked = False
        MnTransManual.Checked = True
        chkMode.Value = 0
    End If
    
    If gIFMode = "Barcode" Then
        'fraBar.Visible = True
'        fraWork.Visible = False
    
        chkMode.Caption = "Barcode"
        MnModeBarcode.Checked = True
        MnModeWorkList.Checked = False
        chkBar.Value = 1
    Else
        'fraBar.Visible = False
'        fraWork.Visible = True
    
        chkMode.Caption = "WorkList"
        MnModeBarcode.Checked = False
        MnModeWorkList.Checked = True
        chkBar.Value = 0
    End If
    
'    If gScreen = "통합" Then
'        cmdSL.Caption = "◀"
'        vasID.Width = 14595
'    Else
'        cmdSL.Caption = "▶"
'        vasID.Width = 7725
'    End If
    
    frmInterface.StatusBar1.Panels(1).Text = gUserID
        
    cboChk.ListIndex = 0
    
    comEqp.CommPort = gSetup.gPort
    comEqp.RTSEnable = gSetup.gRTSEnable
    comEqp.DTREnable = gSetup.gDTREnable
    comEqp.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If

    lblCom1.Caption = "1번 기기" & Space(30) & "포트 : " & gSetup.gPort
    
    comEqp2.CommPort = gSetup2.gPort
    comEqp2.RTSEnable = gSetup2.gRTSEnable
    comEqp2.DTREnable = gSetup2.gDTREnable
    comEqp2.Settings = gSetup2.gSpeed & "," & gSetup2.gParity & "," & gSetup2.gDataBit & "," & gSetup2.gStopBit

    If comEqp2.PortOpen = False Then
        comEqp2.PortOpen = True
    End If
    
    lblCom2.Caption = "2번 기기" & Space(30) & "포트 : " & gSetup2.gPort
    
    comEqp3.CommPort = gSetup3.gPort
    comEqp3.RTSEnable = gSetup3.gRTSEnable
    comEqp3.DTREnable = gSetup3.gDTREnable
    comEqp3.Settings = gSetup3.gSpeed & "," & gSetup3.gParity & "," & gSetup3.gDataBit & "," & gSetup3.gStopBit

    If comEqp3.PortOpen = False Then
        comEqp3.PortOpen = True
    End If
    
    lblCom3.Caption = "3번 기기" & Space(30) & "포트 : " & gSetup3.gPort
    
'    If comEqp.PortOpen Then
'        frmInterface.StatusBar1.Panels(2).Text = "COM" & comEqp.CommPort & " 포트에 연결 되었습니다"
'        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
'        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
'        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
'    Else
'        frmInterface.StatusBar1.Panels(2).Text = "통신포트에 연결 되지 않았습니다"
'        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
'        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
'        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
'    End If

    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
'    '-- osw 추가
     For i = 1 To 1
        If Not Connect_PRServer Then
            MsgBox "연결되지 않았습니다."
            cn_Server_Flag = False
            Exit Sub
        Else
            cn_Server_Flag = True
        End If
    Next
    
    '-- osw 추가
'    For i = 1 To 1
'        If Not Connect_DRServer Then
'            MsgBox "연결되지 않았습니다."
'            cn_Server_Flag = False
'            Exit Sub
'        Else
'            cn_Server_Flag = True
'        End If
'    Next
    
    GetExamCode
    
    SetExamCode
    
    dtpToday = Date
    dtpStartDt = Date
    dtpStopDt = Date
    
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -30), "yyyymmdd")
    
    SQL = "delete from PATRESULT where examdate < '" & sDate & "'"
    Res = SendQuery(gLocal, SQL)
    
    lblUser.Caption = gUserID
    
    If lblUser.Caption = "" Then
        Call picLogin_Click
    End If
    
'    stInterface.Tab = 0

    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    intCrCnt = 0
    '==============================
    
'    Winsock1.LocalPort = CInt(809)
'    Winsock1.Listen
    
'    Call cmdSL_Click
   
    '-- test
'    vasID.MaxRows = 10
    Exit Sub
    
Rst:
    If Err.Number = "8002" Then
        If (MsgBox("포트 번호가 잘못되었습니다." & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            Resume Next
        Else
            End
        End If
    Else
        If (MsgBox(Err.Description & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            Resume Next
        Else
            End
        End If
    End If
End Sub

Private Sub txtBarcode_Change()

End Sub





Private Sub txtBarNum2_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim intRow  As Integer
    
    If KeyAscii = 13 Then
        If Not IsNumeric(txtBarNum2) Then
            StatusBar1.Panels(3).Text = "바코드번호는 숫자만 입력이 가능합니다."
            txtBarNum2 = ""
            Exit Sub
        End If
        
        If Len(txtBarNum2) <> 13 Then
            StatusBar1.Panels(3).Text = "바코드 자릿수를 확인하세요"
            txtBarNum2 = ""
            Exit Sub
        End If
        
        intRow = 0
        With vasID2
            For i = 1 To .MaxRows
                .Row = i
                .Col = colBARCODE
                If Trim(.Text) = Trim(txtBarNum2.Text) Then
                    intRow = i
                End If
            Next
            
            If intRow = 0 Then
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
            End If
            
            Call SetText(vasID2, txtBarNum2.Text, intRow, colBARCODE)
            DoEvents
            
        End With
        
        
        
        If Trim(txtBarNum2) <> "" Then
            'Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"), Trim(txtBarNum))
            Call GetSampleInfoW_PHILL(intRow, vasID2)
        End If
        vasID2.RowHeight(-1) = 12
        txtBarNum2.Text = ""
    End If
    
End Sub

Private Sub txtBarNum3_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim intRow  As Integer
    
    If KeyAscii = 13 Then
        If Not IsNumeric(txtBarNum3) Then
            StatusBar1.Panels(3).Text = "바코드번호는 숫자만 입력이 가능합니다."
            txtBarNum3 = ""
            Exit Sub
        End If
        
        If Len(txtBarNum3) <> 13 Then
            StatusBar1.Panels(3).Text = "바코드 자릿수를 확인하세요"
            txtBarNum3 = ""
            Exit Sub
        End If
        
        intRow = 0
        With vasID3
            For i = 1 To .MaxRows
                .Row = i
                .Col = colBARCODE
                If Trim(.Text) = Trim(txtBarNum.Text) Then
                    intRow = i
                End If
            Next
            
            If intRow = 0 Then
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
            End If
            
            Call SetText(vasID3, txtBarNum3.Text, intRow, colBARCODE)
            DoEvents
            
        End With
        
        
        
        If Trim(txtBarNum3) <> "" Then
            'Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"), Trim(txtBarNum))
            Call GetSampleInfoW_PHILL(intRow, vasID3)
        End If
        vasID3.RowHeight(-1) = 12
        txtBarNum3.Text = ""
    End If
    
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)

    If Winsock1.State <> sckClosed Then
        Winsock1.Close

        Winsock1.Accept requestID
        'MsgBox "장비에 접속되었습니다."
         
        StatusBar1.Panels(2).Text = "장비에 접속되었습니다"
    End If

End Sub

'Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'    Dim strText As String
'    Dim strTmp As String
'
'    Dim strLastSeq  As String
'    Dim strRcvSign  As String
'    Dim strSendAck  As String
'    Dim strRcvCnt   As String
'
'    Dim strNS       As String
'    Dim strNE       As String
'    Dim intNS       As Integer
'    Dim intNE       As Integer
'
'    Dim strSendData  As String
'    Dim varBuffers   As Variant
'    Dim i As Integer
'
'    Winsock1.GetData strText
'
'    strBuffer = strText
'
'    SetRawData "[Rx]" & strBuffer
'    StatusBar1.Panels(3).Text = strBuffer
'
'    varBuffers = Split(strBuffer, vbLf)
'
'    For i = 0 To UBound(varBuffers)
'        strBuffer = varBuffers(i)
'        If strBuffer = "" Then
'            Exit For
'        End If
'        strLastSeq = mGetP(strBuffer, 1, vbTab)
'        strRcvSign = mGetP(strBuffer, 2, vbTab)
'
'        strSendAck = strLastSeq & vbTab & "ACK"
'
'        Select Case UCase(strRcvSign)
'            Case "RESULT"
'                Call EditRcvDataASTM
'                strBuffer = ""
'
'            Case "CONNECT"
'                    Winsock1.SendData strSendAck & vbLf
'                    SetRawData "[Tx]" & strSendAck & vbLf
'
'            Case "RESULTS"
'                    strRcvCnt = CInt(mGetP(strBuffer, 3, vbTab))
'
'                    strNS = strRcvCnt
'                    strNE = CInt(mGetP(strBuffer, 4, vbTab))
'
'                    strNS = strNS - strNE
'                    strNE = strNS + strNE
'
'                    strSendData = strLastSeq & vbTab & "GET" & vbTab & strNS & vbTab & strNE
'                    'strSendData = "-" & strNS & vbTab & "GET" & vbTab & "0" & vbTab & "0"
'
'                    Winsock1.SendData strSendData & vbLf
'                    SetRawData "[Tx]" & strSendData & vbLf
'        End Select
'    Next
'
'
'End Sub

Private Sub SetExamCode()
    Dim i As Integer
    
    
    With vasID
        .MaxCols = colState + UBound(gArrEquip)
        
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeEditCharCase = TypeEditCharCaseSetNone
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            Call SetText(vasID, gArrEquip(i + 1, 4), 0, colState + (i + 1))
            .ColWidth(colState + (i + 1)) = gColWidth
            cboTest.AddItem gArrEquip(i + 1, 4) & Space(20) & "|" & gArrEquip(i + 1, 3)
        Next
        
        cboTest.ListIndex = 0
    End With
    
    With vasID2
        .MaxCols = colState + UBound(gArrEquip)
        
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeEditCharCase = TypeEditCharCaseSetNone
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            Call SetText(vasID2, gArrEquip(i + 1, 4), 0, colState + (i + 1))
            .ColWidth(colState + (i + 1)) = gColWidth
        Next
        
    End With
    
    With vasID3
        .MaxCols = colState + UBound(gArrEquip)
        
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeEditCharCase = TypeEditCharCaseSetNone
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            Call SetText(vasID3, gArrEquip(i + 1, 4), 0, colState + (i + 1))
            .ColWidth(colState + (i + 1)) = gColWidth
        Next
        
    End With
    
End Sub


Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    gAllExam = ""
    SQL = "Select equipcode, examcode, examname, resprec, seqno " & vbCrLf & _
          "  From EQPMASTER " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " Order by  seqno * 10 "
    Res = GetDBSelectVas(gLocal, SQL, vasCode)
    If Res > 0 Then
        ReDim gArrEquip(1 To vasCode.DataRowCnt, 1 To 7)
    Else
        SaveQuery SQL
        Exit Function
    End If
        
    For i = 1 To vasCode.DataRowCnt
        If i = 1 Then
            gAllExam = "'" & Trim(GetText(vasCode, i, 2)) & "'"
        Else
            gAllExam = gAllExam & ",'" & Trim(GetText(vasCode, i, 2)) & "'"
        End If
        
        gArrEquip(i, 1) = i
        For j = 1 To 6
            gArrEquip(i, j + 1) = Trim(GetText(vasCode, i, j))
        Next j
    Next i
    
    GetExamCode = 1
End Function

Private Sub Form_Unload(Cancel As Integer)
    If comEqp.PortOpen = True Then
        comEqp.PortOpen = False
    End If

'    Call dce_close_env      ' Server와 연결을 끊는 곳
    DisConnect_Server
    DisConnect_Local
    Unload Me
    End
End Sub

Private Sub MnExamConfig_Click()
    'frmTestSet.Show
    frmTestSet.Show
    GetExamCode
End Sub

Private Sub MnExit_Click()
    Unload Me
End Sub

Private Sub MnModeBarcode_Click()
    chkMode.Caption = "Barcode"
    MnModeBarcode.Checked = True
    MnModeWorkList.Checked = False
    chkBar.Value = 1
    
    gIFMode = "Barcode"
    Call WritePrivateProfileString("config", "IFMode", gIFMode, App.Path & "\Interface.ini")
 
End Sub

Private Sub MnModeWorkList_Click()
    chkMode.Caption = "WorkList"
    MnModeBarcode.Checked = False
    MnModeWorkList.Checked = True
    chkBar.Value = 0

    gIFMode = "WorkList"
    Call WritePrivateProfileString("config", "IFMode", gIFMode, App.Path & "\Interface.ini")

End Sub

Private Sub MnPrintLand_Click()

    vasID.PrintOrientation = PrintOrientationLandscape '가로출력
    vasID.Action = 13

End Sub

Private Sub MnPrintPort_Click()

    vasID.PrintOrientation = PrintOrientationPortrait '세로출력
    vasID.Action = 13

End Sub

'Private Sub MnScr1_Click()
'    MnScr1.Checked = True
'    MnScr2.Checked = False
'
'    gScreen = "분리"
'    Call WritePrivateProfileString("config", "IFScreen", gScreen, App.Path & "\Interface.ini")
'
'End Sub
'
'Private Sub MnScr2_Click()
'    MnScr1.Checked = False
'    MnScr2.Checked = True
'
'    gScreen = "통합"
'    Call WritePrivateProfileString("config", "IFScreen", gScreen, App.Path & "\Interface.ini")
'
'End Sub

Private Sub MnTConfig_Click()
    frmConfig.Show
End Sub

Private Sub MnTransAuto_Click()
    chkMode.Caption = "Auto"
    MnTransAuto.Checked = True
    MnTransManual.Checked = False
    chkMode.Value = 1

    gSave = "True"
    Call WritePrivateProfileString("config", "AutoSave", gSave, App.Path & "\Interface.ini")

End Sub

Private Sub MnTransManual_Click()
    chkMode.Caption = "Manual"
    MnTransAuto.Checked = False
    MnTransManual.Checked = True
    chkMode.Value = 0
    
    gSave = "False"
    Call WritePrivateProfileString("config", "AutoSave", gSave, App.Path & "\Interface.ini")

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput As String     '송신할 데이터
    
    '-- ASTM TYPE별 Define 해야함.
    '-- ASTM TYPE = Standard
    Select Case intSndPhase
        Case 1  '## Header
            'strOutput = intFrameNo & "H|\^&||| XN-10^00-14^15097^^^^AP795756||||||||E1394-97" & vbCr & ETX
            strOutput = intFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
        Case 2  '## Patient
            'strOutput = intFrameNo & "P|1||||^^|||U|||||^||||||||||||^^^" & vbCr & ETX
            strOutput = intFrameNo & "P|1" & vbCr & ETX
            
            intSndPhase = 4
            intFrameNo = intFrameNo + 1
            
        Case 3  '## No Order
            
        Case 4  '## Order
            If mOrder.NoOrder = True Then
                    
                strOutput = intFrameNo & "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q"
                intSndPhase = 5
            
            Else
                If mOrder.IsSending = False Then   '## 최초 보낼때
                    strOutput = "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q"
                    
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Left(strOutput, 230) & vbCr & ETB
                        intSndPhase = 4
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 5
                    End If
                Else                        '## 남은 문자열이 있을때
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Left(strOutput, 230) & vbCr & ETB
                        intSndPhase = 4
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 5
                    End If
                End If
                
            End If
            
            intFrameNo = intFrameNo + 1
            
        Case 5  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 6
            intFrameNo = intFrameNo + 1
            
        Case 6  '## EOT
            strState = ""
            comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1
            
            Exit Sub
    End Select
    
    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    comEqp.Output = strOutput
    Debug.Print strOutput
    SetRawData "[Tx]" & strOutput
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열의 CheckSum을 구함
'   인수 :
'       - pMsg : 문자열
'   반환 : CheckSum
'-----------------------------------------------------------------------------'
Public Function GetChkSum(ByVal pMsg As String) As String
    Dim lngChkSum   As Long
    Dim i           As Long

    For i = 1 To Len(pMsg)
        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
    Next

    If lngChkSum = 0 Then
        GetChkSum = "00"
    Else
        GetChkSum = Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2)
    End If
End Function

'-- 지금날짜와 검사일자 비교한다
Function DateCompare(ByVal FDate As String) As String
    
    DateCompare = FDate
    If FDate <> Format(Now, "yyyymmdd") Then
        DateCompare = Format(Now, "yyyymmdd")
    End If
    
End Function



Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub

Public Sub SndMore()
    Dim strSndMsg As String
    
    'Call Sleep(1000)
    
    strSndMsg = ">"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) ' & GetChkSum(strSndMsg) & vbCr
    comEqp.Output = strSndMsg & vbCrLf
    
    'SetRawData "[Tx]" & strSndMsg & vbCrLf
    Debug.Print "[SndMore]" & strSndMsg
    
End Sub

Public Sub SndRec()
    Dim strSndMsg As String
    
    strSndMsg = "A"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) '& GetChkSum(strSndMsg)
    comEqp.Output = strSndMsg & vbCrLf
    
End Sub

Private Sub comEqp_OnComm()
    Dim EVMsg       As String
    Dim ERMsg       As String
    Dim Ret         As Long
    Dim strDate     As String
    
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    
    On Error GoTo Rst
    
    Select Case comEqp.CommEvent
        Case comEvReceive

            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If

            Buffer = comEqp.Input
            
            SetRawData "[Rx]" & Buffer
            StatusBar1.Panels(3).Text = Buffer
            
            lngBufLen = Len(Buffer)
            
                For i = 1 To lngBufLen
                    BufChar = Mid$(Buffer, i, 1)
                    Select Case intPhase
                        Case 1
                            Select Case BufChar
                                Case "="
                                    strBuffer1 = ""
                                    intCrCnt = 0
                                    intPhase = 2
                                Case Else
                                    strBuffer1 = strBuffer1 & BufChar
                            End Select
                        Case 2
                            Select Case BufChar
                                Case vbCr '## CR 대기
                                    intCrCnt = intCrCnt + 1
                                    If intCrCnt >= 23 Then
                                        Call EditRcvDataASTM(vasID)
                                        strBuffer1 = ""
                                        intPhase = 1
                                        intCrCnt = 0
                                    End If
                                    strBuffer1 = strBuffer1 & BufChar
                                Case Else
                                    strBuffer1 = strBuffer1 & BufChar
                            
                            End Select
                    End Select
                Next i
            
        Case comEvSend
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
        
        Case comEvCTS
            EVMsg$ = "CTS 변경 감지"
        Case comEvDSR
            EVMsg$ = "DSR 변경 감지"
        Case comEvCD
            EVMsg$ = "CD 변경 감지"
        Case comEvRing
            EVMsg$ = "전화 벨이 울리는 중"
        Case comEvEOF
            EVMsg$ = "EOF 감지"

        '오류 메시지
        Case comBreak
            ERMsg$ = "중단 신호 수신"
        Case comCDTO
            ERMsg$ = "반송파 검출 시간 초과"
        Case comCTSTO
            ERMsg$ = "CTS 시간 초과"
        Case comDCB
            ERMsg$ = "DCB 검색 오류"
        Case comDSRTO
            ERMsg$ = "DSR 시간 초과"
        Case comFrame
            ERMsg$ = "프레이밍 오류"
        Case comOverrun
            ERMsg$ = "패리티 오류"
        Case comRxOver
            ERMsg$ = "수신 버퍼 초과"
        Case comRxParity
            ERMsg$ = "패리티 오류"
        Case comTxFull
            ERMsg$ = "전송 버퍼에 여유가 없음"
        Case Else
            ERMsg$ = "알 수 없는 오류 또는 이벤트"
    End Select

Exit Sub

Rst:
    StatusBar1.Panels(2).Text = Err.Description
    SetRawData "[Error_1]" & Err.Number & ":" & Err.Description

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 접수정보 조회, 표시, 검사오더만들기
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    For i = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, i, colBARCODE)) = pBarNo Then
            intRow = i
            Exit For
        End If
    Next i
    
    If intRow < 0 Then
        intRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If
    
    '-- 장비수신정보 표시
    Call SetText(vasID, pBarNo, intRow, colBARCODE)             '-- 바코드
    Call SetText(vasID, mOrder.RackNo, intRow, colDISKNO)       '-- Rack
    Call SetText(vasID, mOrder.TubePos, intRow, colPOSNO)       '-- Pos
    
    '-- 환자정보 표시
    Call vasActiveCell(vasID, intRow, colBARCODE)
    
    '-- 결과스프레드 지우기
    Call ClearSpread(vasRes)
    
    '-- 검사자 정보 가져오기
    Call GetSampleInfoW(intRow)
    
    '-- 바코드번호에 해당하는 검사코드 가져오기
    gOrderExam = GetOrderExamCode(gEquip, pBarNo)

    '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
    strItems = GetGetEquipExamCode_XN1000(gEquip, pBarNo, intRow)

    '-- 검사채널로 장비오더 만들기
    If Trim(strItems) = "" Then
        mOrder.NoOrder = True
        mOrder.Order = strItems
    Else
        mOrder.NoOrder = False
        mOrder.Order = strItems
    End If
    
    '-- 진행상태(Order) 표시
    Call SetText(vasID, "Order", intRow, colState)
    
    '-- 현재 Row
    gRow = intRow

End Sub

'-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기
Function GetGetEquipExamCode_XN1000(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    Dim strCBC As String
    Dim strDiff As String
    
    GetGetEquipExamCode_XN1000 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 샘플 바코드 번호
    SetRawData "[sBarcode]" & sBarcode
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    
    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""

    strCBC = ""
    strDiff = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            'NRBC%는 오더를 안준다
'            If Trim(gReadBuf(i)) <> "NRBC%" Then
'                strExamCode = strExamCode & "^^^^" & Trim(gReadBuf(i)) & "\"
'            End If
            
            
            If Trim(gReadBuf(i)) = "WBC" Or Trim(gReadBuf(i)) = "RBC" Or Trim(gReadBuf(i)) = "HGB" Or _
                Trim(gReadBuf(i)) = "HCT" Or Trim(gReadBuf(i)) = "MCV" Or Trim(gReadBuf(i)) = "MCH" Or Trim(gReadBuf(i)) = "MCHC" Or _
                Trim(gReadBuf(i)) = "PLT" Or Trim(gReadBuf(i)) = "RDW-SD" Or Trim(gReadBuf(i)) = "RDW-CV" Or Trim(gReadBuf(i)) = "PDW" Or _
                Trim(gReadBuf(i)) = "MPV" Or Trim(gReadBuf(i)) = "P-LCR" Or Trim(gReadBuf(i)) = "PCT" Or Trim(gReadBuf(i)) = "NRBC#" Or Trim(gReadBuf(i)) = "NRBC%" Then
                
                strCBC = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
                
            End If

            If Trim(gReadBuf(i)) = "NEUT#" Or Trim(gReadBuf(i)) = "LYMPH#" Or Trim(gReadBuf(i)) = "MONO#" Or Trim(gReadBuf(i)) = "EO#" Or Trim(gReadBuf(i)) = "BASO#" Or _
                Trim(gReadBuf(i)) = "NEUT%" Or Trim(gReadBuf(i)) = "LYMPH%" Or Trim(gReadBuf(i)) = "MONO%" Or Trim(gReadBuf(i)) = "EO%" Or Trim(gReadBuf(i)) = "BASO%" Or _
                Trim(gReadBuf(i)) = "IG#" Or Trim(gReadBuf(i)) = "IG%" Then
               
                '-- ^^^^LYMPH#\가 두개인 이유는 ETB 를 장비에서 인식하지 못하기 떄문..(그 자리가 230)
                strDiff = "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
                
            End If
        Else
            Exit For
        End If
    Next

    strExamCode = strCBC & strDiff
    
    '-- 오더가 없을 경우 CBC만 검사하도록 한다.
    If strExamCode = "" Then
        strExamCode = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
        strExamCode = strExamCode & "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
    End If
    
    If strExamCode <> "" Then
        strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    End If
    
    GetGetEquipExamCode_XN1000 = strExamCode
    
End Function

'-----------------------------------------------------------------------------'
'   기능 :
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub SetPatInfo(ByVal pBarNo As String, SPD As Object, Optional ByVal pRno As String, Optional ByVal pPno As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strTestDt   As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    For i = 1 To SPD.DataRowCnt
        If Trim(GetText(SPD, i, colBARCODE)) = pBarNo Then
            intRow = i
            Exit For
        End If
    Next i

    If intRow < 0 Then
        intRow = SPD.DataRowCnt + 1
        If SPD.MaxRows < intRow Then
            SPD.MaxRows = intRow
        End If
    End If
    
    
    '-- 장비수신정보 표시
    Call SetText(SPD, "1", intRow, colCheckBox)
    If pBarNo = "" Then
        Call SetText(SPD, mResult.SpcPos, intRow, colBARCODE)
    Else
        Call SetText(SPD, mResult.BarNo, intRow, colBARCODE)
    End If
    Call SetText(SPD, mResult.RackNo, intRow, colDISKNO)
    Call SetText(SPD, mResult.TubePos, intRow, colPOSNO)
    Call SetText(SPD, mResult.RsltDate, intRow, colEXAMDATE)
    Call SetText(SPD, mResult.RsltSeq, intRow, colSAVESEQ)
    
    
    Call vasActiveCell(SPD, intRow, colBARCODE)
    
    '-- 결과스프레드 지우기
    If SPD.Name = "vasID" Then
        Call ClearSpread(vasRes)
    End If
    If SPD.Name = "vasID2" Then
        Call ClearSpread(vasRes2)
    End If
    If SPD.Name = "vasID3" Then
        Call ClearSpread(vasRes3)
    End If
    '-- 검사자 정보 서버테이블에서 가져와 표시(for 워크리스트)  '6,7,8,9
    Call GetSampleInfoW_PHILL(intRow, SPD)
    
    '-- 현재 Row
    gRow = intRow
    
End Sub

Private Sub SendMessage_1200(ByVal MsgHead As String)
    On Error GoTo SendMessage_Error
    
    Dim chksum As Integer
    Dim Buffer As String
    Dim C As Integer
    Dim R As Integer
    Dim Tmp     As String
    Dim OrdVal  As String
    Dim OrdNm   As Variant

    Dim sSendData$
    
    Select Case MsgHead
        Case "ID_DATA"
            Buffer = STX & "ID_DATA" & FS & RS _
                                    & "aMOD" & GS & "LIS" & GS & GS & GS & FS _
                                    & "iIID" & GS & "333" & GS & GS & GS & FS & RS _
                                    & ETX
        Case "SMP_REQ"
            Buffer = STX & "SMP_REQ" & FS & RS & "aMOD" & GS & aMod & GS & GS & GS _
                                        & FS & "iIID" & GS & iIID & GS & GS & GS _
                                        & FS & "rSEQ" & GS & Sample_Seq & GS & GS & GS _
                                        & FS & RS & ETX
            
        Case "SMP_ORD"
    End Select
        
    For C = 1 To Len(Buffer)
        chksum = chksum + Asc(Mid(Buffer, C, 1))
    Next C
    
    sSendData = Buffer & Right("0" & Hex(chksum Mod 256), 2) & EOT
    
    comEqp.Output = sSendData
    
SendMessage_Error:
    If Err <> 0 Then
'        RaiseEvent DispMsg("SendMessage Error : " & Err.Description)
    End If
End Sub

Private Sub GetaModiIID(ByVal sMsg As String)

    Dim tmpData()   As String
    
    '<STX>SYS_READY<FS><RS>aMOD<GS>1265<GS><GS><GS><FS>iIID
    '<GS>12345<GS><GS><GS><FS>aDATE<GS>20Jan2004<GS><GS><GS>
    '<FS>aTIME<GS>13:35:32<GS><GS><GS><FS>iOID<GS>3<GS><GS><GS><FS>
    '<ETX>{chksum}<EOT>

    tmpData() = Split(sMsg, GS)
    
    'aMod
    aMod = Trim(tmpData(1))
    
    'iIID
    iIID = Trim(tmpData(5))

End Sub


Private Function ConvertDateType(ByVal sDate As String) As String
    On Error GoTo ErrRtn
    
    Dim kk%
    Dim sTmp$
    Dim tmpYYYY$, tmpMM$, tmpDD$
    
    ConvertDateType = sDate
    
    tmpYYYY = Right(sDate, 4)
    sDate = Mid(sDate, 1, Len(sDate) - 4)
    
    For kk = 1 To Len(sDate)
        sTmp = Mid(sDate, kk, 1)
        If IsNumeric(sTmp) Then
            tmpDD = tmpDD & sTmp
        Else
            tmpMM = tmpMM & sTmp
        End If
    Next kk
    
    sTmp = tmpDD & Space(1) & tmpMM & Space(1) & tmpYYYY
    
    ConvertDateType = Format(sTmp, "YYYYMMDD")
    
ErrRtn:
    If Err <> 0 Then
        'RaiseEvent DispMsg("ConvertDateType - " & Err.Description)
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataASTM(SPD As Object)
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    Dim strIntResult As String   '수신한 결과(정량)
    Dim strQCResult  As String   '수신한 결과(QC)
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim intCnt       As Integer
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    Dim strTmp      As String
    Dim intIDX      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
    Dim Pos As Integer
    Dim strSeqNo As String
    Dim varORQN As Variant
    Dim strHoleNo    As String
    Dim strFlags        As String
    
    Dim strRcvMsg2      As String
    Dim strRcvMsg3      As String
    Dim strRcvMsg7      As String
    
    
    Dim X   As Integer
    Dim C   As Integer
    Dim MsgID   As String
    Dim RcvBuffer As String
    
    Dim R   As Integer
    Dim x1  As Integer
    Dim x2  As Integer
    Dim AssayNm As String
    Dim RESULT  As String
    Dim EqCd    As String
    Dim OrdCd   As String
    Dim LabNo   As String
    Dim rSeq    As String
    Dim iPID    As String

    Dim sRstDate$, sRstTime$
    Dim MsgBuf$
    
    Dim strCmnt As String
    
On Error GoTo Rst


    If strBuffer1 = "" Then
        Exit Sub
    End If
    
    strBuffer1 = Replace(strBuffer1, vbLf, "")
    varRcvBuf = Split(strBuffer1, vbCr)
    
    If UBound(varRcvBuf) <= 0 Then
        strBuffer1 = ""
        Exit Sub
    End If
    
    If Trim(varRcvBuf(2)) = "SAMPLE RESULTS" Then
        For intRow = 7 To UBound(varRcvBuf)
            strRcvBuf = varRcvBuf(intRow)
            
            Select Case intRow
                Case 7
                    strSeq = mGetP(strRcvBuf, 2, "#")
                    strSeq = Val(strSeq)
                    gRow = 0
                    For i = 1 To SPD.DataRowCnt
                        SPD.Row = i
                        SPD.Col = colCheckBox
                        If SPD.Value = "1" Then
                            gRow = i
                            strBarNo = GetText(SPD, i, colBARCODE)
                            Exit For
                        End If
                    Next i
                        
                    If strBarNo = "" Then
                        SPD.MaxRows = SPD.MaxRows + 1
                        gRow = SPD.MaxRows
                        strBarNo = strSeq
                        'Exit Sub
                    End If
                    
                    With mResult
                        .BarNo = strBarNo
                        '.RackNo = strRackNo
                        '.TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                        .SpcPos = strSeqNo
                    End With
                                        
                    Call SetPatInfo(strBarNo, SPD)
                
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                
                Case 10 To 12
                    strIntBase = Mid(strRcvBuf, 1, 4)
                    strIntBase = Trim(strIntBase)
                    
                    strResult = Mid(strRcvBuf, 5, 6)
                    strResult = Trim(strResult)
                        
                    '절사
                    If UCase(Mid(strIntBase, 1, 2)) = "NA" Or UCase(Mid(strIntBase, 1, 2)) = "CL" Then
                        strResult = mGetP(strResult, 1, ".") '& "." & Mid(mGetP(strResult, 2, "."), 1, 1)
                        strResult = Replace(strResult, ">", "")
                    End If
                    If UCase(Mid(strIntBase, 1, 1)) = "K" Then
                        strResult = mGetP(strResult, 1, ".") & "." & Mid(mGetP(strResult, 2, "."), 1, 1)
                        strResult = Replace(strResult, ">", "")
                    End If
                    
                    If UCase(Mid(strIntBase, 1, 1)) = "K" Then
                        If IsNumeric(strResult) Then
                            'K가   > 8.0  ==> 8.0
                            If CCur(strResult) > 8 Then
                                strResult = "8.0"
                            End If
                        End If
                    End If
                    
                    If strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                        SQL = SQL & "  FROM EQPMASTER"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                        
                        Res = GetDBSelectColumn(gLocal, SQL)
                        
                        '-- 오더 있을 경우
                        If Res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            
                            If SPD.Name = "vasID" Then
                                lsResRow = vasRes.DataRowCnt + 1
                                If vasRes.MaxRows < lsResRow Then
                                    vasRes.MaxRows = lsResRow
                                End If
                            End If
                            
                            If SPD.Name = "vasID2" Then
                                lsResRow = vasRes2.DataRowCnt + 1
                                If vasRes2.MaxRows < lsResRow Then
                                    vasRes2.MaxRows = lsResRow
                                End If
                            End If
                            
                            If SPD.Name = "vasID3" Then
                                lsResRow = vasRes3.DataRowCnt + 1
                                If vasRes3.MaxRows < lsResRow Then
                                    vasRes3.MaxRows = lsResRow
                                End If
                            End If
                            
                            
                            '소수점 처리, 결과 형태 처리
                            lsEquipRes = strResult
                            strResult = SetResult(strResult, strIntBase)
                            lsResult_Buff = strResult
                            
                            '-- Work List
                            SetText SPD, "Result", gRow, colState                 '11 진행상태
                            
                            '-- 결과저장용 seq
                            For intCol = colState + 1 To SPD.MaxCols
                                If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                    SetText SPD, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
                
                            '-- 결과 List
'                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
'                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
'                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
'                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
'                            SetText vasRes, strResult, lsResRow, colRESULT          '결과
'                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
'                            SetText vasRes, strComm, lsResRow, 7                    'Flag
'                            '-- 로컬 저장
'                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                        
                            If SPD.Name = "vasID" Then
                                '-- 결과 List
                                SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                                SetText vasRes, strResult, lsResRow, colRESULT          '결과
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            
                                '-- 로컬 저장
                                Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes, lsEquipRes)
                            End If
                            
                            If SPD.Name = "vasID2" Then
                                '-- 결과 List
                                SetText vasRes2, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                SetText vasRes2, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                SetText vasRes2, lsExamName, lsResRow, colEXAMNAME       '검사명
                                SetText vasRes2, lsEquipRes, lsResRow, colMachResult     '장비결과
                                SetText vasRes2, strResult, lsResRow, colRESULT          '결과
                                SetText vasRes2, lsSeqNo, lsResRow, colSeq               '순번
                                SetText vasRes2, strComm, lsResRow, colFLAG              'Flag
                                
                                '-- 로컬 저장
                                Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes2, lsEquipRes)
                            End If
                            
                            If SPD.Name = "vasID3" Then
                                '-- 결과 List
                                SetText vasRes3, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                SetText vasRes3, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                SetText vasRes3, lsExamName, lsResRow, colEXAMNAME       '검사명
                                SetText vasRes3, lsEquipRes, lsResRow, colMachResult     '장비결과
                                SetText vasRes3, strResult, lsResRow, colRESULT          '결과
                                SetText vasRes3, lsSeqNo, lsResRow, colSeq               '순번
                                SetText vasRes3, strComm, lsResRow, colFLAG              'Flag
                                
                                '-- 로컬 저장
                                Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes3, lsEquipRes)
                            End If
                            
                            lsResult_Buff = ""
                            
                            strState = "R"
                            
                        '-- 오더 없을 경우
                        Else
                        
                                  SQL = "Select examcode, examname, seqno "
                            SQL = SQL & "  From EQPMASTER"
                            SQL = SQL & " Where equipno = '" & gEquip & "' "
                            SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                            Res = GetDBSelectColumn(gLocal, SQL)
                                                    
                            If Res > 0 Then
                                lsExamCode = Trim(gReadBuf(0))
                                lsExamName = Trim(gReadBuf(1))
                                lsSeqNo = Trim(gReadBuf(2))
                                
'                                lsResRow = vasRes.DataRowCnt + 1
'                                If vasRes.MaxRows < lsResRow Then
'                                    vasRes.MaxRows = lsResRow
'                                End If
                                
                                If SPD.Name = "vasID" Then
                                    lsResRow = vasRes.DataRowCnt + 1
                                    If vasRes.MaxRows < lsResRow Then
                                        vasRes.MaxRows = lsResRow
                                    End If
                                End If
                                
                                If SPD.Name = "vasID2" Then
                                    lsResRow = vasRes2.DataRowCnt + 1
                                    If vasRes2.MaxRows < lsResRow Then
                                        vasRes2.MaxRows = lsResRow
                                    End If
                                End If
                                
                                If SPD.Name = "vasID3" Then
                                    lsResRow = vasRes3.DataRowCnt + 1
                                    If vasRes3.MaxRows < lsResRow Then
                                        vasRes3.MaxRows = lsResRow
                                    End If
                                End If
                                
                                '소수점 처리, 결과 형태 처리
                                lsEquipRes = strResult
                                strResult = SetResult(strResult, strIntBase)
                                lsResult_Buff = strResult
                                
                                '-- Work List
                                SetText SPD, "Result", gRow, colState                 '진행상태
                                
                                '-- 결과저장용 seq
                                For intCol = colState + 1 To SPD.MaxCols
                                    If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                        SetText SPD, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
                                
                                If SPD.Name = "vasID" Then
                                    '-- 결과 List
                                    SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                    SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                    SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                                    SetText vasRes, strResult, lsResRow, colRESULT          '결과
                                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                                    SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                
                                    '-- 로컬 저장
                                    Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes, lsEquipRes)
                                End If
                                
                                If SPD.Name = "vasID2" Then
                                    '-- 결과 List
                                    SetText vasRes2, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                    SetText vasRes2, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                    SetText vasRes2, lsExamName, lsResRow, colEXAMNAME       '검사명
                                    SetText vasRes2, lsEquipRes, lsResRow, colMachResult     '장비결과
                                    SetText vasRes2, strResult, lsResRow, colRESULT          '결과
                                    SetText vasRes2, lsSeqNo, lsResRow, colSeq               '순번
                                    SetText vasRes2, strComm, lsResRow, colFLAG              'Flag
                                    
                                    '-- 로컬 저장
                                    Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes2, lsEquipRes)
                                End If
                                
                                If SPD.Name = "vasID3" Then
                                    '-- 결과 List
                                    SetText vasRes3, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                    SetText vasRes3, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                    SetText vasRes3, lsExamName, lsResRow, colEXAMNAME       '검사명
                                    SetText vasRes3, lsEquipRes, lsResRow, colMachResult     '장비결과
                                    SetText vasRes3, strResult, lsResRow, colRESULT          '결과
                                    SetText vasRes3, lsSeqNo, lsResRow, colSeq               '순번
                                    SetText vasRes3, strComm, lsResRow, colFLAG              'Flag
                                    
                                    '-- 로컬 저장
                                    Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes3, lsEquipRes)
                                End If
                                
                                
                                
                                lsResult_Buff = ""
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                    End If
        
                Case 14
                    '## DB에 결과저장
                    If MnTransAuto.Checked = True And strState = "R" Then
                        Res = SaveTransDataW(gRow, SPD)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetText SPD, "0", gRow, colCheckBox
                            SetForeColor SPD, gRow, gRow, 1, colState, 255, 0, 0
                            SetText SPD, "Failed", gRow, colState
                        Else
                            '-- 저장 성공
                            SetBackColor SPD, gRow, gRow, 1, colState, 202, 255, 112
                            SetText SPD, "Trans", gRow, colState
                            SetText SPD, "0", gRow, colCheckBox
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gEquip & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(SPD, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(SPD, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(SPD, gRow, colSAVESEQ)) & vbCrLf
                            
                            Res = SendQuery(gLocal, SQL)
                            If Res = -1 Then
                                SaveQuery SQL
                                Exit Sub
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End If
    
Exit Sub

Rst:
    strBuffer1 = ""
    
    StatusBar1.Panels(2).Text = Err.Description
    SetRawData "[EditRcvDataASTM]" & Err.Number & ":" & Err.Description
    
End Sub


'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataASTM2(SPD As Object)
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    Dim strIntResult As String   '수신한 결과(정량)
    Dim strQCResult  As String   '수신한 결과(QC)
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim intCnt       As Integer
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    Dim strTmp      As String
    Dim intIDX      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
    Dim Pos As Integer
    Dim strSeqNo As String
    Dim varORQN As Variant
    Dim strHoleNo    As String
    Dim strFlags        As String
    
    Dim strRcvMsg2      As String
    Dim strRcvMsg3      As String
    Dim strRcvMsg7      As String
    
    
    Dim X   As Integer
    Dim C   As Integer
    Dim MsgID   As String
    Dim RcvBuffer As String
    
    Dim R   As Integer
    Dim x1  As Integer
    Dim x2  As Integer
    Dim AssayNm As String
    Dim RESULT  As String
    Dim EqCd    As String
    Dim OrdCd   As String
    Dim LabNo   As String
    Dim rSeq    As String
    Dim iPID    As String

    Dim sRstDate$, sRstTime$
    Dim MsgBuf$
    
    Dim strCmnt As String
    
On Error GoTo Rst


    If strBuffer2 = "" Then
        Exit Sub
    End If
    
    strBuffer2 = Replace(strBuffer2, vbLf, "")
    varRcvBuf = Split(strBuffer2, vbCr)
    
    If UBound(varRcvBuf) <= 0 Then
        strBuffer2 = ""
        Exit Sub
    End If
    
    If Trim(varRcvBuf(2)) = "SAMPLE RESULTS" Then
        For intRow = 7 To UBound(varRcvBuf)
            strRcvBuf = varRcvBuf(intRow)
            
            Select Case intRow
                Case 7
                    strSeq = mGetP(strRcvBuf, 2, "#")
                    strSeq = Val(strSeq)
                    gRow = 0
                    For i = 1 To SPD.DataRowCnt
                        SPD.Row = i
                        SPD.Col = colCheckBox
                        If SPD.Value = "1" Then
                            gRow = i
                            strBarNo = GetText(SPD, i, colBARCODE)
                            Exit For
                        End If
                    Next i
                        
                    If strBarNo = "" Then
                        SPD.MaxRows = SPD.MaxRows + 1
                        gRow = SPD.MaxRows
                        strBarNo = strSeq
                        'Exit Sub
                    End If
                    
                    With mResult
                        .BarNo = strBarNo
                        '.RackNo = strRackNo
                        '.TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                        .SpcPos = strSeqNo
                    End With
                                        
                    Call SetPatInfo(strBarNo, SPD)
                
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                
                Case 10 To 12
                    strIntBase = Mid(strRcvBuf, 1, 4)
                    strIntBase = Trim(strIntBase)
                    
                    strResult = Mid(strRcvBuf, 5, 6)
                    strResult = Trim(strResult)
                
                    '절사
                    If UCase(Mid(strIntBase, 1, 2)) = "NA" Or UCase(Mid(strIntBase, 1, 2)) = "CL" Then
                        strResult = mGetP(strResult, 1, ".") '& "." & Mid(mGetP(strResult, 2, "."), 1, 1)
                        strResult = Replace(strResult, ">", "")
                    End If
                    If UCase(Mid(strIntBase, 1, 1)) = "K" Then
                        strResult = mGetP(strResult, 1, ".") & "." & Mid(mGetP(strResult, 2, "."), 1, 1)
                        strResult = Replace(strResult, ">", "")
                    End If
                
                
                    If UCase(Mid(strIntBase, 1, 1)) = "K" Then
                        strResult = Replace(strResult, ">", "")
                        
                        If IsNumeric(strResult) Then
                            'K가   > 8.0  ==> 8.0
                            If CCur(strResult) > 8 Then
                                strResult = "8.0"
                            End If
                        End If
                    End If
                    
                    If strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                        SQL = SQL & "  FROM EQPMASTER"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                        
                        Res = GetDBSelectColumn(gLocal, SQL)
                        
                        '-- 오더 있을 경우
                        If Res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            
                            If SPD.Name = "vasID" Then
                                lsResRow = vasRes.DataRowCnt + 1
                                If vasRes.MaxRows < lsResRow Then
                                    vasRes.MaxRows = lsResRow
                                End If
                            End If
                            
                            If SPD.Name = "vasID2" Then
                                lsResRow = vasRes2.DataRowCnt + 1
                                If vasRes2.MaxRows < lsResRow Then
                                    vasRes2.MaxRows = lsResRow
                                End If
                            End If
                            
                            If SPD.Name = "vasID3" Then
                                lsResRow = vasRes3.DataRowCnt + 1
                                If vasRes3.MaxRows < lsResRow Then
                                    vasRes3.MaxRows = lsResRow
                                End If
                            End If
                            
                            
                            '소수점 처리, 결과 형태 처리
                            lsEquipRes = strResult
                            strResult = SetResult(strResult, strIntBase)
                            lsResult_Buff = strResult
                            
                            '-- Work List
                            SetText SPD, "Result", gRow, colState                 '11 진행상태
                            
                            '-- 결과저장용 seq
                            For intCol = colState + 1 To SPD.MaxCols
                                If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                    SetText SPD, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
                
                            '-- 결과 List
'                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
'                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
'                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
'                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
'                            SetText vasRes, strResult, lsResRow, colRESULT          '결과
'                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
'                            SetText vasRes, strComm, lsResRow, 7                    'Flag
'                            '-- 로컬 저장
'                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                        
                            If SPD.Name = "vasID" Then
                                '-- 결과 List
                                SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                                SetText vasRes, strResult, lsResRow, colRESULT          '결과
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            
                                '-- 로컬 저장
                                Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes, lsEquipRes)
                            End If
                            
                            If SPD.Name = "vasID2" Then
                                '-- 결과 List
                                SetText vasRes2, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                SetText vasRes2, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                SetText vasRes2, lsExamName, lsResRow, colEXAMNAME       '검사명
                                SetText vasRes2, lsEquipRes, lsResRow, colMachResult     '장비결과
                                SetText vasRes2, strResult, lsResRow, colRESULT          '결과
                                SetText vasRes2, lsSeqNo, lsResRow, colSeq               '순번
                                SetText vasRes2, strComm, lsResRow, colFLAG              'Flag
                                
                                '-- 로컬 저장
                                Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes2, lsEquipRes)
                            End If
                            
                            If SPD.Name = "vasID3" Then
                                '-- 결과 List
                                SetText vasRes3, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                SetText vasRes3, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                SetText vasRes3, lsExamName, lsResRow, colEXAMNAME       '검사명
                                SetText vasRes3, lsEquipRes, lsResRow, colMachResult     '장비결과
                                SetText vasRes3, strResult, lsResRow, colRESULT          '결과
                                SetText vasRes3, lsSeqNo, lsResRow, colSeq               '순번
                                SetText vasRes3, strComm, lsResRow, colFLAG              'Flag
                                
                                '-- 로컬 저장
                                Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes3, lsEquipRes)
                            End If
                            
                            lsResult_Buff = ""
                            
                            strState = "R"
                            
                        '-- 오더 없을 경우
                        Else
                        
                                  SQL = "Select examcode, examname, seqno "
                            SQL = SQL & "  From EQPMASTER"
                            SQL = SQL & " Where equipno = '" & gEquip & "' "
                            SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                            Res = GetDBSelectColumn(gLocal, SQL)
                                                    
                            If Res > 0 Then
                                lsExamCode = Trim(gReadBuf(0))
                                lsExamName = Trim(gReadBuf(1))
                                lsSeqNo = Trim(gReadBuf(2))
                                
'                                lsResRow = vasRes.DataRowCnt + 1
'                                If vasRes.MaxRows < lsResRow Then
'                                    vasRes.MaxRows = lsResRow
'                                End If
                                
                                If SPD.Name = "vasID" Then
                                    lsResRow = vasRes.DataRowCnt + 1
                                    If vasRes.MaxRows < lsResRow Then
                                        vasRes.MaxRows = lsResRow
                                    End If
                                End If
                                
                                If SPD.Name = "vasID2" Then
                                    lsResRow = vasRes2.DataRowCnt + 1
                                    If vasRes2.MaxRows < lsResRow Then
                                        vasRes2.MaxRows = lsResRow
                                    End If
                                End If
                                
                                If SPD.Name = "vasID3" Then
                                    lsResRow = vasRes3.DataRowCnt + 1
                                    If vasRes3.MaxRows < lsResRow Then
                                        vasRes3.MaxRows = lsResRow
                                    End If
                                End If
                                
                                '소수점 처리, 결과 형태 처리
                                lsEquipRes = strResult
                                strResult = SetResult(strResult, strIntBase)
                                lsResult_Buff = strResult
                                
                                '-- Work List
                                SetText SPD, "Result", gRow, colState                 '진행상태
                                
                                '-- 결과저장용 seq
                                For intCol = colState + 1 To SPD.MaxCols
                                    If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                        SetText SPD, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
                                
                                If SPD.Name = "vasID" Then
                                    '-- 결과 List
                                    SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                    SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                    SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                                    SetText vasRes, strResult, lsResRow, colRESULT          '결과
                                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                                    SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                
                                    '-- 로컬 저장
                                    Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes, lsEquipRes)
                                End If
                                
                                If SPD.Name = "vasID2" Then
                                    '-- 결과 List
                                    SetText vasRes2, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                    SetText vasRes2, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                    SetText vasRes2, lsExamName, lsResRow, colEXAMNAME       '검사명
                                    SetText vasRes2, lsEquipRes, lsResRow, colMachResult     '장비결과
                                    SetText vasRes2, strResult, lsResRow, colRESULT          '결과
                                    SetText vasRes2, lsSeqNo, lsResRow, colSeq               '순번
                                    SetText vasRes2, strComm, lsResRow, colFLAG              'Flag
                                    
                                    '-- 로컬 저장
                                    Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes2, lsEquipRes)
                                End If
                                
                                If SPD.Name = "vasID3" Then
                                    '-- 결과 List
                                    SetText vasRes3, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                    SetText vasRes3, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                    SetText vasRes3, lsExamName, lsResRow, colEXAMNAME       '검사명
                                    SetText vasRes3, lsEquipRes, lsResRow, colMachResult     '장비결과
                                    SetText vasRes3, strResult, lsResRow, colRESULT          '결과
                                    SetText vasRes3, lsSeqNo, lsResRow, colSeq               '순번
                                    SetText vasRes3, strComm, lsResRow, colFLAG              'Flag
                                    
                                    '-- 로컬 저장
                                    Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes3, lsEquipRes)
                                End If
                                
                                
                                
                                lsResult_Buff = ""
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                    End If
        
                Case 14
                    '## DB에 결과저장
                    If MnTransAuto.Checked = True And strState = "R" Then
                        Res = SaveTransDataW(gRow, SPD)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetText SPD, "0", gRow, colCheckBox
                            SetForeColor SPD, gRow, gRow, 1, colState, 255, 0, 0
                            SetText SPD, "Failed", gRow, colState
                        Else
                            '-- 저장 성공
                            SetBackColor SPD, gRow, gRow, 1, colState, 202, 255, 112
                            SetText SPD, "Trans", gRow, colState
                            SetText SPD, "0", gRow, colCheckBox
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gEquip & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(SPD, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(SPD, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(SPD, gRow, colSAVESEQ)) & vbCrLf
                            
                            Res = SendQuery(gLocal, SQL)
                            If Res = -1 Then
                                SaveQuery SQL
                                Exit Sub
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End If
    
Exit Sub

Rst:
    strBuffer2 = ""
    
    StatusBar1.Panels(2).Text = Err.Description
    SetRawData "[EditRcvDataASTM]" & Err.Number & ":" & Err.Description
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataASTM3(SPD As Object)
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    Dim strIntResult As String   '수신한 결과(정량)
    Dim strQCResult  As String   '수신한 결과(QC)
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim intCnt       As Integer
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    Dim strTmp      As String
    Dim intIDX      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
    Dim Pos As Integer
    Dim strSeqNo As String
    Dim varORQN As Variant
    Dim strHoleNo    As String
    Dim strFlags        As String
    
    Dim strRcvMsg2      As String
    Dim strRcvMsg3      As String
    Dim strRcvMsg7      As String
    
    
    Dim X   As Integer
    Dim C   As Integer
    Dim MsgID   As String
    Dim RcvBuffer As String
    
    Dim R   As Integer
    Dim x1  As Integer
    Dim x2  As Integer
    Dim AssayNm As String
    Dim RESULT  As String
    Dim EqCd    As String
    Dim OrdCd   As String
    Dim LabNo   As String
    Dim rSeq    As String
    Dim iPID    As String

    Dim sRstDate$, sRstTime$
    Dim MsgBuf$
    
    Dim strCmnt As String
    
On Error GoTo Rst


    If strBuffer3 = "" Then
        Exit Sub
    End If
    
    strBuffer3 = Replace(strBuffer3, vbLf, "")
    varRcvBuf = Split(strBuffer3, vbCr)
    
    If UBound(varRcvBuf) <= 0 Then
        strBuffer3 = ""
        Exit Sub
    End If
    
    If Trim(varRcvBuf(2)) = "SAMPLE RESULTS" Then
        For intRow = 7 To UBound(varRcvBuf)
            strRcvBuf = varRcvBuf(intRow)
            
            Select Case intRow
                Case 7
                    strSeq = mGetP(strRcvBuf, 2, "#")
                    strSeq = Val(strSeq)
                    gRow = 0
                    For i = 1 To SPD.DataRowCnt
                        SPD.Row = i
                        SPD.Col = colCheckBox
                        If SPD.Value = "1" Then
                            gRow = i
                            strBarNo = GetText(SPD, i, colBARCODE)
                            Exit For
                        End If
                    Next i
                        
                    If strBarNo = "" Then
                        SPD.MaxRows = SPD.MaxRows + 1
                        gRow = SPD.MaxRows
                        strBarNo = strSeq
                        'Exit Sub
                    End If
                    
                    With mResult
                        .BarNo = strBarNo
                        '.RackNo = strRackNo
                        '.TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                        .SpcPos = strSeqNo
                    End With
                                        
                    Call SetPatInfo(strBarNo, SPD)
                
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                
                Case 10 To 12
                    strIntBase = Mid(strRcvBuf, 1, 4)
                    strIntBase = Trim(strIntBase)
                    
                    strResult = Mid(strRcvBuf, 5, 6)
                    strResult = Trim(strResult)
                
                    '절사
                    If UCase(Mid(strIntBase, 1, 2)) = "NA" Or UCase(Mid(strIntBase, 1, 2)) = "CL" Then
                        strResult = mGetP(strResult, 1, ".") '& "." & Mid(mGetP(strResult, 2, "."), 1, 1)
                        strResult = Replace(strResult, ">", "")
                    End If
                    If UCase(Mid(strIntBase, 1, 1)) = "K" Then
                        strResult = mGetP(strResult, 1, ".") & "." & Mid(mGetP(strResult, 2, "."), 1, 1)
                        strResult = Replace(strResult, ">", "")
                    End If
                
                    If UCase(Mid(strIntBase, 1, 1)) = "K" Then
                        strResult = Replace(strResult, ">", "")
                        
                        If IsNumeric(strResult) Then
                            'K가   > 8.0  ==> 8.0
                            If CCur(strResult) > 8 Then
                                strResult = "8.0"
                            End If
                        End If
                    End If
                    
                    If strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                        SQL = SQL & "  FROM EQPMASTER"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                        
                        Res = GetDBSelectColumn(gLocal, SQL)
                        
                        '-- 오더 있을 경우
                        If Res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            
                            If SPD.Name = "vasID" Then
                                lsResRow = vasRes.DataRowCnt + 1
                                If vasRes.MaxRows < lsResRow Then
                                    vasRes.MaxRows = lsResRow
                                End If
                            End If
                            
                            If SPD.Name = "vasID2" Then
                                lsResRow = vasRes2.DataRowCnt + 1
                                If vasRes2.MaxRows < lsResRow Then
                                    vasRes2.MaxRows = lsResRow
                                End If
                            End If
                            
                            If SPD.Name = "vasID3" Then
                                lsResRow = vasRes3.DataRowCnt + 1
                                If vasRes3.MaxRows < lsResRow Then
                                    vasRes3.MaxRows = lsResRow
                                End If
                            End If
                            
                            
                            '소수점 처리, 결과 형태 처리
                            lsEquipRes = strResult
                            strResult = SetResult(strResult, strIntBase)
                            lsResult_Buff = strResult
                            
                            '-- Work List
                            SetText SPD, "Result", gRow, colState                 '11 진행상태
                            
                            '-- 결과저장용 seq
                            For intCol = colState + 1 To SPD.MaxCols
                                If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                    SetText SPD, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
                
                            '-- 결과 List
'                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
'                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
'                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
'                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
'                            SetText vasRes, strResult, lsResRow, colRESULT          '결과
'                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
'                            SetText vasRes, strComm, lsResRow, 7                    'Flag
'                            '-- 로컬 저장
'                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                        
                            If SPD.Name = "vasID" Then
                                '-- 결과 List
                                SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                                SetText vasRes, strResult, lsResRow, colRESULT          '결과
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            
                                '-- 로컬 저장
                                Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes, lsEquipRes)
                            End If
                            
                            If SPD.Name = "vasID2" Then
                                '-- 결과 List
                                SetText vasRes2, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                SetText vasRes2, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                SetText vasRes2, lsExamName, lsResRow, colEXAMNAME       '검사명
                                SetText vasRes2, lsEquipRes, lsResRow, colMachResult     '장비결과
                                SetText vasRes2, strResult, lsResRow, colRESULT          '결과
                                SetText vasRes2, lsSeqNo, lsResRow, colSeq               '순번
                                SetText vasRes2, strComm, lsResRow, colFLAG              'Flag
                                
                                '-- 로컬 저장
                                Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes2, lsEquipRes)
                            End If
                            
                            If SPD.Name = "vasID3" Then
                                '-- 결과 List
                                SetText vasRes3, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                SetText vasRes3, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                SetText vasRes3, lsExamName, lsResRow, colEXAMNAME       '검사명
                                SetText vasRes3, lsEquipRes, lsResRow, colMachResult     '장비결과
                                SetText vasRes3, strResult, lsResRow, colRESULT          '결과
                                SetText vasRes3, lsSeqNo, lsResRow, colSeq               '순번
                                SetText vasRes3, strComm, lsResRow, colFLAG              'Flag
                                
                                '-- 로컬 저장
                                Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes3, lsEquipRes)
                            End If
                            
                            lsResult_Buff = ""
                            
                            strState = "R"
                            
                        '-- 오더 없을 경우
                        Else
                        
                                  SQL = "Select examcode, examname, seqno "
                            SQL = SQL & "  From EQPMASTER"
                            SQL = SQL & " Where equipno = '" & gEquip & "' "
                            SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                            Res = GetDBSelectColumn(gLocal, SQL)
                                                    
                            If Res > 0 Then
                                lsExamCode = Trim(gReadBuf(0))
                                lsExamName = Trim(gReadBuf(1))
                                lsSeqNo = Trim(gReadBuf(2))
                                
'                                lsResRow = vasRes.DataRowCnt + 1
'                                If vasRes.MaxRows < lsResRow Then
'                                    vasRes.MaxRows = lsResRow
'                                End If
                                
                                If SPD.Name = "vasID" Then
                                    lsResRow = vasRes.DataRowCnt + 1
                                    If vasRes.MaxRows < lsResRow Then
                                        vasRes.MaxRows = lsResRow
                                    End If
                                End If
                                
                                If SPD.Name = "vasID2" Then
                                    lsResRow = vasRes2.DataRowCnt + 1
                                    If vasRes2.MaxRows < lsResRow Then
                                        vasRes2.MaxRows = lsResRow
                                    End If
                                End If
                                
                                If SPD.Name = "vasID3" Then
                                    lsResRow = vasRes3.DataRowCnt + 1
                                    If vasRes3.MaxRows < lsResRow Then
                                        vasRes3.MaxRows = lsResRow
                                    End If
                                End If
                                
                                '소수점 처리, 결과 형태 처리
                                lsEquipRes = strResult
                                strResult = SetResult(strResult, strIntBase)
                                lsResult_Buff = strResult
                                
                                '-- Work List
                                SetText SPD, "Result", gRow, colState                 '진행상태
                                
                                '-- 결과저장용 seq
                                For intCol = colState + 1 To SPD.MaxCols
                                    If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                        SetText SPD, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
                                
                                If SPD.Name = "vasID" Then
                                    '-- 결과 List
                                    SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                    SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                    SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                                    SetText vasRes, strResult, lsResRow, colRESULT          '결과
                                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                                    SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                
                                    '-- 로컬 저장
                                    Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes, lsEquipRes)
                                End If
                                
                                If SPD.Name = "vasID2" Then
                                    '-- 결과 List
                                    SetText vasRes2, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                    SetText vasRes2, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                    SetText vasRes2, lsExamName, lsResRow, colEXAMNAME       '검사명
                                    SetText vasRes2, lsEquipRes, lsResRow, colMachResult     '장비결과
                                    SetText vasRes2, strResult, lsResRow, colRESULT          '결과
                                    SetText vasRes2, lsSeqNo, lsResRow, colSeq               '순번
                                    SetText vasRes2, strComm, lsResRow, colFLAG              'Flag
                                    
                                    '-- 로컬 저장
                                    Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes2, lsEquipRes)
                                End If
                                
                                If SPD.Name = "vasID3" Then
                                    '-- 결과 List
                                    SetText vasRes3, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                                    SetText vasRes3, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                                    SetText vasRes3, lsExamName, lsResRow, colEXAMNAME       '검사명
                                    SetText vasRes3, lsEquipRes, lsResRow, colMachResult     '장비결과
                                    SetText vasRes3, strResult, lsResRow, colRESULT          '결과
                                    SetText vasRes3, lsSeqNo, lsResRow, colSeq               '순번
                                    SetText vasRes3, strComm, lsResRow, colFLAG              'Flag
                                    
                                    '-- 로컬 저장
                                    Call SetLocalDB(gRow, lsResRow, "1", SPD, vasRes3, lsEquipRes)
                                End If
                                
                                
                                
                                lsResult_Buff = ""
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                    End If
        
                Case 14
                    '## DB에 결과저장
                    If MnTransAuto.Checked = True And strState = "R" Then
                        Res = SaveTransDataW(gRow, SPD)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetText SPD, "0", gRow, colCheckBox
                            SetForeColor SPD, gRow, gRow, 1, colState, 255, 0, 0
                            SetText SPD, "Failed", gRow, colState
                        Else
                            '-- 저장 성공
                            SetBackColor SPD, gRow, gRow, 1, colState, 202, 255, 112
                            SetText SPD, "Trans", gRow, colState
                            SetText SPD, "0", gRow, colCheckBox
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gEquip & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(SPD, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(SPD, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(SPD, gRow, colSAVESEQ)) & vbCrLf
                            
                            Res = SendQuery(gLocal, SQL)
                            If Res = -1 Then
                                SaveQuery SQL
                                Exit Sub
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End If
    
Exit Sub

Rst:
    strBuffer3 = ""
    
    StatusBar1.Panels(2).Text = Err.Description
    SetRawData "[EditRcvDataASTM]" & Err.Number & ":" & Err.Description
    
End Sub


Function SetResult(asResult As String, asEquipCode As String)
    Dim i As Integer
    Dim sLVal As String
    Dim sHVal As String
    Dim sEquipCode As String
    Dim sEquipRes As String
    Dim sResult As String
    Dim sPoint As Integer
    Dim sResType As String
    Dim sResFlag As String
    
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    sResFlag = ""
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "select resprec, reflow, refhigh from EQPMASTER where equipcode = '" & sEquipCode & "' AND EQUIPNO = '" & gEquip & "' "
    Res = GetDBSelectColumn(gLocal, SQL)
    
    If IsNumeric(gReadBuf(0)) = True Then
        sPoint = CInt(gReadBuf(0))
        sResType = ""
        For i = 0 To sPoint
            If i = 0 Then
                sResType = "#0"
            ElseIf i = 1 Then
                sResType = sResType & ".0"
            Else
                sResType = sResType & "0"
            End If
        Next
        
        sResult = Format(sEquipRes, sResType)
    Else
        sResult = sEquipRes
    End If
    
''    If IsNumeric(gReadBuf(1)) = True Then
''        sLVal = gReadBuf(1)
''        If CCur(sLVal) > CCur(sEquipRes) Then
''            sResFlag = "H"
''        End If
''    End If
''
''    If IsNumeric(gReadBuf(2)) = True Then
''        sHVal = gReadBuf(2)
''        If CCur(sHVal) < CCur(sEquipRes) Then
''            sResFlag = ">"
''        End If
''    End If
    
    If IsNumeric(gReadBuf(1)) = True And IsNumeric(gReadBuf(2)) = True Then
        sLVal = gReadBuf(1)
        sHVal = gReadBuf(2)
        If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
            sResFlag = ""
        ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
            sResFlag = "H"
        ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
            sResFlag = "L"
        End If
    End If
    
    gsFlag = sResFlag
    SetResult = sResult
    
End Function

' asRow1 = Work List
' asRow2 = 결과 List
Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, SPD As Object, SPDR As Object, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String
    
    sExamDate = Format(dtpToday, "yyyymmddhhmmss")
    If Trim(GetText(SPD, asRow1, colSAVESEQ)) = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
          " WHERE EXAMDATE  = '" & Mid(sExamDate, 1, 8) & "' " & vbCrLf & _
          "   AND EQUIPNO   = '" & gEquip & "' " & vbCrLf & _
          "   AND SAVESEQ   = " & Trim(GetText(SPD, asRow1, colSAVESEQ)) & vbCrLf & _
          "   AND BARCODE   = '" & Trim(GetText(SPD, asRow1, colBARCODE)) & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & Trim(GetText(SPDR, asRow2, colEQUIPCODE)) & "'" & vbCrLf & _
          "   AND EXAMCODE  = '" & Trim(GetText(SPDR, asRow2, colEXAMCODE)) & "'"
          
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "INSERT INTO PATRESULT (" & vbCrLf
    SQL = SQL & "SAVESEQ"                           '저장순번(날짜별)
    SQL = SQL & ", EXAMDATE"                        '검사일자"
    SQL = SQL & ", HOSPDATE"                        '병원접수일자"
    SQL = SQL & ", EQUIPNO"                         '장비코드"
    SQL = SQL & ", BARCODE" & vbCrLf
    SQL = SQL & ", EQUIPCODE"                       '검사채널"
    SQL = SQL & ", EXAMCODE"                        '병원검사코드"
    SQL = SQL & ", EXAMSUBCODE"                     '병원검사코드(SUB)"
    SQL = SQL & ", EXAMNAME"
    SQL = SQL & ", SEQNO" & vbCrLf                  '검사일련번호"
    SQL = SQL & ", SAMPLETYPE"                      '검체유형"
    SQL = SQL & ", INOUT"                           '검체코드
    SQL = SQL & ", DISKNO"
    SQL = SQL & ", POSNO"
    SQL = SQL & ", EQUIPRESULT"                     '장비결과"
    SQL = SQL & ", RESULT" & vbCrLf                 '소수점적용결과"
    SQL = SQL & ", REFFLAG"
    SQL = SQL & ", REFVALUE"
    SQL = SQL & ", CHARTNO"
    SQL = SQL & ", PID"                             '병록번호(내원번호)"
    SQL = SQL & ", PNAME" & vbCrLf
    SQL = SQL & ", PSEX"
    SQL = SQL & ", PAGE"
    SQL = SQL & ", PJUMIN"
    SQL = SQL & ", PANICVALUE"
    SQL = SQL & ", DELTAVALUE" & vbCrLf
    SQL = SQL & ", SENDFLAG"                        '전송구분(0:미전송,1:전송)"
    SQL = SQL & ", SENDDATE"
    SQL = SQL & ", EXAMUID"
    SQL = SQL & ", HOSPITAL)" & vbCrLf
    SQL = SQL & " VALUES (" & vbCrLf
'    SQL = SQL & strSaveSeq
    SQL = SQL & Trim(GetText(SPD, asRow1, colSAVESEQ))
    SQL = SQL & ",'" & sExamDate
    SQL = SQL & "','" & Trim(GetText(SPD, asRow1, colHOSPDATE))
    SQL = SQL & "','" & gEquip
    SQL = SQL & "','" & Trim(GetText(SPD, asRow1, colBARCODE))
    SQL = SQL & "','" & Trim(GetText(SPDR, asRow2, colEQUIPCODE))
    SQL = SQL & "','" & Trim(GetText(SPDR, asRow2, colEXAMCODE))
    SQL = SQL & "','" & Trim(GetText(SPDR, asRow2, colSUBCODE))
    SQL = SQL & "','" & Trim(GetText(SPDR, asRow2, colEXAMNAME))
    SQL = SQL & "','" & Trim(GetText(SPDR, asRow2, colSeq))
    SQL = SQL & "','"
    SQL = SQL & "','" & Trim(GetText(SPD, asRow1, colINOUT))
    SQL = SQL & "','" & Trim(GetText(SPD, asRow1, colDISKNO))
    SQL = SQL & "','" & Trim(GetText(SPD, asRow1, colPOSNO))
    SQL = SQL & "','" & Trim(GetText(SPDR, asRow2, colMachResult))
    SQL = SQL & "','" & Trim(GetText(SPDR, asRow2, colRESULT))
    SQL = SQL & "','" & Trim(GetText(SPDR, asRow2, colFLAG))
    SQL = SQL & "',''"
    SQL = SQL & ",'" & Trim(GetText(SPD, asRow1, colCHARTNO))
    SQL = SQL & "','" & Trim(GetText(SPD, asRow1, colPID))
    SQL = SQL & "','" & Trim(GetText(SPD, asRow1, colPNAME))
    SQL = SQL & "','" & Trim(GetText(SPD, asRow1, colPSEX))
    SQL = SQL & "','" & Trim(GetText(SPD, asRow1, colPAGE))
    SQL = SQL & "',''"
    SQL = SQL & ",''"
    SQL = SQL & ",''"
    SQL = SQL & ",'1'"
    SQL = SQL & ",''"
    SQL = SQL & ",'" & gIFUser
    SQL = SQL & "','')"
    
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function



'-- 오늘 검사한 날짜의 Max + 1 번호를 가져온다
Private Function getMaxTestNum(ByVal strDate As String) As Long

    getMaxTestNum = 1
    
    '-- 결과업데이트
          SQL = "SELECT MAX(SAVESEQ) as SEQ FROM PATRESULT  "
    SQL = SQL & " WHERE MID(EXAMDATE,1,8) = '" & strDate & "' " & vbCrLf
    
    Res = GetDBSelectColumn(gLocal, SQL)
    
    If Res > 0 Then
        If Trim(gReadBuf(0)) = "" Then
            getMaxTestNum = 1
        Else
            getMaxTestNum = Trim(gReadBuf(0)) + 1
        End If
    End If
    
    If getMaxTestNum >= 99999 Then
        getMaxTestNum = 99999
    End If
    
End Function

Private Sub Var_Clear()
    
    gsBarCode = ""
    gsPID = ""
    gsRackNo = ""
    gsPosNo = ""
    gsResDateTime = ""
    gsSeqNo = ""
    gsExamCode = ""
    gsExamName = ""
    gsOrder = ""
    gsResult = ""

End Sub



Private Sub picLogin_Click()

    Dim sMsg As String
    sMsg = "검사자를 입력해주세요."
    lblUser.Caption = InputBox(sMsg, "검사자 입력")
    gUserID = lblUser.Caption
    
    Call WritePrivateProfileString("Server", "IFUser", gUserID, App.Path & "\Interface.ini")

    frmInterface.StatusBar1.Panels(1).Text = gUserID
    
End Sub

Private Sub txtBarNum_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim intRow  As Integer
    
    If KeyAscii = 13 Then
        If Not IsNumeric(txtBarNum) Then
            StatusBar1.Panels(3).Text = "바코드번호는 숫자만 입력이 가능합니다."
            txtBarNum = ""
            Exit Sub
        End If
        
        If Len(txtBarNum) <> 13 Then
            StatusBar1.Panels(3).Text = "바코드 자릿수를 확인하세요"
            txtBarNum = ""
            Exit Sub
        End If
        
        intRow = 0
        With vasID
            For i = 1 To .MaxRows
                .Row = i
                .Col = colBARCODE
                If Trim(.Text) = Trim(txtBarNum.Text) Then
                    intRow = i
                End If
            Next
            
            If intRow = 0 Then
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
            Else
                MsgBox "중복 바코드입니다.", vbOKOnly + vbCritical, Me.Caption
                Exit Sub
            End If
            
            Call SetText(vasID, txtBarNum.Text, intRow, colBARCODE)
            DoEvents
            
        End With
        
        
        
        If Trim(txtBarNum) <> "" Then
            'Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"), Trim(txtBarNum))
            Call GetSampleInfoW_PHILL(intRow, vasID)
        End If
        vasID.RowHeight(-1) = 12
        txtBarNum.Text = ""
    End If
    
End Sub


Private Sub txtRack_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If KeyAscii = 13 Then
        With vasID
            For i = .ActiveRow To .MaxRows
                .Row = i
                .Col = colDISKNO
                .Text = txtRack.Text
                txtRack.Text = txtRack.Text + 1
                If txtRack.Text = "31" Then
                    txtRack.Text = "1"
                End If
            Next
        End With
    End If
    
End Sub

Private Sub vasID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Integer
    
    If BlockRow <= 0 Then
        Exit Sub
    End If
    
    For i = BlockRow To BlockRow2
        vasID.Col = 1
        vasID.Row = i
        If vasID.Value = 0 Then
            vasID.Value = 1
        Else
            vasID.Value = 0
        End If
    Next i
    
End Sub


Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim RS          As ADODB.Recordset
    
    If Row = 0 Then
        With vasID
            .Col = 1: .Col2 = .MaxCols
            .Row = 2: .Row2 = .DataRowCnt
            .SortBy = 0
            .SortKey(1) = Col       '정렬키 열번호

            .SortKeyOrder(1) = SortKeyOrderAscending
    
            .Action = ActionSort
        End With
        Exit Sub
    End If
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
'    lblDate.Caption = Trim(GetText(vasID, Row, colHOSPDATE))
    lsID = Trim(GetText(vasID, Row, colBARCODE))
    lblChangeBar.Caption = lsID
    lblBarcode(0).Caption = lsID
    lblPname(0).Caption = Trim(GetText(vasID, Row, colPNAME))
    lblSaveSeq.Caption = Trim(GetText(vasID, Row, colSAVESEQ))
    lblExamDate.Caption = Trim(GetText(vasID, Row, colEXAMDATE))
    
    If lblSaveSeq.Caption = "" Then
        Exit Sub
    End If
    
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE " & vbCrLf
    SQL = SQL & "  FROM PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf
    SQL = SQL & "   AND SAVESEQ = " & lblSaveSeq.Caption & vbCrLf
    SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
    'SQL = SQL & "   AND EXAMDATE = '" & Mid(Trim(GetText(vasID, Row, colOrdDate)), 1, 8) & "' " & vbCrLf
    SQL = SQL & " GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE "
    SQL = SQL & " ORDER BY SEQNO * 10"
    
    Set RS = cn.Execute(SQL, , 1)

    If Not RS.EOF = True And Not RS.BOF = True Then
        vasRes.MaxRows = 0
        Do Until RS.EOF
            With vasRes
                .MaxRows = .MaxRows + 1
                SetText vasRes, "0", .MaxRows, colCheckBox
                SetText vasRes, Trim(RS.Fields("EQUIPCODE")) & "", .MaxRows, colEQUIPCODE
                SetText vasRes, Trim(RS.Fields("EXAMCODE")) & "", .MaxRows, colEXAMCODE
                SetText vasRes, Trim(RS.Fields("EXAMNAME")) & "", .MaxRows, colEXAMNAME
                SetText vasRes, Trim(RS.Fields("EQUIPRESULT")) & "", .MaxRows, colMachResult
                SetText vasRes, Trim(RS.Fields("RESULT")) & "", .MaxRows, colRESULT
                SetText vasRes, Trim(RS.Fields("SEQNO")) & "", .MaxRows, colSeq
                SetText vasRes, Trim(RS.Fields("REFFLAG")) & "", .MaxRows, colFLAG
                SetText vasRes, Trim(RS.Fields("EXAMSUBCODE")) & "", .MaxRows, colSUBCODE
                
                If Trim(RS.Fields("REFFLAG")) = "H" Then
                    .Row = .MaxRows
                    .Col = colRESULT
                    .ForeColor = vbRed
                ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
                    .Row = .MaxRows
                    .Col = colRESULT
                    .ForeColor = vbBlue
                End If
           
            End With
            RS.MoveNext
        Loop
    End If
    vasRes.RowHeight(-1) = 12
    
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow    As Long
    Dim iCol    As Long
    Dim lsID    As String
    Dim lsTime  As String
    Dim lsPid   As String
    Dim lsSeq   As String
    Dim i       As Integer
    Dim strResult As String
    Dim blnModify As Boolean
    
    blnModify = False
    
    iRow = vasID.ActiveRow
    iCol = vasID.ActiveCol

    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasID.DataRowCnt Then
            Exit Sub
        End If
        If iCol > colState Then
            Exit Sub
        End If
        lsID = Trim(GetText(vasID, iRow, colBARCODE))
        lsPid = Trim(GetText(vasID, iRow, colPID))
        lsSeq = Trim(GetText(vasID, iRow, colSAVESEQ))

'        If lsID = "" Or lsPid = "" Or lsSeq = "" Then
'            Exit Sub
'        End If
        If lsSeq = "" Then
            Exit Sub
        End If

        If MsgBox(lsSeq & " 의 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If

              SQL = "DELETE FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
        SQL = SQL & "   AND PID = '" & lsPid & "' " & vbCrLf
        SQL = SQL & "   AND SAVESEQ = " & lsSeq & vbCrLf
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
        Res = SendQuery(gLocal, SQL)

        If Res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If

        DeleteRow vasID, iRow, iRow
        vasRes.MaxRows = 0
        blnModify = True

    ElseIf KeyCode = vbKeyReturn Then
        If iCol = colBARCODE Then
            'Exit Sub
            
            '-- 바뀐 바코드로 환자정보 불러오기
            Call GetSampleInfoW_PHILL(iRow, vasID)
            
            lsID = Trim(GetText(vasID, iRow, colBARCODE))
            
            
            '-- 바코드 번호가 이전과 틀리다면 업데이트
            'If lsID <> lblChangeBar.Caption Then
            'If lsID <> lblBarcode(0).Caption Then
                      SQL = "UPDATE PATRESULT SET"
                SQL = SQL & " HOSPDATE = '" & Mid(Trim(GetText(vasID, iRow, colHOSPDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & ",BARCODE = '" & lsID & "' " & vbCrLf
                SQL = SQL & ",CHARTNO = '" & Trim(GetText(vasID, iRow, colCHARTNO)) & "' " & vbCrLf
                SQL = SQL & ",PID = '" & Trim(GetText(vasID, iRow, colPID)) & "' " & vbCrLf
                SQL = SQL & ",PNAME = '" & Trim(GetText(vasID, iRow, colPNAME)) & "' " & vbCrLf
                SQL = SQL & ",DISKNO = '" & Trim(GetText(vasID, iRow, colDISKNO)) & "' " & vbCrLf
                SQL = SQL & ",PSEX = '" & Trim(GetText(vasID, iRow, colPSEX)) & "' " & vbCrLf
                SQL = SQL & ",PAGE = '" & Trim(GetText(vasID, iRow, colPAGE)) & "' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(vasID, iRow, colSAVESEQ)) & vbCrLf
                SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Trim(GetText(vasID, iRow, colEXAMDATE)) & "' " & vbCrLf
                'SQL = SQL & "   AND BARCODE = '" & lblBarcode(0).Caption & "' "

                'SetRawData "[SQL]" & SQL
                Res = SendQuery(gLocal, SQL)
                
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If

                blnModify = True

            'End If
        ElseIf iCol = colDISKNO Then
            
        
        Else
            Exit Sub
            vasID.Row = iRow
            vasID.Col = colState
            If Trim(vasID.Text) = "" Then
                Exit Sub
            End If

            '-- 결과만 수정했을 경우의 업데이트는 Delete >> Insert 순으로 한다.
            '-- Delete
                  SQL = "DELETE FROM PATRESULT "
            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
            SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(vasID, iRow, colSAVESEQ)) & vbCrLf
            SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Trim(GetText(vasID, iRow, colEXAMDATE)) & "' " & vbCrLf
            SQL = SQL & "   AND BARCODE = '" & Trim(GetText(vasID, iRow, colBARCODE)) & "' "

            Res = SendQuery(gLocal, SQL)
                
            If Res = -1 Then
                SaveQuery SQL
                Exit Sub
            End If

            '-- Insert
            For i = colState + 1 To vasID.MaxCols
                vasID.Row = iRow
                vasID.Col = i
                If Trim(vasID.Text) <> "" Then
                    '-- 결과 소수점 적용
                    strResult = SetResult(Trim(GetText(vasID, iRow, i)), gArrEquip(i - colState, 2))
                    '-- H/L 일때 색표시
                    If gsFlag = "L" Then
                        vasID.Row = iRow
                        vasID.Col = i
                        vasID.ForeColor = vbBlue
                    ElseIf gsFlag = "H" Then
                        vasID.Row = iRow
                        vasID.Col = i
                        vasID.ForeColor = vbRed
                    End If
                    vasID.Text = strResult

                    SQL = ""
                    SQL = SQL & "INSERT INTO PATRESULT (" & vbCrLf
                    SQL = SQL & "SAVESEQ, EXAMDATE, HOSPDATE, EQUIPNO, BARCODE" & vbCrLf
                    SQL = SQL & ", EQUIPCODE, EXAMCODE, EXAMSUBCODE, EXAMNAME, SEQNO" & vbCrLf
                    SQL = SQL & ", SAMPLETYPE, DISKNO, POSNO, EQUIPRESULT, RESULT" & vbCrLf
                    SQL = SQL & ", REFFLAG, REFVALUE, CHARTNO, PID, PNAME" & vbCrLf
                    SQL = SQL & ", PSEX, PAGE, PJUMIN, PANICVALUE, DELTAVALUE" & vbCrLf
                    SQL = SQL & ", SENDFLAG, SENDDATE, EXAMUID, HOSPITAL)" & vbCrLf
                    SQL = SQL & " VALUES (" & vbCrLf
                    SQL = SQL & Trim(GetText(vasID, iRow, colSAVESEQ))
                    SQL = SQL & ",'" & Trim(GetText(vasID, iRow, colEXAMDATE))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colHOSPDATE))
                    SQL = SQL & "','" & gEquip
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colBARCODE))
                    'equipcode , examcode, examname, resprec, seqno
                    SQL = SQL & "','" & gArrEquip(i - colState, 2) 'Trim(GetText(vasRes, asRow2, colEQUIPCODE))
                    SQL = SQL & "','" & gArrEquip(i - colState, 3) 'Trim(GetText(vasRes, asRow2, colEXAMCODE))
                    SQL = SQL & "','"                              'Trim(GetText(vasRes, asRow2, colSubCode))
                    SQL = SQL & "','" & gArrEquip(i - colState, 4) 'Trim(GetText(vasRes, asRow2, colEXAMNAME))
                    SQL = SQL & "','" & gArrEquip(i - colState, 6) 'Trim(GetText(vasRes, asRow2, colSeq))
                    SQL = SQL & "',''"
                    SQL = SQL & ",'" & Trim(GetText(vasID, iRow, colDISKNO))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPOSNO))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, i)) 'Trim(GetText(vasRes, asRow2, colMachResult))
                    SQL = SQL & "','" & strResult 'Trim(GetText(vasID, iRow, i)) 'Trim(GetText(vasRes, asRow2, colRESULT))
                    SQL = SQL & "','" & gsFlag & "'"
                    SQL = SQL & ",''"
                    SQL = SQL & ",'" & Trim(GetText(vasID, iRow, colCHARTNO))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPID))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPNAME))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPSEX))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPAGE))
                    SQL = SQL & "',''"
                    SQL = SQL & ",''"
                    SQL = SQL & ",''"
                    SQL = SQL & ",'3'"
                    SQL = SQL & ",''"
                    SQL = SQL & ",'" & gIFUser
                    SQL = SQL & "','')"

                    Res = SendQuery(gLocal, SQL)
                    SetText vasID, "수정", iRow, colState

                End If
            Next
            blnModify = True
        End If
        'SetText vasID, "수정", iRow, colState

    End If
    
'    If blnModify = True Then
'        Call cmdRsltSearch_Click
'    End If
    
End Sub

Private Sub vasID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long

    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasID.ActiveRow
        If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Sub

        vasID_Click colBARCODE, lRow
    End If
End Sub


Private Sub vasRes_KeyPress(KeyAscii As Integer)
    Dim strResult   As String
    
    With vasRes
        If KeyAscii = 13 And .ActiveCol = colRESULT And lblBarcode(0).Caption <> "" Then
            '-- 결과 소수점 적용
            strResult = SetResult(Trim(GetText(vasRes, .ActiveRow, colRESULT)), Trim(GetText(vasRes, .ActiveRow, colEQUIPCODE)))
            .Col = colRESULT
            .Text = strResult
            '-- H/L 일때 색표시
            If gsFlag = "L" Then
                vasRes.Row = .ActiveRow
                vasRes.Col = colRESULT
                vasRes.ForeColor = vbBlue
            ElseIf gsFlag = "H" Then
                vasRes.Row = .ActiveRow
                vasRes.Col = colRESULT
                vasRes.ForeColor = vbRed
            End If
            
            SetText vasRes, gsFlag, .ActiveRow, colFLAG
            
            SQL = ""
            SQL = SQL & "UPDATE PATRESULT " & vbCrLf
            SQL = SQL & "   SET RESULT  ='" & strResult & "', " & vbCrLf
            SQL = SQL & "       REFFLAG    = '" & gsFlag & "' " & vbCrLf
            SQL = SQL & " WHERE BARCODE   = '" & Trim(lblBarcode(0).Caption) & "' " & vbCrLf
            SQL = SQL & "   AND MID(EXAMDATE,1,8)  = '" & Trim(lblExamDate.Caption) & "' " & vbCrLf
            SQL = SQL & "   AND SAVESEQ   = " & lblSaveSeq.Caption & vbCrLf
            SQL = SQL & "   AND EQUIPNO   = '" & gEquip & "' " & vbCrLf
            SQL = SQL & "   AND EXAMCODE  = '" & Trim(GetText(vasRes, .ActiveRow, colEXAMCODE)) & "' " & vbCrLf
            SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(vasRes, .ActiveRow, colEQUIPCODE)) & "' " & vbCrLf

            Res = SendQuery(gLocal, SQL)

            If Res = -1 Then
                SaveQuery SQL
                Exit Sub
            End If

        End If
    End With

End Sub


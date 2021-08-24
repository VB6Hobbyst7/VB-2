VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "SPIN32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Anato_Result 
   BorderStyle     =   0  '없음
   Caption         =   "진단서작성"
   ClientHeight    =   9120
   ClientLeft      =   1680
   ClientTop       =   1845
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "ANATO107.frx":0000
   ScaleHeight     =   9120
   ScaleWidth      =   12075
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin VB.ComboBox cmbChief2 
      Height          =   300
      Left            =   7620
      TabIndex        =   59
      ToolTipText     =   "결과완료일 때만 저장됩니다."
      Top             =   6750
      Width           =   1515
   End
   Begin Spin.SpinButton SpinButton1 
      Height          =   375
      Left            =   8730
      TabIndex        =   34
      Top             =   7080
      Width           =   405
      _Version        =   65536
      _ExtentX        =   720
      _ExtentY        =   656
      _StockProps     =   73
      Enabled         =   0   'False
   End
   Begin VB.TextBox txtSlid 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7620
      TabIndex        =   29
      Top             =   7080
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   90
      TabIndex        =   45
      Top             =   6660
      Width           =   6405
      Begin VB.ComboBox cmbEnzyme 
         Height          =   300
         Left            =   1860
         Style           =   2  '드롭다운 목록
         TabIndex        =   57
         Top             =   1890
         Width           =   4335
      End
      Begin VB.ComboBox cmbimmfludye 
         Height          =   300
         Left            =   1860
         Style           =   2  '드롭다운 목록
         TabIndex        =   56
         Top             =   1500
         Width           =   4335
      End
      Begin VB.ComboBox cmbimmdye 
         Height          =   300
         Left            =   1860
         Style           =   2  '드롭다운 목록
         TabIndex        =   55
         Top             =   1110
         Width           =   4335
      End
      Begin VB.ComboBox cmbSpecial 
         Height          =   300
         Left            =   1860
         Style           =   2  '드롭다운 목록
         TabIndex        =   54
         Top             =   720
         Width           =   4335
      End
      Begin Threed.SSPanel pnlFlow 
         Height          =   375
         Left            =   4995
         TabIndex        =   46
         Top             =   240
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   656
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   8.93
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSPanel pnlElectroScope 
         Height          =   375
         Left            =   1875
         TabIndex        =   47
         Top             =   240
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   661
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "효  소  염  색"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   3
         Left            =   210
         TabIndex        =   53
         ToolTipText     =   "특수염색 시약 입력"
         Top             =   1920
         Width           =   1650
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "면역 형광 염색"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   2
         Left            =   210
         TabIndex        =   52
         ToolTipText     =   "특수염색 시약 입력"
         Top             =   1530
         Width           =   1650
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "면  역  염  색"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   51
         ToolTipText     =   "특수염색 시약 입력"
         Top             =   1155
         Width           =   1650
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "특  수  염  색"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   6
         Left            =   210
         TabIndex        =   50
         ToolTipText     =   "특수염색 시약 입력"
         Top             =   750
         Width           =   1650
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Flow Cytometry"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   5
         Left            =   3360
         TabIndex        =   49
         ToolTipText     =   "특수염색 시약 입력"
         Top             =   300
         Width           =   1650
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "전 자 현 미 경"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   48
         ToolTipText     =   "특수염색 시약 입력"
         Top             =   300
         Width           =   1650
      End
   End
   Begin VB.TextBox txtDiagCodeName 
      Enabled         =   0   'False
      Height          =   372
      Left            =   7620
      TabIndex        =   44
      Top             =   8400
      Width           =   3900
   End
   Begin VB.TextBox txtDiagCode 
      Enabled         =   0   'False
      Height          =   372
      Left            =   7620
      TabIndex        =   42
      Top             =   7980
      Width           =   1452
   End
   Begin RichTextLib.RichTextBox txtView 
      Height          =   2172
      Left            =   276
      TabIndex        =   40
      Top             =   4296
      Visible         =   0   'False
      Width           =   8928
      _ExtentX        =   15769
      _ExtentY        =   3836
      _Version        =   393217
      BackColor       =   12648384
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"ANATO107.frx":030A
   End
   Begin Threed.SSCommand cmdViewAdd 
      Height          =   576
      Left            =   11592
      TabIndex        =   39
      Top             =   6276
      Width           =   336
      _Version        =   65536
      _ExtentX        =   572
      _ExtentY        =   995
      _StockProps     =   78
      Caption         =   "V"
   End
   Begin Threed.SSCommand cmdViewDiag 
      Height          =   576
      Left            =   11592
      TabIndex        =   38
      Top             =   5160
      Width           =   336
      _Version        =   65536
      _ExtentX        =   572
      _ExtentY        =   995
      _StockProps     =   78
      Caption         =   "V"
   End
   Begin Threed.SSCommand cmdViewPre 
      Height          =   576
      Left            =   11592
      TabIndex        =   37
      Top             =   4584
      Width           =   336
      _Version        =   65536
      _ExtentX        =   593
      _ExtentY        =   1016
      _StockProps     =   78
      Caption         =   "V"
   End
   Begin Threed.SSCommand cmdViewEye 
      Height          =   576
      Left            =   11592
      TabIndex        =   36
      Top             =   4008
      Width           =   336
      _Version        =   65536
      _ExtentX        =   572
      _ExtentY        =   995
      _StockProps     =   78
      Caption         =   "V"
   End
   Begin VB.ListBox lstSanGu 
      Appearance      =   0  '평면
      BackColor       =   &H00EBF5EB&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   90
      TabIndex        =   10
      Top             =   5148
      Visible         =   0   'False
      Width           =   9090
   End
   Begin VB.Frame frmPhoto 
      Enabled         =   0   'False
      Height          =   510
      Left            =   7620
      TabIndex        =   31
      Top             =   7440
      Width           =   1452
      Begin Threed.SSOption optPhoto 
         Height          =   210
         Index           =   0
         Left            =   165
         TabIndex        =   32
         Top             =   210
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   370
         _StockProps     =   78
         Caption         =   "Yes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   8.99
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optPhoto 
         Height          =   210
         Index           =   1
         Left            =   840
         TabIndex        =   33
         Top             =   210
         Width           =   510
         _Version        =   65536
         _ExtentX        =   900
         _ExtentY        =   370
         _StockProps     =   78
         Caption         =   "No"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   8.99
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin RichTextLib.RichTextBox txtRemark 
      Height          =   1215
      Left            =   4920
      TabIndex        =   27
      Top             =   3840
      Visible         =   0   'False
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   2143
      _Version        =   393217
      BackColor       =   15463915
      TextRTF         =   $"ANATO107.frx":060C
   End
   Begin VB.PictureBox Picture3 
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   11940
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
      Begin FPSpread.vaSpread ssResult 
         Height          =   912
         Left            =   72
         TabIndex        =   12
         Top             =   72
         Width           =   10836
         _Version        =   196608
         _ExtentX        =   19114
         _ExtentY        =   1609
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   8421376
         MaxCols         =   3
         MaxRows         =   200
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   12632256
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "ANATO107.frx":0902
         UserResize      =   0
         VisibleCols     =   3
         VisibleRows     =   200
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   920
         Left            =   10970
         TabIndex        =   13
         Top             =   70
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   1614
         _StockProps     =   78
         Caption         =   "종 료"
         ForeColor       =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO107.frx":14F0
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808000&
      Height          =   4536
      Left            =   9240
      ScaleHeight     =   4470
      ScaleWidth      =   2295
      TabIndex        =   4
      Top             =   3408
      Width           =   2352
      Begin Threed.SSCommand cmdAdditional 
         Height          =   570
         Left            =   1125
         TabIndex        =   35
         Top             =   2850
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2074
         _ExtentY        =   1016
         _StockProps     =   78
         Caption         =   "Additional"
         ForeColor       =   16711680
      End
      Begin Threed.SSCommand cmdSignOut 
         Height          =   570
         Left            =   1125
         TabIndex        =   20
         Top             =   2280
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   1005
         _StockProps     =   78
         Caption         =   "결과완료"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
      Begin Threed.SSCommand cmdFirstDiag 
         Height          =   570
         Left            =   1125
         TabIndex        =   22
         Top             =   1710
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   1005
         _StockProps     =   78
         Caption         =   "판독"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
      Begin Threed.SSCommand cmdPreliminary 
         Height          =   570
         Left            =   1125
         TabIndex        =   16
         Top             =   1140
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   1005
         _StockProps     =   78
         Caption         =   "Preliminary"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
      Begin Threed.SSCommand cmdEyeCheck 
         Height          =   570
         Left            =   1125
         TabIndex        =   17
         Top             =   570
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   1005
         _StockProps     =   78
         Caption         =   "육안검사"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
      Begin Threed.SSCommand cmdJCode 
         Height          =   570
         Left            =   0
         TabIndex        =   18
         ToolTipText     =   "사인아웃 할경우에 DATA 저장됨"
         Top             =   1710
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2074
         _ExtentY        =   1016
         _StockProps     =   78
         Caption         =   "진단코드등록"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   576
         Left            =   1128
         TabIndex        =   19
         Top             =   3900
         Width           =   1176
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   1005
         _StockProps     =   78
         Caption         =   "종   료"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   10.5
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
      Begin Threed.SSCommand cmsCommon 
         Height          =   570
         Left            =   0
         TabIndex        =   21
         Top             =   1140
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   1005
         _StockProps     =   78
         Caption         =   "상용구절"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
      Begin Threed.SSCommand cmdHistory 
         Height          =   570
         Left            =   1125
         TabIndex        =   23
         Top             =   0
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   1005
         _StockProps     =   78
         Caption         =   "병력사항"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
      Begin Threed.SSCommand cmdRecept 
         Height          =   576
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   1176
         _Version        =   65536
         _ExtentX        =   2074
         _ExtentY        =   1016
         _StockProps     =   78
         Caption         =   "환자선택"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
      Begin Threed.SSCommand cmdMacro 
         Height          =   570
         Left            =   0
         TabIndex        =   25
         Top             =   570
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   1005
         _StockProps     =   78
         Caption         =   "MACRO"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  '위 맞춤
      Height          =   1005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12075
      _Version        =   65536
      _ExtentX        =   21299
      _ExtentY        =   1773
      _StockProps     =   15
      Caption         =   "ANATOMIC   PATHOLOGY"
      ForeColor       =   8388608
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Font3D          =   2
      Begin VB.PictureBox Picture1 
         Height          =   465
         Left            =   9540
         ScaleHeight     =   405
         ScaleWidth      =   2115
         TabIndex        =   1
         Top             =   270
         Width           =   2175
         Begin VB.Label Label6 
            Caption         =   "User:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   180
            TabIndex        =   3
            Top             =   90
            Width           =   690
         End
         Begin VB.Label lblUser 
            Caption         =   "********"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   945
            TabIndex        =   2
            Top             =   90
            Width           =   1095
         End
      End
   End
   Begin VB.ListBox lstPtInfo 
      BackColor       =   &H00EBF5EB&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   9240
      TabIndex        =   8
      Top             =   1470
      Width           =   2325
   End
   Begin RichTextLib.RichTextBox txtDiag 
      Height          =   5070
      Left            =   75
      TabIndex        =   26
      Top             =   1440
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   8943
      _Version        =   393217
      BackColor       =   14482170
      ScrollBars      =   2
      TextRTF         =   $"ANATO107.frx":1942
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label11 
      Caption         =   "Chief Sub"
      Height          =   225
      Left            =   6675
      TabIndex        =   58
      Top             =   6810
      Width           =   930
   End
   Begin VB.Label Label10 
      Caption         =   "진단코드명"
      Height          =   300
      Left            =   6675
      TabIndex        =   43
      Top             =   8460
      Width           =   930
   End
   Begin VB.Label Label9 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "진단코드"
      Height          =   300
      Left            =   6675
      TabIndex        =   41
      Top             =   8055
      Width           =   930
   End
   Begin VB.Label Label8 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "육안사진"
      Height          =   225
      Left            =   6675
      TabIndex        =   30
      Top             =   7620
      Width           =   930
   End
   Begin VB.Label Label7 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "보관Slide"
      Height          =   225
      Left            =   6675
      TabIndex        =   28
      Top             =   7155
      Width           =   930
   End
   Begin VB.Label lblRowId 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   1  '단일 고정
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00800000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "작업내용"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   9240
      TabIndex        =   7
      Top             =   3060
      Width           =   2352
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00800000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "판 독 결 과 입 력"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   60
      TabIndex        =   5
      Top             =   1140
      Width           =   9120
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   144
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00800000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "환자정보"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   9240
      TabIndex        =   6
      Top             =   1140
      Width           =   2325
   End
End
Attribute VB_Name = "Anato_Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LbRemark                As Boolean
Dim LbSan                   As Boolean
Dim LbEletroFlag            As Integer
'Dim LbSpeSlideFlag          As Integer
Dim LbSpeGeomchFlag         As Integer
    
Dim SlidCnt

Dim txtViewB                As Boolean

Dim sSpecial()              As String


Private Sub cmdAdditional_Click()
    Dim Response
   
    If lblRowId.Caption = "" Then Exit Sub
    If txtDiag.Text = "" Then Exit Sub
 
    Response = MsgBox("Additional 결과를 저장하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "진단병리과")
    
    If Response = vbNo Then Exit Sub
    
    If Len(Quot(txtDiag.Text)) > 1000 Then
        Response = MsgBox("Additional DATA 길이가 1000Byte가 넘습니다." & vbCrLf & _
                          "1000Byte 이후는 저장되지 않습니다." & vbCrLf & _
                          "저장하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "진단병리과")
        If Response = vbNo Then Exit Sub
    End If
    
    txtDiag.Text = Mid(Quot(txtDiag.Text), 1, 950)
    
    strSQL = ""
    strSQL = strSQL & " UPDATE TWANAT_DIAG "
    strSQL = strSQL & "    SET GBRESULT = '9', "
    strSQL = strSQL & "        DIAGDATE = TO_DATE('" & GsExDate & "','YYYY-MM-DD'), "
    strSQL = strSQL & "        DiagAdd  = '" & txtDiag.Text & "' "
    strSQL = strSQL & "  WHERE ROWID    = '" & lblRowId & "'"

    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "Additional 저장 완료되었습니다.", vbInformation, "진단병리과"
        
        Call frm_Clear
        txtView.Visible = False
        Call Pt_Select
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
    
    
End Sub


Private Sub cmdEyeCheck_Click()
    '육안검사결과
    Dim Response
   
    If lblRowId.Caption = "" Then Exit Sub
    If txtDiag.Text = "" Then Exit Sub

    Response = MsgBox("육안결과를 저장하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "진단병리과")
    
    If Response = vbNo Then Exit Sub
    
    If Len(Quot(txtDiag.Text)) >= 2950 Then
        Response = MsgBox("육안결과 DATA 길이가 3000Byte가 넘습니다." & vbCrLf & _
                          "3000Byte 이후는 저장되지 않습니다." & vbCrLf & _
                          "저장하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "진단병리과")
        If Response = vbNo Then Exit Sub
    End If
    
    txtDiag.Text = Mid(Quot(txtDiag.Text), 1, 2950)
    
    strSQL = ""
    strSQL = strSQL & " UPDATE TWANAT_DIAG "
    strSQL = strSQL & "    SET GBGROSS  = '1', "
'    strSQL = strSQL & "        GBRESULT = '1', "
    strSQL = strSQL & "        DIAGDATE = TO_DATE('" & GsExDate & "','YYYY-MM-DD'), "
    strSQL = strSQL & "        DiagEye  = '" & txtDiag.Text & "' "
    strSQL = strSQL & "  WHERE ROWID    = '" & lblRowId & "'"

    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "육안결과 저장 완료되었습니다.", vbInformation, "진단병리과"
        
        Call frm_Clear
        txtView.Visible = False
        Call Pt_Select
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If

End Sub

Private Sub cmdJCode_Click()
    '진단코드등록
    GDict = "M"
    
    Anato_Jindan_Code.Show vbModal
    
    txtDiagCode.Text = GJindan
    txtDiagCodeName.Text = GPJindan
    
End Sub

Private Sub cmdPreliminary_Click()
    Dim Response
   
    If lblRowId.Caption = "" Then Exit Sub
    If txtDiag.Text = "" Then Exit Sub
 
    Response = MsgBox("Preliminary 결과를 저장하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "진단병리과")
    
    If Response = vbNo Then Exit Sub
    
    If Len(Quot(txtDiag.Text)) > 1000 Then
        Response = MsgBox("Preliminary DATA 길이가 1000Byte가 넘습니다." & vbCrLf & _
                          "1000Byte 이후는 저장되지 않습니다." & vbCrLf & _
                          "저장하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "진단병리과")
        If Response = vbNo Then Exit Sub
    End If
    
    txtDiag.Text = Mid(Quot(txtDiag.Text), 1, 950)
    
    strSQL = ""
    strSQL = strSQL & " UPDATE TWANAT_DIAG "
    strSQL = strSQL & "    SET GBRESULT = '2', "
    strSQL = strSQL & "       CHIEF     = '" & GstrPassIDnumber & "', "
    strSQL = strSQL & "       DIAGDATE  = TO_DATE('" & GsExDate & "','YYYY-MM-DD'), "
    strSQL = strSQL & "       DRREMARK  = '" & Quot(txtRemark.Text) & "', "
'    strSQL = strSQL & "       SpeGeomch = '" & LbSpeGeomchFlag & "', "
    strSQL = strSQL & "       DiagCode  = '" & Trim(txtDiagCode.Text) & "', "
    strSQL = strSQL & "       slid      = '" & txtSlid.Text & "', "
    strSQL = strSQL & "       DiagPre   = '" & txtDiag.Text & "' "
    strSQL = strSQL & "  WHERE ROWID    = '" & lblRowId & "'"

    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "Preliminary 저장 완료되었습니다.", vbInformation, "진단병리과"
        
        Call frm_Clear
        txtView.Visible = False
        Call Pt_Select
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
    
End Sub

Private Sub cmdViewAdd_Click()
    'Additional 결과 조회
    strSQL = ""
    strSQL = strSQL & " SELECT DiagAdd "
    strSQL = strSQL & "   FROM TWANAT_Diag "
    strSQL = strSQL & "  WHERE RowID = '" & LsRowID & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
        txtView.Text = rs.Fields("DiagAdd").Value & ""
    End If
    AdoCloseSet rs

    If txtViewB = True Then
        txtView.Visible = True
        txtViewB = False
    Else
        txtView.Visible = False
        txtViewB = True
    End If

End Sub

Private Sub cmdViewDiag_Click()
    '판독결과 조회
    strSQL = ""
    strSQL = strSQL & " SELECT Descr "
    strSQL = strSQL & "   FROM TWANAT_Diag "
    strSQL = strSQL & "  WHERE RowID = '" & LsRowID & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
        txtView.Text = rs.Fields("Descr").Value & ""
    Else
        txtView.Text = ""
    End If
    AdoCloseSet rs
    
    If txtViewB = True Then
        txtView.Visible = True
        txtViewB = False
    Else
        txtView.Visible = False
        txtViewB = True
    End If
    
End Sub

Private Sub cmdViewEye_Click()
    '육안검사결과 조회
    strSQL = ""
    strSQL = strSQL & " SELECT DiagEye "
    strSQL = strSQL & "   FROM TWANAT_Diag "
    strSQL = strSQL & "  WHERE RowID = '" & LsRowID & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
        txtView.Text = rs.Fields("DiagEye").Value & ""
    Else
        txtView.Text = ""
    End If
    AdoCloseSet rs
    
    If txtViewB = True Then
        txtView.Visible = True
        txtViewB = False
    Else
        txtView.Visible = False
        txtViewB = True
    End If
    
End Sub

Private Sub cmdViewPre_Click()
    'Preliminary결과 조회
    strSQL = ""
    strSQL = strSQL & " SELECT DiagPre "
    strSQL = strSQL & "   FROM TWANAT_Diag "
    strSQL = strSQL & "  WHERE RowID = '" & LsRowID & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
        txtView.Text = rs.Fields("DiagPre").Value & ""
    Else
        txtView.Text = ""
    End If
    AdoCloseSet rs
    
    If txtViewB = True Then
        txtView.Visible = True
        txtViewB = False
    Else
        txtView.Visible = False
        txtViewB = True
    End If

End Sub

Private Sub Form_Activate()
    
    Clipboard.Clear
    Clipboard.SetText "■"
    
'    txtSlid.Text = 1

End Sub


Private Sub Form_Load()
    
    Dim rs                  As ADODB.Recordset
   
    Dim i                   As Integer
    
    lblUser = GstrPassName
    
    txtViewB = True
    lstSanGu.Clear
    GsDiagNo = ""
    GsSpecial = ""
    LbEletroFlag = "0"
    
'    LbSpeSlideFlag = "0"
    LbSpeGeomchFlag = "0"
    
    GsHistology = "NO"
    GsCytology = "NO"
    GsGross = "NO"
    GsFirst = "NO"
    GsComplete = "NO"
    GsJSHistology = "NO"
    GsJSCytology = "NO"
    
    strSQL = ""
    strSQL = strSQL & " SELECT *"
    strSQL = strSQL & " FROM   TWEXAM_REMARK "
    strSQL = strSQL & " WHERE  ExGubun = 'AN'"
    strSQL = strSQL & " ORDER  BY AbbCode"
''    If False = adoSetOpen(strSQL, adoSet) Then Exit Sub
    
    Result = AdoOpenSet(rs, strSQL)
        
    If Result = True Then
        Do Until rs.EOF
            lstSanGu.AddItem rs.Fields("Abbname").Value & ""
            rs.MoveNext
        Loop
    End If
    AdoCloseSet rs
    
    strSQL = ""
    strSQL = strSQL & " SELECT *"
    strSQL = strSQL & " FROM   TWBAS_DOCTOR "
    strSQL = strSQL & " WHERE  DRDEPT1 = 'AP' OR DRDEPT2 = 'AP' "
    strSQL = strSQL & "   AND  GBOUT   = 'N' "
    strSQL = strSQL & " ORDER  BY DRNAME "
    
    Result = AdoOpenSet(rs, strSQL)
        
    If Result = False Then Exit Sub
    
    cmbChief2.AddItem ""
    
    Do Until rs.EOF
        cmbChief2.AddItem rs.Fields("DRCODE").Value & "" & " " & rs.Fields("DRNAME").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Anato_Main.Show

End Sub


Private Sub cmdCancel_Click()

    Picture3.Visible = False

End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    Anato_Main.Show
    
End Sub


Private Sub cmdFirstDiag_Click()
    '판독
    
    Dim Response            As Integer
   
    If lblRowId.Caption = "" Then Exit Sub
    If txtDiag.Text = "" Then Exit Sub
    
    Response = MsgBox(" 판독결과를 저장하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "진단병리과")
    
    If Response = vbNo Then Exit Sub
    
    If Len(Quot(txtDiag.Text)) > 3000 Then
        Response = MsgBox("Preliminary DATA 길이가 3000Byte가 넘습니다." & vbCrLf & _
                          "3000Byte 이후는 저장되지 않습니다." & vbCrLf & _
                          "저장하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "진단병리과")
        If Response = vbNo Then Exit Sub
    End If
    
    txtDiag.Text = Mid(Quot(txtDiag.Text), 1, 2980)
    
    strSQL = ""
    strSQL = strSQL & " UPDATE TWANAT_DIAG"
    strSQL = strSQL & " SET    GBRESULT  = '3',"
    strSQL = strSQL & "        DESCR     = '" & txtDiag.Text & "', "
'    strSQL = strSQL & "       SpeGeomch  = '" & LbSpeGeomchFlag & "', "
'    strSQL = strSQL & "       Diagno     = '" & Quot(GsDiagNo) & "', "
    strSQL = strSQL & "       DIAGDATE   = TO_DATE('" & GsExDate & "','YYYY-MM-DD'), "
    strSQL = strSQL & "       DiagCode   = '" & txtDiagCode.Text & "', "
    strSQL = strSQL & "       slid       = '" & txtSlid.Text & "', "
    If optPhoto(0).Value = True Then
        strSQL = strSQL & "       Photo      = 'Y'"
    Else
        strSQL = strSQL & "       Photo      = 'N'"
    End If
    strSQL = strSQL & " WHERE  ROWID     = '" & lblRowId & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
        
        Call frm_Clear
        txtView.Visible = False
        Call Pt_Select
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If

End Sub


Private Sub cmdHistory_Click()
    '병력사항
    If LbRemark = False Then
        txtRemark.Visible = True
        LbRemark = True
        cmdHistory.Font3D = 2
    Else
        txtRemark.Visible = False
        LbRemark = False
        cmdHistory.Font3D = 0
    End If
    
End Sub


Private Sub cmdMacro_Click()
    
    If lstPtInfo.ListCount <= 1 Then
        MsgBox " 환자를 선택하십시요."
        Exit Sub
    End If
    Anato_Macro_View.Left = 10
    Anato_Macro_View.Top = 1000
    Anato_Macro_View.Show vbModal

End Sub


Private Sub cmdRecept_Click()
    '환자선택
    
    GReceptSeq = 0
    
    lstPtInfo.Clear
    lblRowId = ""
    txtDiag.Text = ""
    
    
    Picture3.Visible = False
    Anato_Jeobsu_View.Left = 485       '3220
    Anato_Jeobsu_View.Top = 900         '900

    Set Anato_Jeobsu_View = Nothing
    
    GAnato_Jeobsu_View = True
    GReceptSeq = 0
    txtView.Visible = False
    
    Anato_Jeobsu_View.Show vbModal
    
    '환자선택에서 넘어온 Data 전처리
    If GAnato_Jeobsu_View = False Then Exit Sub
    
    Call Pt_Select
    
    
End Sub

Private Sub cmdSignOut_Click()
    '결과완료
    
    Dim Response            As Integer
   
    If lblRowId.Caption = "" Then Exit Sub
    
    Call Jindan_Reg
    
'    If GsDiagNo = "" Then
'        Response = MsgBox("진단명을 확인하세요!", vbOKOnly + vbInformation + vbDefaultButton1, "진단병리과")
'        Exit Sub
'    End If
    
    Response = MsgBox("검사결과를 영구 저장하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "진단병리과")
    
    If Response = vbNo Then Exit Sub
    LbSpeGeomchFlag = "1"
        
    If Len(Quot(txtDiag.Text)) > 3000 Then
        Response = MsgBox("Preliminary DATA 길이가 3000Byte가 넘습니다." & vbCrLf & _
                          "3000Byte 이후는 저장되지 않습니다." & vbCrLf & _
                          "저장하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "진단병리과")
        If Response = vbNo Then Exit Sub
    End If
    
    txtDiag.Text = Mid(Quot(txtDiag.Text), 1, 2980)
    
    If Len(GsDiagNo) > 2000 Then
        MsgBox " DIAGNOSIS: 이후의 DATA길이가 2000Byte가 넘습니다." & vbCrLf & _
               " 2000Byte 미만으로 줄이십시요."
        Exit Sub
    End If


    strSQL = ""
    strSQL = strSQL & " UPDATE TWANAT_DIAG "
    strSQL = strSQL & " SET   GBRESULT   = '4', "
    strSQL = strSQL & "       CHIEF      = '" & GstrPassIDnumber & "', "
    strSQL = strSQL & "       DIAGDATE   = TO_DATE('" & GsExDate & "','YYYY-MM-DD'), "
'    strSQL = strSQL & "       DESCR      = '" & txtDiag.Text & "', "
    strSQL = strSQL & "       DRREMARK   = '" & Quot(txtRemark.Text) & "', "
    strSQL = strSQL & "       SpeGeomch  = '" & LbSpeGeomchFlag & "', "
    strSQL = strSQL & "       Diagno     = '" & Quot(GsDiagNo) & "', "
'    strSQL = strSQL & "       DiagCode   = '" & txtDiagCode.Text & "', "
    strSQL = strSQL & "       slid       = '" & txtSlid.Text & "', "
    strSQL = strSQL & "       Chief2     = '" & Mid(cmbChief2.Text, 1, 6) & "', "
    If optPhoto(0).Value = True Then
        strSQL = strSQL & "       Photo      = 'Y'"
    Else
        strSQL = strSQL & "       Photo      = 'N'"
    End If
    
    strSQL = strSQL & " WHERE ROWID      = '" & lblRowId & "' "
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        GsDiagNo = ""
        MsgBox "결과완료가 저장 완료되었습니다.", vbInformation, "진단병리과"
        
        Call frm_Clear
        txtView.Visible = False
        Call Pt_Select
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If

    txtDiag.Locked = False
    Anato_DiagName_Input.txtSummary.Locked = True

End Sub




Private Sub Jindan_Reg()
    '진단명
    Dim rs                  As ADODB.Recordset
    
    Dim LiLength            As Integer
    Dim LiPos               As Integer
    Dim LsSearchChar        As String
    
    If lblRowId.Caption = "" Then Exit Sub
    
    GsDiagNo = ""
    
    strSQL = ""
    strSQL = strSQL & " SELECT GbResult, Diagno "
    strSQL = strSQL & "   FROM TWANAT_DIAG "
    strSQL = strSQL & "  WHERE RowID = '" & lblRowId & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
'    If Result = True And Rowindicator = 1 Then
    If Result Then
        If rs.Fields("GBRESULT").Value & "" = "9" Then
            GsDiagNo = Trim(rs.Fields("DIAGNO").Value & "")
            Exit Sub
        
        End If
        AdoCloseSet rs
    
    End If
    
    If GsDiagNo = "" Then
        LsSearchChar = "DIAGNOSIS:"
        LiPos = InStr(txtDiag.Text, LsSearchChar)
        
        If LiPos <> 0 Then
            LiLength = Len(txtDiag.Text)
            GsDiagNo = Mid(txtDiag.Text, LiPos + 10, LiLength - LiPos + 10)
            
'            If Len(Anato_DiagName_Input.txtSummary.Text) > 1000 Then                                     '300 => 1000
'                GsDiagNo = Left(GsDiagNo, 1000)  '300 => 1000
'            End If
            
        Else
             GsDiagNo = Trim(GsDiagNo)
        End If
        
    Else
        GsDiagNo = GsDiagNo
    End If
        
End Sub


Private Sub cmsCommon_Click()
    
    If LbSan = False Then
        lstSanGu.Visible = True
        LbSan = True
        cmsCommon.Font3D = 2
    Else
        lstSanGu.Visible = False
        LbSan = False
        cmsCommon.Font3D = 0
    End If
    
End Sub


Private Sub lstSanGu_DblClick()
    
    txtDiag.Text = txtDiag.Text & vbCrLf & lstSanGu.List(lstSanGu.ListIndex)
    
    lstSanGu.Visible = False
    
    LbSan = False

End Sub



Private Sub SpinButton1_SpinDown()
    If txtSlid.Text = 1 Then Exit Sub
    txtSlid.Text = txtSlid.Text - 1
    SlidCnt = txtSlid.Text

End Sub

Private Sub SpinButton1_SpinUp()
    txtSlid.Text = txtSlid.Text + 1
    SlidCnt = txtSlid.Text

End Sub


Private Sub ssResult_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ssResult.Height = 2860
    Picture3.Height = 3060
    
End Sub

Private Sub txtDiag_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ssResult.Height = 888
    Picture3.Height = 1028

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ssResult.Height = 888
    Picture3.Height = 1028

End Sub



Private Sub txtSlid_GotFocus()
    txtSlid.SelStart = 0
    txtSlid.SelLength = Len(txtSlid.Text)

End Sub

Private Sub txtSlid_LostFocus()
    SlidCnt = txtSlid.Text

End Sub


Private Sub Pt_Select()
    '선택완료 Sub Routine
    
    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    Dim j                   As Integer
    Dim LsPtNo              As String
    Dim LsOrderDt           As String
    Dim LsRemark            As String
    
    Dim SpecialChar
    
    lstPtInfo.Clear
    lblRowId = ""
    LsRowID = ""   ' 추가 12/14
    
    ReDim sSpecial(0)
    ReDim sSpecial(30)
    
    GReceptSeq = GReceptSeq + 1
    GobjectSS.Row = GReceptSeq
        
    GobjectSS.Col = 1:
    For i = GReceptSeq To GobjectSS.DataRowCnt
        If GobjectSS.Text = 1 Then
            Exit For
        Else
            GReceptSeq = GReceptSeq + 1
            GobjectSS.Row = GReceptSeq
        End If
    Next i
    
    If GReceptSeq > GobjectSS.DataRowCnt Then Exit Sub
    
    '2  접수master
    '3  ptno
    '4  Sname
    '5  육안결과
    '6  sex
    '7  age
    '8  orderdt
    '9  gbgross
    '10 gbresult
    '11 jdate
    '12 slid
    '13 photo
    '14 Special
    '15 immdye
    '16 immfludye
    '17 Enzyme
    '18 Electroscope
    '19 Flow
    '20 Rowid
    
    
    GobjectSS.Col = 2:        lstPtInfo.AddItem GobjectSS.Text
    GobjectSS.Col = 3:        lstPtInfo.AddItem GobjectSS.Text
                              LsPtNo = GobjectSS.Text
                              '''''''''''''''''''
                              GoSub OLD_DATA_READ
    GobjectSS.Col = 4:        lstPtInfo.AddItem GobjectSS.Text
    GobjectSS.Col = 6:        lstPtInfo.AddItem GobjectSS.Text
    GobjectSS.Col = 7:        lstPtInfo.AddItem GobjectSS.Text
    GobjectSS.Col = 8:        lstPtInfo.AddItem GobjectSS.Text
                              LsOrderDt = GobjectSS.Text
    GobjectSS.Col = 11:       lstPtInfo.AddItem GobjectSS.Text
    
    GobjectSS.Col = 12:
                              If GobjectSS.Text <> "" Then
                                  txtSlid.Text = GobjectSS.Text
                              Else
                                  txtSlid.Text = 1
                              End If
    GobjectSS.Col = 13:
                              If GobjectSS.Text = "Y" Then
                                  optPhoto(0).Value = True
                              Else
                                  optPhoto(1).Value = True
                              End If
    
'''    GobjectSS.Col = 14:       pnlSpecial.Caption = "  " & Specode_Get(GobjectSS.Text, 83)
'''    GobjectSS.Col = 15:       pnlimmdye.Caption = "  " & Specode_Get(GobjectSS.Text, 87)
'''    GobjectSS.Col = 16:       pnlimmfludye.Caption = "  " & Specode_Get(GobjectSS.Text, 84)
'''    GobjectSS.Col = 17:       pnlEnzyme.Caption = "  " & Specode_Get(GobjectSS.Text, 86)
    
'    GobjectSS.Col = 18:       pnlElectroScope.Caption = "  " & GobjectSS.Text
    GobjectSS.Col = 19:       pnlFlow.Caption = "  " & GobjectSS.Text
    
    GobjectSS.Col = 20:       lblRowId = GobjectSS.Text
                              LsRowID = GobjectSS.Text
            
    GobjectSS.Col = 5:
                              
                              If RCheck = "1" Then
                                    strSQL = ""
                                    strSQL = strSQL & " SELECT * "
                                    strSQL = strSQL & " FROM   TWANAT_DIAG "
                                    strSQL = strSQL & " WHERE  ROWID  = '" & LsRowID & "' "
                                
                                    Result = AdoOpenSet(rs, strSQL)
                                    If Result Then
                                        Do Until rs.EOF
                                            txtDiag.Text = Trim(rs.Fields("diageye").Value & "")
                                            
                                            rs.MoveNext
                                        Loop
                                    End If
                                    
                                    AdoCloseSet rs
                              ElseIf RCheck = "2" Then
                                    strSQL = ""
                                    strSQL = strSQL & " SELECT * "
                                    strSQL = strSQL & " FROM   TWANAT_DIAG "
                                    strSQL = strSQL & " WHERE  ROWID  = '" & LsRowID & "' "
                                
                                    Result = AdoOpenSet(rs, strSQL)
                                    If Result Then
                                        Do Until rs.EOF
                                            txtDiag.Text = Trim(rs.Fields("Descr").Value & "")
                                            
                                            rs.MoveNext
                                        Loop
                                    End If
                                    
                                    AdoCloseSet rs
                              End If
    
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & " FROM   TWANAT_DIAG "
    strSQL = strSQL & " WHERE  ROWID  = '" & LsRowID & "' "

    Result = AdoOpenSet(rs, strSQL)
    If Result Then
        Do Until rs.EOF
            
            For i = 1 To 30
                SpecialChar = "SPECIAL" & Format(i, "00")
                sSpecial(i) = Trim(rs.Fields(SpecialChar).Value & "")
            Next i
            
            rs.MoveNext
        Loop
    End If
    
    AdoCloseSet rs
    
    
    
'    lstSpecial.Clear
'    lstimmdye.Clear
'    lstimmfludye.Clear
'    lstEnzyme.Clear
    
    For i = 1 To 30
        Select Case sSpecial(i)
                Case "853001" To "853999"   '특수염색
                    cmbSpecial.AddItem sSpecial(i) & "  " & Special_Load(sSpecial(i))
                
                Case "857001" To "857999"   '면역조직화학검사
                    cmbimmdye.AddItem sSpecial(i) & "  " & Special_Load(sSpecial(i))
                
                Case "854001" To "854999"   '조직면역형광검사
                    cmbimmfludye.AddItem sSpecial(i) & "  " & Special_Load(sSpecial(i))
                
                Case "856001" To "856999"   '효소조직화학
                    cmbEnzyme.AddItem sSpecial(i) & "  " & Special_Load(sSpecial(i))
                Case "855001"               '전자현미경검사
                    pnlElectroScope.Caption = " Y"
        End Select
    Next i
    
    cmbSpecial.ListIndex = cmbSpecial.ListCount - 1
    cmbimmdye.ListIndex = cmbimmdye.ListCount - 1
    cmbimmfludye.ListIndex = cmbimmfludye.ListCount - 1
    cmbEnzyme.ListIndex = cmbEnzyme.ListCount - 1
    
    strSQL = ""
    strSQL = strSQL & " SELECT REMARK4 "
    strSQL = strSQL & " FROM   TWOCS_OCLINICAL "
    strSQL = strSQL & " WHERE  PTNO = '" & LsPtNo & "'"
    strSQL = strSQL & " AND    BDATE = to_date('" & LsOrderDt & "','yyyy-mm-dd')"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
        LsRemark = rs.Fields("REMARK4").Value & ""
        txtRemark.Text = ""
        txtRemark.Text = MidH(LsRemark, 1, 30) & vbCrLf
        txtRemark.Text = txtRemark.Text & MidH(LsRemark, 31, 120) & vbCrLf
        txtRemark.Text = txtRemark.Text & MidH(LsRemark, 151, 30) & vbCrLf
        txtRemark.Text = txtRemark.Text & MidH(LsRemark, 181, 120) & vbCrLf
        AdoCloseSet rs
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If Anato_Jeobsu_View.optDiag = True Then
        strSQL = ""
        strSQL = strSQL & " SELECT descr, DiagCode "
        strSQL = strSQL & "   FROM TWANAT_Diag "
        strSQL = strSQL & "  WHERE RowID = '" & LsRowID & "'"
        
        Result = AdoOpenSet(rs, strSQL)
        
        If Result Then
            txtDiag.Text = rs.Fields("DESCR").Value & ""
            txtDiagCode.Text = rs.Fields("DiagCode").Value & ""
            txtDiagCodeName.Text = DiagCodeSearch(txtDiagCode.Text)
            rs.MoveNext
        End If
        AdoCloseSet rs
    End If
    
    If Anato_Jeobsu_View.optAdditional = True Then
        txtDiag.Text = ""
    
        strSQL = ""
        strSQL = strSQL & " SELECT DiagAdd "
        strSQL = strSQL & "   FROM TWANAT_Diag "
        strSQL = strSQL & "  WHERE RowID = '" & LsRowID & "'"
        
        Result = AdoOpenSet(rs, strSQL)
    
        If Result Then
            txtDiag.Text = rs.Fields("DiagAdd").Value & ""
            rs.MoveNext
        End If
        AdoCloseSet rs
    
    End If
    
    Exit Sub
    
'-------------------------------------------------------------------------------------
OLD_DATA_READ:
    Dim adoOLD              As ADODB.Recordset
    
    Call SSInitialize(ssResult)
    
    strSQL = ""
    strSQL = strSQL & " SELECT CLASS, DATEYY, SEQNUM, DIAGNO, GBRESULT, GBGROSS  "
    strSQL = strSQL & "   FROM TWANAT_DIAG  "
    strSQL = strSQL & "  WHERE PTNO     = '" & LsPtNo & "' "
    strSQL = strSQL & "    AND GbResult <> 'X' "
    strSQL = strSQL & " ORDER  BY GBRESULT ASC             "
    
    Result = AdoOpenSet(adoOLD, strSQL)
    
    If Result = False Then Return
    
    ssResult.MaxRows = Rowindicator + 1
    
    Do Until adoOLD.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        
        ssResult.Col = 1:
            ssResult.Text = adoOLD.Fields("CLASS").Value & "-" & _
                                         adoOLD.Fields("DATEYY").Value & "-" & _
                                         adoOLD.Fields("SEQNUM").Value & ""
        
        ssResult.Col = 2:
            ssResult.Text = Replace(adoOLD.Fields("Diagno").Value & "", vbCrLf, "", 1, -1, vbTextCompare)
            ssResult.RowHeight(ssResult.Row) = ssResult.MaxTextRowHeight(ssResult.Row)
            
        ssResult.Col = 3:
            Select Case adoOLD.Fields("GbResult").Value & ""
                Case "0"
                    If adoOLD.Fields("GbGross").Value & "" = "1" Then
                        ssResult.Text = "육안검사"
                    Else
                        ssResult.Text = "접수중"
                    End If
                Case "1"
                    ssResult.Text = "육안검사"
                Case "2"
                    ssResult.Text = "Preliminary"
                Case "3"
                    ssResult.Text = "판독"
                Case "4"
                    ssResult.Text = "결과완료"
                Case "9"
                    ssResult.Text = "Additional"
                Case "X"
                    ssResult.Text = "접수취소"
                Case Else
            End Select
        
        adoOLD.MoveNext
    Loop
    
    AdoCloseSet adoOLD
    
    Picture3.Visible = True
    Return
    
End Sub


Private Sub frm_Clear()
    txtDiag.Text = ""
    txtView.Text = ""
    
    pnlElectroScope.Caption = ""
    pnlFlow.Caption = ""
    
    cmbSpecial.Clear
    cmbimmdye.Clear
    cmbimmfludye.Clear
    cmbEnzyme.Clear
    
    txtSlid.Text = "1"
    optPhoto(1).Value = True
    txtDiagCode.Text = ""
    txtDiagCodeName.Text = ""
    
    
    
End Sub



Private Sub txtView_Click()
    txtView.Visible = False

End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Private Sub cmdGross_Click()
    'GROSS
    
    Dim Response
   
    If lblRowId.Caption = "" Then Exit Sub
 
    Response = MsgBox("GROSS 결과를 저장하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "진단병리과")
    
    If Response = vbNo Then Exit Sub
    
    strSQL = ""
    strSQL = strSQL & " UPDATE TWANAT_DIAG "
    strSQL = strSQL & "    SET GBGROSS   = '1', "
    strSQL = strSQL & "        DESCR     = '" & Quot(txtDiag.Text) & "' "
    strSQL = strSQL & "  WHERE ROWID     = '" & lblRowId & "'"

    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "Gross 저장 완료되었습니다.", vbInformation, "진단병리과"
        
        Call frm_Clear
        Call Pt_Select
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If

End Sub


'미사용
Private Sub cmdSpeGeomch_Click()
    '보관검체
    
    Dim Response
   
    If lblRowId.Caption = "" Then Exit Sub
 
    Response = MsgBox("검체를 보관하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "진단병리과")
    
    If Response = vbNo Then
        Exit Sub
    Else
        LbSpeGeomchFlag = "1"
    End If
        
   
    strSQL = ""
    strSQL = strSQL & " UPDATE TWANAT_DIAG"
    strSQL = strSQL & " SET    SpeGeomch  = '" & LbSpeGeomchFlag & "'"
    strSQL = strSQL & " WHERE  ROWID      = '" & lblRowId & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
'        MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If

End Sub


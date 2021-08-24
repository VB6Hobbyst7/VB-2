VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm313QCCancel_N 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   75
   ClientWidth     =   14745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   14745
   WindowState     =   2  '최대화
   Begin VB.CommandButton BeFoRe 
      Caption         =   "이전"
      Height          =   495
      Left            =   7080
      TabIndex        =   56
      Top             =   8535
      Width           =   1215
   End
   Begin VB.CommandButton NExT 
      Caption         =   "다음"
      Height          =   495
      Left            =   8520
      TabIndex        =   55
      Top             =   8535
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   34
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   33
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "저장(S)"
      Height          =   510
      Left            =   10485
      Style           =   1  '그래픽
      TabIndex        =   32
      Top             =   8535
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblTestList 
      Height          =   5115
      Left            =   75
      TabIndex        =   31
      Top             =   3315
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   9022
      _StockProps     =   64
      AllowUserFormulas=   -1  'True
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14936810
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   12
      MaxRows         =   25
      Protect         =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      SpreadDesigner  =   "frm313QCCancel_N.frx":0000
      VisibleCols     =   4
      VisibleRows     =   24
   End
   Begin MedControls1.LisLabel LisLabel8 
      Height          =   300
      Left            =   11070
      TabIndex        =   24
      Top             =   45
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 바코드 재발행"
      Appearance      =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   2340
      Left            =   11070
      TabIndex        =   19
      Top             =   270
      Width           =   3390
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   12
         Left            =   75
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   480
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "접수 번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   13
         Left            =   75
         TabIndex        =   52
         Top             =   885
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "바코드번호"
         Appearance      =   0
      End
      Begin VB.CommandButton cmdReprint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "재발행"
         Enabled         =   0   'False
         Height          =   510
         Left            =   1785
         Style           =   1  '그래픽
         TabIndex        =   23
         Top             =   1470
         Width           =   1320
      End
      Begin VB.CommandButton cmdReprintList 
         BackColor       =   &H00E0E0E0&
         Caption         =   "재발행 대상"
         Height          =   510
         Left            =   360
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   1470
         Width           =   1320
      End
      Begin MedControls1.LisLabel lblAccNo 
         Height          =   360
         Left            =   1350
         TabIndex        =   20
         Top             =   480
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "03-031014-1"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBarNo 
         Height          =   360
         Left            =   1350
         TabIndex        =   21
         Top             =   885
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "12000020342"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel6 
      Height          =   300
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 컨트롤 정보"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1425
      Left            =   75
      TabIndex        =   1
      Top             =   270
      Width           =   10980
      Begin VB.CommandButton cmdPopCtrl 
         BackColor       =   &H00F4F0F2&
         Height          =   360
         Left            =   3915
         Picture         =   "frm313QCCancel_N.frx":2FE4
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   165
         Width           =   330
      End
      Begin VB.TextBox txtCtrlCd 
         Height          =   375
         Left            =   1425
         MaxLength       =   9
         TabIndex        =   2
         Text            =   "하둘셋넷다여일여아"
         Top             =   165
         Width           =   2490
      End
      Begin MedControls1.LisLabel lblCtrlNm 
         Height          =   360
         Left            =   4275
         TabIndex        =   4
         Top             =   165
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblCtrlDiv 
         Height          =   360
         Left            =   5190
         TabIndex        =   5
         Top             =   570
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "내부정도관리"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblEqp 
         Height          =   360
         Left            =   8325
         TabIndex        =   6
         Top             =   570
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "C001 Coulter Stks"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBuilding 
         Height          =   360
         Left            =   1425
         TabIndex        =   7
         Top             =   975
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "10 본원"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSection 
         Height          =   360
         Left            =   5190
         TabIndex        =   8
         Top             =   975
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "HE Hematology"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWorkarea 
         Height          =   360
         Left            =   8340
         TabIndex        =   9
         Top             =   975
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "03 Hematology"
         Appearance      =   0
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   465
         Left            =   1425
         TabIndex        =   10
         Top             =   480
         Width           =   2490
         Begin VB.OptionButton optLevelCd 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Low"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   13
            Top             =   150
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.OptionButton optLevelCd 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Normal"
            Height          =   180
            Index           =   1
            Left            =   765
            TabIndex        =   12
            Top             =   150
            Width           =   960
         End
         Begin VB.OptionButton optLevelCd 
            BackColor       =   &H00DBE6E6&
            Caption         =   "High"
            Height          =   180
            Index           =   2
            Left            =   1740
            TabIndex        =   11
            Top             =   150
            Width           =   705
         End
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   11
         Left            =   45
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   165
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Control 정보"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   45
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   975
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "건물구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   45
         TabIndex        =   46
         Top             =   570
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Level 구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   7305
         TabIndex        =   47
         Top             =   570
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "검사장비"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   7305
         TabIndex        =   48
         Top             =   975
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Workarea"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   3915
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   570
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "정도관리구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   3915
         TabIndex        =   50
         Top             =   975
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "섹션구분"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   990
      Left            =   75
      TabIndex        =   14
      Top             =   1620
      Width           =   10980
      Begin VB.ComboBox cboLotNo 
         Height          =   300
         Left            =   1425
         Style           =   2  '드롭다운 목록
         TabIndex        =   25
         Top             =   165
         Width           =   3060
      End
      Begin MedControls1.LisLabel lblOpenDt 
         Height          =   360
         Left            =   5760
         TabIndex        =   15
         Top             =   165
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "2003/10/23"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblExpDt 
         Height          =   360
         Left            =   9195
         TabIndex        =   16
         Top             =   165
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "2003/10/23"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblMakeCd 
         Height          =   360
         Left            =   1425
         TabIndex        =   17
         Top             =   570
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "03 Hematology"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRemark 
         Height          =   360
         Left            =   5775
         TabIndex        =   18
         Top             =   570
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "03 Hematology"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   8
         Left            =   4485
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   165
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "시작일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   9
         Left            =   4500
         TabIndex        =   40
         Top             =   570
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "비  고"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   45
         TabIndex        =   41
         Top             =   165
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Lot No."
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   7
         Left            =   45
         TabIndex        =   42
         Top             =   570
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "제조사"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   10
         Left            =   7920
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   165
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "만료일"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel9 
      Height          =   255
      Left            =   75
      TabIndex        =   26
      Top             =   2610
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   450
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 검사항목 정보"
      Appearance      =   0
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00DBE6E6&
      Height          =   540
      Left            =   75
      TabIndex        =   27
      Top             =   2775
      Width           =   14385
      Begin VB.CheckBox chkSelAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체 선택"
         Height          =   180
         Left            =   105
         TabIndex        =   28
         Top             =   225
         Width           =   1110
      End
      Begin MedControls1.LisLabel lblAllCnt 
         Height          =   360
         Left            =   11715
         TabIndex        =   29
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "999"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSelCnt 
         Height          =   360
         Left            =   13635
         TabIndex        =   30
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "999"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   14
         Left            =   12360
         TabIndex        =   53
         Top             =   120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "선택된 항목수"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   15
         Left            =   10440
         TabIndex        =   54
         Top             =   120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "전체 항목수"
         Appearance      =   0
      End
   End
   Begin FPSpread.vaSpread tblWorkList 
      Height          =   735
      Left            =   5040
      TabIndex        =   57
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
      _Version        =   196608
      _ExtentX        =   2566
      _ExtentY        =   1296
      _StockProps     =   64
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   12
      MaxRows         =   3
      ScrollBars      =   0
      SpreadDesigner  =   "frm313QCCancel_N.frx":3096
   End
   Begin VB.Label lblSpcNo 
      Height          =   195
      Left            =   13815
      TabIndex        =   38
      Top             =   2535
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblSpcYY 
      Height          =   195
      Left            =   13785
      TabIndex        =   37
      Top             =   2130
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblColTm 
      Height          =   195
      Left            =   13770
      TabIndex        =   36
      Top             =   1785
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblColDt 
      Height          =   330
      Left            =   13800
      TabIndex        =   35
      Top             =   1395
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frm313QCCancel_N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Coding By Legends
'처방낼때랑 바코드 재출력할때가 좀 얄탁구리함..
'나중에 전체적으로 한번 손봐야 될거 같음...

Public Event LastFormUnload()

Private objQC As clsQcMst
Private objOrder As clsQcOrder

Private CalledMe As Boolean '외부에서 이 화면이 불리워졌는지 판단.

Private mvarParentHwnd As Long
'-------------------------------
' 2009.03.12 양성현 추가
Private currWorkList As Integer
'-------------------------------

Public Property Let ParentHwnd(ByVal vData As Long)
    mvarParentHwnd = vData
End Property

Public Property Get ParentHwnd() As Long
    ParentHwnd = mvarParentHwnd
End Property

Public Sub CallByExternal(ByVal pCtrlCd As String, ByVal pLevelCd As String)
    txtCtrlCd.Text = ""
    Call InitControl
    cboLotNo.Clear
    Call InitLotNo
    
    Call medClearTable(tblTestList)
    tblTestList.MaxRows = 16
    tblTestList.Col = -1
    tblTestList.Row = -1
    tblTestList.BlockMode = True
    tblTestList.CellType = CellTypeStaticText
    tblTestList.BlockMode = False

    Dim Rs As Recordset
    Dim strSQL As String
    
    Set Rs = GetControlInfo(pCtrlCd, pLevelCd)
            
    txtCtrlCd.Text = Rs.Fields("ctrlcd").Value & ""
    lblCtrlNm.Caption = Rs.Fields("ctrlnm").Value & ""
    
    If Rs.Fields("levelcd").Value & "" = "L" Then
        optLevelCd(0).Value = True
    ElseIf Rs.Fields("levelcd").Value & "" = "N" Then
        optLevelCd(1).Value = True
    ElseIf Rs.Fields("levelcd").Value & "" = "H" Then
        optLevelCd(2).Value = True
    End If
    
    lblCtrlDiv.Caption = IIf(Rs.Fields("ctrldiv").Value & "" = "I", "내부정도관리", "외부정도관리")
    lblEqp.Caption = Format(Rs.Fields("eqpcd").Value & "", "!" & String(5, "@")) & Rs.Fields("eqpnm").Value & ""
    lblEqp.ToolTipText = Format(Rs.Fields("eqpcd").Value & "", "!" & String(5, "@")) & Rs.Fields("eqpnm").Value & ""
    lblBuilding.Caption = Format(Rs.Fields("buildcd").Value & "", "!" & String(5, "@")) & Rs.Fields("buildnm").Value & ""
    lblBuilding.ToolTipText = Format(Rs.Fields("buildcd").Value & "", "!" & String(10, "@")) & Rs.Fields("buildnm").Value & ""
    lblSection.Caption = Format(Rs.Fields("sectcd").Value & "", "!" & String(5, "@")) & Rs.Fields("sectnm").Value & ""
    lblSection.ToolTipText = Format(Rs.Fields("sectcd").Value & "", "!" & String(5, "@")) & Rs.Fields("sectnm").Value & ""
    lblWorkarea.Caption = Format(Rs.Fields("workarea").Value & "", "!" & String(5, "@")) & Rs.Fields("workareanm").Value & ""
    lblWorkarea.ToolTipText = Format(Rs.Fields("workarea").Value & "", "!" & String(5, "@")) & Rs.Fields("workareanm").Value & ""
    
    Set Rs = Nothing
    
    Call LoadLotNo
    Call LoadTestItem
End Sub

Private Sub cboLotNo_Click()
    On Error Resume Next
    If Screen.ActiveControl.Name <> cboLotNo.Name Then Exit Sub
    
    Call InitLotNo
    
    lblOpenDt.Caption = medGetP(cboLotNo.Text, 2, COL_DIV)
    lblExpDt.Caption = medGetP(cboLotNo.Text, 3, COL_DIV)
    lblMakeCd.Caption = medGetP(cboLotNo.Text, 4, COL_DIV)
    lblRemark.Caption = medGetP(cboLotNo.Text, 5, COL_DIV)

    lblCtrlNm.Caption = medGetP(cboLotNo.Text, 6, COL_DIV)
    
    Call LoadTestItem
End Sub

Private Sub chkSelAll_Click()
    Dim i As Long, j As Long
    
    On Error Resume Next
    
    If Screen.ActiveControl.Name <> chkSelAll.Name Then Exit Sub
    
    If tblTestList.DataRowCnt = 0 Then Exit Sub
    
    With tblTestList
        For i = 1 To .DataRowCnt
            .Row = i
            For j = 1 To .DataColCnt Step 3
            
            .Col = j
            .Value = IIf(chkSelAll.Value = 1, 1, 0)
            Next j
        Next i
    End With
    
    lblSelCnt.Caption = IIf(chkSelAll.Value = 1, lblAllCnt.Caption, "0")
End Sub

Private Sub cmdClear_Click()
    txtCtrlCd.Text = ""
    Call InitControl
    cboLotNo.Clear
    Call InitLotNo
    
    Call medClearTable(tblTestList)
    tblTestList.MaxRows = 16
    tblTestList.Col = -1
    tblTestList.Row = -1
    tblTestList.BlockMode = True
    tblTestList.CellType = CellTypeStaticText
    tblTestList.BlockMode = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
    If IsLastForm Then Call UnloadForm(Me)
'    If IsLastForm Then
'        If mvarParentHwnd <> 0 Then
'            Call SendMessage(mvarParentHwnd, WM_CLOSE, 0&, 0&)
'        End If
'    End If
End Sub

Private Sub cmdPopCtrl_Click()
    If lblCtrlNm.Caption <> "" Then
        DoEvents
        txtCtrlCd.Text = ""
        Call InitControl
        cboLotNo.Clear
        Call InitLotNo
        
        Call medClearTable(tblTestList)
        tblTestList.MaxRows = 16
        tblTestList.Col = -1
        tblTestList.Row = -1
        tblTestList.BlockMode = True
        tblTestList.CellType = CellTypeStaticText
        tblTestList.BlockMode = False
    End If
    
    DoEvents
    Call LoadControlInfo
    DoEvents
    
    Call LoadLotNo
    DoEvents
    Call LoadTestItem
End Sub
'--------------------------------------------------------
' 2009.03.12 양성현 추가
Private Sub BeFoRe_Click()
    Call BeforeSave
End Sub

Private Sub NExT_Click()
    Call AfterSave
End Sub

Private Sub AfterSave()
Dim strlblEqp    As String
Dim strlblEqpTip As String
Dim strlblBldn   As String
Dim strlblBldnTp As String
Dim strlblSec    As String
Dim strlblSecTip As String
Dim strlblWre    As String
Dim strlblWreTip As String
    
    
    
    
    currWorkList = currWorkList + 1
    With tblWorkList
        If .MaxRows < currWorkList Then
            currWorkList = .MaxRows
            MsgBox "마지막 자료 입니다."
        Else
    
    cboLotNo.Clear
    Call InitLotNo
'        Debug.Print "After : " & currWorkList
            .Row = currWorkList
            .Col = 1:   txtCtrlCd.Text = .Value
            .Col = 2:   lblCtrlNm.Caption = .Value
            .Col = 3:
                If Trim(.Value) = "L" Then
                    optLevelCd(0).Value = True
                ElseIf Trim(.Value) = "N" Then
                    optLevelCd(1).Value = True
                ElseIf Trim(.Value) = "H" Then
                    optLevelCd(2).Value = True
                End If
            .Col = 4:   lblCtrlDiv.Caption = IIf(Trim(.Value) = "I", "내부정도관리", "외부정도관리")
            .Col = 5:   strlblEqp = Trim(.Value)
            .Col = 6:   strlblEqpTip = Trim(.Value)
            .Col = 7:   strlblBldn = Trim(.Value)
            .Col = 8:   strlblBldnTp = Trim(.Value)
            .Col = 9:   strlblSec = Trim(.Value)
            .Col = 10:  strlblSecTip = Trim(.Value)
            .Col = 11:  strlblWre = Trim(.Value)
            .Col = 12:  strlblWreTip = Trim(.Value)
        
            lblEqp.Caption = Format(strlblEqp, "!" & String(5, "@")) & strlblEqpTip
            lblEqp.ToolTipText = Format(strlblBldn, "!" & String(10, "@")) & strlblEqpTip

            lblBuilding.Caption = Format(strlblBldn, "!" & String(5, "@")) & strlblBldnTp
            lblBuilding.ToolTipText = Format(strlblBldn, "!" & String(10, "@")) & strlblBldnTp
        
            lblSection.Caption = Format(strlblSec, "!" & String(5, "@")) & strlblSecTip
            lblSection.ToolTipText = Format(strlblSec, "!" & String(10, "@")) & strlblSecTip
        
            lblWorkarea.Caption = Format(strlblWre, "!" & String(5, "@")) & strlblWreTip
            lblWorkarea.ToolTipText = Format(strlblWre, "!" & String(10, "@")) & strlblWreTip
            DoEvents
            Call LoadLotNo
            DoEvents
            Call LoadTestItem
        End If
    End With
End Sub

Private Sub BeforeSave()
Dim strlblEqp    As String
Dim strlblEqpTip As String
Dim strlblBldn   As String
Dim strlblBldnTp As String
Dim strlblSec    As String
Dim strlblSecTip As String
Dim strlblWre    As String
Dim strlblWreTip As String
    
    
    currWorkList = currWorkList - 1
    If currWorkList < 1 Then
        currWorkList = 1
        MsgBox "처음 자료 입니다."
    Else
'        Debug.Print "Before : " & currWorkList
    
    cboLotNo.Clear
    Call InitLotNo
        With tblWorkList
            .Row = currWorkList
            .Col = 1:   txtCtrlCd.Text = .Value
            .Col = 2:   lblCtrlNm.Caption = .Value
            .Col = 3:
                If Trim(.Value) = "L" Then
                    optLevelCd(0).Value = True
                ElseIf Trim(.Value) = "N" Then
                    optLevelCd(1).Value = True
                ElseIf Trim(.Value) = "H" Then
                    optLevelCd(2).Value = True
                End If
            .Col = 4:   lblCtrlDiv.Caption = IIf(Trim(.Value) = "I", "내부정도관리", "외부정도관리")
            .Col = 5:   strlblEqp = Trim(.Value)
            .Col = 6:   strlblEqpTip = Trim(.Value)
            .Col = 7:   strlblBldn = Trim(.Value)
            .Col = 8:   strlblBldnTp = Trim(.Value)
            .Col = 9:   strlblSec = Trim(.Value)
            .Col = 10:  strlblSecTip = Trim(.Value)
            .Col = 11:  strlblWre = Trim(.Value)
            .Col = 12:  strlblWreTip = Trim(.Value)
        
            lblEqp.Caption = Format(strlblEqp, "!" & String(5, "@")) & strlblEqpTip
            lblEqp.ToolTipText = Format(strlblBldn, "!" & String(10, "@")) & strlblEqpTip

            lblBuilding.Caption = Format(strlblBldn, "!" & String(5, "@")) & strlblBldnTp
            lblBuilding.ToolTipText = Format(strlblBldn, "!" & String(10, "@")) & strlblBldnTp
        
            lblSection.Caption = Format(strlblSec, "!" & String(5, "@")) & strlblSecTip
            lblSection.ToolTipText = Format(strlblSec, "!" & String(10, "@")) & strlblSecTip
        
            lblWorkarea.Caption = Format(strlblWre, "!" & String(5, "@")) & strlblWreTip
            lblWorkarea.ToolTipText = Format(strlblWre, "!" & String(10, "@")) & strlblWreTip
        End With
        DoEvents
        Call LoadLotNo
        DoEvents
        Call LoadTestItem
    End If
End Sub
'----------------------------------------------------------------
' 2009.02.18 양성현 수정
Private Function GetpCtrlCd() As String
    With tblWorkList
        .Row = currWorkList
        .Col = 5
        GetpCtrlCd = Mid(.Value, 1, 2) & "" & Format(medGetP(.Value, 2, "-"), LIS_BarFormat)
    End With
End Function


Private Sub LoadControlInfo(Optional ByVal pCtrlCd As String = "")
'컨트롤의 일반 정보를 불러온다..
    Dim objPop As clsPopUpList
    Dim i As Long
    Dim j As Integer
    Dim strLevel    As String
    
    Set objPop = New clsPopUpList

    With objPop
        .Recordset = GetControlInfo(pCtrlCd)
        
        .FormCaption = "컨트롤 찾기"
        .Delimiter = COL_DIV
'        .FormWidth = 4470
        .FormWidth = 8470
        .ColumnHeaderText = "코드컨트롤명Level구분장비코드장비명건물코드건물명섹션코드섹션명워크애랴코드워크애랴명"
        .ColumnHeaderWidth = "1254.92922775.213629.8583001905.213000001405.213"
        .ColumnHeaderAlign = "002"
        '0 왼쪽, 1 오른쪽, 2 가운데
        
        Call .LoadPopUp
        
        DoEvents
        
        txtCtrlCd.Text = medGetP(.SelectedString, 1, .Delimiter)
        lblCtrlNm.Caption = medGetP(.SelectedString, 2, .Delimiter)
        strLevel = medGetP(.SelectedString, 3, .Delimiter)
        If strLevel = "L" Then
            optLevelCd(0).Value = True
        ElseIf strLevel = "N" Then
            optLevelCd(1).Value = True
        ElseIf strLevel = "H" Then
            optLevelCd(2).Value = True
        End If

        lblCtrlDiv.Caption = IIf(medGetP(.SelectedString, 4, .Delimiter) = "I", "내부정도관리", "외부정도관리")
        lblEqp.Caption = Format(medGetP(.SelectedString, 5, .Delimiter), "!" & String(5, "@")) & medGetP(.SelectedString, 6, .Delimiter)
        lblEqp.ToolTipText = Format(medGetP(.SelectedString, 5, .Delimiter), "!" & String(10, "@")) & medGetP(.SelectedString, 6, .Delimiter)
        lblBuilding.Caption = Format(medGetP(.SelectedString, 7, .Delimiter), "!" & String(5, "@")) & medGetP(.SelectedString, 8, .Delimiter)
        lblBuilding.ToolTipText = Format(medGetP(.SelectedString, 7, .Delimiter), "!" & String(10, "@")) & medGetP(.SelectedString, 8, .Delimiter)
        lblSection.Caption = Format(medGetP(.SelectedString, 9, .Delimiter), "!" & String(5, "@")) & medGetP(.SelectedString, 10, .Delimiter)
        lblSection.ToolTipText = Format(medGetP(.SelectedString, 9, .Delimiter), "!" & String(10, "@")) & medGetP(.SelectedString, 10, .Delimiter)
        lblWorkarea.Caption = Format(medGetP(.SelectedString, 11, .Delimiter), "!" & String(5, "@")) & medGetP(.SelectedString, 12, .Delimiter)
        lblWorkarea.ToolTipText = Format(medGetP(.SelectedString, 11, .Delimiter), "!" & String(10, "@")) & medGetP(.SelectedString, 12, .Delimiter)
        
        .Recordset.MoveFirst

            tblWorkList.MaxRows = .Recordset.RecordCount
            currWorkList = 1
            For i = 1 To tblWorkList.MaxRows
                 tblWorkList.Row = i
            
                If txtCtrlCd.Text = .Recordset.Fields("ctrlcd").Value & "" And strLevel = .Recordset.Fields("levelcd").Value & "" Then
                    currWorkList = i
                End If
                 tblWorkList.Col = 1: tblWorkList.Value = .Recordset.Fields("ctrlcd").Value & ""
                 tblWorkList.Col = 2: tblWorkList.Value = .Recordset.Fields("ctrlnm").Value & ""
                 tblWorkList.Col = 3: tblWorkList.Value = .Recordset.Fields("levelcd").Value & ""
                 tblWorkList.Col = 4: tblWorkList.Value = .Recordset.Fields("ctrldiv").Value & ""
                 tblWorkList.Col = 5: tblWorkList.Value = .Recordset.Fields("eqpcd").Value & ""
                 tblWorkList.Col = 6: tblWorkList.Value = .Recordset.Fields("eqpnm").Value & ""
                 tblWorkList.Col = 7: tblWorkList.Value = .Recordset.Fields("buildcd").Value & ""
                 tblWorkList.Col = 8: tblWorkList.Value = .Recordset.Fields("buildnm").Value & ""
                 tblWorkList.Col = 9: tblWorkList.Value = .Recordset.Fields("sectcd").Value & ""
                 tblWorkList.Col = 10: tblWorkList.Value = .Recordset.Fields("sectnm").Value & ""
                 tblWorkList.Col = 11: tblWorkList.Value = .Recordset.Fields("workarea").Value & ""
                 tblWorkList.Col = 12: tblWorkList.Value = .Recordset.Fields("workareanm").Value & ""
                .Recordset.MoveNext
            Next i

    End With
    
'    Debug.Print "Start : " & currWorkList
    
    Set objPop = Nothing
End Sub

Private Function GetControlInfo(Optional ByVal pCtrlCd As String = "", _
                                Optional ByVal pLevelCd As String = "") As Recordset
    Dim strSQL As String
    
    strSQL = " select distinct a.ctrlcd,a.ctrlnm,a.levelcd,a.ctrldiv,a.eqpcd,c.eqpnm, a.buildcd,d.field1 as buildnm, " & _
            " a.sectcd,e.field1 as sectnm, a.workarea, f.field1 as workareanm " & _
            " from " & T_LAB021 & " a, " & T_LAB023 & " g, " & T_LAB006 & " c, " & T_LAB032 & " d, " & T_LAB032 & " e, " & T_LAB032 & " f " & _
            " where " & DBJ("a.eqpcd*=c.eqpcd") & _
            " and " & DBW("d.cdindex=", LC3_Buildings) & _
            " and a.buildcd=d.cdval1 " & _
            " and " & DBW("e.cdindex=", LC3_Section) & _
            " and a.sectcd=e.cdval1 " & _
            " and " & DBW("f.cdindex=", LC3_WorkArea) & _
            " and a.workarea=f.cdval1 and a.ctrlcd = g.ctrlcd and  g.expdt >= to_char(sysdate,'YYYYMMDD') "

    If pCtrlCd <> "" Then
        strSQL = strSQL & " and " & DBW("a.ctrlcd=", pCtrlCd)
    End If
    
    If pLevelCd <> "" Then
        strSQL = strSQL & " and " & DBW("a.levelcd=", pLevelCd)
    End If

'    strSQL = strSQL & " order by a.ctrlcd,ctrlnm,levelcd"
    strSQL = strSQL & " order by a.ctrlcd, eqpcd, levelcd,ctrlnm"
            
    Set GetControlInfo = New Recordset
    GetControlInfo.Open strSQL, DBConn
End Function

Private Sub LoadLotNo()
    Dim Rs As Recordset
    
    Set Rs = GetLotNo(Trim(cboLotNo.Text))
    
    cboLotNo.Clear
    Do Until Rs.EOF
        cboLotNo.addItem Format(Rs.Fields("lotno").Value & "", "!" & String(100, "@")) & COL_DIV & _
                         Format(Rs.Fields("opendt").Value & "", "####-##-##") & COL_DIV & _
                         Format(Rs.Fields("expdt").Value & "", "####-##-##") & COL_DIV & _
                         Rs.Fields("makecd").Value & "" & COL_DIV & _
                         Rs.Fields("remark").Value & "" & COL_DIV & _
                         Rs.Fields("ctrlnm").Value & ""
        
        Rs.MoveNext
    Loop
    
    If cboLotNo.ListCount > 0 Then
        cboLotNo.ListIndex = 0
        
        lblOpenDt.Caption = medGetP(cboLotNo.List(0), 2, COL_DIV)
        lblExpDt.Caption = medGetP(cboLotNo.List(0), 3, COL_DIV)
        lblMakeCd.Caption = medGetP(cboLotNo.List(0), 4, COL_DIV)
        lblRemark.Caption = medGetP(cboLotNo.List(0), 5, COL_DIV)
    
        lblCtrlNm.Caption = medGetP(cboLotNo.List(0), 6, COL_DIV)
    End If
    
    Set Rs = Nothing
End Sub

Private Function GetLotNo(Optional ByVal pLotNo As String = "") As Recordset
    Dim strSQL As String
    
    strSQL = " select a.lotno,a.opendt,a.expdt,a.makecd,a.remark,b.ctrlnm from " & T_LAB023 & " a, " & T_LAB021 & " b " & _
            " where " & DBW("a.ctrlcd=", Trim(txtCtrlCd.Text)) & _
            " and " & DBW("a.levelcd=", IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H"))) & _
            " and " & DBW("a.expdt>=", Format(GetSystemDate, "yyyyMMdd")) & _
            " and a.ctrlcd=b.ctrlcd " & _
            " and a.levelcd=b.levelcd "
    
    If pLotNo <> "" Then
        strSQL = strSQL & " and " & DBW("lotno=", pLotNo)
    End If
    
    strSQL = strSQL & " order by opendt desc"
    
    Set GetLotNo = New Recordset
    GetLotNo.Open strSQL, DBConn
    
End Function

Private Sub LoadTestItem(Optional ByVal pReprint As Boolean = False)
    Dim objSQL As clsLISSqlQc
    Dim Rs As Recordset
    Dim strSQL As String
    Dim i As Long
    
    Set objSQL = New clsLISSqlQc
        
    If pReprint Then    'Reprint용 데이터 조회
'        strSQL = objSQL.SqlOrderQCItems(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), Trim(medGetP(cboLotno.Text, 1, COL_DIV)), medGetP(lblAccNo.Caption, 1, "-"), "20" & medGetP(lblAccNo.Caption, 2, "-"), medGetP(lblAccNo.Caption, 3, "-"))
        strSQL = " select a.testcd,b.testnm from " & T_LAB026 & " a, " & T_LAB001 & " b " & _
                " where " & DBW("a.workarea=", medGetP(lblAccNo.Caption, 1, "-")) & _
                " and " & DBW("a.accdt=", "20" & medGetP(lblAccNo.Caption, 2, "-")) & _
                " and " & DBW("a.accseq=", medGetP(lblAccNo.Caption, 3, "-")) & _
                " and a.testcd=b.testcd " & _
                " and b.applydt=(select max(applydt) from " & T_LAB001 & " where testcd = b.testcd) "
    Else    '처방용 데이터 조회
        strSQL = objSQL.SqlMstQCItems(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), Trim(medGetP(cboLotNo.Text, 1, COL_DIV)), "D")
    End If
    

    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    Call medClearTable(tblTestList)
    tblTestList.MaxRows = 16
    
    With tblTestList
        .ReDraw = False
        
        .Col = -1
        .Row = -1
        .BlockMode = True
        .CellType = CellTypeStaticText
        .BlockMode = False
        
        Do Until Rs.EOF
                If .DataRowCnt >= .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                
                i = i + 1
                
                .Row = (i + 3) \ 4
                .Col = (((i - 1) * 3) Mod 12) + 1
                .Row2 = .Row: .Col2 = .Col
                .BlockMode = True
                .CellType = CellTypeCheckBox
                .BlockMode = False
                .TypeHAlign = TypeHAlignCenter
                .Value = 1
                .Col = (((i - 1) * 3) Mod 12) + 2: .Value = Rs.Fields("testcd").Value & ""
                .Col = (((i - 1) * 3) Mod 12) + 3: .Value = Rs.Fields("testnm").Value & ""
            
            Rs.MoveNext
        Loop
        
        .ReDraw = True
    End With
    
    If tblTestList.DataRowCnt > 0 Then
        chkSelAll.Value = 1
        lblAllCnt.Caption = i
        lblSelCnt.Caption = i
        
        tblTestList.OperationMode = IIf(pReprint, OperationModeRead, OperationModeNormal)
        chkSelAll.Enabled = IIf(pReprint, False, True)
    End If
        
    Set Rs = Nothing
    Set objSQL = Nothing
End Sub

Private Sub cmdReprint_Click()
    MousePointer = vbHourglass
    Call DoRePrint
    MousePointer = vbDefault
End Sub

Private Sub DoRePrint()
    Dim lngCnt As Long
    Dim lngECnt As Long
    Dim lngSCnt As Long
    Dim i As Long
    
    '프로그래스 바
    
    lngCnt = 0
    Set objOrder = Nothing
    Set objOrder = New clsQcOrder
    
    If SetBarcodeInfo Then     '바코드 출력정보를 담는다
        objOrder.ColCount = 1  '출력할 건수
        If objOrder.PrintBarcodeLabel(1) = False Then
        
        End If
    Else
    
    End If
    
    Set objOrder = Nothing
End Sub

Private Function SetBarcodeInfo() As Boolean
'출력할 검사항목이 없는 경우에는 바코드 에러로 발생시켜야 한다.

    Dim lngCnt As Integer
      
    With objOrder
        .BarCount = 1
        .Controlcd = Trim(txtCtrlCd.Text)
        .ControlNm = Trim(txtCtrlCd.Text)
        .LevelCd = IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H"))
        .BuildNm = ObjSysInfo.BuildingNm
        .WorkArea = medGetP(lblAccNo.Caption, 1, "-")
        .AccDt = "20" & medGetP(lblAccNo.Caption, 2, "-")
        .AccSeq = medGetP(lblAccNo.Caption, 3, "-")
        .SpcYY = lblSpcYY.Caption
        .SpcNo = lblSpcNo.Caption
        .PtId = ""
'        .PtNm = Trim(Mid(lblEqp.Caption, 6))
        .PtNm = Trim(txtCtrlCd.Text)
        .EqpNm = Trim(Mid(lblEqp.Caption, 6))
'        .SpcNm = Trim(txtCtrlCd.Text)
        .SpcNm = IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H"))
        .StoreCd = ""
        .StatFg = ""
'        .WardId = IIf(optLevelcd(0).Value, "L", IIf(optLevelcd(1).Value, "N", "H"))
        .WardId = ""
        .OrdDt = lblColDt.Caption
        .ColTm = lblColTm.Caption
        .TestNames = Replace(objOrder.GetTestNames(.WorkArea, .AccDt, .AccSeq, lngCnt), vbTab, ",")
        
        If lngCnt = 0 Then GoTo Nodata  '출력할 검사항목이 없는 경우 에러처리
                
        '바코드 출력정보를 담아주는 메소드
        Call objOrder.PrintBarcode(1, String(11, " ") & .ColTm)
    End With
    
    SetBarcodeInfo = True
    
    Exit Function
    
Nodata:
    SetBarcodeInfo = False
End Function

Private Sub cmdReprintList_Click()
    Dim objRep As clsQcOrder
    Dim objPop As clsPopUpList
    Dim Rs As Recordset
    Dim strSQL As String
    
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    If cboLotNo.Text = "" Then Exit Sub
    
    Set objRep = New clsQcOrder
    
'    Set Rs = objRep.GetLabNumbers(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), "", _
'                                    Trim(Mid(lblWorkArea.Caption, 1, 5)), _
'                                    Format(DateAdd("d", -7, GetSystemDate), CS_DateDbFormat), _
'                                    Format(GetSystemDate, CS_DateDbFormat))
    
    strSQL = " select distinct a.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_SUBSTR & "( a.accdt,3)" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_CONVERT("CHAR", "a.accseq") & " as accno, " & _
            " b.spcyy" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & FUNC_CONVERT("CHAR", "b.spcno") & " as barno,b.coldt ,b.coltm from " & T_LAB026 & " a," & T_LAB201 & " b " & _
            " where " & DBW("a.ctrlcd=", Trim(txtCtrlCd.Text)) & _
            " and " & DBW("a.levelcd=", IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H"))) & _
            " and " & DBW("a.lotno=", Trim(medGetP(cboLotNo.Text, 1, COL_DIV))) & _
            " and " & DBW("b.coldt   >= ", Format(DateAdd("d", -7, GetSystemDate), CS_DateDbFormat)) & _
            " and " & DBW("b.coldt   <= ", Format(GetSystemDate, CS_DateDbFormat)) & _
            " and a.workarea=b.workarea " & _
            " and a.accdt=b.accdt " & _
            " and a.accseq=b.accseq " & _
            " and " & DBW("b.qcfg", "1", 2) & _
            " order by b.coldt,b.coltm "
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    If Rs.EOF Then
        MsgBox "재발행 대상이 없습니다." & vbNewLine & "7일 이전의 재발행 대상을 보시려면 QC 일괄처방화면을 사용하십시오.", vbExclamation
        GoTo Nodata
    End If
    
    Set objPop = New clsPopUpList
    
    With objPop
        .Recordset = Rs
        .Delimiter = COL_DIV
        .FormCaption = "재발행 대상 리스트"
        .ColumnHeaderText = "접수번호바코드번호채취일자채취시간"
        .ColumnHeaderWidth = "1214.9291319.8111124.7871110.047"
        .FormWidth = 5220
        .LoadPopUp
        
        lblAccNo.Caption = medGetP(.SelectedString, 1, .Delimiter)
        lblSpcYY.Caption = medGetP(medGetP(.SelectedString, 2, .Delimiter), 1, "-")
        lblSpcNo.Caption = medGetP(medGetP(.SelectedString, 2, .Delimiter), 2, "-")
        lblBarNo.Caption = lblSpcYY.Caption & Format(lblSpcNo.Caption, LIS_BarFormat)
        lblColDt.Caption = medGetP(.SelectedString, 3, .Delimiter)
        lblColTm.Caption = medGetP(.SelectedString, 4, .Delimiter)
    End With
    
    Call LoadTestItem(True)
    
    If tblTestList.DataRowCnt > 0 Then
        cmdReprint.Enabled = True
        cmdSave.Enabled = False
    End If
    
Nodata:
    Set Rs = Nothing
    Set objPop = Nothing
    Set objRep = Nothing
End Sub

Private Sub cmdSave_Click()
    If CheckValidation = False Then Exit Sub
    
    Call DoCollection
'---------------------------
' 2009.03.12 양성현 추가
        Call AfterSave
'---------------------------

End Sub

Private Function CheckValidation() As Boolean
    CheckValidation = False
    
    If Trim(txtCtrlCd.Text) = "" Then
        MsgBox "컨트롤 코드를 입력하거나 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    If optLevelCd(0).Value = False And optLevelCd(1).Value = False And optLevelCd(2).Value = False Then
        MsgBox "컨트롤 레벨을 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    If Trim(cboLotNo.Text) = "" Then
        MsgBox "LotNo를 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    If tblTestList.DataRowCnt = 0 Then
        MsgBox "등록된 컨트롤이 아닙니다. 컨트롤을 먼저 등록하십시오.", vbExclamation
        Exit Function
    End If
    
    If Val(lblSelCnt.Caption) = 0 Then
        MsgBox "선택된 항목이 없습니다.", vbExclamation
        Exit Function
    End If
    
    CheckValidation = True
End Function

Private Sub DoCollection()
    MousePointer = vbHourglass
    
    Set objQC = Nothing
    Set objOrder = Nothing
    
    Set objQC = New clsQcMst
    Set objOrder = New clsQcOrder

    If ReadyToCollect Then
        MsgBox "정상적으로 처리되었습니다.", vbInformation
    Else
        MsgBox "처리 도중 오류가 발생하였습니다.", vbExclamation
    End If
    
    Set objQC = Nothing
    Set objOrder = Nothing
    
    MousePointer = vbDefault
End Sub

Private Function ReadyToCollect() As Boolean

    Dim i As Long
    Dim j As Long
    Dim k As Long
    
'이렇게 밖에 못하나..불러올거 다 불러오구.. 채혈할때 또불러오네.. ㅡㅡ;
'나중에 연구좀 더해서 바까야 겠다..

    Call objQC.GetQcData(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), Trim(medGetP(cboLotNo.Text, 1, COL_DIV)))
    Call objQC.GetQCItems(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), Trim(medGetP(cboLotNo.Text, 1, COL_DIV)))
    
    '검사항목이 없는 경우 에러로 간주
    If objQC.ItemCount = 0 Then
        ReadyToCollect = False
        Exit Function
    End If
    
    For i = 1 To objQC.ItemCount
        For j = 1 To tblTestList.DataRowCnt
            tblTestList.Row = j
            For k = 1 To tblTestList.DataColCnt Step 3
                tblTestList.Col = k + 1
                If tblTestList.Value = objQC.Item(i).TestCd Then
                    tblTestList.Col = k
                    objQC.Item(i).Selected = IIf(tblTestList.Value = "1", True, False)
                End If
            Next k
        Next j
    Next i
       
   
    With objOrder
        Set .MyQc = objQC
        
        .SpcYY = LIS_BarDiv & Mid(Format(GetSystemDate, "YYYY"), 4) '검체년도
        
        .PtId = 0                                     '환자ID
        .PtNm = ""
        .Sex = ""                                     '성별
        .AgeDay = 0                                   '환자일령
        .BedInDt = ""                                 '입원일
        .OrdDt = Format(GetSystemDate, CS_DateDbFormat)    '처방일
        
        .Controlcd = Trim(txtCtrlCd.Text)                        'Control코드
        .ControlNm = lblCtrlNm.Caption                         'Control명
        .EqpNm = Trim(Mid(lblEqp.Caption, 6))                            '장비명
        .LevelCd = IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H"))                         'Level코드
        .Lotno = Trim(medGetP(cboLotNo.Text, 1, COL_DIV))                              'Lot Number
        .WardId = ""                                  '병동ID
        .EntDt = Format(GetSystemDate, CS_DateDbFormat)          '입력일
        .DeptCd = ""
        .BuildCd = ObjSysInfo.BuildingCd
        .SpcCd = IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H"))
        .MultiFg = ""
        .QcFg = "1"                                   '내부정도관리
        
        .EntTm = Format(GetSystemDate, CS_TimeDbFormat)          '입력시간
        .EntId = ObjSysInfo.EmpId                         '입력자
        .OrgAccNo = ""                                '원접수번호
        .HosilId = ""                                 '병실ID
        .RoomId = ""                                  '병실ID
        .BedId = ""                                   '침상ID
        .ColDt = Format(GetSystemDate, CS_DateDbFormat)          '채혈일
        .ColId = ObjSysInfo.EmpId                         '채혈자
        .OrgBuildCd = ObjSysInfo.BuildingCd      '** 채혈이 수행되는 건물코드
        .WorkArea = Trim(Mid(lblWorkarea.Caption, 1, 5))
        
        If .DoCollection Then
            '채혈하면서 바로 바코드 발행
            If .PrintBarcodeLabel(1) = False Then      '바코드 출력도중 에러난 경우 채혈은 정상적으로 하고 메시지만 띄워준다.
            
            End If
            
            ReadyToCollect = True
        Else
            ReadyToCollect = False
        End If
    End With
End Function

Private Sub Form_Load()
    txtCtrlCd.Text = ""
    Call InitControl
    cboLotNo.Clear
    Call InitLotNo
    
    Call medClearTable(tblTestList)
    tblTestList.MaxRows = 16
    tblTestList.Col = -1
    tblTestList.Row = -1
    tblTestList.BlockMode = True
    tblTestList.CellType = CellTypeStaticText
    tblTestList.BlockMode = False
End Sub

Private Sub InitControl()
    lblCtrlNm.Caption = ""
    lblCtrlDiv.Caption = ""
    lblEqp.Caption = ""
    lblBuilding.Caption = ""
    lblSection.Caption = ""
    lblWorkarea.Caption = ""
End Sub

Private Sub InitLotNo()
    lblOpenDt.Caption = ""
    lblExpDt.Caption = ""
    lblMakeCd.Caption = ""
    lblRemark.Caption = ""
    
    chkSelAll.Value = 0
    
    lblAllCnt.Caption = ""
    lblSelCnt.Caption = ""
    
    lblAccNo.Caption = ""
    lblBarNo.Caption = ""
    lblColDt.Caption = ""
    lblColTm.Caption = ""
    lblSpcYY.Caption = ""
    lblSpcNo.Caption = ""
    cmdReprint.Enabled = False
    cmdSave.Enabled = True
    chkSelAll.Enabled = True
End Sub

Private Sub optLevelcd_Click(Index As Integer)
    On Error Resume Next
    If Screen.ActiveControl.Name <> optLevelCd(Index).Name Then Exit Sub
    
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    
    cboLotNo.Clear
    Call InitLotNo
    
    Call medClearTable(tblTestList)
    tblTestList.MaxRows = 16
    tblTestList.Col = -1
    tblTestList.Row = -1
    tblTestList.BlockMode = True
    tblTestList.CellType = CellTypeStaticText
    tblTestList.BlockMode = False
    
    Call LoadLotNo
    Call LoadTestItem
    
    If tblTestList.DataRowCnt = 0 Then
        MsgBox "처방 가능한 컨트롤이 존재하지 않습니다.", vbExclamation
    End If
End Sub

Private Sub tblTestList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    On Error Resume Next
    If Screen.ActiveControl.Name <> tblTestList.Name Then Exit Sub
    
    If tblTestList.DataRowCnt = 0 Then Exit Sub
    If (Col Mod 3) <> 1 Then Exit Sub
    
    With tblTestList
        .Row = Row
        .Col = Col
        
        If .Value = 1 Then
            lblSelCnt.Caption = Val(lblSelCnt.Caption) + 1
        Else
            lblSelCnt.Caption = Val(lblSelCnt.Caption) - 1
        End If
    End With
    
    chkSelAll.Value = IIf(lblSelCnt.Caption = lblAllCnt.Caption, 1, 0)
End Sub

Private Sub tblTestList_Click(ByVal Col As Long, ByVal Row As Long)
'    If (Col Mod 3) <> 1 Then Exit Sub
'
'    With tblTestList
'        .Row = Row
'        .Col = Col
'
'        If .Value = 1 Then
'            lblSelCnt.Caption = lblSelCnt.Caption + 1
'        Else
'            lblSelCnt.Caption = Val(lblSelCnt.Caption) - 1
'        End If
'    End With
End Sub

Private Sub txtCtrlCd_Change()
    On Error Resume Next
    If Screen.ActiveControl.Name <> txtCtrlCd.Name Then Exit Sub
    
    If lblCtrlNm.Caption <> "" Then
        Call InitControl
        cboLotNo.Clear
        Call InitLotNo
        
        Call medClearTable(tblTestList)
        tblTestList.MaxRows = 16
        tblTestList.Col = -1
        tblTestList.Row = -1
        tblTestList.BlockMode = True
        tblTestList.CellType = CellTypeStaticText
        tblTestList.BlockMode = False
    End If
End Sub

Private Sub txtCtrlCd_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtCtrlCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtCtrlCd_LostFocus()
    Dim Rs As Recordset
'이따구루 밖에 못할까? 나중에 다른 방법으로 고쳐야지...

    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    If Trim(lblCtrlNm.Caption) <> "" Then Exit Sub
    
    DoEvents
    Set Rs = GetControlInfo(Trim(txtCtrlCd.Text))
    
    If Rs.EOF = False Then
        DoEvents
        Call LoadControlInfo(Trim(txtCtrlCd.Text))
        DoEvents
        Call LoadLotNo
        DoEvents
        Call LoadTestItem
    End If
    
    Set Rs = Nothing
End Sub

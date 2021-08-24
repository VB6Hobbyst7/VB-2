VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "SPIN32.OCX"
Begin VB.Form clp032 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "검사결과관리"
   ClientHeight    =   8040
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   11865
   ControlBox      =   0   'False
   Icon            =   "Clp032.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8040
   ScaleWidth      =   11865
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel2 
      Height          =   510
      Left            =   90
      TabIndex        =   38
      Top             =   450
      Width           =   4920
      _Version        =   65536
      _ExtentX        =   8678
      _ExtentY        =   900
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cmbSLip 
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
         Left            =   1440
         Style           =   2  '드롭다운 목록
         TabIndex        =   39
         Top             =   90
         Width           =   2715
      End
      Begin VB.Label Label9 
         Caption         =   "검사종목변경 :"
         Height          =   240
         Left            =   135
         TabIndex        =   40
         Top             =   135
         Width           =   1275
      End
   End
   Begin FPSpreadADO.fpSpread ssResult 
      Height          =   6675
      Left            =   1845
      TabIndex        =   2
      Top             =   1080
      Width           =   9930
      _Version        =   196608
      _ExtentX        =   17515
      _ExtentY        =   11774
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      ColsFrozen      =   7
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
      MaxCols         =   22
      MaxRows         =   600
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "Clp032.frx":000C
      UserResize      =   0
      VisibleCols     =   22
      VisibleRows     =   500
      Appearance      =   1
   End
   Begin Threed.SSPanel panelExamName 
      Height          =   345
      Left            =   75
      TabIndex        =   35
      Top             =   60
      Width           =   2715
      _Version        =   65536
      _ExtentX        =   4789
      _ExtentY        =   609
      _StockProps     =   15
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   4530
      Left            =   60
      TabIndex        =   22
      Top             =   1095
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
      _ExtentY        =   7990
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
      BorderWidth     =   1
      BevelInner      =   2
      Begin Threed.SSOption optEr 
         Height          =   285
         Left            =   180
         TabIndex        =   45
         Top             =   2115
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "응급Order"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption ssAbnormal 
         Height          =   285
         Left            =   180
         TabIndex        =   44
         Top             =   2340
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Abnormal"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Spin.SpinButton SpinButton2 
         Height          =   300
         Left            =   1275
         TabIndex        =   37
         Top             =   1740
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   529
         _StockProps     =   73
         BackColor       =   -2147483633
         BorderThickness =   0
         ShadowThickness =   1
         TdThickness     =   1
      End
      Begin Spin.SpinButton SpinButton1 
         Height          =   300
         Left            =   1275
         TabIndex        =   36
         Top             =   1425
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   529
         _StockProps     =   73
         BackColor       =   -2147483633
         BorderThickness =   0
         ShadowThickness =   1
         TdThickness     =   1
      End
      Begin VB.TextBox txtStartSlipno 
         BackColor       =   &H00E1FAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1425
         Width           =   1125
      End
      Begin VB.TextBox txtEndSlipno 
         BackColor       =   &H00E1FAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1740
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker dtToJeobsu 
         Height          =   315
         Left            =   135
         TabIndex        =   23
         Top             =   690
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36306
      End
      Begin MSComCtl2.DTPicker dtFromJeobsu 
         Height          =   315
         Left            =   135
         TabIndex        =   24
         Top             =   330
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36306
      End
      Begin Threed.SSPanel panCondition 
         Height          =   1260
         Left            =   420
         TabIndex        =   25
         Top             =   3090
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   2222
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
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         Begin Threed.SSCheck chkRepCd 
            Height          =   255
            Left            =   105
            TabIndex        =   26
            Top             =   705
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "외부의뢰"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.01
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkUnknow 
            Height          =   345
            Left            =   105
            TabIndex        =   27
            Top             =   885
            Visible         =   0   'False
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   609
            _StockProps     =   78
            Caption         =   "미확인"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.01
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkJeobsu 
            Height          =   210
            Left            =   105
            TabIndex        =   28
            Top             =   60
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
            _ExtentY        =   370
            _StockProps     =   78
            Caption         =   "접수중"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   8.99
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkComplete 
            Height          =   255
            Left            =   105
            TabIndex        =   29
            Top             =   480
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "결과완료"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.01
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkPart 
            Height          =   210
            Left            =   105
            TabIndex        =   30
            Top             =   270
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   370
            _StockProps     =   78
            Caption         =   "부분결과"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   8.99
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSOption optCondition 
         Height          =   210
         Left            =   180
         TabIndex        =   31
         Top             =   2850
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   370
         _StockProps     =   78
         Caption         =   "조건별"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   8.99
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optAll 
         Height          =   300
         Left            =   180
         TabIndex        =   32
         Top             =   2565
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "전체"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "[접수기간]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   150
         TabIndex        =   34
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "[접수번호조건]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   150
         TabIndex        =   33
         Top             =   1200
         Width           =   1260
      End
   End
   Begin Threed.SSPanel panelTitle 
      Height          =   1035
      Left            =   6105
      TabIndex        =   4
      Top             =   45
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   1826
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
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.Label lblDrname 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '단일 고정
         Caption         =   "Dr"
         Height          =   255
         Left            =   2820
         TabIndex        =   21
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblDeptname 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '단일 고정
         Caption         =   "Dept"
         Height          =   255
         Left            =   2205
         TabIndex        =   20
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblAge 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '단일 고정
         Caption         =   "Age"
         Height          =   255
         Left            =   1590
         TabIndex        =   19
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblSex 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '단일 고정
         Caption         =   "Sex"
         Height          =   255
         Left            =   975
         TabIndex        =   18
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblRoom 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '단일 고정
         Caption         =   "Room"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblPtNo 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3900
         TabIndex        =   16
         Top             =   90
         Width           =   1425
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3900
         TabIndex        =   15
         Top             =   370
         Width           =   1425
      End
      Begin VB.Label lblGeomsaJa 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3900
         TabIndex        =   14
         Top             =   660
         Width           =   1425
      End
      Begin VB.Label lblJeobsuDt 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1020
         TabIndex        =   13
         Top             =   90
         Width           =   1425
      End
      Begin VB.Label lblSlipNo 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1020
         TabIndex        =   12
         Top             =   370
         Width           =   1425
      End
      Begin VB.Label lblGeomsaDt 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1020
         TabIndex        =   11
         Top             =   660
         Width           =   1425
      End
      Begin VB.Label Label8 
         Caption         =   "검 사 자"
         Height          =   195
         Left            =   3060
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "환 자 명"
         Height          =   195
         Left            =   3060
         TabIndex        =   9
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "환자번호"
         Height          =   195
         Left            =   3060
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "검사일자"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "접수번호"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "접수일자"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
   End
   Begin Threed.SSPanel ssGeomSa 
      Height          =   345
      Left            =   2835
      TabIndex        =   3
      Top             =   60
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "검사결과Report"
      ForeColor       =   8388608
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12.01
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      RoundedCorners  =   0   'False
   End
   Begin MSForms.CommandButton cmdPr 
      Height          =   465
      Left            =   90
      TabIndex        =   43
      Top             =   6795
      Width           =   1725
      Caption         =   "결과출력"
      PicturePosition =   327683
      Size            =   "3043;820"
      Picture         =   "Clp032.frx":2B3E
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdALLcheck 
      Height          =   465
      Left            =   90
      TabIndex        =   42
      Top             =   6345
      Width           =   1725
      Caption         =   "CheckAll"
      PicturePosition =   327683
      Size            =   "3043;820"
      Picture         =   "Clp032.frx":3418
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdInquiry 
      Height          =   465
      Left            =   90
      TabIndex        =   41
      Top             =   5895
      Width           =   1725
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "3043;820"
      Picture         =   "Clp032.frx":3CF2
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuexit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu mnuJob 
      Caption         =   "작업"
      Begin VB.Menu mnuQry 
         Caption         =   "조회확인"
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "행 선택"
         Begin VB.Menu mnuSel 
            Caption         =   "모든행 선택확인"
         End
         Begin VB.Menu mnuDesel 
            Caption         =   "모든행 선택제거"
         End
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAppend 
         Caption         =   "결과등록"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "결과출력"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuPerson 
      Caption         =   "개별조회"
      Begin VB.Menu mnuPtno 
         Caption         =   "등록번호"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSname 
         Caption         =   "환자명"
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "clp032"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LsExamGu(0 To 50)        As String * 2

Public Function Get_General_Status(ByVal sJeobsuDt As String, ByVal iSLno1 As Integer, ByVal iSLno2 As Integer) As String
    Dim adoStatus       As ADODB.Recordset
    
    StrSql = ""
    StrSql = StrSql & " SELECT Status"
    StrSql = StrSql & " FROM   TWEXAM_General"
    StrSql = StrSql & " WHERE  JEOBSUDT  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    SLipno1   = " & iSLno1
    StrSql = StrSql & " AND    SLipno2   = " & iSLno2
    
    If False = adoSetOpen(StrSql, adoStatus) Then
        Get_General_Status = ""
        Exit Function
    End If
    
    Select Case adoStatus.Fields("Status").Value & ""
        Case "R": Get_General_Status = "접수중"
        Case "C":
            If iSLno1 = 42 Then
                Get_General_Status = "최종보고"
            Else
                Get_General_Status = "결과완료"
            End If
        Case "P":
            If iSLno1 = 42 Then
                Get_General_Status = "예비보고"
            Else
                Get_General_Status = "부분결과확인"
            End If
        Case "U": Get_General_Status = "미확인"
        Case "X": Get_General_Status = "Data이상(Panic Or Delta)"
        Case Else: Get_General_Status = ""
    End Select
    
    Call adoSetClose(adoStatus)
    
    
End Function

Public Function Printing_Sens(ByVal sJeobsuDt As String, ByVal iSLno1 As Integer, ByVal iSLno2 As Integer, ByVal sItemCd As String) As Integer
    Dim adoSensRet      As ADODB.Recordset
    Dim iSensCls        As Integer
    
    '420401 = Drug Senstivity (세균검사)
    
    Call SensResultClear
    
    StrSql = ""
    StrSql = StrSql & "  SELECT  a.ItemCd, c.Org_name, d.Codenm AntiName, a.Sens, a.Value"
    StrSql = StrSql & "  FROM    TWEXAM_SENS        a,"
    StrSql = StrSql & "          TWEXAM_GENERAL_Sub b,"
    StrSql = StrSql & "          TWEXAM_ORGLIST     c,"
    StrSql = StrSql & "          TWEXAM_AntiList    d"
    StrSql = StrSql & "  WHERE   a.JEOBSUDT  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & "  AND     a.SLipno1   = " & iSLno1
    StrSql = StrSql & "  AND     a.SLipno2   = " & iSLno2
    StrSql = StrSql & "  AND     a.ItemCd    = '" & sItemCd & "'"
    StrSql = StrSql & "  AND     a.JeobsuDt = b.JeobsuDt(+)"
    StrSql = StrSql & "  AND     a.SLipno1  = b.SLipno1(+)"
    StrSql = StrSql & "  AND     a.SLipno2  = b.SLipno2(+) "
    StrSql = StrSql & "  AND     a.ORACOD   = b.Rcode1"
    StrSql = StrSql & "  AND     a.Oracod   = c.Org_code(+)"
    StrSql = StrSql & "  AND     a.Yakcod   = d.Codeky(+)"
    
    StrSql = StrSql & " UNION ALL "
    StrSql = StrSql & "  SELECT  a.ItemCd, c.Org_name, d.Codenm AntiName, a.Sens, a.Value"
    StrSql = StrSql & "  FROM    TWEXAM_SENS        a,"
    StrSql = StrSql & "          TWEXAM_GENERAL_Sub b,"
    StrSql = StrSql & "          TWEXAM_ORGLIST     c,"
    StrSql = StrSql & "          TWEXAM_AntiList    d"
    StrSql = StrSql & "  WHERE   a.JEOBSUDT  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & "  AND     a.SLipno1   = " & iSLno1
    StrSql = StrSql & "  AND     a.SLipno2   = " & iSLno2
    StrSql = StrSql & "  AND     a.ItemCd    = '" & sItemCd & "'"
    StrSql = StrSql & "  AND     a.JeobsuDt = b.JeobsuDt(+)"
    StrSql = StrSql & "  AND     a.SLipno1  = b.SLipno1(+)"
    StrSql = StrSql & "  AND     a.SLipno2  = b.SLipno2(+) "
    StrSql = StrSql & "  AND     a.ORACOD   = b.Rcode2"
    StrSql = StrSql & "  AND     a.Oracod   = c.Org_code(+)"
    StrSql = StrSql & "  AND     a.Yakcod   = d.Codeky(+)"
    
    StrSql = StrSql & " UNION ALL "
    StrSql = StrSql & "  SELECT  a.ItemCd, c.Org_name, d.Codenm AntiName, a.Sens, a.Value"
    StrSql = StrSql & "  FROM    TWEXAM_SENS        a,"
    StrSql = StrSql & "          TWEXAM_GENERAL_Sub b,"
    StrSql = StrSql & "          TWEXAM_ORGLIST     c,"
    StrSql = StrSql & "          TWEXAM_AntiList    d"
    StrSql = StrSql & "  WHERE   a.JEOBSUDT  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & "  AND     a.SLipno1   = " & iSLno1
    StrSql = StrSql & "  AND     a.SLipno2   = " & iSLno2
    StrSql = StrSql & "  AND     a.ItemCd    = '" & sItemCd & "'"
    StrSql = StrSql & "  AND     a.JeobsuDt = b.JeobsuDt(+)"
    StrSql = StrSql & "  AND     a.SLipno1  = b.SLipno1(+)"
    StrSql = StrSql & "  AND     a.SLipno2  = b.SLipno2(+) "
    StrSql = StrSql & "  AND     a.ORACOD   = b.Rcode3"
    StrSql = StrSql & "  AND     a.Oracod   = c.Org_code(+)"
    StrSql = StrSql & "  AND     a.Yakcod   = d.Codeky(+)"
    
    StrSql = StrSql & " UNION ALL "
    StrSql = StrSql & "  SELECT  a.ItemCd, c.Org_name, d.Codenm AntiName, a.Sens, a.Value"
    StrSql = StrSql & "  FROM    TWEXAM_SENS        a,"
    StrSql = StrSql & "          TWEXAM_GENERAL_Sub b,"
    StrSql = StrSql & "          TWEXAM_ORGLIST     c,"
    StrSql = StrSql & "          TWEXAM_AntiList    d"
    StrSql = StrSql & "  WHERE   a.JEOBSUDT  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & "  AND     a.SLipno1   = " & iSLno1
    StrSql = StrSql & "  AND     a.SLipno2   = " & iSLno2
    StrSql = StrSql & "  AND     a.ItemCd    = '" & sItemCd & "'"
    StrSql = StrSql & "  AND     a.JeobsuDt = b.JeobsuDt(+)"
    StrSql = StrSql & "  AND     a.SLipno1  = b.SLipno1(+)"
    StrSql = StrSql & "  AND     a.SLipno2  = b.SLipno2(+) "
    StrSql = StrSql & "  AND     a.ORACOD   = b.Rcode4"
    StrSql = StrSql & "  AND     a.Oracod   = c.Org_code(+)"
    StrSql = StrSql & "  AND     a.Yakcod   = d.Codeky(+)"
    
    StrSql = StrSql & " UNION ALL "
    StrSql = StrSql & "  SELECT  a.ItemCd, c.Org_name, d.Codenm AntiName, a.Sens, a.Value"
    StrSql = StrSql & "  FROM    TWEXAM_SENS        a,"
    StrSql = StrSql & "          TWEXAM_GENERAL_Sub b,"
    StrSql = StrSql & "          TWEXAM_ORGLIST     c,"
    StrSql = StrSql & "          TWEXAM_AntiList    d"
    StrSql = StrSql & "  WHERE   a.JEOBSUDT  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & "  AND     a.SLipno1   = " & iSLno1
    StrSql = StrSql & "  AND     a.SLipno2   = " & iSLno2
    StrSql = StrSql & "  AND     a.ItemCd    = '" & sItemCd & "'"
    StrSql = StrSql & "  AND     a.JeobsuDt = b.JeobsuDt(+)"
    StrSql = StrSql & "  AND     a.SLipno1  = b.SLipno1(+)"
    StrSql = StrSql & "  AND     a.SLipno2  = b.SLipno2(+) "
    StrSql = StrSql & "  AND     a.ORACOD   = b.Rcode5"
    StrSql = StrSql & "  AND     a.Oracod   = c.Org_code(+)"
    StrSql = StrSql & "  AND     a.Yakcod   = d.Codeky(+)"
    
    If False = adoSetOpen(StrSql, adoSensRet) Then Exit Function
    
    For iSensCls = 0 To adoSensRet.RecordCount - 1
        SensResult.ItemCd(iSensCls) = adoSensRet.Fields("ItemCd").Value & ""
        SensResult.Rcode(iSensCls) = adoSensRet.Fields("Org_name").Value & ""
        SensResult.AntiName(iSensCls) = adoSensRet.Fields("AntiName").Value & ""
        SensResult.Sens(iSensCls) = adoSensRet.Fields("Sens").Value & ""
        SensResult.Value(iSensCls) = adoSensRet.Fields("Value").Value & ""
        adoSensRet.MoveNext
    Next
    Printing_Sens = adoSensRet.RecordCount
    Call adoSetClose(adoSensRet)
    
End Function


Private Sub cmbSLip_Click()
    
    
    If cmbSLip.ListIndex = -1 Then Exit Sub
    
    GiExamNumb = Left(cmbSLip.List(cmbSLip.ListIndex), 2)
    GsExamJong = Mid(cmbSLip.Text, 5, Len(cmbSLip.Text) - 4)
    panelExamName.Caption = GsExamJong
    
    Call SpreadSetClear(ssResult)
    
    ssResult.Col = 1
    ssResult.Row = 0
    ssResult.Text = "C"
    
    lblJeobsuDt.Caption = ""
    lblSlipNo.Caption = ""
    lblGeomsaDt.Caption = ""
    lblPtNo.Caption = ""
    lblName.Caption = ""
    lblGeomsaJa.Caption = ""
    
    
End Sub

Private Sub cmdALLcheck_Click()
    
    Call ssResult_Click(1, 0)
    
End Sub




Private Sub cmdInquiry_Click()
    
    Dim LsPtno       As String * 8
    Dim LsStatus     As String * 1
    Dim LsCodeKy     As String
    Dim LsDrCode     As String * 6
    Dim LsDeptCode   As String * 4
    Dim LiReccnt     As Integer
    Dim i            As Integer
    Dim LsRet
    Dim sFromDate    As String
    Dim sToDate      As String
    
    
    Call SSInitialize(ssResult)
    sFromDate = Format(dtFromJeobsu.Value, "yyyy-MM-dd")
    sToDate = Format(dtToJeobsu.Value, "yyyy-MM-dd")
'---------------------------------------------'
'   검사접수자 DB READ                          '
'---------------------------------------------'
    
    gStrSql = ""
    gStrSql = gStrSql & " SELECT /*+ INDEX (TWBas_Patient INDEX_PATIENT0) */"
    gStrSql = gStrSql & "        a.*, a.RowID,"
    gStrSql = gStrSql & "        TO_CHAR(a.JeobsuDT, 'YYYY-MM-DD') JeobsuDt, "
    gStrSql = gStrSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') OrderDt,  "
    gStrSql = gStrSql & "        TO_CHAR(a.GeomsaDt, 'YYYY-MM-DD') GeomsaDt, "
    gStrSql = gStrSql & "        b.Sname,     b.Sex,  c.Deptnamek, d.Drname, e.Name, f.Codenm"
    gStrSql = gStrSql & " FROM   TWEXAM_GENERAL  a, "
    gStrSql = gStrSql & "        TWBAS_PATIENT   b, "
    gStrSql = gStrSql & "        TWBAS_DEPT      c, "
    gStrSql = gStrSql & "        TWBAS_DOCTOR    d, "
    gStrSql = gStrSql & "        TWBAS_PASS      e, "
    gStrSql = gStrSql & "        TWEXAM_Sample   f  "
    gStrSql = gStrSql & " WHERE  a.JeobsuDt >= TO_DATE('" & sFromDate & "','YYYY-MM-DD')"
    gStrSql = gStrSql & " AND    a.JeobsuDt <= TO_DATE('" & sToDate & "',  'YYYY-MM-DD')"
    gStrSql = gStrSql & " AND    a.Slipno1   = " & GiExamNumb
    gStrSql = gStrSql & " AND    a.GbCh      = 'Y'"
    gStrSql = gStrSql & " AND    a.Slipno2  >= " & Val(txtStartSlipno.Text)
    gStrSql = gStrSql & " AND    a.Ptno      = b.Ptno(+)"
    gStrSql = gStrSql & " AND    a.DeptCode  = c.DeptCode(+)"
    gStrSql = gStrSql & " AND    a.DrCode    = d.DrCode(+)"
    gStrSql = gStrSql & " AND    a.GeomchCD  = f.Code(+)"
    gStrSql = gStrSql & " AND    a.Geomsaja  = e.Idnumber(+)"
    gStrSql = gStrSql & " AND   (e.Programid = ' ' OR e.Programid IS NULL)"
    
    If Not optAll = True Then
        If Val(txtEndSlipno.Text) > 0 Then
            gStrSql = gStrSql & "  AND   a.Slipno2  <= " & Val(txtEndSlipno.Text)
        End If
    End If
    
    '전체 검색
    If optAll = True Then
        gStrSql = gStrSql & "  ORDER BY a.JeobsuDt, a.SlipNo1, a.SlipNo2, a.PtNo  ASC "
    End If
    
    'Panic OR Delta Data 검색
    If ssAbnormal = True Then
        gStrSql = gStrSql & " AND  a.Status = 'X'"
        gStrSql = gStrSql & " ORDER BY a.JeobsuDt, a.SlipNo1, a.SlipNo2, a.PtNo  ASC "
    End If
        
    If optEr = True Then
        gStrSql = gStrSql & " AND  (a.DeptCode = 'ER' OR a.GbER = 'E')"
        gStrSql = gStrSql & " AND   a.Status = 'R'"
    End If
    
        
    ' 조건별 검색 문장
    If optCondition = True Then
        If chkJeobsu = True And chkPart = True And chkComplete = True And chkUnknow = True Then
            gStrSql = gStrSql & " AND  ( a.Status   = 'R' OR a.Status = 'P' OR a.Status = 'C' OR a.Status = 'U' ) "
        ElseIf chkPart = True And chkComplete = True And chkUnknow Then
            gStrSql = gStrSql & " AND  ( a.Status   = 'P' OR a.Status = 'C' OR a.Status = 'U' ) "
        ElseIf chkJeobsu = True And chkComplete = True And chkUnknow Then
            gStrSql = gStrSql & " AND  ( a.Status   = 'R' OR a.Status = 'C' OR a.Status = 'U' ) "
        ElseIf chkJeobsu = True And chkPart = True And chkUnknow Then
            gStrSql = gStrSql & " AND  ( a.Status   = 'R' OR a.Status = 'P' OR a.Status = 'U' ) "
        ElseIf chkJeobsu = True And chkPart = True And chkComplete Then
            gStrSql = gStrSql & " AND  ( a.Status   = 'R' OR a.Status = 'P' OR a.Status = 'C' ) "
        ElseIf chkComplete = True And chkUnknow = True Then
            gStrSql = gStrSql & " AND  ( a.Status   = 'C' OR a.Status = 'U' ) "
        ElseIf chkPart = True And chkUnknow = True Then
            gStrSql = gStrSql & " AND  ( a.Status   = 'P' OR a.Status = 'U' ) "
        ElseIf chkPart = True And chkComplete = True Then
            gStrSql = gStrSql & " AND  ( a.Status   = 'P' OR a.Status = 'C' ) "
        ElseIf chkJeobsu = True And chkUnknow = True Then
            gStrSql = gStrSql & " AND  ( a.Status   = 'R' OR a.Status = 'U' ) "
        ElseIf chkJeobsu = True And chkComplete = True Then
            gStrSql = gStrSql & " AND  ( a.Status   = 'R' OR a.Status = 'C' ) "
        ElseIf chkJeobsu = True And chkPart = True Then
            gStrSql = gStrSql & " AND  ( a.Status   = 'R' OR a.Status = 'P' ) "
        ElseIf chkUnknow = True Then
            gStrSql = gStrSql & " AND  a.Status   = 'U' "
        ElseIf chkComplete = True Then
            gStrSql = gStrSql & " AND  a.Status   = 'C' "
        ElseIf chkPart = True Then
            gStrSql = gStrSql & " AND  a.Status   = 'P' "
        ElseIf chkJeobsu = True Then
            gStrSql = gStrSql & " AND  a.Status   = 'R' "
        End If
        
        If chkRepCd = True Then
            gStrSql = gStrSql & " AND  a.Reporcd   = 'W' "
        End If
        
        gStrSql = gStrSql & "  ORDER BY a.STATUS, a.JeobsuDt, a.SlipNo1, a.STATUS, a.SlipNo2, a.PtNo  ASC             "
    End If
    
    GoSub Spread_Clear
    If False = adoSetOpen(gStrSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        ssResult.Col = 2:     ssResult.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssResult.Col = 3:     ssResult.Text = adoSet.Fields("SlipNo1").Value & ""
        ssResult.Col = 4:     ssResult.Text = adoSet.Fields("SlipNo2").Value & ""
        ssResult.Col = 5:     ssResult.Text = Trim(adoSet.Fields("RoomCode").Value & "")
        ssResult.Col = 6:     ssResult.Text = adoSet.Fields("Ptno").Value & ""
                              LsPtno = adoSet.Fields("Ptno").Value & ""
        ssResult.Col = 7:     ssResult.Text = adoSet.Fields("Sname").Value & ""
        ssResult.Col = 8:     ssResult.Text = adoSet.Fields("Sex").Value & ""
        
        ssResult.Col = 10:    ssResult.Text = adoSet.Fields("GeomsaDt").Value & ""
        ssResult.Col = 11:    ssResult.Text = Format(adoSet.Fields("GeomsaT1").Value, "00") & ":" & _
                                              Format(adoSet.Fields("GeomsaT2").Value, "00")
        If Trim(adoSet.Fields("Name").Value & "") = "" Then
            ssResult.Col = 12: ssResult.Text = GstrIdnumber
        Else
            ssResult.Col = 12:    ssResult.Text = adoSet.Fields("Name").Value & ""
        End If
        ssResult.Col = 18:    ssResult.Text = Format(adoSet.Fields("JeobsuT1").Value, "00") & ":" & _
                                              Format(adoSet.Fields("JeobsuT2").Value, "00")
        ssResult.Col = 14:    ssResult.Text = Trim(adoSet.Fields("Deptnamek").Value & "")
                              LsDeptCode = adoSet.Fields("DeptCode").Value & ""
        
        ssResult.Col = 15:    ssResult.Text = adoSet.Fields("DrName").Value & ""
                              LsDrCode = adoSet.Fields("DrCode").Value & ""
        
        ssResult.Col = 16:    ssResult.Text = Trim(adoSet.Fields("GeomsaCm").Value & "")
        ssResult.Col = 17:    ssResult.Text = adoSet.Fields("AgeYY").Value & ""
        ssResult.Col = 13:    ssResult.Text = Trim(adoSet.Fields("GeomchCD").Value & "")
                              LsCodeKy = adoSet.Fields("GeomchCd").Value & ""
        
        ssResult.Col = 20:    ssResult.Text = adoSet.Fields("ROWID").Value & ""
        ssResult.Col = 21:    ssResult.Text = adoSet.Fields("REPORT1").Value & ""
        ssResult.Col = 22:    ssResult.Text = adoSet.Fields("ORDERDT").Value & ""

        Select Case adoSet.Fields("Status").Value & ""
            Case "R"
                ssResult.Col = 9
                ssResult.Text = "접수중"
                ssResult.ForeColor = RGB(255, 0, 0)
            Case "C"
                ssResult.Col = 21
                If Val(ssResult.Text) > 0 Then
                    ssResult.Col = 9
                    ssResult.Text = "검사완료"
                    ssResult.ForeColor = RGB(0, 128, 128)
                Else
                    ssResult.Col = 9
                    ssResult.Text = "검사완료"
                    ssResult.ForeColor = RGB(0, 0, 255)
                End If
            Case "P"
                ssResult.Col = 9
                ssResult.Text = "부분결과"
                ssResult.ForeColor = RGB(0, 170, 0)
            Case "U"
                ssResult.Col = 9
                ssResult.Text = "미확인"
                ssResult.ForeColor = RGB(0, 0, 0)
            Case "X"
                ssResult.Col = 9
                ssResult.Text = "Data이상"
                ssResult.ForeColor = RGB(0, 170, 0)
                
            Case Else
                ssResult.Col = 9
                ssResult.Text = "접수중"
                ssResult.ForeColor = RGB(255, 0, 0)
        End Select
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    GoSub Spread_Lock_True_Set
        
    ssResult.SetFocus
   
    Exit Sub
   
'/__________________________________________________________________

Spread_Clear:
    ssResult.Row = 1
    ssResult.Row2 = ssResult.DataRowCnt
    ssResult.Col = 1
    ssResult.Col2 = ssResult.DataColCnt
    ssResult.BlockMode = True
    ssResult.Action = SS_ACTION_CLEAR_TEXT
    ssResult.BlockMode = False
    Return

Spread_Lock_True_Set:
    ssResult.Col = 2:    ssResult.Col2 = ssResult.MaxCols
    ssResult.Row = 1:    ssResult.Row2 = ssResult.DataRowCnt
    ssResult.BlockMode = True
    ssResult.Lock = True
    ssResult.BlockMode = False
    Return
    
End Sub

Private Sub cmdPr_Click()
'/----------------------------------------------------------------------------------------------/
'/-----   Spread 를 이용하지 않고 Printer Object 를 이용한 Reporting 방법                       /
'/-----   1999/04/07(Kwak)                                                                      /
'/-----   생성 시키고 Test 미완료 된 상태임.(Remark Printing 을 Page Footer 로 이용하는것이     /
'/-----                                      문제가 있는것 같아 생성시킨 Routine 임).           /
'/----------------------------------------------------------------------------------------------/
    Dim sSLipno1    As String
    Dim sSLipno2    As String
    Dim sPtno       As String * 8
    Dim sJeobsuDt   As String
    Dim sSex        As String
    Dim sGeomsaDt   As String
    Dim sAge        As String
    Dim sRemark     As String
    Dim sSname      As String * 10
    Dim sRitemCd    As String * 8
    Dim sRitemNm    As String * 30
    Dim sDanWi      As String * 6
    Dim sResult     As String * 12
    Dim sResult42   As String
    Dim sMin        As String * 6
    Dim sMax        As String * 6
    Dim sRoomCode   As String
    Dim sDeptName   As String
    Dim sSamplename As String
    Dim iSensCount      As Integer
    Dim iSensOrderName  As String
    
    
    If vbNo = MsgBox(" Print 하시겠습니까?", vbQuestion + vbYesNo, "Printing 할까요?") Then
        Exit Sub
    End If
    
    '서울 건양병원(전산실)에서 일괄 Printing 하기위하여 Sort한 Sub 임.
    '필요할시 Comment 제거후 쓸것(과별,환자번호,검사종목(slipno1)
    'GoSub PrintSpread_Sort

    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
        
        If Left(cmbSLip.Text, 2) = "15" Then
            MsgBox "골수검사는 결과입력Form 에서 개별로 출력이 가능합니다!..."
            Exit Sub
        Else
            If ssResult.Value = True Then
                GoSub Variable_Setting
                GoSub Main_Process
            End If
        End If
    Next
    Exit Sub
    
PrintSpread_Sort:
    ssResult.Row = 1
    ssResult.Row2 = ssResult.DataRowCnt
    ssResult.Col = 1
    ssResult.Col2 = ssResult.DataColCnt
    
    ssResult.SortBy = SS_SORT_BY_ROW
    ssResult.SortKey(1) = 14        '과별
    ssResult.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
    
    ssResult.SortKey(2) = 6        '환자번호
    ssResult.SortKeyOrder(2) = SS_SORT_ORDER_ASCENDING
    
    ssResult.SortKey(3) = 5         '검사종목
    ssResult.SortKeyOrder(5) = SS_SORT_ORDER_ASCENDING

    ssResult.Action = SS_ACTION_SORT

    
    Return
    
Variable_Setting:
    ssResult.Col = 2:  sJeobsuDt = Trim$(ssResult.Text)
    ssResult.Col = 3:  sSLipno1 = Trim$(ssResult.Text)
    ssResult.Col = 4:  sSLipno2 = Trim$(ssResult.Text)
    ssResult.Col = 6:  sPtno = Trim$(ssResult.Text)
    ssResult.Col = 7:  sSname = Trim$(ssResult.Text)
    ssResult.Col = 8:  sSex = Trim$(ssResult.Text)
    ssResult.Col = 10: sGeomsaDt = Trim$(ssResult.Text)
    ssResult.Col = 5: sRoomCode = Trim$(ssResult.Text)
    ssResult.Col = 14: sDeptName = Trim$(ssResult.Text)
    ssResult.Col = 16: sRemark = Trim$(ssResult.Text)
    ssResult.Col = 17: sAge = Trim$(ssResult.Text)
    Return
    
Main_Process:
    
    
    StrSql = ""
    StrSql = StrSql & " SELECT a.ItemCD, a.Result1, a.Chamgo, b.itemNM"
    StrSql = StrSql & " FROM   TWEXAM_GENERAL_SUB  a,"
    StrSql = StrSql & "        TWEXAM_ITEMML       b "
    StrSql = StrSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    a.Slipno1  =  " & Val(sSLipno1)
    StrSql = StrSql & " AND    a.SLipno2  =  " & Val(sSLipno2)
    StrSql = StrSql & " AND    a.PTNO     = '" & sPtno & "'"
    StrSql = StrSql & " AND    a.ITemCD   = b.Codeky"
    StrSql = StrSql & " ORDER  BY a.iTemcd"
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    GoSub PrintHead_RTN
    GoSub Print_OK_RTN
    Printer.EndDoc
    
    
    Call adoSetClose(adoSet)
    
    Return


PrintHead_RTN:
    Dim sDeptCode   As String * 10
    Dim sAgeYY      As String
    Dim sJDT        As String
    Dim sGDT        As String
    Dim sSlipTitle  As String
    Dim sGeomsaJa   As String
    Dim adoSpec     As ADODB.Recordset
    Dim adoGen      As ADODB.Recordset
    
    
    StrSql = ""
    StrSql = StrSql & " SELECT codenm"
    StrSql = StrSql & " FROM   TWEXAM_SPECODE"
    StrSql = StrSql & " WHERE  CODEGU = '12'"
    StrSql = StrSql & " AND    CODEKY = '" & sSLipno1 & "'"
    If adoSetOpen(StrSql, adoSpec) Then
        sSlipTitle = adoSpec.Fields("Codenm").Value & ""
        Call adoSetClose(adoSpec)
    End If
    
    StrSql = ""
    StrSql = StrSql & " SELECT /*+ INDEX (TWBas_DEPT INDEX_DEPT0) */"
    StrSql = StrSql & "        a.RoomCode, a.Sex, a.AgeYY, "
    StrSql = StrSql & "        TO_CHAR(a.JeobsuDT,'YYYY-MM-DD') JeobsuDt,"
    StrSql = StrSql & "        TO_CHAR(a.GeomsaDT,'YYYY-MM-DD') GeomsaDt, "
    StrSql = StrSql & "        a.GeomsaCM, b.DeptNamek, c.Name, d.Codenm"
    StrSql = StrSql & " FROM   TWEXAM_GENERAL a,"
    StrSql = StrSql & "        TWBAS_DEPT     b,"
    StrSql = StrSql & "        TWBas_Pass     c,"
    StrSql = StrSql & "        TWEXAM_Sample  d "
    StrSql = StrSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    a.Ptno     = '" & sPtno & "'"
    StrSql = StrSql & " AND    a.Slipno1  =  " & Val(sSLipno1)
    StrSql = StrSql & " AND    a.Slipno2  =  " & Val(sSLipno2)
    StrSql = StrSql & " AND    a.DeptCode = b.DeptCode(+)"
    StrSql = StrSql & " AND    a.Geomsaja = c.idNumber(+)"
    StrSql = StrSql & " AND    a.GeomchCd = d.Code(+)"
    
    If False = adoSetOpen(StrSql, adoGen) Then Return
    
    sDeptCode = adoGen.Fields("DeptNamek").Value & ""
    sRoomCode = adoGen.Fields("RoomCode").Value & ""
    sSex = adoGen.Fields("Sex").Value & ""
    sAgeYY = adoGen.Fields("AgeYY").Value & ""
    sJDT = adoGen.Fields("JeobsuDt").Value & ""
    sGDT = adoGen.Fields("GeomsaDt").Value & ""
    sRemark = Trim(adoGen.Fields("GeomsaCm").Value & "")
    sGeomsaJa = Trim(adoGen.Fields("Name").Value & "")
    sSamplename = Trim(adoGen.Fields("Codenm").Value & "")
    Call adoSetClose(adoGen)
    
    Printer.FontName = "굴림체"
    Printer.FontSize = "12"
    Printer.FontBold = True
    Printer.FontItalic = True
    
    Printer.Print sSlipTitle & "         Result Report"
    Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━"
    
    Printer.FontName = "굴림체"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.FontItalic = False
    
    Printer.Print "병록No: " & sPtno; Tab(50); "검  체: " & sSamplename
    Printer.Print "성  명: " & Trim(sSname) & "[" & sSex & "/" & sAgeYY & "]"; Tab(50); "검사자: " & sGeomsaJa
    Printer.Print "진료과: " & sDeptName; Tab(50); "접수일: " & sJDT
    Printer.Print "병  실: " & sRoomCode; Tab(50); "검사일: " & sGDT

    If Left(sSLipno1, 1) = "4" Then
        Printer.Print "━━━━━━━━━━━━━━┳━━━━━━━━━━━━━━━━━━━"
        Printer.Print "       검사항목             ┃        검사결과                      "
        Printer.Print "━━━━━━━━━━━━━━┻━━━━━━━━━━━━━━━━━━━"
    Else
        Printer.Print "━━━━━━━━━━━━━━┳━━━━━┳━━━━━━━━┳━━━━"
        Printer.Print "       검사항목             ┃ 검사결과 ┃    참고치      ┃  단위    "
        Printer.Print "━━━━━━━━━━━━━━┻━━━━━┻━━━━━━━━┻━━━━"
    
    End If
    Return
    
    
Print_OK_RTN:
    Dim j            As Integer
    Dim adoGsub      As ADODB.Recordset
    Dim sSensFlag    As String * 1
    Dim nPrcnt       As Integer
    Dim saItemCd(10) As String
    Dim iL           As Integer
    
    sSensFlag = ""
    
    StrSql = ""
    StrSql = StrSql & " SELECT TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt,"
    StrSql = StrSql & "        a.ItemCd, a.Result1, "
    StrSql = StrSql & "        b.ItemNM, b.MinCham, b.MaxCham, b.DanWi, b.MinDanger, b.MaxDanger,"
    StrSql = StrSql & "        b.ResultW"
    StrSql = StrSql & " FROM   TWEXAM_GENERAL_SUB a,"
    StrSql = StrSql & "        TWEXAM_ITEMML      b "
    StrSql = StrSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    a.Ptno     = '" & sPtno & "'"
    StrSql = StrSql & " AND    a.Slipno1  =  " & Val(sSLipno1)
    StrSql = StrSql & " AND    a.Slipno2  =  " & Val(sSLipno2)
    StrSql = StrSql & " AND    a.Verify   = 'Y'"
    StrSql = StrSql & " AND    a.ItemCd   = b.CodeKy(+)"
    StrSql = StrSql & " ORDER  BY a.itemCd"
    
    If False = adoSetOpen(StrSql, adoGsub) Then Return
    
    
    Printer.FontName = "굴림체"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.FontItalic = False
    
    For iL = 0 To 10
        saItemCd(iL) = ""
    Next
    
    iL = 0
    Do Until adoGsub.EOF
        sRitemCd = adoGsub.Fields("ItemCd").Value & ""
        sRitemNm = adoGsub.Fields("ItemNm").Value & ""
        RSet sDanWi = adoGsub.Fields("DanWi").Value & ""
        LSet sResult = adoGsub.Fields("Result1").Value & ""
        
        If Left(sSLipno1, 1) <> "4" Then
            GoSub Get_RefData
            Printer.Print "  " & sRitemNm & sResult & sMin & " ~ " & sMax & sDanWi
        Else
            If Trim(adoGsub.Fields("ResultW").Value & "") = "S" Then
                sSensFlag = "*"
                sResult42 = "다음장참조"
                iSensOrderName = "": iSensOrderName = sRitemNm
                saItemCd(iL) = sRitemCd
                iL = iL + 1
                Printer.Print "  " & sRitemNm & sResult42
            Else
                sResult42 = adoGsub.Fields("Result1").Value & ""
                Printer.Print "  " & sRitemNm & Trim(sResult42)
            End If
        End If
        
        adoGsub.MoveNext
    Loop
    Call adoSetClose(adoGsub)
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print "[ Remark ___________________________: ]"
    Printer.Print sRemark
    Printer.NewPage

    Dim sTmpRcode       As String
    Dim sDispRname      As String * 25
    Dim iCount          As Integer
    
    
    For iL = 0 To 10
        iCount = iL
        If saItemCd(iL) = "" Then
            Exit For
        End If
    Next
    
    If sSensFlag = "*" Then
        For iL = 0 To iCount
            iSensCount = Printing_Sens(sJeobsuDt, Val(sSLipno1), Val(sSLipno2), saItemCd(iL))
            If iSensCount > 0 Then
                Printer.FontName = "굴림체"
                Printer.FontSize = "12"
                Printer.FontBold = True
                Printer.FontItalic = True
            
                Printer.Print sSlipTitle & "(SENSTIVITY)   Result Report"
                Printer.FontName = "굴림체"
                Printer.FontSize = 9
                Printer.FontBold = False
                Printer.FontItalic = False
                
                Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
                Printer.Print "병록No: " & sPtno; Tab(50); "검  체: " & sSamplename
                Printer.Print "성  명: " & Trim(sSname) & "[" & sSex & "/" & sAgeYY & "]"; Tab(50); "검사자: " & sGeomsaJa
                Printer.Print "진료과: " & sDeptName; Tab(50); "접수일: " & sJDT
                Printer.Print "병  실: " & sRoomCode; Tab(50); "검사일: " & sGDT
                Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
                
                Printer.FontName = "바탕체": Printer.FontSize = 10:  Printer.FontBold = True: Printer.FontItalic = False
                Printer.Print "Order    : " & iSensOrderName
                Printer.Print "보고형태 : " & Get_General_Status(sJeobsuDt, Val(sSLipno1), Val(sSLipno2))
                
                For nPrcnt = 0 To iSensCount - 1
                    If sTmpRcode = SensResult.Rcode(nPrcnt) Then
                        Printer.Print "     " & SensResult.AntiName(nPrcnt) & "," & SensResult.Sens(nPrcnt)
                                      
                    Else
                        Printer.Print ""
                        Printer.FontName = "바탕체": Printer.FontSize = 9:  Printer.FontBold = True: Printer.FontItalic = False
                        Printer.Print "@." & _
                                      SensResult.Rcode(nPrcnt) & vbCrLf & _
                                      "     " & SensResult.AntiName(nPrcnt) & "," & SensResult.Sens(nPrcnt)
                    End If
                    sTmpRcode = SensResult.Rcode(nPrcnt)
                Next
                
            End If
            If iCount > 0 Then Printer.NewPage
        Next
    End If
    
    Printer.EndDoc
    Return


Get_RefData:
    Dim adoRef      As ADODB.Recordset
    
    sMin = "": sMax = ""
    
    StrSql = ""
    StrSql = StrSql & " SELECT * "
    StrSql = StrSql & " FROM   TWEXAM_REFDATA"
    StrSql = StrSql & " WHERE  ITEMCODE  = '" & sRitemCd & "'"
    StrSql = StrSql & " AND    AGEMIN   <=  " & Val(sAge)
    StrSql = StrSql & " AND    AGEMAX   >=  " & Val(sAge)
    StrSql = StrSql & " AND    APPDATE   =     (SELECT MAX(APPDATE)"
    StrSql = StrSql & "                         FROM   TWEXAM_REFDATA"
    StrSql = StrSql & "                         WHERE  ITEMCODE = '" & sRitemCd & "'"
    StrSql = StrSql & "                         AND    AGEMIN  <=  " & Val(sAge)
    StrSql = StrSql & "                         AND    AGEMAX  >=  " & Val(sAge) & ")"
    
    If adoSetOpen(StrSql, adoRef) Then
        If sSex = "M" Then
            RSet sMin = adoRef.Fields("M_MIN").Value & ""
            sMax = adoRef.Fields("M_MAX").Value & "": End If
        If sSex = "F" Then
            RSet sMin = adoRef.Fields("F_MIN").Value & ""
            sMax = adoRef.Fields("F_MAX").Value & "": End If
        Call adoSetClose(adoRef)
    End If
    
    Return


Print_Text_Ret:
    Dim sRetCham    As String
    Dim sRet        As String
    Dim adoCham     As ADODB.Recordset
    
    StrSql = ""
    StrSql = StrSql & " SELECT TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt,"
    StrSql = StrSql & "        a.ItemCd, a.Result1, b.ItemNM, b.MinCham, b.MaxCham, a.Chamgo"
    StrSql = StrSql & " FROM   TWEXAM_GENERAL_SUB a,"
    StrSql = StrSql & "        TWEXAM_ITEMML      b "
    StrSql = StrSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    a.Ptno     = '" & sPtno & "'"
    StrSql = StrSql & " AND    a.SLIPNO1  =  " & Val(sSLipno1)
    StrSql = StrSql & " AND    a.Verify   = 'Y'"
    StrSql = StrSql & " AND    a.ItemCd   = b.CodeKy(+)"
    
    If False = adoSetOpen(StrSql, adoCham) Then Return
    
    Do Until adoCham.EOF
        sRitemCd = adoCham.Fields("ItemCd").Value & ""
        sRitemNm = adoCham.Fields("ItemNm").Value & ""
        
        sRet = Trim$(adoCham.Fields("Result1").Value & "")
        sRetCham = Trim$(adoCham.Fields("Chamgo").Value & "")
        GoSub Check_Bit_Chamgo
        Printer.Print sRitemCd & " " & sRitemNm
        Printer.Print "결과:__ "
        Printer.Print "        " & sRet
        If Trim$(sRetCham) <> "" Then
            Printer.Print "참고사항: "
            Printer.Print sRetCham
        End If
        
        adoCham.MoveNext
    Loop
    Call adoSetClose(adoCham)
    
    
    Printer.Print ""
    Printer.Print ""
    If Trim$(sRemark) <> "" Then
        Printer.Print "[ Remark ___________________________: ]"
        Printer.Print sRemark
    End If
    
    
    Return
    
Check_Bit_Chamgo:
    Dim nLength As Double
    Dim sTarget As String
    Dim nCnt    As Integer
    nLength = Len(sRetCham)
    
    nCnt = 1
    For i = 1 To nLength
        If nCnt > 62 Then
            sTarget = sTarget & vbCrLf & Mid(sRetCham, i, 1)
            nCnt = 1
        Else
            sTarget = sTarget & Mid(sRetCham, i, 1)
            nCnt = nCnt + 1
        End If
        
    Next
    sRetCham = sTarget
    Return

End Sub

Public Sub CmdResult_Click()
    
    Dim i    As Integer
    
    Call ssResult_LeaveCell(1, 1, 1, 1, 1)
    
    GiProcess_row = 0
    
    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
        
        If ssResult.Text = "1" Then
            If GiProcess_row = 0 Then
                GiProcess_row = i
            End If
            ssResult.Col = 19
            ssResult.Text = "1"
        Else
            ssResult.Col = 19
            ssResult.Text = ""
        End If
    Next i
    
    If GiProcess_row = 0 Then
        MsgBox "결과를 등록할 선택된 행이 없습니다!...", vbInformation
        Exit Sub
    End If
    
    If cmbSLip.ListIndex = -1 Then Exit Sub
    
    If GiExamNumb = 15 Then
        frmBM.Show vbModal
        Exit Sub
    End If
    
    
    DoEvents: clpSlip1.Show vbModal

    
End Sub



Private Sub cmdSample_Click()
    
    hWndReturn = txtSampleData.hWnd
    frmQryGeom.Show vbModal
    
    For i = Me.ssResult.DataRowCnt To 1 Step -1
        ssResult.Row = i
        ssResult.Col = 13
        If Trim(txtSampleData) <> Trim(ssResult.Text) Then
            ssResult.Action = ActionDeleteRow
        End If
    Next
    
End Sub

Private Sub Form_Load()
 
    
    chkJeobsu.Enabled = True
    chkPart.Enabled = True
    chkComplete.Enabled = True
    chkRepCd.Enabled = True
    chkUnknow.Enabled = True
     
     
    GoSub SLip_Select
    
    lblGeomsaDt = Dual_Date_Get("yyyy-MM-dd")
    dtFromJeobsu.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToJeobsu.Value = Dual_Date_Get("yyyy-MM-dd")
    
    txtStartSlipno = 0
    txtEndSlipno = 0
    
    Exit Sub
    
    
    
SLip_Select:
    StrSql = ""
    StrSql = StrSql & " SELECT *"
    StrSql = StrSql & " FROM   TWEXAM_Specode"
    StrSql = StrSql & " WHERE  Codegu = '12'"
    StrSql = StrSql & " AND    Codeky < '52'"
    StrSql = StrSql & " ORDER  BY Codeky"
    
    cmbSLip.Clear
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        cmbSLip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                        Trim(adoSet.Fields("Codenm").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
        
    Return
    
End Sub
    
Private Sub mnuAbnormal_Click()
    
    frmAbnormal.Show vbModal
    
    
End Sub

Private Sub mnuAppend_Click()
    
    Call CmdResult_Click
    
End Sub

Private Sub mnuDesel_Click()
    Dim i       As Integer
    
    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
        ssResult.Value = "0"
    Next

End Sub

Private Sub mnuExit_Click()

    Unload Me
  
End Sub

Private Sub mnuPexit_Click()
    
    Call cmdExit_Click
    
End Sub

Private Sub mnuItemResult_Click()
    
   
End Sub

Private Sub mnuPrint_Click()
    Call cmdPr_Click
    
End Sub

Private Sub mnuPtno_Click()
    
    gMenuSelect = 1
    frmWhere.Show vbModal
    
End Sub

Private Sub mnuQry_Click()
    
    Call cmdInquiry_Click
    
End Sub

Private Sub mnuSel_Click()
    Dim i       As Integer
    
    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
        ssResult.Value = "1"
    Next

End Sub

Private Sub mnuSname_Click()
    
    gMenuSelect = 2
    frmWhere.Show vbModal

End Sub

Private Sub optAll_Click(Value As Integer)
    
    chkJeobsu.Value = False:    chkJeobsu.Enabled = False
    chkPart.Value = False:      chkPart.Enabled = False
    chkComplete.Value = False:  chkComplete.Enabled = False
    chkRepCd.Value = False:     chkRepCd.Enabled = False
    chkUnknow.Value = False:    chkUnknow.Enabled = False
    
    txtStartSlipno.Text = "0"
    txtEndSlipno.Text = "0"
    
End Sub

Private Sub optCondition_Click(Value As Integer)
    
    chkJeobsu.Enabled = True
    chkPart.Enabled = True
    chkComplete.Enabled = True
    chkRepCd.Enabled = True
    chkUnknow.Enabled = True
    
End Sub

Private Sub SpinButton1_SpinDown()

    txtStartSlipno = Val(txtStartSlipno) - 1
    
End Sub

Private Sub SpinButton1_SpinUp()

    txtStartSlipno = Val(txtStartSlipno) + 1

End Sub


Private Sub SpinButton2_SpinDown()
    
    txtEndSlipno = Val(txtEndSlipno) - 1

End Sub

Private Sub SpinButton2_SpinUp()

    txtEndSlipno = Val(txtEndSlipno) + 1

End Sub





Private Sub SpinButton1_Click()

End Sub


Public Sub ssResult_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim i       As Integer
    
    If Row = 0 And Col = 1 Then
        ssResult.Col = 1:        ssResult.Row = 0
        If ssResult.Text = "A" Then
            ssResult.Col = 1
            ssResult.Row = 0
            ssResult.Text = "C"
            For i = 1 To ssResult.DataRowCnt
                ssResult.Row = i
                ssResult.Text = "0"
            Next i
        Else
            ssResult.Col = 1
            ssResult.Row = 0
            ssResult.Text = "A"
            For i = 1 To ssResult.DataRowCnt
                ssResult.Row = i
                ssResult.Text = "1"
            Next i
        End If
    End If
    
End Sub


Private Sub ssResult_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        If Col > 1 Then
            GoSub SpreadSort_Sub
        End If
    End If
    Exit Sub
    
SpreadSort_Sub:
    ssResult.Col = 1:  ssResult.Col2 = ssResult.MaxCols
    ssResult.Row = 1:  ssResult.Row2 = ssResult.DataRowCnt
    
    ssResult.SortBy = SS_SORT_BY_ROW
    ssResult.SortKey(1) = Col
    If ssResult.SortKeyOrder(1) = SortKeyOrderDescending Then
        ssResult.SortKeyOrder(1) = SortKeyOrderAscending
    Else
        ssResult.SortKeyOrder(1) = SortKeyOrderDescending
    End If
    ssResult.Action = SS_ACTION_SORT
    
    Return
    
End Sub

Private Sub ssResult_GotFocus()

'    ssResult.Col = 1:       ssResult.Row = 1
'    ssResult.Action = SS_ACTION_ACTIVE_CELL
    
End Sub


Private Sub ssResult_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    
    KeyAscii = 0
    
    SendKeys "{tab}"
    
End Sub


Private Sub ssResult_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    If NewRow < 0 Then Exit Sub
    If NewRow > ssResult.DataRowCnt Then Exit Sub
    
    ssResult.Row = NewRow
    
    ssResult.Col = 2:   lblJeobsuDt = ssResult.Text
    ssResult.Col = 3:   lblSlipNo = ssResult.Text
    ssResult.Col = 4:   lblSlipNo = lblSlipNo & "-" & ssResult.Text
    ssResult.Col = 5:   lblRoom = ssResult.Text
    ssResult.Col = 6:   lblPtNo = ssResult.Text
    ssResult.Col = 7:   lblName = ssResult.Text
    ssResult.Col = 8:   lblSex = ssResult.Text
    ssResult.Col = 12:  lblGeomsaJa = ssResult.Text
    ssResult.Col = 14:  lblDeptname = ssResult.Text
    ssResult.Col = 15:  lblDrname = ssResult.Text
    ssResult.Col = 17:  lblAge = ssResult.Text
    ssResult.Col = 10: lblGeomsaDt = ssResult.Text
    
    
End Sub


Private Sub ssResult_LostFocus()
    
    ssResult.Col = 1:       ssResult.Row = 0
    ssResult.Action = SS_ACTION_ACTIVE_CELL
    
End Sub


Private Sub ssResult_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'If Button = 2 Then
    '    PopupMenu mnuJob
    'End If
    
End Sub

Private Sub txtEndSlipno_GotFocus()
    
    txtEndSlipno.SelStart = 0
    txtEndSlipno.SelLength = Len(txtEndSlipno.Text)

End Sub


Private Sub txtEndSlipno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        
        If txtStartSlipno.Text <> "" And _
            txtEndSlipno.Text = "0" Then
            txtEndSlipno.Text = txtStartSlipno.Text
        End If
        GoSub DO_Process
    
    End If
    Exit Sub
    

DO_Process:
    
    DoEvents: Call cmdInquiry_Click
    
    If ssResult.DataRowCnt = 0 Then
        MsgBox "해당 검체번호의 Data 는 없습니다!..."
        Return
    End If
    
    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
        ssResult.Text = "1"
    Next
    ssResult.Row = 0
    ssResult.Col = 1
    ssResult.Text = "A"
    
    DoEvents: Call CmdResult_Click
    
    Return
    
End Sub


Private Sub txtEndSlipno_LostFocus()
   
   GiEndSlipno = Val(txtEndSlipno)

End Sub



Private Sub txtStartSlipno_GotFocus()
    
    txtStartSlipno.SelStart = 0
    txtStartSlipno.SelLength = Len(txtStartSlipno.Text)

End Sub

Private Sub txtStartSlipno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtEndSlipno.SetFocus
        If Trim(txtStartSlipno.Text) = "" Then Exit Sub
        
    End If
    
    
End Sub



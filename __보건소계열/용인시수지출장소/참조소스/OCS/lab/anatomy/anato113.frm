VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Anato_DiagName_Search 
   BorderStyle     =   0  '없음
   Caption         =   "진단명조회"
   ClientHeight    =   8505
   ClientLeft      =   165
   ClientTop       =   1830
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8505
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin VB.Frame Frame2 
      Caption         =   "검체보관 유무"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   8490
      TabIndex        =   52
      Top             =   3240
      Width           =   3315
      Begin Threed.SSOption optG 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   53
         Top             =   240
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "유"
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
      Begin Threed.SSOption optG 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   54
         Top             =   240
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "무"
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "조회조건(진단완료)"
      ForeColor       =   &H00FF0000&
      Height          =   945
      Left            =   8490
      TabIndex        =   47
      Top             =   2250
      Width           =   3315
      Begin MSComCtl2.DTPicker dtToDate 
         Height          =   315
         Left            =   1530
         TabIndex        =   48
         Top             =   570
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36526
      End
      Begin MSComCtl2.DTPicker dtFromDate 
         Height          =   315
         Left            =   1530
         TabIndex        =   49
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36526
      End
      Begin VB.Label Label17 
         Caption         =   "TO"
         Height          =   165
         Left            =   300
         TabIndex        =   51
         Top             =   630
         Width           =   1125
      End
      Begin VB.Label Label16 
         Caption         =   "FROM"
         Height          =   165
         Left            =   300
         TabIndex        =   50
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2820
      Left            =   8505
      ScaleHeight     =   2760
      ScaleWidth      =   3240
      TabIndex        =   44
      Top             =   3900
      Width           =   3300
      Begin FPSpread.vaSpread ssExDate 
         Height          =   2490
         Left            =   0
         TabIndex        =   45
         Top             =   285
         Width           =   3270
         _Version        =   196608
         _ExtentX        =   5768
         _ExtentY        =   4392
         _StockProps     =   64
         BackColorStyle  =   1
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
         GridSolid       =   0   'False
         MaxCols         =   12
         OperationMode   =   1
         ShadowColor     =   12632256
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "ANATO113.frx":0000
         VisibleCols     =   10
         VisibleRows     =   500
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "검사일자"
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
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   3255
      End
   End
   Begin RichTextLib.RichTextBox txtSummary 
      Height          =   420
      Left            =   45
      TabIndex        =   36
      Top             =   7740
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   741
      _Version        =   393217
      BackColor       =   15463903
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"ANATO113.frx":18B6
   End
   Begin RichTextLib.RichTextBox txtDiag 
      Height          =   4860
      Left            =   45
      TabIndex        =   35
      Top             =   2490
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   8573
      _Version        =   393217
      BackColor       =   15463903
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"ANATO113.frx":1B1B
   End
   Begin VB.PictureBox picPanel 
      Height          =   1428
      Left            =   8490
      ScaleHeight     =   1365
      ScaleWidth      =   3270
      TabIndex        =   7
      Top             =   6768
      Width           =   3330
      Begin Threed.SSCommand cmdWork 
         Height          =   700
         Index           =   15
         Left            =   10
         TabIndex        =   15
         Top             =   5055
         Width           =   700
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "ANATO113.frx":1D80
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   705
         Index           =   14
         Left            =   740
         TabIndex        =   14
         Top             =   4340
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO113.frx":209A
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   705
         Index           =   12
         Left            =   740
         TabIndex        =   13
         Top             =   3620
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO113.frx":2974
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   705
         Index           =   16
         Left            =   740
         TabIndex        =   12
         Top             =   5055
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO113.frx":324E
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   705
         Index           =   11
         Left            =   10
         TabIndex        =   11
         Top             =   3620
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   705
         Index           =   13
         Left            =   10
         TabIndex        =   10
         Top             =   4340
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   705
         Index           =   10
         Left            =   740
         TabIndex        =   9
         Top             =   2890
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO113.frx":3568
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   705
         Index           =   9
         Left            =   10
         TabIndex        =   8
         Top             =   2890
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO113.frx":4CFA
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   690
         Left            =   1635
         TabIndex        =   6
         Top             =   690
         Width           =   1650
         _Version        =   65536
         _ExtentX        =   2910
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "종      료"
         ForeColor       =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         Picture         =   "ANATO113.frx":55D4
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   690
         Left            =   1635
         TabIndex        =   4
         Top             =   -15
         Width           =   1650
         _Version        =   65536
         _ExtentX        =   2910
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "인      쇄"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO113.frx":58EE
      End
      Begin Threed.SSCommand cmdStrInq 
         Height          =   690
         Left            =   -15
         TabIndex        =   3
         Top             =   0
         Width           =   1650
         _Version        =   65536
         _ExtentX        =   2910
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "조      회"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO113.frx":5C08
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   690
         Left            =   -15
         TabIndex        =   5
         Top             =   690
         Width           =   1650
         _Version        =   65536
         _ExtentX        =   2910
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "취      소"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO113.frx":5F22
      End
   End
   Begin Threed.SSFrame frmPatient 
      Height          =   1425
      Left            =   60
      TabIndex        =   16
      Top             =   735
      Width           =   8280
      _Version        =   65536
      _ExtentX        =   14605
      _ExtentY        =   2514
      _StockProps     =   14
      Caption         =   "환자정보"
      ForeColor       =   12583104
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin VB.TextBox txtOrganPart 
         BackColor       =   &H00EBF5DF&
         Height          =   300
         Left            =   6450
         TabIndex        =   40
         Top             =   630
         Width           =   1620
      End
      Begin VB.TextBox txtOpname 
         BackColor       =   &H00EBF5DF&
         Height          =   300
         Left            =   6450
         TabIndex        =   39
         Top             =   270
         Width           =   1620
      End
      Begin VB.Label Label12 
         Caption         =   "장  기  명"
         Height          =   225
         Left            =   5640
         TabIndex        =   38
         Top             =   690
         Width           =   810
      End
      Begin VB.Label Label11 
         Caption         =   "수  술  명"
         Height          =   225
         Left            =   5640
         TabIndex        =   37
         Top             =   330
         Width           =   810
      End
      Begin VB.Label lblRoom 
         BackColor       =   &H00EBF5DF&
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
         Height          =   300
         Left            =   6450
         TabIndex        =   34
         Top             =   990
         Width           =   1620
      End
      Begin VB.Label lblSex 
         BackColor       =   &H00EBF5DF&
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
         Height          =   300
         Left            =   3705
         TabIndex        =   33
         Top             =   990
         Width           =   1635
      End
      Begin VB.Label lblDept 
         BackColor       =   &H00EBF5DF&
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
         Height          =   300
         Left            =   3705
         TabIndex        =   32
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblSname 
         BackColor       =   &H00EBF5DF&
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
         Height          =   300
         Left            =   1035
         TabIndex        =   31
         Top             =   990
         Width           =   1635
      End
      Begin VB.Label lblAge 
         BackColor       =   &H00EBF5DF&
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
         Height          =   300
         Left            =   3705
         TabIndex        =   30
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label lblPtNo 
         BackColor       =   &H00EBF5DF&
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
         Height          =   300
         Left            =   1035
         TabIndex        =   29
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "병     실"
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
         Left            =   5640
         TabIndex        =   28
         Top             =   1035
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "환자번호"
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
         Left            =   264
         TabIndex        =   23
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "성    명"
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
         Left            =   264
         TabIndex        =   22
         Top             =   996
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "나    이"
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
         Left            =   2925
         TabIndex        =   21
         Top             =   675
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "성    별"
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
         Left            =   2925
         TabIndex        =   20
         Top             =   990
         Width           =   720
      End
      Begin VB.Label lblAnatoNo 
         BackColor       =   &H00EBF5DF&
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
         Height          =   300
         Left            =   1035
         TabIndex        =   19
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "진 료 과"
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
         Left            =   2925
         TabIndex        =   18
         Top             =   330
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "병리번호"
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
         Left            =   264
         TabIndex        =   17
         Top             =   345
         Width           =   720
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1440
      Left            =   8490
      TabIndex        =   24
      Top             =   735
      Width           =   3315
      _Version        =   65536
      _ExtentX        =   5847
      _ExtentY        =   2540
      _StockProps     =   14
      Caption         =   "조회조건"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin VB.TextBox txtStrOrgan 
         BackColor       =   &H00C8FAC8&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   732
         TabIndex        =   2
         Top             =   1020
         Width           =   2244
      End
      Begin VB.TextBox txtStrOp 
         BackColor       =   &H00C8FAC8&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   732
         TabIndex        =   1
         Top             =   630
         Width           =   2244
      End
      Begin VB.TextBox txtStrInq 
         BackColor       =   &H00C8FAC8&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   732
         TabIndex        =   0
         Top             =   270
         Width           =   2244
      End
      Begin VB.Label Label15 
         Caption         =   "장기명"
         Height          =   210
         Left            =   90
         TabIndex        =   43
         Top             =   1110
         Width           =   630
      End
      Begin VB.Label Label14 
         Caption         =   "수술명"
         Height          =   210
         Left            =   90
         TabIndex        =   42
         Top             =   705
         Width           =   630
      End
      Begin VB.Label Label13 
         Caption         =   "진단명"
         Height          =   204
         Left            =   96
         TabIndex        =   41
         Top             =   348
         Width           =   636
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  '위 맞춤
      Height          =   705
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11880
      _Version        =   65536
      _ExtentX        =   20955
      _ExtentY        =   1244
      _StockProps     =   15
      Caption         =   "진   단   명   조   회"
      ForeColor       =   8388608
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   20.26
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Font3D          =   2
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "진     단     결     과"
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
      TabIndex        =   26
      Top             =   2190
      Width           =   8280
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "진        단        명"
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
      TabIndex        =   25
      Top             =   7440
      Width           =   8310
   End
End
Attribute VB_Name = "Anato_DiagName_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Dim i                       As Integer
Dim sRowID                  As String
Dim LsDiag                  As String

Dim SLIDENO
Dim LsElectro
Dim LsFlow

Dim LiOldRow                As Integer
Dim LsDeptCode(0 To 100)    As String * 2
Dim LsDeptName(0 To 100)    As String


Private Sub Form_Activate()
    
    txtStrInq.SetFocus

End Sub


Private Sub Form_Load()
    
'    dtFromDate.Value = Format(CDate(Dual_Date_Get("yyyy-MM-dd")) - 30, "yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    optG(0).Value = True
    
    Dim rs                  As ADODB.Recordset
    Dim i                   As Integer
    
    strSQL = ""
    strSQL = strSQL & " SELECT DeptCode,   DeptNameK  "
    strSQL = strSQL & " FROM   TWBAS_DEPT "
    strSQL = strSQL & " WHERE  GbJupsu =  '0'         "
    strSQL = strSQL & " ORDER  BY PrintRanking        " '


    Result = AdoOpenSet(rs, strSQL)

    If Result = False Then Exit Sub

    Do Until rs.EOF
        LsDeptCode(i) = rs.Fields("DeptCode").Value & ""
        LsDeptName(i) = rs.Fields("DeptNameK").Value & ""
        rs.MoveNext
        i = i + 1
    Loop
    AdoCloseSet rs
   
End Sub


Private Sub cmdClear_Click()
    
    optG(0).Value = True
    lblAnatoNo = ""
    lblPtno = ""
    lblSname = ""
    lblAge = ""
    lblDept = ""
    lblRoom = ""
    lblSex = ""
    txtDiag.Text = ""
    txtSummary.Text = ""
    
    txtStrInq.Text = ""
    txtStrOp.Text = ""
    txtStrOrgan.Text = ""
    
    Call SSInitialize(ssExDate)
    
    txtStrInq.SetFocus

End Sub


Private Sub cmdExit_Click()
    Unload Me

End Sub


Private Sub cmdPrint_Click()

    Dim i                   As Integer
    Dim LiPageCnt           As Integer
    Dim LiLineCnt           As Integer
    Dim LsAnatNo            As String * 11
    Dim LsPtNo              As String * 8
'    Dim LsSname             As String * 10
    Dim LsSname             As String
    Dim LsSex               As String * 1
    Dim LsAge               As String * 3
    Dim LsDeptName          As String * 18
    Dim LsRoomCode          As String * 6
    Dim LsDiagdate          As String * 10
    Dim LsChief             As String * 10
    Dim LsOpname            As String
    Dim LsOrganPart         As String
    
    Dim lsnumber
    
    Dim LsLineString        As String
    Dim LsBlank             As String * 1
    
    If ssExDate.DataRowCnt = 0 Then Exit Sub
    
    Anato_DiagName_Search.MousePointer = vbHourglass
    
    Printer.FontName = "굴림체"
'    Printer.FontName = "돋음체"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.FontItalic = False
    Printer.FontUnderline = False
    
    GoSub SUB_HEAD_PRINT
    
    For i = 1 To ssExDate.DataRowCnt
        ssExDate.Row = i
        ssExDate.Col = 1:        LsDiagdate = ssExDate.Text
        ssExDate.Col = 2:        LsAnatNo = ssExDate.Text
        ssExDate.Col = 3:        LsPtNo = ssExDate.Text
        ssExDate.Col = 4:        LsSname = RPadH(ssExDate.Text, 10)
        ssExDate.Col = 5:        LsSex = ssExDate.Text
        ssExDate.Col = 6:        LsAge = ssExDate.Text
        ssExDate.Col = 7:        LsDeptName = ssExDate.Text
        ssExDate.Col = 8:        LsRoomCode = ssExDate.Text
        ssExDate.Col = 9:        LsChief = ssExDate.Text
        ssExDate.Col = 10:       LsOpname = ssExDate.Text
        ssExDate.Col = 11:       LsOrganPart = ssExDate.Text
    
        LsBlank = " "
        LsLineString = ""
        LsLineString = LsLineString & lsnumber & LsBlank
        LsLineString = LsLineString & LsAnatNo & LsBlank
        LsLineString = LsLineString & LsPtNo & LsBlank
        LsLineString = LsLineString & LsSname & LsBlank
        LsLineString = LsLineString & LsSex & LsBlank
        LsLineString = LsLineString & LsAge & LsBlank
        LsLineString = LsLineString & LsDiagdate & LsBlank
        LsLineString = LsLineString & LsChief & LsBlank
        LsLineString = LsLineString & LsDeptName & LsBlank
        LsLineString = LsLineString & LsRoomCode
     
        Printer.Print LsLineString
        LsLineString = ""
        LiLineCnt = LiLineCnt + 1
        If LiLineCnt > 58 Then
            LiLineCnt = 0
            Printer.NewPage
            GoSub SUB_LINE_PRINT
            GoSub SUB_HEAD_PRINT
        End If
    Next i
    
    GoSub SUB_LINE_PRINT
    Printer.EndDoc
    
    Anato_DiagName_Search.MousePointer = vbDefault
    
    Exit Sub
    
SUB_HEAD_PRINT:
    Dim LsStrinq            As String * 20
    Dim LsPageCnt           As String * 3
    
    LiPageCnt = LiPageCnt + 1
    RSet LsPageCnt = LiPageCnt
    LsStrinq = txtStrInq
    
    LsLineString = " Case: " & LsStrinq & Space(59) & "Page: " & LsPageCnt
    Printer.Print LsLineString
    GoSub SUB_LINE_PRINT
'                            1         2         3         4         5         6         7         8         9
'                   123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456
'                   123 12345678901 12345678 1234567890 1 123 1234567890 1234567890 123456789012345678 123456 1234
    LsLineString = " NO 병록번호    환자번호 성명       S AGE 검사일자   검사자     진료과             병실   비고"
    Printer.Print LsLineString
    
    GoSub SUB_LINE_PRINT
    
    Return

SUB_LINE_PRINT:
    
    LsLineString = ""
    For j = 1 To 96
        LsLineString = LsLineString & "="
    Next j
    
    Printer.Print LsLineString
    
    Return

End Sub


Private Sub cmdStrInq_Click()
    '조회
    Dim rs                  As ADODB.Recordset
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k(12)               As Integer

    If Trim(txtStrInq.Text) = "" And Trim(txtStrOp.Text) = "" And Trim(txtStrOrgan.Text) = "" Then Exit Sub
    
    If Len(Trim(txtStrInq.Text)) = 1 Then
        MsgBox " 진단명에 2자 이상을 입력하십시요."
        Exit Sub
    End If
    
    If Len(Trim(txtStrOp.Text)) = 1 Then
        MsgBox " 수술명에 2자 이상을 입력하십시요."
        Exit Sub
    End If
    
    If Len(Trim(txtStrOrgan.Text)) = 1 Then
        MsgBox " 장기명에 2자 이상을 입력하십시요."
        Exit Sub
    End If
    
    Anato_DiagName_Search.MousePointer = vbHourglass
    
    lblAnatoNo = ""
    lblPtno = ""
    lblSname = ""
    lblAge = ""
    lblDept = ""
    lblRoom = ""
    lblSex = ""
    txtSummary.Text = ""
    txtDiag.Text = ""
    txtOpname.Text = ""
    txtOrganPart.Text = ""
    
    gSFrDate = Format(dtFromDate.Value, "yyyy-MM-dd")
    gSToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    Call SSInitialize(ssExDate)
     
    strSQL = ""
    strSQL = strSQL & " SELECT TO_CHAR(A.Diagdate, 'YYYY-MM-DD') Diagdate, "
    strSQL = strSQL & "        A.Class,  A.Dateyy, A.Seqnum,   A.Ptno,     A.Sname, "
    strSQL = strSQL & "        A.Sex,    A.AgeYY,  A.DeptCode, A.RoomCode, A.Chief, "
    strSQL = strSQL & "        A.OpName, A.ORGANPART, "  ' D.DXDICT
    strSQL = strSQL & "        B.DRName,   A.RowID "
    strSQL = strSQL & " FROM   TWANAT_DIAG A, TWBAS_DOCTOR B, TWBAS_ILLS C, TWANAT_DICT D "
    strSQL = strSQL & " WHERE  A.OPNAME    = C.ILLCODE(+) "
    strSQL = strSQL & " AND    A.Chief     = B.DRCODE(+) "
    strSQL = strSQL & " AND    A.ORGANPART = D.CODE(+) "
    strSQL = strSQL & " AND    A.DIAGDATE  BETWEEN TO_DATE('" & gSFrDate & "','YYYY-MM-DD') AND TO_DATE('" & gSToDate & "','YYYY-MM-DD')"
'    strSQL = strSQL & " AND    A.Chief     IS NOT NULL "
    If txtStrOp.Text = "" And txtStrOrgan.Text = "" Then
        strSQL = strSQL & "   AND   UPPER(A.Diagno)     Like '%" & UCase(Quot(Trim(txtStrInq.Text))) & "%'"
    ElseIf Trim(txtStrInq) = "" And txtStrOrgan.Text = "" Then
        strSQL = strSQL & "   AND   UPPER(C.ILLNAMEE)     Like '%" & UCase(Quot(Trim(txtStrOp.Text))) & "%'"
    ElseIf Trim(txtStrInq) = "" And txtStrOp.Text = "" Then
        strSQL = strSQL & "   AND   UPPER(D.DXDICT)  Like '%" & UCase(Quot(Trim(txtStrOrgan.Text))) & "%'"
    ElseIf Trim(txtStrInq) = "" Then
        strSQL = strSQL & "   AND   UPPER(C.ILLNAMEE)     Like '%" & UCase(Quot(Trim(txtStrOp.Text))) & "%'"
        strSQL = strSQL & "   AND   UPPER(D.DXDICT)  Like '%" & UCase(Quot(Trim(txtStrOrgan.Text))) & "%'"
    ElseIf txtStrOp.Text = "" Then
        strSQL = strSQL & "   AND   UPPER(A.Diagno)     Like '%" & UCase(Quot(Trim(txtStrInq.Text))) & "%'"
        strSQL = strSQL & "   AND   UPPER(D.DXDICT)  Like '%" & UCase(Quot(Trim(txtStrOrgan.Text))) & "%'"
    ElseIf txtStrOrgan.Text = "" Then
        strSQL = strSQL & "   AND   UPPER(A.Diagno)     Like '%" & UCase(Quot(Trim(txtStrInq.Text))) & "%'"
        strSQL = strSQL & "   AND   UPPER(C.ILLNAMEE)     Like '%" & UCase(Quot(Trim(txtStrOp.Text))) & "%'"
    Else
        strSQL = strSQL & "   AND   UPPER(A.Diagno)     Like '%" & UCase(Quot(Trim(txtStrInq.Text))) & "%'"
        strSQL = strSQL & "   AND   UPPER(C.ILLNAMEE)     Like '%" & UCase(Quot(Trim(txtStrOp.Text))) & "%'"
        strSQL = strSQL & "   AND   UPPER(D.DXDICT)  Like '%" & UCase(Quot(Trim(txtStrOrgan.Text))) & "%'"
    End If
    
    If optG(0).Value = True Then
        strSQL = strSQL & " AND  a.Spegeomch = '1'"
    ElseIf optG(1).Value = True Then
        strSQL = strSQL & " AND  a.Spegeomch = '2'"
    End If
    
    strSQL = strSQL & " ORDER  BY DiagDate DESC, CLASS DESC, DATEYY DESC, SEQNUM DESC "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    ssExDate.MaxRows = Rowindicator + 1

    ssExDate.Row = 0
    ssExDate.Col = 1:     ssExDate.Text = "Diagdate"
    ssExDate.Col = 2:     ssExDate.Text = "접수NO"
    ssExDate.Col = 3:     ssExDate.Text = "Ptno"
    ssExDate.Col = 4:     ssExDate.Text = "Sname"
    ssExDate.Col = 5:     ssExDate.Text = "Sex"
    ssExDate.Col = 6:     ssExDate.Text = "AgeYY"
    ssExDate.Col = 7:     ssExDate.Text = "Deptcode"
    ssExDate.Col = 8:     ssExDate.Text = "Roomcode"
    ssExDate.Col = 9:     ssExDate.Text = "Name"
    ssExDate.Col = 10:    ssExDate.Text = "OpName"
    ssExDate.Col = 11:    ssExDate.Text = "OrganPart"
    
    For i = 1 To 11 '9
        ssExDate.Row = 0
        ssExDate.Col = i
        If LenH(ssExDate.Text) > k(i) Then
            ssExDate.ColWidth(i) = LenH(ssExDate.Text)
            k(i) = LenH(ssExDate.Text)
        Else
            ssExDate.ColWidth(i) = k(i)
        End If
    Next i
    
    
    i = 0
    Do Until rs.EOF
        ssExDate.Row = i + 1
        ssExDate.Col = 1: ssExDate.Text = rs.Fields("DIAGDATE").Value & ""
        If LenH(ssExDate.Text) > k(1) Then
            ssExDate.ColWidth(1) = LenH(ssExDate.Text)
            k(1) = LenH(ssExDate.Text)
        Else
            ssExDate.ColWidth(1) = k(1)
        End If
        
        ssExDate.Col = 2: ssExDate.Text = rs.Fields("Class").Value & "-" & _
                                          rs.Fields("DateYY").Value & "-" & _
                                          rs.Fields("SeqNum").Value & ""
        If LenH(ssExDate.Text) > k(2) Then
            ssExDate.ColWidth(2) = LenH(ssExDate.Text)
            k(2) = LenH(ssExDate.Text)
        Else
            ssExDate.ColWidth(2) = k(2)
        End If
        
        ssExDate.Col = 3: ssExDate.Text = rs.Fields("Ptno").Value & ""
        If LenH(ssExDate.Text) > k(3) Then
            ssExDate.ColWidth(1) = LenH(ssExDate.Text)
            k(3) = LenH(ssExDate.Text)
        Else
            ssExDate.ColWidth(3) = k(3)
        End If
        
        ssExDate.Col = 4: ssExDate.Text = rs.Fields("Sname").Value & ""
        If LenH(ssExDate.Text) > k(4) Then
            ssExDate.ColWidth(4) = LenH(ssExDate.Text)
            k(4) = LenH(ssExDate.Text)
        Else
            ssExDate.ColWidth(4) = k(4)
        End If
        
        ssExDate.Col = 5: ssExDate.Text = rs.Fields("Sex").Value & ""
        If LenH(ssExDate.Text) > k(5) Then
            ssExDate.ColWidth(5) = LenH(ssExDate.Text)
            k(5) = LenH(ssExDate.Text)
        Else
            ssExDate.ColWidth(5) = k(5)
        End If
        
        ssExDate.Col = 6: ssExDate.Text = rs.Fields("AgeYY").Value & ""
        If LenH(ssExDate.Text) > k(6) Then
            ssExDate.ColWidth(6) = LenH(ssExDate.Text)
            k(6) = LenH(ssExDate.Text)
        Else
            ssExDate.ColWidth(6) = k(6)
        End If
        
        ssExDate.Col = 7:
            For j = 0 To 100
                If Trim(rs.Fields("Deptcode").Value & "") = LsDeptCode(j) Then
                    ssExDate.Text = LsDeptName(j)
                    Exit For
                End If
            Next j
        If LenH(ssExDate.Text) > k(7) Then
            ssExDate.ColWidth(7) = LenH(ssExDate.Text)
            k(7) = LenH(ssExDate.Text)
        Else
            ssExDate.ColWidth(7) = k(7)
        End If
        
        ssExDate.Col = 8: ssExDate.Text = rs.Fields("Roomcode").Value & ""
        If LenH(ssExDate.Text) > k(8) Then
            ssExDate.ColWidth(8) = LenH(ssExDate.Text)
            k(8) = LenH(ssExDate.Text)
        Else
            ssExDate.ColWidth(8) = k(8)
        End If
        
        ssExDate.Col = 9: ssExDate.Text = rs.Fields("DRName").Value & ""
        If LenH(ssExDate.Text) > k(9) Then
            ssExDate.ColWidth(9) = LenH(ssExDate.Text)
            k(9) = LenH(ssExDate.Text)
        Else
            ssExDate.ColWidth(9) = k(9)
        End If
        
        If (rs.Fields("DRName").Value & "") = "" Then
            ssExDate.Col = 9: ssExDate.Text = rs.Fields("Chief").Value & ""
        End If
        
        ssExDate.Col = 10: ssExDate.Text = rs.Fields("Opname").Value & ""
        If LenH(ssExDate.Text) > k(10) Then
            ssExDate.ColWidth(10) = LenH(ssExDate.Text)
            k(10) = LenH(ssExDate.Text)
        Else
            ssExDate.ColWidth(10) = k(10)
        End If
        
        ssExDate.Col = 11: ssExDate.Text = rs.Fields("ORGANPART").Value & ""
        If LenH(ssExDate.Text) > k(11) Then
            ssExDate.ColWidth(11) = LenH(ssExDate.Text)
            k(11) = LenH(ssExDate.Text)
        Else
            ssExDate.ColWidth(10) = k(10)
        End If
        
        ssExDate.Col = 12: ssExDate.Text = rs.Fields("RowID").Value & ""
        
        rs.MoveNext: i = i + 1
    Loop
    
    AdoCloseSet rs
    
'    ssExDate.BlockMode = False
    ssExDate.MaxRows = ssExDate.DataRowCnt + 1
        
    Anato_DiagName_Search.MousePointer = vbDefault
    
    
End Sub


Private Sub cmdStrInq_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub



Private Sub ssExDate_Click(ByVal Col As Long, ByVal Row As Long)
    
    
    If Row > ssExDate.DataRowCnt Then Exit Sub
    
    For i = 1 To 11   '9
        ssExDate.Col = i    '1
        ssExDate.Row = LiOldRow
        ssExDate.BackColor = RGB(235, 245, 235)
        ssExDate.ForeColor = RGB(0, 0, 0)
    Next i
    
    For i = 1 To 11     '9
        ssExDate.Col = i    '1
        ssExDate.Row = Row
        ssExDate.BackColor = RGB(0, 0, 128)
        ssExDate.ForeColor = RGB(255, 255, 255)
        LiOldRow = Row
    Next i
    
'            .SS4.ColWidth(3) = 5
'    ssExDate.ColWidth
    
    ssExDate.Col = 12 '10
    ssExDate.Row = Row: sRowID = ssExDate.Text
    
    strSQL = " SELECT * FROM TWANAT_Diag WHERE RowID = '" & sRowID & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    txtDiag.Text = ""
    txtSummary.Text = ""
    
    lblAnatoNo = rs.Fields("Class").Value & "-" & _
                 rs.Fields("DateYY").Value & "-" & _
                 rs.Fields("Seqnum").Value & ""
    lblPtno = rs.Fields("Ptno").Value & ""
    lblSname = rs.Fields("Sname").Value & ""
    
    lblSex = IIf(rs.Fields("Sex").Value & "" = "M", "남자", "여자")
    lblAge = rs.Fields("Ageyy").Value '& ""
    
    For i = 0 To 100
        If Trim(rs.Fields("DeptCode").Value & "") = LsDeptCode(i) Then
            lblDept = LsDeptName(i)
            Exit For
        End If
    Next i
    lblRoom = rs.Fields("RoomCode").Value & ""
    
'    txtDiag.Text = rs.Fields("Descr").Value & ""
    txtDiag.Text = (rs.Fields("DiagEye").Value & "") & (rs.Fields("DiagPre").Value & "") & (rs.Fields("Descr").Value & "") & (rs.Fields("DiagAdd").Value & "")
    
'    txtSummary.Text = rs.Fields("Diagno").Value & ""
    txtSummary.Text = (rs.Fields("Diagcode").Value & "") & " : " & DiagCodeSearch(rs.Fields("Diagcode").Value & "")
    
    txtOpname.Text = rs.Fields("Opname").Value & ""
    txtOrganPart.Text = rs.Fields("OrganPart").Value & ""
    
    AdoCloseSet rs
    
    Call Special_result
    
    txtDiag.Text = txtDiag.Text & vbCrLf & LsDiag
    
    
End Sub


Private Sub txtStrInq_GotFocus()
    txtStrInq.SelStart = 0
    txtStrInq.SelLength = Len(txtStrInq.Text)

End Sub

Private Sub txtStrInq_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub txtStrInq_LostFocus()
    txtStrInq.Text = UCase(txtStrInq.Text)

End Sub

Private Sub txtStrOp_GotFocus()
    txtStrOp.SelStart = 0
    txtStrOp.SelLength = Len(txtStrOp.Text)

End Sub


Private Sub txtStrOp_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub

Private Sub txtStrOp_LostFocus()
    txtStrOp.Text = UCase(txtStrOp.Text)

End Sub

Private Sub txtStrOrgan_GotFocus()
    txtStrOrgan.SelStart = 0
    txtStrOrgan.SelLength = Len(txtStrOrgan.Text)

End Sub

Private Sub txtStrOrgan_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub

Private Sub txtStrOrgan_LostFocus()
    txtStrOrgan.Text = UCase(txtStrOrgan.Text)

End Sub


Private Sub Special_result()

    Dim SpecialChar         As String
    Dim sSpecial(30)        As String
    Dim lsnumber


    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & " FROM   TWANAT_DIAG "
    strSQL = strSQL & " WHERE RowID = '" & sRowID & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    If Result Then
        Do Until rs.EOF
            For j = 1 To 30
                SpecialChar = "SPECIAL" & Format(j, "00")
                sSpecial(j) = Trim(rs.Fields(SpecialChar).Value & "")
            Next j
            
            SLIDENO = Trim(rs.Fields("slid").Value & "")
            
            LsElectro = Trim(rs.Fields("ELECTROSCOPE").Value & "")
            LsFlow = Trim(rs.Fields("FLOW").Value & "")
            
            rs.MoveNext
        Loop
    End If
    
    
    Dim FlagSpecial1
    Dim FlagSpecial2
    Dim FlagSpecial3
    Dim FlagSpecial4
    Dim FlagSpecial5
    
    Dim Flag_Data1
    Dim Flag_Data2
    Dim Flag_Data3
    Dim Flag_Data4
    Dim Flag_Data5
    
    
    For j = 1 To 30
        Select Case sSpecial(j)
                Case "853001" To "853999"
                     FlagSpecial1 = 1
                     Flag_Data1 = Flag_Data1 & Special_Load(sSpecial(j)) & ", "
                Case "857001" To "857999"
                     FlagSpecial2 = 1
                     Flag_Data2 = Flag_Data2 & Special_Load(sSpecial(j)) & ", "
                Case "854001" To "854999"
                     FlagSpecial3 = 1
                     Flag_Data3 = Flag_Data3 & Special_Load(sSpecial(j)) & ", "
                Case "856001" To "856999"
                     FlagSpecial4 = 1
                     Flag_Data4 = Flag_Data4 & Special_Load(sSpecial(j)) & ", "
                Case "855001"
                     FlagSpecial5 = 1
                     Flag_Data5 = "Y "
        End Select
    Next j
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If FlagSpecial1 = 1 Then
         If Len(Flag_Data1) <= 60 Then
             LsDiag = LsDiag & "SPECIAL STAIN : " & Mid(Flag_Data1, 1, Len(Flag_Data1) - 2) & vbCrLf
         
         ElseIf Len(Flag_Data1) > 60 And Len(Flag_Data1) <= 120 Then
             LsDiag = LsDiag & "SPECIAL STAIN : " & Mid(Flag_Data1, 1, 60) & vbCrLf
             LsDiag = LsDiag & "              : " & Mid(Flag_Data1, 61, Len(Flag_Data1) - 62) & vbCrLf
         
         ElseIf Len(Flag_Data1) > 120 And Len(Flag_Data1) <= 180 Then
             LsDiag = LsDiag & "SPECIAL STAIN : " & Mid(Flag_Data1, 1, 60) & vbCrLf
             LsDiag = LsDiag & "              : " & Mid(Flag_Data1, 61, 60) & vbCrLf
             LsDiag = LsDiag & "              : " & Mid(Flag_Data1, 121, Len(Flag_Data1) - 122) & vbCrLf
         
         ElseIf Len(Flag_Data1) > 181 Then
             LsDiag = LsDiag & "SPECIAL STAIN : " & Mid(Flag_Data1, 1, 60) & vbCrLf
             LsDiag = LsDiag & "              : " & Mid(Flag_Data1, 61, 60) & vbCrLf
             LsDiag = LsDiag & "              : " & Mid(Flag_Data1, 121, 60) & vbCrLf
             LsDiag = LsDiag & "              : " & Mid(Flag_Data1, 181, Len(Flag_Data1) - 182) & vbCrLf
         End If
    End If
    
    If FlagSpecial2 = 1 Then
        If Len(Flag_Data2) <= 60 Then
             LsDiag = LsDiag & "IMMUNOHISTOCHEMICAL STAIN : " & Mid(Flag_Data2, 1, Len(Flag_Data2) - 2) & vbCrLf
        
        ElseIf Len(Flag_Data2) > 60 And Len(Flag_Data1) <= 120 Then
             LsDiag = LsDiag & "IMMUNOHISTOCHEMICAL STAIN : " & Mid(Flag_Data2, 1, 60) & vbCrLf
             LsDiag = LsDiag & "                          : " & Mid(Flag_Data2, 61, Len(Flag_Data2) - 62) & vbCrLf
        
        ElseIf Len(Flag_Data2) > 120 And Len(Flag_Data1) <= 180 Then
             LsDiag = LsDiag & "IMMUNOHISTOCHEMICAL STAIN : " & Mid(Flag_Data2, 1, 60) & vbCrLf
             LsDiag = LsDiag & "                          : " & Mid(Flag_Data2, 61, 60) & vbCrLf
             LsDiag = LsDiag & "                          : " & Mid(Flag_Data2, 121, Len(Flag_Data2) - 122) & vbCrLf
        
        ElseIf Len(Flag_Data2) > 180 Then
             LsDiag = LsDiag & "Immunohistochemical stain : " & Mid(Flag_Data2, 1, 60) & vbCrLf
             LsDiag = LsDiag & "                          : " & Mid(Flag_Data2, 61, 60) & vbCrLf
             LsDiag = LsDiag & "                          : " & Mid(Flag_Data2, 121, 60) & vbCrLf
             LsDiag = LsDiag & "                          : " & Mid(Flag_Data2, 181, Len(Flag_Data2) - 182) & vbCrLf
        End If
         
    End If
    If FlagSpecial3 = 1 Then
         LsDiag = LsDiag & "IMMUNOFLUORESCENCE : " & Mid(Flag_Data3, 1, Len(Flag_Data3) - 2) & vbCrLf
    End If
    If FlagSpecial4 = 1 Then
         LsDiag = LsDiag & "ENZYME HISTOCHEMISTRY : " & Mid(Flag_Data4, 1, Len(Flag_Data4) - 2) & vbCrLf
    End If
    
    If LsElectro = "Y" Then
         LsDiag = LsDiag & "ELECTRON MICROSCOPIC EXAM : Y " & vbCrLf
    End If
    
    If LsFlow = "Y" Then
         LsDiag = LsDiag & "FLOW CYTOMETRY : Y " & vbCrLf
    End If
    
    AdoCloseSet rs
    
End Sub


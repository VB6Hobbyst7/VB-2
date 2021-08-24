VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWorkList 
   Caption         =   "의뢰환자 WorkList"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13365
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ListBox lstExam 
      BackColor       =   &H00FFECF8&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7035
      ItemData        =   "frmWorkList.frx":0000
      Left            =   0
      List            =   "frmWorkList.frx":0002
      TabIndex        =   2
      Top             =   2370
      Width           =   3105
   End
   Begin MSComCtl2.MonthView monvCal 
      Height          =   2220
      Left            =   1380
      TabIndex        =   0
      Top             =   930
      Visible         =   0   'False
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   23724033
      CurrentDate     =   37043
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7080
      Left            =   3120
      TabIndex        =   1
      Top             =   2370
      Width           =   10200
      _Version        =   196608
      _ExtentX        =   17992
      _ExtentY        =   12488
      _StockProps     =   64
      ColHeaderDisplay=   0
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   17
      MaxRows         =   25
      RowHeaderDisplay=   0
      ScrollBars      =   2
      SpreadDesigner  =   "frmWorkList.frx":0004
      UserResize      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1530
      Top             =   3810
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2295
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13305
      _Version        =   65536
      _ExtentX        =   23469
      _ExtentY        =   4048
      _StockProps     =   14
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
      Begin VB.ComboBox cboSlip 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1380
         TabIndex        =   23
         Top             =   1125
         Width           =   2355
      End
      Begin VB.CheckBox chkEquip 
         Caption         =   "장비검사 포함"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   450
         Left            =   3795
         TabIndex        =   22
         Top             =   1072
         Value           =   1  '확인
         Width           =   1830
      End
      Begin VB.Frame Frame1 
         Height          =   510
         Left            =   9045
         TabIndex        =   18
         Top             =   105
         Width           =   2700
         Begin VB.OptionButton optSta 
            Caption         =   "모두"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   21
            Top             =   135
            Width           =   870
         End
         Begin VB.OptionButton optSta 
            Caption         =   "접수"
            Height          =   285
            Index           =   1
            Left            =   945
            TabIndex        =   20
            Top             =   135
            Value           =   -1  'True
            Width           =   870
         End
         Begin VB.OptionButton optSta 
            Caption         =   "결과"
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   19
            Top             =   135
            Width           =   870
         End
      End
      Begin VB.CheckBox chkLine 
         Caption         =   "1줄"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5760
         TabIndex        =   17
         Top             =   540
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CheckBox chkPrinter 
         Caption         =   "기본프린터"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9960
         TabIndex        =   16
         Top             =   1260
         Value           =   1  '확인
         Width           =   1680
      End
      Begin VB.ComboBox cboLine 
         Height          =   315
         ItemData        =   "frmWorkList.frx":0878
         Left            =   9015
         List            =   "frmWorkList.frx":0885
         TabIndex        =   15
         Top             =   1245
         Width           =   885
      End
      Begin VB.CheckBox chkProfile 
         Caption         =   "프로파일"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   450
         Left            =   5700
         TabIndex        =   14
         Top             =   1072
         Value           =   1  '확인
         Width           =   1260
      End
      Begin VB.Frame Frame2 
         Height          =   510
         Left            =   9030
         TabIndex        =   10
         Top             =   630
         Width           =   2730
         Begin VB.OptionButton optEmg 
            Caption         =   "일반"
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   13
            Top             =   135
            Width           =   870
         End
         Begin VB.OptionButton optEmg 
            Caption         =   "응급"
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   1
            Left            =   945
            TabIndex        =   12
            Top             =   135
            Width           =   870
         End
         Begin VB.OptionButton optEmg 
            Caption         =   "모두"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   11
            Top             =   135
            Value           =   -1  'True
            Width           =   870
         End
      End
      Begin VB.TextBox txtExam 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   1380
         TabIndex        =   9
         Top             =   1590
         Width           =   1155
      End
      Begin VB.CommandButton cmdCalendar 
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2655
         TabIndex        =   8
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdCalendar 
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2655
         TabIndex        =   7
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox txtEnd 
         Appearance      =   0  '평면
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
         Left            =   3030
         TabIndex        =   6
         Top             =   645
         Width           =   960
      End
      Begin VB.TextBox txtStr 
         Appearance      =   0  '평면
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
         Left            =   3030
         TabIndex        =   5
         Top             =   255
         Width           =   960
      End
      Begin VB.CheckBox chkRemark 
         Caption         =   "비고란"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   11910
         TabIndex        =   4
         Top             =   810
         Value           =   1  '확인
         Width           =   1050
      End
      Begin Threed.SSCommand sscSearch 
         Height          =   525
         Left            =   7800
         TabIndex        =   24
         Top             =   1680
         Width           =   1380
         _Version        =   65536
         _ExtentX        =   2434
         _ExtentY        =   926
         _StockProps     =   78
         Caption         =   "조 회"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel sspItem 
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   25
         Top             =   1140
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "검사파트"
         ForeColor       =   0
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.26
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand sscPrint 
         Height          =   525
         Left            =   9210
         TabIndex        =   26
         Top             =   1680
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   926
         _StockProps     =   78
         Caption         =   "출 력"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand sscClose 
         Height          =   525
         Left            =   10575
         TabIndex        =   27
         Top             =   1680
         Width           =   1380
         _Version        =   65536
         _ExtentX        =   2434
         _ExtentY        =   926
         _StockProps     =   78
         Caption         =   "종 료"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel sspItem 
         Height          =   315
         Index           =   3
         Left            =   7800
         TabIndex        =   28
         Top             =   240
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "진행상태"
         ForeColor       =   0
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel sspItem 
         Height          =   315
         Index           =   4
         Left            =   7800
         TabIndex        =   29
         Top             =   1245
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "LINE"
         ForeColor       =   0
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel sspItem 
         Height          =   315
         Index           =   5
         Left            =   7800
         TabIndex        =   30
         Top             =   750
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "접수형태"
         ForeColor       =   0
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel sspItem 
         Height          =   315
         Index           =   6
         Left            =   150
         TabIndex        =   31
         Top             =   1590
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "검사항목"
         ForeColor       =   0
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.26
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel sspExam 
         Height          =   300
         Left            =   2490
         TabIndex        =   32
         Top             =   1605
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   529
         _StockProps     =   15
         BackColor       =   16772344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   1
      End
      Begin MSMask.MaskEdBox mskReqDate 
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   33
         Top             =   240
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel sspItem 
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   34
         Top             =   240
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "접수범위"
         ForeColor       =   0
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.26
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox mskReqDate 
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   35
         Top             =   630
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   900
         TabIndex        =   36
         Top             =   660
         Width           =   195
      End
   End
   Begin FPSpread.vaSpread vasList_1 
      Height          =   7080
      Left            =   3120
      TabIndex        =   37
      Top             =   2370
      Width           =   10200
      _Version        =   196608
      _ExtentX        =   17992
      _ExtentY        =   12488
      _StockProps     =   64
      ColHeaderDisplay=   0
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   14
      MaxRows         =   25
      RowHeaderDisplay=   0
      ScrollBars      =   2
      SpreadDesigner  =   "frmWorkList.frx":089B
      UserResize      =   1
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub


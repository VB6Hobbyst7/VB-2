VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGJ0101 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  '없음
   Caption         =   "환자 접수"
   ClientHeight    =   7950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11925
   ForeColor       =   &H00400000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dtpLabDate 
      Height          =   315
      Left            =   1230
      TabIndex        =   0
      Top             =   300
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24641539
      CurrentDate     =   36165
   End
   Begin Threed.SSFrame frabasic 
      Height          =   1545
      Left            =   30
      TabIndex        =   22
      Top             =   60
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   2725
      _StockProps     =   14
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtSpcCd 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1020
         Width           =   495
      End
      Begin VB.TextBox txtslipcd 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   1
         Top             =   630
         Width           =   495
      End
      Begin Threed.SSPanel pnlLabDate 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   210
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "접수일자"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlSlipcd 
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   600
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "SLIP"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlSpcCd 
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   990
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "검체"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSCommand cmdsliph 
         Height          =   330
         Left            =   1710
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   630
         Width           =   270
         _Version        =   65536
         _ExtentX        =   476
         _ExtentY        =   582
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "FGJ0101.frx":0000
      End
      Begin Threed.SSCommand cmdspch 
         Height          =   330
         Left            =   1710
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   1020
         Width           =   270
         _Version        =   65536
         _ExtentX        =   476
         _ExtentY        =   582
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "FGJ0101.frx":0122
      End
      Begin VB.Label lblSpcNm 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2010
         TabIndex        =   65
         Top             =   1020
         Width           =   1965
      End
      Begin VB.Label lblSlipNm 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2010
         TabIndex        =   64
         Top             =   630
         Width           =   1965
      End
   End
   Begin Threed.SSFrame fraPerson 
      Height          =   1935
      Left            =   30
      TabIndex        =   24
      Top             =   1530
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   3413
      _StockProps     =   14
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtsex 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3195
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1440
         Width           =   765
      End
      Begin VB.TextBox txtage 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1215
         MaxLength       =   6
         TabIndex        =   7
         Top             =   1440
         Width           =   765
      End
      Begin VB.TextBox txtidright 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         MaxLength       =   7
         TabIndex        =   6
         Top             =   1050
         Width           =   915
      End
      Begin VB.TextBox txtidleft 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1215
         MaxLength       =   6
         TabIndex        =   5
         Top             =   1050
         Width           =   765
      End
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   10  '한글 
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   4
         Top             =   660
         Width           =   2775
      End
      Begin VB.TextBox txtRegNo 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   1995
      End
      Begin Threed.SSPanel pnlregno 
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "등록번호"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlname 
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   630
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "이     름"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlid 
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1020
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "주민번호"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel pnlage 
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   1410
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "나     이"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel pnlsex 
         Height          =   375
         Left            =   2250
         TabIndex        =   46
         Top             =   1410
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "성     별"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin VB.Line Line1 
         X1              =   2070
         X2              =   2160
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin Threed.SSFrame frainfo 
      Height          =   4485
      Left            =   30
      TabIndex        =   28
      Top             =   3390
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   7911
      _StockProps     =   14
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtDr 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   10  '한글 
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   12
         Top             =   1800
         Width           =   1875
      End
      Begin VB.TextBox txtcmt 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         IMEMode         =   10  '한글 
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   14
         Top             =   2880
         Width           =   3855
      End
      Begin VB.CheckBox chkspecial 
         Caption         =   "특진이면 체크"
         Height          =   315
         Left            =   2370
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2190
         Width           =   1485
      End
      Begin VB.TextBox txtspecial 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         MaxLength       =   1
         ScrollBars      =   2  '수직
         TabIndex        =   13
         Text            =   "N"
         Top             =   2190
         Width           =   285
      End
      Begin VB.CheckBox chkem 
         Caption         =   "응급이면 체크"
         Height          =   315
         Left            =   2370
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1410
         Width           =   1485
      End
      Begin VB.TextBox txtem 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         MaxLength       =   1
         ScrollBars      =   2  '수직
         TabIndex        =   11
         Text            =   "N"
         Top             =   1410
         Width           =   285
      End
      Begin VB.TextBox txtward 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   10  '한글 
         Left            =   1200
         MaxLength       =   20
         ScrollBars      =   2  '수직
         TabIndex        =   10
         Top             =   615
         Width           =   2535
      End
      Begin VB.TextBox txtdeptcd 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         MaxLength       =   2
         ScrollBars      =   2  '수직
         TabIndex        =   9
         Top             =   225
         Width           =   375
      End
      Begin Threed.SSPanel pnldept 
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   210
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "진료과"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlward 
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "병실(병동)"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlgubun 
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   990
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "접수구분"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlDr 
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   1770
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Doctor"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlem 
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   1380
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "응급구분"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlspecial 
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   2160
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "특진체크"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlcmt 
         Height          =   345
         Left            =   120
         TabIndex        =   41
         Top             =   2550
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "병명기호(임상소견)"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSFrame fragubun 
         Height          =   405
         Left            =   1200
         TabIndex        =   42
         Top             =   930
         Width           =   2805
         _Version        =   65536
         _ExtentX        =   4948
         _ExtentY        =   714
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
         Begin VB.OptionButton optgubun 
            Caption         =   "수탁"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   1860
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   150
            Width           =   705
         End
         Begin VB.OptionButton optgubun 
            Caption         =   "입원"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   990
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   150
            Width           =   705
         End
         Begin VB.OptionButton optgubun 
            Caption         =   "외래"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   150
            Width           =   705
         End
      End
      Begin Threed.SSPanel pnlLLabNo 
         Height          =   555
         Left            =   60
         TabIndex        =   53
         Top             =   3870
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   979
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
         BevelInner      =   1
         Enabled         =   0   'False
         Begin Threed.SSPanel pnlDLabNo 
            Height          =   375
            Left            =   90
            TabIndex        =   54
            Top             =   90
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "작업번호"
            ForeColor       =   8454143
            BackColor       =   128
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   1
         End
         Begin VB.Label lblLLabseq 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2760
            TabIndex        =   73
            Top             =   120
            Width           =   705
         End
         Begin VB.Label lblLSlipCd 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2250
            TabIndex        =   72
            Top             =   120
            Width           =   525
         End
         Begin VB.Label lblLLabdate 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1170
            TabIndex        =   55
            Top             =   120
            Width           =   1095
         End
      End
      Begin Threed.SSCommand cmddepth 
         Height          =   330
         Left            =   1590
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   210
         Width           =   270
         _Version        =   65536
         _ExtentX        =   476
         _ExtentY        =   582
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "FGJ0101.frx":0244
      End
      Begin VB.Label lbldeptnm 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         TabIndex        =   85
         Top             =   210
         Width           =   2085
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Y/N  or"
         Height          =   180
         Left            =   1590
         TabIndex        =   44
         Top             =   2250
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Y/N  or"
         Height          =   180
         Left            =   1590
         TabIndex        =   43
         Top             =   1470
         Width           =   630
      End
   End
   Begin Threed.SSFrame SSFrame6 
      Height          =   7815
      Left            =   8220
      TabIndex        =   48
      Top             =   60
      Width           =   3675
      _Version        =   65536
      _ExtentX        =   6482
      _ExtentY        =   13785
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
      Begin FPSpread.vaSpread spdsLabNo 
         Height          =   3735
         Left            =   120
         OleObjectBlob   =   "FGJ0101.frx":0366
         TabIndex        =   56
         Top             =   4020
         Width           =   3435
      End
      Begin Threed.SSFrame frasearch 
         Height          =   1515
         Left            =   120
         TabIndex        =   68
         Top             =   1470
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   2672
         _StockProps     =   14
         Caption         =   "접수내역 조회 Option ( 해당일 )"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optsearch 
            Appearance      =   0  '평면
            BackColor       =   &H00C0E0FF&
            Caption         =   "접수 완/미완 List"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   660
            TabIndex        =   33
            Top             =   1080
            Width           =   2115
         End
         Begin VB.OptionButton optsearch 
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            Caption         =   "등록번호"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   660
            TabIndex        =   32
            Top             =   690
            Width           =   2115
         End
         Begin VB.OptionButton optsearch 
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFC0&
            Caption         =   "작업번호"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   660
            TabIndex        =   31
            Top             =   300
            Value           =   -1  'True
            Width           =   2115
         End
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   1005
         Left            =   1260
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   330
         Width           =   1125
         _Version        =   65536
         _ExtentX        =   1976
         _ExtentY        =   1773
         _StockProps     =   78
         Caption         =   "삭제 F4"
         ForeColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Picture         =   "FGJ0101.frx":0673
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   1005
         Left            =   2400
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   330
         Width           =   1125
         _Version        =   65536
         _ExtentX        =   1984
         _ExtentY        =   1773
         _StockProps     =   78
         Caption         =   "종료 ESC"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Picture         =   "FGJ0101.frx":0F4D
      End
      Begin Threed.SSCommand cmdAppend 
         Height          =   1005
         Left            =   120
         TabIndex        =   17
         Top             =   330
         Width           =   1125
         _Version        =   65536
         _ExtentX        =   1976
         _ExtentY        =   1773
         _StockProps     =   78
         Caption         =   "등록 F2"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Picture         =   "FGJ0101.frx":1827
      End
      Begin Threed.SSPanel pnlslabno 
         Height          =   945
         Left            =   120
         TabIndex        =   74
         Top             =   3000
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   1667
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
         Begin VB.TextBox txtslabdate 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   540
            MaxLength       =   8
            TabIndex        =   34
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtsslipcd 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1530
            MaxLength       =   3
            TabIndex        =   35
            Top             =   540
            Width           =   465
         End
         Begin VB.TextBox txtslabseq 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2280
            MaxLength       =   5
            TabIndex        =   36
            Top             =   540
            Width           =   645
         End
         Begin Threed.SSPanel pnltitlelabno 
            Height          =   315
            Left            =   540
            TabIndex        =   81
            Top             =   180
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "작업번호"
            ForeColor       =   0
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
         End
         Begin Threed.SSCommand cmdlabnoh 
            Height          =   330
            Left            =   2010
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   540
            Width           =   270
            _Version        =   65536
            _ExtentX        =   476
            _ExtentY        =   582
            _StockProps     =   78
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            RoundedCorners  =   0   'False
            Picture         =   "FGJ0101.frx":2101
         End
      End
      Begin Threed.SSPanel pnlsRegno 
         Height          =   945
         Left            =   120
         TabIndex        =   69
         Top             =   3000
         Visible         =   0   'False
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   1667
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
         Begin Threed.SSPanel pnls1Regno 
            Height          =   315
            Left            =   540
            TabIndex        =   71
            Top             =   180
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "등록번호"
            ForeColor       =   0
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   1
         End
         Begin VB.TextBox txtsRegno 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   540
            MaxLength       =   15
            TabIndex        =   70
            Top             =   540
            Width           =   1635
         End
      End
      Begin Threed.SSPanel pnlsokres 
         Height          =   945
         Left            =   120
         TabIndex        =   75
         Top             =   3000
         Visible         =   0   'False
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   1667
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
         Begin VB.OptionButton optResOK 
            Caption         =   "미완료"
            Height          =   225
            Index           =   1
            Left            =   1260
            TabIndex        =   79
            Top             =   540
            Width           =   915
         End
         Begin VB.OptionButton optResOK 
            Caption         =   "완료"
            Height          =   225
            Index           =   0
            Left            =   540
            TabIndex        =   77
            Top             =   540
            Width           =   765
         End
         Begin Threed.SSPanel pnlsResOK 
            Height          =   315
            Left            =   540
            TabIndex        =   76
            Top             =   180
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "접수 완/미완 List"
            ForeColor       =   0
            BackColor       =   12640511
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   1
         End
      End
   End
   Begin Threed.SSFrame fraorder 
      Height          =   7815
      Left            =   4140
      TabIndex        =   47
      Top             =   60
      Width           =   4065
      _Version        =   65536
      _ExtentX        =   7170
      _ExtentY        =   13785
      _StockProps     =   14
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin FPSpread.vaSpread spdRoutine 
         Height          =   1725
         Left            =   120
         OleObjectBlob   =   "FGJ0101.frx":2223
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   585
         Width           =   3825
      End
      Begin FPSpread.vaSpread spdorder 
         Height          =   4575
         Left            =   120
         OleObjectBlob   =   "FGJ0101.frx":32C3
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2730
         Width           =   3825
      End
      Begin VB.TextBox txtospccd 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   67
         Top             =   2370
         Width           =   435
      End
      Begin VB.TextBox txtoslipcd 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   66
         Top             =   2370
         Width           =   435
      End
      Begin VB.TextBox txtrpartcd 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   59
         Top             =   225
         Width           =   255
      End
      Begin VB.TextBox txtoordcd 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   16
         Top             =   2370
         Width           =   435
      End
      Begin VB.TextBox txtroutinecd 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1590
         MaxLength       =   3
         TabIndex        =   15
         Top             =   225
         Width           =   465
      End
      Begin Threed.SSPanel pnlRoutine 
         Height          =   360
         Left            =   120
         TabIndex        =   51
         Top             =   210
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "ROUTINE"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSCommand cmdroutineh 
         Height          =   315
         Left            =   2070
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   225
         Width           =   270
         _Version        =   65536
         _ExtentX        =   476
         _ExtentY        =   556
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "FGJ0101.frx":43B9
      End
      Begin Threed.SSPanel pnlorder 
         Height          =   360
         Left            =   120
         TabIndex        =   57
         Top             =   2340
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "검사항목코드"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSCommand cmdordh 
         Height          =   315
         Left            =   2850
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   2370
         Width           =   270
         _Version        =   65536
         _ExtentX        =   476
         _ExtentY        =   556
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "FGJ0101.frx":44DB
      End
      Begin Threed.SSCommand cmdRspdcls 
         Height          =   315
         Left            =   2370
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   225
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "CLEAR"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Picture         =   "FGJ0101.frx":45FD
      End
      Begin Threed.SSCommand cmdospdcls 
         Height          =   315
         Left            =   3150
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   2370
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "CLEAR"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Picture         =   "FGJ0101.frx":4619
      End
      Begin Threed.SSPanel pnllastlabno 
         Height          =   375
         Left            =   120
         TabIndex        =   88
         Top             =   7350
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "최종작업번호"
         ForeColor       =   8454143
         BackColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BorderWidth     =   1
      End
      Begin VB.Label lblLastdate 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         TabIndex        =   91
         Top             =   7380
         Width           =   1095
      End
      Begin VB.Label lblLastSlipCd 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2670
         TabIndex        =   90
         Top             =   7380
         Width           =   525
      End
      Begin VB.Label lblLastSeq 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3210
         TabIndex        =   89
         Top             =   7380
         Width           =   705
      End
   End
End
Attribute VB_Name = "FGJ0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DCJ0101         As DCJ0101  ' DB Class
Dim Sys_Date        As String
Dim UserCd          As String   ' 사용자 ID
Dim Update_F        As Integer  ' 데이타가 수정인지 신규인지 체크
Dim Person_F        As Integer  ' 환자의 기존 존재 여부 체크
Dim CodeHelp_F      As Integer  ' 코드헬프경우 체크
Dim Rnow_row        As Integer  ' Routine Spread Current Row
Dim Onow_row        As Integer  ' TestItem Spread Current Row
Dim iBtn_Ok         As Integer

Private Sub search_clear()
' 조회자료 Clear
    spdRoutine.MaxRows = 0
    spdorder.MaxRows = 0
    spdsLabNo.MaxRows = 0
    
    Rnow_row = 1
    Onow_row = 1
    
    lblLLabdate.Caption = ""
    lblLSlipCd.Caption = ""
    lblLLabseq.Caption = ""

' 신상자료 Clear
    txtdeptcd.Text = ""
    lbldeptnm.Caption = ""
    txtward.Text = ""
    optgubun(0).Value = False
    optgubun(1).Value = False
    optgubun(2).Value = False
    txtem.Text = "N"
    txtDr.Text = ""
    txtspecial.Text = "N"
    txtcmt.Text = ""
    
End Sub
Private Sub chkem_Click()

    If chkem.Value = 1 Then
        txtem.Tag = "1"
        txtem.Text = "Y"
    Else
        txtem.Tag = "0"
        txtem.Text = "N"
    End If

End Sub

Private Sub chkspecial_Click()

    If chkspecial.Value = 1 Then
        txtspecial.Tag = "1"
        txtspecial.Text = "Y"
    Else
        txtspecial.Tag = "0"
        txtspecial.Text = "N"
    End If
    
End Sub

Public Function fGetCurDSN() As String
    Dim sBuf$
    Dim bRetVal As Boolean
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\DSN", "")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\DSN", "", "SemiLIS")
        
        If bRetVal = True Then
            fGetCurDSN = "SemiLIS"
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
            fGetCurDSN = "SemiLIS"
        End If
    Else
        fGetCurDSN = sBuf
    End If
End Function



Private Function SameCheck(tmpLabDate As String, tmpslipcd As String, tmpSpcCd As String) As Boolean
    
    'Dim AdoRecord   As New ADODB.Connection
    Dim AdoRecord   As New ADODB.Recordset
    
    Dim sSql As String
        
    SameCheck = False
    
    sSql = " Select * From H_JUBSU " & _
       "  WHERE LABDATE     = '" & tmpLabDate & "'" & _
       "    AND PARTGBN     = '" & Mid(tmpslipcd, 2, 2) & "'" & _
       "    AND REGNO       = '" & tmpSpcCd & "'"
       
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
        SameCheck = True
    End If

    AdoRecord.Close
    
End Function

Private Sub cmdAppend_Click()
    Dim iCnt        As Integer
    Dim spd_F       As Integer
    Dim RegNo_F     As Integer
    Dim OCnt        As Integer
    
    Dim RtnCd       As String
    Dim Rcd
    Dim Ordcd
    Dim Chk_val
    
    Dim sSex        As String
    Dim sBornYear   As String
    Dim sLabDate    As String
    Dim sIdLeft     As String
    Dim sIdRight    As String
    Dim JubSuGbn    As String
    Dim ErGbn       As String
    Dim SpecialGbn  As String
    Dim sPerson     As String
    Dim sLabSeq     As String       ' 수정시 수정하는 Lab No.
    Dim sJubSu_S    As String
    Dim sJubSu_M    As String
    Dim sJubSu_D    As String
    Dim sTran_JubSu As String       ' 결과등록 성공/실패 결과값
    Dim sDelta      As String
    
    iBtn_Ok = True
    
    If Trim(lblSlipNm.Caption) = "" Then
        ViewMsg "Slip코드를 입력하여 주십시요"
        txtslipcd.SetFocus
        Exit Sub
    End If
    
    If Trim(lblSpcNm.Caption) = "" Then
        ViewMsg "검체코드를 입력하여 주십시요"
        txtSpcCd.SetFocus
        Exit Sub
    End If
    
    If Trim(txtRegNo.Text) = "" Then
        RegNo_F = False
    Else
        RegNo_F = True
    End If
    
    If spdorder.MaxRows = 0 Then
        ViewMsg "입력된 검사항목이 하나도 없습니다."
        txtoordcd.SetFocus
        Exit Sub
    End If

    spd_F = False
    For iCnt = 1 To spdorder.MaxRows
        RtnCd = spdorder.GetText(1, iCnt, Ordcd)
        If Trim(Ordcd) = "1" Then
            spd_F = True
            Exit For
        End If
    Next iCnt
    If spd_F = False Then
        ViewMsg "선택된 검사항목이 하나도 없습니다."
        txtoordcd.SetFocus
        Exit Sub
    End If
        
    sLabDate = Format(dtpLabDate.Value, "YYYYMMDD")
    JubSuGbn = optgubun(0).Tag
    ErGbn = txtem.Tag
    SpecialGbn = txtspecial.Tag
    
    If txtsex.Text = "남" Then
        sSex = "1"
    ElseIf txtsex.Text = "여" Then
        sSex = "2"
    Else
        sSex = "0"
    End If

    If Len(Trim(txtidleft.Text)) < 6 Then
        sIdLeft = ""
    Else
        sIdLeft = Trim(txtidleft.Text)
    End If

    If Len(Trim(txtidright.Text)) < 7 Then
        sIdRight = ""
    Else
        sIdRight = Trim(txtidright.Text)
    End If
    
    If SameCheck(sLabDate, txtslipcd, txtRegNo) Then
        If MsgBox("이미 등록된 자료입니다. 신규 데이터로 등록 하시겠습니까?", vbOKCancel + vbInformation, Me.Caption) <> vbOK Then
            Exit Sub
        End If
    End If
    
    If (Update_F = False Or spdsLabNo.MaxRows = 0) And lblLLabdate = "" Then
        sJubSu_S = ""
        Update_F = False
        sDelta = sLabDate & "|" & txtslipcd.Text & "|"
    Else
        sLabSeq = lblLLabseq.Caption & "|"
        sJubSu_S = sLabSeq
        Update_F = True
    End If
    
    If txtage.Text <> "" Then
        Set DCJ0101 = New DCJ0101
        sBornYear = Str(Val(Left(DCJ0101.Get_Date("DS"), 4)) - Val(txtage.Text))
        Set DCJ0101 = Nothing
    Else
        sBornYear = ""
    End If
    
    sPerson = txtRegNo.Text & "|" & txtname.Text & "|" & sBornYear & "|" & _
              txtidleft.Text & "|" & sSex & "|" & txtidright.Text & "|"
              
    If RegNo_F = False Then
        sPerson = ""
    End If
    
    sJubSu_S = sJubSu_S & sLabDate & "|" & Mid(txtslipcd.Text, 2) & "|" & txtSpcCd.Text & "|" & _
               txtRegNo.Text & "|" & txtname.Text & "|" & _
               txtage.Text & "|" & txtidleft.Text & "|" & _
               sSex & "|" & txtidright.Text & "|"
          
    sJubSu_M = txtrpartcd.Text & "|" & sJubSu_S & txtdeptcd.Text & "|" & txtward.Text & "|" & _
               JubSuGbn & "|" & ErGbn & "|" & txtDr.Text & "|" & _
               txtcmt.Text & "|" & SpecialGbn & "|"
               
    sJubSu_D = txtrpartcd.Text & "|" & sJubSu_S
    OCnt = 0
    For iCnt = 1 To spdorder.MaxRows
        RtnCd = spdorder.GetText(1, iCnt, Chk_val)
        If Chk_val = "1" Then
            OCnt = OCnt + 1
        End If
    Next iCnt
    
    sJubSu_D = sJubSu_D & Trim(Str(OCnt)) & "|"
    For iCnt = 1 To spdorder.MaxRows
        RtnCd = spdorder.GetText(1, iCnt, Chk_val)
        If Chk_val = "1" Then
            RtnCd = spdorder.GetText(4, iCnt, Ordcd)
            RtnCd = spdorder.GetText(5, iCnt, Rcd)
            sJubSu_D = sJubSu_D & Trim(Mid(Ordcd, 7, 3)) & "|" & Trim(Mid(Ordcd, 10, 4)) & "|" & Trim(Rcd) & "|"
        End If
    Next iCnt
    
    Set DCJ0101 = New DCJ0101
    
        sTran_JubSu = DCJ0101.Tran_JubSu(sPerson, Person_F, sJubSu_M, Update_F, sJubSu_D, Update_F, sDelta)
    
    Set DCJ0101 = Nothing
    
    If IsNumeric(sTran_JubSu) Then
        ViewMsg "입력하신 검체의 접수가 등록 되었습니다."
        lblLLabdate.Caption = Format(dtpLabDate.Value, "YYYYMMDD")
        lblLSlipCd.Caption = txtslipcd.Text
        lblLLabseq.Caption = sTran_JubSu
        Person_F = True
    End If
    
    Set DCJ0101 = New DCJ0101
    lblLastSeq = DCJ0101.Get_LastLabNo(Format(dtpLabDate.Value, "YYYYMMDD"), Left(fCurUserSlipCd, 1), Mid(fCurUserSpcCd, 2, 2))
    Set DCJ0101 = Nothing
    iBtn_Ok = False
    txtRegNo.SetFocus
       
    
End Sub

Private Sub cmdAppend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    iBtn_Ok = True
    
End Sub

Private Sub cmdDelete_Click()
    
    Dim RtnCd       As Integer
    Dim iCnt        As Integer
    Dim MsgStr      As String
    Dim spdLabDate
    Dim spdSlipCd
    Dim spdLabSeq
    
' 필수코드 체크
    If lblLLabdate.Caption = "" Or lblLSlipCd.Caption = "" Or lblLLabseq.Caption = "" Then
        ViewMsg "삭제할 작업번호를 먼저 선택하십시요."
        Exit Sub
    End If

' 결과존재여부 체크
    Set DCJ0101 = New DCJ0101
    If DCJ0101.Get_ResultCnt(lblLLabdate.Caption, lblLSlipCd.Caption, lblLLabseq.Caption) > 0 Then
        MsgStr = "정말로 해당 작업번호의 모든 접수내용을 삭제하시겠습니까?"
    Else
        MsgStr = "해당 작업번호의 결과도 함께 삭제되어집니다. 그래도 삭제하시겠습니까?"
    End If
    Set DCJ0101 = Nothing
    
' 삭제여부 확인
    RtnCd = MsgBox(MsgStr, vbYesNo, "Semi-LIS")
    
    If RtnCd = 6 Then ' TEST
        Set DCJ0101 = New DCJ0101
        
        If "SUCCESS" <> DCJ0101.Tran_JubSu_Del(lblLLabdate.Caption & "|" & lblLSlipCd.Caption & "|" & lblLLabseq.Caption & "|") Then
' 결과 삭제/ 접수 삭제
            ViewMsg "삭제가 실패되었습니다."
        Else
        ' 스프레드에서 LabNo지우기
            For iCnt = 1 To spdsLabNo.MaxRows
                RtnCd = spdsLabNo.GetText(1, iCnt, spdLabDate)
                RtnCd = spdsLabNo.GetText(2, iCnt, spdSlipCd)
                RtnCd = spdsLabNo.GetText(3, iCnt, spdLabSeq)
                If Trim(spdLabDate & spdSlipCd & spdLabSeq) = Trim(lblLLabdate.Caption & lblLSlipCd.Caption & lblLLabseq.Caption) Then
                    spdsLabNo.Row = iCnt
                    spdsLabNo.Action = 5 ' SS_DELTE_ROW
                    spdsLabNo.MaxRows = spdsLabNo.MaxRows - 1
                    spdsLabNo.Row = iCnt - 1
                    spdsLabNo.Action = 1 ' SS_ACTIVE_ROW
                End If
            Next iCnt
            
            ViewMsg "삭제가 성공되었습니다."
            Update_F = False
            spdRoutine.MaxRows = 0
            spdorder.MaxRows = 0
            Rnow_row = 1
            Onow_row = 1
            lblLLabdate.Caption = ""
            lblLSlipCd.Caption = ""
            lblLLabseq.Caption = ""
              '마지막 환자 정보 표시
            lblLastSeq = DCJ0101.Get_LastLabNo(Format(dtpLabDate.Value, "YYYYMMDD"), Left(fCurUserSlipCd, 1), Mid(fCurUserSpcCd, 2, 2))
        
        End If
                  
        Set DCJ0101 = Nothing
    End If
    
End Sub


Private Sub cmddepth_Click()

    Dim i%
    Dim j%
    Dim CDept As DCB0101
    Dim sTot01$
    Dim sTot02$
    
    txtdeptcd.SetFocus
    
    Set CDept = New DCB0101
    
    CDept.Get_DEPT
    j = CDept.CurItemCnt
    
    Erase gCodeHlpTable '배열 초기화
    
    ReDim gCodeHlpTable(j) As CodeTBL
    
    With CDept
        sTot01 = .TotField01
        sTot02 = .TotField02
    End With
    
    Set CDept = Nothing
    
    For i = 1 To j
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = GetByOne(sTot01, sTot01)
        gCodeHlpTable(i).sCodeNm = GetByOne(sTot02, sTot02)
    Next
    
    giCodeHlpCnt = j
    
    hWndCd = txtdeptcd.hwnd
    
    FSJ0101.Left = 1600
    FSJ0101.Top = 3930
    
    CodeHelp_F = True
    
    Load FSJ0101
    FSJ0101.Show vbModal

End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    
End Sub

Private Sub cmdlabnoh_Click()

    Dim i%
    Dim j%
    Dim CPart As DCB0101
    Dim sTot01$
    Dim sTot02$
    Dim sTot03$
    
    txtsslipcd.SetFocus
    
    Set CPart = New DCB0101
    
    CPart.Get_PART
    
    j = CPart.CurItemCnt
    
    Erase gCodeHlpTable '배열 초기화
    
    ReDim gCodeHlpTable(j) As CodeTBL
    
    With CPart
        sTot01 = .TotField01
        sTot02 = .TotField02
        sTot03 = .TotField03
    End With
    
    Set CPart = Nothing
    
    For i = 1 To j
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = GetByOne(sTot01, sTot01) & GetByOne(sTot02, sTot02)
        gCodeHlpTable(i).sCodeNm = GetByOne(sTot03, sTot03)
    Next
    
    giCodeHlpCnt = j
    
    hWndCd = txtsslipcd.hwnd
    
    FSJ0101.Left = 6370
    FSJ0101.Top = 3950
        
    CodeHelp_F = True
    
    Load FSJ0101
    FSJ0101.Show vbModal

End Sub

Private Sub cmdordh_Click()
    
    Dim CTestItem As DCB0101
    Dim j%
    Dim i%
    Dim sTot01$, sTot02$, sTot03$, sTot04$, sTot05$
    Dim sTmp1$
    Dim sTmp2$
    Dim vTmp
    
    If txtoslipcd = "" Or txtospccd = "" Then
        MsgBox "먼저 Routine 코드와 Routine 명을 입력한 후, 검사항목을 선택하여 주십시요!!"
        Exit Sub
    End If
    
    txtoordcd.SetFocus
    txtoordcd.Text = ""
    
    Set CTestItem = New DCB0101
    
    vTmp = txtslipcd
    CTestItem.Get_TESTITEM 13, Left$(CStr(vTmp), 1), Mid$(CStr(vTmp), 2, 2), txtSpcCd
        
    j = CTestItem.CurItemCnt
    
    Erase gCodeHlpTable '배열 초기화
    
    With CTestItem
        sTot01 = .TotField01    'PartGbn
        sTot02 = .TotField02    'SpecimenCd
        sTot03 = .TotField03    'TestItemSeq
        sTot04 = .TotField04    'SUBMCD
        sTot05 = .TotField05    'TESTITEMNM
    End With
    
    Set CTestItem = Nothing

    ReDim gCodeHlpTable(j) As CodeTBL
    
    giCodeHlpCnt = 0
    
    For i = 1 To j
        sTmp1 = GetByOne(sTot04, sTot04)
        
        If sTmp1 = "NNNN" Then
            giCodeHlpCnt = giCodeHlpCnt + 1
            gCodeHlpTable(giCodeHlpCnt).sGbn = "N"
            gCodeHlpTable(giCodeHlpCnt).sSeq = Format$(giCodeHlpCnt, "00000")
            gCodeHlpTable(giCodeHlpCnt).sCode = Left(txtoslipcd, 1) & GetByOne(sTot01, sTot01) & _
                                GetByOne(sTot02, sTot02) & GetByOne(sTot03, sTot03)
        
            gCodeHlpTable(giCodeHlpCnt).sCodeNm = GetByOne(sTot05, sTot05)
            
        ElseIf IsNumeric(Left$(sTmp1, 2)) = True And Left$(sTmp1, 2) = "00" Then
            'SUB 원검사만 추가
            giCodeHlpCnt = giCodeHlpCnt + 1
            gCodeHlpTable(giCodeHlpCnt).sGbn = "S" & Left$(sTmp1, 2)
            gCodeHlpTable(giCodeHlpCnt).sSeq = Format$(giCodeHlpCnt, "00000")
            gCodeHlpTable(giCodeHlpCnt).sCode = Left(txtoslipcd, 1) & GetByOne(sTot01, sTot01) & _
                                GetByOne(sTot02, sTot02) & GetByOne(sTot03, sTot03)
        
            gCodeHlpTable(giCodeHlpCnt).sCodeNm = GetByOne(sTot05, sTot05)
            
        Else
            'SUB 원검사 이외의 제외
            Call GetByOne(sTot01, sTot01)
            Call GetByOne(sTot02, sTot02)
            Call GetByOne(sTot03, sTot03)
            Call GetByOne(sTot05, sTot05)
            
        End If
    Next
    
    giCodeHlpMode = 1
    CodeHelp_F = True
    
    Set gCallObject = FGJ0101.txtoordcd
    
    FSJ0201.Left = 6000
    FSJ0201.Top = 3770
    
    Load FSJ0201
    FSJ0201.Show vbModal
    
End Sub

Private Sub cmdospdcls_Click()

    spdRoutine.MaxRows = 0
    spdorder.MaxRows = 0
    
    Rnow_row = 1
    Onow_row = 1
    
    txtoordcd.SetFocus
    
End Sub

Private Sub cmdroutineh_Click()

    Dim i%, j%
    Dim CRoutine As DCB0101
    Dim sField01$, sField02$
    
    txtroutinecd.SetFocus
    txtroutinecd.Text = ""
        
    Set CRoutine = New DCB0101
    
    If txtrpartcd = "" Then
        MsgBox "PART 코드를 입력하여야 그 PART 아래의 ROUTINE 코드를 볼 수 있습니다!!"
        Exit Sub
    Else
        CRoutine.Get_RTN 4, txtrpartcd, txtospccd     'SELECT WITH PARTCD AND SPECIMENCD
    End If
    
    i = CRoutine.CurItemCnt
    
    If i = 0 Then
        MsgBox "PART 코드 - " & txtrpartcd & " 에 설정된 ROUTINE 코드가 없습니다!!"
        Set CRoutine = Nothing
        Exit Sub
    End If
    
    Erase gCodeHlpTable '배열 초기화
    
    ReDim gCodeHlpTable(i) As CodeTBL
    
    sField01 = CRoutine.TotField01
    sField02 = CRoutine.TotField02
    
    For j = 1 To i
        gCodeHlpTable(j).sSeq = Format$(j, "00000")
        gCodeHlpTable(j).sCode = GetByOne(sField01, sField01)
        gCodeHlpTable(j).sCodeNm = GetByOne(sField02, sField02)
    Next

    giCodeHlpCnt = i
    
    giCodeHlpMode = 1
    
    Set gCallObject = FGJ0101.txtroutinecd
    
    FSJ0201.Left = 6200
    FSJ0201.Top = 1200
    
    CodeHelp_F = True
    
    Load FSJ0201
    FSJ0201.Show vbModal

End Sub

Private Sub cmdRspdcls_Click()

    Dim Rcd     As Variant
    Dim ORcd    As Variant
    Dim RtnCd   As Integer
    Dim iCnt    As Integer  ' Routine Spread건수
    Dim sCnt    As Integer  ' 검사항목 Spread건수

' 풀려진 검사항목을 조사, 제거한다.
    For iCnt = 1 To spdRoutine.MaxRows
        RtnCd = spdRoutine.GetText(3, iCnt, Rcd)
        For sCnt = 1 To spdorder.MaxRows
            RtnCd = spdorder.GetText(5, sCnt, ORcd)
            If Rcd = ORcd Then
                spdorder.Row = sCnt
                spdorder.Action = 5 'SS_DELETE_ROW
                spdorder.MaxRows = spdorder.MaxRows - 1
                Onow_row = Onow_row - 1
                sCnt = sCnt - 1
            End If
        Next sCnt
    Next iCnt
    
    spdRoutine.MaxRows = 0
    Rnow_row = 1
    
    txtroutinecd.SetFocus
    
End Sub

Private Sub cmdsliph_Click()

    Dim i%
    Dim j%
    Dim CPart As DCB0101
    Dim sTot01$
    Dim sTot02$
    Dim sTot03$
    Dim tmpSlip As String
    
    txtslipcd.SetFocus
    
    Set CPart = New DCB0101
    
    CPart.Get_PART
    
    tmpSlip = txtslipcd
    
    j = CPart.CurItemCnt
    
    Erase gCodeHlpTable '배열 초기화
    
    ReDim gCodeHlpTable(j) As CodeTBL
    
    With CPart
        sTot01 = .TotField01
        sTot02 = .TotField02
        sTot03 = .TotField03
    End With
    
    Set CPart = Nothing
    
    For i = 1 To j
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = GetByOne(sTot01, sTot01) & GetByOne(sTot02, sTot02)
        gCodeHlpTable(i).sCodeNm = GetByOne(sTot03, sTot03)
    Next
    
    giCodeHlpCnt = j
    
    hWndCd = txtslipcd.hwnd
    
    FSJ0101.Left = 1700
    FSJ0101.Top = 1000
    
' Code Help Flag
    CodeHelp_F = True
    
    Load FSJ0101
    FSJ0101.Show vbModal
    
    If tmpSlip <> txtslipcd Then
        Call search_clear
    End If
    
End Sub

Private Sub cmdspch_Click()

    Dim i%
    Dim j%
    Dim CSpecimen As DCB0101
    Dim sTot01$
    Dim sTot02$
    
    txtSpcCd.SetFocus
        
    Set CSpecimen = New DCB0101
    
    CSpecimen.Get_SPC
    
    j = CSpecimen.CurItemCnt
    
    Erase gCodeHlpTable '배열 초기화
    
    ReDim gCodeHlpTable(j) As CodeTBL
    
    With CSpecimen
        sTot01 = .TotField01
        sTot02 = .TotField02
    End With
    
    Set CSpecimen = Nothing
    
    For i = 1 To j
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = GetByOne(sTot01, sTot01)
        gCodeHlpTable(i).sCodeNm = GetByOne(sTot02, sTot02)
    Next
    
    giCodeHlpCnt = j
    
    hWndCd = txtSpcCd.hwnd
    
    FSJ0101.Left = 1700
    FSJ0101.Top = 1400
    
    CodeHelp_F = True
    
    Load FSJ0101
    FSJ0101.Show vbModal

End Sub

Private Sub dtpLabDate_Change()

    txtslabdate = Format(dtpLabDate.Value, "YYYYMMDD")
    lblLastdate = Format(dtpLabDate.Value, "YYYYMMDD")
    
    Set DCJ0101 = New DCJ0101
    lblLastSeq = DCJ0101.Get_LastLabNo(Format(dtpLabDate.Value, "YYYYMMDD"), Left(fCurUserSlipCd, 1), Mid(fCurUserSpcCd, 2, 2))
    Set DCJ0101 = Nothing
    txtsslipcd.Text = txtslipcd.Text

End Sub

Private Sub dtpLabDate_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        txtslabdate = Format(dtpLabDate.Value, "YYYYMMDD")
        lblLastdate = Format(dtpLabDate.Value, "YYYYMMDD")
        txtslipcd.SetFocus
        KeyCode = 0
    End If

End Sub

Private Sub Form_Activate()

    Dim DefaultCd   As String  ' SlipCd와 LabSeq
    Dim sLastSeq    As String  ' 마지막 번호 취득 키
        
' 기초자료 초기화 작업
    
    If CodeHelp_F = False Then
        
        Person_F = False
        
        spdRoutine.Row = -1
        spdRoutine.Col = 2
        spdRoutine.Col2 = 3
        spdRoutine.BlockMode = True
        spdRoutine.Lock = True
        spdRoutine.BlockMode = False
        
        spdorder.Row = -1
        spdorder.Col = 2
        spdorder.Col2 = 6
        spdorder.BlockMode = True
        spdorder.Lock = True
        spdorder.BlockMode = False
        
        txtslipcd.Text = fCurUserSlipCd
        lblSlipNm.Caption = fCurUserSlipNm
        txtSpcCd.Text = fCurUserSpcCd
        lblSpcNm.Caption = fCurUserSpcNm
    
        Set DCJ0101 = New DCJ0101
        txtslabdate.Text = DCJ0101.Get_Date("DS")
        lblLastSeq = DCJ0101.Get_LastLabNo(Format(dtpLabDate.Value, "YYYYMMDD"), Left(fCurUserSlipCd, 1), Mid(fCurUserSpcCd, 2, 2))
        Set DCJ0101 = Nothing
        lblLastdate = Format(dtpLabDate.Value, "YYYYMMDD")
        lblLastSlipCd = txtslipcd
        txtsslipcd.Text = txtslipcd.Text
        txtRegNo.SetFocus
        
    End If
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyF2
        Call cmdAppend_Click
        KeyCode = 0
    Case vbKeyF4
        Call cmdDelete_Click
        KeyCode = 0
    Case vbKeyEscape
        Call cmdExit_Click
        KeyCode = 0
    End Select
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Dim iCnt    As Integer
    
'If KeyAscii = KEY_RETURN Then
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
        ViewMsg ""
    End If
    
End Sub

Private Sub Form_Load()
    
    Me.Left = 0
    Me.Top = 0
    Me.Width = 11920
    Me.Height = 7950

' 변수 초기화
    Rnow_row = 1
    Onow_row = 1
    CodeHelp_F = False
    iBtn_Ok = False
' 날짜 초기화
    Set DCJ0101 = New DCJ0101
    Sys_Date = DCJ0101.Get_Date("D")
    dtpLabDate.Value = Sys_Date
    Set DCJ0101 = Nothing
          
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call InitRegCurFrmTitle
    
End Sub

Private Sub optgubun_Click(Index As Integer)

    optgubun(0).Tag = Trim(Str(Index))
    
End Sub


Private Sub optResOK_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim RtnCd       As Integer
    Dim RCnt        As Integer
    Dim iCnt        As Integer
    Dim sLabDate    As String
    Dim sJubSu      As String
    Dim sPartCd     As String
    
' 해당일자 취득
    sLabDate = Format(dtpLabDate.Value, "YYYYMMDD")
    sPartCd = txtrpartcd.Text
    
    Set DCJ0101 = New DCJ0101
    
    If Index = 0 Then
' 해당일의 검사미완료 검체번호 조회
        sJubSu = DCJ0101.Get_JubSuList(sLabDate, sPartCd, "YES")
    ElseIf Index = 1 Then
' 해당일의 검사완료 검체번호 조회
        sJubSu = DCJ0101.Get_JubSuList(sLabDate, sPartCd, "NO")
    End If
    
    Set DCJ0101 = Nothing
    
' Spread에 데이타 표시
    If sJubSu <> "" Then
        RCnt = Val(GetByOne(sJubSu, sJubSu))
        For iCnt = 1 To RCnt
            spdsLabNo.MaxRows = iCnt
            Call spdsLabNo.SetText(1, iCnt, sLabDate)
            Call spdsLabNo.SetText(2, iCnt, Trim(GetByOne(sJubSu, sJubSu)))
            Call spdsLabNo.SetText(3, iCnt, Trim(GetByOne(sJubSu, sJubSu)))
        Next iCnt
    Else
        spdsLabNo.MaxRows = 0
        ViewMsg "해당 데이타가 없습니다."
    End If

End Sub

Private Sub optsearch_Click(Index As Integer)

    pnlslabno.Visible = False
    pnlsRegno.Visible = False
    pnlsokres.Visible = False
    
    Select Case Index
        Case 0
            pnlslabno.Visible = True
            txtslabdate.Text = Format(dtpLabDate.Value, "YYYYMMDD")
            txtsslipcd.Text = txtslipcd.Text
            txtslabseq.SetFocus
        Case 1
            pnlsRegno.Visible = True
            txtsRegno.SetFocus
        Case 2
            pnlsokres.Visible = True
            optResOK(0).SetFocus
    End Select
    
End Sub

Private Sub spdRoutine_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    Dim RtnCd   As Integer
    Dim spdCnt  As Integer
    Dim Chk_F
    Dim RRcd
    Dim ORcd
       
    If Row > 0 Then
        RtnCd = spdRoutine.GetText(1, Row, Chk_F)
        RtnCd = spdRoutine.GetText(3, Row, RRcd)
    
        If Chk_F = "0" Then
' 검사항목 제거하기
            For spdCnt = 1 To spdorder.MaxRows
                RtnCd = spdorder.GetText(5, spdCnt, ORcd)
                If RRcd = ORcd Then
                    Call spdorder.SetText(1, spdCnt, "0")
                End If
            Next spdCnt
        ElseIf Chk_F = "1" Then
' 검사항목 재추가하기
            For spdCnt = 1 To spdorder.MaxRows
                RtnCd = spdorder.GetText(5, spdCnt, ORcd)
                If RRcd = ORcd Then
                    Call spdorder.SetText(1, spdCnt, "1")
                End If
            Next spdCnt
        End If
        spdRoutine.Col = 0
        spdRoutine.Row = 0
        spdRoutine.Action = 1
    End If

End Sub

Private Sub spdsLabNo_Click(ByVal Col As Long, ByVal Row As Long)

    Dim RtnCd       As Integer
    Dim RCnt        As Integer
    Dim iCnt        As Integer
    Dim LabDate     As String
    Dim SlipCd      As String
    Dim LabSeq      As String
    Dim sJubSu      As String
    Dim sPrintNm    As String
    Dim sOrdCd      As String
    Dim vTemp       As Variant
    
    If Row < 1 Then Exit Sub
    
    RtnCd = spdsLabNo.GetText(1, Row, vTemp)
    LabDate = vTemp
    RtnCd = spdsLabNo.GetText(2, Row, vTemp)
    SlipCd = vTemp
    RtnCd = spdsLabNo.GetText(3, Row, vTemp)
    LabSeq = vTemp
    
    Set DCJ0101 = New DCJ0101
' 접수 정보 취득하기
        sJubSu = DCJ0101.Get_JubSu(LabDate, SlipCd, LabSeq)
        
        If Trim(sJubSu) = "" Then
            Update_F = False
        Else
' 접수 정보 화면 표시
    ' 기초신상 표시
            Rnow_row = 1
            Onow_row = 1
            dtpLabDate.Value = Mid(LabDate, 1, 4) & "-" & Mid(LabDate, 5, 2) & "-" & Mid(LabDate, 7, 2)
            lblLLabdate.Caption = LabDate
            lblLSlipCd.Caption = SlipCd
            lblLLabseq.Caption = LabSeq
            txtslipcd.Text = SlipCd
            txtSpcCd.Text = GetByOne(sJubSu, sJubSu)        '1
            txtRegNo.Text = GetByOne(sJubSu, sJubSu)        '2
            txtname.Text = GetByOne(sJubSu, sJubSu)         '3
            txtage.Text = GetByOne(sJubSu, sJubSu)        '4
            txtidleft.Text = GetByOne(sJubSu, sJubSu)       '5
            vTemp = GetByOne(sJubSu, sJubSu)                '6
            If vTemp = "1" Or vTemp = "3" Then
                txtsex.Text = "남"
            ElseIf vTemp = "2" Or vTemp = "4" Then
                txtsex.Text = "여"
            Else
                txtsex.Text = ""
            End If
            txtidright.Text = GetByOne(sJubSu, sJubSu)      '7
            txtdeptcd.Text = GetByOne(sJubSu, sJubSu)       '8
            lbldeptnm.Caption = GetByOne(sJubSu, sJubSu)    '9
            txtward.Text = GetByOne(sJubSu, sJubSu)         '10
            optgubun(Val(GetByOne(sJubSu, sJubSu))).Value = True '11
            vTemp = GetByOne(sJubSu, sJubSu)                '12
            If vTemp = "0" Then
                txtem.Text = "N"
                chkem.Value = 0
            ElseIf vTemp = "1" Then
                txtem.Text = "Y"
                chkem.Value = 1
            End If
            txtDr.Text = GetByOne(sJubSu, sJubSu)           '13
            txtcmt.Text = GetByOne(sJubSu, sJubSu)          '14
            vTemp = GetByOne(sJubSu, sJubSu)                '15
            If vTemp = "0" Then
                txtspecial.Text = "N"
                chkspecial.Value = 0
            ElseIf vTemp = "1" Then
                txtspecial.Text = "Y"
                chkspecial.Value = 1
            End If
    ' 접수내용 표시
            RCnt = Val(GetByOne(sJubSu, sJubSu))
            spdorder.MaxRows = RCnt
            For iCnt = 1 To RCnt
                sPrintNm = GetByOne(sJubSu, sJubSu)
                sOrdCd = GetByOne(sJubSu, sJubSu)
                Call spdorder.SetText(1, iCnt, "1")
                If Mid(sOrdCd, 10, 2) <> "NN" Then
                    Call spdorder.SetText(2, iCnt, "S")
                Else
                    Call spdorder.SetText(2, iCnt, "")
                End If
                Call spdorder.SetText(3, iCnt, sPrintNm)  '검사코드
                Call spdorder.SetText(4, iCnt, sOrdCd)    '검사약명
                Call spdorder.SetText(5, iCnt, GetByOne(sJubSu, sJubSu)) '묶음검사코드
                Call spdorder.SetText(6, iCnt, "S")
            Next iCnt
            Onow_row = iCnt
            RCnt = Val(GetByOne(sJubSu, sJubSu))
            spdRoutine.MaxRows = RCnt
            For iCnt = 1 To RCnt
                Call spdRoutine.SetText(1, iCnt, "1")
                Call spdRoutine.SetText(2, iCnt, GetByOne(sJubSu, sJubSu)) '묶음검사코드
                Call spdRoutine.SetText(3, iCnt, GetByOne(sJubSu, sJubSu)) '묶음검사명
                Call spdRoutine.SetText(4, iCnt, "S")
            Next iCnt
            Rnow_row = iCnt
            Update_F = True
            Person_F = True
        End If
    Set DCJ0101 = Nothing
    
    cmdAppend.SetFocus
    
End Sub

Private Sub SSCommand8_Click()

End Sub

Private Sub SSCommand7_Click()

End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub txtage_LostFocus()
    
    txtsex.SetFocus

End Sub


Private Sub txtcmt_GotFocus()

    Call Txt_Highlight(txtcmt)
    FGJ0101.KeyPreview = False
    
End Sub

Private Sub txtcmt_LostFocus()

    FGJ0101.KeyPreview = True
    
End Sub

Private Sub txtdeptcd_Change()

    If Len(txtdeptcd.Text) = txtdeptcd.MaxLength Then
        Set DCJ0101 = New DCJ0101
            lbldeptnm.Caption = DCJ0101.Get_DeptNm(txtdeptcd.Text)
        Set DCJ0101 = Nothing
        If lbldeptnm.Caption = "" Then
            ViewMsg "존재하지 않는 진료과 CODE입니다."
            txtdeptcd.SetFocus
        End If
        If CodeHelp_F = False Then
            txtward.SetFocus
        Else
            SendKeys "{ENTER}"
        End If
    Else
        lbldeptnm.Caption = ""
    End If
    
End Sub

Private Sub txtdeptcd_GotFocus()

    Call Txt_Highlight(txtdeptcd)
    
End Sub

Private Sub txtdeptcd_KeyPress(KeyAscii As Integer)

    CodeHelp_F = False
    
End Sub

Private Sub txtdeptcd_LostFocus()

    If Len(txtdeptcd.Text) < txtdeptcd.MaxLength Then
        txtdeptcd.Text = Format(txtdeptcd.Text, "00")
    End If

End Sub

Private Sub txtDr_GotFocus()

    Call Txt_Highlight(txtDr)
    
End Sub

Private Sub txtem_Change()

    If Trim(txtem.Text) = "Y" Then
        txtem.Tag = "1"
        chkem.Value = 1
    ElseIf Trim(txtem.Text) = "N" Then
        txtem.Tag = "0"
        chkem.Value = 0
    End If
    
End Sub

Private Sub txtem_GotFocus()
    
    Call Txt_Highlight(txtem)
    
End Sub

Private Sub txtem_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtidleft_Change()

    If Len(txtidleft.Text) = txtidleft.MaxLength Then
        txtidright.SetFocus
    End If
    
End Sub

Private Sub txtidleft_GotFocus()
    
    Call Txt_Highlight(txtidleft)
    
End Sub

Private Sub txtidleft_LostFocus()

    If Len(txtidleft.Text) < 4 Then
        txtage.Text = txtidleft.Text
    Else
        txtage.Text = ""
    End If
    
End Sub

Private Sub txtidright_Change()

    If Len(txtidright.Text) = txtidright.MaxLength Then
        txtage.SetFocus
        'txtdeptcd.SetFocus
    End If
    
End Sub

Private Sub txtidright_GotFocus()

    Call Txt_Highlight(txtidright)
    
End Sub

Private Sub txtidright_LostFocus()
' 입력된 숫자가 주민번호가 아닐때 성별 판정
    If Left(Trim(txtidright.Text), 1) <> "" Then
        If Left(Trim(txtidright.Text), 1) = "1" Then
            txtsex.Text = "남"
        ElseIf Left(Trim(txtidright.Text), 1) = "2" Then
            txtsex.Text = "여"
        Else
            txtsex.Text = ""
        End If
    End If
    
    If Len(txtidleft.Text) = 6 And Len(txtidright.Text) = 7 Then
' 주민번호로 나이계산
        If Left(txtidright.Text, 1) = 1 Or Left(txtidright.Text, 1) = 2 Then
            txtage.Text = Val(Left(Sys_Date, 4)) - Val("19" & Left(txtidleft.Text, 2))
        ElseIf Left(txtidright.Text, 1) = 3 Or Left(txtidright.Text, 1) = 4 Then
            txtage.Text = Val(Left(Sys_Date, 4)) - Val("20" & Left(txtidleft.Text, 2))
        Else
            txtage.Text = ""
        End If
' 성별 판정
        If Left(txtidright.Text, 1) = "1" Or Left(txtidright.Text, 1) = "3" Then
            txtsex.Text = "남"
        ElseIf Left(txtidright.Text, 1) = "2" Or Left(txtidright.Text, 1) = "4" Then
            txtsex.Text = "여"
        Else
            txtsex.Text = ""
        End If
    End If
End Sub

Private Sub txtname_Change()
    
    Dim sRegInfo    As String
' 환자등록번호를 키로 하여 기존에 정보가 있는 환자는 신상을 표시
    If Len(txtname.Text) = txtname.MaxLength Then
        
        Set DCJ0101 = New DCJ0101
            sRegInfo = DCJ0101.Get_NameInfo(Trim(txtname.Text))
            If Trim(sRegInfo) <> "" Then
                txtidleft.Text = GetByOne(sRegInfo, sRegInfo)
                txtidright.Text = GetByOne(sRegInfo, sRegInfo)
                txtage.Text = Trim(GetByOne(sRegInfo, sRegInfo))
                txtsex.Text = GetByOne(sRegInfo, sRegInfo)
                txtRegNo.Text = GetByOne(sRegInfo, sRegInfo)
' 기존에 존재하는 환자
                Person_F = True
                If txtname.Tag <> txtname.Text Then
                    Update_F = False
                    Call search_clear
                    
                End If
            Else
' 존재하지 않는 환자
                Person_F = False
                Call search_clear
            End If
        Set DCJ0101 = Nothing
        txtidleft.SetFocus
    End If
End Sub

Private Sub txtname_GotFocus()
    
    Call Txt_Highlight(txtname)
    
End Sub

Private Sub txtname_LostFocus()

    Dim sRegInfo    As String
    
' 진찰권번호 수정인지 처리
   If iBtn_Ok = True Then
      iBtn_Ok = False
      Exit Sub
   End If
   
'   txtRegNo.Text = Format(txtRegNo.Text, "00000")

' 환자등록번호를 키로 하여 기존에 정보가 있는 환자는 신상을 표시
    If txtname.Tag <> txtname.Text Then
        If Len(txtname.Text) < txtname.MaxLength And Trim(txtname) <> "" Then
            Set DCJ0101 = New DCJ0101
                sRegInfo = DCJ0101.Get_NameInfo(Trim(txtname.Text))
                If Trim(sRegInfo) <> "" Then
                    txtidleft.Text = GetByOne(sRegInfo, sRegInfo)
                    txtidright.Text = GetByOne(sRegInfo, sRegInfo)
                    txtage.Text = Trim(GetByOne(sRegInfo, sRegInfo))
                    txtsex.Text = GetByOne(sRegInfo, sRegInfo)
                    txtRegNo.Text = GetByOne(sRegInfo, sRegInfo)
                    Person_F = True
                    If txtname.Tag <> txtname.Text Then
                        Update_F = False
                        Call search_clear
                    End If
                Else
                    Person_F = False
                    Call search_clear
                    txtidleft.Text = ""
                    txtidright.Text = ""
                    txtage.Text = ""
                    txtsex.Text = ""
'                    txtRegNo.Text = ""
                End If
            Set DCJ0101 = Nothing
'            txtidleft.SetFocus
        ElseIf Trim(txtname) = "" Then
            Person_F = False
            Call search_clear
        End If
    Else
        If lblLLabdate = "" Then
            Person_F = False
            Call search_clear
        End If
    End If
End Sub

Private Sub txtoordcd_Change()

    Dim iCnt        As Integer
    Dim Ordcd
    Dim sItemGbn    As String
    Dim sPrintNm    As String
    Dim sOrdNm      As String
    Dim sSubMcd     As String
        
    If Len(txtoordcd.Text) = txtoordcd.MaxLength Then
    
        For iCnt = 1 To spdorder.MaxRows
            Call spdorder.GetText(4, iCnt, Ordcd)
            If Left(Ordcd, 9) = txtoslipcd.Text & txtospccd.Text & txtoordcd.Text Then
                ViewMsg "이미 선택한 항목입니다."
                Call Txt_Highlight(txtoordcd)
                Exit Sub
            End If
        Next iCnt
        
        Set DCJ0101 = New DCJ0101
            sPrintNm = DCJ0101.Get_Order(txtoslipcd.Text, txtospccd.Text, txtoordcd.Text)
            If Trim(sPrintNm) = "" Then
                
            Else
' 검사명과 코드를 화면에 표시한다.
                sSubMcd = GetByOne(sPrintNm, sPrintNm)
                sOrdNm = GetByOne(sPrintNm, sPrintNm)
                
                spdorder.MaxRows = Onow_row
                Call spdorder.SetText(1, Onow_row, "1")
                
                If Left(sSubMcd, 2) <> "NN" Then
                    Call spdorder.SetText(2, Onow_row, "S")
                Else
                    Call spdorder.SetText(2, Onow_row, "")
                End If
                Call spdorder.SetText(3, Onow_row, sOrdNm)
                Call spdorder.SetText(4, Onow_row, txtoslipcd.Text & txtospccd.Text & txtoordcd.Text & sSubMcd)
                Onow_row = Onow_row + 1
            End If
        Set DCJ0101 = Nothing
        
        If CodeHelp_F = False Then
            Call Txt_Highlight(txtoordcd)
        End If
        
    End If
End Sub

Private Sub txtoordcd_GotFocus()
    
    Call Txt_Highlight(txtoordcd)
    
End Sub

Private Sub txtoordcd_KeyPress(KeyAscii As Integer)

    CodeHelp_F = False
    
End Sub

Private Sub txtRegNo_Change()
    
    Dim sRegInfo    As String
' 환자등록번호를 키로 하여 기존에 정보가 있는 환자는 신상을 표시
    If Len(txtRegNo.Text) = txtRegNo.MaxLength Then
        
        Set DCJ0101 = New DCJ0101
            sRegInfo = DCJ0101.Get_RegInfo(txtRegNo.Text)
            If Trim(sRegInfo) <> "" Then
                txtidleft.Text = GetByOne(sRegInfo, sRegInfo)
                txtidright.Text = GetByOne(sRegInfo, sRegInfo)
                txtage.Text = Trim(GetByOne(sRegInfo, sRegInfo))
                txtsex.Text = GetByOne(sRegInfo, sRegInfo)
                txtname.Text = GetByOne(sRegInfo, sRegInfo)
' 기존에 존재하는 환자
                Person_F = True
                If txtRegNo.Tag <> txtRegNo.Text Then
                    Update_F = False
                    Call search_clear
                    
                End If
            Else
' 존재하지 않는 환자
                Person_F = False
                Call search_clear
            End If
        Set DCJ0101 = Nothing
'        txtname.SetFocus
    End If
    
End Sub

Private Sub txtRegNo_GotFocus()

    Call Txt_Highlight(txtRegNo)
    txtRegNo.Tag = txtRegNo.Text
    
End Sub

Private Sub txtRegNo_LostFocus()

    Dim sRegInfo    As String
    
' 진찰권번호 수정인지 처리
   If iBtn_Ok = True Then
      iBtn_Ok = False
      Exit Sub
   End If
   
'   txtRegNo.Text = Format(txtRegNo.Text, "00000")

' 환자등록번호를 키로 하여 기존에 정보가 있는 환자는 신상을 표시
    If txtRegNo.Tag <> txtRegNo.Text Then
        If Len(txtRegNo.Text) < txtRegNo.MaxLength And Trim(txtRegNo) <> "" Then
            Set DCJ0101 = New DCJ0101
                sRegInfo = DCJ0101.Get_RegInfo(txtRegNo.Text)
                If Trim(sRegInfo) <> "" Then
                    txtidleft.Text = GetByOne(sRegInfo, sRegInfo)
                    txtidright.Text = GetByOne(sRegInfo, sRegInfo)
                    txtage.Text = Trim(GetByOne(sRegInfo, sRegInfo))
                    txtsex.Text = GetByOne(sRegInfo, sRegInfo)
                    txtname.Text = GetByOne(sRegInfo, sRegInfo)
                    Person_F = True
                    If txtRegNo.Tag <> txtRegNo.Text Then
                        Update_F = False
                        Call search_clear
                    End If
                Else
                    Person_F = False
                    Call search_clear
                    txtidleft.Text = ""
                    txtidright.Text = ""
                    txtage.Text = ""
                    txtsex.Text = ""
                    txtname.Text = ""
                End If
            Set DCJ0101 = Nothing
            txtname.SetFocus
        ElseIf Trim(txtRegNo) = "" Then
            Person_F = False
            Call search_clear
        End If
    Else
        If lblLLabdate = "" Then
            Person_F = False
            Call search_clear
        End If
    End If
    
End Sub

Private Sub txtroutinecd_Change()

    Dim RtnCd       As Integer
    Dim OrdCnt      As Integer
    Dim iCnt        As Integer
    Dim iSpdRow     As Integer
    Dim D_Chk_F     As Integer      ' 중복되는 Routine인지 체크 Flag
    Dim S_Chk_F     As Integer      ' 포함되는 Routine인지 체크 Flag
    Dim Rcd
    Dim sSpdTestCd
    Dim sRoutine    As String
    Dim sPrintNm    As String
    Dim sTestCd     As String
    Dim sItemGbn    As String
       
    If Len(txtroutinecd.Text) = txtroutinecd.MaxLength Then
' 이미 내려진 묶음처방인지 확인
        For iCnt = 1 To spdRoutine.MaxRows
            Call spdRoutine.GetText(3, iCnt, Rcd)
            If Rcd = txtrpartcd.Text & txtroutinecd.Text Then
                ViewMsg "이미 선택한 항목입니다."
                Call Txt_Highlight(txtroutinecd)
                Exit Sub
            End If
        Next iCnt
        
        Set DCJ0101 = New DCJ0101
        sRoutine = DCJ0101.Get_Routine(txtrpartcd.Text, txtroutinecd.Text, txtSpcCd.Text, UserCd)
        
        If Trim(sRoutine) = "" Then
        
        Else
' Routine명과 풀어진 검사정보를 화면에 표시한다.
    ' Routine명 표시
            spdRoutine.MaxRows = Rnow_row
            Call spdRoutine.SetText(1, Rnow_row, "1")
            Call spdRoutine.SetText(2, Rnow_row, Trim(GetByOne(sRoutine, sRoutine)))
            Call spdRoutine.SetText(3, Rnow_row, txtrpartcd.Text & txtroutinecd.Text)
            Rnow_row = Rnow_row + 1
            OrdCnt = Val(GetByOne(sRoutine, sRoutine))
    ' 풀어진 검사명과 코드를 화면에 표시한다.
            For iCnt = 1 To OrdCnt
            
                sPrintNm = GetByOne(sRoutine, sRoutine)
                sTestCd = GetByOne(sRoutine, sRoutine)
    ' 이미 내려진 검사항목코드와 중복되는 것이 있는지 검사
                D_Chk_F = False
                For iSpdRow = 1 To spdorder.MaxRows
                    RtnCd = spdorder.GetText(4, iSpdRow, sSpdTestCd)
                    If Trim(sSpdTestCd) = Trim(sTestCd) Then
                        D_Chk_F = True
                        Exit For
                    End If
                Next iSpdRow
                If D_Chk_F = False Then
                    spdorder.MaxRows = Onow_row
                    Call spdorder.SetText(1, Onow_row, "1")
                    If Mid(sTestCd, 10, 2) <> "NN" Then
                        Call spdorder.SetText(2, Onow_row, "S")
                    End If
                    Call spdorder.SetText(3, Onow_row, sPrintNm)
                    Call spdorder.SetText(4, Onow_row, sTestCd)
                    Call spdorder.SetText(5, Onow_row, txtrpartcd.Text & txtroutinecd.Text)
                    
                    Onow_row = Onow_row + 1
                Else
                    ViewMsg "중복되는 검사항목이 존재합니다. 확인하여 주십시요"
                End If
            Next iCnt
' 선택된 Routine 검사의 포함관계여부 판단
            S_Chk_F = True
            For iCnt = 1 To spdorder.MaxRows
                RtnCd = spdorder.GetText(5, iCnt, Rcd)
                If Rcd = txtrpartcd.Text & txtroutinecd.Text Then
                    S_Chk_F = False
                    Exit For
                End If
            Next iCnt
            If S_Chk_F = True Then
                spdRoutine.MaxRows = spdRoutine.MaxRows - 1
                Rnow_row = Rnow_row - 1
                ViewMsg "선택하신 Routine코드는 이미 내려지 Routine검사에 포함되는 검사입니다."
            End If
            
        End If
        
        Set DCJ0101 = Nothing
        
        Call Txt_Highlight(txtroutinecd)
    End If
    
End Sub

Private Sub txtroutinecd_GotFocus()
    
    Call Txt_Highlight(txtroutinecd)
    
End Sub

Private Sub txtroutinecd_KeyPress(KeyAscii As Integer)

    CodeHelp_F = False
    
End Sub

Private Sub txtsex_LostFocus()
    
    If txtdeptcd.Enabled = True Then
        txtdeptcd.SetFocus
    End If

End Sub


Private Sub txtslabdate_Change()

    If txtsslipcd.Visible = True And Len(txtslabdate.Text) = txtslabdate.MaxLength Then
        txtsslipcd.SetFocus
    End If
    
End Sub

Private Sub txtslabseq_Change()
    
    Dim RtnCd       As Integer
    Dim RCnt        As Integer
    Dim iCnt        As Integer
    Dim sLabDate    As String
    Dim sJubSu      As String
    Dim sPartCd     As String
    
    If Len(txtslabseq.Text) = txtslabseq.MaxLength Then
' 검체번호가 존재하는지 검사와 만약 존재하면 신상표시
        sLabDate = txtslabdate
        sPartCd = Left(txtsslipcd.Text, 1)
    
        Set DCJ0101 = New DCJ0101
    
' 해당일의 등록번호로 조회
        sJubSu = DCJ0101.Get_JubSuList(sLabDate, sPartCd, Mid(txtsslipcd.Text, 2, 2) & txtslabseq.Text)
    
        Set DCJ0101 = Nothing
    
' Spread에 데이타 표시
        If sJubSu <> "" Then
            RCnt = Val(GetByOne(sJubSu, sJubSu))
            For iCnt = 1 To RCnt
                spdsLabNo.MaxRows = iCnt
                Call spdsLabNo.SetText(1, iCnt, sLabDate)
                Call spdsLabNo.SetText(2, iCnt, Trim(GetByOne(sJubSu, sJubSu)))
                Call spdsLabNo.SetText(3, iCnt, Trim(GetByOne(sJubSu, sJubSu)))
            Next iCnt
            Call spdsLabNo_Click(1, 1)
        Else
            spdsLabNo.MaxRows = 0
            ViewMsg "해당 데이타가 없습니다."
        End If
    End If

End Sub

Private Sub txtslabseq_GotFocus()

    txtslabseq.Tag = txtslabseq.Text
    
End Sub


Private Sub txtslabseq_LostFocus()

    If txtslabseq.Text <> txtslabseq.Tag Then
        If Len(txtslabseq.Text) < txtslabseq.MaxLength Then
            txtslabseq.Text = Format(txtslabseq.Text, "00000")
        End If
    End If
    
End Sub

Private Sub txtSlipcd_Change()
    
    If Len(txtslipcd.Text) = txtslipcd.MaxLength Then
        
        Set DCJ0101 = New DCJ0101
        
        lblSlipNm.Caption = DCJ0101.Get_SlipNm(txtslipcd.Text)
    
        If lblSlipNm.Caption = "" Then
            ViewMsg "존재하지 않는 Slip Code입니다."
        Else
            If Trim(txtslipcd.Text) <> "" Then
                txtrpartcd.Text = Left(Trim(txtslipcd.Text), 1)
                txtoslipcd.Text = Trim(txtslipcd.Text)
            End If
        End If
    
        Set DCJ0101 = Nothing
        
        If CodeHelp_F = False Then
            txtSpcCd.SetFocus
        Else
            SendKeys "{ENTER}"
        End If
    Else
        lblSlipNm.Caption = ""
    End If

End Sub

Private Sub txtSlipcd_GotFocus()
    
    Call Txt_Highlight(txtslipcd)
    txtslipcd.Tag = txtslipcd.Text
    
End Sub

Private Sub txtSlipcd_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CodeHelp_F = False
    
End Sub

Private Sub txtSlipcd_LostFocus()

    If txtslipcd.Tag <> txtslipcd.Text Then
        Call search_clear
    End If
    
End Sub

Private Sub txtSpcCd_Change()
    
    If Len(txtSpcCd.Text) = txtSpcCd.MaxLength Then
        
        Set DCJ0101 = New DCJ0101
        
        lblSpcNm.Caption = Trim(DCJ0101.Get_SpcNm(txtSpcCd.Text))
    
        If lblSpcNm.Caption = "" Then
            ViewMsg "존재하지 않는 검체 Code입니다."
        Else
            txtospccd.Text = txtSpcCd.Text
        End If
    
        Set DCJ0101 = Nothing
        
        If CodeHelp_F = False Then
            txtRegNo.SetFocus
        Else
            SendKeys "{ENTER}"
        End If
    Else
        lblSpcNm.Caption = ""
    End If

End Sub

Private Sub txtSpcCd_GotFocus()
    
    Call Txt_Highlight(txtSpcCd)
    txtSpcCd.Tag = txtSpcCd.Text
    
End Sub

Private Sub txtSpcCd_KeyPress(KeyAscii As Integer)

    CodeHelp_F = False
    
End Sub

Private Sub txtSpcCd_LostFocus()

    If txtSpcCd.Tag <> txtSpcCd.Text Then
        Call search_clear
    End If
    
    If Len(txtSpcCd.Text) < txtSpcCd.MaxLength Then
        txtSpcCd.Text = Format(txtSpcCd.Text, "000")
    End If
    
End Sub

Private Sub txtspecial_Change()

    If Trim(txtspecial.Text) = "Y" Then
        txtspecial.Tag = "1"
        chkspecial.Value = 1
    Else
        txtspecial.Tag = "0"
        chkspecial.Value = 0
    End If
    
End Sub

Private Sub txtspecial_GotFocus()
    
    Call Txt_Highlight(txtspecial)
    
End Sub

Private Sub txtspecial_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txtsRegno_Change()

    Dim RtnCd       As Integer
    Dim RCnt        As Integer
    Dim iCnt        As Integer
    Dim sLabDate    As String
    Dim sJubSu      As String
    Dim sPartCd     As String
    
    If Len(txtsRegno.Text) = txtsRegno.MaxLength Then
' 해당일의 등록번호로 검체번호 조회
        sLabDate = Format(dtpLabDate.Value, "YYYYMMDD")
        sPartCd = Left(txtslipcd, 1)
    
        Set DCJ0101 = New DCJ0101
    
' 해당일의 등록번호로 조회
        sJubSu = DCJ0101.Get_JubSuList(sLabDate, sPartCd, "R" & txtsRegno.Text)
    
        Set DCJ0101 = Nothing
    
' Spread에 데이타 표시
        If sJubSu <> "" Then
            RCnt = Val(GetByOne(sJubSu, sJubSu))
            For iCnt = 1 To RCnt
                spdsLabNo.MaxRows = iCnt
                Call spdsLabNo.SetText(1, iCnt, sLabDate)
                Call spdsLabNo.SetText(2, iCnt, Trim(GetByOne(sJubSu, sJubSu)))
                Call spdsLabNo.SetText(3, iCnt, Trim(GetByOne(sJubSu, sJubSu)))
            Next iCnt
            spdsLabNo.SetFocus
        Else
            spdsLabNo.MaxRows = 0
            ViewMsg "해당 데이타가 없습니다."
            txtsRegno.SetFocus
            Call Txt_Highlight(txtsRegno)
        End If
    End If

End Sub

Private Sub txtsRegno_LostFocus()

    Dim RtnCd       As Integer
    Dim RCnt        As Integer
    Dim iCnt        As Integer
    Dim sLabDate    As String
    Dim sJubSu      As String
    Dim sPartCd     As String
    
' 잘못된 등록번호
    If IsNumeric(Trim(txtsRegno.Text)) = False Then
        ViewMsg "잘못된 등록번호입니다."
        Exit Sub
    End If
    
' 그냥 통과
    If Trim(txtsRegno.Text) = "" Then
        Exit Sub
    End If
' 해당일의 등록번호로 검체번호 조회
    sLabDate = Format(dtpLabDate.Value, "YYYYMMDD")
    sPartCd = Left(txtslipcd, 1)
    
    Set DCJ0101 = New DCJ0101
    
' 해당일의 등록번호로 조회
    sJubSu = DCJ0101.Get_JubSuList(sLabDate, sPartCd, "R" & txtsRegno.Text)
    
    Set DCJ0101 = Nothing
    
' Spread에 데이타 표시
    If sJubSu <> "" Then
        RCnt = Val(GetByOne(sJubSu, sJubSu))
        For iCnt = 1 To RCnt
            spdsLabNo.MaxRows = iCnt
            Call spdsLabNo.SetText(1, iCnt, sLabDate)
            Call spdsLabNo.SetText(2, iCnt, Trim(GetByOne(sJubSu, sJubSu)))
            Call spdsLabNo.SetText(3, iCnt, Trim(GetByOne(sJubSu, sJubSu)))
        Next iCnt
        spdsLabNo.SetFocus
    Else
        spdsLabNo.MaxRows = 0
        ViewMsg "해당 데이타가 없습니다."
    End If


End Sub


Private Sub txtsslipcd_Change()

    If Len(txtsslipcd.Text) = txtsslipcd.MaxLength Then
        If CodeHelp_F = False Then
            txtslabseq.SetFocus
        Else
            SendKeys "{ENTER}"
        End If
    End If
    
End Sub

Private Sub txtsslipcd_KeyPress(KeyAscii As Integer)

    CodeHelp_F = False
    
End Sub

Private Sub txtsslipcd_LostFocus()

    txtsslipcd.Text = UCase(txtsslipcd.Text)
    
End Sub


Private Sub txtward_GotFocus()
    
    Call Txt_Highlight(txtward)
    
End Sub

Private Sub txtward_LostFocus()

    If Len(Trim(txtward.Text)) = 0 Then
        optgubun(0).Value = True
    Else
        optgubun(1).Value = True
    End If
    
End Sub


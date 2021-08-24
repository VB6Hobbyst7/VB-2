VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGS0301 
   BorderStyle     =   0  '없음
   Caption         =   "환자데이터조회 - 이상자 체크"
   ClientHeight    =   7470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame2 
      Height          =   7455
      Left            =   3810
      TabIndex        =   9
      Top             =   0
      Width           =   7965
      _Version        =   65536
      _ExtentX        =   14049
      _ExtentY        =   13150
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
      Begin FPSpread.vaSpread spdList2 
         Height          =   4965
         Left            =   300
         OleObjectBlob   =   "FGS0301.frx":0000
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1770
         Width           =   7350
      End
      Begin Threed.SSFrame SSFrame6 
         Height          =   1545
         Left            =   4590
         TabIndex        =   11
         Top             =   150
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   2725
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Left            =   90
            TabIndex        =   12
            Top             =   480
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "이 름"
            ForeColor       =   0
            BackColor       =   12632256
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   315
            Left            =   90
            TabIndex        =   13
            Top             =   810
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "나 이"
            ForeColor       =   0
            BackColor       =   12632256
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Left            =   90
            TabIndex        =   14
            Top             =   1140
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "성 별"
            ForeColor       =   0
            BackColor       =   12632256
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   315
            Left            =   90
            TabIndex        =   15
            Top             =   150
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "등록번호"
            ForeColor       =   0
            BackColor       =   12632256
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
         Begin VB.Label lblName 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "남궁옥분씨애기"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1110
            TabIndex        =   19
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lblAge 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "130"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1110
            TabIndex        =   18
            Top             =   810
            Width           =   495
         End
         Begin VB.Label lblSex 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "소아"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1110
            TabIndex        =   17
            Top             =   1140
            Width           =   495
         End
         Begin VB.Label lblRegNo 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "720121-1840518"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1110
            TabIndex        =   16
            Top             =   150
            Width           =   1815
         End
      End
      Begin Threed.SSCommand cmdQuery 
         Height          =   945
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "조회 F3"
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FGS0301.frx":03A8
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   945
         Left            =   2760
         TabIndex        =   7
         Top             =   480
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "종료Esc"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FGS0301.frx":0C82
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   945
         Left            =   840
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "인쇄 F5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FGS0301.frx":155C
      End
      Begin Threed.SSPanel SSPanel18 
         Height          =   285
         Left            =   780
         TabIndex        =   20
         Top             =   6780
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "참 고 치"
         ForeColor       =   12640511
         BackColor       =   16576
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
         Alignment       =   8
      End
      Begin Threed.SSPanel SSPanel19 
         Height          =   285
         Left            =   4710
         TabIndex        =   21
         Top             =   6780
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "단  위"
         ForeColor       =   12640511
         BackColor       =   16576
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
         Alignment       =   8
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   285
         Left            =   780
         TabIndex        =   24
         Top             =   7080
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Panic"
         ForeColor       =   16777215
         BackColor       =   16512
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
         Alignment       =   8
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   285
         Left            =   4710
         TabIndex        =   26
         Top             =   7080
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Delta값"
         ForeColor       =   16777215
         BackColor       =   16512
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
         Alignment       =   8
      End
      Begin VB.Label lblDelta단위 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '단일 고정
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7170
         TabIndex        =   28
         Top             =   7080
         Width           =   285
      End
      Begin VB.Label lblDelta 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '단일 고정
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5760
         TabIndex        =   27
         Top             =   7080
         Width           =   1395
      End
      Begin VB.Label lblPanic 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '단일 고정
         Caption         =   "12.5 - 84.7"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1830
         TabIndex        =   25
         Top             =   7080
         Width           =   2625
      End
      Begin VB.Label lbl참고치 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00EAEAFF&
         BorderStyle     =   1  '단일 고정
         Caption         =   "12.5 - 84.7"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1830
         TabIndex        =   23
         Top             =   6780
         Width           =   2625
      End
      Begin VB.Label lbl단위 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00EAEAFF&
         BorderStyle     =   1  '단일 고정
         Caption         =   "g/dl"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5760
         TabIndex        =   22
         Top             =   6780
         Width           =   1695
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   7455
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   3765
      _Version        =   65536
      _ExtentX        =   6641
      _ExtentY        =   13150
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
      Begin FPSpread.vaSpread spdList1 
         Height          =   3825
         Left            =   120
         OleObjectBlob   =   "FGS0301.frx":1E36
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3540
         Width           =   3525
      End
      Begin VB.TextBox txtSlip 
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
         Left            =   210
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "H02"
         Top             =   570
         Width           =   495
      End
      Begin Threed.SSPanel pnlSlip 
         Height          =   345
         Left            =   90
         TabIndex        =   31
         Top             =   210
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   609
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
      Begin Threed.SSCommand cmdSlipHelp 
         Height          =   330
         Left            =   720
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   570
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
         Picture         =   "FGS0301.frx":2EE6
      End
      Begin Threed.SSFrame SSFrame5 
         Height          =   615
         Left            =   120
         TabIndex        =   33
         Top             =   2280
         Width           =   3525
         _Version        =   65536
         _ExtentX        =   6218
         _ExtentY        =   1085
         _StockProps     =   14
         Caption         =   "조회 Option"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optOr 
            Appearance      =   0  '평면
            BackColor       =   &H00FF8080&
            Caption         =   "OR"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1920
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton optAnd 
            Appearance      =   0  '평면
            BackColor       =   &H008080FF&
            Caption         =   "AND"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   540
            TabIndex        =   3
            Top             =   270
            Width           =   1155
         End
      End
      Begin MSComCtl2.DTPicker dtpSLabDate 
         Height          =   315
         Left            =   210
         TabIndex        =   0
         Top             =   1320
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
         CustomFormat    =   "yyy-MM-dd"
         Format          =   24772611
         CurrentDate     =   36165
      End
      Begin Threed.SSPanel pnlLabDate 
         Height          =   345
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   609
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
      Begin MSComCtl2.DTPicker dtpELabDate 
         Height          =   315
         Left            =   1950
         TabIndex        =   1
         Top             =   1320
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
         CustomFormat    =   "yyy-MM-dd"
         Format          =   24772611
         CurrentDate     =   36165
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   615
         Left            =   120
         TabIndex        =   36
         Top             =   2910
         Width           =   3525
         _Version        =   65536
         _ExtentX        =   6218
         _ExtentY        =   1085
         _StockProps     =   14
         Caption         =   "Abnormal Option"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optHigh 
            Appearance      =   0  '평면
            BackColor       =   &H008080FF&
            Caption         =   "High"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   930
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.OptionButton optLow 
            Appearance      =   0  '평면
            BackColor       =   &H008080FF&
            Caption         =   "Low"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   4
            Top             =   270
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.CheckBox chkLow 
            Appearance      =   0  '평면
            BackColor       =   &H00C0C0FF&
            Caption         =   "Low"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   40
            Tag             =   "1"
            Top             =   270
            Width           =   675
         End
         Begin VB.CheckBox chkHigh 
            Appearance      =   0  '평면
            BackColor       =   &H00C0C0FF&
            Caption         =   "High"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   930
            TabIndex        =   39
            Tag             =   "2"
            Top             =   270
            Width           =   675
         End
         Begin VB.CheckBox chkPanic 
            Appearance      =   0  '평면
            BackColor       =   &H00C0C0FF&
            Caption         =   "Panic"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1680
            TabIndex        =   38
            TabStop         =   0   'False
            Tag             =   "4"
            Top             =   270
            Width           =   795
         End
         Begin VB.CheckBox chkDelta 
            Appearance      =   0  '평면
            BackColor       =   &H00C0C0FF&
            Caption         =   "Delta"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2550
            TabIndex        =   37
            TabStop         =   0   'False
            Tag             =   "8"
            Top             =   270
            Width           =   795
         End
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   615
         Left            =   120
         TabIndex        =   42
         Top             =   1650
         Width           =   3525
         _Version        =   65536
         _ExtentX        =   6218
         _ExtentY        =   1085
         _StockProps     =   14
         Caption         =   "판정 구분"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optLH 
            Appearance      =   0  '평면
            BackColor       =   &H00C0C0C0&
            Caption         =   "Low/High"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   90
            TabIndex        =   2
            Top             =   300
            Width           =   1125
         End
         Begin VB.OptionButton optNP 
            Appearance      =   0  '평면
            BackColor       =   &H00C0C0C0&
            Caption         =   "Neg/Pos"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   1230
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   300
            Width           =   1065
         End
         Begin VB.OptionButton optOF 
            Appearance      =   0  '평면
            BackColor       =   &H00C0C0C0&
            Caption         =   "OtherFlag"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2340
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   300
            Width           =   1125
         End
      End
      Begin VB.Label Label2 
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1740
         TabIndex        =   46
         Top             =   1380
         Width           =   195
      End
      Begin VB.Label lblSlip 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "일반혈액검사"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1020
         TabIndex        =   45
         Top             =   570
         Width           =   2535
      End
   End
End
Attribute VB_Name = "FGS0301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim slip As Boolean

Dim low$, high$, panic$, delta$     'Abnormal Option에 대한 각 값들("0":False, "1":True)

Dim 판정Gbn%, 조회Opt%              '판정구분과 조회옵션에 대한 값들

Dim IsTextRow As Integer            '작업번호를 하나씩만 뿌려주기 위한 변수(마지막으로 wirte한 행 기억)


Private Sub chkDelta_Click()
    If delta = "0" Then
        delta = "1"
    Else
        delta = "0"
    End If

End Sub

Private Sub chkHigh_Click()
    If high = "0" Then
        high = "1"
    Else
        high = "0"
    End If
    high = chkHigh.Value
End Sub

Private Sub chkLow_Click()
    If low = "0" Then
        low = "1"
    Else
        low = "0"
    End If
    low = chkLow.Value
End Sub

Private Sub chkPanic_Click()
    If panic = "0" Then
        panic = "1"
    Else
        panic = "0"
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{ESC}"
    End If

End Sub

Private Sub cmdPrint_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{F5}"
    End If

End Sub

Private Sub cmdQuery_Click()
    Dim DCS0301 As DCS0301
    Dim count As Integer
    Dim spdRow As Integer
    Dim code$
    Dim abnormalCd$, deltaGbn$, refGbn$
    Dim male As Boolean     '남자 여자 판별을 위해
    Dim recordField01$, recordField02$, recordField03$, recordField04$
    Dim recordField05$, recordField06$, recordField07$, recordField08$
    Dim recordField09$, recordField10$, recordField11$, recordField12$
    Dim recordField13$, recordField14$, recordField15$, recordField16$
    Dim recordField17$, recordField18$, recordField19$, recordField20$
    Dim recordField21$, recordField22$, recordField23$, recordField24$
    Dim recordField25$, recordField26$, recordField27$, recordField28$
    Dim recordField29$

    Call SpreadClear(2)
    Call lblClear

    If slip = False Then
        ViewMsg "SLIP을 정확히 입력하십시오."
        txtSlip.SetFocus
        Exit Sub
    End If

    If 판정Gbn = 1 And low = "0" And high = "0" And delta = "0" And panic = "0" Then
        ViewMsg "Abnormal Option이 한개도 선택되지 않았습니다."
        Exit Sub
    
    ElseIf 판정Gbn = 2 And high = "0" Then
        ViewMsg "Abnormal Option이 한개도 선택되지 않았습니다."
        Exit Sub
    
    ElseIf 판정Gbn = 3 And high = "0" Then
        ViewMsg "Abnormal Option이 한개도 선택되지 않았습니다."
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    Set DCS0301 = New DCS0301

    count = 0
    For spdRow = 1 To spdList1.MaxRows
        spdList1.Col = 1
        spdList1.Row = spdRow
        If spdList1.text = "1" Then
            spdList1.Col = 3
            code = code & spdList1.text & "|"
            count = count + 1
        End If
    Next

    With DCS0301

        .Get_Contents txtSlip.text, Format(dtpSLabDate.Value, "YYYYMMDD"), _
                Format(dtpELabDate.Value, "YYYYMMDD"), 판정Gbn, 조회Opt, _
                low & high & panic & delta, code, count
            
        count = .Getcount
        If count = 0 Then
            ViewMsg "조회된 내용이 없습니다."
            Set DCS0301 = Nothing
            spdList2.MaxRows = count
            MousePointer = 0
            Exit Sub
        End If
        
        recordField01 = .GetrecordField01
        recordField02 = .GetrecordFiled02
        recordField03 = .GetrecordFiled03
        recordField04 = .GetrecordFiled04
        recordField05 = .GetrecordFiled05
        recordField06 = .GetrecordFiled06
        recordField07 = .GetrecordFiled07
        recordField08 = .GetrecordFiled08
        recordField09 = .GetrecordFiled09
        recordField10 = .GetrecordFiled10
        recordField11 = .GetrecordFiled11
        recordField12 = .GetrecordFiled12
        recordField13 = .GetrecordFiled13
        recordField14 = .GetrecordFiled14
        recordField15 = .GetrecordFiled15
        recordField16 = .GetrecordFiled16
        recordField17 = .GetrecordFiled17
        recordField18 = .GetrecordFiled18
        recordField19 = .GetrecordFiled19
        recordField20 = .GetrecordFiled20
        recordField21 = .GetrecordFiled21
        recordField22 = .GetrecordFiled22
        recordField23 = .GetrecordFiled23
        recordField24 = .GetrecordFiled24
        recordField25 = .GetrecordFiled25
        recordField26 = .GetrecordFiled26
        recordField27 = .GetrecordFiled27
        recordField28 = .GetrecordFiled28
        recordField29 = .GetrecordFiled29

    End With

    Set DCS0301 = Nothing

    With spdList2

    For spdRow = 1 To count

        .MaxRows = spdRow
        
        '작업번호(COL1)
        Call spdCmp(1, spdRow, GetByOne(recordField01, recordField01) & "-" & txtSlip.text _
                    & "-" & GetByOne(recordField02, recordField02))

        '검사항목명(COL2)
        Call .SetText(2, spdRow, GetByOne(recordField03, recordField03))
        
        '결과값(COL3)
        Call .SetText(3, spdRow, GetByOne(recordField04, recordField04))

        'RPD(COL4,5,6)
        
        'ABNORMALCD를 저장
        abnormalCd = GetByOne(recordField05, recordField05)

        If 판정Gbn = 1 Then
            Select Case abnormalCd

                Case "1": Call .SetText(4, spdRow, "L")
                Case "2": Call .SetText(4, spdRow, "H")
                Case "4": Call .SetText(5, spdRow, "P")
                Case "8": Call .SetText(6, spdRow, "D")
                Case "5": Call .SetText(4, spdRow, "L")
                          Call .SetText(5, spdRow, "P")
                Case "9": Call .SetText(4, spdRow, "L")
                          Call .SetText(6, spdRow, "D")
                Case "10": Call .SetText(4, spdRow, "H")
                           Call .SetText(6, spdRow, "D")
                Case "12": Call .SetText(5, spdRow, "P")
                           Call .SetText(6, spdRow, "D")
                Case "13": Call .SetText(4, spdRow, "L")
                           Call .SetText(5, spdRow, "P")
                           Call .SetText(6, spdRow, "D")
                Case "14": Call .SetText(4, spdRow, "H")
                           Call .SetText(5, spdRow, "P")
                           Call .SetText(6, spdRow, "D")
            End Select
            
            Call Skip_Data(recordField29)
            
        ElseIf 판정Gbn = 2 And abnormalCd = "16" Then
            Call .SetText(4, spdRow, "P")
            Call Skip_Data(recordField29)
        ElseIf 판정Gbn = 3 And abnormalCd = "32" Then
            Call .SetText(4, spdRow, GetByOne(recordField29, recordField29))
        Else
            Call Skip_Data(recordField29)
        End If

        '등록번호(COL7)
        Call .SetText(7, spdRow, GetByOne(recordField06, recordField06))

        '이름(COL8)
        Call .SetText(8, spdRow, GetByOne(recordField07, recordField07))
        
        '나이(COL9)
        Call .SetText(9, spdRow, GetByOne(recordField08, recordField08))
        
        '성별(COL10)
        Select Case GetByOne(recordField09, recordField09)
            Case "1": Call .SetText(10, spdRow, "남")
                      male = True
            Case "2": Call .SetText(10, spdRow, "여")
                      male = False
            Case "3": Call .SetText(10, spdRow, "남")
                      male = True
            Case "4": Call .SetText(10, spdRow, "여")
                      male = False
        End Select

        '참고치(COL11)

        refGbn = GetByOne(recordField10, recordField10)

        If refGbn = "2" Then
            If male Then
                Call .SetText(11, spdRow, GetByOne(recordField11, recordField11) & "(-" _
                            & GetByOne(recordField15, recordField15) & ")  -  " _
                            & GetByOne(recordField17, recordField17) & "(+" _
                            & GetByOne(recordField21, recordField21) & ")")
                
                Call Skip_Data(recordField12)
                Call Skip_Data(recordField16)
                Call Skip_Data(recordField18)
                Call Skip_Data(recordField22)
            
            Else
                Call .SetText(11, spdRow, GetByOne(recordField12, recordField12) & "(-" _
                            & GetByOne(recordField16, recordField16) & ")  -  " _
                            & GetByOne(recordField18, recordField18) & "(+" _
                            & GetByOne(recordField22, recordField22) & ")")
                
                Call Skip_Data(recordField11)
                Call Skip_Data(recordField15)
                Call Skip_Data(recordField17)
                Call Skip_Data(recordField21)
            
            End If

            Call Skip_Data(recordField13)
            Call Skip_Data(recordField14)
            Call Skip_Data(recordField19)
            Call Skip_Data(recordField20)

        ElseIf refGbn = "3" Then
            If male Then
                Call .SetText(11, spdRow, "< " & GetByOne(recordField19, recordField19) _
                            & "(+" & GetByOne(recordField21, recordField21) & ")")
                
                Call Skip_Data(recordField20)
                Call Skip_Data(recordField22)

            Else
                Call .SetText(11, spdRow, "< " & GetByOne(recordField20, recordField20) _
                            & "(+" & GetByOne(recordField22, recordField22) & ")")
                
                Call Skip_Data(recordField19)
                Call Skip_Data(recordField21)
            End If
            
            Call Skip_Data(recordField11)
            Call Skip_Data(recordField12)
            Call Skip_Data(recordField13)
            Call Skip_Data(recordField14)
            Call Skip_Data(recordField15)
            Call Skip_Data(recordField16)
            Call Skip_Data(recordField17)
            Call Skip_Data(recordField18)

        ElseIf refGbn = "4" Then
            If male Then
                Call .SetText(11, spdRow, "> " & GetByOne(recordField13, recordField13) _
                            & "(-" & GetByOne(recordField15, recordField15) & ")")
                            
                Call Skip_Data(recordField14)
                Call Skip_Data(recordField16)

            Else
                Call .SetText(11, spdRow, "> " & GetByOne(recordField14, recordField14) _
                            & "(-" & GetByOne(recordField16, recordField16) & ")")
                
                Call Skip_Data(recordField13)
                Call Skip_Data(recordField15)
            End If
            
            Call Skip_Data(recordField11)
            Call Skip_Data(recordField12)
            Call Skip_Data(recordField17)
            Call Skip_Data(recordField18)
            Call Skip_Data(recordField19)
            Call Skip_Data(recordField20)
            Call Skip_Data(recordField21)
            Call Skip_Data(recordField22)

        Else

            Call Skip_Data(recordField11)
            Call Skip_Data(recordField12)
            Call Skip_Data(recordField13)
            Call Skip_Data(recordField14)
            Call Skip_Data(recordField15)
            Call Skip_Data(recordField16)
            Call Skip_Data(recordField17)
            Call Skip_Data(recordField18)
            Call Skip_Data(recordField19)
            Call Skip_Data(recordField20)
            Call Skip_Data(recordField21)
            Call Skip_Data(recordField22)

        End If
            

        '단위(COL12)
        Call .SetText(12, spdRow, GetByOne(recordField23, recordField23))

        'Panic(COL13)
        If GetByOne(recordField24, recordField24) = "1" Then
            Call .SetText(13, spdRow, GetByOne(recordField25, recordField25) & " - " & GetByOne(recordField26, recordField26))
        Else
            Call Skip_Data(recordField25)
            Call Skip_Data(recordField26)
        End If

        'Delta(COL14,15)
        
        deltaGbn = GetByOne(recordField27, recordField27)

        If deltaGbn = "1" Then
            Call .SetText(14, spdRow, GetByOne(recordField28, recordField28))
        ElseIf deltaGbn = "2" Then
            Call .SetText(14, spdRow, GetByOne(recordField28, recordField28))
            Call .SetText(15, spdRow, "%")
        Else
            Call Skip_Data(recordField28)
        End If

    Next

    End With

    MousePointer = 0

    ViewMsg "총 " & count & "개의 자료가 조회되었습니다."

End Sub

Private Sub Skip_Data(data As String)
    
    '파이프 앞의 문자열을 끊어버린다.
    Dim dummy$
    
    dummy = GetByOne(data, data)

End Sub
Private Sub spdCmp(ByVal Col As Long, ByVal Row As Long, ByVal text As String)

    '위 col의 text와 비교해서 같으면 뿌리지 않는다.
    Dim data

    With spdList2

        If Row = 1 Then
            Call .SetText(Col, Row, text)
            IsTextRow = 1
            Exit Sub
        Else
            Call .GetText(Col, IsTextRow, data)
            If data <> text Then
                Call .SetText(Col, Row, text)
                IsTextRow = Row
            End If
        End If
    End With

End Sub

Private Sub cmdQuery_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{F3}"
    End If

End Sub

Private Sub cmdSlipHelp_Click()
    Dim i%
    Dim count%
    Dim DCS0301 As DCS0301
    Dim recordField01$, recordField02$, recordField03$
    
    Set DCS0301 = New DCS0301
    
    DCS0301.Get_SlipCD
    
    count = DCS0301.Getcount
    
    Erase gCodeHlpTable '배열 초기화

    ReDim gCodeHlpTable(count) As CodeTBL

    With DCS0301
        recordField01 = .GetrecordField01
        recordField02 = .GetrecordFiled02
        recordField03 = .GetrecordFiled03
    End With

    Set DCS0301 = Nothing

    For i = 1 To count
        gCodeHlpTable(i).sCode = GetByOne(recordField01, recordField01) & GetByOne(recordField02, recordField02)
        gCodeHlpTable(i).sCodeNm = GetByOne(recordField03, recordField03)
    Next
    
    giCodeHlpCnt = count

    hWndCd = txtSlip.hwnd

    FSS0301.Left = 2500
    FSS0301.Top = 1570

    Load FSS0301
    FSS0301.Show vbModal

    'txtSlip과 lblSlip에 조회내용을 표시한다.
    Call txtSlip_LostFocus
    
    Call GetItem

End Sub

Private Sub dtpELabDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub dtpSLabDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3: Call cmdQuery_Click       '/* 조 회 */
'        Case vbKeyF5: Call cmdPrint_Click       '/* 인 쇄 */
        Case vbKeyEscape: Unload Me             '/* 종 료 */
    End Select

End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    Me.Width = 11920
    Me.Height = 7950

    FGS0301.KeyPreview = True
    
    '폼이 로드될때 디폴트값 설정과 화면 클리어
    txtSlip = fCurUserSlipCd
    lblSlip = fCurUserSlipNm
    dtpSLabDate.Value = Format(Now, "yyyy-mm-dd")
    dtpELabDate.Value = Format(Now, "yyyy-mm-dd")
    optLH.Value = True
    optOr.Value = True

    With spdList1
        .Row = -1
        .Col = -1
        .Lock = True
        
        .Col = 1
        .Lock = False
    End With

    Call SpreadClear(2)

    Call lblClear
    
    slip = True
    low = "0"
    high = "0"
    panic = "0"
    delta = "0"
    판정Gbn = 1
    조회Opt = 2

    Call GetItem

End Sub

Private Sub lblClear()
    lblRegNo.Caption = ""
    lblName.Caption = ""
    lblAge.Caption = ""
    lblSex.Caption = ""
    lbl참고치.Caption = ""
    lbl단위.Caption = ""
    lblPanic.Caption = ""
    lblDelta.Caption = ""
    lblDelta단위.Caption = ""

End Sub
Private Sub SpreadClear(no As Integer)

    '폼이 로드되거나 조회를 실행할때 스프레드 내용을 지운다.
    If no = 1 Then
        With spdList1
            .Row = -1
            .Col = -1
            .text = ""
            .BackColor = 연하늘
        End With
    Else
        With spdList2
            .Row = -1
            .Col = -1
            .text = ""
            .BackColor = 연하늘
        End With
    End If

End Sub

Private Sub optAnd_GotFocus()
    If optLH.Value = True Then
        chkLow.Visible = False
        chkHigh.Visible = False
        optLow.Visible = True
        optHigh.Visible = True
        optLow.Value = False
        optHigh.Value = False
        chkLow.Value = 0
        chkHigh.Value = 0
        
        low = 0
        high = 0
    End If

    조회Opt = 1

End Sub

Private Sub optHigh_GotFocus()
    If high = "0" Then
        high = "1"
        low = "0"
        chkHigh.Value = 1
        chkLow.Value = 0
    Else
        high = "0"
        low = "1"
        chkHigh.Value = 0
        chkLow.Value = 1
    End If

End Sub

Private Sub optLH_GotFocus()
    If optAnd.Value = True Then
        optLow.Visible = True
        optHigh.Visible = True
        chkLow.Visible = False
        chkHigh.Visible = False
    Else
        optLow.Visible = False
        optHigh.Visible = False
        chkLow.Visible = True
        chkHigh.Visible = True
    End If
    
    optAnd.Enabled = True
    optOr.Enabled = True

    optHigh.Caption = "High"
    chkHigh.Caption = "High"
    chkPanic.Enabled = True
    chkDelta.Enabled = True

    판정Gbn = 1
    
    Call GetItem

End Sub

Private Sub optLow_GotFocus()
    If low = "0" Then
        low = "1"
        high = "0"
        chkHigh.Value = 0
        chkLow.Value = 1
    Else
        low = "0"
        high = "1"
        chkLow.Value = 0
        chkHigh = 1
    End If

End Sub

Private Sub optNP_GotFocus()
    optAnd.Enabled = False
    optOr.Enabled = False
    optLow.Visible = False
    optHigh.Visible = False
    chkLow.Visible = False
    chkHigh.Visible = True
    chkLow.Value = 0
    optLow.Value = False
    low = "0"
    chkHigh.Value = 1
    high = "1"
    chkHigh.Caption = "Pos"
    chkPanic.Enabled = False
    chkDelta.Enabled = False
    판정Gbn = 2
    
    Call GetItem

End Sub

Private Sub optOF_GotFocus()
    optAnd.Enabled = False
    optOr.Enabled = False
    optLow.Visible = False
    optHigh.Visible = False
    chkLow.Visible = False
    chkHigh.Visible = True
    chkLow.Value = 0
    optLow.Value = False
    low = "0"
    chkHigh.Value = 1
    high = "1"
    chkHigh.Caption = "OF"
    chkPanic.Enabled = False
    chkDelta.Enabled = False

    판정Gbn = 3
    
    Call GetItem

End Sub

Private Sub optOr_GotFocus()
    If optLH.Value = True Then
        optLow.Visible = False
        optHigh.Visible = False
        chkLow.Visible = True
        chkHigh.Visible = True
        chkLow.Value = 0
        chkHigh.Value = 0
        optLow.Value = False
        optHigh.Value = False

        low = 0
        high = 0
    End If

    조회Opt = 2

End Sub

Private Sub pnlLabDate_DblClick()
    If pnlLabDate.Caption = "접수일자" Then
        pnlLabDate.Caption = "결과완료일"
    ElseIf pnlLabDate.Caption = "결과완료일" Then
        pnlLabDate.Caption = "접수일자"
    End If
End Sub

Private Sub GetItem()
    Dim DCS0301 As DCS0301
    Dim count As Integer
    Dim spdRow As Integer
    Dim recordField01$, recordField02$, recordField03$, recordField04$
    Dim dummy$          'SUB항목이 있는 항목을 그냥 넘기기 위해 사용

    Call SpreadClear(1)
    
    If slip = False Then
        ViewMsg "SLIP을 정확히 입력하십시오."
        txtSlip.SetFocus
        Exit Sub
    End If
    
    Set DCS0301 = New DCS0301

    With DCS0301
        .Get_Item txtSlip.text, 판정Gbn
        count = .Getcount
        If count = 0 Then
            ViewMsg "조회된 내용이 없습니다."
            Set DCS0301 = Nothing
            spdList1.MaxRows = spdRow
            Exit Sub
        End If

        recordField01 = .GetrecordField01
        recordField02 = .GetrecordFiled02
        recordField03 = .GetrecordFiled03
        recordField04 = .GetrecordFiled04
    End With

    Set DCS0301 = Nothing

    For spdRow = 1 To count

        With spdList1
        
        .MaxRows = spdRow
        .Row = -1
        .Col = -1
        .Lock = True
        
        .Col = 1
        .Lock = False
        
        If Left(recordField04, 2) = "00" Then
            dummy = GetByOne(recordField01, recordField01)
            dummy = GetByOne(recordField02, recordField02)
            dummy = GetByOne(recordField03, recordField03)
            dummy = GetByOne(recordField03, recordField03)
            .MaxRows = .MaxRows - 1
            GoTo Skip
        End If

        Call .SetText(1, spdRow, "1")
        Call .SetText(2, spdRow, GetByOne(recordField01, recordField01))

        Call .SetText(3, spdRow, GetByOne(recordField02, recordField02) _
            & GetByOne(recordField03, recordField03) & GetByOne(recordField04, recordField04))
        
        End With

Skip:
    Next

    ViewMsg "총 " & count & "개의 자료가 조회되었습니다."

End Sub

Private Sub spdList2_Click(ByVal Col As Long, ByVal Row As Long)
    Dim regno, name, age, sex As Variant
    Dim 참고치, 단위, panic, delta, delta단위 As Variant

    Call spdReverse(spdList2, 1, 10, Row, Row, 연빨강, 2)
    Call spdList2.GetText(7, Row, regno)
    Call spdList2.GetText(8, Row, name)
    Call spdList2.GetText(9, Row, age)
    Call spdList2.GetText(10, Row, sex)
    Call spdList2.GetText(11, Row, 참고치)
    Call spdList2.GetText(12, Row, 단위)
    Call spdList2.GetText(13, Row, panic)
    Call spdList2.GetText(14, Row, delta)
    Call spdList2.GetText(15, Row, delta단위)

    lblRegNo.Caption = regno
    lblName.Caption = name
    lblAge.Caption = age
    lblSex.Caption = sex
    lbl참고치.Caption = 참고치
    lbl단위.Caption = 단위
    lblPanic.Caption = panic
    lblDelta.Caption = delta
    lblDelta단위.Caption = delta단위

End Sub

Private Sub txtSlip_LostFocus()
    
    Dim DCS0301 As DCS0301
    
    Set DCS0301 = New DCS0301
    
    DCS0301.Get_SlipNm txtSlip.text

    With DCS0301

        If .Getcount = 0 Then
            ViewMsg "입력하신 SLIP은 존재하지 않습니다."
            slip = False
            lblSlip.Caption = ""
            Exit Sub
        End If

        lblSlip.Caption = .GetrecordField01

    End With

    Set DCS0301 = Nothing
    
    txtSlip.text = UCase$(txtSlip.text)
    
    slip = True

End Sub

Private Sub txtSlip_Change()
    If txtSlip.SelStart = txtSlip.MaxLength Then SendKeys "{TAB}"
End Sub

Private Sub txtSlip_GotFocus()
    txtSlip.SelStart = 0
    txtSlip.SelLength = txtSlip.MaxLength

End Sub

Private Sub txtSlip_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub



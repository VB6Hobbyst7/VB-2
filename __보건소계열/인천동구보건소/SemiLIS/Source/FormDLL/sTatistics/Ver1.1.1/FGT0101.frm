VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGT0101 
   BorderStyle     =   0  '없음
   Caption         =   "통계 - 일월년 검사건수"
   ClientHeight    =   7305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11790
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel pnlFooter 
      Height          =   1125
      Left            =   45
      TabIndex        =   33
      Top             =   6150
      Width           =   11715
      _Version        =   65536
      _ExtentX        =   20664
      _ExtentY        =   1984
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
      Begin VB.ComboBox cboCondition 
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
         Left            =   3690
         Style           =   2  '드롭다운 목록
         TabIndex        =   37
         Top             =   360
         Width           =   2055
      End
      Begin Threed.SSFrame fraTGbn 
         Height          =   1005
         Left            =   6870
         TabIndex        =   34
         Top             =   30
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   1773
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
         Begin VB.OptionButton OptTGbn 
            Appearance      =   0  '평면
            BackColor       =   &H00FFC0C0&
            Caption         =   "모두"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   9
            Top             =   210
            Width           =   1245
         End
         Begin VB.OptionButton OptTGbn 
            Appearance      =   0  '평면
            BackColor       =   &H00C0C0FF&
            Caption         =   "이상자"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   150
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   570
            Width           =   1245
         End
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   945
         Left            =   10590
         TabIndex        =   12
         Top             =   90
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
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
         Picture         =   "FGT0101.frx":0000
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   945
         Left            =   8550
         TabIndex        =   10
         Top             =   90
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
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
         Picture         =   "FGT0101.frx":08DA
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   945
         Left            =   9570
         TabIndex        =   11
         Top             =   90
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
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
         Picture         =   "FGT0101.frx":11B4
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Left            =   2310
         TabIndex        =   38
         Top             =   360
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "접수구분"
         ForeColor       =   8454143
         BackColor       =   8388608
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
      End
   End
   Begin Threed.SSFrame fraMain 
      Height          =   1395
      Left            =   30
      TabIndex        =   13
      Top             =   30
      Width           =   11745
      _Version        =   65536
      _ExtentX        =   20717
      _ExtentY        =   2461
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
      Begin Threed.SSFrame fraDate 
         Height          =   735
         Left            =   90
         TabIndex        =   19
         Top             =   540
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
         _ExtentY        =   1296
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
         Begin Threed.SSPanel pnloptDate 
            Height          =   555
            Left            =   60
            TabIndex        =   36
            Top             =   120
            Width           =   3045
            _Version        =   65536
            _ExtentX        =   5371
            _ExtentY        =   979
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
            BevelOuter      =   0
            Begin VB.OptionButton optDate 
               Appearance      =   0  '평면
               BackColor       =   &H00C0E0FF&
               Caption         =   "일별"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   0
               Left            =   90
               TabIndex        =   2
               Top             =   120
               Width           =   795
            End
            Begin VB.OptionButton optDate 
               Appearance      =   0  '평면
               BackColor       =   &H00C0C0FF&
               Caption         =   "월별"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   1
               Left            =   1050
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   120
               Width           =   795
            End
            Begin VB.OptionButton optDate 
               Appearance      =   0  '평면
               BackColor       =   &H00FFC0C0&
               Caption         =   "년별"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   2
               Left            =   2010
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   120
               Width           =   795
            End
         End
         Begin VB.TextBox txtSHour 
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
            Left            =   9210
            MaxLength       =   2
            TabIndex        =   7
            Text            =   "HH"
            Top             =   240
            Width           =   405
         End
         Begin VB.TextBox txtEHour 
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
            Left            =   10110
            MaxLength       =   2
            TabIndex        =   8
            Text            =   "HH"
            Top             =   240
            Width           =   405
         End
         Begin MSComCtl2.DTPicker dtpSLabDate 
            Height          =   315
            Left            =   4530
            TabIndex        =   5
            Top             =   240
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
            Format          =   24510467
            CurrentDate     =   36165
         End
         Begin Threed.SSPanel pnlLabDate 
            Height          =   345
            Left            =   3150
            TabIndex        =   20
            Top             =   240
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "접수일 구간"
            ForeColor       =   8454143
            BackColor       =   8388608
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
         End
         Begin MSComCtl2.DTPicker dtpELabDate 
            Height          =   315
            Left            =   6270
            TabIndex        =   6
            Top             =   240
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
            Format          =   24510467
            CurrentDate     =   36165
         End
         Begin Threed.SSPanel pnlLabTime 
            Height          =   345
            Left            =   7830
            TabIndex        =   22
            Top             =   240
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "접수시간"
            ForeColor       =   8454143
            BackColor       =   8388608
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
         End
         Begin VB.Label lbl시 
            Caption         =   "시"
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
            Left            =   9630
            TabIndex        =   29
            Top             =   270
            Width           =   255
         End
         Begin VB.Label lbl시 
            Caption         =   "시"
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
            Left            =   10530
            TabIndex        =   28
            Top             =   270
            Width           =   255
         End
         Begin VB.Label lblSE 
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
            Left            =   9870
            TabIndex        =   23
            Top             =   270
            Width           =   195
         End
         Begin VB.Label lblinterval 
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
            Left            =   6060
            TabIndex        =   21
            Top             =   300
            Width           =   195
         End
      End
      Begin VB.TextBox txtSlipCd 
         BackColor       =   &H00FFFFFF&
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
         IMEMode         =   8  '영문
         Left            =   6570
         MaxLength       =   3
         TabIndex        =   1
         Top             =   210
         Width           =   480
      End
      Begin VB.TextBox txtPart 
         BackColor       =   &H00FFFFFF&
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
         IMEMode         =   8  '영문
         Left            =   1470
         MaxLength       =   1
         TabIndex        =   0
         Top             =   210
         Width           =   300
      End
      Begin Threed.SSPanel pnlPartCd 
         Height          =   345
         Left            =   90
         TabIndex        =   14
         Top             =   180
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "PART 코드"
         ForeColor       =   8454143
         BackColor       =   8388608
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
      End
      Begin Threed.SSPanel pnlSlipCd 
         Height          =   345
         Left            =   5190
         TabIndex        =   16
         Top             =   180
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "SLIP"
         ForeColor       =   0
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
         BorderWidth     =   1
      End
      Begin Threed.SSCommand cmdSliph 
         Height          =   300
         Left            =   7050
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   210
         Width           =   270
         _Version        =   65536
         _ExtentX        =   476
         _ExtentY        =   529
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
         Picture         =   "FGT0101.frx":1A8E
      End
      Begin Threed.SSCommand cmdParth 
         Height          =   300
         Left            =   1770
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   210
         Width           =   270
         _Version        =   65536
         _ExtentX        =   476
         _ExtentY        =   529
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
         Picture         =   "FGT0101.frx":1BB0
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
         Left            =   7350
         TabIndex        =   18
         Top             =   210
         Width           =   2775
      End
      Begin VB.Label lblPartNm 
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
         Left            =   2070
         TabIndex        =   15
         Top             =   195
         Width           =   2955
      End
   End
   Begin Threed.SSFrame fraYear 
      Height          =   4815
      Left            =   30
      TabIndex        =   24
      Top             =   1320
      Width           =   11745
      _Version        =   65536
      _ExtentX        =   20717
      _ExtentY        =   8493
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
      Begin FPSpread.vaSpread spdYear 
         Height          =   4575
         Left            =   90
         OleObjectBlob   =   "FGT0101.frx":1CD2
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   150
         Width           =   11535
      End
   End
   Begin Threed.SSFrame fraDay 
      Height          =   4815
      Left            =   30
      TabIndex        =   25
      Top             =   1320
      Width           =   11745
      _Version        =   65536
      _ExtentX        =   20717
      _ExtentY        =   8493
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
      Begin FPSpread.vaSpread spdDay 
         Height          =   4575
         Left            =   90
         OleObjectBlob   =   "FGT0101.frx":2181
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   150
         Width           =   11535
      End
   End
   Begin Threed.SSFrame fraMonth 
      Height          =   4815
      Left            =   30
      TabIndex        =   31
      Top             =   1320
      Width           =   11745
      _Version        =   65536
      _ExtentX        =   20717
      _ExtentY        =   8493
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
      Begin FPSpread.vaSpread spdMonth 
         Height          =   4575
         Left            =   90
         OleObjectBlob   =   "FGT0101.frx":2914
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   150
         Width           =   11535
      End
   End
End
Attribute VB_Name = "FGT0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TspdDay
    sDay    As String
    sSpdCol As String
End Type

Private Type tspdMonth
    sMonth  As String
    sSpdCol As String
End Type

Private Type tspdYear
    sYear   As String
    sSpdCol As String
End Type

Dim tspdDayi(31)    As TspdDay
Dim tspdMonthi(12)  As tspdMonth
Dim tspdYeari(10)   As tspdYear

Dim CodeHelp_F      As Integer
Dim bExist          As Integer
Dim DCJ0101         As DCJ0101
Dim DCT0101         As DCT0101
Dim msTotal         As String

Private Sub search_clear(sopt As String)

' 화면 청소
    If sopt = "D" Or sopt = "T" Then
        spdDay.Row = 1
        spdDay.Row2 = spdDay.MaxRows
        spdDay.BlockMode = True
        spdDay.Action = 3
        spdDay.BlockMode = False
        spdDay.MaxRows = 0
    End If
    
    If sopt = "M" Or sopt = "T" Then
        spdMonth.Row = 1
        spdMonth.Row2 = spdMonth.MaxRows
        spdMonth.BlockMode = True
        spdMonth.Action = 3
        spdMonth.BlockMode = False
        spdMonth.MaxRows = 0
    End If
    
    If sopt = "Y" Or sopt = "T" Then
        spdYear.Row = 1
        spdYear.Row2 = spdYear.MaxRows
        spdYear.BlockMode = True
        spdYear.Action = 3
        spdYear.BlockMode = False
        spdYear.MaxRows = 0
    End If
    
End Sub

Private Sub cmdExit_Click()
    
    Unload Me

End Sub

Private Sub cmdParth_Click()

    Dim i%
    Dim iDefaultSeq%
    
    CodeHelp_F = True
    
    Erase gCodeHlpTable '배열 초기화
    
    ReDim gCodeHlpTable(giPartCnt) As CodeTBL
    
    For i = 1 To giPartCnt
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = gPartTable(i).sPartInit
        gCodeHlpTable(i).sCodeNm = gPartTable(i).sPartName
    Next
        
    giCodeHlpCnt = giPartCnt
    
    hWndCd = txtPart.hwnd
    
    FST0101.Left = 1850
    FST0101.Top = 1350
    
    Load FST0101
    FST0101.Show vbModal

End Sub

Private Sub cmdPrint_Click()
       
    Dim iTotal      As Integer
    Dim sSearchUnit As String
    Dim sSearchGbn  As String
    Dim sSLabDate   As String
    Dim sELabDate   As String
    Dim sPartGbn    As String
    Dim sSum        As String
    Dim i           As Integer
    Dim vTmp

' 변수값 정리
    If optDate(0).Value = True Then
        sSearchUnit = "D"
    ElseIf optDate(1).Value = True Then
        sSearchUnit = "M"
    ElseIf optDate(2).Value = True Then
        sSearchUnit = "Y"
    End If

    sSLabDate = Format(dtpSLabDate, "YYYYMMDD")
    If txtSHour.Visible = True Then
        sSLabDate = sSLabDate & txtSHour
    End If
    sELabDate = Format(dtpELabDate, "YYYYMMDD")
    If txtEHour.Visible = True Then
        sELabDate = sELabDate & txtEHour
    End If
    If lblSlipNm <> "" Then
        sPartGbn = Mid(txtSlipCd, 2, 2)
    Else
        sPartGbn = ""
    End If

    If OptTGbn(0).Value = True Then
        sSearchGbn = "N"
    ElseIf OptTGbn(1).Value = True Then
        sSearchGbn = "A"
    End If
    
    If cboCondition.ListIndex = 0 Then
        sSearchGbn = sSearchGbn & Chr(124) & "T" & Chr(124)
    ElseIf cboCondition.ListIndex = 1 Then
        sSearchGbn = sSearchGbn & Chr(124) & "O" & Chr(124)
    ElseIf cboCondition.ListIndex = 2 Then
        sSearchGbn = sSearchGbn & Chr(124) & "I" & Chr(124)
    ElseIf cboCondition.ListIndex = 3 Then
        sSearchGbn = sSearchGbn & Chr(124) & "S" & Chr(124)
    End If
    
    If optDate(0).Value = True Then
        Call spdDay.GetText(1, 1, vTmp)
        
        If vTmp = "" Then
            Call cmdSearch_Click
        End If
        
        If spdDay.MaxRows = 0 Then
            ViewMsg "0건이 출력되었습니다."
            Exit Sub
        End If
        
        sSum = ""
        
        For i = 0 To 31
            If spdDay.ColWidth(i + 3) = 8 Or i = 31 Then
                Call spdDay.GetText(i + 3, spdDay.MaxRows, vTmp)
                
                If vTmp = "" Then
                    vTmp = "0"
                End If
                
                sSum = sSum & CStr(CInt(vTmp)) & "|"
            End If
        Next
    ElseIf optDate(1).Value = True Then
        Call spdMonth.GetText(1, 1, vTmp)
        
        If vTmp = "" Then
            Call cmdSearch_Click
        End If
        
        If spdMonth.MaxRows = 0 Then
            ViewMsg "0건이 출력되었습니다."
            Exit Sub
        End If
        
        sSum = ""
        
        For i = 0 To 12
            If spdMonth.ColWidth(i + 3) = 8 Or i = 12 Then
                Call spdMonth.GetText(i + 3, spdMonth.MaxRows, vTmp)
                If vTmp = "" Then
                    vTmp = "0"
                End If
                
                sSum = sSum & CStr(CInt(vTmp)) & "|"
            End If
        Next
    ElseIf optDate(2).Value = True Then
        Call spdYear.GetText(1, 1, vTmp)
        
        If vTmp = "" Then
            Call cmdSearch_Click
        End If
        
        If spdYear.MaxRows = 0 Then
            ViewMsg "0건이 출력되었습니다."
            Exit Sub
        End If
        
        sSum = ""
        
        For i = 0 To 10
            If spdYear.ColWidth(i + 3) = 8 Or i = 10 Then
                Call spdYear.GetText(i + 3, spdYear.MaxRows, vTmp)
                If vTmp = "" Then
                    vTmp = "0"
                End If
                
                sSum = sSum & CStr(CInt(vTmp)) & "|"
            End If
        Next
    End If
    
    Me.MousePointer = 11

    Set DCT0101 = New DCT0101
    iTotal = DCT0101.Print_ResultCnt(txtPart, sSearchUnit, sSLabDate, sELabDate, sPartGbn, sSearchGbn, msTotal, sSum)
    Set DCT0101 = Nothing

    ViewMsg Trim(Str(iTotal)) + " 건이 출력되었습니다."

    Me.MousePointer = 0
End Sub

Private Sub cmdSearch_Click()
    
    Dim bRtnCd      As Boolean
    
    Dim iCnt        As Integer
    Dim RCnt        As Integer
    Dim ispdRealCnt As Integer
    Dim iMaxCnt     As Integer
    Dim iSCnt       As Integer
    Dim iNCnt       As Integer
    Dim spdNRow     As Integer
    
    Dim lRCnt       As Long
        
    Dim sSearchUnit As String
    Dim sSearchGbn  As String
    Dim sSLabDate   As String
    Dim sELabDate   As String
    Dim sPartGbn    As String
    Dim sTotal      As String
    Dim sLabDate    As String
    Dim sSpcCd      As String
    Dim sOrdCd      As String
    Dim sSubMCd     As String
    Dim sPrintNm    As String
    Dim sOldPrintNm As String
    Dim sRCount     As String
    
    Dim VCount      As Variant
    Dim vTestNm     As Variant
    
' 필수항목 입력체크
    If Trim(lblPartNm.Caption) = "" Then
        ViewMsg "PART코드를 입력하여 주십시요"
        txtPart.SetFocus
        Exit Sub
    End If
    
    If Trim(lblSlipNm.Caption) = "" And Trim(txtSlipCd.Text) <> "" Then
        ViewMsg "SLIP정보를 정확히 처리해 주십시요"
        txtSlipCd.SetFocus
        Exit Sub
    End If
    
    If optDate(0).Value = False And optDate(1).Value = False And optDate(2).Value = False Then
        ViewMsg "검색구간이 선택되지 않았습니다."
        optDate(0).SetFocus
        Exit Sub
    End If
    
    If OptTGbn(0).Value = False And OptTGbn(1).Value = False Then
        ViewMsg "검색대상이 선택되지 않았습니다."
        OptTGbn(0).SetFocus
        Exit Sub
    End If

' 접수일자 구간 검사
    If optDate(0).Value = True Then
        If dtpELabDate.Value - dtpSLabDate > 31 Then
            ViewMsg "일별 조회시 일자구간은 31일을 넘지 못합니다."
            dtpSLabDate.SetFocus
            Exit Sub
        End If
    ElseIf optDate(1).Value = True Then
        If dtpELabDate.Value - dtpSLabDate.Value > 365 Then
            ViewMsg "월별 조회시 구간은 12개월을 넘지 못합니다."
            dtpSLabDate.SetFocus
            Exit Sub
        End If
    ElseIf optDate(2).Value = True Then
        If dtpELabDate.Value - dtpSLabDate.Value > 3650 Then
            dtpSLabDate.SetFocus
            ViewMsg "년별 조회시 구간은 10년을 넘지 못합니다."
        Else
            ViewMsg "시스템이 Down되는 수도 있으니, 가능한 년도 구간을 작게 주십시요"
        End If
    End If
    
' 변수값 정리
    If optDate(0).Value = True Then
        sSearchUnit = "D"
    ElseIf optDate(1).Value = True Then
        sSearchUnit = "M"
    ElseIf optDate(2).Value = True Then
        sSearchUnit = "Y"
    End If
    
    sSLabDate = Format(dtpSLabDate, "YYYYMMDD")
    If txtSHour.Visible = True Then
        sSLabDate = sSLabDate & txtSHour
    End If
    sELabDate = Format(dtpELabDate, "YYYYMMDD")
    If txtEHour.Visible = True Then
        sELabDate = sELabDate & txtEHour
    End If
    If lblSlipNm <> "" Then
        sPartGbn = Mid(txtSlipCd, 2, 2)
    Else
        sPartGbn = ""
    End If

    If OptTGbn(0).Value = True Then
        sSearchGbn = "N"
    ElseIf OptTGbn(1).Value = True Then
        sSearchGbn = "A"
    End If
    
    If cboCondition.ListIndex = 0 Then
        sSearchGbn = sSearchGbn & Chr(124) & "T" & Chr(124)
    ElseIf cboCondition.ListIndex = 1 Then
        sSearchGbn = sSearchGbn & Chr(124) & "O" & Chr(124)
    ElseIf cboCondition.ListIndex = 2 Then
        sSearchGbn = sSearchGbn & Chr(124) & "I" & Chr(124)
    ElseIf cboCondition.ListIndex = 3 Then
        sSearchGbn = sSearchGbn & Chr(124) & "S" & Chr(124)
    End If
    
' 실제 통계 퀴리(검사파트, 검색단위, 접수시작시작, 접수시각종료, 파트구분)
    Me.MousePointer = 11
    
    Set DCT0101 = New DCT0101
    msTotal = DCT0101.Get_ResultCnt(txtPart, sSearchUnit, sSLabDate, sELabDate, sPartGbn, sSearchGbn)
    sTotal = msTotal
    Set DCT0101 = Nothing

' 화면 표시
    ispdRealCnt = 0
    Select Case sSearchUnit
    Case "D"
        Call search_clear(sSearchUnit)
        ' 스프레드 헤드와 Field갯수 조정
        For iCnt = 0 To 30
            If dtpELabDate.Value < (dtpSLabDate.Value + iCnt) Then
                spdDay.ColWidth(iCnt + 3) = 0
            Else
                spdDay.ColWidth(iCnt + 3) = 8
                Call spdDay.SetText(iCnt + 3, 0, Format(dtpSLabDate.Value + iCnt, "MM/DD") & "일")
                ispdRealCnt = ispdRealCnt + 1
                tspdDayi(ispdRealCnt).sDay = Format(dtpSLabDate.Value + iCnt, "YYYYMMDD")
                tspdDayi(ispdRealCnt).sSpdCol = (iCnt + 3)
            End If
        Next iCnt
        iMaxCnt = ispdRealCnt
        
        sOldPrintNm = ""
        spdNRow = 0
        iNCnt = 1
        
        ' 실제 화면에 표시
        RCnt = Val(GetByOne(sTotal, sTotal))
        
        If RCnt = 0 Then
            Me.MousePointer = 0
            Exit Sub
        End If
        
        For iCnt = 1 To RCnt
            sLabDate = GetByOne(sTotal, sTotal)
            sPartGbn = GetByOne(sTotal, sTotal)
            sSpcCd = GetByOne(sTotal, sTotal)
            sOrdCd = GetByOne(sTotal, sTotal)
            sSubMCd = GetByOne(sTotal, sTotal)
            sPrintNm = GetByOne(sTotal, sTotal)
            sRCount = GetByOne(sTotal, sTotal)
            
            If sOldPrintNm <> sPrintNm Then
                sOldPrintNm = sPrintNm
                spdNRow = spdNRow + 1
                If spdDay.MaxRows < spdNRow Then
                    spdDay.MaxRows = spdNRow
                End If
                iNCnt = 1
            End If
            
            Call spdDay.SetText(1, spdNRow, CVar(txtPart & sPartGbn & sSpcCd & sOrdCd) & "")
            Call spdDay.SetText(2, spdNRow, sPrintNm)
            
            For iSCnt = iNCnt To iMaxCnt
                If sLabDate = tspdDayi(iSCnt).sDay Then
                    iNCnt = iSCnt
                    Call spdDay.SetText(Val(tspdDayi(iSCnt).sSpdCol), spdNRow, sRCount & "")
                End If
            Next iSCnt
        Next iCnt
        
        ' 소계 계산
        For iCnt = 1 To spdDay.MaxRows
            lRCnt = 0
            For RCnt = 0 To 30
                If spdDay.ColWidth(RCnt + 3) = 8 Then
                    bRtnCd = spdDay.GetText(RCnt + 3, iCnt, VCount)
                    lRCnt = lRCnt + Val(VCount)
                End If
            Next RCnt
            
            Call spdDay.SetText(34, iCnt, CVar(CStr(Val(lRCnt))) & "")
        Next iCnt
        
        ' 합계 계산
        If spdDay.MaxRows <> 0 Then
            spdDay.MaxRows = spdDay.MaxRows + 1
            Call spdDay.SetText(2, spdDay.MaxRows, ">>>   합         계   <<<")
            spdDay.Col = 1
            spdDay.Col2 = spdDay.MaxCols
            spdDay.Row = spdDay.MaxRows
            spdDay.Row2 = spdDay.MaxRows
            spdDay.BlockMode = True
            spdDay.BackColor = &HC0E0FF
            spdDay.BlockMode = False
        End If
        
        For iCnt = 0 To 31
            lRCnt = 0
            '소계일 경우 포함
            If spdDay.ColWidth(iCnt + 3) = 8 Or iCnt = 31 Then
                For RCnt = 1 To spdDay.MaxRows - 1
                    Call spdDay.GetText(iCnt + 3, RCnt, VCount)
                    lRCnt = lRCnt + Val(VCount)
                Next

                Call spdDay.SetText(iCnt + 3, spdDay.MaxRows, CVar(CStr(Val(lRCnt))) & "")
            End If
        Next
        
        ' 자릿수 점 찍기
        For iCnt = 1 To spdDay.MaxRows
        lRCnt = 0
            For RCnt = 0 To 31
                If spdDay.ColWidth(RCnt + 3) >= 8 Then
                    bRtnCd = spdDay.GetText(RCnt + 3, iCnt, VCount)
                    Call spdDay.SetText(RCnt + 3, iCnt, Format$(Val(VCount), "###,###,###,###") & "")
                End If
            Next RCnt
        Next iCnt
        
    Case "M"
        Call search_clear(sSearchUnit)
        For iCnt = 0 To 11
            If Val(Format(dtpELabDate.Value, "YYYYMM")) < Val(Format(dtpSLabDate.Value, "YYYYMM")) + iCnt Then
                spdMonth.ColWidth(iCnt + 3) = 0
            Else
                spdMonth.ColWidth(iCnt + 3) = 8
                Call spdMonth.SetText(iCnt + 3, 0, Format(dtpSLabDate.Value + (iCnt * 31), "YYYY/MM") & "월")
                ispdRealCnt = ispdRealCnt + 1
                tspdMonthi(ispdRealCnt).sMonth = Format(dtpSLabDate.Value + (iCnt * 31), "YYYYMM")
                tspdMonthi(ispdRealCnt).sSpdCol = (iCnt + 3)
            End If
        Next iCnt
        
        iMaxCnt = ispdRealCnt
        
        sOldPrintNm = ""
        spdNRow = 0
        iNCnt = 0
        
        ' 실제 화면에 표시
        RCnt = Val(GetByOne(sTotal, sTotal))
        
        If RCnt = 0 Then
            Me.MousePointer = 0
            Exit Sub
        End If
        
        For iCnt = 1 To RCnt
            sLabDate = Left(GetByOne(sTotal, sTotal), 6)
            sPartGbn = GetByOne(sTotal, sTotal)
            sSpcCd = GetByOne(sTotal, sTotal)
            sOrdCd = GetByOne(sTotal, sTotal)
            sSubMCd = GetByOne(sTotal, sTotal)
            sPrintNm = GetByOne(sTotal, sTotal)
            sRCount = GetByOne(sTotal, sTotal)
            If sOldPrintNm <> sPrintNm Then
                sOldPrintNm = sPrintNm
                spdNRow = spdNRow + 1
                If spdMonth.MaxRows < spdNRow Then
                    spdMonth.MaxRows = spdNRow
                End If
                iNCnt = 1
            End If
            
            Call spdMonth.SetText(1, spdNRow, CVar(txtPart & sPartGbn & sSpcCd & sOrdCd) & "")
            Call spdMonth.SetText(2, spdNRow, sPrintNm)
            
            For iSCnt = iNCnt To iMaxCnt
                If sLabDate = tspdMonthi(iSCnt).sMonth Then
                    iNCnt = iSCnt
                    Call spdMonth.SetText(Val(tspdMonthi(iSCnt).sSpdCol), spdNRow, sRCount & "")
                End If
            Next iSCnt
        Next iCnt
        
        ' 소계 계산
        For iCnt = 1 To spdMonth.MaxRows
            lRCnt = 0
            For RCnt = 0 To 11
                If spdMonth.ColWidth(RCnt + 3) = 8 Then
                    bRtnCd = spdMonth.GetText(RCnt + 3, iCnt, VCount)
                    lRCnt = lRCnt + Val(VCount)
                End If
            Next RCnt
            If lRCnt <> 0 Then
                Call spdMonth.SetText(15, iCnt, Trim(lRCnt) & "")
            End If
        Next iCnt
        
        ' 합계 계산
        If spdMonth.MaxRows <> 0 Then
            spdMonth.MaxRows = spdMonth.MaxRows + 1
            Call spdMonth.SetText(2, spdMonth.MaxRows, ">>>   합         계   <<<")
            spdMonth.Col = 1
            spdMonth.Col2 = spdMonth.MaxCols
            spdMonth.Row = spdMonth.MaxRows
            spdMonth.Row2 = spdMonth.MaxRows
            spdMonth.BlockMode = True
            spdMonth.BackColor = &HC0E0FF
            spdMonth.BlockMode = False
        End If
        
        For iCnt = 0 To 12
            lRCnt = 0
            '소계일 경우 포함
            If spdMonth.ColWidth(iCnt + 3) = 8 Or iCnt = 12 Then
                For RCnt = 1 To spdMonth.MaxRows - 1
                    Call spdMonth.GetText(iCnt + 3, RCnt, VCount)
                    lRCnt = lRCnt + Val(VCount)
                Next

                Call spdMonth.SetText(iCnt + 3, spdMonth.MaxRows, CVar(CStr(Val(lRCnt))) & "")
            End If
        Next
        
        ' 자릿수 점 찍기
        For iCnt = 1 To spdMonth.MaxRows
        lRCnt = 0
            For RCnt = 0 To 12
                If spdMonth.ColWidth(RCnt + 3) >= 8 Then
                    bRtnCd = spdMonth.GetText(RCnt + 3, iCnt, VCount)
                    Call spdMonth.SetText(RCnt + 3, iCnt, Format$(Val(VCount), "###,###,###,###") & "")
                End If
            Next RCnt
        Next iCnt
        
    Case "Y"
        Call search_clear(sSearchUnit)
        For iCnt = 0 To 9
            If dtpELabDate.Value < (dtpSLabDate.Value + (iCnt * 365)) Then
                spdYear.ColWidth(iCnt + 3) = 0
            Else
                spdYear.ColWidth(iCnt + 3) = 8
                Call spdYear.SetText(iCnt + 3, 0, Format(dtpSLabDate.Value + (iCnt * 365), "YYYY") & "년")
                ispdRealCnt = ispdRealCnt + 1
                tspdYeari(ispdRealCnt).sYear = Format(dtpSLabDate.Value + (iCnt * 365), "YYYY")
                tspdYeari(ispdRealCnt).sSpdCol = (iCnt + 3)
            End If
        Next iCnt
        
        iMaxCnt = ispdRealCnt
        
        sOldPrintNm = ""
        spdNRow = 0
        iNCnt = 0
        
        ' 실제 화면에 표시
        RCnt = Val(GetByOne(sTotal, sTotal))
        
        If RCnt = 0 Then
            Me.MousePointer = 0
            Exit Sub
        End If
        
        For iCnt = 1 To RCnt
            sLabDate = Left(GetByOne(sTotal, sTotal), 4)
            sPartGbn = GetByOne(sTotal, sTotal)
            sSpcCd = GetByOne(sTotal, sTotal)
            sOrdCd = GetByOne(sTotal, sTotal)
            sSubMCd = GetByOne(sTotal, sTotal)
            sPrintNm = GetByOne(sTotal, sTotal)
            sRCount = GetByOne(sTotal, sTotal)
            If sOldPrintNm <> sPrintNm Then
                sOldPrintNm = sPrintNm
                spdNRow = spdNRow + 1
                If spdYear.MaxRows < spdNRow Then
                    spdYear.MaxRows = spdNRow
                End If
                iNCnt = 1
            End If
            
            Call spdYear.SetText(1, spdNRow, CVar(txtPart & sPartGbn & sSpcCd & sOrdCd) & "")
            Call spdYear.SetText(2, spdNRow, sPrintNm)
            
            For iSCnt = iNCnt To iMaxCnt
                If sLabDate = tspdYeari(iSCnt).sYear Then
                    iNCnt = iSCnt
                    Call spdYear.SetText(Val(tspdYeari(iSCnt).sSpdCol), spdNRow, sRCount & "")
                End If
            Next iSCnt
        Next iCnt
        
        ' 소계 계산
        For iCnt = 1 To spdYear.MaxRows
            lRCnt = 0
            For RCnt = 0 To 9
                If spdYear.ColWidth(RCnt + 3) = 8 Then
                    bRtnCd = spdYear.GetText(RCnt + 3, iCnt, VCount)
                    lRCnt = lRCnt + Val(VCount)
                End If
            Next RCnt
            If lRCnt <> 0 Then
                Call spdYear.SetText(34, iCnt, Trim(lRCnt) & "")
            End If
        Next iCnt
        
        ' 합계 계산
        If spdYear.MaxRows <> 0 Then
            spdYear.MaxRows = spdYear.MaxRows + 1
            Call spdYear.SetText(2, spdYear.MaxRows, ">>>   합         계   <<<")
            spdYear.Col = 1
            spdYear.Col2 = spdYear.MaxCols
            spdYear.Row = spdYear.MaxRows
            spdYear.Row2 = spdYear.MaxRows
            spdYear.BlockMode = True
            spdYear.BackColor = &HC0E0FF
            spdYear.BlockMode = False
        End If
        
        For iCnt = 0 To 10
            lRCnt = 0
            '소계일 경우 포함
            If spdYear.ColWidth(iCnt + 3) = 8 Or iCnt = 10 Then
                For RCnt = 1 To spdYear.MaxRows - 1
                    Call spdYear.GetText(iCnt + 3, RCnt, VCount)
                    lRCnt = lRCnt + Val(VCount)
                Next

                Call spdYear.SetText(iCnt + 3, spdYear.MaxRows, CVar(CStr(Val(lRCnt))) & "")
            End If
        Next
        
        ' 자릿수 점 찍기
        For iCnt = 1 To spdYear.MaxRows
        lRCnt = 0
            For RCnt = 0 To 10
                If spdYear.ColWidth(RCnt + 3) >= 8 Then
                    bRtnCd = spdYear.GetText(RCnt + 3, iCnt, VCount)
                    Call spdYear.SetText(RCnt + 3, iCnt, Format$(Val(VCount), "###,###,###,###") & "")
                End If
            Next RCnt
        Next iCnt
    End Select
    
    Me.MousePointer = 0
    
End Sub

Private Sub cmdSliph_Click()

    Dim i%
    Dim j%
    Dim CPart As DCB0101
    Dim sTot01$
    Dim sTot02$
    Dim sTot03$
    
    txtSlipCd.SetFocus
    
    Set CPart = New DCB0101
    
    CPart.Get_PART txtPart
    
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
    
    hWndCd = txtSlipCd.hwnd
    
    FST0101.Left = 7130
    FST0101.Top = 1350
    
' Code Help Flag
    CodeHelp_F = True
    
    Load FST0101
    FST0101.Show vbModal

End Sub

Private Sub dtpELabDate_Change()
    If optDate(0).Value = True Then
        Call search_clear("D")
    ElseIf optDate(1).Value = True Then
        Call search_clear("M")
    ElseIf optDate(2).Value = True Then
        Call search_clear("Y")
    End If
End Sub

Private Sub dtpELabDate_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        dtpELabDate.SetFocus
        KeyCode = 0
    End If
    
End Sub

Private Sub dtpELabDate_LostFocus()

    If dtpSLabDate.Value > dtpELabDate.Value Then
        ViewMsg "검색구간이 잘못되었습니다."
        dtpSLabDate.SetFocus
    End If
    
    If optDate(0).Value = True Then
        If dtpELabDate.Value - dtpSLabDate > 31 Then
            ViewMsg "일별 조회시 일자구간은 31일을 넘지 못합니다."
            dtpSLabDate.SetFocus
        End If
    ElseIf optDate(1).Value = True Then
        If dtpELabDate.Value - dtpSLabDate.Value > 365 Then
            ViewMsg "월별 조회시 구간은 12개월을 넘지 못합니다."
            dtpSLabDate.SetFocus
        End If
    ElseIf optDate(2).Value = True Then
        If dtpELabDate.Value - dtpSLabDate.Value > 3650 Then
            dtpSLabDate.SetFocus
            ViewMsg "년별 조회시 구간은 10년을 넘지 못합니다."
        Else
            ViewMsg "많은 시간을 기다려야 하며, 심하면 시스템이 Down되는 수도 있으니, 년도 구간을 작게 주십시요"
        End If
    End If
    
End Sub

Private Sub dtpSLabDate_Change()
    If optDate(0).Value = True Then
        Call search_clear("D")
    ElseIf optDate(1).Value = True Then
        Call search_clear("M")
    ElseIf optDate(2).Value = True Then
        Call search_clear("Y")
    End If
End Sub

Private Sub dtpSLabDate_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        dtpSLabDate.SetFocus
        KeyCode = 0
    End If
    
End Sub

Private Sub Form_Activate()

    If CodeHelp_F = False Then
        optDate(0).Value = True
        OptTGbn(0).Value = True
        txtPart.SetFocus
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyF3
        Call cmdSearch_Click
    Case vbKeyF5
        Call cmdPrint_Click
    Case vbKeyEscape
        Call cmdExit_Click
    End Select
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ViewMsg ""
    End If
    
End Sub

Private Sub Form_Load()
    
    Me.Left = 0
    Me.Top = 0
    Me.Width = 11920
    Me.Height = 7700
    
    dtpSLabDate.Value = Format(Now, "yyyy-mm-dd")
    dtpELabDate.Value = Format(Now, "yyyy-mm-dd")
    
    cboCondition.AddItem "전체"
    cboCondition.AddItem "외래"
    cboCondition.AddItem "입원"
    cboCondition.AddItem "수탁"
    cboCondition.ListIndex = 0
    
    CodeHelp_F = False
    lblPartNm.Caption = fCurUserPartNm
    txtPart.Text = fCurUserPartCd
    txtSlipCd.Text = fCurUserSlipCd
    lblSlipNm.Caption = fCurUserSlipNm
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call InitRegCurFrmTitle
    
End Sub

Private Sub optDate_Click(Index As Integer)
    
    Select Case Index
    Case 0
        pnlLabDate.Caption = "접수일 구간"
        dtpSLabDate.CustomFormat = "yyy-MM-dd"
        dtpELabDate.CustomFormat = "yyy-MM-dd"
        
        dtpSLabDate.Value = Format(Now, "yyyy-mm-dd")
        dtpELabDate.Value = Format(Now, "yyyy-mm-dd")
    
        pnlLabTime.Visible = True
        txtSHour.Visible = True
        txtEHour.Visible = True
        lbl시(0).Visible = True
        lbl시(1).Visible = True
        lblSE.Visible = True

        txtSHour = "00"
        txtEHour = "23"
        
        fraDay.Left = 30
        fraDay.Top = 1350
        fraDay.Visible = True
        fraMonth.Visible = False
        fraYear.Visible = False
    
        fraMain.ZOrder 0
        optDate(0).SetFocus
    Case 1
        pnlLabDate.Caption = "접수월 구간"
        dtpSLabDate.CustomFormat = "yyy-MM"
        dtpELabDate.CustomFormat = "yyy-MM"
        
        dtpSLabDate.Value = Format(Now, "yyyy-mm-dd")
        dtpELabDate.Value = Format(Now, "yyyy-mm-dd")
    
        pnlLabTime.Visible = False
        txtSHour.Visible = False
        txtEHour.Visible = False
        lbl시(0).Visible = False
        lbl시(1).Visible = False
        lblSE.Visible = False
    
        fraMonth.Left = 30
        fraMonth.Top = 1350
        fraMonth.Visible = True
        fraDay.Visible = False
        fraYear.Visible = False
    
        fraMain.ZOrder 0
        optDate(1).SetFocus
    Case 2
        pnlLabDate.Caption = "접수년 구간"
        dtpSLabDate.CustomFormat = "yyy"
        dtpELabDate.CustomFormat = "yyy"
        
        dtpSLabDate.Value = Format(Now, "yyyy-mm-dd")
        dtpELabDate.Value = Format(Now, "yyyy-mm-dd")
    
        pnlLabTime.Visible = False
        txtSHour.Visible = False
        txtEHour.Visible = False
        lbl시(0).Visible = False
        lbl시(1).Visible = False
        lblSE.Visible = False
    
        fraYear.Left = 30
        fraYear.Top = 1350
        fraYear.Visible = True
        fraDay.Visible = False
        fraMonth.Visible = False
    
        fraMain.ZOrder 0
        optDate(2).SetFocus
    End Select
    
End Sub

Private Sub txtEHour_GotFocus()

    Txt_Highlight txtEHour
    
End Sub

Private Sub txtPart_Change()

    Dim i       As Integer

    On Error GoTo ErrHandler

    If Len(txtPart.Text) = txtPart.MaxLength Then

        For i = 1 To giPartCnt
            lblPartNm.Caption = ""
            If gPartTable(i).sPartInit = txtPart Then
                lblPartNm.Caption = gPartTable(i).sPartName
                bExist = True
                Exit For
            End If
        Next
    
        If lblPartNm.Caption = "" Then
            ViewMsg "존재하지 않는 Part Code입니다."
            Txt_Highlight txtPart
            Exit Sub
        Else
        '    txtSlipCd = ""
        '    lblSlipNm = ""
        End If
        
        If CodeHelp_F = False Then
            txtSlipCd.SetFocus
        Else
            SendKeys "{ENTER}"
        End If
    Else
        lblPartNm.Caption = ""
    End If

ErrHandler:

End Sub

Private Sub txtPart_GotFocus()

    Call Txt_Highlight(txtPart)
    
End Sub

Private Sub txtPart_KeyPress(KeyAscii As Integer)

    CodeHelp_F = False
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtPart_Validate(Cancel As Boolean)

    Dim i%
    
    For i = 1 To giPartCnt
        If gPartTable(i).sPartInit = txtPart Then
            lblPartNm.Caption = gPartTable(i).sPartName
            Exit For
        End If
    Next

End Sub

Private Sub txtSHour_GotFocus()

    Txt_Highlight txtSHour
    
End Sub

Private Sub txtSlipCd_Change()

    If Len(txtSlipCd.Text) = txtSlipCd.MaxLength Then
        
        Set DCJ0101 = New DCJ0101
        
        lblSlipNm.Caption = DCJ0101.Get_SlipNm(txtSlipCd)
    
        Set DCJ0101 = Nothing
        
        If Trim(lblSlipNm.Caption) = "" Then
            ViewMsg "존재하지 않는 Slip Code입니다."
            Txt_Highlight txtSlipCd
            Exit Sub
        End If
    
'''        If CodeHelp_F = False Then
'''            optDate(0).SetFocus
'''        Else
'''            SendKeys "{ENTER}"
'''        End If
    Else
        lblSlipNm.Caption = ""
    End If

End Sub

Private Sub txtSlipCd_GotFocus()

    txtSlipCd.Tag = Trim(txtSlipCd.Text)
    Call Txt_Highlight(txtSlipCd)
    If CodeHelp_F = False Then
        hWndCd = 0
    End If
    
End Sub

Private Sub txtSlipCd_KeyPress(KeyAscii As Integer)

    CodeHelp_F = False
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtSlipCd_LostFocus()
    
    If Trim(txtSlipCd) <> "" Then
        txtPart = Left(txtSlipCd, 1)
    End If
    
    If txtSlipCd.Tag <> txtSlipCd.Text Then
        Call search_clear("T")
    End If

End Sub

VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmEQ공용_Set_Equip_Config_test 
   Caption         =   "ㅗ히셔ㅣ"
   ClientHeight    =   8145
   ClientLeft      =   7245
   ClientTop       =   5595
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   6495
   Begin VB.Frame Frame1 
      Caption         =   "[장비검사코드]"
      Height          =   6795
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6195
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   4095
         ScaleHeight     =   285
         ScaleWidth      =   1965
         TabIndex        =   6
         Top             =   3570
         Width           =   1995
         Begin VB.OptionButton optSERIALYN 
            Caption         =   "사용"
            Height          =   300
            Index           =   0
            Left            =   45
            TabIndex        =   8
            Top             =   0
            Width           =   675
         End
         Begin VB.OptionButton optSERIALYN 
            Caption         =   "미사용"
            Height          =   300
            Index           =   1
            Left            =   960
            TabIndex        =   7
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   2820
         Left            =   3015
         ScaleHeight     =   2790
         ScaleWidth      =   3045
         TabIndex        =   4
         Top             =   630
         Width           =   3075
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   1455
            Left            =   1110
            TabIndex        =   14
            Top             =   855
            Width           =   1830
            _Version        =   393216
            _ExtentX        =   3228
            _ExtentY        =   2566
            _StockProps     =   64
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
            MaxCols         =   1
            ScrollBars      =   2
            SpreadDesigner  =   "frmEQ공용_Set_Equip_Config_test.frx":0000
            VisibleRows     =   1
         End
         Begin VB.TextBox Text1 
            Height          =   330
            IMEMode         =   10  '한글 
            Left            =   1140
            TabIndex        =   12
            Text            =   "txtDEPTCODE"
            Top             =   450
            Width           =   1815
         End
         Begin VB.TextBox txtDEPTCODE 
            Height          =   330
            IMEMode         =   10  '한글 
            Left            =   1140
            TabIndex        =   10
            Text            =   "txtDEPTCODE"
            Top             =   90
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "검  사  명"
            Height          =   180
            Index           =   3
            Left            =   135
            TabIndex        =   13
            Top             =   555
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "장비 코드"
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   11
            Top             =   150
            Width           =   780
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   2580
         Index           =   0
         Left            =   3015
         ScaleHeight     =   2550
         ScaleWidth      =   3045
         TabIndex        =   2
         Top             =   3990
         Width           =   3075
      End
      Begin FPSpread.vaSpread sprEQCD 
         Height          =   6255
         Left            =   135
         TabIndex        =   3
         Top             =   315
         Width           =   2715
         _Version        =   393216
         _ExtentX        =   4789
         _ExtentY        =   11033
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
         MaxCols         =   10
         ScrollBars      =   2
         SpreadDesigner  =   "frmEQ공용_Set_Equip_Config_test.frx":1811
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비코드 상세정보"
         Height          =   180
         Index           =   1
         Left            =   3060
         TabIndex        =   9
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "참고치"
         Height          =   180
         Index           =   0
         Left            =   3060
         TabIndex        =   5
         Top             =   3675
         Width           =   540
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "ㅇㅇㅇㅇㅇㅇㅇ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '하향 대각선
      Height          =   495
      Index           =   1
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   6375
   End
End
Attribute VB_Name = "frmEQ공용_Set_Equip_Config_test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


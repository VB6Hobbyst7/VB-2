VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00F8E4D8&
   Caption         =   "OK SOFT"
   ClientHeight    =   12075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20820
   LinkTopic       =   "Form1"
   ScaleHeight     =   12075
   ScaleWidth      =   20820
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame frameSet 
      BackColor       =   &H00F8E4D8&
      Caption         =   " 시스템 설정 "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Left            =   4080
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   5805
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   1680
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   1470
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1680
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   1110
         Width           =   2295
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "OCS"
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
         Index           =   6
         Left            =   4110
         TabIndex        =   18
         Top             =   1530
         Width           =   435
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "OCS"
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
         Index           =   5
         Left            =   4110
         TabIndex        =   17
         Top             =   1200
         Width           =   435
      End
      Begin VB.Image Image4 
         Height          =   225
         Left            =   390
         Picture         =   "Form1.frx":0000
         Top             =   1500
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "프로토콜"
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
         Index           =   4
         Left            =   600
         TabIndex        =   16
         Top             =   1530
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "OCS"
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
         Index           =   3
         Left            =   600
         TabIndex        =   14
         Top             =   1170
         Width           =   435
      End
      Begin VB.Image Image1 
         Height          =   225
         Left            =   390
         Picture         =   "Form1.frx":03EA
         Top             =   1140
         Width           =   150
      End
   End
   Begin VB.Frame frame1 
      BackColor       =   &H00F8E4D8&
      Height          =   9645
      Left            =   75
      TabIndex        =   9
      Top             =   1710
      Width           =   20685
      Begin VB.CommandButton cmdSL 
         Appearance      =   0  '평면
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
         Height          =   405
         Left            =   90
         TabIndex        =   51
         Top             =   210
         Width           =   465
      End
      Begin FPSpread.vaSpread spdOrder 
         Height          =   9375
         Left            =   60
         TabIndex        =   11
         Top             =   180
         Width           =   10875
         _Version        =   393216
         _ExtentX        =   19182
         _ExtentY        =   16536
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   19
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14548991
         SpreadDesigner  =   "Form1.frx":07D4
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdResult 
         Height          =   9360
         Left            =   10950
         TabIndex        =   10
         Top             =   180
         Width           =   9660
         _Version        =   393216
         _ExtentX        =   17039
         _ExtentY        =   16510
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   10
         MaxRows         =   50
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   13697023
         SpreadDesigner  =   "Form1.frx":4BF3
         TextTip         =   2
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '위 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   20760
      TabIndex        =   8
      Top             =   1035
      Width           =   20820
      Begin VB.Image imgTool 
         Height          =   930
         Index           =   2
         Left            =   3870
         Picture         =   "Form1.frx":549C
         Top             =   -210
         Width           =   1725
      End
      Begin VB.Image imgTool 
         Height          =   930
         Index           =   1
         Left            =   2130
         Picture         =   "Form1.frx":5F00
         Top             =   -210
         Width           =   1725
      End
      Begin VB.Image imgTool 
         Height          =   930
         Index           =   0
         Left            =   390
         Picture         =   "Form1.frx":6964
         Top             =   -210
         Width           =   1725
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   20760
      TabIndex        =   0
      Top             =   0
      Width           =   20820
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   14100
         TabIndex        =   1
         Top             =   510
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
         Format          =   21364736
         CurrentDate     =   40457
      End
      Begin VB.Image ImgSet 
         Height          =   585
         Left            =   1110
         Top             =   120
         Width           =   435
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   12930
         Picture         =   "Form1.frx":73C8
         Top             =   540
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "9600/N/8/1 127.0.0.1 [5005]"
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
         Index           =   2
         Left            =   16980
         TabIndex        =   7
         Top             =   180
         Width           =   2730
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "전남대학교병원 HITACHI 7020[H36] 홍길동[12345]"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   450
         Width           =   10395
      End
      Begin VB.Image Image3 
         Height          =   1065
         Left            =   0
         Picture         =   "Form1.frx":77B2
         Top             =   0
         Width           =   12900
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
         Left            =   13140
         TabIndex        =   5
         Top             =   570
         Width           =   780
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   14610
         Picture         =   "Form1.frx":8EF5
         Top             =   150
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   15555
         Picture         =   "Form1.frx":947F
         Top             =   150
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   16470
         Picture         =   "Form1.frx":9A09
         Top             =   150
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "포트"
         Height          =   180
         Index           =   0
         Left            =   14070
         TabIndex        =   4
         Top             =   180
         Width           =   360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "송신"
         Height          =   195
         Left            =   15045
         TabIndex        =   3
         Top             =   180
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "수신"
         Height          =   195
         Left            =   15930
         TabIndex        =   2
         Top             =   180
         Width           =   420
      End
   End
   Begin VB.Frame frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   9645
      Left            =   60
      TabIndex        =   19
      Top             =   1710
      Width           =   20685
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   8895
         Left            =   13860
         TabIndex        =   21
         Top             =   450
         Width           =   5955
         Begin VB.OptionButton optView 
            BackColor       =   &H00FFFFFF&
            Caption         =   "화면표시"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   4440
            TabIndex        =   50
            Top             =   2580
            Width           =   1125
         End
         Begin VB.OptionButton optView 
            BackColor       =   &H00FFFFFF&
            Caption         =   "화면표시"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   4440
            TabIndex        =   49
            Top             =   2130
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.TextBox txtMuch 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   390
            Width           =   2115
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            TabIndex        =   35
            Top             =   1725
            Width           =   2115
         End
         Begin VB.TextBox txtDec 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            TabIndex        =   34
            Top             =   2160
            Width           =   2115
         End
         Begin VB.TextBox txtCode 
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            TabIndex        =   33
            Top             =   1275
            Width           =   2115
         End
         Begin VB.TextBox txtEquipCode 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            TabIndex        =   32
            Top             =   840
            Width           =   2115
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            TabIndex        =   31
            Top             =   2610
            Width           =   2115
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            TabIndex        =   30
            Top             =   3045
            Width           =   1155
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            TabIndex        =   29
            Top             =   3495
            Width           =   1155
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            TabIndex        =   28
            Top             =   3930
            Width           =   2115
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            TabIndex        =   27
            Top             =   4380
            Width           =   2115
         End
         Begin VB.CommandButton cmdDown 
            Caption         =   "▼"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3810
            TabIndex        =   26
            Top             =   3480
            Width           =   465
         End
         Begin VB.CommandButton cmdUp 
            Caption         =   "▲"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3330
            TabIndex        =   25
            Top             =   3480
            Width           =   465
         End
         Begin VB.CommandButton Command1 
            Caption         =   "▼"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3810
            TabIndex        =   24
            Top             =   3030
            Width           =   465
         End
         Begin VB.CommandButton Command2 
            Caption         =   "▲"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3330
            TabIndex        =   23
            Top             =   3030
            Width           =   465
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            TabIndex        =   22
            Top             =   4830
            Width           =   2115
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   0
            Left            =   840
            Picture         =   "Form1.frx":9F93
            Top             =   450
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "장비코드"
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
            Index           =   8
            Left            =   1110
            TabIndex        =   48
            Top             =   480
            Width           =   780
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   1
            Left            =   840
            Picture         =   "Form1.frx":A37D
            Top             =   900
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "장비명"
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
            Index           =   9
            Left            =   1110
            TabIndex        =   47
            Top             =   930
            Width           =   585
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   2
            Left            =   840
            Picture         =   "Form1.frx":A767
            Top             =   1335
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "장비채널"
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
            Index           =   10
            Left            =   1110
            TabIndex        =   46
            Top             =   1365
            Width           =   780
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   3
            Left            =   840
            Picture         =   "Form1.frx":AB51
            Top             =   1785
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사코드"
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
            Index           =   11
            Left            =   1110
            TabIndex        =   45
            Top             =   1815
            Width           =   780
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   4
            Left            =   840
            Picture         =   "Form1.frx":AF3B
            Top             =   2220
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사명"
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
            Index           =   12
            Left            =   1110
            TabIndex        =   44
            Top             =   2250
            Width           =   585
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   5
            Left            =   840
            Picture         =   "Form1.frx":B325
            Top             =   2670
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사약어"
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
            Index           =   13
            Left            =   1110
            TabIndex        =   43
            Top             =   2700
            Width           =   780
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   6
            Left            =   840
            Picture         =   "Form1.frx":B70F
            Top             =   3105
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "소수점"
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
            Index           =   14
            Left            =   1110
            TabIndex        =   42
            Top             =   3135
            Width           =   585
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   7
            Left            =   840
            Picture         =   "Form1.frx":BAF9
            Top             =   3555
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "순번"
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
            Index           =   15
            Left            =   1110
            TabIndex        =   41
            Top             =   3585
            Width           =   390
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   8
            Left            =   840
            Picture         =   "Form1.frx":BEE3
            Top             =   3990
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Low"
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
            Index           =   16
            Left            =   1110
            TabIndex        =   40
            Top             =   4020
            Width           =   405
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   9
            Left            =   840
            Picture         =   "Form1.frx":C2CD
            Top             =   4440
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "High"
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
            Index           =   17
            Left            =   1110
            TabIndex        =   39
            Top             =   4470
            Width           =   435
         End
         Begin VB.Image imgSave 
            Height          =   1260
            Left            =   3060
            Picture         =   "Form1.frx":C6B7
            Top             =   6150
            Width           =   1290
         End
         Begin VB.Image imgDelete 
            Height          =   1260
            Left            =   1710
            Picture         =   "Form1.frx":E400
            Top             =   6150
            Width           =   1290
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "ex)10.00"
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
            Index           =   7
            Left            =   4470
            TabIndex        =   38
            Top             =   3120
            Width           =   825
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "결과단위"
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
            Index           =   18
            Left            =   1110
            TabIndex        =   37
            Top             =   4920
            Width           =   780
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   10
            Left            =   840
            Picture         =   "Form1.frx":1021A
            Top             =   4890
            Width           =   150
         End
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   8775
         Left            =   270
         TabIndex        =   20
         Top             =   540
         Width           =   12825
         _Version        =   393216
         _ExtentX        =   22622
         _ExtentY        =   15478
         _StockProps     =   64
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   11
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "Form1.frx":10604
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSL_Click()
    If cmdSL.Caption = "▶" Then
        cmdSL.Caption = "◀"
        spdOrder.Width = Me.Width - 200
    Else
        cmdSL.Caption = "▶"
        spdOrder.Width = gScaleWidth
    End If
End Sub

Private Sub Form_Load()
    
    frame1.Visible = True
    frame1.ZOrder 0

End Sub

Private Sub Form_Resize()
    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub
    

    frame1.Width = Me.ScaleWidth - 150
    frame1.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 200
    
    spdOrder.Width = Me.ScaleWidth - spdResult.Width - 280
    spdResult.Left = spdOrder.Left + spdOrder.Width

    spdOrder.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500
    spdResult.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500




End Sub

Private Sub ImgSet_DblClick()

    If frameSet.Visible = True Then
        frameSet.Visible = False
    Else
        frameSet.Visible = True
    End If
    
End Sub

Private Sub imgTool_Click(Index As Integer)

    Select Case Index
        Case 0:
                frame1.Visible = True
                frame1.ZOrder 0
        Case 1:
        
        Case 2:
                frame3.Visible = True
                frame3.ZOrder 0
    End Select
End Sub

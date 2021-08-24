VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00F8E4D8&
   Caption         =   "OK SOFT"
   ClientHeight    =   12915
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15960
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   15315
   ScaleWidth      =   28560
   StartUpPosition =   1  '소유자 가운데
   WindowState     =   2  '최대화
   Begin VB.Frame frame1 
      BackColor       =   &H00F8E4D8&
      Height          =   9645
      Left            =   50
      TabIndex        =   4
      Top             =   1650
      Width           =   20685
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Check1"
         Height          =   315
         Left            =   600
         TabIndex        =   92
         Top             =   240
         Width           =   195
      End
      Begin FPSpread.vaSpread spdOrder 
         Height          =   9375
         Left            =   60
         TabIndex        =   149
         Top             =   180
         Width           =   20535
         _Version        =   393216
         _ExtentX        =   36221
         _ExtentY        =   16536
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   20
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
         MaxCols         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   2
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14548991
         SpreadDesigner  =   "frmMain.frx":0E42
         UserResize      =   2
         ScrollBarTrack  =   1
         ShowScrollTips  =   3
      End
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
         Height          =   375
         Left            =   90
         TabIndex        =   23
         Top             =   210
         Width           =   435
      End
      Begin FPSpread.vaSpread spdResult 
         Height          =   9360
         Left            =   17370
         TabIndex        =   5
         Top             =   180
         Visible         =   0   'False
         Width           =   3210
         _Version        =   393216
         _ExtentX        =   5662
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
         MaxCols         =   12
         MaxRows         =   50
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   13697023
         SpreadDesigner  =   "frmMain.frx":5438
         TextTip         =   2
      End
   End
   Begin VB.Frame FraHidden 
      Caption         =   "HIDDEN CONTROL"
      Height          =   7875
      Left            =   21480
      TabIndex        =   95
      Top             =   2070
      Visible         =   0   'False
      Width           =   6525
      Begin MSComDlg.CommonDialog CFXFile 
         Left            =   3060
         Top             =   330
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame frameSet 
         BackColor       =   &H00FFFFFF&
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
         Height          =   1935
         Left            =   300
         TabIndex        =   136
         Top             =   5340
         Visible         =   0   'False
         Width           =   5025
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   1680
            TabIndex        =   138
            Text            =   "Combo1"
            Top             =   510
            Width           =   2295
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Left            =   1680
            TabIndex        =   137
            Text            =   "Combo1"
            Top             =   1110
            Width           =   2295
         End
         Begin VB.Image Image1 
            Height          =   225
            Left            =   390
            Picture         =   "frmMain.frx":5CD7
            Top             =   540
            Width           =   150
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
            TabIndex        =   142
            Top             =   570
            Width           =   435
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
            TabIndex        =   141
            Top             =   1170
            Width           =   780
         End
         Begin VB.Image Image4 
            Height          =   225
            Left            =   390
            Picture         =   "frmMain.frx":60C1
            Top             =   1140
            Width           =   150
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
            TabIndex        =   140
            Top             =   600
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
            Index           =   6
            Left            =   4110
            TabIndex        =   139
            Top             =   1170
            Width           =   435
         End
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "시스템설정"
         Height          =   375
         Left            =   3660
         TabIndex        =   134
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1470
         TabIndex        =   118
         Top             =   1140
         Width           =   3045
         Begin VB.OptionButton optBarSeq 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Seq 사용"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1770
            TabIndex        =   120
            Top             =   90
            Width           =   1155
         End
         Begin VB.OptionButton optBarSeq 
            BackColor       =   &H00FFFFFF&
            Caption         =   "검체번호 사용"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   119
            Top             =   90
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1470
         TabIndex        =   113
         Top             =   2040
         Width           =   2565
         Begin VB.OptionButton optSaveResult 
            BackColor       =   &H00FFFFFF&
            Caption         =   "LIS결과"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   1260
            TabIndex        =   115
            Top             =   30
            Width           =   1095
         End
         Begin VB.OptionButton optSaveResult 
            BackColor       =   &H00FFFFFF&
            Caption         =   "장비결과"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   90
            TabIndex        =   114
            Top             =   30
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1470
         TabIndex        =   110
         Top             =   1620
         Width           =   1875
         Begin VB.OptionButton optTrans 
            BackColor       =   &H00FFFFFF&
            Caption         =   "자동"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   90
            TabIndex        =   112
            Top             =   30
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton optTrans 
            BackColor       =   &H00FFFFFF&
            Caption         =   "수동"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   930
            TabIndex        =   111
            Top             =   30
            Width           =   765
         End
      End
      Begin VB.Timer tmrReceive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2100
         Top             =   300
      End
      Begin VB.Timer tmrSend 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2580
         Top             =   300
      End
      Begin MSWinsockLib.Winsock wSck 
         Left            =   210
         Top             =   330
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   1380
         Top             =   210
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":64AB
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6A45
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6FDF
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7579
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7E0B
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7F65
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":80BF
               Key             =   "NOF"
            EndProperty
         EndProperty
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   660
         Top             =   300
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         RTSEnable       =   -1  'True
         EOFEnable       =   -1  'True
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1725
         Left            =   180
         TabIndex        =   133
         Top             =   2760
         Width           =   5085
         _Version        =   393216
         _ExtentX        =   8969
         _ExtentY        =   3043
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
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmMain.frx":8219
      End
      Begin FPSpread.vaSpread vasHPVReport 
         Height          =   4155
         Left            =   450
         TabIndex        =   160
         Top             =   3480
         Width           =   16725
         _Version        =   393216
         _ExtentX        =   29501
         _ExtentY        =   7329
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
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   20
         MaxRows         =   11
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ScrollBarShowMax=   0   'False
         SpreadDesigner  =   "frmMain.frx":B25E
         UserResize      =   0
      End
      Begin FPSpread.vaSpread vasPrint 
         Height          =   1905
         Left            =   510
         TabIndex        =   162
         Top             =   2550
         Visible         =   0   'False
         Width           =   4395
         _Version        =   393216
         _ExtentX        =   7752
         _ExtentY        =   3360
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
         GridSolid       =   0   'False
         MaxCols         =   18
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ScrollBarShowMax=   0   'False
         SpreadDesigner  =   "frmMain.frx":D7ED
         UserResize      =   0
      End
      Begin VB.Label Label3 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "바코드사용"
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
         Left            =   390
         TabIndex        =   121
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과적용"
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
         Left            =   390
         TabIndex        =   117
         Top             =   2130
         Width           =   780
      End
      Begin VB.Label Label2 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과전송"
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
         Left            =   390
         TabIndex        =   116
         Top             =   1710
         Width           =   780
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   285
      Left            =   0
      TabIndex        =   131
      Top             =   15030
      Width           =   28560
      _ExtentX        =   50377
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   28530
      TabIndex        =   3
      Top             =   1035
      Width           =   28560
      Begin VB.Frame fraInterface 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   6510
         TabIndex        =   89
         Top             =   -60
         Width           =   14145
         Begin VB.CommandButton cmdRPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "결과출력"
            Height          =   375
            Left            =   4200
            TabIndex        =   161
            Top             =   150
            Width           =   1305
         End
         Begin VB.Frame Frame9 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   11070
            TabIndex        =   154
            Top             =   60
            Width           =   3015
            Begin VB.OptionButton optCheck 
               BackColor       =   &H00FFFFFF&
               Caption         =   "전체"
               Height          =   195
               Index           =   0
               Left            =   210
               TabIndex        =   157
               Top             =   210
               Width           =   675
            End
            Begin VB.OptionButton optCheck 
               BackColor       =   &H00FFFFFF&
               Caption         =   "접수"
               Height          =   195
               Index           =   1
               Left            =   1020
               TabIndex        =   156
               Top             =   210
               Value           =   -1  'True
               Width           =   675
            End
            Begin VB.OptionButton optCheck 
               BackColor       =   &H00FFFFFF&
               Caption         =   "예비결과"
               Height          =   195
               Index           =   2
               Left            =   1890
               TabIndex        =   155
               Top             =   210
               Width           =   1065
            End
         End
         Begin VB.CommandButton cmdResult 
            BackColor       =   &H00FFFFFF&
            Caption         =   "결과받기"
            Height          =   375
            Left            =   2850
            TabIndex        =   151
            Top             =   150
            Width           =   1305
         End
         Begin VB.CommandButton cmdOrder 
            BackColor       =   &H00FFFFFF&
            Caption         =   "오더전송"
            Height          =   375
            Left            =   1500
            TabIndex        =   150
            Top             =   150
            Width           =   1305
         End
         Begin VB.Label Label4 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사구분"
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
            Left            =   10080
            TabIndex        =   158
            Top             =   270
            Width           =   780
         End
         Begin VB.Shape shpC 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   7590
            Top             =   150
            Width           =   1365
         End
         Begin VB.Label lblClear 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "화면정리"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7710
            TabIndex        =   94
            Top             =   240
            Width           =   1125
         End
         Begin VB.Shape shpS 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   6120
            Top             =   150
            Width           =   1365
         End
         Begin VB.Label lblSave 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "선택저장"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6270
            TabIndex        =   93
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label lblWork 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "워크조회"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   90
            Top             =   240
            Width           =   1125
         End
         Begin VB.Shape shpW 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   90
            Top             =   150
            Width           =   1365
         End
      End
      Begin VB.Frame fraResult 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   6510
         TabIndex        =   105
         Top             =   -60
         Visible         =   0   'False
         Width           =   14145
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "결과출력"
            Height          =   375
            Left            =   7650
            TabIndex        =   159
            Top             =   150
            Width           =   1305
         End
         Begin VB.ComboBox cboRstType 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "frmMain.frx":E118
            Left            =   420
            List            =   "frmMain.frx":E11A
            TabIndex        =   128
            Top             =   180
            Width           =   1245
         End
         Begin VB.ComboBox cboState 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "frmMain.frx":E11C
            Left            =   4710
            List            =   "frmMain.frx":E11E
            TabIndex        =   127
            Top             =   180
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   315
            Left            =   1770
            TabIndex        =   107
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   127860737
            CurrentDate     =   40457
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   315
            Left            =   3330
            TabIndex        =   108
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   127860737
            CurrentDate     =   40457
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "~"
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
            Index           =   26
            Left            =   3120
            TabIndex        =   109
            Top             =   240
            Width           =   150
         End
         Begin VB.Image imgGbn 
            Height          =   225
            Left            =   180
            Picture         =   "frmMain.frx":E120
            Top             =   210
            Width           =   150
         End
         Begin VB.Shape shpR 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   6180
            Top             =   150
            Width           =   1365
         End
         Begin VB.Label lblResult 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "결과조회"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6300
            TabIndex        =   106
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "통신설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   4830
         TabIndex        =   45
         Top             =   150
         Width           =   1125
      End
      Begin VB.Shape shpB 
         BackColor       =   &H00C0FFC0&
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00C0FFC0&
         Height          =   435
         Index           =   3
         Left            =   4710
         Top             =   60
         Width           =   1395
      End
      Begin VB.Shape shpB 
         BackColor       =   &H00C0FFC0&
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00C0FFC0&
         Height          =   435
         Index           =   2
         Left            =   3240
         Top             =   60
         Width           =   1395
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   27
         Top             =   150
         Width           =   1125
      End
      Begin VB.Shape shpB 
         BackColor       =   &H00C0FFC0&
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00C0FFC0&
         Height          =   435
         Index           =   1
         Left            =   1770
         Top             =   60
         Width           =   1395
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과조회"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1890
         TabIndex        =   26
         Top             =   150
         Width           =   1125
      End
      Begin VB.Shape shpB 
         BackColor       =   &H00C0FFC0&
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00C0FFC0&
         Height          =   435
         Index           =   0
         Left            =   270
         Top             =   60
         Width           =   1395
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "인터페이스"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   25
         Top             =   150
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   28560
      TabIndex        =   0
      Top             =   0
      Width           =   28560
      Begin VB.Frame fraCommTest 
         Height          =   945
         Left            =   15600
         TabIndex        =   124
         Top             =   30
         Visible         =   0   'False
         Width           =   4935
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send"
            Height          =   735
            Left            =   60
            TabIndex        =   126
            Top             =   150
            Width           =   375
         End
         Begin VB.TextBox txtRcv 
            Height          =   765
            Left            =   450
            MultiLine       =   -1  'True
            TabIndex        =   125
            Top             =   120
            Width           =   4425
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   12630
         TabIndex        =   101
         Top             =   60
         Width           =   2985
         Begin VB.Label lblReceive 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "수신"
            Height          =   195
            Left            =   2010
            TabIndex        =   104
            Top             =   210
            Width           =   420
         End
         Begin VB.Label lblSend 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "송신"
            Height          =   195
            Left            =   1125
            TabIndex        =   103
            Top             =   210
            Width           =   420
         End
         Begin VB.Label lblPort 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "포트"
            Height          =   180
            Left            =   150
            TabIndex        =   102
            Top             =   210
            Width           =   360
         End
         Begin VB.Image imgReceive 
            Height          =   240
            Left            =   2550
            Picture         =   "frmMain.frx":E50A
            Top             =   180
            Width           =   240
         End
         Begin VB.Image imgSend 
            Height          =   240
            Left            =   1635
            Picture         =   "frmMain.frx":EA94
            Top             =   180
            Width           =   240
         End
         Begin VB.Image imgPort 
            Height          =   240
            Left            =   690
            Picture         =   "frmMain.frx":F01E
            Top             =   180
            Width           =   240
         End
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   10020
         TabIndex        =   122
         Top             =   540
         Width           =   2505
         _ExtentX        =   4419
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
         Format          =   127860736
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   27
         Left            =   9150
         TabIndex        =   123
         Top             =   630
         Width           =   720
      End
      Begin VB.Image Image7 
         Height          =   225
         Left            =   8880
         Picture         =   "frmMain.frx":F5A8
         Top             =   600
         Width           =   150
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
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
         Left            =   12840
         TabIndex        =   2
         Top             =   660
         Width           =   75
      End
      Begin VB.Label lblHospInfo 
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
         TabIndex        =   1
         Top             =   450
         Width           =   10485
      End
      Begin VB.Image Image3 
         Height          =   1065
         Left            =   0
         Picture         =   "frmMain.frx":F992
         Top             =   0
         Width           =   12900
      End
   End
   Begin VB.Frame frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   9645
      Left            =   930
      TabIndex        =   48
      Top             =   2370
      Visible         =   0   'False
      Width           =   20685
      Begin VB.OptionButton optComType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   16680
         TabIndex        =   147
         Top             =   480
         Width           =   1125
      End
      Begin VB.Frame frameFILE 
         BackColor       =   &H00FFFFFF&
         Caption         =   " TCP-IP 설정 "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7935
         Left            =   12480
         TabIndex        =   143
         Top             =   900
         Width           =   5325
         Begin VB.TextBox txtRstPath 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   152
            Top             =   2370
            Width           =   4845
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "찾기"
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
            Left            =   4350
            TabIndex        =   148
            Top             =   720
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtOrdPath 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   144
            Top             =   1320
            Width           =   4845
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "결과(CSV) 경로"
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
            Height          =   195
            Index           =   0
            Left            =   510
            TabIndex        =   153
            Top             =   2040
            Width           =   1380
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   17
            Left            =   240
            Picture         =   "frmMain.frx":110D5
            Top             =   2010
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   25
            Left            =   240
            Picture         =   "frmMain.frx":114BF
            Top             =   960
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "오더 경로"
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
            Height          =   195
            Index           =   2
            Left            =   510
            TabIndex        =   146
            Top             =   990
            Width           =   840
         End
         Begin VB.Label lblFileSave 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "저장"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3300
            TabIndex        =   145
            Top             =   6960
            Width           =   1125
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   585
            Left            =   3180
            Top             =   6810
            Width           =   1365
         End
      End
      Begin VB.CommandButton cmdIF 
         Caption         =   "IF 설정"
         Height          =   375
         Left            =   11970
         TabIndex        =   135
         Top             =   8280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "병원정보설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   390
         TabIndex        =   132
         Top             =   270
         Width           =   1965
      End
      Begin VB.OptionButton optComType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   10710
         TabIndex        =   70
         Top             =   510
         Width           =   1125
      End
      Begin VB.OptionButton optComType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4620
         TabIndex        =   69
         Top             =   450
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.Frame frameTCP 
         BackColor       =   &H00FFFFFF&
         Caption         =   " TCP-IP 설정 "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7935
         Left            =   6480
         TabIndex        =   63
         Top             =   900
         Width           =   5325
         Begin VB.OptionButton optTCPType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Client"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   73
            Top             =   390
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton optTCPType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Server"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   3030
            TabIndex        =   72
            Top             =   390
            Width           =   1125
         End
         Begin VB.TextBox txtTCPPort 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            TabIndex        =   68
            Top             =   1320
            Width           =   2445
         End
         Begin VB.TextBox txtTCPIP 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            TabIndex        =   67
            Top             =   930
            Width           =   2445
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   7
            Left            =   840
            Picture         =   "frmMain.frx":118A9
            Top             =   450
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Type"
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
            Height          =   195
            Index           =   18
            Left            =   1110
            TabIndex        =   71
            Top             =   480
            Width           =   465
         End
         Begin VB.Shape shpTcp 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   585
            Left            =   3180
            Top             =   6810
            Width           =   1365
         End
         Begin VB.Label lblTcpSave 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "저장"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3300
            TabIndex        =   66
            Top             =   6960
            Width           =   1125
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Port"
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
            Height          =   195
            Index           =   25
            Left            =   1110
            TabIndex        =   65
            Top             =   1395
            Width           =   375
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   15
            Left            =   840
            Picture         =   "frmMain.frx":11C93
            Top             =   1365
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "IP"
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
            Height          =   195
            Index           =   24
            Left            =   1110
            TabIndex        =   64
            Top             =   990
            Width           =   180
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   10
            Left            =   840
            Picture         =   "frmMain.frx":1207D
            Top             =   960
            Width           =   150
         End
      End
      Begin VB.Frame frameCom 
         BackColor       =   &H00FFFFFF&
         Caption         =   " RS-232 설정 "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7935
         Left            =   420
         TabIndex        =   49
         Top             =   870
         Width           =   5325
         Begin VB.ComboBox cboPort 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMain.frx":12467
            Left            =   2190
            List            =   "frmMain.frx":12469
            TabIndex        =   62
            Top             =   390
            Width           =   2205
         End
         Begin VB.ComboBox cboBaudrate 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMain.frx":1246B
            Left            =   2190
            List            =   "frmMain.frx":1246D
            TabIndex        =   61
            Top             =   780
            Width           =   2205
         End
         Begin VB.ComboBox cboDatabit 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMain.frx":1246F
            Left            =   2190
            List            =   "frmMain.frx":12471
            TabIndex        =   60
            Top             =   1170
            Width           =   2205
         End
         Begin VB.ComboBox cboStartbit 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2190
            TabIndex        =   59
            Top             =   1590
            Width           =   2205
         End
         Begin VB.ComboBox cboStopbit 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2190
            TabIndex        =   58
            Top             =   2070
            Width           =   2205
         End
         Begin VB.ComboBox cboParity 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMain.frx":12473
            Left            =   2190
            List            =   "frmMain.frx":12475
            TabIndex        =   57
            Top             =   2520
            Width           =   2205
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "DataBit"
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
            Height          =   195
            Index           =   33
            Left            =   1110
            TabIndex        =   56
            Top             =   1290
            Width           =   645
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   23
            Left            =   840
            Picture         =   "frmMain.frx":12477
            Top             =   1260
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   22
            Left            =   840
            Picture         =   "frmMain.frx":12861
            Top             =   450
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "통신포트"
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
            Height          =   195
            Index           =   32
            Left            =   1110
            TabIndex        =   55
            Top             =   480
            Width           =   780
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   21
            Left            =   840
            Picture         =   "frmMain.frx":12C4B
            Top             =   855
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Baudrate"
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
            Height          =   195
            Index           =   31
            Left            =   1110
            TabIndex        =   54
            Top             =   885
            Width           =   855
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   20
            Left            =   840
            Picture         =   "frmMain.frx":13035
            Top             =   1695
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Start Bit"
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
            Height          =   195
            Index           =   30
            Left            =   1110
            TabIndex        =   53
            Top             =   1725
            Width           =   705
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   19
            Left            =   840
            Picture         =   "frmMain.frx":1341F
            Top             =   2100
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Stop Bit"
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
            Height          =   195
            Index           =   29
            Left            =   1110
            TabIndex        =   52
            Top             =   2130
            Width           =   705
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   18
            Left            =   840
            Picture         =   "frmMain.frx":13809
            Top             =   2550
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Parity"
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
            Height          =   195
            Index           =   28
            Left            =   1110
            TabIndex        =   51
            Top             =   2580
            Width           =   525
         End
         Begin VB.Label lblComSave 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "저장"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3300
            TabIndex        =   50
            Top             =   6960
            Width           =   1125
         End
         Begin VB.Shape shpCom 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   585
            Left            =   3180
            Top             =   6810
            Width           =   1365
         End
      End
   End
   Begin VB.Frame frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   9645
      Left            =   1230
      TabIndex        =   6
      Top             =   1950
      Visible         =   0   'False
      Width           =   20685
      Begin VB.Frame frameTestSet 
         BackColor       =   &H00FFFFFF&
         Height          =   9315
         Left            =   14730
         TabIndex        =   8
         Top             =   180
         Width           =   5625
         Begin VB.Frame frameOrder 
            BackColor       =   &H00FFFFFF&
            Height          =   2235
            Left            =   210
            TabIndex        =   86
            Top             =   6960
            Visible         =   0   'False
            Width           =   2085
            Begin VB.CommandButton cmdDelete 
               Appearance      =   0  '평면
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   420
               TabIndex        =   91
               Top             =   210
               Width           =   285
            End
            Begin VB.CommandButton cmdAppend 
               Appearance      =   0  '평면
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   87
               Top             =   210
               Width           =   285
            End
            Begin FPSpread.vaSpread spdOrdMst 
               Height          =   1920
               Left            =   90
               TabIndex        =   88
               Top             =   180
               Width           =   1890
               _Version        =   393216
               _ExtentX        =   3334
               _ExtentY        =   3387
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
               MaxCols         =   1
               MaxRows         =   50
               OperationMode   =   2
               RetainSelBlock  =   0   'False
               ScrollBars      =   2
               SelectBlockOptions=   0
               ShadowColor     =   13697023
               SpreadDesigner  =   "frmMain.frx":13BF3
               TextTip         =   2
            End
         End
         Begin VB.ComboBox cboResultType 
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
            ItemData        =   "frmMain.frx":13F7F
            Left            =   1650
            List            =   "frmMain.frx":13F81
            TabIndex        =   42
            Top             =   4470
            Width           =   1575
         End
         Begin VB.Frame frameCut 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   555
            Left            =   1440
            TabIndex        =   32
            Top             =   4740
            Width           =   2565
            Begin VB.OptionButton optCutUse 
               BackColor       =   &H00FFFFFF&
               Caption         =   "사용"
               Height          =   315
               Index           =   1
               Left            =   1320
               TabIndex        =   34
               Top             =   180
               Width           =   1125
            End
            Begin VB.OptionButton optCutUse 
               BackColor       =   &H00FFFFFF&
               Caption         =   "미사용"
               Height          =   315
               Index           =   0
               Left            =   210
               TabIndex        =   33
               Top             =   180
               Value           =   -1  'True
               Width           =   1125
            End
         End
         Begin VB.Frame frameCutOff 
            BackColor       =   &H00FFFFFF&
            Height          =   1545
            Left            =   210
            TabIndex        =   28
            Top             =   5340
            Width           =   5175
            Begin VB.TextBox txtCOHOut 
               Alignment       =   2  '가운데 맞춤
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
               Height          =   300
               Left            =   3480
               TabIndex        =   41
               Top             =   1020
               Width           =   1545
            End
            Begin VB.TextBox txtCOHIn 
               Alignment       =   2  '가운데 맞춤
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
               Height          =   300
               Left            =   1530
               TabIndex        =   39
               Top             =   1020
               Width           =   1185
            End
            Begin VB.ComboBox cboCOH 
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
               ItemData        =   "frmMain.frx":13F83
               Left            =   2730
               List            =   "frmMain.frx":13F85
               TabIndex        =   38
               Top             =   1020
               Width           =   735
            End
            Begin VB.TextBox txtCOMOut 
               Alignment       =   2  '가운데 맞춤
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
               Height          =   300
               Left            =   3480
               TabIndex        =   37
               Top             =   660
               Width           =   1545
            End
            Begin VB.TextBox txtCOLOut 
               Alignment       =   2  '가운데 맞춤
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
               Height          =   300
               Left            =   3480
               TabIndex        =   35
               Top             =   300
               Width           =   1545
            End
            Begin VB.TextBox txtCOLIn 
               Alignment       =   2  '가운데 맞춤
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
               Height          =   300
               Left            =   1530
               TabIndex        =   30
               Top             =   300
               Width           =   1185
            End
            Begin VB.ComboBox cboCOL 
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
               ItemData        =   "frmMain.frx":13F87
               Left            =   2730
               List            =   "frmMain.frx":13F89
               TabIndex        =   29
               Top             =   300
               Width           =   735
            End
            Begin VB.Image Image5 
               Height          =   225
               Index           =   13
               Left            =   210
               Picture         =   "frmMain.frx":13F8B
               Top             =   1080
               Width           =   150
            End
            Begin VB.Image Image5 
               Height          =   225
               Index           =   9
               Left            =   210
               Picture         =   "frmMain.frx":14375
               Top             =   720
               Width           =   150
            End
            Begin VB.Label Label1 
               Appearance      =   0  '평면
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  '투명
               Caption         =   "CutOff (H)"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   21
               Left            =   480
               TabIndex        =   40
               Top             =   1110
               Width           =   840
            End
            Begin VB.Label Label1 
               Appearance      =   0  '평면
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  '투명
               Caption         =   "CutOff (M)"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   17
               Left            =   480
               TabIndex        =   36
               Top             =   750
               Width           =   885
            End
            Begin VB.Label Label1 
               Appearance      =   0  '평면
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  '투명
               Caption         =   "CutOff (L)"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   20
               Left            =   480
               TabIndex        =   31
               Top             =   390
               Width           =   825
            End
            Begin VB.Image Image5 
               Height          =   225
               Index           =   12
               Left            =   210
               Picture         =   "frmMain.frx":1475F
               Top             =   360
               Width           =   150
            End
         End
         Begin VB.TextBox txtRChannel 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            TabIndex        =   24
            Top             =   1770
            Width           =   2115
         End
         Begin VB.TextBox txtEqpCD 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   420
            Width           =   1215
         End
         Begin VB.TextBox txtTestCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            TabIndex        =   20
            Top             =   2220
            Width           =   2115
         End
         Begin VB.TextBox txtTestNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            TabIndex        =   19
            Top             =   2670
            Width           =   2115
         End
         Begin VB.TextBox txtOChannel 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            TabIndex        =   18
            Top             =   1320
            Width           =   2115
         End
         Begin VB.TextBox txtAbbrNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            TabIndex        =   17
            Top             =   3120
            Width           =   2115
         End
         Begin VB.TextBox txtResSpec 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            TabIndex        =   16
            Top             =   3570
            Width           =   1215
         End
         Begin VB.TextBox txtSeq 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            TabIndex        =   15
            Top             =   870
            Width           =   1245
         End
         Begin VB.TextBox txtRefLow 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            TabIndex        =   14
            Top             =   4020
            Width           =   1545
         End
         Begin VB.TextBox txtRefHigh 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3330
            TabIndex        =   13
            Top             =   4020
            Width           =   1545
         End
         Begin VB.CommandButton cmdSeqDown 
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
            Height          =   330
            Left            =   3330
            TabIndex        =   12
            Top             =   840
            Width           =   405
         End
         Begin VB.CommandButton cmdSeqUp 
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
            Height          =   330
            Left            =   2910
            TabIndex        =   11
            Top             =   840
            Width           =   405
         End
         Begin VB.CommandButton cmdSpecDown 
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
            Height          =   330
            Left            =   3330
            TabIndex        =   10
            Top             =   3540
            Width           =   435
         End
         Begin VB.CommandButton cmdSpecUP 
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
            Height          =   330
            Left            =   2880
            TabIndex        =   9
            Top             =   3540
            Width           =   435
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "순번"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   15
            Left            =   600
            TabIndex        =   85
            Top             =   933
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "결과형식"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   22
            Left            =   600
            TabIndex        =   84
            Top             =   4557
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   14
            Left            =   330
            Picture         =   "frmMain.frx":14B49
            Top             =   4527
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "CutOff"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   9
            Left            =   600
            TabIndex        =   83
            Top             =   5010
            Width           =   510
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   1
            Left            =   330
            Picture         =   "frmMain.frx":14F33
            Top             =   4980
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "결과채널"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   19
            Left            =   600
            TabIndex        =   82
            Top             =   1839
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   11
            Left            =   330
            Picture         =   "frmMain.frx":1531D
            Top             =   1809
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   0
            Left            =   330
            Picture         =   "frmMain.frx":15707
            Top             =   450
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "장비코드"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   8
            Left            =   600
            TabIndex        =   81
            Top             =   480
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   2
            Left            =   330
            Picture         =   "frmMain.frx":15AF1
            Top             =   1356
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "오더채널"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   10
            Left            =   600
            TabIndex        =   80
            Top             =   1386
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   3
            Left            =   330
            Picture         =   "frmMain.frx":15EDB
            Top             =   2262
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사코드"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   11
            Left            =   600
            TabIndex        =   79
            Top             =   2292
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   4
            Left            =   330
            Picture         =   "frmMain.frx":162C5
            Top             =   2715
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사명"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   12
            Left            =   600
            TabIndex        =   78
            Top             =   2745
            Width           =   540
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   5
            Left            =   330
            Picture         =   "frmMain.frx":166AF
            Top             =   3168
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사약어"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   13
            Left            =   600
            TabIndex        =   77
            Top             =   3198
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   6
            Left            =   330
            Picture         =   "frmMain.frx":16A99
            Top             =   3621
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "소수점"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   14
            Left            =   600
            TabIndex        =   76
            Top             =   3651
            Width           =   540
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   8
            Left            =   330
            Picture         =   "frmMain.frx":16E83
            Top             =   4074
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "참고치"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   16
            Left            =   600
            TabIndex        =   75
            Top             =   4104
            Width           =   540
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   16
            Left            =   330
            Picture         =   "frmMain.frx":1726D
            Top             =   903
            Width           =   150
         End
         Begin VB.Shape shpA 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   405
            Index           =   3
            Left            =   3990
            Top             =   8550
            Width           =   1335
         End
         Begin VB.Label lblActionTest 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "처방코드"
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
            Height          =   285
            Index           =   3
            Left            =   4080
            TabIndex        =   74
            Top             =   8640
            Width           =   1125
         End
         Begin VB.Image imgDelete 
            Height          =   1260
            Left            =   2280
            Picture         =   "frmMain.frx":17657
            Top             =   5490
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.Shape shpA 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   405
            Index           =   2
            Left            =   3990
            Top             =   7140
            Width           =   1335
         End
         Begin VB.Label lblActionTest 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사저장"
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
            Height          =   285
            Index           =   2
            Left            =   4080
            TabIndex        =   47
            Top             =   7230
            Width           =   1125
         End
         Begin VB.Shape shpA 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   405
            Index           =   1
            Left            =   2580
            Top             =   7140
            Width           =   1335
         End
         Begin VB.Label lblActionTest 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사삭제"
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
            Height          =   285
            Index           =   1
            Left            =   2700
            TabIndex        =   46
            Top             =   7230
            Width           =   1125
         End
         Begin VB.Label lblActionTest 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Refresh"
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
            Height          =   285
            Index           =   0
            Left            =   2670
            TabIndex        =   44
            Top             =   8640
            Width           =   1125
         End
         Begin VB.Shape shpA 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   405
            Index           =   0
            Left            =   2580
            Top             =   8550
            Width           =   1335
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
            Index           =   23
            Left            =   3390
            TabIndex        =   43
            Top             =   4530
            Width           =   825
         End
         Begin VB.Image imgSave 
            Height          =   1260
            Left            =   3840
            Picture         =   "frmMain.frx":19471
            Top             =   5460
            Visible         =   0   'False
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
            Left            =   3930
            TabIndex        =   22
            Top             =   3630
            Width           =   825
         End
      End
      Begin FPSpread.vaSpread spdTest 
         Height          =   9195
         Left            =   270
         TabIndex        =   7
         Top             =   270
         Width           =   14325
         _Version        =   393216
         _ExtentX        =   25268
         _ExtentY        =   16219
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   19
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmMain.frx":1B1BA
      End
   End
   Begin VB.Frame frame2 
      BackColor       =   &H00F8E4D8&
      Height          =   9645
      Left            =   780
      TabIndex        =   96
      Top             =   3240
      Visible         =   0   'False
      Width           =   20685
      Begin VB.CheckBox chkRAll 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Check1"
         Height          =   315
         Left            =   600
         TabIndex        =   129
         Top             =   240
         Width           =   195
      End
      Begin FPSpread.vaSpread spdROrder 
         Height          =   9375
         Left            =   60
         TabIndex        =   99
         Top             =   180
         Width           =   20535
         _Version        =   393216
         _ExtentX        =   36221
         _ExtentY        =   16536
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   20
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
         MaxCols         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14548991
         SpreadDesigner  =   "frmMain.frx":1BD26
         UserResize      =   2
      End
      Begin VB.CommandButton cmdRSL 
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
         Height          =   375
         Left            =   90
         TabIndex        =   130
         Top             =   210
         Visible         =   0   'False
         Width           =   435
      End
      Begin FPSpread.vaSpread spdRResult 
         Height          =   9360
         Left            =   13620
         TabIndex        =   100
         Top             =   180
         Visible         =   0   'False
         Width           =   6960
         _Version        =   393216
         _ExtentX        =   12277
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
         MaxCols         =   12
         MaxRows         =   50
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   13697023
         SpreadDesigner  =   "frmMain.frx":2026E
         TextTip         =   2
      End
      Begin VB.CommandButton Command1 
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
         Height          =   375
         Left            =   90
         TabIndex        =   98
         Top             =   210
         Width           =   435
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Check1"
         Height          =   315
         Left            =   570
         TabIndex        =   97
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "파일"
      Begin VB.Menu mnuExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   "설정"
      Begin VB.Menu mnuComm 
         Caption         =   "통신설정"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTest 
         Caption         =   "검사설정"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBarcodeUse 
         Caption         =   "바코드사용"
         Begin VB.Menu mnuBarcode 
            Caption         =   "바코드사용"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSeqno 
            Caption         =   "순번사용"
         End
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "결과전송"
         Begin VB.Menu mnuSaveAuto 
            Caption         =   "자동"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSaveManual 
            Caption         =   "수동"
         End
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveResult 
         Caption         =   "적용결과"
         Begin VB.Menu mnuEqpResult 
            Caption         =   "장비결과"
         End
         Begin VB.Menu mnuLisResult 
            Caption         =   "LIS결과"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuMenu03 
      Caption         =   "기타"
      Begin VB.Menu mnuHelp01 
         Caption         =   "원격지원(TeamViewer)"
      End
      Begin VB.Menu mnuHelp02 
         Caption         =   "원격지원(LG Uplus)"
      End
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCommTest 
         Caption         =   "통신테스트"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub chkAll_Click()
    Dim iRow As Long
    
    With spdOrder
        If chkAll.Value = 1 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = 1
                
                .Value = 1
            Next iRow
        ElseIf chkAll.Value = 0 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = 1
                
                .Value = 0
            Next iRow
        End If
    End With
End Sub

Private Sub chkRAll_Click()
    Dim iRow As Long
    
    With spdROrder
        If chkRAll.Value = 1 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = 1
                
                .Value = 1
            Next iRow
        ElseIf chkRAll.Value = 0 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = 1
                
                .Value = 0
            Next iRow
        End If
    End With
End Sub

'Private Sub cmdRefresh_Click()
'
'    Call GetTestList
'
'End Sub

Private Sub cmdAppend_Click()

    spdOrdMst.MaxRows = spdOrdMst.MaxRows + 1
    
End Sub

Private Sub cmdConfig_Click()
    
    frmHospInfo.Show vbModal
    
End Sub

Private Sub cmdDelete_Click()
    
    spdOrdMst.Row = spdOrdMst.ActiveRow
    spdOrdMst.Action = ActionDeleteRow
    
    spdOrdMst.MaxRows = spdOrdMst.MaxRows - 1
    
End Sub

Private Sub cmdIF_Click()

    If FraHidden.Visible = True Then
        FraHidden.Visible = False
    Else
        FraHidden.Visible = True
        FraHidden.ZOrder 0
    End If
    
End Sub



Private Sub cmdOrder_Click()
    Dim lngFIleNum  As Long
    Dim strCFXFile  As String
    
    Dim strBarno    As String
    Dim iCnt        As Integer
    Dim varTmp      As Variant
    Dim ORDERPATH   As String
    Dim i           As Integer

    With CFXFile
        .CancelError = True
        .FileName = gComm.ORDPATH & "LIS.lis"
        If Len(Dir(.FileName)) Then
             Close #lngFIleNum
             Kill .FileName
        End If
        lngFIleNum = FreeFile
        
        Open .FileName For Append As #lngFIleNum

        strCFXFile = ""
        For iCnt = 1 To spdOrder.MaxRows
            spdOrder.GetText 1, iCnt, varTmp
            If GetText(spdOrder, iCnt, colCHECKBOX) = "1" Then
                strBarno = GetText(spdOrder, iCnt, colBARCODE)
                If strBarno <> "" Then
                    strCFXFile = strCFXFile & CStr(iCnt) & Space(5 - Len(CStr(iCnt)))
                    strCFXFile = strCFXFile & strBarno & Space(20 - Len(strBarno))
'                    strCFXFile = strCFXFile & "HPV28" & Space(10) & Space(135)
                    strCFXFile = strCFXFile & gAllOrdCd1 & Space(150 - Len(gAllOrdCd1))
                    
                    Call SetText(spdOrder, "", iCnt, colCHECKBOX)
                End If
            End If
        Next
        
        If strCFXFile <> "" Then
            Print #lngFIleNum, strCFXFile
            MsgBox "오더 파일 생성 완료", vbOKOnly + vbInformation, Me.Caption
        End If
        strCFXFile = ""
        Close #lngFIleNum
        
    End With
End Sub

Private Sub cmdPrint_Click()
'    Dim intRow  As Integer
'    Dim intCol  As Integer
'
'    Erase varClipData
'
'
'    With spdROrder
'        For intRow = 1 To .DataRowCnt
'            gRowCnt = .DataRowCnt
'
'            For intCol = 0 To .MaxCols
'                .Row = intRow
'                .Col = intCol
'                If intCol = 1 Then
'                    varClipData(intRow, intCol) = intRow
'                Else
'                    varClipData(intRow, intCol) = .Text
'                End If
'            Next
'
'        Next
'
'        Call SetPrint
'        Erase varClipData
'
'        vasHPVReport.PrintOrientation = PrintOrientationLandscape '가로출력
'        vasHPVReport.Action = 13
'    End With
'
    


    Dim iRow As Integer
    Dim J As Integer
    
    vasPrint.MaxRows = 1
    vasPrint.MaxRows = 2
    
    With spdROrder
        For iRow = 1 To .MaxRows
            If iRow = 1 Then
                J = 1
                'vasPrint.MaxRows =
                vasPrint.Row = 1: vasPrint.Col = 9: vasPrint.Text = "총건수 : " & spdROrder.MaxRows & " 건"
                vasPrint.Row = 1: vasPrint.Col = 12: vasPrint.Text = "검사일 : " & Format(dtpToday, "yyyy-mm-dd")
                vasPrint.Row = 1: vasPrint.Col = 15: vasPrint.Text = "검사자 : " & gHOSP.USERNM
                vasPrint.Row = 1: vasPrint.Col = 18: vasPrint.Text = "담당의 :"
                
                SetText vasPrint, "No", .MaxRows, 1
                SetText vasPrint, Trim(GetText(spdROrder, 0, colHOSPDATE)), 2, 2
                SetText vasPrint, Trim(GetText(spdROrder, 0, colCHARTNO)), 2, 3:   vasPrint.AddCellSpan 3, 2, 2, 1
                SetText vasPrint, Trim(GetText(spdROrder, 0, colPOSNO)), 2, 5:     vasPrint.AddCellSpan 5, 2, 2, 1
                SetText vasPrint, Trim(GetText(spdROrder, 0, colPID)), 2, 7:       vasPrint.AddCellSpan 7, 2, 2, 1
                SetText vasPrint, Trim(GetText(spdROrder, 0, colPNAME)), 2, 9:    vasPrint.AddCellSpan 9, 2, 2, 1
                SetText vasPrint, Trim(GetText(spdROrder, 0, colPSEX)), 2, 11
                SetText vasPrint, Trim(GetText(spdROrder, 0, colPAGE)), 2, 12:     vasPrint.AddCellSpan 13, 2, 6, 1
                SetText vasPrint, Space(30) & "HPV Result", 2, 13
                
                vasPrint.RowHeight(2) = 20
                
                
            End If
            
            'j = 1
            
            .Row = iRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                vasPrint.MaxRows = vasPrint.MaxRows + 1
                J = iRow + 2
                SetText vasPrint, CStr(iRow), J, 1
                SetText vasPrint, Trim(GetText(spdROrder, iRow, colHOSPDATE)), J, 2
                SetText vasPrint, Trim(GetText(spdROrder, iRow, colCHARTNO)), J, 3:   vasPrint.AddCellSpan 3, J, 2, 1
                SetText vasPrint, Trim(GetText(spdROrder, iRow, colPOSNO)), J, 5:     vasPrint.AddCellSpan 5, J, 2, 1
                SetText vasPrint, Trim(GetText(spdROrder, iRow, colPID)), J, 7:       vasPrint.AddCellSpan 7, J, 2, 1
                SetText vasPrint, Trim(GetText(spdROrder, iRow, colPNAME)), J, 9:    vasPrint.AddCellSpan 9, J, 2, 1
                SetText vasPrint, Trim(GetText(spdROrder, iRow, colPSEX)), J, 11
                SetText vasPrint, Trim(GetText(spdROrder, iRow, colPAGE)), J, 12:     vasPrint.AddCellSpan 13, J, 6, 1
                SetText vasPrint, Trim(GetText(spdROrder, iRow, colITEMS)), J, 13
                
                vasPrint.RowHeight(J) = 20
                
            End If
        Next iRow
        
        'vasPrint.RowHeight(-1) = 40

    End With
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "출력할 자료를 선택하세요", , "알 림"
        Exit Sub
    Else
        vasPrint.PrintOrientation = PrintOrientationLandscape '가로출력
        vasPrint.Action = 13
    End If
    

End Sub

Private Sub SetPrint()
    Dim intCnt      As Integer
    Dim i           As Integer
    Dim J           As Integer
    Dim K           As Integer
    Dim varTmp      As Variant
    Dim varTmp2     As Variant
    Dim strTmp      As String
    Dim strINVALID  As String
    
    Dim strHData     As String
    Dim strLData     As String
    Dim strHPData    As String
    Dim strLPData    As String
    
    Dim varHNum     As Variant
    Dim varLNum     As Variant
    
    Dim blnHPos     As Boolean
    Dim blnLPos     As Boolean
    Dim blnMulti    As Boolean
    
    Dim strICData   As String
    Dim varICData   As Variant
    
    Dim varTemp1    As Variant
    Dim varTemp2    As Variant
    
'    blnPos = False
'    blnMulti = False
    
    With vasHPVReport
        .MaxRows = 8
        .MaxRows = (gRowCnt * 4) + 6
        .RowHeight(-1) = 18
        .RowHeight(2) = 25
        
        For i = 0 To gRowCnt - 2
            .AddCellSpan 2, 8 + (i * 4) + 1, 1, 4
            .AddCellSpan 3, 8 + (i * 4) + 1, 1, 4
            .AddCellSpan 19, 8 + (i * 4) + 1, 1, 4
            
            .SetCellBorder 2, 8 + (i * 4) + 1, 2, 8 + (i * 4) + 4, SS_BORDER_TYPE_OUTLINE, vbBlack, CellBorderStyleSolid
            .SetCellBorder 3, 8 + (i * 4) + 1, 3, 8 + (i * 4) + 4, SS_BORDER_TYPE_OUTLINE, vbBlack, CellBorderStyleSolid
            .SetCellBorder 19, 8 + (i * 4) + 1, 19, 8 + (i * 4) + 4, SS_BORDER_TYPE_OUTLINE, vbBlack, CellBorderStyleSolid
            
            .SetCellBorder 2, 8 + (i * 4) + 1, 2, 8 + (i * 4) + 1, SS_BORDER_TYPE_OUTLINE, vbBlack, CellBorderStyleSolid
            .SetCellBorder 3, 8 + (i * 4) + 1, 3, 8 + (i * 4) + 1, SS_BORDER_TYPE_OUTLINE, vbBlack, CellBorderStyleSolid
            .SetCellBorder 19, 8 + (i * 4) + 1, 19, 8 + (i * 4) + 1, SS_BORDER_TYPE_OUTLINE, vbBlack, CellBorderStyleSolid
            
            For K = 1 To 4
                For J = 4 To 18
                    .SetCellBorder J, 8 + (i * 4) + K, J, 8 + (i * 4) + K, SS_BORDER_TYPE_OUTLINE, vbBlack, CellBorderStyleSolid
                    
                    .Row = 8 + (i * 4) + K
                    .Col = J
                    
                    If K = 1 Then
                        Select Case J
                            'A set
                            Case 4:  .SetText J, 8 + (i * 4) + K, "66": .ForeColor = vbRed
                            Case 5:  .SetText J, 8 + (i * 4) + K, "45": .ForeColor = vbRed
                            Case 6:  .SetText J, 8 + (i * 4) + K, "58": .ForeColor = vbRed
                            Case 7:  .SetText J, 8 + (i * 4) + K, "51": .ForeColor = vbRed
                            Case 8:  .SetText J, 8 + (i * 4) + K, "59": .ForeColor = vbRed
                            Case 9:  .SetText J, 8 + (i * 4) + K, "16": .ForeColor = vbRed
                            Case 10: .SetText J, 8 + (i * 4) + K, "33": .ForeColor = vbRed
                            Case 11: .SetText J, 8 + (i * 4) + K, "39": .ForeColor = vbRed
                            Case 12: .SetText J, 8 + (i * 4) + K, "52": .ForeColor = vbRed
                            Case 13: .SetText J, 8 + (i * 4) + K, "35": .ForeColor = vbRed
                            Case 14: .SetText J, 8 + (i * 4) + K, "18": .ForeColor = vbRed
                            Case 15: .SetText J, 8 + (i * 4) + K, "56": .ForeColor = vbRed
                            Case 16: .SetText J, 8 + (i * 4) + K, "68": .ForeColor = vbRed
                            Case 17: .SetText J, 8 + (i * 4) + K, "31": .ForeColor = vbRed
                            Case 18: .SetText J, 8 + (i * 4) + K, "IC": .ForeColor = vbBlack
                        End Select
                    End If
                    
                    If K = 3 Then
                        Select Case J
                            'B set
                            Case 4:  .SetText J, 8 + (i * 4) + K, "26": .ForeColor = vbRed
                            Case 5:  .SetText J, 8 + (i * 4) + K, "69": .ForeColor = vbRed
                            Case 6:  .SetText J, 8 + (i * 4) + K, "73": .ForeColor = vbRed
                            Case 7:  .SetText J, 8 + (i * 4) + K, "42": .ForeColor = vbBlue
                            Case 8:  .SetText J, 8 + (i * 4) + K, "82": .ForeColor = vbRed
                            Case 9:  .SetText J, 8 + (i * 4) + K, "53": .ForeColor = vbRed
                            Case 10: .SetText J, 8 + (i * 4) + K, "43": .ForeColor = vbBlue
                            Case 11: .SetText J, 8 + (i * 4) + K, "54": .ForeColor = vbBlue
                            Case 12: .SetText J, 8 + (i * 4) + K, "70": .ForeColor = vbBlue
                            Case 13: .SetText J, 8 + (i * 4) + K, "61": .ForeColor = vbBlue
                            Case 14: .SetText J, 8 + (i * 4) + K, "6":  .ForeColor = vbBlue
                            Case 15: .SetText J, 8 + (i * 4) + K, "44": .ForeColor = vbBlue
                            Case 16: .SetText J, 8 + (i * 4) + K, "40": .ForeColor = vbBlue
                            Case 17: .SetText J, 8 + (i * 4) + K, "11": .ForeColor = vbBlue
                            Case 18: .SetText J, 8 + (i * 4) + K, "IC": .ForeColor = vbBlack
                        End Select
                    End If
                    
                Next
            Next
        Next
        
        .Row = 2: .Col = 10: .Text = "총건수 : " & gRowCnt & " 건"
        .Row = 2: .Col = 13: .Text = "검사일 : " & Format(varClipData(1, colHOSPDATE), "####-##-##")
        .Row = 2: .Col = 16: .Text = "검사자/결과입력 : 이재왕"
        .Row = 2: .Col = 19: .Text = "담당의:임소영    "
        
        
        ' 초기화
        For i = 6 To .MaxRows - 2 Step 2
            For J = 4 To 17
                .Row = i: .Col = J: .Text = "-"
            Next
        Next
        
        
        strHData = ""
        strLData = ""
        
        
        For intCnt = 1 To 100
            .Row = (intCnt * 4) + 1: .Col = 2: .Text = varClipData(intCnt, 1)
            .Row = (intCnt * 4) + 1: .Col = 3: .Text = varClipData(intCnt, 5)
            
            If varClipData(intCnt, 1) = "" Then
                Exit Sub
            End If
            'varTmp = Split(varClipData(intCnt, colITEMS), vbNewLine)
            varTmp = Split(varClipData(intCnt, colITEMS), vbLf)
        
            strHData = ""
            strLData = ""
            
            For i = 0 To UBound(varTmp)
                If varTmp(i) <> "" Then
                    If mGetP(varTmp(i), 1, ":") = "HPV High Risk Type " Then
                        If UCase(Trim(mGetP(mGetP(varTmp(i), 2, ":"), 1, "("))) = "POSITIVE" Then
                            strHData = "(" & mGetP(mGetP(varTmp(i), 2, ":"), 2, "(")
                            varHNum = Split(strHData, ",")
                            If UBound(varHNum) >= 1 Then
                                blnMulti = True
                            Else
                                blnMulti = False
                            End If
                            blnHPos = True
                        Else
                            strHData = "( )"
                        End If
                    ElseIf mGetP(varTmp(i), 1, ":") = "HPV Low Risk Type " Then
                        If UCase(Trim(mGetP(mGetP(varTmp(i), 2, ":"), 1, "("))) = "POSITIVE" Then
                            strLData = "(" & mGetP(mGetP(varTmp(i), 2, ":"), 2, "(")
                            varLNum = Split(strLData, ",")
                            If UBound(varLNum) >= 1 Then
                                blnMulti = True
                            Else
                                blnMulti = False
                            End If
                            blnLPos = True
                        Else
                            strLData = "( )"
                        End If
                    ElseIf Mid(varTmp(i), 1, 2) = "IC" Then
                        strICData = Mid(varTmp(i), 4)
                        If InStr(strICData, "/") > 0 Then
                            varICData = Split(strICData, "/")
                            If varICData(0) > varICData(1) Then
                                strICData = varICData(0)
                            Else
                                strICData = varICData(1)
                            End If
                        End If
                        .Row = 24: .Col = 16: .Text = "IC " & strICData            'IC Value
                        strINVALID = strICData
                        Exit For
                    End If
                End If
            Next
            
            
            
            
    '        Erase varTemp1
    '        Erase varTemp2
            
            '-- High Set
            If strHData <> "( )" And strHData <> "" Then
                varTemp1 = Trim(strHData)
                varTemp1 = Replace(varTemp1, "(", "")
                varTemp1 = Replace(varTemp1, ")", "")
                varTemp1 = Trim(varTemp1)
                If varTemp1 <> "" Then
                    varTemp1 = Split(varTemp1, ",")
                    For i = 0 To UBound(varTemp1)
                        varTemp2 = Mid(varTemp1(i), 1, InStr(varTemp1(i), "+") - 1)
                        Select Case varTemp2
                            'A set
                            Case "66": .Row = (intCnt * 4) + 2: .Col = 4
                            Case "45": .Row = (intCnt * 4) + 2: .Col = 5
                            Case "58": .Row = (intCnt * 4) + 2: .Col = 6
                            Case "51": .Row = (intCnt * 4) + 2: .Col = 7
                            Case "59": .Row = (intCnt * 4) + 2: .Col = 8
                            Case "16": .Row = (intCnt * 4) + 2: .Col = 9
                            Case "33": .Row = (intCnt * 4) + 2: .Col = 10
                            Case "39": .Row = (intCnt * 4) + 2: .Col = 11
                            Case "52": .Row = (intCnt * 4) + 2: .Col = 12
                            Case "35": .Row = (intCnt * 4) + 2: .Col = 13
                            Case "18": .Row = (intCnt * 4) + 2: .Col = 14
                            Case "56": .Row = (intCnt * 4) + 2: .Col = 15
                            Case "68": .Row = (intCnt * 4) + 2: .Col = 16
                            Case "31": .Row = (intCnt * 4) + 2: .Col = 17
                            'B set
                            Case "26": .Row = (intCnt * 4) + 4: .Col = 4
                            Case "69": .Row = (intCnt * 4) + 4: .Col = 5
                            Case "73": .Row = (intCnt * 4) + 4: .Col = 6
                            Case "42": .Row = (intCnt * 4) + 4: .Col = 7
                            Case "82": .Row = (intCnt * 4) + 4: .Col = 8
                            Case "53": .Row = (intCnt * 4) + 4: .Col = 9
                            Case "43": .Row = (intCnt * 4) + 4: .Col = 10
                            Case "54": .Row = (intCnt * 4) + 4: .Col = 11
                            Case "70": .Row = (intCnt * 4) + 4: .Col = 12
                            Case "61": .Row = (intCnt * 4) + 4: .Col = 13
                            Case "6":  .Row = (intCnt * 4) + 4: .Col = 14
                            Case "44": .Row = (intCnt * 4) + 4: .Col = 15
                            Case "40": .Row = (intCnt * 4) + 4: .Col = 16
                            Case "11": .Row = (intCnt * 4) + 4: .Col = 17
                            
                        End Select
                        .Text = Mid(varTemp1(i), InStr(varTemp1(i), "+"))
                    Next
                End If
            End If
                        
            '-- Low set
            If strLData <> "( )" And strLData <> "" Then
                varTemp1 = Trim(strLData)
                varTemp1 = Replace(varTemp1, "(", "")
                varTemp1 = Replace(varTemp1, ")", "")
                varTemp1 = Trim(varTemp1)
                If varTemp1 <> "" Then
                    varTemp1 = Split(varTemp1, ",")
                    For i = 0 To UBound(varTemp1)
                        varTemp2 = Mid(varTemp1(i), 1, InStr(varTemp1(i), "+") - 1)
                        Select Case varTemp2
                            'A set
                            Case "66": .Row = (intCnt * 4) + 2: .Col = 3
                            Case "45": .Row = (intCnt * 4) + 2: .Col = 4
                            Case "58": .Row = (intCnt * 4) + 2: .Col = 5
                            Case "51": .Row = (intCnt * 4) + 2: .Col = 6
                            Case "59": .Row = (intCnt * 4) + 2: .Col = 7
                            Case "16": .Row = (intCnt * 4) + 2: .Col = 8
                            Case "33": .Row = (intCnt * 4) + 2: .Col = 9
                            Case "39": .Row = (intCnt * 4) + 2: .Col = 10
                            Case "52": .Row = (intCnt * 4) + 2: .Col = 11
                            Case "35": .Row = (intCnt * 4) + 2: .Col = 12
                            Case "18": .Row = (intCnt * 4) + 2: .Col = 13
                            Case "56": .Row = (intCnt * 4) + 2: .Col = 14
                            Case "68": .Row = (intCnt * 4) + 2: .Col = 15
                            Case "31": .Row = (intCnt * 4) + 2: .Col = 16
                            'B set
                            Case "26": .Row = (intCnt * 4) + 4: .Col = 3
                            Case "69": .Row = (intCnt * 4) + 4: .Col = 4
                            Case "73": .Row = (intCnt * 4) + 4: .Col = 5
                            Case "42": .Row = (intCnt * 4) + 4: .Col = 6
                            Case "82": .Row = (intCnt * 4) + 4: .Col = 7
                            Case "53": .Row = (intCnt * 4) + 4: .Col = 8
                            Case "43": .Row = (intCnt * 4) + 4: .Col = 9
                            Case "54": .Row = (intCnt * 4) + 4: .Col = 10
                            Case "70": .Row = (intCnt * 4) + 4: .Col = 11
                            Case "61": .Row = (intCnt * 4) + 4: .Col = 12
                            Case "6":  .Row = (intCnt * 4) + 4: .Col = 13
                            Case "44": .Row = (intCnt * 4) + 4: .Col = 14
                            Case "40": .Row = (intCnt * 4) + 4: .Col = 15
                            Case "11": .Row = (intCnt * 4) + 4: .Col = 16
                            
                        End Select
                        .Text = Mid(varTemp1(i), InStr(varTemp1(i), "+"))
                    Next
                End If
            End If
            
            .Row = (intCnt * 4) + 2: .Col = 18
            .Text = varClipData(intCnt, colOCNT)
            
            .Row = (intCnt * 4) + 4: .Col = 18
            .Text = varClipData(intCnt, colRCNT)
            
            strHPData = ""
            strHPData = strHData
            strHPData = Replace(strHPData, "(", "")
            strHPData = Replace(strHPData, ")", "")
            strHPData = Replace(strHPData, "HPV Low Risk Type", "")
            
            strLPData = ""
            strLPData = strLData
            strLPData = Replace(strLPData, "(", "")
            strLPData = Replace(strLPData, ")", "")
            
            .Row = (intCnt * 4) + 1: .Col = 19: .Text = strHPData & strLPData
            
        Next
        
        
    End With
End Sub

Private Sub cmdResult_Click()
    Dim intRow      As Integer
    Dim intIdx      As Integer
    Dim strSrcfile  As String
    Dim strDestFile As String
    Dim strBuffer   As String
    Dim strtmpBuf   As String
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i, J        As Long
    Dim intCnt      As Integer
    Dim varTmp      As Variant
    Dim RESULTPATH   As String

    Dim strBarno As String
    Dim strOldBarno As String
    Dim strNewBarno As String

On Error GoTo ErrRoutine

    strSrcfile = ""

    With CFXFile
        .InitDir = gComm.RSTPATH
        .FileName = "*.CSV"
        .Filter = "Resource CSV (*.CSV)|*.CSV|All File (*.*)|*.*|"
        .DialogTitle = "CFX96 자료 읽어오기"
        .ShowOpen
    End With
        
    strSrcfile = CFXFile.FileName

    If strSrcfile = "" Then
        Exit Sub
    End If
        
    Open strSrcfile For Input As #3

    strBuffer = ""
    Do While Not EOF(3)
        strBuffer = strBuffer & Input(1, #3)
    Loop

    Close #3
    
    
    varTmp = Split(strBuffer, vbLf)
    J = 1
    For i = 1 To UBound(varTmp)
        ReDim Preserve strRecvData(J)
        strRecvData(J) = varTmp(i)
        'strBuffer = varTmp(i)
        Debug.Print varTmp(i)
        
        strBarno = mGetP(varTmp(i), 1, ",")
        strNewBarno = mGetP(varTmp(i + 1), 1, ",")
        
        If strBarno = "" Then
            strBarno = strOldBarno
        End If
        
        
'        If mGetP(varTmp(i), 3, ",") <> "" And i > 1 Then
'            Call FILE_Protocol
'            Erase strRecvData
'            J = 1
'           ' Exit Sub
'        End If
        
        If strNewBarno <> "" Then
            Call FILE_Protocol
            Erase strRecvData
            J = 1
        Else
            J = J + 1
        End If

        strOldBarno = strBarno
        
        If mGetP(varTmp(i), 3, ",") = "NC" Then
            Exit For
        End If
    Next i
    
    pBuffer = strBuffer
    
    Call FILE_Protocol
    
Exit Sub

    
Exit Sub

ErrRoutine:

End Sub

Private Sub cmdRPrint_Click()


    Dim iRow As Integer
    Dim J As Integer
    
    vasPrint.MaxRows = 1
    vasPrint.MaxRows = 2
    
    With spdOrder
        For iRow = 1 To .MaxRows
            If iRow = 1 Then
                J = 1
                'vasPrint.MaxRows =
                vasPrint.Row = 1: vasPrint.Col = 9: vasPrint.Text = "총건수 : " & spdOrder.MaxRows & " 건"
                vasPrint.Row = 1: vasPrint.Col = 12: vasPrint.Text = "검사일 : " & Format(dtpToday, "yyyy-mm-dd")
                vasPrint.Row = 1: vasPrint.Col = 15: vasPrint.Text = "검사자 : " & gHOSP.USERNM
                vasPrint.Row = 1: vasPrint.Col = 18: vasPrint.Text = "담당의 :"
                
                SetText vasPrint, "No", .MaxRows, 1
                SetText vasPrint, Trim(GetText(spdOrder, 0, colHOSPDATE)), 2, 2
                SetText vasPrint, Trim(GetText(spdOrder, 0, colCHARTNO)), 2, 3:   vasPrint.AddCellSpan 3, 2, 2, 1
                SetText vasPrint, Trim(GetText(spdOrder, 0, colPOSNO)), 2, 5:     vasPrint.AddCellSpan 5, 2, 2, 1
                SetText vasPrint, Trim(GetText(spdOrder, 0, colPID)), 2, 7:       vasPrint.AddCellSpan 7, 2, 2, 1
                SetText vasPrint, Trim(GetText(spdOrder, 0, colPNAME)), 2, 9:    vasPrint.AddCellSpan 9, 2, 2, 1
                SetText vasPrint, Trim(GetText(spdOrder, 0, colPSEX)), 2, 11
                SetText vasPrint, Trim(GetText(spdOrder, 0, colPAGE)), 2, 12:     vasPrint.AddCellSpan 13, 2, 6, 1
                SetText vasPrint, Space(30) & "HPV Result", 2, 13
                
                vasPrint.RowHeight(2) = 20
                
                
            End If
            
            'j = 1
            
            .Row = iRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                vasPrint.MaxRows = vasPrint.MaxRows + 1
                J = iRow + 2
                SetText vasPrint, CStr(iRow), J, 1
                SetText vasPrint, Trim(GetText(spdOrder, iRow, colHOSPDATE)), J, 2
                SetText vasPrint, Trim(GetText(spdOrder, iRow, colCHARTNO)), J, 3:   vasPrint.AddCellSpan 3, J, 2, 1
                SetText vasPrint, Trim(GetText(spdOrder, iRow, colPOSNO)), J, 5:     vasPrint.AddCellSpan 5, J, 2, 1
                SetText vasPrint, Trim(GetText(spdOrder, iRow, colPID)), J, 7:       vasPrint.AddCellSpan 7, J, 2, 1
                SetText vasPrint, Trim(GetText(spdOrder, iRow, colPNAME)), J, 9:    vasPrint.AddCellSpan 9, J, 2, 1
                SetText vasPrint, Trim(GetText(spdOrder, iRow, colPSEX)), J, 11
                SetText vasPrint, Trim(GetText(spdOrder, iRow, colPAGE)), J, 12:     vasPrint.AddCellSpan 13, J, 6, 1
                SetText vasPrint, Trim(GetText(spdOrder, iRow, colITEMS)), J, 13
                
                vasPrint.RowHeight(J) = 20
                
            End If
        Next iRow
        
        'vasPrint.RowHeight(-1) = 40

    End With
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "출력할 자료를 선택하세요", , "알 림"
        Exit Sub
    Else
        vasPrint.PrintOrientation = PrintOrientationLandscape '가로출력
        vasPrint.Action = 13
    End If
    
End Sub

Private Sub cmdSearch_Click()
'    With CFXFile
'        .InitDir = gComm.FILEPATH
'        .FileName = "*.CSV"
'        .Filter = "Resource CSV (*.CSV)|*.CSV|All File (*.*)|*.*|"
'        .DialogTitle = "CFX96 자료 읽어오기"
'        .ShowOpen
'    End With
'
'    strSrcfile = CFXFile.FileName
'
'    If strSrcfile = "" Then
'        Exit Sub
'    End If

End Sub

Private Sub cmdSend_Click()
    Dim i As Integer
    Dim varTmp As Variant
    
    Erase strRecvData
    varTmp = Replace(txtRcv.Text, vbLf, "")
    varTmp = Split(varTmp, vbCr)
    
    For i = 0 To UBound(varTmp)
        ReDim Preserve strRecvData(i + 1)
        strRecvData(i + 1) = varTmp(i)
    Next
    
    Select Case UCase(gHOSP.MACHNM)
        Case "E411"
                Call Phase_Serial_E411
        Case "AU400"
                'Call Phase_Serial_AU400
                Call SerialRcvData_AU400
        Case "AU480"
                Call Phase_Serial_AU480
        Case "XN1000"
                Call SerialRcvData_XN1000
        Case Else
            
    End Select



End Sub

Private Sub cmdSeqDown_Click()
    On Error Resume Next
    
    txtSeq.Text = txtSeq.Text - 1

End Sub

Private Sub cmdSeqUp_Click()
    On Error Resume Next
    
    txtSeq.Text = txtSeq.Text + 1

End Sub

Private Sub cmdSet_Click()

    If frameSet.Visible = True Then
        frameSet.Visible = False
        
    Else
        frameSet.Visible = True
        frameSet.ZOrder 0
    End If
    
End Sub

Private Sub cmdSL_Click()

    If cmdSL.Caption = "▶" Then
        cmdSL.Caption = "◀"
        spdOrder.Width = Me.Width - 400
    Else
        cmdSL.Caption = "▶"
        spdOrder.Width = Me.ScaleWidth - spdResult.Width - 280
    End If
    
End Sub

Private Sub cmdSpecDown_Click()
    On Error Resume Next
    
    txtResSpec.Text = txtResSpec.Text - 1

End Sub

Private Sub cmdSpecUP_Click()
    On Error Resume Next
    
    txtResSpec.Text = txtResSpec.Text + 1

End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    If comEqp.PortOpen = True Then
        comEqp.PortOpen = False
    End If

    Call DisConnect_Server
    
    Call DisConnect_Local
    
    Unload Me
    
    End
    
End Sub

Private Sub fraResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblResult.ForeColor = vbBlack
    shpR.BorderColor = &H808080

End Sub



Private Sub lblFileSave_Click()
    

    If optComType(0).Value = True Then
        Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\OKSOFT.ini")
    ElseIf optComType(1).Value = True Then
        Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\OKSOFT.ini")
    Else
        Call WritePrivateProfileString("COMM", "COMTYPE", "3", App.PATH & "\OKSOFT.ini")
    End If

    
    Call WritePrivateProfileString("COMM", "ORDPATH", txtOrdPath.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("COMM", "RSTPATH", txtRstPath.Text, App.PATH & "\OKSOFT.ini")
    
    GetSetup
    
    GetCommList
    
End Sub

Private Sub lblResult_Click()

    frmMain.spdROrder.MaxRows = 0
    frmMain.spdRResult.MaxRows = 0

    Call GetResultList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), cboRstType.ListIndex, cboState.ListIndex)
    
End Sub

Private Sub lblResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblResult.ForeColor = vbBlack
    shpR.BorderColor = &H808080
    
    lblResult.ForeColor = vbBlue
    shpR.BorderColor = vbCyan
    
End Sub

Private Sub lblSave_Click()
    Dim lRow    As Long
    Dim Res     As Long
    
    With spdOrder
        For lRow = 1 To .DataRowCnt
            .Row = lRow
            .Col = 1
            If .Value = 1 Then
                
                Res = SaveTransData(lRow)
            
                If Res = -1 Then
                    SetForeColor spdOrder, lRow, lRow, 1, colSTATE, 255, 0, 0
                    SetText spdOrder, "Failed", lRow, colSTATE
                Else
                    .Row = lRow
                    .Col = 1
                    .Value = 1
                    
                    SetBackColor spdOrder, lRow, lRow, 1, colSTATE, 202, 255, 112
                    SetText spdOrder, "Trans", lRow, colSTATE
                    
                          SQL = " UPDATE PATRESULT SET " & vbCrLf
                    SQL = SQL & "  SENDFLAG = '2' " & vbCrLf
                    SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.HOSPCD & "' " & vbCrLf
                    SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
                    
                    If DBExec(AdoCn_Local, SQL) Then
                    
                    End If
'                    Res = SendQuery(gLocal, SQL)
'                    If Res = -1 Then
'                        SaveQuery SQL
'                        Exit Sub
'                    End If
                    
                End If
                .Row = lRow
                .Col = 1
                .Value = 0
            End If
        Next lRow
    End With
End Sub

Private Sub mnuBarcode_Click()
    
    mnuBarcode.Checked = True
    mnuSeqno.Checked = False
    
    Call WritePrivateProfileString("HOSP", "BARUSE", "Y", App.PATH & "\OKSOFT.ini")

End Sub

Private Sub mnuComm_Click()
    
    Call lblMenu_Click(3)

End Sub

Private Sub mnuCommTest_Click()

    If fraCommTest.Visible = False Then
        fraCommTest.Visible = True
    Else
        fraCommTest.Visible = False
    End If
    
End Sub

Private Sub mnuEqpResult_Click()
    
    mnuEqpResult.Checked = True
    mnuLisResult.Checked = False
    
    Call WritePrivateProfileString("HOSP", "SAVELIS", "N", App.PATH & "\OKSOFT.ini")

End Sub

Private Sub mnuLisResult_Click()
    
    mnuEqpResult.Checked = False
    mnuLisResult.Checked = True
    
    Call WritePrivateProfileString("HOSP", "SAVELIS", "Y", App.PATH & "\OKSOFT.ini")

End Sub

Private Sub mnuSaveAuto_Click()
    
    mnuSaveAuto.Checked = True
    mnuSaveManual.Checked = False
    
    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "Y", App.PATH & "\OKSOFT.ini")

End Sub

Private Sub mnuSaveManual_Click()
    
    mnuSaveAuto.Checked = False
    mnuSaveManual.Checked = True
    
    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "N", App.PATH & "\OKSOFT.ini")


End Sub

Private Sub mnuSeqno_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = True
    
    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\OKSOFT.ini")

End Sub

Private Sub mnuTest_Click()
    
    Call lblMenu_Click(2)

End Sub

Private Sub spdOrder_KeyPress(KeyAscii As Integer)
    Dim sRow        As Long
    
    If KeyAscii = vbKeyReturn Then
        If colBARCODE = spdOrder.ActiveCol Then
            sRow = spdOrder.ActiveRow
            If GetSampleInfo(sRow, spdROrder) = -1 Then
                MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
            Else
                '정보수정
                SQL = ""
                SQL = SQL & "UPDATE PATRESULT SET "
                SQL = SQL & "  BARCODE = '" & Trim(GetText(spdOrder, sRow, colBARCODE)) & "'" & vbCr
                SQL = SQL & " ,INOUT   = '" & Trim(GetText(spdOrder, sRow, colINOUT)) & "'" & vbCr
                SQL = SQL & " ,CHARTNO = '" & Trim(GetText(spdOrder, sRow, colCHARTNO)) & "'" & vbCr
                SQL = SQL & " ,PID     = '" & Trim(GetText(spdOrder, sRow, colPID)) & "'" & vbCr
                SQL = SQL & " ,PNAME   = '" & Trim(GetText(spdOrder, sRow, colPNAME)) & "'" & vbCr
                SQL = SQL & " ,PSEX    = '" & Trim(GetText(spdOrder, sRow, colPSEX)) & "'" & vbCr
                SQL = SQL & " ,PAGE    = '" & Trim(GetText(spdOrder, sRow, colPAGE)) & "'" & vbCr
                SQL = SQL & " ,PJUMIN  = '" & Trim(GetText(spdOrder, sRow, colPJUMIN)) & "'" & vbCr
                SQL = SQL & " WHERE EXAMDATE = '" & Trim(GetText(spdOrder, sRow, colEXAMDATE)) & "'" & vbCr
                SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, sRow, colSAVESEQ)) & vbCr
                SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.HOSPCD & "' & vbCr"
                'SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, asRow1, colBARCODE)) & "' " & vbCr
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
        End If
    End If
End Sub

Private Sub spdROrder_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol As Integer
    
    '-- 정렬
    If Row = 0 Then
        '-- 정렬 추가
        
        Exit Sub
    End If
    
    '-- 환자정보표시
    
    '-- 결과표시
    If GetPatTRestResult_Search(Row) = -1 Then
        '장비결과가 없을경우 검사명만 보여주기
        spdResult.MaxRows = 0
        With spdOrder
            For intCol = colSTATE + 1 To .MaxCols
                If GetText(spdOrder, Row, intCol) <> "" Then    '◆
                    spdResult.MaxRows = spdResult.MaxRows + 1
                    Call SetText(spdResult, GetText(spdOrder, 0, intCol), spdResult.MaxRows, colRTESTNM)
                    spdResult.RowHeight(-1) = 12
                End If
            Next
        End With
    End If

End Sub

Private Sub spdROrder_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim intRow      As Long
    Dim strTestCd   As String
    Dim strTestNm   As String
    Dim strResult   As String
    Dim strIntBase  As String
    Dim strJudge    As String
    Dim lsID        As String
    Dim lsSeq       As Long
    Dim strExamDate As String

    sRow = spdROrder.ActiveRow
    sCol = spdROrder.ActiveCol
    
    If KeyCode = vbKeyDelete Then
        If sRow < 1 Or sRow > spdROrder.DataRowCnt Then
            Exit Sub
        End If
        
        If sCol > colSTATE Then
            Exit Sub
        End If
        
        lsSeq = Trim(GetText(spdROrder, sRow, colSAVESEQ))
        strExamDate = Trim(GetText(spdROrder, sRow, colEXAMDATE))

        If lsSeq < 1 Then
            Exit Sub
        End If

        If MsgBox(lsSeq & " 의 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If

              SQL = "DELETE FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
        SQL = SQL & "   AND SAVESEQ = " & lsSeq & vbCrLf
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & strExamDate & "' "


        If DBExec(AdoCn_Local, SQL) Then
            '-- 성공
        End If

        DeleteRow spdROrder, sRow, sRow
        spdROrder.MaxRows = spdROrder.MaxRows - 1
        spdRResult.MaxRows = 0
    End If
    
End Sub

Private Sub spdROrder_KeyPress(KeyAscii As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim intRow      As Long
    Dim strTestCd   As String
    Dim strTestNm   As String
    Dim strResult   As String
    Dim strIntBase  As String
    Dim strJudge    As String
    Dim lsID        As String
    Dim lsSeq       As Long
    Dim strExamDate As String
    
    sRow = spdROrder.ActiveRow
    sCol = spdROrder.ActiveCol
    
    If KeyAscii = vbKeyDelete Then
        If sRow < 1 Or sRow > spdROrder.DataRowCnt Then
            Exit Sub
        End If
        If sCol > colSTATE Then
            Exit Sub
        End If
        lsID = Trim(GetText(spdROrder, sRow, colBARCODE))
        lsSeq = Trim(GetText(spdROrder, sRow, colSAVESEQ))
        strExamDate = Trim(GetText(spdROrder, sRow, colEXAMDATE))
        

        If lsSeq < 1 Then
            Exit Sub
        End If

        If MsgBox(lsSeq & " 의 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If

              SQL = "DELETE FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
        'SQL = SQL & "   AND PID = '" & lsPid & "' " & vbCrLf
        SQL = SQL & "   AND SAVESEQ = " & lsSeq & vbCrLf
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & strExamDate & "' "
        
'        Res = SendQuery(gLocal, SQL)
'
'        If Res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
        
        If DBExec(AdoCn_Local, SQL) Then
            '-- 성공
        End If
                
        DeleteRow spdROrder, sRow, sRow
        spdRResult.MaxRows = 0
        'blnModify = True
        
    ElseIf KeyAscii = vbKeyReturn Then
        If spdROrder.ActiveCol = colBARCODE Then
            
            If GetSampleInfo(sRow, spdROrder) = -1 Then
                MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
            Else
                '-- 환자정보수정
                SQL = ""
                SQL = SQL & "UPDATE PATRESULT SET "
                SQL = SQL & "  BARCODE = '" & Trim(GetText(spdROrder, sRow, colBARCODE)) & "'" & vbCr
                SQL = SQL & " ,INOUT   = '" & Trim(GetText(spdROrder, sRow, colINOUT)) & "'" & vbCr
                SQL = SQL & " ,CHARTNO = '" & Trim(GetText(spdROrder, sRow, colCHARTNO)) & "'" & vbCr
                SQL = SQL & " ,PID     = '" & Trim(GetText(spdROrder, sRow, colPID)) & "'" & vbCr
                SQL = SQL & " ,PNAME   = '" & Trim(GetText(spdROrder, sRow, colPNAME)) & "'" & vbCr
                SQL = SQL & " ,PSEX    = '" & Trim(GetText(spdROrder, sRow, colPSEX)) & "'" & vbCr
                SQL = SQL & " ,PAGE    = '" & Trim(GetText(spdROrder, sRow, colPAGE)) & "'" & vbCr
                SQL = SQL & " ,PJUMIN  = '" & Trim(GetText(spdROrder, sRow, colPJUMIN)) & "'" & vbCr
                SQL = SQL & " WHERE EXAMDATE = '" & Trim(GetText(spdROrder, sRow, colEXAMDATE)) & "'" & vbCr
                SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdROrder, sRow, colSAVESEQ)) & vbCr
                SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.MACHCD & "' & vbCr"
                'SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdROrder, asRow1, colBARCODE)) & "' " & vbCr
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            
        ElseIf spdROrder.ActiveCol > colSTATE Then
            strTestNm = GetText(spdROrder, 0, sCol)
            strResult = GetText(spdROrder, sRow, sCol)
            
            For intRow = 1 To spdRResult.MaxRows
                If strTestNm = GetText(spdRResult, intRow, colRTESTNM) Then
                    strTestCd = GetText(spdRResult, intRow, colRTESTCD)
                    strIntBase = GetText(spdRResult, intRow, colRCHANNEL)
                
                    '소수점 처리, 결과판정
                    strResult = SetResult(strResult, strIntBase)
                    strJudge = SetJudge(strResult, strIntBase)
                                                        
                                                        
                    '-- 검사결과수정
                    SQL = ""
                    SQL = SQL & "UPDATE PATRESULT SET "
                    SQL = SQL & "  RESULT   = '" & strResult & "'" & vbCr
                    SQL = SQL & " ,REFJUDGE = '" & strJudge & "'" & vbCr
                    SQL = SQL & " WHERE EXAMDATE = '" & Trim(GetText(spdROrder, sRow, colEXAMDATE)) & "'" & vbCr
                    SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdROrder, sRow, colSAVESEQ)) & vbCr
                    SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCr
                    SQL = SQL & "   AND EXAMCODE = '" & strTestCd & "'" & vbCr
                    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                        Call SetText(spdROrder, strResult, sRow, sCol)
                        Call spdROrder_Click(sCol, sRow)
                    End If
                End If
            Next
        End If
    End If
    
End Sub

Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub

Private Sub comEQP_OnComm()
    Dim EVMsg       As String
    Dim ERMsg       As String
    Dim Ret         As Long
    Dim strDate     As String
    
    Select Case comEqp.CommEvent
        Case comEvReceive

            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If

            pBuffer = comEqp.Input
            
            dtpToday.Value = Now
            
            Call Serial_Protocol

            SetRawData "[Rx]" & pBuffer
            
        Case comEvSend
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
        
        Case comEvCTS
            EVMsg$ = "CTS 변경 감지"
        Case comEvDSR
            EVMsg$ = "DSR 변경 감지"
        Case comEvCD
            EVMsg$ = "CD 변경 감지"
        Case comEvRing
            EVMsg$ = "전화 벨이 울리는 중"
        Case comEvEOF
            EVMsg$ = "EOF 감지"

        '오류 메시지
        Case comBreak
            ERMsg$ = "중단 신호 수신"
        Case comCDTO
            ERMsg$ = "반송파 검출 시간 초과"
        Case comCTSTO
            ERMsg$ = "CTS 시간 초과"
        Case comDCB
            ERMsg$ = "DCB 검색 오류"
        Case comDSRTO
            ERMsg$ = "DSR 시간 초과"
        Case comFrame
            ERMsg$ = "프레이밍 오류"
        Case comOverrun
            ERMsg$ = "패리티 오류"
        Case comRxOver
            ERMsg$ = "수신 버퍼 초과"
        Case comRxParity
            ERMsg$ = "패리티 오류"
        Case comTxFull
            ERMsg$ = "전송 버퍼에 여유가 없음"
        Case Else
            ERMsg$ = "알 수 없는 오류 또는 이벤트"
    End Select

End Sub

Private Sub Form_Load()

On Error GoTo RST

    Me.Width = 20940
    Me.Height = 12585
    
    lblHospInfo.Caption = gHOSP.HOSPNM & "  " & gHOSP.MACHNM & "  " & gHOSP.USERNM & "[" & gHOSP.USERID & "]" '& "버전 " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = gHOSP.MACHNM
    
    Call CtlInitializing
    
    '-- Menu Set
    Call SetMenu
    
    '-- 검사코드
    Call GetTestList
    
    '-- 오더코드
    Call GetOrderMST

    '-- 검사명 보이기
    Call SetExamCode
    
    '-- 통신설정
    Call GetCommList

    If gComm.COMTYPE = "1" Then
        comEqp.CommPort = gComm.COMPORT
        comEqp.RTSEnable = gComm.RTSEnable
        comEqp.DTREnable = gComm.DTREnable
        comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT
    
        If comEqp.PortOpen = False Then
            comEqp.PortOpen = True
        End If
    
        If comEqp.PortOpen Then
            lblStatus.Caption = "COM" & comEqp.CommPort & " 포트에 연결 되었습니다"
            imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        Else
            lblStatus.Caption = "통신포트에 연결 되지 않았습니다"
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        End If
    ElseIf gComm.COMTYPE = "2" Then
        If gComm.TCPTYPE = "1" Then
            wSck.LocalPort = CInt(gComm.TCPPORT)
            wSck.Listen
        
            lblStatus.Caption = "TCP " & gComm.TCPPORT & " 포트에 연결 되었습니다"
            imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        Else
            wSck.Close
            wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)
        
            lblStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 포트에 연결 되었습니다"
            imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        End If
    Else
        imgPort.Visible = False
        imgSend.Visible = False
        imgReceive.Visible = False
        lblPort.Visible = False
        lblSend.Visible = False
        lblReceive.Visible = False
        
        lblStatus.Caption = "결과 경로: " & gComm.RSTPATH
    End If
    
    frame1.Visible = True
    frame1.ZOrder 0

    Call cmdSL_Click
    
    Exit Sub
    
RST:
    frame1.Visible = True
    frame1.ZOrder 0
    
    If Err.Number = "8002" Then
        If (MsgBox("포트 번호가 잘못되었습니다." & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            Resume Next
        Else
            
            End
        End If
    Else
        MsgBox Err.Number & vbNewLine & Err.Description
    End If
    
End Sub

'-- 검사마스터 조회
Public Sub GetCommList()
    Dim i As Integer
    Dim Ret As Integer
    
    If gComm.COMTYPE = "1" Then
        optComType(0).Value = True
        frameCom.Enabled = True
        frameTCP.Enabled = False
        frameFILE.Enabled = False
    ElseIf gComm.COMTYPE = "1" Then
        optComType(1).Value = True
        frameCom.Enabled = False
        frameTCP.Enabled = True
        frameFILE.Enabled = False
    Else
        optComType(2).Value = True
        frameCom.Enabled = False
        frameTCP.Enabled = False
        frameFILE.Enabled = True
    End If
    
    Ret = -1
    For i = 0 To cboPort.ListCount - 1
        If gComm.COMPORT = Trim(cboPort.List(i)) Then
            cboPort.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    
'    If Ret = -1 Then
'        cboPort.ListIndex = 1
'    End If
    
    Ret = -1
    For i = 0 To cboBaudrate.ListCount - 1
        If gComm.SPEED = Trim(cboBaudrate.List(i)) Then
            cboBaudrate.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        cboBaudrate.ListIndex = 4
    End If
    
    Ret = -1
    For i = 0 To cboDatabit.ListCount - 1
        If gComm.DATABIT = Trim(cboDatabit.List(i)) Then
            cboDatabit.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        cboBaudrate.ListIndex = 1
    End If

    Ret = -1
    For i = 0 To cboStartbit.ListCount - 1
        If gComm.STARTBIT = Trim(cboStartbit.List(i)) Then
            cboStartbit.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        cboStartbit.ListIndex = 0
    End If
    
    Ret = -1
    For i = 0 To cboStopbit.ListCount - 1
        If gComm.STOPBIT = Trim(cboStopbit.List(i)) Then
            cboStopbit.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        cboStopbit.ListIndex = 0
    End If
    
    Ret = -1
    For i = 0 To cboParity.ListCount - 1
        If gComm.Parity = Trim(cboParity.List(i)) Then
            cboParity.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        cboParity.ListIndex = 0
    End If
    
    '--------------------------------------------
    
    If gComm.TCPTYPE = "1" Then
        optTCPType(0).Value = True
    Else
        optTCPType(1).Value = True
    End If
    
    txtTCPIP.Text = gComm.TCPIP
    txtTCPPort.Text = gComm.TCPPORT
    txtOrdPath = gComm.ORDPATH
    txtRstPath = gComm.RSTPATH
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub
    
    '-- 인터페이스
    frame1.Width = Me.ScaleWidth - 150
    frame1.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150
    
    spdOrder.Width = Me.ScaleWidth - 300 'spdResult.Width - 400
    spdOrder.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500
    
    'spdResult.Left = spdOrder.Left + spdOrder.Width + 50
    'spdResult.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500

    '-- 결과조회
    frame2.Width = Me.ScaleWidth - 150
    frame2.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150
    
    spdROrder.Width = Me.ScaleWidth - 300 'spdRResult.Width - 500
    spdROrder.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500
    
    'spdRResult.Left = spdOrder.Left + spdROrder.Width + 50
    'spdRResult.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500

    '-- 검사설정
    frame3.Width = Me.ScaleWidth - 150
    frame3.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150
    
    spdTest.Width = Me.ScaleWidth - frameTestSet.Width - 600
    spdTest.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500
    
    frameTestSet.Left = spdTest.Left + spdTest.Width + 50
    frameTestSet.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500

    '-- 통신설정
    frame4.Width = Me.ScaleWidth - 150
    frame4.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150

End Sub





Private Sub fraInterface_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblWork.ForeColor = vbBlack
    lblSave.ForeColor = vbBlack
    lblClear.ForeColor = vbBlack
    shpW.BorderColor = &H808080
    shpS.BorderColor = &H808080
    shpC.BorderColor = &H808080
    
End Sub

Private Sub frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblComSave.ForeColor = vbBlack
    lblTcpSave.ForeColor = vbBlack
    
    shpCom.BorderColor = &H808080
    shpTcp.BorderColor = &H808080
    
    
End Sub

Private Sub frameTestSet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer

    For i = 0 To 3
        lblActionTest(i).ForeColor = vbBlack
        shpA(i).BorderColor = &H808080
    Next
    
End Sub

Private Sub imgDelete_Click()
    Dim Test_Property As Scripting.Dictionary
    Dim objTest_Property As clsCommon
    
    If Trim(txtEqpCD.Text) = "" Then
        MsgBox "검사항목을 먼저 선택하세요", vbCritical, Me.Caption
        Exit Sub
    End If
    
    If Trim(txtOChannel.Text) = "" Then
        MsgBox "검사항목을 먼저 선택하세요", vbCritical, Me.Caption
        Exit Sub
    End If
    
    Set Test_Property = New Scripting.Dictionary

    With Test_Property
        .Add "EQPCD", txtEqpCD.Text
        .Add "OCH", txtOChannel.Text
        .Add "TESTCD", txtTestCd.Text
    End With
    
    Set objTest_Property = New clsCommon
    
    With objTest_Property
        .SetAdoCn AdoCn_Local
        If .DelTestInfo(Test_Property) Then
            '-- 저장 오류
            Call GetTestList
        Else
            '-- 저장 오류
            Call GetTestList
        End If
    End With

End Sub

Private Sub imgSave_Click()
    Dim Test_Property As Scripting.Dictionary
    Dim objTest_Property As clsCommon
    
    If Trim(txtEqpCD.Text) = "" Then
        MsgBox "검사항목을 먼저 선택하세요", vbCritical, Me.Caption
        Exit Sub
    End If
    
    If Trim(txtOChannel.Text) = "" Then
        MsgBox "오더채널을 입력하세요", vbCritical, Me.Caption
        txtOChannel.SetFocus
        Exit Sub
    End If
    
    If Trim(txtRChannel.Text) = "" Then
        MsgBox "결과채널을 입력하세요", vbCritical, Me.Caption
        txtRChannel.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTestCd.Text) = "" Then
        MsgBox "검사코드를 입력하세요", vbCritical, Me.Caption
        txtTestCd.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTestNm.Text) = "" Then
        MsgBox "검사명을 입력하세요", vbCritical, Me.Caption
        txtTestNm.SetFocus
        Exit Sub
    End If
    
    
    Set Test_Property = New Scripting.Dictionary

    With Test_Property
        .Add "EQPCD", txtEqpCD.Text
        .Add "SEQ", txtSeq.Text
        .Add "OCH", txtOChannel.Text
        .Add "RCH", txtRChannel.Text
        .Add "TESTCD", txtTestCd.Text
        .Add "TESTNM", txtTestNm.Text
        .Add "ABBRNM", txtAbbrNm.Text
        .Add "RES", txtResSpec.Text
        .Add "REFL", txtRefLow.Text
        .Add "REFH", txtRefHigh.Text
        .Add "RSTTYPE", cboResultType.Text
        If optCutUse(0).Value = True Then
            .Add "CUTUSE", "Y"
        Else
            .Add "CUTUSE", "N"
        End If
        .Add "COLIN", txtCOLIn.Text
        .Add "COLCP", cboCOL.Text
        .Add "COLOUT", txtCOLOut.Text
        .Add "COHIN", txtCOHIn.Text
        .Add "COHCP", cboCOH.Text
        .Add "COHOUT", txtCOHOut.Text
        .Add "COMOUT", txtCOMOut.Text
    End With
    
    Set objTest_Property = New clsCommon
    
    With objTest_Property
        .SetAdoCn AdoCn_Local
        If .LetTestInfo(Test_Property) Then
            '-- 저장 오류
            Call GetTestList
        Else
            '-- 저장 오류
            Call GetTestList
        End If
    End With

End Sub



Public Sub CtlInitializing()
    Dim intComPortExist As Long
    Dim i As Integer
    
    frame1.Left = 50
    frame1.Top = 1650
    
    frame2.Left = 50
    frame2.Top = 1650
    
    frame3.Left = 50
    frame3.Top = 1650
    
    frame4.Left = 50
    frame4.Top = 1650
    
    dtpToday.Value = Now
    dtpFrom.Value = Now
    dtpTo.Value = Now
    
    
    '-- 인터페이스
    spdOrder.MaxRows = 0
    spdResult.MaxRows = 0
    
    '-- 장비결과
    spdROrder.MaxRows = 0
    spdRResult.MaxRows = 0
        
    '-- 검사코드 설정
    spdTest.MaxRows = 0
    
    cboCOL.AddItem "<"
    cboCOL.AddItem "<="
    cboCOL.ListIndex = 0
    
    cboCOH.AddItem ">"
    cboCOH.AddItem ">="
    cboCOH.ListIndex = 0
    
    cboResultType.AddItem "변함없음"
    cboResultType.AddItem "정량"
    cboResultType.AddItem "정성"
    cboResultType.AddItem "정량(정성)"
    cboResultType.AddItem "정성(정량)"
    cboResultType.ListIndex = 0
    
    txtEqpCD.Text = gHOSP.HOSPCD
    
    '-- 통신설정
    cboPort.AddItem ("1")
    cboPort.AddItem ("2")
    cboPort.AddItem ("3")
    cboPort.AddItem ("4")
    cboPort.AddItem ("5")
    cboPort.AddItem ("6")
    cboPort.AddItem ("7")
    cboPort.AddItem ("8")
    cboPort.AddItem ("9")
    cboPort.AddItem ("10")
    cboPort.AddItem ("11")
    cboPort.AddItem ("12")
    cboPort.AddItem ("13")
    cboPort.AddItem ("14")
    cboPort.AddItem ("15")
    cboPort.AddItem ("16")
    
    cboPort.Clear
    For i = 1 To 16
        intComPortExist = EnumSerPorts(i)
        If intComPortExist > 0 Then
            cboPort.AddItem Trim(str(i))
        End If
    Next
    
    cboBaudrate.AddItem ("150")
    cboBaudrate.AddItem ("300")
    cboBaudrate.AddItem ("600")
    cboBaudrate.AddItem ("1200")
    cboBaudrate.AddItem ("2400")
    cboBaudrate.AddItem ("4800")
    cboBaudrate.AddItem ("9600")
    cboBaudrate.AddItem ("14400")
    cboBaudrate.AddItem ("19200")
    cboBaudrate.AddItem ("115200")
    
    cboDatabit.AddItem ("7")
    cboDatabit.AddItem ("8")
    
    cboStartbit.AddItem ("1")
    cboStartbit.AddItem ("2")
    
    cboStopbit.AddItem ("1")
    cboStopbit.AddItem ("1.5")
    cboStopbit.AddItem ("2")
    
    cboParity.AddItem ("N")
    cboParity.AddItem ("E")
    cboParity.AddItem ("O")
    
    txtTCPIP.Text = ""
    txtTCPPort.Text = ""
    
    '==============================
    intPhase = 1
    intBufCnt = 0
    intFrameNo = 1
    intSndPhase = 0
    strState = ""
    blnIsETB = False
    '==============================
    
    If gHOSP.BARUSE = "Y" Then
        optBarSeq(0).Value = True
    Else
        optBarSeq(1).Value = True
    End If
    
    
    cboState.Clear
    cboState.AddItem "--전체--"
    cboState.AddItem "전송"
    cboState.AddItem "미전송"
    cboState.ListIndex = 0
    
    cboRstType.Clear
    cboRstType.AddItem "검사일자"
    cboRstType.AddItem "접수일자"
    cboRstType.ListIndex = 0
    
End Sub

Private Sub lblActionTest_Click(Index As Integer)
    Dim Test_Property As Scripting.Dictionary
    Dim objTest_Property As clsCommon
    
    If Index = 0 Then
        Call GetTestList
    
    ElseIf Index = 1 Then
        If Trim(txtEqpCD.Text) = "" Then
            MsgBox "검사항목을 먼저 선택하세요", vbCritical, Me.Caption
            Exit Sub
        End If
        
        If Trim(txtOChannel.Text) = "" Then
            MsgBox "검사항목을 먼저 선택하세요", vbCritical, Me.Caption
            Exit Sub
        End If
        
        If MsgBox(txtTestNm.Text & "를 삭제하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
             Exit Sub
        End If
        Set Test_Property = New Scripting.Dictionary
    
        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "OCH", txtOChannel.Text
            .Add "TESTCD", txtTestCd.Text
        End With
        
        Set objTest_Property = New clsCommon
        
        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .DelTestInfo(Test_Property) Then
                '-- 삭제 오류
                'Call GetTestList
            End If
        End With
        
        Call GetTestList
        
    ElseIf Index = 2 Then
        If Trim(txtEqpCD.Text) = "" Then
            MsgBox "검사항목을 먼저 선택하세요", vbCritical, Me.Caption
            Exit Sub
        End If
        
        If Trim(txtOChannel.Text) = "" Then
            MsgBox "오더채널을 입력하세요", vbCritical, Me.Caption
            txtOChannel.SetFocus
            Exit Sub
        End If
        
        If Trim(txtRChannel.Text) = "" Then
            MsgBox "결과채널을 입력하세요", vbCritical, Me.Caption
            txtRChannel.SetFocus
            Exit Sub
        End If
        
        If Trim(txtTestCd.Text) = "" Then
            MsgBox "검사코드를 입력하세요", vbCritical, Me.Caption
            txtTestCd.SetFocus
            Exit Sub
        End If
        
        If Trim(txtTestNm.Text) = "" Then
            MsgBox "검사명을 입력하세요", vbCritical, Me.Caption
            txtTestNm.SetFocus
            Exit Sub
        End If
        
        Set Test_Property = New Scripting.Dictionary
    
        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "SEQ", txtSeq.Text
            .Add "OCH", txtOChannel.Text
            .Add "RCH", txtRChannel.Text
            .Add "TESTCD", txtTestCd.Text
            .Add "TESTNM", txtTestNm.Text
            .Add "ABBRNM", txtAbbrNm.Text
            .Add "RES", txtResSpec.Text
            .Add "REFL", txtRefLow.Text
            .Add "REFH", txtRefHigh.Text
            .Add "RSTTYPE", cboResultType.Text
            If optCutUse(0).Value = True Then
                .Add "CUTUSE", "N"
            Else
                .Add "CUTUSE", "Y"
            End If
            .Add "COLIN", txtCOLIn.Text
            .Add "COLCP", cboCOL.Text
            .Add "COLOUT", txtCOLOut.Text
            .Add "COHIN", txtCOHIn.Text
            .Add "COHCP", cboCOH.Text
            .Add "COHOUT", txtCOHOut.Text
            .Add "COMOUT", txtCOMOut.Text
        End With
        
        Set objTest_Property = New clsCommon
        
        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .LetTestInfo(Test_Property) Then
                '-- 저장 오류
                'Call GetTestList
            End If
        End With
        
        Call GetTestList
        
    ElseIf Index = 3 Then
        If frameOrder.Visible = True Then
            frameOrder.Visible = False
        Else
            frameOrder.Visible = True
        End If
    End If
    
End Sub

Private Sub lblActionTest_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer

    For i = 0 To 2
        lblActionTest(i).ForeColor = vbBlack
        shpA(i).BorderColor = &H808080
    Next
    
    lblActionTest(Index).ForeColor = vbBlue
    shpA(Index).BorderColor = vbCyan


End Sub

Private Sub lblClear_Click()
    
    spdOrder.MaxRows = 0
    spdResult.MaxRows = 0

End Sub

Private Sub lblClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblWork.ForeColor = vbBlack
    lblSave.ForeColor = vbBlack
    lblClear.ForeColor = vbBlack
    shpW.BorderColor = &H808080
    shpS.BorderColor = &H808080
    shpC.BorderColor = &H808080
    
    lblClear.ForeColor = vbBlue
    shpC.BorderColor = vbCyan

End Sub

Private Sub lblComSave_Click()

    If optComType(0).Value = True Then
        Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\OKSOFT.ini")
    ElseIf optComType(1).Value = True Then
        Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\OKSOFT.ini")
    Else
        Call WritePrivateProfileString("COMM", "COMTYPE", "3", App.PATH & "\OKSOFT.ini")
    End If

    
    Call WritePrivateProfileString("COMM", "COMPORT", cboPort.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("COMM", "SPEED", cboBaudrate.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("COMM", "PARITY", cboParity.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("COMM", "DATABIT", cboDatabit.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("COMM", "STARTBIT", cboStartbit.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("COMM", "STOPBIT", cboStopbit.Text, App.PATH & "\OKSOFT.ini")
    
    GetSetup
    
    GetCommList

End Sub

Private Sub lblComSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblComSave.ForeColor = vbBlue
    shpCom.BorderColor = vbCyan

End Sub

Private Sub lblMenu_Click(Index As Integer)

    
    frame1.Visible = False
    frame2.Visible = False
    frame3.Visible = False
    'frame4.Visible = False
    fraInterface.Visible = False
    fraResult.Visible = False
    
    Select Case Index
        Case 0:
                frame1.Visible = True
                frame1.ZOrder 0
        
                fraInterface.Visible = True
        Case 1:
                frame2.Visible = True
                frame2.ZOrder 0
        
                fraResult.Visible = True
        Case 2:
                frame3.Visible = True
                frame3.ZOrder 0
    
                '-- 검사코드
                Call GetTestList
        
        Case 3:
                frame4.Visible = True
                frame4.ZOrder 0
    
                '-- 통신설정
                Call GetCommList
    
    End Select
    
    'vasPrint.ZOrder 0

End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer

    For i = 0 To 3
        lblMenu(i).ForeColor = vbBlack
        shpB(i).BorderColor = vbGreen
    Next
    
    lblMenu(Index).ForeColor = vbBlue
    shpB(Index).BorderColor = vbCyan

End Sub



Private Sub lblSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblWork.ForeColor = vbBlack
    lblSave.ForeColor = vbBlack
    lblClear.ForeColor = vbBlack
    shpW.BorderColor = &H808080
    shpS.BorderColor = &H808080
    shpC.BorderColor = &H808080
    
    lblSave.ForeColor = vbBlue
    shpS.BorderColor = vbCyan

End Sub

Private Sub lblTcpSave_Click()
    

    If optComType(0).Value = True Then
        Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\OKSOFT.ini")
    ElseIf optComType(1).Value = True Then
        Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\OKSOFT.ini")
    Else
        Call WritePrivateProfileString("COMM", "COMTYPE", "3", App.PATH & "\OKSOFT.ini")
    End If

    
    If optTCPType(0).Value = True Then
        Call WritePrivateProfileString("COMM", "TCPTYPE", "1", App.PATH & "\OKSOFT.ini")
    Else
        Call WritePrivateProfileString("COMM", "TCPTYPE", "2", App.PATH & "\OKSOFT.ini")
    End If
    
    Call WritePrivateProfileString("COMM", "TCPIP", txtTCPIP.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("COMM", "TCPPORT", txtTCPPort.Text, App.PATH & "\OKSOFT.ini")
    
    GetSetup
    
    GetCommList

End Sub

Private Sub lblTcpSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblTcpSave.ForeColor = vbBlue
    shpTcp.BorderColor = vbCyan

End Sub

Private Sub lblWork_Click()
    
    frmWorkList.Show vbModal
    
End Sub

Private Sub lblWork_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblWork.ForeColor = vbBlack
    lblSave.ForeColor = vbBlack
    lblClear.ForeColor = vbBlack
    shpW.BorderColor = &H808080
    shpS.BorderColor = &H808080
    shpC.BorderColor = &H808080
    
    lblWork.ForeColor = vbBlue
    shpW.BorderColor = vbCyan

End Sub

Private Sub optComType_Click(Index As Integer)
    
    If Index = 0 Then
        frameCom.Enabled = True
        frameTCP.Enabled = False
    Else
        frameCom.Enabled = False
        frameTCP.Enabled = True
    End If

End Sub

Private Sub optCutUse_Click(Index As Integer)
    If Index = 0 Then
        frameCutOff.Enabled = False
    Else
        frameCutOff.Enabled = True
    End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer

    For i = 0 To 3
        lblMenu(i).ForeColor = vbBlack
        shpB(i).BorderColor = vbGreen
    Next
    
    lblWork.ForeColor = vbBlack
    lblSave.ForeColor = vbBlack
    lblClear.ForeColor = vbBlack
    shpW.BorderColor = &H808080
    shpS.BorderColor = &H808080
    shpC.BorderColor = &H808080
    
    lblResult.ForeColor = vbBlack
    shpR.BorderColor = &H808080
    
    
End Sub



Private Sub spdOrder_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol As Integer
    
    '-- 정렬
    If Row = 0 Then
        '-- 정렬 추가
        
        Exit Sub
    End If
    
    '-- 환자정보표시
    
    '-- 결과표시
    If GetPatTRestResult(Row) = -1 Then
        '장비결과가 없을경우 검사명만 보여주기
        spdResult.MaxRows = 0
        With spdOrder
            For intCol = colSTATE + 1 To .MaxCols
                If GetText(spdOrder, Row, intCol) <> "" Then    '◆
                    spdResult.MaxRows = spdResult.MaxRows + 1
                    Call SetText(spdResult, GetText(spdOrder, 0, intCol), spdResult.MaxRows, colRTESTNM)
                    spdResult.RowHeight(-1) = 12
                End If
            Next
        End With
    End If
        
End Sub

'인터페이스 환자선택시 우측에 검사항목/결과보여주기
Private Function GetPatTRestResult(ByVal asRow As Integer) As Integer
    Dim strBarno As String
    Dim intSeq   As String
    Dim strExamDate As String
    Dim intRow   As Integer
    
On Error GoTo RST

    GetPatTRestResult = -1
    intRow = 0
    
    intSeq = GetText(spdOrder, asRow, colSAVESEQ)
    strExamDate = Mid(GetText(spdOrder, asRow, colEXAMDATE), 1, 8)
    
    If intSeq = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT SEQNO, EXAMNAME, RESULT" & vbCr
    SQL = SQL & "  FROM PATRESULT " & vbCr
    SQL = SQL & " WHERE SAVESEQ = " & intSeq & vbCr
    SQL = SQL & "   AND EXAMDATE = '" & strExamDate & "'" & vbCr
'    SQL = SQL & " ORDER BY SEQNO "
    
    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        With frmMain.spdResult
            .MaxRows = 0
            .MaxRows = AdoRs_Local.RecordCount
            Do Until AdoRs_Local.EOF
                intRow = intRow + 1
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("EXAMNAME").Value & "", intRow, colRTESTNM)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("RESULT").Value & "", intRow, colRLISRESULT)
                AdoRs_Local.MoveNext
            Loop
        End With
        GetPatTRestResult = 1
    End If
    
    AdoRs_Local.Close
    
Exit Function

RST:
    GetPatTRestResult = -1

End Function

'인터페이스 환자선택시 우측에 검사항목/결과보여주기
Public Function GetPatTRestResult_Search(ByVal asRow As Integer) As Integer
    Dim strBarno As String
    Dim intSeq   As String
    Dim strExamDate As String
    Dim intRow   As Integer
    
On Error GoTo RST

    GetPatTRestResult_Search = -1
    intRow = 0
    
    intSeq = GetText(spdROrder, asRow, colSAVESEQ)
    strExamDate = Mid(GetText(spdROrder, asRow, colEXAMDATE), 1, 8)
    
    If intSeq = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT SEQNO,EXAMCODE,EQUIPCODE,EXAMNAME,EQUIPRESULT,RESULT" & vbCr
    SQL = SQL & "  FROM PATRESULT " & vbCr
    SQL = SQL & " WHERE SAVESEQ = " & intSeq & vbCr
    SQL = SQL & "   AND EXAMDATE = '" & strExamDate & "'" & vbCr
'    SQL = SQL & " ORDER BY SEQNO "
    
    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        With frmMain.spdRResult
            .MaxRows = 0
            .MaxRows = AdoRs_Local.RecordCount
            Do Until AdoRs_Local.EOF
                intRow = intRow + 1
                Call SetText(frmMain.spdRResult, AdoRs_Local.Fields("SEQNO").Value & "", intRow, colRSEQNO)
                Call SetText(frmMain.spdRResult, AdoRs_Local.Fields("EXAMCODE").Value & "", intRow, colRTESTCD)
                Call SetText(frmMain.spdRResult, AdoRs_Local.Fields("EXAMNAME").Value & "", intRow, colRTESTNM)
                Call SetText(frmMain.spdRResult, AdoRs_Local.Fields("EQUIPCODE").Value & "", intRow, colRCHANNEL)
                Call SetText(frmMain.spdRResult, AdoRs_Local.Fields("EQUIPRESULT").Value & "", intRow, colRMACHRESULT)
                Call SetText(frmMain.spdRResult, AdoRs_Local.Fields("RESULT").Value & "", intRow, colRLISRESULT)
                AdoRs_Local.MoveNext
            Loop
        End With
        GetPatTRestResult_Search = 1
    End If
    
    AdoRs_Local.Close
    
Exit Function

RST:
    GetPatTRestResult_Search = -1

End Function


Private Sub spdOrdMst_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    
        
    If KeyAscii = vbKeyReturn Then
        '-- Delete
              SQL = ""
        SQL = SQL & "DELETE FROM ORDMASTER "
        
        Call DBExec(AdoCn_Local, SQL)
        
        'Insert
        For intRow = 1 To spdOrdMst.MaxRows
                  SQL = ""
            SQL = SQL & "INSERT INTO ORDMASTER (ORDERCODE,ORDERNAME) VALUES ("
            SQL = SQL & "'" & GetText(spdOrdMst, intRow, 1) & "','')"
            
            Call DBExec(AdoCn_Local, SQL)
        Next
    End If
    
End Sub

Private Sub spdTest_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        Exit Sub
    End If
    
    With spdTest
        txtEqpCD.Text = GetText(spdTest, Row, colLMACHCODE)
        txtSeq.Text = GetText(spdTest, Row, colLSEQNO)
        txtTestCd.Text = GetText(spdTest, Row, colLTESTCD)
        txtOChannel.Text = GetText(spdTest, Row, colLOCHANNEL)
        txtRChannel.Text = GetText(spdTest, Row, colLRCHANNEL)
        txtTestNm.Text = GetText(spdTest, Row, colLTESTNM)
        txtAbbrNm.Text = GetText(spdTest, Row, colLABBRNM)
        txtResSpec.Text = GetText(spdTest, Row, colLRESSPEC)
        txtRefLow.Text = GetText(spdTest, Row, colLLOW)
        txtRefHigh.Text = GetText(spdTest, Row, colLHIGH)
        cboResultType.Text = GetText(spdTest, Row, colLRSTTYPE)
        If GetText(spdTest, Row, colLCUTUSE) = "1" Then
            optCutUse(1).Value = True
        Else
            optCutUse(0).Value = True
        End If
        txtCOLIn.Text = GetText(spdTest, Row, colLCOLIN)
        cboCOL.Text = GetText(spdTest, Row, colLCOLCOMP)
        txtCOLOut = GetText(spdTest, Row, colLCOLOUT)
        txtCOHIn.Text = GetText(spdTest, Row, colLCOHIN)
        cboCOH.Text = GetText(spdTest, Row, colLCOHCOMP)
        txtCOHOut = GetText(spdTest, Row, colLCOHOUT)
        txtCOMOut = GetText(spdTest, Row, colLCOMOUT)
    End With
End Sub

Private Sub wsck_ConnectionRequest(ByVal requestID As Long)

    If wSck.State <> sckClosed Then
        wSck.Close

        wSck.Accept requestID
        lblStatus.Caption = "장비에 접속되었습니다."
    End If

End Sub

Private Sub wsck_DataArrival(ByVal bytesTotal As Long)
    Dim strText As String
    Dim strTmp As String
    
    Dim strLastSeq  As String
    Dim strRcvSign  As String
    Dim strSendAck  As String
    Dim strRcvCnt   As String
    
    Dim strNS       As String
    Dim strNE       As String
    Dim intNS       As Integer
    Dim intNE       As Integer
    
    Dim strSendData  As String
    Dim varBuffers   As Variant
    Dim i As Integer
    Dim lngBufLen As Long
    Dim BufChar     As String
    
    wSck.GetData strText

    pBuffer = strText
    
    dtpToday.Value = Now
    
    Call TCP_Protocol
    
    SetRawData "[Rx]" & pBuffer
    
    
End Sub



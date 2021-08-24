VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00F8E4D8&
   Caption         =   "OK SOFT"
   ClientHeight    =   12915
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   20910
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   15315
   ScaleWidth      =   28560
   StartUpPosition =   1  '소유자 가운데
   WindowState     =   2  '최대화
   Begin VB.Frame frame1 
      BackColor       =   &H00F8E4D8&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9645
      Left            =   50
      TabIndex        =   4
      Top             =   1650
      Width           =   20685
      Begin VB.CheckBox chkTest 
         BackColor       =   &H00F8E4D8&
         Caption         =   "선검사"
         Height          =   315
         Left            =   120
         TabIndex        =   209
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdWorkAll 
         BackColor       =   &H00C0FFC0&
         Caption         =   "일괄등록"
         Height          =   375
         Left            =   1200
         Style           =   1  '그래픽
         TabIndex        =   208
         Top             =   210
         Width           =   2805
      End
      Begin FPSpread.vaSpread spdWork 
         Height          =   9375
         Left            =   60
         TabIndex        =   207
         Top             =   570
         Width           =   3945
         _Version        =   393216
         _ExtentX        =   6959
         _ExtentY        =   16536
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   21
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
         MaxCols         =   21
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14548991
         SpreadDesigner  =   "frmMain.frx":0E42
         UserResize      =   2
      End
      Begin VB.Frame fraDUREADER720 
         BackColor       =   &H00FFFFFF&
         Caption         =   "유린 마이크로 결과"
         Height          =   2295
         Left            =   12360
         TabIndex        =   176
         Top             =   2550
         Visible         =   0   'False
         Width           =   2955
         Begin VB.ComboBox cboBacteria 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1020
            TabIndex        =   180
            Text            =   "Combo3"
            Top             =   1740
            Width           =   1695
         End
         Begin VB.ComboBox cboEpCell 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1020
            TabIndex        =   179
            Text            =   "Combo3"
            Top             =   1275
            Width           =   1695
         End
         Begin VB.ComboBox cboRbcM 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1020
            TabIndex        =   178
            Text            =   "Combo3"
            Top             =   825
            Width           =   1695
         End
         Begin VB.ComboBox cboWbcM 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "frmMain.frx":542E
            Left            =   1020
            List            =   "frmMain.frx":543E
            TabIndex        =   177
            Text            =   "Combo3"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "Bacteria"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   47
            Left            =   120
            TabIndex        =   184
            Top             =   1800
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "E.P Cell"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   46
            Left            =   120
            TabIndex        =   183
            Top             =   1335
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "RBC(M)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   45
            Left            =   120
            TabIndex        =   182
            Top             =   885
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "WBC(M)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   44
            Left            =   120
            TabIndex        =   181
            Top             =   420
            Width           =   720
         End
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   570
         TabIndex        =   80
         Top             =   270
         Visible         =   0   'False
         Width           =   195
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
         Visible         =   0   'False
         Width           =   435
      End
      Begin FPSpread.vaSpread spdResult 
         Height          =   9360
         Left            =   17370
         TabIndex        =   5
         Top             =   180
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
         SpreadDesigner  =   "frmMain.frx":544E
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdOrder 
         Height          =   9375
         Left            =   60
         TabIndex        =   206
         Top             =   180
         Width           =   17235
         _Version        =   393216
         _ExtentX        =   30401
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
         SpreadDesigner  =   "frmMain.frx":5EB4
         UserResize      =   2
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   6510
         TabIndex        =   78
         Top             =   -60
         Width           =   14145
         Begin VB.TextBox txtSeqNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   9630
            TabIndex        =   193
            Text            =   "0"
            Top             =   150
            Width           =   675
         End
         Begin VB.Frame Frame10 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            ForeColor       =   &H00FFFFFF&
            Height          =   585
            Left            =   3120
            TabIndex        =   185
            Top             =   0
            Width           =   4425
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
            Begin VB.Label lblWork 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H80000005&
               BackStyle       =   0  '투명
               Caption         =   "워크조회"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   210
               TabIndex        =   188
               Top             =   240
               Width           =   1125
            End
            Begin VB.Label lblSave 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H80000005&
               BackStyle       =   0  '투명
               Caption         =   "선택저장"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1650
               TabIndex        =   187
               Top             =   240
               Width           =   1125
            End
            Begin VB.Shape shpS 
               BackColor       =   &H00808080&
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               FillColor       =   &H00C0FFC0&
               Height          =   375
               Left            =   1530
               Top             =   150
               Width           =   1365
            End
            Begin VB.Label lblClear 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H80000005&
               BackStyle       =   0  '투명
               Caption         =   "화면정리"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3090
               TabIndex        =   186
               Top             =   240
               Width           =   1125
            End
            Begin VB.Shape shpC 
               BackColor       =   &H00808080&
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               FillColor       =   &H00C0FFC0&
               Height          =   375
               Left            =   2970
               Top             =   150
               Width           =   1365
            End
         End
         Begin VB.CommandButton cmdWork 
            BackColor       =   &H00C0FFFF&
            Caption         =   "워크조회"
            Height          =   405
            Left            =   3120
            Style           =   1  '그래픽
            TabIndex        =   189
            Top             =   120
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.TextBox txtBarcode 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7650
            TabIndex        =   172
            Text            =   "1234567890"
            Top             =   150
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker dtpFrDt 
            Height          =   315
            Left            =   180
            TabIndex        =   173
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
            Format          =   119275521
            CurrentDate     =   40457
         End
         Begin MSComCtl2.DTPicker dtpToDt 
            Height          =   315
            Left            =   1740
            TabIndex        =   174
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
            Format          =   119275521
            CurrentDate     =   40457
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "~"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   43
            Left            =   1530
            TabIndex        =   175
            Top             =   240
            Width           =   150
         End
      End
      Begin VB.Frame fraResult 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   6510
         TabIndex        =   91
         Top             =   -60
         Visible         =   0   'False
         Width           =   14145
         Begin VB.ComboBox cboRstType 
            Appearance      =   0  '평면
            Height          =   300
            ItemData        =   "frmMain.frx":A408
            Left            =   420
            List            =   "frmMain.frx":A40A
            TabIndex        =   114
            Top             =   180
            Width           =   1245
         End
         Begin VB.ComboBox cboState 
            Height          =   300
            ItemData        =   "frmMain.frx":A40C
            Left            =   4710
            List            =   "frmMain.frx":A40E
            TabIndex        =   113
            Top             =   180
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   315
            Left            =   1770
            TabIndex        =   93
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
            Format          =   119275521
            CurrentDate     =   40457
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   315
            Left            =   3330
            TabIndex        =   94
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
            Format          =   119275521
            CurrentDate     =   40457
         End
         Begin VB.Shape shpRX 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   9180
            Top             =   150
            Width           =   1365
         End
         Begin VB.Label lblRExcel 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "엑셀출력"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   9300
            TabIndex        =   205
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label lblRClear 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "화면정리"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7800
            TabIndex        =   149
            Top             =   240
            Width           =   1125
         End
         Begin VB.Shape shpRC 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   7680
            Top             =   150
            Width           =   1365
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "~"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   26
            Left            =   3120
            TabIndex        =   95
            Top             =   240
            Width           =   150
         End
         Begin VB.Image imgGbn 
            Height          =   225
            Left            =   180
            Picture         =   "frmMain.frx":A410
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
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6300
            TabIndex        =   92
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
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
         Left            =   4770
         TabIndex        =   34
         Top             =   150
         Width           =   1275
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
         Height          =   255
         Index           =   2
         Left            =   3300
         TabIndex        =   27
         Top             =   150
         Width           =   1275
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
         Height          =   255
         Index           =   1
         Left            =   1830
         TabIndex        =   26
         Top             =   150
         Width           =   1275
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
         BackColor       =   &H00C0FFFF&
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
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   25
         Top             =   150
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   28560
      TabIndex        =   0
      Top             =   0
      Width           =   28560
      Begin VB.Frame fraCommTest 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   15810
         TabIndex        =   110
         Top             =   60
         Visible         =   0   'False
         Width           =   4935
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clr"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   60
            TabIndex        =   192
            Top             =   630
            Width           =   375
         End
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   90
            TabIndex        =   112
            Top             =   150
            Width           =   375
         End
         Begin VB.TextBox txtRcv 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   111
            Top             =   120
            Width           =   4425
         End
      End
      Begin VB.Frame fraRS232 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   12630
         TabIndex        =   87
         Top             =   60
         Width           =   2985
         Begin VB.Label lblRcv 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "수신"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2010
            TabIndex        =   90
            Top             =   210
            Width           =   420
         End
         Begin VB.Label lblSend 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "송신"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1125
            TabIndex        =   89
            Top             =   210
            Width           =   420
         End
         Begin VB.Label lblPort 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "포트"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   150
            TabIndex        =   88
            Top             =   210
            Width           =   360
         End
         Begin VB.Image imgReceive 
            Height          =   240
            Left            =   2550
            Picture         =   "frmMain.frx":A7FA
            Top             =   180
            Width           =   240
         End
         Begin VB.Image imgSend 
            Height          =   240
            Left            =   1635
            Picture         =   "frmMain.frx":AD84
            Top             =   180
            Width           =   240
         End
         Begin VB.Image imgPort 
            Height          =   240
            Left            =   690
            Picture         =   "frmMain.frx":B30E
            Top             =   180
            Width           =   240
         End
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   10020
         TabIndex        =   108
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
         Format          =   119275520
         CurrentDate     =   40457
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   6090
         Top             =   -30
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         RTSEnable       =   -1  'True
         EOFEnable       =   -1  'True
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6810
         Top             =   30
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wSck 
         Left            =   7500
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label lblCommStatus 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Com"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   12840
         TabIndex        =   171
         Top             =   840
         Width           =   450
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   27
         Left            =   9150
         TabIndex        =   109
         Top             =   630
         Width           =   720
      End
      Begin VB.Image Image7 
         Height          =   225
         Left            =   8880
         Picture         =   "frmMain.frx":B898
         Top             =   600
         Width           =   150
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   12840
         TabIndex        =   2
         Top             =   600
         Width           =   2805
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
         Picture         =   "frmMain.frx":BC82
         Top             =   0
         Width           =   12900
      End
   End
   Begin VB.Frame frame4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9645
      Left            =   930
      TabIndex        =   37
      Top             =   2370
      Visible         =   0   'False
      Width           =   20685
      Begin VB.CommandButton cmdIF 
         Caption         =   "IF 설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11970
         TabIndex        =   120
         Top             =   8280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdConfig 
         BackColor       =   &H00FFFFFF&
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
         Height          =   1245
         Left            =   11910
         Style           =   1  '그래픽
         TabIndex        =   117
         Top             =   990
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
         TabIndex        =   59
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
         TabIndex        =   58
         Top             =   450
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.Frame frameTCP 
         BackColor       =   &H00FFFFFF&
         Caption         =   " TCP-IP 설정 "
         Height          =   7935
         Left            =   6480
         TabIndex        =   52
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
            TabIndex        =   62
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
            TabIndex        =   61
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
            TabIndex        =   57
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
            TabIndex        =   56
            Top             =   930
            Width           =   2445
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   7
            Left            =   840
            Picture         =   "frmMain.frx":D3C5
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
            TabIndex        =   60
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
            TabIndex        =   55
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
            TabIndex        =   54
            Top             =   1395
            Width           =   375
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   15
            Left            =   840
            Picture         =   "frmMain.frx":D7AF
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
            TabIndex        =   53
            Top             =   990
            Width           =   180
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   10
            Left            =   840
            Picture         =   "frmMain.frx":DB99
            Top             =   960
            Width           =   150
         End
      End
      Begin VB.Frame frameCom 
         BackColor       =   &H00FFFFFF&
         Caption         =   " RS-232 설정 "
         Height          =   7935
         Left            =   420
         TabIndex        =   38
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
            ItemData        =   "frmMain.frx":DF83
            Left            =   2190
            List            =   "frmMain.frx":DF85
            TabIndex        =   51
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
            ItemData        =   "frmMain.frx":DF87
            Left            =   2190
            List            =   "frmMain.frx":DF89
            TabIndex        =   50
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
            ItemData        =   "frmMain.frx":DF8B
            Left            =   2190
            List            =   "frmMain.frx":DF8D
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            ItemData        =   "frmMain.frx":DF8F
            Left            =   2190
            List            =   "frmMain.frx":DF91
            TabIndex        =   46
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
            TabIndex        =   45
            Top             =   1290
            Width           =   645
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   23
            Left            =   840
            Picture         =   "frmMain.frx":DF93
            Top             =   1260
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   22
            Left            =   840
            Picture         =   "frmMain.frx":E37D
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
            TabIndex        =   44
            Top             =   480
            Width           =   780
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   21
            Left            =   840
            Picture         =   "frmMain.frx":E767
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
            TabIndex        =   43
            Top             =   885
            Width           =   855
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   20
            Left            =   840
            Picture         =   "frmMain.frx":EB51
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
            TabIndex        =   42
            Top             =   1725
            Width           =   705
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   19
            Left            =   840
            Picture         =   "frmMain.frx":EF3B
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
            TabIndex        =   41
            Top             =   2130
            Width           =   705
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   18
            Left            =   840
            Picture         =   "frmMain.frx":F325
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
            TabIndex        =   40
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
            TabIndex        =   39
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
   Begin VB.Frame frame2 
      BackColor       =   &H00F8E4D8&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9645
      Left            =   780
      TabIndex        =   82
      Top             =   3240
      Visible         =   0   'False
      Width           =   20685
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
         TabIndex        =   116
         Top             =   210
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CheckBox chkRAll 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   115
         Top             =   240
         Width           =   195
      End
      Begin FPSpread.vaSpread spdRResult 
         Height          =   9360
         Left            =   13620
         TabIndex        =   86
         Top             =   180
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
         SpreadDesigner  =   "frmMain.frx":F70F
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdROrder 
         Height          =   9375
         Left            =   60
         TabIndex        =   85
         Top             =   180
         Width           =   13485
         _Version        =   393216
         _ExtentX        =   23786
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
         SpreadDesigner  =   "frmMain.frx":100EF
         UserResize      =   2
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
         TabIndex        =   84
         Top             =   210
         Width           =   435
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   570
         TabIndex        =   83
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Frame frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9645
      Left            =   1230
      TabIndex        =   6
      Top             =   1950
      Visible         =   0   'False
      Width           =   20685
      Begin VB.Frame frameTestSet 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9315
         Left            =   14730
         TabIndex        =   8
         Top             =   180
         Width           =   5835
         Begin VB.CheckBox chkResSpec 
            BackColor       =   &H00FFFFFF&
            Caption         =   "사용유무"
            Height          =   390
            Left            =   4050
            TabIndex        =   152
            Top             =   3540
            Width           =   1185
         End
         Begin VB.CommandButton cmdQCMaster 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "Biorad QC 설정"
            Height          =   345
            Left            =   3870
            Style           =   1  '그래픽
            TabIndex        =   130
            Top             =   4830
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.TextBox txtAnalyte 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            TabIndex        =   128
            Top             =   4860
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.Frame frameOrder 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   210
            TabIndex        =   75
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
               TabIndex        =   79
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
               TabIndex        =   76
               Top             =   210
               Width           =   285
            End
            Begin FPSpread.vaSpread spdOrdMst 
               Height          =   1920
               Left            =   90
               TabIndex        =   77
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
               SpreadDesigner  =   "frmMain.frx":1465E
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
            ItemData        =   "frmMain.frx":14BBB
            Left            =   1650
            List            =   "frmMain.frx":14BBD
            TabIndex        =   31
            Top             =   5220
            Visible         =   0   'False
            Width           =   1575
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
            Left            =   2280
            TabIndex        =   14
            Top             =   4020
            Width           =   1485
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
            Left            =   2280
            TabIndex        =   13
            Top             =   4440
            Width           =   1485
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
         Begin VB.Frame frameCut 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1440
            TabIndex        =   28
            Top             =   5460
            Visible         =   0   'False
            Width           =   2565
            Begin VB.OptionButton optCutUse 
               BackColor       =   &H00FFFFFF&
               Caption         =   "사용"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   1320
               TabIndex        =   30
               Top             =   180
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.OptionButton optCutUse 
               BackColor       =   &H00FFFFFF&
               Caption         =   "미사용"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   210
               TabIndex        =   29
               Top             =   180
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   1125
            End
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   41
            Left            =   1680
            TabIndex        =   151
            Top             =   4530
            Width           =   375
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   40
            Left            =   1680
            TabIndex        =   150
            Top             =   4110
            Width           =   360
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   24
            Left            =   330
            Picture         =   "frmMain.frx":14BBF
            Top             =   4890
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "QC 채널"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   600
            TabIndex        =   129
            Top             =   4920
            Visible         =   0   'False
            Width           =   690
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   15
            Left            =   600
            TabIndex        =   74
            Top             =   933
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "결과형식"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   22
            Left            =   600
            TabIndex        =   73
            Top             =   5310
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   14
            Left            =   330
            Picture         =   "frmMain.frx":14FA9
            Top             =   5280
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "CutOff"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   9
            Left            =   600
            TabIndex        =   72
            Top             =   5700
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   1
            Left            =   330
            Picture         =   "frmMain.frx":15393
            Top             =   5670
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "결과채널"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   19
            Left            =   600
            TabIndex        =   71
            Top             =   1839
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   11
            Left            =   330
            Picture         =   "frmMain.frx":1577D
            Top             =   1809
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   0
            Left            =   330
            Picture         =   "frmMain.frx":15B67
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   8
            Left            =   600
            TabIndex        =   70
            Top             =   480
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   2
            Left            =   330
            Picture         =   "frmMain.frx":15F51
            Top             =   1356
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "오더채널"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   10
            Left            =   600
            TabIndex        =   69
            Top             =   1386
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   3
            Left            =   330
            Picture         =   "frmMain.frx":1633B
            Top             =   2262
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   11
            Left            =   600
            TabIndex        =   68
            Top             =   2292
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   4
            Left            =   330
            Picture         =   "frmMain.frx":16725
            Top             =   2715
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   12
            Left            =   600
            TabIndex        =   67
            Top             =   2745
            Width           =   540
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   5
            Left            =   330
            Picture         =   "frmMain.frx":16B0F
            Top             =   3168
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   13
            Left            =   600
            TabIndex        =   66
            Top             =   3198
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   6
            Left            =   330
            Picture         =   "frmMain.frx":16EF9
            Top             =   3621
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   14
            Left            =   600
            TabIndex        =   65
            Top             =   3651
            Width           =   540
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   8
            Left            =   330
            Picture         =   "frmMain.frx":172E3
            Top             =   4074
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "참고치"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   16
            Left            =   600
            TabIndex        =   64
            Top             =   4104
            Width           =   540
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   16
            Left            =   330
            Picture         =   "frmMain.frx":176CD
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
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   3
            Left            =   4080
            TabIndex        =   63
            Top             =   8640
            Width           =   1125
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
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   4080
            TabIndex        =   36
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
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   2700
            TabIndex        =   35
            Top             =   7230
            Width           =   1125
         End
         Begin VB.Label lblActionTest 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Refresh"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   2670
            TabIndex        =   33
            Top             =   8640
            Visible         =   0   'False
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
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "ex)10.00"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   23
            Left            =   3390
            TabIndex        =   32
            Top             =   5280
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "ex)10.00"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   3930
            TabIndex        =   22
            Top             =   3270
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
         MaxCols         =   28
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmMain.frx":17AB7
      End
   End
   Begin VB.Frame FraHidden 
      Caption         =   "HIDDEN CONTROL"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8805
      Left            =   11880
      TabIndex        =   81
      Top             =   1800
      Visible         =   0   'False
      Width           =   12255
      Begin VB.Frame Frame12 
         Caption         =   "Frame12"
         Height          =   885
         Left            =   5010
         TabIndex        =   201
         Top             =   1770
         Width           =   4245
         Begin VB.CommandButton cmdTVSave 
            Caption         =   "저장"
            Height          =   345
            Left            =   3690
            TabIndex        =   203
            Top             =   360
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.TextBox txtTV 
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
            Left            =   2220
            TabIndex        =   202
            Top             =   390
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "total volume(L)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   42
            Left            =   690
            TabIndex        =   204
            Top             =   465
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   31
            Left            =   420
            Picture         =   "frmMain.frx":189C4
            Top             =   435
            Visible         =   0   'False
            Width           =   150
         End
      End
      Begin VB.CommandButton cmdInit 
         Caption         =   "초기화"
         Height          =   375
         Left            =   7710
         TabIndex        =   200
         Top             =   420
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.TextBox txtNum 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6540
         TabIndex        =   199
         Text            =   "60"
         Top             =   450
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame11"
         Height          =   1935
         Left            =   5370
         TabIndex        =   194
         Top             =   4440
         Width           =   3825
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   900
            TabIndex        =   197
            Text            =   "Text1"
            Top             =   360
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   315
            Left            =   2250
            TabIndex        =   196
            Top             =   330
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   195
            Text            =   "Text2"
            Top             =   360
            Visible         =   0   'False
            Width           =   765
         End
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   1005
            Left            =   150
            TabIndex        =   198
            Top             =   720
            Visible         =   0   'False
            Width           =   3315
            _Version        =   393216
            _ExtentX        =   5847
            _ExtentY        =   1773
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
            SpreadDesigner  =   "frmMain.frx":18DAE
         End
      End
      Begin VB.Timer Timer1 
         Left            =   3030
         Top             =   240
      End
      Begin VB.Timer tmrFlipFlop 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3120
         Top             =   720
      End
      Begin VB.CommandButton cmdEnd 
         Caption         =   "종료"
         Height          =   315
         Left            =   6750
         TabIndex        =   170
         Top             =   1140
         Width           =   795
      End
      Begin VB.Timer tmrComm 
         Enabled         =   0   'False
         Left            =   2670
         Top             =   750
      End
      Begin VB.Frame frameCutOff 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   5070
         TabIndex        =   159
         Top             =   2850
         Visible         =   0   'False
         Width           =   5175
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
            ItemData        =   "frmMain.frx":18FE7
            Left            =   2730
            List            =   "frmMain.frx":18FE9
            TabIndex        =   166
            Top             =   300
            Width           =   735
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
            TabIndex        =   165
            Top             =   300
            Width           =   1185
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
            TabIndex        =   164
            Top             =   300
            Width           =   1545
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
            TabIndex        =   163
            Top             =   660
            Width           =   1545
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
            ItemData        =   "frmMain.frx":18FEB
            Left            =   2730
            List            =   "frmMain.frx":18FED
            TabIndex        =   162
            Top             =   1020
            Width           =   735
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
            TabIndex        =   161
            Top             =   1020
            Width           =   1185
         End
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
            TabIndex        =   160
            Top             =   1020
            Width           =   1545
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   12
            Left            =   210
            Picture         =   "frmMain.frx":18FEF
            Top             =   360
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "CutOff (L)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   20
            Left            =   480
            TabIndex        =   169
            Top             =   390
            Width           =   825
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "CutOff (M)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   17
            Left            =   480
            TabIndex        =   168
            Top             =   750
            Width           =   885
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "CutOff (H)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   21
            Left            =   480
            TabIndex        =   167
            Top             =   1110
            Width           =   840
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   9
            Left            =   210
            Picture         =   "frmMain.frx":193D9
            Top             =   720
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   13
            Left            =   210
            Picture         =   "frmMain.frx":197C3
            Top             =   1080
            Width           =   150
         End
      End
      Begin VB.Timer TimerVESCUBE 
         Left            =   780
         Top             =   360
      End
      Begin VB.Timer Timer2120 
         Enabled         =   0   'False
         Left            =   360
         Top             =   330
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   330
         TabIndex        =   131
         Top             =   6480
         Visible         =   0   'False
         Width           =   5445
         Begin VB.TextBox txtInstrument 
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
            Left            =   1230
            TabIndex        =   141
            Top             =   1020
            Width           =   1185
         End
         Begin VB.TextBox txtLab 
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
            Left            =   1230
            TabIndex        =   140
            Top             =   300
            Width           =   1185
         End
         Begin VB.TextBox txtUnit 
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
            Left            =   1230
            TabIndex        =   139
            Top             =   1380
            Width           =   1185
         End
         Begin VB.TextBox txtReagent 
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
            Left            =   3690
            TabIndex        =   138
            Top             =   1020
            Width           =   1185
         End
         Begin VB.TextBox txtMethod 
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
            Left            =   3690
            TabIndex        =   137
            Top             =   660
            Width           =   1185
         End
         Begin VB.TextBox txtLot 
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
            Left            =   3690
            TabIndex        =   136
            Top             =   300
            Width           =   1185
         End
         Begin VB.TextBox txtTemp 
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
            Left            =   3690
            TabIndex        =   135
            Top             =   1380
            Width           =   1185
         End
         Begin VB.CommandButton cmdLabFind 
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2430
            TabIndex        =   134
            Top             =   300
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton Command3 
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5040
            TabIndex        =   133
            Top             =   300
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton cmdAnalyteFind 
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2430
            TabIndex        =   132
            Top             =   630
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   17
            Left            =   210
            Picture         =   "frmMain.frx":19BAD
            Top             =   1080
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "기기"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   480
            TabIndex        =   148
            Top             =   1110
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Lab"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   34
            Left            =   480
            TabIndex        =   147
            Top             =   390
            Width           =   315
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   25
            Left            =   210
            Picture         =   "frmMain.frx":19F97
            Top             =   360
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   26
            Left            =   210
            Picture         =   "frmMain.frx":1A381
            Top             =   1440
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "단위"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   35
            Left            =   480
            TabIndex        =   146
            Top             =   1470
            Width           =   360
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   27
            Left            =   2670
            Picture         =   "frmMain.frx":1A76B
            Top             =   1080
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   28
            Left            =   2670
            Picture         =   "frmMain.frx":1AB55
            Top             =   720
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "시약"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   36
            Left            =   2940
            TabIndex        =   145
            Top             =   1110
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Method"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   37
            Left            =   2940
            TabIndex        =   144
            Top             =   750
            Width           =   630
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Lot"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   38
            Left            =   2940
            TabIndex        =   143
            Top             =   390
            Width           =   255
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   29
            Left            =   2670
            Picture         =   "frmMain.frx":1AF3F
            Top             =   360
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   30
            Left            =   2670
            Picture         =   "frmMain.frx":1B329
            Top             =   1440
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "온도"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   39
            Left            =   2940
            TabIndex        =   142
            Top             =   1470
            Width           =   360
         End
      End
      Begin VB.Frame frameSet 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 시스템 설정 "
         Height          =   1935
         Left            =   270
         TabIndex        =   121
         Top             =   4470
         Visible         =   0   'False
         Width           =   5025
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1680
            TabIndex        =   123
            Text            =   "Combo1"
            Top             =   510
            Width           =   2295
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1680
            TabIndex        =   122
            Text            =   "Combo1"
            Top             =   1110
            Width           =   2295
         End
         Begin VB.Image Image1 
            Height          =   225
            Left            =   390
            Picture         =   "frmMain.frx":1B713
            Top             =   540
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "OCS"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   600
            TabIndex        =   127
            Top             =   570
            Width           =   435
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "프로토콜"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   600
            TabIndex        =   126
            Top             =   1170
            Width           =   780
         End
         Begin VB.Image Image4 
            Height          =   225
            Left            =   390
            Picture         =   "frmMain.frx":1BAFD
            Top             =   1140
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "OCS"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   5
            Left            =   4110
            TabIndex        =   125
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "OCS"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   6
            Left            =   4110
            TabIndex        =   124
            Top             =   1170
            Width           =   435
         End
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "시스템설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3660
         TabIndex        =   119
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1470
         TabIndex        =   104
         Top             =   1140
         Width           =   6705
         Begin VB.OptionButton optBarSeq 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check"
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
            Index           =   3
            Left            =   4170
            TabIndex        =   191
            Top             =   90
            Width           =   1155
         End
         Begin VB.OptionButton optBarSeq 
            BackColor       =   &H00FFFFFF&
            Caption         =   "R/P 사용"
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
            Index           =   2
            Left            =   2970
            TabIndex        =   190
            Top             =   90
            Width           =   1155
         End
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
            TabIndex        =   106
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
            TabIndex        =   105
            Top             =   90
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1470
         TabIndex        =   99
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
            TabIndex        =   101
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
            TabIndex        =   100
            Top             =   30
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1470
         TabIndex        =   96
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
            TabIndex        =   98
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
            TabIndex        =   97
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
               Picture         =   "frmMain.frx":1BEE7
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1C481
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1CA1B
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1CFB5
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1D847
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1D9A1
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1DAFB
               Key             =   "NOF"
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   885
         Left            =   300
         TabIndex        =   118
         Top             =   2490
         Width           =   4455
         _Version        =   393216
         _ExtentX        =   7858
         _ExtentY        =   1561
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
         SpreadDesigner  =   "frmMain.frx":1DC55
      End
      Begin FPSpread.vaSpread spdQcResult 
         Height          =   825
         Left            =   300
         TabIndex        =   153
         Top             =   3450
         Visible         =   0   'False
         Width           =   4455
         _Version        =   393216
         _ExtentX        =   7858
         _ExtentY        =   1455
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
         SpreadDesigner  =   "frmMain.frx":1DE8E
      End
      Begin VB.Label lblPatInfo 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "홍길"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   4
         Left            =   5640
         TabIndex        =   158
         Top             =   1410
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblPatInfo 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "홍길"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   5670
         TabIndex        =   157
         Top             =   1140
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblPatInfo 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "홍길"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   5640
         TabIndex        =   156
         Top             =   360
         Width           =   2100
      End
      Begin VB.Label lblPatInfo 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "홍길"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   5670
         TabIndex        =   155
         Top             =   630
         Width           =   2580
      End
      Begin VB.Label lblPatInfo 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "홍길"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   5640
         TabIndex        =   154
         Top             =   900
         Width           =   2820
      End
      Begin VB.Image imgDelete 
         Height          =   1260
         Left            =   6030
         Picture         =   "frmMain.frx":1E0C7
         Top             =   7410
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Image imgSave 
         Height          =   1260
         Left            =   7440
         Picture         =   "frmMain.frx":1FEE1
         Top             =   7530
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label3 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "바코드사용"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   107
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과적용"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   103
         Top             =   2130
         Width           =   780
      End
      Begin VB.Label Label2 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과전송"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   102
         Top             =   1710
         Width           =   780
      End
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "파일"
      Begin VB.Menu mnuHosp 
         Caption         =   "병원정보"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
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
         Begin VB.Menu mnuRackPos 
            Caption         =   "Rack/Pos"
         End
         Begin VB.Menu mnuCheckBox 
            Caption         =   "체크순"
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
      Caption         =   "원격지원"
      Begin VB.Menu mnuHelp01 
         Caption         =   "원격지원(TeamViewer)"
      End
      Begin VB.Menu mnuHelp02 
         Caption         =   "원격지원(LG Uplus)"
      End
      Begin VB.Menu mnuHelp03 
         Caption         =   "원격지원(ez Help)"
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

Dim pDel                As Boolean
Dim strOldBarno         As String

Private Sub cboBacteria_Click()

    If Trim(cboBacteria.Text) <> "선택" Then
        'Call setLocalDBMicro("Bacteria", "B0041004", "14", cboBacteria.Text)
        Call setLocalDBMicro("Bacteria", gUrinMic.BACT, "14", cboBacteria.Text)
    End If
    
End Sub

Private Sub cboEpCell_Click()

    If Trim(cboEpCell.Text) <> "선택" Then
        'Call setLocalDBMicro("E.P Cell", "B0041003", "13", cboEpCell.Text)
        Call setLocalDBMicro("E.P Cell", gUrinMic.EPIC, "13", cboEpCell.Text)
    End If
    
End Sub

Private Sub cboRbcM_Click()

    If Trim(cboRbcM.Text) <> "선택" Then
        'Call setLocalDBMicro("RBC(M)", "B0041002", "12", cboRbcM.Text)
        Call setLocalDBMicro("RBC(M)", gUrinMic.RBCM, "12", cboRbcM.Text)
    End If

End Sub

Private Sub cboWbcM_Click()
    
    If Trim(cboWbcM.Text) <> "선택" Then
        'Call setLocalDBMicro("WBC(M)", "B0041001", "11", cboWbcM.Text)
        Call setLocalDBMicro("WBC(M)", gUrinMic.WBCM, "11", cboWbcM.Text)
    End If
    
End Sub


Private Sub setLocalDBMicro(ByVal pChannel As String, ByVal pTestCd As String, ByVal pTestSeq As String, ByVal pResult As String)
    Dim RS1           As ADODB.Recordset
    Dim strExamDate As String
    Dim lngSAVESEQ  As Long
    
On Error GoTo RST

    strExamDate = Trim(Mid(GetText(spdOrder, gRow, colEXAMDATE), 1, 8))
    lngSAVESEQ = GetText(spdOrder, gRow, colSAVESEQ)
    
    SQL = ""
    SQL = SQL & "SELECT COUNT(*) AS CNT FROM PATRESULT " & vbCr
    SQL = SQL & " WHERE EXAMDATE = '" & strExamDate & "' " & vbCr
    SQL = SQL & "   AND SAVESEQ = " & lngSAVESEQ & vbCr
    SQL = SQL & "   AND EQUIPCODE = '" & pChannel & "' " & vbCr
    Set RS1 = AdoCn_Local.Execute(SQL, , 1)
    If Not RS1.EOF = True And Not RS1.BOF = True Then
        If Trim(RS1.Fields("CNT") & "") = 0 Then
            'insert into
            SQL = ""
            SQL = SQL & "INSERT INTO PATRESULT (" & vbCr
            SQL = SQL & "SAVESEQ"                           '저장순번(날짜별)
            SQL = SQL & ", EXAMDATE"                        '검사일자"
            SQL = SQL & ", HOSPDATE"                        '병원접수일자"
            SQL = SQL & ", EQUIPNO"                         '장비코드"
            SQL = SQL & ", BARCODE" & vbCrLf                '검체번호
            SQL = SQL & ", EQUIPCODE"                       '검사채널"
            SQL = SQL & ", ORDERCODE"                       '병원처방코드"
            SQL = SQL & ", EXAMCODE"                        '병원검사코드"
            SQL = SQL & ", EXAMSUBCODE"                     '병원검사코드(SUB)"
            SQL = SQL & ", EXAMNAME"                        '검사명
            SQL = SQL & ", SEQNO" & vbCrLf                  '검사일련번호"
            SQL = SQL & ", SAMPLETYPE"                      '검체유형"
            SQL = SQL & ", INOUT"                           '입/외
            SQL = SQL & ", DISKNO"                          'Rack (VERSACELL 에서는 실제 검사장비코드를 저장한다..)
            SQL = SQL & ", POSNO"                           'Pos
            SQL = SQL & ", EQUIPRESULT"                     '장비결과"
            SQL = SQL & ", RESULT" & vbCrLf                 'LIS 결과"
            SQL = SQL & ", REFJUDGE"                        '판정
            SQL = SQL & ", REFFLAG"                         'flag
            SQL = SQL & ", REFVALUE"                        '참고치
            SQL = SQL & ", CHARTNO"                         '챠트번호
            SQL = SQL & ", PID"                             '병록번호(내원번호)"
            SQL = SQL & ", PNAME" & vbCrLf
            SQL = SQL & ", PSEX"
            SQL = SQL & ", PAGE"
            SQL = SQL & ", PJUMIN"
            SQL = SQL & ", PANICVALUE"
            SQL = SQL & ", DELTAVALUE" & vbCrLf
            SQL = SQL & ", SENDFLAG"                        '전송구분(0:미전송,1:전송)"
            SQL = SQL & ", SENDDATE"
            SQL = SQL & ", EXAMUID"
            SQL = SQL & ", HOSPITAL)" & vbCrLf
            SQL = SQL & " VALUES (" & vbCrLf
            SQL = SQL & Trim(GetText(spdOrder, gRow, colSAVESEQ))
            SQL = SQL & ",'" & strExamDate & "'"
            SQL = SQL & ",'" & Trim(GetText(spdOrder, gRow, colHOSPDATE)) & "'"
            SQL = SQL & ",'" & gHOSP.MACHCD & "'"
            SQL = SQL & ",'" & Trim(GetText(spdOrder, gRow, colBARCODE)) & "'"
            SQL = SQL & ",'" & pChannel & "'"
            SQL = SQL & ",''"
            SQL = SQL & ",'" & pTestCd & "'"
            SQL = SQL & ",''"
            SQL = SQL & ",'" & pChannel & "'"
            SQL = SQL & ",'" & pTestSeq & "'"
            SQL = SQL & ",''"                                                   '검체유형
            SQL = SQL & ",'" & Trim(GetText(spdOrder, gRow, colINOUT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(spdOrder, gRow, colRACKNO)) & "'"
            SQL = SQL & ",'" & Trim(GetText(spdOrder, gRow, colPOSNO)) & "'"
            SQL = SQL & ",'" & pResult & "'"
            SQL = SQL & ",'" & pResult & "'"
            SQL = SQL & ",''"
            SQL = SQL & ",''"
            SQL = SQL & ",''"
            SQL = SQL & ",'" & Trim(GetText(spdOrder, gRow, colCHARTNO)) & "'"
            SQL = SQL & ",'" & Trim(GetText(spdOrder, gRow, colPID)) & "'"
            SQL = SQL & ",'" & Trim(GetText(spdOrder, gRow, colPNAME)) & "'"
            SQL = SQL & ",'" & Trim(GetText(spdOrder, gRow, colPSEX)) & "'"
            SQL = SQL & ",'" & Trim(GetText(spdOrder, gRow, colPAGE)) & "'"
            SQL = SQL & ",'" & Trim(GetText(spdOrder, gRow, colPJUMIN)) & "'"
            SQL = SQL & ",'" & Trim(GetText(spdOrder, gRow, colKEY1)) & "'"     'panic (accseq 대체사용)
            SQL = SQL & ",''"                                                   'delta
            SQL = SQL & ",'0'"                                                  '전송구분(0:미전송,1:전송)
            SQL = SQL & ",''"
            SQL = SQL & ",'" & gHOSP.USERID & "'"
            SQL = SQL & ",'" & gHOSP.HOSPNM & "')"
        Else
            'update
            SQL = ""
            SQL = SQL & "UPDATE PATRESULT SET "
            SQL = SQL & "  BARCODE = '" & Trim(GetText(spdOrder, gRow, colBARCODE)) & "'" & vbCr
            SQL = SQL & " ,INOUT   = '" & Trim(GetText(spdOrder, gRow, colINOUT)) & "'" & vbCr
            SQL = SQL & " ,CHARTNO = '" & Trim(GetText(spdOrder, gRow, colCHARTNO)) & "'" & vbCr
            SQL = SQL & " ,PID     = '" & Trim(GetText(spdOrder, gRow, colPID)) & "'" & vbCr
            SQL = SQL & " ,PNAME   = '" & Trim(GetText(spdOrder, gRow, colPNAME)) & "'" & vbCr
            SQL = SQL & " ,PSEX    = '" & Trim(GetText(spdOrder, gRow, colPSEX)) & "'" & vbCr
            SQL = SQL & " ,PAGE    = '" & Trim(GetText(spdOrder, gRow, colPAGE)) & "'" & vbCr
            SQL = SQL & " ,PJUMIN  = '" & Trim(GetText(spdOrder, gRow, colPJUMIN)) & "'" & vbCr
            SQL = SQL & " ,PANICVALUE  = '" & Trim(GetText(spdOrder, gRow, colKEY1)) & "'" & vbCr
            SQL = SQL & " ,EQUIPRESULT  = '" & pResult & "'" & vbCr
            SQL = SQL & " ,RESULT       = '" & pResult & "'" & vbCr
            SQL = SQL & " WHERE MID(EXAMDATE,1,8) = '" & Mid(Trim(GetText(spdOrder, gRow, colEXAMDATE)), 1, 8) & "'" & vbCr
            SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, gRow, colSAVESEQ)) & vbCr
            SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCr
            SQL = SQL & "   AND EQUIPCODE = '" & pChannel & "' " & vbCr
        End If
    End If
        
    RS1.Close
    
    If DBExec(AdoCn_Local, SQL) Then
        '-- 성공
        'Call spdOrder_Click(1, gRow)
    End If
    
Exit Sub

RST:
End Sub

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

Private Sub chkTest_Click()
    Dim strTest As String
    
    strTest = chkTest.Value
    
    Call WritePrivateProfileString("HOSP", "WORKTEST", strTest, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    
End Sub

Private Sub cmdAnalyteFind_Click()

    frmQCList.Tag = "Analyte조회"
    DoEvents
    frmQCList.Show 'vbModal
    
End Sub

'Private Sub cmdRefresh_Click()
'
'    Call GetTestList
'
'End Sub

Private Sub cmdAppend_Click()

    spdOrdMst.MaxRows = spdOrdMst.MaxRows + 1
    
End Sub

Private Sub cmdClear_Click()
    
    txtRcv.Text = ""
    
End Sub

Private Sub cmdConfig_Click()
    
    frmHospInfo.Show 'vbModal
    
End Sub

Private Sub cmdDelete_Click()
    
    spdOrdMst.Row = spdOrdMst.ActiveRow
    spdOrdMst.Action = ActionDeleteRow
    
    spdOrdMst.MaxRows = spdOrdMst.MaxRows - 1
    
End Sub

Private Sub cmdEnd_Click()

    If MsgBox("장비와 통신중입니다. 종료하시겠습니까?", vbYesNo + vbCritical, "프로그램 종료") = vbYes Then
    
        If comEqp.PortOpen = True Then
            comEqp.PortOpen = False
        End If
    
        Call DisConnect_Server
        
        Call DisConnect_Local
        
        Unload Me
        
        End
    End If
    
End Sub

Private Sub cmdIF_Click()

    If FraHidden.Visible = True Then
        FraHidden.Visible = False
    Else
        FraHidden.Visible = True
        FraHidden.ZOrder 0
    End If
    
End Sub

Private Sub cmdInit_Click()
    
    Call InitialComm
    
End Sub

Private Sub cmdLabFind_Click()

    frmQCList.Caption = "Lab조회"
    frmQCList.Show 'vbModal
    
End Sub

Private Sub cmdQCMaster_Click()

    frmQCMaster.Show 'vbModal
    
End Sub

Private Sub cmdSend_Click()
'    Dim varTmp As Variant
    Dim lngBufLen   As Long
    Dim BufChar     As String
    Dim i           As Integer
    
'    Erase strRecvData
'    varTmp = Replace(txtRcv.Text, vbLf, "")
'    varTmp = Split(varTmp, vbCr)
'
'    For i = 0 To UBound(varTmp)
'        ReDim Preserve strRecvData(i + 1)
'        strRecvData(i + 1) = varTmp(i)
'    Next
    
    
    RcvBuffer = txtRcv.Text
    pBuffer = txtRcv.Text
    
    Select Case UCase(gHOSP.MACHNM)
        Case "LIAISON"
            Call Phase_Serial_LIAISON
        Case "URINSCANPRO"
            'pBuffer = Replace(pBuffer, STX, "")
            Call SerialRcvData_UrinscanPro

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


Private Sub cmdTVSave_Click()
    Dim i As Integer
    
    If Trim(txtTV.Text) <> "" Then
        If IsNumeric(Trim(txtTV.Text)) Then
            If lblPatInfo(1).Caption <> "" And lblPatInfo(2).Caption <> "" Then
                Call SetLocalDB_TV(spdROrder.ActiveRow, 1, 1, txtTV.Text)
                
                Call spdROrder_Click(colBARCODE, spdROrder.ActiveRow)
                        
                For i = 1 To spdRResult.MaxRows
                    Call CalProcess(spdROrder, spdRResult, GetText(spdRResult, i, colRTESTCD), Trim(txtTV.Text))
                Next
            End If
        Else
            MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbCritical, Me.Caption
            txtTV.SetFocus
        End If
    End If

End Sub

Private Sub cmdWork_Click()
    
    Call GetWorkList(Format(dtpFrDt.Value, "yyyymmdd"), Format(dtpToDt.Value, "yyyymmdd"), spdOrder)

End Sub

Private Sub cmdWorkAll_Click()
    Dim i As Integer
        
    pDel = True
    
    With spdWork
        For i = 1 To .DataRowCnt
            If GetText(spdWork, i, colCHECKBOX) = "1" Then
                Call spdWork_DblClick(colBARCODE, i)
            End If
        Next
    End With
    
    spdWork.MaxRows = 0
    pDel = False
    
End Sub

Private Sub Command2_Click()
Dim intTestCnt  As Integer
Dim strSlip1    As String
Dim strSlip2    As String
Dim strBarcode  As String
Dim sqlRet      As Integer
Dim RS            As ADODB.Recordset
'-- 바코드 번호로 오더 조회
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim prm3 As New ADODB.Parameter


On Error GoTo RST
    
    strBarcode = Text1.Text

    Dim strDate     As String
    
    Dim intBcNow    As Integer
    Dim intBcFive   As Integer
    Dim intBcAdd    As Integer
    Dim strADT      As String
    'Dim strSlip1    As String
    'Dim strSlip2    As String
    
    strDate = Format(Now, "yyyy-mm-dd")
    
    intBcNow = DateDiff("d", "1999-01-01", strDate)
    intBcFive = Mid(strBarcode, 1, 5) '06351
    
    intBcAdd = intBcFive - intBcNow
    'strADT = Format(Now + intBcAdd, "yyyymmdd")
    strADT = Format(Now + intBcAdd, "yyyy-mm-dd")


'    If Text2.Text <> "" Then
'        strADT = Text2.Text
'    End If



    strSlip1 = Mid(strBarcode, 6, 2)  '10으로 시작하면 TLA프로시저를 태움
                                      '나머지는 EXAM_INTERFACE_S
    strSlip2 = Mid(strBarcode, 8, 5)  '00001
'MsgBox "8"
    
    '-- SP 사용
    'AdoCn.CursorLocation = adUseClient
    'Set RS = AdoCn.Execute("Exec EXAM_INTERFACE_S '" & Format(Now, "yyyymmdd") & "'," & strSlip1 & "," & strSlip2, sqlRet)
    
'
''    접수일자
''       slipno1
''       slipno2
'
''
''EXAM_INTERFACE_S
''       I_JEOBSUDT Date
''    , I_SLIPNO1       NUMBER
''    , I_SLIPNO2       NUMBER
''
''
''
''EXAM_INTERFACE_ARR_U
''       I_PTNO VARCHAR2
''    , I_JEOBSUDT      DATE
''    , I_SLIPNO1       NUMBER
''    , I_SLIPNO2       NUMBER
''
'
'
'
'
'
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

'MsgBox "9"

    AdoCmd.CommandTimeout = 15
    If strSlip1 = "10" Then
        AdoCmd.CommandText = "TW_MIS_EXAM.EXAM_TLA_INTERFACE_S"
    Else
        AdoCmd.CommandText = "EXAM_INTERFACE_S"
    End If
    AdoCmd.CommandType = adCmdStoredProc
'MsgBox "0"


    If strSlip1 = "10" Then
        Set prm1 = AdoCmd.CreateParameter("I_JEOBSUDT", adDate, adParamInput, 10, strADT)
        AdoCmd.Parameters.Append prm1
        Set prm2 = AdoCmd.CreateParameter("I_BARCODE", adDouble, adParamInput, 12, strBarcode)
        AdoCmd.Parameters.Append prm2
        'Set prm3 = AdoCmd.CreateParameter("I_SLIPNO2", adInteger, adParamInput, 5, strSlip2)
        'AdoCmd.Parameters.Append prm3
    Else
        Set prm1 = AdoCmd.CreateParameter("I_JEOBSUDT", adDate, adParamInput, 10, strADT)
        AdoCmd.Parameters.Append prm1
        Set prm2 = AdoCmd.CreateParameter("I_SLIPNO1", adInteger, adParamInput, 2, strSlip1)
        AdoCmd.Parameters.Append prm2
        Set prm3 = AdoCmd.CreateParameter("I_SLIPNO2", adInteger, adParamInput, 5, strSlip2)
        AdoCmd.Parameters.Append prm3
    End If
    Set RS = New ADODB.Recordset
    RS.Open AdoCmd.Execute
'
'MsgBox "4"

    'ptno, sname, sex, ageyy, deptcode, gber, slipno1, slipno2, itemcd, itemnm, geomchc1, status, result1
    
    intTestCnt = 0
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With vaSpread1
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                vaSpread1.MaxRows = intTestCnt
                
                SetText vaSpread1, Trim(RS.Fields("ptno")) & "", intTestCnt, 1
                SetText vaSpread1, Trim(RS.Fields("sname")) & "", intTestCnt, 2
                SetText vaSpread1, Trim(RS.Fields("sex")) & "", intTestCnt, 3
                SetText vaSpread1, Trim(RS.Fields("ageyy")) & "", intTestCnt, 4
                SetText vaSpread1, Trim(RS.Fields("deptcode")) & "", intTestCnt, 5
'                SetText vaSpread1, Trim(RS.Fields("gber")) & "", intTestCnt, 6
                SetText vaSpread1, Trim(RS.Fields("slipno1")) & "", intTestCnt, 7
                SetText vaSpread1, Trim(RS.Fields("slipno2")) & "", intTestCnt, 8
                SetText vaSpread1, Trim(RS.Fields("itemcd")) & "", intTestCnt, 9
                SetText vaSpread1, Trim(RS.Fields("itemnm")) & "", intTestCnt, 10
                SetText vaSpread1, Trim(RS.Fields("geomchc1")) & "", intTestCnt, 11
                SetText vaSpread1, Trim(RS.Fields("status")) & "", intTestCnt, 12
                SetText vaSpread1, Trim(RS.Fields("result1")) & "", intTestCnt, 13
                
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close


Exit Sub

RST:
    
    MsgBox Err.Number & vbCr & Err.Description
   ' RS.Close
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = 1
    Call cmdEnd_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'If MsgBox("장비와 통신중입니다. 종료하시겠습니까?", vbYesNo + vbCritical, "프로그램 종료") = vbYes Then
    
        If comEqp.PortOpen = True Then
            comEqp.PortOpen = False
        End If
    
        Call DisConnect_Server
        
        Call DisConnect_Local
        
        Unload Me
        
        End
    'End If
    
End Sub

Private Sub fraResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblResult.ForeColor = vbBlack
    lblRClear.ForeColor = vbBlack
    lblRExcel.ForeColor = vbBlack
    shpR.BorderColor = &H808080
    shpRC.BorderColor = &H808080
    shpRX.BorderColor = &H808080

End Sub

Private Sub lblRClear_Click()
    
    spdROrder.MaxRows = 0
    spdRResult.MaxRows = 0
    
End Sub

Private Sub lblRClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblResult.ForeColor = vbBlack
    lblRClear.ForeColor = vbBlack
    lblRExcel.ForeColor = vbBlack
    shpR.BorderColor = &H808080
    shpRC.BorderColor = &H808080
    shpRX.BorderColor = &H808080
    
    lblRClear.ForeColor = vbBlue
    shpRC.BorderColor = vbCyan
    
End Sub

Private Sub lblResult_Click()

    frmMain.spdROrder.MaxRows = 0
    frmMain.spdRResult.MaxRows = 0

    Call GetResultList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), cboRstType.ListIndex, cboState.ListIndex)
    
End Sub

Private Sub lblResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblResult.ForeColor = vbBlack
    lblRClear.ForeColor = vbBlack
    lblRExcel.ForeColor = vbBlack
    shpR.BorderColor = &H808080
    shpRC.BorderColor = &H808080
    shpRX.BorderColor = &H808080
    
    lblResult.ForeColor = vbBlue
    shpR.BorderColor = vbCyan
    
End Sub

Private Sub lblRExcel_Click()
    Dim sFileName As String

    If spdROrder.DataRowCnt < 1 Then
        MsgBox "저장할 자료가 없습니다.", , "알 림"
        Exit Sub
    Else
        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
        CommonDialog1.ShowSave
        sFileName = CommonDialog1.Filename
        SaveExcel sFileName, spdROrder
        MsgBox "엑셀 저장완료", vbOKOnly + vbInformation, Me.Caption
    End If
End Sub

Sub SaveExcel(Filename As String, argSpread As vaSpread)

On Error Resume Next

' Excel Object Library 와 연결합니다.
Dim xlapp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim iRow As Integer
Dim iCol As Integer
Dim i As Integer

    Set xlapp = CreateObject("Excel.Application")
    
    xlapp.DisplayAlerts = False
    
    Set xlBook = xlapp.Workbooks.Add
    
    Set xlSheet = xlBook.Worksheets(1)
     
    For iRow = 0 To argSpread.DataRowCnt
        For iCol = 1 To argSpread.DataColCnt
            argSpread.Row = iRow
            argSpread.Col = iCol
            xlSheet.Cells(iRow + 1, iCol) = argSpread.Text
        Next iCol
    Next iRow
    
    xlBook.SaveAs (Filename)
    xlapp.Quit


End Sub

Private Sub lblRExcel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblResult.ForeColor = vbBlack
    lblRClear.ForeColor = vbBlack
    lblRExcel.ForeColor = vbBlack
    shpR.BorderColor = &H808080
    shpRC.BorderColor = &H808080
    shpRX.BorderColor = &H808080
    
    lblRExcel.ForeColor = vbBlue
    shpRX.BorderColor = vbCyan

End Sub

Private Sub lblSave_Click()
    Dim lRow As Long
    Dim Res  As Integer
    
    For lRow = 1 To spdOrder.DataRowCnt
        spdOrder.Row = lRow
        spdOrder.Col = 1
        If spdOrder.Value = 1 Then
            
            Res = SaveTransData(lRow)
        
            If Res = -1 Then
                SetForeColor spdOrder, lRow, lRow, 1, colSTATE, 255, 0, 0
                SetText spdOrder, "Failed", lRow, colSTATE
            Else
                spdOrder.Row = lRow
                spdOrder.Col = 1
                spdOrder.Value = 1
                
                SetBackColor spdOrder, lRow, lRow, 1, colSTATE, 202, 255, 112
                SetText spdOrder, "Trans", lRow, colSTATE
                
                      SQL = " UPDATE PATRESULT SET " & vbCrLf
                SQL = SQL & "  SENDFLAG = '2' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
                
            End If
            spdOrder.Row = lRow
            spdOrder.Col = 1
            spdOrder.Value = 0
        End If
    Next lRow

End Sub

Private Sub mnuBarcode_Click()
    
    mnuBarcode.Checked = True
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False
    
    Call WritePrivateProfileString("HOSP", "BARUSE", "Y", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

End Sub

Private Sub mnuCheckBox_Click()
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = True
    
    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "3", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

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
    
    Call WritePrivateProfileString("HOSP", "SAVELIS", "N", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

End Sub


Private Sub mnuExit_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuHelp01_Click()

    Call WinExec(App.PATH & "\TeamViewerQS.exe", 1)
    
End Sub

Private Sub mnuHelp02_Click()

    Call WinExec("C:\Program Files (x86)\Internet Explorer\iexplore.exe http://cs1472.com/customer/", 1)


End Sub

Private Sub mnuHelp03_Click()

    Call WinExec("C:\Program Files (x86)\Internet Explorer\iexplore.exe https://939.co.kr/easyqc/", 1)

End Sub

Private Sub mnuHosp_Click()
    
    frmHospInfo.Show vbModal
    
End Sub

Private Sub mnuLisResult_Click()
    
    mnuEqpResult.Checked = False
    mnuLisResult.Checked = True
    
    Call WritePrivateProfileString("HOSP", "SAVELIS", "Y", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

End Sub

Private Sub mnuRackPos_Click()
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = True
    mnuCheckBox.Checked = False
    
    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "2", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

End Sub

Private Sub mnuSaveAuto_Click()
    
    mnuSaveAuto.Checked = True
    mnuSaveManual.Checked = False
    
    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "Y", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

End Sub

Private Sub mnuSaveManual_Click()
    
    mnuSaveAuto.Checked = False
    mnuSaveManual.Checked = True
    
    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "N", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")


End Sub

Private Sub mnuSeqno_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = True
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False
    
    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "1", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

End Sub

Private Sub mnuTest_Click()
    
    Call lblMenu_Click(2)

End Sub

Public Sub spdOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim strNewBarNo As String
    Dim intRow      As Integer
    Dim strSeq      As String
    
    
    sRow = spdOrder.ActiveRow
    sCol = spdOrder.ActiveCol
    
    If sRow = 0 Then
        Exit Sub
    End If
    
    strNewBarNo = GetText(spdOrder, sRow, sCol)
    
    If KeyCode = vbKeyReturn Then
        If colBARCODE = sCol Then
            If GetSampleInfo(sRow, spdOrder) = -1 Then
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
                SQL = SQL & " ,PANICVALUE  = '" & Trim(GetText(spdOrder, sRow, colKEY1)) & "'" & vbCr
                SQL = SQL & " WHERE mid(EXAMDATE,1,8) = '" & Mid(Trim(GetText(spdOrder, sRow, colEXAMDATE)), 1, 8) & "'" & vbCr
                SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, sRow, colSAVESEQ)) & vbCr
                SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCr

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
        ElseIf sCol = colSEQNO Then
            With spdOrder
                strSeq = GetText(spdOrder, .ActiveRow, .ActiveCol)
                If Not IsNumeric(strSeq) Then
                    MsgBox "숫자만 입력이 가능합니다"
                    Exit Sub
                End If
                For intRow = .ActiveRow + 1 To .MaxRows
                    Call SetText(spdOrder, strSeq + 1, intRow, colSEQNO)
                    strSeq = strSeq + 1
                Next
            End With
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If MsgBox(strNewBarNo & " 를 지우시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        DeleteRow spdROrder, sRow, sRow
        spdROrder.MaxRows = spdROrder.MaxRows - 1
        spdRResult.MaxRows = 0
        
    End If
End Sub



Private Sub spdROrder_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol As Integer
    
    If Row = 0 Then
        Call SetSpreadSort(spdROrder, 0)
        Exit Sub
    End If
    
    '-- 환자정보표시
    lblPatInfo(0).Caption = GetText(spdROrder, Row, colPNAME) & " [" & GetText(spdROrder, Row, colPAGE) & "/" & GetText(spdROrder, Row, colPSEX) & "]  "
    lblPatInfo(1).Caption = GetText(spdROrder, Row, colBARCODE)
    lblPatInfo(2).Caption = GetText(spdROrder, Row, colPID)
    lblPatInfo(3).Caption = spdROrder.ActiveRow
    lblPatInfo(4).Caption = GetText(spdROrder, Row, colRACKNO)

    
    txtTV.Text = ""
    
    '-- 결과표시
    If GetPatTRestResult_Search(Row) = -1 Then
        '장비결과가 없을경우 검사명만 보여주기
        spdRResult.MaxRows = 0
        With spdROrder
            For intCol = colSTATE + 1 To .MaxCols
                If GetText(spdROrder, Row, intCol) <> "" Then    '◆
                    spdRResult.MaxRows = spdRResult.MaxRows + 1
                    Call SetText(spdRResult, GetText(spdROrder, 0, intCol), spdRResult.MaxRows, colRTESTNM)
                    spdRResult.RowHeight(-1) = 12
                End If
            Next
        End With
    End If

    'txtTV.SetFocus
    
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
        
    ElseIf KeyCode = vbKeyReturn Then
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

Private Sub InitialComm()
    Dim sSendBuf$
    
    intPhase = 1
    
    msMT = Chr(&H30)
    sSendBuf = msMT & "I " & vbCr & vbLf
    sSendBuf = CheckSum_ADVIA2120(sSendBuf)
    
    comEqp.Output = sSendBuf
    SetRawData "[Tx]" & sSendBuf
       
End Sub

Private Sub spdWork_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer
    
    If Row = 0 And Col <> colCHECKBOX Then
        Call SetSpreadSort(spdWork, 0)
        Exit Sub
    End If
    
    If Row = 0 And Col = colCHECKBOX Then
        If GetText(spdWork, 1, colCHECKBOX) = "1" Then
            For i = 1 To spdWork.DataRowCnt
                Call SetText(spdWork, "0", i, colCHECKBOX)
            Next
        Else
            For i = 1 To spdWork.DataRowCnt
                Call SetText(spdWork, "1", i, colCHECKBOX)
            Next
        End If
    End If
    
    If Row > 0 And Col = colCHECKBOX Then
        If GetText(spdWork, Row, colCHECKBOX) = "1" Then
            Call SetText(spdWork, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdWork, "1", Row, colCHECKBOX)
        End If
    End If
    
End Sub

Private Sub spdWork_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    Dim strBarno_Work   As String
    
    If Row = 0 Then Exit Sub
    
    intWRow = Row
    spdWork.Row = Row
    spdWork.Col = colBARCODE
    strBarno_Work = Trim(spdWork.Text)
    
    With frmMain.spdOrder
        If chkTest.Value = "0" Then
            blnSame = False
            For intORow = 1 To .MaxRows
                .Row = intORow
                .Col = colBARCODE
                If strBarno_Work = Trim(.Text) Then
                    blnSame = True
                    Exit For
                End If
            Next
            If blnSame = False Then
                frmMain.spdOrder.MaxRows = frmMain.spdOrder.MaxRows + 1
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSPECNO), frmMain.spdOrder.MaxRows, colSPECNO)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHECKBOX), frmMain.spdOrder.MaxRows, colCHECKBOX)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colHOSPDATE), frmMain.spdOrder.MaxRows, colHOSPDATE)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colBARCODE), frmMain.spdOrder.MaxRows, colBARCODE)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), frmMain.spdOrder.MaxRows, colSEQNO)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHARTNO), frmMain.spdOrder.MaxRows, colCHARTNO)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPID), frmMain.spdOrder.MaxRows, colPID)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colINOUT), frmMain.spdOrder.MaxRows, colINOUT)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPNAME), frmMain.spdOrder.MaxRows, colPNAME)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPSEX), frmMain.spdOrder.MaxRows, colPSEX)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPAGE), frmMain.spdOrder.MaxRows, colPAGE)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPJUMIN), frmMain.spdOrder.MaxRows, colPJUMIN)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colOCNT), frmMain.spdOrder.MaxRows, colOCNT)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), frmMain.spdOrder.MaxRows, colSEQNO)
                
                varItems = GetText(spdWork, intWRow, colITEMS)
                varItems = Split(varItems, "/")
                For intItems = 0 To UBound(varItems)
                    For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
                        frmMain.spdOrder.Row = 0
                        frmMain.spdOrder.Col = intOCol
                        If varItems(intItems) = Trim(frmMain.spdOrder.Text) Then
                            .Row = frmMain.spdOrder.MaxRows
                            Call SetText(frmMain.spdOrder, "◆", frmMain.spdOrder.MaxRows, intOCol)
                        End If
                    Next
                Next
                
                frmMain.spdOrder.RowHeight(-1) = 12
            End If
        '검사먼저 한 경우
        Else
            blnSame = False
            For intORow = 1 To .MaxRows
                .Row = intORow
                .Col = colHOSPDATE
                If Trim(.Text) = "" Then
                    blnSame = True
                    Exit For
                End If
            Next
            
            If blnSame = True Then
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSPECNO), intORow, colSPECNO)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHECKBOX), intORow, colCHECKBOX)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colHOSPDATE), intORow, colHOSPDATE)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colBARCODE), intORow, colBARCODE)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), intORow, colSEQNO)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHARTNO), intORow, colCHARTNO)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPID), intORow, colPID)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colINOUT), intORow, colINOUT)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPNAME), intORow, colPNAME)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPSEX), intORow, colPSEX)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPAGE), intORow, colPAGE)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPJUMIN), intORow, colPJUMIN)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colOCNT), intORow, colOCNT)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), intORow, colSEQNO)
                    
                '정보수정
                SQL = ""
                SQL = SQL & "UPDATE PATRESULT SET "
                SQL = SQL & "  BARCODE = '" & Trim(GetText(spdOrder, intORow, colBARCODE)) & "'" & vbCr
                SQL = SQL & " ,INOUT   = '" & Trim(GetText(spdOrder, intORow, colINOUT)) & "'" & vbCr
                SQL = SQL & " ,CHARTNO = '" & Trim(GetText(spdOrder, intORow, colCHARTNO)) & "'" & vbCr
                SQL = SQL & " ,PID     = '" & Trim(GetText(spdOrder, intORow, colPID)) & "'" & vbCr
                SQL = SQL & " ,PNAME   = '" & Trim(GetText(spdOrder, intORow, colPNAME)) & "'" & vbCr
                SQL = SQL & " ,PSEX    = '" & Trim(GetText(spdOrder, intORow, colPSEX)) & "'" & vbCr
                SQL = SQL & " ,PAGE    = '" & Trim(GetText(spdOrder, intORow, colPAGE)) & "'" & vbCr
                SQL = SQL & " ,PJUMIN  = '" & Trim(GetText(spdOrder, intORow, colPJUMIN)) & "'" & vbCr
                SQL = SQL & " ,PANICVALUE  = '" & Trim(GetText(spdOrder, intORow, colKEY1)) & "'" & vbCr
                SQL = SQL & " WHERE mid(EXAMDATE,1,8) = '" & Mid(Trim(GetText(spdOrder, intORow, colEXAMDATE)), 1, 8) & "'" & vbCr
                SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, intORow, colSAVESEQ)) & vbCr
                SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCr
    
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            Else
                blnSame = False
                For intORow = 1 To .MaxRows
                    .Row = intORow
                    .Col = colBARCODE
                    If strBarno_Work = Trim(.Text) Then
                        blnSame = True
                        Exit For
                    End If
                Next
                If blnSame = False Then
                    frmMain.spdOrder.MaxRows = frmMain.spdOrder.MaxRows + 1
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSPECNO), frmMain.spdOrder.MaxRows, colSPECNO)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHECKBOX), frmMain.spdOrder.MaxRows, colCHECKBOX)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colHOSPDATE), frmMain.spdOrder.MaxRows, colHOSPDATE)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colBARCODE), frmMain.spdOrder.MaxRows, colBARCODE)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), frmMain.spdOrder.MaxRows, colSEQNO)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHARTNO), frmMain.spdOrder.MaxRows, colCHARTNO)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPID), frmMain.spdOrder.MaxRows, colPID)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colINOUT), frmMain.spdOrder.MaxRows, colINOUT)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPNAME), frmMain.spdOrder.MaxRows, colPNAME)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPSEX), frmMain.spdOrder.MaxRows, colPSEX)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPAGE), frmMain.spdOrder.MaxRows, colPAGE)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPJUMIN), frmMain.spdOrder.MaxRows, colPJUMIN)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colOCNT), frmMain.spdOrder.MaxRows, colOCNT)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), frmMain.spdOrder.MaxRows, colSEQNO)
                    
                    varItems = GetText(spdWork, intWRow, colITEMS)
                    varItems = Split(varItems, "/")
                    For intItems = 0 To UBound(varItems)
                        For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
                            frmMain.spdOrder.Row = 0
                            frmMain.spdOrder.Col = intOCol
                            If varItems(intItems) = Trim(frmMain.spdOrder.Text) Then
                                .Row = frmMain.spdOrder.MaxRows
                                Call SetText(frmMain.spdOrder, "◆", frmMain.spdOrder.MaxRows, intOCol)
                            End If
                        Next
                    Next
                    
                    frmMain.spdOrder.RowHeight(-1) = 12
                End If
            
            End If
            frmMain.spdOrder.RowHeight(-1) = 12
        End If
        
        If pDel = False Then
            Call spdWork.DeleteRows(intWRow, 1)
            spdWork.MaxRows = spdWork.MaxRows - 1
        End If
    End With
    
End Sub

Private Sub Timer2120_Timer()
    Timer2120.Enabled = False
    Timer2120.Interval = 0
    
    Select Case msTimerFlag
        Case "I"
            If msMT = "" Then
                Call InitialComm
                
                msSndPacket = ""
                msTimerFlag = ""
            Else
                mp_bReserveEnd = True
                'PropertyChanged "ReserveEnd"
                
                If mp_bPortOpen = False Then
                    mp_bReserveEnd = False
                    'PropertyChanged "ReserveEnd"
                    
                    comEqp.PortOpen = True
                    mp_bPortOpen = True
                    'PropertyChanged "PortOpen"
                    
                    Sleep 1000
                    
                    Call InitialComm
                    
                    msSndPacket = ""
                    msTimerFlag = ""
                Else
                    'Timer 재가동
                    msTimerFlag = "I"
                    Timer2120.Interval = 1000
                    Timer2120.Enabled = True
                End If
            End If
            
        Case Else
            If msSndPacket = "" Then Exit Sub
    
            comEqp.Output = msSndPacket
            SetRawData "[Tx]" & msSndPacket

            msSndPacket = ""
            msTimerFlag = ""
            
    End Select
End Sub

Private Sub TimerVESCUBE_Timer()
    
    TimerVESCUBE.Enabled = False
    
    Call VesMatic(RcvBuffer)
    
End Sub

Private Sub tmrComm_Timer()
    
    tmrComm.Enabled = False
    tmrFlipFlop.Enabled = False

    
    lblCommStatus.Caption = ""
    
End Sub

Private Sub tmrFlipFlop_Timer()
    lblCommStatus.ForeColor = vbBlue
    
    If lblCommStatus.Visible = True Then
        lblCommStatus.Visible = False
    Else
        lblCommStatus.Visible = True
    End If
    
    
    If lblMenu(0).ForeColor = vbBlack Then
        lblMenu(0).ForeColor = vbBlue
        shpB(0).BorderColor = vbCyan
    Else
        lblMenu(0).ForeColor = vbBlack
        shpB(0).BorderColor = vbGreen
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


'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
'                                 2. 장비수신정보 화면표시,
'                                 3. 처방코드 가져오기,
'                                 4. (처방코드로)검사오더 만들기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Private Sub GetOrder_ADVIA2120(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If
        
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
            
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_ADVIA2120(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If


        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub



'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_ADVIA2120(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String
'    Dim blnCBC      As Boolean
'    Dim blnDIFF     As Boolean
'    Dim blnRETI     As Boolean
'    Dim strPART     As String

    GetEquipExamCode_ADVIA2120 = ""
    strExamCode = ""
    mOrder.SendCnt = 0
    
'    blnCBC = False
'    blnDIFF = False
'    blnRETI = False
'    strPART = ""
    
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    '-- 가져온 검사코드의 파트찾기(CBC,DIFF,RET)
          SQL = "Select DISTINCT RSLTCHANNEL " & vbCr
    SQL = SQL & "  From EQPMASTER " & vbCr
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "'" & vbCr
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")" & vbCr
    SQL = SQL & " Order By RSLTCHANNEL "
        
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            ' "001002003004005009"
            strExamCode = strExamCode & Right("000" & Trim(AdoRs_Local.Fields("RSLTCHANNEL").Value & ""), 3)
            'strExamCode = strExamCode & Format(Trim(AdoRs_Local.Fields("RSLTCHANNEL").Value & ""), "000")
            mOrder.SendCnt = mOrder.SendCnt + 1
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_ADVIA2120 = strExamCode
    
    
End Function

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_ADVIA1800()
    Dim strOutput   As String     '송신할 데이터
    
    Select Case sSndState
        Case ""
            iIdleFlag = CStr(Val(iIdleFlag) + 1)
            
            '## Order 없는 경우
            If mOrder.NoOrder = True Then
                strOutput = ""
                strOutput = intFrameNo & "O" & " " & "0101"
                strOutput = strOutput & "000"                                                   'Sample Count
                strOutput = strOutput & "N"                                                     'Sample classification
                strOutput = strOutput & "2"                                                     'Registration data(0:New, 1:Add, 2:No Request, 3:Sample Delete)
                strOutput = strOutput & Left$(mOrder.BarNo & Space(13), 13)                     'Sample Number
                strOutput = strOutput & Space$(7) & Space$(16) & Space$(16) & "M" & Space$(3)   'Length = 45
                strOutput = strOutput & Space$(8) & " 1.0" & "1" & "1"                          '
                strOutput = strOutput & Space$(1) & ETX
            Else
                '1O 0101010N003498582                                            M            1.011 89M 81M 82M 90M 91M 85M106M103M104M105M 15
                strOutput = ""
                strOutput = intFrameNo & "O" & " " & "0101"
                strOutput = strOutput & Format$(mOrder.SendCnt, "000")                          'Sample Count
                strOutput = strOutput & "N"                                                     'Sample classification
                strOutput = strOutput & "0"                                                     'Registration data(0:New, 1:Add, 2:No Request, 3:Sample Delete)
                strOutput = strOutput & Left$(mOrder.BarNo & Space(13), 13)                     'Sample Number
                strOutput = strOutput & Space$(7) & Space$(16) & Space$(16) & "M" & Space$(3)   'Length = 45
                strOutput = strOutput & Space$(8)                                               '
                strOutput = strOutput & " 1.0"                                                  'Dilution coefficient(4)
                If mOrder.SPCCD = "1" Then                                                      'Sample classification(1:blood serum, 2:urine)
                    strOutput = strOutput & "1"
                Else
                    strOutput = strOutput & "2"
                End If
                strOutput = strOutput & "1"                                                     'Container classification
                strOutput = strOutput & mOrder.Order & Space$(1) & ETX
                
            End If
            
            'n개의 sSndPacket 구성
            ReDim Preserve sSndPacket(Val(iIdleFlag))
            sSndPacket(Val(iIdleFlag)) = STX & strOutput & GetChkSum(strOutput) & vbCr & vbLf
            
            intFrameNo = intFrameNo + 1
            
        Case "E"  '## 처음 Packet 전송
            iOrderFlag = 1
            frmMain.comEqp.Output = sSndPacket(iOrderFlag)
            SetRawData "[Tx]" & sSndPacket(iOrderFlag)
            
            If iOrderFlag = iTotQueryFlag Then
                sSndState = "L"
            Else
                sSndState = "P"
            End If
            
        Case "P"  '## Packet 전송
            iOrderFlag = iOrderFlag + 1
            frmMain.comEqp.Output = sSndPacket(iOrderFlag)
            SetRawData "[Tx]" & sSndPacket(iOrderFlag)
            
            If iOrderFlag = iTotQueryFlag Then
                sSndState = "L"
            Else
                sSndState = "P"
            End If
            
        Case "L"  '## EOT
            'strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            
            iOrderFlag = 0: iPendingFlag = 0: iIdleFlag = 0: iTotQueryFlag = 0
            intFrameNo = 1
            
            Exit Sub
    End Select
    
    If intFrameNo = 8 Then
        intFrameNo = 0
    End If
    
'    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
'    frmMain.comEqp.Output = strOutput
'    SetRawData "[Tx]" & strOutput

End Sub


Private Sub SendOrder_COULTERLH780()
'    On Error GoTo ErrRtn
'
'    Dim sSend   As String * 256
'    Dim sSendStr    As String
'    Dim sChkSum As String
'
'    Select Case intSndPhase
'        Case 1
'
'            Call Get_OrderString
'
'            If pSampleInfo.ORDCNT = 0 Then
'                Exit Sub
'            End If
'
'            comEqp.Output = "01"
'            intSndPhase = intSndPhase + 1
'
'            Exit Sub
'
'        Case 2
'            sSend = pSampleInfo.Kind
'
'            sChkSum = ChkSum_LH750(sSend)
'
'            sSendStr = Chr(2) & Format(1, "00") & sSend & sChkSum & Chr(3)
'
'            comEqp.Output = sSendStr
'
'
'    End Select
'
'ErrRtn:
'    If Err <> 0 Then
'        'err
'    End If
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_ACLELITE()
    Dim strOutput As String     '송신할 데이터
    Dim strDate   As String
    
    Select Case intSndPhase
        Case 1  '## Header
            strDate = Format$(Now, "YYYYMMDDHHMMSS")
            strOutput = intFrameNo & "H|\^&||||||||ACL9000||P|1|" & strDate & vbCr & ETX
            
                                    '1H|\^&|||ACL9000|||||||P|1|20160827080438

            '## 접수정보 유무를 판단하여 SndPhase변경
            If mOrder.NoOrder = True Then
                '## 접수정보가 없는경우
                intSndPhase = 4
            Else
                intSndPhase = 2
            End If
        
            intFrameNo = intFrameNo + 1
        
        Case 2  '## Patient
            'strOutput = intFrameNo & "P|1||" & mOrder.PatId & "|||||" & mOrder.Sex & "||||||||||||||||||||||||||" & vbCr & ETX
            strOutput = intFrameNo & "P|1||" & mOrder.PID & "|||||||||||||||||||||||||||||||" & vbCr & ETX
            
                                     'P|1||                    ||||||||||||||||||||||||||||||||

            intSndPhase = 3
            
            intFrameNo = intFrameNo + 1
            
        Case 3  '## Order
            With mOrder
                'strOutput = intFrameNo & "O|" & CStr(.SendCnt + 1) & "|" & .BarNo & "||" & .Items(.SendCnt + 1) & "|" & .StatFg & "||||||||||||||||||||0||||||" & vbCr & ETX
                 strOutput = intFrameNo & "O|" & CStr(.SendCnt + 1) & "|" & .BarNo & "||" & .Items(.SendCnt + 1) & "|||||||||||||||||||||0||||||" & vbCr & ETX  '  POMIS
                                                                                                                   '||||||||||||||||||||||O||||||               '  ACK
                
                                          'O|1                         |NORMAL9666    ||^^^0013                     |||||||Q||||^|||||||||||F||||||
                
                .SendCnt = .SendCnt + 1
                
                If .Count = .SendCnt Then
                    intSndPhase = 4
                Else
                    intSndPhase = 3
                End If
            End With
            
            intFrameNo = intFrameNo + 1
            
        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1
            
        Case 5  '## EOT
            strState = ""
            comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1
            
            Exit Sub
    End Select
    
    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    comEqp.Output = strOutput
    Debug.Print strOutput
    SetRawData "[Tx]" & strOutput
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_ADVIA2120()
    Dim strOutput   As String     '송신할 데이터
    
    If msMT = "" Then msMT = Chr(&H30)
    
    msMT = Chr(Asc(msMT) + 1)
    
    If msMT > "Z" Then
        msMT = "0"
    End If
    
    '## Order 없는 경우
    If mOrder.NoOrder = True Then
        strOutput = msMT & "N R " & Format(mOrder.BarNo, "00000000000000") & vbCr & vbLf
    Else
        '해당 Work Order가 있는 경우
        'strOutput = msMT & "Y     " & Format(mOrder.BarNo, "00000000000000") & Space(42)
        strOutput = msMT & "Y" & Space(3) & "A" & Space(1) & Format(mOrder.BarNo, "00000000000000") & Space(42)
        strOutput = strOutput & Space(58)
        strOutput = strOutput & Space(14) & vbCr & vbLf
        strOutput = strOutput & mOrder.Order
        strOutput = strOutput & vbCr & vbLf
    End If
        
    strOutput = CheckSum_ADVIA2120(strOutput)
    
    'Delay 2 sec --> 오더쿼리시간 고려하여 Delay 1 sec
    Call Sleep(1000)
    
    comEqp.Output = strOutput
    SetRawData "[Tx]" & strOutput


End Sub

Private Sub SerialRcvData_ADVIA2120()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strOldBarno        As String   '수신한 바코드번호
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strAspect       As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strKind         As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    Dim sBC$
    Dim iStartPos       As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim strUseRes       As String
    Dim intPosS         As Integer
    
'On Error GoTo RST

    With frmMain
        strRcvBuf = RcvBuffer
        
'        strRcvBuf = "OR 00000009855100 020-08           09/01/17 16:23:19   "
'        strRcvBuf = "  1 3.17   2 2.19   3  5.7   4 15.1   5 68.8   6 26.1   7 37.9  51 38.9   8 16.8   9 1.70  10   58  11 10.1  20 41.9  21 35.8  22 14.7  23  1.3  24  0.9  25  5.3  50 2.74 "
'        strBarno = "9855100"
        
        '-- 테스트용 -----------------
        If .fraCommTest.Visible = False Then
            Call SetSQLData("RCV", strRcvBuf, "A")
        End If
        '-- 테스트용 -----------------
    
        sBC = Mid(strRcvBuf, 2, 1)
    
        Select Case sBC
            Case "S"
                Call TransferToken
                
            Case "Q"    '## Request Information
                strBarno = Trim$(Mid$(strRcvBuf, 4, 14))
                
                If IsNumeric(strBarno) Then
                    strBarno = Val(strBarno)
                End If
                
                sRcvState = "Q"
                
                With mOrder
                    .NoOrder = False
                    .BarNo = strBarno
                End With
                
                Call GetOrder_ADVIA2120(strBarno, gHOSP.RSTTYPE)
                
                Call SendOrder_ADVIA2120
                
                .spdQcResult.MaxRows = 0
            Case "R"
                sRcvState = "R"
                
                strBarno = Mid(strRcvBuf, 4, 14)
                strRackNo = Mid(strRcvBuf, 19, 3)
                strTubePos = Mid(strRcvBuf, 23, 2)
                                    
                If IsNumeric(strBarno) Then
                    strBarno = Val(strBarno)
                End If
                                                    
                strKind = strQCFlag("HEMO", strBarno)
                
                With mResult
                    .BarNo = strBarno
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .Kind = strKind
                    .Rerun = ""
                    If strOldBarno <> strBarno Then
                        strOldBarno = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
                    End If
                End With
                
                .spdQcResult.MaxRows = 0
                
                'iStartPos : Result 시작 위치
                iStartPos = InStr(strRcvBuf, vbCr)
                If iStartPos = 0 Then
                    Exit Sub
                End If
                
                iStartPos = iStartPos + 2
                
                For i = 1 To mc_iMaxCnt
                    strTemp1 = Mid(strRcvBuf, iStartPos + 9 * (i - 1), 1)
                    
                    If strTemp1 = vbCr Then Exit For
                
                    ' 특수문자 제거
                    For j = 1 To Len(strRcvBuf)
                        intPosS = InStr(strRcvBuf, "|")
                        If intPosS > 0 Then
                            strRcvBuf = Replace(strRcvBuf, "|" & Mid(Mid(strRcvBuf, intPosS + 1), 1, InStr(Mid(strRcvBuf, intPosS + 1), "|")), "")
                        Else
                            Exit For
                        End If
                    Next
                                                
                    strIntBase = CStr(Val(Trim(Mid(strRcvBuf, iStartPos + 9 * (i - 1), 3))))
                    strResult = Trim(Mid(strRcvBuf, iStartPos + 9 * (i - 1) + 3, 5))
                    strFlag = Trim(Mid(strRcvBuf, iStartPos + 9 * (i - 1) + 3 + 5, 1))
                    
                    If Left(strResult, 1) = "." Then
                        strResult = "0" & strResult
                    End If
                
                    If strIntBase <> "" And strResult <> "" And IsNumeric(strResult) Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '-- 소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strUseRes <> "" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                '-- BIORAD QC 저장
'                                If mResult.Kind = "QC" Then
'                                    strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                    Call SendBioRadQC(strQCData)
'                                End If
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                
'                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
'                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
'                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
'                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
'                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
'                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
'                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '-- 소수점 처리, 결과 형태 처리
                                If strUseRes <> "" Then
                                    strMachResult = strResult
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                '-- BIORAD QC 저장
                                If mResult.Kind = "QC" And strQCAnalyte <> "" Then
                                    strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult, strQCMethod)
                                    'Call SendBioRadQC(strQCData)
                                End If
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                            
                        End If
                        
                    End If
                Next
                
                '-- BIORAD QC 저장
                If mResult.Kind = "QC" Then
                    If .spdQcResult.MaxRows > 0 Then
                        strQCData = ""
                        For i = 1 To .spdQcResult.MaxRows
                            For j = 1 To 16
                                strQCData = strQCData & Trim(GetText(.spdQcResult, i, j)) & "|"
                            Next
                            strQCData = strQCData & vbCrLf
                        Next
                        If strQCData <> "" Then
                            Call SendBioRadQC(strQCData)
                        End If
                    End If
                End If
                                
                Call TransferResultValMsg
                
                .spdResult.RowHeight(-1) = 14
            
                '## DB에 결과저장
                If .optTrans(0).Value = True And strState = "R" Then
                    Res = SaveTransData_MCC(gRow)
                    
                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "Failed", gRow, colSTATE
                    Else
                        '-- 저장 성공
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "저장완료", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX
                        
                              SQL = "Update PATRESULT Set " & vbCrLf
                        SQL = SQL & " sendflag = '2' " & vbCrLf
                        SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                        SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                    End If
                    strState = ""
                End If
        End Select
    End With
'Exit Sub
'RST:
'
'    Screen.MousePointer = vbDefault
'    MsgBox " Error No. : " & Err.Number & vbCrLf & _
'            " Description : " & Err.Description & vbCrLf & _
'            " Source : " & Err.Source & vbCrLf & vbCrLf
'
End Sub

Private Sub TransferToken()
    Dim sSendBuf$
    
    '프로그램 종료 예정 --> 장비쪽이 Slave인 상태에서 종료되도록 TransferToken 안함
    If mp_bReserveEnd Then
        frmMain.comEqp.PortOpen = False
        mp_bPortOpen = False
        Exit Sub
    End If
    
    If msMT = "" Then msMT = Chr(&H30)
    
    msMT = Chr(Asc(msMT) + 1)
   
    If msMT > "Z" Then
        msMT = "0"
    End If
    
    sSendBuf = msMT & "S " & vbCr & vbLf
    
    sSendBuf = CheckSum_ADVIA2120(sSendBuf)
    
    msSndPacket = sSendBuf

    Timer2120.Interval = 5000
    Timer2120.Enabled = True

End Sub


Private Sub TransferResultValMsg()
    'Result Validation Message
    
    Dim sSendBuf$
    
    If msMT = "" Then msMT = Chr(&H30)

    msMT = Chr(Asc(msMT) + 1)
        
    If msMT > "Z" Then
        msMT = "0"
    End If
    
    sSendBuf = msMT & "Z                  0" & vbCr & vbLf
    sSendBuf = CheckSum_ADVIA2120(sSendBuf)
    
    'Delay 1.5 sec --> 결과등록시간 고려하여 Delay 1 sec
    Call Sleep(1000)
    
    comEqp.Output = sSendBuf
    SetRawData "[Tx]" & sSendBuf

End Sub

Private Sub VesMatic(asData As String)
    Dim strHeader As String
    
    Call SetSQLData("RCV", asData, "A")
    
    strHeader = Trim(mGetP(asData, 1, "="))
    
    If strHeader <> "" And IsNumeric(strHeader) Then
        Call SerialRcvData_VESCUBE
    Else
        Exit Sub
    End If
    
End Sub

Public Sub SndMore()
    Dim strSndMsg As String
    
    strSndMsg = ">"
    strSndMsg = STX & strSndMsg & ETX ' & GetChkSum(strSndMsg) & vbCr
    'strSndMsg = strSndMsg & vbCrLf
    
    comEqp.Output = strSndMsg
    
    SetRawData "[Tx]" & strSndMsg
    
End Sub

Public Sub SndRec()
    Dim strSndMsg As String
    
    strSndMsg = "A"
    strSndMsg = STX & strSndMsg & ETX '& GetChkSum(strSndMsg)
    'strSndMsg = strSndMsg & vbCrLf
    
    comEqp.Output = strSndMsg
    
    SetRawData "[Tx]" & strSndMsg
    
End Sub

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_HITACHI7180(argEquipCode As String, argPID As String, Optional intRow As Long) As String
'    Dim i As Integer
'    Dim sExamCode As String
'    Dim strExamCode As String
'    Dim sSpecNo     As String
'    Dim iRow        As Long
'    Dim SpecNo      As String
    
    Dim lngIntBase  As Long
    Dim strItems    As String           '전송할 검사항목
    Dim blnISE      As Boolean          'Na, K, Cl 검사여부
    
    GetEquipExamCode_HITACHI7180 = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    strItems = String$(88, "0")
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   And TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    Call SetSQLData("채널조회", SQL)
    
    'strExamCode = ""
    mOrder.SendCnt = 0
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            lngIntBase = CLng(AdoRs_Local.Fields("SENDCHANNEL").Value)
            
            '## 계산항목: 93~100
            If lngIntBase >= 93 And lngIntBase <= 100 Then
                'GoTo Skip1
            Else
                '## Na, K, Cl 검사여부 Check
                If lngIntBase = 87 Or lngIntBase = 88 Or lngIntBase = 89 Then
                    blnISE = True
                Else
                    Mid$(strItems, lngIntBase, 1) = "1"
                End If
            End If
            
            '## TIBC 이면 UIBC,FE 오더를 준다.
            If lngIntBase = 98 Then
                Mid$(strItems, 22, 1) = "1"     'FE
                Mid$(strItems, 23, 1) = "1"     'UIBC
            End If
            
            
            '## B/C  (025)항목은 계산항목이라 오더를 보내면 안됨(BUN,CREA)
            '## A/G  (026)항목은 계산항목이라 오더를 보내면 안됨
            '## GLOB (032)항목은 계산항목이라 오더를 보내면 안됨
            '## I.Bil(033)항목은 계산항목이라 오더를 보내면 안됨
            '## T.P  (002)항목은 검체가 Urine일때 검사를 하면 안됨
            '## HbA1C(23)항목은 Hgb(20)와 A1C(21) 오더를 보내야 함
            
            '## LDL-C(21)항목은 계산항목이라 오더를 보내면 안됨(CHOL, T.G, HDL-C)
            If lngIntBase = "21" Then
            '    Mid$(strItems, 21, 1) = "0"     'LDL
                'strResult = strTC - ((strTG / 5) + strHDL)
                Mid$(strItems, 11, 1) = "1"     'T-CHOL
                Mid$(strItems, 12, 1) = "1"     'TG
                Mid$(strItems, 13, 1) = "1"     'HDH-C
            End If
            
            mOrder.SendCnt = mOrder.SendCnt + 1
            AdoRs_Local.MoveNext
        Loop
    End If
    
    '## Na, K, Cl 검사여부 Check
    If blnISE Then
        Mid$(strItems, 87, 1) = "1"
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_HITACHI7180 = strItems
    
End Function

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_HITACHI7080(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim lngIntBase  As Long
    Dim strItems    As String           '전송할 검사항목
    Dim blnISE      As Boolean          'Na, K, Cl 검사여부
    
    GetEquipExamCode_HITACHI7080 = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    strItems = String$(48, "0")
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   And TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    Call SetSQLData("채널조회", SQL)
    
    'strExamCode = ""
    mOrder.SendCnt = 0
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            lngIntBase = CLng(AdoRs_Local.Fields("SENDCHANNEL").Value)
            'LDL
            If lngIntBase <> 99 Then
                Mid$(strItems, lngIntBase, 1) = "1"
            End If
            
            mOrder.SendCnt = mOrder.SendCnt + 1
            AdoRs_Local.MoveNext
        Loop
    End If
    
    
    AdoRs_Local.Close
    
    GetEquipExamCode_HITACHI7080 = strItems
    
End Function

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_AU480(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim lngIntBase  As Long
    Dim strItems    As String           '전송할 검사항목
    Dim blnISE      As Boolean          'Na, K, Cl 검사여부
    
    GetEquipExamCode_AU480 = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   And TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    Call SetSQLData("채널조회", SQL)
    
    mOrder.SendCnt = 0
    strItems = ""
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strItems = strItems & "0" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "" & "0")
            'strItems = strItems & Format(Trim(AdoRs_Local.Fields("SENDCHANNEL").Value), "000")
            mOrder.SendCnt = mOrder.SendCnt + 1
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_AU480 = strItems
    
End Function

Private Sub GetOrder_HITACHI7180(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String
    
    intRow = -1
    GetOrder = ""
    
    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If
        
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, Trim(mOrder.Seq), intRow, colSEQNO)
        'Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        'Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
            
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        Call Sleep(200)
        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_HITACHI7180(gHOSP.MACHCD, pBarno, intRow)
        
        mOrder.Func = Replace(mOrder.Func, String(13, "#"), Left(mOrder.BarNo & Space(13), 13))
        
        '-- 검사채널로 장비오더 만들기
        'If Trim(strItems) = "" Then
        If mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            GetOrder = STX & ";" & mOrder.Func & " 88" & mOrder.Order & "100000" & Left(mOrder.PID & Space(30), 30) & ETX  '& vbCrLf
            'GetOrder = STX & ";" & mOrder.Func & " 88" & mOrder.Order & "100000" & Space(30) & ETX  '& vbCrLf
            
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            GetOrder = STX & ";" & mOrder.Func & " 88" & mOrder.Order & "100000" & Left(mOrder.PID & Space(30), 30) & ETX  '& vbCrLf
            'GetOrder = STX & ";" & mOrder.Func & " 88" & mOrder.Order & "100000" & Space(30) & ETX  '& vbCrLf
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If
        
        comEqp.Output = GetOrder
        SetRawData "[Tx]" & GetOrder
        
        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub

Private Sub GetOrder_HITACHI7080(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String
    
    intRow = -1
    GetOrder = ""
    
    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If
        
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, Trim(mOrder.Seq), intRow, colSEQNO)
        'Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        'Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
            
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_HITACHI7080(gHOSP.MACHCD, pBarno, intRow)
        
        '-- 검사채널로 장비오더 만들기
        If mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
            
        End If
        
        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub

Private Sub GetOrder_AU480(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String
    
    intRow = -1
    GetOrder = ""
    
    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If
        
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, Trim(mOrder.Seq), intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
            
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_AU480(gHOSP.MACHCD, pBarno, intRow)
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            'GetOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.SmpType & mOrder.Seq & Space(26 - Len(mOrder.OrgBarNo)) & mOrder.OrgBarNo & "    E" & ETX
            GetOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.SmpType & mOrder.Seq & Space(26 - Len(mOrder.OrgBarNo)) & mOrder.OrgBarNo & "    E" & ETX
            
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            'GetOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.SmpType & mOrder.Seq & Space(26 - Len(mOrder.OrgBarNo)) & mOrder.OrgBarNo & Space(4) & "E" & strItems & ETX
            GetOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.SmpType & mOrder.Seq & Space(26 - Len(mOrder.OrgBarNo)) & mOrder.OrgBarNo & Space(4) & "E" & strItems & ETX
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If
        
        Call Sleep(500)
        
        comEqp.Output = GetOrder
        
        SetRawData "[Tx]" & GetOrder
        
        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub

Private Sub SerialRcvData_HITACHI7180()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strFunction     As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strFunc         As String
    Dim i               As Integer
    Dim strQCTemp       As String
    

    Dim strTC           As String
    Dim strTG           As String
    Dim strHDL          As String
    
    Dim sFunc           As String
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 1, 1)
            
            Select Case strType
                Case ">", "?", "@"      'ANY 수신
                    Call SndMore
                    'Do
                    'Loop Until comEqp.OutBufferCount = 0
                                    
                Case ";"    '## TS inquiry
                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
                    sFunc = Mid(strRcvBuf, 2, 12) & String(13, "#") & Mid(strRcvBuf, 27, 15)
                    sFunc = Mid(strRcvBuf, 2, 40)
                    With mOrder
                        .BarNo = strBarno
                        '.Func = Mid$(strRcvBuf, 2, 2)
                        .Func = sFunc
                        .Function = Mid$(strRcvBuf, 4, 38)
                        .Seq = Mid(strRcvBuf, 4, 5)
                        .RackNo = Mid$(strRcvBuf, 9, 1)
                        .TubePos = Mid$(strRcvBuf, 10, 3)
                        'tmpSeqNo = Mid(RcvBuffer, 4, 5)
                    End With
                    
                    Call GetOrder_HITACHI7180(Trim$(strBarno), gHOSP.RSTTYPE)
                    
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9" '##Result
                    Call SndMore
                    
                Case ":"    '## End
                    '## Control, Calibration 데이터는 무시함
                    strFunc = Mid$(strRcvBuf, 2, 1)
                    If UCase(strFunc) = "H" Or UCase(strFunc) = "G" Or UCase(strFunc) = "F" Then
                        Call SndMore        'MOR Send
                        strState = ""
                        Exit Sub
                    End If
                    
'                    strRackNo = Mid(strRcvBuf, 9, 1)
'                    strTubePos = Trim(Mid(strRcvBuf, 10, 2))
'                    mOrder.Seq = strTubePos
'                    strBarno = Trim(Mid(strRcvBuf, 13, 13))
                    
                    strSeq = Trim(Mid(strRcvBuf, 4, 5))
                    strRackNo = Trim(Mid(strRcvBuf, 9, 1))
                    strTubePos = Trim(Mid(strRcvBuf, 10, 3))
                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
                    
'                tmpSeqNo = Trim(Mid(RcvBuffer, 4, 5))
'                tmpRack = Trim(Mid(RcvBuffer, 9, 1))
'                tmpPos = Trim(Mid(RcvBuffer, 10, 3))
'                tmpID = Trim(Mid(RcvBuffer, 14, 13))
                
                    With mResult
                        .Seq = strSeq
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With
            
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                                    
                    If gRow <= 0 Then
                        '## Mor 전송
                        Call SndMore
                        Exit Sub
                    End If
                    '123456789012345678901234567890123456789012345678901234567890
                    ':n1    10  11                 01212171406       10  3    20   4    20   5   100  10    91  11   188  12   210  13    62  15   0.9  96    33  99    84

                    For i = 51 To Len(strRcvBuf) Step 10
                        strIntBase = Trim(Mid(strRcvBuf, i, 3))
                        strResult = Trim(Mid(strRcvBuf, i + 3, 6))
                        'strComm = Trim(Mid(strRcvBuf, 9, 1))
                        
'                        If strIntBase = "11" Then    'TCHO
'                            strTC = strResult
'                        End If
'
'                        If strIntBase = "12" Then   'TG
'                            strTG = strResult
'                        End If
'
'                        If strIntBase = "13" Then    'HDLC
'                            strHDL = strResult
'                        End If
                        
                        If strIntBase <> "" And strResult <> "" Then
                            If gPatOrdCd <> "" Then
                                SQL = ""
                                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                                SQL = SQL & "  FROM EQPMASTER" & vbCr
                                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                                
                                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                    strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                    strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                    strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                    strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
            
                                    '-- 결과Row 추가
                                    strRstRow = .spdResult.DataRowCnt + 1
                                    If .spdResult.MaxRows < strRstRow Then
                                        .spdResult.MaxRows = strRstRow
                                    End If
            
                                    '소수점 처리, 결과 형태 처리
                                    strMachResult = strResult
                                    If strQCTemp = "1" Then
                                        strResult = SetResult(strResult, strIntBase)
                                    End If
                                    strJudge = SetJudge(strResult, strIntBase)
                                    
                                    '진행상태 표시("결과")
                                    SetText .spdOrder, "결과", gRow, colSTATE
            
                                    '결과값 표시
                                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                        If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                            SetText .spdOrder, strResult, gRow, intCol
                                            Exit For
                                        End If
                                    Next
            
                                    '-- 결과 List
                                    SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                    SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                    SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                    SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                    SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                    SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                    SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                    SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                    
                                    '-- 로컬 저장
                                    SetLocalDB gRow, strRstRow, "1", ""
                                    
                                    strState = "R"
                                    
                                    '-- 결과Count
                                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                        SetText .spdOrder, "1", gRow, colRCNT
                                    Else
                                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                    End If
                                    
                                End If
                            Else
                                SQL = ""
                                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                                SQL = SQL & "  FROM EQPMASTER" & vbCr
                                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                                
                                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                    strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                    strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                    strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                    strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
            
                                    '-- 결과Row 추가
                                    strRstRow = .spdResult.DataRowCnt + 1
                                    If .spdResult.MaxRows < strRstRow Then
                                        .spdResult.MaxRows = strRstRow
                                    End If
            
                                    '소수점 처리, 결과 형태 처리
                                    strMachResult = strResult
                                    If strQCTemp = "1" Then
                                        strResult = SetResult(strResult, strIntBase)
                                    End If
                                    strJudge = SetJudge(strResult, strIntBase)
                                    
                                    '진행상태 표시("결과")
                                    SetText .spdOrder, "결과", gRow, colSTATE
            
                                    '결과값 표시
                                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                        If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                            SetText .spdOrder, strResult, gRow, intCol
                                            Exit For
                                        End If
                                    Next
            
                                    '-- 결과 List
                                    SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                    SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                    SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                    SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                    SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                    SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                    SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                    SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                    
                                    '-- 로컬 저장
                                    SetLocalDB gRow, strRstRow, "1", ""
                                    
                                    If strState <> "R" Then
                                        strState = ""
                                    End If
            
                                    '-- 결과Count
                                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                        SetText .spdOrder, "1", gRow, colRCNT
                                    Else
                                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                    End If
                                End If
                            End If
                        End If
                    Next
                        
                    Call SndMore
                        
                    'LDL 계산
'                    If strTC <> "" And strTG <> "" And strHDL <> "" And IsNumeric(strTC) And IsNumeric(strTG) And IsNumeric(strHDL) Then
'                        strIntBase = "99"
'                        strResult = strTC - ((strTG / 5) + strHDL)
'                        If strResult < 0 Then
'                            strResult = "0"
'                        End If
'                        strTC = ""
'                        strTG = ""
'                        strHDL = ""
'
'                        If gPatOrdCd <> "" Then
'                            SQL = ""
'                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
'                            SQL = SQL & "  FROM EQPMASTER" & vbCr
'                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
'                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
'
'                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
'                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
'
'                                '-- 결과Row 추가
'                                strRstRow = .spdResult.DataRowCnt + 1
'                                If .spdResult.MaxRows < strRstRow Then
'                                    .spdResult.MaxRows = strRstRow
'                                End If
'
'                                '소수점 처리, 결과 형태 처리
'                                strMachResult = strResult
'                                If strQCTemp = "1" Then
'                                    strResult = SetResult(strResult, strIntBase)
'                                End If
'                                strJudge = SetJudge(strResult, strIntBase)
'
'                                '진행상태 표시("결과")
'                                SetText .spdOrder, "결과", gRow, colSTATE
'
'                                '결과값 표시
'                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                        SetText .spdOrder, strResult, gRow, intCol
'                                        Exit For
'                                    End If
'                                Next
'
'                                '-- 결과 List
'                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
'                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
'                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
'                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
'                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
'                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
'                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
'                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
'
'                                '-- 로컬 저장
'                                SetLocalDB gRow, strRstRow, "1", ""
'
'                                strState = "R"
'
'                                '-- 결과Count
'                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                                    SetText .spdOrder, "1", gRow, colRCNT
'                                Else
'                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                                End If
'
'                            End If
'                        Else
'                            SQL = ""
'                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
'                            SQL = SQL & "  FROM EQPMASTER" & vbCr
'                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
'
'                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
'                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
'
'                                '-- 결과Row 추가
'                                strRstRow = .spdResult.DataRowCnt + 1
'                                If .spdResult.MaxRows < strRstRow Then
'                                    .spdResult.MaxRows = strRstRow
'                                End If
'
'                                '소수점 처리, 결과 형태 처리
'                                strMachResult = strResult
'                                If strQCTemp = "1" Then
'                                    strResult = SetResult(strResult, strIntBase)
'                                End If
'                                strJudge = SetJudge(strResult, strIntBase)
'
'                                '진행상태 표시("결과")
'                                SetText .spdOrder, "결과", gRow, colSTATE
'
'                                '결과값 표시
'                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                        SetText .spdOrder, strResult, gRow, intCol
'                                        Exit For
'                                    End If
'                                Next
'
'                                '-- 결과 List
'                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
'                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
'                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
'                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
'                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
'                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
'                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
'                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
'
'                                '-- 로컬 저장
'                                SetLocalDB gRow, strRstRow, "1", ""
'
'                                If strState <> "R" Then
'                                    strState = ""
'                                End If
'
'                                '-- 결과Count
'                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                                    SetText .spdOrder, "1", gRow, colRCNT
'                                Else
'                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                                End If
'                            End If
'                        End If
'                    End If
                    
                    .spdResult.RowHeight(-1) = 14
                        
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_PLIS(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            
            End Select
        Next
    End With

End Sub


'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_HITACHI7080()
    Dim strOutput   As String     '송신할 데이터
    
    strOutput = ";" & mOrder.Function
    strOutput = strOutput & "48"
    strOutput = strOutput & Mid(mOrder.Order, 1, 48)
    strOutput = strOutput & "000000"
    
    'COMMENT란에 BARCODE 표시
    'strOutput = strOutput & "100000" & Left(mOrder.BarNo & Space(30), 30)
    
    Call Sleep(100)
    
    '-- SPE Send(오더전송)
    comEqp.Output = STX & strOutput & ETX '& vbCr & vbLf
    
    SetRawData "[Tx]" & STX & strOutput & ETX '& vbCr & vbLf

End Sub



Private Sub SerialRcvData_HITACHI7080()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strFunction     As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestCodeSub   As String   '검사코드SUB
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strFunc         As String
    Dim i               As Integer
    Dim strQCTemp       As String
    

    Dim strTC           As String
    Dim strTG           As String
    Dim strHDL          As String
    
    Dim sFunc           As String
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 1, 1)
            
            Select Case strType
                Case ">", "?", "@"      'ANY 수신
                    Call SndMore
                    'Do
                    'Loop Until comEqp.OutBufferCount = 0
                
                Case "?", "@"           'REP 수신
                    Sleep (100)
                    Call SndMore        'MOR Send
                    Do
                    '   DoEvents
                    Loop Until frmMain.comEqp.OutBufferCount = 0
                    
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
                    Call SndMore
                                    
                Case ";"    '## TS inquiry
                    strFunc = Mid(strRcvBuf, 2, 39)
                    strSeq = Mid(strRcvBuf, 4, 5)
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Mid(strRcvBuf, 10, 2)
                    strBarno = Trim(Mid(strRcvBuf, 13, 13))
                    
                    With mOrder
                        .Seq = strSeq
                        .BarNo = strBarno
                        .Func = Mid$(strRcvBuf, 2, 1)
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                    End With
                    
                    Call GetOrder_HITACHI7080(Trim$(strBarno), gHOSP.RSTTYPE)
                    
                    Mid(strFunc, 17, 8) = mOrder.BarNo
                    
                    mOrder.Function = strFunc
                    
                    Call SendOrder_HITACHI7080
                    
                    Call SetText(frmMain.spdOrder, "0", gRow, colCHECKBOX)
                                        
                Case ":"    '## End
                    '## Control, Calibration 데이터는 무시함
                    strFunc = Mid$(strRcvBuf, 2, 1)
                    If strFunc = "K" Or strFunc = "L" Or strFunc = "G" Or strFunc = "H" Then
                        Sleep (100)
                        Call SndMore        'MOR Send
                        Do
                        '   DoEvents
                        Loop Until comEqp.OutBufferCount = 0
                        Exit Sub
                    End If
                    
                    If strFunc = "K" Or strFunc = "L" Then
                        Call SndMore        'MOR Send
                        Exit Sub
                    End If
                    
                    Call SndMore            'MOR Send
            
                    If strFunc <> "@" And strFunc <> "M" Then
            
                        strRackNo = Mid(strRcvBuf, 9, 1)
                        strTubePos = Trim(Mid(strRcvBuf, 10, 2))
                        mOrder.Seq = strTubePos
                        strBarno = Trim(Mid(strRcvBuf, 13, 13))
                        
                        With mResult
                            .Seq = strTubePos
                            .BarNo = strBarno
                            .RackNo = strRackNo
                            .TubePos = strTubePos
                            .RsltDate = Format(Now, "yyyymmddhhmmss")
                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                        End With
                
                        Call SetPatInfo_H7080(strBarno, gHOSP.RSTTYPE)
            
                        If gRow <= 0 Then
                            '## Mor 전송
                            Call SndMore
                            Exit Sub
                        End If
                        
                        ':f 01003  1      Biorad 1    0                15 1   6.8  2   4.0  3    45  4    35  5   103  6   1.3  8    68 10   177 11  15.4 12   2.7 13   259 14    83 15    69 16   140 17   5.5

                        strTmp = Mid$(strRcvBuf, 49)
        
                        'For i = 44 To Len(strRcvBuf) Step 10
                        Do While Len(strTmp) >= 9
                            strIntBase = Trim(Mid(strTmp, 1, 2))
                            strResult = Trim(Mid(strTmp, 3, 6))
                            strComm = Trim(Mid(strTmp, 9, 1))
                            
                            If strIntBase = "6" Then    'TCHO
                                strTC = strResult
                            End If

                            If strIntBase = "13" Then   'TG
                                strTG = strResult
                            End If

                            If strIntBase = "11" Then    'HDLC
                                strHDL = strResult
                            End If
                            
                            If strIntBase <> "" And strResult <> "" Then
                                If gPatOrdCd <> "" Then
                                    SQL = ""
                                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                                    SQL = SQL & "  FROM EQPMASTER" & vbCr
                                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                                    SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                                    
                                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                        strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                        strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                        strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                        strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
                
                                        '-- 결과Row 추가
                                        strRstRow = .spdResult.DataRowCnt + 1
                                        If .spdResult.MaxRows < strRstRow Then
                                            .spdResult.MaxRows = strRstRow
                                        End If
                
                                        '소수점 처리, 결과 형태 처리
                                        strMachResult = strResult
                                        If strQCTemp = "1" Then
                                            strResult = SetResult(strResult, strIntBase)
                                        End If
                                        strJudge = SetJudge(strResult, strIntBase)
                                        
                                        '진행상태 표시("결과")
                                        SetText .spdOrder, "결과", gRow, colSTATE
                
                                        '결과값 표시
                                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                            If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                                SetText .spdOrder, strResult, gRow, intCol
                                                strTestCodeSub = gArrEQP(intCol - colSTATE, 7)
                                                Exit For
                                            End If
                                        Next
                
                                        '-- 결과 List
                                        SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                        SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                        SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                        SetText .spdResult, strTestCodeSub, strRstRow, colRSUBCD          '검사코드SUB
                                        SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                        SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                        SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                        SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                        SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                        
                                        '-- 로컬 저장
                                        SetLocalDB gRow, strRstRow, "1", ""
                                        
                                        strState = "R"
                                        
                                        '-- 결과Count
                                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                            SetText .spdOrder, "1", gRow, colRCNT
                                        Else
                                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                        End If
                                        
                                    End If
                                Else
                                    SQL = ""
                                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                                    SQL = SQL & "  FROM EQPMASTER" & vbCr
                                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                                    
                                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                        strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                        strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                        strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                        strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
                
                                        '-- 결과Row 추가
                                        strRstRow = .spdResult.DataRowCnt + 1
                                        If .spdResult.MaxRows < strRstRow Then
                                            .spdResult.MaxRows = strRstRow
                                        End If
                
                                        '소수점 처리, 결과 형태 처리
                                        strMachResult = strResult
                                        If strQCTemp = "1" Then
                                            strResult = SetResult(strResult, strIntBase)
                                        End If
                                        strJudge = SetJudge(strResult, strIntBase)
                                        
                                        '진행상태 표시("결과")
                                        SetText .spdOrder, "결과", gRow, colSTATE
                
                                        '결과값 표시
                                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                            If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                                SetText .spdOrder, strResult, gRow, intCol
                                                strTestCodeSub = gArrEQP(intCol - colSTATE, 7)
                                                Exit For
                                            End If
                                        Next
                
                                        '-- 결과 List
                                        SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                        SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                        SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                        SetText .spdResult, strTestCodeSub, strRstRow, colRSUBCD          '검사코드SUB
                                        SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                        SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                        SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                        SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                        SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                        
                                        '-- 로컬 저장
                                        SetLocalDB gRow, strRstRow, "1", ""
                                        
                                        If strState <> "R" Then
                                            strState = ""
                                        End If
                
                                        '-- 결과Count
                                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                            SetText .spdOrder, "1", gRow, colRCNT
                                        Else
                                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                        End If
                                    End If
                                End If
                            End If
                            strTmp = Mid$(strTmp, 10)
                        Loop
                        
                        'Next
                    
                        .spdResult.RowHeight(-1) = 14
                            
                        'LDL 계산
                        If strTC <> "" And strTG <> "" And strHDL <> "" And IsNumeric(strTC) And IsNumeric(strTG) And IsNumeric(strHDL) Then
                            strIntBase = "99"
                            strResult = strTC - ((strTG / 5) + strHDL)
                            If strResult < 0 Then
                                strResult = "0"
                            End If
                            strTC = ""
                            strTG = ""
                            strHDL = ""

                            If gPatOrdCd <> "" Then
                                SQL = ""
                                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                                SQL = SQL & "  FROM EQPMASTER" & vbCr
                                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "

                                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                    strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                    strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                    strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                    strQCTemp = Trim(RS_L.Fields("QCTemp") & "")

                                    '-- 결과Row 추가
                                    strRstRow = .spdResult.DataRowCnt + 1
                                    If .spdResult.MaxRows < strRstRow Then
                                        .spdResult.MaxRows = strRstRow
                                    End If

                                    '소수점 처리, 결과 형태 처리
                                    strMachResult = strResult
                                    If strQCTemp = "1" Then
                                        strResult = SetResult(strResult, strIntBase)
                                    End If
                                    strJudge = SetJudge(strResult, strIntBase)

                                    '진행상태 표시("결과")
                                    SetText .spdOrder, "결과", gRow, colSTATE

                                    '결과값 표시
                                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                        If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                            SetText .spdOrder, strResult, gRow, intCol
                                            Exit For
                                        End If
                                    Next

                                    '-- 결과 List
                                    SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                    SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                    SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                    SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                    SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                    SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                    SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                    SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치

                                    '-- 로컬 저장
                                    SetLocalDB gRow, strRstRow, "1", ""

                                    strState = "R"

                                    '-- 결과Count
                                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                        SetText .spdOrder, "1", gRow, colRCNT
                                    Else
                                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                    End If

                                End If
                            Else
                                SQL = ""
                                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                                SQL = SQL & "  FROM EQPMASTER" & vbCr
                                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "

                                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                    strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                    strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                    strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                    strQCTemp = Trim(RS_L.Fields("QCTemp") & "")

                                    '-- 결과Row 추가
                                    strRstRow = .spdResult.DataRowCnt + 1
                                    If .spdResult.MaxRows < strRstRow Then
                                        .spdResult.MaxRows = strRstRow
                                    End If

                                    '소수점 처리, 결과 형태 처리
                                    strMachResult = strResult
                                    If strQCTemp = "1" Then
                                        strResult = SetResult(strResult, strIntBase)
                                    End If
                                    strJudge = SetJudge(strResult, strIntBase)

                                    '진행상태 표시("결과")
                                    SetText .spdOrder, "결과", gRow, colSTATE

                                    '결과값 표시
                                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                        If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                            SetText .spdOrder, strResult, gRow, intCol
                                            Exit For
                                        End If
                                    Next

                                    '-- 결과 List
                                    SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                    SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                    SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                    SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                    SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                    SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                    SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                    SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치

                                    '-- 로컬 저장
                                    SetLocalDB gRow, strRstRow, "1", ""

                                    If strState <> "R" Then
                                        strState = ""
                                    End If

                                    '-- 결과Count
                                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                        SetText .spdOrder, "1", gRow, colRCNT
                                    Else
                                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                    End If
                                End If
                            End If
                        End If
                        
                        '## DB에 결과저장
                        If .optTrans(0).Value = True And strState = "R" Then
                            Res = SaveTransData_BIGUBCARE(gRow)
                            
                            If Res = -1 Then
                                '-- 저장 실패
                                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                                SetText .spdOrder, "Failed", gRow, colSTATE
                            Else
                                '-- 저장 성공
                                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                                SetText .spdOrder, "저장완료", gRow, colSTATE
                                SetText .spdOrder, "0", gRow, colCHECKBOX
                                
                                      SQL = "Update PATRESULT Set " & vbCrLf
                                SQL = SQL & " sendflag = '2' " & vbCrLf
                                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                                
                                If DBExec(AdoCn_Local, SQL) Then
                                    '-- 성공
                                End If
                            End If
                            strState = ""
                        End If
                    
                    End If
            End Select
        Next
    End With

End Sub

Private Sub SerialRcvData_VESCUBE()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strOldBarno        As String   '수신한 바코드번호
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strAspect       As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    With frmMain
        strRecvData = Split(RcvBuffer, vbCr)
        For intCnt = 1 To UBound(strRecvData)
            RcvBuffer = strRecvData(intCnt)
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", RcvBuffer, "A")
            End If
            '-- 테스트용 -----------------
                
            '1 = O4ZU70QN0....  48
            '1 = 199297.......   6
            strTemp1 = mGetP(RcvBuffer, 2, "=")
            strBarno = Trim(mGetP(strTemp1, 1, "......."))
            strIntBase = "ESR"
            
            'If Trim(strBarno) <> "" And Len(strBarno) = 6 Then
            If Trim(strBarno) <> "" Then
                With mResult
                    .BarNo = strBarno
                    '.SpcPos = strTubePos & "/" & strRackNo
                    '.Seq = strSeq
                    '.RackNo = strRackNo
                    '.TubePos = strTubePos
                    If strOldBarno <> strBarno Then
                        strOldBarno = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
                    End If
                End With
                
                If gPatOrdCd <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                    SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                    SQL = SQL & "  FROM EQPMASTER" & vbCr
                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                    SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                    
                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                        strTestCode = Trim(RS_L.Fields("TESTCODE"))
                        strTestName = Trim(RS_L.Fields("TESTNAME"))
                        strSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                        '-- 결과Row 추가
                        strRstRow = .spdResult.DataRowCnt + 1
                        If .spdResult.MaxRows < strRstRow Then
                            .spdResult.MaxRows = strRstRow
                        End If
    
                        '소수점 처리, 결과 형태 처리
                        strMachResult = strResult
                        strResult = SetResult(strResult, strIntBase)
                        strJudge = SetJudge(strResult, strIntBase)
                        
                        '진행상태 표시("결과")
                        SetText .spdOrder, "결과", gRow, colSTATE
    
                        '결과값 표시
                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
                            If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                SetText .spdOrder, strResult, gRow, intCol
                                Exit For
                            End If
                        Next
    
                        '-- 결과 List
                        SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                        SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                        SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                        SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                        SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                        SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                        SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                        SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                        
                        '-- 로컬 저장
                        SetLocalDB gRow, strRstRow, "1", ""
                        
                        '-- BIORAD QC 저장
                        'If Mid(strBarno, 1, 2) = "QC" Then
                        '    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
                        'End If
                    
                        
                        strState = "R"
                        
                        '-- 결과Count
                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
                            SetText .spdOrder, "1", gRow, colRCNT
                        Else
                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                        End If
                        
                    End If
                Else
                    SQL = ""
                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                    SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                    SQL = SQL & "  FROM EQPMASTER" & vbCr
                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                    
                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                        strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                        strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                        strSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                        'strQCLab = Trim(RS_L.Fields("QCLab") & "")
                        'strQCLot = Trim(RS_L.Fields("QCLot") & "")
                        'strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                        'strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                        'strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                        'strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                        'strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                        'strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                        '-- 결과Row 추가
                        strRstRow = .spdResult.DataRowCnt + 1
                        If .spdResult.MaxRows < strRstRow Then
                            .spdResult.MaxRows = strRstRow
                        End If
    
                        '소수점 처리, 결과 형태 처리
                        strMachResult = strResult
                        strResult = SetResult(strResult, strIntBase)
                        strJudge = SetJudge(strResult, strIntBase)
                        
                        '진행상태 표시("결과")
                        SetText .spdOrder, "결과", gRow, colSTATE
    
                        '결과값 표시
                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
                            If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                SetText .spdOrder, strResult, gRow, intCol
                                Exit For
                            End If
                        Next
    
                        '-- 결과 List
                        SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                        SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                        SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                        SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                        SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                        SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                        SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                        SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                        
                        '-- 로컬 저장
                        SetLocalDB gRow, strRstRow, "1", ""
                        
                        '-- BIORAD QC 저장
                        'If Mid(strBarno, 1, 2) = "QC" Then
                        '    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
                        'End If
                        
                        If strState <> "R" Then
                            strState = ""
                        End If
    
                        '-- 결과Count
                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
                            SetText .spdOrder, "1", gRow, colRCNT
                        Else
                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                        End If
                    End If
                    
                End If
                
            End If
                            
            .spdResult.RowHeight(-1) = 14
                
            '## DB에 결과저장
            If .optTrans(0).Value = True And strState = "R" Then
                Res = SaveTransData_AMIS(gRow)
                
                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "Failed", gRow, colSTATE
                Else
                    '-- 저장 성공
                    SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                    SetText .spdOrder, "저장완료", gRow, colSTATE
                    SetText .spdOrder, "0", gRow, colCHECKBOX
                    
                          SQL = "Update PATRESULT Set " & vbCrLf
                    SQL = SQL & " sendflag = '2' " & vbCrLf
                    SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                    SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                    SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                    SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                    End If
                End If
                strState = ""
            End If
        Next
    End With

End Sub



Private Sub Phase_Serial_VESCUBE()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)
    TimerVESCUBE.Interval = 5000
    TimerVESCUBE.Enabled = True
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case BufChar
            Case vbLf
            Case vbCr
                    Call VesMatic(RcvBuffer)
                    RcvBuffer = ""
            
            Case Else
                    RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
            
End Sub

Private Sub Phase_Serial_ADVIA2120()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

'    SetRawData "[Rx]" & pBuffer
    lngBufLen = Len(pBuffer)
            
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1                  ' 초기화 확인(MT 대기)
                Select Case BufChar
                    Case Chr(2)          'STX
                        RcvBuffer = ""
                        
                    Case msMT                           ' 0(&H30), Initialize Message 직후이므로
                        RcvBuffer = ""
                        Call TransferToken              'Transfer_Token 시도 후
                        intPhase = 2

                    Case Else
                        'MT이외의 경우(즉, NACK)
                        Call InitialComm
                        intPhase = 1
                        Exit Sub
                End Select

            Case 2                                      ' Token Tranfer에 대한 MT 대기
                Select Case BufChar
                    Case Chr(21)         'NAK
                        Call InitialComm
                        intPhase = 1
                        
                    Case msMT
                        intPhase = 3

                    Case Chr(Asc(msMT) - 1)
                        msMT = Chr(Asc(msMT) - 1)
                        If Asc(msMT) < 0 Then
                            msMT = Chr(&H30)
                        End If
                        Call TransferToken              'Transfer_Token 시도 후
                        intPhase = 2

                    Case Else
                        Call TransferToken
                        intPhase = 2
                End Select

            Case 3                                      ' CheckSum STX인 경우의 오류 방지를 위해, Phase 3 과 4 분리
                Select Case Asc(BufChar)
                    Case 2          'STX
                        RcvBuffer = ""
                        intPhase = 4

                    Case Else
                        intPhase = 3

                End Select

            Case 4                                      'DataEdit(장비쪽에서 보내는 S, Q, R 메세지에 대한) 대기,
                Select Case Asc(BufChar)
                    Case 3            ' ETX
                        msMT = Left(RcvBuffer, 1)
                        Call Sleep(25)  'Delay Time 0.025 sec

                        comEqp.Output = msMT
                        SetRawData "[Tx]" & msMT
                        Call SerialRcvData_ADVIA2120
                        intPhase = 3

                    Case Else
                        RcvBuffer = RcvBuffer & BufChar

                End Select
        End Select
    Next i
    
    
End Sub


'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_VERSACELL()
    Dim strOutput   As String     '송신할 데이터

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&||||62 Flanders-Bartley Road^Flanders^NJ^07921||973-927-2828|N81|||P|1|" & Format(Now, "yyyymmddhhmmss") & "|" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1

        Case 2  '## Patient
            'strOutput = intFrameNo & "P|1|" & mOrder.BarNo & "|||" & frmMain.Han2Eng.HanToEng(mOrder.PName) & "||||" & vbCr & ETX
            strOutput = intFrameNo & "P|1|" & mOrder.BarNo & "|||" & mOrder.PName & "||||" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1

        Case 3  '## Order
            If mOrder.NoOrder = True Then
                '## 접수정보가 없을경우
                strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||||||" & mOrder.SPCCD
                strOutput = intFrameNo & strOutput & vbCr & ETX
                intSndPhase = 4

            Else
                '## 최초 보낼때
                If mOrder.IsSending = False Then
                    strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||||||1"

                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## 남은 문자열이 있을때
                Else
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1

            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    frmMain.comEqp.Output = strOutput
    SetRawData "[Tx]" & strOutput

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_BS240()
    Dim strOutput   As String     '송신할 데이터

    Select Case intSndPhase
        Case 1
            '## 최초 보낼때
'            If mOrder.IsSending = False Then

                strOutput = ""
                strOutput = strOutput & "H|\^&|||Mindry^^|||||||RQ|1394-97|" & Format(Now, "yyyymmddhhmmss") & vbCr
                strOutput = strOutput & "P|1||" & mOrder.BarNo & "||" & mOrder.PName & "||||||A|||||||||||||||||||||||" & vbCr
                strOutput = strOutput & "O|1|" & mOrder.Seq & "||" & mOrder.Order & "|R|" & Format(Now, "yyyymmddhhmmss") & "|" & Format(Now, "yyyymmddhhmmss") & "||||||||||||||||||Q|||||" & vbCr
                strOutput = strOutput & "L|1|N"
                
'                If Len(strOutput) > 230 Then
'                    mOrder.IsSending = True
'                    mOrder.Order = Mid$(strOutput, 231)
'                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'                    intSndPhase = 1
'                Else
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 2
'                End If
            '## 남은 문자열이 있을때
'            Else
'                strOutput = mOrder.Order
'                If Len(strOutput) > 230 Then
'                    mOrder.Order = Mid$(strOutput, 231)
'                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'                    intSndPhase = 1
'                Else
'                    mOrder.IsSending = False
'                    strOutput = intFrameNo & strOutput & vbCr & ETX
'                    intSndPhase = 2
'                End If
'            End If
            intFrameNo = intFrameNo + 1


        Case 2  '## EOT
            strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intSndPhase = 1
            intFrameNo = 1
            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCr & vbLf
    frmMain.comEqp.Output = strOutput
    SetRawData "[Tx]" & strOutput

End Sub

Private Sub Phase_Serial_VERSACELL()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_VERSACELL
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                        intPhase = 4
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        Call SerialRcvData_VERSACELL
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            frmMain.comEqp.Output = ENQ
                            SetRawData "[Tx]" & ENQ
                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
            
End Sub


Private Sub Phase_Serial_ABBOTTRUBY()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case ACK
'                        If strState = "Q" Then
'                            Call SendOrder_VERSACELL
'                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                        intPhase = 4
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        Call SerialRcvData_ABBOTTRUBY
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            frmMain.comEqp.Output = ENQ
                            SetRawData "[Tx]" & ENQ
                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
            
End Sub

Private Sub SerialRcvData_UROMETER720()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strFunction     As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestCodeSub   As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strFunc         As String
    Dim i               As Integer
    Dim strQCTemp       As String
    

    Dim strTC           As String
    Dim strTG           As String
    Dim strHDL          As String
    
'    Call SetSQLData("RCV", RcvBuffer, "A")
    
    With frmMain
        'RcvBuffer = Replace(RcvBuffer, vbLf, "")
        strRecvData = Split(RcvBuffer, vbCrLf)
        
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            Select Case intCnt
                Case 3
                    strSeq = Mid(strRcvBuf, 10)
                    strSeq = Replace(strSeq, ")", "")
                    strSeq = Replace(strSeq, "(", "")
                    strSeq = Val(Trim(strSeq))

            
                    With mResult
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With
            
                    Call SetPatInfo(strSeq, gHOSP.RSTTYPE)

                Case 4 To 13
                    strIntBase = Mid(strRcvBuf, 1, 4)
                    strIntBase = Trim(strIntBase)
                    
                    'strResult = Mid(strRcvBuf, 7, 5) '-- 정성
                    strResult = Mid(strRcvBuf, 8, 4) '-- 정성
                    strResult = Trim(strResult)
            
                    If strIntBase = "pH" Or strIntBase = "p.H" Or strIntBase = "S.G" Then
                        strResult = Trim(Mid(strRcvBuf, 4))  '-- 정량
                        strResult = Replace(strResult, "mg/dl", "")
                        strResult = Replace(strResult, "RBC/ul", "")
                        strResult = Replace(strResult, "WBC/ul", "")
                        
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "=", "")
                    End If
                    
                    '-- URO
                    If strResult = "norm" Then
                        strResult = "-"
                    End If
    '
    '                '-- NIT
                    If strResult = "pos" Then
                        strResult = "+"
                    End If
            
                    Select Case Trim(strResult)
                        Case "+":       strResult = "1+"
                        Case "++":      strResult = "2+"
                        Case "+++":     strResult = "3+"
                        Case "++++":    strResult = "4+"
                        Case "+/-":     strResult = "Trace"
                    End Select

                            
                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        strTestCodeSub = gArrEQP(intCol - colSTATE, 7)
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestCodeSub, strRstRow, colRSUBCD          '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        strTestCodeSub = gArrEQP(intCol - colSTATE, 7)
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestCodeSub, strRstRow, colRSUBCD          '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                        End If
                    End If
                
                Case 14
                            
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_BIGUBCARE(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
                    
            End Select
        Next
    End With

End Sub

Private Sub SerialRcvData_DUREADER720()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strFunction     As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strFunc         As String
    Dim i               As Integer
    Dim strQCTemp       As String
    

    Dim strTC           As String
    Dim strTG           As String
    Dim strHDL          As String
    
    With frmMain
        RcvBuffer = Replace(RcvBuffer, vbLf, "")
        strRecvData = Split(RcvBuffer, vbCr)
        
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            Select Case intCnt
                Case 3
                    strSeq = Mid(strRcvBuf, 12)
                    strSeq = Replace(strSeq, ")", "")
                    strSeq = Replace(strSeq, "(", "")
                    strSeq = Val(Trim(strSeq))

                    strRcvBuf = strRecvData(15)
                    
                    strBarno = Mid(strRcvBuf, 4, 12)
            
                    With mResult
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                        .BarNo = strBarno
                    End With
            
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)

                Case 4 To 13
                    strIntBase = Mid(strRcvBuf, 1, 4)
                    strIntBase = Trim(strIntBase)
                    
                    'strResult = Mid(strRcvBuf, 7, 5) '-- 정성
                    strResult = Mid(strRcvBuf, 8, 4) '-- 정성
                    strResult = Trim(strResult)
            
                    If strIntBase = "pH" Or strIntBase = "p.H" Or strIntBase = "S.G" Then
                        'strResult = Trim(Mid(strRcvBuf, 13))  '-- 정량
                        strResult = Trim(Mid(strRcvBuf, 4))  '-- 정량
                        'strResult = Trim(Mid(strRcvBuf, 12, 7)) '-- 정량
                        strResult = Replace(strResult, "mg/dl", "")
                        strResult = Replace(strResult, "RBC/ul", "")
                        strResult = Replace(strResult, "WBC/ul", "")
                        
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "=", "")
                    End If
                    
                    '-- URO
                    If strResult = "norm" Then
                        strResult = "-"
                    End If
    '
    '                '-- NIT
                    If strResult = "pos" Then
                        strResult = "+"
                    End If
            
                    Select Case Trim(strResult)
                        Case "+":       strResult = "1+"
                        Case "++":      strResult = "2+"
                        Case "+++":     strResult = "3+"
                        Case "++++":    strResult = "4+"
                        Case "+/-":     strResult = "Trace"
                    End Select

                            
                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                        End If
                    End If
                
                Case 14
                            
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_PLIS(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
                    
            End Select
        Next
    End With
    
    Call spdOrder_Click(1, gRow)
    

End Sub


Private Sub Phase_Serial_ETIMAX3000()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case ACK
                        If strState = "Q" Then
'                            Call SendOrder_ETIMAX3000
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                        intPhase = 4
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        Call SerialRcvData_ETIMAX3000
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            frmMain.comEqp.Output = ENQ
                            SetRawData "[Tx]" & ENQ
                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
            
End Sub

Private Sub Phase_Serial_HITACHI7180()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                Erase strRecvData
                
                intBufCnt = 1
                ReDim Preserve strRecvData(intBufCnt)
                
            Case ETX
                Call SerialRcvData_HITACHI7180
                'Erase strRecvData
                
            Case vbCr
            Case vbLf
            Case Else
                If intBufCnt > 0 Then
                    strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                End If
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_HITACHI7080()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                Erase strRecvData
                
                intBufCnt = 1
                ReDim Preserve strRecvData(intBufCnt)
                
            Case ETX
                Call SerialRcvData_HITACHI7080
                'Erase strRecvData
                
            Case vbCr
            Case vbLf
            Case Else
                If intBufCnt > 0 Then
                    strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                End If
        End Select
    Next i
    
End Sub

Private Sub SerialRcvData_UrinscanPro()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strFunction     As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strFunc         As String
    Dim i               As Integer
    Dim strQCTemp       As String
    Dim Pos As Integer
    
    With frmMain
        Pos = InStr(RcvBuffer, "ID_NO")
        If Pos > 0 Then
            RcvBuffer = Replace(RcvBuffer, vbLf, "")
            strRecvData = Split(RcvBuffer, vbCr)
            
            '-- 바코드 번호 찾기
            strRcvBuf = strRecvData(16)
            strBarno = Mid(strRcvBuf, 4, 13)
        
            '-- SEQ 번호 찾기
            strRcvBuf = strRecvData(1)
            strRcvBuf = mGetP(strRcvBuf, 2, ":")
            strRcvBuf = mGetP(strRcvBuf, 1, "-")
            strSeq = Trim(strRcvBuf)
            
            With mResult
                .BarNo = strBarno
                .RsltDate = Format(Now, "yyyymmddhhmmss")
                .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                .TubePos = strSeq
            End With
                    
            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
            
            For intCnt = 4 To UBound(strRecvData)
                strRcvBuf = strRecvData(intCnt)
                
                '-- 테스트용 -----------------
                'If .fraCommTest.Visible = False Then
                '    Call SetSQLData("RCV", strRcvBuf, "A")
                'End If
                '-- 테스트용 -----------------
                
                strType = Trim(Mid$(strRcvBuf, 1, 3))
                strIntBase = strType
                strResult = ""
                
                Select Case strType
                    Case "p.H", "pH", "S.G", "SG", "COL" '## 소숫점 포함 3자리
                        strResult = Trim$(Mid$(strRcvBuf, 4))
                        strResult = Replace(strResult, "mg/dl", "")
                        strResult = Replace(strResult, "RBC/ul", "")
                        strResult = Replace(strResult, "WBC/ul", "")
                        
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "=", "")
                    
                    Case Else
                        strResult = Trim$(Mid$(strRcvBuf, 4, 7))
                        'strResult = Trim(Mid(strRcvBuf, 12))  '-- 정량
                        strResult = Replace(strResult, "mg/dl", "")
                        strResult = Replace(strResult, "RBC/ul", "")
                        strResult = Replace(strResult, "WBC/ul", "")
                        
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "=", "")
                        
                End Select
                
                '기존 해군부대 결과 나가던 형태
'~NOV/28/2009 11:12:10
'Name:             Sex:
'Ward:             Age:
'SD (10)      (0001)
'BLD       -        neg
'BIL       -        neg
'URO    norm    0.1mg/dl
'KET       -        neg
'PRO       -        neg
'NIT       -        neg
'GLU       -        neg
'pH                 7.0
'S.G              1.015
'LEU       -        neg
'(SN=60158391)
'ID(             )
'OP(00000000) LOT(000000) ~


'                Select Case intCnt
'                    Case 3:
'                                    strBarno = Mid(strRcvBuf, 12)
'                                    strBarno = Replace(strBarno, ")", "")
'                                    strBarno = Replace(strBarno, "(", "")
'                                    strBarno = Trim(strBarno)
'                    Case 4 To 13:
'                                    'strResult = Trim(strRcvBuf)
'                                    strEqpcd = Mid(strRcvBuf, 1, 4)
'                                    strEqpcd = Trim(strEqpcd)
'
'                                    strResult = Mid(strRcvBuf, 8, 4) '-- 정성
'                                    strResult = Trim(strResult)
'                                    strResult = strResult & " " & Trim(Mid(strRcvBuf, 15, 12)) '-- 정량
'                                    strResult = Replace(strResult, "mg/dl", "")
'                    'Case Else:  strResult = Trim(strRcvBuf)
'
'                End Select
'
                        
'                Select Case strType
'                    Case "BLD", "BIL", "PRO", "LEU"
'                            '정량
'                            strIntResult = Trim$(Mid$(strRcvBuf, 12, 4))
'                            strIntResult = Replace(strIntResult, "mg/dl", "")
'                            strIntResult = Replace(strIntResult, "RBC/ul", "")
'                            strIntResult = Replace(strIntResult, "WBC/ul", "")
'                            strIntResult = Replace(strIntResult, "<", "")
'                            strIntResult = Replace(strIntResult, ">", "")
'                            strIntResult = Replace(strIntResult, "=", "")
'                            '정성
'                            strResult = Trim$(Mid$(strRcvBuf, 8, 4))
'                            strResult = Replace(strResult, "mg/dl", "")
'                            strResult = Replace(strResult, "RBC/ul", "")
'                            strResult = Replace(strResult, "WBC/ul", "")
'                            strResult = Replace(strResult, "<", "")
'                            strResult = Replace(strResult, ">", "")
'                            strResult = Replace(strResult, "=", "")
'
'                            If IsNumeric(strIntResult) Then
'                                If strResult <> "" Then
'                                    strResult = strIntResult & "(" & strResult & ")"
'                                Else
'                                    strResult = strIntResult
'                                End If
'                            End If
'                    Case "URO"
'                            strResult = Trim$(Mid$(strRcvBuf, 12))
'
'                    Case "p.H", "pH", "S.G", "SG", "COL" '## 소숫점 포함 3자리
'                            strResult = Trim$(Mid$(strRcvBuf, 4))
'                            strResult = Replace(strResult, "mg/dl", "")
'                            strResult = Replace(strResult, "RBC/ul", "")
'                            strResult = Replace(strResult, "WBC/ul", "")
'
'                            strResult = Replace(strResult, "<", "")
'                            strResult = Replace(strResult, ">", "")
'                            strResult = Replace(strResult, "=", "")
'
'                    Case Else
'                            strResult = Trim$(Mid$(strRcvBuf, 4, 7))
'                            'strResult = Trim(Mid(strRcvBuf, 12))  '-- 정량
'                            strResult = Replace(strResult, "mg/dl", "")
'                            strResult = Replace(strResult, "RBC/ul", "")
'                            strResult = Replace(strResult, "WBC/ul", "")
'
'                            strResult = Replace(strResult, "<", "")
'                            strResult = Replace(strResult, ">", "")
'                            strResult = Replace(strResult, "=", "")
'
'                            If Trim(strResult) = "-" Then
'                                strResult = "neg"
'                            End If
'                End Select
    
                If strIntBase <> "" And strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    Else
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            If strState <> "R" Then
                                strState = ""
                            End If
    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    End If
                End If
            Next
                
            '## DB에 결과저장
            If .optTrans(0).Value = True And strState = "R" Then
                Res = SaveTransData_NAVY(gRow)
                
                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "Failed", gRow, colSTATE
                Else
                    '-- 저장 성공
                    SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                    SetText .spdOrder, "저장완료", gRow, colSTATE
                    SetText .spdOrder, "0", gRow, colCHECKBOX
                    
                          SQL = "Update PATRESULT Set " & vbCrLf
                    SQL = SQL & " sendflag = '2' " & vbCrLf
                    SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                    SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                    SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                    SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                    End If
                End If
                strState = ""
            End If
            
            SetText .spdOrder, "0", gRow, colCHECKBOX
            
        End If
    End With

End Sub

Private Sub SerialRcvData_HORIBA()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strFunction     As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestCodeSub   As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strFunc         As String
    Dim i               As Integer
    Dim strQCTemp       As String
    Dim Pos As Integer
    
    With frmMain
        RcvBuffer = Replace(RcvBuffer, vbLf, "")
        strRecvData = Split(RcvBuffer, vbCr)
        
        For intCnt = 0 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            strIntBase = strType
            strResult = ""
            
            Select Case intCnt
                Case 4
                    If InStr(strRcvBuf, "AUTO_SID") > 0 Then
                        strSeq = Mid(strRcvBuf, InStr(strRcvBuf, "AUTO_SID") + 8)
                    '    For i = 1 To vasID.MaxRows
                    '        If Val(GetText(vasID, i, colSPECNO)) = Val(strSeq) Then
                    '            strBarno = GetText(vasID, i, colBARCODE)
                    '            gRow = i
                    '            Exit For
                    '        End If
                    '    Next
                    End If
                    
                    With mResult
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                        .TubePos = strSeq
                    End With
                            
                    Call SetPatInfo(strSeq, gHOSP.RSTTYPE)
                
                Case 9 To 27
                    strIntBase = Trim(Mid(strRcvBuf, 1, 2))
                    strResult = Trim(Mid(strRcvBuf, 3))
                    strResult = Replace(strResult, "h", "")
                    strResult = Replace(strResult, "H", "")
                    strResult = Replace(strResult, "l", "")
                    strResult = Replace(strResult, "L", "")
                    strResult = Replace(strResult, " ", "")
        
                    If strIntBase = "'" Then
                        strIntBase = "|"
                    End If
                    
                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    'If strTestName = gArrEQP(intCol - colSTATE, 5) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        strTestCodeSub = gArrEQP(intCol - colSTATE, 7)
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestCodeSub, strRstRow, colRSUBCD          '검사코드SUB
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        strTestCodeSub = gArrEQP(intCol - colSTATE, 7)
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestCodeSub, strRstRow, colRSUBCD          '검사코드SUB
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                        End If
                    End If
            
                Case 28
            
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_BIGUBCARE(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
                    
                    SetText .spdOrder, "0", gRow, colCHECKBOX
            End Select
        Next
    End With

End Sub



Private Sub Phase_Serial_OSMOPRO()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case ACK
                        If strState = "Q" Then
                            'Call SendOrder_VERSACELL
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                        intPhase = 4
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        Call SerialRcvData_OSMOPRO
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            frmMain.comEqp.Output = ENQ
                            SetRawData "[Tx]" & ENQ
                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
            
End Sub

Private Sub Phase_Serial_COULTERLH780()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case BufChar
            Case STX
                If intPhase = 1 Then
                    RcvBuffer = ""
                End If
                RcvBuffer = RcvBuffer & BufChar
            Case ETX
                RcvBuffer = RcvBuffer & BufChar
                intPhase = intPhase + 1
                If intPhase = 4 Then
                    Call SerialRcvData_COULTERLH780
                    RcvBuffer = ""
                    intPhase = 1
                    Exit For
                End If
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i

'    For i = 1 To lngBufLen
'        BufChar = Mid$(pBuffer, i, 1)
'
'        Select Case intPhase
'            Case 1  '## SYN 대기
'                Select Case BufChar
'                    Case ENQ
'                        If strState = "O" Then
''                            Call SendOrder_COULTERLH780(BufChar)
'                        End If
'                    Case SYN
'                        RcvBuffer = ""
'
'                        If strState = "O" Then
''                            Call SendOrder(BufChar)
'                        Else
'                            'comEqp.Output = SYN
'                            'SetRawData "[Tx]" & SYN
'                            intPhase = 2
'                        End If
'                    Case NAK
'                        If strState = "O" Then
''                            Call SendOrder(BufChar)
'                        End If
'                    Case ACK
'                        If strState = "O" Then
''                            Call SendOrder(BufChar)
'                        End If
''                    Case DLE
'                        If strState = "O" Then
''                            Call SendOrder(BufChar)
'                        End If
'                    Case "A"
'                        If strState = "O" Then
''                            Call SendOrder(BufChar)
'                        End If
'                    Case "B"
'                        If strState = "O" Then
''                            Call SendOrder(BufChar)
'                        End If
'                    Case "C"
'                        If strState = "O" Then
''                            Call SendOrder(BufChar)
'                        End If
'                    Case "D"
'                        If strState = "O" Then
''                            Call SendOrder(BufChar)
'                        End If
'                    Case "E"
'                        If strState = "O" Then
''                            Call SendOrder(BufChar)
'                        End If
'                    Case "F"
'                End Select
'            Case 2  '## Block Count 대기
'                mOrder.BlkCnt = mOrder.BlkCnt + 1
'                If mOrder.BlkCnt = 2 Then
'                    comEqp.Output = ACK
''                    Call mIntLib.WriteLog(ACK, ccPCLog)
'                    mOrder.BlkCnt = 0
'                    intPhase = 3
'                End If
'            Case 3  '## STX 대기
'                Select Case BufChar
'                    Case STX
'                        intPhase = 4
'                End Select
'            Case 4  '## Block Num 대기
'                mOrder.BlkCnt = mOrder.BlkCnt + 1
'                If mOrder.BlkCnt = 2 Then
'                    mOrder.BlkCnt = 0
'                    intPhase = 5
'                End If
'            Case 5  '## ETX 대기
'                Select Case BufChar
'                    Case ETX
'                        'Call ExpectCRC
'                        'comEqp.Output = ACK
'                        'Call mIntLib.WriteLog(ACK, ccPCLog)
'                        intPhase = 6
'                    Case Else
'                        RcvBuffer = RcvBuffer & BufChar
'
'                End Select
'            Case 6  '## STX, SYN 대기
'                Select Case BufChar
'                    Case SYN
'                        Call SerialRcvData_COULTERLH780
'                        'comEqp.Output = ACK
'                        'Call mIntLib.WriteLog(ACK, ccPCLog)
'                        intPhase = 1
'                    Case STX
'                        intPhase = 4
'                End Select
'        End Select
'    Next i
    
'    For i = 1 To lngBufLen
'        BufChar = Mid(pBuffer, i, 1)
'
'        Select Case intPhase
'            Case 1      ''SYN, Blockcount 대기(datablock 이전의 대기상태)
'                Select Case Asc(BufChar)
'                    Case 22     'SYN에 해당
'                        comEqp.Output = Chr(22)     'SYN
'                        RcvBuffer = RcvBuffer & BufChar   'wkBuf
'                        intPhase = 1
'
'                    Case Else   'blockcount-> 2 chars에 해당
'                        comEqp.Output = Chr(6)      'ACK
'                        intPhase = 2
'                End Select
'
'            Case 2      ''datablock 수신 상태(one datablock의 끝인 ETX 이전까지)
'                Select Case Asc(BufChar)
'                    Case 3      'ETX
'                        comEqp.Output = Chr(6)      'ACK
'                        RcvBuffer = RcvBuffer & BufChar
'                        intPhase = 3
'
'                    Case Else
'                        RcvBuffer = RcvBuffer & BufChar
'                        intPhase = 2
'                End Select
'
'            Case 3      ''전송이 끝인지 or 다른 datablock 전송의 시작인지 판단하여 상태 변환
'                Select Case Asc(BufChar)
'                    Case 22     'SYN, 즉 전송의 끝
'                        comEqp.Output = Chr(6)      'ACK
'                        RcvBuffer = RcvBuffer & BufChar
'
'                        Call SerialRcvData_COULTERLH780
'
'                        RcvBuffer = ""
'                        intPhase = 1
'
'                    Case 2  'STX, 즉 다른 datablock 전송 시작
'                        'ix1 = ix1 + 3   'manual dataformat 참조 p.11
'                        ''일단은 다 전송받고 edit_data에서 걸러내는 것으로 바꿈.
'                        RcvBuffer = RcvBuffer & BufChar
'                        intPhase = 2
'
'                End Select
'
'            '--- ORDER 전송 관련
'            Case 4
'                Select Case Asc(BufChar)
'                    Case 5      'ENQ
'                        Call SendOrder_COULTERLH780
'                        intPhase = 5
'
'                    Case 22     'SYN
'                        comEqp.Output = Chr(22)
'                        intPhase = 1
'
'                End Select
'
'            Case 5
'                Select Case Asc(BufChar)
'                    Case 6      'ACK
'                        Call SendOrder_COULTERLH780
'                        intPhase = 6
'
'                    Case Else   'NAK -> RECEIVER ABORT
'                        intPhase = 1
'
'                End Select
'
'            Case 6
'                Select Case Asc(BufChar)
'                    Case 6      'ACK
'                        comEqp.Output = Chr(5)      'ENQ
'                        intPhase = 7
'
'                    Case Else
'                        intPhase = 1
'
'                End Select
'
'            Case 7
'                Select Case Asc(BufChar)
'                    Case 6      'ACK
'
'                    Case 16     'DLE
'                        intPhase = 8
'
'                End Select
'
'            Case 8      'RETURN CODE 대기
'                Select Case Asc(BufChar)
'                    Case 65, 66, 67, 68, 69, 70     'A, B, C, D, E, F
'                        intPhase = 1
'                        intSndPhase = 1
'
'                    Case Else
'                        intPhase = 1
'                        intSndPhase = 1
'
'                End Select
'
'        End Select
'    Next i
            
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
'                                 2. 장비수신정보 화면표시,
'                                 3. 처방코드 가져오기,
'                                 4. (처방코드로)검사오더 만들기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Private Sub GetOrder_Versacell(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If
        
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
            
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_VERSACELL(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If


        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
'                                 2. 장비수신정보 화면표시,
'                                 3. 처방코드 가져오기,
'                                 4. (처방코드로)검사오더 만들기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Private Sub GetOrder_ETIMAX3000(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If
        
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
            
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_ETIMAX3000(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = "^^^UNKNOWN"
            
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If


        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub


'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
'                                 2. 장비수신정보 화면표시,
'                                 3. 처방코드 가져오기,
'                                 4. (처방코드로)검사오더 만들기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Private Sub GetOrder_ACLELITE(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If
        
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
'        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
'        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
            
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_VERSACELL(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If


        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_VERSACELL(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_VERSACELL = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    strExamCode = ""
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            If AdoRs_Local.Fields("SENDCHANNEL").Value & "" <> "990" Then
                strExamCode = strExamCode & "\^^^" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            End If
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_VERSACELL = Mid(strExamCode, 2)
    
End Function

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_ETIMAX3000(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_ETIMAX3000 = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    strExamCode = ""
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            If AdoRs_Local.Fields("SENDCHANNEL").Value & "" <> "990" Then
                strExamCode = strExamCode & "\^^^" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            End If
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_ETIMAX3000 = Mid(strExamCode, 2)
    
End Function

Private Function GetEquipExamCode_ACLELITE(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String
    Dim strTemp     As String
    Dim strIntBase  As String

    GetEquipExamCode_ACLELITE = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    strExamCode = ""
    mOrder.Count = 0
    Erase mOrder.Items
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
                        'If AdoRs_Local.Fields("SENDCHANNEL").Value & "" <> "990" Then
            strIntBase = Mid(Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & ""), 1, 4)
            If strIntBase <> strTemp Then
                strExamCode = strExamCode & "^^^" & strIntBase
                mOrder.Count = mOrder.Count + 1
                ReDim Preserve mOrder.Items(mOrder.Count)
                mOrder.Items(mOrder.Count) = "^^^" & strIntBase
                strTemp = strIntBase
            End If
                        'End If
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_ACLELITE = strExamCode
    
End Function

Private Sub SerialRcvData_VERSACELL()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    'Dim strOldBarno        As String   '수신한 바코드번호
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strAspect       As String
    Dim strEqpNm        As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim strINTRResult   As String
    
    Dim i               As Integer
    Dim j               As Integer
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "P"    '## Patient
                    '2P|1|Multi QC Lv.1|||45731     /||||||||
                    '2P|1|Multi QC Lv3|||45733     /||||||||
                    '2P|1|AMMONIA QC Lv.1|||54181     /||||||||
                    '2P|1|AMMONIA QC Lv.3|||54183     /||||||||
                    
                    If InStr(mGetP(strRcvBuf, 3, "|"), "QC") > 0 Then
                        mResult.Kind = "QC"
                        .spdQcResult.MaxRows = 0
                    Else
                        mResult.Kind = ""
                    End If

                Case "Q"    '## Request Information
                    If mGetP(strRcvBuf, 13, "|") = "A" Then Exit Sub
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = Trim$(mGetP(strTemp1, 2, "^"))
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                        .Seq = mGetP(strTemp1, 3, "^")
                        .RackNo = mGetP(strTemp1, 4, "^")
                        .TubePos = mGetP(strTemp1, 5, "^")
                    End With
                    
                    Call GetOrder_ACLELITE(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "O"
'                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_JWINFO(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                            
                            Call CalProcess(spdOrder, spdResult, strTestCode)
                            
                        End If
                        strState = ""
                        
                    End If
                        
                    mResult.EqpCd = ""
                    
                    strBarno = mGetP(strRcvBuf, 3, "|")
                    
                    If strBarno = "" Then Exit Sub
                    
                    With mResult
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With
                
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                Case "R"
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strResult = mGetP(strRcvBuf, 4, "|")
                    strFlag = UCase(Mid(mGetP(strRcvBuf, 5, "|"), 1, 1))

                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            'SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & " WHERE RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                strResult = getCRRValue(strTestCode, strResult)
                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                
                                '-- BIORAD QC 저장
'                                If Mid(strBarno, 1, 2) = "QC" Then
'                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
'                                End If
                            
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            'SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & " WHERE RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
        
                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                strResult = getCRRValue(strTestCode, strResult)
                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                                                
                                '-- BIORAD QC 저장
                                If mResult.Kind = "QC" Then
                                    strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                    Call SendBioRadQC(strQCData)
                                End If
                                
                                'If strState <> "R" Then
                                    strState = ""
                                'End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                            
                        End If
                        
                    End If
                                
                    .spdResult.RowHeight(-1) = 14
                
                Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_JWINFO(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                            
                            Call CalProcess(spdOrder, spdResult, strTestCode)
                            
                        End If
                        strState = ""
                        
                    End If
                    
'                Case "L"
'                    '-- BIORAD QC 저장
'                    If mResult.Kind = "QC" Then
'                        If .spdQcResult.MaxRows > 0 Then
'                            strQCData = ""
'                            For i = 1 To .spdQcResult.MaxRows
'                                For j = 1 To 16
'                                    strQCData = strQCData & Trim(GetText(.spdQcResult, i, j)) & "|"
'                                Next
'                                strQCData = strQCData & vbCrLf
'                            Next
'                            If strQCData <> "" Then
'                                Call SendBioRadQC(strQCData)
'                            End If
'                        End If
'                    End If
            End Select
        Next
    End With

End Sub

Private Sub SerialRcvData_ABBOTTRUBY()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    'Dim strOldBarno        As String   '수신한 바코드번호
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strAspect       As String
    Dim strEqpNm        As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim strINTRResult   As String
    
    Dim i               As Integer
    Dim j               As Integer
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "P"    '## Patient
'                    If InStr(mGetP(strRcvBuf, 3, "|"), "QC") > 0 Then
'                        mResult.Kind = "QC"
'                        .spdQcResult.MaxRows = 0
'                    Else
'                        mResult.Kind = ""
'                    End If

                Case "O"
                    strBarno = mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^")
                    
                    If strBarno = "" Then Exit Sub
                    '180000006797
                    strBarno = Mid(strBarno, 1, 11)
                    
                    With mResult
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With
                
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                Case "R"
'                    strSeq = mGetP(strRcvBuf, 2, "|")
                    
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 7, "^")
                    strResult = mGetP(strRcvBuf, 4, "|")
                    strFlag = mGetP(strRcvBuf, 7, "|")

                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            'SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & " WHERE RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                'strResult = getCRRValue(strTestCode, strResult)
                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                
                                '-- BIORAD QC 저장
'                                If Mid(strBarno, 1, 2) = "QC" Then
'                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
'                                End If
                            
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            'SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & " WHERE RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
        
                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                'strResult = getCRRValue(strTestCode, strResult)
                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                                                
                                '-- BIORAD QC 저장
'                                If mResult.Kind = "QC" Then
'                                    strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'                                    Call SendBioRadQC(strQCData)
'                                End If
                                
                                'If strState <> "R" Then
                                    strState = ""
                                'End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                            
                        End If
                        
                    End If
                                
                    .spdResult.RowHeight(-1) = 14
                
                Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_PLIS(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                            
                            'Call CalProcess(spdOrder, spdResult, strTestCode)
                            
                        End If
                        strState = ""
                        
                    End If
                    

            End Select
        Next
    End With

End Sub

Private Sub SerialRcvData_ETIMAX3000()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    'Dim strOldBarno        As String   '수신한 바코드번호
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strAspect       As String
    Dim strEqpNm        As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim strINTRResult   As String
    
    Dim i               As Integer
    Dim j               As Integer
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                
                Case "Q"    '## Request Information
                    If mGetP(strRcvBuf, 13, "|") = "A" Then Exit Sub
                    strBarno = Trim(mGetP(strRcvBuf, 3, "|"))
                    mOrder.OrgBarNo = strBarno
                    strBarno = Mid(strBarno, 1, 11)
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                    End With
                    
                    Call GetOrder_ETIMAX3000(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "P"    '## Patient
                    strSeq = mGetP(strRcvBuf, 2, "|")
                    strBarno = mGetP(strRcvBuf, 3, "|")
                    mOrder.Seq = strSeq
                    
                    'mOrder.OrgBarNo = strBarno
                    'strBarno = Mid(strBarno, 1, 11)
                    
                    'If strBarno = "" Then Exit Sub
                    
                    With mResult
                        .Seq = strSeq
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With
                
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                Case "O"    '## Order
'                    strSeq = mGetP(strRcvBuf, 2, "|")
'                    strBarno = mGetP(strRcvBuf, 3, "|")
'                    mOrder.OrgBarNo = strBarno
'                    strBarno = Mid(strBarno, 1, 11)
'
'                    If strBarno = "" Then Exit Sub
'
'                    With mResult
'                        .Seq = strSeq
'                        .BarNo = strBarno
'                        .RsltDate = Format(Now, "yyyymmddhhmmss")
'                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'                    End With
'
'                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'                    strState = "O"
                    
                Case "R"
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strResult = mGetP(strRcvBuf, 4, "|")
                    strFlag = mGetP(strRcvBuf, 5, "|")

                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            'SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & " WHERE RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                'strResult = getCRRValue(strTestCode, strResult)
                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                
                                '-- BIORAD QC 저장
'                                If Mid(strBarno, 1, 2) = "QC" Then
'                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
'                                End If
                            
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            'SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & " WHERE RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
        
                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                'strResult = getCRRValue(strTestCode, strResult)
                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                                                
                                '-- BIORAD QC 저장
'                                If mResult.Kind = "QC" Then
'                                    strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'                                    Call SendBioRadQC(strQCData)
'                                End If
                                
                                'If strState <> "R" Then
                                    strState = ""
                                'End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                            
                        End If
                        
                    End If
                                
                    .spdResult.RowHeight(-1) = 14
                
                Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_PLIS(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                            
                            'Call CalProcess(spdOrder, spdResult, strTestCode)
                            
                        End If
                        strState = ""
                        
                    End If
                    

            End Select
        Next
    End With

End Sub


Private Sub SerialRcvData_ACLELITE()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    'Dim strOldBarno        As String   '수신한 바코드번호
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strAspect       As String
    Dim strEqpNm        As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim strINTRResult   As String
    
    Dim i               As Integer
    Dim j               As Integer
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "P"    '## Patient
                Case "Q"    '## Request Information
                    If mGetP(strRcvBuf, 13, "|") = "A" Then Exit Sub
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = Trim$(mGetP(strTemp1, 2, "^"))
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                        .Seq = mGetP(strTemp1, 3, "^")
                        .RackNo = mGetP(strTemp1, 4, "^")
                        .TubePos = mGetP(strTemp1, 5, "^")
                    End With
                    
                    Call GetOrder_Versacell(strBarno, gHOSP.RSTTYPE)
                    strState = "Q"
                
                Case "O"
                    '3O|1|03498081||^^^FT4  |R||||||||||1|||||||||CENTAURXP|
                    '3O|1|K1924282||^^^aHBs2|R|||||||||||||||||||CENTAURXP|
                    '3O|1|K1924282||^^^aHBs2|R|||||||||||||||||||CENTAURXP|
                    '3O|1|03498303||^^^Na   |R||||||||||1|||||||||ADVIA1800|
                    '3O|1|03498300||^^^Na   |R||||||||||1|||||||||ADVIA1800|

'1H|\^&||||62 Flanders-Bartley Road^Flanders^NJ^07921||973-927-2828|N81|||P|1|20170904051637
'2P|1|Multi QC Lv.1|||45731     /||||||||
'3O|1|PA003||^^^CO2_L|R||20170903||||||||1|||||||||ADVIA1800|
'4R|1|^^^CO2_L|13.6||||N|F||||20170904050713|ADVIA1800
'5L|1|N
'
'1H|\^&||||62 Flanders-Bartley Road^Flanders^NJ^07921||973-927-2828|N81|||P|1|20170904051641
'2P|1|Multi QC Lv3|||45733     /||||||||
'3O|1|PB003||^^^CO2_L|R||20170903||||||||1|||||||||ADVIA1800|
'4R|1|^^^CO2_L|26.5||||N|F||||20170904050722|ADVIA1800
'5L|1|N
'
'1H|\^&||||62 Flanders-Bartley Road^Flanders^NJ^07921||973-927-2828|N81|||P|1|20170904051645
'2P|1|AMMONIA QC Lv.1|||54181     /||||||||
'3O|1|PE002||^^^AMM|R||20170903||||||||1|||||||||ADVIA1800|
'4R|1|^^^AMM|82.7||||N|F||||20170904050734|ADVIA1800
'5L|1|N
'
'1H|\^&||||62 Flanders-Bartley Road^Flanders^NJ^07921||973-927-2828|N81|||P|1|20170904051649
'2P|1|AMMONIA QC Lv.3|||54183     /||||||||
'3O|1|PF002||^^^AMM|R||20170903||||||||1|||||||||ADVIA1800|
'4R|1|^^^AMM|440.4||||N|F||||20170904050737|ADVIA1800
'5L|1|N
                        
                    mResult.EqpCd = ""
                    
                    strBarno = mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^")
                    strRackNo = mGetP(mGetP(strRcvBuf, 3, "|"), 2, "^")
                    strTubePos = mGetP(mGetP(strRcvBuf, 3, "|"), 3, "^")
                    
                    strEqpNm = mGetP(strRcvBuf, 25, "|")
                    If strEqpNm = "" Then
                        strEqpNm = mGetP(strRcvBuf, 26, "|")
                    End If
                    
                    If strEqpNm <> "" Then
                        If UCase(strEqpNm) = "CENTAURXP" Then
                            mResult.EqpCd = gCENXPCD
                        ElseIf UCase(strEqpNm) = "ADVIA1800" Then
                            mResult.EqpCd = gADV18CD
                        End If
                    End If
                    
                    With mResult
                        .BarNo = strBarno
                        .SpcPos = strTubePos & "/" & strRackNo
                        .Seq = strSeq
                        .RackNo = mResult.EqpCd         'strRackNo
                        .TubePos = Mid(strEqpNm, 1, 3)  'strTubePos
                        'If strOldBarno <> strBarno Then
                        '    strOldBarno = strBarno
                            .RsltDate = Format(Now, "yyyymmddhhmmss")
                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    
                            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                            
                        'End If
                    End With
                
                    
                Case "R"
                    '6R|2|^^^aHBs2^^^1^COFF|1.00|mIU/mL||<|N|F||||20170831143313|CENTAURXP
                    
                    '4R|1|^^^CKMB^^^1^DOSE|2.30  |ng/mL|| |N|F||||20170831143543|CENTAURXP
                    '5R|2|^^^CKMB^^^1^COFF|1.00  |ng/mL|| |N|F||||20170831143543|CENTAURXP
                    '6R|3|^^^CKMB^^^1^RLU |14171 |     || |N|F||||20170831143543|CENTAURXP
                    
'                    4R|1|^^^CKMB^^^1^DOSE|0.83|ng/mL|||N|F||||20170902091523|CENTAURXP
'                    5R|2|^^^CKMB^^^1^COFF|1.00|ng/mL|||N|F||||20170902091523|CENTAURXP
'                    6R|3|^^^CKMB^^^1^RLU|8078||||N|F||||20170902091523|CENTAURXP

'                    4R|1|^^^TnIUltra^^^1^DOSE|0.000|ng/mL|||N|F||||20170902091509|CENTAURXP
'                    5R|2|^^^TnIUltra^^^1^COFF|1.000|ng/mL|||N|F||||20170902091509|CENTAURXP
'                    6R|3|^^^TnIUltra^^^1^RLU|1169||||N|F||||20170902091509|CENTAURXP
                    

'                    4R|1|^^^aHCV^^^1^INTR|NR||||N|F||||20170902110124|CENTAURXP
'                    6R|2|^^^aHCV^^^1^INDX|0.08||||N|F||||20170902110124|CENTAURXP
'                    0R|3|^^^aHCV^^^1^COFF|1.00|Index|||N|F||||20170902110124|CENTAURXP
'                    2R|4|^^^aHCV^^^1^RLU|18780||||N|F||||20170902110124|CENTAURXP
                    
                    strTemp1 = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strIntBase = strTemp1
                    strAspect = mGetP(mGetP(strRcvBuf, 3, "|"), 8, "^")
                    strTemp2 = mGetP(strRcvBuf, 4, "|")
                    strFlag = mGetP(strRcvBuf, 7, "|")                  '<
                    strIntResult = mGetP(strRcvBuf, 4, "|")
                    
                    'mResult.EqpNm = mGetP(strRcvBuf, 14, "|")           'CENTAURXP / ADVIA1800
                    If mResult.EqpCd = gCENXPCD Then
                    'If strAspect = "INDX" Or strAspect = "INTR" Or strAspect = "DOSE" Or strAspect = "RLU" Or strAspect = "COFF" Then
                        If strIntBase = "HBsII" Or strIntBase = "EHIV" Then 'INDX
                            strIntBase = strIntBase & "_" & strAspect
                            If strAspect = "INTR" Then  '정성결과
                                strINTRResult = strIntResult
                            End If
                            If strAspect = "INDX" Then
                                If UCase(strINTRResult) = "REACT" Then
                                    strResult = "POSITIVE" & "(" & strIntResult & ")"
                                Else
                                    strResult = "NEGATIVE" & "(" & strIntResult & ")"
                                End If
                            End If
                        
                        ElseIf strIntBase = "aHBs2" Or strIntBase = "aHAVT" Or strIntBase = "aHAVM" Then
                            strIntBase = strIntBase & "_" & strAspect
                            If strAspect = "INTR" Then
                                strINTRResult = strIntResult
                            End If
                            If strAspect = "DOSE" Then
                                If UCase(strINTRResult) = "REACT" Then
                                    strResult = "POSITIVE" & "(" & strIntResult & ")"
                                Else
                                    strResult = "NEGATIVE" & "(" & strIntResult & ")"
                                End If
                            End If
                            
                        ElseIf strIntBase = "aHCV" Then
                            strIntBase = strIntBase & "_" & strAspect
                            If strAspect = "INTR" Then
                                strINTRResult = strIntResult
                            End If
                            If strAspect = "INDX" Then
                                If UCase(strINTRResult) = "REACT" Then
                                    strResult = "Reactive" & "(" & strIntResult & ")"
                                Else
                                    strResult = "Non-reactive" & "(" & strIntResult & ")"
                                End If
                            End If
                        Else
                            If strAspect = "DOSE" Then
                                strIntBase = strIntBase & "_" & strAspect
                                strResult = strIntResult
                            End If
                        End If
                    Else
                        strResult = strIntResult
                    End If
                    
                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            'SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & " WHERE RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                strResult = getCRRValue(strTestCode, strResult)
                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                
                                '-- BIORAD QC 저장
'                                If Mid(strBarno, 1, 2) = "QC" Then
'                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
'                                End If
                            
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            'SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & " WHERE RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
        
                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                strResult = getCRRValue(strTestCode, strResult)
                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                                                
                                '-- BIORAD QC 저장
                                If mResult.Kind = "QC" Then
                                    strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                    Call SendBioRadQC(strQCData)
                                End If
                                
                                'If strState <> "R" Then
                                    strState = ""
                                'End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                            
                        End If
                        
                    End If
                                
                    .spdResult.RowHeight(-1) = 14
                
                

                    
'                Case "C"    '## Comment
'                    '## Abnormal 결과일때 Comment 저장
'                    If strFlag <> "N" Then
'                        strTemp1 = mGetP(strRcvBuf, 4, "|")
'                        strComm = mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")
'                    End If
'
'                Case "L"
'                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_PLIS(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                            
                            Call CalProcess(spdOrder, spdResult, strTestCode)
                            
                        End If
                        strState = ""
                        
                    End If
                    
'                Case "L"
'                    '-- BIORAD QC 저장
'                    If mResult.Kind = "QC" Then
'                        If .spdQcResult.MaxRows > 0 Then
'                            strQCData = ""
'                            For i = 1 To .spdQcResult.MaxRows
'                                For j = 1 To 16
'                                    strQCData = strQCData & Trim(GetText(.spdQcResult, i, j)) & "|"
'                                Next
'                                strQCData = strQCData & vbCrLf
'                            Next
'                            If strQCData <> "" Then
'                                Call SendBioRadQC(strQCData)
'                            End If
'                        End If
'                    End If
            End Select
        Next
    End With

End Sub


Private Sub SerialRcvData_AU480()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    'Dim strOldBarno        As String   '수신한 바코드번호
    Dim strBarno        As String   '수신한 바코드번호
    Dim strOrgBarno        As String   '수신한 바코드번호
    
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strSmpType      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strAspect       As String
    Dim strEqpNm        As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim strINTRResult   As String
    
    Dim i               As Integer
    Dim j               As Integer
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
        '    strRcvBuf = RcvBuffer
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 1, 2)
            
            Select Case strType
                Case "R "    '## Inquiry Order
                    'R 000101 0001                1608270009
                    'R 000502N0001              201803130103

                    'strBarNo = Trim(Mid(strRcvBuf, 14, 20))
                    strBarno = Trim(Mid(strRcvBuf, 14, 26))
                    strOrgBarno = strBarno
                    strBarno = Mid(strBarno, 1, 12)
                    strRackNo = Mid(strRcvBuf, 3, 4)
                    strTubePos = Mid(strRcvBuf, 7, 2)
                    
                    'strSeq = Trim(Mid(strRcvBuf, 9, 5))

                    strSmpType = Mid(strRcvBuf, 9, 1)
                    strSeq = Mid(strRcvBuf, 10, 4)
                    
                    
'                .RACK = Mid(RcvBuffer, 3, 4)
'                .Pos = Mid(RcvBuffer, 7, 2)
'                .SEQNO = Mid(RcvBuffer, 10, 4)
'                .ID = Trim(Mid(RcvBuffer, 14, 20))
                    
                    
                    With mOrder
                        .BarNo = strBarno
                        .Seq = strSeq
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .OrgBarNo = strOrgBarno
                        '.SmpType = strSmpType
                        .SmpType = Space$(1)
                    End With
                    
                    If strBarno = "" Then
                        strBarno = "오더없음_" & Trim(strSeq)
                        'Exit Sub
                    End If
                    
                    Call GetOrder_AU480(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "D "    '## Result
                    'D 000101 0001                1608270009    E001   9.3  002   5.8  
                    strBarno = Trim$(Mid$(strRcvBuf, 14, 26))
                    strRackNo = Mid(strRcvBuf, 3, 4)
                    strTubePos = Mid(strRcvBuf, 7, 2)
                    strSeq = Trim(Mid(strRcvBuf, 9, 5))
                        
                    If strBarno = "" Then
                        strBarno = "오더없음_" & strSeq
                        'Exit Sub
                    End If
                    
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    End With
                    
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                            
                    
                    strTmp = Mid$(strRcvBuf, 45)

                    Do While Len(strTmp) >= 11
                        strIntBase = Mid$(strTmp, 2, 2)
                        strResult = Mid$(strTmp, 4, 6)
                        strResult = Trim(strResult)
                        strComm = Mid$(strTmp, 10, 1)
                        
                        If strIntBase <> "" And strResult <> "" Then
                            If gPatOrdCd <> "" Then
                                SQL = ""
                                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                                SQL = SQL & "  FROM EQPMASTER" & vbCr
                                'SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                SQL = SQL & " WHERE RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                                
                                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                    strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                    strTestName = Trim(RS_L.Fields("TESTNAME"))
                                    strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                    strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
            
                                    '-- 결과Row 추가
                                    strRstRow = .spdResult.DataRowCnt + 1
                                    If .spdResult.MaxRows < strRstRow Then
                                        .spdResult.MaxRows = strRstRow
                                    End If
            
                                    '소수점 처리, 결과 형태 처리
                                    strMachResult = strResult
                                    If strQCTemp = "1" Then
                                        strResult = SetResult(strResult, strIntBase)
                                    End If
                                    'strJudge = SetJudge(strResult, strIntBase)
                                    strJudge = ""
                                    
                                    '진행상태 표시("결과")
                                    SetText .spdOrder, "결과", gRow, colSTATE
            
                                    '결과값 표시
                                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                        If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                            SetText .spdOrder, strResult, gRow, intCol
                                            Exit For
                                        End If
                                    Next
            
                                    '-- 결과 List
                                    SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                    SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                    SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                    SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                    SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                    SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                    SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                    SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                    
                                    '-- 로컬 저장
                                    SetLocalDB gRow, strRstRow, "1", ""
                                    
                                    strState = "R"
                                    
                                    '-- 결과Count
                                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                        SetText .spdOrder, "1", gRow, colRCNT
                                    Else
                                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                    End If
                                    
                                End If
                            Else
                                SQL = ""
                                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                                SQL = SQL & "  FROM EQPMASTER" & vbCr
                                'SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                SQL = SQL & " WHERE RSLTCHANNEL = '" & strIntBase & "' "
                                
                                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                    strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                    strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                    strSeqNo = Trim(RS_L.Fields("SEQNO"))
            
                                    strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                    strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                    strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                    strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                    strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                    strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                    strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                    strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
            
                                    '-- 결과Row 추가
                                    strRstRow = .spdResult.DataRowCnt + 1
                                    If .spdResult.MaxRows < strRstRow Then
                                        .spdResult.MaxRows = strRstRow
                                    End If
            
                                    '소수점 처리, 결과 형태 처리
                                    strMachResult = strResult
                                    If strQCTemp = "1" Then
                                        strResult = SetResult(strResult, strIntBase)
                                    End If
                                    'strJudge = SetJudge(strResult, strIntBase)
                                    strJudge = ""
                                    
                                    '진행상태 표시("결과")
                                    SetText .spdOrder, "결과", gRow, colSTATE
            
                                    '결과값 표시
                                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                        If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                            SetText .spdOrder, strResult, gRow, intCol
                                            Exit For
                                        End If
                                    Next
            
                                    '-- 결과 List
                                    SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                    SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                    SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                    SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                    SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                    SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                    SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                    SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                    
                                    '-- 로컬 저장
                                    SetLocalDB gRow, strRstRow, "1", ""
                                    
                                    'If strState <> "R" Then
                                        strState = ""
                                    'End If
            
                                    '-- 결과Count
                                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                        SetText .spdOrder, "1", gRow, colRCNT
                                    Else
                                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                    End If
                                End If
                            End If
                        End If
                        strTmp = Mid$(strTmp, 12)
                    Loop
                    .spdResult.RowHeight(-1) = 14
                
'                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_MSINFOTEC(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                            
                            'Call CalProcess(spdOrder, spdResult, strTestCode)
                            
                        End If
                        strState = ""
                        
                    End If
            End Select
        Next
    End With

End Sub


Public Sub SerialRcvData_OSMOPRO()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    'Dim strOldBarno        As String   '수신한 바코드번호
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strAspect       As String
    Dim strEqpNm        As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim strINTRResult   As String
    Dim strKind         As String
    
    Dim i               As Integer
    Dim j               As Integer
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "P"    '## Patient
                Case "Q"    '## Request Information
                Case "O"
                    'O|1|3MA005||^^^||20161110082002|20161110082002||0.000|||||20161110082002||||||||20161110082002|||<CR><ETX
                    
                    strBarno = mGetP(strRcvBuf, 3, "|")
                    
                    If UCase(strBarno) = "S1" Or UCase(strBarno) = "S3" Or UCase(strBarno) = "U1" Or UCase(strBarno) = "U2" Then       'Control Result
                        strKind = "QC"
                    Else
                        strKind = ""
                    End If
                    
                    With mResult
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                        .Kind = strKind
                    End With
                
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    
                                        
                Case "R"
                    'R|1|^^^OSMO|51|mOsm/Kg H2O||N|N|F||OperatorID|20161027142723|| 17010095A<CR><ETX>
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strResult = mGetP(strRcvBuf, 4, "|")
                    
                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTEMP")) & ""
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                                                
                                '-- BIORAD QC 저장
                                If mResult.Kind = "QC" Then
                                    
                                    strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                    
                                    Call SendBioRadQC(strQCData)
                                    
                                End If
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                            
                        End If
                        
                    End If
                                
                    .spdResult.RowHeight(-1) = 14

                Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_MCC(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                            
                        End If
                        strState = ""
                        
                    End If
            End Select
        Next
    End With

End Sub


'''Public Sub SerialRcvData_CT500()
'''    Dim RS_L            As ADODB.Recordset
'''    Dim strRcvBuf       As String   '수신한 Data
'''    Dim strType         As String   '수신한 Record Type
'''    Dim strBarno        As String   '수신한 바코드번호
'''    Dim strSeq          As String   '수신한 Sequence
'''    Dim strRackNo       As String   '수신한 Rack Or Disk No
'''    Dim strTubePos      As String   '수신한 Tube Position
'''    Dim strIntBase      As String   '수신한 장비기준 검사명
'''    Dim strMachResult   As String   '수신한 장비결과
'''    Dim strResult       As String   '수신한 결과(정성)
'''    Dim strIntResult    As String   '수신한 결과(정량)
'''    Dim varResult       As Variant
'''    Dim strQCResult     As String   '수신한 결과(QC)
'''    Dim strFlag         As String   '수신한 Abnormal Flag
'''    Dim strComm         As String   '수신한 Comment
'''
'''    Dim strOrderCode     As String   '처방코드
'''    Dim strTestCode      As String   '검사코드
'''    Dim strTestName      As String   '검사명
'''    Dim strSeqNo         As String   '로컬DB 검사Seq
'''
'''    Dim strRstRow        As String   '결과스프레드 현재 Row
'''    Dim intCnt          As Integer  '통신 Frame 갯수
'''    Dim intCol          As Integer  '결과컬럼 갯수
'''    Dim strJudge        As String   '결과판정
'''    Dim Res             As Integer
'''
'''    Dim strTmp          As String
'''    Dim strOldBarno     As String
'''    Dim strQCData       As String
'''    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
'''
'''    With frmMain
'''        strRcvBuf = RcvBuffer
'''        strRcvBuf = Replace(strRcvBuf, vbLf, "")
'''
''''#4-723      17-08-28
''''ID = 3495464
''''Color: STRAW
''''Clarity:
''''GLU NEGATIVE
''''BIL NEGATIVE
''''KET NEGATIVE
''''SG 1.025
''''BLO NEGATIVE
''''pH 6#
''''PRO NEGATIVE
''''URO      0.2 E.U./dL
''''NIT NEGATIVE
''''LEU NEGATIVE
''''
'''
'''
'''        '-- 테스트용 -----------------
'''        If .fraCommTest.Visible = False Then
'''            Call SetSQLData("RCV", strRcvBuf, "A")
'''        End If
'''        '-- 테스트용 -----------------
'''
'''        If Mid(strRcvBuf, 1, 3) = "ID=" Then
'''            miLineNo = 1
'''            mColor = False
'''            'strBarno = Trim(Mid(strRcvBuf, 5, 12))
'''            strBarno = Trim(Mid(strRcvBuf, 4))
'''            mResult.BarNo = strBarno
'''            If strBarno = "1" Or strBarno = "2" Then
'''                mResult.Kind = "QC"
'''            End If
'''
'''            With mResult
'''                .BarNo = strBarno
'''                .SpcPos = strSeq
'''                .Seq = strSeq
'''                .RackNo = strRackNo
'''                .TubePos = strTubePos
'''                If strOldBarno <> strBarno Then
'''                    strOldBarno = strBarno
'''                    .RsltDate = Format(Now, "yyyymmddhhmmss")
'''                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'''                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'''                End If
'''            End With
'''
'''            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'''
'''            If gRow <= 0 Then
'''                Exit Sub
'''            End If
'''        Else
'''            If miLineNo = 1 Then
'''                strIntBase = Trim(mGetP(strRcvBuf, 1, Space$(1)))
'''                If Right(strIntBase, 1) = "*" Then
'''                    strIntBase = Mid(strIntBase, 1, Len(strIntBase) - 1)
'''                End If
'''                strResult = Trim(mGetP(strRcvBuf, 2, Space$(1)))
'''                If strResult = "" Then
'''                    If Len(strIntBase) = 3 Then
'''                        strResult = Trim(Mid(strRcvBuf, 8))
'''                    Else
'''                        strResult = Trim(Mid(strRcvBuf, 9))
'''                    End If
'''                End If
'''                strResult = Replace(strResult, "E.U./dL", "")
'''                strResult = Trim(strResult)
'''
'''
'''                If strIntBase = "Color:" Then
'''                    mColor = True
'''                End If
'''
'''                '--QC
'''                If Len(mResult.BarNo) <= 5 Then
'''                    strResult = Replace(strResult, "<", "")
'''                    strResult = Replace(strResult, ">", "")
'''                    strResult = Replace(strResult, "=", "")
'''                End If
'''
'''RST:
'''                If strIntBase <> "" And strResult <> "" Then
'''                    If gPatOrdCd <> "" Then
'''                        SQL = ""
'''                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH, QCTemp " & vbCr
'''                        SQL = SQL & "  FROM EQPMASTER" & vbCr
'''                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'''                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
'''                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
'''
'''                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'''                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'''                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
'''                            strTestName = Trim(RS_L.Fields("TESTNAME"))
'''                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
'''                            strQCTemp = Trim(RS_L.Fields("QCTEMP"))
'''
'''                            '-- 결과Row 추가
'''                            strRstRow = .spdResult.DataRowCnt + 1
'''                            If .spdResult.MaxRows < strRstRow Then
'''                                .spdResult.MaxRows = strRstRow
'''                            End If
'''
'''                            '소수점 처리, 결과 형태 처리
'''                            strMachResult = strResult
'''                            If strQCTemp = "1" Then
'''                                strResult = SetResult(strResult, strIntBase)
'''                            End If
'''                            strJudge = SetJudge(strResult, strIntBase)
'''
'''                            '진행상태 표시("결과")
'''                            SetText .spdOrder, "결과", gRow, colSTATE
'''
'''                            '결과값 표시
'''                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
'''                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
'''                                    SetText .spdOrder, strResult, gRow, intCol
'''                                    Exit For
'''                                End If
'''                            Next
'''
'''                            '-- 결과 List
'''                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
'''                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
'''                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'''                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
'''                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
'''                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
'''                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
'''                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
'''                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
'''
'''                            '-- 로컬 저장
'''                            SetLocalDB gRow, strRstRow, "1", ""
'''
'''                            strState = "R"
'''
'''                            '-- BIORAD QC 저장
''''                            If mResult.Kind = "QC" Then
''''                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
''''
''''                                Call SendBioRadQC(strQCData)
''''                            End If
'''
'''                            '-- 결과Count
'''                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
'''                                SetText .spdOrder, "1", gRow, colRCNT
'''                            Else
'''                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'''                            End If
'''
'''                        End If
'''                    Else
'''                        SQL = ""
'''                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH, QCTemp,QCAnalyte " & vbCr
'''                        SQL = SQL & "  FROM EQPMASTER" & vbCr
'''                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'''                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
'''
'''                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'''                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'''                            strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'''                            strTestName = Trim(RS_L.Fields("TESTNAME") & "")
'''                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
'''                            strQCTemp = Trim(RS_L.Fields("QCTEMP")) & ""
'''                            strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
'''
'''                            '-- 결과Row 추가
'''                            strRstRow = .spdResult.DataRowCnt + 1
'''                            If .spdResult.MaxRows < strRstRow Then
'''                                .spdResult.MaxRows = strRstRow
'''                            End If
'''
'''                            '소수점 처리, 결과 형태 처리
'''                            strMachResult = strResult
'''                            If strQCTemp = "1" Then
'''                                strResult = SetResult(strResult, strIntBase)
'''                            End If
'''                            strJudge = SetJudge(strResult, strIntBase)
'''
'''                            '진행상태 표시("결과")
'''                            SetText .spdOrder, "결과", gRow, colSTATE
'''
'''                            '결과값 표시
'''                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
'''                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
'''                                    SetText .spdOrder, strResult, gRow, intCol
'''                                    Exit For
'''                                End If
'''                            Next
'''
'''                            '-- 결과 List
'''                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
'''                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
'''                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'''                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
'''                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
'''                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
'''                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
'''                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
'''                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
'''
'''                            '-- 로컬 저장
'''                            SetLocalDB gRow, strRstRow, "1", ""
'''
'''                            If strState <> "R" Then
'''                                strState = ""
'''                            End If
'''
'''                            '-- BIORAD QC 저장
'''                            If mResult.Kind = "QC" Then
'''                                strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, mResult.BarNo, strQCAnalyte, strResult)
'''
'''                                Call SendBioRadQC(strQCData)
'''                            End If
'''
'''                            '-- 결과Count
'''                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
'''                                SetText .spdOrder, "1", gRow, colRCNT
'''                            Else
'''                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'''                            End If
'''                        End If
'''
'''                    End If
'''
'''                End If
'''
'''                .spdResult.RowHeight(-1) = 14
'''
'''                '## DB에 결과저장
'''                If .optTrans(0).Value = True And strState = "R" Then
'''                    Res = SaveTransData_MCC(gRow)
'''
'''                    If Res = -1 Then
'''                        '-- 저장 실패
'''                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'''                        SetText .spdOrder, "Failed", gRow, colSTATE
'''                    Else
'''                        '-- 저장 성공
'''                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'''                        SetText .spdOrder, "저장완료", gRow, colSTATE
'''                        SetText .spdOrder, "0", gRow, colCHECKBOX
'''
'''                              SQL = "Update PATRESULT Set " & vbCrLf
'''                        SQL = SQL & " sendflag = '2' " & vbCrLf
'''                        SQL = SQL & " Where equipno = '" & gHOSP.machCD & "' " & vbCrLf
'''                        SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'''                        SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'''                        SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
'''
'''                        If DBExec(AdoCn_Local, SQL) Then
'''                            '-- 성공
'''                        End If
'''                    End If
'''                    strState = ""
'''                End If
'''
'''                If mColor = False And strIntBase = "LEU" Then
'''                    strIntBase = "Color:"
'''                    strResult = "YELLOW"
'''                    GoTo RST
'''                End If
'''            End If
'''        End If
'''    End With
'''
'''End Sub

Public Sub SerialRcvData_CT500()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim varResult       As Variant
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strOldBarno     As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
'    With frmMain
'        strRcvBuf = RcvBuffer
'        strRcvBuf = Replace(strRcvBuf, vbLf, "")
        
'#4-723      17-08-28
'ID = 3495464
'Color: STRAW
'Clarity:
'GLU NEGATIVE
'BIL NEGATIVE
'KET NEGATIVE
'SG 1.025
'BLO NEGATIVE
'pH 6#
'PRO NEGATIVE
'URO      0.2 E.U./dL
'NIT NEGATIVE
'LEU NEGATIVE
'

    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            If Mid(strRcvBuf, 1, 3) = "ID=" Then
                strBarno = Trim(Mid(strRcvBuf, 4))
                mResult.BarNo = strBarno
                If strBarno = "1" Or strBarno = "2" Then
                    mResult.Kind = "QC"
                End If
                
                With mResult
                    .BarNo = strBarno
                    .SpcPos = strSeq
                    .Seq = strSeq
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                End With
                
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                            
                If gRow <= 0 Then
                    Exit Sub
                End If
            Else
                If intCnt > 2 Then
                    strIntBase = Trim(mGetP(strRcvBuf, 1, Space$(1)))
                    If Right(strIntBase, 1) = "*" Then
                        strIntBase = Mid(strIntBase, 1, Len(strIntBase) - 1)
                    End If
                    strResult = Trim(mGetP(strRcvBuf, 2, Space$(1)))
                    If strResult = "" Then
                        If Len(strIntBase) = 3 Then
                            strResult = Trim(Mid(strRcvBuf, 8))
                        Else
                            strResult = Trim(Mid(strRcvBuf, 9))
                        End If
                    End If
                    strResult = Replace(strResult, "E.U./dL", "")
                    strResult = Trim(strResult)
                    
                    
                    If strIntBase = "Color:" Then
                        mColor = True
                    End If
                        
                    '--QC
                    If Len(mResult.BarNo) <= 5 Then
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "=", "")
                    End If
                    
RST:
                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH, QCTemp " & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTEMP"))
                                
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
            
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
            
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                strState = "R"
                                
                                '-- BIORAD QC 저장
    '                            If mResult.Kind = "QC" Then
    '                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
    '
    '                                Call SendBioRadQC(strQCData)
    '                            End If
                        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH, QCTemp,QCAnalyte " & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTEMP")) & ""
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
            
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
            
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
            
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
            
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
            
                                '-- BIORAD QC 저장
                                If mResult.Kind = "QC" Then
                                    strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, mResult.BarNo, strQCAnalyte, strResult)
                                    
                                    Call SendBioRadQC(strQCData)
                                End If
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                            
                        End If
                        
                    End If
                                
                    .spdResult.RowHeight(-1) = 14
                                
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_MCC(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
                    
                    If mColor = False And strIntBase = "LEU" Then
                        strIntBase = "Color:"
                        strResult = "YELLOW"
                        GoTo RST
                    End If
                End If
            End If
        Next
    End With

End Sub

Private Sub SerialRcvData_XP300()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim varRcvBuf       As Variant
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim varResult       As Variant
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    Dim strOrderCode    As String   '처방코드
    Dim strTestCode     As String   '검사코드
    Dim strTestSubCode  As String   '검사코드
    Dim strTestName     As String   '검사명
    Dim strSeqNo        As String   '로컬DB 검사Seq
    
    Dim strRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strDecYN        As String   '결과판정여부
    Dim strJudge        As String   '결과판정
    
    Dim strTmp          As String
    Dim strOldBarno     As String
    Dim strQCData       As String
    Dim i               As Integer
    Dim Res             As Integer
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                Case "P"    '## Patient
                Case "O"
                    strBarno = Trim(mGetP(mGetP(strRcvBuf, 4, "|"), 3, "^"))
                    strRackNo = Trim(mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^"))
                    strTubePos = Trim(mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^"))

                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With

                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                Case "R"
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                    strResult = mGetP(strRcvBuf, 4, "|")
                    strFlag = mGetP(strRcvBuf, 7, "|")
                    
                    If InStr(strResult, "^") > 0 Then
                        '## 정성결과 저장
                        strResult = mGetP(strResult, 2, "^")
                    Else
                        '## 정량결과 저장
                        strResult = strResult
                    End If
                    
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH  " & vbCr
                        SQL = SQL & ", QCTemp AS DECYN                              " & vbCr
                        SQL = SQL & "  FROM EQPMASTER                               " & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'        " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'      " & vbCr
                        '처방이 있을경우
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ")     " & vbCr
                        End If

                            
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("DECYN") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                                                                                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    '서브코드
                                    strTestSubCode = gArrEQP(intCol - colSTATE, 17)
                                    
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestSubCode, strRstRow, colRSUBCD          '검사SUB코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    End If
                                
                    .spdResult.RowHeight(-1) = 14
                    
                Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_KOMAIN(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                            
                        End If
                        strState = ""
                        
                    End If

            End Select
        Next
    End With

End Sub

Private Sub SerialRcvData_BS240()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim varRcvBuf       As Variant
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim varResult       As Variant
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    Dim strOrderCode    As String   '처방코드
    Dim strTestCode     As String   '검사코드
    Dim strTestSubCode  As String   '검사코드
    Dim strTestName     As String   '검사명
    Dim strSeqNo        As String   '로컬DB 검사Seq
    
    Dim strRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strDecYN        As String   '결과판정여부
    Dim strJudge        As String   '결과판정
    
    Dim strTmp          As String
    Dim strOldBarno     As String
    Dim strQCData       As String
    Dim i               As Integer
    Dim Res             As Integer
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    strRecvData = Split(RcvBuffer, vbCr)
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 1, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 2, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                    'H|\^&||| Product Model ^01.03.07.03^123456|||||||RQ|1394-97|20090910102501<CR>
                    'Q|1|^SAMPLE123||||||||||A<CR>
                    'L|1|N<CR>
                    
                    
                    '1H|\^&|||Mindry^^|||||||RQ|1394-97|20180611140650
                    'L|1|N

                    mOrder.BSMType = mGetP(strRcvBuf, 9, "|")
                    mOrder.BSMaker = mGetP(strRcvBuf, 3, "|")
                    mOrder.BSMchNm = mGetP(strRcvBuf, 4, "|")
                    mOrder.BSDtTm = Format(Now, "yyyymmddhhmmss")
                
                Case "Q"    '## Request Information
                    'Q|1|||||20180611000000|20180611140650|||||O
                    
'                    1H|\^&|||Mindry^^|||||||RQ|1394-97|20180611144830
'
'Q|6||1|1||||||||O
'
'L|6|N
'
'A6
                    strBarno = mGetP(strRcvBuf, 4, "|")
                    strSeq = mGetP(strRcvBuf, 4, "|")
                    
                    With mOrder
                        .Seq = strSeq
                    End With
                    
                    strState = "Q"
                    intSndPhase = 1
                    
                    Call GetOrder_BS240(strBarno, gHOSP.RSTTYPE)
                    
                Case "P"    '## Patient
                
                Case "O"
                    strBarno = Trim(mGetP(mGetP(strRcvBuf, 4, "|"), 3, "^"))
                    strRackNo = Trim(mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^"))
                    strTubePos = Trim(mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^"))

                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With

                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                Case "R"
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                    strResult = mGetP(strRcvBuf, 4, "|")
                    strFlag = mGetP(strRcvBuf, 7, "|")
                    
                    If InStr(strResult, "^") > 0 Then
                        '## 정성결과 저장
                        strResult = mGetP(strResult, 2, "^")
                    Else
                        '## 정량결과 저장
                        strResult = strResult
                    End If
                    
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH  " & vbCr
                        SQL = SQL & ", QCTemp AS DECYN                              " & vbCr
                        SQL = SQL & "  FROM EQPMASTER                               " & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'        " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'      " & vbCr
                        '처방이 있을경우
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ")     " & vbCr
                        End If

                            
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("DECYN") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                                                                                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    End If
                                
                    .spdResult.RowHeight(-1) = 14
                    
                Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_KOMAIN(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                            
                        End If
                        strState = ""
                        
                    End If
            End Select
        Next
    End With

End Sub


Public Sub SerialRcvData_XI921F()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim varRcvBuf       As Variant
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim varResult       As Variant
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    Dim strOrderCode    As String   '처방코드
    Dim strTestCode     As String   '검사코드
    Dim strTestSubCode  As String   '검사코드
    Dim strTestName     As String   '검사명
    Dim strSeqNo        As String   '로컬DB 검사Seq
    
    Dim strRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strDecYN        As String   '결과판정여부
    Dim strJudge        As String   '결과판정
    
    Dim strTmp          As String
    Dim strOldBarno     As String
    Dim strQCData       As String
    Dim i               As Integer
    Dim Res             As Integer
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    
    '001 000000000000000000 063 5.23 151.2 111.4 1.44 7.62
    'Note：Serial number+ Space +ID+ Space + Flag byte + Space +K+ result +Space+ Na+result + Space + Cl- result + Space + Ca2+ result + Space + PH result+ Space.
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            varRcvBuf = Split(strRcvBuf, Space(1))
                        
            If varRcvBuf(0) <> "" Then
                strSeq = Trim(Mid(strRcvBuf, 1, 3)) 'varRcvBuf(0)
                strBarno = Trim(Mid(strRcvBuf, 5, 20)) 'varRcvBuf(2)
                strBarno = Val(strBarno)
                mResult.BarNo = strBarno
                
                With mResult
                    .BarNo = strBarno
                    .Seq = strSeq
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                End With
                
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                            
                If gRow <= 0 Then
                    Exit Sub
                End If
                                    
                '3:k, 4:na, 5:cl, 6:ca2, 7:ph
                For i = 1 To 3 'UBound(varRcvBuf)
                    strIntBase = i
                    If i = 1 Then
                        strResult = Mid(strRcvBuf, 29, 5)
                    ElseIf i = 2 Then
                        strResult = Mid(strRcvBuf, 34, 6)
                    ElseIf i = 3 Then
                        strResult = Mid(strRcvBuf, 40, 6)
                    End If
                    
                    strResult = Trim(strResult)
                    
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH  " & vbCr
                        SQL = SQL & ", QCTemp AS DECYN                              " & vbCr
                        SQL = SQL & "  FROM EQPMASTER                               " & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'        " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'      " & vbCr
                        '처방이 있을경우
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ")     " & vbCr
                        End If
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strDecYN = Trim(RS_L.Fields("DECYN"))
                            
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리
                            strMachResult = strResult
                            If strDecYN = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            
                            '로컬에서 판정
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
        
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    '처방코드
                                    strOrderCode = gArrEQP(intCol - colSTATE, 16)
                                    '서브코드
                                    strTestSubCode = gArrEQP(intCol - colSTATE, 17)
                                    Exit For
                                End If
                            Next
        
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestSubCode, strRstRow, colRSUBCD          '검사SUB코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL            '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT      '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT           'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            strState = "R"
                            
                            '-- BIORAD QC 저장
'                            If mResult.Kind = "QC" Then
'                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                Call SendBioRadQC(strQCData)
'                            End If
                    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    End If
                                
                    .spdResult.RowHeight(-1) = 14
                Next
                
                '## DB에 결과저장
                If .optTrans(0).Value = True And gPatOrdCd <> "" And strState = "R" Then
                    
                    Res = SaveTransData(gRow)
                    
                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "저장실패", gRow, colSTATE
                    Else
                        '-- 저장 성공
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "저장완료", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX
                        
                              SQL = "Update PATRESULT Set " & vbCrLf
                        SQL = SQL & " sendflag = '2' " & vbCrLf
                        SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                        SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                    End If
                    strState = ""
                End If
                    
            End If
        Next
    End With

End Sub




Public Sub SerialRcvData_COULTERLH780()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim varResult       As Variant
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strOldBarno     As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    Dim tmpIFCd$, tmpRst$, tmpFlag$
    Dim sTmp$, sTmp1$, sTmp2$, sTotIFCd$
    Dim sIFCd() As String
    Dim iPos%, iPos2%, ii%

            
    With frmMain
    
        '-- 테스트용 -----------------
        If .fraCommTest.Visible = False Then
            Call SetSQLData("RCV", RcvBuffer, "A")
        End If
        '-- 테스트용 -----------------
        
        'Data를 Edit하기 편리하도록
        '<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>에서
        '[DATA Block]부분만 제외하고 msRcvBuffer 제거한다.
        Do
            iPos = InStr(1, RcvBuffer, Chr(2))
            
            '<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>
            If iPos = 0 Then
                Exit Do
            End If
            
            sTmp1 = Left$(RcvBuffer, iPos - 1)
            sTmp2 = Mid$(RcvBuffer, iPos + 3)
            
            RcvBuffer = ""
            RcvBuffer = sTmp1 & sTmp2
        Loop While iPos <> 0
        
        Do
            iPos = InStr(1, RcvBuffer, Chr(3))
            
            '<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>
            If iPos = 0 Then
                Exit Do
            End If
            
            sTmp1 = Left$(RcvBuffer, iPos - 5)
            sTmp2 = Mid$(RcvBuffer, iPos + 1)
            
            RcvBuffer = ""
            RcvBuffer = sTmp1 & sTmp2
        Loop While iPos <> 0
        
        '작업번호 구하기
        iPos = InStr(RcvBuffer, "ID1")
        If iPos > 0 Then
            sTmp2 = Mid(RcvBuffer, iPos + 4, 16)
            ii = InStr(1, sTmp2, vbCr)
            If ii <> 0 Then
                sTmp2 = Mid(sTmp2, 1, ii - 1)
            End If
            strBarno = sTmp2
        End If
        
        iPos = InStr(RcvBuffer, "CASSPOS")
        If iPos > 0 Then
            sTmp1 = Mid(RcvBuffer, iPos + 9, 6)
                
            strRackNo = Left(sTmp1, 4)
            strTubePos = Right(sTmp1, 2)
        End If
        
        '-- ??
'        mResult.BarNo = strBarno
'        If strBarno = "1" Or strBarno = "2" Then
'            mResult.Kind = "QC"
'        End If
        '-- ??
        strBarno = Mid(strBarno, 1, 11)
        With mResult
            '1           20180118    18000000146 7   02  11          58069   최복순                      6           RDW/MO%/EO%/BA%/LY%/NE%

            .BarNo = strBarno
            '.SpcPos = strSeq
            '.Seq = strSeq
            .RackNo = strRackNo
            .TubePos = strTubePos
            .RsltDate = Format(Now, "yyyymmddhhmmss")
            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
        End With
        
        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
        
        '장비에서 검사할 수 있는 모든 항목 저장
        sTotIFCd = "WBC|RBC|HGB|HCT|MCV|MCH|MCHC|RDW|PLT|PCT|MPV|PDW|" _
                & "LY#|MO#|NE#|EO#|BA#|NRBC#|LY%|MO%|NE%|EO%|BA%|NRBC%|" _
                & "RET%|RET#|MRV|MSCV|IRF|HLR%|HLR#"
                
        sIFCd() = Split(sTotIFCd, Chr(124))
        
        '검사명, 검사결과값 얻기
        For ii = 0 To UBound(sIFCd())
            iPos = InStr(RcvBuffer, Trim(sIFCd(ii)))
            
            If iPos > 0 Then
                sTmp = Trim(Mid(RcvBuffer, iPos + 4, 3))
                If sTmp = "Pop" Then
                    iPos = 0
                ElseIf sTmp = "IS" Then
                    iPos = InStr(iPos + 4, RcvBuffer, Trim(sIFCd(ii)))
                End If
            End If
            
            If iPos > 0 Then
                iPos2 = InStr(iPos, RcvBuffer, Chr(13))
                sTmp = Trim(Mid(RcvBuffer, iPos, iPos2 - iPos))
                
                tmpIFCd = Trim(sIFCd(ii))
                
                sTmp = Trim(Mid(sTmp, Len(tmpIFCd) + 1))
                
                iPos2 = InStr(sTmp, " ")
                If iPos2 > 0 Then
                    tmpRst = Trim(Mid(sTmp, 1, iPos2))
                    tmpFlag = Trim(Mid(sTmp, iPos2))
                Else
                    tmpRst = Trim(sTmp)
                    tmpFlag = ""
                End If
                
    '            tmpRst = Trim(Mid(sTmp, 5, 6))
    '            tmpFlag = Trim(Mid(sTmp, 10))
            
                '-- 결과의 자릿수가 부족해 뒤의 Flag도 표시되는 경우 처리
                iPos = InStr(1, tmpRst, " ")
                If iPos <> 0 Then
                    tmpRst = Trim(Mid(tmpRst, 1, iPos - 1))
                End If
                
                'STKS가 업그레이드 된 후 MCHC결과를 잘라내면 SOH가 뒤에 붙는 현상
                If IsNumeric(Right$(tmpRst, 1)) = True Then
                Else
                    tmpRst = Left$(tmpRst, Len(tmpRst) - 1)
                End If
                
                strIntBase = tmpIFCd
                strResult = tmpRst
                strFlag = tmpFlag
                
                If strIntBase <> "" And strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH, QCTemp " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTEMP"))
                            
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
        
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
        
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            strState = "R"
                            
                            '-- BIORAD QC 저장
'                            If mResult.Kind = "QC" Then
'                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                Call SendBioRadQC(strQCData)
'                            End If
                    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    Else
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH, QCTemp,QCAnalyte " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTEMP")) & ""
                            strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
        
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
        
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
        
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
        
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            If strState <> "R" Then
                                strState = ""
                            End If
        
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, mResult.BarNo, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                        
                    End If
                    
                End If
            End If
        Next ii
    
        .spdResult.RowHeight(-1) = 14
                    
        '## DB에 결과저장
        If .optTrans(0).Value = True And strState = "R" Then
            Res = SaveTransData_PLIS(gRow)
            
            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "Failed", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX
                
                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
        End If
    
    End With

End Sub




Private Sub Phase_Serial_CT500()
     Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

'    For i = 1 To lngBufLen
'        BufChar = Mid$(pBuffer, i, 1)
'        Select Case BufChar
'            Case STX
'                RcvBuffer = ""
'
'                miLineNo = 0
'            Case vbCr
'                Call SerialRcvData_CT500
'
'                miLineNo = 1
'
'                RcvBuffer = ""
'
'            Case ETX
'                RcvBuffer = ""
'                miLineNo = 0
'
'            Case Else
'                RcvBuffer = RcvBuffer & BufChar
'        End Select
'    Next i

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                'If intBufCnt = 0 Then
                    intBufCnt = 1
                    Erase strRecvData
                    ReDim Preserve strRecvData(intBufCnt)
                'Else
                '    intBufCnt = intBufCnt + 1
                '    ReDim Preserve strRecvData(intBufCnt)
                'End If
            Case vbCr
                intBufCnt = intBufCnt + 1
                ReDim Preserve strRecvData(intBufCnt)
            Case vbLf
            
            Case ETX
                Call SerialRcvData_CT500
                Erase strRecvData
                intBufCnt = 0
                
            Case Else
                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
        End Select
    Next i

End Sub

Public Sub Phase_TCP_BS240_ASTM()

End Sub

Public Sub TCPRcvData_BS240_HL7()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarno     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult  As String   '수신한 결과(QC)
    Dim varResult       As Variant
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim intCnt       As Integer
    
    Dim strOrderCode    As String   '처방코드
    Dim strTestCode     As String   '검사코드
    Dim strTestSubCode  As String   '검사코드
    Dim strTestName     As String   '검사명
    Dim strSeqNo        As String   '로컬DB 검사Seq
    
    Dim strTmp      As String
    
    Dim strTGResult As String
    Dim strCHOLResult As String
    Dim strHDLResult As String
    Dim intCol As Integer
    
    Dim blnResult     As Boolean
    
    Dim strRstRow       As String   '결과스프레드 현재 Row
    Dim strDecYN        As String   '결과판정여부
    Dim strJudge        As String   '결과판정
    
    Dim strQCData       As String
    Dim i               As Integer
    Dim Res             As Integer
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    Dim strSndBuffer    As String
    
    blnResult = False
    
    '-- LDL 계산용
    strTGResult = ""
    strCHOLResult = ""
    strHDLResult = ""
    
    With frmMain
        For intCnt = 0 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            'SetRawData "[Rcv]" & strRcvBuf
            
            strType = mGetP(strRcvBuf, 1, "|")
            
            Select Case strType
                Case "MSH"
                    mOrder.BSMType = mGetP(strRcvBuf, 9, "|")
                    mOrder.BSMaker = mGetP(strRcvBuf, 3, "|")
                    mOrder.BSMchNm = mGetP(strRcvBuf, 4, "|")
                    mOrder.BSDtTm = Format(Now, "yyyymmddhhmmss")
                
                    'With mOrder
                    '    .BSMType = strMType
                    '    .BSMaker = strMaker
                    '    .BSMchNm = strMchNm
                    '    .BSDtTm = strDtTm
                    'End With
                    
                    Select Case mOrder.BSMType
                        Case "ORU^R01"  '==> ACK^R01
                                           strSndBuffer = SB & "MSH|^~\&|||" & mOrder.BSMaker & "|" & mOrder.BSMchNm & "|" & mOrder.BSDtTm & "||ACK^R01|" & mOrder.BSMType & "|P|2.3.1||||0||ASCII|||" & vbCr
                            strSndBuffer = strSndBuffer & "MSA|AA|" & mOrder.BSMType & "|Message accepted|||0|" & vbCr
                            strSndBuffer = strSndBuffer & EB & vbCr
                            
                            SetRawData "[Tx]" & strSndBuffer
                            wSck.SendData strSndBuffer
                        
                        '-- 오더 준비
                        Case "QRY^Q02"  '==> QCK^Q02
                                           strSndBuffer = SB & "MSH|^~\&|||" & mOrder.BSMaker & "|" & mOrder.BSMchNm & "|" & mOrder.BSDtTm & "||QCK^Q02|" & mOrder.BSMType & "|P|2.3.1||||0||ASCII|||" & vbCr
                            strSndBuffer = strSndBuffer & "MSA|AA|" & mOrder.BSMType & "|Message accepted|||0|" & vbCr
                            strSndBuffer = strSndBuffer & "ERR|0|" & vbCr '& EB & vbCr
                            strSndBuffer = strSndBuffer & "QAK|SR|OK|" & vbCr
                            strSndBuffer = strSndBuffer & EB & vbCr
                            
                            SetRawData "[Tx]" & strSndBuffer
                            wSck.SendData strSndBuffer
                        
                        '-- 오더 전송
                        Case "ACK^Q03"
                            '-- 최초이후전송
                            Call GetOrder_BS240_HL7(strBarno, gHOSP.RSTTYPE)
                            
                    End Select
                
                Case "QRD"
                    'QRD|20180611153634|R|D|1|||RD||OTH|||T|

                    mOrder.BSQryId = mGetP(strRcvBuf, 5, "|")
                    mOrder.Seq = mOrder.BSQryId
                    
                Case "QRF"
                    'QRF||||1|1|RCT|COR|ALL||

                    mOrder.BSModel = mGetP(strRcvBuf, 2, "|")
                    mOrder.BSSTime = mGetP(strRcvBuf, 3, "|")
                    mOrder.BSETime = mGetP(strRcvBuf, 4, "|")
                    
                    mOrder.BSQRF = strRcvBuf
                    
                    'QRF||||1|7|RCT|COR|ALL||
                    mOrder.Seq = mGetP(strRcvBuf, 5, "|")
                    mOrder.BSBeginNo = mGetP(strRcvBuf, 5, "|")
                    mOrder.BSFinishNo = mGetP(strRcvBuf, 6, "|")
                    
                    '-- 최초오더전송
                    intSndPhase = 1
                    
                    Call GetOrder_BS240_HL7(strBarno, gHOSP.RSTTYPE)
                
                Case "PID"
                    'PID|1|2|||테스트-5|||O|||||||||||||||||||||||
                    'PID|1|||||||O|||||||||||||||||||||||

                    mOrder.BSMType = mGetP(strRcvBuf, 2, "|")
                    strBarno = Trim(mGetP(strRcvBuf, 3, "|"))
                    'strBarno = mGetP(mGetP(strRcvBuf, 6, "|"), 2, "-")
                    
                    If strBarno = "" Then
                        strBarno = Trim(mGetP(strRcvBuf, 2, "|"))
                    End If
                    
                    mResult.BarNo = strBarno
                    
                    If Trim(strBarno) <> Trim(strOldBarno) Then
                        strOldBarno = strBarno
                    
                        With mResult
                            .BarNo = Trim(strBarno)
                            .Seq = strSeq
                            .RsltDate = Format(Now, "yyyymmddhhmmss")
                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                        End With
                    End If
                    
                    Call SetPatInfo_BS240_HL7(strBarno, gHOSP.RSTTYPE)
                    
                Case "OBR"
                    '-- 인터페이스 응답
'                                   strSndBuffer = SB & "MSH|^~\&|||" & mOrder.BSMaker & "|" & mOrder.BSMchNm & "|" & mOrder.BSDtTm & "||ACK^R01|" & mOrder.BSMType & "|P|2.3.1||||0||ASCII|||" & vbCr
'                    strSndBuffer = strSndBuffer & "MSA|AA|" & mOrder.BSMType & "|Message accepted|||0|" & vbCr
'                    strSndBuffer = strSndBuffer & EB & vbCr
'
'                    SetRawData "[Tx]" & strSndBuffer
'                    wSck.SendData strSndBuffer
        
        
                    'OBR|28|6|CHOL|^|Serum|20180529164220|20180529164044|20180529175810|||1|1|Normal|26411|20190131||M|255.000000|18.000000|249.413219|mg/dL|||||||||||||||||||||||||||
        
        
                    'strBarNo = Trim$(mGetP(strRcvBuf, 3, "|"))
                    strSeq = Trim$(mGetP(strRcvBuf, 4, "|"))
                    
                    If strBarno = "" Then
                        strBarno = strSeq
                    End If
                    
'                    blnResult = False
'                    For i = 1 To spdOrder.DataRowCnt
'                        If Trim(GetText(spdOrder, i, colBARCODE)) = Trim(strBarNo) Then
'                            If InStr(Trim(GetText(spdOrder, i, colSTATE)), "오더") > 0 Then
'                                With mResult
'                                    .BarNo = strBarNo
'                                    .Seq = strSeq
'                                    .RsltDate = Format(Now, "yyyymmddhhmmss")
'                                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'                                End With
'                                blnResult = True
'                                Exit For
'                            End If
'                        End If
'                    Next
                    
'                    If blnResult = False Then
'                        With mResult
'                            .BarNo = strBarNo
'                            .Seq = strSeq
'                            .RsltDate = Format(Now, "yyyymmddhhmmss")
'                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'                        End With
'                    End If
                    
'                    Call SetPatInfo_BS240_HL7(strBarNo, gHOSP.RSTTYPE)
                
                Case "OBX"
                    strIntBase = Trim(mGetP(strRcvBuf, 4, "|"))
                    strResult = Trim$(mGetP(strRcvBuf, 6, "|"))
    
    '                '-- TG 결과저장(==>12 로변경)
    '                If Val(strIntBase) = 12 Then
    '                    strTGResult = Trim(strResult)
    '                End If
    '                '-- Chol 결과저장(==>9 로변경)
    '                If Val(strIntBase) = 9 Then
    '                    strCHOLResult = Trim(strResult)
    '                End If
    '                '-- HDL 결과저장(==>14 로변경)
    '                If Val(strIntBase) = 14 Then
    '                    strHDLResult = Trim(strResult)
    '                End If
    
    
RST:
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH  " & vbCr
                        SQL = SQL & ", QCTemp AS DECYN                              " & vbCr
                        SQL = SQL & "  FROM EQPMASTER                               " & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'        " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'      " & vbCr
                        '처방이 있을경우
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ")     " & vbCr
                            strState = "R"
                        Else
                            strState = ""
                        End If
    
                            
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("DECYN") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                                                                                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    '서브코드
                                    strTestSubCode = gArrEQP(intCol - colSTATE, 17)
                                    
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestSubCode, strRstRow, colRSUBCD          '검사SUB코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            'strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        Else
                            strState = ""
                        End If
                    End If
                                
                    .spdResult.RowHeight(-1) = 14
    
                    '-- LDL 계산식
    '                If strTGResult <> "" And strCHOLResult <> "" And strHDLResult <> "" Then
    '                    strIntBase = "99"
    '                    strResult = strCHOLResult - ((strTGResult / 5) + strHDLResult)
    '                    strCHOLResult = ""
    '                    strTGResult = ""
    '                    strHDLResult = ""
    '                    GoTo Rst
    '                End If
                
            End Select
        Next
    
        '## DB에 결과저장
        If .optTrans(0).Value = True And strState = "R" Then
            Res = SaveTransData_KOMAIN(gRow)
            
            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "Failed", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX
                
                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
                
            End If
            strState = ""
            
        End If
    End With
    
End Sub


Private Sub Phase_Serial_XP300()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case ACK
                        If strState = "Q" Then
                            '
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        Call SerialRcvData_XP300
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            frmMain.comEqp.Output = ENQ
                            SetRawData "[Tx]" & ENQ
                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
            
End Sub

Private Sub Phase_Serial_XI921F()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    
    '123456789012345678901234567890123456789012345678901234567890
    '001 000000000000000000 063 5.23 151.2 111.4 1.44 7.62
    'Note：Serial number+ Space +ID+ Space + Flag byte + Space +K+ result +Space+ Na+result + Space + Cl- result + Space + Ca2+ result + Space + PH result+ Space.
    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        If miLineNo = 0 And RcvBuffer = "" Then
            intBufCnt = 1
            Erase strRecvData
            ReDim Preserve strRecvData(intBufCnt)
            RcvBuffer = RcvBuffer & BufChar
        End If
        
        miLineNo = miLineNo + 1
        
        strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
         
        If miLineNo = 63 Then
            Call SerialRcvData_XI921F
            Erase strRecvData
            RcvBuffer = ""
            intBufCnt = 0
            miLineNo = 0
        End If
    Next i
    
    
'    Timer1.Interval = 3000
'    Timer1.Enabled = True
    
End Sub

Private Sub Phase_Serial_RAPIDLAB348()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## STX 대기
                Select Case BufChar
                    Case STX
                        intPhase = 2
                        intBufCnt = 1
                        Erase strRecvData
                        ReDim Preserve strRecvData(intBufCnt)
                
                End Select
            Case 2      '## ETX 대기
                Select Case BufChar
                    Case ETX
                        intPhase = 3
                    Case Else
                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                End Select
            Case 3      '## EOT 대기
                Select Case BufChar
                    Case EOT
                        Call SerialRcvData_RAPIDLAB348
                        intPhase = 1
                End Select
        End Select
    Next i
    
End Sub


Private Sub Phase_Serial_BS240()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case ENQ
                comEqp.Output = ACK
                SetRawData "[Tx]" & ACK
                
            Case EOT
                If strState = "Q" Then
                    comEqp.Output = ENQ
                    SetRawData "[Tx]" & ENQ
                End If
                'comEqp.Output = ACK
                'SetRawData "[Tx]" & ACK
                
            Case ACK
                If strState = "Q" Then
                    Call SendOrder_BS240
                End If
            Case STX
                RcvBuffer = ""
        
            Case ETX
                Call SerialRcvData_BS240
                RcvBuffer = ""
                
                comEqp.Output = ACK
                SetRawData "[Tx]" & ACK
                
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
    
End Sub



Private Sub SerialRcvData_RAPIDLAB348()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strOldBarno        As String   '수신한 바코드번호
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strAspect       As String
    Dim strQCChannel    As String
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strIDRecord     As String   '수신한 Identifyer Record
    Dim strWorkNo       As String   '수신한 WorkNo
    Dim AssayNm         As String
    
    Dim Pos1            As Long
    Dim Pos2            As Long
    Dim x1              As Long
    Dim x2              As Long
    
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    Dim strctHb     As String
    Dim strO2SAT    As String
    Dim strPO2      As String

    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strIDRecord = Trim$(mGetP(strRcvBuf, 1, FS))
            
            If strIDRecord = "SMP_NEW_DATA" Or strIDRecord = "SMP_EDIT_DATA" Then
                '## WorkNo 조회
                Pos1 = InStr(strRcvBuf, "rSEQ")
                If Pos1 > 0 Then
                    Pos2 = InStr(Mid$(strRcvBuf, Pos1), FS)
                    strSeq = Format$(mGetP(Mid$(strRcvBuf, Pos1, Pos2), 2, GS), "#####")
                    strSeq = Val(strSeq)
                Else
                    '## NOTE: WorkNo가 전송되지 않은 에러처리
                    Exit Sub
                End If
                
                '## 바코드번호 조회
                Pos1 = 0: Pos2 = 0
                Pos1 = InStr(strRcvBuf, "iPID")
                If Pos1 > 0 Then
                    Pos2 = InStr(Mid$(strRcvBuf, Pos1), FS)
                    strBarno = Format$(mGetP(Mid$(strRcvBuf, Pos1, Pos2), 2, GS), String$(9, "#"))
                Else
                    '## NOTE: 바코드번호가 전송되지 않은 에러처리
                End If
                
                With mResult
                    .BarNo = strBarno
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .Rerun = ""
                    'If strOldBarno <> strBarno Then
                        'strOldBarno = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
                    'End If
                End With
                          
                x1 = 1
                Do While InStr(x1, strRcvBuf, FS & "m") <> 0
                    x1 = InStr(x1, strRcvBuf, FS & "m")
                    x2 = InStr(x1, strRcvBuf, GS)
        
            '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                    'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                    strIntBase = Mid(strRcvBuf, x1 + 1, x2 - (x1 + 1))
                    x2 = x2 + 1
                    x1 = InStr(x2, strRcvBuf, GS)
                    strResult = Mid(strRcvBuf, x2, x1 - x2)
                    
                    If strIntBase = "mPO2" Then
                        strPO2 = strResult
                    End If

                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("SEQNO") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                strResult = getCRRValue(strTestCode, strResult)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                '-- BIORAD QC 저장
'                                If mResult.Kind = "QC" Then
'                                    strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                    Call SendBioRadQC(strQCData)
'                                End If
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
        
                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                strResult = getCRRValue(strTestCode, strResult)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                '-- BIORAD QC 저장
'                                If mResult.Kind = "QC" Then
'
'                                    strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                    Call SendBioRadQC(strQCData)
'
'                                End If
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                        End If
                    End If
                Loop
                
                x1 = 1
                Do While InStr(x1, strRcvBuf, FS & "c") <> 0
                    x1 = InStr(x1, strRcvBuf, FS & "c")
                    x2 = InStr(x1, strRcvBuf, GS)
            
            '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                    'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                    strIntBase = Mid(strRcvBuf, x1 + 1, x2 - (x1 + 1))
                    x2 = x2 + 1
                    x1 = InStr(x2, strRcvBuf, GS)
                    strResult = Mid(strRcvBuf, x2, x1 - x2)
            
                    If strIntBase = "ctHb(est)" Then
                        strctHb = strResult
                    End If
                    
                    If strIntBase = "cO2SAT" Then
                        strO2SAT = strResult
                    End If
                
                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                strResult = getCRRValue(strTestCode, strResult)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                '-- BIORAD QC 저장
'                                If mResult.Kind = "QC" Then
'                                    strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                    Call SendBioRadQC(strQCData)
'                                End If
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
        
                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                strResult = getCRRValue(strTestCode, strResult)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                '-- BIORAD QC 저장
'                                If mResult.Kind = "QC" Then
'
'                                    strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                    Call SendBioRadQC(strQCData)
'
'                                End If
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                        End If
                    End If
                Loop
            End If
            
            'O2CT = (1.39ctHb x O2SAT/100) + (0.00314pO2)
            strResult = ""
            If strctHb <> "" And strO2SAT <> "" And strPO2 <> "" Then
                strResult = ((1.39 * strctHb) * (strO2SAT / 100)) + (0.00314 * strPO2)
                strResult = Format(strResult, "##.00")
                strResult = Mid(strResult, 1, InStr(strResult, ".") + 1)
                strIntBase = "O2CT"
            End If
            
            If strResult <> "" Then
                SQL = ""
                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                SQL = SQL & "  FROM EQPMASTER" & vbCr
                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                
                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                    strTestCode = Trim(RS_L.Fields("TESTCODE"))
                    strTestName = Trim(RS_L.Fields("TESTNAME"))
                    strSeqNo = Trim(RS_L.Fields("SEQNO"))

                    '-- 결과Row 추가
                    strRstRow = .spdResult.DataRowCnt + 1
                    If .spdResult.MaxRows < strRstRow Then
                        .spdResult.MaxRows = strRstRow
                    End If

                    '소수점 처리, 결과 형태 처리
                    strMachResult = strResult
                    If strQCTemp = "1" Then
                        strResult = SetResult(strResult, strIntBase)
                    End If
                    strJudge = SetJudge(strResult, strIntBase)
                    
                    'CRR 적용
                    strResult = getCRRValue(strTestCode, strResult)
                    
                    '진행상태 표시("결과")
                    SetText .spdOrder, "결과", gRow, colSTATE

                    '결과값 표시
                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                        If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                            SetText .spdOrder, strResult, gRow, intCol
                            Exit For
                        End If
                    Next

                    '-- 결과 List
                    SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                    SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                    SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                    SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                    SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                    SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                    SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                    SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                    
                    '-- 로컬 저장
                    SetLocalDB gRow, strRstRow, "1", ""
                    
                    '-- BIORAD QC 저장
                    If mResult.Kind = "QC" Then
                        strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                        
                        Call SendBioRadQC(strQCData)
                    End If
                    
                    strState = "R"
                    
                    '-- 결과Count
                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                        SetText .spdOrder, "1", gRow, colRCNT
                    Else
                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                    End If
                End If
            End If
            .spdResult.RowHeight(-1) = 14
        
            '#########  QC Define ##########################################################
            If strIDRecord = "QC_NEW_DATA" Or strIDRecord = "QC_EDIT_DATA" Then
                
                .spdQcResult.MaxRows = 0
                '## Type 조회
                Pos1 = InStr(strRcvBuf, "rTYPE")
                If Pos1 > 0 Then
                    Pos2 = InStr(Mid$(strRcvBuf, Pos1), FS)
                    strBarno = mGetP(Mid$(strRcvBuf, Pos1, Pos2), 2, GS)
                    'strBarno = Val(strBarno)
                Else
                    '## NOTE: WorkNo가 전송되지 않은 에러처리
                    Exit Sub
                End If
                
                '## Level 조회
                Pos1 = 0: Pos2 = 0
                Pos1 = InStr(strRcvBuf, "iQLEV")
                If Pos1 > 0 Then
                    Pos2 = InStr(Mid$(strRcvBuf, Pos1), FS)
                    strQCLevel = mGetP(Mid$(strRcvBuf, Pos1, Pos2), 2, GS)
                Else
                    '## NOTE: 바코드번호가 전송되지 않은 에러처리
                End If
                
                
                '## QC 채널 조회
                Pos1 = 0: Pos2 = 0
                Pos1 = InStr(strRcvBuf, "iQFILE")
                If Pos1 > 0 Then
                    Pos2 = InStr(Mid$(strRcvBuf, Pos1), FS)
                    strQCChannel = mGetP(Mid$(strRcvBuf, Pos1, Pos2), 2, GS)
                Else
                    '## NOTE: 바코드번호가 전송되지 않은 에러처리
                End If
                
                With mResult
                    .BarNo = strBarno
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .Rerun = ""
                    .Kind = "QC"
                    'If strOldBarno <> strBarno Then
                        'strOldBarno = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
                    'End If
                    
                    Call SetText(frmMain.spdOrder, strQCChannel, gRow, colPID)
                    Call SetText(frmMain.spdOrder, strQCLevel, gRow, colPNAME)
                End With
                          
                x1 = 1
                Do While InStr(x1, strRcvBuf, FS & "m") <> 0
                    x1 = InStr(x1, strRcvBuf, FS & "m")
                    x2 = InStr(x1, strRcvBuf, GS)
        
            '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                    'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                    strIntBase = Mid(strRcvBuf, x1 + 1, x2 - (x1 + 1))
                    x2 = x2 + 1
                    x1 = InStr(x2, strRcvBuf, GS)
                    strResult = Mid(strRcvBuf, x2, x1 - x2)

                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        'SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCAnalyte = Trim(RS_L.Fields("QCAnalyte"))
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                'strQCData = GetQCResult_Detail(gHOSP.LABCD, strQCChannel, strQCAnalyte, strResult)
                                strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strQCChannel, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    End If
                Loop
                
                
                
                x1 = 1
                Do While InStr(x1, strRcvBuf, FS & "c") <> 0
                    x1 = InStr(x1, strRcvBuf, FS & "c")
                    x2 = InStr(x1, strRcvBuf, GS)
            
            '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                    'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                    strIntBase = Mid(strRcvBuf, x1 + 1, x2 - (x1 + 1))
                    x2 = x2 + 1
                    x1 = InStr(x2, strRcvBuf, GS)
                    strResult = Mid(strRcvBuf, x2, x1 - x2)
            
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strQCChannel, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                        
                    End If
                Loop
                
                Exit Sub
            End If
            '#########  QC Define ##########################################################
        
            '## DB에 결과저장
            If .optTrans(0).Value = True And strState = "R" Then
                Res = SaveTransData_MCC(gRow)
                
                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "Failed", gRow, colSTATE
                Else
                    '-- 저장 성공
                    SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                    SetText .spdOrder, "저장완료", gRow, colSTATE
                    SetText .spdOrder, "0", gRow, colCHECKBOX
                    
                          SQL = "Update PATRESULT Set " & vbCrLf
                    SQL = SQL & " sendflag = '2' " & vbCrLf
                    SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                    SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                    SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                    SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                    End If
                End If
                strState = ""
            End If
        
        Next
        
        
    End With

End Sub

Private Sub Phase_Serial_PFA200()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case vbCr
                Call SerialRcvData_PFA200
                
                RcvBuffer = ""
                
                miLineNo = miLineNo + 1
                
            Case Is <> 10
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i

End Sub


Private Sub SerialRcvData_PFA200()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim varResult       As Variant
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strOldBarno     As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    With frmMain
        strRcvBuf = RcvBuffer
        strRcvBuf = Replace(strRcvBuf, vbLf, "")
'Buffer = ""
'Buffer = Buffer & "PFA-100" & vbCrLf
'Buffer = Buffer & "REV. 2.20   S/N: 3954 " & vbCrLf
'Buffer = Buffer & "05/31/10       01:12 PM" & vbCrLf
'Buffer = Buffer & "ID#: 010000159846" & vbCrLf
'Buffer = Buffer & "Test Type: Collagen/ADP" & vbCrLf
'Buffer = Buffer & "SAMPLE  A:   114 SEC" & vbCrLf
'Buffer = Buffer & "cs: 6781" & vbCrLf

        '-- 테스트용 -----------------
        If .fraCommTest.Visible = False Then
            Call SetSQLData("RCV", strRcvBuf, "A")
        End If
        '-- 테스트용 -----------------
        
        If UCase(Mid(strRcvBuf, 1, 3)) = "PFA" Then
            miLineNo = 1
            
        ElseIf Mid(strRcvBuf, 1, 3) = "ID#" Then
            strBarno = Trim(Mid(strRcvBuf, 5, 12))
            mResult.BarNo = strBarno
    
        ElseIf Mid(strRcvBuf, 1, 10) = "Test Type:" Then
            strIntBase = Trim(Mid(strRcvBuf, 11))
            mResult.IntBase = Trim(strIntBase)
            
        ElseIf Mid(strRcvBuf, 1, 3) = "QC:" Then
            mResult.Kind = "QC"
    
        ElseIf (Mid(strRcvBuf, 1, 9) = "SAMPLE A:") Or (Mid(strRcvBuf, 1, 9) = "SAMPLE B:") Then
            strResult = Mid(strRcvBuf, 10)
            If InStr(UCase(strRcvBuf), "SEC") = 0 Then
                strResult = Trim(strResult)
            Else
                varResult = Split(strResult, "Sec")
                strResult = Trim(varResult(0))
                strFlag = ""
                
                If Left(strResult, 1) = ">" And IsNumeric(Right(strResult, 1)) <> True Then
                    '비정상 결과 & Flag
                    strResult = Mid(strResult, 1, Len(strResult) - 1)
                End If
            End If
            
            mResult.RESULT = strResult
            
        Else
            If miLineNo = 7 And mResult.Kind <> "QC" Then
                If Trim(strRcvBuf) <> "" Then
                    strFlag = Trim(strRcvBuf)
                End If
            End If
            
            If UCase(Mid(strRcvBuf, 1, 3)) = "CS:" Or (miLineNo >= 7 And mResult.Kind <> "QC") Or (miLineNo >= 8 And mResult.Kind = "QC") Then
                strBarno = mResult.BarNo
                strIntBase = mResult.IntBase
                strResult = mResult.RESULT
                
                With mResult
                    .BarNo = strBarno
                    .SpcPos = strSeq
                    .Seq = strSeq
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    If strOldBarno <> strBarno Then
                        strOldBarno = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    End If
                End With
                
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                            
                If gRow <= 0 Then
                    Exit Sub
                End If
                
                            
                If strIntBase <> "" And strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTEMP"))
        
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
        
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
        
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
        
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            strState = "R"
                            
                            '-- BIORAD QC 저장
'                            If mResult.Kind = "QC" Then
'                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                Call SendBioRadQC(strQCData)
'                            End If
                    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    Else
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTEMP"))
        
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
        
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
        
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
        
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            If strState <> "R" Then
                                strState = ""
                            End If
        
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                        
                    End If
                    
                End If
                            
                .spdResult.RowHeight(-1) = 14
                            
                '## DB에 결과저장
                If .optTrans(0).Value = True And strState = "R" Then
                    Res = SaveTransData_MCC(gRow)
                    
                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "Failed", gRow, colSTATE
                    Else
                        '-- 저장 성공
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "저장완료", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX
                        
                              SQL = "Update PATRESULT Set " & vbCrLf
                        SQL = SQL & " sendflag = '2' " & vbCrLf
                        SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                        SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                    End If
                    strState = ""
                End If
            End If
        End If
    End With

End Sub

Private Sub Phase_Serial_AFIAS6()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case "$" 'SOH
                Erase strRecvData
                
                intBufCnt = 1
                ReDim Preserve strRecvData(intBufCnt)
            Case vbCr
                Call SerialRcvData_AFIAS6
            Case Else
                If intBufCnt > 0 Then
                    strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                End If
        End Select
    Next i

End Sub


Private Sub SerialRcvData_AFIAS6()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strRcvBuf = strRecvData(intCnt)
            strBarno = Trim(mGetP(strRcvBuf, 5, "|"))
            strRackNo = ""
            strTubePos = ""
            strSeq = ""
                        
            With mResult
                .BarNo = strBarno
                .SpcPos = strSeq
                .Seq = strSeq
                .RackNo = strRackNo
                .TubePos = strTubePos
                .RsltDate = Format(Now, "yyyymmddhhmmss")
                .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
            End With
            
            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
            If gRow <= 0 Then
                Exit Sub
            End If
                        
            strIntBase = mGetP(strRcvBuf, 8, "|")
            strResult = mGetP(strRcvBuf, 11, "|")
                        
            If strIntBase <> "" And strResult <> "" Then
                If gPatOrdCd <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                    SQL = SQL & "  FROM EQPMASTER" & vbCr
                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                    SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                    
                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                        strTestCode = Trim(RS_L.Fields("TESTCODE"))
                        strTestName = Trim(RS_L.Fields("TESTNAME"))
                        strSeqNo = Trim(RS_L.Fields("SEQNO"))

                        '-- 결과Row 추가
                        strRstRow = .spdResult.DataRowCnt + 1
                        If .spdResult.MaxRows < strRstRow Then
                            .spdResult.MaxRows = strRstRow
                        End If

                        '소수점 처리, 결과 형태 처리
                        strMachResult = strResult
                        strResult = SetResult(strResult, strIntBase)
                        strJudge = SetJudge(strResult, strIntBase)
                        
                        '진행상태 표시("결과")
                        SetText .spdOrder, "결과", gRow, colSTATE

                        '결과값 표시
                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
                            If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                SetText .spdOrder, strResult, gRow, intCol
                                Exit For
                            End If
                        Next

                        '-- 결과 List
                        SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                        SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                        SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                        SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                        SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                        SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                        SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                        SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                        
                        '-- 로컬 저장
                        SetLocalDB gRow, strRstRow, "1", ""
                        
                        strState = "R"
                        
                        '-- 결과Count
                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
                            SetText .spdOrder, "1", gRow, colRCNT
                        Else
                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                        End If
                        
                    End If
                Else
                    SQL = ""
                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                    SQL = SQL & "  FROM EQPMASTER" & vbCr
                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                    
                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                        strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                        strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                        strSeqNo = Trim(RS_L.Fields("SEQNO"))

                        '-- 결과Row 추가
                        strRstRow = .spdResult.DataRowCnt + 1
                        If .spdResult.MaxRows < strRstRow Then
                            .spdResult.MaxRows = strRstRow
                        End If

                        '소수점 처리, 결과 형태 처리
                        strMachResult = strResult
                        strResult = SetResult(strResult, strIntBase)
                        strJudge = SetJudge(strResult, strIntBase)
                        
                        '진행상태 표시("결과")
                        SetText .spdOrder, "결과", gRow, colSTATE

                        '결과값 표시
                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
                            If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                SetText .spdOrder, strResult, gRow, intCol
                                Exit For
                            End If
                        Next

                        '-- 결과 List
                        SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                        SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                        SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                        SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                        SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                        SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                        SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                        SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                        
                        '-- 로컬 저장
                        SetLocalDB gRow, strRstRow, "1", ""
                        
                        If strState <> "R" Then
                            strState = ""
                        End If

                        '-- 결과Count
                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
                            SetText .spdOrder, "1", gRow, colRCNT
                        Else
                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                        End If
                    End If
                    
                End If
                
            End If
                        
            .spdResult.RowHeight(-1) = 14
                        
            '## DB에 결과저장
            If .optTrans(0).Value = True And strState = "R" Then
                Res = SaveTransData_MCC(gRow)
                
                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "Failed", gRow, colSTATE
                Else
                    '-- 저장 성공
                    SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                    SetText .spdOrder, "저장완료", gRow, colSTATE
                    SetText .spdOrder, "0", gRow, colCHECKBOX
                    
                          SQL = "Update PATRESULT Set " & vbCrLf
                    SQL = SQL & " sendflag = '2' " & vbCrLf
                    SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                    SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                    SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                    SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                    End If
                End If
                strState = ""
            End If
        Next
    End With

End Sub


Private Sub Phase_Serial_ADVIA1800()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long


On Error GoTo RST

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        'Erase strRecvData
                        RcvBuffer = ""
                        sRcvState = "": sSndState = ""
                        intPhase = 2
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case Else
                        intPhase = 1
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case STX
'                        intBufCnt = 1
'                        Erase strRecvData
'                        ReDim Preserve strRecvData(intBufCnt)
                        RcvBuffer = ""
                    Case EOT
                        Select Case sRcvState
                            Case "Q"
                                intPhase = 3
                                iTotQueryFlag = iPendingFlag
                                iPendingFlag = 0
                                
                                'Order전송 Start
                                frmMain.comEqp.Output = ENQ
                                sSndState = "E"
                                
                            Case "R"
                                intPhase = 1
                        End Select
                        
                        sRcvState = ""
                    
                    Case ENQ
                        'Erase strRecvData
                        RcvBuffer = ""
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case vbLf
                        intPhase = 2
                        If RcvBuffer <> "" Then
                            Call SerialRcvData_ADVIA1800
                            RcvBuffer = ""
                        End If
                        
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case vbCr
                    Case ETB
                    
                    Case Else
                        intPhase = 2
'                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        RcvBuffer = RcvBuffer & BufChar
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case ACK
                        Select Case sSndState
                            Case "E"        '<ENQ> 전송 후의 상태
                                Call SendOrder_ADVIA1800
                        
                            Case "P"        '<Packet> 전송 후의 상태
                                Call SendOrder_ADVIA1800
                                                
                            Case "L"        '마지막 <Packet> 전송 후의 상태
                                Call SendOrder_ADVIA1800
                                
                                'Order관련 초기화
                                sSndState = ""
                                Erase sSndPacket
                                intPhase = 1
                        End Select
                    
                    Case ENQ
                        'Erase strRecvData
                        RcvBuffer = ""
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case NAK
                        Select Case sSndState
                            Case "E"
                                frmMain.comEqp.Output = Chr(5)
                                intPhase = 3
                            Case "P"
                                frmMain.comEqp.Output = sSndPacket(iOrderFlag)
                                intPhase = 3
                            Case "L"
                                frmMain.comEqp.Output = sSndPacket(iOrderFlag)
                                intPhase = 3
                        End Select
                        
                    Case 4      'EOT
                        'Erase strRecvData
                        RcvBuffer = ""
                        intPhase = 1
                        sRcvState = "": sSndState = ""
                        'Order관련 초기화
                        iPendingFlag = 0: iTotQueryFlag = 0
                        
                End Select
        End Select
    Next i
                
    Exit Sub
    
RST:
     
                strErrMsg = "위    치 : " & gHOSP.MACHNM & "GetTest" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    strErrMsg = strErrMsg & "ORDER    : " & mOrder.BarNo & vbNewLine
    strErrMsg = strErrMsg & "RESLLT   : " & mResult.BarNo & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
            
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
'                                 2. 장비수신정보 화면표시,
'                                 3. 처방코드 가져오기,
'                                 4. (처방코드로)검사오더 만들기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Private Sub GetOrder_ADVIA1800(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If
        
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
            
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_ADVIA1800(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If


        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub


'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_ADVIA1800(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_ADVIA1800 = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    strExamCode = ""
    mOrder.SendCnt = 0
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            ' " 89M 81M 82M 90M 91M108M 85M"
            If Trim(AdoRs_Local.Fields("SENDCHANNEL").Value) <> "" Then
                strExamCode = strExamCode & Right(Space(3) & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & ""), 3) & "M"
                mOrder.SendCnt = mOrder.SendCnt + 1
            End If
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_ADVIA1800 = strExamCode
    
End Function



Private Sub SerialRcvData_ADVIA1800()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strOldBarno        As String   '수신한 바코드번호
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strAspect       As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    Dim strTmp          As String
    Dim i               As Integer
    Dim iBCpos          As Integer
    
    Dim iTBlockNo   As Integer
    Dim iCBlockNo   As Integer
    Dim iItemNo     As Integer
    Dim strKind     As String
    Dim iPos        As Integer
    
    Dim varIntBase()    As String
    Dim varResult()     As String
    Dim varFlag()       As String
    
    Dim strUseRes       As String
    Dim strCREA         As String
    Dim strTP           As String

On Error GoTo RST
    
    
    iBCpos = 2
    
    With frmMain
        'For intCnt = 1 To UBound(strRecvData)
        '    strRcvBuf = strRecvData(intCnt)
            
            strRcvBuf = RcvBuffer
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, iBCpos, 1)
            
            Select Case strType
                Case "q"    '## Request Information(Batch)
                    sRcvState = "Q"
                    sSndState = ""
                    
                Case "Q"    '## Request Information
                    sRcvState = "Q"
                    sSndState = ""
                
                    iTmpPendingFlag = Val(Mid$(strRcvBuf, iBCpos + 6, 2))
                    iPendingFlag = iPendingFlag + iTmpPendingFlag
                    
                    For i = 1 To iPendingFlag
                        strBarno = Trim$(Mid$(strRcvBuf, iBCpos + 9 + 13 * (i - 1), 13))
                        
                        With mOrder
                            .NoOrder = False
                            .BarNo = strBarno
                        End With
                        
                        Call GetOrder_ADVIA1800(strBarno, gHOSP.RSTTYPE)
                        Call SendOrder_ADVIA1800
                    Next
                
                Case "R"
                    sRcvState = "R"
                    
                    iTBlockNo = Val(Mid$(strRcvBuf, iBCpos + 2, 2))
                    iCBlockNo = Val(Mid$(strRcvBuf, iBCpos + 4, 2))
                    iItemNo = Val(Mid$(strRcvBuf, iBCpos + 6, 3))
                    
                    iBCpos = iBCpos + 6
                    
                    strKind = Mid$(strRcvBuf, iBCpos + 17, 1)       'N:Sample, C:Control
                    strBarno = Trim$(Mid$(strRcvBuf, iBCpos + 19, 13))
                                    
                    strTemp2 = Trim$(Mid$(strRcvBuf, iBCpos + 32, 7))
                    iPos = InStr(strTemp2, "-")
                             
                    If iPos = 0 Then
                        strRackNo = ""
                        strTubePos = ""
                    Else
                        strRackNo = Mid$(strTemp2, 1, iPos - 1)
                        strTubePos = Mid$(strTemp2, iPos + 1)
                    End If
                    
                    If strKind = "C" Then       'Control Result
                        strKind = "QC"
                    Else
                        strKind = ""
                    End If
                    
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Kind = strKind
                        .Rerun = ""
                        If strOldBarno <> strBarno Then
                            strOldBarno = strBarno
                            .RsltDate = Format(Now, "yyyymmddhhmmss")
                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    
                            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                            
                        End If
                    End With
                    
                    ReDim Preserve varIntBase(iItemNo)
                    ReDim Preserve varResult(iItemNo)
                    ReDim Preserve varFlag(iItemNo)
                    
                    If iCBlockNo = 1 Then
                        For i = 1 To iItemNo
                            varIntBase(i) = Trim$(Mid(strRcvBuf, iBCpos + 89 + 19 * (i - 1), 3))
                            varResult(i) = Trim(Mid(strRcvBuf, iBCpos + 89 + 4 + 19 * (i - 1), 8))
                            varFlag(i) = Trim(Mid(strRcvBuf, iBCpos + 89 + 8 + 4 + 19 * (i - 1), 3))
                            
                            If InStr(varFlag(i), "R") > 0 Then
                                mResult.Rerun = "R"
                                varFlag(i) = Replace(varFlag(i), "R", "")
                            End If
                        Next i
                    Else
                        For i = 1 To iItemNo
                            varIntBase(i) = Trim$(Mid(strRcvBuf, iBCpos + 39 + 19 * (i - 1), 3))
                            varResult(i) = Trim(Mid(strRcvBuf, iBCpos + 39 + 4 + 19 * (i - 1), 8))
                            varFlag(i) = Trim(Mid(strRcvBuf, iBCpos + 39 + 8 + 4 + 19 * (i - 1), 3))
                            
                            If InStr(varFlag(i), "R") > 0 Then
                                mResult.Rerun = "R"
                                varFlag(i) = Replace(varFlag(i), "R", "")
                            End If
                        Next i
                    End If
                    
                    If mResult.Rerun = "R" Then       'Rerun Result
                        mResult.Kind = mResult.Kind & "R"
                    End If
                    
                    strCREA = ""
                    strTP = ""
                    
                    For i = 1 To iItemNo
                        strIntBase = varIntBase(i)
                        strResult = varResult(i)
                        
'91  C3750N3 CREA(단회)
'110 C2200-1 TP(단회뇨)
                        If strIntBase = "91" Then
                            strCREA = strResult
                        End If
                        If strIntBase = "110" Then
                            strTP = strResult
                        End If

RST1:
                        If strIntBase <> "" And strResult <> "" And strResult <> "ERROR" Then
                            If gPatOrdCd <> "" Then
                                SQL = ""
                                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                                SQL = SQL & "  FROM EQPMASTER" & vbCr
                                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                                
                                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                    strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                    strTestName = Trim(RS_L.Fields("TESTNAME"))
                                    strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                    strQCTemp = Trim(RS_L.Fields("QCTEMP")) & ""
                                    strQCAnalyte = Trim(RS_L.Fields("QCAnalyte")) & ""
                                    
                                    
                                    'LDH  체액
                                    'If strTestCode = "C2590N1" Or strTestCode = "C2590N2" Then '혈액 Or strTestCode = "B2590"
                                    If strTestCode = "B2590N1" Or strTestCode = "B2590N2" Then '혈액 Or strTestCode = "B2590"
                                        If IsNumeric(strResult) Then
                                            strResult = strResult / 6
                                        End If
                                    End If
                                    
                                    '-- 결과Row 추가
                                    strRstRow = .spdResult.DataRowCnt + 1
                                    If .spdResult.MaxRows < strRstRow Then
                                        .spdResult.MaxRows = strRstRow
                                    End If
            
                                    '소수점 처리, 결과 형태 처리
                                    strMachResult = strResult
                                    If strQCTemp = "1" Then
                                        strResult = SetResult(strResult, strIntBase)
                                    End If
                                    strJudge = SetJudge(strResult, strIntBase)
                                    
                                    'CRR 적용
                                    If strKind <> "QC" Then
                                        strResult = getCRRValue(strTestCode, strResult)
                                    End If
                                    
                                    
                                    '진행상태 표시("결과")
                                    SetText .spdOrder, "결과", gRow, colSTATE
            
                                    '결과값 표시
                                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                        If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                            SetText .spdOrder, strResult, gRow, intCol
                                            Exit For
                                        End If
                                    Next
            
                                    '-- 결과 List
                                    SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                    SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                    SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                    SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                    SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                    SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                    SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                    SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                    
                                    '-- High, Low 색 표시
                                    If strJudge <> "" Then
                                        SetForeColor .spdResult, strRstRow, strRstRow, colRMACHRESULT, colRLISRESULT, 255, 0, 0
                                    End If
                                                                        
                                    '-- 로컬 저장
                                    SetLocalDB gRow, strRstRow, "1", ""
                                    
                                    '-- BIORAD QC 저장
                                    If mResult.Kind = "QC" Then
                                        strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                        
                                        Call SendBioRadQC(strQCData)
                                    End If
                                    
                                    strState = "R"
                                    
                                    '-- 결과Count
                                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                        SetText .spdOrder, "1", gRow, colRCNT
                                    Else
                                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                    End If
                                    
                                    If strTestCode = "C3730N1" Or strTestCode = "C3750" Or strTestCode = "C7230" Or strTestCode = "C3750N3" Or strTestCode = "C2302N6" Then
                                        Call CalProcess(spdOrder, spdResult, strTestCode)
                                    End If

                                End If
                            Else
                                SQL = ""
                                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                                SQL = SQL & "  FROM EQPMASTER" & vbCr
                                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                                
                                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                    strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                    strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                    strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                    strQCTemp = Trim(RS_L.Fields("QCTEMP")) & ""
                                    strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                    
                                    'LDH  체액
                                    'If strTestCode = "C2590N1" Or strTestCode = "C2590N2" Then '혈액LDH Or strTestCode = "B2590"
                                    If strTestCode = "B2590N1" Or strTestCode = "B2590N2" Then '혈액LDH Or strTestCode = "B2590"
                                        If IsNumeric(strResult) Then
                                            strResult = strResult / 6
                                        End If
                                    End If
                                    
                                    '-- 결과Row 추가
                                    strRstRow = .spdResult.DataRowCnt + 1
                                    If .spdResult.MaxRows < strRstRow Then
                                        .spdResult.MaxRows = strRstRow
                                    End If
            
                                    '소수점 처리, 결과 형태 처리
                                    strMachResult = strResult
                                    If strQCTemp = "1" Then
                                        strResult = SetResult(strResult, strIntBase)
                                    End If
                                    strJudge = SetJudge(strResult, strIntBase)
                                    
                                    'CRR 적용
                                    If strKind <> "QC" Then
                                        strResult = getCRRValue(strTestCode, strResult)
                                    End If
                                    
                                    '진행상태 표시("결과")
                                    SetText .spdOrder, "결과", gRow, colSTATE
            
                                    '결과값 표시
                                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                        If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                            SetText .spdOrder, strResult, gRow, intCol
                                            Exit For
                                        End If
                                    Next
            
                                    '-- 결과 List
                                    SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                    SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                    SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                    SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                    SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                    SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                    SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                    SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                                    
                                    '-- High, Low 색 표시
                                    If strJudge <> "" Then
                                        SetForeColor .spdResult, strRstRow, strRstRow, colRMACHRESULT, colRLISRESULT, 255, 0, 0
                                    End If
                                    
                                    '-- 로컬 저장
                                    SetLocalDB gRow, strRstRow, "1", ""
                                    
                                    '-- BIORAD QC 저장
                                    If mResult.Kind = "QC" Then
                                        
                                        strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                        
                                        Call SendBioRadQC(strQCData)
                                        
                                    End If
                                    
                                    If strState <> "R" Then
                                        strState = ""
                                    End If
            
                                    '-- 결과Count
                                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                        SetText .spdOrder, "1", gRow, colRCNT
                                    Else
                                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                    End If
                                End If
                                
                            End If
                            
                        End If
                    Next
                    
                    If strTP <> "" And strCREA <> "" And IsNumeric(strTP) And IsNumeric(strCREA) Then
                        strIntBase = "PRCR"
                        strResult = strTP / strCREA
                        strResult = Format(strResult, "#,##0.00")
                        strTP = ""
                        strCREA = ""
                        GoTo RST1
                    End If
                    
                    .spdResult.RowHeight(-1) = 14
                
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_MCC(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                            
                            If strTestCode = "C3730N1" Or strTestCode = "C3750" Or strTestCode = "C7230" Or strTestCode = "C3750N3" Or strTestCode = "C2302N6" Then
                                Call CalProcess(spdOrder, spdResult, strTestCode)
                            End If
                            
                        End If
                        strState = ""
                    End If
            End Select
        'Next
    End With

    Exit Sub
    
RST:
     
                strErrMsg = "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_ADVIA1800" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    strErrMsg = strErrMsg & "ORDER    : " & mOrder.BarNo & vbNewLine
    strErrMsg = strErrMsg & "RESLLT   : " & mResult.BarNo & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
End Sub


Private Sub SerialRcvData_RAPIDPOINT500()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strOldBarno        As String   '수신한 바코드번호
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strAspect       As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq
    
    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String

    Dim X   As Integer
    Dim C   As Integer
    Dim MsgID   As String
    
    Dim R   As Integer
    Dim x1  As Integer
    Dim x2  As Integer
    Dim AssayNm As String
    Dim RESULT  As String
    Dim EqCd    As String
    Dim OrdCd   As String
    Dim LabNo   As String
    Dim rSeq    As String
    Dim iPID    As String
    Dim iQID    As String

    Dim sRstDate$, sRstTime$
    Dim MsgBuf$
    
'    Dim strQCResult As String
    
    Dim iQLEV$, iQLOT$, strAnalyte$
    Dim db_tmp As String * 100
    
    Dim Pos1            As Long
    Dim Pos2            As Long
    Dim strQCChannel    As String
    
    With frmMain
        '-- 테스트용 -----------------
        If .fraCommTest.Visible = False Then
            Call SetSQLData("RCV", RcvBuffer, "A")
        End If
        '-- 테스트용 -----------------
        
        X = InStr(1, RcvBuffer, FS)
        If RcvBuffer <> "" Then
            MsgID = Mid(RcvBuffer, 2, X - 2)
        End If
        Select Case MsgID
            Case "ID_REQ"
                Call SendMessage_1200("ID_DATA")
            Case "SMP_START"
            Case "SMP_NEW_AV"
                Do Until X = 0
                    X = InStr(X, RcvBuffer, "r")
                    If X = 0 Then Exit Do
                    If Mid(RcvBuffer, X, 4) = "rSEQ" Then
                        X = X + 5
                        C = InStr(X, RcvBuffer, GS)
                        Sample_Seq = Mid(RcvBuffer, X, C - X)
                    End If
                    Call GetaModiIID(RcvBuffer)
                    Call SendMessage_1200("SMP_REQ")
                Loop
            
            Case "SYS_READY"
            Case "SYS_NOT_READY"
            Case "SMP_NEW_DATA", "SMP_EDIT_DATA"
                GoTo RST
            Case "CAL_ABORT"
            Case "QC_START"
            Case "QC_NEW_AV"
                Do Until X = 0
                    X = InStr(X, RcvBuffer, "r")
                    If X = 0 Then Exit Do
                    If Mid(RcvBuffer, X, 4) = "rSEQ" Then
                        X = X + 5
                        C = InStr(X, RcvBuffer, GS)
                        Sample_Seq = Mid(RcvBuffer, X, C - X)
                    End If
                    Call GetaModiIID(RcvBuffer)
                    Call SendMessage_1200("QC_REQ")
                Loop
            Case "QC_NEW_DATA", "QC_EDIT_DATA"
                GoTo RST
        End Select
            
        Exit Sub

RST:
        MsgBuf = RcvBuffer
        
        If MsgID = "SMP_NEW_DATA" Or MsgID = "SMP_EDIT_DATA" Then
            'aMod
            x1 = 1
            x1 = InStr(x1, MsgBuf, "aMod") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                aMod = Mid(MsgBuf, x1, x2 - x1)
            End If
        
            'iIID
            x1 = 1
            x1 = InStr(x1, MsgBuf, "iIID") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                iIID = Mid(MsgBuf, x1, x2 - x1)
            End If
        
            'rSEQ
            x1 = 1
            x1 = InStr(x1, MsgBuf, "rSEQ") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                rSeq = Mid(MsgBuf, x1, x2 - x1)
            End If
        
            'PID
            x1 = 1
            x1 = InStr(x1, MsgBuf, "iPID") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                iPID = Mid(MsgBuf, x1, x2 - x1)
            End If
            'DATE
            x1 = 1
            x1 = InStr(x1, MsgBuf, "rDATE") + 6
            If x1 <> 6 Then
                x2 = InStr(x1, MsgBuf, GS)
                sRstDate = Mid(MsgBuf, x1, x2 - x1)
                sRstDate = ConvertDateType(sRstDate)
            End If
            'TIME
            x1 = 1
            x1 = InStr(x1, MsgBuf, "rTIME") + 6
            If x1 <> 6 Then
                x2 = InStr(x1, MsgBuf, GS)
                sRstTime = Mid(MsgBuf, x1, x2 - x1)
                sRstTime = Format(sRstTime, "HHNNSS")
            End If
        
            x2 = 0
        
            '접수번호, SeqNo
            strBarno = Trim(iPID)
            strSeq = Trim(rSeq)
            
            If strBarno = "" Or Not IsNumeric(strBarno) Then
                Exit Sub
            End If
            
            With mResult
                .BarNo = strBarno
                .RackNo = strRackNo
                .TubePos = strTubePos
                .Rerun = ""
                If strOldBarno <> strBarno Then
                    strOldBarno = strBarno
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
            
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                End If
            End With
            
            strState = "O"
                        
            '----------------------------------------------------------------------------------------
            '   Measured Data
            '----------------------------------------------------------------------------------------
            x1 = 1
            Do While InStr(x1, MsgBuf, FS & "m") <> 0
                x1 = InStr(x1, MsgBuf, FS & "m")
                x2 = InStr(x1, MsgBuf, GS)
        
        '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                strIntBase = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))
                Debug.Print strIntBase
                x2 = x2 + 1
                x1 = InStr(x2, MsgBuf, GS)
                strResult = Mid(MsgBuf, x2, x1 - x2)
        
                SetRawData "[결과]" & strIntBase & "," & strResult
                
                If strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    Else
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            strQCLab = Trim(RS_L.Fields("QCLab") & "")
                            strQCLot = Trim(RS_L.Fields("QCLot") & "")
                            strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                            strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                            strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                            strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                            strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                                
                            End If
                            
                            If strState <> "R" Then
                                strState = ""
                            End If
    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    End If
                End If
            Loop
                    
                    
            '----------------------------------------------------------------------------------------
            '   Calibrated Data
            '----------------------------------------------------------------------------------------
            x1 = 1
            Do While InStr(x1, MsgBuf, FS & "c") <> 0
                x1 = InStr(x1, MsgBuf, FS & "c")
                x2 = InStr(x1, MsgBuf, GS)
        
        '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                strIntBase = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))
                x2 = x2 + 1
                x1 = InStr(x2, MsgBuf, GS)
                strResult = Mid(MsgBuf, x2, x1 - x2)
                
                If strIntBase <> "" And strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    Else
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            strQCLab = Trim(RS_L.Fields("QCLab") & "")
                            strQCLot = Trim(RS_L.Fields("QCLot") & "")
                            strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                            strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                            strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                            strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                            strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                                
                            End If
                            
                            If strState <> "R" Then
                                strState = ""
                            End If
    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    End If
                End If
            Loop
            
            .spdResult.RowHeight(-1) = 14
            
            '## DB에 결과저장
            If .optTrans(0).Value = True And strState = "R" Then
                Res = SaveTransData_MCC(gRow)
                
                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "Failed", gRow, colSTATE
                Else
                    '-- 저장 성공
                    SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                    SetText .spdOrder, "저장완료", gRow, colSTATE
                    SetText .spdOrder, "0", gRow, colCHECKBOX
                    
                          SQL = "Update PATRESULT Set " & vbCrLf
                    SQL = SQL & " sendflag = '2' " & vbCrLf
                    SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                    SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                    SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                    SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                    End If
                End If
                strState = ""
            End If

            
        '>> If MsgID = "SMP_NEW_DATA" Or MsgID = "SMP_EDIT_DATA" Then
        
        ElseIf MsgID = "QC_NEW_DATA" Or MsgID = "QC_EDIT_DATA" Then
            '## Type 조회
            Pos1 = InStr(MsgBuf, "rTYPE")
            If Pos1 > 0 Then
                Pos2 = InStr(Mid$(MsgBuf, Pos1), FS)
                strBarno = mGetP(Mid$(MsgBuf, Pos1, Pos2), 2, GS)
                'strBarno = Val(strBarno)
            Else
                '## NOTE: WorkNo가 전송되지 않은 에러처리
                Exit Sub
            End If

            '## Level 조회
            Pos1 = 0: Pos2 = 0
            Pos1 = InStr(MsgBuf, "iQLEV")
            If Pos1 > 0 Then
                Pos2 = InStr(Mid$(MsgBuf, Pos1), FS)
                strQCLevel = mGetP(Mid$(MsgBuf, Pos1, Pos2), 2, GS)
            Else
                '## NOTE: 바코드번호가 전송되지 않은 에러처리
            End If


            '## QC 채널 조회
            Pos1 = 0: Pos2 = 0
            Pos1 = InStr(MsgBuf, "iQFILE")
            If Pos1 > 0 Then
                Pos2 = InStr(Mid$(MsgBuf, Pos1), FS)
                strQCChannel = mGetP(Mid$(MsgBuf, Pos1, Pos2), 2, GS)
            Else
                '## NOTE: 바코드번호가 전송되지 않은 에러처리
            End If
            
            strQCChannel = strQCLevel
            
'            x1 = 1
'            x1 = InStr(x1, MsgBuf, "aMod") + 5
'            If x1 <> 5 Then
'                x2 = InStr(x1, MsgBuf, GS)
'                aMod = Mid(MsgBuf, x1, x2 - x1)
'            End If
'
'            'iIID
'            x1 = 1
'            x1 = InStr(x1, MsgBuf, "iIID") + 5
'            If x1 <> 5 Then
'                x2 = InStr(x1, MsgBuf, GS)
'                iIID = Mid(MsgBuf, x1, x2 - x1)
'            End If
'
'            'rSEQ
'            x1 = 1
'            x1 = InStr(x1, MsgBuf, "rSEQ") + 5
'            If x1 <> 5 Then
'                x2 = InStr(x1, MsgBuf, GS)
'                rSeq = Mid(MsgBuf, x1, x2 - x1)
'            End If
'
'            'PID
'            x1 = 1
'            x1 = InStr(x1, MsgBuf, "iPID") + 5
'            If x1 <> 5 Then
'                x2 = InStr(x1, MsgBuf, GS)
'                iPID = Mid(MsgBuf, x1, x2 - x1)
'            End If
'            'DATE
'            x1 = 1
'            x1 = InStr(x1, MsgBuf, "rDATE") + 6
'            If x1 <> 6 Then
'                x2 = InStr(x1, MsgBuf, GS)
'                sRstDate = Mid(MsgBuf, x1, x2 - x1)
'                sRstDate = ConvertDateType(sRstDate)
'            End If
'            'TIME
'            x1 = 1
'            x1 = InStr(x1, MsgBuf, "rTIME") + 6
'            If x1 <> 6 Then
'                x2 = InStr(x1, MsgBuf, GS)
'                sRstTime = Mid(MsgBuf, x1, x2 - x1)
'                sRstTime = Format(sRstTime, "HHNNSS")
'            End If
'
'            x2 = 0
'
'            '접수번호, SeqNo
'            strBarno = Trim(iPID)
'            strSeq = Trim(rSeq)
'
'            If strBarno = "" Or Not IsNumeric(strBarno) Then
'                Exit Sub
'            End If
            
            With mResult
                .BarNo = strBarno
                .RackNo = strRackNo
                .TubePos = strTubePos
                .Rerun = ""
                .Kind = "QC"
                'If strOldBarno <> strBarno Then
                    'strOldBarno = strBarno
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
            
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                'End If
                    
                Call SetText(frmMain.spdOrder, strQCChannel, gRow, colPID)
                Call SetText(frmMain.spdOrder, strQCLevel, gRow, colPNAME)
                
            End With
            
            strState = "O"
                        
            '----------------------------------------------------------------------------------------
            '   Measured Data
            '----------------------------------------------------------------------------------------
            x1 = 1
            Do While InStr(x1, MsgBuf, FS & "m") <> 0
                x1 = InStr(x1, MsgBuf, FS & "m")
                x2 = InStr(x1, MsgBuf, GS)
        
        '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                strIntBase = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))
        
                x2 = x2 + 1
                x1 = InStr(x2, MsgBuf, GS)
                strResult = Mid(MsgBuf, x2, x1 - x2)
        
                SetRawData "[결과]" & strIntBase & "," & strResult
                
                If strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strQCChannel, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    Else
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            strQCLab = Trim(RS_L.Fields("QCLab") & "")
                            strQCLot = Trim(RS_L.Fields("QCLot") & "")
                            strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                            strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                            strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                            strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                            strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                
                                'strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strQCChannel, strQCAnalyte, strResult)
                                
                                
                                Call SendBioRadQC(strQCData)
                                
                            End If
                            
                            If strState <> "R" Then
                                strState = ""
                            End If
    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    End If
                End If
            Loop
                    
                    
            '----------------------------------------------------------------------------------------
            '   Calibrated Data
            '----------------------------------------------------------------------------------------
            x1 = 1
            Do While InStr(x1, strRcvBuf, FS & "c") <> 0
                x1 = InStr(x1, strRcvBuf, FS & "c")
                x2 = InStr(x1, strRcvBuf, GS)
        
        '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                strIntBase = Mid(strRcvBuf, x1 + 1, x2 - (x1 + 1))
                x2 = x2 + 1
                x1 = InStr(x2, strRcvBuf, GS)
                strResult = Mid(strRcvBuf, x2, x1 - x2)
                
                If strIntBase <> "" And strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    Else
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            strQCLab = Trim(RS_L.Fields("QCLab") & "")
                            strQCLot = Trim(RS_L.Fields("QCLot") & "")
                            strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                            strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                            strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                            strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                            strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                                
                            End If
                            
                            If strState <> "R" Then
                                strState = ""
                            End If
    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    End If
                End If
                
            Loop
            
            '## DB에 결과저장
            If .optTrans(0).Value = True And strState = "R" Then
                Res = SaveTransData_MCC(gRow)
                
                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "Failed", gRow, colSTATE
                Else
                    '-- 저장 성공
                    SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                    SetText .spdOrder, "저장완료", gRow, colSTATE
                    SetText .spdOrder, "0", gRow, colCHECKBOX
                    
                          SQL = "Update PATRESULT Set " & vbCrLf
                    SQL = SQL & " sendflag = '2' " & vbCrLf
                    SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                    SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                    SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                    SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                    End If
                End If
                strState = ""
            End If
            
            .spdResult.RowHeight(-1) = 14
            
            
        End If
        
        
        
    End With

End Sub

Private Sub SendMessage_1200(ByVal MsgHead As String)
    Dim chksum As Integer
    Dim Buffer As String
    Dim C As Integer
'    Dim R As Integer

    Dim sSendData$
    
    Select Case MsgHead
        Case "ID_DATA"
            Buffer = STX & "ID_DATA" & FS & R_S _
                                    & "aMOD" & GS & "LIS" & GS & GS & GS & FS _
                                    & "iIID" & GS & "333" & GS & GS & GS & FS & R_S _
                                    & ETX
        Case "SMP_REQ"
            Buffer = STX & "SMP_REQ" & FS & R_S & "aMOD" & GS & aMod & GS & GS & GS _
                                        & FS & "iIID" & GS & iIID & GS & GS & GS _
                                        & FS & "rSEQ" & GS & Sample_Seq & GS & GS & GS _
                                        & FS & R_S & ETX
            
        Case "QC_REQ"
            Buffer = STX & "QC_REQ" & FS & R_S & "aMOD" & GS & aMod & GS & GS & GS _
                                        & FS & "iIID" & GS & iIID & GS & GS & GS _
                                        & FS & "rSEQ" & GS & Sample_Seq & GS & GS & GS _
                                        & FS & R_S & ETX
            
        Case "SMP_ORD"
    End Select
        
    For C = 1 To Len(Buffer)
        chksum = chksum + Asc(Mid(Buffer, C, 1))
    Next C
    
    sSendData = Buffer & Right("0" & Hex(chksum Mod 256), 2) & EOT
    
    frmMain.comEqp.Output = sSendData
    
    SetRawData "[Tx]" & sSendData
    
End Sub

Private Sub GetaModiIID(ByVal sMsg As String)

    Dim tmpData()   As String
    
    '<STX>SYS_READY<FS><RS>aMOD<GS>1265<GS><GS><GS><FS>iIID
    '<GS>12345<GS><GS><GS><FS>aDATE<GS>20Jan2004<GS><GS><GS>
    '<FS>aTIME<GS>13:35:32<GS><GS><GS><FS>iOID<GS>3<GS><GS><GS><FS>
    '<ETX>{chksum}<EOT>

    tmpData() = Split(sMsg, GS)
    
    'aMod
    aMod = Trim(tmpData(1))
    
    'iIID
    iIID = Trim(tmpData(5))

End Sub


Private Function ConvertDateType(ByVal sDate As String) As String
    On Error GoTo ErrRtn
    
    Dim kk%
    Dim sTmp$
    Dim tmpYYYY$, tmpMM$, tmpDD$
    
    ConvertDateType = sDate
    
    tmpYYYY = Right(sDate, 4)
    sDate = Mid(sDate, 1, Len(sDate) - 4)
    
    For kk = 1 To Len(sDate)
        sTmp = Mid(sDate, kk, 1)
        If IsNumeric(sTmp) Then
            tmpDD = tmpDD & sTmp
        Else
            tmpMM = tmpMM & sTmp
        End If
    Next kk
    
    sTmp = tmpDD & Space(1) & tmpMM & Space(1) & tmpYYYY
    
    ConvertDateType = Format(sTmp, "YYYYMMDD")
    
ErrRtn:
    If Err <> 0 Then
        'RaiseEvent DispMsg("ConvertDateType - " & Err.Description)
    End If
End Function


Private Sub Phase_Serial_RAPIDPOINT500()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                AckOn = False
                RcvBuffer = BufChar
                
            Case EOT
                If AckOn = False Then
                    frmMain.comEqp.Output = STX & ACK & ETX & "0B" & EOT        'Ack Message
                    Call SerialRcvData_RAPIDPOINT500
                End If
            
            Case ACK
                AckOn = True
                RcvBuffer = RcvBuffer & BufChar
            
            Case Else
                RcvBuffer = RcvBuffer & BufChar
                
        End Select
    Next i
            
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_ACLTOP()


    Dim strOutput   As String     '송신할 데이터
    Dim blnLast     As Boolean
    Dim intRow      As Integer
    Dim strBarno    As String
    Dim strItems    As String

    blnLast = False

    With frmMain.spdOrder
        If intSndPhase <= 3 Then
            For intRow = 1 To .DataRowCnt
                If GetText(frmMain.spdOrder, intRow, colCHECKBOX) = "1" And GetText(frmMain.spdOrder, intRow, colSTATE) = "오더준비" Then
                    strBarno = Trim(GetText(frmMain.spdOrder, intRow, colBARCODE))
                    strItems = Trim(GetText(frmMain.spdOrder, intRow, colKEY1))
                    If intSndPhase = 3 Then
                        .Row = intRow
                        .Col = colCHECKBOX: .Text = "0"
                        .Col = colSTATE:    .Text = "오더전송"

                        If intRow = .DataRowCnt Then
                            blnLast = True
                        End If

                    End If
                    Exit For
                End If
            Next
        End If
    End With

    If intRow = frmMain.spdOrder.DataRowCnt Then
        blnLast = True
    End If

    Select Case intSndPhase
        Case 1  '## Header
        '''''            strOutput = "H|@^\|" & mOrder.MsgID & "||" & mOrder.Receiver & "|||||" & mOrder.Sender & "||P|" & mOrder.Version & "|" & Format(Now, "yyyyMMddHHmmss") & vbCr
            strOutput = intFrameNo & "H|@^\|" & mOrder.MsgID & "||" & mOrder.Receiver & "|||||" & mOrder.Sender & "||P|" & mOrder.Version & "|" & Format(Now, "yyyyMMddHHmmss") & vbCr & ETB
            intSndPhase = 2
            intFrameNo = intFrameNo + 1

        Case 2  '## Patient
'''''        strOutput = strOutput & "P|" & mPNo & "||||^||||||||" & vbCr
            strOutput = intFrameNo & "P|" & mPNo & "||||^||||||||" & vbCr & ETB
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
            mPNo = mPNo + 1

        Case 3  '## Order
            '## 최초 보낼때
            If mOrder.IsSending = False Then
'''''         = strOutput & "O|1|" & strBarno & "||" & strItems & "|R|" & Format(Now, "yyyyMMddHHmmss") & "|||||A||||P||||||||||Q" & vbCr
                strOutput = "O|1|" & strBarno & "||" & strItems & "|R|" & Format(Now, "yyyyMMddHHmmss") & "|||||A||||P||||||||||Q"
                If Len(strOutput) > 230 Then
                    mOrder.IsSending = True
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 3
                Else
                    strOutput = intFrameNo & strOutput & vbCr & ETB
                    If blnLast = True Then
                        intSndPhase = 4
                    Else
                        intSndPhase = 2
                    End If
                End If
            '## 남은 문자열이 있을때
            Else
                strOutput = mOrder.Order
                If Len(strOutput) > 230 Then
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 3
                Else
                    mOrder.IsSending = False
                    strOutput = intFrameNo & strOutput & vbCr & ETB
                    If blnLast = True Then
                        intSndPhase = 4
                    Else
                        intSndPhase = 2
                    End If
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
'''''            strOutput = strOutput & "L|1|N"
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1

            Exit Sub
    End Select


    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    frmMain.comEqp.Output = strOutput
    SetRawData "[Tx]" & strOutput

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_LIAISON()


    Dim strOutput   As String     '송신할 데이터
    Dim blnLast     As Boolean
    Dim intRow      As Integer
    Dim strBarno    As String
    Dim strPtId     As String
    Dim strItems    As String

    blnLast = False

    With frmMain.spdOrder
        If intSndPhase <= 3 Then
            For intRow = 1 To .DataRowCnt
                If GetText(frmMain.spdOrder, intRow, colCHECKBOX) = "1" And GetText(frmMain.spdOrder, intRow, colSTATE) = "오더준비" Then
                    strBarno = Trim(GetText(frmMain.spdOrder, intRow, colBARCODE))
                    strPtId = Trim(GetText(frmMain.spdOrder, intRow, colPID))
                    strItems = Trim(GetText(frmMain.spdOrder, intRow, colKEY1))
                    If intSndPhase = 3 Then
                        .Row = intRow
                        .Col = colCHECKBOX: .Text = "0"
                        .Col = colSTATE:    .Text = "오더전송"

                        If intRow = .DataRowCnt Then
                            blnLast = True
                        End If

                    End If
                    Exit For
                End If
            Next
        End If
    End With

    If intRow = frmMain.spdOrder.DataRowCnt Then
        blnLast = True
    End If

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&|||" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
            mPNo = 0
            
        Case 2  '## Patient
            If strBarno = "" Then
                strOutput = intFrameNo & "L|1|N" & vbCr & ETX
                intSndPhase = 5
                intFrameNo = intFrameNo + 1
            Else
                strOutput = intFrameNo & "P|" & CStr(mPNo) & "||" & strPtId & "||^|||||||||||||||||||||" & vbCr & ETX
                intSndPhase = 3
                intFrameNo = intFrameNo + 1
                mPNo = mPNo + 1
            End If
            
        Case 3  '## Order
            strOutput = intFrameNo & "O|1|" & strBarno & "||" & strItems & "|N|" & Format(Now, "yyyymmdd") & "|||||||||||||||||||O|" & vbCr & ETX
            
            If blnLast = True Then
                intSndPhase = 4
            Else
                intSndPhase = 2
            End If
            
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1
            intSndPhase = 1
            Exit Sub
    End Select


    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    frmMain.comEqp.Output = strOutput
    SetRawData "[Tx]" & strOutput

End Sub


Private Sub Phase_Serial_ACLTOP()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        intBufCnt = 0
                        Erase strRecvData
                        intPhase = 2
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_ACLTOP
                        Else
                            frmMain.comEqp.Output = ACK
                            SetRawData "[Tx]" & ACK
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                        Else
                            intBufCnt = intBufCnt + 1
                        End If
                        ReDim Preserve strRecvData(intBufCnt)
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intPhase = 3
                    Case EOT
                        intPhase = 1
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case vbLf
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = IIf(blnIsETB = False, 4, 2)
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        Call SerialRcvData_ACLTOP
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            frmMain.comEqp.Output = ENQ
                            SetRawData "[Tx]" & ENQ
                        End If
                        intPhase = 1
                End Select
        End Select
    Next i

End Sub

Private Sub Phase_Serial_LIAISON()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        intBufCnt = 0
                        Erase strRecvData
                        intPhase = 2
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_LIAISON
                        Else
                            frmMain.comEqp.Output = ACK
                            SetRawData "[Tx]" & ACK
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                        Else
                            intBufCnt = intBufCnt + 1
                        End If
                        ReDim Preserve strRecvData(intBufCnt)
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intPhase = 3
                    Case EOT
                        intPhase = 1
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case vbLf
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = IIf(blnIsETB = False, 4, 2)
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        Call SerialRcvData_LIAISON
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            frmMain.comEqp.Output = ENQ
                            SetRawData "[Tx]" & ENQ
                        End If
                        intPhase = 1
                End Select
        End Select
    Next i

End Sub

Private Sub Phase_Serial_PATHFAST()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case BufChar
            Case ENQ
                intBufCnt = 1
                Erase strRecvData
                ReDim Preserve strRecvData(intBufCnt)
                comEqp.Output = ACK
                SetRawData "[Tx]" & ACK
            Case STX
                intBufCnt = intBufCnt + 1
                ReDim Preserve strRecvData(intBufCnt)
                
            Case vbLf
                comEqp.Output = ACK
                SetRawData "[Tx]" & ACK
                
            Case EOT
                Call SerialRcvData_PATHFAST
                intBufCnt = 1
                Erase strRecvData
                ReDim Preserve strRecvData(intBufCnt)
            
            Case Else
                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
        End Select
    Next i

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
'                                 2. 장비수신정보 화면표시,
'                                 3. 처방코드 가져오기,
'                                 4. (처방코드로)검사오더 만들기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Private Sub GetOrder_ACLTOP(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String

    intRow = -1

    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 12

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_ACLTOP(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""

            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더준비", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems

            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더준비", intRow, colSTATE)
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, strItems, intRow, colKEY1)
        End If

        SetText frmMain.spdOrder, "1", intRow, colCHECKBOX

        '-- 현재 Row
        gRow = intRow

    End With

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
'                                 2. 장비수신정보 화면표시,
'                                 3. 처방코드 가져오기,
'                                 4. (처방코드로)검사오더 만들기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Private Sub GetOrder_LIAISON(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String

    intRow = -1

    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, pBarno, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 12

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_LIAISON(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""

            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더준비", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems

            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더준비", intRow, colSTATE)
            
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, strItems, intRow, colKEY1)
        End If

        SetText frmMain.spdOrder, "1", intRow, colCHECKBOX

        '-- 현재 Row
        gRow = intRow

    End With

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
'                                 2. 장비수신정보 화면표시,
'                                 3. 처방코드 가져오기,
'                                 4. (처방코드로)검사오더 만들기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Private Sub GetOrder_BS240_HL7(ByVal pBarno As String, ByVal pType As String) ', _
                               ByVal pMaker As String, ByVal pMchNm As String, _
                               ByVal pModel As String, ByVal pSTime As String, _
                               ByVal pETime As String, ByVal pQryId As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    Dim strSend     As String
    Dim blnLast     As Boolean
    
    
    Dim iCnt As Integer
    
    For iCnt = mOrder.BSBeginNo To mOrder.BSFinishNo
        
        mOrder.Seq = iCnt
        
        intRow = -1
    
        '-- 1. 접수정보 조회
        With frmMain
            '-- 바코드 사용
            If .optBarSeq(0).Value = True Then
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                        intRow = i
                        Exit For
                    End If
                Next i
            Else
                Select Case pType
                    '-- Seq
                    Case "1"
                        For i = 1 To .spdOrder.DataRowCnt
                            If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" And Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                                pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                                mOrder.BarNo = pBarno
                                intRow = i
                                Exit For
                            End If
                        Next i
                    '-- Rack/Pos
                    Case "2"
                        For i = 1 To .spdOrder.DataRowCnt
                            If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                                pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                                intRow = i
                                Exit For
                            End If
                        Next i
                    '-- Check Top
                    Case "3"
                        For i = 1 To .spdOrder.DataRowCnt
                            If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                                pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                                mOrder.BarNo = pBarno
                                mOrder.PName = Trim(GetText(frmMain.spdOrder, i, colPNAME))
                                mOrder.PID = Trim(GetText(frmMain.spdOrder, i, colPID))
                                intRow = i
                                Exit For
                            End If
                        Next i
                End Select
            End If
    
            '-- 스프레드에서 못찾았음..
            If intRow < 0 Then
                Exit Sub
                'intRow = .spdOrder.DataRowCnt + 1
                'If .spdOrder.MaxRows < intRow Then
                '    .spdOrder.MaxRows = intRow
                'End If
            End If
    
            '-- 장비수신정보 화면표시
    
            '-- 결과스프레드 지우기
            .spdResult.MaxRows = 0
    
            '-- 검사자 정보 가져오기
            Call GetSampleInfo(intRow, .spdOrder)
    
            .spdOrder.RowHeight(-1) = 12
    
            '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
            strItems = GetEquipExamCode_BS240_HL7(gHOSP.MACHCD, pBarno, intRow)
    
            '-- 검사채널로 장비오더 만들기
            If Trim(strItems) = "" Then
                mOrder.NoOrder = True
                mOrder.Order = ""
    
                '-- 진행상태(Order) 표시
                Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
                          
                '-- 오더문자열만들기(오더없을때)
                          strSend = SB & "MSH|^~\&|||" & mOrder.BSMaker & "|" & mOrder.BSMchNm & "|" & mOrder.BSDtTm & "||DSR^Q03|1|P|2.3.1||||0||ASCII|||" & vbCr
                strSend = strSend & "MSA|AA|" & CStr(intSndPhase) & "|Message accepted|||0|" & vbCr
                strSend = strSend & "ERR|0|" & vbCr
                strSend = strSend & "QAK|SR|NF|" & vbCr
                strSend = strSend & EB & vbCr
            
            Else
                mOrder.NoOrder = False
                mOrder.Order = strItems
    
                '-- 진행상태(Order) 표시
                Call SetText(frmMain.spdOrder, "오더준비", intRow, colSTATE)
                
                '-- 오더문자열만들기
                          strSend = SB & "MSH|^~\&|||" & mOrder.BSMaker & "|" & mOrder.BSMchNm & "|" & mOrder.BSDtTm & "||DSR^Q03|1|P|2.3.1||||0||ASCII|||" & vbCr
                strSend = strSend & "MSA|AA|" & CStr(intSndPhase) & "|Message accepted|||0|" & vbCr
                strSend = strSend & "ERR|0|" & vbCr
                strSend = strSend & "QAK|SR|OK|" & vbCr
                'strSend = strSend & "QRD|" & mOrder.BSDtTm & "|R|D|" & mOrder.BSQryId & "|||RD|" & mOrder.BarNo & "|OTH|||T|" & vbCr
                strSend = strSend & "QRD|" & mOrder.BSDtTm & "|R|D|" & mOrder.BSQryId & "|||RD||OTH|||T|" & vbCr
                
                '                           QRF|||                                                          '|||RCT|COR|ALL||
    
                
                        'mOrder.BSModel = mGetP(strRcvBuf, 2, "|")
                        'mOrder.BSSTime = mGetP(strRcvBuf, 3, "|")
                        'mOrder.BSETime = mGetP(strRcvBuf, 4, "|")
                        
                
                strSend = strSend & "QRF|" & mOrder.BSModel & "|" & mOrder.BSSTime & "|" & mOrder.BSETime & "|||RCT|COR|ALL||" & vbCr
                
               ' strSend = strSend & mOrder.BSQRF & vbCr
                
                strSend = strSend & "DSP|1|||||" & vbCr       'HOSPITAL ID
                strSend = strSend & "DSP|2|||||" & vbCr                         'Bed No
                'strSend = strSend & "DSP|3||" & mOrder.PName & "-" & mOrder.BarNo & "|||" & vbCr      'Patient Name - 바코드번호
                strSend = strSend & "DSP|3||" & mOrder.PName & "|||" & vbCr       'Patient Name
                strSend = strSend & "DSP|4|||||" & vbCr                         'Birth Date
                strSend = strSend & "DSP|5|||||" & vbCr                         'Gender
                strSend = strSend & "DSP|6|||||" & vbCr                         'Blood Type
                strSend = strSend & "DSP|7|||||" & vbCr
                strSend = strSend & "DSP|8|||||" & vbCr
                strSend = strSend & "DSP|9|||||" & vbCr
                strSend = strSend & "DSP|10|||||" & vbCr
                strSend = strSend & "DSP|11|||||" & vbCr
                strSend = strSend & "DSP|12|||||" & vbCr
                strSend = strSend & "DSP|13|||||" & vbCr
                strSend = strSend & "DSP|14|||||" & vbCr
                strSend = strSend & "DSP|15|||||" & vbCr                        'Patient Type
                strSend = strSend & "DSP|16|||||" & vbCr
                strSend = strSend & "DSP|17|||||" & vbCr                        'Pay type
                strSend = strSend & "DSP|18|||||" & vbCr
                strSend = strSend & "DSP|19|||||" & vbCr
                strSend = strSend & "DSP|20|||||" & vbCr
                '바코드로 사용할때만 필요
                'If gHOSP.BARUSE = "Y" Then
                    strSend = strSend & "DSP|21||" & mOrder.BarNo & "|||" & vbCr                    'barcode
                'Else
                '    strSend = strSend & "DSP|21|||||" & vbCr                    'barcode
                'End If
                strSend = strSend & "DSP|22||" & mOrder.PID & "|||" & vbCr                'sample id
                strSend = strSend & "DSP|23||" & Format(Now, "yyyymmddhhmmss") & "|||" & vbCr   'sending time
                strSend = strSend & "DSP|24||N|||" & vbCr                                       'STAT (응급여부)
                strSend = strSend & "DSP|25|||||" & vbCr
                strSend = strSend & "DSP|26||serum|||" & vbCr                                   'sample type
                strSend = strSend & "DSP|27|||||" & vbCr                                        'doctor(처방의)
                strSend = strSend & "DSP|28|||||" & vbCr                                        'send department
                strSend = strSend & strItems
                
                blnLast = False
                For i = 1 To spdOrder.DataRowCnt
                    If Trim(GetText(spdOrder, i, colCHECKBOX)) = "1" And Trim(GetText(spdOrder, i, colSTATE)) = "" Then
                        blnLast = True
                        Exit For
                    End If
                Next i
                
                If blnLast = True Then
                    strSend = strSend & "DSC|" & CStr(intSndPhase) & "|" & vbCr
                Else
                    strSend = strSend & "DSC||" & vbCr
                End If
                
                strSend = strSend & EB & vbCr
                
                '-- 진행상태(Order) 표시
                Call SetText(frmMain.spdOrder, "0", intRow, colCHECKBOX)
                Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
                
                
                SetRawData "[Tx]" & strSend
                wSck.SendData strSend
                intSndPhase = intSndPhase + 1
            End If
    
            '-- 현재 Row
            gRow = intRow
    
        End With
    Next
    
End Sub

Private Sub GetOrder_BS240(ByVal pBarno As String, ByVal pType As String) ', _
                               ByVal pMaker As String, ByVal pMchNm As String, _
                               ByVal pModel As String, ByVal pSTime As String, _
                               ByVal pETime As String, ByVal pQryId As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    Dim strSend     As String
    Dim blnLast     As Boolean
    
    intRow = -1

    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            mOrder.PName = Trim(GetText(frmMain.spdOrder, i, colPNAME))
                            mOrder.PID = Trim(GetText(frmMain.spdOrder, i, colPID))
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            Exit Sub
            'intRow = .spdOrder.DataRowCnt + 1
            'If .spdOrder.MaxRows < intRow Then
            '    .spdOrder.MaxRows = intRow
            'End If
        End If

        '-- 장비수신정보 화면표시

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 12

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_BS240(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""

            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems

            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더준비", intRow, colSTATE)
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)

        End If

        '-- 현재 Row
        gRow = intRow

    End With

End Sub


'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_ACLTOP(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_ACLTOP = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    strExamCode = ""

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strExamCode = strExamCode & "@^^^" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close

    GetEquipExamCode_ACLTOP = Mid(strExamCode, 2)

End Function

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_LIAISON(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_LIAISON = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    strExamCode = ""

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strExamCode = strExamCode & "\^^^" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "^")
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close

    GetEquipExamCode_LIAISON = Mid(strExamCode, 2)

End Function

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_BS240_HL7(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_BS240_HL7 = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    strExamCode = ""
    i = 0
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strExamCode = strExamCode & "DSP|" & CStr(28 + i) & "||" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value) & "" & "^^^|||" & vbCr
            i = i + 1
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close

    GetEquipExamCode_BS240_HL7 = strExamCode

End Function

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_BS240(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_BS240 = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    strExamCode = ""
    i = 0
'Assay No.
'12
'Code
'Assay Name
'0-99999.0
'Max = 99999#
'20
'Test
'Dilution Rate
'String
'4
'Void and reserved.
'Repeat Num
'String
'2
'Void and reserved
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strExamCode = strExamCode & "\" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "^^^")
            
            i = i + 1
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close

    GetEquipExamCode_BS240 = Mid(strExamCode, 2)

End Function

Private Sub SerialRcvData_ACLTOP()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strAspect       As String

    Dim strTemp1        As String
    Dim strTemp2        As String

    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq

    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim varBarno        As Variant
    Dim i               As Integer

    Dim strUseRes       As String

    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)

            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If

            Select Case strType
                Case "H"    '## Header
                    '1H|@^\|<1504128210_21570><1504128210_21571>||acl|||||LIS||P|1394-97|20170830172330
                    mOrder.MsgID = Trim(mGetP(strRcvBuf, 3, "|"))
                    mOrder.Sender = Trim(mGetP(strRcvBuf, 5, "|"))
                    mOrder.Receiver = Trim(mGetP(strRcvBuf, 10, "|"))
                    mOrder.Version = Trim(mGetP(strRcvBuf, 13, "|"))

                Case "P"    '## Patient
                Case "Q"    '## Request Information

                    'Q|1|^1001@^1002@^1003@^1004@^1005@^1006@^1008||||||||||O@N
                    'Q|1|^198772||||||||||O@N
                    'Q|1|^1310250941@^1310250867||||||||||O@N


                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strTemp1 = Replace(strTemp1, "^", "")

                    varBarno = Split(strTemp1, "@")

                    For i = 0 To UBound(varBarno)
                        mOrder.BarNo = varBarno(i)
                        Call GetOrder_ACLTOP(varBarno(i), gHOSP.RSTTYPE)
                    Next

'                    With mOrder
'                        .NoOrder = False
'                        .BarNo = strBarno
'                        .Seq = mGetP(strTemp1, 3, "^")
'                        .RackNo = mGetP(strTemp1, 4, "^")
'                        .TubePos = mGetP(strTemp1, 5, "^")
'                    End With

                   ' Call GetOrder(strBarno, gHOSP.RSTTYPE)
                    strState = "Q"
                    mPNo = 1

                Case "O"
                    strBarno = mGetP(strRcvBuf, 3, "|")

                    With mResult
                        .BarNo = strBarno
                        '.SpcPos = strTubePos & "/" & strRackNo
                        '.Seq = strSeq
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        If strOldBarno <> strBarno Then
                            strOldBarno = strBarno
                            .RsltDate = Format(Now, "yyyymmddhhmmss")
                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))

                            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)

                        End If
                    End With

                Case "R"
                    'R|1|^^^131|28.4|s||N||F@V||SysAdmin^SysAdmin||20170826150315|
                    'R|1|^^^541|103.6|D mAbs||N||F@V||SysAdmin^SysAdmin||2017090108
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    If strIntBase = "131" Then
                        strIntBase = strIntBase & UCase(mGetP(strRcvBuf, 5, "|"))
                    End If

                    'R|1|^^^2241|0.3|microg/mLFEU||N||F@V||SysAdmin^SysAdmin||2017
                    'R|2|^^^2241|172|ng/mL||N||F@V||SysAdmin^SysAdmin||20170901083115|acl^03^2

                    ' D-Dimer
                    If strIntBase = "2241" Then
                        If mGetP(strRcvBuf, 5, "|") = "microg/mLFEU" Then
                            strIntBase = strIntBase
                        Else
                            strIntBase = ""
                        End If
                    End If

                    strResult = mGetP(strRcvBuf, 4, "|")

                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "

                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")

                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If

                                '-- 소수점 처리, 결과 형태 처리
                                If strUseRes <> "" Then
                                    strMachResult = strResult
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)

                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE

                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next

                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치

                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""

                                '-- BIORAD QC 저장
                                If Mid(strBarno, 1, 2) = "QC" Then
                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
                                End If


                                strState = "R"

                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If

                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "

                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""

                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")

                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If

                                '-- 소수점 처리, 결과 형태 처리
                                If strUseRes <> "" Then
                                    strMachResult = strResult
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)


                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE

                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next

                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치

                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""

                                '-- BIORAD QC 저장
                                If Mid(strBarno, 1, 2) = "QC" Then
                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
                                End If

                                If strState <> "R" Then
                                    strState = ""
                                End If

                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If

                        End If

                    End If

                    .spdResult.RowHeight(-1) = 14

                Case "C"    '## Comment
                    '## Abnormal 결과일때 Comment 저장
                    If strFlag <> "N" Then
                        strTemp1 = mGetP(strRcvBuf, 4, "|")
                        strComm = mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")
                    End If

                Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_MCC(gRow)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

End Sub

Private Sub SerialRcvData_LIAISON()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strAspect       As String

    Dim strTemp1        As String
    Dim strTemp2        As String

    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq

    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim varBarno        As Variant
    Dim i               As Integer

    Dim strUseRes       As String

    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)

            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If

            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
'''                    strTemp1 = mGetP(strRcvBuf, 3, "|")
'''                    strTemp1 = Replace(strTemp1, "^", "")
'''
'''                    varBarno = Split(strTemp1, "@")
'''
'''                    For i = 0 To UBound(varBarno)
'''                        mOrder.BarNo = varBarno(i)
'''                        Call GetOrder_ACLTOP(varBarno(i), gHOSP.RSTTYPE)
'''                    Next

                    strBarno = mGetP(strRcvBuf, 3, "|")
                    With mOrder
                        .BarNo = strBarno
                    End With
                    
                    Call GetOrder_LIAISON(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                    mPNo = 0

                Case "P"    '## Patient
                    strState = "P"
                Case "O"
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_KYU(gRow)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
                
                    strState = "O"
                    
                    strBarno = mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^")

                    With mResult
                        .BarNo = strBarno
                        If strOldBarno <> strBarno Then
                            strOldBarno = strBarno
                            .RsltDate = Format(Now, "yyyymmddhhmmss")
                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))

                            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)

                        End If
                    End With

                Case "R"
                    strIntBase = Trim$(mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^"))
                    strIntResult = mGetP(strRcvBuf, 4, "|")
                    strResult = mGetP(strRcvBuf, 7, "|")
                    
                    If IsNumeric(strIntResult) Then
                        strIntResult = SetResult(strIntResult, strIntBase)
                    End If
                    
                    'strIntResult = Replace(strIntResult, "<", "")
                    'strIntResult = Replace(strIntResult, ">", "")
            
'                    If strIntBase = "Myco-M" Then
'                        If strIntResult >= 27 Then
'                            strResult = "POSITIVE:" & ">" & strIntResult
'                        ElseIf strIntResult >= 10 Then
'                            strResult = "POSITIVE:" & strIntResult
'                        ElseIf strIntResult <= 0.1 Then
'                            strResult = "NEGATIVE:" & "<" & strIntResult
'                        ElseIf strIntResult < 10 Then
'                            strResult = "NEGATIVE:" & strIntResult
'                        Else
'                            strResult = "EQUIVOCAL:" & strIntResult
'                        End If
'                    End If
                    
                    If strIntBase = "IGF-I" Then
                        '수치만
                        strResult = strIntResult
                    Else
                        If Not IsNumeric(strIntResult) Then
                            If strResult = "N" Then
                                strResult = "Negative" & "(" & strIntResult & ")"
                            ElseIf strResult = "P" Then
                                strResult = "Positive" & "(" & strIntResult & ")"
                            ElseIf strResult = "<" Then
                                strResult = "Negative" & "(" & strIntResult & ")"
                            ElseIf strResult = ">" Then
                                strResult = "Positive" & "(" & strIntResult & ")"
                            
                            Else
                                strResult = strResult & "(" & strIntResult & ")"
                            End If
                        Else
                            Select Case strIntBase
                                Case "TXAB" 'C. difficile Toxins A&B
                                        If strIntResult < 0.9 Then
                                            strResult = "Negative"
                                        ElseIf strIntResult >= 1.1 Then
                                            strResult = "Positive"
                                        Else
                                            strResult = "Equivocal"
                                        End If
                                
                                Case "Myco-M" 'Mycoplasma IgM
                                        If strIntResult < 10 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 10 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "MeaslesG" 'Measles IgG
                                        If strIntResult < 13.5 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 16.5 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "Mump-G" 'Mumps IgG
                                        If strIntResult < 9 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 11 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "Mump-M" 'Mumps IgM
                                        If strIntResult < 0.9 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 1.1 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "VZV-G" 'VZV IgG
                                        If strIntResult < 135 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 165 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "VZV-M" 'VZV IgM
                                        If strIntResult < 0.9 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 1.1 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "EBV-M" 'EBV IgM
                                        If strIntResult < 20 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 40 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "EBNA-G" 'EBNA IgG
                                        If strIntResult < 5 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 20 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "VCA-G" 'VCA IgG
                                        If strIntResult < 20 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 20 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "EA-G" 'EA IgG
                                        If strIntResult < 10 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 40 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "Tox-G2" 'Toxo IgG II
                                        If strIntResult < 7.2 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 8.8 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "ToxoM" 'Toxo IgM
                                        If strIntResult < 6 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 8 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "RubG" 'Rubella IgG
                                        If strIntResult < 10 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 10 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "RubM" 'Rubella IgM
                                        If strIntResult < 20 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 25 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "CMVGII" 'CMV IgG II
                                        If strIntResult < 12 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 14 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "CMV-MII" 'CMV IgM II
                                        If strIntResult < 18 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 22 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "HSV-G" 'HSV 1/2 IgG
                                        If strIntResult < 0.9 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 1.1 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "HSV-M" 'HSV 1/2 IgM
                                        If strIntResult < 0.9 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 1.1 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                            End Select
                        End If
'                        If strResult = "N" Then
'                            strResult = "Negative" & "(" & strIntResult & ")"
'                        ElseIf strResult = "P" Then
'                            strResult = "Positive" & "(" & strIntResult & ")"
'                        ElseIf strResult = "<" Then
'                            strResult = "Negative" & "(" & strIntResult & ")"
'                        ElseIf strResult = ">" Then
'                            strResult = "Positive" & "(" & strIntResult & ")"
'                        Else
'                            strResult = ""
'                        End If
                    
                    End If
                    
                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "

                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")

                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If

                                '-- 소수점 처리, 결과 형태 처리
                                If strUseRes <> "" Then
                                    strMachResult = strResult
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)

                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE

                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next

                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치

                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""

                                '-- BIORAD QC 저장
                                'If Mid(strBarno, 1, 2) = "QC" Then
                                '    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
                                'End If

                                strState = "R"

                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If

                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "

                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""

                                'strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                'strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                'strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                'strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                'strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                'strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                'strQCTemp = Trim(RS_L.Fields("QCTemp") & "")

                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If

                                '-- 소수점 처리, 결과 형태 처리
                                If strUseRes <> "" Then
                                    strMachResult = strResult
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)


                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE

                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next

                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치

                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""

                                '-- BIORAD QC 저장
                                'If Mid(strBarno, 1, 2) = "QC" Then
                                '    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
                                'End If

                                If strState <> "R" Then
                                    strState = ""
                                End If

                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If

                        End If

                    End If

                    .spdResult.RowHeight(-1) = 14

                'Case "C"    '## Comment

                'Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_KYU(gRow)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

End Sub

Private Sub SerialRcvData_PATHFAST()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim strAspect       As String

    Dim strTemp1        As String
    Dim strTemp2        As String

    Dim strOrderCode     As String   '처방코드
    Dim strTestCode      As String   '검사코드
    Dim strTestName      As String   '검사명
    Dim strSeqNo         As String   '로컬DB 검사Seq

    Dim strRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim varBarno        As Variant
    Dim i               As Integer

    Dim strUseRes       As String
'    Dim blnOrder        As Boolean

'    blnOrder = False

    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)

            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If

            Select Case strType
                Case "H"    '## Header
                    mResult.BarNo = ""
                Case "P"    '## Patient
                Case "Q"    '## Request Information
                Case "O"
                    strBarno = mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^")
                    strSeq = mGetP(mGetP(strRcvBuf, 3, "|"), 2, "^")
                    If strBarno = "" Then
                        strBarno = strSeq
                    '    Exit Sub
                    End If
                    
                    With mResult
                        '.BarNo = strBarno
                        '.RackNo = strRackNo
                        '.TubePos = strTubePos
                        If mResult.BarNo = "" Then
                            'strOldBarno = strBarno
                            .RsltDate = Format(Now, "yyyymmddhhmmss")
                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        End If
                    End With
                    
                Case "R"
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                    strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")

                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "

                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                'strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""
                                'strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")

                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If

                                '-- 소수점 처리, 결과 형태 처리
                                If strUseRes <> "" Then
                                    strMachResult = strResult
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)

                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE

                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next

                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치

                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""

                                '-- BIORAD QC 저장
                                'If Mid(strBarno, 1, 2) = "QC" Then
                                '    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
                                'End If


                                strState = "R"

                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If

                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "

                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""

                                'strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                'strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                'strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                'strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                'strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                'strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                'strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                'strQCTemp = Trim(RS_L.Fields("QCTemp") & "")

                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If

                                '-- 소수점 처리, 결과 형태 처리
                                If strUseRes <> "" Then
                                    strMachResult = strResult
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)


                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE

                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next

                                '-- 결과 List
                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치

                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""

                                '-- BIORAD QC 저장
                                'If Mid(strBarno, 1, 2) = "QC" Then
                                '    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
                                'End If

                                If strState <> "R" Then
                                    strState = ""
                                End If

                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If

                        End If

                    End If

                    .spdResult.RowHeight(-1) = 14

                Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_AMIS(gRow)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

End Sub


Private Sub comEQP_OnComm()
    Dim EVMsg       As String
    Dim ERMsg       As String
    Dim Ret         As Long
    Dim strDate     As String
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    
    
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
            
            If fraInterface.Visible = False Then
                tmrComm.Interval = 20000
                tmrComm.Enabled = True
                
                tmrFlipFlop.Interval = 500
                tmrFlipFlop.Enabled = True
                
                lblCommStatus.Caption = "장비 검사결과가 수신되었습니다. 인터페이스 창에서 확인하세요!"
            End If
            
            SetRawData "[Rx]" & pBuffer

            
            Select Case UCase(gHOSP.MACHNM)
                Case "BS240"
                    Call Phase_Serial_BS240
                    
                Case "XP300"
                    Call Phase_Serial_XP300
                
                Case "XI921F"
                    Call Phase_Serial_XI921F
                
                Case "LIAISON"
                    Call Phase_Serial_LIAISON
                    
                Case "URINSCANPRO"

                    lngBufLen = Len(pBuffer)
                
                    For i = 1 To lngBufLen
                        BufChar = Mid$(pBuffer, i, 1)
                        Select Case intPhase
                            Case 1
                                Select Case BufChar
                                    Case STX
                                        dtpToday = Now
                                        RcvBuffer = ""
                                        intPhase = 2
                                    Case Else
                                        RcvBuffer = RcvBuffer & BufChar
                                End Select
                            Case 2
                            
                                Select Case BufChar
                                    Case ETX
                                        Call SerialRcvData_UrinscanPro
                                        RcvBuffer = ""
                                        intPhase = 1
                                    Case Else
                                        RcvBuffer = RcvBuffer & BufChar
                                End Select
                        End Select
                    Next i
                
                Case "HITACHI7080"
                    Call Phase_Serial_HITACHI7080
                    
                Case "HORIBA"
                    lngBufLen = Len(pBuffer)
                
                    For i = 1 To lngBufLen
                        BufChar = Mid$(pBuffer, i, 1)
                        Select Case BufChar
                            Case STX
                                dtpToday = Now
                                RcvBuffer = ""
                                RcvBuffer = RcvBuffer & BufChar
                            Case ETX
                                Call SerialRcvData_HORIBA
                                RcvBuffer = ""
                            Case Else
                                RcvBuffer = RcvBuffer & BufChar
                        End Select
                    Next i
                
                Case "UROMETER720"
                    lngBufLen = Len(pBuffer)
                    
                    For i = 1 To lngBufLen
                        BufChar = Mid$(pBuffer, i, 1)
                        Select Case intPhase
                            Case 1
                                Select Case BufChar
                                    Case "~"
                                        RcvBuffer = ""
                                        intPhase = 2
                                    Case Else
                                        RcvBuffer = RcvBuffer & BufChar
                                End Select
                            Case 2
                            
                                Select Case BufChar
                                    Case "~"
                                        Call SerialRcvData_UROMETER720
                                        RcvBuffer = ""
                                        intPhase = 1
                                    Case Else
                                        RcvBuffer = RcvBuffer & BufChar
                                End Select
                        End Select
                    Next i
                    
            End Select
            
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
    
    'Me.Caption = gHOSP.MACHNM
    Me.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [장비와 인터페이스]     ◈◈◈◈◈"
    
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
        fraRS232.Visible = True
        
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
    Else
        If gComm.TCPTYPE = "1" Then
            wSck.LocalPort = CInt(gComm.TCPPORT)
            wSck.Listen

            lblStatus.Caption = "TCP " & gComm.TCPPORT & " 포트로 연결합니다"
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Visible = False
            imgReceive.Visible = False
            lblSend.Visible = False
            lblRcv.Visible = False
            'imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            'imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        Else
            wSck.Close
            wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)

            lblStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 포트로 연결합니다"
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Visible = False
            imgReceive.Visible = False
            lblSend.Visible = False
            lblRcv.Visible = False
            'imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            'imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        End If
    End If
    
    
    If gWORKTEST = "0" Then
        chkTest.Value = "0"
    Else
        chkTest.Value = "1"
    End If
    
    pDel = False
    
    If gHOSP.MACHNM = "ADVIA2120" Or gHOSP.MACHNM = "ADVIA2120-1" Or gHOSP.MACHNM = "ADVIA2120-2" Then
        cmdInit.Visible = True
        Call InitialComm
    Else
        cmdInit.Visible = False
    End If
        
    lblMenu(0).BackColor = &HFFFFC0
    Frame1.Visible = True
    Frame1.ZOrder 0

    
    '변수 초기화(Advia1650)
    iPendingFlag = 0: iTotQueryFlag = 0: iTmpPendingFlag = 0: iIdleFlag = 0
    iOrderFlag = 0: iResultFlag = 0
    sRcvState = "": sSndState = ""
    intPhase = 1
    
    miLineNo = 0
    
    'spdOrder.MaxRows = 10
    'spdOrder.RowHeight(-1) = 12
    
    Exit Sub
    
RST:
    Frame1.Visible = True
    Frame1.ZOrder 0
    
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
    Else
        optComType(1).Value = True
        frameCom.Enabled = False
        frameTCP.Enabled = True
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
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub
    
    '-- 인터페이스
    Frame1.Width = Me.ScaleWidth - 150
    Frame1.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150
    
    If gWORKPOS = "P" Then
        cmdWorkAll.Visible = False
        chkTest.Visible = False
        spdOrder.Width = Me.ScaleWidth - spdResult.Width - 400
        spdOrder.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500
        spdResult.Left = spdOrder.Left + spdOrder.Width + 50
        spdResult.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500
    Else
        cmdSL.Left = spdOrder.Left + 50
        chkAll.Left = spdOrder.Left + 550
        
        spdWork.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - cmdWorkAll.Height - 500
        spdOrder.Left = spdWork.Width + 100
        spdOrder.Width = Me.ScaleWidth - spdWork.Width - spdResult.Width - 400
        spdOrder.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500
        spdResult.Left = spdOrder.Left + spdOrder.Width + 50
        spdResult.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500
    End If
    

    fraDUREADER720.Left = spdResult.Left
    fraDUREADER720.Top = spdResult.Top + 7000
    
    '-- 결과조회
    Frame2.Width = Me.ScaleWidth - 150
    Frame2.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150
    
    spdROrder.Width = Me.ScaleWidth - spdRResult.Width - 500
    spdROrder.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500
    
    spdRResult.Left = spdROrder.Left + spdROrder.Width + 50
    spdRResult.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500
    
    '-- 검사설정
    frame3.Width = Me.ScaleWidth - 150
    frame3.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150
    
    spdTest.Width = Me.ScaleWidth - frameTestSet.Width - 600
    spdTest.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500
    
    frameTestSet.Left = spdTest.Left + spdTest.Width + 50
    frameTestSet.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500

    '-- 통신설정
    Frame4.Width = Me.ScaleWidth - 150
    Frame4.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150

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
    
    Frame1.Left = 50
    Frame1.Top = 1650
    
    Frame2.Left = 50
    Frame2.Top = 1650
    
    frame3.Left = 50
    frame3.Top = 1650
    
    Frame4.Left = 50
    Frame4.Top = 1650
    
    dtpToday.Value = Now
    dtpFrom.Value = Now
    dtpTo.Value = Now
    
    
    '-- 인터페이스
    spdWork.MaxRows = 0
    spdOrder.MaxRows = 0
    spdResult.MaxRows = 0
    
    If gWORKPOS = "P" Then
        spdWork.Visible = False
    Else
        spdWork.Visible = True
    End If
    txtSeqNo.Text = 1
    
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
            cboPort.AddItem Trim(Str(i))
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
    
    txtTV.Text = ""
    lblPatInfo(0).Caption = ""
    lblPatInfo(1).Caption = ""
    lblPatInfo(2).Caption = ""
    
    lblCommStatus.Caption = ""
    
    txtBarcode.Text = ""
    
    dtpFrDt.Value = Now
    dtpToDt.Value = Now
    
    '-- Urin Micro
    cboWbcM.Clear
    cboRbcM.Clear
    cboEpCell.Clear
    cboBacteria.Clear
    
    
    
    cboWbcM.AddItem "선택"
    cboWbcM.AddItem "0-1"
    cboWbcM.AddItem "1-4"
    cboWbcM.AddItem "5-10"
    cboWbcM.AddItem "10-20"
    cboWbcM.AddItem "30↑"
    cboWbcM.AddItem "few"
    cboWbcM.AddItem "some"
    cboWbcM.AddItem "many"
    
    cboWbcM.ListIndex = 0
    
    cboRbcM.AddItem "선택"
    cboRbcM.AddItem "0-1"
    cboRbcM.AddItem "1-4"
    cboRbcM.AddItem "5-10"
    cboRbcM.AddItem "10-20"
    cboRbcM.AddItem "30↑"
    cboRbcM.AddItem "few"
    cboRbcM.AddItem "some"
    cboRbcM.AddItem "many"
    
    cboRbcM.ListIndex = 0
    
    cboEpCell.AddItem "선택"
    cboEpCell.AddItem "0-1"
    cboEpCell.AddItem "1-4"
    cboEpCell.AddItem "5-10"
    cboEpCell.AddItem "10-20"
    cboEpCell.AddItem "30↑"
    cboEpCell.AddItem "few"
    cboEpCell.AddItem "some"
    cboEpCell.AddItem "many"
    
    cboEpCell.ListIndex = 0
    
    cboBacteria.AddItem "선택"
    cboBacteria.AddItem "0-1"
    cboBacteria.AddItem "1-4"
    cboBacteria.AddItem "5-10"
    cboBacteria.AddItem "10-20"
    cboBacteria.AddItem "30↑"
    cboBacteria.AddItem "few"
    cboBacteria.AddItem "some"
    cboBacteria.AddItem "many"
    
    cboBacteria.ListIndex = 0
    
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
        
'        If Trim(txtOChannel.Text) = "" Then
'            MsgBox "검사항목을 먼저 선택하세요", vbCritical, Me.Caption
'            Exit Sub
'        End If
        
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
        
'        If Trim(txtOChannel.Text) = "" Then
'            MsgBox "오더채널을 입력하세요", vbCritical, Me.Caption
'            txtOChannel.SetFocus
'            Exit Sub
'        End If
'
'        If Trim(txtRChannel.Text) = "" Then
'            MsgBox "결과채널을 입력하세요", vbCritical, Me.Caption
'            txtRChannel.SetFocus
'            Exit Sub
'        End If
        
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
            '-- QC
            .Add "LAB", txtLab.Text
            .Add "LOT", txtLot.Text
            .Add "ANALYTE", txtAnalyte.Text
            .Add "METHOD", txtMethod.Text
            .Add "INSTRUMENT", txtInstrument.Text
            .Add "REAGENT", txtReagent.Text
            .Add "UNIT", txtUnit.Text
            
            '.Add "TEMP", txtTemp.Text
            .Add "TEMP", chkResSpec.Value

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
    
    spdWork.MaxRows = 0
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
        Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Else
        Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    End If

    If optComType(0).Value = True Then
        Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Else
        Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    End If
    
    Call WritePrivateProfileString("COMM", "COMPORT", cboPort.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("COMM", "SPEED", cboBaudrate.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("COMM", "PARITY", cboParity.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("COMM", "DATABIT", cboDatabit.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("COMM", "STARTBIT", cboStartbit.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("COMM", "STOPBIT", cboStopbit.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    
    GetSetup
    
    GetCommList
    
    MsgBox "통신정보가 변경되었습니다.", vbInformation + vbOKCancel, Me.Caption

End Sub

Private Sub lblComSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblComSave.ForeColor = vbBlue
    shpCom.BorderColor = vbCyan

End Sub

Private Sub lblMenu_Click(Index As Integer)

    
    Frame1.Visible = False
    Frame2.Visible = False
    frame3.Visible = False
    'frame4.Visible = False
    fraInterface.Visible = False
    fraResult.Visible = False
    
    fraDUREADER720.Visible = False

    lblPatInfo(0).Caption = ""
    lblPatInfo(1).Caption = ""
    lblPatInfo(2).Caption = ""
    lblPatInfo(3).Caption = ""
    lblPatInfo(4).Caption = ""
    
'    lblMenu(0).FontUnderline = False
'    lblMenu(1).FontUnderline = False
'    lblMenu(2).FontUnderline = False
'    lblMenu(3).FontUnderline = False
'    lblMenu(Index).FontUnderline = True
    
    lblMenu(0).BackColor = vbWhite
    lblMenu(1).BackColor = vbWhite
    lblMenu(2).BackColor = vbWhite
    lblMenu(3).BackColor = vbWhite
    lblMenu(Index).BackColor = &HFFFFC0
    
    shpB(0).BorderColor = vbGreen
    shpB(1).BorderColor = vbGreen
    shpB(2).BorderColor = vbGreen
    shpB(3).BorderColor = vbGreen
    shpB(Index).BorderColor = vbBlue
    
     
    Select Case Index
        Case 0:
                Frame1.Visible = True
                Frame1.ZOrder 0
        
                fraInterface.Visible = True
                frmMain.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [장비와 인터페이스]     ◈◈◈◈◈"
                
                tmrComm.Enabled = False
                tmrFlipFlop.Enabled = False
                
                lblCommStatus.Caption = ""
        Case 1:
                Frame2.Visible = True
                Frame2.ZOrder 0
        
                fraResult.Visible = True
                frmMain.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [검사 결과 조회]     ◈◈◈◈◈"
        Case 2:
                frame3.Visible = True
                frame3.ZOrder 0
    
                '-- 검사코드
                Call GetTestList
                frmMain.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [검사 코드 설정]     ◈◈◈◈◈"
        
        Case 3:
                Frame4.Visible = True
                Frame4.ZOrder 0
    
                '-- 통신설정
                Call GetCommList
                frmMain.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [장비 통신 설정]     ◈◈◈◈◈"
    
    End Select
    
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
        Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Else
        Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    End If
    
    If optTCPType(0).Value = True Then
        Call WritePrivateProfileString("COMM", "TCPTYPE", "1", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Else
        Call WritePrivateProfileString("COMM", "TCPTYPE", "2", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    End If
    
    Call WritePrivateProfileString("COMM", "TCPIP", txtTCPIP.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("COMM", "TCPPORT", txtTCPPort.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    
    GetSetup
    
    GetCommList

    MsgBox "통신정보가 변경되었습니다.", vbInformation + vbOKCancel, Me.Caption

End Sub

Private Sub lblTcpSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblTcpSave.ForeColor = vbBlue
    shpTcp.BorderColor = vbCyan

End Sub

Private Sub lblWork_Click()
    
    If gWORKPOS = "P" Then
        frmWorkList.Show 'vbModal
    Else
        'Call GetWorkList_Main(Format(dtpFrDt.Value, "yyyymmdd"), Format(dtpToDt.Value, "yyyymmdd"), spdWork)
        Call GetWorkList(Format(dtpFrDt.Value, "yyyy-mm-dd"), Format(dtpToDt.Value, "yyyy-mm-dd"), spdWork)
    End If
    
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
    Dim intCol  As Integer
    Dim i       As Integer
    
    '-- 정렬
'    If Row = 0 Then
'        '-- 정렬 추가
'
'        Exit Sub
'    End If
    
    
    If Row = 0 And Col = colCHECKBOX Then
        If GetText(spdOrder, 1, colCHECKBOX) = "1" Then
            For i = 1 To spdOrder.DataRowCnt
                Call SetText(spdOrder, "0", i, colCHECKBOX)
            Next
        Else
            For i = 1 To spdOrder.DataRowCnt
                Call SetText(spdOrder, "1", i, colCHECKBOX)
            Next
        End If
        Exit Sub
    End If
    
    If Row > 0 And Col = colCHECKBOX Then
        If GetText(spdOrder, Row, colCHECKBOX) = "1" Then
            Call SetText(spdOrder, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdOrder, "1", Row, colCHECKBOX)
        End If
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
    
    If gHOSP.MACHNM = "DUREADER720" Then
        fraDUREADER720.Visible = True
        cboWbcM.SetFocus
        gRow = Row
    Else
        fraDUREADER720.Visible = False
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
    SQL = SQL & "SELECT DISTINCT SEQNO,EXAMCODE,EQUIPCODE,EXAMNAME,EQUIPRESULT,RESULT,REFJUDGE" & vbCr
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
                If AdoRs_Local.Fields("REFJUDGE").Value & "" <> "" Then
                    SetForeColor frmMain.spdRResult, intRow, intRow, colRLISRESULT, colRLISRESULT, 255, 0, 0
                End If
                
                If AdoRs_Local.Fields("EXAMCODE").Value & "" = "24HRS-V" Then
                    txtTV.Text = AdoRs_Local.Fields("RESULT").Value & ""
                End If
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
        Call SetSpreadSort(spdTest, 0)
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
        '-- QC
        txtLab = GetText(spdTest, Row, colLQCLab)
        txtLot = GetText(spdTest, Row, colLQCLot)
        txtAnalyte = GetText(spdTest, Row, colLQCAnalyte)
        txtMethod = GetText(spdTest, Row, colLQCMethod)
        txtInstrument = GetText(spdTest, Row, colLQCInstrument)
        txtReagent = GetText(spdTest, Row, colLQCReagent)
        txtUnit = GetText(spdTest, Row, colLQCUnit)
        txtTemp = GetText(spdTest, Row, colLQCTemp)
        
        If GetText(spdTest, Row, colLUseResSpec) = "1" Then
            chkResSpec.Value = "1"
        Else
            chkResSpec.Value = "0"
        End If
    
    End With
End Sub

Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Len(Trim(txtBarcode)) > 0 And KeyCode = vbKeyReturn Then
        mResult.BarNo = Trim(txtBarcode.Text)
        Call SetPatInfo_WithBar(mResult.BarNo, gHOSP.RSTTYPE)
        txtBarcode.SelStart = 0
        txtBarcode.SelLength = Len(txtBarcode.Text)
    End If
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
'                                 2. 장비수신정보 화면표시,
'                                 3. 처방코드 가져오기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Public Sub SetPatInfo_WithBar(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    With frmMain
        For i = 1 To .spdOrder.DataRowCnt
            If IsNumeric(pBarno) And IsNumeric(Trim(GetText(frmMain.spdOrder, i, colBARCODE))) Then
                If Val(Trim(GetText(frmMain.spdOrder, i, colBARCODE))) = Val(pBarno) Then
                    If Trim(GetText(frmMain.spdOrder, i, colSTATE)) = "" Then
                        intRow = i
                        Exit For
                    End If
                End If
            End If
        Next i
        
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        Call SetText(.spdOrder, "1", intRow, colCHECKBOX)
        
        '-- 장비결과인덱스 화면표시
        Call SetText(.spdOrder, mResult.RsltSeq, intRow, colSAVESEQ)
        Call SetText(.spdOrder, mResult.RsltDate, intRow, colEXAMDATE)
        
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mResult.BarNo, intRow, colBARCODE)
        'Call SetText(.spdOrder, mResult.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mResult.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mResult.TubePos, intRow, colPOSNO)
        'Call SetText(.spdOrder, Format(frmMain.txtSeqNo.Text, "#0"), intRow, colSEQNO)
        
        'SetText SPD, Format(frmMain.txtSeqNo.Text, "#0"), asRow, colSEQNO
    
        '-- 환자정보 표시
        'Call vasActiveCell(.spdOrder, intRow, colBARCODE)
        
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
    
    End With
    
    '-- 현재 Row
    gRow = intRow
    
End Sub





Private Sub txtSeqNo_KeyPress(KeyAscii As Integer)
    Dim strSeq  As String
    Dim i       As Integer
    
    If KeyAscii = vbKeyReturn Then
        DoEvents
        
        With spdOrder
            .Row = .ActiveRow
            .Col = .ActiveCol
            
            strSeq = Trim(txtSeqNo.Text)
            
            If IsNumeric(strSeq) Then
                For i = .ActiveRow To .DataRowCnt
                    Call SetText(spdOrder, Format(strSeq, "#0"), i, colSEQNO)
                    strSeq = Val(strSeq) + 1
                Next
            Else
                MsgBox strSeq & " : 숫자만 입력이 가능합니다", vbOKOnly + vbCritical, Me.Caption
            End If
        End With
    End If
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)

    Dim strTestNm   As String
    
    If KeyAscii = vbKeyReturn Then
        If Trim(txtTestCd.Text) <> "" Then
            strTestNm = GetTest(Trim(txtTestCd.Text))
            If strTestNm <> "" Then
                txtTestNm = strTestNm
                txtAbbrNm = strTestNm
            End If
        End If
    End If
End Sub

Private Sub txtTV_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call cmdTVSave_Click
    End If
    
End Sub

Private Sub wSck_ConnectionRequest(ByVal requestID As Long)
            
    If wSck.State <> sckClosed Then
        wSck.Close

        wSck.Accept requestID
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        If gComm.TCPTYPE = "1" Then
            lblStatus.Caption = "TCP " & gComm.TCPPORT & " 포트로 연결 되었습니다"
        Else
            lblStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 포트로 연결 되었습니다"
        End If
    End If
            
End Sub

Private Sub wSck_DataArrival(ByVal bytesTotal As Long)
    Dim strText     As String
    Dim varBuffers  As Variant
    
    
    wSck.GetData strText
    SetRawData "[Rx]" & strText
    
    '-- 컴파일시 제외할 것!!
    strText = Replace(strText, vbLf, "")
    
    varBuffers = Split(strText, vbCr)
    
    If UBound(varBuffers) > 0 Then
        strRecvData = varBuffers
    
        dtpToday.Value = Now
        
        Select Case UCase(gHOSP.MACHNM)
            Case "BS240PRO_HL7"
                Call TCPRcvData_BS240_HL7
    
            Case "BS240_ASTM"
                'Call TCPRcvData_BS240_ASTM
    
        End Select
    End If

End Sub

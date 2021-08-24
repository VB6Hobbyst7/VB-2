VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "OK SOFT"
   ClientHeight    =   11880
   ClientLeft      =   60
   ClientTop       =   -1830
   ClientWidth     =   21900
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
   ScaleHeight     =   11880
   ScaleWidth      =   21900
   StartUpPosition =   1  '소유자 가운데
   WindowState     =   2  '최대화
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
      Height          =   7125
      Left            =   18540
      TabIndex        =   62
      Top             =   3210
      Visible         =   0   'False
      Width           =   11265
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
         Height          =   405
         Left            =   8370
         TabIndex        =   232
         Top             =   570
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
            TabIndex        =   235
            Top             =   150
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
            TabIndex        =   234
            Top             =   150
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
            TabIndex        =   233
            Top             =   150
            Width           =   360
         End
         Begin VB.Image imgReceive 
            Height          =   240
            Left            =   2550
            Picture         =   "frmMain.frx":0E42
            Top             =   120
            Width           =   240
         End
         Begin VB.Image imgSend 
            Height          =   240
            Left            =   1635
            Picture         =   "frmMain.frx":13CC
            Top             =   120
            Width           =   240
         End
         Begin VB.Image imgPort 
            Height          =   240
            Left            =   690
            Picture         =   "frmMain.frx":1956
            Top             =   120
            Width           =   240
         End
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
         Left            =   540
         TabIndex        =   214
         Top             =   4680
         Visible         =   0   'False
         Width           =   5445
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
            TabIndex        =   224
            Top             =   630
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
            TabIndex        =   223
            Top             =   300
            Visible         =   0   'False
            Width           =   195
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
            TabIndex        =   222
            Top             =   300
            Visible         =   0   'False
            Width           =   195
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
            TabIndex        =   221
            Top             =   1380
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
            TabIndex        =   220
            Top             =   300
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
            TabIndex        =   219
            Top             =   660
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
            TabIndex        =   218
            Top             =   1020
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
            TabIndex        =   217
            Top             =   1380
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
            TabIndex        =   216
            Top             =   300
            Width           =   1185
         End
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
            TabIndex        =   215
            Top             =   1020
            Width           =   1185
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
            TabIndex        =   231
            Top             =   1470
            Width           =   360
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   30
            Left            =   2670
            Picture         =   "frmMain.frx":1EE0
            Top             =   1440
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   29
            Left            =   2670
            Picture         =   "frmMain.frx":22CA
            Top             =   360
            Width           =   150
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
            TabIndex        =   230
            Top             =   390
            Width           =   255
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
            TabIndex        =   229
            Top             =   750
            Width           =   630
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
            TabIndex        =   228
            Top             =   1110
            Width           =   360
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   28
            Left            =   2670
            Picture         =   "frmMain.frx":26B4
            Top             =   720
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   27
            Left            =   2670
            Picture         =   "frmMain.frx":2A9E
            Top             =   1080
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
            TabIndex        =   227
            Top             =   1470
            Width           =   360
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   26
            Left            =   210
            Picture         =   "frmMain.frx":2E88
            Top             =   1440
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   25
            Left            =   210
            Picture         =   "frmMain.frx":3272
            Top             =   360
            Width           =   150
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
            TabIndex        =   226
            Top             =   390
            Width           =   315
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
            TabIndex        =   225
            Top             =   1110
            Width           =   360
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   17
            Left            =   210
            Picture         =   "frmMain.frx":365C
            Top             =   1080
            Width           =   150
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Height          =   2415
         Left            =   5250
         TabIndex        =   197
         Top             =   1890
         Visible         =   0   'False
         Width           =   5565
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
            ItemData        =   "frmMain.frx":3A46
            Left            =   1770
            List            =   "frmMain.frx":3A48
            TabIndex        =   208
            Top             =   1860
            Visible         =   0   'False
            Width           =   2175
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
            Height          =   375
            Left            =   1590
            TabIndex        =   205
            Top             =   210
            Width           =   2565
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
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   207
               Top             =   210
               Value           =   -1  'True
               Width           =   1125
            End
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
               Height          =   180
               Index           =   1
               Left            =   1320
               TabIndex        =   206
               Top             =   210
               Width           =   1125
            End
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
            ItemData        =   "frmMain.frx":3A4A
            Left            =   2790
            List            =   "frmMain.frx":3A4C
            TabIndex        =   204
            Top             =   630
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
            Left            =   1590
            TabIndex        =   203
            Top             =   630
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
            Left            =   3540
            TabIndex        =   202
            Top             =   630
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
            Left            =   3540
            TabIndex        =   201
            Top             =   990
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
            ItemData        =   "frmMain.frx":3A4E
            Left            =   2790
            List            =   "frmMain.frx":3A50
            TabIndex        =   200
            Top             =   1350
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
            Left            =   1590
            TabIndex        =   199
            Top             =   1350
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
            Left            =   3540
            TabIndex        =   198
            Top             =   1350
            Width           =   1545
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   14
            Left            =   450
            Picture         =   "frmMain.frx":3A52
            Top             =   1920
            Visible         =   0   'False
            Width           =   150
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
            Left            =   720
            TabIndex        =   213
            Top             =   1950
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   1
            Left            =   240
            Picture         =   "frmMain.frx":3E3C
            Top             =   300
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
            Left            =   510
            TabIndex        =   212
            Top             =   330
            Width           =   510
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   12
            Left            =   240
            Picture         =   "frmMain.frx":4226
            Top             =   690
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
            Left            =   510
            TabIndex        =   211
            Top             =   720
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
            Left            =   510
            TabIndex        =   210
            Top             =   1080
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
            Left            =   510
            TabIndex        =   209
            Top             =   1440
            Width           =   840
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   9
            Left            =   240
            Picture         =   "frmMain.frx":4610
            Top             =   1050
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   13
            Left            =   240
            Picture         =   "frmMain.frx":49FA
            Top             =   1410
            Width           =   150
         End
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
         Left            =   4530
         TabIndex        =   113
         Text            =   "60"
         Top             =   210
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Timer Timer1 
         Left            =   2070
         Top             =   270
      End
      Begin VB.Timer tmrFlipFlop 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2910
         Top             =   270
      End
      Begin VB.Timer tmrComm 
         Enabled         =   0   'False
         Left            =   2490
         Top             =   270
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
         TabIndex        =   81
         Top             =   1200
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
            TabIndex        =   111
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
            TabIndex        =   110
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
            TabIndex        =   83
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
            TabIndex        =   82
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
         TabIndex        =   76
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
            TabIndex        =   78
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
            TabIndex        =   77
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
         TabIndex        =   73
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
            TabIndex        =   75
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
            TabIndex        =   74
            Top             =   30
            Width           =   765
         End
      End
      Begin VB.Timer tmrReceive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1230
         Top             =   270
      End
      Begin VB.Timer tmrSend 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1650
         Top             =   270
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   3450
         Top             =   210
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4DE4
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":537E
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5918
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5EB2
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6744
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":689E
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":69F8
               Key             =   "NOF"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6B52
               Key             =   "ON"
               Object.Tag             =   "연결성공"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6E2D
               Key             =   "OFF"
               Object.Tag             =   "연결실패"
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   885
         Left            =   300
         TabIndex        =   94
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
         SpreadDesigner  =   "frmMain.frx":7107
      End
      Begin FPSpread.vaSpread spdQcResult 
         Height          =   825
         Left            =   300
         TabIndex        =   100
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
         SpreadDesigner  =   "frmMain.frx":734A
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5610
         Top             =   210
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imlMenu 
         Left            =   3510
         Top             =   690
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":758D
               Key             =   "INTERFACE"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7E67
               Key             =   "SEARCH"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8181
               Key             =   "TESTSET"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8A39
               Key             =   "COMMSET"
            EndProperty
         EndProperty
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
         TabIndex        =   84
         Top             =   1290
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
         TabIndex        =   80
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
         TabIndex        =   79
         Top             =   1710
         Width           =   780
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '아래 맞춤
      Height          =   405
      Left            =   0
      TabIndex        =   163
      Top             =   11475
      Width           =   21900
      _ExtentX        =   38629
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   17754
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17754
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      ScaleWidth      =   21900
      TabIndex        =   11
      Top             =   0
      Width           =   21900
      Begin VB.Frame Frame18 
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
         Height          =   405
         Left            =   12630
         TabIndex        =   192
         Top             =   540
         Width           =   1485
         Begin VB.Image imgConn 
            Height          =   255
            Left            =   990
            Top             =   120
            Width           =   345
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "장비연결"
            ForeColor       =   &H00C000C0&
            Height          =   180
            Left            =   120
            TabIndex        =   193
            Top             =   150
            Width           =   780
         End
      End
      Begin VB.Frame Frame13 
         Height          =   675
         Left            =   7050
         TabIndex        =   173
         Top             =   120
         Visible         =   0   'False
         Width           =   1515
         Begin MSCommLib.MSComm comEqp 
            Left            =   180
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
            RThreshold      =   1
            RTSEnable       =   -1  'True
            EOFEnable       =   -1  'True
         End
      End
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
         Left            =   15690
         TabIndex        =   87
         Top             =   0
         Visible         =   0   'False
         Width           =   5175
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
            TabIndex        =   112
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
            TabIndex        =   89
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
            TabIndex        =   88
            Top             =   120
            Width           =   4425
         End
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   9810
         TabIndex        =   85
         Top             =   450
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
         CalendarBackColor=   16777215
         Format          =   139264000
         CurrentDate     =   40457
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Com1 포트에 연결되었습니다"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   12720
         TabIndex        =   194
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lblCommStatus 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Com"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   12720
         TabIndex        =   101
         Top             =   120
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
         Left            =   9000
         TabIndex        =   86
         Top             =   540
         Width           =   720
      End
      Begin VB.Image Image7 
         Height          =   225
         Left            =   8760
         Picture         =   "frmMain.frx":9313
         Top             =   540
         Width           =   150
      End
      Begin VB.Label lblHospInfo 
         BackStyle       =   0  '투명
         Caption         =   "전남대학교병원 HITACHI 7020[H36] 홍길동[12345]"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1740
         TabIndex        =   12
         Top             =   510
         Width           =   7005
      End
      Begin VB.Image Image3 
         Height          =   1065
         Left            =   0
         Picture         =   "frmMain.frx":96FD
         Top             =   0
         Width           =   12900
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
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
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   21870
      TabIndex        =   13
      Top             =   1035
      Width           =   21900
      Begin VB.Frame Frame17 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   6450
         TabIndex        =   180
         Top             =   -60
         Width           =   2175
         Begin VB.Shape shpB 
            BackColor       =   &H00C0FFC0&
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   435
            Index           =   3
            Left            =   690
            Shape           =   4  '둥근 사각형
            Top             =   210
            Width           =   1395
         End
         Begin VB.Label lblMenu 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "통신설정 "
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
            Height          =   315
            Index           =   3
            Left            =   750
            TabIndex        =   182
            ToolTipText     =   "통신 및 인터페이스를 설정힙니다."
            Top             =   270
            Width           =   1275
         End
         Begin VB.Image Image13 
            Height          =   480
            Left            =   120
            Picture         =   "frmMain.frx":AE40
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Frame Frame16 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   4290
         TabIndex        =   179
         Top             =   -60
         Width           =   2175
         Begin VB.Label lblMenu 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "검사설정 "
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
            Height          =   315
            Index           =   2
            Left            =   750
            TabIndex        =   181
            ToolTipText     =   "검사항목을 설정합니다."
            Top             =   270
            Width           =   1275
         End
         Begin VB.Shape shpB 
            BackColor       =   &H00C0FFC0&
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   435
            Index           =   2
            Left            =   690
            Shape           =   4  '둥근 사각형
            Top             =   210
            Width           =   1395
         End
         Begin VB.Image Image12 
            Height          =   465
            Left            =   120
            Picture         =   "frmMain.frx":B70A
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   2130
         TabIndex        =   176
         Top             =   -60
         Width           =   2175
         Begin VB.Label lblMenu 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "결과조회 "
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
            Height          =   315
            Index           =   1
            Left            =   750
            TabIndex        =   177
            ToolTipText     =   "인터페이스된 검사결과를 조회합니다."
            Top             =   270
            Width           =   1275
         End
         Begin VB.Shape shpB 
            BackColor       =   &H00C0FFC0&
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   435
            Index           =   1
            Left            =   690
            Shape           =   4  '둥근 사각형
            Top             =   210
            Width           =   1395
         End
         Begin VB.Image Image11 
            Height          =   480
            Left            =   120
            Picture         =   "frmMain.frx":BFB2
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Frame Frame14 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   -30
         TabIndex        =   174
         Top             =   -60
         Width           =   2175
         Begin VB.Image Image10 
            Height          =   480
            Left            =   120
            Picture         =   "frmMain.frx":C2BC
            Top             =   180
            Width           =   480
         End
         Begin VB.Shape shpB 
            BackColor       =   &H00C0FFC0&
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   435
            Index           =   0
            Left            =   690
            Shape           =   4  '둥근 사각형
            Top             =   210
            Width           =   1425
         End
         Begin VB.Label lblMenu 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            Caption         =   "인터페이스 "
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
            Height          =   315
            Index           =   0
            Left            =   720
            TabIndex        =   175
            ToolTipText     =   "검사장비와 인터페이스를 합니다."
            Top             =   270
            Width           =   1335
         End
      End
      Begin VB.Frame fraInterface 
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
         Height          =   585
         Left            =   8760
         TabIndex        =   60
         Top             =   30
         Width           =   12405
         Begin VB.TextBox txtSeqNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   15.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   7770
            TabIndex        =   236
            Text            =   "1"
            Top             =   90
            Width           =   765
         End
         Begin VB.Frame Frame10 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BorderStyle     =   0  '없음
            ForeColor       =   &H00FFFFFF&
            Height          =   585
            Left            =   3270
            TabIndex        =   106
            Top             =   0
            Width           =   4425
            Begin VB.Shape shpW 
               BackColor       =   &H00808080&
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               FillColor       =   &H00C0FFC0&
               Height          =   375
               Left            =   90
               Shape           =   4  '둥근 사각형
               Top             =   150
               Width           =   1365
            End
            Begin VB.Label lblWork 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H80000005&
               BackStyle       =   0  '투명
               Caption         =   "워크조회"
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   210
               TabIndex        =   109
               Top             =   240
               Width           =   1125
            End
            Begin VB.Label lblSave 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H80000005&
               BackStyle       =   0  '투명
               Caption         =   "선택저장"
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   1680
               TabIndex        =   108
               Top             =   240
               Width           =   1125
            End
            Begin VB.Shape shpS 
               BackColor       =   &H00808080&
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               FillColor       =   &H00C0FFC0&
               Height          =   375
               Left            =   1560
               Shape           =   4  '둥근 사각형
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
               Height          =   225
               Left            =   3120
               TabIndex        =   107
               Top             =   240
               Width           =   1125
            End
            Begin VB.Shape shpC 
               BackColor       =   &H00808080&
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               FillColor       =   &H00C0FFC0&
               Height          =   375
               Left            =   3000
               Shape           =   4  '둥근 사각형
               Top             =   150
               Width           =   1365
            End
         End
         Begin VB.TextBox txtBarcode 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10470
            TabIndex        =   102
            Text            =   "1234567890"
            Top             =   180
            Visible         =   0   'False
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker dtpFrDt 
            Height          =   345
            Left            =   60
            TabIndex        =   103
            Top             =   180
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   139264001
            CurrentDate     =   40457
         End
         Begin MSComCtl2.DTPicker dtpToDt 
            Height          =   345
            Left            =   1680
            TabIndex        =   104
            Top             =   180
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   139264001
            CurrentDate     =   40457
         End
         Begin VB.Label Label4 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "바코드"
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
            Left            =   9810
            TabIndex        =   165
            Top             =   240
            Visible         =   0   'False
            Width           =   540
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
            TabIndex        =   105
            Top             =   240
            Width           =   150
         End
      End
      Begin VB.Frame fraSet 
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
         Height          =   585
         Left            =   8760
         TabIndex        =   166
         Top             =   30
         Visible         =   0   'False
         Width           =   8505
         Begin VB.CommandButton cmdTestOptSet 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "검사옵션 설정"
            Height          =   435
            Left            =   6630
            Style           =   1  '그래픽
            TabIndex        =   170
            Top             =   90
            Width           =   1665
         End
         Begin VB.CommandButton Command6 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "화면 설정"
            Height          =   435
            Left            =   4560
            Style           =   1  '그래픽
            TabIndex        =   169
            Top             =   90
            Width           =   1665
         End
         Begin VB.CommandButton cmdEMRSet 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "EMR 설정"
            Height          =   435
            Left            =   420
            MaskColor       =   &H00FFFFFF&
            Style           =   1  '그래픽
            TabIndex        =   168
            Top             =   90
            Width           =   1665
         End
         Begin VB.CommandButton cmdConfig 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "병원정보 설정"
            Height          =   435
            Left            =   2490
            Style           =   1  '그래픽
            TabIndex        =   167
            Top             =   90
            Width           =   1665
         End
         Begin VB.Image Image9 
            Height          =   225
            Left            =   6390
            Picture         =   "frmMain.frx":CB86
            Top             =   210
            Width           =   150
         End
         Begin VB.Image Image8 
            Height          =   225
            Left            =   4320
            Picture         =   "frmMain.frx":CF70
            Top             =   210
            Width           =   150
         End
         Begin VB.Image Image6 
            Height          =   225
            Left            =   2250
            Picture         =   "frmMain.frx":D35A
            Top             =   210
            Width           =   150
         End
         Begin VB.Image Image2 
            Height          =   225
            Left            =   180
            Picture         =   "frmMain.frx":D744
            Top             =   210
            Width           =   150
         End
      End
      Begin VB.Frame fraResult 
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
         Height          =   585
         Left            =   8760
         TabIndex        =   68
         Top             =   30
         Visible         =   0   'False
         Width           =   12525
         Begin VB.CommandButton cmdWork 
            BackColor       =   &H00C0FFFF&
            Caption         =   "통계조회"
            Height          =   405
            Left            =   9000
            Style           =   1  '그래픽
            TabIndex        =   195
            Top             =   120
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.ComboBox cboRstType 
            Appearance      =   0  '평면
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
            ItemData        =   "frmMain.frx":DB2E
            Left            =   420
            List            =   "frmMain.frx":DB30
            TabIndex        =   91
            Top             =   180
            Width           =   1245
         End
         Begin VB.ComboBox cboState 
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
            ItemData        =   "frmMain.frx":DB32
            Left            =   4890
            List            =   "frmMain.frx":DB34
            TabIndex        =   90
            Top             =   180
            Width           =   1245
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   315
            Left            =   1710
            TabIndex        =   70
            Top             =   180
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   139264001
            CurrentDate     =   40457
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   315
            Left            =   3360
            TabIndex        =   71
            Top             =   180
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   139264001
            CurrentDate     =   40457
         End
         Begin VB.Shape shpRS 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   11910
            Shape           =   4  '둥근 사각형
            Top             =   120
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label lblRSave 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "선택저장"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   12030
            TabIndex        =   164
            Top             =   210
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Shape shpRX 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   10500
            Shape           =   4  '둥근 사각형
            Top             =   150
            Width           =   1365
         End
         Begin VB.Label lblRExcel 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "엑셀출력"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   10620
            TabIndex        =   114
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label lblRClear 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "화면정리"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7710
            TabIndex        =   99
            Top             =   240
            Width           =   1125
         End
         Begin VB.Shape shpRC 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   7590
            Shape           =   4  '둥근 사각형
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
            Left            =   3180
            TabIndex        =   72
            Top             =   240
            Width           =   150
         End
         Begin VB.Image imgGbn 
            Height          =   225
            Left            =   180
            Picture         =   "frmMain.frx":DB36
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
            Shape           =   4  '둥근 사각형
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
            TabIndex        =   69
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
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
         Index           =   4
         Left            =   60
         TabIndex        =   178
         Top             =   60
         Width           =   1275
      End
      Begin VB.Shape shpB 
         BackColor       =   &H00C0FFC0&
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00C0FFC0&
         Height          =   345
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   1395
      End
   End
   Begin VB.Frame frame2 
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
      Left            =   0
      TabIndex        =   63
      Top             =   1800
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
         TabIndex        =   93
         Top             =   210
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CheckBox chkRAll 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   92
         Top             =   240
         Width           =   195
      End
      Begin FPSpread.vaSpread spdRResult 
         Height          =   9360
         Left            =   13620
         TabIndex        =   67
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
         GrayAreaBackColor=   16777215
         MaxCols         =   12
         MaxRows         =   50
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16777215
         SpreadDesigner  =   "frmMain.frx":DF20
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdROrder 
         Height          =   9375
         Left            =   60
         TabIndex        =   66
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
         GrayAreaBackColor=   16777215
         MaxCols         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16777215
         SpreadDesigner  =   "frmMain.frx":E66C
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
         TabIndex        =   65
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
         TabIndex        =   64
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
      Left            =   0
      TabIndex        =   16
      Top             =   1800
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
         TabIndex        =   17
         Top             =   180
         Width           =   5835
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
            TabIndex        =   196
            Top             =   870
            Width           =   1245
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
            Left            =   2340
            TabIndex        =   186
            Top             =   4440
            Width           =   855
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
            Left            =   2340
            TabIndex        =   185
            Top             =   4020
            Width           =   855
         End
         Begin VB.TextBox txtRefLowF 
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
            Left            =   4080
            TabIndex        =   184
            Top             =   4020
            Width           =   855
         End
         Begin VB.TextBox txtRefHighF 
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
            Left            =   4080
            TabIndex        =   183
            Top             =   4440
            Width           =   855
         End
         Begin VB.CheckBox chkResSpec 
            BackColor       =   &H00FFFFFF&
            Caption         =   "사용"
            Height          =   390
            Left            =   1650
            TabIndex        =   7
            Top             =   3540
            Width           =   705
         End
         Begin VB.CommandButton cmdQCMaster 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "Biorad QC 설정"
            Height          =   345
            Left            =   4140
            Style           =   1  '그래픽
            TabIndex        =   98
            Top             =   7560
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
            Left            =   1920
            TabIndex        =   96
            Top             =   7590
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
            TabIndex        =   57
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
               TabIndex        =   61
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
               TabIndex        =   58
               Top             =   210
               Width           =   285
            End
            Begin FPSpread.vaSpread spdOrdMst 
               Height          =   1920
               Left            =   90
               TabIndex        =   59
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
               SpreadDesigner  =   "frmMain.frx":11669
               TextTip         =   2
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
            TabIndex        =   3
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
            TabIndex        =   18
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
            TabIndex        =   4
            Top             =   2220
            Width           =   2115
         End
         Begin VB.TextBox txtTestNm 
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
            TabIndex        =   5
            Top             =   2670
            Width           =   3255
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
            TabIndex        =   2
            Top             =   1320
            Width           =   2115
         End
         Begin VB.TextBox txtAbbrNm 
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
            TabIndex        =   6
            Top             =   3120
            Width           =   3255
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
            Left            =   2400
            TabIndex        =   8
            Top             =   3540
            Width           =   435
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
            TabIndex        =   1
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
            TabIndex        =   0
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
            Top             =   3510
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
            Top             =   3510
            Width           =   435
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "남성하한"
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
            Left            =   1500
            TabIndex        =   190
            Top             =   4110
            Width           =   720
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "남성상한"
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
            Left            =   1500
            TabIndex        =   189
            Top             =   4500
            Width           =   720
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "여성상한"
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
            Index           =   49
            Left            =   3300
            TabIndex        =   188
            Top             =   4500
            Width           =   720
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "여성하한"
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
            Index           =   50
            Left            =   3300
            TabIndex        =   187
            Top             =   4080
            Width           =   720
         End
         Begin VB.Shape shpA 
            BackColor       =   &H00808080&
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   1395
            Index           =   0
            Left            =   930
            Shape           =   3  '원형
            Top             =   4980
            Width           =   1695
         End
         Begin VB.Shape shpA 
            BackColor       =   &H00808080&
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   1395
            Index           =   1
            Left            =   2430
            Shape           =   3  '원형
            Top             =   4980
            Width           =   1845
         End
         Begin VB.Image imgTestSet 
            Height          =   1260
            Index           =   1
            Left            =   2700
            Picture         =   "frmMain.frx":1198C
            Top             =   5040
            Width           =   1290
         End
         Begin VB.Image imgTestSet 
            Height          =   1260
            Index           =   0
            Left            =   1140
            Picture         =   "frmMain.frx":137A6
            Top             =   5040
            Width           =   1290
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   24
            Left            =   600
            Picture         =   "frmMain.frx":154EF
            Top             =   7620
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
            Left            =   870
            TabIndex        =   97
            Top             =   7650
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
            TabIndex        =   56
            Top             =   933
            Width           =   360
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
            TabIndex        =   55
            Top             =   1839
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   11
            Left            =   330
            Picture         =   "frmMain.frx":158D9
            Top             =   1809
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   0
            Left            =   330
            Picture         =   "frmMain.frx":15CC3
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
            TabIndex        =   54
            Top             =   480
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   2
            Left            =   330
            Picture         =   "frmMain.frx":160AD
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
            TabIndex        =   53
            Top             =   1386
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   3
            Left            =   330
            Picture         =   "frmMain.frx":16497
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
            TabIndex        =   52
            Top             =   2292
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   4
            Left            =   330
            Picture         =   "frmMain.frx":16881
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
            TabIndex        =   51
            Top             =   2745
            Width           =   540
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   5
            Left            =   330
            Picture         =   "frmMain.frx":16C6B
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
            TabIndex        =   50
            Top             =   3198
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   6
            Left            =   330
            Picture         =   "frmMain.frx":17055
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
            TabIndex        =   49
            Top             =   3651
            Width           =   540
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   8
            Left            =   330
            Picture         =   "frmMain.frx":1743F
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
            TabIndex        =   48
            Top             =   4104
            Width           =   540
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   16
            Left            =   330
            Picture         =   "frmMain.frx":17829
            Top             =   903
            Width           =   150
         End
         Begin VB.Shape shpA 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   405
            Index           =   2
            Left            =   3990
            Top             =   8550
            Visible         =   0   'False
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
            TabIndex        =   47
            Top             =   8640
            Visible         =   0   'False
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
            TabIndex        =   20
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
            Index           =   3
            Left            =   2580
            Top             =   8550
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin FPSpread.vaSpread spdTest 
         Height          =   9195
         Left            =   210
         TabIndex        =   191
         Top             =   270
         Width           =   14325
         _Version        =   393216
         _ExtentX        =   25268
         _ExtentY        =   16219
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
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
         GrayAreaBackColor=   16777215
         MaxCols         =   30
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16777215
         SpreadDesigner  =   "frmMain.frx":17C13
      End
   End
   Begin VB.Frame frame1 
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
      Left            =   15
      TabIndex        =   14
      Top             =   1800
      Width           =   20685
      Begin VB.CommandButton cmdWorkSearch 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   990
         Style           =   1  '그래픽
         TabIndex        =   172
         Top             =   210
         Width           =   1305
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00FFFFFF&
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
         Left            =   720
         TabIndex        =   171
         Top             =   240
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CheckBox chkTest 
         BackColor       =   &H00FFFFFF&
         Caption         =   "선검사"
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
         Left            =   120
         TabIndex        =   118
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdWorkAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "워크등록  ▷▷ "
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
         Left            =   2280
         Style           =   1  '그래픽
         TabIndex        =   117
         Top             =   210
         Width           =   1725
      End
      Begin FPSpread.vaSpread spdWork 
         Height          =   9375
         Left            =   60
         TabIndex        =   116
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
         GrayAreaBackColor=   16777215
         MaxCols         =   21
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16777215
         SpreadDesigner  =   "frmMain.frx":18A53
         UserResize      =   2
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
         TabIndex        =   19
         Top             =   210
         Visible         =   0   'False
         Width           =   435
      End
      Begin FPSpread.vaSpread spdResult 
         Height          =   9360
         Left            =   12870
         TabIndex        =   15
         Top             =   180
         Width           =   7710
         _Version        =   393216
         _ExtentX        =   13600
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
         GrayAreaBackColor=   16777215
         MaxCols         =   12
         MaxRows         =   50
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16777215
         SpreadDesigner  =   "frmMain.frx":1BB14
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdOrder 
         Height          =   9375
         Left            =   60
         TabIndex        =   115
         Top             =   180
         Width           =   12765
         _Version        =   393216
         _ExtentX        =   22516
         _ExtentY        =   16536
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         ColHeaderDisplay=   0
         ColsFrozen      =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   16777215
         MaxCols         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16777215
         ShadowDark      =   16777215
         SpreadDesigner  =   "frmMain.frx":1C27B
         UserResize      =   2
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
      Left            =   0
      TabIndex        =   21
      Top             =   1800
      Visible         =   0   'False
      Width           =   20685
      Begin VB.Frame fraView 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 화면 설정 "
         Height          =   7935
         Left            =   11460
         TabIndex        =   125
         Top             =   900
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   15
            Left            =   1620
            TabIndex        =   162
            Top             =   7140
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   14
            Left            =   1620
            TabIndex        =   161
            Top             =   6840
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   15
            Left            =   420
            TabIndex        =   160
            Top             =   7140
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   14
            Left            =   420
            TabIndex        =   159
            Top             =   6900
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   17
            Left            =   2790
            TabIndex        =   158
            Top             =   6240
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   16
            Left            =   2790
            TabIndex        =   157
            Top             =   5894
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   13
            Left            =   2790
            TabIndex        =   156
            Top             =   5518
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   12
            Left            =   2790
            TabIndex        =   155
            Top             =   5142
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   11
            Left            =   2790
            TabIndex        =   154
            Top             =   4766
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   10
            Left            =   2790
            TabIndex        =   153
            Top             =   4390
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   9
            Left            =   2790
            TabIndex        =   152
            Top             =   4014
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   8
            Left            =   2790
            TabIndex        =   151
            Top             =   3638
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   7
            Left            =   2790
            TabIndex        =   150
            Top             =   3262
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   6
            Left            =   2790
            TabIndex        =   149
            Top             =   2886
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   5
            Left            =   2790
            TabIndex        =   148
            Top             =   2510
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   4
            Left            =   2790
            TabIndex        =   147
            Top             =   2134
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   3
            Left            =   2790
            TabIndex        =   146
            Top             =   1758
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   2
            Left            =   2790
            TabIndex        =   145
            Top             =   1382
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   1
            Left            =   2790
            TabIndex        =   144
            Top             =   1006
            Width           =   1515
         End
         Begin VB.TextBox txtColumn 
            Height          =   315
            Index           =   0
            Left            =   2790
            TabIndex        =   143
            Top             =   630
            Width           =   1515
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   17
            Left            =   540
            TabIndex        =   141
            Top             =   6330
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   16
            Left            =   540
            TabIndex        =   140
            Top             =   5954
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   13
            Left            =   540
            TabIndex        =   139
            Top             =   5578
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   12
            Left            =   540
            TabIndex        =   138
            Top             =   5202
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   11
            Left            =   540
            TabIndex        =   137
            Top             =   4826
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   10
            Left            =   540
            TabIndex        =   136
            Top             =   4450
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   9
            Left            =   540
            TabIndex        =   135
            Top             =   4074
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   8
            Left            =   540
            TabIndex        =   134
            Top             =   3698
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   7
            Left            =   540
            TabIndex        =   133
            Top             =   3322
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   6
            Left            =   540
            TabIndex        =   132
            Top             =   2946
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   5
            Left            =   540
            TabIndex        =   131
            Top             =   2570
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   4
            Left            =   540
            TabIndex        =   130
            Top             =   2194
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   3
            Left            =   540
            TabIndex        =   129
            Top             =   1818
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   2
            Left            =   540
            TabIndex        =   128
            Top             =   1442
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   1
            Left            =   540
            TabIndex        =   127
            Top             =   1066
            Width           =   1995
         End
         Begin VB.CheckBox chkColumn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "저장순번"
            Height          =   180
            Index           =   0
            Left            =   540
            TabIndex        =   126
            Top             =   690
            Width           =   1995
         End
         Begin VB.Label lblView 
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
            Left            =   3240
            TabIndex        =   142
            Top             =   7110
            Width           =   1125
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   585
            Left            =   3120
            Top             =   6930
            Width           =   1365
         End
      End
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
         Left            =   17430
         TabIndex        =   95
         Top             =   8280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton optComType 
         BackColor       =   &H00808080&
         Caption         =   "TCP-IP 사용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   1
         Left            =   5910
         TabIndex        =   43
         Top             =   360
         Width           =   5325
      End
      Begin VB.OptionButton optComType 
         BackColor       =   &H00808080&
         Caption         =   "RS-2232 사용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   0
         Left            =   450
         TabIndex        =   42
         Top             =   360
         Value           =   -1  'True
         Width           =   5295
      End
      Begin VB.Frame frameTCP 
         BackColor       =   &H00FFFFFF&
         Caption         =   " TCP-IP 상세설정 "
         ForeColor       =   &H00808080&
         Height          =   7935
         Left            =   5910
         TabIndex        =   36
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   41
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
            TabIndex        =   40
            Top             =   930
            Width           =   2445
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   7
            Left            =   840
            Picture         =   "frmMain.frx":1F1AC
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
            TabIndex        =   44
            Top             =   480
            Width           =   465
         End
         Begin VB.Shape shpTcp 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   4
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
               Size            =   14.25
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
            TabIndex        =   38
            Top             =   1395
            Width           =   375
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   15
            Left            =   840
            Picture         =   "frmMain.frx":1F596
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
            TabIndex        =   37
            Top             =   990
            Width           =   180
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   10
            Left            =   840
            Picture         =   "frmMain.frx":1F980
            Top             =   960
            Width           =   150
         End
      End
      Begin VB.Frame frameCom 
         BackColor       =   &H00FFFFFF&
         Caption         =   " RS-232 상세설정 "
         ForeColor       =   &H00808080&
         Height          =   7935
         Left            =   420
         TabIndex        =   22
         Top             =   870
         Width           =   5325
         Begin VB.CheckBox chkDTR 
            BackColor       =   &H00FFFFFF&
            Caption         =   "True"
            Height          =   315
            Left            =   2610
            TabIndex        =   124
            Top             =   3510
            Value           =   1  '확인
            Width           =   1785
         End
         Begin VB.CheckBox chkRTS 
            BackColor       =   &H00FFFFFF&
            Caption         =   "True"
            Height          =   315
            Left            =   2610
            TabIndex        =   119
            Top             =   3000
            Value           =   1  '확인
            Width           =   1785
         End
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
            ItemData        =   "frmMain.frx":1FD6A
            Left            =   2190
            List            =   "frmMain.frx":1FD6C
            TabIndex        =   35
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
            ItemData        =   "frmMain.frx":1FD6E
            Left            =   2190
            List            =   "frmMain.frx":1FD70
            TabIndex        =   34
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
            ItemData        =   "frmMain.frx":1FD72
            Left            =   2190
            List            =   "frmMain.frx":1FD74
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            ItemData        =   "frmMain.frx":1FD76
            Left            =   2190
            List            =   "frmMain.frx":1FD78
            TabIndex        =   30
            Top             =   2520
            Width           =   2205
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "DTREnable"
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
            Index           =   48
            Left            =   1110
            TabIndex        =   123
            Top             =   3570
            Width           =   1035
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   35
            Left            =   840
            Picture         =   "frmMain.frx":1FD7A
            Top             =   3540
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "RTSEnable"
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
            Left            =   1110
            TabIndex        =   120
            Top             =   3060
            Width           =   1035
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   32
            Left            =   840
            Picture         =   "frmMain.frx":20164
            Top             =   3030
            Width           =   150
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
            TabIndex        =   29
            Top             =   1290
            Width           =   645
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   23
            Left            =   840
            Picture         =   "frmMain.frx":2054E
            Top             =   1260
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   22
            Left            =   840
            Picture         =   "frmMain.frx":20938
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
            TabIndex        =   28
            Top             =   480
            Width           =   780
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   21
            Left            =   840
            Picture         =   "frmMain.frx":20D22
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
            TabIndex        =   27
            Top             =   885
            Width           =   855
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   20
            Left            =   840
            Picture         =   "frmMain.frx":2110C
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
            TabIndex        =   26
            Top             =   1725
            Width           =   705
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   19
            Left            =   840
            Picture         =   "frmMain.frx":214F6
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
            TabIndex        =   25
            Top             =   2130
            Width           =   705
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   18
            Left            =   840
            Picture         =   "frmMain.frx":218E0
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
            TabIndex        =   24
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
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3300
            TabIndex        =   23
            Top             =   6960
            Width           =   1125
         End
         Begin VB.Shape shpCom 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   4
            FillColor       =   &H00C0FFC0&
            Height          =   585
            Left            =   3180
            Top             =   6810
            Width           =   1365
         End
      End
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
      Index           =   23
      Left            =   270
      TabIndex        =   122
      Top             =   30
      Width           =   525
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   34
      Left            =   0
      Picture         =   "frmMain.frx":21CCA
      Top             =   0
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
      Index           =   7
      Left            =   270
      TabIndex        =   121
      Top             =   30
      Width           =   525
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   33
      Left            =   0
      Picture         =   "frmMain.frx":220B4
      Top             =   0
      Width           =   150
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "파일"
      Visible         =   0   'False
      Begin VB.Menu mnuHosp 
         Caption         =   "병원정보"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEMR 
         Caption         =   "EMR설정"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   "설정"
      Visible         =   0   'False
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
      Visible         =   0   'False
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
Dim gMnuIdx             As Integer




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

Private Sub chkResSpec_Click()

    If chkResSpec.Value = "1" Then
        txtResSpec.Enabled = True
        cmdSpecUP.Enabled = True
        cmdSpecDown.Enabled = True
    Else
        txtResSpec.Enabled = False
        cmdSpecUP.Enabled = False
        cmdSpecDown.Enabled = False
    End If
    
End Sub

Private Sub chkTest_Click()
    Dim strTest As String
    
    strTest = chkTest.Value
    
    Call WritePrivateProfileString("HOSP", "WORKTEST", strTest, App.PATH & "\INI\" & gMACH & ".ini")
    
End Sub

Private Sub cmdAnalyteFind_Click()

    frmQCList.tag = "Analyte조회"
    DoEvents
    frmQCList.Show 'vbModal
    
End Sub

'Private Sub cmdRefresh_Click()
'
'    Call GetTestList
'
'End Sub

Private Sub cmdAppend_Click()

    spdOrdMst.maxrows = spdOrdMst.maxrows + 1
    
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
    
    spdOrdMst.maxrows = spdOrdMst.maxrows - 1
    
End Sub

Private Sub cmdEMRSet_Click()
    
    frmEMRInfo.Show 'vbModal
    
End Sub

Private Sub cmdEnd_Click()

    If MsgBox("장비와 통신중입니다. 종료하시겠습니까?", vbYesNo + vbCritical, "프로그램 종료") = vbYes Then
    
        If comEqp.PortOpen = True Then
            comEqp.PortOpen = False
        End If
    
        If gDBTYPE <> "99" Then
            Call DisConnect_Server
            
            Call DisConnect_Local
        End If
        
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


Private Sub cmdLabFind_Click()

    frmQCList.Caption = "Lab조회"
    frmQCList.Show 'vbModal
    
End Sub

Private Sub cmdQCMaster_Click()

    frmQCMaster.Show 'vbModal
    
End Sub

Private Sub cmdSend_Click()
            pBuffer = txtRcv.Text
            
            Select Case UCase(gHOSP.MACHNM)
                'Case "ACCESS2":                 Call Phase_Serial_ACCESS2
                'Case "AU480":                   Call Phase_Serial_AU480
                'Case "MICROS60":                Call Phase_Serial_MICROS60
                'Case "HORIBA":                  Call Phase_Serial_HORIBA
                'Case "ISMART30":                Call Phase_Serial_ISMART30
                'Case "UROMETER720":             Call Phase_Serial_UROMETER720
                'Case "HITACHI7020":             Call Phase_Serial_HITACHI7020
                Case "LIAISON":                 Call Phase_Serial_LIAISON

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



Private Sub cmdTestOptSet_Click()
    
    frmTestOptSet.Show

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
    
    spdWork.maxrows = 0
    spdOrder.RowHeight(-1) = 15
    pDel = False
    
End Sub

Private Sub cmdWorkSearch_Click()
        
    Call GetWorkList(Format(dtpFrDt.Value, "yyyymmdd"), Format(dtpToDt.Value, "yyyymmdd"), spdWork)
    
    spdWork.RowHeight(-1) = 15

End Sub



Private Sub Command6_Click()
    
    frmScreenSet.Show 'vbModal

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
    lblRSave.ForeColor = vbBlack
    shpR.BorderColor = &H808080
    shpRC.BorderColor = &H808080
    shpRX.BorderColor = &H808080
    shpRS.BorderColor = &H808080

End Sub

Private Sub Image3_DblClick()
    If fraCommTest.Visible = False Then
        fraCommTest.Visible = True
    Else
        fraCommTest.Visible = False
    End If
End Sub

Private Sub imgTestSet_Click(Index As Integer)
    Dim Test_Property As Scripting.Dictionary
    Dim objTest_Property As clsCommon
    
    If Index = 1 Then
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
        
    ElseIf Index = 0 Then
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
            .Add "REFLF", txtRefLowF.Text
            .Add "REFHF", txtRefHighF.Text
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
        
    End If

End Sub

Private Sub imgTestSet_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer

    For i = 0 To 1
'        imgTestSet(i).ForeColor = vbBlack
        shpA(i).BorderColor = vbWhite
    Next
    
'    imgTestSet(Index).ForeColor = vbBlue
    shpA(Index).BorderColor = vbCyan

End Sub

Private Sub lblRClear_Click()
    
    spdROrder.maxrows = 0
    spdRResult.maxrows = 0
    
End Sub

Private Sub lblRClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblResult.ForeColor = vbBlack
    lblRClear.ForeColor = vbBlack
    lblRExcel.ForeColor = vbBlack
    lblRSave.ForeColor = vbBlack
    shpR.BorderColor = &H808080
    shpRC.BorderColor = &H808080
    shpRX.BorderColor = &H808080
    shpRS.BorderColor = &H808080
    
    lblRClear.ForeColor = vbBlue
    shpRC.BorderColor = vbCyan
    
End Sub

Private Sub lblResult_Click()

    frmMain.spdROrder.maxrows = 0
    frmMain.spdRResult.maxrows = 0

    Call GetResultList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), cboRstType.ListIndex, cboState.ListIndex)
    
End Sub

Private Sub lblResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblResult.ForeColor = vbBlack
    lblRClear.ForeColor = vbBlack
    lblRExcel.ForeColor = vbBlack
    lblRSave.ForeColor = vbBlack
    shpR.BorderColor = &H808080
    shpRC.BorderColor = &H808080
    shpRX.BorderColor = &H808080
    shpRS.BorderColor = &H808080
    
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
    lblRSave.ForeColor = vbBlack
    shpR.BorderColor = &H808080
    shpRC.BorderColor = &H808080
    shpRX.BorderColor = &H808080
    shpRS.BorderColor = &H808080
    
    lblRExcel.ForeColor = vbBlue
    shpRX.BorderColor = vbCyan

End Sub

Private Sub lblRSave_Click()
    Dim lRow As Long
    Dim Res  As Integer
    
    For lRow = 1 To spdROrder.DataRowCnt
        spdROrder.Row = lRow
        spdROrder.Col = 1
        If spdROrder.Value = 1 Then
            
            Res = SaveTransDataR(lRow)
        
            If Res = -1 Then
                SetForeColor spdROrder, lRow, lRow, 1, colSTATE, 255, 0, 0
                SetText spdROrder, "저장실패", lRow, colSTATE
            Else
                spdROrder.Row = lRow
                spdROrder.Col = 1
                spdROrder.Value = 1
                
                SetBackColor spdROrder, lRow, lRow, 1, colSTATE, 202, 255, 112
                SetText spdROrder, "저장완료", lRow, colSTATE
                
                      SQL = " UPDATE PATRESULT SET " & vbCrLf
                SQL = SQL & "  SENDFLAG = '2' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdROrder, lRow, colBARCODE)) & "' "
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
                
            End If
            spdROrder.Row = lRow
            spdROrder.Col = 1
            spdROrder.Value = 0
        End If
    Next lRow

End Sub

Private Sub lblRSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblResult.ForeColor = vbBlack
    lblRClear.ForeColor = vbBlack
    lblRExcel.ForeColor = vbBlack
    lblRSave.ForeColor = vbBlack
    shpR.BorderColor = &H808080
    shpRC.BorderColor = &H808080
    shpRX.BorderColor = &H808080
    shpRS.BorderColor = &H808080
    
    lblRSave.ForeColor = vbBlue
    shpRS.BorderColor = vbCyan

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
                SetText spdOrder, "저장실패", lRow, colSTATE
            Else
                spdOrder.Row = lRow
                spdOrder.Col = 1
                spdOrder.Value = 1
                
                SetBackColor spdOrder, lRow, lRow, 1, colSTATE, 202, 255, 112
                SetText spdOrder, "저장완료", lRow, colSTATE
                
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

Private Sub lblView_Click()
    Dim i          As Integer
    Dim strSPDView As String
    Dim strSPDSize As String
    
    strSPDView = ""
    
    For i = 0 To 17
        strSPDView = strSPDView & IIf(chkColumn(i).Value = "1", "1", "0")
        strSPDSize = strSPDSize & txtColumn(i).Text & "|"
    Next
    
    Call WritePrivateProfileString("VIEW", "SPDVIEW", strSPDView, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("VIEW", "SPDSIZE", strSPDSize, App.PATH & "\INI\" & gMACH & ".ini")

    '-- 컬럼보이기설정
    Call SetColumnView
    
    MsgBox "컬럼정보가 변경되었습니다.", vbInformation + vbOKOnly, Me.Caption

End Sub

Private Sub mnuBarcode_Click()
    
    mnuBarcode.Checked = True
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False
    
    Call WritePrivateProfileString("HOSP", "BARUSE", "Y", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuCheckBox_Click()
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = True
    
    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "3", App.PATH & "\INI\" & gMACH & ".ini")

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

Private Sub mnuEMR_Click()
    
    frmEMRInfo.Show 'vbModal
    
End Sub

Private Sub mnuEqpResult_Click()
    
    mnuEqpResult.Checked = True
    mnuLisResult.Checked = False
    
    Call WritePrivateProfileString("HOSP", "SAVELIS", "N", App.PATH & "\INI\" & gMACH & ".ini")

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
    
    Call WritePrivateProfileString("HOSP", "SAVELIS", "Y", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuRackPos_Click()
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = True
    mnuCheckBox.Checked = False
    
    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "2", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuSaveAuto_Click()
    
    mnuSaveAuto.Checked = True
    mnuSaveManual.Checked = False
    
    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "Y", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuSaveManual_Click()
    
    mnuSaveAuto.Checked = False
    mnuSaveManual.Checked = True
    
    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "N", App.PATH & "\INI\" & gMACH & ".ini")


End Sub

Private Sub mnuSeqno_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = True
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False
    
    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "1", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuTest_Click()
    
    Call lblMenu_Click(2)

End Sub

Private Sub spdOrder_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 And Col <> colCHECKBOX Then
        If spdOrder.UserColAction = 1 Then
            Call SetSpreadSort(spdOrder, 0)
        Else
            Call SetSpreadSort(spdOrder, 1)
        End If
        Exit Sub
    End If

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
                For intRow = .ActiveRow + 1 To .maxrows
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
        spdROrder.maxrows = spdROrder.maxrows - 1
        spdRResult.maxrows = 0
        
    End If
End Sub



Private Sub spdROrder_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol As Integer
    
    If Row = 0 Then
'        If spdROrder.UserColAction = 1 Then
'            Call SetSpreadSort(spdROrder, 0)
'        Else
'            Call SetSpreadSort(spdROrder, 1)
'        End If
        
        Exit Sub
    End If
    

    StatusBar.Panels(1).Text = GetText(spdROrder, Row, colPNAME) & " [" & GetText(spdROrder, Row, colPSEX) & "/" & GetText(spdROrder, Row, colPAGE) & "] " & _
                               "B.No:" & GetText(spdROrder, Row, colBARCODE) & " P.ID:" & GetText(spdROrder, Row, colPID)
    
    '-- 결과표시
    If GetPatTRestResult_Search(Row) = -1 Then
        '장비결과가 없을경우 검사명만 보여주기
        spdRResult.maxrows = 0
        With spdROrder
            For intCol = colSTATE + 1 To .MaxCols
                If GetText(spdROrder, Row, intCol) <> "" Then    '◆
                    spdRResult.maxrows = spdRResult.maxrows + 1
                    Call SetText(spdRResult, GetText(spdROrder, 0, intCol), spdRResult.maxrows, colRTESTNM)
                    spdRResult.RowHeight(-1) = 15
                End If
            Next
        End With
    End If

    spdRResult.RowHeight(-1) = 15
    
    'txtTV.SetFocus
    
End Sub

Private Sub spdROrder_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        If spdROrder.UserColAction = 1 Then
            Call SetSpreadSort(spdROrder, 0)
        Else
            Call SetSpreadSort(spdROrder, 1)
        End If
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
        spdRResult.maxrows = 0
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
            
            For intRow = 1 To spdRResult.maxrows
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



Private Sub spdTest_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        If spdTest.UserColAction = 1 Then
            Call SetSpreadSort(spdTest, 0)
        Else
            Call SetSpreadSort(spdTest, 1)
        End If
        
    End If

End Sub

Private Sub spdWork_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer
    
    
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
    
    
    If Row = 0 And Col <> colCHECKBOX Then
        If spdWork.UserColAction = 1 Then
            Call SetSpreadSort(spdWork, 0)
        Else
            Call SetSpreadSort(spdWork, 1)
        End If
        Exit Sub
    End If
    
    
    If Row = 0 Then Exit Sub
    
    intWRow = Row
    spdWork.Row = Row
    spdWork.Col = colBARCODE
    strBarno_Work = Trim(spdWork.Text)
    
    With frmMain.spdOrder
        If chkTest.Value = "0" Then
            blnSame = False
            For intORow = 1 To .maxrows
                .Row = intORow
                .Col = colBARCODE
                If strBarno_Work = Trim(.Text) Then
                    blnSame = True
                    Exit For
                End If
            Next
            If blnSame = False Then
                frmMain.spdOrder.maxrows = frmMain.spdOrder.maxrows + 1
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSPECNO), frmMain.spdOrder.maxrows, colSPECNO)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHECKBOX), frmMain.spdOrder.maxrows, colCHECKBOX)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colHOSPDATE), frmMain.spdOrder.maxrows, colHOSPDATE)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colBARCODE), frmMain.spdOrder.maxrows, colBARCODE)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), frmMain.spdOrder.maxrows, colSEQNO)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHARTNO), frmMain.spdOrder.maxrows, colCHARTNO)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPID), frmMain.spdOrder.maxrows, colPID)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colINOUT), frmMain.spdOrder.maxrows, colINOUT)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPNAME), frmMain.spdOrder.maxrows, colPNAME)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPSEX), frmMain.spdOrder.maxrows, colPSEX)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPAGE), frmMain.spdOrder.maxrows, colPAGE)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPJUMIN), frmMain.spdOrder.maxrows, colPJUMIN)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colOCNT), frmMain.spdOrder.maxrows, colOCNT)
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), frmMain.spdOrder.maxrows, colSEQNO)
                
                varItems = GetText(spdWork, intWRow, colITEMS)
                varItems = Split(varItems, "/")
                For intItems = 0 To UBound(varItems)
                    For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
                        frmMain.spdOrder.Row = 0
                        frmMain.spdOrder.Col = intOCol
                        If varItems(intItems) = Trim(frmMain.spdOrder.Text) Then
                            .Row = frmMain.spdOrder.maxrows
                            Call SetText(frmMain.spdOrder, "◆", frmMain.spdOrder.maxrows, intOCol)
                        End If
                    Next
                Next
                
                frmMain.spdOrder.RowHeight(-1) = 12
            End If
        '검사먼저 한 경우
        Else
            blnSame = False
            For intORow = 1 To .maxrows
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
                For intORow = 1 To .maxrows
                    .Row = intORow
                    .Col = colBARCODE
                    If strBarno_Work = Trim(.Text) Then
                        blnSame = True
                        Exit For
                    End If
                Next
                If blnSame = False Then
                    frmMain.spdOrder.maxrows = frmMain.spdOrder.maxrows + 1
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSPECNO), frmMain.spdOrder.maxrows, colSPECNO)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHECKBOX), frmMain.spdOrder.maxrows, colCHECKBOX)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colHOSPDATE), frmMain.spdOrder.maxrows, colHOSPDATE)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colBARCODE), frmMain.spdOrder.maxrows, colBARCODE)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), frmMain.spdOrder.maxrows, colSEQNO)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHARTNO), frmMain.spdOrder.maxrows, colCHARTNO)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPID), frmMain.spdOrder.maxrows, colPID)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colINOUT), frmMain.spdOrder.maxrows, colINOUT)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPNAME), frmMain.spdOrder.maxrows, colPNAME)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPSEX), frmMain.spdOrder.maxrows, colPSEX)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPAGE), frmMain.spdOrder.maxrows, colPAGE)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPJUMIN), frmMain.spdOrder.maxrows, colPJUMIN)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colOCNT), frmMain.spdOrder.maxrows, colOCNT)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), frmMain.spdOrder.maxrows, colSEQNO)
                    
                    varItems = GetText(spdWork, intWRow, colITEMS)
                    varItems = Split(varItems, "/")
                    For intItems = 0 To UBound(varItems)
                        For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
                            frmMain.spdOrder.Row = 0
                            frmMain.spdOrder.Col = intOCol
                            If varItems(intItems) = Trim(frmMain.spdOrder.Text) Then
                                .Row = frmMain.spdOrder.maxrows
                                Call SetText(frmMain.spdOrder, "◆", frmMain.spdOrder.maxrows, intOCol)
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
            spdWork.maxrows = spdWork.maxrows - 1
        End If
        
        .RowHeight(-1) = 15
    End With
    
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
            strItems = strItems & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "" & "0")
            'strItems = strItems & Format(Trim(AdoRs_Local.Fields("SENDCHANNEL").Value), "000")
            mOrder.SendCnt = mOrder.SendCnt + 1
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_AU480 = strItems
    
End Function

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_ACCESS2(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim lngIntBase  As Long
    Dim strItems    As String           '전송할 검사항목
    Dim blnISE      As Boolean          'Na, K, Cl 검사여부
    
    GetEquipExamCode_ACCESS2 = ""
    
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
            strItems = strItems & "^^^" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "\")
            mOrder.SendCnt = mOrder.SendCnt + 1
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_ACCESS2 = Mid(strItems, 1, Len(strItems) - 1)
    
End Function

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_LIAISON(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim lngIntBase  As Long
    Dim strItems    As String           '전송할 검사항목
    Dim blnISE      As Boolean          'Na, K, Cl 검사여부
    
    GetEquipExamCode_LIAISON = ""
    
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
            strItems = strItems & "\^^^" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "^")
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_LIAISON = Mid(strItems, 2)
    
End Function

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
            If .spdOrder.maxrows < intRow Then
                .spdOrder.maxrows = intRow
            End If
        End If
    
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, Trim(mOrder.Seq), intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
            
        '-- 결과스프레드 지우기
        .spdResult.maxrows = 0
    
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


Private Sub GetOrder_AU680(ByVal pBarno As String, ByVal pType As String)

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
            If .spdOrder.maxrows < intRow Then
                .spdOrder.maxrows = intRow
            End If
        End If
    
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
            
        '-- 결과스프레드 지우기
        .spdResult.maxrows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_AU680(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        'If Trim(strItems) = "" Then
        If mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            GetOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & Space(1) & mOrder.Seq & mOrder.BarNo & Space(4) & "E" & ETX
            
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            GetOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(26 - Len(mOrder.BarNo)) & mOrder.BarNo & Space(4) & "E" & strItems & ETX
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If
        
        comEqp.Output = GetOrder
        SetRawData "[Tx]" & GetOrder
        
        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub

Private Sub GetOrder_ACCESS2(ByVal pBarno As String, ByVal pType As String)

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
            If .spdOrder.maxrows < intRow Then
                .spdOrder.maxrows = intRow
            End If
        End If
    
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, Trim(mOrder.Seq), intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
            
        '-- 결과스프레드 지우기
        .spdResult.maxrows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_ACCESS2(gHOSP.MACHCD, pBarno, intRow)
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        End If
        
        '-- 진행상태(Order) 표시
        Call SetText(frmMain.spdOrder, "오더", intRow, colSTATE)
        
        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub

Private Sub GetOrder_LIAISON(ByVal pBarno As String, ByVal pType As String)

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
            If .spdOrder.maxrows < intRow Then
                .spdOrder.maxrows = intRow
            End If
        End If
    
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
            
        '-- 결과스프레드 지우기
        .spdResult.maxrows = 0
    
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
        
            '-- Order 저장
            Call SetText(frmMain.spdOrder, strItems, intRow, colKEY1)
        
        End If
                
        SetText .spdOrder, "1", intRow, colCHECKBOX
                
        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub


'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_AU680(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_AU680 = ""
    
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
            strExamCode = strExamCode & Format(Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & ""), "000")
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_AU680 = Mid(strExamCode, 2)
    
End Function



Private Sub SerialRcvData_AU680()
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
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strFunc         As String
    Dim i               As Integer
    Dim strQCTemp       As String
    
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 1, 2)
            
            Select Case strType
                Case "R "    '## Inquiry Order
                    strBarno = Trim(Mid(strRcvBuf, 14, 26))
                    strRackNo = Mid(strRcvBuf, 3, 4)
                    strTubePos = Mid(strRcvBuf, 7, 2)
                    strSeq = Mid(strRcvBuf, 9, 5)
                    
                    With mOrder
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .Seq = strSeq
                        .TubePos = strTubePos
                    End With
                    
                    Call GetOrder_AU680(Trim$(strBarno), gHOSP.RSTTYPE)
                    
                Case "D "    '## Result
                    strRackNo = Mid(strRcvBuf, 3, 4)
                    strTubePos = Mid(strRcvBuf, 7, 2)
                    strSeq = Mid(strRcvBuf, 10, 4)
                    strBarno = Trim$(Mid$(strRcvBuf, 14, 26))
                    
                    strTmp = Mid$(strRcvBuf, 45)
            
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With
            
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                                
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                        
                        
                    strTmp = Mid$(strRcvBuf, 51)
    
                    Do While Len(strTmp) >= 10
                        strIntBase = Mid$(strTmp, 2, 2)
                        strResult = Trim(Mid$(strTmp, 4, 6))
                        strComm = Mid$(strTmp, 10, 1)
                        
                        If strIntBase <> "" And strResult <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            If gPatOrdCd <> "" Then
                                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            End If
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.maxrows < lsRstRow Then
                                    .spdResult.maxrows = lsRstRow
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
                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        End If
                        strTmp = Mid$(strTmp, 12)
                    Loop
                    
                    .spdResult.RowHeight(-1) = 14
                        
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
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


Private Sub Phase_Serial_AU680()
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
            Case ETB
            Case ETX
                Call SerialRcvData_AU680
                
            Case Else
                If intBufCnt > 0 Then
                    strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                End If
        End Select
    Next i

End Sub

Private Sub Phase_Serial_ACCESS2()
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
                            Call SendOrder_ACCESS2
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
                    Case vbCr, vbLf
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
                        dtpToday.Value = Now
                        Call SerialRcvData_ACCESS2
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
                        dtpToday.Value = Now
                        intPhase = 1
                        intBufCnt = 0
                        
                        Call SerialRcvData_LIAISON
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            frmMain.comEqp.Output = ENQ
                            SetRawData "[Tx]" & ENQ
                        End If
                End Select
        End Select
    Next i
            
End Sub

Private Sub Phase_Serial_AU480()
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
            Case ETB
            Case ETX
                dtpToday = Now
                Call SerialRcvData_AU480
                
            Case Else
                If intBufCnt > 0 Then
                    strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                End If
        End Select
    Next i

End Sub

'Private Sub SerialRcvData_HORIBA()
'    Dim RS_L            As ADODB.Recordset
'    Dim strRcvBuf       As String   '수신한 Data
'    Dim strType         As String   '수신한 Record Type
'    Dim strBarno        As String   '수신한 바코드번호
'    Dim strSeq          As String   '수신한 Sequence
'    Dim strRackNo       As String   '수신한 Rack Or Disk No
'    Dim strTubePos      As String   '수신한 Tube Position
'    Dim strIntBase      As String   '수신한 장비기준 검사명
'    Dim strMachResult   As String   '수신한 장비결과
'    Dim strResult       As String   '수신한 결과(정성)
'    Dim strIntResult    As String   '수신한 결과(정량)
'    Dim strQCResult     As String   '수신한 결과(QC)
'    Dim strFlag         As String   '수신한 Abnormal Flag
'    Dim strComm         As String   '수신한 Comment
'    Dim strFunction     As String
'
'    Dim strOrderCode     As String   '처방코드
'    Dim strTestCode      As String   '검사코드
'    Dim strTestCodeSub   As String   '검사코드
'    Dim strTestName      As String   '검사명
'    Dim strSeqNo         As String   '로컬DB 검사Seq
'
'    Dim strRstRow        As String   '결과스프레드 현재 Row
'    Dim intCnt          As Integer  '통신 Frame 갯수
'    Dim intCol          As Integer  '결과컬럼 갯수
'    Dim strJudge        As String   '결과판정
'    Dim Res             As Integer
'
'    Dim strTmp          As String
'    Dim strFunc         As String
'    Dim i               As Integer
'    Dim strQCTemp       As String
'    Dim Pos As Integer
'
'    With frmMain
'        RcvBuffer = Replace(RcvBuffer, vbLf, "")
'        strRecvData = Split(RcvBuffer, vbCr)
'
'        For intCnt = 0 To UBound(strRecvData)
'            strRcvBuf = strRecvData(intCnt)
'
'            strIntBase = strType
'            strResult = ""
'
'            Select Case intCnt
'                Case 4
'                    If InStr(strRcvBuf, "AUTO_SID") > 0 Then
'                        strSeq = Mid(strRcvBuf, InStr(strRcvBuf, "AUTO_SID") + 8)
'                    End If
'
'                    With mResult
'                        .BarNo = strBarno
'                        .RsltDate = Format(Now, "yyyymmddhhmmss")
'                        .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
'                        .TubePos = strSeq
'                    End With
'
'                    Call SetPatInfo(strSeq, gHOSP.RSTTYPE)
'
'                Case 9 To 27
'                    strIntBase = Trim(Mid(strRcvBuf, 1, 2))
'                    strResult = Trim(Mid(strRcvBuf, 3))
'                    strResult = Replace(strResult, "h", "")
'                    strResult = Replace(strResult, "H", "")
'                    strResult = Replace(strResult, "l", "")
'                    strResult = Replace(strResult, "L", "")
'                    strResult = Replace(strResult, " ", "")
'
'                    If strIntBase = "'" Then
'                        strIntBase = "|"
'                    End If
'
'                    If strIntBase <> "" And strResult <> "" Then
'                        SQL = ""
'                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
'                        SQL = SQL & "  FROM EQPMASTER" & vbCr
'                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
'                        If gPatOrdCd <> "" Then
'                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
'                        End If
'
'                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                            strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                            strTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
'                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
'
'                            '-- 결과Row 추가
'                            strRstRow = .spdResult.DataRowCnt + 1
'                            If .spdResult.MaxRows < strRstRow Then
'                                .spdResult.MaxRows = strRstRow
'                            End If
'
'                            '소수점 처리, 결과 형태 처리
'                            strMachResult = strResult
'                            If strQCTemp = "1" Then
'                                strResult = SetResult(strResult, strIntBase)
'                            End If
'                            strJudge = SetJudge(strResult, strIntBase)
'
'                            '진행상태 표시("결과")
'                            SetText .spdOrder, "결과", gRow, colSTATE
'
'                            '결과값 표시
'                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                    SetText .spdOrder, strResult, gRow, intCol
'                                    Exit For
'                                End If
'                            Next
'
'                            '-- 결과 List
'                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
'                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
'                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'                            SetText .spdResult, strTestCodeSub, strRstRow, colRSUBCD          '검사코드SUB
'                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
'                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
'                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
'                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
'                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
'                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
'
'                            '-- 로컬 저장
'                            SetLocalDB gRow, strRstRow, "1", ""
'
'                            strState = "R"
'
'                            '-- 결과Count
'                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                                SetText .spdOrder, "1", gRow, colRCNT
'                            Else
'                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                            End If
'                        End If
'                    End If
'
'                Case 28
'
'                    '## DB에 결과저장
'                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
'                        Res = SaveTransData(gRow)
'
'                        If Res = -1 Then
'                            '-- 저장 실패
'                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                            SetText .spdOrder, "저장실패", gRow, colSTATE
'                        Else
'                            '-- 저장 성공
'                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                            SetText .spdOrder, "저장완료", gRow, colSTATE
'                            SetText .spdOrder, "0", gRow, colCHECKBOX
'
'                                  SQL = "Update PATRESULT Set " & vbCrLf
'                            SQL = SQL & " sendflag = '2' " & vbCrLf
'                            SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
'                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
'
'                            If DBExec(AdoCn_Local, SQL) Then
'                                '-- 성공
'                            End If
'                        End If
'                        strState = ""
'
'                        SetText .spdOrder, "0", gRow, colCHECKBOX
'
'                    End If
'
'            End Select
'        Next
'    End With
'
'End Sub

Private Sub SerialRcvData_HORIBA()
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
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim strINTRResult   As String
    
    Dim i               As Integer
    Dim J               As Integer
    
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
                    strBarno = Trim(mGetP(strRcvBuf, 3, "|"))
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With

                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                Case "R"
                    strIntBase = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^"))
                    strResult = Trim(mGetP(strRcvBuf, 4, "|"))
                    
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                            lsTestName = Trim(RS_L.Fields("TESTNAME"))
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.maxrows < lsRstRow Then
                                .spdResult.maxrows = lsRstRow
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
                                If lsTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    If strJudge = "H" Or strJudge = "L" Then
                                        .spdOrder.ForeColor = vbRed
                                    Else
                                        .spdOrder.ForeColor = vbBlack
                                    End If
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            If strJudge = "H" Or strJudge = "L" Then
                                .spdResult.ForeColor = vbRed
                            Else
                                .spdResult.ForeColor = vbBlack
                            End If
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
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
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
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
                            
'                            Call CalProcess(spdOrder, spdResult, lsTestCode)
                            
                        End If
                        strState = ""
                        
                    End If

            End Select
        Next
    End With

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
            
                    If strIntBase = "pH" Or strIntBase = "p.H" Or strIntBase = "S.G" Or strIntBase = "SG" Then
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
            
'                    Select Case Trim(strResult)
'                        Case "+":       strResult = "1+"
'                        Case "++":      strResult = "2+"
'                        Case "+++":     strResult = "3+"
'                        Case "++++":    strResult = "4+"
'                        'Case "+/-":     strResult = "Trace"
'                    End Select

                            
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.maxrows < strRstRow Then
                                .spdResult.maxrows = strRstRow
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
                                    strTestCodeSub = gArrEQP(intCol - colSTATE, 16)
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
                    End If
                
                Case 14
                            
                    '## DB에 결과저장
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then

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
                    
            End Select
        Next
    End With

End Sub


Private Sub Phase_Serial_ISMART30()
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
                        Call SerialRcvData_ISMART30
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

Private Sub Phase_Serial_MICROS60()
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
                        dtpToday.Value = Now
                        Call SerialRcvData_MICROS60
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

Private Sub SerialRcvData_ISMART30()
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
    Dim strTemp1        As String
    Dim strTemp2        As String

    Dim strTC           As String
    Dim strTG           As String
    Dim strHDL          As String
    
    Dim strHBA1C        As String
    
    Dim sFunc           As String
    
    
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

                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With

                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                Case "R"
                    strTemp1 = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strTemp2 = mGetP(strRcvBuf, 4, "|")
                    strFlag = mGetP(strRcvBuf, 7, "|")
                    strIntBase = strTemp1
                    strResult = ""
                    If InStr(strTemp2, "^") > 0 Then
                        '## 정성결과 저장
                        strResult = mGetP(strTemp2, 2, "^")
                    Else
                        '## 정량결과 저장
                        strResult = strTemp2
                    End If
                    
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.maxrows < strRstRow Then
                                .spdResult.maxrows = strRstRow
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
                                    '서브코드
                                    strTestCodeSub = gArrEQP(intCol - colSTATE, 17)
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
                    End If
                                
                    .spdResult.RowHeight(-1) = 14
                    
                Case "L"
                    '## DB에 결과저장
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
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

Private Sub SerialRcvData_MICROS60()
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
    Dim strTemp1        As String
    Dim strTemp2        As String

    Dim strTC           As String
    Dim strTG           As String
    Dim strHDL          As String
    
    Dim strHBA1C        As String
    
    Dim sFunc           As String
    
    Dim strEOS          As String
    Dim strBAS          As String
    Dim strNEU          As String
    Dim strGRA          As String
    
    
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
                    strBarno = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^"))
                    strRackNo = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 2, "^"))
                    strTubePos = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 3, "^"))

                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With

                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                Case "R"
                    strIntBase = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^"))
                    strResult = Trim(mGetP(strRcvBuf, 4, "|"))
                    
'                    If strIntBase = "EOS%" Then
'                        strEOS = strResult
'                    End If
'                    If strIntBase = "BAS%" Then
'                        strBAS = strResult
'                    End If
'                    If strIntBase = "NEU%" Then
'                        strNEU = strResult
'                    End If
RST:
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.maxrows < strRstRow Then
                                .spdResult.maxrows = strRstRow
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
                                    '서브코드
                                    strTestCodeSub = gArrEQP(intCol - colSTATE, 17)
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
                    End If
                                
                    .spdResult.RowHeight(-1) = 14
                    
'                    If strEOS <> "" And strBAS <> "" And strNEU <> "" Then
'                        If IsNumeric(strEOS) And IsNumeric(strBAS) And IsNumeric(strNEU) Then
'                            strResult = CCur(strEOS) + CCur(strBAS) + CCur(strNEU)
'                            strIntBase = "GRA%"
'                            strEOS = ""
'                            strBAS = ""
'                            strNEU = ""
'                            GoTo RST
'                        End If
'                    End If
                    
                    
                Case "L"
                    '## DB에 결과저장
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
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

'Private Sub Phase_Serial_HORIBA()
'    Dim Buffer      As Variant
'    Dim BufChar     As String
'    Dim lngBufLen   As Long
'    Dim i           As Long
'
'    lngBufLen = Len(pBuffer)
'
'    For i = 1 To lngBufLen
'        BufChar = Mid$(pBuffer, i, 1)
'        Select Case BufChar
'            Case STX
'                    dtpToday = Now
'                    RcvBuffer = ""
'                    RcvBuffer = RcvBuffer & BufChar
'            Case ETX
'                    Call SerialRcvData_HORIBA
'                    RcvBuffer = ""
'            Case Else
'                    RcvBuffer = RcvBuffer & BufChar
'        End Select
'    Next i
'
'End Sub

Private Sub Phase_Serial_HORIBA()
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
                        Call SerialRcvData_HORIBA
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

Private Sub Phase_Serial_UROMETER720()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

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
                        dtpToday = Now
                        Call SerialRcvData_UROMETER720
                        RcvBuffer = ""
                        intPhase = 1
                    Case Else
                        RcvBuffer = RcvBuffer & BufChar
                End Select
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_HITACHI7020()
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
                Call SerialRcvData_HITACHI7020
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

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_ACCESS2()
    Dim strOutput   As String     '송신할 데이터
    Dim strSpcCd    As String
    
    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&|||Host LIS|||||ACCESS||P|1" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1

        Case 2  '## Patient
            strOutput = intFrameNo & "P|1|" & mOrder.BarNo & vbCr & ETX

            intSndPhase = 3
            intFrameNo = intFrameNo + 1
        
        Case 3  '## Order
            'Specimen Convert
            Select Case Trim(mOrder.SmpType)
                Case "1"
                    strSpcCd = "Amniotic"
                Case "2"
                    strSpcCd = "Blood"
                Case "3"
                    strSpcCd = "Cervical"
                Case "4"
                    strSpcCd = "CSF"
                Case "5"
                    strSpcCd = "Plasma"
                Case "6"
                    strSpcCd = "Ratio"
                Case "7"
                    strSpcCd = "Saliva"
                Case "8"
                    strSpcCd = "Serum"
                Case "9"
                    strSpcCd = "Synovial"
                Case "10"
                    strSpcCd = "Urethral"
                Case "11"
                    strSpcCd = "Urine"
                Case "12"
                    strSpcCd = "Other"
                Case Else
                    strSpcCd = ""
            End Select
            
            If mOrder.NoOrder = True Then

                strOutput = intFrameNo & "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||A||||" & strSpcCd
                intSndPhase = 4

            Else
                If mOrder.IsSending = False Then   '## 최초 보낼때
                    strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||A||||" & strSpcCd

                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Left(strOutput, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                Else                        '## 남은 문자열이 있을때
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Left(strOutput, 230) & vbCr & ETB
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
            strOutput = intFrameNo & "L|1" & vbCr & ETX
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

Private Sub SerialRcvData_ACCESS2()
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
    Dim J               As Integer
    
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
                Case "Q"
                    strTmp = Trim(mGetP(strRcvBuf, 3, "|"))
                    strBarno = Trim(mGetP(strTmp, 2, "^"))
                    
                    With mOrder
                        .BarNo = strBarno
                    End With
                    
                    Call GetOrder_LIAISON(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "P"    '## Patient
                    strState = "P"
                
                Case "O"
                    '## DB에 결과저장
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
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
                    
                    strBarno = mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^")
                    strFlag = ""
                    strFlag = Mid(strBarno, 1, 1)
                    
                    With mResult
                        .BarNo = strBarno
                        If strFlag = "#" Then
                            .Kind = "QC"
                        Else
                            .Kind = ""
                        End If
                        If strOldBarno <> strBarno Then
                            strOldBarno = strBarno
                            .RsltDate = Format(Now, "yyyymmddhhmmss")
                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))

                            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)

                        End If
                    End With
                    
                    strState = "O"
                    
                Case "R"
                    strIntBase = Trim$(mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^"))
                    strIntResult = mGetP(strRcvBuf, 4, "|")
                    strResult = mGetP(strRcvBuf, 7, "|")
                    
                    If IsNumeric(strIntResult) Then
                        strIntResult = SetResult(strIntResult, strIntBase)
                    End If
                    
                    If mResult.Kind = "QC" Then
                        strResult = strIntResult
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, "=", "")
                    Else
                        If Not IsNumeric(strIntResult) Then
                            Select Case strResult
                                Case "N":   strResult = "Negative" & "(" & strIntResult & ")"
                                Case "P":   strResult = "Positive" & "(" & strIntResult & ")"
                                Case "<":   strResult = "Negative" & "(" & strIntResult & ")"
                                Case ">":   strResult = "Positive" & "(" & strIntResult & ")"
                                Case Else:  strResult = strResult & "(" & strIntResult & ")"
                            End Select
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
                                Case "MEAS-G" 'Measles IgG
                                        If strIntResult < 13.5 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 16.5 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "Mump-G", "MUMP-G" 'Mumps IgG
                                        If strIntResult < 9 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 11 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "Mump-M", "MUMP-M" 'Mumps IgM
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
                                Case "ToxoM", "Tox-M10" 'Toxo IgM
                                        If strIntResult < 6 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 8 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "RubG", "RUBG10" 'Rubella IgG
                                        If strIntResult < 10 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 10 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "RubM", "Rub-M" 'Rubella IgM
                                        If strIntResult < 20 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 25 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "CMVGII", "CMVG" 'CMV IgG II
                                        If strIntResult < 12 Then
                                            strResult = "Negative" & "(" & strIntResult & ")"
                                        ElseIf strIntResult >= 14 Then
                                            strResult = "Positive" & "(" & strIntResult & ")"
                                        Else
                                            strResult = "Equivocal" & "(" & strIntResult & ")"
                                        End If
                                Case "CMV-MII", "CMV-M" 'CMV IgM II
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
                    End If
                    
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER   " & vbCr
                        SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.maxrows < strRstRow Then
                                .spdResult.maxrows = strRstRow
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
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
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
            End Select
        Next
    End With

End Sub

Private Sub SerialRcvData_LIAISON()
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
    Dim strOrgIntResult As String
    Dim strIntFlag      As String
    
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
    Dim J               As Integer
    Dim strRData()    As String
    
    With frmMain
        ReDim Preserve strRData(UBound(strRecvData))
        
        For i = 1 To UBound(strRecvData)
            strRData(i) = strRecvData(i)
        Next
        
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", intCnt & ":" & strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"
                    strBarno = Trim(mGetP(strRcvBuf, 3, "|"))
                    
                    With mOrder
                        .BarNo = strBarno
                    End With
                    
                    Call GetOrder_LIAISON(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                    mPNo = 0
                
                Case "P"    '## Patient
                Case "O"
                    '## DB에 결과저장
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
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
                    
                    strBarno = Trim(mGetP(strRcvBuf, 3, "|"))
                    
                    strTemp1 = Trim(mGetP(strRcvBuf, 4, "|"))
                    strRackNo = mGetP(strTemp1, 2, "^")
                    strTubePos = mGetP(strTemp1, 3, "^")
                    
                    strRackNo = Format(strRackNo, "0000")
                    strTubePos = Format(strTubePos, "00")
                    
                    If Mid(strBarno, 1, 1) = "#" Then
                        mResult.Kind = "QC"
                    Else
                        mResult.Kind = ""
                    End If
                    
                    With mResult
                        .Seq = strSeq
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With
                
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                Case "R"
                    strIntBase = Trim$(mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^"))
                    strIntResult = mGetP(strRcvBuf, 4, "|")
                    strResult = mGetP(strRcvBuf, 7, "|")
                    strIntFlag = ""
                    
                    'If IsNumeric(strIntResult) Then
                        strOrgIntResult = strIntResult
                        If Not IsNumeric(strIntResult) Then
                            strIntFlag = Mid(strIntResult, 1, 1)
                            strIntResult = Replace(strIntResult, ">", "")
                            strIntResult = Replace(strIntResult, "<", "")
                            strIntResult = Replace(strIntResult, "=", "")
                            
                            strIntResult = SetResult(strIntResult, strIntBase)
                            strIntResult = strIntFlag & strIntResult
                        Else
                            strIntResult = SetResult(strIntResult, strIntBase)
                        End If
                    'End If
                    
                    If mResult.Kind = "QC" Then
                        strResult = strIntResult
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, "=", "")
                    Else
                        If strIntBase = "IGF-I" Or strIntBase = "ALDO" Then
                            '수치만
'                            3O|1|#81417391||^^^IGF-I^|R||||||Q||||||||||||||F
'                            4R|1|^^^IGF-I^^DOSE|150.2|ng/mL||N||F||||20180710124636
                            
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
                                    Case "Myco-M" 'Mycoplasma IgM
                                            If strIntResult < 10 Then
                                                strResult = "Negative" & "(" & strIntResult & ")"
                                            ElseIf strIntResult >= 10 Then
                                                strResult = "Positive" & "(" & strIntResult & ")"
'                                            Else
'                                                strResult = "Equivocal" & "(" & strIntResult & ")"
                                            End If
                                    
                                    Case "Myco-G" 'Mycoplasma IgG
                                            If strIntResult < 10 Then
                                                strResult = "Negative" & "(" & strIntResult & ")"
                                            ElseIf strIntResult >= 10 Then
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
                        End If
                    End If
                    
                    If strIntBase <> "" And strResult <> "" And strIntResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER   " & vbCr
                        SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.maxrows < strRstRow Then
                                .spdResult.maxrows = strRstRow
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
                            
                            '-- 수치결과가 안넘어오면 저장하지 않는다.
                            If strOrgIntResult = "" Then
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
                                
                    .spdResult.RowHeight(-1) = 14
                
                Case "L"
                    '## DB에 결과저장
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
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
    Dim J               As Integer
    
    Dim strGFR          As String
    Dim strCrea         As String
    
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
                        strIntBase = Mid$(strTmp, 1, 3)
                        strResult = Mid$(strTmp, 4, 6)
                        strResult = Trim(strResult)
                        strComm = Mid$(strTmp, 10, 1)
                        
                        
                        If strIntBase = "009" Then   'CREA
                            'GFR 계산
                            strGFR = ""
                            strCrea = strResult
                            
                            If CCur(strResult) > 0 Then
                                '18세 이상만 적용
                                If IsNumeric(strCrea) And CCur(mPatient.age) > 18 Then
                                    If mPatient.sex = "M" Then
                                        strGFR = 186 * (strCrea ^ -1.154) * (CCur(mPatient.age) ^ -0.203)
                                    ElseIf mPatient.sex = "F" Then
                                        strGFR = 186 * (strCrea ^ -1.154) * (CCur(mPatient.age) ^ -0.203) * 0.742
                                    End If
                                    
                                    If strGFR <> "" Then
                                        strGFR = Format(strGFR, "##0.00")
                                        If strGFR <= 120 Then
                                            strGFR = Round(strGFR, 2)
                                        ElseIf strGFR > 120 Then
                                            strGFR = "> 120"
                                        End If
                                    End If
                                End If
                            Else
                                strGFR = "Error"
                            End If
                            
                            If CCur(strResult) < 0.2 Then strResult = "< 0.2"
                        End If
'RST:
                        If strIntBase <> "" And strResult <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            If gPatOrdCd <> "" Then
                                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            End If
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.maxrows < strRstRow Then
                                    .spdResult.maxrows = strRstRow
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
                        End If
                        strTmp = Mid$(strTmp, 12)
                    Loop
                    
                    '-- GFR 저장
                    If strGFR <> "" Then
                        strIntBase = "088"
                        strResult = strGFR
                        
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.maxrows < strRstRow Then
                                .spdResult.maxrows = strRstRow
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
                    End If
                    
                    .spdResult.RowHeight(-1) = 14
                
'                    '## DB에 결과저장
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
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
                            
                            'Call CalProcess(spdOrder, spdResult, strTestCode)
                            
                        End If
                        strState = ""
                        
                    End If
            End Select
        Next
    End With

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

Private Sub GetOrder_HITACHI7020(ByVal pBarno As String, ByVal pType As String)

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
                            mOrder.PNAME = Trim(GetText(frmMain.spdOrder, i, colPNAME))
                            mOrder.PID = Trim(GetText(frmMain.spdOrder, i, colPID))
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            'Exit Sub
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.maxrows < intRow Then
                .spdOrder.maxrows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시

        '-- 결과스프레드 지우기
        .spdResult.maxrows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 12

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_HITACHI7020(gHOSP.MACHCD, pBarno, intRow)

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

        End If

        '-- 현재 Row
        gRow = intRow

    End With

End Sub

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_HITACHI7020(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim lngIntBase  As Long
    Dim strItems    As String           '전송할 검사항목
    Dim blnISE      As Boolean          'Na, K, Cl 검사여부
    
    GetEquipExamCode_HITACHI7020 = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    strItems = String$(37, "0")
    
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
    
    GetEquipExamCode_HITACHI7020 = strItems
    
End Function

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_HITACHI7020()
    Dim strOutput   As String     '송신할 데이터
    
    strOutput = ";" & mOrder.Function
    strOutput = strOutput & " 37"
    strOutput = strOutput & Mid(mOrder.Order, 1, 37)
    strOutput = strOutput & "00000"
    
    Call Sleep(100)
    
    '-- SPE Send(오더전송)
    comEqp.Output = STX & strOutput & ETX '& vbCr & vbLf
    SetRawData "[Tx]" & STX & strOutput & ETX '& vbCr & vbLf

End Sub

Private Sub SerialRcvData_HITACHI7020()
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
    Dim J               As Integer
    
    Dim strGFR          As String
    Dim strCrea         As String
    Dim strFunc         As String
    Dim strFunction     As String
    

    Dim strTC           As String
    Dim strTG           As String
    Dim strHDL          As String
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
        '    strRcvBuf = RcvBuffer
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 1, 1)
            
            Select Case strType
                Case ">"                'ANY 수신
                    Sleep (100)
                    Call SndMore        'MOR Send
                    Do
                    Loop Until comEqp.OutBufferCount = 0
                
                Case "?"                'REP 수신
                    Sleep (100)
                    Call SndMore        'MOR Send
                    Do
                    '   DoEvents
                    Loop Until frmMain.comEqp.OutBufferCount = 0
                
                Case "?"                'SUS 수신
                    Sleep (100)
                    Call SndMore        'MOR Send
                    Do
                    '   DoEvents
                    Loop Until frmMain.comEqp.OutBufferCount = 0
                    
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
                    Call SndMore
                
                Case ";"                'SPE
                    strFunc = Mid(strRcvBuf, 2, 1)              ' Function
                    strSeq = Mid(strRcvBuf, 4, 5)               ' Sample No
                    strRackNo = Mid(strRcvBuf, 9, 1)            ' Rack No
                    strTubePos = Mid(strRcvBuf, 10, 3)          ' Pos No
                    strBarno = Trim(Mid(strRcvBuf, 13, 13))     ' Barcode
                    strFunction = Mid(strRcvBuf, 2, 12) & String(13, "#") & Mid(strRcvBuf, 27, 15)

                    If strFunc = "K" Or strFunc = "L" Or strFunc = "G" Or strFunc = "H" Then
                         Exit Sub
                     End If
                    
                    With mOrder
                        .Func = strFunc
                        .Seq = strSeq
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .BarNo = strBarno
                    End With
                    
                    Call GetOrder_HITACHI7020(Trim$(strBarno), gHOSP.RSTTYPE)
                    
                    strFunction = Replace(strFunction, String(13, "#"), Left(mOrder.BarNo & Space(13), 13))
                    
                    mOrder.Function = strFunction
                    
                    Call SendOrder_HITACHI7020
                    
                    Call SetText(frmMain.spdOrder, "0", gRow, colCHECKBOX)
                    Call SetText(frmMain.spdOrder, "오더전송", gRow, colSTATE)
                
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
                    
                    Call SndMore            'MOR Send
            
                    If strFunc <> "@" And strFunc <> "M" Then
                        'QC
                        If strFunc = "F" Then
                            strBarno = Trim(Mid(strRcvBuf, 6, 10))
                            strBarno = "QC" & strBarno
                        End If
                        
                        strRackNo = Mid(strRcvBuf, 9, 1)
                        strTubePos = Trim(Mid(strRcvBuf, 10, 3))
                        mOrder.Seq = strTubePos
                        strBarno = Trim(Mid(strRcvBuf, 14, 13))
                        
                        With mResult
                            .Seq = strTubePos
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
                        
                        ':f 01003  1      Biorad 1    0                15 1   6.8  2   4.0  3    45  4    35  5   103  6   1.3  8    68 10   177 11  15.4 12   2.7 13   259 14    83 15    69 16   140 17   5.5

                        strTmp = Mid$(strRcvBuf, 44)
                        'Do While Len(strTmp) >= 9
        
        
                        For i = 1 To Len(strTmp) Step 10
                            strIntBase = Trim(Mid(strTmp, i, 3))
                            strResult = Trim(Mid(strTmp, i + 3, 6))
                            strComm = Trim(Mid(strTmp, i + 9, 1))
                            
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
                                SQL = ""
                                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                                SQL = SQL & "  FROM EQPMASTER" & vbCr
                                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                                If gPatOrdCd <> "" Then
                                    SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                                End If
                                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                    strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                    strTestName = Trim(RS_L.Fields("TESTNAME"))
                                    strSeqNo = Trim(RS_L.Fields("SEQNO"))
                                    strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
            
                                    '-- 결과Row 추가
                                    strRstRow = .spdResult.DataRowCnt + 1
                                    If .spdResult.maxrows < strRstRow Then
                                        .spdResult.maxrows = strRstRow
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
                            End If
                        Next
                    End If
                    
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

                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            strTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")

                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.maxrows < strRstRow Then
                                .spdResult.maxrows = strRstRow
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
                    End If

'                    '## DB에 결과저장
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
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
                            
                            'Call CalProcess(spdOrder, spdResult, strTestCode)
                            
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
                'Case "ACCESS2":                 Call Phase_Serial_ACCESS2
                'Case "AU480":                   Call Phase_Serial_AU480
                'Case "MICROS60":                Call Phase_Serial_MICROS60
                'Case "HORIBA":                  Call Phase_Serial_HORIBA
                'Case "ISMART30":                Call Phase_Serial_ISMART30
                'Case "UROMETER720":             Call Phase_Serial_UROMETER720
                'Case "HITACHI7020":             Call Phase_Serial_HITACHI7020
                Case "LIAISON":                 Call Phase_Serial_LIAISON

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


Private Sub SetColumnView()
    Dim i       As Integer
    Dim varSize As Variant
    
    varSize = Split(gCOLSIZE, "|")
    
    For i = 0 To UBound(varSize) - 1
        spdOrder.Col = i + 2
        If Mid(gCOLVIEW, i + 1, 1) = 1 Then
            spdOrder.ColHidden = False
            chkColumn(i).Value = "1"
            frmScreenSet.chkColumn(i).Value = "1"
        Else
            spdOrder.ColHidden = True
            chkColumn(i).Value = "0"
            frmScreenSet.chkColumn(i).Value = "0"
        End If
        spdOrder.ColWidth(i + 2) = varSize(i)
    
    
        '워크리스트
        If i >= 2 Then
            frmWorkList.spdWork.Col = i + 2
            If Mid(gCOLVIEW, i + 1, 1) = 1 Then
                frmWorkList.spdWork.ColHidden = False
                chkColumn(i).Value = "1"
            Else
                frmWorkList.spdWork.ColHidden = True
                chkColumn(i).Value = "0"
            End If
            frmWorkList.spdWork.ColWidth(i + 2) = varSize(i)
        End If
    Next

    For i = 0 To UBound(varSize) - 1
        spdROrder.Col = i + 2
        If Mid(gCOLVIEW, i + 1, 1) = 1 Then
            spdROrder.ColHidden = False
            chkColumn(i).Value = "1"
        Else
            spdROrder.ColHidden = True
            chkColumn(i).Value = "0"
        End If
        spdROrder.ColWidth(i + 2) = varSize(i)
    
    Next

End Sub

Private Sub Form_Load()
    Dim lngConnect  As Long
    Dim strMsg      As String

On Error GoTo RST
'On Error Resume Next

    Me.Width = 20940
    Me.Height = 12585
    
    lblHospInfo.Caption = gHOSP.HOSPNM & "  " & gHOSP.MACHNM & "  " & gHOSP.USERNM & "[" & gHOSP.USERID & "]" '& "버전 " & App.Major & "." & App.Minor & "." & App.Revision
    
    'Me.Caption = gHOSP.MACHNM
    Me.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [장비와 인터페이스]     ◈◈◈◈◈"
    
    Call CtlInitializing
        
    '-- Menu Set
    'Call SetMenu
    
    '-- 컬럼보이기설정
    Call SetColumnView
    
    '-- 검사코드
    Call GetTestList
    
    '-- 오더코드
    Call GetOrderMST

    '-- 검사명 보이기
    Call SetExamCode
    
    '-- 통신설정
    Call GetCommList

    '-- 통신열기
    Call OpenCommunication
    
    If gWORKTEST = "0" Then
        chkTest.Value = "0"
    Else
        chkTest.Value = "1"
    End If
    
    '순번사용
    If gHOSP.RSTTYPE = "1" Then
        txtSeqNo.Visible = True
    Else
        txtSeqNo.Visible = False
    End If
    
    
    pDel = False
    
        
    If gComm.COMTYPE = "" Then
        lblMenu(3).BackColor = &HFFFFC0
        frame4.Visible = True
        frame4.ZOrder 0
        Call lblMenu_Click(3)
    Else
        lblMenu(0).BackColor = &HFFFFC0
        frame1.Visible = True
        frame1.ZOrder 0
    End If
    
    '변수 초기화(Advia1650)
    iPendingFlag = 0: iTotQueryFlag = 0: iTmpPendingFlag = 0: iIdleFlag = 0
    iOrderFlag = 0: iResultFlag = 0
    sRcvState = "": sSndState = ""
    intPhase = 1
    
    miLineNo = 0
        
    Call setStatusBar
    
    shpB(0).BorderColor = vbBlue
    lblMenu(0).ForeColor = vbRed
    lblMenu(0).FontBold = True
    
    lngConnect = dce_setenv(App.PATH & "\sl.env", "", "")
    If lngConnect = 0 Then Call dce_error(strMsg)
    
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

Private Sub setStatusBar()
    Dim strBarUse   As String
    Dim strSaveAuto As String
    Dim strRstType  As String
    Dim strSaveLIS  As String
    
    If gHOSP.BARUSE = "Y" Then
        strBarUse = "◆ 바코드 사용 "
        strRstType = ""
    Else
        strBarUse = "◆ 바코드 미사용 "
        
        Select Case gHOSP.RSTTYPE
            Case "1": strRstType = "◆ 장비 순번 사용 "
            Case "2": strRstType = "◆ 장비 RACK/POS 사용 "
            Case "3": strRstType = "◆ IF 워크리스트 사용 "
        End Select
    End If
    
    If gHOSP.SAVELIS = "1" Then
        strSaveLIS = "◆ 장비결과 전송 "
    Else
        strSaveLIS = "◆ LIS결과 전송 "
    End If

    
    If gHOSP.SAVEAUTO = "Y" Then
        strSaveAuto = "◆ 자동전송 "
    Else
        strSaveAuto = "◆ 수동저장(선택저장) "
    End If
    
    
    StatusBar.Panels(1).MinWidth = frmMain.Width * 0.5
    StatusBar.Panels(2).MinWidth = frmMain.Width * 0.25
    StatusBar.Panels(3).MinWidth = frmMain.Width * 0.25
    
    StatusBar.Panels(3).Text = strBarUse & strRstType & strSaveLIS & strSaveAuto
'    StatusBar.Panels(2).Text = IIf(gHOSP.SAVEAUTO = "Y", "자동전송", "수동저장(선택저장)")

End Sub

Public Sub OpenCommunication()

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
            StatusBar.Panels(2).Text = "COM" & comEqp.CommPort & " 포트에 연결 되었습니다"
            imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        
            imgConn.Picture = imlStatus.ListImages("ON").ExtractIcon
            imgConn.ToolTipText = "연결성공"
        
        Else
            lblStatus.Caption = "통신포트에 연결 되지 않았습니다"
            StatusBar.Panels(2).Text = "통신포트에 연결 되지 않았습니다"
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        
            imgConn.Picture = imlStatus.ListImages("OFF").ExtractIcon
            imgConn.ToolTipText = "연결실패"
        End If
    ElseIf gComm.COMTYPE = "2" Then
        If gComm.TCPTYPE = "1" Then
'            wSck.LocalPort = CInt(gComm.TCPPORT)
'            wSck.Listen

            lblStatus.Caption = "TCP " & gComm.TCPPORT & " 포트로 연결합니다"
            StatusBar.Panels(2).Text = "TCP " & gComm.TCPPORT & " 포트로 연결합니다"
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Visible = False
            imgReceive.Visible = False
            lblSend.Visible = False
            lblRcv.Visible = False
            'imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            'imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        Else
'            wSck.Close
'            wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)

            lblStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 포트로 연결합니다"
            StatusBar.Panels(2).Text = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 포트로 연결합니다"
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Visible = False
            imgReceive.Visible = False
            lblSend.Visible = False
            lblRcv.Visible = False
            'imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            'imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        End If
    ElseIf gComm.COMTYPE = "" Then
    
    End If
End Sub


Public Sub GetCommList()
    Dim i As Integer
    Dim Ret As Integer
    
    optComType(0).BackColor = &H808080
    optComType(1).BackColor = &H808080
    
    If gComm.COMTYPE = "1" Then
        optComType(0).Value = True
        frameCom.Enabled = True
        frameTCP.Enabled = False
        optComType(0).BackColor = &HFF8080
    Else
        optComType(1).Value = True
        frameCom.Enabled = False
        frameTCP.Enabled = True
        optComType(1).BackColor = &HFF8080
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
    
    Me.Top = 0
    
    '-- 인터페이스
    frame1.Top = 1850
    frame2.Top = 1850
    frame3.Top = 1850
    frame4.Top = 1850
    frame1.Width = Me.ScaleWidth - 150
    frame1.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150 - StatusBar.Height
    
    If gWORKPOS = "P" Then
        cmdWorkSearch.Visible = False
        cmdWorkAll.Visible = False
        chkTest.Visible = False
        
        spdOrder.Width = Me.ScaleWidth - spdResult.Width - 400
        spdOrder.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500 - StatusBar.Height
        spdResult.Left = spdOrder.Left + spdOrder.Width + 50
        spdResult.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500 - StatusBar.Height
    Else
        cmdSL.Left = spdOrder.Left + 50
        chkAll.Left = spdOrder.Left + 550
        
        spdWork.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - cmdWorkAll.Height - 500 - StatusBar.Height
        spdOrder.Left = spdWork.Width + 100
        spdOrder.Width = Me.ScaleWidth - spdWork.Width - spdResult.Width - 400
        spdOrder.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500 - StatusBar.Height
        spdResult.Left = spdOrder.Left + spdOrder.Width + 50
        spdResult.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500 - StatusBar.Height
    End If
    
    
    '-- 결과조회
    frame2.Width = Me.ScaleWidth - 150
    frame2.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150 - StatusBar.Height
    
    spdROrder.Width = Me.ScaleWidth - spdRResult.Width - 500
    spdROrder.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500 - StatusBar.Height
    
    spdRResult.Left = spdROrder.Left + spdROrder.Width + 50
    spdRResult.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500 - StatusBar.Height
    
    '-- 검사설정
    frame3.Width = Me.ScaleWidth - 150
    frame3.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150 - StatusBar.Height
    
    spdTest.Width = Me.ScaleWidth - frameTestSet.Width - 600
    spdTest.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500 - StatusBar.Height
    
    frameTestSet.Left = spdTest.Left + spdTest.Width + 50
    frameTestSet.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500 - StatusBar.Height

    '-- 통신설정
    frame4.Width = Me.ScaleWidth - 150
    frame4.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150 - StatusBar.Height

    Call setStatusBar
    
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

'    For i = 0 To 3
'        lblActionTest(i).ForeColor = vbBlack
'        'shpA(i).BorderColor = &H808080
'        shpA(i).BorderColor = vbWhite
'    Next
    
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
    spdWork.maxrows = 0
    spdOrder.maxrows = 0
    spdResult.maxrows = 0
    
    If gWORKPOS = "P" Then
        spdWork.Visible = False
    Else
        spdWork.Visible = True
    End If
    txtSeqNo.Text = 1
    
    '-- 장비결과
    spdROrder.maxrows = 0
    spdRResult.maxrows = 0
        
    '-- 검사코드 설정
    spdTest.maxrows = 0
    
    cboCOL.AddItem ""
    cboCOL.AddItem "<"
    cboCOL.AddItem "<="
    cboCOL.ListIndex = 0
    
    cboCOH.AddItem ""
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
    cboBaudrate.AddItem ("38400")
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
    
    
    lblCommStatus.Caption = ""
    
    txtBarcode.Text = ""
    
    dtpFrDt.Value = Now
    dtpToDt.Value = Now
    

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
    
    ElseIf Index = 4 Then
'        If Trim(txtEqpCD.Text) = "" Then
'            MsgBox "검사항목을 먼저 선택하세요", vbCritical, Me.Caption
'            Exit Sub
'        End If
'
'        If Trim(txtTestCd.Text) = "" Then
'            MsgBox "검사코드를 입력하세요", vbCritical, Me.Caption
'            txtTestCd.SetFocus
'            Exit Sub
'        End If
'
'        If Trim(txtTestNm.Text) = "" Then
'            MsgBox "검사명을 입력하세요", vbCritical, Me.Caption
'            txtTestNm.SetFocus
'            Exit Sub
'        End If
'
'        Set Test_Property = New Scripting.Dictionary
'
'        With Test_Property
'            .Add "EQPCD", txtEqpCD.Text
'            .Add "SEQ", txtSeq.Text
'            .Add "OCH", txtOChannel.Text
'            .Add "RCH", txtRChannel.Text
'            .Add "TESTCD", txtTestCd.Text
'            .Add "TESTNM", txtTestNm.Text
'            .Add "ABBRNM", txtAbbrNm.Text
'            .Add "RES", txtResSpec.Text
'            .Add "REFL", txtRefLow.Text
'            .Add "REFH", txtRefHigh.Text
'            .Add "RSTTYPE", cboResultType.Text
'            If optCutUse(0).Value = True Then
'                .Add "CUTUSE", "N"
'            Else
'                .Add "CUTUSE", "Y"
'            End If
'            .Add "COLIN", txtCOLIn.Text
'            .Add "COLCP", cboCOL.Text
'            .Add "COLOUT", txtCOLOut.Text
'            .Add "COHIN", txtCOHIn.Text
'            .Add "COHCP", cboCOH.Text
'            .Add "COHOUT", txtCOHOut.Text
'            .Add "COMOUT", txtCOMOut.Text
'            '-- QC
'            .Add "LAB", txtLab.Text
'            .Add "LOT", txtLot.Text
'            .Add "ANALYTE", txtAnalyte.Text
'            .Add "METHOD", txtMethod.Text
'            .Add "INSTRUMENT", txtInstrument.Text
'            .Add "REAGENT", txtReagent.Text
'            .Add "UNIT", txtUnit.Text
'
'            .Add "TEMP", chkResSpec.Value
'
'        End With
'
'        Set objTest_Property = New clsCommon
'
'        With objTest_Property
'            .SetAdoCn AdoCn_Local
'            If Not .EditTestInfo(Test_Property) Then
'                '-- 저장 오류
'                'Call GetTestList
'            End If
'        End With
'
'        Call GetTestList
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
    
    spdWork.maxrows = 0
    spdOrder.maxrows = 0
    spdResult.maxrows = 0

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
        Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\INI\" & gMACH & ".ini")
    End If

    If optComType(0).Value = True Then
        Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    Call WritePrivateProfileString("COMM", "COMPORT", cboPort.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("COMM", "SPEED", cboBaudrate.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("COMM", "PARITY", cboParity.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("COMM", "DATABIT", cboDatabit.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("COMM", "STARTBIT", cboStartbit.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("COMM", "STOPBIT", cboStopbit.Text, App.PATH & "\INI\" & gMACH & ".ini")
    If chkRTS.Value = "1" Then
        Call WritePrivateProfileString("COMM", "RTSEnable", "True", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        Call WritePrivateProfileString("COMM", "RTSEnable", "False", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    If chkRTS.Value = "1" Then
        Call WritePrivateProfileString("COMM", "DTREnable", "True", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        Call WritePrivateProfileString("COMM", "DTREnable", "False", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    Call GetSetup
    
    Call GetCommList
    
    'Call OpenCommunication
    
    MsgBox "통신정보가 변경되었습니다.", vbInformation + vbOKOnly, Me.Caption

End Sub

Private Sub lblComSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblComSave.ForeColor = vbBlue
    shpCom.BorderColor = vbCyan

End Sub

Private Sub lblMenu_Click(Index As Integer)

    gMnuIdx = Index
    
    frame1.Visible = False
    frame2.Visible = False
    frame3.Visible = False
    'frame4.Visible = False
    fraInterface.Visible = False
    fraResult.Visible = False
    fraSet.Visible = False

    
    lblMenu(0).BackColor = vbWhite
    lblMenu(1).BackColor = vbWhite
    lblMenu(2).BackColor = vbWhite
    lblMenu(3).BackColor = vbWhite
    lblMenu(Index).BackColor = &HFFFFC0
    
    lblMenu(0).FontBold = False
    lblMenu(1).FontBold = False
    lblMenu(2).FontBold = False
    lblMenu(3).FontBold = False
    
    shpB(0).BorderColor = vbGreen
    shpB(1).BorderColor = vbGreen
    shpB(2).BorderColor = vbGreen
    shpB(3).BorderColor = vbGreen
    
     
    Select Case Index
        Case 0:
                frame1.Visible = True
                frame1.ZOrder 0
                chkAll.ZOrder 0
                fraInterface.Visible = True
                frmMain.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [장비와 인터페이스]     ◈◈◈◈◈"
                
                tmrComm.Enabled = False
                tmrFlipFlop.Enabled = False
                
                lblCommStatus.Caption = ""
        Case 1:
                frame2.Visible = True
                frame2.ZOrder 0
                chkRAll.ZOrder 0
        
                fraResult.Visible = True
                frmMain.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [검사 결과 조회]     ◈◈◈◈◈"
        Case 2:
                frame3.Visible = True
                frame3.ZOrder 0
    
                '-- 검사코드
                Call GetTestList
                frmMain.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [검사 코드 설정]     ◈◈◈◈◈"
        
        Case 3:
                frame4.Visible = True
                frame4.ZOrder 0
    
                fraSet.Visible = True
                
                '-- 통신설정
                Call GetCommList
                
'                '-- 화면설정
'                Call SetColumnName
                
                frmMain.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [장비 통신 설정]     ◈◈◈◈◈"
    
    End Select
    
    shpB(Index).BorderColor = vbBlue
    lblMenu(Index).ForeColor = vbRed
    lblMenu(Index).FontBold = True
    
End Sub

Private Sub SetColumnName()
    Dim i As Integer
    
    chkColumn(0).Caption = "검사일시"
    chkColumn(1).Caption = "저장순번"
    chkColumn(2).Caption = "접수일자"
    chkColumn(3).Caption = "검체번호 (바코드)"
    chkColumn(4).Caption = "Seq"
    chkColumn(5).Caption = "RACK"
    chkColumn(6).Caption = "POS"
    chkColumn(7).Caption = "입원/외래"
    chkColumn(8).Caption = "챠트번호"
    chkColumn(9).Caption = "환자번호"
    chkColumn(10).Caption = "환자이름"
    chkColumn(11).Caption = "성별"
    chkColumn(12).Caption = "나이"
    chkColumn(13).Caption = "주민번호"
'    chkColumn(14).Caption = ""
'    chkColumn(15).Caption = ""
    chkColumn(16).Caption = "오더갯수"
    chkColumn(17).Caption = "결과갯수"
    
    For i = 0 To 17
        txtColumn(i).Alignment = 2
        txtColumn(i).Text = spdOrder.ColWidth(i + 2)
    Next

End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer

    For i = 0 To 3
        If gMnuIdx <> i Then
            lblMenu(i).ForeColor = vbBlack
            shpB(i).BorderColor = vbGreen
        End If
    Next
    
    If gMnuIdx <> Index Then
        lblMenu(Index).ForeColor = vbBlue
        shpB(Index).BorderColor = vbCyan
    End If

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
        Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    If optTCPType(0).Value = True Then
        Call WritePrivateProfileString("COMM", "TCPTYPE", "1", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        Call WritePrivateProfileString("COMM", "TCPTYPE", "2", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    Call WritePrivateProfileString("COMM", "TCPIP", txtTCPIP.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("COMM", "TCPPORT", txtTCPPort.Text, App.PATH & "\INI\" & gMACH & ".ini")
    
    Call GetSetup
    
    Call GetCommList
    
'    Call OpenCommunication

    MsgBox "통신정보가 변경되었습니다.", vbInformation + vbOKOnly, Me.Caption

End Sub

Private Sub lblTcpSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblTcpSave.ForeColor = vbBlue
    shpTcp.BorderColor = vbCyan

End Sub

Private Sub lblWork_Click()
    
    If gWORKPOS = "P" Then
        frmWorkList.Show 'vbModal
    Else
        Call GetWorkList(Format(dtpFrDt.Value, "yyyymmdd"), Format(dtpToDt.Value, "yyyymmdd"), spdWork)
        
        spdWork.RowHeight(-1) = 15

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
        optComType(0).BackColor = &H808080
        optComType(1).BackColor = &H808080
        optComType(0).BackColor = &HFF8080
    Else
        frameCom.Enabled = False
        frameTCP.Enabled = True
        optComType(0).BackColor = &H808080
        optComType(1).BackColor = &H808080
        optComType(1).BackColor = &HFF8080
    End If

End Sub


Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer

    For i = 0 To 3
        If gMnuIdx <> i Then
            lblMenu(i).ForeColor = vbBlack
            shpB(i).BorderColor = vbGreen
        End If
    Next
    
    lblWork.ForeColor = vbBlack
    lblSave.ForeColor = vbBlack
    lblClear.ForeColor = vbBlack
    lblResult.ForeColor = vbBlack
    
    shpW.BorderColor = &H808080
    shpS.BorderColor = &H808080
    shpC.BorderColor = &H808080
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
    StatusBar.Panels(1).Text = GetText(spdOrder, Row, colPNAME) & " [" & GetText(spdOrder, Row, colPSEX) & "/" & GetText(spdOrder, Row, colPAGE) & "] " & _
                               "B.No:" & GetText(spdOrder, Row, colBARCODE) & " P.ID:" & GetText(spdOrder, Row, colPID)
    
    '-- 결과표시
    If GetPatTRestResult(Row) = -1 Then
        '장비결과가 없을경우 검사명만 보여주기
        spdResult.maxrows = 0
        With spdOrder
            For intCol = colSTATE + 1 To .MaxCols
                If GetText(spdOrder, Row, intCol) <> "" Then    '◆
                    spdResult.maxrows = spdResult.maxrows + 1
                    Call SetText(spdResult, GetText(spdOrder, 0, intCol), spdResult.maxrows, colRTESTNM)
                    spdResult.RowHeight(-1) = 15
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
    Dim strBarcode  As String
    
On Error GoTo RST

    GetPatTRestResult = -1
    intRow = 0
    
    intSeq = GetText(spdOrder, asRow, colSAVESEQ)
    strExamDate = Mid(GetText(spdOrder, asRow, colEXAMDATE), 1, 8)
    strBarcode = GetText(spdOrder, asRow, colBARCODE)
    
    If intSeq = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT SEQNO, EQUIPCODE,EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT,REFFLAG,REFJUDGE" & vbCr
    SQL = SQL & "  FROM PATRESULT " & vbCr
    SQL = SQL & " WHERE SAVESEQ  = " & intSeq & vbCr
    SQL = SQL & "   AND EXAMDATE = '" & strExamDate & "'" & vbCr
    SQL = SQL & "   AND BARCODE  = '" & strBarcode & "'" & vbCr
'    SQL = SQL & " ORDER BY SEQNO "
    
    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        With frmMain.spdResult
            .maxrows = 0
            .maxrows = AdoRs_Local.RecordCount
            Do Until AdoRs_Local.EOF
                intRow = intRow + 1
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("SEQNO").Value & "", intRow, colRSEQNO)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("EQUIPCODE").Value & "", intRow, colRCHANNEL)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("EXAMCODE").Value & "", intRow, colRTESTCD)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("EXAMNAME").Value & "", intRow, colRTESTNM)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("EQUIPRESULT").Value & "", intRow, colRMACHRESULT)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("RESULT").Value & "", intRow, colRLISRESULT)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("REFFLAG").Value & "", intRow, colRFLAG)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("REFJUDGE").Value & "", intRow, colRJUDGE)
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
            .maxrows = 0
            .maxrows = AdoRs_Local.RecordCount
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
        For intRow = 1 To spdOrdMst.maxrows
                  SQL = ""
            SQL = SQL & "INSERT INTO ORDMASTER (ORDERCODE,ORDERNAME) VALUES ("
            SQL = SQL & "'" & GetText(spdOrdMst, intRow, 1) & "','')"
            
            Call DBExec(AdoCn_Local, SQL)
        Next
    End If
    
End Sub

Private Sub spdTest_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
'        If spdTest.UserColAction = 1 Then
'            Call SetSpreadSort(spdTest, 0)
'        Else
'            Call SetSpreadSort(spdTest, 1)
'        End If
        
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
        txtRefLowF.Text = GetText(spdTest, Row, colLLOWF)
        txtRefHighF.Text = GetText(spdTest, Row, colLHIGHF)
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
            txtResSpec.Enabled = True
            cmdSpecUP.Enabled = True
            cmdSpecDown.Enabled = True
        Else
            chkResSpec.Value = "0"
            txtResSpec.Enabled = False
            cmdSpecUP.Enabled = False
            cmdSpecDown.Enabled = False
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
            If .spdOrder.maxrows < intRow Then
                .spdOrder.maxrows = intRow
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
        .spdResult.maxrows = 0
    
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



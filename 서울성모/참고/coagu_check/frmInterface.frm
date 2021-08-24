VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   Caption         =   " CoaguCheck Interface_건국대학교병원"
   ClientHeight    =   10980
   ClientLeft      =   855
   ClientTop       =   300
   ClientWidth     =   14925
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInterface.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11843.6
   ScaleMode       =   0  '사용자
   ScaleWidth      =   30500.29
   Begin VB.CheckBox chkMode 
      Caption         =   "AUTO"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   8205
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   120
      Value           =   1  '확인
      Width           =   900
   End
   Begin Threed.SSCommand cmdResMach 
      Height          =   615
      Left            =   9150
      TabIndex        =   82
      Top             =   120
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "결과매칭"
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
   Begin Threed.SSCommand Command_close 
      Height          =   615
      Left            =   13740
      TabIndex        =   81
      Top             =   120
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "종료"
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
   Begin Threed.SSCommand Command_setup 
      Height          =   615
      Left            =   12585
      TabIndex        =   80
      Top             =   120
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "코드설정"
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
   Begin Threed.SSCommand Command_Config 
      Height          =   615
      Left            =   11445
      TabIndex        =   79
      Top             =   120
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "통신설정"
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
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1035
      Left            =   15030
      TabIndex        =   25
      Top             =   1170
      Visible         =   0   'False
      Width           =   3465
      Begin VB.Timer tm_H232 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   1440
         Top             =   360
      End
      Begin VB.Timer tmResRequest 
         Interval        =   4000
         Left            =   2280
         Top             =   360
      End
      Begin VB.Timer tm_H232_2 
         Enabled         =   0   'False
         Left            =   1860
         Top             =   360
      End
      Begin VB.Timer tmErr 
         Interval        =   1000
         Left            =   2700
         Top             =   360
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   60
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         InputLen        =   1
         RThreshold      =   1
         RTSEnable       =   -1  'True
         EOFEnable       =   -1  'True
      End
      Begin MSCommLib.MSComm MSComm2 
         Left            =   660
         Top             =   210
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         InputLen        =   1
         RThreshold      =   1
         RTSEnable       =   -1  'True
         EOFEnable       =   -1  'True
      End
   End
   Begin Threed.SSPanel spErr 
      Height          =   435
      Left            =   15030
      TabIndex        =   24
      Top             =   9045
      Visible         =   0   'False
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   767
      _StockProps     =   15
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCol1 
      Caption         =   "<"
      Height          =   315
      Left            =   17190
      TabIndex        =   23
      Top             =   5805
      Visible         =   0   'False
      Width           =   255
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   7995
      _Version        =   65536
      _ExtentX        =   14102
      _ExtentY        =   1191
      _StockProps     =   15
      Caption         =   "  CoaguCheck Interface"
      ForeColor       =   16777215
      BackColor       =   11494691
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   345
         Left            =   5235
         TabIndex        =   13
         Top             =   180
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21430273
         CurrentDate     =   40534
      End
      Begin MSComCtl2.DTPicker dtpToday_1 
         Height          =   345
         Left            =   3600
         TabIndex        =   22
         Top             =   210
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21430273
         CurrentDate     =   40534
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4185
         TabIndex        =   1
         Top             =   270
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   705
      Left            =   17190
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin FPSpread.vaSpread vasResTemp 
      Height          =   2775
      Left            =   15030
      TabIndex        =   9
      Top             =   2250
      Visible         =   0   'False
      Width           =   3375
      _Version        =   393216
      _ExtentX        =   5953
      _ExtentY        =   4895
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
      SpreadDesigner  =   "frmInterface.frx":030A
   End
   Begin FPSpread.vaSpread vasCode 
      Height          =   2805
      Left            =   15030
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
      _Version        =   393216
      _ExtentX        =   3625
      _ExtentY        =   4948
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
      SpreadDesigner  =   "frmInterface.frx":053F
   End
   Begin FPSpread.vaSpread vasTemp1 
      Height          =   1125
      Left            =   15030
      TabIndex        =   6
      Top             =   7875
      Visible         =   0   'False
      Width           =   1755
      _Version        =   393216
      _ExtentX        =   3096
      _ExtentY        =   1984
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
      SpreadDesigner  =   "frmInterface.frx":4A00
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   975
      Left            =   16830
      TabIndex        =   5
      Top             =   7875
      Visible         =   0   'False
      Width           =   1545
      _Version        =   393216
      _ExtentX        =   2725
      _ExtentY        =   1720
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
      SpreadDesigner  =   "frmInterface.frx":4C35
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   10605
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   6165
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5467
            MinWidth        =   5467
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2011-07-05"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오후 1:34"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "메디메이트 ☎(02)6205-1751"
            TextSave        =   "메디메이트 ☎(02)6205-1751"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame7 
      Height          =   9840
      Left            =   30
      TabIndex        =   3
      Top             =   780
      Width           =   14835
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         ScaleHeight     =   345
         ScaleWidth      =   14415
         TabIndex        =   88
         Top             =   9330
         Width           =   14445
         Begin VB.Label Label7 
            BackStyle       =   0  '투명
            Caption         =   "[장비2]"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4770
            TabIndex        =   92
            Top             =   120
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lblConnect2 
            BackStyle       =   0  '투명
            Caption         =   "연결 대기중."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5700
            TabIndex        =   91
            Top             =   120
            Visible         =   0   'False
            Width           =   3555
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "[장비]"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   90
            TabIndex        =   90
            Top             =   90
            Width           =   735
         End
         Begin VB.Label lblConnect 
            BackStyle       =   0  '투명
            Caption         =   "연결 대기중."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   840
            TabIndex        =   89
            Top             =   90
            Width           =   3555
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "[매칭후결과]"
         Height          =   4725
         Left            =   210
         TabIndex        =   73
         Top             =   4560
         Width           =   14535
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   585
            Left            =   9630
            TabIndex        =   99
            Top             =   2745
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.TextBox Text2 
            Height          =   915
            Left            =   5160
            MultiLine       =   -1  'True
            TabIndex        =   98
            Top             =   2700
            Visible         =   0   'False
            Width           =   4485
         End
         Begin Threed.SSCommand cmdReset 
            Height          =   345
            Left            =   5160
            TabIndex        =   84
            Top             =   210
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   609
            _StockProps     =   78
            Caption         =   "화면초기화"
         End
         Begin Threed.SSCommand cmdCall 
            Height          =   345
            Left            =   2760
            TabIndex        =   83
            Top             =   210
            Width           =   2325
            _Version        =   65536
            _ExtentX        =   4101
            _ExtentY        =   609
            _StockProps     =   78
            Caption         =   "Local Data 불러오기"
         End
         Begin VB.TextBox txtTest 
            Height          =   795
            Left            =   1050
            TabIndex        =   78
            Top             =   6930
            Width           =   4995
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Command5"
            Height          =   735
            Left            =   6090
            TabIndex        =   77
            Top             =   6930
            Width           =   1845
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   870
            TabIndex        =   74
            Top             =   720
            Width           =   195
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   225
            Left            =   300
            TabIndex        =   75
            Top             =   750
            Width           =   465
            _Version        =   65536
            _ExtentX        =   820
            _ExtentY        =   397
            _StockProps     =   15
            Caption         =   "번호"
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
         End
         Begin FPSpread.vaSpread vasExam 
            Height          =   3855
            Left            =   150
            TabIndex        =   76
            Top             =   660
            Width           =   14235
            _Version        =   393216
            _ExtentX        =   25109
            _ExtentY        =   6800
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   14
            MaxRows         =   30
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":4E6A
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   1140
            TabIndex        =   85
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
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
            Format          =   21430273
            CurrentDate     =   40534
         End
         Begin VB.Label Label10 
            Caption         =   "검사일자"
            Height          =   255
            Left            =   210
            TabIndex        =   86
            Top             =   300
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "[장비검사결과]"
         Height          =   4305
         Left            =   8280
         TabIndex        =   61
         Top             =   240
         Width           =   6405
         Begin FPSpread.vaSpread vasResult 
            Height          =   3495
            Left            =   150
            TabIndex        =   62
            Top             =   690
            Width           =   6105
            _Version        =   393216
            _ExtentX        =   10769
            _ExtentY        =   6165
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   8
            MaxRows         =   30
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":5A7F
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   555
            Left            =   6450
            TabIndex        =   63
            Top             =   600
            Visible         =   0   'False
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   979
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
            Begin VB.TextBox txtReqS 
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
               Height          =   285
               Left            =   120
               TabIndex        =   66
               Text            =   "1"
               Top             =   180
               Width           =   795
            End
            Begin VB.TextBox txtReqE 
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
               Height          =   285
               Left            =   1200
               TabIndex        =   65
               Text            =   "1"
               Top             =   180
               Width           =   795
            End
            Begin VB.CommandButton cmd_Req_Res2 
               Caption         =   "장비2"
               Height          =   405
               Left            =   3030
               TabIndex        =   64
               Top             =   120
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label Label2 
               Caption         =   "-"
               Height          =   315
               Left            =   990
               TabIndex        =   67
               Top             =   210
               Width           =   255
            End
         End
         Begin MSComCtl2.DTPicker dtpResOnly 
            Height          =   315
            Left            =   1080
            TabIndex        =   68
            Top             =   270
            Width           =   1515
            _ExtentX        =   2672
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
            Format          =   21430273
            CurrentDate     =   40534
         End
         Begin Threed.SSCommand cmdResSch 
            Height          =   345
            Left            =   2640
            TabIndex        =   70
            Top             =   240
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   609
            _StockProps     =   78
            Caption         =   "결과조회"
         End
         Begin Threed.SSCommand cmd_Req_Res 
            Height          =   345
            Left            =   5220
            TabIndex        =   71
            Top             =   240
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   609
            _StockProps     =   78
            Caption         =   "Import"
         End
         Begin Threed.SSCommand cmdResClear 
            Height          =   345
            Left            =   3840
            TabIndex        =   72
            Top             =   240
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   609
            _StockProps     =   78
            Caption         =   "화면초기화"
         End
         Begin VB.Label Label8 
            Caption         =   "검사일자"
            Height          =   255
            Left            =   150
            TabIndex        =   69
            Top             =   330
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "[처방목록]"
         Height          =   4305
         Left            =   225
         TabIndex        =   59
         Top             =   240
         Width           =   7965
         Begin FPSpread.vaSpread vasList 
            Height          =   3495
            Left            =   180
            TabIndex        =   60
            Top             =   690
            Width           =   7665
            _Version        =   393216
            _ExtentX        =   13520
            _ExtentY        =   6165
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   8
            MaxRows         =   30
            Protect         =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":6410
         End
         Begin MSComCtl2.DTPicker dtpSDeptDate 
            Height          =   315
            Left            =   1170
            TabIndex        =   93
            Top             =   210
            Width           =   1605
            _ExtentX        =   2831
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
            Format          =   21430273
            CurrentDate     =   40534
         End
         Begin MSComCtl2.DTPicker dtpEDeptDate 
            Height          =   315
            Left            =   3030
            TabIndex        =   94
            Top             =   210
            Width           =   1605
            _ExtentX        =   2831
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
            Format          =   21430273
            CurrentDate     =   40534
         End
         Begin Threed.SSCommand cmdListSch 
            Height          =   345
            Left            =   4860
            TabIndex        =   95
            Top             =   210
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   609
            _StockProps     =   78
            Caption         =   "처방목록조회"
         End
         Begin VB.Label Label6 
            Caption         =   "채혈일자"
            Height          =   255
            Left            =   210
            TabIndex        =   97
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "-"
            Height          =   195
            Left            =   2820
            TabIndex        =   96
            Top             =   270
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdResInsert 
         Caption         =   "결과수기입력"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   11640
         TabIndex        =   14
         Top             =   3660
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label lblUser 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   13590
         TabIndex        =   12
         Top             =   8880
         Width           =   1035
      End
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   5850
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   5310
      Width           =   2175
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Left            =   420
      TabIndex        =   8
      Top             =   3450
      Width           =   2055
   End
   Begin Threed.SSPanel spNotTrans 
      Height          =   6315
      Left            =   600
      TabIndex        =   15
      Top             =   9960
      Width           =   6615
      _Version        =   65536
      _ExtentX        =   11668
      _ExtentY        =   11139
      _StockProps     =   15
      BackColor       =   15724527
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread vasServerTemp 
         Height          =   1215
         Left            =   1740
         TabIndex        =   21
         Top             =   3120
         Visible         =   0   'False
         Width           =   3225
         _Version        =   393216
         _ExtentX        =   5689
         _ExtentY        =   2143
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
         SpreadDesigner  =   "frmInterface.frx":6DEC
      End
      Begin VB.CommandButton cmdNotExit 
         Caption         =   "종료"
         Height          =   375
         Left            =   5070
         TabIndex        =   20
         Top             =   150
         Width           =   1305
      End
      Begin FPSpread.vaSpread vasNotResult 
         Height          =   5475
         Left            =   210
         TabIndex        =   19
         Top             =   600
         Width           =   6195
         _Version        =   393216
         _ExtentX        =   10927
         _ExtentY        =   9657
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
         MaxCols         =   5
         MaxRows         =   30
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":7021
      End
      Begin VB.CommandButton cmdNotResult 
         Caption         =   "조회"
         Height          =   375
         Left            =   3270
         TabIndex        =   18
         Top             =   150
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpNotDate 
         Height          =   315
         Left            =   1500
         TabIndex        =   16
         Top             =   180
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21430273
         CurrentDate     =   40589
      End
      Begin VB.Label Label3 
         Caption         =   "검사일자"
         Height          =   195
         Left            =   540
         TabIndex        =   17
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3015
      Left            =   -30
      TabIndex        =   26
      Top             =   10200
      Visible         =   0   'False
      Width           =   9705
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   525
         Left            =   4260
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   41
         Top             =   270
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "WorkList"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Picture         =   "frmInterface.frx":77FC
         Style           =   1  '그래픽
         TabIndex        =   29
         Top             =   750
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmdResSave 
         Caption         =   "결과저장"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   90
         TabIndex        =   28
         Top             =   270
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CommandButton cmdQC 
         Caption         =   "QC"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2130
         TabIndex        =   27
         Top             =   270
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox Text_Today 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   420
         TabIndex        =   30
         Text            =   "2002/02/18"
         Top             =   120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   1455
         Left            =   1710
         TabIndex        =   37
         Top             =   1200
         Visible         =   0   'False
         Width           =   3315
         _Version        =   393216
         _ExtentX        =   5847
         _ExtentY        =   2566
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
         MaxCols         =   17
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmInterface.frx":80C6
         UserResize      =   2
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   1455
         Left            =   5100
         TabIndex        =   38
         Top             =   1200
         Visible         =   0   'False
         Width           =   2655
         _Version        =   393216
         _ExtentX        =   4683
         _ExtentY        =   2566
         _StockProps     =   64
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   16777215
         MaxCols         =   8
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":E1AE
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   3495
      Left            =   -420
      TabIndex        =   31
      Top             =   10560
      Visible         =   0   'False
      Width           =   6435
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   585
         Left            =   2550
         TabIndex        =   40
         Top             =   2820
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   240
         TabIndex        =   39
         Top             =   2820
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.TextBox txtBuff 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   36
         Top             =   2310
         Visible         =   0   'False
         Width           =   4770
      End
      Begin VB.TextBox txtBuff2 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   35
         Top             =   1830
         Visible         =   0   'False
         Width           =   4770
      End
      Begin FPSpread.vaSpread vas_print 
         Height          =   1395
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Visible         =   0   'False
         Width           =   1995
         _Version        =   393216
         _ExtentX        =   3519
         _ExtentY        =   2461
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
         MaxCols         =   8
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":11FCA
      End
      Begin FPSpread.vaSpread vasOrder 
         Height          =   1395
         Left            =   2190
         TabIndex        =   33
         Top             =   360
         Visible         =   0   'False
         Width           =   1995
         _Version        =   393216
         _ExtentX        =   3519
         _ExtentY        =   2461
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
         MaxCols         =   3
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":1392B
      End
      Begin FPSpread.vaSpread vasTransChk 
         Height          =   1395
         Left            =   4230
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
         _Version        =   393216
         _ExtentX        =   3625
         _ExtentY        =   2461
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
         SpreadDesigner  =   "frmInterface.frx":15166
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   6975
      Left            =   -150
      TabIndex        =   42
      Top             =   10440
      Visible         =   0   'False
      Width           =   7425
      Begin VB.TextBox Text_ini 
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
         Left            =   0
         TabIndex        =   58
         Top             =   0
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton cmdUp 
         Height          =   525
         Left            =   0
         Picture         =   "frmInterface.frx":1539B
         Style           =   1  '그래픽
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton cmdDown 
         Height          =   525
         Left            =   1170
         Picture         =   "frmInterface.frx":154CA
         Style           =   1  '그래픽
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2490
         Picture         =   "frmInterface.frx":155FC
         ScaleHeight     =   285
         ScaleWidth      =   315
         TabIndex        =   55
         Top             =   60
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton Command4 
         Caption         =   "test"
         Height          =   615
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton cmdUser 
         Caption         =   "사용자관리"
         Height          =   465
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Width           =   1665
      End
      Begin VB.TextBox txtUser 
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
         Height          =   315
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "결과 출력"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   50
         Top             =   5160
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton chkStart 
         Caption         =   "시작"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   49
         Top             =   5160
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox txtBarcode 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   390
         TabIndex        =   48
         Top             =   5220
         Visible         =   0   'False
         Width           =   2385
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   855
         Left            =   480
         TabIndex        =   43
         Top             =   450
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "미등록건수"
         ForeColor       =   16777215
         BackColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   735
         Left            =   540
         TabIndex        =   44
         Top             =   1650
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "어제"
         ForeColor       =   16777215
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel spMissResPast 
         Height          =   2475
         Left            =   480
         TabIndex        =   45
         Top             =   2610
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   4366
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   16777215
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   48
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   735
         Left            =   3690
         TabIndex        =   46
         Top             =   1650
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "오늘"
         ForeColor       =   16777215
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel spMissResNow 
         Height          =   2475
         Left            =   3630
         TabIndex        =   47
         Top             =   2580
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   4366
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   16777215
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   48
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "보고자 :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin Threed.SSCommand cmdSend 
      Height          =   615
      Left            =   10290
      TabIndex        =   87
      Top             =   120
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "결과전송"
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
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===========================================

'vasList
'barcode, pid, kornm, bcno,  deptname
Const colCheckBox = 1
Const colBarcode = 2
Const colPID = 3
Const colPName = 4
Const colBCNO = 5
Const colDept = 6
Const colLSampNo = 7
Const colReady = 7
Const colLExamDate = 8
Const colExamName = 8
Const colLExamTime = 9
Const colLEquipCode = 10
Const colLExamCode = 11
Const colLExamName = 12
Const colLResult = 13
Const colState = 14


'vasResult

'SampNo , ExamDate, ExamTime, EquipCode, ExamCode, ExamName, Result

Const ColSampNo = 2
Const ColExamDate = 3
Const ColExamTime = 4
Const ColEquipCode = 5
Const ColExamCode = 6
Const colRExamName = 7
Const ColResult = 8


'vasID
''', 검체번호, 검사일자, 검사시간, 장비코드, 검사코드, 검사명, 장비결과, 결과, 상태, 비고
'''
'''Const colCheckBox = 1
'''Const colExamDate = 3
'''Const colExamTime = 4
'''Const colEquipCode = 5
'''Const colExamCode = 6
'''Const colExamName = 7
'''Const colEquipRes = 8
'''Const colResult = 9
'''Const colState = 10
'''Const colErrState = 11


Dim colResult1 As Long

Dim gRow As Long

Dim gsBarCode As String
Dim gsPID As String
Dim gsRackNo As String
Dim gsPosNo As String
Dim gsResDateTime As String
Dim gsSeqNo As String
Dim gsExamCode As String
Dim gsExamName As String
Dim gsOrder As String
Dim gsResult As String

Dim PreRack As String
Dim PrePos As String
Dim PreRow As Long

'************************************************
Dim in_spc_no$, spc_no$(), TST_CD$(), TST_NM$()
Dim spc_cd$(), TST_FRCT_CD$(), tst_frct_nm$()
Dim tst_dte$(), tst_time$(), work_no$()
Dim PT_NO$(), PT_NM$(), sex$(), birthday$(), intbase$()

Dim acpt_no$()

Dim rv As Integer
Dim vTemp As String
'************************************************

Function SetResult(asResult As String, aiItem As Integer) As String
    Dim iFloat As Integer

    If Not IsNumeric(asResult) Then
        Exit Function
    End If

    Select Case aiItem
'    Case 1, 14, 15, 16, 17, 18, 30
'        iFloat = 2
    Case 8
        iFloat = 0
    Case 1, 2, 14, 15, 16, 17, 18, 24, 30
        iFloat = 2
    Case Else
        iFloat = 1
    End Select

    If iFloat = 0 Then
        SetResult = CStr(CCur(asResult))
    Else
        If aiItem = 1 Then
            SetResult = CStr(Format(CCur(Left(asResult, 5 - iFloat)) & "." & Right(asResult, iFloat), "#0.0"))
        ElseIf aiItem = 14 Or aiItem = 15 Or aiItem = 16 Or aiItem = 17 Or aiItem = 18 Then
            SetResult = CStr(CCur(Left(asResult, 5 - iFloat)) & "." & Right(asResult, iFloat))
        Else
            SetResult = CStr(CCur(Left(asResult, 4 - iFloat)) & "." & Right(asResult, iFloat))
        End If
    End If
    
        
'사용안함
'    If aiItem = 1 Then
'        SetResult = Format(SetResult, "#0.0")
'    End If
End Function

Public Function Check_Result(argBarCode As String, argPID As String, argExamCode As String, _
                            argResult As String, ByVal argRow As Integer, asSex As String) As Integer
    Dim sDiffRet, sDiffRet1 As String
    Dim PreResult   As String
    
    Dim sResClassCode As String     '결과종류
    Dim sLow        As String       '참조치
    Dim sHigh       As String
    Dim RefRet      As String
    Dim sPanicGubun As String
    Dim sPanicLow   As String       'Panic
    Dim sPanicHigh  As String
    Dim PanicRet    As String
    Dim sDeltaGubun As String
    Dim sDeltaLow   As String       'Delta
    Dim sDeltaHigh  As String
    Dim DeltaRet    As String
    
    Dim sTmpRece1, sTmpRet1 As String
    Dim sTmpRece2, sTmpRet2 As String
    Dim sMax_ReceNo As String
    Dim i           As Integer
    Dim sReceNo     As String
    Dim sPID        As String
    
    Dim sTmpStr As String
    
    Check_Result = -1
    
    If argBarCode = "" Then
        Exit Function
    End If
    
    If argExamCode = "" Then
        Exit Function
    End If
    

    RefRet = ""
    PanicRet = ""
    DeltaRet = ""
    
    sDiffRet = argResult
    If sDiffRet = "" Then
        Check_Result = -1
        Exit Function
    End If
    
    SQL = " Use NeoSoft"
    res = SendQuery(gServer, SQL)
    
    SQL = " Select LABM_MAN_FRES, LABM_MAN_TRES, LABM_WOM_FRES, LABM_WOM_TRES " & CR & _
          "From CC_LABM " & CR & _
          " Where LABN_ID = '" & Trim(argExamCode) & "' "
          
    res = db_select_Col(gServer, SQL)
    
'    sResClassCode = Trim(gReadBuf(0))
    
'    If sResClassCode = "1" Then '숫자
'참조치 체크
        sLow = ""
        sHigh = ""
        
        '숫자인지 아닌지 확인
        If IsNumeric(sDiffRet) = False Then
           'MsgBox "결과형식이 일치하지 않습니다.", vbInformation, "알림"
           Check_Result = -1
           Exit Function
        End If
        
'        If IsNumeric(gReadBuf(13)) Then
'            If CInt(gReadBuf(13)) > 0 Then
'                sTmpStr = "#0."
'                For i = 1 To CInt(gReadBuf(13))
'                    sTmpStr = sTmpStr & "0"
'                Next i
'            Else
'                sTmpStr = "#0"
'            End If
'            sDiffRet = Format(sDiffRet, sTmpStr)
'            SetText vasRes, sDiffRet, argRow, colResult
'            SetText vasRes, sDiffRet, argRow, colResult1
'        End If
        
        Select Case asSex
        Case "M", ""
            sLow = Trim(gReadBuf(0))
            sHigh = Trim(gReadBuf(1))
        Case "F"
            sLow = Trim(gReadBuf(2))
            sHigh = Trim(gReadBuf(3))
        End Select
        
        If sLow = "" And sHigh = "" Then
            RefRet = ""
        ElseIf sLow = "" And sHigh <> "" Then   '이상
            If CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            End If
        ElseIf sLow <> "" And sHigh = "" Then   '이하
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            End If
        Else
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            ElseIf CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            ElseIf CCur(sLow) <= CCur(sDiffRet) And CCur(sHigh) <= CCur(sDiffRet) Then
                RefRet = ""
            End If
        End If


''Panic 체크
'        sPanicLow = ""
'        sPanicHigh = ""
'
'        sPanicGubun = Trim(gReadBuf(5))
'
'        Select Case asSex
'        Case "M", ""
'            sPanicLow = Trim(gReadBuf(6))
'            sPanicHigh = Trim(gReadBuf(7))
'        Case "F"
'            sPanicLow = Trim(gReadBuf(8))
'            sPanicHigh = Trim(gReadBuf(9))
'        End Select
'
'        If sPanicGubun = "0" Then '상한/하한
'            If sPanicLow = "" Or sPanicHigh = "" Then
'                PanicRet = ""
'            Else
'                If CCur(sPanicLow) > CCur(sDiffRet) Then
'                    PanicRet = "L"
'                ElseIf CCur(sPanicHigh) < CCur(sDiffRet) Then
'                    PanicRet = "H"
'                ElseIf CCur(sPanicLow) <= CCur(sDiffRet) And CCur(sPanicHigh) <= CCur(sDiffRet) Then
'                    PanicRet = ""
'                End If
'            End If
'        ElseIf sPanicGubun = "1" Then 'percent
'            If sPanicLow = "" Then
'                PanicRet = ""
'            Else
'                If CCur(sPanicLow) - CCur(sDiffRet) > 0 Then
'                    If ((CCur(sPanicLow) - CCur(sDiffRet)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
'                        PanicRet = "L"
'                    Else
'                        PanicRet = ""
'                    End If
'                ElseIf CCur(sPanicHigh) - CCur(sDiffRet) < 0 Then
'                    If ((CCur(sDiffRet) - CCur(sPanicLow)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
'                        PanicRet = "H"
'                    Else
'                        PanicRet = ""
'                    End If
'                Else
'                    PanicRet = ""
'                End If
'            End If
'        End If
'
'
''Delta 체크
'        sDeltaLow = ""
'        sDeltaHigh = ""
'
'        sTmpRece1 = ""
'        sTmpRet1 = ""
'        sTmpRece2 = ""
'        sTmpRet2 = ""
'        PreResult = ""
'
'        sMax_ReceNo = ""
''        sTmpRece1 = Trim(argForm.dtpReceDate.Value)
'        sReceNo = argBarCode
'
'        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
'              " Where HID = '115' " & CR & _
'              " And PID = '" & Trim(argPID) & "' " & CR & _
'              " And ReceNo < '" & argBarCode & "' " & CR & _
'              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
'              " Group By Result"
'
'        res = db_select_Col(gServer, SQL)
'
'        If res > 0 And gReadBuf(0) <> "" Then
'            PreResult = gReadBuf(0)
'        Else
'            PreResult = ""
'        End If
'
'        If PreResult <> "" Then
'          'PreResult = Trim(gReadBuf(0))
'          sDeltaGubun = Trim(gReadBuf(10))
'
'          sDeltaLow = Trim(gReadBuf(11))
'          sDeltaHigh = Trim(gReadBuf(12))
'
'            '이전결과에서 현재결과 뺀값이 sDiffRet임 (2002년 3월 15일 수정)
''            sDiffRet = PreResult - sDiffRet
'            sDiffRet1 = sDiffRet - PreResult
'            If sDeltaGubun = "0" Then '상한/하한
'                If sDeltaLow = "" Or sDeltaHigh = "" Then
'                    DeltaRet = ""
'                Else
'                    If CCur(sDeltaLow) > CCur(sDiffRet1) Then
'                        DeltaRet = "L"
'                    ElseIf CCur(sDeltaHigh) < CCur(sDiffRet1) Then
'                        DeltaRet = "H"
'                    ElseIf CCur(sDeltaLow) <= CCur(sDiffRet1) And CCur(sDeltaHigh) <= CCur(sDiffRet1) Then
'                        DeltaRet = ""
'                    End If
'                End If
'
'            ElseIf sDeltaGubun = "1" Then 'percent
'               If CInt(PreResult) = 0 Or CInt(sDiffRet) = 0 Then
'                  DeltaRet = ""
'               Else
'                   If sDeltaLow = "" Then
'                        DeltaRet = ""
'                    Else
'                        If (Abs(CCur(PreResult) - CCur(sDiffRet)) / CCur(PreResult)) * 100 >= CCur(sDeltaLow) Then
'                            DeltaRet = "D"
'                        Else
'                            DeltaRet = ""
'                        End If
'                    End If
'               End If
'            End If
'        End If
'
'    ElseIf sResClassCode = "2" Then '문자
'        Dim sRefValue As String
'        Dim sPanicValue As String
'        Dim sResult As String
'
'        sLow = ""
'        sLow = UCase(Trim(GetText(argTable, argRow, iresRefValue)))
'
'        '2003/03/17 이상은 수정
'        '검사 항목 결과 참조 코드 체크에서 1 이상일 경우만 판정되게
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            Exit Function
'        End If
'
'        '2002년 3월 12일 +-에서 +/-로 수정
'        '2002년 5월 13일 NON-REACTIVE 판정 안돼서 추가
'        '2003년 2월 4일 이상은 수정 - 0-1로 참조치는 1이나 판정됨
'        '=================================================================================
'        '2002년 5월 13일 1 : 40 미만 판정 안됨
'        '2002년 6월 11일 (결과참조가 1:로 시작하면 판정체크 안하게 수정)
'        If Trim(Left(sDiffRet, 3)) = "1 :" Or Trim(Left(sDiffRet, 3)) = "1:" Then
'            Exit Function
'        End If
'        '=================================================================================
'
'        Select Case UCase(sDiffRet)
'        Case "-", "NEGATIVE", "음성", "1", "NON-REACTIVE", "0-1"
'            sResult = 1
'        Case "+/-", "2", "+-", "2-5"
'            sResult = 2
'        Case "+", "POSITIVE", "양성", "3", "6-10"
'            sResult = 3
'        Case "++", "4", "11-20"
'            sResult = 4
'        Case "+++", "5", "21-30"
'            sResult = 5
'        Case "++++", "6"
'            sResult = 6
'        Case "+++++", "7"
'            sResult = 7
'        Case "++++++", "8"
'            sResult = 8
'        Case Else
'            sResult = sDiffRet
'        End Select
'        'sLow = "0-2"
'        If Trim(sLow) <> "" Then
'            Select Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "음성", "1", "NON-REACTIVE", "0-2"
'                sRefValue = 1
'            Case "+/-", "2", "+-"
'                sRefValue = 2
'            Case "+", "POSITIVE", "양성", "3"
'                sRefValue = 3
'            Case "++", "4"
'                sRefValue = 4
'            Case "+++", "5"
'                sRefValue = 5
'            Case "++++", "6"
'                sRefValue = 6
'            Case "+++++", "7"
'                sRefValue = 7
'            Case "++++++", "8"
'                sRefValue = 8
'            Case Else
'                If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'                    RefRet = Trim(GetText(argTable, argRow, iresDecision))
'                ElseIf UCase(sDiffRet) <> UCase(sLow) Then
'                    RefRet = sDiffRet
'                End If
'            End Select
'            If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'
'            ElseIf sRefValue < sResult Then
''                RefRet = "H"
'                RefRet = sDiffRet
'
''                argTable.Row = argRow
''                argTable.Col = iresDecision
''                argTable.ForeColor = RGB(205, 55, 0)
'
'
'            End If
'        End If
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            RefRet = Trim(GetText(argTable, argRow, iresDecision))
'        End If
'        sLow = ""
'        sLow = Trim(GetText(argTable, argRow, iresPanicValue))
'        If Trim(sLow) <> "" Then
'            Select Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "음성"
'                sPanicValue = 1
'            Case "+/-"
'                sPanicValue = 2
'            Case "+", "POSITIVE", "양성"
'                sPanicValue = 3
'            Case "++"
'                sPanicValue = 4
'            Case "+++"
'                sPanicValue = 5
'            Case "++++"
'                sPanicValue = 6
'            Case "+++++"
'                sPanicValue = 7
'            Case "++++++"
'                sPanicValue = 8
'            Case Else
'                If UCase(sDiffRet) > UCase(sLow) Then
'                    PanicRet = sDiffRet
'                End If
'            End Select
'            If sPanicValue < sResult Then
'                'PanicRet = "H"
'                PanicRet = sDiffRet
'            End If
'        End If
'
'        'Delta Check
'        sMax_ReceNo = ""
'        DeltaRet = ""
'        sReceNo = Trim(GetText(argForm.vasPatient, 1, 1))
'        sPID = Trim(GetText(argForm.vasPatient, 1, 3))
'
'        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
'              " Where PID = '" & sPID & "' " & CR & _
'              " And ReceNo < '" & sReceNo & "' " & CR & _
'              " And ExamCode = '" & Trim(GetText(argTable, argRow, iresExamCode)) & "' " & CR & _
'              " Group By Result"
'
'        res = db_select_Col(SQL)
'
'        If res > 0 And gReadBuf(0) <> "" Then
'               If sDiffRet <> gReadBuf(0) Then
'                  DeltaRet = "D"
'               End If
'        Else
'            DeltaRet = ""
'        End If
'    End If
'
'    SetText vasRes, RefRet, argRow, colRCheck
'
'
'    '2002년 2월 15일 수정 (판정시 H, L 일때 글자 색깔 변화)
'    '2002년 3월 14일 수정 (판정시 L일때는 파란색 그 외는 빨간색)
'    If RefRet = "L" Then
'        vasRes.Row = argRow
'        vasRes.Col = colRCheck
'        vasRes.ForeColor = RGB(65, 105, 225)
'    Else
'        vasRes.Row = argRow
'        vasRes.Col = colRCheck
'        vasRes.ForeColor = RGB(205, 55, 0)
'    End If
'
'
'    Check_Result = 1

End Function

Private Sub chkAll_Click()
    Dim iRow As Integer
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasExam.DataRowCnt
            vasExam.Row = iRow
            vasExam.Col = 1
            
            vasExam.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasExam.DataRowCnt
            vasExam.Row = iRow
            vasExam.Col = 1
            
            vasExam.Value = 0
        Next iRow
    End If

End Sub

Private Sub chkMode_Click()
    If chkMode.Value = 1 Then
        chkMode.Caption = "Auto"
    Else
        chkMode.Caption = "Manual"
    End If
End Sub

Private Sub cmd_print_Click()
'''    Dim iRow As Integer
'''    Dim j As Integer
'''
'''    Dim sCurDate As String
'''    Dim sSerDate As String
'''    Dim sHead As String
'''    Dim sFoot As String
'''
'''    ClearSpread vas_print
'''
'''    j = 1
'''
'''    For iRow = 1 To vasExam.DataRowCnt
'''        vasExam.Row = iRow
'''        vasExam.Col = 1
'''
'''        If vasExam.Value = 1 Then
'''
'''            SetText vas_print, Trim(GetText(vasExam, iRow, 2)), j, 1
'''            SetText vas_print, Trim(GetText(vasExam, iRow, 3)), j, 2
'''            SetText vas_print, Trim(GetText(vasExam, iRow, 4)), j, 3
'''            SetText vas_print, Trim(GetText(vasExam, iRow, 5)), j, 4
'''            SetText vas_print, Trim(GetText(vasExam, iRow, 6)), j, 5
'''            SetText vas_print, Trim(GetText(vasExam, iRow, colResult)), j, 6
'''            SetText vas_print, Trim(GetText(vasExam, iRow, colResult + 4)), j, 7
'''            SetText vas_print, Trim(GetText(vasExam, iRow, vasExam.MaxCols)), j, 8
''''            SetText vas_print, Trim(GetText(vasExam, iRow, 10)), j, 9
'''
'''            j = j + 1
'''        End If
'''    Next iRow
'''
'''    If vas_print.DataRowCnt < 1 Then
'''        MsgBox "출력할 자료가 없습니다.", , "알 림"
'''        Exit Sub
'''    End If
'''
'''    SetText vas_print, Trim(GetText(vasExam, 0, 2)), 0, 1
'''    SetText vas_print, Trim(GetText(vasExam, 0, 3)), 0, 2
'''    SetText vas_print, Trim(GetText(vasExam, 0, 4)), 0, 3
'''    SetText vas_print, Trim(GetText(vasExam, 0, 5)), 0, 4
'''    SetText vas_print, Trim(GetText(vasExam, 0, 6)), 0, 5
'''    SetText vas_print, Trim(GetText(vasExam, 0, colResult)), 0, 6
'''    SetText vas_print, Trim(GetText(vasExam, 0, colResult + 4)), 0, 7
'''    SetText vas_print, Trim(GetText(vasExam, 0, vasExam.MaxCols)), 0, 8
'''
'''
'''    sCurDate = Text_Today
'''
'''    sSerDate = Text_Today
'''
'''    vas_print.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
'''    vas_print.PrintAbortMsg = "인쇄중 입니다 ..."
'''    vas_print.PrintJobName = "TaqMan WorkList 출력"
'''
'''    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & " 일산백병원 진단검사의학과"
'''
'''    vas_print.PrintHeader = sHead
'''    vas_print.PrintFooter = sFoot
'''
'''    vas_print.PrintMarginTop = 680
'''    vas_print.PrintMarginBottom = 680
''''현재 SS가 비대칭으로 출력함
''''    vaslist.PrintMarginLeft = 720
'''    vas_print.PrintMarginLeft = 0
'''    vas_print.PrintMarginRight = 0
'''
'''    vas_print.PrintColor = True
'''    vas_print.PrintGrid = True
'''
''''Set printing range
'''    vas_print.PrintType = 0  'SS_PRINT_ALL(default)
'''
'''    vas_print.PrintShadows = True
'''
'''    vas_print.Action = 13 'SS_ACTION_PRINT

End Sub

Private Sub cmd_Req_Res_Click()

    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    txtReqS.Text = "1"
    txtReqE.Text = "5"
    
    lblConnect.Caption = "연결 대기중."
    lblConnect.ForeColor = &HFF&
    MSComm1.Output = H232_Function(H232_Connect)
    Save_Raw_Data "[Tx1 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & H232_Function(H232_Connect)
    H232_Connect_state = False
    tm_H232.Enabled = True
    
End Sub

Private Sub cmd_Req_Res2_Click()

    If MSComm2.PortOpen = False Then
        MSComm2.PortOpen = True
    End If
    lblConnect2.Caption = "연결 대기중."
    lblConnect2.ForeColor = &HFF&
    MSComm2.Output = H232_Function_2(H232_Connect)
    Save_Raw_Data "[Tx2 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & H232_Function_2(H232_Connect)
    H232_Connect_state_2 = False
    tm_H232_2.Enabled = True
End Sub

'Private Sub chkStart_Click()
'    If Timer1.Enabled = True Then
'        Timer1.Enabled = False
'        chkStart.Caption = "시작"
'    Else
'        gWait = 0
'        gRow = 0
'        Timer1.Enabled = True
'        chkStart.Caption = "종료"
'    End If
'End Sub

Private Sub cmdCall_Click()
    Dim iRow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Y As Integer
    Dim sResFlag As String
    Dim sRes As String
    
    Dim sResult As String
    
    ClearSpread vasExam
    
   ''', 검체번호, 검사일자, 검사시간, 장비코드, 검사코드, 검사명, 장비결과, 결과, 상태, 비고
    SQL = "select barcode, pid, pname, bcno, dept, sampno, examdate, examtime, equipcode, examcode, examname, result, sendflag " & _
          "from pat_res " & _
          "where examdate = '" & Format(dtpExamDate, "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          "group by barcode, pid, pname, bcno, dept, sampno, examdate, examtime, equipcode, examcode, examname, result, sendflag "
    res = db_select_Vas(gLocal, SQL, vasExam, vasExam.DataRowCnt + 1, 2)
    
'''    ClearSpread vaSpread1
'''
'''    res = db_select_Vas(gLocal, SQL, vaSpread1)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
'''    vasSort vasExam, colRack, colTube
    
    For iRow = 1 To vasExam.DataRowCnt
        
        Select Case Trim(GetText(vasExam, iRow, colState))
        Case "0"
            SetBackColor vasExam, iRow, iRow, colCheckBox, colState, 255, 250, 205
            SetText vasExam, "결과", iRow, colState
        Case "1"
            SetBackColor vasExam, iRow, iRow, 1, colState, 202, 255, 112
            SetText vasExam, "완료", iRow, colState
        Case Else
            SetBackColor vasExam, iRow, iRow, 1, colState, 255, 255, 255
            SetText vasExam, "", iRow, colState
        End Select
    
        '결과 불러오기
        ClearSpread vasTemp
        
    Next iRow
    vasExam.RowHeight(-1) = 14
'''    TransCheck
'''    vasExam.RowHeight(-1) = 17
End Sub

Private Sub cmdCol1_Click()
'''    If cmdCol1.Caption = "<" Then
'''        cmdCol1.Caption = ">"
'''        vasExam.Col = 4
'''        vasExam.ColHidden = True
'''        vasExam.Col = 5
'''        vasExam.ColHidden = True
'''        vasExam.Col = 11
'''        vasExam.ColHidden = True
'''        vasExam.Col = 13
'''        vasExam.ColHidden = True
'''
'''    Else
'''        cmdCol1.Caption = "<"
'''        vasExam.Col = 4
'''        vasExam.ColHidden = False
'''        vasExam.Col = 5
'''        vasExam.ColHidden = False
'''        vasExam.Col = 11
'''        vasExam.ColHidden = False
'''        vasExam.Col = 13
'''        vasExam.ColHidden = False
'''    End If
    
End Sub

Private Sub cmdDown_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow + 1
    vasActiveCell vasID, lRow + 1, 2
    vasID_Click 2, lRow + 1
End Sub

Private Sub cmdListSch_Click()
    Dim sWorkListURL As String
    Dim i As Long
    
'''http://10.90.10.228:8090/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetCoaguChekHJList&Data1=&Data2=
    sWorkListURL = "http://10.20.200.1/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetCoaguChekHJList&Data1=" & Format(dtpSDeptDate.Value, "yyyymmdd") & "&Data2=" & Format(dtpEDeptDate.Value, "yyyymmdd")
'''
'''    sWorkListURL = "http://10.90.10.228:8090/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetCoaguChekHJList&Data1=&Data2="
    URLstart sWorkListURL
    
    
    
'''    ControlURL1.Start "D:\인터페이스_진행\서울건국대학교병원\GyeolGwaIF.xml"
    Save_Raw_Data "[XML]" & gStrXML
    
    
    
    i = InStr(1, Trim(gStrXML), "</root>")
    If i > 0 Then
        gStrXML = Mid(Trim(gStrXML), 1, i + 6)
    End If
    
    Online_Param gStrXML
    
    ClearSpread vasList
    
    If vasList.MaxRows < giIndex + 1 Then
        vasList.MaxRows = giIndex + 1
    End If
    
    Call HJList_Parsing
    
    For i = 0 To giIndex
        SetText vasList, gWorkList(i).pid, i + 1, colPID
        SetText vasList, gWorkList(i).barcode, i + 1, colBarcode
        SetText vasList, gWorkList(i).bcno, i + 1, colBCNO
        SetText vasList, gWorkList(i).kornm, i + 1, colPName
        SetText vasList, gWorkList(i).deptname, i + 1, colDept
        SetText vasList, gWorkList(i).ExamName, i + 1, colExamName
    Next i

    vasList.MaxRows = vasList.DataRowCnt
    If vasList.DataRowCnt = 1 Then
        Call vasList_DblClick(3, 1)
    End If
    vasList.RowHeight(-1) = 14
End Sub

Private Function HJList_Parsing()
    Dim i As Integer
    Dim pid_pos As Integer
    Dim barcode_pos As Integer
    Dim bcno_pos As Integer
    Dim kornm_pos As Integer
    Dim deptname_pos As Integer
    
    For i = 0 To giIndex
        pid_pos = InStr(1, gWorkList(i).HJList, ",")
        kornm_pos = InStr(pid_pos + 1, gWorkList(i).HJList, ",")
        bcno_pos = InStr(kornm_pos + 1, gWorkList(i).HJList, ",")
        barcode_pos = InStr(bcno_pos + 1, gWorkList(i).HJList, ",")
        deptname_pos = InStr(barcode_pos + 1, gWorkList(i).HJList, ",")
            
        gWorkList(i).pid = Mid(gWorkList(i).HJList, 1, pid_pos - 1)
        gWorkList(i).kornm = Mid(gWorkList(i).HJList, pid_pos + 1, kornm_pos - pid_pos - 1)
        gWorkList(i).bcno = Mid(gWorkList(i).HJList, kornm_pos + 1, bcno_pos - kornm_pos - 1)
        gWorkList(i).barcode = Mid(gWorkList(i).HJList, bcno_pos + 1, barcode_pos - bcno_pos - 1)
        gWorkList(i).deptname = Mid(gWorkList(i).HJList, barcode_pos + 1, deptname_pos - barcode_pos - 1)
        gWorkList(i).ExamName = Mid(gWorkList(i).HJList, deptname_pos + 1)
    Next i
    
End Function



Private Sub cmdNotExit_Click()
    spNotTrans.Visible = False
End Sub

Private Sub cmdNotResult_Click()
    ClearSpread vasNotResult
    
    SQL = "select '', barcode, examcode, examname, result from pat_res " & vbCrLf & _
          "where examdate = '" & Format(dtpNotDate.Value, "yyyymmdd") & "' and sendflag <> 'C'"
    res = db_select_Vas(gLocal, SQL, vasNotResult)
    
End Sub

Private Sub cmdQC_Click()
'''    frmQCResSch.Show
End Sub

Private Sub cmdResClear_Click()
    ClearSpread vasResult
    
End Sub

Private Sub cmdReset_Click()
    Dim i As Integer

    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
'''    vasID.Row = 1
'''    vasID.Row2 = vasID.MaxRows
'''    vasID.Col = 1
'''    vasID.Col2 = vasID.MaxCols
'''    vasID.BlockMode = True
'''    vasID.BackColor = RGB(255, 255, 255)
'''    vasID.Action = 3
'''    vasID.BlockMode = False

'''    ClearSpread vasID
'''    ClearSpread vasRes
    ClearSpread vasExam
    
    dtpExamDate = Date
    dtpToday = Format(Date, "yyyy/mm/dd")
'''    TransCheck
    gRow = 0
End Sub

Private Sub cmdResInsert_Click()
    Dim iRow As Long
    Dim iCol As Long
    Dim lsNewBarcode As String
    Dim lsOldBarcode As String
    Dim lsBarcode As String
    
    Dim rv As Integer
    Dim i As Long
    
    iRow = vasExam.DataRowCnt + 1
    
    If iRow > vasExam.MaxRows Then
        vasExam.MaxRows = iRow
    End If
    
    lsNewBarcode = InputBox("변경할 검체번호를 입력하세요.", "검체번호변경")
    
    lsBarcode = Left(lsNewBarcode, 11)
    SQL = "p_interfacequery '1', '" & lsBarcode & "'"
    res = db_select_Col(gServer, SQL)
    
    If res < 1 Then
    Else
    
        SetText vasExam, lsBarcode, iRow, colBarcode
        
        SetText vasExam, gReadBuf(0), iRow, colPID
'''        SetText vasExam, gReadBuf(9), iRow, colReceDate
'''        SetText vasExam, gReadBuf(10), iRow, colReceno
'''        SetText vasExam, gReadBuf(11), iRow, colSeqNo
        SetText vasExam, gReadBuf(32), iRow, colPName
    End If
    
End Sub

Private Sub cmdResMach_Click()
'''체크박스, 바코드번호, 등록번호, 성명, 검체번호, 진료과, 검사번호, 검사일자, 검사시간, 장비코드, 검사코드, 검사명, 결과
    Dim i As Long
    Dim iList As Long
    Dim iRes As Long
    Dim j As Long
    Dim iRow As Long
    Dim sBarcode As String
    Dim sPID As String
    Dim sPName As String
    Dim sBCNO As String
    Dim sDept As String
    Dim sSampNo As String
    Dim sExamDate As String
    Dim sExamTime As String
    Dim sEquipCode As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim sResult As String
    
    iList = -1
    
    For i = 1 To vasList.DataRowCnt
        vasList.Col = 1
        vasList.Row = i
        If vasList.Value = 1 Then
            iList = i
            Exit For
        End If
    Next
    
    iRes = -1
    
    For i = 1 To vasResult.DataRowCnt
        vasResult.Col = 1
        vasResult.Row = i
        If vasResult.Value = 1 Then
            iRes = i
            Exit For
        End If
    Next
    
    If iList = -1 Or iRes = -1 Then
        Exit Sub
    End If
    
    sBarcode = Trim(GetText(vasList, iList, colBarcode))
    sPID = Trim(GetText(vasList, iList, colPID))
    sPName = Trim(GetText(vasList, iList, colPName))
    sBCNO = Trim(GetText(vasList, iList, colBCNO))
    sDept = Trim(GetText(vasList, iList, colDept))
    sSampNo = Trim(GetText(vasResult, iRes, ColSampNo))
    sExamDate = Trim(GetText(vasResult, iRes, ColExamDate))
    sExamTime = Trim(GetText(vasResult, iRes, ColExamTime))
    sEquipCode = Trim(GetText(vasResult, iRes, ColEquipCode))
    sExamCode = Trim(GetText(vasResult, iRes, ColExamCode))
    sExamName = Trim(GetText(vasList, iList, colExamName))
    sResult = Trim(GetText(vasResult, iRes, ColResult))
    
    SQL = "select barcode from pat_res where equipno = '" & gEquip & "' and barcode = '" & sBarcode & "'"
    res = db_select_Col(gLocal, SQL)
    
    If res > 0 Then
    Else
        DeleteRow vasResult, iRes, iRes
        DeleteRow vasList, iList, iList
        
        SQL = "insert into pat_res(equipno, barcode, pid, pname, bcno, " & vbCrLf & _
              "dept, sampno, examdate, examtime, equipcode, " & vbCrLf & _
              "examcode, examname, result, sendflag) " & vbCrLf & _
              "values('" & gEquip & "', '" & sBarcode & "', '" & sPID & "', '" & sPName & "', '" & sBCNO & "', " & vbCrLf & _
              "'" & sDept & "', '" & sSampNo & "', '" & sExamDate & "', '" & sExamTime & "', '" & sEquipCode & "', " & vbCrLf & _
              "'" & sExamCode & "', '" & sExamName & "', '" & sResult & "', '0')"
        
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
            
        End If
        SQL = "update sampres set mflag = '1' " & vbCrLf & _
              "where sampno = '" & sSampNo & "' and examdate = '" & sExamDate & "' and examtime = '" & sExamTime & "'"
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
            
        End If
        
        iRow = -1
        For i = 1 To vasExam.DataRowCnt
            If Trim(GetText(vasExam, i, colBarcode)) = sBarcode Then
                iRow = i
                Exit For
            End If
            
        Next
        If iRow = -1 Then
            iRow = vasExam.DataRowCnt + 1
        End If
        
        If iRow > vasExam.MaxRows Then
            vasExam.MaxRows = iRow
        End If
    
        SetText vasExam, sBarcode, iRow, colBarcode
        SetText vasExam, sPID, iRow, colPID
        SetText vasExam, sPName, iRow, colPName
        SetText vasExam, sBCNO, iRow, colBCNO
        SetText vasExam, sDept, iRow, colDept
        SetText vasExam, sSampNo, iRow, colLSampNo
        SetText vasExam, sExamDate, iRow, colLExamDate
        SetText vasExam, sExamTime, iRow, colLExamTime
        SetText vasExam, sEquipCode, iRow, colLEquipCode
        SetText vasExam, sExamCode, iRow, colLExamCode
        SetText vasExam, sExamName, iRow, colLExamName
        SetText vasExam, sResult, iRow, colLResult
        SetText vasExam, "결과", iRow, colState
        
        vasExam.RowHeight(iRow) = 14
        
        For i = 1 To vasExam.DataRowCnt
            vasExam.Col = colCheckBox
            vasExam.Row = i
            
            If i = iRow Then
                vasExam.Value = 1
            Else
                vasExam.Value = 0
            End If
        Next
        
        
    End If
    If chkMode.Value = 1 Then
        Call cmdSend_Click
    End If
End Sub

Private Sub MachRes(asList As Long, asRes As Long)
    Dim i As Long
    Dim iList As Long
    Dim iRes As Long
    Dim j As Long
    Dim iRow As Long
    Dim sBarcode As String
    Dim sPID As String
    Dim sPName As String
    Dim sBCNO As String
    Dim sDept As String
    Dim sSampNo As String
    Dim sExamDate As String
    Dim sExamTime As String
    Dim sEquipCode As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim sResult As String
    
    iList = -1
    iRes = -1
    
    iList = asList
    iRes = asRes
    
    If iList = -1 Or iRes = -1 Then
        Exit Sub
    End If
    
    sBarcode = Trim(GetText(vasList, iList, colBarcode))
    sPID = Trim(GetText(vasList, iList, colPID))
    sPName = Trim(GetText(vasList, iList, colPName))
    sBCNO = Trim(GetText(vasList, iList, colBCNO))
    sDept = Trim(GetText(vasList, iList, colDept))
    sSampNo = Trim(GetText(vasResult, iRes, ColSampNo))
    sExamDate = Trim(GetText(vasResult, iRes, ColExamDate))
    sExamTime = Trim(GetText(vasResult, iRes, ColExamTime))
    sEquipCode = Trim(GetText(vasResult, iRes, ColEquipCode))
    sExamCode = Trim(GetText(vasResult, iRes, ColExamCode))
    sExamName = Trim(GetText(vasResult, iRes, colExamName))
    sResult = Trim(GetText(vasResult, iRes, ColResult))
    
    SQL = "select barcode from pat_res where equipno = '" & gEquip & "' and barcode = '" & sBarcode & "'"
    res = db_select_Col(gLocal, SQL)
    
    If res > 0 Then
    Else
        DeleteRow vasResult, iRes, iRes
        DeleteRow vasList, iList, iList
        
        SQL = "insert into pat_res(equipno, barcode, pid, pname, bcno, " & vbCrLf & _
              "dept, sampno, examdate, examtime, equipcode, " & vbCrLf & _
              "examcode, examname, result, sendflag) " & vbCrLf & _
              "values('" & gEquip & "', '" & sBarcode & "', '" & sPID & "', '" & sPName & "', '" & sBCNO & "', " & vbCrLf & _
              "'" & sDept & "', '" & sSampNo & "', '" & sExamDate & "', '" & sExamTime & "', '" & sEquipCode & "', " & vbCrLf & _
              "'" & sExamCode & "', '" & sExamName & "', '" & sResult & "', '0')"
        
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
            
        End If
        
        SQL = "update sampres set mflag = '1' " & vbCrLf & _
              "where sampno = '" & sSampNo & "' and examdate = '" & sExamDate & "' and examtime = '" & sExamTime & "'"
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
        
        iRow = -1
        For i = 1 To vasExam.DataRowCnt
            If Trim(GetText(vasExam, i, colBarcode)) = sBarcode Then
                iRow = i
                Exit For
            End If
            
        Next
        If iRow = -1 Then
            iRow = vasExam.DataRowCnt + 1
        End If
        
        If iRow > vasExam.MaxRows Then
            vasExam.MaxRows = iRow
        End If
    
        SetText vasExam, sBarcode, iRow, colBarcode
        SetText vasExam, sPID, iRow, colPID
        SetText vasExam, sPName, iRow, colPName
        SetText vasExam, sBCNO, iRow, colBCNO
        SetText vasExam, sDept, iRow, colDept
        SetText vasExam, sSampNo, iRow, colLSampNo
        SetText vasExam, sExamDate, iRow, colLExamDate
        SetText vasExam, sExamTime, iRow, colLExamTime
        SetText vasExam, sEquipCode, iRow, colLEquipCode
        SetText vasExam, sExamCode, iRow, colLExamCode
        SetText vasExam, sExamName, iRow, colLExamName
        SetText vasExam, sResult, iRow, colLResult
        SetText vasExam, "결과", iRow, colState
        
        vasExam.RowHeight(iRow) = 14
        
        For i = 1 To vasExam.DataRowCnt
            vasExam.Col = colCheckBox
            vasExam.Row = i
            
            If i = iRow Then
                vasExam.Value = 1
            Else
                vasExam.Value = 0
            End If
        Next
        
        If chkMode.Value = 1 Then
            Call cmdSend_Click
        End If
        
    End If
End Sub

Private Sub cmdResSave_Click()
    'Proc_Result txtBarcode
End Sub

Private Sub cmdResSch_Click()
    ClearSpread vasResult
    
    SQL = "select '', SampNo , ExamDate, ExamTime, EquipCode, ExamCode, ExamName, Result from sampres " & vbCrLf & _
          "where examdate = '" & Format(dtpResOnly, "yyyymmdd") & "' and mflag = '0' " & vbCrLf & _
          "group by ExamDate, ExamTime, SampNo, EquipCode, ExamCode, ExamName, Result"
    res = db_select_Vas(gLocal, SQL, vasResult)
    
    vasResult.MaxRows = vasResult.DataRowCnt
    vasResult.RowHeight(-1) = 14
    
End Sub

Private Sub cmdSend_Click()
    Dim lRow As Long
    
'''    If Trim(txtUser.Text) = "" Then
'''        MsgBox "사용자ID를 입력하세요."
'''        Exit Sub
'''
'''    End If
    
    
    For lRow = 1 To vasExam.DataRowCnt
        vasExam.Row = lRow
        vasExam.Col = 1
        If vasExam.Value = 1 Then
            If Len(Trim(GetText(vasExam, lRow, colBarcode))) = 11 Then
                res = Insert_Data(lRow)
            
'            vasExam.Value = 0
                If res = -1 Then
                    SetForeColor vasExam, lRow, lRow, colCheckBox, colCheckBox, 255, 0, 0
                    SetText vasExam, "실패", lRow, colState
                    Err_Data lRow
                Else
                    vasExam.Row = lRow
                    vasExam.Col = 1
                    
                    
                    SetBackColor vasExam, lRow, lRow, 1, colState, 202, 255, 112
                    SetText vasExam, "완료", lRow, colState
                    
                End If
            End If
        End If
        
        vasExam.Row = lRow
        vasExam.Col = 1
        vasExam.Value = 0
    Next lRow
    
'''    TransCheck
End Sub

Private Sub Err_Data(asRow As Long)
    Dim sErrData
    
    sErrData = ""
    
'''    If Trim(GetText(vasExam, asRow, colReceCode)) <> Trim(GetText(vasExam, asRow, colExamName)) Then
'''        sErrData = "처방항목과 다른 항목결과입니다."
'''    End If
'''
'''    If Trim(GetText(vasExam, asRow, colResult)) = "Aborted" Then
'''        sErrData = "결과값이 [Aborted] 입니다."
'''    End If
'''
'''
'''
'''    If sErrData <> "" Then
'''        SQL = "update pat_res set bigo = '" & sErrData & "' " & vbCrLf & _
'''              "where barcode = '" & Trim(GetText(vasExam, asRow, colBarcode)) & "' " & vbCrLf & _
'''              "and examdate = '" & Trim(GetText(vasExam, asRow, colExamDate)) & "' and examtime = '" & Trim(GetText(vasExam, asRow, colExamTime)) & "'"
'''        Res = SendQuery(gLocal, SQL)
'''        SetText vasExam, sErrData, asRow, colErrState
'''
'''        Save_Raw_Data "[SQL" & Res & "]" & SQL
'''    End If
'''
'''    spErr.Caption = Trim(GetText(vasExam, asRow, colBarcode)) & " 결과전송 실패"
'''    tmErr.Enabled = True
    
End Sub
Private Sub cmdUp_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow - 1
    vasActiveCell vasID, lRow - 1, 2
    vasID_Click 2, lRow - 1
End Sub

Private Sub cmdUser_Click()
    frmUser.Show 0
    
End Sub

Private Sub cmdWorkList_Click()
'''    Timer1.Enabled = False
'''    frmWorkList.Show
End Sub

Private Sub Command_close_Click()
    Unload Me
End Sub

Private Sub Command_config_Click()
    frmConfig.Show 1
End Sub


Private Sub Command_setup_Click()
    frmOrderCode.Show 1
    GetExamCode
End Sub

Private Sub Command2_Click()
    H232 Text2.Text, "1"
    Text2.Text = ""
End Sub

Private Sub Command3_Click()
    Dim s As String
    Dim i As Long
    
    For i = 1 To Len(txtData)
    
    s = Mid(txtData, i, 1)

    If H232_1 = "1" Then
        If s = chrNACK Then
            Save_Raw_Data "[Rx1 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & chrNACK
            H232_Connect_state = True
            tm_H232.Enabled = False
            txtBuff.Text = ""
'''            MSComm1.Output = H232_Function(H232_SerialNum)
            Save_Raw_Data "[Tx1 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & H232_Function(H232_SerialNum)
        End If
        
    ElseIf H232_1 = "2" Then
        If s = chrEOT Then
            
            If Trim(HSerial(txtBuff.Text)) = "" Then
                lblConnect.Caption = "Cobas h232 연결 성공"
            Else
                lblConnect.Caption = "Cobas h232 연결 성공 (SN:" & HSerial(txtBuff.Text) & ")"
            End If

            lblConnect.ForeColor = &H808000
            H232_1 = "3"
'''            MSComm1.Output = chrACK
            
            
        End If
        
    ElseIf H232_1 = "3" Then
        If s = chrACK Then
'''            MSComm1.Output = "a"
            txtBuff = ""
        ElseIf s = "a" Then
'''            MSComm1.Output = chrTAB
            H232_s_1 = "1"
        
        ElseIf s = chrTAB Then
            If H232_s_1 = "1" Then
'''                MSComm1.Output = Trim(txtReqS.Text)
                H232_s_1 = "2"
            ElseIf H232_s_1 = "2" Then
'''                MSComm1.Output = Trim(txtReqE.Text)
                H232_s_1 = "3"
                
            ElseIf H232_s_1 = "3" Then
'''                MSComm1.Output = "0"
                
            End If
        ElseIf s = "0" Then
'''            MSComm1.Output = vbCr
            H232_1 = "4"
        Else
            txtBuff = txtBuff & s
            If txtBuff = Trim(txtReqS.Text) Or txtBuff = Trim(txtReqE.Text) Then
'''                MSComm1.Output = chrTAB
            End If
            txtBuff = ""
            
        End If
    
    ElseIf H232_1 = "4" Then
        If s = chrSTX Then
            txtBuff = chrSTX
        ElseIf s = chrEOT Or s = chrETX Then
            txtBuff = txtBuff & s
            Save_Raw_Data "[Rx1 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & txtBuff
            H232 txtBuff.Text, "1"
            
            txtBuff.Text = ""
'''            MSComm1.Output = chrACK
            Save_Raw_Data "[Tx1 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & chrACK
        
        Else
            txtBuff = txtBuff & s
        End If
        
    End If
    Next
    
End Sub

Private Sub Command4_Click()
    H232 txtData, "1"
    txtData = ""
   
End Sub

Private Sub Command5_Click()
    H232 txtTest.Text, "1"
    
    txtTest.Text = ""
End Sub

'Private Sub Command4_Click()
'    Amplicor_INIT
'End Sub

Private Sub Form_Load()
    Dim sDate As String
        
    '메인화면 관련
    Me.Left = 0
    Me.Top = 0
'''    Me.Height = 11190
'''    Me.Width = 15360
    
    gResCol = 16
    
    '변수 초기화
    cmdReset_Click
    
    'ini파일에서 정보 가져오기
    GetSetup
    
    '컴포트 오픈
    
    MSComm1.CommPort = gSetup.gPort
    MSComm1.RTSEnable = gSetup.gRTSEnable
    MSComm1.DTREnable = gSetup.gDTREnable
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit


'''        If MSComm1.PortOpen = False Then
'''            MSComm1.PortOpen = True
'''        End If

    MSComm2.CommPort = gSetup2.gPort
    MSComm2.RTSEnable = gSetup2.gRTSEnable
    MSComm2.DTREnable = gSetup2.gDTREnable
    MSComm2.Settings = gSetup2.gSpeed & "," & gSetup2.gParity & "," & gSetup2.gDataBit & "," & gSetup2.gStopBit

'''        If MSComm2.PortOpen = False Then
'''            MSComm2.PortOpen = True
'''        End If
    
'''    If Not Connect_Server Then
'''        MsgBox "연결되지 않았습니다."
'''        cn_Server_Flag = False
'''        Exit Sub
'''    Else
'''        cn_Server_Flag = True
'''    End If

    '서버접속
'''    TP_OK = False
'''    Call tuxedo_Send.TP_INIT
    
    
    '로컬접속
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
    gTimerReq = 0
    
'    If Text_Today.Text = "2007-07-12" Then
'''        SQL = " Alter Table pat_res Alter Column diskno text(20) "
'''        Res = SendQuery(gLocal, SQL)
'''
'''        SQL = " Alter Table pat_res Alter Column posno text(20) "
'''        Res = SendQuery(gLocal, SQL)
'    End If
    
    '서버접속
'    cn_Server_Flag = dce_setenv("client.env", "", "")

    '검사일자
    
    'dtpToday = Format(CDate(GetDateFull), "yyyy/mm/dd")
    'dtpToday_1 = Format(CDate(GetDateFull), "yyyy/mm/dd")
    dtpToday = Format(CDate(Date), "yyyy/mm/dd")
    dtpToday_1 = Format(CDate(Date), "yyyy/mm/dd")

    '====================로컬 DB지우기 - 30일 보관======================
    sDate = Format(DateAdd("y", dtpToday, -30), "yyyymmdd")
    
    SQL = "Delete from pat_res where examdate < '" & sDate & "' "
    SendQuery gLocal, SQL
    '===================================================================
    
    '검사코드
    GetExamCode
    
'''    SQL = " Select unit From pat_res "
'''    Res = db_select_Col(gLocal, SQL)
'''    If Res = -1 Then
'''        SQL = " Alter Table pat_res Add Column unit text(20) "
'''        Res = SendQuery(gLocal, SQL)
'''    End If
'''
'''    SQL = " Alter Table pat_res Alter Column result text(50) "
'''    Res = SendQuery(gLocal, SQL)
'''
'''    SQL = " Alter Table pat_res Alter Column sampletype text(20) "
'''    Res = SendQuery(gLocal, SQL)
    
    txtBuff = ""
    txtData = ""
    
    vasExam.MaxRows = 0
    spNotTrans.Visible = False
    dtpResOnly = Date
    dtpSDeptDate = Date
    dtpEDeptDate = Date
    dtpExamDate = Date
    ClearSpread vasResult
    
    
    H232_Connect_state = False
    H232_Connect_state_2 = False
    
    cmdCol1_Click
    tmErr.Enabled = False

    defClr
    gRow = 0
    chkMode.Value = 0

    
End Sub

Function AutoRece(asReceDate As String, asBarcode As String, asUserID As String, Optional asRecePart As String = "O", Optional asReceCnd As String = "A", Optional asReceDept As String = "PA  ") As Integer
    'Dim strAutoRece As String
    'AutoRece = -1
    
    'strAutoRece = asReceCnd & asRecePart & asReceDept & asReceDate & asBarcode & asUserID

    'AutoRece = tuxedo_Send.TP_PUT_RESULT("ACA0119A", strAutoRece)
    
    'Save_Raw_Data AutoRece & "<ACA0119A>" & strAutoRece
    
End Function


Sub PatInfo(argSpread As vaSpread, asSpecID As String, asRow As Long)

    Dim lsBarcode As String
    Dim rv As Integer
    Dim i As Long
    
    ClearSpread vasCode
    
    lsBarcode = asSpecID
    SQL = "p_interfacequery '1', '" & lsBarcode & "'"
    res = db_select_Col(gServer, SQL)

                                                               
    If res < 1 Then
    Else
        SetText argSpread, gReadBuf(0), asRow, colPID
''        SetText argSpread, gReadBuf(9), asRow, colReceDate
''        SetText argSpread, gReadBuf(10), asRow, colReceno
'''        SetText argSpread, gReadBuf(11), asRow, colSeqNo
        SetText argSpread, gReadBuf(32), asRow, colPName
'''        gReadBuf(6) = "LEREL401"
        SQL = "select examname from equipexam where examcode = '" & Trim(gReadBuf(6)) & "'"
        res = db_select_Col(gLocal, SQL)
''        SetText argSpread, Trim(gReadBuf(0)), asRow, colReceCode
        
        Save_Raw_Data "<" & lsBarcode & ">" & gReadBuf(32) & ":" & Trim(gReadBuf(6))
        
    End If
    
End Sub

Function EquipExamCode(argEquipCode As String, argPID As String, argSENO As String, argSEQN As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
    Dim i As Integer
    Dim sExamCode As String

    EquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    ClearSpread vasTemp1
    sExamCode = ""
    
    SQL = " Select examcode From EquipExam " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(argEquipCode) & "' "
    res = db_select_Vas(gLocal, SQL, vasTemp1)
    
    If vasTemp1.DataRowCnt < 1 Then
        Exit Function
    End If
    
    For i = 1 To vasTemp1.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & ",'" & Trim(GetText(vasTemp1, i, 1)) & "'"
        Else
            sExamCode = "'" & Trim(GetText(vasTemp1, i, 1)) & "'"
        End If
    Next i

    SQL = " Select SUCD From LRESULT " & CR & _
          " Where PAID = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and SENO = " & argSENO & vbCrLf & _
          "   and SEQN = " & argSEQN & vbCrLf & _
          "   and SUCD in ( " & sExamCode & ")  "
          
    res = db_select_Col(gServer, SQL)
  
    If gReadBuf(0) <> "" Then
        EquipExamCode = Trim(gReadBuf(0))
    End If
    
End Function

Function GetExamCode() As Integer
    Dim i As Long
    Dim j As Long
    
    gAllExam = ""
    
    ClearSpread vasTemp
    GetExamCode = -1
    
    '장비코드,검사코드,검사명,참고치_Low,참고치_High,오더구분
    SQL = "Select equipcode, examcode, examname, reflow, refhigh, ordgubun, subcode " & vbCrLf & _
          "From equipexam " & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "order by  seqno "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    If res > 0 Then
        ReDim gArrEquip(1 To vasTemp.DataRowCnt, 1 To 8)
    Else
        SaveQuery SQL
        Exit Function
    End If
         
'''    vasExam.MaxCols = vasTemp.DataRowCnt * 5 + colResult - 1
''''    SetText vasExam, "COPY", 0, vasExam.MaxCols
'''    colResult1 = vasTemp.DataRowCnt * 4 + colResult
    
''''    vasExam.ColWidth(13) = 0
'''    vasExam.ColWidth(14) = 0
'''    vasExam.ColWidth(15) = 0
'''    vasExam.ColWidth(16) = 0
'''    vasExam.ColWidth(17) = 0
    
    For i = 1 To vasTemp.DataRowCnt
        gArrEquip(i, 1) = i
        For j = 1 To 7
            gArrEquip(i, j + 1) = Trim(GetText(vasTemp, i, j))
        Next j
        
        '검사명 디스플레이
        'MsgBox colResult + (i - 1) * 4
'''        SetText vasExam, gArrEquip(i, 4), 0, colResult + (i - 1) * 4
'''        vasExam.ColWidth(colResult + (i - 1) * 4) = 15
'''        vasExam.ColWidth(colResult + (i - 1) * 4 + 1) = 0
'''        vasExam.ColWidth(colResult + (i - 1) * 4 + 2) = 0
'''        vasExam.ColWidth(colResult + (i - 1) * 4 + 3) = 0
'''
'''        vasExam.ColWidth(colResult1 + i - 1) = 0
        
        If gAllExam = "" Then
            gAllExam = "'" & Trim(GetText(vasTemp, i, 2)) & "'"
        Else
            gAllExam = gAllExam & ",'" & Trim(GetText(vasTemp, i, 2)) & "'"
        End If
        
    Next i
    
    GetExamCode = 1
End Function

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

'    Call dce_close_env      ' Server와 연결을 끊는 곳
    DisConnect_Local
    DisConnect_Server
    
    
    Unload Me
    
    End
    
End Sub

Public Function HConnectState(asData As String) As Boolean
    Dim lsData As String
    Dim i As Integer
    
    HConnectState = False
    
    lsData = asData
    i = InStr(1, lsData, chrSTX)
    
    If i > 0 Then
        lsData = Mid(lsData, i + 1)
        i = InStr(1, lsData, chrHT)
        If i > 0 Then
            lsData = Mid(lsData, i + 1)
            i = InStr(1, lsData, chrHT)
            If i > 0 Then
                lsData = Mid(lsData, 1, i - 1)
                If lsData = "00FD" Or lsData = "00FF" Then
                    HConnectState = True
                End If
            End If
        End If
    End If
End Function

Public Function HSerial(asData As String) As String

    Dim lsData As String
    Dim i As Integer
    
    HSerial = ""
    
    lsData = asData
    i = InStr(1, lsData, chrSTX)
    
    If i > 0 Then
        lsData = Mid(lsData, i + 1)
        i = InStr(1, lsData, chrHT)
        If i > 0 Then
            lsData = Mid(lsData, i + 1)
            i = InStr(1, lsData, chrHT)
            If i > 0 Then
                lsData = Mid(lsData, 1, i - 1)
                HSerial = lsData
            End If
        End If
    End If
End Function

Private Sub H232(asData As String, asEquip As String)

    Dim sdata As String
    Dim i As Integer
    Dim sStr(1 To 30) As String
    Dim j As Long
    Dim k As Long
    Dim iRow As Integer
    
    Dim sBarcode As String
    Dim sEquipCode As String
    Dim sResult As String
    Dim sEquipRes As String
    Dim sExamDate As String
    Dim sExamTime As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim sSeqNo As String
    Dim sResType As String
    Dim sResPoint As String
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim liRet
    Dim sGiho As String
    Dim sFlag As String
    Dim sResFlag As String
    Dim iReceState As Integer
    
    sdata = Mid(Trim(asData), 2)
    
    i = 1
    
    For j = 1 To Len(sdata)
        If Mid(sdata, j, 1) = Chr(9) Then
            i = i + 1
            sStr(i) = ""
        Else
            sStr(i) = sStr(i) & Mid(sdata, j, 1)
        End If
        
    Next
    
    sBarcode = Trim(sStr(12))


    sEquipCode = Trim(sStr(18))
    sResult = Trim(sStr(2))
    sEquipRes = sResult
    sExamDate = "20" & Trim(sStr(4))
    sExamTime = Trim(sStr(3))
    
    If Trim(sResult) = "" Then
        Exit Sub
    End If
    

    sExamCode = ""
    sExamName = ""
    
    gReadBuf(0) = ""
    SQL = "Select EquipCode, ExamCode, ExamName, SeqNo, RSGubun, resprec, RefLow, RefHigh " & vbCrLf & _
          "from equipexam where equipno = '" & gEquip & "' and EquipCode = '" & sEquipCode & "' "
    res = db_select_Col(gLocal, SQL)
    If Trim(gReadBuf(0)) = sEquipCode Then
        sExamCode = Trim(gReadBuf(1))
        sExamName = Trim(gReadBuf(2))
        sSeqNo = Trim(gReadBuf(3))
        sResType = Trim(gReadBuf(4))
        sResPoint = Trim(gReadBuf(5))
        sRefLow = Trim(gReadBuf(6))
        sRefHigh = Trim(gReadBuf(7))
    Else
        sEquipCode = ""
    End If
    
    
    If sResPoint = "0" Then
        sResult = Format(sResult, "#0")
    ElseIf sResPoint = "1" Then
        sResult = Format(sResult, "#0.0")
    ElseIf sResPoint = "2" Then
        sResult = Format(sResult, "#0.00")
    ElseIf sResPoint = "3" Then
        sResult = Format(sResult, "#0.000")
    ElseIf sResPoint = "4" Then
        sResult = Format(sResult, "#0.0000")
    Else
        sResult = sResult
    End If
    
    If IsNumeric(sResult) = True And IsNumeric(sRefLow) = True Then
        If CCur(sResult) < CCur(sRefLow) Then
            sResult = sRefLow
        End If
        
    End If
    
    If IsNumeric(sResult) = True And IsNumeric(sRefHigh) = True Then
        If CCur(sResult) > CCur(sRefHigh) Then
            sResult = sRefHigh
        End If
        
    End If

    SQL = "select sampno from sampres " & vbCrLf & _
          "where equipno = '" & gEquip & "' and sampno = '" & sBarcode & "' " & vbCrLf & _
          "  and examdate = '" & sExamDate & "' and examtime = '" & sExamTime & "'"
    res = db_select_Col(gLocal, SQL)
    
    If res > 0 Then
        Exit Sub
        
    Else
        SQL = "insert into sampres(equipno, sampno, examdate, examtime, equipcode, examcode, examname, result, mflag) " & vbCrLf & _
              "values('" & gEquip & "', '" & sBarcode & "', '" & sExamDate & "', '" & sExamTime & "', " & vbCrLf & _
              "'" & sEquipCode & "', '" & sExamCode & "', '" & sExamName & "', '" & sResult & "', '0')"
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            'Exit Function
        End If
        
        iRow = -1
        
        For i = 1 To vasResult.DataRowCnt
            If Trim(GetText(vasResult, i, ColSampNo)) = sBarcode And Trim(GetText(vasResult, i, ColExamDate)) = sExamDate And Trim(GetText(vasResult, i, ColExamTime)) = sExamTime Then
                iRow = i
                Exit For
            End If
        Next
        If iRow = -1 Then
            iRow = vasResult.DataRowCnt + 1
            
        End If
        If iRow > vasResult.MaxRows Then
            vasResult.MaxRows = iRow
        End If
        
        SetText vasResult, sBarcode, iRow, ColSampNo
        SetText vasResult, sExamDate, iRow, ColExamDate
        SetText vasResult, sExamTime, iRow, ColExamTime
        SetText vasResult, sEquipCode, iRow, ColEquipCode
        SetText vasResult, sExamCode, iRow, ColExamCode
        SetText vasResult, sExamName, iRow, colRExamName
        SetText vasResult, sResult, iRow, ColResult
        
        vasResult.RowHeight(iRow) = 12
        
        For j = 1 To vasList.DataRowCnt
            If Trim(GetText(vasList, j, colReady)) = "대기" Then
                MachRes j, CLng(iRow)
                Exit For
            End If
        Next
    End If

'''    cmdResSch_Click
End Sub


Sub TaqMan(ArgData As String)
'''    Dim i, j, k, x As Integer
'''
'''    Dim iCnt As Integer
'''    Dim jCnt As Integer
'''    Dim aCnt As Integer
'''    Dim bCnt As Integer
'''
'''    Dim lsTmp As String
'''
'''    Dim sDate As String '필요없을 수도 있음
'''    Dim sGubun As String
'''    Dim sPID As String
'''    Dim sReceNo As String
'''    Dim sSpecID As String
'''    Dim sTestID As String
'''    Dim sExamCode As String
'''    Dim sExamName As String
'''    Dim sResClassCode As String
'''    Dim sFlag As String
'''    Dim sResult2 As String
'''
'''    Dim sAg_Res As String
'''    Dim sAb_Res As String
'''
'''    Dim sGiho As String
'''
'''    Dim sExamCode_All As String
'''
'''    Dim lRow As Long
'''    Dim lCol As Long
'''
'''    Dim lResRow As Long
'''
'''    Dim slen, sLen2 As String
'''    Dim iRCnt As Integer
'''
'''    Dim lsEquipCode As String
'''    Dim lsExamCode As String
'''    Dim lsExamName As String
'''    Dim lsSeqNo As String
'''    Dim lsResType As String
'''    Dim lsResPoint As String
'''    Dim lsRefLow As String
'''    Dim lsRefHigh As String
'''    Dim lsRes As String
'''    Dim sResult As String
'''    Dim sResultT As String
'''    Dim sQCFlag As String
'''    Dim sQC As String
'''    Dim lsResFlag As String
'''    Dim lsResNumeric As String
'''    Dim liRet As Integer
'''
'''
'''    Select Case Mid(ArgData, 2, 1)
'''    Case "H"    'Header
'''        gPreRow = -1
'''
'''        iCnt = 0
'''
'''    Case "P"    'Patient
'''        gPatFlag = -1
'''
'''        ClearSpread vasRes
'''    Case "O"    'Order
'''        iCnt = 0
'''
'''        gSpecID = ""
'''        sResult = ""
'''        sResultT = ""
'''        sQCFlag = "N"
'''        sQC = ""
'''
'''        i = InStr(1, ArgData, "|")
'''        Do While i > 0
'''            iCnt = iCnt + 1
'''            Select Case iCnt
'''            Case 3  '검체번호(QC인 경우 LotNo)
'''                sPID = Left(ArgData, i - 1)
'''                sSpecID = sPID
'''                gSpecID = sSpecID
'''
'''            Case 4  'sample position
'''                lsTmp = Left(ArgData, i - 1)
'''                If gSpecID = "" Then
'''                    gSpecID = lsTmp
'''                End If
'''
'''            Case 12 'QC 여부
'''                lsTmp = Left(ArgData, i - 1)
'''                If Left(lsTmp, 1) = "Q" Then
'''                    sQCFlag = "Y"
'''                Else
'''                    Exit Do
'''                End If
'''            Case 15 'QC Level
'''                lsTmp = Mid(ArgData, i + 1)
'''                If InStr(1, lsTmp, "^") > 0 Then
'''                    sQC = Left(lsTmp, InStr(1, lsTmp, "^") - 1)
'''                End If
'''
'''                Exit Do
'''            End Select
'''
'''            ArgData = Mid(ArgData, i + 1)
'''            i = InStr(1, ArgData, "|")
'''        Loop
'''
'''
'''        glRow = -1
'''        For i = 1 To vasExam.DataRowCnt
'''            If Trim(GetText(vasExam, i, colBarcode)) = gSpecID Then
'''                glRow = i
'''
'''
'''                Exit For
'''            End If
'''        Next i
'''
'''        If glRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
'''            glRow = vasExam.DataRowCnt + 1
'''            If glRow > vasExam.MaxRows Then
'''                vasExam.MaxRows = glRow + 1
'''            End If
'''        End If
'''
'''        SetText vasExam, gSpecID, glRow, colBarcode
'''        vasActiveCell vasID, glRow, colBarcode
'''
'''        If Trim(GetText(vasExam, glRow, colPID)) = "" Then
'''            PatInfo vasExam, gSpecID, glRow
'''        End If
'''
'''        If gPatFlag = -1 Then
'''            ClearSpread vasRes
''''            vasExam_Click colBarcode, glRow
'''
'''            gPatFlag = 1
'''        End If
'''
'''        gPreSpecID = gSpecID
'''        gPreRow = glRow
'''
'''    Case "R"    'Result
'''        gRecodeType = "R"
'''
'''        SetText vasExam, "결과", glRow, colState
'''        SetBackColor vasExam, glRow, glRow, colCheckBox, colState, 255, 250, 205
'''
'''        sExamCode = ""
'''        sResClassCode = ""
'''        sExamName = ""
'''        'sResult = ""
'''
'''        iCnt = 0
'''        i = InStr(1, ArgData, "|")
'''        Do While i > 0
'''            iCnt = iCnt + 1
'''            lsTmp = Left(ArgData, i - 1)
'''            Select Case iCnt
'''            Case 3
'''                lsEquipCode = Mid(lsTmp, 4)
'''
'''            Case 4
'''                sTestID = lsEquipCode
'''                sResult = Trim(lsTmp)
'''                sResult2 = Trim(sResult)
'''                sResult = Result_Set(sTestID, sResult)
'''
''''                res = Result_Set(sTestID, sResult)
''''                If res > 0 Then
''''                    lResRow = res
'''
'''                    k = -1
'''                    For j = LBound(gArrEquip) To UBound(gArrEquip)
'''                        If Trim(gArrEquip(j, 2)) = Trim(sTestID) Then
'''                            k = j
'''                            Exit For
'''                        End If
'''                    Next j
'''
'''                    If k > 0 Then
'''                        vasExam.SetText colResult + (k - 1) * 4, glRow, sResult
''''                        For x = 20 To 40
''''                        vasExam.SetText x, glRow, sResult
''''                        Next
'''                        If sTestID = "HBMCAP96" And sResult2 <> "" Then
'''                            If IsNumeric(sResult2) = True Then
'''                                lsResNumeric = CCur(sResult2) * 5.82
'''                            ElseIf Mid(sResult2, 1, 1) = ">" Or Mid(sResult2, 1, 1) = "<" Then
'''                                If IsNumeric(Mid(sResult2, 2)) = True Then
'''                                    lsResNumeric = CCur(Mid(sResult2, 2)) * 5.82
'''                                End If
'''                            Else
'''                                lsResNumeric = sResult2
'''                            End If
'''                            lsResNumeric = Result_Set_1(sTestID, lsResNumeric)
''''                            If Right(lsResNumeric, 1) = "." Then
''''                                lsResNumeric = Mid(lsResNumeric, 1, Len(lsResNumeric) - 1)
''''                            End If
'''                            vasExam.SetText colResult + (k) * 4, glRow, Trim(lsResFlag & lsResNumeric)
'''
'''                        ElseIf sTestID = "HB2CAP96" And sResult2 <> "" Then
'''                            If IsNumeric(sResult2) = True Then
'''                                lsResNumeric = CCur(sResult2) * 5.82
'''                            ElseIf Mid(sResult2, 1, 1) = ">" Or Mid(sResult2, 1, 1) = "<" Then
'''                                If IsNumeric(Mid(sResult2, 2)) = True Then
'''                                    lsResNumeric = CCur(Mid(sResult2, 2)) * 5.82
'''                                End If
'''                            Else
'''                                lsResNumeric = sResult2
'''                            End If
'''                            lsResNumeric = Result_Set_1(sTestID, lsResNumeric)
''''                            If Right(lsResNumeric, 1) = "." Then
''''                                lsResNumeric = Mid(lsResNumeric, 1, Len(lsResNumeric) - 1)
''''                            End If
'''                            vasExam.SetText colResult + (k) * 4, glRow, Trim(lsResFlag & lsResNumeric)
'''
'''                        ElseIf sTestID = "HCMCAP96" And sResult2 <> "" Then
'''                            If IsNumeric(sResult2) = True Then
'''                                lsResNumeric = CCur(sResult2) * 2.7
'''                            ElseIf Mid(sResult2, 1, 1) = ">" Or Mid(sResult2, 1, 1) = "<" Then
'''                                If IsNumeric(Mid(sResult2, 2)) = True Then
'''                                    lsResNumeric = CCur(Mid(sResult2, 2)) * 2.7
'''                                End If
'''                            Else
'''                                lsResNumeric = sResult2
'''                            End If
'''                            lsResNumeric = Result_Set_1(sTestID, lsResNumeric)
''''                            If Right(lsResNumeric, 1) = "." Then
''''                                lsResNumeric = Mid(lsResNumeric, 1, Len(lsResNumeric) - 1)
''''                            End If
'''                            vasExam.SetText colResult + (k) * 4, glRow, Trim(lsResFlag & lsResNumeric)
'''
'''                        End If
'''                    End If
'''
'''
'''
'''                    gReadBuf(0) = ""
'''                    SQL = "Select EquipCode, ExamCode, ExamName, SeqNo, RSGubun, resprec, RefLow, RefHigh " & vbCrLf & _
'''                          "from equipexam where equipno = '" & gEquip & "' and EquipCode = '" & lsEquipCode & "' "
'''                    res = db_select_Col(gLocal, SQL)
'''                    If Trim(gReadBuf(0)) = lsEquipCode Then
'''                        lsExamCode = Trim(gReadBuf(1))
'''                        lsExamName = Trim(gReadBuf(2))
'''                        lsSeqNo = Trim(gReadBuf(3))
'''                        lsResType = Trim(gReadBuf(4))
'''                        lsResPoint = Trim(gReadBuf(5))
'''                        lsRefLow = Trim(gReadBuf(6))
'''                        lsRefHigh = Trim(gReadBuf(7))
'''                    Else
'''                        lsEquipCode = ""
'''                    End If
'''                    SQL = "Delete FROM pat_res " & vbCrLf & _
'''                          "WHERE examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'''                          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'''                          "  AND equipcode = '" & lsEquipCode & "'" & vbCrLf & _
'''                          "  AND barcode = '" & Trim(GetText(vasExam, glRow, 2)) & "' "
'''                    res = SendQuery(gLocal, SQL)
'''                    If res = -1 Then
'''                        SaveQuery SQL
'''                        'Exit Function
'''                    End If
'''                    If sTestID = "HBMCAP96" Then
'''                        SQL = "INSERT INTO pat_res (examdate, equipno, " & _
'''                                "barcode, sampletype, receno, " & _
'''                                "pid, pname, jumin, page, psex, " & _
'''                                "resdate, seqno, diskno, posno, " & _
'''                                "equipcode, examcode, " & _
'''                                "result, sendflag, examname, " & _
'''                                "refflag,panicflag, deltaflag, unit, refvalue, panicvalue, subcode, recedate ) " & vbCrLf & _
'''                              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'''                              "'" & Trim(GetText(vasExam, glRow, 2)) & "','" & Trim(GetText(vasExam, glRow, 3)) & "', '" & Trim(GetText(vasExam, glRow, 4)) & "', " & _
'''                              "'" & Trim(GetText(vasExam, glRow, 5)) & "', '" & Trim(GetText(vasExam, glRow, 6)) & "', '" & Trim(GetText(vasExam, glRow, 9)) & "', 0, '" & Trim(GetText(vasExam, glRow, 7)) & "', " & _
'''                              "'" & Format(Date, "yyyymmdd") & "', '" & lsSeqNo & "', '', '', " & vbCrLf & _
'''                              "'" & lsEquipCode & "', 'LMOHL613', " & _
'''                              "'" & sResult & "', 'B', '" & lsExamName & "', " & vbCrLf & _
'''                              "'', '', '', '', " & _
'''                              "'', '', '00','" & Trim(GetText(vasExam, glRow, colReceDate)) & "' ) "
'''                        res = SendQuery(gLocal, SQL)
'''                        If res = -1 Then
'''                            SaveQuery SQL
'''                            'Exit Function
'''                        End If
'''                        SQL = "INSERT INTO pat_res (examdate, equipno, " & _
'''                                "barcode, sampletype, receno, " & _
'''                                "pid, pname, jumin, page, psex, " & _
'''                                "resdate, seqno, diskno, posno, " & _
'''                                "equipcode, examcode, " & _
'''                                "result, sendflag, examname, " & _
'''                                "refflag,panicflag, deltaflag, unit, refvalue, panicvalue, subcode, recedate ) " & vbCrLf & _
'''                              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'''                              "'" & Trim(GetText(vasExam, glRow, 2)) & "','" & Trim(GetText(vasExam, glRow, 3)) & "', '" & Trim(GetText(vasExam, glRow, 4)) & "', " & _
'''                              "'" & Trim(GetText(vasExam, glRow, 5)) & "', '" & Trim(GetText(vasExam, glRow, 6)) & "', '" & Trim(GetText(vasExam, glRow, 9)) & "', 0, '" & Trim(GetText(vasExam, glRow, 7)) & "', " & _
'''                              "'" & Format(Date, "yyyymmdd") & "', '" & lsSeqNo & "', '', '', " & vbCrLf & _
'''                              "'" & lsEquipCode & "', 'LMOHL614', " & _
'''                              "'" & lsResNumeric & "', 'B', '" & lsExamName & "', " & vbCrLf & _
'''                              "'', '', '', '', " & _
'''                              "'', '','00','" & Trim(GetText(vasExam, glRow, colReceDate)) & "' ) "
'''                        res = SendQuery(gLocal, SQL)
'''                        If res = -1 Then
'''                            SaveQuery SQL
'''                            'Exit Function
'''                        End If
'''                    ElseIf sTestID = "HB2CAP96" Then
'''                        SQL = "INSERT INTO pat_res (examdate, equipno, " & _
'''                                "barcode, sampletype, receno, " & _
'''                                "pid, pname, jumin, page, psex, " & _
'''                                "resdate, seqno, diskno, posno, " & _
'''                                "equipcode, examcode, " & _
'''                                "result, sendflag, examname, " & _
'''                                "refflag,panicflag, deltaflag, unit, refvalue, panicvalue, subcode, recedate ) " & vbCrLf & _
'''                              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'''                              "'" & Trim(GetText(vasExam, glRow, 2)) & "','" & Trim(GetText(vasExam, glRow, 3)) & "', '" & Trim(GetText(vasExam, glRow, 4)) & "', " & _
'''                              "'" & Trim(GetText(vasExam, glRow, 5)) & "', '" & Trim(GetText(vasExam, glRow, 6)) & "', '" & Trim(GetText(vasExam, glRow, 9)) & "', 0, '" & Trim(GetText(vasExam, glRow, 7)) & "', " & _
'''                              "'" & Format(Date, "yyyymmdd") & "', '" & lsSeqNo & "', '', '', " & vbCrLf & _
'''                              "'" & lsEquipCode & "', 'LMOHL613', " & _
'''                              "'" & sResult & "', 'B', '" & lsExamName & "', " & vbCrLf & _
'''                              "'', '', '', '', " & _
'''                              "'', '', '00','" & Trim(GetText(vasExam, glRow, colReceDate)) & "' ) "
'''                        res = SendQuery(gLocal, SQL)
'''                        If res = -1 Then
'''                            SaveQuery SQL
'''                            'Exit Function
'''                        End If
'''                        SQL = "INSERT INTO pat_res (examdate, equipno, " & _
'''                                "barcode, sampletype, receno, " & _
'''                                "pid, pname, jumin, page, psex, " & _
'''                                "resdate, seqno, diskno, posno, " & _
'''                                "equipcode, examcode, " & _
'''                                "result, sendflag, examname, " & _
'''                                "refflag,panicflag, deltaflag, unit, refvalue, panicvalue, subcode, recedate ) " & vbCrLf & _
'''                              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'''                              "'" & Trim(GetText(vasExam, glRow, 2)) & "','" & Trim(GetText(vasExam, glRow, 3)) & "', '" & Trim(GetText(vasExam, glRow, 4)) & "', " & _
'''                              "'" & Trim(GetText(vasExam, glRow, 5)) & "', '" & Trim(GetText(vasExam, glRow, 6)) & "', '" & Trim(GetText(vasExam, glRow, 9)) & "', 0, '" & Trim(GetText(vasExam, glRow, 7)) & "', " & _
'''                              "'" & Format(Date, "yyyymmdd") & "', '" & lsSeqNo & "', '', '', " & vbCrLf & _
'''                              "'" & lsEquipCode & "', 'LMOHL614', " & _
'''                              "'" & lsResNumeric & "', 'B', '" & lsExamName & "', " & vbCrLf & _
'''                              "'', '', '', '', " & _
'''                              "'', '','00','" & Trim(GetText(vasExam, glRow, colReceDate)) & "' ) "
'''                        res = SendQuery(gLocal, SQL)
'''                        If res = -1 Then
'''                            SaveQuery SQL
'''                            'Exit Function
'''                        End If
'''                    ElseIf sTestID = "HCMCAP96" Then
'''                        SQL = "INSERT INTO pat_res (examdate, equipno, " & _
'''                                "barcode, sampletype, receno, " & _
'''                                "pid, pname, jumin, page, psex, " & _
'''                                "resdate, seqno, diskno, posno, " & _
'''                                "equipcode, examcode, " & _
'''                                "result, sendflag, examname, " & _
'''                                "refflag,panicflag, deltaflag, unit, refvalue, panicvalue, subcode, recedate ) " & vbCrLf & _
'''                              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'''                              "'" & Trim(GetText(vasExam, glRow, 2)) & "','" & Trim(GetText(vasExam, glRow, 3)) & "', '" & Trim(GetText(vasExam, glRow, 4)) & "', " & _
'''                              "'" & Trim(GetText(vasExam, glRow, 5)) & "', '" & Trim(GetText(vasExam, glRow, 6)) & "', '" & Trim(GetText(vasExam, glRow, 9)) & "', 0, '" & Trim(GetText(vasExam, glRow, 7)) & "', " & _
'''                              "'" & Format(Date, "yyyymmdd") & "', '" & lsSeqNo & "', '', '', " & vbCrLf & _
'''                              "'" & lsEquipCode & "', 'LMOHL622', " & _
'''                              "'" & sResult & "', 'B', '" & lsExamName & "', " & vbCrLf & _
'''                              "'', '', '', '', " & _
'''                              "'', '', '00','" & Trim(GetText(vasExam, glRow, colReceDate)) & "' ) "
'''                        res = SendQuery(gLocal, SQL)
'''                        If res = -1 Then
'''                            SaveQuery SQL
'''                            'Exit Function
'''                        End If
'''                        SQL = "INSERT INTO pat_res (examdate, equipno, " & _
'''                                "barcode, sampletype, receno, " & _
'''                                "pid, pname, jumin, page, psex, " & _
'''                                "resdate, seqno, diskno, posno, " & _
'''                                "equipcode, examcode, " & _
'''                                "result, sendflag, examname, " & _
'''                                "refflag,panicflag, deltaflag, unit, refvalue, panicvalue, subcode, recedate ) " & vbCrLf & _
'''                              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'''                              "'" & Trim(GetText(vasExam, glRow, 2)) & "','" & Trim(GetText(vasExam, glRow, 3)) & "', '" & Trim(GetText(vasExam, glRow, 4)) & "', " & _
'''                              "'" & Trim(GetText(vasExam, glRow, 5)) & "', '" & Trim(GetText(vasExam, glRow, 6)) & "', '" & Trim(GetText(vasExam, glRow, 9)) & "', 0, '" & Trim(GetText(vasExam, glRow, 7)) & "', " & _
'''                              "'" & Format(Date, "yyyymmdd") & "', '" & lsSeqNo & "', '', '', " & vbCrLf & _
'''                              "'" & lsEquipCode & "', 'LMOHL623', " & _
'''                              "'" & lsResNumeric & "', 'B', '" & lsExamName & "', " & vbCrLf & _
'''                              "'', '', '', '', " & _
'''                              "'', '','00','" & Trim(GetText(vasExam, glRow, colReceDate)) & "' ) "
'''                        res = SendQuery(gLocal, SQL)
'''                        If res = -1 Then
'''                            SaveQuery SQL
'''                            'Exit Function
'''                        End If
'''                    Else
'''
'''                        SQL = "INSERT INTO pat_res (examdate, equipno, " & _
'''                                "barcode, sampletype, receno, " & _
'''                                "pid, pname, jumin, page, psex, " & _
'''                                "resdate, seqno, diskno, posno, " & _
'''                                "equipcode, examcode, " & _
'''                                "result, sendflag, examname, " & _
'''                                "refflag,panicflag, deltaflag, unit, refvalue, panicvalue, subcode, recedate ) " & vbCrLf & _
'''                              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'''                              "'" & Trim(GetText(vasExam, glRow, 2)) & "','" & Trim(GetText(vasExam, glRow, 3)) & "', '" & Trim(GetText(vasExam, glRow, 4)) & "', " & _
'''                              "'" & Trim(GetText(vasExam, glRow, 5)) & "', '" & Trim(GetText(vasExam, glRow, 6)) & "', '" & Trim(GetText(vasExam, glRow, 9)) & "', 0, '" & Trim(GetText(vasExam, glRow, 7)) & "', " & _
'''                              "'" & Format(Date, "yyyymmdd") & "', '" & lsSeqNo & "', '', '', " & vbCrLf & _
'''                              "'" & lsEquipCode & "', '" & lsExamCode & "', " & _
'''                              "'" & sResult & "', 'B', '" & lsExamName & "', " & vbCrLf & _
'''                              "'', '', '', '', " & _
'''                              "'', '' ,'' ,'" & Trim(GetText(vasExam, glRow, colReceDate)) & "') "
'''                        res = SendQuery(gLocal, SQL)
'''                        If res = -1 Then
'''                            SaveQuery SQL
'''                            'Exit Function
'''                        End If
'''                    End If
'''
''''                    SQL = "select rstdate, rstval from klc010m " & vbCrLf & _
''''                          "where pid = '" & Trim(GetText(vasID, glRow, colPID)) & "' " & vbCrLf & _
''''                          "and PARTCD  || SLIPCD  || SPCCD || ordcd = '" & gReadBuf(0) & "' " & vbCrLf & _
''''                          "and subsqno = '01' order by rstdate desc"
''''
''''                    res = db_select_Col(gServer, SQL)
'''
''''                    vasID.SetText colPdate, glRow, Trim(gReadBuf(0))
''''                    vasID.SetText colPResult, glRow, Trim(gReadBuf(1))
'''
''''                    Save_Local_One glRow, lResRow, "1"
''''                End If
'''                Exit Sub
'''            Case 5
''''                sUnit = lsTmp
'''            Case 7
''''                sRef = lsTmp
'''                Exit Do
'''            End Select
'''
'''            ArgData = Mid(ArgData, i + 1)
'''            i = InStr(1, ArgData, "|")
'''        Loop
'''
'''        gMsgFlag = ""
'''        gHeadRecode = ""
'''        txtBuff.Text = ""
'''
'''
'''    Case "L"    '자료수신 완료
'''        If chkMode.Value = 1 Then
'''            liRet = -1
'''            If glRow < 1 Then
'''                Exit Sub
'''            End If
'''
'''            If Trim(GetText(vasExam, glRow, colPID)) <> "" Then
'''                liRet = Insert_Data(glRow)
'''            End If
'''
'''            If liRet = -1 Then
'''                SetBackColor vasExam, glRow, glRow, colState, colState, 255, 0, 0
'''                SetText vasExam, "실패", glRow, colState
'''            Else
'''                SetBackColor vasExam, glRow, glRow, 1, colState, 202, 255, 112
'''                SetText vasExam, "완료", glRow, colState
'''
'''                'Local 상태를 서버전송(C)로 바꿈
'''                SQL = " Update pat_res Set " & vbCrLf & _
'''                      " sendflag = 'C' " & vbCrLf & _
'''                      " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'''                      " And equipno = '" & gEquip & "' " & vbCrLf & _
'''                      " And barcode = '" & Trim(GetText(vasExam, glRow, colBarcode)) & "' " '& vbCrLf & _
'''                      " And equipcode = '" & lsEquipCode & "' "
'''                res = SendQuery(gLocal, SQL)
'''                If res = -1 Then
'''                    SaveQuery SQL
'''                    Exit Sub
'''                End If
'''            End If
'''        End If
'''''        i = InStr(1, argData, chrETX)
'''''        gMsgEnd = Mid(argData, 3, i - 2)
''''        vasID_Click colBarcode, glRow
''''
''''
''''        If chkMode.Value = 1 And glRow > 0 And glRow <= vasID.DataRowCnt And gRecodeType = "R" Then
''''            If Trim(GetText(vasID, glRow, colPSex)) = "Q" Then Exit Sub
''''            'res = Insert_Data(gPreRow)
''''            'Res = ToServer(glRow)
''''            res = ToServer_TXT(glRow)
''''            If res = 1 Then
''''                SetBackColor vasID, glRow, glRow, colCheckBox, colState, 202, 255, 112
''''                SetText vasID, "완료", gPreRow, colState
''''            ElseIf res = -1 Then
''''                SetForeColor vasID, glRow, glRow, 255, 0, 0
''''                SetText vasID, "실패", gPreRow, colState
''''            End If
''''        End If
'''
'''    End Select
        
End Sub

Function Result_Set(ByVal asTest As String, ByVal asRes As String) As String
    Dim sGiho As String
    Dim sRes As String
    Dim sRes1 As String
    Dim sFormat As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim ii As Integer
    Dim sResi As Integer
    Dim sRes2 As String
        
        
    Dim iRCnt
    
    Dim i, j As Integer
    Dim lResRow As Integer
    
    Result_Set = ""
    
    If Trim(asTest) = "" Then Exit Function
    
    SQL = "Select EquipCode, ExamCode, ExamName, '', resprec " & vbCrLf & _
          " " & vbCrLf & _
          "from EquipExam " & vbCrLf & _
          "where Equipno = '" & gEquip & "' " & vbCrLf & _
          "  and EquipCode = '" & Trim(asTest) & "' "
    res = db_select_Col(gLocal, SQL)
    If res < 1 Then Exit Function
    If Trim(gReadBuf(0)) <> Trim(asTest) Then Exit Function
    
    sGiho = ""
    sRes = ""
    sRes1 = ""
    
    sExamCode = Trim(gReadBuf(1))
    sExamName = Trim(gReadBuf(2))
    
    If Trim(sExamCode) = "" Then Exit Function
    
    
        
    For i = 1 To Len(asRes)
        If IsNumeric(Mid(asRes, i, 1)) = True Or Mid(asRes, i, 1) = "." Then
            sRes = sRes & Mid(asRes, i, 1)
        ElseIf Mid(asRes, i, 1) = "<" Or Mid(asRes, i, 1) = ">" Then
            sGiho = sGiho & Mid(asRes, i, 1)
        End If
    Next i
    If asRes = "Target Not Detected" Then
        sRes = "Not Detected"
    ElseIf asRes = "Invalid" Then
        sRes = "Invalid"
    ElseIf asRes = "Failed" Then
        sRes = "Failed"
    End If
    If IsNumeric(asRes) Then
        If asTest = "HCMCAP96" Then
            If CCur(asRes) < 15 Then
                sRes = "Detected<15"
            ElseIf CCur(asRes) > 6.9 * 10 ^ 7 Then
                sRes = ">6.90X(10^7)"
            End If
        ElseIf asTest = "HB2CAP96" Then
            If CCur(asRes) < 20 Then
                sRes = "Detected<20"
            ElseIf CCur(asRes) > 1.7 * 10 ^ 8 Then
                sRes = ">1.70X(10^8)"
            End If
        ElseIf asTest = "HBMCAP96" Then
            If CCur(asRes) < 12 Then
                sRes = "Detected<12"
            ElseIf CCur(asRes) > 1.1 * 10 ^ 8 Then
                sRes = ">1.10X(10^8)"
            End If
        End If
    Else

    End If
    If IsNumeric(sRes) Then

        
        sResi = 0
        sRes2 = ""
        For ii = 1 To Len(sRes)
            If IsNumeric(Mid(sRes, ii, 1)) = True Then
                sResi = sResi + 1
            Else
            End If
            If sResi > 4 Then
                
                If Mid(sRes, ii, 1) = "." Then
                    sRes2 = sRes2 & Mid(sRes, ii, 1)
                Else
                    sRes2 = sRes2 & "0"
                End If
            Else
                
                sRes2 = sRes2 & Mid(sRes, ii, 1)
            End If
        Next
'        sRes1 = sRes2
        sRes1 = Format(CCur(sRes2), "#0.00")
        sRes1 = Format(CCur(sRes1), "0.00E+0")
        sRes1 = Format(Left(sRes1, InStr(1, sRes1, "E") - 1), "#0.00") & "X(10^" & Right(sRes1, Len(sRes1) - InStr(1, sRes1, "E") - 1) & ")"
        
        
'        sRes1 = Format(CCur(sRes), "#0.00")
'        sRes1 = Format(CCur(sRes1), "0.00E+0")
'        sRes1 = Left(sRes1, InStr(1, sRes1, "E") - 1) & "*(10^" & Right(sRes1, Len(sRes1) - InStr(1, sRes1, "E") - 1) & ")"
    Else
        sRes1 = sRes
    End If
    
           
    Result_Set = sGiho & sRes1
End Function

Function Result_Set_1(ByVal asTest As String, ByVal asRes As String) As String
    Dim sGiho As String
    Dim sRes As String
    Dim sRes1 As String
    Dim sFormat As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim ii As Integer
    Dim sResi As Integer
    Dim sRes2 As String
    Dim iRCnt
    
    Dim i, j As Integer
    Dim lResRow As Integer
    
    Result_Set_1 = ""
    
    If Trim(asTest) = "" Then Exit Function
    
    SQL = "Select EquipCode, ExamCode, ExamName, '', resprec " & vbCrLf & _
          " " & vbCrLf & _
          "from EquipExam " & vbCrLf & _
          "where Equipno = '" & gEquip & "' " & vbCrLf & _
          "  and EquipCode = '" & Trim(asTest) & "' "
    res = db_select_Col(gLocal, SQL)
    If res < 1 Then Exit Function
    If Trim(gReadBuf(0)) <> Trim(asTest) Then Exit Function
    
    sGiho = ""
    sRes = ""
    sRes1 = ""
    
    sExamCode = Trim(gReadBuf(1))
    sExamName = Trim(gReadBuf(2))
    
    If Trim(sExamCode) = "" Then Exit Function
    
    
        
    For i = 1 To Len(asRes)
        If IsNumeric(Mid(asRes, i, 1)) = True Or Mid(asRes, i, 1) = "." Then
            sRes = sRes & Mid(asRes, i, 1)
        ElseIf Mid(asRes, i, 1) = "<" Or Mid(asRes, i, 1) = ">" Then
            sGiho = sGiho & Mid(asRes, i, 1)
        End If
    Next i
    If asRes = "Target Not Detected" Then
        sRes = "Not Detected"
    ElseIf asRes = "Invalid" Then
        sRes = "Invalid"
    ElseIf asRes = "Failed" Then
        sRes = "Failed"
    End If
    If IsNumeric(asRes) Then
        If asTest = "HCMCAP96" Then
            If CCur(asRes) < 41 Then
                sRes = "Detected<41"
            ElseIf CCur(asRes) > 1.86 * 10 ^ 8 Then
                sRes = ">1.86X(10^8)"
            End If
        ElseIf asTest = "HBMCAP96" Then
            If CCur(asRes) < 70 Then
                sRes = "Detected<70"
            ElseIf CCur(asRes) > 6.4 * 10 ^ 8 Then
                sRes = ">6.40X(10^8)"
            End If
        ElseIf asTest = "HB2CAP96" Then
            If CCur(asRes) < 116 Then
                sRes = "Detected<116"
            ElseIf CCur(asRes) > 9.89 * 10 ^ 8 Then
                sRes = ">9.89X(10^8)"
            End If
        End If
    Else

    End If
    If IsNumeric(sRes) Then
    
        
        sResi = 0
        sRes2 = ""
        For ii = 1 To Len(sRes)
            If IsNumeric(Mid(sRes, ii, 1)) = True Then
                sResi = sResi + 1
            Else
            End If
            If sResi > 4 Then
                If Mid(sRes, ii, 1) = "." Then
                    sRes2 = sRes2 & Mid(sRes, ii, 1)
                Else
                    sRes2 = sRes2 & "0"
                End If
            Else
                
                sRes2 = sRes2 & Mid(sRes, ii, 1)
            End If
        Next
'        sRes1 = sRes2
        sRes1 = Format(CCur(sRes2), "#0.00")
        sRes1 = Format(CCur(sRes1), "0.00E+0")
        sRes1 = Format(Left(sRes1, InStr(1, sRes1, "E") - 1), "#0.00") & "X(10^" & Right(sRes1, Len(sRes1) - InStr(1, sRes1, "E") - 1) & ")"
            
'        sRes1 = Format(CCur(sRes), "#0.00")
'        sRes1 = Format(CCur(sRes1), "0.00E+0")
'        sRes1 = Left(sRes1, InStr(1, sRes1, "E") - 1) & "*(10^" & Right(sRes1, Len(sRes1) - InStr(1, sRes1, "E") - 1) & ")"
    Else
        sRes1 = sRes
    End If
    
           
    Result_Set_1 = sGiho & sRes1
End Function

Sub COBAS_Amplicor()
'''    Dim myVar As String
'''    Dim i, j, k, a As Long
'''    Dim liRet As Integer
'''    Dim lsData As String
'''    Dim lsTmp As String
'''
'''    Dim lsRing As String
'''    Dim lsOrdDate As String
'''    Dim lsTube As String
'''    Dim lsOrdType As String
'''    Dim lsSpcID As String
'''
'''    Dim lsTestType As String
'''    Dim lsResDate As String
'''    Dim lsQualRes As String
'''    Dim lsQualFlag As String
'''    Dim lsQualResRaw As String
'''    Dim lsQualQS As String
'''    Dim lsQuanRes As String
'''    Dim lsQuanFlag As String
'''    Dim lsQuanResRaw As String
'''    Dim lsQuanQS As String
'''
'''    Dim lsID As String
'''    Dim lRow As Long
'''    Dim lsEquipCode As String
'''    Dim lsExamCode As String
'''    Dim lsExamName As String
'''    Dim lsSeqNo As String
'''    Dim lsResType As String
'''    Dim lsResPoint As String
'''    Dim lsRefLow As String
'''    Dim lsRefHigh As String
'''    Dim lsRes As String
'''
'''    myVar = Trim(txtBuff)
'''    i = InStr(1, myVar, chrLF)
'''    Do While i > 0
'''        lsData = Left(myVar, i - 1)
'''        myVar = Mid(myVar, i + 1)
'''
'''        Select Case Left(lsData, 2)
'''        'Result.Order processing
'''        Case "00"   'Result Selection
'''        Case "01"   'A-ring ID
'''            lsRing = Trim(Mid(lsData, 4, 6))
'''        Case "02"   'Order Date / Time
'''            lsOrdDate = Trim(Mid(lsData, 4))
'''        Case "03"   'Order Run Mode
'''        Case "04"   'A-tube Position
'''            lsTube = Trim(Mid(lsData, 4, 2))
'''        Case "05"   'Order Type
'''            lsOrdType = Trim(Mid(lsData, 4, 1))
'''        Case "06"   'Specimen Information
'''            lsSpcID = Trim(Mid(lsData, 4, 2))
'''        Case "07"   'Test ID
'''            lsEquipCode = Trim(Mid(lsData, 4, 3))
'''
'''            lsExamCode = ""
'''            lsExamName = ""
'''            lsSeqNo = ""
'''            lsResType = ""
'''            lsResPoint = ""
'''            lsRefLow = ""
'''            lsRefHigh = ""
'''
'''            lsRes = ""
'''
'''            SQL = "Select EquipCode, ExamCode, ExamName, SeqNo, RSGubun, resprec, RefLow, RefHigh " & vbCrLf & _
'''                  "from equipexam where equipno = '" & gEquip & "' and EquipCode = '" & lsEquipCode & "' "
'''            res = db_select_Col(gLocal, SQL)
'''            If Trim(gReadBuf(0)) = lsEquipCode Then
'''                lsExamCode = Trim(gReadBuf(1))
'''                lsExamName = Trim(gReadBuf(2))
'''                lsSeqNo = Trim(gReadBuf(3))
'''                lsResType = Trim(gReadBuf(4))
'''                lsResPoint = Trim(gReadBuf(5))
'''                lsRefLow = Trim(gReadBuf(6))
'''                lsRefHigh = Trim(gReadBuf(7))
'''            Else
'''                lsEquipCode = ""
'''            End If
'''        Case "08"   'Test Type
'''            lsTestType = Trim(Mid(lsData, 4, 1))
'''        Case "10"   'Result Date / Time
'''            lsResDate = Trim(Mid(lsData, 4))
''''            If IsDate(lsResDate) Then
''''                lsResDate = Format(lsResDate, "yyyy-mm-dd hh:nn:ss")
''''            End If
'''            lsResDate = Mid(lsResDate, 7, 4) & "-" & Mid(lsResDate, 4, 2) & "-" & Mid(lsResDate, 1, 2) & Mid(lsResDate, 11)
'''        Case "11"   'Qualitative Result
'''            lsQualRes = Trim(Mid(lsData, 4, 1))
'''            Select Case lsQualRes
'''
'''            Case "1"
'''                lsQualRes = "Positive"
'''            Case "2"
'''                lsQualRes = "Negative"
'''            Case "3"
'''                lsQualRes = "Trace"
'''            Case Else
'''                lsQualRes = ""
'''            End Select
'''            lsQualFlag = Trim(Mid(lsData, 6, 8))
'''
'''            If lsResType = "T" Then
'''                lsRes = lsQualRes
'''
'''                lRow = -1
'''                For lRow = 1 To vasExam.DataRowCnt
'''                    If Trim(GetText(vasExam, lRow, colRack)) = lsRing And _
'''                       Trim(GetText(vasExam, lRow, colTube)) = lsTube Then
'''                        gRow = lRow
'''                        Exit For
'''                    End If
'''                Next lRow
'''
'''                If lRow = -1 Then
'''                    SQL = "Select barcode, pid, pname, examtype, diskno, posno " & vbCrLf & _
'''                          "from pat_res " & vbCrLf & _
'''                          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'''                          "  and equipno = '" & gEquip & "' " & vbCrLf & _
'''                          "  and diskno = '" & lsRing & "'   " & vbCrLf & _
'''                          "  and posno = '" & lsTube & "' " & vbCrLf & _
'''                          "  and equipcode = '" & lsEquipCode & "' "
'''                    res = db_select_Vas(gLocal, SQL, vasExam, vasExam.DataRowCnt + 1, 2)
'''                    If res = 0 Then
'''                        lRow = vasExam.DataRowCnt + 1
'''                        gRow = lRow
'''
'''                        vasExam.SetText 2, lRow, lsRing & "-" & lsTube
'''                    Else
'''                        lRow = vasExam.DataRowCnt
'''                        gRow = lRow
'''                    End If
'''                Else
'''                    gRow = lRow
'''                End If
'''                'gRow = gRow + 1
'''                lRow = gRow
'''
'''                vasExam.SetText 10, lRow, lsRing
'''                vasExam.SetText 11, lRow, lsTube
'''                vasExam.SetText 12, lRow, "수신"
'''
'''                For j = 1 To UBound(gArrEquip)
'''                    If gArrEquip(j, 2) = lsEquipCode Then
'''
'''                        SetText vasExam, lsRes, gRow, gResCol + j
'''
''''                        If gArrExamRes(liEquipCode).RefFlag = "H" Then
''''                            SetForeColor vasList, lRow, lRow, gResCol + j, gResCol + j, 255, 127, 0
''''                        ElseIf gArrExamRes(liEquipCode).RefFlag = "L" Then
''''                            SetForeColor vasList, lRow, lRow, gResCol + j, gResCol + j, 0, 127, 255
''''                        Else
''''                            SetForeColor vasList, lRow, lRow, gResCol + j, gResCol + j, 0, 0, 0
''''                        End If
'''
''''                        Save_Local_One lRow, i, "A"
'''
'''                        SQL = "Delete FROM pat_res " & vbCrLf & _
'''                              "WHERE examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'''                              "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'''                              "  AND equipcode = '" & lsEquipCode & "'" & vbCrLf & _
'''                              "  AND barcode = '" & Trim(GetText(vasExam, lRow, 2)) & "' "
'''                        res = SendQuery(gLocal, SQL)
'''                        If res = -1 Then
'''                            SaveQuery SQL
'''                            'Exit Function
'''                        End If
'''
'''                        SQL = "INSERT INTO pat_res (examdate, equipno, " & _
'''                                "barcode, sampletype, receno, " & _
'''                                "pid, pname, jumin, page, psex, " & _
'''                                "resdate, seqno, diskno, posno, " & _
'''                                "equipcode, examcode, " & _
'''                                "result, sendflag, examname, " & _
'''                                "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
'''                              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'''                              "'" & Trim(GetText(vasExam, lRow, 2)) & "','" & Trim(GetText(vasExam, lRow, 3)) & "', '" & Trim(GetText(vasExam, lRow, 4)) & "', " & _
'''                              "'" & Trim(GetText(vasExam, lRow, 5)) & "', '" & Trim(GetText(vasExam, lRow, 6)) & "', '" & Trim(GetText(vasExam, lRow, 9)) & "', 0, '" & Trim(GetText(vasExam, lRow, 7)) & "', " & _
'''                              "'" & lsResDate & "', '" & lsSeqNo & "', '" & Trim(GetText(vasExam, lRow, 6)) & "', '" & Trim(GetText(vasExam, lRow, 7)) & "', " & vbCrLf & _
'''                              "'" & lsEquipCode & "', '" & lsExamCode & "', " & _
'''                              "'" & lsRes & "', 'B', '" & lsExamName & "', " & vbCrLf & _
'''                              "'', '', '', '', " & _
'''                              "'', '' ) "
'''                        res = SendQuery(gLocal, SQL)
'''                        If res = -1 Then
'''                            SaveQuery SQL
'''                            'Exit Function
'''                        End If
'''
'''                        If chkMode.Value = 1 Then        'Auto
'''                            liRet = -1
'''
'''                            If Trim(GetText(vasExam, lRow, colPID)) <> "" Then
'''                                liRet = Insert_Data(lRow)
'''                            End If
'''
'''                            If liRet = -1 Then
'''                                SetBackColor vasExam, lRow, lRow, colCheckBox, colCheckBox, 255, 0, 0
'''                                SetText vasExam, "실패", lRow, colState
'''                            Else
'''                                SetBackColor vasExam, lRow, lRow, 1, colState, 202, 255, 112
'''                                SetText vasExam, "완료", lRow, colState
'''
'''                                'Local 상태를 서버전송(C)로 바꿈
'''                                SQL = " Update pat_res Set " & vbCrLf & _
'''                                      " sendflag = 'C' " & vbCrLf & _
'''                                      " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'''                                      " And equipno = '" & gEquip & "' " & vbCrLf & _
'''                                      " And barcode = '" & Trim(GetText(vasExam, lRow, colBarcode)) & "' " & vbCrLf & _
'''                                      " And equipcode = '" & lsEquipCode & "' "
'''                                res = SendQuery(gLocal, SQL)
'''                                If res = -1 Then
'''                                    SaveQuery SQL
'''                                    Exit Sub
'''                                End If
'''                            End If
'''
''''                        Else
''''                            SetBackColor vasExam, lRow, lRow, 1, 1, 255, 0, 0
''''                            SetText vasExam, "실패", lRow, gResCol
'''                        End If
'''
'''                        Exit For
'''                    End If
'''                Next j
'''
'''            Else
'''                lsRes = ""
'''            End If
'''
'''            PreRow = lRow
'''            PreRack = lsRing
'''            PrePos = lsTube
'''
'''        Case "12"   'Qualitative Raw Data
'''        If lsResType = "T" And lsOrdType = "1" Then
'''            lsQualResRaw = Mid(lsData, 4, 5)
'''            Debug.Print lsQualResRaw
'''            lsTmp = ""
'''            For j = 1 To Len(lsQualResRaw)
'''                If IsNumeric(Mid(lsQualResRaw, j, 1)) Then
'''                    lsTmp = lsTmp & Mid(lsQualResRaw, j, 1)
'''                Else
'''                    lsTmp = lsTmp & "0"
'''                End If
'''
'''            Next j
'''            lsQualResRaw = lsTmp
'''            lsQualResRaw = Left(lsQualResRaw, 2) & "." & Mid(lsQualResRaw, 3)
'''            If IsNumeric(lsQualResRaw) Then
'''                lsQualResRaw = Format(CCur(lsQualResRaw), "#0.000")
'''            End If
'''            If IsNumeric(lsRefHigh) Then
'''                If CCur(lsRefHigh) <= CCur(lsQualResRaw) Then
'''                    lsRes = "Positive"
'''                End If
'''            End If
'''            If IsNumeric(lsRefLow) Then
'''                If CCur(lsRefLow) > CCur(lsQualResRaw) Then
'''                    lsRes = "Negative"
'''                End If
'''            End If
'''            If IsNumeric(lsRefHigh) And IsNumeric(lsRefLow) Then
'''                If CCur(lsRefLow) <= CCur(lsQualResRaw) And CCur(lsRefHigh) > CCur(lsQualResRaw) Then
'''                    lsRes = "Trace"
'''                End If
'''            End If
'''
'''            'Debug.Print lsQualResRaw
'''            'Debug.Print lsRes
'''
'''            For j = 1 To UBound(gArrEquip)
'''                If gArrEquip(j, 2) = lsEquipCode Then
'''
'''                    SetText vasExam, lsRes, lRow, gResCol + j
'''
''''                        If gArrExamRes(liEquipCode).RefFlag = "H" Then
''''                            SetForeColor vasList, lRow, lRow, gResCol + j, gResCol + j, 255, 127, 0
''''                        ElseIf gArrExamRes(liEquipCode).RefFlag = "L" Then
''''                            SetForeColor vasList, lRow, lRow, gResCol + j, gResCol + j, 0, 127, 255
''''                        Else
''''                            SetForeColor vasList, lRow, lRow, gResCol + j, gResCol + j, 0, 0, 0
''''                        End If
'''
''''                        Save_Local_One lRow, i, "A"
'''                    SQL = "Delete FROM pat_res " & vbCrLf & _
'''                          "WHERE examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'''                          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'''                          "  AND equipcode = '" & lsEquipCode & "'" & vbCrLf & _
'''                          "  AND barcode = '" & Trim(GetText(vasExam, lRow, colBarcode)) & "' "
'''                    res = SendQuery(gLocal, SQL)
'''                    If res = -1 Then
'''                        SaveQuery SQL
'''                        'Exit Function
'''                    End If
'''
'''                    SQL = "INSERT INTO pat_res (examdate, equipno, " & _
'''                            "barcode, sampletype, receno, " & _
'''                            "pid, pname, jumin, page, psex, " & _
'''                            "resdate, seqno, diskno, posno, " & _
'''                            "equipcode, examcode, " & _
'''                            "result, sendflag, examname, " & _
'''                            "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
'''                          "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'''                          "'" & Trim(GetText(vasExam, lRow, 2)) & "','" & Trim(GetText(vasExam, lRow, 3)) & "', '" & Trim(GetText(vasExam, lRow, 4)) & "', " & _
'''                          "'" & Trim(GetText(vasExam, lRow, 5)) & "', '" & Trim(GetText(vasExam, lRow, 6)) & "', '" & Trim(GetText(vasExam, lRow, 9)) & "', 0, '" & Trim(GetText(vasExam, lRow, 7)) & "', " & _
'''                          "'" & lsResDate & "', '" & lsSeqNo & "', '" & Trim(GetText(vasExam, lRow, 6)) & "', '" & Trim(GetText(vasExam, lRow, 7)) & "', " & vbCrLf & _
'''                          "'" & lsEquipCode & "', '" & lsExamCode & "', " & _
'''                          "'" & lsRes & "', 'B', '" & lsExamName & "', " & vbCrLf & _
'''                          "'', '', '', '', " & _
'''                          "'', '' ) "
'''                    res = SendQuery(gLocal, SQL)
'''                    If res = -1 Then
'''                        SaveQuery SQL
'''                        'Exit Function
'''                    End If
'''
'''                    If chkMode.Value = 1 Then        'Auto
'''                        liRet = -1
'''
'''                        If Trim(GetText(vasExam, lRow, colPID)) <> "" Then
'''                            liRet = Insert_Data(lRow)
'''                        End If
'''
'''                        If liRet = -1 Then
'''                            SetBackColor vasExam, lRow, lRow, colCheckBox, colCheckBox, 255, 0, 0
'''                            SetText vasExam, "실패", lRow, colState
'''                        Else
'''                            SetBackColor vasExam, lRow, lRow, 1, colState, 202, 255, 112
'''                            SetText vasExam, "완료", lRow, colState
'''
'''                            'Local 상태를 서버전송(C)로 바꿈
'''                            SQL = " Update pat_res Set " & vbCrLf & _
'''                                  " sendflag = 'C' " & vbCrLf & _
'''                                  " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'''                                  " And equipno = '" & gEquip & "' " & vbCrLf & _
'''                                  " And barcode = '" & Trim(GetText(vasExam, lRow, colBarcode)) & "' " & vbCrLf & _
'''                                  " And equipcode = '" & lsEquipCode & "' "
'''                            res = SendQuery(gLocal, SQL)
'''                            If res = -1 Then
'''                                SaveQuery SQL
'''                                Exit Sub
'''                            End If
'''                        End If
'''
''''                    Else
''''                        SetBackColor vasExam, lRow, lRow, 1, 1, 255, 0, 0
''''                        SetText vasExam, "실패", lRow, gResCol
'''                    End If
'''
'''                    Exit For
'''                End If
'''            Next j
'''        End If
'''        Case "13"   'Quantitative Result
'''            lsQuanRes = Trim(Mid(lsData, 4, 10))
'''
'''            lsQuanFlag = Trim(Mid(lsData, 15, 8))
'''
'''            If lsResType <> "T" Then
'''                lsRes = lsQuanRes
'''
'''                If IsNumeric(lsRes) Then
'''                    If Mid(lsQuanFlag, 5, 1) = "8" Then
'''                        lsRes = lsExamName & " : not detected"
'''                    Else
'''                        If IsNumeric(lsResType) Then
'''                            lsRes = Format(lsRes * CCur(lsResType), "#0.0000000")
'''                            lsTmp = ""
'''                            k = 0
'''                            For j = 1 To Len(lsRes)
'''                                If IsNumeric(Mid(lsRes, j, 1)) Then
'''                                    k = k + 1
'''                                    If k > 3 Then
'''                                        lsTmp = lsTmp & "0"
'''                                    Else
'''                                        lsTmp = lsTmp & Mid(lsRes, j, 1)
'''                                    End If
'''                                Else
'''                                    lsTmp = lsTmp & Mid(lsRes, j, 1)
'''                                End If
'''                            Next j
'''                            lsRes = lsTmp
'''                            lsRes = Format(lsRes, "0.00E+00")
'''
'''                            j = InStr(1, lsRes, "E")
'''                            If j > 0 Then
'''                                If Mid(lsRes, j + 1, 1) = "-" Then
'''                                    lsRes = Left(lsRes, j - 1) & "x10^" & Mid(lsRes, j + 1, 1) & CInt(Mid(lsRes, j + 2))
'''                                Else
'''                                    lsRes = Left(lsRes, j - 1) & "x10^" & CInt(Mid(lsRes, j + 2))
'''                                End If
'''                            End If
'''                        End If
'''                    End If
'''                Else
'''                    lsRes = lsExamName & " : not detected"
'''                End If
'''
'''
'''                lRow = -1
'''                For lRow = 1 To vasExam.DataRowCnt
'''                    If Trim(GetText(vasExam, lRow, colRack)) = lsRing And _
'''                       Trim(GetText(vasExam, lRow, colTube)) = lsTube Then
'''                        gRow = lRow
'''                        Exit For
'''                    End If
'''                Next lRow
'''
'''                If lRow = -1 Then
'''                    SQL = "Select barcode, pid, pname, examtype, diskno, posno " & vbCrLf & _
'''                          "from pat_res " & vbCrLf & _
'''                          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'''                          "  and equipno = '" & gEquip & "' " & vbCrLf & _
'''                          "  and diskno = '" & lsRing & "'   " & vbCrLf & _
'''                          "  and posno = '" & lsTube & "' " & vbCrLf & _
'''                          "  and equipcode = '" & lsEquipCode & "' "
'''                    res = db_select_Vas(gLocal, SQL, vasExam, vasExam.DataRowCnt + 1, 2)
'''                    If res = 0 Then
'''                        lRow = vasExam.DataRowCnt + 1
'''                        gRow = lRow
'''
'''                        vasExam.vasExam 2, lRow, lsRing & "-" & lsTube
'''                    Else
'''                        lRow = vasExam.DataRowCnt
'''                        gRow = lRow
'''                    End If
'''                Else
'''                    gRow = lRow
'''                End If
'''
''''                gRow = gRow + 1
'''                lRow = gRow
'''
'''                vasExam.SetText colRack, lRow, lsRing
'''                vasExam.SetText colTube, lRow, lsTube
'''                vasExam.SetText colState, lRow, "수신"
'''
'''                For j = 1 To UBound(gArrEquip)
'''                    If gArrEquip(j, 2) = lsEquipCode Then
'''
'''                        SetText vasExam, lsRes, gRow, gResCol + j
'''
'''
'''                        SQL = "Delete FROM pat_res " & vbCrLf & _
'''                              "WHERE examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'''                              "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'''                              "  AND equipcode = '" & lsEquipCode & "'" & vbCrLf & _
'''                              "  AND barcode = '" & Trim(GetText(vasExam, lRow, 2)) & "' "
'''                        res = SendQuery(gLocal, SQL)
'''                        If res = -1 Then
'''                            SaveQuery SQL
'''                            'Exit Function
'''                        End If
'''
'''                        SQL = "INSERT INTO pat_res (examdate, equipno, " & _
'''                                "barcode, sampletype, receno, " & _
'''                                "pid, pname, jumin, page, psex, " & _
'''                                "resdate, seqno, diskno, posno, " & _
'''                                "equipcode, examcode, " & _
'''                                "result, sendflag, examname, " & _
'''                                "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
'''                              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'''                              "'" & Trim(GetText(vasExam, lRow, 2)) & "','" & Trim(GetText(vasExam, lRow, 3)) & "', '" & Trim(GetText(vasExam, lRow, 4)) & "', " & _
'''                              "'" & Trim(GetText(vasExam, lRow, 5)) & "', '" & Trim(GetText(vasExam, lRow, 6)) & "', '" & Trim(GetText(vasExam, lRow, 9)) & "', 0, '" & Trim(GetText(vasExam, lRow, 7)) & "', " & _
'''                              "'" & lsResDate & "', '" & lsSeqNo & "', '" & Trim(GetText(vasExam, lRow, 10)) & "', '" & Trim(GetText(vasExam, lRow, 11)) & "', " & vbCrLf & _
'''                              "'" & lsEquipCode & "', '" & lsExamCode & "', " & _
'''                              "'" & lsRes & "', 'B', '" & lsExamName & "', " & vbCrLf & _
'''                              "'', '', '', '', " & _
'''                              "'', '' ) "
'''                        res = SendQuery(gLocal, SQL)
'''                        If res = -1 Then
'''                            SaveQuery SQL
'''                            'Exit Function
'''                        End If
'''
'''                        If chkMode.Value = 1 Then        'Auto
'''                            liRet = -1
'''
'''                            If Trim(GetText(vasExam, lRow, colPID)) <> "" Then
'''                                liRet = Insert_Data(lRow)
'''                            End If
'''
'''                            If liRet = -1 Then
'''                                SetBackColor vasExam, lRow, lRow, colCheckBox, colCheckBox, 255, 0, 0
'''                                SetText vasExam, "실패", lRow, colState
'''                            Else
'''                                SetBackColor vasExam, lRow, lRow, 1, colState, 202, 255, 112
'''                                SetText vasExam, "완료", lRow, colState
'''
'''                                'Local 상태를 서버전송(C)로 바꿈
'''                                SQL = " Update pat_res Set " & vbCrLf & _
'''                                      " sendflag = 'C' " & vbCrLf & _
'''                                      " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'''                                      " And equipno = '" & gEquip & "' " & vbCrLf & _
'''                                      " And barcode = '" & Trim(GetText(vasExam, lRow, colBarcode)) & "' " & vbCrLf & _
'''                                      " And equipcode = '" & lsEquipCode & "' "
'''                                res = SendQuery(gLocal, SQL)
'''                                If res = -1 Then
'''                                    SaveQuery SQL
'''                                    Exit Sub
'''                                End If
'''                            End If
''''                        Else
''''                            SetBackColor vasExam, lRow, lRow, 1, 1, 255, 0, 0
''''                            SetText vasExam, "실패", lRow, gResCol
'''                        End If
'''
'''
'''                        Exit For
'''                    End If
'''                Next j
'''
'''            Else
'''                lsRes = ""
'''            End If
'''
'''            PreRow = lRow
'''            PreRack = lsRing
'''            PrePos = lsTube
'''
'''        Case "14"   'Quantitative Raw Data
'''        Case "15"   'Quantitative QS Raw Data
'''        Case "16"   'Quantitative Raw Data ID
'''        Case "17"   'Result Print/Send Status
'''        Case "20"   'Quantitative QS/Control Values
'''        Case "99"   'Result / Order Manipulation Response
'''        'Status processing
'''        Case "00"   'State Selection
'''        Case "41"   'A-ring Load
'''        Case "42"   'Reagent Load
'''        Case "43"   'Cassette Load
'''        Case "90"   'System Status
'''        Case "91"   'TC Status
'''        Case "92"   'DP Status
'''        Case "95"   'File Summary
'''        'Control processing
'''        Case "98"   'Protocol Software Version
'''        Case "99"   'General Response/Error Code
'''        End Select
'''        i = InStr(1, myVar, chrLF)
'''    Loop
End Sub

Sub Save_Raw_Data(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    sFileName = Format(dtpToday, "yyyymmdd")
    
    Open App.Path & "\Log\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

Sub HLC_G7(asData As String)

'''    Dim myVar As String
'''    Dim MyRet As String
'''
'''    Dim i As Integer
'''    Dim ii As Integer
'''    Dim j As Integer
'''    Dim k As Integer
'''    Dim x As Integer
'''    Dim lCol As Long
'''
'''    Dim iRow As Integer
'''    Dim lRow As Long
'''    Dim liRet As Integer
'''
'''    Dim lsSampleNo As String
'''    Dim lsSampleType As String
'''    Dim lsDate As String
'''    Dim lsTime As String
'''    Dim lsID As String
'''    Dim lsID2 As String
'''    Dim lsID1 As String
'''    Dim lsIDInfo As String
'''
'''    Dim lsPID As String
'''    Dim lsPName As String
'''    Dim lsPsex As String
'''
'''    Dim lsData As String
'''
'''    Dim lsTestID As String
'''    Dim lsResult As String
'''    Dim sResult As String
'''    Dim lsFlag As String
'''    Dim lsExamCode As String
'''    Dim lsExamCode1 As String
'''
'''    Dim lsRsCode As String
'''    Dim lsExamName As String
'''
'''    Dim sDate As String
'''    Dim m As Integer
'''    Dim n As Integer
'''
'''    sDate = Format(Text_Today, "yyyymmdd")
'''
'''    lsSampleNo = CStr(CLng(Trim(Mid(asData, 8, 4))))       '순번
'''
'''    lsID = Trim(Mid(asData, 64, 12))               '바코드번호
'''    lsID = Left(lsID, 11)
'''
'''    lRow = -1
'''    For i = 1 To vasExam.DataRowCnt
'''        If lsID <> "" Then  '메뉴얼일경우 바코드 없음
'''            If Trim(GetText(vasExam, i, colBarcode)) = lsID Then
'''                lRow = i
'''                Exit For
'''            End If
'''        End If
'''    Next i
'''
'''    If lRow = -1 Then
'''        lRow = vasExam.DataRowCnt + 1
'''        If vasExam.MaxRows < lRow Then
'''            vasExam.MaxRows = lRow
'''        End If
'''    End If
'''
'''    SetText vasExam, lsSampleNo, lRow, colRack
'''
'''    SetText vasExam, lsID, lRow, colBarcode
'''    'vasExam.SetText colPos, lRow, lsID      '바코드번호를 임시컬럼에 넣기
'''
'''    vasActiveCell vasExam, lRow, colPID
'''
'''    If Len(lsID) = 11 Then
'''        PatInfo vasExam, lsID, lRow
'''    End If
'''
'''    lsTestID = "1"      '장비코드
'''    lsResult = Trim(Mid(asData, 38, 5))
'''    sResult = lsResult
'''    For i = 1 To UBound(gArrEquip)
'''        If Trim(lsTestID) = gArrEquip(i, 2) Then
'''            k = gArrEquip(i, 1)
'''            lCol = (gArrEquip(i, 1) - 1)
'''            lsExamCode = Trim(gArrEquip(i, 3))
'''
'''            SetText vasExam, lsResult, lRow, colResult + lCol * 4
'''            SetText vasExam, lsResult, lRow, colResult1 + lCol
'''        End If
'''    Next i
'''
'''
'''    'If lsSampleType = "P" Then
'''        ClearSpread vasResTemp
'''
'''        '샘플의 검사항목 가져오기
'''        lsID1 = Mid(lsID, 1, 11)
'''        in_spc_no$ = lsID1       '검체번호
'''
'''        rv = sl_Hitache_examdata_select&(in_spc_no$, spc_no$(), tst_cd$(), tst_nm$(), _
'''                                            spc_cd$(), tst_frct_cd$(), tst_frct_nm$(), _
'''                                            tst_dte$(), tst_time$(), work_no$(), pt_no$(), _
'''                                            pt_nm$(), sex$(), birthday$(), intbase$())
'''
'''        If rv < 1 Then
'''            SetForeColor vasExam, lRow, lRow, 1, colState, 255, 0, 0
'''            SetText vasExam, "없음", lRow, colState
'''        Else
'''            If lsSampleType = "P" And Trim(GetText(vasExam, lRow, colPID)) = "" Then
'''                SetText vasExam, pt_no(0), lRow, colPID                     '챠트번호
'''                SetText vasExam, pt_nm(0), lRow, colPName                   '환자성명
'''                SetText vasExam, sex(0), lRow, colPSex                      '성별
''''                        CalAgeSex birthday(0) & "0000000", Trim(Text_Today.Text)
''''                        SetText vasExam, gPatGen.Age, lRow, colPAge                 '나이
'''                SetText vasExam, spc_cd(0), lRow, colSeqNo                  '검체종류
'''                SetText vasExam, tst_dte(0), lRow, colDate
'''                SetText vasExam, tst_time(0), lRow, colTime
'''                SetText vasExam, work_no(0), lRow, colReceno                '접수번호
'''            End If
'''
'''            '처방난 검사항목, 검사명 가져오기
'''            vasResTemp.MaxRows = rv + 1
'''
'''            i = 0
'''
'''            Do While i < rv
'''                SetText vasResTemp, Trim(tst_cd$(i)), i + 1, 1
'''                SetText vasResTemp, Trim(tst_nm$(i)), i + 1, 2
'''
'''                i = i + 1
'''            Loop
'''        End If
'''    'End If
'''
'''    'Local에 저장하기
'''    Save_Local_One_2 lRow, lCol + 1, "A"
'''
'''
'''    If chkMode.Value = 1 Then        'Auto
'''        liRet = -1
'''
'''        If Trim(GetText(vasExam, lRow, colPID)) <> "" Then
'''            liRet = Insert_Data(lRow)
'''        End If
'''
'''        If liRet = -1 Then
'''            SetBackColor vasExam, lRow, lRow, colCheckBox, colCheckBox, 255, 0, 0
'''            SetText vasExam, "실패", lRow, colState
'''        Else
'''            SetBackColor vasExam, lRow, lRow, 1, colState, 202, 255, 112
'''            SetText vasExam, "완료", lRow, colState
'''
'''            'Local 상태를 서버전송(B)로 바꿈
'''            SQL = " Update pat_res Set " & vbCrLf & _
'''                  " sendflag = 'B' " & vbCrLf & _
'''                  " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'''                  " And equipno = '" & gEquip & "' " & vbCrLf & _
'''                  " And barcode = '" & Trim(GetText(vasExam, lRow, colBarcode)) & "' "
'''            res = SendQuery(gLocal, SQL)
'''            If res = -1 Then
'''                SaveQuery SQL
'''                Exit Sub
'''            End If
'''        End If
'''
'''    Else    'Manual
'''        gReadBuf(0) = ""
'''        '수신중========================================================
'''        SetText vasExam, "수신완료", lRow, colState
'''        SetBackColor vasExam, lRow, lRow, 1, 1, 0, 128, 64
'''        '==============================================================
'''    End If
        
End Sub

'''Function Make_Order(argNo As String, argRow As Integer) As String
''''Order Text 만들기...
''''2006/02/07 이상은 - RET 항목은 따로 Order 전송해야 함
'''
'''    'Dim sRetOrder(2) As String     'Order Text넣을 변수
'''    Dim sRetOrder(3) As String
'''
'''    Dim sOrder As String
'''
'''    Dim i As Integer
'''    Dim j As Integer
'''    Dim k As Integer
'''
'''    Dim lsExamCode As String
'''    Dim sExamCode As String
'''    Dim sEquipCode As String
'''    Dim sOrdGubun As String
'''
'''    Dim iCnt_Ord As Integer    'Order 갯수
'''
'''    Dim llRow As Long
'''
'''    Dim in_spc_no$, spc_no$(), tst_cd$(), tst_nm$()
'''    Dim spc_cd$(), tst_frct_cd$(), tst_frct_nm$()
'''    Dim tst_dte$(), tst_time$(), work_no$()
'''    Dim pt_no$(), pt_nm$(), sex$(), birthday$(), intbase$()
'''
'''    Dim acpt_no$()
'''
'''    Dim lsSpcCode As String
'''
'''    Dim rv As Integer
'''    Dim vTemp As String
'''
'''    If argNo = "" Then
'''        Exit Function
'''    End If
'''
'''    Make_Order = -1
'''
'''    '환자정보 및 처방항목 조회
'''    argNo = Mid(argNo, 1, 11)
'''
'''    in_spc_no$ = argNo       '검체번호
'''
'''    rv = sl_Hitache_examdata_select&(in_spc_no$, spc_no$(), tst_cd$(), tst_nm$(), _
'''                                        spc_cd$(), tst_frct_cd$(), tst_frct_nm$(), _
'''                                        tst_dte$(), tst_time$(), work_no$(), pt_no$(), _
'''                                        pt_nm$(), sex$(), birthday$(), intbase$())
'''
'''    If rv < 1 Then
'''        Make_Order = 0
'''
'''        SetForeColor vasExam, argRow, argRow, 1, colState, 255, 0, 0
'''        SetText vasExam, "없음", argRow, colState
'''
'''        'CBC+Diff
'''        sOrder = "11111111" & "1111111111" & _
'''                 "11111" & "00" & "0000000100" & "000000000000000"
'''
'''        Make_Order = sOrder
'''
'''        Exit Function
'''    Else
'''        SetText vasExam, pt_no(0), argRow, colPID                     '챠트번호
'''        SetText vasExam, pt_nm(0), argRow, colPName                   '환자성명
'''        SetText vasExam, sex(0), argRow, colPSex                      '성별
''''        CalAgeSex birthday(0) & "0000000", Trim(Text_Today.Text)
''''        SetText vasexam, gPatGen.Age, argRow, colPAge                 '나이
'''        SetText vasExam, spc_cd(0), argRow, colSeqNo                  '검체종류
'''        SetText vasExam, tst_dte(0), argRow, colDate
'''        SetText vasExam, tst_time(0), argRow, colTime
'''        SetText vasExam, work_no(0), argRow, colReceno                '접수번호
'''
'''        lsSpcCode = spc_cd(0)
'''
'''        '검사항목 관련
'''        ClearSpread vasCode
'''
'''        i = 0
'''        Do While i < rv
'''            '검사항목, 검사명 디스플레이
'''            vasCode.MaxRows = rv + 1
'''
'''            SetText vasCode, Trim(tst_cd$(i)), i + 1, 1
'''            SetText vasCode, Trim(tst_nm$(i)), i + 1, 2
'''
'''            i = i + 1
'''        Loop
'''
'''        'Order 갯수
'''        iCnt_Ord = i - 1
'''        SetText vasExam, CStr(iCnt_Ord), argRow, colOrd
'''
'''        'Order 만들기
'''        For i = 1 To 3
'''            sRetOrder(i) = "0"
'''        Next i
'''
'''        k = 1
'''        Do While k <= vasCode.DataRowCnt
'''            sExamCode = Trim(GetText(vasCode, k, 1))
'''
'''            For j = 1 To UBound(gArrEquip())
'''                If sExamCode = gArrEquip(j, 3) Then
'''                    Select Case gArrEquip(j, 7)
'''                    Case "C"        'CBC
'''                        sRetOrder(1) = "1"
'''                    Case "D"        'Diff
'''                        sRetOrder(2) = "1"
'''                    Case "R"        'Ret
'''                        sRetOrder(3) = "1"
'''                    End Select
'''
'''                    Exit For
'''                End If
'''            Next j
'''
'''            k = k + 1
'''        Loop
'''
'''        sOrder = ""
'''
'''        For i = 1 To 3
'''            sOrder = sOrder & sRetOrder(i)
'''        Next i
'''
'''        '2006.02.06 이상은
'''        'RET%는 Diff임
'''        'NRBC%는 사용 안함
'''        If sOrder <> "" And sOrder = "100" Then          'CBC
'''            sOrder = "11111111" & "0000000000" & _
'''                     "11111" & "00" & "0000000100" & "000000000000000"
'''
'''        ElseIf sOrder <> "" And sOrder = "110" Then      'CBC+Diff
'''            sOrder = "11111111" & "1111111111" & _
'''                     "11111" & "00" & "0000000100" & "000000000000000"
'''        ElseIf sOrder <> "" And sOrder = "111" Then     'CBC+Diff+Ret
'''            sOrder = "11111111" & "1111111111" & _
'''                     "11111" & "00" & "1100000100" & "000000000000000"
'''        End If
'''        Make_Order = sOrder
'''    End If
'''
'''End Function

Function Save_Local_One_2(ByVal asRow1 As Long, ByVal asCol As Long, asSend As String)
'    Dim sCnt As String
'    Dim sExamDate As String
'
'    Dim sEquipCode As String
'    Dim sExamCode As String
'    Dim sExamName As String
'    Dim sResult As String
'
'    sExamDate = GetDateFull
'
'    sEquipCode = ""
'    sExamCode = ""
'    sExamName = ""
'    sResult = ""
'
'
'    sEquipCode = Trim(gArrEquip(asCol, 2))
'    sExamCode = Trim(gArrEquip(asCol, 3))
'    sExamName = Trim(gArrEquip(asCol, 4))
'
'    sResult = Trim(GetText(vasExam, asRow1, colResult + (asCol - 1) * 4))
'
'    'If UCase(Left(Trim(GetText(vasID, asRow1, colPJumin)), 1)) = "F" Then
'    If Trim(GetText(vasID, asRow1, colState)) = "QC" Then
'        Save_Local_QC Trim(Text_Today.Text) & " " & Format(Time, "hh:nn:ss"), _
'                      Trim(GetText(vasExam, asRow1, colBarcode)), _
'                      sEquipCode, _
'                      sResult, _
'                      sResult
'        Exit Function
'    End If
'
'    sCnt = ""
'    If sEquipCode = "" Then Exit Function
'
'    SQL = "select count(*) from pat_res " & vbCrLf & _
'          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'          "  and equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and barcode = '" & Trim(GetText(vasExam, asRow1, colBarcode)) & "' " & vbCrLf & _
'          "  and equipcode = '" & CStr(sEquipCode) & "'"
'
'    res = db_select_Col(gLocal, SQL)
'    sCnt = Trim(gReadBuf(0))
'    If res = -1 Then
'        SaveQuery SQL, 1
'        Exit Function
'    End If
'
'    If Not IsNumeric(sCnt) Then
'        sCnt = "0"
'    End If
'
'    If Not IsNumeric(GetText(vasExam, asRow1, colPAge)) Then
'        SetText vasExam, "0", asRow1, colPAge
'    End If
''    If Not IsDate(Trim(GetText(vasExam, asRow, colExamDate))) Then
''        SetText vasExam, "1900-01-01", asRow, colExamDate
''    End If
'
'    If sCnt = "0" Then
'        SQL = "INSERT INTO pat_res (examdate, equipno, barcode, seqno, diskno, posno, " & _
'              "pid, pname, jumin, page, psex, resdate, receno, " & _
'              "equipcode, examcode, result, result1, sendflag, examname, " & _
'              "refflag, refvalue, panicvalue, recedate ) " & vbCrLf & _
'              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'              "'" & Trim(GetText(vasExam, asRow1, colBarcode)) & "', '" & Trim(GetText(vasExam, asRow1, colSeqNo)) & "'," & _
'              "'" & Trim(GetText(vasExam, asRow1, colRack)) & "', '" & Trim(GetText(vasExam, asRow1, colPos)) & "', " & _
'              "'" & Trim(GetText(vasExam, asRow1, colPID)) & "', " & vbCrLf & _
'              "'" & Trim(GetText(vasExam, asRow1, colPName)) & "', '" & Trim(GetText(vasID, asRow1, colPJumin)) & "', " & _
'              "'" & Trim(GetText(vasExam, asRow1, colPAge)) & "', '" & Trim(GetText(vasExam, asRow1, colPSex)) & "', " & _
'              "'" & sExamDate & "', '" & Trim(GetText(vasExam, asRow1, colReceno)) & "', " & vbCrLf & _
'              "'" & CStr(sEquipCode) & "', '" & sExamCode & "',  " & _
'              "'" & sResult & "', '" & sResult & "', '" & asSend & "', '" & sExamName & "', " & vbCrLf & _
'              "'',  " & _
'              "'" & Trim(GetText(vasExam, asRow1, colOrd)) & "', '" & Trim(GetText(vasExam, asRow1, colRes)) & "', '" & Trim(GetText(vasExam, asRow1, colDate)) & "') "
'
'        res = SendQuery(gLocal, SQL)
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Function
'        End If
'    Else
'        SQL = " Update pat_res Set " & vbCrLf & _
'              " result = '" & sResult & "', " & vbCrLf & _
'              " resdate = '" & sExamDate & "' " & vbCrLf & _
'              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'              " And equipno = '" & gEquip & "' " & vbCrLf & _
'              " And barcode = '" & Trim(GetText(vasExam, asRow1, colBarcode)) & "' " & vbCrLf & _
'              " And equipcode = '" & CStr(sEquipCode) & "' " & vbCrLf & _
'              " And examcode = '" & sExamCode & "' "
'
'        res = SendQuery(gLocal, SQL)
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Function
'        End If
'
'    End If
    
End Function

Function Save_Local_One_1(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
'    Dim sCnt As String
'    Dim sExamDate As String
'
'    sExamDate = GetDateFull
'
'    'If UCase(Left(Trim(GetText(vasID, asRow1, colPJumin)), 1)) = "F" Then
'    If Trim(GetText(vasID, asRow1, colState)) = "QC" Then
'        Save_Local_QC Trim(Text_Today.Text) & " " & Format(Time, "hh:nn:ss"), _
'                      Trim(GetText(vasID, asRow1, colBarcode)), _
'                      Trim(GetText(vasRes, asRow2, colEquipCode)), _
'                      Trim(GetText(vasRes, asRow2, colResult)), _
'                      Trim(GetText(vasRes, asRow2, colResult1))
'        Exit Function
'    End If
'
'    sCnt = ""
'    If Trim(GetText(vasRes, asRow2, colEquipCode)) = "" Then Exit Function
'
'    SQL = "select count(*) from pat_res " & vbCrLf & _
'          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'          "  and equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and barcode = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
'          "  and equipcode = '" & CStr(GetText(vasRes, asRow2, colEquipCode)) & "'"
'
'    res = db_select_Col(gLocal, SQL)
'    sCnt = Trim(gReadBuf(0))
'    If res = -1 Then
'        SaveQuery SQL, 1
'        Exit Function
'    End If
'
'    If Not IsNumeric(sCnt) Then
'        sCnt = "0"
'    End If
'
'    If Not IsNumeric(GetText(vasID, asRow1, colPAge)) Then
'        SetText vasID, "0", asRow1, colPAge
'    End If
''    If Not IsDate(Trim(GetText(vasExam, asRow, colExamDate))) Then
''        SetText vasExam, "1900-01-01", asRow, colExamDate
''    End If
'
'    If sCnt = "0" Then
'        SQL = "INSERT INTO pat_res (examdate, equipno, barcode, seqno, diskno, posno, " & _
'              "pid, pname, jumin, page, psex, resdate, receno, " & _
'              "equipcode, examcode, result, result1, sendflag, examname, " & _
'              "refflag, refvalue, panicvalue, recedate ) " & vbCrLf & _
'              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'              "'" & Trim(GetText(vasID, asRow1, colBarcode)) & "', '" & Trim(GetText(vasID, asRow1, colSeqNo)) & "'," & _
'              "'" & Trim(GetText(vasID, asRow1, colRack)) & "', '" & Trim(GetText(vasID, asRow1, colPos)) & "', " & _
'              "'" & Trim(GetText(vasID, asRow1, colPID)) & "', " & vbCrLf & _
'              "'" & Trim(GetText(vasID, asRow1, colPName)) & "', '" & Trim(GetText(vasID, asRow1, colPJumin)) & "', " & _
'              "'" & Trim(GetText(vasID, asRow1, colPAge)) & "', '" & Trim(GetText(vasID, asRow1, colPSex)) & "', " & _
'              "'" & sExamDate & "', '" & Trim(GetText(vasID, asRow1, colReceno)) & "', " & vbCrLf & _
'              "'" & CStr(GetText(vasRes, asRow2, colEquipCode)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "',  " & _
'              "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', '" & Trim(GetText(vasRes, asRow2, colResult1)) & "', '" & asSend & "', '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
'              "'" & Trim(GetText(vasRes, asRow2, colRCheck)) & "',  " & _
'              "'" & Trim(GetText(vasID, asRow1, colOrd)) & "', '" & Trim(GetText(vasID, asRow1, colRes)) & "', '" & Trim(GetText(vasID, asRow1, colDate)) & "') "
'
'        res = SendQuery(gLocal, SQL)
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Function
'        End If
'    Else
'        SQL = " Update pat_res Set " & vbCrLf & _
'              " result = '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf & _
'              " resdate = '" & sExamDate & "' " & vbCrLf & _
'              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'              " And equipno = '" & gEquip & "' " & vbCrLf & _
'              " And barcode = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
'              " And equipcode = '" & CStr(GetText(vasRes, asRow2, colEquipCode)) & "' " & vbCrLf & _
'              " And examcode = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "' "
'
'        res = SendQuery(gLocal, SQL)
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Function
'        End If
'
'    End If
    
End Function


Function Insert_Data(ByVal argSpcRow As Integer) As Integer
'서버의 데이타 베이스에 저장
    Dim iRow As Integer
    Dim i As Integer
    Dim j As Integer

'    Dim clsResult As New clsResult
    Dim sBarcode, sTstcode As String
    Dim sPID As String
    Dim rc As Integer
    Dim mCnt As Integer
    Dim oerrmsg$
    Dim ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$()
    Dim sSampleSeq As String
    Dim sSampleDate As String
    Dim sChartNo As String
    Dim sPart As String
    Dim sSubCode As String
    Dim sExamCode As String
    Dim sResult As String
    Dim sSendFlag As String
    Dim sExamFlag As Boolean
    Dim sExamDT As String
    Dim sAllRes As String
    Dim sDate As String
    Dim sResultURL As String
    
    Dim sExamDate As String
    Dim sExamTime As String
    
''' -. 환자의 조회된 정보에서 체크 및 전송된 결과에서 결과 체크 후 결과전송
''' -. 전송시 호출 URL : http://10.90.10.228:8090/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetPOCT&Data1=바코드번호1결과수치  &Data2=H10
''' -. 파라미터 설명 : 구분자 (ETX) 와 (ETB) 사용하고 있음. Data2는 장비코드임
'''                    바코드번호(검체번호아님), 결과수치, 접수일시(14자리)만 대입해서 넘겨주시면 됩니다.
'''                    혹시 결과수치에 + 문자 사용된다면 ＋로 대체해야 합니다.
    
    Insert_Data = -1
      
    sBarcode = Trim(GetText(vasExam, argSpcRow, colBarcode))
    sResult = Trim(GetText(vasExam, argSpcRow, colLResult))
    
    'Data1=바코드번호 1 결과수치 접수일시
    
    If Len(sExamDate) <> 8 Then
        sExamDate = Format(Date, "yyyymmdd")
    End If
    If Len(sExamTime) <> 6 Then
        sExamTime = Format(Time, "hhmmss")
    End If
    
    sResultURL = "Data1=" & sBarcode & chrETX & "1" & chrETB & sResult & chrETB & " " & chrETB & " " & chrETB & chrETX & "&Data2=" & gEquip
    
    
    '//// 10.90.10.228:8090 <--- 테스트 서버
    sResultURL = "http://10.20.200.1/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetPOCT&" & sResultURL
    'Local에서 환자별로 결과값 가져오기
    URLstart sResultURL
        
'''    Res = -1
'''    Res = tuxedo_Send.TP_PUT_RESULT("HAMA0101", sAllRes)

    Save_Raw_Data res & "<URL>" & sResultURL
    
    If res = -1 Then
        Insert_Data = -1
        Exit Function
        
    Else
        SQL = " Update pat_res Set " & vbCrLf & _
              " sendflag = '1' " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(sBarcode) & "' "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        Insert_Data = 1
    End If
    
    
    
End Function

Sub Var_Clear()
    gsBarCode = ""
    gsPID = ""
    gsRackNo = ""
    gsPosNo = ""
    gsResDateTime = ""
    gsSeqNo = ""
    gsExamCode = ""
    gsExamName = ""
    gsOrder = ""
    gsResult = ""
End Sub

Private Sub MSComm1_OnComm()
    Dim s As String
    
    s = MSComm1.Input
    
    txtData = txtData & s
    If H232_1 = "1" Then
        If s = chrNACK Then
            Save_Raw_Data "[Rx1 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & chrNACK
            H232_Connect_state = True
            tm_H232.Enabled = False
            txtBuff.Text = ""
            MSComm1.Output = H232_Function(H232_CState)
            Save_Raw_Data "[Tx1 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & H232_Function(H232_CState)
'''            H232_1 = "3"
'''            MSComm1.Output = "a"
            txtBuff = ""
            
        End If
        
    ElseIf H232_1 = "2" Then
        If s = "" Then
            Save_Raw_Data "[Rx1 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & txtBuff
            lblConnect.Caption = "Cobas h232 연결 성공"
            lblConnect.ForeColor = &H808000
            H232_1 = "3"
            MSComm1.Output = chrACK
        ElseIf s = chrVT Then
            MSComm1.Output = chrCR
            
        Else
            txtBuff = txtBuff & s
            
        End If
    
    ElseIf H232_1 = "3" Then
        If s = chrACK Then
            MSComm1.Output = "a"
            txtBuff = ""
        ElseIf s = "a" Then
            MSComm1.Output = chrTAB
            H232_s_1 = "1"
        
        ElseIf s = chrTAB Then
            If H232_s_1 = "1" Then
                MSComm1.Output = Trim(txtReqS.Text)
                H232_s_1 = "2"
            ElseIf H232_s_1 = "2" Then
                MSComm1.Output = Trim(txtReqE.Text)
                H232_s_1 = "3"
                
            ElseIf H232_s_1 = "3" Then
                MSComm1.Output = "0"
                
            End If
        ElseIf s = "0" Then
            MSComm1.Output = vbCr
            H232_1 = "4"
        Else
            txtBuff = txtBuff & s
            If txtBuff = Trim(txtReqS.Text) Or txtBuff = Trim(txtReqE.Text) Then
                MSComm1.Output = chrTAB
            End If
            txtBuff = ""
            
        End If
    
    ElseIf H232_1 = "4" Then
        If s = chrSTX Then
            txtBuff = chrSTX
        ElseIf s = chrEOT Or s = chrETX Then
            txtBuff = txtBuff & s
            Save_Raw_Data "[Rx1 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & txtBuff
            H232 txtBuff.Text, "2"
            
            txtBuff.Text = ""
            MSComm1.Output = chrACK
            Save_Raw_Data "[Tx1 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & chrACK
            If s = chrEOT Then
                If MSComm1.PortOpen = True Then
                    MSComm1.PortOpen = False
                End If
                
            End If
            
        Else
            txtBuff = txtBuff & s
        End If
        
    End If
 
End Sub

Private Sub MSComm2_OnComm()
    Dim s As String
    
    s = MSComm2.Input
    
    txtData = txtData & s
    If H232_2 = "1" Then
        If s = chrNACK Then
            Save_Raw_Data "[Rx2 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & chrNACK
            H232_Connect_state_2 = True
            tm_H232_2.Enabled = False
            txtBuff2.Text = ""
            MSComm2.Output = H232_Function_2(H232_CState)
            Save_Raw_Data "[Tx2 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & H232_Function_2(H232_CState)
'''            H232_2 = "3"
'''            MSComm2.Output = "a"
            txtBuff2 = ""
            
        End If
        
    ElseIf H232_2 = "2" Then
        If s = "" Then
            Save_Raw_Data "[Rx2 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & txtBuff2
            lblConnect2.Caption = "Cobas h232 연결 성공"
            lblConnect2.ForeColor = &H808000
            H232_2 = "3"
            MSComm2.Output = chrACK
        ElseIf s = chrVT Then
            MSComm2.Output = chrCR
            
        Else
            txtBuff2 = txtBuff2 & s
            
        End If
    
    ElseIf H232_2 = "3" Then
        If s = chrACK Then
            MSComm2.Output = "a"
            txtBuff2 = ""
        ElseIf s = "a" Then
            MSComm2.Output = chrTAB
            H232_s_2 = "1"
        
        ElseIf s = chrTAB Then
            If H232_s_2 = "1" Then
                MSComm2.Output = Trim(txtReqS.Text)
                H232_s_2 = "2"
            ElseIf H232_s_2 = "2" Then
                MSComm2.Output = Trim(txtReqE.Text)
                H232_s_2 = "3"
                
            ElseIf H232_s_2 = "3" Then
                MSComm2.Output = "0"
                
            End If
        ElseIf s = "0" Then
            MSComm2.Output = vbCr
            H232_2 = "4"
        Else
            txtBuff2 = txtBuff & s
            If txtBuff2 = Trim(txtReqS.Text) Or txtBuff2 = Trim(txtReqE.Text) Then
                MSComm2.Output = chrTAB
            End If
            txtBuff2 = ""
            
        End If
    
    ElseIf H232_2 = "4" Then
        If s = chrSTX Then
            txtBuff2 = chrSTX
        ElseIf s = chrEOT Or s = chrETX Then
            txtBuff2 = txtBuff2 & s
            Save_Raw_Data "[Rx2 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & txtBuff2
            H232 txtBuff2.Text, "2"
            
            txtBuff2.Text = ""
            MSComm2.Output = chrACK
            Save_Raw_Data "[Tx2 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & chrACK
            If s = chrEOT Then
                If MSComm2.PortOpen = True Then
                    MSComm2.PortOpen = False
                End If
                
            End If
            
        Else
            txtBuff2 = txtBuff2 & s
        End If
        
    End If
    
End Sub

Private Sub Scriptlet1_onscriptletevent(ByVal name As String, ByVal eventData As Variant)

End Sub

Private Sub spErr_Click()
    spErr.Caption = ""
    tmErr.Enabled = False
End Sub

Private Sub Text_Today_GotFocus()
    SelectFocus Text_Today
End Sub

Private Sub Text_Today_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCall_Click
    ElseIf KeyCode = vbKeyF7 Then
'''        frmQCResSch.Show
    End If
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub tm_H232_2_Timer()

    If MSComm2.PortOpen = False Then
        MSComm2.PortOpen = True
    End If
        
    If H232_Connect_state_2 = False Then
        MSComm2.Output = H232_Function_2(H232_Connect)
    End If
End Sub

Private Sub tm_H232_Timer()

    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
        
    If H232_Connect_state = False Then
        MSComm1.Output = H232_Function(H232_Connect)
    End If
    
End Sub

Private Sub tmErr_Timer()
    Beep
    
End Sub

Private Sub tmResRequest_Timer()
    gTimerReq = gTimerReq + 1
    
    dtpToday = Date
    dtpToday_1 = Date - 1
    
    
    If gTimerReq = 3 Then
    
        If MSComm1.PortOpen = False Then
            MSComm1.PortOpen = True
        End If
    
        txtReqS = "1"
        txtReqE = "1"
        lblConnect.Caption = "연결 대기중."
        lblConnect.ForeColor = &HFF&
        MSComm1.Output = H232_Function(H232_Connect)
        Save_Raw_Data "[Tx1 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & H232_Function(H232_Connect)
        TransCheck
        H232_Connect_state = False
        tm_H232.Enabled = True
        gTimerReq = 0
        
'''    ElseIf gTimerReq = 3 Then


'''        If MSComm2.PortOpen = False Then
'''            MSComm2.PortOpen = True
'''        End If
'''
'''        txtReqS = "1"
'''        txtReqE = "1"
'''        lblConnect2.Caption = "연결 대기중."
'''        lblConnect2.ForeColor = &HFF&
'''        MSComm2.Output = H232_Function_2(H232_Connect)
'''        Save_Raw_Data "[Tx2 " & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "]" & H232_Function_2(H232_Connect)
'''        TransCheck
'''        H232_Connect_state_2 = False
'''        tm_H232_2.Enabled = True
    
    End If
    
End Sub


Private Sub defClr()
    SSPanel6.BackColor = &H800000
    spMissResNow.BackColor = &H800000
    spMissResNow.Caption = 0
    spMissResPast.Caption = 0
    SSPanel4.BackColor = &H800000
    spMissResPast.BackColor = &H800000
End Sub

Private Sub TransCheck()
'''    SQL = "select count(*) from pat_res where examdate = '" & Format(Date, "yyyymmdd") & "' and sendflag <> 'C' and receno <> ''"
'''    Res = db_select_Col(gLocal, SQL)
'''
'''    If IsNumeric(Trim(gReadBuf(0))) = True Then
'''        spMissResNow.Caption = Trim(gReadBuf(0))
'''        If Trim(gReadBuf(0)) = "0" Then
'''            SSPanel6.BackColor = &H800000
'''            spMissResNow.BackColor = &H800000
'''        Else
'''            SSPanel6.BackColor = &HFF&
'''            spMissResNow.BackColor = &HFF&
'''        End If
'''
'''    Else
'''        SSPanel6.BackColor = &H800000
'''        spMissResNow.BackColor = &H800000
'''        spMissResNow.Caption = 0
'''    End If
'''
'''    SQL = "select count(*) from pat_res where examdate = '" & Format(Date - 1, "yyyymmdd") & "' and sendflag <> 'C' and receno <> ''"
'''    Res = db_select_Col(gLocal, SQL)
'''
'''    If IsNumeric(Trim(gReadBuf(0))) = True Then
'''        spMissResPast.Caption = Trim(gReadBuf(0))
'''        If Trim(gReadBuf(0)) = "0" Then
'''            SSPanel4.BackColor = &H800000
'''            spMissResPast.BackColor = &H800000
'''        Else
'''            SSPanel4.BackColor = &HFF&
'''            spMissResPast.BackColor = &HFF&
'''        End If
'''    Else
'''
'''        spMissResPast.Caption = 0
'''        SSPanel4.BackColor = &H800000
'''        spMissResPast.BackColor = &H800000
'''    End If
    
End Sub

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sID As String

    If KeyCode = 8 Then
        txtUser.Text = ""
        lblUser.Caption = ""
        lblUser.Caption = ""
        Exit Sub
        
    ElseIf KeyCode = 13 Then
'        sID = Trim(txtUser.Text) & Chr(KeyCode)
'        txtUser.Text = sID
        lblUser.Caption = ""
        SQL = "select user_name from pword where user_id = '" & Trim(txtUser.Text) & "'"
        res = db_select_Col(gServer, SQL)
        If res > 0 Then
            lblUser.Caption = Trim(gReadBuf(0))
        Else
            MsgBox "잘못된 사용자 ID 입니다."
            txtUser.Text = ""
            Exit Sub
        End If
    End If
    
End Sub

Private Sub vasExam_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim iRow As Long
    Dim i As Long
    
    For i = 1 To vasExam.DataRowCnt
        
    Next
    
'''    Dim iRow As Long
'''    Dim iCol As Long
'''    Dim lsNewBarcode As String
'''    Dim lsOldBarcode As String
'''    Dim lsBarcode As String
'''
'''    Dim rv As Integer
'''    Dim i As Long
'''
'''
'''
'''    iRow = Row
'''    iCol = Col
'''
'''    If iCol = colBarcode Then
'''
'''        UserState = False
'''        frmMod.Show 1
'''        If UserState = False Then
'''            Exit Sub
'''        End If
'''
'''
'''        lsOldBarcode = Trim(GetText(vasExam, iRow, colBarcode))
'''        lsNewBarcode = InputBox("변경할 검체번호를 입력하세요.", "검체번호변경")
'''
''''''        SQL = "select barcode from pat_res where barcode = '" & lsNewBarcode & "'"
''''''        res = db_select_Col(gLocal, SQL)
''''''
''''''        If Trim(gReadBuf(0)) = lsNewBarcode Then
''''''            MsgBox "이미 입력된 바코드 번호입니다. "
''''''            Exit Sub
''''''        End If
'''
'''        If Trim(lsNewBarcode) <> "" Then
'''            lsBarcode = Left(lsNewBarcode, 11)
'''            SQL = "p_interfacequery '1', '" & lsBarcode & "'"
'''            res = db_select_Col(gServer, SQL)
'''
'''            If res < 1 Then
'''                SQL = "update pat_res set barcode = '" & lsNewBarcode & "' " & vbCrLf & _
'''                      "where equipno = '" & gEquip & "' and barcode = '" & lsOldBarcode & "' " & vbCrLf & _
'''                      "and examdate = '" & Trim(GetText(vasExam, iRow, ColExamDate)) & "' and examtime = '" & Trim(GetText(vasExam, iRow, ColExamTime)) & "'"
'''                res = SendQuery(gLocal, SQL)
'''
'''            Else
'''                SQL = "update pat_res set barcode = '" & lsNewBarcode & "', pid = '" & Trim(gReadBuf(0)) & "', " & vbCrLf & _
'''                      "recedate = '" & Trim(gReadBuf(9)) & "', receno = '" & Trim(gReadBuf(10)) & "', " & vbCrLf & _
'''                      "seqno = '" & Trim(gReadBuf(11)) & "', pname = '" & Trim(gReadBuf(32)) & "' " & vbCrLf & _
'''                      "where equipno = '" & gEquip & "' and barcode = '" & lsOldBarcode & "' " & vbCrLf & _
'''                      "and examdate = '" & Trim(GetText(vasExam, iRow, ColExamDate)) & "' and examtime = '" & Trim(GetText(vasExam, iRow, ColExamTime)) & "'"
'''                res = SendQuery(gLocal, SQL)
'''
'''                Save_Raw_Data "[SQL]" & SQL
'''
'''                SetText vasExam, gReadBuf(0), iRow, colPID
'''                SetText vasExam, gReadBuf(9), iRow, colReceDate
'''                SetText vasExam, gReadBuf(10), iRow, colReceno
''''''                SetText vasExam, gReadBuf(11), iRow, colSeqNo
'''                SetText vasExam, gReadBuf(32), iRow, colPName
'''                SetText vasExam, lsNewBarcode, iRow, colBarcode
'''
'''            End If
'''
'''        End If
'''    End If
    
End Sub

Private Sub vasExam_KeyDown(KeyCode As Integer, Shift As Integer)
'''    Dim i As Integer
'''    Dim lRow As Long
'''    Dim lCol As Long
'''    Dim lsID As String
'''    Dim sAnalyzerID As String
'''    Dim k As Integer
'''    Dim sEquipCode As String
'''    Dim sExamCode As String
'''    Dim sExamName As String
'''    Dim sSeqNo As String
'''    Dim sResType As String
'''    Dim sResPoint As String
'''    Dim sRefLow As String
'''    Dim sRefHigh As String
'''    Dim sResult As String
'''    Dim sErrData As String
'''    Dim sExamDate As String
'''    Dim sExamTime As String
'''    Dim sBarcode As String
'''
'''    lRow = vasExam.ActiveRow
'''    lCol = vasExam.ActiveCol
'''
'''    If KeyCode = vbKeyReturn Then
'''
'''        UserState = False
'''        frmMod.Show 1
'''        If UserState = False Then
'''            Exit Sub
'''        End If
'''
'''
'''        gReadBuf(0) = ""
'''        sResult = Trim(GetText(vasExam, lRow, ColResult))
'''        sErrData = Trim(GetText(vasExam, lRow, colErrState))
'''        sExamDate = Trim(GetText(vasExam, lRow, ColExamDate))
'''        sExamTime = Trim(GetText(vasExam, lRow, ColExamTime))
'''        sBarcode = Trim(GetText(vasExam, lRow, colBarcode))
'''
'''        If Trim(sResult) = "" Then
'''            SQL = "select barcode from worklist where barcode = '" & sBarcode & "' and resdatetime = '" & sExamTime & "'"
'''            res = db_select_Col(gLocal, SQL)
'''            If Trim(gReadBuf(0)) = sBarcode Then
'''            Else
'''                SQL = "insert into worklist(barcode, ResDateTime) values('" & sBarcode & "', '" & sExamTime & "')"
'''                res = SendQuery(gLocal, SQL)
'''            End If
'''
'''            SQL = "delete from pat_res " & vbCrLf & _
'''                  "where equipno = '" & gEquip & "' and barcode = '" & Trim(GetText(vasExam, lRow, colBarcode)) & "' " & vbCrLf & _
'''                  "and examdate = '" & Trim(GetText(vasExam, lRow, ColExamDate)) & "' and examtime = '" & Trim(GetText(vasExam, lRow, ColExamTime)) & "' " & vbCrLf & _
'''                  "and examcode = '" & Trim(GetText(vasExam, lRow, ColExamCode)) & "'"
'''            res = SendQuery(gLocal, SQL)
'''            DeleteRow vasExam, lRow, lRow
'''
'''        Else
'''            SQL = "update from pat_res set result = '" & sResult & "', bigo = '" & sErrData & "' " & vbCrLf & _
'''                  "where equipno = '" & gEquip & "' and barcode = '" & Trim(GetText(vasExam, lRow, colBarcode)) & "' " & vbCrLf & _
'''                  "and examdate = '" & Trim(GetText(vasExam, lRow, ColExamDate)) & "' and examtime = '" & Trim(GetText(vasExam, lRow, ColExamTime)) & "' " & vbCrLf & _
'''                  "and examcode = '" & Trim(GetText(vasExam, lRow, ColExamCode)) & "'"
'''            res = SendQuery(gLocal, SQL)
'''        End If
'''
'''
'''
'''    End If
End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        vasSort vasID, Col
    End If
    
    If Row < 0 Or Row > vasID.DataRowCnt Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    End If
    
    If Row = 1 Then
        cmdUp.Enabled = False
        cmdDown.Enabled = True
    ElseIf Row = vasID.DataRowCnt Then
        cmdUp.Enabled = True
        cmdDown.Enabled = False
    Else
        cmdUp.Enabled = True
        cmdDown.Enabled = True
    End If
End Sub

Private Sub vasID_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lsCnt As String
    Dim lsID As String
    Dim lsDate As String
    
    Dim iRow As Integer
    
    'cmdCall_Click
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasID, Row, colBarcode))
    
    'Local에서 불러오기
    ClearSpread vasRes

    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "Select a.equipcode, a.examcode, b.examname, a.result, a.result1, b.seqno " & vbCrLf & _
          "from pat_res a, equipexam b " & vbCrLf & _
          "where a.examdate = '" & Format(dtpToday, "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
          "  and b.equipno = a.equipno " & vbCrLf & _
          "  and b.equipcode = a.equipcode " & vbCrLf & _
          "  and b.examcode = a.examcode " & vbCrLf & _
          "order by b.seqno, a.equipcode "
    res = db_select_Vas(gLocal, SQL, vasRes)
    'SaveQuery SQL
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    vasRes.MaxRows = vasRes.DataRowCnt
    'vasSort vasRes, 5, 2
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Integer
    Dim lsID As String
    
    iRow = vasID.ActiveRow
    If KeyCode = vbKeyDelete Then
            
        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        lsID = Trim(GetText(vasID, iRow, colBarcode))
            
        SQL = " Delete From pat_res " & vbCrLf & _
              " Where examdate = '" & Format(dtpToday, "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & lsID & "' "
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
            
        DeleteRow vasID, iRow, iRow
        ClearSpread vasRes
    End If
End Sub

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'Dim iRow As Integer
'Dim lsID As String
'
'    If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'        Exit Sub
'    End If
'
'    iRow = Row
'
'    lsID = Trim(GetText(vasID, iRow, colBarcode))
'
'    SQL = " Delete From pat_res " & vbCrLf & _
'          " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'          " And equipno = '" & gEquip & "' " & vbCrLf & _
'          " And barcode = '" & lsID & "' "
'    res = SendQuery(gLocal, SQL)
'
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
'
'    DeleteRow vasID, iRow, iRow
End Sub

Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim i As Long
    Dim iRow As Long
    
    SetBackColor vasList, 1, vasList.DataRowCnt, colCheckBox, colReady, 255, 255, 255
    If Trim(GetText(vasList, Row, colReady)) = "대기" Then
        SetText vasList, "", Row, colReady
        Exit Sub
    End If
    
    For i = 1 To vasList.DataRowCnt
        If i = Row Then
            SetText vasList, "대기", i, colReady
            SetBackColor vasList, i, i, colCheckBox, colReady, 255, 220, 200
        Else
            SetText vasList, "", i, colReady
        End If
        
    Next
End Sub

Private Sub vasRes_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vasResRow As Long
    Dim vasResCol As Long
    Dim vasIDRow As Long

    Dim lCCR, lM_C_ratio, lP_C_ratio As Long
    Dim sCCR, sCrea_S, sCrea_U, sM_ALB_U, sTP_U As String

    Dim sResult As String
    Dim sResult1 As String

    Dim i As Integer

    Dim sTotalVol As String

'    vasIDRow = vasID.ActiveRow
'    vasResRow = vasRes.ActiveRow
'    vasResCol = vasRes.ActiveCol
'
'    If KeyCode = vbKeyReturn Then
'
'        If vasResCol = colResult Then
'
'            If Trim(GetText(vasRes, vasResRow, colEquipCode)) = "***" Then
'                sTotalVol = Trim(GetText(vasRes, vasResRow, colResult))
'                Save_Local_One_1 vasIDRow, vasResRow, "A"
'
'                If IsNumeric(sTotalVol) Then
'                    lCCR = -1
'                    sCCR = ""
'                    sCrea_S = ""
'                    sCrea_U = ""
'                    sM_ALB_U = ""
'                    sTP_U = ""
'
'                    i = 1
'                    Do While i <= vasRes.DataRowCnt
'                        Select Case Trim(GetText(vasRes, i, colExamCode))
'                        Case "L3117", "L3101", "L3102", "L3103"  'Microalbumun(24hr),Na,K,Cl
'                            sResult = Trim(GetText(vasRes, i, colResult))
'                            SetText vasRes, sResult, i, colResult1
'                            If IsNumeric(sResult) Then
'                                sResult = CCur(sResult) * CCur(sTotalVol) / 1000
'                                SetText vasRes, sResult, i, colResult
'                            End If
'
'                            Save_Local_One_1 vasIDRow, i, "A"
'
'                        Case "L3104", "L3106", "L3107", "L3109" 'Ca,Pi,UA,Protein(24hr)
'                            sResult = Trim(GetText(vasRes, i, colResult))
'                            SetText vasRes, sResult, i, colResult1
'                            If IsNumeric(sResult) Then
'                                sResult = CCur(sResult) * CCur(sTotalVol) / 100
'                                SetText vasRes, sResult, i, colResult
'                            End If
'
'                            Save_Local_One_1 vasIDRow, i, "A"
'                        Case "L31094", "L31095" 'Protein 16hr, 8hr
'                            sResult = Trim(GetText(vasRes, i, colResult))
'                            SetText vasRes, sResult, i, colResult1
'                            If IsNumeric(sResult) Then
'                                sResult = CCur(sResult) * CCur(sTotalVol) / 100
'                                SetText vasRes, sResult, i, colResult
'                            End If
'
'                            Save_Local_One_1 vasIDRow, i, "A"
'                        Case "L31111", "L31112", "L31123" 'Creatinie 16hr, 8hr,24hr
'                            sResult = Trim(GetText(vasRes, i, colResult))
'                            SetText vasRes, sResult, i, colResult1
'                            If IsNumeric(sResult) Then
'                                sResult = CCur(sResult) * CCur(sTotalVol) / 100 / 1000
'                                SetText vasRes, sResult, i, colResult
'                            End If
'
'                            Save_Local_One_1 vasIDRow, i, "A"
'                        Case "L3041"    'Serum Creatinine
'                            sCrea_S = Trim(GetText(vasRes, i, colResult))
'
'                            Save_Local_One_1 vasIDRow, i, "A"
'                        Case "L31121"   'CCR
'                            sCCR = Trim(GetText(vasRes, i, colResult))
'                            lCCR = i
'                        Case "L31171"   'Microalbumin(random)
'                            sM_ALB_U = Trim(GetText(vasRes, i, colResult))
'                        Case "L31110"  'Creatinine(random)
'                            sCrea_U = Trim(GetText(vasRes, i, colResult))
'                        Case "L31090"   'Protein(random)
'                            sTP_U = Trim(GetText(vasRes, i, colResult))
'                        Case "L31172"   'Microalbumin / creatinine (random urine)
'                            lM_C_ratio = i
'                        Case "L31172"   'protein / creatinie (random)
'                            lP_C_ratio = i
'                        End Select
'                        i = i + 1
'                    Loop
'
'                    If lCCR > 0 And lCCR <= vasRes.DataRowCnt And IsNumeric(sCCR) = True And IsNumeric(sCrea_S) = True Then
'                        sResult = CCur(sCCR) * CCur(sTotalVol) / 1440 / CCur(sCrea_S)
'                        SetText vasRes, sResult, lCCR, colResult
'                        Save_Local_One_1 vasIDRow, i, "A"
'                    End If
'
'                    If IsNumeric(sM_ALB_U) = True And IsNumeric(sCrea_U) = True Then
'                        sResult = Format(CCur(sM_ALB_U) / CCur(sCrea_U), "0.00") * 100
'                        If lM_C_ratio > 0 And lM_C_ratio <= vasRes.DataRowCnt Then
'                            SetText vasRes, sResult, lM_C_ratio, colResult
'                        Else
'                            i = vasRes.DataRowCnt + 1
'                            If i > vasRes.maxrows Then
'                                vasRes.maxrows = i
'                            End If
'
'                            SetText vasRes, "101", i, colEquipCode
'                            SetText vasRes, "L31172", i, colExamCode
'                            SetText vasRes, "Microalbumin / Urine Creatinine", i, colExamName
'                            SetText vasRes, sResult, i, colResult
'                            SetText vasRes, sResult, i, colResult1
'                        End If
'
'                        Save_Local_One_1 vasIDRow, i, "A"
'                    End If
'
'                    If IsNumeric(sTP_U) = True And IsNumeric(sCrea_U) = True Then
'                        sResult = Format(CCur(sTP_U) / CCur(sCrea_U), "0.00") * 1000
'                        If lP_C_ratio > 0 And lP_C_ratio <= vasRes.DataRowCnt Then
'                            SetText vasRes, sResult, lM_C_ratio, colResult
'                        Else
'                            i = vasRes.DataRowCnt + 1
'                            If i > vasRes.maxrows Then
'                                vasRes.maxrows = i
'                            End If
'
'                            SetText vasRes, "102", i, colEquipCode
'                            SetText vasRes, "L31201", i, colExamCode
'                            SetText vasRes, "Urine Protein / Urine Creatinine", i, colExamName
'                            SetText vasRes, sResult, i, colResult
'                            SetText vasRes, sResult, i, colResult1
'                        End If
'
'                        Save_Local_One_1 vasIDRow, i, "A"
'                    End If
'                End If
'            Else
'                sResult = Trim(GetText(vasRes, vasResRow, colResult))
'                If MsgBox("저장하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "주의!!!  확인!!!") = vbYes Then
'                    sResult = Trim(GetText(vasRes, vasResRow, colResult))
'
'                    SQL = " update pat_res set " & vbCrLf & _
'                          " Result = '" & sResult & "' " & vbCrLf & _
'                          " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'                          " And equipno = '" & gEquip & "' " & vbCrLf & _
'                          " And barcode = '" & Trim(GetText(vasID, vasIDRow, colBarcode)) & "' " & vbCrLf & _
'                          " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' "
'                    res = SendQuery(gLocal, SQL)
'
'                    If res = -1 Then
'                        SaveQuery SQL
'                        Exit Sub
'                    End If
'
'                    'SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
'
'                End If
'            End If
'
'
'        End If
'    ElseIf KeyCode = vbKeyDelete Then
'        If MsgBox("해당 환자의 " & Trim(GetText(vasRes, vasResRow, colExamName)) & " 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'            Exit Sub
'        End If
'
'        SQL = " Delete From pat_res " & vbCrLf & _
'              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'              " And equipno = '" & gEquip & "' " & vbCrLf & _
'              " And barcode = '" & Trim(GetText(vasID, vasIDRow, colBarcode)) & "' " & vbCrLf & _
'              " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' " & vbCrLf & _
'              " and examcode =  '" & Trim(GetText(vasRes, vasResRow, colExamCode)) & "' "
'        res = SendQuery(gLocal, SQL)
'
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'
'        DeleteRow vasRes, vasResRow, vasResRow
'
'    End If
End Sub

'''Sub Get_Request()
'''    Dim in_spc_no$, spc_no$(), tst_cd$(), tst_nm$()
'''    Dim spc_cd$(), tst_frct_cd$(), tst_frct_nm$()
'''    Dim tst_dte$(), tst_time$(), work_no$()
'''    Dim pt_no$(), pt_nm$(), sex$(), birthday$(), intbase$()
'''
'''    Dim rv As Integer
'''    Dim i As Integer
'''    Dim vTemp As String
'''
'''
'''
'''    in_spc_no$ = Trim(txtBarcode.Text)
'''    rv = sl_Olympus_examdata_select&(in_spc_no$, spc_no$(), tst_cd$(), tst_nm$(), _
'''                                        spc_cd$(), tst_frct_cd$(), tst_frct_nm$(), _
'''                                        tst_dte$(), tst_time$(), work_no$(), pt_no$(), _
'''                                        pt_nm$(), sex$(), birthday$(), intbase$())
'''
'''End Sub

Function Save_Local_QC(asExamDate As String, asBarcode As String, asExamCode As String, asRes1 As String, asRes2 As String)
    Dim sResDateTime As String
    Dim sControl As String
    Dim sLotNo As String
    
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sRefFlag As String
    
    Dim sCnt As String
    
    sResDateTime = Format(CDate(asExamDate), "yyyymmdd hhnnss")
    'sControl = Trim(Left(asBarcode, 2))
    'sLotNo = Trim(Mid(asBarcode, 3))
    sControl = asBarcode
    sRefFlag = ""
    
    SQL = "Select t_mean, t_sd from qcexam " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and validstart >= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and valiend <= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and levelname = '" & sControl & "' " & vbCrLf & _
          "  and equipcode = '" & asExamCode & "' "
    res = db_select_Col(gLocal, SQL)
    If res > 0 Then
        If IsNumeric(gReadBuf(0)) And IsNumeric(gReadBuf(1)) Then
            sRefLow = CCur(gReadBuf(0)) - CCur(gReadBuf(1))
            sRefHigh = CCur(gReadBuf(0)) + CCur(gReadBuf(1))
            If CCur(sRefHigh) < CCur(asRes2) Then
                sRefFlag = "H"
            End If
            If CCur(sRefLow) > CCur(asRes2) Then
                sRefFlag = "L"
            End If
        End If
    End If
    
    sCnt = ""
    SQL = "Select count(*) from qc_res " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and examtime = '" & Mid(sResDateTime, 10, 6) & "' " & vbCrLf & _
          "  and levelname = '" & sControl & "' " & vbCrLf & _
          "  and equipcode = '" & asExamCode & "' "
    res = db_select_Var(gLocal, SQL, sCnt)
    If res <= 0 Then
        SaveQuery SQL
        db_RollBack gLocal
        Exit Function
    End If
    res = db_select_Var(gLocal, SQL, sCnt)
    If res <= 0 Then
        SaveQuery SQL
        Exit Function
    End If
    If Not IsNumeric(sCnt) Then sCnt = "0"
    
    If CInt(sCnt) > 0 Then
        SQL = "delete from qc_res " & vbCrLf & _
              "where equipno = '" & gEquip & "' " & vbCrLf & _
              "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
              "  and examtime = '" & Mid(sResDateTime, 9, 4) & "' " & vbCrLf & _
              "  and levelname = '" & sControl & "' " & vbCrLf & _
              "  and equipcode = '" & asExamCode & "' "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            'db_RollBack gLocal
            SaveQuery SQL
            Exit Function
        End If
    End If
    SQL = "Insert into qc_res (equipno, examdate, examtime, levelname, equipcode, sresult, result, resflag, remark, examuid, lotno) " & vbCrLf & _
          "values ('" & gEquip & "', '" & Left(sResDateTime, 8) & "', '" & Mid(sResDateTime, 10, 4) & "', '" & sControl & "', '" & asExamCode & "', '" & asRes1 & "', '" & asRes2 & "', '" & sRefFlag & "','','', '" & sLotNo & "') "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        'db_RollBack gLocal
        SaveQuery SQL
        Exit Function
    End If
    
End Function



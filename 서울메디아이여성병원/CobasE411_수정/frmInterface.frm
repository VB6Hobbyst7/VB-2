VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Interface Program"
   ClientHeight    =   10440
   ClientLeft      =   240
   ClientTop       =   750
   ClientWidth     =   15225
   FillColor       =   &H0000FFFF&
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
   MaxButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   15225
   Begin TabDlg.SSTab SSTab1 
      Height          =   9615
      Left            =   90
      TabIndex        =   10
      Top             =   750
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   16960
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Interface"
      TabPicture(0)   =   "frmInterface.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "결과확인"
      TabPicture(1)   =   "frmInterface.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   9120
         Left            =   150
         TabIndex        =   24
         Top             =   330
         Width           =   14760
         Begin FPSpread.vaSpread vasResTemp 
            Height          =   2355
            Left            =   420
            TabIndex        =   38
            Top             =   6120
            Visible         =   0   'False
            Width           =   11265
            _Version        =   393216
            _ExtentX        =   19870
            _ExtentY        =   4154
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
            SpreadDesigner  =   "frmInterface.frx":047A
         End
         Begin VB.CommandButton cmdVasListWidth 
            Caption         =   ">>"
            Height          =   405
            Left            =   210
            TabIndex        =   37
            Top             =   1110
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.CheckBox ChkAll 
            Height          =   255
            Left            =   720
            TabIndex        =   35
            Top             =   1170
            Width           =   225
         End
         Begin VB.Frame Frame4 
            Caption         =   "[검사결과조회]"
            Height          =   735
            Left            =   180
            TabIndex        =   25
            Top             =   210
            Width           =   14385
            Begin VB.TextBox txtBarcode 
               Height          =   315
               Left            =   11760
               TabIndex        =   30
               Top             =   270
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.ComboBox cmbTransGubun 
               Height          =   315
               ItemData        =   "frmInterface.frx":06BE
               Left            =   3330
               List            =   "frmInterface.frx":06CB
               TabIndex        =   29
               Text            =   "전체"
               Top             =   270
               Width           =   1395
            End
            Begin VB.CommandButton cmdCall 
               Caption         =   "데이터 불러오기"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   4860
               TabIndex        =   28
               Top             =   210
               Width           =   1815
            End
            Begin VB.CommandButton cmdListClear 
               Caption         =   "화면초기화"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   6720
               TabIndex        =   27
               Top             =   210
               Width           =   1275
            End
            Begin VB.CommandButton cmdListTrans 
               Caption         =   "검사결과수동전송"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   8040
               TabIndex        =   26
               Top             =   210
               Width           =   1905
            End
            Begin MSComCtl2.DTPicker dtpExamDate 
               Height          =   315
               Left            =   1110
               TabIndex        =   31
               Top             =   270
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               _Version        =   393216
               Format          =   21299201
               CurrentDate     =   40780
            End
            Begin VB.Label Label4 
               Caption         =   "Barcode 검색"
               Height          =   225
               Left            =   10380
               TabIndex        =   34
               Top             =   330
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label Label2 
               Caption         =   "검사일자"
               Height          =   225
               Left            =   180
               TabIndex        =   33
               Top             =   330
               Width           =   915
            End
            Begin VB.Label Label3 
               Caption         =   "구분"
               Height          =   225
               Left            =   2820
               TabIndex        =   32
               Top             =   330
               Width           =   555
            End
         End
         Begin FPSpread.vaSpread vasList 
            Height          =   7875
            Left            =   180
            TabIndex        =   49
            Top             =   1080
            Width           =   14385
            _Version        =   393216
            _ExtentX        =   25374
            _ExtentY        =   13891
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            ColsFrozen      =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridColor       =   16777215
            MaxCols         =   103
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":06E3
         End
         Begin FPSpread.vaSpread vasListRes 
            Height          =   7875
            Left            =   9030
            TabIndex        =   42
            Top             =   1080
            Width           =   5535
            _Version        =   393216
            _ExtentX        =   9763
            _ExtentY        =   13891
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
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
            MaxCols         =   9
            MaxRows         =   100
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmInterface.frx":3B93
         End
      End
      Begin VB.Frame Frame3 
         Height          =   9120
         Left            =   -74850
         TabIndex        =   16
         Top             =   360
         Width           =   14760
         Begin VB.CommandButton Command4 
            Caption         =   "Command4"
            Height          =   585
            Left            =   1890
            TabIndex        =   50
            Top             =   5910
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   525
            Left            =   7530
            TabIndex        =   47
            Top             =   -120
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   1155
            Left            =   12690
            TabIndex        =   22
            Top             =   7800
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Height          =   1125
            Left            =   7200
            TabIndex        =   21
            Top             =   7860
            Visible         =   0   'False
            Width           =   5475
         End
         Begin FPSpread.vaSpread vasTMaxRes 
            Height          =   2235
            Left            =   840
            TabIndex        =   46
            Top             =   6630
            Visible         =   0   'False
            Width           =   5625
            _Version        =   393216
            _ExtentX        =   9922
            _ExtentY        =   3942
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
            SpreadDesigner  =   "frmInterface.frx":4668
         End
         Begin FPSpread.vaSpread vasTMaxList 
            Height          =   2805
            Left            =   7200
            TabIndex        =   45
            Top             =   3810
            Visible         =   0   'False
            Width           =   3525
            _Version        =   393216
            _ExtentX        =   6218
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
            SpreadDesigner  =   "frmInterface.frx":48AC
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   315
            Left            =   8040
            TabIndex        =   44
            Top             =   330
            Visible         =   0   'False
            Width           =   1695
         End
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   1155
            Left            =   7200
            TabIndex        =   43
            Top             =   4230
            Visible         =   0   'False
            Width           =   4065
            _Version        =   393216
            _ExtentX        =   7170
            _ExtentY        =   2037
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
            SpreadDesigner  =   "frmInterface.frx":4AF0
         End
         Begin VB.TextBox txtData 
            Height          =   1215
            Left            =   11580
            TabIndex        =   41
            Top             =   6600
            Visible         =   0   'False
            Width           =   2715
         End
         Begin FPSpread.vaSpread vasOrderBuf 
            Height          =   1215
            Left            =   7200
            TabIndex        =   40
            Top             =   6600
            Visible         =   0   'False
            Width           =   4395
            _Version        =   393216
            _ExtentX        =   7752
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
            SpreadDesigner  =   "frmInterface.frx":4D34
         End
         Begin FPSpread.vaSpread vasOrder 
            Height          =   1215
            Left            =   7200
            TabIndex        =   39
            Top             =   5400
            Visible         =   0   'False
            Width           =   4395
            _Version        =   393216
            _ExtentX        =   7752
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
            SpreadDesigner  =   "frmInterface.frx":91FA
         End
         Begin VB.CommandButton cmdVasIDWidth 
            Caption         =   ">>"
            Height          =   405
            Left            =   210
            TabIndex        =   36
            Top             =   810
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox txtBuff 
            Height          =   1215
            Left            =   11580
            TabIndex        =   20
            Top             =   5400
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "화면초기화"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   180
            TabIndex        =   19
            Top             =   270
            Width           =   1425
         End
         Begin VB.CommandButton cmd_Trans 
            Caption         =   "검사결과수동전송"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1680
            TabIndex        =   18
            Top             =   270
            Width           =   2085
         End
         Begin VB.CheckBox chkA 
            Height          =   255
            Left            =   720
            TabIndex        =   17
            Top             =   870
            Width           =   225
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   8175
            Left            =   180
            TabIndex        =   23
            Top             =   780
            Width           =   14385
            _Version        =   393216
            _ExtentX        =   25374
            _ExtentY        =   14420
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            ColsFrozen      =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridColor       =   16777215
            MaxCols         =   103
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":D6C0
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   8175
            Left            =   9030
            TabIndex        =   48
            Top             =   780
            Width           =   5535
            _Version        =   393216
            _ExtentX        =   9763
            _ExtentY        =   14420
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
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
            MaxCols         =   9
            MaxRows         =   100
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmInterface.frx":10B70
         End
      End
   End
   Begin Threed.SSPanel sspMode 
      Height          =   525
      Left            =   2040
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   926
      _Version        =   131072
      ForeColor       =   16777215
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   10.5
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "전송모드"
      BevelWidth      =   3
      BorderWidth     =   5
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2730
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   4096
      InputLen        =   1
      RThreshold      =   1
      EOFEnable       =   -1  'True
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   555
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   979
      _Version        =   131072
      ForeColor       =   16777215
      BackColor       =   11494691
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "INTERFACE"
      BevelOuter      =   0
      Alignment       =   1
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   3420
         Top             =   60
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   6420
         Top             =   60
      End
      Begin VB.TextBox txtUID 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Enabled         =   0   'False
         Height          =   270
         Left            =   9000
         TabIndex        =   15
         Top             =   180
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8640
         Picture         =   "frmInterface.frx":11645
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   14
         Top             =   180
         Width           =   315
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   11760
         TabIndex        =   13
         Top             =   120
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63242240
         CurrentDate     =   40778
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
         Left            =   10740
         TabIndex        =   12
         Top             =   210
         Width           =   900
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '투명
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
         Height          =   225
         Left            =   5310
         TabIndex        =   11
         Top             =   210
         Visible         =   0   'False
         Width           =   1725
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   6585
      Left            =   -840
      TabIndex        =   2
      Top             =   4050
      Visible         =   0   'False
      Width           =   8835
      Begin VB.TextBox txtMsg 
         ForeColor       =   &H000000C0&
         Height          =   825
         Left            =   7830
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   9
         Top             =   3300
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtErr 
         Height          =   1035
         Left            =   4440
         TabIndex        =   8
         Top             =   5100
         Width           =   1935
      End
      Begin VB.TextBox txtDate 
         Height          =   405
         Left            =   330
         TabIndex        =   5
         Top             =   1260
         Width           =   2325
      End
      Begin VB.TextBox txtAll 
         Height          =   375
         Left            =   300
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   870
         Width           =   2055
      End
      Begin VB.TextBox txtTemp 
         Height          =   375
         Left            =   300
         TabIndex        =   3
         Top             =   450
         Width           =   2055
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   4455
         Left            =   4020
         TabIndex        =   6
         Top             =   0
         Width           =   3555
         _Version        =   393216
         _ExtentX        =   6271
         _ExtentY        =   7858
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmInterface.frx":11BCF
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   2535
         Left            =   150
         TabIndex        =   7
         Top             =   2130
         Visible         =   0   'False
         Width           =   3555
         _Version        =   393216
         _ExtentX        =   6271
         _ExtentY        =   4471
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
         SpreadDesigner  =   "frmInterface.frx":160F9
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "메인"
      Begin VB.Menu mnuExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnuConf 
      Caption         =   "설정"
      Begin VB.Menu mnuCodeConfig 
         Caption         =   "코드설정"
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "통신설정"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "전송"
      Begin VB.Menu mnuAuto 
         Caption         =   "자동전송"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuManual 
         Caption         =   "수동전송"
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "pop"
      Visible         =   0   'False
      Begin VB.Menu subUp 
         Caption         =   "검체번호 수정"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu subDel 
         Caption         =   "검체결과 삭제"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const colCheckBox = 1
Const colBarCode = 2
Const colPID = 3
Const colPName = 4
Const colPSex = 5
Const colPAge = 6
Const colReceDate = 7
Const colRackPos = 8

'''Const colRcnt = 5
Const colState = 9
Const colRStart = 9

' 장비코드 검사코드 검사명 수치결과 문자결과 seq
Const colEquipExam = 1
Const colExamCode = 2
Const colExamName = 3
Const colResValue = 4
Const colResult = 5
Const colSeq = 6
Const colResDate = 7
Const colResTime = 8
Const colRef = 9

Public gRow As Long
Dim sOrder As String
Dim ConfirmData As String
Dim sSampleType As String
Dim lsFlag As String
Dim llRow As Long



Private Sub chkA_Click()
    Dim iRow As Integer
    
    If chkA.Value = 1 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 1
        Next iRow
    ElseIf chkA.Value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 0
        Next iRow
    End If
End Sub

Private Sub ChkAll_Click()
    Dim iRow As Integer
    
    If ChkAll.Value = 1 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            
            vasList.Value = 1
        Next iRow
    ElseIf ChkAll.Value = 0 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            
            vasList.Value = 0
        Next iRow
    End If
End Sub

Private Sub cmd_Trans_Click()
'선택전송
Dim VasidRow As Integer
Dim VasResRow As Integer
Dim iRow As Integer
Dim liRet As Integer
Dim iNumber As Integer

    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:선택전송") = vbCancel Then
        Exit Sub
    End If

    If vasID.DataRowCnt < 1 Then
        MsgBox "저장할 데이터가 없습니다."
        Exit Sub
    End If
    
    'db_BeginTran gServer
'''    Connect_Server
    For VasidRow = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = VasidRow
        
        If vasID.Value = 1 Then '체크된 열은 저장이 안됨
'        If vasID.Value = "" Then
        
            liRet = -1
            
    
            liRet = Insert_Data(VasidRow, vasID)

            
            If liRet = 1 Then
                'db_Commit gServer
                
                SetBackColor vasID, VasidRow, VasidRow, colCheckBox, colCheckBox, 202, 255, 112
                SetText vasID, "완료", VasidRow, colState
            Else
                SetBackColor vasID, VasidRow, VasidRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasID, "실패", VasidRow, colState
            End If
            vasID.Col = 1
            vasID.Row = VasidRow
            vasID.Value = 0
        Else
        
        End If
    Next VasidRow
    
End Sub

Function Insert_Data(argSpcRow As Integer, argSpread As vaSpread) As Integer
    Dim iRow        As Integer
    Dim i           As Integer
    Dim j           As Integer
    Dim lsID        As String
    Dim sResult     As String
    Dim sResult1    As String       '수치결과
    Dim sResult2    As String       '문자결과
    
    Dim iPos        As Integer
    Dim iPos1       As Integer
    Dim sORD_CD     As String
    Dim sSPCCD      As String
    Dim sSEQ_NO     As String
    
    Dim sDecision   As String
    Dim sPanicFlag  As String
    Dim sDeltaFlag  As String
    Dim sDPA_GB     As String       'DELTA/PANIC 동시 발생시 ('DP'로 변경)
    
    Dim sCnt        As String
    
    Dim sResultCD   As String
    Dim sAllResult  As String
    Dim sEquipCode  As String
    Dim sReceCode   As String
    Dim sTransDate As String
    Dim sTransTime As String
    
    Dim sRsltSqno As String
    Dim sResValue As String
    Dim sRcpnSqno As String
    Dim sExamCode As String
    Dim sResGubun As String
    Dim sTransRes As String
    Dim sTmaxStr As String
    Dim llRow As Long
    Dim sSmyr As String
    Dim sSmsn As String
    Dim sSms1 As String
    Dim sReceDate As String
    Dim sSlip As String
    Dim sWkno As String
    Dim sMachNo As String
    Dim sResFlag As Boolean
    Dim sRefFlag As String
    
    Dim strTmaxRes As String
    Dim iCol As Long
    Dim strMsg As String
    Dim lsReceDate As String
    Dim lsReceNo As String
    
    Dim strResState As String
    Dim smBarcode As String
'''    Dim sExamCode As String
    Dim smSpcCode As String
    Dim smReceDate As String
    Dim smPID As String
    Dim smPName As String
    Dim smPIO As String
    Dim smPSex As String
    Dim smPAge As String
    Dim smResGB As String
    Dim smNumRes As String
    Dim smStrRes As String
    Dim smSTate As String
    Dim strExamDate As String
    Dim strExamTime As String
    Dim Ret As Long
    Dim FilNum
    Dim sFileName As String
    Dim FindFile As String
    
    
    
    Insert_Data = -1
    
    
    
    sResFlag = False
    lsID = ""
    lsID = Trim(GetText(argSpread, argSpcRow, colBarCode))
    strExamDate = Trim(GetText(argSpread, argSpcRow, colPID))
    strExamTime = Trim(GetText(argSpread, argSpcRow, colReceDate))

    If IsNumeric(lsID) = False Then Exit Function
    If Len(lsID) < 8 Then Exit Function
    'Local에서 환자별로 결과값 가져오기
    ClearSpread vasResTemp
    
    SQL = " Select equipcode, examcode, resvalue, result, refflag " & vbCrLf & _
          " From pat_res  " & vbCrLf & _
          " Where barcode = '" & lsID & "' and resvalue <> '' and pid = '" & strExamDate & "' and recedate = '" & strExamTime & "' "

    res = db_select_Vas(gLocal, SQL, vasResTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
        
    For j = 1 To vasResTemp.DataRowCnt
        sExamCode = Trim(GetText(vasResTemp, j, 2))
        sResValue = Trim(GetText(vasResTemp, j, 3))
        sResult = Trim(GetText(vasResTemp, j, 4))
        sRefFlag = Trim(GetText(vasResTemp, j, 5))
        
        SQL = "select resgubun from equipexam where equipcode = '" & Trim(GetText(vasResTemp, j, 1)) & "'"
        res = db_select_Col(gLocal, SQL)
        
        sResGubun = Trim(gReadBuf(0))
        
        If sResGubun = "1" Then '문자
            sTransRes = sResult & "(" & sResValue & ")"
            
        Else
            sTransRes = sResValue
           
        End If
        
'''        pifrslip char(6)   슬립코드
'''        pifrlbno char(12)  차트번호
'''        pifrcode char(6)   수가코드
'''        pifrsubc char(6)   수가세부코드
'''        pifrrslt char(25)  결과값
'''        pifrcpmv Char(10)  '999999'
'''        pifrendt char(8)   결과일
'''        pifrentm char(6)   결과시간
        gReadBuf(0) = ""
        
        SQL = "SELECT pifrrslt FROM PIFRSLTM.db "
        SQL = SQL & vbCrLf & " WHERE pifrlbno = '" & lsID & "' "
        SQL = SQL & vbCrLf & "   AND pifrcode = '" & sExamCode & "' "
        SQL = SQL & vbCrLf & "   AND pifrendt = '" & strExamDate & "' "
        SQL = SQL & vbCrLf & "   AND pifrentm = '" & strExamTime & "' "

        res = db_select_Col(gServer, SQL)
        
        
        If res = 0 Then
            SQL = "INSERT INTO PIFRSLTM.db (pifrslip, pifrlbno, pifrcode, pifrsubc, pifrrslt, pifrcpmv, pifrendt, pifrentm) "
            SQL = SQL & vbCrLf & " values ('F016A', '" & lsID & "', '" & sExamCode & "', '', "
            SQL = SQL & vbCrLf & "        '" & sTransRes & "', '999999', '" & strExamDate & "', '" & strExamTime & "')"
            'res = SendQuery(gServer, SQL)
            
            'Ret = WinExec("rcp.exe -a h01.dat med.lab:/usr/tmp/h01.dat", 2)
            
            'argSQL의 내용을 파일로 저장
            FilNum = FreeFile
            
            FindFile = Dir("c:\insert.sql")
            If FindFile <> "" Then
                Kill "c:\insert.sql"     '전송완료가 됐을때 파일지우기
            End If
            
            Open "c:\insert.sql" For Append As FilNum
            
            Print #FilNum, SQL
            Close FilNum
    
            Ret = WinExec("C:\RS232\execbde.exe C:\insert.sql c:\err.txt", 2)
            
        Else
            SQL = "UPDATE PIFRSLTM.db SET pifrrslt = '" & sTransRes & "' "
            SQL = SQL & vbCrLf & " WHERE pifrlbno = '" & lsID & "' "
            SQL = SQL & vbCrLf & "   AND pifrcode = '" & sExamCode & "' "
            SQL = SQL & vbCrLf & "   AND pifrendt = '" & strExamDate & "' "
            SQL = SQL & vbCrLf & "   AND pifrentm = '" & strExamTime & "' "
            'res = SendQuery(gServer, SQL)
        
            'argSQL의 내용을 파일로 저장
            FilNum = FreeFile
            
            FindFile = Dir("c:\update.sql")
            If FindFile <> "" Then
                Kill "c:\update.sql"     '전송완료가 됐을때 파일지우기
            End If
            
            Open "c:\update.sql" For Append As FilNum
            
            Print #FilNum, SQL
            Close FilNum
    
            Ret = WinExec("C:\RS232\execbde.exe C:\update.sql c:\err.txt", 2)
        
        End If
        
        
        
'''        SQL = "UPDATE LRESULT  "
'''        SQL = SQL & vbCrLf & "     SET RSFL = 'Y' "
'''        SQL = SQL & vbCrLf & "     , RSLT = '" & sTransRes & "'"
'''        If sRefFlag = "H" Or sRefFlag = "L" Then
'''            SQL = SQL & vbCrLf & "     , HLFL = '" & sRefFlag & "'"
'''        Else
'''        End If
'''        SQL = SQL & vbCrLf & "     , RSDT = sysdate "
'''        SQL = SQL & vbCrLf & "     , RSID = '" & gUserID & "' "
'''        SQL = SQL & vbCrLf & "   WHERE SPNO = '" & lsID & "' "
'''        SQL = SQL & vbCrLf & "     AND ORCD = '" & sExamCode & "' "
'''        SQL = SQL & vbCrLf & "     AND OKFL = 'N' "
'''
'''        res = SendQuery(gServer, SQL)
        
    Next
    
    SQL = "update pat_res " & vbCrLf & _
          " set sendflag = '2' " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(argSpread, argSpcRow, colBarCode)) & "' "
    res = SendQuery(gLocal, SQL)
                    
    Insert_Data = 1
    
End Function


Function Insert_Data_1(argSpcRow As Integer) As Integer
    Dim iRow        As Integer
    Dim i           As Integer
    Dim j           As Integer
    Dim lsID        As String
    Dim sResult     As String
    Dim sResult1    As String       '수치결과
    Dim sResult2    As String       '문자결과
    
    Dim iPos        As Integer
    Dim iPos1       As Integer
    Dim sORD_CD     As String
    Dim sSPCCD      As String
    Dim sSEQ_NO     As String
    
    Dim sDecision   As String
    Dim sPanicFlag  As String
    Dim sDeltaFlag  As String
    Dim sDPA_GB     As String       'DELTA/PANIC 동시 발생시 ('DP'로 변경)
    
    Dim sCnt        As String
    
    Dim sResultCD   As String
    Dim sAllResult  As String
    Dim sEquipCode  As String
    Dim sReceCode   As String
    Dim sTransDate As String
    Dim sTransTime As String
    
    Dim sRsltSqno As String
    Dim sResValue As String
    Dim sRcpnSqno As String
    Dim sExamCode As String
    Dim sResGubun As String
    Dim sTransRes As String
    Dim sTmaxStr As String
    Dim llRow As Long
    Dim sSmyr As String
    Dim sSmsn As String
    Dim sSms1 As String
    Dim sReceDate As String
    Dim sSlip As String
    Dim sWkno As String
    Dim sMachNo As String
    Dim sResFlag As Boolean
    Dim sRefFlag As String
    
    Dim strTmaxRes As String
    Dim iCol As Long
    Dim strMsg As String
    Dim lsReceDate As String
    Dim lsReceNo As String
    
    
    Insert_Data_1 = -1
    
    
    sResFlag = False
    lsID = ""
    lsID = Trim(GetText(vasList, argSpcRow, colBarCode))
    
'''    If Len(lsID) = 8 And IsNumeric(lsID) = True Then
'''    Else
'''        Exit Function
'''    End If
'''
'''    lsReceDate = Format(Date, "yyyy") & Mid(lsID, 1, 4)
'''    lsReceNo = Mid(lsID, 5)
'''    lsReceNo = CStr(CCur(lsReceNo))
   
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread vasResTemp
    
    SQL = " Select equipcode, examcode, resvalue, result, refflag " & vbCrLf & _
          " From pat_res  " & vbCrLf & _
          " Where barcode = '" & lsID & "' and resvalue <> '' "

    res = db_select_Vas(gLocal, SQL, vasResTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
        
    For j = 1 To vasResTemp.DataRowCnt
        sExamCode = Trim(GetText(vasResTemp, j, 2))
        sResValue = Trim(GetText(vasResTemp, j, 3))
        sResult = Trim(GetText(vasResTemp, j, 4))
        sRefFlag = Trim(GetText(vasResTemp, j, 5))
        
        SQL = "select resgubun from equipexam where equipcode = '" & Trim(GetText(vasResTemp, j, 1)) & "'"
        res = db_select_Col(gLocal, SQL)
        
        sResGubun = Trim(gReadBuf(0))
        
        If sResGubun = "1" Then '문자
            sTransRes = sResult & "(" & sResValue & ")"
            
        Else
            sTransRes = sResValue
            sResult = ""
        End If
        SQL = "UPDATE LRESULT  "
        SQL = SQL & vbCrLf & "     SET RSFL = 'Y' "
        SQL = SQL & vbCrLf & "     , RSLT = '" & sTransRes & "'"
        If sRefFlag = "H" Or sRefFlag = "L" Then
            SQL = SQL & vbCrLf & "     , HLFL = '" & sRefFlag & "'"
        Else
        End If
        SQL = SQL & vbCrLf & "     , RSDT = sysdate "
        SQL = SQL & vbCrLf & "     , RSID = '" & gUserID & "' "
        SQL = SQL & vbCrLf & "   WHERE SPNO = '" & lsID & "' "
        SQL = SQL & vbCrLf & "     AND ORCD = '" & sExamCode & "' "
        SQL = SQL & vbCrLf & "     AND OKFL = 'N' "
        
        res = SendQuery(gServer, SQL)
        
    Next
    
    SQL = "update pat_res " & vbCrLf & _
          " set sendflag = '2' " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(vasList, argSpcRow, colBarCode)) & "' "
    res = SendQuery(gLocal, SQL)
                    
'    Insert_Data = 1
    
'    Insert_Data_1 = -1
'
'
'    sResFlag = False
'    lsID = ""
'    lsID = Trim(GetText(vasList, argSpcRow, colBarCode))
'
'    If Len(lsID) = 8 And IsNumeric(lsID) = True Then
'    Else
'        Exit Function
'    End If
'
'    lsReceDate = Format(Date, "yyyy") & Mid(lsID, 1, 4)
'    lsReceNo = Mid(lsID, 5)
'    lsReceNo = CStr(CCur(lsReceNo))
'
'
'    'Local에서 환자별로 결과값 가져오기
'    ClearSpread vasResTemp
'
'    SQL = " Select equipcode, examcode, resvalue, result, refflag " & vbCrLf & _
'          " From pat_res  " & vbCrLf & _
'          " Where barcode = '" & lsID & "' and resvalue <> '' "
'
'    res = db_select_Vas(gLocal, SQL, vasResTemp)
'
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    For j = 1 To vasResTemp.DataRowCnt
'        sExamCode = Trim(GetText(vasResTemp, j, 2))
'        sResValue = Trim(GetText(vasResTemp, j, 3))
'        sResult = Trim(GetText(vasResTemp, j, 4))
'        sRefFlag = Trim(GetText(vasResTemp, j, 5))
'
'        SQL = "select resgubun from equipexam where equipcode = '" & Trim(GetText(vasResTemp, j, 1)) & "'"
'        res = db_select_Col(gLocal, SQL)
'
'        sResGubun = Trim(gReadBuf(0))
'
'        If sResGubun = "1" Then '문자
'            sTransRes = sResult & "(" & sResValue & ")"
'
'        Else
'            sTransRes = sResValue
'            sResult = ""
'        End If
'
'        SQL = " Update trures"
'        SQL = SQL & vbCrLf & "   set result_value = '" & sTransRes & "',"
'        SQL = SQL & vbCrLf & "      result_decision = '" & sRefFlag & "',"
'        SQL = SQL & vbCrLf & "      machine = '" & gEquip & "'"
'        SQL = SQL & vbCrLf & "where request_date = '" & lsReceDate & "'"
'        SQL = SQL & vbCrLf & "  and exam_no = " & lsReceNo & " "
'        SQL = SQL & vbCrLf & "  and exam_part = 'S'"
'        SQL = SQL & vbCrLf & "  and exam_code = '" & sExamCode & "' "
'        SQL = SQL & vbCrLf & "  and end_report = '' "
'
'        res = SendQuery(gServer, SQL)
'
'    Next
'
'
'
'
'    SQL = "update pat_res " & vbCrLf & _
'          " set sendflag = '2' " & vbCrLf & _
'          " Where equipno = '" & gEquip & "' " & vbCrLf & _
'          " And barcode = '" & Trim(GetText(vasList, argSpcRow, colBarCode)) & "' "
'    res = SendQuery(gLocal, SQL)
'
    Insert_Data_1 = 1
    
End Function

Private Sub cmdCall_Click()
    Dim i As Long
    Dim varSendFlag
    Dim j As Long
    Dim X As Long
    Dim strResult As String
    
    
    ClearSpread vasList
    ClearSpread vasListRes, 1, 1
    vasListRes.MaxRows = 0
    
    varSendFlag = cmbTransGubun.ListIndex

    SQL = "select '', barcode, pid, pname, psex, page, recedate, diskno, sendflag from pat_res " & vbCrLf & _
          " where examdate = '" & Format(dtpExamDate, "yyyymmdd") & "' "
    If varSendFlag = 1 Or varSendFlag = 2 Then
        SQL = SQL & " and sendflag = '" & varSendFlag & "' "
    Else
        SQL = SQL & " and sendflag <> '0' "
    End If
    
    SQL = SQL & vbCrLf & " group by barcode, pid, pname,  sendflag, recedate, diskno, psex, page"
    res = db_select_Vas(gLocal, SQL, vasList)

    
    vasList.MaxRows = vasList.DataRowCnt
    For i = 1 To vasList.DataRowCnt
        If GetText(vasList, i, colState) = "1" Then
            SetText vasList, "Result", i, colState
            
        ElseIf GetText(vasList, i, colState) = "2" Then
            SetText vasList, "Trans", i, colState
            SetBackColor vasList, i, i, colBarCode, colState, 255, 255, 180
        End If
    Next
    
    ClearSpread vasResTemp
    
    SQL = "select barcode, equipcode, resvalue, result from pat_res " & vbCrLf & _
          " where examdate = '" & Format(dtpExamDate, "yyyymmdd") & "' "
    If varSendFlag = 1 Or varSendFlag = 2 Then
        SQL = SQL & " and sendflag = '" & varSendFlag & "' "
    Else
        SQL = SQL & " and sendflag <> '0' "
    End If
    
    SQL = SQL & vbCrLf & " group by barcode, equipcode, resvalue, result"
    res = db_select_Vas(gLocal, SQL, vasResTemp)
    
'''    gArr_Exam(i, 1) = i    '순서
'''    gArr_Exam(i, 2) = Trim(GetText(vasTemp, i, 2))    '장비코드
'''    gArr_Exam(i, 3) = Trim(GetText(vasTemp, i, 3))    '검사명
    
    For i = 1 To vasResTemp.DataRowCnt
        For j = 1 To vasList.DataRowCnt
            If Trim(GetText(vasResTemp, i, 1)) = Trim(GetText(vasList, j, colBarCode)) Then
                For X = 1 To vasList.MaxCols - colRStart
                    If Trim(GetText(vasResTemp, i, 2)) = Trim(gArr_Exam(X, 2)) Then
                        If gArr_Exam(X, 4) = "0" Then
                            strResult = Trim(GetText(vasResTemp, i, 3))
                        ElseIf gArr_Exam(X, 4) = "1" Then
                            strResult = Trim(GetText(vasResTemp, i, 4)) & "(" & Trim(GetText(vasResTemp, i, 3)) & ")"
                        Else
                            strResult = Trim(GetText(vasResTemp, i, 3))
                        End If
                        
                        SetText vasList, strResult, j, colRStart + CCur(gArr_Exam(X, 1))
                        Exit For
                    End If
                Next X
                Exit For
            End If
        Next j
    Next i

End Sub

Private Sub cmdClear_Click()
Dim iNumber As Integer
    
    txtMsg.Text = ""
    
    ClearSpread vasID, 1, 1
    vasID.MaxRows = 0
    ClearSpread vasRes, 1, 1
    vasRes.MaxRows = 0
End Sub

Private Sub cmdListClear_Click()
    Dim iNumber As Integer
    
    txtMsg.Text = ""
    
    ClearSpread vasList, 1, 1
    vasList.MaxRows = 0
    ClearSpread vasListRes, 1, 1
    vasListRes.MaxRows = 0
End Sub

Private Sub cmdListTrans_Click()
'선택전송
Dim VasidRow As Integer
Dim VasResRow As Integer
Dim iRow As Integer
Dim liRet As Integer
Dim iNumber As Integer

    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:선택전송") = vbCancel Then
        Exit Sub
    End If

'''    If txtUID.Text = "" Then
'''        MsgBox "사용자 확인을 해 주십시오"
'''        txtUID.SetFocus
'''        Exit Sub
'''    End If
    
    If vasList.DataRowCnt < 1 Then
        MsgBox "저장할 데이터가 없습니다."
        Exit Sub
    End If
    
    'db_BeginTran gServer
'''    Connect_Server
    For VasidRow = 1 To vasList.DataRowCnt
        vasList.Col = 1
        vasList.Row = VasidRow
        
        If vasList.Value = 1 Then '체크된 열은 저장이 안됨
'        If vasID.Value = "" Then
        
            liRet = -1
'''            liRet = Insert_Data(VasidRow)
            liRet = Insert_Data(VasidRow, vasList)
            
            If liRet = 1 Then
                'db_Commit gServer
                
                SetBackColor vasList, VasidRow, VasidRow, colBarCode, colState, 255, 255, 180
                SetText vasList, "Trans", VasidRow, colState
            Else
                SetBackColor vasList, VasidRow, VasidRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasList, "Failed", VasidRow, colState
            End If
            vasList.Col = 1
            vasList.Row = VasidRow
            vasList.Value = 0
        Else
        
        End If
    Next VasidRow
    
End Sub

Private Sub cmdVasIDWidth_Click()
    Dim i As Integer
    
    
    If cmdVasIDWidth.Caption = ">>" Then
        vasID.Width = 14385
        cmdVasIDWidth.Caption = "<<"
        
        vasID.Visible = False
        For i = colRStart + 1 To vasID.MaxCols
            vasID.Col = i
            vasID.ColHidden = False
        Next
        vasID.Visible = True
        vasID.ScrollBars = ScrollBarsBoth
    Else
        vasID.Width = 8745
        cmdVasIDWidth.Caption = ">>"
        vasID.Visible = False
        For i = colRStart + 1 To vasID.MaxCols
            vasID.Col = i
            vasID.ColHidden = True
        Next
        vasID.Visible = True
        vasID.ScrollBars = ScrollBarsVertical
    End If
End Sub

Private Sub cmdVasListWidth_Click()
    Dim i As Integer
    
    If cmdVasListWidth.Caption = ">>" Then
        vasList.Width = 14385
        cmdVasListWidth.Caption = "<<"
        
        vasList.Visible = False
        For i = colRStart + 1 To vasList.MaxCols
            vasList.Col = i
            vasList.ColHidden = False
        Next
        vasList.Visible = True
        vasList.ScrollBars = ScrollBarsBoth
    Else
        vasList.Width = 8745
        cmdVasListWidth.Caption = ">>"
        vasList.Visible = False
        For i = colRStart + 1 To vasList.MaxCols
            vasList.Col = i
            vasList.ColHidden = True
        Next
        vasList.Visible = True
        vasList.ScrollBars = ScrollBarsVertical
    End If
End Sub

Private Sub Command1_Click()
    Dim sSigFlag As String
'''    Cobas8000 "" & "1H|\^&|5694||cobas 8000^1.02|||||host|RSUPL^REAL|P|1|20120409034132|" & vbCr & _
'''              "P|1|" & vbCr & _
'''              "O|1|1226741641|0^30001^1^^QC^SC^not|^^^12^1|R||||||P||||1||||||||||F|" & vbCr & _
'''              "C|1|I|^^^^|G|" & vbCr & _
'''              "R|1|^^^12/1/not|0.75|mmol/L||||F||bmserv^SYSTEM|20120409153939|20120409034132|ISD2" & vbLf & _
'''              "2E^1^MU1#ISE#1#1^3^14^Current|" & vbCr & _
'''              "C|1|I|0|I|" & vbCr & _
'''              "L|1|N|" & vbCr & _
'''              "FF" & vbLf & chrEOT
                
        sSigFlag = Cobas8000(Text1.Text)
        
        Text1.Text = ""
End Sub

Sub Var_Clear()
    gOrderMessage = ""
    
    gBarCode = ""
'''    sBarCode = ""
'''    sSeqNo = ""
'''    sDiskno = ""
'''    sPosno = ""
    sSampleType = ""
'''    txtpat = ""
    llRow = -1
End Sub

Public Sub CobasProg(argData As String)
    Dim i As Integer
    Dim j As Integer
    Dim X As Integer
    Dim iCnt As Integer
    Dim jCnt As Integer
    Dim aCnt As Integer
    Dim bCnt As Integer
    
    Dim lsTemp As String

    Dim sDate As String
    Dim sGubun As String
    Dim sPID As String
    Dim sReceNo As String
    Dim sSpecID As String
    Dim sTestID As String
    Dim sExamCode As String
    Dim sExamCode1 As String
    Dim sExamName As String
    Dim sSeqNo As String
    Dim sSeq As String
    Dim sResClassCode As String
    Dim sFlag As String
    Dim sResult, sResult2 As String
    Dim sResult1 As String
    Dim sResValue As String
    Dim sGiho As String
    Dim sExamDate As String
    Dim sResDateTime As String
    
    Dim sRefFlag As String
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sPanicFlag As String
    Dim sDeltaFlag As String
    Dim sCnt As String
    Dim sPage As String
    
    Dim lsRefLow As String
    Dim lsRefHigh As String
    Dim lsRefRev As String
    Dim lsResDate As String
    Dim vResult As Variant
    
    
    Dim sExamCode_All As String
    Dim sPart_All As String
    Dim sBarcode As String
    Dim sBarCode1 As String
    Dim sBarcode2 As String
    Dim sOrDate As String
    Dim sRackPos As String
    
    Dim lRow As Long
    Dim lCol As Long
    
    Dim lResRow As Long
    
    Dim jRow As Integer
    
    Dim slen, sLen2 As String
    Dim aCount As Integer
    Dim iRCnt As Integer
    Dim lsLotNo As String
    Dim sAllExamCode As String
    Dim sSampleInfo As String
    Dim sCommentNum As String
    Dim sCmtProc As Boolean
    Dim strSpcData As Variant
    Dim sMin As String
    Dim sMax As String
    
    On Error GoTo ErrRes:
    
    Select Case Mid(argData, 1, 1)
    Case "H"    'Header
        gCmtFlag = ""
        gPreRow = -1
        
        Var_Clear
    Case "P"    'Patient
        gPatFlag = -1
        
    Case "O"    'Test Order
        gCmtFlag = ""
        gRecodeType = "O"
        ClearSpread vasRes
        aCount = aCount + 1
        
        
        iCnt = 0
        
        jCnt = 0
        
        For i = 1 To Len(argData)
            If Mid(argData, i, 1) = "|" Then
                iCnt = iCnt + 1
                Select Case iCnt
                Case 2  'PID
                    slen = InStr(i + 1, argData, "|")
                    sPID = Mid(argData, i + 1, slen - i - 1)
                    strSpcData = Split(argData, "^")
                    
                    sSpecID = Trim(sPID)
                    
                    If IsNumeric(sSpecID) = True Then
                        sSpecID = Format(sSpecID, "00000000")
                    End If
                    
                    gSpecID = sSpecID
                    
                Case 3
                    
                    slen = InStr(i + 1, argData, "|")
                    sPID = Mid(argData, i + 1, slen - i - 1)
                    strSpcData = Split(argData, "^")
            
                    sRackPos = Trim(strSpcData(1)) & "-" & Trim(strSpcData(2))
                    
                Case 11
                    slen = InStr(i + 1, argData, "|")
                    If slen > 0 Then
                        If Mid(argData, i + 1, slen - i - 1) = "Q" Then
                            sSampleType = "Q"
                        Else
                            sSampleType = "P"
                        End If
                    Else
                        sSampleType = "P"
                    End If
                Case 22
                    sResDateTime = ""
                        
                        slen = InStr(i + 1, argData, "|")
                        If slen > 0 Then
                            sResDateTime = Mid(argData, i + 1, slen - i - 1)
                        End If
                    If sResDateTime = "" Then
                        lsResDate = Format(Date, "yyyymmdd")
                        sExamDate = Format(Time, "hhmmss")
                    Else
                        lsResDate = Mid(sResDateTime, 1, 8)
                        sExamDate = Mid(sResDateTime, 9, 6)
                        
                    End If
                            
                        
                End Select
            End If

        Next i
        If sSampleType = "P" Then
        
            glRow = -1
            For lRow = 1 To vasID.DataRowCnt
                If Trim(GetText(vasID, lRow, colBarCode)) = gSpecID And Trim(GetText(vasID, lRow, colPID)) = lsResDate And Trim(GetText(vasID, lRow, colReceDate)) = sExamDate Then
                    glRow = lRow
                    
                    If gPatFlag = -1 Then
                        vasID_Click 2, glRow
                        
                        gPatFlag = 1
                        vasActiveCell vasID, glRow, 2
                    End If
    
                    Exit For
                End If
            Next lRow
            
            '2004/06/16 이상은========================================================
            'Order 전송뒤 Clear시 다시 바코드 스캔 안하고 결과 넘어오도록 수정
            If glRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
                glRow = vasID.DataRowCnt + 1
                If glRow > vasID.MaxRows Then
                    vasID.MaxRows = glRow + 1
                End If
                vasActiveCell vasID, glRow, colBarCode
                SetText vasID, sSpecID, glRow, colBarCode
                SetText vasID, sRackPos, glRow, colRackPos
                SetText vasID, lsResDate, glRow, colPID
                
                SetText vasID, sExamDate, glRow, colReceDate
                
            End If
            '==========================================================================
        
        ElseIf sSampleType = "Q" Then
            glRow = -1

            If glRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
                glRow = vasID.DataRowCnt + 1
                If glRow > vasID.MaxRows Then
                    vasID.MaxRows = glRow + 1
                End If
                vasActiveCell vasID, glRow, colBarCode
                SetText vasID, sSpecID, glRow, colBarCode
            End If
            '==========================================================================
            SetText vasID, "QC", glRow, colPName
        End If
        
        gPreSpecID = sSpecID
        
        gPreRow = glRow
        
        SetText vasID, "Result", glRow, colState
        
    Case "C"
        
    Case "R"    'Result
        gtestid = ""
        gResultRes = ""
        gCmtFlag = "R"
            
        iCnt = 0
    
        sExamCode = ""
        sResClassCode = ""
        sExamName = ""
        sResult = ""
        
            aCnt = 0
            
            For i = 1 To Len(argData)
                If Mid(argData, i, 1) = "|" Then
                   aCnt = aCnt + 1
                   Select Case aCnt
                   Case 2
                        slen = InStr(i + 1, argData, "|")
                        sTestID = Trim(Mid(argData, i + 1, slen - i - 1))
                        sTestID = Mid(sTestID, 4)
                        slen = InStr(1, sTestID, "/")
                        sTestID = Mid(sTestID, 1, slen - 1)
                        gtestid = sTestID
                   Case 3
                        slen = InStr(i + 1, argData, "|")
                        sResValue = Trim(Mid(argData, i + 1, slen - i - 1))
                        
                        slen = InStr(1, sResValue, "^")
                        
                        If slen > 0 Then
                            sResValue = Mid(sResValue, slen + 1)
                        End If
                        
                        gResultRes = sResValue
                        
                        SQL = "select examcode, examname, seqno from equipexam "
                        SQL = SQL & vbCrLf & " where equipcode = '" & gtestid & "' "
                        
                        res = db_select_Col(gLocal, SQL)
                        
                        sExamCode = Trim(gReadBuf(0))
                        sExamName = Trim(gReadBuf(1))
                        sSeq = Trim(gReadBuf(2))
                        
                        sGiho = ""
                        If Left(sResValue, 1) = "<" Or Left(sResValue, 1) = ">" Then
                            sGiho = Left(sResValue, 1)
                            sResValue = Trim(Mid(sResValue, 2))
                        End If
                        
                        sResult1 = Result_Set(gtestid, sResValue)
                        
                        vResult = Split(sResult1, "/")
                        
                        sResValue = sGiho & vResult(0)
                        sResult = vResult(1)
                        
                        gReadBuf(0) = ""
                        gReadBuf(1) = ""
                        sMin = ""
                        sMax = ""
                        sRefFlag = ""

                    End Select
                End If
            Next i
        
        If gtestid <> "" And sResValue <> "" Then
            If sSampleType = "P" Then
                sExamCode_All = ""
                sPart_All = ""
                
                lResRow = -1
                For j = 1 To vasRes.DataRowCnt
                    If Trim(sExamCode) = Trim(GetText(vasRes, j, colExamCode)) Then
                        lResRow = j
                        Exit For
                    End If
                Next j
                
                If lResRow = -1 Then
                    lResRow = vasRes.DataRowCnt + 1
                    If lResRow > vasRes.MaxRows Then
                        vasRes.MaxRows = lResRow
                    End If
                End If

                SetText vasRes, gtestid, lResRow, colEquipExam '장비코드
                SetText vasRes, sExamCode, lResRow, colExamCode '검사코드
                SetText vasRes, sExamName, lResRow, colExamName '검사명
                SetText vasRes, sSeq, lResRow, colSeq '순서
                SetText vasRes, sResValue, lResRow, colResValue '결과수치
                SetText vasRes, sResult, lResRow, colResult '문자결과
                SetText vasRes, sExamDate, lResRow, colResDate '검사일자
                SetText vasRes, lsResDate, lResRow, colResTime '검사시간
                SetText vasRes, sRefFlag, lResRow, colRef
                
                
                Save_Local_One glRow, lResRow, "1"
                    
                vasID_Click colBarCode, glRow
                
                SQL = "select resgubun from equipexam " & vbCrLf & _
                      "where equipno = '" & gEquip & "' and examcode = '" & sExamCode & "'"
                res = db_select_Col(gLocal, SQL)
                
                If Trim(gReadBuf(0)) = "1" Then
                    SetPositionResult glRow, gtestid, sResult & "(" & sResValue & ")"
                Else
                
                    SetPositionResult glRow, gtestid, sResValue
                End If

                
'''                SetText vasID, vasRes.DataRowCnt, glRow, colRcnt
                
            ElseIf sSampleType = "Q" Then   'QC결과
                
            End If
        End If
        
        gMsgFlag = ""
        gHeadRecode = ""
        txtData.Text = ""
        
    Case "Q"    'Request
'''        gRecodeType = "Q"
'''
'''        ClearSpread vasTemp
'''        ClearSpread vasOrder
'''        ClearSpread vasOrderBuf
'''
'''        slen = InStr(1, argData, "|")
'''        argData = Mid(argData, slen + 1)
'''
'''        slen = InStr(1, argData, "|")
'''        argData = Mid(argData, slen + 1)
'''
'''        slen = InStr(1, argData, "|")
'''        gSpecID = Mid(argData, 1, slen - 1)     '검체번호
'''
'''        gSpecID = Mid(gSpecID, 3)
'''
'''        slen = InStr(1, gSpecID, "^")
'''
''''''        gSpecID = Trim(Mid(gSpecID, 1, slen - 1))
'''
'''        gSampleInfo = ""
'''        sSampleInfo = Trim(Mid(gSpecID, slen + 1))         'sampleinfo
'''
'''        gSpecID = Trim(Mid(gSpecID, 1, slen - 1))        '검체번호
'''        gSampleInfo = sSampleInfo
'''
'''        slen = InStr(1, sSampleInfo, "^")
''''''        gSampleInfo = Mid(sSampleInfo, 1, slen)
'''        sSampleInfo = Mid(sSampleInfo, slen + 1)
'''
'''        slen = InStr(1, sSampleInfo, "^")
''''''        gSampleInfo = gSampleInfo & Mid(sSampleInfo, 1, slen)
'''        sRackPos = Mid(sSampleInfo, 1, slen - 1)
'''        sSampleInfo = Mid(sSampleInfo, slen + 1)
'''
'''
'''        slen = InStr(1, sSampleInfo, "^")
''''''        gSampleInfo = gSampleInfo & Mid(sSampleInfo, 1, slen - 1)
'''        sRackPos = sRackPos & "-" & Mid(sSampleInfo, 1, slen - 1)
'''        sSampleInfo = Mid(sSampleInfo, slen + 1)
'''
'''        slen = InStr(1, sSampleInfo, "^")
''''''        gSampleInfo = gSampleInfo & Mid(sSampleInfo, 1, slen)
'''        sSampleInfo = Mid(sSampleInfo, slen + 1)
'''
'''        slen = InStr(1, sSampleInfo, "^")
''''''        gSampleInfo = gSampleInfo & Mid(sSampleInfo, 1, slen)
'''        sSampleInfo = Mid(sSampleInfo, slen + 1)
''''''
''''''
''''''        slen = InStr(1, sSampleInfo, "^")
''''''        gSampleInfo = gSampleInfo & Mid(sSampleInfo, 1, slen - 1)
''''''        sSampleInfo = Mid(sSampleInfo, slen + 1)
''''''
'''
'''        glRow = vasID.DataRowCnt + 1
'''        If vasID.MaxRows < glRow + 1 Then
'''            vasID.MaxRows = glRow + 1
'''        End If
'''
'''        glRow = -1
'''        For lRow = 1 To vasID.DataRowCnt
'''            If Trim(GetText(vasID, lRow, colBarCode)) = gSpecID Then
'''                glRow = lRow
'''                vasActiveCell vasID, glRow, colBarCode
'''                SetText vasID, "Order", glRow, colState
'''                SetText vasID, sRackPos, glRow, colRackPos
'''                Exit For
'''            End If
'''        Next lRow
'''
'''        '==========================================================================
'''        'Order 전송뒤 Clear시 다시 바코드 스캔 안하고 결과 넘어오도록 수정
'''        If glRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
'''            glRow = vasID.DataRowCnt + 1
'''            If glRow > vasID.MaxRows Then
'''                vasID.MaxRows = glRow + 1
'''            End If
'''
'''            vasActiveCell vasID, glRow, colBarCode
'''            SetText vasID, gSpecID, glRow, colBarCode
'''            SetText vasID, "Order", glRow, colState
'''            SetText vasID, sRackPos, glRow, colRackPos
'''        End If
'''        '==========================================================================
'''
'''        'Order 만들기
'''        Make_Order gSpecID, glRow
        
    Case "L"    '자료수신 완료

        If gRecodeType = "Q" Then
        Else
        
            If mnuAuto.Checked = True Then
                If glRow > 0 And glRow <= vasID.DataRowCnt Then
                    res = -1
                    res = Insert_Data(CInt(glRow), vasID)
                    If res = 1 Then
'                            SetBackColor vasID, gPreRow, gPreRow, colCheckBox, colState, 202, 255, 112
                        SetBackColor vasID, glRow, glRow, 2, colState, 220, 250, 220
                        SetText vasID, "Trans", glRow, colState
                    ElseIf res = -1 Then
                        SetBackColor vasID, glRow, glRow, colCheckBox, colState, 255, 0, 0
                        SetText vasID, "Failed", glRow, colState
                    End If
                End If
            End If
        End If
    End Select
    
ErrRes:
    Exit Sub
    
End Sub

Private Function Result_Set(asExamCode As String, asResult As String) As String
    Dim strRefH As String
    Dim strRefM As String
    Dim strRefL As String
    Dim cRefH As String
    Dim cRefL As String
    Dim strResGubun As String
    Dim strLEquil As String
    Dim strHEquil As String
    Dim i As Integer
    Dim strRespRec As String
    Dim strPointFormat As String
    Dim cRepH As String
    Dim cRepL As String
    Dim strGiho As String
    Dim strResult As String
    Dim strResValue As String
    Dim strRefFlag As String
    
    
    On Error GoTo ErrRes:
    
    Result_Set = ""
    strRefFlag = ""
    
    strResValue = asResult
    
    If IsNumeric(strResValue) = False Then
        Result_Set = strResValue & "/" & strResValue & "/" & strRefFlag
        Exit Function
    End If
    
    For i = 1 To 11
        gReadBuf(i - 1) = ""
    Next
    
    SQL = "SELECT REPLOW, REPHIGH, REFLOW, REFHIGH, LSTRING, MSTRING, HSTRING, LEQUIL, HEQUIL, RESPREC, RESGUBUN " & vbCrLf & _
          "FROM EQUIPEXAM WHERE EQUIPNO = '" & gEquip & "' AND EQUIPCODE = '" & asExamCode & "'"
    res = db_select_Col(gLocal, SQL)
    
    cRepL = Trim(gReadBuf(0))
    cRepH = Trim(gReadBuf(1))
    cRefL = Trim(gReadBuf(2))
    cRefH = Trim(gReadBuf(3))
    strRefL = Trim(gReadBuf(4))
    strRefM = Trim(gReadBuf(5))
    strRefH = Trim(gReadBuf(6))
    strLEquil = Trim(gReadBuf(7))
    strHEquil = Trim(gReadBuf(8))
    strRespRec = Trim(gReadBuf(9))
    strResGubun = Trim(gReadBuf(10))
    
    If IsNumeric(cRepL) = True Then
        If CCur(cRepL) > CCur(strResValue) Then
            strRefFlag = "L"
'''            strResValue = cRepL
        End If
    End If
    
    If IsNumeric(cRepH) = True Then
        If CCur(cRepH) < CCur(strResValue) Then
            strRefFlag = "H"
'''            strResValue = cRepH
        End If
    End If
    
    If strResGubun = "1" Then '문자
        If IsNumeric(cRefL) = True Then
            If strLEquil = "1" Then
                If CCur(cRefL) >= CCur(strResValue) Then
                    strResult = strRefL
                End If
            Else
                If CCur(cRefL) > CCur(strResValue) Then
                    strResult = strRefL
                End If
            End If
        End If
        
        If IsNumeric(cRefH) = True Then
            If strHEquil = "1" Then
                If CCur(cRefH) <= CCur(strResValue) Then
                    strResult = strRefH
                End If
            Else
                If CCur(cRefH) < CCur(strResValue) Then
                    strResult = strRefH
                End If
            End If
        End If
        
        If IsNumeric(cRefL) = True And IsNumeric(cRefH) = True Then
            If strLEquil = "1" And strHEquil = "1" Then
                If CCur(cRefL) <= CCur(strResValue) And CCur(cRefH) >= CCur(strResValue) Then
                    strResult = strRefM
                End If
            ElseIf strLEquil = "1" And strHEquil = "0" Then
                If CCur(cRefL) <= CCur(strResValue) And CCur(cRefH) > CCur(strResValue) Then
                    strResult = strRefM
                End If
            ElseIf strLEquil = "0" And strHEquil = "1" Then
                If CCur(cRefL) < CCur(strResValue) And CCur(cRefH) >= CCur(strResValue) Then
                    strResult = strRefM
                End If
            Else
                If CCur(cRefL) < CCur(strResValue) And CCur(cRefH) > CCur(strResValue) Then
                    strResult = strRefM
                End If
                
            End If
        End If
        
    End If
    
    
    If IsNumeric(strRespRec) = True And strRespRec <> "9" Then
        If strRespRec = "0" Then
            strResValue = Format(strResValue, "####0")
        Else
            strPointFormat = ""
            For i = 1 To CInt(strRespRec)
                strPointFormat = strPointFormat & "0"
            Next
    
            strPointFormat = "##0." & strPointFormat
    
            strResValue = Format(strResValue, strPointFormat)
        End If
        
    Else
        strResValue = strResValue
    End If
    
    Result_Set = strGiho & strResValue & "/" & strResult & "/" & strRefFlag
    Exit Function
    
ErrRes:
    
    Result_Set = strResValue & "/" & strResValue & "/" & strRefFlag
    Exit Function
    
End Function

Private Sub Init_Form()
    frmInterface.Caption = gEquipName & " Interface Program"
    SSPanel1.Caption = "     " & gEquipName & "  INTERFACE"
End Sub

Private Sub Command2_Click()
'''        Dim i As Integer
'''    'exammuch, examcode, examsuga, examname, examdumy, examtype
'''    For i = 1 To vasTMaxList.DataRowCnt
'''        SQL = "select examcode from equipexam " & vbCrLf & _
'''              "where equipno = '" & Trim(GetText(vasTMaxList, i, 1)) & "'  " & vbCrLf & _
'''              "and equipcode = '" & Trim(GetText(vasTMaxList, i, 2)) & "' and examcode = '" & Trim(GetText(vasTMaxList, i, 3)) & "'"
'''        res = db_select_Col(gLocal, SQL)
'''
'''        If res = 0 Then
'''            If Trim(GetText(vasTMaxList, i, 2)) <> "" And Trim(GetText(vasTMaxList, i, 2)) <> "" Then
'''                SQL = "insert into equipexam(equipno, equipcode, examcode, examname) " & vbCrLf & _
'''                      "values('" & Trim(GetText(vasTMaxList, i, 1)) & "', '" & Trim(GetText(vasTMaxList, i, 2)) & "', " & vbCrLf & _
'''                      "'" & Trim(GetText(vasTMaxList, i, 3)) & "', '" & Trim(GetText(vasTMaxList, i, 4)) & "')"
'''                res = SendQuery(gLocal, SQL)
'''            End If
'''
'''
'''        End If
'''
'''
'''    Next
'''
'''    ClearSpread vasTMaxList
End Sub

Private Sub Command3_Click()
    MsgBox Conv_Kor_Eng("이지성")
'''    Dim strMsg As String
'''    Dim strTmaxRes As String
'''
'''    Dim i As Long
'''    Dim j As Long
'''    Dim iRow As Long
'''    Dim iCol As Long
'''
'''
'''
'''    SetText vasID, "1226741641", 1, colBarCode
'''
'''    Get_Sample_Info 1
'''
'''    ClearSpread vasTMaxRes
'''
'''
'''    iRow = 1
'''    iCol = 1
'''
'''    If chkCobas.Value = 1 Then
'''        strMsg = "C" & "mach" & vbTab & gCobas & vbTab
'''        strMsg = strMsg & "smyr" & vbTab & "12" & vbTab
'''        strMsg = strMsg & "smsn" & vbTab & "2674164" & vbTab
'''        strMsg = strMsg & "sms1" & vbTab & "1" & vbLf
'''
'''        strTmaxRes = gTMAX.TP_CALL("HAMA010A", strMsg, "")
'''
'''        If Trim(Mid(strTmaxRes, 1, 10)) = "0" Then
'''            strTmaxRes = Cut_KorEng(strTmaxRes, 100)
'''            i = InStr(1, strTmaxRes, vbTab)
'''
''''''            iRow = 1
''''''            iCol = 1
'''
'''            While i > 0
'''                SetText vasTMaxRes, Mid(strTmaxRes, 1, i - 1), iRow, iCol
'''                If iCol = 1 Then
'''                    SetText vasTMaxRes, gArchitect, iRow, 12
'''                End If
'''
'''                iCol = iCol + 1
'''                strTmaxRes = Mid(strTmaxRes, i + 1)
'''
'''
'''                i = InStr(1, strTmaxRes, vbTab)
'''                j = InStr(1, strTmaxRes, vbLf)
'''                If i > 0 And i > j And j > 0 Then
'''                    SetText vasTMaxRes, Mid(strTmaxRes, 1, j - 1), iRow, iCol + 1
'''                    iRow = iRow + 1
'''                    iCol = 1
'''                    strTmaxRes = Mid(strTmaxRes, j + 1)
'''                    i = InStr(1, strTmaxRes, vbTab)
'''                End If
'''            Wend
'''
'''        End If
'''    End If
'''
'''
'''    If chkArchitect.Value = 1 Then
'''
'''        strMsg = "C" & "mach" & vbTab & gArchitect & vbTab
'''        strMsg = strMsg & "smyr" & vbTab & "12" & vbTab
'''        strMsg = strMsg & "smsn" & vbTab & "2674164" & vbTab
'''        strMsg = strMsg & "sms1" & vbTab & "1" & vbLf
'''
'''        strTmaxRes = gTMAX.TP_CALL("HAMA010A", strMsg, "")
'''
'''        If Trim(Mid(strTmaxRes, 1, 10)) = "0" Then
'''            strTmaxRes = Cut_KorEng(strTmaxRes, 100)
'''            i = InStr(1, strTmaxRes, vbTab)
'''
''''''            iRow = 1
''''''            iCol = 1
'''
'''            While i > 0
'''                SetText vasTMaxRes, Mid(strTmaxRes, 1, i - 1), iRow, iCol
'''                If iCol = 1 Then
'''                    SetText vasTMaxRes, gArchitect, iRow, 12
'''                End If
'''
'''                iCol = iCol + 1
'''                strTmaxRes = Mid(strTmaxRes, i + 1)
'''
'''
'''                i = InStr(1, strTmaxRes, vbTab)
'''                j = InStr(1, strTmaxRes, vbLf)
'''                If i > 0 And i > j And j > 0 Then
'''                    SetText vasTMaxRes, Mid(strTmaxRes, 1, j - 1), iRow, iCol + 1
'''                    iRow = iRow + 1
'''                    iCol = 1
'''                    strTmaxRes = Mid(strTmaxRes, j + 1)
'''                    i = InStr(1, strTmaxRes, vbTab)
'''                End If
'''            Wend
'''
'''        End If
'''    End If
    
    
    
    
'''    strMsg = "O" & "smyr" & vbTab & "12" & vbTab
'''    strMsg = strMsg & "smsn" & vbTab & "2674164" & vbTab
'''    strMsg = strMsg & "sms1" & vbTab & "1" & vbTab
'''    strMsg = strMsg & "mach" & vbTab & "110" & vbLf


'''    strMsg = "P" & "mach" & vbTab & "110" & vbTab
'''    strMsg = strMsg & "slip" & vbTab & "LSR" & vbTab
'''    strMsg = strMsg & "slip1" & vbTab & "LCR" & vbTab
'''    strMsg = strMsg & "slip2" & vbTab & "" & vbTab
'''    strMsg = strMsg & "slip3" & vbTab & "" & vbTab
'''    strMsg = strMsg & "date" & vbTab & "20120409" & vbTab
'''    strMsg = strMsg & "todt" & vbTab & "20120410" & vbTab
'''    strMsg = strMsg & "flag" & vbTab & "Y" & vbLf
    
    
'''    strMsg = "C" & "mach" & vbTab & "110" & vbTab
'''    strMsg = strMsg & "smyr" & vbTab & "12" & vbTab
'''    strMsg = strMsg & "smsn" & vbTab & "2674164" & vbTab
'''    strMsg = strMsg & "sms1" & vbTab & "1" & vbLf
'''
'''
'''
'''    res = gTMAX.TP_CALL("HAMA010A", strMsg, "")
End Sub

Private Sub Command4_Click()
    ClearSpread vasTMaxRes
    SQL = "SELECT pifrslip, pifrlbno, pifrcode, pifrsubc, pifrrslt, pifrcpmv, pifrendt, pifrentm FROM PIFRSLTM.db "
    SQL = SQL & "  WHERE pifrlbno = '||R|||||'  "
    res = db_select_Vas(gServer, SQL, vasTMaxRes)
    
    SQL = "insert into PIFRSLTM.db(pifrslip, pifrlbno, pifrcode, pifrsubc, pifrrslt, pifrcpmv, pifrendt, pifrentm) "
    SQL = SQL & vbCrLf & "values('F011  ','12345678  ','21AC  ','      ','0.3       ','999999    ', '20151026', '090000') "
    res = SendQuery(gServer, SQL)
    
End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer

    '1. 화면 및 변수 초기화
    '2. 데이타베이스에 Connect 하기 - Local - Server
    '3. Ini 내용 불러오기    GetSetup
    '4. Comport Open
    
    'Timer interval = 3000 -> 10000
    
    Me.Left = 0
    Me.Top = 0
    
    cmdClear_Click
        
    GetSetup    'ini에서 DB정보 불러오기
    
    Init_Form
    
    If Not Connect_Server Then
        MsgBox "연결되지 않았습니다."
        Exit Sub
    End If
    
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        Exit Sub
    End If

    MSComm1.CommPort = gSetup.gPort
    MSComm1.RTSEnable = "True"
    MSComm1.DTREnable = "True"
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    
    
    '수정후 꼭 풀기
    
'''    WinSock_Listen Winsock1
    

    
'''    lblUser.Caption = gExamUID
'''    txtUID.Text = gExamUID

    raw_data = ""

    

    dtpToday = Date
    dtpExamDate = Date
  
    '====================로컬 DB지우기 - 30일 보관======================
    sDate = Format(DateAdd("y", CDate(dtpToday), -gLocalExpDate), "yyyymmdd")
    
    SQL = "Delete from pat_res where examdate < '" & sDate & "' "
    res = SendQuery(gLocal, SQL)
    '===================================================================

    '검사코드 가져오기
    GetExamCode

    ClearSpread vasCode

    vasID.MaxRows = 1
    vasID.ColsFrozen = 9
    vasRes.MaxRows = 20
    vasList.MaxRows = 1
    
    vasList.ColsFrozen = 9
    
    vasListRes.MaxRows = 20
    
    
'''    vasID.Visible = False
'''    For i = colRStart + 1 To vasID.MaxCols
'''        vasID.Col = i
'''        vasID.ColHidden = True
'''    Next
'''    vasID.Visible = True
    
'''    vasList.Visible = False
'''    For i = colRStart + 1 To vasList.MaxCols
'''        vasList.Col = i
'''        vasList.ColHidden = True
'''    Next
'''    vasID.Visible = True
    
    
    
    SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
''    WritePrivateProfileString "config", "UID", txtUID.Text, App.Path & "\interface.ini"
'''    gTMAX.TP_TERM
'''    DisConnect_Server
    DisConnect_Local
    
    Unload frmLogin
    
End Sub

Sub GetExamCode()
'검사코드를 array에 저장
    Dim i As Integer
    
    gAllExam = ""
    gOrderExam = ""
    gReceExam = ""
    
    
    For i = 1 To 100
        gArr_Exam(i, 1) = ""
        gArr_Exam(i, 2) = ""
        gArr_Exam(i, 3) = ""
    Next i
    
    ClearSpread vasTemp
    
'''    SQL = "Select SeqNo, EquipCode, ExamName, resgubun From EquipExam where Equipno = '" & gEquip & "' "
'''    SQL = SQL & vbCrLf & "GROUP BY SeqNo, EquipCode, ExamName, resgubun "
'''    SQL = SQL & vbCrLf & " Order by SeqNo"
    SQL = "Select SeqNo, EquipCode, ExamName, resgubun From EquipExam "
    SQL = SQL & vbCrLf & "GROUP BY SeqNo, EquipCode, ExamName, resgubun "
    SQL = SQL & vbCrLf & " Order by SeqNo"
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    vasID.MaxCols = colRStart + vasTemp.DataRowCnt
    vasList.MaxCols = colRStart + vasTemp.DataRowCnt
    
    For i = 1 To vasTemp.DataRowCnt
        If IsNumeric(Trim(GetText(vasTemp, i, 1))) = True Then
            gArr_Exam(i, 1) = i    '순서
            gArr_Exam(i, 2) = Trim(GetText(vasTemp, i, 2))    '장비코드
            gArr_Exam(i, 3) = Trim(GetText(vasTemp, i, 3))    '검사명
            gArr_Exam(i, 4) = Trim(GetText(vasTemp, i, 4))    '결과구분
            
            SetText vasID, Trim(GetText(vasTemp, i, 3)), 0, colRStart + i
            SetText vasList, Trim(GetText(vasTemp, i, 3)), 0, colRStart + i
            
        End If
        
    Next i
    
'''    For i = 1 To 100
'''        gArr_Exam(i, 1) = ""
'''        gArr_Exam(i, 2) = ""
'''        gArr_Exam(i, 3) = ""
'''    Next i


    ClearSpread vasTemp

    SQL = "Select ExamCode From EquipExam where Equipno = '" & gEquip & "' "

    res = db_select_Vas(gLocal, SQL, vasTemp)

    For i = 1 To vasTemp.DataRowCnt

        If Trim(GetText(vasTemp, i, 1)) <> "" Then
            If gAllExam = "" Then
                gAllExam = "'" & Trim(GetText(vasTemp, i, 1)) & "'"
            Else
                gAllExam = gAllExam & ",'" & Trim(GetText(vasTemp, i, 1)) & "'"
            End If
        End If
    Next i
    
End Sub


Private Sub mnuAuto_Click()
    mnuManual.Checked = False
    mnuAuto.Checked = True
End Sub

Private Sub mnuCodeConfig_Click()
    frmEquipExam.SSPanel1.Caption = "  " & gEquipName & " 장비 코드 설정"
    frmEquipExam.Show 1
    GetExamCode
End Sub

Private Sub mnuConfig_Click()
    frmConfig.SSPanel_machine.Caption = gEquipName
    frmConfig.Show 1
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuManual_Click()
    mnuManual.Checked = True
    mnuAuto.Checked = False
End Sub

Function Make_Order(asSpecid As String, asRow As Long) As String

    Dim sDate As String
    
    Dim sCnt As String
    
    Dim sOCnt As Long
    Dim sRetOrder As String
    Dim sOrder As String
    
    Dim iRow As Long
    Dim llRow As Long
    Dim llRow_Order As Long
    
    Dim sBarcode As String
    Dim sBarCode1 As String
    Dim sBarcode2 As String
    Dim sPID As String
    Dim sPName As String
    Dim sSex As String
    Dim sAge As String
    Dim sOrDate As String
    Dim sEmgFlag As String
    Dim sSampType As String
    
    Dim sEquipNo As String
    
    Dim sExamCode As String
    Dim sExamName As String
    Dim sEquipCode As String
    Dim sTempCode As String
    Dim iTmpCnt As Integer
    Dim GluFlag As Integer
    Dim CholFlag As Integer
    Dim sSeqNo As String
    
    Dim i As Integer
    Dim j As Integer
    Dim jRow As Integer
    Dim iCnt As String
    Dim jCnt As String
    Dim sSmyr As String
    Dim sSmsn As String
    Dim sSms1 As String
    
    Dim sOrderFlag As String
    Dim strMsg As String
    Dim sReceExam As String
    Dim strTmaxRes As String
    Dim sReceCode As String
    Dim sReqEquipCode As String
    Dim lsReceDate As String
    Dim lsReceNo As String
    Dim lsAllReceCode As String
    Dim sReceDate As String
    Dim sRackPos As String
    Dim sEM As String
    Dim strDBBar As String
    
    
'''    If Len(asSpecid) = 8 And IsNumeric(asSpecid) = True Then
'''    Else
'''        Exit Function
'''    End If
    
    sBarcode = asSpecid
    strDBBar = sBarcode
    
    If IsNumeric(strDBBar) = True Then
    
        
        strDBBar = Format(strDBBar, "###0")
        strDBBar = Mid(strDBBar, 1, Len(strDBBar) - 1)
        
        
        
        
        
        'Server에서 Order 가져오기
        
        SQL = "SELECT ORD_CD  "
        SQL = SQL & vbCrLf & "   FROM MCCSI.LIS_INTERFACE1_V "
        SQL = SQL & vbCrLf & "  WHERE BCODE_NO = " & strDBBar & ""
        SQL = SQL & vbCrLf & "    AND ORD_CD IN (" & gAllExam & ")"
'''        SQL = SQL & vbCrLf & "    AND STS_CD = '0'"
        
        res = db_select_Row(gServer, SQL)
        lsAllReceCode = ""
        For i = 1 To res
            If lsAllReceCode = "" Then
                lsAllReceCode = "'" & Trim(gReadBuf(i - 1)) & "'"
                
            Else
                lsAllReceCode = lsAllReceCode & ", '" & Trim(gReadBuf(i - 1)) & "'"
            End If
        Next
        
        If lsAllReceCode = "" Then
            lsAllReceCode = "''"
        End If
        
        'Server에서 Order 가져오기
        
        SQL = "SELECT DISTINCT ER_YN  "
        SQL = SQL & vbCrLf & "   FROM MCCSI.LIS_INTERFACE1_V "
        SQL = SQL & vbCrLf & "  WHERE BCODE_NO = " & strDBBar & " "
        SQL = SQL & vbCrLf & "    AND ORD_CD IN (" & gAllExam & ")"
'''        SQL = SQL & vbCrLf & "    AND STS_CD = '0'"
        
        res = db_select_Col(gServer, SQL)
        
        sEM = ""
        sEM = Trim(gReadBuf(0))
        
        
    Else
        If lsAllReceCode = "" Then
            lsAllReceCode = "''"
        End If
    End If
    
    ClearSpread vasTemp
    
    SQL = "select equipcode, equipno, examcode, EXAMNAME, SEQNO from equipexam "
    SQL = SQL & vbCrLf & "where equipno = '" & gEquip & "' and examcode in (" & lsAllReceCode & ") "
    res = db_select_Vas(gLocal, SQL, vasTemp)

    If Trim(GetText(vasID, glRow, colPID)) = "" Then
        Get_Sample_Info glRow
    End If
    
    sPID = Trim(GetText(vasID, glRow, colPID))
    sPName = Trim(GetText(vasID, glRow, colPName))
'''    sPName = Conv_Kor_Eng(Trim(GetText(vasID, glRow, colPName)))
    sAge = Trim(GetText(vasID, glRow, colPAge))
    sSex = Trim(GetText(vasID, glRow, colPSex))
    sReceDate = Trim(GetText(vasID, glRow, colReceDate))
    sRackPos = Trim(GetText(vasID, glRow, colRackPos))
    If IsNumeric(sAge) = False Then
        sAge = "0"
    End If
    
    If sSex = "M" Or sSex = "F" Then
    Else
        sSex = "U"
    End If
    
    iCnt = 0
'''    ClearSpread vasTemp
    ClearSpread vasOrder
    ClearSpread vasOrderBuf
    
'''    H|\^&|||host|||||cobas 8000^1.02|TSDWN|P|1|
'''    P|1||||^||||
'''    O|1|" & barcode & "|0^40002^3^^S1^SC|^^^989\^^^990\^^^991|R||||||A||||1||||||||||O
'''    L|1|N

    'Order 생성하기==================================================

    sEmgFlag = lsFlag
    
    llRow_Order = 1

    gCurMsgCnt = 1
    
'''1H|\^&|||ASTM-Host
'''59
'''2P|1||325618
'''70
'''3O|1|04020060|41^@5^1|^^^900^0\^^^410^0|R||||||N||||||||||||||O
'''A8
'''4L|1
'''3D
    'HeadH|\^&|||cobas-e411^1|||||host|RSUPL^REAL|P|1
'1H|\^&|||host^2|||||cobas6000|TSDWN^BATCH|P|1
'''P|1|||||||M||||||40^Y
'''O|1|            P110055695|0^5004^2^^S1^|^^^250|R||20111004133833||||N||||1||||||||||O
'''L|1|N
'''9C

    gHeader = "H|\^&|||host^1|||||cobas-e411|TSDWN^REPLY|P|1"
    gPatient = "P|1"
    
    sReqEquipCode = ""
    For i = 1 To vasTemp.DataRowCnt
        If Trim(GetText(vasTemp, i, 1)) <> "" Then
            sReqEquipCode = sReqEquipCode & "^^^" & Trim(GetText(vasTemp, i, 1)) & "^\"
        End If
        
        sEquipCode = Trim(GetText(vasTemp, i, 1))
        sEquipNo = Trim(GetText(vasTemp, i, 2))
        sExamCode = Trim(GetText(vasTemp, i, 3))
        
'''        SQL = "SELECT EQUIPCODE, EXAMNAME, SEQNO FROM EQUIPEXAM " & vbCrLf & _
'''              "WHERE equipno = '" & sEquipNo & "' and  EXAMCODE = '" & sExamCode & "' "
'''        res = db_select_Col(gLocal, SQL)
        
'''        sEquipCode = Trim(gReadBuf(0))
        sExamName = Trim(GetText(vasTemp, i, 4))
        sSeqNo = Trim(GetText(vasTemp, i, 5))
        
        SetPositionResult asRow, sEquipCode, "*"
        
        SQL = "select barcode from pat_res where barcode = '" & Trim(asSpecid) & "' and examcode = '" & Trim(sExamCode) & "'"
        res = db_select_Col(gLocal, SQL)
        If res = 0 Then
        
            SQL = "insert into pat_res(equipno, examdate, recedate, barcode, examcode, equipcode, result, resvalue, pname, pid, " & vbCrLf & _
                  "                    seqno, page, psex, examname, diskno) " & vbCrLf & _
                  " values('" & gEquip & "', '" & Format(Date, "yyyymmdd") & "', '" & Trim(sReceDate) & "', " & vbCrLf & _
                  " '" & Trim(asSpecid) & "','" & Trim(sExamCode) & "','" & Trim(sEquipCode) & "', '', '', " & vbCrLf & _
                  "'" & Trim(sPName) & "', '" & Trim(sPID) & "', '" & Trim(sSeqNo) & "', '" & Trim(sAge) & "', '" & Trim(sSex) & "', '" & Trim(sExamName) & "', '" & sRackPos & "')"
            res = SendQuery(gLocal, SQL)
        End If
            
    Next
    
    If sReqEquipCode <> "" Then
        sReqEquipCode = Mid(sReqEquipCode, 1, Len(sReqEquipCode) - 1)
    End If
    
    sEmgFlag = "R"
    
    sSampType = "1"
    
'''    i = InStr(1, gSampleInfo, "^")
'''    sSampType = Mid(gSampleInfo, i + 1)
'''
'''    i = InStr(1, sSampType, "^")
'''    sSampType = Mid(sSampType, i + 1)
'''
'''    i = InStr(1, sSampType, "^")
'''    sSampType = Mid(sSampType, i + 1)
'''
'''    i = InStr(1, sSampType, "^")
'''    sSampType = Mid(sSampType, i + 1)
'''
'''    i = InStr(1, sSampType, "^")
'''    sSampType = Mid(sSampType, 1, i - 1)
'''
'''    sSampType = Mid(sSampType, 2)
'''
'''    i = InStr(1, gSampleInfo, "^")
'''
'''    gSampleInfo = Mid(gSampleInfo, i)
    
    
    '''3O|1|04020060|41^@5^1|^^^900^0\^^^410^0|R||||||N||||||||||||||O

    'O|1|500169|^50017^3^^S1^SC|^^^8706^|R||||||A||||1||||||||||O
    ''''O|1|            P110055695|0^5004^2^^S1^|^^^250|R||20111004133833||||N||||1||||||||||O
    '응급여부 sEM : 응급(S) 일반(R)
    
    If sEM = "Y" Then
        sEM = "S"
    Else
        sEM = "R"
    End If
    
    sRetOrder = "O|1|" & Trim(asSpecid) & "|" & gSampleInfo & "|" & sReqEquipCode & "|" & sEM & "||||||A||||1||||||||||O"

    gMsgEnd = "L|1|N"
    
    'Order 전송하기==============================================
    gOrderMessage = gHeader & vbCr
    gOrderMessage = gOrderMessage & gPatient & vbCr
    gOrderMessage = gOrderMessage & sRetOrder & vbCr
    gOrderMessage = gOrderMessage & gMsgEnd & vbCr
    
    
End Function

Private Sub SetPositionResult(asRow As Long, asEquipCode As String, asResult As String)
    Dim strEquipCode As String
    Dim strResult As String
    Dim lngRow As Long
    Dim i As Integer
    
    lngRow = asRow
    strEquipCode = asEquipCode
    strResult = asResult

    For i = colRStart + 1 To vasID.MaxCols
        If Trim(gArr_Exam(i - colRStart, 2)) = Trim(strEquipCode) Then
            SetText vasID, strResult, lngRow, i
            Exit For
        End If
    Next
End Sub

Public Function GetExamCode_Equip(argCode As String, argReceNo As String, argDate As String) As Integer
'검체번호에 존재하는 장비번호 해당하는 검사코드 가져오기

    Dim i As Integer
    Dim sExamCode As String
     
    sExamCode = ""
    GetExamCode_Equip = -1
    ClearSpread frmInterface.vaSpread1
    
    If argCode = "" Then
        Exit Function
    End If
    
    sExamCode = ""
    SQL = "Select ExamCode From EquipExam" & vbCrLf & _
          "Where Equip = '" & gEquip & "'" & vbCrLf & _
          "  And EquipCode = '" & argCode & "' "
    res = db_select_Vas(gServer, SQL, frmInterface.vaSpread1)
    
    For i = 1 To frmInterface.vaSpread1.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & ",'" & Trim(GetText(frmInterface.vaSpread1, i, 1)) & "'"
        Else
            sExamCode = "'" & Trim(GetText(frmInterface.vaSpread1, i, 1)) & "'"
        End If
    Next i
     
    gAllExam1 = sExamCode
    
    GetExamCode_Equip = 1
    
End Function


Function Get_Sample_Info(ByVal asRow As Long) As Integer
    Dim sID As String
    
    Dim lsPID As String
    Dim lsPname As String
    Dim lsDate As String
    Dim lsSDate As String
    Dim lsEDate As String
    Dim iRow As Integer
    Dim iCol As Integer
    Dim strTmaxRes As String
    Dim i As Long
    Dim j As Long
    Dim lsRow As Long
    Dim strMsg As String
    Dim lsReceNo As String
    
    
    '환자정보 가져오기
    sID = Trim(GetText(vasID, asRow, colBarCode))   '샘플 바코드 번호
    
    If IsNumeric(sID) = False Then
        Exit Function
        
    End If
    sID = Format(sID, "###0")
    
    sID = Mid(sID, 1, Len(sID) - 1)
    
'''    If IsNumeric(sID) = True Then
       
      '챠트번호, 환자명, 주민번호
      
      
    SQL = "SELECT PTNT_NO, PTNT_NM, READING_YMD, SEX, AGE "
    SQL = SQL & vbCrLf & "   FROM MCCSI.LIS_INTERFACE1_V "
    SQL = SQL & vbCrLf & "  WHERE BCODE_NO = " & sID & ""
    
    SQL = SQL & vbCrLf & " GROUP BY PTNT_NO, PTNT_NM, READING_YMD, SEX, AGE"
    
    res = db_select_Col(gServer, SQL)
    
    If res > 0 Then
        SetText vasID, Trim(gReadBuf(0)), asRow, colPID
        SetText vasID, Trim(gReadBuf(1)), asRow, colPName
        SetText vasID, Trim(gReadBuf(2)), asRow, colReceDate
        SetText vasID, Trim(gReadBuf(3)), asRow, colPSex
        SetText vasID, Trim(gReadBuf(4)), asRow, colPAge
    
    End If
'''    Else
'''
'''    End If
    
End Function

'''Function Get_QC_Info(ByVal asRow As Long) As Integer
'''    Dim sID, lsPID As String
'''
'''    '샘플 환자 정보 가져오기
'''    sID = Trim(GetText(vasID, asRow, colBarCode))   '샘플 바코드 번호
'''
'''    '사용하는 정보만 가져오기
'''    SQL = "select distinct QCBarcode, LotNo, QCInLevel, LabCode   " & vbCrLf & _
'''          "from QCInItem " & vbCrLf & _
'''          "where EquipCode = '" & gEquip & "' " & vbCrLf & _
'''          "  and QCBarcode = '" & sID & "' " & vbCrLf & _
'''          "  and AppDate <= '" & Trim(dtpToday) & "' " & vbCrLf & _
'''          "  and UseFlag = 'Y' "
'''    res = db_select_Col(gServer, SQL)
'''    If res = 1 Then
'''        SetText vasID, gReadBuf(1), asRow, colPID
'''        SetText vasID, gReadBuf(2), asRow, colPName
'''        SetText vasID, gReadBuf(3), asRow, colJumin
'''    End If
'''
'''End Function

Function SetResult(asResult As String, asExamCode As String) As String
'DB에서 불러오기
'    Dim iFloat As Integer
    Dim iFloat As String
    
    If Not IsNumeric(asResult) Then
        Exit Function
    End If

'    Select Case aiItem
'    Case 7, 16
'        iFloat = 2
'    Case 14
'        iFloat = 0
'    Case Else
'        iFloat = 1
'    End Select
'
'    If iFloat = 0 Then
'        SetResult = CStr(CCur(asResult))
'    Else
'        SetResult = CStr(CCur(Left(asResult, 5 - iFloat)) & "." & Right(asResult, iFloat))
'    End If
 
    gReadBuf(0) = ""
    
    SQL = " Select Point From ExamMaster " & vbCrLf & _
          " Where HID = '115' " & vbCrLf & _
          " And ExamCode = '" & Trim(asExamCode) & "' " & vbCrLf & _
          " And UseFlag = 'Y' "
    res = db_select_Col(gServer, SQL)
    
    iFloat = gReadBuf(0)
    
    '2004/05/31 이상은
    'ASO 관리자에는 소수점 2자리로 셋팅되어 있으나 1자리로 할 것
    If asExamCode = "C4633AJ" Then   'ASO
        iFloat = 1
    End If
    
    Select Case iFloat
    Case 0
        SetResult = Format(asResult, "#,##0")
    Case 1
        SetResult = Format(asResult, "#,##0.0")
    Case 2
        SetResult = Format(asResult, "#,##0.00")
    Case 3
        SetResult = Format(asResult, "#,##0.000")
    Case Else
    
    End Select
    
End Function



Private Sub MSComm1_OnComm()

    Dim sTmp As String
    Dim strSendData
    Dim strResFlag
    Dim sSigFlag As String
    Dim sStemp As String

    sTmp = MSComm1.Input
    
    Select Case sTmp
    Case chrENQ
        txtBuff.Text = sTmp
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
        MSComm1.Output = chrACK
        
    Case chrLF
        txtBuff.Text = txtBuff.Text & sTmp
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
        MSComm1.Output = chrACK
    Case chrEOT
        txtBuff.Text = txtBuff.Text & sTmp
        gOrderMessage = ""
        gOrderCnt = 1
        comSend = "stENQ"
        Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & txtBuff
        sSigFlag = Cobas8000(txtBuff.Text)
        If sSigFlag = "Q" Then

            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrENQ
            MSComm1.Output = chrENQ
        End If
    Case chrACK
        Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & chrACK
        If comSend = "stENQ" Then
            sStemp = SendOrder
            MSComm1.Output = sStemp
            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & sStemp

        ElseIf comSend = "stOrder" Then
            MSComm1.Output = chrEOT
            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrEOT

        End If
        
    Case Else
        txtBuff.Text = txtBuff.Text & sTmp
        
    End Select
    


End Sub

Private Sub sspMode_Click()
    If sspMode.Caption = "수정모드" Then
        sspMode.Caption = "전송모드"
        sspMode.BackColor = &HFF0000
        sspMode.ForeColor = &HFFFFFF
        vasRes.OperationMode = 1
        
    ElseIf sspMode.Caption = "전송모드" Then
        sspMode.Caption = "수정모드"
        sspMode.BackColor = &H8000&
        sspMode.ForeColor = &HFFFFFF
        vasRes.OperationMode = 0
        
        vasActiveCell vasRes, 1, colResult
        vasRes.SetFocus
    End If

End Sub


'Private Sub subUp_Click()
'Dim sValue As String
'Dim sTmp As String
'Dim i As Integer
'Dim j As Integer
'
'    sTmp = ""
'
'    vasID.Row = vasID.ActiveRow
'    vasID.Col = vasID.ActiveCol
'
'    sTmp = vasID.Text
'
'    sValue = InputBox("변경할 검체번호를 입력하세요")
'
'    If Trim(sValue) <> "" Then
'        If MsgBox("" & sTmp & "를 " & sValue & "로 수정하시겠습니까?", vbYesNo, "확인") = vbYes Then
'            SetText vasID, sValue, vasID.Row, vasID.Col
'
'            If Trim(GetText(vasID, vasID.Row, colBarCode)) <> "" Then
'                Get_Sample_Info vasID.Row
'
'                For i = 1 To vasRes.DataRowCnt
'                    Save_Local_One vasID.Row, i, "A"
'                Next
'            End If
'        End If
'    End If
'End Sub

'''Private Sub txtToday_KeyDown(KeyCode As Integer, Shift As Integer)
'''    Dim i As Integer
'''
'''    If KeyCode = vbKeyReturn Then
'''
'''    SQL = "select barcode, receno, pid, pname, pjumin, psex, page, '', sendflag from pat_res " & vbCrLf & _
'''          "where examdate = '" & Format(Trim(txtToday), "yyyymmdd") & "' and equipno = '0025' " & vbCrLf & _
'''          "group by barcode, receno, pid, pname, pjumin, psex, page,  sendflag"
'''    res = db_select_Vas(gLocal, SQL, vasID, vasID.DataRowCnt + 1, 2)
'''
'''    For i = 1 To vasID.DataRowCnt
'''        If GetText(vasID, i, colState) = "A" Then
'''            SetText vasID, "수신완료", i, colState
'''            SetBackColor vasID, i, i, colCheckBox, colCheckBox, 100, 122, 255
'''        ElseIf GetText(vasID, i, colState) = "B" Then
'''            SetText vasID, "전송완료", i, colState
'''            SetBackColor vasID, i, i, colCheckBox, colCheckBox, 202, 255, 112
'''        End If
'''    Next
'''    End If
'''End Sub

'''Private Sub Timer1_Timer()
''''''    If dtpToday <> Date Then
''''''        dtpToday = Date
''''''    End If
'''
'''End Sub

Private Sub txtUID_KeyDown(KeyCode As Integer, Shift As Integer)
'''    If KeyCode = 13 Then
'''        gExamUID = txtUID.Text
'''        Call WritePrivateProfileString("CONFIG", "UID", txtUID.Text, App.Path & "\Interface.ini")
'''    End If
End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim i As Integer
    
    Dim lsTempBarCode As String
    Dim lsPID As String
    Dim lsPname As String
    Dim lsSex As String
    Dim lsAge As String
    
    '샘플번호에 해당 하는 검사결과 Local Databse에서 가져오기
    
    ClearSpread vasRes
    vasRes.MaxRows = 0
    
    lsID = Trim(GetText(vasID, Row, colBarCode))
        
    
    SQL = "select equipcode, examcode, examname, resvalue, result, seqno, examdate, examtime, refflag " & vbCrLf & _
          "FROM pat_res " & vbCrLf & _
          "WHERE  " & vbCrLf & _
          "  " & vbCrLf & _
          "  Barcode = '" & Trim(GetText(vasID, Row, colBarCode)) & "' " & vbCrLf & _
          "  order by seqno, equipcode"

    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

End Sub

Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    Dim sExamTime As String
    Dim strPAge As String
    
    strPAge = Trim(GetText(vasID, asRow1, colPAge))
    If IsNumeric(strPAge) = False Then
        strPAge = "0"
    End If
    
    
    sExamDate = ""
    sExamDate = Trim(GetText(vasRes, asRow2, colResDate))
    sExamTime = Trim(GetText(vasRes, asRow2, colResTime))
    
    If Trim(sExamDate) = "" Then
        sExamDate = Format(Date, "yyyymmdd")
    End If
    
    
    SQL = "delete  FROM pat_res " & vbCrLf & _
          "WHERE " & vbCrLf & _
          "  equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasID, asRow1, colBarCode)) & "' "

    res = db_select_Row(gLocal, SQL)
    
'''    If res > 0 Then
'''        SQL = "update pat_res set resvalue = '" & Trim(GetText(vasRes, asRow2, colResValue)) & "', " & vbCrLf & _
'''              "result = '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf & _
'''              "sendflag = '" & asSend & "', " & vbCrLf & _
'''              "examdate = '" & sExamDate & "', examtime = '" & sExamTime & "', refflag = '" & Trim(GetText(vasRes, asRow2, colRef)) & "' " & vbCrLf & _
'''              "WHERE  " & vbCrLf & _
'''              "  equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "'" & vbCrLf & _
'''              "  AND barcode = '" & Trim(GetText(vasID, asRow1, colBarCode)) & "' "
'''        res = SendQuery(gLocal, SQL)
'''
'''    Else
        SQL = "insert into pat_res(equipno, examdate, barcode, equipcode, examcode, " & vbCrLf & _
              "sendflag, seqno, examname, resvalue, " & vbCrLf & _
              "result, examtime, pid, pname, refflag, psex, page, diskno, recedate) " & vbCrLf & _
              "values('" & gEquip & "', '" & sExamDate & "', '" & Trim(GetText(vasID, asRow1, colBarCode)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " & vbCrLf & _
              "'" & asSend & "', '" & Trim(GetText(vasRes, asRow2, colSeq)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colExamName)) & "', '" & Trim(GetText(vasRes, asRow2, colResValue)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf & _
              "'" & sExamTime & "', '" & Trim(GetText(vasID, asRow1, colPID)) & "', '" & Trim(GetText(vasID, asRow1, colPName)) & "', '" & Trim(GetText(vasRes, asRow2, colRef)) & "', '" & Trim(GetText(vasID, asRow1, colPSex)) & "', '" & strPAge & "', '" & Trim(GetText(vasID, asRow1, colRackPos)) & "', '" & Trim(GetText(vasID, asRow1, colReceDate)) & "') "
        res = SendQuery(gLocal, SQL)
'''    End If
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Private Sub vasID_KeyPress(KeyAscii As Integer)
    Dim sSpecID As String
    Dim llRow As Long
    Dim iRow As Long

    If KeyAscii = 13 Then

        llRow = vasID.ActiveRow
        sSpecID = Trim(GetText(vasID, llRow, colBarCode))

        '샘플의 환자 정보 가져오기
        Get_Sample_Info llRow
        
        For iRow = 1 To vasRes.DataRowCnt
            Save_Local_One llRow, iRow, "A"
        Next
    End If
End Sub

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If

'    PopupMenu mnuPop
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim i As Integer
    
    Dim lsTempBarCode As String
    Dim lsPID As String
    Dim lsPname As String
    Dim lsSex As String
    Dim lsAge As String
    
    '샘플번호에 해당 하는 검사결과 Local Databse에서 가져오기
    
    ClearSpread vasListRes
    vasListRes.MaxRows = 0
    
    lsID = Trim(GetText(vasList, Row, colBarCode))


    SQL = "select equipcode, examcode, examname, resvalue, result, seqno, examdate, examtime, refflag " & vbCrLf & _
          "FROM pat_res " & vbCrLf & _
          "WHERE Barcode = '" & Trim(GetText(vasList, Row, colBarCode)) & "' " & vbCrLf & _
          "  order by seqno, equipcode"


    res = db_select_Vas(gLocal, SQL, vasListRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If


End Sub

Private Sub vasres_rightclick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim VasidRow As Integer
    Dim VasResRow As Integer
    
    VasidRow = vasID.ActiveRow
    VasResRow = vasRes.ActiveRow
    If VasidRow < 1 Or VasidRow > vasID.DataRowCnt Then
        Exit Sub
    End If
    If VasResRow < 1 Or VasResRow > vasRes.DataRowCnt Then
        Exit Sub
    End If
    
    PopupMenu mnuPop

End Sub

Private Sub subDel_Click()
    Dim i As Long
    Dim VasidRow As Integer
    Dim VasResRow As Integer
    Dim X As Long
    Dim j As Long
    Dim C, r, c2, r2

    VasidRow = vasID.ActiveRow
    VasResRow = vasRes.ActiveRow
    If VasidRow < 1 Or VasidRow > vasID.DataRowCnt Then
        Exit Sub
    End If
    If VasResRow < 1 Or VasResRow > vasRes.DataRowCnt Then
        Exit Sub
    End If

    If vasRes.IsBlockSelected Or vasRes.SelectionCount Then

        vasRes.BlockMode = True
'        db_BeginTran gLocal
        
        For X = 0 To vasRes.SelectionCount - 1
            vasRes.GetSelection X, C, r, c2, r2
            vasRes.Col = C
            vasRes.Col2 = c2
            vasRes.Row = r
            vasRes.Row2 = r2
            If IsNumeric(r) = True And IsNumeric(r2) = True Then
                If CInt(r) > 0 And CInt(r2) > 0 Then
                    For j = r To r2
                        SQL = "Delete from pat_res where barcode = '" & Trim(GetText(vasID, VasidRow, colBarCode)) & "' " & vbCrLf & _
                              "and equipcode = '" & Trim(GetText(vasRes, j, colEquipExam)) & "' "
                        res = SendQuery(gLocal, SQL)
                        
                    Next
                End If
            End If
        Next X
        vasRes.BlockMode = False
'        db_Commit gLocal
        

    End If

'    SQL = "Delete from pat_res where barcode = '" & Trim(GetText(vasID, VasidRow, colBarCode)) & "' " & vbCrLf & _
'          "and equipcode = '" & Trim(GetText(vasRes, VasResRow, colEquipExam)) & "' "
'    res = SendQuery(gLocal, SQL)
    
    vasID_Click colBarCode, VasidRow
    vasRes_Click 3, 1
End Sub

'Private Sub subResDel_Click()
'    Dim i As Long
'    i = vasID.ActiveRow
'    vasID.DeleteRows i, 1
'    If i > vasID.DataRowCnt Then
'        i = vasID.DataRowCnt
'    End If
'    vasID.MaxRows = vasID.DataRowCnt
'    vasActiveCell vasID, i, colBarCode
'    vasID.SetFocus
'End Sub


Private Sub vasRes_Click(ByVal Col As Long, ByVal Row As Long)
   vasRes.Row = vasRes.ActiveRow
   vasRes.Col = vasRes.ActiveCol
   ConfirmData = vasRes.Value
    
End Sub

'''Public Function HL7_Ack(argMSH As String) As String
'''    Dim strMSH As String
'''    Dim strACK As String
'''    Dim strDateTime As String
'''    Dim strSplit() As String
'''    Dim strSigNum As String
'''
'''    Dim i As Integer
'''    Dim j As Integer
'''
'''    strMSH = argMSH
'''    strSplit = Split(strMSH, "|")
'''
'''
'''    '''MSH|^~\&|cobas 8000||host||20130104114005||OUL^R22^REAL|31777||2.5||||AA||UNICODE UTF-8|
'''    strDateTime = Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
'''
'''
'''    strACK = Chr(11)
'''    strACK = strACK & "MSH|^~\&|host||cobas 8000||" & strDateTime & "||ACK|38753||2.5||||NE||UNICODE·UTF-8|" & vbCr
'''    strACK = strACK & "MSA|AA|38749||" & vbCr
'''    strACK = strACK & Chr(28) & vbCr
'''
'''
'''End Function

''''WinSock Control ==============================================================================================================
'''Public Sub WinSock_Listen(argWinSock As Winsock)
'''    Dim sWinSockPort As String
'''
'''
'''    sWinSockPort = gSetup.gPort
'''
'''
'''    If sWinSockPort = "0" Or IsNumeric(sWinSockPort) = False Then
'''        Exit Sub
'''    End If
'''
'''    If argWinSock.State <> sckClosed Then
'''        argWinSock.Close
'''    End If
'''
'''    argWinSock.LocalPort = sWinSockPort
'''    argWinSock.Listen
'''
''''''    If EquipNum = 1 Then
''''''        lblConnect1.Caption = "연결 대기중..."
''''''    Else
''''''        lblConnect2.Caption = "연결 대기중..."
''''''    End If
'''
'''End Sub
'''
'''Private Sub Winsock1_Close()
'''
'''    If Winsock1.State <> sckClosed Then
'''        Winsock1.Close
'''    End If
'''    Winsock1.LocalPort = gSetup.gPort
'''    Winsock1.Listen
'''
'''
''''''    lblConnect1.Caption = "연결 대기중..."
'''
'''End Sub
'''
'''Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'''    If Winsock1.State <> sckClosed Then
'''        Winsock1.Close
'''    End If
'''
'''    Winsock1.Accept requestID
''''''    lblConnect1.Caption = "연결[" & requestID & "]" & Winsock1.RemoteHostIP
'''End Sub
'''
'''Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'''
'''    Dim sTmp As String
'''    Dim strSendData
'''    Dim strResFlag
'''    Dim sSigFlag As String
'''    Dim sStemp As String
'''
'''    Winsock1.GetData sTmp
''''''    Winsock1.SendData sTmp
'''    Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & sTmp
'''
'''    If InStr(1, sTmp, chrENQ) > 0 Then
'''        txtBuff.Text = sTmp
'''
'''        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
'''        Winsock1.SendData chrACK
'''    End If
'''
'''    If InStr(1, sTmp, chrLF) > 0 Then
'''        txtBuff.Text = txtBuff.Text & sTmp
'''        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
'''        Winsock1.SendData chrACK
'''    End If
'''
'''    If InStr(1, sTmp, chrEOT) > 0 Then
'''
'''        txtBuff.Text = txtBuff.Text & sTmp
'''        gOrderMessage = ""
'''        gOrderCnt = 1
'''        comSend = "stENQ"
'''
'''        sSigFlag = Cobas8000(txtBuff.Text)
'''        If sSigFlag = "Q" Then
'''
'''            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrENQ
'''            Winsock1.SendData chrENQ
'''        End If
'''
'''    End If
'''
'''    If InStr(1, sTmp, chrACK) > 0 Then
'''        Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & chrACK
'''        If comSend = "stENQ" Then
'''            sStemp = SendOrder
'''            Winsock1.SendData sStemp
'''            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & sStemp
'''
'''        ElseIf comSend = "stOrder" Then
'''            Winsock1.SendData chrEOT
'''            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrEOT
'''
'''        End If
'''
'''    End If
'''
'''
'''End Sub

Function SendOrder() As String

    Dim sSendOrder As String
    
    If Len(gOrderMessage) > 240 Then

        If gOrderCnt = 8 Then
            gOrderCnt = 0
        End If

        sSendOrder = CStr(gOrderCnt) & Left(gOrderMessage, 240) & chrETB
        gOrderMessage = Mid(gOrderMessage, 241)

        sSendOrder = chrSTX & sSendOrder & CheckSum(sSendOrder) & chrCR & chrLF
'''        SaveQuery sSendOrder, 1

        gOrderCnt = gOrderCnt + 1
        comSend = "stENQ"

'''        gPreMsg = sSendOrder
'''        Save_Raw_Data "[TX]" & sSendOrder
        SendOrder = sSendOrder
'''        MSComm1.Output = sSendOrder

    Else
        If gOrderCnt = 8 Then
            gOrderCnt = 0
        End If
        
        sSendOrder = CStr(gOrderCnt) & gOrderMessage & chrETX
        sSendOrder = chrSTX & sSendOrder & CheckSum(sSendOrder) & chrCR & chrLF
                
        gOrderMessage = ""
        comSend = "stOrder"
        
'''        gPreMsg = sSendOrder
'''        Save_Raw_Data "[TX]" & sSendOrder
        SendOrder = sSendOrder
'''        MSComm1.Output = sSendOrder
    End If
End Function

Function Cobas8000(asVar As String) As String
    Dim i As Integer
    Dim iIndex As Integer
    
    Dim lsData As String
    Dim lsTemp As String
    
    Dim lsHead As String
    Dim lsPatient As String
    Dim lsRequest As String
    Dim lsOrder As String

    Dim lsMessage As String
    
    Dim lsMSGflag As String
    
    Cobas8000 = ""
    
    lsMessage = ""
    
    If asVar = "" Then
        Exit Function
    End If
    
    ClearSpread vasRes
    
    
    iIndex = 0
    lsData = asVar
    
    lsData = Replace(lsData, chrENQ, "")
    lsData = Replace(lsData, chrEOT, "")
    
    i = InStr(1, lsData, chrSTX)
    
    While i > 0
        lsData = Mid(lsData, 1, i - 1) & Mid(lsData, i + 2)
        i = InStr(1, lsData, chrSTX)
    Wend
    
    i = InStr(1, lsData, chrLF)
    
    While i > 0
        lsData = Mid(lsData, 1, i - 4) & Mid(lsData, i + 1)
        i = InStr(1, lsData, vbLf)
    Wend
    
    
    lsData = Replace(lsData, chrETB, "")
    lsData = Replace(lsData, chrETX, "")
    
    
    i = InStr(1, lsData, Chr(13))
    Do While i > 0
        lsTemp = Mid(lsData, 1, i - 1)
        lsData = Mid(lsData, i + 1)
        
        
        
        Select Case Left(lsTemp, 1)
        Case "Q"
            lsMSGflag = "Q"
        Case "O"
            lsMSGflag = "O"
        End Select
        
        '-- Define
        Call CobasProg(lsTemp)
        
        i = InStr(1, lsData, chrCR)
    Loop
    
    Cobas8000 = lsMSGflag
End Function

